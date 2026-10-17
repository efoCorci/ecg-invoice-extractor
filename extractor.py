import os
import io
import re
import asyncio
from typing import List, Dict, Any, Optional

import pymupdf
from PIL import Image

try:
    import winocr
    HAS_WINOCR = True
except ImportError:
    HAS_WINOCR = False


def clean_number(val: Optional[str]) -> Optional[float]:
    """Parse string to float/int, handling commas, currency prefixes, and OCR artifacts."""
    if not val:
        return None
    val_str = str(val).replace('GHS', '').replace('CHS', '').replace('kWh', '').replace('Mh', '').strip()
    # Normalize multiple dots: e.g. 2.866.77 -> 2866.77
    if val_str.count('.') > 1:
        parts = val_str.split('.')
        val_str = ''.join(parts[:-1]) + '.' + parts[-1]
    # Remove commas and spaces between digits: '2, 999.30' -> '2999.30'
    val_str = re.sub(r'(?<=\d)[\s,]+(?=\d)', '', val_str)
    val_str = val_str.replace(',', '').strip()

    m = re.search(r'[-+]?\d*\.?\d+', val_str)
    if m:
        cleaned = m.group(0)
        try:
            if '.' in cleaned:
                return round(float(cleaned), 2)
            return int(cleaned)
        except ValueError:
            return None
    return None


def extract_field(patterns, text: str, default: str = "") -> str:
    """Try regex patterns in order until a match is found."""
    if isinstance(patterns, str):
        patterns = [patterns]
    for pattern in patterns:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            res = m.group(1).strip()
            if res:
                return res
    return default


def parse_invoice_text(text: str, filename: str = "") -> Dict[str, Any]:
    """Extract required 15 fields from raw text (digital or OCR) of an ECG invoice page."""
    
    # 1. Customer Name
    customer_name = extract_field([
        r'Customer Name\s*\n\s*([^\n\r]+)',
        r'Customer\s*Name:?\s*(?:\n|\s+)\s*([A-Za-z\s,\.]+?)(?=\s+Phone|\s+Account|\s+Service|\n|$)',
        r'Customer\s*Name:\s*([^\n\r]+)'
    ], text)

    # 2. Phone Number
    phone_number = extract_field([
        r'Phone Number\s*\n\s*([0-9\+\-\s]+)',
        r'(?:Phone\s*(?:Number|umber)|Phone:?)\s*(?:\n|\s+)\s*([0-9\+\-\s]{9,15})',
        r'Phone Number\s*\n\s*([^\n\r]+)'
    ], text)

    # 3. Digital Address
    digital_address = extract_field([
        r'Digital Address\s*(?:\n|\s+)\s*([A-Za-z0-9\-]+|N/A)',
        r'Digital Address\s*\n\s*([^\n\r]+)'
    ], text, default="N/A")

    # 4. Account Number
    account_number = extract_field([
        r'Account Number\s*\n\s*([0-9A-Za-z\-]+)',
        r'Account\s*(?:Number|No):?\s*(?:\n|\s+)\s*([0-9A-Za-z\-]+)'
    ], text)

    # 5. Meter Number
    meter_number = extract_field([
        r'Meter Number\s*\n\s*([0-9A-Za-z\-]+)',
        r'Meter\s*(?:Number|No):?\s*(?:\n|\s+)\s*([0-9A-Za-z\-]+)'
    ], text)

    # 6. Bill Period
    bill_period = extract_field([
        r'(\d{2}/\d{2}/\d{4}\s*-\s*\d{2}/\d{2}/\d{4}(?:\s*\(\d+\s*days\))?)',
        r'Bill\s*Period\s*(?:\n|\s+)\s*([0-9/]+\s*-\s*[^\n]+?(?:\(\d+\s*days\))?)',
        r'Bill Period\s*\n\s*([^\n\r]+)'
    ], text)

    # 7. Bill Number
    bill_number = extract_field([
        r'Bill Number\s*\n\s*([0-9A-Za-z\-]+)',
        r'Bil[lt]\s*(?:Number|No):?\s*(?:\n|\s+)\s*([0-9A-Za-z\-]+)'
    ], text)

    # 8. Issuing Center
    issuing_center = extract_field([
        r'Issuing Center\s*\n\s*([^\n\r]+)',
        r'Issuing\s*Center\s*(?:\n|\s+)\s*([A-Za-z\s]+?)(?=\s+Due\s*Date|\s+Issue\s*Date|\n|$)'
    ], text, default="Unknown District")
    # Clean OCR noise in center if any
    if issuing_center:
        issuing_center = re.sub(r'[\r\n].*', '', issuing_center).strip()

    # 9. Issue Date
    issue_date = extract_field([
        r'Issue Date\s*\n\s*(\d{2}/\d{2}/\d{4})',
        r'Issue\s*Date:?\s*(?:\n|\s+)\s*(\d{2}/\d{2}/\d{4})'
    ], text)

    # 10. Due Date
    due_date = extract_field([
        r'Due Date\s*\n\s*(\d{2}/\d{2}/\d{4})',
        r'Due\s*Date\s*(?:\n|\s+)\s*(\d{2}/\d{2}/\d{4})',
        r'Payment\s+due\s+by\s*(\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+,?\s+\d{4})'
    ], text)

    # 11 & 12. Previous Reading and Current Reading
    prev_reading = None
    curr_reading = None

    # Try paired match first (common in both digital & OCR)
    m_paired = re.search(r'Previous Reading\s*\n?\s*Current Reading\s*\n?\s*([\d,\.]+)\s*(?:kWh|Mh)?\s*\n?\s*(?:\([A-Za-z0-9]+\)\s*)?([\d,\.]+)\s*(?:kWh|Mh)?', text, re.I)
    if m_paired:
        prev_reading = clean_number(m_paired.group(1))
        curr_reading = clean_number(m_paired.group(2))
    else:
        prev_reading_raw = extract_field([
            r'Previous Reading\s*\n\s*([\d,\.]+)',
            r'Previous\s*Reading\s*(?:\n|\s+)\s*([\d,\.]+)'
        ], text)
        curr_reading_raw = extract_field([
            r'Current Reading\s*\n\s*([\d,\.]+)',
            r'Current\s*Reading\s*(?:\n|\s+)\s*([\d,\.]+)'
        ], text)
        prev_reading = clean_number(prev_reading_raw)
        curr_reading = clean_number(curr_reading_raw)

    # 13. Consumption
    consumption_raw = extract_field([
        r'Consumption\s*\n\s*([\d,\.]+)\s*kWh',
        r'Consumption\s*(?:\n|\s+)\s*([\d,\.]+)\s*(?:kWh|Mh)?',
        r'September\s+\d{4}\s+Consumption\s*\n?\s*([\d,\.]+)\s*kWh'
    ], text)
    consumption = clean_number(consumption_raw)

    # 14 & 15. Current Bill & Previous Balance
    current_bill = None
    previous_balance = None

    # In digital text or standard OCR:
    cb_raw = extract_field([
        r'Current Bill\s*(?:GHS)?\s*([-\d,\.]+)',
        r'[A-Za-z]+\s+\d{4}\s+Bill\s+GHS\s*([-\d,\.]+)'
    ], text)
    current_bill = clean_number(cb_raw)

    pb_raw = extract_field([
        r'Previous Balance\s*(?:GHS)?\s*([-\d,\.]+)'
    ], text)
    previous_balance = clean_number(pb_raw)

    # In scanned OCR where amounts appear in block: GHS A \n GHS B \n GHS C \n GHS D
    if current_bill is None or previous_balance is None:
        m_fin = re.search(r'GHS\s*([\d,\.]+)\s*\n\s*GHS\s*([-\d,\.]+)\s*\n\s*GHS\s*([-\d,\.]+)\s*\n\s*GHS\s*([-\d,\.]+)', text)
        if m_fin:
            if current_bill is None:
                current_bill = clean_number(m_fin.group(1))
            if previous_balance is None:
                previous_balance = clean_number(m_fin.group(2))

    # 16. Amount Payable
    amount_payable_raw = extract_field([
        r'Amount\s*Payable\s*(?:Amount\s*Payable\s*)?(?:GHS|CHS)?\s*([-\d,\.]+)',
        r'Amount Payable\s*\n\s*GHS\s*([-\d,\.]+)',
        r'(?:GHS|CHS)?\s*([-\d,\.]+)\s*\n\s*Payment due by'
    ], text)
    amount_payable = clean_number(amount_payable_raw)

    return {
        "File Name": filename,
        "Customer Name": customer_name,
        "Phone Number": phone_number,
        "Digital Address": digital_address,
        "Account Number": account_number,
        "Meter Number": meter_number,
        "Bill Period": bill_period,
        "Bill Number": bill_number,
        "Issuing Center": issuing_center,
        "Issue Date": issue_date,
        "Due Date": due_date,
        "Previous Reading": prev_reading,
        "Current Reading": curr_reading,
        "Consumption": consumption,
        "Current Bill": current_bill,
        "Previous Balance": previous_balance,
        "Amount Payable": amount_payable,
    }


async def _ocr_page_async(page) -> str:
    """Render a PyMuPDF page to image and OCR it using winocr."""
    if not HAS_WINOCR:
        return ""
    try:
        pix = page.get_pixmap(dpi=150)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        res = await winocr.recognize_pil(img, "en")
        return "\n".join([l.text for l in res.lines])
    except Exception as e:
        print(f"[!] OCR error on page: {e}")
        return ""


async def _process_pdf_doc_async(doc: pymupdf.Document, filename: str) -> List[Dict[str, Any]]:
    """Process all pages of a PyMuPDF document, performing OCR on scanned pages asynchronously."""
    num_pages = len(doc)
    page_texts: List[str] = [""] * num_pages

    ocr_tasks = []
    ocr_indices = []

    for i in range(num_pages):
        page = doc[i]
        text = page.get_text()
        if len(text.strip()) > 50 and any(k in text.lower() for k in ["customer", "bill", "electricity", "account"]):
            page_texts[i] = text
        else:
            # Needs OCR
            ocr_indices.append(i)
            ocr_tasks.append(_ocr_page_async(page))

    if ocr_tasks:
        ocr_results = await asyncio.gather(*ocr_tasks)
        for idx, ocr_text in zip(ocr_indices, ocr_results):
            page_texts[idx] = ocr_text

    invoices: List[Dict[str, Any]] = []
    for i, p_text in enumerate(page_texts):
        if not p_text.strip():
            continue
        page_label = f"{filename} (Page {i + 1})" if num_pages > 1 else filename
        inv = parse_invoice_text(p_text, filename=page_label)
        # Check if page has meaningful invoice content
        if inv.get("Customer Name") or inv.get("Account Number") or inv.get("Meter Number") or inv.get("Amount Payable") is not None:
            invoices.append(inv)

    return invoices


def extract_from_pdf_path(pdf_path: str) -> List[Dict[str, Any]]:
    """Extract all invoices across all pages of a PDF file on disk."""
    filename = os.path.basename(pdf_path)
    doc = pymupdf.open(pdf_path)
    try:
        return asyncio.run(_process_pdf_doc_async(doc, filename))
    finally:
        doc.close()


def extract_from_pdf_stream(stream, filename: str = "") -> List[Dict[str, Any]]:
    """Extract all invoices across all pages from a byte stream."""
    stream_bytes = stream.read()
    doc = pymupdf.open(stream=stream_bytes, filetype="pdf")
    try:
        return asyncio.run(_process_pdf_doc_async(doc, filename or "uploaded.pdf"))
    finally:
        doc.close()


if __name__ == "__main__":
    multi_file = r"c:\Users\wwukpo\Desktop\ECG BILLS\ecg00.pdf"
    if os.path.exists(multi_file):
        print(f"Testing extraction on {multi_file}...")
        results = extract_from_pdf_path(multi_file)
        print(f"Extracted {len(results)} invoices from {multi_file}:")
        for inv in results[:5]:
            print(f"  [{inv['File Name']}] {inv['Customer Name']} | {inv['Issuing Center']} | GHS {inv['Amount Payable']}")
