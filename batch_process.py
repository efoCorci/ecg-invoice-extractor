import os
import glob
import argparse
from extractor import extract_from_pdf_path
from excel_builder import generate_excel_workbook

def process_directory(folder_path: str, output_excel: str, group_by: str = "center"):
    print(f"Scanning directory: {folder_path}")
    pdf_files = glob.glob(os.path.join(folder_path, "*.pdf"))
    if not pdf_files:
        print(f"No PDF files found in {folder_path}")
        return

    print(f"Found {len(pdf_files)} PDF file(s). Extracting data...")
    records = []
    for idx, pdf in enumerate(pdf_files, 1):
        try:
            invoices = extract_from_pdf_path(pdf)
            records.extend(invoices)
            print(f"     -> Found {len(invoices)} bill(s)")
            for data in invoices[:3]:
                print(f"        * [{data.get('File Name')}] {data.get('Customer Name')} | {data.get('Issuing Center')} | Amount: {data.get('Amount Payable')}")
            if len(invoices) > 3:
                print(f"        * ... and {len(invoices) - 3} more bill(s)")
        except Exception as e:
            print(f"     [!] Failed to extract {pdf}: {e}")

    if not records:
        print("No valid records extracted.")
        return

    print(f"\nGenerating Excel workbook grouped by {group_by.upper()}...")
    wb = generate_excel_workbook(records, group_by=group_by)
    wb.save(output_excel)
    print(f"Done! Excel file saved to:\n  {output_excel}")
    print(f"Worksheets created: {wb.sheetnames}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract ECG PDF bills and export to grouped Excel workbook.")
    parser.add_argument("--folder", default=r"c:\Users\wwukpo\Desktop\ECG BILLS", help="Folder containing PDF bills")
    parser.add_argument("--output", default="", help="Output Excel file path")
    parser.add_argument("--group-by", choices=["center", "customer", "both"], default="center", help="Grouping strategy for sheets (default: center)")
    args = parser.parse_args()

    out_file = args.output
    if not out_file:
        out_file = os.path.join(args.folder, "ECG_Extracted_Invoices.xlsx")

    process_directory(args.folder, out_file, group_by=args.group_by)
