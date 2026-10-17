import os
import io
import glob
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename

from extractor import extract_from_pdf_path, extract_from_pdf_stream
from excel_builder import generate_excel_bytes

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB max upload

DEFAULT_SAMPLE_DIR = r"c:\Users\wwukpo\Desktop\ECG BILLS"


@app.after_request
def add_header(response):
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.route("/")
def index():
    return render_template("index.html", default_folder=DEFAULT_SAMPLE_DIR)


@app.route("/api/upload", methods=["POST"])
def upload_files():
    """Process uploaded PDF files."""
    if 'files' not in request.files:
        return jsonify({"error": "No files provided"}), 400

    uploaded_files = request.files.getlist('files')
    if not uploaded_files or (len(uploaded_files) == 1 and uploaded_files[0].filename == ''):
        return jsonify({"error": "No selected file"}), 400

    results = []
    errors = []

    for file in uploaded_files:
        if not file.filename.lower().endswith('.pdf'):
            continue
        try:
            filename = secure_filename(file.filename) or file.filename
            invoices = extract_from_pdf_stream(file.stream, filename=filename)
            results.extend(invoices)
        except Exception as e:
            errors.append({"file": file.filename, "error": str(e)})

    return jsonify({
        "success": True,
        "count": len(results),
        "invoices": results,
        "errors": errors
    })


@app.route("/api/scan-folder", methods=["POST"])
def scan_folder():
    """Scan a local directory path for PDF files and extract them."""
    data = request.get_json() or {}
    folder_path = data.get("folder_path", "").strip()

    if not folder_path:
        folder_path = DEFAULT_SAMPLE_DIR

    if not os.path.exists(folder_path):
        return jsonify({"error": f"Folder not found: {folder_path}"}), 404

    pdf_files = glob.glob(os.path.join(folder_path, "*.pdf"))
    if not pdf_files:
        return jsonify({"error": f"No PDF files found in folder: {folder_path}"}), 404

    results = []
    errors = []

    for pdf_path in pdf_files:
        try:
            invoices = extract_from_pdf_path(pdf_path)
            results.extend(invoices)
        except Exception as e:
            errors.append({"file": os.path.basename(pdf_path), "error": str(e)})

    return jsonify({
        "success": True,
        "folder": folder_path,
        "count": len(results),
        "invoices": results,
        "errors": errors
    })


@app.route("/api/sample", methods=["GET"])
def load_sample():
    """Load the default sample bill if present."""
    sample_file = os.path.join(DEFAULT_SAMPLE_DIR, "ABORE HOUSE ELECTRICITY BILL.pdf")
    if not os.path.exists(sample_file):
        # Look for any pdf in directory
        pdfs = glob.glob(os.path.join(DEFAULT_SAMPLE_DIR, "*.pdf"))
        if pdfs:
            sample_file = pdfs[0]
        else:
            return jsonify({"error": "Sample file not found"}), 404

    try:
        invoices = extract_from_pdf_path(sample_file)
        return jsonify({
            "success": True,
            "invoices": invoices,
            "count": len(invoices)
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/export-excel", methods=["POST"])
def export_excel():
    """Generate and download Excel file from provided invoice data."""
    payload = request.get_json() or {}
    invoices = payload.get("invoices", [])
    group_by = payload.get("group_by", "center")  # "center", "customer", or "both"

    if not invoices:
        return jsonify({"error": "No invoice data provided to export"}), 400

    try:
        excel_bytes = generate_excel_bytes(invoices, group_by=group_by)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"ECG_Invoices_Extracted_{timestamp}.xlsx"

        return send_file(
            io.BytesIO(excel_bytes),
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": f"Failed to generate Excel: {str(e)}"}), 500


if __name__ == "__main__":
    print("Starting ECG Invoice Extractor Server on http://localhost:5000 ...")
    app.run(host="0.0.0.0", port=5000, debug=False)
