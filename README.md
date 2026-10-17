# ⚡ ECG Invoice Data Extractor & Excel Exporter

A modern, fast web application and command-line batch tool to extract structured data from Electricity Company of Ghana (ECG) PDF invoices and automatically organize them into formatted, downloadable Excel worksheets grouped by **Issuing Center**.

---

## 🎯 Features

- **Extracted Fields (15 Required Data Points + Metadata)**:
  1. `Customer Name`
  2. `Phone Number` (stored as text with leading zeros preserved, e.g., `0549517857`)
  3. `Digital Address` (e.g., `N/A` or GhanaPost GPS address)
  4. `Account Number`
  5. `Meter Number`
  6. `Bill Period` (e.g., `01/09/2026 - 30/09/2026 (30 days)`)
  7. `Bill Number`
  8. `Issuing Center` (e.g., `Manso Nkwanta District`)
  9. `Issue Date`
  10. `Due Date`
  11. `Previous Reading` (Clean numeric format `#,##0`)
  12. `Current Reading` (Clean numeric format `#,##0`)
  13. `Consumption` (kWh)
  14. `Current Bill` (GHS, numeric currency `#,##0.00`)
  15. `Previous Balance` (GHS, numeric currency `#,##0.00`)
  16. `Amount Payable` (GHS, numeric currency `#,##0.00;[Red]-#,##0.00`)
  17. `File Name` (source tracking)

- **Smart Grouping & Fuzzy Similarity**:
  - Invoices with the same or similar Issuing Center (or Customer Name) are automatically merged and appended onto the same worksheet using similarity matching (e.g. minor typos, abbreviation differences like `Dist.` vs `District` or case variations are normalized).
  - Each worksheet is named after the **Issuing Center** (or Customer).
  - An overarching **All Bills Summary** tab is included when multiple centers exist.
  - Interactive **Group Sheets By** selector in the UI:
    - *Issuing Center* (Default)
    - *Customer Name*
    - *Customer + Issuing Center*
  - Automated `=SUM(...)` formulas at the bottom of numeric columns (Consumption, Current Bill, Previous Balance, Amount Payable).
  - Frozen header rows, auto-sized columns, and professional corporate navy styling.

- **Interactive Editing**:
  - Click the eye icon next to any bill in the preview table to view and edit its fields directly before exporting.

- **Dual Access Modes**:
  - **Modern Web App**: Drag & drop files, scan folder (`c:\Users\wwukpo\Desktop\ECG BILLS`), live preview table with center tabs, KPI counters, search bar, and one-click Excel download.
  - **CLI Batch Script**: Command-line batch runner for automated headless directory processing (`--group-by center|customer|both`).

---

## 🚀 How to Launch the Web App

### Option A: Double-Click the Launcher
Simply double-click the included:
```
run_app.bat
```
This automatically starts the local server and opens your browser to `http://localhost:5000`.

### Option B: Terminal Command
```bash
python app.py
```
Then visit [http://localhost:5000](http://localhost:5000) in your browser.

---

## 💻 How to Use the CLI Batch Script

To scan a folder and generate the Excel file without opening a browser:
```bash
python batch_process.py --folder "c:\Users\wwukpo\Desktop\ECG BILLS"
```
The output file `ECG_Extracted_Invoices.xlsx` will be generated directly inside the folder.
