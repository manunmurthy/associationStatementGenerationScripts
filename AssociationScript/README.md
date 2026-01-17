# HDFC Bank Statement to Excel Converter

Convert HDFC bank statements (PDF or Excel) to professional Excel reports with transaction analysis and flat number identification.

## Quick Start

### 1. Install Dependencies
```bash
pip install -r requirements.txt
# Optional (PDF export):
pip install reportlab
```

### 2. Place Input File
Put your HDFC bank statement file in the `input/` folder:
- PDF: `*.pdf`
- Excel: `*.xls` or `*.xlsx`

### 3. Run the Script
```bash
python hdfc_complete_converter.py
```

Choose your input format (PDF or Excel) when prompted.

## Output
The script generates an Excel file in the `output/` folder and (optionally) a PDF per sheet if `reportlab` is installed. The workbook contains these sheets:
- **Summary**: Transaction overview and flat breakdown
- **Credits**: All credit transactions with flat numbers
- **Debits**: All debit transactions
- **Entries with Flats**: All transactions that have a detected flat number (one row per transaction). Note: the Balance column was removed from this sheet to keep entries concise.
- **Flat Details**: Payment breakdown by flat (payments in separate columns, totals)

Each sheet is also exported to a separate PDF named like `HDFC_Complete_Statement_YYYYMMDD_HHMMSS_<SheetName>.pdf` if `reportlab` is available.

## New / Notable Behavior
- "Entries with Flats" lists flats in the configured order (A-blocks, then B-blocks, then C-blocks). For flats with no transactions a placeholder row is added.
- "Flat Details" groups payments per flat and shows totals and multi-payment detection.
- PDF export is automatic when `reportlab` is installed. If not installed, PDFs will be skipped and only Excel is saved.

## Features
✅ Dual input support (PDF & Excel)  
✅ Automatic flat number detection (A001-A320, B001-B312, C001-C318)  
✅ Credit & debit separation  
✅ Professional Excel formatting  
✅ Per-sheet PDF export (optional)  

## Project Structure
```
AssociationScript/
├── README.md
├── requirements.txt
├── hdfc_complete_converter.py    # Main script
├── src/                          # Core modules
├── config/                       # Bank format configs
├── input/                        # Place files here
└── output/                       # Generated Excel & PDFs
```

---
Built for efficient association management and financial tracking.
