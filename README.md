# SP Shah | Job Register Audit Tool

A standalone client-side HTML evaluation tool and analytical engine designed specifically for **SP Shah** to audit import Job Registers, evaluate billing turnaround time (TAT), identify Late Billing SLA breaches, and export formatted multi-sheet Excel audit reports.

---

## 🚀 Ready-to-Use Standalone Browser App

Open either of the following files directly in any web browser (Chrome, Edge, Firefox, Brave, Safari):

- **`SP Shah Import&Export-Job Audit.html`** *(or `index.html`)*
- **100% Client-Side & Private**: All processing happens locally in browser memory. No data is sent over the network.
- **Bi-Format Compatibility**: Supports both formatted `.xlsx` workbooks and raw legacy port Crystal Reports `.xls` (BIFF8) files.
- **Identical Interface & Output**: Retains the exact original SP Shah internal audit interface, layout, styling, and multi-sheet ExcelJS formatting.

---

## 🎯 Audit Business Rules & Evaluation

1. **Turnaround Difference**:
   $$\text{Days to Bill} = \text{Bill Date} - \text{Delivery Date}$$
   - **Delivery Date**: Detected from title row column matching `Delivery Date \n IGM Date \n BE Date` (Column G).
   - **Bill Date**: Detected from subordinate row under `Bill No. \n Date` (Column M).

2. **Late Billing Evaluation**:
   - **On-Time Billing** ($\le 2\text{ Days}$): Marked as `"Within 2-day threshold"`.
   - **Late Billing** ($> 2\text{ Days}$): Flagged in red with comment `"Late billing: bill date is X days after delivery"`.
   - **Review / Advance Billed** ($< 0\text{ Days}$): `"Review: bill date precedes delivery date"`.
   - **Missing Dates**: `"Review: missing or unreadable bill/delivery date"`.

---

## 📊 Exported Audit Workbook (`SP_Shah_Job_Register_Audit_Report.xlsx`)

When clicking **"Export audit workbook"**, ExcelJS generates a styled `.xlsx` report preserving the exact original formatting:

1. **`Late Billing`**: Primary register containing `Client Name`, `Group Sr. No.`, `Invoice / B/L No.`, `Delivery Date`, `Bill Date`, `Days to Bill`, `Status / Comments`, and `Source Title Row`. (Late rows highlighted in soft red `#FFFFF1F2` with bold red text `#FFB42332`).
2. **`Executive Summary`**: High-level audit metrics, grouped jobs count, late billing cases, late billing rate %, and revenue exposure.
3. **`Client Revenue Analysis`**: Distinct client breakdown by job volume, late billing rate %, and invoice value.
4. **`Billing Aging`**: Aging distribution of billing TAT.
5. **`Data Quality`**: Header detection notes, grouped jobs validation, and missing date checks.
6. **`Batch History`**: Session batch logs tracking changes between file uploads.
7. **`Source Data`**: Copy of the entire raw worksheet with styled header and auto-filters.

---

## 🛠️ Repository Structure

```
Tools/
├── SP Shah Import&Export-Job Audit.html   # Main standalone tool (ready to use)
├── index.html                            # Root web entrypoint
├── README.md                             # Documentation & business logic
├── .gitignore                            # Data protection rules
├── python/                               # Headless Python audit tools
│   ├── job_register_processor.py         # Batch openpyxl/xlrd processor
│   ├── run_audit.py                      # CLI runner
│   ├── app.py                            # FastAPI web server
│   ├── run_tool.sh                       # CLI launcher
│   └── start_web_ui.sh                   # Web dashboard launcher
└── vendor/                               # Embedded libraries
    ├── xlsx.full.min.js                  # SheetJS v0.18.5
    └── chart.umd.min.js                  # Chart.js v4.4.1
```
