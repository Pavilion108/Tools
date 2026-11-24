# SP Shah & Co. — Import Job Register Auditor & Late Billing TAT Analyzer

An auditor-grade analytics and automation system built specifically for **SP Shah & Co.** (Customs Broker / CHA / Logistics) to process import port Job Registers (both legacy Crystal Decisions `.xls` and formatted `.xlsx` workbooks), automatically group client entries, evaluate billing turnaround time (TAT), detect SLA breaches, and generate comprehensive multi-sheet Excel audit reports.

---

## 🎯 Core Audit Rules & Logic

1. **Turnaround Difference Formula**:
   $$\text{Difference (Days)} = \text{Bill Date} - \text{Delivery Date}$$

2. **Late Billing SLA Breach**:
   - Flagged whenever $\text{Difference} > 2\text{ Days}$.
   - Evaluated as non-compliant against the firm's 2-day billing TAT SLA.

3. **Status Classifications & Audit Categories**:
   - **On-Time Billing** ($\le 2\text{ Days}$): Billed within compliance benchmark.
   - **Late Billing** ($> 2\text{ Days}$): Delays working capital; flagged for auditor review.
   - **Critical Unbilled Deliveries**: Cargo delivered on record, but no bill date found (immediate cash flow and revenue leakage risk).
   - **Advance Billed** ($< 0\text{ Days}$): Bill raised before cargo delivery.
   - **Data Integrity Anomalies** ($> 100\text{ Days}$ or $<-60\text{ Days}$): Input typographical errors (e.g. year `2020` instead of `2026`).

---

## 🧭 Smart Job Register Parsing & Column Orientation

The engine auto-detects single- or multi-row headers across the first 20 rows of any uploaded workbook using fuzzy keyword matching:

| Field | Source Header Match | Row Level | Extraction Rule |
| :--- | :--- | :--- | :--- |
| **Group Sr. No** | Auto-Generated Counter | All Rows | Sequential integer (`1, 2, ... 5074`) assigned across all rows of a client entry. |
| **Sr. No** | `Sr.No` | Title Row | An integer on the title row (`1, 2, 3...`). Blank rows below belong to that entry until the next integer appears. |
| **Job No** | `JobNo \| Recd. Dt.` | Title Row | Port Job Number (e.g. `1`, `2`). |
| **Recd. Date** | `JobNo \| Recd. Dt.` | Subordinate Row | Cargo receipt date parsed from the date in the Job No column. |
| **Client Name** | `Client Name \| Address` | Title Row | Name of the importer/client. |
| **Client Address** | `Client Name \| Address` | Subordinate Row | Physical address parsed from the client column on the row where `JobNo` contains a date. |
| **Delivery Date** | Column matching `Delivery`, `Date`, `IGM Date`, `BE Date` | **Title Row (Col G)** | Cargo delivery date (e.g. `14/05/2026`). |
| **Bill Date** | Column with `Bill No.` and `Date` | **Subordinate Row (Col M)** | Billing invoice date (e.g. `15/05/2026`). |
| **B/L & Invoice No** | `B/LNo \| Invo. No.` | Title & Sub Rows | B/L number on title row; invoice details (`@` syntax) on sub-row. |
| **Customs Duty & Val** | `Duty Paid`, `Invo. Val.` | Title & Sub Rows | Duty Paid and Assessable / CIF values. |

---

## 🚀 Quick Start — Standalone Single-File Browser App

Double-click or drag **`SP_Shah_Job_Register_Auditor.html`** into any modern web browser (Google Chrome, Microsoft Edge, Brave, Firefox, Safari).

- **Zero Installation**: No Python, Node.js, or server required.
- **100% Client-Side Privacy**: All processing runs locally in browser memory. No data is sent to external servers.
- **Bi-Format Compatibility**: Embedded **SheetJS v0.18.5** handles both formatted `.xlsx` (OpenXML) and legacy `.xls` (BIFF8 / Crystal Decisions).
- **Interactive Visuals**: Embedded **Chart.js v4.4.1** renders SLA compliance donuts, TAT aging buckets, and top client delay charts.
- **1-Click Multi-Sheet Export**: Generates a 7-sheet professional audit workbook `.xlsx`.

---

## 📊 Exported Audit Workbook Structure (7 Sheets)

When clicking **"Export Audit Workbook (.xlsx)"**, the engine generates:

1. **`Late_Billing_Register`**:
   The primary register containing dedicated columns: `Group Sr. No`, `Sr. No`, `Job No`, `Client Name`, `Client Address`, `B/L No`, `B/L Date`, `Recd Date`, `Invoice No & Date`, `Delivery Date`, `Bill No`, `Bill Date`, `Difference (Days)`, `Billing Status`, `Aging Category`, `Auditor Comments`, `Customs Duty Paid (INR)`, and `Assessable Value (INR)`.
2. **`Executive_Audit_Summary`**:
   High-level KPI metrics, SLA compliance rates, total duty exposure, and executive findings.
3. **`Client_Revenue_TAT_Analysis`**:
   Pivot of all clients ranked by late billing volume, average TAT days, total customs duty, assessable turnover, and assigned risk levels.
4. **`Billing_Aging_Drilldown`**:
   Bucketed aging distribution (`0-2d On-Time`, `3-5d Low`, `6-10d Moderate`, `11-20d Significant`, `21-30d High`, `>30d Severe`, `Unbilled`, `Advance`, `Anomalies`).
5. **`Unbilled_Revenue_Exposure`**:
   Critical exception report isolating delivered jobs with missing bill dates (revenue leakage risk).
6. **`Data_Integrity_Anomalies`**:
   Date inconsistencies and typographical errors flagged for data correction.
7. **`Enriched_Raw_Register`**:
   Full raw register enriched with `Group Sr. No`, parsed dates, difference, status, and auditor comments.

---

## 🛠️ Modifying & Extending the Tool

The project is structured with clean, human-readable source code so you or AI assistants can easily make changes:

```
Tools/
├── SP_Shah_Job_Register_Auditor.html  # Standalone single-file distribution (built)
├── build.py                           # Python bundler script
├── README.md                          # Documentation & business rules
├── .gitignore                         # Protects confidential data files
├── src/                               # Clean, editable frontend source code
│   ├── styles.css                     # UI stylesheet & responsive theme
│   └── audit_engine.js                # Core parsing, grouping & export engine
├── vendor/                            # Embedded client libraries
│   ├── xlsx.full.min.js               # SheetJS v0.18.5
│   └── chart.umd.min.js               # Chart.js v4.4.1
└── python/                            # Python engine & CLI utilities
    ├── job_register_processor.py      # Batch openpyxl/xlrd processor
    ├── run_audit.py                   # CLI audit runner
    ├── app.py                         # FastAPI web dashboard
    ├── run_tool.sh                    # CLI runner shell script
    └── start_web_ui.sh                # Web server shell script
```

### To customize logic or UI:
1. Edit **`src/styles.css`** to change styling, colors, or badges.
2. Edit **`src/audit_engine.js`** to change SLA thresholds (e.g. modify `diffDays > 2`), add new columns, or adjust auditor comments.
3. Re-bundle the standalone HTML file by running:
   ```bash
   python3 build.py
   ```
4. Commit your changes to git and push!

---

## 🐍 Headless Python Processing (Alternative)

For automated server pipelines or scheduled cron jobs:

```bash
cd python/
python3 run_audit.py /path/to/Job_Register.xlsx
```
Or start the FastAPI local dashboard:
```bash
cd python/
./start_web_ui.sh
```

---

## 🔒 Code Security & Protection Advice

If you distribute this tool to team members or external clients and want to prevent unauthorized tampering or programmatic alterations:

1. **Client-Side Obfuscation**: You can run an obfuscator (e.g. `javascript-obfuscator src/audit_engine.js --compact true --self-defending true`) before bundling into `dist/`.
2. **Server-Side API Option**: Use the provided `python/app.py` FastAPI server where the parsing engine and business logic run on a private server, exposing only read-only views to the browser.
3. **Integrity Checksums**: Keep this git repository as your immutable single source of truth (`main` branch) to track all revisions.
