#!/usr/bin/env python3
"""
SP Shah & Co. - Single-File Bundler
Compiles src/styles.css, src/audit_engine.js, and vendor libraries
into the distribution-ready SP_Shah_Job_Register_Auditor.html file.
"""

import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(BASE_DIR, 'src')
VENDOR_DIR = os.path.join(BASE_DIR, 'vendor')

if not os.path.exists(VENDOR_DIR):
    VENDOR_DIR = '/tmp/vendor'

sheetjs_file = os.path.join(VENDOR_DIR, 'xlsx.full.min.js')
chartjs_file = os.path.join(VENDOR_DIR, 'chart.umd.min.js')

for f_path in (sheetjs_file, chartjs_file):
    if not os.path.exists(f_path):
        print(f"Error: Missing vendor file: {f_path}")
        sys.exit(1)

with open(sheetjs_file, 'r', encoding='utf-8') as f:
    sheetjs_code = f.read()

with open(chartjs_file, 'r', encoding='utf-8') as f:
    chartjs_code = f.read()

with open(os.path.join(SRC_DIR, 'styles.css'), 'r', encoding='utf-8') as f:
    css_code = f.read()

with open(os.path.join(SRC_DIR, 'audit_engine.js'), 'r', encoding='utf-8') as f:
    js_code = f.read()

html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>SP Shah | Import Job Register Audit & Late Billing Analyzer</title>
  <style>
{css_code}
  </style>
  <script>
    // SheetJS (xlsx.full.min.js v0.18.5)
{sheetjs_code}
  </script>
  <script>
    // Chart.js (chart.umd.min.js v4.4.1)
{chartjs_code}
  </script>
</head>
<body>
  <header>
    <div class="brand">
      <div class="brand-icon">SP</div>
      <div class="brand-text">
        <h1>SP SHAH & CO. — JOB REGISTER AUDITOR</h1>
        <p>Customs Broker & Logistics Late Billing Compliance System (TAT SLA > 2 Days)</p>
      </div>
    </div>
    <div class="header-actions">
      <button id="btn-export-excel" class="btn btn-success" disabled>
        📥 Export Audit Workbook (.xlsx)
      </button>
      <button id="btn-export-csv" class="btn btn-outline" disabled>
        📄 Export Late Jobs (CSV)
      </button>
      <button id="btn-clear" class="btn btn-outline" style="display:none;">
        ✕ Clear
      </button>
    </div>
  </header>

  <div class="container">
    <div id="upload-zone" class="upload-card">
      <input type="file" id="file-input" accept=".xlsx,.xls,.csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">
      <div class="upload-icon">📂</div>
      <div class="upload-title">Drop your Job Register File (.xlsx or legacy .xls) here</div>
      <div class="upload-desc">
        Processes directly in your browser. 100% private & secure — no data leaves this machine.
        Supports raw Crystal Decisions .xls and formatted .xlsx job registers.
      </div>
      <div class="badge-bar">
        <span class="chip">✓ Group Sr. No Engine</span>
        <span class="chip">✓ Auto Header Detection</span>
        <span class="chip">✓ Late Billing (&gt; 2 Days) Flagging</span>
        <span class="chip">✓ 7-Sheet Audit Report Export</span>
        <span class="chip">✓ Offline Standalone</span>
      </div>
    </div>

    <div id="status-bar" class="status-bar">
      <span id="status-text">Ready. Please upload an import job register.</span>
      <span id="status-time"></span>
    </div>

    <div id="kpi-section" class="kpi-grid" style="display:none;">
      <div class="kpi-card kpi-primary">
        <div class="kpi-title">Total Import Jobs</div>
        <div id="kpi-total" class="kpi-val">0</div>
        <div class="kpi-sub"><span id="kpi-total-sub">Grouped client entries</span></div>
      </div>
      <div class="kpi-card kpi-danger">
        <div class="kpi-title">Late Billing (> 2 Days)</div>
        <div id="kpi-late" class="kpi-val" style="color:var(--danger);">0</div>
        <div class="kpi-sub">
          <span id="kpi-late-pct" class="kpi-badge badge-danger">0%</span>
          <span>SLA Non-Compliant</span>
        </div>
      </div>
      <div class="kpi-card kpi-success">
        <div class="kpi-title">On-Time Billing (≤ 2d)</div>
        <div id="kpi-ontime" class="kpi-val" style="color:var(--success);">0</div>
        <div class="kpi-sub">
          <span id="kpi-ontime-pct" class="kpi-badge badge-success">0%</span>
          <span>SLA Compliant</span>
        </div>
      </div>
      <div class="kpi-card kpi-warning">
        <div class="kpi-title">Critical Unbilled</div>
        <div id="kpi-unbilled" class="kpi-val" style="color:var(--warning);">0</div>
        <div class="kpi-sub">
          <span class="kpi-badge badge-warning">Revenue Risk</span>
          <span id="kpi-unbilled-sub">Delivered, missing bill</span>
        </div>
      </div>
      <div class="kpi-card kpi-info">
        <div class="kpi-title">Total Customs Duty</div>
        <div id="kpi-duty" class="kpi-val">₹0 Cr</div>
        <div class="kpi-sub">
          <span>Assessable: </span>
          <span id="kpi-assessable">₹0 Cr</span>
        </div>
      </div>
      <div class="kpi-card kpi-purple">
        <div class="kpi-title">Avg Billing TAT</div>
        <div id="kpi-tat" class="kpi-val">0.0 d</div>
        <div class="kpi-sub">
          <span>Target SLA: ≤ 2.0 Days</span>
        </div>
      </div>
    </div>

    <div id="charts-section" class="charts-row" style="display:none;">
      <div class="chart-card">
        <div class="chart-header">
          <div class="chart-title">Billing SLA Compliance Distribution</div>
        </div>
        <div class="chart-box">
          <canvas id="chart-status"></canvas>
        </div>
      </div>
      <div class="chart-card">
        <div class="chart-header">
          <div class="chart-title">Turnaround Aging Buckets (Days)</div>
        </div>
        <div class="chart-box">
          <canvas id="chart-aging"></canvas>
        </div>
      </div>
      <div class="chart-card">
        <div class="chart-header">
          <div class="chart-title">Top 8 Clients with Highest Late Billing</div>
        </div>
        <div class="chart-box">
          <canvas id="chart-clients"></canvas>
        </div>
      </div>
    </div>

    <div id="table-section" class="table-card" style="display:none;">
      <div class="table-toolbar">
        <div class="filter-pills">
          <button class="pill active" data-filter="all">All Jobs (<span id="count-all">0</span>)</button>
          <button class="pill" data-filter="late">Late Billing (> 2d) (<span id="count-late">0</span>)</button>
          <button class="pill" data-filter="ontime">On-Time (≤ 2d) (<span id="count-ontime">0</span>)</button>
          <button class="pill" data-filter="unbilled">Unbilled (<span id="count-unbilled">0</span>)</button>
          <button class="pill" data-filter="advance">Advance Billed (<span id="count-advance">0</span>)</button>
          <button class="pill" data-filter="anomaly">Date Anomalies (<span id="count-anomaly">0</span>)</button>
        </div>
        <div class="search-box">
          <span class="search-icon">🔍</span>
          <input type="text" id="search-input" placeholder="Search Client, Job No, Bill No...">
        </div>
      </div>

      <div class="table-responsive">
        <table>
          <thead>
            <tr>
              <th data-sort="group_sr_no">Group Sr.No ↕</th>
              <th data-sort="job_no">Job No ↕</th>
              <th data-sort="client_name">Client Name & Address ↕</th>
              <th data-sort="delivery_date">Delivery Date ↕</th>
              <th data-sort="bill_date">Bill No & Date ↕</th>
              <th data-sort="diff_days" style="text-align:center;">Difference ↕</th>
              <th data-sort="billing_status">Status ↕</th>
              <th>Auditor Comments</th>
              <th data-sort="duty_paid" style="text-align:right;">Customs Duty (₹) ↕</th>
            </tr>
          </thead>
          <tbody id="table-body">
          </tbody>
        </table>
      </div>

      <div class="table-footer">
        <div id="page-info" class="page-info">Showing 0 of 0 jobs</div>
        <div class="page-controls">
          <button id="btn-prev" class="btn btn-secondary btn-sm" disabled>‹ Previous</button>
          <span id="page-current" style="font-weight:700; margin: 0 8px;">1</span>
          <button id="btn-next" class="btn btn-secondary btn-sm" disabled>Next ›</button>
        </div>
      </div>
    </div>
  </div>

  <div id="detail-modal" class="modal-backdrop">
    <div class="modal-card">
      <div class="modal-header">
        <div>
          <h2 id="modal-client" style="font-size:18px; font-weight:800;">Client Detail</h2>
          <p id="modal-sub" style="font-size:12px; color:var(--text-muted);"></p>
        </div>
        <button id="modal-close-btn" class="modal-close">&times;</button>
      </div>

      <div id="modal-finding-box" class="audit-finding-box">
        <div id="modal-finding-title" class="audit-finding-title">Late Billing Audit Observation</div>
        <div id="modal-finding-desc" class="audit-finding-desc"></div>
      </div>

      <div class="detail-grid">
        <div class="detail-item">
          <div class="detail-label">Group Sr. No / Sr. No</div>
          <div id="m-group-sr" class="detail-value"></div>
        </div>
        <div class="detail-item">
          <div class="detail-label">Job No / Recd Date</div>
          <div id="m-job-recd" class="detail-value"></div>
        </div>
        <div class="detail-item">
          <div class="detail-label">Delivery Date (Title Row)</div>
          <div id="m-delivery-date" class="detail-value" style="font-weight:700; color:var(--accent);"></div>
        </div>
        <div class="detail-item">
          <div class="detail-label">Bill No & Bill Date (Sub-Row)</div>
          <div id="m-bill-info" class="detail-value" style="font-weight:700;"></div>
        </div>
        <div class="detail-item">
          <div class="detail-label">Turnaround Difference</div>
          <div id="m-diff" class="detail-value"></div>
        </div>
        <div class="detail-item">
          <div class="detail-label">Aging Category</div>
          <div id="m-aging" class="detail-value"></div>
        </div>
        <div class="detail-item">
          <div class="detail-label">B/L No & B/L Date</div>
          <div id="m-bl-info" class="detail-value"></div>
        </div>
        <div class="detail-item">
          <div class="detail-label">Invoice Details</div>
          <div id="m-invo-info" class="detail-value"></div>
        </div>
        <div class="detail-item">
          <div class="detail-label">Customs Duty Paid</div>
          <div id="m-duty" class="detail-value"></div>
        </div>
        <div class="detail-item">
          <div class="detail-label">Assessable Value</div>
          <div id="m-assessable" class="detail-value"></div>
        </div>
      </div>

      <div class="detail-item" style="margin-bottom:16px;">
        <div class="detail-label">Client Address (Extracted from Job Entry)</div>
        <div id="m-address" class="detail-value" style="font-size:12.5px;"></div>
      </div>
    </div>
  </div>

  <script>
{js_code}
  </script>
</body>
</html>
"""

dist_file = os.path.join(BASE_DIR, 'SP_Shah_Job_Register_Auditor.html')
with open(dist_file, 'w', encoding='utf-8') as f:
    f.write(html_template)

print(f"Build complete: {dist_file} ({len(html_template):,} bytes)")
