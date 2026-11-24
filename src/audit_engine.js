(function() {
      // State
      let allJobs = [];
      let filteredJobs = [];
      let rawSheetRows = [];
      let currentFilter = 'all';
      let searchQuery = '';
      let sortCol = 'group_sr_no';
      let sortAsc = true;
      let currentPage = 1;
      const pageSize = 50;

      let chartStatus = null;
      let chartAging = null;
      let chartClients = null;

      // DOM Elements
      const fileInput = document.getElementById('file-input');
      const uploadZone = document.getElementById('upload-zone');
      const statusBar = document.getElementById('status-bar');
      const statusText = document.getElementById('status-text');
      const statusTime = document.getElementById('status-time');
      const kpiSection = document.getElementById('kpi-section');
      const chartsSection = document.getElementById('charts-section');
      const tableSection = document.getElementById('table-section');
      const tableBody = document.getElementById('table-body');
      const searchInput = document.getElementById('search-input');
      const btnExportExcel = document.getElementById('btn-export-excel');
      const btnExportCsv = document.getElementById('btn-export-csv');
      const btnClear = document.getElementById('btn-clear');
      const btnPrev = document.getElementById('btn-prev');
      const btnNext = document.getElementById('btn-next');
      const pageInfo = document.getElementById('page-info');
      const pageCurrent = document.getElementById('page-current');

      // Modal Elements
      const detailModal = document.getElementById('detail-modal');
      const modalCloseBtn = document.getElementById('modal-close-btn');

      // Drag and drop events
      ['dragenter', 'dragover'].forEach(name => {
        uploadZone.addEventListener(name, (e) => {
          e.preventDefault();
          uploadZone.classList.add('dragover');
        }, false);
      });

      ['dragleave', 'drop'].forEach(name => {
        uploadZone.addEventListener(name, (e) => {
          e.preventDefault();
          uploadZone.classList.remove('dragover');
        }, false);
      });

      uploadZone.addEventListener('drop', (e) => {
        const files = e.dataTransfer.files;
        if (files && files.length > 0) handleFile(files[0]);
      });

      fileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files.length > 0) handleFile(e.target.files[0]);
      });

      modalCloseBtn.addEventListener('click', () => detailModal.classList.remove('active'));
      detailModal.addEventListener('click', (e) => {
        if (e.target === detailModal) detailModal.classList.remove('active');
      });

      function showStatus(msg, time) {
        statusBar.classList.add('active');
        statusText.textContent = msg;
        statusTime.textContent = time ? '(' + time + ')' : '';
      }

      // File Reader
      function handleFile(file) {
        const t0 = performance.now();
        showStatus('Reading ' + file.name + '...');

        const reader = new FileReader();
        reader.onload = function(e) {
          try {
            showStatus('Parsing spreadsheet sheets & detecting client rows...');
            setTimeout(() => {
              const data = new Uint8Array(e.target.result);
              const wb = XLSX.read(data, { type: 'array', cellDates: true });
              const sheetName = wb.SheetNames[0];
              const sheet = wb.Sheets[sheetName];
              const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, dateNF: 'dd/mm/yyyy' });
              rawSheetRows = rows;

              const auditResult = processJobRegister(rows);
              allJobs = auditResult.jobs;

              const t1 = performance.now();
              const elapsedSec = ((t1 - t0) / 1000).toFixed(2);
              showStatus('Successfully parsed ' + allJobs.length.toLocaleString() + ' client jobs from ' + file.name, elapsedSec + 's');

              renderUI(auditResult);
              btnExportExcel.disabled = false;
              btnExportCsv.disabled = false;
              btnClear.style.display = 'inline-flex';
            }, 20);
          } catch (err) {
            console.error(err);
            alert('Failed to parse file: ' + err.message);
            showStatus('Error parsing file: ' + err.message);
          }
        };
        reader.readAsArrayBuffer(file);
      }

      // Audit Engine
      function parseDate(val) {
        if (val === null || val === undefined || val === '') return null;
        if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
        if (typeof val === 'number') {
          if (val >= 30000 && val <= 60000) {
            const d = new Date(Math.round((val - 25569) * 86400 * 1000));
            return isNaN(d.getTime()) ? null : d;
          }
        }
        let s = String(val).trim();
        if (!s || /^(none|null|nan|-)$/i.test(s)) return null;
        if (s.includes(' ')) s = s.split(' ')[0];

        const m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$/);
        if (m) {
          let day = parseInt(m[1], 10);
          let mon = parseInt(m[2], 10);
          let yr = parseInt(m[3], 10);
          if (yr < 100) yr += 2000;
          if (mon > 12 && day <= 12) {
            const t = day; day = mon; mon = t;
          }
          const d = new Date(yr, mon - 1, day);
          return isNaN(d.getTime()) ? null : d;
        }
        const iso = s.match(/^(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})$/);
        if (iso) {
          const yr = parseInt(iso[1], 10);
          const mon = parseInt(iso[2], 10);
          const day = parseInt(iso[3], 10);
          const d = new Date(yr, mon - 1, day);
          return isNaN(d.getTime()) ? null : d;
        }
        return null;
      }

      function parseNumber(val) {
        if (val === null || val === undefined) return 0;
        if (typeof val === 'number') return isNaN(val) ? 0 : val;
        const s = String(val).replace(/[, ₹$]/g, '').trim();
        const n = parseFloat(s);
        return isNaN(n) ? 0 : n;
      }

      function formatDate(d) {
        if (!d) return '';
        const day = String(d.getDate()).padStart(2, '0');
        const mon = String(d.getMonth() + 1).padStart(2, '0');
        const yr = d.getFullYear();
        return `${day}/${mon}/${yr}`;
      }

      function processJobRegister(rows) {
        let headerIdx = 0;
        let bestScore = -1;
        const keywords = ['sr.no', 'sr no', 'jobno', 'client', 'b/l', 'delivery', 'bill no', 'duty', 'invo'];

        for (let i = 0; i < Math.min(20, rows.length); i++) {
          const r = rows[i] || [];
          const text = r.map(c => String(c || '')).join(' ').toLowerCase();
          const score = keywords.filter(kw => text.includes(kw)).length;
          if (score > bestScore) {
            bestScore = score;
            headerIdx = i;
          }
        }

        const rawHeaders = (rows[headerIdx] || []).map(c => String(c || '').trim());
        const combinedHeaders = [...rawHeaders];
        let dataStartRow = headerIdx + 1;

        for (let n = headerIdx + 1; n < Math.min(headerIdx + 4, rows.length); n++) {
          const nextR = (rows[n] || []).map(c => String(c || '').trim());
          const subText = nextR.join(' ').toLowerCase();
          if (['recd. dt', 'recd dt', 'address', 'invo. no', 'thoka no', 'net wt'].some(kw => subText.includes(kw))) {
            dataStartRow = n + 1;
            for (let c = 0; c < Math.max(combinedHeaders.length, nextR.length); c++) {
              const sub = nextR[c] || '';
              if (sub) {
                if (combinedHeaders[c] && !combinedHeaders[c].toLowerCase().includes(sub.toLowerCase())) {
                  combinedHeaders[c] += '|' + sub;
                } else if (!combinedHeaders[c]) {
                  combinedHeaders[c] = sub;
                }
              }
            }
          }
        }

        const colMap = {};
        let bestDelScore = -1;
        let bestDelCol = -1;

        combinedHeaders.forEach((h, i) => {
          const hc = h.toLowerCase().replace(/\r?\n/g, ' ').trim();
          if (['sr.no', 'sr no', 's.no'].some(k => hc.includes(k)) && !('sr_no' in colMap)) colMap.sr_no = i;
          else if (['jobno', 'job no'].some(k => hc.includes(k)) && !('job_no' in colMap)) colMap.job_no = i;
          else if (['client', 'customer', 'party'].some(k => hc.includes(k)) && !('client_name' in colMap)) colMap.client_name = i;
          else if (['b/l', 'bl no', 'blno'].some(k => hc.includes(k)) && !('bl_no' in colMap)) colMap.bl_no = i;
          else if (hc === 'date' && !('bl_date' in colMap)) colMap.bl_date = i;
          else if (['igm item', 'igm'].some(k => hc.includes(k)) && hc.includes('item') && !('igm_item' in colMap)) colMap.igm_item = i;
          else if (hc.includes('marks') && !('marks_no' in colMap)) colMap.marks_no = i;
          else if (['pkg', 'wt'].some(k => hc.includes(k)) && !('pkgs_wt' in colMap)) colMap.pkgs_wt = i;
          else if (['cth', 'desc'].some(k => hc.includes(k)) && !('cargo_desc' in colMap)) colMap.cargo_desc = i;
          else if (['invo', 'cif', 'ass. val'].some(k => hc.includes(k)) && hc.includes('val') && !('invo_val' in colMap)) colMap.invo_val = i;
          else if (hc.includes('duty') && !('duty_paid' in colMap)) colMap.duty_paid = i;

          const delKw = ['delivery', 'igm date', 'be date', 'date'];
          let dScore = delKw.filter(kw => hc.includes(kw)).length;
          if (hc.includes('delivery') && (hc.includes('igm') || hc.includes('be date'))) dScore += 5;
          if (dScore > bestDelScore) {
            bestDelScore = dScore;
            bestDelCol = i;
          }

          if (hc.includes('bill no') && hc.includes('date') && !('bill_col' in colMap)) colMap.bill_col = i;
          else if (hc.includes('bill no') && !('bill_col' in colMap)) colMap.bill_col = i;
        });

        if (bestDelCol >= 0) colMap.delivery_col = bestDelCol;

        const defaults = {
          sr_no: 0, job_no: 1, client_name: 2, bl_no: 3,
          bl_date: 4, igm_item: 5, delivery_col: 6, marks_no: 7,
          pkgs_wt: 8, cargo_desc: 9, invo_val: 10, duty_paid: 11,
          bill_col: 12
        };
        for (const k in defaults) {
          if (!(k in colMap)) colMap[k] = defaults[k];
        }

        const cSr = colMap.sr_no;
        const cJob = colMap.job_no;
        const cClient = colMap.client_name;
        const cBl = colMap.bl_no;
        const cBlDate = colMap.bl_date || 4;
        const cIgm = colMap.igm_item || 5;
        const cDel = colMap.delivery_col;
        const cMarks = colMap.marks_no || 7;
        const cPkg = colMap.pkgs_wt || 8;
        const cCargo = colMap.cargo_desc || 9;
        const cInvo = colMap.invo_val || 10;
        const cDuty = colMap.duty_paid || 11;
        const cBill = colMap.bill_col;

        const rawJobs = [];
        let current = null;
        let groupCounter = 0;

        for (let r = dataStartRow; r < rows.length; r++) {
          const row = rows[r] || [];
          const srVal = row[cSr];
          let isTitle = false;
          let parsedSr = null;

          if (srVal !== null && srVal !== undefined) {
            const s = String(srVal).trim();
            if (s && !['sr', 'total', 'grand', 'page'].some(w => s.toLowerCase().includes(w))) {
              const num = parseFloat(s);
              if (!isNaN(num) && Math.floor(num) === num && num > 0) {
                parsedSr = parseInt(num, 10);
                isTitle = true;
              }
            }
          }

          if (isTitle) {
            groupCounter++;
            if (current) rawJobs.push(current);
            current = {
              group_sr_no: groupCounter,
              sr_no: parsedSr,
              title_row: row,
              sub_rows: []
            };
          } else {
            if (current) current.sub_rows.push(row);
          }
        }
        if (current) rawJobs.push(current);

        const parsedJobs = [];
        let lateCount = 0, onTimeCount = 0, unbilledCount = 0, advanceCount = 0, anomalyCount = 0;
        let totalDuty = 0, totalAssessable = 0;
        let tatSum = 0, tatCount = 0;

        for (const item of rawJobs) {
          const t = item.title_row;
          const subs = item.sub_rows;

          const delDate = parseDate(t[cDel]);
          const blDate = parseDate(t[cBlDate]);
          const dutyPaid = parseNumber(t[cDuty]);
          const invoVal = parseNumber(t[cInvo]);

          let billDate = null;
          let clientAddress = '';
          let recdDate = null;
          let invoNoFull = '';
          let thokaNo = '';
          let igmDate = null;
          let beDate = null;
          let assVal = 0;
          let cifVal = 0;

          for (const s of subs) {
            if (s[cBill] !== undefined && s[cBill] !== null && !billDate) {
              const b = parseDate(s[cBill]);
              if (b) billDate = b;
            }
            if (s[cJob] !== undefined && s[cJob] !== null && !recdDate) {
              const r = parseDate(s[cJob]);
              if (r) recdDate = r;
            }
            if (s[cClient] !== undefined && s[cClient] !== null && !clientAddress) {
              const a = String(s[cClient]).trim();
              if (a) clientAddress = a;
            }
            if (s[cBl] !== undefined && s[cBl] !== null && !invoNoFull) {
              const inv = String(s[cBl]).trim();
              if (inv) invoNoFull = inv;
            }
            if (s[cIgm] !== undefined && s[cIgm] !== null && !thokaNo) {
              const th = String(s[cIgm]).trim();
              if (th) thokaNo = th;
            }
            if (s[cDel] !== undefined && s[cDel] !== null) {
              const d = parseDate(s[cDel]);
              if (d) {
                if (!igmDate) igmDate = d;
                else if (!beDate) beDate = d;
              }
            }
            if (s[cInvo] !== undefined && s[cInvo] !== null) {
              const n = parseNumber(s[cInvo]);
              if (n > 0) {
                if (assVal === 0) assVal = n;
                else if (cifVal === 0) cifVal = n;
              }
            }
          }

          if (assVal === 0 && invoVal > 0) assVal = invoVal;

          totalDuty += dutyPaid;
          totalAssessable += assVal;

          let diffDays = null;
          let billingStatus = '';
          let isLate = false;
          let comments = '';
          let agingBucket = '';

          if (delDate && billDate) {
            const msDiff = billDate.getTime() - delDate.getTime();
            diffDays = Math.round(msDiff / (1000 * 60 * 60 * 24));

            if (diffDays <= 2 && diffDays >= 0) agingBucket = '0-2 Days (On-Time)';
            else if (diffDays >= 3 && diffDays <= 5) agingBucket = '3-5 Days (Low Delay)';
            else if (diffDays >= 6 && diffDays <= 10) agingBucket = '6-10 Days (Moderate Delay)';
            else if (diffDays >= 11 && diffDays <= 20) agingBucket = '11-20 Days (Significant Delay)';
            else if (diffDays >= 21 && diffDays <= 30) agingBucket = '21-30 Days (High Delay)';
            else if (diffDays > 30 && diffDays <= 100) agingBucket = '> 30 Days (Severe Delay)';
            else if (diffDays > 100) agingBucket = 'Date Anomaly (> 100d)';
            else if (diffDays < 0) agingBucket = 'Advance Billed (< 0d)';

            if (diffDays > 100 || diffDays < -60) {
              billingStatus = 'Date Anomaly';
              isLate = true;
              anomalyCount++;
              comments = `⚠️ DATA ERROR: Difference is ${diffDays} days. Possible year typo (Del: ${formatDate(delDate)}, Bill: ${formatDate(billDate)}).`;
            } else if (diffDays > 2) {
              billingStatus = 'Late Billing';
              isLate = true;
              lateCount++;
              comments = `🚨 LATE BILLING: Delayed by ${diffDays} days (Exceeds 2-day SLA limit). Mention Reason for delay.`;
            } else if (diffDays === 0) {
              billingStatus = 'Same-Day Billing';
              isLate = false;
              onTimeCount++;
              comments = '✅ SAME-DAY: Billed on delivery date (Compliant).';
            } else if (diffDays >= 1 && diffDays <= 2) {
              billingStatus = 'On-Time Billing';
              isLate = false;
              onTimeCount++;
              comments = `✅ ON-TIME: Billed in ${diffDays} day(s) (Within 2-day limit).`;
            } else if (diffDays < 0) {
              billingStatus = 'Advance Billed';
              isLate = false;
              advanceCount++;
              comments = `ℹ️ ADVANCE: Bill raised ${Math.abs(diffDays)} day(s) before recorded delivery date.`;
            }

            if (diffDays >= 0 && diffDays <= 100) {
              tatSum += diffDays;
              tatCount++;
            }
          } else if (delDate && !billDate) {
            billingStatus = 'Unbilled Delivery';
            isLate = true;
            unbilledCount++;
            agingBucket = 'Unbilled Revenue Exposure';
            comments = `🔥 CRITICAL: Cargo delivered on ${formatDate(delDate)} but NO Bill Date found. Working capital & revenue leakage risk.`;
          } else if (!delDate && billDate) {
            billingStatus = 'Undelivered / Billed';
            isLate = false;
            agingBucket = 'Undelivered / In-Transit';
            comments = `ℹ️ IN-TRANSIT: Bill date ${formatDate(billDate)} present, but cargo delivery date pending.`;
          } else {
            billingStatus = 'Incomplete Dates';
            isLate = false;
            comments = 'Missing both delivery and bill dates.';
          }

          parsedJobs.push({
            group_sr_no: item.group_sr_no,
            sr_no: item.sr_no,
            job_no: String(t[cJob] || '').trim(),
            client_name: String(t[cClient] || '').trim(),
            client_address: clientAddress,
            bl_no: String(t[cBl] || '').trim(),
            bl_date: formatDate(blDate),
            recd_date: formatDate(recdDate),
            invo_no_full: invoNoFull,
            delivery_date: formatDate(delDate),
            bill_no: String(t[cBill] || '').trim(),
            bill_date: formatDate(billDate),
            diff_days: diffDays,
            is_late: isLate,
            billing_status: billingStatus,
            aging_bucket: agingBucket,
            comments: comments,
            duty_paid: dutyPaid,
            assessable_value: assVal
          });
        }

        const avgTat = tatCount > 0 ? (tatSum / tatCount).toFixed(1) : '0.0';

        return {
          totalJobs: parsedJobs.length,
          lateCount,
          onTimeCount,
          unbilledCount,
          advanceCount,
          anomalyCount,
          totalDuty,
          totalAssessable,
          avgTat,
          jobs: parsedJobs
        };
      }

      // UI Renderers
      function renderUI(res) {
        kpiSection.style.display = 'grid';
        chartsSection.style.display = 'grid';
        tableSection.style.display = 'flex';

        // Update KPIs
        document.getElementById('kpi-total').textContent = res.totalJobs.toLocaleString();
        document.getElementById('kpi-late').textContent = res.lateCount.toLocaleString();
        document.getElementById('kpi-late-pct').textContent = ((res.lateCount / res.totalJobs) * 100).toFixed(1) + '%';
        document.getElementById('kpi-ontime').textContent = res.onTimeCount.toLocaleString();
        document.getElementById('kpi-ontime-pct').textContent = ((res.onTimeCount / res.totalJobs) * 100).toFixed(1) + '%';
        document.getElementById('kpi-unbilled').textContent = res.unbilledCount.toLocaleString();
        document.getElementById('kpi-duty').textContent = '₹' + (res.totalDuty / 10000000).toFixed(2) + ' Cr';
        document.getElementById('kpi-assessable').textContent = '₹' + (res.totalAssessable / 10000000).toFixed(2) + ' Cr';
        document.getElementById('kpi-tat').textContent = res.avgTat + ' d';

        // Filter Counts
        document.getElementById('count-all').textContent = res.totalJobs;
        document.getElementById('count-late').textContent = res.lateCount;
        document.getElementById('count-ontime').textContent = res.onTimeCount;
        document.getElementById('count-unbilled').textContent = res.unbilledCount;
        document.getElementById('count-advance').textContent = res.advanceCount;
        document.getElementById('count-anomaly').textContent = res.anomalyCount;

        // Render Charts
        renderCharts(res);

        // Render Table
        applyFilterAndSearch();
      }

      function renderCharts(res) {
        if (chartStatus) chartStatus.destroy();
        if (chartAging) chartAging.destroy();
        if (chartClients) chartClients.destroy();

        // Status Chart
        const ctxStatus = document.getElementById('chart-status').getContext('2d');
        chartStatus = new Chart(ctxStatus, {
          type: 'doughnut',
          data: {
            labels: ['Late (> 2d)', 'On-Time (≤ 2d)', 'Unbilled', 'Advance', 'Anomalies'],
            datasets: [{
              data: [res.lateCount, res.onTimeCount, res.unbilledCount, res.advanceCount, res.anomalyCount],
              backgroundColor: ['#ef4444', '#10b981', '#f59e0b', '#0284c7', '#8b5cf6'],
              borderWidth: 2,
              borderColor: '#ffffff'
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
              legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } }
            }
          }
        });

        // Aging Buckets Chart
        const bucketCounts = {
          '0-2d (On-Time)': 0,
          '3-5d (Low)': 0,
          '6-10d (Mod)': 0,
          '11-20d (Sig)': 0,
          '21-30d (High)': 0,
          '> 30d (Severe)': 0,
          'Unbilled': res.unbilledCount
        };

        res.jobs.forEach(j => {
          if (j.diff_days !== null) {
            const d = j.diff_days;
            if (d >= 0 && d <= 2) bucketCounts['0-2d (On-Time)']++;
            else if (d >= 3 && d <= 5) bucketCounts['3-5d (Low)']++;
            else if (d >= 6 && d <= 10) bucketCounts['6-10d (Mod)']++;
            else if (d >= 11 && d <= 20) bucketCounts['11-20d (Sig)']++;
            else if (d >= 21 && d <= 30) bucketCounts['21-30d (High)']++;
            else if (d > 30 && d <= 100) bucketCounts['> 30d (Severe)']++;
          }
        });

        const ctxAging = document.getElementById('chart-aging').getContext('2d');
        chartAging = new Chart(ctxAging, {
          type: 'bar',
          data: {
            labels: Object.keys(bucketCounts),
            datasets: [{
              label: 'Job Count',
              data: Object.values(bucketCounts),
              backgroundColor: ['#10b981', '#fbbf24', '#f97316', '#ef4444', '#b91c1c', '#7f1d1d', '#f59e0b'],
              borderRadius: 4
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
              x: { ticks: { font: { size: 10 } } },
              y: { beginAtZero: true }
            }
          }
        });

        // Top Clients Chart
        const clientLateMap = {};
        res.jobs.forEach(j => {
          if (j.billing_status === 'Late Billing') {
            const name = j.client_name || 'UNKNOWN';
            clientLateMap[name] = (clientLateMap[name] || 0) + 1;
          }
        });

        const sortedClients = Object.entries(clientLateMap)
          .sort((a, b) => b[1] - a[1])
          .slice(0, 8);

        const ctxClients = document.getElementById('chart-clients').getContext('2d');
        chartClients = new Chart(ctxClients, {
          type: 'bar',
          indexAxis: 'y',
          data: {
            labels: sortedClients.map(c => c[0].length > 18 ? c[0].slice(0, 16) + '...' : c[0]),
            datasets: [{
              label: 'Late Jobs',
              data: sortedClients.map(c => c[1]),
              backgroundColor: '#ef4444',
              borderRadius: 4
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
              x: { beginAtZero: true },
              y: { ticks: { font: { size: 10 } } }
            }
          }
        });
      }

      // Filtering & Table Logic
      function applyFilterAndSearch() {
        filteredJobs = allJobs.filter(j => {
          // Filter pill
          if (currentFilter === 'late' && j.billing_status !== 'Late Billing') return false;
          if (currentFilter === 'ontime' && !j.billing_status.includes('On-Time') && !j.billing_status.includes('Same-Day')) return false;
          if (currentFilter === 'unbilled' && j.billing_status !== 'Unbilled Delivery') return false;
          if (currentFilter === 'advance' && j.billing_status !== 'Advance Billed') return false;
          if (currentFilter === 'anomaly' && j.billing_status !== 'Date Anomaly') return false;

          // Search text
          if (searchQuery) {
            const q = searchQuery.toLowerCase();
            const text = `${j.client_name} ${j.job_no} ${j.bill_no} ${j.bl_no} ${j.client_address} ${j.comments}`.toLowerCase();
            if (!text.includes(q)) return false;
          }

          return true;
        });

        // Sort
        filteredJobs.sort((a, b) => {
          let va = a[sortCol];
          let vb = b[sortCol];
          if (va === null || va === undefined) va = sortAsc ? Infinity : -Infinity;
          if (vb === null || vb === undefined) vb = sortAsc ? Infinity : -Infinity;
          if (typeof va === 'string') return sortAsc ? va.localeCompare(vb) : vb.localeCompare(va);
          return sortAsc ? va - vb : vb - va;
        });

        currentPage = 1;
        renderTablePage();
      }

      function renderTablePage() {
        const total = filteredJobs.length;
        const totalPages = Math.ceil(total / pageSize) || 1;
        if (currentPage > totalPages) currentPage = totalPages;
        if (currentPage < 1) currentPage = 1;

        const startIdx = (currentPage - 1) * pageSize;
        const pageItems = filteredJobs.slice(startIdx, startIdx + pageSize);

        tableBody.innerHTML = '';

        pageItems.forEach(j => {
          const tr = document.createElement('tr');
          if (j.billing_status === 'Late Billing') tr.className = 'row-late';
          else if (j.billing_status === 'Unbilled Delivery') tr.className = 'row-unbilled';

          let statusBadgeClass = 'badge-success';
          if (j.billing_status === 'Late Billing') statusBadgeClass = 'badge-danger';
          else if (j.billing_status === 'Unbilled Delivery') statusBadgeClass = 'badge-warning';
          else if (j.billing_status === 'Date Anomaly') statusBadgeClass = 'badge-purple';
          else if (j.billing_status === 'Advance Billed') statusBadgeClass = 'badge-info';

          let diffDisplay = '-';
          if (j.diff_days !== null) {
            const sign = j.diff_days > 0 ? '+' : '';
            diffDisplay = `<span class="tat-badge ${j.diff_days > 2 ? 'badge-danger' : 'badge-success'}">${sign}${j.diff_days}d</span>`;
          }

          tr.innerHTML = `
            <td style="font-weight:700; color:var(--text-muted);">${j.group_sr_no}</td>
            <td style="font-weight:700;">#${j.job_no}</td>
            <td class="client-cell">
              <div class="client-title">${escapeHtml(j.client_name)}</div>
              <div class="client-sub">${escapeHtml(j.client_address || 'Address pending')}</div>
            </td>
            <td style="font-weight:600; color:var(--accent);">${j.delivery_date || '-'}</td>
            <td>
              <div style="font-weight:600;">${escapeHtml(j.bill_no || '-')}</div>
              <div style="font-size:11px; color:var(--text-muted);">${j.bill_date || '-'}</div>
            </td>
            <td style="text-align:center;">${diffDisplay}</td>
            <td><span class="kpi-badge ${statusBadgeClass}">${j.billing_status}</span></td>
            <td style="font-size:11.5px; max-width:280px;">${escapeHtml(j.comments)}</td>
            <td style="text-align:right; font-weight:700;">₹${Number(j.duty_paid || 0).toLocaleString('en-IN')}</td>
          `;

          tr.addEventListener('click', () => openDetailModal(j));
          tableBody.appendChild(tr);
        });

        pageInfo.textContent = `Showing ${total === 0 ? 0 : startIdx + 1} - ${Math.min(startIdx + pageSize, total)} of ${total.toLocaleString()} jobs`;
        pageCurrent.textContent = `${currentPage} / ${totalPages}`;
        btnPrev.disabled = currentPage <= 1;
        btnNext.disabled = currentPage >= totalPages;
      }

      function openDetailModal(j) {
        document.getElementById('modal-client').textContent = j.client_name || 'Job Details';
        document.getElementById('modal-sub').textContent = `Group Sr. No #${j.group_sr_no} | Job #${j.job_no} | Status: ${j.billing_status}`;

        const fBox = document.getElementById('modal-finding-box');
        const fTitle = document.getElementById('modal-finding-title');
        const fDesc = document.getElementById('modal-finding-desc');

        if (j.billing_status === 'Late Billing') {
          fBox.className = 'audit-finding-box';
          fTitle.textContent = '🚨 Late Billing Audit Observation (TAT SLA > 2 Days)';
          fDesc.textContent = j.comments;
        } else if (j.billing_status === 'Unbilled Delivery') {
          fBox.className = 'audit-finding-box';
          fTitle.textContent = '🔥 Critical Revenue Leakage Observation';
          fDesc.textContent = j.comments;
        } else {
          fBox.className = 'audit-finding-box compliant';
          fTitle.textContent = '✅ Billing SLA Compliant';
          fDesc.textContent = j.comments;
        }

        document.getElementById('m-group-sr').textContent = `${j.group_sr_no} (Raw Sr.No: ${j.sr_no || '-'})`;
        document.getElementById('m-job-recd').textContent = `${j.job_no || '-'} (Recd: ${j.recd_date || '-'})`;
        document.getElementById('m-delivery-date').textContent = j.delivery_date || 'Missing';
        document.getElementById('m-bill-info').textContent = `${j.bill_no || '-'} on ${j.bill_date || 'Not Billed'}`;
        document.getElementById('m-diff').textContent = j.diff_days !== null ? `${j.diff_days} Days (${j.diff_days > 2 ? 'Breach' : 'Compliant'})` : 'N/A';
        document.getElementById('m-aging').textContent = j.aging_bucket || '-';
        document.getElementById('m-bl-info').textContent = `${j.bl_no || '-'} (${j.bl_date || '-'})`;
        document.getElementById('m-invo-info').textContent = j.invo_no_full || '-';
        document.getElementById('m-duty').textContent = '₹' + Number(j.duty_paid || 0).toLocaleString('en-IN');
        document.getElementById('m-assessable').textContent = '₹' + Number(j.assessable_value || 0).toLocaleString('en-IN');
        document.getElementById('m-address').textContent = j.client_address || 'No address line detected in sub-rows';

        detailModal.classList.add('active');
      }

      function escapeHtml(str) {
        return String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
      }

      // UI Listeners
      document.querySelectorAll('.pill').forEach(btn => {
        btn.addEventListener('click', () => {
          document.querySelectorAll('.pill').forEach(p => p.classList.remove('active'));
          btn.classList.add('active');
          currentFilter = btn.dataset.filter;
          applyFilterAndSearch();
        });
      });

      searchInput.addEventListener('input', (e) => {
        searchQuery = e.target.value.trim();
        applyFilterAndSearch();
      });

      document.querySelectorAll('th[data-sort]').forEach(th => {
        th.addEventListener('click', () => {
          const col = th.dataset.sort;
          if (sortCol === col) {
            sortAsc = !sortAsc;
          } else {
            sortCol = col;
            sortAsc = true;
          }
          applyFilterAndSearch();
        });
      });

      btnPrev.addEventListener('click', () => {
        if (currentPage > 1) {
          currentPage--;
          renderTablePage();
        }
      });

      btnNext.addEventListener('click', () => {
        const totalPages = Math.ceil(filteredJobs.length / pageSize);
        if (currentPage < totalPages) {
          currentPage++;
          renderTablePage();
        }
      });

      btnClear.addEventListener('click', () => {
        allJobs = [];
        filteredJobs = [];
        rawSheetRows = [];
        fileInput.value = '';
        kpiSection.style.display = 'none';
        chartsSection.style.display = 'none';
        tableSection.style.display = 'none';
        statusBar.classList.remove('active');
        btnExportExcel.disabled = true;
        btnExportCsv.disabled = true;
        btnClear.style.display = 'none';
      });

      // Export Multi-Sheet Excel Workbook
      btnExportExcel.addEventListener('click', () => {
        if (allJobs.length === 0) return;
        showStatus('Generating complete 7-sheet Audit Workbook (.xlsx)...');

        setTimeout(() => {
          try {
            const wb = XLSX.utils.book_new();

            // Sheet 1: Late_Billing_Register
            const s1Headers = [
              'Group Sr. No', 'Sr. No', 'Job No', 'Client Name', 'Client Address',
              'B/L No', 'B/L Date', 'Recd Date', 'Invoice No & Date',
              'Delivery Date', 'Bill No', 'Bill Date', 'Difference (Days)',
              'Billing Status', 'Aging Category', 'Auditor Comments',
              'Customs Duty Paid (INR)', 'Assessable Value (INR)'
            ];
            const s1Rows = [s1Headers];
            for (const j of allJobs) {
              s1Rows.push([
                j.group_sr_no, j.sr_no, j.job_no, j.client_name, j.client_address,
                j.bl_no, j.bl_date, j.recd_date, j.invo_no_full,
                j.delivery_date, j.bill_no, j.bill_date, j.diff_days !== null ? j.diff_days : '',
                j.billing_status, j.aging_bucket, j.comments,
                j.duty_paid, j.assessable_value
              ]);
            }
            const ws1 = XLSX.utils.aoa_to_sheet(s1Rows);
            XLSX.utils.book_append_sheet(wb, ws1, 'Late_Billing_Register');

            // Sheet 2: Executive_Audit_Summary
            const total = allJobs.length;
            const late = allJobs.filter(j => j.billing_status === 'Late Billing').length;
            const ontime = allJobs.filter(j => j.billing_status.includes('On-Time') || j.billing_status.includes('Same-Day')).length;
            const unbilled = allJobs.filter(j => j.billing_status === 'Unbilled Delivery').length;
            const advance = allJobs.filter(j => j.billing_status === 'Advance Billed').length;
            const anomaly = allJobs.filter(j => j.billing_status === 'Date Anomaly').length;
            const duty = allJobs.reduce((acc, j) => acc + (j.duty_paid || 0), 0);

            const s2Rows = [
              ['SP SHAH & CO. - CUSTOMS BROKER & LOGISTICS AUDIT REPORT'],
              ['EXECUTIVE SUMMARY & BILLING TAT COMPLIANCE EVALUATION'],
              ['Generated On:', new Date().toLocaleString()],
              [''],
              ['Audit Key Performance Indicator (KPI)', 'Count / Metric', '% of Total Jobs', 'Duty / Financial Exposure (INR)'],
              ['Total Port Import Jobs Audited', total, '100.00%', duty],
              ['On-Time Billing (<= 2 Days SLA)', ontime, ((ontime / total) * 100).toFixed(2) + '%', '-'],
              ['Late Billing (> 2 Days Delay)', late, ((late / total) * 100).toFixed(2) + '%', 'Working Capital & Delay Risk'],
              ['Critical Unbilled Deliveries (Missing Bill Date)', unbilled, ((unbilled / total) * 100).toFixed(2) + '%', 'Immediate Leakage Risk'],
              ['Advance Billed (Bill Date < Delivery Date)', advance, ((advance / total) * 100).toFixed(2) + '%', 'Early Invoicing'],
              ['Data Discrepancies & Date Anomalies', anomaly, ((anomaly / total) * 100).toFixed(2) + '%', 'Input Typo Review Required'],
              [''],
              ['Audit SLA Benchmark:', '<= 2 Business Days between Cargo Delivery & Bill Date'],
              ['Compliance Status:', late > 0 ? 'NON-COMPLIANT (TAT SLA Breached on ' + ((late / total) * 100).toFixed(1) + '% jobs)' : 'COMPLIANT']
            ];
            const ws2 = XLSX.utils.aoa_to_sheet(s2Rows);
            XLSX.utils.book_append_sheet(wb, ws2, 'Executive_Audit_Summary');

            // Sheet 3: Client_Revenue_TAT_Analysis
            const clientMap = {};
            for (const j of allJobs) {
              const cname = j.client_name || 'UNKNOWN CLIENT';
              if (!clientMap[cname]) {
                clientMap[cname] = { total: 0, late: 0, onTime: 0, unbilled: 0, duty: 0, assessable: 0, tatSum: 0, tatCount: 0 };
              }
              const c = clientMap[cname];
              c.total++;
              c.duty += j.duty_paid || 0;
              c.assessable += j.assessable_value || 0;
              if (j.billing_status === 'Late Billing') c.late++;
              if (j.billing_status.includes('On-Time') || j.billing_status.includes('Same-Day')) c.onTime++;
              if (j.billing_status === 'Unbilled Delivery') c.unbilled++;
              if (j.diff_days !== null && j.diff_days >= 0 && j.diff_days <= 100) {
                c.tatSum += j.diff_days;
                c.tatCount++;
              }
            }

            const s3Headers = [
              'Client Name', 'Total Jobs', 'Late Billing Jobs', 'Late Billing %',
              'On-Time Jobs', 'Unbilled Deliveries', 'Average Billing TAT (Days)',
              'Total Customs Duty (INR)', 'Total Assessable Value (INR)', 'Audit Risk Level'
            ];
            const s3Rows = [s3Headers];
            const sortedClients = Object.keys(clientMap).sort((a, b) => clientMap[b].late - clientMap[a].late);
            for (const name of sortedClients) {
              const c = clientMap[name];
              const latePct = c.total > 0 ? ((c.late / c.total) * 100).toFixed(1) + '%' : '0.0%';
              const avgTat = c.tatCount > 0 ? (c.tatSum / c.tatCount).toFixed(1) : 'N/A';
              let risk = 'Low';
              if (c.unbilled > 0 || c.late > 10) risk = 'High';
              else if (c.late > 2) risk = 'Medium';

              s3Rows.push([
                name, c.total, c.late, latePct, c.onTime, c.unbilled, avgTat, c.duty, c.assessable, risk
              ]);
            }
            const ws3 = XLSX.utils.aoa_to_sheet(s3Rows);
            XLSX.utils.book_append_sheet(wb, ws3, 'Client_Revenue_TAT_Analysis');

            // Sheet 4: Billing_Aging_Drilldown
            const bucketMap = {};
            const bucketOrder = [
              '0-2 Days (On-Time)', '3-5 Days (Low Delay)', '6-10 Days (Moderate Delay)',
              '11-20 Days (Significant Delay)', '21-30 Days (High Delay)', '> 30 Days (Severe Delay)',
              'Unbilled Revenue Exposure', 'Advance Billed (< 0d)', 'Date Anomaly (> 100d)', 'Undelivered / In-Transit'
            ];
            for (const b of bucketOrder) {
              bucketMap[b] = { count: 0, duty: 0, assessable: 0 };
            }
            for (const j of allJobs) {
              const b = j.aging_bucket || 'Other';
              if (!bucketMap[b]) bucketMap[b] = { count: 0, duty: 0, assessable: 0 };
              bucketMap[b].count++;
              bucketMap[b].duty += j.duty_paid || 0;
              bucketMap[b].assessable += j.assessable_value || 0;
            }
            const s4Headers = ['Aging Category', 'Job Count', '% of Total Jobs', 'Customs Duty Exposure (INR)', 'Assessable Value (INR)', 'Auditor Evaluation'];
            const s4Rows = [s4Headers];
            for (const b of bucketOrder) {
              const d = bucketMap[b] || { count: 0, duty: 0, assessable: 0 };
              const pct = total > 0 ? ((d.count / total) * 100).toFixed(2) + '%' : '0.0%';
              let evalText = 'TAT Compliant';
              if (b.includes('Severe') || b.includes('Unbilled')) evalText = 'Critical Attention Required';
              else if (b.includes('Significant') || b.includes('High')) evalText = 'Elevated Revenue Risk';
              else if (b.includes('Low') || b.includes('Moderate')) evalText = 'Minor Delay SLA';
              else if (b.includes('Anomaly')) evalText = 'Input Data Error';
              s4Rows.push([b, d.count, pct, d.duty, d.assessable, evalText]);
            }
            const ws4 = XLSX.utils.aoa_to_sheet(s4Rows);
            XLSX.utils.book_append_sheet(wb, ws4, 'Billing_Aging_Drilldown');

            // Sheet 5: Unbilled_Revenue_Exposure
            const s5Headers = ['Group Sr. No', 'Job No', 'Client Name', 'Client Address', 'B/L No', 'Delivery Date', 'Duty Paid (INR)', 'Assessable Value (INR)', 'Auditor Comment'];
            const s5Rows = [s5Headers];
            for (const j of allJobs.filter(x => x.billing_status === 'Unbilled Delivery')) {
              s5Rows.push([j.group_sr_no, j.job_no, j.client_name, j.client_address, j.bl_no, j.delivery_date, j.duty_paid, j.assessable_value, j.comments]);
            }
            const ws5 = XLSX.utils.aoa_to_sheet(s5Rows);
            XLSX.utils.book_append_sheet(wb, ws5, 'Unbilled_Revenue_Exposure');

            // Sheet 6: Data_Integrity_Anomalies
            const s6Headers = ['Group Sr. No', 'Job No', 'Client Name', 'Delivery Date', 'Bill Date', 'Difference (Days)', 'Auditor Finding'];
            const s6Rows = [s6Headers];
            for (const j of allJobs.filter(x => x.billing_status === 'Date Anomaly')) {
              s6Rows.push([j.group_sr_no, j.job_no, j.client_name, j.delivery_date, j.bill_date, j.diff_days, j.comments]);
            }
            const ws6 = XLSX.utils.aoa_to_sheet(s6Rows);
            XLSX.utils.book_append_sheet(wb, ws6, 'Data_Integrity_Anomalies');

            // Sheet 7: Enriched_Raw_Register
            if (rawSheetRows && rawSheetRows.length > 0) {
              const ws7 = XLSX.utils.aoa_to_sheet(rawSheetRows);
              XLSX.utils.book_append_sheet(wb, ws7, 'Enriched_Raw_Register');
            }

            const now = new Date();
            const dateStr = now.toISOString().slice(0, 10).replace(/-/g, '');
            const filename = `SP_Shah_Audit_Report_${dateStr}.xlsx`;
            XLSX.writeFile(wb, filename);

            showStatus('Audit Workbook exported successfully: ' + filename);
          } catch (err) {
            console.error(err);
            alert('Export error: ' + err.message);
            showStatus('Export failed: ' + err.message);
          }
        }, 50);
      });

      // Export CSV
      btnExportCsv.addEventListener('click', () => {
        const lateJobs = allJobs.filter(j => j.billing_status === 'Late Billing');
        if (lateJobs.length === 0) {
          alert('No late billing jobs to export.');
          return;
        }
        const csvRows = [
          ['Group Sr. No', 'Job No', 'Client Name', 'Client Address', 'Delivery Date', 'Bill No', 'Bill Date', 'Difference (Days)', 'Comments', 'Customs Duty (INR)']
        ];
        lateJobs.forEach(j => {
          csvRows.push([
            j.group_sr_no, j.job_no, `"${j.client_name.replace(/"/g, '""')}"`, `"${(j.client_address || '').replace(/"/g, '""')}"`,
            j.delivery_date, j.bill_no, j.bill_date, j.diff_days, `"${j.comments.replace(/"/g, '""')}"`, j.duty_paid
          ]);
        });
        const csvContent = "data:text/csv;charset=utf-8," + csvRows.map(e => e.join(",")).join("\n");
        const encodedUri = encodeURI(csvContent);
        const link = document.createElement("a");
        link.setAttribute("href", encodedUri);
        link.setAttribute("download", `SP_Shah_Late_Billing_${new Date().toISOString().slice(0, 10)}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
      });

    })();