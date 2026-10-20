import os, glob, re, json, datetime, openpyxl

folder = os.path.dirname(os.path.abspath(__file__))

def get_excel_file():
    files = glob.glob(os.path.join(folder, "*.xlsx"))
    return files[0] if files else None

def fdate(v):
    if not v: return ""
    if hasattr(v, "strftime"): return v.strftime("%d/%m/%Y")
    if isinstance(v, (int, float)):
        try: return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))).strftime("%d/%m/%Y")
        except: return str(v)
    return str(v).strip()

def sgroup(s):
    s = str(s or "").strip()
    m = re.match(r"^\.?(\d+)", s)
    g = {"00":"00 Customer Docs","01":"01 Estimate Pending","02":"02 Intimation Pending","03":"03 Surveyor Assign","04":"04 Survey Pending","05":"05 Survey Approval","06":"06 Investigation","07":"07 Customer Approval","08":"08 Store Clarification","09":"09 Parts Pending","11":"11 Dismantling","12":"12 Mech Work","13":"13 Denting","14":"14 Painting","15":"15 Fitting","16":"16 Final Inspection","17":"17 Ready","18":"18 Total Loss Process","19":"19 Total Loss","20":"20 No Claim"}
    if m:
        p = m.group(1).zfill(2)
        if p in g: return g[p]
    if "TOTAL LOSS" in s.upper(): return "19 Total Loss"
    if "NO CLAIM" in s.upper(): return "20 No Claim"
    return s[:25]

def read_data(fpath):
    wb = openpyxl.load_workbook(fpath, data_only=True)
    ws = wb["DVPR"] if "DVPR" in wb.sheetnames else wb.active
    rows = list(ws.iter_rows(min_row=3, values_only=True))
    data = []
    for i, r in enumerate(rows, 1):
        if not r or (not r[1] and not r[3]): continue
        st = str(r[15] or "").strip()
        if not st or st == "None": continue
        try: age = int(float(str(r[19]))) if r[19] not in ("", None, "#N/A") else 0
        except: age = 0
        try: sage = int(float(str(r[14]))) if r[14] not in ("", None, "#N/A") else 0
        except: sage = 0
        data.append({
            "id": i,
            "customer": str(r[1] or "").strip(),
            "contact": str(r[2] or "").strip(),
            "reg": str(r[3] or "").strip(),
            "model": str(r[4] or "").strip(),
            "advisor": str(r[5] or "").strip(),
            "ro": str(r[6] or "#N/A").strip(),
            "rodate": fdate(r[7]),
            "rostatus": str(r[9] or "#N/A").strip(),
            "promise": fdate(r[10]),
            "insurance": str(r[12] or "Not Specified").strip(),
            "repair": str(r[13] or "MEDIUM").strip(),
            "statusage": sage,
            "status": st,
            "group": sgroup(st),
            "statusdate": fdate(r[16]),
            "panel": str(r[17] or "").strip(),
            "inward": fdate(r[18]),
            "age": age,
            "bucket": str(r[20] or "").strip()
        })
    return data

def build_html(data, fname):
    jdata = json.dumps(data, ensure_ascii=False)
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=5,viewport-fit=cover">
<title>ASM Hyundai | Operations Dashboard</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
:root{{--bg:#fff5f7;--bg-card:#ffffff;--rose-50:#fff1f2;--rose-100:#ffe4e6;--rose-200:#fecdd3;--rose-300:#fda4af;--rose-400:#fb7185;--rose-500:#f43f5e;--rose-600:#e11d48;--rose-700:#be123c;--rose-800:#9f1239;--rose-900:#881337;--text-main:#0f172a;--text-sub:#334155;--text-muted:#64748b;--border:#ffe4e6;--border-dark:#fecdd3;--radius:16px;--shadow:0 4px 20px -2px rgba(244,63,94,0.07),0 2px 6px -1px rgba(0,0,0,0.03)}}
*{{margin:0;padding:0;box-sizing:border-box;-webkit-tap-highlight-color:transparent}}
html,body{{max-width:100%;overflow-x:hidden;-webkit-text-size-adjust:100%}}
body{{font-family:'Plus Jakarta Sans',sans-serif;background:linear-gradient(180deg,#fff0f3 0%,#fff5f7 400px,#ffffff 100%);color:var(--text-main);min-height:100vh;padding:20px 24px 60px;padding-left:max(24px,env(safe-area-inset-left));padding-right:max(24px,env(safe-area-inset-right))}}
.container{{max-width:1480px;margin:0 auto;width:100%;min-width:0}}
.top-header{{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:16px;background:var(--bg-card);padding:18px 24px;border-radius:var(--radius);border:1px solid var(--border);box-shadow:var(--shadow);margin-bottom:20px}}
.brand{{display:flex;align-items:center;gap:14px}}
.brand-icon{{width:48px;height:48px;border-radius:14px;background:linear-gradient(135deg,var(--rose-500),var(--rose-700));display:flex;align-items:center;justify-content:center;color:#fff;font-weight:800;font-size:22px;box-shadow:0 8px 18px rgba(225,29,72,0.32);flex-shrink:0}}
.brand-title{{font-size:22px;font-weight:800;color:var(--rose-900);line-height:1.2;letter-spacing:-0.02em}}
.brand-subtitle{{font-size:12px;font-weight:700;color:var(--rose-500);text-transform:uppercase;letter-spacing:0.08em}}
.header-actions{{display:flex;align-items:center;gap:10px;flex-wrap:wrap}}
.live-badge{{display:inline-flex;align-items:center;gap:6px;background:var(--rose-50);border:1px solid var(--rose-200);color:var(--rose-700);padding:6px 14px;border-radius:999px;font-size:12px;font-weight:700;white-space:nowrap}}
.live-dot{{width:8px;height:8px;border-radius:50%;background:#10b981;box-shadow:0 0 0 3px rgba(16,185,129,0.2);animation:pulse 1.8s infinite}}
@keyframes pulse{{0%,100%{{transform:scale(1);opacity:1}}50%{{transform:scale(1.2);opacity:0.7}}}}
.btn{{background:linear-gradient(135deg,var(--rose-500),var(--rose-600));color:#fff;border:none;padding:10px 18px;border-radius:10px;font-size:13px;font-weight:700;font-family:inherit;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;gap:6px;box-shadow:0 4px 12px rgba(225,29,72,0.25);transition:all 0.15s ease;white-space:nowrap;min-height:40px}}
.btn:hover{{background:linear-gradient(135deg,var(--rose-600),var(--rose-700));transform:translateY(-1px)}}
.btn-sec{{background:var(--rose-50);color:var(--rose-700);border:1px solid var(--rose-200);box-shadow:none}}
.btn-sec:hover{{background:var(--rose-100);color:var(--rose-900)}}
.stats-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,160px),1fr));gap:16px;margin-bottom:20px}}
.stat-card{{background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius);padding:18px 20px;box-shadow:var(--shadow);position:relative;overflow:hidden;transition:transform 0.15s,box-shadow 0.15s;min-width:0}}
.stat-card:hover{{transform:translateY(-2px);box-shadow:0 8px 25px -4px rgba(244,63,94,0.14)}}
.stat-card::before{{content:'';position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,var(--rose-400),var(--rose-600))}}
.stat-label{{font-size:12px;font-weight:700;color:var(--rose-800);text-transform:uppercase;letter-spacing:0.04em;margin-bottom:6px}}
.stat-val{{font-size:34px;font-weight:800;color:var(--text-main);line-height:1;margin-bottom:6px}}
.stat-hint{{font-size:12px;color:var(--text-muted);font-weight:600}}
.charts-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,260px),1fr));gap:18px;margin-bottom:20px}}
.chart-card{{background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius);padding:20px;box-shadow:var(--shadow);display:flex;flex-direction:column;min-width:0;overflow:hidden}}
.chart-head{{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px}}
.chart-title{{font-size:16px;font-weight:800;color:var(--rose-900);letter-spacing:-0.01em}}
.chart-badge{{font-size:11px;font-weight:700;background:var(--rose-50);color:var(--rose-700);padding:3px 10px;border-radius:999px;border:1px solid var(--rose-200)}}
.chart-container{{position:relative;flex:1;min-height:260px;width:100%;max-width:100%}}
.chart-container canvas{{max-width:100%!important}}
.card-sec{{background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius);padding:20px;box-shadow:var(--shadow);margin-bottom:20px}}
.sec-head{{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:10px}}
.sec-title{{font-size:17px;font-weight:800;color:var(--rose-900)}}
.controls-bar{{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:16px}}
.search-box{{flex:1;min-width:min(100%,260px);position:relative}}
.search-input{{width:100%;padding:11px 14px 11px 38px;border-radius:10px;border:1px solid var(--border-dark);font-family:inherit;font-size:16px;font-weight:500;color:var(--text-main);background:#fff;outline:none;transition:border-color 0.15s;min-height:42px}}
.search-input:focus{{border-color:var(--rose-500);box-shadow:0 0 0 3px rgba(244,63,94,0.12)}}
.search-icon{{position:absolute;left:12px;top:50%;transform:translateY(-50%);color:var(--rose-400);width:16px;height:16px}}
.filters-wrap{{display:flex;align-items:center;gap:8px;flex-wrap:wrap}}
.select-filter{{padding:10px 12px;border-radius:10px;border:1px solid var(--border-dark);font-family:inherit;font-size:12px;font-weight:700;color:var(--text-sub);background:#fff;outline:none;min-height:42px}}
.pills-bar{{display:flex;align-items:center;gap:8px;overflow-x:auto;-webkit-overflow-scrolling:touch;padding-bottom:8px;margin-bottom:16px}}
.pill{{padding:7px 15px;border-radius:999px;font-size:12px;font-weight:700;background:var(--rose-50);color:var(--rose-800);border:1px solid var(--rose-200);cursor:pointer;transition:all 0.15s;white-space:nowrap;flex-shrink:0}}
.pill:hover,.pill.active{{background:var(--rose-600);color:#fff;border-color:var(--rose-600);box-shadow:0 3px 10px rgba(225,29,72,0.22)}}
.table-wrap{{overflow-x:auto;-webkit-overflow-scrolling:touch;border:1px solid var(--border);border-radius:12px}}
table{{width:100%;border-collapse:collapse;font-size:13px;text-align:left}}
th{{background:var(--rose-50);color:var(--rose-900);padding:13px 14px;font-weight:800;font-size:12px;text-transform:uppercase;letter-spacing:0.04em;border-bottom:1px solid var(--border-dark);white-space:nowrap}}
td{{padding:12px 14px;border-bottom:1px solid var(--border);color:var(--text-main);font-weight:500;white-space:nowrap}}
tr:nth-child(even){{background:#fffbfc}}
tr:hover{{background:var(--rose-50)}}
.badge{{display:inline-flex;align-items:center;padding:4px 10px;border-radius:999px;font-size:11px;font-weight:700;line-height:1.3;white-space:nowrap}}
.badge-ready{{background:#dcfce7;color:#166534;border:1px solid #bbf7d0}}
.badge-parts{{background:var(--rose-100);color:var(--rose-800);border:1px solid var(--rose-200)}}
.badge-work{{background:#e0e7ff;color:#3730a3;border:1px solid #c7d2fe}}
.badge-survey{{background:#fef3c7;color:#92400e;border:1px solid #fde68a}}
.badge-other{{background:#f1f5f9;color:#475569;border:1px solid #cbd5e1}}
.badge-age-danger{{background:#fee2e2;color:#991b1b;border:1px solid #fecaca;font-weight:800}}
.badge-age-warning{{background:#fef3c7;color:#92400e;border:1px solid #fde68a;font-weight:700}}
.badge-age-ok{{background:#dcfce7;color:#166534;border:1px solid #bbf7d0}}
.pagination{{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-top:16px;padding-top:12px}}
.page-info{{font-size:13px;font-weight:600;color:var(--text-muted)}}
.page-btns{{display:flex;gap:6px;flex-wrap:wrap;justify-content:center}}
.mobile-cards{{display:none}}
.vcard{{background:#fff;border:1px solid var(--border);border-radius:14px;padding:14px;box-shadow:var(--shadow)}}
.vcard-top{{display:flex;justify-content:space-between;align-items:flex-start;gap:10px;margin-bottom:10px}}
.vcard-reg{{font-size:16px;font-weight:800;color:var(--rose-900);word-break:break-word}}
.vcard-name{{font-size:13px;font-weight:600;color:var(--text-sub);margin-top:2px}}
.vcard-grid{{display:grid;grid-template-columns:1fr 1fr;gap:8px}}
.vcard-item{{background:var(--rose-50);border-radius:10px;padding:8px 10px}}
.vcard-item span{{display:block;font-size:10px;font-weight:800;color:var(--rose-700);text-transform:uppercase;letter-spacing:.04em}}
.vcard-item b{{display:block;font-size:12px;font-weight:700;color:var(--text-main);margin-top:3px;word-break:break-word;white-space:normal}}
.vcard-status{{grid-column:1/-1}}
.page-btn{{padding:7px 13px;border-radius:8px;border:1px solid var(--border-dark);background:#fff;font-family:inherit;font-size:12px;font-weight:700;color:var(--text-sub);cursor:pointer;min-height:36px}}
.page-btn.active{{background:var(--rose-600);color:#fff;border-color:var(--rose-600)}}
.gsheet-box{{display:none;background:var(--rose-50);border:1px solid var(--rose-200);border-radius:12px;padding:14px;margin-bottom:16px}}
.gsheet-box.open{{display:flex;align-items:center;gap:10px;flex-wrap:wrap}}
.gsheet-input{{flex:1;min-width:0;padding:9px 12px;border:1px solid var(--border-dark);border-radius:8px;font-family:inherit;font-size:16px;min-height:40px}}
@media(max-width:900px){{
.stats-grid{{grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}}
.charts-grid{{grid-template-columns:1fr}}
}}
@media(max-width:768px){{
body{{padding:10px 8px 48px;padding-left:max(8px,env(safe-area-inset-left));padding-right:max(8px,env(safe-area-inset-right));padding-bottom:max(48px,env(safe-area-inset-bottom))}}
.top-header{{padding:12px;flex-direction:column;align-items:stretch;gap:10px}}
.brand{{gap:10px}}
.brand-icon{{width:40px;height:40px;font-size:18px}}
.brand-title{{font-size:16px}}
.brand-subtitle{{font-size:11px}}
.header-actions{{width:100%;display:grid;grid-template-columns:1fr 1fr;gap:8px}}
.live-badge{{grid-column:span 2;justify-content:center}}
.btn{{width:100%;min-height:44px}}
.stat-card{{padding:12px 10px}}
.stat-val{{font-size:22px}}
.stat-label{{font-size:10px}}
.stat-hint{{font-size:10px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.chart-card{{padding:12px}}
.chart-head{{gap:8px}}
.chart-title{{font-size:14px}}
.chart-container{{min-height:210px;height:210px}}
.card-sec{{padding:12px}}
.controls-bar{{flex-direction:column;align-items:stretch}}
.search-box{{width:100%;min-width:0}}
.filters-wrap{{width:100%;display:grid;grid-template-columns:1fr 1fr;gap:8px}}
.select-filter{{width:100%;font-size:14px;min-height:44px}}
.filters-wrap .btn-sec{{grid-column:span 2}}
.table-wrap{{display:none}}
.mobile-cards{{display:flex;flex-direction:column;gap:10px}}
.pagination{{flex-direction:column;align-items:center;gap:10px}}
.gsheet-box.open{{flex-direction:column;align-items:stretch}}
.gsheet-input{{width:100%;min-width:0}}
.pills-bar{{margin:0 -4px 14px;padding:0 4px 10px}}
}}
@media(max-width:420px){{
.vcard-grid{{grid-template-columns:1fr}}
.stat-val{{font-size:20px}}
}}
</style>
</head>
<body>
<div class="container">
<header class="top-header">
<div class="brand">
<div class="brand-icon">H</div>
<div>
<div class="brand-title">ASM HYUNDAI | BODYSHOP</div>
<div class="brand-subtitle">Operations Dashboard</div>
</div>
</div>
<div class="header-actions">
<span class="live-badge"><span class="live-dot" id="liveDot"></span><span id="syncText">Live Excel Sync</span></span>
<button class="btn btn-sec" onclick="toggleGSheet()">Google Sheet</button>
<button class="btn" onclick="triggerRefresh()">Refresh Data</button>
</div>
</header>
<div class="gsheet-box" id="gsheetBox">
<input type="text" id="gsheetUrl" class="gsheet-input" placeholder="Enter Google Sheets published CSV URL (File -> Share -> Publish to web -> CSV)">
<button class="btn" onclick="connectGSheet()">Connect Sheet</button>
<button class="btn btn-sec" onclick="toggleGSheet()">Close</button>
</div>
<div class="stats-grid">
<div class="stat-card"><div class="stat-label">Total Vehicles</div><div class="stat-val" id="kpiTotal">0</div><div class="stat-hint" id="kpiExcelFile">{fname}</div></div>
<div class="stat-card"><div class="stat-label">Ready For Delivery</div><div class="stat-val" style="color:#16a34a" id="kpiReady">0</div><div class="stat-hint">Ready to Handover</div></div>
<div class="stat-card"><div class="stat-label">In Bodyshop Work</div><div class="stat-val" style="color:#4f46e5" id="kpiWork">0</div><div class="stat-hint">Denting, Paint, Fitting</div></div>
<div class="stat-card"><div class="stat-label">Parts Pending</div><div class="stat-val" style="color:var(--rose-600)" id="kpiParts">0</div><div class="stat-hint">Waiting For Parts</div></div>
<div class="stat-card"><div class="stat-label">Survey / Estimate</div><div class="stat-val" style="color:#d97706" id="kpiSurvey">0</div><div class="stat-hint">Surveyor & Approvals</div></div>
<div class="stat-card"><div class="stat-label">Ageing 21+ Days</div><div class="stat-val" style="color:#dc2626" id="kpiAgeing">0</div><div class="stat-hint">Requires Attention</div></div>
</div>
<div class="charts-grid">
<div class="chart-card">
<div class="chart-head"><div class="chart-title">Workshop Stage Breakdown</div><div class="chart-badge">Pipeline</div></div>
<div class="chart-container"><canvas id="stageChart"></canvas></div>
</div>
<div class="chart-card">
<div class="chart-head"><div class="chart-title">Service Advisor Workload</div><div class="chart-badge" id="advCountHint">Advisors</div></div>
<div class="chart-container"><canvas id="advisorChart"></canvas></div>
</div>
<div class="chart-card">
<div class="chart-head"><div class="chart-title">Vehicle Ageing Distribution</div><div class="chart-badge">Days in Shop</div></div>
<div class="chart-container"><canvas id="ageingChart"></canvas></div>
</div>
<div class="chart-card">
<div class="chart-head"><div class="chart-title">Top Insurance Companies</div><div class="chart-badge">Volume</div></div>
<div class="chart-container"><canvas id="insChart"></canvas></div>
</div>
</div>
<div class="card-sec">
<div class="sec-head"><div class="sec-title">Vehicle Inventory Matrix</div><div class="stat-hint" id="rowCountHint">0 Vehicles Found</div></div>
<div class="pills-bar">
<div class="pill active" onclick="setPill('ALL',this)">All (<span id="pillAll">0</span>)</div>
<div class="pill" onclick="setPill('PARTS',this)">Parts Pending (<span id="pillParts">0</span>)</div>
<div class="pill" onclick="setPill('SURVEY',this)">Survey/Estimate (<span id="pillSurvey">0</span>)</div>
<div class="pill" onclick="setPill('WORK',this)">Bodyshop (<span id="pillWork">0</span>)</div>
<div class="pill" onclick="setPill('READY',this)">Ready (<span id="pillReady">0</span>)</div>
<div class="pill" onclick="setPill('AGE21',this)">Age 21+ (<span id="pillAge">0</span>)</div>
</div>
<div class="controls-bar">
<div class="search-box">
<svg class="search-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
<input type="text" id="searchInput" class="search-input" placeholder="Search by Reg No, Customer, Model, Advisor, Insurance..." oninput="handleSearch()">
</div>
<div class="filters-wrap">
<select id="advFilter" class="select-filter" onchange="applyFilters()"><option value="">All Advisors</option></select>
<select id="insFilter" class="select-filter" onchange="applyFilters()"><option value="">All Insurance</option></select>
<select id="pageLimit" class="select-filter" onchange="changePageSize()"><option value="25">25 per page</option><option value="50">50 per page</option><option value="100">100 per page</option><option value="9999">Show All</option></select>
<button class="btn btn-sec" onclick="resetFilters()">Reset Filters</button>
</div>
</div>
<div class="table-wrap">
<table>
<thead>
<tr>
<th>#</th>
<th>Reg No.</th>
<th>Customer Name</th>
<th>Model</th>
<th>Advisor</th>
<th>Status</th>
<th>Status Age</th>
<th>Total Age</th>
<th>Insurance Company</th>
<th>RO No.</th>
<th>Promise Date</th>
</tr>
</thead>
<tbody id="tableBody"></tbody>
</table>
</div>
<div class="mobile-cards" id="mobileCards"></div>
<div class="pagination">
<div class="page-info" id="pageInfo">Showing 0 of 0</div>
<div class="page-btns" id="pageBtns"></div>
</div>
</div>
</div>
<script>
let RAW_DATA = {jdata};
let filtered = [...RAW_DATA];
let curPage = 1;
let pageSize = 25;
let curPill = 'ALL';
let charts = {{}};

Chart.defaults.font.family = "'Plus Jakarta Sans', sans-serif";
Chart.defaults.color = '#334155';

function badgeClass(st){{
    let s = String(st||'');
    if(s.startsWith('17')) return 'badge-ready';
    if(s.startsWith('09')) return 'badge-parts';
    if(/^(11|12|13|14|15|16)/.test(s)) return 'badge-work';
    if(/^(00|01|02|03|04|05|07|08)/.test(s)) return 'badge-survey';
    return 'badge-other';
}}

function ageBadge(a){{
    let n = Number(a)||0;
    if(n>=21) return `<span class="badge badge-age-danger">${{n}}d</span>`;
    if(n>=14) return `<span class="badge badge-age-warning">${{n}}d</span>`;
    return `<span class="badge badge-age-ok">${{n}}d</span>`;
}}

function initFilters(){{
    let advs = [...new Set(RAW_DATA.map(d=>d.advisor).filter(Boolean))].sort();
    let inss = [...new Set(RAW_DATA.map(d=>d.insurance).filter(Boolean))].sort();
    let aEl = document.getElementById('advFilter');
    let iEl = document.getElementById('insFilter');
    let selA = aEl.value; let selI = iEl.value;
    aEl.innerHTML = '<option value="">All Advisors</option>' + advs.map(a=>`<option value="${{a}}">${{a}}</option>`).join('');
    iEl.innerHTML = '<option value="">All Insurance</option>' + inss.map(i=>`<option value="${{i}}">${{i}}</option>`).join('');
    aEl.value = selA; iEl.value = selI;
}}

function renderKPIs(){{
    let tot = RAW_DATA.length;
    let rdy = RAW_DATA.filter(d=>String(d.status).startsWith('17')).length;
    let wrk = RAW_DATA.filter(d=>/^(11|12|13|14|15|16)/.test(String(d.status))).length;
    let prt = RAW_DATA.filter(d=>String(d.status).startsWith('09')).length;
    let srv = RAW_DATA.filter(d=>/^(01|02|03|04|05)/.test(String(d.status))).length;
    let age = RAW_DATA.filter(d=>(Number(d.age)||0)>=21).length;

    document.getElementById('kpiTotal').textContent = tot;
    document.getElementById('kpiReady').textContent = rdy;
    document.getElementById('kpiWork').textContent = wrk;
    document.getElementById('kpiParts').textContent = prt;
    document.getElementById('kpiSurvey').textContent = srv;
    document.getElementById('kpiAgeing').textContent = age;

    document.getElementById('pillAll').textContent = tot;
    document.getElementById('pillParts').textContent = prt;
    document.getElementById('pillSurvey').textContent = srv;
    document.getElementById('pillWork').textContent = wrk;
    document.getElementById('pillReady').textContent = rdy;
    document.getElementById('pillAge').textContent = age;
}}

function renderGraphicalCharts(){{
    let stageCounts = {{'Ready':0, 'Parts Wait':0, 'Fitting':0, 'Final Inspect':0, 'Survey/Appr':0, 'Denting/Paint':0, 'Investigation':0}};
    RAW_DATA.forEach(d=>{{
        let s = String(d.status||'');
        if(s.startsWith('17')) stageCounts['Ready']++;
        else if(s.startsWith('09')) stageCounts['Parts Wait']++;
        else if(s.startsWith('15')) stageCounts['Fitting']++;
        else if(s.startsWith('16')) stageCounts['Final Inspect']++;
        else if(/^(01|02|03|04|05)/.test(s)) stageCounts['Survey/Appr']++;
        else if(/^(13|14)/.test(s)) stageCounts['Denting/Paint']++;
        else if(/^(06|18|19|20)/.test(s)) stageCounts['Investigation']++;
    }});

    let sLabels = Object.keys(stageCounts);
    let sVals = Object.values(stageCounts);
    let sColors = ['#10b981', '#f43f5e', '#6366f1', '#8b5cf6', '#f59e0b', '#fb7185', '#94a3b8'];

    if(charts.stage){{
        charts.stage.data.datasets[0].data = sVals;
        charts.stage.update();
    }} else {{
        let ctx = document.getElementById('stageChart').getContext('2d');
        charts.stage = new Chart(ctx, {{
            type: 'doughnut',
            data: {{
                labels: sLabels,
                datasets: [{{ data: sVals, backgroundColor: sColors, borderWidth: 2, borderColor: '#ffffff', hoverOffset: 6 }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{ position: 'bottom', labels: {{ boxWidth: 10, font: {{ size: 10, weight: '600' }}, padding: 8 }} }}
                }},
                cutout: '62%'
            }}
        }});
    }}

    let advMap = {{}};
    RAW_DATA.forEach(d=>{{ let a = d.advisor||'Other'; advMap[a] = (advMap[a]||0)+1; }});
    let sAdvs = Object.entries(advMap).sort((a,b)=>b[1]-a[1]).slice(0,7);
    let aLabels = sAdvs.map(x=>x[0]);
    let aVals = sAdvs.map(x=>x[1]);
    document.getElementById('advCountHint').textContent = Object.keys(advMap).length + ' Advisors';

    if(charts.advisor){{
        charts.advisor.data.labels = aLabels;
        charts.advisor.data.datasets[0].data = aVals;
        charts.advisor.update();
    }} else {{
        let ctx = document.getElementById('advisorChart').getContext('2d');
        charts.advisor = new Chart(ctx, {{
            type: 'bar',
            data: {{
                labels: aLabels,
                datasets: [{{
                    label: 'Vehicles',
                    data: aVals,
                    backgroundColor: '#fb7185',
                    hoverBackgroundColor: '#e11d48',
                    borderRadius: 8
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{ legend: {{ display: false }} }},
                scales: {{
                    y: {{ beginAtZero: true, grid: {{ color: '#ffe4e6' }} }},
                    x: {{ grid: {{ display: false }}, ticks: {{ font: {{ weight: '700', size: 11 }} }} }}
                }}
            }}
        }});
    }}

    let bMap = {{'0-7d':0, '8-14d':0, '15-20d':0, '21-30d':0, '31-60d':0, '60d+':0}};
    RAW_DATA.forEach(d=>{{
        let a = Number(d.age)||0;
        if(a<=7) bMap['0-7d']++;
        else if(a<=14) bMap['8-14d']++;
        else if(a<=20) bMap['15-20d']++;
        else if(a<=30) bMap['21-30d']++;
        else if(a<=60) bMap['31-60d']++;
        else bMap['60d+']++;
    }});
    let bLabels = Object.keys(bMap);
    let bVals = Object.values(bMap);
    let bColors = ['#10b981', '#34d399', '#fde047', '#fb923c', '#f43f5e', '#be123c'];

    if(charts.ageing){{
        charts.ageing.data.datasets[0].data = bVals;
        charts.ageing.update();
    }} else {{
        let ctx = document.getElementById('ageingChart').getContext('2d');
        charts.ageing = new Chart(ctx, {{
            type: 'bar',
            data: {{
                labels: bLabels,
                datasets: [{{
                    label: 'Vehicles',
                    data: bVals,
                    backgroundColor: bColors,
                    borderRadius: 8
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{ legend: {{ display: false }} }},
                scales: {{
                    y: {{ beginAtZero: true, grid: {{ color: '#ffe4e6' }} }},
                    x: {{ grid: {{ display: false }}, ticks: {{ font: {{ weight: '700', size: 11 }} }} }}
                }}
            }}
        }});
    }}

    let insMap = {{}};
    RAW_DATA.forEach(d=>{{ let ins = d.insurance||'Other'; insMap[ins] = (insMap[ins]||0)+1; }});
    let sIns = Object.entries(insMap).sort((a,b)=>b[1]-a[1]).slice(0,6);
    let iLabels = sIns.map(x=>x[0].length>18?x[0].slice(0,18)+'...':x[0]);
    let iVals = sIns.map(x=>x[1]);

    if(charts.ins){{
        charts.ins.data.labels = iLabels;
        charts.ins.data.datasets[0].data = iVals;
        charts.ins.update();
    }} else {{
        let ctx = document.getElementById('insChart').getContext('2d');
        charts.ins = new Chart(ctx, {{
            type: 'bar',
            data: {{
                labels: iLabels,
                datasets: [{{
                    label: 'Vehicles',
                    data: iVals,
                    backgroundColor: '#fda4af',
                    hoverBackgroundColor: '#f43f5e',
                    borderRadius: 8
                }}]
            }},
            options: {{
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{ legend: {{ display: false }} }},
                scales: {{
                    x: {{ beginAtZero: true, grid: {{ color: '#ffe4e6' }} }},
                    y: {{ grid: {{ display: false }}, ticks: {{ font: {{ weight: '700', size: 10 }} }} }}
                }}
            }}
        }});
    }}
}}

function applyFilters(){{
    let q = (document.getElementById('searchInput').value||'').toLowerCase().trim();
    let adv = document.getElementById('advFilter').value;
    let ins = document.getElementById('insFilter').value;

    filtered = RAW_DATA.filter(d=>{{
        if(adv && d.advisor !== adv) return false;
        if(ins && d.insurance !== ins) return false;
        if(curPill === 'PARTS' && !String(d.status).startsWith('09')) return false;
        if(curPill === 'SURVEY' && !/^(01|02|03|04|05)/.test(String(d.status))) return false;
        if(curPill === 'WORK' && !/^(11|12|13|14|15|16)/.test(String(d.status))) return false;
        if(curPill === 'READY' && !String(d.status).startsWith('17')) return false;
        if(curPill === 'AGE21' && (Number(d.age)||0)<21) return false;
        if(q){{
            let text = [d.reg, d.customer, d.model, d.advisor, d.insurance, d.ro, d.status].join(' ').toLowerCase();
            if(!text.includes(q)) return false;
        }}
        return true;
    }});

    curPage = 1;
    renderTable();
}}

function handleSearch(){{ applyFilters(); }}

function setPill(type, el){{
    curPill = type;
    document.querySelectorAll('.pill').forEach(p=>p.classList.remove('active'));
    el.classList.add('active');
    applyFilters();
}}

function resetFilters(){{
    document.getElementById('searchInput').value = '';
    document.getElementById('advFilter').value = '';
    document.getElementById('insFilter').value = '';
    curPill = 'ALL';
    document.querySelectorAll('.pill').forEach(p=>p.classList.remove('active'));
    document.querySelector('.pill').classList.add('active');
    applyFilters();
}}

function changePageSize(){{
    pageSize = Number(document.getElementById('pageLimit').value);
    curPage = 1;
    renderTable();
}}

function renderTable(){{
    let total = filtered.length;
    let start = (curPage-1)*pageSize;
    let pageItems = filtered.slice(start, start + pageSize);
    let cards = document.getElementById('mobileCards');

    document.getElementById('rowCountHint').textContent = total + ' Vehicles Found';
    let tbody = document.getElementById('tableBody');
    if(total === 0){{
        tbody.innerHTML = '<tr><td colspan="11" style="text-align:center;padding:30px;color:var(--text-muted)">No matching vehicles found</td></tr>';
        cards.innerHTML = '<div class="vcard" style="text-align:center;color:var(--text-muted)">No matching vehicles found</div>';
        document.getElementById('pageInfo').textContent = 'Showing 0 of 0';
        document.getElementById('pageBtns').innerHTML = '';
        return;
    }}

    tbody.innerHTML = pageItems.map((d,idx)=>{{
        return `<tr>
<td>${{start + idx + 1}}</td>
<td><strong style="color:var(--rose-900)">${{d.reg||'--'}}</strong></td>
<td>${{d.customer||'--'}}</td>
<td>${{d.model||'--'}}</td>
<td><span style="font-weight:700;color:var(--text-sub)">${{d.advisor||'--'}}</span></td>
<td><span class="badge ${{badgeClass(d.status)}}">${{d.status||'--'}}</span></td>
<td>${{d.statusage}}d</td>
<td>${{ageBadge(d.age)}}</td>
<td>${{d.insurance||'--'}}</td>
<td>${{d.ro||'--'}}</td>
<td>${{d.promise||'--'}}</td>
</tr>`;
    }}).join('');

    cards.innerHTML = pageItems.map((d,idx)=>{{
        return `<article class="vcard">
<div class="vcard-top">
<div>
<div class="vcard-reg">${{d.reg||'--'}}</div>
<div class="vcard-name">${{d.customer||'--'}} · ${{d.model||'--'}}</div>
</div>
${{ageBadge(d.age)}}
</div>
<div class="vcard-grid">
<div class="vcard-item vcard-status"><span>Status</span><b><span class="badge ${{badgeClass(d.status)}}">${{d.status||'--'}}</span></b></div>
<div class="vcard-item"><span>Advisor</span><b>${{d.advisor||'--'}}</b></div>
<div class="vcard-item"><span>Insurance</span><b>${{d.insurance||'--'}}</b></div>
<div class="vcard-item"><span>RO No.</span><b>${{d.ro||'--'}}</b></div>
<div class="vcard-item"><span>Promise</span><b>${{d.promise||'--'}}</b></div>
</div>
</article>`;
    }}).join('');

    let end = Math.min(start + pageSize, total);
    document.getElementById('pageInfo').textContent = `Showing ${{start + 1}} to ${{end}} of ${{total}}`;

    let totalPages = Math.ceil(total/pageSize);
    let btns = '';
    if(totalPages > 1){{
        if(curPage > 1) btns += `<button class="page-btn" onclick="gotoPage(${{curPage-1}})">Prev</button>`;
        for(let p=1; p<=Math.min(totalPages,7); p++){{
            btns += `<button class="page-btn ${{p===curPage?'active':''}}" onclick="gotoPage(${{p}})">${{p}}</button>`;
        }}
        if(totalPages > 7) btns += `<span style="padding:4px">...</span><button class="page-btn" onclick="gotoPage(${{totalPages}})">${{totalPages}}</button>`;
        if(curPage < totalPages) btns += `<button class="page-btn" onclick="gotoPage(${{curPage+1}})">Next</button>`;
    }}
    document.getElementById('pageBtns').innerHTML = btns;
}}

function gotoPage(p){{ curPage = p; renderTable(); }}

function toggleGSheet(){{
    let b = document.getElementById('gsheetBox');
    b.classList.toggle('open');
}}

function connectGSheet(){{
    let url = document.getElementById('gsheetUrl').value.trim();
    if(!url){{ alert('Please enter Google Sheets CSV URL'); return; }}
    fetchGSheetCSV(url);
}}

function parseCSV(text){{
    let lines = text.split(/\\r?\\n/).filter(l=>l.trim());
    if(lines.length < 3) return [];
    let rows = lines.map(line=>{{
        let res = []; let inQuote = false; let cur = '';
        for(let i=0; i<line.length; i++){{
            let c = line[i];
            if(c === '"') inQuote = !inQuote;
            else if(c === ',' && !inQuote){{ res.push(cur.trim()); cur = ''; }}
            else cur += c;
        }}
        res.push(cur.trim());
        return res;
    }});
    let data = [];
    for(let i=2; i<rows.length; i++){{
        let r = rows[i];
        if(!r[1] && !r[3]) continue;
        let st = r[15]||'';
        if(!st) continue;
        data.push({{
            id: i-1,
            customer: r[1]||'',
            contact: r[2]||'',
            reg: r[3]||'',
            model: r[4]||'',
            advisor: r[5]||'',
            ro: r[6]||'#N/A',
            rodate: r[7]||'',
            rostatus: r[9]||'#N/A',
            promise: r[10]||'',
            insurance: r[12]||'Not Specified',
            repair: r[13]||'MEDIUM',
            statusage: Number(r[14])||0,
            status: st,
            group: st.slice(0,25),
            statusdate: r[16]||'',
            panel: r[17]||'',
            inward: r[18]||'',
            age: Number(r[19])||0,
            bucket: r[20]||''
        }});
    }}
    return data;
}}

function fetchGSheetCSV(url){{
    document.getElementById('syncText').textContent = 'Fetching Google Sheet...';
    fetch(url)
    .then(r=>r.text())
    .then(t=>{{
        let ndata = parseCSV(t);
        if(ndata.length > 0){{
            RAW_DATA = ndata;
            document.getElementById('syncText').textContent = 'Google Sheet Live (' + ndata.length + ')';
            document.getElementById('kpiExcelFile').textContent = 'Google Sheet Linked';
            initFilters();
            renderKPIs();
            renderGraphicalCharts();
            applyFilters();
            toggleGSheet();
        }}
    }})
    .catch(e=>{{
        alert('Could not fetch Google Sheet CSV: ' + e);
        document.getElementById('syncText').textContent = 'Sheet Error';
    }});
}}

function pollLocalServer(){{
    if(location.protocol === 'http:' || location.protocol === 'https:'){{
        fetch('/api/data')
        .then(r=>r.json())
        .then(res=>{{
            if(res && res.data && res.data.length > 0){{
                if(JSON.stringify(RAW_DATA) !== JSON.stringify(res.data)){{
                    RAW_DATA = res.data;
                    document.getElementById('syncText').textContent = 'Live Excel Sync (' + res.data.length + ')';
                    initFilters();
                    renderKPIs();
                    renderGraphicalCharts();
                    applyFilters();
                }}
            }}
        }})
        .catch(()=>{{}});
    }}
}}

function triggerRefresh(){{
    if(location.protocol.startsWith('http')){{
        pollLocalServer();
    }} else {{
        location.reload();
    }}
}}

window.addEventListener('DOMContentLoaded', ()=>{{
    initFilters();
    renderKPIs();
    if(typeof Chart !== 'undefined') renderGraphicalCharts();
    applyFilters();
    if(location.protocol.startsWith('http')){{
        setInterval(pollLocalServer, 3000);
    }}
}});
</script>
</body>
</html>"""

def main():
    efile = get_excel_file()
    if not efile:
        print("Excel file not found")
        return
    bname = os.path.basename(efile)
    data = read_data(efile)
    html = build_html(data, bname)
    outpath = os.path.join(folder, "DVPR_Live_Dashboard.html")
    with open(outpath, "w", encoding="utf-8") as f:
        f.write(html)
    idxpath = os.path.join(folder, "index.html")
    with open(idxpath, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"OK: {len(data)} rows written to {outpath} and {idxpath}")

if __name__ == "__main__":
    main()
