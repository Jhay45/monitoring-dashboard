# monitoring-dashboard
Real-time online monitoring dashboard with live metrics, system status, and performance analytics
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Sorsogon 1 Area — Monitoring Dashboard</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{--bg:#161b2c;--card:rgba(30,36,59,.82);--text:#e6e9f8;--muted:#8a94b8;--gold:#d4af37;--green:#10b981;--red:#ef4444;--blue:#3b82f6;--orange:#f97316;--pink:#ec4899;--teal:#14b8a6;--sky:#0ea5e9;--lime:#84cc16;--slate:#64748b;--amber:#f59e0b;--border:rgba(138,148,184,.22)}
*{margin:0;padding:0;box-sizing:border-box;font-family:'Inter',sans-serif}
body{min-height:100vh;padding:12px;padding-bottom:40px;color:var(--text);background:radial-gradient(ellipse 600px 400px at 15% 8%,rgba(212,175,55,.18),transparent 60%),radial-gradient(ellipse 500px 500px at 90% 20%,rgba(59,130,246,.14),transparent 60%),radial-gradient(ellipse 700px 500px at 50% 100%,rgba(139,92,246,.12),transparent 60%),linear-gradient(160deg,#0b0f1e,#161b2c 45%,#1a1f38);background-attachment:fixed}
.container{max-width:1100px;margin:0 auto}
h1{font-size:18px;font-weight:800;margin-bottom:2px}
.subtitle{font-size:11px;color:var(--muted);margin-bottom:12px}
.save-status{background:rgba(37,44,72,.55);border:1px solid var(--border);border-radius:10px;padding:9px 11px;margin-bottom:10px;display:flex;align-items:center;gap:8px;font-size:12px;flex-wrap:wrap}
.branch-row{display:flex;gap:5px;margin-bottom:10px;flex-wrap:wrap}
.branch-btn-top{padding:8px 6px;border-radius:9px;font-weight:700;border:2px solid var(--border);background:rgba(30,36,59,.7);color:var(--text);cursor:pointer;font-size:11px;flex:1;min-width:68px}
.branch-btn-top.active{border-color:var(--gold);background:rgba(212,175,55,.18);color:var(--gold)}
.branch-btn-top .bc{display:block;font-size:9px;opacity:.7}
.clear-btn-group{position:relative;margin-left:auto}
.clear-data-btn{background:var(--red);color:#fff;border:none;border-radius:7px;padding:7px 10px;font-weight:700;font-size:11px;cursor:pointer;min-height:34px}
.clear-dropdown{position:absolute;top:100%;right:0;margin-top:6px;background:rgba(30,36,59,.97);border:1px solid var(--border);border-radius:10px;padding:11px;min-width:210px;z-index:100;display:none;box-shadow:0 8px 24px rgba(0,0,0,.4)}
.clear-dropdown.show{display:block}
.clear-dropdown label{font-size:11px;color:var(--muted);margin-bottom:5px;display:block;font-weight:600}
.clear-dropdown select{width:100%;margin-bottom:9px;padding:6px;border-radius:6px;background:rgba(37,44,72,.7);border:1px solid var(--border);color:var(--text)}
.clear-confirm-btn{background:var(--red);color:#fff;border:none;border-radius:6px;padding:8px;font-weight:700;width:100%;cursor:pointer;min-height:36px}
.clear-all-btn{background:transparent;color:var(--muted);border:1px solid var(--muted);border-radius:6px;padding:6px;font-size:11px;margin-top:7px;width:100%;cursor:pointer}
.tabs{display:flex;gap:5px;margin-bottom:14px;flex-wrap:wrap;position:sticky;top:0;background:rgba(11,15,30,.9);padding:7px 0;z-index:10;border-radius:11px}
.tab-btn{padding:9px 6px;border-radius:9px;font-weight:700;border:none;cursor:pointer;font-size:11px;flex:1;min-width:72px}
.tab-btn.active{transform:scale(1.04);box-shadow:0 0 12px rgba(0,0,0,.4)}
.tb-over{background:#8b5cf6;color:#fff}.tb-nm{background:var(--gold);color:#000}.tb-bd{background:var(--orange);color:#fff}.tb-del{background:var(--red);color:#fff}.tb-kam{background:var(--pink);color:#fff}.tb-oc{background:var(--sky);color:#fff}.tb-staff{background:var(--lime);color:#000}.tb-inc{background:var(--amber);color:#000}.tb-upd{background:var(--slate);color:#fff}.tb-sync{background:var(--teal);color:#fff}.tb-exc{background:#475569;color:#fff}
.tab-content{display:none}.tab-content.active{display:block}
.card{background:var(--card);border-radius:14px;padding:15px;margin-bottom:14px;border:1px solid var(--border)}
.card-title{font-size:13px;font-weight:700;margin-bottom:12px;display:flex;justify-content:space-between;flex-wrap:wrap;gap:5px}
.badge{font-size:9px;background:rgba(16,185,129,.15);color:#4ade80;padding:2px 7px;border-radius:9px;font-weight:600}
.badge-orange{background:rgba(249,115,22,.15);color:#fdba74}
.grid-form{display:grid;grid-template-columns:repeat(auto-fit,minmax(125px,1fr));gap:9px;margin-bottom:12px}
.grid-form label{font-size:10px;color:var(--muted);margin-bottom:3px;display:block;font-weight:600}
input,select,textarea{width:100%;background:rgba(37,44,72,.7);border:1px solid var(--border);color:var(--text);padding:9px 10px;border-radius:7px;font-size:13px}
input:focus,select:focus,textarea:focus{outline:none;border-color:var(--gold)}
input[readonly]{color:#94a3b8;background:rgba(0,0,0,.25)}
.btn-row{display:flex;gap:5px;flex-wrap:wrap;margin-top:5px}
button{padding:9px 12px;border-radius:7px;font-weight:700;border:none;cursor:pointer;font-size:11px;min-height:36px}
.btn-save{background:var(--green);color:#fff}.btn-update{background:var(--blue);color:#fff}.btn-cancel{background:transparent;border:1px solid var(--muted);color:var(--muted)}.btn-clear{background:transparent;border:1px solid var(--gold);color:var(--gold)}.btn-export{background:transparent;border:1px solid var(--blue);color:var(--blue)}.btn-orange{background:var(--orange);color:#fff}.btn-red{background:var(--red);color:#fff}.btn-pink{background:var(--pink);color:#fff}.btn-sky{background:var(--sky);color:#fff}.btn-teal{background:var(--teal);color:#fff}.btn-amber{background:var(--amber);color:#000}.btn-lock{background:var(--gold);color:#000}.btn-unlock{background:transparent;border:1px solid var(--gold);color:var(--gold)}.btn-view-on{background:var(--blue);color:#fff}.btn-view-off{background:transparent;border:1px solid var(--muted);color:var(--muted)}.btn-exc{background:#475569;color:#fff}
.hidden{display:none!important}
table{width:100%;border-collapse:collapse;font-size:10px;margin-top:8px}
th,td{padding:7px 5px;text-align:left;border-bottom:1px solid var(--border)}
th{color:var(--muted);font-weight:700;text-transform:uppercase;font-size:8px}
.text-center{text-align:center}
.action-btns{display:flex;gap:4px;justify-content:center}
.btn-edit{background:transparent;border:1px solid var(--blue);color:var(--blue);padding:5px 8px;border-radius:5px;font-size:10px;min-height:28px}
.btn-del{background:transparent;border:1px solid var(--red);color:var(--red);padding:5px 8px;border-radius:5px;font-size:10px;min-height:28px}
.filter-bar{display:flex;gap:7px;flex-wrap:wrap;align-items:center;background:rgba(37,44,72,.55);padding:10px;border-radius:10px;border:1px solid var(--border);margin-bottom:10px}
.filter-bar label{font-size:10px;color:var(--muted);font-weight:600}
.filter-bar select,.filter-bar input{padding:7px 9px;border-radius:7px;font-size:11px;background:rgba(30,36,59,.8);border:1px solid var(--border);color:var(--text);min-width:90px;width:auto}
.btn-apply{background:var(--blue);color:#fff}.btn-reset{background:transparent;border:1px solid var(--muted);color:var(--muted)}.btn-up{background:var(--green);color:#fff}.btn-tpl{background:transparent;border:1px solid var(--teal);color:var(--teal)}
.upload-result{margin-top:9px;padding:9px;border-radius:7px;font-size:11px;display:none}
.upload-success{background:rgba(16,185,129,.1);border:1px solid rgba(16,185,129,.3);color:#4ade80;display:block}
.upload-error{background:rgba(239,68,68,.1);border:1px solid rgba(239,68,68,.3);color:#fca5a5;display:block}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:8px;margin-bottom:12px}
.kpi-card{background:rgba(37,44,72,.6);border-radius:10px;padding:11px;text-align:center;border:1px solid var(--border)}
.kpi-card .kl{font-size:9px;color:var(--muted);margin-bottom:4px;font-weight:600;text-transform:uppercase}
.kpi-card .kv{font-size:18px;font-weight:800}
.kg{color:var(--green)}.ko{color:var(--orange)}.kr{color:var(--red)}.kp{color:var(--pink)}.kb{color:var(--blue)}.ks{color:var(--sky)}.klime{color:var(--lime)}.kslate{color:var(--slate)}.kamb{color:var(--amber)}.kexc{color:#94a3b8}
.total-row td{font-weight:800;background:rgba(212,175,55,.1);color:var(--gold);border-top:2px solid var(--gold)}
.week-row{display:flex;gap:7px;align-items:center;flex-wrap:wrap;margin-bottom:10px}
.week-row label{font-size:10px;color:var(--muted);font-weight:600}
.week-row select{width:auto;min-width:130px;padding:7px 9px;font-size:11px}
.lock-tag{display:inline-block;background:rgba(212,175,55,.15);color:var(--gold);padding:2px 6px;border-radius:5px;font-size:9px;font-weight:700;margin-left:4px}
.locked-row{opacity:.7;background:rgba(212,175,55,.05)}
.sync-box{width:100%;min-height:80px;font-family:monospace;font-size:10px;background:rgba(37,44,72,.5);border:1px solid var(--border);border-radius:7px;color:var(--text);padding:10px;margin:8px 0}
.sync-note{font-size:10px;color:var(--muted);margin-top:6px}
.ok{color:var(--green);font-weight:800}.no{color:var(--red);font-weight:800}.partial{color:var(--orange);font-weight:700}
.rank-1{background:rgba(212,175,55,.12)!important}.rank-2{background:rgba(192,192,192,.08)!important}.rank-3{background:rgba(205,127,50,.08)!important}
.table-wrap{overflow-x:auto}
.branch-label{font-size:10px;color:var(--gold);font-weight:700;margin-bottom:7px}
.client-pair{display:flex;gap:5px;align-items:center;margin-bottom:6px}
.client-pair input{flex:1;min-width:0}.client-pair select{flex:1;max-width:170px}
.add-client-btn{background:transparent;border:1px dashed var(--border);color:var(--muted);padding:6px 10px;border-radius:6px;font-size:10px;width:100%;margin-bottom:10px;min-height:30px}
.remove-client-btn{background:transparent;border:1px solid var(--red);color:var(--red);padding:4px 8px;border-radius:5px;font-size:10px;min-height:28px;flex:0 0 auto}

.staff-btn-wrap{position:relative;margin-left:auto}
.staff-list-btn{background:var(--blue);color:#fff;border:none;border-radius:7px;padding:7px 12px;font-weight:700;font-size:11px;cursor:pointer;min-height:34px}
.staff-dropdown{position:absolute;top:100%;right:0;margin-top:6px;background:rgba(30,36,59,.97);border:1px solid var(--border);border-radius:10px;padding:11px;min-width:260px;max-width:340px;z-index:100;display:none;box-shadow:0 8px 24px rgba(0,0,0,.4)}
.staff-dropdown.show{display:block}
.staff-dropdown label{font-size:11px;color:var(--muted);margin-bottom:5px;display:block;font-weight:600}
.staff-row-item{display:flex;gap:5px;align-items:center;margin-bottom:6px}
.staff-row-item input{flex:1;padding:6px;border-radius:5px;background:rgba(37,44,72,.7);border:1px solid var(--border);color:var(--text);font-size:11px}
.staff-row-item select{flex:1;max-width:130px;padding:6px;border-radius:5px;background:rgba(37,44,72,.7);border:1px solid var(--border);color:var(--text);font-size:11px}
.staff-save-row{margin-top:10px;display:flex;gap:6px}
.btn-sm{padding:4px 7px;font-size:10px;min-height:26px}
</style>
</head>
<body>
<div class="container">
<h1>📊 Sorsogon 1 Area — Monitoring Dashboard</h1>
<p class="subtitle">Prepared by Jayruben Garlando — Area Head</p>

<div class="save-status">
<span>📍 <strong id="curBranchLabel">Consolidated — All Branches</strong></span>
<span style="color:var(--muted)">| ✅ Data Saved Locally</span>
<div class="staff-btn-wrap"><button class="staff-list-btn" onclick="toggleStaffDropdown()">👤 Staff List</button>
<div class="staff-dropdown" id="staffDropdown">
<label>📋 Manage Staff — Name & Assigned Branch</label>
<div id="staffListContainer"></div>
<button class="btn-save btn-sm" style="width:100%;margin:8px 0" onclick="addStaffRow()">+ ADD NEW STAFF</button>
<div class="staff-save-row">
<button class="btn-save" style="flex:1" onclick="saveStaffList()">💾 SAVE LIST</button>
<button class="btn-clear" style="flex:1" onclick="resetStaffList()">↩️ RESET</button>
</div>
<small style="margin-top:8px;display:block;color:var(--muted)">Changes update all staff dropdowns automatically.</small>
</div></div>
<div class="clear-btn-group">
<button class="clear-data-btn" onclick="toggleClearDropdown()">🔒 Clear Data</button>
<div class="clear-dropdown" id="clearDropdown">
<label>Select Month:</label>
<select id="clearMonthSelect" onchange="updateClearDays()">
<option value="">-- Choose Month --</option>
<option value="01">January</option><option value="02">February</option><option value="03">March</option>
<option value="04">April</option><option value="05">May</option><option value="06">June</option>
<option value="07">July</option><option value="08">August</option><option value="09">September</option>
<option value="10">October</option><option value="11">November</option><option value="12">December</option>
</select>
<label>Select Date (Day):</label>
<select id="clearDaySelect"><option value="">-- All Days --</option></select>
<button class="clear-confirm-btn" onclick="clearDataByMonth()">⚠️ CLEAR SELECTED MONTH / DATE</button>
<button class="clear-all-btn" onclick="clearAllData()">❌ CLEAR ALL DATA</button>
</div>
</div>
</div>

<div class="branch-row">
<button class="branch-btn-top active" onclick="selectBranch('',this)">📊 Consolidated<span class="bc">ALL</span></button>
<button class="branch-btn-top" onclick="selectBranch('B0814',this)">Donsol<span class="bc">B0814</span></button>
<button class="branch-btn-top" onclick="selectBranch('B0392',this)">Pilar 1<span class="bc">B0392</span></button>
<button class="branch-btn-top" onclick="selectBranch('B0739',this)">Castilla<span class="bc">B0739</span></button>
<button class="branch-btn-top" onclick="selectBranch('B1659',this)">Castilla 2<span class="bc">B1659</span></button>
<button class="branch-btn-top" onclick="selectBranch('B1812',this)">Pilar 2<span class="bc">B1812</span></button>
</div>

<div class="tabs">
<button class="tab-btn tb-over active" onclick="switchTab('overview',this)">📊 Overview</button>
<button class="tab-btn tb-nm" onclick="switchTab('nmoffset',this)">📈 NM & Off</button>
<button class="tab-btn tb-bd" onclick="switchTab('baddebt',this)">💰 Bad Debt</button>
<button class="tab-btn tb-del" onclick="switchTab('delinquent',this)">⚠️ Delinq</button>
<button class="tab-btn tb-exc" onclick="switchTab('excuses',this)">📝 Excuses</button>
<button class="tab-btn tb-kam" onclick="switchTab('kamustahan',this)">🤝 Kamust</button>
<button class="tab-btn tb-oc" onclick="switchTab('offset',this)">📋 Offset</button>
<button class="tab-btn tb-staff" onclick="switchTab('staff',this)">👥 Staff Perf</button>
<button class="tab-btn tb-inc" onclick="switchTab('incentives',this)">🎁 Incentives</button>
<button class="tab-btn tb-upd" onclick="switchTab('update',this)">✅ Update</button>
<button class="tab-btn tb-sync" onclick="switchTab('sync',this)">🔄 Sync</button>
</div>

<!-- OVERVIEW -->
<div class="tab-content active" id="tab-overview">
<div class="card">
<p class="card-title" id="ovTitle">📊 OVERVIEW — All Branches</p>
<div class="filter-bar">
<label>📅 Month:</label>
<select id="filterMonth" onchange="updateDayOptions()">
<option value="">-- All --</option><option value="01">Jan</option><option value="02">Feb</option><option value="03">Mar</option>
<option value="04">Apr</option><option value="05">May</option><option value="06">Jun</option>
<option value="07">Jul</option><option value="08">Aug</option><option value="09">Sep</option>
<option value="10">Oct</option><option value="11">Nov</option><option value="12">Dec</option>
</select>
<label>📆 Day:</label><select id="filterDay"><option value="">-- All --</option></select>
<button class="btn-apply" onclick="applyDateFilter()">🔍 APPLY</button>
<button class="btn-reset" onclick="resetDateFilter()">↩️ RESET</button>
<div style="width:100%;height:1px;background:var(--border);margin:3px 0"></div>
<label>📁 Excel — Month:</label>
<select id="uploadMonth"><option value="">-- Select --</option>
<option value="01">Jan</option><option value="02">Feb</option><option value="03">Mar</option><option value="04">Apr</option>
<option value="05">May</option><option value="06">Jun</option><option value="07">Jul</option><option value="08">Aug</option>
<option value="09">Sep</option><option value="10">Oct</option><option value="11">Nov</option><option value="12">Dec</option>
</select>
<input type="number" id="uploadDay" min="1" max="31" placeholder="Day" style="width:60px">
<input type="file" id="excelFile" accept=".xlsx,.xls,.csv" style="display:none" onchange="handleFileUpload(event)">
<button class="btn-up" onclick="document.getElementById('excelFile').click()">⬆️ UPLOAD</button>
<button class="btn-tpl" onclick="downloadTemplate()">📋 TEMPLATE</button>
</div>
<div id="uploadResult" class="upload-result"></div>
<div class="kpi-grid" id="overviewKPI"></div>
<div class="table-wrap"><table>
<thead><tr><th>Branch</th><th>NM</th><th>Off</th><th>Net</th><th>BD ₱</th><th>Tgt ₱</th><th>%</th><th>Del</th><th>Kam</th><th>OffC</th><th>Exc</th></tr></thead>
<tbody id="overviewBody"></tbody>
</table></div>
</div>
</div>

<!-- NM & OFFSET -->
<div class="tab-content" id="tab-nmoffset">
<div class="card" id="nmFormCard">
<p class="card-title"><span>📈 NM & Offset — Daily Entry</span><span class="badge">PERMANENT</span></p>
<div class="branch-label" id="nmBranchLabel"></div>
<div class="grid-form">
<div style="grid-column:1/-1"><label>Staff Name</label><select id="nmStaff"><option value="">-- Select Staff --</option></select></div>
<div><label>Month</label><select id="nmMonth" onchange="updateDays('nm')"><option value="">-- Month --</option>
<option value="01">Jan</option><option value="02">Feb</option><option value="03">Mar</option><option value="04">Apr</option>
<option value="05">May</option><option value="06">Jun</option><option value="07">Jul</option><option value="08">Aug</option>
<option value="09">Sep</option><option value="10">Oct</option><option value="11">Nov</option><option value="12">Dec</option>
</select></div>
<div><label>Date</label><select id="nmDay"><option value="">-- Month first --</option></select></div>
<div><label>New Member</label><input type="number" id="nmValue" value="0" min="0" oninput="calcNetNM()"></div>
<div><label>Offset</label><input type="number" id="offsetValue" value="0" min="0" oninput="calcNetNM()"></div>
<div><label>Target NM</label><input type="number" id="targetNM" value="0" min="0"></div>
<div><label>Target Offset</label><input type="number" id="targetOffset" value="0" min="0"></div>
<div style="grid-column:1/-1"><label>Net NM (Auto)</label><input type="text" id="netNM" value="0" readonly style="color:var(--green);font-weight:700"></div>
</div>
<div class="btn-row">
<button class="btn-save" id="nmSaveBtn" onclick="saveNM()">💾 SAVE</button>
<button class="btn-update hidden" id="nmUpdateBtn" onclick="updateNM()">🔄 UPDATE</button>
<button class="btn-cancel hidden" id="nmCancelBtn" onclick="cancelEditNM()">✕ CANCEL</button>
<button class="btn-clear" onclick="clearNMForm()">🔄 Clear</button>
<button class="btn-export" onclick="exportNM()">📂 Export</button>
</div>
</div>
<div class="card">
<p class="card-title">📋 NM & Offset Records</p>
<div class="table-wrap"><table>
<thead><tr><th>#</th><th>Branch</th><th>Staff</th><th>Mo</th><th>Date</th><th>NM</th><th>Off</th><th>Net</th><th>TgtN</th><th>TgtO</th><th class="text-center">Act</th></tr></thead>
<tbody id="nmBody"><tr><td colspan="11" class="text-center" style="color:var(--muted);padding:20px">No entries yet.</td></tr></tbody>
</table></div>
</div>
</div>

<!-- BAD DEBT -->
<div class="tab-content" id="tab-baddebt">
<div class="card" id="bdFormCard">
<p class="card-title"><span>💰 BAD DEBT — Daily Entry</span><span class="badge badge-orange">PERMANENT</span></p>
<div class="branch-label" id="bdBranchLabel"></div>
<div class="grid-form">
<div style="grid-column:1/-1"><label>Staff Name</label><select id="bdStaff" onchange="autoFillBranchFromStaff('bd')"><option value="">-- Select Staff --</option></select></div>
<div><label>Month</label><select id="bdMonth" onchange="updateDays('bd')"><option value="">-- Month --</option>
<option value="01">Jan</option><option value="02">Feb</option><option value="03">Mar</option><option value="04">Apr</option>
<option value="05">May</option><option value="06">Jun</option><option value="07">Jul</option><option value="08">Aug</option>
<option value="09">Sep</option><option value="10">Oct</option><option value="11">Nov</option><option value="12">Dec</option>
</select></div>
<div><label>Date</label><select id="bdDay"><option value="">-- Month first --</option></select></div>
<div style="grid-column:1/-1"><label>Client / Remarks</label><input type="text" id="bdClient" placeholder="Client name or remarks"></div>
<div><label>Amount Collected ₱</label><input type="number" id="bdAmount" value="0" min="0" step="0.01" oninput="calcBalanceBD()"></div>
<div><label>Week</label><select id="bdWorkingDay"><option value="">-- Select --</option>
<option value="Week 1">Week 1</option><option value="Week 2">Week 2</option><option value="Week 3">Week 3</option>
<option value="Week 4">Week 4</option><option value="Week 5">Week 5</option>
</select></div>
<div><label>Target ₱</label><input type="number" id="bdTarget" value="0" min="0" step="0.01" oninput="calcBalanceBD()"></div>
<div style="grid-column:1/-1"><label>Remaining Balance (Auto)</label><input type="text" id="bdBalance" value="0" readonly style="color:var(--orange);font-weight:700"></div>
</div>
<div class="btn-row">
<button class="btn-orange" id="bdSaveBtn" onclick="saveBD()">💾 SAVE</button>
<button class="btn-update hidden" id="bdUpdateBtn" onclick="updateBD()">🔄 UPDATE</button>
<button class="btn-cancel hidden" id="bdCancelBtn" onclick="cancelEditBD()">✕ CANCEL</button>
<button class="btn-clear" onclick="clearBDForm()">🔄 Clear</button>
<button class="btn-export" onclick="exportBD()">📂 Export</button>
</div>
</div>
<div class="card">
<p class="card-title">📊 Collection Summary <span id="bdWeekLabel"></span></p>
<div class="week-row"><label>🗓️ Week Filter:</label><select id="bdWeekFilter" onchange="renderBD()">
<option value="">-- All Weeks --</option>
<option value="Week 1">Week 1</option><option value="Week 2">Week 2</option>
<option value="Week 3">Week 3</option><option value="Week 4">Week 4</option><option value="Week 5">Week 5</option>
</select></div>
<div class="kpi-grid" id="bdSummary"></div>
<div class="table-wrap"><table>
<thead><tr><th>#</th><th>Branch</th><th>Staff</th><th>Mo</th><th>Date</th><th>Week</th><th>Client</th><th>Amount</th><th>Tgt</th><th class="text-center">Act</th></tr></thead>
<tbody id="bdBody"><tr><td colspan="10" class="text-center" style="color:var(--muted);padding:20px">No entries yet.</td></tr></tbody>
</table></div>
</div>
</div>

<!-- DELINQUENT -->
<div class="tab-content" id="tab-delinquent">
<div class="card">
<p class="card-title"><span>⚠️ DELINQUENT — Daily Entry</span><span class="badge" style="background:rgba(239,68,68,.15);color:#fca5a5">PERMANENT</span></p>
<div class="branch-label" id="delBranchLabel"></div>
<div class="grid-form">
<div><label>Month</label><select id="delMonth" onchange="updateDays('del')"><option value="">-- Month --</option>
<option value="01">Jan</option><option value="02">Feb</option><option value="03">Mar</option><option value="04">Apr</option>
<option value="05">May</option><option value="06">Jun</option><option value="07">Jul</option><option value="08">Aug</option>
<option value="09">Sep</option><option value="10">Oct</option><option value="11">Nov</option><option value="12">Dec</option>
</select></div>
<div><label>Date</label><select id="delDay"><option value="">-- Month first --</option></select></div>
<div style="grid-column:1/-1"><label>Staff Name</label><select id="delStaff" onchange="autoFillBranchFromStaff('del')"><option value="">-- Select Staff --</option></select></div>
</div>
<div id="delClientsContainer"></div>
<button class="add-client-btn" onclick="addDelClientRow()">➕ Add Client / Reason</button>
<div class="btn-row">
<button class="btn-red" id="delSaveBtn" onclick="saveDel()">💾 SAVE</button>
<button class="btn-update hidden" id="delUpdateBtn" onclick="updateDel()">🔄 UPDATE</button>
<button class="btn-cancel hidden" id="delCancelBtn" onclick="cancelEditDel()">✕ CANCEL</button>
<button class="btn-clear" onclick="clearDelForm()">🔄 Clear</button>
<button class="btn-export" onclick="exportDel()">📂 Export</button>
</div>
</div>
<div class="card">
<p class="card-title">📋 Delinquent Records</p>
<div class="kpi-grid" id="delSummary"></div>
<div class="table-wrap"><table>
<thead><tr><th>#</th><th>Branch</th><th>Mo</th><th>Date</th><th>Staff</th><th>Clients & Reasons</th><th class="text-center">Act</th></tr></thead>
<tbody id="delBody"><tr><td colspan="7" class="text-center" style="color:var(--muted);padding:20px">No entries yet.</td></tr></tbody>
</table></div>
</div>
</div>

<!-- EXCUSES (NO FILTER) -->
<div class="tab-content" id="tab-excuses">
<div class="card">
<p class="card-title"><span>📝 EXCUSES — Entry</span><span class="badge" style="background:rgba(100,116,139,.2);color:#cbd5e1">PERMANENT</span></p>
<div class="branch-label" id="excBranchLabel"></div>
<div class="grid-form">
<div><label>Month</label><select id="excMonth" onchange="updateDays('exc')"><option value="">-- Month --</option>
<option value="01">Jan</option><option value="02">Feb</option><option value="03">Mar</option><option value="04">Apr</option>
<option value="05">May</option><option value="06">Jun</option><option value="07">Jul</option><option value="08">Aug</option>
<option value="09">Sep</option><option value="10">Oct</option><option value="11">Nov</option><option value="12">Dec</option>
</select></div>
<div><label>Date</label><select id="excDay"><option value="">-- Month first --</option></select></div>
<div style="grid-column:1/-1"><label>Staff Name</label><select id="excStaff" onchange="autoFillBranchFromStaff('exc')"><option value="">-- Select Staff --</option></select></div>
<div style="grid-column:1/-1"><label>Name of Client(s)</label><input type="text" id="excClient" placeholder="Client name(s)"></div>
<div style="grid-column:1/-1"><label>Reason</label><select id="excReason">
<option value="">-- Select Reason --</option>
<option value="NCNO">NCNO</option>
<option value="NCNW">NCNW</option>
<option value="Sick">Sick</option>
<option value="Hospitalization">Hospitalization</option>
<option value="Family Problem">Family Problem</option>
<option value="Bankruptcy">Bankruptcy</option>
<option value="Death of client">Death of client</option>
<option value="Run away">Run away</option>
</select></div>
</div>
<div class="btn-row">
<button class="btn-exc" id="excSaveBtn" onclick="saveEXC()">💾 SAVE</button>
<button class="btn-update hidden" id="excUpdateBtn" onclick="updateEXC()">🔄 UPDATE</button>
<button class="btn-cancel hidden" id="excCancelBtn" onclick="cancelEditEXC()">✕ CANCEL</button>
<button class="btn-clear" onclick="clearEXCForm()">🔄 Clear</button>
<button class="btn-export" onclick="exportEXC()">📂 Export</button>
</div>
</div>
<div class="card">
<p class="card-title">📋 Excuses Records</p>
<div class="kpi-grid" id="excSummary">
<div class="kpi-card"><div class="kl">Total Excuses</div><div class="kv kexc" id="excTotalCount">0</div></div>
</div>
<div class="table-wrap"><table>
<thead><tr><th>#</th><th>Branch</th><th>Mo</th><th>Date</th><th>Staff</th><th>Client(s)</th><th>Reason</th><th class="text-center">Act</th></tr></thead>
<tbody id="excBody"><tr><td colspan="8" class="text-center" style="color:var(--muted);padding:20px">No entries yet.</td></tr></tbody>
</table></div>
</div>
</div>

<!-- KAMUSTAHAN -->
<div class="tab-content" id="tab-kamustahan">
<div class="card">
<p class="card-title"><span>🤝 KAMUSTAHAN — Daily Entry</span><span class="badge" style="background:rgba(236,72,153,.15);color:#f9a8d4">PERMANENT</span></p>
<div class="branch-label" id="kamBranchLabel"></div>
<div class="grid-form">
<div><label>Month</label><select id="kamMonth" onchange="updateDays('kam')"><option value="">-- Month --</option>
<option value="01">Jan</option><option value="02">Feb</option><option value="03">Mar</option><option value="04">Apr</option>
<option value="05">May</option><option value="06">Jun</option><option value="07">Jul</option><option value="08">Aug</option>
<option value="09">Sep</option><option value="10">Oct</option><option value="11">Nov</option><option value="12">Dec</option>
</select></div>
<div><label>Date</label><select id="kamDay"><option value="">-- Month first --</option></select></div>
<div style="grid-column:1/-1"><label>Staff Name</label><select id="kamStaff" onchange="autoFillBranchFromStaff('kam')"><option value="">-- Select Staff --</option></select></div>
<div><label>No. of Kamustahan</label><input type="number" id="kamCount" value="0" min="0"></div>
<div style="grid-column:1/-1"><label>Update / Remarks</label><textarea id="kamRemarks" rows="2" placeholder="Enter remarks..."></textarea></div>
</div>
<div class="btn-row">
<button class="btn-pink" id="kamSaveBtn" onclick="saveKam()">💾 SAVE</button>
<button class="btn-update hidden" id="kamUpdateBtn" onclick="updateKam()">🔄 UPDATE</button>
<button class="btn-cancel hidden" id="kamCancelBtn" onclick="cancelEditKam()">✕ CANCEL</button>
<button class="btn-clear" onclick="clearKamForm()">🔄 Clear</button>
<button class="btn-export" onclick="exportKam()">📂 Export</button>
</div>
</div>
<div class="card">
<p class="card-title">📋 Kamustahan Records</p>
<div class="kpi-grid" id="kamSummary"></div>
<div class="table-wrap"><table>
<thead><tr><th>#</th><th>Branch</th><th>Mo</th><th>Date</th><th>Staff</th><th>Count</th><th>Remarks</th><th class="text-center">Act</th></tr></thead>
<tbody id="kamBody"><tr><td colspan="8" class="text-center" style="color:var(--muted);padding:20px">No entries yet.</td></tr></tbody>
</table></div>
</div>
</div>

<!-- OFFSET CLIENT -->
<div class="tab-content" id="tab-offset">
<div class="card">
<p class="card-title"><span>📋 OFFSET CLIENT — Entry</span><span class="badge" style="background:rgba(14,165,233,.15);color:#7dd3fc">PERMANENT</span></p>
<div class="branch-label" id="ocBranchLabel"></div>
<div class="grid-form">
<div><label>Month</label><select id="ocMonth" onchange="updateDays('oc')"><option value="">-- Month --</option>
<option value="01">Jan</option><option value="02">Feb</option><option value="03">Mar</option><option value="04">Apr</option>
<option value="05">May</option><option value="06">Jun</option><option value="07">Jul</option><option value="08">Aug</option>
<option value="09">Sep</option><option value="10">Oct</option><option value="11">Nov</option><option value="12">Dec</option>
</select></div>
<div><label>Date</label><select id="ocDay"><option value="">-- Month first --</option></select></div>
<div style="grid-column:1/-1"><label>Staff Name</label><select id="ocMFO" onchange="autoFillBranchFromStaff('oc')"><option value="">-- Select Staff --</option></select></div>
<div style="grid-column:1/-1"><label>Name of Client</label><input type="text" id="ocClient" placeholder="Client name"></div>
<div style="grid-column:1/-1"><label>Remarks</label><select id="ocRemarks"><option value="">-- Select --</option>
<option value="Good Clients">Good Clients</option>
<option value="Subsequent Kamustahan">Subsequent Kamustahan</option>
<option value="Dummy">Dummy</option>
<option value="Loan Share">Loan Share</option>
<option value="Attitude">Attitude</option>
</select></div>
</div>
<div class="btn-row">
<button class="btn-sky" id="ocSaveBtn" onclick="saveOC()">💾 SAVE</button>
<button class="btn-update hidden" id="ocUpdateBtn" onclick="updateOC()">🔄 UPDATE</button>
<button class="btn-cancel hidden" id="ocCancelBtn" onclick="cancelEditOC()">✕ CANCEL</button>
<button class="btn-clear" onclick="clearOCForm()">🔄 Clear</button>
<button class="btn-export" onclick="exportOC()">📂 Export</button>
</div>
</div>
<div class="card">
<p class="card-title">📋 Offset Client Records</p>
<div class="kpi-grid" id="ocSummary"></div>
<div class="table-wrap"><table>
<thead><tr><th>#</th><th>Branch</th><th>Mo</th><th>Date</th><th>MFO</th><th>Client</th><th>Remarks</th><th class="text-center">Act</th></tr></thead>
<tbody id="ocBody"><tr><td colspan="8" class="text-center" style="color:var(--muted);padding:20px">No entries yet.</td></tr></tbody>
</table></div>
</div>
</div>

<!-- STAFF PERFORMANCE -->
<div class="tab-content" id="tab-staff">
<div class="card">
<p class="card-title">👥 STAFF PERFORMANCE — Monthly Consolidated</p>
<div class="filter-bar">
<label>📅 Month:</label>
<select id="staffMonth" onchange="updateStaffDays()">
<option value="">-- All Months --</option>
<option value="01">Jan</option><option value="02">Feb</option><option value="03">Mar</option><option value="04">Apr</option>
<option value="05">May</option><option value="06">Jun</option><option value="07">Jul</option><option value="08">Aug</option>
<option value="09">Sep</option><option value="10">Oct</option><option value="11">Nov</option><option value="12">Dec</option>
</select>
<label>📆 Date (Day):</label><select id="staffDay" onchange="renderStaffPerf()"><option value="">-- All Days --</option></select>
<button class="btn-export" onclick="exportStaffPerf()">📂 EXPORT</button>
</div>
<div class="kpi-grid" id="staffKPI"></div>
<div class="table-wrap"><table>
<thead><tr><th>Rank</th><th>Staff Name</th><th>Branch</th><th>Month</th><th>NM</th><th>Offset</th><th>Net NM</th><th>Bad Debt ₱</th><th>Delinquent</th><th>Kamustahan</th></tr></thead>
<tbody id="staffBody"></tbody>
</table></div>
</div>
</div>

<!-- INCENTIVES -->
<div class="tab-content" id="tab-incentives">
<div class="card">
<p class="card-title"><span>🎁 INCENTIVES — Entry</span><span class="badge" style="background:rgba(245,158,11,.15);color:#fcd34d">PERMANENT</span></p>
<div class="grid-form">
<div style="grid-column:1/-1"><label>Staff Name</label><select id="incStaff" onchange="autoFillBranchFromStaff('inc')"><option value="">-- Select Staff --</option></select></div>
<div><label>Branch Code</label><input type="text" id="incBranchCode" readonly style="background:rgba(0,0,0,.25)" placeholder="Auto from staff">
<option value="">-- Select --</option>
<option value="B0814">B0814</option><option value="B0392">B0392</option><option value="B0739">B0739</option>
<option value="B1659">B1659</option><option value="B1812">B1812</option>
</select></div>
<div><label>Branch Name</label><input type="text" id="incBranchName" readonly style="background:rgba(0,0,0,.25)" placeholder="Auto-filled"></div>
<div><label>ID Number</label><input type="text" id="incID" placeholder="ID No."></div>
<div><label>Last Name</label><input type="text" id="incLast" placeholder="Last Name"></div>
<div><label>First Name</label><input type="text" id="incFirst" placeholder="First Name"></div>
<div><label>Middle Name</label><input type="text" id="incMiddle" placeholder="Middle Name"></div>
<div><label>Designation</label><input type="text" id="incDesignation" placeholder="Designation"></div>
<div><label>Level</label><input type="text" id="incLevel" placeholder="e.g. Level 1"></div>
<div><label>No. of Borrowers</label><input type="text" id="incBorrowers" placeholder="e.g. 50"></div>
<div><label>Portfolio ₱</label><input type="number" id="incPortfolio" value="0" min="0" step="0.01"></div>
<div><label>PAR %</label><input type="number" id="incPAR" value="0" min="0" step="0.01"></div>
<div><label>Amount ₱</label><input type="text" id="incAmount" placeholder="e.g. 5000"></div>
<div style="grid-column:1/-1; border-top:1px dashed var(--border); padding-top:8px; margin-top:4px"><label style="color:var(--gold)">📈 INCREASE BORROWER — Incentive Details</label></div>
<div><label>Number of Increase Borrower</label><input type="text" id="incIncBorrowers" placeholder="e.g. 10"></div>
<div><label>Level for Increase of Borrowers</label><input type="text" id="incIncLevel" placeholder="e.g. Level 1"></div>
<div><label>Amount incentive of Increase Borrower</label><input type="text" id="incIncAmount" placeholder="e.g. 2000"></div>
</div>
<div class="btn-row">
<button class="btn-amber" id="incSaveBtn" onclick="saveINC()">💾 SAVE</button>
<button class="btn-update hidden" id="incUpdateBtn" onclick="updateINC()">🔄 UPDATE</button>
<button class="btn-cancel hidden" id="incCancelBtn" onclick="cancelEditINC()">✕ CANCEL</button>
<button class="btn-clear" onclick="clearINCForm()">🔄 Clear</button>
<button class="btn-export" onclick="exportINC()">📂 Export</button>
</div>
</div>
<div class="card">
<p class="card-title">📊 Level Summary</p>
<div class="kpi-grid" id="incKPI"></div>
<div class="table-wrap"><table>
<thead><tr><th>Level</th><th>No. of Staff</th><th>Borrowers</th><th>Amount ₱</th><th>% of Total</th></tr></thead>
<tbody id="incLevelBody"></tbody>
</table></div>
</div>
<div class="card">
<p class="card-title">📋 Incentives Records</p>
<div class="table-wrap"><table>
<thead><tr><th>#</th><th>Branch Code</th><th>Branch Name</th><th>ID</th><th>Name</th><th>Designation</th><th>Level</th><th>Borrowers</th><th>Portfolio ₱</th><th>PAR%</th><th>Amount ₱</th><th>Increase Borrowers</th><th>Increase Level</th><th>Increase Amount</th><th class="text-center">Act</th></tr></thead>
<tbody id="incBody"><tr><td colspan="16" class="text-center" style="color:var(--muted);padding:20px">No entries yet.</td></tr></tbody>
</table></div>
</div>
</div>

<!-- UPDATE TRACKER -->
<div class="tab-content" id="tab-update">
<div class="card">
<p class="card-title">✅ UPDATE TRACKER</p>
<div class="filter-bar">
<label>📅 Month:</label>
<select id="updateMonth" onchange="updateUpdateDays()">
<option value="">-- All Time --</option>
<option value="01">Jan</option><option value="02">Feb</option><option value="03">Mar</option><option value="04">Apr</option>
<option value="05">May</option><option value="06">Jun</option><option value="07">Jul</option><option value="08">Aug</option>
<option value="09">Sep</option><option value="10">Oct</option><option value="11">Nov</option><option value="12">Dec</option>
</select>
<label>📆 Date:</label><select id="updateDay" onchange="renderUpdateTracker()"><option value="">-- All Days --</option></select>
<button class="btn-view-on" id="viewStaffBtn" onclick="switchUpdateView('staff',this)">👥 STAFF STATUS</button>
<button class="btn-view-off" id="viewBranchBtn" onclick="switchUpdateView('branch',this)">🏢 BRANCH STATUS</button>
</div>
<div class="kpi-grid" id="updateKPI"></div>
<div id="staffView">
<div class="table-wrap"><table>
<thead><tr><th>Staff Name</th><th>Branch</th><th>NM & Off</th><th>Bad Debt</th><th>Delinquent</th><th>Kamustahan</th><th>Status</th></tr></thead>
<tbody id="updateStaffBody"></tbody>
</table></div>
</div>
<div id="branchView" class="hidden">
<div class="table-wrap"><table>
<thead><tr><th>Branch</th><th>Total Staff</th><th>Submitted</th><th>NM ✅</th><th>BD ✅</th><th>Del ✅</th><th>Kam ✅</th><th>Complete</th><th>%</th><th>Status</th></tr></thead>
<tbody id="updateBranchBody"></tbody>
</table></div>
</div>
</div>
</div>

<!-- SYNC & LOCK -->
<div class="tab-content" id="tab-sync">
<div class="card">
<p class="card-title">📅 LOCK / UNLOCK SPECIFIC DATE <span style="font-size:9px;color:var(--muted)">Requires password</span></p>
<div class="grid-form">
<div><label>Month</label><select id="dateLockMonth" onchange="updateDateLockDays()">
<option value="">-- Month --</option>
<option value="01">Jan</option><option value="02">Feb</option><option value="03">Mar</option><option value="04">Apr</option>
<option value="05">May</option><option value="06">Jun</option><option value="07">Jul</option><option value="08">Aug</option>
<option value="09">Sep</option><option value="10">Oct</option><option value="11">Nov</option><option value="12">Dec</option>
</select></div>
<div><label>Date (Day)</label><select id="dateLockDay"><option value="">-- Select Day --</option></select></div>
</div>
<div class="btn-row">
<button class="btn-lock" onclick="lockSpecificDate()">🔒 LOCK THIS DATE</button>
<button class="btn-unlock" onclick="unlockSpecificDate()">🔓 UNLOCK THIS DATE</button>
</div>
<div id="dateLockedList" style="margin-top:10px;font-size:11px"></div>
</div>

<div class="card">
<p class="card-title">📤 SHARE DASHBOARD</p>
<p class="sync-note">Generate a shareable sync code. Send it to a colleague — they paste it in the box below and click CONNECT to load all your data.</p>
<textarea id="shareLink" class="sync-box" readonly placeholder="Click GENERATE SYNC CODE below..."></textarea>
<div class="btn-row">
<button class="btn-teal" onclick="generateSyncCode()">🔑 GENERATE SYNC CODE</button>
<button class="btn-export" onclick="copySyncCode()">📋 COPY</button>
<button class="btn-amber" onclick="downloadShareableFile()">📥 DOWNLOAD SHAREABLE FILE</button>
</div>
</div>

<div class="card">
<p class="card-title">🔄 CONNECT / IMPORT</p>
<textarea id="syncIn" class="sync-box" placeholder="Paste sync code here, then click CONNECT..."></textarea>
<div class="btn-row">
<button class="btn-save" onclick="connectSync()">🔗 CONNECT & IMPORT</button>
</div>
<p class="sync-note">Sync merges records by ID — it will not delete your existing data.</p>
</div>
</div>

</div>

<script>
const BRANCHES = {'B0814':'Donsol','B0392':'Pilar 1','B0739':'Castilla','B1659':'Castilla 2','B1812':'Pilar 2'};
let STAFF = [];
const LOCK_PASSWORD = 'admin123';
const LS = 'sorsogon1_';
const MONTHS = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
let currentBranch = '';
let editing = {nm:null, bd:null, del:null, kam:null, oc:null, inc:null, exc:null};
let dateFilter = {month:'', day:''};

function store(k,v){ try{ localStorage.setItem(LS+k, JSON.stringify(v)); }catch(e){ alert('Storage full!'); } }
function load(k){ try{ return JSON.parse(localStorage.getItem(LS+k))||[]; }catch(e){ return []; } }
function uid(){ return Date.now().toString(36)+Math.random().toString(36).slice(2,7); }
function esc(s){ return String(s==null?'':s).replace(/[&<>"]/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }
function bName(c){ return BRANCHES[c]||(c?'Unknown':'All'); }
function num(v){ const n=parseFloat(v); return isNaN(n)?0:n; }
function daysIn(m){ if(!m) return 31; return new Date(2026, parseInt(m,10), 0).getDate(); }
function fillDays(id){
  const sel=document.getElementById(id); if(!sel) return;
  const pair=document.getElementById(id.replace('Day','Month')); const mv=pair?pair.value:'';
  const max=mv?daysIn(mv):31; const cur=sel.value;
  sel.innerHTML='<option value="">-- Select Day --</option>';
  for(let i=1;i<=max;i++){ const d=String(i).padStart(2,'0'); sel.innerHTML+=`<option value="${d}" ${cur===d?'selected':''}>${d}</option>`; }
}
function lockedDates(){ return load('lockedDates'); }
function isLocked(m,d){ return m&&d&&lockedDates().indexOf(m+'-'+d)!==-1; }
function checkLock(m,d){ if(isLocked(m,d)){ alert('🔒 This date ('+m+'-'+d+') is LOCKED. Editing not allowed.'); return true; } return false; }
function branchOfStaff(n){ const s=STAFF.find(x=>x.name===n); return s?s.branch:currentBranch; }
function byBranch(a){ return currentBranch? a.filter(r=>r.branch===currentBranch):a; }
function byDate(a){ return (!dateFilter.month&&!dateFilter.day)?a:a.filter(r=>(!dateFilter.month||r.month===dateFilter.month)&&(!dateFilter.day||r.day===dateFilter.day)); }

function toggleStaffDropdown(){ document.getElementById('staffDropdown').classList.toggle('show'); document.getElementById('clearDropdown').classList.remove('show'); renderStaffList(); }
function renderStaffList(){
  const c=document.getElementById('staffListContainer'); c.innerHTML='';
  const opts=Object.keys(BRANCHES).map(b=>`<option value="${b}">${b} — ${BRANCHES[b]}</option>`).join('');
  STAFF.forEach(s=>{ c.innerHTML+=`<div class="staff-row-item"><input type="text" class="staff-name-inp" value="${esc(s.name)}" placeholder="Staff name"><select class="staff-branch-sel">${Object.keys(BRANCHES).map(b=>`<option value="${b}" ${s.branch===b?'selected':''}>${b} — ${BRANCHES[b]}</option>`).join('')}</select><button class="btn-del btn-sm" onclick="this.parentElement.remove()">✕</button></div>`; });
}
function addStaffRow(){
  const c=document.getElementById('staffListContainer'); const div=document.createElement('div'); div.className='staff-row-item';
  div.innerHTML=`<input type="text" class="staff-name-inp" placeholder="Staff name"><select class="staff-branch-sel">${Object.keys(BRANCHES).map(b=>`<option value="${b}">${b} — ${BRANCHES[b]}</option>`).join('')}</select><button class="btn-del btn-sm" onclick="this.parentElement.remove()">✕</button>`;
  c.appendChild(div);
}
function saveStaffList(){
  const rows=document.querySelectorAll('#staffListContainer .staff-row-item'); const list=[];
  rows.forEach(r=>{ const n=r.querySelector('.staff-name-inp').value.trim(); if(n) list.push({name:n, branch:r.querySelector('.staff-branch-sel').value}); });
  STAFF=list; store('staff',list); populateStaffDropdowns(); alert('✅ Staff list saved! All dropdowns updated.');
}
function resetStaffList(){ renderStaffList(); }
function populateStaffDropdowns(){
  const opts='<option value="">-- Select Staff --</option>'+STAFF.map(s=>`<option value="${esc(s.name)}">${esc(s.name)} (${bName(s.branch)})</option>`).join('');
  ['nmStaff','bdStaff','delStaff','kamStaff','ocMFO','excStaff','incStaff'].forEach(id=>{ const el=document.getElementById(id); if(el) el.innerHTML=opts; });
}
function autoFillBranchFromStaff(prefix){
  const el=document.getElementById(prefix==='oc'?'ocMFO':prefix+'Staff'); const name=el?el.value:'';
  const s=STAFF.find(x=>x.name===name); const b=s?s.branch:currentBranch;
  const lbl=document.getElementById(prefix+'BranchLabel'); if(lbl) lbl.textContent=b?('🏢 Branch: '+bName(b)+' ('+b+')'):'🏢 Branch: —';
  if(prefix==='inc'){ document.getElementById('incBranchCode').value=b||''; document.getElementById('incBranchName').value=b?bName(b):''; }
}
function init(){
  STAFF = load('staff');
  populateStaffDropdowns();
  addDelClientRow();
  renderAll();
  renderLockedList();
}

function switchTab(name,btn){
  document.querySelectorAll('.tab-content').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('tab-'+name).classList.add('active');
  btn.classList.add('active');
  renderAll();
}
function selectBranch(code,btn){
  currentBranch=code;
  document.querySelectorAll('.branch-btn-top').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('curBranchLabel').textContent=code?(bName(code)+' — '+code):'Consolidated — All Branches';
  ['nm','bd','del','kam','oc','exc'].forEach(p=>{ const el=document.getElementById(p+'BranchLabel'); if(el) el.textContent=code?('🏢 Branch: '+bName(code)+' ('+code+')'):'🏢 Branch: Auto-detected from Staff / MFO'; });
  renderAll();
}
function updateDays(p){ fillDays(p+'Day'); }
function updateDayOptions(){ fillDays('filterDay'); }
function updateStaffDays(){ fillDays('staffDay'); renderStaffPerf(); }
function updateUpdateDays(){ fillDays('updateDay'); renderUpdateTracker(); }
function updateDateLockDays(){ fillDays('dateLockDay'); }
function updateClearDays(){ fillDays('clearDaySelect'); }

function toggleClearDropdown(){ document.getElementById('clearDropdown').classList.toggle('show'); }
document.addEventListener('click', function(e){ if(!e.target.closest('.clear-btn-group')) document.getElementById('clearDropdown').classList.remove('show'); if(!e.target.closest('.staff-btn-wrap')) document.getElementById('staffDropdown').classList.remove('show'); });
function clearDataByMonth(){
  const m=document.getElementById('clearMonthSelect').value, d=document.getElementById('clearDaySelect').value;
  if(!m){ alert('Select a month first.'); return; }
  if(!confirm('Clear ALL records for '+MONTHS[parseInt(m)]+(d?(' '+d):'')+'? Cannot be undone.')) return;
  ['nm','bd','del','kam','oc','exc'].forEach(k=>{ store(k, load(k).filter(r=>!(r.month===m&&(!d||r.day===d)))); });
  alert('✅ Cleared.'); document.getElementById('clearDropdown').classList.remove('show'); renderAll();
}
function clearAllData(){
  if(!confirm('⚠️ CLEAR ALL DATA? Cannot be undone!')) return;
  if(!confirm('Really clear EVERYTHING?')) return;
  ['nm','bd','del','kam','oc','inc','exc','lockedDates'].forEach(k=>localStorage.removeItem(LS+k));
  alert('✅ All data cleared.'); document.getElementById('clearDropdown').classList.remove('show'); renderAll();
}

function applyDateFilter(){ dateFilter.month=document.getElementById('filterMonth').value; dateFilter.day=document.getElementById('filterDay').value; renderOverview(); }
function resetDateFilter(){ dateFilter={month:'',day:''}; document.getElementById('filterMonth').value=''; document.getElementById('filterDay').innerHTML='<option value="">-- All --</option>'; renderOverview(); }

function renderOverview(){
  const nm=byDate(byBranch(load('nm'))), bd=byDate(byBranch(load('bd'))), del=byDate(byBranch(load('del'))), kam=byDate(byBranch(load('kam'))), oc=byDate(byBranch(load('oc'))), exc=byDate(byBranch(load('exc')));
  document.getElementById('ovTitle').textContent='📊 OVERVIEW — '+(currentBranch?bName(currentBranch):'All Branches')+(dateFilter.month?(' · '+MONTHS[parseInt(dateFilter.month)]+(dateFilter.day?(' '+dateFilter.day):'')):'');
  const b={}; Object.keys(BRANCHES).forEach(c=>b[c]={nm:0,off:0,net:0,bd:0,bdt:0,del:0,kam:0,oc:0,exc:0});
  nm.forEach(r=>{ if(!b[r.branch])b[r.branch]={nm:0,off:0,net:0,bd:0,bdt:0,del:0,kam:0,oc:0,exc:0}; b[r.branch].nm+=num(r.nm); b[r.branch].off+=num(r.off); b[r.branch].net+=num(r.nm)-num(r.off); });
  bd.forEach(r=>{ if(!b[r.branch])b[r.branch]={nm:0,off:0,net:0,bd:0,bdt:0,del:0,kam:0,oc:0,exc:0}; b[r.branch].bd+=num(r.amount); b[r.branch].bdt+=num(r.target); });
  del.forEach(r=>{ if(!b[r.branch])b[r.branch]={nm:0,off:0,net:0,bd:0,bdt:0,del:0,kam:0,oc:0,exc:0}; b[r.branch].del+=(r.clients?r.clients.length:1); });
  kam.forEach(r=>{ if(!b[r.branch])b[r.branch]={nm:0,off:0,net:0,bd:0,bdt:0,del:0,kam:0,oc:0,exc:0}; b[r.branch].kam+=num(r.count); });
  oc.forEach(r=>{ if(!b[r.branch])b[r.branch]={nm:0,off:0,net:0,bd:0,bdt:0,del:0,kam:0,oc:0,exc:0}; b[r.branch].oc+=1; });
  exc.forEach(r=>{ if(!b[r.branch])b[r.branch]={nm:0,off:0,net:0,bd:0,bdt:0,del:0,kam:0,oc:0,exc:0}; b[r.branch].exc+=1; });
  let t={nm:0,off:0,net:0,bd:0,bdt:0,del:0,kam:0,oc:0,exc:0}; let rows='';
  Object.keys(b).forEach(c=>{ const x=b[c]; const pct=x.bdt>0?((x.bd/x.bdt)*100).toFixed(1):'0';
    rows+=`<tr><td><strong>${esc(bName(c))}</strong><br><span style="color:var(--muted);font-size:8px">${c}</span></td><td>${x.nm}</td><td>${x.off}</td><td style="color:var(--green);font-weight:700">${x.net}</td><td>₱${x.bd.toLocaleString()}</td><td>₱${x.bdt.toLocaleString()}</td><td>${pct}%</td><td>${x.del}</td><td>${x.kam}</td><td>${x.oc}</td><td>${x.exc}</td></tr>`;
    ['nm','off','net','bd','bdt','del','kam','oc','exc'].forEach(k=>t[k]+=x[k]); });
  const tp=t.bdt>0?((t.bd/t.bdt)*100).toFixed(1):'0';
  rows+=`<tr class="total-row"><td>TOTAL</td><td>${t.nm}</td><td>${t.off}</td><td>${t.net}</td><td>₱${t.bd.toLocaleString()}</td><td>₱${t.bdt.toLocaleString()}</td><td>${tp}%</td><td>${t.del}</td><td>${t.kam}</td><td>${t.oc}</td><td>${t.exc}</td></tr>`;
  document.getElementById('overviewBody').innerHTML=rows;
  document.getElementById('overviewKPI').innerHTML=`
    <div class="kpi-card"><div class="kl">Total NM</div><div class="kv kg">${t.nm}</div></div>
    <div class="kpi-card"><div class="kl">Total Offset</div><div class="kv ko">${t.off}</div></div>
    <div class="kpi-card"><div class="kl">Net NM</div><div class="kv kb">${t.net}</div></div>
    <div class="kpi-card"><div class="kl">Bad Debt Coll.</div><div class="kv kg">₱${t.bd.toLocaleString()}</div></div>
    <div class="kpi-card"><div class="kl">Delinquent</div><div class="kv kr">${t.del}</div></div>
    <div class="kpi-card"><div class="kl">Kamustahan</div><div class="kv kp">${t.kam}</div></div>
    <div class="kpi-card"><div class="kl">Offset Clients</div><div class="kv ks">${t.oc}</div></div>
    <div class="kpi-card"><div class="kl">Excuses</div><div class="kv kexc">${t.exc}</div></div>`;
}

/* NM */
function calcNetNM(){ document.getElementById('netNM').value=num(document.getElementById('nmValue').value)-num(document.getElementById('offsetValue').value); }
function clearNMForm(){ ['nmStaff','nmMonth','nmDay'].forEach(id=>document.getElementById(id).value=''); ['nmValue','offsetValue','targetNM','targetOffset'].forEach(id=>document.getElementById(id).value='0'); calcNetNM(); editing.nm=null; document.getElementById('nmSaveBtn').classList.remove('hidden'); document.getElementById('nmUpdateBtn').classList.add('hidden'); document.getElementById('nmCancelBtn').classList.add('hidden'); }
function saveNM(){
  const staff=document.getElementById('nmStaff').value, m=document.getElementById('nmMonth').value, d=document.getElementById('nmDay').value;
  if(!staff||!m||!d){ alert('Select Staff, Month and Date.'); return; }
  if(checkLock(m,d)) return;
  const a=load('nm'); a.push({id:uid(), branch:branchOfStaff(staff), staff, month:m, day:d, nm:num(document.getElementById('nmValue').value), off:num(document.getElementById('offsetValue').value), tnm:num(document.getElementById('targetNM').value), toff:num(document.getElementById('targetOffset').value)});
  store('nm',a); clearNMForm(); renderNM(); renderAll();
}
function editNM(id){ const r=load('nm').find(x=>x.id===id); if(!r)return; if(checkLock(r.month,r.day))return; editing.nm=id;
  document.getElementById('nmStaff').value=r.staff; document.getElementById('nmMonth').value=r.month; fillDays('nmDay'); document.getElementById('nmDay').value=r.day;
  document.getElementById('nmValue').value=r.nm; document.getElementById('offsetValue').value=r.off; document.getElementById('targetNM').value=r.tnm; document.getElementById('targetOffset').value=r.toff; calcNetNM();
  document.getElementById('nmSaveBtn').classList.add('hidden'); document.getElementById('nmUpdateBtn').classList.remove('hidden'); document.getElementById('nmCancelBtn').classList.remove('hidden'); document.getElementById('nmFormCard').scrollIntoView({behavior:'smooth'}); }
function updateNM(){ const m=document.getElementById('nmMonth').value, d=document.getElementById('nmDay').value; if(checkLock(m,d))return; const staff=document.getElementById('nmStaff').value;
  const a=load('nm'); const i=a.findIndex(x=>x.id===editing.nm); if(i===-1)return;
  a[i]={...a[i], branch:branchOfStaff(staff), staff, month:m, day:d, nm:num(document.getElementById('nmValue').value), off:num(document.getElementById('offsetValue').value), tnm:num(document.getElementById('targetNM').value), toff:num(document.getElementById('targetOffset').value)};
  store('nm',a); clearNMForm(); renderNM(); renderAll(); }
function cancelEditNM(){ clearNMForm(); }
function deleteNM(id){ if(!confirm('Delete this NM record?'))return; store('nm', load('nm').filter(x=>x.id!==id)); renderNM(); renderAll(); }
function renderNM(){
  const a=byBranch(load('nm')).sort((x,y)=>(x.month+x.day).localeCompare(y.month+y.day));
  const b=document.getElementById('nmBody');
  if(!a.length){ b.innerHTML='<tr><td colspan="11" class="text-center" style="color:var(--muted);padding:20px">No entries yet.</td></tr>'; return; }
  b.innerHTML=a.map((r,i)=>`<tr class="${isLocked(r.month,r.day)?'locked-row':''}"><td>${i+1}</td><td>${esc(bName(r.branch))}</td><td>${esc(r.staff)}</td><td>${MONTHS[parseInt(r.month)]}</td><td>${r.day}${isLocked(r.month,r.day)?' <span class="lock-tag">LOCKED</span>':''}</td><td>${r.nm}</td><td>${r.off}</td><td style="color:var(--green);font-weight:700">${r.nm-r.off}</td><td>${r.tnm}</td><td>${r.toff}</td><td class="text-center"><div class="action-btns"><button class="btn-edit" onclick="editNM('${r.id}')">Edit</button><button class="btn-del" onclick="deleteNM('${r.id}')">Del</button></div></td></tr>`).join('');
}
function exportNM(){ exportXLSX(byBranch(load('nm')), 'NM_Offset', ['Branch','Staff','Month','Day','NM','Offset','Net NM','Target NM','Target Offset'], r=>[bName(r.branch),r.staff,r.month,r.day,r.nm,r.off,r.nm-r.off,r.tnm,r.toff]); }

/* BAD DEBT */
function calcBalanceBD(){ document.getElementById('bdBalance').value=(num(document.getElementById('bdTarget').value)-num(document.getElementById('bdAmount').value)).toFixed(2); }
function clearBDForm(){ ['bdStaff','bdMonth','bdDay','bdClient','bdWorkingDay'].forEach(id=>document.getElementById(id).value=''); ['bdAmount','bdTarget'].forEach(id=>document.getElementById(id).value='0'); calcBalanceBD(); editing.bd=null; document.getElementById('bdSaveBtn').classList.remove('hidden'); document.getElementById('bdUpdateBtn').classList.add('hidden'); document.getElementById('bdCancelBtn').classList.add('hidden'); }
function saveBD(){
  const staff=document.getElementById('bdStaff').value.trim(), m=document.getElementById('bdMonth').value, d=document.getElementById('bdDay').value;
  if(!staff||!m||!d){ alert('Fill Staff, Month and Date.'); return; }
  if(checkLock(m,d)) return;
  const a=load('bd'); a.push({id:uid(), branch:currentBranch||branchOfStaff(staff)||'B0814', staff, month:m, day:d, client:document.getElementById('bdClient').value.trim(), amount:num(document.getElementById('bdAmount').value), week:document.getElementById('bdWorkingDay').value, target:num(document.getElementById('bdTarget').value)});
  store('bd',a); clearBDForm(); renderBD(); renderAll();
}
function editBD(id){ const r=load('bd').find(x=>x.id===id); if(!r)return; if(checkLock(r.month,r.day))return; editing.bd=id;
  document.getElementById('bdStaff').value=r.staff; document.getElementById('bdMonth').value=r.month; fillDays('bdDay'); document.getElementById('bdDay').value=r.day;
  document.getElementById('bdClient').value=r.client||''; document.getElementById('bdAmount').value=r.amount; document.getElementById('bdWorkingDay').value=r.week||''; document.getElementById('bdTarget').value=r.target; calcBalanceBD();
  document.getElementById('bdSaveBtn').classList.add('hidden'); document.getElementById('bdUpdateBtn').classList.remove('hidden'); document.getElementById('bdCancelBtn').classList.remove('hidden'); document.getElementById('bdFormCard').scrollIntoView({behavior:'smooth'}); }
function updateBD(){ const m=document.getElementById('bdMonth').value, d=document.getElementById('bdDay').value; if(checkLock(m,d))return;
  const a=load('bd'); const i=a.findIndex(x=>x.id===editing.bd); if(i===-1)return;
  a[i]={...a[i], staff:document.getElementById('bdStaff').value.trim(), month:m, day:d, client:document.getElementById('bdClient').value.trim(), amount:num(document.getElementById('bdAmount').value), week:document.getElementById('bdWorkingDay').value, target:num(document.getElementById('bdTarget').value)};
  store('bd',a); clearBDForm(); renderBD(); renderAll(); }
function cancelEditBD(){ clearBDForm(); }
function deleteBD(id){ if(!confirm('Delete this Bad Debt record?'))return; store('bd', load('bd').filter(x=>x.id!==id)); renderBD(); renderAll(); }
function renderBD(){
  const wf=document.getElementById('bdWeekFilter').value;
  let a=byBranch(load('bd')); if(wf)a=a.filter(r=>r.week===wf); a=a.sort((x,y)=>(x.month+x.day).localeCompare(y.month+y.day));
  document.getElementById('bdWeekLabel').textContent=wf?('— '+wf):'';
  const ta=a.reduce((s,r)=>s+num(r.amount),0), tt=a.reduce((s,r)=>s+num(r.target),0);
  document.getElementById('bdSummary').innerHTML=`<div class="kpi-card"><div class="kl">Entries</div><div class="kv kb">${a.length}</div></div><div class="kpi-card"><div class="kl">Collected</div><div class="kv kg">₱${ta.toLocaleString()}</div></div><div class="kpi-card"><div class="kl">Target</div><div class="kv ko">₱${tt.toLocaleString()}</div></div><div class="kpi-card"><div class="kl">Remaining</div><div class="kv kr">₱${(tt-ta).toLocaleString()}</div></div>`;
  const b=document.getElementById('bdBody');
  if(!a.length){ b.innerHTML='<tr><td colspan="10" class="text-center" style="color:var(--muted);padding:20px">No entries yet.</td></tr>'; return; }
  b.innerHTML=a.map((r,i)=>`<tr class="${isLocked(r.month,r.day)?'locked-row':''}"><td>${i+1}</td><td>${esc(bName(r.branch))}</td><td>${esc(r.staff)}</td><td>${MONTHS[parseInt(r.month)]}</td><td>${r.day}${isLocked(r.month,r.day)?' <span class="lock-tag">LOCKED</span>':''}</td><td>${esc(r.week||'-')}</td><td>${esc(r.client||'-')}</td><td>₱${num(r.amount).toLocaleString()}</td><td>₱${num(r.target).toLocaleString()}</td><td class="text-center"><div class="action-btns"><button class="btn-edit" onclick="editBD('${r.id}')">Edit</button><button class="btn-del" onclick="deleteBD('${r.id}')">Del</button></div></td></tr>`).join('');
}
function exportBD(){ exportXLSX(byBranch(load('bd')), 'Bad_Debt', ['Branch','Staff','Month','Day','Week','Client','Amount','Target'], r=>[bName(r.branch),r.staff,r.month,r.day,r.week,r.client,r.amount,r.target]); }

/* DELINQUENT */
function addDelClientRow(n,rs){
  const c=document.getElementById('delClientsContainer'); const row=document.createElement('div'); row.className='client-pair';
  row.innerHTML=`<input type="text" placeholder="Client name" class="del-cn" value="${esc(n||'')}"><input type="text" placeholder="Reason" class="del-cr" value="${esc(rs||'')}"><button type="button" class="remove-client-btn" onclick="this.parentElement.remove()">✕</button>`;
  c.appendChild(row);
}
function clearDelForm(){ ['delMonth','delDay','delStaff'].forEach(id=>document.getElementById(id).value=''); document.getElementById('delClientsContainer').innerHTML=''; addDelClientRow(); editing.del=null; document.getElementById('delSaveBtn').classList.remove('hidden'); document.getElementById('delUpdateBtn').classList.add('hidden'); document.getElementById('delCancelBtn').classList.add('hidden'); }
function readDelClients(){ return Array.from(document.querySelectorAll('#delClientsContainer .client-pair')).map(row=>{ const n=row.querySelector('.del-cn').value.trim(), rs=row.querySelector('.del-cr').value.trim(); return n?{name:n,reason:rs}:null; }).filter(Boolean); }
function saveDel(){
  const m=document.getElementById('delMonth').value, d=document.getElementById('delDay').value, staff=document.getElementById('delStaff').value.trim();
  if(!m||!d||!staff){ alert('Fill Month, Date and Staff.'); return; }
  if(checkLock(m,d)) return;
  const cl=readDelClients(); if(!cl.length){ alert('Add at least one client.'); return; }
  const a=load('del'); a.push({id:uid(), branch:currentBranch||branchOfStaff(staff)||'B0814', month:m, day:d, staff, clients:cl});
  store('del',a); clearDelForm(); renderDel(); renderAll();
}
function editDel(id){ const r=load('del').find(x=>x.id===id); if(!r)return; if(checkLock(r.month,r.day))return; editing.del=id;
  document.getElementById('delMonth').value=r.month; fillDays('delDay'); document.getElementById('delDay').value=r.day; document.getElementById('delStaff').value=r.staff;
  document.getElementById('delClientsContainer').innerHTML=''; (r.clients||[]).forEach(c=>addDelClientRow(c.name,c.reason));
  document.getElementById('delSaveBtn').classList.add('hidden'); document.getElementById('delUpdateBtn').classList.remove('hidden'); document.getElementById('delCancelBtn').classList.remove('hidden'); }
function updateDel(){ const m=document.getElementById('delMonth').value, d=document.getElementById('delDay').value; if(checkLock(m,d))return;
  const a=load('del'); const i=a.findIndex(x=>x.id===editing.del); if(i===-1)return;
  a[i]={...a[i], month:m, day:d, staff:document.getElementById('delStaff').value.trim(), clients:readDelClients()};
  store('del',a); clearDelForm(); renderDel(); renderAll(); }
function cancelEditDel(){ clearDelForm(); }
function deleteDel(id){ if(!confirm('Delete this Delinquent record?'))return; store('del', load('del').filter(x=>x.id!==id)); renderDel(); renderAll(); }
function renderDel(){
  const a=byBranch(load('del')).sort((x,y)=>(x.month+x.day).localeCompare(y.month+y.day));
  const tc=a.reduce((s,r)=>s+(r.clients?r.clients.length:0),0);
  document.getElementById('delSummary').innerHTML=`<div class="kpi-card"><div class="kl">Records</div><div class="kv kb">${a.length}</div></div><div class="kpi-card"><div class="kl">Delinquent Clients</div><div class="kv kr">${tc}</div></div>`;
  const b=document.getElementById('delBody');
  if(!a.length){ b.innerHTML='<tr><td colspan="8" class="text-center" style="color:var(--muted);padding:20px">No entries yet.</td></tr>'; return; }
  b.innerHTML=a.map((r,i)=>{ const cr=(r.clients||[]).map(c=>`<div>• <strong>${esc(c.name)}</strong> — <span style="color:var(--muted)">${esc(c.reason||'-')}</span></div>`).join('');
    return `<tr class="${isLocked(r.month,r.day)?'locked-row':''}"><td>${i+1}</td><td>${esc(bName(r.branch))}</td><td>${MONTHS[parseInt(r.month)]}</td><td>${r.day}${isLocked(r.month,r.day)?' <span class="lock-tag">LOCKED</span>':''}</td><td>${esc(r.staff)}</td><td>${cr||'-'}</td><td class="text-center"><div class="action-btns"><button class="btn-edit" onclick="editDel('${r.id}')">Edit</button><button class="btn-del" onclick="deleteDel('${r.id}')">Del</button></div></td></tr>`; }).join('');
}
function exportDel(){ const flat=[]; byBranch(load('del')).forEach(r=>(r.clients||[]).forEach(c=>flat.push([bName(r.branch),r.month,r.day,r.staff,c.name,c.reason]))); exportXLSX(flat, 'Delinquent', ['Branch','Month','Day','Staff','Client','Reason'], x=>x); }

/* EXCUSES (no filter) */
function clearEXCForm(){ ['excMonth','excDay','excStaff','excClient','excReason'].forEach(id=>{ const el=document.getElementById(id); if(el)el.value=''; }); editing.exc=null; document.getElementById('excSaveBtn').classList.remove('hidden'); document.getElementById('excUpdateBtn').classList.add('hidden'); document.getElementById('excCancelBtn').classList.add('hidden'); }
function saveEXC(){
  const m=document.getElementById('excMonth').value, d=document.getElementById('excDay').value, staff=document.getElementById('excStaff').value, reason=document.getElementById('excReason').value;
  if(!m||!d||!staff||!reason){ alert('Fill Month, Date, Staff and Reason.'); return; }
  if(checkLock(m,d)) return;
  const a=load('exc'); a.push({id:uid(), branch:branchOfStaff(staff)||currentBranch||'B0814', month:m, day:d, staff, client:document.getElementById('excClient').value.trim(), reason});
  store('exc',a); clearEXCForm(); renderEXC(); renderAll();
}
function editEXC(id){ const r=load('exc').find(x=>x.id===id); if(!r)return; if(checkLock(r.month,r.day))return; editing.exc=id;
  document.getElementById('excMonth').value=r.month; fillDays('excDay'); document.getElementById('excDay').value=r.day;
  document.getElementById('excStaff').value=r.staff||''; document.getElementById('excClient').value=r.client||''; document.getElementById('excReason').value=r.reason||'';
  document.getElementById('excSaveBtn').classList.add('hidden'); document.getElementById('excUpdateBtn').classList.remove('hidden'); document.getElementById('excCancelBtn').classList.remove('hidden'); }
function updateEXC(){ const m=document.getElementById('excMonth').value, d=document.getElementById('excDay').value; if(checkLock(m,d))return;
  const a=load('exc'); const i=a.findIndex(x=>x.id===editing.exc); if(i===-1)return;
  const staff=document.getElementById('excStaff').value; a[i]={...a[i], branch:branchOfStaff(staff)||currentBranch||'B0814', month:m, day:d, staff, client:document.getElementById('excClient').value.trim(), reason:document.getElementById('excReason').value};
  store('exc',a); clearEXCForm(); renderEXC(); renderAll(); }
function cancelEditEXC(){ clearEXCForm(); }
function deleteEXC(id){ if(!confirm('Delete this Excuse record?'))return; store('exc', load('exc').filter(x=>x.id!==id)); renderEXC(); renderAll(); }
function renderEXC(){
  const a=byBranch(load('exc')).sort((x,y)=>(x.month+x.day).localeCompare(y.month+y.day));
  document.getElementById('excTotalCount').textContent=a.length;
  const b=document.getElementById('excBody');
  if(!a.length){ b.innerHTML='<tr><td colspan="8" class="text-center" style="color:var(--muted);padding:20px">No entries yet.</td></tr>'; return; }
  b.innerHTML=a.map((r,i)=>`<tr class="${isLocked(r.month,r.day)?'locked-row':''}"><td>${i+1}</td><td>${esc(bName(r.branch))}</td><td>${MONTHS[parseInt(r.month)]}</td><td>${r.day}${isLocked(r.month,r.day)?' <span class="lock-tag">LOCKED</span>':''}</td><td>${esc(r.staff||'-')}</td><td>${esc(r.client||'-')}</td><td><strong>${esc(r.reason||'-')}</strong></td><td class="text-center"><div class="action-btns"><button class="btn-edit" onclick="editEXC('${r.id}')">Edit</button><button class="btn-del" onclick="deleteEXC('${r.id}')">Del</button></div></td></tr>`).join('');
}
function exportEXC(){ exportXLSX(byBranch(load('exc')), 'Excuses', ['Branch','Month','Day','Staff','Client','Reason'], r=>[bName(r.branch),r.month,r.day,r.staff||'',r.client,r.reason]); }

/* KAMUSTAHAN */
function clearKamForm(){ ['kamMonth','kamDay','kamStaff','kamRemarks'].forEach(id=>document.getElementById(id).value=''); document.getElementById('kamCount').value='0'; editing.kam=null; document.getElementById('kamSaveBtn').classList.remove('hidden'); document.getElementById('kamUpdateBtn').classList.add('hidden'); document.getElementById('kamCancelBtn').classList.add('hidden'); }
function saveKam(){ const m=document.getElementById('kamMonth').value, d=document.getElementById('kamDay').value, staff=document.getElementById('kamStaff').value.trim(); if(!m||!d||!staff){alert('Fill Month, Date and Staff.');return;} if(checkLock(m,d))return;
  const a=load('kam'); a.push({id:uid(), branch:currentBranch||branchOfStaff(staff)||'B0814', month:m, day:d, staff, count:num(document.getElementById('kamCount').value), remarks:document.getElementById('kamRemarks').value.trim()}); store('kam',a); clearKamForm(); renderKam(); renderAll(); }
function editKam(id){ const r=load('kam').find(x=>x.id===id); if(!r)return; if(checkLock(r.month,r.day))return; editing.kam=id;
  document.getElementById('kamMonth').value=r.month; fillDays('kamDay'); document.getElementById('kamDay').value=r.day; document.getElementById('kamStaff').value=r.staff; document.getElementById('kamCount').value=r.count; document.getElementById('kamRemarks').value=r.remarks||'';
  document.getElementById('kamSaveBtn').classList.add('hidden'); document.getElementById('kamUpdateBtn').classList.remove('hidden'); document.getElementById('kamCancelBtn').classList.remove('hidden'); }
function updateKam(){ const m=document.getElementById('kamMonth').value, d=document.getElementById('kamDay').value; if(checkLock(m,d))return;
  const a=load('kam'); const i=a.findIndex(x=>x.id===editing.kam); if(i===-1)return;
  a[i]={...a[i], month:m, day:d, staff:document.getElementById('kamStaff').value.trim(), count:num(document.getElementById('kamCount').value), remarks:document.getElementById('kamRemarks').value.trim()};
  store('kam',a); clearKamForm(); renderKam(); renderAll(); }
function cancelEditKam(){ clearKamForm(); }
function deleteKam(id){ if(!confirm('Delete this Kamustahan record?'))return; store('kam', load('kam').filter(x=>x.id!==id)); renderKam(); renderAll(); }
function renderKam(){
  const a=byBranch(load('kam')).sort((x,y)=>(x.month+x.day).localeCompare(y.month+y.day)); const tot=a.reduce((s,r)=>s+num(r.count),0);
  document.getElementById('kamSummary').innerHTML=`<div class="kpi-card"><div class="kl">Records</div><div class="kv kb">${a.length}</div></div><div class="kpi-card"><div class="kl">Total Kamustahan</div><div class="kv kp">${tot}</div></div>`;
  const b=document.getElementById('kamBody');
  if(!a.length){ b.innerHTML='<tr><td colspan="8" class="text-center" style="color:var(--muted);padding:20px">No entries yet.</td></tr>'; return; }
  b.innerHTML=a.map((r,i)=>`<tr class="${isLocked(r.month,r.day)?'locked-row':''}"><td>${i+1}</td><td>${esc(bName(r.branch))}</td><td>${MONTHS[parseInt(r.month)]}</td><td>${r.day}${isLocked(r.month,r.day)?' <span class="lock-tag">LOCKED</span>':''}</td><td>${esc(r.staff)}</td><td>${r.count}</td><td>${esc(r.remarks||'-')}</td><td class="text-center"><div class="action-btns"><button class="btn-edit" onclick="editKam('${r.id}')">Edit</button><button class="btn-del" onclick="deleteKam('${r.id}')">Del</button></div></td></tr>`).join('');
}
function exportKam(){ exportXLSX(byBranch(load('kam')), 'Kamustahan', ['Branch','Month','Day','Staff','Count','Remarks'], r=>[bName(r.branch),r.month,r.day,r.staff,r.count,r.remarks]); }

/* OFFSET CLIENT */
function clearOCForm(){ ['ocMonth','ocDay','ocMFO','ocClient','ocRemarks'].forEach(id=>document.getElementById(id).value=''); editing.oc=null; document.getElementById('ocSaveBtn').classList.remove('hidden'); document.getElementById('ocUpdateBtn').classList.add('hidden'); document.getElementById('ocCancelBtn').classList.add('hidden'); }
function saveOC(){ const m=document.getElementById('ocMonth').value, d=document.getElementById('ocDay').value, mfo=document.getElementById('ocMFO').value, client=document.getElementById('ocClient').value.trim(); if(!m||!d||!mfo||!client){alert('Fill Month, Date, MFO and Client.');return;} if(checkLock(m,d))return;
  const a=load('oc'); a.push({id:uid(), branch:branchOfStaff(mfo)||currentBranch||'B0814', month:m, day:d, mfo, client, remarks:document.getElementById('ocRemarks').value}); store('oc',a); clearOCForm(); renderOC(); renderAll(); }
function editOC(id){ const r=load('oc').find(x=>x.id===id); if(!r)return; if(checkLock(r.month,r.day))return; editing.oc=id;
  document.getElementById('ocMonth').value=r.month; fillDays('ocDay'); document.getElementById('ocDay').value=r.day; document.getElementById('ocMFO').value=r.mfo; document.getElementById('ocClient').value=r.client; document.getElementById('ocRemarks').value=r.remarks||'';
  document.getElementById('ocSaveBtn').classList.add('hidden'); document.getElementById('ocUpdateBtn').classList.remove('hidden'); document.getElementById('ocCancelBtn').classList.remove('hidden'); }
function updateOC(){ const m=document.getElementById('ocMonth').value, d=document.getElementById('ocDay').value; if(checkLock(m,d))return;
  const a=load('oc'); const i=a.findIndex(x=>x.id===editing.oc); if(i===-1)return; const mfo=document.getElementById('ocMFO').value;
  a[i]={...a[i], branch:branchOfStaff(mfo)||currentBranch||'B0814', month:m, day:d, mfo, client:document.getElementById('ocClient').value.trim(), remarks:document.getElementById('ocRemarks').value};
  store('oc',a); clearOCForm(); renderOC(); renderAll(); }
function cancelEditOC(){ clearOCForm(); }
function deleteOC(id){ if(!confirm('Delete this Offset Client record?'))return; store('oc', load('oc').filter(x=>x.id!==id)); renderOC(); renderAll(); }
function renderOC(){
  const a=byBranch(load('oc')).sort((x,y)=>(x.month+x.day).localeCompare(y.month+y.day));
  document.getElementById('ocSummary').innerHTML=`<div class="kpi-card"><div class="kl">Total Offset Clients</div><div class="kv ks">${a.length}</div></div>`;
  const b=document.getElementById('ocBody');
  if(!a.length){ b.innerHTML='<tr><td colspan="8" class="text-center" style="color:var(--muted);padding:20px">No entries yet.</td></tr>'; return; }
  b.innerHTML=a.map((r,i)=>`<tr class="${isLocked(r.month,r.day)?'locked-row':''}"><td>${i+1}</td><td>${esc(bName(r.branch))}</td><td>${MONTHS[parseInt(r.month)]}</td><td>${r.day}${isLocked(r.month,r.day)?' <span class="lock-tag">LOCKED</span>':''}</td><td>${esc(r.mfo)}</td><td>${esc(r.client)}</td><td>${esc(r.remarks||'-')}</td><td class="text-center"><div class="action-btns"><button class="btn-edit" onclick="editOC('${r.id}')">Edit</button><button class="btn-del" onclick="deleteOC('${r.id}')">Del</button></div></td></tr>`).join('');
}
function exportOC(){ exportXLSX(byBranch(load('oc')), 'Offset_Clients', ['Branch','Month','Day','MFO','Client','Remarks'], r=>[bName(r.branch),r.month,r.day,r.mfo,r.client,r.remarks]); }

/* STAFF PERF */
function renderStaffPerf(){
  const m=document.getElementById('staffMonth').value, d=document.getElementById('staffDay').value;
  const nm=load('nm'), bd=load('bd'), del=load('del'), kam=load('kam');
  const match=r=>(!m||r.month===m)&&(!d||r.day===d)&&(!currentBranch||r.branch===currentBranch);
  const map={}; STAFF.forEach(s=>map[s.name]={staff:s.name, branch:s.branch, nm:0, off:0, net:0, bd:0, del:0, kam:0});
  nm.filter(match).forEach(r=>{ if(!map[r.staff])map[r.staff]={staff:r.staff,branch:r.branch,nm:0,off:0,net:0,bd:0,del:0,kam:0}; map[r.staff].nm+=num(r.nm); map[r.staff].off+=num(r.off); map[r.staff].net+=num(r.nm)-num(r.off); });
  bd.filter(match).forEach(r=>{ if(!map[r.staff])map[r.staff]={staff:r.staff,branch:r.branch,nm:0,off:0,net:0,bd:0,del:0,kam:0}; map[r.staff].bd+=num(r.amount); });
  del.filter(match).forEach(r=>{ (r.clients||[]).forEach(()=>{ if(!map[r.staff])map[r.staff]={staff:r.staff,branch:r.branch,nm:0,off:0,net:0,bd:0,del:0,kam:0}; map[r.staff].del+=1; }); });
  kam.filter(match).forEach(r=>{ if(!map[r.staff])map[r.staff]={staff:r.staff,branch:r.branch,nm:0,off:0,net:0,bd:0,del:0,kam:0}; map[r.staff].kam+=num(r.count); });
  const list=Object.values(map).sort((a,b)=>b.net-a.net);
  let tn=0,tb=0; list.forEach(r=>{tn+=r.net;tb+=r.bd;});
  document.getElementById('staffKPI').innerHTML=`<div class="kpi-card"><div class="kl">Staff Tracked</div><div class="kv kb">${list.length}</div></div><div class="kpi-card"><div class="kl">Total Net NM</div><div class="kv kg">${tn}</div></div><div class="kpi-card"><div class="kl">Total Bad Debt</div><div class="kv ko">₱${tb.toLocaleString()}</div></div>`;
  const b=document.getElementById('staffBody');
  if(!list.length){ b.innerHTML='<tr><td colspan="10" class="text-center" style="color:var(--muted);padding:20px">No data.</td></tr>'; return; }
  b.innerHTML=list.map((r,i)=>`<tr class="${i===0?'rank-1':i===1?'rank-2':i===2?'rank-3':''}"><td><strong>${i+1}</strong></td><td>${esc(r.staff)}</td><td>${esc(bName(r.branch))}</td><td>${m?MONTHS[parseInt(m)]:'All'}</td><td>${r.nm}</td><td>${r.off}</td><td style="color:var(--green);font-weight:700">${r.net}</td><td>₱${r.bd.toLocaleString()}</td><td>${r.del}</td><td>${r.kam}</td></tr>`).join('');
}
function exportStaffPerf(){ const m=document.getElementById('staffMonth').value, d=document.getElementById('staffDay').value;
  const nm=load('nm'), bd=load('bd'), del=load('del'), kam=load('kam');
  const match=r=>(!m||r.month===m)&&(!d||r.day===d); const map={};
  nm.filter(match).forEach(r=>{ if(!map[r.staff])map[r.staff]={staff:r.staff,branch:r.branch,nm:0,off:0,bd:0,del:0,kam:0}; map[r.staff].nm+=num(r.nm); map[r.staff].off+=num(r.off); });
  bd.filter(match).forEach(r=>{ if(!map[r.staff])map[r.staff]={staff:r.staff,branch:r.branch,nm:0,off:0,bd:0,del:0,kam:0}; map[r.staff].bd+=num(r.amount); });
  del.filter(match).forEach(r=>{ if(!map[r.staff])map[r.staff]={staff:r.staff,branch:r.branch,nm:0,off:0,bd:0,del:0,kam:0}; map[r.staff].del+=(r.clients?r.clients.length:1); });
  kam.filter(match).forEach(r=>{ if(!map[r.staff])map[r.staff]={staff:r.staff,branch:r.branch,nm:0,off:0,bd:0,del:0,kam:0}; map[r.staff].kam+=num(r.count); });
  const list=Object.values(map).sort((a,b)=>(b.nm-b.off)-(a.nm-a.off));
  exportXLSX(list, 'Staff_Performance', ['Staff','Branch','NM','Offset','Net NM','Bad Debt','Delinquent','Kamustahan'], r=>[r.staff,bName(r.branch),r.nm,r.off,r.nm-r.off,r.bd,r.del,r.kam]); }

/* INCENTIVES */
function autoFillIncBranch(){ document.getElementById('incBranchName').value=BRANCHES[document.getElementById('incBranchCode').value]||''; }
function clearINCForm(){ ['incBranchCode','incBranchName','incID','incLast','incFirst','incMiddle','incDesignation','incLevel'].forEach(id=>document.getElementById(id).value=''); ['incBorrowers','incPortfolio','incPAR','incAmount','incIncBorrowers','incIncLevel','incIncAmount'].forEach(id=>document.getElementById(id).value=''); editing.inc=null; document.getElementById('incSaveBtn').classList.remove('hidden'); document.getElementById('incUpdateBtn').classList.add('hidden'); document.getElementById('incCancelBtn').classList.add('hidden'); }
function readINC(){ return {staff:document.getElementById('incStaff').value, branchCode:document.getElementById('incBranchCode').value, branchName:document.getElementById('incBranchName').value, id:document.getElementById('incID').value.trim(), last:document.getElementById('incLast').value.trim(), first:document.getElementById('incFirst').value.trim(), middle:document.getElementById('incMiddle').value.trim(), designation:document.getElementById('incDesignation').value.trim(), level:document.getElementById('incLevel').value, borrowers:num(document.getElementById('incBorrowers').value), portfolio:num(document.getElementById('incPortfolio').value), par:num(document.getElementById('incPAR').value), amount:num(document.getElementById('incAmount').value), incBorrowers:document.getElementById('incIncBorrowers').value.trim(), incLevel:document.getElementById('incIncLevel').value.trim(), incAmount:document.getElementById('incIncAmount').value.trim()}; }
function saveINC(){ const v=readINC(); if(!v.branchCode||!v.last||!v.first||!v.level){alert('Fill Branch Code, Name and Level.');return;} const a=load('inc'); a.push({id:uid(),...v}); store('inc',a); clearINCForm(); renderINC(); }
function editINC(id){ const r=load('inc').find(x=>x.id===id); if(!r)return; editing.inc=id;
  document.getElementById('incStaff').value=r.staff||''; document.getElementById('incBranchCode').value=r.branchCode; document.getElementById('incBranchName').value=r.branchName; document.getElementById('incID').value=r.id||''; document.getElementById('incLast').value=r.last; document.getElementById('incFirst').value=r.first; document.getElementById('incMiddle').value=r.middle||''; document.getElementById('incDesignation').value=r.designation||''; document.getElementById('incLevel').value=r.level; document.getElementById('incBorrowers').value=r.borrowers; document.getElementById('incPortfolio').value=r.portfolio; document.getElementById('incPAR').value=r.par; document.getElementById('incAmount').value=r.amount; document.getElementById('incIncBorrowers').value=r.incBorrowers||''; document.getElementById('incIncLevel').value=r.incLevel||''; document.getElementById('incIncAmount').value=r.incAmount||'';
  document.getElementById('incSaveBtn').classList.add('hidden'); document.getElementById('incUpdateBtn').classList.remove('hidden'); document.getElementById('incCancelBtn').classList.remove('hidden'); }
function updateINC(){ const a=load('inc'); const i=a.findIndex(x=>x.id===editing.inc); if(i===-1)return; a[i]={...a[i],...readINC()}; store('inc',a); clearINCForm(); renderINC(); }
function cancelEditINC(){ clearINCForm(); }
function deleteINC(id){ if(!confirm('Delete this Incentive record?'))return; store('inc', load('inc').filter(x=>x.id!==id)); renderINC(); }
function renderINC(){
  const a=load('inc'); const lv={}; let ta=0,tb=0;
  a.forEach(r=>{ const l=r.level||'Unspecified'; if(!lv[l])lv[l]={count:0,borr:0,amt:0}; lv[l].count++; lv[l].borr+=r.borrowers; lv[l].amt+=r.amount; ta+=r.amount; tb+=r.borrowers; });
  document.getElementById('incKPI').innerHTML=`<div class="kpi-card"><div class="kl">Staff w/ Incentive</div><div class="kv kamb">${a.length}</div></div><div class="kpi-card"><div class="kl">Total Borrowers</div><div class="kv klime">${tb}</div></div><div class="kpi-card"><div class="kl">Total Amount</div><div class="kv kg">₱${ta.toLocaleString()}</div></div>`;
  let rows=''; Object.keys(lv).forEach(l=>{ const v=lv[l]; rows+=`<tr><td><strong>${esc(l)}</strong></td><td>${v.count}</td><td>${v.borr}</td><td>₱${v.amt.toLocaleString()}</td><td>${ta>0?((v.amt/ta)*100).toFixed(1):0}%</td></tr>`; });
  rows+=`<tr class="total-row"><td>TOTAL</td><td>${a.length}</td><td>${tb}</td><td>₱${ta.toLocaleString()}</td><td>100%</td></tr>`;
  document.getElementById('incLevelBody').innerHTML=rows||'<tr><td colspan="5" class="text-center" style="color:var(--muted);padding:15px">No entries.</td></tr>';
  const b=document.getElementById('incBody');
  if(!a.length){ b.innerHTML='<tr><td colspan="12" class="text-center" style="color:var(--muted);padding:20px">No entries yet.</td></tr>'; return; }
  b.innerHTML=a.map((r,i)=>`<tr><td>${i+1}</td><td>${esc(r.staff||'-')}</td><td>${esc(r.branchCode)}</td><td>${esc(r.branchName)}</td><td>${esc(r.id||'-')}</td><td>${esc(r.last+', '+r.first+(r.middle?' '+r.middle:''))}</td><td>${esc(r.designation||'-')}</td><td>${esc(r.level)}</td><td>${r.borrowers}</td><td>₱${r.portfolio.toLocaleString()}</td><td>${r.par}%</td><td>₱${r.amount.toLocaleString()}</td><td>${esc(r.incBorrowers||'-')}</td><td>${esc(r.incLevel||'-')}</td><td>${esc(r.incAmount||'-')}</td><td class="text-center"><div class="action-btns"><button class="btn-edit" onclick="editINC('${r.id}')">Edit</button><button class="btn-del" onclick="deleteINC('${r.id}')">Del</button></div></td></tr>`).join('');
}
function exportINC(){ exportXLSX(load('inc'), 'Incentives', ['Staff','Branch Code','Branch Name','ID','Last','First','Middle','Designation','Level','Borrowers','Portfolio','PAR%','Amount','Increase Borrowers','Increase Level','Increase Amount'], r=>[r.staff||'',r.branchCode,r.branchName,r.id,r.last,r.first,r.middle,r.designation,r.level,r.borrowers,r.portfolio,r.par,r.amount,r.incBorrowers||'',r.incLevel||'',r.incAmount||'']); }

/* UPDATE TRACKER */
function switchUpdateView(mode){
  document.getElementById('staffView').classList.toggle('hidden', mode!=='staff');
  document.getElementById('branchView').classList.toggle('hidden', mode!=='branch');
  document.getElementById('viewStaffBtn').className=mode==='staff'?'btn-view-on':'btn-view-off';
  document.getElementById('viewBranchBtn').className=mode==='branch'?'btn-view-on':'btn-view-off';
}
function renderUpdateTracker(){
  const m=document.getElementById('updateMonth').value, d=document.getElementById('updateDay').value;
  const nm=load('nm'), bd=load('bd'), del=load('del'), kam=load('kam');
  const has=(arr,s)=>arr.some(r=>r.staff===s&&(!m||r.month===m)&&(!d||r.day===d));
  const set={}; STAFF.forEach(s=>set[s.name]={staff:s.name,branch:s.branch,nm:false,bd:false,del:false,kam:false});
  [...nm,...bd,...del,...kam].forEach(r=>{ if(!set[r.staff])set[r.staff]={staff:r.staff,branch:r.branch,nm:false,bd:false,del:false,kam:false}; });
  Object.keys(set).forEach(n=>{ const s=set[n]; s.nm=has(nm,n); s.bd=has(bd,n); s.del=has(del,n); s.kam=has(kam,n); s.score=[s.nm,s.bd,s.del,s.kam].filter(Boolean).length; });
  const list=Object.values(set);
  const comp=list.filter(s=>s.score===4).length, part=list.filter(s=>s.score>0&&s.score<4).length, none=list.filter(s=>s.score===0).length;
  document.getElementById('updateKPI').innerHTML=`<div class="kpi-card"><div class="kl">Total Staff</div><div class="kv kb">${list.length}</div></div><div class="kpi-card"><div class="kl">Complete</div><div class="kv kg">${comp}</div></div><div class="kpi-card"><div class="kl">Partial</div><div class="kv ko">${part}</div></div><div class="kpi-card"><div class="kl">Not Submitted</div><div class="kv kr">${none}</div></div>`;
  document.getElementById('updateStaffBody').innerHTML=list.map(s=>{ const st=s.score===4?'<span class="ok">✅ COMPLETE</span>':s.score>0?'<span class="partial">⏳ PARTIAL ('+s.score+'/4)</span>':'<span class="no">❌ NONE</span>';
    return `<tr><td>${esc(s.staff)}</td><td>${esc(bName(s.branch))}</td><td class="text-center">${s.nm?'<span class="ok">✅</span>':'<span class="no">❌</span>'}</td><td class="text-center">${s.bd?'<span class="ok">✅</span>':'<span class="no">❌</span>'}</td><td class="text-center">${s.del?'<span class="ok">✅</span>':'<span class="no">❌</span>'}</td><td class="text-center">${s.kam?'<span class="ok">✅</span>':'<span class="no">❌</span>'}</td><td>${st}</td></tr>`; }).join('')||'<tr><td colspan="7" class="text-center" style="color:var(--muted);padding:15px">No staff data.</td></tr>';
  const bb={}; Object.keys(BRANCHES).forEach(c=>bb[c]={total:0,sub:0,nm:0,bd:0,del:0,kam:0,comp:0});
  list.forEach(s=>{ const b=s.branch||'B0814'; if(!bb[b])bb[b]={total:0,sub:0,nm:0,bd:0,del:0,kam:0,comp:0}; bb[b].total++; if(s.score>0)bb[b].sub++; if(s.nm)bb[b].nm++; if(s.bd)bb[b].bd++; if(s.del)bb[b].del++; if(s.kam)bb[b].kam++; if(s.score===4)bb[b].comp++; });
  document.getElementById('updateBranchBody').innerHTML=Object.keys(bb).map(c=>{ const b=bb[c]; const pct=b.total>0?((b.comp/b.total)*100).toFixed(0):0; const st=pct==100?'<span class="ok">✅ DONE</span>':pct>=50?'<span class="partial">⏳ ONGOING</span>':'<span class="no">❌ BEHIND</span>';
    return `<tr><td><strong>${esc(bName(c))}</strong> <span style="color:var(--muted);font-size:8px">${c}</span></td><td>${b.total}</td><td>${b.sub}</td><td>${b.nm}</td><td>${b.bd}</td><td>${b.del}</td><td>${b.kam}</td><td>${b.comp}</td><td>${pct}%</td><td>${st}</td></tr>`; }).join('');
}

/* LOCK */
function renderLockedList(){ const ld=lockedDates(); const el=document.getElementById('dateLockedList'); if(!ld.length){el.innerHTML='<span style="color:var(--muted)">No locked dates.</span>';return;} el.innerHTML='<strong style="color:var(--gold)">🔒 Locked dates:</strong> '+ld.map(x=>MONTHS[parseInt(x.slice(0,2))]+' '+x.slice(3)).join(', '); }
function lockSpecificDate(){ const m=document.getElementById('dateLockMonth').value, d=document.getElementById('dateLockDay').value; if(!m||!d){alert('Select month and day.');return;} const p=prompt('Enter lock password:'); if(p!==LOCK_PASSWORD){alert('Wrong password.');return;} const ld=lockedDates(); const k=m+'-'+d; if(ld.indexOf(k)===-1){ld.push(k);store('lockedDates',ld);} alert('🔒 Locked.'); renderLockedList(); renderAll(); }
function unlockSpecificDate(){ const m=document.getElementById('dateLockMonth').value, d=document.getElementById('dateLockDay').value; if(!m||!d){alert('Select month and day.');return;} const p=prompt('Enter unlock password:'); if(p!==LOCK_PASSWORD){alert('Wrong password.');return;} store('lockedDates', lockedDates().filter(x=>x!==m+'-'+d)); alert('🔓 Unlocked.'); renderLockedList(); renderAll(); }

/* SYNC / SHARE */
function allData(){ return {nm:load('nm'),bd:load('bd'),del:load('del'),kam:load('kam'),oc:load('oc'),inc:load('inc'),exc:load('exc'),lockedDates:load('lockedDates')}; }
function encodeSync(){ return btoa(unescape(encodeURIComponent(JSON.stringify(allData())))); }
function generateSyncCode(){ const c=encodeSync(); document.getElementById('shareLink').value=c; alert('✅ Sync code generated! Click COPY to copy it.'); }
function copySyncCode(){ const t=document.getElementById('shareLink'); if(!t.value){alert('Generate code first.');return;} t.select(); document.execCommand('copy'); alert('✅ Copied! Paste it in the CONNECT box on another device.'); }
function connectSync(){ const code=document.getElementById('syncIn').value.trim(); if(!code){alert('Paste a sync code first.');return;}
  try{ const data=JSON.parse(decodeURIComponent(escape(atob(code))));
    ['nm','bd','del','kam','oc','inc','exc'].forEach(k=>{ if(!Array.isArray(data[k]))return; const cur=load(k); const ids=new Set(cur.map(x=>x.id)); data[k].forEach(r=>{ if(r.id&&!ids.has(r.id))cur.push(r); }); store(k,cur); });
    if(Array.isArray(data.lockedDates))store('lockedDates', data.lockedDates);
    alert('✅ Sync complete! Records merged.'); document.getElementById('syncIn').value=''; renderAll();
  }catch(e){ alert('❌ Invalid sync code.'); } }
function downloadShareableFile(){ const data=JSON.stringify(allData());
  const seed=`<script>try{var D=JSON.parse('${data.replace(/\\/g,'\\\\').replace(/'/g,"\\'")}');Object.keys(D).forEach(function(k){localStorage.setItem('${LS}'+k, JSON.stringify(D[k]));});}catch(e){}<\/script>`;
  const html=document.documentElement.outerHTML.replace('</body>', seed+'</body>');
  const blob=new Blob([html],{type:'text/html'}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='Sorsogon1_Dashboard_Shareable.html'; a.click(); }

/* EXCEL */
function exportXLSX(rows, sheet, headers, mapFn){ if(!rows.length){alert('No data to export.');return;} const data=[headers,...rows.map(mapFn)]; const ws=XLSX.utils.aoa_to_sheet(data); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,sheet); XLSX.writeFile(wb, sheet+'_'+new Date().toISOString().slice(0,10)+'.xlsx'); }
function downloadTemplate(){ const ws=XLSX.utils.aoa_to_sheet([['Branch','Staff','Month(MM)','Day(DD)','NM','Offset','Target NM','Target Offset'],['B0814','Sample Staff','10','01',5,1,10,2]]); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'NM_Upload'); XLSX.writeFile(wb,'NM_Upload_Template.xlsx'); }
function handleFileUpload(e){ const file=e.target.files[0]; if(!file)return; const m=document.getElementById('uploadMonth').value, d=document.getElementById('uploadDay').value; const res=document.getElementById('uploadResult');
  const rd=new FileReader(); rd.onload=function(ev){ try{ const wb=XLSX.read(ev.target.result,{type:'array'}); const ws=wb.Sheets[wb.SheetNames[0]]; const rows=XLSX.utils.sheet_to_json(ws,{header:1});
    let count=0; const a=load('nm'); rows.slice(1).forEach(r=>{ if(!r[1])return; const branch=r[0]&&BRANCHES[r[0]]?r[0]:(currentBranch||'B0814');
      a.push({id:uid(), branch, staff:String(r[1]), month:m||String(r[2]||'').padStart(2,'0'), day:(d?String(d).padStart(2,'0'):String(r[3]||'').padStart(2,'0')), nm:num(r[4]), off:num(r[5]), tnm:num(r[6]), toff:num(r[7])}); count++; });
    store('nm',a); res.className='upload-result upload-success'; res.textContent='✅ Uploaded '+count+' NM records!'; renderNM(); renderAll();
  }catch(err){ res.className='upload-result upload-error'; res.textContent='❌ Upload failed: '+err.message; } e.target.value=''; };
  rd.readAsArrayBuffer(file); }

function renderAll(){ renderOverview(); renderNM(); renderBD(); renderDel(); renderEXC(); renderKam(); renderOC(); renderStaffPerf(); renderINC(); renderUpdateTracker(); renderLockedList(); }
init();
</script>
</body>
</html>
