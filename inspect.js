// inspect.js — IPQC / OQC Inspection Module v9.68

/* ── State ─────────────────────────────────────────────── */
let _inspList    = [];
let _inspOrders  = [];
let _inspInited  = false;
let _inspFilter  = 'all';
let _inspPoints  = [];
let _inspEditId  = null;   // inspId ที่กำลังแก้ไข (null = สร้างใหม่)
let _inspViewRow = null;

/* ── Default check points per type ─────────────────────── */
const _INSP_DEF = {
  IPQC: [
    { cp:'OD',    smin:'', smax:'', unit:'mm',  vis:false, actual:'' },
    { cp:'ID',    smin:'', smax:'', unit:'mm',  vis:false, actual:'' },
    { cp:'H',     smin:'', smax:'', unit:'mm',  vis:false, actual:'' },
    { cp:'ผิวงาน', smin:'', smax:'', unit:'', vis:true,  actual:'' },
  ],
  OQC: [
    { cp:'OD',   smin:'', smax:'', unit:'mm',  vis:false, actual:'' },
    { cp:'ID',   smin:'', smax:'', unit:'mm',  vis:false, actual:'' },
    { cp:'H',    smin:'', smax:'', unit:'mm',  vis:false, actual:'' },
    { cp:'ผิวชุบ',                smin:'', smax:'', unit:'', vis:true,  actual:'' },
    { cp:'สีชุบ – ไม่มีคราบน้ำไหล', smin:'', smax:'', unit:'', vis:true,  actual:'' },
    { cp:'ความหนาชุบ',            smin:'10', smax:'30', unit:'μm', vis:false, actual:'' },
  ],
};

/* ── Helpers ────────────────────────────────────────────── */
function _inspUrl()  { return localStorage.getItem('ptts_script_url') || ''; }
function _inspToday(){ const d=new Date(); return d.toISOString().slice(0,10); }

function _inspInputToSheet(s) {
  if (!s) return '';
  const [y,m,d] = s.split('-');
  return d+'/'+m+'/'+(parseInt(y)+543);
}
function _inspSheetToInput(s) {
  if (!s) return '';
  const p = s.split('/');
  if (p.length !== 3) return '';
  const yr = parseInt(p[2]);
  const ce = yr > 2400 ? yr-543 : yr;
  return String(ce).padStart(4,'0')+'-'+p[1]+'-'+p[0];
}

function _inspBadge(result) {
  const cls = {OK:'b-pass',Hold:'b-hold',NG:'b-fail'}[result]||'';
  return `<span class="isp-badge ${cls}">${result||'—'}</span>`;
}
function _inspTypeBadge(type) {
  return type==='IPQC'
    ? '<span class="isp-badge b-ipqc">IPQC</span>'
    : '<span class="isp-badge b-oqc">OQC</span>';
}

/* ── Entry point ────────────────────────────────────────── */
async function _inspInit() {
  const el = document.getElementById('tab-inspect');
  if (!el) return;
  if (!_inspInited) {
    el.innerHTML = _inspShell();
    _inspInited = true;
  }
  await _inspLoad();
}

/* ── Shell HTML + CSS ───────────────────────────────────── */
function _inspShell() {
  return `
<style>
#_ispWrap{padding:16px 20px 48px}
.isp-topbar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:14px}
.isp-topbar-title{font-size:.95rem;font-weight:800;color:var(--c1);margin-right:4px}
.isp-toggle{display:flex;background:var(--hover);border-radius:20px;padding:3px;gap:2px}
.isp-tgl{font-size:.72rem;font-weight:700;padding:4px 13px;border-radius:16px;border:none;background:transparent;color:var(--t3);cursor:pointer;transition:all .15s}
.isp-tgl.act{background:var(--c1);color:#fff}
.isp-tgl.act-oqc{background:#7c3aed;color:#fff}
.isp-spacer{flex:1}
.isp-btn-pri{font-size:.78rem;padding:7px 16px;border-radius:20px;border:none;background:var(--c1);color:#fff;cursor:pointer;font-weight:700;transition:all .15s}
.isp-btn-pri:hover{opacity:.85;transform:translateY(-1px)}
.isp-btn-out{font-size:.78rem;padding:6px 14px;border-radius:20px;border:1.5px solid var(--border);background:transparent;color:var(--t2);cursor:pointer;font-weight:600;transition:all .15s}
.isp-btn-out:hover{border-color:var(--c1);color:var(--c1)}
.isp-stat-row{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px}
.isp-stat{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:12px 14px;box-shadow:0 1px 4px rgba(0,0,0,.04)}
.isp-stat-lbl{font-size:.68rem;color:var(--t3);font-weight:700;margin-bottom:3px}
.isp-stat-val{font-size:1.5rem;font-weight:800;color:var(--t1);line-height:1.1}
.isp-stat-sub{font-size:.68rem;margin-top:2px}
.isp-c-pass{color:#16a34a}.isp-c-fail{color:#dc2626}.isp-c-hold{color:#d97706}.isp-c-blue{color:var(--c1)}
.isp-card{background:var(--card);border:1px solid var(--border);border-radius:12px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.04)}
.isp-card-hdr{padding:10px 14px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:8px}
.isp-card-title{font-size:.7rem;font-weight:800;color:var(--t3);text-transform:uppercase;letter-spacing:.05em}
table.isp-tbl{width:100%;border-collapse:collapse;font-size:.8rem}
table.isp-tbl th{padding:8px 10px;font-size:.68rem;font-weight:700;color:var(--t3);text-align:left;border-bottom:1.5px solid var(--border);background:var(--hover);white-space:nowrap}
table.isp-tbl td{padding:8px 10px;border-bottom:1px solid var(--border);color:var(--t1);vertical-align:middle}
table.isp-tbl tr:last-child td{border-bottom:none}
table.isp-tbl tbody tr:hover td{background:var(--hover)}
.isp-badge{display:inline-flex;align-items:center;padding:2px 9px;border-radius:99px;font-size:.68rem;font-weight:700;white-space:nowrap}
.b-pass{background:#dcfce7;color:#15803d}
.b-fail{background:#fee2e2;color:#dc2626}
.b-hold{background:#fef9c3;color:#b45309}
.b-ipqc{background:#dbeafe;color:#1d4ed8}
.b-oqc{background:#ede9fe;color:#7c3aed}
.isp-btn-sm{font-size:.7rem;padding:3px 9px;border-radius:7px;border:1px solid var(--border);background:var(--card);cursor:pointer;color:var(--t2);font-weight:600;margin-right:2px}
.isp-btn-view{font-size:.68rem;padding:3px 8px;border-radius:6px;border:1.5px solid #2563eb;background:#eff6ff;cursor:pointer;color:#1d4ed8;font-weight:700;white-space:nowrap}
.isp-btn-view:hover{background:#dbeafe}
.isp-btn-edit{font-size:.68rem;padding:3px 8px;border-radius:6px;border:1.5px solid #d97706;background:#fffbeb;cursor:pointer;color:#b45309;font-weight:700;white-space:nowrap}
.isp-btn-edit:hover{background:#fef3c7}
.isp-btn-del{font-size:.68rem;padding:3px 8px;border-radius:6px;border:1.5px solid #dc2626;background:#fef2f2;cursor:pointer;color:#dc2626;font-weight:700;white-space:nowrap}
.isp-btn-del:hover{background:#fee2e2}
.isp-empty{text-align:center;padding:40px;color:var(--t3);font-size:.9rem}
/* Panel */
#_ispPanel{position:fixed;inset:0;background:rgba(15,23,42,.45);z-index:9000;display:none}
#_ispPanelBox{position:fixed;right:0;top:0;bottom:0;height:100%;width:min(620px,100%);background:var(--bg-card,#ffffff);overflow-y:auto;display:flex;flex-direction:column;box-shadow:-6px 0 32px rgba(0,0,0,.22)}
[data-theme="dark"] #_ispPanelBox{background:#1e2535}
.isp-panel-hdr{padding:13px 18px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px;position:sticky;top:0;background:inherit;z-index:10;flex-shrink:0}
.isp-panel-title{font-size:.92rem;font-weight:800;color:var(--t1)}
.isp-panel-close{margin-left:auto;width:28px;height:28px;border-radius:50%;border:1px solid var(--border);background:transparent;cursor:pointer;color:var(--t3);font-size:13px;display:flex;align-items:center;justify-content:center}
.isp-panel-body{padding:16px 18px;flex:1}
.isp-panel-footer{padding:12px 18px;border-top:1px solid var(--border);display:flex;gap:8px;flex-shrink:0;position:sticky;bottom:0;background:inherit}
.isp-panel-footer .isp-btn-pri{flex:1;padding:9px;font-size:.82rem}
/* Form */
.isp-field-row{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px}
.isp-field{display:flex;flex-direction:column;gap:4px}
.isp-field.full{grid-column:1/-1}
.isp-field label{font-size:.68rem;font-weight:700;color:var(--t3)}
.isp-field input,.isp-field select,.isp-field textarea{font-family:inherit;font-size:.8rem;padding:7px 10px;border:1.5px solid var(--border);border-radius:8px;color:var(--t1);background:var(--card);outline:none;width:100%}
.isp-field input:focus,.isp-field select:focus{border-color:var(--c1)}
.isp-info-box{background:var(--hover);border:1px solid var(--border);border-radius:10px;padding:10px 12px;margin-bottom:12px;font-size:.78rem}
.isp-info-box .isp-info-lbl{font-size:.65rem;font-weight:700;color:var(--t3);margin-bottom:5px}
.isp-info-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:5px;font-size:.75rem}
.isp-info-grid .ik{color:var(--t3)}
.isp-sep{font-size:.68rem;font-weight:800;color:var(--t3);text-transform:uppercase;letter-spacing:.06em;margin:14px 0 8px;padding-bottom:6px;border-bottom:1.5px solid var(--border)}
/* Check-point table */
table.isp-cp{width:100%;border-collapse:collapse;font-size:.76rem;margin-bottom:8px}
table.isp-cp th{font-size:.65rem;color:var(--t3);font-weight:700;padding:5px 7px;border-bottom:1.5px solid var(--border);text-align:left;background:var(--hover)}
table.isp-cp td{padding:5px 7px;border-bottom:1px solid var(--border);vertical-align:middle}
table.isp-cp input,table.isp-cp select{font-family:inherit;font-size:.75rem;padding:3px 6px;border:1px solid var(--border);border-radius:6px;background:var(--card);color:var(--t1);width:100%}
.isp-ok{display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;border-radius:50%;font-size:11px;font-weight:700}
.isp-ok-y{background:#dcfce7;color:#15803d}
.isp-ok-n{background:#fee2e2;color:#dc2626}
.isp-ok-na{background:var(--hover);color:var(--t3)}
/* Result */
.isp-result-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:12px}
.isp-res-box{background:var(--hover);border:1.5px solid var(--border);border-radius:10px;padding:9px;text-align:center}
.isp-res-box label{font-size:.65rem;font-weight:700;color:var(--t3);display:block;margin-bottom:3px}
.isp-res-box input{text-align:center;font-weight:800;font-size:1rem;border:none;background:transparent;color:var(--t1);width:100%;font-family:inherit}
.isp-result-btns{display:flex;gap:8px;margin-bottom:12px}
.isp-rb{flex:1;padding:8px;border-radius:10px;border:2px solid var(--border);background:var(--card);cursor:pointer;font-weight:700;font-size:.8rem;transition:all .15s;color:var(--t2)}
.isp-rb.sel-ok{border-color:#16a34a;background:#dcfce7;color:#15803d}
.isp-rb.sel-hold{border-color:#d97706;background:#fef9c3;color:#b45309}
.isp-rb.sel-ng{border-color:#dc2626;background:#fee2e2;color:#dc2626}
@media(max-width:640px){
  .isp-stat-row{grid-template-columns:1fr 1fr}
  .isp-field-row{grid-template-columns:1fr}
  #_ispPanelBox{width:100%}
}
</style>
<div id="_ispWrap">
  <div class="isp-topbar">
    <span class="isp-topbar-title">🔍 ใบตรวจสอบ</span>
    <div class="isp-toggle">
      <button class="isp-tgl act" id="_ispTAll"  onclick="_inspSetFilter('all')">ทั้งหมด</button>
      <button class="isp-tgl"     id="_ispTI"    onclick="_inspSetFilter('IPQC')">IPQC</button>
      <button class="isp-tgl"     id="_ispTO"    onclick="_inspSetFilter('OQC')">OQC</button>
    </div>
    <div class="isp-spacer"></div>
    <button class="isp-btn-pri" onclick="_inspOpenPanel()">+ สร้างใบตรวจ</button>
  </div>
  <div class="isp-stat-row" id="_ispStats"></div>
  <div class="isp-card">
    <div class="isp-card-hdr"><span class="isp-card-title">รายการใบตรวจ</span></div>
    <div style="overflow-x:auto">
      <table class="isp-tbl">
        <thead><tr>
          <th>เลขที่</th><th>ประเภท</th><th>No.PO</th><th>สินค้า / Spec</th>
          <th style="text-align:center">ตรวจ</th><th style="text-align:center">ผ่าน</th>
          <th>วันที่</th><th>ผู้ตรวจ</th><th>ผล</th><th style="width:155px"></th>
        </tr></thead>
        <tbody id="_ispTbody"><tr><td colspan="10" class="isp-empty">กำลังโหลด...</td></tr></tbody>
      </table>
    </div>
  </div>
</div>
<div id="_ispPanel" onclick="if(event.target===this)_inspClosePanel()">
  <div id="_ispPanelBox">
    <div class="isp-panel-hdr">
      <span class="isp-panel-title" id="_ispPanelTitle">สร้างใบตรวจ</span>
      <button class="isp-panel-close" onclick="_inspClosePanel()">✕</button>
    </div>
    <div class="isp-panel-body" id="_ispPanelBody"></div>
    <div class="isp-panel-footer" id="_ispPanelFooter"></div>
  </div>
</div>
`;
}

/* ── Load list ──────────────────────────────────────────── */
async function _inspLoad() {
  const url = _inspUrl();
  if (!url) { _inspRenderEmpty('กรุณาตั้งค่า Script URL ในแท็บ "ตั้งค่า" ก่อน'); return; }
  try {
    const r = await fetch(url + '?action=getInspections');
    const d = await r.json();
    _inspList = d.rows || [];
    _inspRender();
  } catch(e) {
    _inspRenderEmpty('โหลดข้อมูลไม่ได้: ' + e.message);
  }
}

function _inspRenderEmpty(msg) {
  const tb = document.getElementById('_ispTbody');
  if (tb) tb.innerHTML = `<tr><td colspan="10" class="isp-empty">${msg}</td></tr>`;
  _inspRenderStats([]);
}

/* ── Render ─────────────────────────────────────────────── */
function _inspRender() {
  let rows = _inspList;
  if (_inspFilter !== 'all') rows = rows.filter(r => r.type === _inspFilter);
  _inspRenderStats(rows);
  _inspRenderTable(rows);
}

function _inspRenderStats(rows) {
  const el = document.getElementById('_ispStats');
  if (!el) return;
  const total = rows.length;
  const pass  = rows.filter(r => r.result === 'OK').length;
  const hold  = rows.filter(r => r.result === 'Hold').length;
  const fail  = rows.filter(r => r.result === 'NG').length;
  el.innerHTML = `
    <div class="isp-stat"><div class="isp-stat-lbl">ใบตรวจทั้งหมด</div><div class="isp-stat-val isp-c-blue">${total}</div><div class="isp-stat-sub" style="color:var(--t3)">รายการ</div></div>
    <div class="isp-stat"><div class="isp-stat-lbl">ผ่าน (OK)</div><div class="isp-stat-val isp-c-pass">${pass}</div><div class="isp-stat-sub isp-c-pass">${total?Math.round(pass/total*100):0}%</div></div>
    <div class="isp-stat"><div class="isp-stat-lbl">Hold / Rework</div><div class="isp-stat-val isp-c-hold">${hold}</div><div class="isp-stat-sub isp-c-hold">${total?Math.round(hold/total*100):0}%</div></div>
    <div class="isp-stat"><div class="isp-stat-lbl">ไม่ผ่าน (NG)</div><div class="isp-stat-val isp-c-fail">${fail}</div><div class="isp-stat-sub isp-c-fail">${total?Math.round(fail/total*100):0}%</div></div>
  `;
}

function _inspRenderTable(rows) {
  const tb = document.getElementById('_ispTbody');
  if (!tb) return;
  if (!rows.length) { tb.innerHTML = `<tr><td colspan="10" class="isp-empty">ไม่มีข้อมูล</td></tr>`; return; }
  tb.innerHTML = rows.map(r => `
    <tr>
      <td><strong>${r.inspId||'—'}</strong></td>
      <td>${_inspTypeBadge(r.type)}</td>
      <td>${r.orderId||'—'}</td>
      <td style="max-width:170px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${r.orderRef||''}">${r.orderRef||r.workType||'—'}</td>
      <td style="text-align:center">${r.qtyCheck != null ? r.qtyCheck : '—'}</td>
      <td style="text-align:center">${r.qtyPass != null ? r.qtyPass : '—'}</td>
      <td>${r.date||'—'}</td>
      <td>${r.inspector||'—'}</td>
      <td>${_inspBadge(r.result)}</td>
      <td>
        <div style="display:flex;gap:4px;justify-content:flex-end;flex-wrap:nowrap">
          <button class="isp-btn-view" onclick="_inspViewDetail('${r.inspId}')">🔍 ดู</button>
          <button class="isp-btn-edit" onclick="_inspEditLoad('${r.inspId}')">✏️ แก้ไข</button>
          <button class="isp-btn-del"  onclick="_inspDelete('${r.inspId}')">🗑 ลบ</button>
        </div>
      </td>
    </tr>
  `).join('');
}

/* ── Filter toggle ──────────────────────────────────────── */
function _inspSetFilter(f) {
  _inspFilter = f;
  [['_ispTAll','all'],['_ispTI','IPQC'],['_ispTO','OQC']].forEach(([id,val]) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.classList.remove('act','act-oqc');
    if (val === f) el.classList.add(f === 'OQC' ? 'act-oqc' : 'act');
  });
  _inspRender();
}

/* ── Load orders for dropdown ───────────────────────────── */
async function _inspLoadOrders() {
  if (_inspOrders.length) return;
  const url = _inspUrl();
  if (!url) return;
  try {
    const r = await fetch(url + '?action=getOrders');
    const d = await r.json();
    _inspOrders = (d.rows || d || []).filter(row => Array.isArray(row) && row[1]);
  } catch(e) { _inspOrders = []; }
}

/* ── Open create panel ──────────────────────────────────── */
async function _inspOpenPanel() {
  _inspEditId = null;
  await _inspLoadOrders();
  document.getElementById('_ispPanelTitle').textContent = _inspEditId ? '✏️ แก้ไขใบตรวจ' : '+ สร้างใบตรวจ';
  document.getElementById('_ispPanelBody').innerHTML = _inspCreateFormHtml();
  document.getElementById('_ispPanelFooter').innerHTML = `
    <button class="isp-btn-out" onclick="_inspClosePanel()">ยกเลิก</button>
    <button class="isp-btn-pri" onclick="_inspSave(false)">💾 บันทึก</button>
    <button class="isp-btn-pri" style="background:#7c3aed" onclick="_inspSave(true)">🖨 บันทึก + พิมพ์</button>
  `;
  document.getElementById('_ispPanel').style.display = 'block';
  _inspSetDefaultPoints('OQC');
  _inspRenderCpTable();
}

function _inspClosePanel() {
  document.getElementById('_ispPanel').style.display = 'none';
}

/* ── Create form HTML ───────────────────────────────────── */
function _inspCreateFormHtml() {
  const orderOpts = _inspOrders.map(o => {
    const noPO = String(o[1]||'');
    const wt   = String(o[4]||'');
    const od   = String(o[34]||'');
    const id_  = String(o[35]||'');
    const h    = String(o[36]||'');
    const spec = od ? ('OD'+od+(id_?'xID'+id_:'')+(h?'xH'+h:'')+'mm') : '';
    const lbl  = noPO + (spec?' · '+spec : wt?' · '+wt:'');
    return `<option value="${noPO}" data-wt="${wt}" data-od="${od}" data-id="${id_}" data-h="${h}">${lbl}</option>`;
  }).join('');
  return `
<div class="isp-field-row">
  <div class="isp-field">
    <label>ประเภทการตรวจ</label>
    <select id="_ispFType" onchange="_inspOnTypeChange()">
      <option value="OQC">OQC — ก่อนส่งมอบ</option>
      <option value="IPQC">IPQC — ระหว่างผลิต</option>
    </select>
  </div>
  <div class="isp-field">
    <label>Order (No.PO)</label>
    <select id="_ispFOrder" onchange="_inspOnOrderChange()">
      <option value="">— เลือก Order —</option>
      ${orderOpts}
    </select>
  </div>
</div>
<div class="isp-info-box">
  <div class="isp-info-lbl">ข้อมูลจาก Order (อัตโนมัติ)</div>
  <div class="isp-info-grid">
    <span><span class="ik">ประเภทงาน: </span><strong id="_ispFWt">—</strong></span>
    <span><span class="ik">OD: </span><span id="_ispFOd">—</span></span>
    <span><span class="ik">ID: </span><span id="_ispFId">—</span></span>
    <span><span class="ik">H: </span><span id="_ispFH">—</span></span>
    <span style="grid-column:1/-1"><span class="ik">Spec: </span><span id="_ispFSpec">—</span></span>
  </div>
</div>
<div class="isp-field-row">
  <div class="isp-field">
    <label>วันที่ตรวจ</label>
    <input type="date" id="_ispFDate" value="${_inspToday()}">
  </div>
  <div class="isp-field">
    <label>ผู้ตรวจ</label>
    <input type="text" id="_ispFInspector" placeholder="ชื่อผู้ตรวจ">
  </div>
</div>
<div class="isp-sep">จุดตรวจ (Check Points)</div>
<div id="_ispCpWrap"></div>
<button class="isp-btn-out" style="width:100%;margin-bottom:14px;font-size:.72rem" onclick="_inspAddPoint()">+ เพิ่มจุดตรวจ</button>
<div class="isp-sep">สรุปผล</div>
<div class="isp-result-grid">
  <div class="isp-res-box"><label>จำนวนตรวจ</label><input id="_ispFQtyC" type="number" min="0" placeholder="0"></div>
  <div class="isp-res-box"><label>ผ่าน</label><input id="_ispFQtyP" type="number" min="0" placeholder="0" style="color:#15803d"></div>
  <div class="isp-res-box"><label>ไม่ผ่าน</label><input id="_ispFQtyF" type="number" min="0" placeholder="0" style="color:#dc2626"></div>
</div>
<div style="font-size:.68rem;font-weight:700;color:var(--t3);margin-bottom:6px">ผลการตรวจ</div>
<div class="isp-result-btns">
  <button class="isp-rb sel-ok" id="_ispROK" onclick="_inspSelectResult('OK')">✓ OK</button>
  <button class="isp-rb"          id="_ispRHold" onclick="_inspSelectResult('Hold')">⚠ Hold</button>
  <button class="isp-rb"          id="_ispRNG" onclick="_inspSelectResult('NG')">✗ NG</button>
</div>
<div class="isp-field full" style="margin-bottom:0">
  <label>หมายเหตุ / Action ที่ต้องทำ</label>
  <textarea id="_ispFNote" rows="2" placeholder="ระบุถ้ามี..."></textarea>
</div>
`;
}

/* ── Form event handlers ────────────────────────────────── */
function _inspOnTypeChange() {
  const type = document.getElementById('_ispFType')?.value || 'OQC';
  _inspSetDefaultPoints(type);
  _inspRenderCpTable();
}

function _inspOnOrderChange() {
  const sel = document.getElementById('_ispFOrder');
  if (!sel) return;
  const opt = sel.options[sel.selectedIndex];
  if (!opt || !opt.value) return;
  const wt  = opt.dataset.wt  || '';
  const od  = opt.dataset.od  || '';
  const id_ = opt.dataset.id  || '';
  const h   = opt.dataset.h   || '';
  const spec = od ? ('OD'+od+(id_?'xID'+id_:'')+(h?'xH'+h:'')+'mm') : '';
  const s = (id, v) => { const el=document.getElementById(id); if(el) el.textContent=v||'—'; };
  s('_ispFWt', wt); s('_ispFOd', od?od+' mm':''); s('_ispFId', id_?id_+' mm':''); s('_ispFH', h?h+' mm':'');
  s('_ispFSpec', spec||wt||'');
  // Auto-fill tolerances into OD/ID/H check points
  _inspPoints.forEach(p => {
    if (p.cp==='OD' && od)  { const n=parseFloat(od);  if(!isNaN(n)){p.smin=(n-0.2).toFixed(1); p.smax=(n+0.2).toFixed(1);} }
    if (p.cp==='ID' && id_) { const n=parseFloat(id_); if(!isNaN(n)){p.smin=(n-0.2).toFixed(1); p.smax=(n+0.2).toFixed(1);} }
    if (p.cp==='H'  && h)   { const n=parseFloat(h);   if(!isNaN(n)){p.smin=(n-0.5).toFixed(1); p.smax=(n+0.5).toFixed(1);} }
  });
  _inspRenderCpTable();
}

function _inspSetDefaultPoints(type) {
  _inspPoints = (_INSP_DEF[type] || _INSP_DEF.OQC).map(p => ({ ...p }));
}

function _inspSelectResult(v) {
  ['OK','Hold','NG'].forEach(r => {
    const btn = document.getElementById('_ispR'+r);
    if (btn) btn.className = 'isp-rb' + (r===v?' sel-'+r.toLowerCase():'');
  });
}

function _inspGetResult() {
  for (const v of ['OK','Hold','NG']) {
    const btn = document.getElementById('_ispR'+v);
    if (btn && btn.className.includes('sel-')) return v;
  }
  return 'OK';
}

/* ── Check-point table ──────────────────────────────────── */
function _inspRenderCpTable() {
  const wrap = document.getElementById('_ispCpWrap');
  if (!wrap) return;
  wrap.innerHTML = `
<table class="isp-cp">
  <thead><tr>
    <th style="width:22px">#</th><th>Check Point</th>
    <th style="width:68px">Min</th><th style="width:68px">Max</th>
    <th style="width:72px">ค่าจริง</th><th style="width:46px">หน่วย</th>
    <th style="width:36px">ผล</th><th style="width:24px"></th>
  </tr></thead>
  <tbody>${_inspPoints.map((p,i) => _inspCpRowHtml(i,p)).join('')}</tbody>
</table>`;
}

function _inspCpRowHtml(i, p) {
  if (p.vis) {
    return `<tr>
      <td style="color:var(--t3);font-size:.7rem">${i+1}</td>
      <td><input value="${p.cp}" oninput="_inspPoints[${i}].cp=this.value" style="min-width:110px"></td>
      <td colspan="2" style="color:var(--t3);font-size:.68rem;padding-left:6px">Visual</td>
      <td><select onchange="_inspPoints[${i}].actual=this.value;_inspRefreshOk(${i})">
        <option value="" ${!p.actual?'selected':''}>—</option>
        <option value="OK" ${p.actual==='OK'?'selected':''}>OK</option>
        <option value="NG" ${p.actual==='NG'?'selected':''}>NG</option>
      </select></td>
      <td></td>
      <td id="_ispOk${i}">${_inspOkDot(p)}</td>
      <td><button class="isp-btn-sm" style="padding:1px 5px;font-size:.62rem" onclick="_inspRemovePoint(${i})">✕</button></td>
    </tr>`;
  }
  return `<tr>
    <td style="color:var(--t3);font-size:.7rem">${i+1}</td>
    <td><input value="${p.cp}" oninput="_inspPoints[${i}].cp=this.value"></td>
    <td><input type="number" value="${p.smin}" step="0.1" oninput="_inspPoints[${i}].smin=this.value;_inspRefreshOk(${i})"></td>
    <td><input type="number" value="${p.smax}" step="0.1" oninput="_inspPoints[${i}].smax=this.value;_inspRefreshOk(${i})"></td>
    <td><input type="number" value="${p.actual||''}" step="0.01" oninput="_inspPoints[${i}].actual=this.value;_inspRefreshOk(${i})"></td>
    <td><input value="${p.unit}" oninput="_inspPoints[${i}].unit=this.value" style="width:42px"></td>
    <td id="_ispOk${i}">${_inspOkDot(p)}</td>
    <td><button class="isp-btn-sm" style="padding:1px 5px;font-size:.62rem" onclick="_inspRemovePoint(${i})">✕</button></td>
  </tr>`;
}

function _inspOkDot(p) {
  if (!p.actual && p.actual !== 0) return '<span class="isp-ok isp-ok-na">—</span>';
  let ok;
  if (p.vis) {
    ok = (p.actual === 'OK');
  } else {
    const v  = parseFloat(p.actual);
    const mn = p.smin !== '' ? parseFloat(p.smin) : null;
    const mx = p.smax !== '' ? parseFloat(p.smax) : null;
    ok = (mn===null||v>=mn) && (mx===null||v<=mx);
  }
  return ok ? '<span class="isp-ok isp-ok-y">✓</span>' : '<span class="isp-ok isp-ok-n">✗</span>';
}

function _inspRefreshOk(i) {
  const cell = document.getElementById('_ispOk'+i);
  if (cell) cell.innerHTML = _inspOkDot(_inspPoints[i]);
}

function _inspAddPoint() {
  _inspPoints.push({ cp:'', smin:'', smax:'', actual:'', unit:'', vis:false });
  _inspRenderCpTable();
}

function _inspRemovePoint(i) {
  _inspPoints.splice(i, 1);
  _inspRenderCpTable();
}

/* ── Save ───────────────────────────────────────────────── */
async function _inspSave(andPrint) {
  const url = _inspUrl();
  if (!url) { alert('กรุณาตั้งค่า Script URL ก่อน'); return; }
  // เปิด print window ก่อน fetch — ป้องกัน browser บล็อก popup
  let _printWin = null;
  if (andPrint) {
    _printWin = window.open('', '_blank');
    if (!_printWin) { alert('Browser บล็อก popup กรุณาอนุญาตก่อน'); return; }
    _printWin.document.write('<html><body style="font-family:sans-serif;padding:20px">กำลังบันทึก...</body></html>');
  }
  const type      = document.getElementById('_ispFType')?.value      || 'OQC';
  const orderId   = document.getElementById('_ispFOrder')?.value     || '';
  const dateInput = document.getElementById('_ispFDate')?.value      || '';
  const inspector = (document.getElementById('_ispFInspector')?.value||'').trim();
  const qtyCheck  = parseInt(document.getElementById('_ispFQtyC')?.value)||0;
  const qtyPass   = parseInt(document.getElementById('_ispFQtyP')?.value)||0;
  const qtyFail   = parseInt(document.getElementById('_ispFQtyF')?.value)||0;
  const result    = _inspGetResult();
  const note      = (document.getElementById('_ispFNote')?.value||'').trim();
  const wt        = (document.getElementById('_ispFWt')?.textContent||'').replace('—','').trim();
  const spec      = (document.getElementById('_ispFSpec')?.textContent||'').replace('—','').trim();
  if (!orderId)   { alert('กรุณาเลือก Order'); return; }
  if (!inspector) { alert('กรุณาระบุชื่อผู้ตรวจ'); return; }
  const items = _inspPoints.map((p,i) => {
    let ok = '';
    if (p.actual !== '' && p.actual !== undefined) {
      if (p.vis) { ok = p.actual==='OK' ? 'Y' : 'N'; }
      else {
        const v  = parseFloat(p.actual);
        const mn = p.smin!=='' ? parseFloat(p.smin) : null;
        const mx = p.smax!=='' ? parseFloat(p.smax) : null;
        ok = ((mn===null||v>=mn)&&(mx===null||v<=mx)) ? 'Y' : 'N';
      }
    }
    return { seq:i+1, checkPoint:p.cp, specMin:p.smin, specMax:p.smax,
             actual:p.actual, unit:p.unit, ok, defectType:'' };
  });
  const payload = {
    action:'saveInspection', type, orderId,
    orderRef: spec||wt,
    date: _inspInputToSheet(dateInput),
    inspector, workType:wt, spec,
    qtyCheck, qtyPass, qtyFail, result, note, items,
  };
  try {
    const res = await fetch(url, { method:'POST', mode:'cors', body:JSON.stringify(payload) });
    const d   = await res.json();
    if (d.status !== 'ok') throw new Error(d.message||'error');
    _inspClosePanel();
    await _inspLoad();
    if (andPrint && _printWin) {
      const row = _inspList.find(x => x.inspId === d.inspId);
      if (row) _inspPrintRow(row, items, _printWin);
      else _printWin.close();
    }
  } catch(e) {
    if (_printWin) _printWin.close();
    alert('บันทึกไม่ได้: ' + e.message);
  }
}

/* ── View detail ────────────────────────────────────────── */
/* helper — render panel body จาก header + items (items=null = loading) */
function _inspRenderPanelBody(h, items) {
  const itemsSection = items === null
    ? `<div class="isp-sep">จุดตรวจ</div>
       <div style="text-align:center;padding:20px;color:var(--t3)">⏳ กำลังโหลดจุดตรวจ...</div>`
    : `<div class="isp-sep">จุดตรวจ</div>
<table class="isp-cp">
  <thead><tr><th>#</th><th>Check Point</th><th>Min</th><th>Max</th><th>ค่าจริง</th><th>หน่วย</th><th>ผล</th></tr></thead>
  <tbody>${items.map((it,i)=>`<tr>
    <td style="color:var(--t3)">${i+1}</td>
    <td>${it.checkPoint||'—'}</td>
    <td style="text-align:center">${it.specMin||'—'}</td>
    <td style="text-align:center">${it.specMax||'—'}</td>
    <td style="text-align:center"><strong>${it.actual||'—'}</strong></td>
    <td>${it.unit||''}</td>
    <td>${it.ok==='Y'?'<span class="isp-ok isp-ok-y">✓</span>':it.ok==='N'?'<span class="isp-ok isp-ok-n">✗</span>':'<span class="isp-ok isp-ok-na">—</span>'}</td>
  </tr>`).join('')}</tbody>
</table>
${h.note?`<div style="font-size:.8rem;color:var(--t2);margin-top:4px"><strong>หมายเหตุ:</strong> ${h.note}</div>`:''}`;

  document.getElementById('_ispPanelTitle').textContent = h.inspId || '';
  document.getElementById('_ispPanelBody').innerHTML = `
<div class="isp-info-box">
  <div class="isp-info-grid" style="grid-template-columns:1fr 1fr">
    <span><span class="ik">ประเภท: </span>${_inspTypeBadge(h.type)}</span>
    <span><span class="ik">No.PO: </span><strong>${h.orderId||'—'}</strong></span>
    <span><span class="ik">วันที่ตรวจ: </span>${h.date||'—'}</span>
    <span><span class="ik">ผู้ตรวจ: </span><strong>${h.inspector||'—'}</strong></span>
    <span style="grid-column:1/-1"><span class="ik">สินค้า / Spec: </span>${h.orderRef||h.workType||'—'}</span>
  </div>
</div>
<div class="isp-stat-row" style="grid-template-columns:1fr 1fr 1fr">
  <div class="isp-stat"><div class="isp-stat-lbl">ตรวจ</div><div class="isp-stat-val isp-c-blue">${h.qtyCheck||0}</div></div>
  <div class="isp-stat"><div class="isp-stat-lbl">ผ่าน</div><div class="isp-stat-val isp-c-pass">${h.qtyPass||0}</div></div>
  <div class="isp-stat"><div class="isp-stat-lbl">ไม่ผ่าน</div><div class="isp-stat-val isp-c-fail">${h.qtyFail||0}</div></div>
</div>
<div style="text-align:center;margin-bottom:12px">${_inspBadge(h.result)}</div>
${itemsSection}`;

  // footer — ถ้ายัง loading items ให้ disable print
  if (items === null) {
    document.getElementById('_ispPanelFooter').innerHTML = `
      <button class="isp-btn-out" onclick="_inspClosePanel()">ปิด</button>
      <button class="isp-btn-pri" disabled style="opacity:.5">⏳ โหลด...</button>`;
  } else {
    document.getElementById('_ispPanelFooter').innerHTML = `
      <button class="isp-btn-out" onclick="_inspClosePanel()">ปิด</button>
      <button class="isp-btn-pri" onclick="_inspPrintRow(${JSON.stringify(h)},${JSON.stringify(items)})">🖨 พิมพ์</button>`;
  }
}

async function _inspViewDetail(inspId) {
  // แสดง panel ทันทีจาก cache — ไม่ต้องรอ fetch
  const hCached = _inspList.find(x => x.inspId === inspId);
  if (hCached) {
    _inspRenderPanelBody(hCached, null);        // แสดงทันที, items = loading
    document.getElementById('_ispPanel').style.display = 'block';
  }

  // โหลด items จาก server ใน background
  const url = _inspUrl();
  if (!url) return;
  try {
    const r = await fetch(url + '?action=getInspectionDetail&inspId=' + encodeURIComponent(inspId));
    const d = await r.json();
    if (d.status !== 'ok' || !d.header) {
      if (!hCached) { alert('ไม่พบข้อมูล'); return; }
      // มี cache อยู่แล้ว — แสดงว่าโหลด items ไม่ได้ แต่ไม่ปิด panel
      document.getElementById('_ispPanelBody').querySelector('.isp-sep') &&
        (_inspRenderPanelBody(hCached, []));
      return;
    }
    _inspRenderPanelBody(d.header, d.items || []);
    if (!hCached) document.getElementById('_ispPanel').style.display = 'block';
  } catch(e) {
    if (!hCached) alert('โหลดไม่ได้: ' + e.message);
  }
}

/* ── Print by ID ────────────────────────────────────────── */
async function _inspPrintById(inspId) {
  const url = _inspUrl();
  if (!url) return;
  // เปิด window ก่อน fetch — ต้องอยู่ใน user gesture context
  const win = window.open('', '_blank');
  if (!win) { alert('Browser บล็อก popup กรุณาอนุญาตก่อน'); return; }
  win.document.write('<html><body style="font-family:sans-serif;padding:20px">กำลังโหลดข้อมูล...</body></html>');
  try {
    const r = await fetch(url + '?action=getInspectionDetail&inspId=' + encodeURIComponent(inspId));
    const d = await r.json();
    if (d.status !== 'ok' || !d.header) { win.close(); alert('ไม่พบข้อมูล'); return; }
    _inspPrintRow(d.header, d.items || [], win);
  } catch(e) { win.close(); alert('พิมพ์ไม่ได้: ' + e.message); }
}

/* ── Print document ─────────────────────────────────────── */
function _inspPrintRow(h, items, win) {
  const co = (() => { try { return JSON.parse(localStorage.getItem('ptts_company')||'{}'); } catch(e){ return {}; } })();
  const coName = co.name || 'บริษัท PTTS';
  const typeLabel = h.type === 'IPQC' ? 'IPQC — ใบตรวจสอบระหว่างผลิต' : 'OQC — ใบตรวจสอบก่อนส่งมอบ';
  const rc = h.result==='OK' ? '#16a34a' : h.result==='NG' ? '#dc2626' : '#d97706';
  const rcBg = h.result==='OK' ? '#f0fdf4' : h.result==='NG' ? '#fef2f2' : '#fefce8';
  const cpRows = (items||[]).map((it,i) => {
    const okHtml = it.ok==='Y' ? '<span style="color:#16a34a;font-weight:700">✓ OK</span>'
                 : it.ok==='N' ? '<span style="color:#dc2626;font-weight:700">✗ NG</span>' : '—';
    return `<tr>
      <td style="text-align:center">${i+1}</td>
      <td>${it.checkPoint||''}</td>
      <td style="text-align:center">${it.specMin||'—'}</td>
      <td style="text-align:center">${it.specMax||'—'}</td>
      <td style="text-align:center"><strong>${it.actual||'—'}</strong></td>
      <td style="text-align:center">${it.unit||''}</td>
      <td style="text-align:center">${okHtml}</td>
    </tr>`;
  }).join('');
  const html = `<!DOCTYPE html><html lang="th"><head>
<meta charset="UTF-8"><title>${h.inspId||'ใบตรวจ'}</title>
<link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Sarabun',sans-serif;font-size:13px;color:#1e293b;background:#fff;padding:24px 28px}
.wrap{max-width:720px;margin:0 auto}
.hdr{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:14px;padding-bottom:12px;border-bottom:2px solid #1e293b}
.co-name{font-size:1rem;font-weight:800}.co-sub{font-size:.7rem;color:#64748b;margin-top:2px}
.doc-type{font-size:.65rem;color:#64748b;font-weight:700;text-transform:uppercase;letter-spacing:.04em;text-align:right}
.doc-id{font-size:1.15rem;font-weight:800;text-align:right}
.info-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:11px 14px;margin-bottom:12px;font-size:.8rem}
.ik{color:#64748b;font-size:.7rem;display:block;margin-bottom:1px}
.summary{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:12px;text-align:center}
.sb{border:1px solid #e2e8f0;border-radius:8px;padding:8px}.sb .sk{font-size:.68rem;color:#64748b}.sb .sv{font-size:1.3rem;font-weight:800}
.res-box{border:2px solid ${rc};border-radius:10px;background:${rcBg};text-align:center;padding:8px;margin-bottom:14px}
.res-box .rl{font-size:.68rem;font-weight:700;color:${rc}}.res-box .rv{font-size:1.6rem;font-weight:800;color:${rc}}
table.cp{width:100%;border-collapse:collapse;font-size:.8rem;margin-bottom:14px}
table.cp th{background:#1e293b;color:#fff;padding:7px 10px;font-size:.7rem;font-weight:700}
table.cp th:nth-child(2){text-align:left}
table.cp td{padding:7px 10px;border-bottom:1px solid #e2e8f0;vertical-align:middle;text-align:center}
table.cp td:nth-child(2){text-align:left}
.note{border:1px solid #e2e8f0;border-radius:8px;padding:10px;font-size:.8rem;margin-bottom:14px}
.sign-row{display:grid;grid-template-columns:1fr 1fr 1fr;gap:20px;margin-top:28px}
.sign-box{border-top:1px solid #94a3b8;padding-top:6px;text-align:center;font-size:.7rem;color:#64748b}
@media print{body{padding:10px}.wrap{max-width:100%}}
</style></head><body>
<div class="wrap">
  <div class="hdr">
    <div><div class="co-name">${coName}</div>${co.address?`<div class="co-sub">${co.address}</div>`:''}</div>
    <div><div class="doc-type">${typeLabel}</div><div class="doc-id">${h.inspId||'—'}</div></div>
  </div>
  <div class="info-grid">
    <div><span class="ik">No.PO / Order</span><strong>${h.orderId||'—'}</strong></div>
    <div><span class="ik">วันที่ตรวจ</span>${h.date||'—'}</div>
    <div><span class="ik">สินค้า / Spec</span>${h.orderRef||h.workType||'—'}</div>
    <div><span class="ik">ผู้ตรวจ</span><strong>${h.inspector||'—'}</strong></div>
  </div>
  <div class="summary">
    <div class="sb"><div class="sk">จำนวนตรวจ</div><div class="sv" style="color:#2563eb">${h.qtyCheck||0}</div></div>
    <div class="sb"><div class="sk">ผ่าน</div><div class="sv" style="color:#16a34a">${h.qtyPass||0}</div></div>
    <div class="sb"><div class="sk">ไม่ผ่าน</div><div class="sv" style="color:#dc2626">${h.qtyFail||0}</div></div>
  </div>
  <div class="res-box"><div class="rl">ผลการตรวจ</div><div class="rv">${h.result||'—'}</div></div>
  <table class="cp">
    <thead><tr><th style="width:30px">#</th><th>Check Point</th><th style="width:65px">Min</th><th style="width:65px">Max</th><th style="width:75px">ค่าจริง</th><th style="width:50px">หน่วย</th><th style="width:65px">ผล</th></tr></thead>
    <tbody>${cpRows}</tbody>
  </table>
  ${h.note?`<div class="note"><strong>หมายเหตุ:</strong> ${h.note}</div>`:''}
  <div class="sign-row">
    <div class="sign-box">ผู้ตรวจ<br><br><br>${h.inspector||'.................................'}</div>
    <div class="sign-box">หัวหน้าฝ่ายผลิต<br><br><br>.................................</div>
    <div class="sign-box">QC อนุมัติ<br><br><br>.................................</div>
  </div>
</div>
<script>(function(){
  function go(){ try{ window.focus(); window.print(); }catch(e){} }
  if(document.fonts&&document.fonts.ready){ document.fonts.ready.then(go).catch(function(){ setTimeout(go,700); }); }
  else { setTimeout(go,700); }
})();<\/script>
</body></html>`;
  if (!win) {
    win = window.open('','_blank');
    if (!win) { alert('Browser บล็อก popup กรุณาอนุญาตก่อน'); return; }
  }
  win.document.open();
  win.document.write(html);
  win.document.close();
}

/* ── Delete inspection ──────────────────────────────────── */
async function _inspDelete(inspId) {
  const confirmed = await Swal.fire({
    title: 'ลบใบตรวจ?',
    html: `<div style="font-size:.9rem">เลขที่ <strong>${inspId}</strong><br><span style="color:#dc2626">ไม่สามารถกู้คืนได้</span></div>`,
    icon: 'warning',
    showCancelButton: true,
    confirmButtonColor: '#dc2626',
    cancelButtonColor: '#6b7280',
    confirmButtonText: 'ลบเลย',
    cancelButtonText: 'ยกเลิก',
  });
  if (!confirmed.isConfirmed) return;
  const url = _inspUrl();
  if (!url) return;
  try {
    const r = await fetch(url, {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ action:'deleteInspection', inspId })
    });
    const d = await r.json();
    if (d.status !== 'ok') { Swal.fire('ผิดพลาด','ลบไม่สำเร็จ: ' + (d.message||''),'error'); return; }
    Swal.fire({ title:'ลบแล้ว', icon:'success', timer:1200, showConfirmButton:false });
    _inspList = _inspList.filter(x => x.inspId !== inspId);
    _inspRenderList();
  } catch(e) { Swal.fire('Error', e.message, 'error'); }
}

/* ── Edit inspection: load data into form ───────────────── */
async function _inspEditLoad(inspId) {
  // ดึง header จาก cache ก่อน
  const hc = _inspList.find(x => x.inspId === inspId);
  // โหลด detail (ต้องการ items)
  const url = _inspUrl();
  if (!url) return;
  Swal.fire({ title:'กำลังโหลด...', didOpen:()=>Swal.showLoading(), allowOutsideClick:false });
  try {
    const r = await fetch(url + '?action=getInspectionDetail&inspId=' + encodeURIComponent(inspId));
    const d = await r.json();
    Swal.close();
    if (d.status !== 'ok' || !d.header) { Swal.fire('ผิดพลาด','ไม่พบข้อมูล','error'); return; }
    const h = d.header; const items = d.items || [];

    // เปิด panel ในโหมดสร้าง แล้วค่อย prefill
    _inspEditId = inspId;
    _inspOpenPanel();

    // prefill fields
    document.getElementById('_ispFType').value      = h.type      || 'OQC';
    document.getElementById('_ispFOrderId').value   = h.orderId   || '';
    document.getElementById('_ispFOrderRef').value  = h.orderRef  || '';
    document.getElementById('_ispFDate').value      = (() => {
      // convert dd/MM/yyyy (BE) → yyyy-MM-dd (CE) for input[type=date]
      const p = (h.date||'').split('/');
      if (p.length === 3) {
        const yr = parseInt(p[2]) > 2400 ? parseInt(p[2])-543 : parseInt(p[2]);
        return `${yr}-${p[1].padStart(2,'0')}-${p[0].padStart(2,'0')}`;
      }
      return '';
    })();
    document.getElementById('_ispFInspector').value = h.inspector || '';
    document.getElementById('_ispFQtyC').value      = h.qtyCheck  || '';
    document.getElementById('_ispFQtyP').value      = h.qtyPass   || '';
    document.getElementById('_ispFQtyF').value      = h.qtyFail   || '';
    document.getElementById('_ispFResult').value    = h.result    || 'OK';
    const noteEl = document.getElementById('_ispFNote');
    if (noteEl) noteEl.value = h.note || '';

    // load check points
    _inspPoints = items.map(it => ({
      checkPoint: it.checkPoint || '',
      specMin:    it.specMin    || '',
      specMax:    it.specMax    || '',
      actual:     it.actual     || '',
      unit:       it.unit       || '',
      ok:         it.ok         || '',
      defectType: it.defectType || '',
    }));
    _inspRenderPoints();

    // อัปเดต panel title และ trigger type change
    document.getElementById('_ispPanelTitle').textContent = '✏️ แก้ไขใบตรวจ ' + inspId;
    if (typeof _inspOnTypeChange === 'function') _inspOnTypeChange();
  } catch(e) { Swal.close(); Swal.fire('Error', e.message, 'error'); }
}

