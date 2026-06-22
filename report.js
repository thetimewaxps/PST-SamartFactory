// ══════════════════════════════════════════════════════════════
// report.js — Report Hub (tab-report)   v1.0
// ══════════════════════════════════════════════════════════════

// ── caches (reset ได้ด้วย rptClearCache) ──────────────────────
let _rptInvoices  = null;
let _rptBills     = null;
let _rptPOs       = null;
let _rptPlates    = null;
let _rptCustomers = null;
let _rptSuppliers = null;

function rptClearCache() {
  _rptInvoices = _rptBills = _rptPOs = _rptPlates = _rptCustomers = _rptSuppliers = null;
}

// ── utils ──────────────────────────────────────────────────────
function _rptUrl() { return localStorage.getItem('ptts_script_url') || ''; }

function _rptBaht(n) {
  return Number(n || 0).toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// 'dd/MM/yyyy' (CE หรือ BE) → 'YYYY-MM' CE เพื่อ sort/group
function _rptMonthKey(dateStr) {
  var m = String(dateStr || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return '';
  var yr = parseInt(m[3]);
  if (yr > 2500) yr -= 543;
  return yr + '-' + String(m[2]).padStart(2, '0');
}

// 'YYYY-MM' CE → 'ม.ค. 2569' (พ.ศ.)
function _rptMonthLabel(key) {
  var MO = ['','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  if (!key) return '';
  var p = key.split('-');
  return MO[parseInt(p[1])] + ' ' + (parseInt(p[0]) + 543);
}

// สร้าง dropdown html จาก list ของ monthKey
function _rptMonthSelect(id, months, selected) {
  return '<select id="' + id + '" onchange="' + id + 'Changed()" style="border:1px solid var(--bc-div);border-radius:8px;padding:6px 10px;font-family:Sarabun,sans-serif;font-size:.85rem;color:var(--t1);background:var(--bg-card);cursor:pointer">' +
    months.map(function(k) {
      return '<option value="' + k + '"' + (k === selected ? ' selected' : '') + '>' + _rptMonthLabel(k) + '</option>';
    }).join('') +
  '</select>';
}

// สร้าง print window (pattern มาตรฐาน)
function _rptOpenPrint(html) {
  var win = window.open('', '_blank');
  if (!win) { alert('กรุณาอนุญาต popup'); return; }
  win.document.write(html);
  win.document.close();
}

// ── โหลดข้อมูล (lazy + cache) ─────────────────────────────────
async function _rptFetch(action, key) {
  var url = _rptUrl();
  if (!url) throw new Error('ยังไม่ได้ตั้งค่า Script URL');
  var res = await fetch(url + '?action=' + action, { mode: 'cors' });
  var d   = await res.json();
  if (d.status === 'error') throw new Error(d.message);
  return d[key] || [];
}

async function _rptGetInvoices()  { if (!_rptInvoices)  _rptInvoices  = await _rptFetch('getInvoices',   'invoices');  return _rptInvoices; }
async function _rptGetBills()     { if (!_rptBills)     _rptBills     = await _rptFetch('getBillingNotes','bills');     return _rptBills; }
async function _rptGetPOs()       { if (!_rptPOs)       _rptPOs       = (await _rptFetch('getPurchaseOrders','headers')) ; return _rptPOs; }
async function _rptGetPlates()    { if (!_rptPlates)    _rptPlates    = await _rptFetch('getPlatingNotes','plates');    return _rptPlates; }
async function _rptGetCustomers() { if (!_rptCustomers) _rptCustomers = await _rptFetch('getCustomers',  'customers'); return _rptCustomers; }
async function _rptGetSuppliers() { if (!_rptSuppliers) _rptSuppliers = await _rptFetch('getSuppliers',  'suppliers'); return _rptSuppliers; }

// ── shared print CSS ───────────────────────────────────────────
function _rptPrintCss() {
  return '<link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700;800&display=swap" rel="stylesheet">' +
  '<style>*{box-sizing:border-box;margin:0;padding:0}' +
  'body{font-family:Sarabun,sans-serif;font-size:11.5px;color:#1e293b;padding:14px 18px}' +
  '.hdr{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:3px solid #1e293b;padding-bottom:8px;margin-bottom:12px}' +
  '.hdr-title{font-size:1rem;font-weight:900;color:#64748b}' +
  '.hdr-main{font-size:1.4rem;font-weight:900;color:#1d4ed8}' +
  '.hdr-sub{font-size:.78rem;font-weight:700;color:#7c3aed;margin-top:2px}' +
  '.hdr-right{text-align:right;font-size:.68rem;color:#64748b;line-height:1.8}' +
  'table{width:100%;border-collapse:collapse;font-size:.79rem}' +
  'thead tr{background:#1e293b}' +
  'th{color:#fff;padding:6px 8px;font-size:.68rem;font-weight:700;text-align:left}' +
  'th.r{text-align:right}th.c{text-align:center}' +
  'td{padding:5px 8px;border-bottom:1px solid #e2e8f0;vertical-align:middle}' +
  'td.r{text-align:right}td.c{text-align:center}' +
  '.sum-row td{font-weight:800;background:#f1f5f9;border-top:2px solid #94a3b8}' +
  '.badge{display:inline-block;font-size:.7rem;font-weight:700;padding:2px 8px;border-radius:5px}' +
  '.footer{margin-top:16px;font-size:.68rem;color:#94a3b8;text-align:right}' +
  '@media print{body{padding:6px 10px}@page{size:A4 portrait;margin:10mm}}' +
  '</style>';
}

function _rptPrintScript() {
  return '<script>(function(){if(document.fonts&&document.fonts.ready){document.fonts.ready.then(function(){window.print()}).catch(function(){setTimeout(function(){window.print()},700)})}else{setTimeout(function(){window.print()},700)}})()\x3c/script>';
}

// ══════════════════════════════════════════════════════════════
// REPORT 1 — ยอดขายรายเดือน
// ══════════════════════════════════════════════════════════════
var _rpt1MonthList = [];
var _rpt1SelMonth  = '';
var _rpt1Inv       = [];
var _rpt1CustMap   = {};

async function rptSales() {
  var url = _rptUrl();
  if (!url) { Swal.fire('ยังไม่ตั้งค่า', 'กรุณาตั้งค่า Script URL ในแท็บ ตั้งค่า', 'warning'); return; }
  Swal.fire({ title: 'กำลังโหลด...', didOpen: function() { Swal.showLoading(); }, allowOutsideClick: false });
  try {
    var [invs, custs] = await Promise.all([_rptGetInvoices(), _rptGetCustomers()]);
    _rpt1Inv = invs;
    _rpt1CustMap = {};
    custs.forEach(function(c) { _rpt1CustMap[c.code] = c.name; });

    // สร้าง month list (desc)
    var mset = {};
    invs.forEach(function(iv) { var k = _rptMonthKey(iv.date); if (k) mset[k] = true; });
    _rpt1MonthList = Object.keys(mset).sort().reverse();
    if (!_rpt1MonthList.length) { Swal.fire('ไม่มีข้อมูล', 'ยังไม่มีใบกำกับภาษีในระบบ', 'info'); return; }
    _rpt1SelMonth = _rpt1MonthList[0];

    Swal.hideLoading(); Swal.close();
    _rpt1ShowSwal();
  } catch (e) {
    Swal.fire('โหลดไม่ได้', e.message, 'error');
  }
}

function _rpt1ShowSwal() {
  window.rpt1MonthChanged = function() {
    var sel = document.getElementById('rpt1MonthSel');
    if (sel) { _rpt1SelMonth = sel.value; var c = Swal.getHtmlContainer(); if (c) c.innerHTML = _rpt1Html(); }
  };
  Swal.fire({
    title: '💰 ยอดขายรายเดือน',
    background: 'var(--bg-card)', color: 'var(--t1)', width: 760,
    html: _rpt1Html(),
    showCancelButton: true,
    cancelButtonText: '🖨 พิมพ์',
    cancelButtonColor: '#0369a1',
    confirmButtonText: 'ปิด',
    confirmButtonColor: '#2563eb',
    reverseButtons: true,
  }).then(function(res) {
    delete window.rpt1MonthChanged;
    if (res.dismiss === Swal.DismissReason.cancel) _rpt1Print();
  });
}

function _rpt1Html() {
  var filtered = _rpt1Inv.filter(function(iv) { return _rptMonthKey(iv.date) === _rpt1SelMonth; });
  var sumSub = 0, sumVat = 0, sumTotal = 0;
  filtered.forEach(function(iv) { sumSub += iv.subtotal; sumVat += iv.vat; sumTotal += iv.total; });

  var rows = filtered.map(function(iv, i) {
    var bg = i % 2 === 0 ? '' : 'background:rgba(0,0,0,.02)';
    var custName = _rpt1CustMap[iv.customerCode] || iv.customerCode;
    return '<tr style="' + bg + '">' +
      '<td style="font-size:.72rem;color:var(--t3)">' + iv.invoiceNo + '</td>' +
      '<td style="font-size:.72rem">' + iv.date + '</td>' +
      '<td style="font-size:.78rem;font-weight:600">' + custName + '</td>' +
      '<td style="text-align:right">' + _rptBaht(iv.subtotal) + '</td>' +
      '<td style="text-align:right;color:#64748b">' + _rptBaht(iv.vat) + '</td>' +
      '<td style="text-align:right;font-weight:700;color:var(--c1)">' + _rptBaht(iv.total) + '</td>' +
    '</tr>';
  }).join('');

  return '<div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;flex-wrap:wrap">' +
    '<span style="font-size:.85rem;color:var(--t2)">เดือน:</span>' +
    _rptMonthSelect('rpt1MonthSel', _rpt1MonthList, _rpt1SelMonth) +
    '<span style="font-size:.8rem;color:var(--t3)">' + filtered.length + ' รายการ</span>' +
  '</div>' +
  (filtered.length
    ? '<div style="overflow-x:auto;max-height:380px;overflow-y:auto">' +
      '<table style="width:100%;border-collapse:collapse;font-size:.79rem">' +
      '<thead><tr style="background:rgba(37,99,235,.1);position:sticky;top:0">' +
        '<th style="padding:6px 8px;text-align:left;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">เลขที่</th>' +
        '<th style="padding:6px 8px;text-align:left;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">วันที่</th>' +
        '<th style="padding:6px 8px;text-align:left;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">ลูกค้า</th>' +
        '<th style="padding:6px 8px;text-align:right;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">ยอดก่อน VAT</th>' +
        '<th style="padding:6px 8px;text-align:right;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">VAT</th>' +
        '<th style="padding:6px 8px;text-align:right;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">รวม</th>' +
      '</tr></thead>' +
      '<tbody>' + rows + '</tbody>' +
      '<tfoot><tr style="background:rgba(37,99,235,.07);font-weight:800">' +
        '<td colspan="3" style="padding:6px 8px;font-size:.78rem;color:var(--t1)">รวมทั้งเดือน ' + _rptMonthLabel(_rpt1SelMonth) + '</td>' +
        '<td style="padding:6px 8px;text-align:right;color:var(--c1)">' + _rptBaht(sumSub) + '</td>' +
        '<td style="padding:6px 8px;text-align:right;color:#64748b">' + _rptBaht(sumVat) + '</td>' +
        '<td style="padding:6px 8px;text-align:right;color:#16a34a;font-size:.88rem">' + _rptBaht(sumTotal) + '</td>' +
      '</tr></tfoot>' +
      '</table></div>'
    : '<div style="padding:32px;text-align:center;color:var(--t3)">ไม่มีรายการในเดือนนี้</div>');
}

function _rpt1Print() {
  var filtered = _rpt1Inv.filter(function(iv) { return _rptMonthKey(iv.date) === _rpt1SelMonth; });
  var sumSub = 0, sumVat = 0, sumTotal = 0;
  filtered.forEach(function(iv) { sumSub += iv.subtotal; sumVat += iv.vat; sumTotal += iv.total; });
  var rows = filtered.map(function(iv, i) {
    var bg = i % 2 === 0 ? '#f8fafc' : '#fff';
    return '<tr style="background:' + bg + '">' +
      '<td class="c" style="color:#64748b;font-size:.7rem">' + (i+1) + '</td>' +
      '<td style="font-size:.7rem;color:#64748b">' + iv.invoiceNo + '</td>' +
      '<td>' + iv.date + '</td>' +
      '<td>' + (_rpt1CustMap[iv.customerCode] || iv.customerCode) + '</td>' +
      '<td class="r">' + _rptBaht(iv.subtotal) + '</td>' +
      '<td class="r" style="color:#64748b">' + _rptBaht(iv.vat) + '</td>' +
      '<td class="r" style="font-weight:700;color:#1d4ed8">' + _rptBaht(iv.total) + '</td>' +
    '</tr>';
  }).join('');
  var html = '<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8"><title>ยอดขาย ' + _rptMonthLabel(_rpt1SelMonth) + '</title>' +
    _rptPrintCss() + '</head><body>' +
    '<div class="hdr"><div>' +
      '<div class="hdr-title">รายงานยอดขาย</div>' +
      '<div class="hdr-main">เดือน ' + _rptMonthLabel(_rpt1SelMonth) + '</div>' +
    '</div><div class="hdr-right">จำนวน ' + filtered.length + ' รายการ<br>พิมพ์เมื่อ: ' + new Date().toLocaleDateString('th-TH') + '</div></div>' +
    '<table><thead><tr>' +
      '<th class="c" style="width:28px">#</th>' +
      '<th style="width:90px">เลขที่</th>' +
      '<th style="width:90px">วันที่</th>' +
      '<th>ลูกค้า</th>' +
      '<th class="r" style="width:90px">ก่อน VAT</th>' +
      '<th class="r" style="width:70px">VAT</th>' +
      '<th class="r" style="width:90px">รวม</th>' +
    '</tr></thead><tbody>' + rows + '</tbody>' +
    '<tfoot><tr class="sum-row">' +
      '<td colspan="4" style="padding:6px 8px">รวมทั้งเดือน</td>' +
      '<td class="r">' + _rptBaht(sumSub) + '</td>' +
      '<td class="r">' + _rptBaht(sumVat) + '</td>' +
      '<td class="r" style="color:#1d4ed8">' + _rptBaht(sumTotal) + '</td>' +
    '</tr></tfoot></table>' +
    '<div class="footer">รายงานนี้สร้างโดยระบบ PTTS Cost Breakdown</div>' +
    _rptPrintScript() + '</body></html>';
  _rptOpenPrint(html);
}

// ══════════════════════════════════════════════════════════════
// REPORT 2 — ค้างชำระ (Outstanding Billing)
// ══════════════════════════════════════════════════════════════
var _rpt2Bills   = [];
var _rpt2CustMap = {};

async function rptOutstanding() {
  var url = _rptUrl();
  if (!url) { Swal.fire('ยังไม่ตั้งค่า', 'กรุณาตั้งค่า Script URL ในแท็บ ตั้งค่า', 'warning'); return; }
  Swal.fire({ title: 'กำลังโหลด...', didOpen: function() { Swal.showLoading(); }, allowOutsideClick: false });
  try {
    var [bills, custs] = await Promise.all([_rptGetBills(), _rptGetCustomers()]);
    _rpt2Bills = bills;
    _rpt2CustMap = {};
    custs.forEach(function(c) { _rpt2CustMap[c.code] = c.name; });
    Swal.hideLoading(); Swal.close();
    _rpt2ShowSwal();
  } catch (e) {
    Swal.fire('โหลดไม่ได้', e.message, 'error');
  }
}

// แปลง payTerm string → จำนวนวัน (เพื่อคำนวณ due date)
function _rpt2ParseDays(payTerm) {
  if (!payTerm) return 0;
  var s = String(payTerm).toLowerCase();
  if (s === 'cod' || s === 'cash') return 0;
  var m = s.match(/(\d+)/);
  return m ? parseInt(m[1]) : 30;
}

// 'dd/MM/yyyy' (BE) → Date object
function _rpt2ParseDate(dateStr) {
  var m = String(dateStr || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return null;
  var yr = parseInt(m[3]);
  if (yr > 2500) yr -= 543;
  return new Date(yr, parseInt(m[2]) - 1, parseInt(m[1]));
}

function _rpt2ShowSwal() {
  var today = new Date(); today.setHours(0,0,0,0);

  // เตรียม + sort by overdue ก่อน
  var enriched = _rpt2Bills.map(function(b) {
    var d = _rpt2ParseDate(b.date);
    var days = _rpt2ParseDays(b.payTerm);
    var due = d ? new Date(d.getTime() + days * 86400000) : null;
    var overdueDays = due ? Math.floor((today - due) / 86400000) : 0;
    return Object.assign({}, b, { _due: due, _overdueDays: overdueDays });
  }).sort(function(a, b) { return b._overdueDays - a._overdueDays; });

  var html = _rpt2Html(enriched, today);

  Swal.fire({
    title: '🧾 ใบวางบิลทั้งหมด',
    background: 'var(--bg-card)', color: 'var(--t1)', width: 780,
    html: html,
    showCancelButton: enriched.length > 0,
    cancelButtonText: '🖨 พิมพ์',
    cancelButtonColor: '#0369a1',
    confirmButtonText: 'ปิด',
    confirmButtonColor: '#2563eb',
    reverseButtons: true,
  }).then(function(res) {
    if (res.dismiss === Swal.DismissReason.cancel) _rpt2Print(enriched);
  });
}

function _rpt2Html(enriched, today) {
  var sumNet = 0;
  enriched.forEach(function(b) { sumNet += b.netAmount; });

  var rows = enriched.map(function(b, i) {
    var custName = _rpt2CustMap[b.customerCode] || b.customerCode;
    var dueStr = b._due ? b._due.toLocaleDateString('th-TH', { year: 'numeric', month: 'short', day: 'numeric' }) : '—';
    var od = b._overdueDays;
    var odColor = od > 30 ? '#dc2626' : od > 0 ? '#d97706' : '#16a34a';
    var odLabel = od > 0 ? 'เกิน ' + od + ' วัน' : od === 0 ? 'ครบวันนี้' : 'เหลือ ' + (-od) + ' วัน';
    var bg = i % 2 === 0 ? '' : 'background:rgba(0,0,0,.02)';
    return '<tr style="' + bg + '">' +
      '<td style="font-size:.72rem;color:var(--t3);white-space:nowrap">' + b.billNo + '</td>' +
      '<td style="font-size:.72rem">' + b.date + '</td>' +
      '<td style="font-size:.78rem;font-weight:600">' + custName + '</td>' +
      '<td style="text-align:right">' + _rptBaht(b.netAmount) + '</td>' +
      '<td style="font-size:.72rem;text-align:center;white-space:nowrap">' + (b.payTerm || '—') + '</td>' +
      '<td style="font-size:.72rem;text-align:center;white-space:nowrap">' + dueStr + '</td>' +
      '<td style="text-align:center"><span style="font-size:.7rem;font-weight:700;color:' + odColor + ';background:' + odColor.replace('#','rgba(') + ',.1);padding:2px 8px;border-radius:5px">' + odLabel + '</span></td>' +
    '</tr>';
  }).join('');

  return (enriched.length
    ? '<div style="margin-bottom:8px;font-size:.78rem;color:var(--t3)">รวม ' + enriched.length + ' ใบ · ยอดสุทธิรวม <strong style="color:var(--c1)">' + _rptBaht(sumNet) + '</strong> บาท</div>' +
      '<div style="overflow-x:auto;max-height:380px;overflow-y:auto">' +
      '<table style="width:100%;border-collapse:collapse;font-size:.79rem">' +
      '<thead><tr style="background:rgba(37,99,235,.1);position:sticky;top:0">' +
        '<th style="padding:6px 8px;text-align:left;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">เลขที่</th>' +
        '<th style="padding:6px 8px;text-align:left;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">วันที่</th>' +
        '<th style="padding:6px 8px;text-align:left;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">ลูกค้า</th>' +
        '<th style="padding:6px 8px;text-align:right;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">ยอดสุทธิ</th>' +
        '<th style="padding:6px 8px;text-align:center;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">เครดิต</th>' +
        '<th style="padding:6px 8px;text-align:center;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">ครบกำหนด</th>' +
        '<th style="padding:6px 8px;text-align:center;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">สถานะ</th>' +
      '</tr></thead><tbody>' + rows + '</tbody></table></div>'
    : '<div style="padding:32px;text-align:center;color:var(--t3)">ยังไม่มีใบวางบิลในระบบ</div>');
}

function _rpt2Print(enriched) {
  var today = new Date(); today.setHours(0,0,0,0);
  var sumNet = 0;
  enriched.forEach(function(b) { sumNet += b.netAmount; });
  var rows = enriched.map(function(b, i) {
    var custName = _rpt2CustMap[b.customerCode] || b.customerCode;
    var dueStr = b._due ? b._due.toLocaleDateString('th-TH', { year: 'numeric', month: 'short', day: 'numeric' }) : '—';
    var od = b._overdueDays;
    var odColor = od > 30 ? '#dc2626' : od > 0 ? '#d97706' : '#16a34a';
    var odLabel = od > 0 ? 'เกิน ' + od + ' วัน' : od === 0 ? 'ครบวันนี้' : 'เหลือ ' + (-od) + ' วัน';
    var bg = i % 2 === 0 ? '#f8fafc' : '#fff';
    return '<tr style="background:' + bg + '">' +
      '<td class="c" style="color:#64748b;font-size:.7rem">' + (i+1) + '</td>' +
      '<td>' + b.billNo + '</td>' +
      '<td>' + b.date + '</td>' +
      '<td>' + custName + '</td>' +
      '<td class="r">' + _rptBaht(b.netAmount) + '</td>' +
      '<td class="c">' + (b.payTerm || '—') + '</td>' +
      '<td class="c">' + dueStr + '</td>' +
      '<td class="c"><span style="color:' + odColor + ';font-weight:700">' + odLabel + '</span></td>' +
    '</tr>';
  }).join('');
  var html = '<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8"><title>ใบวางบิลทั้งหมด</title>' +
    _rptPrintCss() + '</head><body>' +
    '<div class="hdr"><div>' +
      '<div class="hdr-title">รายงาน</div>' +
      '<div class="hdr-main">ใบวางบิลทั้งหมด</div>' +
    '</div><div class="hdr-right">' + enriched.length + ' ใบ<br>พิมพ์เมื่อ: ' + new Date().toLocaleDateString('th-TH') + '</div></div>' +
    '<table><thead><tr>' +
      '<th class="c" style="width:28px">#</th>' +
      '<th style="width:85px">เลขที่</th>' +
      '<th style="width:80px">วันที่</th>' +
      '<th>ลูกค้า</th>' +
      '<th class="r" style="width:85px">ยอดสุทธิ</th>' +
      '<th class="c" style="width:65px">เครดิต</th>' +
      '<th class="c" style="width:95px">ครบกำหนด</th>' +
      '<th class="c" style="width:90px">สถานะ</th>' +
    '</tr></thead><tbody>' + rows + '</tbody>' +
    '<tfoot><tr class="sum-row"><td colspan="4" style="padding:6px 8px">รวมทั้งหมด</td>' +
      '<td class="r">' + _rptBaht(sumNet) + '</td><td colspan="3"></td></tr></tfoot></table>' +
    '<div class="footer">รายงานนี้สร้างโดยระบบ PTTS Cost Breakdown</div>' +
    _rptPrintScript() + '</body></html>';
  _rptOpenPrint(html);
}

// ══════════════════════════════════════════════════════════════
// REPORT 3 — ยอดซื้อ PO รายเดือน
// ══════════════════════════════════════════════════════════════
var _rpt3MonthList = [];
var _rpt3SelMonth  = '';
var _rpt3POs       = [];
var _rpt3SuppMap   = {};

async function rptPO() {
  var url = _rptUrl();
  if (!url) { Swal.fire('ยังไม่ตั้งค่า', 'กรุณาตั้งค่า Script URL ในแท็บ ตั้งค่า', 'warning'); return; }
  Swal.fire({ title: 'กำลังโหลด...', didOpen: function() { Swal.showLoading(); }, allowOutsideClick: false });
  try {
    var res = await fetch(_rptUrl() + '?action=getPurchaseOrders', { mode: 'cors' });
    var d   = await res.json();
    if (d.status === 'error') throw new Error(d.message);
    // headers เป็น array ของ array (raw rows จาก Code.gs)
    var rawHeaders = d.headers || [];
    var supps = await _rptGetSuppliers();
    _rpt3SuppMap = {};
    supps.forEach(function(s) { _rpt3SuppMap[s.code] = s.name; });

    // แปลง raw header array → object ตาม PO_COLS
    // poNo=0, issueDate=1, wantDate=2, supplierCode=3, refOrders=4, payTerm=5, deliverTerm=6, subtotal=7, vat=8, total=9, status=10
    _rpt3POs = rawHeaders.map(function(r) {
      return {
        poNo:         String(r[0]||'').trim(),
        issueDate:    String(r[1]||'').trim(),
        supplierCode: String(r[3]||'').trim(),
        subtotal:     parseFloat(r[7])||0,
        vat:          parseFloat(r[8])||0,
        total:        parseFloat(r[9])||0,
        status:       String(r[10]||'').trim(),
      };
    }).filter(function(p){ return p.poNo; });

    var mset = {};
    _rpt3POs.forEach(function(p) { var k = _rptMonthKey(p.issueDate); if (k) mset[k] = true; });
    _rpt3MonthList = Object.keys(mset).sort().reverse();
    if (!_rpt3MonthList.length) { Swal.fire('ไม่มีข้อมูล', 'ยังไม่มีใบสั่งซื้อในระบบ', 'info'); return; }
    _rpt3SelMonth = _rpt3MonthList[0];

    Swal.hideLoading(); Swal.close();
    _rpt3ShowSwal();
  } catch (e) {
    Swal.fire('โหลดไม่ได้', e.message, 'error');
  }
}

function _rpt3ShowSwal() {
  window.rpt3MonthChanged = function() {
    var sel = document.getElementById('rpt3MonthSel');
    if (sel) { _rpt3SelMonth = sel.value; var c = Swal.getHtmlContainer(); if (c) c.innerHTML = _rpt3Html(); }
  };
  Swal.fire({
    title: '🛒 ยอดซื้อ PO รายเดือน',
    background: 'var(--bg-card)', color: 'var(--t1)', width: 760,
    html: _rpt3Html(),
    showCancelButton: true,
    cancelButtonText: '🖨 พิมพ์',
    cancelButtonColor: '#0369a1',
    confirmButtonText: 'ปิด',
    confirmButtonColor: '#2563eb',
    reverseButtons: true,
  }).then(function(res) {
    delete window.rpt3MonthChanged;
    if (res.dismiss === Swal.DismissReason.cancel) _rpt3Print();
  });
}

function _rpt3Html() {
  var filtered = _rpt3POs.filter(function(p) { return _rptMonthKey(p.issueDate) === _rpt3SelMonth; });
  var sumSub = 0, sumVat = 0, sumTotal = 0;
  filtered.forEach(function(p) { sumSub += p.subtotal; sumVat += p.vat; sumTotal += p.total; });

  var rows = filtered.map(function(p, i) {
    var bg = i % 2 === 0 ? '' : 'background:rgba(0,0,0,.02)';
    var suppName = _rpt3SuppMap[p.supplierCode] || p.supplierCode;
    var stColor = p.status === 'อนุมัติ' ? '#16a34a' : p.status === 'ยกเลิก' ? '#dc2626' : '#d97706';
    return '<tr style="' + bg + '">' +
      '<td style="font-size:.72rem;color:var(--t3)">' + p.poNo + '</td>' +
      '<td style="font-size:.72rem">' + p.issueDate + '</td>' +
      '<td style="font-size:.78rem;font-weight:600">' + suppName + '</td>' +
      '<td style="text-align:right">' + _rptBaht(p.subtotal) + '</td>' +
      '<td style="text-align:right;color:#64748b">' + _rptBaht(p.vat) + '</td>' +
      '<td style="text-align:right;font-weight:700;color:var(--c1)">' + _rptBaht(p.total) + '</td>' +
      '<td style="text-align:center"><span style="font-size:.7rem;font-weight:700;color:' + stColor + ';background:' + stColor.replace('#','rgba(') + ',.1);padding:2px 8px;border-radius:5px">' + (p.status||'—') + '</span></td>' +
    '</tr>';
  }).join('');

  return '<div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;flex-wrap:wrap">' +
    '<span style="font-size:.85rem;color:var(--t2)">เดือน:</span>' +
    _rptMonthSelect('rpt3MonthSel', _rpt3MonthList, _rpt3SelMonth) +
    '<span style="font-size:.8rem;color:var(--t3)">' + filtered.length + ' รายการ</span>' +
  '</div>' +
  (filtered.length
    ? '<div style="overflow-x:auto;max-height:380px;overflow-y:auto">' +
      '<table style="width:100%;border-collapse:collapse;font-size:.79rem">' +
      '<thead><tr style="background:rgba(37,99,235,.1);position:sticky;top:0">' +
        '<th style="padding:6px 8px;text-align:left;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">เลขที่ PO</th>' +
        '<th style="padding:6px 8px;text-align:left;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">วันที่</th>' +
        '<th style="padding:6px 8px;text-align:left;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">ซัพพลายเออร์</th>' +
        '<th style="padding:6px 8px;text-align:right;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">ก่อน VAT</th>' +
        '<th style="padding:6px 8px;text-align:right;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">VAT</th>' +
        '<th style="padding:6px 8px;text-align:right;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">รวม</th>' +
        '<th style="padding:6px 8px;text-align:center;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">สถานะ</th>' +
      '</tr></thead>' +
      '<tbody>' + rows + '</tbody>' +
      '<tfoot><tr style="background:rgba(37,99,235,.07);font-weight:800">' +
        '<td colspan="3" style="padding:6px 8px;font-size:.78rem;color:var(--t1)">รวมทั้งเดือน ' + _rptMonthLabel(_rpt3SelMonth) + '</td>' +
        '<td style="padding:6px 8px;text-align:right;color:var(--c1)">' + _rptBaht(sumSub) + '</td>' +
        '<td style="padding:6px 8px;text-align:right;color:#64748b">' + _rptBaht(sumVat) + '</td>' +
        '<td style="padding:6px 8px;text-align:right;color:#16a34a;font-size:.88rem">' + _rptBaht(sumTotal) + '</td>' +
        '<td></td>' +
      '</tr></tfoot>' +
      '</table></div>'
    : '<div style="padding:32px;text-align:center;color:var(--t3)">ไม่มีรายการในเดือนนี้</div>');
}

function _rpt3Print() {
  var filtered = _rpt3POs.filter(function(p) { return _rptMonthKey(p.issueDate) === _rpt3SelMonth; });
  var sumSub = 0, sumVat = 0, sumTotal = 0;
  filtered.forEach(function(p) { sumSub += p.subtotal; sumVat += p.vat; sumTotal += p.total; });
  var rows = filtered.map(function(p, i) {
    var bg = i % 2 === 0 ? '#f8fafc' : '#fff';
    return '<tr style="background:' + bg + '">' +
      '<td class="c" style="color:#64748b;font-size:.7rem">' + (i+1) + '</td>' +
      '<td>' + p.poNo + '</td>' +
      '<td>' + p.issueDate + '</td>' +
      '<td>' + (_rpt3SuppMap[p.supplierCode] || p.supplierCode) + '</td>' +
      '<td class="r">' + _rptBaht(p.subtotal) + '</td>' +
      '<td class="r" style="color:#64748b">' + _rptBaht(p.vat) + '</td>' +
      '<td class="r" style="font-weight:700;color:#1d4ed8">' + _rptBaht(p.total) + '</td>' +
      '<td class="c">' + (p.status||'—') + '</td>' +
    '</tr>';
  }).join('');
  var html = '<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8"><title>ยอดซื้อ PO ' + _rptMonthLabel(_rpt3SelMonth) + '</title>' +
    _rptPrintCss() + '</head><body>' +
    '<div class="hdr"><div>' +
      '<div class="hdr-title">รายงานยอดซื้อใบสั่งซื้อ</div>' +
      '<div class="hdr-main">เดือน ' + _rptMonthLabel(_rpt3SelMonth) + '</div>' +
    '</div><div class="hdr-right">' + filtered.length + ' รายการ<br>พิมพ์เมื่อ: ' + new Date().toLocaleDateString('th-TH') + '</div></div>' +
    '<table><thead><tr>' +
      '<th class="c" style="width:28px">#</th>' +
      '<th style="width:90px">เลขที่ PO</th>' +
      '<th style="width:80px">วันที่</th>' +
      '<th>ซัพพลายเออร์</th>' +
      '<th class="r" style="width:85px">ก่อน VAT</th>' +
      '<th class="r" style="width:65px">VAT</th>' +
      '<th class="r" style="width:85px">รวม</th>' +
      '<th class="c" style="width:65px">สถานะ</th>' +
    '</tr></thead><tbody>' + rows + '</tbody>' +
    '<tfoot><tr class="sum-row"><td colspan="4" style="padding:6px 8px">รวมทั้งเดือน</td>' +
      '<td class="r">' + _rptBaht(sumSub) + '</td>' +
      '<td class="r">' + _rptBaht(sumVat) + '</td>' +
      '<td class="r" style="color:#1d4ed8">' + _rptBaht(sumTotal) + '</td>' +
      '<td></td></tr></tfoot></table>' +
    '<div class="footer">รายงานนี้สร้างโดยระบบ PTTS Cost Breakdown</div>' +
    _rptPrintScript() + '</body></html>';
  _rptOpenPrint(html);
}

// ══════════════════════════════════════════════════════════════
// REPORT 4 — ค่าชุบแยกโรงชุบ
// ══════════════════════════════════════════════════════════════
var _rpt4MonthList = [];
var _rpt4SelMonth  = '';
var _rpt4Plates    = [];
var _rpt4SuppMap   = {};

async function rptPlating() {
  var url = _rptUrl();
  if (!url) { Swal.fire('ยังไม่ตั้งค่า', 'กรุณาตั้งค่า Script URL ในแท็บ ตั้งค่า', 'warning'); return; }
  Swal.fire({ title: 'กำลังโหลด...', didOpen: function() { Swal.showLoading(); }, allowOutsideClick: false });
  try {
    var [plates, supps] = await Promise.all([_rptGetPlates(), _rptGetSuppliers()]);
    _rpt4Plates = plates;
    _rpt4SuppMap = {};
    supps.forEach(function(s) { _rpt4SuppMap[s.code] = s.name; });

    var mset = {};
    plates.forEach(function(p) { var k = _rptMonthKey(p.date); if (k) mset[k] = true; });
    _rpt4MonthList = Object.keys(mset).sort().reverse();
    if (!_rpt4MonthList.length) { Swal.fire('ไม่มีข้อมูล', 'ยังไม่มีประวัติส่งชุบในระบบ', 'info'); return; }
    _rpt4SelMonth = _rpt4MonthList[0];

    Swal.hideLoading(); Swal.close();
    _rpt4ShowSwal();
  } catch (e) {
    Swal.fire('โหลดไม่ได้', e.message, 'error');
  }
}

// คำนวณยอดรวมจาก items ของ platingNote
function _rpt4Subtotal(items) {
  var t = 0;
  (items || []).forEach(function(it) {
    var qty   = parseFloat(it.qty)   || 0;
    var price = parseFloat(it.price) || 0;
    t += qty * price;
  });
  return t;
}

function _rpt4ShowSwal() {
  window.rpt4MonthChanged = function() {
    var sel = document.getElementById('rpt4MonthSel');
    if (sel) { _rpt4SelMonth = sel.value; var c = Swal.getHtmlContainer(); if (c) c.innerHTML = _rpt4Html(); }
  };
  Swal.fire({
    title: '🧪 ค่าชุบแยกโรงชุบ',
    background: 'var(--bg-card)', color: 'var(--t1)', width: 740,
    html: _rpt4Html(),
    showCancelButton: true,
    cancelButtonText: '🖨 พิมพ์',
    cancelButtonColor: '#0369a1',
    confirmButtonText: 'ปิด',
    confirmButtonColor: '#2563eb',
    reverseButtons: true,
  }).then(function(res) {
    delete window.rpt4MonthChanged;
    if (res.dismiss === Swal.DismissReason.cancel) _rpt4Print();
  });
}

function _rpt4Html() {
  var filtered = _rpt4Plates.filter(function(p) { return _rptMonthKey(p.date) === _rpt4SelMonth; });

  // aggregate ตาม supplier
  var suppAgg = {};
  filtered.forEach(function(p) {
    var code = p.supplierCode || '—';
    if (!suppAgg[code]) suppAgg[code] = { count: 0, total: 0, notes: [] };
    var sub = _rpt4Subtotal(p.items);
    suppAgg[code].count++;
    suppAgg[code].total += sub;
    suppAgg[code].notes.push({ platingNo: p.platingNo, date: p.date, sub: sub });
  });

  var suppList = Object.keys(suppAgg).sort(function(a, b) { return suppAgg[b].total - suppAgg[a].total; });
  var grandTotal = 0;
  suppList.forEach(function(k) { grandTotal += suppAgg[k].total; });

  var rows = suppList.map(function(code, i) {
    var agg = suppAgg[code];
    var suppName = _rpt4SuppMap[code] || code;
    var pct = grandTotal > 0 ? ((agg.total / grandTotal) * 100).toFixed(1) : '0.0';
    var bg = i % 2 === 0 ? '' : 'background:rgba(0,0,0,.02)';
    return '<tr style="' + bg + '">' +
      '<td style="font-size:.78rem;font-weight:600">' + suppName + '</td>' +
      '<td style="font-size:.72rem;color:var(--t3)">' + code + '</td>' +
      '<td style="text-align:center;font-size:.82rem">' + agg.count + '</td>' +
      '<td style="text-align:right;font-weight:700;color:var(--c1)">' + _rptBaht(agg.total) + '</td>' +
      '<td style="text-align:center;font-size:.72rem;color:var(--t3)">' + pct + '%</td>' +
    '</tr>';
  }).join('');

  return '<div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;flex-wrap:wrap">' +
    '<span style="font-size:.85rem;color:var(--t2)">เดือน:</span>' +
    _rptMonthSelect('rpt4MonthSel', _rpt4MonthList, _rpt4SelMonth) +
    '<span style="font-size:.8rem;color:var(--t3)">' + filtered.length + ' ใบส่งชุบ · ' + suppList.length + ' โรงชุบ</span>' +
  '</div>' +
  (filtered.length
    ? '<div style="overflow-x:auto;max-height:380px;overflow-y:auto">' +
      '<table style="width:100%;border-collapse:collapse;font-size:.79rem">' +
      '<thead><tr style="background:rgba(37,99,235,.1);position:sticky;top:0">' +
        '<th style="padding:6px 8px;text-align:left;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">โรงชุบ</th>' +
        '<th style="padding:6px 8px;text-align:left;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">รหัส</th>' +
        '<th style="padding:6px 8px;text-align:center;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">จำนวนใบ</th>' +
        '<th style="padding:6px 8px;text-align:right;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">ยอดรวม</th>' +
        '<th style="padding:6px 8px;text-align:center;font-size:.67rem;color:var(--t3);border-bottom:2px solid var(--bc-div)">สัดส่วน</th>' +
      '</tr></thead>' +
      '<tbody>' + rows + '</tbody>' +
      '<tfoot><tr style="background:rgba(37,99,235,.07);font-weight:800">' +
        '<td colspan="3" style="padding:6px 8px;font-size:.78rem;color:var(--t1)">รวมทั้งเดือน ' + _rptMonthLabel(_rpt4SelMonth) + '</td>' +
        '<td style="padding:6px 8px;text-align:right;color:#16a34a;font-size:.88rem">' + _rptBaht(grandTotal) + '</td>' +
        '<td style="padding:6px 8px;text-align:center;font-size:.72rem;color:var(--t3)">100%</td>' +
      '</tr></tfoot>' +
      '</table></div>'
    : '<div style="padding:32px;text-align:center;color:var(--t3)">ไม่มีรายการในเดือนนี้</div>');
}

function _rpt4Print() {
  var filtered = _rpt4Plates.filter(function(p) { return _rptMonthKey(p.date) === _rpt4SelMonth; });
  var suppAgg = {};
  filtered.forEach(function(p) {
    var code = p.supplierCode || '—';
    if (!suppAgg[code]) suppAgg[code] = { count: 0, total: 0 };
    suppAgg[code].count++;
    suppAgg[code].total += _rpt4Subtotal(p.items);
  });
  var suppList = Object.keys(suppAgg).sort(function(a,b){ return suppAgg[b].total - suppAgg[a].total; });
  var grandTotal = 0;
  suppList.forEach(function(k){ grandTotal += suppAgg[k].total; });
  var rows = suppList.map(function(code, i) {
    var agg = suppAgg[code];
    var pct = grandTotal > 0 ? ((agg.total / grandTotal) * 100).toFixed(1) : '0.0';
    var bg = i % 2 === 0 ? '#f8fafc' : '#fff';
    return '<tr style="background:' + bg + '">' +
      '<td class="c" style="color:#64748b;font-size:.7rem">' + (i+1) + '</td>' +
      '<td>' + (_rpt4SuppMap[code] || code) + '</td>' +
      '<td>' + code + '</td>' +
      '<td class="c">' + agg.count + '</td>' +
      '<td class="r" style="font-weight:700;color:#1d4ed8">' + _rptBaht(agg.total) + '</td>' +
      '<td class="c">' + pct + '%</td>' +
    '</tr>';
  }).join('');
  var html = '<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8"><title>ค่าชุบ ' + _rptMonthLabel(_rpt4SelMonth) + '</title>' +
    _rptPrintCss() + '</head><body>' +
    '<div class="hdr"><div>' +
      '<div class="hdr-title">รายงานค่าชุบ</div>' +
      '<div class="hdr-main">เดือน ' + _rptMonthLabel(_rpt4SelMonth) + '</div>' +
    '</div><div class="hdr-right">' + filtered.length + ' ใบ · ' + suppList.length + ' โรงชุบ<br>พิมพ์เมื่อ: ' + new Date().toLocaleDateString('th-TH') + '</div></div>' +
    '<table><thead><tr>' +
      '<th class="c" style="width:28px">#</th>' +
      '<th>โรงชุบ</th>' +
      '<th style="width:80px">รหัส</th>' +
      '<th class="c" style="width:65px">จำนวนใบ</th>' +
      '<th class="r" style="width:95px">ยอดรวม</th>' +
      '<th class="c" style="width:65px">สัดส่วน</th>' +
    '</tr></thead><tbody>' + rows + '</tbody>' +
    '<tfoot><tr class="sum-row"><td colspan="4" style="padding:6px 8px">รวมทั้งเดือน</td>' +
      '<td class="r" style="color:#1d4ed8">' + _rptBaht(grandTotal) + '</td>' +
      '<td class="c">100%</td></tr></tfoot></table>' +
    '<div class="footer">รายงานนี้สร้างโดยระบบ PTTS Cost Breakdown</div>' +
    _rptPrintScript() + '</body></html>';
  _rptOpenPrint(html);
}
