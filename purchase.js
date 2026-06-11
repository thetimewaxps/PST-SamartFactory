// ════════════════════════════════════════════════════════════
//  purchase.js — ระบบใบสั่งซื้อ (Purchase Order)
//  เชื่อมกับ Order ผ่านฟิลด์ "อ้างอิง No.PO/No.Quo"
// ════════════════════════════════════════════════════════════

const PO_HEADER_COLS = {
  poNo:0, issueDate:1, wantDate:2, supplierCode:3, refOrders:4,
  payTerm:5, deliverTerm:6, subtotal:7, vat:8, total:9,
  status:10, createdBy:11, note:12, update:13
};
const PO_ITEM_COLS = { poNo:0, seq:1, name:2, spec:3, qty:4, unit:5, unitPrice:6, lineTotal:7 };

let _poCache = [];          // header rows
let _poItemsCache = {};     // poNo -> items rows
let _supplierCache = [];
let _poEditingNo = null;    // poNo ที่กำลังแก้ไข, null = สร้างใหม่
let _poItems = [];          // รายการสินค้าในการ์ดที่กำลังแก้ไข
let _poPage = 1;
const PO_PAGE_SIZE = 20;

// ── โหลด suppliers ──
async function fetchSuppliers() {
  if (!SCRIPT_URL) return;
  try {
    const res = await fetch(SCRIPT_URL + '?action=getSuppliers', {mode:'cors'});
    const data = await res.json();
    _supplierCache = data.suppliers || [];
    const sel = $('po_supplier');
    if (sel) {
      const cur = sel.value;
      sel.innerHTML = '<option value="">— เลือก Supplier —</option>' +
        _supplierCache.map(s => `<option value="${s.code}">${s.name} (${s.code})</option>`).join('');
      if (cur) sel.value = cur;
    }
  } catch (err) { console.error('fetchSuppliers', err); }
}

// ── โหลดใบสั่งซื้อทั้งหมด ──
async function fetchPurchaseOrders() {
  const tbody = $('poBody');
  if (!tbody) return;
  if (!SCRIPT_URL) {
    tbody.innerHTML = `<tr><td colspan="7" style="padding:30px;text-align:center;color:var(--t3);font-size:.8rem">⚠️ ยังไม่ได้ตั้งค่า Script URL</td></tr>`;
    return;
  }
  tbody.innerHTML = `<tr><td colspan="7" style="padding:30px;text-align:center;color:var(--t3);font-size:.8rem">↻ กำลังโหลด…</td></tr>`;
  try {
    const res = await fetch(SCRIPT_URL + '?action=getPurchaseOrders', {mode:'cors'});
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const data = await res.json();
    if (data.status === 'error') throw new Error(data.message || 'unknown');
    _poCache = (data.headers || []).slice().reverse(); // ใหม่สุดก่อน
    _poItemsCache = data.items || {};
    _poPage = 1;
    renderPOTable();
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="7" style="padding:30px;text-align:center;color:#f87171;font-size:.8rem">โหลดข้อมูลไม่สำเร็จ: ${err.message}</td></tr>`;
  }
}

function _poSupplierName(code) {
  const s = _supplierCache.find(x => x.code === code);
  return s ? s.name : (code || '—');
}

function _poStatusBadge(status) {
  const map = {
    'ร่าง':          {bg:'rgba(148,163,184,.15)', fg:'#94a3b8', bd:'rgba(148,163,184,.35)'},
    'ส่งแล้ว':       {bg:'rgba(56,189,248,.15)',  fg:'#0ea5e9', bd:'rgba(56,189,248,.35)'},
    'ได้รับของแล้ว': {bg:'rgba(34,197,94,.15)',   fg:'#16a34a', bd:'rgba(34,197,94,.35)'},
    'ยกเลิก':        {bg:'rgba(248,113,113,.15)', fg:'#ef4444', bd:'rgba(248,113,113,.35)'},
  };
  const c = map[status] || map['ร่าง'];
  return `<span style="display:inline-block;padding:2px 8px;border-radius:20px;font-size:.65rem;font-weight:600;
    background:${c.bg};color:${c.fg};border:1px solid ${c.bd}">${status||'ร่าง'}</span>`;
}

// ── เรนเดอร์ตารางรายการ PO ──
function renderPOTable() {
  const tbody = $('poBody');
  if (!tbody) return;
  const q = ($('poSearch')?.value || '').trim().toLowerCase();
  let rows = _poCache;
  if (q) {
    rows = rows.filter(r => [r[PO_HEADER_COLS.poNo], r[PO_HEADER_COLS.refOrders], _poSupplierName(r[PO_HEADER_COLS.supplierCode])]
      .some(v => String(v||'').toLowerCase().includes(q)));
  }
  if (rows.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" style="padding:30px;text-align:center;color:var(--t3);font-size:.8rem">ไม่มีใบสั่งซื้อ</td></tr>`;
    if ($('poPager')) $('poPager').innerHTML = '';
    return;
  }
  const totalPages = Math.max(1, Math.ceil(rows.length / PO_PAGE_SIZE));
  if (_poPage > totalPages) _poPage = totalPages;
  if (_poPage < 1) _poPage = 1;
  const startIdx = (_poPage - 1) * PO_PAGE_SIZE;
  const pageRows = rows.slice(startIdx, startIdx + PO_PAGE_SIZE);

  tbody.innerHTML = pageRows.map((r, ri) => {
    const poNo   = String(r[PO_HEADER_COLS.poNo]||'');
    const total  = parseFloat(r[PO_HEADER_COLS.total]) || 0;
    const status = String(r[PO_HEADER_COLS.status]||'ร่าง');
    const rowBg  = ri % 2 === 0 ? '' : 'background:var(--pair-bg)';
    return `<tr style="${rowBg};border-bottom:1px solid var(--bc-div)">
      <td style="padding:8px 10px;font-size:.78rem;font-weight:600;color:var(--c1);white-space:nowrap">${poNo}</td>
      <td style="padding:8px 10px;font-size:.72rem;color:var(--t3);white-space:nowrap">${r[PO_HEADER_COLS.issueDate]||'—'}</td>
      <td style="padding:8px 10px;font-size:.78rem;color:var(--t1)">${_poSupplierName(r[PO_HEADER_COLS.supplierCode])}</td>
      <td style="padding:8px 10px;font-size:.72rem;color:var(--t2)">${r[PO_HEADER_COLS.refOrders]||'—'}</td>
      <td style="padding:8px 10px;text-align:right;font-size:.78rem;font-weight:600;color:var(--c1);white-space:nowrap">${total ? total.toLocaleString('th-TH',{minimumFractionDigits:2}) : '—'} <span style="font-size:.65rem">฿</span></td>
      <td style="padding:8px 10px;white-space:nowrap">${_poStatusBadge(status)}</td>
      <td style="padding:8px 10px;text-align:center;white-space:nowrap">
        <button class="btn-fx" onclick="_poEdit('${poNo}')" style="padding:5px 10px;border-radius:7px;border:none;background:#2563eb;color:#fff;font-size:.7rem;cursor:pointer;font-family:Sarabun,sans-serif;margin:1px">✏️ แก้ไข</button>
        <button class="btn-fx" onclick="_poPrint('${poNo}')" style="padding:5px 10px;border-radius:7px;border:none;background:#16a34a;color:#fff;font-size:.7rem;cursor:pointer;font-family:Sarabun,sans-serif;margin:1px">🖨️ พิมพ์</button>
        <button class="btn-fx" onclick="_poDelete('${poNo}')" style="padding:5px 8px;border-radius:7px;border:1px solid rgba(248,113,113,.35);background:rgba(248,113,113,.1);color:#f87171;font-size:.7rem;cursor:pointer;margin:1px">🗑️</button>
      </td>
    </tr>`;
  }).join('');

  const pager = $('poPager');
  if (pager) {
    if (totalPages <= 1) {
      pager.innerHTML = `<span>ทั้งหมด ${rows.length} รายการ</span>`;
    } else {
      pager.innerHTML = `
        <button class="btn-fx" onclick="_poGoPage(${_poPage-1})" ${_poPage<=1?'disabled':''}
          style="padding:5px 12px;border-radius:7px;border:1px solid var(--bc-card);background:var(--bg-card);color:var(--t1);font-size:.75rem;cursor:pointer;${_poPage<=1?'opacity:.4;cursor:not-allowed':''}">‹ ก่อนหน้า</button>
        <span>หน้า ${_poPage} / ${totalPages} (ทั้งหมด ${rows.length} รายการ)</span>
        <button class="btn-fx" onclick="_poGoPage(${_poPage+1})" ${_poPage>=totalPages?'disabled':''}
          style="padding:5px 12px;border-radius:7px;border:1px solid var(--bc-card);background:var(--bg-card);color:var(--t1);font-size:.75rem;cursor:pointer;${_poPage>=totalPages?'opacity:.4;cursor:not-allowed':''}">ถัดไป ›</button>
      `;
    }
  }
}
function _poGoPage(p) { _poPage = p; renderPOTable(); }

// ── จัดการรายการสินค้าในฟอร์ม ──
function _poAddItemRow() {
  _poItems.push({ name:'', spec:'', qty:'', unit:'', unitPrice:'' });
  _poRenderItemsEditor();
}
function _poRemoveItemRow(idx) {
  _poItems.splice(idx, 1);
  _poRenderItemsEditor();
  _poRecalcTotals();
}
function _poItemChanged(idx, field, value) {
  if (!_poItems[idx]) return;
  _poItems[idx][field] = value;
  if (field === 'qty' || field === 'unitPrice') {
    const qty = parseFloat(_poItems[idx].qty) || 0;
    const price = parseFloat(_poItems[idx].unitPrice) || 0;
    const lineEl = $('poItemTotal_' + idx);
    if (lineEl) lineEl.textContent = (qty*price).toLocaleString('th-TH',{minimumFractionDigits:2});
    _poRecalcTotals();
  }
}
function _poRenderItemsEditor() {
  const wrap = $('poItemsBody');
  if (!wrap) return;
  if (_poItems.length === 0) {
    wrap.innerHTML = `<tr><td colspan="8" style="padding:14px;text-align:center;color:var(--t3);font-size:.78rem">ยังไม่มีรายการ — กด "เพิ่มรายการ"</td></tr>`;
    return;
  }
  wrap.innerHTML = _poItems.map((it, idx) => {
    const lineTotal = (parseFloat(it.qty)||0) * (parseFloat(it.unitPrice)||0);
    return `<tr>
      <td style="padding:6px 8px;text-align:center;font-size:.76rem;color:var(--t3)">${idx+1}</td>
      <td style="padding:6px 8px"><input type="text" value="${(it.name||'').replace(/"/g,'&quot;')}" oninput="_poItemChanged(${idx},'name',this.value)"
        style="width:100%;box-sizing:border-box;padding:5px 7px;border-radius:6px;border:1px solid var(--bc-input);background:var(--bg-input);color:var(--t1);font-size:.78rem;font-family:Sarabun,sans-serif"></td>
      <td style="padding:6px 8px"><input type="text" value="${(it.spec||'').replace(/"/g,'&quot;')}" oninput="_poItemChanged(${idx},'spec',this.value)"
        style="width:100%;box-sizing:border-box;padding:5px 7px;border-radius:6px;border:1px solid var(--bc-input);background:var(--bg-input);color:var(--t1);font-size:.78rem;font-family:Sarabun,sans-serif"></td>
      <td style="padding:6px 8px"><input type="number" value="${it.qty||''}" oninput="_poItemChanged(${idx},'qty',this.value)"
        style="width:100%;box-sizing:border-box;padding:5px 7px;border-radius:6px;border:1px solid var(--bc-input);background:var(--bg-input);color:var(--t1);font-size:.78rem;text-align:center;font-family:Sarabun,sans-serif"></td>
      <td style="padding:6px 8px"><input type="text" value="${(it.unit||'').replace(/"/g,'&quot;')}" oninput="_poItemChanged(${idx},'unit',this.value)"
        style="width:100%;box-sizing:border-box;padding:5px 7px;border-radius:6px;border:1px solid var(--bc-input);background:var(--bg-input);color:var(--t1);font-size:.78rem;text-align:center;font-family:Sarabun,sans-serif"></td>
      <td style="padding:6px 8px"><input type="number" value="${it.unitPrice||''}" oninput="_poItemChanged(${idx},'unitPrice',this.value)"
        style="width:100%;box-sizing:border-box;padding:5px 7px;border-radius:6px;border:1px solid var(--bc-input);background:var(--bg-input);color:var(--t1);font-size:.78rem;text-align:right;font-family:Sarabun,sans-serif"></td>
      <td style="padding:6px 8px;text-align:right;font-size:.78rem;font-weight:600;color:var(--c1);white-space:nowrap" id="poItemTotal_${idx}">${lineTotal.toLocaleString('th-TH',{minimumFractionDigits:2})}</td>
      <td style="padding:6px 8px;text-align:center">
        <button class="btn-fx" onclick="_poRemoveItemRow(${idx})" style="border:none;background:none;color:#f87171;cursor:pointer;font-size:.9rem">🗑️</button>
      </td>
    </tr>`;
  }).join('');
}
function _poRecalcTotals() {
  const subtotal = _poItems.reduce((s,it) => s + (parseFloat(it.qty)||0)*(parseFloat(it.unitPrice)||0), 0);
  const vat = subtotal * 0.07;
  const total = subtotal + vat;
  if ($('po_subtotal')) $('po_subtotal').textContent = subtotal.toLocaleString('th-TH',{minimumFractionDigits:2});
  if ($('po_vat'))      $('po_vat').textContent      = vat.toLocaleString('th-TH',{minimumFractionDigits:2});
  if ($('po_total'))    $('po_total').textContent    = total.toLocaleString('th-TH',{minimumFractionDigits:2});
  return { subtotal, vat, total };
}

// ── สร้างใบใหม่ / เคลียร์ฟอร์ม ──
async function _poNewForm() {
  _poEditingNo = null;
  _poItems = [];
  $('po_poNo').value = 'กำลังสร้างเลขที่...';
  $('po_issueDate').value = _todayStr();
  $('po_wantDate').value = '';
  $('po_supplier').value = '';
  $('po_refOrders').value = '';
  $('po_payTerm').value = '';
  $('po_deliverTerm').value = '';
  $('po_status').value = 'ร่าง';
  $('po_note').value = '';
  _poRenderItemsEditor();
  _poRecalcTotals();
  $('po_formTitle').textContent = '🧾 สร้างใบสั่งซื้อใหม่';
  $('po_saveBtn').textContent = '💾 บันทึกใบสั่งซื้อ';
  if (SCRIPT_URL) {
    try {
      const res = await fetch(SCRIPT_URL + '?action=getNextPONo', {mode:'cors'});
      const data = await res.json();
      $('po_poNo').value = data.nextPO || '';
    } catch (err) {
      $('po_poNo').value = '';
    }
  } else {
    $('po_poNo').value = '';
  }
}

// ── โหลด PO มาแก้ไข ──
function _poEdit(poNo) {
  const r = _poCache.find(row => String(row[PO_HEADER_COLS.poNo]) === String(poNo));
  if (!r) return;
  _poEditingNo = poNo;
  $('po_poNo').value = poNo;
  $('po_issueDate').value = _ordDateToInput(r[PO_HEADER_COLS.issueDate]);
  $('po_wantDate').value  = _ordDateToInput(r[PO_HEADER_COLS.wantDate]);
  $('po_supplier').value  = r[PO_HEADER_COLS.supplierCode] || '';
  $('po_refOrders').value = r[PO_HEADER_COLS.refOrders] || '';
  $('po_payTerm').value   = r[PO_HEADER_COLS.payTerm] || '';
  $('po_deliverTerm').value = r[PO_HEADER_COLS.deliverTerm] || '';
  $('po_status').value    = r[PO_HEADER_COLS.status] || 'ร่าง';
  $('po_createdBy').value = r[PO_HEADER_COLS.createdBy] || '';
  $('po_note').value      = r[PO_HEADER_COLS.note] || '';

  const items = _poItemsCache[poNo] || [];
  _poItems = items.map(it => ({
    name: it[PO_ITEM_COLS.name] || '',
    spec: it[PO_ITEM_COLS.spec] || '',
    qty:  it[PO_ITEM_COLS.qty] || '',
    unit: it[PO_ITEM_COLS.unit] || '',
    unitPrice: it[PO_ITEM_COLS.unitPrice] || '',
  }));
  _poRenderItemsEditor();
  _poRecalcTotals();
  $('po_formTitle').textContent = `✏️ แก้ไขใบสั่งซื้อ ${poNo}`;
  $('po_saveBtn').textContent = '💾 บันทึกการแก้ไข';
  document.getElementById('tab-po')?.scrollIntoView?.({behavior:'smooth', block:'start'});
}

// ── บันทึก PO ──
async function _poSave() {
  const poNo = String($('po_poNo').value || '').trim();
  if (!poNo || poNo === 'กำลังสร้างเลขที่...') {
    Swal.fire({icon:'warning', title:'ยังไม่ได้เลขที่ PO', text:'กรุณารอสักครู่หรือกด "ใบใหม่" อีกครั้ง', toast:true, position:'top-end', showConfirmButton:false, timer:2200});
    return;
  }
  if (!$('po_supplier').value) {
    Swal.fire({icon:'warning', title:'กรุณาเลือก Supplier', toast:true, position:'top-end', showConfirmButton:false, timer:2200});
    return;
  }
  if (_poItems.length === 0) {
    Swal.fire({icon:'warning', title:'กรุณาเพิ่มรายการสินค้าอย่างน้อย 1 รายการ', toast:true, position:'top-end', showConfirmButton:false, timer:2200});
    return;
  }
  const { subtotal, vat, total } = _poRecalcTotals();
  const header = [
    poNo,
    _ordDateToSheet($('po_issueDate').value || _todayStr()),
    _ordDateToSheet($('po_wantDate').value || ''),
    $('po_supplier').value,
    $('po_refOrders').value || '',
    $('po_payTerm').value || '',
    $('po_deliverTerm').value || '',
    subtotal, vat, total,
    $('po_status').value || 'ร่าง',
    $('po_createdBy').value || '',
    $('po_note').value || '',
    _ordDateToSheet(_todayStr()),
  ];
  const items = _poItems.map((it, idx) => [
    poNo, idx+1, it.name||'', it.spec||'', parseFloat(it.qty)||0, it.unit||'', parseFloat(it.unitPrice)||0,
    (parseFloat(it.qty)||0) * (parseFloat(it.unitPrice)||0)
  ]);

  const btn = $('po_saveBtn');
  const btnOldText = btn ? btn.textContent : '';
  if (btn) { btn.disabled = true; btn.textContent = '⏳ กำลังบันทึก...'; }
  Swal.fire({
    title: 'กำลังบันทึกใบสั่งซื้อ...', allowOutsideClick:false, allowEscapeKey:false,
    background:'#0d1b2a', color:'#cce4ff',
    didOpen: () => Swal.showLoading(),
  });
  try {
    await fetch(SCRIPT_URL, { method:'POST', mode:'no-cors',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ action:'savePurchaseOrder', poNo, header, items }) });
    await fetchPurchaseOrders();
    _poEditingNo = poNo;
    $('po_formTitle').textContent = `✏️ แก้ไขใบสั่งซื้อ ${poNo}`;
    if (btn) btn.textContent = '💾 บันทึกการแก้ไข';
    Swal.fire({icon:'success', title:'บันทึกใบสั่งซื้อแล้ว ✅',
      html:`เลขที่ PO: <b>${poNo}</b>`,
      background:'#0d1b2a', color:'#cce4ff',
      confirmButtonColor:'#6366f1', timer:1800, showConfirmButton:false});
  } catch (err) {
    if (btn) btn.textContent = btnOldText;
    Swal.fire({icon:'error', title:'บันทึกไม่สำเร็จ', text:String(err.message||err),
      background:'#0d1b2a', color:'#cce4ff', confirmButtonColor:'#6366f1'});
  } finally {
    if (btn) btn.disabled = false;
  }
}

// ── ลบ PO ──
async function _poDelete(poNo) {
  const ok = await Swal.fire({
    icon:'warning', title:'ลบใบสั่งซื้อนี้?', text:poNo,
    showCancelButton:true, confirmButtonText:'ลบ', cancelButtonText:'ยกเลิก',
    confirmButtonColor:'#ef4444', background:'var(--bg-deep)', color:'var(--t1)'
  });
  if (!ok.isConfirmed) return;
  Swal.fire({
    title: 'กำลังลบใบสั่งซื้อ...', allowOutsideClick:false, allowEscapeKey:false,
    background:'#0d1b2a', color:'#cce4ff',
    didOpen: () => Swal.showLoading(),
  });
  try {
    await fetch(SCRIPT_URL, { method:'POST', mode:'no-cors',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ action:'deletePurchaseOrder', poNo }) });
    await fetchPurchaseOrders();
    Swal.fire({icon:'success', title:'ลบแล้ว ✅', html:`เลขที่ PO: <b>${poNo}</b>`,
      background:'#0d1b2a', color:'#cce4ff',
      confirmButtonColor:'#6366f1', timer:1500, showConfirmButton:false});
    if (_poEditingNo === poNo) _poNewForm();
  } catch (err) {
    Swal.fire({icon:'error', title:'ลบไม่สำเร็จ', text:String(err.message||err),
      background:'#0d1b2a', color:'#cce4ff', confirmButtonColor:'#6366f1'});
  }
}

// ── พิมพ์ใบสั่งซื้อ A4 ──
function _poPrint(poNo) {
  const r = _poCache.find(row => String(row[PO_HEADER_COLS.poNo]) === String(poNo));
  if (!r) return;
  const items = _poItemsCache[poNo] || [];
  const supplier = _supplierCache.find(s => s.code === r[PO_HEADER_COLS.supplierCode]) || {};
  const fmt = n => (parseFloat(n)||0).toLocaleString('th-TH',{minimumFractionDigits:2});
  const subtotal = parseFloat(r[PO_HEADER_COLS.subtotal]) || 0;
  const vat = parseFloat(r[PO_HEADER_COLS.vat]) || 0;
  const total = parseFloat(r[PO_HEADER_COLS.total]) || 0;

  const itemRows = items.map(it => `
    <tr>
      <td style="text-align:center">${it[PO_ITEM_COLS.seq]}</td>
      <td>${it[PO_ITEM_COLS.name]||''}</td>
      <td>${it[PO_ITEM_COLS.spec]||''}</td>
      <td style="text-align:center">${it[PO_ITEM_COLS.qty]||''}</td>
      <td style="text-align:center">${it[PO_ITEM_COLS.unit]||''}</td>
      <td style="text-align:right">${fmt(it[PO_ITEM_COLS.unitPrice])}</td>
      <td style="text-align:right">${fmt(it[PO_ITEM_COLS.lineTotal])}</td>
    </tr>`).join('');

  const win = window.open('', '_blank');
  if (!win) return;
  win.document.write(`
    <html><head><title>ใบสั่งซื้อ ${poNo}</title>
    <style>
      @page { size: A4; margin: 16mm; }
      body{font-family:'Sarabun',Tahoma,sans-serif;color:#111;font-size:13px}
      h1{margin:0;font-size:22px}
      .sub{color:#666;font-size:12px;margin-bottom:14px}
      .grid{display:flex;justify-content:space-between;gap:24px;margin-bottom:16px}
      .box{flex:1;border:1px solid #ccc;border-radius:6px;padding:10px 12px}
      .box h3{margin:0 0 6px;font-size:13px;color:#444}
      table{width:100%;border-collapse:collapse;margin-top:10px}
      th,td{border:1px solid #ccc;padding:6px 8px;font-size:12px}
      th{background:#f3f4f6}
      .totals{margin-top:10px;display:flex;justify-content:flex-end}
      .totals table{width:280px}
      .totals td{border:none;padding:4px 8px}
      .totals tr.grand td{font-weight:700;font-size:14px;border-top:2px solid #333}
      .sign{display:flex;justify-content:space-between;margin-top:60px;text-align:center}
      .sign div{width:200px;border-top:1px solid #333;padding-top:6px;font-size:12px}
    </style></head><body>
    <h1>ใบสั่งซื้อ / Purchase Order</h1>
    <div class="sub">เลขที่: <b>${poNo}</b> &nbsp;|&nbsp; วันที่ออก: ${r[PO_HEADER_COLS.issueDate]||'—'} &nbsp;|&nbsp; วันที่ต้องการรับ: ${r[PO_HEADER_COLS.wantDate]||'—'}</div>
    <div class="grid">
      <div class="box">
        <h3>ผู้ขาย (Supplier)</h3>
        <div><b>${supplier.name||r[PO_HEADER_COLS.supplierCode]||'—'}</b></div>
        <div>${supplier.address||''}</div>
        <div>เลขผู้เสียภาษี: ${supplier.taxId||'—'}</div>
        <div>ติดต่อ: ${supplier.contact||'—'}</div>
      </div>
      <div class="box">
        <h3>เงื่อนไข</h3>
        <div>การชำระเงิน: ${r[PO_HEADER_COLS.payTerm]||'—'}</div>
        <div>การส่งมอบ: ${r[PO_HEADER_COLS.deliverTerm]||'—'}</div>
        <div>อ้างอิง No.PO/Quo: ${r[PO_HEADER_COLS.refOrders]||'—'}</div>
        <div>สถานะ: ${r[PO_HEADER_COLS.status]||'—'}</div>
      </div>
    </div>
    <table>
      <thead><tr><th style="width:5%">ลำดับ</th><th>รายการ</th><th>สเปค</th><th style="width:8%">จำนวน</th><th style="width:8%">หน่วย</th><th style="width:14%">ราคา/หน่วย</th><th style="width:14%">รวม</th></tr></thead>
      <tbody>${itemRows}</tbody>
    </table>
    <div class="totals">
      <table>
        <tr><td>รวมเป็นเงิน</td><td style="text-align:right">${fmt(subtotal)} บาท</td></tr>
        <tr><td>ภาษีมูลค่าเพิ่ม 7%</td><td style="text-align:right">${fmt(vat)} บาท</td></tr>
        <tr class="grand"><td>รวมสุทธิ</td><td style="text-align:right">${fmt(total)} บาท</td></tr>
      </table>
    </div>
    ${r[PO_HEADER_COLS.note] ? `<div style="margin-top:10px"><b>หมายเหตุ:</b> ${r[PO_HEADER_COLS.note]}</div>` : ''}
    <div class="sign">
      <div>ผู้จัดทำ<br>${r[PO_HEADER_COLS.createdBy]||''}</div>
      <div>ผู้อนุมัติ</div>
    </div>
    </body></html>`);
  win.document.close();
  win.focus();
  setTimeout(() => win.print(), 400);
}
