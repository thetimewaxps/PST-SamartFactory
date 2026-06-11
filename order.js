// ══════════════════════════════════════════════════════
// ══ TAB: Order — เชื่อมกับชีต "Order" (ใช้ร่วมกับ AppSheet) ══
// ══════════════════════════════════════════════════════
// คอลัมน์ A-AB (28 คอลัมน์, 0-based) — L-T (11-19) เป็นช่องสำรองของ AppSheet ห้ามแก้
const ORDER_COLS = {
  noQuo: 0, noPO: 1, orderDate: 2, mold: 3, workType: 4, productList: 5,
  qty: 6, material: 7, switch_: 8,
  price: 10,        // K = ราคาขาย
  note: 11,         // L = หมายเหตุ
  poFile: 12,       // M = รูปภาพPO. (ไฟล์แนบ PO)
  statusDeliver: 15,// P = สถานะส่งงาน (แสดงในตาราง, อ่านอย่างเดียว)
  process: 16,      // Q = Process (สถานะงานที่แก้ไขได้ — กำลังผลิต/ส่งซุป/.../เรียบร้อย)
  update: 20, linkImages: 21, status: 22, totalTax: 23, add: 24, my: 25,
  wantDate: 26, customer: 27
};
const ORDER_NUM_COLS = 28;
let _orderCache = [];
let _ordEditNoPO = null;

function _todayStr() {
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}
// แปลงวันที่ dd/MM/yyyy (จาก Sheet) <-> yyyy-MM-dd (สำหรับ <input type=date>)
function _ordDateToInput(s) {
  s = String(s||'').trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0');
  return s;
}
function _ordDateToSheet(s) {
  s = String(s||'').trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return `${parseInt(m[3])}/${parseInt(m[2])}/${parseInt(m[1])+543}`;
  return s;
}

// อัปเดตพรีวิวข้อมูลจากใบเสนอราคาที่เปิดอยู่
// แถวข้อมูลจากแท็บ DATA ที่เลือกผ่านปุ่ม "📦 AddOrder" — ใช้แทนระบบเดิมที่ดึงจากฟอร์มที่เปิดอยู่ (ช้า/สับสน)
let _ordSourceRow = null;

function updateOrderPreview() {
  const r = _ordSourceRow;
  if (!r) {
    ['ord_previewNoQuo','ord_previewCustomer','ord_previewWorkType','ord_previewProductList','ord_previewQty','ord_previewPrice'].forEach(id => {
      if ($(id)) $(id).textContent = '—';
    });
    if ($('ord_productList')) $('ord_productList').value = '';
    if ($('ord_material'))    $('ord_material').textContent = '—';
    _ordUpdateCreateBtn();
    return;
  }
  if ($('ord_previewNoQuo'))    $('ord_previewNoQuo').textContent    = r[DT.noQuo] || '—';
  if ($('ord_previewCustomer')) $('ord_previewCustomer').textContent = (() => {
    const sel = $('f_contact');
    const val = r[DT.contact] || '';
    if (sel) {
      const opt = Array.from(sel.options).find(o => o.value === val);
      if (opt) return opt.text || '—';
    }
    return val || '—';
  })();
  if ($('ord_previewWorkType')) $('ord_previewWorkType').textContent = r[DT.workType] || '—';
  if ($('ord_previewQty'))      $('ord_previewQty').textContent      = r[DT.unit] || '—';
  if ($('ord_previewPrice'))    $('ord_previewPrice').textContent    = r[DT.sellPrice] || '—';

  // รายการสินค้า = ขนาด (OD×ID×H) จาก DATA — ห้ามแก้ไขในการ์ด Order
  const od = r[DT.od]||'', id2 = r[DT.id]||'', h = r[DT.h]||'';
  const sizeVal = (od && id2 && h) ? `${od}×${id2}×${h}` : (r[DT.size]||'');
  const plEl = $('ord_productList');
  if (plEl) plEl.value = sizeVal;
  if ($('ord_previewProductList')) $('ord_previewProductList').textContent = sizeVal || '—';

  // แม่พิมพ์ — ตรวจสอบ OD จาก DATA เทียบกับตารางแม่พิมพ์ ว่ามีหรือไม่ (แก้ไขเพิ่มเติมได้ ไม่ทับค่าที่แก้เอง)
  _ordAutoFill('ord_mold', _ordCheckMoldOd());

  // หมายเหตุ — ดึงค่ามาให้อัตโนมัติ แต่แก้ไขเพิ่มเติมได้ (ไม่ทับค่าที่แก้เอง)
  _ordAutoFill('ord_note', r[DT.remark] || '');

  // วัตถุดิบ — ดึงจาก DATA แสดงเป็น label อย่างเดียว
  if ($('ord_material')) $('ord_material').textContent = r[DT.rawMat] || '—';

  _ordUpdateCreateBtn();
}
// เปลี่ยนข้อความ/พฤติกรรมปุ่มหลัก ตามว่าการ์ดมีข้อมูลจาก DATA แล้วหรือยัง
function _ordUpdateCreateBtn() {
  const btn = $('ord_createBtn');
  if (!btn) return;
  btn.innerHTML = _ordSourceRow ? '➕ สร้าง Order' : '🚀 เริ่มสร้าง Order';
}
// กดปุ่มหลัก: ถ้ายังไม่มีข้อมูลจาก DATA → ไปแท็บ DATA เพื่อกด "Order", ถ้ามีแล้ว → สร้าง Order
function _ordCreateBtnClick() {
  if (!_ordSourceRow) {
    switchTab('data');
    return;
  }
  createOrder();
}
// ตรวจสอบว่า OD ของรายการ ตรงกับตารางแม่พิมพ์หรือไม่ -> คืนค่า OD ถ้าตรง, "ไม่มี" ถ้าไม่ตรง
function _ordCheckMoldOd() {
  const odVal = parseFloat(_ordSourceRow ? _ordSourceRow[DT.od] : NaN);
  if (!odVal) return 'ไม่มี';
  const row = (_moldData || []).find(m => m.od === odVal);
  if (!row) return 'ไม่มี';
  const ids = row.ids || [];
  const noMold = ids.some(v => v.includes('ไม่มีพิมพ์'));
  if (noMold || ids.length === 0) return 'ไม่มี';
  return String(odVal);
}
// เติมค่าอัตโนมัติลงช่อง โดยไม่ทับค่าที่ผู้ใช้แก้ไขเองแล้ว
function _ordAutoFill(id, val) {
  const el = $(id);
  if (!el) return;
  if (el.value === '' || el.value === el.dataset.autoVal) {
    el.value = val;
    el.dataset.autoVal = val;
  }
}

// เคลียร์การ์ด "สร้าง Order" กลับสู่สถานะว่าง (ใช้ทั้งตอนสร้างสำเร็จ และตอนพบ No.Quo+No.PO ซ้ำ)
function _ordResetCard() {
  if ($('ord_noPO')) $('ord_noPO').value = '';
  if ($('ord_note'))  { $('ord_note').value = ''; delete $('ord_note').dataset.autoVal; }
  if ($('ord_mold'))  { $('ord_mold').value = ''; delete $('ord_mold').dataset.autoVal; }
  if ($('ord_workStatus')) $('ord_workStatus').value = 'ปรกติ';
  _ordSourceRow = null;
  updateOrderPreview();
  _ordClearPoFile('ord');
}

// ปุ่ม "📦 AddOrder" จากตาราง DATA (เฉพาะแถวที่สถานะ = "รอสรุป")
function dtAddOrder(idx) {
  const tbody = $('dtBody');
  const rows = tbody && tbody._filteredRows;
  if (!rows || !rows[idx]) return;
  const r = rows[idx];

  Swal.fire({
    title: 'สร้าง Order จากใบเสนอราคานี้?',
    html: `<small>No.Quo: <b>${r[DT.noQuo]}</b> — ข้อมูลในการ์ด "สร้าง Order" จะถูกแทนที่ด้วยข้อมูลจากรายการนี้</small>`,
    icon: 'question', showCancelButton: true,
    confirmButtonText: '✅ ดึงข้อมูล', cancelButtonText: 'ยกเลิก',
    background: '#0a1c2e', color: '#f1f5f9',
    confirmButtonColor: '#f59e0b', cancelButtonColor: '#475569'
  }).then(res => {
    if (!res.isConfirmed) return;
    _ordResetCard();   // ล้างค่าเก่าก่อน
    _ordSourceRow = r; // แล้วค่อยตั้งแถวใหม่
    switchTab('order');
  });
}

// เมื่อเลือกไฟล์ PO ใหม่ — แสดงชื่อไฟล์ + รูปตัวอย่าง (ถ้าเป็นรูปภาพ) + ปุ่มลบ
function _ordPoFileChanged(prefix) {
  const input = $(prefix + '_poFile');
  const file  = input?.files?.[0];
  const nameEl  = $(prefix + '_poFileName');
  const clearEl = $(prefix + '_poFileClear');
  const imgEl   = $(prefix + '_poFilePreview');
  if (!file) { _ordClearPoFile(prefix); return; }
  if (nameEl)  nameEl.textContent = file.name;
  if (clearEl) clearEl.style.display = 'inline-block';
  if (imgEl) {
    if (file.type && file.type.startsWith('image/')) {
      imgEl.src = URL.createObjectURL(file);
      imgEl.style.display = 'block';
    } else {
      imgEl.style.display = 'none';
      imgEl.removeAttribute('src');
    }
  }
}
// ลบไฟล์ PO ที่เลือกไว้
function _ordClearPoFile(prefix) {
  const input  = $(prefix + '_poFile');
  const nameEl = $(prefix + '_poFileName');
  const clearEl = $(prefix + '_poFileClear');
  const imgEl  = $(prefix + '_poFilePreview');
  if (input)  input.value = '';
  if (nameEl) nameEl.textContent = 'ยังไม่ได้เลือกไฟล์ใด';
  if (clearEl) clearEl.style.display = 'none';
  if (imgEl) { imgEl.style.display = 'none'; imgEl.removeAttribute('src'); }
}

// แปลงไฟล์เป็น base64 (ตัด prefix data:...;base64, ออก)
function _ordFileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result || '';
      const idx = result.indexOf('base64,');
      resolve(idx >= 0 ? result.slice(idx + 7) : result);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}
// อัปโหลดไฟล์ PO ไปยัง Drive ผ่าน Apps Script แล้วคืน URL
async function _ordUploadPoFile(file) {
  const base64 = await _ordFileToBase64(file);
  const res = await fetch(SCRIPT_URL, {
    method: 'POST', mode: 'cors',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify({ action: 'uploadOrderFile', fileName: file.name, mimeType: file.type, base64 })
  });
  const data = await res.json();
  if (data && data.status === 'ok') return data.url;
  throw new Error((data && data.message) || 'upload failed');
}

// สร้าง Order ใหม่จากใบเสนอราคาที่เปิดอยู่
async function createOrder() {
  if (!SCRIPT_URL) {
    Swal.fire({icon:'info',title:'ยังไม่ตั้งค่า URL',text:'กรุณาใส่ Apps Script URL ก่อน',background:'#0d1b2a',color:'#cce4ff',confirmButtonColor:'#6366f1'});
    return;
  }
  const noQuo = _ordSourceRow ? (_ordSourceRow[DT.noQuo] || '') : '';
  const noPO  = $('ord_noPO').value.trim();
  if (!noQuo) {
    Swal.fire({icon:'warning',title:'ไม่พบข้อมูลใบเสนอราคา',text:'กรุณากดปุ่ม "📦 AddOrder" จากตาราง DATA ก่อนสร้าง Order',background:'#0d1b2a',color:'#cce4ff',confirmButtonColor:'#6366f1'});
    return;
  }
  if (!noPO) {
    Swal.fire({icon:'warning',title:'กรุณาใส่ No.PO',background:'#0d1b2a',color:'#cce4ff',confirmButtonColor:'#6366f1'});
    return;
  }

  // ตรวจสอบว่ากรอกข้อมูลครบถ้วนก่อนบันทึก
  const missing = [];
  if (!$('ord_orderDate')?.value)            missing.push('วันที่สั่งซื้อ');
  if (!$('ord_wantDate')?.value)             missing.push('วันที่ต้องการ');
  if (!($('ord_mold')?.value || '').trim())  missing.push('แม่พิมพ์');
  if (!($('ord_note')?.value || '').trim())  missing.push('หมายเหตุ');
  if (!$('ord_poFile')?.files?.[0])          missing.push('แนบไฟล์ PO');
  if (missing.length) {
    Swal.fire({
      icon:'warning', title:'กรอกข้อมูลไม่ครบถ้วน',
      html:'กรุณากรอก/แนบข้อมูลให้ครบก่อนสร้าง Order:<br><b>' + missing.join(', ') + '</b>',
      background:'#0d1b2a', color:'#cce4ff', confirmButtonColor:'#6366f1'
    });
    return;
  }

  // ตรวจสอบ No.Quo ซ้ำ (คอลัมน์ A ห้ามซ้ำ)
  const dupQuo = (_orderCache || []).some(r =>
    String(r[ORDER_COLS.noQuo] || '').trim() === String(noQuo).trim()
  );
  if (dupQuo) {
    await Swal.fire({
      icon: 'error', title: 'พบ No.Quo ซ้ำ',
      html: `มี Order ที่ใช้ <b>No.Quo: ${noQuo}</b> อยู่แล้ว<br><span style="font-size:.8rem;color:#8b8aaa">No.Quo ห้ามซ้ำ ไม่สามารถสร้าง Order นี้ได้</span>`,
      background:'#0d1b2a', color:'#cce4ff',
      confirmButtonText: 'ตกลง', confirmButtonColor:'#dc2626'
    });
    if ($('ord_noPO')) { $('ord_noPO').value = ''; $('ord_noPO').focus(); }
    return;
  }

  // ตรวจสอบ No.PO ซ้ำ (คอลัมน์ B ห้ามซ้ำ)
  const dupPO = (_orderCache || []).some(r =>
    String(r[ORDER_COLS.noPO] || '').trim() === String(noPO).trim()
  );
  if (dupPO) {
    await Swal.fire({
      icon: 'error', title: 'พบ No.PO ซ้ำ',
      html: `มี Order ที่ใช้ <b>No.PO: ${noPO}</b> อยู่แล้ว<br><span style="font-size:.8rem;color:#8b8aaa">No.PO ห้ามซ้ำ ไม่สามารถสร้าง Order นี้ได้</span>`,
      background:'#0d1b2a', color:'#cce4ff',
      confirmButtonText: 'ตกลง', confirmButtonColor:'#dc2626'
    });
    if ($('ord_noPO')) { $('ord_noPO').value = ''; $('ord_noPO').focus(); }
    return;
  }

  // ยืนยันก่อนสร้าง Order จริง
  const confirmRes = await Swal.fire({
    icon: 'question', title: 'ยืนยันการสร้าง Order?',
    html: `No.Quo: <b>${noQuo}</b><br>No.PO: <b>${noPO}</b>`,
    showCancelButton: true,
    confirmButtonText: '✅ ยืนยัน', cancelButtonText: 'ยกเลิก',
    background: '#0a1c2e', color: '#f1f5f9',
    confirmButtonColor: '#22c55e', cancelButtonColor: '#475569'
  });
  if (!confirmRes.isConfirmed) return;

  const customer = ($('ord_previewCustomer')?.textContent || '').trim().replace(/^—$/, '') || '';

  const row = new Array(ORDER_NUM_COLS).fill('');
  row[ORDER_COLS.noQuo]       = noQuo;
  row[ORDER_COLS.noPO]        = noPO;
  row[ORDER_COLS.orderDate]   = _ordDateToSheet($('ord_orderDate').value || _todayStr());
  row[ORDER_COLS.mold]        = $('ord_mold')?.value || '';
  row[ORDER_COLS.workType]    = _ordSourceRow[DT.workType] || '';
  row[ORDER_COLS.productList] = $('ord_productList').value || '';
  row[ORDER_COLS.qty]         = _ordSourceRow[DT.unit] || '';
  row[ORDER_COLS.material]    = ($('ord_material')?.textContent || '').trim().replace(/^—$/, '') || '';
  row[ORDER_COLS.price]       = _ordSourceRow[DT.sellPrice] || '';
  row[ORDER_COLS.note]        = $('ord_note').value || '';
  row[ORDER_COLS.process]     = $('ord_status').value || 'กำลังผลิต';
  row[ORDER_COLS.status]      = $('ord_workStatus')?.value || 'ปรกติ';
  row[ORDER_COLS.wantDate]    = _ordDateToSheet($('ord_wantDate').value || '');
  row[ORDER_COLS.customer]    = customer;

  const createBtn = $('ord_createBtn');
  const statusEl  = $('ord_createStatus');
  if (createBtn) createBtn.disabled = true;

  try {
    const poFile = $('ord_poFile')?.files?.[0];
    if (poFile) {
      if (statusEl) statusEl.textContent = '⏳ กำลังอัปโหลดไฟล์ PO...';
      row[ORDER_COLS.poFile] = await _ordUploadPoFile(poFile);
    }

    await fetch(SCRIPT_URL, { method:'POST', mode:'no-cors',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ action:'addOrder', row }) });

    // อัปเดตสถานะของรายการต้นทาง (DATA) เป็น "ผ่าน" — แถวเดียวเท่านั้น (ตรวจ noQuo ก่อนส่ง)
    if (statusEl) statusEl.textContent = '⏳ กำลังอัปเดตสถานะ DATA...';
    await fetch(SCRIPT_URL, { method:'POST', mode:'no-cors',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ action:'updateDataStatus', noQuo, status:'ผ่าน' }) });

    Swal.fire({icon:'success',title:'สร้าง Order แล้ว ✅',
      html:`บันทึก Order เรียบร้อย<br><span style="font-size:.8rem;color:#8b8aaa">อัปเดตสถานะ DATA ของ No.Quo: <b>${noQuo}</b> เป็น "ผ่าน" แล้ว</span>`,
      background:'#0d1b2a',color:'#cce4ff',
      confirmButtonColor:'#6366f1', timer:2200, showConfirmButton:false});
    _ordResetCard();
    if (statusEl) statusEl.textContent = '';
    setTimeout(fetchOrders, 1200);
  } catch (err) {
    Swal.fire({icon:'error',title:'เกิดข้อผิดพลาด',text:'ส่งข้อมูลไม่สำเร็จ',background:'#0d1b2a',color:'#cce4ff',confirmButtonColor:'#6366f1'});
    if (statusEl) statusEl.textContent = '';
  } finally {
    if (createBtn) createBtn.disabled = false;
  }
}

// ── ลบรายการ Order: ลบแถวใน sheet "Order" + ตั้งสถานะ DATA ของ No.Quo เป็น "รอสรุป" ──
async function deleteOrderRow(noPO, noQuo) {
  if (!SCRIPT_URL) {
    Swal.fire({icon:'warning',title:'ยังไม่ได้ตั้งค่า Script URL',background:'#0d1b2a',color:'#cce4ff',confirmButtonColor:'#6366f1'});
    return;
  }
  const r = await Swal.fire({
    icon:'warning', title:'ลบรายการ Order นี้?',
    html:`<div style="text-align:left;font-size:.85rem;line-height:1.8">
      <div>📋 <b>No.Quo:</b> ${noQuo || '—'}</div>
      <div>📦 <b>No.PO:</b> ${noPO || '—'}</div>
      <div style="color:#f87171;margin-top:6px">⚠️ จะลบแถวนี้ออกจากชีต Order และตั้งสถานะ DATA ของ No.Quo เป็น "รอสรุป"</div>
    </div>`,
    showCancelButton:true, confirmButtonText:'🗑️ ลบ', cancelButtonText:'ยกเลิก',
    background:'#0d1b2a', color:'#cce4ff',
    confirmButtonColor:'#dc2626', cancelButtonColor:'#475569'
  });
  if (!r.isConfirmed) return;

  Swal.fire({title:'กำลังลบ…', background:'#0d1b2a', color:'#cce4ff', allowOutsideClick:false, showConfirmButton:false, didOpen:()=>Swal.showLoading()});
  try {
    await fetch(SCRIPT_URL, { method:'POST', mode:'no-cors',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ action:'deleteOrder', noPO }) });

    if (noQuo) {
      await fetch(SCRIPT_URL, { method:'POST', mode:'no-cors',
        headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ action:'updateDataStatus', noQuo, status:'รอสรุป' }) });
    }

    Swal.fire({icon:'success', title:'ลบแล้ว ✅',
      html:`ลบรายการ Order เรียบร้อย${noQuo ? `<br><span style="font-size:.8rem;color:#8b8aaa">ตั้งสถานะ DATA ของ No.Quo: <b>${noQuo}</b> เป็น "รอสรุป" แล้ว</span>` : ''}`,
      background:'#0d1b2a', color:'#cce4ff', confirmButtonColor:'#6366f1', timer:2200, showConfirmButton:false});
    setTimeout(fetchOrders, 1200);
  } catch (err) {
    Swal.fire({icon:'error', title:'เกิดข้อผิดพลาด', text:'ลบไม่สำเร็จ', background:'#0d1b2a', color:'#cce4ff', confirmButtonColor:'#6366f1'});
  }
}

// โหลดรายการ Order จากชีต "Order"
async function fetchOrders() {
  const tbody = $('ordBody');
  if (!tbody) return;
  if (!SCRIPT_URL) {
    tbody.innerHTML = `<tr><td colspan="10" style="padding:30px;text-align:center;color:var(--t3);font-size:.8rem">⚠️ ยังไม่ได้ตั้งค่า Script URL</td></tr>`;
    return;
  }
  tbody.innerHTML = `<tr><td colspan="10" style="padding:30px;text-align:center;color:var(--t3);font-size:.8rem">↻ กำลังโหลด…</td></tr>`;
  try {
    const res = await fetch(SCRIPT_URL + '?action=getOrders', {mode:'cors'});
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const data = await res.json();
    if (data.status === 'error') throw new Error(data.message || 'unknown');
    _orderCache = (data.rows || []).slice().reverse(); // ใหม่สุดก่อน
    _initSeenIfEmpty(SEEN_KEY_ORDER, _orderCache.map(r => r[ORDER_COLS.noPO]));
    _ordPage = 1;
    renderOrderTable();
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="10" style="padding:30px;text-align:center;color:#f87171;font-size:.8rem">โหลดข้อมูลไม่สำเร็จ: ${err.message}</td></tr>`;
  }
}

const ORD_PAGE_SIZE = 20;
let _ordPage = 1;

function _ordGoPage(p) {
  _ordPage = p;
  renderOrderTable();
}

function _ordTruncate(s, n) {
  s = String(s||'');
  return s.length > n ? s.slice(0, n-1) + '…' : s;
}

function _ordResetFilters() {
  $('ordSearch').value = '';
  $('ordFilterStatus').value = '_not_done';
  $('ordFilterCustomer').value = '';
  $('ordFilterWantFrom').value = '';
  $('ordFilterWantTo').value = '';
  _ordPage = 1;
  renderOrderTable();
}

function renderOrderTable() {
  const tbody = $('ordBody');
  if (!tbody) return;
  const q = ($('ordSearch')?.value || '').trim().toLowerCase();
  let rows = _orderCache;
  if (q) {
    rows = rows.filter(r => [r[ORDER_COLS.noQuo], r[ORDER_COLS.noPO], r[ORDER_COLS.customer], r[ORDER_COLS.productList]]
      .some(v => String(v||'').toLowerCase().includes(q)));
  }

  // ── ตัวกรอง: สถานะงาน (ค่าเริ่มต้น = ซ่อนรายการที่ "เรียบร้อย") ──
  const statusFilter = $('ordFilterStatus')?.value ?? '_not_done';
  if (statusFilter === '_not_done') {
    rows = rows.filter(r => String(r[ORDER_COLS.process]||'').trim() !== 'เรียบร้อย');
  } else if (statusFilter !== '_all') {
    rows = rows.filter(r => String(r[ORDER_COLS.process]||'').trim() === statusFilter);
  }

  // ── ตัวกรอง: ลูกค้า ──
  const custFilter = ($('ordFilterCustomer')?.value || '').trim().toLowerCase();
  if (custFilter) {
    rows = rows.filter(r => String(r[ORDER_COLS.customer]||'').toLowerCase().includes(custFilter));
  }

  // ── ตัวกรอง: ช่วงวันที่ต้องการ ──
  const wantFrom = $('ordFilterWantFrom')?.value || '';
  const wantTo   = $('ordFilterWantTo')?.value || '';
  if (wantFrom || wantTo) {
    rows = rows.filter(r => {
      const iso = _ordDateToInput(r[ORDER_COLS.wantDate]);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return false;
      if (wantFrom && iso < wantFrom) return false;
      if (wantTo   && iso > wantTo)   return false;
      return true;
    });
  }
  if (rows.length === 0) {
    tbody.innerHTML = `<tr><td colspan="10" style="padding:30px;text-align:center;color:var(--t3);font-size:.8rem">ไม่มีข้อมูล Order</td></tr>`;
    $('ordPager').innerHTML = '';
    return;
  }

  const totalPages = Math.max(1, Math.ceil(rows.length / ORD_PAGE_SIZE));
  if (_ordPage > totalPages) _ordPage = totalPages;
  if (_ordPage < 1) _ordPage = 1;
  const startIdx = (_ordPage - 1) * ORD_PAGE_SIZE;
  const pageRows = rows.slice(startIdx, startIdx + ORD_PAGE_SIZE);

  tbody.innerHTML = pageRows.map(r => {
    const noPO  = String(r[ORDER_COLS.noPO]||'');
    const price = parseFloat(r[ORDER_COLS.price]) || 0;
    const productList = String(r[ORDER_COLS.productList]||'');
    const customer    = String(r[ORDER_COLS.customer]||'');
    const isNewRow = _isNewItem(SEEN_KEY_ORDER, noPO);
    return `<tr style="border-bottom:1px solid var(--bc-div)">
      <td style="padding:8px 8px;font-size:.78rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${r[ORDER_COLS.noQuo]||'—'}</td>
      <td style="padding:8px 8px;font-size:.78rem;font-weight:600;color:var(--c1);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${noPO||'—'}${_newBadge(isNewRow)}</td>
      <td style="padding:8px 8px;font-size:.78rem;white-space:nowrap">${r[ORDER_COLS.orderDate]||'—'}</td>
      <td style="padding:8px 8px;font-size:.78rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${customer.replace(/"/g,'&quot;')}">${_ordTruncate(customer,12)||'—'}</td>
      <td style="padding:8px 8px;font-size:.78rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${productList.replace(/"/g,'&quot;')}">${_ordTruncate(productList,28)||'—'}</td>
      <td style="padding:8px 8px;font-size:.78rem;text-align:center">${r[ORDER_COLS.qty]||'—'}</td>
      <td style="padding:8px 8px;font-size:.78rem;text-align:right;white-space:nowrap">${price ? price.toLocaleString('th-TH',{minimumFractionDigits:2}) : '—'}</td>
      <td style="padding:8px 8px;font-size:.78rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${r[ORDER_COLS.process]||'—'}</td>
      <td style="padding:8px 8px;font-size:.78rem;white-space:nowrap">${r[ORDER_COLS.wantDate]||'—'}</td>
      <td style="padding:8px 6px;text-align:center">
        <div style="display:flex;flex-wrap:wrap;gap:4px;justify-content:center">
          <button onclick="showOrderDetail('${noPO.replace(/'/g,"\\'")}')"
            style="padding:5px 8px;border-radius:7px;border:1px solid rgba(34,197,94,.35);
            background:rgba(34,197,94,.1);color:#4ade80;font-size:.72rem;cursor:pointer;white-space:nowrap">👁️ ดู</button>
          <button onclick="openEditOrder('${noPO.replace(/'/g,"\\'")}')"
            style="padding:5px 8px;border-radius:7px;border:1px solid rgba(99,102,241,.35);
            background:rgba(99,102,241,.1);color:#9b8fff;font-size:.72rem;cursor:pointer;white-space:nowrap">✏️ แก้ไข</button>
          <button onclick="deleteOrderRow('${noPO.replace(/'/g,"\\'")}','${String(r[ORDER_COLS.noQuo]||'').replace(/'/g,"\\'")}')"
            style="padding:5px 8px;border-radius:7px;border:1px solid rgba(248,113,113,.35);
            background:rgba(248,113,113,.1);color:#f87171;font-size:.72rem;cursor:pointer;white-space:nowrap">🗑️ ลบ</button>
        </div>
      </td>
    </tr>`;
  }).join('');

  // ── ตัวควบคุมแบ่งหน้า ──
  const pager = $('ordPager');
  if (pager) {
    if (totalPages <= 1) {
      pager.innerHTML = `<span>ทั้งหมด ${rows.length} รายการ</span>`;
    } else {
      pager.innerHTML = `
        <button onclick="_ordGoPage(${_ordPage-1})" ${_ordPage<=1?'disabled':''}
          style="padding:5px 12px;border-radius:7px;border:1px solid var(--bc-card);background:var(--bg-card);color:var(--t1);font-size:.75rem;cursor:pointer;${_ordPage<=1?'opacity:.4;cursor:not-allowed':''}">‹ ก่อนหน้า</button>
        <span>หน้า ${_ordPage} / ${totalPages} (ทั้งหมด ${rows.length} รายการ)</span>
        <button onclick="_ordGoPage(${_ordPage+1})" ${_ordPage>=totalPages?'disabled':''}
          style="padding:5px 12px;border-radius:7px;border:1px solid var(--bc-card);background:var(--bg-card);color:var(--t1);font-size:.75rem;cursor:pointer;${_ordPage>=totalPages?'opacity:.4;cursor:not-allowed':''}">ถัดไป ›</button>
      `;
    }
  }
}

// ── สีของ badge ตาม Process ──
function _ordProcessColor(process) {
  const p = String(process||'').trim();
  if (p === 'เรียบร้อย')   return {bg:'rgba(34,197,94,.15)',  fg:'#4ade80', bd:'rgba(34,197,94,.35)'};
  if (p === 'กำลังผลิต')   return {bg:'rgba(245,158,11,.15)', fg:'#fbbf24', bd:'rgba(245,158,11,.35)'};
  if (p === 'รอConfirm')   return {bg:'rgba(148,163,184,.15)',fg:'#cbd5e1', bd:'rgba(148,163,184,.35)'};
  if (p === 'Stock' || p === 'FG รอเรียก') return {bg:'rgba(99,102,241,.15)', fg:'#9b8fff', bd:'rgba(99,102,241,.35)'};
  return {bg:'rgba(56,189,248,.15)', fg:'#7dd3fc', bd:'rgba(56,189,248,.35)'};
}
// ── สีของ badge ตามสถานะส่งงาน ──
function _ordDeliverColor(s) {
  const v = String(s||'').trim();
  if (v.includes('ส่งแล้ว') || v.includes('เรียบร้อย')) return {bg:'rgba(34,197,94,.15)', fg:'#4ade80', bd:'rgba(34,197,94,.35)'};
  if (v.includes('ยังไม่ส่ง') || v === '') return {bg:'rgba(248,113,113,.15)', fg:'#f87171', bd:'rgba(248,113,113,.35)'};
  return {bg:'rgba(56,189,248,.15)', fg:'#7dd3fc', bd:'rgba(56,189,248,.35)'};
}
// ── แตกข้อความรูปภาพหลายรูป (คั่นด้วย , / ; / ขึ้นบรรทัดใหม่) ──
function _ordSplitImages(s) {
  return String(s||'').split(/[\n,;]+/).map(x => x.trim()).filter(Boolean);
}

// ── ลำดับขั้นตอนงาน สำหรับ Timeline (ความคืบหน้า) ──
const ORDER_FLOW_STEPS = [
  {label:'รับออเดอร์',          icon:'📋'},
  {label:'ตรวจสอบข้อมูล',       icon:'✅'},
  {label:'กำลังผลิต',           icon:'⚙️'},
  {label:'QC ตรวจสอบคุณภาพ',    icon:'🔍'},
  {label:'ส่งสินค้า',            icon:'🚚'},
];
// คืนค่า index ของขั้นตอนปัจจุบัน (0-4) หรือ 5 = เสร็จสิ้นทั้งหมด
function _ordCurrentStepIdx(r) {
  const p = String(r[ORDER_COLS.process]||'').trim();
  const delivered = String(r[ORDER_COLS.statusDeliver]||'').includes('ส่งแล้ว');
  if (delivered || p === 'เรียบร้อย') return 5;
  if (p === 'รอConfirm') return 1;
  if (p === 'กำลังผลิต') return 2;
  if (['ส่งซุป','ส่งตัวอย่างเทส+รอสรุป','ส่งยังไม่ครบ'].includes(p)) return 3;
  if (p === 'FG รอเรียก' || p === 'Stock') return 4;
  return 2; // ค่าเริ่มต้น
}
// สร้าง HTML ของ Timeline แนวตั้ง
function _ordTimelineHtml(r) {
  const cur = _ordCurrentStepIdx(r);
  const orderDate = String(r[ORDER_COLS.orderDate]||'').trim() || '—';
  const wantDate  = String(r[ORDER_COLS.wantDate]||'').trim() || '—';
  return ORDER_FLOW_STEPS.map((s,i) => {
    const isLast = i === ORDER_FLOW_STEPS.length - 1;
    let dot, dateTxt, badge;
    if (i < cur) {
      dot = `<div style="width:26px;height:26px;border-radius:50%;background:rgba(34,197,94,.18);border:2px solid #22c55e;
        display:flex;align-items:center;justify-content:center;font-size:.85rem;color:#4ade80;flex-shrink:0">✓</div>`;
      dateTxt = i === 0 ? orderDate : '';
      badge = `<span style="font-size:.66rem;font-weight:700;color:#4ade80;background:rgba(34,197,94,.12);
        border-radius:6px;padding:1px 7px">เสร็จสิ้น</span>`;
    } else if (i === cur) {
      dot = `<div style="width:26px;height:26px;border-radius:50%;background:rgba(245,158,11,.18);border:2px solid #f59e0b;
        display:flex;align-items:center;justify-content:center;font-size:.8rem;flex-shrink:0">${s.icon}</div>`;
      dateTxt = wantDate !== '—' ? `กำหนด: ${wantDate}` : '';
      badge = `<span style="font-size:.66rem;font-weight:700;color:#fbbf24;background:rgba(245,158,11,.12);
        border-radius:6px;padding:1px 7px">กำลังดำเนินการ</span>`;
    } else {
      dot = `<div style="width:26px;height:26px;border-radius:50%;background:rgba(148,163,184,.12);border:2px solid #475569;
        display:flex;align-items:center;justify-content:center;font-size:.8rem;color:#64748b;flex-shrink:0">${s.icon}</div>`;
      dateTxt = '';
      badge = `<span style="font-size:.66rem;font-weight:700;color:var(--t3);background:rgba(148,163,184,.1);
        border-radius:6px;padding:1px 7px">รอดำเนินการ</span>`;
    }
    return `
      <div style="display:flex;gap:10px">
        <div style="display:flex;flex-direction:column;align-items:center">
          ${dot}
          ${!isLast ? `<div style="flex:1;width:2px;min-height:26px;background:${i < cur ? '#22c55e' : 'rgba(148,163,184,.25)'};margin:2px 0"></div>` : ''}
        </div>
        <div style="padding-bottom:18px">
          <div style="font-size:.82rem;font-weight:700;color:${i<=cur?'var(--t1)':'var(--t3)'}">${s.label}</div>
          ${dateTxt ? `<div style="font-size:.7rem;color:var(--t3);margin:2px 0">${dateTxt}</div>` : ''}
          <div style="margin-top:3px">${badge}</div>
        </div>
      </div>`;
  }).join('');
}

// ── การ์ด: ดูรายละเอียด Order ──
function showOrderDetail(noPO) {
  const r = _orderCache.find(row => String(row[ORDER_COLS.noPO]) === String(noPO));
  if (!r) return;
  _addSeen(SEEN_KEY_ORDER, noPO);
  renderOrderTable();

  const g = (k) => String(r[ORDER_COLS[k]] ?? '').trim() || '—';
  const procColor = _ordProcessColor(r[ORDER_COLS.process]);
  const delColor  = _ordDeliverColor(r[ORDER_COLS.statusDeliver]);
  const price = parseFloat(r[ORDER_COLS.price]) || 0;
  const noQuo = String(r[ORDER_COLS.noQuo]||'');
  const escPO   = noPO.replace(/'/g,"\\'");
  const escQuo  = noQuo.replace(/'/g,"\\'");

  const poUrl = String(r[ORDER_COLS.poFile]||'').trim();
  const jobImgs = _ordSplitImages(r[ORDER_COLS.linkImages]);
  const shipDate = String(r[ORDER_COLS.update]||'').trim() || '—';

  const rowField = (icon, label, value) => `
    <div style="display:flex;align-items:flex-start;gap:8px;padding:6px 0">
      <div style="font-size:1rem;line-height:1.4;flex-shrink:0">${icon}</div>
      <div style="min-width:0">
        <div style="font-size:.68rem;color:var(--t3)">${label}</div>
        <div style="font-size:.85rem;font-weight:600;color:var(--t1);word-break:break-word">${value}</div>
      </div>
    </div>`;

  const imgCard = (label, url) => `
    <div style="text-align:center">
      <div style="font-size:.68rem;color:var(--t3);margin-bottom:4px">${label}</div>
      ${url
        ? `<div style="position:relative">
             <img src="${url}" style="width:100%;height:90px;object-fit:cover;border-radius:8px;border:1px solid var(--bc-card)">
             <button onclick="_ordZoomImage('${String(url).replace(/'/g,"\\'")}','${label.replace(/'/g,"\\'")}')"
               style="position:absolute;bottom:5px;right:5px;width:24px;height:24px;border-radius:50%;border:none;
               background:rgba(0,0,0,.55);color:#fff;font-size:.78rem;cursor:pointer;display:flex;align-items:center;justify-content:center">🔍</button>
           </div>`
        : `<div style="width:100%;height:90px;border-radius:8px;border:1px dashed var(--bc-card);
             display:flex;align-items:center;justify-content:center;color:var(--t3);font-size:.7rem">ไม่มีรูป</div>`}
    </div>`;

  const html = `
    <div style="text-align:left;font-family:inherit">
      <!-- หัวการ์ด -->
      <div style="display:flex;flex-wrap:wrap;justify-content:space-between;align-items:flex-start;gap:8px;
        padding-bottom:10px;margin-bottom:10px;border-bottom:1px solid var(--bc-div)">
        <div style="display:flex;gap:18px;flex-wrap:wrap">
          <div><div style="font-size:.68rem;color:var(--t3)">No.Quo</div><div style="font-size:1rem;font-weight:800;color:var(--c1)">${g('noQuo')}</div></div>
          <div><div style="font-size:.68rem;color:var(--t3)">No.PO</div><div style="font-size:1rem;font-weight:800;color:#9b8fff">${g('noPO')}</div></div>
          <div><div style="font-size:.68rem;color:var(--t3)">ลูกค้า</div><div style="font-size:1rem;font-weight:700;color:var(--t1)">${g('customer')}</div></div>
        </div>
        <div style="display:flex;gap:6px;flex-wrap:wrap">
          <span style="padding:4px 10px;border-radius:999px;font-size:.72rem;font-weight:700;
            background:${procColor.bg};color:${procColor.fg};border:1px solid ${procColor.bd}">${g('process')}</span>
          <span style="padding:4px 10px;border-radius:999px;font-size:.72rem;font-weight:700;
            background:${delColor.bg};color:${delColor.fg};border:1px solid ${delColor.bd}">📦 ${g('statusDeliver')}</span>
        </div>
      </div>

      <div style="display:grid;grid-template-columns:1.5fr 1fr;gap:16px;align-items:start">
        <!-- ── คอลัมน์ซ้าย: รายละเอียด ── -->
        <div>
          <!-- ข้อมูลหลัก -->
          <div style="font-size:.7rem;font-weight:700;color:var(--c1);margin-bottom:2px">📋 ข้อมูลหลัก</div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:0 14px;background:var(--bc-div);
            border-radius:10px;padding:6px 12px;margin-bottom:12px">
            ${rowField('📅','วันที่สั่งซื้อ', g('orderDate'))}
            ${rowField('🚚','วันที่ส่งงาน', shipDate)}
            ${rowField('⏰','วันที่ต้องการ', g('wantDate'))}
            ${rowField('🏷️','สถานะงาน', g('status'))}
            ${rowField('🧱','แม่พิมพ์', g('mold'))}
            ${rowField('📐','แบบงาน', g('workType'))}
            ${rowField('📦','รายการสินค้า', g('productList'))}
            ${rowField('🔢','จำนวน', g('qty'))}
            ${rowField('🧪','วัตถุดิบ', g('material'))}
            ${rowField('💰','ราคาเสนอขาย', price ? price.toLocaleString('th-TH',{minimumFractionDigits:2}) + ' บาท' : '—')}
          </div>

          <!-- รูปภาพ -->
          <div style="font-size:.7rem;font-weight:700;color:var(--c1);margin-bottom:6px">🖼️ รูปภาพ</div>
          <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:12px">
            ${imgCard('รูปภาพ PO.', poUrl)}
            ${imgCard('รูปงาน 1', jobImgs[0]||'')}
            ${imgCard('รูปงาน 2', jobImgs[1]||'')}
          </div>

          <!-- หมายเหตุ / Note -->
          <div style="font-size:.7rem;font-weight:700;color:var(--c1);margin-bottom:6px">📝 หมายเหตุ / Note</div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:4px">
            <div style="background:var(--bc-div);border-radius:10px;padding:10px 12px;
              font-size:.8rem;color:var(--t1);white-space:pre-wrap;word-break:break-word;min-height:48px">
              <div style="font-size:.66rem;color:var(--t3);margin-bottom:3px">หมายเหตุ</div>${g('note')}</div>
            <div style="background:var(--bc-div);border-radius:10px;padding:10px 12px;
              font-size:.8rem;color:var(--t1);white-space:pre-wrap;word-break:break-word;min-height:48px">
              <div style="font-size:.66rem;color:var(--t3);margin-bottom:3px">Note</div>—</div>
          </div>
        </div>

        <!-- ── คอลัมน์ขวา: Timeline ── -->
        <div>
          <div style="font-size:.7rem;font-weight:700;color:var(--c1);margin-bottom:8px">📈 ความคืบหน้า (Timeline)</div>
          <div style="background:var(--bc-div);border-radius:10px;padding:12px 12px 0 12px">
            ${_ordTimelineHtml(r)}
          </div>
        </div>
      </div>

      <!-- ปุ่ม Action -->
      <div style="display:flex;flex-wrap:wrap;gap:8px;justify-content:flex-end;margin-top:14px;
        padding-top:10px;border-top:1px solid var(--bc-div)">
        <button onclick="_ordQuickChangeProcess('${escPO}')"
          style="padding:7px 12px;border-radius:8px;border:1px solid rgba(245,158,11,.35);
          background:rgba(245,158,11,.1);color:#d97706;font-size:.78rem;font-weight:700;cursor:pointer">🔄 เปลี่ยนสถานะ</button>
        <button onclick="_ordMarkDelivered('${escPO}')"
          style="padding:7px 12px;border-radius:8px;border:1px solid rgba(34,197,94,.35);
          background:rgba(34,197,94,.1);color:#16a34a;font-size:.78rem;font-weight:700;cursor:pointer">✈️ ส่งงาน</button>
        <button onclick="_ordPrintDetail('${escPO}')"
          style="padding:7px 12px;border-radius:8px;border:1px solid var(--bc-card);
          background:var(--bc-div);color:var(--t1);font-size:.78rem;font-weight:700;cursor:pointer">🖨️ พิมพ์</button>
        <button onclick="Swal.close(); openEditOrder('${escPO}')"
          style="padding:7px 12px;border-radius:8px;border:1px solid rgba(99,102,241,.35);
          background:rgba(99,102,241,.1);color:#6366f1;font-size:.78rem;font-weight:700;cursor:pointer">✏️ แก้ไข</button>
        <button onclick="_ordMoreMenu('${escPO}','${escQuo}')"
          style="padding:7px 10px;border-radius:8px;border:1px solid var(--bc-card);
          background:var(--bc-div);color:var(--t1);font-size:.78rem;font-weight:700;cursor:pointer">⋯</button>
      </div>
    </div>`;

  Swal.fire({
    title: '📋 รายละเอียด Order',
    html, width: 880,
    background: 'var(--bg-card)', color: 'var(--t1)',
    showConfirmButton: true,
    confirmButtonText: '✕ ปิด', confirmButtonColor: '#475569'
  });
}

// ── ดูรูปขยาย ──
function _ordZoomImage(url, label) {
  Swal.fire({
    title: label, imageUrl: url, imageAlt: label,
    width: 'min(92vw, 720px)',
    background: 'var(--bg-card)', color: 'var(--t1)',
    confirmButtonText: 'ปิด', confirmButtonColor: '#6366f1'
  });
}

// ── เปลี่ยนสถานะงาน (Process) แบบเร็ว จากการ์ดรายละเอียด ──
async function _ordQuickChangeProcess(noPO) {
  const r = _orderCache.find(row => String(row[ORDER_COLS.noPO]) === String(noPO));
  if (!r) return;
  const options = {};
  ['กำลังผลิต','ส่งซุป','ส่งตัวอย่างเทส+รอสรุป','ส่งยังไม่ครบ','รอConfirm','FG รอเรียก','Stock','เรียบร้อย']
    .forEach(v => options[v] = v);

  const { value: newProcess } = await Swal.fire({
    title: '🔄 เปลี่ยนสถานะงาน (Process)',
    input: 'select',
    inputOptions: options,
    inputValue: r[ORDER_COLS.process] || '',
    showCancelButton: true,
    confirmButtonText: '✅ บันทึก', cancelButtonText: 'ยกเลิก',
    background: 'var(--bg-card)', color: 'var(--t1)',
    confirmButtonColor: '#6366f1', cancelButtonColor: '#475569'
  });
  if (!newProcess) return;

  const row = r.slice();
  while (row.length < ORDER_NUM_COLS) row.push('');
  row[ORDER_COLS.process] = newProcess;

  Swal.fire({title:'กำลังบันทึก…', background:'var(--bg-card)', color:'var(--t1)', allowOutsideClick:false, showConfirmButton:false, didOpen:()=>Swal.showLoading()});
  try {
    await fetch(SCRIPT_URL, { method:'POST', mode:'no-cors',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ action:'updateOrder', noPO, row }) });
    await fetchOrders();
    Swal.fire({icon:'success', title:'บันทึกแล้ว ✅', background:'var(--bg-card)', color:'var(--t1)',
      confirmButtonColor:'#6366f1', timer:1100, showConfirmButton:false});
    setTimeout(() => showOrderDetail(noPO), 900);
  } catch (err) {
    Swal.fire({icon:'error', title:'เกิดข้อผิดพลาด', text:'บันทึกไม่สำเร็จ', background:'var(--bg-card)', color:'var(--t1)', confirmButtonColor:'#6366f1'});
  }
}

// ── ตั้งสถานะส่งงาน = "ส่งแล้ว" + บันทึกวันที่ส่งงาน ──
async function _ordMarkDelivered(noPO) {
  const r = _orderCache.find(row => String(row[ORDER_COLS.noPO]) === String(noPO));
  if (!r) return;
  const res = await Swal.fire({
    icon:'question', title:'ยืนยันส่งงาน?',
    html:`No.PO: <b>${noPO}</b><br><span style="font-size:.8rem;color:var(--t3)">จะตั้งสถานะส่งงาน = "ส่งแล้ว" และบันทึกวันที่ส่งงานวันนี้</span>`,
    showCancelButton:true, confirmButtonText:'✅ ยืนยัน', cancelButtonText:'ยกเลิก',
    background:'var(--bg-card)', color:'var(--t1)',
    confirmButtonColor:'#22c55e', cancelButtonColor:'#475569'
  });
  if (!res.isConfirmed) return;

  const row = r.slice();
  while (row.length < ORDER_NUM_COLS) row.push('');
  row[ORDER_COLS.statusDeliver] = 'ส่งแล้ว';
  row[ORDER_COLS.update] = _ordDateToSheet(_todayStr());

  Swal.fire({title:'กำลังบันทึก…', background:'var(--bg-card)', color:'var(--t1)', allowOutsideClick:false, showConfirmButton:false, didOpen:()=>Swal.showLoading()});
  try {
    await fetch(SCRIPT_URL, { method:'POST', mode:'no-cors',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ action:'updateOrder', noPO, row }) });
    await fetchOrders();
    Swal.fire({icon:'success', title:'ส่งงานแล้ว ✅', background:'var(--bg-card)', color:'var(--t1)',
      confirmButtonColor:'#6366f1', timer:1100, showConfirmButton:false});
    setTimeout(() => showOrderDetail(noPO), 900);
  } catch (err) {
    Swal.fire({icon:'error', title:'เกิดข้อผิดพลาด', text:'บันทึกไม่สำเร็จ', background:'var(--bg-card)', color:'var(--t1)', confirmButtonColor:'#6366f1'});
  }
}

// ── พิมพ์การ์ดรายละเอียด Order ──
function _ordPrintDetail(noPO) {
  const r = _orderCache.find(row => String(row[ORDER_COLS.noPO]) === String(noPO));
  if (!r) return;
  const g = (k) => String(r[ORDER_COLS[k]] ?? '').trim() || '—';
  const price = parseFloat(r[ORDER_COLS.price]) || 0;
  const win = window.open('', '_blank');
  if (!win) return;
  win.document.write(`
    <html><head><title>Order ${g('noPO')}</title>
    <style>
      body{font-family:'Sarabun',Tahoma,sans-serif;padding:24px;color:#111}
      h2{margin:0 0 4px}
      .row{display:flex;gap:24px;margin-bottom:6px;flex-wrap:wrap}
      .lbl{font-size:.75rem;color:#666}
      .val{font-size:1rem;font-weight:700}
      table{width:100%;border-collapse:collapse;margin-top:12px}
      td{padding:6px 8px;border:1px solid #ccc;font-size:.85rem;vertical-align:top}
      td.lbl2{background:#f3f4f6;font-weight:700;width:160px}
    </style></head><body>
    <h2>ใบสั่งงาน Order</h2>
    <div class="row">
      <div><div class="lbl">No.Quo</div><div class="val">${g('noQuo')}</div></div>
      <div><div class="lbl">No.PO</div><div class="val">${g('noPO')}</div></div>
      <div><div class="lbl">ลูกค้า</div><div class="val">${g('customer')}</div></div>
    </div>
    <table>
      <tr><td class="lbl2">วันที่สั่งซื้อ</td><td>${g('orderDate')}</td><td class="lbl2">วันที่ต้องการ</td><td>${g('wantDate')}</td></tr>
      <tr><td class="lbl2">แม่พิมพ์</td><td>${g('mold')}</td><td class="lbl2">แบบงาน</td><td>${g('workType')}</td></tr>
      <tr><td class="lbl2">รายการสินค้า</td><td>${g('productList')}</td><td class="lbl2">จำนวน</td><td>${g('qty')}</td></tr>
      <tr><td class="lbl2">วัตถุดิบ</td><td>${g('material')}</td><td class="lbl2">ราคาเสนอขาย</td><td>${price ? price.toLocaleString('th-TH',{minimumFractionDigits:2})+' บาท' : '—'}</td></tr>
      <tr><td class="lbl2">Process</td><td>${g('process')}</td><td class="lbl2">สถานะส่งงาน</td><td>${g('statusDeliver')}</td></tr>
      <tr><td class="lbl2">หมายเหตุ</td><td colspan="3">${g('note')}</td></tr>
    </table>
    </body></html>`);
  win.document.close();
  win.focus();
  setTimeout(() => win.print(), 400);
}

// ── เมนู "เพิ่มเติม" ──
function _ordMoreMenu(noPO, noQuo) {
  Swal.fire({
    title: '⋯ เพิ่มเติม',
    html: `<div style="display:flex;flex-direction:column;gap:8px;text-align:left">
      <button onclick="deleteOrderRow('${noPO}','${noQuo}')"
        style="padding:9px 12px;border-radius:8px;border:1px solid rgba(248,113,113,.35);
        background:rgba(248,113,113,.1);color:#f87171;font-size:.82rem;font-weight:700;cursor:pointer;text-align:left">🗑️ ลบ Order นี้</button>
    </div>`,
    showConfirmButton:false, showCloseButton:true,
    background:'var(--bg-card)', color:'var(--t1)'
  });
}

function openEditOrder(noPO) {
  const r = _orderCache.find(row => String(row[ORDER_COLS.noPO]) === String(noPO));
  if (!r) return;
  _addSeen(SEEN_KEY_ORDER, noPO);
  renderOrderTable();
  _ordEditNoPO = noPO;
  $('ordEdit_noPO').textContent  = noPO;
  $('ordEdit_orderDate').value   = _ordDateToInput(r[ORDER_COLS.orderDate]);
  $('ordEdit_wantDate').value    = _ordDateToInput(r[ORDER_COLS.wantDate]);
  $('ordEdit_status').value      = r[ORDER_COLS.process] || '';
  $('ordEdit_qty').value         = r[ORDER_COLS.qty] || '';
  $('ordEdit_price').value       = r[ORDER_COLS.price] || '';
  $('ordEdit_productList').value = r[ORDER_COLS.productList] || '';
  $('ordEdit_material').value    = r[ORDER_COLS.material] || '';
  $('ordEdit_note').value        = r[ORDER_COLS.note] || '';
  if ($('ordEdit_workStatus')) $('ordEdit_workStatus').value = r[ORDER_COLS.status] || 'ปรกติ';
  _ordClearPoFile('ordEdit');
  const poUrl = r[ORDER_COLS.poFile] || '';
  if ($('ordEdit_poLinkWrap')) {
    $('ordEdit_poLinkWrap').innerHTML = poUrl
      ? `<a href="${poUrl}" target="_blank" rel="noopener">📎 ดูไฟล์ PO เดิม</a>`
      : '';
  }
  $('orderEditModal').style.display = 'flex';
}
function closeOrderEdit() {
  $('orderEditModal').style.display = 'none';
  _ordEditNoPO = null;
}
async function saveOrderEdit() {
  if (!_ordEditNoPO) return;
  const r = _orderCache.find(row => String(row[ORDER_COLS.noPO]) === String(_ordEditNoPO));
  if (!r) return;
  const row = r.slice();
  while (row.length < ORDER_NUM_COLS) row.push('');
  row[ORDER_COLS.orderDate]   = _ordDateToSheet($('ordEdit_orderDate').value);
  row[ORDER_COLS.wantDate]    = _ordDateToSheet($('ordEdit_wantDate').value);
  row[ORDER_COLS.process]     = $('ordEdit_status').value;
  row[ORDER_COLS.status]      = $('ordEdit_workStatus')?.value || row[ORDER_COLS.status] || 'ปรกติ';
  row[ORDER_COLS.qty]         = $('ordEdit_qty').value;
  row[ORDER_COLS.price]       = $('ordEdit_price').value;
  row[ORDER_COLS.productList] = $('ordEdit_productList').value;
  row[ORDER_COLS.material]    = $('ordEdit_material').value;
  row[ORDER_COLS.note]        = $('ordEdit_note').value;
  row[ORDER_COLS.poFile]      = r[ORDER_COLS.poFile] || '';

  const saveBtn  = $('ordEdit_saveBtn');
  const statusEl = $('ordEdit_status_msg');
  if (saveBtn) saveBtn.disabled = true;

  try {
    const poFile = $('ordEdit_poFile')?.files?.[0];
    if (poFile) {
      if (statusEl) statusEl.textContent = '⏳ กำลังอัปโหลดไฟล์ PO...';
      row[ORDER_COLS.poFile] = await _ordUploadPoFile(poFile);
    }

    await fetch(SCRIPT_URL, { method:'POST', mode:'no-cors',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ action:'updateOrder', noPO: _ordEditNoPO, row }) });

    Swal.fire({icon:'success',title:'บันทึกแล้ว ✅',background:'#0d1b2a',color:'#cce4ff',
      confirmButtonColor:'#6366f1', timer:1300, showConfirmButton:false});
    if (statusEl) statusEl.textContent = '';
    closeOrderEdit();
    setTimeout(fetchOrders, 1000);
  } catch (err) {
    Swal.fire({icon:'error',title:'เกิดข้อผิดพลาด',text:'บันทึกไม่สำเร็จ',background:'#0d1b2a',color:'#cce4ff',confirmButtonColor:'#6366f1'});
    if (statusEl) statusEl.textContent = '';
  } finally {
    if (saveBtn) saveBtn.disabled = false;
  }
}
      