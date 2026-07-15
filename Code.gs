// ═══════════════════════════════════════════════════════════════
//  PTTS Cost Breakdown — Google Apps Script
//  วางใน Extensions → Apps Script ของ Google Sheet แล้ว Deploy
// ═══════════════════════════════════════════════════════════════

var SHEET_NAME = 'DATA';   // ← ชื่อ tab ที่ต้องการบันทึก

// ── doPost: รับข้อมูลจากเว็บฟอร์ม ────────────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.getActiveSpreadsheet();

    // ── deleteRow ────────────────────────────────────────────────
    if (data.action === 'deleteRow') return deleteRow(data.noQuo);

    // ── updateRow ────────────────────────────────────────────────
    if (data.action === 'updateRow') return updateRow(data.noQuo, data.row);

    // ── addMold ──────────────────────────────────────────────────
    if (data.action === 'addMold') {
      var shM = ss.getSheetByName('Modl');
      if (!shM) return jsonOut({ status:'error', message:'ไม่พบ sheet Modl' });
      var row = [data.od].concat(data.ids || []);
      shM.appendRow(row);
      return jsonOut({ status:'ok', message:'เพิ่ม OD ' + data.od + ' แล้ว' });
    }

    // ── updateMold ───────────────────────────────────────────────
    if (data.action === 'updateMold') {
      var shU = ss.getSheetByName('Modl');
      if (!shU) return jsonOut({ status:'error', message:'ไม่พบ sheet Modl' });
      var lastU = shU.getLastRow();
      for (var r = 2; r <= lastU; r++) {
        var odVal = parseFloat(shU.getRange(r, 1).getValue());
        if (Math.abs(odVal - parseFloat(data.od)) < 0.01) {
          shU.getRange(r, 1, 1, Math.max(shU.getLastColumn(), (data.ids||[]).length + 1)).clearContent();
          var newRow = [data.od].concat(data.ids || []);
          shU.getRange(r, 1, 1, newRow.length).setValues([newRow]);
          return jsonOut({ status:'ok', message:'อัปเดต OD ' + data.od + ' แล้ว' });
        }
      }
      return jsonOut({ status:'error', message:'ไม่พบ OD ' + data.od });
    }

    // ── deleteMold ───────────────────────────────────────────────
    if (data.action === 'deleteMold') {
      var shD = ss.getSheetByName('Modl');
      if (!shD) return jsonOut({ status:'error', message:'ไม่พบ sheet Modl' });
      var lastD = shD.getLastRow();
      for (var d = 2; d <= lastD; d++) {
        var odD = parseFloat(shD.getRange(d, 1).getValue());
        if (Math.abs(odD - parseFloat(data.od)) < 0.01) {
          shD.deleteRow(d);
          return jsonOut({ status:'ok', message:'ลบ OD ' + data.od + ' แล้ว' });
        }
      }
      return jsonOut({ status:'error', message:'ไม่พบ OD ' + data.od });
    }

    // ── saveLaborConfig: บันทึกอัตราค่าแรง + กระบวนการผลิต (sync ทุกเครื่อง) ──
    if (data.action === 'saveLaborConfig') return saveLaborConfig(data);

    // ── saveDrafts: บันทึกแบบร่างทั้งหมด (sync ทุกเครื่อง) ──
    if (data.action === 'saveDrafts') return saveDrafts(data);

    // ── savePlatingNoNeed: บันทึกรายการ "ไม่ต้องชุบ" (sync ทุกเครื่อง) ──
    if (data.action === 'savePlatingNoNeed') return savePlatingNoNeed(data);

    // ── savePlatingPriceMemory: บันทึกความจำราคาใบส่งชุบ (sync ทุกเครื่อง) ──
    if (data.action === 'savePlatingPriceMemory') return savePlatingPriceMemory(data);

    // ── updateMat: เพิ่ม/แก้ไขราคาวัตถุดิบใน sheet MAT-ฝา / MAT-ตะแกรง ──
    if (data.action === 'updateMat') return updateMat(data);

    // ── deleteMat: ลบรายการวัตถุดิบออกจาก sheet MAT-ฝา / MAT-ตะแกรง ──
    if (data.action === 'deleteMat') return deleteMat(data);

    // ── addOrder: สร้าง Order ใหม่ใน sheet "Order" (ใช้ร่วมกับ AppSheet) ──
    if (data.action === 'addOrder') return addOrder(data);

    // ── updateOrder: แก้ไข/อัปเดตสถานะ Order (ค้นหาด้วย No.PO) ──
    if (data.action === 'updateOrder') return updateOrder(data.noPO, data.row);

    // ── deleteOrder: ลบแถวใน sheet "Order" ตาม No.PO ──
    if (data.action === 'deleteOrder') return deleteOrder(data.noPO);

    // ── updateDataStatus: ตั้งคอลัมน์ "สถานะ" (A) ของ DATA ตาม No.Quo (B) เป็นค่าใหม่ ──
    if (data.action === 'updateDataStatus') return updateDataStatus(data.noQuo, data.status);

    // ── uploadOrderFile: อัปโหลดไฟล์แนบ PO ขึ้น Drive ─────────────
    if (data.action === 'uploadOrderFile') return uploadOrderFile(data);

    // ── savePurchaseOrder: สร้าง/แก้ไขใบสั่งซื้อ (header + items) ──
    if (data.action === 'savePurchaseOrder') return savePurchaseOrder(data);

    if (data.action === 'deletePOSupplierItem') return deletePOSupplierItem(data);

    // ── deletePurchaseOrder: ลบใบสั่งซื้อ (header + items) ──
    if (data.action === 'deletePurchaseOrder') return deletePurchaseOrder(data.poNo);

    // ── saveCompanyInfo: บันทึกข้อมูลบริษัท (sheet "Company") ──
    if (data.action === 'saveCompanyInfo') return saveCompanyInfo(data);

    // ── uploadLogo: อัปโหลดโลโก้บริษัทขึ้น Drive ─────────────────
    if (data.action === 'uploadLogo') return uploadLogo(data);
    if (data.action === 'uploadPOItemImage') return uploadPOItemImage(data);

    // ── saveCustomer: เพิ่ม/แก้ไขลูกค้า (sheet "Customers") ──────
    if (data.action === 'saveCustomer') return saveCustomer(data);

    // ── deleteCustomer: ลบลูกค้า (sheet "Customers") ─────────────
    if (data.action === 'deleteCustomer') return deleteCustomer(data.code);

    // ── saveSupplier: เพิ่ม/แก้ไข Supplier (sheet "Suppliers") ───
    if (data.action === 'saveSupplier') return saveSupplier(data);

    // ── deleteSupplier: ลบ Supplier (sheet "Suppliers") ──────────
    if (data.action === 'deleteSupplier') return deleteSupplier(data.code);

    // ── saveInvoiceRecord: ออกเลขที่ใบกำกับภาษี + บันทึก log ────
    if (data.action === 'saveInvoiceRecord') return saveInvoiceRecord(data);

    // ── updateInvoiceRecord: แก้ไขรายละเอียดใบกำกับที่ออกไปแล้ว ──
    if (data.action === 'updateInvoiceRecord') return updateInvoiceRecord(data);

    // ── deleteInvoiceRecord: ลบใบกำกับที่ออกไปแล้ว ──────────────
    if (data.action === 'deleteInvoiceRecord') return deleteInvoiceRecord(data.invoiceNo, data.revertProcess);

    // ── stampOrderPO: ประทับ invoiceNo/invoiceDate ลง col AC/AD ใน Order ──
    // data: { noPOs:[...], invoiceNo, invoiceDate }
    if (data.action === 'stampOrderPO') return stampOrderPO(data);

    if (data.action === 'saveBillingNote') return saveBillingNote(data);
    if (data.action === 'deleteBillingNote') return deleteBillingNote(data);

    // ── ใบส่งชุบ (PlatingNote) ──────────────────────────────────
    if (data.action === 'savePlatingNote') return savePlatingNote(data);
    if (data.action === 'deletePlatingNote') return deletePlatingNote(data);
    if (data.action === 'updatePlatingNote') return updatePlatingNote(data);

    // ── saveItem / deleteItem: รายการสินค้า/บริการที่ใช้บ่อย (sheet "ItemMaster") ──
    if (data.action === 'saveItem') return saveItem(data);
    if (data.action === 'deleteItem') return deleteItem(data.code);

    // ── sendChatMessage: แชททีม (sheet "Chat") ─────────────────────
    if (data.action === 'sendChatMessage') return sendChatMessage(data);

    // ── saveExpenseReceipt: ใบเสร็จรายจ่าย ───────────────────────
    if (data.action === 'saveExpenseReceipt') {
      var expSS = SpreadsheetApp.getActiveSpreadsheet();
      var expSh = expSS.getSheetByName('ExpenseReceipts');
      if (!expSh) {
        expSh = expSS.insertSheet('ExpenseReceipts');
        expSh.appendRow(['ref','date','category','payMethod','vendorName','vendorTaxId','vendorAddress','items','total','note','createdAt']);
      }
      var d = data.data || {};
      expSh.appendRow([
        d.ref||'', d.date||'', d.category||'', d.payMethod||'',
        d.vendorName||'', d.vendorTaxId||'', d.vendorAddress||'',
        JSON.stringify(d.items||[]), d.total||0, d.note||'',
        new Date().toLocaleString('th-TH')
      ]);
      return jsonOut({ status:'ok' });
    }

    // ── deleteExpenseReceipt: ลบใบเสร็จรายจ่าย ───────────────────
    if (data.action === 'deleteExpenseReceipt') {
      var delSS = SpreadsheetApp.getActiveSpreadsheet();
      var delSh = delSS.getSheetByName('ExpenseReceipts');
      if (!delSh) return jsonOut({ status:'ok' });
      var delRows = delSh.getDataRange().getValues();
      for (var di = delRows.length - 1; di >= 1; di--) {
        if (String(delRows[di][0]) === String(data.ref)) {
          delSh.deleteRow(di + 1);
          break;
        }
      }
      return jsonOut({ status:'ok' });
    }

    // ── saveGeneralQuotation: ใบเสนอราคาทั่วไป → ชีต Quotations ──
    if (data.action === 'saveGeneralQuotation') return saveGeneralQuotation(data);
    if (data.action === 'deleteGeneralQuotation') return deleteGeneralQuotation(data);

    // ── saveRFQ: สร้าง/แก้ไขใบขอราคา (sheet "RFQ") ──
    if (data.action === 'saveRFQ') return saveRFQ(data);

    // ── deleteRFQ: ลบใบขอราคา ────────────────────────────────────
    if (data.action === 'deleteRFQ') return deleteRFQ(data);

    // ── WI: บันทึก/ลบ Work Instruction ──────────────────────────
    if (data.action === 'uploadWIImage') return uploadWIImage(data);
    if (data.action === 'saveWI') return saveWI(data);
    if (data.action === 'deleteWI') return deleteWI(data);
    if (data.action === 'setOrderWI')     return setOrderWI(data);
    if (data.action === 'saveMatStock')        return saveMatStock(data);
    if (data.action === 'updateMatStockBatch') return updateMatStockBatch(data);
    if (data.action === 'adjustMatStockBatch') return adjustMatStockBatch(data);

    if (data.action === 'saveInspection')   return saveInspection(data);
    if (data.action === 'updateInspection') return updateInspection(data);
    if (data.action === 'deleteInspection') return deleteInspection(data.inspId);

    // ── HR Module ──────────────────────────────────────────────────
    if (data.action === 'saveHREmployee')   return saveHREmployee(data);
    if (data.action === 'deleteHREmployee') return deleteHREmployee(data);
    if (data.action === 'saveHRAttendance') return saveHRAttendance(data);
    if (data.action === 'seedHRHolidays')    return seedHRHolidays(data);
    if (data.action === 'deleteHRHoliday')  return deleteHRHoliday(data);
    if (data.action === 'hrLogin')          return hrLogin(data);
    if (data.action === 'saveLoanRequest')  return saveLoanRequest(data);
    if (data.action === 'approveLoan')      return approveLoan(data);
    if (data.action === 'rejectLoan')       return rejectLoan(data);
    if (data.action === 'closeLoan')        return closeLoan(data);
    if (data.action === 'deleteLoanRequest') return deleteLoanRequest(data);
    if (data.action === 'markAdvanceDeducted') return markAdvanceDeducted(data);
    if (data.action === 'saveHRLoanContract')  return saveHRLoanContract(data);
    if (data.action === 'recordHRLoanPayment') return recordHRLoanPayment(data);
    if (data.action === 'recordSalaryPayment') return recordSalaryPayment(data);
    if (data.action === 'getSalaryPayments')   return getSalaryPayments(e.parameter);
    if (data.action === 'deleteSalaryPayment')  return deleteSalaryPayment(data);
    if (data.action === 'payrollDeductLoans')   return payrollDeductLoans(data);
    if (data.action === 'uploadSlipImage')   return uploadSlipImage(data);
    if (data.action === 'uploadEmployeeFile') return uploadEmployeeFile(data);
    if (data.action === 'saveHRPayCfg')        return saveHRPayCfg(data);
    if (data.action === 'addHRSalaryHistory')  return addHRSalaryHistory(data);

    // ── default: บันทึก Cost row ──────────────────────────────────
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return jsonOut({ status:'error', message:'ไม่พบ sheet ' + SHEET_NAME });
    sheet.appendRow(data.row);
    scheduleTelegram();   // ยิง Telegram แบบ async ไม่บล็อก response
    return jsonOut({ status:'ok', message:'บันทึกสำเร็จ' });

  } catch (err) {
    return jsonOut({ status:'error', message: err.toString() });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── doGet: Health check + data fetching ──────────────────────
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;

  if (action === 'getCosts') {
    var rawLim = e.parameter.limit;
    var lim = (rawLim !== undefined && rawLim !== '') ? (parseInt(rawLim) || 0) : 200;  // default 200, &limit=0 = ทั้งหมด
    return getCosts(lim);
  }

  if (action === 'getOrders') {
    var rawLimO = e.parameter.limit;
    var limO = (rawLimO !== undefined && rawLimO !== '') ? (parseInt(rawLimO) || 0) : 0; // default 0 = ทั้งหมด
    return getOrders(limO);
  }

  if (action === 'getOrderPoImage') {
    return getOrderPoImage(e.parameter.path || '', e.parameter.noPO || '');
  }

  if (action === 'getOrderQty') {
    return getOrderQty(e.parameter.noPO || '');
  }

  if (action === 'getPurchaseOrders') return getPurchaseOrders();

  if (action === 'getSuppliers') return getSuppliers();

  if (action === 'getPOSupplierItems') return getPOSupplierItems();

  if (action === 'getNextPONo') return getNextPONo();

  if (action === 'getNextNo') {
    try {
      var ss2     = SpreadsheetApp.getActiveSpreadsheet();
      var dataSheet = ss2.getSheetByName(SHEET_NAME);
      var lastRow2  = dataSheet ? dataSheet.getLastRow() : 1;
      var nextNo    = 1;
      if (lastRow2 > 1) {
        var lastVal = String(dataSheet.getRange(lastRow2, 2).getValue() || '');
        var parsed  = parseInt(lastVal.replace(/^Q-/i, ''), 10);
        nextNo = isNaN(parsed) ? (lastRow2 - 1 + 1) : (parsed + 1);
      }
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', nextNo: nextNo }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'getMAT') {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();

      function readMatData(sheetName) {
        var sh = ss.getSheetByName(sheetName);
        if (!sh) return [];
        var last = sh.getLastRow();
        if (last < 3) return [];
        var vals = sh.getRange(3, 2, last - 2, 7).getValues(); // B..H
        return vals.map(function(r) {
          var code     = String(r[0]).trim();
          var name     = String(r[1]).trim();     // col C = รายการวัตถุดิบ
          var price    = parseFloat(r[3]) || 0;  // col E = ราคา+30%
          var priceBuy = parseFloat(r[4]) || 0;  // col F = ราคาซื้อจริง
          var w        = parseFloat(r[5]) || 0;  // col G = กว้าง (มิล)
          var l        = parseFloat(r[6]) || 0;  // col H = ยาว (มิล)
          return code ? { code: code, name: name, price: price, priceBuy: priceBuy, w: w, l: l } : null;
        }).filter(Boolean);
      }

      var matFlap = readMatData('MAT-ฝา');
      var matMesh = readMatData('MAT-ตะแกรง');

      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', matFlap: matFlap, matMesh: matMesh }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'getSpecMat') {
    try {
      var ss3 = SpreadsheetApp.getActiveSpreadsheet();
      var specSh = ss3.getSheetByName('Spec Mat');
      if (!specSh) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'error', message: 'ไม่พบ sheet ชื่อ Spec Mat' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var last3 = specSh.getLastRow();
      var specs = {};
      if (last3 >= 4) {
        var sv = specSh.getRange(4, 1, last3 - 3, 3).getValues();
        sv.forEach(function(r) {
          var name = String(r[0]).trim();
          var w = parseFloat(r[1]) || 0;
          var l = parseFloat(r[2]) || 0;
          if (name && w && l) specs[name] = { w: w, l: l };
        });
      }
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', specs: specs }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'getContacts') {
    try {
      var ssC = SpreadsheetApp.getActiveSpreadsheet();
      var shC = ssC.getSheetByName('ชื่อผู้ติดต่อ');
      if (!shC) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'error', message: 'ไม่พบ sheet ชื่อผู้ติดต่อ' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var lastC = shC.getLastRow();
      var contacts = [];
      if (lastC >= 2) {
        var valsC = shC.getRange(2, 1, lastC - 1, 2).getValues();
        contacts = valsC.map(function(r) {
          var name = String(r[0]).trim();
          var co   = String(r[1]).trim();
          return name ? { name: name, company: co } : null;
        }).filter(Boolean);
      }
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', contacts: contacts }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'getWorkTypes') {
    try {
      var ssW  = SpreadsheetApp.getActiveSpreadsheet();
      var shW  = ssW.getSheetByName('แบบงาน');
      if (!shW) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'error', message: 'ไม่พบ sheet ชื่อ แบบงาน' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var lastW = shW.getLastRow();
      var types = [];
      if (lastW >= 2) {
        var valsW = shW.getRange(2, 2, lastW - 1, 1).getValues();
        types = valsW.map(function(r) { return String(r[0]).trim(); }).filter(Boolean);
      }
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', types: types }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'getModl') {
    try {
      var ssM = SpreadsheetApp.getActiveSpreadsheet();
      var shM = ssM.getSheetByName('Modl');
      if (!shM) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'error', message: 'ไม่พบ sheet ชื่อ Modl' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var lastM = shM.getLastRow();
      var lastCol = shM.getLastColumn();
      var molds = [];
      if (lastM >= 2 && lastCol >= 1) {
        var valsM = shM.getRange(2, 1, lastM - 1, lastCol).getValues();
        valsM.forEach(function(r) {
          var od = parseFloat(r[0]);
          if (isNaN(od)) return;
          var ids = [];
          for (var c = 1; c < r.length; c++) {
            var v = String(r[c]).trim();
            if (v && v !== '' && v !== '0') ids.push(v);
          }
          molds.push({ od: od, ids: ids });
        });
      }
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', molds: molds }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'getDrafts') {
    try {
      var ssD = SpreadsheetApp.getActiveSpreadsheet();
      var shD2 = ssD.getSheetByName('Drafts');
      if (!shD2) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'ok', drafts: [] }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var rawD = shD2.getRange('B1').getValue();
      var drafts = rawD ? JSON.parse(rawD) : [];
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', drafts: drafts }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'getPlatingNoNeed') return getPlatingNoNeed();

  if (action === 'getPlatingPriceMemory') return getPlatingPriceMemory();

  if (action === 'getCompanyInfo') return getCompanyInfo();

  if (action === 'getCustomers') return getCustomers();

  if (action === 'getItemMaster') return getItemMaster();

  if (action === 'getNextInvoiceNo') {
    var typ = e.parameter.type || 'full';
    return jsonOut({ status: 'ok', nextNo: _peekNextInvoiceNo(typ) });
  }

  if (action === 'getInvoices') return getInvoices();

  if (action === 'getNextBillNo') return jsonOut({ status: 'ok', billNo: _peekNextBillNo() });

  if (action === 'getChatMessages') return getChatMessages();

  if (action === 'runBackupNow') {
    try {
      var bkResult = backupSpreadsheet();
      return jsonOut({ status: 'ok', backup: bkResult });
    } catch (e) {
      return jsonOut({ status: 'error', message: e.message });
    }
  }

  if (action === 'getBillingNotes') return getBillingNotes();

  if (action === 'getNextPlatingNo') return jsonOut({ status: 'ok', platingNo: _peekNextPlatingNo() });

  if (action === 'getPlatingNotes') return getPlatingNotes();

  if (action === 'getRFQList') return getRFQList();

  if (action === 'getWIList') return getWIList();

  if (action === 'getSalesDashboard') {
    var months = parseInt(e.parameter.months) || 6;
    return getSalesDashboard(months);
  }

  if (action === 'getMatStock')    return getMatStock();
  if (action === 'getMatMovement') return getMatMovement(e.parameter.matCode || '');

  if (action === 'getHREmployees')  return getHREmployees();
  if (action === 'getHRAttendance') return getHRAttendance(e.parameter);
  if (action === 'getHRHolidays')   return getHRHolidays();
  if (action === 'getHRPayCfg')        return getHRPayCfg();
  if (action === 'getHRSalaryHistory') return getHRSalaryHistory(e.parameter);
  if (action === 'getLoanRequests')    return getLoanRequests(e.parameter);
  if (action === 'getHRLoanContracts') return getHRLoanContracts(e.parameter);
  if (action === 'getHRLoanPayments')  return getHRLoanPayments(e.parameter);
  if (action === 'getSalaryPayments')  return getSalaryPayments(e.parameter);

  if (action === 'getInspections') {
    var ispType    = e.parameter.type    || '';
    var ispOrderId = e.parameter.orderId || '';
    return getInspections(ispType, ispOrderId);
  }

  if (action === 'getInspectionDetail') {
    var ispId = e.parameter.inspId || '';
    return getInspectionDetail(ispId);
  }

  if (action === 'getGeneralQuotations') {
    try {
      var gqSS = SpreadsheetApp.getActiveSpreadsheet();
      var gqSh = gqSS.getSheetByName('Quotations');
      if (!gqSh) return jsonOut({ data: [] });
      var gqRows = gqSh.getDataRange().getValues();
      if (gqRows.length <= 1) return jsonOut({ data: [] });
      var gqHeaders = gqRows[0];
      var gqData = gqRows.slice(1).map(function(r) {
        var obj = {};
        gqHeaders.forEach(function(h, i) { obj[h] = r[i]; });
        try { if (typeof obj.items === 'string') obj.items = JSON.parse(obj.items || '[]'); } catch(e) { obj.items = []; }
        // format date: Date object → ISO string
        if (obj.date instanceof Date) {
          obj.date = Utilities.formatDate(obj.date, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
        } else {
          obj.date = String(obj.date || '');
        }
        return obj;
      });
      return jsonOut({ data: gqData });
    } catch(err) {
      return jsonOut({ data: [], error: err.toString() });
    }
  }

  if (action === 'getExpenseReceipts') {
    try {
      var expGS = SpreadsheetApp.getActiveSpreadsheet();
      var expGSh = expGS.getSheetByName('ExpenseReceipts');
      if (!expGSh) return jsonOut({ data: [] });
      var expRows = expGSh.getDataRange().getValues();
      if (expRows.length <= 1) return jsonOut({ data: [] });
      var expHeaders = expRows[0];
      var expData = expRows.slice(1).map(function(r) {
        var obj = {};
        expHeaders.forEach(function(h, i) { obj[h] = r[i]; });
        try { if (typeof obj.items === 'string') obj.items = JSON.parse(obj.items || '[]'); } catch(e) { obj.items = []; }
        return obj;
      });
      return jsonOut({ data: expData });
    } catch(err) {
      return jsonOut({ data: [], error: err.toString() });
    }
  }

  if (action === 'getLaborConfig') {
    try {
      var ssL = SpreadsheetApp.getActiveSpreadsheet();
      var shL = ssL.getSheetByName('Config');
      if (!shL) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'empty' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var raw = shL.getRange('B1').getValue();
      if (!raw) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'empty' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var cfg = JSON.parse(raw);
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', config: cfg }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'PTTS Cost Breakdown API พร้อมใช้งาน' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── getCosts: ดึงข้อมูลจาก sheet DATA (limit = rows ล่าสุด, 0 = ทั้งหมด)
function getCosts(limit) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sh    = ss.getSheetByName('DATA') || ss.getSheets()[0];
    var last  = sh.getLastRow();
    var total = Math.max(0, last - 1);
    if (total === 0) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', rows: [], total: 0 }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var lim      = (limit && limit > 0 && limit < total) ? limit : total;
    var startRow = last - lim + 1;
    var values   = sh.getRange(startRow, 1, lim, sh.getLastColumn()).getValues();
    var rows = values.map(function(r) {
      return r.map(function(cell) {
        if (cell instanceof Date) {
          return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        }
        return cell;
      });
    });
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', rows: rows, total: total, limited: lim < total }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: e.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── updateRow: อัปเดตแถวที่ตรงกับ No.Quo ─────────────────────
function updateRow(noQuo, row) {
  try {
    if (!noQuo) throw new Error('noQuo is required');
    if (!Array.isArray(row) || row.length === 0) throw new Error('row data is required');
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var sh   = ss.getSheetByName('DATA') || ss.getSheets()[0];
    var last = sh.getLastRow();
    if (last < 2) {
      sh.appendRow(row);
      try { sendTelegramOnChangeManual(); } catch(tErr) { Logger.log('Telegram error: ' + tErr); }
      return jsonOut({ status: 'ok', updated: noQuo, note: 'appended' });
    }
    // อ่าน column B ทั้งหมดครั้งเดียว → ไม่ต้อง loop อ่านทีละ cell
    var colB = sh.getRange(2, 2, last - 1, 1).getValues(); // [[val],[val],...]
    var found = -1;
    for (var i = 0; i < colB.length; i++) {
      if (String(colB[i][0]) == String(noQuo)) { found = i + 2; break; }
    }
    if (found !== -1) {
      sh.getRange(found, 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    scheduleTelegram();   // ยิง Telegram แบบ async ไม่บล็อก response
    return jsonOut({ status: 'ok', updated: noQuo, note: found !== -1 ? 'updated' : 'appended' });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── updateDataStatus: อัปเดตคอลัมน์ "สถานะ" (A) ของ sheet DATA ──
// ค้นหาแถวที่ No.Quo (คอลัมน์ B) ตรงกับ noQuo เป๊ะๆ (รายการเดียว) แล้วตั้งค่าคอลัมน์ A = status
function updateDataStatus(noQuo, status) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!noQuo) throw new Error('noQuo is required');
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var sh   = ss.getSheetByName('DATA') || ss.getSheets()[0];
    var last = sh.getLastRow();
    if (last < 2) throw new Error('ไม่พบ ' + noQuo);
    var colB  = sh.getRange(2, 2, last - 1, 1).getValues();
    var found = -1, count = 0;
    for (var i = 0; i < colB.length; i++) {
      if (String(colB[i][0]).trim() === String(noQuo).trim()) { found = i + 2; count++; }
    }
    if (count === 0) throw new Error('ไม่พบ ' + noQuo);
    if (count > 1) throw new Error('พบ No.Quo ซ้ำ ' + count + ' แถว — ยกเลิกเพื่อความปลอดภัย');
    sh.getRange(found, 1).setValue(status || 'ผ่าน');
    return jsonOut({ status: 'ok', updated: noQuo, row: found });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── saveHRPayCfg: บันทึก pay cycle config ลง Config sheet (A=HRPayCfg, B=JSON) ──
function saveHRPayCfg(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Config');
    if (!sh) sh = ss.insertSheet('Config');
    // หา row ที่มี A = 'HRPayCfg' หรือสร้างแถวใหม่
    var lastRow = sh.getLastRow();
    var found = -1;
    for (var i = 1; i <= lastRow; i++) {
      if (sh.getRange(i, 1).getValue() === 'HRPayCfg') { found = i; break; }
    }
    var payload = JSON.stringify(data.cfg || {});
    if (found > 0) {
      sh.getRange(found, 2).setValue(payload);
    } else {
      sh.appendRow(['HRPayCfg', payload]);
    }
    return jsonOut({ status: 'ok' });
  } catch(e) { return jsonOut({ status: 'error', message: e.message }); }
}

// ── getHRPayCfg: อ่าน pay cycle config จาก Config sheet ──
function getHRPayCfg() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Config');
    if (!sh) return jsonOut({ status: 'ok', cfg: null });
    var lastRow = sh.getLastRow();
    for (var i = 1; i <= lastRow; i++) {
      if (sh.getRange(i, 1).getValue() === 'HRPayCfg') {
        var raw = sh.getRange(i, 2).getValue();
        try { return jsonOut({ status: 'ok', cfg: JSON.parse(raw || 'null') }); }
        catch(e2) { return jsonOut({ status: 'ok', cfg: null }); }
      }
    }
    return jsonOut({ status: 'ok', cfg: null });
  } catch(e) { return jsonOut({ status: 'error', message: e.message }); }
}

// ══════════════════════════════════════════════════════════════════
// HR_SalaryHistory — Phase 1: บันทึก/ดึงประวัติการปรับเงินเดือน
// ไม่แตะ _hrCalcPayslip เลย — เพิ่มข้อมูลอย่างเดียว
// Schema: empId | effectiveDate | salary | dailyRate | note | recordedAt
// ══════════════════════════════════════════════════════════════════

function _getSalaryHistorySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('HR_SalaryHistory');
  if (!sh) {
    sh = ss.insertSheet('HR_SalaryHistory');
    sh.appendRow(['empId','effectiveDate','salary','dailyRate','note','recordedAt']);
    sh.setFrozenRows(1);
  }
  return sh;
}

// POST action: addHRSalaryHistory
// body: { empId, effectiveDate (yyyy-MM-dd), salary, dailyRate, note }
function addHRSalaryHistory(data) {
  try {
    if (!data.empId || !data.effectiveDate || (!data.salary && !data.dailyRate)) {
      return jsonOut({ status: 'error', message: 'ข้อมูลไม่ครบ (empId, effectiveDate, salary หรือ dailyRate)' });
    }
    var sh = _getSalaryHistorySheet();
    var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
    sh.appendRow([
      String(data.empId).trim(),
      String(data.effectiveDate).trim(),   // yyyy-MM-dd จาก client
      Number(data.salary)   || 0,
      Number(data.dailyRate)|| 0,
      String(data.note     || '').trim(),
      now
    ]);
    return jsonOut({ status: 'ok' });
  } catch(e) { return jsonOut({ status: 'error', message: e.message }); }
}

// GET action: getHRSalaryHistory?empId=xxx  (ถ้าไม่ส่ง empId → คืนทั้งหมด)
function getHRSalaryHistory(params) {
  try {
    var sh = _getSalaryHistorySheet();
    var rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return jsonOut({ status: 'ok', data: [] });
    var filterEmp = (params && params.empId) ? String(params.empId).trim() : '';
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      var empId        = String(r[0]||'').trim();
      if (filterEmp && empId !== filterEmp) continue;
      // effectiveDate อาจเป็น Date object (Sheets auto-parse) → format เป็น yyyy-MM-dd เสมอ
      var effRaw = r[1];
      var effStr = '';
      if (effRaw instanceof Date) {
        effStr = Utilities.formatDate(effRaw, 'Asia/Bangkok', 'yyyy-MM-dd');
      } else {
        effStr = String(effRaw||'').trim();
      }
      result.push({
        empId:         empId,
        effectiveDate: effStr,
        salary:        Number(r[2])||0,
        dailyRate:     Number(r[3])||0,
        note:          String(r[4]||'').trim(),
        recordedAt:    String(r[5]||'').trim()
      });
    }
    // เรียงตาม effectiveDate ใหม่→เก่า
    result.sort(function(a,b){ return b.effectiveDate.localeCompare(a.effectiveDate); });
    return jsonOut({ status: 'ok', data: result });
  } catch(e) { return jsonOut({ status: 'error', message: e.message }); }
}

// ── saveLaborConfig: บันทึกอัตราค่าแรง + กระบวนการผลิต ลง sheet "Config" ──
// data: { ratePerMin, ratePerDay, hoursPerDay, processes: [...] }
// เก็บเป็น JSON string ไว้ที่ sheet "Config" เซลล์ B1 (A1 = label)
function saveLaborConfig(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Config');
    if (!sh) sh = ss.insertSheet('Config');
    var payload = {
      ratePerMin:  parseFloat(data.ratePerMin)  || 0,
      ratePerDay:  parseFloat(data.ratePerDay)  || 0,
      hoursPerDay: parseFloat(data.hoursPerDay) || 8,
      processes:   data.processes || []
    };
    sh.getRange('A1').setValue('LaborConfig');
    sh.getRange('B1').setValue(JSON.stringify(payload));
    return jsonOut({ status: 'ok', message: 'บันทึกค่าแรงแล้ว' });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── saveDrafts: บันทึกแบบร่างทั้งหมด (array) ลง sheet "Drafts" ──
// data: { drafts: [...] } — เก็บเป็น JSON string ไว้ที่เซลล์ B1 (A1 = label)
// เครื่องไหนบันทึก/ลบแบบร่าง จะ push array ทั้งก้อนทับของเดิม (sync ทุกเครื่อง)
function saveDrafts(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Drafts');
    if (!sh) sh = ss.insertSheet('Drafts');
    sh.getRange('A1').setValue('Drafts');
    sh.getRange('B1').setValue(JSON.stringify(data.drafts || []));
    return jsonOut({ status: 'ok', message: 'บันทึกแบบร่างแล้ว', count: (data.drafts||[]).length });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── getPlatingNoNeed: อ่านรายการ No.PO ที่มาร์ค "ไม่ต้องชุบ" จาก sheet "PlatingNoNeed" (B1 = JSON array) ──
function getPlatingNoNeed() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('PlatingNoNeed');
    var list = [];
    if (sh) {
      var raw = sh.getRange('B1').getValue();
      list = raw ? JSON.parse(raw) : [];
    }
    return jsonOut({ status: 'ok', list: list });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── savePlatingNoNeed: บันทึกรายการ No.PO ที่มาร์ค "ไม่ต้องชุบ" ทั้งก้อน (sync ทุกเครื่อง) ──
// data: { list: [...] } — เก็บเป็น JSON string ไว้ที่เซลล์ B1 (A1 = label)
function savePlatingNoNeed(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('PlatingNoNeed');
    if (!sh) sh = ss.insertSheet('PlatingNoNeed');
    sh.getRange('A1').setValue('PlatingNoNeed');
    sh.getRange('B1').setValue(JSON.stringify(data.list || []));
    return jsonOut({ status: 'ok', message: 'บันทึกรายการไม่ต้องชุบแล้ว', count: (data.list||[]).length });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── getPlatingPriceMemory: อ่านความจำราคาใบส่งชุบ จาก sheet "PlatingPriceMemory" (B1 = JSON array) ──
// แต่ละรายการ: { key, description, top, bot, meshOut, meshIn, price }
function getPlatingPriceMemory() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('PlatingPriceMemory');
    var list = [];
    if (sh) {
      var raw = sh.getRange('B1').getValue();
      list = raw ? JSON.parse(raw) : [];
    }
    return jsonOut({ status: 'ok', list: list });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── savePlatingPriceMemory: บันทึกความจำราคาใบส่งชุบทั้งก้อน (sync ทุกเครื่อง) ──
// data: { list: [...] } — เก็บเป็น JSON string ไว้ที่เซลล์ B1 (A1 = label)
function savePlatingPriceMemory(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('PlatingPriceMemory');
    if (!sh) sh = ss.insertSheet('PlatingPriceMemory');
    sh.getRange('A1').setValue('PlatingPriceMemory');
    sh.getRange('B1').setValue(JSON.stringify(data.list || []));
    return jsonOut({ status: 'ok', message: 'บันทึกความจำราคาแล้ว', count: (data.list||[]).length });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── updateMat: เพิ่ม/แก้ไขราคาวัตถุดิบ (MAT-ฝา / MAT-ตะแกรง) ──
// data: { type: 'flap'|'mesh', code, name, price, priceBuy, w, l }
// คอลัมน์: B=code, C=name, D=(ไม่ใช้), E=ราคา+30%, F=ราคาซื้อจริง, G=กว้าง, H=ยาว (เริ่มแถว 3)
function updateMat(data) {
  try {
    var code = String(data.code || '').trim();
    if (!code) throw new Error('code is required');
    var sheetName = (data.type === 'mesh') ? 'MAT-ตะแกรง' : 'MAT-ฝา';
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return jsonOut({ status: 'error', message: 'ไม่พบ sheet ' + sheetName });

    var name     = data.name || '';
    var price    = parseFloat(data.price) || 0;
    var priceBuy = parseFloat(data.priceBuy) || 0;
    var w        = parseFloat(data.w) || 0;
    var l        = parseFloat(data.l) || 0;

    var last = sh.getLastRow();
    var found = -1;
    if (last >= 3) {
      var codes = sh.getRange(3, 2, last - 2, 1).getValues();
      for (var i = 0; i < codes.length; i++) {
        if (String(codes[i][0]).trim() === code) { found = i + 3; break; }
      }
    }
    if (found !== -1) {
      sh.getRange(found, 3).setValue(name);   // C = ชื่อรายการ
      sh.getRange(found, 5).setValue(price);  // E = ราคา+30%
      sh.getRange(found, 6).setValue(priceBuy); // F = ราคาซื้อจริง
      if (w) sh.getRange(found, 7).setValue(w); // G = กว้าง
      if (l) sh.getRange(found, 8).setValue(l); // H = ยาว
      return jsonOut({ status: 'ok', updated: code });
    } else {
      var newRow = (last < 2) ? 3 : last + 1;
      sh.getRange(newRow, 2).setValue(code);
      sh.getRange(newRow, 3).setValue(name);
      sh.getRange(newRow, 5).setValue(price);
      sh.getRange(newRow, 6).setValue(priceBuy);
      if (w) sh.getRange(newRow, 7).setValue(w); // G = กว้าง
      if (l) sh.getRange(newRow, 8).setValue(l); // H = ยาว
      return jsonOut({ status: 'ok', added: code });
    }
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── deleteMat: ลบรายการวัตถุดิบ (MAT-ฝา / MAT-ตะแกรง) ──
// data: { type: 'flap'|'mesh', code }
function deleteMat(data) {
  try {
    var code = String(data.code || '').trim();
    if (!code) throw new Error('code is required');
    var sheetName = (data.type === 'mesh') ? 'MAT-ตะแกรง' : 'MAT-ฝา';
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return jsonOut({ status: 'error', message: 'ไม่พบ sheet ' + sheetName });

    var last = sh.getLastRow();
    if (last >= 3) {
      var codes = sh.getRange(3, 2, last - 2, 1).getValues();
      for (var i = 0; i < codes.length; i++) {
        if (String(codes[i][0]).trim() === code) {
          sh.deleteRow(i + 3);
          return jsonOut({ status: 'ok', deleted: code });
        }
      }
    }
    return jsonOut({ status: 'error', message: 'ไม่พบ ' + code });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── getOrders: ดึงข้อมูลจาก sheet "Order" (ใช้ร่วมกับ AppSheet) ──
// คอลัมน์ A-AB (28 คอลัมน์) — A=No.Quo, B=No.PO, ... AB=ลูกค้า
function getOrders(limit) {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var sh   = ss.getSheetByName('Order');
    if (!sh) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', rows: [], total: 0 }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var lastRaw = sh.getLastRow();
    var last = 1;
    if (lastRaw >= 2) {
      var colB = sh.getRange(2, 2, lastRaw - 1, 1).getValues();
      for (var i = colB.length - 1; i >= 0; i--) {
        if (String(colB[i][0]).trim() !== '') { last = i + 2; break; }
      }
    }
    var total = Math.max(0, last - 1);
    if (total === 0) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', rows: [], total: 0 }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var lim      = (limit && limit > 0 && limit < total) ? limit : total;
    var startRow = last - lim + 1;
    var numCols  = Math.max(28, sh.getLastColumn());
    var values   = sh.getRange(startRow, 1, lim, numCols).getValues();
    var rows = values.map(function(r) {
      return r.map(function(cell) {
        if (cell instanceof Date) {
          return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        }
        return cell;
      });
    });
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', rows: rows, total: total, limited: lim < total }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: e.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── addOrder: เพิ่มแถวใหม่ใน sheet "Order" ───────────────────
// data.row = array ค่าตามลำดับคอลัมน์ A-AB (28 ค่า)
function addOrder(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!Array.isArray(data.row) || data.row.length === 0) throw new Error('row data is required');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Order');
    if (!sh) sh = ss.insertSheet('Order');
    // หาแถวสุดท้ายที่มีข้อมูลจริง (เช็คคอลัมน์ A No.Quo และ B No.PO)
    // เพราะ getLastRow()/appendRow() อาจนับรวมแถวที่มีแค่ format ว่างๆ ทำให้เพิ่มแถวห่างจากข้อมูลจริงมาก
    var maxRows = sh.getMaxRows();
    var lastDataRow = 1;
    if (maxRows > 1) {
      var ab = sh.getRange(2, 1, maxRows - 1, 2).getValues();
      for (var i = ab.length - 1; i >= 0; i--) {
        if (String(ab[i][0]).trim() !== '' || String(ab[i][1]).trim() !== '') { lastDataRow = i + 2; break; }
      }
    }
    var targetRow = lastDataRow > 1 ? lastDataRow + 1 : 2;
    sh.getRange(targetRow, 1, 1, data.row.length).setValues([data.row]);
    return jsonOut({ status: 'ok', message: 'สร้าง Order แล้ว' });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── updateOrder: อัปเดตแถวที่ตรงกับ No.PO (คอลัมน์ B) ─────────
// ถ้าไม่พบ No.PO เดิม จะ append แถวใหม่แทน
function updateOrder(noPO, row) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!noPO) throw new Error('noPO is required');
    if (!Array.isArray(row) || row.length === 0) throw new Error('row data is required');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Order');
    if (!sh) sh = ss.insertSheet('Order');
    var last = sh.getLastRow();
    if (last < 2) {
      sh.appendRow(row);
      return jsonOut({ status: 'ok', updated: noPO, note: 'appended' });
    }
    var colB  = sh.getRange(2, 2, last - 1, 1).getValues();
    var found = -1;
    for (var i = 0; i < colB.length; i++) {
      if (String(colB[i][0]) == String(noPO)) { found = i + 2; break; }
    }
    // ── sync column P (statusDeliver) ตาม Process (column Q) ──
    // P index 15 (0-based), Q index 16 (0-based)
    if (row.length > 16) {
      var proc = String(row[16] || '').trim();
      row[15] = (proc === 'เรียบร้อย') ? true : false;
    }
    if (found !== -1) {
      sh.getRange(found, 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    return jsonOut({ status: 'ok', updated: noPO, note: found !== -1 ? 'updated' : 'appended' });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── หาหรือสร้างโฟลเดอร์ตาม path ซ้อนกัน เช่น ['appsheet','data','PTSytem-1001399782','ORDER_Images'] ──
function _getOrCreateFolderPath(pathParts) {
  var folder = DriveApp.getRootFolder();
  for (var i = 0; i < pathParts.length; i++) {
    var name = pathParts[i];
    var it = folder.getFoldersByName(name);
    folder = it.hasNext() ? it.next() : folder.createFolder(name);
  }
  return folder;
}

// ── uploadOrderFile: รับไฟล์ base64 → บันทึกลง Drive โฟลเดอร์ ORDER_Images ของ AppSheet ──
// data: { fileName, mimeType, base64, noPO }
function uploadOrderFile(data) {
  try {
    if (!data.base64) throw new Error('base64 is required');
    var folder = _getOrCreateFolderPath(['appsheet', 'data', 'PTSytem-1001399782', 'ORDER_Images']);

    // ตั้งชื่อไฟล์ตามรูปแบบของ AppSheet: {เลขที่PO}.รูปภาพPO..{HHmmss}.{ext}
    var origName = data.fileName || 'PO_file.jpg';
    var extMatch = origName.match(/\.([a-zA-Z0-9]+)$/);
    var ext = extMatch ? extMatch[1] : 'jpg';
    var noPO = String(data.noPO || '').trim() || 'PO';
    var hhmmss = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'HHmmss');
    var fileName = noPO + '.รูปภาพPO..' + hhmmss + '.' + ext;

    var bytes = Utilities.base64Decode(data.base64);
    var blob  = Utilities.newBlob(bytes, data.mimeType || 'application/octet-stream', fileName);
    var file  = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var relPath = 'ORDER_Images/' + fileName;
    var appUrl = 'https://www.appsheet.com/template/gettablefileurl?appName=PTSytem-1001399782&tablename=ORDER&filename='
      + encodeURIComponent(relPath);

    // url      → บันทึกลงคอลัมน์ V (ลิงก์เปิดดูรูปแบบ AppSheet)
    // path     → บันทึกลงคอลัมน์ M (path สัมพัทธ์ เช่น ORDER_Images/{เลขที่PO}.รูปภาพPO..{ชื่อไฟล์}.jpg)
    return jsonOut({ status: 'ok', url: appUrl, path: relPath, fileId: file.getId() });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── getOrderPoImage: อ่านรูป PO จาก Drive แล้ว return base64 ─────────────
// GET ?action=getOrderPoImage&path=ORDER_Images/...  หรือ &noPO=2207050
function getOrderPoImage(path, noPO) {
  try {
    var folder = _getOrCreateFolderPath(['appsheet', 'data', 'PTSytem-1001399782', 'ORDER_Images']);
    var file = null;

    // วิธี 1: ค้นหาจาก path ตรงๆ
    if (path) {
      var fileName = path.replace(/^.*ORDER_Images\//, '');
      var byName = folder.getFilesByName(fileName);
      if (byName.hasNext()) file = byName.next();
    }

    // วิธี 2: ถ้าไม่มี path หรือหาไม่เจอ → ค้นหาจาก noPO (prefix match)
    if (!file && noPO) {
      var allFiles = folder.getFiles();
      while (allFiles.hasNext()) {
        var f = allFiles.next();
        var n = f.getName();
        // ไฟล์ PO: {noPO}.รูปภาพPO..{time}.{ext}
        // ไฟล์ Drawing: {noPO}.Drawing..{time}.{ext} หรือ {noPO}.รูปงาน..{time}.{ext}
        if (n.indexOf(noPO + '.') === 0) { file = f; break; }
      }
    }

    if (!file) return jsonOut({ status:'error', message:'ไม่พบรูป PO' });
    var blob = file.getBlob();
    var mime = blob.getContentType() || 'image/jpeg';
    var b64  = Utilities.base64Encode(blob.getBytes());
    return jsonOut({ status:'ok', src: 'data:' + mime + ';base64,' + b64 });
  } catch(e) {
    return jsonOut({ status:'error', message: e.toString() });
  }
}

// ── getOrderQty: ดึงจำนวนชิ้นจาก Order sheet ตาม noPO ───────────────────
function getOrderQty(noPO) {
  try {
    if (!noPO) return jsonOut({ status:'error', message:'no noPO' });
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Order');
    if (!sh || sh.getLastRow() < 2) return jsonOut({ status:'ok', qty:'' });
    var data = sh.getRange(2, 1, sh.getLastRow()-1, 8).getValues(); // A-H (cols 1-8)
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(noPO).trim()) { // col B = noPO (index 1)
        return jsonOut({ status:'ok', qty: String(data[i][6] || '') }); // col G = qty (index 6)
      }
    }
    return jsonOut({ status:'ok', qty:'' });
  } catch(e) {
    return jsonOut({ status:'error', message: e.toString() });
  }
}

// ── deleteOrder: ลบแถวที่ตรงกับ No.PO (คอลัมน์ B) ใน sheet "Order" ──
function deleteOrder(noPO) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!noPO) throw new Error('noPO is required');
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var sh   = ss.getSheetByName('Order');
    if (!sh) throw new Error('ไม่พบ sheet Order');
    var last = sh.getLastRow();
    if (last >= 2) {
      // อ่านคอลัมน์ B (No.PO) ทั้งหมดในครั้งเดียว แทนการเรียก getRange ทีละแถว (ช้ามากถ้าชีตมีหลายพันแถว)
      var colB = sh.getRange(2, 2, last - 1, 1).getValues();
      for (var i = colB.length - 1; i >= 0; i--) {
        if (String(colB[i][0]) == String(noPO)) {
          sh.deleteRow(i + 2);
          return jsonOut({ status: 'ok', deleted: noPO });
        }
      }
    }
    throw new Error('ไม่พบ ' + noPO);
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── getSuppliers: ดึงรายชื่อ supplier จาก sheet "Suppliers" (A-F) ──
// A=รหัส, B=ชื่อบริษัท, C=ที่อยู่, D=เลขผู้เสียภาษี, E=ผู้ติดต่อ/เบอร์, F=หมายเหตุ
function getSuppliers() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Suppliers');
    if (!sh) return jsonOut({ status: 'ok', suppliers: [] });
    var last = sh.getLastRow();
    if (last < 2) return jsonOut({ status: 'ok', suppliers: [] });
    var vals = sh.getRange(2, 1, last - 1, 8).getValues();
    var suppliers = vals.map(function(r) {
      return {
        code:         String(r[0]||'').trim(),
        name:         String(r[1]||'').trim(),
        address:      String(r[2]||'').trim(),
        taxId:        String(r[3]||'').trim(),
        contact:      String(r[4]||'').trim(),
        note:         String(r[5]||'').trim(),
        leadTimeDays: Number(r[6]||0),
        matCodes:     String(r[7]||'').trim()
      };
    }).filter(function(s){ return s.code; });
    return jsonOut({ status: 'ok', suppliers: suppliers });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── saveSupplier: เพิ่ม/แก้ไข Supplier (ค้นหาด้วย code; ถ้าไม่มี code = เพิ่มใหม่) ──
// data: { code, name, address, taxId, contact, note }
function saveSupplier(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Suppliers');
    if (!sh) sh = ss.insertSheet('Suppliers');

    var name = String(data.name || '').trim();
    if (!name) throw new Error('name is required');
    var code = String(data.code || '').trim();
    var last = sh.getLastRow();

    var found = -1;
    if (code && last >= 1) {
      var codes = sh.getRange(1, 1, last, 1).getValues();
      for (var i = 0; i < codes.length; i++) {
        if (String(codes[i][0]).trim() === code) { found = i + 1; break; }
      }
    }
    if (!code) code = 'S' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'yyMMddHHmmss');

    var row = [
      code, name,
      String(data.address      || '').trim(),
      String(data.taxId        || '').trim(),
      String(data.contact      || '').trim(),
      String(data.note         || '').trim(),
      Number(data.leadTimeDays || 0),
      String(data.matCodes     || '').trim()
    ];
    if (found !== -1) {
      sh.getRange(found, 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    return jsonOut({ status: 'ok', code: code, note: found !== -1 ? 'updated' : 'added' });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── deleteSupplier: ลบ Supplier ตามรหัส ───────────────────────
function deleteSupplier(code) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    code = String(code || '').trim();
    if (!code) throw new Error('code is required');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Suppliers');
    if (!sh) throw new Error('ไม่พบ sheet Suppliers');
    var last = sh.getLastRow();
    for (var i = last; i >= 1; i--) {
      if (String(sh.getRange(i, 1).getValue()).trim() === code) {
        sh.deleteRow(i);
        return jsonOut({ status: 'ok', deleted: code });
      }
    }
    throw new Error('ไม่พบ ' + code);
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── getPurchaseOrders: ดึงข้อมูลใบสั่งซื้อ ──
// header จาก sheet "PurchaseOrders" (A-N, 14 คอลัมน์)
// items จาก sheet "PO_Items" (A-I, 9 คอลัมน์ — I = imageUrl) จัดกลุ่มตาม poNo (คอลัมน์ A)
function getPurchaseOrders() {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var shH = ss.getSheetByName('PurchaseOrders');
    var shI = ss.getSheetByName('PO_Items');
    var fmt = function(r) {
      return r.map(function(cell) {
        if (cell instanceof Date) return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        return cell;
      });
    };
    var headers = [];
    if (shH) {
      var lastH = shH.getLastRow();
      if (lastH >= 2) {
        headers = shH.getRange(2, 1, lastH - 1, 14).getValues().map(fmt);
      }
    }
    var itemsByPO = {};
    if (shI) {
      var lastI = shI.getLastRow();
      if (lastI >= 2) {
        shI.getRange(2, 1, lastI - 1, 9).getValues().forEach(function(r) {
          var poNo = String(r[0]||'').trim();
          if (!poNo) return;
          if (!itemsByPO[poNo]) itemsByPO[poNo] = [];
          itemsByPO[poNo].push(fmt(r));
        });
      }
    }
    return jsonOut({ status: 'ok', headers: headers, items: itemsByPO });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── getPOSupplierItems: ดึงรายการสินค้าที่เคยสั่งซื้อ แยกตาม supplier จาก sheet "PO_SupplierItems" (A-G) ──
// A=supplierCode, B=name, C=spec, D=unit, E=unitPrice(ล่าสุด), F=imageUrl, G=updatedAt
function getPOSupplierItems() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('PO_SupplierItems');
    if (!sh) return jsonOut({ status: 'ok', items: [] });
    var last = sh.getLastRow();
    if (last < 2) return jsonOut({ status: 'ok', items: [] });
    var vals = sh.getRange(2, 1, last - 1, 7).getValues();
    var items = vals.map(function(r) {
      return {
        supplierCode: String(r[0]||'').trim(),
        name:      String(r[1]||'').trim(),
        spec:      String(r[2]||'').trim(),
        unit:      String(r[3]||'').trim(),
        unitPrice: r[4],
        imageUrl:  String(r[5]||'').trim(),
        updatedAt: r[6]
      };
    }).filter(function(it){ return it.supplierCode && it.name; });
    return jsonOut({ status: 'ok', items: items });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── deletePOSupplierItem: ลบรายการออกจากคลังสินค้าของ supplier (sheet "PO_SupplierItems") ──
// data: { supplierCode, name }
function deletePOSupplierItem(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var supplierCode = String(data.supplierCode || '').trim();
    var name = String(data.name || '').trim();
    if (!supplierCode || !name) throw new Error('supplierCode and name are required');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('PO_SupplierItems');
    if (!sh) return jsonOut({ status: 'ok', deleted: false });
    var last = sh.getLastRow();
    if (last >= 2) {
      var vals = sh.getRange(2, 1, last - 1, 2).getValues();
      for (var i = vals.length; i >= 1; i--) {
        if (String(vals[i-1][0]).trim().toLowerCase() === supplierCode.toLowerCase() &&
            String(vals[i-1][1]).trim().toLowerCase() === name.toLowerCase()) {
          sh.deleteRow(i + 1);
        }
      }
    }
    return jsonOut({ status: 'ok', deleted: true });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── getNextPONo: สร้างเลขที่ใบสั่งซื้อใหม่ รูปแบบ PUR-YYMM-XX ──
function getNextPONo() {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var sh  = ss.getSheetByName('PurchaseOrders');
    var now = new Date();
    var prefix = 'PUR-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyMM');
    var maxSeq = 0;
    if (sh) {
      var last = sh.getLastRow();
      if (last >= 2) {
        sh.getRange(2, 1, last - 1, 1).getValues().forEach(function(r) {
          var v = String(r[0]||'').trim();
          if (v.indexOf(prefix + '-') === 0) {
            var seq = parseInt(v.substring(prefix.length + 1), 10);
            if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
          }
        });
      }
    }
    var nextPO = prefix + '-' + ('0' + (maxSeq + 1)).slice(-2);
    return jsonOut({ status: 'ok', nextPO: nextPO });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── savePurchaseOrder: สร้าง/แก้ไขใบสั่งซื้อ (header + items) ──
// data: { poNo, header: [A..N ตามลำดับคอลัมน์ PurchaseOrders],
//         items: [[A..H ตามลำดับคอลัมน์ PO_Items], ...] }
// ถ้าพบ poNo เดิมในคอลัมน์ A ของ PurchaseOrders → อัปเดตแถวนั้น (และลบ/เขียน items ใหม่ทั้งหมด)
// ถ้าไม่พบ → สร้างแถวใหม่
// ── _upsertPOSupplierItems: อัปเดต/เพิ่มรายการสินค้าในคลังของ supplier (sheet "PO_SupplierItems") ──
// item array: [poNo, seq, name, spec, qty, unit, unitPrice, lineTotal, imageUrl]
// match ด้วย supplierCode + name (case-insensitive, trim) — ถ้าพบ อัปเดต spec/unit/unitPrice(ล่าสุด)/imageUrl/updatedAt, ถ้าไม่พบ เพิ่มแถวใหม่
function _upsertPOSupplierItems(ss, supplierCode, items) {
  var sh = ss.getSheetByName('PO_SupplierItems');
  if (!sh) {
    sh = ss.insertSheet('PO_SupplierItems');
    sh.appendRow(['supplierCode', 'name', 'spec', 'unit', 'unitPrice', 'imageUrl', 'updatedAt']);
  }
  var last = sh.getLastRow();
  var existing = last >= 2 ? sh.getRange(2, 1, last - 1, 2).getValues() : [];
  var now = new Date();
  items.forEach(function(item) {
    var name = String(item[2] || '').trim();
    if (!name) return;
    var spec = item[3] || '';
    var unit = item[5] || '';
    var unitPrice = item[6] || 0;
    var imageUrl = item[8] || '';
    var foundRow = -1;
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i][0]).trim().toLowerCase() === supplierCode.toLowerCase() &&
          String(existing[i][1]).trim().toLowerCase() === name.toLowerCase()) {
        foundRow = i + 2;
        break;
      }
    }
    if (foundRow !== -1) {
      var rowVals = sh.getRange(foundRow, 1, 1, 7).getValues()[0];
      sh.getRange(foundRow, 1, 1, 7).setValues([[
        supplierCode, name, spec || rowVals[2], unit || rowVals[3], unitPrice, imageUrl || rowVals[5], now
      ]]);
    } else {
      sh.appendRow([supplierCode, name, spec, unit, unitPrice, imageUrl, now]);
      existing.push([supplierCode, name]);
    }
  });
}

function savePurchaseOrder(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var poNo = String(data.poNo || '').trim();
    if (!poNo) throw new Error('poNo is required');
    if (!Array.isArray(data.header) || data.header.length === 0) throw new Error('header is required');
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── header: PurchaseOrders ──
    var shH = ss.getSheetByName('PurchaseOrders');
    if (!shH) shH = ss.insertSheet('PurchaseOrders');
    var lastH  = shH.getLastRow();
    var foundH = -1;
    if (lastH >= 2) {
      var poNos = shH.getRange(2, 1, lastH - 1, 1).getValues();
      for (var i = 0; i < poNos.length; i++) {
        if (String(poNos[i][0]).trim() === poNo) { foundH = i + 2; break; }
      }
    }
    if (foundH !== -1) {
      shH.getRange(foundH, 1, 1, data.header.length).setValues([data.header]);
    } else {
      shH.appendRow(data.header);
    }

    // ── items: PO_Items (ลบของเดิมทั้งหมดของ poNo นี้ แล้วเขียนใหม่) ──
    var shI = ss.getSheetByName('PO_Items');
    if (!shI) shI = ss.insertSheet('PO_Items');
    var lastI = shI.getLastRow();
    if (lastI >= 2) {
      var colA = shI.getRange(2, 1, lastI - 1, 1).getValues();
      for (var j = colA.length; j >= 1; j--) {
        if (String(colA[j-1][0]).trim() === poNo) shI.deleteRow(j + 1);
      }
    }
    if (Array.isArray(data.items)) {
      data.items.forEach(function(item) { shI.appendRow(item); });
    }

    // ── อัปเดตคลังรายการสินค้าต่อ supplier: PO_SupplierItems (upsert ตาม supplierCode+name) ──
    var supplierCode = String(data.header[3] || '').trim();
    if (supplierCode && Array.isArray(data.items) && data.items.length) {
      _upsertPOSupplierItems(ss, supplierCode, data.items);
    }

    return jsonOut({ status: 'ok', saved: poNo, note: foundH !== -1 ? 'updated' : 'created' });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── deletePurchaseOrder: ลบใบสั่งซื้อ (header ใน PurchaseOrders + รายการใน PO_Items) ──
function deletePurchaseOrder(poNo) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    poNo = String(poNo || '').trim();
    if (!poNo) throw new Error('poNo is required');
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var shH = ss.getSheetByName('PurchaseOrders');
    var deletedHeader = false;
    if (shH) {
      var lastH = shH.getLastRow();
      for (var i = lastH; i >= 2; i--) {
        if (String(shH.getRange(i, 1).getValue()).trim() === poNo) {
          shH.deleteRow(i);
          deletedHeader = true;
          break;
        }
      }
    }

    var shI = ss.getSheetByName('PO_Items');
    if (shI) {
      var lastI = shI.getLastRow();
      for (var j = lastI; j >= 2; j--) {
        if (String(shI.getRange(j, 1).getValue()).trim() === poNo) shI.deleteRow(j);
      }
    }

    if (!deletedHeader) throw new Error('ไม่พบ ' + poNo);
    return jsonOut({ status: 'ok', deleted: poNo });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── deleteRow: ลบแถวที่ตรงกับ No.Quo ─────────────────────────
function deleteRow(noQuo) {
  try {
    if (!noQuo) throw new Error('noQuo is required');
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var sh   = ss.getSheetByName('DATA') || ss.getSheets()[0];
    var last = sh.getLastRow();
    for (var i = last; i >= 2; i--) {
      if (sh.getRange(i, 2).getValue() == noQuo) {
        sh.deleteRow(i);
        return jsonOut({ status: 'ok', deleted: noQuo });
      }
    }
    throw new Error('ไม่พบ ' + noQuo);
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ══════════════════════════════════════════════════════════════
//  ใบกำกับภาษี: Company / Customers / Invoices
// ══════════════════════════════════════════════════════════════

// ── getCompanyInfo: อ่านชีต "Company" (A=key, B=value) ───────
function getCompanyInfo() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Company');
    var info = {
      name: '', nameEn: '', address: '', addressEn: '', phone: '', email: '', taxId: '', logoUrl: ''
    };
    if (!sh) return jsonOut({ status: 'ok', info: info });
    var last = sh.getLastRow();
    if (last < 1) return jsonOut({ status: 'ok', info: info });
    var vals = sh.getRange(1, 1, last, 2).getValues();
    var map = {
      company_name: 'name', company_name_en: 'nameEn', company_address: 'address',
      company_address_en: 'addressEn',
      company_phone: 'phone', company_email: 'email', company_taxid: 'taxId',
      company_logo_url: 'logoUrl'
    };
    vals.forEach(function(r) {
      var key = String(r[0] || '').trim();
      var val = String(r[1] || '').trim();
      if (map[key]) info[map[key]] = val;
    });
    return jsonOut({ status: 'ok', info: info });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── saveCompanyInfo: เขียนค่าลงชีต "Company" (สร้างแถวถ้ายังไม่มี) ──
// data: { name, nameEn, address, addressEn, phone, email, taxId, logoUrl }
function saveCompanyInfo(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Company');
    if (!sh) sh = ss.insertSheet('Company');

    var fields = {
      company_name:     data.name    || '',
      company_name_en:  data.nameEn  || '',
      company_address:  data.address || '',
      company_address_en: data.addressEn || '',
      company_phone:    data.phone   || '',
      company_email:    data.email   || '',
      company_taxid:    data.taxId   || '',
      company_logo_url: data.logoUrl || ''
    };

    var last = sh.getLastRow();
    var keyRows = {};
    if (last >= 1) {
      var keys = sh.getRange(1, 1, last, 1).getValues();
      for (var i = 0; i < keys.length; i++) {
        var k = String(keys[i][0] || '').trim();
        if (k) keyRows[k] = i + 1;
      }
    }
    Object.keys(fields).forEach(function(key) {
      // logoUrl: อัปเดตเฉพาะเมื่อส่งค่ามา หรือสั่งลบ (clearLogo) — ไม่ลบของเดิมถ้าไม่ได้เปลี่ยนโลโก้
      if (key === 'company_logo_url' && !data.logoUrl && !data.clearLogo && keyRows[key]) return;
      if (keyRows[key]) {
        sh.getRange(keyRows[key], 2).setValue(fields[key]);
      } else {
        sh.appendRow([key, fields[key]]);
      }
    });
    return jsonOut({ status: 'ok', message: 'บันทึกข้อมูลบริษัทแล้ว' });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── uploadLogo: รับไฟล์ base64 → บันทึกลง Drive คืน URL ───────
// data: { fileName, mimeType, base64 }
function uploadLogo(data) {
  try {
    if (!data.base64) throw new Error('base64 is required');
    var folder = _getOrCreateFolderPath(['appsheet', 'data', 'PTSytem-1001399782', 'Company']);

    var origName = data.fileName || 'logo.png';
    var extMatch = origName.match(/\.([a-zA-Z0-9]+)$/);
    var ext = extMatch ? extMatch[1] : 'png';
    var fileName = 'logo.' + ext;

    var bytes = Utilities.base64Decode(data.base64);
    var blob  = Utilities.newBlob(bytes, data.mimeType || 'image/png', fileName);
    var file  = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // ใช้ URL รูปแบบ lh3.googleusercontent.com — แสดงผลใน <img> ได้แน่นอนกว่า drive.google.com/uc
    // (ตัวหลังมักเจอหน้า "ยืนยันการสแกนไวรัส" ทำให้รูปไม่ขึ้น)
    var url = 'https://lh3.googleusercontent.com/d/' + file.getId() + '=w400';
    return jsonOut({ status: 'ok', url: url, fileId: file.getId() });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── uploadPOItemImage: รับไฟล์ base64 → บันทึกลง Drive คืน URL (รูปรายการในใบสั่งซื้อ) ──
// data: { fileName, mimeType, base64 }
function uploadPOItemImage(data) {
  try {
    if (!data.base64) throw new Error('base64 is required');
    var folder = _getOrCreateFolderPath(['appsheet', 'data', 'PTSytem-1001399782', 'PO_Items']);

    var origName = data.fileName || 'item.png';
    var extMatch = origName.match(/\.([a-zA-Z0-9]+)$/);
    var ext = extMatch ? extMatch[1] : 'png';
    var fileName = 'item_' + new Date().getTime() + '.' + ext;

    var bytes = Utilities.base64Decode(data.base64);
    var blob  = Utilities.newBlob(bytes, data.mimeType || 'image/png', fileName);
    var file  = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var url = 'https://lh3.googleusercontent.com/d/' + file.getId() + '=w400';
    return jsonOut({ status: 'ok', url: url, fileId: file.getId() });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── getCustomers: ดึงรายชื่อลูกค้าจากชีต "Customers" (A-I) ────
// A=รหัสลูกค้า, B=ชื่อ/บริษัท, C=สาขา, D=เลขผู้เสียภาษี, E=ที่อยู่, F=เบอร์โทร, G=ผู้ติดต่อ, H=หัก ณ ที่จ่าย, I=เครดิต
function getCustomers() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Customers');
    if (!sh) return jsonOut({ status: 'ok', customers: [] });
    var last = sh.getLastRow();
    if (last < 2) return jsonOut({ status: 'ok', customers: [] }); // แถว 1 = หัวตาราง
    var vals = sh.getRange(2, 1, last - 1, 9).getValues();
    var customers = vals.map(function(r) {
      return {
        code:    String(r[0] || '').trim(),
        name:    String(r[1] || '').trim(),
        branch:  String(r[2] || '').trim(),
        taxId:   String(r[3] || '').trim(),
        address: String(r[4] || '').trim(),
        phone:   String(r[5] || '').trim(),
        contact: String(r[6] || '').trim(),
        wht:     r[7] === true || String(r[7]).toLowerCase() === 'true',
        credit:  String(r[8] || '').trim()
      };
    }).filter(function(c) { return c.name; });
    return jsonOut({ status: 'ok', customers: customers });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── saveCustomer: เพิ่ม/แก้ไขลูกค้า (ค้นหาด้วย code; ถ้าไม่มี code = เพิ่มใหม่) ──
// data: { code, name, branch, taxId, address, phone, contact }
function saveCustomer(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Customers');
    if (!sh) sh = ss.insertSheet('Customers');

    var name = String(data.name || '').trim();
    if (!name) throw new Error('name is required');
    var code = String(data.code || '').trim();
    var last = sh.getLastRow();

    var found = -1;
    if (code && last >= 1) {
      var codes = sh.getRange(1, 1, last, 1).getValues();
      for (var i = 0; i < codes.length; i++) {
        if (String(codes[i][0]).trim() === code) { found = i + 1; break; }
      }
    }
    if (!code) code = 'C' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'yyMMddHHmmss');

    var row = [
      code, name,
      String(data.branch  || '').trim(),
      String(data.taxId   || '').trim(),
      String(data.address || '').trim(),
      String(data.phone   || '').trim(),
      String(data.contact || '').trim(),
      data.wht ? true : false,
      String(data.credit  || '').trim()
    ];
    if (found !== -1) {
      sh.getRange(found, 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    return jsonOut({ status: 'ok', code: code, note: found !== -1 ? 'updated' : 'added' });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── deleteCustomer: ลบลูกค้าตามรหัส ───────────────────────────
function deleteCustomer(code) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    code = String(code || '').trim();
    if (!code) throw new Error('code is required');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Customers');
    if (!sh) throw new Error('ไม่พบ sheet Customers');
    var last = sh.getLastRow();
    for (var i = last; i >= 1; i--) {
      if (String(sh.getRange(i, 1).getValue()).trim() === code) {
        sh.deleteRow(i);
        return jsonOut({ status: 'ok', deleted: code });
      }
    }
    throw new Error('ไม่พบ ' + code);
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── getItemMaster: ดึงรายการสินค้า/บริการที่ใช้บ่อยจากชีต "ItemMaster" (A-D) ──
// A=รหัส, B=ชื่อรายการ, C=หน่วย, D=ราคา/หน่วย
function getItemMaster() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('ItemMaster');
    if (!sh) return jsonOut({ status: 'ok', items: [] });
    var last = sh.getLastRow();
    if (last < 2) return jsonOut({ status: 'ok', items: [] }); // แถว 1 = หัวตาราง
    var vals = sh.getRange(2, 1, last - 1, 4).getValues();
    var items = vals.map(function(r) {
      return {
        code:  String(r[0] || '').trim(),
        name:  String(r[1] || '').trim(),
        unit:  String(r[2] || '').trim(),
        price: Number(r[3] || 0)
      };
    }).filter(function(it) { return it.name; });
    return jsonOut({ status: 'ok', items: items });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── saveItem: เพิ่ม/แก้ไขรายการสินค้า/บริการ (ค้นหาด้วย code; ถ้าไม่มี code = เพิ่มใหม่) ──
// data: { code, name, unit, price }
function saveItem(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('ItemMaster');
    if (!sh) {
      sh = ss.insertSheet('ItemMaster');
      sh.appendRow(['code','name','unit','price']);
    }

    var name = String(data.name || '').trim();
    if (!name) throw new Error('name is required');
    var code = String(data.code || '').trim();
    var last = sh.getLastRow();

    var found = -1;
    if (code && last >= 1) {
      var codes = sh.getRange(1, 1, last, 1).getValues();
      for (var i = 0; i < codes.length; i++) {
        if (String(codes[i][0]).trim() === code) { found = i + 1; break; }
      }
    }
    if (!code) code = 'I' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'yyMMddHHmmss');

    var row = [
      code, name,
      String(data.unit  || 'ชิ้น').trim(),
      Number(data.price || 0)
    ];
    if (found !== -1) {
      sh.getRange(found, 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    return jsonOut({ status: 'ok', code: code, note: found !== -1 ? 'updated' : 'added' });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── deleteItem: ลบรายการสินค้า/บริการตามรหัส ─────────────────
function deleteItem(code) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    code = String(code || '').trim();
    if (!code) throw new Error('code is required');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('ItemMaster');
    if (!sh) throw new Error('ไม่พบ sheet ItemMaster');
    var last = sh.getLastRow();
    for (var i = last; i >= 1; i--) {
      if (String(sh.getRange(i, 1).getValue()).trim() === code) {
        sh.deleteRow(i);
        return jsonOut({ status: 'ok', deleted: code });
      }
    }
    throw new Error('ไม่พบ ' + code);
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── getInvoices: ดึงรายการใบกำกับภาษีทั้งหมดจากชีต "Invoices" ──
// คอลัมน์: A=เลขที่ใบกำกับ B=วันที่ออก C=ประเภท D=รหัสลูกค้า
//          E=รายการPO(comma) F=ยอดก่อนVAT G=VAT H=ยอดรวม I=สถานะ J=ผู้ออก K=รายการสินค้า(JSON)
function getInvoices() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Invoices');
    if (!sh) return jsonOut({ status: 'ok', invoices: [] });
    var last = sh.getLastRow();
    if (last < 1) return jsonOut({ status: 'ok', invoices: [] });
    var vals = sh.getRange(1, 1, last, 12).getValues();
    var invoices = vals.map(function(r) {
      var dateVal = r[1];
      var dateStr = (dateVal instanceof Date)
        ? Utilities.formatDate(dateVal, Session.getScriptTimeZone() || 'GMT+7', 'dd/MM/yyyy')
        : String(dateVal || '').trim();
      var items = [];
      try { items = r[10] ? JSON.parse(r[10]) : []; } catch (e2) { items = []; }
      return {
        invoiceNo:   String(r[0] || '').trim(),
        date:        dateStr,
        type:        String(r[2] || '').trim(),
        customerCode:String(r[3] || '').trim(),
        poList:      String(r[4] || '').trim(),
        subtotal:    parseFloat(r[5]) || 0,
        vat:         parseFloat(r[6]) || 0,
        total:       parseFloat(r[7]) || 0,
        status:      String(r[8] || '').trim(),
        issuedBy:    String(r[9] || '').trim(),
        items:       items,
        billNo:      String(r[11] || '').trim()
      };
    }).filter(function(x) { return x.invoiceNo; });
    return jsonOut({ status: 'ok', invoices: invoices });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── ใบกำกับภาษี: เลขที่เอกสาร รูปแบบ {YY}{MM}{SSS} ────────────
// YY = ปี ค.ศ. 2 หลักท้าย, MM = เดือน 2 หลัก, SSS = เลขลำดับ 3 หลัก
// ใช้เลขลำดับร่วมกันทั้งแบบเต็มรูปและอย่างย่อในเดือนเดียวกัน
function _invoicePrefix(type) {
  return (type === 'short') ? 'ABB' : 'INV';
}
function _peekNextInvoiceNo(type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Invoices');
  var prefix = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'yyMM');
  var maxSeq = 0;
  if (sh) {
    var last = sh.getLastRow();
    if (last >= 1) {
      sh.getRange(1, 1, last, 1).getValues().forEach(function(r) {
        var v = String(r[0] || '').trim();
        if (v.indexOf(prefix) === 0 && v.length === prefix.length + 3) {
          var seq = parseInt(v.substring(prefix.length), 10);
          if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
        }
      });
    }
  }
  return prefix + ('00' + (maxSeq + 1)).slice(-3);
}

// ── saveInvoiceRecord: ออกเลขที่ใบกำกับใหม่ + log + mark PO ที่เกี่ยวข้อง ──
// data: { type:'full'|'short', customerCode, poList:[noPO,...],
//         subtotal, vat, total, issuedBy }
// คอลัมน์ Invoices: A=เลขที่ใบกำกับ B=วันที่ออก C=ประเภท D=รหัสลูกค้า
//                   E=รายการPO(comma) F=ยอดก่อนVAT G=VAT H=ยอดรวม I=สถานะ J=ผู้ออก
// คอลัมน์ Order: AC(29)=เลขที่ใบกำกับ AD(30)=วันที่ออกใบกำกับ
function saveInvoiceRecord(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Invoices');
    if (!sh) sh = ss.insertSheet('Invoices');

    var type = (data.type === 'short') ? 'short' : 'full';
    var invoiceNo = _peekNextInvoiceNo(type);
    // ใช้วันที่จาก payload (พ.ศ.) ถ้ามี, ไม่งั้น generate วันนี้เป็น พ.ศ.
    var dateStr = (function() {
      if (data.date && String(data.date).match(/^\d{2}\/\d{2}\/\d{4}$/)) return String(data.date);
      var now = new Date();
      var dd = Utilities.formatDate(now, Session.getScriptTimeZone() || 'GMT+7', 'dd');
      var mm = Utilities.formatDate(now, Session.getScriptTimeZone() || 'GMT+7', 'MM');
      var yyyy = parseInt(Utilities.formatDate(now, Session.getScriptTimeZone() || 'GMT+7', 'yyyy'), 10) + 543;
      return dd + '/' + mm + '/' + yyyy;
    })();
    var poList = Array.isArray(data.poList) ? data.poList : [];

    var items = Array.isArray(data.items) ? data.items : [];

    sh.appendRow([
      invoiceNo, dateStr, (type === 'short' ? 'อย่างย่อ' : 'เต็มรูป'),
      String(data.customerCode || ''), poList.join(', '),
      parseFloat(data.subtotal) || 0, parseFloat(data.vat) || 0, parseFloat(data.total) || 0,
      'ออกแล้ว', String(data.issuedBy || ''), JSON.stringify(items)
    ]);

    // ── mark Order rows ที่เกี่ยวข้อง + เปลี่ยน Process เป็น "เตรียมส่ง" ──
    var shO = ss.getSheetByName('Order');
    if (shO && poList.length) {
      var lastO = shO.getLastRow();
      if (lastO >= 2) {
        var colB = shO.getRange(2, 2, lastO - 1, 1).getValues();
        for (var i = 0; i < colB.length; i++) {
          var noPO = String(colB[i][0]).trim();
          if (poList.indexOf(noPO) !== -1) {
            shO.getRange(i + 2, 29).setValue(invoiceNo); // AC
            shO.getRange(i + 2, 30).setValue(dateStr);   // AD
            var curProc = String(shO.getRange(i + 2, 17).getValue() || '').trim();
            if (curProc !== 'เรียบร้อย') {
              shO.getRange(i + 2, 17).setValue('เตรียมส่ง'); // Q = Process
            }
          }
        }
      }
    }

    return jsonOut({ status: 'ok', invoiceNo: invoiceNo, date: dateStr });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── updateInvoiceRecord: แก้ไขรายละเอียดใบกำกับเดิม (ไม่เปลี่ยนเลขที่) ──
// data: { invoiceNo, date(dd/MM/yyyy), type, customerCode, poList:[noPO,...],
//         subtotal, vat, total, status }
function updateInvoiceRecord(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Invoices');
    if (!sh) throw new Error('ไม่พบ sheet Invoices');

    var invoiceNo = String(data.invoiceNo || '').trim();
    if (!invoiceNo) throw new Error('ไม่พบเลขที่ใบกำกับ');

    var last = sh.getLastRow();
    var found = -1;
    if (last >= 1) {
      var col = sh.getRange(1, 1, last, 1).getValues();
      for (var i = 0; i < col.length; i++) {
        if (String(col[i][0]).trim() === invoiceNo) { found = i + 1; break; }
      }
    }
    if (found === -1) throw new Error('ไม่พบใบกำกับ ' + invoiceNo);

    var type = (data.type === 'short') ? 'short' : 'full';
    var poList = Array.isArray(data.poList) ? data.poList : [];
    var dateStr = String(data.date || '').trim();
    var items = Array.isArray(data.items) ? data.items : [];

    sh.getRange(found, 2, 1, 10).setValues([[
      dateStr, (type === 'short' ? 'อย่างย่อ' : 'เต็มรูป'),
      String(data.customerCode || ''), poList.join(', '),
      parseFloat(data.subtotal) || 0, parseFloat(data.vat) || 0, parseFloat(data.total) || 0,
      String(data.status || 'ออกแล้ว'), String(data.issuedBy || ''), JSON.stringify(items)
    ]]);

    return jsonOut({ status: 'ok', invoiceNo: invoiceNo });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── stampOrderPO: ประทับ invoiceNo/invoiceDate ลง col AC/AD ใน Order ──
// data: { noPOs:[...], invoiceNo, invoiceDate }
function stampOrderPO(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var shO = ss.getSheetByName('Order');
    if (!shO) throw new Error('ไม่พบ sheet Order');
    var noPOs = Array.isArray(data.noPOs)
      ? data.noPOs.map(function(x){ return String(x).trim(); }).filter(Boolean)
      : [];
    if (!noPOs.length) throw new Error('ไม่พบเลขที่ PO');
    var invoiceNo   = String(data.invoiceNo   || '').trim();
    var invoiceDate = String(data.invoiceDate || '').trim();
    var lastO = shO.getLastRow();
    var stamped = [];
    if (lastO >= 2) {
      var colB = shO.getRange(2, 2, lastO - 1, 1).getValues(); // col B = noPO
      for (var i = 0; i < colB.length; i++) {
        var noPO = String(colB[i][0]).trim();
        if (noPOs.indexOf(noPO) !== -1) {
          shO.getRange(i + 2, 29).setValue(invoiceNo);    // AC
          shO.getRange(i + 2, 30).setValue(invoiceDate);  // AD
          stamped.push(noPO);
        }
      }
    }
    return jsonOut({ status: 'ok', stamped: stamped });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── deleteInvoiceRecord: ลบใบกำกับออกจากชีต Invoices + เคลียร์ AC/AD ใน Order ──
function deleteInvoiceRecord(invoiceNo, revertProcess) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    invoiceNo = String(invoiceNo || '').trim();
    if (!invoiceNo) throw new Error('ไม่พบเลขที่ใบกำกับ');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Invoices');
    if (!sh) throw new Error('ไม่พบ sheet Invoices');
    var last = sh.getLastRow();
    var found = -1;
    for (var i = last; i >= 1; i--) {
      if (String(sh.getRange(i, 1).getValue()).trim() === invoiceNo) { found = i; break; }
    }
    if (found === -1) throw new Error('ไม่พบใบกำกับ ' + invoiceNo);
    sh.deleteRow(found);

    // ── เคลียร์ AC (เลขที่ใบกำกับ) / AD (วันที่ออกใบกำกับ) ใน Order ที่ตรงกับเลขนี้ ──
    var shO = ss.getSheetByName('Order');
    if (shO) {
      var lastO = shO.getLastRow();
      if (lastO >= 2) {
        var colAC = shO.getRange(2, 29, lastO - 1, 1).getValues(); // AC = 29
        for (var j = 0; j < colAC.length; j++) {
          if (String(colAC[j][0]).trim() === invoiceNo) {
            shO.getRange(j + 2, 29, 1, 2).clearContent(); // AC, AD
            if (revertProcess) shO.getRange(j + 2, 17).setValue(revertProcess); // Q = Process
          }
        }
      }
    }

    return jsonOut({ status: 'ok', deleted: invoiceNo });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ══════════════════════════════════════════════════════
// ══ ใบวางบิล (BillingNote) ═══════════════════════════════
// ══════════════════════════════════════════════════════
// คอลัมน์ BillingNote: A=เลขที่ใบวางบิล B=วันที่ออก C=รหัสลูกค้า
//   D=รายการเลขที่ใบกำกับ(comma) E=ยอดรวม F=หัก ณ ที่จ่าย 3% G=ยอดรับสุทธิ
//   H=เงื่อนไขชำระเงิน I=ผู้ออก

// ── เสนอเลขที่ใบวางบิลถัดไป รูปแบบ BI{ปี พ.ศ. 2 หลัก}/{เลขลำดับ 2 หลัก} ──
// อิงจากปี พ.ศ. ปัจจุบัน แยกรันแต่ละปี (ผู้ใช้แก้ไข/ตั้งใหม่เองได้เสมอ)
function _peekNextBillNo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('BillingNote');
  var now = new Date();
  var yy  = String((now.getFullYear() + 543) % 100);
  if (yy.length < 2) yy = '0' + yy;
  var prefix = 'BI' + yy + '/';
  var maxSeq = 0;
  if (sh) {
    var last = sh.getLastRow();
    if (last >= 1) {
      sh.getRange(1, 1, last, 1).getValues().forEach(function(r) {
        var v = String(r[0] || '').trim();
        if (v.indexOf(prefix) === 0) {
          var seq = parseInt(v.substring(prefix.length), 10);
          if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
        }
      });
    }
  }
  return prefix + ('0' + (maxSeq + 1)).slice(-2);
}

// ── saveBillingNote: บันทึกใบวางบิลใหม่ + มาร์คใบกำกับที่เลือกว่าวางบิลแล้ว ──
// data: { billNo, billDate(dd/MM/yyyy), customerCode, invoiceNos:[...],
//         sumTotal, whtAmount, netAmount, payTerm, issuedBy }
function saveBillingNote(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('BillingNote');
    if (!sh) sh = ss.insertSheet('BillingNote');

    var billNo = String(data.billNo || '').trim();
    if (!billNo) throw new Error('กรุณากรอกเลขที่ใบวางบิล');

    var invoiceNos = Array.isArray(data.invoiceNos) ? data.invoiceNos.map(function(x){ return String(x).trim(); }).filter(Boolean) : [];
    if (!invoiceNos.length) throw new Error('ไม่พบรายการใบกำกับที่เลือก');

    var dateStr = String(data.billDate || '').trim() ||
      Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'dd/MM/yyyy');

    sh.appendRow([
      billNo, dateStr, String(data.customerCode || ''), invoiceNos.join(', '),
      parseFloat(data.sumTotal) || 0, parseFloat(data.whtAmount) || 0, parseFloat(data.netAmount) || 0,
      String(data.payTerm || ''), String(data.issuedBy || '')
    ]);

    // ── มาร์คใบกำกับที่เลือกว่า "วางบิลแล้ว" (คอลัมน์ L=12 ของ Invoices) ──
    var shI = ss.getSheetByName('Invoices');
    if (shI) {
      var lastI = shI.getLastRow();
      if (lastI >= 1) {
        var colA = shI.getRange(1, 1, lastI, 1).getValues();
        for (var i = 0; i < colA.length; i++) {
          var invNo = String(colA[i][0]).trim();
          if (invoiceNos.indexOf(invNo) !== -1) {
            shI.getRange(i + 1, 12).setValue(billNo); // L = billNo
          }
        }
      }
    }

    return jsonOut({ status: 'ok', billNo: billNo, date: dateStr });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── getBillingNotes: ดึงประวัติใบวางบิลทั้งหมด (สำหรับดู/พิมพ์ซ้ำในอนาคต) ──
function getBillingNotes() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('BillingNote');
    if (!sh) return jsonOut({ status: 'ok', bills: [] });
    var last = sh.getLastRow();
    if (last < 1) return jsonOut({ status: 'ok', bills: [] });
    var vals = sh.getRange(1, 1, last, 9).getValues();
    var bills = vals.map(function(r) {
      var dateVal = r[1];
      var dateStr = (dateVal instanceof Date)
        ? Utilities.formatDate(dateVal, Session.getScriptTimeZone() || 'GMT+7', 'dd/MM/yyyy')
        : String(dateVal || '').trim();
      return {
        billNo:      String(r[0] || '').trim(),
        date:        dateStr,
        customerCode:String(r[2] || '').trim(),
        invoiceNos:  String(r[3] || '').trim(),
        sumTotal:    parseFloat(r[4]) || 0,
        whtAmount:   parseFloat(r[5]) || 0,
        netAmount:   parseFloat(r[6]) || 0,
        payTerm:     String(r[7] || '').trim(),
        issuedBy:    String(r[8] || '').trim()
      };
    }).filter(function(x) { return x.billNo; });
    return jsonOut({ status: 'ok', bills: bills });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── deleteBillingNote: ยกเลิกใบวางบิล — ลบแถวใน BillingNote + เคลียร์ billNo ในใบกำกับที่เกี่ยวข้อง ──
// data: { billNo }
function deleteBillingNote(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var billNo = String(data.billNo || '').trim();
    if (!billNo) throw new Error('ไม่พบเลขที่ใบวางบิล');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('BillingNote');
    var found = false;
    if (sh) {
      var last = sh.getLastRow();
      if (last >= 1) {
        var colA = sh.getRange(1, 1, last, 1).getValues();
        for (var i = colA.length - 1; i >= 0; i--) {
          if (String(colA[i][0]).trim() === billNo) {
            sh.deleteRow(i + 1);
            found = true;
            break;
          }
        }
      }
    }
    if (!found) throw new Error('ไม่พบใบวางบิลเลขที่ ' + billNo);

    // ── เคลียร์ billNo ในใบกำกับที่ถูกมาร์คไว้ (คอลัมน์ L=12 ของ Invoices) ──
    var shI = ss.getSheetByName('Invoices');
    if (shI) {
      var lastI = shI.getLastRow();
      if (lastI >= 1) {
        var colL = shI.getRange(1, 12, lastI, 1).getValues();
        for (var j = 0; j < colL.length; j++) {
          if (String(colL[j][0]).trim() === billNo) {
            shI.getRange(j + 1, 12).setValue('');
          }
        }
      }
    }

    return jsonOut({ status: 'ok', billNo: billNo });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── Async Telegram via one-time trigger ──────────────────────
function scheduleTelegram() {
  try {
    // ลบ trigger เก่าที่ค้างอยู่ก่อน (ป้องกัน accumulate)
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'sendTelegramOnChangeManual') {
        ScriptApp.deleteTrigger(t);
      }
    });
    // สร้าง trigger ใหม่ ยิงหลัง 5 วินาที
    ScriptApp.newTrigger('sendTelegramOnChangeManual')
      .timeBased()
      .after(5000)
      .create();
  } catch(e) {
    Logger.log('scheduleTelegram error: ' + e);
  }
}

// ══════════════════════════════════════════════════════
// ══ ใบส่งชุบ (PlatingNote) ═══════════════════════════════
// ══════════════════════════════════════════════════════
// คอลัมน์ PlatingNote: A=เลขที่ใบส่งชุบ B=วันที่ออก C=รหัสร้านชุบ
//   D=รายการ(JSON) E=ผู้ออก F=รายการเลขที่ Order(comma)

// ── เสนอเลขที่ใบส่งชุบถัดไป รูปแบบ ZP{ปี พ.ศ. 2 หลัก}/{เลขลำดับ 2 หลัก} ──
// อิงจากปี พ.ศ. ปัจจุบัน แยกรันแต่ละปี (ผู้ใช้แก้ไข/ตั้งใหม่เองได้เสมอ)
function _peekNextPlatingNo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('PlatingNote');
  var now = new Date();
  var yy  = String((now.getFullYear() + 543) % 100);
  if (yy.length < 2) yy = '0' + yy;
  var prefix = 'ZP' + yy + '/';
  var maxSeq = 0;
  if (sh) {
    var last = sh.getLastRow();
    if (last >= 1) {
      sh.getRange(1, 1, last, 1).getValues().forEach(function(r) {
        var v = String(r[0] || '').trim();
        if (v.indexOf(prefix) === 0) {
          var seq = parseInt(v.substring(prefix.length), 10);
          if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
        }
      });
    }
  }
  return prefix + ('0' + (maxSeq + 1)).slice(-2);
}

// ── sendChatMessage: บันทึกข้อความแชททีม (sheet "Chat") ──
// data: { sender, message }
function sendChatMessage(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Chat');
    if (!sh) {
      sh = ss.insertSheet('Chat');
      sh.appendRow(['เวลา', 'ผู้ส่ง', 'ข้อความ']);
    }
    var sender = String(data.sender || 'ไม่ระบุ').trim() || 'ไม่ระบุ';
    var message = String(data.message || '').trim();
    if (!message) throw new Error('ข้อความว่าง');
    sh.appendRow([new Date(), sender, message]);
    return jsonOut({ status: 'ok' });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── getChatMessages: ดึงข้อความแชททีมล่าสุด (สูงสุด 200 ข้อความ) ──
function getChatMessages() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Chat');
    if (!sh) return jsonOut({ status: 'ok', messages: [] });
    var last = sh.getLastRow();
    if (last < 2) return jsonOut({ status: 'ok', messages: [] });
    var startRow = Math.max(2, last - 199);
    var vals = sh.getRange(startRow, 1, last - startRow + 1, 3).getValues();
    var tz = Session.getScriptTimeZone() || 'GMT+7';
    var messages = vals
      .filter(function (row) { return String(row[2] || '').trim() !== ''; })
      .map(function (row) {
        var t = row[0];
        var timeStr = (t instanceof Date)
          ? Utilities.formatDate(t, tz, 'dd/MM/yyyy HH:mm')
          : String(t || '');
        return { time: timeStr, sender: String(row[1] || ''), message: String(row[2] || '') };
      });
    return jsonOut({ status: 'ok', messages: messages });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── savePlatingNote: บันทึกใบส่งชุบใหม่ + เปลี่ยนสถานะ Process ของ Order ที่เลือกเป็น "อยู่ระหว่างชุบ" ──
// data: { platingNo, platingDate(dd/MM/yyyy), supplierCode, items:[...], orderNos:[...], issuedBy }
function savePlatingNote(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('PlatingNote');
    if (!sh) sh = ss.insertSheet('PlatingNote');

    var platingNo = String(data.platingNo || '').trim();
    if (!platingNo) throw new Error('กรุณากรอกเลขที่ใบส่งชุบ');

    var items = Array.isArray(data.items) ? data.items : [];
    var orderNos = Array.isArray(data.orderNos) ? data.orderNos.map(function(x){ return String(x).trim(); }).filter(Boolean) : [];

    var dateStr = String(data.platingDate || '').trim() ||
      Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'dd/MM/yyyy');

    sh.appendRow([
      platingNo, dateStr, String(data.supplierCode || ''),
      JSON.stringify(items), String(data.issuedBy || ''), orderNos.join(', ')
    ]);

    // ── เปลี่ยน Process ของ Order ที่เกี่ยวข้องเป็น "ส่งชุป" ──
    if (orderNos.length) {
      var shO2 = ss.getSheetByName('Order');
      if (shO2) {
        var lastO2 = shO2.getLastRow();
        if (lastO2 >= 2) {
          var colB2 = shO2.getRange(2, 2, lastO2 - 1, 1).getValues();
          for (var ip = 0; ip < colB2.length; ip++) {
            if (orderNos.indexOf(String(colB2[ip][0]).trim()) !== -1) {
              var curProc2 = String(shO2.getRange(ip + 2, 17).getValue() || '').trim();
              if (curProc2 !== 'เตรียมส่ง' && curProc2 !== 'เรียบร้อย') {
                shO2.getRange(ip + 2, 17).setValue('ส่งชุป'); // Q = Process
              }
            }
          }
        }
      }
    }

    return jsonOut({ status: 'ok', platingNo: platingNo, date: dateStr });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── getPlatingNotes: ดึงประวัติใบส่งชุบทั้งหมด (สำหรับดู/พิมพ์ซ้ำ) ──
function getPlatingNotes() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('PlatingNote');
    if (!sh) return jsonOut({ status: 'ok', plates: [] });
    var last = sh.getLastRow();
    if (last < 1) return jsonOut({ status: 'ok', plates: [] });
    var vals = sh.getRange(1, 1, last, 6).getValues();
    var plates = vals.map(function(r) {
      var dateVal = r[1];
      var dateStr = (dateVal instanceof Date)
        ? Utilities.formatDate(dateVal, Session.getScriptTimeZone() || 'GMT+7', 'dd/MM/yyyy')
        : String(dateVal || '').trim();
      var items = [];
      try { items = r[3] ? JSON.parse(r[3]) : []; } catch (e2) { items = []; }
      return {
        platingNo:   String(r[0] || '').trim(),
        date:        dateStr,
        supplierCode:String(r[2] || '').trim(),
        items:       items,
        issuedBy:    String(r[4] || '').trim(),
        orderNos:    String(r[5] || '').trim()
      };
    }).filter(function(x) { return x.platingNo; });
    return jsonOut({ status: 'ok', plates: plates });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  }
}

// ── deletePlatingNote: ยกเลิกใบส่งชุบ — ลบแถวใน PlatingNote + revert Process ของ Order ──
// data: { platingNo, revertProcess?, orderNos?:[...] }
function deletePlatingNote(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var platingNo = String(data.platingNo || '').trim();
    if (!platingNo) throw new Error('ไม่พบเลขที่ใบส่งชุบ');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('PlatingNote');
    var found = false;
    var savedOrderNos = [];
    if (sh) {
      var last = sh.getLastRow();
      if (last >= 1) {
        var colA = sh.getRange(1, 1, last, 6).getValues();
        for (var i = colA.length - 1; i >= 0; i--) {
          if (String(colA[i][0]).trim() === platingNo) {
            // เก็บ orderNos จากคอลัมน์ที่ 6 (F) ก่อนลบ
            savedOrderNos = String(colA[i][5] || '').split(',').map(function(s){ return s.trim(); }).filter(Boolean);
            sh.deleteRow(i + 1);
            found = true;
            break;
          }
        }
      }
    }
    if (!found) throw new Error('ไม่พบใบส่งชุบเลขที่ ' + platingNo);

    // ── revert Process ของ Order ที่เกี่ยวข้อง (ถ้าผู้ใช้เลือก) ──
    var revertProcess = String(data.revertProcess || '').trim();
    var orderNos = Array.isArray(data.orderNos) ? data.orderNos : savedOrderNos;
    if (revertProcess && orderNos.length) {
      var shO = ss.getSheetByName('Order');
      if (shO) {
        var lastO = shO.getLastRow();
        if (lastO >= 2) {
          var colB = shO.getRange(2, 2, lastO - 1, 1).getValues();
          for (var j = 0; j < colB.length; j++) {
            if (orderNos.indexOf(String(colB[j][0]).trim()) !== -1) {
              shO.getRange(j + 2, 17).setValue(revertProcess); // Q = Process
            }
          }
        }
      }
    }

    return jsonOut({ status: 'ok', platingNo: platingNo });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── updatePlatingNote: แก้ไขใบส่งชุบที่มีอยู่ (วันที่, ร้านชุบ, รายการสินค้า) ──
// data: { platingNo, date(dd/MM/yyyy)?, supplierCode?, items?:[...] }
function updatePlatingNote(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var platingNo = String(data.platingNo || '').trim();
    if (!platingNo) throw new Error('ไม่พบเลขที่ใบส่งชุบ');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('PlatingNote');
    if (!sh) throw new Error('ไม่พบ Sheet PlatingNote');
    var last = sh.getLastRow();
    if (last < 1) throw new Error('ไม่พบข้อมูลในชีต PlatingNote');
    var colA = sh.getRange(1, 1, last, 1).getValues();
    var found = false;
    for (var i = 0; i < colA.length; i++) {
      if (String(colA[i][0]).trim() === platingNo) {
        if (data.date !== undefined && data.date !== null)
          sh.getRange(i + 1, 2).setValue(String(data.date));
        if (data.supplierCode !== undefined && data.supplierCode !== null)
          sh.getRange(i + 1, 3).setValue(String(data.supplierCode));
        if (data.items !== undefined && data.items !== null)
          sh.getRange(i + 1, 4).setValue(JSON.stringify(data.items));
        found = true;
        break;
      }
    }
    if (!found) throw new Error('ไม่พบใบส่งชุบเลขที่ ' + platingNo);
    return jsonOut({ status: 'ok', platingNo: platingNo });
  } catch (e) {
    return jsonOut({ status: 'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════════
//  BACKUP ระบบสำรองข้อมูล — คัดลอกไฟล์ Spreadsheet ทั้งไฟล์เป็นรายวัน (ตี 4)
// ═══════════════════════════════════════════════════════════════

var BACKUP_FOLDER_NAME = 'PTTS Backups';  // ชื่อโฟลเดอร์เก็บไฟล์สำรองใน Drive
var BACKUP_KEEP_COUNT  = 7;               // เก็บไฟล์สำรองล่าสุดไว้กี่ไฟล์ (เกินกว่านี้ลบของเก่าออก) — รายวัน 7 ไฟล์ = 1 อาทิตย์

// ── backupSpreadsheet: คัดลอกไฟล์ Spreadsheet ปัจจุบันทั้งไฟล์ ไปเก็บในโฟลเดอร์ "PTTS Backups" ──
// เรียกอัตโนมัติด้วย time-driven trigger (ดู createDailyBackupTrigger)
// หรือเรียกเองได้ผ่าน doGet ?action=runBackupNow (สำหรับทดสอบ)
function backupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());

  // หา/สร้างโฟลเดอร์ปลายทาง
  var folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);

  // ตั้งชื่อไฟล์สำรองตามวันเวลา
  var tz = Session.getScriptTimeZone() || 'GMT+7';
  var ts = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH'h'mm");
  var backupName = 'Backup - ' + ss.getName() + ' - ' + ts;

  var copy = file.makeCopy(backupName, folder);

  // ลบไฟล์สำรองเก่าที่เกินจำนวนที่ตั้งไว้ (เรียงตามวันที่สร้าง เก่าสุดลบก่อน)
  var files = [];
  var it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    files.push(f);
  }
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  for (var i = BACKUP_KEEP_COUNT; i < files.length; i++) {
    files[i].setTrashed(true);
  }

  return { name: copy.getName(), id: copy.getId(), url: copy.getUrl(), folder: folder.getName() };
}

// ── createDailyBackupTrigger: รันครั้งเดียวจาก Apps Script editor เพื่อตั้งระบบ backup อัตโนมัติทุกวัน (ตี 4) ──
// วิธีใช้: เปิด Extensions → Apps Script → เลือกฟังก์ชัน "createDailyBackupTrigger" จาก dropdown ด้านบน → กด Run 
// ── saveGeneralQuotation: บันทึกใบเสนอราคาทั่วไป → ชีต Quotations ──
function saveGeneralQuotation(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Quotations');
    if (!sh) {
      sh = ss.insertSheet('Quotations');
      sh.appendRow(['refNo','date','customer','items','remark','subtotal','vatAmt','grand','createdAt']);
      sh.getRange(1,1,1,9).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    // ถ้า refNo ซ้ำ — update แถวเดิม
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.refNo)) {
        sh.getRange(i + 1, 1, 1, 9).setValues([[
          data.refNo, data.date, data.customer,
          JSON.stringify(data.items || []),
          data.remark || '',
          data.subtotal || 0, data.vatAmt || 0, data.grand || 0,
          new Date().toISOString()
        ]]);
        sh.getRange(i + 1, 2).setNumberFormat('@STRING@'); // force date as text
        return jsonOut({ status:'ok', updated: true });
      }
    }
    sh.appendRow([
      data.refNo, data.date, data.customer,
      JSON.stringify(data.items || []),
      data.remark || '',
      data.subtotal || 0, data.vatAmt || 0, data.grand || 0,
      new Date().toISOString()
    ]);
    sh.getRange(sh.getLastRow(), 2).setNumberFormat('@STRING@'); // force date as text
    return jsonOut({ status:'ok', updated: false });
  } catch(err) {
    return jsonOut({ status:'error', message: err.toString() });
  }
}

// ── deleteGeneralQuotation: ลบใบเสนอราคาทั่วไป ──
function deleteGeneralQuotation(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Quotations');
    if (!sh) return jsonOut({ status:'ok' });
    var rows = sh.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) === String(data.refNo)) {
        sh.deleteRow(i + 1);
        return jsonOut({ status:'ok' });
      }
    }
    return jsonOut({ status:'ok' });
  } catch(err) {
    return jsonOut({ status:'error', message: err.toString() });
  }
}


// ══════════════════════════════════════════════════════════════
//  RFQ — ใบขอราคา  (sheet "RFQ")
// ══════════════════════════════════════════════════════════════
// คอลัมน์: rfqNo | date | supplier | items (JSON) | remark | createdAt

function saveRFQ(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('RFQ');
    if (!sh) {
      sh = ss.insertSheet('RFQ');
      sh.appendRow(['rfqNo', 'date', 'supplier', 'items', 'remark', 'createdAt']);
      sh.getRange(1, 1, 1, 6).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    var rows = sh.getDataRange().getValues();
    // upsert: ถ้า rfqNo ซ้ำให้ update แถวเดิม
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.rfqNo)) {
        sh.getRange(i + 1, 1, 1, 6).setValues([[
          data.rfqNo, data.date || '', data.supplier || '',
          typeof data.items === 'string' ? data.items : JSON.stringify(data.items || []),
          data.remark || '', new Date().toISOString()
        ]]);
        return jsonOut({ status: 'ok', updated: true });
      }
    }
    sh.appendRow([
      data.rfqNo, data.date || '', data.supplier || '',
      typeof data.items === 'string' ? data.items : JSON.stringify(data.items || []),
      data.remark || '', new Date().toISOString()
    ]);
    return jsonOut({ status: 'ok', updated: false });
  } catch (err) {
    return jsonOut({ status: 'error', message: err.toString() });
  }
}

function getRFQList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('RFQ');
    if (!sh || sh.getLastRow() < 2) return jsonOut({ status: 'ok', data: [] });
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
    var list = rows.filter(function(r) { return String(r[0]).trim(); }).map(function(r) {
      var itemsParsed = [];
      try { itemsParsed = JSON.parse(String(r[3]) || '[]'); } catch(e) {}
      var dateVal = r[1];
      var dateStr = '';
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
      } else {
        dateStr = String(dateVal);
      }
      return {
        rfqNo:     String(r[0]),
        date:      dateStr,
        supplier:  String(r[2]),
        items:     itemsParsed,
        remark:    String(r[4]),
        createdAt: String(r[5])
      };
    });
    return jsonOut({ status: 'ok', data: list });
  } catch (err) {
    return jsonOut({ status: 'error', message: err.toString() });
  }
}

function deleteRFQ(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('RFQ');
    if (!sh) return jsonOut({ status: 'ok' });
    var rows = sh.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) === String(data.rfqNo)) {
        sh.deleteRow(i + 1);
        return jsonOut({ status: 'ok' });
      }
    }
    return jsonOut({ status: 'ok' });
  } catch (err) {
    return jsonOut({ status: 'error', message: err.toString() });
  }
}

// ── WI (Work Instruction) ──────────────────────────────────────
function getWIList() {
  // อ่าน 1 step = 1 row แล้ว group ตาม wiId
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('WI');
    if (!sh) return jsonOut({ status: 'ok', data: [] });
    var rows = sh.getDataRange().getValues();
    if (rows.length < 2) return jsonOut({ status: 'ok', data: [] });
    var map = {};
    var order = [];
    rows.slice(1).forEach(function(r) {
      var wiId = String(r[0]||''); if (!wiId) return;
      if (!map[wiId]) {
        var _dt9 = r[9];
        var _updAt = (_dt9 instanceof Date)
          ? Utilities.formatDate(_dt9, 'Asia/Bangkok', 'dd/MM/yyyy')
          : String(_dt9||'').trim();
        map[wiId] = { wiId:wiId, od:r[1], id:r[2], h:r[3], workType:String(r[4]||''), remark:String(r[5]||''), updatedAt:_updAt, steps:[] };
        order.push(wiId);
      }
      var stepIdx = parseInt(r[6])||0;
      var caption = String(r[7]||'');
      var img     = String(r[8]||'');
      if (caption || img) map[wiId].steps[stepIdx] = { caption:caption, img:img };
    });
    var list = order.map(function(id){ return map[id]; });
    return jsonOut({ status: 'ok', data: list });
  } catch (err) {
    return jsonOut({ status: 'error', message: err.toString() });
  }
}


function uploadWIImage(data) {
  try {
    var folderId = '1lI_bHFKULL0fQYYIfNkUl7p148_QflSu';
    var folder = DriveApp.getFolderById(folderId);
    var base64Data = String(data.base64 || '').replace(/^data:image\/\w+;base64,/, '');
    var mimeType = data.mimeType || 'image/jpeg';
    var filename = data.filename || ('wi_' + new Date().getTime() + '.jpg');
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, filename);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://lh3.googleusercontent.com/d/' + file.getId();
    return jsonOut({ status: 'ok', url: url });
  } catch(err) {
    return jsonOut({ status: 'error', message: err.toString() });
  }
}

function saveWI(data) {
  // เก็บ 1 step = 1 row เพื่อหลีกเลี่ยง 50K char limit ของ Sheets cell
  // cols: wiId | od | id | h | workType | remark | stepIndex | stepCaption | stepImg | updatedAt
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('WI');
    if (!sh) {
      sh = ss.insertSheet('WI');
      sh.appendRow(['wiId','od','id','h','workType','remark','stepIndex','stepCaption','stepImg','updatedAt']);
    }
    var _d = new Date();
    var now = Utilities.formatDate(_d,'Asia/Bangkok','dd/MM/') + String(parseInt(Utilities.formatDate(_d,'Asia/Bangkok','yyyy'))+543);
    var steps = data.steps || [];
    // ลบ row เก่าของ wiId นี้ทั้งหมด (upsert)
    var rows = sh.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) === String(data.wiId)) sh.deleteRow(i + 1);
    }
    // เขียน row ใหม่ทีละ step
    if (steps.length === 0) {
      sh.appendRow([data.wiId, data.od, data.id||0, data.h, data.workType||'', data.remark||'', 0, '', '', now]);
    } else {
      steps.forEach(function(s, idx) {
        sh.appendRow([data.wiId, data.od, data.id||0, data.h, data.workType||'', data.remark||'', idx, s.caption||'', s.img||'', now]);
      });
    }
    return jsonOut({ status: 'ok', wiId: data.wiId });
  } catch(err) {
    return jsonOut({ status: 'error', message: err.toString() });
  }
}

function getSalesDashboard(months) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Order');
    if (!sh) return jsonOut({ status: 'ok', months: [], topProducts: [], topCustomers: [], kpi: {} });
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return jsonOut({ status: 'ok', months: [], topProducts: [], topCustomers: [], kpi: {} });
    var numCols = Math.max(38, sh.getLastColumn());
    var values = sh.getRange(2, 1, lastRow - 1, numCols).getValues();

    var now = new Date();
    // cutoff = วันแรกของเดือน N เดือนที่แล้ว
    var cutoff = new Date(now.getFullYear(), now.getMonth() - months + 1, 1);
    var prevCutoff = new Date(now.getFullYear(), now.getMonth() - months * 2 + 1, 1);

    function parseRow(r) {
      var rawDate = r[2];
      var dateStr;
      if (rawDate instanceof Date) {
        // Sheets คืน Date object — แปลงเป็น dd/MM/yyyy (CE) ก่อน
        var dd = String(rawDate.getDate()).padStart(2,'0');
        var mm = String(rawDate.getMonth()+1).padStart(2,'0');
        var yyyy = rawDate.getFullYear(); // CE แล้ว (Apps Script ใช้ CE)
        dateStr = dd + '/' + mm + '/' + yyyy;
      } else {
        dateStr = String(rawDate || '');
      }
      if (!dateStr) return null;
      var parts = dateStr.split('/');
      if (parts.length !== 3) return null;
      var yr = parseInt(parts[2]);
      var yearCE = yr > 2400 ? yr - 543 : yr;
      var mo = parseInt(parts[1]);
      var dy = parseInt(parts[0]);
      var d = new Date(yearCE, mo - 1, dy);
      if (isNaN(d.getTime())) return null;
      var monthKey = yearCE + '-' + String(mo).padStart(2, '0');
      var workType = String(r[4]  || '');
      var customer = String(r[27] || '');
      var qty      = Number(r[6])  || 0;
      var totalTax = Number(r[23]) || 0;
      var price    = Number(r[10]) || 0;
      var revenue  = totalTax > 0 ? totalTax : price * qty;
      var od  = String(r[34] || '');
      var id_ = String(r[35] || '');
      var h   = String(r[36] || '');
      var spec = od ? ('OD' + od + (parseFloat(id_) > 0 ? 'xID' + id_ : '') + (parseFloat(h) > 0 ? 'xH' + h : '') + 'mm.') : '';
      return { monthKey: monthKey, month: mo, yearCE: yearCE, d: d,
               workType: workType, customer: customer, qty: qty, revenue: revenue, spec: spec };
    }

    // ── Aggregate ──────────────────────────────────────
    var monthMap   = {};
    var productMap = {};
    var customerMap = {};
    var kpi = { revenue:0, count:0, qty:0, prevRevenue:0, prevQty:0, prevCount:0 };

    for (var i = 0; i < values.length; i++) {
      var row = parseRow(values[i]);
      if (!row) continue;

      // prev period
      if (row.d >= prevCutoff && row.d < cutoff) {
        kpi.prevRevenue += row.revenue;
        kpi.prevQty     += row.qty;
        kpi.prevCount   += 1;
      }
      // current period
      if (row.d < cutoff) continue;

      kpi.revenue += row.revenue;
      kpi.qty     += row.qty;
      kpi.count   += 1;

      // monthly
      if (!monthMap[row.monthKey]) monthMap[row.monthKey] = { monthKey: row.monthKey, month: row.month, year: row.yearCE, revenue: 0, count: 0 };
      monthMap[row.monthKey].revenue += row.revenue;
      monthMap[row.monthKey].count   += 1;

      // product (workType)
      if (!row.workType) continue;
      if (!productMap[row.workType]) productMap[row.workType] = { workType: row.workType, count: 0, revenue: 0, qty: 0, specMap: {} };
      productMap[row.workType].count   += 1;
      productMap[row.workType].revenue += row.revenue;
      productMap[row.workType].qty     += row.qty;
      if (row.spec) {
        productMap[row.workType].specMap[row.spec] = (productMap[row.workType].specMap[row.spec] || 0) + 1;
      }

      // customer
      if (!row.customer) continue;
      if (!customerMap[row.customer]) customerMap[row.customer] = { customer: row.customer, count: 0, revenue: 0 };
      customerMap[row.customer].count   += 1;
      customerMap[row.customer].revenue += row.revenue;
    }

    // ── Build month array (sorted) ────────────────────
    var monthArr = Object.values(monthMap).sort(function(a,b){ return a.monthKey > b.monthKey ? 1 : -1; });

    // ── Top 5 products ────────────────────────────────
    var topProducts = Object.values(productMap).sort(function(a,b){ return b.revenue - a.revenue; }).slice(0,5).map(function(p) {
      var specText = '';
      var maxFreq = 0;
      Object.keys(p.specMap).forEach(function(s){ if (p.specMap[s] > maxFreq){ maxFreq = p.specMap[s]; specText = s; } });
      return { workType: p.workType, count: p.count, revenue: p.revenue, qty: p.qty, specText: specText };
    });

    // ── Top 5 customers ───────────────────────────────
    var topCustomers = Object.values(customerMap).sort(function(a,b){ return b.revenue - a.revenue; }).slice(0,5);

    return jsonOut({ status: 'ok', months: monthArr, topProducts: topProducts, topCustomers: topCustomers, kpi: kpi });

  } catch(err) {
    return jsonOut({ status: 'error', message: err.toString() });
  }
}

// ── Inspection: getInspections ────────────────────────────
// Schema: inspId | type | orderId | orderRef | date | inspector |
//         workType | spec | qtyCheck | qtyPass | qtyFail | result | note | createdAt | items(JSON)
//         [0]      [1]    [2]        [3]        [4]   [5]          [6]       [7]   [8]       [9]      [10]     [11]    [12]  [13]         [14]
function getInspections(type, orderId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Inspections');
    if (!sh || sh.getLastRow() < 2) return jsonOut({ status:'ok', rows:[] });
    var data = sh.getDataRange().getValues();
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (!r[0] || String(r[0]) === 'inspId') continue;
      if (type    && String(r[1]) !== type)    continue;
      if (orderId && String(r[2]) !== orderId) continue;
      rows.push({
        inspId:   String(r[0]),
        type:     String(r[1]),
        orderId:  String(r[2]),
        orderRef: String(r[3]),
        date:     r[4] instanceof Date ? Utilities.formatDate(r[4],'Asia/Bangkok','dd/MM/yyyy') : String(r[4]),
        inspector:String(r[5]),
        workType: String(r[6]),
        spec:     String(r[7]),
        qtyCheck: Number(r[8])  || 0,
        qtyPass:  Number(r[9])  || 0,
        qtyFail:  Number(r[10]) || 0,
        result:   String(r[11]),
        note:     String(r[12]),
        createdAt:String(r[13])
      });
    }
    rows.reverse();
    return jsonOut({ status:'ok', rows:rows });
  } catch(err) {
    return jsonOut({ status:'error', message: err.toString() });
  }
}

// ── Inspection: getInspectionDetail ──────────────────────
function getInspectionDetail(inspId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Inspections');
    if (!sh || sh.getLastRow() < 2) return jsonOut({ status:'ok', header:null, items:[] });
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(inspId)) continue;
      var r = data[i];
      var header = {
        inspId:   String(r[0]),  type:     String(r[1]),
        orderId:  String(r[2]),  orderRef: String(r[3]),
        date:     r[4] instanceof Date ? Utilities.formatDate(r[4],'Asia/Bangkok','dd/MM/yyyy') : String(r[4]),
        inspector:String(r[5]),  workType: String(r[6]),
        spec:     String(r[7]),
        qtyCheck: Number(r[8])||0, qtyPass: Number(r[9])||0, qtyFail: Number(r[10])||0,
        result:   String(r[11]), note: String(r[12])
      };
      var items = [];
      try { items = JSON.parse(String(r[14]) || '[]'); } catch(e) { items = []; }
      return jsonOut({ status:'ok', header:header, items:items });
    }
    return jsonOut({ status:'ok', header:null, items:[] });
  } catch(err) {
    return jsonOut({ status:'error', message: err.toString() });
  }
}

// ── Inspection: saveInspection ────────────────────────────
function saveInspection(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Inspections');
    if (!sh) sh = ss.insertSheet('Inspections');
    if (sh.getLastRow() === 0) {
      sh.appendRow(['inspId','type','orderId','orderRef','date','inspector',
                    'workType','spec','qtyCheck','qtyPass','qtyFail','result','note','createdAt','items']);
    }
    // ลบ Insp_Items sheet ถ้ายังมีอยู่
    var shi = ss.getSheetByName('Insp_Items');
    if (shi) ss.deleteSheet(shi);

    var type   = String(data.type || 'OQC');
    var year   = new Date().getFullYear();
    var seqNo  = Math.max(sh.getLastRow(), 1);
    var inspId = type + '-' + year + '-' + String(seqNo).padStart(4,'0');
    var now    = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
    var itemsJson = JSON.stringify(data.items || []);

    sh.appendRow([
      inspId,
      type,
      String(data.orderId   || ''),
      String(data.orderRef  || ''),
      String(data.date      || ''),
      String(data.inspector || ''),
      String(data.workType  || ''),
      String(data.spec      || ''),
      Number(data.qtyCheck) || 0,
      Number(data.qtyPass)  || 0,
      Number(data.qtyFail)  || 0,
      String(data.result    || 'OK'),
      String(data.note      || ''),
      now,
      itemsJson
    ]);
    return jsonOut({ status:'ok', inspId:inspId });
  } catch(err) {
    return jsonOut({ status:'error', message: err.toString() });
  }
}

// ── Inspection: deleteInspection ─────────────────────────
function deleteInspection(inspId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Inspections');
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getDataRange().getValues();
      for (var i = rows.length - 1; i >= 1; i--) {
        if (String(rows[i][0]) === String(inspId)) { sh.deleteRow(i + 1); break; }
      }
    }
    return jsonOut({ status:'ok' });
  } catch(err) {
    return jsonOut({ status:'error', message: err.toString() });
  }
}

// ── Inspection: updateInspection ─────────────────────────
function updateInspection(data) {
  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var sh     = ss.getSheetByName('Inspections');
    if (!sh)   return jsonOut({ status:'error', message:'ไม่พบ sheet Inspections' });
    var inspId = String(data.inspId || '');
    if (!inspId) return jsonOut({ status:'error', message:'ไม่มี inspId' });
    var rows   = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== inspId) continue;
      var rowNum    = i + 1;
      var itemsJson = JSON.stringify(data.items || []);
      sh.getRange(rowNum, 2, 1, 14).setValues([[
        String(data.type      || ''),
        String(data.orderId   || ''),
        String(data.orderRef  || ''),
        String(data.date      || ''),
        String(data.inspector || ''),
        String(data.workType  || ''),
        String(data.spec      || ''),
        Number(data.qtyCheck) || 0,
        Number(data.qtyPass)  || 0,
        Number(data.qtyFail)  || 0,
        String(data.result    || 'OK'),
        String(data.note      || ''),
        String(rows[i][13]),   // createdAt — คงค่าเดิม
        itemsJson
      ]]);
      return jsonOut({ status:'ok', inspId:inspId });
    }
    return jsonOut({ status:'error', message:'ไม่พบ inspId: ' + inspId });
  } catch(err) {
    return jsonOut({ status:'error', message: err.toString() });
  }
}



// ══════════════════════════════════════════════════════════════
// MAT STOCK CONTROL
// ══════════════════════════════════════════════════════════════

// ── getMatStock: ดึง stock ทั้งหมดจากชีต MatStock ──
function getMatStock() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('MatStock');
    if (!sh) return jsonOut({ status:'ok', rows:[] });
    var last = sh.getLastRow();
    if (last < 2) return jsonOut({ status:'ok', rows:[] });
    var vals = sh.getRange(2, 1, last - 1, 8).getValues();
    // ป้องกัน Google Sheets คืน Date object สำหรับ cell ที่มี format วันที่
    var safeNum = function(v) {
      if (v instanceof Date) return 0;
      var n = parseFloat(v);
      return isNaN(n) ? 0 : n;
    };
    var rows = vals.map(function(r) {
      return {
        matCode:     String(r[0]||'').trim(),
        matName:     String(r[1]||'').trim(),
        unit:        String(r[2]||'แผ่น').trim(),
        stockQty:    safeNum(r[3]),
        minStock:    safeNum(r[4]),
        orderQty:    safeNum(r[5]),
        suppCode:    String(r[6]||'').trim(),
        lastUpdated: (r[7] instanceof Date)
          ? Utilities.formatDate(r[7], Session.getScriptTimeZone()||'Asia/Bangkok', 'dd/MM/yyyy HH:mm')
          : String(r[7]||'').trim()
      };
    }).filter(function(r){ return r.matCode; });
    return jsonOut({ status:'ok', rows:rows });
  } catch(e) {
    return jsonOut({ status:'error', message:e.message });
  }
}

// ── saveMatStock: เพิ่ม/แก้ไข 1 แถวใน MatStock (CRUD) ──
// data: { matCode, matName, unit, stockQty, minStock, orderQty, suppCode }
function saveMatStock(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('MatStock');
    if (!sh) {
      sh = ss.insertSheet('MatStock');
      sh.appendRow(['matCode','matName','unit','stockQty','minStock','orderQty','suppCode','lastUpdated']);
    }
    var code = String(data.matCode || '').trim();
    if (!code) throw new Error('matCode is required');
    var now  = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
    var row  = [
      code,
      String(data.matName  || '').trim(),
      String(data.unit     || 'แผ่น').trim(),
      Number(data.stockQty || 0),
      Number(data.minStock || 0),
      Number(data.orderQty || 0),
      String(data.suppCode || '').trim(),
      now
    ];
    var last = sh.getLastRow();
    var found = -1;
    if (last >= 2) {
      var codes = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < codes.length; i++) {
        if (String(codes[i][0]).trim() === code) { found = i + 2; break; }
      }
    }
    if (found !== -1) {
      sh.getRange(found, 1, 1, 8).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    if (data.delete) {
      // deleteMatStock: ลบแถว
      if (found !== -1) sh.deleteRow(found);
      return jsonOut({ status:'ok', action:'deleted', matCode:code });
    }
    return jsonOut({ status:'ok', action: found !== -1 ? 'updated' : 'added', matCode:code });
  } catch(e) {
    return jsonOut({ status:'error', message:e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── updateMatStockBatch: หักหรือเพิ่ม stock หลาย MAT พร้อมกัน + บันทึก Movement ──
// data: {
//   movements: [{ matCode, qty, type:'IN'|'OUT'|'ADJUST', ref, note }],
//   by: 'ชื่อผู้บันทึก'
// }
function updateMatStockBatch(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var stockSh= ss.getSheetByName('MatStock');
    var moveSh = ss.getSheetByName('MatMovement');
    if (!stockSh) throw new Error('ไม่พบชีต MatStock');
    if (!moveSh) {
      moveSh = ss.insertSheet('MatMovement');
      moveSh.appendRow(['date','matCode','type','qty','stockAfter','ref','note','by']);
    }

    var moves   = data.movements || [];
    var by      = String(data.by || 'system').trim();
    var tz      = Session.getScriptTimeZone() || 'Asia/Bangkok';
    var now     = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
    var results = [];

    // โหลด MatStock ทั้งหมดไว้ใน memory
    var lastStock = stockSh.getLastRow();
    var stockVals = lastStock >= 2 ? stockSh.getRange(2, 1, lastStock - 1, 8).getValues() : [];
    var stockMap  = {}; // code → rowIndex (1-based ใน sheet = i+2)
    stockVals.forEach(function(r, i) {
      var c = String(r[0]||'').trim();
      if (c) stockMap[c] = { rowIdx: i + 2, data: r };
    });

    moves.forEach(function(m) {
      var code = String(m.matCode || '').trim();
      if (!code) return;
      var qty  = Number(m.qty)  || 0;
      var type = String(m.type || 'OUT').toUpperCase();
      var ref  = String(m.ref  || '').trim();
      var note = String(m.note || '').trim();

      var entry = stockMap[code];
      var currentQty = entry ? Number(entry.data[3]||0) : 0;
      var newQty;
      if (type === 'IN')     newQty = currentQty + qty;
      else if (type === 'OUT') newQty = currentQty - qty;
      else                   newQty = qty; // ADJUST = ตั้งค่าใหม่

      if (entry) {
        stockSh.getRange(entry.rowIdx, 4, 1, 2).setValues([[newQty, now]]);
      }
      // บันทึก Movement
      moveSh.appendRow([now, code, type, qty, newQty, ref, note, by]);
      results.push({ matCode:code, type:type, qty:qty, stockAfter:newQty });
    });

    return jsonOut({ status:'ok', results:results });
  } catch(e) {
    return jsonOut({ status:'error', message:e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── adjustMatStockBatch: ปรับ stock จากการตรวจนับจริง ──
// data: { items: [{matCode, newQty, note}], by, ref }
// บันทึก MatMovement ประเภท 'COUNT' (ตรวจนับ) สำหรับทุกรายการที่มีส่วนต่าง
function adjustMatStockBatch(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var stockSh = ss.getSheetByName('MatStock');
    var moveSh  = ss.getSheetByName('MatMovement');
    if (!stockSh) throw new Error('ไม่พบชีต MatStock');
    if (!moveSh) {
      moveSh = ss.insertSheet('MatMovement');
      moveSh.appendRow(['date','matCode','type','qty','stockAfter','ref','note','by']);
    }

    var items  = data.items  || [];
    var by     = String(data.by  || 'system').trim();
    var ref    = String(data.ref || '').trim();
    var tz     = Session.getScriptTimeZone() || 'Asia/Bangkok';
    var now    = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
    var results = [];

    // โหลด MatStock ทั้งหมดไว้ใน memory
    var lastRow   = stockSh.getLastRow();
    var stockVals = lastRow >= 2 ? stockSh.getRange(2, 1, lastRow - 1, 8).getValues() : [];
    var stockMap  = {};
    stockVals.forEach(function(r, i) {
      var c = String(r[0]||'').trim();
      if (c) stockMap[c] = { rowIdx: i + 2, currentQty: Number(r[3]||0) };
    });

    items.forEach(function(item) {
      var code   = String(item.matCode || '').trim();
      if (!code) return;
      var newQty = Number(item.newQty);
      if (isNaN(newQty)) return;
      var note   = String(item.note || '').trim();

      var entry = stockMap[code];
      var oldQty = entry ? entry.currentQty : 0;
      var diff   = newQty - oldQty; // บวก = เพิ่ม, ลบ = ขาด

      if (entry) {
        stockSh.getRange(entry.rowIdx, 4, 1, 2).setValues([[newQty, now]]);
      }
      // บันทึก Movement ทุกรายการ (แม้ diff=0 เพื่อเป็นหลักฐานว่าตรวจนับแล้ว)
      moveSh.appendRow([now, code, 'COUNT', diff, newQty, ref, note || 'ตรวจนับจริง', by]);
      results.push({ matCode: code, oldQty: oldQty, newQty: newQty, diff: diff });
    });

    return jsonOut({ status:'ok', results: results });
  } catch(e) {
    return jsonOut({ status:'error', message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ── getMatMovement: ดูประวัติ Movement (กรองตาม matCode ถ้าระบุ) ──
function getMatMovement(matCode) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('MatMovement');
    if (!sh) return jsonOut({ status:'ok', rows:[] });
    var last = sh.getLastRow();
    if (last < 2) return jsonOut({ status:'ok', rows:[] });
    var vals = sh.getRange(2, 1, last - 1, 8).getValues();
    var code = String(matCode || '').trim();
    var rows = vals.map(function(r) {
      return {
        date:       String(r[0]||'').trim(),
        matCode:    String(r[1]||'').trim(),
          type:       String(r[2]||'').trim(),
        qty:        Number(r[3]||0),
        stockAfter: Number(r[4]||0),
        ref:        String(r[5]||'').trim(),
        note:       String(r[6]||'').trim(),
        by:         String(r[7]||'').trim()
      };
    }).filter(function(r){
      return r.matCode && (!code || r.matCode === code);
    });
    rows.reverse(); // ล่าสุดก่อน
    return jsonOut({ status:'ok', rows:rows });
  } catch(e) {
    return jsonOut({ status:'error', message:e.message });
  }
}

// ══════════════════════════════════════════════════════════════
// HR Module Functions
// ══════════════════════════════════════════════════════════════

function hrLogin(data) {
  try {
    var phone = String(data.phone || '').replace(/[^0-9]/g, '');
    var pin   = String(data.pin   || '');
    if (!phone || !pin) return jsonOut({ status: 'error', message: 'กรุณากรอกเบอร์โทรและ PIN' });
    // normalize 9 digits → add leading 0
    if (phone.length === 9) phone = '0' + phone;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('HR_Employees');
    if (!sh) return jsonOut({ status: 'error', message: 'ไม่พบชีต HR_Employees' });

    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      var rowPhone = String(r[14] || '').replace(/[^0-9]/g, '');
      if (rowPhone.length === 9) rowPhone = '0' + rowPhone;
      var rowPin = String(r[13] || '');
      if (rowPhone !== phone || rowPin !== pin) continue;
      // ตรวจ active
      var active = String(r[10] || '').toLowerCase();
      if (active === 'false' || active === '0' || active === 'inactive') {
        return jsonOut({ status: 'error', message: 'บัญชีถูกระงับ กรุณาติดต่อ HR' });
      }
      return jsonOut({
        status: 'ok',
        emp: {
          empId:         String(r[0]  || ''),
          name:          String(r[1]  || ''),
          dept:          String(r[2]  || ''),
          position:      String(r[3]  || ''),
          type:          String(r[4]  || 'monthly'),
          advanceBudget: parseFloat(r[12]) || 0,
          phone:         String(r[14] || ''),
          profileUrl:    String(r[17] || ''),
        }
      });
    }
    return jsonOut({ status: 'error', message: 'เบอร์โทรหรือ PIN ไม่ถูกต้อง' });
  } catch(e) {
    return jsonOut({ status: 'error', message: e.toString() });
  }
}

function getHREmployees() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('HR_Employees');
    if (!sh) return jsonOut({ data: [] });
    var rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return jsonOut({ data: [] });
    var data = rows.slice(1).map(function(r) {
      return {
        empId: String(r[0]||''), name: String(r[1]||''), dept: String(r[2]||''),
        position: String(r[3]||''), type: String(r[4]||'monthly'),
        salary: Number(r[5]||0), dailyRate: Number(r[6]||0),
        otRateWD: Number(r[7]||100), otRateSun: Number(r[8]||200),
        startDate: String(r[9]||''), active: r[10] !== false && r[10] !== 'FALSE',
        payCycle: String(r[11]||'default'), advanceBudget: Number(r[12]||0),
        phone: (function(p){ p=String(p||''); return (p && /^\d{9}$/.test(p)) ? '0'+p : p; })(r[14]),
        idCard: String(r[15]||''), address: String(r[16]||''),
        profileUrl: String(r[17]||''), idCardUrl: String(r[18]||''),
        pin: String(r[13]||''),
        ssoEnabled: r[19] !== false && r[19] !== 'FALSE',
        allowances: (function(a){ try{ return JSON.parse(String(a||'[]')); }catch(e){ return []; } })(r[20])
      };
    }).filter(function(e) { return e.active && e.empId; });
    return jsonOut({ data: data });
  } catch(e) { return jsonOut({ data: [], error: e.toString() }); }
}

function saveHREmployee(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('HR_Employees');
    if (!sh) {
      sh = ss.insertSheet('HR_Employees');
      sh.appendRow(['empId','name','dept','position','type','salary','dailyRate','otRateWD','otRateSun','startDate','active','payCycle','advanceBudget','pin','phone','idCard','address','profileUrl','idCardUrl','ssoEnabled','allowances']);
    }
    var newRow = [
      data.empId, data.name, data.dept||'', data.position||'', data.type||'monthly',
      data.salary||0, data.dailyRate||0, data.otRateWD||100, data.otRateSun||200,
      data.startDate||'', data.active!==false,
      data.payCycle||'default', data.advanceBudget||0, data.pin||'', data.phone||'',
      data.idCard||'', data.address||'', data.profileUrl||'', data.idCardUrl||'',
      data.ssoEnabled !== false,
      JSON.stringify(data.allowances || [])
    ];
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.empId)) {
        // ถ้า PIN ว่าง → เก็บ PIN เดิมไว้ (ไม่ overwrite)
        if (!data.pin) newRow[13] = String(rows[i][13] || '');
        sh.getRange(i+1, 1, 1, newRow.length).setValues([newRow]);
        sh.getRange(i+1, 14).setNumberFormat('@'); // force text for pin
        sh.getRange(i+1, 15).setNumberFormat('@'); // force text for phone
        return jsonOut({ status:'ok' });
      }
    }
    sh.appendRow(newRow);
    sh.getRange(sh.getLastRow(), 14).setNumberFormat('@'); // force text for pin
    sh.getRange(sh.getLastRow(), 15).setNumberFormat('@'); // force text for phone
    return jsonOut({ status:'ok' });
  } catch(e) { return jsonOut({ status:'error', message: e.toString() }); }
}

function deleteHREmployee(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('HR_Employees');
    if (!sh) return jsonOut({ status:'ok' });
    var rows = sh.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) === String(data.empId)) {
        sh.deleteRow(i + 1);
        return jsonOut({ status:'ok' });
      }
    }
    return jsonOut({ status:'ok' });
  } catch(e) { return jsonOut({ status:'error', message: e.toString() }); }
}

function saveHRAttendance(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('HR_Attendance');
    if (!sh) {
      sh = ss.insertSheet('HR_Attendance');
      sh.appendRow(['empId','name','dept','date','dow','clockIn','lastScan','status','lateMin','otHours','otRate','otNote','period','importedAt']);
    }
    var now = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });
    var toInsert = data.rows || [];
    if (!toInsert.length) return jsonOut({ status:'ok', saved: 0, skipped: 0 });

    var checkOnly    = !!data.checkOnly;    // แค่ตรวจ ไม่เขียน
    var skipExisting = !!data.skipExisting; // ข้ามของเดิม (ไม่ทับ)

    // อ่านข้อมูลเดิมทั้งหมดครั้งเดียว สร้าง map empId+date → rowIndex (1-based)
    var existingRows = sh.getDataRange().getValues();
    var existMap = {};
    for (var i = 1; i < existingRows.length; i++) {
      var k = String(existingRows[i][0]) + '|' + _hrFmtDate(existingRows[i][3]);
      existMap[k] = i + 1;
    }

    var updates    = []; // { rowNum, values }
    var inserts    = []; // [values]
    var skippedList = []; // { empId, date, name }

    toInsert.forEach(function(r) {
      var newRow = [
        r.empId, r.empName||r.name||'', r.dept||'', r.date||'',
        r.dow||0, r.clockIn||'', r.lastScan||'', r.status||'present',
        r.lateMin||0, r.otHours||0, r.otRate||1, r.otNote||'', r.period||'', now
      ];
      var k = String(r.empId) + '|' + String(r.date);
      if (existMap[k]) {
        if (skipExisting || checkOnly) {
          // ข้าม — บันทึก info เพื่อแจ้งกลับ
          skippedList.push({ empId: r.empId, date: r.date, name: r.empName||r.name||'' });
        } else {
          updates.push({ rowNum: existMap[k], values: newRow });
        }
      } else {
        inserts.push(newRow);
        existMap[k] = -1;
      }
    });

    // checkOnly — ไม่เขียนอะไรเลย
    if (checkOnly) {
      return jsonOut({ status:'ok', saved: 0, skipped: skippedList.length, skippedList: skippedList });
    }

    // เขียนจริง
    updates.forEach(function(u) {
      sh.getRange(u.rowNum, 1, 1, u.values.length).setValues([u.values]);
    });
    if (inserts.length) {
      var lastRow = sh.getLastRow();
      sh.getRange(lastRow + 1, 1, inserts.length, inserts[0].length).setValues(inserts);
    }

    return jsonOut({ status:'ok', saved: inserts.length + updates.length, skipped: skippedList.length });
  } catch(e) { return jsonOut({ status:'error', message: e.toString() }); }
}

function _hrFmtDate(v) {
  // Sheets อาจแปลง '01/06/2026' เป็น Date object อัตโนมัติ — แปลงกลับ dd/MM/yyyy
  if (v instanceof Date) {
    var tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
    return Utilities.formatDate(v, tz, 'dd/MM/yyyy');
  }
  return String(v || '');
}

function _hrFmtTime(v) {
  // Sheets อาจแปลง '08:30' เป็น Date object — แปลงกลับ HH:mm
  if (!v && v !== 0) return '';
  if (v instanceof Date) {
    var tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
    return Utilities.formatDate(v, tz, 'HH:mm');
  }
  var s = String(v).trim();
  if (!s || s === '0' || s === '0.0') return '';
  // HH:mm:ss → HH:mm
  var m = s.match(/^(\d{1,2}):(\d{2})/);
  if (m) return ('0'+m[1]).slice(-2) + ':' + m[2];
  return s;
}

function deleteHRHoliday(data) {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var sh   = ss.getSheetByName('HR_Holidays');
    if (!sh) return jsonOut({ status: 'ok' });
    var target = String(data.date || '').trim();
    var rows   = sh.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {
      // Sheets อาจแปลง text date เป็น Date object → ต้อง format กลับให้ตรง
      var raw = rows[i][0];
      var cellDate;
      if (raw instanceof Date) {
        var dd = ('0' + raw.getDate()).slice(-2);
        var mm = ('0' + (raw.getMonth() + 1)).slice(-2);
        var yr = raw.getFullYear();
        if (yr < 2500) yr += 543; // CE → BE
        cellDate = dd + '/' + mm + '/' + yr;
      } else {
        cellDate = String(raw || '').trim();
      }
      if (cellDate === target) {
        sh.deleteRow(i + 1);
        return jsonOut({ status: 'ok', deleted: target });
      }
    }
    return jsonOut({ status: 'ok', deleted: null });
  } catch(e) { return jsonOut({ status: 'error', message: e.toString() }); }
}

function seedHRHolidays(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('HR_Holidays');
    if (!sh) {
      sh = ss.insertSheet('HR_Holidays');
      sh.appendRow(['date', 'name', 'type']);
    }
    // เก็บวันที่ที่มีอยู่แล้วเพื่อข้าม (handle Date object ที่ Sheets แปลงอัตโนมัติ)
    var rows = sh.getDataRange().getValues();
    var existing = {};
    rows.slice(1).forEach(function(r) {
      var raw = r[0];
      var d;
      if (raw instanceof Date) {
        var dd = ('0' + raw.getDate()).slice(-2);
        var mm = ('0' + (raw.getMonth() + 1)).slice(-2);
        var yr = raw.getFullYear();
        if (yr < 2500) yr += 543;
        d = dd + '/' + mm + '/' + yr;
      } else {
        d = String(raw || '').trim();
      }
      if (d) existing[d] = true;
    });
    var added = 0, skipped = 0;
    (data.holidays || []).forEach(function(h) {
      if (existing[h.date]) { skipped++; return; }
      sh.appendRow([h.date, h.name, h.type]);
      existing[h.date] = true;
      added++;
    });
    return jsonOut({ status: 'ok', added: added, skipped: skipped });
  } catch(e) { return jsonOut({ status: 'error', message: e.toString() }); }
}

function getHRHolidays() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('HR_Holidays');
    if (!sh) {
      sh = ss.insertSheet('HR_Holidays');
      sh.appendRow(['date', 'name', 'type']);
      // ── วันหยุดราชการไทย 2569 (2026) ──────────────────────────
      var holidays2569 = [
        ['01/01/2569', 'วันขึ้นปีใหม่',                              'นักขัตฤกษ์'],
        ['02/01/2569', 'วันหยุดชดเชยวันขึ้นปีใหม่',                  'ชดเชย'],
        ['20/02/2569', 'วันมาฆบูชา',                                 'นักขัตฤกษ์'],
        ['06/04/2569', 'วันจักรี',                                   'นักขัตฤกษ์'],
        ['13/04/2569', 'วันสงกรานต์',                                'นักขัตฤกษ์'],
        ['14/04/2569', 'วันสงกรานต์',                                'นักขัตฤกษ์'],
        ['15/04/2569', 'วันสงกรานต์',                                'นักขัตฤกษ์'],
        ['01/05/2569', 'วันแรงงานแห่งชาติ',                          'นักขัตฤกษ์'],
        ['31/05/2569', 'วันวิสาขบูชา',                               'นักขัตฤกษ์'],
        ['03/06/2569', 'วันเฉลิมพระชนมพรรษาสมเด็จพระราชินี',        'นักขัตฤกษ์'],
        ['28/07/2569', 'วันเฉลิมพระชนมพรรษา ร.10',                  'นักขัตฤกษ์'],
        ['12/08/2569', 'วันแม่แห่งชาติ',                             'นักขัตฤกษ์'],
        ['13/10/2569', 'วันคล้ายวันสวรรคต ร.9',                     'นักขัตฤกษ์'],
        ['23/10/2569', 'วันปิยมหาราช',                               'นักขัตฤกษ์'],
        ['05/12/2569', 'วันคล้ายวันพระบรมราชสมภพ ร.9 (วันพ่อ)',     'นักขัตฤกษ์'],
        ['07/12/2569', 'วันหยุดชดเชยวันพ่อ',                         'ชดเชย'],
        ['10/12/2569', 'วันรัฐธรรมนูญ',                              'นักขัตฤกษ์'],
        ['31/12/2569', 'วันสิ้นปี',                                  'นักขัตฤกษ์'],
      ];
      holidays2569.forEach(function(row) { sh.appendRow(row); });
    }
    var rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return jsonOut({ data: [] });
    var data = rows.slice(1).map(function(r) {
      var raw = r[0];
      var dateStr;
      if (raw instanceof Date) {
        var dd  = ('0' + raw.getDate()).slice(-2);
        var mm  = ('0' + (raw.getMonth() + 1)).slice(-2);
        var yr  = raw.getFullYear();
        if (yr < 2500) yr += 543; // CE → BE (Thai locale: 2026 → 2569)
        dateStr = dd + '/' + mm + '/' + yr;
      } else {
        dateStr = String(raw || '');
      }
      return { date: dateStr, name: String(r[1]||''), type: String(r[2]||'') };
    }).filter(function(r) { return r.date; });
    return jsonOut({ data: data });
  } catch(e) { return jsonOut({ data: [], error: e.toString() }); }
}

function getHRAttendance(params) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('HR_Attendance');
    if (!sh) return jsonOut({ data: [] });
    var rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return jsonOut({ data: [] });
    var filterMonth = String(params.month||'');
    var filterEmpId = String(params.empId||'');
    var data = rows.slice(1).map(function(r) {
      return {
        empId: String(r[0]||''), name: String(r[1]||''), empName: String(r[1]||''),
        dept: String(r[2]||''), date: _hrFmtDate(r[3]),
        dow: Number(r[4]||0), clockIn: _hrFmtTime(r[5]), lastScan: _hrFmtTime(r[6]),
        status: String(r[7]||'present'), lateMin: Number(r[8]||0),
        otHours: Number(r[9]||0), otRate: Number(r[10]||1), period: String(r[11]||'')
      };
    }).filter(function(r) {
      if (!r.empId) return false;
      if (filterEmpId && r.empId !== filterEmpId) return false;
      if (filterMonth) {
        var p = r.date.split('/');
        var key = p.length >= 3 ? p[2] + '-' + (p[1].length < 2 ? '0'+p[1] : p[1]) : '';
        if (key !== filterMonth) return false;
      }
      return true;
    });
    return jsonOut({ data: data });
  } catch(e) { return jsonOut({ data: [], error: e.toString() }); }
}

// ══════════════════════════════════════════════════════════════
// LOAN REQUESTS — เงินเบิกล่วงหน้า / เงินกู้บริษัท
// ══════════════════════════════════════════════════════════════
var LOAN_COLS = ['requestId','empId','empName','dept','type','amount','reason',
  'requestDate','status','approvedBy','approvedDate','repayPeriods',
  'installmentAmt','deductMonth','notes',
  'transferDate','transferAmount','transferNote','closedDate','slipImageUrl',
  'deducted','deductedDate'];

var SLIP_FOLDER_ID = '178PJZtwzzIalwx124hd6H6frDpLCNW-N';

// ── PTS folder helper ──
function _getPtsSubFolder(subName) {
  var root = DriveApp.getRootFolder();
  var ptsIter = root.getFoldersByName('PTS-Management System');
  var pts = ptsIter.hasNext() ? ptsIter.next() : root.createFolder('PTS-Management System');
  var subIter = pts.getFoldersByName(subName);
  return subIter.hasNext() ? subIter.next() : pts.createFolder(subName);
}

function uploadEmployeeFile(data) {
  try {
    var folder = _getPtsSubFolder('HR-Employees');
    var bytes = Utilities.base64Decode(data.base64);
    var blob  = Utilities.newBlob(bytes, data.mimeType || 'image/jpeg', data.filename || 'file.jpg');
    var file  = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    // ใช้ thumbnail URL — โหลดได้ใน <img> โดยตรง
    var fid = file.getId();
    var url = 'https://drive.google.com/thumbnail?id=' + fid + '&sz=w400-h400';
    var viewUrl = 'https://drive.google.com/uc?id=' + fid + '&export=view';
    return jsonOut({ status: 'ok', url: url, viewUrl: viewUrl, fileId: fid });
  } catch(e) { return jsonOut({ status: 'error', message: e.toString() }); }
}

function _loanSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('LoanRequests');
  if (!sh) {
    sh = ss.insertSheet('LoanRequests');
    sh.appendRow(LOAN_COLS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function _genLoanId() {
  var now = new Date();
  var d = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyyMMdd');
  var sh = _loanSheet();
  var last = sh.getLastRow();
  var seq = String(last).padStart(3,'0');
  return 'REQ-' + d + '-' + seq;
}

function saveLoanRequest(data) {
  try {
    var sh = _loanSheet();
    var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');

    // ตรวจ budget เงินเบิก
    if (data.type === 'advance') {
      var budget = parseFloat(data.advanceBudget) || 0;
      if (budget > 0) {
        var rows = sh.getDataRange().getValues();
        var mon = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM');
        var used = 0;
        for (var i = 1; i < rows.length; i++) {
          if (String(rows[i][1]) !== String(data.empId)) continue;
          if (String(rows[i][4]) !== 'advance') continue;
          if (rows[i][8] === 'rejected') continue;
          // requestDate format dd/MM/yyyy HH:mm -> extract yyyy-MM
          var rd = String(rows[i][7] || '');
          var rdKey = rd.length >= 10 ? rd.slice(6,10) + '-' + rd.slice(3,5) : '';
          if (rdKey !== mon) continue;
          used += parseFloat(rows[i][5]) || 0;
        }
        if (used + parseFloat(data.amount) > budget) {
          return jsonOut({ status: 'budget_exceeded',
            message: 'เกินวงเงิน — ใช้ไปแล้ว ' + used.toLocaleString() + ' บาท / วงเงิน ' + budget.toLocaleString() + ' บาท' });
        }
      }
    }

    var id = _genLoanId();
    var installAmt = data.type === 'loan'
      ? Math.ceil(parseFloat(data.amount) / (parseInt(data.repayPeriods) || 1))
      : parseFloat(data.amount);

    sh.appendRow([
      id, data.empId, data.empName || '', data.dept || '',
      data.type, parseFloat(data.amount) || 0, data.reason || '',
      now, 'pending', '', '', parseInt(data.repayPeriods) || 1,
      installAmt, '', data.notes || ''
    ]);
    return jsonOut({ status: 'ok', requestId: id });
  } catch(e) { return jsonOut({ status: 'error', message: e.toString() }); }
}

function getLoanRequests(params) {
  try {
    var sh = _loanSheet();
    var rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return jsonOut({ data: [] });
    var filterEmp    = String(params.empId  || '');
    var filterStatus = String(params.status || '');
    var data = rows.slice(1).map(function(r) {
      return {
        requestId: String(r[0]||''), empId: String(r[1]||''), empName: String(r[2]||''),
        dept: String(r[3]||''), type: String(r[4]||''), amount: Number(r[5]||0),
        reason: String(r[6]||''), requestDate: r[7] instanceof Date ? Utilities.formatDate(r[7],'Asia/Bangkok','dd/MM/yyyy HH:mm') : String(r[7]||''), status: String(r[8]||'pending'),
        approvedBy: String(r[9]||''), approvedDate: String(r[10]||''),
        repayPeriods: Number(r[11]||1), installmentAmt: Number(r[12]||0),
        deductMonth: r[13] instanceof Date ? Utilities.formatDate(r[13],'Asia/Bangkok','yyyy-MM') : String(r[13]||''),
        notes: String(r[14]||''),
        transferDate: String(r[15]||''), transferAmount: Number(r[16]||0),
        transferNote: String(r[17]||''), closedDate: String(r[18]||''),
        slipImageUrl: String(r[19]||''),
        deducted: r[20] === true || String(r[20]||'') === 'TRUE',
        deductedDate: String(r[21]||'')
      };
    }).filter(function(r) {
      if (!r.requestId) return false;
      if (filterEmp    && r.empId  !== filterEmp)    return false;
      if (filterStatus && r.status !== filterStatus) return false;
      return true;
    });
    return jsonOut({ data: data });
  } catch(e) { return jsonOut({ data: [], error: e.toString() }); }
}

function approveLoan(data) {
  try {
    var sh   = _loanSheet();
    var rows = sh.getDataRange().getValues();
    var now  = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== String(data.requestId)) continue;
      sh.getRange(i+1, 9).setValue('approved');
      sh.getRange(i+1, 10).setValue(data.approvedBy || 'ผู้จัดการ');
      sh.getRange(i+1, 11).setValue(now);
      if (data.deductMonth) sh.getRange(i+1, 14).setValue(data.deductMonth);
      if (data.notes)       sh.getRange(i+1, 15).setValue(data.notes);
      return jsonOut({ status: 'ok' });
    }
    return jsonOut({ status: 'error', message: 'ไม่พบ requestId' });
  } catch(e) { return jsonOut({ status: 'error', message: e.toString() }); }
}

function rejectLoan(data) {
  try {
    var sh   = _loanSheet();
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== String(data.requestId)) continue;
      sh.getRange(i+1, 9).setValue('rejected');
      if (data.notes) sh.getRange(i+1, 15).setValue(data.notes);
      return jsonOut({ status: 'ok' });
    }
    return jsonOut({ status: 'error', message: 'ไม่พบ requestId' });
  } catch(e) { return jsonOut({ status: 'error', message: e.toString() }); }
}

function uploadSlipImage(data) {
  try {
    var folder = DriveApp.getFolderById(SLIP_FOLDER_ID);
    var bytes  = Utilities.base64Decode(data.base64);
    var blob   = Utilities.newBlob(bytes, data.mimeType || 'image/jpeg', data.filename || 'slip.jpg');
    var file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/uc?id=' + file.getId() + '&export=view';
    return jsonOut({ status: 'ok', url: url, fileId: file.getId() });
  } catch(e) { return jsonOut({ status: 'error', message: e.toString() }); }
}

function closeLoan(data) {
  try {
    var sh   = _loanSheet();
    var rows = sh.getDataRange().getValues();
    var now  = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== String(data.requestId)) continue;
      if (rows[i][8] !== 'approved') return jsonOut({ status: 'error', message: 'สถานะต้องเป็น approved ก่อนปิดรายการ' });
      sh.getRange(i+1, 9).setValue('closed');
      if (data.deductMonth)     sh.getRange(i+1, 14).setValue(data.deductMonth);
      if (data.transferDate)    sh.getRange(i+1, 16).setValue(data.transferDate);
      if (data.transferAmount)  sh.getRange(i+1, 17).setValue(parseFloat(data.transferAmount) || 0);
      if (data.transferNote)    sh.getRange(i+1, 18).setValue(data.transferNote);
      sh.getRange(i+1, 19).setValue(now);
      if (data.slipImageUrl)    sh.getRange(i+1, 20).setValue(data.slipImageUrl);
      return jsonOut({ status: 'ok' });
    }
    return jsonOut({ status: 'error', message: 'ไม่พบ requestId' });
  } catch(e) { return jsonOut({ status: 'error', message: e.toString() }); }
}

// ═══════════════════════════════════════════════════════�

function deleteLoanRequest(data) {
  try {
    var sh   = _loanSheet();
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== String(data.requestId)) continue;
      sh.deleteRow(i + 1);
      return jsonOut({ status: 'ok' });
    }
    return jsonOut({ status: 'error', message: 'ไม่พบ requestId' });
  } catch(e) { return jsonOut({ status: 'error', message: e.toString() }); }
}

function markAdvanceDeducted(data) {
  // data.requestIds = array of requestId ที่ถูกหักในสลิปงวดนี้
  try {
    var sh   = _loanSheet();
    var rows = sh.getDataRange().getValues();
    var ids  = (data.requestIds || []).map(String);
    var now  = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
    var count = 0;
    for (var i = 1; i < rows.length; i++) {
      if (ids.indexOf(String(rows[i][0])) === -1) continue;
      if (rows[i][8] !== 'closed') continue; // เฉพาะ closed เท่านั้น
      sh.getRange(i+1, 21).setValue(true);   // col 21 = deducted
      sh.getRange(i+1, 22).setValue(now);    // col 22 = deductedDate
      count++;
    }
    return jsonOut({ status: 'ok', marked: count });
  } catch(e) { return jsonOut({ status: 'error', message: e.toString() }); }
}

// ═══════════════════════════════════════════════════════
// HR LOAN CONTRACTS — ตารางสัญญาเงินกู้ (HR_Loans)
// ═══════════════════════════════════════════════════════

var HR_LOAN_COLS = ['loanId','empId','empName','type','originalAmount','outstanding',
  'installmentAmt','payDays','startDate','reason','status','closedDate','createdDate','notes'];

var HR_LOAN_PAY_COLS = ['paymentId','loanId','empId','amount','payDate','periodLabel',
  'outstandingAfter','note','recordedBy','recordedDate'];

function _hrLoanContractSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('HR_Loans');
  if (!sh) {
    sh = ss.insertSheet('HR_Loans');
    sh.appendRow(HR_LOAN_COLS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function _hrLoanPaymentSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('HR_LoanPayments');
  if (!sh) {
    sh = ss.insertSheet('HR_LoanPayments');
    sh.appendRow(HR_LOAN_PAY_COLS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function _genHRLoanId() {
  var now = new Date();
  var d = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyyMMdd');
  var sh = _hrLoanContractSheet();
  var seq = String(Math.max(1, sh.getLastRow())).padStart(3,'0');
  return 'LN-' + d + '-' + seq;
}

function _genHRPaymentId() {
  var now = new Date();
  var d = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyyMMdd');
  var sh = _hrLoanPaymentSheet();
  var seq = String(Math.max(1, sh.getLastRow())).padStart(3,'0');
  return 'PAY-' + d + '-' + seq;
}

function saveHRLoanContract(data) {
  try {
    var sh = _hrLoanContractSheet();
    var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
    var rows = sh.getDataRange().getValues();
    if (data.loanId) {
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) !== String(data.loanId)) continue;
        if (data.empName      !== undefined) sh.getRange(i+1,3).setValue(data.empName);
        if (data.type         !== undefined) sh.getRange(i+1,4).setValue(data.type);
        if (data.outstanding  !== undefined) sh.getRange(i+1,6).setValue(parseFloat(data.outstanding)||0);
        if (data.installmentAmt!== undefined)sh.getRange(i+1,7).setValue(parseFloat(data.installmentAmt)||0);
        if (data.payDays      !== undefined) sh.getRange(i+1,8).setValue(data.payDays);
        if (data.reason       !== undefined) sh.getRange(i+1,10).setValue(data.reason);
        if (data.status       !== undefined) sh.getRange(i+1,11).setValue(data.status);
        if (data.status === 'closed')        sh.getRange(i+1,12).setValue(now);
        if (data.notes        !== undefined) sh.getRange(i+1,14).setValue(data.notes);
        return jsonOut({ status:'ok', loanId: data.loanId });
      }
    }
    var loanId = _genHRLoanId();
    sh.appendRow([
      loanId,
      data.empId || '',
      data.empName || '',
      data.type || 'loan',
      parseFloat(data.originalAmount || data.outstanding || 0),
      parseFloat(data.outstanding || 0),
      parseFloat(data.installmentAmt || 0),
      data.payDays || '1,16',
      data.startDate || now.slice(0,10),
      data.reason || '',
      data.status || 'active',
      '',
      now,
      data.notes || ''
    ]);
    return jsonOut({ status:'ok', loanId: loanId });
  } catch(e) { return jsonOut({ status:'error', message: e.toString() }); }
}

function getHRLoanContracts(params) {
  try {
    var sh = _hrLoanContractSheet();
    var rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return jsonOut({ data:[] });
    var empId = params && params.empId ? String(params.empId) : null;
    var data = rows.slice(1).map(function(r) {
      return {
        loanId: String(r[0]||''), empId: String(r[1]||''), empName: String(r[2]||''),
        type: String(r[3]||'loan'),
        originalAmount: Number(r[4]||0), outstanding: Number(r[5]||0),
        installmentAmt: Number(r[6]||0), payDays: String(r[7]||'1,16'),
        startDate: String(r[8]||''), reason: String(r[9]||''),
        status: String(r[10]||'active'), closedDate: String(r[11]||''),
        createdDate: String(r[12]||''), notes: String(r[13]||'')
      };
    }).filter(function(l) {
      if (!l.loanId) return false;
      if (empId) return l.empId === empId;
      return true;
    });
    return jsonOut({ data: data });
  } catch(e) { return jsonOut({ data:[], error: e.toString() }); }
}

// payrollDeductLoans — หักยอดเงินกู้/เบิกล่วงหน้าพร้อมกันหลายรายการในงวดเดียว
// data.contracts = [{loanId, amount}]
// data.requests  = [{requestId, amount}] (เบิกล่วงหน้า — ไม่มี outstanding ต้องอัป)
function payrollDeductLoans(data) {
  try {
    var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
    var results = { contracts: [], requests: [], errors: [] };

    // ── หักสัญญาเงินกู้ ──
    var contracts = data.contracts || [];
    for (var ci = 0; ci < contracts.length; ci++) {
      var c = contracts[ci];
      try {
        var r = recordHRLoanPayment({
          loanId: c.loanId,
          amount: c.amount,
          periodLabel: data.periodLabel || '',
          note: 'หักจากเงินเดือน',
          recordedBy: data.recordedBy || 'payroll'
        });
        var rObj = JSON.parse(r.getContent());
        results.contracts.push({ loanId: c.loanId, status: rObj.status, paymentId: rObj.paymentId });
        if (rObj.status !== 'ok') results.errors.push('loanId ' + c.loanId + ': ' + rObj.message);
      } catch(ce) { results.errors.push('loanId ' + c.loanId + ': ' + ce.toString()); }
    }

    // ── เบิกล่วงหน้า (request) — ไม่ต้องอัป outstanding แค่ log ไว้ใน HR_LoanPayments ─
    // (การเปลี่ยนสถานะ request ทำผ่าน markAdvanceDeducted แยกต่างหาก)

    return jsonOut({ status: 'ok', results: results });
  } catch(e) { return jsonOut({ status: 'error', message: e.toString() }); }
}

function recordHRLoanPayment(data) {
  try {
    var pSh = _hrLoanPaymentSheet();
    var lSh = _hrLoanContractSheet();
    var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
    var lRows = lSh.getDataRange().getValues();
    var lIdx = -1;
    for (var i = 1; i < lRows.length; i++) {
      if (String(lRows[i][0]) === String(data.loanId)) { lIdx = i; break; }
    }
    if (lIdx < 0) return jsonOut({ status:'error', message:'ไม่พบ loanId' });
    var outstanding = parseFloat(lRows[lIdx][5] || 0);
    var amt = parseFloat(data.amount || 0);
    if (amt <= 0) return jsonOut({ status:'error', message:'ยอดชำระต้องมากกว่า 0' });
    var newOutstanding = Math.round(Math.max(0, outstanding - amt) * 100) / 100;
    lSh.getRange(lIdx+1, 6).setValue(newOutstanding);
    if (newOutstanding === 0) {
      lSh.getRange(lIdx+1, 11).setValue('closed');
      lSh.getRange(lIdx+1, 12).setValue(now);
    }
    var payId = _genHRPaymentId();
    pSh.appendRow([
      payId, data.loanId, data.empId || lRows[lIdx][1],
      amt, data.payDate || now.slice(0,10),
      data.periodLabel || '',
      newOutstanding,
      data.note || '',
      data.recordedBy || '',
      now
    ]);
    return jsonOut({ status:'ok', paymentId: payId, outstandingAfter: newOutstanding, closed: newOutstanding === 0 });
  } catch(e) { return jsonOut({ status:'error', message: e.toString() }); }
}

function getHRLoanPayments(params) {
  try {
    var sh = _hrLoanPaymentSheet();
    var rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return jsonOut({ data:[] });
    var loanId = params && params.loanId ? String(params.loanId) : null;
    var empId  = params && params.empId  ? String(params.empId)  : null;
    var data = rows.slice(1).map(function(r) {
      return {
        paymentId: String(r[0]||''), loanId: String(r[1]||''), empId: String(r[2]||''),
        amount: Number(r[3]||0), payDate: (r[4] instanceof Date) ? Utilities.formatDate(r[4], 'Asia/Bangkok', 'dd/MM/yyyy') : String(r[4]||'').trim(),
        periodLabel: String(r[5]||''), outstandingAfter: Number(r[6]||0),
        note: String(r[7]||''), recordedBy: String(r[8]||''), recordedDate: String(r[9]||'')
      };
    }).filter(function(p) {
      if (!p.paymentId) return false;
      if (loanId && p.loanId !== loanId) return false;
      if (empId  && p.empId  !== empId)  return false;
      return true;
    });
    return jsonOut({ data: data });
  } catch(e) { return jsonOut({ data:[], error: e.toString() }); }
}

// ══════════════════════════════════════════════════════════════════
// HR_SalaryPayments — Phase 2
// Schema: payId | empId | empName | month | period | amount
//         transferDate | note | slipUrl | recordedBy | deductItemsJson
//         baseSalary | dailyRate | recordedAt
// baseSalary + dailyRate = snapshot ณ เวลาจ่าย (ไม่กระทบ _hrCalcPayslip)
// ══════════════════════════════════════════════════════════════════

function _salaryPaymentSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('HR_SalaryPayments');
  if (!sh) {
    sh = ss.insertSheet('HR_SalaryPayments');
    sh.appendRow([
      'payId','empId','empName','month','period','amount',
      'transferDate','note','slipUrl','recordedBy','deductItemsJson',
      'baseSalary','dailyRate','recordedAt'
    ]);
    sh.setFrozenRows(1);
  } else {
    // เติม header ที่ขาด (รองรับชีตที่มีอยู่ก่อนด้วย schema เก่า)
    var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var needed = [
      [11,'deductItemsJson'],[12,'baseSalary'],[13,'dailyRate'],[14,'recordedAt']
    ];
    needed.forEach(function(pair) {
      var col = pair[0], name = pair[1];
      if (!hdr[col - 1] || String(hdr[col - 1]).trim() === '') {
        sh.getRange(1, col).setValue(name);
      }
    });
  }
  return sh;
}

function _genSalaryPayId() {
  return 'SP-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMddHHmmss')
    + '-' + Math.floor(Math.random()*1000);
}

// POST action: recordSalaryPayment
// body: { empId, empName, month, period, amount, transferDate, note, slipUrl,
//         recordedBy, deductItems[], baseSalary, dailyRate, allowancesJson }
function recordSalaryPayment(data) {
  try {
    if (!data.empId || !data.month || !data.amount) {
      return jsonOut({ status: 'error', message: 'ข้อมูลไม่ครบ (empId, month, amount)' });
    }

    // ── Snapshot baseSalary/dailyRate ──
    // รับจาก client (hr.js จะส่งมา) หรือ fallback ดึงจาก HR_Employees
    var snapshotSalary   = Number(data.baseSalary)  || 0;
    var snapshotDaily    = Number(data.dailyRate)    || 0;
    // fallback เฉพาะเมื่อทั้ง salary และ dailyRate เป็น 0 (ไม่ได้ส่งมาจาก client)
    // HR_Employees schema: [0]empId [1]name [2]dept [3]position [4]type [5]salary [6]dailyRate
    if (!snapshotSalary && !snapshotDaily) {
      var empSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HR_Employees');
      if (empSh) {
        var empRows = empSh.getDataRange().getValues();
        for (var ei = 1; ei < empRows.length; ei++) {
          if (String(empRows[ei][0]).trim() === String(data.empId).trim()) {
            snapshotSalary = Number(empRows[ei][5]) || 0;  // col F = salary
            snapshotDaily  = Number(empRows[ei][6]) || 0;  // col G = dailyRate
            break;
          }
        }
      }
    }

    var sh  = _salaryPaymentSheet();
    var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
    var payId = _genSalaryPayId();
    var deductJson = '';
    try { deductJson = JSON.stringify(data.deductItems || []); } catch(e2) {}
    var allowJson = '';
    try { allowJson = data.allowancesJson ? String(data.allowancesJson) : JSON.stringify(data.allowances || []); } catch(e3) {}

    sh.appendRow([
      payId,
      String(data.empId).trim(),
      String(data.empName || '').trim(),
      String(data.month   || '').trim(),
      String(data.period  || 'all').trim(),
      Number(data.amount) || 0,
      String(data.transferDate || '').trim(),
      String(data.note    || '').trim(),
      String(data.slipUrl || '').trim(),
      String(data.recordedBy || 'admin').trim(),
      deductJson,
      snapshotSalary,
      snapshotDaily,
      now,
      allowJson
    ]);
    // Force month (col D=4) และ transferDate (col G=7) เป็น plain text
    var newRow = sh.getLastRow();
    var monthCell = sh.getRange(newRow, 4);
    monthCell.setNumberFormat('@STRING@');
    monthCell.setValue(String(data.month || '').trim());
    var xfCell = sh.getRange(newRow, 7);
    xfCell.setNumberFormat('@STRING@');
    xfCell.setValue(String(data.transferDate || '').trim());
    return jsonOut({ status: 'ok', payId: payId });
  } catch(e) { return jsonOut({ status: 'error', message: e.message }); }
}

// GET action: getSalaryPayments?month=yyyy-MM&empId=xxx
function getSalaryPayments(params) {
  try {
    var sh = _salaryPaymentSheet();
    var rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return jsonOut({ status:'ok', data: [] });
    var fMonth = (params && params.month) ? String(params.month).trim() : '';
    var fEmpId = (params && params.empId) ? String(params.empId).trim() : '';
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      // month อาจถูก Sheets auto-parse เป็น Date object → format กลับเป็น yyyy-MM
      var monthRaw = r[3];
      var rMonth = (monthRaw instanceof Date)
        ? Utilities.formatDate(monthRaw, 'Asia/Bangkok', 'yyyy-MM')
        : String(monthRaw||'').trim();
      var rEmpId = String(r[1]||'').trim();
      if (fMonth && rMonth !== fMonth) continue;
      if (fEmpId && rEmpId !== fEmpId) continue;
      var deductItems = [];
      try { deductItems = JSON.parse(String(r[10]||'[]')); } catch(e2) {}
      var allowances = [];
      try { allowances = JSON.parse(String(r[14]||'[]')); } catch(e4) {}
      result.push({
        payId:        String(r[0]||'').trim(),
        empId:        rEmpId,
        empName:      String(r[2]||'').trim(),
        month:        rMonth,
        period:       String(r[4]||'all').trim(),
        amount:       Number(r[5])||0,
        transferDate: (r[6] instanceof Date)
          ? Utilities.formatDate(r[6], 'Asia/Bangkok', 'dd/MM/yyyy')
          : String(r[6]||'').trim(),
        note:         String(r[7]||'').trim(),
        slipUrl:      String(r[8]||'').trim(),
        recordedBy:   String(r[9]||'').trim(),
        deductItems:  deductItems,
        baseSalary:   Number(r[11])||0,
        dailyRate:    Number(r[12])||0,
        recordedAt:   String(r[13]||'').trim(),
        allowances:   allowances
      });
    }
    return jsonOut({ status:'ok', data: result });
  } catch(e) { return jsonOut({ status:'ok', data:[], error: e.message }); }
}

// POST action: deleteSalaryPayment
// body: { payId }
function deleteSalaryPayment(data) {
  try {
    var sh = _salaryPaymentSheet();
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.payId).trim()) {
        sh.deleteRow(i + 1);
        return jsonOut({ status: 'ok' });
      }
    }
    return jsonOut({ status: 'error', message: 'ไม่พบ payId: ' + data.payId });
  } catch(e) { return jsonOut({ status: 'error', message: e.message }); }
}

