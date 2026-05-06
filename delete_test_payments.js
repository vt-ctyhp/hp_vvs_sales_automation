// ============================================================
// delete_test_payments.gs
// Xem trước + xóa rows test của thanhvu@ctyhp.vn
// khỏi 400 Payments Ledger
// ============================================================

var DELETE_SUBMITTED_BY = 'thanhvu@ctyhp.vn'; // ← đổi nếu cần

// ── Bước 1: Xem trước — chạy cái này trước ──────────────────
// Hiện danh sách rows sẽ bị xóa, CHƯA xóa gì cả.
function previewTestPayments() {
  var ui   = SpreadsheetApp.getUi();
  var info = _getTestRows_();

  if (!info.rows.length) {
    ui.alert('✅ Không tìm thấy row nào của "' + DELETE_SUBMITTED_BY + '" trong Payments sheet.\n\nKhông có gì để xóa.');
    return;
  }

  var preview = info.rows.slice(0, 20).map(function(r) {
    return 'Row ' + r.sheetRow + ' | ' + r.paymentId + ' | ' + r.date + ' | $' + r.gross + ' | ' + r.docType;
  }).join('\n');

  var more = info.rows.length > 20 ? '\n... và ' + (info.rows.length - 20) + ' rows nữa' : '';

  ui.alert(
    '⚠️ TÌM THẤY ' + info.rows.length + ' ROWS CỦA "' + DELETE_SUBMITTED_BY + '"\n\n' +
    preview + more + '\n\n' +
    'Chạy deleteTestPayments() để XÓA VĨNH VIỄN.\n' +
    'Hoặc không làm gì nếu muốn giữ lại.'
  );
}

// ── Bước 2: Xóa thật — chạy SAU KHI đã xem preview ─────────
function deleteTestPayments() {
  var ui   = SpreadsheetApp.getUi();
  var info = _getTestRows_();

  if (!info.rows.length) {
    ui.alert('Không tìm thấy row nào. Không có gì để xóa.');
    return;
  }

  var confirm = ui.alert(
    '⚠️ XÁC NHẬN XÓA VĨNH VIỄN',
    'Sắp xóa ' + info.rows.length + ' rows của "' + DELETE_SUBMITTED_BY + '".\n\n' +
    'Hành động này KHÔNG THỂ UNDO.\n\n' +
    'Bạn có chắc chắn muốn xóa không?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    ui.alert('Đã huỷ. Không có gì bị xóa.');
    return;
  }

  // Xóa từ dưới lên để row index không bị lệch
  var sh = info.sheet;
  var rowNums = info.rows.map(function(r) { return r.sheetRow; }).sort(function(a,b){ return b-a; });

  rowNums.forEach(function(r) { sh.deleteRow(r); });
  try { if (typeof swInvalidatePaymentReadModelsAfterWrite_ === 'function') swInvalidatePaymentReadModelsAfterWrite_(null, 'Test payments deleted'); } catch (_) {}

  ui.alert('✅ Đã xóa ' + rowNums.length + ' rows của "' + DELETE_SUBMITTED_BY + '" khỏi Payments sheet.');
}

// ── Helper: tìm rows cần xóa ─────────────────────────────────
function _getTestRows_() {
  var id  = (PropertiesService.getScriptProperties().getProperty('LEDGER_FILE_ID') || '').trim();
  if (!id) throw new Error('LEDGER_FILE_ID chưa set. Chạy setLedgerFileId() trước.');

  var ledger = SpreadsheetApp.openById(id);
  var sheets = ledger.getSheets();
  var sh     = sheets.filter(function(s){ return /payment/i.test(s.getName()); })[0] || sheets[0];

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { sheet: sh, rows: [] };

  // SubmittedBy = col W = index 22 (0-based)
  var SUBMITTED_BY_COL = 23; // 1-based = col W

  var data = sh.getRange(2, 1, lastRow - 1, 47).getValues();
  var target = DELETE_SUBMITTED_BY.trim().toLowerCase();

  var rows = [];
  data.forEach(function(row, i) {
    var submittedBy = String(row[22] || '').trim().toLowerCase(); // col W index 22
    if (submittedBy === target) {
      rows.push({
        sheetRow:  i + 2, // 1-based, +1 for header
        paymentId: String(row[0]  || ''), // col A PAYMENT_ID
        date:      String(row[7]  || ''), // col H PaymentDateTime
        gross:     String(row[11] || ''), // col L AmountGross
        docType:   String(row[5]  || '')  // col F DocType
      });
    }
  });

  return { sheet: sh, rows: rows };
}
