function doGet(e) {
  var p = (e && e.parameter) || {};
  var app = String(p.app || p.view || p.ui || '').toLowerCase();
  var op = String(p.op || p.action || p.a || '').toLowerCase();

  if (app === 'receipt' || app === 'receipts' || app === 'ipad') {
    return ipad_receiptDoGet(e);
  }

  if (op) {
    return askControllerDoGet_(e);
  }

  return sw_taskQueueDoGet_(e);
}

function sw_taskQueueDoGet_(e) {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Sales Appointment Workflow')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
