function doGet(e) {
  var p = (e && e.parameter) || {};
  var app = String(p.app || p.view || p.ui || '').toLowerCase();
  var hasPaymentLaunchToken = !!(p.launch || p.paymentLaunch || p.token);
  var op = String(p.op || p.action || p.a || '').toLowerCase();

  if (hasPaymentLaunchToken || app === 'receipt' || app === 'receipts' || app === 'payment' || app === 'payments' || app === 'ipad') {
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
