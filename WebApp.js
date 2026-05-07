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

function sw_taskQueueDoPost_(e) {
  var started = new Date().getTime();
  try {
    var body = swWebAppJsonBody_(e);
    var action = swTrim_((body && body.action) || '');
    if (action === 'sw_login') {
      var loginResult = sw_login(
        body.email || '',
        body.password || '',
        body.options || {}
      );
      loginResult._swWebApp = {
        action: action,
        transport: 'webAppPost',
        serverElapsedMs: new Date().getTime() - started
      };
      return swWebAppJson_(loginResult);
    }
    if (action === 'sw_getBootstrap') {
      var bootstrapResult = sw_getBootstrap(swTrim_(body.token || ''));
      bootstrapResult._swWebApp = {
        action: action,
        transport: 'webAppPost',
        serverElapsedMs: new Date().getTime() - started
      };
      return swWebAppJson_(bootstrapResult);
    }
    return swWebAppErr_('BAD_ACTION', 'Unknown workflow POST action: ' + action);
  } catch (err) {
    return swWebAppErr_('POST_ERR', err && (err.message || err.stack) ? (err.message || err.stack) : String(err));
  }
}

function swWebAppJsonBody_(e) {
  if (!e || !e.postData || !/json/i.test(String(e.postData.type || ''))) return {};
  try {
    return JSON.parse(e.postData.contents || '{}');
  } catch (_) {}
  return {};
}

function swWebAppJson_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj || {}))
    .setMimeType(ContentService.MimeType.JSON);
}

function swWebAppErr_(code, message) {
  return swWebAppJson_({
    ok: false,
    error: code || 'ERROR',
    message: String(message || '')
  });
}
