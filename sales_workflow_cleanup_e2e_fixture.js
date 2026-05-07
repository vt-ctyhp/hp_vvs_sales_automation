/**
 * Admin-only helpers for cleanup workflow end-to-end testing.
 *
 * These helpers do not change cleanup workflow behavior. They only expose
 * spreadsheet/service info and append a clearly-labeled synthetic stale row in
 * 00_Master Appointments so the existing generator can produce real cleanup
 * tasks for browser automation.
 */

function sw_adminGetWorkflowSpreadsheetInfo(authToken) {
  if (typeof authToken === 'object') authToken = '';
  var ss = swSpreadsheet_();
  var user = swAuthUserForApi_(ss, authToken || '');
  if (!user || !user.isAdmin) throw new Error('Admin access required.');
  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    spreadsheetUrl: ss.getUrl(),
    serviceUrl: ScriptApp.getService().getUrl(),
    webAppUrl: ScriptApp.getService().getUrl(),
    sheets: SW_SHEETS
  };
}

function sw_adminCreateCleanupE2eFixture(authToken, options) {
  if (typeof authToken === 'object' && options == null) {
    options = authToken;
    authToken = '';
  }
  options = options || {};

  sw_setupSalesWorkflow();
  var ss = swSpreadsheet_();
  var user = swAuthUserForApi_(ss, authToken || '');
  if (!user || !user.isAdmin) throw new Error('Admin access required.');

  var master = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!master) throw new Error('Missing sheet: ' + SW_SHEETS.MASTER);

  var lastCol = Math.max(master.getLastColumn(), 1);
  var headers = master.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var idx = swAppointmentColumnIndex_(headers);
  var row = [];
  for (var i = 0; i < lastCol; i++) row[i] = '';

  var now = new Date();
  var staleDays = Math.max(31, Number(options.staleDays || 62) || 62);
  var touchAt = new Date(now.getTime() - (staleDays * 86400000));
  var tz = swTimezone_();
  var stamp = Utilities.formatDate(now, tz, 'yyyyMMdd_HHmmss');
  var suffix = Utilities.getUuid().slice(0, 8).toUpperCase();
  var fixtureId = 'E2E_CLEANUP_' + stamp + '_' + suffix;
  var rootApptId = fixtureId;
  var apptId = fixtureId + '_APPT';
  var customerName = swTrim_(options.customerName || ('Test Customer Cleanup E2E ' + stamp + ' ' + suffix));
  var email = swNormEmail_(options.email || ('test.cleanup.e2e.' + stamp.toLowerCase() + '.' + suffix.toLowerCase() + '@example.com'));
  var phone = swTrim_(options.phone || '555-010-' + suffix.slice(0, 4));
  var clientAdvisor = swTrim_(options.clientAdvisorName || 'Lyn');
  var clientAdvisorEmail = swNormEmail_(options.clientAdvisorEmail || 'lyn@ctyhp.com');
  var joc = swTrim_(options.jocName || 'Mark');
  var jocEmail = swNormEmail_(options.jocEmail || 'oc002@ctyhp.com');
  var salesStage = swTrim_(options.salesStage || 'Hot Lead');
  var convStatus = swTrim_(options.convStatus || 'Quotation Requested');
  var customOrder = swTrim_(options.customOrder || '3D Requested');
  var centerStone = swTrim_(options.centerStone || 'No Center Stone');
  var nextSteps = swTrim_(options.nextSteps || ('E2E stale cleanup fixture created by ' + (user.email || user.name || 'admin') + ' on ' + stamp + '.'));
  var visitDate = swTrim_(options.visitDate || Utilities.formatDate(touchAt, tz, 'M/d/yyyy'));
  var visitTime = swTrim_(options.visitTime || '2:00 PM');
  var visitType = swTrim_(options.visitType || 'Engagement Ring');
  var bookedAtIso = swTrim_(options.bookedAt || swIso_(touchAt));
  var updatedAtIso = swTrim_(options.updatedAt || swIso_(touchAt));
  var brand = swTrim_(options.brand || 'VVS');

  function setAt(index, value) {
    if (!(index >= 0) || index >= row.length) return;
    row[index] = value == null ? '' : value;
  }

  setAt(idx.appt, apptId);
  setAt(idx.root, rootApptId);
  setAt(idx.uid, fixtureId);
  setAt(idx.name, customerName);
  setAt(idx.email, email);
  setAt(idx.emailLower, email);
  setAt(idx.phone, phone);
  setAt(idx.phoneNorm, swNormPhone_(phone));
  setAt(idx.brand, brand);
  setAt(idx.bookedAt, bookedAtIso);
  setAt(idx.visitDate, visitDate);
  setAt(idx.visitTime, visitTime);
  setAt(idx.visitType, visitType);
  setAt(idx.status, swTrim_(options.status || 'Booked'));
  setAt(idx.active, swTrim_(options.active || 'Y'));
  setAt(idx.assignedRep, clientAdvisor);
  setAt(idx.assignedRepEmail, clientAdvisorEmail);
  setAt(idx.assistedRep, joc);
  setAt(idx.assistedRepEmail, jocEmail);
  setAt(idx.salesStage, salesStage);
  setAt(idx.convStatus, convStatus);
  setAt(idx.customOrder, customOrder);
  setAt(idx.inProduction, swTrim_(options.inProduction || ''));
  setAt(idx.centerStoneStatus, centerStone);
  setAt(idx.nextSteps, nextSteps);
  setAt(idx.updatedAt, updatedAtIso);

  master.appendRow(row);
  var rowNumber = master.getLastRow();

  return {
    ok: true,
    fixtureId: fixtureId,
    rowNumber: rowNumber,
    rootApptId: rootApptId,
    apptId: apptId,
    customerName: customerName,
    email: email,
    clientAdvisor: clientAdvisor,
    clientAdvisorEmail: clientAdvisorEmail,
    joc: joc,
    jocEmail: jocEmail,
    salesStage: salesStage,
    convStatus: convStatus,
    customOrder: customOrder,
    staleDays: staleDays,
    touchAt: bookedAtIso,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    serviceUrl: ScriptApp.getService().getUrl()
  };
}
