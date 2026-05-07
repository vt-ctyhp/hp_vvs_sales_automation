/**
 * One-time test customer data cleanup.
 *
 * Dry run:
 *   sw_previewTestDataCleanupOnce()
 *
 * Apply:
 *   sw_applyTestDataCleanupOnce({ confirmationToken: '<token from preview>' })
 *
 * This deletes source rows only. Generated _SW_* read-model tabs are left alone
 * and are invalidated/rebuilt after source rows are deleted.
 */

var SW_TEST_DATA_CLEANUP_PREVIEW_PROPERTY_ = 'SW_TEST_DATA_CLEANUP_ONCE_LAST_PREVIEW';
var SW_TEST_DATA_CLEANUP_PREVIEW_MAX_AGE_MS_ = 4 * 60 * 60 * 1000;
var SW_TEST_DATA_CLEANUP_REASON_ = 'One-time test customer data cleanup';

var SW_TEST_DATA_CLEANUP_WORKFLOW_SOURCE_ORDER_ = [
  '00_Master Appointments',
  '02_Form_Inbox',
  '_ExternalBookingEvents',
  '_IntakeQueue',
  '_SalesTaskQueue',
  '_SalesTaskLog',
  '_AppointmentArtifacts',
  '_SalesDataCleanup',
  '03_Client_Status_Log',
  '05_Wax_Requests',
  '07_Root_Index'
];

var SW_TEST_DATA_CLEANUP_READ_MODEL_SHEETS_ = {
  '_SW_TaskReadModel': true,
  '_SW_CustomerReadModel': true,
  '_SW_DiamondReadModel': true,
  '_SW_DiamondRootReadModel': true,
  '_SW_AppointmentReadModel': true,
  '_SW_CalendarMonthReadModel': true,
  '_SW_PaymentReadModel': true,
  '_SW_AdminDashboardReadModel': true,
  '_SW_ReadModelMeta': true
};

function sw_previewTestDataCleanupOnce(options) {
  options = options || {};
  swRequireTestDataCleanupAdmin_(options);
  var plan = swBuildTestDataCleanupPlan_(options);
  swStoreTestDataCleanupPreview_(plan);
  swLogTestDataCleanupPlan_(plan, 'SW_TEST_DATA_CLEANUP_PREVIEW');
  return swPublicTestDataCleanupResult_(plan);
}

function sw_applyTestDataCleanupOnce(options) {
  options = swNormalizeTestDataCleanupApplyOptions_(options);
  swRequireTestDataCleanupAdmin_(options);

  var lock = LockService.getScriptLock();
  lock.waitLock(Number(options.lockWaitMs || 30000));
  try {
    var plan = swBuildTestDataCleanupPlan_(options);
    var validation = swValidateTestDataCleanupPreview_(plan, options);
    plan.previewValidation = validation;
    if (!validation.ok) {
      plan.ok = false;
      plan.errors.push({ type: 'preview', message: validation.message });
      swLogTestDataCleanupPlan_(plan, 'SW_TEST_DATA_CLEANUP_APPLY_BLOCKED');
      return swPublicTestDataCleanupResult_(plan);
    }

    var confirmation = swConfirmTestDataCleanupApply_(plan, options);
    plan.confirmation = confirmation;
    if (!confirmation.ok) {
      plan.ok = false;
      plan.errors.push({ type: 'confirmation', message: confirmation.message || 'Apply was not confirmed.' });
      swLogTestDataCleanupPlan_(plan, 'SW_TEST_DATA_CLEANUP_APPLY_CANCELLED');
      return swPublicTestDataCleanupResult_(plan);
    }

    if (plan.totalCandidateRows > 0) {
      plan.deletedRows = swDeleteTestDataCleanupRows_(plan);
      plan.deletedCount = plan.deletedRows.length;
      plan.invalidation = swInvalidateAfterTestDataCleanup_(plan, options);
      swRecordTestDataCleanupResultFailures_(plan, plan.invalidation, 'invalidation');
      if (options.rebuildReadModels !== false) {
        plan.readModelRebuild = swRebuildAfterTestDataCleanup_();
        swRecordTestDataCleanupResultFailures_(plan, plan.readModelRebuild, 'readModelRebuild');
      }
    }

    plan.ok = plan.errors.length === 0;
    swLogTestDataCleanupPlan_(plan, 'SW_TEST_DATA_CLEANUP_APPLY');
    return swPublicTestDataCleanupResult_(plan);
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function swBuildTestDataCleanupPlan_(options) {
  options = options || {};
  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  var plan = {
    ok: true,
    apply: options.apply === true,
    createdAt: swTestDataCleanupIso_(new Date()),
    workflowSpreadsheetId: swTestDataCleanupSpreadsheetId_(ss),
    workflowSpreadsheetName: swTestDataCleanupSpreadsheetName_(ss),
    confirmationToken: '',
    fingerprint: '',
    totalCandidateRows: 0,
    deletedCount: 0,
    keyset: swNewTestDataCleanupKeyset_(),
    sources: [],
    skippedSources: [],
    candidates: [],
    deletedRows: [],
    invalidation: null,
    readModelRebuild: null,
    warnings: [],
    errors: []
  };

  swAddTestDataCleanupWarning_(plan, 'Generated _SW_* read-model tabs are not row-deleted. They are invalidated/rebuilt after source cleanup.');
  swAddTestDataCleanupWarning_(plan, 'Payment ledger cleanup uses exact RootApptID/APPT_ID/SO matches from the workflow keyset, not name-only matching.');
  swAddTestDataCleanupWarning_(plan, 'Drive files and folders referenced by artifact rows are not trashed by this function.');

  swScanWorkflowTestDataCleanupSources_(ss, plan, options);
  swScanPaymentLedgerForTestDataCleanup_(plan, options);
  swScanDiamond200ForTestDataCleanup_(plan, options);

  plan.candidates.sort(swCompareTestDataCleanupCandidates_);
  plan.totalCandidateRows = plan.candidates.length;
  plan.fingerprint = swTestDataCleanupFingerprint_(plan);
  plan.confirmationToken = 'DELETE_TEST_DATA_' + plan.fingerprint.slice(0, 10).toUpperCase();
  plan.summary = swTestDataCleanupSummary_(plan);
  return plan;
}

function swRequireTestDataCleanupAdmin_(options) {
  options = options || {};
  var token = swTestDataCleanupTrim_(options.authToken || options.token || '');
  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  if (typeof swAuthUserForApi_ !== 'function') {
    throw new Error('Admin access required, but workflow auth helpers are unavailable.');
  }
  var user = swAuthUserForApi_(ss, token);
  if (!user || !user.isAdmin) throw new Error('Admin access required for one-time test data cleanup.');
  return user;
}

function swScanWorkflowTestDataCleanupSources_(ss, plan, options) {
  var sheetNames = swWorkflowTestDataCleanupSheetNames_(options);
  for (var i = 0; i < sheetNames.length; i++) {
    var name = sheetNames[i];
    if (SW_TEST_DATA_CLEANUP_READ_MODEL_SHEETS_[name]) {
      plan.skippedSources.push({ workbookKey: 'workflow', sheetName: name, reason: 'generated read model' });
      continue;
    }
    var sh = ss.getSheetByName(name);
    if (!sh) {
      plan.skippedSources.push({ workbookKey: 'workflow', sheetName: name, reason: 'missing sheet' });
      continue;
    }
    var target = {
      workbookKey: 'workflow',
      workbookLabel: 'Workflow workbook',
      spreadsheet: ss,
      sheet: sh,
      sheetName: sh.getName(),
      headerRows: 1,
      dataStartRow: 2,
      allowDirectTestMatch: name !== '_SalesTaskLog' && name !== '07_Root_Index' && name !== '_AppointmentArtifacts',
      exactKeyOnly: name === '_SalesTaskLog' || name === '07_Root_Index' || name === '_AppointmentArtifacts',
      deleteOrder: swTestDataCleanupDeleteOrder_(name, 'workflow')
    };
    swScanSheetForTestDataCleanup_(target, plan, options);
  }
}

function swWorkflowTestDataCleanupSheetNames_(options) {
  var names = SW_TEST_DATA_CLEANUP_WORKFLOW_SOURCE_ORDER_.slice();
  if (typeof SW_SHEETS !== 'undefined') {
    names = [
      SW_SHEETS.MASTER || '00_Master Appointments',
      '02_Form_Inbox',
      SW_SHEETS.EXTERNAL_BOOKING_EVENTS || '_ExternalBookingEvents',
      '_IntakeQueue',
      SW_SHEETS.TASKS || '_SalesTaskQueue',
      SW_SHEETS.LOG || '_SalesTaskLog',
      SW_SHEETS.APPOINTMENT_ARTIFACTS || '_AppointmentArtifacts',
      SW_SHEETS.DATA_CLEANUP || '_SalesDataCleanup',
      '03_Client_Status_Log',
      '05_Wax_Requests',
      '07_Root_Index'
    ];
  }
  if (options && options.workflowSheets && options.workflowSheets.length) {
    names = options.workflowSheets.slice();
  }
  return swUniqueStrings_(names);
}

function swScanPaymentLedgerForTestDataCleanup_(plan, options) {
  if (options && options.includePayments === false) return;
  var target = null;
  try {
    if (typeof rp_getLedgerTarget === 'function') {
      target = rp_getLedgerTarget();
    } else if (typeof pr_getLedger_ === 'function' && typeof pr_getPaymentsSheet_ === 'function') {
      var ledger = pr_getLedger_();
      target = { ss: ledger, sh: pr_getPaymentsSheet_(ledger), resolved: { ledgerFileKey: 'pr_getLedger_', ledgerSheetKey: 'pr_getPaymentsSheet_' } };
    }
  } catch (err) {
    plan.skippedSources.push({ workbookKey: 'paymentsLedger', reason: err && err.message ? err.message : String(err) });
    return;
  }
  if (!target || !target.sh) {
    plan.skippedSources.push({ workbookKey: 'paymentsLedger', reason: 'payments ledger helper unavailable' });
    return;
  }
  var ledgerSafety = swPaymentLedgerTargetSheetIsSafe_(target);
  if (!ledgerSafety.ok) {
    plan.skippedSources.push({ workbookKey: 'paymentsLedger', sheetName: target.sh.getName(), reason: ledgerSafety.reason });
    return;
  }

  swScanSheetForTestDataCleanup_({
    workbookKey: 'paymentsLedger',
    workbookLabel: 'Payments ledger',
    spreadsheet: target.ss || target.sh.getParent(),
    sheet: target.sh,
    sheetName: target.sh.getName(),
    headerRows: 1,
    dataStartRow: 2,
    allowDirectTestMatch: false,
    exactKeyOnly: true,
    deleteOrder: 10,
    note: target.resolved || {}
  }, plan, options);
}

function swPaymentLedgerTargetSheetIsSafe_(target) {
  var actual = '';
  try { actual = target && target.sh ? target.sh.getName() : ''; } catch (_) {}
  if (!actual) return { ok: false, reason: 'payments ledger sheet name unavailable' };

  var expected = 'Payments';
  var configured = false;
  try {
    if (typeof RP_KEY_ALIASES !== 'undefined' && RP_KEY_ALIASES.LEDGER_SHEET_NAME && typeof rp_propOneOf_ === 'function') {
      var sheetRes = rp_propOneOf_(RP_KEY_ALIASES.LEDGER_SHEET_NAME, { required: false, label: 'Payments sheet name' });
      if (sheetRes && sheetRes.value) {
        expected = String(sheetRes.value).trim() || expected;
        configured = true;
      }
    }
  } catch (_) {}

  if (configured && actual !== expected) {
    return { ok: false, reason: 'configured payments sheet "' + expected + '" was not resolved; refusing fallback tab "' + actual + '"' };
  }
  if (!configured && actual !== expected && !/payment/i.test(actual)) {
    return { ok: false, reason: 'default payments tab was not found; refusing fallback tab "' + actual + '"' };
  }
  return { ok: true };
}

function swScanDiamond200ForTestDataCleanup_(plan, options) {
  if (options && options.includeDiamonds === false) return;
  var target = null;
  try {
    if (typeof swDiamond200Target_ === 'function') target = swDiamond200Target_();
  } catch (err) {
    plan.skippedSources.push({ workbookKey: 'diamonds200', reason: err && err.message ? err.message : String(err) });
    return;
  }
  if (!target || !target.sheet) {
    plan.skippedSources.push({ workbookKey: 'diamonds200', reason: 'diamond workbook helper unavailable' });
    return;
  }

  swScanSheetForTestDataCleanup_({
    workbookKey: 'diamonds200',
    workbookLabel: '200 diamond workbook',
    spreadsheet: target.ss || target.sheet.getParent(),
    sheet: target.sheet,
    sheetName: target.sheet.getName(),
    headerRows: 2,
    dataStartRow: 3,
    allowDirectTestMatch: false,
    exactKeyOnly: true,
    deleteOrder: 20,
    note: { tab: target.tab || target.sheet.getName() }
  }, plan, options);
}

function swScanSheetForTestDataCleanup_(target, plan, options) {
  var sh = target.sheet;
  var lr = sh.getLastRow();
  var lc = sh.getLastColumn();
  var sourceInfo = {
    workbookKey: target.workbookKey,
    workbookLabel: target.workbookLabel,
    spreadsheetId: swTestDataCleanupSpreadsheetId_(target.spreadsheet),
    spreadsheetName: swTestDataCleanupSpreadsheetName_(target.spreadsheet),
    sheetName: target.sheetName,
    sheetId: sh.getSheetId(),
    lastRow: lr,
    lastColumn: lc,
    note: target.note || {}
  };
  plan.sources.push(sourceInfo);

  if (lr < target.dataStartRow || lc < 1) return;

  var headerInfo = swReadTestDataCleanupHeaders_(sh, target.headerRows || 1, lc);
  var columns = swTestDataCleanupColumns_(headerInfo);
  var usableColumns = swTestDataCleanupUsableIndexes_(columns);
  if (!usableColumns.length) {
    plan.skippedSources.push({ workbookKey: target.workbookKey, sheetName: target.sheetName, reason: 'no usable key/name/email columns' });
    return;
  }

  var rowCount = lr - target.dataStartRow + 1;
  var values = sh.getRange(target.dataStartRow, 1, rowCount, lc).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    var rowNumber = target.dataStartRow + i;
    var row = values[i];
    var rec = swTestDataCleanupRecordFromRow_(row, columns);
    var match = swTestDataCleanupMatchRecord_(rec, target, plan.keyset, options);
    if (!match.matched) continue;

    var candidate = swTestDataCleanupCandidate_(target, sourceInfo, rowNumber, rec, match);
    if (swAddTestDataCleanupCandidate_(plan, candidate)) {
      swAddRecordToTestDataCleanupKeyset_(plan.keyset, rec, match);
    }
  }
}

function swReadTestDataCleanupHeaders_(sh, headerRows, lastColumn) {
  var rows = sh.getRange(1, 1, Math.max(1, headerRows), lastColumn).getDisplayValues();
  var headers = [];
  for (var c = 0; c < lastColumn; c++) {
    var pieces = [];
    for (var r = 0; r < rows.length; r++) {
      var piece = swTestDataCleanupTrim_(rows[r][c]);
      if (piece) pieces.push(piece);
    }
    headers.push(pieces.join(' '));
  }
  return {
    headers: headers,
    map: swTestDataCleanupHeaderMap_(headers)
  };
}

function swTestDataCleanupColumns_(headerInfo) {
  var H = headerInfo.map;
  return {
    root: swTestDataCleanupPick_(H, ['RootApptID', 'Root Appt ID', 'Root Appointment ID', 'ROOT', 'Root_ID']),
    appt: swTestDataCleanupPick_(H, ['APPT_ID', 'Appt ID', 'Appointment ID']),
    uid: swTestDataCleanupPick_(H, ['CalendlyEventUID', 'Calendly Event UID', 'Admin: Calendly Event UID', 'Acuity ID', 'UID', 'External Booking UID', 'ProviderAppointmentID', 'Provider Appointment ID']),
    taskId: swTestDataCleanupPick_(H, ['TaskID', 'Task ID']),
    name: swTestDataCleanupPick_(H, ['Customer Name', 'Customer', 'Client Name', 'Name', 'Full Name']),
    emailLower: swTestDataCleanupPick_(H, ['EmailLower', 'Email Lower']),
    email: swTestDataCleanupPick_(H, ['Email', 'Email Address', 'E-mail']),
    phone: swTestDataCleanupPick_(H, ['PhoneNorm', 'Phone Norm', 'Phone', 'Phone Number', 'Mobile', 'Tel']),
    so: swTestDataCleanupPick_(H, ['SO#', 'SO #', 'SO', 'SO Number', 'Sales Order', 'Sales Order #']),
    brand: swTestDataCleanupPick_(H, ['Brand', 'Company']),
    visitDate: swTestDataCleanupPick_(H, ['Visit Date', 'Appointment Date', 'Date', 'PaymentDateTime', 'Payment DateTime']),
    status: swTestDataCleanupPick_(H, ['Status', 'DocStatus', 'Doc Status']),
    paymentId: swTestDataCleanupPick_(H, ['PAYMENT_ID', 'Payment ID', 'PaymentId']),
    docNumber: swTestDataCleanupPick_(H, ['DocNumber', 'Doc #', 'Document Number']),
    payloadJson: swTestDataCleanupPick_(H, ['Payload JSON', 'Payload', 'Request JSON', 'Raw JSON', 'RawPayloadJSON', 'JSON', 'Event JSON'])
  };
}

function swTestDataCleanupUsableIndexes_(columns) {
  var out = [];
  Object.keys(columns || {}).forEach(function (key) {
    var idx = Number(columns[key]);
    if (isFinite(idx) && idx >= 0 && out.indexOf(idx) < 0) out.push(idx);
  });
  return out;
}

function swTestDataCleanupRecordFromRow_(row, columns) {
  var rec = {
    root: swTestDataCleanupCleanId_(swTestDataCleanupCell_(row, columns.root)),
    appt: swTestDataCleanupCleanId_(swTestDataCleanupCell_(row, columns.appt)),
    uid: swTestDataCleanupCleanId_(swTestDataCleanupCell_(row, columns.uid)),
    taskId: swTestDataCleanupCleanId_(swTestDataCleanupCell_(row, columns.taskId)),
    customerName: swTestDataCleanupTrim_(swTestDataCleanupCell_(row, columns.name)),
    email: swTestDataCleanupNormEmail_(swTestDataCleanupCell_(row, columns.emailLower) || swTestDataCleanupCell_(row, columns.email)),
    phone: swTestDataCleanupNormPhone_(swTestDataCleanupCell_(row, columns.phone)),
    so: swTestDataCleanupCleanId_(swTestDataCleanupCell_(row, columns.so)),
    brand: swTestDataCleanupTrim_(swTestDataCleanupCell_(row, columns.brand)),
    visitDate: swTestDataCleanupTrim_(swTestDataCleanupCell_(row, columns.visitDate)),
    status: swTestDataCleanupTrim_(swTestDataCleanupCell_(row, columns.status)),
    paymentId: swTestDataCleanupCleanId_(swTestDataCleanupCell_(row, columns.paymentId)),
    docNumber: swTestDataCleanupCleanId_(swTestDataCleanupCell_(row, columns.docNumber)),
    payloadJson: swTestDataCleanupTrim_(swTestDataCleanupCell_(row, columns.payloadJson))
  };
  swMergePayloadIntoTestDataCleanupRecord_(rec);
  return rec;
}

function swMergePayloadIntoTestDataCleanupRecord_(rec) {
  if (!rec.payloadJson) return;
  var payload = swParseTestDataCleanupPayload_(rec.payloadJson);
  if (!payload) return;

  rec.root = rec.root || swTestDataCleanupCleanId_(payload.root);
  rec.appt = rec.appt || swTestDataCleanupCleanId_(payload.appt);
  rec.uid = rec.uid || swTestDataCleanupCleanId_(payload.uid);
  rec.customerName = rec.customerName || swTestDataCleanupTrim_(payload.customerName);
  rec.email = rec.email || swTestDataCleanupNormEmail_(payload.email);
  rec.phone = rec.phone || swTestDataCleanupNormPhone_(payload.phone);
  rec.so = rec.so || swTestDataCleanupCleanId_(payload.so);
}

function swParseTestDataCleanupPayload_(text) {
  var parsed = null;
  try { parsed = JSON.parse(text); } catch (_) {}
  if (!parsed) return null;

  var out = {};
  swWalkTestDataCleanupPayload_(parsed, out, 0);
  return out;
}

function swWalkTestDataCleanupPayload_(value, out, depth) {
  if (depth > 6 || value == null) return;
  if (Array.isArray(value)) {
    for (var i = 0; i < value.length; i++) swWalkTestDataCleanupPayload_(value[i], out, depth + 1);
    return;
  }
  if (typeof value !== 'object') return;

  Object.keys(value).forEach(function (key) {
    var v = value[key];
    var hk = swTestDataCleanupHeaderKey_(key);
    if (typeof v !== 'object' || v == null) {
      if (!out.root && (hk === 'rootapptid' || hk === 'rootappt' || hk === 'root')) out.root = v;
      if (!out.appt && (hk === 'apptid' || hk === 'appointmentid')) out.appt = v;
      if (!out.uid && (hk === 'calendlyeventuid' || hk === 'externalbookinguid' || hk === 'uid')) out.uid = v;
      if (!out.customerName && (hk === 'customername' || hk === 'clientname' || hk === 'name' || hk === 'fullname')) out.customerName = v;
      if (!out.email && (hk === 'email' || hk === 'emaillower' || hk === 'emailaddress')) out.email = v;
      if (!out.phone && (hk === 'phone' || hk === 'phonenumber' || hk === 'phonenorm')) out.phone = v;
      if (!out.so && (hk === 'so' || hk === 'sonumber' || hk === 'salesorder' || hk === 'salesordernumber')) out.so = v;
    }
    swWalkTestDataCleanupPayload_(v, out, depth + 1);
  });
}

function swTestDataCleanupMatchRecord_(rec, target, keyset, options) {
  var keyMatch = swTestDataCleanupExactKeyMatch_(rec, keyset);
  if (keyMatch.matched) return keyMatch;

  if (target.exactKeyOnly) return { matched: false };

  if (target.allowDirectTestMatch !== false) {
    var direct = swTestDataCleanupDirectMatch_(rec, options);
    if (direct.matched) return direct;
  }

  return { matched: false };
}

function swTestDataCleanupExactKeyMatch_(rec, keyset) {
  if (rec.taskId && keyset.taskIds[rec.taskId]) return { matched: true, matchType: 'taskId', matchedField: 'TaskID', matchedValue: rec.taskId, reason: 'exact task ID from test data keyset' };
  if (rec.root && keyset.roots[rec.root]) return { matched: true, matchType: 'root', matchedField: 'RootApptID', matchedValue: rec.root, reason: 'exact root ID from test data keyset' };
  if (rec.appt && keyset.appts[rec.appt]) return { matched: true, matchType: 'appt', matchedField: 'APPT_ID', matchedValue: rec.appt, reason: 'exact appointment ID from test data keyset' };
  if (rec.uid && keyset.uids[rec.uid]) return { matched: true, matchType: 'uid', matchedField: 'CalendlyEventUID', matchedValue: rec.uid, reason: 'exact booking UID from test data keyset' };
  if (rec.so && keyset.sos[rec.so]) return { matched: true, matchType: 'so', matchedField: 'SO#', matchedValue: rec.so, reason: 'exact SO from test data keyset' };
  return { matched: false };
}

function swTestDataCleanupDirectMatch_(rec, options) {
  var name = swTestDataCleanupTrim_(rec.customerName);
  var email = swTestDataCleanupNormEmail_(rec.email);
  var mode = swTestDataCleanupNorm_(options && options.matchMode || 'strict');

  if (name && swTestDataCleanupTextLooksTest_(name, mode)) {
    return { matched: true, matchType: 'directName', matchedField: 'Customer Name', matchedValue: name, reason: 'customer name contains a test-data token' };
  }
  if (email && swTestDataCleanupEmailLooksTest_(email, mode)) {
    return { matched: true, matchType: 'directEmail', matchedField: 'Email', matchedValue: email, reason: 'email contains a test-data token' };
  }
  return { matched: false };
}

function swTestDataCleanupTextLooksTest_(value, mode) {
  var text = swTestDataCleanupNorm_(value);
  if (!text) return false;
  if (mode === 'contains') return /test|testing|tester|sample|dummy|fake/.test(text);
  var spaced = text.replace(/[^a-z0-9]+/g, ' ');
  return /(^| )(test|testing|tester|testclient|testcustomer|sample|dummy|fake)([0-9]*)( |$)/.test(spaced);
}

function swTestDataCleanupEmailLooksTest_(email, mode) {
  email = swTestDataCleanupNormEmail_(email);
  if (!email) return false;
  if (mode === 'contains') return /test|testing|tester|sample|dummy|fake/.test(email);

  var parts = email.split('@');
  var local = parts[0] || '';
  var domain = parts.slice(1).join('@') || '';
  var token = /(^|[._+\-])(test|testing|tester|testclient|testcustomer|sample|dummy|fake)([0-9]*)([._+\-]|$)/;
  return token.test(local) || token.test(domain);
}

function swTestDataCleanupCandidate_(target, sourceInfo, rowNumber, rec, match) {
  return {
    workbookKey: target.workbookKey,
    workbookLabel: target.workbookLabel,
    spreadsheetId: sourceInfo.spreadsheetId,
    spreadsheetName: sourceInfo.spreadsheetName,
    sheetName: target.sheetName,
    sheetId: target.sheet.getSheetId(),
    rowNumber: rowNumber,
    deleteOrder: target.deleteOrder || 100,
    root: rec.root || '',
    appt: rec.appt || '',
    uid: rec.uid || '',
    taskId: rec.taskId || '',
    so: rec.so || '',
    customerName: rec.customerName || '',
    email: rec.email || '',
    phone: rec.phone || '',
    brand: rec.brand || '',
    visitDate: rec.visitDate || '',
    status: rec.status || '',
    paymentId: rec.paymentId || '',
    docNumber: rec.docNumber || '',
    matchType: match.matchType || '',
    matchedField: match.matchedField || '',
    matchedValue: match.matchedValue || '',
    reason: match.reason || '',
    headerRows: target.headerRows || 1,
    rowFingerprint: swTestDataCleanupRowFingerprint_(rec)
  };
}

function swAddTestDataCleanupCandidate_(plan, candidate) {
  var key = [
    candidate.workbookKey,
    candidate.spreadsheetId,
    candidate.sheetId,
    candidate.rowNumber
  ].join('|');
  if (!plan._candidateKeys) plan._candidateKeys = {};
  if (plan._candidateKeys[key]) return false;
  plan._candidateKeys[key] = true;
  plan.candidates.push(candidate);
  return true;
}

function swAddRecordToTestDataCleanupKeyset_(keyset, rec, match) {
  match = match || {};
  swSetIfValue_(keyset.roots, rec.root);
  swSetIfValue_(keyset.appts, rec.appt);
  swSetIfValue_(keyset.uids, rec.uid);
  swSetIfValue_(keyset.taskIds, rec.taskId);
  if (match.matchType !== 'directName' && match.matchType !== 'directEmail') {
    swSetIfValue_(keyset.sos, rec.so);
  }
  swSetIfValue_(keyset.emails, rec.email);
  swSetIfValue_(keyset.phones, rec.phone);
}

function swNewTestDataCleanupKeyset_() {
  return {
    roots: {},
    appts: {},
    uids: {},
    taskIds: {},
    sos: {},
    emails: {},
    phones: {}
  };
}

function swSetIfValue_(set, value) {
  value = swTestDataCleanupTrim_(value);
  if (value) set[value] = true;
}

function swDeleteTestDataCleanupRows_(plan) {
  var deleted = [];
  var groups = {};
  for (var i = 0; i < plan.candidates.length; i++) {
    var c = plan.candidates[i];
    var key = [c.workbookKey, c.spreadsheetId, c.sheetId].join('|');
    if (!groups[key]) groups[key] = { sample: c, rows: [] };
    groups[key].rows.push(c);
  }

  var groupList = Object.keys(groups).map(function (key) { return groups[key]; });
  groupList.sort(function (a, b) {
    var ao = a.sample.deleteOrder || 100;
    var bo = b.sample.deleteOrder || 100;
    if (ao !== bo) return ao - bo;
    return String(a.sample.sheetName).localeCompare(String(b.sample.sheetName));
  });

  for (var g = 0; g < groupList.length; g++) {
    var group = groupList[g];
    var sh = swOpenTestDataCleanupSheet_(group.sample);
    if (!sh) {
      plan.errors.push({ type: 'sheet', sheetName: group.sample.sheetName, message: 'Sheet unavailable at apply time.' });
      continue;
    }

    group.rows.sort(function (a, b) { return b.rowNumber - a.rowNumber; });
    for (var r = 0; r < group.rows.length; r++) {
      var row = group.rows[r];
      try {
        var recheck = swCurrentTestDataCleanupRowMatchesCandidate_(sh, row);
        if (!recheck.ok) {
          plan.errors.push({
            type: 'rowRevalidation',
            workbookKey: row.workbookKey,
            sheetName: row.sheetName,
            rowNumber: row.rowNumber,
            message: recheck.message || 'Current row no longer matches previewed candidate.'
          });
          continue;
        }
        sh.deleteRow(row.rowNumber);
        deleted.push(row);
      } catch (err) {
        plan.errors.push({
          type: 'row',
          workbookKey: row.workbookKey,
          sheetName: row.sheetName,
          rowNumber: row.rowNumber,
          message: err && err.message ? err.message : String(err)
        });
      }
    }
  }
  return deleted;
}

function swCurrentTestDataCleanupRowMatchesCandidate_(sh, candidate) {
  try {
    if (!sh || candidate.rowNumber < 1 || candidate.rowNumber > sh.getLastRow()) {
      return { ok: false, message: 'Candidate row is outside the current sheet range.' };
    }
    var lc = sh.getLastColumn();
    var headerInfo = swReadTestDataCleanupHeaders_(sh, candidate.headerRows || 1, lc);
    var columns = swTestDataCleanupColumns_(headerInfo);
    var row = sh.getRange(candidate.rowNumber, 1, 1, lc).getDisplayValues()[0];
    var rec = swTestDataCleanupRecordFromRow_(row, columns);
    var currentFingerprint = swTestDataCleanupRowFingerprint_(rec);
    if (currentFingerprint !== candidate.rowFingerprint) {
      return { ok: false, message: 'Candidate row fingerprint changed since preview.' };
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err && err.message ? err.message : String(err) };
  }
}

function swOpenTestDataCleanupSheet_(candidate) {
  var ss = null;
  try {
    if (candidate.workbookKey === 'workflow') {
      ss = swTestDataCleanupWorkflowSpreadsheet_();
    } else {
      ss = SpreadsheetApp.openById(candidate.spreadsheetId);
    }
    return ss.getSheetByName(candidate.sheetName);
  } catch (_) {}
  return null;
}

function swValidateTestDataCleanupPreview_(plan, options) {
  if (plan.totalCandidateRows === 0) return { ok: true, empty: true };

  var preview = swReadTestDataCleanupPreview_();
  if (!preview) {
    return { ok: false, message: 'Run sw_previewTestDataCleanupOnce() before apply.' };
  }
  if (preview.workflowSpreadsheetId !== plan.workflowSpreadsheetId) {
    return { ok: false, message: 'Preview was created for a different workflow spreadsheet.' };
  }
  if (preview.fingerprint !== plan.fingerprint) {
    return { ok: false, message: 'Current candidate rows differ from the stored preview. Run preview again before apply.' };
  }
  var previewAt = swTestDataCleanupDateMs_(preview.createdAt);
  if (previewAt && new Date().getTime() - previewAt > SW_TEST_DATA_CLEANUP_PREVIEW_MAX_AGE_MS_) {
    return { ok: false, message: 'Stored preview is older than four hours. Run preview again before apply.' };
  }
  if (options.confirmationToken && options.confirmationToken !== preview.confirmationToken) {
    return { ok: false, message: 'Confirmation token does not match the stored preview token.' };
  }
  return { ok: true, previewCreatedAt: preview.createdAt };
}

function swConfirmTestDataCleanupApply_(plan, options) {
  if (plan.totalCandidateRows === 0) return { ok: true, skipped: true, message: 'No candidate rows.' };
  if (options.confirmationToken && options.confirmationToken === plan.confirmationToken) {
    return { ok: true, source: 'options.confirmationToken' };
  }

  try {
    var ui = SpreadsheetApp.getUi();
    var preview = plan.candidates.slice(0, 12).map(function (row) {
      return row.workbookKey + ' / ' + row.sheetName + ' row ' + row.rowNumber + ': ' +
        (row.customerName || row.email || row.root || row.appt || row.so || row.taskId || '(matched row)');
    }).join('\n');
    var more = plan.totalCandidateRows > 12 ? '\n...and ' + (plan.totalCandidateRows - 12) + ' more row(s)' : '';
    var prompt = ui.prompt(
      'Confirm test data cleanup',
      'This will permanently delete ' + plan.totalCandidateRows + ' source row(s), bottom-up.\n\n' +
        preview + more + '\n\n' +
        'Type this token to apply:\n' + plan.confirmationToken,
      ui.ButtonSet.OK_CANCEL
    );
    var button = prompt.getSelectedButton();
    var response = swTestDataCleanupTrim_(prompt.getResponseText());
    if (button === ui.Button.OK && response === plan.confirmationToken) {
      return { ok: true, source: 'ui.prompt' };
    }
    return { ok: false, source: 'ui.prompt', message: 'Confirmation token was not entered.' };
  } catch (_) {}

  return { ok: false, source: 'unavailable', message: 'No UI confirmation was available and no confirmationToken option was provided.' };
}

function swNormalizeTestDataCleanupApplyOptions_(options) {
  if (typeof options === 'string') return { apply: true, confirmationToken: options };
  options = options || {};
  options.apply = true;
  return options;
}

function swStoreTestDataCleanupPreview_(plan) {
  var payload = {
    workflowSpreadsheetId: plan.workflowSpreadsheetId,
    workflowSpreadsheetName: plan.workflowSpreadsheetName,
    fingerprint: plan.fingerprint,
    confirmationToken: plan.confirmationToken,
    totalCandidateRows: plan.totalCandidateRows,
    summary: plan.summary,
    createdAt: plan.createdAt
  };
  try {
    PropertiesService.getDocumentProperties().setProperty(
      SW_TEST_DATA_CLEANUP_PREVIEW_PROPERTY_,
      JSON.stringify(payload)
    );
  } catch (err) {
    plan.errors.push({ type: 'previewStore', message: err && err.message ? err.message : String(err) });
  }
}

function swReadTestDataCleanupPreview_() {
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty(SW_TEST_DATA_CLEANUP_PREVIEW_PROPERTY_);
    return raw ? JSON.parse(raw) : null;
  } catch (_) {}
  return null;
}

function swInvalidateAfterTestDataCleanup_(plan, options) {
  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  var out = {
    appointment: null,
    payment: null,
    diamond: null,
    taskList: null,
    taskDashboard: null,
    customerSearch: null,
    artifactRoots: 0,
    broad: null
  };
  var reason = SW_TEST_DATA_CLEANUP_REASON_;
  var touched = swTestDataCleanupTouchedWorkbooks_(plan.deletedRows || []);

  try {
    if (typeof swInvalidateAppointmentReadModelsAfterWrite_ === 'function') {
      swInvalidateAppointmentReadModelsAfterWrite_(ss, reason);
      out.appointment = { ok: true, helper: 'swInvalidateAppointmentReadModelsAfterWrite_' };
    }
  } catch (err1) {
    out.appointment = { ok: false, error: err1 && err1.message ? err1.message : String(err1) };
  }

  if (touched.paymentsLedger) {
    try {
      if (typeof swInvalidatePaymentReadModelsAfterWrite_ === 'function') {
        swInvalidatePaymentReadModelsAfterWrite_(ss, reason);
        out.payment = { ok: true, helper: 'swInvalidatePaymentReadModelsAfterWrite_' };
      }
    } catch (err2) {
      out.payment = { ok: false, error: err2 && err2.message ? err2.message : String(err2) };
    }
  }

  if (touched.diamonds200) {
    try {
      if (typeof swInvalidateDiamondReadModelsAfterWrite_ === 'function') {
        swInvalidateDiamondReadModelsAfterWrite_(ss, reason);
        out.diamond = { ok: true, helper: 'swInvalidateDiamondReadModelsAfterWrite_' };
      }
    } catch (err3) {
      out.diamond = { ok: false, error: err3 && err3.message ? err3.message : String(err3) };
    }
  }

  try {
    if (typeof swInvalidateTaskListCache_ === 'function') {
      swInvalidateTaskListCache_(ss);
      out.taskList = { ok: true, helper: 'swInvalidateTaskListCache_' };
    }
  } catch (err4) {
    out.taskList = { ok: false, error: err4 && err4.message ? err4.message : String(err4) };
  }

  try {
    if (typeof swInvalidateTaskDashboardProjectionCache_ === 'function') {
      swInvalidateTaskDashboardProjectionCache_(ss);
      out.taskDashboard = { ok: true, helper: 'swInvalidateTaskDashboardProjectionCache_' };
    }
  } catch (err5) {
    out.taskDashboard = { ok: false, error: err5 && err5.message ? err5.message : String(err5) };
  }

  try {
    if (typeof swInvalidateCustomerSearchReadModelCache_ === 'function') {
      swInvalidateCustomerSearchReadModelCache_(ss);
      out.customerSearch = { ok: true, helper: 'swInvalidateCustomerSearchReadModelCache_' };
    }
  } catch (err6) {
    out.customerSearch = { ok: false, error: err6 && err6.message ? err6.message : String(err6) };
  }

  if (typeof swInvalidateAppointmentArtifactRowsForRoot_ === 'function') {
    var roots = swUniqueStrings_(Object.keys(plan.keyset.roots || {}));
    roots.forEach(function (root) {
      try {
        swInvalidateAppointmentArtifactRowsForRoot_(ss, root);
        out.artifactRoots++;
      } catch (_) {}
    });
  }

  try {
    if (typeof swMarkWorkflowReadModelsStale_ === 'function') {
      swMarkWorkflowReadModelsStale_(ss, reason);
      out.broad = { ok: true, helper: 'swMarkWorkflowReadModelsStale_' };
    }
  } catch (err7) {
    out.broad = { ok: false, error: err7 && err7.message ? err7.message : String(err7) };
  }

  return out;
}

function swRebuildAfterTestDataCleanup_() {
  try {
    if (typeof sw_rebuildWorkflowReadModels === 'function') {
      return sw_rebuildWorkflowReadModels({ reason: SW_TEST_DATA_CLEANUP_REASON_ });
    }
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
  return { ok: true, skipped: true, reason: 'sw_rebuildWorkflowReadModels unavailable' };
}

function swTestDataCleanupTouchedWorkbooks_(rows) {
  var out = {};
  (rows || []).forEach(function (row) {
    out[row.workbookKey] = true;
  });
  return out;
}

function swTestDataCleanupDeleteOrder_(sheetName, workbookKey) {
  if (workbookKey === 'paymentsLedger') return 10;
  if (workbookKey === 'diamonds200') return 20;
  if (sheetName === '_AppointmentArtifacts') return 30;
  if (sheetName === '_SalesTaskQueue') return 40;
  if (sheetName === '_SalesTaskLog') return 45;
  if (sheetName === '_IntakeQueue') return 50;
  if (sheetName === '02_Form_Inbox') return 60;
  if (sheetName === '_SalesDataCleanup') return 65;
  if (sheetName === '03_Client_Status_Log') return 70;
  if (sheetName === '05_Wax_Requests') return 72;
  if (sheetName === '07_Root_Index') return 74;
  if (sheetName === '00_Master Appointments') return 90;
  return 80;
}

function swTestDataCleanupSummary_(plan) {
  var bySource = {};
  (plan.candidates || []).forEach(function (row) {
    var key = row.workbookKey + ':' + row.sheetName;
    if (!bySource[key]) bySource[key] = { workbookKey: row.workbookKey, sheetName: row.sheetName, candidates: 0, deleted: 0 };
    bySource[key].candidates++;
  });
  (plan.deletedRows || []).forEach(function (row) {
    var key = row.workbookKey + ':' + row.sheetName;
    if (!bySource[key]) bySource[key] = { workbookKey: row.workbookKey, sheetName: row.sheetName, candidates: 0, deleted: 0 };
    bySource[key].deleted++;
  });
  return {
    totalCandidateRows: plan.totalCandidateRows || 0,
    deletedCount: plan.deletedCount || 0,
    sources: Object.keys(bySource).sort().map(function (key) { return bySource[key]; }),
    keyset: swPublicTestDataCleanupKeysetCounts_(plan.keyset),
    skippedSources: plan.skippedSources || [],
    warnings: plan.warnings || [],
    errors: plan.errors || []
  };
}

function swPublicTestDataCleanupResult_(plan) {
  plan.summary = swTestDataCleanupSummary_(plan);
  return {
    ok: plan.ok !== false,
    apply: plan.apply === true,
    createdAt: plan.createdAt,
    workflowSpreadsheetId: plan.workflowSpreadsheetId,
    workflowSpreadsheetName: plan.workflowSpreadsheetName,
    totalCandidateRows: plan.totalCandidateRows,
    deletedCount: plan.deletedCount || 0,
    confirmationToken: plan.confirmationToken,
    fingerprint: plan.fingerprint,
    summary: plan.summary,
    candidates: (plan.candidates || []).map(swPublicTestDataCleanupCandidate_),
    deletedRows: (plan.deletedRows || []).map(swPublicTestDataCleanupCandidate_),
    previewValidation: plan.previewValidation || null,
    confirmation: plan.confirmation || null,
    invalidation: plan.invalidation || null,
    readModelRebuild: plan.readModelRebuild || null
  };
}

function swPublicTestDataCleanupCandidate_(row) {
  return {
    workbookKey: row.workbookKey,
    spreadsheetName: row.spreadsheetName,
    sheetName: row.sheetName,
    rowNumber: row.rowNumber,
    root: row.root,
    appt: row.appt,
    so: row.so,
    taskId: row.taskId,
    customerName: row.customerName,
    email: row.email,
    brand: row.brand,
    visitDate: row.visitDate,
    status: row.status,
    paymentId: row.paymentId,
    docNumber: row.docNumber,
    matchedField: row.matchedField,
    matchedValue: row.matchedValue,
    reason: row.reason
  };
}

function swPublicTestDataCleanupKeysetCounts_(keyset) {
  return {
    roots: Object.keys(keyset.roots || {}).length,
    appts: Object.keys(keyset.appts || {}).length,
    uids: Object.keys(keyset.uids || {}).length,
    taskIds: Object.keys(keyset.taskIds || {}).length,
    sos: Object.keys(keyset.sos || {}).length,
    emails: Object.keys(keyset.emails || {}).length,
    phones: Object.keys(keyset.phones || {}).length
  };
}

function swLogTestDataCleanupPlan_(plan, label) {
  try {
    Logger.log(label + ' ' + JSON.stringify(swPublicTestDataCleanupResult_(plan)));
  } catch (_) {}
}

function swRecordTestDataCleanupResultFailures_(plan, result, type) {
  if (!swTestDataCleanupResultHasFailure_(result)) return;
  plan.errors.push({
    type: type || 'postDelete',
    message: 'Post-delete ' + (type || 'operation') + ' reported a failure.',
    result: result
  });
}

function swTestDataCleanupResultHasFailure_(value) {
  if (!value || typeof value !== 'object') return false;
  if (value.ok === false) return true;
  if (Array.isArray(value)) {
    for (var i = 0; i < value.length; i++) {
      if (swTestDataCleanupResultHasFailure_(value[i])) return true;
    }
    return false;
  }
  var keys = Object.keys(value);
  for (var k = 0; k < keys.length; k++) {
    if (swTestDataCleanupResultHasFailure_(value[keys[k]])) return true;
  }
  return false;
}

function swTestDataCleanupFingerprint_(plan) {
  var rows = (plan.candidates || []).map(function (row) {
    return [
      row.workbookKey,
      row.spreadsheetId,
      row.sheetId,
      row.rowNumber,
      row.root,
      row.appt,
      row.uid,
      row.taskId,
      row.so,
      swTestDataCleanupNorm_(row.customerName),
      swTestDataCleanupNormEmail_(row.email),
      row.paymentId,
      row.docNumber
    ].join('|');
  });
  return swTestDataCleanupHash_(rows.join('\n'));
}

function swTestDataCleanupRowFingerprint_(rec) {
  rec = rec || {};
  return swTestDataCleanupHash_([
    rec.root || '',
    rec.appt || '',
    rec.uid || '',
    rec.taskId || '',
    rec.so || '',
    swTestDataCleanupNorm_(rec.customerName || ''),
    swTestDataCleanupNormEmail_(rec.email || ''),
    rec.paymentId || '',
    rec.docNumber || ''
  ].join('|'));
}

function swTestDataCleanupHash_(text) {
  text = String(text || '');
  try {
    var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
    return bytes.map(function (b) {
      var v = b < 0 ? b + 256 : b;
      return ('0' + v.toString(16)).slice(-2);
    }).join('');
  } catch (_) {}

  var hash = 0;
  for (var i = 0; i < text.length; i++) {
    hash = ((hash << 5) - hash) + text.charCodeAt(i);
    hash |= 0;
  }
  return ('00000000' + (hash >>> 0).toString(16)).slice(-8);
}

function swCompareTestDataCleanupCandidates_(a, b) {
  if (a.deleteOrder !== b.deleteOrder) return a.deleteOrder - b.deleteOrder;
  if (a.workbookKey !== b.workbookKey) return a.workbookKey < b.workbookKey ? -1 : 1;
  if (a.sheetName !== b.sheetName) return a.sheetName < b.sheetName ? -1 : 1;
  return a.rowNumber - b.rowNumber;
}

function swTestDataCleanupHeaderMap_(headers) {
  var map = {};
  (headers || []).forEach(function (header, idx) {
    var raw = swTestDataCleanupTrim_(header);
    if (!raw) return;
    if (map[raw] == null) map[raw] = idx;
    var key = swTestDataCleanupHeaderKey_(raw);
    if (map[key] == null) map[key] = idx;
  });
  return map;
}

function swTestDataCleanupPick_(map, aliases) {
  for (var i = 0; i < aliases.length; i++) {
    var raw = aliases[i];
    if (map[raw] != null) return map[raw];
    var key = swTestDataCleanupHeaderKey_(raw);
    if (map[key] != null) return map[key];
  }
  return -1;
}

function swTestDataCleanupCell_(row, idx) {
  idx = Number(idx);
  return isFinite(idx) && idx >= 0 ? row[idx] : '';
}

function swTestDataCleanupWorkflowSpreadsheet_() {
  if (typeof swSpreadsheet_ === 'function') return swSpreadsheet_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet and swSpreadsheet_ is unavailable.');
  return ss;
}

function swTestDataCleanupSpreadsheetId_(ss) {
  try { return ss.getId(); } catch (_) {}
  return '';
}

function swTestDataCleanupSpreadsheetName_(ss) {
  try { return ss.getName(); } catch (_) {}
  return '';
}

function swAddTestDataCleanupWarning_(plan, message) {
  if (!plan.warnings) plan.warnings = [];
  if (plan.warnings.indexOf(message) < 0) plan.warnings.push(message);
}

function swUniqueStrings_(values) {
  var seen = {};
  var out = [];
  (values || []).forEach(function (value) {
    value = swTestDataCleanupTrim_(value);
    if (!value || seen[value]) return;
    seen[value] = true;
    out.push(value);
  });
  return out;
}

function swTestDataCleanupHeaderKey_(value) {
  return swTestDataCleanupNorm_(value).replace(/[^a-z0-9]+/g, '');
}

function swTestDataCleanupNorm_(value) {
  return swTestDataCleanupTrim_(value).toLowerCase();
}

function swTestDataCleanupTrim_(value) {
  return String(value == null ? '' : value).trim();
}

function swTestDataCleanupNormEmail_(value) {
  return swTestDataCleanupTrim_(value).toLowerCase();
}

function swTestDataCleanupNormPhone_(value) {
  return swTestDataCleanupTrim_(value).replace(/\D/g, '');
}

function swTestDataCleanupCleanId_(value) {
  return swTestDataCleanupTrim_(value).replace(/^'/, '');
}

function swTestDataCleanupIso_(date) {
  try {
    if (typeof swIso_ === 'function') return swIso_(date);
  } catch (_) {}
  return date && date.toISOString ? date.toISOString() : String(date || '');
}

function swTestDataCleanupDateMs_(value) {
  if (!value) return 0;
  var d = new Date(value);
  var ms = d.getTime();
  return isNaN(ms) ? 0 : ms;
}
