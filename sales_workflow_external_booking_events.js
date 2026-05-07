/**
 * Provider-aware external booking event queue.
 *
 * v1 is Acuity-only. Calendly stays on its existing standalone queue/webhook
 * project and is intentionally not imported here.
 */

var SW_EXTERNAL_BOOKING_EVENT_BATCH_SIZE_ = 20;
var SW_EXTERNAL_BOOKING_EVENT_LOCK_WAIT_MS_ = 30000;
var SW_HPAPP_ACUITY_QUEUE_SPREADSHEET_ID_PROP_ = 'HPAPP_ACUITY_QUEUE_SPREADSHEET_ID';
var SW_EXTERNAL_BOOKING_QUEUE_SPREADSHEET_ID_PROP_ = 'EXTERNAL_BOOKING_QUEUE_SPREADSHEET_ID';
var SW_EXTERNAL_BOOKING_EVENT_DONE_STATUSES_ = {
  DONE: true,
  SKIPPED_DUP: true,
  SKIPPED_CANCELED: true,
  DONE_NO_ROW: true,
  DONE_NO_PRIOR: true,
  DONE_NO_CHANGE: true,
  SKIPPED_UNSUPPORTED: true,
  SKIPPED_IGNORED_ACTION: true
};

function sw_processExternalBookingEvents(options) {
  var redirected = typeof swOrchRedirectLegacyTrigger_ === 'function'
    ? swOrchRedirectLegacyTrigger_('sw_processExternalBookingEvents', options)
    : null;
  if (redirected) return redirected;

  options = options || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(Number(options.lockWaitMs || SW_EXTERNAL_BOOKING_EVENT_LOCK_WAIT_MS_))) {
    return { ok: true, skipped: true, reason: 'LOCK_BUSY' };
  }

  try {
    var sh = swEnsureExternalBookingEventsSheet_();
    var pending = swExternalBookingPendingRows_(sh, options);
    var result = {
      ok: true,
      processed: 0,
      submitted: 0,
      rescheduled: 0,
      edited: 0,
      canceled: 0,
      skipped: 0,
      errors: 0,
      formSubmitted: 0,
      statuses: {},
      checkedAt: new Date().toISOString()
    };

    pending.forEach(function (event) {
      var attempts = Number(event.attempts || 0) + 1;
      var rowResult;
      try {
        rowResult = swProcessExternalBookingEvent_(event);
        var status = rowResult.queueStatus || 'DONE';
        swExternalBookingUpdateRow_(sh, event.H, event.rowNumber, {
          Status: status,
          Attempts: attempts,
          ProcessedAt: new Date(),
          ResolvedUID: rowResult.resolvedUid || '',
          MasterRow: rowResult.masterRow || '',
          ResultJSON: swExternalBookingStringify_({
            action: event.action,
            providerAppointmentId: event.providerAppointmentId,
            outcome: rowResult.outcome || status,
            detailId: rowResult.detailId || '',
            submitted: Number(rowResult.submitted || 0),
            rescheduled: Number(rowResult.rescheduled || 0),
            edited: Number(rowResult.edited || 0),
            canceled: Number(rowResult.canceled || 0)
          }),
          Error: ''
        });
        result.processed++;
        result.submitted += Number(rowResult.submitted || 0);
        result.rescheduled += Number(rowResult.rescheduled || 0);
        result.edited += Number(rowResult.edited || 0);
        result.canceled += Number(rowResult.canceled || 0);
        result.formSubmitted += Number(rowResult.formSubmitted || 0);
        if (/^SKIPPED/.test(status) || /^DONE_NO/.test(status)) result.skipped++;
        result.statuses[status] = Number(result.statuses[status] || 0) + 1;
      } catch (err) {
        var gaveUp = attempts >= 5;
        var errorStatus = gaveUp ? 'ERROR_GAVE_UP' : 'RETRY';
        swExternalBookingUpdateRow_(sh, event.H, event.rowNumber, {
          Status: errorStatus,
          Attempts: attempts,
          ProcessedAt: new Date(),
          Error: swExternalBookingError_(err)
        });
        result.processed++;
        result.errors++;
        result.statuses[errorStatus] = Number(result.statuses[errorStatus] || 0) + 1;
      }
    });

    result.ok = result.errors === 0;
    return result;
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function swEnsureExternalBookingEventsSheet_(ss) {
  var queueSs = swExternalBookingQueueSpreadsheet_(ss);
  return swEnsureSheet_(queueSs, SW_SHEETS.EXTERNAL_BOOKING_EVENTS, SW_EXTERNAL_BOOKING_EVENT_HEADERS);
}

function swExternalBookingQueueSpreadsheet_(fallbackSs) {
  var id = '';
  try {
    var props = PropertiesService.getScriptProperties();
    id = swExternalBookingTrim_(props.getProperty(SW_HPAPP_ACUITY_QUEUE_SPREADSHEET_ID_PROP_) ||
      props.getProperty(SW_EXTERNAL_BOOKING_QUEUE_SPREADSHEET_ID_PROP_) || '');
  } catch (_) {}
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (err) {
      throw new Error('Unable to open HPAPP Acuity queue spreadsheet: ' + swExternalBookingError_(err));
    }
  }
  return fallbackSs || swSpreadsheet_();
}

function swExternalBookingPendingRows_(sh, options) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var values = sh.getRange(1, 1, lastRow, Math.max(sh.getLastColumn(), SW_EXTERNAL_BOOKING_EVENT_HEADERS.length)).getValues();
  var headers = values[0] || [];
  var H = swExternalBookingHeaderMap_(headers);
  var maxRows = Math.max(1, Number(options.maxRows || SW_EXTERNAL_BOOKING_EVENT_BATCH_SIZE_) || SW_EXTERNAL_BOOKING_EVENT_BATCH_SIZE_);
  var onlyTestRunId = swExternalBookingTrim_(options.testRunId || '');
  var out = [];

  for (var r = 1; r < values.length; r++) {
    if (out.length >= maxRows) break;
    var row = values[r];
    var provider = swExternalBookingTrim_(row[H.Provider]).toLowerCase();
    var status = swExternalBookingTrim_(row[H.Status] || 'PENDING').toUpperCase();
    var testRunId = swExternalBookingTrim_(row[H.TestRunID] || '');
    if (provider !== 'acuity') continue;
    if (status !== 'PENDING' && status !== 'RETRY') continue;
    if (onlyTestRunId && testRunId !== onlyTestRunId) continue;
    out.push(swExternalBookingEventFromRow_(row, H, r + 1));
  }
  return out;
}

function swExternalBookingHeaderMap_(headers) {
  var H = {};
  headers.forEach(function (h, i) {
    var key = swExternalBookingTrim_(h);
    if (key && H[key] === undefined) H[key] = i;
  });
  SW_EXTERNAL_BOOKING_EVENT_HEADERS.forEach(function (h) {
    if (H[h] === undefined) throw new Error('Missing _ExternalBookingEvents header: ' + h);
  });
  return H;
}

function swExternalBookingEventFromRow_(row, H, rowNumber) {
  var rawPayloadJson = swExternalBookingTrim_(row[H.RawPayloadJSON] || '');
  var rawPayload = swExternalBookingParseJson_(rawPayloadJson, {});
  return {
    H: H,
    rowNumber: rowNumber,
    receivedAt: row[H.ReceivedAt],
    provider: swExternalBookingTrim_(row[H.Provider]).toLowerCase(),
    action: swExternalBookingNormalizeAction_(row[H.Action]),
    providerAppointmentId: swExternalBookingTrim_(row[H.ProviderAppointmentID]),
    calendarId: swExternalBookingTrim_(row[H.CalendarID]),
    appointmentTypeId: swExternalBookingTrim_(row[H.AppointmentTypeID]),
    rawPayloadJson: rawPayloadJson,
    rawPayload: rawPayload,
    signatureVerified: swExternalBookingTruthy_(row[H.SignatureVerified]),
    status: swExternalBookingTrim_(row[H.Status]),
    attempts: Number(row[H.Attempts] || 0),
    testRunId: swExternalBookingTrim_(row[H.TestRunID])
  };
}

function swProcessExternalBookingEvent_(event) {
  if (event.provider !== 'acuity') {
    return { queueStatus: 'SKIPPED_UNSUPPORTED', outcome: 'unsupportedProvider' };
  }
  if (!event.signatureVerified) {
    throw new Error('External booking event was not signature verified.');
  }
  if (!event.providerAppointmentId) {
    throw new Error('Missing ProviderAppointmentID.');
  }
  if (event.action === 'changed') {
    return { queueStatus: 'SKIPPED_IGNORED_ACTION', outcome: 'ignoredChangedEvent' };
  }
  if (['scheduled', 'rescheduled', 'canceled'].indexOf(event.action) < 0) {
    return { queueStatus: 'SKIPPED_IGNORED_ACTION', outcome: 'ignoredAction:' + event.action };
  }

  var SP = PropertiesService.getScriptProperties();
  var formId = SP.getProperty('FORM_ID');
  if (!formId) throw new Error('Missing Script Property: FORM_ID');

  if (event.action === 'canceled') {
    return swProcessAcuityCanceledExternalBooking_(event);
  }

  var userId = SP.getProperty('ACUITY_USER_ID');
  var apiKey = SP.getProperty('ACUITY_API_KEY');
  var appt = swExternalBookingAcuityDetail_(event, userId, apiKey);
  if (!appt || !appt.id) appt = Object.assign({}, appt || {}, { id: event.providerAppointmentId });

  if (event.action === 'scheduled') {
    return swProcessAcuityScheduledExternalBooking_(event, appt, formId);
  }
  return swProcessAcuityRescheduledExternalBooking_(event, appt, formId);
}

function swProcessAcuityScheduledExternalBooking_(event, appt, formId) {
  var uid = swExternalBookingTrim_(appt.id || event.providerAppointmentId);
  if (swAcuityExternalCanceled_(uid)) {
    return { queueStatus: 'SKIPPED_CANCELED', outcome: 'alreadyCanceled', resolvedUid: uid, detailId: uid };
  }
  var existing = swFindExternalBookingUid_(uid);
  if (existing.found || swAcuityExternalDone_(uid)) {
    return {
      queueStatus: 'SKIPPED_DUP',
      outcome: 'alreadyHandled',
      resolvedUid: uid,
      masterRow: existing.masterRow || '',
      detailId: uid
    };
  }

  var fieldMap = acuityToFormFieldMap_(appt);
  acuitySubmitToForm_(formId, fieldMap);
  swAcuityMarkExternalDone_(uid);
  return {
    queueStatus: 'DONE',
    outcome: 'scheduledSubmitted',
    submitted: 1,
    formSubmitted: 1,
    resolvedUid: uid,
    detailId: uid
  };
}

function swProcessAcuityRescheduledExternalBooking_(event, appt, formId) {
  var baseUid = swExternalBookingTrim_(appt.id || event.providerAppointmentId);
  var newUid = acuityStableRescheduleUid_(appt);
  var existingNew = newUid ? swFindExternalBookingUid_(newUid) : { found: false };
  if ((existingNew.found || (newUid && swAcuityExternalDone_(newUid))) && newUid) {
    return {
      queueStatus: 'SKIPPED_DUP',
      outcome: 'rescheduleAlreadyHandled',
      resolvedUid: newUid,
      masterRow: existingNew.masterRow || '',
      detailId: baseUid
    };
  }

  var existingBase = swFindExternalBookingUid_(baseUid);
  if (!existingBase.found && !swAcuityExternalDone_(baseUid)) {
    var fieldMap = acuityToFormFieldMap_(appt);
    acuitySubmitToForm_(formId, fieldMap);
    swAcuityMarkExternalDone_(baseUid);
    return {
      queueStatus: 'DONE_NO_PRIOR',
      outcome: 'rescheduleNoPriorSubmittedAsScheduled',
      submitted: 1,
      formSubmitted: 1,
      resolvedUid: baseUid,
      detailId: baseUid
    };
  }

  var outcome = acuityHandleExisting_(appt, formId);
  if (outcome === 'rescheduled') {
    swAcuityMarkExternalDone_(baseUid);
    if (newUid) swAcuityMarkExternalDone_(newUid);
    return {
      queueStatus: 'DONE',
      outcome: 'rescheduled',
      rescheduled: 1,
      formSubmitted: 1,
      resolvedUid: newUid || baseUid,
      masterRow: existingBase.masterRow || '',
      detailId: baseUid
    };
  }
  if (outcome === 'edited') {
    swAcuityMarkExternalDone_(baseUid);
    return {
      queueStatus: 'DONE',
      outcome: 'editedExisting',
      edited: 1,
      resolvedUid: baseUid,
      masterRow: existingBase.masterRow || '',
      detailId: baseUid
    };
  }
  return {
    queueStatus: 'DONE_NO_CHANGE',
    outcome: outcome || 'unchanged',
    resolvedUid: existingNew.found ? newUid : baseUid,
    masterRow: existingNew.masterRow || existingBase.masterRow || '',
    detailId: baseUid
  };
}

function swProcessAcuityCanceledExternalBooking_(event) {
  var uid = swExternalBookingTrim_(event.providerAppointmentId);
  var existing = swFindLatestAcuityMasterRow_(uid);
  var canceled = acuityCancelOnMaster_(uid);
  swAcuityMarkExternalCanceled_(uid);
  swAcuityMarkExternalDone_(uid);
  if (canceled) {
    return {
      queueStatus: 'DONE',
      outcome: 'canceled',
      canceled: 1,
      resolvedUid: uid,
      masterRow: existing.masterRow || '',
      detailId: uid
    };
  }
  return {
    queueStatus: 'DONE_NO_ROW',
    outcome: 'cancelNoActiveRow',
    resolvedUid: uid,
    detailId: uid
  };
}

function swExternalBookingAcuityDetail_(event, userId, apiKey) {
  var payload = event.rawPayload || {};
  var detail = payload.mockAppointment || payload.appointment || payload.appointmentDetail || payload.detail || null;
  if (detail) {
    if (!detail.id) detail.id = event.providerAppointmentId;
    return detail;
  }
  if (!userId || !apiKey) {
    throw new Error('Missing Script Properties: ACUITY_USER_ID / ACUITY_API_KEY');
  }
  return acuityFetchAppointmentDetail_(userId, apiKey, { id: event.providerAppointmentId });
}

function swAcuityMarkExternalDone_(uid) {
  uid = swExternalBookingTrim_(uid);
  if (!uid) return;
  try { PropertiesService.getScriptProperties().setProperty('ACUITY:DONE:' + uid, '1'); } catch (_) {}
  try { CacheService.getScriptCache().remove('MASTER_UIDS_CACHE'); } catch (_) {}
}

function swAcuityExternalDone_(uid) {
  uid = swExternalBookingTrim_(uid);
  if (!uid) return false;
  try { return PropertiesService.getScriptProperties().getProperty('ACUITY:DONE:' + uid) === '1'; } catch (_) {}
  return false;
}

function swAcuityMarkExternalCanceled_(uid) {
  uid = swExternalBookingTrim_(uid);
  if (!uid) return;
  try { PropertiesService.getScriptProperties().setProperty('ACUITY:CANCELED:' + uid, '1'); } catch (_) {}
}

function swAcuityExternalCanceled_(uid) {
  uid = swExternalBookingTrim_(uid);
  if (!uid) return false;
  try { return PropertiesService.getScriptProperties().getProperty('ACUITY:CANCELED:' + uid) === '1'; } catch (_) {}
  return false;
}

function swFindExternalBookingUid_(uid) {
  uid = swExternalBookingTrim_(uid);
  if (!uid) return { found: false };
  var master = swFindMasterAppointmentUid_(uid);
  if (master.found) return master;
  if (swFindFormInboxUid_(uid)) return { found: true, source: '02_Form_Inbox' };
  return { found: false };
}

function swFindMasterAppointmentUid_(uid) {
  var ss = swSpreadsheet_();
  var sh = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!sh || sh.getLastRow() < 2) return { found: false };
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var H = {};
  hdr.forEach(function (h, i) { if (h) H[String(h).trim()] = i + 1; });
  if (!H['CalendlyEventUID']) return { found: false };
  var vals = sh.getRange(2, H['CalendlyEventUID'], sh.getLastRow() - 1, 1).getValues();
  var found = null;
  for (var i = 0; i < vals.length; i++) {
    if (swExternalBookingTrim_(vals[i][0]) === uid) found = { found: true, source: SW_SHEETS.MASTER, masterRow: i + 2 };
  }
  return found || { found: false };
}

function swFindLatestAcuityMasterRow_(baseUid) {
  baseUid = swExternalBookingTrim_(baseUid);
  var ss = swSpreadsheet_();
  var sh = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!baseUid || !sh || sh.getLastRow() < 2) return { found: false };
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var H = {};
  hdr.forEach(function (h, i) { if (h) H[String(h).trim()] = i + 1; });
  if (!H['CalendlyEventUID']) return { found: false };
  var lastRow = sh.getLastRow();
  var vals = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  var candidate = null;
  vals.forEach(function (row, i) {
    var uid = swExternalBookingTrim_(row[H['CalendlyEventUID'] - 1]);
    if (uid !== baseUid && uid.indexOf(baseUid + '_R') !== 0) return;
    var status = H['Status'] ? swExternalBookingTrim_(row[H['Status'] - 1]) : '';
    if (/canceled|rescheduled/i.test(status)) return;
    candidate = { found: true, source: SW_SHEETS.MASTER, masterRow: i + 2, uid: uid };
  });
  return candidate || { found: false };
}

function swFindFormInboxUid_(uid) {
  var ss = swSpreadsheet_();
  var sh = ss.getSheetByName('02_Form_Inbox');
  if (!sh || sh.getLastRow() < 2) return false;
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var idx = swExternalBookingHeaderIndex_(hdr, ['Admin: Calendly Event UID', 'CalendlyEventUID', 'Calendly Event UID', 'Acuity ID']);
  if (idx < 0) return false;
  var vals = sh.getRange(2, idx + 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (swExternalBookingTrim_(vals[i][0]) === uid) return true;
  }
  return false;
}

function swExternalBookingUpdateRow_(sh, H, rowNumber, values) {
  Object.keys(values || {}).forEach(function (name) {
    if (H[name] === undefined) return;
    sh.getRange(rowNumber, H[name] + 1).setValue(values[name]);
  });
}

function swExternalBookingNormalizeAction_(value) {
  var s = swExternalBookingTrim_(value).toLowerCase();
  s = s.replace(/^appointment[\._-]?/, '');
  if (s === 'cancelled') return 'canceled';
  if (s === 'reschedule') return 'rescheduled';
  if (s === 'schedule') return 'scheduled';
  return s;
}

function swExternalBookingHeaderIndex_(headers, names) {
  var low = (headers || []).map(function (h) { return swExternalBookingTrim_(h).toLowerCase(); });
  for (var i = 0; i < names.length; i++) {
    var idx = low.indexOf(swExternalBookingTrim_(names[i]).toLowerCase());
    if (idx >= 0) return idx;
  }
  return -1;
}

function swExternalBookingParseJson_(raw, fallback) {
  try {
    if (!raw) return fallback;
    return JSON.parse(String(raw));
  } catch (_) {
    return fallback;
  }
}

function swExternalBookingStringify_(value) {
  try { return JSON.stringify(value || {}); } catch (_) { return '{}'; }
}

function swExternalBookingTruthy_(value) {
  return /^(true|yes|y|1)$/i.test(swExternalBookingTrim_(value));
}

function swExternalBookingTrim_(value) {
  return String(value === null || value === undefined ? '' : value).trim();
}

function swExternalBookingError_(err) {
  return err && (err.stack || err.message) ? String(err.stack || err.message) : String(err || '');
}

function sw_testInjectExternalBookingEvent(options) {
  options = options || {};
  var sh = swEnsureExternalBookingEventsSheet_();
  var now = new Date();
  var action = swExternalBookingNormalizeAction_(options.action || 'scheduled');
  var id = swExternalBookingTrim_(options.providerAppointmentId || options.id || ('TEST' + now.getTime()));
  var testRunId = swExternalBookingTrim_(options.testRunId || ('ACUITY_EXTERNAL_TEST_' + now.getTime()));
  var rawPayload = options.rawPayload || {};
  if (options.mockAppointment) rawPayload.mockAppointment = options.mockAppointment;
  if (!rawPayload.testRunId) rawPayload.testRunId = testRunId;
  if (!rawPayload.providerAppointmentId) rawPayload.providerAppointmentId = id;

  var row = [
    now,
    'acuity',
    action,
    id,
    swExternalBookingTrim_(options.calendarId || (rawPayload.calendarID || rawPayload.calendarId || '')),
    swExternalBookingTrim_(options.appointmentTypeId || (rawPayload.appointmentTypeID || rawPayload.appointmentTypeId || '')),
    swExternalBookingStringify_(rawPayload),
    options.signatureVerified === false ? 'FALSE' : 'TRUE',
    'PENDING',
    0,
    '',
    '',
    '',
    '',
    '',
    testRunId
  ];
  sh.appendRow(row);
  return { ok: true, sheet: SW_SHEETS.EXTERNAL_BOOKING_EVENTS, row: sh.getLastRow(), testRunId: testRunId, id: id, action: action };
}

function sw_configureHpAppAcuityQueue(options) {
  options = options || {};
  var spreadsheetId = swExternalBookingTrim_(options.spreadsheetId || options.queueSpreadsheetId || '');
  if (!spreadsheetId) throw new Error('Missing spreadsheetId.');
  var props = PropertiesService.getScriptProperties();
  props.setProperty(SW_HPAPP_ACUITY_QUEUE_SPREADSHEET_ID_PROP_, spreadsheetId);
  props.setProperty(SW_EXTERNAL_BOOKING_QUEUE_SPREADSHEET_ID_PROP_, spreadsheetId);
  if (options.acuityUserId) props.setProperty('ACUITY_USER_ID', swExternalBookingTrim_(options.acuityUserId));
  if (options.acuityApiKey) props.setProperty('ACUITY_API_KEY', swExternalBookingTrim_(options.acuityApiKey));
  var sh = swEnsureExternalBookingEventsSheet_();
  swStyleSheet_(sh);
  return {
    ok: true,
    spreadsheetId: spreadsheetId,
    sheetName: sh.getName(),
    headers: sh.getRange(1, 1, 1, SW_EXTERNAL_BOOKING_EVENT_HEADERS.length).getDisplayValues()[0]
  };
}

function sw_getHpAppAcuityQueueConfig() {
  var props = PropertiesService.getScriptProperties();
  var spreadsheetId = swExternalBookingTrim_(props.getProperty(SW_HPAPP_ACUITY_QUEUE_SPREADSHEET_ID_PROP_) ||
    props.getProperty(SW_EXTERNAL_BOOKING_QUEUE_SPREADSHEET_ID_PROP_) || '');
  var ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : swSpreadsheet_();
  var sh = swEnsureExternalBookingEventsSheet_(ss);
  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    sheetName: sh.getName(),
    rows: Math.max(0, sh.getLastRow() - 1),
    headers: sh.getRange(1, 1, 1, SW_EXTERNAL_BOOKING_EVENT_HEADERS.length).getDisplayValues()[0],
    usingExternalSpreadsheet: !!spreadsheetId
  };
}

function sw_clearExternalBookingTestFlags(options) {
  options = options || {};
  var ids = options.ids || [];
  if (options.id) ids.push(options.id);
  var deleted = [];
  var props = PropertiesService.getScriptProperties();
  ids.forEach(function (id) {
    id = swExternalBookingTrim_(id);
    if (!id) return;
    ['ACUITY:DONE:' + id, 'ACUITY:CANCELED:' + id].forEach(function (key) {
      props.deleteProperty(key);
      deleted.push(key);
    });
  });
  return { ok: true, deleted: deleted };
}
