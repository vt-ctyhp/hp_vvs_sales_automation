/**
 * Phase 1 duplicate audit.
 *
 * Read-only: scans 00_Master Appointments and logs duplicate appointment risks.
 * It does not write to the spreadsheet.
 */
function duplicate_phase1_auditOnly() {
  var ss = duplicateAuditSpreadsheet_();
  var sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet: 00_Master Appointments');

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    Logger.log('PHASE1_DUPLICATE_AUDIT ' + JSON.stringify({
      ok: true,
      spreadsheetName: ss.getName(),
      spreadsheetId: ss.getId(),
      masterRows: Math.max(0, lastRow - 1),
      message: 'No master data rows.'
    }));
    return { ok: true, masterRows: 0 };
  }

  var values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headers = values[0].map(function (h) { return String(h || '').trim(); });
  var rows = values.slice(1);
  var H = duplicateAuditHeaderMap_(headers);

  var idx = {
    appt: duplicateAuditPick_(H, ['APPT_ID', 'Appt ID', 'Appointment ID']),
    root: duplicateAuditPick_(H, ['RootApptID', 'Root Appt ID', 'Root Appointment ID']),
    uid: duplicateAuditPick_(H, ['CalendlyEventUID', 'Calendly Event UID', 'Admin: Calendly Event UID', 'Acuity ID']),
    visitNo: duplicateAuditPick_(H, ['Visit #', 'Visit Number', 'Visit No']),
    name: duplicateAuditPick_(H, ['Customer Name', 'Customer', 'Client Name', 'Name']),
    emailLower: duplicateAuditPick_(H, ['EmailLower', 'Email Lower']),
    email: duplicateAuditPick_(H, ['Email', 'Email Address', 'E-mail']),
    phoneNorm: duplicateAuditPick_(H, ['PhoneNorm', 'Phone Norm']),
    phone: duplicateAuditPick_(H, ['Phone', 'Phone Number', 'Mobile', 'Tel']),
    brand: duplicateAuditPick_(H, ['Brand', 'Company']),
    visitDate: duplicateAuditPick_(H, ['Visit Date', 'Appointment Date']),
    visitTime: duplicateAuditPick_(H, ['Visit Time', 'Appointment Time']),
    visitType: duplicateAuditPick_(H, ['Visit Type', 'Appointment Type']),
    status: duplicateAuditPick_(H, ['Status']),
    active: duplicateAuditPick_(H, ['Active?', 'Active', 'Is Active']),
    timestamp: duplicateAuditPick_(H, ['Timestamp', 'Created At', 'Submitted At']),
    rescheduledFrom: duplicateAuditPick_(H, ['RescheduledFromUID', 'Rescheduled From UID']),
    rescheduledTo: duplicateAuditPick_(H, ['RescheduledToUID', 'Rescheduled To UID']),
    canceledAt: duplicateAuditPick_(H, ['CanceledAt', 'Canceled At'])
  };

  var missing = [];
  ['appt', 'uid', 'visitNo', 'name', 'brand', 'visitDate', 'visitTime', 'visitType', 'status', 'active'].forEach(function (k) {
    if (idx[k] < 0) missing.push(k);
  });

  var records = rows.map(function (row, i) {
    var rowIndex = i + 2;
    var email = duplicateAuditNormEmail_(duplicateAuditCell_(row, idx.emailLower) || duplicateAuditCell_(row, idx.email));
    var phone = duplicateAuditNormPhone_(duplicateAuditCell_(row, idx.phoneNorm) || duplicateAuditCell_(row, idx.phone));
    var status = duplicateAuditCell_(row, idx.status);
    var active = duplicateAuditCell_(row, idx.active);
    var rec = {
      rowIndex: rowIndex,
      appt: duplicateAuditTrim_(duplicateAuditCell_(row, idx.appt)),
      root: duplicateAuditTrim_(duplicateAuditCell_(row, idx.root)),
      uid: duplicateAuditTrim_(duplicateAuditCell_(row, idx.uid)),
      visitNo: duplicateAuditTrim_(duplicateAuditCell_(row, idx.visitNo)),
      name: duplicateAuditTrim_(duplicateAuditCell_(row, idx.name)),
      email: email,
      phone: phone,
      brand: duplicateAuditNormKey_(duplicateAuditCell_(row, idx.brand)),
      visitDate: duplicateAuditNormKey_(duplicateAuditCell_(row, idx.visitDate)),
      visitTime: duplicateAuditNormKey_(duplicateAuditCell_(row, idx.visitTime)),
      visitType: duplicateAuditNormKey_(duplicateAuditCell_(row, idx.visitType)),
      status: duplicateAuditTrim_(status),
      activeRaw: duplicateAuditTrim_(active),
      timestamp: duplicateAuditTrim_(duplicateAuditCell_(row, idx.timestamp)),
      rescheduledFrom: duplicateAuditTrim_(duplicateAuditCell_(row, idx.rescheduledFrom)),
      rescheduledTo: duplicateAuditTrim_(duplicateAuditCell_(row, idx.rescheduledTo)),
      canceledAt: duplicateAuditTrim_(duplicateAuditCell_(row, idx.canceledAt)),
      nonBlankCount: duplicateAuditNonBlankCount_(row)
    };
    rec.isCurrent = duplicateAuditIsCurrent_(rec.status, rec.activeRaw);
    rec.score = duplicateAuditCanonicalScore_(rec);
    return rec;
  });

  var byAppt = duplicateAuditGroup_(records, function (r) { return r.appt; });
  var byUid = duplicateAuditGroup_(records, function (r) { return r.uid; });
  var byFingerprint = {};
  records.forEach(function (r) {
    var base = [r.brand, r.visitDate, r.visitTime, r.visitType].join('|');
    if (!r.visitDate || !r.visitTime) return;
    if (r.email) duplicateAuditPush_(byFingerprint, base + '|email:' + r.email, r);
    if (r.phone) duplicateAuditPush_(byFingerprint, base + '|phone:' + r.phone, r);
  });

  var duplicateAppts = duplicateAuditIssues_(byAppt, 'DUPLICATE_APPT_ID');
  var duplicateUids = duplicateAuditIssues_(byUid, 'DUPLICATE_UID');
  var duplicateFingerprints = duplicateAuditIssues_(byFingerprint, 'DUPLICATE_APPOINTMENT_FINGERPRINT');
  var multipleActiveAppts = duplicateAuditIssues_(byAppt, 'MULTIPLE_ACTIVE_APPT_ID', function (group) {
    return group.filter(function (r) { return r.isCurrent; }).length > 1;
  });
  var multipleActiveUids = duplicateAuditIssues_(byUid, 'MULTIPLE_ACTIVE_UID', function (group) {
    return group.filter(function (r) { return r.isCurrent; }).length > 1;
  });
  var blankVisitRows = records.filter(function (r) {
    return !r.visitNo && (r.appt || r.uid || r.name || r.email || r.phone);
  });
  var blankVisitInDuplicateAppts = blankVisitRows.filter(function (r) {
    return r.appt && byAppt[r.appt] && byAppt[r.appt].length > 1;
  });
  var statusConflicts = records.filter(function (r) {
    var s = String(r.status || '').toLowerCase();
    var a = String(r.activeRaw || '').toLowerCase();
    return (a === 'yes' || a === 'true' || a === '1') && /cancel|resched|duplicate|superseded/.test(s);
  });

  var summary = {
    ok: true,
    generatedAt: new Date().toISOString(),
    spreadsheetName: ss.getName(),
    spreadsheetId: ss.getId(),
    masterRows: records.length,
    missingExpectedColumns: missing,
    counts: {
      duplicateApptIdGroups: duplicateAppts.length,
      duplicateApptIdRows: duplicateAuditRowsInIssues_(duplicateAppts),
      duplicateUidGroups: duplicateUids.length,
      duplicateUidRows: duplicateAuditRowsInIssues_(duplicateUids),
      duplicateFingerprintGroups: duplicateFingerprints.length,
      duplicateFingerprintRows: duplicateAuditRowsInIssues_(duplicateFingerprints),
      multipleActiveApptIdGroups: multipleActiveAppts.length,
      multipleActiveUidGroups: multipleActiveUids.length,
      blankVisitNumberRows: blankVisitRows.length,
      blankVisitNumberRowsInsideDuplicateApptId: blankVisitInDuplicateAppts.length,
      activeStatusConflicts: statusConflicts.length
    },
    examples: {
      duplicateApptIds: duplicateAuditExampleIssues_(duplicateAppts, 15),
      duplicateUids: duplicateAuditExampleIssues_(duplicateUids, 15),
      duplicateFingerprints: duplicateAuditExampleIssues_(duplicateFingerprints, 15),
      multipleActiveApptIds: duplicateAuditExampleIssues_(multipleActiveAppts, 15),
      blankVisitNumbers: blankVisitRows.slice(0, 20).map(duplicateAuditBriefRecord_),
      statusConflicts: statusConflicts.slice(0, 20).map(duplicateAuditBriefRecord_)
    }
  };

  Logger.log('PHASE1_DUPLICATE_AUDIT_SUMMARY ' + JSON.stringify(summary.counts));
  Logger.log('PHASE1_DUPLICATE_AUDIT_DETAILS ' + JSON.stringify(summary, null, 2));
  return summary;
}

/**
 * Read-only focused spot check for duplicate Calendly/Acuity UID groups.
 * Run this after duplicate_phase1_auditOnly if the full output is too large.
 */
function duplicate_phase1_spotCheckDuplicateUids() {
  var ss = duplicateAuditSpreadsheet_();
  var sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet: 00_Master Appointments');

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    Logger.log('PHASE1_DUPLICATE_UID_SPOTCHECK ' + JSON.stringify({ ok: true, groups: [] }));
    return { ok: true, groups: [] };
  }

  var values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headers = values[0].map(function (h) { return String(h || '').trim(); });
  var H = duplicateAuditHeaderMap_(headers);
  var idx = {
    uid: duplicateAuditPick_(H, ['CalendlyEventUID', 'Calendly Event UID', 'Admin: Calendly Event UID', 'Acuity ID']),
    appt: duplicateAuditPick_(H, ['APPT_ID', 'Appt ID', 'Appointment ID']),
    root: duplicateAuditPick_(H, ['RootApptID', 'Root Appt ID', 'Root Appointment ID']),
    visitNo: duplicateAuditPick_(H, ['Visit #', 'Visit Number', 'Visit No']),
    name: duplicateAuditPick_(H, ['Customer Name', 'Customer', 'Client Name', 'Name']),
    emailLower: duplicateAuditPick_(H, ['EmailLower', 'Email Lower']),
    email: duplicateAuditPick_(H, ['Email', 'Email Address', 'E-mail']),
    phoneNorm: duplicateAuditPick_(H, ['PhoneNorm', 'Phone Norm']),
    phone: duplicateAuditPick_(H, ['Phone', 'Phone Number', 'Mobile', 'Tel']),
    brand: duplicateAuditPick_(H, ['Brand', 'Company']),
    visitDate: duplicateAuditPick_(H, ['Visit Date', 'Appointment Date']),
    visitTime: duplicateAuditPick_(H, ['Visit Time', 'Appointment Time']),
    visitType: duplicateAuditPick_(H, ['Visit Type', 'Appointment Type']),
    status: duplicateAuditPick_(H, ['Status']),
    active: duplicateAuditPick_(H, ['Active?', 'Active', 'Is Active']),
    timestamp: duplicateAuditPick_(H, ['Timestamp', 'Created At', 'Submitted At']),
    rescheduledFrom: duplicateAuditPick_(H, ['RescheduledFromUID', 'Rescheduled From UID']),
    rescheduledTo: duplicateAuditPick_(H, ['RescheduledToUID', 'Rescheduled To UID']),
    canceledAt: duplicateAuditPick_(H, ['CanceledAt', 'Canceled At'])
  };

  var byUid = {};
  values.slice(1).forEach(function (row, i) {
    var uid = duplicateAuditTrim_(duplicateAuditCell_(row, idx.uid));
    if (!uid) return;
    var rec = {
      row: i + 2,
      appt: duplicateAuditTrim_(duplicateAuditCell_(row, idx.appt)),
      root: duplicateAuditTrim_(duplicateAuditCell_(row, idx.root)),
      visitNo: duplicateAuditTrim_(duplicateAuditCell_(row, idx.visitNo)),
      name: duplicateAuditTrim_(duplicateAuditCell_(row, idx.name)),
      email: duplicateAuditNormEmail_(duplicateAuditCell_(row, idx.emailLower) || duplicateAuditCell_(row, idx.email)),
      phone: duplicateAuditNormPhone_(duplicateAuditCell_(row, idx.phoneNorm) || duplicateAuditCell_(row, idx.phone)),
      brand: duplicateAuditTrim_(duplicateAuditCell_(row, idx.brand)),
      visitDate: duplicateAuditTrim_(duplicateAuditCell_(row, idx.visitDate)),
      visitTime: duplicateAuditTrim_(duplicateAuditCell_(row, idx.visitTime)),
      visitType: duplicateAuditTrim_(duplicateAuditCell_(row, idx.visitType)),
      status: duplicateAuditTrim_(duplicateAuditCell_(row, idx.status)),
      active: duplicateAuditTrim_(duplicateAuditCell_(row, idx.active)),
      timestamp: duplicateAuditTrim_(duplicateAuditCell_(row, idx.timestamp)),
      rescheduledFrom: duplicateAuditTrim_(duplicateAuditCell_(row, idx.rescheduledFrom)),
      rescheduledTo: duplicateAuditTrim_(duplicateAuditCell_(row, idx.rescheduledTo)),
      canceledAt: duplicateAuditTrim_(duplicateAuditCell_(row, idx.canceledAt))
    };
    rec.isCurrent = duplicateAuditIsCurrent_(rec.status, rec.active);
    duplicateAuditPush_(byUid, uid, rec);
  });

  var groups = Object.keys(byUid).filter(function (uid) {
    return byUid[uid].length > 1;
  }).sort().map(function (uid) {
    var rows = byUid[uid];
    return {
      uid: uid,
      rowCount: rows.length,
      activeRows: rows.filter(function (r) { return r.isCurrent; }).map(function (r) { return r.row; }),
      rows: rows
    };
  });

  Logger.log('PHASE1_DUPLICATE_UID_SPOTCHECK_SUMMARY ' + JSON.stringify({
    duplicateUidGroups: groups.length,
    duplicateUidRows: groups.reduce(function (sum, g) { return sum + g.rowCount; }, 0),
    groupsWithMultipleActiveRows: groups.filter(function (g) { return g.activeRows.length > 1; }).length
  }));
  groups.forEach(function (g, i) {
    Logger.log('PHASE1_DUPLICATE_UID_GROUP_' + (i + 1) + ' ' + JSON.stringify(g, null, 2));
  });
  return { ok: true, groups: groups };
}

/**
 * Read-only cause classifier for remaining duplicate/bad-data patterns.
 *
 * This looks beyond duplicate UID groups. It compares Master against
 * 02_Form_Inbox and _IntakeQueue so we can infer whether bad rows came from:
 * - upstream duplicate submissions
 * - resolver/queue replay
 * - missing/different UID fallback duplicates
 * - expected reschedule history
 * - status/active/visit-number hygiene issues
 */
function duplicate_phase1_identifyRemainingCauses() {
  var ss = duplicateAuditSpreadsheet_();
  var master = duplicateCauseReadAppointmentSheet_(ss, '00_Master Appointments', 'MASTER');
  var inbox = duplicateCauseReadAppointmentSheet_(ss, '02_Form_Inbox', 'FORM_INBOX');
  var queue = duplicateCauseReadQueue_(ss);

  var masterByUid = duplicateCauseGroup_(master.records, function (r) { return r.uid; });
  var inboxByUid = duplicateCauseGroup_(inbox.records, function (r) { return r.uid; });
  var masterByFingerprint = duplicateCauseGroupFingerprints_(master.records);
  var inboxByFingerprint = duplicateCauseGroupFingerprints_(inbox.records);
  var masterByAppt = duplicateCauseGroup_(master.records, function (r) { return r.appt; });

  var buckets = {
    UID_REPLAY_MULTIPLE_ACTIVE: [],
    UID_DUPLICATE_STATUS_CONFLICT: [],
    UID_REUSED_FOR_RESCHEDULE_HISTORY: [],
    FINGERPRINT_DUPLICATE_MULTIPLE_ACTIVE: [],
    FINGERPRINT_DUPLICATE_WITH_BLANK_UID: [],
    FINGERPRINT_DUPLICATE_WITH_DIFFERENT_UIDS: [],
    CURRENT_ROW_BLANK_VISIT_NUMBER: [],
    BLANK_VISIT_NUMBER_PAIRED_WITH_FILLED_ROW: [],
    ACTIVE_STATUS_CONFLICT: [],
    RESCHEDULE_HISTORY_MISSING_LINKS: [],
    APPT_ID_IS_ROOT_HISTORY_NOT_DUPLICATE: [],
    APPT_ID_SHARED_ACROSS_DIFFERENT_CONTACTS: [],
    INTAKE_QUEUE_REPLAY_RISK: []
  };

  Object.keys(masterByUid).forEach(function (uid) {
    var group = masterByUid[uid];
    if (!uid || group.length < 2) return;

    var activeRows = group.filter(function (r) { return r.isCurrent; });
    var fingerprints = duplicateCauseUnique_(group.map(function (r) { return duplicateCausePrimaryFingerprint_(r); }));
    var statuses = duplicateCauseUnique_(group.map(function (r) { return r.statusNorm || '(blank)'; }));
    var issue = duplicateCauseIssue_(uid, group, {
      sourceSystem: duplicateCauseUidSource_(uid),
      masterRowCount: group.length,
      activeRowCount: activeRows.length,
      inboxUidCount: (inboxByUid[uid] || []).length,
      fingerprints: fingerprints,
      statuses: statuses,
      likelyOrigin: duplicateCauseLikelyOrigin_(group.length, (inboxByUid[uid] || []).length)
    });

    if (activeRows.length > 1 && fingerprints.length <= 1) {
      buckets.UID_REPLAY_MULTIPLE_ACTIVE.push(issue);
    } else if (fingerprints.length > 1) {
      buckets.UID_REUSED_FOR_RESCHEDULE_HISTORY.push(issue);
    } else {
      buckets.UID_DUPLICATE_STATUS_CONFLICT.push(issue);
    }
  });

  Object.keys(masterByFingerprint).forEach(function (key) {
    var group = masterByFingerprint[key];
    if (!key || group.length < 2) return;

    var activeRows = group.filter(function (r) { return r.isCurrent; });
    var uids = duplicateCauseUnique_(group.map(function (r) { return r.uid || '(blank)'; }));
    var nonBlankUids = uids.filter(function (u) { return u !== '(blank)'; });
    var blankUidCount = group.filter(function (r) { return !r.uid; }).length;
    var sameUidOnly = nonBlankUids.length === 1 && blankUidCount === 0;
    var issue = duplicateCauseIssue_(key, group, {
      masterRowCount: group.length,
      activeRowCount: activeRows.length,
      inboxFingerprintCount: (inboxByFingerprint[key] || []).length,
      uniqueUids: uids,
      blankUidCount: blankUidCount,
      likelyOrigin: duplicateCauseLikelyOrigin_(group.length, (inboxByFingerprint[key] || []).length)
    });

    if (sameUidOnly) return; // already represented by UID buckets.
    if (activeRows.length > 1) buckets.FINGERPRINT_DUPLICATE_MULTIPLE_ACTIVE.push(issue);
    if (blankUidCount > 0) buckets.FINGERPRINT_DUPLICATE_WITH_BLANK_UID.push(issue);
    if (nonBlankUids.length > 1) buckets.FINGERPRINT_DUPLICATE_WITH_DIFFERENT_UIDS.push(issue);
  });

  master.records.forEach(function (r) {
    if (r.isCurrent && !r.visitNo && (r.appt || r.uid || r.name || r.email || r.phone)) {
      buckets.CURRENT_ROW_BLANK_VISIT_NUMBER.push(duplicateCauseBrief_(r));
    }

    if (r.isCurrent && /cancel|resched|duplicate|superseded/.test(r.statusNorm)) {
      buckets.ACTIVE_STATUS_CONFLICT.push(duplicateCauseBrief_(r));
    }

    if (/resched/.test(r.statusNorm) && !r.isCurrent && !r.rescheduledTo && !r.rescheduledFrom) {
      buckets.RESCHEDULE_HISTORY_MISSING_LINKS.push(duplicateCauseBrief_(r));
    }
  });

  Object.keys(masterByFingerprint).forEach(function (key) {
    var group = masterByFingerprint[key];
    if (!key || group.length < 2) return;
    var hasBlank = group.some(function (r) { return !r.visitNo; });
    var hasFilled = group.some(function (r) { return !!r.visitNo; });
    if (hasBlank && hasFilled) {
      buckets.BLANK_VISIT_NUMBER_PAIRED_WITH_FILLED_ROW.push(duplicateCauseIssue_(key, group, {
        blankRows: group.filter(function (r) { return !r.visitNo; }).map(function (r) { return r.row; }),
        filledRows: group.filter(function (r) { return !!r.visitNo; }).map(function (r) { return r.row; })
      }));
    }
  });

  Object.keys(masterByAppt).forEach(function (appt) {
    var group = masterByAppt[appt];
    if (!appt || group.length < 2) return;
    var contacts = duplicateCauseUnique_(group.map(function (r) { return duplicateCauseContactKey_(r); }).filter(Boolean));
    var occurrences = duplicateCauseUnique_(group.map(function (r) {
      return [r.visitDateNorm, r.visitTimeNorm, r.visitTypeNorm, r.visitNo].join('|');
    }).filter(function (v) { return v.replace(/\|/g, ''); }));
    var issue = duplicateCauseIssue_(appt, group, {
      contactCount: contacts.length,
      occurrenceCount: occurrences.length,
      contacts: contacts.slice(0, 5),
      note: contacts.length <= 1 && occurrences.length > 1
        ? 'Repeated APPT_ID appears to be customer/root history, not exact duplicate rows.'
        : 'Repeated APPT_ID crosses contact keys; review before using APPT_ID as identity.'
    });
    if (contacts.length <= 1 && occurrences.length > 1) {
      buckets.APPT_ID_IS_ROOT_HISTORY_NOT_DUPLICATE.push(issue);
    } else if (contacts.length > 1) {
      buckets.APPT_ID_SHARED_ACROSS_DIFFERENT_CONTACTS.push(issue);
    }
  });

  queue.risks.forEach(function (risk) {
    buckets.INTAKE_QUEUE_REPLAY_RISK.push(risk);
  });

  var summary = {
    ok: true,
    generatedAt: new Date().toISOString(),
    spreadsheetName: ss.getName(),
    spreadsheetId: ss.getId(),
    rowCounts: {
      master: master.records.length,
      formInbox: inbox.records.length,
      intakeQueue: queue.records.length
    },
    missingSheets: []
      .concat(master.missing ? ['00_Master Appointments'] : [])
      .concat(inbox.missing ? ['02_Form_Inbox'] : [])
      .concat(queue.missing ? ['_IntakeQueue'] : []),
    bucketCounts: duplicateCauseBucketCounts_(buckets)
  };

  Logger.log('PHASE1_CAUSE_AUDIT_SUMMARY ' + JSON.stringify(summary));
  Object.keys(buckets).forEach(function (name) {
    var items = buckets[name];
    Logger.log('PHASE1_CAUSE_AUDIT_BUCKET_' + name + ' ' + JSON.stringify({
      count: items.length,
      examples: items.slice(0, 20)
    }, null, 2));
  });
  return { summary: summary, buckets: buckets };
}

/**
 * Read-only spot check for the Dorian Chen style case.
 * This catches rows that may not share UID but share customer/contact fields.
 */
function duplicate_phase1_spotCheckDorianChen() {
  return duplicate_phase1_spotCheckCustomerText_('dorian');
}

function duplicate_phase1_spotCheckCustomerText_(needle) {
  var ss = duplicateAuditSpreadsheet_();
  var master = duplicateCauseReadAppointmentSheet_(ss, '00_Master Appointments', 'MASTER');
  var q = duplicateAuditNormKey_(needle);
  var rows = master.records.filter(function (r) {
    return [
      r.name,
      r.email,
      r.phone,
      r.appt,
      r.root,
      r.uid
    ].join(' ').toLowerCase().indexOf(q) >= 0;
  });
  var out = {
    ok: true,
    query: needle,
    rowCount: rows.length,
    rows: rows.map(duplicateCauseBrief_)
  };
  Logger.log('PHASE1_CUSTOMER_SPOTCHECK_' + q.replace(/[^a-z0-9]+/g, '_').toUpperCase() + ' ' + JSON.stringify(out, null, 2));
  return out;
}

function duplicateAuditSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID') || props.getProperty('MASTER_FILE_ID') || '';
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (_) {}
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet and no SPREADSHEET_ID script property.');
  return ss;
}

function duplicateAuditHeaderMap_(headers) {
  var map = {};
  headers.forEach(function (h, i) {
    var raw = String(h || '').trim();
    if (!raw) return;
    map[raw] = i;
    map[duplicateAuditHeaderKey_(raw)] = i;
  });
  return map;
}

function duplicateAuditHeaderKey_(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function duplicateAuditPick_(map, names) {
  for (var i = 0; i < names.length; i++) {
    var raw = names[i];
    if (map[raw] != null) return map[raw];
    var key = duplicateAuditHeaderKey_(raw);
    if (map[key] != null) return map[key];
  }
  return -1;
}

function duplicateAuditCell_(row, idx) {
  return idx >= 0 ? row[idx] : '';
}

function duplicateAuditTrim_(value) {
  return String(value == null ? '' : value).trim();
}

function duplicateAuditNormKey_(value) {
  return duplicateAuditTrim_(value).toLowerCase().replace(/\s+/g, ' ');
}

function duplicateAuditNormEmail_(value) {
  return duplicateAuditTrim_(value).toLowerCase();
}

function duplicateAuditNormPhone_(value) {
  var d = duplicateAuditTrim_(value).replace(/\D+/g, '');
  if (d.length > 10 && d[0] === '1') d = d.slice(1);
  return d.length >= 7 ? d : '';
}

function duplicateAuditNonBlankCount_(row) {
  var n = 0;
  row.forEach(function (v) { if (duplicateAuditTrim_(v)) n++; });
  return n;
}

function duplicateAuditIsCurrent_(status, activeRaw) {
  var s = duplicateAuditNormKey_(status);
  var a = duplicateAuditNormKey_(activeRaw);
  if (a === 'yes' || a === 'true' || a === '1') return true;
  if (a === 'no' || a === 'false' || a === '0') return false;
  return !/cancel|resched|duplicate|superseded|inactive/.test(s);
}

function duplicateAuditCanonicalScore_(r) {
  var score = 0;
  if (r.isCurrent) score += 1000;
  if (r.uid) score += 100;
  if (r.visitNo) score += 80;
  if (r.appt) score += 40;
  if (r.root) score += 20;
  score += Math.min(r.nonBlankCount || 0, 50);
  score += Math.min(r.rowIndex || 0, 99999) / 100000;
  return score;
}

function duplicateAuditGroup_(records, keyFn) {
  var out = {};
  records.forEach(function (r) {
    var key = duplicateAuditTrim_(keyFn(r));
    if (!key) return;
    duplicateAuditPush_(out, key, r);
  });
  return out;
}

function duplicateAuditPush_(obj, key, value) {
  if (!obj[key]) obj[key] = [];
  obj[key].push(value);
}

function duplicateAuditIssues_(groups, issueType, predicate) {
  var out = [];
  Object.keys(groups).forEach(function (key) {
    var group = groups[key] || [];
    if (group.length < 2) return;
    if (predicate && !predicate(group)) return;
    var sorted = group.slice().sort(function (a, b) {
      return b.score - a.score;
    });
    out.push({
      issueType: issueType,
      key: key,
      rows: group.map(function (r) { return r.rowIndex; }),
      keepCandidate: sorted[0].rowIndex,
      records: sorted
    });
  });
  out.sort(function (a, b) {
    return b.records.length - a.records.length || String(a.key).localeCompare(String(b.key));
  });
  return out;
}

function duplicateAuditRowsInIssues_(issues) {
  var rows = {};
  issues.forEach(function (issue) {
    issue.rows.forEach(function (row) { rows[row] = true; });
  });
  return Object.keys(rows).length;
}

function duplicateAuditExampleIssues_(issues, limit) {
  return issues.slice(0, limit).map(function (issue) {
    return {
      issueType: issue.issueType,
      key: issue.key,
      rows: issue.rows,
      keepCandidate: issue.keepCandidate,
      records: issue.records.map(duplicateAuditBriefRecord_)
    };
  });
}

function duplicateAuditBriefRecord_(r) {
  return {
    row: r.rowIndex,
    appt: r.appt,
    root: r.root,
    visitNo: r.visitNo,
    name: r.name,
    email: r.email,
    phone: r.phone,
    brand: r.brand,
    visitDate: r.visitDate,
    visitTime: r.visitTime,
    visitType: r.visitType,
    status: r.status,
    active: r.activeRaw,
    uid: r.uid,
    timestamp: r.timestamp
  };
}

function duplicateCauseReadAppointmentSheet_(ss, sheetName, sourceName) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return { missing: true, records: [] };

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { missing: false, records: [] };

  var values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headers = values[0].map(function (h) { return String(h || '').trim(); });
  var H = duplicateAuditHeaderMap_(headers);
  var idx = {
    appt: duplicateAuditPick_(H, ['APPT_ID', 'Appt ID', 'Appointment ID']),
    root: duplicateAuditPick_(H, ['RootApptID', 'Root Appt ID', 'Root Appointment ID']),
    uid: duplicateAuditPick_(H, [
      'CalendlyEventUID',
      'Calendly Event UID',
      'Admin: Calendly Event UID',
      'Acuity ID',
      'Event UID',
      'UID'
    ]),
    visitNo: duplicateAuditPick_(H, ['Visit #', 'Visit Number', 'Visit No']),
    name: duplicateAuditPick_(H, ['Customer Name', 'Customer', 'Client Name', 'Name', 'Full Name']),
    emailLower: duplicateAuditPick_(H, ['EmailLower', 'Email Lower']),
    email: duplicateAuditPick_(H, ['Email', 'Email Address', 'E-mail']),
    phoneNorm: duplicateAuditPick_(H, ['PhoneNorm', 'Phone Norm']),
    phone: duplicateAuditPick_(H, ['Phone', 'Phone Number', 'Mobile', 'Tel']),
    brand: duplicateAuditPick_(H, ['Brand', 'Company']),
    visitDate: duplicateAuditPick_(H, ['Visit Date', 'Appointment Date', 'Date']),
    visitTime: duplicateAuditPick_(H, ['Visit Time', 'Appointment Time', 'Time']),
    visitType: duplicateAuditPick_(H, ['Visit Type', 'Appointment Type']),
    status: duplicateAuditPick_(H, ['Status']),
    active: duplicateAuditPick_(H, ['Active?', 'Active', 'Is Active']),
    timestamp: duplicateAuditPick_(H, ['Timestamp', 'Created At', 'Submitted At', 'QueuedAt']),
    rescheduledFrom: duplicateAuditPick_(H, ['RescheduledFromUID', 'Rescheduled From UID']),
    rescheduledTo: duplicateAuditPick_(H, ['RescheduledToUID', 'Rescheduled To UID']),
    canceledAt: duplicateAuditPick_(H, ['CanceledAt', 'Canceled At']),
    source: duplicateAuditPick_(H, ['Source', 'Source (normalized)']),
    notes: duplicateAuditPick_(H, ['Automation Notes', 'Notes', 'Error'])
  };

  var records = values.slice(1).map(function (row, i) {
    var rec = duplicateCauseRecordFromValues_(row, idx);
    rec.row = i + 2;
    rec.rowIndex = rec.row;
    rec.sheet = sheetName;
    rec.sourceSheet = sourceName;
    return rec;
  });

  return { missing: false, records: records, headers: headers };
}

function duplicateCauseRecordFromValues_(row, idx) {
  var status = duplicateAuditTrim_(duplicateAuditCell_(row, idx.status));
  var active = duplicateAuditTrim_(duplicateAuditCell_(row, idx.active));
  var rec = {
    appt: duplicateAuditTrim_(duplicateAuditCell_(row, idx.appt)),
    root: duplicateAuditTrim_(duplicateAuditCell_(row, idx.root)),
    uid: duplicateAuditTrim_(duplicateAuditCell_(row, idx.uid)),
    visitNo: duplicateAuditTrim_(duplicateAuditCell_(row, idx.visitNo)),
    name: duplicateAuditTrim_(duplicateAuditCell_(row, idx.name)),
    email: duplicateAuditNormEmail_(duplicateAuditCell_(row, idx.emailLower) || duplicateAuditCell_(row, idx.email)),
    phone: duplicateAuditNormPhone_(duplicateAuditCell_(row, idx.phoneNorm) || duplicateAuditCell_(row, idx.phone)),
    brand: duplicateAuditTrim_(duplicateAuditCell_(row, idx.brand)),
    visitDate: duplicateAuditTrim_(duplicateAuditCell_(row, idx.visitDate)),
    visitTime: duplicateAuditTrim_(duplicateAuditCell_(row, idx.visitTime)),
    visitType: duplicateAuditTrim_(duplicateAuditCell_(row, idx.visitType)),
    status: status,
    active: active,
    timestamp: duplicateAuditTrim_(duplicateAuditCell_(row, idx.timestamp)),
    rescheduledFrom: duplicateAuditTrim_(duplicateAuditCell_(row, idx.rescheduledFrom)),
    rescheduledTo: duplicateAuditTrim_(duplicateAuditCell_(row, idx.rescheduledTo)),
    canceledAt: duplicateAuditTrim_(duplicateAuditCell_(row, idx.canceledAt)),
    source: duplicateAuditTrim_(duplicateAuditCell_(row, idx.source)),
    notes: duplicateAuditTrim_(duplicateAuditCell_(row, idx.notes))
  };
  rec.brandNorm = duplicateAuditNormKey_(rec.brand);
  rec.visitDateNorm = duplicateAuditNormKey_(rec.visitDate);
  rec.visitTimeNorm = duplicateAuditNormKey_(rec.visitTime);
  rec.visitTypeNorm = duplicateAuditNormKey_(rec.visitType);
  rec.statusNorm = duplicateAuditNormKey_(rec.status);
  rec.isCurrent = duplicateAuditIsCurrent_(rec.status, rec.active);
  rec.contactKey = duplicateCauseContactKey_(rec);
  rec.fingerprint = duplicateCausePrimaryFingerprint_(rec);
  return rec;
}

function duplicateCauseReadQueue_(ss) {
  var sh = ss.getSheetByName('_IntakeQueue');
  if (!sh) return { missing: true, records: [], risks: [] };

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { missing: false, records: [], risks: [] };

  var values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headers = values[0].map(function (h) { return String(h || '').trim(); });
  var H = duplicateAuditHeaderMap_(headers);
  var idx = {
    queuedAt: duplicateAuditPick_(H, ['QueuedAt', 'Queued At']),
    status: duplicateAuditPick_(H, ['Status']),
    masterRowIndex: duplicateAuditPick_(H, ['MasterRowIndex', 'Master Row Index']),
    resolvedUid: duplicateAuditPick_(H, ['ResolvedUID', 'Resolved UID']),
    brand: duplicateAuditPick_(H, ['Brand', 'Company']),
    payload: duplicateAuditPick_(H, ['Payload']),
    processedAt: duplicateAuditPick_(H, ['ProcessedAt', 'Processed At']),
    error: duplicateAuditPick_(H, ['Error'])
  };

  var records = values.slice(1).map(function (row, i) {
    var payloadText = duplicateAuditTrim_(duplicateAuditCell_(row, idx.payload));
    var payload = duplicateCauseParsePayload_(payloadText);
    var rec = {
      row: i + 2,
      sheet: '_IntakeQueue',
      sourceSheet: 'INTAKE_QUEUE',
      queuedAt: duplicateAuditTrim_(duplicateAuditCell_(row, idx.queuedAt)),
      status: duplicateAuditTrim_(duplicateAuditCell_(row, idx.status)),
      masterRowIndex: duplicateAuditTrim_(duplicateAuditCell_(row, idx.masterRowIndex)),
      uid: duplicateAuditTrim_(duplicateAuditCell_(row, idx.resolvedUid)),
      brand: duplicateAuditTrim_(duplicateAuditCell_(row, idx.brand)) || duplicateCausePayloadValue_(payload, ['brand', 'Company', 'Brand']),
      payloadText: payloadText,
      processedAt: duplicateAuditTrim_(duplicateAuditCell_(row, idx.processedAt)),
      error: duplicateAuditTrim_(duplicateAuditCell_(row, idx.error)),
      name: duplicateCausePayloadValue_(payload, ['name', 'Customer Name', 'customerName', 'Name']),
      email: duplicateAuditNormEmail_(duplicateCausePayloadValue_(payload, ['email', 'Email', 'EmailLower'])),
      phone: duplicateAuditNormPhone_(duplicateCausePayloadValue_(payload, ['phone', 'Phone', 'PhoneNorm'])),
      visitDate: duplicateCausePayloadValue_(payload, ['visitDate', 'Visit Date']),
      visitTime: duplicateCausePayloadValue_(payload, ['visitTime', 'Visit Time']),
      visitType: duplicateCausePayloadValue_(payload, ['visitType', 'Visit Type'])
    };
    rec.brandNorm = duplicateAuditNormKey_(rec.brand);
    rec.visitDateNorm = duplicateAuditNormKey_(rec.visitDate);
    rec.visitTimeNorm = duplicateAuditNormKey_(rec.visitTime);
    rec.visitTypeNorm = duplicateAuditNormKey_(rec.visitType);
    rec.contactKey = duplicateCauseContactKey_(rec);
    rec.fingerprint = duplicateCausePrimaryFingerprint_(rec);
    return rec;
  });

  var openRows = records.filter(function (r) {
    return /pending|running|error/i.test(r.status || '');
  });
  var byUid = duplicateCauseGroup_(openRows, function (r) { return r.uid; });
  var byFingerprint = duplicateCauseGroup_(openRows, function (r) { return r.fingerprint; });
  var risks = [];

  Object.keys(byUid).forEach(function (uid) {
    var group = byUid[uid];
    if (uid && group.length > 1) {
      risks.push(duplicateCauseIssue_(uid, group, {
        queueIssue: 'Multiple open queue rows share the same ResolvedUID.'
      }));
    }
  });

  Object.keys(byFingerprint).forEach(function (key) {
    var group = byFingerprint[key];
    if (key && group.length > 1) {
      risks.push(duplicateCauseIssue_(key, group, {
        queueIssue: 'Multiple open queue rows share the same appointment fingerprint.'
      }));
    }
  });

  return { missing: false, records: records, risks: risks };
}

function duplicateCauseParsePayload_(text) {
  if (!text) return {};
  try {
    var parsed = JSON.parse(text);
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (_) {
    return {};
  }
}

function duplicateCausePayloadValue_(payload, names) {
  payload = payload || {};
  for (var i = 0; i < names.length; i++) {
    var k = names[i];
    if (payload[k] != null && payload[k] !== '') return duplicateAuditTrim_(payload[k]);
  }
  return '';
}

function duplicateCauseGroup_(records, keyFn) {
  var out = {};
  records.forEach(function (r) {
    var key = duplicateAuditTrim_(keyFn(r));
    if (!key) return;
    duplicateAuditPush_(out, key, r);
  });
  return out;
}

function duplicateCauseGroupFingerprints_(records) {
  return duplicateCauseGroup_(records, function (r) {
    return duplicateCausePrimaryFingerprint_(r);
  });
}

function duplicateCausePrimaryFingerprint_(r) {
  if (!r) return '';
  var contact = duplicateCauseContactKey_(r);
  if (!contact || !r.visitDateNorm) return '';
  return [
    r.brandNorm || '',
    r.visitDateNorm || '',
    r.visitTimeNorm || '',
    r.visitTypeNorm || '',
    contact
  ].join('|');
}

function duplicateCauseContactKey_(r) {
  if (!r) return '';
  if (r.email) return 'email:' + r.email;
  if (r.phone) return 'phone:' + r.phone;
  return '';
}

function duplicateCauseIssue_(key, group, extra) {
  extra = extra || {};
  var records = (group || []).slice().sort(function (a, b) {
    return (a.row || a.rowIndex || 0) - (b.row || b.rowIndex || 0);
  });
  var out = {
    key: key,
    rows: records.map(function (r) { return r.row || r.rowIndex; }),
    records: records.map(duplicateCauseBrief_)
  };
  Object.keys(extra).forEach(function (k) { out[k] = extra[k]; });
  return out;
}

function duplicateCauseBrief_(r) {
  return {
    row: r.row || r.rowIndex || '',
    sheet: r.sheet || '',
    sourceSheet: r.sourceSheet || '',
    appt: r.appt || '',
    root: r.root || '',
    uid: r.uid || '',
    visitNo: r.visitNo || '',
    name: r.name || '',
    email: r.email || '',
    phone: r.phone || '',
    brand: r.brand || '',
    visitDate: r.visitDate || '',
    visitTime: r.visitTime || '',
    visitType: r.visitType || '',
    status: r.status || '',
    active: r.active || '',
    timestamp: r.timestamp || r.queuedAt || '',
    rescheduledFrom: r.rescheduledFrom || '',
    rescheduledTo: r.rescheduledTo || '',
    canceledAt: r.canceledAt || '',
    source: r.source || '',
    queueStatus: r.status && r.sourceSheet === 'INTAKE_QUEUE' ? r.status : '',
    processedAt: r.processedAt || '',
    error: r.error || ''
  };
}

function duplicateCauseUnique_(values) {
  var seen = {};
  var out = [];
  values.forEach(function (v) {
    v = duplicateAuditTrim_(v);
    if (!v || seen[v]) return;
    seen[v] = true;
    out.push(v);
  });
  return out;
}

function duplicateCauseLikelyOrigin_(masterCount, upstreamCount) {
  if (upstreamCount >= masterCount && upstreamCount > 1) {
    return 'Likely upstream duplicate submissions before resolver.';
  }
  if (upstreamCount === 1 && masterCount > 1) {
    return 'Likely resolver, poller, or queue replay appended more than once from one upstream row.';
  }
  if (upstreamCount === 0 && masterCount > 1) {
    return 'Not visible in Form_Inbox; likely direct master edit, legacy import, or a source path not logged there.';
  }
  if (upstreamCount > 1 && upstreamCount < masterCount) {
    return 'Upstream duplicated and downstream likely replayed/appended again.';
  }
  return 'Unclear; inspect rows.';
}

function duplicateCauseUidSource_(uid) {
  uid = duplicateAuditTrim_(uid);
  if (!uid) return 'Missing UID';
  if (/^\d+_R\d+$/i.test(uid)) return 'Acuity synthetic reschedule UID';
  if (/^\d+$/.test(uid)) return 'Acuity appointment ID';
  if (/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(uid)) return 'Calendly event UID';
  return 'Unknown UID format';
}

function duplicateCauseBucketCounts_(buckets) {
  var out = {};
  Object.keys(buckets).forEach(function (k) {
    out[k] = (buckets[k] || []).length;
  });
  return out;
}
