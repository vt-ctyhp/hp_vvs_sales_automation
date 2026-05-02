/**
 * Phase 2 data model, lifecycle audit, and conservative cleanup tools.
 *
 * 00_Master Appointments remains the historical ledger. These routines build
 * generated read models and only cleanup rows that are exact high-confidence
 * replay duplicates. Do not run apply cleanup until reps are paused.
 */

var PHASE2_MODEL_TABS = {
  CURRENT_CUSTOMERS: '_Model_CurrentCustomers',
  APPOINTMENT_EVENTS: '_Model_AppointmentEvents',
  DATA_QUALITY: '_Model_DataQuality'
};

function phase2_lifecycleAuditOnly() {
  var model = phase2ModelBuild_();
  var out = {
    ok: true,
    generatedAt: model.generatedAt,
    masterRows: model.records.length,
    issueCounts: phase2ModelIssueCounts_(model.issues),
    issues: model.issues.slice(0, 200)
  };
  Logger.log('PHASE2_LIFECYCLE_AUDIT_SUMMARY ' + JSON.stringify({
    ok: out.ok,
    generatedAt: out.generatedAt,
    masterRows: out.masterRows,
    issueCounts: out.issueCounts
  }));
  Logger.log('PHASE2_LIFECYCLE_AUDIT_DETAILS ' + JSON.stringify(out, null, 2));
  return out;
}

function phase2_refreshModelTabs() {
  var model = phase2ModelBuild_();
  var ss = phase2ModelSpreadsheet_();

  var currentHeaders = [
    'GeneratedAt',
    'RootApptID',
    'CurrentMasterRow',
    'CurrentAPPT_ID',
    'Customer Name',
    'EmailLower',
    'PhoneNorm',
    'Brand',
    'Lifecycle State',
    'Current Visit Date',
    'Current Visit Time',
    'Current Visit Type',
    'Current Status',
    'Assigned Rep',
    'Assisted Rep',
    'SO#',
    'Client Status Report URL',
    'Row Count',
    'History Rows',
    'Data Quality Flags'
  ];

  var eventHeaders = [
    'GeneratedAt',
    'MasterRow',
    'RootApptID',
    'APPT_ID',
    'CalendlyEventUID',
    'Customer Name',
    'EmailLower',
    'PhoneNorm',
    'Brand',
    'Visit Date',
    'Visit Time',
    'Visit Type',
    'Visit #',
    'Status',
    'Active?',
    'Event Role',
    'RescheduledFromUID',
    'RescheduledToUID',
    'CanceledAt',
    'Fingerprint'
  ];

  var qualityHeaders = [
    'GeneratedAt',
    'Issue Type',
    'Severity',
    'Key',
    'Rows',
    'RootApptID',
    'APPT_ID',
    'UID',
    'Details'
  ];

  var currentRows = model.currentCustomers.map(function (r) {
    return [
      model.generatedAt,
      r.root,
      r.current ? r.current.row : '',
      r.current ? r.current.appt : '',
      r.current ? r.current.name : r.name,
      r.current ? r.current.email : r.email,
      r.current ? r.current.phone : r.phone,
      r.current ? r.current.brand : r.brand,
      r.current ? r.current.lifecycleState : 'historical',
      r.current ? r.current.visitDate : '',
      r.current ? r.current.visitTime : '',
      r.current ? r.current.visitType : '',
      r.current ? r.current.status : '',
      r.current ? r.current.assignedRep : '',
      r.current ? r.current.assistedRep : '',
      r.current ? r.current.so : '',
      r.current ? r.current.clientStatusReportUrl : '',
      r.rowCount,
      r.historyRows,
      r.flags.join('; ')
    ];
  });

  var eventRows = model.records.map(function (r) {
    return [
      model.generatedAt,
      r.row,
      r.root,
      r.appt,
      r.uid,
      r.name,
      r.email,
      r.phone,
      r.brand,
      r.visitDate,
      r.visitTime,
      r.visitType,
      r.visitNo,
      r.status,
      r.active,
      r.lifecycleState,
      r.rescheduledFrom,
      r.rescheduledTo,
      r.canceledAt,
      r.fingerprint
    ];
  });

  var qualityRows = model.issues.map(function (issue) {
    return [
      model.generatedAt,
      issue.type,
      issue.severity,
      issue.key,
      issue.rows.join(','),
      issue.root || '',
      issue.appt || '',
      issue.uid || '',
      issue.details || ''
    ];
  });

  phase2ModelReplaceSheet_(ss, PHASE2_MODEL_TABS.CURRENT_CUSTOMERS, currentHeaders, currentRows);
  phase2ModelReplaceSheet_(ss, PHASE2_MODEL_TABS.APPOINTMENT_EVENTS, eventHeaders, eventRows);
  phase2ModelReplaceSheet_(ss, PHASE2_MODEL_TABS.DATA_QUALITY, qualityHeaders, qualityRows);

  var summary = {
    ok: true,
    generatedAt: model.generatedAt,
    tabs: PHASE2_MODEL_TABS,
    counts: {
      currentCustomers: currentRows.length,
      appointmentEvents: eventRows.length,
      dataQualityIssues: qualityRows.length
    }
  };
  Logger.log('PHASE2_MODEL_REFRESH ' + JSON.stringify(summary));
  return summary;
}

function phase2_cleanupHighConfidenceDuplicatesDryRun() {
  var model = phase2ModelBuild_();
  var plan = phase2ModelCleanupPlan_(model);
  var out = {
    ok: true,
    generatedAt: model.generatedAt,
    duplicateRowsToMark: plan.rowsToMark.length,
    plan: plan
  };
  Logger.log('PHASE2_CLEANUP_DRY_RUN ' + JSON.stringify(out, null, 2));
  return out;
}

function phase2_cleanupHighConfidenceDuplicatesApply() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var model = phase2ModelBuild_();
    var plan = phase2ModelCleanupPlan_(model);
    var ss = phase2ModelSpreadsheet_();
    var sh = ss.getSheetByName('00_Master Appointments');
    if (!sh) throw new Error('Missing sheet: 00_Master Appointments');
    var H = phase2ModelHeaderMap_(sh);
    var now = new Date().toISOString();

    plan.rowsToMark.forEach(function (item) {
      if (H['Status']) sh.getRange(item.row, H['Status']).setValue('Duplicate');
      if (H['Active?']) sh.getRange(item.row, H['Active?']).setValue('No');
      if (H['Automation Notes']) {
        var prev = String(sh.getRange(item.row, H['Automation Notes']).getValue() || '');
        var note = 'Phase2 cleanup: high-confidence duplicate of row ' + item.keepRow + ' @ ' + now;
        sh.getRange(item.row, H['Automation Notes']).setValue(prev ? prev + '\n' + note : note);
      }
    });

    var out = {
      ok: true,
      generatedAt: now,
      rowsMarked: plan.rowsToMark.length,
      rows: plan.rowsToMark
    };
    Logger.log('PHASE2_CLEANUP_APPLY ' + JSON.stringify(out, null, 2));
    return out;
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function phase2ModelBuild_() {
  var generatedAt = new Date().toISOString();
  var master = phase2ModelReadMaster_();
  var records = master.records;
  var issues = phase2ModelIssues_(records);
  var issuesByRoot = {};

  issues.forEach(function (issue) {
    if (!issue.root) return;
    if (!issuesByRoot[issue.root]) issuesByRoot[issue.root] = [];
    issuesByRoot[issue.root].push(issue.type);
  });

  var byRoot = phase2ModelGroup_(records, function (r) { return r.root || r.appt; });
  var currentCustomers = Object.keys(byRoot).filter(Boolean).sort().map(function (root) {
    var group = byRoot[root];
    var currentRows = group.filter(function (r) { return r.isCurrent; });
    var current = phase2ModelBestRow_(currentRows.length ? currentRows : group);
    return {
      root: root,
      current: current,
      name: current ? current.name : '',
      email: current ? current.email : '',
      phone: current ? current.phone : '',
      brand: current ? current.brand : '',
      rowCount: group.length,
      historyRows: group.filter(function (r) { return !r.isCurrent; }).length,
      flags: phase2ModelUnique_(issuesByRoot[root] || [])
    };
  });

  return {
    generatedAt: generatedAt,
    headers: master.headers,
    records: records,
    issues: issues,
    currentCustomers: currentCustomers
  };
}

function phase2ModelReadMaster_() {
  var ss = phase2ModelSpreadsheet_();
  var sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet: 00_Master Appointments');

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { headers: [], records: [] };

  var values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headers = values[0].map(function (h) { return String(h || '').trim(); });
  var H = phase2ModelHeaderMapFromHeaders_(headers);
  var idx = {
    appt: phase2ModelPick_(H, ['APPT_ID', 'Appt ID', 'Appointment ID']),
    root: phase2ModelPick_(H, ['RootApptID', 'Root Appt ID', 'Root Appointment ID']),
    uid: phase2ModelPick_(H, ['CalendlyEventUID', 'Calendly Event UID', 'Admin: Calendly Event UID', 'Acuity ID']),
    visitNo: phase2ModelPick_(H, ['Visit #', 'Visit Number', 'Visit No']),
    name: phase2ModelPick_(H, ['Customer Name', 'Customer', 'Client Name', 'Name']),
    emailLower: phase2ModelPick_(H, ['EmailLower', 'Email Lower']),
    email: phase2ModelPick_(H, ['Email', 'Email Address', 'E-mail']),
    phoneNorm: phase2ModelPick_(H, ['PhoneNorm', 'Phone Norm']),
    phone: phase2ModelPick_(H, ['Phone', 'Phone Number', 'Mobile', 'Tel']),
    brand: phase2ModelPick_(H, ['Brand', 'Company']),
    visitDate: phase2ModelPick_(H, ['Visit Date', 'Appointment Date', 'Date']),
    visitTime: phase2ModelPick_(H, ['Visit Time', 'Appointment Time', 'Time']),
    visitType: phase2ModelPick_(H, ['Visit Type', 'Appointment Type']),
    status: phase2ModelPick_(H, ['Status']),
    active: phase2ModelPick_(H, ['Active?', 'Active', 'Is Active']),
    assignedRep: phase2ModelPick_(H, ['Assigned Rep', 'Rep']),
    assistedRep: phase2ModelPick_(H, ['Assisted Rep', 'Assistant Rep']),
    so: phase2ModelPick_(H, ['SO#', 'SO #', 'SO']),
    reportUrl: phase2ModelPick_(H, ['Client Status Report URL', 'Report URL']),
    rescheduledFrom: phase2ModelPick_(H, ['RescheduledFromUID', 'Rescheduled From UID']),
    rescheduledTo: phase2ModelPick_(H, ['RescheduledToUID', 'Rescheduled To UID']),
    canceledAt: phase2ModelPick_(H, ['CanceledAt', 'Canceled At'])
  };

  var records = values.slice(1).map(function (row, i) {
    var rec = {
      row: i + 2,
      appt: phase2ModelTrim_(phase2ModelCell_(row, idx.appt)),
      root: phase2ModelTrim_(phase2ModelCell_(row, idx.root)),
      uid: phase2ModelTrim_(phase2ModelCell_(row, idx.uid)),
      visitNo: phase2ModelTrim_(phase2ModelCell_(row, idx.visitNo)),
      name: phase2ModelTrim_(phase2ModelCell_(row, idx.name)),
      email: phase2ModelNormEmail_(phase2ModelCell_(row, idx.emailLower) || phase2ModelCell_(row, idx.email)),
      phone: phase2ModelNormPhone_(phase2ModelCell_(row, idx.phoneNorm) || phase2ModelCell_(row, idx.phone)),
      brand: phase2ModelTrim_(phase2ModelCell_(row, idx.brand)),
      visitDate: phase2ModelTrim_(phase2ModelCell_(row, idx.visitDate)),
      visitTime: phase2ModelTrim_(phase2ModelCell_(row, idx.visitTime)),
      visitType: phase2ModelTrim_(phase2ModelCell_(row, idx.visitType)),
      status: phase2ModelTrim_(phase2ModelCell_(row, idx.status)),
      active: phase2ModelTrim_(phase2ModelCell_(row, idx.active)),
      assignedRep: phase2ModelTrim_(phase2ModelCell_(row, idx.assignedRep)),
      assistedRep: phase2ModelTrim_(phase2ModelCell_(row, idx.assistedRep)),
      so: phase2ModelTrim_(phase2ModelCell_(row, idx.so)),
      clientStatusReportUrl: phase2ModelTrim_(phase2ModelCell_(row, idx.reportUrl)),
      rescheduledFrom: phase2ModelTrim_(phase2ModelCell_(row, idx.rescheduledFrom)),
      rescheduledTo: phase2ModelTrim_(phase2ModelCell_(row, idx.rescheduledTo)),
      canceledAt: phase2ModelTrim_(phase2ModelCell_(row, idx.canceledAt))
    };
    rec.root = rec.root || rec.appt;
    rec.statusNorm = phase2ModelNormKey_(rec.status);
    rec.activeNorm = phase2ModelNormKey_(rec.active);
    rec.isCurrent = phase2ModelIsCurrent_(rec);
    rec.fingerprint = phase2ModelFingerprint_(rec);
    rec.lifecycleState = phase2ModelLifecycleState_(rec);
    rec.score = phase2ModelScore_(rec);
    return rec;
  });

  return { headers: headers, records: records };
}

function phase2ModelIssues_(records) {
  var issues = [];
  var current = records.filter(function (r) { return r.isCurrent; });
  var byUid = phase2ModelGroup_(records, function (r) { return r.uid; });
  var currentByUid = phase2ModelGroup_(current, function (r) { return r.uid; });
  var currentByFingerprint = phase2ModelGroup_(current, function (r) { return r.fingerprint; });
  var byAppt = phase2ModelGroup_(records, function (r) { return r.appt; });

  Object.keys(currentByUid).forEach(function (uid) {
    var group = currentByUid[uid];
    if (uid && group.length > 1) {
      issues.push(phase2ModelIssue_('MULTIPLE_CURRENT_UID', 'high', uid, group, 'More than one current row shares the same external UID.'));
    }
  });

  Object.keys(currentByFingerprint).forEach(function (fp) {
    var group = currentByFingerprint[fp];
    if (fp && group.length > 1) {
      issues.push(phase2ModelIssue_('MULTIPLE_CURRENT_FINGERPRINT', 'high', fp, group, 'More than one current row shares the same appointment fingerprint.'));
    }
  });

  records.forEach(function (r) {
    if ((r.activeNorm === 'yes' || r.activeNorm === 'true' || r.activeNorm === '1') &&
        /cancel|resched|duplicate|superseded|inactive/.test(r.statusNorm)) {
      issues.push(phase2ModelIssue_('ACTIVE_STATUS_CONFLICT', 'high', r.uid || r.appt || String(r.row), [r], 'Active row has historical/canceled status.'));
    }
    if (/resched/.test(r.statusNorm) && !r.isCurrent && !r.rescheduledFrom && !r.rescheduledTo) {
      issues.push(phase2ModelIssue_('RESCHEDULE_HISTORY_MISSING_LINKS', 'medium', r.appt || String(r.row), [r], 'Historical reschedule row is missing old/new UID links.'));
    }
    if (!r.root && r.appt) {
      issues.push(phase2ModelIssue_('MISSING_ROOT_APPT_ID', 'medium', r.appt, [r], 'RootApptID is blank; downstream code may fall back to APPT_ID.'));
    }
  });

  Object.keys(byAppt).forEach(function (appt) {
    var group = byAppt[appt];
    if (!appt || group.length < 2) return;
    var contacts = phase2ModelUnique_(group.map(function (r) { return phase2ModelContactKey_(r); }).filter(Boolean));
    var occurrences = phase2ModelUnique_(group.map(function (r) {
      return [phase2ModelNormDate_(r.visitDate), phase2ModelNormTime_(r.visitTime), phase2ModelNormKey_(r.visitType), r.visitNo].join('|');
    }).filter(function (v) { return v.replace(/\|/g, ''); }));
    var type = contacts.length > 1 ? 'APPT_ID_SHARED_ACROSS_CONTACTS' : 'APPT_ID_REPEATED_AS_HISTORY';
    var severity = contacts.length > 1 ? 'medium' : 'info';
    issues.push(phase2ModelIssue_(type, severity, appt, group,
      contacts.length > 1
        ? 'Repeated APPT_ID crosses contacts; do not use APPT_ID alone as identity.'
        : 'Repeated APPT_ID appears to be history/root chain context, not exact duplicate proof.'
    ));
  });

  Object.keys(byUid).forEach(function (uid) {
    var group = byUid[uid];
    if (!uid || group.length < 2) return;
    var fingerprints = phase2ModelUnique_(group.map(function (r) { return r.fingerprint; }).filter(Boolean));
    if (fingerprints.length > 1) {
      issues.push(phase2ModelIssue_('UID_REUSED_ACROSS_OCCURRENCES', 'medium', uid, group, 'Same external UID appears on different appointment fingerprints.'));
    }
  });

  return issues.sort(function (a, b) {
    var rank = { high: 0, medium: 1, low: 2, info: 3 };
    return (rank[a.severity] || 9) - (rank[b.severity] || 9) || String(a.type).localeCompare(String(b.type));
  });
}

function phase2ModelCleanupPlan_(model) {
  var groups = {};
  model.records.filter(function (r) { return r.isCurrent && r.fingerprint; }).forEach(function (r) {
    var key = r.uid ? ('uidfp|' + r.uid + '|' + r.fingerprint) : ('blankuidfp|' + r.fingerprint);
    if (!groups[key]) groups[key] = [];
    groups[key].push(r);
  });

  var rowsToMark = [];
  var groupsOut = [];
  Object.keys(groups).forEach(function (key) {
    var group = groups[key];
    if (group.length < 2) return;
    var keep = phase2ModelBestRow_(group);
    var mark = group.filter(function (r) { return r.row !== keep.row; });
    mark.forEach(function (r) {
      rowsToMark.push({
        row: r.row,
        keepRow: keep.row,
        appt: r.appt,
        root: r.root,
        uid: r.uid,
        fingerprint: r.fingerprint,
        reason: 'Exact current UID/fingerprint replay duplicate.'
      });
    });
    groupsOut.push({
      key: key,
      keepRow: keep.row,
      duplicateRows: mark.map(function (r) { return r.row; })
    });
  });

  return {
    groups: groupsOut,
    rowsToMark: rowsToMark.sort(function (a, b) { return a.row - b.row; })
  };
}

function phase2ModelReplaceSheet_(ss, name, headers, rows) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.setFrozenRows(1);
  try { sh.autoResizeColumns(1, headers.length); } catch (_) {}
  try { sh.hideSheet(); } catch (_) {}
  try {
    var protections = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (!protections.length) protections = [sh.protect().setDescription(name + ' generated model tab')];
    protections.forEach(function (p) {
      p.setWarningOnly(true);
      p.setDescription(name + ' generated model tab');
    });
  } catch (_) {}
}

function phase2ModelIssue_(type, severity, key, group, details) {
  var first = (group && group[0]) || {};
  return {
    type: type,
    severity: severity,
    key: key,
    rows: (group || []).map(function (r) { return r.row; }),
    root: first.root || '',
    appt: first.appt || '',
    uid: first.uid || '',
    details: details
  };
}

function phase2ModelIssueCounts_(issues) {
  var out = {};
  (issues || []).forEach(function (issue) {
    out[issue.type] = (out[issue.type] || 0) + 1;
  });
  return out;
}

function phase2ModelBestRow_(records) {
  if (!records || !records.length) return null;
  return records.slice().sort(function (a, b) {
    return b.score - a.score || b.row - a.row;
  })[0];
}

function phase2ModelScore_(r) {
  var score = 0;
  if (r.isCurrent) score += 1000;
  if (r.uid) score += 100;
  if (r.visitNo) score += 80;
  if (r.appt) score += 40;
  if (r.root) score += 20;
  score += Math.min(r.row || 0, 99999) / 100000;
  return score;
}

function phase2ModelLifecycleState_(r) {
  var s = r.statusNorm || '';
  if (/duplicate|superseded/.test(s)) return 'duplicate/superseded';
  if (/cancel/.test(s) || r.canceledAt) return 'canceled';
  if (/resched/.test(s) || r.rescheduledTo) return 'rescheduled-away';
  if (/no.?show/.test(s)) return 'no-show';
  if (/complete/.test(s)) return r.isCurrent ? 'completed/current' : 'historical';
  if (r.isCurrent) return 'scheduled/current';
  return 'historical';
}

function phase2ModelIsCurrent_(r) {
  var a = r.activeNorm || '';
  if (/cancel|resched|duplicate|superseded|inactive/.test(r.statusNorm || '') || r.rescheduledTo || r.canceledAt) return false;
  if (a === 'yes' || a === 'true' || a === '1') return true;
  if (a === 'no' || a === 'false' || a === '0') return false;
  return true;
}

function phase2ModelFingerprint_(r) {
  var contact = phase2ModelContactKey_(r);
  var d = phase2ModelNormDate_(r.visitDate);
  var t = phase2ModelNormTime_(r.visitTime);
  if (!contact || !d || !t) return '';
  return [
    phase2ModelNormKey_(r.brand),
    d,
    t,
    phase2ModelNormKey_(r.visitType),
    contact
  ].join('|');
}

function phase2ModelContactKey_(r) {
  if (!r) return '';
  if (r.email) return 'email:' + r.email;
  if (r.phone) return 'phone:' + r.phone;
  return '';
}

function phase2ModelGroup_(records, keyFn) {
  var out = {};
  records.forEach(function (r) {
    var key = phase2ModelTrim_(keyFn(r));
    if (!key) return;
    if (!out[key]) out[key] = [];
    out[key].push(r);
  });
  return out;
}

function phase2ModelUnique_(values) {
  var seen = {};
  var out = [];
  values.forEach(function (v) {
    v = phase2ModelTrim_(v);
    if (!v || seen[v]) return;
    seen[v] = true;
    out.push(v);
  });
  return out;
}

function phase2ModelSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID') || props.getProperty('MASTER_FILE_ID') || '';
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (_) {}
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet and no SPREADSHEET_ID script property.');
  return ss;
}

function phase2ModelHeaderMap_(sh) {
  return phase2ModelHeaderMapFromHeaders_(sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0]);
}

function phase2ModelHeaderMapFromHeaders_(headers) {
  var map = {};
  headers.forEach(function (h, i) {
    var raw = phase2ModelTrim_(h);
    if (!raw) return;
    map[raw] = i + 1;
    map[phase2ModelHeaderKey_(raw)] = i + 1;
  });
  return map;
}

function phase2ModelHeaderKey_(value) {
  return phase2ModelTrim_(value).toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function phase2ModelPick_(map, names) {
  for (var i = 0; i < names.length; i++) {
    var raw = names[i];
    if (map[raw] != null) return map[raw] - 1;
    var key = phase2ModelHeaderKey_(raw);
    if (map[key] != null) return map[key] - 1;
  }
  return -1;
}

function phase2ModelCell_(row, idx) {
  return idx >= 0 ? row[idx] : '';
}

function phase2ModelTrim_(value) {
  return String(value == null ? '' : value).trim();
}

function phase2ModelNormKey_(value) {
  return phase2ModelTrim_(value).toLowerCase().replace(/\s+/g, ' ');
}

function phase2ModelNormEmail_(value) {
  return phase2ModelTrim_(value).toLowerCase();
}

function phase2ModelNormPhone_(value) {
  var d = phase2ModelTrim_(value).replace(/\D+/g, '');
  if (d.length > 10 && d[0] === '1') d = d.slice(1);
  return d.length >= 7 ? d : '';
}

function phase2ModelNormDate_(value) {
  value = phase2ModelTrim_(value);
  if (!value) return '';
  var iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3];
  var mdY = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(value);
  if (mdY) return mdY[3] + '-' + String(mdY[1]).padStart(2, '0') + '-' + String(mdY[2]).padStart(2, '0');
  var dt = new Date(value);
  if (!isNaN(dt.getTime())) return Utilities.formatDate(dt, Session.getScriptTimeZone() || 'America/Los_Angeles', 'yyyy-MM-dd');
  return phase2ModelNormKey_(value);
}

function phase2ModelNormTime_(value) {
  value = phase2ModelTrim_(value);
  if (!value) return '';
  var m12 = /^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)$/i.exec(value);
  if (m12) {
    var h = parseInt(m12[1], 10);
    var ap = m12[3].toUpperCase();
    if (ap === 'AM' && h === 12) h = 0;
    if (ap === 'PM' && h !== 12) h += 12;
    return String(h).padStart(2, '0') + ':' + m12[2];
  }
  var m24 = /^(\d{1,2}):(\d{2})(?::\d{2})?$/.exec(value);
  if (m24) return String(parseInt(m24[1], 10)).padStart(2, '0') + ':' + m24[2];
  return phase2ModelNormKey_(value);
}
