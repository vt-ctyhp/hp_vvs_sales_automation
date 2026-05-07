/**
 * Unified Sales Workflow Inbox: append-only notification log, broadcast sends,
 * and read-model serving helpers.
 */

var SW_INBOX_DEFAULT_LIMIT = 50;
var SW_INBOX_DEFAULT_WINDOW_DAYS = 7;

function sw_getInboxNotifications(authToken, options) {
  return swTimed_('sw_getInboxNotifications', function () {
    var mark = swStepTimer_('sw_getInboxNotifications');
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    mark('requiredSheets');
    var user = swAuthUserForApi_(ss, authToken);
    mark('identity', { canBroadcast: swInboxCanSendBroadcast_(user) });
    options = swInboxNormalizeListOptions_(options);
    var projected = swTryGetInboxNotificationsFromReadModel_(ss, user, options);
    if (projected && projected.ok) {
      mark('read', {
        source: projected.source || '',
        rows: projected.sourceRows || 0,
        returned: (projected.notifications || []).length,
        ageSeconds: projected.readModelAgeSeconds || 0
      });
      return projected;
    }
    var fallback = swReadInboxNotificationsFromLog_(ss, user, options, projected && projected.fallbackReason);
    mark('read', {
      source: fallback.source || '',
      rows: fallback.sourceRows || 0,
      returned: (fallback.notifications || []).length,
      fallbackReason: fallback.fallbackReason || ''
    });
    return fallback;
  });
}

function sw_getInboxBroadcastOptions(authToken) {
  return swTimed_('sw_getInboxBroadcastOptions', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken);
    if (!swInboxCanSendBroadcast_(user)) throw new Error('Broadcast access requires Admin or Diamond Order Admin.');
    return {
      ok: true,
      recipientModes: ['all', 'roles', 'users'],
      roleOptions: swInboxRoleOptions_(),
      users: swInboxActiveWorkflowUsers_(ss).map(function (row) {
        return {
          email: row.email,
          name: row.name,
          roles: row.roles,
          roleLabels: row.roleLabels
        };
      })
    };
  });
}

function sw_sendInboxBroadcast(authToken, payload) {
  return swTimed_('sw_sendInboxBroadcast', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken);
    if (!swInboxCanSendBroadcast_(user)) throw new Error('Broadcast access requires Admin or Diamond Order Admin.');
    payload = payload || {};
    var subject = swTrim_(payload.subject || payload.title || '');
    var body = swTrim_(payload.body || payload.message || '');
    if (!subject) throw new Error('Subject is required.');
    if (!body) throw new Error('Body is required.');
    var resolved = swInboxResolveBroadcastRecipients_(ss, payload);
    if (!resolved.emails.length) throw new Error('Select at least one active recipient.');
    var now = swIso_(new Date());
    var senderRoles = (user.roles || []).join(',');
    var notificationId = 'INBOX-BCAST-' + Utilities.getUuid();
    var record = {
      notificationId: notificationId,
      createdAt: now,
      kind: 'broadcast',
      eventType: 'BROADCAST',
      title: subject,
      body: body,
      badgeLabel: 'Message',
      senderName: user.name || user.email || 'Workflow Admin',
      senderEmail: user.email || '',
      senderRole: senderRoles,
      recipientMode: resolved.mode,
      roleTargets: resolved.roleTargets,
      userTargets: resolved.userTargets,
      visibleEmails: resolved.emails,
      fingerprint: notificationId,
      payload: {
        recipientCount: resolved.emails.length
      }
    };
    swInboxAppendRecord_(ss, record);
    return {
      ok: true,
      notificationId: notificationId,
      recipientCount: resolved.emails.length,
      message: 'Broadcast sent to ' + resolved.emails.length + ' recipient(s).'
    };
  });
}

function swBuildInboxReadModel_(ss, builtAt) {
  var started = new Date().getTime();
  var builtAtIso = swIso_(builtAt || new Date());
  try {
    var records = swInboxReadLogRecords_(ss);
    records.sort(swInboxSortDesc_);
    var byMonth = {};
    records.forEach(function (rec) {
      var key = swInboxMonthKey_(rec.createdAt);
      if (!byMonth[key]) byMonth[key] = [];
      byMonth[key].push(swInboxPublicRecord_(rec));
    });
    var rows = Object.keys(byMonth).sort().reverse().map(function (monthKey) {
      var items = byMonth[monthKey] || [];
      var values = [
        monthKey,
        items.length,
        swStringify_(items),
        builtAtIso
      ];
      values.push(swReadModelSearchText_(values.slice(0, 3)));
      return values;
    });
    var write = swWriteReadModelSheet_(ss, SW_SHEETS.READ_MODEL_INBOX, SW_INBOX_READ_MODEL_HEADERS, rows);
    write.sourceRows = records.length;
    write.outputRows = rows.length;
    return write;
  } catch (err) {
    return swReadModelErrorResult_(err, started, SW_SHEETS.READ_MODEL_INBOX);
  }
}

function swTryGetInboxNotificationsFromReadModel_(ss, user, options) {
  var config = [];
  try { config = swReadConfig_(ss, true); } catch (_) {}
  if (!swReadModelServingFlag_(config, 'READ_MODEL_SERVE_INBOX', 'Y')) {
    return { ok: false, fallbackReason: 'disabled' };
  }
  var status = swReadModelFreshStatus_(ss, 'inbox', SW_SHEETS.READ_MODEL_INBOX);
  if (!status.fresh) return { ok: false, fallbackReason: status.reason || 'notFresh', readModelAgeSeconds: status.ageSeconds || 0 };
  var sh = ss.getSheetByName(SW_SHEETS.READ_MODEL_INBOX);
  var rows = swReadSheetObjectsExpectedHeaders_(sh, SW_INBOX_READ_MODEL_HEADERS);
  var records = [];
  rows.forEach(function (row) {
    var items = swParseJson_(row['Notifications JSON'] || '[]', []);
    if (!Array.isArray(items)) return;
    items.forEach(function (item) { records.push(item); });
  });
  return swInboxPageResponse_(records, user, options, {
    source: 'inboxReadModelSheet',
    sourceRows: rows.length,
    readModelAgeSeconds: status.ageSeconds || 0,
    fallbackReason: ''
  });
}

function swReadInboxNotificationsFromLog_(ss, user, options, fallbackReason) {
  var records = swInboxReadLogRecords_(ss).map(swInboxPublicRecord_);
  return swInboxPageResponse_(records, user, options, {
    source: 'inboxLog',
    sourceRows: records.length,
    fallbackReason: fallbackReason || ''
  });
}

function swInboxPageResponse_(records, user, options, meta) {
  records = (records || []).filter(function (rec) {
    if (!swInboxRecordVisibleToUser_(rec, user)) return false;
    if (options.filter === 'schedule' && rec.kind !== 'schedule') return false;
    if (options.filter === 'broadcast' && rec.kind !== 'broadcast') return false;
    return true;
  }).sort(swInboxSortDesc_);

  var cursor = swInboxDecodeCursor_(options.cursor);
  var sinceMs = options.cursor ? 0 : new Date().getTime() - options.windowDays * 24 * 60 * 60 * 1000;
  var olderWindowAvailable = false;
  if (cursor && cursor.beforeAtMs) {
    records = records.filter(function (rec) {
      var ms = swInboxDateMs_(rec.createdAt);
      if (ms < cursor.beforeAtMs) return true;
      return ms === cursor.beforeAtMs && String(rec.id || rec.notificationId || '') < String(cursor.beforeId || '');
    });
  } else if (sinceMs) {
    olderWindowAvailable = records.some(function (rec) {
      return swInboxDateMs_(rec.createdAt) < sinceMs;
    });
    records = records.filter(function (rec) {
      return swInboxDateMs_(rec.createdAt) >= sinceMs;
    });
  }

  var limit = options.limit;
  var page = records.slice(0, limit + 1);
  var hasMore = page.length > limit;
  if (hasMore) page = page.slice(0, limit);
  var last = page.length ? page[page.length - 1] : null;
  var nextCursor = hasMore && last ? swInboxEncodeCursor_(last.createdAt, last.id || last.notificationId || '') : '';
  if (!hasMore && !options.cursor && olderWindowAvailable) {
    hasMore = true;
    nextCursor = swInboxEncodeCursor_(swIso_(new Date(sinceMs)), '');
  }
  var summary = { total: page.length, schedule: 0, broadcast: 0 };
  page.forEach(function (rec) {
    if (rec.kind === 'schedule') summary.schedule++;
    if (rec.kind === 'broadcast') summary.broadcast++;
  });
  meta = meta || {};
  return {
    ok: true,
    generatedAt: swIso_(new Date()),
    source: meta.source || '',
    sourceRows: meta.sourceRows || 0,
    fallbackReason: meta.fallbackReason || '',
    readModelAgeSeconds: meta.readModelAgeSeconds || 0,
    windowDays: options.windowDays,
    nextCursor: nextCursor,
    hasMore: hasMore,
    summary: summary,
    canSendBroadcast: swInboxCanSendBroadcast_(user),
    notifications: page
  };
}

function swInboxLogAppointmentScheduleChangeFromRows_(ss, eventType, rowNumber, previousRowNumber) {
  try {
    ss = ss || swSpreadsheet_();
    var rec = swInboxAppointmentByRow_(ss, rowNumber);
    if (!rec) return { ok: false, skipped: true, reason: 'missingAppointment' };
    var previous = previousRowNumber ? swInboxAppointmentByRow_(ss, previousRowNumber) : null;
    var record = swInboxScheduleRecord_(eventType, rec, previous);
    return swInboxAppendRecord_(ss, record);
  } catch (err) {
    try {
      Logger.log('SW_INBOX_SCHEDULE_LOG_ERROR ' + JSON.stringify({ eventType: eventType, rowNumber: rowNumber, error: err && err.message ? err.message : String(err) }));
    } catch (_) {}
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

function swInboxScheduleRecord_(eventType, rec, previous) {
  eventType = swTrim_(eventType || '').toUpperCase();
  var createdAt = swIso_(new Date());
  var visitAt = swVisitDateTime_(rec, swTimezone_());
  var bookedAtMs = swInboxDateMs_(rec.bookedAtRaw || rec.bookedAt);
  var bookedWithin24 = false;
  if (eventType === 'NEW_APPOINTMENT' && visitAt && bookedAtMs) {
    var diff = visitAt.getTime() - bookedAtMs;
    bookedWithin24 = diff >= 0 && diff <= 24 * 60 * 60 * 1000;
  }
  var title = 'Appointment update';
  var badge = 'Schedule';
  if (eventType === 'NEW_APPOINTMENT') {
    title = bookedWithin24 ? 'Last-minute appointment booked' : 'New appointment booked';
    badge = bookedWithin24 ? 'Within 24h' : 'New';
  } else if (eventType === 'APPOINTMENT_RESCHEDULED') {
    title = 'Appointment rescheduled';
    badge = 'Rescheduled';
  } else if (eventType === 'APPOINTMENT_CANCELED') {
    title = 'Appointment canceled';
    badge = 'Canceled';
  }
  var body = swInboxScheduleBody_(eventType, rec, previous, bookedWithin24);
  var fingerprint = [
    'schedule',
    eventType,
    rec.appt || rec.uid || rec.row || '',
    previous && (previous.appt || previous.uid || previous.row) || '',
    rec.visitDate || '',
    rec.visitTime || '',
    rec.bookedAt || '',
    rec.canceledAt || ''
  ].join('|');
  return {
    notificationId: 'INBOX-SCHED-' + swInboxDigest_(fingerprint).slice(0, 24),
    createdAt: createdAt,
    kind: 'schedule',
    eventType: eventType,
    title: title,
    body: body,
    badgeLabel: badge,
    senderName: 'System',
    senderEmail: '',
    senderRole: 'System',
    recipientMode: 'all',
    roleTargets: [],
    userTargets: [],
    visibleEmails: [],
    root: rec.root || rec.appt || '',
    appt: rec.appt || '',
    previousAppt: previous && previous.appt || '',
    customerName: rec.name || '',
    brand: rec.brand || '',
    visitDate: rec.visitDate || '',
    visitTime: rec.visitTime || '',
    previousVisitDate: previous && previous.visitDate || '',
    previousVisitTime: previous && previous.visitTime || '',
    bookedAt: rec.bookedAt || '',
    bookedWithin24: bookedWithin24,
    fingerprint: fingerprint,
    payload: {
      row: rec.row || '',
      previousRow: previous && previous.row || '',
      assignedRep: rec.assignedRep || '',
      assistedRep: rec.assistedRep || '',
      status: rec.status || ''
    }
  };
}

function swInboxScheduleBody_(eventType, rec, previous, bookedWithin24) {
  var who = rec.name || 'No customer';
  var next = [rec.visitDate, swFormatAppointmentTime_(rec.visitTime || '')].filter(Boolean).join(' ');
  var owners = [
    rec.assignedRep ? 'Advisor ' + rec.assignedRep : '',
    rec.assistedRep ? 'JOC ' + rec.assistedRep : ''
  ].filter(Boolean).join(' | ');
  var parts = [];
  if (eventType === 'APPOINTMENT_RESCHEDULED') {
    var prev = previous ? [previous.visitDate, swFormatAppointmentTime_(previous.visitTime || '')].filter(Boolean).join(' ') : '';
    parts.push(who + ' moved from ' + (prev || 'the previous time') + ' to ' + (next || 'the new time') + '.');
  } else if (eventType === 'APPOINTMENT_CANCELED') {
    parts.push(who + ' canceled ' + (next || 'their appointment') + '.');
  } else {
    parts.push(who + ' booked ' + (next || 'a new appointment') + '.');
    if (bookedWithin24) parts.push('Booked within 24 hours of the visit.');
  }
  if (owners) parts.push(owners + '.');
  return parts.join(' ');
}

function swInboxAppointmentByRow_(ss, rowNumber) {
  rowNumber = Number(rowNumber || 0);
  if (!rowNumber || rowNumber < 2) return null;
  var sh = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!sh || rowNumber > sh.getLastRow()) return null;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
  var idx = swAppointmentColumnIndex_(headers);
  var indexes = swAppointmentColumnIndexes_(idx);
  var values = swReadSelectedRows_(sh, rowNumber, 1, indexes, 'values')[0] || [];
  var display = swReadSelectedRows_(sh, rowNumber, 1, indexes, 'display')[0] || [];
  return swAppointmentRecordFromRows_(display, values, idx, rowNumber);
}

function swInboxAppendRecord_(ss, record) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.INBOX_LOG, SW_INBOX_LOG_HEADERS);
  var fingerprint = swTrim_(record && record.fingerprint || '');
  if (fingerprint && swInboxFingerprintExists_(sh, fingerprint)) {
    return { ok: true, skipped: true, reason: 'duplicateFingerprint' };
  }
  var row = swInboxLogRow_(record || {});
  sh.getRange(sh.getLastRow() + 1, 1, 1, SW_INBOX_LOG_HEADERS.length).setValues([row]);
  try { swMarkWorkflowReadModelsStale_(ss, 'Inbox notification appended', 'inbox'); } catch (_) {}
  return { ok: true, notificationId: record.notificationId || '', rowNumber: sh.getLastRow() };
}

function swInboxFingerprintExists_(sh, fingerprint) {
  if (!fingerprint || sh.getLastRow() < 2) return false;
  var headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), SW_INBOX_LOG_HEADERS.length)).getDisplayValues()[0].map(swTrim_);
  var col = swPickIndex_(swHeaderMapFromArray_(headers), ['Fingerprint']) + 1;
  if (!col) return false;
  var values = sh.getRange(2, col, sh.getLastRow() - 1, 1).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    if (swTrim_(values[i][0]) === fingerprint) return true;
  }
  return false;
}

function swInboxLogRow_(rec) {
  return [
    rec.notificationId || '',
    rec.createdAt || swIso_(new Date()),
    rec.kind || '',
    rec.eventType || '',
    rec.title || '',
    rec.body || '',
    rec.badgeLabel || '',
    rec.senderName || '',
    rec.senderEmail || '',
    rec.senderRole || '',
    rec.recipientMode || '',
    swStringify_(rec.roleTargets || []),
    swStringify_(rec.userTargets || []),
    swStringify_(rec.visibleEmails || []),
    rec.root || '',
    rec.appt || '',
    rec.previousAppt || '',
    rec.customerName || '',
    rec.brand || '',
    rec.visitDate || '',
    rec.visitTime || '',
    rec.previousVisitDate || '',
    rec.previousVisitTime || '',
    rec.bookedAt || '',
    rec.bookedWithin24 ? 'Y' : '',
    rec.fingerprint || '',
    swStringify_(rec.payload || {})
  ];
}

function swInboxReadLogRecords_(ss) {
  var sh = ss.getSheetByName(SW_SHEETS.INBOX_LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  return swReadSheetObjectsExpectedHeaders_(sh, SW_INBOX_LOG_HEADERS).map(swInboxLogRecordFromRow_).filter(function (rec) {
    return !!rec.notificationId;
  });
}

function swInboxLogRecordFromRow_(row) {
  row = row || {};
  return {
    id: row['NotificationID'] || '',
    notificationId: row['NotificationID'] || '',
    createdAt: row['Created At'] || '',
    kind: row['Kind'] || '',
    eventType: row['Event Type'] || '',
    title: row['Title'] || '',
    body: row['Body'] || '',
    badgeLabel: row['Badge Label'] || '',
    senderName: row['Sender Name'] || '',
    senderEmail: row['Sender Email'] || '',
    senderRole: row['Sender Role'] || '',
    recipientMode: row['Recipient Mode'] || '',
    roleTargets: swParseJson_(row['Role Targets JSON'] || '[]', []),
    userTargets: swParseJson_(row['User Targets JSON'] || '[]', []),
    visibleEmails: swParseJson_(row['Visible Emails JSON'] || '[]', []),
    root: row['RootApptID'] || '',
    appt: row['APPT_ID'] || '',
    previousAppt: row['Previous APPT_ID'] || '',
    customerName: row['Customer Name'] || '',
    brand: row['Brand'] || '',
    visitDate: row['Visit Date'] || '',
    visitTime: row['Visit Time'] || '',
    previousVisitDate: row['Previous Visit Date'] || '',
    previousVisitTime: row['Previous Visit Time'] || '',
    bookedAt: row['Booked At'] || '',
    bookedWithin24: swTruthy_(row['Booked Within 24 Hours?']),
    fingerprint: row['Fingerprint'] || '',
    payload: swParseJson_(row['Payload JSON'] || '{}', {})
  };
}

function swInboxPublicRecord_(rec) {
  rec = rec || {};
  return {
    id: rec.id || rec.notificationId || '',
    notificationId: rec.notificationId || rec.id || '',
    createdAt: rec.createdAt || '',
    kind: rec.kind || '',
    eventType: rec.eventType || '',
    title: rec.title || '',
    body: rec.body || '',
    badgeLabel: rec.badgeLabel || '',
    senderName: rec.senderName || '',
    senderRole: rec.senderRole || '',
    recipientMode: rec.recipientMode || '',
    visibleEmails: rec.visibleEmails || [],
    root: rec.root || '',
    appt: rec.appt || '',
    previousAppt: rec.previousAppt || '',
    customerName: rec.customerName || '',
    brand: rec.brand || '',
    visitDate: rec.visitDate || '',
    visitTime: rec.visitTime || '',
    previousVisitDate: rec.previousVisitDate || '',
    previousVisitTime: rec.previousVisitTime || '',
    bookedAt: rec.bookedAt || '',
    bookedWithin24: !!rec.bookedWithin24,
    payload: rec.payload || {}
  };
}

function swInboxNormalizeListOptions_(options) {
  options = options || {};
  var limit = Math.round(Number(options.limit || SW_INBOX_DEFAULT_LIMIT));
  if (!isFinite(limit) || limit <= 0) limit = SW_INBOX_DEFAULT_LIMIT;
  limit = Math.min(limit, 100);
  var days = Math.round(Number(options.windowDays || SW_INBOX_DEFAULT_WINDOW_DAYS));
  if (!isFinite(days) || days <= 0) days = SW_INBOX_DEFAULT_WINDOW_DAYS;
  var filter = swNorm_(options.filter || 'all');
  if (['all', 'schedule', 'broadcast'].indexOf(filter) < 0) filter = 'all';
  return {
    cursor: swTrim_(options.cursor || ''),
    limit: limit,
    windowDays: days,
    filter: filter
  };
}

function swInboxCanSendBroadcast_(user) {
  user = user || {};
  return !!(user.isAdmin || user.isDiamondOrderAdmin);
}

function swInboxActiveWorkflowUsers_(ss) {
  return swAuthReadPublicUserRowsCached_(ss).filter(function (row) {
    return row.email && swWorkflowUserActive_(row);
  }).map(function (row) {
    var roles = swAuthRoles_(row.roles || '');
    return {
      email: swNormEmail_(row.email),
      name: row.name || row.email,
      roles: roles,
      roleLabels: roles.map(swInboxRoleLabel_).join(', ')
    };
  });
}

function swInboxRoleOptions_() {
  return (typeof swAuthRoleOptions_ === 'function' ? swAuthRoleOptions_() : []).map(function (role) {
    return {
      value: role.value,
      label: role.label || swInboxRoleLabel_(role.value)
    };
  });
}

function swInboxResolveBroadcastRecipients_(ss, payload) {
  payload = payload || {};
  var mode = swNorm_(payload.recipientMode || payload.mode || '');
  if (['all', 'roles', 'users'].indexOf(mode) < 0) throw new Error('Choose one recipient mode.');
  var active = swInboxActiveWorkflowUsers_(ss);
  var roleTargets = swInboxCanonicalRoleTargets_(payload.roleTargets || payload.roles || []);
  var userTargets = swInboxCanonicalEmailTargets_(payload.userTargets || payload.users || []);
  var emails = [];
  if (mode === 'all') {
    emails = active.map(function (row) { return row.email; });
  } else if (mode === 'roles') {
    if (!roleTargets.length) throw new Error('Select at least one role.');
    active.forEach(function (row) {
      var matches = row.roles.some(function (role) {
        return roleTargets.some(function (target) { return swAuthHasRole_([role], target); });
      });
      if (matches) emails.push(row.email);
    });
  } else if (mode === 'users') {
    if (!userTargets.length) throw new Error('Select at least one user.');
    var allowed = {};
    active.forEach(function (row) { allowed[row.email] = true; });
    userTargets.forEach(function (email) { if (allowed[email]) emails.push(email); });
  }
  return {
    mode: mode,
    roleTargets: roleTargets,
    userTargets: userTargets,
    emails: swInboxUnique_(emails)
  };
}

function swInboxCanonicalRoleTargets_(roles) {
  if (!Array.isArray(roles)) roles = String(roles || '').split(/[,\n;]/);
  var allowed = {};
  swInboxRoleOptions_().forEach(function (role) {
    allowed[swAuthRoleKey_(role.value)] = role.value;
  });
  var out = [];
  roles.forEach(function (role) {
    var canonical = allowed[swAuthRoleKey_(swAuthCanonicalRole_(role))];
    if (canonical && out.indexOf(canonical) < 0) out.push(canonical);
  });
  return out;
}

function swInboxCanonicalEmailTargets_(emails) {
  if (!Array.isArray(emails)) emails = String(emails || '').split(/[,\n;]/);
  return swInboxUnique_(emails.map(swNormEmail_).filter(Boolean));
}

function swInboxRecordVisibleToUser_(rec, user) {
  rec = rec || {};
  if (rec.kind === 'schedule') return true;
  var email = swNormEmail_(user && user.email || '');
  if (!email) return false;
  var visible = Array.isArray(rec.visibleEmails) ? rec.visibleEmails : [];
  return visible.map(swNormEmail_).indexOf(email) >= 0;
}

function swInboxSortDesc_(a, b) {
  var am = swInboxDateMs_(a && a.createdAt);
  var bm = swInboxDateMs_(b && b.createdAt);
  if (am !== bm) return bm - am;
  return String(b && (b.id || b.notificationId) || '').localeCompare(String(a && (a.id || a.notificationId) || ''));
}

function swInboxDateMs_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value.getTime();
  var d = new Date(swTrim_(value || ''));
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function swInboxMonthKey_(value) {
  var d = new Date(value || '');
  if (isNaN(d.getTime())) d = new Date();
  return d.getFullYear() + '-' + swPad2_(d.getMonth() + 1);
}

function swInboxEncodeCursor_(createdAt, id) {
  var text = swStringify_({ beforeAt: createdAt || '', beforeId: id || '' });
  return Utilities.base64EncodeWebSafe(Utilities.newBlob(text).getBytes());
}

function swInboxDecodeCursor_(cursor) {
  cursor = swTrim_(cursor || '');
  if (!cursor) return null;
  try {
    var text = Utilities.newBlob(Utilities.base64DecodeWebSafe(cursor)).getDataAsString();
    var parsed = swParseJson_(text, null);
    if (!parsed) return null;
    return {
      beforeAt: parsed.beforeAt || '',
      beforeAtMs: swInboxDateMs_(parsed.beforeAt),
      beforeId: parsed.beforeId || ''
    };
  } catch (_) {}
  return null;
}

function swInboxDigest_(text) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text || ''), Utilities.Charset.UTF_8);
  var out = '';
  for (var i = 0; i < bytes.length; i++) {
    var v = bytes[i];
    if (v < 0) v += 256;
    out += ('0' + v.toString(16)).slice(-2);
  }
  return out;
}

function swInboxUnique_(values) {
  var seen = {};
  var out = [];
  (values || []).forEach(function (value) {
    value = swTrim_(value || '');
    if (!value || seen[value]) return;
    seen[value] = true;
    out.push(value);
  });
  return out;
}

function swInboxRoleLabel_(role) {
  if (role === SW_OWNER_ROLES.SALES_REP) return 'Client Advisor';
  if (role === SW_OWNER_ROLES.JOC) return 'JOC';
  if (role === SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN) return 'Diamond Order Admin';
  if (role === SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT) return 'Diamond Order Assistant';
  if (role === 'Admin') return 'Admin';
  return role || '';
}
