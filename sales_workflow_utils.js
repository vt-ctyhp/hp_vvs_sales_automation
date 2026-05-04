/**
 * Sales workflow utilities: dates, normalization, lookup helpers, JSON, timing, and config resolution.
 */

function swVisitDateTime_(rec, tz) {
  var dateParts = swDateParts_(rec.visitDateRaw, rec.visitDate);
  if (!dateParts) return null;
  var timeParts = swTimeParts_(rec.visitTimeRaw, rec.visitTime);
  return new Date(dateParts.y, dateParts.m - 1, dateParts.d, timeParts.h, timeParts.min, 0, 0);
}

function swDateParts_(raw, display) {
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return { y: raw.getFullYear(), m: raw.getMonth() + 1, d: raw.getDate() };
  }
  var s = swTrim_(display || raw);
  if (!s) return null;
  var iso = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
  if (iso) return { y: Number(iso[1]), m: Number(iso[2]), d: Number(iso[3]) };
  var mdy = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/.exec(s);
  if (mdy) {
    var y = Number(mdy[3]);
    if (y < 100) y += 2000;
    return { y: y, m: Number(mdy[1]), d: Number(mdy[2]) };
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) return { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() };
  return null;
}

function swTimeParts_(raw, display) {
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return { h: raw.getHours(), min: raw.getMinutes() };
  }
  if (typeof raw === 'number' && isFinite(raw)) {
    var totalMinutes = Math.round((raw % 1) * 24 * 60);
    return { h: Math.floor(totalMinutes / 60) % 24, min: totalMinutes % 60 };
  }
  var s = swTrim_(display || raw);
  if (!s) return { h: 9, min: 0 };
  var m12 = /^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)$/i.exec(s);
  if (m12) {
    var h = Number(m12[1]);
    var ap = m12[3].toUpperCase();
    if (ap === 'AM' && h === 12) h = 0;
    if (ap === 'PM' && h !== 12) h += 12;
    return { h: h, min: Number(m12[2]) };
  }
  var m24 = /^(\d{1,2}):(\d{2})/.exec(s);
  if (m24) return { h: Number(m24[1]), min: Number(m24[2]) };
  return { h: 9, min: 0 };
}

function swFormatAppointmentTime_(display, raw) {
  var s = swTrim_(display);
  if (!s && raw == null) return '';
  var parsed = swParseTimeParts_(raw, s);
  if (!parsed) return s;
  var hour12 = parsed.h % 12 || 12;
  var suffix = parsed.h < 12 ? 'am' : 'pm';
  return swPad2_(hour12) + ':' + swPad2_(parsed.min) + ' ' + suffix;
}

function swParseTimeParts_(raw, display) {
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return { h: raw.getHours(), min: raw.getMinutes() };
  }
  if (typeof raw === 'number' && isFinite(raw)) {
    var totalMinutes = Math.round((raw % 1) * 24 * 60);
    return { h: Math.floor(totalMinutes / 60) % 24, min: totalMinutes % 60 };
  }
  var s = swTrim_(display || raw);
  if (!s) return null;
  var m12 = /^(\d{1,2}):(\d{2})(?::\d{2}(?:\.\d+)?)?\s*(AM|PM)$/i.exec(s);
  if (m12) {
    var hour12 = Number(m12[1]);
    if (!hour12) hour12 = 12;
    hour12 = ((hour12 - 1) % 12) + 1;
    var suffix = m12[3].toUpperCase();
    var hour = hour12;
    if (suffix === 'PM' && hour !== 12) hour += 12;
    if (suffix === 'AM' && hour === 12) hour = 0;
    return { h: hour, min: Number(m12[2]) };
  }
  var m24 = /^(\d{1,2}):(\d{2})(?::\d{2}(?:\.\d+)?)?$/.exec(s);
  if (m24) {
    var hour24 = Number(m24[1]);
    var minute = Number(m24[2]);
    if (hour24 >= 0 && hour24 <= 23 && minute >= 0 && minute <= 59) {
      return { h: hour24, min: minute };
    }
  }
  return null;
}

function swPad2_(value) {
  return String(value).length < 2 ? '0' + value : String(value);
}

function swDayOfDue_(visitAt) {
  return new Date(visitAt.getFullYear(), visitAt.getMonth(), visitAt.getDate(), 8, 0, 0, 0);
}

function swDateAddHours_(date, hours) {
  return new Date(date.getTime() + hours * 60 * 60 * 1000);
}

function swDateKey_(date) {
  if (!(date instanceof Date)) date = new Date(date);
  if (isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, swTimezone_(), 'yyyy-MM-dd');
}

function swDateValue_(iso) {
  if (!iso) return 9999999999999;
  var d = new Date(iso);
  return isNaN(d.getTime()) ? 9999999999999 : d.getTime();
}

function swSnoozeDateValue_(task) {
  if (!task || !task.snoozeUntil) return 0;
  var d = new Date(task.snoozeUntil);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function swTaskSnoozedInFuture_(task, nowMs) {
  if (!task || task.status !== SW_STATUSES.SNOOZED) return false;
  var snoozeMs = swSnoozeDateValue_(task);
  return snoozeMs && snoozeMs > (nowMs || new Date().getTime());
}

function swTaskPendingLike_(task, nowMs) {
  if (!task) return false;
  if (task.status === SW_STATUSES.PENDING) return true;
  if (task.status === SW_STATUSES.SNOOZED) return !swTaskSnoozedInFuture_(task, nowMs || new Date().getTime());
  return false;
}

function swIsOverdue_(task, nowMs) {
  if (!swTaskPendingLike_(task, nowMs)) return false;
  return swDateValue_(task.dueAt) < nowMs;
}

function swTaskDueForQueue_(task, nowMs) {
  if (!swTaskPendingLike_(task, nowMs)) return false;
  if (!task.dueAt) return true;
  return swDateValue_(task.dueAt) <= nowMs;
}

function swDueLabel_(task, nowMs) {
  if (swTaskSnoozedInFuture_(task, nowMs)) {
    var sd = new Date(task.snoozeUntil);
    return 'Snoozed until ' + Utilities.formatDate(sd, swTimezone_(), 'MMM d, h:mm a');
  }
  var t = swDateValue_(task.dueAt);
  if (t === 9999999999999) return 'No due time';
  var diff = t - nowMs;
  var mins = Math.round(Math.abs(diff) / 60000);
  if (diff < 0) {
    if (mins < 60) return 'Overdue ' + mins + 'm';
    if (mins < 1440) return 'Overdue ' + Math.round(mins / 60) + 'h';
    return 'Overdue ' + Math.round(mins / 1440) + 'd';
  }
  if (mins < 60) return 'Due in ' + mins + 'm';
  if (mins < 1440) return 'Due in ' + Math.round(mins / 60) + 'h';
  return 'Due in ' + Math.round(mins / 1440) + 'd';
}

function swHeaderMapFromArray_(headers) {
  var map = {};
  headers.forEach(function (h, i) {
    var raw = swTrim_(h);
    if (!raw) return;
    map[raw] = i;
    map[swHeaderKey_(raw)] = i;
  });
  return map;
}

function swPickIndex_(map, names) {
  for (var i = 0; i < names.length; i++) {
    if (map[names[i]] != null) return map[names[i]];
    var key = swHeaderKey_(names[i]);
    if (map[key] != null) return map[key];
  }
  return -1;
}

function swHeaderKey_(value) {
  return swTrim_(value).toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function swCell_(row, idx) {
  return idx >= 0 ? row[idx] : '';
}

function swTrim_(value) {
  return String(value == null ? '' : value).trim();
}

function swNorm_(value) {
  return swTrim_(value).toLowerCase().replace(/\s+/g, ' ');
}

function swNormEmail_(value) {
  return swTrim_(value).toLowerCase();
}

function swNormPhone_(value) {
  var d = swTrim_(value).replace(/\D+/g, '');
  if (d.length > 10 && d.charAt(0) === '1') d = d.slice(1);
  return d.length >= 7 ? d : '';
}

function swTruthy_(value) {
  var s = swNorm_(value);
  return s === 'y' || s === 'yes' || s === 'true' || s === '1' || s === 'working' || s === 'available';
}

function swUnique_(values) {
  var seen = {};
  var out = [];
  values.forEach(function (v) {
    v = swTrim_(v);
    if (!v || seen[v]) return;
    seen[v] = true;
    out.push(v);
  });
  return out;
}

function swStringify_(obj) {
  return JSON.stringify(obj || {});
}

function swTimed_(operation, fn) {
  var started = new Date().getTime();
  try {
    return fn();
  } finally {
    try {
      Logger.log('SW_TIMING ' + JSON.stringify({
        operation: operation,
        ms: new Date().getTime() - started
      }));
    } catch (_) {}
  }
}

function swStepTimer_(operation) {
  var started = new Date().getTime();
  var last = started;
  return function (step, extra) {
    var now = new Date().getTime();
    try {
      Logger.log('SW_TIMING_STEP ' + JSON.stringify({
        operation: operation,
        step: step,
        ms: now - last,
        totalMs: now - started,
        extra: extra || {}
      }));
    } catch (_) {}
    last = now;
  };
}

function swParseJson_(text, fallback) {
  if (!text) return fallback;
  try {
    var parsed = JSON.parse(text);
    return parsed == null ? fallback : parsed;
  } catch (_) {
    return fallback;
  }
}

function swDeepValue_(obj, path) {
  var cur = obj;
  for (var i = 0; i < path.length; i++) {
    if (!cur || cur[path[i]] == null) return '';
    cur = cur[path[i]];
  }
  return cur;
}

function swIso_(date) {
  return (date || new Date()).toISOString();
}

function swTimezone_() {
  try { return Session.getScriptTimeZone() || 'America/Los_Angeles'; } catch (_) {}
  return 'America/Los_Angeles';
}

function swConfigValue_(configRows, section, key, fallback) {
  var targetSection = swNorm_(section);
  var targetKey = swNorm_(key);
  for (var i = 0; i < configRows.length; i++) {
    if (swNorm_(configRows[i]['Section']) === targetSection && swNorm_(configRows[i]['Key']) === targetKey) {
      return configRows[i]['Value'] || fallback || '';
    }
  }
  return fallback || '';
}

function swMapLinkForBrand_(configRows, brand) {
  var key = swBrandConfigKey_(brand);
  return key ? swConfigValue_(configRows, 'SYSTEM', 'MAP_LINK_' + key, '') : '';
}

function swLocationMsgForBrand_(configRows, brand) {
  var key = swBrandConfigKey_(brand);
  return key ? swConfigValue_(configRows, 'SYSTEM', 'LOCATION_MSG_' + key, '') : '';
}

function swWelcomeMessageForBrand_(configRows, brand) {
  var key = swBrandConfigKey_(brand);
  return key ? swConfigValue_(configRows, 'SYSTEM', 'WELCOME_MSG_' + key, '') : '';
}

function swWelcomeImageForBrand_(configRows, brand) {
  var key = swBrandConfigKey_(brand);
  return key ? swConfigValue_(configRows, 'SYSTEM', 'WELCOME_IMAGE_' + key, '') : '';
}

function swBrandConfigKey_(brand) {
  var b = swNorm_(brand);
  if (!b) return '';
  if (/\bvvs\b/.test(b) || b.indexOf('vvs') >= 0) return 'VVS';
  if (b.indexOf('hung') >= 0 || b.indexOf('phat') >= 0 || b.indexOf('hpusa') >= 0 || b.indexOf('hp usa') >= 0) {
    return 'HUNG_PHAT';
  }
  return '';
}

function swUserHasConfigRole_(config, email, role) {
  email = swNormEmail_(email);
  if (!email) return false;
  role = swNorm_(role);
  for (var i = 0; i < config.length; i++) {
    if (swNorm_(config[i]['Role']) !== role) continue;
    if (!swTruthy_(config[i]['Active?'] || 'Y')) continue;
    if (swNormEmail_(config[i]['Email']) === email) return true;
  }
  return false;
}

function swUserMatchesRoster_(roster, name, email) {
  var n = swNorm_(name);
  var e = swNormEmail_(email);
  for (var i = 0; i < roster.length; i++) {
    if (e && swNormEmail_(roster[i].email) === e) return true;
    if (n && swNorm_(roster[i].name) === n) return true;
  }
  return false;
}

function swLookupEmailByName_(ss, name, ctx) {
  name = swNorm_(name);
  if (!name) return '';
  if (ctx && ctx.peopleIndex) return ctx.peopleIndex.emailByName[name] || '';
  var sh = ss.getSheetByName(SW_SHEETS.DROPDOWN);
  if (!sh || sh.getLastRow() < 2) return '';
  var values = sh.getDataRange().getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var pairs = [
    [swPickIndex_(H, ['Assigned Rep']), swPickIndex_(H, ['Assigned Rep Email'])],
    [swPickIndex_(H, ['Assisted Rep']), swPickIndex_(H, ['Assisted Rep Email'])]
  ];
  for (var i = 1; i < values.length; i++) {
    for (var p = 0; p < pairs.length; p++) {
      var nameCol = pairs[p][0];
      var emailCol = pairs[p][1];
      if (nameCol >= 0 && emailCol >= 0 && swNorm_(values[i][nameCol]) === name) {
        return swNormEmail_(values[i][emailCol]);
      }
    }
  }
  return '';
}

function swLookupNameByEmail_(ss, email, ctx) {
  email = swNormEmail_(email);
  if (!email) return '';
  if (ctx && ctx.peopleIndex) return ctx.peopleIndex.nameByEmail[email] || '';
  var sh = ss.getSheetByName(SW_SHEETS.DROPDOWN);
  if (sh && sh.getLastRow() >= 2) {
    var values = sh.getDataRange().getDisplayValues();
    var headers = values[0].map(function (h) { return swTrim_(h); });
    var H = swHeaderMapFromArray_(headers);
    var pairs = [
      [swPickIndex_(H, ['Assigned Rep']), swPickIndex_(H, ['Assigned Rep Email'])],
      [swPickIndex_(H, ['Assisted Rep']), swPickIndex_(H, ['Assisted Rep Email'])]
    ];
    for (var i = 1; i < values.length; i++) {
      for (var p = 0; p < pairs.length; p++) {
        var nameCol = pairs[p][0];
        var emailCol = pairs[p][1];
        if (nameCol >= 0 && emailCol >= 0 && swNormEmail_(values[i][emailCol]) === email) {
          return swTrim_(values[i][nameCol]);
        }
      }
    }
  }

  var config = swReadConfig_(ss);
  for (var c = 0; c < config.length; c++) {
    if (swNormEmail_(config[c]['Email']) === email) return swTrim_(config[c]['Name'] || config[c]['Key']);
  }
  return '';
}
