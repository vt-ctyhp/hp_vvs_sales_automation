/**
 * AutoAssign_Phase1.gs
 * Sales Appointment Auto-Assignment — Phase 1
 *
 * Không có top-level const/var.
 * Không có function onOpen (gọi addAutoAssignMenu() từ onOpen hiện có).
 * Tên hàm đủ mô tả để tránh trùng với project hiện có.
 *
 * ĐỂ KẾT NỐI: thêm addAutoAssignMenu(); vào cuối onOpen() hiện có.
 */


// ─────────────────────────────────────────────────────────────
// CONFIG — bọc trong function, không phải top-level const
// ─────────────────────────────────────────────────────────────

function autoAssignSheets_() {
  return {
    master:   '00_Master Appointments',
    roster:   '10_Roster_Schedule',
    qualif:   'Rep Qualifications',
    cache:    'Daily Availability Cache',
    log:      'Assignment Log',
    settings: 'Settings',
    dropdown: 'Dropdown',
    sysLog:   'Log',
    changes:  'Schedule Changes',
  };
}

// Cột trong 00_Master Appointments (đếm từ A=1)
function autoAssignMasterCols_() {
  return {
    APPT_ID:            1,   // A
    ASSIGNED_REP:       12,  // L
    ASSISTED_REP:       14,  // N
    ASSISTED_REP_EMAIL: 15,  // O
    CUSTOMER_NAME:      16,  // P
    ACTIVE:             26,  // Z
    STATUS:             27,  // AA
    VISIT_DATE:         29,  // AC
    VISIT_TIME:         30,  // AD
    DIAMOND_TYPE:       35,  // AI
  };
}


// ═════════════════════════════════════════════════════════════
// MENU — thêm addAutoAssignMenu() vào cuối onOpen() hiện có
// ═════════════════════════════════════════════════════════════

function addAutoAssignMenu() {
  SpreadsheetApp.getUi()
    .createMenu('🤖 Auto-Assign')
    .addItem('▶ Build Availability Cache',       'buildAvailabilityCache')
    .addItem('▶ Assign Today\'s Appointments',   'assignTodayAppointments')
    .addItem('▶ Full Daily Setup (cả hai)',       'runDailySetup')
    .addSeparator()
    .addItem('⚙ One-time: Create Tabs',          'setupAutoAssignTabs')
    .addItem('⚙ One-time: Install Triggers',     'installAutoAssignTriggers')
    .addSeparator()
    .addItem('🔍 View Today\'s Log',             'showAutoAssignLog')
    .addToUi();
}


// ═════════════════════════════════════════════════════════════
// ENTRY POINTS
// ═════════════════════════════════════════════════════════════

function runDailySetup() {
  autoAssignLog_('runDailySetup', 'START');
  var s = autoAssignGetSettings_();
  if (s['SYSTEM_ACTIVE'] === 'N') {
    autoAssignLog_('runDailySetup', 'SYSTEM_ACTIVE=N — skipped');
    return;
  }
  try {
    buildAvailabilityCache();
    assignTodayAppointments();
    sendMorningStaffingChat_();
    autoAssignLog_('runDailySetup', 'DONE');
  } catch (e) {
    autoAssignLog_('runDailySetup', 'ERROR: ' + e.message);
    alertNoRepAvailable_('⚠️ Daily setup failed:\n' + e.message);
  }
}


// ═════════════════════════════════════════════════════════════
// STEP 1 — BUILD DAILY AVAILABILITY CACHE
// ═════════════════════════════════════════════════════════════

function buildAvailabilityCache() {
  var sh    = autoAssignSheets_();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var today = new Date();
  var todayStr = toDateStr_(today);
  var dayNames = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  var todayDay = dayNames[today.getDay()];

  autoAssignLog_('buildAvailabilityCache', 'Building ' + todayStr + ' (' + todayDay + ')');

  var rosterSh = ss.getSheetByName(sh.roster);
  if (!rosterSh) throw new Error('Sheet not found: ' + sh.roster);
  var rosterData = rosterSh.getDataRange().getValues();
  var rosterHdr  = rosterData[0].map(function(h){ return h.toString().trim(); });
  var repCol = rosterHdr.indexOf('Rep');
  var dayCol = rosterHdr.indexOf(todayDay);
  if (repCol < 0) throw new Error('"Rep" column missing in Roster');
  if (dayCol < 0) throw new Error('Column "' + todayDay + '" missing in Roster');

  var rosterMap = {};
  for (var i = 1; i < rosterData.length; i++) {
    var name = rosterData[i][repCol].toString().trim();
    if (!name) continue;
    rosterMap[name] = rosterData[i][dayCol].toString().trim().toUpperCase() === 'Y';
  }

  var qualMap     = getRepQualMap_(ss, sh);
  var overrideMap = getTodayScheduleOverrides_(ss, sh, todayStr);
  var now = new Date();
  var rows = [];

  for (var repName in rosterMap) {
    var scheduled = rosterMap[repName];
    var qual      = qualMap[repName]     || {};
    var override  = overrideMap[repName] || null;

    var overrideStatus = 'Working';
    var availFrom  = '';
    var availUntil = '';
    var isAvail    = scheduled;

    if (override) {
      overrideStatus = override.changeType;
      availFrom      = override.availFrom  || '';
      availUntil     = override.availUntil || '';
      if (override.changeType === 'Full-day off') isAvail = false;
    }

    var finalAvail = scheduled && (qual.active !== false) && isAvail;
    var repEmail   = qual.email || getRepEmailFromDropdown_(ss, sh, repName);

    rows.push([
      todayStr, repName, todayDay,
      scheduled  ? 'Y' : 'N',
      overrideStatus, availFrom, availUntil,
      finalAvail ? 'Y' : 'N',
      qual.lab   ? 'Y' : 'N',
      qual.natural || 'None',
      repEmail,
      toDateTimeStr_(now),
    ]);
  }

  var cacheSh = getOrCreateTab_(ss, sh.cache);
  cacheSh.clearContents();
  var hdr = ['Cache Date','Rep Name','Day of Week','Scheduled?','Override Status',
             'Available From','Available Until','Is Available','Lab Qualified',
             'Natural Role','Rep Email','Last Updated'];
  cacheSh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#dcfce7');
  if (rows.length > 0) {
    cacheSh.getRange(2, 1, rows.length, hdr.length).setValues(rows);
    cacheSh.autoResizeColumns(1, hdr.length);
  }
  autoAssignLog_('buildAvailabilityCache', 'Done — ' + rows.length + ' reps');
}


// ═════════════════════════════════════════════════════════════
// STEP 2 — ASSIGN TODAY'S APPOINTMENTS
// ═════════════════════════════════════════════════════════════

function assignTodayAppointments() {
  var sh       = autoAssignSheets_();
  var c        = autoAssignMasterCols_();
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var todayStr = toDateStr_(new Date());
  var skipStatuses = ['Canceled','Rescheduled','Completed','No Show'];

  autoAssignLog_('assignTodayAppointments', 'Running for ' + todayStr);

  var cache = getAvailabilityCache_(ss, sh, todayStr);
  if (cache.length === 0) {
    autoAssignLog_('assignTodayAppointments', 'Cache empty — rebuilding');
    buildAvailabilityCache();
    cache = getAvailabilityCache_(ss, sh, todayStr);
  }

  var masterSh = ss.getSheetByName(sh.master);
  if (!masterSh) throw new Error('Sheet not found: ' + sh.master);
  var data = masterSh.getDataRange().getValues();

  var assigned = 0, skipped = 0, noRep = 0;

  for (var i = 1; i < data.length; i++) {
    var row    = data[i];
    var apptId = row[c.APPT_ID - 1].toString().trim();
    if (!apptId) continue;

    var active = row[c.ACTIVE  - 1].toString().trim();
    var status = row[c.STATUS  - 1].toString().trim();
    if (active !== 'Yes' && active !== 'Y') { skipped++; continue; }
    if (skipStatuses.indexOf(status) >= 0)  { skipped++; continue; }

    var visitDateRaw = row[c.VISIT_DATE - 1];
    var visitDateStr = visitDateRaw instanceof Date
      ? toDateStr_(visitDateRaw)
      : visitDateRaw.toString().trim().substring(0, 10);
    if (visitDateStr !== todayStr) { skipped++; continue; }

    if (row[c.ASSISTED_REP - 1].toString().trim()) { skipped++; continue; }

    var diamondType  = row[c.DIAMOND_TYPE  - 1].toString().trim();
    var customerName = row[c.CUSTOMER_NAME - 1].toString().trim();
    var visitTimeRaw = row[c.VISIT_TIME    - 1];
    var visitTimeStr = visitTimeRaw instanceof Date
      ? Utilities.formatDate(visitTimeRaw, Session.getScriptTimeZone(), 'HH:mm')
      : visitTimeRaw.toString().trim().substring(0, 5);

    var result = pickAssistantRep_(ss, sh, cache, diamondType, visitTimeStr, todayStr);

    if (!result) {
      var alertMsg = '⚠️ No Assisted Rep available\n'
        + 'Customer: ' + customerName + ' (' + apptId + ')\n'
        + 'Time: ' + visitTimeStr + ' | Diamond: ' + diamondType;
      autoAssignLog_('assignTodayAppointments', 'NO REP: ' + apptId);
      alertNoRepAvailable_(alertMsg);
      writeAssignmentLog_(ss, sh, {
        date: todayStr, apptId: apptId, customerName: customerName,
        visitTime: visitTimeStr, diamondType: diamondType,
        pool: 'N/A', assignedRep: '', orderNum: 0,
        assignType: 'No Rep Available', assignedBy: 'System',
        notes: 'No qualified rep — Paul alerted',
      });
      noRep++;
      continue;
    }

    masterSh.getRange(i + 1, c.ASSISTED_REP)      .setValue(result.repName);
    masterSh.getRange(i + 1, c.ASSISTED_REP_EMAIL) .setValue(result.repEmail);

    writeAssignmentLog_(ss, sh, {
      date: todayStr, apptId: apptId, customerName: customerName,
      visitTime: visitTimeStr, diamondType: diamondType,
      pool: result.pool, assignedRep: result.repName,
      orderNum: result.orderNum, assignType: 'Auto', assignedBy: 'System',
      notes: '',
    });

    autoAssignLog_('assignTodayAppointments',
      '✓ ' + result.repName + ' → ' + apptId + ' [' + diamondType + ']');
    assigned++;
  }

  var summary = 'assigned:' + assigned + ' noRep:' + noRep + ' skipped:' + skipped;
  autoAssignLog_('assignTodayAppointments', summary);
  return summary;
}


// ═════════════════════════════════════════════════════════════
// ROUND-ROBIN REP PICKER
// ═════════════════════════════════════════════════════════════

function pickAssistantRep_(ss, sh, cache, diamondType, visitTime, todayStr) {
  var isNatural = /natural/i.test(diamondType);
  var isLab     = /lab/i.test(diamondType);
  var available = cache.filter(function(r){ return r.isAvailable; });
  var pool = [], poolLabel = '';

  if (isNatural) {
    var primary = available.filter(function(r){
      return r.naturalRole === 'Primary' && isRepFreeAtTime_(r, visitTime);
    });
    if (primary.length > 0) {
      pool = primary; poolLabel = 'Natural Primary';
    } else {
      pool = available.filter(function(r){
        return r.naturalRole === 'Backup' && isRepFreeAtTime_(r, visitTime);
      });
      poolLabel = 'Natural Backup';
    }
  } else if (isLab) {
    pool = available.filter(function(r){
      return r.labQualified && isRepFreeAtTime_(r, visitTime);
    });
    poolLabel = 'Lab Pool';
  } else {
    pool = available.filter(function(r){ return isRepFreeAtTime_(r, visitTime); });
    poolLabel = 'General';
  }

  if (pool.length === 0) return null;

  var todayLog = getTodayAssignmentLog_(ss, sh, todayStr);
  var countMap = {};
  pool.forEach(function(r){ countMap[r.repName] = 0; });
  todayLog.forEach(function(entry){
    if (countMap.hasOwnProperty(entry.assignedRep)) countMap[entry.assignedRep]++;
  });

  pool.sort(function(a, b){
    var diff = (countMap[a.repName] || 0) - (countMap[b.repName] || 0);
    return diff !== 0 ? diff : a.repName.localeCompare(b.repName);
  });

  var chosen = pool[0];
  return {
    repName:  chosen.repName,
    repEmail: chosen.repEmail,
    pool:     poolLabel,
    orderNum: (countMap[chosen.repName] || 0) + 1,
  };
}

function isRepFreeAtTime_(repCache, visitTimeStr) {
  if (!repCache.isAvailable) return false;
  var visitMins = timeStrToMins_(visitTimeStr);
  if (visitMins === null) return true;
  if (repCache.availFrom) {
    var from = timeStrToMins_(repCache.availFrom);
    if (from !== null && visitMins < from) return false;
  }
  if (repCache.availUntil) {
    var until = timeStrToMins_(repCache.availUntil);
    if (until !== null && visitMins >= until) return false;
  }
  return true;
}


// ═════════════════════════════════════════════════════════════
// DATA READERS
// ═════════════════════════════════════════════════════════════

function getRepQualMap_(ss, sh) {
  var sheet = ss.getSheetByName(sh.qualif);
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};
  var h   = data[0].map(function(s){ return s.toString().trim(); });
  var col = function(k){ return h.indexOf(k); };
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var row  = data[i];
    var name = col('Rep Name') >= 0 ? row[col('Rep Name')].toString().trim() : '';
    if (!name) continue;
    map[name] = {
      email:   col('Rep Email')       >= 0 ? row[col('Rep Email')].toString().trim()       : '',
      lab:     col('Lab Diamond')     >= 0 ? row[col('Lab Diamond')].toString().trim().toUpperCase()     === 'Y' : false,
      natural: col('Natural Diamond') >= 0 ? row[col('Natural Diamond')].toString().trim() : 'None',
      active:  col('Active?')         >= 0 ? row[col('Active?')].toString().trim().toUpperCase()         !== 'N' : true,
    };
  }
  return map;
}

function getTodayScheduleOverrides_(ss, sh, todayStr) {
  var sheet = ss.getSheetByName(sh.changes);
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};
  var h   = data[0].map(function(s){ return s.toString().trim(); });
  var col = function(k){ return h.indexOf(k); };
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var row     = data[i];
    var repName = col('Rep Name') >= 0 ? row[col('Rep Name')].toString().trim() : '';
    if (!repName) continue;
    var changeDate = row[col('Change Date')];
    var ds = changeDate instanceof Date
      ? toDateStr_(changeDate)
      : changeDate.toString().trim().substring(0, 10);
    if (ds !== todayStr) continue;
    map[repName] = {
      changeType: col('Change Type')     >= 0 ? row[col('Change Type')].toString().trim()     : 'Full-day off',
      availFrom:  col('Available From')  >= 0 ? row[col('Available From')].toString().trim()  : '',
      availUntil: col('Available Until') >= 0 ? row[col('Available Until')].toString().trim() : '',
    };
  }
  return map;
}

function getAvailabilityCache_(ss, sh, todayStr) {
  var sheet = ss.getSheetByName(sh.cache);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var h   = data[0].map(function(s){ return s.toString().trim(); });
  var get = function(row, key){
    var idx = h.indexOf(key);
    return idx >= 0 ? row[idx].toString().trim() : '';
  };
  return data.slice(1).filter(function(row){
    return get(row, 'Cache Date').substring(0, 10) === todayStr;
  }).map(function(row){
    return {
      repName:      get(row, 'Rep Name'),
      repEmail:     get(row, 'Rep Email'),
      availFrom:    get(row, 'Available From'),
      availUntil:   get(row, 'Available Until'),
      isAvailable:  get(row, 'Is Available')  === 'Y',
      labQualified: get(row, 'Lab Qualified') === 'Y',
      naturalRole:  get(row, 'Natural Role'),
    };
  });
}

function getTodayAssignmentLog_(ss, sh, todayStr) {
  var sheet = ss.getSheetByName(sh.log);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getDataRange().getValues();
  var h    = data[0].map(function(s){ return s.toString().trim(); });
  var dateIdx = h.indexOf('Log Date');
  var repIdx  = h.indexOf('Assigned Rep');
  return data.slice(1).filter(function(row){
    var d  = row[dateIdx];
    var ds = d instanceof Date ? toDateStr_(d) : d.toString().trim().substring(0, 10);
    return ds === todayStr;
  }).map(function(row){
    return { assignedRep: row[repIdx] ? row[repIdx].toString().trim() : '' };
  });
}

function getRepEmailFromDropdown_(ss, sh, repName) {
  var sheet = ss.getSheetByName(sh.dropdown);
  if (!sheet) return '';
  var data  = sheet.getDataRange().getValues();
  var h     = data[0].map(function(s){ return s.toString().trim(); });
  var nameCol  = h.indexOf('Assigned Rep');
  var emailCol = h.indexOf('Assigned Rep Email');
  if (nameCol < 0 || emailCol < 0) return '';
  for (var i = 1; i < data.length; i++) {
    if (data[i][nameCol].toString().trim() === repName) {
      return data[i][emailCol].toString().trim();
    }
  }
  return '';
}

function autoAssignGetSettings_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(autoAssignSheets_().settings);
  if (!sh) return {};
  var data = sh.getDataRange().getValues();
  var map  = {};
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0].toString().trim();
    var val = data[i][1].toString().trim();
    if (key) map[key] = val;
  }
  return map;
}


// ═════════════════════════════════════════════════════════════
// DATA WRITER
// ═════════════════════════════════════════════════════════════

function writeAssignmentLog_(ss, sh, entry) {
  var sheet   = getOrCreateTab_(ss, sh.log);
  var headers = ['Log Date','APPT_ID','Customer Name','Visit Time','Diamond Type',
                 'Pool Used','Assigned Rep','Assignment Order','Assignment Type',
                 'Assigned By','Assigned At','Notes'];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#fef9c3');
  }
  sheet.appendRow([
    entry.date, entry.apptId, entry.customerName, entry.visitTime, entry.diamondType,
    entry.pool, entry.assignedRep, entry.orderNum, entry.assignType,
    entry.assignedBy, toDateTimeStr_(new Date()), entry.notes,
  ]);
}


// ═════════════════════════════════════════════════════════════
// NOTIFICATIONS
// ═════════════════════════════════════════════════════════════

function sendMorningStaffingChat_() {
  var settings = autoAssignGetSettings_();
  var webhook  = settings['GOOGLE_CHAT_WEBHOOK'];
  if (!webhook) { autoAssignLog_('sendMorningStaffingChat_', 'No webhook — skipped'); return; }

  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var sh       = autoAssignSheets_();
  var cache    = getAvailabilityCache_(ss, sh, toDateStr_(new Date()));

  var available  = cache.filter(function(r){ return r.isAvailable; });
  var labReps    = available.filter(function(r){ return r.labQualified; })             .map(function(r){ return r.repName; });
  var natPrimary = available.filter(function(r){ return r.naturalRole === 'Primary'; }).map(function(r){ return r.repName; });
  var natBackup  = available.filter(function(r){ return r.naturalRole === 'Backup';  }).map(function(r){ return r.repName; });
  var notWorking = cache.filter(function(r){ return !r.isAvailable; })                .map(function(r){ return r.repName; });

  var dayLabel = new Date().toLocaleDateString('en-US', { weekday:'long', month:'short', day:'numeric' });
  var text = '🌅 *Good morning — Staffing for ' + dayLabel + '*\n\n'
    + '✅ *Available (' + available.length + '):*  ' + (available.map(function(r){ return r.repName; }).join(', ') || 'None') + '\n'
    + '💎 *Lab Diamond:*  '       + (labReps.join(', ')    || 'None') + '\n'
    + '💍 *Natural — Primary:*  ' + (natPrimary.join(', ') || 'None') + '\n'
    + (natBackup.length  ? '💍 *Natural — Backup:*  '  + natBackup.join(', ')  + '\n' : '')
    + (notWorking.length ? '⛔ *Not working:*  '        + notWorking.join(', ') + '\n' : '')
    + '\n_Auto-assignment active._';

  postToGoogleChat_(webhook, text);
}

function alertNoRepAvailable_(message) {
  var webhook = autoAssignGetSettings_()['GOOGLE_CHAT_WEBHOOK'];
  if (!webhook) return;
  postToGoogleChat_(webhook, '🚨 *Alert — Action needed*\n' + message);
}

function postToGoogleChat_(webhookUrl, text) {
  try {
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ text: text }),
    });
  } catch (e) {
    autoAssignLog_('postToGoogleChat_', 'Failed: ' + e.message);
  }
}


// ═════════════════════════════════════════════════════════════
// ONE-TIME SETUP
// ═════════════════════════════════════════════════════════════

function setupAutoAssignTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = autoAssignSheets_();

  var qualSh = getOrCreateTab_(ss, sh.qualif);
  if (qualSh.getLastRow() === 0) {
    var qHdr = ['Rep Name','Rep Email','Lab Diamond','Natural Diamond','General Appointment','Active?','Notes'];
    qualSh.appendRow(qHdr);
    qualSh.getRange(1, 1, 1, qHdr.length).setFontWeight('bold').setBackground('#dbeafe');
    qualSh.setColumnWidths(1, qHdr.length, 150);
    prefillQualificationsTab_(ss, sh, qualSh);
  }

  var cacheSh = getOrCreateTab_(ss, sh.cache);
  if (cacheSh.getLastRow() === 0) {
    var cHdr = ['Cache Date','Rep Name','Day of Week','Scheduled?','Override Status',
                'Available From','Available Until','Is Available','Lab Qualified',
                'Natural Role','Rep Email','Last Updated'];
    cacheSh.appendRow(cHdr);
    cacheSh.getRange(1, 1, 1, cHdr.length).setFontWeight('bold').setBackground('#dcfce7');
  }

  var logSh = getOrCreateTab_(ss, sh.log);
  if (logSh.getLastRow() === 0) {
    var lHdr = ['Log Date','APPT_ID','Customer Name','Visit Time','Diamond Type',
                'Pool Used','Assigned Rep','Assignment Order','Assignment Type',
                'Assigned By','Assigned At','Notes'];
    logSh.appendRow(lHdr);
    logSh.getRange(1, 1, 1, lHdr.length).setFontWeight('bold').setBackground('#fef9c3');
  }

  var settingsSh = getOrCreateTab_(ss, sh.settings);
  if (settingsSh.getLastRow() === 0) {
    var sRows = [
      ['Setting Key',         'Value',             'Notes'],
      ['GOOGLE_CHAT_WEBHOOK', '',                  'Paste your Google Chat Space webhook URL here'],
      ['PAUL_EMAIL',          'os003@ctyhp.com',   'Receives alerts when no rep is available'],
      ['DAILY_RESET_HOUR',    '7',                 'Hour of daily setup trigger (24h)'],
      ['AUTO_REASSIGN',       'Y',                 'Y = auto-reassign on schedule change'],
      ['SYSTEM_ACTIVE',       'Y',                 'N = pause all automation (maintenance mode)'],
    ];
    settingsSh.getRange(1, 1, sRows.length, 3).setValues(sRows);
    settingsSh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#f3e8ff');
    settingsSh.setColumnWidth(1, 220);
    settingsSh.setColumnWidth(2, 300);
    settingsSh.setColumnWidth(3, 340);
  }

  SpreadsheetApp.getUi().alert(
    '✅ Setup complete!\n\n' +
    '• Rep Qualifications — điền cột Natural Diamond (None / Primary / Backup)\n' +
    '• Settings — paste Google Chat webhook URL\n\n' +
    'Sau đó chạy installAutoAssignTriggers()'
  );
}

function prefillQualificationsTab_(ss, sh, qualSh) {
  var dropSh = ss.getSheetByName(sh.dropdown);
  if (!dropSh) return;
  var data  = dropSh.getDataRange().getValues();
  var h     = data[0].map(function(s){ return s.toString().trim(); });
  var nameCol  = h.indexOf('Assigned Rep');
  var emailCol = h.indexOf('Assigned Rep Email');
  if (nameCol < 0) return;
  var seen = {};
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var name  = data[i][nameCol].toString().trim();
    var email = emailCol >= 0 ? data[i][emailCol].toString().trim() : '';
    if (!name || seen[name]) continue;
    seen[name] = true;
    rows.push([name, email, 'Y', 'None', 'Y', 'Y', '']);
  }
  if (rows.length > 0) qualSh.getRange(2, 1, rows.length, 7).setValues(rows);
}

function installAutoAssignTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    var fn = t.getHandlerFunction();
    if (fn === 'runDailySetup' || fn === 'assignTodayAppointments') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('runDailySetup')
    .timeBased().everyDays(1).atHour(7).create();
  ScriptApp.newTrigger('assignTodayAppointments')
    .timeBased().everyHours(1).create();
  autoAssignLog_('installAutoAssignTriggers', 'Installed daily 7am + hourly');
  SpreadsheetApp.getUi().alert(
    '✅ Triggers installed!\n' +
    '• runDailySetup — every day at 7am\n' +
    '• assignTodayAppointments — every hour'
  );
}

function showAutoAssignLog() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var sh       = autoAssignSheets_();
  var todayStr = toDateStr_(new Date());
  var log      = getTodayAssignmentLog_(ss, sh, todayStr);
  var msg = log.length === 0
    ? 'No assignments recorded today yet.'
    : 'Assignments today (' + todayStr + '): ' + log.length + '\n\n'
      + log.map(function(l, i){ return (i + 1) + '. ' + l.assignedRep; }).join('\n');
  SpreadsheetApp.getUi().alert('📋 Today\'s Log', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}


// ═════════════════════════════════════════════════════════════
// UTILITIES
// ═════════════════════════════════════════════════════════════

function getOrCreateTab_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); autoAssignLog_('getOrCreateTab_', 'Created: ' + name); }
  return sh;
}

function toDateStr_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function toDateTimeStr_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function timeStrToMins_(timeStr) {
  if (!timeStr) return null;
  var parts = timeStr.toString().trim().split(':');
  if (parts.length < 2) return null;
  var h = parseInt(parts[0], 10), m = parseInt(parts[1], 10);
  if (isNaN(h) || isNaN(m)) return null;
  return h * 60 + m;
}

function autoAssignLog_(fn, msg) {
  var logSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(autoAssignSheets_().sysLog);
  if (logSh) logSh.appendRow([toDateTimeStr_(new Date()), fn, msg]);
  console.log('[' + fn + '] ' + msg);
}

function createRosterSchedule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('10_Roster_Schedule');
  if (sh) {
    SpreadsheetApp.getUi().alert('Tab 10_Roster_Schedule đã tồn tại rồi.');
    return;
  }

  sh = ss.insertSheet('10_Roster_Schedule');

  // Headers
  var headers = ['Rep', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun',
                 'Assisted Coverage Enabled?', 'Assisted Coverage Partner'];
  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#dbeafe');

  // Pre-fill reps từ Dropdown tab
  var dropSh = ss.getSheetByName('Dropdown');
  var rows = [];
  if (dropSh) {
    var data  = dropSh.getDataRange().getValues();
    var h     = data[0].map(function(s){ return s.toString().trim(); });
    var nameCol = h.indexOf('Assigned Rep');
    var seen  = {};
    for (var i = 1; i < data.length; i++) {
      var name = nameCol >= 0 ? data[i][nameCol].toString().trim() : '';
      if (!name || seen[name]) continue;
      seen[name] = true;
      // Default: làm việc Mon-Fri, nghỉ Sat-Sun
      rows.push([name, 'Y','Y','Y','Y','Y','N','N', 'Y', '']);
    }
  }

  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sh.autoResizeColumns(1, headers.length);
  SpreadsheetApp.getUi().alert(
    '✅ Tạo xong 10_Roster_Schedule!\n\n' +
    'Vào tab đó và chỉnh lại:\n' +
    '• Y = làm việc ngày đó\n' +
    '• N = nghỉ ngày đó'
  );
}