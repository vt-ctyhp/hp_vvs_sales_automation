const LABEL_STATUS_MAP_ = {
  'completed' : 'Completed',
  'confirmed' : 'Confirmed',
  'no-show'   : 'No-Show',
  'canceled'  : 'Canceled',
};

function installLabelSyncTrigger() {
  const FN = 'acuityLabelSync';
  // Xóa trigger cũ nếu có
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === FN)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger(FN).timeBased().everyMinutes(1).create();
  Logger.log('✅ Label sync trigger installed: every 5 minutes');
}

function uninstallLabelSyncTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'acuityLabelSync')
    .forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('🗑️ Label sync trigger removed');
}

function acuityLabelSync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('acuityLabelSync: locked, skipping');
    return;
  }

  try {
    const SP     = PropertiesService.getScriptProperties();
    const userId = SP.getProperty('ACUITY_USER_ID');
    const apiKey = SP.getProperty('ACUITY_API_KEY');
    if (!userId || !apiKey) throw new Error('Missing credentials');

    const now    = new Date();
    const future = new Date(now.getTime() + 30 * 24 * 3600 * 1000);
    const minDate = Utilities.formatDate(now,    'UTC', "yyyy-MM-dd'T'00:00:00'Z'");
    const maxDate = Utilities.formatDate(future, 'UTC', "yyyy-MM-dd'T'23:59:59'Z'");

    const auth = 'Basic ' + Utilities.base64Encode(userId + ':' + apiKey);
    const url  = ACUITY_CFG.BASE_URL
               + '/appointments'
               + '?minDate=' + encodeURIComponent(minDate)
               + '&maxDate=' + encodeURIComponent(maxDate)
               + '&max=300';

    const resp = UrlFetchApp.fetch(url, {
      method: 'get', headers: { Authorization: auth }, muteHttpExceptions: true,
    });
    if (resp.getResponseCode() !== 200) {
      Logger.log('❌ API error: ' + resp.getResponseCode());
      return;
    }

    const appointments = JSON.parse(resp.getContentText());
    Logger.log('Fetched ' + appointments.length + ' appointments');
    if (!appointments.length) return;

    const ss  = SpreadsheetApp.getActive();
    const sh  = ss.getSheetByName('00_Master Appointments');
    if (!sh) { Logger.log('❌ Sheet not found'); return; }

    const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const H   = {};
    hdr.forEach((h, i) => { if (h) H[String(h).trim()] = i + 1; });
    if (!H['CalendlyEventUID'] || !H['Status']) { Logger.log('❌ Missing columns'); return; }

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    // Batch read toàn bộ UID + Status
    const uidCol    = sh.getRange(2, H['CalendlyEventUID'], lastRow - 1, 1).getValues();
    const statusCol = sh.getRange(2, H['Status'],           lastRow - 1, 1).getValues();

    // Build map — giữ TẤT CẢ rows, kể cả Canceled
    const masterMap = {};
    for (let i = 0; i < uidCol.length; i++) {
      const uid = String(uidCol[i][0]).trim();
      const st  = String(statusCol[i][0] || '').trim();
      if (!uid) continue;
      // Nếu uid đã có → giữ dòng mới nhất (row lớn hơn)
      if (!masterMap[uid] || (i + 2) > masterMap[uid].row) {
        masterMap[uid] = { row: i + 2, status: st };
      }
    }

    let updated = 0;

    for (const appt of appointments) {
      const uid       = String(appt.id || '').trim();
      const newStatus = labelToStatus_(appt);
      if (!newStatus) continue;

      // Tìm row — ưu tiên _R reschedule mới nhất
      const entry = findMasterEntry_(uid, masterMap);
      if (!entry) {
        Logger.log('⚠️ Not found on sheet: uid=' + uid);
        continue;
      }

      const { row, status: curStatus } = entry;
      Logger.log('Check uid=' + uid + ' row=' + row + ' curStatus="' + curStatus + '" newStatus="' + newStatus + '"');

      // Chỉ block Rescheduled — mọi status khác đều cho update
      if (/rescheduled/i.test(curStatus)) {
        Logger.log('⏭️ Skip Rescheduled');
        continue;
      }
      if (curStatus === newStatus) {
        Logger.log('⏭️ Already ' + newStatus);
        continue;
      }

      // ✅ Update
      sh.getRange(row, H['Status']).setValue(newStatus);

      if (H['Automation Notes']) {
        const prev = sh.getRange(row, H['Automation Notes']).getValue() || '';
        const note = '[Label Sync] ' + curStatus + ' → ' + newStatus
                   + ' @ ' + Utilities.formatDate(new Date(), ACUITY_CFG.TZ, 'yyyy-MM-dd HH:mm:ss');
        sh.getRange(row, H['Automation Notes']).setValue(prev ? prev + '\n' + note : note);
      }

      updated++;
      Logger.log('🏷️ UPDATED row=' + row + ' uid=' + uid + ' | ' + curStatus + ' → ' + newStatus);
    }

    if (updated > 0) CacheService.getScriptCache().remove('MASTER_UIDS_CACHE');
    Logger.log('✅ Done — checked=' + appointments.length + ' updated=' + updated);

  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function labelToStatus_(appt) {
  for (const label of (appt.labels || [])) {
    const name  = String(label.name || '').trim().toLowerCase();
    const match = LABEL_STATUS_MAP_[name];
    if (match) return match;
  }
  if (appt.noShow === true) return 'No Show';
  return '';
}

function findMasterEntry_(uid, masterMap) {
  // Pass 1: tìm _R mới nhất (reschedule) — không bị Rescheduled
  let best = null;
  for (const key of Object.keys(masterMap)) {
    if (!key.startsWith(uid + '_R')) continue;
    const entry = masterMap[key];
    if (/rescheduled/i.test(entry.status)) continue; // ← bỏ check canceled
    if (!best || entry.row > best.row) best = entry;
  }
  if (best) return best;

  // Pass 2: dòng gốc — kể cả Canceled
  const orig = masterMap[uid];
  if (orig && !/rescheduled/i.test(orig.status)) return orig; // ← bỏ check canceled
  return null;
}

function debugLabelSyncDetail() {
  const SP     = PropertiesService.getScriptProperties();
  const userId = SP.getProperty('ACUITY_USER_ID');
  const apiKey = SP.getProperty('ACUITY_API_KEY');
  const auth   = 'Basic ' + Utilities.base64Encode(userId + ':' + apiKey);

  const TARGET_ID = '1686838998'; // ← ID mới

  const resp = UrlFetchApp.fetch(
    ACUITY_CFG.BASE_URL + '/appointments/' + TARGET_ID,
    { method: 'get', headers: { Authorization: auth }, muteHttpExceptions: true }
  );
  const appt = JSON.parse(resp.getContentText());

  Logger.log('=== LABEL INFO ===');
  (appt.labels || []).forEach(l => {
    Logger.log('Label name RAW: "' + l.name + '"');
    Logger.log('Label lowercase: "' + l.name.toLowerCase() + '"');
    Logger.log('Map match: "' + (LABEL_STATUS_MAP_[l.name.toLowerCase()] || 'NOT FOUND ❌') + '"');
  });

  Logger.log('');
  Logger.log('=== MASTER SHEET LOOKUP ===');

  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getSheetByName('00_Master Appointments');
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const H   = {};
  hdr.forEach((h, i) => { if (h) H[String(h).trim()] = i + 1; });

  const lastRow = sh.getLastRow();
  const uids    = sh.getRange(2, H['CalendlyEventUID'], lastRow - 1, 1).getValues();

  let found = false;
  for (let i = 0; i < uids.length; i++) {
    const uid = String(uids[i][0]).trim();
    if (!uid.includes(TARGET_ID)) continue;
    found = true;
    const row    = i + 2;
    const status = String(sh.getRange(row, H['Status']).getValue() || '').trim();
    Logger.log('✅ Found row=' + row + ' uid="' + uid + '" status="' + status + '"');
  }

  if (!found) {
    Logger.log('❌ UID ' + TARGET_ID + ' NOT FOUND trên Master Sheet!');
    Logger.log('→ Appointment test này chưa được submit vào sheet');
    Logger.log('→ Hãy test với 1 appointment THẬT đã có trên sheet');
  }
}

function rp_debugFindOnEditWalkinSlides() {
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log('=== ALL TRIGGERS (%s) ===', triggers.length);
  
  triggers.forEach(t => {
    try {
      Logger.log('Function: %s | EventType: %s | Source: %s',
        t.getHandlerFunction(),
        t.getEventType(),
        t.getTriggerSource()
      );
    } catch(e) {
      Logger.log('Function: %s | Error reading trigger: %s',
        t.getHandlerFunction(), e.message);
    }
  });

  // Tìm onEdit_WalkinSlides
  Logger.log('\n=== FIND onEdit_WalkinSlides ===');
  const found = triggers.filter(t => t.getHandlerFunction() === 'onEdit_WalkinSlides');
  Logger.log(found.length ? '✅ Found: ' + found.length + ' trigger(s)' : '❌ Not found');

  // Thử chạy thẳng hàm để xem lỗi
  Logger.log('\n=== TEST RUN ===');
  try {
    onEdit_WalkinSlides({ 
      range: SpreadsheetApp.getActive().getActiveSheet().getActiveRange(),
      source: SpreadsheetApp.getActive()
    });
    Logger.log('✅ Run OK');
  } catch(e) {
    Logger.log('❌ Error: ' + e.message);
    Logger.log('Stack: ' + e.stack);
  }
}