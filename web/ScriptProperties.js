function listAllProperties() {
  const props = PropertiesService.getScriptProperties().getProperties();
  for (const [key, value] of Object.entries(props)) {
    console.log(`${key} = ${value}`);
  }
}
function setOneProperty() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("DEBUG", "FALSE");
}

function addOrUpdateProperty() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("CHUNK_OVERLAP_SECONDS", "3");
}

function deleteOneProperty() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("DEBUG_STRATEGIST=true");
}

/** ===== Phase 0: Feature flag + sanity checks ===== */
function setRemindersInAckFlag_(on) {
  const sp = PropertiesService.getScriptProperties();
  sp.setProperty('REMINDERS_IN_ACK', on ? 'TRUE' : 'FALSE');
  const v = sp.getProperty('REMINDERS_IN_ACK');
  Logger.log('REMINDERS_IN_ACK = %s', v);
  return v;
}
function getRemindersInAckFlag_() {
  const sp = PropertiesService.getScriptProperties();
  return /true/i.test(sp.getProperty('REMINDERS_IN_ACK') || '');
}

function phase0_checkCoreProps_() {
  const sp = PropertiesService.getScriptProperties();
  const must = ['SPREADSHEET_ID'];  // minimal set for our work today
  const missing = must.filter(k => !sp.getProperty(k));
  if (missing.length) {
    Logger.log('⚠️ Missing Script Properties: %s', missing.join(', '));
  } else {
    Logger.log('✅ Core Script Properties OK');
  }
  return { missing, all: sp.getProperties() };
}

function phase0_initFlagOff(){
  phase0_checkCoreProps_();
  setRemindersInAckFlag_(true); // keep OFF for the shadow-prep phase
}


/** ===== Phase 0: timestamped backup of core tabs ===== */
function phase0_backupTabs() {
  const ss = SpreadsheetApp.getActive();
  const tz = Session.getScriptTimeZone() || 'America/Los_Angeles';
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH.mm');

  const tabs = [
    '04_Reminders_Queue',
    '15_Reminders_Log',
    '06_Acknowledgement_Log',
    '09_Ack_Dashboard',
    '07_Root_Index',
    '08_Reps_Map',
    '10_Roster_Schedule'
  ];

  const backup = SpreadsheetApp.create(`${ss.getName()} — Phase0 Backup — ${stamp}`);
  tabs.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) { Logger.log('Skip (not found): %s', name); return; }
    const copy = sh.copyTo(backup);
    copy.setName(`${name} (backup ${stamp})`);
  });
  const sheet1 = backup.getSheetByName('Sheet1'); // default new-file tab
  if (sheet1) backup.deleteSheet(sheet1);
  const url = backup.getUrl();
  Logger.log('✅ Backup created: %s', url);
  return url;
}

/** ===== Phase 0: Trigger audit ===== */
function phase0_listKeyTriggers() {
  const want = [
    'ack_runMorningFlow',
    'ack_middayQueuesRefresh',
    'ack_lateDayDashboardRefresh',
    'Remind.remindersDailyCron',
    'Remind.remindersHourlySafetyNet'
  ];
  const have = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  Logger.log('Current triggers:\n- ' + have.join('\n- '));
  const missing = want.filter(w => !have.includes(w));
  if (missing.length) Logger.log('⚠️ Missing expected triggers: %s', missing.join(', '));
  else Logger.log('✅ All expected triggers present');
  return { have, missing };
}

/** ===== Phase 0: Header guards ===== */
function phase0_ensureRemindersHeaders() {
  const sh = SpreadsheetApp.getActive().getSheetByName('04_Reminders_Queue');
  if (!sh) { Logger.log('⚠️ Missing sheet: 04_Reminders_Queue'); return; }
  const need = [
    'id','soNumber','type','firstDueDate','nextDueAt','recurrence','status','snoozeUntil',
    'assignedRepName','assignedRepEmail','assistedRepName','assistedRepEmail',
    'customerName','nextSteps',
    'createdAt','createdBy','confirmedAt','confirmedBy','lastSentAt','attempts','lastError',
    'cancelReason','lastAdminAction','lastAdminBy'
  ];
  const have = sh.getRange(1,1,1,Math.max(sh.getLastColumn(), need.length)).getDisplayValues()[0].map(x => String(x||'').trim());
  let changed = false;
  for (let i = 0; i < need.length; i++) {
    if ((have[i]||'') !== need[i]) { sh.getRange(1, i+1).setValue(need[i]); changed = true; }
  }
  if (changed) sh.setFrozenRows(1);
  Logger.log('✅ 04_Reminders_Queue headers: OK');
}

function phase0_ensureAckLogHeaders() {
  const sh = SpreadsheetApp.getActive().getSheetByName('06_Acknowledgement_Log');
  if (!sh) { Logger.log('⚠️ Missing sheet: 06_Acknowledgement_Log'); return; }
  const lr = sh.getLastRow();
  // Only write headers if the first row is blank / not initialized
  if (lr < 1 || !String(sh.getRange(1,1).getValue() || '').trim()) {
    const need = [
      'Log Date','Timestamp','RootApptID','Rep','Role',
      'Ack Status','Ack Note','Ack By (Email/Name)',
      'Customer (at log)','Sales Stage (at log)','Conversion Status (at log)',
      'Custom Order Status (at log)','In Production Status (at log)','Center Stone Order Status (at log)',
      'Next Steps (at log)','Last Updated By (at log)','Last Updated At (at log)',
      'Client Status Report URL'
    ];
    sh.getRange(1,1,1,need.length).setValues([need]);
    sh.setFrozenRows(1);
    Logger.log('🆕 Initialized 06_Acknowledgement_Log headers.');
  } else {
    Logger.log('✅ 06_Acknowledgement_Log headers already present.');
  }
}

