/**
 * Recovery backup utilities.
 *
 * These functions are intentionally manual/menu-driven. They do not run from
 * triggers and they do not change production sheets except when explicitly
 * creating a backup copy or restoring properties in the file where they run.
 */

const RECOVERY_PROPS_SHEET = '__RECOVERY_SCRIPT_PROPERTIES';
const RECOVERY_MANIFEST_SHEET = '__RECOVERY_BACKUP_MANIFEST';
const RECOVERY_TRIGGERS_SHEET = '__RECOVERY_TRIGGER_MANIFEST';
const RECOVERY_FOLDER_NAME = 'VVS Recovery Backups';

function recovery_downloadScriptProperties() {
  if (!recovery_confirmSensitiveBackup_('Download Script Properties JSON')) return null;

  const payload = recovery_buildBackupPayload_({
    sourceSpreadsheet: recovery_getActiveSpreadsheet_(),
    includeTriggers: true
  });
  const json = JSON.stringify(payload, null, 2);
  recovery_showJsonDownloadDialog_(payload.fileName, json);
  return payload;
}

function recovery_createScriptPropertiesDriveBackup() {
  if (!recovery_confirmSensitiveBackup_('Create Script Properties Drive Backup')) return null;

  const ss = recovery_getActiveSpreadsheet_();
  const sourceFile = DriveApp.getFileById(ss.getId());
  const folder = recovery_getOrCreateBackupFolder_(sourceFile);
  const payload = recovery_buildBackupPayload_({
    sourceSpreadsheet: ss,
    includeTriggers: true
  });

  const json = JSON.stringify(payload, null, 2);
  const file = folder.createFile(payload.fileName, json, 'application/json');
  const result = {
    ok: true,
    backupFileId: file.getId(),
    backupFileUrl: file.getUrl(),
    folderUrl: folder.getUrl()
  };
  Logger.log('Script Properties backup created: %s', file.getUrl());
  recovery_alert_('Script Properties backup created.\n\n' + file.getUrl());
  return result;
}

function recovery_createFullRecoveryCopy() {
  if (!recovery_confirmSensitiveBackup_('Create Full Spreadsheet Recovery Copy')) return null;

  const source = recovery_getActiveSpreadsheet_();
  const sourceFile = DriveApp.getFileById(source.getId());
  const folder = recovery_getOrCreateBackupFolder_(sourceFile);
  const stamp = recovery_timestamp_();
  const copyName = recovery_safeFileName_(source.getName() + ' - RECOVERY COPY - ' + stamp);

  const copyFile = sourceFile.makeCopy(copyName, folder);
  const copy = SpreadsheetApp.openById(copyFile.getId());

  const payload = recovery_buildBackupPayload_({
    sourceSpreadsheet: source,
    copiedSpreadsheet: copy,
    includeTriggers: true
  });
  recovery_writeBackupSheets_(copy, payload);

  const jsonFile = folder.createFile(payload.fileName, JSON.stringify(payload, null, 2), 'application/json');
  const result = {
    ok: true,
    sourceSpreadsheetId: source.getId(),
    recoveryCopyId: copy.getId(),
    recoveryCopyUrl: copy.getUrl(),
    backupJsonFileId: jsonFile.getId(),
    backupJsonFileUrl: jsonFile.getUrl(),
    folderUrl: folder.getUrl()
  };

  Logger.log('Recovery copy created: %s', copy.getUrl());
  Logger.log('Recovery JSON created: %s', jsonFile.getUrl());
  recovery_alert_(
    'Recovery copy created.\n\n' +
    'Copy:\n' + copy.getUrl() + '\n\n' +
    'Properties JSON:\n' + jsonFile.getUrl() + '\n\n' +
    'Open the copy and run recovery_restoreScriptPropertiesFromBackupSheet() before using it as a replacement.'
  );
  return result;
}

function recovery_restoreScriptPropertiesFromBackupSheet() {
  return recovery_restoreScriptPropertiesFromBackupSheet_(false);
}

function recovery_restoreExactScriptPropertiesFromBackupSheet() {
  return recovery_restoreScriptPropertiesFromBackupSheet_(true);
}

function recovery_restoreScriptPropertiesFromBackupSheet_(useExactOriginalValues) {
  const ss = recovery_getActiveSpreadsheet_();
  const sh = ss.getSheetByName(RECOVERY_PROPS_SHEET);
  if (!sh) throw new Error('Missing backup sheet: ' + RECOVERY_PROPS_SHEET);

  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) throw new Error(RECOVERY_PROPS_SHEET + ' has no backed-up properties.');

  const headers = values[0].map(function(h){ return String(h || '').trim(); });
  const keyCol = headers.indexOf('Key');
  const originalCol = headers.indexOf('Original Value');
  const safeCol = headers.indexOf('Copy-Safe Restore Value');
  if (keyCol < 0 || originalCol < 0 || safeCol < 0) {
    throw new Error(RECOVERY_PROPS_SHEET + ' is missing required columns.');
  }

  const label = useExactOriginalValues
    ? 'Restore EXACT original Script Properties'
    : 'Restore copy-safe Script Properties';
  if (!recovery_confirmSensitiveBackup_(label)) return null;

  const props = PropertiesService.getScriptProperties();
  let restored = 0;
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][keyCol] || '').trim();
    if (!key) continue;
    const value = useExactOriginalValues ? values[i][originalCol] : values[i][safeCol];
    props.setProperty(key, String(value || ''));
    restored++;
  }

  const result = {
    ok: true,
    restored: restored,
    spreadsheetId: ss.getId(),
    mode: useExactOriginalValues ? 'exact-original' : 'copy-safe'
  };
  Logger.log('Restored %s Script Properties using mode=%s', restored, result.mode);
  recovery_alert_('Restored ' + restored + ' Script Properties.\n\nMode: ' + result.mode);
  return result;
}

function recovery_buildBackupPayload_(opts) {
  opts = opts || {};
  const source = opts.sourceSpreadsheet || recovery_getActiveSpreadsheet_();
  const copied = opts.copiedSpreadsheet || null;
  const props = PropertiesService.getScriptProperties().getProperties();
  const keys = Object.keys(props).sort();
  const timestamp = new Date();
  const stamp = recovery_timestamp_(timestamp);

  const sourceInfo = recovery_spreadsheetInfo_(source);
  const copiedInfo = copied ? recovery_spreadsheetInfo_(copied) : null;
  const payload = {
    schema: 'vvs-recovery-backup/v1',
    createdAt: timestamp.toISOString(),
    createdAtLocal: Utilities.formatDate(timestamp, source.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    createdBy: recovery_activeUserEmail_(),
    scriptId: recovery_scriptId_(),
    sourceSpreadsheet: sourceInfo,
    copiedSpreadsheet: copiedInfo,
    scriptProperties: {},
    scriptPropertiesCopySafe: {},
    triggers: opts.includeTriggers ? recovery_triggerManifest_() : [],
    notes: [
      'Spreadsheet Drive copies should carry the container-bound Apps Script code.',
      'Script Properties and installable triggers are backed up separately here.',
      'For a recovery copy, restore copy-safe properties first so SPREADSHEET_ID points to the copy, not the original.',
      'Reinstall time triggers from the copied project after restore.'
    ]
  };

  keys.forEach(function(key){
    const value = props[key];
    payload.scriptProperties[key] = value;
    payload.scriptPropertiesCopySafe[key] =
      key === 'SPREADSHEET_ID' && copied ? copied.getId() : value;
  });

  payload.fileName = recovery_safeFileName_(
    source.getName() + ' - script-properties-backup - ' + stamp + '.json'
  );
  return payload;
}

function recovery_writeBackupSheets_(copy, payload) {
  const propRows = [[
    'Key',
    'Original Value',
    'Copy-Safe Restore Value',
    'Note'
  ]];
  Object.keys(payload.scriptProperties).sort().forEach(function(key){
    const note = key === 'SPREADSHEET_ID'
      ? 'Copy-safe value points to the recovery copy. Original value is preserved for audit.'
      : '';
    propRows.push([
      key,
      payload.scriptProperties[key],
      payload.scriptPropertiesCopySafe[key],
      note
    ]);
  });

  const manifestRows = [
    ['Field', 'Value'],
    ['Schema', payload.schema],
    ['Created At', payload.createdAt],
    ['Created By', payload.createdBy],
    ['Source Spreadsheet ID', payload.sourceSpreadsheet.id],
    ['Source Spreadsheet URL', payload.sourceSpreadsheet.url],
    ['Recovery Copy ID', payload.copiedSpreadsheet ? payload.copiedSpreadsheet.id : ''],
    ['Recovery Copy URL', payload.copiedSpreadsheet ? payload.copiedSpreadsheet.url : ''],
    ['Source Script ID', payload.scriptId],
    ['Restore Copy-Safe Function', 'recovery_restoreScriptPropertiesFromBackupSheet'],
    ['Restore Exact Original Function', 'recovery_restoreExactScriptPropertiesFromBackupSheet'],
    ['Trigger Reinstall Note', 'Run the project trigger installers after restoring properties.']
  ];

  const triggerRows = [[
    'Handler Function',
    'Event Type',
    'Trigger Source',
    'Trigger Source ID',
    'Unique ID'
  ]];
  payload.triggers.forEach(function(t){
    triggerRows.push([
      t.handlerFunction,
      t.eventType,
      t.triggerSource,
      t.triggerSourceId,
      t.uniqueId
    ]);
  });

  recovery_writeHiddenSheet_(copy, RECOVERY_PROPS_SHEET, propRows);
  recovery_writeHiddenSheet_(copy, RECOVERY_MANIFEST_SHEET, manifestRows);
  recovery_writeHiddenSheet_(copy, RECOVERY_TRIGGERS_SHEET, triggerRows);
}

function recovery_writeHiddenSheet_(ss, name, rows) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();
  if (rows.length && rows[0].length) {
    sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, rows[0].length);
  }
  sh.hideSheet();
  return sh;
}

function recovery_showJsonDownloadDialog_(fileName, json) {
  const html = HtmlService.createHtmlOutput(
    '<!doctype html><html><head><base target="_top">' +
    '<style>body{font-family:Arial,sans-serif;margin:16px}button{margin:0 0 12px 0;padding:8px 12px}textarea{width:100%;height:360px;font-family:monospace;font-size:12px}</style>' +
    '</head><body>' +
    '<button onclick="download()">Download JSON</button>' +
    '<textarea readonly id="json"></textarea>' +
    '<script>' +
    'const fileName=' + JSON.stringify(fileName) + ';' +
    'const content=' + JSON.stringify(json) + ';' +
    'document.getElementById("json").value=content;' +
    'function download(){const blob=new Blob([content],{type:"application/json"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=fileName;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove();},500);}' +
    '</script></body></html>'
  ).setWidth(720).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Script Properties Backup');
}

function recovery_getActiveSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet is available for recovery backup.');
  return ss;
}

function recovery_getOrCreateBackupFolder_(sourceFile) {
  let parent = null;
  try {
    const parents = sourceFile.getParents();
    parent = parents.hasNext() ? parents.next() : null;
  } catch (_) {}
  parent = parent || DriveApp.getRootFolder();

  const existing = parent.getFoldersByName(RECOVERY_FOLDER_NAME);
  return existing.hasNext() ? existing.next() : parent.createFolder(RECOVERY_FOLDER_NAME);
}

function recovery_spreadsheetInfo_(ss) {
  return {
    id: ss.getId(),
    name: ss.getName(),
    url: ss.getUrl(),
    timeZone: ss.getSpreadsheetTimeZone(),
    activeSheetName: ss.getActiveSheet() ? ss.getActiveSheet().getName() : ''
  };
}

function recovery_triggerManifest_() {
  return ScriptApp.getProjectTriggers().map(function(t){
    return {
      handlerFunction: recovery_triggerValue_(t, 'getHandlerFunction'),
      eventType: recovery_triggerValue_(t, 'getEventType'),
      triggerSource: recovery_triggerValue_(t, 'getTriggerSource'),
      triggerSourceId: recovery_triggerValue_(t, 'getTriggerSourceId'),
      uniqueId: recovery_triggerValue_(t, 'getUniqueId')
    };
  });
}

function recovery_triggerValue_(trigger, methodName) {
  try {
    return trigger && typeof trigger[methodName] === 'function'
      ? String(trigger[methodName]() || '')
      : '';
  } catch (_) {
    return '';
  }
}

function recovery_scriptId_() {
  try {
    return ScriptApp.getScriptId ? ScriptApp.getScriptId() : '';
  } catch (_) {
    return '';
  }
}

function recovery_activeUserEmail_() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (_) {
    return '';
  }
}

function recovery_timestamp_(date) {
  const d = date || new Date();
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'America/Los_Angeles', 'yyyy-MM-dd HH.mm.ss');
}

function recovery_safeFileName_(name) {
  return String(name || 'recovery-backup')
    .replace(/[\\/:*?"<>|#%{}~&]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function recovery_confirmSensitiveBackup_(actionLabel) {
  try {
    const ui = SpreadsheetApp.getUi();
    const button = ui.alert(
      actionLabel,
      'This includes Script Properties and may include private IDs, API keys, webhook URLs, or other secrets. Continue?',
      ui.ButtonSet.YES_NO
    );
    return button === ui.Button.YES;
  } catch (_) {
    return true;
  }
}

function recovery_alert_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (_) {
    Logger.log(message);
  }
}
