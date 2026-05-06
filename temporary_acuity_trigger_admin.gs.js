/** Temporary Acuity trigger ownership helpers.
 *
 * Use these from the Apps Script editor while signed in as the account that
 * should own the Acuity automation. Delete this file after the handoff is done.
 */

var TMP_ACUITY_POLL_HANDLER_ = 'acuityPollAndSubmit';
var TMP_ACUITY_LABEL_HANDLER_ = 'acuityLabelSync';
var TMP_ACUITY_TRIGGER_PROPS_ = ['ACUITY_USER_ID', 'ACUITY_API_KEY', 'FORM_ID'];

function tmp_uninstallAcuityTrigger() {
  return tmp_deleteOwnedTriggersForHandler_(TMP_ACUITY_POLL_HANDLER_);
}

function tmp_installAcuityTrigger() {
  if (typeof sw_installBackgroundOrchestratorTrigger === 'function') {
    var orchestrator = sw_installBackgroundOrchestratorTrigger();
    Logger.log('TMP_ACUITY_INSTALL_TRIGGER_ORCHESTRATED ' + JSON.stringify(orchestrator));
    return orchestrator;
  }
  var removed = tmp_deleteOwnedTriggersForHandler_(TMP_ACUITY_POLL_HANDLER_);
  var trigger = ScriptApp.newTrigger(TMP_ACUITY_POLL_HANDLER_)
    .timeBased()
    .everyMinutes(1)
    .create();
  var result = {
    ok: true,
    handler: TMP_ACUITY_POLL_HANDLER_,
    cadence: 'every 1 minute',
    removedBeforeInstall: removed.deleted,
    trigger: tmp_describeTrigger_(trigger)
  };
  Logger.log('TMP_ACUITY_INSTALL_TRIGGER ' + JSON.stringify(result));
  return result;
}

function tmp_reinstallAcuityTrigger() {
  return tmp_installAcuityTrigger();
}

function tmp_uninstallAcuityLabelSyncTrigger() {
  return tmp_deleteOwnedTriggersForHandler_(TMP_ACUITY_LABEL_HANDLER_);
}

function tmp_installAcuityLabelSyncTrigger() {
  if (typeof sw_installBackgroundOrchestratorTrigger === 'function') {
    var orchestrator = sw_installBackgroundOrchestratorTrigger();
    Logger.log('TMP_ACUITY_INSTALL_LABEL_TRIGGER_ORCHESTRATED ' + JSON.stringify(orchestrator));
    return orchestrator;
  }
  var removed = tmp_deleteOwnedTriggersForHandler_(TMP_ACUITY_LABEL_HANDLER_);
  var trigger = ScriptApp.newTrigger(TMP_ACUITY_LABEL_HANDLER_)
    .timeBased()
    .everyMinutes(1)
    .create();
  var result = {
    ok: true,
    handler: TMP_ACUITY_LABEL_HANDLER_,
    cadence: 'every 1 minute',
    removedBeforeInstall: removed.deleted,
    trigger: tmp_describeTrigger_(trigger)
  };
  Logger.log('TMP_ACUITY_INSTALL_LABEL_TRIGGER ' + JSON.stringify(result));
  return result;
}

function tmp_reinstallAcuityLabelSyncTrigger() {
  return tmp_installAcuityLabelSyncTrigger();
}

function tmp_listMyAcuityTriggers() {
  var handlers = {};
  handlers[TMP_ACUITY_POLL_HANDLER_] = true;
  handlers[TMP_ACUITY_LABEL_HANDLER_] = true;
  if (typeof SW_ORCH_HANDLER !== 'undefined') handlers[SW_ORCH_HANDLER] = true;

  var triggers = ScriptApp.getProjectTriggers()
    .filter(function(trigger) {
      return handlers[trigger.getHandlerFunction()] === true;
    })
    .map(tmp_describeTrigger_);

  var result = {
    ok: true,
    count: triggers.length,
    triggers: triggers
  };
  Logger.log('TMP_ACUITY_MY_TRIGGERS ' + JSON.stringify(result));
  return result;
}

function tmp_checkAcuityScriptProperties() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var checks = {};
  var missing = [];

  TMP_ACUITY_TRIGGER_PROPS_.forEach(function(key) {
    var value = props[key];
    var present = value !== null && value !== undefined && String(value).trim() !== '';
    if (!present) missing.push(key);
    checks[key] = {
      present: present,
      length: present ? String(value).length : 0,
      preview: present ? tmp_redactValue_(key, value) : ''
    };
  });

  var result = {
    ok: missing.length === 0,
    checkedAt: new Date().toISOString(),
    missing: missing,
    properties: checks,
    activeSpreadsheet: tmp_checkActiveSpreadsheet_(),
    formAccess: tmp_checkConfiguredForm_(props.FORM_ID)
  };
  Logger.log('TMP_ACUITY_SCRIPT_PROPERTIES_CHECK ' + JSON.stringify(result));
  return result;
}

function tmp_deleteOwnedTriggersForHandler_(handler) {
  var deleted = 0;
  var inspected = 0;
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() !== handler) return;
    inspected++;
    ScriptApp.deleteTrigger(trigger);
    deleted++;
  });
  var result = {
    ok: true,
    handler: handler,
    inspected: inspected,
    deleted: deleted,
    note: 'Only triggers owned by the current account are visible and deletable.'
  };
  Logger.log('TMP_ACUITY_DELETE_TRIGGERS ' + JSON.stringify(result));
  return result;
}

function tmp_describeTrigger_(trigger) {
  return {
    handler: trigger.getHandlerFunction(),
    eventType: tmp_safeTriggerValue_(trigger, 'getEventType'),
    source: tmp_safeTriggerValue_(trigger, 'getTriggerSource'),
    sourceId: tmp_safeTriggerValue_(trigger, 'getTriggerSourceId'),
    uniqueId: tmp_safeTriggerValue_(trigger, 'getUniqueId')
  };
}

function tmp_safeTriggerValue_(trigger, methodName) {
  try {
    if (!trigger || typeof trigger[methodName] !== 'function') return '';
    var value = trigger[methodName]();
    return value === null || value === undefined ? '' : String(value);
  } catch (err) {
    return '';
  }
}

function tmp_redactValue_(key, value) {
  var text = String(value || '');
  if (!text) return '';
  if (key === 'FORM_ID') {
    return text.length <= 12 ? text : text.slice(0, 6) + '...' + text.slice(-6);
  }
  if (text.length <= 8) return '***';
  return text.slice(0, 2) + '***' + text.slice(-4);
}

function tmp_checkActiveSpreadsheet_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { ok: false, message: 'No active spreadsheet is available.' };
    var master = ss.getSheetByName('00_Master Appointments');
    return {
      ok: !!master,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      hasMasterAppointmentsSheet: !!master
    };
  } catch (err) {
    return {
      ok: false,
      message: err && err.message ? err.message : String(err)
    };
  }
}

function tmp_checkConfiguredForm_(formId) {
  formId = String(formId || '').trim();
  if (!formId) {
    return {
      ok: false,
      message: 'FORM_ID is missing.'
    };
  }
  try {
    var form = FormApp.openById(formId);
    return {
      ok: true,
      formIdPreview: tmp_redactValue_('FORM_ID', formId),
      formTitle: form.getTitle()
    };
  } catch (err) {
    return {
      ok: false,
      formIdPreview: tmp_redactValue_('FORM_ID', formId),
      message: err && err.message ? err.message : String(err)
    };
  }
}
