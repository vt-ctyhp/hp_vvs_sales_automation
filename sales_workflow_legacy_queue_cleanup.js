/** Legacy ACK/Q queue retirement utilities.
 *
 * Dry-run first:
 *   sw_dryRunCleanupLegacyQueueWorkflow()
 *
 * Apply after reviewing the log:
 *   sw_applyCleanupLegacyQueueWorkflow()
 */

var SW_LEGACY_QUEUE_RETIRED_TABS_ = [
  '06_Acknowledgement_Log',
  '08_Reps_Map',
  '09_Ack_Dashboard',
  '12_Ack_Policies',
  '13_Morning_Snapshot',
  '14_Snapshot_Log'
];

var SW_LEGACY_QUEUE_PROTECTED_TABS_ = [
  '00_Master Appointments',
  '00_Dashboard',
  '03_Client_Status_Log',
  '04_Reminders_Queue',
  '07_Root_Index',
  '10_Roster_Schedule',
  '15_Reminders_Log',
  '100_Metrics_View',
  '200_ Diamond Tracker',
  'Schedule Changes',
  'Rep Qualifications'
];

var SW_LEGACY_QUEUE_TRIGGER_HANDLERS_ = [
  'ack_runMorningFlow',
  'ack_middayQueuesRefresh',
  'ack_lateDayDashboardRefresh',
  'ack_injectRemindersForAllReps',
  'ack_installDailyTriggers',
  'ack_installHourlyDashboardTrigger',
  'buildAckDashboard',
  'buildTodaysQueuesAll',
  'buildTodaysQueuesAll_WithReminders',
  'openMyQueue',
  'refreshMyQueue',
  'refreshMyQueueHybrid',
  'submitMyQueue',
  'submitMyQueueUnified',
  'recomputeAckStatusSummary'
];

function sw_dryRunCleanupLegacyQueueWorkflow() {
  return sw_cleanupLegacyQueueWorkflow({ apply: false });
}

function sw_applyCleanupLegacyQueueWorkflow() {
  return sw_cleanupLegacyQueueWorkflow({ apply: true });
}

function sw_cleanupLegacyQueueWorkflow(options) {
  options = options || {};
  var apply = options.apply === true;
  var qPrefix = String(options.qPrefix || 'Q_');
  var ss = (typeof swSpreadsheet_ === 'function')
    ? swSpreadsheet_()
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No spreadsheet is available for legacy queue cleanup.');

  var protectedSet = swLegacyQueueNameSet_(SW_LEGACY_QUEUE_PROTECTED_TABS_);
  var retiredSet = swLegacyQueueNameSet_(SW_LEGACY_QUEUE_RETIRED_TABS_);
  var triggerSet = swLegacyQueueNameSet_(SW_LEGACY_QUEUE_TRIGGER_HANDLERS_);
  var result = {
    ok: true,
    apply: apply,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    qPrefix: qPrefix,
    protectedTabs: SW_LEGACY_QUEUE_PROTECTED_TABS_.slice(),
    candidateSheets: [],
    skippedProtectedSheets: [],
    candidateTriggers: [],
    deletedSheets: [],
    deletedTriggers: [],
    errors: []
  };

  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    var isQueueTab = qPrefix && name.indexOf(qPrefix) === 0;
    var isRetiredTab = retiredSet[name] === true;
    if (!isQueueTab && !isRetiredTab) return;

    var info = {
      name: name,
      sheetId: sheet.getSheetId(),
      reason: isQueueTab ? 'legacy Q queue tab' : 'retired ACK support tab',
      maxRows: sheet.getMaxRows(),
      maxColumns: sheet.getMaxColumns(),
      lastRow: sheet.getLastRow(),
      lastColumn: sheet.getLastColumn()
    };

    if (protectedSet[name] === true) {
      result.skippedProtectedSheets.push(info);
      return;
    }
    result.candidateSheets.push(info);
  });

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    var handler = trigger.getHandlerFunction();
    if (triggerSet[handler] !== true) return;
    result.candidateTriggers.push(swDescribeLegacyQueueTrigger_(trigger));
  });

  if (apply) {
    result.candidateTriggers.forEach(function(triggerInfo) {
      try {
        var match = swFindProjectTriggerByUid_(triggerInfo.uniqueId, triggerInfo.handler);
        if (match) {
          ScriptApp.deleteTrigger(match);
          result.deletedTriggers.push(triggerInfo);
        }
      } catch (err) {
        result.errors.push({
          type: 'trigger',
          handler: triggerInfo.handler,
          message: err && err.message ? err.message : String(err)
        });
      }
    });

    result.candidateSheets.forEach(function(sheetInfo) {
      try {
        var sheet = ss.getSheetByName(sheetInfo.name);
        if (!sheet) return;
        if (ss.getSheets().length <= 1) {
          throw new Error('Cannot delete the last remaining sheet.');
        }
        ss.deleteSheet(sheet);
        result.deletedSheets.push(sheetInfo);
      } catch (err) {
        result.errors.push({
          type: 'sheet',
          name: sheetInfo.name,
          message: err && err.message ? err.message : String(err)
        });
      }
    });
  }

  result.ok = result.errors.length === 0;
  Logger.log('SW_LEGACY_QUEUE_CLEANUP ' + JSON.stringify(result));
  return result;
}

function swLegacyQueueNameSet_(values) {
  var out = {};
  values.forEach(function(value) {
    out[String(value)] = true;
  });
  return out;
}

function swDescribeLegacyQueueTrigger_(trigger) {
  return {
    handler: trigger.getHandlerFunction(),
    uniqueId: swSafeTriggerValue_(trigger, 'getUniqueId'),
    eventType: swSafeTriggerValue_(trigger, 'getEventType'),
    source: swSafeTriggerValue_(trigger, 'getTriggerSource'),
    sourceId: swSafeTriggerValue_(trigger, 'getTriggerSourceId')
  };
}

function swFindProjectTriggerByUid_(uniqueId, handler) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var trigger = triggers[i];
    var triggerUid = swSafeTriggerValue_(trigger, 'getUniqueId');
    if (uniqueId && triggerUid === uniqueId) return trigger;
    if (!uniqueId && trigger.getHandlerFunction() === handler) return trigger;
  }
  return null;
}

function swSafeTriggerValue_(trigger, methodName) {
  try {
    if (!trigger || typeof trigger[methodName] !== 'function') return '';
    var value = trigger[methodName]();
    return value === null || value === undefined ? '' : String(value);
  } catch (err) {
    return '';
  }
}
