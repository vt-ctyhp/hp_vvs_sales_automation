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
  '03_Client_Status_Log',
  '04_Reminders_Queue',
  '07_Root_Index',
  '10_Roster_Schedule',
  '15_Reminders_Log',
  '200_ Diamond Tracker',
  'Schedule Changes'
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

var SW_LEGACY_APPOINTMENT_TRIGGER_HANDLERS_ = [
  'processUploadQueue',
  'processSummariesWorker',
  'processIntakeQueue',
  'ensureBootstrapForRecentRows_'
];

var SW_LEGACY_APPOINTMENT_RETIRED_TABS_ = [
  '_upload_queue'
];

var SW_LEGACY_APPOINTMENT_RETAINED_TABS_ = [
  '_AppointmentArtifacts',
  '_IntakeQueue'
];

var SW_LEGACY_SHEET_DASHBOARD_TABS_ = [
  '00_Dashboard',
  '02_Client by Stage',
  '02_Clients by Stage',
  '100_Metrics_View',
  '99_3D_Status_Map',
  'Drill_KPI',
  'Drill_Unified',
  'Debug_Cohort2nd',
  'Audit_Orphans'
];

var SW_LEGACY_SHEET_DASHBOARD_TRIGGER_HANDLERS_ = [
  'refreshDashboardHourly',
  'runOnceToBuildAll',
  'buildMetricsView_',
  'writeDashboard_',
  'snapshotKpisForHistory_',
  'timedRefreshHandler',
  'buildUnifiedDrillDown_',
  'runBuildUnifiedDrillDown',
  'P15_alertOnHoldOrders'
];

function sw_dryRunCleanupLegacyAppointmentAutomation() {
  return sw_cleanupLegacyAppointmentAutomation({ apply: false });
}

function sw_applyCleanupLegacyAppointmentAutomation() {
  return sw_cleanupLegacyAppointmentAutomation({
    apply: true,
    deleteRetiredSheets: true,
    installHourlyRepair: true
  });
}

function sw_retireLegacyAppointmentTrigger_(handlerName) {
  return sw_cleanupLegacyAppointmentAutomation({
    apply: true,
    handlerOnly: handlerName,
    deleteRetiredSheets: false,
    installHourlyRepair: handlerName === 'ensureBootstrapForRecentRows_'
  });
}

function sw_cleanupLegacyAppointmentAutomation(options) {
  options = options || {};
  var apply = options.apply === true;
  var handlerOnly = String(options.handlerOnly || '');
  var deleteRetiredSheets = options.deleteRetiredSheets === true;
  var installHourlyRepair = options.installHourlyRepair === true;
  var ss = (typeof swSpreadsheet_ === 'function')
    ? swSpreadsheet_()
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No spreadsheet is available for appointment automation cleanup.');

  var retiredTriggerSet = swLegacyQueueNameSet_(SW_LEGACY_APPOINTMENT_TRIGGER_HANDLERS_);
  var retiredTabSet = swLegacyQueueNameSet_(SW_LEGACY_APPOINTMENT_RETIRED_TABS_);
  var result = {
    ok: true,
    apply: apply,
    handlerOnly: handlerOnly,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    candidateTriggers: [],
    deletedTriggers: [],
    candidateSheets: [],
    deletedSheets: [],
    retainedSheets: [],
    installedRepairTrigger: false,
    errors: []
  };

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    var handler = trigger.getHandlerFunction();
    var isRetired = retiredTriggerSet[handler] === true;
    var isRepair = installHourlyRepair && handler === 'repairMissingUrls_';
    if (handlerOnly) {
      if (handler !== handlerOnly && !isRepair) return;
    } else if (!isRetired && !isRepair) {
      return;
    }
    result.candidateTriggers.push(swDescribeLegacyQueueTrigger_(trigger));
  });

  if (!handlerOnly && deleteRetiredSheets) {
    ss.getSheets().forEach(function(sheet) {
      var name = sheet.getName();
      if (retiredTabSet[name] === true) {
        result.candidateSheets.push({
          name: name,
          sheetId: sheet.getSheetId(),
          reason: 'retired appointment upload queue',
          lastRow: sheet.getLastRow(),
          lastColumn: sheet.getLastColumn()
        });
      }
    });
  }

  SW_LEGACY_APPOINTMENT_RETAINED_TABS_.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      result.retainedSheets.push({
        name: name,
        reason: name === '_IntakeQueue'
          ? 'current iPad handoff queue used by ipad_runIntakeNow'
          : 'current appointment artifact workflow table',
        lastRow: sheet.getLastRow(),
        lastColumn: sheet.getLastColumn()
      });
    }
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
        if (ss.getSheets().length <= 1) throw new Error('Cannot delete the last remaining sheet.');
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

    if (installHourlyRepair && typeof repairMissingUrls_ === 'function') {
      try {
        var hasRepair = ScriptApp.getProjectTriggers().some(function(trigger) {
          return trigger.getHandlerFunction() === 'repairMissingUrls_';
        });
        if (!hasRepair) {
          ScriptApp.newTrigger('repairMissingUrls_').timeBased().everyHours(1).create();
          result.installedRepairTrigger = true;
        }
      } catch (err) {
        result.errors.push({
          type: 'trigger',
          handler: 'repairMissingUrls_',
          message: err && err.message ? err.message : String(err)
        });
      }
    }
  }

  result.ok = result.errors.length === 0;
  Logger.log('SW_LEGACY_APPOINTMENT_AUTOMATION_CLEANUP ' + JSON.stringify(result));
  return result;
}

function sw_dryRunCleanupLegacySheetDashboards() {
  return sw_cleanupLegacySheetDashboards({ apply: false });
}

function sw_applyCleanupLegacySheetDashboards() {
  return sw_cleanupLegacySheetDashboards({ apply: true });
}

function sw_cleanupLegacySheetDashboards(options) {
  options = options || {};
  var apply = options.apply === true;
  var ss = (typeof swSpreadsheet_ === 'function')
    ? swSpreadsheet_()
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No spreadsheet is available for legacy sheet dashboard cleanup.');

  var retiredSet = swLegacyQueueNameSet_(SW_LEGACY_SHEET_DASHBOARD_TABS_);
  var retiredTriggerSet = swLegacyQueueNameSet_(SW_LEGACY_SHEET_DASHBOARD_TRIGGER_HANDLERS_);
  var result = {
    ok: true,
    apply: apply,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    candidateSheets: [],
    deletedSheets: [],
    candidateTriggers: [],
    deletedTriggers: [],
    skippedSheets: [],
    errors: []
  };

  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (retiredSet[name] !== true) return;
    result.candidateSheets.push({
      name: name,
      sheetId: sheet.getSheetId(),
      lastRow: sheet.getLastRow(),
      lastColumn: sheet.getLastColumn()
    });
  });

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    var handler = trigger.getHandlerFunction();
    if (retiredTriggerSet[handler] !== true) return;
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
          result.skippedSheets.push({
            name: sheetInfo.name,
            reason: 'cannot delete the last remaining sheet'
          });
          return;
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
  Logger.log('SW_LEGACY_SHEET_DASHBOARD_CLEANUP ' + JSON.stringify(result));
  return result;
}

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
