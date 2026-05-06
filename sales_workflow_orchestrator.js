/**
 * Single background trigger orchestrator.
 *
 * This consolidates time-based background jobs without changing event handlers:
 * onFormSubmit and onOpen remain event-driven. Script Properties store only
 * fixed-size control state that is overwritten each run.
 */

var SW_ORCH_HANDLER = 'sw_backgroundOrchestrator';
var SW_ORCH_LEASE_KEY = 'SW_ORCH_LEASE_JSON';
var SW_ORCH_INTAKE_KEY = 'SW_ORCH_INTAKE_JSON';
var SW_ORCH_STATE_KEY = 'SW_ORCH_STATE_JSON';
var SW_ORCH_LEASE_MS = 8 * 60 * 1000;
var SW_ORCH_STALE_HEARTBEAT_MS = 8 * 60 * 1000;
var SW_ORCH_INTAKE_DRAIN_MS = 2 * 60 * 1000;
var SW_ORCH_INTAKE_STALE_MS = 45 * 60 * 1000;
var SW_ORCH_HOURLY_MS = 60 * 60 * 1000;
var SW_ORCH_HEAVY_MAINTENANCE_PROP = 'SW_ORCH_ENABLE_HEAVY_MAINTENANCE';

function sw_backgroundOrchestrator(e) {
  return swOrchTimed_('sw_backgroundOrchestrator', function () {
    var runId = swOrchRunId_();
    var startedAt = new Date();
    var lease = swOrchAcquireLease_(runId, e);
    if (!lease.ok) {
      var skippedSummary = {
        ok: true,
        runId: runId,
        skipped: true,
        reason: lease.reason || 'LEASE_ACTIVE',
        activeLease: lease.lease || null,
        startedAt: swOrchIso_(startedAt),
        source: swOrchSource_(e),
        jobs: []
      };
      try {
        Logger.log('SW_ORCH_SUMMARY ' + JSON.stringify(swOrchSummaryForLog_(skippedSummary)));
      } catch (_) {}
      return skippedSummary;
    }

    var state = swOrchReadState_();
    var summary = {
      ok: true,
      runId: runId,
      startedAt: swOrchIso_(startedAt),
      source: swOrchSource_(e),
      jobs: []
    };

    try {
      var intake = swOrchIntakeStatus_();
      if (intake.defer) {
        summary.skipped = true;
        summary.reason = intake.reason;
        summary.intake = intake;
        return summary;
      }

      var acuityJob = swOrchRunJob_(summary, runId, 'acuityPollAndSubmit', function () {
        return typeof acuityPollAndSubmit === 'function'
          ? acuityPollAndSubmit()
          : { ok: true, skipped: true, reason: 'acuityPollAndSubmit unavailable' };
      });
      if (acuityJob.ok) {
        state.lastAcuityPollAt = swOrchIso_(new Date());
      }

      if (swOrchAcuitySubmitted_(acuityJob.result)) {
        swOrchMarkIntakeDrain_('acuity submitted Google Form responses', {
          submitted: Number(acuityJob.result.submitted || 0),
          rescheduled: Number(acuityJob.result.rescheduled || 0)
        });
        summary.deferred = true;
        summary.reason = 'ACUITY_FORM_SUBMIT_DRAIN';
        swOrchSaveState_(state);
        return summary;
      }

      if (swOrchMaybeDeferForIntake_(summary)) {
        swOrchSaveState_(state);
        return summary;
      }
      var labelJob = swOrchRunJob_(summary, runId, 'acuityLabelSync', function () {
        return typeof acuityLabelSync === 'function'
          ? acuityLabelSync()
          : { ok: true, skipped: true, reason: 'acuityLabelSync unavailable' };
      });
      if (labelJob.ok) state.lastAcuityLabelSyncAt = swOrchIso_(new Date());

      if (swOrchMaybeDeferForIntake_(summary)) {
        swOrchSaveState_(state);
        return summary;
      }
      var automationJob = swOrchRunJob_(summary, runId, 'sw_processAppointmentAutomation', function () {
        return typeof sw_processAppointmentAutomation === 'function'
          ? sw_processAppointmentAutomation({
              source: 'orchestrator',
              lockWaitMs: 1500,
              deferTaskGeneration: !swOrchHeavyMaintenanceEnabled_(e)
            })
          : { ok: true, skipped: true, reason: 'sw_processAppointmentAutomation unavailable' };
      });
      if (automationJob.ok) {
        state.lastAppointmentAutomationAt = swOrchIso_(new Date());
        if (automationJob.result && automationJob.result.generatedTasks) {
          state.lastTaskGenerationAt = swOrchIso_(new Date());
        }
      }

      if (swOrchMaybeDeferForIntake_(summary)) {
        swOrchSaveState_(state);
        return summary;
      }
      var heavyMaintenanceEnabled = swOrchHeavyMaintenanceEnabled_(e);
      if (!heavyMaintenanceEnabled) {
        swOrchRecordSkip_(summary, 'sw_generateSalesWorkflowTasks', 'heavyMaintenanceDisabledForIntakeSafety');
      } else if (!automationJob.result || !automationJob.result.generatedTasks) {
        var taskDue = swOrchIntervalDue_(state.lastTaskGenerationAt, SW_ORCH_HOURLY_MS);
        if (taskDue.due) {
          var taskJob = swOrchRunJob_(summary, runId, 'sw_generateSalesWorkflowTasks', function () {
            return typeof sw_generateSalesWorkflowTasks === 'function'
              ? sw_generateSalesWorkflowTasks()
              : { ok: true, skipped: true, reason: 'sw_generateSalesWorkflowTasks unavailable' };
          });
          if (taskJob.ok) state.lastTaskGenerationAt = swOrchIso_(new Date());
        } else {
          swOrchRecordSkip_(summary, 'sw_generateSalesWorkflowTasks', taskDue.reason);
        }
      } else {
        swOrchRecordSkip_(summary, 'sw_generateSalesWorkflowTasks', 'alreadyGeneratedByAppointmentAutomation');
      }

      if (swOrchMaybeDeferForIntake_(summary)) {
        swOrchSaveState_(state);
        return summary;
      }
      var repairDue = swOrchIntervalDue_(state.lastRepairAt, SW_ORCH_HOURLY_MS);
      if (repairDue.due) {
        var repairJob = swOrchRunJob_(summary, runId, 'repairMissingUrls_', function () {
          return typeof repairMissingUrls_ === 'function'
            ? repairMissingUrls_({ force: true, source: 'orchestrator' })
            : { ok: true, skipped: true, reason: 'repairMissingUrls_ unavailable' };
        });
        if (repairJob.ok) state.lastRepairAt = swOrchIso_(new Date());
      } else {
        swOrchRecordSkip_(summary, 'repairMissingUrls_', repairDue.reason);
      }

      if (swOrchMaybeDeferForIntake_(summary)) {
        swOrchSaveState_(state);
        return summary;
      }
      var readModelDue = swOrchReadModelsDue_();
      if (!heavyMaintenanceEnabled) {
        swOrchRecordSkip_(summary, 'sw_rebuildWorkflowReadModels', 'heavyMaintenanceDisabledForIntakeSafety');
      } else if (readModelDue.due) {
        var readModelJob = swOrchRunJob_(summary, runId, 'sw_rebuildWorkflowReadModels', function () {
          return typeof sw_rebuildWorkflowReadModels === 'function'
            ? sw_rebuildWorkflowReadModels({ reason: 'orchestrator' })
            : { ok: true, skipped: true, reason: 'sw_rebuildWorkflowReadModels unavailable' };
        });
        if (readModelJob.ok) state.lastReadModelAt = swOrchIso_(new Date());
      } else {
        swOrchRecordSkip_(summary, 'sw_rebuildWorkflowReadModels', readModelDue.reason);
      }

      state.lastRunAt = swOrchIso_(new Date());
      state.lastRunOk = summary.ok !== false;
      swOrchSaveState_(state);
      return summary;
    } finally {
      swOrchReleaseLease_(runId);
      try {
        Logger.log('SW_ORCH_SUMMARY ' + JSON.stringify(swOrchSummaryForLog_(summary)));
      } catch (_) {}
    }
  });
}

function sw_runBackgroundMaintenanceOnce() {
  return sw_backgroundOrchestrator({
    source: 'manualMaintenance',
    forceMaintenance: true
  });
}

function sw_installBackgroundOrchestratorTrigger() {
  var removed = sw_removeBackgroundWorkerTriggers_();
  var orchestrators = ScriptApp.getProjectTriggers().filter(function (trigger) {
    return trigger.getHandlerFunction() === SW_ORCH_HANDLER;
  });
  for (var i = 1; i < orchestrators.length; i++) {
    try {
      ScriptApp.deleteTrigger(orchestrators[i]);
      removed.push(SW_ORCH_HANDLER + ':duplicate');
    } catch (err) {
      removed.push(SW_ORCH_HANDLER + ':duplicate:ERROR:' + swOrchErrorMessage_(err));
    }
  }
  var exists = orchestrators.length > 0;
  if (!exists) {
    ScriptApp.newTrigger(SW_ORCH_HANDLER).timeBased().everyMinutes(5).create();
  }
  try {
    PropertiesService.getScriptProperties().setProperty('REPAIR_MISSING_URLS_TRIGGER_MODE', 'orchestrated');
  } catch (_) {}
  return {
    ok: true,
    handler: SW_ORCH_HANDLER,
    cadence: 'every 5 minutes',
    removed: removed
  };
}

function sw_removeBackgroundWorkerTriggers_() {
  var handlers = {
    acuityPollAndSubmit: true,
    acuityLabelSync: true,
    sw_processAppointmentAutomation: true,
    sw_generateSalesWorkflowTasks: true,
    sw_rebuildWorkflowReadModels: true,
    repairMissingUrls_: true,
    processUploadQueue: true,
    processSummariesWorker: true,
    processIntakeQueue: true,
    ensureBootstrapForRecentRows_: true,
    cs_onOpenTrigger_: true
  };
  var removed = [];
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    var fn = trigger.getHandlerFunction();
    if (!handlers[fn]) return;
    try {
      ScriptApp.deleteTrigger(trigger);
      removed.push(fn);
    } catch (err) {
      removed.push(fn + ':ERROR:' + swOrchErrorMessage_(err));
    }
  });
  return removed;
}

function sw_removeBackgroundOrchestratorTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() !== SW_ORCH_HANDLER) return;
    ScriptApp.deleteTrigger(trigger);
    removed++;
  });
  return { ok: true, handler: SW_ORCH_HANDLER, removed: removed };
}

function sw_getBackgroundOrchestratorStatus() {
  var lease = swOrchReadJsonProperty_(SW_ORCH_LEASE_KEY, null);
  var status = {
    ok: true,
    handler: SW_ORCH_HANDLER,
    lease: lease,
    leaseStale: swOrchLeaseStale_(lease),
    heavyMaintenanceEnabled: swOrchHeavyMaintenanceEnabled_({}),
    intake: swOrchReadJsonProperty_(SW_ORCH_INTAKE_KEY, {}),
    state: swOrchReadState_(),
    triggers: swOrchListRelevantTriggers_()
  };
  try {
    Logger.log('SW_ORCH_STATUS ' + JSON.stringify(status));
  } catch (_) {}
  return status;
}

function sw_clearBackgroundOrchestratorLease() {
  return swOrchWithScriptLock_(5000, function () {
    var props = PropertiesService.getScriptProperties();
    var prior = swOrchReadJsonProperty_(SW_ORCH_LEASE_KEY, null);
    props.deleteProperty(SW_ORCH_LEASE_KEY);
    var result = {
      ok: true,
      cleared: !!prior,
      priorLease: prior
    };
    try {
      Logger.log('SW_ORCH_LEASE_CLEARED ' + JSON.stringify(result));
    } catch (_) {}
    return result;
  });
}

function swOrchRedirectLegacyTrigger_(handler, e) {
  if (!e || !e.triggerUid) return null;
  if (handler === SW_ORCH_HANDLER) return null;
  if (typeof sw_backgroundOrchestrator !== 'function') return null;
  try {
    return sw_backgroundOrchestrator({
      legacyHandler: handler,
      legacyTriggerUid: e.triggerUid
    });
  } catch (err) {
    try {
      Logger.log('SW_ORCH_LEGACY_REDIRECT_ERROR ' + JSON.stringify({
        handler: handler,
        error: swOrchErrorMessage_(err)
      }));
    } catch (_) {}
    return {
      ok: true,
      skipped: true,
      reason: 'LEGACY_REDIRECT_ERROR',
      handler: handler,
      error: swOrchErrorMessage_(err)
    };
  }
}

function swOrchMarkIntakeStart_(reason) {
  return swOrchUpdateIntake_(function (state) {
    var now = new Date();
    state.active = Math.max(0, Number(state.active || 0)) + 1;
    state.lastReason = reason || 'intake start';
    state.lastStartedAt = swOrchIso_(now);
    state.updatedAt = swOrchIso_(now);
    state.drainUntil = swOrchIso_(new Date(now.getTime() + SW_ORCH_INTAKE_DRAIN_MS));
    return state;
  });
}

function swOrchMarkIntakeFinish_(reason) {
  return swOrchUpdateIntake_(function (state) {
    var now = new Date();
    state.active = Math.max(0, Number(state.active || 0) - 1);
    state.lastReason = reason || 'intake finish';
    state.lastFinishedAt = swOrchIso_(now);
    state.updatedAt = swOrchIso_(now);
    state.drainUntil = swOrchIso_(new Date(now.getTime() + SW_ORCH_INTAKE_DRAIN_MS));
    return state;
  });
}

function swOrchMarkIntakeDrain_(reason, details) {
  return swOrchUpdateIntake_(function (state) {
    var now = new Date();
    state.active = Math.max(0, Number(state.active || 0));
    state.lastReason = reason || 'intake drain';
    state.lastDetails = details || {};
    state.updatedAt = swOrchIso_(now);
    state.drainUntil = swOrchIso_(new Date(now.getTime() + SW_ORCH_INTAKE_DRAIN_MS));
    return state;
  });
}

function swOrchUpdateIntake_(mutator) {
  return swOrchWithScriptLock_(5000, function () {
    var state = swOrchReadJsonProperty_(SW_ORCH_INTAKE_KEY, {});
    state = mutator(state || {}) || state || {};
    PropertiesService.getScriptProperties().setProperty(SW_ORCH_INTAKE_KEY, JSON.stringify(state));
    return state;
  });
}

function swOrchIntakeStatus_() {
  var state = swOrchReadJsonProperty_(SW_ORCH_INTAKE_KEY, {});
  var nowMs = new Date().getTime();
  var active = Math.max(0, Number(state.active || 0));
  var updatedMs = swOrchDateMs_(state.updatedAt || state.lastStartedAt || '');
  var drainMs = swOrchDateMs_(state.drainUntil || '');
  if (active > 0 && (!updatedMs || nowMs - updatedMs < SW_ORCH_INTAKE_STALE_MS)) {
    return { defer: true, reason: 'INTAKE_ACTIVE', active: active, drainUntil: state.drainUntil || '' };
  }
  if (drainMs && drainMs > nowMs) {
    return { defer: true, reason: 'INTAKE_DRAIN', active: active, drainUntil: state.drainUntil || '' };
  }
  return { defer: false, active: active, drainUntil: state.drainUntil || '' };
}

function swOrchMaybeDeferForIntake_(summary) {
  var intake = swOrchIntakeStatus_();
  if (!intake.defer) return false;
  summary.deferred = true;
  summary.reason = intake.reason;
  summary.intake = intake;
  return true;
}

function swOrchAcquireLease_(runId, sourceEvent) {
  return swOrchWithScriptLock_(5000, function () {
    var now = new Date();
    var nowMs = now.getTime();
    var existing = swOrchReadJsonProperty_(SW_ORCH_LEASE_KEY, null);
    if (existing && existing.runId && !swOrchLeaseStale_(existing) && swOrchDateMs_(existing.expiresAt) > nowMs) {
      return { ok: false, reason: 'LEASE_ACTIVE', lease: existing };
    }
    var lease = {
      runId: runId,
      startedAt: swOrchIso_(now),
      heartbeatAt: swOrchIso_(now),
      expiresAt: swOrchIso_(new Date(nowMs + SW_ORCH_LEASE_MS)),
      source: swOrchSource_(sourceEvent)
    };
    PropertiesService.getScriptProperties().setProperty(SW_ORCH_LEASE_KEY, JSON.stringify(lease));
    return { ok: true, lease: lease };
  });
}

function swOrchLeaseStale_(lease) {
  if (!lease || !lease.runId) return false;
  var nowMs = new Date().getTime();
  var expiresMs = swOrchDateMs_(lease.expiresAt || '');
  if (expiresMs && expiresMs <= nowMs) return true;
  var heartbeatMs = swOrchDateMs_(lease.heartbeatAt || lease.startedAt || '');
  if (!heartbeatMs) return true;
  return nowMs - heartbeatMs > SW_ORCH_STALE_HEARTBEAT_MS;
}

function swOrchHeartbeat_(runId, currentJob) {
  try {
    swOrchWithScriptLock_(1000, function () {
      var lease = swOrchReadJsonProperty_(SW_ORCH_LEASE_KEY, null);
      if (!lease || lease.runId !== runId) return null;
      var now = new Date();
      lease.heartbeatAt = swOrchIso_(now);
      lease.expiresAt = swOrchIso_(new Date(now.getTime() + SW_ORCH_LEASE_MS));
      lease.currentJob = currentJob || '';
      PropertiesService.getScriptProperties().setProperty(SW_ORCH_LEASE_KEY, JSON.stringify(lease));
      return lease;
    });
  } catch (_) {}
}

function swOrchReleaseLease_(runId) {
  try {
    swOrchWithScriptLock_(5000, function () {
      var lease = swOrchReadJsonProperty_(SW_ORCH_LEASE_KEY, null);
      if (lease && lease.runId === runId) {
        PropertiesService.getScriptProperties().deleteProperty(SW_ORCH_LEASE_KEY);
      }
      return true;
    });
  } catch (_) {}
}

function swOrchRunJob_(summary, runId, name, fn) {
  swOrchHeartbeat_(runId, name);
  var started = new Date();
  var job = {
    name: name,
    startedAt: swOrchIso_(started)
  };
  try {
    job.result = fn();
    job.ok = !(job.result && job.result.ok === false);
  } catch (err) {
    job.ok = false;
    job.error = swOrchErrorMessage_(err);
    summary.ok = false;
  }
  job.ms = new Date().getTime() - started.getTime();
  summary.jobs.push(job);
  try {
    Logger.log('SW_ORCH_JOB ' + JSON.stringify(swOrchJobForLog_(job)));
  } catch (_) {}
  swOrchHeartbeat_(runId, '');
  return job;
}

function swOrchRecordSkip_(summary, name, reason) {
  summary.jobs.push({
    name: name,
    ok: true,
    skipped: true,
    reason: reason || 'notDue',
    ms: 0
  });
}

function swOrchAcuitySubmitted_(result) {
  if (!result) return false;
  return Number(result.submitted || 0) > 0 || Number(result.rescheduled || 0) > 0 || Number(result.formSubmitted || 0) > 0;
}

function swOrchHeavyMaintenanceEnabled_(e) {
  if (e && e.forceMaintenance) return true;
  try {
    return /^true$/i.test(String(PropertiesService.getScriptProperties().getProperty(SW_ORCH_HEAVY_MAINTENANCE_PROP) || ''));
  } catch (_) {
    return false;
  }
}

function swOrchSource_(e) {
  if (e && e.legacyHandler) return 'legacy:' + e.legacyHandler;
  if (e && e.source) return String(e.source);
  if (e && e.forceMaintenance) return 'manualMaintenance';
  return 'timeBased';
}

function swOrchIntervalDue_(lastIso, intervalMs) {
  var lastMs = swOrchDateMs_(lastIso);
  if (!lastMs) return { due: true, reason: 'neverRun' };
  var elapsed = new Date().getTime() - lastMs;
  if (elapsed >= intervalMs) return { due: true, reason: 'intervalElapsed' };
  return { due: false, reason: 'notDue' };
}

function swOrchReadModelsDue_() {
  if (typeof sw_getWorkflowReadModelStatus !== 'function') {
    return { due: false, reason: 'statusUnavailable' };
  }
  try {
    var status = sw_getWorkflowReadModelStatus();
    if (status && status.allFresh) return { due: false, reason: 'allFresh' };
    return { due: true, reason: 'staleOrMissing', status: status };
  } catch (err) {
    return { due: true, reason: 'statusError', error: swOrchErrorMessage_(err) };
  }
}

function swOrchReadState_() {
  return swOrchReadJsonProperty_(SW_ORCH_STATE_KEY, {});
}

function swOrchSaveState_(state) {
  PropertiesService.getScriptProperties().setProperty(SW_ORCH_STATE_KEY, JSON.stringify(state || {}));
}

function swOrchReadJsonProperty_(key, fallback) {
  var raw = '';
  try {
    raw = PropertiesService.getScriptProperties().getProperty(key) || '';
    return raw ? JSON.parse(raw) : fallback;
  } catch (_) {
    return fallback;
  }
}

function swOrchWithScriptLock_(waitMs, fn) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    locked = lock.tryLock(waitMs || 3000);
    if (!locked) return { ok: false, reason: 'SCRIPT_LOCK_BUSY' };
    return fn();
  } finally {
    if (locked) {
      try { lock.releaseLock(); } catch (_) {}
    }
  }
}

function swOrchListRelevantTriggers_() {
  var handlers = {};
  handlers[SW_ORCH_HANDLER] = true;
  [
    'acuityPollAndSubmit',
    'acuityLabelSync',
    'sw_processAppointmentAutomation',
    'sw_generateSalesWorkflowTasks',
    'sw_rebuildWorkflowReadModels',
    'repairMissingUrls_',
    'onFormSubmit',
    'onOpen',
    'cs_onOpenTrigger_'
  ].forEach(function (name) { handlers[name] = true; });
  var out = [];
  try {
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
      var fn = trigger.getHandlerFunction();
      if (!handlers[fn]) return;
      out.push({
        handler: fn,
        eventType: swOrchSafeTriggerValue_(trigger, 'getEventType'),
        source: swOrchSafeTriggerValue_(trigger, 'getTriggerSource'),
        sourceId: swOrchSafeTriggerValue_(trigger, 'getTriggerSourceId'),
        uniqueId: swOrchSafeTriggerValue_(trigger, 'getUniqueId')
      });
    });
  } catch (err) {
    out.push({ error: swOrchErrorMessage_(err) });
  }
  return out;
}

function swOrchSafeTriggerValue_(trigger, methodName) {
  try {
    if (!trigger || typeof trigger[methodName] !== 'function') return '';
    var value = trigger[methodName]();
    return value == null ? '' : String(value);
  } catch (_) {
    return '';
  }
}

function swOrchRunId_() {
  return 'orch_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000000);
}

function swOrchIso_(date) {
  try {
    if (typeof swIso_ === 'function') return swIso_(date);
  } catch (_) {}
  return (date instanceof Date ? date : new Date(date)).toISOString();
}

function swOrchDateMs_(value) {
  if (!value) return 0;
  var ms = new Date(value).getTime();
  return isNaN(ms) ? 0 : ms;
}

function swOrchErrorMessage_(err) {
  if (!err) return '';
  return err && err.message ? String(err.message) : String(err);
}

function swOrchTimed_(operation, fn) {
  if (typeof swTimed_ === 'function') return swTimed_(operation, fn);
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

function swOrchResultSummary_(result) {
  if (!result || typeof result !== 'object') return result == null ? null : String(result);
  var keys = [
    'ok', 'skipped', 'reason', 'message',
    'submitted', 'rescheduled', 'edited', 'canceled', 'updated',
    'checkedExisting', 'existingCandidates', 'checkedCanceled', 'canceledCandidates',
    'deferredExisting', 'deferredCanceled',
    'taskGenerationDeferred',
    'processed', 'errors', 'generatedTasks', 'created', 'repaired',
    'allFresh', 'totalMs'
  ];
  var out = {};
  keys.forEach(function (key) {
    if (result[key] !== undefined) out[key] = result[key];
  });
  return out;
}

function swOrchJobForLog_(job) {
  return {
    name: job.name,
    ok: job.ok,
    skipped: !!job.skipped,
    reason: job.reason || '',
    error: job.error || '',
    ms: job.ms || 0,
    result: swOrchResultSummary_(job.result)
  };
}

function swOrchSummaryForLog_(summary) {
  return {
    ok: summary.ok,
    runId: summary.runId,
    startedAt: summary.startedAt,
    source: summary.source || '',
    skipped: !!summary.skipped,
    deferred: !!summary.deferred,
    reason: summary.reason || '',
    jobs: (summary.jobs || []).map(swOrchJobForLog_)
  };
}
