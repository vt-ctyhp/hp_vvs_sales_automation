/**
 * Sales workflow read models.
 *
 * Phase 1 is shadow-only: these generated tabs are built for benchmarking and
 * comparison, but the web app still serves from the existing source sheets.
 */

var SW_READ_MODEL_VERSION = 'phase1-v1';
var SW_READ_MODEL_DEFAULT_TTL_SECONDS = 10 * 60;
var SW_READ_MODEL_REFRESH_HANDLER = 'sw_rebuildWorkflowReadModels';
var SW_READ_MODEL_INVALIDATED_THIS_EXECUTION_ = {};
var SW_TASK_DASHBOARD_CACHE_SECONDS = 10 * 60;
var SW_TASK_DASHBOARD_MEMORY_CACHE_ = {};

function sw_rebuildWorkflowReadModels(options) {
  options = options || {};
  var ss = swSpreadsheet_();
  var lock = LockService.getDocumentLock() || LockService.getScriptLock();
  lock.waitLock(28000);
  try {
    return swRebuildWorkflowReadModelsUnlocked_(ss, options);
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function sw_getWorkflowReadModelStatus() {
  return swWorkflowReadModelStatus_(swSpreadsheet_());
}

function sw_invalidateWorkflowReadModels(reason) {
  var ss = swSpreadsheet_();
  var result = swMarkWorkflowReadModelsStale_(ss, reason || 'Manual invalidation');
  if (result.invalidated) return result;
  return {
    ok: true,
    invalidated: false,
    reason: swTrim_(reason || ''),
    message: 'No read-model metadata exists yet.'
  };
}

function swMarkWorkflowReadModelsStale_(ss, reason, modelName) {
  var sh = ss.getSheetByName(SW_SHEETS.READ_MODEL_META);
  var now = swIso_(new Date());
  modelName = swTrim_(modelName || '');
  if (!modelName || modelName === 'tasks') {
    try { swInvalidateTaskDashboardProjectionCache_(ss); } catch (_) {}
  }
  var key = ss.getId() + ':' + (modelName || 'all');
  if (SW_READ_MODEL_INVALIDATED_THIS_EXECUTION_[key]) {
    return {
      ok: true,
      invalidated: true,
      skippedDuplicate: true,
      reason: swTrim_(reason || ''),
      models: 0,
      invalidatedAt: now
    };
  }
  if (!sh || sh.getLastRow() < 2) {
    return {
      ok: true,
      invalidated: false,
      reason: swTrim_(reason || ''),
      message: 'No read-model metadata exists yet.'
    };
  }

  var headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), SW_READ_MODEL_META_HEADERS.length)).getDisplayValues()[0].map(swTrim_);
  var H = swHeaderMapFromArray_(headers);
  var statusCol = swPickIndex_(H, ['Status']) + 1;
  var invalidatedCol = swPickIndex_(H, ['Invalidated At']) + 1;
  var notesCol = swPickIndex_(H, ['Notes']) + 1;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), SW_READ_MODEL_META_HEADERS.length)).getDisplayValues();
  var changed = 0;
  rows.forEach(function (row, idx) {
    var rowModel = swTrim_(row[swPickIndex_(H, ['Model'])] || '');
    if (modelName && rowModel !== modelName) return;
    var rowNumber = idx + 2;
    if (statusCol > 0) sh.getRange(rowNumber, statusCol).setValue('STALE');
    if (invalidatedCol > 0) sh.getRange(rowNumber, invalidatedCol).setValue(now);
    if (notesCol > 0) sh.getRange(rowNumber, notesCol).setValue(swTrim_(reason || 'Manual invalidation'));
    changed++;
  });
  SW_READ_MODEL_INVALIDATED_THIS_EXECUTION_[key] = true;

  return {
    ok: true,
    invalidated: changed > 0,
    models: changed,
    invalidatedAt: now,
    reason: swTrim_(reason || '')
  };
}

function sw_installWorkflowReadModelRefreshTrigger() {
  sw_removeWorkflowReadModelRefreshTriggers();
  ScriptApp.newTrigger(SW_READ_MODEL_REFRESH_HANDLER).timeBased().everyMinutes(5).create();
  return {
    ok: true,
    handler: SW_READ_MODEL_REFRESH_HANDLER,
    cadence: 'every 5 minutes'
  };
}

function sw_removeWorkflowReadModelRefreshTriggers() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() !== SW_READ_MODEL_REFRESH_HANDLER) return;
    ScriptApp.deleteTrigger(trigger);
    removed++;
  });
  return { ok: true, removed: removed, handler: SW_READ_MODEL_REFRESH_HANDLER };
}

function sw_measureWorkflowReadModelBuildSpeed(options) {
  var started = new Date().getTime();
  var result = sw_rebuildWorkflowReadModels(options || {});
  result.totalMs = new Date().getTime() - started;
  Logger.log('SW_READ_MODEL_BENCHMARK ' + JSON.stringify(result));
  return result;
}

function swRebuildWorkflowReadModelsUnlocked_(ss, options) {
  options = options || {};
  var started = new Date().getTime();
  var builtAt = new Date();
  var builtAtIso = swIso_(builtAt);
  var ttlSeconds = swReadModelTtlSeconds_(ss, options);
  var expiresAtIso = swIso_(new Date(builtAt.getTime() + ttlSeconds * 1000));
  var meta = [];
  var out = {
    ok: true,
    version: SW_READ_MODEL_VERSION,
    builtAt: builtAtIso,
    expiresAt: expiresAtIso,
    ttlSeconds: ttlSeconds,
    models: {}
  };

  var taskResult = swBuildTaskReadModel_(ss, builtAt);
  meta.push(swReadModelMetaRow_('tasks', SW_SHEETS.TASKS, taskResult, builtAtIso, expiresAtIso));
  out.models.tasks = taskResult;
  if (!taskResult.ok) out.ok = false;

  var customerResult = swBuildCustomerReadModel_(ss, builtAt);
  meta.push(swReadModelMetaRow_('customers', SW_SHEETS.MASTER, customerResult, builtAtIso, expiresAtIso));
  out.models.customers = customerResult;
  if (!customerResult.ok) out.ok = false;

  var metaResult = swWriteReadModelSheet_(ss, SW_SHEETS.READ_MODEL_META, SW_READ_MODEL_META_HEADERS, meta);
  out.models.meta = metaResult;
  if (!metaResult.ok) out.ok = false;

  out.totalMs = new Date().getTime() - started;
  Logger.log('SW_READ_MODEL_REBUILD ' + JSON.stringify(swWorkflowReadModelLogSummary_(out)));
  return out;
}

function swBuildTaskReadModel_(ss, builtAt) {
  var started = new Date().getTime();
  try {
    var state = swReadTaskListState_(ss, false);
    var nowMs = builtAt.getTime();
    var rows = (state.tasks || []).map(function (task) {
      return swTaskReadModelRow_(task, nowMs);
    });
    var write = swWriteReadModelSheet_(ss, SW_SHEETS.READ_MODEL_TASKS, SW_TASK_READ_MODEL_HEADERS, rows);
    write.sourceRows = state.tasks ? state.tasks.length : 0;
    write.outputRows = rows.length;
    write.buildMs = new Date().getTime() - started;
    try { swCacheTaskListState_(ss, state); } catch (_) {}
    var projections = swBuildTaskDashboardProjections_(ss, state, builtAt);
    write.projectionUsers = projections.users || 0;
    write.projectionKeys = projections.keys || 0;
    write.projectionMs = projections.buildMs || 0;
    write.projectionError = projections.error || '';
    return write;
  } catch (err) {
    return swReadModelErrorResult_(err, started);
  }
}

function swBuildCustomerReadModel_(ss, builtAt) {
  var started = new Date().getTime();
  try {
    var appointments = swReadAppointments_(ss);
    var groups = swAdminDashboardRowsByRoot_(appointments);
    var master = ss.getSheetByName(SW_SHEETS.MASTER);
    var masterGid = master ? master.getSheetId() : '';
    var rows = [];
    Object.keys(groups).sort().forEach(function (root) {
      var rootRows = groups[root] || [];
      var activeRows = rootRows.filter(function (rec) { return swIsAppointmentActive_(rec); });
      var rec = swAdminDashboardLatestRow_(activeRows.length ? activeRows : rootRows);
      if (!rec) return;
      var stage = swAdminDashboardPipelineStage_(rec, rootRows);
      rows.push(swCustomerReadModelRow_(ss, masterGid, root, rec, rootRows, activeRows, stage));
    });
    var write = swWriteReadModelSheet_(ss, SW_SHEETS.READ_MODEL_CUSTOMERS, SW_CUSTOMER_READ_MODEL_HEADERS, rows);
    write.sourceRows = appointments.length;
    write.outputRows = rows.length;
    write.buildMs = new Date().getTime() - started;
    return write;
  } catch (err) {
    return swReadModelErrorResult_(err, started);
  }
}

function swTaskReadModelRow_(task, nowMs) {
  task = task || {};
  var values = [
    task.taskId || '',
    task.root || '',
    task.appt || '',
    task.customerName || '',
    task.brand || '',
    task.visitDate || '',
    swFormatAppointmentTime_(task.visitTime || ''),
    task.visitType || '',
    task.lifecycleStage || '',
    task.taskType || '',
    task.taskTitle || '',
    task.ownerRole || '',
    task.intendedOwner || '',
    swNormEmail_(task.intendedOwnerEmail || ''),
    task.currentOwner || '',
    swNormEmail_(task.currentOwnerEmail || ''),
    task.coverageReason || '',
    task.dueAt || '',
    task.status || '',
    task.primaryAction || '',
    task.snoozeUntil || '',
    task.snoozeReason || '',
    task.rowNumber || '',
    swTaskPendingLike_(task, nowMs) ? 'Y' : 'N',
    swTaskDueForQueue_(task, nowMs) ? 'Y' : 'N'
  ];
  values.push(swReadModelSearchText_(values));
  return values;
}

function swCustomerReadModelRow_(ss, masterGid, root, rec, rootRows, activeRows, stage) {
  var visit = swAdminDashboardVisitSummary_(rootRows);
  var sourceRows = (rootRows || []).map(function (row) { return row.row || ''; }).filter(Boolean);
  var masterUrl = masterGid && rec.row
    ? 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + masterGid + '&range=A' + rec.row
    : '';
  var values = [
    root || '',
    rec.appt || '',
    rec.row || '',
    rec.name || '',
    rec.email || '',
    rec.phone || '',
    rec.brand || '',
    rec.assignedRep || '',
    rec.assignedRepEmail || '',
    rec.assistedRep || '',
    rec.assistedRepEmail || '',
    rec.visitDate || '',
    rec.visitTime || '',
    rec.visitType || '',
    visit.next || '',
    visit.last || '',
    (rootRows || []).length,
    (activeRows || []).length,
    (activeRows || []).length ? 'Y' : 'N',
    stage.key || '',
    stage.label || '',
    rec.salesStage || '',
    rec.convStatus || '',
    rec.customOrder || '',
    rec.inProduction || '',
    rec.centerStoneStatus || '',
    rec.so || '',
    rec.orderTotal || '',
    rec.paidToDate || '',
    rec.remainingBalance || '',
    rec.lastPaymentDate || '',
    rec.quotationUrl || '',
    rec.clientFolder || '',
    rec.reportUrl || '',
    rec.tracker3dUrl || '',
    rec.deadline3d || '',
    rec.productionDeadline || '',
    rec.waxStatus || '',
    rec.waxDeadlineAdmin || '',
    rec.dvStonesSummary || '',
    rec.nextSteps || '',
    rec.updatedAt || '',
    swStringify_(sourceRows),
    ''
  ];
  values[values.length - 1] = swReadModelSearchText_(values.concat([masterUrl]));
  return values;
}

function swWorkflowReadModelStatus_(ss) {
  var nowMs = new Date().getTime();
  var metaRows = swReadModelMetaRows_(ss);
  var byModel = {};
  metaRows.forEach(function (row) {
    byModel[row['Model']] = row;
  });
  var models = swReadModelDefinitions_().map(function (def) {
    var sh = ss.getSheetByName(def.sheet);
    var meta = byModel[def.model] || {};
    var builtAt = swTrim_(meta['Built At']);
    var expiresAt = swTrim_(meta['Expires At']);
    var expiresMs = swReadModelDateMs_(expiresAt);
    var stale = !builtAt || swTrim_(meta['Status']) !== 'OK' || (expiresMs && expiresMs < nowMs);
    return {
      model: def.model,
      sheet: def.sheet,
      exists: !!sh,
      rows: sh ? Math.max(0, sh.getLastRow() - 1) : 0,
      columns: sh ? sh.getLastColumn() : 0,
      version: meta['Version'] || '',
      builtAt: builtAt,
      expiresAt: expiresAt,
      ageSeconds: builtAt ? Math.max(0, Math.round((nowMs - swReadModelDateMs_(builtAt)) / 1000)) : 0,
      fresh: !!(sh && !stale),
      stale: stale,
      status: meta['Status'] || (sh ? 'UNKNOWN' : 'MISSING'),
      error: meta['Error'] || ''
    };
  });
  var triggers = 0;
  try {
    triggers = ScriptApp.getProjectTriggers().filter(function (trigger) {
      return trigger.getHandlerFunction() === SW_READ_MODEL_REFRESH_HANDLER;
    }).length;
  } catch (_) {}
  return {
    ok: true,
    generatedAt: swIso_(new Date()),
    version: SW_READ_MODEL_VERSION,
    refreshHandler: SW_READ_MODEL_REFRESH_HANDLER,
    refreshTriggers: triggers,
    models: models,
    allFresh: models.every(function (model) { return model.fresh; })
  };
}

function swReadModelDefinitions_() {
  return [
    { model: 'tasks', sheet: SW_SHEETS.READ_MODEL_TASKS },
    { model: 'customers', sheet: SW_SHEETS.READ_MODEL_CUSTOMERS }
  ];
}

function swReadModelTtlSeconds_(ss, options) {
  var ttl = Number(options && options.ttlSeconds);
  if (!isFinite(ttl) || ttl <= 0) {
    try {
      ttl = Number(swConfigValue_(swReadConfig_(ss, true), 'SYSTEM', 'READ_MODEL_TTL_SECONDS', String(SW_READ_MODEL_DEFAULT_TTL_SECONDS)));
    } catch (_) {
      ttl = SW_READ_MODEL_DEFAULT_TTL_SECONDS;
    }
  }
  if (!isFinite(ttl) || ttl <= 0) ttl = SW_READ_MODEL_DEFAULT_TTL_SECONDS;
  return Math.max(SW_READ_MODEL_DEFAULT_TTL_SECONDS, Math.min(Math.floor(ttl), 30 * 60));
}

function swWriteReadModelSheet_(ss, sheetName, headers, rows) {
  var started = new Date().getTime();
  try {
    var sh = swEnsureSheet_(ss, sheetName, headers);
    var values = [headers].concat(rows || []);
    swSizeReadModelSheet_(sh, values.length, headers.length);
    sh.clearContents();
    sh.getRange(1, 1, values.length, headers.length).setValues(values);
    sh.setFrozenRows(1);
    try { sh.hideSheet(); } catch (_) {}
    return {
      ok: true,
      sheet: sheetName,
      outputRows: rows ? rows.length : 0,
      buildMs: new Date().getTime() - started
    };
  } catch (err) {
    return swReadModelErrorResult_(err, started, sheetName);
  }
}

function swSizeReadModelSheet_(sh, targetRows, targetCols) {
  targetRows = Math.max(2, Number(targetRows) || 2);
  targetCols = Math.max(1, Number(targetCols) || 1);
  var maxRows = sh.getMaxRows();
  var maxCols = sh.getMaxColumns();
  if (maxRows < targetRows) sh.insertRowsAfter(maxRows, targetRows - maxRows);
  if (maxCols < targetCols) sh.insertColumnsAfter(maxCols, targetCols - maxCols);
  maxRows = sh.getMaxRows();
  maxCols = sh.getMaxColumns();
  if (maxRows > targetRows + 50) sh.deleteRows(targetRows + 1, maxRows - targetRows);
  if (maxCols > targetCols + 10) sh.deleteColumns(targetCols + 1, maxCols - targetCols);
}

function swReadModelMetaRow_(model, sourceSheet, result, builtAtIso, expiresAtIso) {
  result = result || {};
  return [
    model,
    SW_READ_MODEL_VERSION,
    builtAtIso,
    expiresAtIso,
    sourceSheet || '',
    result.sourceRows || 0,
    result.outputRows || 0,
    result.buildMs || 0,
    result.ok ? 'OK' : 'ERROR',
    result.error || '',
    '',
    result.notes || ''
  ];
}

function swReadModelMetaRows_(ss) {
  var sh = ss.getSheetByName(SW_SHEETS.READ_MODEL_META);
  if (!sh || sh.getLastRow() < 2) return [];
  return swReadSheetObjectsExpectedHeaders_(sh, SW_READ_MODEL_META_HEADERS);
}

function swReadModelErrorResult_(err, started, sheetName) {
  return {
    ok: false,
    sheet: sheetName || '',
    sourceRows: 0,
    outputRows: 0,
    buildMs: new Date().getTime() - started,
    error: err && err.message ? err.message : String(err)
  };
}

function swWorkflowReadModelLogSummary_(result) {
  result = result || {};
  var models = {};
  Object.keys(result.models || {}).forEach(function (key) {
    var model = result.models[key] || {};
    models[key] = {
      ok: model.ok !== false,
      sourceRows: model.sourceRows || 0,
      outputRows: model.outputRows || 0,
      buildMs: model.buildMs || 0,
      projectionUsers: model.projectionUsers || 0,
      projectionKeys: model.projectionKeys || 0,
      projectionMs: model.projectionMs || 0,
      error: model.error || ''
    };
  });
  return {
    ok: result.ok !== false,
    version: result.version || '',
    builtAt: result.builtAt || '',
    expiresAt: result.expiresAt || '',
    ttlSeconds: result.ttlSeconds || 0,
    totalMs: result.totalMs || 0,
    models: models
  };
}

function swBenchmarkSalesWorkflowReadModelSummary_(status) {
  status = status || {};
  var out = {
    allFresh: !!status.allFresh,
    refreshTriggers: status.refreshTriggers || 0,
    models: {}
  };
  (status.models || []).forEach(function (model) {
    out.models[model.model] = {
      exists: !!model.exists,
      fresh: !!model.fresh,
      rows: model.rows || 0,
      ageSeconds: model.ageSeconds || 0,
      status: model.status || '',
      error: model.error || ''
    };
  });
  return out;
}

function swReadModelSearchText_(values) {
  return swNorm_((values || []).join(' ')).slice(0, 4000);
}

function swReadModelDateMs_(value) {
  var s = swTrim_(value);
  if (!s) return 0;
  var d = new Date(s);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function swReadTaskListStateForDashboard_(ss, config) {
  var readModel = swTryReadTaskListStateFromReadModel_(ss, config);
  if (readModel && readModel.state) return readModel;

  var state = swReadTaskListState_(ss, true);
  return {
    source: 'taskQueue',
    fallbackReason: readModel ? readModel.fallbackReason || '' : '',
    ageSeconds: readModel ? readModel.ageSeconds || 0 : 0,
    state: state
  };
}

function swTryReadTaskListStateFromReadModel_(ss, config) {
  if (!swTaskReadModelServingEnabled_(config)) {
    return { source: 'taskQueue', fallbackReason: 'disabled', state: null };
  }
  try {
    var status = swTaskReadModelStatus_(ss);
    if (!status.fresh) {
      return {
        source: 'taskQueue',
        fallbackReason: status.reason || 'notFresh',
        ageSeconds: status.ageSeconds || 0,
        state: null
      };
    }

    var cachedState = swReadCachedTaskListState_(ss);
    if (cachedState && cachedState.tasks) {
      return {
        source: 'taskReadModelCache',
        fallbackReason: '',
        ageSeconds: status.ageSeconds || 0,
        state: cachedState
      };
    }

    var sh = ss.getSheetByName(SW_SHEETS.READ_MODEL_TASKS);
    var rows = swReadSheetObjectsExpectedHeaders_(sh, SW_TASK_READ_MODEL_HEADERS);
    var tasks = rows.map(swTaskFromReadModelRow_).filter(function (task) {
      return !!task.taskId;
    });
    var state = swBuildTaskStateFromTasks_(tasks, [], {});
    try { swCacheTaskListState_(ss, state); } catch (_) {}
    return {
      source: 'taskReadModelSheet',
      fallbackReason: '',
      ageSeconds: status.ageSeconds || 0,
      state: state
    };
  } catch (err) {
    try {
      Logger.log('SW_READ_MODEL_TASK_FALLBACK ' + JSON.stringify({
        reason: err && err.message ? err.message : String(err)
      }));
    } catch (_) {}
    return {
      source: 'taskQueue',
      fallbackReason: err && err.message ? err.message : String(err),
      state: null
    };
  }
}

function swTaskReadModelServingEnabled_(config) {
  return swNorm_(swConfigValue_(config || [], 'SYSTEM', 'READ_MODEL_SERVE_TASKS', 'Y')) !== 'n';
}

function swTaskReadModelStatus_(ss) {
  var sh = ss.getSheetByName(SW_SHEETS.READ_MODEL_TASKS);
  if (!sh) return { fresh: false, reason: 'missingSheet' };
  var meta = null;
  var rows = swReadModelMetaRows_(ss);
  for (var i = 0; i < rows.length; i++) {
    if (swTrim_(rows[i]['Model']) === 'tasks') {
      meta = rows[i];
      break;
    }
  }
  if (!meta) return { fresh: false, reason: 'missingMeta' };
  if (swTrim_(meta['Version']) !== SW_READ_MODEL_VERSION) return { fresh: false, reason: 'versionMismatch' };
  if (swTrim_(meta['Status']) !== 'OK') return { fresh: false, reason: 'status:' + swTrim_(meta['Status']) };
  if (swTrim_(meta['Invalidated At'])) return { fresh: false, reason: 'invalidated' };
  var builtAtMs = swReadModelDateMs_(meta['Built At']);
  var expiresAtMs = swReadModelDateMs_(meta['Expires At']);
  var nowMs = new Date().getTime();
  var ageSeconds = builtAtMs ? Math.max(0, Math.round((nowMs - builtAtMs) / 1000)) : 0;
  if (!builtAtMs || !expiresAtMs) return { fresh: false, reason: 'missingDates', ageSeconds: ageSeconds };
  if (expiresAtMs < nowMs) return { fresh: false, reason: 'expired', ageSeconds: ageSeconds };
  return {
    fresh: true,
    reason: '',
    ageSeconds: ageSeconds,
    builtAt: meta['Built At'] || '',
    expiresAt: meta['Expires At'] || '',
    rows: Math.max(0, sh.getLastRow() - 1)
  };
}

function swTaskFromReadModelRow_(row) {
  row = row || {};
  return {
    taskId: row['TaskID'] || '',
    root: row['RootApptID'] || '',
    appt: row['APPT_ID'] || '',
    customerName: row['Customer Name'] || '',
    brand: row['Brand'] || '',
    visitDate: row['Visit Date'] || '',
    visitTime: swFormatAppointmentTime_(row['Visit Time'] || ''),
    visitType: row['Visit Type'] || '',
    lifecycleStage: row['Lifecycle Stage'] || '',
    taskType: row['Task Type'] || '',
    taskTitle: row['Task Title'] || '',
    ownerRole: row['Owner Role'] || '',
    intendedOwner: row['Intended Owner'] || '',
    intendedOwnerEmail: swNormEmail_(row['Intended Owner Email'] || ''),
    currentOwner: row['Current Owner'] || '',
    currentOwnerEmail: swNormEmail_(row['Current Owner Email'] || ''),
    coverageReason: row['Coverage Reason'] || '',
    dueAt: row['Due At'] || '',
    status: row['Status'] || SW_STATUSES.PENDING,
    primaryAction: row['Primary Action'] || '',
    snoozeUntil: row['Snooze Until'] || '',
    snoozeReason: row['Snooze Reason'] || '',
    rowNumber: Number(row['Row Number'] || 0) || 0
  };
}

function swBuildTaskDashboardProjections_(ss, state, builtAt) {
  var started = new Date().getTime();
  try {
    var config = swReadConfig_(ss, true);
    var cleanupTabEnabled = typeof swDataCleanupCampaignTabEnabled_ === 'function'
      ? swDataCleanupCampaignTabEnabled_(config)
      : false;
    var users = swTaskDashboardProjectionUsers_(ss);
    swClearTaskDashboardProjectionCaches_(ss);

    var builtAtIso = swIso_(builtAt);
    var expiresAtIso = swIso_(new Date(builtAt.getTime() + SW_TASK_DASHBOARD_CACHE_SECONDS * 1000));
    var keys = [];
    users.forEach(function (user) {
      var buckets = swBuildVisibleTaskBuckets_(state, user, { cleanupCampaignTabEnabled: cleanupTabEnabled });
      var counts = {
        mine: buckets.mine.length,
        cleanup: buckets.cleanup.length,
        coverage: buckets.coverage.length,
        admin: buckets.admin.length
      };
      var base = {
        ok: true,
        version: SW_READ_MODEL_VERSION,
        builtAt: builtAtIso,
        expiresAt: expiresAtIso,
        email: user.email,
        signature: swTaskDashboardUserSignature_(user, cleanupTabEnabled),
        totalTasks: (state.tasks || []).length,
        counts: counts,
        ageSeconds: 0
      };
      keys.push(swPutTaskDashboardProjection_(ss, user, cleanupTabEnabled, 'bootstrap', swMergeObjects_(base, {
        tasks: buckets.mine
      })));
      ['mine', 'cleanup', 'coverage', 'admin'].forEach(function (viewName) {
        keys.push(swPutTaskDashboardProjection_(ss, user, cleanupTabEnabled, 'view:' + viewName, swMergeObjects_(base, {
          view: viewName,
          tasks: buckets[viewName] || []
        })));
      });
    });
    swPutTaskDashboardProjectionIndex_(ss, keys.filter(Boolean));
    swClearTaskDashboardInvalidation_(ss);
    return {
      ok: true,
      users: users.length,
      keys: keys.filter(Boolean).length,
      buildMs: new Date().getTime() - started,
      error: ''
    };
  } catch (err) {
    return {
      ok: false,
      users: 0,
      keys: 0,
      buildMs: new Date().getTime() - started,
      error: err && err.message ? err.message : String(err)
    };
  }
}

function swReadTaskDashboardBootstrapProjection_(ss, user, config) {
  return swReadTaskDashboardProjection_(ss, user, config, 'bootstrap');
}

function swReadTaskDashboardViewProjection_(ss, user, viewName, config) {
  viewName = swTrim_(viewName || 'mine');
  if (['mine', 'cleanup', 'coverage', 'admin'].indexOf(viewName) < 0) return null;
  return swReadTaskDashboardProjection_(ss, user, config, 'view:' + viewName);
}

function swReadTaskDashboardProjection_(ss, user, config, projectionType) {
  if (!swTaskReadModelServingEnabled_(config)) return null;
  var cleanupTabEnabled = typeof swDataCleanupCampaignTabEnabled_ === 'function'
    ? swDataCleanupCampaignTabEnabled_(config || [])
    : false;
  user = user || {};
  if (!user.email) return null;
  var key = swTaskDashboardProjectionKey_(ss, user, cleanupTabEnabled, projectionType);
  var payload = swTaskDashboardProjectionCacheGet_(key);
  if (!payload || payload.version !== SW_READ_MODEL_VERSION) return null;
  if (payload.signature !== swTaskDashboardUserSignature_(user, cleanupTabEnabled)) return null;
  var nowMs = new Date().getTime();
  var builtAtMs = swReadModelDateMs_(payload.builtAt);
  var expiresAtMs = swReadModelDateMs_(payload.expiresAt);
  if (!builtAtMs || !expiresAtMs || expiresAtMs < nowMs) return null;
  var invalidatedAt = swTaskDashboardInvalidatedAt_(ss);
  if (invalidatedAt && swReadModelDateMs_(invalidatedAt) > builtAtMs) return null;
  payload.ageSeconds = Math.max(0, Math.round((nowMs - builtAtMs) / 1000));
  payload.source = 'taskDashboardProjection';
  return payload;
}

function swTaskDashboardProjectionUsers_(ss) {
  var rows = [];
  try {
    rows = swAuthReadUserRows_(ss, true);
  } catch (_) {}
  var byEmail = {};
  rows.forEach(function (row) {
    if (!swTruthy_(row['Active?'] || '')) return;
    var user = swAuthUserFromRow_(row);
    if (!user.email || byEmail[user.email]) return;
    byEmail[user.email] = user;
  });
  return Object.keys(byEmail).sort().map(function (email) { return byEmail[email]; });
}

function swPutTaskDashboardProjection_(ss, user, cleanupTabEnabled, projectionType, payload) {
  var key = swTaskDashboardProjectionKey_(ss, user, cleanupTabEnabled, projectionType);
  swTaskDashboardProjectionCachePut_(key, payload);
  return key;
}

function swTaskDashboardProjectionKey_(ss, user, cleanupTabEnabled, projectionType) {
  return 'sw:taskDashboard:v1:' + ss.getId() + ':' +
    encodeURIComponent(swNormEmail_(user && user.email || '')) + ':' +
    encodeURIComponent(swTaskDashboardUserSignature_(user, cleanupTabEnabled)) + ':' +
    encodeURIComponent(swTrim_(projectionType || 'bootstrap'));
}

function swTaskDashboardUserSignature_(user, cleanupTabEnabled) {
  user = user || {};
  var parts = [
    swNormEmail_(user.email || ''),
    swNorm_(user.name || ''),
    (user.roles || []).map(swNorm_).sort().join(','),
    user.isAdmin ? 'admin' : '',
    user.isJoc ? 'joc' : '',
    user.isRep ? 'rep' : '',
    user.isDiamondOrderAdmin ? 'diamondAdmin' : '',
    user.isDiamondOrderAssistant ? 'diamondAssistant' : '',
    cleanupTabEnabled ? 'cleanupY' : 'cleanupN'
  ];
  return parts.join('|').replace(/[^a-z0-9@._|,-]+/gi, '').slice(0, 220);
}

function swTaskDashboardProjectionCacheGet_(key) {
  try {
    var memory = SW_TASK_DASHBOARD_MEMORY_CACHE_[key];
    if (memory && memory.expiresAt > new Date().getTime()) return memory.payload || null;
  } catch (_) {}
  var payload = swTaskListCacheGet_(key);
  if (!payload) return null;
  try {
    SW_TASK_DASHBOARD_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_TASK_DASHBOARD_CACHE_SECONDS * 1000,
      payload: payload
    };
  } catch (_) {}
  return payload;
}

function swTaskDashboardProjectionCachePut_(key, payload) {
  try {
    SW_TASK_DASHBOARD_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_TASK_DASHBOARD_CACHE_SECONDS * 1000,
      payload: payload
    };
  } catch (_) {}
  swTaskListCachePut_(key, payload);
}

function swPutTaskDashboardProjectionIndex_(ss, keys) {
  try {
    CacheService.getScriptCache().put(swTaskDashboardProjectionIndexKey_(ss), swStringify_(keys || []), SW_TASK_DASHBOARD_CACHE_SECONDS);
  } catch (_) {}
}

function swClearTaskDashboardProjectionCaches_(ss) {
  var keys = [];
  try {
    var text = CacheService.getScriptCache().get(swTaskDashboardProjectionIndexKey_(ss));
    keys = text ? swParseJson_(text, []) : [];
  } catch (_) {}
  (keys || []).forEach(function (key) {
    try { delete SW_TASK_DASHBOARD_MEMORY_CACHE_[key]; } catch (_) {}
    try { swTaskListCacheRemove_(key); } catch (_) {}
  });
  try { CacheService.getScriptCache().remove(swTaskDashboardProjectionIndexKey_(ss)); } catch (_) {}
}

function swTaskDashboardProjectionIndexKey_(ss) {
  return 'sw:taskDashboard:index:v1:' + ss.getId();
}

function swInvalidateTaskDashboardProjectionCache_(ss) {
  swClearTaskDashboardProjectionCaches_(ss);
  try {
    CacheService.getScriptCache().put(swTaskDashboardInvalidationKey_(ss), swIso_(new Date()), SW_TASK_DASHBOARD_CACHE_SECONDS);
  } catch (_) {}
}

function swClearTaskDashboardInvalidation_(ss) {
  try { CacheService.getScriptCache().remove(swTaskDashboardInvalidationKey_(ss)); } catch (_) {}
}

function swTaskDashboardInvalidatedAt_(ss) {
  try {
    return CacheService.getScriptCache().get(swTaskDashboardInvalidationKey_(ss)) || '';
  } catch (_) {}
  return '';
}

function swTaskDashboardInvalidationKey_(ss) {
  return 'sw:taskDashboard:invalidated:v1:' + ss.getId();
}

function swMergeObjects_(base, extra) {
  var out = {};
  Object.keys(base || {}).forEach(function (key) { out[key] = base[key]; });
  Object.keys(extra || {}).forEach(function (key) { out[key] = extra[key]; });
  return out;
}
