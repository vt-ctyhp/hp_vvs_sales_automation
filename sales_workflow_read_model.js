/**
 * Sales workflow read models.
 *
 * Phase 1 is shadow-only: these generated tabs are built for benchmarking and
 * comparison, but the web app still serves from the existing source sheets.
 */

var SW_READ_MODEL_VERSION = 'phase1-v1';
var SW_READ_MODEL_DEFAULT_TTL_SECONDS = 5 * 60;
var SW_READ_MODEL_REFRESH_HANDLER = 'sw_rebuildWorkflowReadModels';

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
  var sh = ss.getSheetByName(SW_SHEETS.READ_MODEL_META);
  var now = swIso_(new Date());
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
  var rowCount = sh.getLastRow() - 1;
  if (statusCol > 0) sh.getRange(2, statusCol, rowCount, 1).setValue('STALE');
  if (invalidatedCol > 0) sh.getRange(2, invalidatedCol, rowCount, 1).setValue(now);
  if (notesCol > 0) sh.getRange(2, notesCol, rowCount, 1).setValue(swTrim_(reason || 'Manual invalidation'));

  return {
    ok: true,
    invalidated: true,
    models: rowCount,
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
  return Math.max(60, Math.min(Math.floor(ttl), 30 * 60));
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
