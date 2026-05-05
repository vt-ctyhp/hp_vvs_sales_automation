/**
 * Sales workflow task repository: task state, queue projections, row writes, and task logs.
 */

var SW_TASK_LIST_CACHE_SECONDS = 2 * 60;
var SW_TASK_LIST_CACHE_CHUNK_SIZE = 75000;
var SW_TASK_LIST_MEMORY_CACHE_ = {};

function swReadTaskState_(ss, readOnly, options) {
  var sh = readOnly
    ? swGetRequiredSheet_(ss, SW_SHEETS.TASKS)
    : swEnsureSheet_(ss, SW_SHEETS.TASKS, SW_TASK_HEADERS);
  var rows = swReadSheetObjects_(sh);
  var rawTasks = [];
  rows.forEach(function (r) {
    var t = swTaskFromRow_(r);
    if (t.taskId) {
      rawTasks.push(t);
    }
  });
  return swBuildTaskStateFromTasks_(rawTasks, rows, options);
}

function swReadTaskListState_(ss, readOnly, options) {
  options = options || {};
  if (readOnly && !options.includeDuplicates) {
    var cached = swReadCachedTaskListState_(ss);
    if (cached) return cached;
  }

  var sh = readOnly
    ? swGetRequiredSheet_(ss, SW_SHEETS.TASKS)
    : swEnsureSheet_(ss, SW_SHEETS.TASKS, SW_TASK_HEADERS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { rows: [], byId: {}, tasks: [] };

  var statusCol = swTaskHeaderColumn_('Status');
  var primaryActionCol = swTaskHeaderColumn_('Primary Action');
  var snoozeReasonCol = swTaskHeaderColumn_('Snooze Reason');
  if (statusCol <= 0 || sh.getLastColumn() < statusCol) return swReadTaskState_(ss, readOnly, options);

  var rowCount = lastRow - 1;
  var rows = sh.getRange(2, 1, rowCount, statusCol).getDisplayValues();
  var extraRows = [];
  if (primaryActionCol > 0 && snoozeReasonCol >= primaryActionCol && sh.getLastColumn() >= primaryActionCol) {
    var extraWidth = Math.min(sh.getLastColumn(), snoozeReasonCol) - primaryActionCol + 1;
    extraRows = sh.getRange(2, primaryActionCol, rowCount, extraWidth).getDisplayValues();
  }
  var rawTasks = [];

  for (var i = 0; i < rows.length; i++) {
    var t = swTaskListFromListValues_(rows[i], extraRows[i] || [], i + 2);
    if (!t.taskId) continue;
    rawTasks.push(t);
  }

  var state = swBuildTaskStateFromTasks_(rawTasks, [], options);
  if (readOnly && !options.includeDuplicates) swCacheTaskListState_(ss, state);
  return state;
}

function swReadTaskListForRoot_(ss, rootApptId) {
  var root = swTrim_(rootApptId);
  if (!root) return [];
  var cached = swReadCachedTaskListState_(ss);
  if (cached && cached.tasks) {
    return cached.tasks.filter(function (task) {
      return swTrim_(task.root) === root || swTrim_(task.appt) === root;
    });
  }

  var sh = swGetRequiredSheet_(ss, SW_SHEETS.TASKS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var rowCount = lastRow - 1;
  var rootCol = swTaskHeaderColumn_('RootApptID');
  var apptCol = swTaskHeaderColumn_('APPT_ID');
  var statusCol = swTaskHeaderColumn_('Status');
  var primaryActionCol = swTaskHeaderColumn_('Primary Action');
  var snoozeReasonCol = swTaskHeaderColumn_('Snooze Reason');
  if (rootCol <= 0 || apptCol <= 0 || statusCol <= 0) return [];

  var roots = sh.getRange(2, rootCol, rowCount, 1).getDisplayValues();
  var appts = sh.getRange(2, apptCol, rowCount, 1).getDisplayValues();
  var out = [];
  for (var i = 0; i < rowCount; i++) {
    if (swTrim_(roots[i][0]) !== root && swTrim_(appts[i][0]) !== root) continue;
    var rowNumber = i + 2;
    var coreRow = sh.getRange(rowNumber, 1, 1, statusCol).getDisplayValues()[0];
    var extraRow = [];
    if (primaryActionCol > 0 && snoozeReasonCol >= primaryActionCol && sh.getLastColumn() >= primaryActionCol) {
      var extraWidth = Math.min(sh.getLastColumn(), snoozeReasonCol) - primaryActionCol + 1;
      extraRow = sh.getRange(rowNumber, primaryActionCol, 1, extraWidth).getDisplayValues()[0];
    }
    var task = swTaskListFromListValues_(coreRow, extraRow, rowNumber);
    if (task.taskId) out.push(task);
  }
  return out;
}

function swReadCachedTaskListState_(ss) {
  var key = swTaskListCacheKey_(ss);
  try {
    var memory = SW_TASK_LIST_MEMORY_CACHE_[key];
    if (memory && memory.expiresAt > new Date().getTime()) return swBuildTaskStateFromTasks_(memory.tasks || [], [], {});
  } catch (_) {}

  var payload = swTaskListCacheGet_(key);
  if (!payload || !payload.tasks) return null;
  try {
    SW_TASK_LIST_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_TASK_LIST_CACHE_SECONDS * 1000,
      tasks: payload.tasks || []
    };
  } catch (_) {}
  return swBuildTaskStateFromTasks_(payload.tasks || [], [], {});
}

function swCacheTaskListState_(ss, state) {
  var key = swTaskListCacheKey_(ss);
  var tasks = state && state.tasks ? state.tasks : [];
  try {
    SW_TASK_LIST_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_TASK_LIST_CACHE_SECONDS * 1000,
      tasks: tasks
    };
  } catch (_) {}
  swTaskListCachePut_(key, { cachedAt: swIso_(new Date()), tasks: tasks });
}

function swInvalidateTaskListCache_(ss) {
  var key = swTaskListCacheKey_(ss);
  try { delete SW_TASK_LIST_MEMORY_CACHE_[key]; } catch (_) {}
  swTaskListCacheRemove_(key);
}

function swTaskListCacheKey_(ss) {
  return 'sw:taskListState:v2:' + ss.getId();
}

function swTaskListCacheGet_(key) {
  try {
    var cache = CacheService.getScriptCache();
    var metaText = cache.get(key + ':meta');
    var meta = metaText ? swParseJson_(metaText, null) : null;
    if (!meta || !meta.chunks) return null;
    var text = '';
    for (var i = 0; i < meta.chunks; i++) {
      var chunk = cache.get(key + ':' + i);
      if (chunk == null) return null;
      text += chunk;
    }
    return swParseJson_(text, null);
  } catch (_) {}
  return null;
}

function swTaskListCachePut_(key, payload) {
  try {
    var cache = CacheService.getScriptCache();
    swTaskListCacheRemove_(key);
    var text = swStringify_(payload);
    var chunks = Math.ceil(text.length / SW_TASK_LIST_CACHE_CHUNK_SIZE);
    if (!chunks || chunks > 20) return;
    for (var i = 0; i < chunks; i++) {
      cache.put(key + ':' + i, text.slice(i * SW_TASK_LIST_CACHE_CHUNK_SIZE, (i + 1) * SW_TASK_LIST_CACHE_CHUNK_SIZE), SW_TASK_LIST_CACHE_SECONDS);
    }
    cache.put(key + ':meta', swStringify_({ chunks: chunks }), SW_TASK_LIST_CACHE_SECONDS);
  } catch (_) {}
}

function swTaskListCacheRemove_(key) {
  try {
    var cache = CacheService.getScriptCache();
    var metaText = cache.get(key + ':meta');
    var meta = metaText ? swParseJson_(metaText, null) : null;
    var keys = [key + ':meta'];
    var chunks = meta && meta.chunks ? Number(meta.chunks) : 20;
    for (var i = 0; i < chunks; i++) keys.push(key + ':' + i);
    cache.removeAll(keys);
  } catch (_) {}
}

function swBuildTaskStateFromTasks_(rawTasks, rows, options) {
  options = options || {};
  var byId = {};
  var taskIds = [];
  var duplicateTaskIds = {};

  (rawTasks || []).forEach(function (t) {
    if (!t || !t.taskId) return;
    if (!byId[t.taskId]) {
      byId[t.taskId] = t;
      taskIds.push(t.taskId);
      return;
    }
    duplicateTaskIds[t.taskId] = true;
    byId[t.taskId] = swBetterTaskRecord_(byId[t.taskId], t);
  });

  var tasks = taskIds.map(function (taskId) { return byId[taskId]; }).filter(Boolean);
  var out = { rows: rows || [], byId: byId, tasks: tasks };
  if (options.includeDuplicates) {
    out.allTasks = rawTasks || [];
    out.duplicateTaskIds = Object.keys(duplicateTaskIds);
  }
  return out;
}

function swBetterTaskRecord_(a, b) {
  if (!a) return b;
  if (!b) return a;
  var ar = swTaskCanonicalRank_(a);
  var br = swTaskCanonicalRank_(b);
  if (br !== ar) return br > ar ? b : a;
  var au = swTaskTimestampValue_(a.updatedAt || a.createdAt);
  var bu = swTaskTimestampValue_(b.updatedAt || b.createdAt);
  if (bu !== au) return bu > au ? b : a;
  return (b.rowNumber || 0) > (a.rowNumber || 0) ? b : a;
}

function swTaskCanonicalRank_(t) {
  var statusRank = 0;
  if (t.status === SW_STATUSES.BLOCKED) statusRank = 100;
  if (t.status === SW_STATUSES.SNOOZED) statusRank = 150;
  if (t.status === SW_STATUSES.PENDING) statusRank = 200;
  if (t.status === SW_STATUSES.COMPLETED) statusRank = 300;
  if (t.claimedBy) statusRank += 20;
  return statusRank;
}

function swTaskTimestampValue_(value) {
  if (!value) return 0;
  var d = new Date(value);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function swTaskHeaderColumn_(header) {
  return SW_TASK_HEADERS.indexOf(header) + 1;
}

function swTaskListFromListValues_(coreRow, extraRow, rowNumber) {
  var primaryActionCol = swTaskHeaderColumn_('Primary Action');
  function val(header) {
    var idx = SW_TASK_HEADERS.indexOf(header);
    if (idx < 0) return '';
    if (idx < coreRow.length) return coreRow[idx];
    var extraIdx = idx - primaryActionCol + 1;
    return extraIdx >= 0 && extraIdx < extraRow.length ? extraRow[extraIdx] : '';
  }
  return {
    taskId: val('TaskID'),
    root: val('RootApptID'),
    appt: val('APPT_ID'),
    customerName: val('Customer Name'),
    brand: val('Brand'),
    visitDate: val('Visit Date'),
    visitTime: swFormatAppointmentTime_(val('Visit Time')),
    visitType: val('Visit Type'),
    lifecycleStage: val('Lifecycle Stage'),
    taskType: val('Task Type'),
    taskTitle: val('Task Title'),
    ownerRole: val('Owner Role'),
    intendedOwner: val('Intended Owner'),
    intendedOwnerEmail: swNormEmail_(val('Intended Owner Email')),
    currentOwner: val('Current Owner'),
    currentOwnerEmail: swNormEmail_(val('Current Owner Email')),
    coverageReason: val('Coverage Reason'),
    dueAt: val('Due At'),
    status: val('Status') || SW_STATUSES.PENDING,
    primaryAction: val('Primary Action'),
    snoozeUntil: val('Snooze Until'),
    snoozeReason: val('Snooze Reason'),
    rowNumber: rowNumber || 0
  };
}

function swTaskListFromValues_(row, primaryAction, rowNumber) {
  function val(header) {
    var idx = SW_TASK_HEADERS.indexOf(header);
    return idx >= 0 && idx < row.length ? row[idx] : '';
  }
  return {
    taskId: val('TaskID'),
    root: val('RootApptID'),
    appt: val('APPT_ID'),
    customerName: val('Customer Name'),
    brand: val('Brand'),
    visitDate: val('Visit Date'),
    visitTime: swFormatAppointmentTime_(val('Visit Time')),
    visitType: val('Visit Type'),
    lifecycleStage: val('Lifecycle Stage'),
    taskType: val('Task Type'),
    taskTitle: val('Task Title'),
    ownerRole: val('Owner Role'),
    intendedOwner: val('Intended Owner'),
    intendedOwnerEmail: swNormEmail_(val('Intended Owner Email')),
    currentOwner: val('Current Owner'),
    currentOwnerEmail: swNormEmail_(val('Current Owner Email')),
    coverageReason: val('Coverage Reason'),
    dueAt: val('Due At'),
    status: val('Status') || SW_STATUSES.PENDING,
    primaryAction: primaryAction || '',
    snoozeUntil: val('Snooze Until'),
    snoozeReason: val('Snooze Reason'),
    rowNumber: rowNumber || 0
  };
}

function swTaskFromRow_(r) {
  return {
    taskId: r['TaskID'] || '',
    root: r['RootApptID'] || '',
    appt: r['APPT_ID'] || '',
    customerName: r['Customer Name'] || '',
    brand: r['Brand'] || '',
    visitDate: r['Visit Date'] || '',
    visitTime: swFormatAppointmentTime_(r['Visit Time'] || ''),
    visitType: r['Visit Type'] || '',
    lifecycleStage: r['Lifecycle Stage'] || '',
    taskType: r['Task Type'] || '',
    taskTitle: r['Task Title'] || '',
    ownerRole: r['Owner Role'] || '',
    intendedOwner: r['Intended Owner'] || '',
    intendedOwnerEmail: swNormEmail_(r['Intended Owner Email'] || ''),
    currentOwner: r['Current Owner'] || '',
    currentOwnerEmail: swNormEmail_(r['Current Owner Email'] || ''),
    coverageReason: r['Coverage Reason'] || '',
    dueAt: r['Due At'] || '',
    status: r['Status'] || SW_STATUSES.PENDING,
    dependencyTaskId: r['Dependency TaskID'] || '',
    createdAt: r['Created At'] || '',
    updatedAt: r['Updated At'] || '',
    completedBy: r['Completed By'] || '',
    completedByEmail: swNormEmail_(r['Completed By Email'] || ''),
    completedAt: r['Completed At'] || '',
    claimedBy: r['Claimed By'] || '',
    claimedAt: r['Claimed At'] || '',
    lastEvent: r['Last Event'] || '',
    payloadJson: r['Payload JSON'] || '',
    templateKey: r['Template Key'] || '',
    instructions: r['Instructions'] || '',
    primaryAction: r['Primary Action'] || '',
    snoozeUntil: r['Snooze Until'] || '',
    snoozeReason: r['Snooze Reason'] || '',
    snoozedBy: r['Snoozed By'] || '',
    snoozedAt: r['Snoozed At'] || '',
    rowNumber: r.__rowNumber || 0
  };
}

function swListVisibleTasks_(ss, user, view) {
  var state = swReadTaskState_(ss);
  return swListVisibleTasksFromState_(state, user, view);
}

function swListVisibleTasksFromState_(state, user, view, options) {
  return swBuildVisibleTaskBuckets_(state, user, options)[view || 'mine'] || [];
}

function swBuildVisibleTaskBuckets_(state, user, options) {
  options = options || {};
  var now = new Date().getTime();
  var tasks = state.tasks || Object.keys(state.byId || {}).map(function (id) { return state.byId[id]; });
  var buckets = { mine: [], cleanup: [], coverage: [], admin: [] };
  tasks.forEach(function (t) {
    var due = swTaskDueForQueue_(t, now);
    var cleanupCampaignTask = typeof swIsDataCleanupTaskType_ === 'function' &&
      swIsDataCleanupTaskType_(t.taskType) &&
      swNorm_(t.lifecycleStage) === swNorm_('Cleanup Campaign');
    var jocCoverageTask = t.ownerRole === 'JOC' &&
      (!!t.coverageReason || swNorm_(t.currentOwner) === swNorm_('JOC Coverage'));
    if (options.cleanupCampaignTabEnabled && cleanupCampaignTask && due &&
        (user.isAdmin || (swTaskOwnedByUser_(t, user) && !jocCoverageTask))) {
      buckets.cleanup.push(t);
    }
    if (due && swTaskOwnedByUser_(t, user) && !(options.cleanupCampaignTabEnabled && cleanupCampaignTask)) {
      buckets.mine.push(t);
    }
    if ((user.isJoc || user.isAdmin) && due && jocCoverageTask) {
      buckets.coverage.push(t);
    }
    if (user.isAdmin && t.status !== SW_STATUSES.COMPLETED) {
      buckets.admin.push(t);
    }
  });

  buckets.mine = swSortAndPublishTasks_(buckets.mine, now);
  buckets.cleanup = swSortAndPublishTasks_(buckets.cleanup, now);
  buckets.coverage = swSortAndPublishTasks_(buckets.coverage, now);
  buckets.admin = swSortAndPublishTasks_(buckets.admin, now);
  return buckets;
}

function swTaskStateMayNeedNameIdentity_(state) {
  var tasks = state.tasks || [];
  for (var i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    if (!swTaskPendingLike_(t, new Date().getTime())) continue;
    if (t.currentOwnerEmail) continue;
    if (swNorm_(t.ownerRole) === swNorm_(SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN) ||
        swNorm_(t.ownerRole) === swNorm_(SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT)) continue;
    if (!swTrim_(t.currentOwner)) continue;
    if (swNorm_(t.currentOwner) === swNorm_('JOC Coverage')) continue;
    return true;
  }
  return false;
}

function swSortAndPublishTasks_(tasks, now) {
  tasks.sort(function (a, b) {
    var ao = swIsOverdue_(a, now) ? 0 : 1;
    var bo = swIsOverdue_(b, now) ? 0 : 1;
    if (ao !== bo) return ao - bo;
    return swDateValue_(a.dueAt) - swDateValue_(b.dueAt);
  });
  return tasks.map(function (t) { return swPublicTask_(t, now); });
}

function swPublicTask_(t, nowMs) {
  return {
    taskId: t.taskId,
    root: t.root,
    appt: t.appt,
    customerName: t.customerName,
    brand: t.brand,
    visitDate: t.visitDate,
    visitTime: t.visitTime,
    visitType: t.visitType,
    lifecycleStage: t.lifecycleStage,
    taskType: t.taskType,
    taskTitle: t.taskTitle,
    ownerRole: t.ownerRole,
    intendedOwner: t.intendedOwner,
    currentOwner: t.currentOwner,
    coverageReason: t.coverageReason,
    dueAt: t.dueAt,
    dueLabel: swDueLabel_(t, nowMs || new Date().getTime()),
    status: t.status,
    primaryAction: t.primaryAction,
    snoozeUntil: t.snoozeUntil || '',
    snoozeReason: t.snoozeReason || ''
  };
}

function swGetTaskById_(ss, taskId) {
  return swReadTaskRowById_(ss, taskId, false);
}

function swReadTaskRowById_(ss, taskId, readOnly) {
  var sh = readOnly
    ? swGetRequiredSheet_(ss, SW_SHEETS.TASKS)
    : swEnsureSheet_(ss, SW_SHEETS.TASKS, SW_TASK_HEADERS);
  if (readOnly) {
    var cachedTask = swCachedTaskListRowById_(ss, taskId);
    if (cachedTask && cachedTask.rowNumber) {
      var cachedRow = swReadTaskRowAtNumber_(sh, cachedTask.rowNumber);
      if (cachedRow && String(cachedRow.taskId) === String(taskId)) return cachedRow;
    }
    return swReadTaskRowByIdOnePass_(sh, taskId);
  }

  var rowNumber = swFindTaskRow_(sh, taskId);
  if (!rowNumber) return null;

  return swReadTaskRowAtNumber_(sh, rowNumber);
}

function swCachedTaskListRowById_(ss, taskId) {
  if (!taskId) return null;
  var cached = swReadCachedTaskListState_(ss);
  return cached && cached.byId ? cached.byId[taskId] : null;
}

function swReadTaskRowAtNumber_(sh, rowNumber) {
  rowNumber = Number(rowNumber);
  if (!isFinite(rowNumber) || rowNumber < 2 || rowNumber > sh.getLastRow()) return null;
  var colCount = Math.min(sh.getLastColumn(), SW_TASK_HEADERS.length);
  var values = sh.getRange(rowNumber, 1, 1, colCount).getDisplayValues()[0];
  return swTaskFromValues_(values, rowNumber);
}

function swReadTaskRowByIdOnePass_(sh, taskId) {
  return swFindCanonicalTaskInSheet_(sh, taskId);
}

function swTaskFromValues_(values, rowNumber) {
  var row = { __rowNumber: rowNumber };
  for (var i = 0; i < SW_TASK_HEADERS.length; i++) {
    row[SW_TASK_HEADERS[i]] = i < values.length ? values[i] : '';
  }
  return swTaskFromRow_(row);
}

function swBeginDeferredTaskWrites_(ss, state) {
  var taskSheet = swEnsureSheet_(ss, SW_SHEETS.TASKS, SW_TASK_HEADERS);
  var logSheet = swEnsureSheet_(ss, SW_SHEETS.LOG, SW_LOG_HEADERS);
  state.deferWrites = true;
  state.nextTaskRow = taskSheet.getLastRow() + 1;
  state.pendingTaskRows = [];
  state.pendingLogRows = [];
  state.taskAppendStartRow = state.nextTaskRow;
  state.logAppendStartRow = logSheet.getLastRow() + 1;
}

function swFlushDeferredTaskWrites_(ss, state) {
  if (!state || !state.deferWrites) return;
  var wroteTasks = false;
  if (state.pendingTaskRows && state.pendingTaskRows.length) {
    var taskSheet = swEnsureSheet_(ss, SW_SHEETS.TASKS, SW_TASK_HEADERS);
    taskSheet.getRange(state.taskAppendStartRow, 1, state.pendingTaskRows.length, SW_TASK_HEADERS.length)
      .setValues(state.pendingTaskRows);
    wroteTasks = true;
  }
  if (state.pendingLogRows && state.pendingLogRows.length) {
    var logSheet = swEnsureSheet_(ss, SW_SHEETS.LOG, SW_LOG_HEADERS);
    logSheet.getRange(logSheet.getLastRow() + 1, 1, state.pendingLogRows.length, SW_LOG_HEADERS.length)
      .setValues(state.pendingLogRows);
  }
  state.deferWrites = false;
  state.pendingTaskRows = [];
  state.pendingLogRows = [];
  if (wroteTasks) swInvalidateTaskListCache_(ss);
}

function swQueueOrAppendTaskRow_(ss, state, task) {
  if (state && state.deferWrites) {
    task.rowNumber = state.nextTaskRow++;
    state.pendingTaskRows.push(swTaskToRow_(task));
    return task.rowNumber;
  }
  return swAppendTaskRow_(ss, task);
}

function swQueueOrAppendTaskLog_(ss, state, eventType, task, actor, fromOwner, toOwner, details) {
  if (state && state.deferWrites) {
    state.pendingLogRows.push(swTaskLogRow_(eventType, task, actor, fromOwner, toOwner, details));
    return;
  }
  swAppendTaskLog_(ss, eventType, task, actor, fromOwner, toOwner, details);
}

function swAppendTaskRow_(ss, task) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.TASKS, SW_TASK_HEADERS);
  var row = swTaskToRow_(task);
  sh.appendRow(row);
  task.rowNumber = sh.getLastRow();
  swInvalidateTaskListCache_(ss);
  return task.rowNumber;
}

function swWriteTaskRow_(ss, task) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.TASKS, SW_TASK_HEADERS);
  if (!task.rowNumber) {
    var found = swFindTaskRow_(sh, task.taskId);
    task.rowNumber = found;
  }
  if (!task.rowNumber) {
    swAppendTaskRow_(ss, task);
    return;
  }
  sh.getRange(task.rowNumber, 1, 1, SW_TASK_HEADERS.length).setValues([swTaskToRow_(task)]);
  swInvalidateTaskListCache_(ss);
}

function swTaskToRow_(task) {
  var map = {
    'TaskID': task.taskId,
    'RootApptID': task.root,
    'APPT_ID': task.appt,
    'Customer Name': task.customerName,
    'Brand': task.brand,
    'Visit Date': task.visitDate,
    'Visit Time': task.visitTime,
    'Visit Type': task.visitType,
    'Lifecycle Stage': task.lifecycleStage,
    'Task Type': task.taskType,
    'Task Title': task.taskTitle,
    'Owner Role': task.ownerRole,
    'Intended Owner': task.intendedOwner,
    'Intended Owner Email': task.intendedOwnerEmail,
    'Current Owner': task.currentOwner,
    'Current Owner Email': task.currentOwnerEmail,
    'Coverage Reason': task.coverageReason,
    'Due At': task.dueAt,
    'Status': task.status,
    'Dependency TaskID': task.dependencyTaskId,
    'Created At': task.createdAt,
    'Updated At': task.updatedAt,
    'Completed By': task.completedBy,
    'Completed By Email': task.completedByEmail,
    'Completed At': task.completedAt,
    'Claimed By': task.claimedBy,
    'Claimed At': task.claimedAt,
    'Last Event': task.lastEvent,
    'Payload JSON': task.payloadJson,
    'Template Key': task.templateKey,
    'Instructions': task.instructions,
    'Primary Action': task.primaryAction,
    'Snooze Until': task.snoozeUntil,
    'Snooze Reason': task.snoozeReason,
    'Snoozed By': task.snoozedBy,
    'Snoozed At': task.snoozedAt
  };
  return SW_TASK_HEADERS.map(function (h) { return map[h] == null ? '' : map[h]; });
}

function swFindTaskRow_(sh, taskId) {
  var task = swFindCanonicalTaskInSheet_(sh, taskId);
  return task ? task.rowNumber : 0;
}

function swFindCanonicalTaskInSheet_(sh, taskId) {
  if (!taskId || sh.getLastRow() < 2) return null;
  var rowCount = sh.getLastRow() - 1;
  var colCount = Math.min(sh.getLastColumn(), SW_TASK_HEADERS.length);
  var ids = sh.getRange(2, 1, rowCount, 1).getDisplayValues();
  var best = null;
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) !== String(taskId)) continue;
    var rowNumber = i + 2;
    var row = sh.getRange(rowNumber, 1, 1, colCount).getDisplayValues()[0];
    best = swBetterTaskRecord_(best, swTaskFromValues_(row, rowNumber));
  }
  return best;
}

function swAppendTaskLog_(ss, eventType, task, actor, fromOwner, toOwner, details) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.LOG, SW_LOG_HEADERS);
  sh.appendRow(swTaskLogRow_(eventType, task, actor, fromOwner, toOwner, details));
}

function swTaskLogRow_(eventType, task, actor, fromOwner, toOwner, details) {
  actor = actor || swSystemUser_();
  return [
    swIso_(new Date()),
    eventType,
    task.taskId || '',
    task.root || '',
    task.appt || '',
    task.taskType || '',
    actor.name || '',
    actor.email || '',
    fromOwner || '',
    toOwner || '',
    task.status || '',
    swStringify_(details || {})
  ];
}

function swDuplicateTaskCleanupPlan_(state) {
  var byId = {};
  var allTasks = (state && state.allTasks) || (state && state.tasks) || [];
  allTasks.forEach(function (t) {
    if (!t || !t.taskId) return;
    if (!byId[t.taskId]) byId[t.taskId] = [];
    byId[t.taskId].push(t);
  });

  var groups = [];
  var rowsToBlock = [];
  Object.keys(byId).sort().forEach(function (taskId) {
    var group = byId[taskId];
    if (group.length < 2) return;
    var keep = group.reduce(function (best, t) {
      return swBetterTaskRecord_(best, t);
    }, null);
    var statusCounts = {};
    group.forEach(function (t) {
      var status = t.status || '(blank)';
      statusCounts[status] = (statusCounts[status] || 0) + 1;
    });
    var block = group.filter(function (t) {
      return keep && t.rowNumber !== keep.rowNumber &&
        t.status !== SW_STATUSES.COMPLETED &&
        t.status !== SW_STATUSES.BLOCKED;
    });
    block.forEach(function (t) {
      rowsToBlock.push({
        task: t,
        row: t.rowNumber,
        keepRow: keep.rowNumber,
        taskId: t.taskId,
        taskType: t.taskType,
        customerName: t.customerName,
        currentOwner: t.currentOwner,
        reason: 'Duplicate TaskID cleanup; kept row ' + keep.rowNumber + '.'
      });
    });
    groups.push({
      taskId: taskId,
      rowCount: group.length,
      rows: group.map(function (t) { return t.rowNumber; }),
      keepRow: keep ? keep.rowNumber : '',
      statuses: statusCounts,
      rowsToBlock: block.map(function (t) { return t.rowNumber; }),
      sample: swDuplicateTaskBrief_(keep || group[0])
    });
  });

  return {
    duplicateGroups: groups,
    rowsToBlock: rowsToBlock
  };
}

function swDuplicateTaskAuditOutput_(state, plan) {
  state = state || {};
  plan = plan || swDuplicateTaskCleanupPlan_(state);
  var allTasks = state.allTasks || state.tasks || [];
  var duplicateRows = plan.duplicateGroups.reduce(function (sum, g) {
    return sum + g.rowCount;
  }, 0);
  var statusMix = {};
  var extraRowsByType = {};
  plan.duplicateGroups.forEach(function (g) {
    var names = Object.keys(g.statuses).sort().join(' + ');
    statusMix[names] = (statusMix[names] || 0) + 1;
    var type = g.sample.taskType || '(blank)';
    extraRowsByType[type] = (extraRowsByType[type] || 0) + Math.max(0, g.rowCount - 1);
  });
  var pendingGroups = plan.duplicateGroups.filter(function (g) {
    return !!g.statuses[SW_STATUSES.PENDING];
  });
  return {
    ok: true,
    generatedAt: swIso_(new Date()),
    readOnly: true,
    summary: {
      totalPhysicalRows: allTasks.length,
      uniqueTaskIds: state.tasks ? state.tasks.length : 0,
      duplicateTaskIdGroups: plan.duplicateGroups.length,
      duplicateTaskIdRows: duplicateRows,
      extraDuplicateRows: duplicateRows - plan.duplicateGroups.length,
      duplicateGroupsWithPendingRows: pendingGroups.length,
      pendingRowsToBlock: plan.rowsToBlock.length,
      duplicateStatusMix: statusMix,
      extraRowsByType: extraRowsByType
    },
    examples: plan.duplicateGroups.slice(0, 50)
  };
}

function swCleanupDuplicateTasks_(apply) {
  var lock = LockService.getDocumentLock() || LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var state = swReadTaskState_(ss, false, { includeDuplicates: true });
    var plan = swDuplicateTaskCleanupPlan_(state);
    var out = swDuplicateTaskAuditOutput_(state, plan);
    out.readOnly = !apply;
    out.rowsPlannedToBlock = plan.rowsToBlock.map(function (item) {
      return swDuplicateTaskCleanupItem_(item);
    });

    if (!apply) {
      Logger.log('SW_DUPLICATE_TASK_CLEANUP_DRY_RUN ' + JSON.stringify(out, null, 2));
      return out;
    }

    var now = swIso_(new Date());
    plan.rowsToBlock.forEach(function (item) {
      var t = item.task;
      var fromOwner = t.currentOwner || '';
      t.status = SW_STATUSES.BLOCKED;
      t.coverageReason = item.reason;
      t.updatedAt = now;
      t.lastEvent = 'BLOCK';
      swWriteTaskRow_(ss, t);
      swAppendTaskLog_(ss, 'BLOCK', t, swSystemUser_(), fromOwner, t.currentOwner, {
        reason: 'Duplicate TaskID cleanup',
        keepRow: item.keepRow,
        duplicateRow: item.row
      });
    });

    out.applied = true;
    out.rowsBlocked = plan.rowsToBlock.length;
    Logger.log('SW_DUPLICATE_TASK_CLEANUP_APPLY ' + JSON.stringify(out, null, 2));
    return out;
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function swDuplicateTaskCleanupItem_(item) {
  return {
    row: item.row,
    keepRow: item.keepRow,
    taskId: item.taskId,
    taskType: item.taskType,
    customerName: item.customerName,
    currentOwner: item.currentOwner,
    reason: item.reason
  };
}

function swDuplicateTaskBrief_(t) {
  t = t || {};
  return {
    row: t.rowNumber || '',
    taskId: t.taskId || '',
    root: t.root || '',
    appt: t.appt || '',
    taskType: t.taskType || '',
    taskTitle: t.taskTitle || '',
    customerName: t.customerName || '',
    currentOwner: t.currentOwner || '',
    dueAt: t.dueAt || '',
    status: t.status || '',
    createdAt: t.createdAt || '',
    updatedAt: t.updatedAt || ''
  };
}
