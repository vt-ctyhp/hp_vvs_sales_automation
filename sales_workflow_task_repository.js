/**
 * Sales workflow task repository: task state, queue projections, row writes, and task logs.
 */

function swReadTaskState_(ss, readOnly) {
  var sh = readOnly
    ? swGetRequiredSheet_(ss, SW_SHEETS.TASKS)
    : swEnsureSheet_(ss, SW_SHEETS.TASKS, SW_TASK_HEADERS);
  var rows = swReadSheetObjects_(sh);
  var byId = {};
  var tasks = [];
  rows.forEach(function (r) {
    var t = swTaskFromRow_(r);
    if (t.taskId) {
      byId[t.taskId] = t;
      tasks.push(t);
    }
  });
  return { rows: rows, byId: byId, tasks: tasks };
}

function swReadTaskListState_(ss, readOnly) {
  var sh = readOnly
    ? swGetRequiredSheet_(ss, SW_SHEETS.TASKS)
    : swEnsureSheet_(ss, SW_SHEETS.TASKS, SW_TASK_HEADERS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { rows: [], byId: {}, tasks: [] };

  var statusCol = swTaskHeaderColumn_('Status');
  if (statusCol <= 0 || sh.getLastColumn() < statusCol) return swReadTaskState_(ss, readOnly);

  var main = sh.getRange(2, 1, lastRow - 1, statusCol).getDisplayValues();
  var actionCol = swTaskHeaderColumn_('Primary Action');
  var actions = actionCol > 0 && sh.getLastColumn() >= actionCol
    ? sh.getRange(2, actionCol, lastRow - 1, 1).getDisplayValues()
    : [];
  var byId = {};
  var tasks = [];

  for (var i = 0; i < main.length; i++) {
    var t = swTaskListFromValues_(main[i], actions[i] ? actions[i][0] : '', i + 2);
    if (!t.taskId) continue;
    byId[t.taskId] = t;
    tasks.push(t);
  }

  return { rows: [], byId: byId, tasks: tasks };
}

function swTaskHeaderColumn_(header) {
  return SW_TASK_HEADERS.indexOf(header) + 1;
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
    visitTime: val('Visit Time'),
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
    visitTime: r['Visit Time'] || '',
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
    rowNumber: r.__rowNumber || 0
  };
}

function swListVisibleTasks_(ss, user, view) {
  var state = swReadTaskState_(ss);
  return swListVisibleTasksFromState_(state, user, view);
}

function swListVisibleTasksFromState_(state, user, view) {
  return swBuildVisibleTaskBuckets_(state, user)[view || 'mine'] || [];
}

function swBuildVisibleTaskBuckets_(state, user) {
  var now = new Date().getTime();
  var tasks = state.tasks || Object.keys(state.byId || {}).map(function (id) { return state.byId[id]; });
  var buckets = { mine: [], coverage: [], admin: [] };
  tasks.forEach(function (t) {
    if (t.status === SW_STATUSES.PENDING && swTaskDueForQueue_(t, now) && swTaskOwnedByUser_(t, user)) {
      buckets.mine.push(t);
    }
    if ((user.isJoc || user.isAdmin) && t.ownerRole === 'JOC' &&
      swTaskDueForQueue_(t, now) &&
      t.status === SW_STATUSES.PENDING &&
      (!!t.coverageReason || swNorm_(t.currentOwner) === swNorm_('JOC Coverage'))) {
      buckets.coverage.push(t);
    }
    if (user.isAdmin && t.status !== SW_STATUSES.COMPLETED) {
      buckets.admin.push(t);
    }
  });

  buckets.mine = swSortAndPublishTasks_(buckets.mine, now);
  buckets.coverage = swSortAndPublishTasks_(buckets.coverage, now);
  buckets.admin = swSortAndPublishTasks_(buckets.admin, now);
  return buckets;
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
    primaryAction: t.primaryAction
  };
}

function swGetTaskById_(ss, taskId) {
  return swReadTaskRowById_(ss, taskId, false);
}

function swReadTaskRowById_(ss, taskId, readOnly) {
  var sh = readOnly
    ? swGetRequiredSheet_(ss, SW_SHEETS.TASKS)
    : swEnsureSheet_(ss, SW_SHEETS.TASKS, SW_TASK_HEADERS);
  var rowNumber = swFindTaskRow_(sh, taskId);
  if (!rowNumber) return null;

  var colCount = Math.min(sh.getLastColumn(), SW_TASK_HEADERS.length);
  var values = sh.getRange(rowNumber, 1, 1, colCount).getDisplayValues()[0];
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
  if (state.pendingTaskRows && state.pendingTaskRows.length) {
    var taskSheet = swEnsureSheet_(ss, SW_SHEETS.TASKS, SW_TASK_HEADERS);
    taskSheet.getRange(state.taskAppendStartRow, 1, state.pendingTaskRows.length, SW_TASK_HEADERS.length)
      .setValues(state.pendingTaskRows);
  }
  if (state.pendingLogRows && state.pendingLogRows.length) {
    var logSheet = swEnsureSheet_(ss, SW_SHEETS.LOG, SW_LOG_HEADERS);
    logSheet.getRange(logSheet.getLastRow() + 1, 1, state.pendingLogRows.length, SW_LOG_HEADERS.length)
      .setValues(state.pendingLogRows);
  }
  state.deferWrites = false;
  state.pendingTaskRows = [];
  state.pendingLogRows = [];
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
    'Primary Action': task.primaryAction
  };
  return SW_TASK_HEADERS.map(function (h) { return map[h] == null ? '' : map[h]; });
}

function swFindTaskRow_(sh, taskId) {
  if (sh.getLastRow() < 2) return 0;
  var idRange = sh.getRange(2, 1, sh.getLastRow() - 1, 1);
  try {
    var found = idRange
      .createTextFinder(String(taskId))
      .useRegularExpression(false)
      .matchEntireCell(true)
      .matchCase(true)
      .findNext();
    if (found) return found.getRow();
  } catch (_) {}

  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(taskId)) return i + 2;
  }
  return 0;
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
