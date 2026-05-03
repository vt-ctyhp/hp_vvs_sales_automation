/**
 * Sales workflow public API: functions called by the web app, triggers, and admins.
 */

// Setup and generation.

/**
 * Mutating setup: ensures workflow sheets, styling, config rows, and templates exist.
 */
function sw_setupSalesWorkflow() {
  var ss = swSpreadsheet_();
  var taskSheet = swEnsureSheet_(ss, SW_SHEETS.TASKS, SW_TASK_HEADERS);
  var logSheet = swEnsureSheet_(ss, SW_SHEETS.LOG, SW_LOG_HEADERS);
  var configSheet = swEnsureSheet_(ss, SW_SHEETS.CONFIG, SW_CONFIG_HEADERS);
  var templateSheet = swEnsureSheet_(ss, SW_SHEETS.TEMPLATES, SW_TEMPLATE_HEADERS);

  swStyleSheet_(taskSheet);
  swStyleSheet_(logSheet);
  swStyleSheet_(configSheet);
  swStyleSheet_(templateSheet);

  swSeedConfig_(configSheet);
  swSeedTemplates_(templateSheet);

  return {
    ok: true,
    sheets: SW_SHEETS,
    message: 'Sales workflow sheets are ready.'
  };
}

/**
 * Mutating generation: creates or updates workflow tasks from master appointments.
 */
function sw_generateSalesWorkflowTasks() {
  return swTimed_('sw_generateSalesWorkflowTasks', function () {
    sw_setupSalesWorkflow();

    var ss = swSpreadsheet_();
    var ctx = swBuildContext_(ss, true);
    if (swNorm_(swConfigValue_(ctx.config, 'SYSTEM', 'FEATURE_ENABLED', 'Y')) === 'n') {
      return {
        ok: true,
        generatedAt: swIso_(new Date()),
        scannedAppointments: 0,
        created: 0,
        updated: 0,
        blocked: 0,
        skippedOld: 0,
        systemCompleted: 0,
        paused: true
      };
    }
    var masterRows = swReadAppointments_(ss);
    var taskState = swReadTaskState_(ss);
    swBeginDeferredTaskWrites_(ss, taskState);
    var now = new Date();
    var summary = {
      ok: true,
      generatedAt: swIso_(now),
      scannedAppointments: masterRows.length,
      created: 0,
      updated: 0,
      blocked: 0,
      skippedOld: 0,
      systemCompleted: 0
    };

    masterRows.forEach(function (rec) {
      if (!rec.root && !rec.appt) return;
      if (!swIsWorkflowRelevant_(rec, now, ctx)) {
        summary.skippedOld++;
        return;
      }

      if (!swIsAppointmentActive_(rec)) {
        summary.blocked += swBlockTasksForAppointment_(ss, taskState, rec, 'Appointment is no longer active/current.');
        return;
      }

      swGenerateTasksForAppointment_(ss, taskState, ctx, rec, now, summary);
    });

    swFlushDeferredTaskWrites_(ss, taskState);
    return summary;
  });
}

/**
 * Mutating generation wrapper: refreshes owner assignment through the normal generator.
 */
function sw_refreshTaskOwners() {
  var summary = sw_generateSalesWorkflowTasks();
  summary.ownerRefresh = true;
  return summary;
}

/**
 * Mutating setup: replaces Sales Workflow generation and owner-refresh triggers.
 */
function sw_installSalesWorkflowTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    var fn = trigger.getHandlerFunction();
    if (fn === 'sw_generateSalesWorkflowTasks' || fn === 'sw_refreshTaskOwners') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('sw_generateSalesWorkflowTasks').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('sw_refreshTaskOwners').timeBased().everyDays(1).atHour(7).create();
  return {
    ok: true,
    message: 'Installed hourly task generation and daily 7am owner refresh.'
  };
}

// Read-only UI calls.

/**
 * Read-only UI bootstrap: returns current user, view counts, and initial My Queue tasks.
 */
function sw_getBootstrap() {
  return swTimed_('sw_getBootstrap', function () {
    var mark = swStepTimer_('sw_getBootstrap');
    var ss = swSpreadsheet_();
    mark('spreadsheet');
    swRequireWorkflowReadSheets_(ss, { templates: false });
    mark('requiredSheets');
    var ctx = swBuildIdentityContext_(ss, true);
    mark('identity');
    var user = swCurrentUser_(ss, ctx);
    mark('currentUser', { isAdmin: user.isAdmin, isJoc: user.isJoc });
    var state = swReadTaskListState_(ss, true);
    mark('taskListRead', { tasks: state.tasks.length });
    var buckets = swBuildVisibleTaskBuckets_(state, user);
    mark('taskBuckets', {
      mine: buckets.mine.length,
      coverage: buckets.coverage.length,
      admin: buckets.admin.length
    });
    return {
      ok: true,
      user: user,
      tasks: buckets.mine,
      counts: {
        mine: buckets.mine.length,
        coverage: buckets.coverage.length,
        admin: user.isAdmin ? buckets.admin.length : 0
      },
      views: {
        mine: true,
        coverage: user.isJoc || user.isAdmin,
        admin: user.isAdmin
      },
      message: 'Connected. Use Generate Tasks to create or refresh the queue.'
    };
  });
}

/**
 * Read-only UI list: returns tasks visible to the current user for the requested view.
 */
function sw_getMyTasks(view) {
  return swTimed_('sw_getMyTasks', function () {
    var mark = swStepTimer_('sw_getMyTasks');
    var viewName = view || 'mine';
    var ss = swSpreadsheet_();
    mark('spreadsheet');
    swRequireWorkflowReadSheets_(ss, { templates: false });
    mark('requiredSheets');
    var user = swCurrentUserForTaskListView_(ss, viewName, true);
    mark('identity');
    mark('currentUser', { isAdmin: user.isAdmin, isJoc: user.isJoc });
    var state = swReadTaskListState_(ss, true);
    mark('taskListRead', { tasks: state.tasks.length });
    var tasks = swListVisibleTasksFromState_(state, user, viewName);
    mark('filter', { view: viewName, tasks: tasks.length });
    return {
      ok: true,
      view: viewName,
      user: user,
      tasks: tasks
    };
  });
}

/**
 * Read-only admin list: returns filterable admin-visible tasks.
 */
function sw_adminGetTasks(filters) {
  return swTimed_('sw_adminGetTasks', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var ctx = swBuildIdentityContext_(ss, true);
    var user = swCurrentUser_(ss, ctx);
    if (!user.isAdmin) throw new Error('Admin access required.');
    var state = swReadTaskListState_(ss, true);
    var tasks = swListVisibleTasksFromState_(state, user, 'admin');
    filters = filters || {};
    if (filters.status) {
      tasks = tasks.filter(function (t) { return t.status === filters.status; });
    }
    if (filters.ownerRole) {
      tasks = tasks.filter(function (t) { return t.ownerRole === filters.ownerRole; });
    }
    return { ok: true, tasks: tasks };
  });
}

/**
 * Read-only detail: returns task payload, rendered template data, and allowed actions.
 */
function sw_getTaskDetail(taskId) {
  return swTimed_('sw_getTaskDetail', function () {
    var mark = swStepTimer_('sw_getTaskDetail');
    var ss = swSpreadsheet_();
    mark('spreadsheet');
    swRequireWorkflowReadSheets_(ss);
    mark('requiredSheets');
    var ctx = swBuildTaskDetailContext_(ss, true);
    mark('detailContext');
    var user = swCurrentUser_(ss, ctx);
    mark('currentUser', { isAdmin: user.isAdmin, isJoc: user.isJoc });
    var task = swReadTaskRowById_(ss, taskId, true);
    mark('taskRowLookup');
    if (!task) throw new Error('Task not found: ' + taskId);
    if (!swCanViewTask_(task, user)) throw new Error('You do not have access to this task.');

    var payload = swParseJson_(task.payloadJson, {});
    mark('payloadParse');
    var template = ctx.templates[task.taskType] || swDefaultTemplate_(task.taskType);
    var renderData = swRenderDataForTask_(task, payload);
    var renderedTemplate = template.template ? swRenderTemplate_(template.template, renderData) : '';
    var renderedAttachmentUrl = template.attachmentUrl ? swRenderTemplate_(template.attachmentUrl, renderData) : '';
    var renderedAttachmentLabel = template.attachmentLabel ? swRenderTemplate_(template.attachmentLabel, renderData) : '';
    var attachments = swAttachmentsForTask_(task, template, renderData);
    var missingFields = swMissingFieldsForTask_(task, template, renderData);
    var checklist = swParseJson_(template.checklistJson, []);
    mark('render');

    return {
      ok: true,
      user: user,
      task: task,
      payload: payload,
      renderedTemplate: renderedTemplate,
      attachment: {
        label: renderedAttachmentLabel,
        url: renderedAttachmentUrl
      },
      attachments: attachments,
      missingFields: missingFields,
      checklist: checklist,
      canComplete: swCanActOnTask_(task, user),
      canClaim: swCanClaimTask_(task, user),
      canAdmin: user.isAdmin
    };
  });
}

// Task actions.

/**
 * Mutating task action: marks acknowledge-style tasks complete through the standard path.
 */
function sw_acknowledgeTask(taskId, data) {
  data = data || {};
  data.acknowledged = true;
  return sw_completeTask(taskId, data);
}

/**
 * Mutating task action: validates and completes a pending task, then refreshes generation.
 */
function sw_completeTask(taskId, data) {
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swCurrentUser_(ss);
  var task = swGetTaskById_(ss, taskId);
  if (!task) throw new Error('Task not found: ' + taskId);
  if (!swCanActOnTask_(task, user)) throw new Error('You are not the current owner for this task.');
  if (task.status !== SW_STATUSES.PENDING) throw new Error('Only pending tasks can be completed.');

  data = data || {};
  swValidateCompletion_(ss, task, data);

  var template = swTemplateForType_(ss, task.taskType);
  var payload = swParseJson_(task.payloadJson, {});
  var renderData = swRenderDataForTask_(task, payload);
  var renderedTemplate = template.template ? swRenderTemplate_(template.template, renderData) : '';
  var renderedAttachments = swAttachmentsForTask_(task, template, renderData);
  payload.completion = data;
  payload.renderedTemplate = renderedTemplate;
  payload.renderedAttachments = renderedAttachments;
  payload.completedBy = user.name || user.email;
  payload.completedByEmail = user.email;
  payload.completedAt = swIso_(new Date());

  var oldOwner = task.currentOwner;
  task.status = SW_STATUSES.COMPLETED;
  task.completedBy = user.name || user.email;
  task.completedByEmail = user.email;
  task.completedAt = payload.completedAt;
  task.updatedAt = payload.completedAt;
  task.lastEvent = 'COMPLETE';
  task.payloadJson = swStringify_(payload);
  swWriteTaskRow_(ss, task);
  swAppendTaskLog_(ss, 'COMPLETE', task, user, oldOwner, task.currentOwner, data);

  var generation = sw_generateSalesWorkflowTasks();
  return {
    ok: true,
    task: swGetTaskById_(ss, taskId),
    generation: generation
  };
}

/**
 * Mutating task action: lets an eligible user claim a pending coverage task.
 */
function sw_claimTask(taskId) {
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swCurrentUser_(ss);
  var task = swGetTaskById_(ss, taskId);
  if (!task) throw new Error('Task not found: ' + taskId);
  if (!swCanClaimTask_(task, user)) throw new Error('This task is not available for you to claim.');

  var fromOwner = task.currentOwner;
  var now = swIso_(new Date());
  task.currentOwner = user.name || user.email;
  task.currentOwnerEmail = user.email;
  task.coverageReason = task.coverageReason || 'CLAIMED_FROM_COVERAGE';
  task.claimedBy = user.name || user.email;
  task.claimedAt = now;
  task.updatedAt = now;
  task.lastEvent = 'CLAIM';
  swWriteTaskRow_(ss, task);
  swAppendTaskLog_(ss, 'CLAIM', task, user, fromOwner, task.currentOwner, {});
  return { ok: true, task: swGetTaskById_(ss, taskId) };
}

/**
 * Mutating task action: records that the user copied a task template.
 */
function sw_logTemplateCopied(taskId) {
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swCurrentUser_(ss);
  var task = swGetTaskById_(ss, taskId);
  if (!task) throw new Error('Task not found: ' + taskId);
  if (!swCanViewTask_(task, user)) throw new Error('You do not have access to this task.');
  swAppendTaskLog_(ss, 'TEMPLATE_COPY', task, user, task.currentOwner, task.currentOwner, {});
  return { ok: true };
}

// Admin actions.

/**
 * Mutating admin action: reassigns a pending task to a named owner.
 */
function sw_adminReassignTask(taskId, ownerName, ownerEmail, reason) {
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swCurrentUser_(ss);
  if (!user.isAdmin) throw new Error('Admin access required.');

  var task = swGetTaskById_(ss, taskId);
  if (!task) throw new Error('Task not found: ' + taskId);
  if (task.status !== SW_STATUSES.PENDING) throw new Error('Only pending tasks can be reassigned.');

  var fromOwner = task.currentOwner;
  var now = swIso_(new Date());
  task.currentOwner = swTrim_(ownerName);
  task.currentOwnerEmail = swTrim_(ownerEmail).toLowerCase();
  task.coverageReason = swTrim_(reason) || 'ADMIN_REASSIGNED';
  task.updatedAt = now;
  task.lastEvent = 'REASSIGN';
  swWriteTaskRow_(ss, task);
  swAppendTaskLog_(ss, 'REASSIGN', task, user, fromOwner, task.currentOwner, {
    reason: reason || ''
  });
  return { ok: true, task: swGetTaskById_(ss, taskId) };
}

/**
 * Mutating admin action: blocks a task and records the reason.
 */
function sw_adminBlockTask(taskId, reason) {
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swCurrentUser_(ss);
  if (!user.isAdmin) throw new Error('Admin access required.');

  var task = swGetTaskById_(ss, taskId);
  if (!task) throw new Error('Task not found: ' + taskId);

  var oldStatus = task.status;
  task.status = SW_STATUSES.BLOCKED;
  task.coverageReason = swTrim_(reason) || 'ADMIN_BLOCKED';
  task.updatedAt = swIso_(new Date());
  task.lastEvent = 'BLOCK';
  swWriteTaskRow_(ss, task);
  swAppendTaskLog_(ss, 'BLOCK', task, user, task.currentOwner, task.currentOwner, {
    fromStatus: oldStatus,
    reason: reason || ''
  });
  return { ok: true, task: swGetTaskById_(ss, taskId) };
}

/**
 * Mutating admin action: returns a blocked task to pending status.
 */
function sw_adminUnblockTask(taskId, reason) {
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swCurrentUser_(ss);
  if (!user.isAdmin) throw new Error('Admin access required.');

  var task = swGetTaskById_(ss, taskId);
  if (!task) throw new Error('Task not found: ' + taskId);

  task.status = SW_STATUSES.PENDING;
  task.coverageReason = swTrim_(reason) || '';
  task.updatedAt = swIso_(new Date());
  task.lastEvent = 'UNBLOCK';
  swWriteTaskRow_(ss, task);
  swAppendTaskLog_(ss, 'UNBLOCK', task, user, task.currentOwner, task.currentOwner, {
    reason: reason || ''
  });
  return { ok: true, task: swGetTaskById_(ss, taskId) };
}

// Diagnostics and tests.

/**
 * Read-only diagnostic: logs server-side speed for initial queue load paths.
 */
function sw_measureSalesWorkflowSpeed() {
  var started = new Date().getTime();
  var out = {
    ok: true,
    generatedAt: swIso_(new Date()),
    readOnly: true,
    note: 'Measures Apps Script server time only. Browser/network rendering time is not included.',
    steps: []
  };
  var bootstrap = null;
  var mine = null;
  var coverage = null;
  var admin = null;

  swBenchmarkSalesWorkflowStep_(out, 'sw_getBootstrap', function () {
    bootstrap = sw_getBootstrap();
    return {
      counts: bootstrap.counts || {},
      views: bootstrap.views || {},
      initialTasks: bootstrap.tasks ? bootstrap.tasks.length : 0
    };
  });

  swBenchmarkSalesWorkflowStep_(out, 'sw_getMyTasks:mine', function () {
    mine = sw_getMyTasks('mine');
    return swBenchmarkSalesWorkflowListSummary_(mine);
  });

  if (bootstrap && bootstrap.views && bootstrap.views.coverage) {
    swBenchmarkSalesWorkflowStep_(out, 'sw_getMyTasks:coverage', function () {
      coverage = sw_getMyTasks('coverage');
      return swBenchmarkSalesWorkflowListSummary_(coverage);
    });
  }

  if (bootstrap && bootstrap.views && bootstrap.views.admin) {
    swBenchmarkSalesWorkflowStep_(out, 'sw_getMyTasks:admin', function () {
      admin = sw_getMyTasks('admin');
      return swBenchmarkSalesWorkflowListSummary_(admin);
    });
  }

  var detailTaskId = swBenchmarkSalesWorkflowFirstTaskId_([mine, coverage, admin]);
  if (detailTaskId) {
    swBenchmarkSalesWorkflowStep_(out, 'sw_getTaskDetail', function () {
      var detail = sw_getTaskDetail(detailTaskId);
      return {
        taskId: detailTaskId,
        taskType: detail.task ? detail.task.taskType : '',
        attachments: detail.attachments ? detail.attachments.length : 0,
        missingFields: detail.missingFields ? detail.missingFields.length : 0
      };
    });
  } else {
    out.steps.push({
      operation: 'sw_getTaskDetail',
      skipped: true,
      reason: 'No visible task was available for detail timing.'
    });
  }

  out.totalMs = new Date().getTime() - started;
  Logger.log('SW_BENCHMARK_SUMMARY ' + JSON.stringify(out, null, 2));
  return out;
}

/**
 * Mutating diagnostic: runs setup and generation twice to confirm duplicate-safe generation.
 */
function sw_testSalesWorkflowDryRun() {
  var setup = sw_setupSalesWorkflow();
  var first = sw_generateSalesWorkflowTasks();
  var second = sw_generateSalesWorkflowTasks();
  Logger.log('SALES_WORKFLOW_TEST_DRY_RUN ' + JSON.stringify({
    setup: setup,
    first: first,
    second: second,
    duplicateSafe: second.created === 0
  }, null, 2));
  return { ok: true, setup: setup, first: first, second: second };
}

function swBenchmarkSalesWorkflowStep_(out, operation, fn) {
  var started = new Date().getTime();
  var step = {
    operation: operation,
    ok: true
  };
  try {
    step.result = fn() || {};
  } catch (err) {
    step.ok = false;
    step.error = err && err.message ? err.message : String(err);
  } finally {
    step.ms = new Date().getTime() - started;
    out.steps.push(step);
    Logger.log('SW_BENCHMARK_STEP ' + JSON.stringify(step));
  }
  return step;
}

function swBenchmarkSalesWorkflowListSummary_(res) {
  return {
    view: res && res.view ? res.view : '',
    tasks: res && res.tasks ? res.tasks.length : 0
  };
}

function swBenchmarkSalesWorkflowFirstTaskId_(responses) {
  for (var i = 0; i < responses.length; i++) {
    var tasks = responses[i] && responses[i].tasks ? responses[i].tasks : [];
    if (tasks.length && tasks[0].taskId) return tasks[0].taskId;
  }
  return '';
}
