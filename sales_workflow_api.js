/**
 * Sales workflow public API: functions called by the web app, triggers, and admins.
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

function sw_refreshTaskOwners() {
  var summary = sw_generateSalesWorkflowTasks();
  summary.ownerRefresh = true;
  return summary;
}

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

function sw_getMyTasks(view) {
  return swTimed_('sw_getMyTasks', function () {
    var mark = swStepTimer_('sw_getMyTasks');
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
    var tasks = swListVisibleTasksFromState_(state, user, view || 'mine');
    mark('filter', { view: view || 'mine', tasks: tasks.length });
    return {
      ok: true,
      view: view || 'mine',
      user: user,
      tasks: tasks
    };
  });
}

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

function sw_acknowledgeTask(taskId, data) {
  data = data || {};
  data.acknowledged = true;
  return sw_completeTask(taskId, data);
}

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
