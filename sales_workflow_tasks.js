/**
 * Sales Appointment Workflow Web App.
 *
 * Creates a role-aware task ledger for the core appointment workflow. The
 * master sheet remains the historical record; this module provides the daily
 * work queue, ownership, coverage, templates, and audit trail.
 */

var SW_SHEETS = {
  MASTER: '00_Master Appointments',
  TASKS: '_SalesTaskQueue',
  LOG: '_SalesTaskLog',
  CONFIG: '_SalesWorkflowConfig',
  TEMPLATES: '_SalesWorkflowTemplates',
  ROSTER: '10_Roster_Schedule',
  SCHEDULE_CHANGES: 'Schedule Changes',
  DROPDOWN: 'Dropdown'
};

var SW_STATUSES = {
  PENDING: 'Pending',
  COMPLETED: 'Completed',
  BLOCKED: 'Blocked'
};

var SW_TASKS = {
  ASSIGN: 'ASSIGN_APPOINTMENT',
  WELCOME: 'SEND_WELCOME',
  HYBRID: 'SEND_HYBRID_WELCOME',
  MAP: 'SEND_MAP_INSTRUCTIONS',
  REVIEW: 'REVIEW_APPOINTMENT',
  CHECKLIST: 'APPOINTMENT_DAY_CHECKLIST',
  PROCESS: 'PROCESS_APPOINTMENT_DATA',
  APPROVE: 'APPROVE_RECAP_MESSAGE',
  FINAL: 'SEND_FINAL_RECAP'
};

var SW_TASK_HEADERS = [
  'TaskID',
  'RootApptID',
  'APPT_ID',
  'Customer Name',
  'Brand',
  'Visit Date',
  'Visit Time',
  'Visit Type',
  'Lifecycle Stage',
  'Task Type',
  'Task Title',
  'Owner Role',
  'Intended Owner',
  'Intended Owner Email',
  'Current Owner',
  'Current Owner Email',
  'Coverage Reason',
  'Due At',
  'Status',
  'Dependency TaskID',
  'Created At',
  'Updated At',
  'Completed By',
  'Completed By Email',
  'Completed At',
  'Claimed By',
  'Claimed At',
  'Last Event',
  'Payload JSON',
  'Template Key',
  'Instructions',
  'Primary Action'
];

var SW_LOG_HEADERS = [
  'Event At',
  'Event Type',
  'TaskID',
  'RootApptID',
  'APPT_ID',
  'Task Type',
  'Actor Name',
  'Actor Email',
  'From Owner',
  'To Owner',
  'Status',
  'Details JSON'
];

var SW_CONFIG_HEADERS = [
  'Section',
  'Key',
  'Value',
  'Role',
  'Name',
  'Email',
  'Active?',
  'Priority',
  'Notes'
];

var SW_TEMPLATE_HEADERS = [
  'Task Type',
  'Task Title',
  'Instructions',
  'Template',
  'Attachment Label',
  'Attachment URL',
  'Checklist JSON',
  'Primary Action'
];

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
  sw_setupSalesWorkflow();

  var ss = swSpreadsheet_();
  var ctx = swBuildContext_(ss);
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

  return summary;
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
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swCurrentUser_(ss);
  var tasks = swListVisibleTasks_(ss, user, 'mine');
  return {
    ok: true,
    user: user,
    counts: {
      mine: tasks.length,
      coverage: swListVisibleTasks_(ss, user, 'coverage').length,
      admin: user.isAdmin ? swListVisibleTasks_(ss, user, 'admin').length : 0
    },
    views: {
      mine: true,
      coverage: user.isJoc || user.isAdmin,
      admin: user.isAdmin
    },
    message: 'Connected. Use Generate Tasks to create or refresh the queue.'
  };
}

function sw_getMyTasks(view) {
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swCurrentUser_(ss);
  return {
    ok: true,
    view: view || 'mine',
    user: user,
    tasks: swListVisibleTasks_(ss, user, view || 'mine')
  };
}

function sw_adminGetTasks(filters) {
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swCurrentUser_(ss);
  if (!user.isAdmin) throw new Error('Admin access required.');
  var tasks = swListVisibleTasks_(ss, user, 'admin');
  filters = filters || {};
  if (filters.status) {
    tasks = tasks.filter(function (t) { return t.status === filters.status; });
  }
  if (filters.ownerRole) {
    tasks = tasks.filter(function (t) { return t.ownerRole === filters.ownerRole; });
  }
  return { ok: true, tasks: tasks };
}

function sw_getTaskDetail(taskId) {
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swCurrentUser_(ss);
  var task = swGetTaskById_(ss, taskId);
  if (!task) throw new Error('Task not found: ' + taskId);
  if (!swCanViewTask_(task, user)) throw new Error('You do not have access to this task.');

  var payload = swParseJson_(task.payloadJson, {});
  var template = swTemplateForType_(ss, task.taskType);
  var renderData = swRenderDataForTask_(task, payload);
  var renderedTemplate = template.template ? swRenderTemplate_(template.template, renderData) : '';
  var renderedAttachmentUrl = template.attachmentUrl ? swRenderTemplate_(template.attachmentUrl, renderData) : '';
  var renderedAttachmentLabel = template.attachmentLabel ? swRenderTemplate_(template.attachmentLabel, renderData) : '';
  var attachments = swAttachmentsForTask_(task, template, renderData);
  var missingFields = swMissingFieldsForTask_(task, template, renderData);
  var checklist = swParseJson_(template.checklistJson, []);

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
  var state = swReadTaskState_(ss);
  var task = state.byId[taskId];
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

function swGenerateTasksForAppointment_(ss, state, ctx, rec, now, summary) {
  var visitAt = swVisitDateTime_(rec, ctx.tz);
  var assign = swBuildTask_(ss, state, ctx, rec, SW_TASKS.ASSIGN, 'System', now, '', now, {});
  assign.status = SW_STATUSES.COMPLETED;
  assign.completedBy = 'System';
  assign.completedAt = assign.completedAt || swIso_(now);
  assign.lastEvent = assign.lastEvent || 'AUTO_COMPLETE';
  swUpsertTask_(ss, state, assign, summary);
  summary.systemCompleted++;

  var within24 = visitAt && visitAt.getTime() >= now.getTime() - (6 * 60 * 60 * 1000) &&
    visitAt.getTime() <= now.getTime() + (24 * 60 * 60 * 1000);
  var dueNow = now;

  if (within24) {
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.HYBRID, 'JOC', dueNow, assign.taskId, now, {}), summary);
  } else {
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.WELCOME, 'JOC', dueNow, assign.taskId, now, {}), summary);
    if (visitAt) {
      swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.MAP, 'JOC', swDateAddHours_(visitAt, -48), '', now, {}), summary);
    }
  }

  if (visitAt) {
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.REVIEW, 'SALES_REP', swDateAddHours_(visitAt, -24), '', now, {}), summary);
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.CHECKLIST, 'SALES_REP', swDayOfDue_(visitAt), '', now, {}), summary);
  }

  var checklistId = swTaskId_(rec, SW_TASKS.CHECKLIST);
  var processId = swTaskId_(rec, SW_TASKS.PROCESS);
  var approveId = swTaskId_(rec, SW_TASKS.APPROVE);

  if (swTaskCompleted_(state, checklistId)) {
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.PROCESS, 'JOC', dueNow, checklistId, now, {}), summary);
  }

  if (swTaskCompleted_(state, processId)) {
    var processPayload = swTaskPayload_(state, processId);
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.APPROVE, 'SALES_REP', dueNow, processId, now, {
      recapDraft: swDeepValue_(processPayload, ['completion', 'recapText']) || ''
    }), summary);
  }

  if (swTaskCompleted_(state, approveId)) {
    var approvePayload = swTaskPayload_(state, approveId);
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.FINAL, 'JOC', dueNow, approveId, now, {
      approvedText: swDeepValue_(approvePayload, ['completion', 'approvedText']) ||
        swDeepValue_(approvePayload, ['completion', 'recapText']) || ''
    }), summary);
  }
}

function swBuildTask_(ss, state, ctx, rec, taskType, ownerRole, dueAt, dependencyTaskId, now, extraPayload) {
  var template = ctx.templates[taskType] || swDefaultTemplate_(taskType);
  var existing = state.byId[swTaskId_(rec, taskType)] || null;
  var owner = swResolveOwner_(ss, ctx, rec, ownerRole, dueAt || now, existing);
  var payload = {
    appointment: {
      row: rec.row,
      root: rec.root,
      appt: rec.appt,
      uid: rec.uid,
      customerName: rec.name,
      email: rec.email,
      phone: rec.phone,
      brand: rec.brand,
      visitDate: rec.visitDate,
      visitTime: rec.visitTime,
      visitType: rec.visitType,
      assignedRep: rec.assignedRep,
      assignedRepEmail: rec.assignedRepEmail,
      assistedRep: rec.assistedRep,
      assistedRepEmail: rec.assistedRepEmail,
      clientFolder: rec.clientFolder,
      reportUrl: rec.reportUrl
    },
    extra: extraPayload || {}
  };
  payload.extra.mapLink = payload.extra.mapLink || swMapLinkForBrand_(ctx.config, rec.brand);
  payload.extra.locationMsg = payload.extra.locationMsg || swLocationMsgForBrand_(ctx.config, rec.brand);
  if (taskType === SW_TASKS.WELCOME || taskType === SW_TASKS.HYBRID) {
    payload.extra.welcomeMessage = payload.extra.welcomeMessage || swWelcomeMessageForBrand_(ctx.config, rec.brand);
    payload.extra.welcomeImageUrl = payload.extra.welcomeImageUrl || swWelcomeImageForBrand_(ctx.config, rec.brand);
  }

  return {
    taskId: swTaskId_(rec, taskType),
    root: rec.root,
    appt: rec.appt,
    customerName: rec.name,
    brand: rec.brand,
    visitDate: rec.visitDate,
    visitTime: rec.visitTime,
    visitType: rec.visitType,
    lifecycleStage: swLifecycleForTask_(taskType),
    taskType: taskType,
    taskTitle: template.taskTitle,
    ownerRole: ownerRole,
    intendedOwner: owner.intendedOwner,
    intendedOwnerEmail: owner.intendedOwnerEmail,
    currentOwner: owner.currentOwner,
    currentOwnerEmail: owner.currentOwnerEmail,
    coverageReason: owner.coverageReason,
    dueAt: dueAt ? swIso_(dueAt) : '',
    status: SW_STATUSES.PENDING,
    dependencyTaskId: dependencyTaskId || '',
    createdAt: existing ? existing.createdAt : swIso_(now),
    updatedAt: swIso_(now),
    completedBy: existing ? existing.completedBy : '',
    completedByEmail: existing ? existing.completedByEmail : '',
    completedAt: existing ? existing.completedAt : '',
    claimedBy: existing ? existing.claimedBy : '',
    claimedAt: existing ? existing.claimedAt : '',
    lastEvent: existing ? existing.lastEvent : 'CREATE',
    payloadJson: swStringify_(payload),
    templateKey: taskType,
    instructions: template.instructions,
    primaryAction: template.primaryAction,
    rowNumber: existing ? existing.rowNumber : 0
  };
}

function swUpsertTask_(ss, state, nextTask, summary) {
  var existing = state.byId[nextTask.taskId];
  if (!existing) {
    swAppendTaskRow_(ss, nextTask);
    state.byId[nextTask.taskId] = swGetTaskById_(ss, nextTask.taskId);
    swAppendTaskLog_(ss, nextTask.status === SW_STATUSES.COMPLETED ? 'AUTO_COMPLETE' : 'CREATE', nextTask, swSystemUser_(), '', nextTask.currentOwner, {});
    summary.created++;
    return;
  }

  if (existing.status === SW_STATUSES.COMPLETED || existing.claimedBy) return;

  var changed = false;
  [
    'customerName', 'brand', 'visitDate', 'visitTime', 'visitType', 'taskTitle',
    'ownerRole', 'intendedOwner', 'intendedOwnerEmail', 'coverageReason',
    'dueAt', 'dependencyTaskId', 'payloadJson', 'instructions', 'primaryAction'
  ].forEach(function (k) {
    if (String(existing[k] || '') !== String(nextTask[k] || '')) {
      existing[k] = nextTask[k] || '';
      changed = true;
    }
  });

  if (String(existing.currentOwner || '') !== String(nextTask.currentOwner || '') ||
      String(existing.currentOwnerEmail || '') !== String(nextTask.currentOwnerEmail || '')) {
    var fromOwner = existing.currentOwner;
    existing.currentOwner = nextTask.currentOwner || '';
    existing.currentOwnerEmail = nextTask.currentOwnerEmail || '';
    existing.lastEvent = 'ASSIGN';
    swAppendTaskLog_(ss, 'ASSIGN', existing, swSystemUser_(), fromOwner, existing.currentOwner, {
      coverageReason: existing.coverageReason || ''
    });
    changed = true;
  }

  if (changed) {
    existing.updatedAt = swIso_(new Date());
    swWriteTaskRow_(ss, existing);
    summary.updated++;
  }
}

function swBlockTasksForAppointment_(ss, state, rec, reason) {
  var count = 0;
  Object.keys(state.byId).forEach(function (taskId) {
    var t = state.byId[taskId];
    if (t.root !== rec.root && t.appt !== rec.appt) return;
    if (t.status === SW_STATUSES.COMPLETED || t.status === SW_STATUSES.BLOCKED) return;
    t.status = SW_STATUSES.BLOCKED;
    t.coverageReason = reason;
    t.updatedAt = swIso_(new Date());
    t.lastEvent = 'BLOCK';
    swWriteTaskRow_(ss, t);
    swAppendTaskLog_(ss, 'BLOCK', t, swSystemUser_(), t.currentOwner, t.currentOwner, { reason: reason });
    count++;
  });
  return count;
}

function swResolveOwner_(ss, ctx, rec, ownerRole, dueAt, existing) {
  if (existing && (existing.status === SW_STATUSES.COMPLETED || existing.claimedBy)) {
    return {
      intendedOwner: existing.intendedOwner,
      intendedOwnerEmail: existing.intendedOwnerEmail,
      currentOwner: existing.currentOwner,
      currentOwnerEmail: existing.currentOwnerEmail,
      coverageReason: existing.coverageReason
    };
  }

  if (ownerRole === 'System') {
    return {
      intendedOwner: 'System',
      intendedOwnerEmail: '',
      currentOwner: 'System',
      currentOwnerEmail: '',
      coverageReason: ''
    };
  }

  if (ownerRole === 'SALES_REP') {
    var repName = rec.assignedRep || '';
    var repEmail = rec.assignedRepEmail || swLookupEmailByName_(ss, repName) || '';
    return {
      intendedOwner: repName,
      intendedOwnerEmail: repEmail,
      currentOwner: repName || 'Admin Review',
      currentOwnerEmail: repEmail,
      coverageReason: repName ? '' : 'UNASSIGNED_REP'
    };
  }

  if (ownerRole === 'JOC') {
    return swResolveJocOwner_(ss, ctx, rec, dueAt, existing);
  }

  return {
    intendedOwner: '',
    intendedOwnerEmail: '',
    currentOwner: '',
    currentOwnerEmail: '',
    coverageReason: 'UNASSIGNED_OWNER_ROLE'
  };
}

function swResolveJocOwner_(ss, ctx, rec, dueAt, existing) {
  var intendedName = rec.assistedRep || '';
  var intendedEmail = rec.assistedRepEmail || swLookupEmailByName_(ss, intendedName) || '';
  if (!intendedName) {
    return {
      intendedOwner: '',
      intendedOwnerEmail: '',
      currentOwner: 'JOC Coverage',
      currentOwnerEmail: '',
      coverageReason: 'NO_ASSISTED_REP'
    };
  }

  var ownerDate = dueAt && dueAt.getTime && dueAt.getTime() > new Date().getTime() ? dueAt : new Date();
  var intendedAvail = swAvailabilityFor_(ss, intendedName, ownerDate);
  if (intendedAvail.available) {
    return {
      intendedOwner: intendedName,
      intendedOwnerEmail: intendedEmail,
      currentOwner: intendedName,
      currentOwnerEmail: intendedEmail,
      coverageReason: ''
    };
  }

  return {
    intendedOwner: intendedName,
    intendedOwnerEmail: intendedEmail,
    currentOwner: 'JOC Coverage',
    currentOwnerEmail: '',
    coverageReason: swAssistedCoverageReason_(intendedAvail)
  };
}

function swAssistedCoverageReason_(availability) {
  if (!availability || !availability.reason) return 'ASSISTED_REP_UNAVAILABLE';
  if (availability.reason === 'NOT_SCHEDULED') return 'ASSISTED_REP_NOT_SCHEDULED';
  if (availability.reason === 'OUT_OF_OFFICE') return 'ASSISTED_REP_OUT_OF_OFFICE';
  if (availability.reason === 'NO_ROSTER' || availability.reason === 'NO_ROSTER_ROW' || availability.reason === 'ROSTER_SCHEMA_INCOMPLETE') {
    return 'ASSISTED_REP_SCHEDULE_MISSING';
  }
  return 'ASSISTED_REP_UNAVAILABLE';
}

function swAvailabilityFor_(ss, personName, date) {
  personName = swTrim_(personName);
  if (!personName) return { known: false, available: false, reason: 'NO_NAME' };

  var roster = ss.getSheetByName(SW_SHEETS.ROSTER);
  if (!roster || roster.getLastRow() < 2) return { known: false, available: false, reason: 'NO_ROSTER' };

  var values = roster.getDataRange().getValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var repCol = swPickIndex_(H, ['Rep', 'Name', 'Team Member']);
  var day = Utilities.formatDate(date || new Date(), swTimezone_(), 'EEE');
  var dayCol = swPickIndex_(H, [day]);
  if (repCol < 0 || dayCol < 0) return { known: false, available: false, reason: 'ROSTER_SCHEMA_INCOMPLETE' };

  var normTarget = swNorm_(personName);
  var scheduled = null;
  for (var i = 1; i < values.length; i++) {
    var rowName = swTrim_(values[i][repCol]);
    if (!rowName) continue;
    if (swNorm_(rowName) === normTarget) {
      scheduled = swTruthy_(values[i][dayCol]);
      break;
    }
  }
  if (scheduled == null) return { known: false, available: false, reason: 'NO_ROSTER_ROW' };
  if (!scheduled) return { known: true, available: false, reason: 'NOT_SCHEDULED' };

  var override = swScheduleOverride_(ss, personName, date);
  if (override && /off|ooo|out|vacation|pto|sick/i.test(override.changeType || '')) {
    return { known: true, available: false, reason: 'OUT_OF_OFFICE' };
  }

  return { known: true, available: true, reason: 'SCHEDULED' };
}

function swScheduleOverride_(ss, personName, date) {
  var sh = ss.getSheetByName(SW_SHEETS.SCHEDULE_CHANGES);
  if (!sh || sh.getLastRow() < 2) return null;
  var values = sh.getDataRange().getValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var nameCol = swPickIndex_(H, ['Rep Name', 'Rep', 'Name']);
  var dateCol = swPickIndex_(H, ['Change Date', 'Date']);
  var typeCol = swPickIndex_(H, ['Change Type', 'Status', 'Override Status']);
  if (nameCol < 0 || dateCol < 0) return null;

  var targetDate = swDateKey_(date);
  var targetName = swNorm_(personName);
  for (var i = 1; i < values.length; i++) {
    if (swNorm_(values[i][nameCol]) !== targetName) continue;
    if (swDateKey_(values[i][dateCol]) !== targetDate) continue;
    return { changeType: typeCol >= 0 ? swTrim_(values[i][typeCol]) : 'Full-day off' };
  }
  return null;
}

function swBuildContext_(ss) {
  var config = swReadConfig_(ss);
  return {
    tz: swTimezone_(),
    config: config,
    assistedRoster: swReadAssistedRoster_(ss),
    templates: swReadTemplates_(ss),
    admins: swReadAdmins_(ss),
    lookbackDays: Number(swConfigValue_(config, 'SYSTEM', 'WORKFLOW_LOOKBACK_DAYS', '14')) || 14,
    futureDays: Number(swConfigValue_(config, 'SYSTEM', 'WORKFLOW_FUTURE_DAYS', '365')) || 365
  };
}

function swReadAppointments_(ss) {
  var sh = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!sh) throw new Error('Missing sheet: ' + SW_SHEETS.MASTER);
  if (sh.getLastRow() < 2 || sh.getLastColumn() < 1) return [];

  var values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var display = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  var headers = display[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var idx = {
    appt: swPickIndex_(H, ['APPT_ID', 'Appt ID', 'Appointment ID']),
    root: swPickIndex_(H, ['RootApptID', 'Root Appt ID', 'Root Appointment ID']),
    uid: swPickIndex_(H, ['CalendlyEventUID', 'Calendly Event UID', 'Admin: Calendly Event UID', 'Acuity ID', 'UID']),
    name: swPickIndex_(H, ['Customer Name', 'Customer', 'Client Name', 'Name']),
    emailLower: swPickIndex_(H, ['EmailLower', 'Email Lower']),
    email: swPickIndex_(H, ['Email', 'Email Address', 'E-mail']),
    phoneNorm: swPickIndex_(H, ['PhoneNorm', 'Phone Norm']),
    phone: swPickIndex_(H, ['Phone', 'Phone Number', 'Mobile', 'Tel']),
    brand: swPickIndex_(H, ['Brand', 'Company']),
    visitDate: swPickIndex_(H, ['Visit Date', 'Appointment Date', 'Date']),
    visitTime: swPickIndex_(H, ['Visit Time', 'Appointment Time', 'Time']),
    visitType: swPickIndex_(H, ['Visit Type', 'Appointment Type']),
    status: swPickIndex_(H, ['Status']),
    active: swPickIndex_(H, ['Active?', 'Active', 'Is Active']),
    assignedRep: swPickIndex_(H, ['Assigned Rep', 'Rep', 'Owner']),
    assignedRepEmail: swPickIndex_(H, ['Assigned Rep Email', 'Rep Email', 'Owner Email']),
    assistedRep: swPickIndex_(H, ['Assisted Rep', 'Assistant Rep']),
    assistedRepEmail: swPickIndex_(H, ['Assisted Rep Email', 'Assistant Rep Email']),
    clientFolder: swPickIndex_(H, ['Client Folder', 'ClientFolderURL', 'Client Folder URL']),
    reportUrl: swPickIndex_(H, ['Client Status Report URL', 'Report URL'])
  };

  var out = [];
  for (var i = 1; i < values.length; i++) {
    var drow = display[i];
    var vrow = values[i];
    var rec = {
      row: i + 1,
      appt: swTrim_(swCell_(drow, idx.appt)),
      root: swTrim_(swCell_(drow, idx.root)),
      uid: swTrim_(swCell_(drow, idx.uid)),
      name: swTrim_(swCell_(drow, idx.name)),
      email: swNormEmail_(swCell_(drow, idx.emailLower) || swCell_(drow, idx.email)),
      phone: swNormPhone_(swCell_(drow, idx.phoneNorm) || swCell_(drow, idx.phone)),
      brand: swTrim_(swCell_(drow, idx.brand)),
      visitDate: swTrim_(swCell_(drow, idx.visitDate)),
      visitTime: swTrim_(swCell_(drow, idx.visitTime)),
      visitType: swTrim_(swCell_(drow, idx.visitType)),
      visitDateRaw: swCell_(vrow, idx.visitDate),
      visitTimeRaw: swCell_(vrow, idx.visitTime),
      status: swTrim_(swCell_(drow, idx.status)),
      active: swTrim_(swCell_(drow, idx.active)),
      assignedRep: swTrim_(swCell_(drow, idx.assignedRep)),
      assignedRepEmail: swNormEmail_(swCell_(drow, idx.assignedRepEmail)),
      assistedRep: swTrim_(swCell_(drow, idx.assistedRep)),
      assistedRepEmail: swNormEmail_(swCell_(drow, idx.assistedRepEmail)),
      clientFolder: swTrim_(swCell_(drow, idx.clientFolder)),
      reportUrl: swTrim_(swCell_(drow, idx.reportUrl))
    };
    rec.root = rec.root || rec.appt;
    rec.statusNorm = swNorm_(rec.status);
    rec.activeNorm = swNorm_(rec.active);
    out.push(rec);
  }
  return out;
}

function swIsWorkflowRelevant_(rec, now, ctx) {
  var visitAt = swVisitDateTime_(rec, ctx.tz);
  if (!visitAt) return !!(rec.appt || rec.root || rec.name);
  var min = now.getTime() - (ctx.lookbackDays * 24 * 60 * 60 * 1000);
  var max = now.getTime() + (ctx.futureDays * 24 * 60 * 60 * 1000);
  return visitAt.getTime() >= min && visitAt.getTime() <= max;
}

function swIsAppointmentActive_(rec) {
  var s = rec.statusNorm || '';
  var a = rec.activeNorm || '';
  if (/cancel|resched|duplicate|superseded|inactive|no show/.test(s)) return false;
  if (a === 'no' || a === 'n' || a === 'false' || a === '0') return false;
  return true;
}

function swReadTaskState_(ss) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.TASKS, SW_TASK_HEADERS);
  var rows = swReadSheetObjects_(sh);
  var byId = {};
  rows.forEach(function (r) {
    var t = swTaskFromRow_(r);
    if (t.taskId) byId[t.taskId] = t;
  });
  return { rows: rows, byId: byId };
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
  var now = new Date().getTime();
  var tasks = Object.keys(state.byId).map(function (id) { return state.byId[id]; });
  tasks = tasks.filter(function (t) {
    if (view === 'admin') return user.isAdmin && t.status !== SW_STATUSES.COMPLETED;
    if (view === 'coverage') return (user.isJoc || user.isAdmin) && t.ownerRole === 'JOC' &&
      swTaskDueForQueue_(t, now) &&
      t.status === SW_STATUSES.PENDING &&
      (!!t.coverageReason || swNorm_(t.currentOwner) === swNorm_('JOC Coverage'));
    return t.status === SW_STATUSES.PENDING && swTaskDueForQueue_(t, now) && swTaskOwnedByUser_(t, user);
  });
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
  var state = swReadTaskState_(ss);
  return state.byId[taskId] || null;
}

function swAppendTaskRow_(ss, task) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.TASKS, SW_TASK_HEADERS);
  var row = swTaskToRow_(task);
  sh.appendRow(row);
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
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(taskId)) return i + 2;
  }
  return 0;
}

function swAppendTaskLog_(ss, eventType, task, actor, fromOwner, toOwner, details) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.LOG, SW_LOG_HEADERS);
  actor = actor || swSystemUser_();
  sh.appendRow([
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
  ]);
}

function swCurrentUser_(ss) {
  var email = '';
  try { email = swNormEmail_(Session.getActiveUser().getEmail()); } catch (_) {}
  var config = swReadConfig_(ss);
  var assistedRoster = swReadAssistedRoster_(ss);
  var admins = swReadAdmins_(ss);
  var name = email ? swLookupNameByEmail_(ss, email) : '';
  if (!name && email) name = email;
  var isAdmin = admins.length === 0 || admins.indexOf(email) >= 0 || swUserHasConfigRole_(config, email, 'Admin');
  var isJoc = swUserMatchesRoster_(assistedRoster, name, email) || swUserHasConfigRole_(config, email, 'JOC');
  return {
    email: email,
    name: name,
    isAdmin: isAdmin,
    isJoc: isJoc,
    isRep: !!name
  };
}

function swSystemUser_() {
  return { name: 'System', email: '', isAdmin: true, isJoc: false };
}

function swTaskOwnedByUser_(task, user) {
  var email = swNormEmail_(user.email);
  if (email && swNormEmail_(task.currentOwnerEmail) === email) return true;
  if (swNorm_(user.name) && swNorm_(task.currentOwner) === swNorm_(user.name)) return true;
  return false;
}

function swCanViewTask_(task, user) {
  if (user.isAdmin) return true;
  if (swTaskOwnedByUser_(task, user)) return true;
  if (swCanClaimTask_(task, user)) return true;
  return false;
}

function swCanActOnTask_(task, user) {
  if (user.isAdmin) return true;
  return swTaskOwnedByUser_(task, user);
}

function swCanClaimTask_(task, user) {
  if (task.status !== SW_STATUSES.PENDING) return false;
  if (!(user.isJoc || user.isAdmin)) return false;
  if (task.ownerRole !== 'JOC') return false;
  if (swTaskOwnedByUser_(task, user)) return false;
  return !!task.coverageReason || swNorm_(task.currentOwner) === swNorm_('JOC Coverage');
}

function swValidateCompletion_(ss, task, data) {
  var template = swTemplateForType_(ss, task.taskType);
  var payload = swParseJson_(task.payloadJson, {});
  var renderData = swRenderDataForTask_(task, payload);
  var missingTemplateFields = swMissingFieldsForTask_(task, template, renderData);
  if (missingTemplateFields.length) {
    throw new Error('Missing template fields before completion: ' + missingTemplateFields.join(', '));
  }

  var checklist = swParseJson_(template.checklistJson, []);
  if (checklist && checklist.length) {
    var checked = data.checklist || {};
    var missing = [];
    checklist.forEach(function (item) {
      if (item.required !== false && !checked[item.id]) missing.push(item.label || item.id);
    });
    if (missing.length) throw new Error('Complete required checklist items: ' + missing.join(', '));
  }
  if (task.taskType === SW_TASKS.PROCESS && !swTrim_(data.recapText)) {
    throw new Error('Enter the recap draft before completing this task.');
  }
  if (task.taskType === SW_TASKS.APPROVE && !swTrim_(data.approvedText)) {
    throw new Error('Enter the approved recap text before finalizing.');
  }
}

function swReadConfig_(ss) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.CONFIG, SW_CONFIG_HEADERS);
  return swReadSheetObjects_(sh);
}

function swReadAdmins_(ss) {
  var config = swReadConfig_(ss);
  var emails = [];
  config.forEach(function (r) {
    if (swNorm_(r['Section']) === 'system' && swNorm_(r['Key']) === 'adminemails') {
      String(r['Value'] || '').split(/[,\n;]/).forEach(function (e) {
        e = swNormEmail_(e);
        if (e) emails.push(e);
      });
    }
    if (swNorm_(r['Role']) === 'admin' && swTruthy_(r['Active?'] || 'Y')) {
      var em = swNormEmail_(r['Email']);
      if (em) emails.push(em);
    }
  });
  return swUnique_(emails);
}

function swReadAssistedRoster_(ss) {
  var sh = ss.getSheetByName(SW_SHEETS.DROPDOWN);
  var out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var values = sh.getDataRange().getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var nameCol = swPickIndex_(H, ['Assisted Rep', 'Assistant Rep']);
  var emailCol = swPickIndex_(H, ['Assisted Rep Email', 'Assistant Rep Email']);
  if (nameCol < 0) return out;
  var seen = {};
  for (var i = 1; i < values.length; i++) {
    var name = swTrim_(values[i][nameCol]);
    var email = emailCol >= 0 ? swNormEmail_(values[i][emailCol]) : '';
    var key = swNorm_(name) + '|' + email;
    if (!name || seen[key]) continue;
    seen[key] = true;
    out.push({ name: name, email: email });
  }
  return out;
}

function swReadTemplates_(ss) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.TEMPLATES, SW_TEMPLATE_HEADERS);
  var rows = swReadSheetObjects_(sh);
  var out = {};
  rows.forEach(function (r) {
    var type = swTrim_(r['Task Type']);
    if (!type) return;
    out[type] = {
      taskTitle: r['Task Title'] || type,
      instructions: r['Instructions'] || '',
      template: r['Template'] || '',
      attachmentLabel: r['Attachment Label'] || '',
      attachmentUrl: r['Attachment URL'] || '',
      checklistJson: r['Checklist JSON'] || '',
      primaryAction: r['Primary Action'] || 'Complete'
    };
  });
  return out;
}

function swTemplateForType_(ss, taskType) {
  return swReadTemplates_(ss)[taskType] || swDefaultTemplate_(taskType);
}

function swDefaultTemplate_(taskType) {
  var all = swDefaultTemplates_();
  for (var i = 0; i < all.length; i++) {
    if (all[i][0] === taskType) {
      return {
        taskTitle: all[i][1],
        instructions: all[i][2],
        template: all[i][3],
        attachmentLabel: all[i][4],
        attachmentUrl: all[i][5],
        checklistJson: all[i][6],
        primaryAction: all[i][7]
      };
    }
  }
  return { taskTitle: taskType, instructions: '', template: '', checklistJson: '', primaryAction: 'Complete' };
}

function swSeedConfig_(sh) {
  swMigrateConfigRows_(sh);
  var rows = [
    ['SYSTEM', 'FEATURE_ENABLED', 'Y', '', '', '', 'Y', '', 'Set N to pause workflow generation.'],
    ['SYSTEM', 'ADMIN_EMAILS', '', '', '', '', 'Y', '', 'Comma-separated manager/admin emails. Blank means all users can administer during setup.'],
    ['SYSTEM', 'MAP_LINK_VVS', '', '', '', '', 'Y', '', 'Map or instructions link for VVS appointments.'],
    ['SYSTEM', 'MAP_LINK_HUNG_PHAT', '', '', '', '', 'Y', '', 'Map or instructions link for Hung Phat / HPUSA appointments.'],
    ['SYSTEM', 'LOCATION_MSG_VVS', '', '', '', '', 'Y', '', 'Store/location message for VVS map/instructions templates.'],
    ['SYSTEM', 'LOCATION_MSG_HUNG_PHAT', '', '', '', '', 'Y', '', 'Store/location message for Hung Phat / HPUSA map/instructions templates.'],
    ['SYSTEM', 'WELCOME_MSG_VVS', '', '', '', '', 'Y', '', 'Welcome to Your Ring Journey text message for VVS appointments.'],
    ['SYSTEM', 'WELCOME_MSG_HUNG_PHAT', '', '', '', '', 'Y', '', 'Welcome to Your Ring Journey text message for Hung Phat / HPUSA appointments.'],
    ['SYSTEM', 'WELCOME_IMAGE_VVS', '', '', '', '', 'Y', '', 'Welcome to Your Ring Journey image URL for VVS appointments.'],
    ['SYSTEM', 'WELCOME_IMAGE_HUNG_PHAT', '', '', '', '', 'Y', '', 'Welcome to Your Ring Journey image URL for Hung Phat / HPUSA appointments.'],
    ['SYSTEM', 'WORKFLOW_LOOKBACK_DAYS', '14', '', '', '', 'Y', '', 'Do not generate new tasks for appointments older than this.'],
    ['SYSTEM', 'WORKFLOW_FUTURE_DAYS', '365', '', '', '', 'Y', '', 'Generate upcoming appointment workflow tasks through this many days out.'],
    ['SYSTEM', 'JOC_OWNER_SOURCE', '00_Master Appointments: Assisted Rep', '', '', '', 'Y', '', 'JOC ownership comes from each appointment row, not primary/backup config.'],
    ['USER', 'ADMIN_1', '', 'Admin', '', '', 'Y', '1', 'Optional admin row.'],
    ['SYSTEM', 'SHARED_JOC_QUEUE', 'JOC Coverage', '', '', '', 'Y', '', 'Used when no scheduled JOC is available.']
  ];
  swAppendMissingConfigRows_(sh, rows);
}

function swMigrateConfigRows_(sh) {
  if (sh.getLastRow() < 2) return;
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, SW_CONFIG_HEADERS.length).getDisplayValues();
  var byKey = {};
  values.forEach(function (row, i) {
    byKey[swNorm_(row[0]) + '|' + swNorm_(row[1])] = i + 2;
  });

  var renames = {
    locationlabelvvs: 'LOCATION_MSG_VVS',
    locationlabelhungphat: 'LOCATION_MSG_HUNG_PHAT'
  };

  for (var r = values.length - 1; r >= 0; r--) {
    var rowIndex = r + 2;
    var section = swNorm_(values[r][0]);
    var key = swHeaderKey_(values[r][1]);
    if (section !== 'system') continue;
    if (key === 'maplink' || key === 'locationlabel') {
      sh.deleteRow(rowIndex);
      continue;
    }
    if (renames[key]) {
      var targetKey = swNorm_('SYSTEM') + '|' + swNorm_(renames[key]);
      var targetRow = byKey[targetKey];
      if (targetRow && targetRow !== rowIndex) {
        if (!swTrim_(sh.getRange(targetRow, 3).getDisplayValue()) && swTrim_(values[r][2])) {
          sh.getRange(targetRow, 3).setValue(values[r][2]);
        }
        sh.deleteRow(rowIndex);
      } else {
        sh.getRange(rowIndex, 2).setValue(renames[key]);
      }
    }
  }
}

function swAppendMissingConfigRows_(sh, rows) {
  var existing = {};
  if (sh.getLastRow() > 1) {
    var values = sh.getRange(2, 1, sh.getLastRow() - 1, SW_CONFIG_HEADERS.length).getDisplayValues();
    values.forEach(function (row) {
      var key = swNorm_(row[0]) + '|' + swNorm_(row[1]);
      if (key !== '|') existing[key] = true;
    });
  }
  var append = rows.filter(function (row) {
    return !existing[swNorm_(row[0]) + '|' + swNorm_(row[1])];
  });
  if (append.length) {
    sh.getRange(sh.getLastRow() + 1, 1, append.length, SW_CONFIG_HEADERS.length).setValues(append);
  }
}

function swSeedTemplates_(sh) {
  swMigrateTemplateRows_(sh);
  if (sh.getLastRow() > 1) return;
  var rows = swDefaultTemplates_();
  sh.getRange(2, 1, rows.length, SW_TEMPLATE_HEADERS.length).setValues(rows);
}

function swMigrateTemplateRows_(sh) {
  if (sh.getLastRow() < 2) return;
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, SW_TEMPLATE_HEADERS.length).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rowIndex = i + 2;
    var taskType = swTrim_(row[0]);
    if (taskType === SW_TASKS.WELCOME) {
      if (String(row[3] || '').indexOf('welcomeMessage') < 0) {
        sh.getRange(rowIndex, 4).setValue('{{welcomeMessage}}');
      }
      if (String(row[5] || '').indexOf('welcomeImageUrl') < 0) {
        sh.getRange(rowIndex, 5).setValue('Welcome Journey Image');
        sh.getRange(rowIndex, 6).setValue('{{welcomeImageUrl}}');
      }
    }
    if (taskType === SW_TASKS.MAP && String(row[3] || '').indexOf('locationMsg') < 0) {
      sh.getRange(rowIndex, 4).setValue('{{locationMsg}}\n{{mapLink}}');
    }
    if (taskType === SW_TASKS.HYBRID && String(row[3] || '').indexOf('locationMsg') < 0) {
      sh.getRange(rowIndex, 4).setValue(String(row[3] || '') + '\n\n{{locationMsg}}\n{{mapLink}}');
    }
  }
}

function swDefaultTemplates_() {
  return [
    [SW_TASKS.ASSIGN, 'Assign Appointment', 'System-owned assignment record. No manual action needed.', '', '', '', '', 'Assigned'],
    [SW_TASKS.WELCOME, 'Send Welcome to Your Ring Journey Text', 'Send the brand-specific welcome message and welcome image, then mark it sent.', '{{welcomeMessage}}', 'Welcome Journey Image', '{{welcomeImageUrl}}', '', 'Mark Sent'],
    [SW_TASKS.HYBRID, 'Send Hybrid Welcome + Instructions', 'Appointment is within 24 hours. Send the combined welcome and instructions.', 'Hi {{customerName}}, we are looking forward to seeing you {{appointmentDate}} at {{appointmentTime}}. Please review the map/instructions before you arrive. Your stylist is {{assignedRep}}.\n\n{{locationMsg}}\n{{mapLink}}', 'Map / Instructions', '{{mapLink}}', '', 'Mark Sent'],
    [SW_TASKS.MAP, 'Send Map & Instructions', 'Send the map and appointment instructions.', '{{locationMsg}}\n{{mapLink}}', 'Map / Instructions', '{{mapLink}}', '', 'Mark Sent'],
    [SW_TASKS.REVIEW, 'Review Appointment Folder', 'Review the intake form, inspiration images, and customer folder before the appointment.', '', 'Client Folder', '{{clientFolder}}', '', 'Acknowledged & Reviewed'],
    [SW_TASKS.CHECKLIST, 'Appointment Day Checklist', 'Complete each appointment-day item before marking complete.', '', '', '', '[{"id":"printed_intake","label":"Printed intake form","required":true},{"id":"recorded_appointment","label":"Recorded appointment","required":true},{"id":"uploaded_recap","label":"Uploaded recap","required":true},{"id":"uploaded_photos","label":"Uploaded intake photos","required":true},{"id":"goody_bag","label":"Gave goody bag","required":true}]', 'Complete Checklist'],
    [SW_TASKS.PROCESS, 'Process Appointment Data', 'Upload the recording, generate the recap draft, and submit it here.', '', 'Client Folder', '{{clientFolder}}', '', 'Submit Recap Draft'],
    [SW_TASKS.APPROVE, 'Approve/Edit Recap Message', 'Review the JOC recap draft. Edit if needed, then finalize.', '{{recapDraft}}', '', '', '', 'Finalized'],
    [SW_TASKS.FINAL, 'Send Final Recap Text', 'Send the finalized recap message, then mark it sent.', '{{approvedText}}', '', '', '', 'Mark Sent']
  ];
}

function swRenderDataForTask_(task, payload) {
  payload = payload || {};
  var appt = payload.appointment || {};
  var extra = payload.extra || {};
  var completion = payload.completion || {};
  return {
    customerName: task.customerName || appt.customerName || '',
    brand: task.brand || appt.brand || '',
    appointmentDate: task.visitDate || appt.visitDate || '',
    appointmentTime: task.visitTime || appt.visitTime || '',
    visitType: task.visitType || appt.visitType || '',
    assignedRep: appt.assignedRep || '',
    assignedRepEmail: appt.assignedRepEmail || '',
    assistedRep: appt.assistedRep || '',
    assistedRepEmail: appt.assistedRepEmail || '',
    clientFolder: appt.clientFolder || '',
    reportUrl: appt.reportUrl || '',
    mapLink: extra.mapLink || '',
    locationMsg: extra.locationMsg || '',
    welcomeMessage: extra.welcomeMessage || '',
    welcomeImageUrl: extra.welcomeImageUrl || '',
    recapDraft: extra.recapDraft || completion.recapText || '',
    approvedText: extra.approvedText || completion.approvedText || extra.recapDraft || ''
  };
}

function swAttachmentsForTask_(task, template, data) {
  var out = [];
  var primaryUrl = template.attachmentUrl ? swRenderTemplate_(template.attachmentUrl, data) : '';
  var primaryLabel = template.attachmentLabel ? swRenderTemplate_(template.attachmentLabel, data) : '';
  swPushAttachment_(out, primaryLabel, primaryUrl);

  if (task.taskType === SW_TASKS.WELCOME || task.taskType === SW_TASKS.HYBRID) {
    swPushAttachment_(out, 'Welcome Journey Image', data.welcomeImageUrl || '');
  }
  return out;
}

function swPushAttachment_(out, label, url) {
  url = swTrim_(url);
  if (!url) return;
  for (var i = 0; i < out.length; i++) {
    if (swTrim_(out[i].url) === url) return;
  }
  out.push({ label: swTrim_(label) || url, url: url });
}

function swRenderTemplate_(template, data) {
  return String(template || '').replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g, function (_, key) {
    return data[key] == null ? '' : String(data[key]);
  });
}

function swMissingFieldsForTask_(task, template, data) {
  var text = [template.template || '', template.attachmentUrl || ''].join('\n');
  if (task.taskType === SW_TASKS.WELCOME) {
    text += '\n{{welcomeMessage}}\n{{welcomeImageUrl}}';
  }
  if (task.taskType === SW_TASKS.HYBRID) {
    text += '\n{{mapLink}}\n{{locationMsg}}\n{{welcomeImageUrl}}';
  }
  return swMissingTemplateFields_(text, data);
}

function swMissingTemplateFields_(template, data) {
  var missing = {};
  String(template || '').replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g, function (_, key) {
    if (!data[key]) missing[key] = true;
    return '';
  });
  return Object.keys(missing).sort();
}

function swTaskId_(rec, taskType) {
  return ['SW', rec.root || rec.appt || 'NO_ROOT', rec.appt || 'NO_APPT', taskType].join('|');
}

function swTaskCompleted_(state, taskId) {
  return !!(state.byId[taskId] && state.byId[taskId].status === SW_STATUSES.COMPLETED);
}

function swTaskPayload_(state, taskId) {
  return state.byId[taskId] ? swParseJson_(state.byId[taskId].payloadJson, {}) : {};
}

function swLifecycleForTask_(taskType) {
  var map = {};
  map[SW_TASKS.ASSIGN] = 'Booked';
  map[SW_TASKS.WELCOME] = 'Booked';
  map[SW_TASKS.HYBRID] = 'Booked';
  map[SW_TASKS.MAP] = 'Pre-Appointment';
  map[SW_TASKS.REVIEW] = 'Pre-Appointment';
  map[SW_TASKS.CHECKLIST] = 'Appointment Day';
  map[SW_TASKS.PROCESS] = 'Post-Appointment';
  map[SW_TASKS.APPROVE] = 'Post-Appointment';
  map[SW_TASKS.FINAL] = 'Final Follow-Up';
  return map[taskType] || '';
}

function swVisitDateTime_(rec, tz) {
  var dateParts = swDateParts_(rec.visitDateRaw, rec.visitDate);
  if (!dateParts) return null;
  var timeParts = swTimeParts_(rec.visitTimeRaw, rec.visitTime);
  return new Date(dateParts.y, dateParts.m - 1, dateParts.d, timeParts.h, timeParts.min, 0, 0);
}

function swDateParts_(raw, display) {
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return { y: raw.getFullYear(), m: raw.getMonth() + 1, d: raw.getDate() };
  }
  var s = swTrim_(display || raw);
  if (!s) return null;
  var iso = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
  if (iso) return { y: Number(iso[1]), m: Number(iso[2]), d: Number(iso[3]) };
  var mdy = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/.exec(s);
  if (mdy) {
    var y = Number(mdy[3]);
    if (y < 100) y += 2000;
    return { y: y, m: Number(mdy[1]), d: Number(mdy[2]) };
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) return { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() };
  return null;
}

function swTimeParts_(raw, display) {
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return { h: raw.getHours(), min: raw.getMinutes() };
  }
  var s = swTrim_(display || raw);
  if (!s) return { h: 9, min: 0 };
  var m12 = /^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)$/i.exec(s);
  if (m12) {
    var h = Number(m12[1]);
    var ap = m12[3].toUpperCase();
    if (ap === 'AM' && h === 12) h = 0;
    if (ap === 'PM' && h !== 12) h += 12;
    return { h: h, min: Number(m12[2]) };
  }
  var m24 = /^(\d{1,2}):(\d{2})/.exec(s);
  if (m24) return { h: Number(m24[1]), min: Number(m24[2]) };
  return { h: 9, min: 0 };
}

function swDayOfDue_(visitAt) {
  return new Date(visitAt.getFullYear(), visitAt.getMonth(), visitAt.getDate(), 8, 0, 0, 0);
}

function swDateAddHours_(date, hours) {
  return new Date(date.getTime() + hours * 60 * 60 * 1000);
}

function swDateKey_(date) {
  if (!(date instanceof Date)) date = new Date(date);
  if (isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, swTimezone_(), 'yyyy-MM-dd');
}

function swDateValue_(iso) {
  if (!iso) return 9999999999999;
  var d = new Date(iso);
  return isNaN(d.getTime()) ? 9999999999999 : d.getTime();
}

function swIsOverdue_(task, nowMs) {
  return task.status === SW_STATUSES.PENDING && swDateValue_(task.dueAt) < nowMs;
}

function swTaskDueForQueue_(task, nowMs) {
  if (!task.dueAt) return true;
  return swDateValue_(task.dueAt) <= nowMs;
}

function swDueLabel_(task, nowMs) {
  var t = swDateValue_(task.dueAt);
  if (t === 9999999999999) return 'No due time';
  var diff = t - nowMs;
  var mins = Math.round(Math.abs(diff) / 60000);
  if (diff < 0) {
    if (mins < 60) return 'Overdue ' + mins + 'm';
    return 'Overdue ' + Math.round(mins / 60) + 'h';
  }
  if (mins < 60) return 'Due in ' + mins + 'm';
  if (mins < 1440) return 'Due in ' + Math.round(mins / 60) + 'h';
  return 'Due in ' + Math.round(mins / 1440) + 'd';
}

function swSpreadsheet_() {
  var id = '';
  try {
    id = swTrim_(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || PropertiesService.getScriptProperties().getProperty('MASTER_FILE_ID'));
  } catch (_) {}
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (_) {}
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet and no SPREADSHEET_ID script property.');
  return ss;
}

function swEnsureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    var existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
    var col = existing.length;
    headers.forEach(function (h) {
      if (existing.indexOf(h) < 0) {
        col++;
        sh.getRange(1, col).setValue(h);
      }
    });
  }
  return sh;
}

function swStyleSheet_(sh) {
  try {
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#EFE8DD').setFontColor('#2A2725');
    sh.autoResizeColumns(1, Math.min(sh.getLastColumn(), 12));
  } catch (_) {}
}

function swReadSheetObjects_(sh) {
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < 1) return [];
  var values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var obj = { __rowNumber: i + 1 };
    var blank = true;
    for (var j = 0; j < headers.length; j++) {
      if (!headers[j]) continue;
      obj[headers[j]] = values[i][j];
      if (values[i][j] !== '') blank = false;
    }
    if (!blank) out.push(obj);
  }
  return out;
}

function swHeaderMapFromArray_(headers) {
  var map = {};
  headers.forEach(function (h, i) {
    var raw = swTrim_(h);
    if (!raw) return;
    map[raw] = i;
    map[swHeaderKey_(raw)] = i;
  });
  return map;
}

function swPickIndex_(map, names) {
  for (var i = 0; i < names.length; i++) {
    if (map[names[i]] != null) return map[names[i]];
    var key = swHeaderKey_(names[i]);
    if (map[key] != null) return map[key];
  }
  return -1;
}

function swHeaderKey_(value) {
  return swTrim_(value).toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function swCell_(row, idx) {
  return idx >= 0 ? row[idx] : '';
}

function swTrim_(value) {
  return String(value == null ? '' : value).trim();
}

function swNorm_(value) {
  return swTrim_(value).toLowerCase().replace(/\s+/g, ' ');
}

function swNormEmail_(value) {
  return swTrim_(value).toLowerCase();
}

function swNormPhone_(value) {
  var d = swTrim_(value).replace(/\D+/g, '');
  if (d.length > 10 && d.charAt(0) === '1') d = d.slice(1);
  return d.length >= 7 ? d : '';
}

function swTruthy_(value) {
  var s = swNorm_(value);
  return s === 'y' || s === 'yes' || s === 'true' || s === '1' || s === 'working' || s === 'available';
}

function swUnique_(values) {
  var seen = {};
  var out = [];
  values.forEach(function (v) {
    v = swTrim_(v);
    if (!v || seen[v]) return;
    seen[v] = true;
    out.push(v);
  });
  return out;
}

function swStringify_(obj) {
  return JSON.stringify(obj || {});
}

function swParseJson_(text, fallback) {
  if (!text) return fallback;
  try {
    var parsed = JSON.parse(text);
    return parsed == null ? fallback : parsed;
  } catch (_) {
    return fallback;
  }
}

function swDeepValue_(obj, path) {
  var cur = obj;
  for (var i = 0; i < path.length; i++) {
    if (!cur || cur[path[i]] == null) return '';
    cur = cur[path[i]];
  }
  return cur;
}

function swIso_(date) {
  return (date || new Date()).toISOString();
}

function swTimezone_() {
  try { return Session.getScriptTimeZone() || 'America/Los_Angeles'; } catch (_) {}
  return 'America/Los_Angeles';
}

function swConfigValue_(configRows, section, key, fallback) {
  var targetSection = swNorm_(section);
  var targetKey = swNorm_(key);
  for (var i = 0; i < configRows.length; i++) {
    if (swNorm_(configRows[i]['Section']) === targetSection && swNorm_(configRows[i]['Key']) === targetKey) {
      return configRows[i]['Value'] || fallback || '';
    }
  }
  return fallback || '';
}

function swMapLinkForBrand_(configRows, brand) {
  var key = swBrandConfigKey_(brand);
  return key ? swConfigValue_(configRows, 'SYSTEM', 'MAP_LINK_' + key, '') : '';
}

function swLocationMsgForBrand_(configRows, brand) {
  var key = swBrandConfigKey_(brand);
  return key ? swConfigValue_(configRows, 'SYSTEM', 'LOCATION_MSG_' + key, '') : '';
}

function swWelcomeMessageForBrand_(configRows, brand) {
  var key = swBrandConfigKey_(brand);
  return key ? swConfigValue_(configRows, 'SYSTEM', 'WELCOME_MSG_' + key, '') : '';
}

function swWelcomeImageForBrand_(configRows, brand) {
  var key = swBrandConfigKey_(brand);
  return key ? swConfigValue_(configRows, 'SYSTEM', 'WELCOME_IMAGE_' + key, '') : '';
}

function swBrandConfigKey_(brand) {
  var b = swNorm_(brand);
  if (!b) return '';
  if (/\bvvs\b/.test(b) || b.indexOf('vvs') >= 0) return 'VVS';
  if (b.indexOf('hung') >= 0 || b.indexOf('phat') >= 0 || b.indexOf('hpusa') >= 0 || b.indexOf('hp usa') >= 0) {
    return 'HUNG_PHAT';
  }
  return '';
}

function swUserHasConfigRole_(config, email, role) {
  email = swNormEmail_(email);
  if (!email) return false;
  role = swNorm_(role);
  for (var i = 0; i < config.length; i++) {
    if (swNorm_(config[i]['Role']) !== role) continue;
    if (!swTruthy_(config[i]['Active?'] || 'Y')) continue;
    if (swNormEmail_(config[i]['Email']) === email) return true;
  }
  return false;
}

function swUserMatchesRoster_(roster, name, email) {
  var n = swNorm_(name);
  var e = swNormEmail_(email);
  for (var i = 0; i < roster.length; i++) {
    if (e && swNormEmail_(roster[i].email) === e) return true;
    if (n && swNorm_(roster[i].name) === n) return true;
  }
  return false;
}

function swLookupEmailByName_(ss, name) {
  name = swNorm_(name);
  if (!name) return '';
  var sh = ss.getSheetByName(SW_SHEETS.DROPDOWN);
  if (!sh || sh.getLastRow() < 2) return '';
  var values = sh.getDataRange().getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var pairs = [
    [swPickIndex_(H, ['Assigned Rep']), swPickIndex_(H, ['Assigned Rep Email'])],
    [swPickIndex_(H, ['Assisted Rep']), swPickIndex_(H, ['Assisted Rep Email'])]
  ];
  for (var i = 1; i < values.length; i++) {
    for (var p = 0; p < pairs.length; p++) {
      var nameCol = pairs[p][0];
      var emailCol = pairs[p][1];
      if (nameCol >= 0 && emailCol >= 0 && swNorm_(values[i][nameCol]) === name) {
        return swNormEmail_(values[i][emailCol]);
      }
    }
  }
  return '';
}

function swLookupNameByEmail_(ss, email) {
  email = swNormEmail_(email);
  if (!email) return '';
  var sh = ss.getSheetByName(SW_SHEETS.DROPDOWN);
  if (sh && sh.getLastRow() >= 2) {
    var values = sh.getDataRange().getDisplayValues();
    var headers = values[0].map(function (h) { return swTrim_(h); });
    var H = swHeaderMapFromArray_(headers);
    var pairs = [
      [swPickIndex_(H, ['Assigned Rep']), swPickIndex_(H, ['Assigned Rep Email'])],
      [swPickIndex_(H, ['Assisted Rep']), swPickIndex_(H, ['Assisted Rep Email'])]
    ];
    for (var i = 1; i < values.length; i++) {
      for (var p = 0; p < pairs.length; p++) {
        var nameCol = pairs[p][0];
        var emailCol = pairs[p][1];
        if (nameCol >= 0 && emailCol >= 0 && swNormEmail_(values[i][emailCol]) === email) {
          return swTrim_(values[i][nameCol]);
        }
      }
    }
  }

  var config = swReadConfig_(ss);
  for (var c = 0; c < config.length; c++) {
    if (swNormEmail_(config[c]['Email']) === email) return swTrim_(config[c]['Name'] || config[c]['Key']);
  }
  return '';
}
