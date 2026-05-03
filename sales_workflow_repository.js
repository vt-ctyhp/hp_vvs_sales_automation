/**
 * Sales workflow repository: sheet reads/writes, setup, task state, config, templates, and identity.
 */

function swReadRosterAvailabilityIndex_(ss) {
  var out = { exists: false, schemaOk: false, byName: {} };
  var roster = ss.getSheetByName(SW_SHEETS.ROSTER);
  if (!roster || roster.getLastRow() < 2) return out;
  out.exists = true;

  var values = roster.getDataRange().getValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var repCol = swPickIndex_(H, ['Rep', 'Name', 'Team Member']);
  var dayCols = {
    Sun: swPickIndex_(H, ['Sun']),
    Mon: swPickIndex_(H, ['Mon']),
    Tue: swPickIndex_(H, ['Tue']),
    Wed: swPickIndex_(H, ['Wed']),
    Thu: swPickIndex_(H, ['Thu']),
    Fri: swPickIndex_(H, ['Fri']),
    Sat: swPickIndex_(H, ['Sat'])
  };
  out.schemaOk = repCol >= 0 && Object.keys(dayCols).some(function (day) { return dayCols[day] >= 0; });
  if (!out.schemaOk) return out;

  for (var i = 1; i < values.length; i++) {
    var rowName = swTrim_(values[i][repCol]);
    if (!rowName) continue;
    var row = { name: rowName, days: {} };
    Object.keys(dayCols).forEach(function (day) {
      row.days[day] = dayCols[day] >= 0 ? swTruthy_(values[i][dayCols[day]]) : null;
    });
    out.byName[swNorm_(rowName)] = row;
  }
  return out;
}

function swReadScheduleChangesIndex_(ss) {
  var out = { byNameDate: {} };
  var sh = ss.getSheetByName(SW_SHEETS.SCHEDULE_CHANGES);
  if (!sh || sh.getLastRow() < 2) return out;

  var values = sh.getDataRange().getValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var nameCol = swPickIndex_(H, ['Rep Name', 'Rep', 'Name']);
  var dateCol = swPickIndex_(H, ['Change Date', 'Date']);
  var typeCol = swPickIndex_(H, ['Change Type', 'Status', 'Override Status']);
  if (nameCol < 0 || dateCol < 0) return out;

  for (var i = 1; i < values.length; i++) {
    var name = swNorm_(values[i][nameCol]);
    var date = swDateKey_(values[i][dateCol]);
    if (!name || !date) continue;
    out.byNameDate[name + '|' + date] = {
      changeType: typeCol >= 0 ? swTrim_(values[i][typeCol]) : 'Full-day off'
    };
  }
  return out;
}

function swBuildContext_(ss, readOnly) {
  var config = swReadConfig_(ss, readOnly);
  var peopleIndex = swReadPeopleIndex_(ss, config);
  return {
    tz: swTimezone_(),
    config: config,
    peopleIndex: peopleIndex,
    assistedRoster: peopleIndex.assistedRoster,
    templates: swReadTemplates_(ss, readOnly),
    admins: swReadAdminsFromConfig_(config),
    rosterIndex: swReadRosterAvailabilityIndex_(ss),
    scheduleChangesIndex: swReadScheduleChangesIndex_(ss),
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

function swReadTaskState_(ss, readOnly) {
  var sh = readOnly
    ? swGetRequiredSheet_(ss, SW_SHEETS.TASKS)
    : swEnsureSheet_(ss, SW_SHEETS.TASKS, SW_TASK_HEADERS);
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
  return swListVisibleTasksFromState_(state, user, view);
}

function swListVisibleTasksFromState_(state, user, view) {
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

function swCurrentUser_(ss, ctx) {
  var email = '';
  try { email = swNormEmail_(Session.getActiveUser().getEmail()); } catch (_) {}
  ctx = ctx || {};
  var config = ctx.config || swReadConfig_(ss);
  var assistedRoster = ctx.assistedRoster || swReadAssistedRoster_(ss);
  var admins = ctx.admins || swReadAdminsFromConfig_(config);
  var peopleIndex = ctx.peopleIndex || swReadPeopleIndex_(ss, config);
  var name = email ? (peopleIndex.nameByEmail[email] || '') : '';
  if (!name && email && !ctx.peopleIndex) name = swLookupNameByEmail_(ss, email);
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

function swReadConfig_(ss, readOnly) {
  var sh = readOnly
    ? swGetRequiredSheet_(ss, SW_SHEETS.CONFIG)
    : swEnsureSheet_(ss, SW_SHEETS.CONFIG, SW_CONFIG_HEADERS);
  return swReadSheetObjects_(sh);
}

function swReadAdmins_(ss) {
  return swReadAdminsFromConfig_(swReadConfig_(ss));
}

function swReadAdminsFromConfig_(config) {
  var emails = [];
  (config || []).forEach(function (r) {
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

function swReadPeopleIndex_(ss, config) {
  var out = {
    nameByEmail: {},
    emailByName: {},
    assistedRoster: []
  };
  var sh = ss.getSheetByName(SW_SHEETS.DROPDOWN);
  if (sh && sh.getLastRow() >= 2) {
    var values = sh.getDataRange().getDisplayValues();
    var headers = values[0].map(function (h) { return swTrim_(h); });
    var H = swHeaderMapFromArray_(headers);
    var pairs = [
      [swPickIndex_(H, ['Assigned Rep']), swPickIndex_(H, ['Assigned Rep Email'])],
      [swPickIndex_(H, ['Assisted Rep']), swPickIndex_(H, ['Assisted Rep Email'])]
    ];
    var assistedNameCol = swPickIndex_(H, ['Assisted Rep', 'Assistant Rep']);
    var assistedEmailCol = swPickIndex_(H, ['Assisted Rep Email', 'Assistant Rep Email']);
    var seenAssisted = {};
    for (var i = 1; i < values.length; i++) {
      pairs.forEach(function (pair) {
        var nameCol = pair[0];
        var emailCol = pair[1];
        if (nameCol < 0 || emailCol < 0) return;
        var name = swTrim_(values[i][nameCol]);
        var email = swNormEmail_(values[i][emailCol]);
        if (name && email) {
          out.emailByName[swNorm_(name)] = out.emailByName[swNorm_(name)] || email;
          out.nameByEmail[email] = out.nameByEmail[email] || name;
        }
      });

      if (assistedNameCol >= 0) {
        var assistedName = swTrim_(values[i][assistedNameCol]);
        var assistedEmail = assistedEmailCol >= 0 ? swNormEmail_(values[i][assistedEmailCol]) : '';
        var assistedKey = swNorm_(assistedName) + '|' + assistedEmail;
        if (assistedName && !seenAssisted[assistedKey]) {
          seenAssisted[assistedKey] = true;
          out.assistedRoster.push({ name: assistedName, email: assistedEmail });
        }
      }
    }
  }

  (config || []).forEach(function (row) {
    var email = swNormEmail_(row['Email']);
    var name = swTrim_(row['Name'] || row['Key']);
    if (email && name) {
      out.nameByEmail[email] = out.nameByEmail[email] || name;
      out.emailByName[swNorm_(name)] = out.emailByName[swNorm_(name)] || email;
    }
  });

  return out;
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

function swReadTemplates_(ss, readOnly) {
  var sh = readOnly
    ? swGetRequiredSheet_(ss, SW_SHEETS.TEMPLATES)
    : swEnsureSheet_(ss, SW_SHEETS.TEMPLATES, SW_TEMPLATE_HEADERS);
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

function swRequireWorkflowReadSheets_(ss) {
  [SW_SHEETS.TASKS, SW_SHEETS.CONFIG, SW_SHEETS.TEMPLATES].forEach(function (name) {
    swGetRequiredSheet_(ss, name);
  });
}

function swGetRequiredSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name + '. Run sw_setupSalesWorkflow first.');
  return sh;
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
