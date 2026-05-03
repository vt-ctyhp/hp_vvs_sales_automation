/**
 * Sales workflow generation: appointment-to-task rules, owner resolution, and dependencies.
 */

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
    swQueueOrAppendTaskRow_(ss, state, nextTask);
    state.byId[nextTask.taskId] = nextTask;
    swQueueOrAppendTaskLog_(ss, state, nextTask.status === SW_STATUSES.COMPLETED ? 'AUTO_COMPLETE' : 'CREATE', nextTask, swSystemUser_(), '', nextTask.currentOwner, {});
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
    var repEmail = rec.assignedRepEmail || swLookupEmailByName_(ss, repName, ctx) || '';
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
  var intendedEmail = rec.assistedRepEmail || swLookupEmailByName_(ss, intendedName, ctx) || '';
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
  var intendedAvail = swAvailabilityFor_(ss, intendedName, ownerDate, ctx);
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

function swAvailabilityFor_(ss, personName, date, ctx) {
  personName = swTrim_(personName);
  if (!personName) return { known: false, available: false, reason: 'NO_NAME' };

  ctx = ctx || {};
  var rosterIndex = ctx.rosterIndex || swReadRosterAvailabilityIndex_(ss);
  if (!rosterIndex.exists) return { known: false, available: false, reason: 'NO_ROSTER' };
  if (!rosterIndex.schemaOk) return { known: false, available: false, reason: 'ROSTER_SCHEMA_INCOMPLETE' };

  var day = Utilities.formatDate(date || new Date(), swTimezone_(), 'EEE');
  var row = rosterIndex.byName[swNorm_(personName)];
  if (!row) return { known: false, available: false, reason: 'NO_ROSTER_ROW' };

  var scheduled = row.days[day];
  if (scheduled == null) return { known: false, available: false, reason: 'ROSTER_SCHEMA_INCOMPLETE' };
  if (!scheduled) return { known: true, available: false, reason: 'NOT_SCHEDULED' };

  var override = swScheduleOverride_(ss, personName, date, ctx);
  if (override && /off|ooo|out|vacation|pto|sick/i.test(override.changeType || '')) {
    return { known: true, available: false, reason: 'OUT_OF_OFFICE' };
  }

  return { known: true, available: true, reason: 'SCHEDULED' };
}

function swScheduleOverride_(ss, personName, date, ctx) {
  var targetDate = swDateKey_(date);
  var targetName = swNorm_(personName);
  ctx = ctx || {};
  var scheduleIndex = ctx.scheduleChangesIndex || swReadScheduleChangesIndex_(ss);
  return scheduleIndex.byNameDate[targetName + '|' + targetDate] || null;
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
