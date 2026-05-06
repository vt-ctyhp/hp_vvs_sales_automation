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
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.REVIEW, SW_OWNER_ROLES.SALES_REP, swDateAddHours_(visitAt, -24), '', now, {}), summary);
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.CHECKLIST, SW_OWNER_ROLES.SALES_REP, swDayOfDue_(visitAt), '', now, {}), summary);
  }

  swGenerateDiamondWorkflowTasks_(ss, state, ctx, rec, now, summary, visitAt);
  swGeneratePostConsultTasks_(ss, state, ctx, rec, now, summary);

  var checklistId = swTaskId_(rec, SW_TASKS.CHECKLIST);
  var approveId = swTaskId_(rec, SW_TASKS.APPROVE);
  var appointmentOutcome = typeof swAppointmentOutcomeForRoot_ === 'function' ? swAppointmentOutcomeForRoot_(state, rec) : '';
  var isNoShow = typeof swIsNoShowOutcome_ === 'function' ? swIsNoShowOutcome_(appointmentOutcome) : false;
  var aiSummary = (ctx.appointmentSummaryByRoot && ctx.appointmentSummaryByRoot[rec.root || rec.appt || '']) ||
    (typeof swSummaryExtraForRoot_ === 'function' ? swSummaryExtraForRoot_(ss, rec.root || rec.appt || '') : { ready: false });

  if (swTaskCompleted_(state, checklistId) && !isNoShow && aiSummary.ready) {
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.APPROVE, SW_OWNER_ROLES.SALES_REP, dueNow, checklistId, now, {
      artifactId: aiSummary.artifactId || '',
      workflowStage: aiSummary.workflowStage || '',
      transcriptDocUrl: aiSummary.transcriptDocUrl || '',
      summaryDocUrl: aiSummary.summaryDocUrl || '',
      summaryJsonUrl: aiSummary.summaryJsonUrl || '',
      salesBrief: aiSummary.salesBrief || '',
      reviewFlags: aiSummary.reviewFlags || '',
      clientFollowUpDraft: aiSummary.clientFollowUpDraft || '',
      recapDraft: aiSummary.recapDraft || ''
    }), summary);
  }

  if (swTaskCompleted_(state, approveId)) {
    var approvePayload = swTaskPayload_(state, approveId);
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.FINAL, 'JOC', dueNow, approveId, now, {
      approvedText: swDeepValue_(approvePayload, ['completion', 'approvedText']) ||
        swDeepValue_(approvePayload, ['completion', 'recapText']) || '',
      artifactId: swDeepValue_(approvePayload, ['extra', 'artifactId']) || '',
      transcriptDocUrl: swDeepValue_(approvePayload, ['extra', 'transcriptDocUrl']) || '',
      summaryDocUrl: swDeepValue_(approvePayload, ['extra', 'summaryDocUrl']) || '',
      summaryJsonUrl: swDeepValue_(approvePayload, ['extra', 'summaryJsonUrl']) || '',
      salesBrief: swDeepValue_(approvePayload, ['extra', 'salesBrief']) || '',
      reviewFlags: swDeepValue_(approvePayload, ['extra', 'reviewFlags']) || ''
    }), summary);
  }
}

function swBuildTask_(ss, state, ctx, rec, taskType, ownerRole, dueAt, dependencyTaskId, now, extraPayload) {
  var template = ctx.templates[taskType] || swDefaultTemplate_(taskType);
  var existing = state.byId[swTaskId_(rec, taskType)] || null;
  var owner = swResolveOwner_(ss, ctx, rec, ownerRole, dueAt || now, existing);
  var visitTime = swFormatAppointmentTime_(rec.visitTime, rec.visitTimeRaw);
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
      visitTime: visitTime,
      visitType: rec.visitType,
      diamondType: rec.diamondType,
      assignedRep: rec.assignedRep,
      assignedRepEmail: rec.assignedRepEmail,
      assistedRep: rec.assistedRep,
      assistedRepEmail: rec.assistedRepEmail,
      clientFolder: rec.clientFolder,
      reportUrl: rec.reportUrl,
      quotationUrl: rec.quotationUrl,
      tracker3dUrl: rec.tracker3dUrl,
      salesStage: rec.salesStage,
      convStatus: rec.convStatus,
      customOrder: rec.customOrder,
      inProduction: rec.inProduction,
      nextSteps: rec.nextSteps,
      designRequest: rec.designRequest,
      deadline3d: rec.deadline3d,
      productionDeadline: rec.productionDeadline,
      waxStatus: rec.waxStatus,
      waxDeadlineAdmin: rec.waxDeadlineAdmin,
      waxRequestUrl: rec.waxRequestUrl,
      centerStoneStatus: rec.centerStoneStatus,
      dvStonesSummary: rec.dvStonesSummary,
      dvCustomerLookingFor: rec.dvCustomerLookingFor,
      dvVarietyStrategy: rec.dvVarietyStrategy,
      dvCustomerRequirementsJson: rec.dvCustomerRequirementsJson,
      so: rec.so,
      orderFolder: rec.orderFolder
    },
    extra: extraPayload || {}
  };
  payload.extra.mapLink = payload.extra.mapLink || swMapLinkForBrand_(ctx.config, rec.brand);
  payload.extra.locationMsg = payload.extra.locationMsg || swLocationMsgForBrand_(ctx.config, rec.brand);
  if (taskType === SW_TASKS.WELCOME || taskType === SW_TASKS.HYBRID) {
    payload.extra.welcomeMessage = payload.extra.welcomeMessage || swWelcomeMessageForBrand_(ctx.config, rec.brand);
    payload.extra.welcomeImageUrl = payload.extra.welcomeImageUrl || swWelcomeImageForBrand_(ctx.config, rec.brand);
  }
  if (taskType === SW_TASKS.HYBRID) {
    payload.extra.hybridMessage = payload.extra.hybridMessage || swHybridMessageForBrand_(ctx.config, rec.brand);
  }

  return {
    taskId: swTaskId_(rec, taskType),
    root: rec.root,
    appt: rec.appt,
    customerName: rec.name,
    brand: rec.brand,
    visitDate: rec.visitDate,
    visitTime: visitTime,
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
    snoozeUntil: existing ? existing.snoozeUntil : '',
    snoozeReason: existing ? existing.snoozeReason : '',
    snoozedBy: existing ? existing.snoozedBy : '',
    snoozedAt: existing ? existing.snoozedAt : '',
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
  if (existing.status === SW_STATUSES.BLOCKED && !swShouldReviveGeneratedTask_(existing)) return;

  var changed = false;
  var revivedFromBlocked = false;
  var previousBlockReason = '';

  if (swShouldReviveGeneratedTask_(existing)) {
    previousBlockReason = existing.coverageReason || '';
    existing.status = SW_STATUSES.PENDING;
    existing.coverageReason = nextTask.coverageReason || '';
    existing.lastEvent = 'UNBLOCK';
    revivedFromBlocked = true;
    changed = true;
  }

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

  if (revivedFromBlocked) {
    swAppendTaskLog_(ss, 'UNBLOCK', existing, swSystemUser_(), existing.currentOwner, existing.currentOwner, {
      reason: 'Appointment is active/current again.',
      previousReason: previousBlockReason || ''
    });
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
    if (!swTaskMatchesAppointmentInstance_(t, rec)) return;
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

function swTaskMatchesAppointmentInstance_(task, rec) {
  var recAppt = swTrim_(rec && rec.appt);
  var taskAppt = swTrim_(task && task.appt);
  if (recAppt && taskAppt) return taskAppt === recAppt;
  if (recAppt) return !taskAppt && swTrim_(task && task.root) === recAppt;

  var recRoot = swTrim_(rec && rec.root);
  if (!recRoot) return false;
  return swTrim_(task && task.root) === recRoot;
}

function swShouldReviveGeneratedTask_(task) {
  if (!task || task.status !== SW_STATUSES.BLOCKED) return false;
  var reason = swTrim_(task.coverageReason);
  if (reason === SW_INACTIVE_APPOINTMENT_BLOCK_REASON) return true;
  // Older refreshes could clear the auto-block reason while reassigning a blocked row.
  return !reason && task.lastEvent === 'ASSIGN';
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

  if (ownerRole === SW_OWNER_ROLES.SYSTEM || ownerRole === 'System') {
    return {
      intendedOwner: 'System',
      intendedOwnerEmail: '',
      currentOwner: 'System',
      currentOwnerEmail: '',
      coverageReason: ''
    };
  }

  if (swWorkflowRoleMatches_(ownerRole, SW_OWNER_ROLES.SALES_REP)) {
    var rep = swCanonicalWorkflowOwnerForRole_(ss, ctx, rec.assignedRep, rec.assignedRepEmail, SW_OWNER_ROLES.SALES_REP);
    var repName = rep ? rep.name : swTrim_(rec.assignedRep || '');
    var repEmail = rep ? rep.email : '';
    if ((rec.assignedRep || rec.assignedRepEmail) && !rep) {
      return {
        intendedOwner: repName || swTrim_(rec.assignedRepEmail || ''),
        intendedOwnerEmail: '',
        currentOwner: 'Admin Review',
        currentOwnerEmail: '',
        coverageReason: 'UNRESOLVED_REP'
      };
    }
    return {
      intendedOwner: repName,
      intendedOwnerEmail: repEmail,
      currentOwner: repName || 'Admin Review',
      currentOwnerEmail: repEmail,
      coverageReason: repName ? '' : 'UNASSIGNED_REP'
    };
  }

  if (ownerRole === SW_OWNER_ROLES.JOC || ownerRole === 'JOC') {
    return swResolveJocOwner_(ss, ctx, rec, dueAt, existing);
  }

  if (ownerRole === SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN || ownerRole === SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT) {
    return swResolveRoleQueueOwner_(ctx, ownerRole);
  }

  return {
    intendedOwner: '',
    intendedOwnerEmail: '',
    currentOwner: '',
    currentOwnerEmail: '',
    coverageReason: 'UNASSIGNED_OWNER_ROLE'
  };
}

function swResolveRoleQueueOwner_(ctx, ownerRole) {
  var label = swRoleQueueLabel_(ctx, ownerRole);
  return {
    intendedOwner: label,
    intendedOwnerEmail: '',
    currentOwner: label,
    currentOwnerEmail: '',
    coverageReason: ''
  };
}

function swResolveConfigRoleOwner_(ctx, ownerRole) {
  ctx = ctx || {};
  var config = ctx.config || [];
  var candidates = config.filter(function (row) {
    return swNorm_(row['Role']) === swNorm_(ownerRole) && swTruthy_(row['Active?'] || 'Y') && swNormEmail_(row['Email']);
  }).sort(function (a, b) {
    return (Number(a['Priority']) || 999) - (Number(b['Priority']) || 999);
  });
  if (candidates.length) {
    return {
      intendedOwner: swTrim_(candidates[0]['Name'] || candidates[0]['Key']),
      intendedOwnerEmail: swNormEmail_(candidates[0]['Email']),
      currentOwner: swTrim_(candidates[0]['Name'] || candidates[0]['Key']) || swNormEmail_(candidates[0]['Email']),
      currentOwnerEmail: swNormEmail_(candidates[0]['Email']),
      coverageReason: ''
    };
  }

  var sharedKey = ownerRole === SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN
    ? 'SHARED_DIAMOND_ORDER_ADMIN_QUEUE'
    : 'SHARED_DIAMOND_ORDER_ASSISTANT_QUEUE';
  var shared = swConfigValue_(config, 'SYSTEM', sharedKey,
    ownerRole === SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN ? 'Diamond Order Admin Coverage' : 'Diamond Order Assistant Coverage');
  return {
    intendedOwner: '',
    intendedOwnerEmail: '',
    currentOwner: shared,
    currentOwnerEmail: '',
    coverageReason: 'UNASSIGNED_' + ownerRole
  };
}

function swRoleQueueLabel_(ctx, ownerRole) {
  if (ownerRole === SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN) {
    return 'Diamond Order Admin';
  }
  if (ownerRole === SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT) {
    return 'Diamond Order Assistant';
  }
  return String(ownerRole || '');
}

function swResolveJocOwner_(ss, ctx, rec, dueAt, existing) {
  var joc = swCanonicalWorkflowOwnerForRole_(ss, ctx, rec.assistedRep, rec.assistedRepEmail, SW_OWNER_ROLES.JOC);
  var intendedName = joc ? joc.name : swTrim_(rec.assistedRep || '');
  var intendedEmail = joc ? joc.email : '';
  if (!intendedName) {
    return {
      intendedOwner: '',
      intendedOwnerEmail: '',
      currentOwner: 'JOC Coverage',
      currentOwnerEmail: '',
      coverageReason: 'NO_ASSISTED_REP'
    };
  }
  if (!joc) {
    return {
      intendedOwner: intendedName,
      intendedOwnerEmail: '',
      currentOwner: 'JOC Coverage',
      currentOwnerEmail: '',
      coverageReason: 'UNRESOLVED_ASSISTED_REP'
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

  var coverage = swFindAvailableJocCoverageOwner_(ss, ctx, intendedName, ownerDate);
  if (coverage) {
    return {
      intendedOwner: intendedName,
      intendedOwnerEmail: intendedEmail,
      currentOwner: coverage.name,
      currentOwnerEmail: coverage.email,
      coverageReason: swAssistedCoverageReason_(intendedAvail) + '_COVERED_BY_JOC'
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
  if (row.active === false) return { known: true, available: false, reason: 'INACTIVE' };

  var override = swScheduleOverride_(ss, personName, date, ctx);
  if (override && /off|ooo|out|vacation|pto|sick/i.test(override.changeType || '')) {
    return { known: true, available: false, reason: 'OUT_OF_OFFICE' };
  }
  if (override && (override.availableFrom || override.availableUntil || swTruthy_(override.changeType || ''))) {
    if (!swScheduleOverrideAllowsTime_(override, date)) {
      return { known: true, available: false, reason: 'PARTIAL_UNAVAILABLE' };
    }
    return { known: true, available: true, reason: 'OVERRIDE_WORKING' };
  }

  var scheduled = row.days[day];
  if (scheduled == null) return { known: false, available: false, reason: 'ROSTER_SCHEMA_INCOMPLETE' };
  if (!scheduled) return { known: true, available: false, reason: 'NOT_SCHEDULED' };

  return { known: true, available: true, reason: 'SCHEDULED' };
}

function swScheduleOverride_(ss, personName, date, ctx) {
  var targetDate = swDateKey_(date);
  var targetName = swNorm_(personName);
  ctx = ctx || {};
  var scheduleIndex = ctx.scheduleChangesIndex || swReadScheduleChangesIndex_(ss);
  var rosterIndex = ctx.rosterIndex || swReadRosterAvailabilityIndex_(ss);
  var rosterRow = rosterIndex.byName ? rosterIndex.byName[targetName] : null;
  if (rosterRow && rosterRow.email && scheduleIndex.byEmailDate[rosterRow.email + '|' + targetDate]) {
    return scheduleIndex.byEmailDate[rosterRow.email + '|' + targetDate];
  }
  return scheduleIndex.byNameDate[targetName + '|' + targetDate] || null;
}

function swScheduleOverrideAllowsTime_(override, date) {
  var mins = swDateMinutes_(date);
  if (mins == null) return true;
  var from = swTimeStringToMinutes_(override && override.availableFrom);
  var until = swTimeStringToMinutes_(override && override.availableUntil);
  if (from != null && mins < from) return false;
  if (until != null && mins >= until) return false;
  return true;
}

function swDateMinutes_(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return null;
  return date.getHours() * 60 + date.getMinutes();
}

function swTimeStringToMinutes_(value) {
  var s = swTrim_(value);
  if (!s) return null;
  var m12 = /^(\d{1,2}):(\d{2})\s*(AM|PM)$/i.exec(s);
  if (m12) {
    var h12 = Number(m12[1]);
    var ap = m12[3].toUpperCase();
    if (ap === 'AM' && h12 === 12) h12 = 0;
    if (ap === 'PM' && h12 !== 12) h12 += 12;
    return h12 * 60 + Number(m12[2]);
  }
  var m24 = /^(\d{1,2}):(\d{2})/.exec(s);
  if (!m24) return null;
  return Number(m24[1]) * 60 + Number(m24[2]);
}

function swFindAvailableJocCoverageOwner_(ss, ctx, intendedName, date) {
  ctx = ctx || {};
  var preferredPartner = '';
  var rosterIndex = ctx.rosterIndex || swReadRosterAvailabilityIndex_(ss);
  var intendedRow = rosterIndex.byName ? rosterIndex.byName[swNorm_(intendedName)] : null;
  if (intendedRow && intendedRow.coverageEnabled !== false) preferredPartner = swTrim_(intendedRow.coveragePartner);

  var candidates = swAvailableScheduledPeopleForRole_(ss, ctx, SW_OWNER_ROLES.JOC, date, intendedName);
  if (!candidates.length) return null;
  candidates.sort(function (a, b) {
    var ap = preferredPartner && swNorm_(a.name) === swNorm_(preferredPartner) ? 0 : 1;
    var bp = preferredPartner && swNorm_(b.name) === swNorm_(preferredPartner) ? 0 : 1;
    if (ap !== bp) return ap - bp;
    return String(a.name || '').localeCompare(String(b.name || ''));
  });
  return candidates[0];
}

function swAvailableScheduledPeopleForRole_(ss, ctx, role, date, excludeName) {
  ctx = ctx || {};
  var people = ctx.employeeSchedulePeople || null;
  if (!people) {
    people = typeof swReadEmployeeSchedulePeople_ === 'function' ? swReadEmployeeSchedulePeople_(ss) : [];
    ctx.employeeSchedulePeople = people;
  }
  var exclude = swNorm_(excludeName);
  var out = [];
  (people || []).forEach(function (person) {
    if (!person || !person.name) return;
    if (exclude && swNorm_(person.name) === exclude) return;
    if (person.active === false) return;
    if (typeof swEmployeeHasRole_ === 'function' && !swEmployeeHasRole_(person, role)) return;
    var available = swAvailabilityFor_(ss, person.name, date, ctx);
    if (!available.available) return;
    out.push({
      name: person.name,
      email: swNormEmail_(person.email || '') || swLookupEmailByName_(ss, person.name, ctx),
      defaultJoc: person.defaultJoc || '',
      skills: person.skills || swDefaultRepSkills_()
    });
  });
  return out;
}

function swCanonicalWorkflowOwnerForRole_(ss, ctx, name, email, role) {
  ctx = ctx || {};
  if (!ctx.workflowOwnerIndexByRole) {
    var index = {};
    swReadCanonicalWorkflowPeople_(ss, { schedulableOnly: true, activeOnly: true }).forEach(function (user) {
      [SW_OWNER_ROLES.SALES_REP, SW_OWNER_ROLES.JOC].forEach(function (candidateRole) {
        if (!swWorkflowUserHasSchedulableRole_(user, candidateRole)) return;
        var key = swWorkflowRoleKey_(candidateRole);
        if (!index[key]) index[key] = { byEmail: {}, byName: {} };
        if (user.email && !index[key].byEmail[user.email]) index[key].byEmail[user.email] = user;
        if (user.name && !index[key].byName[swNorm_(user.name)]) index[key].byName[swNorm_(user.name)] = user;
      });
    });
    ctx.workflowOwnerIndexByRole = index;
  }
  var bucket = ctx.workflowOwnerIndexByRole[swWorkflowRoleKey_(role)] || { byEmail: {}, byName: {} };
  email = swNormEmail_(email || '');
  name = swTrim_(name || '');
  var user = email ? bucket.byEmail[email] : null;
  if (!user && name) user = bucket.byName[swNorm_(name)] || null;
  return user ? { name: user.name, email: user.email, role: user.scheduleRole || user.role || '' } : null;
}

function swPrepareClientAdvisorRoundRobin_(ss, ctx, appointments) {
  ctx = ctx || {};
  var enabled = swTruthy_(swConfigValue_(ctx.config || [], 'SYSTEM', 'CLIENT_ADVISOR_ROUND_ROBIN', 'N'));
  var people = typeof swReadEmployeeSchedulePeople_ === 'function' ? swReadEmployeeSchedulePeople_(ss) : [];
  ctx.employeeSchedulePeople = people;
  var state = {
    enabled: enabled,
    appointments: appointments || [],
    countsByDate: {},
    busyBySlotOwner: {},
    assignedByRoot: {},
    advisors: people.filter(function (person) {
      return person && person.active !== false && swEmployeeHasRole_(person, SW_OWNER_ROLES.SALES_REP);
    }).map(function (person) {
      return {
        name: person.name,
        email: swNormEmail_(person.email || '') || swLookupEmailByName_(ss, person.name, ctx),
        defaultJoc: person.defaultJoc || '',
        skills: person.skills || swDefaultRepSkills_()
      };
    })
  };
  ctx.clientAdvisorRoundRobin = state;
  if (!enabled) return state;

  (appointments || []).forEach(function (rec) {
    if (!rec || (!rec.assignedRep && !rec.assignedRepEmail)) return;
    var visitAt = swVisitDateTime_(rec, ctx.tz);
    if (!visitAt) return;
    if (!swIsAppointmentActive_(rec)) return;
    var existingOwner = swCanonicalWorkflowOwnerForRole_(ss, ctx, rec.assignedRep, rec.assignedRepEmail, SW_OWNER_ROLES.SALES_REP);
    if (!existingOwner) return;
    var dateKey = swDateKey_(visitAt);
    var ownerKey = swNorm_(existingOwner.name);
    if (!state.countsByDate[dateKey]) state.countsByDate[dateKey] = {};
    state.countsByDate[dateKey][ownerKey] = (state.countsByDate[dateKey][ownerKey] || 0) + 1;
    var existingJoc = swCanonicalWorkflowOwnerForRole_(ss, ctx, rec.assistedRep, rec.assistedRepEmail, SW_OWNER_ROLES.JOC);
    if (rec.root) state.assignedByRoot[rec.root] = {
      name: existingOwner.name,
      email: existingOwner.email,
      defaultJoc: existingJoc ? existingJoc.name : ''
    };
    if (swIsAppointmentActive_(rec)) swTrackClientAdvisorBusySlot_(state, existingOwner.name, visitAt, rec);
  });
  return state;
}

function swMaybeAutoAssignClientAdvisor_(ss, ctx, rec, summary) {
  var rr = ctx && ctx.clientAdvisorRoundRobin;
  if (!rr || !rr.enabled || !rec) return false;
  var visitAt = swVisitDateTime_(rec, ctx.tz);
  if (!visitAt) return false;
  var root = rec.root || rec.appt || '';

  var currentOwner = (rec.assignedRep || rec.assignedRepEmail)
    ? swCanonicalClientAdvisorOwner_(ss, ctx, rec.assignedRep, rec.assignedRepEmail)
    : null;
  var preferredPrevious = swPickPreviousClientAdvisorForAppointment_(ss, ctx, rr, visitAt, rec);
  var currentIsPreferredPrevious = preferredPrevious && currentOwner &&
    swNorm_(preferredPrevious.name) === swNorm_(currentOwner.name);
  if (preferredPrevious && !currentIsPreferredPrevious) {
    return swApplyClientAdvisorAssignment_(ss, ctx, rr, rec, visitAt, preferredPrevious, summary, 'autoReassignedClientAdvisors');
  }

  var currentUnavailable = false;
  var currentUnqualified = false;
  var currentBackupShouldYield = false;
  if (rec.assignedRep || rec.assignedRepEmail) {
    var currentAvail = currentOwner ? swAvailabilityFor_(ss, currentOwner.name, visitAt, ctx) : null;
    currentUnavailable = !currentOwner || (currentAvail.known && !currentAvail.available) ||
      (currentOwner && swClientAdvisorHasAppointmentConflict_(rr, currentOwner.name, visitAt, rec));
    var existingAdvisor = currentOwner ? swRoundRobinAdvisorByName_(rr, currentOwner.name) : null;
    currentUnqualified = !existingAdvisor || !swAdvisorQualifiedForAppointment_(existingAdvisor, rec);
    currentBackupShouldYield = !currentIsPreferredPrevious && !currentUnavailable && !currentUnqualified &&
      swClientAdvisorBackupShouldYield_(ss, ctx, rr, existingAdvisor, visitAt, rec);
    if (!currentUnavailable && !currentUnqualified && !currentBackupShouldYield) {
      if (swCurrentClientAdvisorNeedsLinkedJoc_(ss, ctx, existingAdvisor, rec)) {
        return swApplyClientAdvisorAssignment_(ss, ctx, rr, rec, visitAt, existingAdvisor, summary, 'autoLinkedJocFromAdvisor');
      }
      return false;
    }
  }

  var existingRootAssignment = !rec.assignedRep && root ? swExistingRootClientAdvisorAssignment_(ss, ctx, rr, visitAt, rec, root) : null;
  var chosen = existingRootAssignment || swPickClientAdvisorRoundRobin_(ss, ctx, rr, visitAt, rec, rec.assignedRep || '');
  if (!chosen || !chosen.name) {
    summary.clientAdvisorRoundRobinNoOwner = (summary.clientAdvisorRoundRobinNoOwner || 0) + 1;
    return false;
  }

  return swApplyClientAdvisorAssignment_(ss, ctx, rr, rec, visitAt, chosen, summary,
    (currentUnavailable || currentUnqualified || currentBackupShouldYield) ? 'autoReassignedClientAdvisors' : 'autoAssignedClientAdvisors');
}

function swApplyClientAdvisorAssignment_(ss, ctx, rr, rec, visitAt, advisor, summary, summaryKey) {
  if (!advisor || !advisor.name) return false;
  var previousName = swNorm_(rec.assignedRep || '');
  var previousEmail = swNormEmail_(rec.assignedRepEmail || '');
  var nextName = swNorm_(advisor.name || '');
  var nextEmail = swNormEmail_(advisor.email || '');
  var ownerChanged = previousName !== nextName && (!previousEmail || !nextEmail || previousEmail !== nextEmail);
  var linkedJoc = swTrim_(advisor.defaultJoc || '');
  var linkedJocUser = linkedJoc ? swCanonicalWorkflowOwnerForRole_(ss, ctx, linkedJoc, '', SW_OWNER_ROLES.JOC) : null;
  var linkedJocName = linkedJocUser ? linkedJocUser.name : linkedJoc;
  var linkedJocEmail = linkedJocUser ? linkedJocUser.email : '';

  if (!swWriteAppointmentOwnerAssignmentToMaster_(ss, rec, advisor, linkedJocName, linkedJocEmail)) return false;
  rec.assignedRep = advisor.name;
  rec.assignedRepEmail = advisor.email || '';
  rec.assistedRep = linkedJocName || '';
  rec.assistedRepEmail = linkedJocEmail || '';
  if (rec.root || rec.appt) rr.assignedByRoot[rec.root || rec.appt] = advisor;
  if (ownerChanged) {
    swCountClientAdvisorAssignment_(rr, advisor.name, visitAt);
    swTrackClientAdvisorBusySlot_(rr, advisor.name, visitAt, rec);
  }
  summary = summary || {};
  if (summaryKey) summary[summaryKey] = (summary[summaryKey] || 0) + 1;
  return true;
}

function swCurrentClientAdvisorNeedsLinkedJoc_(ss, ctx, advisor, rec) {
  if (!advisor || !advisor.defaultJoc) return false;
  var linkedJoc = swCanonicalWorkflowOwnerForRole_(ss, ctx, advisor.defaultJoc, '', SW_OWNER_ROLES.JOC);
  if (!linkedJoc) return false;
  if (!rec.assistedRep && !rec.assistedRepEmail) return true;
  var currentJoc = swCanonicalWorkflowOwnerForRole_(ss, ctx, rec.assistedRep, rec.assistedRepEmail, SW_OWNER_ROLES.JOC);
  if (currentJoc) {
    if (linkedJoc.email && currentJoc.email && linkedJoc.email === currentJoc.email) return false;
    return swNorm_(linkedJoc.name) !== swNorm_(currentJoc.name);
  }
  return swNorm_(rec.assistedRep) !== swNorm_(linkedJoc.name);
}

function swExistingRootClientAdvisorAssignment_(ss, ctx, rr, visitAt, rec, root) {
  var existing = root && rr.assignedByRoot ? rr.assignedByRoot[root] : null;
  if (!existing || !existing.name) return null;
  var advisor = swRoundRobinAdvisorByName_(rr, existing.name);
  if (!advisor) return null;
  if (!swAdvisorQualifiedForAppointment_(advisor, rec)) return null;
  if (!swAvailabilityFor_(ss, advisor.name, visitAt, ctx).available) return null;
  if (swClientAdvisorHasAppointmentConflict_(rr, advisor.name, visitAt, rec)) return null;
  if (swClientAdvisorBackupShouldYield_(ss, ctx, rr, advisor, visitAt, rec)) return null;
  return advisor;
}

function swClientAdvisorBackupShouldYield_(ss, ctx, rr, advisor, visitAt, rec) {
  if (swAdvisorTierForAppointment_(advisor, rec) !== 'Backup') return false;
  return swHasPrimaryClientAdvisorCandidate_(ss, ctx, rr, visitAt, rec, advisor.name);
}

function swHasPrimaryClientAdvisorCandidate_(ss, ctx, rr, visitAt, rec, excludeName) {
  return (rr.advisors || []).some(function (advisor) {
    if (!advisor || !advisor.name) return false;
    if (excludeName && swNorm_(advisor.name) === swNorm_(excludeName)) return false;
    if (swAdvisorTierForAppointment_(advisor, rec) !== 'Primary') return false;
    if (!swAvailabilityFor_(ss, advisor.name, visitAt, ctx).available) return false;
    if (swClientAdvisorHasAppointmentConflict_(rr, advisor.name, visitAt, rec)) return false;
    return true;
  });
}

function swPickPreviousClientAdvisorForAppointment_(ss, ctx, rr, visitAt, rec) {
  if (!swShouldPreferPreviousClientAdvisor_(rec)) return null;
  var previous = swFindPreviousClientAdvisorForAppointment_(ss, ctx, rr, visitAt, rec);
  if (!previous || !previous.name) return null;
  var advisor = swRoundRobinAdvisorByName_(rr, previous.name);
  if (!advisor) return null;
  if (!swAdvisorQualifiedForAppointment_(advisor, rec)) return null;
  if (!swAvailabilityFor_(ss, advisor.name, visitAt, ctx).available) return null;
  if (swClientAdvisorHasAppointmentConflict_(rr, advisor.name, visitAt, rec)) return null;
  return advisor;
}

function swFindPreviousClientAdvisorForAppointment_(ss, ctx, rr, visitAt, rec) {
  var now = ctx && ctx.now instanceof Date ? ctx.now : new Date();
  var currentMs = visitAt instanceof Date ? visitAt.getTime() : 0;
  var nowMs = now.getTime();
  var candidates = [];
  (rr.appointments || []).forEach(function (other) {
    if (!other || !swIsAppointmentActive_(other)) return;
    if (other.row && rec.row && other.row === rec.row) return;
    if (other.appt && rec.appt && other.appt === rec.appt) return;
    if (!swPreviousAppointmentMatches_(rec, other)) return;
    var otherVisitAt = swVisitDateTime_(other, ctx.tz);
    if (!otherVisitAt) return;
    var otherMs = otherVisitAt.getTime();
    if (!(otherMs < currentMs && otherMs < nowMs)) return;
    var owner = swCanonicalClientAdvisorOwner_(ss, ctx, other.assignedRep, other.assignedRepEmail);
    if (!owner) return;
    candidates.push({ visitAt: otherVisitAt, owner: owner });
  });
  if (!candidates.length) return null;
  candidates.sort(function (a, b) {
    return b.visitAt.getTime() - a.visitAt.getTime();
  });
  return candidates[0].owner;
}

function swPreviousAppointmentMatches_(rec, other) {
  var root = swTrim_((rec && rec.root) || (rec && rec.appt) || '');
  if (root) return swTrim_((other && other.root) || (other && other.appt) || '') === root;
  var email = swNormEmail_((rec && rec.email) || '');
  if (email && email === swNormEmail_((other && other.email) || '')) return true;
  var phone = swNormPhone_((rec && rec.phone) || '');
  return !!(phone && phone === swNormPhone_((other && other.phone) || ''));
}

function swShouldPreferPreviousClientAdvisor_(rec) {
  var visitType = swNorm_((rec && rec.visitType) || '');
  if (visitType.indexOf('diamond viewing') >= 0) return true;
  if (/\b(2nd|second|follow up|follow-up|returning)\b/.test(visitType)) return true;
  var visitNumber = Number(String((rec && rec.visitNumber) || '').replace(/[^0-9.]+/g, ''));
  return isFinite(visitNumber) && visitNumber > 1;
}

function swCanonicalClientAdvisorOwner_(ss, ctx, name, email) {
  var owner = swCanonicalWorkflowOwnerForRole_(ss, ctx, name, email, SW_OWNER_ROLES.SALES_REP);
  if (owner) return owner;
  var raw = swTrim_(name || '');
  if (!raw) return null;
  var pieces = raw.split(/[,\n;]/);
  for (var i = 0; i < pieces.length; i++) {
    var cleaned = swStripWorkflowOwnerQualifier_(pieces[i]);
    if (!cleaned) continue;
    owner = swCanonicalWorkflowOwnerForRole_(ss, ctx, cleaned, '', SW_OWNER_ROLES.SALES_REP);
    if (owner) return owner;
  }
  return null;
}

function swStripWorkflowOwnerQualifier_(value) {
  return swTrim_(String(value || '').replace(/\s*\([^)]*\)\s*/g, ' '));
}

function swRoundRobinAdvisorByName_(rr, name) {
  var target = swNorm_(name);
  for (var i = 0; i < (rr.advisors || []).length; i++) {
    if (swNorm_(rr.advisors[i].name) === target) return rr.advisors[i];
  }
  return null;
}

function swPickClientAdvisorRoundRobin_(ss, ctx, rr, visitAt, rec, excludeName) {
  var dateKey = swDateKey_(visitAt);
  var counts = rr.countsByDate[dateKey] || {};
  var candidates = (rr.advisors || []).filter(function (advisor) {
    if (!advisor || !advisor.name) return false;
    if (excludeName && swNorm_(advisor.name) === swNorm_(excludeName)) return false;
    if (!swAvailabilityFor_(ss, advisor.name, visitAt, ctx).available) return false;
    if (swClientAdvisorHasAppointmentConflict_(rr, advisor.name, visitAt, rec)) return false;
    return swAdvisorQualifiedForAppointment_(advisor, rec);
  });
  candidates = swPrioritizeDiamondAdvisorPool_(candidates, rec);
  if (!candidates.length) return null;
  candidates.sort(function (a, b) {
    var ac = counts[swNorm_(a.name)] || 0;
    var bc = counts[swNorm_(b.name)] || 0;
    if (ac !== bc) return ac - bc;
    return String(a.name || '').localeCompare(String(b.name || ''));
  });
  return candidates[0];
}

function swAdvisorQualifiedForAppointment_(advisor, rec) {
  return swAdvisorTierForAppointment_(advisor, rec) !== 'None';
}

function swAdvisorTierForAppointment_(advisor, rec) {
  var skills = (advisor && advisor.skills) || swDefaultRepSkills_();
  var kind = swAppointmentAdvisorKind_(rec);
  if (kind === 'natural') return swNormalizeNaturalSkill_(skills.naturalDiamond);
  if (kind === 'lab') return swNormalizeLabSkill_(skills.labDiamond);
  if (kind === 'diamond') {
    return swBestAdvisorSkillTier_([
      swNormalizeLabSkill_(skills.labDiamond),
      swNormalizeNaturalSkill_(skills.naturalDiamond)
    ]);
  }
  return swNormalizeGeneralSkill_(skills.generalAppointment);
}

function swBestAdvisorSkillTier_(tiers) {
  tiers = tiers || [];
  for (var i = 0; i < tiers.length; i++) {
    if (tiers[i] === 'Primary') return 'Primary';
  }
  for (var j = 0; j < tiers.length; j++) {
    if (tiers[j] === 'Backup') return 'Backup';
  }
  return 'None';
}

function swPrioritizeDiamondAdvisorPool_(candidates, rec) {
  return swPrioritizeAdvisorPoolByTier_(candidates, function (advisor) {
    return swAdvisorTierForAppointment_(advisor, rec);
  });
}

function swPrioritizeNaturalAdvisorPool_(candidates, rec) {
  if (swAppointmentDiamondKind_(rec) !== 'natural') return candidates;
  return swPrioritizeAdvisorPoolByTier_(candidates, function (advisor) {
    return swNormalizeNaturalSkill_((advisor.skills || {}).naturalDiamond);
  });
}

function swPrioritizeAdvisorPoolByTier_(candidates, tierFn) {
  var primary = candidates.filter(function (advisor) {
    return tierFn(advisor) === 'Primary';
  });
  return primary.length ? primary : candidates.filter(function (advisor) {
    return tierFn(advisor) === 'Backup';
  });
}

function swCountClientAdvisorAssignment_(rr, advisorName, visitAt) {
  var dateKey = swDateKey_(visitAt);
  if (!rr.countsByDate[dateKey]) rr.countsByDate[dateKey] = {};
  var ownerKey = swNorm_(advisorName);
  rr.countsByDate[dateKey][ownerKey] = (rr.countsByDate[dateKey][ownerKey] || 0) + 1;
}

function swTrackClientAdvisorBusySlot_(rr, advisorName, visitAt, rec) {
  var slotKey = swAppointmentSlotKey_(visitAt);
  var ownerKey = swNorm_(advisorName);
  if (!slotKey || !ownerKey) return;
  if (!rr.busyBySlotOwner) rr.busyBySlotOwner = {};
  if (!rr.busyBySlotOwner[slotKey]) rr.busyBySlotOwner[slotKey] = {};
  if (!rr.busyBySlotOwner[slotKey][ownerKey]) rr.busyBySlotOwner[slotKey][ownerKey] = [];
  rr.busyBySlotOwner[slotKey][ownerKey].push(swAppointmentConflictRef_(rec));
}

function swClientAdvisorHasAppointmentConflict_(rr, advisorName, visitAt, rec) {
  var slotKey = swAppointmentSlotKey_(visitAt);
  var ownerKey = swNorm_(advisorName);
  var rows = rr && rr.busyBySlotOwner && rr.busyBySlotOwner[slotKey] && rr.busyBySlotOwner[slotKey][ownerKey];
  if (!rows || !rows.length) return false;
  var current = swAppointmentConflictRef_(rec);
  return rows.some(function (item) {
    return !swSameAppointmentConflictRef_(item, current);
  });
}

function swAppointmentSlotKey_(visitAt) {
  if (!(visitAt instanceof Date) || isNaN(visitAt.getTime())) return '';
  return swDateKey_(visitAt) + '|' + swPad2_(visitAt.getHours()) + ':' + swPad2_(visitAt.getMinutes());
}

function swAppointmentConflictRef_(rec) {
  return {
    row: Number(rec && rec.row) || 0,
    root: swTrim_((rec && rec.root) || ''),
    appt: swTrim_((rec && rec.appt) || '')
  };
}

function swSameAppointmentConflictRef_(a, b) {
  if (a.row && b.row && a.row === b.row) return true;
  if (a.appt && b.appt && a.appt === b.appt) return true;
  return !!(a.root && b.root && a.root === b.root);
}

function swAppointmentDiamondKind_(rec) {
  var kind = swAppointmentAdvisorKind_(rec);
  return kind === 'natural' || kind === 'lab' ? kind : '';
}

function swAppointmentAdvisorKind_(rec) {
  var s = swNorm_([
    rec && rec.diamondType,
    rec && rec.dvCustomerRequirementsJson,
    rec && rec.dvCustomerLookingFor,
    rec && rec.visitType
  ].join(' '));
  if (s.indexOf('natural') >= 0) return 'natural';
  if (s.indexOf('lab') >= 0) return 'lab';
  if (swNorm_((rec && rec.visitType) || '').indexOf('diamond viewing') >= 0) return 'diamond';
  return 'general';
}

function swWriteAppointmentOwnerAssignmentToMaster_(ss, rec, advisor, linkedJoc, linkedJocEmail) {
  var row = Number(rec && rec.row);
  if (!row || row < 2 || typeof swEnsureMasterOwnerHeaders_ !== 'function') return false;
  var master = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!master) return false;
  var headers = swEnsureMasterOwnerHeaders_(master);
  master.getRange(row, headers.assignedRep).setValue(advisor.name || '');
  master.getRange(row, headers.assignedRepEmail).setValue(advisor.email || '');
  var canonicalJoc = linkedJoc && !linkedJocEmail ? swCanonicalWorkflowOwnerForRole_(ss, {}, linkedJoc, '', SW_OWNER_ROLES.JOC) : null;
  master.getRange(row, headers.assistedRep).setValue(linkedJoc || '');
  master.getRange(row, headers.assistedRepEmail).setValue(linkedJocEmail || (canonicalJoc ? canonicalJoc.email : ''));
  return true;
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
  map[SW_TASKS.POST_CONSULT_STATUS] = 'Post-Consult Ops';
  map[SW_TASKS.START_3D] = 'Post-Consult Ops';
  map[SW_TASKS.RECORD_3D_DEADLINE] = 'Post-Consult Ops';
  map[SW_TASKS.REQUEST_WAX] = 'Wax';
  map[SW_TASKS.UPDATE_WAX] = 'Wax';
  map[SW_TASKS.DIAMOND_PROPOSE] = 'Diamond Viewing';
  map[SW_TASKS.DIAMOND_QUOTE] = 'Diamond Viewing';
  map[SW_TASKS.DIAMOND_ORDER] = 'Diamond Order';
  map[SW_TASKS.DIAMOND_TRACK] = 'Diamond Tracking';
  map[SW_TASKS.DIAMOND_DELIVERY] = 'Diamond Delivery';
  map[SW_TASKS.DIAMOND_DECISIONS] = 'Diamond Decisions';
  map[SW_TASKS.DIAMOND_RETURN] = 'Diamond Return';
  map[SW_TASKS.DIAMOND_ORDER_ACK_REP] = 'Diamond Ordered';
  map[SW_TASKS.DIAMOND_ORDER_ACK_JOC] = 'Diamond Ordered';
  map[SW_TASKS.DIAMOND_ETA_REP] = 'Diamond ETA Risk';
  map[SW_TASKS.DIAMOND_ETA_JOC] = 'Diamond ETA Risk';
  map[SW_TASKS.DATA_CLEANUP_REVIEW] = 'Data Cleanup';
  map[SW_TASKS.DATA_CLEANUP_CONFIRM] = 'Data Cleanup';
  map[SW_TASKS.DATA_CLEANUP_REVISE] = 'Data Cleanup';
  return map[taskType] || '';
}
