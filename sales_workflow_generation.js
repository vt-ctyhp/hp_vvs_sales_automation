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
      internalSummary: aiSummary.internalSummary || '',
      customerInsights: aiSummary.customerInsights || '',
      recommendedNextSteps: aiSummary.recommendedNextSteps || '',
      confidenceFlags: aiSummary.confidenceFlags || '',
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
      summaryDocUrl: swDeepValue_(approvePayload, ['extra', 'summaryDocUrl']) || '',
      summaryJsonUrl: swDeepValue_(approvePayload, ['extra', 'summaryJsonUrl']) || '',
      internalSummary: swDeepValue_(approvePayload, ['extra', 'internalSummary']) || '',
      customerInsights: swDeepValue_(approvePayload, ['extra', 'customerInsights']) || '',
      recommendedNextSteps: swDeepValue_(approvePayload, ['extra', 'recommendedNextSteps']) || '',
      confidenceFlags: swDeepValue_(approvePayload, ['extra', 'confidenceFlags']) || ''
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
    var repName = rec.assignedRep || '';
    var repEmail = swLookupEmailByName_(ss, repName, ctx) || '';
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
  var intendedName = rec.assistedRep || '';
  var intendedEmail = swLookupEmailByName_(ss, intendedName, ctx) || swSchedulePersonEmailByName_(ctx, intendedName);
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

function swSchedulePersonEmailByName_(ctx, name) {
  var target = swNorm_(name);
  if (!target) return '';
  var people = (ctx && ctx.employeeSchedulePeople) || [];
  for (var i = 0; i < people.length; i++) {
    if (swNorm_(people[i].name) === target) return swNormEmail_(people[i].email || '');
  }
  return '';
}

function swPrepareClientAdvisorRoundRobin_(ss, ctx, appointments) {
  ctx = ctx || {};
  var enabled = swTruthy_(swConfigValue_(ctx.config || [], 'SYSTEM', 'CLIENT_ADVISOR_ROUND_ROBIN', 'N'));
  var people = typeof swReadEmployeeSchedulePeople_ === 'function' ? swReadEmployeeSchedulePeople_(ss) : [];
  ctx.employeeSchedulePeople = people;
  var state = {
    enabled: enabled,
    countsByDate: {},
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
    if (!rec || !rec.assignedRep) return;
    var visitAt = swVisitDateTime_(rec, ctx.tz);
    if (!visitAt) return;
    var dateKey = swDateKey_(visitAt);
    var ownerKey = swNorm_(rec.assignedRep);
    if (!state.countsByDate[dateKey]) state.countsByDate[dateKey] = {};
    state.countsByDate[dateKey][ownerKey] = (state.countsByDate[dateKey][ownerKey] || 0) + 1;
    if (rec.root) state.assignedByRoot[rec.root] = {
      name: rec.assignedRep,
      email: rec.assignedRepEmail || swLookupEmailByName_(ss, rec.assignedRep, ctx),
      defaultJoc: rec.assistedRep || ''
    };
  });
  return state;
}

function swMaybeAutoAssignClientAdvisor_(ss, ctx, rec, summary) {
  var rr = ctx && ctx.clientAdvisorRoundRobin;
  if (!rr || !rr.enabled || !rec) return false;
  var visitAt = swVisitDateTime_(rec, ctx.tz);
  if (!visitAt) return false;

  var currentUnavailable = false;
  if (rec.assignedRep) {
    var currentAvail = swAvailabilityFor_(ss, rec.assignedRep, visitAt, ctx);
    currentUnavailable = currentAvail.known && !currentAvail.available;
    if (!currentUnavailable) {
      var existingAdvisor = swRoundRobinAdvisorByName_(rr, rec.assignedRep);
      if (existingAdvisor && existingAdvisor.defaultJoc && !rec.assistedRep) {
        var existingLinkedEmail = swLookupEmailByName_(ss, existingAdvisor.defaultJoc, ctx) || swSchedulePersonEmailByName_(ctx, existingAdvisor.defaultJoc);
        if (swWriteAppointmentOwnerAssignmentToMaster_(ss, rec, existingAdvisor, existingAdvisor.defaultJoc, existingLinkedEmail)) {
          rec.assistedRep = existingAdvisor.defaultJoc;
          rec.assistedRepEmail = existingLinkedEmail;
          summary.autoLinkedJocFromAdvisor = (summary.autoLinkedJocFromAdvisor || 0) + 1;
          return true;
        }
      }
      return false;
    }
  }

  var root = rec.root || rec.appt || '';
  var existingRootAssignment = !rec.assignedRep && root ? rr.assignedByRoot[root] : null;
  var chosen = existingRootAssignment || swPickClientAdvisorRoundRobin_(ss, ctx, rr, visitAt, rec, rec.assignedRep || '');
  if (!chosen || !chosen.name) {
    summary.clientAdvisorRoundRobinNoOwner = (summary.clientAdvisorRoundRobinNoOwner || 0) + 1;
    return false;
  }

  var linkedJoc = swTrim_(chosen.defaultJoc || '');
  var linkedJocEmail = linkedJoc ? (swLookupEmailByName_(ss, linkedJoc, ctx) || swSchedulePersonEmailByName_(ctx, linkedJoc)) : '';
  if (!swWriteAppointmentOwnerAssignmentToMaster_(ss, rec, chosen, linkedJoc, linkedJocEmail)) return false;
  rec.assignedRep = chosen.name;
  rec.assignedRepEmail = chosen.email || '';
  rec.assistedRep = linkedJoc || rec.assistedRep || '';
  rec.assistedRepEmail = linkedJoc ? linkedJocEmail : rec.assistedRepEmail;
  if (root) rr.assignedByRoot[root] = chosen;
  var dateKey = swDateKey_(visitAt);
  if (!rr.countsByDate[dateKey]) rr.countsByDate[dateKey] = {};
  var ownerKey = swNorm_(chosen.name);
  rr.countsByDate[dateKey][ownerKey] = (rr.countsByDate[dateKey][ownerKey] || 0) + 1;
  if (currentUnavailable) summary.autoReassignedClientAdvisors = (summary.autoReassignedClientAdvisors || 0) + 1;
  else summary.autoAssignedClientAdvisors = (summary.autoAssignedClientAdvisors || 0) + 1;
  return true;
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
    return swAdvisorQualifiedForAppointment_(advisor, rec);
  });
  candidates = swPrioritizeNaturalAdvisorPool_(candidates, rec);
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
  var skills = advisor.skills || swDefaultRepSkills_();
  var kind = swAppointmentDiamondKind_(rec);
  if (kind === 'natural') return swNormalizeNaturalSkill_(skills.naturalDiamond) !== 'None';
  if (kind === 'lab') return swTruthy_(skills.labDiamond);
  return swTruthy_(skills.generalAppointment);
}

function swPrioritizeNaturalAdvisorPool_(candidates, rec) {
  if (swAppointmentDiamondKind_(rec) !== 'natural') return candidates;
  var primary = candidates.filter(function (advisor) {
    return swNormalizeNaturalSkill_((advisor.skills || {}).naturalDiamond) === 'Primary';
  });
  return primary.length ? primary : candidates.filter(function (advisor) {
    return swNormalizeNaturalSkill_((advisor.skills || {}).naturalDiamond) === 'Backup';
  });
}

function swAppointmentDiamondKind_(rec) {
  var s = swNorm_((rec && rec.diamondType) || (rec && rec.dvCustomerRequirementsJson) || (rec && rec.dvCustomerLookingFor) || '');
  if (s.indexOf('natural') >= 0) return 'natural';
  if (s.indexOf('lab') >= 0) return 'lab';
  return '';
}

function swWriteAppointmentOwnerAssignmentToMaster_(ss, rec, advisor, linkedJoc, linkedJocEmail) {
  var row = Number(rec && rec.row);
  if (!row || row < 2 || typeof swEnsureMasterOwnerHeaders_ !== 'function') return false;
  var master = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!master) return false;
  var headers = swEnsureMasterOwnerHeaders_(master);
  master.getRange(row, headers.assignedRep).setValue(advisor.name || '');
  master.getRange(row, headers.assignedRepEmail).setValue(advisor.email || '');
  if (linkedJoc) {
    master.getRange(row, headers.assistedRep).setValue(linkedJoc);
    master.getRange(row, headers.assistedRepEmail).setValue(linkedJocEmail || swLookupEmailByName_(ss, linkedJoc, {}) || '');
  }
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
