/**
 * Stale customer data cleanup workflow.
 *
 * One-time campaign cases use a temporary Cleanup tab while enabled. The same
 * stale-record logic continues to create normal queue tasks after the campaign
 * tab is disabled.
 */

var SW_DATA_CLEANUP_STATUS = {
  OPEN: 'Open',
  PENDING_CONFIRMATION: 'Pending Confirmation',
  RETURNED: 'Returned',
  APPLIED: 'Applied'
};

function swIsDataCleanupTaskType_(taskType) {
  return [
    SW_TASKS.DATA_CLEANUP_REVIEW,
    SW_TASKS.DATA_CLEANUP_CONFIRM,
    SW_TASKS.DATA_CLEANUP_REVISE
  ].indexOf(taskType) >= 0;
}

function swDataCleanupEnabled_(config) {
  return swNorm_(swConfigValue_(config || [], 'SYSTEM', 'DATA_CLEANUP_ENABLED', 'Y')) !== 'n';
}

function swDataCleanupCampaignTabEnabled_(config) {
  return swNorm_(swConfigValue_(config || [], 'SYSTEM', 'DATA_CLEANUP_CAMPAIGN_TAB_ENABLED', 'Y')) !== 'n';
}

function swGenerateDataCleanupTasks_(ss, taskState, ctx, appointments, now, summary) {
  ctx = ctx || {};
  now = now || new Date();
  var config = ctx.config || swReadConfig_(ss, true);
  var out = {
    enabled: swDataCleanupEnabled_(config),
    campaignTabEnabled: swDataCleanupCampaignTabEnabled_(config),
    scannedRoots: 0,
    candidates: 0,
    casesCreated: 0,
    tasksCreated: 0,
    skippedUnresolved: 0,
    skippedResolvedSameTouch: 0,
    tabDisabled: false
  };
  if (!out.enabled) return out;

  var staleDays = Number(swConfigValue_(config, 'SYSTEM', 'DATA_CLEANUP_STALE_DAYS', '30')) || 30;
  var campaignId = swTrim_(swConfigValue_(config, 'SYSTEM', 'DATA_CLEANUP_CAMPAIGN_ID', 'ONE_TIME_2026_05')) || 'ONE_TIME_2026_05';
  var groups = swDataCleanupRowsByRoot_(appointments || []);
  var currentByRoot = swDataCleanupCurrentRowsByRoot_(groups);
  var rootIndex = swDataCleanupReadRootIndex_(ss);
  var caseState = swReadDataCleanupCaseState_(ss);
  var roots = Object.keys(currentByRoot);
  out.existingJocTasksRestored = swDataCleanupRestoreCampaignJocTasks_(ss, taskState, ctx, caseState, campaignId, now);
  if (summary && out.existingJocTasksRestored) summary.updated += out.existingJocTasksRestored;

  roots.forEach(function (root) {
    out.scannedRoots++;
    if (caseState.unresolvedByRoot[root]) {
      out.skippedUnresolved++;
      return;
    }

    var rec = currentByRoot[root];
    if (!rec || !swIsAppointmentActive_(rec)) return;
    var stage = swDataCleanupPipelineStage_(rec, groups[root] || []);
    if (!swDataCleanupStageNeedsReview_(stage.key)) return;

    var lastTouch = swDataCleanupLastTouch_(rec, groups[root] || [], rootIndex.byRoot[root], now);
    if (!lastTouch) return;
    var stale = Math.floor((swDataCleanupStartOfDay_(now).getTime() - swDataCleanupStartOfDay_(lastTouch).getTime()) / 86400000);
    if (stale < staleDays) return;
    out.candidates++;

    var caseId = swDataCleanupCaseId_(campaignId, root, lastTouch);
    if (caseState.byId[caseId] && !swDataCleanupCaseUnresolved_(caseState.byId[caseId])) {
      out.skippedResolvedSameTouch++;
      return;
    }

    var createdAt = swIso_(now);
    var cleanupCase = {
      caseId: caseId,
      campaignId: campaignId,
      root: root,
      appt: rec.appt || '',
      customerName: rec.name || 'No customer',
      brand: rec.brand || '',
      clientAdvisor: rec.assignedRep || '',
      clientAdvisorEmail: swLookupEmailByName_(ss, rec.assignedRep || '', ctx) || rec.assignedRepEmail || '',
      joc: rec.assistedRep || '',
      jocEmail: swLookupEmailByName_(ss, rec.assistedRep || '', ctx) || rec.assistedRepEmail || '',
      stageKey: stage.key,
      currentSalesStage: rec.salesStage || '',
      currentConvStatus: rec.convStatus || '',
      currentCustomOrder: rec.customOrder || '',
      currentInProduction: rec.inProduction || '',
      lastTouchAt: swIso_(lastTouch),
      staleDays: stale,
      status: SW_DATA_CLEANUP_STATUS.OPEN,
      campaignTab: out.campaignTabEnabled ? 'Y' : 'N',
      proposalJson: '',
      proposedBy: '',
      proposedByEmail: '',
      proposedRole: '',
      proposedAt: '',
      confirmationBy: '',
      confirmationEmail: '',
      confirmedAt: '',
      returnReason: '',
      returnedBy: '',
      returnedAt: '',
      revisionCount: 0,
      appliedAt: '',
      appliedResultJson: '',
      createdAt: createdAt,
      updatedAt: createdAt
    };
    swWriteDataCleanupCase_(ss, cleanupCase);
    caseState.byId[caseId] = cleanupCase;
    caseState.unresolvedByRoot[root] = cleanupCase;
    out.casesCreated++;

    var repTask = swBuildDataCleanupTask_(ss, taskState, ctx, rec, cleanupCase, SW_TASKS.DATA_CLEANUP_REVIEW, SW_OWNER_ROLES.SALES_REP, now, '', now, {
      phase: 'review',
      taskId: swDataCleanupTaskId_(caseId, 'REVIEW', 'REP', 0)
    });
    var jocTask = swBuildDataCleanupTask_(ss, taskState, ctx, rec, cleanupCase, SW_TASKS.DATA_CLEANUP_REVIEW, SW_OWNER_ROLES.JOC, now, '', now, {
      phase: 'review',
      taskId: swDataCleanupTaskId_(caseId, 'REVIEW', 'JOC', 0)
    });
    swUpsertTask_(ss, taskState, repTask, summary);
    swUpsertTask_(ss, taskState, jocTask, summary);
    out.tasksCreated += 2;
  });

  if (out.campaignTabEnabled && !out.casesCreated && !swDataCleanupHasUnresolvedCampaign_(caseState, campaignId)) {
    swDataCleanupSetConfigValue_(ss, 'DATA_CLEANUP_CAMPAIGN_TAB_ENABLED', 'N');
    out.tabDisabled = true;
    out.campaignTabEnabled = false;
  }

  return out;
}

function swBuildDataCleanupTask_(ss, state, ctx, rec, cleanupCase, taskType, ownerRole, dueAt, dependencyTaskId, now, options) {
  options = options || {};
  var template = (ctx.templates && ctx.templates[taskType]) || swDefaultTemplate_(taskType);
  var owner = swDataCleanupUsesDirectCampaignJocOwner_(cleanupCase, ownerRole)
    ? swDataCleanupDirectCampaignJocOwner_(ss, ctx, rec)
    : swResolveOwner_(ss, ctx, rec, ownerRole, dueAt || now, null);
  if (options.ownerName || options.ownerEmail) {
    owner.currentOwner = options.ownerName || owner.currentOwner;
    owner.currentOwnerEmail = swNormEmail_(options.ownerEmail || owner.currentOwnerEmail);
    owner.intendedOwner = options.ownerName || owner.intendedOwner;
    owner.intendedOwnerEmail = swNormEmail_(options.ownerEmail || owner.intendedOwnerEmail);
    owner.coverageReason = '';
  }
  var taskId = options.taskId || swDataCleanupTaskId_(cleanupCase.caseId, taskType, ownerRole, cleanupCase.revisionCount || 0);
  var existing = state && state.byId ? state.byId[taskId] : null;
  var visitTime = swFormatAppointmentTime_(rec.visitTime, rec.visitTimeRaw);
  var publicCase = swDataCleanupPublicCase_(cleanupCase);
  var payload = {
    appointment: {
      row: rec.row || '',
      root: rec.root || cleanupCase.root || '',
      appt: rec.appt || cleanupCase.appt || '',
      uid: rec.uid || '',
      customerName: rec.name || cleanupCase.customerName || '',
      email: rec.email || '',
      phone: rec.phone || '',
      brand: rec.brand || cleanupCase.brand || '',
      visitDate: rec.visitDate || '',
      visitTime: visitTime,
      visitType: rec.visitType || '',
      assignedRep: rec.assignedRep || cleanupCase.clientAdvisor || '',
      assignedRepEmail: rec.assignedRepEmail || cleanupCase.clientAdvisorEmail || '',
      assistedRep: rec.assistedRep || cleanupCase.joc || '',
      assistedRepEmail: rec.assistedRepEmail || cleanupCase.jocEmail || '',
      clientFolder: rec.clientFolder || '',
      reportUrl: rec.reportUrl || '',
      quotationUrl: rec.quotationUrl || '',
      tracker3dUrl: rec.tracker3dUrl || '',
      salesStage: rec.salesStage || cleanupCase.currentSalesStage || '',
      convStatus: rec.convStatus || cleanupCase.currentConvStatus || '',
      customOrder: rec.customOrder || cleanupCase.currentCustomOrder || '',
      inProduction: rec.inProduction || cleanupCase.currentInProduction || '',
      centerStoneStatus: rec.centerStoneStatus || '',
      nextSteps: rec.nextSteps || '',
      designRequest: rec.designRequest || '',
      deadline3d: rec.deadline3d || '',
      productionDeadline: rec.productionDeadline || '',
      waxStatus: rec.waxStatus || '',
      waxDeadlineAdmin: rec.waxDeadlineAdmin || '',
      waxRequestUrl: rec.waxRequestUrl || '',
      dvStonesSummary: rec.dvStonesSummary || '',
      so: rec.so || '',
      orderFolder: rec.orderFolder || '',
      remainingBalance: rec.remainingBalance || '',
      orderTotal: rec.orderTotal || '',
      paidToDate: rec.paidToDate || '',
      lastPaymentDate: rec.lastPaymentDate || '',
      orderDate: rec.orderDate || ''
    },
    extra: {
      cleanupCase: publicCase,
      cleanupPhase: options.phase || '',
      cleanupProposalSummary: swDataCleanupProposalSummary_(publicCase.proposal || {}),
      cleanupProposedBy: cleanupCase.proposedBy || '',
      cleanupReturnReason: cleanupCase.returnReason || ''
    }
  };

  return {
    taskId: taskId,
    root: rec.root || cleanupCase.root || '',
    appt: rec.appt || cleanupCase.appt || '',
    customerName: rec.name || cleanupCase.customerName || '',
    brand: rec.brand || cleanupCase.brand || '',
    visitDate: rec.visitDate || '',
    visitTime: visitTime,
    visitType: rec.visitType || '',
    lifecycleStage: cleanupCase.campaignTab === 'Y' ? 'Cleanup Campaign' : 'Data Cleanup',
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

function swValidateDataCleanupCompletion_(task, data) {
  if (!swIsDataCleanupTaskType_(task.taskType)) return;
  data = data || {};
  if (task.taskType === SW_TASKS.DATA_CLEANUP_CONFIRM) {
    var decision = swNorm_(data.cleanupDecision);
    if (decision !== 'approve' && decision !== 'return') throw new Error('Choose Approve or Return before submitting cleanup confirmation.');
    if (decision === 'return' && !swTrim_(data.cleanupReturnReason)) throw new Error('Enter a return reason before returning this cleanup.');
    return;
  }

  var proposal = swDataCleanupProposalFromData_(data);
  if (!swTrim_(proposal.salesStage)) throw new Error('Select Sales Stage before submitting cleanup.');
  if (!swTrim_(proposal.convStatus)) throw new Error('Select Conversion Status before submitting cleanup.');
  var verification = proposal.verification || {};
  ['ownersVerified', 'contactVerified', 'opsStatusReviewed', 'nextStepsCurrent'].forEach(function (key) {
    if (!verification[key]) throw new Error('Complete all cleanup verification checks before submitting.');
  });
  if (swDataCleanupProposalIsLost_(proposal)) {
    if (!swTrim_(proposal.lostLeadReason)) throw new Error('Select a Lost Lead reason before submitting cleanup.');
    if (!swTrim_(proposal.lostLeadNotes)) throw new Error('Enter Lost Lead reason notes before submitting cleanup.');
  }
}

function swHandleDataCleanupTaskCompletion_(ss, task, data, user) {
  if (!swIsDataCleanupTaskType_(task.taskType)) return null;
  var cleanupCase = swReadDataCleanupCaseByTask_(ss, task);
  if (!cleanupCase) throw new Error('Cleanup case not found for this task.');
  if (cleanupCase.status === SW_DATA_CLEANUP_STATUS.APPLIED) throw new Error('This cleanup case has already been applied.');

  if (task.taskType === SW_TASKS.DATA_CLEANUP_CONFIRM) {
    if (cleanupCase.status !== SW_DATA_CLEANUP_STATUS.PENDING_CONFIRMATION) {
      throw new Error(swDataCleanupInactiveTaskMessage_(ss, task) || 'This cleanup confirmation is no longer active. Refresh Queue to load the current cleanup task.');
    }
    return swCompleteDataCleanupConfirmation_(ss, task, cleanupCase, data, user);
  }
  if (task.taskType === SW_TASKS.DATA_CLEANUP_REVIEW && cleanupCase.status !== SW_DATA_CLEANUP_STATUS.OPEN) {
    throw new Error(swDataCleanupInactiveTaskMessage_(ss, task) || 'This cleanup review is no longer active. Refresh Queue to load the current cleanup task.');
  }
  if (task.taskType === SW_TASKS.DATA_CLEANUP_REVISE && cleanupCase.status !== SW_DATA_CLEANUP_STATUS.RETURNED) {
    throw new Error(swDataCleanupInactiveTaskMessage_(ss, task) || 'This cleanup revision is no longer active. Refresh Queue to load the current cleanup task.');
  }
  return swCompleteDataCleanupProposal_(ss, task, cleanupCase, data, user);
}

function swCompleteDataCleanupProposal_(ss, task, cleanupCase, data, user) {
  var timingStep = typeof swStepTimer_ === 'function' ? swStepTimer_('swCompleteDataCleanupProposal_') : null;
  var now = new Date();
  var proposal = swDataCleanupProposalFromData_(data);
  var proposerRole = task.ownerRole || SW_OWNER_ROLES.SALES_REP;
  cleanupCase.status = SW_DATA_CLEANUP_STATUS.PENDING_CONFIRMATION;
  cleanupCase.proposalJson = swStringify_(proposal);
  cleanupCase.proposedBy = user.name || user.email || task.currentOwner || '';
  cleanupCase.proposedByEmail = user.email || task.currentOwnerEmail || '';
  cleanupCase.proposedRole = proposerRole;
  cleanupCase.proposedAt = swIso_(now);
  cleanupCase.returnReason = '';
  cleanupCase.returnedBy = '';
  cleanupCase.returnedAt = '';
  cleanupCase.updatedAt = cleanupCase.proposedAt;
  if (timingStep) timingStep('prepare_case_update', { caseId: cleanupCase.caseId });
  swWriteDataCleanupCase_(ss, cleanupCase);
  if (timingStep) timingStep('swWriteDataCleanupCase_', { caseId: cleanupCase.caseId });

  swDataCleanupBlockTasksForCase_(ss, cleanupCase.caseId, task.taskId, user, 'SUPERSEDED_BY_PROPOSAL', function (candidate) {
    return candidate.taskType === SW_TASKS.DATA_CLEANUP_REVIEW || candidate.taskType === SW_TASKS.DATA_CLEANUP_REVISE;
  });
  if (timingStep) timingStep('swDataCleanupBlockTasksForCase_', { caseId: cleanupCase.caseId });

  var payload = swParseJson_(task.payloadJson, {});
  var rec = swDataCleanupRecFromTask_(task, payload, cleanupCase);
  var confirmRole = swDataCleanupOppositeRole_(proposerRole);
  var ctx = swBuildDataCleanupImmediateContext_(ss);
  if (timingStep) timingStep('swBuildDataCleanupImmediateContext_', { caseId: cleanupCase.caseId });
  var confirmTask = swBuildDataCleanupTask_(ss, null, ctx, rec, cleanupCase, SW_TASKS.DATA_CLEANUP_CONFIRM, confirmRole, now, task.taskId, now, {
    phase: 'confirm',
    taskId: swDataCleanupTaskId_(cleanupCase.caseId, 'CONFIRM', swDataCleanupRoleKey_(confirmRole), Number(cleanupCase.revisionCount) || 0)
  });
  if (timingStep) timingStep('swBuildDataCleanupTask_', { taskId: confirmTask.taskId });
  swDataCleanupUpsertImmediateTask_(ss, confirmTask, user, 'CREATE');
  if (timingStep) timingStep('swDataCleanupUpsertImmediateTask_', { taskId: confirmTask.taskId });
  return {
    action: 'DATA_CLEANUP_PROPOSED',
    caseId: cleanupCase.caseId,
    confirmationOwner: confirmTask.currentOwner,
    confirmationTaskId: confirmTask.taskId
  };
}

function swCompleteDataCleanupConfirmation_(ss, task, cleanupCase, data, user) {
  var timingStep = typeof swStepTimer_ === 'function' ? swStepTimer_('swCompleteDataCleanupConfirmation_') : null;
  var decision = swNorm_(data.cleanupDecision);
  var now = new Date();
  if (decision === 'return') {
    cleanupCase.status = SW_DATA_CLEANUP_STATUS.RETURNED;
    cleanupCase.returnReason = swTrim_(data.cleanupReturnReason);
    cleanupCase.returnedBy = user.name || user.email || '';
    cleanupCase.returnedAt = swIso_(now);
    cleanupCase.revisionCount = (Number(cleanupCase.revisionCount) || 0) + 1;
    cleanupCase.updatedAt = cleanupCase.returnedAt;
    if (timingStep) timingStep('prepare_return_case_update', { caseId: cleanupCase.caseId });
    swWriteDataCleanupCase_(ss, cleanupCase);
    if (timingStep) timingStep('swWriteDataCleanupCase_', { caseId: cleanupCase.caseId });

    var payload = swParseJson_(task.payloadJson, {});
    var rec = swDataCleanupRecFromTask_(task, payload, cleanupCase);
    var reviseRole = cleanupCase.proposedRole || SW_OWNER_ROLES.SALES_REP;
    var ctx = swBuildDataCleanupImmediateContext_(ss);
    if (timingStep) timingStep('swBuildDataCleanupImmediateContext_', { caseId: cleanupCase.caseId });
    var reviseTask = swBuildDataCleanupTask_(ss, null, ctx, rec, cleanupCase, SW_TASKS.DATA_CLEANUP_REVISE, reviseRole, now, task.taskId, now, {
      phase: 'revise',
      ownerName: cleanupCase.proposedBy,
      ownerEmail: cleanupCase.proposedByEmail,
      taskId: swDataCleanupTaskId_(cleanupCase.caseId, 'REVISE', swDataCleanupRoleKey_(reviseRole), Number(cleanupCase.revisionCount) || 0)
    });
    if (timingStep) timingStep('swBuildDataCleanupTask_', { taskId: reviseTask.taskId });
    swDataCleanupUpsertImmediateTask_(ss, reviseTask, user, 'CREATE');
    if (timingStep) timingStep('swDataCleanupUpsertImmediateTask_', { taskId: reviseTask.taskId });
    return {
      action: 'DATA_CLEANUP_RETURNED',
      caseId: cleanupCase.caseId,
      revisionTaskId: reviseTask.taskId,
      reason: cleanupCase.returnReason
    };
  }

  var proposal = swParseJson_(cleanupCase.proposalJson, {});
  var result = swApplyDataCleanupProposal_(ss, task, cleanupCase, proposal, user, now);
  if (timingStep) timingStep('swApplyDataCleanupProposal_', { caseId: cleanupCase.caseId });
  cleanupCase.status = SW_DATA_CLEANUP_STATUS.APPLIED;
  cleanupCase.confirmationBy = user.name || user.email || '';
  cleanupCase.confirmationEmail = user.email || '';
  cleanupCase.confirmedAt = swIso_(now);
  cleanupCase.appliedAt = cleanupCase.confirmedAt;
  cleanupCase.appliedResultJson = swStringify_(result);
  cleanupCase.updatedAt = cleanupCase.confirmedAt;
  swWriteDataCleanupCase_(ss, cleanupCase);
  if (timingStep) timingStep('swWriteDataCleanupCase_', { caseId: cleanupCase.caseId });
  swDataCleanupBlockTasksForCase_(ss, cleanupCase.caseId, task.taskId, user, 'CLEANUP_APPLIED');
  if (timingStep) timingStep('swDataCleanupBlockTasksForCase_', { caseId: cleanupCase.caseId });
  return {
    action: 'DATA_CLEANUP_APPLIED',
    caseId: cleanupCase.caseId,
    result: result
  };
}

function swBuildDataCleanupImmediateContext_(ss) {
  var ctx = typeof swBuildTaskDetailContext_ === 'function'
    ? swBuildTaskDetailContext_(ss, true)
    : {};
  ctx.rosterIndex = swReadRosterAvailabilityIndex_(ss);
  ctx.scheduleChangesIndex = swReadScheduleChangesIndex_(ss);
  return ctx;
}

function swApplyDataCleanupProposal_(ss, task, cleanupCase, proposal, user, now) {
  var row = swMasterRowForTask_(ss, task);
  if (!(row >= 2)) throw new Error('Could not resolve Master row for cleanup writeback.');
  var payload = swParseJson_(task.payloadJson, {});
  var appt = payload.appointment || {};
  var writebackPayload = {
    assignedRep: appt.assignedRep || cleanupCase.clientAdvisor || '',
    assistedRep: appt.assistedRep || cleanupCase.joc || '',
    salesStage: swTrim_(proposal.salesStage),
    convStatus: swTrim_(proposal.convStatus),
    customOrder: swTrim_(proposal.customOrder),
    cosAllowedEmpty: !swTrim_(proposal.customOrder),
    inProduction: swTrim_(proposal.inProduction),
    centerStone: swTrim_(proposal.centerStone),
    nextSteps: swTrim_(proposal.nextSteps),
    orderDate: swTrim_(proposal.orderDate),
    deadline3d: swTrim_(proposal.deadline3d),
    prodDeadline: swTrim_(proposal.prodDeadline),
    wax: null,
    waxSummary: '',
    notebookLMLink: swTrim_(proposal.notebookLMLink)
  };
  var result;
  if (typeof cs_submitFromDialogForRow_ === 'function') {
    result = cs_submitFromDialogForRow_(row, writebackPayload);
  } else {
    swSetMasterActiveRowForTask_(ss, task);
    result = cs_submitFromDialog(writebackPayload);
  }
  if (result && result.ok === false) throw new Error(result.error || 'Cleanup writeback failed.');
  swDataCleanupWriteMasterAudit_(ss, task, cleanupCase, proposal, now || new Date());
  return { action: 'CUSTOMER_DATA_CLEANUP_WRITEBACK', summary: result && result.summary ? result.summary : result };
}

function swDataCleanupWriteMasterAudit_(ss, task, cleanupCase, proposal, now) {
  var sh = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!sh) throw new Error('Missing sheet: ' + SW_SHEETS.MASTER);
  var row = swMasterRowForTask_(ss, task);
  if (!(row >= 2)) throw new Error('Could not resolve Master row for cleanup writeback.');
  var cReason = swDataCleanupEnsureMasterColumn_(sh, 'Lost Lead Reason');
  var cNotes = swDataCleanupEnsureMasterColumn_(sh, 'Lost Lead Reason Notes');
  var cReviewed = swDataCleanupEnsureMasterColumn_(sh, 'Data Cleanup Reviewed At');
  var cConfirmed = swDataCleanupEnsureMasterColumn_(sh, 'Data Cleanup Confirmed At');
  var lost = swDataCleanupProposalIsLost_(proposal);
  sh.getRange(row, cReason).setValue(lost ? swTrim_(proposal.lostLeadReason) : '');
  sh.getRange(row, cNotes).setValue(lost ? swTrim_(proposal.lostLeadNotes) : '');
  sh.getRange(row, cReviewed).setValue(cleanupCase.proposedAt || '');
  sh.getRange(row, cConfirmed).setValue(swIso_(now || new Date()));
}

function swEnsureDataCleanupMasterHeaders_(ss) {
  try {
    var sh = ss && ss.getSheetByName ? ss.getSheetByName(SW_SHEETS.MASTER) : null;
    if (!sh) return;
    swDataCleanupEnsureMasterColumn_(sh, 'Lost Lead Reason');
    swDataCleanupEnsureMasterColumn_(sh, 'Lost Lead Reason Notes');
    swDataCleanupEnsureMasterColumn_(sh, 'Data Cleanup Reviewed At');
    swDataCleanupEnsureMasterColumn_(sh, 'Data Cleanup Confirmed At');
  } catch (e) {
    try { Logger.log('swEnsureDataCleanupMasterHeaders_ skipped: ' + e.message); } catch (_) {}
  }
}

function swDataCleanupEnsureMasterColumn_(sh, header) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (swHeaderKey_(headers[i]) === swHeaderKey_(header)) return i + 1;
  }
  var col = sh.getLastColumn() + 1;
  sh.getRange(1, col).setValue(header);
  return col;
}

function swDataCleanupProposalFromData_(data) {
  data = data || {};
  var p = data.cleanupProposal || {};
  var verification = p.verification || data.cleanupVerification || {};
  return {
    salesStage: swTrim_(p.salesStage || data.cleanupSalesStage || ''),
    convStatus: swTrim_(p.convStatus || data.cleanupConvStatus || ''),
    customOrder: swTrim_(p.customOrder || data.cleanupCustomOrder || ''),
    inProduction: swTrim_(p.inProduction || data.cleanupInProduction || ''),
    centerStone: swTrim_(p.centerStone || data.cleanupCenterStone || ''),
    orderDate: swTrim_(p.orderDate || data.cleanupOrderDate || ''),
    deadline3d: swTrim_(p.deadline3d || data.cleanupDeadline3d || ''),
    prodDeadline: swTrim_(p.prodDeadline || data.cleanupProdDeadline || ''),
    nextSteps: swTrim_(p.nextSteps || data.cleanupNextSteps || ''),
    notebookLMLink: swTrim_(p.notebookLMLink || data.cleanupNotebookLMLink || ''),
    lostLeadReason: swTrim_(p.lostLeadReason || data.cleanupLostLeadReason || ''),
    lostLeadNotes: swTrim_(p.lostLeadNotes || data.cleanupLostLeadNotes || ''),
    cleanupNotes: swTrim_(p.cleanupNotes || data.cleanupNotes || ''),
    verification: {
      ownersVerified: !!verification.ownersVerified,
      contactVerified: !!verification.contactVerified,
      opsStatusReviewed: !!verification.opsStatusReviewed,
      nextStepsCurrent: !!verification.nextStepsCurrent
    }
  };
}

function swDataCleanupProposalIsLost_(proposal) {
  proposal = proposal || {};
  return /lost/i.test(swTrim_(proposal.salesStage)) || /lost/i.test(swTrim_(proposal.convStatus));
}

function swDataCleanupProposalSummary_(proposal) {
  proposal = proposal || {};
  var parts = [];
  if (proposal.salesStage) parts.push('Sales Stage: ' + proposal.salesStage);
  if (proposal.convStatus) parts.push('Conversion: ' + proposal.convStatus);
  if (proposal.customOrder) parts.push('Custom Order: ' + proposal.customOrder);
  if (proposal.inProduction) parts.push('In Production: ' + proposal.inProduction);
  if (proposal.centerStone) parts.push('Center Stone: ' + proposal.centerStone);
  if (proposal.lostLeadReason) parts.push('Lost reason: ' + proposal.lostLeadReason);
  if (proposal.nextSteps) parts.push('Next steps: ' + proposal.nextSteps);
  return parts.join('\n') || 'No proposal captured yet.';
}

function swDataCleanupFormOptions_(ss, task) {
  var out = {
    salesStages: ['Lead', 'Follow-Up Required', 'Viewing Scheduled', 'Order In Progress', 'Lost Lead'],
    convStatuses: ['Quotation Requested', 'Viewing Scheduled', 'Deposit Paid', 'Confirmed Order', 'Order In Progress', 'Lost Lead'],
    customOrderStatuses: ['', '3D Requested', '3D Revision Requested', '3D Received', 'Approved for Production', 'Waiting Production Timeline', 'In Production', 'Order Completed'],
    inProductionStatuses: ['', 'CAD Approved', 'Casting', 'Setting', 'QC', 'Production Completed'],
    centerStoneStatuses: ['', 'No Center Stone', 'Need to Propose', 'Viewing Scheduled', 'Ordered', 'In Stock', 'Customer Approved'],
    lostLeadReasons: ['No response', 'Budget/timeline not ready', 'Chose another jeweler', 'Unable to contact', 'Cancelled/no-show and no rebook', 'Duplicate/invalid', 'Other']
  };
  try {
    if (typeof readDropdowns_ === 'function') {
      var lists = readDropdowns_() || {};
      out.salesStages = lists.salesStages && lists.salesStages.length ? lists.salesStages : out.salesStages;
      out.convStatuses = lists.convStatuses && lists.convStatuses.length ? lists.convStatuses : out.convStatuses;
      out.customOrderStatuses = lists.customOrderStatuses && lists.customOrderStatuses.length ? [''].concat(lists.customOrderStatuses) : out.customOrderStatuses;
      out.inProductionStatuses = lists.inProductionStatuses && lists.inProductionStatuses.length ? [''].concat(lists.inProductionStatuses) : out.inProductionStatuses;
      out.centerStoneStatuses = lists.centerStoneStatuses && lists.centerStoneStatuses.length ? [''].concat(lists.centerStoneStatuses) : out.centerStoneStatuses;
    }
  } catch (_) {}
  return out;
}

function swReadDataCleanupCaseState_(ss) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.DATA_CLEANUP, SW_DATA_CLEANUP_HEADERS);
  var rows = swReadSheetObjectsExpectedHeaders_(sh, SW_DATA_CLEANUP_HEADERS);
  var out = { byId: {}, unresolvedByRoot: {}, rows: [] };
  rows.forEach(function (row) {
    var item = swDataCleanupCaseFromRow_(row);
    if (!item.caseId) return;
    out.byId[item.caseId] = item;
    out.rows.push(item);
    if (swDataCleanupCaseUnresolved_(item)) out.unresolvedByRoot[item.root] = item;
  });
  return out;
}

function swReadDataCleanupCaseByTask_(ss, task) {
  var payload = swParseJson_(task.payloadJson, {});
  var caseId = swDeepValue_(payload, ['extra', 'cleanupCase', 'caseId']);
  if (!caseId) {
    var parts = String(task.taskId || '').split('|');
    caseId = parts.length >= 3 && parts[0] === 'SWC' ? parts[1] + '|' + parts[2] + '|' + parts[3] + '|' + parts[4] : '';
  }
  if (!caseId) return null;
  var state = swReadDataCleanupCaseState_(ss);
  return state.byId[caseId] || null;
}

function swDataCleanupCaseFromRow_(row) {
  var proposalText = row['Proposal JSON'] || '';
  return {
    rowNumber: row.__rowNumber || 0,
    caseId: row['CaseID'] || '',
    campaignId: row['Campaign ID'] || '',
    root: row['RootApptID'] || '',
    appt: row['APPT_ID'] || '',
    customerName: row['Customer Name'] || '',
    brand: row['Brand'] || '',
    clientAdvisor: row['Client Advisor'] || '',
    clientAdvisorEmail: swNormEmail_(row['Client Advisor Email'] || ''),
    joc: row['JOC'] || '',
    jocEmail: swNormEmail_(row['JOC Email'] || ''),
    stageKey: row['Stage Key'] || '',
    currentSalesStage: row['Current Sales Stage'] || '',
    currentConvStatus: row['Current Conversion Status'] || '',
    currentCustomOrder: row['Current Custom Order Status'] || '',
    currentInProduction: row['Current In Production Status'] || '',
    lastTouchAt: row['Last Touch At'] || '',
    staleDays: Number(row['Stale Days']) || 0,
    status: row['Status'] || '',
    campaignTab: row['Campaign Tab?'] || '',
    proposalJson: proposalText,
    proposal: swParseJson_(proposalText, {}),
    proposedBy: row['Proposed By'] || '',
    proposedByEmail: swNormEmail_(row['Proposed By Email'] || ''),
    proposedRole: row['Proposed Role'] || '',
    proposedAt: row['Proposed At'] || '',
    confirmationBy: row['Confirmation By'] || '',
    confirmationEmail: swNormEmail_(row['Confirmation Email'] || ''),
    confirmedAt: row['Confirmed At'] || '',
    returnReason: row['Return Reason'] || '',
    returnedBy: row['Returned By'] || '',
    returnedAt: row['Returned At'] || '',
    revisionCount: Number(row['Revision Count']) || 0,
    appliedAt: row['Applied At'] || '',
    appliedResultJson: row['Applied Result JSON'] || '',
    createdAt: row['Created At'] || '',
    updatedAt: row['Updated At'] || ''
  };
}

function swWriteDataCleanupCase_(ss, cleanupCase) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.DATA_CLEANUP, SW_DATA_CLEANUP_HEADERS);
  if (!cleanupCase.rowNumber) {
    var row = swFindDataCleanupCaseRow_(sh, cleanupCase.caseId);
    cleanupCase.rowNumber = row;
  }
  var values = swDataCleanupCaseToRow_(cleanupCase);
  if (cleanupCase.rowNumber) {
    sh.getRange(cleanupCase.rowNumber, 1, 1, SW_DATA_CLEANUP_HEADERS.length).setValues([values]);
    return cleanupCase.rowNumber;
  }
  sh.appendRow(values);
  cleanupCase.rowNumber = sh.getLastRow();
  return cleanupCase.rowNumber;
}

function swDataCleanupCaseToRow_(c) {
  var map = {
    'CaseID': c.caseId,
    'Campaign ID': c.campaignId,
    'RootApptID': c.root,
    'APPT_ID': c.appt,
    'Customer Name': c.customerName,
    'Brand': c.brand,
    'Client Advisor': c.clientAdvisor,
    'Client Advisor Email': c.clientAdvisorEmail,
    'JOC': c.joc,
    'JOC Email': c.jocEmail,
    'Stage Key': c.stageKey,
    'Current Sales Stage': c.currentSalesStage,
    'Current Conversion Status': c.currentConvStatus,
    'Current Custom Order Status': c.currentCustomOrder,
    'Current In Production Status': c.currentInProduction,
    'Last Touch At': c.lastTouchAt,
    'Stale Days': c.staleDays,
    'Status': c.status,
    'Campaign Tab?': c.campaignTab,
    'Proposal JSON': c.proposalJson,
    'Proposed By': c.proposedBy,
    'Proposed By Email': c.proposedByEmail,
    'Proposed Role': c.proposedRole,
    'Proposed At': c.proposedAt,
    'Confirmation By': c.confirmationBy,
    'Confirmation Email': c.confirmationEmail,
    'Confirmed At': c.confirmedAt,
    'Return Reason': c.returnReason,
    'Returned By': c.returnedBy,
    'Returned At': c.returnedAt,
    'Revision Count': c.revisionCount,
    'Applied At': c.appliedAt,
    'Applied Result JSON': c.appliedResultJson,
    'Created At': c.createdAt,
    'Updated At': c.updatedAt
  };
  return SW_DATA_CLEANUP_HEADERS.map(function (h) { return map[h] == null ? '' : map[h]; });
}

function swFindDataCleanupCaseRow_(sh, caseId) {
  if (!caseId || sh.getLastRow() < 2) return 0;
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(caseId)) return i + 2;
  }
  return 0;
}

function swDataCleanupCaseUnresolved_(cleanupCase) {
  return cleanupCase && cleanupCase.status !== SW_DATA_CLEANUP_STATUS.APPLIED;
}

function swDataCleanupHasUnresolvedCampaign_(caseState, campaignId) {
  var rows = (caseState && caseState.rows) || [];
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].campaignId === campaignId && rows[i].campaignTab === 'Y' && swDataCleanupCaseUnresolved_(rows[i])) return true;
  }
  return false;
}

function swDataCleanupPublicCase_(cleanupCase) {
  return {
    caseId: cleanupCase.caseId,
    campaignId: cleanupCase.campaignId,
    root: cleanupCase.root,
    status: cleanupCase.status,
    staleDays: cleanupCase.staleDays,
    lastTouchAt: cleanupCase.lastTouchAt,
    currentSalesStage: cleanupCase.currentSalesStage,
    currentConvStatus: cleanupCase.currentConvStatus,
    currentCustomOrder: cleanupCase.currentCustomOrder,
    currentInProduction: cleanupCase.currentInProduction,
    proposedBy: cleanupCase.proposedBy,
    proposedRole: cleanupCase.proposedRole,
    proposedAt: cleanupCase.proposedAt,
    proposal: swParseJson_(cleanupCase.proposalJson, cleanupCase.proposal || {}),
    returnReason: cleanupCase.returnReason,
    revisionCount: cleanupCase.revisionCount
  };
}

function swDataCleanupUsesDirectCampaignJocOwner_(cleanupCase, ownerRole) {
  return cleanupCase &&
    cleanupCase.campaignTab === 'Y' &&
    swWorkflowRoleMatches_(ownerRole, SW_OWNER_ROLES.JOC);
}

function swDataCleanupDirectCampaignJocOwner_(ss, ctx, rec) {
  rec = rec || {};
  var intendedName = swTrim_(rec.assistedRep || '');
  var intendedEmail = swNormEmail_(rec.assistedRepEmail || '');
  var owner = swDataCleanupCanonicalJocOwner_(ss, ctx, intendedName, intendedEmail);
  if (owner) {
    return {
      intendedOwner: owner.name,
      intendedOwnerEmail: owner.email,
      currentOwner: owner.name,
      currentOwnerEmail: owner.email,
      coverageReason: ''
    };
  }
  if (intendedName || intendedEmail) {
    return {
      intendedOwner: intendedName || intendedEmail,
      intendedOwnerEmail: intendedEmail,
      currentOwner: intendedName || intendedEmail,
      currentOwnerEmail: intendedEmail,
      coverageReason: 'UNRESOLVED_ASSISTED_REP'
    };
  }
  return {
    intendedOwner: '',
    intendedOwnerEmail: '',
    currentOwner: 'Admin Review',
    currentOwnerEmail: '',
    coverageReason: 'NO_ASSISTED_REP'
  };
}

function swDataCleanupCanonicalJocOwner_(ss, ctx, name, email) {
  var exact = swCanonicalWorkflowOwnerForRole_(ss, ctx, name, email, SW_OWNER_ROLES.JOC);
  if (exact) return exact;
  var parts = swDataCleanupOwnerNameParts_(name);
  var matched = [];
  parts.forEach(function (part) {
    var owner = swCanonicalWorkflowOwnerForRole_(ss, ctx, part, '', SW_OWNER_ROLES.JOC);
    if (!owner) return;
    for (var i = 0; i < matched.length; i++) {
      if (swNormEmail_(matched[i].email) === swNormEmail_(owner.email)) return;
    }
    matched.push(owner);
  });
  return matched.length ? matched[0] : null;
}

function swDataCleanupOwnerNameParts_(name) {
  return String(name || '')
    .split(/\s*(?:,|\/|&|\band\b)\s*/i)
    .map(function (part) { return swTrim_(part); })
    .filter(Boolean);
}

function swDataCleanupRestoreCampaignJocTasks_(ss, taskState, ctx, caseState, campaignId, now) {
  var byId = (taskState && taskState.byId) || {};
  var casesById = (caseState && caseState.byId) || {};
  var count = 0;
  Object.keys(byId).forEach(function (taskId) {
    var task = byId[taskId];
    if (!swDataCleanupShouldRestoreCampaignJocTask_(task, campaignId)) return;
    var cleanupCase = casesById[swDataCleanupCaseIdFromTask_(task)] || {};
    var payload = swParseJson_(task.payloadJson, {});
    var rec = swDataCleanupRecFromTask_(task, payload, cleanupCase);
    rec.assistedRep = rec.assistedRep || task.intendedOwner || '';
    rec.assistedRepEmail = rec.assistedRepEmail || task.intendedOwnerEmail || '';
    var owner = swDataCleanupDirectCampaignJocOwner_(ss, ctx, rec);
    if (swNorm_(task.currentOwner) === swNorm_(owner.currentOwner) &&
        swNormEmail_(task.currentOwnerEmail) === swNormEmail_(owner.currentOwnerEmail) &&
        swNorm_(task.intendedOwner) === swNorm_(owner.intendedOwner) &&
        swNormEmail_(task.intendedOwnerEmail) === swNormEmail_(owner.intendedOwnerEmail) &&
        task.coverageReason === owner.coverageReason) {
      return;
    }

    var fromOwner = task.currentOwner;
    task.intendedOwner = owner.intendedOwner;
    task.intendedOwnerEmail = owner.intendedOwnerEmail;
    task.currentOwner = owner.currentOwner;
    task.currentOwnerEmail = owner.currentOwnerEmail;
    task.coverageReason = owner.coverageReason;
    task.updatedAt = swIso_(now || new Date());
    task.lastEvent = 'ASSIGN';
    swWriteTaskRow_(ss, task);
    swAppendTaskLog_(ss, 'ASSIGN', task, swSystemUser_(), fromOwner, task.currentOwner, {
      reason: 'One-time data cleanup JOC work stays in the Cleanup tab.',
      coverageReason: task.coverageReason
    });
    count++;
  });
  return count;
}

function swDataCleanupShouldRestoreCampaignJocTask_(task, campaignId) {
  if (!task || task.claimedBy) return false;
  if (task.status !== SW_STATUSES.PENDING && task.status !== SW_STATUSES.SNOOZED) return false;
  if (!swWorkflowRoleMatches_(task.ownerRole, SW_OWNER_ROLES.JOC)) return false;
  if (!swIsDataCleanupTaskType_(task.taskType)) return false;
  if (swNorm_(task.lifecycleStage) !== swNorm_('Cleanup Campaign')) return false;
  return swDataCleanupCampaignIdFromTask_(task) === campaignId;
}

function swDataCleanupInactiveTaskMessage_(ss, task) {
  if (!task || !swIsCleanupCampaignTask_(task)) return '';
  var cleanupCase = swReadDataCleanupCaseByTask_(ss, task);
  if (cleanupCase && cleanupCase.status === SW_DATA_CLEANUP_STATUS.PENDING_CONFIRMATION) {
    return 'This cleanup case already has a submitted proposal. Refresh Queue and open the confirmation task.';
  }
  if (cleanupCase && cleanupCase.status === SW_DATA_CLEANUP_STATUS.RETURNED) {
    return 'This cleanup case was returned for revision. Refresh Queue and open the revision task.';
  }
  if (cleanupCase && cleanupCase.status === SW_DATA_CLEANUP_STATUS.APPLIED) {
    return 'This cleanup case has already been applied. Refresh Queue to clear it from the list.';
  }
  if (task.status === SW_STATUSES.COMPLETED) {
    return 'This cleanup task was already completed. Refresh Queue to load the current cleanup list.';
  }
  if (task.status === SW_STATUSES.BLOCKED) {
    return 'This cleanup task is no longer active' + (task.coverageReason ? ' (' + task.coverageReason + ')' : '') + '. Refresh Queue to load the current cleanup task.';
  }
  if (task.status === SW_STATUSES.SNOOZED) {
    return 'This cleanup task is snoozed until later. Refresh Queue to load currently due cleanup work.';
  }
  return 'This cleanup task is no longer active. Refresh Queue to load the current cleanup task.';
}

function swDataCleanupCampaignIdFromTask_(task) {
  var payload = swParseJson_(task && task.payloadJson, {});
  var campaignId = swDeepValue_(payload, ['extra', 'cleanupCase', 'campaignId']);
  if (campaignId) return campaignId;
  var parts = String(task && task.taskId || '').split('|');
  return parts.length >= 3 && parts[0] === 'SWC' && parts[1] === 'DC' ? parts[2] : '';
}

function swDataCleanupCaseIdFromTask_(task) {
  var payload = swParseJson_(task && task.payloadJson, {});
  var caseId = swDeepValue_(payload, ['extra', 'cleanupCase', 'caseId']);
  if (caseId) return caseId;
  var parts = String(task && task.taskId || '').split('|');
  return parts.length >= 5 && parts[0] === 'SWC'
    ? [parts[1], parts[2], parts[3], parts[4]].join('|')
    : '';
}

function swDataCleanupUpsertImmediateTask_(ss, task, actor, eventType) {
  var existing = swGetTaskById_(ss, task.taskId);
  if (existing && existing.status !== SW_STATUSES.COMPLETED) {
    task.rowNumber = existing.rowNumber;
    task.createdAt = existing.createdAt || task.createdAt;
    swWriteTaskRow_(ss, task);
    swAppendTaskLog_(ss, eventType || 'UPDATE', task, actor || swSystemUser_(), existing.currentOwner, task.currentOwner, {});
    return;
  }
  swAppendTaskRow_(ss, task);
  swAppendTaskLog_(ss, eventType || 'CREATE', task, actor || swSystemUser_(), '', task.currentOwner, {});
}

function swDataCleanupBlockTasksForCase_(ss, caseId, exceptTaskId, actor, reason, predicate) {
  var state = swReadTaskState_(ss);
  Object.keys(state.byId || {}).forEach(function (taskId) {
    var t = state.byId[taskId];
    if (taskId === exceptTaskId) return;
    if (String(taskId).indexOf('SWC|' + caseId + '|') !== 0) return;
    if (predicate && !predicate(t)) return;
    if (t.status === SW_STATUSES.COMPLETED || t.status === SW_STATUSES.BLOCKED) return;
    t.status = SW_STATUSES.BLOCKED;
    t.coverageReason = reason || 'DATA_CLEANUP_SUPERSEDED';
    t.updatedAt = swIso_(new Date());
    t.lastEvent = 'BLOCK';
    swWriteTaskRow_(ss, t);
    swAppendTaskLog_(ss, 'BLOCK', t, actor || swSystemUser_(), t.currentOwner, t.currentOwner, { reason: reason || '' });
  });
}

function swDataCleanupRowsByRoot_(appointments) {
  var groups = {};
  (appointments || []).forEach(function (rec) {
    var root = swTrim_(rec.root || rec.appt);
    if (!root) return;
    if (!groups[root]) groups[root] = [];
    groups[root].push(rec);
  });
  return groups;
}

function swDataCleanupCurrentRowsByRoot_(groups) {
  var out = {};
  Object.keys(groups || {}).forEach(function (root) {
    var rows = groups[root] || [];
    var active = rows.filter(function (rec) { return swIsAppointmentActive_(rec); });
    out[root] = swDataCleanupLatestRow_(active.length ? active : rows);
  });
  return out;
}

function swDataCleanupLatestRow_(rows) {
  var tz = swTimezone_();
  return (rows || []).reduce(function (best, rec) {
    if (!best) return rec;
    var bv = swDataCleanupRowSortValue_(best, tz);
    var rv = swDataCleanupRowSortValue_(rec, tz);
    return rv >= bv ? rec : best;
  }, null);
}

function swDataCleanupRowSortValue_(rec, tz) {
  var visit = swVisitDateTime_(rec, tz);
  if (visit) return visit.getTime();
  var updated = swDataCleanupDateTimeValue_(rec.updatedAtRaw, rec.updatedAt);
  if (updated) return updated.getTime();
  return Number(rec.row || 0);
}

function swDataCleanupPipelineStage_(rec, rootRows) {
  var sales = swNorm_(rec.salesStage);
  var conv = swNorm_(rec.convStatus);
  var custom = swNorm_(rec.customOrder);
  var inProd = swNorm_(rec.inProduction);
  var combined = [sales, conv, custom, inProd].join(' ');
  if (/lost/.test(combined)) return { key: 'lost', label: 'Lost Lead' };
  if (/won/.test(combined) || /order completed/.test(custom) || /production completed/.test(inProd)) return { key: 'won', label: 'Won / Completed' };
  if (custom === 'in production' || (inProd && !/none|n\/a|na/.test(inProd))) return { key: 'inProduction', label: 'In Production' };
  if (/deposit|confirmed order|order in progress|approved for production|waiting production|3d requested|3d revision|3d received/.test(combined)) return { key: 'deposit', label: 'Deposit / Order In Progress' };
  if (/appointment|viewing scheduled|scheduled/.test(combined) || swDataCleanupHasFutureVisit_(rootRows)) return { key: 'appointment', label: 'Appointment / Viewing Scheduled' };
  if (/follow/.test(combined)) return { key: 'followUp', label: 'Follow-Up' };
  if (/hot/.test(combined)) return { key: 'hotLead', label: 'Hot Lead' };
  return { key: 'lead', label: 'Lead' };
}

function swDataCleanupStageNeedsReview_(stageKey) {
  return stageKey === 'lead' || stageKey === 'hotLead' || stageKey === 'followUp';
}

function swDataCleanupHasFutureVisit_(rows) {
  var tz = swTimezone_();
  var today = swDataCleanupStartOfDay_(new Date()).getTime();
  for (var i = 0; i < (rows || []).length; i++) {
    if (!swIsAppointmentActive_(rows[i])) continue;
    var visit = swVisitDateTime_(rows[i], tz);
    if (visit && visit.getTime() >= today) return true;
  }
  return false;
}

function swDataCleanupLastTouch_(rec, rootRows, rootIndexTouch, now) {
  var candidates = [];
  if (rootIndexTouch) candidates.push(rootIndexTouch);
  var updated = swDataCleanupDateTimeValue_(rec.updatedAtRaw, rec.updatedAt);
  if (updated) candidates.push(updated);
  var paid = swDataCleanupDateTimeValue_(rec.lastPaymentDateRaw, rec.lastPaymentDate);
  if (paid) candidates.push(paid);
  (rootRows || []).forEach(function (row) {
    var visit = swVisitDateTime_(row, swTimezone_());
    if (visit && visit.getTime() <= now.getTime()) candidates.push(visit);
    var booked = swDataCleanupDateTimeValue_(row.bookedAtRaw, row.bookedAt);
    if (booked) candidates.push(booked);
  });
  return candidates.reduce(function (best, item) {
    if (!item || isNaN(item.getTime())) return best;
    if (!best || item.getTime() > best.getTime()) return item;
    return best;
  }, null);
}

function swDataCleanupReadRootIndex_(ss) {
  var out = { byRoot: {} };
  var sh = ss.getSheetByName('07_Root_Index');
  if (!sh || sh.getLastRow() < 2) return out;
  var values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var display = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  var H = swHeaderMapFromArray_(display[0].map(function (h) { return swTrim_(h); }));
  var cRoot = swPickIndex_(H, ['RootApptID', 'Root Appt ID', 'ROOT', 'Root_ID']);
  var cUpdated = swPickIndex_(H, ['Updated At', 'UpdatedAt', 'Last Updated']);
  if (cRoot < 0 || cUpdated < 0) return out;
  for (var i = 1; i < values.length; i++) {
    var root = swTrim_(swCell_(display[i], cRoot));
    var when = swDataCleanupDateTimeValue_(swCell_(values[i], cUpdated), swCell_(display[i], cUpdated));
    if (!root || !when) continue;
    if (!out.byRoot[root] || when.getTime() > out.byRoot[root].getTime()) out.byRoot[root] = when;
  }
  return out;
}

function swDataCleanupDateTimeValue_(raw, display) {
  if (raw instanceof Date && !isNaN(raw.getTime())) return raw;
  var s = swTrim_(display || raw);
  if (!s) return null;
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  var parts = swDateParts_(raw, display);
  return parts ? new Date(parts.y, parts.m - 1, parts.d, 0, 0, 0, 0) : null;
}

function swDataCleanupStartOfDay_(date) {
  date = date || new Date();
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
}

function swDataCleanupCaseId_(campaignId, root, lastTouch) {
  var touchKey = lastTouch ? Utilities.formatDate(lastTouch, swTimezone_(), 'yyyyMMdd') : 'notouch';
  return ['DC', campaignId, root, touchKey].join('|');
}

function swDataCleanupTaskId_(caseId, phase, ownerKey, revision) {
  return ['SWC', caseId, phase, ownerKey, revision || 0].join('|');
}

function swDataCleanupRoleKey_(role) {
  return swWorkflowRoleMatches_(role, SW_OWNER_ROLES.SALES_REP) ? 'REP' : 'JOC';
}

function swDataCleanupOppositeRole_(role) {
  return swWorkflowRoleMatches_(role, SW_OWNER_ROLES.SALES_REP) ? SW_OWNER_ROLES.JOC : SW_OWNER_ROLES.SALES_REP;
}

function swDataCleanupRecFromTask_(task, payload, cleanupCase) {
  payload = payload || {};
  var appt = payload.appointment || {};
  return {
    row: appt.row || '',
    root: appt.root || cleanupCase.root || '',
    appt: appt.appt || cleanupCase.appt || '',
    uid: appt.uid || '',
    name: appt.customerName || cleanupCase.customerName || '',
    email: appt.email || '',
    phone: appt.phone || '',
    brand: appt.brand || cleanupCase.brand || '',
    visitDate: appt.visitDate || task.visitDate || '',
    visitTime: appt.visitTime || task.visitTime || '',
    visitType: appt.visitType || task.visitType || '',
    assignedRep: appt.assignedRep || cleanupCase.clientAdvisor || '',
    assignedRepEmail: appt.assignedRepEmail || cleanupCase.clientAdvisorEmail || '',
    assistedRep: appt.assistedRep || cleanupCase.joc || '',
    assistedRepEmail: appt.assistedRepEmail || cleanupCase.jocEmail || '',
    clientFolder: appt.clientFolder || '',
    reportUrl: appt.reportUrl || '',
    quotationUrl: appt.quotationUrl || '',
    tracker3dUrl: appt.tracker3dUrl || '',
    salesStage: appt.salesStage || cleanupCase.currentSalesStage || '',
    convStatus: appt.convStatus || cleanupCase.currentConvStatus || '',
    customOrder: appt.customOrder || cleanupCase.currentCustomOrder || '',
    inProduction: appt.inProduction || cleanupCase.currentInProduction || '',
    centerStoneStatus: appt.centerStoneStatus || '',
    nextSteps: appt.nextSteps || '',
    designRequest: appt.designRequest || '',
    deadline3d: appt.deadline3d || '',
    productionDeadline: appt.productionDeadline || '',
    waxStatus: appt.waxStatus || '',
    waxDeadlineAdmin: appt.waxDeadlineAdmin || '',
    waxRequestUrl: appt.waxRequestUrl || '',
    dvStonesSummary: appt.dvStonesSummary || '',
    so: appt.so || '',
    orderFolder: appt.orderFolder || '',
    remainingBalance: appt.remainingBalance || '',
    orderTotal: appt.orderTotal || '',
    paidToDate: appt.paidToDate || '',
    lastPaymentDate: appt.lastPaymentDate || '',
    orderDate: appt.orderDate || ''
  };
}

function swDataCleanupSetConfigValue_(ss, key, value) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.CONFIG, SW_CONFIG_HEADERS);
  var rows = swReadSheetObjectsExpectedHeaders_(sh, SW_CONFIG_HEADERS);
  for (var i = 0; i < rows.length; i++) {
    if (swNorm_(rows[i]['Section']) === 'system' && swNorm_(rows[i]['Key']) === swNorm_(key)) {
      sh.getRange(rows[i].__rowNumber, 3).setValue(value);
      return;
    }
  }
  sh.appendRow(['SYSTEM', key, value, '', '', '', 'Y', '', 'Set by data cleanup workflow.']);
}
