/**
 * Sales workflow post-consult operations: client status, Start 3D, deadlines,
 * wax request/update tasks, and dashboard row-based completion adapters.
 */

function swGeneratePostConsultTasks_(ss, state, ctx, rec, now, summary) {
  var checklistId = swTaskId_(rec, SW_TASKS.CHECKLIST);
  if (!swTaskCompleted_(state, checklistId)) return;

  var statusId = swTaskId_(rec, SW_TASKS.POST_CONSULT_STATUS);
  var checklistPayload = swTaskPayload_(state, checklistId);
  swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.POST_CONSULT_STATUS, SW_OWNER_ROLES.JOC, now, checklistId, now, {
    consultHandoff: swDeepValue_(checklistPayload, ['completion']) || {}
  }), summary);

  if (!swTaskCompleted_(state, statusId)) return;

  var statusCompletion = swPostConsultCompletion_(state, rec, SW_TASKS.POST_CONSULT_STATUS);
  var no3d = swPostConsultNo3D_(state, rec);
  var startId = swTaskId_(rec, SW_TASKS.START_3D);
  var startCompleted = swTaskCompleted_(state, startId);
  var hasStarted3d = !!(rec.so || rec.tracker3dUrl || startCompleted);

  if (!no3d && !rec.so && !startCompleted) {
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.START_3D, SW_OWNER_ROLES.JOC, now, statusId, now, {
      clientStatus: statusCompletion || {},
      designRequest: statusCompletion.designRequest || rec.designRequest || '',
      nextSteps: statusCompletion.nextSteps || rec.nextSteps || '',
      waxNeeded: statusCompletion.waxNeeded || ''
    }), summary);
  }

  if (!no3d && hasStarted3d && !rec.deadline3d) {
    var deadlineDependency = startCompleted ? startId : statusId;
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.RECORD_3D_DEADLINE, SW_OWNER_ROLES.JOC, swNextDayAt930_(now), deadlineDependency, now, {
      soNumber: rec.so || swDeepValue_(swTaskPayload_(state, startId), ['completion', 'so']) || '',
      tracker3dUrl: rec.tracker3dUrl || ''
    }), summary);
  }

  var wantsWax = swPostConsultWantsWax_(state, rec);
  var waxActive = swPostConsultActiveWaxRequests_(ctx, rec.root);
  if (wantsWax && !waxActive.length && !swTaskCompleted_(state, swTaskId_(rec, SW_TASKS.REQUEST_WAX))) {
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.REQUEST_WAX, SW_OWNER_ROLES.JOC, now, startCompleted ? startId : statusId, now, {
      waxNeededBy: statusCompletion.waxNeededBy || swDeepValue_(swTaskPayload_(state, startId), ['completion', 'waxNeededBy']) || '',
      waxPriority: statusCompletion.waxPriority || swDeepValue_(swTaskPayload_(state, startId), ['completion', 'waxPriority']) || '',
      soNumber: rec.so || swDeepValue_(swTaskPayload_(state, startId), ['completion', 'so']) || ''
    }), summary);
  }

  var waxNeedsUpdate = swPostConsultWaxRequestsNeedingUpdate_(ctx, rec.root);
  if (waxNeedsUpdate.length) {
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.UPDATE_WAX, SW_OWNER_ROLES.JOC, now, '', now, {
      waxRequests: waxNeedsUpdate,
      waxRequestSummary: swWaxRequestSummary_(waxNeedsUpdate),
      waxRequestUrl: waxNeedsUpdate[0].rowUrl || waxNeedsUpdate[0].link || ''
    }), summary);
  }
}

function swPostConsultCompletion_(state, rec, taskType) {
  var payload = swTaskPayload_(state, swTaskId_(rec, taskType));
  return (payload && payload.completion) || {};
}

function swPostConsultNo3D_(state, rec) {
  var status = swPostConsultCompletion_(state, rec, SW_TASKS.POST_CONSULT_STATUS);
  if (swNorm_(status.threeDNeeded) === 'no') return true;
  var start = swPostConsultCompletion_(state, rec, SW_TASKS.START_3D);
  return !!start.no3d || swNorm_(start.start3dMode) === 'no3d';
}

function swPostConsultWantsWax_(state, rec) {
  var status = swPostConsultCompletion_(state, rec, SW_TASKS.POST_CONSULT_STATUS);
  var start = swPostConsultCompletion_(state, rec, SW_TASKS.START_3D);
  return swNorm_(status.waxNeeded) === 'yes' || swNorm_(start.waxNeeded) === 'yes';
}

function swPostConsultActiveWaxRequests_(ctx, root) {
  return ((ctx && ctx.waxIndex && ctx.waxIndex.activeByRoot) || {})[root] || [];
}

function swPostConsultWaxRequestsNeedingUpdate_(ctx, root) {
  return ((ctx && ctx.waxIndex && ctx.waxIndex.needsUpdateByRoot) || {})[root] || [];
}

function swWaxRequestSummary_(items) {
  return (items || []).map(function (w) {
    return [w.id || 'Wax request', w.status || 'No status', w.adminDeadline ? ('Admin due ' + w.adminDeadline) : 'No admin deadline'].join(' - ');
  }).join('\n');
}

function swNextDayAt930_(date) {
  date = date || new Date();
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1, 9, 30, 0, 0);
}

function swIsPostConsultTaskType_(taskType) {
  return [
    SW_TASKS.POST_CONSULT_STATUS,
    SW_TASKS.START_3D,
    SW_TASKS.RECORD_3D_DEADLINE,
    SW_TASKS.REQUEST_WAX,
    SW_TASKS.UPDATE_WAX
  ].indexOf(taskType) >= 0;
}

function swValidatePostConsultCompletion_(task, data) {
  if (!swIsPostConsultTaskType_(task.taskType)) return;
  data = data || {};
  if (task.taskType === SW_TASKS.POST_CONSULT_STATUS) {
    if (!swTrim_(data.salesStage)) throw new Error('Select Sales Stage before submitting Client Status.');
    if (!swTrim_(data.convStatus)) throw new Error('Select Conversion Status before submitting Client Status.');
    if (!swTrim_(data.threeDNeeded)) throw new Error('Select whether 3D is needed.');
    if (swNorm_(data.threeDNeeded) === 'no' && !swTrim_(data.no3dReason)) {
      throw new Error('Enter a reason when marking 3D not needed.');
    }
  }
  if (task.taskType === SW_TASKS.START_3D) {
    if (data.no3d || swNorm_(data.start3dMode) === 'no3d') {
      if (!swTrim_(data.no3dReason)) throw new Error('Enter a reason when marking 3D not needed.');
      return;
    }
    if (!swTrim_(data.brand)) throw new Error('Select a brand before starting 3D.');
    if (!swTrim_(data.so)) throw new Error('Enter the SO number before starting 3D.');
    if (!swTrim_(data.odooUrl)) throw new Error('Enter the Odoo SO URL before starting 3D.');
  }
  if (task.taskType === SW_TASKS.RECORD_3D_DEADLINE) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(swTrim_(data.deadline3d))) throw new Error('Select a valid 3D deadline.');
  }
  if (task.taskType === SW_TASKS.REQUEST_WAX) {
    if (!swTrim_(data.soMo)) throw new Error('Enter the SO/MO number for the wax request.');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(swTrim_(data.neededByRep))) throw new Error('Select Needed By (Rep) for the wax request.');
    if (!swTrim_(data.priority)) throw new Error('Select wax priority.');
  }
  if (task.taskType === SW_TASKS.UPDATE_WAX) {
    var updates = data.waxUpdates || [];
    if (!updates.length) throw new Error('Enter at least one wax update.');
    updates.forEach(function (u) {
      if (!swTrim_(u.id)) throw new Error('Wax update is missing a request ID.');
      if (!swTrim_(u.status)) throw new Error('Select a wax status for ' + u.id + '.');
      var done = /(completed|canceled|cancelled)/i.test(swTrim_(u.status));
      if (!done && !swTrim_(u.adminDeadline)) throw new Error('Enter an admin deadline for ' + u.id + '.');
    });
  }
}

function swHandlePostConsultTaskCompletion_(ss, task, data, user) {
  if (!swIsPostConsultTaskType_(task.taskType)) return null;
  if (task.taskType === SW_TASKS.POST_CONSULT_STATUS) return swCompleteClientStatusTask_(ss, task, data, user);
  if (task.taskType === SW_TASKS.START_3D) return swCompleteStart3DTask_(ss, task, data, user);
  if (task.taskType === SW_TASKS.RECORD_3D_DEADLINE) return swComplete3DDeadlineTask_(ss, task, data, user);
  if (task.taskType === SW_TASKS.REQUEST_WAX) return swCompleteWaxRequestTask_(ss, task, data, user);
  if (task.taskType === SW_TASKS.UPDATE_WAX) return swCompleteWaxUpdateTask_(ss, task, data, user);
  return null;
}

function swCompleteClientStatusTask_(ss, task, data, user) {
  var payload = swParseJson_(task.payloadJson, {});
  var appt = payload.appointment || {};
  var rowNum = swSetMasterActiveRowForTask_(ss, task);
  var result = cs_submitFromDialogForRow_(rowNum, {
    assignedRep: appt.assignedRep || '',
    assistedRep: appt.assistedRep || '',
    salesStage: swTrim_(data.salesStage),
    convStatus: swTrim_(data.convStatus),
    customOrder: swTrim_(data.customOrder),
    cosAllowedEmpty: !swTrim_(data.customOrder),
    inProduction: swTrim_(data.inProduction),
    centerStone: swTrim_(data.centerStone),
    nextSteps: swTrim_(data.nextSteps),
    orderDate: swTrim_(data.orderDate),
    deadline3d: swTrim_(data.deadline3d),
    prodDeadline: swTrim_(data.prodDeadline),
    wax: null,
    waxSummary: '',
    notebookLMLink: swTrim_(data.notebookLMLink)
  }, ss);
  if (result && result.ok === false) throw new Error(result.error || 'Client Status update failed.');
  return { action: 'CLIENT_STATUS_SUBMITTED', result: result && result.summary ? result.summary : result };
}

function swCompleteStart3DTask_(ss, task, data, user) {
  if (data.no3d || swNorm_(data.start3dMode) === 'no3d') {
    return { action: 'NO_3D_NEEDED', reason: swTrim_(data.no3dReason) };
  }
  swSetMasterActiveRowForTask_(ss, task);
  var result = saveAssignedSO({
    brand: swTrim_(data.brand),
    so: swTrim_(data.so),
    odooUrl: swTrim_(data.odooUrl),
    designRequest: swTrim_(data.designRequest),
    shortTag: swTrim_(data.shortTag),
    forceOverwrite: !!data.forceOverwrite,
    designForm: {
      Shape: swTrim_(data.shape),
      RingStyle: swTrim_(data.ringStyle),
      Notes: swTrim_(data.designRequest)
    }
  });
  if (result && result.ok === false) throw new Error(result.error || 'Start 3D failed.');
  return { action: 'START_3D_SUBMITTED', result: result && result.summary ? result.summary : result };
}

function swComplete3DDeadlineTask_(ss, task, data, user) {
  swSetMasterActiveRowForTask_(ss, task);
  var result = (typeof Deadlines !== 'undefined' && Deadlines.saveRecordDeadline)
    ? Deadlines.saveRecordDeadline({ kind: '3D', dateIso: swTrim_(data.deadline3d) })
    : saveRecordDeadline({ kind: '3D', dateIso: swTrim_(data.deadline3d) });
  if (result && result.ok === false) throw new Error(result.error || '3D deadline save failed.');
  return { action: '3D_DEADLINE_RECORDED', result: result };
}

function swCompleteWaxRequestTask_(ss, task, data, user) {
  var payload = swParseJson_(task.payloadJson, {});
  var appt = payload.appointment || {};
  var root = swTrim_(appt.root || task.root || task.appt);
  var result = wax_onRequestSubmit_({
    rootApptId: root,
    soMo: swTrim_(data.soMo),
    neededByRep: swTrim_(data.neededByRep),
    priority: swTrim_(data.priority),
    requestedBy: (user && (user.email || user.name)) || ''
  });
  if (result && result.ok === false) throw new Error(result.error || 'Wax request failed.');
  return { action: 'WAX_REQUEST_CREATED', result: result };
}

function swCompleteWaxUpdateTask_(ss, task, data, user) {
  var result = wax_adminCommitFromDialog({ updates: data.waxUpdates || [] });
  if (result && result.ok === false) throw new Error(result.error || 'Wax update failed.');
  return { action: 'WAX_REQUEST_UPDATED', result: result };
}

function swSetMasterActiveRowForTask_(ss, task) {
  var row = swMasterRowForTask_(ss, task);
  var sh = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!sh || !row) throw new Error('Could not resolve Master row for this task.');
  ss.setActiveSheet(sh);
  ss.setActiveRange(sh.getRange(row, 1));
  return row;
}

function swMasterRowForTask_(ss, task) {
  var payload = swParseJson_(task.payloadJson, {});
  var row = Number(swDeepValue_(payload, ['appointment', 'row']) || 0);
  if (row > 1) return row;
  var sh = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!sh || sh.getLastRow() < 2) return 0;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  var H = swHeaderMapFromArray_(headers);
  var cRoot = swPickIndex_(H, ['RootApptID', 'APPT_ID']);
  if (cRoot < 0) return 0;
  var values = sh.getRange(2, cRoot + 1, sh.getLastRow() - 1, 1).getDisplayValues();
  var want = swTrim_(task.root || task.appt);
  for (var i = 0; i < values.length; i++) {
    if (swTrim_(values[i][0]) === want) return i + 2;
  }
  return 0;
}

var SW_TASK_FORM_OPTIONS_MEMORY_CACHE_ = {};

function swTaskFormOptions_(ss, task) {
  var taskType = task && task.taskType ? task.taskType : '';
  var cacheKey = '';
  if (taskType) {
    try {
      var cacheTaskType = taskType;
      if (typeof swIsDataCleanupTaskType_ === 'function' && swIsDataCleanupTaskType_(taskType)) {
        cacheTaskType = 'DATA_CLEANUP';
      }
      cacheKey = 'sw:formOptions:v1:' + ss.getId() + ':' + cacheTaskType;
      var memory = SW_TASK_FORM_OPTIONS_MEMORY_CACHE_[cacheKey];
      if (memory && memory.expiresAt > new Date().getTime()) return memory.value || {};
      var cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) {
        var cachedValue = swParseJson_(cached, {});
        SW_TASK_FORM_OPTIONS_MEMORY_CACHE_[cacheKey] = {
          expiresAt: new Date().getTime() + 10 * 60 * 1000,
          value: cachedValue
        };
        return cachedValue;
      }
    } catch (_) {}
  }
  var out;
  if (task && typeof swIsDataCleanupTaskType_ === 'function' && swIsDataCleanupTaskType_(task.taskType)) {
    out = swDataCleanupFormOptions_(ss, task);
    swCacheTaskFormOptions_(cacheKey, out);
    return out;
  }
  if (!task || !swIsPostConsultTaskType_(task.taskType)) return {};
  out = {
    salesStages: ['Lead', 'Follow-Up Required', 'Viewing Scheduled', 'Order In Progress', 'Lost Lead'],
    convStatuses: ['Quotation Requested', 'Viewing Scheduled', 'Deposit Paid', 'Confirmed Order', 'Order In Progress', 'Lost Lead'],
    customOrderStatuses: ['', '3D Requested', '3D Revision Requested', '3D Received', 'Approved for Production', 'Waiting Production Timeline', 'In Production', 'Order Completed'],
    inProductionStatuses: ['', 'CAD Approved', 'Casting', 'Setting', 'QC', 'Production Completed'],
    centerStoneStatuses: ['', 'No Center Stone', 'Need to Propose', 'Viewing Scheduled', 'Ordered', 'In Stock', 'Customer Approved'],
    waxStatuses: []
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
  try {
    if (typeof wax_statusOptions === 'function') out.waxStatuses = wax_statusOptions();
  } catch (_) {}
  if (!out.waxStatuses.length) out.waxStatuses = ['Wax Requested', 'In Progress', 'Completed', 'Canceled'];
  swCacheTaskFormOptions_(cacheKey, out);
  return out;
}

function swCacheTaskFormOptions_(cacheKey, out) {
  if (!cacheKey) return;
  try {
    SW_TASK_FORM_OPTIONS_MEMORY_CACHE_[cacheKey] = {
      expiresAt: new Date().getTime() + 10 * 60 * 1000,
      value: out || {}
    };
  } catch (_) {}
  try {
    var json = JSON.stringify(out || {});
    if (json.length <= 90000) CacheService.getScriptCache().put(cacheKey, json, 10 * 60);
  } catch (_) {}
}
