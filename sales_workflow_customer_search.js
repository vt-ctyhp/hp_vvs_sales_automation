/**
 * Sales Workflow customer search: dashboard kanban lookup and direct ops actions.
 */

var SW_CUSTOMER_SEARCH_MAX_CARDS_PER_COLUMN = 30;
var SW_CUSTOMER_SEARCH_LOG_SCAN_ROWS = 500;

function sw_searchCustomers(authToken, query, filters) {
  return swTimed_('sw_searchCustomers', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken);
    swRequireCustomerSearchUser_(user);

    filters = swCustomerSearchNormalizeFilters_(filters);
    query = swTrim_(query || filters.query || '');

    var appointments = swReadAppointments_(ss);
    var rows = swCustomerSearchFilteredRows_(appointments, query, filters);
    var groups = swAdminDashboardRowsByRoot_(rows);
    var master = ss.getSheetByName(SW_SHEETS.MASTER);
    var masterGid = master ? master.getSheetId() : '';
    var columnsByKey = {};
    SW_ADMIN_DASHBOARD_COLUMNS.forEach(function (col) {
      columnsByKey[col.key] = { key: col.key, label: col.label, count: 0, cards: [], hiddenCount: 0 };
    });

    Object.keys(groups).forEach(function (root) {
      var rootRows = groups[root];
      var active = rootRows.filter(function (rec) { return swIsAppointmentActive_(rec); });
      var rec = swAdminDashboardLatestRow_(active.length ? active : rootRows);
      if (!rec) return;
      var stage = swAdminDashboardPipelineStage_(rec, rootRows);
      var column = columnsByKey[stage.key] || columnsByKey.lead;
      var card = swCustomerSearchCard_(ss, masterGid, root, rec, rootRows, stage);
      column.count++;
      if (column.cards.length < SW_CUSTOMER_SEARCH_MAX_CARDS_PER_COLUMN) {
        column.cards.push(card);
      } else {
        column.hiddenCount++;
      }
    });

    return {
      ok: true,
      generatedAt: swIso_(new Date()),
      query: query,
      filters: swCustomerSearchPublicFilters_(filters),
      filterOptions: swCustomerSearchFilterOptions_(appointments),
      kanban: {
        columns: SW_ADMIN_DASHBOARD_COLUMNS.map(function (col) { return columnsByKey[col.key]; })
      }
    };
  });
}

function sw_getCustomerSearchDetail(authToken, rootApptId) {
  return swTimed_('sw_getCustomerSearchDetail', function () {
    var mark = swStepTimer_('sw_getCustomerSearchDetail');
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    mark('requiredSheets');
    var user = swAuthUserForApi_(ss, authToken);
    mark('identity');
    swRequireCustomerSearchUser_(user);
    var out = swCustomerSearchDetailPayload_(ss, user, rootApptId);
    mark('payload', {
      appointments: out.appointments ? out.appointments.length : 0,
      tasks: out.tasks ? out.tasks.length : 0,
      logs: out.logs ? out.logs.length : 0
    });
    return out;
  });
}

function sw_customerSearchUpdateStatus(authToken, rootApptId, payload) {
  return swTimed_('sw_customerSearchUpdateStatus', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken);
    swRequireCustomerSearchUser_(user);
    payload = payload || {};

    var lock = LockService.getDocumentLock() || LockService.getScriptLock();
    lock.waitLock(28000);
    try {
      var target = swCustomerSearchActivateRoot_(ss, rootApptId);
      var result = cs_submitFromDialog({
        assignedRep: target.rec.assignedRep || '',
        assistedRep: target.rec.assistedRep || '',
        salesStage: swTrim_(payload.salesStage),
        convStatus: swTrim_(payload.convStatus),
        customOrder: swTrim_(payload.customOrder),
        cosAllowedEmpty: !swTrim_(payload.customOrder),
        inProduction: swTrim_(payload.inProduction),
        centerStone: swTrim_(payload.centerStone),
        nextSteps: swTrim_(payload.nextSteps),
        orderDate: swTrim_(payload.orderDate),
        deadline3d: swTrim_(payload.deadline3d),
        prodDeadline: swTrim_(payload.prodDeadline),
        wax: null,
        waxSummary: '',
        notebookLMLink: swTrim_(payload.notebookLMLink)
      });
      if (result && result.ok === false) throw new Error(result.error || 'Client status update failed.');
      swCustomerSearchLog_(ss, 'CUSTOMER_SEARCH_STATUS_UPDATE', target.rec, user, payload, result);
    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }
    return swCustomerSearchDetailPayload_(ss, user, rootApptId);
  });
}

function sw_customerSearchUpdate3DDeadline(authToken, rootApptId, payload) {
  return swTimed_('sw_customerSearchUpdate3DDeadline', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken);
    swRequireCustomerSearchUser_(user);
    payload = payload || {};
    var deadline = swTrim_(payload.deadline3d || payload.dateIso || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(deadline)) throw new Error('Select a valid 3D deadline.');

    var lock = LockService.getDocumentLock() || LockService.getScriptLock();
    lock.waitLock(28000);
    try {
      var target = swCustomerSearchActivateRoot_(ss, rootApptId);
      var result = (typeof Deadlines !== 'undefined' && Deadlines.saveRecordDeadline)
        ? Deadlines.saveRecordDeadline({ kind: '3D', dateIso: deadline })
        : saveRecordDeadline({ kind: '3D', dateIso: deadline });
      if (result && result.ok === false) throw new Error(result.error || '3D deadline update failed.');
      swCustomerSearchLog_(ss, 'CUSTOMER_SEARCH_3D_DEADLINE', target.rec, user, { deadline3d: deadline, note: swTrim_(payload.note) }, result);
    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }
    return swCustomerSearchDetailPayload_(ss, user, rootApptId);
  });
}

function sw_customerSearchSubmit3DRevision(authToken, rootApptId, payload) {
  return swTimed_('sw_customerSearchSubmit3DRevision', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken);
    swRequireCustomerSearchUser_(user);
    payload = payload || {};
    var form = payload.form || payload;
    if (!swTrim_(form.DesignNotes)) throw new Error('Enter revision design notes.');

    var lock = LockService.getDocumentLock() || LockService.getScriptLock();
    lock.waitLock(28000);
    try {
      var target = swCustomerSearchActivateRoot_(ss, rootApptId);
      var result = submit3DRevision({ form: form });
      if (result && result.ok === false) throw new Error(result.error || '3D revision failed.');
      swCustomerSearchLog_(ss, 'CUSTOMER_SEARCH_3D_REVISION', target.rec, user, form, result);
    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }
    return swCustomerSearchDetailPayload_(ss, user, rootApptId);
  });
}

function sw_customerSearchRequestWax(authToken, rootApptId, payload) {
  return swTimed_('sw_customerSearchRequestWax', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken);
    swRequireCustomerSearchUser_(user);
    payload = payload || {};
    if (!swTrim_(payload.soMo)) throw new Error('Enter the SO/MO number for the wax request.');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(swTrim_(payload.neededByRep))) throw new Error('Select Needed By (Rep).');
    if (!swTrim_(payload.priority)) throw new Error('Select wax priority.');

    var lock = LockService.getDocumentLock() || LockService.getScriptLock();
    lock.waitLock(28000);
    try {
      var target = swCustomerSearchResolveRoot_(ss, rootApptId);
      var result = wax_onRequestSubmit_({
        rootApptId: target.root,
        soMo: swTrim_(payload.soMo),
        neededByRep: swTrim_(payload.neededByRep),
        priority: swTrim_(payload.priority),
        requestedBy: (user && (user.email || user.name)) || ''
      });
      if (result && result.ok === false) throw new Error(result.error || 'Wax request failed.');
      swCustomerSearchLog_(ss, 'CUSTOMER_SEARCH_WAX_REQUEST', target.rec, user, payload, result);
    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }
    return swCustomerSearchDetailPayload_(ss, user, rootApptId);
  });
}

function swRequireCustomerSearchUser_(user) {
  if (!(user && (user.isAdmin || user.isJoc || user.isRep))) {
    throw new Error('Customer Search access requires Client Advisor, JOC, or Admin role.');
  }
}

function swCustomerSearchNormalizeFilters_(filters) {
  filters = filters || {};
  var activeRaw = filters.activeOnly;
  var activeOnly = !(activeRaw === false || String(activeRaw || '').toLowerCase() === 'false');
  return {
    query: swTrim_(filters.query || ''),
    brand: swTrim_(filters.brand || ''),
    clientAdvisor: swTrim_(filters.clientAdvisor || ''),
    joc: swTrim_(filters.joc || ''),
    activeOnly: activeOnly
  };
}

function swCustomerSearchPublicFilters_(filters) {
  return {
    query: filters.query || '',
    brand: filters.brand || '',
    clientAdvisor: filters.clientAdvisor || '',
    joc: filters.joc || '',
    activeOnly: filters.activeOnly !== false
  };
}

function swCustomerSearchFilteredRows_(appointments, query, filters) {
  query = swTrim_(query);
  var q = swNorm_(query);
  var qPhone = swNormPhone_(query);
  var baseRows = (appointments || []).filter(function (rec) {
    if (filters.activeOnly && !swIsAppointmentActive_(rec)) return false;
    if (filters.brand && swNorm_(rec.brand) !== swNorm_(filters.brand)) return false;
    if (filters.clientAdvisor && swNorm_(rec.assignedRep) !== swNorm_(filters.clientAdvisor)) return false;
    if (filters.joc && swNorm_(rec.assistedRep) !== swNorm_(filters.joc)) return false;
    return true;
  });
  if (!q) return baseRows;

  var matchedRoots = {};
  baseRows.forEach(function (rec) {
    if (!swCustomerSearchRecordMatches_(rec, q, qPhone)) return;
    var root = swTrim_(rec.root || rec.appt);
    if (root) matchedRoots[root] = true;
  });
  return baseRows.filter(function (rec) {
    return !!matchedRoots[swTrim_(rec.root || rec.appt)];
  });
}

function swCustomerSearchRecordMatches_(rec, q, qPhone) {
  var fields = [
    rec.name, rec.email, rec.phone, rec.root, rec.appt, rec.so, rec.brand,
    rec.assignedRep, rec.assistedRep, rec.visitType, rec.visitDate,
    rec.salesStage, rec.convStatus, rec.customOrder, rec.nextSteps
  ];
  var text = swNorm_(fields.join(' '));
  if (text.indexOf(q) >= 0) return true;
  if (qPhone && swNormPhone_(rec.phone).indexOf(qPhone) >= 0) return true;
  return false;
}

function swCustomerSearchFilterOptions_(appointments) {
  var brands = [];
  var advisors = [];
  var jocs = [];
  (appointments || []).forEach(function (rec) {
    if (!swIsAppointmentActive_(rec)) return;
    if (rec.brand) brands.push(rec.brand);
    if (rec.assignedRep) advisors.push(rec.assignedRep);
    if (rec.assistedRep) jocs.push(rec.assistedRep);
  });
  return {
    brands: swUnique_(brands).sort(),
    clientAdvisors: swUnique_(advisors).sort(),
    jocs: swUnique_(jocs).sort()
  };
}

function swCustomerSearchCard_(ss, masterGid, root, rec, rootRows, stage) {
  var card = swAdminDashboardCustomerCard_(ss, masterGid, root, rec, rootRows, stage, { byRoot: {}, bySo: {} });
  card.email = rec.email || '';
  card.phone = rec.phone || '';
  card.deadline3d = rec.deadline3d || '';
  card.productionDeadline = rec.productionDeadline || '';
  card.waxStatus = rec.waxStatus || '';
  card.waxDeadlineAdmin = rec.waxDeadlineAdmin || '';
  card.centerStoneStatus = rec.centerStoneStatus || '';
  card.badges = swCustomerSearchBadges_(rec);
  return card;
}

function swCustomerSearchBadges_(rec) {
  var badges = [];
  if (!swTrim_(rec.assignedRep)) badges.push('Missing Advisor');
  if (!swTrim_(rec.assistedRep)) badges.push('Missing JOC');
  if (/3d revision/i.test(swTrim_(rec.customOrder))) badges.push('3D Revision');
  var d3 = swCustomerSearchDateMs_(rec.deadline3d);
  if (d3 && d3 < swCustomerSearchTodayMs_() && /3d/i.test(swTrim_(rec.customOrder))) badges.push('3D Overdue');
  var waxDeadline = swCustomerSearchDateMs_(rec.waxDeadlineAdmin);
  if (waxDeadline && waxDeadline < swCustomerSearchTodayMs_() && !/complete|cancel/i.test(swTrim_(rec.waxStatus))) badges.push('Wax Issue');
  return badges;
}

function swCustomerSearchDetailPayload_(ss, user, rootApptId) {
  var mark = swStepTimer_('swCustomerSearchDetailPayload');
  var target = swCustomerSearchResolveRoot_(ss, rootApptId);
  mark('resolveRoot', { rows: target.rows.length });
  var master = ss.getSheetByName(SW_SHEETS.MASTER);
  var masterGid = master ? master.getSheetId() : '';
  var stage = swAdminDashboardPipelineStage_(target.rec, target.rows);
  var card = swCustomerSearchCard_(ss, masterGid, target.root, target.rec, target.rows, stage);
  mark('card');
  var now = new Date().getTime();
  var state = swReadTaskListState_(ss, true);
  var tasks = (state.tasks || []).filter(function (t) {
    return (swTrim_(t.root) === target.root || swTrim_(t.appt) === target.root) && t.status !== SW_STATUSES.COMPLETED;
  }).map(function (t) {
    return swPublicTask_(t, now);
  });
  mark('tasks', { tasks: tasks.length });
  var logs = swCustomerSearchRecentLogs_(ss, target.root);
  mark('logs', { logs: logs.length });
  var formOptions = swTaskFormOptions_(ss, { taskType: SW_TASKS.POST_CONSULT_STATUS });
  mark('formOptions', { groups: formOptions ? Object.keys(formOptions).length : 0 });

  return {
    ok: true,
    generatedAt: swIso_(new Date()),
    user: user,
    root: target.root,
    card: card,
    appointments: target.rows.map(swCustomerSearchPublicAppointment_),
    tasks: tasks,
    logs: logs,
    formOptions: formOptions,
    actions: {
      updateStatus: true,
      update3dDeadline: true,
      submit3dRevision: true,
      requestWax: true
    }
  };
}

function swCustomerSearchPublicAppointment_(rec) {
  return {
    row: rec.row || '',
    root: rec.root || '',
    appt: rec.appt || '',
    customerName: rec.name || '',
    brand: rec.brand || '',
    visitDate: rec.visitDate || '',
    visitTime: rec.visitTime || '',
    visitType: rec.visitType || '',
    status: rec.status || '',
    assignedRep: rec.assignedRep || '',
    assistedRep: rec.assistedRep || '',
    so: rec.so || '',
    salesStage: rec.salesStage || '',
    conversionStatus: rec.convStatus || '',
    customOrderStatus: rec.customOrder || '',
    active: swIsAppointmentActive_(rec)
  };
}

function swCustomerSearchRecentLogs_(ss, root) {
  var sh = ss.getSheetByName(SW_SHEETS.LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  var last = sh.getLastRow();
  var start = Math.max(2, last - SW_CUSTOMER_SEARCH_LOG_SCAN_ROWS + 1);
  var values = sh.getRange(start, 1, last - start + 1, Math.min(sh.getLastColumn(), SW_LOG_HEADERS.length)).getDisplayValues();
  var out = [];
  for (var i = values.length - 1; i >= 0 && out.length < 20; i--) {
    var row = values[i];
    if (swTrim_(row[3]) !== root && swTrim_(row[4]) !== root) continue;
    out.push({
      eventAt: row[0] || '',
      eventType: row[1] || '',
      taskId: row[2] || '',
      root: row[3] || '',
      appt: row[4] || '',
      taskType: row[5] || '',
      actorName: row[6] || '',
      actorEmail: row[7] || '',
      status: row[10] || '',
      detailsJson: row[11] || ''
    });
  }
  return out;
}

function swCustomerSearchResolveRoot_(ss, rootApptId) {
  var want = swTrim_(rootApptId);
  if (!want) throw new Error('Missing customer/root id.');
  var rows = typeof swReadAppointmentsForRoot_ === 'function'
    ? swReadAppointmentsForRoot_(ss, want)
    : swReadAppointments_(ss).filter(function (rec) {
      return swTrim_(rec.root) === want || swTrim_(rec.appt) === want;
    });
  if (!rows.length) throw new Error('Customer not found: ' + want);
  var active = rows.filter(function (rec) { return swIsAppointmentActive_(rec); });
  var rec = swAdminDashboardLatestRow_(active.length ? active : rows);
  var root = swTrim_(rec.root || rec.appt || want);
  return { root: root, rec: rec, rows: rows };
}

function swCustomerSearchActivateRoot_(ss, rootApptId) {
  var target = swCustomerSearchResolveRoot_(ss, rootApptId);
  var sh = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!sh || !target.rec.row) throw new Error('Could not resolve Master row for this customer.');
  ss.setActiveSheet(sh);
  ss.setActiveRange(sh.getRange(target.rec.row, 1));
  return target;
}

function swCustomerSearchLog_(ss, eventType, rec, user, payload, result) {
  swAppendTaskLog_(ss, eventType, {
    taskId: '',
    root: rec.root || rec.appt || '',
    appt: rec.appt || '',
    taskType: eventType,
    status: SW_STATUSES.COMPLETED
  }, user, rec.assignedRep || '', rec.assignedRep || '', {
    payload: payload || {},
    result: result || {}
  });
}

function swCustomerSearchDateMs_(value) {
  value = swTrim_(value);
  if (!value) return 0;
  var iso = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(value);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])).getTime();
  var mdy = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/.exec(value);
  if (mdy) {
    var y = Number(mdy[3]);
    if (y < 100) y += 2000;
    return new Date(y, Number(mdy[1]) - 1, Number(mdy[2])).getTime();
  }
  var d = new Date(value);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function swCustomerSearchTodayMs_() {
  var today = new Date();
  return new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
}
