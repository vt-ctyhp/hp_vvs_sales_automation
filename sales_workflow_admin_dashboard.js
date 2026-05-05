/**
 * Admin dashboard read API for the Sales Workflow web app.
 * This module is read-only: it aggregates Master appointment rows, task health,
 * and the external payments ledger when configured.
 */

var SW_ADMIN_DASHBOARD_COLUMNS = [
  { key: 'lead', label: 'Lead' },
  { key: 'hotLead', label: 'Hot Lead' },
  { key: 'followUp', label: 'Follow-Up' },
  { key: 'appointment', label: 'Appointment / Viewing Scheduled' },
  { key: 'deposit', label: 'Deposit / Order In Progress' },
  { key: 'inProduction', label: 'In Production' },
  { key: 'won', label: 'Won / Completed' },
  { key: 'lost', label: 'Lost Lead' }
];

var SW_ADMIN_DASHBOARD_WINDOWS = [
  { value: 'today', label: 'Today' },
  { value: 'last7', label: 'Last 7 days' },
  { value: 'last30', label: 'Last 30 days' },
  { value: 'thisMonth', label: 'This month' },
  { value: 'custom', label: 'Custom' }
];

var SW_ADMIN_DASHBOARD_STAGE_WEIGHTS = {
  lead: 0.10,
  hotLead: 0.25,
  followUp: 0.20,
  appointment: 0.35,
  deposit: 0.85,
  inProduction: 0.95,
  won: 0,
  lost: 0
};
var SW_ADMIN_DASHBOARD_AUX_CACHE_SECONDS = 5 * 60;

/**
 * Read-only admin dashboard payload.
 */
function sw_getAdminDashboard(authToken, filters) {
  return swTimed_('sw_getAdminDashboard', function () {
    if (typeof authToken === 'object' && filters == null) {
      filters = authToken;
      authToken = '';
    }

    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken);
    if (!user.isAdmin) throw new Error('Admin access required.');

    var mark = swStepTimer_('sw_getAdminDashboard');
    filters = swAdminDashboardNormalizeFilters_(filters);
    mark('normalize', { mode: filters.mode });
    var appointments = swReadAppointments_(ss);
    mark('appointments', { rows: appointments.length });
    var scope = swAdminDashboardBuildScope_(appointments, filters);
    var warnings = [];
    var payments = swAdminDashboardReadPayments_(scope, filters, warnings);
    mark('payments', { receipts: payments.receipts ? payments.receipts.length : 0 });
    var state = swReadTaskListState_(ss, true);
    var tasks = swAdminDashboardOpenTasksFromState_(state);
    mark('tasks', { rows: tasks.length });
    var indexes = swAdminDashboardBuildIndexes_(appointments);
    mark('indexes', { roots: Object.keys(indexes.currentByRoot || {}).length });
    var currentByRoot = indexes.currentByRoot;
    var metrics = swAdminDashboardMetrics_(appointments, tasks, currentByRoot, payments, filters);
    mark('metrics');
    var healthContext = filters.mode === 'health'
      ? swAdminDashboardHealthContext_(ss, appointments, currentByRoot, payments, tasks, filters, warnings, indexes)
      : null;
    var health = healthContext ? swAdminDashboardHealth_(healthContext, metrics) : null;
    mark('health', { included: !!health });
    var kanban = filters.mode === 'kanban'
      ? swAdminDashboardKanban_(ss, appointments, payments, filters, indexes)
      : null;
    mark('kanban', { included: !!kanban });
    var filterOptions = swAdminDashboardFilterOptions_(ss, appointments, filters);
    mark('filterOptions', {
      brands: filterOptions.brands ? filterOptions.brands.length : 0,
      clientAdvisors: filterOptions.clientAdvisors ? filterOptions.clientAdvisors.length : 0,
      jocs: filterOptions.jocs ? filterOptions.jocs.length : 0
    });

    return {
      ok: true,
      generatedAt: swIso_(new Date()),
      filters: swAdminDashboardPublicFilters_(filters),
      filterOptions: filterOptions,
      metrics: metrics,
      health: health,
      kanban: kanban,
      taskCount: tasks.length,
      warnings: warnings
    };
  });
}

function swAdminDashboardNormalizeFilters_(filters) {
  filters = filters || {};
  var preset = swTrim_(filters.windowPreset || filters.window || '');
  if (!preset) preset = (filters.startDate || filters.endDate) ? 'custom' : 'last7';
  preset = swAdminDashboardNormalizeWindowPreset_(preset);

  var window = swAdminDashboardWindowForPreset_(preset);
  var start = preset === 'custom'
    ? (swAdminDashboardParseDate_(filters.startDate) || window.start)
    : window.start;
  var end = preset === 'custom'
    ? (swAdminDashboardParseDate_(filters.endDate) || window.end)
    : window.end;
  start = swAdminDashboardStartOfDay_(start);
  end = swAdminDashboardEndOfDay_(end);
  if (end.getTime() < start.getTime()) {
    var tmp = start;
    start = swAdminDashboardStartOfDay_(end);
    end = swAdminDashboardEndOfDay_(tmp);
  }
  return {
    start: start,
    end: end,
    startDate: swAdminDashboardDateKey_(start),
    endDate: swAdminDashboardDateKey_(end),
    windowPreset: preset,
    windowLabel: swAdminDashboardWindowLabel_(preset, start, end),
    mode: swAdminDashboardNormalizeMode_(filters.mode || filters.dashboardMode || filters.view),
    brand: swTrim_(filters.brand),
    clientAdvisor: swTrim_(filters.clientAdvisor),
    joc: swTrim_(filters.joc),
    includeClosed: filters.includeClosed === true || String(filters.includeClosed || '').toLowerCase() === 'true'
  };
}

function swAdminDashboardNormalizeMode_(mode) {
  var key = swNorm_(mode);
  return key === 'kanban' ? 'kanban' : 'health';
}

function swAdminDashboardPublicFilters_(filters) {
  return {
    startDate: filters.startDate,
    endDate: filters.endDate,
    windowPreset: filters.windowPreset || 'last7',
    windowLabel: filters.windowLabel || '',
    mode: filters.mode || 'health',
    brand: filters.brand || '',
    clientAdvisor: filters.clientAdvisor || '',
    joc: filters.joc || '',
    includeClosed: !!filters.includeClosed
  };
}

function swAdminDashboardNormalizeWindowPreset_(preset) {
  var key = swNorm_(preset).replace(/[^a-z0-9]/g, '');
  if (key === 'today') return 'today';
  if (key === 'last30' || key === '30days' || key === 'last30days') return 'last30';
  if (key === 'thismonth' || key === 'month') return 'thisMonth';
  if (key === 'custom') return 'custom';
  return 'last7';
}

function swAdminDashboardWindowForPreset_(preset) {
  var now = new Date();
  var today = swAdminDashboardStartOfDay_(now);
  var start = new Date(today);
  if (preset === 'today') {
    start = today;
  } else if (preset === 'last30') {
    start = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 29);
  } else if (preset === 'thisMonth') {
    start = new Date(today.getFullYear(), today.getMonth(), 1);
  } else {
    start = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 6);
  }
  var end = swAdminDashboardEndOfDay_(today);
  return { start: start, end: end };
}

function swAdminDashboardWindowLabel_(preset, start, end) {
  for (var i = 0; i < SW_ADMIN_DASHBOARD_WINDOWS.length; i++) {
    if (SW_ADMIN_DASHBOARD_WINDOWS[i].value === preset && preset !== 'custom') {
      return SW_ADMIN_DASHBOARD_WINDOWS[i].label;
    }
  }
  return [swAdminDashboardDateKey_(start), swAdminDashboardDateKey_(end)].filter(Boolean).join(' to ');
}

function swAdminDashboardParseDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  var s = swTrim_(value);
  if (!s) return null;
  var iso = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  var mdy = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/.exec(s);
  if (mdy) {
    var y = Number(mdy[3]);
    if (y < 100) y += 2000;
    return new Date(y, Number(mdy[1]) - 1, Number(mdy[2]));
  }
  var parsed = new Date(s);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function swAdminDashboardStartOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
}

function swAdminDashboardEndOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59, 999);
}

function swAdminDashboardDateKey_(date) {
  return Utilities.formatDate(date, swTimezone_(), 'yyyy-MM-dd');
}

function swAdminDashboardDateTimeValue_(raw, display) {
  if (raw instanceof Date && !isNaN(raw.getTime())) return raw;
  var s = swTrim_(display || raw);
  if (!s) return null;
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  return swAdminDashboardParseDate_(s);
}

function swAdminDashboardInRange_(date, filters) {
  return date instanceof Date && !isNaN(date.getTime()) &&
    date.getTime() >= filters.start.getTime() &&
    date.getTime() <= filters.end.getTime();
}

function swAdminDashboardBuildScope_(appointments, filters) {
  var out = {
    global: !(filters.brand || filters.clientAdvisor || filters.joc),
    roots: {},
    sos: {},
    brandOnly: !!(filters.brand && !filters.clientAdvisor && !filters.joc),
    brand: filters.brand || ''
  };
  (appointments || []).forEach(function (rec) {
    if (!swAdminDashboardRecordMatchesOwnerFilters_(rec, filters)) return;
    var root = swAdminDashboardCleanId_(rec.root || rec.appt);
    var so = swAdminDashboardCleanId_(rec.so);
    if (root) out.roots[root] = true;
    if (so) out.sos[so] = true;
  });
  return out;
}

function swAdminDashboardRecordMatchesOwnerFilters_(rec, filters) {
  if (filters.brand && swNorm_(rec.brand) !== swNorm_(filters.brand)) return false;
  if (filters.clientAdvisor && swNorm_(rec.assignedRep) !== swNorm_(filters.clientAdvisor)) return false;
  if (filters.joc && swNorm_(rec.assistedRep) !== swNorm_(filters.joc)) return false;
  return true;
}

function swAdminDashboardCleanId_(value) {
  return swTrim_(value).replace(/^'/, '');
}

function swAdminDashboardMetrics_(appointments, tasks, currentByRoot, payments, filters) {
  var metrics = {
    bookingsCreated: 0,
    consultationsScheduled: 0,
    viewingsScheduled: 0,
    consultationsActual: 0,
    viewingsActual: 0,
    paymentsCount: payments.count || 0,
    paymentsNet: payments.netAmount || 0,
    firstDepositCount: payments.firstDepositCount || 0,
    firstDepositNet: payments.firstDepositNet || 0,
    adminOpenTasks: 0,
    adminOverdueTasks: 0
  };

  var tz = swTimezone_();
  (appointments || []).forEach(function (rec) {
    if (!swAdminDashboardRecordMatchesOwnerFilters_(rec, filters)) return;

    var bookedAt = swAdminDashboardDateTimeValue_(rec.bookedAtRaw, rec.bookedAt);
    if (swAdminDashboardInRange_(bookedAt, filters) && !swTrim_(rec.rescheduledFromUid)) {
      metrics.bookingsCreated++;
    }

    var visitAt = swVisitDateTime_(rec, tz);
    if (!swAdminDashboardInRange_(visitAt, filters)) return;

    var type = swAdminDashboardVisitType_(rec);
    if (!type) return;
    if (swIsAppointmentActive_(rec)) {
      if (type === 'consultation') metrics.consultationsScheduled++;
      if (type === 'viewing') metrics.viewingsScheduled++;
    }
    if (swAdminDashboardIsActualStatus_(rec.status)) {
      if (type === 'consultation') metrics.consultationsActual++;
      if (type === 'viewing') metrics.viewingsActual++;
    }
  });

  var nowMs = new Date().getTime();
  (tasks || []).forEach(function (task) {
    if (!swAdminDashboardTaskMatchesFilters_(task, currentByRoot, filters)) return;
    metrics.adminOpenTasks++;
    if (swTaskPendingLike_(task, nowMs) && swDateValue_(task.dueAt) < nowMs) {
      metrics.adminOverdueTasks++;
    }
  });

  return metrics;
}

function swAdminDashboardVisitType_(rec) {
  var type = swNorm_(rec.visitType);
  if (type === 'appointment' || type.indexOf('consult') >= 0) return 'consultation';
  if (type === 'diamond viewing' || type.indexOf('viewing') >= 0) return 'viewing';
  return '';
}

function swAdminDashboardIsActualStatus_(status) {
  var s = swNorm_(status);
  return s === 'completed' || s === 'attended' || s === 'done';
}

function swAdminDashboardTaskMatchesFilters_(task, currentByRoot, filters) {
  var rec = currentByRoot[swAdminDashboardCleanId_(task.root || task.appt)] || null;
  if (rec) return swAdminDashboardRecordMatchesOwnerFilters_(rec, filters);
  if (filters.brand && swNorm_(task.brand) !== swNorm_(filters.brand)) return false;
  if (filters.clientAdvisor && swNorm_(task.currentOwner || task.intendedOwner) !== swNorm_(filters.clientAdvisor)) return false;
  if (filters.joc && swNorm_(task.currentOwner || task.intendedOwner) !== swNorm_(filters.joc)) return false;
  return true;
}

function swAdminDashboardOpenTasksFromState_(state) {
  return (state && state.tasks ? state.tasks : []).filter(function (task) {
    return task && task.taskId && task.status !== SW_STATUSES.COMPLETED;
  });
}

function swAdminDashboardReadSelectedDisplayColumns_(sh, indexes) {
  var rowCount = sh.getLastRow() - 1;
  if (rowCount <= 0) return [];
  var columns = [];
  var seen = {};
  (indexes || []).forEach(function (idx) {
    idx = Number(idx);
    if (!isFinite(idx) || idx < 0 || seen[idx]) return;
    seen[idx] = true;
    columns.push(idx);
  });
  columns.sort(function (a, b) { return a - b; });

  var out = [];
  for (var i = 0; i < rowCount; i++) out.push([]);
  var start = null;
  var group = [];
  function flush() {
    if (start == null || !group.length) return;
    var width = group[group.length - 1] - start + 1;
    var values = sh.getRange(2, start + 1, rowCount, width).getDisplayValues();
    for (var r = 0; r < values.length; r++) {
      for (var c = 0; c < group.length; c++) {
        out[r][group[c]] = values[r][group[c] - start];
      }
    }
  }
  columns.forEach(function (idx) {
    if (start == null) {
      start = idx;
      group = [idx];
      return;
    }
    if (idx === group[group.length - 1] + 1) {
      group.push(idx);
      return;
    }
    flush();
    start = idx;
    group = [idx];
  });
  flush();
  return out;
}

function swAdminDashboardBuildIndexes_(appointments) {
  var groups = swAdminDashboardRowsByRoot_(appointments);
  var currentByRoot = {};
  Object.keys(groups).forEach(function (root) {
    var rows = groups[root];
    var active = rows.filter(function (rec) { return swIsAppointmentActive_(rec); });
    currentByRoot[root] = swAdminDashboardLatestRow_(active.length ? active : rows);
  });
  return {
    groups: groups,
    currentByRoot: currentByRoot
  };
}

function swAdminDashboardFilterOptions_(ss, appointments, filters) {
  var brands = {};
  var advisors = {};
  var jocs = {};
  (appointments || []).forEach(function (rec) {
    if (rec.brand) brands[rec.brand] = true;
    if (rec.assignedRep) advisors[rec.assignedRep] = true;
    if (rec.assistedRep) jocs[rec.assistedRep] = true;
  });
  try {
    var assignment = swReadAssignmentOptions_(ss);
    (assignment.salesReps || []).forEach(function (item) {
      if (item && item.name) advisors[item.name] = true;
    });
    (assignment.jocReps || []).forEach(function (item) {
      if (item && item.name) jocs[item.name] = true;
    });
  } catch (_) {}
  if (filters.brand) brands[filters.brand] = true;
  if (filters.clientAdvisor) advisors[filters.clientAdvisor] = true;
  if (filters.joc) jocs[filters.joc] = true;
  return {
    windows: SW_ADMIN_DASHBOARD_WINDOWS,
    brands: swAdminDashboardSortedKeys_(brands),
    clientAdvisors: swAdminDashboardSortedKeys_(advisors),
    jocs: swAdminDashboardSortedKeys_(jocs)
  };
}

function swAdminDashboardSortedKeys_(map) {
  return Object.keys(map || {}).filter(Boolean).sort(function (a, b) {
    return String(a).localeCompare(String(b));
  });
}

function swAdminDashboardReadPayments_(scope, filters, warnings) {
  var out = {
    count: 0,
    netAmount: 0,
    firstDepositCount: 0,
    firstDepositNet: 0,
    receipts: [],
    firstDeposits: [],
    firstByKey: {},
    byRoot: {},
    bySo: {}
  };

  var target = null;
  try {
    target = swAdminDashboardPaymentsSheet_();
  } catch (err) {
    warnings.push('Payments ledger unavailable: ' + (err && err.message ? err.message : err));
    return out;
  }
  if (!target || !target.sh) {
    warnings.push('Payments ledger unavailable.');
    return out;
  }

  var sh = target.sh;
  var lr = sh.getLastRow();
  var lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return out;

  var headers = sh.getRange(1, 1, 1, lc).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var C = {
    root: swPickIndex_(H, ['RootApptID', 'APPT_ID', 'Root Appt ID', 'Appointment ID']),
    so: swPickIndex_(H, ['SO#', 'SO', 'SO Number', 'Sales Order', 'Sales Order #']),
    brand: swPickIndex_(H, ['Brand']),
    docType: swPickIndex_(H, ['DocType', 'Doc Type', 'Document Type', 'Type']),
    docStatus: swPickIndex_(H, ['DocStatus', 'Doc Status', 'Status']),
    when: swPickIndex_(H, ['PaymentDateTime', 'Payment DateTime', 'Payment Date/Time', 'Payment Date', 'Paid At']),
    amountNet: swPickIndex_(H, ['AmountNet', 'Net', 'Net Amount']),
    amountGross: swPickIndex_(H, ['AmountGross', 'Gross', 'Amount']),
    balance: swPickIndex_(H, ['Balance_SO', 'Balance SO', 'BalanceDue', 'Balance Due']),
    orderTotal: swPickIndex_(H, ['Order_Total_SO', 'Order Total SO', 'OrderTotalValue', 'Order Total'])
  };
  if (C.docType < 0 || C.when < 0 || (C.amountNet < 0 && C.amountGross < 0)) {
    warnings.push('Payments ledger is missing DocType, PaymentDateTime, or amount columns.');
    return out;
  }

  var rowCount = lr - 1;
  var indexes = swAdminDashboardPaymentColumnIndexes_(C);
  var values = swReadSelectedRows_(sh, 2, rowCount, indexes, 'values');
  var receipts = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var docType = swTrim_(swCell_(row, C.docType));
    if (!/receipt/i.test(docType)) continue;
    var status = swNorm_(swCell_(row, C.docStatus));
    if (/void|replaced|cancel|draft|deleted/.test(status)) continue;

    var root = swAdminDashboardCleanId_(swCell_(row, C.root));
    var so = swAdminDashboardCleanId_(swCell_(row, C.so));
    var brand = swTrim_(swCell_(row, C.brand));
    if (!swAdminDashboardPaymentInScope_(root, so, brand, scope, filters)) continue;

    var when = swAdminDashboardDateTimeValue_(swCell_(row, C.when), swCell_(row, C.when));
    if (!when) continue;

    var net = swAdminDashboardNumber_(C.amountNet >= 0 ? swCell_(row, C.amountNet) : swCell_(row, C.amountGross));
    var balance = C.balance >= 0 ? swAdminDashboardNumberOrBlank_(swCell_(row, C.balance)) : '';
    var orderTotal = C.orderTotal >= 0 ? swAdminDashboardNumberOrBlank_(swCell_(row, C.orderTotal)) : '';
    var key = root || so;
    if (!key) continue;

    var receipt = { root: root, so: so, key: key, when: when, net: net, balance: balance, orderTotal: orderTotal };
    receipts.push(receipt);
    out.receipts.push(receipt);

    if (swAdminDashboardInRange_(when, filters)) {
      out.count++;
      out.netAmount += net;
    }

    swAdminDashboardApplyPaymentSummary_(out, root, so, when, net, balance, orderTotal);
  }

  var firstByKey = {};
  receipts.forEach(function (item) {
    var existing = firstByKey[item.key];
    if (!existing || item.when.getTime() < existing.when.getTime()) firstByKey[item.key] = item;
  });
  Object.keys(firstByKey).forEach(function (key) {
    var item = firstByKey[key];
    out.firstByKey[key] = item;
    if (swAdminDashboardInRange_(item.when, filters)) {
      out.firstDepositCount++;
      out.firstDepositNet += item.net;
      out.firstDeposits.push(item);
    }
  });

  return out;
}

function swAdminDashboardPaymentColumnIndexes_(columns) {
  var out = [];
  Object.keys(columns || {}).forEach(function (key) {
    var col = Number(columns[key]);
    if (isFinite(col) && col >= 0) out.push(col);
  });
  return out;
}

function swAdminDashboardPaymentsSheet_() {
  if (typeof rp_getLedgerTarget === 'function') {
    return rp_getLedgerTarget();
  }
  if (typeof pr_getLedger_ === 'function' && typeof pr_getPaymentsSheet_ === 'function') {
    var ledger = pr_getLedger_();
    return { sh: pr_getPaymentsSheet_(ledger) };
  }
  throw new Error('No payments ledger helper is available.');
}

function swAdminDashboardPaymentInScope_(root, so, brand, scope, filters) {
  if (scope.global) return true;
  if (root && scope.roots[root]) return true;
  if (so && scope.sos[so]) return true;
  if (scope.brandOnly && filters.brand && swNorm_(brand) === swNorm_(filters.brand)) return true;
  return false;
}

function swAdminDashboardApplyPaymentSummary_(out, root, so, when, net, balance, orderTotal) {
  var targets = [];
  if (root) targets.push({ map: out.byRoot, key: root });
  if (so) targets.push({ map: out.bySo, key: so });
  targets.forEach(function (target) {
    if (!target.map[target.key]) {
      target.map[target.key] = {
        paymentCount: 0,
        paidNet: 0,
        lastPaymentDate: '',
        lastPaymentMs: 0,
        balanceDue: '',
        orderTotal: ''
      };
    }
    var item = target.map[target.key];
    item.paymentCount++;
    item.paidNet += net;
    if (when.getTime() >= item.lastPaymentMs) {
      item.lastPaymentMs = when.getTime();
      item.lastPaymentDate = swAdminDashboardDateKey_(when);
      if (balance !== '') item.balanceDue = balance;
      if (orderTotal !== '') item.orderTotal = orderTotal;
    }
  });
}

function swAdminDashboardNumber_(value) {
  var n = Number(String(value == null ? '' : value).replace(/[^0-9.\-]/g, ''));
  return isFinite(n) ? n : 0;
}

function swAdminDashboardNumberOrBlank_(value) {
  if (value === '' || value == null) return '';
  var n = swAdminDashboardNumber_(value);
  return isFinite(n) ? n : '';
}

function swAdminDashboardHealthContext_(ss, appointments, currentByRoot, payments, tasks, filters, warnings, indexes) {
  var mark = swStepTimer_('swAdminDashboardHealthContext');
  indexes = indexes || {};
  var groups = indexes.groups || swAdminDashboardRowsByRoot_(appointments);
  mark('groups', { roots: Object.keys(groups || {}).length });
  var rootIndex = swAdminDashboardReadRootIndex_(ss, warnings);
  mark('rootIndex', { available: rootIndex.available, cacheHit: !!rootIndex.cacheHit });
  var statusLog = swAdminDashboardReadStatusLog_(ss, appointments, warnings, currentByRoot);
  mark('statusLog', { available: statusLog.available, cacheHit: !!statusLog.cacheHit });
  var stageWeights = swAdminDashboardStageWeights_(ss);
  mark('stageWeights');
  var master = ss.getSheetByName(SW_SHEETS.MASTER);
  var masterGid = master ? master.getSheetId() : '';
  var rows = [];
  var tz = swTimezone_();

  Object.keys(currentByRoot || {}).forEach(function (root) {
    var rec = currentByRoot[root];
    if (!rec || !swAdminDashboardRecordMatchesOwnerFilters_(rec, filters)) return;

    var rootRows = groups[root] || [];
    var stage = swAdminDashboardPipelineStage_(rec, rootRows);
    if (!filters.includeClosed && (stage.key === 'won' || stage.key === 'lost')) return;

    var pay = (payments.byRoot && payments.byRoot[root]) ||
      (rec.so && payments.bySo && payments.bySo[swAdminDashboardCleanId_(rec.so)]) ||
      {};
    var firstPayment = (payments.firstByKey && (payments.firstByKey[root] || payments.firstByKey[swAdminDashboardCleanId_(rec.so)])) || null;
    var firstVisit = swAdminDashboardFirstVisit_(rootRows, tz);
    var latestVisit = swAdminDashboardLatestVisit_(rootRows, tz);
    var bookedAt = swAdminDashboardDateTimeValue_(rec.bookedAtRaw, rec.bookedAt);
    var updatedAt = swAdminDashboardDateTimeValue_(rec.updatedAtRaw, rec.updatedAt);
    var lastPaymentDate = pay.lastPaymentMs ? new Date(pay.lastPaymentMs) : swAdminDashboardDateTimeValue_(rec.lastPaymentDateRaw, rec.lastPaymentDate);
    var lastTouch = rootIndex.byRoot[root] || updatedAt || lastPaymentDate || null;
    var orderDate = swAdminDashboardDateTimeValue_(rec.orderDateRaw, rec.orderDate);
    var d3Deadline = swAdminDashboardDateTimeValue_(rec.deadline3d, rec.deadline3d);
    var productionDeadline = swAdminDashboardDateTimeValue_(rec.productionDeadline, rec.productionDeadline);
    var d3 = statusLog.threeDByRoot[root] || {};
    var prod = statusLog.productionByRoot[root] || {};
    var d3Pending = !!d3.pending || /3d requested|3d revision/.test(swNorm_(rec.customOrder));
    var d3RequestDate = d3.requestDate || null;
    var d3AgeDays = d3RequestDate ? swAdminDashboardCalendarDays_(d3RequestDate, new Date()) : null;
    var threeDDeadlineOverdue = d3Pending && d3Deadline && swAdminDashboardStartOfDay_(d3Deadline).getTime() < swAdminDashboardStartOfDay_(new Date()).getTime();
    var threeDBlocked = d3Pending && ((d3AgeDays != null && d3AgeDays > 3) || threeDDeadlineOverdue);
    var orderAgeDays = orderDate ? swAdminDashboardCalendarDays_(orderDate, new Date()) : null;
    var productionDrag = stage.key === 'inProduction' && (
      (productionDeadline && swAdminDashboardStartOfDay_(productionDeadline).getTime() < swAdminDashboardStartOfDay_(new Date()).getTime()) ||
      (orderAgeDays != null && orderAgeDays > 30)
    );

    var orderTotal = swAdminDashboardBestNumber_([pay.orderTotal, rec.orderTotal]);
    var paidNet = swAdminDashboardBestNumber_([pay.paidNet, rec.paidToDate]);
    var balanceDue = swAdminDashboardBestNumber_([pay.balanceDue, rec.remainingBalance]);
    var budgetMid = swAdminDashboardBudgetMidpoint_(rec);
    var valueForWeight = orderTotal > 0 ? orderTotal : budgetMid;
    var stageWeight = stageWeights[stage.key] != null ? stageWeights[stage.key] : 0;
    var isOpen = stage.key !== 'won' && stage.key !== 'lost';
    var weightedValue = isOpen ? valueForWeight * stageWeight : 0;
    var quietHours = lastTouch ? Math.floor((new Date().getTime() - lastTouch.getTime()) / 3600000) : null;
    var d3Moves = swAdminDashboardNumber_(rec.deadline3dMoves);
    var productionMoves = swAdminDashboardNumber_(rec.productionDeadlineMoves);

    rows.push({
      root: root,
      appt: rec.appt || '',
      row: rec.row || '',
      customerName: rec.name || 'No customer',
      brand: rec.brand || '',
      clientAdvisor: rec.assignedRep || 'Unassigned',
      joc: rec.assistedRep || 'Unassigned',
      source: rec.source || '',
      stageKey: stage.key,
      stageLabel: stage.label,
      isOpen: isOpen,
      isAppointmentActive: swIsAppointmentActive_(rec),
      firstVisit: firstVisit,
      latestVisit: latestVisit,
      bookedAt: bookedAt,
      firstDepositDate: firstPayment ? firstPayment.when : null,
      firstDepositNet: firstPayment ? Number(firstPayment.net || 0) : 0,
      paidNet: paidNet,
      balanceDue: balanceDue,
      orderTotal: orderTotal,
      budgetMid: budgetMid,
      valueForWeight: valueForWeight,
      stageWeight: stageWeight,
      weightedValue: weightedValue,
      lastTouch: lastTouch,
      quietHours: quietHours,
      d3Pending: d3Pending,
      d3RequestDate: d3RequestDate,
      d3AgeDays: d3AgeDays,
      threeDDeadlineOverdue: threeDDeadlineOverdue,
      threeDBlocked: threeDBlocked,
      productionDeadline: productionDeadline,
      productionDrag: productionDrag,
      productionStageUpdatedAt: prod.updatedAt || null,
      orderDate: orderDate,
      orderAgeDays: orderAgeDays,
      escalations: d3Moves + productionMoves,
      nextSteps: rec.nextSteps || '',
      so: rec.so || '',
      masterUrl: masterGid && rec.row ? ('https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + masterGid + '&range=A' + rec.row) : ''
    });
  });
  mark('rows', { rows: rows.length });

  return {
    ss: ss,
    appointments: appointments || [],
    rows: rows,
    payments: payments || {},
    tasks: tasks || [],
    filters: filters,
    rootIndex: rootIndex,
    statusLog: statusLog,
    stageWeights: stageWeights
  };
}

function swAdminDashboardHealth_(ctx, metrics) {
  var rows = ctx.rows || [];
  var openRows = rows.filter(function (row) { return row.isOpen; });
  var noTouchRows = openRows.filter(function (row) {
    return row.quietHours != null && row.quietHours > 48;
  });
  var threeDRows = openRows.filter(function (row) { return row.threeDBlocked; });
  var productionRows = openRows.filter(function (row) { return row.productionDrag; });
  var receivableRows = openRows.filter(function (row) { return row.balanceDue > 0; });

  return {
    asOfDate: swAdminDashboardDateKey_(new Date()),
    windowLabel: ctx.filters.windowLabel || '',
    snapshot: swAdminDashboardHealthSnapshot_(noTouchRows, threeDRows, productionRows),
    pulse: swAdminDashboardHealthPulse_(ctx, metrics),
    pipeline: swAdminDashboardHealthPipeline_(ctx, openRows),
    trend: swAdminDashboardHealthTrend_(ctx),
    leadSources: swAdminDashboardHealthLeadSources_(ctx),
    advisorScorecard: swAdminDashboardHealthAdvisorScorecard_(ctx),
    topDeals: swAdminDashboardHealthTopDeals_(openRows),
    receivables: swAdminDashboardHealthReceivables_(receivableRows)
  };
}

function swAdminDashboardHealthSnapshot_(noTouchRows, threeDRows, productionRows) {
  return [
    {
      key: 'noTouch48',
      title: 'Hot leads going cold',
      value: noTouchRows.length,
      caption: 'active customers without a touch in 48h',
      footnote: swAdminDashboardCurrency_(swAdminDashboardSum_(noTouchRows, 'weightedValue')) + ' weighted pipeline at risk',
      tone: 'danger'
    },
    {
      key: 'threeDBlocked',
      title: '3D approval blocked',
      value: threeDRows.length,
      caption: '3D requests older than 3 days or overdue',
      footnote: swAdminDashboardCountWaiting_(threeDRows) + ' waiting >3 days',
      tone: 'warn'
    },
    {
      key: 'productionDrag',
      title: 'Production drag',
      value: productionRows.length,
      caption: 'orders with overdue production or 30d+ age',
      footnote: swAdminDashboardCurrency_(swAdminDashboardSum_(productionRows, 'balanceDue')) + ' outstanding',
      tone: 'info'
    }
  ];
}

function swAdminDashboardHealthPulse_(ctx, metrics) {
  var medianDays = swAdminDashboardMedianLeadToDeposit_(ctx.rows || [], ctx.filters);
  var winRate = swAdminDashboardWinRate90_(ctx.rows || []);
  return [
    {
      key: 'bookings',
      label: 'Bookings created',
      value: metrics.bookingsCreated || 0,
      caption: 'new appointments booked in window',
      trend: ''
    },
    {
      key: 'deposits',
      label: 'Deposits collected',
      value: swAdminDashboardCurrency_(metrics.firstDepositNet || 0),
      caption: (metrics.firstDepositCount || 0) + ' first-time deposit(s)',
      trend: ''
    },
    {
      key: 'winRate',
      label: 'Win rate (90d)',
      value: winRate.available ? (Math.round(winRate.value) + '%') : 'NA',
      caption: winRate.available ? (winRate.won + ' won / ' + winRate.closed + ' closed') : 'no won/lost records in 90d',
      trend: ''
    },
    {
      key: 'leadToDeposit',
      label: 'Median lead to deposit',
      value: medianDays == null ? 'NA' : (medianDays + 'd'),
      caption: medianDays == null ? 'no first deposits in window' : 'from first visit to first deposit',
      trend: ''
    }
  ];
}

function swAdminDashboardHealthPipeline_(ctx, openRows) {
  var byKey = {};
  SW_ADMIN_DASHBOARD_COLUMNS.forEach(function (col) {
    byKey[col.key] = {
      key: col.key,
      label: col.label,
      count: 0,
      weightedValue: 0,
      receivables: 0
    };
  });
  openRows.forEach(function (row) {
    var item = byKey[row.stageKey] || byKey.lead;
    item.count++;
    item.weightedValue += row.weightedValue || 0;
    item.receivables += row.balanceDue || 0;
  });
  var stages = SW_ADMIN_DASHBOARD_COLUMNS.filter(function (col) {
    return col.key !== 'won' && col.key !== 'lost';
  }).map(function (col) {
    return byKey[col.key];
  });
  var largest = stages.reduce(function (best, item) {
    if (!best || item.count > best.count) return item;
    return best;
  }, null);
  return {
    activeCustomers: openRows.length,
    weightedValue: swAdminDashboardSum_(openRows, 'weightedValue'),
    outstandingReceivables: swAdminDashboardSum_(openRows, 'balanceDue'),
    stages: stages,
    insight: largest && largest.count
      ? (largest.label + ' is the largest active stage by count.')
      : 'No active pipeline customers match these filters.'
  };
}

function swAdminDashboardHealthTrend_(ctx) {
  var end = swAdminDashboardEndOfWeek_(ctx.filters.end);
  var weeks = [];
  var maxBookings = 0;
  var maxDeposits = 0;
  for (var i = 11; i >= 0; i--) {
    var weekStart = swAdminDashboardAddDays_(swAdminDashboardStartOfWeek_(end), -7 * i);
    var weekEnd = swAdminDashboardEndOfDay_(swAdminDashboardAddDays_(weekStart, 6));
    var bookings = swAdminDashboardBookingsInRange_(ctx.appointments, ctx.filters, weekStart, weekEnd);
    var deposits = swAdminDashboardFirstDepositsInRange_(ctx.payments, weekStart, weekEnd);
    maxBookings = Math.max(maxBookings, bookings);
    maxDeposits = Math.max(maxDeposits, deposits);
    weeks.push({
      label: Utilities.formatDate(weekStart, swTimezone_(), 'MMM d'),
      startDate: swAdminDashboardDateKey_(weekStart),
      endDate: swAdminDashboardDateKey_(weekEnd),
      bookings: bookings,
      firstDeposits: deposits
    });
  }
  return {
    weeks: weeks,
    maxBookings: maxBookings,
    maxDeposits: maxDeposits
  };
}

function swAdminDashboardHealthLeadSources_(ctx) {
  var cutoff = swAdminDashboardAddDays_(swAdminDashboardStartOfDay_(new Date()), -89);
  var counts = {};
  var total = 0;
  var hasSource = false;
  (ctx.rows || []).forEach(function (row) {
    if (!row.source) return;
    hasSource = true;
    var when = row.bookedAt || row.firstVisit || row.latestVisit;
    if (when && when.getTime() < cutoff.getTime()) return;
    var source = row.source || 'Did not disclose';
    counts[source] = (counts[source] || 0) + 1;
    total++;
  });
  if (!hasSource) {
    return {
      available: false,
      rows: [],
      note: 'Lead source data is unavailable or empty in 00_Master Appointments.'
    };
  }
  var rows = Object.keys(counts).map(function (source) {
    return {
      source: source,
      count: counts[source],
      pct: total ? Math.round((counts[source] / total) * 100) : 0
    };
  }).sort(function (a, b) {
    if (b.count !== a.count) return b.count - a.count;
    return String(a.source).localeCompare(String(b.source));
  }).slice(0, 8);
  return {
    available: true,
    rows: rows,
    note: total + ' sourced customer(s) in the last 90 days'
  };
}

function swAdminDashboardHealthAdvisorScorecard_(ctx) {
  var groups = {};
  (ctx.rows || []).forEach(function (row) {
    var key = row.clientAdvisor || 'Unassigned';
    if (!groups[key]) {
      groups[key] = {
        advisor: key,
        customers: 0,
        weightedPipeline: 0,
        collected: 0,
        won: 0,
        closed: 0,
        noTouch: 0,
        threeDOverdue: 0,
        escalations: 0
      };
    }
    var item = groups[key];
    if (row.isOpen) {
      item.customers++;
      item.weightedPipeline += row.weightedValue || 0;
      if (row.quietHours != null && row.quietHours > 48) item.noTouch++;
      if (row.threeDBlocked) item.threeDOverdue++;
      if (row.escalations > 0 || row.productionDrag) item.escalations++;
    }
    item.collected += row.paidNet || 0;
    if (row.stageKey === 'won' || row.stageKey === 'lost') {
      item.closed++;
      if (row.stageKey === 'won') item.won++;
    }
  });
  return {
    rows: Object.keys(groups).map(function (key) {
      var item = groups[key];
      item.winRate = item.closed ? Math.round((item.won / item.closed) * 100) : null;
      return item;
    }).sort(function (a, b) {
      return (b.customers - a.customers) || String(a.advisor).localeCompare(String(b.advisor));
    })
  };
}

function swAdminDashboardHealthTopDeals_(openRows) {
  return openRows.filter(function (row) {
    return row.weightedValue > 0;
  }).sort(function (a, b) {
    return b.weightedValue - a.weightedValue;
  }).slice(0, 6).map(function (row) {
    return {
      root: row.root,
      customerName: row.customerName,
      advisor: row.clientAdvisor,
      brand: row.brand,
      stageLabel: row.stageLabel,
      weightedValue: row.weightedValue,
      quietDays: row.quietHours == null ? null : Math.floor(row.quietHours / 24),
      source: row.source || '',
      masterUrl: row.masterUrl || ''
    };
  });
}

function swAdminDashboardHealthReceivables_(receivableRows) {
  var sorted = receivableRows.sort(function (a, b) {
    return b.balanceDue - a.balanceDue;
  });
  return {
    openBalance: swAdminDashboardSum_(receivableRows, 'balanceDue'),
    customers: receivableRows.length,
    rows: sorted.slice(0, 6).map(function (row) {
      return {
        root: row.root,
        customerName: row.customerName,
        advisor: row.clientAdvisor,
        brand: row.brand,
        balanceDue: row.balanceDue,
        stageLabel: row.stageLabel,
        lastPaymentDate: row.lastTouch ? swAdminDashboardDateKey_(row.lastTouch) : '',
        masterUrl: row.masterUrl || ''
      };
    })
  };
}

function swAdminDashboardReadRootIndex_(ss, warnings) {
  var cached = swAdminDashboardCachedRootIndex_(ss);
  if (cached) return cached;

  var out = { available: false, byRoot: {}, cacheHit: false };
  var sh = ss.getSheetByName('07_Root_Index');
  if (!sh || sh.getLastRow() < 2) {
    warnings.push('Last-touch data unavailable: 07_Root_Index is missing or empty.');
    return out;
  }
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var C = {
    root: swPickIndex_(H, ['RootApptID', 'Root Appt ID', 'ROOT', 'Root_ID']),
    updatedAt: swPickIndex_(H, ['Updated At', 'UpdatedAt', 'Last Updated'])
  };
  if (C.root < 0 || C.updatedAt < 0) {
    warnings.push('Last-touch data unavailable: 07_Root_Index is missing RootApptID or Updated At.');
    return out;
  }
  var display = swAdminDashboardReadSelectedDisplayColumns_(sh, [C.root, C.updatedAt]);
  for (var i = 0; i < display.length; i++) {
    var root = swAdminDashboardCleanId_(swCell_(display[i], C.root));
    var when = swAdminDashboardDateTimeValue_(swCell_(display[i], C.updatedAt), swCell_(display[i], C.updatedAt));
    if (!root || !when) continue;
    if (!out.byRoot[root] || when.getTime() > out.byRoot[root].getTime()) out.byRoot[root] = when;
  }
  out.available = true;
  swAdminDashboardCacheRootIndex_(ss, out);
  return out;
}

function swAdminDashboardReadStatusLog_(ss, appointments, warnings, rootScope) {
  var cached = swAdminDashboardCachedStatusLog_(ss);
  if (cached) return cached;

  var out = { available: false, threeDByRoot: {}, productionByRoot: {}, cacheHit: false };
  var sh = ss.getSheetByName('03_Client_Status_Log');
  if (!sh || sh.getLastRow() < 2) {
    warnings.push('Status-log timing unavailable: 03_Client_Status_Log is missing or empty.');
    return out;
  }
  var apptToRoot = {};
  (appointments || []).forEach(function (rec) {
    if (rec.appt && rec.root) apptToRoot[rec.appt] = rec.root;
  });
  var targetRoots = {};
  Object.keys(rootScope || {}).forEach(function (root) {
    root = swAdminDashboardCleanId_(root);
    if (root) targetRoots[root] = true;
  });
  var hasTargetRoots = Object.keys(targetRoots).length > 0;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var C = {
    root: swPickIndex_(H, ['RootApptID', 'Root Appt ID', 'ROOT']),
    appt: swPickIndex_(H, ['APPT_ID', 'Appt ID', 'APPTID', 'Appointment ID']),
    custom: swPickIndex_(H, ['Custom Order Status', 'COS', 'Status', 'Order Status']),
    inProduction: swPickIndex_(H, ['In Production Status', 'IPS']),
    updatedAt: swPickIndex_(H, ['Updated At', 'UpdatedAt', 'Timestamp'])
  };
  if (C.updatedAt < 0 || (C.root < 0 && C.appt < 0)) {
    warnings.push('Status-log timing unavailable: 03_Client_Status_Log is missing an appointment/root id or Updated At.');
    return out;
  }
  var display = swAdminDashboardReadSelectedDisplayColumns_(sh, [C.root, C.appt, C.custom, C.inProduction, C.updatedAt]);
  for (var i = 0; i < display.length; i++) {
    var root = swAdminDashboardCleanId_(swCell_(display[i], C.root));
    var appt = swAdminDashboardCleanId_(swCell_(display[i], C.appt));
    root = root || apptToRoot[appt] || '';
    if (!root) continue;
    if (hasTargetRoots && !targetRoots[root]) continue;
    var when = swAdminDashboardDateTimeValue_(swCell_(display[i], C.updatedAt), swCell_(display[i], C.updatedAt));
    if (!when) continue;
    var custom = swNorm_(swCell_(display[i], C.custom));
    var ips = swTrim_(swCell_(display[i], C.inProduction));

    if (custom) {
      var d3 = out.threeDByRoot[root] || { requestDate: null, resolveDate: null, pending: false };
      if (/^3d requested|^3d revision requested/.test(custom)) {
        if (!d3.requestDate || when.getTime() > d3.requestDate.getTime()) d3.requestDate = when;
      }
      if (/^3d received|^3d waiting approval|^3d approved|approved for production|waiting production timeline|in production|order completed/.test(custom)) {
        if (!d3.resolveDate || when.getTime() > d3.resolveDate.getTime()) d3.resolveDate = when;
      }
      d3.pending = !!(d3.requestDate && (!d3.resolveDate || d3.resolveDate.getTime() < d3.requestDate.getTime()));
      out.threeDByRoot[root] = d3;
    }

    if (ips) {
      var prod = out.productionByRoot[root] || { status: '', updatedAt: null };
      if (!prod.updatedAt || when.getTime() > prod.updatedAt.getTime()) {
        prod.status = ips;
        prod.updatedAt = when;
      }
      out.productionByRoot[root] = prod;
    }
  }
  out.available = true;
  swAdminDashboardCacheStatusLog_(ss, out);
  return out;
}

function swAdminDashboardCachedRootIndex_(ss) {
  try {
    var cached = CacheService.getScriptCache().get(swAdminDashboardAuxCacheKey_(ss, 'rootIndex'));
    var parsed = cached ? swParseJson_(cached, null) : null;
    if (!parsed || !parsed.byRoot) return null;
    var out = { available: !!parsed.available, byRoot: {}, cacheHit: true };
    Object.keys(parsed.byRoot || {}).forEach(function (root) {
      var ms = Number(parsed.byRoot[root]);
      if (isFinite(ms) && ms > 0) out.byRoot[root] = new Date(ms);
    });
    return out.available ? out : null;
  } catch (_) {}
  return null;
}

function swAdminDashboardCacheRootIndex_(ss, rootIndex) {
  try {
    var byRoot = {};
    Object.keys((rootIndex && rootIndex.byRoot) || {}).forEach(function (root) {
      byRoot[root] = swAdminDashboardDateMs_(rootIndex.byRoot[root]);
    });
    swAdminDashboardPutAuxCache_(ss, 'rootIndex', { available: !!(rootIndex && rootIndex.available), byRoot: byRoot });
  } catch (_) {}
}

function swAdminDashboardCachedStatusLog_(ss) {
  try {
    var cached = CacheService.getScriptCache().get(swAdminDashboardAuxCacheKey_(ss, 'statusLog'));
    var parsed = cached ? swParseJson_(cached, null) : null;
    if (!parsed || !parsed.available) return null;
    var out = { available: true, threeDByRoot: {}, productionByRoot: {}, cacheHit: true };
    Object.keys(parsed.threeDByRoot || {}).forEach(function (root) {
      var item = parsed.threeDByRoot[root] || {};
      out.threeDByRoot[root] = {
        requestDate: item.requestDate ? new Date(Number(item.requestDate)) : null,
        resolveDate: item.resolveDate ? new Date(Number(item.resolveDate)) : null,
        pending: !!item.pending
      };
    });
    Object.keys(parsed.productionByRoot || {}).forEach(function (root) {
      var item = parsed.productionByRoot[root] || {};
      out.productionByRoot[root] = {
        status: item.status || '',
        updatedAt: item.updatedAt ? new Date(Number(item.updatedAt)) : null
      };
    });
    return out;
  } catch (_) {}
  return null;
}

function swAdminDashboardCacheStatusLog_(ss, statusLog) {
  try {
    var threeDByRoot = {};
    Object.keys((statusLog && statusLog.threeDByRoot) || {}).forEach(function (root) {
      var item = statusLog.threeDByRoot[root] || {};
      threeDByRoot[root] = {
        requestDate: swAdminDashboardDateMs_(item.requestDate),
        resolveDate: swAdminDashboardDateMs_(item.resolveDate),
        pending: !!item.pending
      };
    });
    var productionByRoot = {};
    Object.keys((statusLog && statusLog.productionByRoot) || {}).forEach(function (root) {
      var item = statusLog.productionByRoot[root] || {};
      productionByRoot[root] = {
        status: item.status || '',
        updatedAt: swAdminDashboardDateMs_(item.updatedAt)
      };
    });
    swAdminDashboardPutAuxCache_(ss, 'statusLog', {
      available: !!(statusLog && statusLog.available),
      threeDByRoot: threeDByRoot,
      productionByRoot: productionByRoot
    });
  } catch (_) {}
}

function swAdminDashboardPutAuxCache_(ss, name, payload) {
  try {
    var text = swStringify_(payload);
    if (text.length < 90000) CacheService.getScriptCache().put(swAdminDashboardAuxCacheKey_(ss, name), text, SW_ADMIN_DASHBOARD_AUX_CACHE_SECONDS);
  } catch (_) {}
}

function swAdminDashboardAuxCacheKey_(ss, name) {
  return 'sw:adminDashboardAux:v2:' + ss.getId() + ':' + name;
}

function swAdminDashboardDateMs_(value) {
  return value instanceof Date && !isNaN(value.getTime()) ? value.getTime() : 0;
}

function swAdminDashboardStageWeights_(ss) {
  var weights = {};
  Object.keys(SW_ADMIN_DASHBOARD_STAGE_WEIGHTS).forEach(function (key) {
    weights[key] = SW_ADMIN_DASHBOARD_STAGE_WEIGHTS[key];
  });
  var sh = ss.getSheetByName('00_Dashboard');
  if (!sh) return weights;
  try {
    var values = sh.getRange('AS30:AT60').getValues();
    values.forEach(function (row) {
      var label = swTrim_(row[0]);
      var value = Number(row[1]);
      if (!label || !isFinite(value)) return;
      var key = swAdminDashboardStageKeyFromLabel_(label);
      if (key) weights[key] = value;
    });
  } catch (_) {}
  return weights;
}

function swAdminDashboardStageKeyFromLabel_(label) {
  var s = swNorm_(label);
  if (/lost/.test(s)) return 'lost';
  if (/won|completed/.test(s)) return 'won';
  if (/production/.test(s)) return 'inProduction';
  if (/deposit|order/.test(s)) return 'deposit';
  if (/appointment|viewing|consult/.test(s)) return 'appointment';
  if (/follow/.test(s)) return 'followUp';
  if (/hot/.test(s)) return 'hotLead';
  if (/lead/.test(s)) return 'lead';
  return '';
}

function swAdminDashboardFirstVisit_(rows, tz) {
  var best = null;
  (rows || []).forEach(function (rec) {
    var visit = swVisitDateTime_(rec, tz);
    if (visit && (!best || visit.getTime() < best.getTime())) best = visit;
  });
  return best;
}

function swAdminDashboardLatestVisit_(rows, tz) {
  var best = null;
  (rows || []).forEach(function (rec) {
    var visit = swVisitDateTime_(rec, tz);
    if (visit && (!best || visit.getTime() > best.getTime())) best = visit;
  });
  return best;
}

function swAdminDashboardBudgetMidpoint_(rec) {
  var min = swAdminDashboardNumberOrBlank_(rec.budgetMin);
  var max = swAdminDashboardNumberOrBlank_(rec.budgetMax);
  if (min !== '' && max !== '') return (Number(min) + Number(max)) / 2;
  if (max !== '') return Number(max);
  if (min !== '') return Number(min);
  return 0;
}

function swAdminDashboardBestNumber_(values) {
  for (var i = 0; i < (values || []).length; i++) {
    if (values[i] === 0) return 0;
    var n = swAdminDashboardNumberOrBlank_(values[i]);
    if (n !== '') return Number(n);
  }
  return 0;
}

function swAdminDashboardSum_(rows, key) {
  return (rows || []).reduce(function (sum, row) {
    return sum + (Number(row[key]) || 0);
  }, 0);
}

function swAdminDashboardCountWaiting_(rows) {
  return (rows || []).filter(function (row) {
    return row.d3AgeDays != null && row.d3AgeDays > 3;
  }).length;
}

function swAdminDashboardCurrency_(value) {
  var n = Number(value || 0);
  return '$' + Math.round(n).toLocaleString();
}

function swAdminDashboardCalendarDays_(start, end) {
  if (!start || !end) return null;
  var a = swAdminDashboardStartOfDay_(start).getTime();
  var b = swAdminDashboardStartOfDay_(end).getTime();
  return Math.max(0, Math.round((b - a) / 86400000));
}

function swAdminDashboardMedianLeadToDeposit_(rows, filters) {
  var diffs = [];
  (rows || []).forEach(function (row) {
    if (!row.firstVisit || !row.firstDepositDate) return;
    if (!swAdminDashboardInRange_(row.firstDepositDate, filters)) return;
    diffs.push(swAdminDashboardCalendarDays_(row.firstVisit, row.firstDepositDate));
  });
  return swAdminDashboardMedian_(diffs);
}

function swAdminDashboardMedian_(values) {
  var nums = (values || []).filter(function (n) {
    return n != null && isFinite(Number(n));
  }).map(Number).sort(function (a, b) { return a - b; });
  if (!nums.length) return null;
  var mid = Math.floor(nums.length / 2);
  return nums.length % 2 ? nums[mid] : Math.round((nums[mid - 1] + nums[mid]) / 2);
}

function swAdminDashboardWinRate90_(rows) {
  var cutoff = swAdminDashboardAddDays_(swAdminDashboardStartOfDay_(new Date()), -89);
  var won = 0;
  var closed = 0;
  (rows || []).forEach(function (row) {
    if (row.stageKey !== 'won' && row.stageKey !== 'lost') return;
    var when = row.lastTouch || row.latestVisit || row.firstDepositDate;
    if (when && when.getTime() < cutoff.getTime()) return;
    closed++;
    if (row.stageKey === 'won') won++;
  });
  return {
    available: closed > 0,
    value: closed ? (won / closed) * 100 : 0,
    won: won,
    closed: closed
  };
}

function swAdminDashboardStartOfWeek_(date) {
  var d = swAdminDashboardStartOfDay_(date);
  var day = d.getDay();
  var diff = day === 0 ? -6 : 1 - day;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + diff);
}

function swAdminDashboardEndOfWeek_(date) {
  return swAdminDashboardEndOfDay_(swAdminDashboardAddDays_(swAdminDashboardStartOfWeek_(date), 6));
}

function swAdminDashboardAddDays_(date, days) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() + days);
}

function swAdminDashboardBookingsInRange_(appointments, filters, start, end) {
  var count = 0;
  (appointments || []).forEach(function (rec) {
    if (!swAdminDashboardRecordMatchesOwnerFilters_(rec, filters)) return;
    var bookedAt = swAdminDashboardDateTimeValue_(rec.bookedAtRaw, rec.bookedAt);
    if (bookedAt && bookedAt.getTime() >= start.getTime() && bookedAt.getTime() <= end.getTime() && !swTrim_(rec.rescheduledFromUid)) {
      count++;
    }
  });
  return count;
}

function swAdminDashboardFirstDepositsInRange_(payments, start, end) {
  var count = 0;
  var firstByKey = payments.firstByKey || {};
  Object.keys(firstByKey).forEach(function (key) {
    var item = firstByKey[key];
    if (item && item.when && item.when.getTime() >= start.getTime() && item.when.getTime() <= end.getTime()) {
      count++;
    }
  });
  return count;
}

function swAdminDashboardKanban_(ss, appointments, payments, filters, indexes) {
  indexes = indexes || {};
  var groups = indexes.groups || swAdminDashboardRowsByRoot_(appointments);
  var currentByRoot = indexes.currentByRoot || swAdminDashboardCurrentRowsByRoot_(appointments);
  var master = ss.getSheetByName(SW_SHEETS.MASTER);
  var masterGid = master ? master.getSheetId() : '';
  var columnsByKey = {};
  SW_ADMIN_DASHBOARD_COLUMNS.forEach(function (col) {
    columnsByKey[col.key] = {
      key: col.key,
      label: col.label,
      count: 0,
      hiddenCount: 0,
      cards: []
    };
  });

  Object.keys(currentByRoot).forEach(function (root) {
    var rec = currentByRoot[root];
    if (!swIsAppointmentActive_(rec) && !filters.includeClosed) return;
    if (!swAdminDashboardRecordMatchesOwnerFilters_(rec, filters)) return;
    var stage = swAdminDashboardPipelineStage_(rec, groups[root] || []);
    if (!filters.includeClosed && (stage.key === 'won' || stage.key === 'lost')) return;
    var col = columnsByKey[stage.key] || columnsByKey.lead;
    var card = swAdminDashboardCustomerCard_(ss, masterGid, root, rec, groups[root] || [], stage, payments);
    col.count++;
    if (col.cards.length < 60) col.cards.push(card);
    else col.hiddenCount++;
  });

  return {
    columns: SW_ADMIN_DASHBOARD_COLUMNS.map(function (col) { return columnsByKey[col.key]; })
  };
}

function swAdminDashboardRowsByRoot_(appointments) {
  var groups = {};
  (appointments || []).forEach(function (rec) {
    var root = swAdminDashboardCleanId_(rec.root || rec.appt);
    if (!root) return;
    if (!groups[root]) groups[root] = [];
    groups[root].push(rec);
  });
  return groups;
}

function swAdminDashboardCurrentRowsByRoot_(appointments) {
  var groups = swAdminDashboardRowsByRoot_(appointments);
  var out = {};
  Object.keys(groups).forEach(function (root) {
    var rows = groups[root];
    var active = rows.filter(function (rec) { return swIsAppointmentActive_(rec); });
    out[root] = swAdminDashboardLatestRow_(active.length ? active : rows);
  });
  return out;
}

function swAdminDashboardLatestRow_(rows) {
  var tz = swTimezone_();
  return (rows || []).reduce(function (best, rec) {
    if (!best) return rec;
    var bv = swAdminDashboardRowSortValue_(best, tz);
    var rv = swAdminDashboardRowSortValue_(rec, tz);
    return rv >= bv ? rec : best;
  }, null);
}

function swAdminDashboardRowSortValue_(rec, tz) {
  var visit = swVisitDateTime_(rec, tz);
  if (visit) return visit.getTime();
  return Number(rec.row || 0);
}

function swAdminDashboardPipelineStage_(rec, rootRows) {
  var sales = swNorm_(rec.salesStage);
  var conv = swNorm_(rec.convStatus);
  var custom = swNorm_(rec.customOrder);
  var inProd = swNorm_(rec.inProduction);
  var combined = [sales, conv, custom, inProd].join(' ');

  if (/lost/.test(combined)) return { key: 'lost', label: 'Lost Lead' };
  if (/won/.test(combined) || /order completed/.test(custom) || /production completed/.test(inProd)) {
    return { key: 'won', label: 'Won / Completed' };
  }
  if (custom === 'in production' || (inProd && !/none|n\/a|na/.test(inProd))) {
    return { key: 'inProduction', label: 'In Production' };
  }
  if (/deposit|confirmed order|order in progress|approved for production|waiting production|3d requested|3d revision|3d received/.test(combined)) {
    return { key: 'deposit', label: 'Deposit / Order In Progress' };
  }
  if (/appointment|viewing scheduled|scheduled/.test(combined) || swAdminDashboardHasFutureVisit_(rootRows)) {
    return { key: 'appointment', label: 'Appointment / Viewing Scheduled' };
  }
  if (/follow/.test(combined)) return { key: 'followUp', label: 'Follow-Up' };
  if (/hot/.test(combined)) return { key: 'hotLead', label: 'Hot Lead' };
  return { key: 'lead', label: 'Lead' };
}

function swAdminDashboardHasFutureVisit_(rows) {
  var tz = swTimezone_();
  var today = swAdminDashboardStartOfDay_(new Date()).getTime();
  for (var i = 0; i < (rows || []).length; i++) {
    if (!swIsAppointmentActive_(rows[i])) continue;
    var visit = swVisitDateTime_(rows[i], tz);
    if (visit && visit.getTime() >= today) return true;
  }
  return false;
}

function swAdminDashboardCustomerCard_(ss, masterGid, root, rec, rootRows, stage, payments) {
  var visit = swAdminDashboardVisitSummary_(rootRows);
  var pay = (payments.byRoot && payments.byRoot[root]) ||
    (rec.so && payments.bySo && payments.bySo[swAdminDashboardCleanId_(rec.so)]) ||
    {};
  return {
    root: root,
    appt: rec.appt || '',
    row: rec.row || '',
    customerName: rec.name || '',
    brand: rec.brand || '',
    clientAdvisor: rec.assignedRep || '',
    joc: rec.assistedRep || '',
    visitType: rec.visitType || '',
    visitDate: rec.visitDate || '',
    visitTime: rec.visitTime || '',
    nextVisit: visit.next || '',
    lastVisit: visit.last || '',
    salesStage: rec.salesStage || '',
    conversionStatus: rec.convStatus || '',
    customOrderStatus: rec.customOrder || '',
    inProductionStatus: rec.inProduction || '',
    stageKey: stage.key,
    stageLabel: stage.label,
    so: rec.so || '',
    nextSteps: rec.nextSteps || '',
    paymentCount: pay.paymentCount || 0,
    paidNet: pay.paidNet || 0,
    balanceDue: pay.balanceDue === 0 ? 0 : (pay.balanceDue || ''),
    orderTotal: pay.orderTotal === 0 ? 0 : (pay.orderTotal || ''),
    lastPaymentDate: pay.lastPaymentDate || '',
    source: rec.source || '',
    budgetMin: rec.budgetMin || '',
    budgetMax: rec.budgetMax || '',
    remainingBalance: rec.remainingBalance || '',
    orderDate: rec.orderDate || '',
    updatedAt: rec.updatedAt || '',
    clientFolder: rec.clientFolder || '',
    reportUrl: rec.reportUrl || '',
    quotationUrl: rec.quotationUrl || '',
    tracker3dUrl: rec.tracker3dUrl || '',
    orderFolder: rec.orderFolder || '',
    masterUrl: masterGid && rec.row ? ('https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + masterGid + '&range=A' + rec.row) : ''
  };
}

function swAdminDashboardVisitSummary_(rows) {
  var tz = swTimezone_();
  var now = new Date().getTime();
  var next = null;
  var last = null;
  (rows || []).forEach(function (rec) {
    var visit = swVisitDateTime_(rec, tz);
    if (!visit) return;
    if (visit.getTime() >= now && (!next || visit.getTime() < next.visit.getTime())) {
      next = { visit: visit, rec: rec };
    }
    if (visit.getTime() < now && (!last || visit.getTime() > last.visit.getTime())) {
      last = { visit: visit, rec: rec };
    }
  });
  return {
    next: next ? swAdminDashboardVisitLabel_(next.rec, next.visit) : '',
    last: last ? swAdminDashboardVisitLabel_(last.rec, last.visit) : ''
  };
}

function swAdminDashboardVisitLabel_(rec, visit) {
  return [
    swAdminDashboardDateKey_(visit),
    rec.visitTime || '',
    rec.visitType || ''
  ].filter(Boolean).join(' ');
}
