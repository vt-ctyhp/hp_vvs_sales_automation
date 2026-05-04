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

    filters = swAdminDashboardNormalizeFilters_(filters);
    var appointments = swReadAppointments_(ss);
    var scope = swAdminDashboardBuildScope_(appointments, filters);
    var warnings = [];
    var payments = swAdminDashboardReadPayments_(scope, filters, warnings);
    var state = swReadTaskListState_(ss, true);
    var tasks = swListVisibleTasksFromState_(state, user, 'admin');
    var currentByRoot = swAdminDashboardCurrentRowsByRoot_(appointments);

    return {
      ok: true,
      generatedAt: swIso_(new Date()),
      filters: swAdminDashboardPublicFilters_(filters),
      filterOptions: swAdminDashboardFilterOptions_(ss, appointments, filters),
      metrics: swAdminDashboardMetrics_(appointments, tasks, currentByRoot, payments, filters),
      kanban: swAdminDashboardKanban_(ss, appointments, payments, filters),
      tasks: tasks,
      warnings: warnings
    };
  });
}

function swAdminDashboardNormalizeFilters_(filters) {
  filters = filters || {};
  var week = swAdminDashboardDefaultWeek_();
  var start = swAdminDashboardParseDate_(filters.startDate) || week.start;
  var end = swAdminDashboardParseDate_(filters.endDate) || week.end;
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
    brand: swTrim_(filters.brand),
    clientAdvisor: swTrim_(filters.clientAdvisor),
    joc: swTrim_(filters.joc),
    includeClosed: filters.includeClosed === true || String(filters.includeClosed || '').toLowerCase() === 'true'
  };
}

function swAdminDashboardPublicFilters_(filters) {
  return {
    startDate: filters.startDate,
    endDate: filters.endDate,
    brand: filters.brand || '',
    clientAdvisor: filters.clientAdvisor || '',
    joc: filters.joc || '',
    includeClosed: !!filters.includeClosed
  };
}

function swAdminDashboardDefaultWeek_() {
  var now = new Date();
  var start = swAdminDashboardStartOfDay_(now);
  var day = start.getDay();
  var diff = day === 0 ? -6 : 1 - day;
  start = new Date(start.getFullYear(), start.getMonth(), start.getDate() + diff);
  var end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6, 23, 59, 59, 999);
  return { start: start, end: end };
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

  var values = sh.getRange(1, 1, lr, lc).getValues();
  var display = sh.getRange(1, 1, lr, lc).getDisplayValues();
  var headers = display[0].map(function (h) { return swTrim_(h); });
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

  var receipts = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var drow = display[i];
    var docType = swTrim_(swCell_(drow, C.docType));
    if (!/receipt/i.test(docType)) continue;
    var status = swNorm_(swCell_(drow, C.docStatus));
    if (/void|replaced|cancel|draft|deleted/.test(status)) continue;

    var root = swAdminDashboardCleanId_(swCell_(drow, C.root));
    var so = swAdminDashboardCleanId_(swCell_(drow, C.so));
    var brand = swTrim_(swCell_(drow, C.brand));
    if (!swAdminDashboardPaymentInScope_(root, so, brand, scope, filters)) continue;

    var when = swAdminDashboardDateTimeValue_(swCell_(row, C.when), swCell_(drow, C.when));
    if (!when) continue;

    var net = swAdminDashboardNumber_(C.amountNet >= 0 ? swCell_(row, C.amountNet) : swCell_(row, C.amountGross));
    var balance = C.balance >= 0 ? swAdminDashboardNumberOrBlank_(swCell_(row, C.balance)) : '';
    var orderTotal = C.orderTotal >= 0 ? swAdminDashboardNumberOrBlank_(swCell_(row, C.orderTotal)) : '';
    var key = root || so;
    if (!key) continue;

    receipts.push({ root: root, so: so, key: key, when: when, net: net, balance: balance, orderTotal: orderTotal });

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
    if (swAdminDashboardInRange_(item.when, filters)) {
      out.firstDepositCount++;
      out.firstDepositNet += item.net;
    }
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

function swAdminDashboardKanban_(ss, appointments, payments, filters) {
  var groups = swAdminDashboardRowsByRoot_(appointments);
  var currentByRoot = swAdminDashboardCurrentRowsByRoot_(appointments);
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
