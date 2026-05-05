/**
 * Admin dashboard read API for the Sales Workflow web app.
 * This module is read-only: it aggregates Master appointment rows and the
 * external payments ledger when configured.
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

var SW_ADMIN_DASHBOARD_PAYMENTS_CACHE_SECONDS = 2 * 60;
var SW_ADMIN_DASHBOARD_PAYMENTS_MEMORY_CACHE_ = {};

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
    mark('normalize');
    var appointments = swReadAppointments_(ss);
    mark('appointments', { rows: appointments.length });
    var scope = swAdminDashboardBuildScope_(appointments, filters);
    var warnings = [];
    var payments = swAdminDashboardReadPayments_(scope, filters, warnings);
    mark('payments', { receipts: payments.receipts ? payments.receipts.length : 0, cacheHit: !!payments.cacheHit });
    var indexes = swAdminDashboardBuildIndexes_(appointments);
    mark('indexes', { roots: Object.keys(indexes.currentByRoot || {}).length });
    var kanban = swAdminDashboardKanban_(ss, appointments, payments, filters, indexes);
    mark('kanban', { included: true });
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
      kanban: kanban,
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
    windowPreset: filters.windowPreset || 'last7',
    windowLabel: filters.windowLabel || '',
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
    bySo: {},
    cacheHit: false
  };

  var source = swAdminDashboardReadPaymentReceiptRows_(warnings);
  out.cacheHit = !!(source && source.cacheHit);
  var values = source && source.rows ? source.rows : [];
  var receipts = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var root = row.root || '';
    var so = row.so || '';
    var brand = row.brand || '';
    if (!swAdminDashboardPaymentInScope_(root, so, brand, scope, filters)) continue;

    var when = new Date(Number(row.whenMs || 0));
    if (isNaN(when.getTime())) continue;
    var net = Number(row.net || 0);
    var balance = row.balance === '' || row.balance == null ? '' : Number(row.balance);
    var orderTotal = row.orderTotal === '' || row.orderTotal == null ? '' : Number(row.orderTotal);
    var key = row.key || root || so;
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

function swAdminDashboardReadPaymentReceiptRows_(warnings) {
  var target = null;
  try {
    target = swAdminDashboardPaymentsSheet_();
  } catch (err) {
    warnings.push('Payments ledger unavailable: ' + (err && err.message ? err.message : err));
    return { rows: [], cacheHit: false };
  }
  if (!target || !target.sh) {
    warnings.push('Payments ledger unavailable.');
    return { rows: [], cacheHit: false };
  }

  var cacheKey = swAdminDashboardPaymentsCacheKey_(target);
  var cached = swAdminDashboardCachedPaymentReceiptRows_(cacheKey);
  if (cached !== null) return { rows: cached, cacheHit: true };

  var sh = target.sh;
  var lr = sh.getLastRow();
  var lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return { rows: [], cacheHit: false };

  var headers = sh.getRange(1, 1, 1, lc).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var C = {
    root: swPickIndex_(H, ['RootApptID', 'APPT_ID', 'Root Appt ID', 'Appointment ID']),
    so: swPickIndex_(H, ['SO#', 'SO', 'SO Number', 'Sales Order', 'Sales Order #']),
    brand: swPickIndex_(H, ['Brand']),
    paymentId: swPickIndex_(H, ['PAYMENT_ID', 'PaymentId', 'Payment ID']),
    docType: swPickIndex_(H, ['DocType', 'Doc Type', 'Document Type', 'Type']),
    docNumber: swPickIndex_(H, ['DocNumber', 'Doc #', 'Document Number']),
    docStatus: swPickIndex_(H, ['DocStatus', 'Doc Status', 'Status']),
    when: swPickIndex_(H, ['PaymentDateTime', 'Payment DateTime', 'Payment Date/Time', 'Payment Date', 'Paid At']),
    method: swPickIndex_(H, ['Method', 'Payment Method', 'Tender']),
    amountNet: swPickIndex_(H, ['AmountNet', 'Net', 'Net Amount']),
    amountGross: swPickIndex_(H, ['AmountGross', 'Gross', 'Amount']),
    balance: swPickIndex_(H, ['Balance_SO', 'Balance SO', 'BalanceDue', 'Balance Due']),
    orderTotal: swPickIndex_(H, ['Order_Total_SO', 'Order Total SO', 'OrderTotalValue', 'Order Total'])
  };
  if (C.docType < 0 || C.when < 0 || (C.amountNet < 0 && C.amountGross < 0)) {
    warnings.push('Payments ledger is missing DocType, PaymentDateTime, or amount columns.');
    return { rows: [], cacheHit: false };
  }

  var rowCount = lr - 1;
  var indexes = swAdminDashboardPaymentColumnIndexes_(C);
  var block = swAdminDashboardReadPaymentBlock_(sh, rowCount, indexes);
  var rows = [];
  for (var i = 0; i < block.rows.length; i++) {
    var row = block.rows[i];
    var docType = swTrim_(swAdminDashboardPaymentBlockCell_(row, C.docType, block.offset));
    if (!/receipt/i.test(docType)) continue;
    var status = swNorm_(swAdminDashboardPaymentBlockCell_(row, C.docStatus, block.offset));
    if (/void|replaced|cancel|draft|deleted/.test(status)) continue;
    var root = swAdminDashboardCleanId_(swAdminDashboardPaymentBlockCell_(row, C.root, block.offset));
    var so = swAdminDashboardCleanId_(swAdminDashboardPaymentBlockCell_(row, C.so, block.offset));
    var key = root || so;
    if (!key) continue;
    var whenRaw = swAdminDashboardPaymentBlockCell_(row, C.when, block.offset);
    var when = swAdminDashboardDateTimeValue_(whenRaw, whenRaw);
    if (!when) continue;
    var netRaw = C.amountNet >= 0
      ? swAdminDashboardPaymentBlockCell_(row, C.amountNet, block.offset)
      : swAdminDashboardPaymentBlockCell_(row, C.amountGross, block.offset);
    rows.push({
      root: root,
      so: so,
      key: key,
      brand: swTrim_(swAdminDashboardPaymentBlockCell_(row, C.brand, block.offset)),
      paymentId: swTrim_(swAdminDashboardPaymentBlockCell_(row, C.paymentId, block.offset)),
      docType: docType,
      docNumber: swTrim_(swAdminDashboardPaymentBlockCell_(row, C.docNumber, block.offset)),
      method: swTrim_(swAdminDashboardPaymentBlockCell_(row, C.method, block.offset)),
      whenMs: when.getTime(),
      net: swAdminDashboardNumber_(netRaw),
      gross: C.amountGross >= 0 ? swAdminDashboardNumber_(swAdminDashboardPaymentBlockCell_(row, C.amountGross, block.offset)) : swAdminDashboardNumber_(netRaw),
      balance: C.balance >= 0 ? swAdminDashboardNumberOrBlank_(swAdminDashboardPaymentBlockCell_(row, C.balance, block.offset)) : '',
      orderTotal: C.orderTotal >= 0 ? swAdminDashboardNumberOrBlank_(swAdminDashboardPaymentBlockCell_(row, C.orderTotal, block.offset)) : ''
    });
  }
  swAdminDashboardCachePaymentReceiptRows_(cacheKey, rows);
  return { rows: rows, cacheHit: false };
}

function swAdminDashboardReadPaymentBlock_(sh, rowCount, indexes) {
  var columns = swSelectedColumnIndexes_(indexes);
  var minCol = columns.length ? columns[0] : 0;
  var maxCol = columns.length ? columns[columns.length - 1] : 0;
  return {
    offset: minCol,
    rows: columns.length ? sh.getRange(2, minCol + 1, rowCount, maxCol - minCol + 1).getValues() : []
  };
}

function swAdminDashboardPaymentBlockCell_(row, idx, offset) {
  return idx >= 0 ? row[idx - offset] : '';
}

function swAdminDashboardPaymentColumnIndexes_(columns) {
  var out = [];
  Object.keys(columns || {}).forEach(function (key) {
    var col = Number(columns[key]);
    if (isFinite(col) && col >= 0) out.push(col);
  });
  return out;
}

function swAdminDashboardCachedPaymentReceiptRows_(cacheKey) {
  if (!cacheKey) return null;
  try {
    var memory = SW_ADMIN_DASHBOARD_PAYMENTS_MEMORY_CACHE_[cacheKey];
    if (memory && memory.expiresAt > new Date().getTime()) return memory.rows || [];
  } catch (_) {}
  try {
    var cached = CacheService.getScriptCache().get(cacheKey);
    var parsed = cached ? swParseJson_(cached, null) : null;
    if (!parsed || !Array.isArray(parsed.rows)) return null;
    SW_ADMIN_DASHBOARD_PAYMENTS_MEMORY_CACHE_[cacheKey] = {
      expiresAt: new Date().getTime() + SW_ADMIN_DASHBOARD_PAYMENTS_CACHE_SECONDS * 1000,
      rows: parsed.rows || []
    };
    return parsed.rows || [];
  } catch (_) {}
  return null;
}

function swAdminDashboardCachePaymentReceiptRows_(cacheKey, rows) {
  if (!cacheKey) return;
  rows = rows || [];
  try {
    SW_ADMIN_DASHBOARD_PAYMENTS_MEMORY_CACHE_[cacheKey] = {
      expiresAt: new Date().getTime() + SW_ADMIN_DASHBOARD_PAYMENTS_CACHE_SECONDS * 1000,
      rows: rows
    };
  } catch (_) {}
  try {
    var text = swStringify_({ cachedAt: swIso_(new Date()), rows: rows });
    if (text.length < 90000) CacheService.getScriptCache().put(cacheKey, text, SW_ADMIN_DASHBOARD_PAYMENTS_CACHE_SECONDS);
  } catch (_) {}
}

function swAdminDashboardPaymentsCacheKey_(target) {
  try {
    var sh = target && target.sh;
    if (!sh) return '';
    return 'sw:adminDashboardPayments:v2:' + sh.getParent().getId() + ':' + sh.getSheetId();
  } catch (_) {}
  return '';
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
