/**
 * Sales workflow public API: functions called by the web app, triggers, and admins.
 */

// Setup and generation.

/**
 * Mutating setup: ensures workflow sheets, styling, config rows, and templates exist.
 */
function sw_setupSalesWorkflow() {
  var ss = swSpreadsheet_();
  var taskSheet = swEnsureSheet_(ss, SW_SHEETS.TASKS, SW_TASK_HEADERS);
  var logSheet = swEnsureSheet_(ss, SW_SHEETS.LOG, SW_LOG_HEADERS);
  var configSheet = swEnsureSheet_(ss, SW_SHEETS.CONFIG, SW_CONFIG_HEADERS);
  var templateSheet = swEnsureSheet_(ss, SW_SHEETS.TEMPLATES, SW_TEMPLATE_HEADERS);
  var usersSheet = swEnsureSheet_(ss, SW_SHEETS.USERS, SW_AUTH_USER_HEADERS);

  swStyleSheet_(taskSheet);
  swStyleSheet_(logSheet);
  swStyleSheet_(configSheet);
  swStyleSheet_(templateSheet);
  swStyleSheet_(usersSheet);

  swSeedConfig_(configSheet);
  swSeedTemplates_(templateSheet);
  swSeedAuthUsers_(usersSheet);

  return {
    ok: true,
    sheets: SW_SHEETS,
    message: 'Sales workflow sheets are ready.'
  };
}

/**
 * Mutating generation: creates or updates workflow tasks from master appointments.
 */
function sw_generateSalesWorkflowTasks() {
  return swTimed_('sw_generateSalesWorkflowTasks', function () {
    var lock = LockService.getDocumentLock() || LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      sw_setupSalesWorkflow();

      var ss = swSpreadsheet_();
      var ctx = swBuildContext_(ss, true);
      if (swNorm_(swConfigValue_(ctx.config, 'SYSTEM', 'FEATURE_ENABLED', 'Y')) === 'n') {
        return {
          ok: true,
          generatedAt: swIso_(new Date()),
          scannedAppointments: 0,
          created: 0,
          updated: 0,
          blocked: 0,
          skippedOld: 0,
          systemCompleted: 0,
          paused: true
        };
      }
      var masterRows = swReadAppointments_(ss);
      var taskState = swReadTaskState_(ss);
      swBeginDeferredTaskWrites_(ss, taskState);
      var now = new Date();
      var summary = {
        ok: true,
        generatedAt: swIso_(now),
        scannedAppointments: masterRows.length,
        created: 0,
        updated: 0,
        blocked: 0,
        skippedOld: 0,
        systemCompleted: 0
      };

      masterRows.forEach(function (rec) {
        if (!rec.root && !rec.appt) return;
        if (!swIsWorkflowRelevant_(rec, now, ctx)) {
          summary.skippedOld++;
          return;
        }

        if (!swIsAppointmentActive_(rec)) {
          summary.blocked += swBlockTasksForAppointment_(ss, taskState, rec, 'Appointment is no longer active/current.');
          return;
        }

        swGenerateTasksForAppointment_(ss, taskState, ctx, rec, now, summary);
      });

      swFlushDeferredTaskWrites_(ss, taskState);
      return summary;
    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }
  });
}

/**
 * Editor-only, read-only duplicate task audit. Logs to Apps Script only.
 */
function sw_auditDuplicateTasks() {
  var ss = swSpreadsheet_();
  swRequireWorkflowReadSheets_(ss, { templates: false });
  var state = swReadTaskState_(ss, true, { includeDuplicates: true });
  var plan = swDuplicateTaskCleanupPlan_(state);
  var out = swDuplicateTaskAuditOutput_(state, plan);
  Logger.log('SW_DUPLICATE_TASK_AUDIT_SUMMARY ' + JSON.stringify(out.summary));
  Logger.log('SW_DUPLICATE_TASK_AUDIT_DETAILS ' + JSON.stringify(out, null, 2));
  return out;
}

/**
 * Editor-only dry run for duplicate task cleanup. Does not write.
 */
function sw_cleanupDuplicateTasksDryRun() {
  return swCleanupDuplicateTasks_(false);
}

/**
 * Editor-only cleanup for duplicate task rows. Marks extra pending rows Blocked.
 */
function sw_cleanupDuplicateTasksApply() {
  return swCleanupDuplicateTasks_(true);
}

/**
 * Mutating generation wrapper: refreshes owner assignment through the normal generator.
 */
function sw_refreshTaskOwners() {
  var summary = sw_generateSalesWorkflowTasks();
  summary.ownerRefresh = true;
  return summary;
}

/**
 * Mutating setup: replaces Sales Workflow generation and owner-refresh triggers.
 */
function sw_installSalesWorkflowTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    var fn = trigger.getHandlerFunction();
    if (fn === 'sw_generateSalesWorkflowTasks' || fn === 'sw_refreshTaskOwners') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('sw_generateSalesWorkflowTasks').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('sw_refreshTaskOwners').timeBased().everyDays(1).atHour(7).create();
  return {
    ok: true,
    message: 'Installed hourly task generation and daily 7am owner refresh.'
  };
}

// Read-only UI calls.

/**
 * Read-only UI bootstrap: returns current user, view counts, and initial My Queue tasks.
 */
function sw_getBootstrap(authToken) {
  return swTimed_('sw_getBootstrap', function () {
    var mark = swStepTimer_('sw_getBootstrap');
    var ss = swSpreadsheet_();
    mark('spreadsheet');
    swRequireWorkflowReadSheets_(ss, { templates: false });
    mark('requiredSheets');
    var user = swAuthUserForApi_(ss, authToken);
    mark('identity', { mode: authToken ? 'passwordSession' : 'appsScriptIdentity' });
    var state = swReadTaskListState_(ss, true);
    mark('taskListRead', { tasks: state.tasks.length });
    mark('currentUser', { isAdmin: user.isAdmin, isJoc: user.isJoc });
    var buckets = swBuildVisibleTaskBuckets_(state, user);
    mark('taskBuckets', {
      mine: buckets.mine.length,
      coverage: buckets.coverage.length,
      admin: buckets.admin.length
    });
    return {
      ok: true,
      user: user,
      tasks: buckets.mine,
      counts: {
        mine: buckets.mine.length,
        coverage: buckets.coverage.length,
        admin: user.isAdmin ? buckets.admin.length : 0
      },
      views: {
        mine: true,
        calendar: true,
        inStockDiamonds: true,
        diamondTracking: user.isAdmin || user.isDiamondOrderAdmin || user.isDiamondOrderAssistant,
        coverage: user.isJoc || user.isAdmin,
        admin: user.isAdmin
      },
      message: 'Connected. Use Generate Tasks to create or refresh the queue.'
    };
  });
}

/**
 * Read-only UI list: returns tasks visible to the current user for the requested view.
 */
function sw_getMyTasks(authToken, view) {
  return swTimed_('sw_getMyTasks', function () {
    var mark = swStepTimer_('sw_getMyTasks');
    if (!view && /^(mine|coverage|admin)$/i.test(String(authToken || ''))) {
      view = authToken;
      authToken = '';
    }
    var viewName = view || 'mine';
    var ss = swSpreadsheet_();
    mark('spreadsheet');
    swRequireWorkflowReadSheets_(ss, { templates: false });
    mark('requiredSheets');
    var user = swAuthUserForApi_(ss, authToken);
    mark('identity');
    mark('currentUser', { isAdmin: user.isAdmin, isJoc: user.isJoc });
    var state = swReadTaskListState_(ss, true);
    mark('taskListRead', { tasks: state.tasks.length });
    var tasks = swListVisibleTasksFromState_(state, user, viewName);
    mark('filter', { view: viewName, tasks: tasks.length });
    return {
      ok: true,
      view: viewName,
      user: user,
      tasks: tasks
    };
  });
}

/**
 * Read-only UI calendar: returns active upcoming appointments for one calendar month.
 */
function sw_getCalendarAppointments(authToken, monthKey) {
  return swTimed_('sw_getCalendarAppointments', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    swAuthUserForApi_(ss, authToken);

    var tz = swTimezone_();
    var month = swCalendarMonthRange_(monthKey);
    var today = new Date();
    var todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0, 0);
    var appointments = swReadAppointments_(ss).filter(function (rec) {
      if (!swIsAppointmentActive_(rec)) return false;
      var visitAt = swVisitDateTime_(rec, tz);
      if (!visitAt) return false;
      if (visitAt.getTime() < todayStart.getTime()) return false;
      return visitAt.getTime() >= month.start.getTime() && visitAt.getTime() < month.end.getTime();
    }).sort(function (a, b) {
      var av = swVisitDateTime_(a, tz);
      var bv = swVisitDateTime_(b, tz);
      return av.getTime() - bv.getTime() || String(a.name).localeCompare(String(b.name));
    }).map(function (rec) {
      var visitAt = swVisitDateTime_(rec, tz);
      return {
        id: ['CAL', rec.row, rec.root || rec.appt || ''].join('|'),
        row: rec.row,
        root: rec.root,
        appt: rec.appt,
        customerName: rec.name,
        brand: rec.brand,
        visitDate: rec.visitDate,
        visitTime: rec.visitTime,
        visitType: rec.visitType,
        dateKey: swDateKey_(visitAt),
        sortAt: swIso_(visitAt),
        assignedRep: rec.assignedRep,
        assistedRep: rec.assistedRep,
        status: rec.status,
        clientFolder: rec.clientFolder,
        reportUrl: rec.reportUrl,
        quotationUrl: rec.quotationUrl,
        tracker3dUrl: rec.tracker3dUrl,
        isDiamondViewing: swDiamondIsViewingAppointment_(rec)
      };
    });

    return {
      ok: true,
      monthKey: month.key,
      monthLabel: Utilities.formatDate(month.start, tz, 'MMMM yyyy'),
      prevMonthKey: swCalendarMonthKey_(new Date(month.start.getFullYear(), month.start.getMonth() - 1, 1)),
      nextMonthKey: swCalendarMonthKey_(new Date(month.start.getFullYear(), month.start.getMonth() + 1, 1)),
      todayKey: swDateKey_(todayStart),
      appointmentCount: appointments.length,
      appointments: appointments
    };
  });
}

/**
 * Read-only UI diamond tracking dashboard for diamond order roles.
 */
function sw_getDiamondTrackingDashboard(authToken) {
  return swTimed_('sw_getDiamondTrackingDashboard', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken);
    if (!(user.isAdmin || user.isDiamondOrderAdmin || user.isDiamondOrderAssistant)) {
      throw new Error('Diamond order access required.');
    }

    var target = swDiamond200Target_();
    if (!target || !target.sheet) {
      return { ok: true, available: false, rows: [], stats: {}, missingColumns: [] };
    }

    var sh = target.sheet;
    var lr = sh.getLastRow();
    var lc = sh.getLastColumn();
    if (lr < 3 || lc < 1) {
      return { ok: true, available: true, spreadsheetUrl: target.ss.getUrl(), tab: target.tab, rows: [], stats: {} };
    }

    var hm = swDiamond200HeaderMap_(sh);
    var C = {
      root: swDiamondFind200Column_(hm, ['RootApptID', 'APPT_ID', 'Root Appt ID']),
      customerName: swDiamondFind200Column_(hm, ['Customer Name', 'Client Name', 'Customer']),
      appointment: swDiamondFind200Column_(hm, ['Customer Appt Time & Date', 'Customer Appointment Date', 'Appointment Date']),
      assignedRep: swDiamondFind200Column_(hm, ['Assigned Rep', 'Sales Rep']),
      vendor: swDiamondFind200Column_(hm, ['Vendor']),
      shape: swDiamondFind200Column_(hm, ['Shape']),
      carat: swDiamondFind200Column_(hm, ['Carat']),
      color: swDiamondFind200Column_(hm, ['Color']),
      clarity: swDiamondFind200Column_(hm, ['Clarity']),
      certNo: swDiamondFind200Column_(hm, ['Certificate No', 'Cert #', 'Cert No', 'Certificate #']),
      orderStatus: swDiamondFind200Column_(hm, ['Order Status', 'OrderStatus']),
      stoneStatus: swDiamondFind200Column_(hm, ['Stone Status', 'StoneStatus']),
      decision: swDiamondFind200Column_(hm, ['Stone Decision (PO, Return)', 'Stone Decision', 'StoneDecision']),
      orderDate: swDiamondFind200Column_(hm, ['Purchased / Ordered Date', 'Purchased/Ordered Date', 'PurchasedOrderedDate']),
      returnDueDate: swDiamondFind200Column_(hm, ['Return DUE DATE', 'Return Due Date', 'Return Due']),
      trackingEta: swDiamondFind200Column_(hm, ['Tracking ETA', 'Tracking ETA Date', 'ETA Date', 'ETA']),
      trackingStatus: swDiamondFind200Column_(hm, ['Tracking Status', 'ETA Status', 'Shipment Status']),
      carrier: swDiamondFind200Column_(hm, ['Carrier', 'Shipping Carrier']),
      trackingNumber: swDiamondFind200Column_(hm, ['Tracking Number', 'Tracking #', 'Tracking No']),
      trackingUrl: swDiamondFind200Column_(hm, ['Tracking URL', 'Tracking Link'])
    };
    var missingColumns = [];
    if (!C.trackingEta) missingColumns.push('Tracking ETA');
    if (!C.trackingStatus) missingColumns.push('Tracking Status');

    var values = sh.getRange(3, 1, lr - 2, lc).getDisplayValues();
    var today = new Date();
    var todayMs = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
    var warningMs = todayMs + 7 * 24 * 60 * 60 * 1000;
    var rows = [];
    var stats = {
      total: 0,
      onTheWay: 0,
      delivered: 0,
      returns: 0,
      missingEta: 0,
      issues: 0
    };

    values.forEach(function (row, i) {
      var orderStatus = swDiamondCell_(row, C.orderStatus);
      var stoneStatus = swDiamondCell_(row, C.stoneStatus);
      var decision = swDiamondCell_(row, C.decision);
      var trackingEta = swDiamondCell_(row, C.trackingEta);
      var trackingStatus = swDiamondCell_(row, C.trackingStatus);
      var orderNorm = swNorm_(orderStatus);
      var stoneNorm = swNorm_(stoneStatus);
      var decisionNorm = swNorm_(decision);
      var trackingNorm = swNorm_(trackingStatus);
      var relevant = orderNorm === 'on the way' || orderNorm === 'delivered' ||
        decisionNorm === 'return' || trackingEta || trackingStatus ||
        stoneNorm.indexOf('return in progress') >= 0;
      if (!relevant) return;

      var etaMs = swDiamondDateValue_(trackingEta);
      var returnMs = swDiamondDateValue_(swDiamondCell_(row, C.returnDueDate));
      var issue = '';
      if (orderNorm === 'on the way' && !trackingEta) issue = 'Missing ETA';
      if (!issue && /(delay|unavailable|cancel|concern|problem)/.test(trackingNorm)) issue = 'Tracking concern';
      if (!issue && etaMs && etaMs < todayMs && orderNorm === 'on the way') issue = 'ETA overdue';
      if (!issue && decisionNorm === 'return' && returnMs && returnMs <= warningMs) issue = returnMs < todayMs ? 'Return overdue' : 'Return due soon';

      var out = {
        rowIndex: i + 3,
        root: swDiamondCell_(row, C.root),
        customerName: swDiamondCell_(row, C.customerName),
        appointment: swDiamondCell_(row, C.appointment),
        assignedRep: swDiamondCell_(row, C.assignedRep),
        vendor: swDiamondCell_(row, C.vendor),
        certNo: swDiamondCell_(row, C.certNo),
        diamond: [swDiamondCell_(row, C.shape), swDiamondCell_(row, C.carat), swDiamondCell_(row, C.color), swDiamondCell_(row, C.clarity)].filter(Boolean).join(' '),
        orderStatus: orderStatus,
        stoneStatus: stoneStatus,
        decision: decision,
        orderDate: swDiamondCell_(row, C.orderDate),
        returnDueDate: swDiamondCell_(row, C.returnDueDate),
        trackingEta: trackingEta,
        trackingStatus: trackingStatus,
        carrier: swDiamondCell_(row, C.carrier),
        trackingNumber: swDiamondCell_(row, C.trackingNumber),
        trackingUrl: swDiamondCell_(row, C.trackingUrl),
        issue: issue
      };
      rows.push(out);
      stats.total++;
      if (orderNorm === 'on the way') stats.onTheWay++;
      if (orderNorm === 'delivered') stats.delivered++;
      if (decisionNorm === 'return' || stoneNorm.indexOf('return in progress') >= 0) stats.returns++;
      if (orderNorm === 'on the way' && !trackingEta) stats.missingEta++;
      if (issue) stats.issues++;
    });

    rows.sort(function (a, b) {
      if (!!a.issue !== !!b.issue) return a.issue ? -1 : 1;
      var av = swDiamondDateValue_(a.trackingEta || a.returnDueDate) || 9999999999999;
      var bv = swDiamondDateValue_(b.trackingEta || b.returnDueDate) || 9999999999999;
      return av - bv;
    });

    return {
      ok: true,
      available: true,
      generatedAt: swIso_(new Date()),
      spreadsheetUrl: target.ss.getUrl(),
      spreadsheetName: target.ss.getName(),
      tab: target.tab,
      missingColumns: missingColumns,
      stats: stats,
      rows: rows.slice(0, 200)
    };
  });
}

/**
 * Read-only UI stock view: all workflow users can see in-store diamonds and
 * return due dates before proposing stones for a Diamond Viewing appointment.
 */
function sw_getInStockDiamonds(authToken) {
  return swTimed_('sw_getInStockDiamonds', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    swAuthUserForApi_(ss, authToken);

    var target = swDiamond200Target_();
    if (!target || !target.sheet) {
      return { ok: true, available: false, rows: [], stats: {} };
    }

    var sh = target.sheet;
    var lr = sh.getLastRow();
    var lc = sh.getLastColumn();
    if (lr < 3 || lc < 1) {
      return { ok: true, available: true, spreadsheetUrl: target.ss.getUrl(), tab: target.tab, rows: [], stats: {} };
    }

    var config = ss.getSheetByName(SW_SHEETS.CONFIG)
      ? swReadSheetObjectsExpectedHeaders_(ss.getSheetByName(SW_SHEETS.CONFIG), SW_CONFIG_HEADERS)
      : [];
    var returnWindow = Number(swConfigValue_(config, 'SYSTEM', 'DIAMOND_RETURN_WINDOW_DAYS', '30')) || 30;
    var hm = swDiamond200HeaderMap_(sh);
    var C = {
      root: swDiamondFind200Column_(hm, ['RootApptID', 'APPT_ID', 'Root Appt ID']),
      customerName: swDiamondFind200Column_(hm, ['Customer Name', 'Client Name', 'Customer']),
      appointment: swDiamondFind200Column_(hm, ['Customer Appt Time & Date', 'Customer Appointment Date', 'Appointment Date']),
      assignedRep: swDiamondFind200Column_(hm, ['Assigned Rep', 'Sales Rep']),
      company: swDiamondFind200Column_(hm, ['Company', 'Brand']),
      vendor: swDiamondFind200Column_(hm, ['Vendor']),
      stoneType: swDiamondFind200Column_(hm, ['Stone Type', 'StoneType']),
      shape: swDiamondFind200Column_(hm, ['Shape']),
      carat: swDiamondFind200Column_(hm, ['Carat']),
      color: swDiamondFind200Column_(hm, ['Color']),
      clarity: swDiamondFind200Column_(hm, ['Clarity']),
      lab: swDiamondFind200Column_(hm, ['LAB', 'Lab', 'Grading Lab']),
      certNo: swDiamondFind200Column_(hm, ['Certificate No', 'Cert #', 'Cert No', 'Certificate #']),
      measurement: swDiamondFind200Column_(hm, ['Measurements', 'Measurement', 'Meas.', 'Meas']),
      ratio: swDiamondFind200Column_(hm, ['L/W Ratio', 'L-W Ratio', 'LW Ratio', 'Ratio']),
      orderStatus: swDiamondFind200Column_(hm, ['Order Status', 'OrderStatus']),
      stoneStatus: swDiamondFind200Column_(hm, ['Stone Status', 'StoneStatus']),
      decision: swDiamondFind200Column_(hm, ['Stone Decision (PO, Return)', 'Stone Decision', 'StoneDecision']),
      orderDate: swDiamondFind200Column_(hm, ['Purchased / Ordered Date', 'Purchased/Ordered Date', 'PurchasedOrderedDate']),
      memoDate: swDiamondFind200Column_(hm, ['Memo/ Invoice Date', 'Memo / Invoice Date', 'Memo Invoice Date']),
      returnDueDate: swDiamondFind200Column_(hm, ['Return DUE DATE', 'Return Due Date', 'Return Due'])
    };

    var values = sh.getRange(3, 1, lr - 2, lc).getDisplayValues();
    var today = new Date();
    var todayMs = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
    var warningMs = todayMs + 7 * 24 * 60 * 60 * 1000;
    var rows = [];
    var stats = {
      total: 0,
      available: 0,
      returnSoon: 0,
      returnOverdue: 0,
      noReturnDate: 0
    };

    values.forEach(function (row, i) {
      var orderStatus = swDiamondCell_(row, C.orderStatus);
      var stoneStatus = swDiamondCell_(row, C.stoneStatus);
      var decision = swDiamondCell_(row, C.decision);
      var orderNorm = swNorm_(orderStatus);
      var stoneNorm = swNorm_(stoneStatus);
      var decisionNorm = swNorm_(decision);
      var isInStock = stoneNorm.indexOf('in stock') >= 0 || orderNorm === 'delivered';
      var unavailable = /return in progress|returned|sold|customer purchased/.test(stoneNorm) ||
        decisionNorm === 'purchase' || decisionNorm === 'purchased';
      if (!isInStock || unavailable) return;

      var returnDueDate = swDiamondCell_(row, C.returnDueDate) ||
        swDiamondReturnDueDate_(swDiamondCell_(row, C.orderDate), returnWindow);
      var returnMs = swDiamondDateValue_(returnDueDate);
      var issue = '';
      if (!returnDueDate) issue = 'No return date';
      else if (returnMs < todayMs) issue = 'Return overdue';
      else if (returnMs <= warningMs) issue = 'Return soon';

      var daysUntilReturn = returnMs ? Math.ceil((returnMs - todayMs) / (24 * 60 * 60 * 1000)) : '';
      rows.push({
        rowIndex: i + 3,
        root: swDiamondCell_(row, C.root),
        customerName: swDiamondCell_(row, C.customerName),
        appointment: swDiamondCell_(row, C.appointment),
        assignedRep: swDiamondCell_(row, C.assignedRep),
        company: swDiamondCell_(row, C.company),
        vendor: swDiamondCell_(row, C.vendor),
        stoneType: swDiamondCell_(row, C.stoneType),
        certNo: swDiamondCell_(row, C.certNo),
        diamond: [swDiamondCell_(row, C.shape), swDiamondCell_(row, C.carat), swDiamondCell_(row, C.color), swDiamondCell_(row, C.clarity)].filter(Boolean).join(' '),
        measurement: swDiamondCell_(row, C.measurement),
        ratio: swDiamondCell_(row, C.ratio),
        lab: swDiamondCell_(row, C.lab),
        orderStatus: orderStatus,
        stoneStatus: stoneStatus,
        decision: decision,
        orderDate: swDiamondCell_(row, C.orderDate),
        memoDate: swDiamondCell_(row, C.memoDate),
        returnDueDate: returnDueDate,
        daysUntilReturn: daysUntilReturn,
        issue: issue,
        availabilityLabel: returnDueDate ? ('Available until ' + returnDueDate) : 'Return date missing'
      });
      stats.total++;
      if (!issue) stats.available++;
      if (issue === 'Return soon') stats.returnSoon++;
      if (issue === 'Return overdue') stats.returnOverdue++;
      if (issue === 'No return date') stats.noReturnDate++;
    });

    rows.sort(function (a, b) {
      if (!!a.issue !== !!b.issue) return a.issue ? -1 : 1;
      var av = swDiamondDateValue_(a.returnDueDate) || 9999999999999;
      var bv = swDiamondDateValue_(b.returnDueDate) || 9999999999999;
      return av - bv || String(a.diamond).localeCompare(String(b.diamond));
    });

    return {
      ok: true,
      available: true,
      generatedAt: swIso_(new Date()),
      spreadsheetUrl: target.ss.getUrl(),
      spreadsheetName: target.ss.getName(),
      tab: target.tab,
      stats: stats,
      rows: rows.slice(0, 300)
    };
  });
}

function swCalendarMonthRange_(monthKey) {
  var now = new Date();
  var match = /^(\d{4})-(\d{2})$/.exec(swTrim_(monthKey));
  var year = match ? Number(match[1]) : now.getFullYear();
  var month = match ? Number(match[2]) - 1 : now.getMonth();
  var start = new Date(year, month, 1, 0, 0, 0, 0);
  var end = new Date(year, month + 1, 1, 0, 0, 0, 0);
  return {
    key: swCalendarMonthKey_(start),
    start: start,
    end: end
  };
}

function swCalendarMonthKey_(date) {
  return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0');
}

/**
 * Read-only admin list: returns filterable admin-visible tasks.
 */
function sw_adminGetTasks(authToken, filters) {
  return swTimed_('sw_adminGetTasks', function () {
    if (typeof authToken === 'object' && filters == null) {
      filters = authToken;
      authToken = '';
    }
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken);
    if (!user.isAdmin) throw new Error('Admin access required.');
    var state = swReadTaskListState_(ss, true);
    var tasks = swListVisibleTasksFromState_(state, user, 'admin');
    filters = filters || {};
    if (filters.status) {
      tasks = tasks.filter(function (t) { return t.status === filters.status; });
    }
    if (filters.ownerRole) {
      tasks = tasks.filter(function (t) { return t.ownerRole === filters.ownerRole; });
    }
    return { ok: true, tasks: tasks };
  });
}

/**
 * Read-only detail: returns task payload, rendered template data, and allowed actions.
 */
function sw_getTaskDetail(authToken, taskId) {
  return swTimed_('sw_getTaskDetail', function () {
    if (!taskId && /^SW\|/.test(String(authToken || ''))) {
      taskId = authToken;
      authToken = '';
    }
    var mark = swStepTimer_('sw_getTaskDetail');
    var ss = swSpreadsheet_();
    mark('spreadsheet');
    swRequireWorkflowReadSheets_(ss);
    mark('requiredSheets');
    var ctx = { templates: swReadTemplates_(ss, true) };
    mark('detailContext', { mode: authToken ? 'passwordSession' : 'appsScriptIdentity' });
    var user = swAuthUserForApi_(ss, authToken);
    mark('currentUser', { isAdmin: user.isAdmin, isJoc: user.isJoc });
    var task = swReadTaskRowById_(ss, taskId, true);
    mark('taskRowLookup');
    if (!task) throw new Error('Task not found: ' + taskId);
    if (!swCanViewTask_(task, user)) throw new Error('You do not have access to this task.');

    var payload = swParseJson_(task.payloadJson, {});
    mark('payloadParse');
    var template = ctx.templates[task.taskType] || swDefaultTemplate_(task.taskType);
    var renderData = swRenderDataForTask_(task, payload);
    var renderedTemplate = template.template ? swRenderTemplate_(template.template, renderData) : '';
    var renderedAttachmentUrl = template.attachmentUrl ? swRenderTemplate_(template.attachmentUrl, renderData) : '';
    var renderedAttachmentLabel = template.attachmentLabel ? swRenderTemplate_(template.attachmentLabel, renderData) : '';
    var attachments = swAttachmentsForTask_(task, template, renderData);
    var missingFields = swMissingFieldsForTask_(task, template, renderData);
    var checklist = swParseJson_(template.checklistJson, []);
    mark('render');

    return {
      ok: true,
      user: user,
      task: task,
      payload: payload,
      renderedTemplate: renderedTemplate,
      attachment: {
        label: renderedAttachmentLabel,
        url: renderedAttachmentUrl
      },
      attachments: attachments,
      missingFields: missingFields,
      checklist: checklist,
      canComplete: swCanActOnTask_(task, user),
      canClaim: swCanClaimTask_(task, user),
      canAdmin: user.isAdmin
    };
  });
}

// Task actions.

/**
 * Mutating task action: marks acknowledge-style tasks complete through the standard path.
 */
function sw_acknowledgeTask(authToken, taskId, data) {
  if (!taskId && /^SW\|/.test(String(authToken || ''))) {
    taskId = authToken;
    authToken = '';
  }
  if (typeof taskId === 'object' && data == null) {
    data = taskId;
    taskId = authToken;
    authToken = '';
  }
  data = data || {};
  data.acknowledged = true;
  return sw_completeTask(authToken, taskId, data);
}

/**
 * Mutating task action: validates and completes a pending task, then refreshes generation.
 */
function sw_completeTask(authToken, taskId, data) {
  if (!taskId && /^SW\|/.test(String(authToken || ''))) {
    taskId = authToken;
    authToken = '';
  }
  if (typeof taskId === 'object' && data == null) {
    data = taskId;
    taskId = authToken;
    authToken = '';
  }
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swAuthUserForApi_(ss, authToken);
  var task = swGetTaskById_(ss, taskId);
  if (!task) throw new Error('Task not found: ' + taskId);
  if (!swCanActOnTask_(task, user)) throw new Error('You are not the current owner for this task.');
  if (task.status !== SW_STATUSES.PENDING) throw new Error('Only pending tasks can be completed.');

  data = data || {};
  swValidateCompletion_(ss, task, data);
  var diamondAction = swDiamondHandleTaskCompletion_(ss, task, data, user);

  var template = swTemplateForType_(ss, task.taskType);
  var payload = swParseJson_(task.payloadJson, {});
  var renderData = swRenderDataForTask_(task, payload);
  var renderedTemplate = template.template ? swRenderTemplate_(template.template, renderData) : '';
  var renderedAttachments = swAttachmentsForTask_(task, template, renderData);
  payload.completion = data;
  payload.renderedTemplate = renderedTemplate;
  payload.renderedAttachments = renderedAttachments;
  payload.completedBy = user.name || user.email;
  payload.completedByEmail = user.email;
  payload.completedAt = swIso_(new Date());
  if (diamondAction) payload.diamondAction = diamondAction;

  var oldOwner = task.currentOwner;
  task.status = SW_STATUSES.COMPLETED;
  task.completedBy = user.name || user.email;
  task.completedByEmail = user.email;
  task.completedAt = payload.completedAt;
  task.updatedAt = payload.completedAt;
  task.lastEvent = 'COMPLETE';
  task.payloadJson = swStringify_(payload);
  swWriteTaskRow_(ss, task);
  swAppendTaskLog_(ss, 'COMPLETE', task, user, oldOwner, task.currentOwner, data);

  var generation = sw_generateSalesWorkflowTasks();
  return {
    ok: true,
    task: swGetTaskById_(ss, taskId),
    generation: generation
  };
}

/**
 * Mutating task action: lets an eligible user claim a pending coverage task.
 */
function sw_claimTask(authToken, taskId) {
  if (!taskId && /^SW\|/.test(String(authToken || ''))) {
    taskId = authToken;
    authToken = '';
  }
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swAuthUserForApi_(ss, authToken);
  var task = swGetTaskById_(ss, taskId);
  if (!task) throw new Error('Task not found: ' + taskId);
  if (!swCanClaimTask_(task, user)) throw new Error('This task is not available for you to claim.');

  var fromOwner = task.currentOwner;
  var now = swIso_(new Date());
  task.currentOwner = user.name || user.email;
  task.currentOwnerEmail = user.email;
  task.coverageReason = task.coverageReason || 'CLAIMED_FROM_COVERAGE';
  task.claimedBy = user.name || user.email;
  task.claimedAt = now;
  task.updatedAt = now;
  task.lastEvent = 'CLAIM';
  swWriteTaskRow_(ss, task);
  swAppendTaskLog_(ss, 'CLAIM', task, user, fromOwner, task.currentOwner, {});
  return { ok: true, task: swGetTaskById_(ss, taskId) };
}

/**
 * Mutating task action: records that the user copied a task template.
 */
function sw_logTemplateCopied(authToken, taskId) {
  if (!taskId && /^SW\|/.test(String(authToken || ''))) {
    taskId = authToken;
    authToken = '';
  }
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swAuthUserForApi_(ss, authToken);
  var task = swGetTaskById_(ss, taskId);
  if (!task) throw new Error('Task not found: ' + taskId);
  if (!swCanViewTask_(task, user)) throw new Error('You do not have access to this task.');
  swAppendTaskLog_(ss, 'TEMPLATE_COPY', task, user, task.currentOwner, task.currentOwner, {});
  return { ok: true };
}

// Admin actions.

/**
 * Mutating admin action: reassigns a pending task to a named owner.
 */
function sw_adminReassignTask(authToken, taskId, ownerName, ownerEmail, reason) {
  if (/^SW\|/.test(String(authToken || ''))) {
    reason = ownerEmail;
    ownerEmail = ownerName;
    ownerName = taskId;
    taskId = authToken;
    authToken = '';
  }
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swAuthUserForApi_(ss, authToken);
  if (!user.isAdmin) throw new Error('Admin access required.');

  var task = swGetTaskById_(ss, taskId);
  if (!task) throw new Error('Task not found: ' + taskId);
  if (task.status !== SW_STATUSES.PENDING) throw new Error('Only pending tasks can be reassigned.');

  var fromOwner = task.currentOwner;
  var now = swIso_(new Date());
  task.currentOwner = swTrim_(ownerName);
  task.currentOwnerEmail = swTrim_(ownerEmail).toLowerCase();
  task.coverageReason = swTrim_(reason) || 'ADMIN_REASSIGNED';
  task.updatedAt = now;
  task.lastEvent = 'REASSIGN';
  swWriteTaskRow_(ss, task);
  swAppendTaskLog_(ss, 'REASSIGN', task, user, fromOwner, task.currentOwner, {
    reason: reason || ''
  });
  return { ok: true, task: swGetTaskById_(ss, taskId) };
}

/**
 * Mutating admin action: blocks a task and records the reason.
 */
function sw_adminBlockTask(authToken, taskId, reason) {
  if (/^SW\|/.test(String(authToken || ''))) {
    reason = taskId;
    taskId = authToken;
    authToken = '';
  }
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swAuthUserForApi_(ss, authToken);
  if (!user.isAdmin) throw new Error('Admin access required.');

  var task = swGetTaskById_(ss, taskId);
  if (!task) throw new Error('Task not found: ' + taskId);

  var oldStatus = task.status;
  task.status = SW_STATUSES.BLOCKED;
  task.coverageReason = swTrim_(reason) || 'ADMIN_BLOCKED';
  task.updatedAt = swIso_(new Date());
  task.lastEvent = 'BLOCK';
  swWriteTaskRow_(ss, task);
  swAppendTaskLog_(ss, 'BLOCK', task, user, task.currentOwner, task.currentOwner, {
    fromStatus: oldStatus,
    reason: reason || ''
  });
  return { ok: true, task: swGetTaskById_(ss, taskId) };
}

/**
 * Mutating admin action: returns a blocked task to pending status.
 */
function sw_adminUnblockTask(authToken, taskId, reason) {
  if (/^SW\|/.test(String(authToken || ''))) {
    reason = taskId;
    taskId = authToken;
    authToken = '';
  }
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swAuthUserForApi_(ss, authToken);
  if (!user.isAdmin) throw new Error('Admin access required.');

  var task = swGetTaskById_(ss, taskId);
  if (!task) throw new Error('Task not found: ' + taskId);

  task.status = SW_STATUSES.PENDING;
  task.coverageReason = swTrim_(reason) || '';
  task.updatedAt = swIso_(new Date());
  task.lastEvent = 'UNBLOCK';
  swWriteTaskRow_(ss, task);
  swAppendTaskLog_(ss, 'UNBLOCK', task, user, task.currentOwner, task.currentOwner, {
    reason: reason || ''
  });
  return { ok: true, task: swGetTaskById_(ss, taskId) };
}

// Diagnostics and tests.

/**
 * Read-only setup review for the login + Diamond Viewing workflow rollout.
 */
function sw_reviewDiamondWorkflowSetup() {
  var ss = swSpreadsheet_();
  var out = {
    ok: true,
    generatedAt: swIso_(new Date()),
    sheets: {},
    diamondRoles: {},
    diamondTemplates: {},
    authUsers: [],
    diamondTracking: {},
    accessModel: {
      diamondOrderAdmin: 'Role-based via _SalesWorkflowUsers role DIAMOND_ORDER_ADMIN.',
      diamondOrderAssistant: 'Role-based via _SalesWorkflowUsers role DIAMOND_ORDER_ASSISTANT.',
      configNameEmailRequired: false
    }
  };

  [SW_SHEETS.CONFIG, SW_SHEETS.TEMPLATES, SW_SHEETS.USERS, SW_SHEETS.TASKS].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    out.sheets[name] = {
      exists: !!sh,
      rows: sh ? Math.max(0, sh.getLastRow() - 1) : 0
    };
  });

  var config = ss.getSheetByName(SW_SHEETS.CONFIG)
    ? swReadSheetObjectsExpectedHeaders_(ss.getSheetByName(SW_SHEETS.CONFIG), SW_CONFIG_HEADERS)
    : [];
  [SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN, SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT].forEach(function (role) {
    out.diamondRoles[role] = config.filter(function (row) {
      return swNorm_(row['Role']) === swNorm_(role);
    }).map(function (row) {
      return {
        key: row['Key'],
        name: row['Name'],
        email: row['Email'],
        active: row['Active?'],
        priority: row['Priority']
      };
    });
  });

  var templates = ss.getSheetByName(SW_SHEETS.TEMPLATES)
    ? swReadSheetObjectsExpectedHeaders_(ss.getSheetByName(SW_SHEETS.TEMPLATES), SW_TEMPLATE_HEADERS)
    : [];
  [
    SW_TASKS.DIAMOND_PROPOSE,
    SW_TASKS.DIAMOND_QUOTE,
    SW_TASKS.DIAMOND_ORDER,
    SW_TASKS.DIAMOND_TRACK,
    SW_TASKS.DIAMOND_DELIVERY,
    SW_TASKS.DIAMOND_DECISIONS,
    SW_TASKS.DIAMOND_RETURN,
    SW_TASKS.DIAMOND_ORDER_ACK_REP,
    SW_TASKS.DIAMOND_ORDER_ACK_JOC,
    SW_TASKS.DIAMOND_ETA_REP,
    SW_TASKS.DIAMOND_ETA_JOC
  ].forEach(function (taskType) {
    var row = templates.filter(function (t) { return t['Task Type'] === taskType; })[0];
    out.diamondTemplates[taskType] = {
      exists: !!row,
      title: row ? row['Task Title'] : '',
      primaryAction: row ? row['Primary Action'] : ''
    };
  });

  var users = ss.getSheetByName(SW_SHEETS.USERS)
    ? swReadSheetObjectsExpectedHeaders_(ss.getSheetByName(SW_SHEETS.USERS), SW_AUTH_USER_HEADERS)
    : [];
  out.authUsers = users.map(function (row) {
    return {
      email: row['Email'],
      name: row['Name'],
      roles: row['Roles'],
      active: row['Active?'],
      passwordSet: !!row['Password Hash'],
      lastLoginAt: row['Last Login At']
    };
  });

  var target = swDiamond200Target_();
  var missingTrackingColumns = [];
  if (target && target.sheet) {
    var trackingHm = swDiamond200HeaderMap_(target.sheet);
    if (!swDiamondFind200Column_(trackingHm, ['Tracking ETA', 'Tracking ETA Date', 'ETA Date', 'ETA'])) missingTrackingColumns.push('Tracking ETA');
    if (!swDiamondFind200Column_(trackingHm, ['Tracking Status', 'ETA Status', 'Shipment Status'])) missingTrackingColumns.push('Tracking Status');
  }
  out.diamondTracking = {
    available: !!(target && target.sheet),
    spreadsheetName: target && target.ss ? target.ss.getName() : '',
    tab: target ? target.tab : '',
    url: target && target.ss ? target.ss.getUrl() : '',
    missingColumns: missingTrackingColumns
  };
  Logger.log('SW_DIAMOND_WORKFLOW_SETUP_REVIEW ' + JSON.stringify(out, null, 2));
  return out;
}

/**
 * Read-only diagnostic: logs server-side speed for initial queue load paths.
 */
function sw_measureSalesWorkflowSpeed() {
  var started = new Date().getTime();
  var out = {
    ok: true,
    generatedAt: swIso_(new Date()),
    readOnly: true,
    note: 'Measures Apps Script server time only. Browser/network rendering time is not included.',
    steps: []
  };
  var bootstrap = null;
  var mine = null;
  var coverage = null;
  var admin = null;

  swBenchmarkSalesWorkflowStep_(out, 'sw_getBootstrap', function () {
    bootstrap = sw_getBootstrap();
    return {
      counts: bootstrap.counts || {},
      views: bootstrap.views || {},
      initialTasks: bootstrap.tasks ? bootstrap.tasks.length : 0
    };
  });

  swBenchmarkSalesWorkflowStep_(out, 'sw_getMyTasks:mine', function () {
    mine = sw_getMyTasks('mine');
    return swBenchmarkSalesWorkflowListSummary_(mine);
  });

  if (bootstrap && bootstrap.views && bootstrap.views.coverage) {
    swBenchmarkSalesWorkflowStep_(out, 'sw_getMyTasks:coverage', function () {
      coverage = sw_getMyTasks('coverage');
      return swBenchmarkSalesWorkflowListSummary_(coverage);
    });
  }

  if (bootstrap && bootstrap.views && bootstrap.views.admin) {
    swBenchmarkSalesWorkflowStep_(out, 'sw_getMyTasks:admin', function () {
      admin = sw_getMyTasks('admin');
      return swBenchmarkSalesWorkflowListSummary_(admin);
    });
  }

  var detailTaskId = swBenchmarkSalesWorkflowFirstTaskId_([mine, coverage, admin]);
  if (detailTaskId) {
    swBenchmarkSalesWorkflowStep_(out, 'sw_getTaskDetail', function () {
      var detail = sw_getTaskDetail(detailTaskId);
      return {
        taskId: detailTaskId,
        taskType: detail.task ? detail.task.taskType : '',
        attachments: detail.attachments ? detail.attachments.length : 0,
        missingFields: detail.missingFields ? detail.missingFields.length : 0
      };
    });
  } else {
    out.steps.push({
      operation: 'sw_getTaskDetail',
      skipped: true,
      reason: 'No visible task was available for detail timing.'
    });
  }

  out.totalMs = new Date().getTime() - started;
  Logger.log('SW_BENCHMARK_SUMMARY ' + JSON.stringify(out, null, 2));
  return out;
}

/**
 * Mutating diagnostic: runs setup and generation twice to confirm duplicate-safe generation.
 */
function sw_testSalesWorkflowDryRun() {
  var setup = sw_setupSalesWorkflow();
  var first = sw_generateSalesWorkflowTasks();
  var second = sw_generateSalesWorkflowTasks();
  Logger.log('SALES_WORKFLOW_TEST_DRY_RUN ' + JSON.stringify({
    setup: setup,
    first: first,
    second: second,
    duplicateSafe: second.created === 0
  }, null, 2));
  return { ok: true, setup: setup, first: first, second: second };
}

function swBenchmarkSalesWorkflowStep_(out, operation, fn) {
  var started = new Date().getTime();
  var step = {
    operation: operation,
    ok: true
  };
  try {
    step.result = fn() || {};
  } catch (err) {
    step.ok = false;
    step.error = err && err.message ? err.message : String(err);
  } finally {
    step.ms = new Date().getTime() - started;
    out.steps.push(step);
    Logger.log('SW_BENCHMARK_STEP ' + JSON.stringify(step));
  }
  return step;
}

function swBenchmarkSalesWorkflowListSummary_(res) {
  return {
    view: res && res.view ? res.view : '',
    tasks: res && res.tasks ? res.tasks.length : 0
  };
}

function swBenchmarkSalesWorkflowFirstTaskId_(responses) {
  for (var i = 0; i < responses.length; i++) {
    var tasks = responses[i] && responses[i].tasks ? responses[i].tasks : [];
    if (tasks.length && tasks[0].taskId) return tasks[0].taskId;
  }
  return '';
}
