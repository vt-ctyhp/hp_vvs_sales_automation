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
  var cleanupSheet = swEnsureSheet_(ss, SW_SHEETS.DATA_CLEANUP, SW_DATA_CLEANUP_HEADERS);
  var artifactSheet = swEnsureAppointmentArtifactsSheet_(ss);

  swStyleSheet_(taskSheet);
  swStyleSheet_(logSheet);
  swStyleSheet_(configSheet);
  swStyleSheet_(templateSheet);
  swStyleSheet_(usersSheet);
  swStyleSheet_(cleanupSheet);
  swStyleSheet_(artifactSheet);

  swSeedConfig_(configSheet);
  swSeedTemplates_(templateSheet);
  swSeedAuthUsers_(usersSheet);
  swNormalizeWorkflowUserRoleLabels_(ss);
  swEnsureDiamondRequirementMasterHeaders_(ss);
  if (typeof swEnsureDataCleanupMasterHeaders_ === 'function') swEnsureDataCleanupMasterHeaders_(ss);

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
      ctx.appointmentSummaryByRoot = typeof swAppointmentSummaryIndex_ === 'function' ? swAppointmentSummaryIndex_(ss) : {};
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
          summary.blocked += swBlockTasksForAppointment_(ss, taskState, rec, SW_INACTIVE_APPOINTMENT_BLOCK_REASON);
          return;
        }

        swGenerateTasksForAppointment_(ss, taskState, ctx, rec, now, summary);
      });

      if (typeof swGenerateDataCleanupTasks_ === 'function') {
        summary.dataCleanup = swGenerateDataCleanupTasks_(ss, taskState, ctx, masterRows, now, summary);
      }

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
 * Mutating setup: replaces Sales Workflow generation and appointment automation triggers.
 */
function sw_installSalesWorkflowTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    var fn = trigger.getHandlerFunction();
    if (fn === 'sw_generateSalesWorkflowTasks' ||
        fn === 'processUploadQueue' || fn === 'processSummariesWorker' ||
        fn === 'sw_processAppointmentAutomation') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('sw_generateSalesWorkflowTasks').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('sw_processAppointmentAutomation').timeBased().everyMinutes(5).create();
  return {
    ok: true,
    message: 'Installed hourly queue refresh and 5-minute appointment automation.'
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
    var config = swReadConfig_(ss, true);
    var cleanupTabEnabled = typeof swDataCleanupCampaignTabEnabled_ === 'function'
      ? swDataCleanupCampaignTabEnabled_(config)
      : false;
    var state = swReadTaskListState_(ss, true);
    mark('taskListRead', { tasks: state.tasks.length });
    mark('currentUser', { isAdmin: user.isAdmin, isJoc: user.isJoc });
    var buckets = swBuildVisibleTaskBuckets_(state, user, { cleanupCampaignTabEnabled: cleanupTabEnabled });
    mark('taskBuckets', {
      mine: buckets.mine.length,
      cleanup: buckets.cleanup.length,
      coverage: buckets.coverage.length,
      admin: buckets.admin.length
    });
    return {
      ok: true,
      user: user,
      tasks: buckets.mine,
      counts: {
        mine: buckets.mine.length,
        cleanup: buckets.cleanup.length,
        coverage: buckets.coverage.length,
        admin: user.isAdmin ? buckets.admin.length : 0
      },
      views: {
        mine: true,
        customerSearch: user.isAdmin || user.isJoc || user.isRep,
        calendar: true,
        inStockDiamonds: true,
        diamondTracking: user.isAdmin || user.isDiamondOrderAdmin || user.isDiamondOrderAssistant,
        bulkReturns: user.isAdmin || user.isDiamondOrderAdmin,
        cleanup: cleanupTabEnabled,
        coverage: user.isJoc || user.isAdmin,
        adminDashboard: user.isAdmin,
        admin: user.isAdmin
      },
      message: 'Connected. Use Refresh Queue to create or refresh the queue.'
    };
  });
}

/**
 * Read-only UI list: returns tasks visible to the current user for the requested view.
 */
function sw_getMyTasks(authToken, view) {
  return swTimed_('sw_getMyTasks', function () {
    var mark = swStepTimer_('sw_getMyTasks');
    if (!view && /^(mine|cleanup|coverage|admin)$/i.test(String(authToken || ''))) {
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
    var config = swReadConfig_(ss, true);
    var cleanupTabEnabled = typeof swDataCleanupCampaignTabEnabled_ === 'function'
      ? swDataCleanupCampaignTabEnabled_(config)
      : false;
    var state = swReadTaskListState_(ss, true);
    mark('taskListRead', { tasks: state.tasks.length });
    var tasks = swListVisibleTasksFromState_(state, user, viewName, { cleanupCampaignTabEnabled: cleanupTabEnabled });
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
      assignedRep: swDiamondFind200Column_(hm, ['Client Advisor', 'Assigned Rep', 'Sales Rep']),
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
    var returnWarning = Number(swConfigValue_(config, 'SYSTEM', 'DIAMOND_RETURN_WARNING_DAYS', '7')) || 7;
    var hm = swDiamond200HeaderMap_(sh);
    var C = {
      root: swDiamondFind200Column_(hm, ['RootApptID', 'APPT_ID', 'Root Appt ID']),
      customerName: swDiamondFind200Column_(hm, ['Customer Name', 'Client Name', 'Customer']),
      appointment: swDiamondFind200Column_(hm, ['Customer Appt Time & Date', 'Customer Appointment Date', 'Appointment Date']),
      assignedRep: swDiamondFind200Column_(hm, ['Client Advisor', 'Assigned Rep', 'Sales Rep']),
      joc: swDiamondFind200Column_(hm, ['JOC', 'Assisted Rep', 'Assistant Rep']),
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
    var warningMs = todayMs + returnWarning * 24 * 60 * 60 * 1000;
    var rows = [];
    var stats = {
      total: 0,
      available: 0,
      returnSoon: 0,
      returnOverdue: 0,
      noReturnDate: 0,
      assignmentMissing: 0,
      warningDays: returnWarning
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

      var shape = swDiamondCell_(row, C.shape);
      var carat = swDiamondCell_(row, C.carat);
      var color = swDiamondCell_(row, C.color);
      var clarity = swDiamondCell_(row, C.clarity);
      var daysUntilReturn = returnMs ? Math.ceil((returnMs - todayMs) / (24 * 60 * 60 * 1000)) : '';
      var root = swDiamondCell_(row, C.root);
      var customerName = swDiamondCell_(row, C.customerName);
      var assignedRep = swDiamondCell_(row, C.assignedRep);
      var joc = swDiamondCell_(row, C.joc);
      var assignmentMissing = !root || !customerName || !assignedRep || !joc;
      rows.push({
        rowIndex: i + 3,
        root: root,
        customerName: customerName,
        appointment: swDiamondCell_(row, C.appointment),
        assignedRep: assignedRep,
        joc: joc,
        company: swDiamondCell_(row, C.company),
        vendor: swDiamondCell_(row, C.vendor),
        stoneType: swDiamondCell_(row, C.stoneType),
        certNo: swDiamondCell_(row, C.certNo),
        shape: shape,
        carat: carat,
        color: color,
        clarity: clarity,
        diamond: [shape, carat, color, clarity].filter(Boolean).join(' '),
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
        warningDays: returnWarning,
        issue: issue,
        assignmentMissing: assignmentMissing,
        availabilityLabel: returnDueDate ? ('Available until ' + returnDueDate) : 'Return date missing'
      });
      stats.total++;
      if (!issue) stats.available++;
      if (issue === 'Return soon') stats.returnSoon++;
      if (issue === 'Return overdue') stats.returnOverdue++;
      if (issue === 'No return date') stats.noReturnDate++;
      if (assignmentMissing) stats.assignmentMissing++;
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

/**
 * Read-only UI return picker: diamond order admins can select stones for one
 * bulk return shipment before writing Return in Progress to 200_.
 */
function sw_getBulkReturnCandidates(authToken) {
  return swTimed_('sw_getBulkReturnCandidates', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken);
    swRequireDiamondBulkReturnUser_(user);

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
    var returnWarning = Number(swConfigValue_(config, 'SYSTEM', 'DIAMOND_RETURN_WARNING_DAYS', '7')) || 7;
    var hm = swDiamond200HeaderMap_(sh);
    var C = {
      root: swDiamondFind200Column_(hm, ['RootApptID', 'APPT_ID', 'Root Appt ID']),
      customerName: swDiamondFind200Column_(hm, ['Customer Name', 'Client Name', 'Customer']),
      assignedRep: swDiamondFind200Column_(hm, ['Client Advisor', 'Assigned Rep', 'Sales Rep']),
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
      returnDueDate: swDiamondFind200Column_(hm, ['Return DUE DATE', 'Return Due Date', 'Return Due'])
    };

    var values = sh.getRange(3, 1, lr - 2, lc).getDisplayValues();
    var today = new Date();
    var todayMs = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
    var warningMs = todayMs + returnWarning * 24 * 60 * 60 * 1000;
    var rows = [];
    var stats = {
      total: 0,
      available: 0,
      returnSoon: 0,
      returnOverdue: 0,
      noReturnDate: 0,
      warningDays: returnWarning
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

      rows.push({
        rowIndex: i + 3,
        root: swDiamondCell_(row, C.root),
        customerName: swDiamondCell_(row, C.customerName),
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
        returnDueDate: returnDueDate,
        daysUntilReturn: returnMs ? Math.ceil((returnMs - todayMs) / (24 * 60 * 60 * 1000)) : '',
        warningDays: returnWarning,
        issue: issue
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
      rows: rows.slice(0, 500)
    };
  });
}

/**
 * Mutating UI action: marks selected 200_ rows as Return in Progress for a
 * single bulk shipment and records a shared shipment note.
 */
function sw_bulkMarkDiamondsReturnInProgress(authToken, payload) {
  return swTimed_('sw_bulkMarkDiamondsReturnInProgress', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken);
    swRequireDiamondBulkReturnUser_(user);

    payload = payload || {};
    var seen = {};
    var rowIndexes = (payload.rowIndexes || []).map(function (rowIndex) {
      return Number(rowIndex);
    }).filter(function (rowIndex) {
      if (!(rowIndex >= 3) || seen[rowIndex]) return false;
      seen[rowIndex] = true;
      return true;
    });
    if (!rowIndexes.length) throw new Error('Select at least one diamond to return.');
    if (rowIndexes.length > 250) throw new Error('Select 250 or fewer diamonds per bulk return shipment.');

    var lock = LockService.getDocumentLock() || LockService.getScriptLock();
    lock.waitLock(28000);
    try {
      var target = swDiamond200Target_();
      if (!target || !target.sheet) throw new Error('Diamond tracking sheet is unavailable.');
      var sh = target.sheet;
      var hm = swDiamond200HeaderMap_(sh);
      var C = {
        root: swDiamondFind200Column_(hm, ['RootApptID', 'APPT_ID', 'Root Appt ID']),
        certNo: swDiamondFind200Column_(hm, ['Certificate No', 'Cert #', 'Cert No', 'Certificate #']),
        orderStatus: swDiamondFind200Column_(hm, ['Order Status', 'OrderStatus']),
        stoneStatus: swDiamondFind200Column_(hm, ['Stone Status', 'StoneStatus']),
        decision: swDiamondFind200Column_(hm, ['Stone Decision (PO, Return)', 'Stone Decision', 'StoneDecision'])
      };
      if (!C.stoneStatus) throw new Error('Stone Status column is missing in 200_.');
      if (!C.decision) C.decision = swDiamondEnsure200Column_(sh, 'Stone Decision (PO, Return)');
      var cNotes = swDiamondEnsure200Column_(sh, 'Return Notes');
      var lastRow = sh.getLastRow();
      var lastCol = sh.getLastColumn();
      var now = swIso_(new Date());
      var note = swTrim_(payload.note || '');
      var line = 'Bulk return in progress @ ' + now + (user && user.email ? ' by ' + user.email : '') + (note ? ' | ' + note : '');
      var updatedRows = [];
      var skippedRows = [];
      var touchedAppointments = {};
      var updatesForJsonByAppt = {};

      rowIndexes.forEach(function (rowIndex) {
        if (rowIndex > lastRow) {
          skippedRows.push({ rowIndex: rowIndex, reason: 'Row no longer exists' });
          return;
        }
        var row = sh.getRange(rowIndex, 1, 1, lastCol).getDisplayValues()[0];
        var orderStatus = swDiamondCell_(row, C.orderStatus);
        var stoneStatus = swDiamondCell_(row, C.stoneStatus);
        var decision = swDiamondCell_(row, C.decision);
        var stoneNorm = swNorm_(stoneStatus);
        var orderNorm = swNorm_(orderStatus);
        var decisionNorm = swNorm_(decision);
        var isInStock = stoneNorm.indexOf('in stock') >= 0 || orderNorm === 'delivered';
        var unavailable = /return in progress|returned|sold|customer purchased/.test(stoneNorm) ||
          decisionNorm === 'purchase' || decisionNorm === 'purchased';
        if (!isInStock || unavailable) {
          skippedRows.push({ rowIndex: rowIndex, reason: 'Row is no longer return-eligible' });
          return;
        }

        var nextStatus = swDiamondMergeStatus_(stoneStatus, 'Return in Progress');
        if (typeof cd_writeBypassValidation_ === 'function') {
          cd_writeBypassValidation_(sh, rowIndex, C.stoneStatus, nextStatus);
        } else {
          sh.getRange(rowIndex, C.stoneStatus).setValue(nextStatus);
        }
        sh.getRange(rowIndex, C.decision).setValue('Return');
        var existingNotes = swTrim_(sh.getRange(rowIndex, cNotes).getDisplayValue());
        sh.getRange(rowIndex, cNotes).setValue(existingNotes ? existingNotes + '\n' + line : line);

        var root = swDiamondCell_(row, C.root);
        var certNo = swDiamondCell_(row, C.certNo);
        if (root) {
          touchedAppointments[root] = true;
          if (!updatesForJsonByAppt[root]) updatesForJsonByAppt[root] = [];
          updatesForJsonByAppt[root].push({ certNo: certNo, decision: 'Return', hold: null });
        }
        updatedRows.push(rowIndex);
      });

      Object.keys(touchedAppointments).forEach(function (root) {
        try {
          if (typeof dp_update100JsonLinesWithDecisions_ === 'function') {
            dp_update100JsonLinesWithDecisions_(root, updatesForJsonByAppt[root] || []);
          }
        } catch (_) {}
        try {
          if (typeof dp_computeCountsForAppointment_ === 'function' && typeof dp_refresh100QuickRef_ === 'function') {
            var counts = dp_computeCountsForAppointment_(sh, hm, root);
            dp_refresh100QuickRef_(root, counts, sh, hm);
          }
        } catch (_) {}
      });

      try {
        swAppendTaskLog_(ss, 'BULK_DIAMOND_RETURN', {
          taskId: '',
          root: '',
          appt: '',
          taskType: 'BULK_RETURN_DIAMONDS',
          status: SW_STATUSES.COMPLETED
        }, user, '', '', {
          updatedRows: updatedRows,
          skippedRows: skippedRows,
          note: note
        });
      } catch (_) {}

      return {
        ok: true,
        status: 'Return in Progress',
        updatedRows: updatedRows,
        skippedRows: skippedRows,
        updatedCount: updatedRows.length,
        spreadsheetUrl: target.ss.getUrl(),
        tab: target.tab
      };
    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }
  });
}

function swRequireDiamondBulkReturnUser_(user) {
  if (!(user && (user.isAdmin || user.isDiamondOrderAdmin))) {
    throw new Error('Diamond order admin access required.');
  }
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
    var renderedTemplate = swRenderedCopyableTemplateForTask_(task, template, renderData);
    var renderedAttachmentUrl = template.attachmentUrl ? swRenderTemplate_(template.attachmentUrl, renderData) : '';
    var renderedAttachmentLabel = template.attachmentLabel ? swRenderTemplate_(template.attachmentLabel, renderData) : '';
    var attachments = swAttachmentsForTask_(task, template, renderData);
    var missingFields = swMissingFieldsForTask_(task, template, renderData);
    var checklist = swParseJson_(template.checklistJson, []);
    mark('render');
    var formOptions = typeof swTaskFormOptions_ === 'function' ? swTaskFormOptions_(ss, task) : {};
    mark('formOptions', { groups: formOptions ? Object.keys(formOptions).length : 0 });
    var appointmentArtifacts = typeof swPublicAppointmentArtifacts_ === 'function'
      ? swPublicAppointmentArtifacts_(ss, task.root || task.appt || '')
      : [];
    mark('appointmentArtifacts', { artifacts: appointmentArtifacts.length });
    var assignmentOptions = user.isAdmin ? swReadAssignmentOptions_(ss) : {};
    mark('assignmentOptions', {
      salesReps: assignmentOptions.salesReps ? assignmentOptions.salesReps.length : 0,
      jocReps: assignmentOptions.jocReps ? assignmentOptions.jocReps.length : 0
    });

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
      formOptions: formOptions,
      appointmentArtifacts: appointmentArtifacts,
      assignmentOptions: assignmentOptions,
      missingFields: missingFields,
      checklist: checklist,
      canComplete: swCanActOnTask_(task, user),
      canClaim: swCanClaimTask_(task, user),
      canAdmin: user.isAdmin
    };
  });
}

function sw_getAppointmentUploadFolder(authToken, taskId, artifactType) {
  return swTimed_('sw_getAppointmentUploadFolder', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss);
    var user = swAuthUserForApi_(ss, authToken);
    var task = swReadTaskRowById_(ss, taskId, true);
    if (!task) throw new Error('Task not found: ' + taskId);
    if (!swCanActOnTask_(task, user)) throw new Error('You are not the current owner for this appointment task.');

    var payload = swParseJson_(task.payloadJson, {});
    var root = swTrim_(task.root || task.appt || (payload.appointment && (payload.appointment.root || payload.appointment.appt)) || '');
    if (!root) throw new Error('Missing RootApptID for appointment upload folder.');
    if (typeof swEnsureAppointmentFolderForRoot_ !== 'function' ||
        typeof swArtifactDriveDropFolder_ !== 'function') {
      throw new Error('Appointment artifact folder helpers are not available.');
    }

    var type = typeof swNormalizeDriveUploadArtifactType_ === 'function'
      ? swNormalizeDriveUploadArtifactType_(artifactType)
      : (swTrim_(artifactType) || 'APPOINTMENT_RECORDING');
    var folders = swEnsureAppointmentFolderForRoot_(ss, root);
    var folder = swArtifactDriveDropFolder_(folders, type);
    return {
      ok: true,
      rootApptId: root,
      artifactType: type,
      folderId: folder.getId(),
      url: folder.getUrl()
    };
  });
}

function sw_syncAppointmentDriveUploads(authToken, taskId) {
  return swTimed_('sw_syncAppointmentDriveUploads', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss);
    var user = swAuthUserForApi_(ss, authToken);
    var task = swReadTaskRowById_(ss, taskId, true);
    if (!task) throw new Error('Task not found: ' + taskId);
    if (!swCanActOnTask_(task, user)) throw new Error('You are not the current owner for this appointment task.');

    var payload = swParseJson_(task.payloadJson, {});
    var root = swTrim_(task.root || task.appt || (payload.appointment && (payload.appointment.root || payload.appointment.appt)) || '');
    if (!root) throw new Error('Missing RootApptID for appointment upload sync.');
    if (typeof swSyncAppointmentDriveUploads_ !== 'function') throw new Error('Appointment Drive upload sync is not available.');

    var created = swSyncAppointmentDriveUploads_(ss, root, task.taskId || taskId, user);
    return {
      ok: true,
      rootApptId: root,
      registered: created.length,
      artifacts: typeof swPublicAppointmentArtifacts_ === 'function' ? swPublicAppointmentArtifacts_(ss, root) : []
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
  if (!swTaskPendingLike_(task, new Date().getTime())) throw new Error('Only pending or due snoozed tasks can be completed.');

  data = data || {};
  swValidateCompletion_(ss, task, data);
  var diamondAction = swDiamondHandleTaskCompletion_(ss, task, data, user);
  var postConsultAction = typeof swHandlePostConsultTaskCompletion_ === 'function'
    ? swHandlePostConsultTaskCompletion_(ss, task, data, user)
    : null;
  var dataCleanupAction = typeof swHandleDataCleanupTaskCompletion_ === 'function'
    ? swHandleDataCleanupTaskCompletion_(ss, task, data, user)
    : null;
  var appointmentAction = typeof swHandleAppointmentCompletion_ === 'function'
    ? swHandleAppointmentCompletion_(ss, task, data, user)
    : null;
  var approvalAction = task.taskType === SW_TASKS.APPROVE && typeof swMarkAppointmentSummaryApproved_ === 'function'
    ? swMarkAppointmentSummaryApproved_(ss, task.root || task.appt || '', data.approvedText || '', user)
    : null;
  var jocHandoffAction = task.taskType === SW_TASKS.FINAL && typeof swMarkAppointmentJocHandoff_ === 'function'
    ? swMarkAppointmentJocHandoff_(ss, task.root || task.appt || '', user)
    : null;

  var template = swTemplateForType_(ss, task.taskType);
  var payload = swParseJson_(task.payloadJson, {});
  var renderData = swRenderDataForTask_(task, payload);
  var renderedTemplate = swRenderedCopyableTemplateForTask_(task, template, renderData);
  var renderedAttachments = swAttachmentsForTask_(task, template, renderData);
  payload.completion = data;
  payload.renderedTemplate = renderedTemplate;
  payload.renderedAttachments = renderedAttachments;
  payload.completedBy = user.name || user.email;
  payload.completedByEmail = user.email;
  payload.completedAt = swIso_(new Date());
  if (diamondAction) payload.diamondAction = diamondAction;
  if (postConsultAction) payload.postConsultAction = postConsultAction;
  if (dataCleanupAction) payload.dataCleanupAction = dataCleanupAction;
  if (appointmentAction) payload.appointmentAction = appointmentAction;
  if (approvalAction) payload.approvalAction = approvalAction;
  if (jocHandoffAction) payload.jocHandoffAction = jocHandoffAction;

  var oldOwner = task.currentOwner;
  task.status = SW_STATUSES.COMPLETED;
  task.completedBy = user.name || user.email;
  task.completedByEmail = user.email;
  task.completedAt = payload.completedAt;
  task.updatedAt = payload.completedAt;
  task.lastEvent = 'COMPLETE';
  task.snoozeUntil = '';
  task.snoozeReason = '';
  task.snoozedBy = '';
  task.snoozedAt = '';
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
 * Mutating task action: snoozes a pending workflow task until a future date.
 * Snoozed tasks are hidden from active queues and do not count late until then.
 */
function sw_snoozeTask(authToken, taskId, data) {
  if (/^SW\|/.test(String(authToken || ''))) {
    data = taskId;
    taskId = authToken;
    authToken = '';
  }
  data = data || {};
  var untilDate = swTrim_(data.untilDate || data.date || '');
  var reason = swTrim_(data.reason || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(untilDate)) throw new Error('Select a valid snooze date.');
  if (!reason) throw new Error('Enter a snooze reason.');

  var parts = untilDate.split('-').map(Number);
  var until = new Date(parts[0], parts[1] - 1, parts[2], 9, 30, 0, 0);
  var today = new Date();
  var todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0, 0);
  if (until.getTime() < todayStart.getTime()) throw new Error('Snooze date must be today or later.');

  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swAuthUserForApi_(ss, authToken);
  var task = swGetTaskById_(ss, taskId);
  if (!task) throw new Error('Task not found: ' + taskId);
  if (!swCanActOnTask_(task, user)) throw new Error('You are not the current owner for this task.');
  if (task.status !== SW_STATUSES.PENDING && task.status !== SW_STATUSES.SNOOZED) {
    throw new Error('Only pending or snoozed tasks can be snoozed.');
  }

  var oldOwner = task.currentOwner;
  var now = swIso_(new Date());
  task.status = SW_STATUSES.SNOOZED;
  task.snoozeUntil = swIso_(until);
  task.snoozeReason = reason;
  task.snoozedBy = user.name || user.email;
  task.snoozedAt = now;
  task.updatedAt = now;
  task.lastEvent = 'SNOOZE';
  swWriteTaskRow_(ss, task);
  swAppendTaskLog_(ss, 'SNOOZE', task, user, oldOwner, task.currentOwner, {
    untilDate: untilDate,
    snoozeUntil: task.snoozeUntil,
    reason: reason
  });
  return { ok: true, task: swGetTaskById_(ss, taskId) };
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
 * Mutating admin action: assigns appointment-level Client Advisor and JOC owner
 * on the Master appointment row, then refreshes workflow task owners.
 */
function sw_adminAssignAppointmentOwners(authToken, taskId, data) {
  if (/^SW\|/.test(String(authToken || ''))) {
    data = taskId;
    taskId = authToken;
    authToken = '';
  }
  data = data || {};
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swAuthUserForApi_(ss, authToken);
  if (!user.isAdmin) throw new Error('Admin access required.');

  var task = swGetTaskById_(ss, taskId);
  if (!task) throw new Error('Task not found: ' + taskId);
  var row = (typeof swMasterRowForTask_ === 'function') ? swMasterRowForTask_(ss, task) : 0;
  if (!row) throw new Error('Could not resolve Master row for this task.');

  var master = ss.getSheetByName(SW_SHEETS.MASTER);
  var headers = swEnsureMasterOwnerHeaders_(master);
  var values = master.getRange(row, 1, 1, master.getLastColumn()).getDisplayValues()[0];
  var root = headers.root ? swTrim_(values[headers.root - 1]) : '';

  var assignedName = swTrim_(data.assignedRep);
  var assistedName = swTrim_(data.assistedRep);
  var reason = swTrim_(data.reason);
  var options = swReadAssignmentOptions_(ss);
  var assignedEmail = swTrim_(data.assignedRepEmail) || swAssignmentEmailForName_(assignedName, options.salesReps);
  var assistedEmail = swTrim_(data.assistedRepEmail) || swAssignmentEmailForName_(assistedName, options.jocReps);

  var targetRows = [row];
  if (root && headers.root) {
    var roots = master.getRange(2, headers.root, Math.max(0, master.getLastRow() - 1), 1).getDisplayValues();
    targetRows = [];
    roots.forEach(function (r, i) {
      if (swTrim_(r[0]) === root) targetRows.push(i + 2);
    });
    if (!targetRows.length) targetRows = [row];
  }

  targetRows.forEach(function (r) {
    master.getRange(r, headers.assignedRep).setValue(assignedName);
    master.getRange(r, headers.assignedRepEmail).setValue(assignedEmail);
    master.getRange(r, headers.assistedRep).setValue(assistedName);
    master.getRange(r, headers.assistedRepEmail).setValue(assistedEmail);
  });

  swAppendTaskLog_(ss, 'APPOINTMENT_OWNER_ASSIGN', task, user, task.currentOwner, task.currentOwner, {
    assignedRep: assignedName,
    assignedRepEmail: assignedEmail,
    assistedRep: assistedName,
    assistedRepEmail: assistedEmail,
    reason: reason,
    rootApptId: root,
    rowsUpdated: targetRows
  });

  var generation = sw_generateSalesWorkflowTasks();
  return {
    ok: true,
    rowsUpdated: targetRows.length,
    assignedRep: assignedName,
    assignedRepEmail: assignedEmail,
    assistedRep: assistedName,
    assistedRepEmail: assistedEmail,
    reason: reason,
    generation: generation
  };
}

function swEnsureMasterOwnerHeaders_(master) {
  if (!master) throw new Error('Missing sheet: ' + SW_SHEETS.MASTER);
  var required = [
    { header: 'Assigned Rep', aliases: ['Client Advisor', 'Assigned Rep', 'Rep', 'Owner'] },
    { header: 'Assigned Rep Email', aliases: ['Client Advisor Email', 'Assigned Rep Email', 'Rep Email', 'Owner Email'] },
    { header: 'Assisted Rep', aliases: ['Assisted Rep', 'Assistant Rep'] },
    { header: 'Assisted Rep Email', aliases: ['Assisted Rep Email', 'Assistant Rep Email'] }
  ];
  var headers = master.getRange(1, 1, 1, Math.max(1, master.getLastColumn())).getDisplayValues()[0].map(function (h) {
    return swTrim_(h);
  });
  var H = swHeaderMapFromArray_(headers);
  required.forEach(function (item) {
    if (swPickIndex_(H, item.aliases) >= 0) return;
    master.getRange(1, master.getLastColumn() + 1).setValue(item.header);
    headers.push(item.header);
    H = swHeaderMapFromArray_(headers);
  });
  return {
    root: swPickIndex_(H, ['RootApptID', 'APPT_ID']) + 1,
    assignedRep: swPickIndex_(H, ['Client Advisor', 'Assigned Rep', 'Rep', 'Owner']) + 1,
    assignedRepEmail: swPickIndex_(H, ['Client Advisor Email', 'Assigned Rep Email', 'Rep Email', 'Owner Email']) + 1,
    assistedRep: swPickIndex_(H, ['Assisted Rep', 'Assistant Rep']) + 1,
    assistedRepEmail: swPickIndex_(H, ['Assisted Rep Email', 'Assistant Rep Email']) + 1
  };
}

function swReadAssignmentOptions_(ss) {
  var cacheKey = '';
  try {
    cacheKey = 'sw:assignmentOptions:v1:' + ss.getId();
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return swParseJson_(cached, { salesReps: [], jocReps: [] });
  } catch (_) {}
  var out = { salesReps: [], jocReps: [] };
  var sh = ss.getSheetByName(SW_SHEETS.DROPDOWN);
  if (sh && sh.getLastRow() >= 2 && sh.getLastColumn() >= 1) {
    var values = sh.getDataRange().getDisplayValues();
    var H = swHeaderMapFromArray_(values[0].map(function (h) { return swTrim_(h); }));
    swPushAssignmentOptionsFromColumns_(out.salesReps, values, swPickIndex_(H, ['Client Advisor', 'Assigned Rep']), swPickIndex_(H, ['Client Advisor Email', 'Assigned Rep Email']));
    swPushAssignmentOptionsFromColumns_(out.jocReps, values, swPickIndex_(H, ['Assisted Rep', 'Assistant Rep']), swPickIndex_(H, ['Assisted Rep Email', 'Assistant Rep Email']));
  }

  try {
    swAuthReadUserRows_(ss, false).forEach(function (row) {
      var roles = swAuthRoles_(row['Roles']);
      var item = { name: swTrim_(row['Name']) || swNormEmail_(row['Email']), email: swNormEmail_(row['Email']) };
      if (swAuthHasRole_(roles, SW_OWNER_ROLES.SALES_REP)) swPushAssignmentOption_(out.salesReps, item);
      if (swAuthHasRole_(roles, 'JOC')) swPushAssignmentOption_(out.jocReps, item);
    });
  } catch (_) {}

  out.salesReps.sort(swAssignmentOptionSort_);
  out.jocReps.sort(swAssignmentOptionSort_);
  if (cacheKey) {
    try {
      var json = JSON.stringify(out);
      if (json.length <= 90000) CacheService.getScriptCache().put(cacheKey, json, 300);
    } catch (_) {}
  }
  return out;
}

function swPushAssignmentOptionsFromColumns_(target, values, nameCol, emailCol) {
  if (nameCol < 0) return;
  for (var i = 1; i < values.length; i++) {
    swPushAssignmentOption_(target, {
      name: swTrim_(values[i][nameCol]),
      email: emailCol >= 0 ? swNormEmail_(values[i][emailCol]) : ''
    });
  }
}

function swPushAssignmentOption_(target, item) {
  if (!item || !item.name) return;
  for (var i = 0; i < target.length; i++) {
    if (swNorm_(target[i].name) === swNorm_(item.name) || (item.email && swNormEmail_(target[i].email) === swNormEmail_(item.email))) {
      if (!target[i].email && item.email) target[i].email = item.email;
      return;
    }
  }
  target.push({ name: item.name, email: item.email || '' });
}

function swAssignmentEmailForName_(name, options) {
  name = swNorm_(name);
  for (var i = 0; i < (options || []).length; i++) {
    if (swNorm_(options[i].name) === name) return options[i].email || '';
  }
  return '';
}

function swAssignmentOptionSort_(a, b) {
  return String(a.name || '').localeCompare(String(b.name || ''));
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
    masterDiamondRequirementColumns: {},
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

  var master = ss.getSheetByName(SW_SHEETS.MASTER);
  if (master) {
    var masterHeaders = master.getRange(1, 1, 1, Math.max(master.getLastColumn(), 1)).getDisplayValues()[0];
    var masterMap = swHeaderMapFromArray_(masterHeaders);
    [
      'DV Customer Looking For',
      'DV Variety Strategy',
      'DV Customer Requirements (JSON)'
    ].forEach(function (header) {
      out.masterDiamondRequirementColumns[header] = swPickIndex_(masterMap, [header]) >= 0;
    });
  }

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
 * Read-only diagnostic: logs server-side speed for the major dashboard load
 * paths. Mutating actions are intentionally not executed.
 *
 * Optional second argument:
 *   { detailLimit: 8, customerSearchQueries: [''], calendarMonths: ['2026-05'] }
 */
function sw_measureSalesWorkflowSpeed(authToken, options) {
  if (authToken && typeof authToken === 'object' && options == null) {
    options = authToken;
    authToken = '';
  }
  authToken = swTrim_(authToken);
  options = swBenchmarkSalesWorkflowOptions_(options);

  var started = new Date().getTime();
  var out = {
    ok: true,
    generatedAt: swIso_(new Date()),
    readOnly: true,
    note: 'Measures Apps Script server time only. Browser/network rendering time is not included. Mutating workflow actions are skipped.',
    options: {
      detailLimit: options.detailLimit,
      includeTaskDetails: options.includeTaskDetails,
      customerSearchQueries: options.customerSearchQueries,
      calendarMonths: options.calendarMonths,
      adminDashboardFilters: options.adminDashboardFilters
    },
    steps: []
  };
  var bootstrap = null;
  var queueResponses = [];
  var customerSearchResponses = [];

  swBenchmarkSalesWorkflowStep_(out, 'sw_getBootstrap', function () {
    bootstrap = sw_getBootstrap(authToken);
    return {
      counts: bootstrap.counts || {},
      views: bootstrap.views || {},
      initialTasks: bootstrap.tasks ? bootstrap.tasks.length : 0
    };
  });

  swBenchmarkSalesWorkflowQueueStep_(out, queueResponses, authToken, 'mine');
  ['cleanup', 'coverage', 'admin'].forEach(function (viewName) {
    if (bootstrap && bootstrap.views && bootstrap.views[viewName]) {
      swBenchmarkSalesWorkflowQueueStep_(out, queueResponses, authToken, viewName);
    } else {
      swBenchmarkSalesWorkflowSkip_(out, 'sw_getMyTasks:' + viewName, 'View not visible for this user.');
    }
  });

  if (options.includeTaskDetails) {
    var detailTasks = swBenchmarkSalesWorkflowDetailSamples_(queueResponses, options.detailLimit);
    if (detailTasks.length) {
      detailTasks.forEach(function (sample) {
        swBenchmarkSalesWorkflowStep_(out, 'sw_getTaskDetail:' + swBenchmarkSalesWorkflowLabel_(sample.taskType), function () {
          var detail = sw_getTaskDetail(authToken, sample.taskId);
          return swBenchmarkSalesWorkflowTaskDetailSummary_(detail, sample);
        });
      });
    } else {
      swBenchmarkSalesWorkflowSkip_(out, 'sw_getTaskDetail', 'No visible task was available for detail timing.');
    }
  } else {
    swBenchmarkSalesWorkflowSkip_(out, 'sw_getTaskDetail', 'Task detail timing disabled by options.');
  }

  if (bootstrap && bootstrap.views && bootstrap.views.customerSearch) {
    options.customerSearchQueries.forEach(function (query) {
      swBenchmarkSalesWorkflowStep_(out, 'sw_searchCustomers:' + swBenchmarkSalesWorkflowLabel_(query || 'initial'), function () {
        var filters = {
          query: query,
          activeOnly: true
        };
        var res = sw_searchCustomers(authToken, query, filters);
        customerSearchResponses.push(res);
        return swBenchmarkSalesWorkflowCustomerSearchSummary_(res);
      });
    });
    var customerRoot = swBenchmarkSalesWorkflowFirstCustomerRoot_(customerSearchResponses);
    if (customerRoot && options.includeCustomerDetail) {
      swBenchmarkSalesWorkflowStep_(out, 'sw_getCustomerSearchDetail', function () {
        var detail = sw_getCustomerSearchDetail(authToken, customerRoot);
        return swBenchmarkSalesWorkflowCustomerDetailSummary_(detail);
      });
    } else {
      swBenchmarkSalesWorkflowSkip_(out, 'sw_getCustomerSearchDetail', customerRoot ? 'Customer detail timing disabled by options.' : 'No customer card was available for detail timing.');
    }
  } else {
    swBenchmarkSalesWorkflowSkip_(out, 'sw_searchCustomers', 'Customer Search is not visible for this user.');
    swBenchmarkSalesWorkflowSkip_(out, 'sw_getCustomerSearchDetail', 'Customer Search is not visible for this user.');
  }

  options.calendarMonths.forEach(function (monthKey) {
    swBenchmarkSalesWorkflowStep_(out, 'sw_getCalendarAppointments:' + monthKey, function () {
      var res = sw_getCalendarAppointments(authToken, monthKey);
      return {
        monthKey: res.monthKey || monthKey,
        monthLabel: res.monthLabel || '',
        appointments: res.appointmentCount || (res.appointments ? res.appointments.length : 0)
      };
    });
  });

  if (bootstrap && bootstrap.views && bootstrap.views.inStockDiamonds) {
    swBenchmarkSalesWorkflowStep_(out, 'sw_getInStockDiamonds', function () {
      return swBenchmarkSalesWorkflowRowsSummary_(sw_getInStockDiamonds(authToken));
    });
  } else {
    swBenchmarkSalesWorkflowSkip_(out, 'sw_getInStockDiamonds', 'In-Stock Diamonds is not visible for this user.');
  }

  if (bootstrap && bootstrap.views && bootstrap.views.diamondTracking) {
    swBenchmarkSalesWorkflowStep_(out, 'sw_getDiamondTrackingDashboard', function () {
      return swBenchmarkSalesWorkflowRowsSummary_(sw_getDiamondTrackingDashboard(authToken));
    });
  } else {
    swBenchmarkSalesWorkflowSkip_(out, 'sw_getDiamondTrackingDashboard', 'Diamond Tracking is not visible for this user.');
  }

  if (bootstrap && bootstrap.views && bootstrap.views.bulkReturns) {
    swBenchmarkSalesWorkflowStep_(out, 'sw_getBulkReturnCandidates', function () {
      return swBenchmarkSalesWorkflowRowsSummary_(sw_getBulkReturnCandidates(authToken));
    });
  } else {
    swBenchmarkSalesWorkflowSkip_(out, 'sw_getBulkReturnCandidates', 'Bulk Returns is not visible for this user.');
  }

  if (bootstrap && bootstrap.views && bootstrap.views.adminDashboard) {
    options.adminDashboardFilters.forEach(function (filters, index) {
      swBenchmarkSalesWorkflowStep_(out, 'sw_getAdminDashboard:' + (filters.windowPreset || ('filters' + (index + 1))), function () {
        return swBenchmarkSalesWorkflowAdminDashboardSummary_(sw_getAdminDashboard(authToken, filters));
      });
    });
  } else {
    swBenchmarkSalesWorkflowSkip_(out, 'sw_getAdminDashboard', 'Admin Dashboard is not visible for this user.');
  }

  if (bootstrap && bootstrap.views && bootstrap.views.admin) {
    swBenchmarkSalesWorkflowStep_(out, 'sw_adminGetTasks:pending', function () {
      return swBenchmarkSalesWorkflowListSummary_(sw_adminGetTasks(authToken, { status: SW_STATUSES.PENDING }));
    });
    if (swSpreadsheet_().getSheetByName(SW_SHEETS.USERS)) {
      swBenchmarkSalesWorkflowStep_(out, 'sw_adminListWorkflowUsers', function () {
        var res = sw_adminListWorkflowUsers(authToken);
        return {
          users: res.users ? res.users.length : 0,
          roles: res.roleOptions ? res.roleOptions.length : 0
        };
      });
    } else {
      swBenchmarkSalesWorkflowSkip_(out, 'sw_adminListWorkflowUsers', 'Workflow user sheet does not exist.');
    }
  } else {
    swBenchmarkSalesWorkflowSkip_(out, 'sw_adminGetTasks:pending', 'Admin Review is not visible for this user.');
    swBenchmarkSalesWorkflowSkip_(out, 'sw_adminListWorkflowUsers', 'Admin controls are not visible for this user.');
  }

  out.totalMs = new Date().getTime() - started;
  out.summary = swBenchmarkSalesWorkflowSummary_(out.steps);
  out.ok = out.summary.failedSteps === 0;
  Logger.log('SW_BENCHMARK_SUMMARY ' + JSON.stringify(swBenchmarkSalesWorkflowLogSummary_(out)));
  return out;
}

/**
 * Read-only diagnostic: explains why tasks for a named/email owner do or do not
 * appear in that user's My Queue.
 */
function sw_diagnoseTaskVisibilityForOwner(authToken, ownerNameOrEmail) {
  if (ownerNameOrEmail == null && !/^sw_/i.test(String(authToken || ''))) {
    ownerNameOrEmail = authToken;
    authToken = '';
  }

  var ss = swSpreadsheet_();
  swRequireWorkflowReadSheets_(ss, { templates: false });
  var requester = swAuthUserForApi_(ss, authToken);
  var target = swDiagnosticTargetUser_(ss, ownerNameOrEmail, requester);
  if (!requester.isAdmin &&
      swNormEmail_(requester.email) !== swNormEmail_(target.email) &&
      swNorm_(requester.name) !== swNorm_(target.name)) {
    throw new Error('Admin access required to diagnose another user.');
  }

  var now = new Date().getTime();
  var state = swReadTaskListState_(ss, true);
  var statusCounts = {};
  var hiddenReasons = {};
  var examples = [];
  var matched = 0;
  var dueNow = 0;
  var visible = 0;

  (state.tasks || []).forEach(function (task) {
    if (!swDiagnosticTaskOwnerMatches_(task, target)) return;
    matched++;
    var status = task.status || '(blank)';
    statusCounts[status] = (statusCounts[status] || 0) + 1;
    var isDue = swTaskDueForQueue_(task, now);
    var isOwned = swTaskOwnedByUser_(task, target);
    var inMine = isDue && isOwned;
    if (isDue) dueNow++;
    if (inMine) visible++;
    if (!inMine) {
      var reason = !swTaskPendingLike_(task, now)
        ? 'status_not_pending'
        : !isDue
          ? 'not_due_yet_or_snoozed'
          : !isOwned
            ? 'owner_identity_or_role_mismatch'
            : 'filtered';
      hiddenReasons[reason] = (hiddenReasons[reason] || 0) + 1;
    }
    if (examples.length < 25) {
      examples.push({
        row: task.rowNumber || '',
        taskId: task.taskId || '',
        taskType: task.taskType || '',
        taskTitle: task.taskTitle || '',
        customerName: task.customerName || '',
        ownerRole: task.ownerRole || '',
        intendedOwner: task.intendedOwner || '',
        intendedOwnerEmail: task.intendedOwnerEmail || '',
        currentOwner: task.currentOwner || '',
        currentOwnerEmail: task.currentOwnerEmail || '',
        dueAt: task.dueAt || '',
        status: task.status || '',
        visibleInMine: inMine
      });
    }
  });

  var out = {
    ok: true,
    generatedAt: swIso_(new Date()),
    readOnly: true,
    targetUser: swDiagnosticPublicUser_(target),
    summary: {
      matchedOwnerTasks: matched,
      dueNow: dueNow,
      visibleInMine: visible,
      statusCounts: statusCounts,
      hiddenReasons: hiddenReasons
    },
    examples: examples
  };
  Logger.log('SW_TASK_VISIBILITY_DIAGNOSTIC ' + JSON.stringify(out, null, 2));
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

function swBenchmarkSalesWorkflowSkip_(out, operation, reason) {
  var step = {
    operation: operation,
    ok: true,
    skipped: true,
    reason: reason || 'Skipped.',
    ms: 0
  };
  out.steps.push(step);
  Logger.log('SW_BENCHMARK_STEP ' + JSON.stringify(step));
  return step;
}

function swBenchmarkSalesWorkflowOptions_(options) {
  options = options || {};
  var detailLimit = Number(options.detailLimit || options.maxDetailTasks || 8);
  if (!isFinite(detailLimit) || detailLimit < 0) detailLimit = 8;
  detailLimit = Math.min(Math.floor(detailLimit), 25);
  return {
    detailLimit: detailLimit,
    includeTaskDetails: options.includeTaskDetails !== false,
    includeCustomerDetail: options.includeCustomerDetail !== false,
    customerSearchQueries: swBenchmarkSalesWorkflowStringList_(
      options.customerSearchQueries || options.searchQueries,
      ['']
    ),
    calendarMonths: swBenchmarkSalesWorkflowStringList_(
      options.calendarMonths,
      [swCalendarMonthKey_(new Date())]
    ),
    adminDashboardFilters: swBenchmarkSalesWorkflowFilterList_(
      options.adminDashboardFilters,
      [{ windowPreset: options.windowPreset || 'last7' }]
    )
  };
}

function swBenchmarkSalesWorkflowStringList_(value, fallback) {
  var list = [];
  if (Object.prototype.toString.call(value) === '[object Array]') {
    list = value;
  } else if (value != null) {
    list = [value];
  }
  list = list.map(function (item) { return swTrim_(item); });
  if (!list.length) list = fallback || [];
  return list.slice(0, 5);
}

function swBenchmarkSalesWorkflowFilterList_(value, fallback) {
  var list = Object.prototype.toString.call(value) === '[object Array]'
    ? value
    : value ? [value] : [];
  if (!list.length) list = fallback || [];
  return list.map(function (item) { return item || {}; }).slice(0, 5);
}

function swBenchmarkSalesWorkflowQueueStep_(out, responses, authToken, viewName) {
  swBenchmarkSalesWorkflowStep_(out, 'sw_getMyTasks:' + viewName, function () {
    var res = sw_getMyTasks(authToken, viewName);
    responses.push(res);
    return swBenchmarkSalesWorkflowListSummary_(res);
  });
}

function swBenchmarkSalesWorkflowListSummary_(res) {
  var tasks = res && res.tasks ? res.tasks : [];
  return {
    view: res && res.view ? res.view : '',
    tasks: tasks.length,
    taskTypes: swBenchmarkSalesWorkflowTaskTypeCounts_(tasks)
  };
}

function swBenchmarkSalesWorkflowFirstTaskId_(responses) {
  for (var i = 0; i < responses.length; i++) {
    var tasks = responses[i] && responses[i].tasks ? responses[i].tasks : [];
    if (tasks.length && tasks[0].taskId) return tasks[0].taskId;
  }
  return '';
}

function swBenchmarkSalesWorkflowTaskTypeCounts_(tasks) {
  var counts = {};
  (tasks || []).forEach(function (task) {
    var key = task && task.taskType ? task.taskType : '(blank)';
    counts[key] = (counts[key] || 0) + 1;
  });
  return counts;
}

function swBenchmarkSalesWorkflowDetailSamples_(responses, limit) {
  var out = [];
  var seenTypes = {};
  var seenIds = {};
  responses = responses || [];

  responses.forEach(function (res) {
    var tasks = res && res.tasks ? res.tasks : [];
    tasks.forEach(function (task) {
      if (out.length >= limit || !task || !task.taskId) return;
      var type = task.taskType || '(blank)';
      if (seenTypes[type]) return;
      seenTypes[type] = true;
      seenIds[task.taskId] = true;
      out.push({
        taskId: task.taskId,
        taskType: type,
        view: res.view || '',
        title: task.taskTitle || '',
        customerName: task.customerName || ''
      });
    });
  });

  responses.forEach(function (res) {
    var tasks = res && res.tasks ? res.tasks : [];
    tasks.forEach(function (task) {
      if (out.length >= limit || !task || !task.taskId || seenIds[task.taskId]) return;
      seenIds[task.taskId] = true;
      out.push({
        taskId: task.taskId,
        taskType: task.taskType || '(blank)',
        view: res.view || '',
        title: task.taskTitle || '',
        customerName: task.customerName || ''
      });
    });
  });

  return out;
}

function swBenchmarkSalesWorkflowTaskDetailSummary_(detail, sample) {
  return {
    taskId: sample.taskId,
    sourceView: sample.view || '',
    taskType: detail && detail.task ? detail.task.taskType : sample.taskType,
    customerName: detail && detail.task ? detail.task.customerName : sample.customerName,
    attachments: detail && detail.attachments ? detail.attachments.length : 0,
    appointmentArtifacts: detail && detail.appointmentArtifacts ? detail.appointmentArtifacts.length : 0,
    missingFields: detail && detail.missingFields ? detail.missingFields.length : 0,
    checklist: detail && detail.checklist ? detail.checklist.length : 0,
    formOptionGroups: detail && detail.formOptions ? Object.keys(detail.formOptions).length : 0,
    canComplete: !!(detail && detail.canComplete),
    canClaim: !!(detail && detail.canClaim),
    canAdmin: !!(detail && detail.canAdmin)
  };
}

function swBenchmarkSalesWorkflowCustomerSearchSummary_(res) {
  var columns = res && res.kanban && res.kanban.columns ? res.kanban.columns : [];
  var cards = 0;
  var hidden = 0;
  columns.forEach(function (col) {
    cards += col.cards ? col.cards.length : 0;
    hidden += Number(col.hiddenCount || 0);
  });
  return {
    query: res && res.query ? res.query : '',
    activeOnly: !(res && res.filters && res.filters.activeOnly === false),
    columns: columns.length,
    cards: cards,
    hiddenCards: hidden
  };
}

function swBenchmarkSalesWorkflowFirstCustomerRoot_(responses) {
  for (var i = 0; i < responses.length; i++) {
    var columns = responses[i] && responses[i].kanban && responses[i].kanban.columns
      ? responses[i].kanban.columns
      : [];
    for (var c = 0; c < columns.length; c++) {
      var cards = columns[c].cards || [];
      for (var j = 0; j < cards.length; j++) {
        if (cards[j] && cards[j].root) return cards[j].root;
      }
    }
  }
  return '';
}

function swBenchmarkSalesWorkflowCustomerDetailSummary_(detail) {
  return {
    root: detail && detail.root ? detail.root : '',
    appointments: detail && detail.appointments ? detail.appointments.length : 0,
    tasks: detail && detail.tasks ? detail.tasks.length : 0,
    logs: detail && detail.logs ? detail.logs.length : 0,
    actions: detail && detail.actions ? Object.keys(detail.actions).filter(function (key) {
      return detail.actions[key];
    }).length : 0
  };
}

function swBenchmarkSalesWorkflowRowsSummary_(res) {
  return {
    available: !!(res && res.available !== false),
    rows: res && res.rows ? res.rows.length : 0,
    stats: res && res.stats ? res.stats : {},
    missingColumns: res && res.missingColumns ? res.missingColumns : [],
    tab: res && res.tab ? res.tab : ''
  };
}

function swBenchmarkSalesWorkflowAdminDashboardSummary_(res) {
  var columns = res && res.kanban && res.kanban.columns ? res.kanban.columns : [];
  var cards = 0;
  columns.forEach(function (col) {
    cards += col.cards ? col.cards.length : 0;
  });
  var metrics = res && res.metrics ? res.metrics : {};
  return {
    window: res && res.filters ? res.filters.windowPreset : '',
    windowLabel: res && res.filters ? res.filters.windowLabel : '',
    bookingsCreated: metrics.bookingsCreated || 0,
    paymentsCount: metrics.paymentsCount || 0,
    adminOpenTasks: metrics.adminOpenTasks || 0,
    kanbanCards: cards,
    taskRows: res && res.taskCount != null ? res.taskCount : (res && res.tasks ? res.tasks.length : 0),
    warnings: res && res.warnings ? res.warnings.length : 0
  };
}

function swBenchmarkSalesWorkflowSummary_(steps) {
  var summary = {
    completedSteps: 0,
    skippedSteps: 0,
    failedSteps: 0,
    slowest: []
  };
  (steps || []).forEach(function (step) {
    if (step.skipped) {
      summary.skippedSteps++;
      return;
    }
    summary.completedSteps++;
    if (step.ok === false) summary.failedSteps++;
    summary.slowest.push({
      operation: step.operation,
      ms: step.ms || 0,
      ok: step.ok !== false
    });
  });
  summary.slowest.sort(function (a, b) { return b.ms - a.ms; });
  summary.slowest = summary.slowest.slice(0, 8);
  return summary;
}

function swBenchmarkSalesWorkflowLogSummary_(out) {
  return {
    ok: out.ok,
    generatedAt: out.generatedAt,
    readOnly: out.readOnly,
    totalMs: out.totalMs,
    summary: out.summary,
    options: out.options,
    steps: (out.steps || []).map(function (step) {
      return {
        operation: step.operation,
        ok: step.ok,
        skipped: !!step.skipped,
        reason: step.reason || '',
        error: step.error || '',
        ms: step.ms || 0,
        result: step.result || {}
      };
    })
  };
}

function swBenchmarkSalesWorkflowLabel_(value) {
  var label = swTrim_(value || 'blank').replace(/[^A-Za-z0-9_:-]+/g, '_');
  return label ? label.slice(0, 60) : 'blank';
}

function swDiagnosticTargetUser_(ss, ownerNameOrEmail, fallbackUser) {
  var query = swTrim_(ownerNameOrEmail);
  var ctx = swBuildIdentityContext_(ss, true);
  var email = query.indexOf('@') >= 0 ? swNormEmail_(query) : '';
  var name = email ? swLookupNameByEmail_(ss, email, ctx) : query;
  if (!email && name) email = swLookupEmailByName_(ss, name, ctx);

  var authRow = email ? swAuthFindUserRowReadOnly_(ss, email) : null;
  if (!authRow && name) {
    var authRows = swAuthReadUserRows_(ss, true);
    for (var i = 0; i < authRows.length; i++) {
      if (swNorm_(authRows[i]['Name']) === swNorm_(name)) {
        authRow = authRows[i];
        break;
      }
    }
  }
  if (authRow) return swAuthUserFromRow_(authRow);

  if (!query && fallbackUser) return fallbackUser;
  return {
    email: email,
    name: name || email || query,
    roles: [SW_OWNER_ROLES.SALES_REP],
    isAdmin: false,
    isJoc: false,
    isRep: true,
    isDiamondOrderAdmin: false,
    isDiamondOrderAssistant: false
  };
}

function swDiagnosticTaskOwnerMatches_(task, user) {
  if (!task || !user) return false;
  var email = swNormEmail_(user.email);
  var name = swNorm_(user.name);
  if (email && (swNormEmail_(task.currentOwnerEmail) === email || swNormEmail_(task.intendedOwnerEmail) === email)) return true;
  if (name && (swNorm_(task.currentOwner) === name || swNorm_(task.intendedOwner) === name)) return true;
  return false;
}

function swDiagnosticPublicUser_(user) {
  user = user || {};
  return {
    email: swNormEmail_(user.email),
    name: swTrim_(user.name),
    roles: user.roles || [],
    isAdmin: !!user.isAdmin,
    isJoc: !!user.isJoc,
    isRep: !!user.isRep,
    isDiamondOrderAdmin: !!user.isDiamondOrderAdmin,
    isDiamondOrderAssistant: !!user.isDiamondOrderAssistant
  };
}
