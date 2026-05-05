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
  var rosterSheet = swEnsureSheet_(ss, SW_SHEETS.ROSTER, SW_EMPLOYEE_SCHEDULE_HEADERS);
  var scheduleChangesSheet = swEnsureSheet_(ss, SW_SHEETS.SCHEDULE_CHANGES, SW_SCHEDULE_CHANGE_HEADERS);

  swStyleSheet_(taskSheet);
  swStyleSheet_(logSheet);
  swStyleSheet_(configSheet);
  swStyleSheet_(templateSheet);
  swStyleSheet_(usersSheet);
  swStyleSheet_(cleanupSheet);
  swStyleSheet_(artifactSheet);
  swStyleSheet_(rosterSheet);
  swStyleSheet_(scheduleChangesSheet);

  swSeedConfig_(configSheet);
  if (typeof swClearConfigCache_ === 'function') swClearConfigCache_(ss);
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
      swPrepareClientAdvisorRoundRobin_(ss, ctx, masterRows);
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

        swMaybeAutoAssignClientAdvisor_(ss, ctx, rec, summary);
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
        fn === 'processIntakeQueue' || fn === 'ensureBootstrapForRecentRows_' ||
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
    var user;
    var identityMode = authToken ? 'passwordSession' : 'appsScriptIdentity';
    if (authToken) {
      user = swAuthUserForApi_(ss, authToken);
    } else if (typeof swBuildBootstrapUser_ === 'function') {
      var bootstrapUser = swBuildBootstrapUser_(ss, true);
      user = bootstrapUser.user;
      identityMode = bootstrapUser.lightweight ? 'appsScriptIdentityLightweight' : 'appsScriptIdentity';
    } else {
      user = swAuthUserForApi_(ss, authToken);
    }
    mark('identity', { mode: identityMode });
    return swBuildBootstrapResponse_(ss, user, mark);
  });
}

function swBuildBootstrapResponse_(ss, user, mark) {
  mark = mark || function () {};
  var config = swReadConfig_(ss, true);
  mark('config');
  var cleanupTabEnabled = typeof swDataCleanupCampaignTabEnabled_ === 'function'
    ? swDataCleanupCampaignTabEnabled_(config)
    : false;
  var projected = typeof swReadTaskDashboardBootstrapProjection_ === 'function'
    ? swReadTaskDashboardBootstrapProjection_(ss, user, config)
    : null;
  if (projected && projected.ok) {
    mark('taskListRead', {
      tasks: projected.totalTasks || 0,
      source: projected.source || 'taskDashboardProjection',
      fallbackReason: '',
      ageSeconds: projected.ageSeconds || 0
    });
    mark('currentUser', { isAdmin: user.isAdmin, isJoc: user.isJoc });
    mark('taskBuckets', {
      mine: projected.counts.mine || 0,
      cleanup: projected.counts.cleanup || 0,
      coverage: projected.counts.coverage || 0,
      admin: projected.counts.admin || 0,
      source: projected.source || 'taskDashboardProjection'
    });
    return {
      ok: true,
      user: user,
      tasks: projected.tasks || [],
      counts: {
        mine: projected.counts.mine || 0,
        cleanup: projected.counts.cleanup || 0,
        coverage: projected.counts.coverage || 0,
        admin: user.isAdmin ? projected.counts.admin || 0 : 0
      },
      views: swBootstrapViewsForUser_(user, cleanupTabEnabled),
      message: 'Connected. Use Refresh Queue to create or refresh the queue.'
    };
  }

  var taskRead = swReadTaskListStateForDashboard_(ss, config);
  var state = taskRead.state;
  mark('taskListRead', {
    tasks: state.tasks.length,
    source: taskRead.source,
    fallbackReason: taskRead.fallbackReason || '',
    ageSeconds: taskRead.ageSeconds || 0
  });
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
    views: swBootstrapViewsForUser_(user, cleanupTabEnabled),
    message: 'Connected. Use Refresh Queue to create or refresh the queue.'
  };
}

function swBootstrapViewsForUser_(user, cleanupTabEnabled) {
  user = user || {};
  return {
    mine: true,
    customerSearch: user.isAdmin || user.isJoc || user.isRep,
    calendar: true,
    inStockDiamonds: true,
    diamondTracking: user.isAdmin || user.isDiamondOrderAdmin || user.isDiamondOrderAssistant,
    bulkReturns: user.isAdmin || user.isDiamondOrderAdmin,
    cleanup: !!cleanupTabEnabled,
    coverage: user.isJoc || user.isAdmin,
    adminDashboard: user.isAdmin,
    employeeSchedules: user.isAdmin,
    admin: user.isAdmin
  };
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
    var projected = typeof swReadTaskDashboardViewProjection_ === 'function'
      ? swReadTaskDashboardViewProjection_(ss, user, viewName, config)
      : null;
    if (projected && projected.ok) {
      mark('taskListRead', {
        tasks: projected.totalTasks || 0,
        source: projected.source || 'taskDashboardProjection',
        fallbackReason: '',
        ageSeconds: projected.ageSeconds || 0
      });
      mark('filter', { view: viewName, tasks: (projected.tasks || []).length, source: projected.source || 'taskDashboardProjection' });
      return {
        ok: true,
        view: viewName,
        user: user,
        tasks: projected.tasks || []
      };
    }

    var taskRead = swReadTaskListStateForDashboard_(ss, config);
    var state = taskRead.state;
    mark('taskListRead', {
      tasks: state.tasks.length,
      source: taskRead.source,
      fallbackReason: taskRead.fallbackReason || '',
      ageSeconds: taskRead.ageSeconds || 0
    });
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
    var projected = typeof swTryGetCalendarAppointmentsFromReadModel_ === 'function'
      ? swTryGetCalendarAppointmentsFromReadModel_(ss, month.key)
      : null;
    if (projected && projected.ok) return projected;

    var today = new Date();
    var todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0, 0);
    var rows = swReadAppointments_(ss).filter(function (rec) {
      if (!swIsAppointmentActive_(rec)) return false;
      var visitAt = swVisitDateTime_(rec, tz);
      if (!visitAt) return false;
      if (visitAt.getTime() < todayStart.getTime()) return false;
      return visitAt.getTime() >= month.start.getTime() && visitAt.getTime() < month.end.getTime();
    }).sort(function (a, b) {
      var av = swVisitDateTime_(a, tz);
      var bv = swVisitDateTime_(b, tz);
      return av.getTime() - bv.getTime() || String(a.name).localeCompare(String(b.name));
    });
    var aiBriefByRoot = rows.length && typeof swAppointmentAiBriefIndex_ === 'function'
      ? swAppointmentAiBriefIndex_(ss)
      : {};
    var appointments = rows.map(function (rec) {
      var visitAt = swVisitDateTime_(rec, tz);
      var root = rec.root || rec.appt || '';
      var aiBrief = root && aiBriefByRoot[root] && typeof swAppointmentAiBriefCompact_ === 'function'
        ? swAppointmentAiBriefCompact_(aiBriefByRoot[root])
        : { hasAiBrief: false, reviewFlagCount: 0, latestAiBriefUpdatedAt: '' };
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
        isDiamondViewing: swDiamondIsViewingAppointment_(rec),
        hasAiBrief: !!aiBrief.hasAiBrief,
        reviewFlagCount: Number(aiBrief.reviewFlagCount || 0),
        latestAiBriefUpdatedAt: aiBrief.latestAiBriefUpdatedAt || ''
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

function sw_getAppointmentAiBrief(authToken, rootApptId) {
  return swTimed_('sw_getAppointmentAiBrief', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    swAuthUserForApi_(ss, authToken);
    var root = swTrim_(rootApptId);
    if (!root) throw new Error('Missing RootApptID.');
    var brief = typeof swAppointmentAiBriefForRoot_ === 'function'
      ? swAppointmentAiBriefForRoot_(ss, root)
      : null;
    return {
      ok: true,
      rootApptId: root,
      aiBrief: brief
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

    var config = ss.getSheetByName(SW_SHEETS.CONFIG)
      ? swReadSheetObjectsExpectedHeaders_(ss.getSheetByName(SW_SHEETS.CONFIG), SW_CONFIG_HEADERS)
      : [];
    var projected = typeof swTryGetDiamondTrackingDashboardFromReadModel_ === 'function'
      ? swTryGetDiamondTrackingDashboardFromReadModel_(ss, config)
      : null;
    if (projected && projected.ok) return projected;

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

    var values = swDiamondRead200Rows_(sh, 3, lr - 2, C, lc);
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

    var config = ss.getSheetByName(SW_SHEETS.CONFIG)
      ? swReadSheetObjectsExpectedHeaders_(ss.getSheetByName(SW_SHEETS.CONFIG), SW_CONFIG_HEADERS)
      : [];
    var projected = typeof swTryGetInStockDiamondsFromReadModel_ === 'function'
      ? swTryGetInStockDiamondsFromReadModel_(ss, config)
      : null;
    if (projected && projected.ok) return projected;

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

    var values = swDiamondRead200Rows_(sh, 3, lr - 2, C, lc);
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

    var config = ss.getSheetByName(SW_SHEETS.CONFIG)
      ? swReadSheetObjectsExpectedHeaders_(ss.getSheetByName(SW_SHEETS.CONFIG), SW_CONFIG_HEADERS)
      : [];
    var projected = typeof swTryGetBulkReturnCandidatesFromReadModel_ === 'function'
      ? swTryGetBulkReturnCandidatesFromReadModel_(ss, config)
      : null;
    if (projected && projected.ok) return projected;

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

    var values = swDiamondRead200Rows_(sh, 3, lr - 2, C, lc);
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

      try { if (typeof swInvalidateDiamondReadModelsAfterWrite_ === 'function') swInvalidateDiamondReadModelsAfterWrite_(ss, 'Bulk diamond return marked'); } catch (_) {}
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

function sw_adminGetEmployeeSchedules(authToken) {
  return swTimed_('sw_adminGetEmployeeSchedules', function () {
    var ss = swSpreadsheet_();
    sw_setupSalesWorkflow();
    var user = swAuthUserForApi_(ss, authToken);
    if (!user.isAdmin) throw new Error('Admin access required.');
    return swReadEmployeeScheduleAdminData_(ss);
  });
}

function sw_adminSaveEmployeeSchedules(authToken, data) {
  return swTimed_('sw_adminSaveEmployeeSchedules', function () {
    data = data || {};
    var ss = swSpreadsheet_();
    sw_setupSalesWorkflow();
    var user = swAuthUserForApi_(ss, authToken);
    if (!user.isAdmin) throw new Error('Admin access required.');

    var people = data.people || data.rows || [];
    if (!Array.isArray(people)) throw new Error('Schedule rows must be an array.');
    var canonicalPeople = swCanonicalizeEmployeeScheduleRowsForWrite_(ss, people);
    var written = swWriteEmployeeRosterRows_(ss, canonicalPeople, user);
    if (data.settings && data.settings.clientAdvisorRoundRobin != null) {
      swSetWorkflowConfigValue_(ss, 'SYSTEM', 'CLIENT_ADVISOR_ROUND_ROBIN',
        swTruthy_(data.settings.clientAdvisorRoundRobin) ? 'Y' : 'N');
    }
    var generation = null;
    try { generation = sw_generateSalesWorkflowTasks(); } catch (err) { generation = { ok: false, error: swTrim_(err && err.message || err) }; }
    return {
      ok: true,
      updated: written,
      generation: generation,
      schedule: swReadEmployeeScheduleAdminData_(ss)
    };
  });
}

function sw_adminUpsertScheduleChange(authToken, data) {
  return swTimed_('sw_adminUpsertScheduleChange', function () {
    data = data || {};
    var ss = swSpreadsheet_();
    sw_setupSalesWorkflow();
    var user = swAuthUserForApi_(ss, authToken);
    if (!user.isAdmin) throw new Error('Admin access required.');
    var row = swUpsertScheduleChangeRow_(ss, data, user);
    var generation = null;
    try { generation = sw_generateSalesWorkflowTasks(); } catch (err) { generation = { ok: false, error: swTrim_(err && err.message || err) }; }
    return {
      ok: true,
      rowNumber: row,
      generation: generation,
      schedule: swReadEmployeeScheduleAdminData_(ss)
    };
  });
}

function sw_adminDeleteScheduleChange(authToken, name, date) {
  return swTimed_('sw_adminDeleteScheduleChange', function () {
    var ss = swSpreadsheet_();
    sw_setupSalesWorkflow();
    var user = swAuthUserForApi_(ss, authToken);
    if (!user.isAdmin) throw new Error('Admin access required.');
    var deleted = swDeleteScheduleChangeRow_(ss, name, date);
    var generation = null;
    try { generation = sw_generateSalesWorkflowTasks(); } catch (err) { generation = { ok: false, error: swTrim_(err && err.message || err) }; }
    return {
      ok: true,
      deleted: deleted,
      generation: generation,
      schedule: swReadEmployeeScheduleAdminData_(ss)
    };
  });
}

function sw_adminAuditWorkflowPeopleData(authToken) {
  return swTimed_('sw_adminAuditWorkflowPeopleData', function () {
    var ss = swSpreadsheet_();
    sw_setupSalesWorkflow();
    var user = swAuthUserForApi_(ss, authToken);
    if (!user.isAdmin) throw new Error('Admin access required.');
    return swAuditWorkflowPeopleData_(ss);
  });
}

function sw_adminMigrateWorkflowPeople(authToken, options) {
  return swTimed_('sw_adminMigrateWorkflowPeople', function () {
    var ss = swSpreadsheet_();
    sw_setupSalesWorkflow();
    var user = swAuthUserForApi_(ss, authToken);
    if (!user.isAdmin) throw new Error('Admin access required.');
    options = options || {};
    var dryRun = options.dryRun !== false;
    var actions = [];
    var warnings = [];
    var before = swAuditWorkflowPeopleData_(ss);

    swReadCanonicalWorkflowPeople_(ss, { schedulableOnly: true, includeInactive: true }).forEach(function (person) {
      actions.push('Ensure roster extension for ' + person.name + ' <' + person.email + '>');
      if (!dryRun) {
        try {
          swEnsureOrSyncRosterForWorkflowUser_(ss, person, user);
        } catch (err) {
          warnings.push('Skipped roster extension for ' + person.name + ': ' + swTrim_(err && err.message || err));
        }
      }
    });

    swMigrateWorkflowIdentitySheet_(ss, SW_SHEETS.ROSTER, ['Rep', 'Name', 'Team Member'], ['Email', 'Rep Email'], ['Role', 'Roles'], '', dryRun, actions);
    swMigrateWorkflowIdentitySheet_(ss, SW_SHEETS.SCHEDULE_CHANGES, ['Rep Name', 'Rep', 'Name'], ['Email', 'Rep Email'], ['Role', 'Roles'], '', dryRun, actions);
    swMigrateDefaultJocPairsFromDropdown_(ss, dryRun, actions, warnings);

    if (options.clearDropdownIdentityData) {
      swBackupAndClearDropdownIdentityData_(ss, dryRun, actions, warnings);
    }

    if (!dryRun) {
      try { if (typeof swClearAssignmentOptionsMemoryCache_ === 'function') swClearAssignmentOptionsMemoryCache_(ss); } catch (_) {}
      try { CacheService.getScriptCache().remove('sw:assignmentOptions:v1:' + ss.getId()); } catch (_) {}
    }
    var generation = null;
    if (!dryRun) {
      try { generation = sw_generateSalesWorkflowTasks(); } catch (err) { generation = { ok: false, error: swTrim_(err && err.message || err) }; }
    }
    return {
      ok: true,
      dryRun: dryRun,
      generatedAt: swIso_(new Date()),
      actions: actions,
      warnings: warnings,
      generation: generation,
      auditBefore: before,
      auditAfter: dryRun ? null : swAuditWorkflowPeopleData_(ss)
    };
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
    var canAct = swCanActOnTask_(task, user);

    var payload = swParseJson_(task.payloadJson, {});
    mark('payloadParse');
    var appointmentAiBrief = swTaskDetailAppointmentAiBrief_(ss, task, payload);
    mark('aiBrief', { hasAiBrief: !!(appointmentAiBrief && appointmentAiBrief.hasAiBrief) });
    payload = swTaskDetailHydrateAiBriefPayload_(task, payload, appointmentAiBrief);
    mark('payloadHydrate');
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
    var appointmentArtifacts = swTaskDetailShouldLoadAppointmentArtifacts_(task) && typeof swPublicAppointmentArtifacts_ === 'function'
      ? swPublicAppointmentArtifacts_(ss, task.root || task.appt || '')
      : [];
    mark('appointmentArtifacts', { artifacts: appointmentArtifacts.length });
    var appointmentUploadFolders = canAct && swTaskDetailShouldLoadAppointmentArtifacts_(task)
      ? swCachedAppointmentUploadFoldersForTask_(ss, task)
      : {};
    mark('appointmentUploadFolders', { folders: Object.keys(appointmentUploadFolders).length });
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
      appointmentAiBrief: appointmentAiBrief,
      appointmentUploadFolders: appointmentUploadFolders,
      assignmentOptions: assignmentOptions,
      missingFields: missingFields,
      checklist: checklist,
      canComplete: canAct,
      canClaim: swCanClaimTask_(task, user),
      canAdmin: user.isAdmin
    };
  });
}

function swTaskDetailAppointmentAiBrief_(ss, task, payload) {
  if (!(task && (task.taskType === SW_TASKS.APPROVE || task.taskType === SW_TASKS.FINAL))) return null;
  if (typeof swAppointmentAiBriefForRoot_ !== 'function') return null;
  var root = swTrim_(task.root || task.appt ||
    swDeepValue_(payload, ['appointment', 'root']) ||
    swDeepValue_(payload, ['appointment', 'appt']) || '');
  if (!root) return null;
  return swAppointmentAiBriefForRoot_(ss, root);
}

function swTaskDetailHydrateAiBriefPayload_(task, payload, aiBrief) {
  payload = payload || {};
  if (!(task && aiBrief && aiBrief.hasAiBrief)) return payload;
  payload.extra = payload.extra || {};
  var extra = payload.extra;
  extra.artifactId = extra.artifactId || aiBrief.artifactId || '';
  extra.workflowStage = extra.workflowStage || aiBrief.workflowStage || '';
  extra.transcriptDocUrl = extra.transcriptDocUrl || aiBrief.transcriptDocUrl || '';
  extra.summaryDocUrl = extra.summaryDocUrl || aiBrief.summaryDocUrl || '';
  extra.summaryJsonUrl = extra.summaryJsonUrl || aiBrief.summaryJsonUrl || '';
  extra.salesBrief = aiBrief.salesBrief || extra.salesBrief || '';
  extra.reviewFlags = (aiBrief.reviewFlags || []).length ? aiBrief.reviewFlags.join('\n') : (extra.reviewFlags || '');
  if (task.taskType === SW_TASKS.APPROVE) {
    extra.clientFollowUpDraft = aiBrief.clientFollowUpDraft || extra.clientFollowUpDraft || '';
    extra.recapDraft = aiBrief.clientFollowUpDraft || extra.recapDraft || '';
  }
  return payload;
}

function swTaskDetailShouldLoadAppointmentArtifacts_(task) {
  return task && task.taskType === SW_TASKS.CHECKLIST;
}

function swCachedAppointmentUploadFoldersForTask_(ss, task) {
  var out = {};
  if (typeof swDriveUploadArtifactTypes_ !== 'function' ||
      typeof swCachedAppointmentUploadFolderInfo_ !== 'function') return out;
  var root = swTrim_(task && (task.root || task.appt) || '');
  if (!root) return out;
  swDriveUploadArtifactTypes_().forEach(function (type) {
    var info = swCachedAppointmentUploadFolderInfo_(ss, root, type);
    if (info && info.url) out[type] = info;
  });
  return out;
}

function sw_getAppointmentUploadFolder(authToken, taskId, artifactType) {
  return swTimed_('sw_getAppointmentUploadFolder', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
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
    if (typeof swAppointmentDriveDropFolderInfoForRoot_ === 'function') {
      return swAppointmentDriveDropFolderInfoForRoot_(ss, root, type);
    }
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
    var automation = null;
    var automationError = '';
    try {
      if (typeof sw_processAppointmentAutomation === 'function') {
        automation = sw_processAppointmentAutomation();
      }
    } catch (err) {
      automationError = swTrim_(err && err.message || err);
    }
    return {
      ok: true,
      rootApptId: root,
      registered: created.length,
      automation: automation,
      automationError: automationError,
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
  if (appointmentAction || approvalAction || jocHandoffAction) {
    try { if (typeof swInvalidateAppointmentReadModelsAfterWrite_ === 'function') swInvalidateAppointmentReadModelsAfterWrite_(ss, 'Appointment task completion updated source data'); } catch (_) {}
  }
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

/**
 * Read-only client diagnostic: records browser-side startup timing from the
 * HtmlService web app. This intentionally does not authenticate or touch sheets
 * so the logging call cannot slow down startup.
 */
function sw_logClientLoadTiming(authToken, payload) {
  if (authToken && typeof authToken === 'object' && payload == null) {
    payload = authToken;
    authToken = '';
  }
  var out = swSanitizeClientLoadTiming_(payload || {});
  out.loggedAt = swIso_(new Date());
  out.tokenPresent = !!authToken || !!out.tokenPresent;
  Logger.log('SW_CLIENT_LOAD_TIMING ' + JSON.stringify(out));
  return { ok: true };
}

function swSanitizeClientLoadTiming_(payload) {
  payload = payload || {};
  return {
    event: swClientLoadString_(payload.event, 60),
    build: swClientLoadString_(payload.build, 80),
    clientAt: swClientLoadString_(payload.clientAt, 40),
    mode: swClientLoadString_(payload.mode, 40),
    finalState: swClientLoadString_(payload.finalState, 40),
    tokenPresentOnInit: !!payload.tokenPresentOnInit,
    tokenPresent: !!payload.tokenPresent,
    visibilityState: swClientLoadString_(payload.visibilityState, 40),
    documentHidden: !!payload.documentHidden,
    hasFocus: !!payload.hasFocus,
    userEmail: swClientLoadString_(payload.userEmail, 120),
    view: swClientLoadString_(payload.view, 40),
    taskCount: swClientLoadNumber_(payload.taskCount),
    counts: swClientLoadNumberMap_(payload.counts, 12),
    timings: swClientLoadNumberMap_(payload.timings, 24),
    marks: swClientLoadNumberMap_(payload.marks, 40),
    navigation: swClientLoadNumberMap_(payload.navigation, 12),
    viewport: swClientLoadNumberMap_(payload.viewport, 8),
    userAgent: swClientLoadString_(payload.userAgent, 220)
  };
}

function swClientLoadString_(value, maxLen) {
  return swTrim_(value).slice(0, maxLen || 120);
}

function swClientLoadNumber_(value) {
  value = Number(value || 0);
  return isFinite(value) ? Math.round(value) : 0;
}

function swClientLoadNumberMap_(map, maxKeys) {
  var out = {};
  map = map || {};
  Object.keys(map).slice(0, maxKeys || 20).forEach(function (key) {
    out[swClientLoadString_(key, 60)] = swClientLoadNumber_(map[key]);
  });
  return out;
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

  var reason = swTrim_(data.reason);
  var options = swReadAssignmentOptions_(ss);
  var assignedOwner = swAssignmentCanonicalOwner_(data.assignedRep, data.assignedRepEmail, options.salesReps, 'Client Advisor');
  var assistedOwner = swAssignmentCanonicalOwner_(data.assistedRep, data.assistedRepEmail, options.jocReps, 'JOC');
  var assignedName = assignedOwner.name;
  var assignedEmail = assignedOwner.email;
  var assistedName = assistedOwner.name;
  var assistedEmail = assistedOwner.email;

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
  try { if (typeof swInvalidateAppointmentReadModelsAfterWrite_ === 'function') swInvalidateAppointmentReadModelsAfterWrite_(ss, 'Appointment owners assigned'); } catch (_) {}
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

var SW_ASSIGNMENT_OPTIONS_MEMORY_CACHE_ = {};

function swClearAssignmentOptionsMemoryCache_(ss) {
  try { delete SW_ASSIGNMENT_OPTIONS_MEMORY_CACHE_['sw:assignmentOptions:v1:' + ss.getId()]; } catch (_) {}
}

function swReadAssignmentOptions_(ss) {
  var cacheKey = '';
  try {
    cacheKey = 'sw:assignmentOptions:v1:' + ss.getId();
    var memory = SW_ASSIGNMENT_OPTIONS_MEMORY_CACHE_[cacheKey];
    if (memory && memory.expiresAt > new Date().getTime()) return memory.value;
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) {
      var cachedValue = swParseJson_(cached, { salesReps: [], jocReps: [] });
      SW_ASSIGNMENT_OPTIONS_MEMORY_CACHE_[cacheKey] = {
        expiresAt: new Date().getTime() + 300000,
        value: cachedValue
      };
      return cachedValue;
    }
  } catch (_) {}
  var out = { salesReps: [], jocReps: [] };
  try {
    swReadCanonicalWorkflowPeople_(ss, { schedulableOnly: true, activeOnly: true }).forEach(function (row) {
      var roles = row.roles || swAuthRoles_(row.role || '');
      var item = {
        name: row.name || row.email,
        email: row.email
      };
      if (swAuthHasRole_(roles, SW_OWNER_ROLES.SALES_REP)) swPushAssignmentOption_(out.salesReps, item);
      if (swAuthHasRole_(roles, 'JOC')) swPushAssignmentOption_(out.jocReps, item);
    });
  } catch (_) {}

  out.salesReps.sort(swAssignmentOptionSort_);
  out.jocReps.sort(swAssignmentOptionSort_);
  if (cacheKey) {
    try {
      SW_ASSIGNMENT_OPTIONS_MEMORY_CACHE_[cacheKey] = {
        expiresAt: new Date().getTime() + 300000,
        value: out
      };
    } catch (_) {}
  }
  if (cacheKey) {
    try {
      var json = JSON.stringify(out);
      if (json.length <= 90000) CacheService.getScriptCache().put(cacheKey, json, 300);
    } catch (_) {}
  }
  return out;
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

function swAssignmentCanonicalOwner_(name, email, options, label) {
  name = swTrim_(name || '');
  email = swNormEmail_(email || '');
  if (!name && !email) return { name: '', email: '' };
  var byName = null;
  var byEmail = null;
  (options || []).forEach(function (option) {
    if (!option) return;
    if (name && swNorm_(option.name) === swNorm_(name)) byName = byName || option;
    if (email && swNormEmail_(option.email) === email) byEmail = byEmail || option;
  });
  if (byName && byEmail && swNormEmail_(byName.email) !== swNormEmail_(byEmail.email)) {
    throw new Error(label + ' name/email conflict: "' + name + '" does not match ' + email + '.');
  }
  var owner = byEmail || byName;
  if (!owner) throw new Error(label + ' must be an active canonical workflow user: ' + (name || email));
  return {
    name: owner.name || name || email,
    email: swNormEmail_(owner.email || email)
  };
}

function swAssignmentOptionSort_(a, b) {
  return String(a.name || '').localeCompare(String(b.name || ''));
}

function swWriteEmployeeRosterRows_(ss, people, actor) {
  var sh = swEnsureEmployeeScheduleSheets_(ss).roster;
  var headers = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getDisplayValues()[0].map(function (h) {
    return swTrim_(h);
  });
  var now = swIso_(new Date());
  var actorLabel = actor ? (actor.name || actor.email || '') : '';
  var rows = [];
  var seen = {};

  (people || []).forEach(function (person) {
    person = person || {};
    var name = swTrim_(person.name || person.rep || person['Rep']);
    if (!name) return;
    var key = swNorm_(name);
    if (seen[key]) return;
    seen[key] = true;
    var days = person.days || {};
    var role = swNormalizeEmployeeRoleList_(person.role || person.roles || '');
    var active = person.active == null ? true : swTruthy_(person.active);
    var coverageEnabled = person.coverageEnabled == null ? true : swTruthy_(person.coverageEnabled);
    var skills = person.skills || {};
    var valuesByHeaderKey = {
      rep: name,
      name: name,
      teammember: name,
      email: swNormEmail_(person.email || ''),
      repemail: swNormEmail_(person.email || ''),
      role: role,
      roles: role,
      active: active ? 'Y' : 'N',
      mon: swTruthy_(days.Mon) ? 'Y' : 'N',
      tue: swTruthy_(days.Tue) ? 'Y' : 'N',
      wed: swTruthy_(days.Wed) ? 'Y' : 'N',
      thu: swTruthy_(days.Thu) ? 'Y' : 'N',
      fri: swTruthy_(days.Fri) ? 'Y' : 'N',
      sat: swTruthy_(days.Sat) ? 'Y' : 'N',
      sun: swTruthy_(days.Sun) ? 'Y' : 'N',
      defaultjoc: swTrim_(person.defaultJoc || ''),
      linkedjoc: swTrim_(person.defaultJoc || ''),
      jocpartner: swTrim_(person.defaultJoc || ''),
      assistedcoverageenabled: coverageEnabled ? 'Y' : 'N',
      coverageenabled: coverageEnabled ? 'Y' : 'N',
      assistedcoveragepartner: swTrim_(person.coveragePartner || ''),
      coveragepartner: swTrim_(person.coveragePartner || ''),
      labdiamond: swTruthy_(skills.labDiamond) ? 'Y' : 'N',
      lab: swTruthy_(skills.labDiamond) ? 'Y' : 'N',
      naturaldiamond: swNormalizeNaturalSkill_(skills.naturalDiamond || 'None'),
      natural: swNormalizeNaturalSkill_(skills.naturalDiamond || 'None'),
      generalappointment: swTruthy_(skills.generalAppointment) ? 'Y' : 'N',
      general: swTruthy_(skills.generalAppointment) ? 'Y' : 'N',
      skillnotes: swTrim_(person.skillNotes || ''),
      updatedat: now,
      updatedby: actorLabel
    };
    rows.push(headers.map(function (header) {
      var headerKey = swHeaderKey_(header);
      return valuesByHeaderKey[headerKey] == null ? '' : valuesByHeaderKey[headerKey];
    }));
  });

  var oldRows = Math.max(0, sh.getLastRow() - 1);
  if (oldRows > 0) sh.getRange(2, 1, oldRows, sh.getLastColumn()).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  return rows.length;
}

function swUpsertScheduleChangeRow_(ss, data, actor) {
  var sh = swEnsureEmployeeScheduleSheets_(ss).changes;
  var headers = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getDisplayValues()[0].map(function (h) {
    return swTrim_(h);
  });
  var name = swTrim_(data.name || data.rep || data['Rep Name']);
  var date = swScheduleDateKey_(data.date || data.changeDate || data['Change Date']);
  if (!name) throw new Error('Select an employee.');
  if (!date) throw new Error('Select a valid change date.');
  var peopleIndex = swCanonicalWorkflowPeopleIndex_(ss, { schedulableOnly: true, includeInactive: true });
  var email = swNormEmail_(data.email || '');
  var canonicalUser = (email && peopleIndex.byEmail[email]) || peopleIndex.byName[swNorm_(name)] || null;
  if (!canonicalUser) throw new Error('Schedule override employee must be a workflow user with Client Advisor or JOC access.');
  name = canonicalUser.name;
  email = canonicalUser.email;
  var role = canonicalUser.scheduleRole || swWorkflowSchedulableRoleList_(canonicalUser.roles || '');
  var changeType = swTrim_(data.changeType || data.status || 'Full-day off');
  var now = swIso_(new Date());
  var actorLabel = actor ? (actor.name || actor.email || '') : '';
  var valuesByHeaderKey = {
    repname: name,
    rep: name,
    name: name,
    email: email,
    repemail: email,
    role: role,
    roles: role,
    changedate: date,
    date: date,
    changetype: changeType,
    status: changeType,
    overridestatus: changeType,
    availablefrom: swTrim_(data.availableFrom || ''),
    from: swTrim_(data.availableFrom || ''),
    availableuntil: swTrim_(data.availableUntil || ''),
    until: swTrim_(data.availableUntil || ''),
    notes: swTrim_(data.notes || ''),
    note: swTrim_(data.notes || ''),
    updatedat: now,
    updatedby: actorLabel
  };
  var rowNumber = swFindScheduleChangeRow_(sh, name, date);
  if (!rowNumber) rowNumber = sh.getLastRow() + 1;
  sh.getRange(rowNumber, 1, 1, headers.length).setValues([headers.map(function (header) {
    var headerKey = swHeaderKey_(header);
    return valuesByHeaderKey[headerKey] == null ? '' : valuesByHeaderKey[headerKey];
  })]);
  return rowNumber;
}

function swDeleteScheduleChangeRow_(ss, name, date) {
  var sh = swEnsureEmployeeScheduleSheets_(ss).changes;
  var rowNumber = swFindScheduleChangeRow_(sh, name, date);
  if (!rowNumber) return false;
  sh.deleteRow(rowNumber);
  return true;
}

function swFindScheduleChangeRow_(sh, name, date) {
  name = swNorm_(name);
  date = swScheduleDateKey_(date);
  if (!name || !date || !sh || sh.getLastRow() < 2) return 0;
  var values = sh.getDataRange().getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var nameCol = swPickIndex_(H, ['Rep Name', 'Rep', 'Name']);
  var dateCol = swPickIndex_(H, ['Change Date', 'Date']);
  if (nameCol < 0 || dateCol < 0) return 0;
  for (var i = 1; i < values.length; i++) {
    if (swNorm_(values[i][nameCol]) === name && swScheduleDateKey_(values[i][dateCol]) === date) return i + 1;
  }
  return 0;
}

function swScheduleEmailForName_(ss, name) {
  var target = swNorm_(name);
  if (!target) return '';
  var people = swReadEmployeeSchedulePeople_(ss);
  for (var i = 0; i < people.length; i++) {
    if (swNorm_(people[i].name) === target) return swNormEmail_(people[i].email);
  }
  return '';
}

function swAuditWorkflowPeopleData_(ss) {
  var userIndex = swWorkflowAuditUserIndex_(ss);
  var out = {
    ok: true,
    generatedAt: swIso_(new Date()),
    counts: {
      users: userIndex.users.length,
      activeSchedulableUsers: userIndex.users.filter(function (u) { return u.active && u.scheduleRole; }).length,
      rosterRows: 0,
      scheduleChangeRows: 0,
      dropdownIdentities: 0,
      activeAppointments: 0
    },
    duplicateActiveEmails: [],
    duplicateActiveNames: [],
    conflicts: [],
    orphanRosterRows: [],
    orphanScheduleChanges: [],
    dropdownOnlyIdentities: [],
    appointmentOwnersNotMapped: []
  };

  Object.keys(userIndex.activeEmails).forEach(function (email) {
    if (userIndex.activeEmails[email].length > 1) out.duplicateActiveEmails.push({ email: email, users: userIndex.activeEmails[email] });
  });
  Object.keys(userIndex.activeScheduleNames).forEach(function (name) {
    if (userIndex.activeScheduleNames[name].length > 1) out.duplicateActiveNames.push({ name: name, users: userIndex.activeScheduleNames[name] });
  });

  swReadEmployeeRosterRows_(ss).forEach(function (row) {
    out.counts.rosterRows++;
    var match = swAuditResolveCanonicalUser_(userIndex, row.email, row.name, '');
    if (!match.user) out.orphanRosterRows.push(swAuditIdentityIssue_('roster', row, match.reason));
    else swAuditPushIdentityConflict_(out, 'roster', row, match);
  });

  swReadEmployeeScheduleChanges_(ss).forEach(function (row) {
    out.counts.scheduleChangeRows++;
    var match = swAuditResolveCanonicalUser_(userIndex, row.email, row.name, '');
    if (!match.user) out.orphanScheduleChanges.push(swAuditIdentityIssue_('scheduleChange', row, match.reason));
    else swAuditPushIdentityConflict_(out, 'scheduleChange', row, match);
  });

  swReadDropdownIdentityEntries_(ss).forEach(function (entry) {
    out.counts.dropdownIdentities++;
    var match = swAuditResolveCanonicalUser_(userIndex, entry.email, entry.name, entry.role);
    if (!match.user) out.dropdownOnlyIdentities.push(swAuditIdentityIssue_('dropdown', entry, match.reason));
    else swAuditPushIdentityConflict_(out, 'dropdown', entry, match);
  });

  try {
    swReadAppointments_(ss).forEach(function (rec) {
      if (typeof swIsAppointmentActive_ === 'function' && !swIsAppointmentActive_(rec)) return;
      out.counts.activeAppointments++;
      swAuditAppointmentOwner_(out, userIndex, rec, SW_OWNER_ROLES.SALES_REP, rec.assignedRep, rec.assignedRepEmail);
      swAuditAppointmentOwner_(out, userIndex, rec, SW_OWNER_ROLES.JOC, rec.assistedRep, rec.assistedRepEmail);
    });
  } catch (err) {
    out.conflicts.push({ source: 'appointments', reason: 'APPOINTMENT_AUDIT_FAILED', message: swTrim_(err && err.message || err) });
  }

  return out;
}

function swWorkflowAuditUserIndex_(ss) {
  var users = [];
  var byEmail = {};
  var byEmailAll = {};
  var byName = {};
  var activeEmails = {};
  var activeScheduleNames = {};
  swAuthReadUserRows_(ss, true).forEach(function (row) {
    var email = swNormEmail_(row['Email']);
    if (!email) return;
    var roles = swAuthRoles_(row['Roles']);
    var user = {
      rowNumber: row.__rowNumber || 0,
      email: email,
      name: swTrim_(row['Name']) || email,
      roles: roles,
      role: roles.join(','),
      scheduleRole: swWorkflowSchedulableRoleList_(roles),
      active: swWorkflowUserActive_(row)
    };
    users.push(user);
    if (!byEmail[email]) byEmail[email] = user;
    if (!byEmailAll[email]) byEmailAll[email] = [];
    byEmailAll[email].push(user);
    if (!byName[swNorm_(user.name)]) byName[swNorm_(user.name)] = [];
    byName[swNorm_(user.name)].push(user);
    if (user.active) {
      if (!activeEmails[email]) activeEmails[email] = [];
      activeEmails[email].push({ name: user.name, email: user.email, roles: user.role });
      if (user.scheduleRole) {
        var nameKey = swNorm_(user.name);
        if (!activeScheduleNames[nameKey]) activeScheduleNames[nameKey] = [];
        activeScheduleNames[nameKey].push({ name: user.name, email: user.email, roles: user.role });
      }
    }
  });
  return { users: users, byEmail: byEmail, byEmailAll: byEmailAll, byName: byName, activeEmails: activeEmails, activeScheduleNames: activeScheduleNames };
}

function swAuditResolveCanonicalUser_(userIndex, email, name, requiredRole) {
  email = swNormEmail_(email || '');
  name = swTrim_(name || '');
  var user = null;
  if (email) {
    var emailMatches = (userIndex.byEmailAll && userIndex.byEmailAll[email]) || (userIndex.byEmail[email] ? [userIndex.byEmail[email]] : []);
    var activeEmailMatches = emailMatches.filter(function (candidate) { return candidate.active !== false; });
    if (activeEmailMatches.length > 1) return { user: null, reason: 'DUPLICATE_EMAIL_MATCHES' };
    if (activeEmailMatches.length === 1) user = activeEmailMatches[0];
    else if (emailMatches.length === 1) user = emailMatches[0];
    else if (emailMatches.length > 1) return { user: null, reason: 'DUPLICATE_EMAIL_MATCHES' };
  }
  var matchType = user ? 'email' : '';
  if (!user && name) {
    var matches = userIndex.byName[swNorm_(name)] || [];
    var activeNameMatches = matches.filter(function (candidate) { return candidate.active !== false; });
    if (activeNameMatches.length === 1) {
      user = activeNameMatches[0];
      matchType = 'name';
    } else if (activeNameMatches.length > 1) {
      return { user: null, reason: 'DUPLICATE_NAME_MATCHES' };
    } else if (matches.length === 1) {
      user = matches[0];
      matchType = 'name';
    } else if (matches.length > 1) {
      return { user: null, reason: 'DUPLICATE_NAME_MATCHES' };
    }
  }
  if (!user) return { user: null, reason: email ? 'EMAIL_NOT_FOUND' : 'NAME_NOT_FOUND' };
  var roleMismatch = requiredRole && !swWorkflowUserHasSchedulableRole_(user, requiredRole);
  return {
    user: user,
    matchType: matchType,
    reason: '',
    roleMismatch: !!roleMismatch,
    inactive: user.active === false,
    nameMismatch: name && swNorm_(name) !== swNorm_(user.name),
    emailMismatch: email && email !== user.email
  };
}

function swAuditIdentityIssue_(source, row, reason) {
  return {
    source: source,
    rowNumber: row.rowNumber || row.sourceRow || '',
    name: row.name || '',
    email: row.email || '',
    role: row.role || '',
    reason: reason || 'NO_CANONICAL_USER'
  };
}

function swAuditPushIdentityConflict_(out, source, row, match) {
  if (source === 'roster' && row && row.active === false && match && match.inactive) return;
  var reasons = [];
  if (match.roleMismatch) reasons.push('ROLE_MISMATCH');
  if (match.inactive) reasons.push('INACTIVE_USER');
  if (match.nameMismatch) reasons.push('NAME_MISMATCH');
  if (match.emailMismatch) reasons.push('EMAIL_MISMATCH');
  if (!reasons.length) return;
  out.conflicts.push({
    source: source,
    rowNumber: row.rowNumber || row.sourceRow || '',
    name: row.name || '',
    email: row.email || '',
    canonicalName: match.user.name,
    canonicalEmail: match.user.email,
    canonicalRoles: match.user.role,
    reasons: reasons
  });
}

function swAuditAppointmentOwner_(out, userIndex, rec, role, name, email) {
  name = swTrim_(name || '');
  email = swNormEmail_(email || '');
  if (!name && !email) return;
  var match = swAuditResolveCanonicalUser_(userIndex, email, name, role);
  if (!match.user || match.roleMismatch || match.inactive || match.nameMismatch || match.emailMismatch) {
    out.appointmentOwnersNotMapped.push({
      rowNumber: rec.row,
      root: rec.root || '',
      appt: rec.appt || '',
      customerName: rec.name || '',
      role: role,
      name: name,
      email: email,
      reason: !match.user ? match.reason : [match.roleMismatch && 'ROLE_MISMATCH', match.inactive && 'INACTIVE_USER', match.nameMismatch && 'NAME_MISMATCH', match.emailMismatch && 'EMAIL_MISMATCH'].filter(Boolean).join(',')
    });
  }
}

function swMigrateWorkflowIdentitySheet_(ss, sheetName, nameAliases, emailAliases, roleAliases, requiredRole, dryRun, actions) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return 0;
  var values = sh.getDataRange().getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var nameCol = swPickIndex_(H, nameAliases);
  var emailCol = swPickIndex_(H, emailAliases);
  var roleCol = roleAliases && roleAliases.length ? swPickIndex_(H, roleAliases) : -1;
  if (nameCol < 0 && emailCol < 0) return 0;
  var userIndex = swWorkflowAuditUserIndex_(ss);
  var updated = 0;
  for (var i = 1; i < values.length; i++) {
    var name = nameCol >= 0 ? swTrim_(values[i][nameCol]) : '';
    var email = emailCol >= 0 ? swNormEmail_(values[i][emailCol]) : '';
    if (!name && !email) continue;
    var match = swAuditResolveCanonicalUser_(userIndex, email, name, requiredRole || '');
    if (!match.user || match.roleMismatch || match.inactive) continue;
    var changes = [];
    if (nameCol >= 0 && name !== match.user.name) changes.push({ col: nameCol + 1, value: match.user.name, label: 'name' });
    if (emailCol >= 0 && email !== match.user.email) changes.push({ col: emailCol + 1, value: match.user.email, label: 'email' });
    if (roleCol >= 0 && swTrim_(values[i][roleCol]) !== match.user.scheduleRole) changes.push({ col: roleCol + 1, value: match.user.scheduleRole, label: 'role' });
    if (!changes.length) continue;
    updated++;
    actions.push((dryRun ? 'Would update ' : 'Updated ') + sheetName + ' row ' + (i + 1) + ' identity to ' + match.user.name + ' <' + match.user.email + '>');
    if (!dryRun) changes.forEach(function (change) { sh.getRange(i + 1, change.col).setValue(change.value); });
  }
  return updated;
}

function swMigrateDefaultJocPairsFromDropdown_(ss, dryRun, actions, warnings) {
  var pairs = swReadDropdownAdvisorJocPairs_(ss);
  if (!pairs.length) return 0;
  var userIndex = swWorkflowAuditUserIndex_(ss);
  var roster = swEnsureEmployeeScheduleSheets_(ss).roster;
  var headers = roster.getRange(1, 1, 1, Math.max(1, roster.getLastColumn())).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var defaultJocCol = swPickIndex_(H, ['Default JOC', 'Linked JOC', 'JOC Partner']);
  if (defaultJocCol < 0) return 0;
  var updated = 0;
  pairs.forEach(function (pair) {
    var advisor = swAuditResolveCanonicalUser_(userIndex, pair.advisorEmail, pair.advisorName, SW_OWNER_ROLES.SALES_REP);
    var joc = swAuditResolveCanonicalUser_(userIndex, pair.jocEmail, pair.jocName, SW_OWNER_ROLES.JOC);
    if (!advisor.user || advisor.roleMismatch || advisor.inactive || !joc.user || joc.roleMismatch || joc.inactive) return;
    var rowNumber = swFindRosterRowForWorkflowUser_(roster, advisor.user.email, advisor.user.name);
    if (!rowNumber) {
      if (!dryRun) swEnsureOrSyncRosterForWorkflowUser_(ss, advisor.user, swSystemUser_());
      rowNumber = dryRun ? 0 : swFindRosterRowForWorkflowUser_(roster, advisor.user.email, advisor.user.name);
    }
    if (!rowNumber && !dryRun) return;
    var current = rowNumber ? swTrim_(roster.getRange(rowNumber, defaultJocCol + 1).getDisplayValue()) : '';
    if (current) return;
    updated++;
    actions.push((dryRun ? 'Would set ' : 'Set ') + advisor.user.name + ' default JOC to ' + joc.user.name + ' from Dropdown row ' + pair.sourceRow);
    if (!dryRun && rowNumber) roster.getRange(rowNumber, defaultJocCol + 1).setValue(joc.user.name);
  });
  if (!updated && pairs.length) warnings.push('No default JOC pairs needed migration from Dropdown.');
  return updated;
}

function swReadDropdownIdentityEntries_(ss) {
  var out = [];
  var sh = ss.getSheetByName(SW_SHEETS.DROPDOWN);
  if (!sh || sh.getLastRow() < 2) return out;
  var values = sh.getDataRange().getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var groups = [
    { role: SW_OWNER_ROLES.SALES_REP, names: ['Client Advisor', 'Assigned Rep'], emails: ['Client Advisor Email', 'Assigned Rep Email'] },
    { role: SW_OWNER_ROLES.JOC, names: ['Assisted Rep', 'Assistant Rep', 'JOC'], emails: ['Assisted Rep Email', 'Assistant Rep Email', 'JOC Email'] }
  ];
  groups.forEach(function (group) {
    var nameCol = swPickIndex_(H, group.names);
    var emailCol = swPickIndex_(H, group.emails);
    if (nameCol < 0 && emailCol < 0) return;
    for (var i = 1; i < values.length; i++) {
      var name = nameCol >= 0 ? swTrim_(values[i][nameCol]) : '';
      var email = emailCol >= 0 ? swNormEmail_(values[i][emailCol]) : '';
      if (!name && !email) continue;
      out.push({ sourceRow: i + 1, rowNumber: i + 1, role: group.role, name: name, email: email });
    }
  });
  return out;
}

function swReadDropdownAdvisorJocPairs_(ss) {
  var out = [];
  var sh = ss.getSheetByName(SW_SHEETS.DROPDOWN);
  if (!sh || sh.getLastRow() < 2) return out;
  var values = sh.getDataRange().getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var advisorCol = swPickIndex_(H, ['Client Advisor', 'Assigned Rep']);
  var advisorEmailCol = swPickIndex_(H, ['Client Advisor Email', 'Assigned Rep Email']);
  var jocCol = swPickIndex_(H, ['JOC', 'Assisted Rep', 'Assistant Rep']);
  var jocEmailCol = swPickIndex_(H, ['JOC Email', 'Assisted Rep Email', 'Assistant Rep Email']);
  if (advisorCol < 0 || jocCol < 0) return out;
  for (var i = 1; i < values.length; i++) {
    var advisorName = swTrim_(values[i][advisorCol]);
    var jocName = swTrim_(values[i][jocCol]);
    if (!advisorName || !jocName) continue;
    out.push({
      sourceRow: i + 1,
      advisorName: advisorName,
      advisorEmail: advisorEmailCol >= 0 ? swNormEmail_(values[i][advisorEmailCol]) : '',
      jocName: jocName,
      jocEmail: jocEmailCol >= 0 ? swNormEmail_(values[i][jocEmailCol]) : ''
    });
  }
  return out;
}

function swBackupAndClearDropdownIdentityData_(ss, dryRun, actions, warnings) {
  var sh = ss.getSheetByName(SW_SHEETS.DROPDOWN);
  if (!sh || sh.getLastRow() < 2) return 0;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
  var identityCols = swDropdownIdentityColumnIndexes_(headers);
  if (!identityCols.length) return 0;
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();
  var backupRows = [];
  values.forEach(function (row, i) {
    var hasIdentity = identityCols.some(function (col) { return swTrim_(row[col]); });
    if (!hasIdentity) return;
    var backup = [swIso_(new Date()), i + 2];
    identityCols.forEach(function (col) { backup.push(row[col]); });
    backupRows.push(backup);
  });
  if (!backupRows.length) {
    warnings.push('Dropdown identity columns were already empty.');
    return 0;
  }
  actions.push((dryRun ? 'Would back up and clear ' : 'Backed up and cleared ') + backupRows.length + ' Dropdown identity row(s).');
  if (dryRun) return backupRows.length;
  var backupName = '_SW_DropdownIdentityBackup';
  var backupSheet = ss.getSheetByName(backupName) || ss.insertSheet(backupName);
  if (backupSheet.getLastRow() === 0) {
    backupSheet.getRange(1, 1, 1, 2 + identityCols.length).setValues([['Backed Up At', 'Source Row'].concat(identityCols.map(function (col) {
      return headers[col] || ('Column ' + (col + 1));
    }))]);
  }
  backupSheet.getRange(backupSheet.getLastRow() + 1, 1, backupRows.length, backupRows[0].length).setValues(backupRows);
  identityCols.forEach(function (col) {
    sh.getRange(2, col + 1, sh.getLastRow() - 1, 1).clearContent();
  });
  swStyleSheet_(backupSheet);
  return backupRows.length;
}

function swDropdownIdentityColumnIndexes_(headers) {
  var H = swHeaderMapFromArray_(headers || []);
  var names = [
    'Client Advisor', 'Assigned Rep', 'Client Advisor Email', 'Assigned Rep Email',
    'Assisted Rep', 'Assistant Rep', 'Assisted Rep Email', 'Assistant Rep Email',
    'JOC', 'JOC Email'
  ];
  var seen = {};
  var out = [];
  names.forEach(function (name) {
    var col = swPickIndex_(H, [name]);
    if (col >= 0 && !seen[col]) {
      seen[col] = true;
      out.push(col);
    }
  });
  out.sort(function (a, b) { return a - b; });
  return out;
}

function swSetWorkflowConfigValue_(ss, section, key, value) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.CONFIG, SW_CONFIG_HEADERS);
  var targetSection = swNorm_(section);
  var targetKey = swNorm_(key);
  if (sh.getLastRow() >= 2) {
    var values = sh.getRange(2, 1, sh.getLastRow() - 1, SW_CONFIG_HEADERS.length).getDisplayValues();
    for (var i = 0; i < values.length; i++) {
      if (swNorm_(values[i][0]) === targetSection && swNorm_(values[i][1]) === targetKey) {
        sh.getRange(i + 2, 3).setValue(value);
        if (typeof swClearConfigCache_ === 'function') swClearConfigCache_(ss);
        return i + 2;
      }
    }
  }
  var row = ['', '', '', '', '', '', 'Y', '', ''];
  row[0] = section;
  row[1] = key;
  row[2] = value;
  sh.appendRow(row);
  if (typeof swClearConfigCache_ === 'function') swClearConfigCache_(ss);
  return sh.getLastRow();
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
 * Read-only diagnostic: logs server-side speed for startup and the major
 * dashboard load paths. Mutating workflow actions are intentionally not
 * executed.
 *
 * Optional second argument:
 *   { startupOnly: true, detailLimit: 8, customerSearchQueries: [''], calendarMonths: ['2026-05'] }
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
    note: 'Measures Apps Script server time only, including HtmlService generation. Browser/network rendering time is not included. Mutating workflow actions are skipped.',
    options: {
      startupOnly: options.startupOnly,
      includeStartup: options.includeStartup,
      bootstrapRepeats: options.bootstrapRepeats,
      includeReadModelStatus: options.includeReadModelStatus,
      includeLoginBootstrap: !!(options.loginEmail && options.loginPassword),
      detailLimit: options.detailLimit,
      includeTaskDetails: options.includeTaskDetails,
      includeCustomerDetail: options.includeCustomerDetail,
      customerSearchQueries: options.customerSearchQueries,
      calendarMonths: options.calendarMonths,
      adminDashboardFilters: options.adminDashboardFilters
    },
    steps: []
  };
  var bootstrap = null;
  var queueResponses = [];
  var customerSearchResponses = [];

  if (options.includeStartup) {
    swBenchmarkSalesWorkflowStep_(out, 'webApp:taskQueueHtml', function () {
      return swBenchmarkSalesWorkflowHtmlSummary_();
    });
  }

  if (options.includeReadModelStatus && typeof sw_getWorkflowReadModelStatus === 'function') {
    swBenchmarkSalesWorkflowStep_(out, 'sw_getWorkflowReadModelStatus', function () {
      return swBenchmarkSalesWorkflowReadModelSummary_(sw_getWorkflowReadModelStatus());
    });
  }

  if (!bootstrap) {
    swBenchmarkSalesWorkflowStep_(out, 'sw_getBootstrap', function () {
      bootstrap = sw_getBootstrap(authToken);
      return swBenchmarkSalesWorkflowBootstrapSummary_(bootstrap);
    });
  }

  if (options.includeStartup) {
    for (var repeat = 2; repeat <= options.bootstrapRepeats; repeat++) {
      swBenchmarkSalesWorkflowStep_(out, 'sw_getBootstrap:warm' + repeat, function () {
        return swBenchmarkSalesWorkflowBootstrapSummary_(sw_getBootstrap(authToken));
      });
    }

    if (options.loginEmail && options.loginPassword) {
      swBenchmarkSalesWorkflowStep_(out, 'sw_login+bootstrap', function () {
        return swBenchmarkSalesWorkflowLoginBootstrapSummary_(options.loginEmail, options.loginPassword);
      });
    } else {
      swBenchmarkSalesWorkflowSkip_(out, 'sw_login+bootstrap', 'Provide loginEmail and loginPassword in options to measure the password sign-in path.');
    }

    if (options.startupOnly) {
      return swBenchmarkSalesWorkflowFinalize_(out, started);
    }
  }

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
        appointments: res.appointmentCount || (res.appointments ? res.appointments.length : 0),
        source: res.source || '',
        readModelAgeSeconds: res.readModelAgeSeconds || 0
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

  return swBenchmarkSalesWorkflowFinalize_(out, started);
}

function sw_measureSalesWorkflowStartupSpeed(options) {
  options = options || {};
  options.startupOnly = true;
  options.includeStartup = true;
  return sw_measureSalesWorkflowSpeed('', options);
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

function swBenchmarkSalesWorkflowFinalize_(out, started) {
  out.totalMs = new Date().getTime() - started;
  out.summary = swBenchmarkSalesWorkflowSummary_(out.steps);
  out.summary.startup = swBenchmarkSalesWorkflowStartupSummary_(out.steps);
  out.ok = out.summary.failedSteps === 0;
  Logger.log('SW_BENCHMARK_SUMMARY ' + JSON.stringify(swBenchmarkSalesWorkflowLogSummary_(out)));
  return out;
}

function swBenchmarkSalesWorkflowHtmlSummary_() {
  var html = typeof sw_taskQueueDoGet_ === 'function'
    ? sw_taskQueueDoGet_({ parameter: {} })
    : HtmlService.createHtmlOutputFromFile('Index');
  var content = html && html.getContent ? html.getContent() : '';
  var buildMatch = /SW_DASHBOARD_BUILD\s*=\s*['"]([^'"]+)['"]/.exec(content);
  return {
    bytes: content.length,
    kb: Math.round(content.length / 1024),
    build: buildMatch ? buildMatch[1] : '',
    hasLoginScreen: content.indexOf('loginScreen') >= 0,
    hasAppShell: content.indexOf('appShell') >= 0
  };
}

function swBenchmarkSalesWorkflowBootstrapSummary_(bootstrap) {
  bootstrap = bootstrap || {};
  return {
    counts: bootstrap.counts || {},
    views: bootstrap.views || {},
    initialTasks: bootstrap.tasks ? bootstrap.tasks.length : 0,
    user: swBenchmarkSalesWorkflowUserSummary_(bootstrap.user)
  };
}

function swBenchmarkSalesWorkflowLoginBootstrapSummary_(email, password) {
  var res = sw_login(email, password, { includeBootstrap: true });
  var bootstrap = res.bootstrap || {};
  return {
    user: swBenchmarkSalesWorkflowUserSummary_(res.user),
    hasToken: !!res.token,
    expiresInSeconds: res.expiresInSeconds || 0,
    bootstrap: swBenchmarkSalesWorkflowBootstrapSummary_(bootstrap)
  };
}

function swBenchmarkSalesWorkflowUserSummary_(user) {
  user = user || {};
  return {
    email: user.email || '',
    name: user.name || '',
    roles: user.roles || [],
    isAdmin: !!user.isAdmin,
    isJoc: !!user.isJoc,
    isRep: !!user.isRep,
    isDiamondOrderAdmin: !!user.isDiamondOrderAdmin,
    isDiamondOrderAssistant: !!user.isDiamondOrderAssistant
  };
}

function swBenchmarkSalesWorkflowStartupSummary_(steps) {
  var html = swBenchmarkSalesWorkflowStepByOperation_(steps, 'webApp:taskQueueHtml');
  var bootstrap = swBenchmarkSalesWorkflowStepByOperation_(steps, 'sw_getBootstrap');
  var login = swBenchmarkSalesWorkflowStepByOperation_(steps, 'sw_login+bootstrap');
  var warm = (steps || []).filter(function (step) {
    return step && step.ok && !step.skipped && /^sw_getBootstrap:warm/.test(step.operation || '');
  }).map(function (step) {
    return step.ms || 0;
  }).filter(function (ms) {
    return ms > 0;
  });
  var htmlMs = html && html.ok && !html.skipped ? html.ms || 0 : 0;
  var bootstrapMs = bootstrap && bootstrap.ok && !bootstrap.skipped ? bootstrap.ms || 0 : 0;
  var loginMs = login && login.ok && !login.skipped ? login.ms || 0 : 0;
  return {
    htmlMs: htmlMs,
    htmlKb: html && html.result ? html.result.kb || 0 : 0,
    returningSessionBootstrapMs: bootstrapMs,
    returningSessionServerMs: htmlMs && bootstrapMs ? htmlMs + bootstrapMs : 0,
    warmBootstrapBestMs: warm.length ? Math.min.apply(null, warm) : 0,
    warmBootstrapWorstMs: warm.length ? Math.max.apply(null, warm) : 0,
    signInBootstrapMs: loginMs,
    firstSignInServerMs: htmlMs && loginMs ? htmlMs + loginMs : 0,
    signInMeasured: !!(login && login.ok && !login.skipped)
  };
}

function swBenchmarkSalesWorkflowStepByOperation_(steps, operation) {
  for (var i = 0; i < (steps || []).length; i++) {
    if (steps[i] && steps[i].operation === operation) return steps[i];
  }
  return null;
}

function swBenchmarkSalesWorkflowStep_(out, operation, fn) {
  var started = new Date().getTime();
  var timingCapture = typeof swTimingCaptureStart_ === 'function' ? swTimingCaptureStart_() : [];
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
    if (typeof swTimingCaptureStop_ === 'function') {
      step.timings = swTimingCaptureStop_(timingCapture);
    }
    step.ms = new Date().getTime() - started;
    out.steps.push(step);
    Logger.log('SW_BENCHMARK_STEP ' + JSON.stringify(swBenchmarkSalesWorkflowStepForLog_(step)));
  }
  return step;
}

function swBenchmarkSalesWorkflowStepForLog_(step) {
  return {
    operation: step.operation,
    ok: step.ok,
    skipped: !!step.skipped,
    reason: step.reason || '',
    error: step.error || '',
    ms: step.ms || 0,
    result: step.result || {}
  };
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
  var bootstrapRepeats = Number(options.bootstrapRepeats || options.startupRepeats || 2);
  if (!isFinite(bootstrapRepeats) || bootstrapRepeats < 1) bootstrapRepeats = 1;
  bootstrapRepeats = Math.min(Math.floor(bootstrapRepeats), 5);
  return {
    startupOnly: !!options.startupOnly,
    includeStartup: options.includeStartup !== false,
    bootstrapRepeats: bootstrapRepeats,
    includeReadModelStatus: options.includeReadModelStatus === true || String(options.includeReadModelStatus || '').toLowerCase() === 'true',
    loginEmail: swNormEmail_(options.loginEmail || options.email || ''),
    loginPassword: String(options.loginPassword || options.password || ''),
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
    tab: res && res.tab ? res.tab : '',
    source: res && res.source ? res.source : '',
    readModelAgeSeconds: res && res.readModelAgeSeconds ? res.readModelAgeSeconds : 0
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
    warnings: res && res.warnings ? res.warnings.length : 0,
    source: res && res.source ? res.source : '',
    readModelAgeSeconds: res && res.readModelAgeSeconds ? res.readModelAgeSeconds : 0
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
      ok: step.ok !== false,
      timings: swBenchmarkSalesWorkflowTimingSummary_(step.timings)
    });
  });
  summary.slowest.sort(function (a, b) { return b.ms - a.ms; });
  summary.slowest = summary.slowest.slice(0, 8);
  return summary;
}

function swBenchmarkSalesWorkflowTimingSummary_(timings) {
  var out = [];
  (timings || []).forEach(function (item) {
    if (!item || item.type !== 'step') return;
    out.push({
      operation: item.operation || '',
      step: item.step || '',
      ms: item.ms || 0,
      totalMs: item.totalMs || 0,
      extra: item.extra || {}
    });
  });
  if (out.length > 12) out = out.slice(0, 12);
  return out;
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
