/**
 * Additive infrastructure read models for expensive cross-workbook/dashboard reads.
 *
 * Source sheets remain writable sources of truth. These helpers build and serve
 * narrow hidden projections in the 100 workbook, with live-source fallbacks.
 */

var SW_DIAMOND_READ_MODEL_CACHE_SECONDS = 10 * 60;
var SW_CALENDAR_MONTH_READ_MODEL_CACHE_SECONDS = 10 * 60;
var SW_ADMIN_DASHBOARD_READ_MODEL_CACHE_SECONDS = 10 * 60;
var SW_DIAMOND_READ_MODEL_MEMORY_CACHE_ = {};
var SW_CALENDAR_MONTH_READ_MODEL_MEMORY_CACHE_ = {};
var SW_ADMIN_DASHBOARD_READ_MODEL_MEMORY_CACHE_ = {};

function swBuildDiamondReadModels_(ss, builtAt) {
  var started = new Date().getTime();
  var builtAtIso = swIso_(builtAt || new Date());
  try {
    var target = swDiamond200Target_();
    var records = [];
    var sourceName = '';
    if (target && target.sheet) {
      var sh = target.sheet;
      sourceName = target.tab || sh.getName();
      var sourceMeta = swDiamondReadModelSourceMeta_(target, sourceName);
      var lr = sh.getLastRow();
      var lc = sh.getLastColumn();
      if (lr >= 3 && lc >= 1) {
        var hm = swDiamond200HeaderMap_(sh);
        var C = swDiamondReadModelColumns_(hm);
        var values = swDiamondRead200Rows_(sh, 3, lr - 2, C, lc);
        values.forEach(function (row, i) {
          var rec = swDiamondReadModelRecordFromRow_(row, i + 3, C, sourceMeta);
          if (rec.root || rec.certNo || rec.vendor || rec.diamond) records.push(rec);
        });
      }
    }

    var rowValues = records.map(swDiamondReadModelRow_);
    var rowWrite = swWriteReadModelSheet_(ss, SW_SHEETS.READ_MODEL_DIAMONDS, SW_DIAMOND_READ_MODEL_HEADERS, rowValues);
    var rootRows = swDiamondRootReadModelRows_(records, builtAtIso);
    var rootWrite = swWriteReadModelSheet_(ss, SW_SHEETS.READ_MODEL_DIAMOND_ROOTS, SW_DIAMOND_ROOT_READ_MODEL_HEADERS, rootRows);
    var cacheResult = swCacheDiamondReadModelRecords_(ss, records, { builtAt: builtAtIso });
    return {
      ok: rowWrite.ok !== false && rootWrite.ok !== false,
      sheet: SW_SHEETS.READ_MODEL_DIAMONDS,
      sourceRows: records.length,
      outputRows: rowValues.length,
      rootRows: rootRows.length,
      buildMs: new Date().getTime() - started,
      rowBuildMs: rowWrite.buildMs || 0,
      rootBuildMs: rootWrite.buildMs || 0,
      sourceSheet: sourceName,
      cacheOk: cacheResult.ok === true,
      cacheChunks: cacheResult.chunks || 0,
      cacheBytes: cacheResult.bytes || 0,
      cacheError: cacheResult.reason || '',
      error: rowWrite.error || rootWrite.error || ''
    };
  } catch (err) {
    return swReadModelErrorResult_(err, started, SW_SHEETS.READ_MODEL_DIAMONDS);
  }
}

function swDiamondReadModelSourceMeta_(target, sourceName) {
  var meta = {
    sourceSpreadsheetUrl: '',
    sourceSpreadsheetName: '',
    sourceTab: sourceName || ''
  };
  try { meta.sourceSpreadsheetUrl = target && target.ss ? target.ss.getUrl() : ''; } catch (_) {}
  try { meta.sourceSpreadsheetName = target && target.ss ? target.ss.getName() : ''; } catch (_) {}
  return meta;
}

function swDiamondReadModelColumns_(hm) {
  var C = swDiamond200Columns_(hm);
  C.customerName = swDiamondFind200Column_(hm, ['Customer Name', 'Client Name', 'Customer']);
  C.appointment = swDiamondFind200Column_(hm, ['Customer Appt Time & Date', 'Customer Appointment Date', 'Appointment Date']);
  C.assignedRep = swDiamondFind200Column_(hm, ['Client Advisor', 'Assigned Rep', 'Sales Rep']);
  C.joc = swDiamondFind200Column_(hm, ['JOC', 'Assisted Rep', 'Assistant Rep']);
  C.company = swDiamondFind200Column_(hm, ['Company', 'Brand']);
  C.requestDate = swDiamondFind200Column_(hm, ['Request Date', 'Requested Date']);
  C.requestedBy = swDiamondFind200Column_(hm, ['Requested By', 'Request By']);
  C.orderedBy = swDiamondFind200Column_(hm, ['Ordered By', 'Purchased By']);
  C.memoDate = swDiamondFind200Column_(hm, ['Memo/ Invoice Date', 'Memo / Invoice Date', 'Memo Invoice Date', 'Memo/Invoice Date']);
  C.loupeOrder = swDiamondFind200Column_(hm, ['Loupe360 Order #', 'Loupe360 Order#', 'Vendor Order Number', 'Vendor Order #']);
  C.invoice = swDiamondFind200Column_(hm, ['Invoice Number', 'Invoice #']);
  C.syncAt = swDiamondFind200Column_(hm, ['Loupe360 Last Sync At', 'Last Loupe360 Sync At']);
  return C;
}

function swDiamondReadModelRecordFromRow_(row, rowIndex, C, sourceMeta) {
  var shape = swDiamondCell_(row, C.shape);
  var carat = swDiamondCell_(row, C.carat);
  var color = swDiamondCell_(row, C.color);
  var clarity = swDiamondCell_(row, C.clarity);
  var diamond = [shape, carat, color, clarity].filter(Boolean).join(' ');
  return {
    sourceRow: rowIndex,
    root: swDiamondCell_(row, C.root),
    customerName: swDiamondCell_(row, C.customerName),
    appointment: swDiamondCell_(row, C.appointment),
    assignedRep: swDiamondCell_(row, C.assignedRep),
    joc: swDiamondCell_(row, C.joc),
    company: swDiamondCell_(row, C.company),
    vendor: swDiamondCell_(row, C.vendor),
    stoneType: swDiamondCell_(row, C.stoneType),
    shape: shape,
    carat: carat,
    color: color,
    clarity: clarity,
    lab: swDiamondCell_(row, C.lab),
    certNo: swDiamondCell_(row, C.certNo),
    measurement: swDiamondCell_(row, C.measurement),
    ratio: swDiamondCell_(row, C.ratio),
    orderStatus: swDiamondCell_(row, C.orderStatus),
    stoneStatus: swDiamondCell_(row, C.stoneStatus),
    decision: swDiamondCell_(row, C.decision),
    requestDate: swDiamondCell_(row, C.requestDate),
    requestedBy: swDiamondCell_(row, C.requestedBy),
    orderedBy: swDiamondCell_(row, C.orderedBy),
    orderDate: swDiamondCell_(row, C.orderDate),
    memoDate: swDiamondCell_(row, C.memoDate),
    returnDueDate: swDiamondCell_(row, C.returnDueDate),
    trackingEta: swDiamondCell_(row, C.trackingEta),
    trackingStatus: swDiamondCell_(row, C.trackingStatus),
    carrier: swDiamondCell_(row, C.carrier),
    trackingNumber: swDiamondCell_(row, C.trackingNumber),
    trackingUrl: swDiamondCell_(row, C.trackingUrl),
    trackingNotes: swDiamondCell_(row, C.trackingNotes),
    loupeOrder: swDiamondCell_(row, C.loupeOrder),
    invoice: swDiamondCell_(row, C.invoice),
    syncAt: swDiamondCell_(row, C.syncAt),
    diamond: diamond,
    sourceSpreadsheetUrl: sourceMeta && sourceMeta.sourceSpreadsheetUrl || '',
    sourceSpreadsheetName: sourceMeta && sourceMeta.sourceSpreadsheetName || '',
    sourceTab: sourceMeta && sourceMeta.sourceTab || ''
  };
}

function swDiamondReadModelRow_(rec) {
  var values = [
    rec.sourceRow || '',
    rec.root || '',
    rec.customerName || '',
    rec.appointment || '',
    rec.assignedRep || '',
    rec.joc || '',
    rec.company || '',
    rec.vendor || '',
    rec.stoneType || '',
    rec.shape || '',
    rec.carat || '',
    rec.color || '',
    rec.clarity || '',
    rec.lab || '',
    rec.certNo || '',
    rec.measurement || '',
    rec.ratio || '',
    rec.orderStatus || '',
    rec.stoneStatus || '',
    rec.decision || '',
    rec.requestDate || '',
    rec.requestedBy || '',
    rec.orderedBy || '',
    rec.orderDate || '',
    rec.memoDate || '',
    rec.returnDueDate || '',
    rec.trackingEta || '',
    rec.trackingStatus || '',
    rec.carrier || '',
    rec.trackingNumber || '',
    rec.trackingUrl || '',
    rec.trackingNotes || '',
    rec.loupeOrder || '',
    rec.invoice || '',
    rec.syncAt || '',
    rec.diamond || '',
    rec.sourceSpreadsheetUrl || '',
    rec.sourceSpreadsheetName || '',
    rec.sourceTab || ''
  ];
  values.push(swReadModelSearchText_(values));
  return values;
}

function swDiamondRootReadModelRows_(records, builtAtIso) {
  var byRoot = {};
  (records || []).forEach(function (rec) {
    var root = swTrim_(rec.root);
    if (!root) return;
    var item = byRoot[root];
    if (!item) {
      item = {
        root: root,
        customerName: rec.customerName || '',
        assignedRep: rec.assignedRep || '',
        joc: rec.joc || '',
        company: rec.company || '',
        stoneCount: 0,
        proposing: 0,
        onTheWay: 0,
        delivered: 0,
        inStock: 0,
        returns: 0,
        purchases: 0,
        issues: 0,
        missingAssignment: 0,
        earliestReturnMs: 0,
        earliestReturnDue: '',
        latestEtaMs: 0,
        latestEta: '',
        sourceRows: []
      };
      byRoot[root] = item;
    }
    item.stoneCount++;
    item.sourceRows.push(rec.sourceRow);
    if (!item.customerName && rec.customerName) item.customerName = rec.customerName;
    if (!item.assignedRep && rec.assignedRep) item.assignedRep = rec.assignedRep;
    if (!item.joc && rec.joc) item.joc = rec.joc;
    if (!item.company && rec.company) item.company = rec.company;
    var order = swNorm_(rec.orderStatus);
    var stone = swNorm_(rec.stoneStatus);
    var decision = swNorm_(rec.decision);
    if (order === 'proposing') item.proposing++;
    if (order === 'on the way') item.onTheWay++;
    if (order === 'delivered') item.delivered++;
    if (stone.indexOf('in stock') >= 0) item.inStock++;
    if (decision === 'return' || stone.indexOf('return in progress') >= 0) item.returns++;
    if (decision === 'purchase' || decision === 'purchased') item.purchases++;
    if (!rec.customerName || !rec.assignedRep || !rec.joc) item.missingAssignment++;
    if (order === 'on the way' && !rec.trackingEta) item.issues++;
    var dueMs = swDiamondDateValue_(rec.returnDueDate);
    if (dueMs && dueMs < 9999999999999 && (!item.earliestReturnMs || dueMs < item.earliestReturnMs)) {
      item.earliestReturnMs = dueMs;
      item.earliestReturnDue = rec.returnDueDate;
    }
    var etaMs = swDiamondDateValue_(rec.trackingEta);
    if (etaMs && etaMs < 9999999999999 && etaMs > item.latestEtaMs) {
      item.latestEtaMs = etaMs;
      item.latestEta = rec.trackingEta;
    }
  });
  return Object.keys(byRoot).sort().map(function (root) {
    var item = byRoot[root];
    var summary = swDiamondCountsSummary_({
      proposing: item.proposing,
      onTheWay: item.onTheWay,
      delivered: item.delivered,
      total: item.stoneCount
    });
    var values = [
      item.root,
      item.customerName,
      item.assignedRep,
      item.joc,
      item.company,
      item.stoneCount,
      item.proposing,
      item.onTheWay,
      item.delivered,
      item.inStock,
      item.returns,
      item.purchases,
      item.issues,
      item.missingAssignment,
      item.earliestReturnDue,
      item.latestEta,
      summary,
      swStringify_(item.sourceRows),
      builtAtIso
    ];
    values.push(swReadModelSearchText_(values));
    return values;
  });
}

function swBuildAppointmentReadModels_(ss, builtAt) {
  var started = new Date().getTime();
  var builtAtIso = swIso_(builtAt || new Date());
  try {
    var appointments = swReadAppointments_(ss);
    var rows = appointments.map(swAppointmentReadModelRow_);
    var appointmentWrite = swWriteReadModelSheet_(ss, SW_SHEETS.READ_MODEL_APPOINTMENTS, SW_APPOINTMENT_READ_MODEL_HEADERS, rows);
    var calendarRows = swCalendarMonthReadModelRows_(ss, appointments, builtAtIso);
    var calendarWrite = swWriteReadModelSheet_(ss, SW_SHEETS.READ_MODEL_CALENDAR_MONTHS, SW_CALENDAR_MONTH_READ_MODEL_HEADERS, calendarRows);
    var calendarCache = swCacheCalendarMonthReadModels_(ss, calendarRows, { builtAt: builtAtIso });
    return {
      ok: appointmentWrite.ok !== false && calendarWrite.ok !== false,
      sheet: SW_SHEETS.READ_MODEL_APPOINTMENTS,
      sourceRows: appointments.length,
      outputRows: rows.length,
      calendarMonths: calendarRows.length,
      buildMs: new Date().getTime() - started,
      appointmentBuildMs: appointmentWrite.buildMs || 0,
      calendarBuildMs: calendarWrite.buildMs || 0,
      cacheOk: calendarCache.ok === true,
      cacheChunks: calendarCache.chunks || 0,
      cacheBytes: calendarCache.bytes || 0,
      cacheError: calendarCache.reason || '',
      error: appointmentWrite.error || calendarWrite.error || ''
    };
  } catch (err) {
    return swReadModelErrorResult_(err, started, SW_SHEETS.READ_MODEL_APPOINTMENTS);
  }
}

function swAppointmentReadModelRow_(rec) {
  rec = rec || {};
  var visitAt = swVisitDateTime_(rec, swTimezone_());
  var values = [
    rec.row || '',
    rec.appt || '',
    rec.root || '',
    rec.uid || '',
    rec.name || '',
    rec.email || '',
    rec.phone || '',
    rec.brand || '',
    rec.bookedAt || '',
    rec.canceledAt || '',
    rec.rescheduledFromUid || '',
    rec.rescheduledToUid || '',
    rec.visitDate || '',
    rec.visitTime || '',
    visitAt ? visitAt.getTime() : '',
    rec.visitType || '',
    rec.diamondType || '',
    rec.status || '',
    rec.active || '',
    rec.assignedRep || '',
    rec.assignedRepEmail || '',
    rec.assistedRep || '',
    rec.assistedRepEmail || '',
    rec.clientFolder || '',
    rec.reportUrl || '',
    rec.quotationUrl || '',
    rec.tracker3dUrl || '',
    rec.salesStage || '',
    rec.convStatus || '',
    rec.customOrder || '',
    rec.inProduction || '',
    rec.nextSteps || '',
    rec.designRequest || '',
    rec.deadline3d || '',
    rec.productionDeadline || '',
    rec.waxStatus || '',
    rec.waxDeadlineAdmin || '',
    rec.waxRequestUrl || '',
    rec.centerStoneStatus || '',
    rec.dvStonesSummary || '',
    rec.dvCustomerLookingFor || '',
    rec.dvVarietyStrategy || '',
    rec.dvCustomerRequirementsJson || '',
    rec.so || '',
    rec.orderFolder || '',
    rec.source || '',
    rec.budgetMin || '',
    rec.budgetMax || '',
    rec.orderTotal || '',
    rec.paidToDate || '',
    rec.remainingBalance || '',
    rec.lastPaymentDate || '',
    rec.orderDate || '',
    rec.updatedAt || '',
    rec.deadline3dMoves || '',
    rec.productionDeadlineMoves || ''
  ];
  values.push(swReadModelSearchText_(values));
  return values;
}

function swAppointmentFromReadModelRow_(row) {
  row = row || {};
  var rec = {
    row: Number(row['Row Number'] || 0) || 0,
    appt: row['APPT_ID'] || '',
    root: row['RootApptID'] || row['APPT_ID'] || '',
    uid: row['UID'] || '',
    name: row['Customer Name'] || '',
    email: row['Email'] || '',
    phone: row['Phone'] || '',
    brand: row['Brand'] || '',
    bookedAt: row['Booked At'] || '',
    bookedAtRaw: row['Booked At'] || '',
    canceledAt: row['Canceled At'] || '',
    canceledAtRaw: row['Canceled At'] || '',
    rescheduledFromUid: row['Rescheduled From UID'] || '',
    rescheduledToUid: row['Rescheduled To UID'] || '',
    visitDate: row['Visit Date'] || '',
    visitTime: row['Visit Time'] || '',
    visitDateRaw: row['Visit Date'] || '',
    visitTimeRaw: row['Visit Time'] || '',
    visitAtMs: Number(row['Visit At Ms'] || 0) || 0,
    visitType: row['Visit Type'] || '',
    diamondType: row['Diamond Type'] || '',
    status: row['Status'] || '',
    active: row['Active?'] || '',
    assignedRep: row['Client Advisor'] || '',
    assignedRepEmail: row['Client Advisor Email'] || '',
    assistedRep: row['JOC'] || '',
    assistedRepEmail: row['JOC Email'] || '',
    clientFolder: row['Client Folder'] || '',
    reportUrl: row['Client Status Report URL'] || '',
    quotationUrl: row['Quotation URL'] || '',
    tracker3dUrl: row['3D Tracker URL'] || '',
    salesStage: row['Sales Stage'] || '',
    convStatus: row['Conversion Status'] || '',
    customOrder: row['Custom Order Status'] || '',
    inProduction: row['In Production Status'] || '',
    nextSteps: row['Next Steps'] || '',
    designRequest: row['Design Request'] || '',
    deadline3d: row['3D Deadline'] || '',
    productionDeadline: row['Production Deadline'] || '',
    waxStatus: row['Wax Print Status'] || '',
    waxDeadlineAdmin: row['Wax Deadline (Admin)'] || '',
    waxRequestUrl: row['Wax Request URL'] || '',
    centerStoneStatus: row['Center Stone Status'] || '',
    dvStonesSummary: row['DV Stones Summary'] || '',
    dvCustomerLookingFor: row['DV Customer Looking For'] || '',
    dvVarietyStrategy: row['DV Variety Strategy'] || '',
    dvCustomerRequirementsJson: row['DV Customer Requirements (JSON)'] || '',
    so: row['SO#'] || '',
    orderFolder: row['Order Folder'] || '',
    source: row['Source'] || '',
    budgetMin: row['Budget Min'] || '',
    budgetMax: row['Budget Max'] || '',
    orderTotal: row['Order Total'] || '',
    paidToDate: row['Paid-to-Date'] || '',
    remainingBalance: row['Remaining Balance'] || '',
    lastPaymentDate: row['Last Payment Date'] || '',
    lastPaymentDateRaw: row['Last Payment Date'] || '',
    orderDate: row['Order Date'] || '',
    orderDateRaw: row['Order Date'] || '',
    updatedAt: row['Updated At'] || '',
    updatedAtRaw: row['Updated At'] || '',
    deadline3dMoves: row['3D Deadline Moves'] || '',
    productionDeadlineMoves: row['Production Deadline Moves'] || ''
  };
  rec.statusNorm = swNorm_(rec.status);
  rec.activeNorm = swNorm_(rec.active);
  return rec;
}

function swCalendarMonthReadModelRows_(ss, appointments, builtAtIso) {
  var byMonth = {};
  var tz = swTimezone_();
  var aiBriefByRoot = {};
  try {
    aiBriefByRoot = typeof swAppointmentAiBriefIndex_ === 'function'
      ? swAppointmentAiBriefIndex_(ss)
      : {};
  } catch (_) {}
  (appointments || []).forEach(function (rec) {
    if (!swIsAppointmentActive_(rec)) return;
    var visitAt = swVisitDateTime_(rec, tz);
    if (!visitAt) return;
    var monthKey = swCalendarMonthKey_(new Date(visitAt.getFullYear(), visitAt.getMonth(), 1));
    if (!byMonth[monthKey]) byMonth[monthKey] = [];
    byMonth[monthKey].push(swCalendarAppointmentPayloadFromRec_(ss, rec, visitAt, aiBriefByRoot));
  });
  return Object.keys(byMonth).sort().map(function (monthKey) {
    var month = swCalendarMonthRange_(monthKey);
    var appointmentsForMonth = byMonth[monthKey].sort(function (a, b) {
      return String(a.sortAt).localeCompare(String(b.sortAt)) || String(a.customerName).localeCompare(String(b.customerName));
    });
    var json = swStringify_(appointmentsForMonth);
    var values = [
      monthKey,
      Utilities.formatDate(month.start, tz, 'MMMM yyyy'),
      appointmentsForMonth.length,
      json,
      builtAtIso
    ];
    values.push(swReadModelSearchText_(values.slice(0, 3)));
    return values;
  });
}

function swCalendarAppointmentPayloadFromRec_(ss, rec, visitAt, aiBriefByRoot) {
  var root = rec.root || rec.appt || '';
  var aiBrief = root && aiBriefByRoot && aiBriefByRoot[root] && typeof swAppointmentAiBriefCompact_ === 'function'
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
}

function swBuildPaymentReadModel_(ss, builtAt) {
  var started = new Date().getTime();
  try {
    var warnings = [];
    var source = typeof swAdminDashboardReadPaymentReceiptRows_ === 'function'
      ? swAdminDashboardReadPaymentReceiptRows_(warnings, { forceLive: true })
      : { rows: [] };
    var rows = (source.rows || []).map(swPaymentReadModelRow_);
    var write = swWriteReadModelSheet_(ss, SW_SHEETS.READ_MODEL_PAYMENTS, SW_PAYMENT_READ_MODEL_HEADERS, rows);
    write.sourceRows = source.rows ? source.rows.length : 0;
    write.outputRows = rows.length;
    write.buildMs = new Date().getTime() - started;
    write.warnings = warnings.length;
    write.notes = warnings.join(' | ').slice(0, 500);
    return write;
  } catch (err) {
    return swReadModelErrorResult_(err, started, SW_SHEETS.READ_MODEL_PAYMENTS);
  }
}

function swPaymentReadModelRow_(row) {
  row = row || {};
  var values = [
    row.sourceRow || '',
    row.root || '',
    row.so || '',
    row.key || row.root || row.so || '',
    row.brand || '',
    row.paymentId || '',
    row.docType || '',
    row.docNumber || '',
    row.docStatus || '',
    row.method || '',
    row.whenMs ? swIso_(new Date(Number(row.whenMs))) : '',
    row.whenMs || '',
    row.net || 0,
    row.gross || 0,
    row.balance === '' || row.balance == null ? '' : row.balance,
    row.orderTotal === '' || row.orderTotal == null ? '' : row.orderTotal
  ];
  values.push(swReadModelSearchText_(values));
  return values;
}

function swBuildAdminDashboardReadModel_(ss, builtAt) {
  var started = new Date().getTime();
  try {
    if (typeof swAdminDashboardBuildPayload_ !== 'function') {
      var emptyWrite = swWriteReadModelSheet_(ss, SW_SHEETS.READ_MODEL_ADMIN_DASHBOARD, SW_ADMIN_DASHBOARD_READ_MODEL_HEADERS, []);
      emptyWrite.sourceRows = 0;
      emptyWrite.outputRows = 0;
      emptyWrite.buildMs = new Date().getTime() - started;
      emptyWrite.notes = 'adminDashboardBuildPayloadUnavailable';
      return emptyWrite;
    }
    var builtAtIso = swIso_(builtAt || new Date());
    var presets = ['today', 'last7', 'last30', 'thisMonth'];
    var rows = [];
    var oversized = 0;
    presets.forEach(function (preset) {
      var filters = swAdminDashboardNormalizeFilters_({ windowPreset: preset });
      var payload = swAdminDashboardBuildPayload_(ss, filters, { source: 'adminDashboardReadModelBuild' });
      payload.source = 'adminDashboardReadModel';
      payload.readModelBuiltAt = builtAtIso;
      var text = swStringify_(payload);
      if (text.length > 48000) {
        oversized++;
        return;
      }
      var values = [
        swAdminDashboardReadModelKey_(filters),
        filters.windowPreset,
        filters.brand || '',
        filters.clientAdvisor || '',
        filters.joc || '',
        filters.includeClosed ? 'Y' : 'N',
        builtAtIso,
        text,
        text.length
      ];
      values.push(swReadModelSearchText_(values.slice(0, 7)));
      rows.push(values);
    });
    var write = swWriteReadModelSheet_(ss, SW_SHEETS.READ_MODEL_ADMIN_DASHBOARD, SW_ADMIN_DASHBOARD_READ_MODEL_HEADERS, rows);
    var adminCache = swCacheAdminDashboardReadModels_(ss, rows, { builtAt: builtAtIso });
    write.sourceRows = presets.length;
    write.outputRows = rows.length;
    write.buildMs = new Date().getTime() - started;
    write.oversizedPayloads = oversized;
    write.cacheOk = adminCache.ok === true;
    write.cacheChunks = adminCache.chunks || 0;
    write.cacheBytes = adminCache.bytes || 0;
    write.cacheError = adminCache.reason || '';
    if (oversized) write.notes = oversized + ' oversized admin dashboard payload(s) skipped.';
    return write;
  } catch (err) {
    return swReadModelErrorResult_(err, started, SW_SHEETS.READ_MODEL_ADMIN_DASHBOARD);
  }
}

function swTryGetCalendarAppointmentsFromReadModel_(ss, monthKey) {
  var config = swReadConfig_(ss, true);
  if (!swReadModelServingFlag_(config, 'READ_MODEL_SERVE_APPOINTMENTS', 'Y')) return null;
  var status = swReadModelFreshStatus_(ss, 'calendarMonths', SW_SHEETS.READ_MODEL_CALENDAR_MONTHS);
  if (!status.fresh) return null;
  var month = swCalendarMonthRange_(monthKey);
  var cached = swCachedCalendarMonthReadModels_(ss, status);
  var found = cached && cached.months ? cached.months[month.key] : null;
  if (!found) {
    var sh = ss.getSheetByName(SW_SHEETS.READ_MODEL_CALENDAR_MONTHS);
    var rows = swReadSheetObjectsExpectedHeaders_(sh, SW_CALENDAR_MONTH_READ_MODEL_HEADERS);
    for (var i = 0; i < rows.length; i++) {
      if (swTrim_(rows[i]['Month Key']) === month.key) {
        found = {
          monthKey: rows[i]['Month Key'] || '',
          monthLabel: rows[i]['Month Label'] || '',
          appointmentCount: rows[i]['Appointment Count'] || '',
          appointments: swParseJson_(rows[i]['Appointments JSON'] || '[]', [])
        };
        break;
      }
    }
    try { swCacheCalendarMonthReadModels_(ss, rows, status); } catch (_) {}
  }
  if (!found) return null;
  var appointments = found.appointments || [];
  if (!Array.isArray(appointments)) appointments = [];
  var today = new Date();
  var todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0, 0);
  appointments = appointments.filter(function (appt) {
    var sortAt = new Date(appt.sortAt || '');
    return !isNaN(sortAt.getTime()) && sortAt.getTime() >= todayStart.getTime();
  }).sort(function (a, b) {
    return String(a.sortAt).localeCompare(String(b.sortAt)) || String(a.customerName).localeCompare(String(b.customerName));
  });
  return {
    ok: true,
    source: 'calendarMonthReadModel',
    readModelAgeSeconds: status.ageSeconds || 0,
    monthKey: month.key,
    monthLabel: found.monthLabel || Utilities.formatDate(month.start, swTimezone_(), 'MMMM yyyy'),
    prevMonthKey: swCalendarMonthKey_(new Date(month.start.getFullYear(), month.start.getMonth() - 1, 1)),
    nextMonthKey: swCalendarMonthKey_(new Date(month.start.getFullYear(), month.start.getMonth() + 1, 1)),
    todayKey: swDateKey_(todayStart),
    appointmentCount: appointments.length,
    appointments: appointments
  };
}

function swTryGetInStockDiamondsFromReadModel_(ss, config) {
  var records = swReadDiamondReadModelRecords_(ss, config);
  if (!records) return null;
  var returnWindow = Number(swConfigValue_(config || [], 'SYSTEM', 'DIAMOND_RETURN_WINDOW_DAYS', '30')) || 30;
  var returnWarning = Number(swConfigValue_(config || [], 'SYSTEM', 'DIAMOND_RETURN_WARNING_DAYS', '7')) || 7;
  var today = new Date();
  var todayMs = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
  var warningMs = todayMs + returnWarning * 24 * 60 * 60 * 1000;
  var rows = [];
  var stats = { total: 0, available: 0, returnSoon: 0, returnOverdue: 0, noReturnDate: 0, assignmentMissing: 0, warningDays: returnWarning };
  records.rows.forEach(function (rec) {
    var orderNorm = swNorm_(rec.orderStatus);
    var stoneNorm = swNorm_(rec.stoneStatus);
    var decisionNorm = swNorm_(rec.decision);
    var isInStock = stoneNorm.indexOf('in stock') >= 0 || orderNorm === 'delivered';
    var unavailable = /return in progress|returned|sold|customer purchased/.test(stoneNorm) ||
      decisionNorm === 'purchase' || decisionNorm === 'purchased';
    if (!isInStock || unavailable) return;
    var returnDueDate = rec.returnDueDate || swDiamondReturnDueDate_(rec.orderDate, returnWindow);
    var returnMs = swDiamondDateValue_(returnDueDate);
    var issue = '';
    if (!returnDueDate) issue = 'No return date';
    else if (returnMs < todayMs) issue = 'Return overdue';
    else if (returnMs <= warningMs) issue = 'Return soon';
    var daysUntilReturn = returnMs ? Math.ceil((returnMs - todayMs) / (24 * 60 * 60 * 1000)) : '';
    var assignmentMissing = !rec.root || !rec.customerName || !rec.assignedRep || !rec.joc;
    rows.push({
      rowIndex: Number(rec.sourceRow || 0),
      root: rec.root,
      customerName: rec.customerName,
      appointment: rec.appointment,
      assignedRep: rec.assignedRep,
      joc: rec.joc,
      company: rec.company,
      vendor: rec.vendor,
      stoneType: rec.stoneType,
      certNo: rec.certNo,
      shape: rec.shape,
      carat: rec.carat,
      color: rec.color,
      clarity: rec.clarity,
      diamond: rec.diamond,
      measurement: rec.measurement,
      ratio: rec.ratio,
      lab: rec.lab,
      orderStatus: rec.orderStatus,
      stoneStatus: rec.stoneStatus,
      decision: rec.decision,
      orderDate: rec.orderDate,
      memoDate: rec.memoDate,
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
  rows.sort(swDiamondReturnRowSort_);
  return swDiamondReadModelResult_('diamondReadModel', records, stats, rows.slice(0, 300), []);
}

function swTryGetBulkReturnCandidatesFromReadModel_(ss, config) {
  var result = swTryGetInStockDiamondsFromReadModel_(ss, config);
  if (!result) return null;
  result.source = 'diamondReadModel';
  result.rows = (result.rows || []).slice(0, 500);
  return result;
}

function swTryGetDiamondTrackingDashboardFromReadModel_(ss, config) {
  var records = swReadDiamondReadModelRecords_(ss, config || []);
  if (!records) return null;
  var today = new Date();
  var todayMs = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
  var warningMs = todayMs + 7 * 24 * 60 * 60 * 1000;
  var rows = [];
  var stats = { total: 0, onTheWay: 0, delivered: 0, returns: 0, missingEta: 0, issues: 0 };
  records.rows.forEach(function (rec) {
    var orderNorm = swNorm_(rec.orderStatus);
    var stoneNorm = swNorm_(rec.stoneStatus);
    var decisionNorm = swNorm_(rec.decision);
    var trackingNorm = swNorm_(rec.trackingStatus);
    var relevant = orderNorm === 'on the way' || orderNorm === 'delivered' ||
      decisionNorm === 'return' || rec.trackingEta || rec.trackingStatus ||
      stoneNorm.indexOf('return in progress') >= 0;
    if (!relevant) return;
    var etaMs = swDiamondDateValue_(rec.trackingEta);
    var returnMs = swDiamondDateValue_(rec.returnDueDate);
    var issue = '';
    if (orderNorm === 'on the way' && !rec.trackingEta) issue = 'Missing ETA';
    if (!issue && /(delay|unavailable|cancel|concern|problem)/.test(trackingNorm)) issue = 'Tracking concern';
    if (!issue && etaMs && etaMs < todayMs && orderNorm === 'on the way') issue = 'ETA overdue';
    if (!issue && decisionNorm === 'return' && returnMs && returnMs <= warningMs) issue = returnMs < todayMs ? 'Return overdue' : 'Return due soon';
    rows.push({
      rowIndex: Number(rec.sourceRow || 0),
      root: rec.root,
      customerName: rec.customerName,
      appointment: rec.appointment,
      assignedRep: rec.assignedRep,
      vendor: rec.vendor,
      certNo: rec.certNo,
      diamond: rec.diamond,
      orderStatus: rec.orderStatus,
      stoneStatus: rec.stoneStatus,
      decision: rec.decision,
      orderDate: rec.orderDate,
      returnDueDate: rec.returnDueDate,
      trackingEta: rec.trackingEta,
      trackingStatus: rec.trackingStatus,
      carrier: rec.carrier,
      trackingNumber: rec.trackingNumber,
      trackingUrl: rec.trackingUrl,
      issue: issue
    });
    stats.total++;
    if (orderNorm === 'on the way') stats.onTheWay++;
    if (orderNorm === 'delivered') stats.delivered++;
    if (decisionNorm === 'return' || stoneNorm.indexOf('return in progress') >= 0) stats.returns++;
    if (orderNorm === 'on the way' && !rec.trackingEta) stats.missingEta++;
    if (issue) stats.issues++;
  });
  rows.sort(function (a, b) {
    if (!!a.issue !== !!b.issue) return a.issue ? -1 : 1;
    var av = swDiamondDateValue_(a.trackingEta || a.returnDueDate) || 9999999999999;
    var bv = swDiamondDateValue_(b.trackingEta || b.returnDueDate) || 9999999999999;
    return av - bv;
  });
  return swDiamondReadModelResult_('diamondReadModel', records, stats, rows.slice(0, 200), []);
}

function swDiamondReturnRowSort_(a, b) {
  if (!!a.issue !== !!b.issue) return a.issue ? -1 : 1;
  var av = swDiamondDateValue_(a.returnDueDate) || 9999999999999;
  var bv = swDiamondDateValue_(b.returnDueDate) || 9999999999999;
  return av - bv || String(a.diamond).localeCompare(String(b.diamond));
}

function swDiamondReadModelResult_(source, records, stats, rows, missingColumns) {
  var first = records.rows && records.rows.length ? records.rows[0] : {};
  return {
    ok: true,
    available: true,
    source: source,
    readModelAgeSeconds: records.ageSeconds || 0,
    generatedAt: swIso_(new Date()),
    spreadsheetUrl: first.sourceSpreadsheetUrl || '',
    spreadsheetName: first.sourceSpreadsheetName || '',
    tab: first.sourceTab || '',
    missingColumns: missingColumns || [],
    stats: stats || {},
    rows: rows || []
  };
}

function swReadDiamondReadModelRecords_(ss, config) {
  if (!swReadModelServingFlag_(config || [], 'READ_MODEL_SERVE_DIAMONDS', 'Y')) return null;
  var status = swReadModelFreshStatus_(ss, 'diamonds', SW_SHEETS.READ_MODEL_DIAMONDS);
  if (!status.fresh) return null;
  var cached = swCachedDiamondReadModelRecords_(ss, status);
  if (cached) return cached;
  var sh = ss.getSheetByName(SW_SHEETS.READ_MODEL_DIAMONDS);
  var rows = swReadSheetObjectsExpectedHeaders_(sh, SW_DIAMOND_READ_MODEL_HEADERS).map(swDiamondReadModelRecordFromObject_);
  try { swCacheDiamondReadModelRecords_(ss, rows, status); } catch (_) {}
  return { rows: rows, ageSeconds: status.ageSeconds || 0 };
}

function swDiamondReadModelRecordFromObject_(row) {
  row = row || {};
  return {
    sourceRow: row['Source Row'] || '',
    root: row['RootApptID'] || '',
    customerName: row['Customer Name'] || '',
    appointment: row['Customer Appt Time & Date'] || '',
    assignedRep: row['Client Advisor'] || '',
    joc: row['JOC'] || '',
    company: row['Company'] || '',
    vendor: row['Vendor'] || '',
    stoneType: row['Stone Type'] || '',
    shape: row['Shape'] || '',
    carat: row['Carat'] || '',
    color: row['Color'] || '',
    clarity: row['Clarity'] || '',
    lab: row['LAB'] || '',
    certNo: row['Certificate No'] || '',
    measurement: row['Measurements'] || '',
    ratio: row['L/W Ratio'] || '',
    orderStatus: row['Order Status'] || '',
    stoneStatus: row['Stone Status'] || '',
    decision: row['Stone Decision'] || '',
    requestDate: row['Request Date'] || '',
    requestedBy: row['Requested By'] || '',
    orderedBy: row['Ordered By'] || '',
    orderDate: row['Purchased / Ordered Date'] || '',
    memoDate: row['Memo/ Invoice Date'] || '',
    returnDueDate: row['Return DUE DATE'] || '',
    trackingEta: row['Tracking ETA'] || '',
    trackingStatus: row['Tracking Status'] || '',
    carrier: row['Carrier'] || '',
    trackingNumber: row['Tracking Number'] || '',
    trackingUrl: row['Tracking URL'] || '',
    trackingNotes: row['Tracking Notes'] || '',
    loupeOrder: row['Loupe360 Order #'] || '',
    invoice: row['Invoice Number'] || '',
    syncAt: row['Loupe360 Last Sync At'] || '',
    diamond: row['Diamond Label'] || '',
    sourceSpreadsheetUrl: row['Source Spreadsheet URL'] || '',
    sourceSpreadsheetName: row['Source Spreadsheet Name'] || '',
    sourceTab: row['Source Tab'] || ''
  };
}

function swReadPaymentReceiptRowsFromReadModel_(ss, warnings) {
  var config = [];
  try { config = swReadConfig_(ss, true); } catch (_) {}
  if (!swReadModelServingFlag_(config, 'READ_MODEL_SERVE_PAYMENTS', 'Y')) return null;
  var status = swReadModelFreshStatus_(ss, 'payments', SW_SHEETS.READ_MODEL_PAYMENTS);
  if (!status.fresh) return null;
  var sh = ss.getSheetByName(SW_SHEETS.READ_MODEL_PAYMENTS);
  var rows = swReadSheetObjectsExpectedHeaders_(sh, SW_PAYMENT_READ_MODEL_HEADERS).map(function (row) {
    return {
      sourceRow: Number(row['Source Row'] || 0) || 0,
      root: swAdminDashboardCleanId_(row['RootApptID']),
      so: swAdminDashboardCleanId_(row['SO#']),
      key: row['Key'] || row['RootApptID'] || row['SO#'] || '',
      brand: row['Brand'] || '',
      paymentId: row['Payment ID'] || '',
      docType: row['Doc Type'] || '',
      docNumber: row['Doc Number'] || '',
      docStatus: row['Doc Status'] || '',
      method: row['Method'] || '',
      whenMs: Number(row['Payment At Ms'] || 0) || 0,
      net: swAdminDashboardNumber_(row['Amount Net']),
      gross: swAdminDashboardNumber_(row['Amount Gross']),
      balance: row['Balance Due'] === '' ? '' : swAdminDashboardNumberOrBlank_(row['Balance Due']),
      orderTotal: row['Order Total'] === '' ? '' : swAdminDashboardNumberOrBlank_(row['Order Total'])
    };
  });
  return { rows: rows, cacheHit: true, source: 'paymentReadModel', ageSeconds: status.ageSeconds || 0 };
}

function swTryReadAdminDashboardFromReadModel_(ss, filters) {
  var config = [];
  try { config = swReadConfig_(ss, true); } catch (_) {}
  if (!swReadModelServingFlag_(config, 'READ_MODEL_SERVE_ADMIN', 'Y')) return null;
  if (!swAdminDashboardReadModelEligible_(filters)) return null;
  var status = swReadModelFreshStatus_(ss, 'adminDashboard', SW_SHEETS.READ_MODEL_ADMIN_DASHBOARD);
  if (!status.fresh) return null;
  var key = swAdminDashboardReadModelKey_(filters);
  var cached = swCachedAdminDashboardReadModels_(ss, status);
  if (cached && cached.payloads && cached.payloads[key]) {
    var cachedPayload = cached.payloads[key];
    if (!cachedPayload || cachedPayload.ok === false) return null;
    cachedPayload.source = 'adminDashboardReadModel';
    cachedPayload.readModelAgeSeconds = status.ageSeconds || 0;
    return cachedPayload;
  }
  var sh = ss.getSheetByName(SW_SHEETS.READ_MODEL_ADMIN_DASHBOARD);
  var rows = swReadSheetObjectsExpectedHeaders_(sh, SW_ADMIN_DASHBOARD_READ_MODEL_HEADERS);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i]['Key'] !== key) continue;
    var payload = swParseJson_(rows[i]['Payload JSON'] || '', null);
    if (!payload || payload.ok === false) return null;
    payload.source = 'adminDashboardReadModel';
    payload.readModelAgeSeconds = status.ageSeconds || 0;
    try { swCacheAdminDashboardReadModels_(ss, rows, status); } catch (_) {}
    return payload;
  }
  return null;
}

function swAdminDashboardReadModelEligible_(filters) {
  filters = filters || {};
  if (filters.brand || filters.clientAdvisor || filters.joc || filters.includeClosed) return false;
  return ['today', 'last7', 'last30', 'thisMonth'].indexOf(filters.windowPreset || '') >= 0;
}

function swAdminDashboardReadModelKey_(filters) {
  filters = filters || {};
  return [
    filters.windowPreset || 'last7',
    swNorm_(filters.brand || ''),
    swNorm_(filters.clientAdvisor || ''),
    swNorm_(filters.joc || ''),
    filters.includeClosed ? 'closed' : 'open'
  ].join('|');
}

function swDiamondReadModelCacheKey_(ss) {
  return 'sw:diamondReadModel:v1:' + ss.getId();
}

function swCalendarMonthReadModelCacheKey_(ss) {
  return 'sw:calendarMonthReadModel:v1:' + ss.getId();
}

function swAdminDashboardReadModelCacheKey_(ss) {
  return 'sw:adminDashboardReadModel:v1:' + ss.getId();
}

function swCacheDiamondReadModelRecords_(ss, records, status) {
  var key = swDiamondReadModelCacheKey_(ss);
  var payload = {
    cachedAt: swIso_(new Date()),
    modelBuiltAt: status && status.builtAt || '',
    rows: records || []
  };
  try {
    SW_DIAMOND_READ_MODEL_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_DIAMOND_READ_MODEL_CACHE_SECONDS * 1000,
      modelBuiltAt: payload.modelBuiltAt,
      rows: payload.rows
    };
  } catch (_) {}
  if (typeof swTaskListCachePut_ === 'function') return swTaskListCachePut_(key, payload);
  return { ok: false, reason: 'chunkCacheUnavailable', chunks: 0, bytes: 0 };
}

function swCachedDiamondReadModelRecords_(ss, status) {
  var key = swDiamondReadModelCacheKey_(ss);
  var builtAt = status && status.builtAt || '';
  var now = new Date().getTime();
  try {
    var memory = SW_DIAMOND_READ_MODEL_MEMORY_CACHE_[key];
    if (memory && memory.expiresAt > now && (!builtAt || memory.modelBuiltAt === builtAt)) {
      return { rows: memory.rows || [], ageSeconds: status.ageSeconds || 0, cacheHit: true };
    }
  } catch (_) {}
  try {
    var payload = typeof swTaskListCacheGet_ === 'function' ? swTaskListCacheGet_(key) : null;
    if (payload && (!builtAt || payload.modelBuiltAt === builtAt)) {
      SW_DIAMOND_READ_MODEL_MEMORY_CACHE_[key] = {
        expiresAt: now + SW_DIAMOND_READ_MODEL_CACHE_SECONDS * 1000,
        modelBuiltAt: payload.modelBuiltAt || '',
        rows: payload.rows || []
      };
      return { rows: payload.rows || [], ageSeconds: status.ageSeconds || 0, cacheHit: true };
    }
  } catch (_) {}
  return null;
}

function swInvalidateDiamondReadModelCache_(ss) {
  var key = swDiamondReadModelCacheKey_(ss);
  try { delete SW_DIAMOND_READ_MODEL_MEMORY_CACHE_[key]; } catch (_) {}
  try { if (typeof swTaskListCacheRemove_ === 'function') swTaskListCacheRemove_(key); } catch (_) {}
}

function swCacheCalendarMonthReadModels_(ss, rows, status) {
  var key = swCalendarMonthReadModelCacheKey_(ss);
  var months = {};
  (rows || []).forEach(function (row) {
    var monthKey = '';
    var monthLabel = '';
    var appointmentCount = 0;
    var appointments = [];
    if (Array.isArray(row)) {
      monthKey = swTrim_(row[0]);
      monthLabel = row[1] || '';
      appointmentCount = Number(row[2] || 0) || 0;
      appointments = swParseJson_(row[3] || '[]', []);
    } else {
      monthKey = swTrim_(row['Month Key']);
      monthLabel = row['Month Label'] || '';
      appointmentCount = Number(row['Appointment Count'] || 0) || 0;
      appointments = swParseJson_(row['Appointments JSON'] || '[]', []);
    }
    if (!monthKey) return;
    if (!Array.isArray(appointments)) appointments = [];
    months[monthKey] = {
      monthKey: monthKey,
      monthLabel: monthLabel,
      appointmentCount: appointmentCount,
      appointments: appointments
    };
  });
  var payload = {
    cachedAt: swIso_(new Date()),
    modelBuiltAt: status && status.builtAt || '',
    months: months
  };
  try {
    SW_CALENDAR_MONTH_READ_MODEL_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_CALENDAR_MONTH_READ_MODEL_CACHE_SECONDS * 1000,
      modelBuiltAt: payload.modelBuiltAt,
      months: payload.months
    };
  } catch (_) {}
  if (typeof swTaskListCachePut_ === 'function') return swTaskListCachePut_(key, payload);
  return { ok: false, reason: 'chunkCacheUnavailable', chunks: 0, bytes: 0 };
}

function swCachedCalendarMonthReadModels_(ss, status) {
  var key = swCalendarMonthReadModelCacheKey_(ss);
  var builtAt = status && status.builtAt || '';
  var now = new Date().getTime();
  try {
    var memory = SW_CALENDAR_MONTH_READ_MODEL_MEMORY_CACHE_[key];
    if (memory && memory.expiresAt > now && (!builtAt || memory.modelBuiltAt === builtAt)) {
      return { months: memory.months || {}, cacheHit: true };
    }
  } catch (_) {}
  try {
    var payload = typeof swTaskListCacheGet_ === 'function' ? swTaskListCacheGet_(key) : null;
    if (payload && (!builtAt || payload.modelBuiltAt === builtAt)) {
      SW_CALENDAR_MONTH_READ_MODEL_MEMORY_CACHE_[key] = {
        expiresAt: now + SW_CALENDAR_MONTH_READ_MODEL_CACHE_SECONDS * 1000,
        modelBuiltAt: payload.modelBuiltAt || '',
        months: payload.months || {}
      };
      return { months: payload.months || {}, cacheHit: true };
    }
  } catch (_) {}
  return null;
}

function swInvalidateCalendarMonthReadModelCache_(ss) {
  var key = swCalendarMonthReadModelCacheKey_(ss);
  try { delete SW_CALENDAR_MONTH_READ_MODEL_MEMORY_CACHE_[key]; } catch (_) {}
  try { if (typeof swTaskListCacheRemove_ === 'function') swTaskListCacheRemove_(key); } catch (_) {}
}

function swCacheAdminDashboardReadModels_(ss, rows, status) {
  var key = swAdminDashboardReadModelCacheKey_(ss);
  var payloads = {};
  (rows || []).forEach(function (row) {
    var rowKey = Array.isArray(row) ? row[0] : row['Key'];
    var text = Array.isArray(row) ? row[7] : row['Payload JSON'];
    if (!rowKey || !text) return;
    var payload = swParseJson_(text, null);
    if (payload) payloads[rowKey] = payload;
  });
  var payload = {
    cachedAt: swIso_(new Date()),
    modelBuiltAt: status && status.builtAt || '',
    payloads: payloads
  };
  try {
    SW_ADMIN_DASHBOARD_READ_MODEL_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_ADMIN_DASHBOARD_READ_MODEL_CACHE_SECONDS * 1000,
      modelBuiltAt: payload.modelBuiltAt,
      payloads: payload.payloads
    };
  } catch (_) {}
  if (typeof swTaskListCachePut_ === 'function') return swTaskListCachePut_(key, payload);
  return { ok: false, reason: 'chunkCacheUnavailable', chunks: 0, bytes: 0 };
}

function swCachedAdminDashboardReadModels_(ss, status) {
  var key = swAdminDashboardReadModelCacheKey_(ss);
  var builtAt = status && status.builtAt || '';
  var now = new Date().getTime();
  try {
    var memory = SW_ADMIN_DASHBOARD_READ_MODEL_MEMORY_CACHE_[key];
    if (memory && memory.expiresAt > now && (!builtAt || memory.modelBuiltAt === builtAt)) {
      return { payloads: memory.payloads || {}, cacheHit: true };
    }
  } catch (_) {}
  try {
    var payload = typeof swTaskListCacheGet_ === 'function' ? swTaskListCacheGet_(key) : null;
    if (payload && (!builtAt || payload.modelBuiltAt === builtAt)) {
      SW_ADMIN_DASHBOARD_READ_MODEL_MEMORY_CACHE_[key] = {
        expiresAt: now + SW_ADMIN_DASHBOARD_READ_MODEL_CACHE_SECONDS * 1000,
        modelBuiltAt: payload.modelBuiltAt || '',
        payloads: payload.payloads || {}
      };
      return { payloads: payload.payloads || {}, cacheHit: true };
    }
  } catch (_) {}
  return null;
}

function swInvalidateAdminDashboardReadModelCache_(ss) {
  var key = swAdminDashboardReadModelCacheKey_(ss);
  try { delete SW_ADMIN_DASHBOARD_READ_MODEL_MEMORY_CACHE_[key]; } catch (_) {}
  try { if (typeof swTaskListCacheRemove_ === 'function') swTaskListCacheRemove_(key); } catch (_) {}
}

function swReadModelFreshStatus_(ss, modelName, sheetName) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return { fresh: false, reason: 'missingSheet' };
  var meta = null;
  var rows = swReadModelMetaRows_(ss);
  for (var i = 0; i < rows.length; i++) {
    if (swTrim_(rows[i]['Model']) === modelName) {
      meta = rows[i];
      break;
    }
  }
  if (!meta) return { fresh: false, reason: 'missingMeta' };
  var metaVersion = swTrim_(meta['Version']);
  if (metaVersion !== SW_READ_MODEL_VERSION) {
    return { fresh: false, reason: 'versionMismatch', actualVersion: metaVersion, expectedVersion: SW_READ_MODEL_VERSION };
  }
  if (swTrim_(meta['Status']) !== 'OK') return { fresh: false, reason: 'status:' + swTrim_(meta['Status']) };
  if (swTrim_(meta['Invalidated At'])) return { fresh: false, reason: 'invalidated' };
  var builtAtMs = swReadModelDateMs_(meta['Built At']);
  var expiresAtMs = swReadModelDateMs_(meta['Expires At']);
  var nowMs = new Date().getTime();
  var ageSeconds = builtAtMs ? Math.max(0, Math.round((nowMs - builtAtMs) / 1000)) : 0;
  if (!builtAtMs || !expiresAtMs) return { fresh: false, reason: 'missingDates', ageSeconds: ageSeconds };
  if (expiresAtMs < nowMs) return { fresh: false, reason: 'expired', ageSeconds: ageSeconds };
  return { fresh: true, reason: '', ageSeconds: ageSeconds, builtAt: meta['Built At'] || '', expiresAt: meta['Expires At'] || '', rows: Math.max(0, sh.getLastRow() - 1) };
}

function swReadModelServingFlag_(config, key, fallback) {
  return swNorm_(swConfigValue_(config || [], 'SYSTEM', key, fallback || 'Y')) !== 'n';
}

function swInvalidateDiamondReadModelsAfterWrite_(ss, reason) {
  try { swInvalidateDiamondReadModelCache_(ss || swSpreadsheet_()); } catch (_) {}
  try { swMarkWorkflowReadModelsStale_(ss || swSpreadsheet_(), reason || 'Diamond source updated', 'diamonds'); } catch (_) {}
  try { swMarkWorkflowReadModelsStale_(ss || swSpreadsheet_(), reason || 'Diamond source updated', 'diamondRoots'); } catch (_) {}
  try { swMarkWorkflowReadModelsStale_(ss || swSpreadsheet_(), reason || 'Diamond source updated', 'customers'); } catch (_) {}
  try { swMarkWorkflowReadModelsStale_(ss || swSpreadsheet_(), reason || 'Diamond source updated', 'appointments'); } catch (_) {}
  try { swMarkWorkflowReadModelsStale_(ss || swSpreadsheet_(), reason || 'Diamond source updated', 'adminDashboard'); } catch (_) {}
}

function swInvalidatePaymentReadModelsAfterWrite_(ss, reason) {
  try { swInvalidateAdminDashboardReadModelCache_(ss || swSpreadsheet_()); } catch (_) {}
  try { swMarkWorkflowReadModelsStale_(ss || swSpreadsheet_(), reason || 'Payment source updated', 'payments'); } catch (_) {}
  try { swMarkWorkflowReadModelsStale_(ss || swSpreadsheet_(), reason || 'Payment source updated', 'customers'); } catch (_) {}
  try { swMarkWorkflowReadModelsStale_(ss || swSpreadsheet_(), reason || 'Payment source updated', 'appointments'); } catch (_) {}
  try { swMarkWorkflowReadModelsStale_(ss || swSpreadsheet_(), reason || 'Payment source updated', 'adminDashboard'); } catch (_) {}
  try {
    if (typeof swInvalidateCustomerSearchReadModelCache_ === 'function') {
      swInvalidateCustomerSearchReadModelCache_(ss || swSpreadsheet_());
    }
  } catch (_) {}
}

function swInvalidateAppointmentReadModelsAfterWrite_(ss, reason) {
  try { swInvalidateCalendarMonthReadModelCache_(ss || swSpreadsheet_()); } catch (_) {}
  try { swInvalidateAdminDashboardReadModelCache_(ss || swSpreadsheet_()); } catch (_) {}
  try { swMarkWorkflowReadModelsStale_(ss || swSpreadsheet_(), reason || 'Appointment source updated', 'appointments'); } catch (_) {}
  try { swMarkWorkflowReadModelsStale_(ss || swSpreadsheet_(), reason || 'Appointment source updated', 'calendarMonths'); } catch (_) {}
  try { swMarkWorkflowReadModelsStale_(ss || swSpreadsheet_(), reason || 'Appointment source updated', 'customers'); } catch (_) {}
  try { swMarkWorkflowReadModelsStale_(ss || swSpreadsheet_(), reason || 'Appointment source updated', 'adminDashboard'); } catch (_) {}
}
