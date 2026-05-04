/**
 * Sales workflow data repository: context, source appointment reads, roster, and schedule indexes.
 */

function swReadRosterAvailabilityIndex_(ss) {
  var out = { exists: false, schemaOk: false, byName: {} };
  var roster = ss.getSheetByName(SW_SHEETS.ROSTER);
  if (!roster || roster.getLastRow() < 2) return out;
  out.exists = true;

  var values = roster.getDataRange().getValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var repCol = swPickIndex_(H, ['Rep', 'Name', 'Team Member']);
  var dayCols = {
    Sun: swPickIndex_(H, ['Sun']),
    Mon: swPickIndex_(H, ['Mon']),
    Tue: swPickIndex_(H, ['Tue']),
    Wed: swPickIndex_(H, ['Wed']),
    Thu: swPickIndex_(H, ['Thu']),
    Fri: swPickIndex_(H, ['Fri']),
    Sat: swPickIndex_(H, ['Sat'])
  };
  out.schemaOk = repCol >= 0 && Object.keys(dayCols).some(function (day) { return dayCols[day] >= 0; });
  if (!out.schemaOk) return out;

  for (var i = 1; i < values.length; i++) {
    var rowName = swTrim_(values[i][repCol]);
    if (!rowName) continue;
    var row = { name: rowName, days: {} };
    Object.keys(dayCols).forEach(function (day) {
      row.days[day] = dayCols[day] >= 0 ? swTruthy_(values[i][dayCols[day]]) : null;
    });
    out.byName[swNorm_(rowName)] = row;
  }
  return out;
}

function swReadScheduleChangesIndex_(ss) {
  var out = { byNameDate: {} };
  var sh = ss.getSheetByName(SW_SHEETS.SCHEDULE_CHANGES);
  if (!sh || sh.getLastRow() < 2) return out;

  var values = sh.getDataRange().getValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var nameCol = swPickIndex_(H, ['Rep Name', 'Rep', 'Name']);
  var dateCol = swPickIndex_(H, ['Change Date', 'Date']);
  var typeCol = swPickIndex_(H, ['Change Type', 'Status', 'Override Status']);
  if (nameCol < 0 || dateCol < 0) return out;

  for (var i = 1; i < values.length; i++) {
    var name = swNorm_(values[i][nameCol]);
    var date = swDateKey_(values[i][dateCol]);
    if (!name || !date) continue;
    out.byNameDate[name + '|' + date] = {
      changeType: typeCol >= 0 ? swTrim_(values[i][typeCol]) : 'Full-day off'
    };
  }
  return out;
}

function swBuildIdentityContext_(ss, readOnly) {
  var config = swReadConfig_(ss, readOnly);
  var peopleIndex = swReadPeopleIndex_(ss, config);
  var admins = swReadAdminsFromConfig_(config);
  return {
    tz: swTimezone_(),
    config: config,
    peopleIndex: peopleIndex,
    assistedRoster: peopleIndex.assistedRoster,
    admins: admins,
    lookbackDays: Number(swConfigValue_(config, 'SYSTEM', 'WORKFLOW_LOOKBACK_DAYS', '14')) || 14,
    futureDays: Number(swConfigValue_(config, 'SYSTEM', 'WORKFLOW_FUTURE_DAYS', '365')) || 365
  };
}

function swBuildTaskDetailContext_(ss, readOnly) {
  var ctx = swBuildIdentityContext_(ss, readOnly);
  ctx.templates = swReadTemplates_(ss, readOnly);
  return ctx;
}

function swBuildTaskDetailReadContext_(ss, readOnly) {
  var user = swCurrentUserConfigOnly_(ss, readOnly);
  if (user.isAdmin) {
    return {
      user: user,
      templates: swReadTemplates_(ss, readOnly),
      lightweight: true
    };
  }

  var ctx = swBuildTaskDetailContext_(ss, readOnly);
  ctx.user = swCurrentUser_(ss, ctx);
  ctx.lightweight = false;
  return ctx;
}

function swBuildContext_(ss, readOnly) {
  var ctx = swBuildTaskDetailContext_(ss, readOnly);
  ctx.rosterIndex = swReadRosterAvailabilityIndex_(ss);
  ctx.scheduleChangesIndex = swReadScheduleChangesIndex_(ss);
  ctx.waxIndex = swReadWaxRequestIndex_(ss);
  return ctx;
}

function swReadWaxRequestIndex_(ss) {
  var out = { byRoot: {}, activeByRoot: {}, needsUpdateByRoot: {}, statusOptions: [] };
  var sheetName = (typeof WAX !== 'undefined' && WAX.SHEET) ? WAX.SHEET : '05_Wax_Requests';
  var sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < 1) return out;

  var values = sh.getDataRange().getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var C = {
    id: swPickIndex_(H, ['WaxRequestID']),
    root: swPickIndex_(H, ['RootApptID']),
    so: swPickIndex_(H, ['SO/MO Number', 'SO Number', 'SO#']),
    customer: swPickIndex_(H, ['Customer Name']),
    priority: swPickIndex_(H, ['Priority']),
    status: swPickIndex_(H, ['Wax Print Status']),
    repNeed: swPickIndex_(H, ['Needed By (Rep)', 'Needed by (Rep)', 'Rep Needed By']),
    adminDeadline: swPickIndex_(H, ['Wax Deadline (Admin)', 'Wax Admin Deadline']),
    estPrint: swPickIndex_(H, ['Estimated Print Date']),
    completed: swPickIndex_(H, ['Completed Print Date']),
    notes: swPickIndex_(H, ['Status Notes']),
    link: swPickIndex_(H, ['Master Row Link'])
  };
  var now = new Date();
  var todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0, 0).getTime();

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var root = swTrim_(swCell_(row, C.root));
    if (!root) continue;
    var status = swTrim_(swCell_(row, C.status));
    var statusNorm = swNorm_(status);
    var active = !/(^|\s)(completed|canceled|cancelled)(\s|$)/.test(statusNorm);
    var adminDeadline = swTrim_(swCell_(row, C.adminDeadline));
    var adminMs = adminDeadline ? swDateValue_(adminDeadline) : 0;
    var needsUpdate = active && (!status || !adminDeadline || (adminMs && adminMs < todayStart));
    var rowUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + sh.getSheetId() + '&range=A' + (i + 1);
    var item = {
      id: swTrim_(swCell_(row, C.id)),
      root: root,
      so: swTrim_(swCell_(row, C.so)),
      customerName: swTrim_(swCell_(row, C.customer)),
      priority: swTrim_(swCell_(row, C.priority)),
      status: status,
      repNeed: swTrim_(swCell_(row, C.repNeed)),
      adminDeadline: adminDeadline,
      estPrint: swTrim_(swCell_(row, C.estPrint)),
      completed: swTrim_(swCell_(row, C.completed)),
      notes: swTrim_(swCell_(row, C.notes)),
      link: swTrim_(swCell_(row, C.link)) || rowUrl,
      rowUrl: rowUrl,
      active: active,
      needsUpdate: needsUpdate
    };
    if (!out.byRoot[root]) out.byRoot[root] = [];
    out.byRoot[root].push(item);
    if (active) {
      if (!out.activeByRoot[root]) out.activeByRoot[root] = [];
      out.activeByRoot[root].push(item);
    }
    if (needsUpdate) {
      if (!out.needsUpdateByRoot[root]) out.needsUpdateByRoot[root] = [];
      out.needsUpdateByRoot[root].push(item);
    }
  }
  try {
    if (typeof wax_statusOptions === 'function') out.statusOptions = wax_statusOptions();
  } catch (_) {}
  return out;
}

function swReadAppointments_(ss) {
  var sh = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!sh) throw new Error('Missing sheet: ' + SW_SHEETS.MASTER);
  if (sh.getLastRow() < 2 || sh.getLastColumn() < 1) return [];

  var values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var display = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  var headers = display[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var idx = {
    appt: swPickIndex_(H, ['APPT_ID', 'Appt ID', 'Appointment ID']),
    root: swPickIndex_(H, ['RootApptID', 'Root Appt ID', 'Root Appointment ID']),
    uid: swPickIndex_(H, ['CalendlyEventUID', 'Calendly Event UID', 'Admin: Calendly Event UID', 'Acuity ID', 'UID']),
    name: swPickIndex_(H, ['Customer Name', 'Customer', 'Client Name', 'Name']),
    emailLower: swPickIndex_(H, ['EmailLower', 'Email Lower']),
    email: swPickIndex_(H, ['Email', 'Email Address', 'E-mail']),
    phoneNorm: swPickIndex_(H, ['PhoneNorm', 'Phone Norm']),
    phone: swPickIndex_(H, ['Phone', 'Phone Number', 'Mobile', 'Tel']),
    brand: swPickIndex_(H, ['Brand', 'Company']),
    visitDate: swPickIndex_(H, ['Visit Date', 'Appointment Date', 'Date']),
    visitTime: swPickIndex_(H, ['Visit Time', 'Appointment Time', 'Time']),
    visitType: swPickIndex_(H, ['Visit Type', 'Appointment Type']),
    status: swPickIndex_(H, ['Status']),
    active: swPickIndex_(H, ['Active?', 'Active', 'Is Active']),
    assignedRep: swPickIndex_(H, ['Assigned Rep', 'Rep', 'Owner']),
    assignedRepEmail: swPickIndex_(H, ['Assigned Rep Email', 'Rep Email', 'Owner Email']),
    assistedRep: swPickIndex_(H, ['Assisted Rep', 'Assistant Rep']),
    assistedRepEmail: swPickIndex_(H, ['Assisted Rep Email', 'Assistant Rep Email']),
    clientFolder: swPickIndex_(H, ['Client Folder', 'ClientFolderURL', 'Client Folder URL']),
    reportUrl: swPickIndex_(H, ['Client Status Report URL', 'Report URL']),
    quotationUrl: swPickIndex_(H, ['Quotation URL', 'QuotationURL', 'Quote URL']),
    tracker3d: swPickIndex_(H, ['3D Tracker', '3D Log', '3D Tracker URL']),
    salesStage: swPickIndex_(H, ['Sales Stage']),
    convStatus: swPickIndex_(H, ['Conversion Status']),
    customOrder: swPickIndex_(H, ['Custom Order Status']),
    inProduction: swPickIndex_(H, ['In Production Status']),
    nextSteps: swPickIndex_(H, ['Next Steps']),
    designRequest: swPickIndex_(H, ['Design Request']),
    deadline3d: swPickIndex_(H, ['3D Deadline']),
    productionDeadline: swPickIndex_(H, ['Production Deadline', 'Prod. Deadline']),
    waxStatus: swPickIndex_(H, ['Wax Print Status']),
    waxDeadlineAdmin: swPickIndex_(H, ['Wax Deadline (Admin)', 'Wax Admin Deadline']),
    waxRequestUrl: swPickIndex_(H, ['Wax Request URL']),
    centerStoneStatus: swPickIndex_(H, ['Center Stone Order Status', 'Center Stone Status', 'CSOS', 'Diamond Memo Status', 'DV Status']),
    dvStonesJson: swPickIndex_(H, ['DV Stones (JSON Lines)', 'DV Stones JSON Lines', 'DV Stones-JSON Lines']),
    dvStonesSummary: swPickIndex_(H, ['DV Stones Summary', 'DV Stones- Summary']),
    dvCustomerLookingFor: swPickIndex_(H, ['DV Customer Looking For', 'Diamond Customer Looking For', 'Customer Diamond Requirements']),
    dvVarietyStrategy: swPickIndex_(H, ['DV Variety Strategy', 'Diamond Variety Strategy']),
    dvCustomerRequirementsJson: swPickIndex_(H, ['DV Customer Requirements (JSON)', 'DV Customer Requirements JSON', 'Customer Diamond Requirements JSON']),
    so: swPickIndex_(H, ['SO#', 'SO #', 'SO']),
    orderFolder: swPickIndex_(H, ['Order Folder', '05-3D Folder'])
  };

  var out = [];
  for (var i = 1; i < values.length; i++) {
    var drow = display[i];
    var vrow = values[i];
    var rec = {
      row: i + 1,
      appt: swTrim_(swCell_(drow, idx.appt)),
      root: swTrim_(swCell_(drow, idx.root)),
      uid: swTrim_(swCell_(drow, idx.uid)),
      name: swTrim_(swCell_(drow, idx.name)),
      email: swNormEmail_(swCell_(drow, idx.emailLower) || swCell_(drow, idx.email)),
      phone: swNormPhone_(swCell_(drow, idx.phoneNorm) || swCell_(drow, idx.phone)),
      brand: swTrim_(swCell_(drow, idx.brand)),
      visitDate: swTrim_(swCell_(drow, idx.visitDate)),
      visitTime: swFormatAppointmentTime_(swCell_(drow, idx.visitTime), swCell_(vrow, idx.visitTime)),
      visitType: swTrim_(swCell_(drow, idx.visitType)),
      visitDateRaw: swCell_(vrow, idx.visitDate),
      visitTimeRaw: swCell_(vrow, idx.visitTime),
      status: swTrim_(swCell_(drow, idx.status)),
      active: swTrim_(swCell_(drow, idx.active)),
      assignedRep: swTrim_(swCell_(drow, idx.assignedRep)),
      assignedRepEmail: swNormEmail_(swCell_(drow, idx.assignedRepEmail)),
      assistedRep: swTrim_(swCell_(drow, idx.assistedRep)),
      assistedRepEmail: swNormEmail_(swCell_(drow, idx.assistedRepEmail)),
      clientFolder: swTrim_(swCell_(drow, idx.clientFolder)),
      reportUrl: swTrim_(swCell_(drow, idx.reportUrl)),
      quotationUrl: swTrim_(swCell_(drow, idx.quotationUrl)),
      tracker3dUrl: swTrim_(swCell_(drow, idx.tracker3d)),
      salesStage: swTrim_(swCell_(drow, idx.salesStage)),
      convStatus: swTrim_(swCell_(drow, idx.convStatus)),
      customOrder: swTrim_(swCell_(drow, idx.customOrder)),
      inProduction: swTrim_(swCell_(drow, idx.inProduction)),
      nextSteps: swTrim_(swCell_(drow, idx.nextSteps)),
      designRequest: swTrim_(swCell_(drow, idx.designRequest)),
      deadline3d: swTrim_(swCell_(drow, idx.deadline3d)),
      productionDeadline: swTrim_(swCell_(drow, idx.productionDeadline)),
      waxStatus: swTrim_(swCell_(drow, idx.waxStatus)),
      waxDeadlineAdmin: swTrim_(swCell_(drow, idx.waxDeadlineAdmin)),
      waxRequestUrl: swTrim_(swCell_(drow, idx.waxRequestUrl)),
      centerStoneStatus: swTrim_(swCell_(drow, idx.centerStoneStatus)),
      dvStonesJson: swTrim_(swCell_(drow, idx.dvStonesJson)),
      dvStonesSummary: swTrim_(swCell_(drow, idx.dvStonesSummary)),
      dvCustomerLookingFor: swTrim_(swCell_(drow, idx.dvCustomerLookingFor)),
      dvVarietyStrategy: swTrim_(swCell_(drow, idx.dvVarietyStrategy)),
      dvCustomerRequirementsJson: swTrim_(swCell_(drow, idx.dvCustomerRequirementsJson)),
      so: swTrim_(swCell_(drow, idx.so)),
      orderFolder: swTrim_(swCell_(drow, idx.orderFolder))
    };
    rec.root = rec.root || rec.appt;
    rec.statusNorm = swNorm_(rec.status);
    rec.activeNorm = swNorm_(rec.active);
    out.push(rec);
  }
  return out;
}
