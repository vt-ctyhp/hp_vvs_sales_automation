/**
 * One-time Loupe360 diamond sync for the 200_ tracker.
 *
 * The upload flow converts a spreadsheet upload into a temporary Google Sheet,
 * previews the tracker delta, then applies by Certificate No. The uploaded
 * source is treated as current for shipment, status, and spec facts while
 * customer/advisor/JOC assignment fields remain manual workflow data.
 */

var SW_LOUPE360_VENDOR = 'Loupe360';
var SW_LOUPE360_SYNC_HEADER = 'Loupe360 Last Sync At';

function sw_previewLoupe360DiamondSync(form) {
  return swTimed_('sw_previewLoupe360DiamondSync', function () {
    form = form || {};
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, form.swAuthToken || form.authToken || '');
    swRequireDiamondSyncUser_(user);

    var blobs = swLoupe360FormBlobs_(form.loupe360DiamondFile);
    if (!blobs.length) throw new Error('Choose a Loupe360 spreadsheet file.');
    var temp = null;
    try {
      temp = swLoupe360UploadToTempSheet_(blobs[0], user);
      var source = swLoupe360ReadSourceRows_(temp.id);
      var target = swLoupe360ReadTarget_();
      var plan = swLoupe360BuildPlan_(source.rows, target);
      return swLoupe360PreviewResponse_(temp, source, target, plan);
    } catch (err) {
      if (temp && temp.id) {
        try { DriveApp.getFileById(temp.id).setTrashed(true); } catch (_) {}
      }
      throw err;
    }
  });
}

function sw_applyLoupe360DiamondSync(authToken, syncId) {
  return swTimed_('sw_applyLoupe360DiamondSync', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken || '');
    swRequireDiamondSyncUser_(user);
    syncId = swTrim_(syncId);
    if (!syncId) throw new Error('Missing sync preview. Upload the Loupe360 spreadsheet again.');

    var lock = LockService.getDocumentLock() || LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var source = swLoupe360ReadSourceRows_(syncId);
      var target = swLoupe360ReadTarget_();
      var cSync = swDiamondEnsure200Column_(target.sheet, SW_LOUPE360_SYNC_HEADER);
      target = swLoupe360ReadTarget_();
      target.columns.syncAt = cSync;
      var plan = swLoupe360BuildPlan_(source.rows, target);
      var result = swLoupe360ApplyPlan_(target, plan, user);

      try {
        swAppendTaskLog_(ss, 'LOUPE360_DIAMOND_SYNC', {
          taskId: '',
          root: '',
          appt: '',
          taskType: 'LOUPE360_DIAMOND_SYNC',
          status: SW_STATUSES.COMPLETED
        }, user, '', '', {
          sourceRows: source.rows.length,
          updated: result.updated,
          appended: result.appended,
          skipped: result.skipped,
          statusOverwrites: plan.stats.statusOverwrites,
          duplicateSourceCerts: plan.stats.duplicateSourceCerts,
          duplicateTrackerCerts: plan.stats.duplicateTrackerCerts
        });
      } catch (_) {}

      var generation = null;
      try { generation = sw_generateSalesWorkflowTasks(); } catch (genErr) {
        generation = { ok: false, error: swTrim_(genErr && genErr.message || genErr) };
      }
      try { DriveApp.getFileById(syncId).setTrashed(true); } catch (_) {}
      return {
        ok: true,
        updated: result.updated,
        appended: result.appended,
        skipped: result.skipped,
        statusOverwrites: plan.stats.statusOverwrites,
        duplicateSourceCerts: plan.stats.duplicateSourceCerts,
        duplicateTrackerCerts: plan.stats.duplicateTrackerCerts,
        assignmentMissing: swLoupe360CountMissingAssignments_(),
        spreadsheetUrl: target.ss.getUrl(),
        tab: target.tab,
        generation: generation
      };
    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }
  });
}

function sw_assignInStockDiamond(authToken, payload) {
  return swTimed_('sw_assignInStockDiamond', function () {
    payload = payload || {};
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken || '');
    swRequireInStockDiamondAssignmentUser_(user);

    var rowIndex = Number(payload.rowIndex);
    if (!(rowIndex >= 3)) throw new Error('Missing 200_ row.');
    var target = swDiamond200Target_();
    if (!target || !target.sheet) throw new Error('Diamond tracking sheet is unavailable.');
    var sh = target.sheet;
    if (rowIndex > sh.getLastRow()) throw new Error('Diamond row no longer exists.');

    var lock = LockService.getDocumentLock() || LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var hm = swDiamond200HeaderMap_(sh);
      var C = {
        root: swDiamondFind200Column_(hm, ['RootApptID', 'APPT_ID', 'Root Appt ID']),
        customerName: swDiamondFind200Column_(hm, ['Customer Name', 'Client Name', 'Customer']),
        assignedRep: swDiamondFind200Column_(hm, ['Client Advisor', 'Assigned Rep', 'Sales Rep']),
        joc: swDiamondFind200Column_(hm, ['JOC', 'Assisted Rep', 'Assistant Rep']),
        certNo: swDiamondFind200Column_(hm, ['Certificate No', 'Cert #', 'Cert No', 'Certificate #']),
        orderStatus: swDiamondFind200Column_(hm, ['Order Status', 'OrderStatus']),
        stoneStatus: swDiamondFind200Column_(hm, ['Stone Status', 'StoneStatus']),
        decision: swDiamondFind200Column_(hm, ['Stone Decision (PO, Return)', 'Stone Decision', 'StoneDecision'])
      };
      if (!C.customerName) throw new Error('Customer Name column is missing in 200_.');
      if (!C.assignedRep) throw new Error('Client Advisor / Assigned Rep column is missing in 200_.');
      if (!C.root) throw new Error('RootApptID column is missing in 200_.');
      if (!C.joc) C.joc = swDiamondEnsure200Column_(sh, 'JOC');

      var row = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
      var expectedCert = swTrim_(payload.certNo);
      var actualCert = swDiamondCell_(row, C.certNo);
      if (expectedCert && actualCert && swLoupe360CertKey_(expectedCert) !== swLoupe360CertKey_(actualCert)) {
        throw new Error('Diamond row changed. Reload in-stock diamonds and try again.');
      }
      if (!swLoupe360IsAssignableStockRow_(row, C)) {
        throw new Error('This diamond is no longer eligible for stock assignment.');
      }

      var customerName = swTrim_(payload.customerName);
      var root = swTrim_(payload.root);
      var assignedRep = swTrim_(payload.assignedRep);
      var joc = swTrim_(payload.joc);
      if (!customerName && !root && !assignedRep && !joc) {
        throw new Error('Enter at least one assignment field.');
      }

      if (customerName) sh.getRange(rowIndex, C.customerName).setValue(customerName);
      if (root) sh.getRange(rowIndex, C.root).setValue(root);
      if (assignedRep) sh.getRange(rowIndex, C.assignedRep).setValue(assignedRep);
      if (joc) sh.getRange(rowIndex, C.joc).setValue(joc);

      try {
        swAppendTaskLog_(ss, 'IN_STOCK_DIAMOND_ASSIGN', {
          taskId: '',
          root: root || swDiamondCell_(row, C.root),
          appt: '',
          taskType: 'IN_STOCK_DIAMOND_ASSIGN',
          status: SW_STATUSES.COMPLETED
        }, user, '', assignedRep || '', {
          rowIndex: rowIndex,
          certNo: actualCert,
          customerName: customerName,
          root: root,
          assignedRep: assignedRep,
          joc: joc
        });
      } catch (_) {}

      var finalRoot = root || swDiamondCell_(row, C.root);
      if (finalRoot) {
        try {
          if (typeof dp_computeCountsForAppointment_ === 'function' && typeof dp_refresh100QuickRef_ === 'function') {
            var counts = dp_computeCountsForAppointment_(sh, hm, finalRoot);
            dp_refresh100QuickRef_(finalRoot, counts, sh, hm);
          }
        } catch (_) {}
      }
      var generation = null;
      try { generation = sw_generateSalesWorkflowTasks(); } catch (genErr) {
        generation = { ok: false, error: swTrim_(genErr && genErr.message || genErr) };
      }
      return { ok: true, rowIndex: rowIndex, generation: generation };
    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }
  });
}

function swRequireDiamondSyncUser_(user) {
  if (!(user && user.isDiamondOrderAdmin)) {
    throw new Error('Diamond order admin access required.');
  }
}

function swRequireInStockDiamondAssignmentUser_(user) {
  if (!(user && (user.isAdmin || user.isRep || user.isJoc || user.isDiamondOrderAdmin))) {
    throw new Error('Workflow user access required.');
  }
}

function swLoupe360FormBlobs_(value) {
  if (!value) return [];
  var values = Array.isArray(value) ? value : [value];
  return values.filter(function (blob) {
    return blob && typeof blob.getBytes === 'function' && blob.getBytes().length;
  });
}

function swLoupe360UploadToTempSheet_(blob, user) {
  var now = new Date();
  var name = 'Loupe360 Diamond Sync ' + Utilities.formatDate(now, swTimezone_(), 'yyyyMMdd-HHmmss');
  var upload = blob.copyBlob ? blob.copyBlob() : Utilities.newBlob(blob.getBytes(), blob.getContentType(), blob.getName ? blob.getName() : name);
  upload.setName(name);
  var file = Drive.Files.insert({
    title: name,
    mimeType: MimeType.GOOGLE_SHEETS,
    description: 'Temporary diamond sync upload for ' + (user && user.email || 'workflow user')
  }, upload, {
    convert: true,
    supportsAllDrives: true
  });
  return { id: file.id, name: name, url: file.alternateLink || ('https://docs.google.com/spreadsheets/d/' + file.id) };
}

function swLoupe360ReadSourceRows_(spreadsheetId) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sh = ss.getSheets()[0];
  if (!sh || sh.getLastRow() < 2) throw new Error('The Loupe360 spreadsheet has no data rows.');
  var display = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  var values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var headers = display[0].map(swTrim_);
  var H = swHeaderMapFromArray_(headers);
  var required = ['ReportNo', 'OrderStatus', 'Shape', 'Carats', 'Col', 'Clar', 'Lab'];
  var missing = required.filter(function (h) { return swPickIndex_(H, [h]) < 0; });
  if (missing.length) throw new Error('Loupe360 spreadsheet is missing column(s): ' + missing.join(', '));

  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var raw = values[r];
    var disp = display[r];
    var certNo = swLoupe360Cell_(disp, H, ['ReportNo']);
    if (!certNo) continue;
    rows.push(swLoupe360NormalizeSourceRow_(raw, disp, H, r + 1));
  }
  if (!rows.length) throw new Error('No Loupe360 rows with ReportNo were found.');
  return { spreadsheetId: spreadsheetId, sheetName: sh.getName(), rows: rows };
}

function swLoupe360NormalizeSourceRow_(raw, disp, H, sourceRow) {
  function text(names) { return swLoupe360Cell_(disp, H, names); }
  function rawValue(names) { var idx = swPickIndex_(H, names); return idx >= 0 ? raw[idx] : ''; }
  var length = text(['Length']);
  var width = text(['Width']);
  var height = text(['Height']);
  var measurements = [length, width, height].filter(Boolean).join(' x ');
  var status = swTrim_(text(['OrderStatus'])).toUpperCase().replace(/\s+/g, '_');
  var growth = text(['GrowthMethod']);
  var stoneType = /cvd|hpht|lab/i.test(growth) || /^LG/i.test(text(['ReportNo'])) ? 'Lab Diamond' : '';
  return {
    sourceRow: sourceRow,
    certNo: text(['ReportNo']),
    certKey: swLoupe360CertKey_(text(['ReportNo'])),
    orderNumber: text(['OrderNumber']),
    invoiceNumber: text(['InvoiceNumber']),
    orderedBy: text(['OrderedBy']),
    orderDate: swLoupe360DateValue_(rawValue(['OrderDate']) || text(['OrderDate'])),
    expectedDelivery: swLoupe360DateValue_(rawValue(['ExpectedDelivery']) || text(['ExpectedDelivery'])),
    forecastedDelivery: swLoupe360DateValue_(rawValue(['ForecastedDelivery']) || text(['ForecastedDelivery'])),
    sourceStatus: status,
    sourceStatusLabel: swLoupe360StatusLabel_(status),
    vendor: SW_LOUPE360_VENDOR,
    stoneType: stoneType,
    shape: swLoupe360Title_(text(['Shape'])),
    carat: text(['Carats']),
    color: text(['Col']),
    clarity: text(['Clar']),
    cut: text(['Cut']),
    pol: text(['Pol']),
    sym: text(['Symm']),
    fluorIntensity: text(['Flo']),
    fluorColor: text(['FloCol']),
    measurements: measurements,
    ratio: text(['Ratio']),
    lab: text(['Lab']),
    growthMethod: growth
  };
}

function swLoupe360Cell_(row, H, names) {
  var idx = swPickIndex_(H, names);
  return idx >= 0 ? swTrim_(row[idx]) : '';
}

function swLoupe360ReadTarget_() {
  var target = swDiamond200Target_();
  if (!target || !target.sheet) throw new Error('Diamond tracking sheet is unavailable.');
  var sh = target.sheet;
  var hm = swDiamond200HeaderMap_(sh);
  var columns = {
    root: swDiamondFind200Column_(hm, ['RootApptID', 'APPT_ID', 'Root Appt ID']),
    customerName: swDiamondFind200Column_(hm, ['Customer Name', 'Client Name', 'Customer']),
    appointment: swDiamondFind200Column_(hm, ['Customer Appt Time & Date', 'Customer Appointment Date', 'Appointment Date']),
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
    requestDate: swDiamondFind200Column_(hm, ['Request Date', 'RequestDate']),
    requestedBy: swDiamondFind200Column_(hm, ['Requested By', 'RequestedBy']),
    orderedBy: swDiamondFind200Column_(hm, ['Ordered By', 'OrderedBy']),
    vendorOrderNumber: swDiamondFind200Column_(hm, ['Loupe360 Order #', 'Vendor Order Number', 'Order Number', 'OrderNumber']),
    invoiceNumber: swDiamondFind200Column_(hm, ['Invoice Number', 'Invoice #', 'InvoiceNumber']),
    returnDueDate: swDiamondFind200Column_(hm, ['Return DUE DATE', 'Return Due Date', 'Return Due']),
    trackingEta: swDiamondFind200Column_(hm, ['Tracking ETA', 'Tracking ETA Date', 'ETA Date', 'ETA']),
    trackingStatus: swDiamondFind200Column_(hm, ['Tracking Status', 'ETA Status', 'Shipment Status']),
    cut: swDiamondFind200Column_(hm, ['Cut']),
    pol: swDiamondFind200Column_(hm, ['Pol.', 'Pol', 'Polish']),
    sym: swDiamondFind200Column_(hm, ['Sym.', 'Sym', 'Symmetry']),
    fluorIntensity: swDiamondFind200Column_(hm, ['Fluor.Intesity', 'Fluor.Intensity', 'Fluor Intensity']),
    fluorColor: swDiamondFind200Column_(hm, ['Fluor.Color', 'Fluor Color', 'Fluorescence Color']),
    syncAt: swDiamondFind200Column_(hm, [SW_LOUPE360_SYNC_HEADER])
  };
  if (!columns.certNo) throw new Error('Certificate No column is missing in 200_.');
  if (!columns.orderStatus) throw new Error('Order Status column is missing in 200_.');
  if (!columns.stoneStatus) throw new Error('Stone Status column is missing in 200_.');

  var rows = [];
  var byCert = {};
  var lr = sh.getLastRow();
  var lc = sh.getLastColumn();
  if (lr >= 3 && lc >= 1) {
    var display = sh.getRange(3, 1, lr - 2, lc).getDisplayValues();
    display.forEach(function (row, i) {
      var rowIndex = i + 3;
      var cert = swDiamondCell_(row, columns.certNo);
      var key = swLoupe360CertKey_(cert);
      var rec = { rowIndex: rowIndex, row: row, certNo: cert, certKey: key };
      rows.push(rec);
      if (key) {
        if (!byCert[key]) byCert[key] = [];
        byCert[key].push(rec);
      }
    });
  }
  return {
    ss: target.ss,
    sheet: sh,
    tab: target.tab,
    hm: hm,
    columns: columns,
    rows: rows,
    byCert: byCert,
    lastCol: lc
  };
}

function swLoupe360BuildPlan_(sourceRows, target) {
  var sourceByCert = {};
  var sourceOrder = [];
  var duplicateSource = 0;
  var plan = {
    updates: [],
    appends: [],
    skipped: [],
    conflicts: [],
    stats: {
      sourceRows: sourceRows.length,
      matched: 0,
      updated: 0,
      appended: 0,
      skippedInactiveMissing: 0,
      statusOverwrites: 0,
      duplicateSourceCerts: 0,
      duplicateTrackerCerts: 0
    }
  };
  Object.keys(target.byCert || {}).forEach(function (key) {
    if ((target.byCert[key] || []).length > 1) plan.stats.duplicateTrackerCerts++;
  });

  sourceRows.forEach(function (src) {
    if (!src.certKey) return;
    if (sourceByCert[src.certKey]) {
      duplicateSource++;
    } else {
      sourceOrder.push(src.certKey);
    }
    sourceByCert[src.certKey] = src;
  });

  sourceOrder.forEach(function (certKey) {
    var src = sourceByCert[certKey];
    var matches = target.byCert[src.certKey] || [];
    if (!matches.length) {
      if (swLoupe360IsActiveSource_(src)) {
        plan.appends.push(src);
      } else {
        plan.stats.skippedInactiveMissing++;
        plan.skipped.push({ certNo: src.certNo, reason: 'Inactive source row not present in 200_' });
      }
      return;
    }
    plan.stats.matched += matches.length;
    matches.forEach(function (match) {
      var update = swLoupe360BuildUpdate_(src, match, target.columns);
      if (update.statusOverwrite) plan.stats.statusOverwrites++;
      if (update.changes.length) plan.updates.push(update);
    });
  });

  plan.stats.duplicateSourceCerts = duplicateSource;
  plan.stats.updated = plan.updates.length;
  plan.stats.appended = plan.appends.length;
  return plan;
}

function swLoupe360BuildUpdate_(src, targetRow, C) {
  var current = targetRow.row;
  var changes = [];
  var desired = swLoupe360DesiredStatus_(src);
  var statusOverwrite = swLoupe360WorkflowStatusConflict_(current, C, desired);

  swLoupe360MaybeChange_(changes, current, C.orderStatus, desired.orderStatus);
  swLoupe360MaybeChange_(changes, current, C.stoneStatus, desired.stoneStatus);
  swLoupe360MaybeChange_(changes, current, C.vendor, SW_LOUPE360_VENDOR);
  swLoupe360MaybeChange_(changes, current, C.stoneType, src.stoneType);
  swLoupe360MaybeChange_(changes, current, C.shape, src.shape);
  swLoupe360MaybeChange_(changes, current, C.carat, src.carat);
  swLoupe360MaybeChange_(changes, current, C.color, src.color);
  swLoupe360MaybeChange_(changes, current, C.clarity, src.clarity);
  swLoupe360MaybeChange_(changes, current, C.lab, src.lab);
  swLoupe360MaybeChange_(changes, current, C.measurement, src.measurements);
  swLoupe360MaybeChange_(changes, current, C.ratio, src.ratio);
  swLoupe360MaybeChange_(changes, current, C.cut, src.cut);
  swLoupe360MaybeChange_(changes, current, C.pol, src.pol);
  swLoupe360MaybeChange_(changes, current, C.sym, src.sym);
  swLoupe360MaybeChange_(changes, current, C.fluorIntensity, src.fluorIntensity);
  swLoupe360MaybeChange_(changes, current, C.fluorColor, src.fluorColor);
  swLoupe360MaybeChange_(changes, current, C.orderDate, src.orderDate);
  swLoupe360MaybeChange_(changes, current, C.vendorOrderNumber, src.orderNumber);
  swLoupe360MaybeChange_(changes, current, C.invoiceNumber, src.invoiceNumber);
  if (desired.trackingEta) swLoupe360MaybeChange_(changes, current, C.trackingEta, desired.trackingEta);
  if (desired.trackingStatus) swLoupe360MaybeChange_(changes, current, C.trackingStatus, desired.trackingStatus);
  if (desired.returnDueDate && !swDiamondCell_(current, C.returnDueDate)) {
    swLoupe360MaybeChange_(changes, current, C.returnDueDate, desired.returnDueDate);
  }
  return {
    rowIndex: targetRow.rowIndex,
    certNo: src.certNo,
    source: src,
    changes: changes,
    statusOverwrite: statusOverwrite
  };
}

function swLoupe360MaybeChange_(changes, row, col, next) {
  if (!col || next == null || next === '') return;
  var cur = swDiamondCell_(row, col);
  var nextText = swLoupe360DisplayValue_(next);
  if (swNorm_(cur) === swNorm_(nextText)) return;
  changes.push({ col: col, value: next });
}

function swLoupe360DesiredStatus_(src) {
  var status = src.sourceStatus;
  var out = {
    orderStatus: '',
    stoneStatus: '',
    trackingEta: '',
    trackingStatus: '',
    returnDueDate: '',
    destructive: false
  };
  if (status === 'DELIVERED') {
    out.orderStatus = 'Delivered';
    out.stoneStatus = 'In Stock';
    out.returnDueDate = src.orderDate ? swLoupe360AddDays_(src.orderDate, 30) : '';
  } else if (status === 'SHIPPED' || status === 'IN_CUSTOMS') {
    out.orderStatus = 'On the Way';
    out.trackingEta = src.forecastedDelivery || src.expectedDelivery || '';
    out.trackingStatus = src.sourceStatusLabel;
  } else if (status === 'RETURNED') {
    out.orderStatus = 'Returned';
    out.stoneStatus = 'Returned';
    out.destructive = true;
  } else if (status === 'CANCELLED') {
    out.orderStatus = 'Cancelled';
    out.stoneStatus = 'Unavailable';
    out.destructive = true;
  } else if (status === 'NOT_AVAILABLE') {
    out.orderStatus = 'Not Available';
    out.stoneStatus = 'Unavailable';
    out.destructive = true;
  }
  return out;
}

function swLoupe360ApplyPlan_(target, plan, user) {
  var sh = target.sheet;
  var C = target.columns;
  var now = swIso_(new Date());
  var updated = 0;
  plan.updates.forEach(function (item) {
    item.changes.forEach(function (change) {
      swLoupe360SetCell_(sh, item.rowIndex, change.col, change.value);
    });
    if (C.syncAt) sh.getRange(item.rowIndex, C.syncAt).setValue(now);
    updated++;
  });

  var appended = 0;
  if (plan.appends.length) {
    var start = Math.max(sh.getLastRow() + 1, 3);
    if (sh.getMaxRows() < start + plan.appends.length - 1) {
      sh.insertRowsAfter(sh.getMaxRows(), start + plan.appends.length - 1 - sh.getMaxRows());
    }
    var rows = plan.appends.map(function (src) {
      return swLoupe360BuildAppendRow_(src, target, user, now);
    });
    try {
      if (sh.getLastRow() >= 3) {
        sh.getRange(3, 1, 1, target.lastCol).copyTo(sh.getRange(start, 1, rows.length, target.lastCol), { formatOnly: true });
      }
    } catch (_) {}
    sh.getRange(start, 1, rows.length, target.lastCol).setValues(rows);
    appended = rows.length;
  }
  return {
    updated: updated,
    appended: appended,
    skipped: plan.skipped.length + plan.conflicts.length
  };
}

function swLoupe360BuildAppendRow_(src, target, user, now) {
  var C = target.columns;
  var row = new Array(target.lastCol).fill('');
  function put(col, value) {
    if (col && value != null && value !== '') row[col - 1] = value;
  }
  var desired = swLoupe360DesiredStatus_(src);
  put(C.vendor, SW_LOUPE360_VENDOR);
  put(C.stoneType, src.stoneType || 'Lab Diamond');
  put(C.shape, src.shape);
  put(C.carat, src.carat);
  put(C.color, src.color);
  put(C.clarity, src.clarity);
  put(C.lab, src.lab);
  put(C.certNo, src.certNo);
  put(C.measurement, src.measurements);
  put(C.ratio, src.ratio);
  put(C.cut, src.cut);
  put(C.pol, src.pol);
  put(C.sym, src.sym);
  put(C.fluorIntensity, src.fluorIntensity);
  put(C.fluorColor, src.fluorColor);
  put(C.orderStatus, desired.orderStatus);
  put(C.stoneStatus, desired.stoneStatus);
  put(C.orderDate, src.orderDate);
  put(C.vendorOrderNumber, src.orderNumber);
  put(C.invoiceNumber, src.invoiceNumber);
  put(C.requestDate, new Date());
  put(C.requestedBy, user && user.email || '');
  put(C.orderedBy, src.orderedBy);
  put(C.trackingEta, desired.trackingEta);
  put(C.trackingStatus, desired.trackingStatus);
  put(C.returnDueDate, desired.returnDueDate);
  put(C.syncAt, now);
  return row;
}

function swLoupe360SetCell_(sh, row, col, value) {
  if (!col) return;
  if (typeof cd_writeBypassValidation_ === 'function') {
    cd_writeBypassValidation_(sh, row, col, value);
  } else {
    sh.getRange(row, col).setValue(value);
  }
}

function swLoupe360PreviewResponse_(temp, source, target, plan) {
  return {
    ok: true,
    syncId: temp.id,
    sourceRows: source.rows.length,
    sourceSheet: source.sheetName,
    spreadsheetUrl: target.ss.getUrl(),
    tab: target.tab,
    stats: plan.stats,
    samples: {
      appends: plan.appends.slice(0, 8).map(swLoupe360PreviewRow_),
      updates: plan.updates.slice(0, 8).map(function (row) {
        return {
          rowIndex: row.rowIndex,
          certNo: row.certNo,
          changeCount: row.changes.length,
          status: row.source.sourceStatusLabel
        };
      }),
      skipped: plan.skipped.slice(0, 8),
      conflicts: plan.conflicts.slice(0, 8)
    },
    assignmentMissing: swLoupe360CountMissingAssignmentsAfterPlan_(target, plan)
  };
}

function swLoupe360PreviewRow_(row) {
  return {
    certNo: row.certNo,
    diamond: [row.shape, row.carat, row.color, row.clarity].filter(Boolean).join(' '),
    status: row.sourceStatusLabel
  };
}

function swLoupe360CountMissingAssignments_() {
  try {
    var target = swDiamond200Target_();
    if (!target || !target.sheet || target.sheet.getLastRow() < 3) return 0;
    var sh = target.sheet;
    var hm = swDiamond200HeaderMap_(sh);
    var C = {
      root: swDiamondFind200Column_(hm, ['RootApptID', 'APPT_ID', 'Root Appt ID']),
      customerName: swDiamondFind200Column_(hm, ['Customer Name', 'Client Name', 'Customer']),
      assignedRep: swDiamondFind200Column_(hm, ['Client Advisor', 'Assigned Rep', 'Sales Rep']),
      joc: swDiamondFind200Column_(hm, ['JOC', 'Assisted Rep', 'Assistant Rep']),
      orderStatus: swDiamondFind200Column_(hm, ['Order Status', 'OrderStatus']),
      stoneStatus: swDiamondFind200Column_(hm, ['Stone Status', 'StoneStatus']),
      decision: swDiamondFind200Column_(hm, ['Stone Decision (PO, Return)', 'Stone Decision', 'StoneDecision'])
    };
    var rows = sh.getRange(3, 1, sh.getLastRow() - 2, sh.getLastColumn()).getDisplayValues();
    var count = 0;
    rows.forEach(function (row) {
      if (!swLoupe360NeedsAssignmentStatus_(row, C)) return;
      if (swLoupe360AssignmentMissing_(row, C)) {
        count++;
      }
    });
    return count;
  } catch (_) {
    return 0;
  }
}

function swLoupe360CountMissingAssignmentsAfterPlan_(target, plan) {
  var C = target.columns;
  var updatesByRow = {};
  plan.updates.forEach(function (item) {
    if (!updatesByRow[item.rowIndex]) updatesByRow[item.rowIndex] = [];
    Array.prototype.push.apply(updatesByRow[item.rowIndex], item.changes || []);
  });
  var count = 0;
  target.rows.forEach(function (rec) {
    var row = rec.row.slice();
    (updatesByRow[rec.rowIndex] || []).forEach(function (change) {
      if (change.col) row[change.col - 1] = change.value;
    });
    if (!swLoupe360NeedsAssignmentStatus_(row, C)) return;
    if (swLoupe360AssignmentMissing_(row, C)) {
      count++;
    }
  });
  plan.appends.forEach(function (src) {
    var desired = swLoupe360DesiredStatus_(src);
    if (swLoupe360SourceNeedsAssignment_(src, desired)) count++;
  });
  return count;
}

function swLoupe360IsAssignableStockRow_(row, C) {
  return swLoupe360NeedsAssignmentStatus_(row, C);
}

function swLoupe360AssignmentMissing_(row, C) {
  if (!swDiamondCell_(row, C.customerName)) return true;
  if (!swDiamondCell_(row, C.root)) return true;
  if (!swDiamondCell_(row, C.assignedRep)) return true;
  // Keep JOC optional so an older/lean 200_ tracker is not treated as missing
  // JOC on every active diamond just because the column does not exist.
  return !!(C.joc && !swDiamondCell_(row, C.joc));
}

function swLoupe360NeedsAssignmentStatus_(row, C) {
  var orderStatus = swDiamondCell_(row, C.orderStatus);
  var stoneStatus = swDiamondCell_(row, C.stoneStatus);
  var decision = swDiamondCell_(row, C.decision);
  var orderNorm = swNorm_(orderStatus);
  var stoneNorm = swNorm_(stoneStatus);
  var decisionNorm = swNorm_(decision);
  var inactive = orderNorm === 'returned' || orderNorm === 'cancelled' || orderNorm === 'not available' ||
    /returned|return in progress|unavailable|cancelled/.test(stoneNorm) ||
    decisionNorm === 'return' || decisionNorm === 'returned';
  if (inactive) return false;
  return orderNorm === 'on the way' || orderNorm === 'delivered' ||
    stoneNorm.indexOf('in stock') >= 0 ||
    /purchased|customer purchased/.test(stoneNorm) ||
    decisionNorm === 'purchase' || decisionNorm === 'purchased';
}

function swLoupe360WorkflowStatusConflict_(row, C, desired) {
  if (!desired || (!desired.orderStatus && !desired.stoneStatus)) return false;
  var orderNorm = swNorm_(swDiamondCell_(row, C.orderStatus));
  var stoneNorm = swNorm_(swDiamondCell_(row, C.stoneStatus));
  var decisionNorm = swNorm_(swDiamondCell_(row, C.decision));
  var protectedRow = /return in progress|returned|sold|customer purchased|purchased/.test(stoneNorm) ||
    orderNorm === 'returned' || orderNorm === 'cancelled' || orderNorm === 'not available' ||
    decisionNorm === 'purchase' || decisionNorm === 'purchased' || decisionNorm === 'return';
  return protectedRow && (
    (desired.orderStatus && swNorm_(desired.orderStatus) !== orderNorm) ||
    (desired.stoneStatus && swNorm_(desired.stoneStatus) !== stoneNorm)
  );
}

function swLoupe360SourceNeedsAssignment_(src, desired) {
  var status = src && src.sourceStatus;
  return status === 'SHIPPED' || status === 'IN_CUSTOMS' || status === 'DELIVERED' ||
    swNorm_(desired && desired.orderStatus) === 'on the way' ||
    swNorm_(desired && desired.orderStatus) === 'delivered';
}

function swLoupe360IsActiveSource_(src) {
  return src.sourceStatus === 'DELIVERED' || src.sourceStatus === 'SHIPPED' || src.sourceStatus === 'IN_CUSTOMS';
}

function swLoupe360CertKey_(value) {
  return swTrim_(value).toUpperCase().replace(/[^A-Z0-9]+/g, '');
}

function swLoupe360StatusLabel_(status) {
  var labels = {
    DELIVERED: 'Delivered',
    SHIPPED: 'Shipped',
    IN_CUSTOMS: 'In Customs',
    RETURNED: 'Returned',
    CANCELLED: 'Cancelled',
    NOT_AVAILABLE: 'Not Available'
  };
  return labels[status] || swLoupe360Title_(String(status || '').replace(/_/g, ' '));
}

function swLoupe360Title_(value) {
  return swTrim_(value).toLowerCase().replace(/\b[a-z]/g, function (ch) { return ch.toUpperCase(); });
}

function swLoupe360DateValue_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value;
  if (typeof value === 'number' && isFinite(value) && value > 20000) {
    return new Date(new Date(1899, 11, 30).getTime() + value * 24 * 60 * 60 * 1000);
  }
  var text = swTrim_(value);
  if (!text) return '';
  var n = Number(text);
  if (isFinite(n) && n > 20000) return new Date(new Date(1899, 11, 30).getTime() + n * 24 * 60 * 60 * 1000);
  var parsed = new Date(text);
  return isNaN(parsed.getTime()) ? '' : parsed;
}

function swLoupe360AddDays_(dateValue, days) {
  var date = swLoupe360DateValue_(dateValue);
  if (!date) return '';
  return new Date(date.getTime() + Number(days || 0) * 24 * 60 * 60 * 1000);
}

function swLoupe360DisplayValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, swTimezone_(), 'yyyy-MM-dd');
  }
  return swTrim_(value);
}
