/**
 * Diamond Viewing task adapter for Sales Workflow.
 *
 * 200_ remains the source of truth for diamond status, tracking ETA, decisions,
 * and return deadlines. Sales Workflow task payloads cache small snapshots for
 * card rendering and fast task detail loads.
 */

function swGenerateDiamondWorkflowTasks_(ss, state, ctx, rec, now, summary, visitAt) {
  if (!swDiamondIsViewingAppointment_(rec)) return;

  var dueNow = now;
  var diamond = swDiamondSnapshotForRec_(ss, ctx, rec, visitAt);
  var base = swDiamondPayloadExtra_(rec, diamond);

  swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.DIAMOND_PROPOSE, SW_OWNER_ROLES.SALES_REP, dueNow, '', now, base), summary);
  swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.DIAMOND_QUOTE, SW_OWNER_ROLES.JOC, dueNow, '', now, base), summary);

  if (diamond.counts.proposing > 0) {
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.DIAMOND_ORDER, SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN, dueNow, '', now, base), summary);
  }

  if (diamond.counts.onTheWay > 0) {
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.DIAMOND_TRACK, SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT, dueNow, '', now, base), summary);
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.DIAMOND_DELIVERY, SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN, dueNow, '', now, base), summary);
  }

  var decisionDue = visitAt && visitAt.getTime() > now.getTime() ? swDayOfDue_(visitAt) : dueNow;
  if (diamond.counts.delivered > 0) {
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.DIAMOND_DECISIONS, SW_OWNER_ROLES.JOC, decisionDue, '', now, base), summary);
  }

  if (diamond.returnRows.length) {
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.DIAMOND_RETURN, SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT, dueNow, '', now, base), summary);
  }

  if (diamond.etaIssue) {
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.DIAMOND_ETA_REP, SW_OWNER_ROLES.SALES_REP, dueNow, '', now, base), summary);
    swUpsertTask_(ss, state, swBuildTask_(ss, state, ctx, rec, SW_TASKS.DIAMOND_ETA_JOC, SW_OWNER_ROLES.JOC, dueNow, '', now, base), summary);
  }
}

function swDiamondIsViewingAppointment_(rec) {
  return /diamond\s*viewing/i.test(String(rec && rec.visitType || ''));
}

function swDiamondSnapshotForRec_(ss, ctx, rec, visitAt) {
  var rows = swDiamondRowsForRec_(ss, ctx, rec);
  var returnWindow = Number(swConfigValue_(ctx.config || [], 'SYSTEM', 'DIAMOND_RETURN_WINDOW_DAYS', '30')) || 30;
  var returnWarning = Number(swConfigValue_(ctx.config || [], 'SYSTEM', 'DIAMOND_RETURN_WARNING_DAYS', '7')) || 7;
  var now = new Date();
  var warningMs = now.getTime() + returnWarning * 24 * 60 * 60 * 1000;
  var counts = {
    total: rows.length,
    proposing: 0,
    onTheWay: 0,
    delivered: 0,
    returnDecision: 0,
    purchaseDecision: 0
  };
  var actionRows = [];
  var returnRows = [];
  var etaIssue = '';

  rows.forEach(function (row) {
    var order = swNorm_(row.orderStatus);
    var stone = swNorm_(row.stoneStatus);
    var decision = swNorm_(row.decision);
    if (order === 'proposing') counts.proposing++;
    if (order === 'on the way') counts.onTheWay++;
    if (order === 'delivered' || stone.indexOf('in stock') >= 0) counts.delivered++;
    if (decision === 'return') counts.returnDecision++;
    if (decision === 'purchase') counts.purchaseDecision++;

    row.returnDueDate = row.returnDueDate || swDiamondReturnDueDate_(row.orderDate, returnWindow);
    if (decision === 'return' && swDiamondDateValue_(row.returnDueDate) <= warningMs) returnRows.push(row);
    if (order === 'proposing' || order === 'on the way' || order === 'delivered' || decision === 'return') actionRows.push(row);

    var trackingStatus = swNorm_(row.trackingStatus);
    var etaValue = swDiamondDateValue_(row.trackingEta);
    if (!etaIssue && (/(delay|unavailable|cancel|problem|concern)/i.test(trackingStatus) ||
        (visitAt && etaValue && etaValue > visitAt.getTime()))) {
      etaIssue = swDiamondEtaIssueText_(row, visitAt);
    }
  });

  var summary = rec.dvStonesSummary || swDiamondCountsSummary_(counts);
  var proposalTarget = visitAt ? swDateKey_(swDateAddHours_(visitAt, -14 * 24)) : '';
  return {
    trackerUrl: swDiamondTrackerUrl_(),
    quotationUrl: rec.quotationUrl || '',
    tracker3dUrl: rec.tracker3dUrl || '',
    centerStoneStatus: rec.centerStoneStatus || '',
    summary: summary,
    proposalTarget: proposalTarget,
    counts: counts,
    rows: actionRows.slice(0, 25),
    returnRows: returnRows.slice(0, 25),
    etaIssue: etaIssue
  };
}

function swDiamondPayloadExtra_(rec, diamond) {
  diamond = diamond || {};
  return {
    quotationUrl: diamond.quotationUrl || rec.quotationUrl || '',
    tracker3dUrl: diamond.tracker3dUrl || rec.tracker3dUrl || '',
    diamondTrackerUrl: diamond.trackerUrl || '',
    diamondSummary: diamond.summary || '',
    diamondProposalTarget: diamond.proposalTarget || '',
    diamondActionSummary: swDiamondActionSummary_(diamond),
    diamondEtaIssue: diamond.etaIssue || '',
    manufacturingMessage: swDiamondManufacturingMessage_(rec, diamond),
    diamond: diamond
  };
}

function swDiamondRowsForRec_(ss, ctx, rec) {
  ctx = ctx || {};
  if (!ctx.diamondRowsByRoot) ctx.diamondRowsByRoot = swDiamondBuildRowsByRoot_();
  var rows = ctx.diamondRowsByRoot[swTrim_(rec.root)] || [];
  return rows.slice();
}

function swDiamondBuildRowsByRoot_() {
  var out = {};
  var target = swDiamond200Target_();
  if (!target || !target.sheet) return out;

  var sh = target.sheet;
  var lr = sh.getLastRow();
  var lc = sh.getLastColumn();
  if (lr < 3 || lc < 1) return out;
  var hm = swDiamond200HeaderMap_(sh);
  var C = swDiamond200Columns_(hm);
  if (!C.root) return out;
  var values = sh.getRange(3, 1, lr - 2, lc).getDisplayValues();
  values.forEach(function (row, i) {
    var root = swTrim_(row[C.root - 1]);
    if (!root) return;
    var rec = {
      rowIndex: i + 3,
      root: root,
      certNo: swDiamondCell_(row, C.certNo),
      vendor: swDiamondCell_(row, C.vendor),
      stoneType: swDiamondCell_(row, C.stoneType),
      shape: swDiamondCell_(row, C.shape),
      carat: swDiamondCell_(row, C.carat),
      color: swDiamondCell_(row, C.color),
      clarity: swDiamondCell_(row, C.clarity),
      lab: swDiamondCell_(row, C.lab),
      measurement: swDiamondCell_(row, C.measurement),
      ratio: swDiamondCell_(row, C.ratio),
      orderStatus: swDiamondCell_(row, C.orderStatus),
      stoneStatus: swDiamondCell_(row, C.stoneStatus),
      decision: swDiamondCell_(row, C.decision),
      orderDate: swDiamondCell_(row, C.orderDate),
      returnDueDate: swDiamondCell_(row, C.returnDueDate),
      trackingEta: swDiamondCell_(row, C.trackingEta),
      trackingStatus: swDiamondCell_(row, C.trackingStatus),
      carrier: swDiamondCell_(row, C.carrier),
      trackingNumber: swDiamondCell_(row, C.trackingNumber),
      trackingUrl: swDiamondCell_(row, C.trackingUrl),
      trackingNotes: swDiamondCell_(row, C.trackingNotes)
    };
    if (!out[root]) out[root] = [];
    out[root].push(rec);
  });
  return out;
}

function swDiamond200Target_() {
  try {
    if (typeof dp_get200Sheet_ === 'function') return dp_get200Sheet_();
  } catch (e) {
    try { Logger.log('SW_DIAMOND_200_UNAVAILABLE ' + e.message); } catch (_) {}
  }
  return null;
}

function swDiamond200HeaderMap_(sh) {
  if (typeof dp_headerMapFor200_ === 'function') return dp_headerMapFor200_(sh);
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  return { byExact: swHeaderMapFromArray_(headers), byNorm: swHeaderMapFromArray_(headers) };
}

function swDiamond200Columns_(hm) {
  return {
    root: swDiamondFind200Column_(hm, ['RootApptID', 'APPT_ID', 'Root Appt ID']),
    certNo: swDiamondFind200Column_(hm, ['Certificate No', 'Cert #', 'Cert No', 'Certificate #']),
    vendor: swDiamondFind200Column_(hm, ['Vendor']),
    stoneType: swDiamondFind200Column_(hm, ['Stone Type', 'StoneType']),
    shape: swDiamondFind200Column_(hm, ['Shape']),
    carat: swDiamondFind200Column_(hm, ['Carat']),
    color: swDiamondFind200Column_(hm, ['Color']),
    clarity: swDiamondFind200Column_(hm, ['Clarity']),
    lab: swDiamondFind200Column_(hm, ['LAB', 'Lab', 'Grading Lab']),
    measurement: swDiamondFind200Column_(hm, ['Measurements', 'Measurement', 'Meas.', 'Meas']),
    ratio: swDiamondFind200Column_(hm, ['L/W Ratio', 'L-W Ratio', 'LW Ratio', 'Ratio']),
    orderStatus: swDiamondFind200Column_(hm, ['Order Status', 'OrderStatus']),
    stoneStatus: swDiamondFind200Column_(hm, ['Stone Status', 'StoneStatus']),
    decision: swDiamondFind200Column_(hm, ['Stone Decision (PO, Return)', 'Stone Decision', 'StoneDecision']),
    orderDate: swDiamondFind200Column_(hm, ['Purchased / Ordered Date', 'Purchased/Ordered Date', 'PurchasedOrderedDate']),
    returnDueDate: swDiamondFind200Column_(hm, ['Return DUE DATE', 'Return Due Date', 'Return Due']),
    trackingEta: swDiamondFind200Column_(hm, ['Tracking ETA', 'Tracking ETA Date', 'ETA Date', 'ETA']),
    trackingStatus: swDiamondFind200Column_(hm, ['Tracking Status', 'ETA Status', 'Shipment Status']),
    carrier: swDiamondFind200Column_(hm, ['Carrier', 'Shipping Carrier']),
    trackingNumber: swDiamondFind200Column_(hm, ['Tracking Number', 'Tracking #', 'Tracking No']),
    trackingUrl: swDiamondFind200Column_(hm, ['Tracking URL', 'Tracking Link']),
    trackingNotes: swDiamondFind200Column_(hm, ['Tracking Notes', 'ETA Notes', 'Shipping Notes'])
  };
}

function swDiamondFind200Column_(hm, aliases) {
  if (!hm) return 0;
  if (typeof dp_findHeaderIndex_ === 'function') {
    try {
      var found = dp_findHeaderIndex_(hm, aliases, false);
      return found > 0 ? found : 0;
    } catch (_) {}
  }
  var maps = [hm.byExact || hm, hm.byNorm || {}];
  for (var m = 0; m < maps.length; m++) {
    for (var i = 0; i < aliases.length; i++) {
      var raw = aliases[i];
      if (maps[m][raw] != null) return maps[m][raw] + 1;
      var key = swHeaderKey_(raw);
      if (maps[m][key] != null) return maps[m][key] + 1;
    }
  }
  return 0;
}

function swDiamondEnsure200Column_(sh, header) {
  var headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getDisplayValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (swHeaderKey_(headers[i]) === swHeaderKey_(header)) return i + 1;
  }
  var col = sh.getLastColumn() + 1;
  sh.getRange(1, col).setValue(header);
  return col;
}

function swDiamondCell_(row, col) {
  return col ? swTrim_(row[col - 1]) : '';
}

function swDiamondTrackerUrl_() {
  var target = swDiamond200Target_();
  return target && target.ss ? target.ss.getUrl() : '';
}

function swDiamondActionSummary_(diamond) {
  diamond = diamond || {};
  if (diamond.etaIssue) return diamond.etaIssue;
  if (diamond.returnRows && diamond.returnRows.length) return diamond.returnRows.length + ' return row(s) due soon or overdue.';
  var c = diamond.counts || {};
  return [
    'Proposing: ' + (c.proposing || 0),
    'On the Way: ' + (c.onTheWay || 0),
    'Delivered: ' + (c.delivered || 0),
    'Return: ' + (c.returnDecision || 0),
    'Purchase: ' + (c.purchaseDecision || 0)
  ].join(' | ');
}

function swDiamondCountsSummary_(counts) {
  counts = counts || {};
  return 'Proposed: ' + (counts.proposing || 0) +
    ' | On the Way: ' + (counts.onTheWay || 0) +
    ' | Delivered: ' + (counts.delivered || 0) +
    ' | Total: ' + (counts.total || 0);
}

function swDiamondEtaIssueText_(row, visitAt) {
  var parts = [];
  if (row.trackingStatus) parts.push('Status: ' + row.trackingStatus);
  if (row.trackingEta) parts.push('ETA: ' + row.trackingEta);
  if (visitAt) parts.push('Appointment: ' + swDateKey_(visitAt));
  if (row.certNo) parts.push('Cert: ' + row.certNo);
  return parts.join(' | ') || 'Diamond ETA/status needs review.';
}

function swDiamondManufacturingMessage_(rec, diamond) {
  var rows = ((diamond && diamond.rows) || []).filter(function (row) {
    return swNorm_(row.decision) === 'purchase' || swNorm_(row.stoneStatus).indexOf('customer purchased') >= 0;
  });
  if (!rows.length) rows = ((diamond && diamond.rows) || []).filter(function (row) {
    return swNorm_(row.orderStatus) === 'delivered' || swNorm_(row.stoneStatus).indexOf('in stock') >= 0;
  }).slice(0, 3);
  var lines = [
    'Manufacturing, confirmed diamond dimensions for ' + (rec.name || 'customer') + (rec.so ? ' / SO ' + rec.so : '') + ':'
  ];
  rows.forEach(function (row) {
    lines.push('- ' + [
      row.shape,
      row.carat ? row.carat + 'ct' : '',
      row.color,
      row.clarity,
      row.measurement ? 'Measurements ' + row.measurement : '',
      row.ratio ? 'L/W ' + row.ratio : '',
      row.certNo ? 'Cert ' + row.certNo : ''
    ].filter(Boolean).join(' | '));
  });
  lines.push('Please confirm the CAD/manufacturing dimensions match the latest 3D tracker details before production.');
  return lines.join('\n');
}

function swDiamondReturnDueDate_(orderDate, days) {
  var t = swDiamondDateValue_(orderDate);
  if (!t) return '';
  var d = new Date(t + (Number(days) || 30) * 24 * 60 * 60 * 1000);
  return swDateKey_(d);
}

function swDiamondDateValue_(value) {
  if (!value) return 0;
  var d = new Date(value);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function swDiamondHandleTaskCompletion_(ss, task, data, user) {
  data = data || {};
  if (task.taskType === SW_TASKS.DIAMOND_ORDER) return swDiamondCompleteOrder_(task, data, user);
  if (task.taskType === SW_TASKS.DIAMOND_TRACK) return swDiamondCompleteTracking_(task, data, user);
  if (task.taskType === SW_TASKS.DIAMOND_DELIVERY) return swDiamondCompleteDelivery_(task, data);
  if (task.taskType === SW_TASKS.DIAMOND_DECISIONS) return swDiamondCompleteDecisions_(task, data);
  return null;
}

function swDiamondCompleteOrder_(task, data, user) {
  var decisions = data.diamondOrderDecisions || [];
  var items = decisions.filter(function (item) { return item && item.decision; }).map(function (item) {
    return {
      rowIndex: Number(item.rowIndex),
      rootApptId: task.root,
      decision: item.decision,
      orderedBy: data.orderedBy || (user && user.email) || '',
      orderedDate: data.orderedDate || ''
    };
  });
  if (!items.length) return { skipped: true, reason: 'No order decisions supplied.' };
  if (typeof dp_submitOrderApprovals !== 'function') throw new Error('Diamond order approval function is not available.');
  return dp_submitOrderApprovals({
    applyDefaultsToAll: true,
    defaultOrderedBy: data.orderedBy || (user && user.email) || '',
    defaultOrderDate: data.orderedDate || '',
    items: items
  });
}

function swDiamondCompleteTracking_(task, data, user) {
  var target = swDiamond200Target_();
  if (!target || !target.sheet) throw new Error('Diamond tracking sheet is unavailable.');
  var sh = target.sheet;
  var rows = swDeepValue_(swParseJson_(task.payloadJson, {}), ['extra', 'diamond', 'rows']) || [];
  var onTheWay = rows.filter(function (row) { return swNorm_(row.orderStatus) === 'on the way'; });
  if (!onTheWay.length) return { skipped: true, reason: 'No on-the-way stones found.' };

  var cEta = swDiamondEnsure200Column_(sh, 'Tracking ETA');
  var cStatus = swDiamondEnsure200Column_(sh, 'Tracking Status');
  var cCarrier = swDiamondEnsure200Column_(sh, 'Carrier');
  var cNumber = swDiamondEnsure200Column_(sh, 'Tracking Number');
  var cUrl = swDiamondEnsure200Column_(sh, 'Tracking URL');
  var cNotes = swDiamondEnsure200Column_(sh, 'Tracking Notes');
  var cChecked = swDiamondEnsure200Column_(sh, 'Last Tracking Check At');
  var now = swIso_(new Date());
  onTheWay.forEach(function (row) {
    var r = Number(row.rowIndex);
    if (!(r >= 3)) return;
    sh.getRange(r, cEta).setValue(data.trackingEta || '');
    sh.getRange(r, cStatus).setValue(data.trackingStatus || '');
    sh.getRange(r, cCarrier).setValue(data.carrier || '');
    sh.getRange(r, cNumber).setValue(data.trackingNumber || '');
    sh.getRange(r, cUrl).setValue(data.trackingUrl || '');
    sh.getRange(r, cNotes).setValue(data.trackingNotes || '');
    sh.getRange(r, cChecked).setValue(now + (user && user.email ? ' by ' + user.email : ''));
  });
  return { ok: true, updatedRows: onTheWay.map(function (row) { return row.rowIndex; }) };
}

function swDiamondCompleteDelivery_(task, data) {
  var rows = swDeepValue_(swParseJson_(task.payloadJson, {}), ['extra', 'diamond', 'rows']) || [];
  var selected = rows.filter(function (row) { return swNorm_(row.orderStatus) === 'on the way'; }).map(function (row) {
    return { rowIndex: Number(row.rowIndex), rootApptId: task.root, memoDate: data.memoDate || '', selected: true };
  });
  if (!selected.length) return { skipped: true, reason: 'No on-the-way stones found.' };
  if (typeof dp_submitConfirmDelivery !== 'function') throw new Error('Diamond delivery confirmation function is not available.');
  return dp_submitConfirmDelivery({
    applyDefaultToAll: true,
    defaultMemoDate: data.memoDate || '',
    items: selected
  });
}

function swDiamondCompleteDecisions_(task, data) {
  var decisions = data.diamondDecisions || [];
  var items = decisions.filter(function (item) { return item && item.decision; }).map(function (item) {
    return {
      rowIndex: Number(item.rowIndex),
      rootApptId: task.root,
      decision: item.decision,
      hold: !!item.hold
    };
  });
  if (!items.length) return { skipped: true, reason: 'No diamond decisions supplied.' };
  if (typeof dp_submitStoneDecisions !== 'function') throw new Error('Diamond decision function is not available.');
  return dp_submitStoneDecisions({ items: items });
}

function sw_refreshDiamondQuoteFromTracking(authToken, taskId) {
  return swDiamondQuoteAction_(authToken, taskId, 'diamonds');
}

function sw_refreshDiamondQuoteFrom3D(authToken, taskId) {
  return swDiamondQuoteAction_(authToken, taskId, 'settings');
}

function sw_refreshDiamondQuoteAll(authToken, taskId) {
  return swDiamondQuoteAction_(authToken, taskId, 'all');
}

function swDiamondQuoteAction_(authToken, taskId, mode) {
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var user = swAuthUserForApi_(ss, authToken);
  var task = swGetTaskById_(ss, taskId);
  if (!task) throw new Error('Task not found: ' + taskId);
  if (!swCanViewTask_(task, user)) throw new Error('You do not have access to this task.');
  var payload = swParseJson_(task.payloadJson, {});
  var appt = payload.appointment || {};
  var latestDiamond = swDiamondSnapshotForRec_(ss, swBuildIdentityContext_(ss, true), {
    root: task.root || appt.root,
    appt: task.appt || appt.appt,
    name: task.customerName || appt.customerName,
    visitDate: task.visitDate || appt.visitDate,
    visitTime: task.visitTime || appt.visitTime,
    quotationUrl: swDeepValue_(payload, ['extra', 'quotationUrl']) || appt.quotationUrl || '',
    tracker3dUrl: swDeepValue_(payload, ['extra', 'tracker3dUrl']) || appt.tracker3dUrl || '',
    dvStonesSummary: swDeepValue_(payload, ['extra', 'diamondSummary']) || appt.dvStonesSummary || '',
    centerStoneStatus: appt.centerStoneStatus || ''
  }, null);
  payload.extra = payload.extra || {};
  payload.extra.diamond = latestDiamond;
  var quotationUrl = swDeepValue_(payload, ['extra', 'quotationUrl']) || swDeepValue_(payload, ['appointment', 'quotationUrl']);
  if (!quotationUrl) throw new Error('Quotation URL is missing.');

  var out = { ok: true, mode: mode };
  if (mode === 'diamonds' || mode === 'all') out.diamonds = swDiamondWriteQuoteDiamonds_(quotationUrl, payload);
  if (mode === 'settings' || mode === 'all') out.settings = swDiamondWriteQuoteSettings_(quotationUrl, payload);
  swAppendTaskLog_(ss, 'DIAMOND_QUOTE_REFRESH', task, user, task.currentOwner, task.currentOwner, out);
  return out;
}

function swDiamondWriteQuoteDiamonds_(quotationUrl, payload) {
  if (typeof uq_writeDiamondsToQuote_ !== 'function') throw new Error('Quotation diamond writer is not available.');
  var rows = swDeepValue_(payload, ['extra', 'diamond', 'rows']) || [];
  var records = rows.filter(function (row) { return row.certNo; }).map(function (row) {
    return {
      certNo: row.certNo,
      vendor: row.vendor || '',
      stoneType: row.stoneType || '',
      shape: row.shape || '',
      carat: row.carat || '',
      color: row.color || '',
      clarity: row.clarity || '',
      lab: row.lab || '',
      ratio: row.ratio || '',
      measurement: row.measurement || '',
      orderStatus: row.orderStatus || '',
      stoneStatus: row.stoneStatus || '',
      customerName: swDeepValue_(payload, ['appointment', 'customerName']) || '',
      apptTimeDate: [swDeepValue_(payload, ['appointment', 'visitDate']), swDeepValue_(payload, ['appointment', 'visitTime'])].filter(Boolean).join(' '),
      assignedRep: swDeepValue_(payload, ['appointment', 'assignedRep']) || '',
      company: swDeepValue_(payload, ['appointment', 'brand']) || ''
    };
  });
  if (!records.length) throw new Error('No diamond rows with certificate numbers are available for quotation refresh.');
  return uq_writeDiamondsToQuote_(quotationUrl, records);
}

function swDiamondWriteQuoteSettings_(quotationUrl, payload) {
  if (typeof uq_writeSettingsToQuote_ !== 'function') throw new Error('Quotation setting writer is not available.');
  var trackerUrl = swDeepValue_(payload, ['extra', 'tracker3dUrl']);
  if (!trackerUrl) throw new Error('3D Tracker URL is missing.');
  var record = swDiamondLatest3DSettingRecord_(trackerUrl);
  return uq_writeSettingsToQuote_(quotationUrl, [record]);
}

function swDiamondLatest3DSettingRecord_(trackerUrl) {
  var fileId = (typeof uq_extractFileId_ === 'function') ? uq_extractFileId_(trackerUrl) : swDiamondExtractFileId_(trackerUrl);
  if (!fileId) throw new Error('3D Tracker URL not recognized.');
  var ss = SpreadsheetApp.openById(fileId);
  var shLog = ss.getSheetByName('Log') || ss.getSheetByName('3D Log') || ss.getSheetByName('3D Revision Log');
  if (!shLog || shLog.getLastRow() < 2) throw new Error('Tracker Log tab has no usable rows.');
  var lr = shLog.getLastRow();
  var lc = shLog.getLastColumn();
  var headers = shLog.getRange(1, 1, 1, lc).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var row = shLog.getRange(lr, 1, 1, lc).getDisplayValues()[0];
  function get(names) {
    var idx = swPickIndex_(H, names);
    return idx >= 0 ? row[idx] : '';
  }
  var style = get(['Ring Style', 'RingStyle']);
  var metal = get(['Metal', 'Metal Type', 'Metal (Type)']);
  var band = swTrim_(get(['Band Width (mm)', 'BandWidthMM', 'Band Width'])).replace(/\s*mm\s*$/i, '');
  var size = get(['US Size', 'USSize']);
  var product = get(['Product', 'Setting Name', 'Design Name']);
  return {
    product: String(product || (style ? style + ' Setting' : '')),
    styleDetail: String(style || ''),
    metal: String(metal || ''),
    bandWidth: band,
    ringSize: String(size || ''),
    freeUpgrade: '',
    onlineRetailerPrice: '',
    brilliantEarthPriceAfterTax: '',
    vvsPrice: '',
    yourSavings: '',
    link: ''
  };
}

function swDiamondExtractFileId_(url) {
  var m = /[-\w]{25,}/.exec(String(url || ''));
  return m ? m[0] : '';
}
