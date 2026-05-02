var AR_SHEET_NAME = '00_Master Appointments';

var AR_CFG = {
  NEXT_2ND_VIEW_KEYWORDS: ['diamond viewing ready'],
  FOLLOW_UP_KEYWORDS:     ['needs follow-up']
};

// ── Open dialog ───────────────────────────────────────────────
function openApptReportDialog() {
  var html = HtmlService.createTemplateFromFile('dlg_appt_report_v1')
    .evaluate()
    .setWidth(1280).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, 'Appointment Sales Report');
}

// ── Column map helpers ────────────────────────────────────────
function ar_buildColMap_(headers) {
  var map = {};
  headers.forEach(function(h, i) {
    var k = String(h || '').trim();
    if (k) map[k] = i;
  });
  return map;
}

function ar_getCol_(colMap) {
  var names = Array.prototype.slice.call(arguments, 1);
  for (var i = 0; i < names.length; i++) {
    if (colMap.hasOwnProperty(names[i])) return colMap[names[i]];
  }
  return -1;
}

// ── Get sheet ─────────────────────────────────────────────────
function ar_getSheet_() {
  var masterId = (PropertiesService.getScriptProperties().getProperty('MASTER_FILE_ID') || '').trim();
  var ss = masterId
    ? SpreadsheetApp.openById(masterId)
    : SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(AR_SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + AR_SHEET_NAME + '" not found.');
  return sh;
}

// ── Filter options for dropdowns ─────────────────────────────
function appt_getFilterOptions() {
  try {
    var sh      = ar_getSheet_();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, brands: [], reps: [], visitTypes: [] };

    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var colMap  = ar_buildColMap_(headers);
    var rows    = lastRow - 1;

    var uniq = function() {
      var names  = Array.prototype.slice.call(arguments);
      var colIdx = ar_getCol_.apply(null, [colMap].concat(names));
      if (colIdx < 0) return [];
      var vals = sh.getRange(2, colIdx + 1, rows, 1).getValues();
      var seen = {}, out = [];
      vals.forEach(function(r) {
        var v = String(r[0] || '').trim();
        if (v && !seen[v]) { seen[v] = true; out.push(v); }
      });
      return out.sort();
    };

    return {
      ok:         true,
      brands:     uniq('Brand'),
      reps:       uniq('Assigned Rep'),
      visitTypes: uniq('Visit Type')
    };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Main report data ──────────────────────────────────────────
function appt_getReportData(filters) {
  try {
    filters = filters || {};
    var sh      = ar_getSheet_();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, rows: [], summary: ar_emptySummary_() };

    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var colMap  = ar_buildColMap_(headers);

    var C = {
      apptId:            ar_getCol_(colMap, 'APPT_ID'),
      rootApptId:        ar_getCol_(colMap, 'RootApptID', 'Root Appt ID'),
      brand:             ar_getCol_(colMap, 'Brand'),
      visitNum:          ar_getCol_(colMap, 'Visit #', 'Visit#'),
      customerName:      ar_getCol_(colMap, 'Customer Name'),
      assignedRep:       ar_getCol_(colMap, 'Assigned Rep'),
      salesStage:        ar_getCol_(colMap, 'Sales Stage'),
      conversionStatus:  ar_getCol_(colMap, 'Conversion_Status', 'Conversion Status'),
      centerStoneStatus: ar_getCol_(colMap, 'Center Stone Order Status'),
      ackStatus:         ar_getCol_(colMap, 'Ack Status'),
      status:            ar_getCol_(colMap, 'Status'),
      visitType:         ar_getCol_(colMap, 'Visit Type'),
      visitDate:         ar_getCol_(colMap, 'Visit Date'),
      visitTime:         ar_getCol_(colMap, 'Visit Time'),
      location:          ar_getCol_(colMap, 'Location'),
      soNum:             ar_getCol_(colMap, 'SO#', 'SO'),
      orderTotal:        ar_getCol_(colMap, 'Order Total'),
      paidToDate:        ar_getCol_(colMap, 'Paid-to-Date', 'Paid to Date'),
      remainingBalance:  ar_getCol_(colMap, 'Remaining Balance'),
      cashInGross:       ar_getCol_(colMap, 'Cash-in (Gross)', 'Cash In Gross'),
      lastPaymentDate:   ar_getCol_(colMap, 'Last Payment Date'),
      nextSteps:         ar_getCol_(colMap, 'Next Steps')
    };

    var allData = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

    var dateFrom = filters.dateFrom ? new Date(filters.dateFrom + 'T00:00:00') : null;
    var dateTo   = filters.dateTo   ? new Date(filters.dateTo   + 'T23:59:59') : null;

    var isActive = function(arr) { return arr && arr.length > 0; };
    var matchArr = function(arr, val) {
      var v = String(val || '').trim().toLowerCase();
      return arr.some(function(f) { return String(f).trim().toLowerCase() === v; });
    };

    var tz     = Session.getScriptTimeZone() || 'America/Los_Angeles';
    var fmtDate = function(v) {
      if (!v) return '';
      var d = v instanceof Date ? v : new Date(v);
      return isNaN(d) ? String(v) : Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    };
    var fmtS = function(v) { return String(v == null ? '' : v).trim(); };
    var fmtN = function(v) { return (v === '' || v == null) ? 0 : (Number(v) || 0); };
    var get  = function(row, idx) { return idx >= 0 ? row[idx] : ''; };

    var rows = [];

    for (var i = 0; i < allData.length; i++) {
      var row = allData[i];

      // Date filter
      if (dateFrom || dateTo) {
        var rawDate = get(row, C.visitDate);
        var vd = rawDate instanceof Date ? rawDate : (rawDate ? new Date(rawDate) : null);
        if (!vd || isNaN(vd)) continue;
        if (dateFrom && vd < dateFrom) continue;
        if (dateTo   && vd > dateTo)   continue;
      }

      // Field filters
      if (isActive(filters.brands)     && !matchArr(filters.brands,     get(row, C.brand)))      continue;
      if (isActive(filters.visitTypes) && !matchArr(filters.visitTypes, get(row, C.visitType)))  continue;
      if (isActive(filters.reps)       && !matchArr(filters.reps,       get(row, C.assignedRep))) continue;

      var vtStr = fmtS(get(row, C.visitType));
      var vtLc  = vtStr.toLowerCase();
      var stStr = fmtS(get(row, C.status));
      var stLc  = stStr.toLowerCase();

      // Visit type classification
      var isFirstTime  = vtLc === 'appointment';
      var isSecondView = vtLc === 'diamond viewing';
      var isWalkIn     = vtLc.indexOf('walk') !== -1;
      var isOnline     = vtLc === 'online customer';
      var isFunnel     = isFirstTime || isSecondView; // only these two in conversion funnel

      // Attendance classification
      var isNoShow    = stLc.indexOf('no show') !== -1 || stLc.indexOf('no-show') !== -1 || stLc === 'absent';
      var isCompleted = !isNoShow && (
        stLc.indexOf('complet') !== -1 || stLc.indexOf('done')    !== -1 ||
        stLc.indexOf('attend')  !== -1 || stLc.indexOf('visited') !== -1
      );
      var isScheduled = !isNoShow && !isCompleted;

      // Revenue
      var orderTotal  = fmtN(get(row, C.orderTotal));
      var remaining   = fmtN(get(row, C.remainingBalance));
      var cashIn      = fmtN(get(row, C.cashInGross));
      var hasDeposit  = cashIn > 0;
      var isClosed    = orderTotal > 0.005 && remaining <= 0.005; // Remaining Balance = 0 → closed

      rows.push({
        rowIdx:            i + 2,
        apptId:            fmtS(get(row, C.apptId)),
        rootApptId:        fmtS(get(row, C.rootApptId)),
        brand:             fmtS(get(row, C.brand)),
        visitNum:          fmtN(get(row, C.visitNum)),
        customerName:      fmtS(get(row, C.customerName)),
        assignedRep:       fmtS(get(row, C.assignedRep)),
        salesStage:        fmtS(get(row, C.salesStage)),
        conversionStatus:  fmtS(get(row, C.conversionStatus)),
        centerStoneStatus: fmtS(get(row, C.centerStoneStatus)),
        ackStatus:         fmtS(get(row, C.ackStatus)),
        status:            stStr,
        visitType:         vtStr,
        visitDate:         fmtDate(get(row, C.visitDate)),
        visitTime:         fmtS(get(row, C.visitTime)),
        location:          fmtS(get(row, C.location)),
        soNum:             fmtS(get(row, C.soNum)),
        orderTotal:        orderTotal,
        paidToDate:        fmtN(get(row, C.paidToDate)),
        remainingBalance:  remaining,
        cashInGross:       cashIn,
        lastPaymentDate:   fmtDate(get(row, C.lastPaymentDate)),
        nextSteps:         fmtS(get(row, C.nextSteps)),
        isFirstTime:       isFirstTime,
        isSecondView:      isSecondView,
        isWalkIn:          isWalkIn,
        isOnline:          isOnline,
        isFunnel:          isFunnel,
        isNoShow:          isNoShow,
        isCompleted:       isCompleted,
        isScheduled:       isScheduled,
        hasDeposit:        hasDeposit,
        isClosed:          isClosed
      });
    }

    // ── Aggregation ─────────────────────────────────────────
    var cnt = function(arr, fn) { return arr.filter(fn).length; };
    var sum = function(arr, key) { return arr.reduce(function(a, r) { return a + (r[key] || 0); }, 0); };

    var first  = rows.filter(function(r) { return r.isFirstTime; });
    var second = rows.filter(function(r) { return r.isSecondView; });

    // Activity
    var totalScheduled   = rows.length;
    var totalCompleted   = cnt(rows, function(r) { return r.isCompleted; });
    var totalNoShow      = cnt(rows, function(r) { return r.isNoShow; });
    var firstScheduled   = first.length;
    var firstCompleted   = cnt(first,  function(r) { return r.isCompleted; });
    var firstNoShow      = cnt(first,  function(r) { return r.isNoShow; });
    var secondScheduled  = second.length;
    var secondCompleted  = cnt(second, function(r) { return r.isCompleted; });
    var secondNoShow     = cnt(second, function(r) { return r.isNoShow; });
    var walkInCount      = cnt(rows,   function(r) { return r.isWalkIn; });
    var onlineCount      = cnt(rows,   function(r) { return r.isOnline; });

    // Conversions
    var firstWithDeposit = cnt(first,  function(r) { return r.isCompleted && r.hasDeposit; });
    var secondClosed     = cnt(second, function(r) { return r.isCompleted && r.isClosed; });

    // Revenue: deduplicate closed orders by SO#
    var closedOrders    = {};
    var totalClosedValue = 0;
    rows.filter(function(r) { return r.isClosed && r.soNum; }).forEach(function(r) {
      if (!closedOrders[r.soNum]) {
        closedOrders[r.soNum] = r.orderTotal;
        totalClosedValue += r.orderTotal;
      }
    });
    var totalDeposits = sum(rows, 'cashInGross');

    // Pipeline: "Deposit awaiting 2nd view"
    // Root Appt IDs that have a deposit but no Diamond Viewing row
    var rootsWithDeposit     = {};
    var rootsWithSecondView  = {};
    rows.forEach(function(r) {
      var key = r.rootApptId || r.apptId;
      if (!key) return;
      if (r.hasDeposit)  rootsWithDeposit[key]    = true;
      if (r.isSecondView) rootsWithSecondView[key] = true;
    });
    var depositAwaitingSecondView = Object.keys(rootsWithDeposit).filter(function(k) {
      return !rootsWithSecondView[k];
    }).length;

    
    var nextSecondView = cnt(rows, function(r) {
      var v = r.centerStoneStatus.toLowerCase();
      return AR_CFG.NEXT_2ND_VIEW_KEYWORDS.some(function(k) { return v.indexOf(k) !== -1; });
    });
    var needsFollowUp = cnt(rows, function(r) {
      var v = r.ackStatus.toLowerCase();
      return AR_CFG.FOLLOW_UP_KEYWORDS.some(function(k) { return v.indexOf(k) !== -1; });
    });

    // KPIs
    var pct = function(n, d) { return d > 0 ? n / d : 0; };
    var showRate        = pct(totalCompleted,   totalScheduled);
    var firstConvRate   = pct(firstWithDeposit, firstCompleted);
    var secondCloseRate = pct(secondClosed,      secondCompleted);

    return {
      ok:   true,
      rows: rows,
      summary: {
        totalScheduled:            totalScheduled,
        totalCompleted:            totalCompleted,
        totalNoShow:               totalNoShow,
        firstScheduled:            firstScheduled,
        firstCompleted:            firstCompleted,
        firstNoShow:               firstNoShow,
        secondScheduled:           secondScheduled,
        secondCompleted:           secondCompleted,
        secondNoShow:              secondNoShow,
        walkInCount:               walkInCount,
        onlineCount:               onlineCount,
        firstWithDeposit:          firstWithDeposit,
        secondClosed:              secondClosed,
        totalDeposits:             totalDeposits,
        totalClosedValue:          totalClosedValue,
        closedOrderCount:          Object.keys(closedOrders).length,
        depositAwaitingSecondView: depositAwaitingSecondView,
        nextSecondView:            nextSecondView,
        needsFollowUp:             needsFollowUp,
        showRate:                  showRate,
        firstConvRate:             firstConvRate,
        secondCloseRate:           secondCloseRate
      }
    };

  } catch(e) {
    Logger.log('appt_getReportData error: ' + e.stack);
    return { ok: false, error: e.message };
  }
}

function ar_emptySummary_() {
  return {
    totalScheduled:0, totalCompleted:0, totalNoShow:0,
    firstScheduled:0, firstCompleted:0, firstNoShow:0,
    secondScheduled:0, secondCompleted:0, secondNoShow:0,
    walkInCount:0, onlineCount:0,
    firstWithDeposit:0, secondClosed:0,
    totalDeposits:0, totalClosedValue:0, closedOrderCount:0,
    depositAwaitingSecondView:0,
    nextSecondView:0, needsFollowUp:0,
    showRate:0, firstConvRate:0, secondCloseRate:0
  };
}

// ── Briefing: save & load by date key ────────────────────────
// ── Briefing: lưu và đọc từ sheet 00_Briefing_Log ────────────
var BRIEFING_SHEET = '00_Briefing_Log';
var BRIEFING_COLS  = ['Date', 'No-Show Reasons', 'Objections', 'What Worked', 'Blockers', 'Saved By', 'Last Updated'];

function ar_getBriefingSheet_() {
  var masterId = (PropertiesService.getScriptProperties().getProperty('MASTER_FILE_ID') || '').trim();
  var ss = masterId ? SpreadsheetApp.openById(masterId) : SpreadsheetApp.getActive();

  var sh = ss.getSheetByName(BRIEFING_SHEET);

  // Tự tạo sheet + header nếu chưa có
  if (!sh) {
    sh = ss.insertSheet(BRIEFING_SHEET);
    var hdr = sh.getRange(1, 1, 1, BRIEFING_COLS.length);
    hdr.setValues([BRIEFING_COLS]);
    hdr.setFontWeight('bold');
    hdr.setBackground('#fef08a');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 110);
    sh.setColumnWidth(2, 200);
    sh.setColumnWidth(3, 200);
    sh.setColumnWidth(4, 200);
    sh.setColumnWidth(5, 200);
    sh.setColumnWidth(6, 130);
    sh.setColumnWidth(7, 150);
  }
  return sh;
}

function appt_saveBriefing(payload) {
  try {
    var sh      = ar_getBriefingSheet_();
    var dateKey = payload.dateKey || 'LATEST';
    var data    = payload.data || {};
    var savedBy = Session.getActiveUser().getEmail() || 'unknown';
    var now     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    var lastRow = sh.getLastRow();
    var values  = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 1).getValues() : [];

    // Tìm row đã có ngày này chưa
    var targetRow = -1;
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === dateKey) {
        targetRow = i + 2; // +2 vì bắt đầu từ row 2 (row 1 là header)
        break;
      }
    }

    var rowData = [dateKey, data.noshow || '', data.objection || '', data.worked || '', data.blockers || '', savedBy, now];

    if (targetRow > 0) {
      // Cập nhật row đã có
      sh.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      // Thêm row mới
      sh.appendRow(rowData);
    }

    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function appt_loadBriefing(dateKey) {
  try {
    var sh      = ar_getBriefingSheet_();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, data: null };

    var key    = dateKey || 'LATEST';
    var values = sh.getRange(2, 1, lastRow - 1, BRIEFING_COLS.length).getValues();

    // Tìm row khớp dateKey, ưu tiên row cuối nếu có nhiều bản
    var found = null;
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === key) {
        found = values[i];
      }
    }

    if (!found) return { ok: true, data: null };

    return {
      ok: true,
      data: {
        noshow:    found[1] || '',
        objection: found[2] || '',
        worked:    found[3] || '',
        blockers:  found[4] || ''
      }
    };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}
