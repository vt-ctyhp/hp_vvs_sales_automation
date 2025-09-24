/**
 * Diamonds_ConfirmDelivery_v1.gs — Phase 3 (Dialog C: Confirm Delivery)
 *
 * Depends on Phase 1/2 shared helpers in Diamonds_v1.gs / Diamonds_OrderApprove_v1.gs:
 *   dp_get200Sheet_(), dp_headerMapFor200_(), dp_headerMap_()
 *   dp_findHeaderIndex_(), dp_setCellByHeader_(), dp_aliases200_, dp_aliases100_
 *   dp_getCurrentUserEmail_(), dp_norm_(), dp_computeCountsForAppointment_()
 *   dp_find100RowByRootApptId_(), dp_buildSnapshotFrom200_()
 *
 * This module adds:
 *   - dp_openConfirmDeliveryDiamonds()            // menu opener
 *   - dp_bootstrapForConfirmDialog_()             // returns all "On the Way" stones (not In Stock)
 *   - dp_submitConfirmDelivery(payload)           // writes Stone Status, Memo/Invoice Date, Return DUE DATE
 *   - dp_refresh100QuickRef_(rootApptId, counts)  // refreshes summary/JSON only (no CSOS change)
 *   - alias extensions for "Memo/ Invoice Date" & "Return DUE DATE"
 */

// -------------------------------------------------------------
// MENU OPENER
// -------------------------------------------------------------
function dp_openConfirmDeliveryDiamonds() {
  var ui = SpreadsheetApp.getUi();
  var bootstrap;
  try {
    bootstrap = dp_bootstrapForConfirmDialog_();
  } catch (e) {
    ui.alert('💎 Diamonds — Confirm Delivery', (e && e.message ? e.message : String(e)), ui.ButtonSet.OK);
    return;
  }
  var t = HtmlService.createTemplateFromFile('dlg_confirm_delivery_v1');
  t.bootstrap = bootstrap; // inject payload
  var html = t.evaluate().setWidth(1250).setHeight(720);
  ui.showModalDialog(html, '💎 Confirm Delivery');
}

// -------------------------------------------------------------
// BOOTSTRAP (load all "On the Way" stones not yet In Stock)
// -------------------------------------------------------------
function dp_bootstrapForConfirmDialog_() {
  var target = dp_get200Sheet_();
  var sh200 = target.sheet;
  var hm200 = dp_headerMapFor200_(sh200);
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

  // Columns we need
  var cOrder     = dp_findHeaderIndex_(hm200, dp_aliases200_['Order Status'], true);
  var cStoneStat = dp_findHeaderIndex_(hm200, dp_aliases200_['Stone Status'], true);
  var cMemo      = dp_findHeaderIndex_(hm200, dp_aliases200_['Memo/ Invoice Date'], true);
  var cDue       = dp_findHeaderIndex_(hm200, dp_aliases200_['Return DUE DATE'], true);

  var cVendor    = dp_findHeaderIndex_(hm200, dp_aliases200_['Vendor'], true);
  var cType      = dp_findHeaderIndex_(hm200, dp_aliases200_['Stone Type'], true);
  var cShape     = dp_findHeaderIndex_(hm200, dp_aliases200_['Shape'], true);
  var cCarat     = dp_findHeaderIndex_(hm200, dp_aliases200_['Carat'], true);
  var cColor     = dp_findHeaderIndex_(hm200, dp_aliases200_['Color'], true);
  var cClarity   = dp_findHeaderIndex_(hm200, dp_aliases200_['Clarity'], true);
  var cLab       = dp_findHeaderIndex_(hm200, dp_aliases200_['LAB'], true);
  var cCert      = dp_findHeaderIndex_(hm200, dp_aliases200_['Certificate No'], true);
  var cCompany   = dp_findHeaderIndex_(hm200, dp_aliases200_['Company'], true);
  var cAssigned  = dp_findHeaderIndex_(hm200, dp_aliases200_['Assigned Rep'], true);
  var cRoot      = dp_findHeaderIndex_(hm200, dp_aliases200_['RootApptID'], true);
  var cCust      = dp_findHeaderIndex_(hm200, dp_aliases200_['Customer Name'], true);
  var cAppt      = dp_findHeaderIndex_(hm200, dp_aliases200_['Customer Appt Time & Date'], true);
  var cOrderedBy = dp_findHeaderIndex_(hm200, dp_aliases200_['Ordered By'], true);
  var cOrdDate   = dp_findHeaderIndex_(hm200, dp_aliases200_['Purchased / Ordered Date'], true);

    Logger.log('Columns -> Order:' + cOrder + ', StoneStatus:' + cStoneStat +
           ', Memo:' + cMemo + ', Due:' + cDue);
  var lastRow = sh200.getLastRow();
  var items = [];
  if (lastRow >= 3) {
    var rng = sh200.getRange(3, 1, lastRow - 2, sh200.getLastColumn()).getDisplayValues();
    for (var i=0; i<rng.length; i++) {
      var rowIdx = i + 3, row = rng[i];
      var order = String(row[cOrder - 1] || '').trim();
      if (!/^On the Way$/i.test(order)) continue;

      var stoneStatus = String(row[cStoneStat - 1] || '').trim();
      if (/^In Stock$/i.test(stoneStatus)) continue; // already confirmed

      items.push({
        rowIndex: rowIdx,
        rootApptId: String(row[cRoot - 1] || '').trim(),
        company: String(row[cCompany - 1] || '').trim(),
        customerName: String(row[cCust - 1] || '').trim(),
        apptTimeDate: String(row[cAppt - 1] || '').trim(),
        assignedRep: String(row[cAssigned - 1] || '').trim(),
        vendor: String(row[cVendor - 1] || '').trim(),
        stoneType: String(row[cType - 1] || '').trim(),
        shape: String(row[cShape - 1] || '').trim(),
        carat: String(row[cCarat - 1] || '').trim(),
        color: String(row[cColor - 1] || '').trim(),
        clarity: String(row[cClarity - 1] || '').trim(),
        lab: String(row[cLab - 1] || '').trim(),
        certNo: String(row[cCert - 1] || '').trim(),
        orderedBy: String(row[cOrderedBy - 1] || '').trim(),
        orderedDate: String(row[cOrdDate - 1] || '').trim()
      });
    }
  }

  // Optional: try to show SO# from 100_ (if header exists)
  // We won't join per item (expensive) — but we can display RootApptID and Customer name; SO# can be added later if needed.

  // Defaults: Memo = today; Return DUE DATE = +20 days will be computed server-side
  var todayIso = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  return {
    items: items,
    defaults: {
      memoDateDefault: todayIso
    },
    meta: { spreadsheetUrl: target.ss.getUrl(), tab: target.tab }
  };
}

// -------------------------------------------------------------
// SUBMIT: set Stone Status = In Stock, write Memo/Invoice Date & Return DUE DATE
// -------------------------------------------------------------
/**
 * payload: {
 *   defaultMemoDate: 'YYYY-MM-DD' or '',
 *   applyDefaultToAll: boolean,
 *   items: [{ rowIndex, rootApptId, memoDate? , selected: true|false }, ...]
 * }
 */
function dp_submitConfirmDelivery(payload) {
  Logger.log('dp_submitConfirmDelivery: items=' + (payload && payload.items ? payload.items.length : 0) +
            ', applyDefault=' + !!payload.applyDefaultToAll +
            ', defaultMemo=' + (payload.defaultMemoDate || ''));

  // define target FIRST, then log it
  var target = dp_get200Sheet_();
  Logger.log('200_ target: ' + target.ss.getName() + ' / ' + target.tab);

  var sh200 = target.sheet;
  var hm200 = dp_headerMapFor200_(sh200);
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var lastCol = sh200.getLastColumn();

  var lock = LockService.getDocumentLock();
  lock.waitLock(28 * 1000);
  try {
    if (!payload || !Array.isArray(payload.items) || payload.items.length === 0) {
      throw new Error('No items to update.');
    }

    // Resolve columns for writes
    var cOrder     = dp_findHeaderIndex_(hm200, dp_aliases200_['Order Status'], true);
    var cCert = dp_findHeaderIndex_(hm200, dp_aliases200_['Certificate No'], true);
    var deliveredKeys = []; // for optional 400_ sync
    var cStoneStat = dp_findHeaderIndex_(hm200, dp_aliases200_['Stone Status'], true);
    var cMemo      = dp_findHeaderIndex_(hm200, dp_aliases200_['Memo/ Invoice Date'], true);
    var cDue       = dp_findHeaderIndex_(hm200, dp_aliases200_['Return DUE DATE'], true);

    var touchedAppointments = {};
    var updatedRows = [];
    var skipped = [];

    var useDefault = !!payload.applyDefaultToAll;
    var defaultMemoDate = dp_parseIsoDateOrEmpty_(payload.defaultMemoDate, tz) || new Date(); // fallback = today

    // Process selection
    for (var i=0; i<payload.items.length; i++) {
      var it = payload.items[i];
      // Items are already filtered as "selected" on the client.
      // Treat every item that arrives here as selected.
      if (!it || !it.rowIndex) continue;

      var r = Number(it.rowIndex);
      if (!(r >= 3)) { skipped.push({rowIndex:r, reason:'Row index invalid'}); continue; }

      var rowRng = sh200.getRange(r, 1, 1, lastCol);
      var rowDisp = rowRng.getDisplayValues()[0];
      var rowVals = rowRng.getValues()[0];

      var order = String(rowDisp[cOrder - 1] || '').trim();
      if (!/^On the Way$/i.test(order)) {
        skipped.push({rowIndex:r, reason: 'Order Status is not "On the Way" ('+order+')'});
        continue;
      }

      // Compute memo date
      var memoDate = useDefault
        ? defaultMemoDate
        : (dp_parseIsoDateOrEmpty_(it.memoDate, tz) || defaultMemoDate);

      // Compute due date = memo + 20 days (C15)
      var dueDate = dp_addDays_(memoDate, 20);

      // Merge Stone Status with "In Stock" (do not replace)
      var mergedStoneStatus = cd_mergeStoneStatus_(String(rowDisp[cStoneStat - 1] || ''), 'In Stock');

      // Set values
      rowVals[cStoneStat - 1] = mergedStoneStatus; // e.g., "Diamond Viewing ; In Stock"
      rowVals[cOrder - 1]     = 'Delivered';       // <-- Order Status
      rowVals[cMemo  - 1]     = memoDate;          // Date object
      rowVals[cDue   - 1]     = dueDate;           // Date object (+20 days)

      rowRng.setValues([rowVals]);

      // 🔎 Read-back verification (logs the actual cell values right after write)
      try {
        var rb = sh200.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
        Logger.log('POST-WRITE row '+r+': ' + JSON.stringify({
          orderStatus: rb[cOrder-1],
          stoneStatus: rb[cStoneStat-1],
          memoDate: rb[cMemo-1],
          dueDate: rb[cDue-1]
        }, null, 2));
      } catch (e) {
        Logger.log('Readback error for row '+r+': ' + (e && e.message ? e.message : e));
      }

      updatedRows.push({rowIndex: r, rootApptId: it.rootApptId});

      // Keep keys for optional 400_ sync
      deliveredKeys.push({
        rowIndex: r,
        rootApptId: it.rootApptId,
        certNo: String(rowDisp[cCert - 1] || '').trim()
      });

      if (it.rootApptId) touchedAppointments[it.rootApptId] = true;
    }

    // Refresh 100_ and set CSOS based on Delivered counts
    // After writing all selected rows…
    var results = [];
    for (var apptId in touchedAppointments) {
      // Recount after writes
      var counts = dp_computeCountsForAppointment_(sh200, hm200, apptId);
      var csos   = dp_decideCsosAfterDelivery_(counts);

      // Keep existing helper (if present) to set CSOS etc.
      var applied = null;
      if (typeof dp_applyCsosAndRefresh100_ === 'function') {
        try { applied = dp_applyCsosAndRefresh100_(apptId, csos, sh200, hm200, counts); } catch (e) {
          Logger.log('dp_applyCsosAndRefresh100_ error: ' + (e && e.message ? e.message : e));
        }
      } else {
        // Minimal CSOS write if your helper isn't present
        try {
          var loc = dp_find100RowByRootApptId_(apptId);
          if (loc) dp_setCellByHeader_(loc.sheet, loc.headerMap, loc.rowIndex, dp_aliases100_['Center Stone Order Status'], csos);
        } catch(e) {
          Logger.log('Set CSOS (fallback) error: ' + (e && e.message ? e.message : e));
        }
      }

      // ✅ Always refresh the DV Stones Summary on 100_ (idempotent)
      try {
        // If you only want to force the summary line (lighter), call the helper below:
      dp_update100SummaryOnly_(apptId, counts);
      } catch (e) {
        Logger.log('DV summary refresh error: ' + (e && e.message ? e.message : e));
      }

      // Optional reminders hook
      try { dp_onCsosChanged_(apptId, csos); } catch (e) {}

      results.push({ rootApptId: apptId, csos: csos, counts: counts, applied: applied });
    }

    return {
      ok: true,
      updatedCount: updatedRows.length,
      updatedRows: updatedRows,
      skipped: skipped,
      appointments: results,
      meta: { spreadsheetUrl: target.ss.getUrl(), tab: target.tab }
    };

  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// -------------------------------------------------------------
// 100_ quick reference refresh (no CSOS change)
// -------------------------------------------------------------
// Write DV Stones Summary on ALL 100_ rows for a RootApptID.
function dp_update100SummaryOnly_(rootApptId, counts) {
  var locs = dp_find100RowsByRootApptId_(rootApptId);
  if (!locs || !locs.length) return { ok:false, updated:0, reason:'RootApptID not found in 100_' };

  var summary = 'Proposed: ' + (counts.proposing || 0) +
                ' • On the Way: ' + (counts.onTheWay || 0) +
                ' • Not Approved: ' + (counts.notApproved || 0) +
                ' • In Stock: ' + (counts.inStock || 0) +
                ' • Total: ' + (counts.total || 0);

  var n = 0;
  locs.forEach(function(loc){
    var sh100 = loc.sheet, hm100 = loc.headerMap, r = loc.rowIndex;
    var colSum = dp_findHeaderIndex_(hm100, dp_aliases100_['DV Stones Summary'], true);
    sh100.getRange(r, colSum).setValue(summary);
    n++;
  });
  return { ok:true, updated: n, summary: summary };
}
/**
 * Set CSOS on ALL matching 100_ rows and also update JSON-lines + Summary.
 * Returns { ok, updated, rows: [rowIndex...] }.
 */
function dp_applyCsosAndRefresh100_(rootApptId, csos, sh200, hm200, counts) {
  var locs = dp_find100RowsByRootApptId_(rootApptId);
  if (!locs || !locs.length) return { ok:false, updated:0, reason:'RootApptID not found in 100_' };

  var snap = dp_buildSnapshotFrom200_(rootApptId, sh200, hm200);
  var summary = 'Proposed: ' + (counts.proposing||0) +
                ' • On the Way: ' + (counts.onTheWay||0) +
                ' • Not Approved: ' + (counts.notApproved||0) +
                ' • In Stock: ' + (counts.inStock||0) +
                ' • Total: ' + (counts.total||0);

  var rowsTouched = [];
  locs.forEach(function(loc){
    var sh100 = loc.sheet, hm100 = loc.headerMap, r = loc.rowIndex;
    // CSOS
    dp_setCellByHeader_(sh100, hm100, r, dp_aliases100_['Center Stone Order Status'], csos);
    // JSON lines + Summary
    var colJson = dp_findHeaderIndex_(hm100, dp_aliases100_['DV Stones (JSON Lines)'], true);
    var colSum  = dp_findHeaderIndex_(hm100, dp_aliases100_['DV Stones Summary'], true);
    sh100.getRange(r, colJson).setValue(snap.jsonLines);
    sh100.getRange(r, colSum).setValue(summary);
    rowsTouched.push(r);
  });

  Logger.log('dp_applyCsosAndRefresh100_: CSOS "'+csos+'" on ' + rowsTouched.length + ' row(s) for RootApptID=' + rootApptId);
  return { ok:true, updated: rowsTouched.length, rows: rowsTouched };
}



// -------------------------------------------------------------
// UTIL: add days in local timezone
// -------------------------------------------------------------
function dp_addDays_(dateObj, days) {
  var d = new Date(dateObj);
  d.setDate(d.getDate() + Number(days||0));
  return d;
}

// -------------------------------------------------------------
// Extend 200_ aliases for Memo/Due (newline & variant-safe)
// -------------------------------------------------------------
try {
  if (typeof dp_aliases200_ !== 'undefined') {
    dp_aliases200_['Memo/ Invoice Date'] = [
      'Memo/ Invoice Date','Memo / Invoice Date','Memo Invoice Date','Memo- Invoice Date','Memo/Invoice Date'
    ];
    dp_aliases200_['Return DUE DATE'] = [
      'Return DUE DATE','Return\nDUE DATE','Return Due Date','Return DUE','Return Due'
    ];
  }
} catch(e) {}

// Join tokens with ", " like Sheets multi-select; dedup + preserve order.
function cd_mergeStoneStatus_(current, toAdd) {
  var cur = String(current || '').trim();
  var add = String(toAdd || '').trim();
  if (!cur) return add || '';

  // split on common delimiters ; | , •
  var parts = cur.split(/\s*[;|,]\s*|\s*•\s*/g)
                 .map(function(s){ return s.trim(); })
                 .filter(Boolean);

  if (add && !parts.some(function(p){ return p.toLowerCase() === add.toLowerCase(); })) {
    parts.push(add);
  }
  return parts.join(', ');   // <-- comma + space
}

function dp_decideCsosAfterDelivery_(counts) {
  // All delivered when no remaining Proposing/On the Way and at least one Delivered.
  if ((counts.delivered || 0) > 0 && (counts.onTheWay || 0) === 0 && (counts.proposing || 0) === 0) {
    return 'Diamond Memo – Delivered';
  }
  if ((counts.delivered || 0) > 0) {
    return 'Diamond Memo – SOME Delivered';
  }
  // fallback (shouldn't happen in this flow)
  return 'Diamond Memo – Proposed';
}

// --- Legacy → Canon shims (safe no-ops if the name already exists in this file) ---
if (typeof headerMap_ !== 'function') {
  function headerMap_(sh){ return headerMap__canon(sh); }
}
if (typeof ensureHeaders_ !== 'function') {
  function ensureHeaders_(sh, labels){ return ensureHeaders__canon(sh, labels); }
}
if (typeof getMasterSheet_ !== 'function') {
  function getMasterSheet_(ss){ return getMasterSheet__canon(ss); }
}
if (typeof getOrdersSheet_ !== 'function') {
  function getOrdersSheet_(wb){ return getOrdersSheet__canon(wb); }
}
if (typeof coerceSOTextColumn_ !== 'function') {
  function coerceSOTextColumn_(sh, H){ return coerceSOTextColumn__canon(sh, H); }
}
if (typeof existsSOInMaster_ !== 'function') {
  function existsSOInMaster_(sh, brand, so, skipRow){ return existsSOInMaster__canon(sh, brand, so, skipRow); }
}



