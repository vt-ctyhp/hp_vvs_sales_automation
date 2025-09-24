/**
 * Diamonds_OrderApprove_v1.gs — Phase 2 (Dialog B: Order/Approve Manager)
 * Dependencies (from Phase 1 / shared):
 *  - dp_get200Sheet_(), dp_headerMapFor200_(), dp_headerMap_()
 *  - dp_findHeaderIndex_(), dp_setCellByHeader_(), dp_aliases200_, dp_aliases100_
 *  - dp_computeCountsForAppointment_(), dp_norm_()
 *  - dp_onCsosChanged_(rootApptId, newStatus)  // optional hook; safe no-op if not wired
 *
 * This module:
 *  - dp_openOrderApproveDiamonds(): menu entrypoint
 *  - dp_bootstrapForOrderDialog_(): returns proposing stones list + defaults
 *  - dp_submitOrderApprovals(payload): writes 200_ updates, then updates 100_ CSOS & quick reference
 *  - helpers for 100_ locate/update and snapshot refresh from 200_
 */

// ------------------------------ PUBLIC ENTRYPOINTS ------------------------------

/** Menu entrypoint to open Dialog B. */
function dp_openOrderApproveDiamonds() {
  var ui = SpreadsheetApp.getUi();

  // Pre-compute everything and inject into HTML (more reliable than client bootstrap)
  var bootstrap;
  try {
    bootstrap = dp_bootstrapForOrderDialog_();
  } catch (e) {
    ui.alert('💎 Diamonds — Order/Approve', (e && e.message ? e.message : String(e)), ui.ButtonSet.OK);
    return;
  }

  var t = HtmlService.createTemplateFromFile('dlg_order_approve_diamonds_v1');
  t.bootstrap = bootstrap; // inject data
  var html = t.evaluate().setWidth(1250).setHeight(720);
  ui.showModalDialog(html, '💎 Order/Approve Diamonds');
}

/** Build bootstrap payload: proposing stones across 200_, defaults, and dropdowns. */
function dp_bootstrapForOrderDialog_() {
  var target = dp_get200Sheet_();
  var sh200 = target.sheet;
  var hm200 = dp_headerMapFor200_(sh200);
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

  var colOrder      = dp_findHeaderIndex_(hm200, dp_aliases200_['Order Status'], true);
  var colVendor     = dp_findHeaderIndex_(hm200, dp_aliases200_['Vendor'], true);
  var colType       = dp_findHeaderIndex_(hm200, dp_aliases200_['Stone Type'], true);
  var colShape      = dp_findHeaderIndex_(hm200, dp_aliases200_['Shape'], true);
  var colCarat      = dp_findHeaderIndex_(hm200, dp_aliases200_['Carat'], true);
  var colColor      = dp_findHeaderIndex_(hm200, dp_aliases200_['Color'], true);
  var colClarity    = dp_findHeaderIndex_(hm200, dp_aliases200_['Clarity'], true);
  var colLab        = dp_findHeaderIndex_(hm200, dp_aliases200_['LAB'], true);
  var colCert       = dp_findHeaderIndex_(hm200, dp_aliases200_['Certificate No'], true);
  var colAssigned   = dp_findHeaderIndex_(hm200, dp_aliases200_['Assigned Rep'], true);
  var colRoot       = dp_findHeaderIndex_(hm200, dp_aliases200_['RootApptID'], true);
  var colCust       = dp_findHeaderIndex_(hm200, dp_aliases200_['Customer Name'], true);
  var colAppt       = dp_findHeaderIndex_(hm200, dp_aliases200_['Customer Appt Time & Date'], true);
  var colCompany    = dp_findHeaderIndex_(hm200, dp_aliases200_['Company'], true);
  var colReqBy      = dp_findHeaderIndex_(hm200, dp_aliases200_['Requested By'], false);
  var colReqDate    = dp_findHeaderIndex_(hm200, dp_aliases200_['Request Date'], false);

  var last = sh200.getLastRow();
  if (last < 3) {
    return {
      items: [],
      defaults: dp_defaultsForOrder_(),
      dropdowns: dp_orderDropdowns_(),
      meta: { spreadsheetUrl: target.ss.getUrl(), tab: target.tab }
    };
  }

  // Scan all data rows; collect where Order Status == Proposing
  var rng = sh200.getRange(3, 1, last - 2, sh200.getLastColumn()).getDisplayValues();
  var items = [];
  for (var i = 0; i < rng.length; i++) {
    var rowIdx = i + 3;
    var row = rng[i];

    var order = String(row[colOrder - 1] || '').trim();
    if (!/^Proposing$/i.test(order)) continue;

    items.push({
      rowIndex: rowIdx,
      rootApptId: String(row[colRoot - 1] || '').trim(),
      company: String(row[colCompany - 1] || '').trim(),
      customerName: String(row[colCust - 1] || '').trim(),
      apptTimeDate: String(row[colAppt - 1] || '').trim(),
      assignedRep: String(row[colAssigned - 1] || '').trim(),
      vendor: String(row[colVendor - 1] || '').trim(),
      stoneType: String(row[colType - 1] || '').trim(),
      shape: String(row[colShape - 1] || '').trim(),
      carat: String(row[colCarat - 1] || '').trim(),
      color: String(row[colColor - 1] || '').trim(),
      clarity: String(row[colClarity - 1] || '').trim(),
      lab: String(row[colLab - 1] || '').trim(),
      certNo: String(row[colCert - 1] || '').trim(),
      requestedBy: (colReqBy > -1 ? String(row[colReqBy - 1] || '').trim() : ''),
      requestDate: (colReqDate > -1 ? String(row[colReqDate - 1] || '').trim() : '')
    });
  }

  // Sort by Customer, then RootApptID, then RequestDate desc (string sort is fine as display)
  items.sort(function(a,b){
    if (a.customerName !== b.customerName) return a.customerName.localeCompare(b.customerName);
    if (a.rootApptId !== b.rootApptId) return a.rootApptId.localeCompare(b.rootApptId);
    return String(b.requestDate).localeCompare(String(a.requestDate));
  });

  return {
    items: items,
    defaults: dp_defaultsForOrder_(),
    dropdowns: dp_orderDropdowns_(),
    meta: { spreadsheetUrl: target.ss.getUrl(), tab: target.tab }
  };
}

/**
 * Submit handler for Dialog B.
 * payload: {
 *   defaultOrderedBy: string,
 *   defaultOrderDate: 'YYYY-MM-DD' or '',
 *   applyDefaultsToAll: boolean,
 *   items: [{ rowIndex, rootApptId, decision: 'On the Way'|'Not Approved', orderedBy?, orderedDate? }, ...]
 * }
 */
function dp_submitOrderApprovals(payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(28 * 1000);
  try {
    if (!payload || !Array.isArray(payload.items) || payload.items.length === 0) {
      throw new Error('No changes to save.');
    }

    var target = dp_get200Sheet_();
    var sh200 = target.sheet;
    var hm200 = dp_headerMapFor200_(sh200);
    var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    var lastCol = sh200.getLastColumn();

    var cOrder   = dp_findHeaderIndex_(hm200, dp_aliases200_['Order Status'], true);
    var cOrdBy   = dp_findHeaderIndex_(hm200, dp_aliases200_['Ordered By'], true);
    var cOrdDate = dp_findHeaderIndex_(hm200, dp_aliases200_['Purchased / Ordered Date'], true);

    var updated = [];
    var skipped = [];
    var touchByAppt = {}; // rootApptId -> true

    var useDefaults = !!payload.applyDefaultsToAll;
    var defBy = String(payload.defaultOrderedBy || '').trim() || dp_getCurrentUserEmail_();
    var defDate = dp_parseIsoDateOrEmpty_(payload.defaultOrderDate, tz); // may be null

    // Process each row
    for (var i = 0; i < payload.items.length; i++) {
      var it = payload.items[i];
      if (!it || !it.rowIndex || !it.decision) continue; // ignore empty lines
      var r = Number(it.rowIndex);
      if (!(r >= 3)) { skipped.push({rowIndex:r, reason:'Row index invalid'}); continue; }

      // Read the whole row (one call) so we can safely write it back
      var rowRng = sh200.getRange(r, 1, 1, lastCol);
      var rowVals = rowRng.getValues()[0];
      var rowDisp = sh200.getRange(r, 1, 1, lastCol).getDisplayValues()[0]; // for checks

      var currentOrder = String(rowDisp[cOrder - 1] || '').trim();
      if (!/^Proposing$/i.test(currentOrder)) {
        skipped.push({rowIndex:r, reason:'Order Status is no longer Proposing (' + currentOrder + ')'});
        continue;
      }

      // Compute values
      var newOrder = it.decision; // 'On the Way' | 'Not Approved'
      var ordBy = useDefaults ? defBy : (String(it.orderedBy || '').trim() || defBy);
      var ordDate = null;
      if (newOrder === 'On the Way') {
        var iso = useDefaults ? payload.defaultOrderDate : (it.orderedDate || payload.defaultOrderDate);
        ordDate = dp_parseIsoDateOrEmpty_(iso, tz) || new Date(); // fallback = today
      } else {
        // Not Approved -> clear date
        ordDate = '';
      }

      // Write into the row array at absolute columns
      rowVals[cOrder - 1]   = newOrder;
      rowVals[cOrdBy - 1]   = ordBy;
      rowVals[cOrdDate - 1] = ordDate;

      // Single write per row
      rowRng.setValues([rowVals]);

      updated.push({rowIndex:r, rootApptId: it.rootApptId, set: {order:newOrder, by:ordBy}});
      if (it.rootApptId) touchByAppt[it.rootApptId] = true;
    }

    // For each touched appointment, recompute counts, set CSOS on 100_, refresh quick reference
    var resultsByAppt = [];
    var sh200Again = target.sheet; // same
    var hm200Again = hm200;

    for (var apptId in touchByAppt) {
      var counts = dp_computeCountsForAppointment_(sh200Again, hm200Again, apptId);
      var csos = dp_decideCsosFromCounts_(counts, apptId);
      var applied = dp_applyCsosAndRefresh100_(apptId, csos, sh200Again, hm200Again, counts);
      resultsByAppt.push({rootApptId: apptId, csos: csos, counts: counts, applied: applied});
      try { dp_onCsosChanged_(apptId, csos); } catch (e) {} // reminders hook (safe)
    }

    return {
      ok: true,
      updatedCount: updated.length,
      updatedRows: updated,
      skipped: skipped,
      appointments: resultsByAppt,
      meta: { spreadsheetUrl: target.ss.getUrl(), tab: target.tab }
    };

  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// ------------------------------ INTERNAL HELPERS ------------------------------

function dp_orderDropdowns_() {
  return {
    decisions: ['On the Way','Not Approved']
  };
}

function dp_defaultsForOrder_() {
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var todayIso = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  return {
    orderedByDefault: dp_getCurrentUserEmail_(),
    orderedDateDefault: todayIso
  };
}

/** Parse 'YYYY-MM-DD' in local tz -> Date, else null/'' -> null. */
function dp_parseIsoDateOrEmpty_(iso, tz) {
  var s = String(iso || '').trim();
  if (!s) return null;
  // 'YYYY-MM-DDT00:00:00' local
  var parts = s.split('-');
  if (parts.length !== 3) return null;
  var y = Number(parts[0]), m = Number(parts[1]) - 1, d = Number(parts[2]);
  var dt = new Date(y, m, d, 0, 0, 0, 0);
  if (isNaN(dt)) return null;
  return dt;
}

/** Decide CSOS label based on counts. */
function dp_decideCsosFromCounts_(counts, rootApptId) {
  // counts: {proposing,onTheWay,notApproved,inStock,total}
  if (counts.onTheWay > 0 && counts.proposing === 0 && counts.notApproved === 0) {
    return 'Diamond Memo – On the Way';
  }
  if (counts.onTheWay > 0) {
    return 'Diamond Memo – SOME On the Way';
  }
  if (counts.onTheWay === 0 && counts.notApproved > 0 && counts.proposing === 0) {
    return 'Diamond Memo – NONE APPROVED';
  }
  // If still only proposing -> keep Proposed (do not override)
  return 'Diamond Memo – Proposed';
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

/** Find the 100_ row by RootApptID (exact match). */
function dp_find100RowByRootApptId_(rootApptId) {
  var ss = SpreadsheetApp.getActive();
  // Try active sheet first; if not 100_, look up by name
  var sh = ss.getActiveSheet();
  var hm = dp_headerMap_(sh);
  var maybeRoot = dp_findHeaderIndex_(hm, dp_aliases100_['RootApptID'], false);

  if (maybeRoot < 0 || !/00[_\s-]*Master\s*Appointments/i.test(sh.getName())) {
    var fallback = ss.getSheets().find(function(s){ return /00[_\s-]*Master\s*Appointments/i.test(s.getName()); });
    if (!fallback) return null;
    sh = fallback;
    hm = dp_headerMap_(sh);
    maybeRoot = dp_findHeaderIndex_(hm, dp_aliases100_['RootApptID'], true);
  }

  var col = maybeRoot;
  var last = sh.getLastRow();
  if (last < 2) return null;

  var vals = sh.getRange(2, col, last-1, 1).getDisplayValues();
  for (var i=0; i<vals.length; i++) {
    if (String(vals[i][0] || '').trim() === String(rootApptId)) {
      return { sheet: sh, rowIndex: i+2, headerMap: hm };
    }
  }
  return null;
}

/** Build the appointment snapshot (JSON-lines) from 200_. */
function dp_buildSnapshotFrom200_(rootApptId, sh200, hm200) {
  var cRoot     = dp_findHeaderIndex_(hm200, dp_aliases200_['RootApptID'], true);
  var cVendor   = dp_findHeaderIndex_(hm200, dp_aliases200_['Vendor'], true);
  var cType     = dp_findHeaderIndex_(hm200, dp_aliases200_['Stone Type'], true);
  var cShape    = dp_findHeaderIndex_(hm200, dp_aliases200_['Shape'], true);
  var cCarat    = dp_findHeaderIndex_(hm200, dp_aliases200_['Carat'], true);
  var cColor    = dp_findHeaderIndex_(hm200, dp_aliases200_['Color'], true);
  var cClarity  = dp_findHeaderIndex_(hm200, dp_aliases200_['Clarity'], true);
  var cLab      = dp_findHeaderIndex_(hm200, dp_aliases200_['LAB'], true);
  var cCert     = dp_findHeaderIndex_(hm200, dp_aliases200_['Certificate No'], true);
  var cOrder    = dp_findHeaderIndex_(hm200, dp_aliases200_['Order Status'], true);
  var cStoneSt  = dp_findHeaderIndex_(hm200, dp_aliases200_['Stone Status'], true);
  var cReqDate  = dp_findHeaderIndex_(hm200, dp_aliases200_['Request Date'], false);
  var cOrdBy    = dp_findHeaderIndex_(hm200, dp_aliases200_['Ordered By'], true);
  var cOrdDate  = dp_findHeaderIndex_(hm200, dp_aliases200_['Purchased / Ordered Date'], true);

  var last = sh200.getLastRow();
  if (last < 3) return { jsonLines: '' };

  var rows = sh200.getRange(3, 1, last-2, sh200.getLastColumn()).getDisplayValues();
  var lines = [];
  for (var i=0;i<rows.length;i++) {
    var row = rows[i];
    if (String(row[cRoot - 1] || '').trim() !== String(rootApptId)) continue;

    var item = {
      vendor: row[cVendor - 1] || '',
      stoneType: row[cType - 1] || '',
      shape: row[cShape - 1] || '',
      carat: row[cCarat - 1] || '',
      color: row[cColor - 1] || '',
      clarity: row[cClarity - 1] || '',
      lab: row[cLab - 1] || '',
      certNo: row[cCert - 1] || '',
      orderStatus: row[cOrder - 1] || '',
      stoneStatus: row[cStoneSt - 1] || '',
      requestDate: (cReqDate > -1 ? row[cReqDate - 1] : ''),
      orderedBy: row[cOrdBy - 1] || '',
      orderedDate: row[cOrdDate - 1] || ''
    };
    lines.push(JSON.stringify(item));
  }
  return { jsonLines: lines.join('\n') };
}

// ---- extend aliases (non-breaking) for 100_ "SO#" if you later want to surface it ----
try {
  if (typeof dp_aliases100_ !== 'undefined' && !dp_aliases100_['SO#']) {
    dp_aliases100_['SO#'] = ['SO#','SO Number','Sales Order #','SO'];
  }
} catch (e) {}

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



