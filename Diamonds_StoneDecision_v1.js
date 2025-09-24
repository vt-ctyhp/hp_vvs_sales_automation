/** 
 * Phase 4 — Dialog D: Client Stone Decisions (Purchase / Return / Hold)
 * - Opens from the active 100_ client row
 * - Edits 200_ "Stone Decision (PO, Return)" and "Stone Status" (adds/removes tokens)
 * - Refreshes 100_ JSON lines and summary; does NOT change CSOS here
 */

// -------------------------------------------------------------
// MENU OPENER
// -------------------------------------------------------------
function dp_openStoneDecisions() {
  var ui = SpreadsheetApp.getUi();
  var bootstrap;
  try {
    bootstrap = dp_bootstrapForStoneDecisions_();
  } catch (e) {
    ui.alert('💎 Diamonds — Client Decisions', (e && e.message ? e.message : String(e)) +
      '\n\nOpen 100_ (00_Master Appointments) and select a customer data row, then try again.', ui.ButtonSet.OK);
    return;
  }
  var t = HtmlService.createTemplateFromFile('dlg_stone_decision_v1');
  t.bootstrap = bootstrap; // inject payload
  var html = t.evaluate().setWidth(1200).setHeight(720);
  ui.showModalDialog(html, '💎 Record Client Decisions');
}

// -------------------------------------------------------------
// BOOTSTRAP: load all stones for the active 100_ row (by RootApptID)
// -------------------------------------------------------------
function dp_bootstrapForStoneDecisions_() {
  var ctx = dp_getActiveMasterRowContext_(); // validates 100_ selection and core headers
  
  // Try to read SO# for top bar (best-effort)
  var soNumber = '';
  try {
    if (typeof dp_aliases100_ !== 'undefined') {
      dp_aliases100_['SO#'] = dp_aliases100_['SO#'] || ['SO#','SO #','SO Number','SO No','Sales Order #','Sales Order Number','SO'];
    }
    var colSO = dp_findHeaderIndex_(ctx.headerMap, dp_aliases100_['SO#'] || ['SO#'], false);
    if (colSO > -1) {
      soNumber = String(ctx.sheet.getRange(ctx.rowIndex, colSO).getDisplayValue() || '').trim();
    }
  } catch (e) {}

  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var apptAt = dp_composeApptDateTimeString_(ctx.visitDateValue, ctx.visitTimeValue, tz);

  // 200_ target and columns
  var target = dp_get200Sheet_();
  var sh200 = target.sheet;
  var hm200 = dp_headerMapFor200_(sh200);

  var cRoot      = dp_findHeaderIndex_(hm200, dp_aliases200_['RootApptID'], true);
  var cOrder     = dp_findHeaderIndex_(hm200, dp_aliases200_['Order Status'], true);
  var cStoneStat = dp_findHeaderIndex_(hm200, dp_aliases200_['Stone Status'], true);
  var cDecision  = dp_findHeaderIndex_(hm200, dp_aliases200_['Stone Decision (PO, Return)'], true);

  var cVendor  = dp_findHeaderIndex_(hm200, dp_aliases200_['Vendor'], true);
  var cType    = dp_findHeaderIndex_(hm200, dp_aliases200_['Stone Type'], true);
  var cShape   = dp_findHeaderIndex_(hm200, dp_aliases200_['Shape'], true);
  var cCarat   = dp_findHeaderIndex_(hm200, dp_aliases200_['Carat'], true);
  var cColor   = dp_findHeaderIndex_(hm200, dp_aliases200_['Color'], true);
  var cClarity = dp_findHeaderIndex_(hm200, dp_aliases200_['Clarity'], true);
  var cLab     = dp_findHeaderIndex_(hm200, dp_aliases200_['LAB'], true);
  var cCert    = dp_findHeaderIndex_(hm200, dp_aliases200_['Certificate No'], true);

  var items = [];
  var lastRow = sh200.getLastRow();
  if (lastRow >= 3) {
    var data = sh200.getRange(3, 1, lastRow - 2, sh200.getLastColumn()).getDisplayValues();
    for (var i = 0; i < data.length; i++) {
      var row = data[i], rIdx = i + 3;
      if (String(row[cRoot - 1] || '').trim() !== String(ctx.rootApptId)) continue;

      var stoneStatus = String(row[cStoneStat - 1] || '').trim();
      items.push({
        rowIndex: rIdx,
        rootApptId: ctx.rootApptId,
        orderStatus: String(row[cOrder - 1] || '').trim(),
        stoneStatus: stoneStatus,
        // Quick flags for UI
        inStock: /\bIn Stock\b/i.test(stoneStatus),
        hold: /\bHOLD for Cust\.\b/i.test(stoneStatus),
        purchased: /\bCustomer Purchased\b/i.test(stoneStatus),
        decision: String(row[cDecision - 1] || '').trim(), // "" | Purchase | Return
        // Specs
        vendor: String(row[cVendor - 1] || '').trim(),
        stoneType: String(row[cType - 1] || '').trim(),
        shape: String(row[cShape - 1] || '').trim(),
        carat: String(row[cCarat - 1] || '').trim(),
        color: String(row[cColor - 1] || '').trim(),
        clarity: String(row[cClarity - 1] || '').trim(),
        lab: String(row[cLab - 1] || '').trim(),
        certNo: String(row[cCert - 1] || '').trim()
      });
    }
  }

  return {
    context: {
      rootApptId: ctx.rootApptId,
      customerName: ctx.customerName,
      soNumber: soNumber || '',
      apptAt: apptAt || '',
      companyBrand: ctx.companyBrand || '',
      assignedRep: ctx.assignedRep || ''
    },
    dropdowns: {
      decisions: ['Purchase', 'Return']
    },
    items: items,
    meta: { spreadsheetUrl: target.ss.getUrl(), tab: target.tab }
  };
}

// -------------------------------------------------------------
// SUBMIT: apply Decision + Hold to 200_ and refresh 100_ JSON/Summary
// -------------------------------------------------------------
/**
 * payload: {
 *   items: [{ rowIndex, rootApptId, decision?, hold? }, ...]   // only edited rows
 * }
 */
function dp_submitStoneDecisions(payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(28 * 1000);

  try {
    if (!payload || !Array.isArray(payload.items) || payload.items.length === 0) {
      throw new Error('Nothing to save.');
    }

    var target = dp_get200Sheet_();
    var sh200 = target.sheet;
    var hm200 = dp_headerMapFor200_(sh200);
    var lastCol = sh200.getLastColumn();

    // Resolve columns
    var cStoneStat = dp_findHeaderIndex_(hm200, dp_aliases200_['Stone Status'], true);
    var cDecision  = dp_findHeaderIndex_(hm200, dp_aliases200_['Stone Decision (PO, Return)'], true);
    var cCert      = dp_findHeaderIndex_(hm200, dp_aliases200_['Certificate No'], true);

    var updated = 0;
    var touchedAppointments = {};
    var updatesForJsonByAppt = {}; // rootApptId -> [{certNo, decision?, hold?}, ...]

    payload.items.forEach(function(it){
      if (!it || !it.rowIndex) return;

      var r = Number(it.rowIndex);
      if (!(r >= 3)) return;

      var rowRng  = sh200.getRange(r, 1, 1, lastCol);
      var rowDisp = rowRng.getDisplayValues()[0];
      var rowVals = rowRng.getValues()[0];

      var currentStatus = String(rowDisp[cStoneStat - 1] || '').trim();
      var nextStatus = currentStatus;

      // Toggle HOLD for Cust.
      if (typeof it.hold === 'boolean') {
        nextStatus = it.hold
          ? cd_mergeStoneStatus_(nextStatus, 'HOLD for Cust.')
          : cd_removeStoneStatusToken_(nextStatus, 'HOLD for Cust.');
      }

      // Apply Decision
      var decisionToSet = '';
      if (it.decision === 'Purchase') {
        decisionToSet = 'Purchase';
        // add "Customer Purchased" into Stone Status
        nextStatus = cd_mergeStoneStatus_(nextStatus, 'Customer Purchased');
      } else if (it.decision === 'Return') {
        decisionToSet = 'Return';
        // remove "Customer Purchased" if present
        nextStatus = cd_removeStoneStatusToken_(nextStatus, 'Customer Purchased');
      } // else leave Decision unchanged

      // --- Write Stone Status with validation-safe path
      cd_writeBypassValidation_(sh200, r, cStoneStat, nextStatus);

      // --- Write Stone Decision if provided
      if (decisionToSet) {
        sh200.getRange(r, cDecision).setValue(decisionToSet);
      }

      updated++;
      if (it.rootApptId) {
        touchedAppointments[it.rootApptId] = true;
        var cert = String(rowDisp[cCert - 1] || '').trim();
        (updatesForJsonByAppt[it.rootApptId] = updatesForJsonByAppt[it.rootApptId] || [])
          .push({ certNo: cert, decision: decisionToSet || null, hold: (typeof it.hold === 'boolean' ? !!it.hold : null) });
      }
    });

    // Refresh 100_ JSON lines + summary (no CSOS changes in this phase)
    var results = [];
    for (var apptId in touchedAppointments) {
      // 1) update JSON lines with decisions/hold flags (best-effort)
      try { dp_update100JsonLinesWithDecisions_(apptId, updatesForJsonByAppt[apptId] || []); } catch (e) {}

      // 2) refresh summary snapshot
      var counts = dp_computeCountsForAppointment_(sh200, hm200, apptId);
      dp_refresh100QuickRef_(apptId, counts, sh200, hm200);
      results.push({ rootApptId: apptId, counts: counts });
    }

    return {
      ok: true,
      updatedCount: updated,
      appointments: results,
      meta: { spreadsheetUrl: target.ss.getUrl(), tab: target.tab }
    };

  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// -------------------------------------------------------------
// JSON lines updater for 100_ (adds decision / hold per certNo)
// -------------------------------------------------------------
function dp_update100JsonLinesWithDecisions_(rootApptId, updates) {
  if (!updates || !updates.length) return { ok:true, updated:0 };

  var loc = dp_find100RowByRootApptId_(rootApptId);
  if (!loc) return { ok:false, reason:'RootApptID not found in 100_' };

  var sh100 = loc.sheet, row = loc.rowIndex, hm100 = loc.headerMap;
  var colJson = dp_findHeaderIndex_(hm100, dp_aliases100_['DV Stones (JSON Lines)'], true);

  var raw = String(sh100.getRange(row, colJson).getValue() || '').trim();
  if (!raw) return { ok:true, updated:0 };

  var lines = raw.split(/\r?\n/);
  var mapByCert = {};
  try {
    // build quick index by certNo (lowercased)
    for (var i = 0; i < lines.length; i++) {
      var obj;
      try { obj = JSON.parse(lines[i]); } catch (e) { continue; }
      if (!obj) continue;
      var key = String((obj.certNo || '')).trim().toLowerCase();
      if (key) mapByCert[key] = { obj: obj, idx: i };
    }

    var changed = 0;
    updates.forEach(function(u){
      var key = String((u.certNo || '')).trim().toLowerCase();
      if (!key || !mapByCert[key]) return;
      var rec = mapByCert[key].obj;
      var idx = mapByCert[key].idx;
      var dirty = false;

      if (u.decision !== null && typeof u.decision !== 'undefined' && u.decision !== '') {
        if (rec.decision !== u.decision) { rec.decision = u.decision; dirty = true; }
      }
      if (u.hold !== null && typeof u.hold !== 'undefined') {
        if (rec.hold !== u.hold) { rec.hold = !!u.hold; dirty = true; }
      }
      if (dirty) {
        lines[idx] = JSON.stringify(rec);
        changed++;
      }
    });

    if (changed > 0) {
      sh100.getRange(row, colJson).setValue(lines.join('\n'));
    }
    return { ok:true, updated: changed };

  } catch (e) {
    Logger.log('dp_update100JsonLinesWithDecisions_ error: ' + (e && e.message ? e.message : e));
    return { ok:false, error: String(e) };
  }
}

// -------------------------------------------------------------
// Helpers: Stone Status join/remove + validation-safe write
// -------------------------------------------------------------

// Join tokens with ", " like Sheets' multi-select dropdowns. Dedup + preserve order.
function cd_mergeStoneStatus_(current, toAdd) {
  var cur = String(current || '').trim();
  var add = String(toAdd || '').trim();
  if (!cur) return add || '';
  var parts = cur.split(/\s*[;|,]\s*|\s*•\s*/g)
                 .map(function(s){ return s.trim(); })
                 .filter(Boolean);
  if (add && !parts.some(function(p){ return p.toLowerCase() === add.toLowerCase(); })) {
    parts.push(add);
  }
  return parts.join(', ');
}

// Remove a single token (case-insensitive) and re-join with ", "
function cd_removeStoneStatusToken_(current, removeToken) {
  var cur = String(current || '').trim();
  if (!cur) return '';
  var rm = String(removeToken || '').trim().toLowerCase();
  var parts = cur.split(/\s*[;|,]\s*|\s*•\s*/g)
                 .map(function(s){ return s.trim(); })
                 .filter(Boolean)
                 .filter(function(p){ return p.toLowerCase() !== rm; });
  return parts.join(', ');
}

/**
 * Validation-safe setter for a single cell:
 *  - capture current rule
 *  - clear validation
 *  - write value
 *  - restore original rule
 */
function cd_writeBypassValidation_(sheet, row, col, value) {
  var rng = sheet.getRange(row, col);
  var originalRule = rng.getDataValidation();  // may be null
  rng.clearDataValidations();
  rng.setValue(value);
  if (originalRule) rng.setDataValidation(originalRule);
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



