/** File: report_server.gs (v1.1)
*  Role: Reporting (By Status / By Rep) + dropdown lists + export.
*  Changes (v1.1):
*    - Booked Appointment (Sales Stage = "Appointment"): omit "Order Total" and "Paid-to-Date" columns.
*    - Highlight blank Assigned Rep / Assisted Rep cells in light yellow (#fff9c4) in PDF and sheet.
*/

// Public helpers used by menus and Client Status
function report_getDropdownLists_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Dropdown');
  if (!sh) throw new Error('Sheet not found: "Dropdown"');

  const values = sh.getDataRange().getValues();
  if (!values.length) return {};
  const headers = values[0].map(String);
  const rows = values.slice(1);
  const dedup = (arr) => {
    const seen = new Set(), out = [];
    for (const v of arr) {
      const s = (v == null ? '' : String(v)).trim();
      if (!s || seen.has(s)) continue; seen.add(s); out.push(s);
    }
    return out;
  };
  const col = (name) => {
    const i = headers.indexOf(name);
    if (i === -1) return [];
    return dedup(rows.map(r => r[i]));
  };
  const assignedReps  = col('Assigned Rep');
  const assistedReps0 = col('Assisted Rep');
  const assistedReps  = assistedReps0.length ? assistedReps0 : assignedReps;

  return {
    salesStage:        col('Sales Stage'),
    conversionStatus:  col('Conversion Status'),
    customOrderStatus: col('Custom Order Status'),
    centerStoneStatus: col('Center Stone Order Status'),
    assignedReps,
    assistedReps,
  };
}


// ---- Menus call these openers (thin wrappers in v1_sales_menu also call them) ----
function report_ping(){ return 'pong'; }


// ===== Data read + cache =====
function report_getMasterData_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Sheet not found: "00_Master Appointments"');

  const key = 'report:master:v1:' + sh.getSheetId() + ':' + sh.getLastRow() + 'x' + sh.getLastColumn();
  const cached = report_cacheGet_(key);
  if (cached && cached.headers && cached.rows) return cached;

  const values  = sh.getDataRange().getValues();
  if (values.length < 2) return { headers: [], rows: [] };

  const headers = values[0].map(String);
  const rows    = values.slice(1);

  const payload = { headers, rows };
  report_cachePut_(key, payload, 60);
  return payload;
}
function report_headerIndex_(headers) {
  const H = {}; headers.forEach((h, i) => H[String(h)] = i); return H;
}

// Return the first matching header index from a list of possible names, or -1
function report_firstHeaderIndex_(headers, names){
  for (var i = 0; i < names.length; i++) {
    var j = headers.indexOf(names[i]);
    if (j !== -1) return j;
  }
  return -1;
}

// -------- Assisted Rep email → name map (from "Dropdown") with 10-min cache --------
function report_getAssistedEmailToNameMap_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Dropdown');
    if (!sh) return {};
    const values = sh.getDataRange().getValues();
    if (!values || values.length < 2) return {};

    const hdr = values[0].map(h => String(h || '').trim().toLowerCase());
    const iName  = hdr.indexOf('assisted rep');         // name column
    const iEmail = hdr.indexOf('assisted rep email');   // email column
    if (iEmail < 0) return {};

    const map = {};
    for (let r = 1; r < values.length; r++) {
      const email = String(values[r][iEmail] || '').trim().toLowerCase();
      if (!email) continue;
      const name = String(iName >= 0 ? (values[r][iName] || '') : '').trim();
      map[email] = name || email; // fallback to email if name blank
    }
    return map;
  } catch (_) {
    return {};
  }
}

function report_cachedAssistedMap_() {
  try {
    const cache = CacheService.getUserCache();
    const key = 'report:assistEmailMap:v1';
    const hit = cache.get(key);
    if (hit) { try { return JSON.parse(hit); } catch (_e) {} }
    const fresh = report_getAssistedEmailToNameMap_();
    try {
      const json = JSON.stringify(fresh);
      if (json.length <= 90000) cache.put(key, json, 600); // 10 minutes
    } catch (_e) {}
    return fresh;
  } catch (_e) {
    return report_getAssistedEmailToNameMap_();
  }
}


function report_shapeResult_(headers, rows) {
  const cols = [
    'APPT_ID','Customer Name','Assigned Rep','Brand','SO#',
    'Sales Stage','Conversion Status','Custom Order Status','Center Stone Order Status',
    'Next Steps','Client Status Report URL'
  ];
  const idx = cols.map(c => headers.indexOf(c));
  const shaped = rows.map(r => idx.map(i => (i >= 0 ? r[i] : '')));
  return { headers: cols, rows: shaped, total: rows.length, previewLimit: 1000 };
}

// Shape columns for the By Status report:
// puts Order Total + Paid-to-Date immediately AFTER Visit Date (unless opts.omitPaymentCols).
// ALSO: Append " (LastUpdatedByNameOrEmail)" into the Assisted Rep cell (no new columns).
// NEW: When includeProductionCols=true, insert "In Production Status" + "Production Deadline" AFTER "Custom Order Status".
function report_shapeStatusResult_(headers, rows, opts) {
  var includeProd = !!(opts && opts.includeProductionCols);
  var omitPayCols = !!(opts && opts.omitPaymentCols);

  var baseCols = [
    'APPT_ID','Customer Name','Assigned Rep','Assisted Rep','Brand','SO#','Visit Date',
    // ⬇️ we'll insert Order Total + Paid-to-Date right here (unless omitted)
    'Sales Stage','Conversion Status','Custom Order Status','Center Stone Order Status',
    'Next Steps','Client Status Report URL'
  ];

  // Source indices in Master
  var baseIdx      = baseCols.map(function(c){ return headers.indexOf(c); });
  var iAssistedSrc = headers.indexOf('Assisted Rep');

  // Last Updated By (email) — robust aliases
  var iUpdatedBy = report_firstHeaderIndex_(headers, [
    'Last Updated By','Updated By','Updated by','Last Updated by','UpdatedBy','Updated By (email)','Updated By Email'
  ]);

  // Payment columns
  var iConv   = headers.indexOf('Conversion Status');
  var iTotal  = report_firstHeaderIndex_(headers, ['Order Total']);
  var iPaidTo = report_firstHeaderIndex_(headers, [
    'Paid-to-Date','Paid to Date','Paid To Date','Total Pay To Date','Total Paid To Date'
  ]);

  // NEW: Production columns (source indices; robust aliases)
  var iProdStatus   = report_firstHeaderIndex_(headers, ['In Production Status','Production Status']);
  var iProdDeadline = report_firstHeaderIndex_(headers, ['Production Deadline','Production Due','Production Due Date','Est. Completion Date']);

  // Ack column (if present on master)
  var iAck = report_firstHeaderIndex_(headers, ['Ack Status']);

  // Build "today at 00:00" in script TZ for Days Since Last Update calc
  var tz       = Session.getScriptTimeZone();
  var todayYMD = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var todayMs  = new Date(todayYMD + 'T00:00:00').getTime();

  // Output headers + insert positions
  var outHeaders = baseCols.slice();

  // Insert payment columns after Visit Date (unless omitted)
  var visitPos = outHeaders.indexOf('Visit Date');
  var payInsertAt = (visitPos === -1) ? 7 : (visitPos + 1);
  if (!omitPayCols) {
    outHeaders.splice(payInsertAt, 0, 'Order Total', 'Total Pay To Date');
  }

  // Where Assisted Rep lands in the output
  var outAssistedIdx = outHeaders.indexOf('Assisted Rep');

  // Build the mapping once
  var emailToName = report_cachedAssistedMap_();

  // Optional: insert production columns AFTER "Custom Order Status"
  var prodInsertAt = -1;
  if (includeProd) {
    var cosIdx = outHeaders.indexOf('Custom Order Status');
    if (cosIdx !== -1) {
      prodInsertAt = cosIdx + 1;
      outHeaders.splice(prodInsertAt, 0, 'In Production Status', 'Production Deadline');
    }
  }

  // Finally, append Ack + Days at the far right
  outHeaders.push('Ack Status', 'Days Since Last Update');

  var shapedRows = rows.map(function(r){
    var rowOut = baseIdx.map(function(i){ return (i >= 0 ? r[i] : ''); });

    // Assisted Rep + (LastUpdatedByNameOrEmail)
    try {
      var assistedVal = (iAssistedSrc >= 0 ? String(r[iAssistedSrc] || '') : '').trim();
      var lupEmail    = (iUpdatedBy   >= 0 ? String(r[iUpdatedBy]   || '') : '').trim().toLowerCase();
      if (assistedVal && lupEmail && outAssistedIdx >= 0) {
        var disp = emailToName[lupEmail] || lupEmail; // map to name, fallback to email
        rowOut[outAssistedIdx] = assistedVal + ' (' + disp + ')';
      }
    } catch (_e) { /* noop */ }

    // Insert payment values after Visit Date (read values directly from Master when present)
    if (!omitPayCols) {
      var vTotal  = (iTotal  >= 0) ? r[iTotal]  : '';
      var vPayTo  = (iPaidTo >= 0) ? r[iPaidTo] : '';
      rowOut.splice(payInsertAt, 0, vTotal, vPayTo);
    }

    // NEW: Insert production columns (values) after Custom Order Status
    if (prodInsertAt !== -1) {
      var vProdStatus   = (iProdStatus   >= 0) ? r[iProdStatus]   : '';
      var vProdDeadline = (iProdDeadline >= 0) ? r[iProdDeadline] : '';
      rowOut.splice(prodInsertAt, 0, vProdStatus, vProdDeadline);
    }

    // Append Ack Status (from master if present) + Days Since Last Update (computed)
    var ack = (iAck >= 0) ? r[iAck] : '';
    var days = report_daysSinceFromRow_(headers, r, todayMs);
    rowOut.push(ack, days);

    return rowOut;
  });

  return { headers: outHeaders, rows: shapedRows, total: rows.length, previewLimit: 1000 };
}


// Shape columns for the By Rep report:
// puts Order Total + Paid-to-Date immediately AFTER Visit Date (unless opts.omitPaymentCols).
// ALSO: Append " (LastUpdatedByNameOrEmail)" into the Assisted Rep cell (no new columns).
// NEW: When includeProductionCols=true, insert "In Production Status" + "Production Deadline" AFTER "Custom Order Status".
function report_shapeRepsResult_(headers, rows, opts) {
  var includeProd = !!(opts && opts.includeProductionCols);
  var omitPayCols = !!(opts && opts.omitPaymentCols);

  var baseCols = [
    'APPT_ID','Customer Name','Assigned Rep','Assisted Rep','Brand','SO#','Visit Date',
    // ⬇️ we'll insert Order Total + Paid-to-Date right here (unless omitted)
    'Sales Stage','Conversion Status','Custom Order Status','Center Stone Order Status',
    'Next Steps','Client Status Report URL'
  ];

  var baseIdx      = baseCols.map(function(c){ return headers.indexOf(c); });
  var iAssistedSrc = headers.indexOf('Assisted Rep');

  var iUpdatedBy = report_firstHeaderIndex_(headers, [
    'Last Updated By','Updated By','Updated by','Last Updated by','UpdatedBy','Updated By (email)','Updated By Email'
  ]);

  var iConv   = headers.indexOf('Conversion Status');
  var iTotal  = report_firstHeaderIndex_(headers, ['Order Total']);
  var iPaidTo = report_firstHeaderIndex_(headers, [
    'Paid-to-Date','Paid to Date','Paid To Date','Total Pay To Date','Total Paid To Date'
  ]);

  // NEW: Production columns (source indices; robust aliases)
  var iProdStatus   = report_firstHeaderIndex_(headers, ['In Production Status','Production Status']);
  var iProdDeadline = report_firstHeaderIndex_(headers, ['Production Deadline','Production Due','Production Due Date','Est. Completion Date']);

  // Ack column (if present on master)
  var iAck = report_firstHeaderIndex_(headers, ['Ack Status']);

  // Build "today at 00:00" in script TZ for Days Since Last Update calc
  var tz       = Session.getScriptTimeZone();
  var todayYMD = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var todayMs  = new Date(todayYMD + 'T00:00:00').getTime();

  var outHeaders = baseCols.slice();
  var visitPos = outHeaders.indexOf('Visit Date');
  var payInsertAt = (visitPos === -1) ? 7 : (visitPos + 1);
  if (!omitPayCols) {
    outHeaders.splice(payInsertAt, 0, 'Order Total', 'Total Pay To Date');
  }

  var outAssistedIdx = outHeaders.indexOf('Assisted Rep');

  // Optional: insert production columns AFTER "Custom Order Status"
  var prodInsertAt = -1;
  if (includeProd) {
    var cosIdx = outHeaders.indexOf('Custom Order Status');
    if (cosIdx !== -1) {
      prodInsertAt = cosIdx + 1;
      outHeaders.splice(prodInsertAt, 0, 'In Production Status', 'Production Deadline');
    }
  }

  // Finally, append Ack + Days at the far right
  outHeaders.push('Ack Status', 'Days Since Last Update');

  var emailToName = report_cachedAssistedMap_();

  var shapedRows = rows.map(function(r){
    var rowOut = baseIdx.map(function(i){ return (i >= 0 ? r[i] : ''); });

    // Assisted Rep + (LastUpdatedByNameOrEmail)
    try {
      var assistedVal = (iAssistedSrc >= 0 ? String(r[iAssistedSrc] || '') : '').trim();
      var lupEmail    = (iUpdatedBy   >= 0 ? String(r[iUpdatedBy]   || '') : '').trim().toLowerCase();
      if (assistedVal && lupEmail && outAssistedIdx >= 0) {
        var disp = emailToName[lupEmail] || lupEmail;
        rowOut[outAssistedIdx] = assistedVal + ' (' + disp + ')';
      }
    } catch (_e) { /* noop */ }

    if (!omitPayCols) {
      var vTotal  = (iTotal  >= 0) ? r[iTotal]  : '';
      var vPayTo  = (iPaidTo >= 0) ? r[iPaidTo] : '';
      rowOut.splice(payInsertAt, 0, vTotal, vPayTo);
    }

    // NEW: Insert production columns (values) after Custom Order Status
    if (prodInsertAt !== -1) {
      var vProdStatus   = (iProdStatus   >= 0) ? r[iProdStatus]   : '';
      var vProdDeadline = (iProdDeadline >= 0) ? r[iProdDeadline] : '';
      rowOut.splice(prodInsertAt, 0, vProdStatus, vProdDeadline);
    }

    // Append Ack Status (from master if present) + Days Since Last Update (computed)
    var ack = (iAck >= 0) ? r[iAck] : '';
    var days = report_daysSinceFromRow_(headers, r, todayMs);
    rowOut.push(ack, days);

    return rowOut;
  });

  return { headers: outHeaders, rows: shapedRows, total: rows.length, previewLimit: 1000 };
}


// ===== Filters =====
function report_tokenizeMulti_(v) {
  if (v == null) return [];
  return String(v).replace(/\u00A0/g, ' ').split(/[\n\r,;|•]+/g).map(s => s.trim()).filter(Boolean);
}

// Convert a value to a comparable timestamp (ms). Returns null if blank/invalid.
function report_toTime_(v) {
  if (v instanceof Date && !isNaN(v)) return v.getTime();
  if (v == null) return null;
  var s = String(v).trim();
  if (!s) return null;
  var d = new Date(s);
  return isNaN(d) ? null : d.getTime();
}

// Compute integer days since "last update" for a given row.
// Priority: Updated At → Booked At (ISO). Returns '' if both are blank/invalid.
// Comparison is against "today at 00:00" in script TZ (to avoid off-by-one).
function report_daysSinceFromRow_(headers, row, todayMs) {
  var iUpdatedAt = report_firstHeaderIndex_(headers, [
    'Updated At','Last Updated At','Last Updated','UpdatedAt','Updated At (ISO)'
  ]);
  var iBookedIso = report_firstHeaderIndex_(headers, [
    'Booked At (ISO)','Booked At','Appointment Booked At (ISO)','Appt Booked (ISO)'
  ]);

  var t = (iUpdatedAt >= 0) ? report_toTime_(row[iUpdatedAt]) : null;
  if (t == null && iBookedIso >= 0) t = report_toTime_(row[iBookedIso]);
  if (t == null) return '';

  var oneDay = 24 * 60 * 60 * 1000;
  var days = Math.floor((todayMs - t) / oneDay);
  if (!isFinite(days)) return '';
  return Math.max(0, days);
}


// yyyy-MM-dd for neat preview display (preserves typed Date for export).
function report_formatDateYMD_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = (v == null ? '' : String(v)).trim();
  if (!s) return '';
  var d = new Date(s);
  return isNaN(d) ? s : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// Collapse rows to only the latest row per Root Appt ID (by most recent Visit Date).
// Robust header matching for the root column; falls back to APPT_ID if root not found.
// If Visit Date is blank/invalid for all rows in a group, chooses the last seen row.
function report_pickLatestRowsByRootApptId_(headers, rows) {
  // Try to find a "root appointment id" column with robust aliases
  var iRoot = report_firstHeaderIndex_(headers, [
    'Root Appt ID','Root Appointment ID','Root ApptID','RootApptID','rootapptid','ROOTAPPTID','Appt Root ID','APPT_ROOT'
  ]);
  // Fallback: group by APPT_ID (no dedupe if APPT_IDs are unique)
  if (iRoot < 0) iRoot = headers.indexOf('APPT_ID');

  // Visit Date index (for recency)
  var iVisit = headers.indexOf('Visit Date');

  // Nothing we can do if neither exists
  if (iRoot < 0) return rows;

  var byKey = new Map();

  for (var idx = 0; idx < rows.length; idx++) {
    var r = rows[idx];
    var key = String(r[iRoot] == null ? '' : r[iRoot]).trim();
    // If there is still no key, treat this row as its own group to avoid dropping data
    if (!key) key = '__row__' + idx;

    var t = (iVisit >= 0) ? report_toTime_(r[iVisit]) : null;
    var cand = { row: r, t: (t == null ? -Infinity : t), idx: idx };

    var prev = byKey.get(key);
    if (!prev || (cand.t > prev.t) || (cand.t === prev.t && cand.idx > prev.idx)) {
      byKey.set(key, cand);
    }
  }

  var out = [];
  byKey.forEach(function(v){ out.push(v.row); });
  return out;
}


function report_runStatus(payload) {
  const data = report_getMasterData_();
  const H = report_headerIndex_(data.headers);

  // 1) Collapse to latest row per rootapptid first (applies to ALL reports)
  const latestRows = report_pickLatestRowsByRootApptId_(data.headers, data.rows);

  // 2) Then apply status filters to those latest rows
  const sets = {
    salesStage:        new Set(payload?.salesStage || []),
    conversionStatus:  new Set(payload?.conversionStatus || []),
    customOrderStatus: new Set(payload?.customOrderStatus || []),
    centerStoneStatus: new Set(payload?.centerStoneStatus || [])
  };

  const out = [];
  for (const r of latestRows) {
    const vStage  = String(r[H['Sales Stage']] || '').trim();
    const vConv   = String(r[H['Conversion Status']] || '').trim();
    const vCust   = String(r[H['Custom Order Status']] || '').trim();
    const vCenter = String(r[H['Center Stone Order Status']] || '').trim();

    const mStage  = sets.salesStage.size        ? sets.salesStage.has(vStage)      : true;
    const mConv   = sets.conversionStatus.size  ? sets.conversionStatus.has(vConv) : true;
    const mCust   = sets.customOrderStatus.size ? sets.customOrderStatus.has(vCust): true;
    const mCenter = sets.centerStoneStatus.size ? sets.centerStoneStatus.has(vCenter) : true;

    if (mStage && mConv && mCust && mCenter) out.push(r);
  }

  // --- NEW: sort by Visit Date (oldest → newest) if column exists ---
  if (data.headers.indexOf('Visit Date') !== -1) {
    var vi = data.headers.indexOf('Visit Date');
    out.sort(function(a, b) {
      var ta = report_toTime_(a[vi]);
      var tb = report_toTime_(b[vi]);
      if (ta == null && tb == null) return 0;   // both blank
      if (ta == null) return 1;                 // blanks last
      if (tb == null) return -1;
      return ta - tb;                           // ascending
    });
  }

  // --- NEW: shape with Visit Date included ---
  var includeProd = Array.isArray(payload && payload.customOrderStatus)
    && payload.customOrderStatus.some(function(s){ return String(s).trim().toLowerCase() === 'in production'; });

  // Omit payment columns *only* for Booked Appointment (Sales Stage = "Appointment")
  var omitPay = Array.isArray(payload && payload.salesStage)
    && payload.salesStage.length === 1
    && String(payload.salesStage[0] || '').trim().toLowerCase() === 'appointment';

  var shaped  = report_shapeStatusResult_(data.headers, out, {
    includeProductionCols: includeProd,
    omitPaymentCols: omitPay
  });

  var summary = report_buildSummary_(data.headers, out);
  var warn    = out.length > 10000 ? 'Large result (>10,000 rows). Consider narrowing filters.' : '';

  // For preview, format the Visit Date nicely (export keeps typed Dates)
  if (payload && payload._mode === 'preview') {
    var dtIdx = shaped.headers.indexOf('Visit Date');
    var otIdx = shaped.headers.indexOf('Order Total');
    var ptIdx = shaped.headers.indexOf('Total Pay To Date');
    var pdIdx = shaped.headers.indexOf('Production Deadline'); // NEW

    for (var i = 0; i < shaped.rows.length; i++) {
      if (dtIdx !== -1) shaped.rows[i][dtIdx] = report_formatDateYMD_(shaped.rows[i][dtIdx]);
      if (otIdx !== -1) shaped.rows[i][otIdx] = report_fmtCurrency0_(shaped.rows[i][otIdx]);
      if (ptIdx !== -1) shaped.rows[i][ptIdx] = report_fmtCurrency0_(shaped.rows[i][ptIdx]);
      if (pdIdx !== -1) shaped.rows[i][pdIdx] = report_formatDateYMD_(shaped.rows[i][pdIdx]); // NEW
    }
    shaped.rows = shaped.rows.slice(0, Math.min(shaped.previewLimit || 1000, shaped.rows.length));
  }

  return Object.assign({}, shaped, { summary: summary, warn: warn });

}


function report_runReps(payload) {
  const data = report_getMasterData_();
  const H = report_headerIndex_(data.headers);

  // 1) Collapse to latest row per rootapptid first (applies to ALL reports)
  const latestRows = report_pickLatestRowsByRootApptId_(data.headers, data.rows);

  // 2) Then apply rep + status filters on those latest rows
  const assignedSet = new Set((payload?.assigned || []).map(String));
  const assistedSet = new Set((payload?.assisted || []).map(String));
  const anyRepSelected = (assignedSet.size + assistedSet.size) > 0;

  const stageSet   = new Set((payload?.salesStage || []).map(String));
  const convSet    = new Set((payload?.conversionStatus || []).map(String));
  const customSet  = new Set((payload?.customOrderStatus || []).map(String));
  const centerSet  = new Set((payload?.centerStoneStatus || []).map(String));

  const out = [];
  for (const r of latestRows) {
    // Rep matching (OR within assigned/assisted; applies only if any selected)
    const assignedTokens = report_tokenizeMulti_(r[H['Assigned Rep']]);
    const assistedTokens = report_tokenizeMulti_(r[H['Assisted Rep']]);
    const assignedMatch = assignedSet.size ? assignedTokens.some(t => assignedSet.has(t)) : false;
    const assistedMatch = assistedSet.size ? assistedTokens.some(t => assistedSet.has(t)) : false;
    const repMatch = anyRepSelected ? (assignedMatch || assistedMatch) : true;

    // Status filters (AND across lists; OR within each list)
    const vStage  = String(r[H['Sales Stage']] || '').trim();
    const vConv   = String(r[H['Conversion Status']] || '').trim();
    const vCust   = String(r[H['Custom Order Status']] || '').trim();
    const vCenter = String(r[H['Center Stone Order Status']] || '').trim();

    const stageMatch  = stageSet.size  ? stageSet.has(vStage)  : true;
    const convMatch   = convSet.size   ? convSet.has(vConv)    : true;
    const customMatch = customSet.size ? customSet.has(vCust)  : true;
    const centerMatch = centerSet.size ? centerSet.has(vCenter): true;

    if (repMatch && stageMatch && convMatch && customMatch && centerMatch) out.push(r);
  }

  // Sort by Visit Date (oldest → newest) when the column exists; blanks last.
  var vi = data.headers.indexOf('Visit Date');
  if (vi !== -1) {
    out.sort(function(a, b) {
      var ta = report_toTime_(a[vi]);
      var tb = report_toTime_(b[vi]);
      if (ta == null && tb == null) return 0;
      if (ta == null) return 1;   // blanks last
      if (tb == null) return -1;
      return ta - tb;             // ascending
    });
  }

  // Shape with Visit Date included.
  var includeProd = Array.from(customSet).some(function(v){ return String(v || '').trim().toLowerCase() === 'in production'; });

  // For rep report we keep payment columns by default (no change)
  var shaped  = report_shapeRepsResult_(data.headers, out, { includeProductionCols: includeProd, omitPaymentCols: false });

  var summary = report_buildSummary_(data.headers, out);
  var warn    = out.length > 10000 ? 'Large result (>10,000 rows). Consider narrowing filters.' : '';

  // For preview: present Visit Date as yyyy-MM-dd (export keeps typed Date).
  if (payload && payload._mode === 'preview') {
    var dtIdx = shaped.headers.indexOf('Visit Date');
    var otIdx = shaped.headers.indexOf('Order Total');
    var ptIdx = shaped.headers.indexOf('Total Pay To Date');
    var pdIdx = shaped.headers.indexOf('Production Deadline'); // NEW

    for (var i = 0; i < shaped.rows.length; i++) {
      if (dtIdx !== -1) shaped.rows[i][dtIdx] = report_formatDateYMD_(shaped.rows[i][dtIdx]);
      if (otIdx !== -1) shaped.rows[i][otIdx] = report_fmtCurrency0_(shaped.rows[i][otIdx]);
      if (ptIdx !== -1) shaped.rows[i][ptIdx] = report_fmtCurrency0_(shaped.rows[i][ptIdx]);
      if (pdIdx !== -1) shaped.rows[i][pdIdx] = report_formatDateYMD_(shaped.rows[i][pdIdx]); // NEW
    }
    shaped.rows = shaped.rows.slice(0, Math.min(shaped.previewLimit || 1000, shaped.rows.length));
  }
  return Object.assign({}, shaped, { summary: summary, warn: warn });
}


// ===== Export =====
function report_shortCriteria_(type, payload) {
  if (type === 'status') {
    const parts = [];
    for (const k of ['salesStage','conversionStatus','customOrderStatus','centerStoneStatus']) {
      if (Array.isArray(payload?.[k]) && payload[k].length) parts.push(payload[k].slice(0, 2).join('|'));
    }
    return parts.join(' • ') || 'All';
  }
  if (type === 'rep') {
    const A = (payload?.assigned || []).slice(0,2).join('|');
    const S = (payload?.assisted || []).slice(0,2).join('|');
    const parts = [];
    if (A) parts.push('A:' + A);
    if (S) parts.push('S:' + S);
    return parts.join(' • ') || 'All';
  }
  return 'All';
}
function report_buildSheetName_(type, payload) {
  const title = (type === 'status') ? 'By Status' : 'By Rep';
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH.mm');
  const crit = report_shortCriteria_(type, payload) || 'All';
  let name = `Report — ${title} — ${crit} — ${now}`;
  if (name.length <= 100) return name;
  const hash = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(payload))
  ).slice(0, 4);
  const compact = `Report — ${title} — ${now} — ${hash}`;
  if (compact.length <= 100) return compact;
  return compact.slice(0, 100);
}
// Build final filename (no spreadsheet title, no SO#, special Rep rule)
function report_buildFriendlyExportName_(type, payload, res) {
  // 1) Determine chosen conversion status (prefer explicit single selection)
  var chosenStatus = '';
  if (payload && Array.isArray(payload.conversionStatus) && payload.conversionStatus.length === 1) {
    chosenStatus = String(payload.conversionStatus[0] || '').trim();
  }
  // If not explicitly chosen, infer the most frequent Conversion Status from result rows
  if (!chosenStatus && res && res.headers && res.rows && res.rows.length) {
    var iConv = res.headers.indexOf('Conversion Status');
    if (iConv !== -1) {
      var counts = {};
      res.rows.forEach(function(r){ var k = String(r[iConv] || '').trim() || '(blank)'; counts[k] = (counts[k]||0)+1; });
      chosenStatus = Object.keys(counts).sort(function(a,b){ return (counts[b]-counts[a]) || String(a).localeCompare(String(b)); })[0] || 'All';
    }
  }
  // Fallback to Sales Stage / Custom Order Status when not set by Conversion
  if (!chosenStatus && payload && Array.isArray(payload.salesStage) && payload.salesStage.length === 1) {
    chosenStatus = String(payload.salesStage[0] || '').trim();
  }
  if (!chosenStatus && payload && Array.isArray(payload.customOrderStatus) && payload.customOrderStatus.length === 1) {
    chosenStatus = String(payload.customOrderStatus[0] || '').trim();
  }

  chosenStatus = chosenStatus || 'All';

  // Cosmetic label tweaks for filenames
  // e.g. Sales Stage "Appointment" should display as "Booked Appointment"
  var DISPLAY_MAP = {
    'Appointment': 'Booked Appointment'
  };
  if (DISPLAY_MAP[chosenStatus]) {
    chosenStatus = DISPLAY_MAP[chosenStatus];
  }

  // 2) Special statuses that should include "Rep###:" prefix
  var special = new Set(['deposit paid','in production','viewing scheduled']);
  var isSpecial = special.has(chosenStatus.toLowerCase());

  // 3) Rep prefix rule
  // Primary: allow a cosmetic override via payload.repLabel (e.g., "Rep101") — no data filtering implied.
  // Fallback: if not provided, and for special statuses, infer from data when ALL rows share one Assigned Rep.
  var repPrefix = '';
  var repLabel = (payload && payload.repLabel) ? String(payload.repLabel).trim() : '';

  if (repLabel) {
    repPrefix = repLabel + ': ';
  } else if (isSpecial) {
    var selectedRep = (payload && Array.isArray(payload.assigned) && payload.assigned.length === 1)
      ? String(payload.assigned[0] || '').trim()
      : '';
    if (!selectedRep && res && res.headers) {
      var iAssigned = res.headers.indexOf('Assigned Rep');
      if (iAssigned !== -1 && Array.isArray(res.rows) && res.rows.length) {
        var bag = new Set();
        res.rows.forEach(function(r){
          var v = String(r[iAssigned] || '').trim();
          if (v) bag.add(v);
        });
        if (bag.size === 1) selectedRep = Array.from(bag)[0];
      }
    }
    if (selectedRep) repPrefix = selectedRep + ': ';
  }

  // 4) Timestamp + user email
  var tz   = Session.getScriptTimeZone() || 'America/Los_Angeles';
  var ts   = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd h:mma');
  var user = report_getGeneratorEmail_() || '';

  // 5) Build final title — NOTE: no SO#/APPT_ID segment anymore
  // Desired patterns:
  //  - Special statuses: "Rep###: Deposit Paid - 2025-09-09 11_32AM (gen. by user@…)"
  //  - Others: "Follow Up - 2025-09-09 11_32AM (gen. by user@…)"
  var title = (repPrefix ? repPrefix : '') + chosenStatus + ' - ' + ts + (user ? (' (gen. by ' + user + ')') : '');

  // Ensure under sheet-name limit; PDF filename uses this verbatim.
  if (title.length > 100) title = title.slice(0, 100);
  return title;
}

function report_fetchPdfBlob_(ss, sh) {
  var url = report_buildPdfUrl_(ss, sh);
  var token = ScriptApp.getOAuthToken();
  var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  return resp.getBlob();
}

/** -------------------- HTML→PDF Export (no temp tab left; no Drive file) -------------------- **/

// Minimal HTML escaper for safe rendering
function report_htmlEscape_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function(m){
    return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[m];
  });
}

// Turn a cell into a safe string; URL → <a>Open</a>; format Visit Date
function report_cellToHtml_(header, value) {
  var h = String(header || '');
  var s = (value == null ? '' : value);

  // Date fields: normalize to yyyy-MM-dd
  if (/^(visit date|production deadline|production due|production due date|est\.?\s*completion date)$/i.test(h)) {
    return report_htmlEscape_(report_formatDateYMD_(s));
  }

  // Currency: Order Total / Total Pay To Date (render as $#,### no decimals)
  if (/^(order total|paid[- ]?to[- ]?date|total pay to date|total paid[- ]?to[- ]?date)$/i.test(h)) {
    return report_htmlEscape_(report_fmtCurrency0_(s));
  }

  // Client Status Report URL: show "Open" link if it's a URL
  if (/client status report url/i.test(h)) {
    var url = String(s || '').trim();
    if (/^https?:\/\//i.test(url)) {
      return '<a href="' + report_htmlEscape_(url) + '" target="_blank">Open</a>';
    }
    return '';
  }

  // Generic URL auto-link (optional; comment out if you only want the one above)
  var str = String(s);
  if (/^https?:\/\//i.test(str)) {
    return '<a href="' + report_htmlEscape_(str) + '" target="_blank">' + report_htmlEscape_(str) + '</a>';
  }

  return report_htmlEscape_(str);
}

// Short human summary of the filters (for the meta line above the table)
function report_buildFiltersSummary_(type, payload) {
  try {
    if (type === 'status') {
      var parts = [];
      var add = function(label, arr) {
        var a = Array.isArray(arr) ? arr.filter(Boolean) : [];
        parts.push(label + ': ' + (a.length ? report_htmlEscape_(a.join(', ')) : 'All'));
      };
      add('Sales Stage',        payload && payload.salesStage);
      add('Conversion',         payload && payload.conversionStatus);
      add('Custom Order',       payload && payload.customOrderStatus);
      add('Center Stone',       payload && payload.centerStoneStatus);
      return parts.join(' • ');
    }
    if (type === 'rep') {
      var A = (payload && Array.isArray(payload.assigned) && payload.assigned.length) ? payload.assigned.join(', ') : 'All';
      var S = (payload && Array.isArray(payload.assisted) && payload.assisted.length) ? payload.assisted.join(', ') : 'All';
      return 'Assigned: ' + report_htmlEscape_(A) + ' • Assisted: ' + report_htmlEscape_(S);
    }
  } catch (e) {}
  return '';
}

// Build the full HTML for the PDF (table headers/rows from shaped result)
function report_buildReportHtml_(type, payload, res, title) {
  var filtersMeta = report_buildFiltersSummary_(type, payload);
  var rowCount = (res && Array.isArray(res.rows)) ? res.rows.length : 0;

  // Build THEAD
  var thead = '<tr>' + res.headers.map(function(h){
    return '<th>' + report_htmlEscape_(h) + '</th>';
  }).join('') + '</tr>';

  // Build TBODY with overdue highlighting for Production Deadline when still In Production
  var idxCOS      = res.headers.indexOf('Custom Order Status');
  var idxPD       = res.headers.indexOf('Production Deadline');
  var idxAssigned = res.headers.indexOf('Assigned Rep');
  var idxAssist   = res.headers.indexOf('Assisted Rep');
  // NEW: Ack/Days indices for conditional highlight
  var idxAck      = res.headers.indexOf('Ack Status');
  var idxDays     = res.headers.indexOf('Days Since Last Update');

  // Only the Booked Appointment report (Sales Stage = "Appointment") should highlight blank reps
  var highlightBlankReps =
    (type === 'status') &&
    Array.isArray(payload && payload.salesStage) &&
    payload.salesStage.length === 1 &&
    String(payload.salesStage[0] || '').trim().toLowerCase() === 'appointment';

  // Build "today at 00:00" in script TZ (ms + yyyy-MM-dd string for fallback)
  var tz       = Session.getScriptTimeZone();
  var todayYMD = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var todayMs  = new Date(todayYMD + 'T00:00:00').getTime();

  function normStatus(x){
    return String(x == null ? '' : x).replace(/\u00A0/g,' ').trim().toLowerCase();
  }

  var tbody = (res.rows || []).map(function(r){
    var tds = '';
    var status = (idxCOS >= 0) ? normStatus(r[idxCOS]) : '';

    // PRECOMPUTE Ack/Days for this row (used to color the Ack cell)
    var _ackVal  = (idxAck  >= 0) ? String(r[idxAck]  || '').trim() : '';
    var _daysVal = (idxDays >= 0) ? Number(r[idxDays] || 0) : 0;
    // accept ASCII hyphen, NB hyphen, en/em dashes, or space
    var _ackNeedsFU = /needs\s*follow(?:[-\s\u2010-\u2015])?up/i.test(_ackVal) && _daysVal >= 3;

    for (var c = 0; c < res.headers.length; c++) {
      var h = res.headers[c];
      var v = r[c];

      var bgColor = null;
      var fontClr = null;

      // flag overdue: PD exists, status is In Production, and PD < today
      if (c === idxPD && idxPD >= 0 && status === 'in production') {
        var ms  = report_toTime_(v);         // null or epoch ms
        var ymd = report_formatDateYMD_(v);  // '' or 'yyyy-MM-dd'
        var overdue = (ms != null) ? (ms < todayMs)
                    : (ymd && ymd < todayYMD);
        if (overdue) {
          bgColor = '#ff3b30';
          fontClr = '#000000';
        }
      }

      // NEW: light-yellow highlight for blank Assigned/Assisted Rep in Booked Appointment report
      if (highlightBlankReps && (c === idxAssigned || c === idxAssist)) {
        var blank = (String(v == null ? '' : v).trim() === '');
        if (blank) bgColor = '#fff9c4';
      }

      // NEW: independent highlights (Ack Status & Days Since)
      // Ack Status → orange if "Needs follow-up" (hyphen tolerant)
      if (c === idxAck && idxAck >= 0) {
        var ackStr = String(v == null ? '' : v).trim();
        if (/needs\s*follow[-\s]?up/i.test(ackStr)) {
          bgColor = '#ffa726';
          fontClr = '#000000';
        }
      }

      // Days Since Last Update → orange if >= 3
      if (c === idxDays && idxDays >= 0) {
        var dnum = Number(v);
        if (isFinite(dnum) && dnum >= 3) {
          bgColor = '#ffa726';
          fontClr = '#000000';
        }
      }

      var tdStyle = '';
      var tdBgAttr = '';
      if (bgColor) {
        tdStyle  = ' style="background-color:' + bgColor + (fontClr ? ';color:' + fontClr : '') + ';"';
        tdBgAttr = ' bgcolor="' + bgColor + '"';
      }

      tds += '<td' + tdBgAttr + tdStyle + '>' + report_cellToHtml_(h, v) + '</td>';
    }

    return '<tr>' + tds + '</tr>';
  }).join('');

  // Simple, clean CSS similar to Appointment Summary
  var html =
  '<!doctype html>' +
  '<html><head><meta charset="utf-8">' +
  '<style>' +
  '  @page { size: Letter landscape; margin: 12mm; }' +
  '  /* Keep background colors in print/PDF */' +
  '  html,body,table,thead,tbody,tr,td,th{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }' +
  '  body { font:12px/1.45 system-ui,-apple-system,Segoe UI,Roboto,Arial; color:#111; }' +
  '  h1 { font-size:16px; margin:0 0 6px; }' +
  '  .meta { font-size:12px; color:#555; margin:0 0 10px; }' +
  '  table { width:100%; border-collapse:collapse; font-size:11px; }' +
  '  thead th { background:#f4f4f4; border:1px solid #ddd; text-align:left; padding:6px 8px; }' +
  '  tbody td { border:1px solid #e5e5e5; padding:6px 8px; vertical-align:top; }' +
  '</style></head><body>' +

  '<h1>' + report_htmlEscape_(title) + '</h1>' +
  '<div class="meta">' + report_htmlEscape_(filtersMeta) + (filtersMeta ? ' • ' : '') + 'Rows: ' + rowCount + '</div>' +
  '<table><thead>' + thead + '</thead><tbody>' + tbody + '</tbody></table>' +
  '</body></html>';

  return html;
}

/**
 * Export the current Status/Rep result as a PDF (HTML→PDF), returned as base64 bytes.
 * Called by HTML with google.script.run.report_exportPdf('status'|'rep', payload)
 * Returns: {ok:true, fileName, bytesBase64, mimeType} on success.
 */
function report_exportPdf(type, payload) {
  try {
    var res = (type === 'status')
      ? report_runStatus(Object.assign({}, payload || {}, {_mode: undefined}))
      : (type === 'rep')
      ? report_runReps(Object.assign({}, payload || {}, {_mode: undefined}))
      : (function(){ throw new Error('Unknown report type: ' + type); })();

    // NEW: shape sanity logging
    Logger.log('[exportPdf] type=%s rows=%s cols=%s',
               type, (res && res.rows && res.rows.length) || 0, (res && res.headers && res.headers.length) || 0);

    report_sortRowsByVisitDate_(res);

    var baseName = report_buildFriendlyExportName_(type, payload, res);
    var fileName = baseName + '.pdf';
    var html = report_buildReportHtml_(type, payload, res, baseName);

    // NEW: sizes going into PDF conversion
    Logger.log('[exportPdf] baseName="%s" htmlLen=%s', baseName, (html || '').length);

    var blob = Utilities.newBlob(html, 'text/html', 'report.html').getAs('application/pdf');

    // NEW: resulting blob meta
    Logger.log('[exportPdf] blobBytes=%s mime=%s', (blob && blob.getBytes && blob.getBytes().length) || -1, blob && blob.getContentType());

    return {
      ok: true,
      fileName: fileName,
      bytesBase64: Utilities.base64Encode(blob.getBytes()),
      mimeType: blob.getContentType()
    };
  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

// ===== Formatting + Summary =====
function report_applyExportFormatting_(sh, res) {
  var lastRow = sh.getLastRow();
  var lastCol = res.headers.length;
  if (lastRow < 1 || lastCol < 1) return;

  sh.setFrozenRows(1);
  var header = sh.getRange(1, 1, 1, lastCol);
  header.setFontWeight('bold').setBackground('#f5f5f5');

  sh.autoResizeColumns(1, lastCol);

  // Number format for currency columns
  var otIdx = res.headers.indexOf('Order Total');           // 0-based
  var ptIdx = res.headers.indexOf('Paid-to-Date');          // 0-based
  if (ptIdx === -1) ptIdx = res.headers.indexOf('Total Pay To Date'); // back-compat

  if (lastRow > 1) {
    if (otIdx !== -1) {
      sh.getRange(2, otIdx + 1, lastRow - 1, 1).setNumberFormat('$#,##0');
    }
    if (ptIdx !== -1) {
      sh.getRange(2, ptIdx + 1, lastRow - 1, 1).setNumberFormat('$#,##0');
    }
  }

  // Integer format for Days Since Last Update
  var dIdx = res.headers.indexOf('Days Since Last Update'); // 0-based
  if (dIdx !== -1 && lastRow > 1) {
    sh.getRange(2, dIdx + 1, lastRow - 1, 1).setNumberFormat('0');
  }

  // === Overdue highlighting for Production Deadline when still "In Production" ===
  var cosIdx = res.headers.indexOf('Custom Order Status');    // 0-based
  var pdIdx  = res.headers.indexOf('Production Deadline');    // 0-based
  if (lastRow > 1 && cosIdx !== -1 && pdIdx !== -1) {
    var nRows = lastRow - 1;

    var cosVals = sh.getRange(2, cosIdx + 1, nRows, 1).getValues(); // strings
    var pdVals  = sh.getRange(2, pdIdx  + 1, nRows, 1).getValues(); // dates or strings

    var bg = new Array(nRows);
    var fc = new Array(nRows);
    var fw = new Array(nRows);

    // "today at 00:00" in script TZ (ms + yyyy-MM-dd for fallback)
    var tz       = Session.getScriptTimeZone();
    var todayYMD = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    var todayMs  = new Date(todayYMD + 'T00:00:00').getTime();

    for (var i = 0; i < nRows; i++) {
      var status = String(cosVals[i][0] || '').trim().toLowerCase();
      var pv     = pdVals[i][0];

      // robust ms (typed Date or parsable string)
      var ms  = report_toTime_(pv);
      var ymd = (function(v){
        var s = (v == null ? '' : String(v)).trim();
        if (!s) return '';
        var d = new Date(s);
        return isNaN(d) ? s : Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      })(pv);

      var overdue = (ms != null) ? (ms < todayMs)
                  : (ymd && ymd < todayYMD);

      var hit = (status === 'in production') && overdue;

      bg[i] = [ hit ? '#ff3b30' : null ];
      fc[i] = [ hit ? '#000000' : null ];
      fw[i] = [ 'normal' ];
    }

    var pdRange = sh.getRange(2, pdIdx + 1, nRows, 1);
    pdRange.setBackgrounds(bg).setFontColors(fc).setFontWeights(fw);
    // ensure a readable date format
    pdRange.setNumberFormat('yyyy-mm-dd');
  }

  // Date format for Production Deadline (yyyy-mm-dd)
  var pdIdx2 = res.headers.indexOf('Production Deadline'); // 0-based
  if (pdIdx2 !== -1 && lastRow > 1) {
    sh.getRange(2, pdIdx2 + 1, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
  }

  // (Optional) Set reasonable column widths for the new columns
  var ipsIdx = res.headers.indexOf('In Production Status');
  if (ipsIdx !== -1) sh.setColumnWidth(ipsIdx + 1, 220);
  if (pdIdx2  !== -1) sh.setColumnWidth(pdIdx2  + 1, 140);

  // Ensure Next Steps column is wide enough for wrapping
  var nextIdx = res.headers.indexOf('Next Steps');
  if (nextIdx !== -1) {
    sh.setColumnWidth(nextIdx + 1, 320);
  }

  // ✅ Wrap all cells (instead of CLIP) and align top
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, lastCol)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setVerticalAlignment('top');
  }

  SpreadsheetApp.flush();

  // Fixed row height ~48px for consistency
  if (lastRow > 1) {
    var DATA_ROW_HEIGHT = 48;
    var start = 2, remain = lastRow - 1, BATCH = 500;
    while (remain > 0) {
      var n = Math.min(BATCH, remain);
      sh.setRowHeights(start, n, DATA_ROW_HEIGHT);
      start += n; remain -= n;
    }
    SpreadsheetApp.flush();
  }

  // NEW: light-yellow highlight for blank Assigned/Assisted Rep cells (all reports)
  if (lastRow > 1) {
    var idxAssigned = res.headers.indexOf('Assigned Rep');
    var idxAssist   = res.headers.indexOf('Assisted Rep');
    var nRows = lastRow - 1;

    if (idxAssigned !== -1) {
      var valsA = sh.getRange(2, idxAssigned + 1, nRows, 1).getDisplayValues();
      var bgsA  = valsA.map(function(row){ return [ String(row[0] || '').trim() ? null : '#fff9c4' ]; });
      sh.getRange(2, idxAssigned + 1, nRows, 1).setBackgrounds(bgsA);
    }
    if (idxAssist !== -1) {
      var valsS = sh.getRange(2, idxAssist + 1, nRows, 1).getDisplayValues();
      var bgsS  = valsS.map(function(row){ return [ String(row[0] || '').trim() ? null : '#fff9c4' ]; });
      sh.getRange(2, idxAssist + 1, nRows, 1).setBackgrounds(bgsS);
    }
  }

  // NEW: independent highlights (Ack Status & Days Since Last Update)
  var idxAck  = res.headers.indexOf('Ack Status');
  var idxDays = res.headers.indexOf('Days Since Last Update');

  if (lastRow > 1) {
    var n = lastRow - 1;

    // Ack Status → orange if "Needs follow-up"
    if (idxAck !== -1) {
      var ackVals = sh.getRange(2, idxAck + 1, n, 1).getDisplayValues();
      var bgsAck = new Array(n), fcsAck = new Array(n);
      for (var i = 0; i < n; i++) {
        var a = String(ackVals[i][0] || '').trim();
        var hitAck = /needs\s*follow[-\s]?up/i.test(a);
        bgsAck[i] = [ hitAck ? '#ffa726' : null ];
        fcsAck[i] = [ hitAck ? '#000000' : null ];
      }
      sh.getRange(2, idxAck + 1, n, 1).setBackgrounds(bgsAck).setFontColors(fcsAck);
    }

    // Days Since Last Update → orange if >= 3
    if (idxDays !== -1) {
      var daysVals = sh.getRange(2, idxDays + 1, n, 1).getValues();
      var bgsDays = new Array(n), fcsDays = new Array(n);
      for (var j = 0; j < n; j++) {
        var d = Number(daysVals[j][0] || 0);
        var hitDays = isFinite(d) && d >= 3;
        bgsDays[j] = [ hitDays ? '#ffa726' : null ];
        fcsDays[j] = [ hitDays ? '#000000' : null ];
      }
      sh.getRange(2, idxDays + 1, n, 1).setBackgrounds(bgsDays).setFontColors(fcsDays);
    }
  }

  var full = sh.getRange(1, 1, lastRow, lastCol);
  if (sh.getFilter()) sh.getFilter().remove();
  full.createFilter();
  full.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
}

function report_countMap_(arr) {
  const m = {}; for (const v0 of arr) {
    const v = (v0 == null ? '' : String(v0)).trim();
    const key = v || '(blank)'; m[key] = (m[key] || 0) + 1;
  } return m;
}
function report_buildSummary_(headers, rows) {
  const H = {}; headers.forEach((h, i) => H[h] = i);
  const salesStages=[], conversions=[], customs=[], centers=[], unifiedReps=[];
  for (const r of rows) {
    salesStages.push(r[H['Sales Stage']]);
    conversions.push(r[H['Conversion Status']]);
    customs.push(r[H['Custom Order Status']]);
    centers.push(r[H['Center Stone Order Status']]);
    const assigned = report_tokenizeMulti_(r[H['Assigned Rep']]);
    const assisted = report_tokenizeMulti_(r[H['Assisted Rep']]);
    const bag = new Set([...assigned, ...assisted]);
    if (bag.size === 0) bag.add('(blank)');
    for (const rep of bag) unifiedReps.push(rep);
  }
  return {
    totalRows: rows.length,
    groups: {
      rep:                   report_countMap_(unifiedReps),
      salesStage:            report_countMap_(salesStages),
      conversionStatus:      report_countMap_(conversions),
      customOrderStatus:     report_countMap_(customs),
      centerStoneStatus:     report_countMap_(centers),
    }
  };
}
function report_buildSummarySheetNameFromBase_(baseName) {
  const name1 = baseName + ' — Summary';
  if (name1.length <= 100) return name1;
  const name2 = 'Summary — ' + baseName;
  if (name2.length <= 100) return name2;
  return name2.slice(0, 100);
}
function report_writeSummarySheet_(ss, baseName, summary) {
  if (!summary || !summary.groups) return null;
  const name = report_buildSummarySheetNameFromBase_(baseName);
  const sh = ss.insertSheet(name);
  if (!sh) throw new Error('Failed to create summary sheet');

  const sections = [
    ['Rep (Assigned or Assisted)',      'rep'],
    ['Sales Stage',                     'salesStage'],
    ['Conversion Status',               'conversionStatus'],
    ['Custom Order Status',             'customOrderStatus'],
    ['Center Stone Order Status',       'centerStoneStatus'],
  ];

  let row = 1;
  for (const [label, key] of sections) {
    const map = summary.groups[key] || {};
    const entries = Object.keys(map).map(k => [k, map[k]]);
    entries.sort((a,b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])));
    sh.getRange(row, 1).setValue(label).setFontWeight('bold'); row += 1;
    sh.getRange(row, 1, 1, 2).setValues([['Value','Count']]).setFontWeight('bold').setBackground('#f5f5f5'); row += 1;
    if (entries.length) { sh.getRange(row, 1, entries.length, 2).setValues(entries); row += entries.length; }
    row += 1;
  }
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.autoResizeColumns(1, 2);
    sh.getRange(1, 1, lastRow, 2).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  }
  return { sheetName: name };
}


// ===== Small cache =====
function report_cache_()           { return CacheService.getUserCache(); }
function report_cacheGet_(key)     { try { const s = report_cache_().get(key); return s ? JSON.parse(s) : null; } catch(e) { return null; } }
function report_cachePut_(key, obj, ttlSec) {
  try {
    const json = JSON.stringify(obj);
    if (json.length <= 90000) report_cache_().put(key, json, ttlSec || 60);
  } catch (e) { /* noop */ }
}

function report_buildPdfUrl_(ss, sh) {
  var base = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export';
  var params = {
    format: 'pdf', gid: String(sh.getSheetId()),
    portrait: 'false', fitw: 'true', size: 'letter',
    sheetnames: 'false', printtitle: 'false', pagenumbers: 'false',
    gridlines: 'false', fzr: 'true',
    top_margin: '0.5', bottom_margin: '0.5', left_margin: '0.5', right_margin: '0.5'
  };
  var q = Object.keys(params).map(k => k + '=' + encodeURIComponent(params[k])).join('&');
  return base + '?' + q;
}

// Sort shaped rows by Visit Date ascending (oldest → newest); blanks last
function report_sortRowsByVisitDate_(res) {
  var col = res.headers.indexOf('Visit Date'); // 0-based
  if (col === -1) return res;                  // nothing to sort on

  function ts(v) {
    if (v == null || v === '') return Infinity;                     // blanks last
    if (Object.prototype.toString.call(v) === '[object Date]') {
      var t = v.getTime(); return isNaN(t) ? Infinity : t;
    }
    var d = new Date(v);
    return isNaN(d.getTime()) ? Infinity : d.getTime();
  }

  res.rows.sort(function (a, b) { return ts(a[col]) - ts(b[col]); });
  return res;
}

/**
 * Replace raw URLs in the named column with =HYPERLINK(url, "Open")
 * Safe on blanks/non-URLs. Writes real formulas.
 */
function report_setHyperlinkLabel_(sh, res, headerName, label) {
  var colIdx = res.headers.indexOf(headerName); // 0-based
  if (colIdx === -1) return;

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return; // no data rows

  var rng = sh.getRange(2, colIdx + 1, lastRow - 1, 1);
  var vals = rng.getValues();
  var out  = new Array(vals.length);

  for (var i = 0; i < vals.length; i++) {
    var url = vals[i][0] != null ? String(vals[i][0]).trim() : '';
    if (url && /^https?:\/\//i.test(url)) {
      var safeUrl = url.replace(/"/g, '""'); // escape quotes
      out[i] = ['=HYPERLINK("' + safeUrl + '","' + label + '")'];
    } else {
      out[i] = [''];
    }
  }
  rng.setValues(out);
}

function report_getGeneratorEmail_() {
  if (typeof REPORT_EXPORT_USER === 'string' && REPORT_EXPORT_USER.trim()) {
    return REPORT_EXPORT_USER.trim();
  }
  var u = '';
  try { u = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  if (!u) {
    try { u = Session.getEffectiveUser().getEmail() || ''; } catch (e2) {}
  }
  return u;
}

function report_fmtCurrency0_(v) {
  // Accepts number or string; strips non-digits; rounds to 0 decimals; adds $ and commas
  if (v == null || v === '') return '';
  var n = Number(String(v).replace(/[^\d.-]/g, ''));
  if (!isFinite(n)) return '';
  return '$' + Math.round(n).toLocaleString('en-US');
}

/** Common helper: export a Rep-filtered PDF (single Rep; one or more Conversion Statuses) and download it. */
function report_menu_exportRepPdf_(repId, conversionStatuses) {
  var ui = SpreadsheetApp.getUi();
  try {
    if (!repId) throw new Error('Missing Rep ID.');
    var payload = {
      assigned: [String(repId)],
      // Note: report_runReps ignores statuses by default; leaving here for future filter extension.
      conversionStatus: Array.isArray(conversionStatuses) ? conversionStatuses.slice() : [String(conversionStatuses)],
      _mode: undefined
    };
    var out = report_exportPdf('rep', payload);
    if (!out || !out.ok) throw new Error(out && out.error ? out.error : 'Export failed.');
    report_showDownloadDialog_(out.fileName, out.bytesBase64, out.mimeType || 'application/pdf');
  } catch (e) {
    ui.alert('Quick PDF Export — ' + String(repId), (e && e.message) ? e.message : String(e), ui.ButtonSet.OK);
  }
}

/** Small helper so all quick exports use the base64 PDF pipeline (no external URLs). */
function report_quickStatusPdfDownload_(titleOrRepLabel, statusFilters) {
  var payload = Object.assign({ repLabel: String(titleOrRepLabel || '') }, statusFilters || {});
  Logger.log('[quickStatus] payload=%s', JSON.stringify(payload));

  var out = report_exportPdf('status', payload);
  Logger.log('[quickStatus] export ok=%s file="%s" err=%s',
             !!(out && out.ok), out && out.fileName, out && out.error);

  if (!out || !out.ok) throw new Error(out && out.error ? out.error : 'Export failed.');
  report_showDownloadDialog_(out.fileName, out.bytesBase64, out.mimeType || 'application/pdf');
}


/** Show a tiny auto-download dialog from base64 PDF bytes (robust; no external URLs). */
function report_showDownloadDialog_(fileName, b64, mime) {
  var safeName = String(fileName || 'report.pdf').replace(/[\\/:*?"<>|]/g, '-');
  var html =
'<!doctype html><html><head><meta charset="utf-8"><style>' +
'body{font:12px/1.45 system-ui,-apple-system,Segoe UI,Roboto,Arial;margin:10px}' +
'#log{white-space:pre-wrap;background:#f8f8f8;border:1px solid #e5e5e5;padding:6px 8px;max-height:160px;overflow:auto;margin-top:8px}' +
'button{padding:6px 10px;margin-top:8px}' +
'</style></head><body>' +
'<div><b>Preparing your PDF…</b></div>' +
'<div id="log"></div><button id="btn" style="display:none">Download manually</button>' +
'<script>(function(){' +
'  var NAME='+JSON.stringify(safeName)+', B64='+JSON.stringify(b64||"")+', MIME='+JSON.stringify(mime||"application/pdf")+';' +
'  var logEl=null, btn=null;' +
'  function w(x){ try{ (logEl||(logEl=document.getElementById("log"))).textContent+=(x+"\\n"); }catch(_){}}' +
'  function start(){' +
'    try{' +
'      if(!document.body){ return setTimeout(start,30); }' +
'      var bc=atob(B64); var bytes=new Uint8Array(bc.length); for(var i=0;i<bc.length;i++){bytes[i]=bc.charCodeAt(i);}' +
'      var blob=new Blob([bytes],{type:MIME}); var url=URL.createObjectURL(blob);' +
'      var a=document.createElement("a"); a.href=url; a.download=NAME; document.body.appendChild(a);' +
'      try{ a.click(); w("Download started."); }catch(e){ w("Auto-click blocked: "+e.message); }' +
'      btn=document.getElementById("btn"); if(btn){ btn.style.display="inline-block"; btn.onclick=function(){ try{ a.click(); }catch(e){ alert(e.message); } }; }' +
'      setTimeout(function(){ try{ URL.revokeObjectURL(url); }catch(_){ } try{ google.script.host.close(); }catch(_){ } }, 1200);' +
'    }catch(e){ w("Error: "+(e&&e.message||e)); try{ alert("Download failed: "+(e&&e.message||e)); }catch(_){ } try{ google.script.host.close(); }catch(_){ } }' +
'  }' +
'  if(document.readyState==="loading"){ document.addEventListener("DOMContentLoaded", start); } else { setTimeout(start,0); }' +
'})();</script></body></html>';
  SpreadsheetApp.getUi()
    .showModalDialog(HtmlService.createHtmlOutput(html).setWidth(380).setHeight(180), 'PDF Export');
}


/**
 * Lightweight client->server logger.
 * Usage from HTML: google.script.run.report__clientLog_(level, where, msg, extra)
 */
function report__clientLog_(level, where, msg, extra) {
  var L = String(level || 'INFO').toUpperCase();
  var W = String(where || 'downloadDlg');
  var M = String(msg || '');
  var E = '';
  try { E = extra ? JSON.stringify(extra) : ''; } catch (_e) {}
  Logger.log('[%s][%s] %s %s', L, W, M, E);
}


/** Small helper so all quick exports use the base64 PDF pipeline (no external URLs). */
function report_quickStatusPdfDownload_(titleOrRepLabel, statusFilters) {
  var payload = Object.assign({ repLabel: String(titleOrRepLabel || '') }, statusFilters || {});
  Logger.log('[quickStatus] payload=%s', JSON.stringify(payload));

  var out = report_exportPdf('status', payload);
  Logger.log('[quickStatus] export ok=%s file="%s" err=%s',
             !!(out && out.ok), out && out.fileName, out && out.error);

  if (!out || !out.ok) throw new Error(out && out.error ? out.error : 'Export failed.');
  report_showDownloadDialog_(out.fileName, out.bytesBase64, out.mimeType || 'application/pdf');
}


/** Menu actions — one per quick export (By Status) */
function report_menu_export_BookedAppointment() {
  report_quickStatusPdfDownload_('Booked Appointment', { salesStage: ['Appointment'] });
}
function report_menu_export_ViewingScheduled() {
  report_quickStatusPdfDownload_('Viewing Scheduled', { conversionStatus: ['Viewing Scheduled'] });
}
function report_menu_export_DepositPaid() {
  report_quickStatusPdfDownload_('Deposit Paid', { conversionStatus: ['Deposit Paid'] });
}
function report_menu_export_InProduction() {
  report_quickStatusPdfDownload_('In Production', { customOrderStatus: ['In Production'] });
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
