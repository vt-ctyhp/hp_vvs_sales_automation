/** UpdateQuotation_Settings_v1.gs — v1.4 (robust URL + Ring Size + debug)
*  Role: Upsert the Ring Setting Options table in a Quotation (client copy).
*  Source of truth: UI payload (records array) or "Load from tracker" baseline.
*
*  Finds the Ring Settings table:
*    1) Preferred named range QUOTE_SETTINGS_ANCHOR (top-left header cell).
*    2) Fallback: header scan (first 50 rows × 30 cols) for “Product / Style”.
*
*  Upsert key:
*    - Key = Product (case-insensitive). If blank, uses Style Detail as fallback.
*    - Updates matching row; otherwise appends to the first blank row in the block.
*
*  Logging:
*    - Toggle with Script Property UQ_DEBUG = TRUE|FALSE
*    - Helpers: uq_enableDebug(), uq_disableDebug(), uq_log_(), uq_err_()
*
*  Dependencies (optional; for reading Quotation URL and Tracker URL from 100_):
*    - dp_getActiveMasterRowContext_()
*    - dp_findHeaderIndex_()
*/

// ───────────────────────────── Debug toggles ─────────────────────────────
function uq_enableDebug(){ PropertiesService.getScriptProperties().setProperty('UQ_DEBUG','TRUE'); }
function uq_disableDebug(){ PropertiesService.getScriptProperties().setProperty('UQ_DEBUG','FALSE'); }
function uq_isDebug_() {
  try { return /^(TRUE|1|YES)$/i.test(String(PropertiesService.getScriptProperties().getProperty('UQ_DEBUG')||'FALSE')); }
  catch (_){ return false; }
}
function uq_log_(){ if (uq_isDebug_()) try { Logger.log.apply(Logger, arguments); } catch(_){ } }
function uq_err_(label, e) {
  const msg = (e && e.stack) ? e.stack : (e && e.message) ? e.message : String(e);
  Logger.log('❌ ' + label + ': ' + msg);
}

// ───────────────────────────── Constants ─────────────────────────────
/** Named range on Quotation file: first header cell of Ring Settings table */
const UQ_SETTINGS_ANCHOR_NAME = 'QUOTE_SETTINGS_ANCHOR';
/** Header scan width (columns from header to the right) */
const UQ_SETTINGS_SCAN_WIDTH = 20;
/** Heuristic header tokens to locate the Ring Settings header row when no named range exists */
const UQ_SETTINGS_HEADER_TOKENS = ['Product', 'Style', 'Style Detail', 'Ring Setting', 'Metal'];

/** Column aliases to find "Quotation URL" on 100_ */
var UQ100_ALIASES = { 'Quotation URL': ['Quotation URL','QuotationURL','Quote URL','QuoteURL'] };
const UQ_SETTINGS_QURL_ALIASES =
  (typeof UQ100_ALIASES !== 'undefined') ? UQ100_ALIASES
  : ({ 'Quotation URL': ['Quotation URL','QuotationURL','Quote URL','QuoteURL'] });

/** Column name candidates in the Quotation header (we map whatever exists) */
const UQ_SETTINGS_COLS = {
  keyPrimary:   ['Product','Product Name','Item','SKU','Code'],
  keyFallback:  ['Style Detail','Style','Ring Style','Setting Style'],
  metal:        ['Metal','Metal Type'],
  bandWidth:    ['Band Width','Band Width (mm)','Width','Width (mm)'],
  ringSize:     ['Ring Size','US Size','Size'],
  freeUpgrade:  ['Free Upgrade','Bonus','Included Upgrade'],
  priceRetail:  ['Online Retailer Price','Retail Price','Competitor Price'],
  priceBEAT:    ['Brilliant Earth Price After Tax','BE Price After Tax','BE (After Tax)'],
  priceVVS:     ['VVS Price','Our Price'],
  savings:      ['Your Savings!','Savings'],
  link:         ['Link','URL','Product URL']
};

// ───────────────────────────── Menu opener ─────────────────────────────
function uq_openUpdateQuotationSettings() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService
    .createHtmlOutputFromFile('dlg_update_quote_settings_v1') // <— EXACT NAME
    .setWidth(1040)
    .setHeight(700)
    .setTitle('Update Quotation — Ring Settings');
  ui.showModalDialog(html, '🧾 Update Quotation — Ring Settings');
}

// ───────────────────────────── Bootstrap (100_) ─────────────────────────────
/** Read Quotation URL from active 100_ row (robust: rich link, display, HYPERLINK()). */
function uq_bootstrapSettings() {
  if (typeof dp_getActiveMasterRowContext_ !== 'function') {
    return { ok:false, reason:'NO_100_CTX' };
  }
  const ctx = dp_getActiveMasterRowContext_();
  uq_log_('[100_] Context row:', JSON.stringify({
    rowIndex: ctx.rowIndex,
    rootApptId: ctx.rootApptId,
    customer: ctx.customerName,
    brand: ctx.companyBrand,
    assignedRep: ctx.assignedRep
  }, null, 2));

  let qUrl = '';
  try {
    const colQ = (typeof dp_findHeaderIndex_ === 'function')
      ? dp_findHeaderIndex_(ctx.headerMap, UQ_SETTINGS_QURL_ALIASES['Quotation URL'], false)
      : -1;
    if (colQ > -1) {
      const rng = ctx.sheet.getRange(ctx.rowIndex, colQ);

      // 1) rich-text link
      try {
        const rt = rng.getRichTextValue && rng.getRichTextValue();
        qUrl = (rt && rt.getLinkUrl && rt.getLinkUrl()) || '';
      } catch(_){}

      // 2) plain URL in cell
      if (!qUrl) qUrl = String(rng.getDisplayValue() || '').trim();

      // 3) HYPERLINK("url","label") fallback
      if (!qUrl) {
        const f = (rng.getFormula && rng.getFormula()) || '';
        const m = f.match(/^\s*=\s*HYPERLINK\s*\(\s*"([^"]+)"/i) ||
                  f.match(/^\s*=\s*HYPERLINK\s*\(\s*'([^']+)'/i);
        if (m && m[1]) qUrl = m[1].trim();
      }
    }
  } catch (e) {
    uq_err_('uq_bootstrapSettings/QuotationURL', e);
  }
  uq_log_('[100_] Quotation URL resolved to:', qUrl || '(blank)');

  return {
    ok:true,
    quotationUrl: qUrl || '',
    rootApptId: ctx.rootApptId,
    customerName: ctx.customerName,
    brand: ctx.companyBrand || '',
    assignedRep: ctx.assignedRep || ''
  };
}

// ───────────────────────────── Submit (server) ─────────────────────────────
/**
* Upsert Ring Setting Options in the client’s Quotation.
* payload = { quotationUrl, records:[{product,styleDetail,metal,bandWidth,ringSize, ...}] }
*/
function uq_submitUpdateQuotationSettings(payload) {
  const lock = LockService.getDocumentLock(); lock.waitLock(28 * 1000);
  try {
    uq_log_('== uq_submitUpdateQuotationSettings ==\nPayload:', JSON.stringify(payload || {}, null, 2));
    if (!payload || !payload.quotationUrl) throw new Error('Quotation URL is missing on 100_.');
    if (!Array.isArray(payload.records) || !payload.records.length) throw new Error('No setting records to upsert.');

    const result = uq_writeSettingsToQuote_(payload.quotationUrl, payload.records);
    uq_log_('Settings write result:', JSON.stringify(result, null, 2));

    return {
      ok: true,
      ...result,
      message: `Settings updated: ${result.updated} updated, ${result.appended} added.`
    };
  } catch (e) {
    uq_err_('uq_submitUpdateQuotationSettings', e);
    throw e;
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

// ───────────────────────────── Writer (Quotation) ─────────────────────────────
function uq_writeSettingsToQuote_(quotationUrl, records) {
  uq_log_('[QUOTE] settings → URL:', quotationUrl);

  // Allow ID or URL
  const tryId = uq_extractFileId_(quotationUrl);
  const ssQuote = tryId ? SpreadsheetApp.openById(tryId) : SpreadsheetApp.openByUrl(quotationUrl);
  uq_log_('[QUOTE] Opened file:', ssQuote.getName(), 'ID:', ssQuote.getId());

  // 1) Anchor lookup (named range preferred)
  let rc = null, sh = null, headerRow = -1, headerCol = -1, usedAnchor = false;
  try { rc = ssQuote.getRangeByName(UQ_SETTINGS_ANCHOR_NAME); } catch(_){}
  if (rc) {
    sh = rc.getSheet();
    headerRow = rc.getRow();
    headerCol = rc.getColumn();
    usedAnchor = true;
    uq_log_('[QUOTE] Named range found:', UQ_SETTINGS_ANCHOR_NAME, '→', sh.getName()+'!'+rc.getA1Notation().replace(/^.*!/,''));
  } else {
    // 2) Fallback header scan
    const sheets = ssQuote.getSheets();
    outer:
    for (let i=0;i<sheets.length;i++){
      const s = sheets[i];
      const r = Math.min(50, s.getLastRow());
      const c = Math.min(30, s.getLastColumn());
      if (r < 1 || c < 1) continue;
      const data = s.getRange(1,1,r,c).getDisplayValues();
      for (let rr=0; rr<r; rr++){
        for (let cc=0; cc<c; cc++){
          const cell = String(data[rr][cc] || '');
          if (!cell) continue;
          const hit = UQ_SETTINGS_HEADER_TOKENS.some(k => cell.toLowerCase().indexOf(k.toLowerCase()) >= 0);
          if (hit) { sh = s; headerRow = rr+1; headerCol = cc+1; break outer; }
        }
      }
    }
    uq_log_('[QUOTE] Heuristic header:', sh ? `${sh.getName()} R${headerRow}C${headerCol}` : 'NOT FOUND');
  }
  if (!sh || headerRow < 1 || headerCol < 1) {
    throw new Error(`Could not locate Ring Settings header. Add named range “${UQ_SETTINGS_ANCHOR_NAME}” to the first header cell.`);
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const width   = Math.min(UQ_SETTINGS_SCAN_WIDTH, (lastCol - headerCol + 1));
  const height  = Math.max(0, lastRow - headerRow); // (kept identical to existing behavior)
  uq_log_('[QUOTE] Window → sheet:', sh.getName(), 'headerRow:', headerRow, 'headerCol:', headerCol, 'height:', height, 'width:', width, 'anchor?', usedAnchor);

  const rangeBlock = sh.getRange(headerRow, headerCol, height, width);
  const grid = rangeBlock.getDisplayValues();
  const hdr  = (grid[0] || []).map(h => String(h || '').trim());
  uq_log_('[QUOTE] Settings header row:', JSON.stringify(hdr));

  // flexible header matcher
  const colIdx = (nameCandidates) => {
    for (const k of nameCandidates) {
      const needle = k.toLowerCase().replace(/\s+/g,'');
      for (let j=0;j<hdr.length;j++){
        const h = hdr[j].toLowerCase().replace(/\s+/g,'');
        if (h.indexOf(needle) >= 0) return j;
      }
    }
    return -1;
  };

  // Build column map
  const IDX = {
    keyPrimary:  colIdx(UQ_SETTINGS_COLS.keyPrimary),
    keyFallback: colIdx(UQ_SETTINGS_COLS.keyFallback),
    metal:       colIdx(UQ_SETTINGS_COLS.metal),
    bandWidth:   colIdx(UQ_SETTINGS_COLS.bandWidth),
    ringSize:    colIdx(UQ_SETTINGS_COLS.ringSize),
    freeUpgrade: colIdx(UQ_SETTINGS_COLS.freeUpgrade),
    priceRetail: colIdx(UQ_SETTINGS_COLS.priceRetail),
    priceBEAT:   colIdx(UQ_SETTINGS_COLS.priceBEAT),
    priceVVS:    colIdx(UQ_SETTINGS_COLS.priceVVS),
    savings:     colIdx(UQ_SETTINGS_COLS.savings),
    link:        colIdx(UQ_SETTINGS_COLS.link)
  };
  if (IDX.keyPrimary < 0 && IDX.keyFallback < 0) {
    throw new Error('Quotation Ring Settings header must include a “Product” OR “Style Detail” (or similar) column.');
  }

  // Index existing rows
  const keyForRow = (rowArr) => {
    const p = IDX.keyPrimary >= 0 ? String(rowArr[IDX.keyPrimary] || '').trim() : '';
    const f = IDX.keyFallback >= 0 ? String(rowArr[IDX.keyFallback] || '').trim() : '';
    const k = p || f;
    return k ? k.toLowerCase() : '';
  };

  const write = rangeBlock.getValues();
  const existing = new Map();
  let firstBlankOffset = -1;
  for (let r=1; r<grid.length; r++){
    const k = keyForRow(grid[r]);
    if (!k) { if (firstBlankOffset < 0) firstBlankOffset = r; continue; }
    existing.set(k, r); // r is row offset within block
  }
  uq_log_('[QUOTE] existing settings:', existing.size, 'firstBlankOffset:', firstBlankOffset);

  // Upsert
  let updated = 0, appended = 0;
  const putAt = (rowOffset, idx, val) => { if (idx >= 0 && val != null && val !== '') write[rowOffset][idx] = val; };

  records.forEach(rec => {
    const key = (String(rec.product || '').trim() || String(rec.styleDetail || '').trim()).toLowerCase();
    if (!key) return; // skip rows with no key at all

    const out = {
      keyPrimary:   rec.product || '',
      keyFallback:  rec.styleDetail || '',
      metal:        rec.metal || '',
      bandWidth:    rec.bandWidth || '',
      ringSize:     rec.ringSize || '',
      freeUpgrade:  rec.freeUpgrade || '',
      priceRetail:  rec.onlineRetailerPrice || '',
      priceBEAT:    rec.brilliantEarthPriceAfterTax || '',
      priceVVS:     rec.vvsPrice || '',
      savings:      rec.yourSavings || '',
      link:         rec.link || ''
    };

    if (existing.has(key)) {
      const rr = existing.get(key);
      putAt(rr, IDX.keyPrimary,  out.keyPrimary);
      putAt(rr, IDX.keyFallback, out.keyFallback);
      putAt(rr, IDX.metal,       out.metal);
      putAt(rr, IDX.bandWidth,   out.bandWidth);
      putAt(rr, IDX.ringSize,    out.ringSize);
      putAt(rr, IDX.freeUpgrade, out.freeUpgrade);
      putAt(rr, IDX.priceRetail, out.priceRetail);
      putAt(rr, IDX.priceBEAT,   out.priceBEAT);
      putAt(rr, IDX.priceVVS,    out.priceVVS);
      putAt(rr, IDX.savings,     out.savings);
      putAt(rr, IDX.link,        out.link);
      updated++;
      uq_log_('↻ update setting key=', key, 'at grid offset', rr);
    } else {
      // Append within current block if a blank exists
      let rr = -1;
      for (let r=1; r<write.length; r++){
        const isBlank = keyForRow(write[r]) === '';
        if (isBlank) { rr = r; break; }
      }
      if (rr >= 0) {
        putAt(rr, IDX.keyPrimary,  out.keyPrimary);
        putAt(rr, IDX.keyFallback, out.keyFallback);
        putAt(rr, IDX.metal,       out.metal);
        putAt(rr, IDX.bandWidth,   out.bandWidth);
        putAt(rr, IDX.ringSize,    out.ringSize);
        putAt(rr, IDX.freeUpgrade, out.freeUpgrade);
        putAt(rr, IDX.priceRetail, out.priceRetail);
        putAt(rr, IDX.priceBEAT,   out.priceBEAT);
        putAt(rr, IDX.priceVVS,    out.priceVVS);
        putAt(rr, IDX.savings,     out.savings);
        putAt(rr, IDX.link,        out.link);
        appended++;
        uq_log_('＋ append setting key=', key, 'at grid offset', rr);
      } else {
        // Extend by one row after current block
        const newR = headerRow + grid.length;
        sh.insertRowsAfter(headerRow + (grid.length - 1), 1);
        const rowVals = new Array(width).fill('');
        if (IDX.keyPrimary   >= 0) rowVals[IDX.keyPrimary]   = out.keyPrimary;
        if (IDX.keyFallback  >= 0) rowVals[IDX.keyFallback]  = out.keyFallback;
        if (IDX.metal        >= 0) rowVals[IDX.metal]        = out.metal;
        if (IDX.bandWidth    >= 0) rowVals[IDX.bandWidth]    = out.bandWidth;
        if (IDX.ringSize     >= 0) rowVals[IDX.ringSize]     = out.ringSize;
        if (IDX.freeUpgrade  >= 0) rowVals[IDX.freeUpgrade]  = out.freeUpgrade;
        if (IDX.priceRetail  >= 0) rowVals[IDX.priceRetail]  = out.priceRetail;
        if (IDX.priceBEAT    >= 0) rowVals[IDX.priceBEAT]    = out.priceBEAT;
        if (IDX.priceVVS     >= 0) rowVals[IDX.priceVVS]     = out.priceVVS;
        if (IDX.savings      >= 0) rowVals[IDX.savings]      = out.savings;
        if (IDX.link         >= 0) rowVals[IDX.link]         = out.link;
        sh.getRange(newR, headerCol, 1, width).setValues([rowVals]);
        appended++;
        uq_log_('＋ append-by-extension setting key=', key, 'at sheet row', newR);
      }
    }
  });

  // Commit buffered block writes
  rangeBlock.setValues(write);
  uq_log_('[QUOTE] Settings commit: updated =', updated, 'appended =', appended);

  return {
    sheetName: sh.getName(),
    updated, appended, upserted: (updated + appended),
    tableSize: Math.max(0, height - 1),
    anchorFound: usedAnchor
  };
}

// ───────────────────────────── Utilities ─────────────────────────────
/** Robustly extract a Drive file ID from *any* Google URL or return '' */
function uq_extractFileId_(url) {
  const s = String(url || '').trim();
  if (!s) return '';
  let m;
  m = s.match(/\/d\/([-\w]{25,})/);          if (m) return m[1];     // .../d/<id>/
  m = s.match(/[?&]id=([-\w]{25,})/);        if (m) return m[1];     // ?id=<id>
  m = s.match(/[-\w]{25,}/);                 if (m) return m[0];     // last resort token
  return '';
}

// ───────────────────────────── Tracker helpers (Ring Size supported) ─────────────────────────────
/** Returns [{ id,label,rowIndex,tsISO,revision,ringStyle,metal,band,size }] for the 3D Tracker “Log”. */
function uq_getTrackerVersions() {
  var ctx = (typeof dp_getActiveMasterRowContext_ === 'function') ? dp_getActiveMasterRowContext_() : null;
  if (!ctx) throw new Error('Open 100_ and select a customer row.');

  var sh = ctx.sheet, H = ctx.headerMap, r = ctx.rowIndex;
  var c = (H.byExact && H.byExact['3D Tracker']) || (H.byNorm && H.byNorm['3dtracker']) || 0;
  if (!c) throw new Error('Column "3D Tracker" not found on 100_.');

  var rng = sh.getRange(r, c);
  var url = '';
  try {
    var rt = rng.getRichTextValue();
    url = (rt && rt.getLinkUrl && rt.getLinkUrl()) || '';
  } catch(_){}
  if (!url) url = String(rng.getDisplayValue() || '').trim();
  if (!url) throw new Error('No 3D Tracker URL on this row.');

  var fileId = uq_extractFileId_(url);
  if (!fileId) throw new Error('3D Tracker URL not recognized (need a direct link to the tracker spreadsheet).');

  var ss = SpreadsheetApp.openById(fileId);
  var shLog = ss.getSheetByName('Log') || ss.getSheetByName('3D Log') || ss.getSheetByName('3D Revision Log');
  if (!shLog) throw new Error('Tracker Log tab not found.');

  var lr = shLog.getLastRow(), lc = shLog.getLastColumn();
  if (lr < 2) return [];

  var headers = shLog.getRange(1,1,1,lc).getDisplayValues()[0].map(function(h){return String(h||'').trim();});
  var Hm = {}; headers.forEach(function(h,i){ if(h) Hm[h]=i; });

  function pick(names, row){ for (var i=0;i<names.length;i++){ var j=Hm[names[i]]; if (j!=null) return row[j]; } return ''; }

  var vals = shLog.getRange(2,1,lr-1,lc).getValues();
  var out = [], rowIndex = 2;
  for (var i=0; i<vals.length; i++, rowIndex++) {
    var row = vals[i];

    var ts = pick(['Timestamp','Date','Created At','Updated'], row);
    var rev = pick(['Revision #','Revision','Version'], row);

    var style = pick(['Ring Style','RingStyle'], row);
    var metal = pick(['Metal','Metal Type','Metal (Type)'], row);
    var rawBand = pick(['Band Width (mm)','BandWidthMM','Band Width'], row);
    var band = String(rawBand || '').trim();
    var bandClean = band.replace(/\s*mm\s*$/i, '');

    var size  = pick(['US Size','USSize'], row);

    var lbl = Utilities.formatString(
      '%s%s — %s%s%s',
      (rev ? ('Rev ' + rev + ' • ') : ''),
      (ts instanceof Date ? Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(ts||'')),
      String(style||'Setting'),
      (metal ? (' • ' + metal) : ''),
      (bandClean ? (' • ' + bandClean + ' mm') : '')
    );

    out.push({
      id: String(rowIndex),
      label: lbl,
      rowIndex: rowIndex,
      tsISO: (ts instanceof Date) ? ts.toISOString() : String(ts||''),
      revision: String(rev||''),
      ringStyle: String(style||''),
      metal: String(metal||''),
      band: String(band||''),
      size: String(size||'')
    });
  }
  out.sort(function(a,b){ return String(b.rowIndex).localeCompare(String(a.rowIndex)); });
  uq_log_('[Tracker] versions found:', out.length);
  return out;
}

/** Build a baseline Ring Setting record from a chosen Log row id. */
function uq_getSettingFromTrackerRow(rowId) {
  rowId = String(rowId||'').trim();
  if (!rowId) throw new Error('Missing tracker row id.');

  var ctx = (typeof dp_getActiveMasterRowContext_ === 'function') ? dp_getActiveMasterRowContext_() : null;
  if (!ctx) throw new Error('Open 100_ and select a customer row.');

  var sh = ctx.sheet, H = ctx.headerMap, r = ctx.rowIndex;
  var c = (H.byExact && H.byExact['3D Tracker']) || (H.byNorm && H.byNorm['3dtracker']) || 0;
  if (!c) throw new Error('Column "3D Tracker" not found on 100_.');

  var rng = sh.getRange(r, c);
  var url = '';
  try { var rt = rng.getRichTextValue(); url = (rt && rt.getLinkUrl && rt.getLinkUrl()) || ''; } catch(_){}
  if (!url) url = String(rng.getDisplayValue() || '').trim();
  var fileId = uq_extractFileId_(url);
  if (!fileId) throw new Error('Invalid 3D Tracker URL');

  var ss = SpreadsheetApp.openById(fileId);
  var shLog = ss.getSheetByName('Log') || ss.getSheetByName('3D Log') || ss.getSheetByName('3D Revision Log');
  if (!shLog) throw new Error('Tracker Log tab not found.');

  var lr = shLog.getLastRow(), lc = shLog.getLastColumn();
  var idx = Number(rowId); if (!(idx>=2 && idx<=lr)) throw new Error('Tracker row not found: ' + rowId);

  var headers = shLog.getRange(1,1,1,lc).getDisplayValues()[0].map(function(h){return String(h||'').trim();});
  var Hm = {}; headers.forEach(function(h,i){ if(h) Hm[h]=i; });
  function get(names, row){ for (var i=0;i<names.length;i++){ var j=Hm[names[i]]; if (j!=null) return row[j]; } return ''; }

  var row = shLog.getRange(idx, 1, 1, lc).getDisplayValues()[0];

  var style = get(['Ring Style','RingStyle'], row);
  var metal = get(['Metal','Metal Type','Metal (Type)'], row);
  var rawBand = get(['Band Width (mm)','BandWidthMM','Band Width'], row);
  var bandClean = String(rawBand || '').trim().replace(/\s*mm\s*$/i, '');
  var size  = get(['US Size','USSize'], row);
  var product = get(['Product','Setting Name','Design Name'], row);

  // Compose baseline
  return {
    product: String(product || (style ? (style + ' Setting') : '')),
    styleDetail: String(style || ''),
    metal: String(metal || ''),
    bandWidth: bandClean,
    ringSize: String(size || ''),
    freeUpgrade: '',
    onlineRetailerPrice: '',
    brilliantEarthPriceAfterTax: '',
    vvsPrice: '',
    yourSavings: '',
    link: ''
  };
}

function _test_bootstrap_Settings() {
  Logger.log(JSON.stringify(uq_bootstrapSettings(), null, 2));
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



