// /** UpdateQuotation_v1.gs — v1.1 (instrumented)
//  *  Step 1: Menu + Diamonds dialog shell (200_-only read, no writing)
//  *  Reuses Diamonds_v1 helpers: dp_getActiveMasterRowContext_(), dp_get200Sheet_(),
//  *  dp_headerMapFor200_(), dp_findHeaderIndex_()
//  *
//  *  Logging:
//  *    - Toggle with Script Property UQ_DEBUG = TRUE|FALSE (helpers below)
//  *    - All major steps and decisions are logged with clear tags.
//  */

// // ──────────────────────────────────────────────────────────────────────────────
// // Debug logging helpers
// // ──────────────────────────────────────────────────────────────────────────────
// function uq_enableDebug(){ PropertiesService.getScriptProperties().setProperty('UQ_DEBUG','TRUE'); }
// function uq_disableDebug(){ PropertiesService.getScriptProperties().setProperty('UQ_DEBUG','FALSE'); }

// function uq_isDebug_() {
//   try {
//     const v = (PropertiesService.getScriptProperties().getProperty('UQ_DEBUG') || 'FALSE').toUpperCase();
//     return v === 'TRUE' || v === '1' || v === 'YES';
//   } catch (e) { return false; }
// }
// function uq_log_() {
//   if (!uq_isDebug_()) return;
//   try { Logger.log.apply(Logger, arguments); } catch (e) {}
// }
// function uq_err_(label, e) {
//   const msg = (e && e.stack) ? e.stack : (e && e.message) ? e.message : String(e);
//   Logger.log('❌ ' + label + ': ' + msg);
// }

// // ──────────────────────────────────────────────────────────────────────────────
// // 100_ helper: column aliases for Quotation URL
// // ──────────────────────────────────────────────────────────────────────────────
// var UQ100_ALIASES = {
//   'Quotation URL': ['Quotation URL','QuotationURL','Quote URL','QuoteURL']
// };

// // Named range anchor inside the Quotation file. (top-left header cell of Diamonds table)
// var UQ_ANCHOR_NAME_DIAMONDS = 'QUOTE_DIAMONDS_ANCHOR';

// // Fallback: if the named range isn’t found, look for any header cell that contains one of:
// var UQ_DIAMONDS_HEADER_ALIASES = ['Cert', 'Certificate', 'Cert #', 'Certificate No'];

// // ──────────────────────────────────────────────────────────────────────────────
// // Menu
// // ──────────────────────────────────────────────────────────────────────────────
// function uq_openUpdateQuotationDiamonds() {
//   var ui = SpreadsheetApp.getUi();
//   try {
//     uq_log_('== uq_openUpdateQuotationDiamonds() ==');
//     var bootstrap = uq_bootstrapDiamonds_(); // throws helpful errors
//     uq_log_('Bootstrap meta →', JSON.stringify(bootstrap && bootstrap.meta || {}, null, 2));
//     uq_log_('Bootstrap context →', JSON.stringify(bootstrap && bootstrap.context || {}, null, 2));
//     uq_log_('Bootstrap items count:', (bootstrap && bootstrap.items ? bootstrap.items.length : 0));

//     var t = HtmlService.createTemplateFromFile('dlg_update_quote_diamonds_v1');
//     t.bootstrap = bootstrap;
//     var html = t.evaluate().setWidth(1200).setHeight(720);
//     ui.showModalDialog(html, '🧾 Update Quotation — Diamonds');
//   } catch (e) {
//     uq_err_('uq_openUpdateQuotationDiamonds', e);
//     ui.alert('🧾 Update Quotation — Diamonds', (e && e.message ? e.message : String(e)) +
//       '\n\nOpen 100_ (00_Master Appointments) and select a customer row, then try again.', ui.ButtonSet.OK);
//   }
// }

// // ──────────────────────────────────────────────────────────────────────────────
// // Bootstrap: active 100_ context + stones from 200_ by RootApptID (read-only)
// // ──────────────────────────────────────────────────────────────────────────────
// function uq_bootstrapDiamonds_() {
//   uq_log_('== uq_bootstrapDiamonds_() ==');

//   // 100_ selection & context (rootApptId, customerName, brand, etc.)
//   var ctx = dp_getActiveMasterRowContext_();
//   uq_log_('[100_] row context:', JSON.stringify({
//     rowIndex: ctx.rowIndex, rootApptId: ctx.rootApptId, customer: ctx.customerName,
//     brand: ctx.companyBrand, assignedRep: ctx.assignedRep,
//     visitDate: ctx.visitDateStr, visitTime: ctx.visitTimeStr
//   }, null, 2));

//   // Read Quotation URL from the same 100_ row (best effort)
//   var qUrl = '';
//   try {
//     var colQ = dp_findHeaderIndex_(ctx.headerMap, UQ100_ALIASES['Quotation URL'], false);
//     if (colQ > -1) qUrl = String(ctx.sheet.getRange(ctx.rowIndex, colQ).getDisplayValue() || '').trim();
//     uq_log_('[100_] Quotation URL:', qUrl || '(blank)');
//   } catch (e) {
//     uq_err_('Read Quotation URL', e);
//   }

//   // 200_ target + header map (two-row combiner)
//   var target = dp_get200Sheet_();
//   var sh200  = target.sheet;
//   var hm200  = dp_headerMapFor200_(sh200);
//   uq_log_('[200_] workbook:', target.ss.getName(), 'tab:', target.tab, 'rows:', sh200.getLastRow(), 'cols:', sh200.getLastColumn());

//   // Resolve columns we’ll display
//   function cReq(k){ return dp_findHeaderIndex_(hm200, dp_aliases200_[k], true); }
//   var cRoot   = cReq('RootApptID');
//   var cOrder  = cReq('Order Status');
//   var cStoneS = cReq('Stone Status');
//   var cVendor = cReq('Vendor');
//   var cType   = cReq('Stone Type');
//   var cShape  = cReq('Shape');
//   var cCarat  = cReq('Carat');
//   var cColor  = cReq('Color');
//   var cClarity= cReq('Clarity');
//   var cLab    = cReq('LAB');
//   var cCert   = cReq('Certificate No');
//   var cCust   = cReq('Customer Name');
//   var cAppt   = cReq('Customer Appt Time & Date');
//   var cRep    = cReq('Assigned Rep');
//   var cBrand  = cReq('Company');

//   var items = [];
//   var last = sh200.getLastRow();
//   if (last >= 3) {
//     var data = sh200.getRange(3, 1, last - 2, sh200.getLastColumn()).getDisplayValues();
//     var want = String(ctx.rootApptId || '').trim();
//     for (var i = 0; i < data.length; i++) {
//       var row = data[i], rIdx = i + 3;
//       if (String(row[cRoot - 1] || '').trim() !== want) continue;

//       items.push({
//         rowIndex: rIdx, rootApptId: want,
//         // specs
//         vendor: row[cVendor - 1] || '',
//         stoneType: row[cType - 1] || '',
//         shape: row[cShape - 1] || '',
//         carat: row[cCarat - 1] || '',
//         color: row[cColor - 1] || '',
//         clarity: row[cClarity - 1] || '',
//         lab: row[cLab - 1] || '',
//         certNo: row[cCert - 1] || '',
//         // status
//         orderStatus: row[cOrder - 1] || '',
//         stoneStatus: row[cStoneS - 1] || '',
//         // preview
//         customerName: row[cCust - 1] || '',
//         apptTimeDate: row[cAppt - 1] || '',
//         assignedRep: row[cRep - 1] || '',
//         company: row[cBrand - 1] || ''
//       });
//     }
//   }
//   uq_log_('[200_] Matched stones for rootApptId', ctx.rootApptId, '→', items.length);

//   return {
//     context: {
//       rootApptId: ctx.rootApptId,
//       customerName: ctx.customerName,
//       visitAt: (ctx.visitDateStr || '') + (ctx.visitTimeStr ? (' ' + ctx.visitTimeStr) : ''),
//       companyBrand: ctx.companyBrand || '',
//       assignedRep: ctx.assignedRep || '',
//       quotationUrl: qUrl || ''
//     },
//     items: items,
//     meta: { spreadsheetUrl: target.ss.getUrl(), tab: target.tab }
//   };
// }

// // ──────────────────────────────────────────────────────────────────────────────
// // Public API — dialog submit
// // ──────────────────────────────────────────────────────────────────────────────
// function uq_submitUpdateQuotationDiamonds(payload) {
//   const lock = LockService.getDocumentLock();
//   lock.waitLock(28 * 1000);
//   try {
//     uq_log_('== uq_submitUpdateQuotationDiamonds() ==');
//     uq_log_('Payload:', JSON.stringify(payload || {}, null, 2));
//     if (!payload || !payload.rootApptId) throw new Error('Missing rootApptId.');
//     if (!payload.quotationUrl) throw new Error('Quotation URL not found on 100_.');
//     if (!Array.isArray(payload.rowIndexes) || payload.rowIndexes.length === 0) throw new Error('No stones selected.');

//     // Source of truth → 200_
//     const target = dp_get200Sheet_();
//     const sh200  = target.sheet;
//     const hm200  = dp_headerMapFor200_(sh200);
//     uq_log_('[200_] Using:', target.ss.getName(), 'tab:', target.tab);

//     const last   = sh200.getLastRow();
//     const lastCol= sh200.getLastColumn();
//     uq_log_('[200_] size rows×cols:', last, '×', lastCol);

//     // Resolve columns
//     const c = (k, req=true) => dp_findHeaderIndex_(hm200, dp_aliases200_[k], req);
//     const C = {
//       Vendor: c('Vendor'),
//       Type: c('Stone Type'),
//       Shape: c('Shape'),
//       Carat: c('Carat'),
//       Color: c('Color'),
//       Clarity: c('Clarity'),
//       LAB: c('LAB'),
//       Cert: c('Certificate No'),
//       OrderStatus: c('Order Status'),
//       StoneStatus: c('Stone Status'),
//       Customer: c('Customer Name'),
//       Appt: c('Customer Appt Time & Date'),
//       Rep: c('Assigned Rep'),
//       Company: c('Company'),
//       Root: c('RootApptID')
//     };

//     // Read selected rows in one batch
//     const rowsNeeded = payload.rowIndexes.map(Number).filter(r => r >= 3 && r <= last);
//     uq_log_('Selected 200_ rows:', JSON.stringify(rowsNeeded));
//     const rects = rowsNeeded.map(r => sh200.getRange(r, 1, 1, lastCol));
//     const vals  = rects.map(rng => rng.getDisplayValues()[0]);

//     // Normalize → records (key by certNo)
//     const recs = vals.map((row, i) => {
//       const certNo = String(row[C.Cert - 1] || '').trim();
//       if (!certNo) throw new Error('Row '+ rowsNeeded[i] + ': missing Certificate No (cannot upsert Quotation).');
//       return {
//         certNo:   certNo,
//         vendor:   row[C.Vendor - 1] || '',
//         stoneType:row[C.Type - 1] || '',
//         shape:    row[C.Shape - 1] || '',
//         carat:    row[C.Carat - 1] || '',
//         color:    row[C.Color - 1] || '',
//         clarity:  row[C.Clarity - 1] || '',
//         lab:      row[C.LAB - 1] || '',
//         orderStatus: String(row[C.OrderStatus - 1] || '').trim(),
//         stoneStatus: String(row[C.StoneStatus - 1] || '').trim(),
//         customerName: row[C.Customer - 1] || '',
//         apptTimeDate: row[C.Appt - 1] || '',
//         assignedRep:  row[C.Rep - 1] || '',
//         company:      row[C.Company - 1] || ''
//       };
//     });
//     uq_log_('Prepared records from 200_:', JSON.stringify(recs, null, 2));

//     const result = uq_writeDiamondsToQuote_(payload.quotationUrl, recs);
//     uq_log_('Write result:', JSON.stringify(result, null, 2));

//     return {
//       ok: true,
//       upserted: result.upserted,
//       appended: result.appended,
//       updated:  result.updated,
//       tableSize: result.tableSize,
//       sheetName: result.sheetName,
//       anchorFound: result.anchorFound,
//       message: 'Quotation updated: ' + result.updated + ' updated, ' + result.appended + ' added.'
//     };
//   } catch (e) {
//     uq_err_('uq_submitUpdateQuotationDiamonds', e);
//     throw e;
//   } finally {
//     try { lock.releaseLock(); } catch(e) {}
//   }
// }

// // ──────────────────────────────────────────────────────────────────────────────
// /** Core writer:
//  *  - Opens the Quotation by URL (the client’s **copy**, not the template)
//  *  - Looks for named range QUOTE_DIAMONDS_ANCHOR; if missing, scans headers heuristically
//  *  - Reads Diamonds table under the header, indexes by Cert#
//  *  - Upserts rows (update existing by Cert#, otherwise append)
//  *  - Logs every major decision
//  */
// // ──────────────────────────────────────────────────────────────────────────────
// function uq_writeDiamondsToQuote_(quotationUrl, records) {
//   uq_log_('== uq_writeDiamondsToQuote_() ==');
//   uq_log_('[QUOTE] URL:', quotationUrl);

//   // Open the live **Quotation file** for this client (from 100_ "Quotation URL")
//   const ssQuote = SpreadsheetApp.openByUrl(quotationUrl);
//   uq_log_('[QUOTE] Opened:', ssQuote.getName(), 'ID:', ssQuote.getId());
//   const sheets  = ssQuote.getSheets().map(s => s.getName());
//   uq_log_('[QUOTE] Tabs:', JSON.stringify(sheets));

//   // 1) Locate anchor (preferred: named range)
//   let anchor = null, rc = null;
//   try {
//     rc = ssQuote.getRangeByName(UQ_ANCHOR_NAME_DIAMONDS);
//     if (rc) {
//       anchor = [rc.getSheet().getName(), rc.getA1Notation().replace(/^.*!/, '')]; // ["Sheet", "B4"]
//       uq_log_('[QUOTE] Named range found:', UQ_ANCHOR_NAME_DIAMONDS, '→', rc.getA1Notation());
//     } else {
//       uq_log_('[QUOTE] Named range NOT found:', UQ_ANCHOR_NAME_DIAMONDS);
//     }
//   } catch (e) {
//     uq_err_('getRangeByName', e);
//   }

//   // 2) If no named range, heuristically find a header cell with "Cert"/"Certificate"
//   let sh = null, headerRow = -1, headerCol = -1;
//   if (anchor) {
//     sh = ssQuote.getSheetByName(anchor[0]);
//     headerRow = rc.getRow();
//     headerCol = rc.getColumn();
//     uq_log_('[QUOTE] Using anchor → sheet:', sh.getName(), 'row:', headerRow, 'col:', headerCol);
//   } else {
//     const allSheets = ssQuote.getSheets();
//     outer:
//     for (let i = 0; i < allSheets.length; i++) {
//       const s = allSheets[i];
//       const r = Math.min(50, s.getLastRow());
//       const c = Math.min(30, s.getLastColumn());
//       if (r < 1 || c < 1) continue;
//       const data = s.getRange(1, 1, r, c).getDisplayValues();
//       for (let rr = 0; rr < r; rr++) {
//         for (let cc = 0; cc < c; cc++) {
//           const cell = String(data[rr][cc] || '');
//           if (!cell) continue;
//           const hit = UQ_DIAMONDS_HEADER_ALIASES.some(k => cell.toLowerCase().indexOf(k.toLowerCase()) >= 0);
//           if (hit) {
//             sh = s; headerRow = rr + 1; headerCol = cc + 1;
//             uq_log_('[QUOTE] Heuristic header hit on sheet', sh.getName(), 'at', 'R'+headerRow+'C'+headerCol, 'cell:', cell);
//             break outer;
//           }
//         }
//       }
//     }
//   }
//   if (!sh || headerRow < 1 || headerCol < 1) {
//     throw new Error('Could not locate Diamonds table header in Quotation. Add named range “' + UQ_ANCHOR_NAME_DIAMONDS + '” to the header.');
//   }

//   // 3) Read existing table block under header (conservative width for v1)
//   const lastRow = sh.getLastRow();
//   const lastCol = sh.getLastColumn();
//   const width   = Math.min(20, (lastCol - headerCol + 1));
//   const startDataRow = headerRow + 1;
//   const height  = Math.max(0, lastRow - headerRow); // includes header
//   uq_log_('[QUOTE] Table window → sheet:', sh.getName(), 'headerRow:', headerRow, 'headerCol:', headerCol, 'height:', height, 'width:', width);

//   const rng  = sh.getRange(headerRow, headerCol, height, width);
//   const grid = rng.getDisplayValues();

//   // Build header map (first row of the grid)
//   const hdr = (grid[0] || []).map(h => String(h || '').trim());
//   uq_log_('[QUOTE] Header row:', JSON.stringify(hdr));

//   const colIdx = (nameCandidates) => {
//     for (const k of nameCandidates) {
//       const needle = k.toLowerCase().replace(/\s+/g,'');
//       for (let j=0; j<hdr.length; j++) {
//         const h = hdr[j].toLowerCase().replace(/\s+/g,'');
//         if (h.indexOf(needle) >= 0) return j;
//       }
//     }
//     return -1;
//   };

//   const idxCert = colIdx(['Certificate No', 'Cert #', 'Certificate', 'Cert']);
//   if (idxCert < 0) throw new Error('Quotation Diamonds header must include a "Certificate No" (or similar) column.');
//   uq_log_('[QUOTE] idxCert:', idxCert);

//   // Build index of existing rows (grid[1..N])
//   const existing = new Map();
//   let firstBlankRowOffset = -1;
//   for (let r = 1; r < grid.length; r++) {
//     const cert = String(grid[r][idxCert] || '').trim();
//     if (!cert) { firstBlankRowOffset = (firstBlankRowOffset < 0 ? r : firstBlankRowOffset); continue; }
//     existing.set(cert.toLowerCase(), r); // row offset within grid block
//   }
//   uq_log_('[QUOTE] Existing rows:', existing.size, 'firstBlankOffset:', firstBlankRowOffset);

//   // Prepare a mutable write buffer for the block
//   const write = sh.getRange(headerRow, headerCol, height, width).getValues();

//   let updated = 0, appended = 0;
//   records.forEach(rec => {
//     const key = rec.certNo.trim().toLowerCase();

//     // Compose sparse output row with header-aware placements
//     const out = new Array(width).fill('');
//     const put = (cols, val) => {
//       const at = colIdx(Array.isArray(cols) ? cols : [cols]);
//       if (at >= 0) out[at] = val;
//     };

//     put(['Certificate No','Cert #','Certificate','Cert'], rec.certNo);
//     put(['LAB','Lab'], rec.lab);
//     put(['Shape'], rec.shape);
//     put(['Carat'], rec.carat);
//     put(['Color'], rec.color);
//     put(['Clarity'], rec.clarity);
//     put(['Vendor'], rec.vendor);
//     put(['Type','Stone Type'], rec.stoneType);
//     put(['Order Status'], rec.orderStatus);
//     put(['Stone Status'], rec.stoneStatus);

//     if (existing.has(key)) {
//       // Update in place
//       const rr = existing.get(key);
//       for (let j=0;j<width;j++) {
//         if (out[j] !== '') write[rr][j] = out[j];
//       }
//       updated++;
//       uq_log_('↻ update', rec.certNo, 'at grid row offset', rr);
//     } else {
//       // Append into first blank row within the buffered block if available
//       let rr = -1;
//       for (let r = 1; r < write.length; r++) {
//         const isBlank = String(write[r][idxCert] || '').trim() === '';
//         if (isBlank) { rr = r; break; }
//       }
//       if (rr >= 0) {
//         for (let j=0;j<width;j++) write[rr][j] = out[j];
//         uq_log_('＋ append-in-block', rec.certNo, 'at grid row offset', rr);
//       } else {
//         // Extend sheet by one row at the end of the current table window
//         const newR = headerRow + (grid.length); // first row after current block
//         sh.insertRowsAfter(headerRow + height - 1, 1);
//         sh.getRange(newR, headerCol, 1, width).setValues([out]);
//         uq_log_('＋ append-by-extension', rec.certNo, 'at sheet row', newR);
//       }
//       appended++;
//     }
//   });

//   // Commit buffered writes
//   sh.getRange(headerRow, headerCol, height, width).setValues(write);
//   uq_log_('[QUOTE] commit: updated =', updated, 'appended =', appended);

//   return {
//     sheetName: sh.getName(),
//     updated: updated,
//     appended: appended,
//     upserted: updated + appended,
//     tableSize: Math.max(0, height - 1), // excluding header
//     anchorFound: !!anchor
//   };
// }

// // --- Legacy → Canon shims (safe no-ops if the name already exists in this file) ---
// if (typeof headerMap_ !== 'function') {
//   function headerMap_(sh){ return headerMap__canon(sh); }
// }
// if (typeof ensureHeaders_ !== 'function') {
//   function ensureHeaders_(sh, labels){ return ensureHeaders__canon(sh, labels); }
// }
// if (typeof getMasterSheet_ !== 'function') {
//   function getMasterSheet_(ss){ return getMasterSheet__canon(ss); }
// }
// if (typeof getOrdersSheet_ !== 'function') {
//   function getOrdersSheet_(wb){ return getOrdersSheet__canon(wb); }
// }
// if (typeof coerceSOTextColumn_ !== 'function') {
//   function coerceSOTextColumn_(sh, H){ return coerceSOTextColumn__canon(sh, H); }
// }
// if (typeof existsSOInMaster_ !== 'function') {
//   function existsSOInMaster_(sh, brand, so, skipRow){ return existsSOInMaster__canon(sh, brand, so, skipRow); }
// }


/** UpdateQuotation_v1.gs — v1.2 (Ratio + Measurement)
 *  Step 1: Menu + Diamonds dialog shell (200_-only read, no writing)
 *  Reuses Diamonds_v1 helpers: dp_getActiveMasterRowContext_(), dp_get200Sheet_(),
 *  dp_headerMapFor200_(), dp_findHeaderIndex_()
 *
 *  v1.2 changes:
 *    - Added Ratio and Measurement columns to bootstrap, submit, and quote writer.
 *
 *  Logging:
 *    - Toggle with Script Property UQ_DEBUG = TRUE|FALSE (helpers below)
 *    - All major steps and decisions are logged with clear tags.
 */

// ──────────────────────────────────────────────────────────────────────────────
// Debug logging helpers
// ──────────────────────────────────────────────────────────────────────────────
function uq_enableDebug(){ PropertiesService.getScriptProperties().setProperty('UQ_DEBUG','TRUE'); }
function uq_disableDebug(){ PropertiesService.getScriptProperties().setProperty('UQ_DEBUG','FALSE'); }

function uq_isDebug_() {
  try {
    const v = (PropertiesService.getScriptProperties().getProperty('UQ_DEBUG') || 'FALSE').toUpperCase();
    return v === 'TRUE' || v === '1' || v === 'YES';
  } catch (e) { return false; }
}
function uq_log_() {
  if (!uq_isDebug_()) return;
  try { Logger.log.apply(Logger, arguments); } catch (e) {}
}
function uq_err_(label, e) {
  const msg = (e && e.stack) ? e.stack : (e && e.message) ? e.message : String(e);
  Logger.log('❌ ' + label + ': ' + msg);
}

// ──────────────────────────────────────────────────────────────────────────────
// 100_ helper: column aliases for Quotation URL
// ──────────────────────────────────────────────────────────────────────────────
var UQ100_ALIASES = {
  'Quotation URL': ['Quotation URL','QuotationURL','Quote URL','QuoteURL']
};

// Named range anchor inside the Quotation file. (top-left header cell of Diamonds table)
var UQ_ANCHOR_NAME_DIAMONDS = 'QUOTE_DIAMONDS_ANCHOR';

// Fallback: if the named range isn't found, look for any header cell that contains one of:
var UQ_DIAMONDS_HEADER_ALIASES = ['Cert', 'Certificate', 'Cert #', 'Certificate No'];

// ── v1.2: Local column aliases for Ratio and Measurement on 200_ ──────────────
// These are used as fallback in case dp_aliases200_ does not define them yet.
var UQ_200_RATIO_ALIASES       = ['Ratio', 'L/W Ratio', 'LW Ratio', 'Length/Width'];
var UQ_200_MEASUREMENT_ALIASES = ['Measurement', 'Measurements', 'Dimensions', 'Dimension'];

// ── v1.2: Column aliases for Ratio and Measurement inside the Quotation sheet ─
var UQ_QUOTE_RATIO_ALIASES       = ['Ratio', 'L/W Ratio', 'LW Ratio', 'Length/Width'];
var UQ_QUOTE_MEASUREMENT_ALIASES = ['Meas.', 'Meas', 'Measurement', 'Measurements', 'Dimensions', 'Dimension'];

// ──────────────────────────────────────────────────────────────────────────────
// Menu
// ──────────────────────────────────────────────────────────────────────────────
function uq_openUpdateQuotationDiamonds() {
  var ui = SpreadsheetApp.getUi();
  try {
    uq_log_('== uq_openUpdateQuotationDiamonds() ==');
    var bootstrap = uq_bootstrapDiamonds_(); // throws helpful errors
    uq_log_('Bootstrap meta →', JSON.stringify(bootstrap && bootstrap.meta || {}, null, 2));
    uq_log_('Bootstrap context →', JSON.stringify(bootstrap && bootstrap.context || {}, null, 2));
    uq_log_('Bootstrap items count:', (bootstrap && bootstrap.items ? bootstrap.items.length : 0));

    var t = HtmlService.createTemplateFromFile('dlg_update_quote_diamonds_v1');
    t.bootstrap = bootstrap;
    var html = t.evaluate().setWidth(1200).setHeight(720);
    ui.showModalDialog(html, '🧾 Update Quotation — Diamonds');
  } catch (e) {
    uq_err_('uq_openUpdateQuotationDiamonds', e);
    ui.alert('🧾 Update Quotation — Diamonds', (e && e.message ? e.message : String(e)) +
      '\n\nOpen 100_ (00_Master Appointments) and select a customer row, then try again.', ui.ButtonSet.OK);
  }
}

// ──────────────────────────────────────────────────────────────────────────────
// Bootstrap: active 100_ context + stones from 200_ by RootApptID (read-only)
// ──────────────────────────────────────────────────────────────────────────────
function uq_bootstrapDiamonds_() {
  uq_log_('== uq_bootstrapDiamonds_() ==');

  // 100_ selection & context (rootApptId, customerName, brand, etc.)
  var ctx = dp_getActiveMasterRowContext_();
  uq_log_('[100_] row context:', JSON.stringify({
    rowIndex: ctx.rowIndex, rootApptId: ctx.rootApptId, customer: ctx.customerName,
    brand: ctx.companyBrand, assignedRep: ctx.assignedRep,
    visitDate: ctx.visitDateStr, visitTime: ctx.visitTimeStr
  }, null, 2));

  // Read Quotation URL from the same 100_ row (best effort)
  var qUrl = '';
  try {
    var colQ = dp_findHeaderIndex_(ctx.headerMap, UQ100_ALIASES['Quotation URL'], false);
    if (colQ > -1) qUrl = String(ctx.sheet.getRange(ctx.rowIndex, colQ).getDisplayValue() || '').trim();
    uq_log_('[100_] Quotation URL:', qUrl || '(blank)');
  } catch (e) {
    uq_err_('Read Quotation URL', e);
  }

  // 200_ target + header map (two-row combiner)
  var target = dp_get200Sheet_();
  var sh200  = target.sheet;
  var hm200  = dp_headerMapFor200_(sh200);
  uq_log_('[200_] workbook:', target.ss.getName(), 'tab:', target.tab, 'rows:', sh200.getLastRow(), 'cols:', sh200.getLastColumn());

  // Resolve columns we'll display
  function cReq(k){ return dp_findHeaderIndex_(hm200, dp_aliases200_[k], true); }

  // v1.2: optional resolver for columns that may not be in dp_aliases200_ yet
  function cOpt(aliases){ return dp_findHeaderIndex_(hm200, aliases, false); }

  var cRoot   = cReq('RootApptID');
  var cOrder  = cReq('Order Status');
  var cStoneS = cReq('Stone Status');
  var cVendor = cReq('Vendor');
  var cType   = cReq('Stone Type');
  var cShape  = cReq('Shape');
  var cCarat  = cReq('Carat');
  var cColor  = cReq('Color');
  var cClarity= cReq('Clarity');
  var cLab    = cReq('LAB');
  var cCert   = cReq('Certificate No');
  var cCust   = cReq('Customer Name');
  var cAppt   = cReq('Customer Appt Time & Date');
  var cRep    = cReq('Assigned Rep');
  var cBrand  = cReq('Company');

  // v1.2: Ratio and Measurement (optional — soft-fail if column absent in 200_)
  var cRatio  = (dp_aliases200_ && dp_aliases200_['Ratio'])
                ? cOpt(dp_aliases200_['Ratio'])
                : cOpt(UQ_200_RATIO_ALIASES);
  var cMeasure= (dp_aliases200_ && dp_aliases200_['Measurement'])
                ? cOpt(dp_aliases200_['Measurement'])
                : cOpt(UQ_200_MEASUREMENT_ALIASES);

  uq_log_('[200_] cRatio:', cRatio, 'cMeasure:', cMeasure);

  var items = [];
  var last = sh200.getLastRow();
  if (last >= 3) {
    var data = sh200.getRange(3, 1, last - 2, sh200.getLastColumn()).getDisplayValues();
    var want = String(ctx.rootApptId || '').trim();
    for (var i = 0; i < data.length; i++) {
      var row = data[i], rIdx = i + 3;
      if (String(row[cRoot - 1] || '').trim() !== want) continue;

      items.push({
        rowIndex: rIdx, rootApptId: want,
        // specs
        vendor:   row[cVendor - 1] || '',
        stoneType:row[cType - 1] || '',
        shape:    row[cShape - 1] || '',
        carat:    row[cCarat - 1] || '',
        color:    row[cColor - 1] || '',
        clarity:  row[cClarity - 1] || '',
        lab:      row[cLab - 1] || '',
        certNo:   row[cCert - 1] || '',
        // v1.2: new fields
        ratio:       (cRatio   > 0 ? (row[cRatio   - 1] || '') : ''),
        measurement: (cMeasure > 0 ? (row[cMeasure - 1] || '') : ''),
        // status
        orderStatus: row[cOrder - 1] || '',
        stoneStatus: row[cStoneS - 1] || '',
        // preview
        customerName: row[cCust - 1] || '',
        apptTimeDate: row[cAppt - 1] || '',
        assignedRep:  row[cRep - 1] || '',
        company:      row[cBrand - 1] || ''
      });
    }
  }
  uq_log_('[200_] Matched stones for rootApptId', ctx.rootApptId, '→', items.length);

  return {
    context: {
      rootApptId: ctx.rootApptId,
      customerName: ctx.customerName,
      visitAt: (ctx.visitDateStr || '') + (ctx.visitTimeStr ? (' ' + ctx.visitTimeStr) : ''),
      companyBrand: ctx.companyBrand || '',
      assignedRep: ctx.assignedRep || '',
      quotationUrl: qUrl || ''
    },
    items: items,
    meta: { spreadsheetUrl: target.ss.getUrl(), tab: target.tab }
  };
}

// ──────────────────────────────────────────────────────────────────────────────
// Public API — dialog submit
// ──────────────────────────────────────────────────────────────────────────────
function uq_submitUpdateQuotationDiamonds(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(28 * 1000);
  try {
    uq_log_('== uq_submitUpdateQuotationDiamonds() ==');
    uq_log_('Payload:', JSON.stringify(payload || {}, null, 2));
    if (!payload || !payload.rootApptId) throw new Error('Missing rootApptId.');
    if (!payload.quotationUrl) throw new Error('Quotation URL not found on 100_.');
    if (!Array.isArray(payload.rowIndexes) || payload.rowIndexes.length === 0) throw new Error('No stones selected.');

    // Source of truth → 200_
    const target = dp_get200Sheet_();
    const sh200  = target.sheet;
    const hm200  = dp_headerMapFor200_(sh200);
    uq_log_('[200_] Using:', target.ss.getName(), 'tab:', target.tab);

    const last   = sh200.getLastRow();
    const lastCol= sh200.getLastColumn();
    uq_log_('[200_] size rows×cols:', last, '×', lastCol);

    // Resolve columns
    const c = (k, req=true) => dp_findHeaderIndex_(hm200, dp_aliases200_[k], req);
    const cOpt2 = (aliases) => dp_findHeaderIndex_(hm200, aliases, false);

    const C = {
      Vendor:      c('Vendor'),
      Type:        c('Stone Type'),
      Shape:       c('Shape'),
      Carat:       c('Carat'),
      Color:       c('Color'),
      Clarity:     c('Clarity'),
      LAB:         c('LAB'),
      Cert:        c('Certificate No'),
      OrderStatus: c('Order Status'),
      StoneStatus: c('Stone Status'),
      Customer:    c('Customer Name'),
      Appt:        c('Customer Appt Time & Date'),
      Rep:         c('Assigned Rep'),
      Company:     c('Company'),
      Root:        c('RootApptID'),
      // v1.2: Ratio and Measurement
      Ratio:       (dp_aliases200_ && dp_aliases200_['Ratio'])
                     ? cOpt2(dp_aliases200_['Ratio'])
                     : cOpt2(UQ_200_RATIO_ALIASES),
      Measurement: (dp_aliases200_ && dp_aliases200_['Measurement'])
                     ? cOpt2(dp_aliases200_['Measurement'])
                     : cOpt2(UQ_200_MEASUREMENT_ALIASES)
    };

    uq_log_('[200_] C.Ratio:', C.Ratio, 'C.Measurement:', C.Measurement);

    // Read selected rows in one batch
    const rowsNeeded = payload.rowIndexes.map(Number).filter(r => r >= 3 && r <= last);
    uq_log_('Selected 200_ rows:', JSON.stringify(rowsNeeded));
    const rects = rowsNeeded.map(r => sh200.getRange(r, 1, 1, lastCol));
    const vals  = rects.map(rng => rng.getDisplayValues()[0]);

    // Normalize → records (key by certNo)
    const recs = vals.map((row, i) => {
      const certNo = String(row[C.Cert - 1] || '').trim();
      if (!certNo) throw new Error('Row '+ rowsNeeded[i] + ': missing Certificate No (cannot upsert Quotation).');
      return {
        certNo:      certNo,
        vendor:      row[C.Vendor - 1] || '',
        stoneType:   row[C.Type - 1] || '',
        shape:       row[C.Shape - 1] || '',
        carat:       row[C.Carat - 1] || '',
        color:       row[C.Color - 1] || '',
        clarity:     row[C.Clarity - 1] || '',
        lab:         row[C.LAB - 1] || '',
        // v1.2: Ratio and Measurement (gracefully blank if column not found)
        ratio:       (C.Ratio       > 0 ? (row[C.Ratio       - 1] || '') : ''),
        measurement: (C.Measurement > 0 ? (row[C.Measurement - 1] || '') : ''),
        orderStatus: String(row[C.OrderStatus - 1] || '').trim(),
        stoneStatus: String(row[C.StoneStatus - 1] || '').trim(),
        customerName: row[C.Customer - 1] || '',
        apptTimeDate: row[C.Appt - 1] || '',
        assignedRep:  row[C.Rep - 1] || '',
        company:      row[C.Company - 1] || ''
      };
    });
    uq_log_('Prepared records from 200_:', JSON.stringify(recs, null, 2));

    const result = uq_writeDiamondsToQuote_(payload.quotationUrl, recs);
    uq_log_('Write result:', JSON.stringify(result, null, 2));

    return {
      ok: true,
      upserted: result.upserted,
      appended: result.appended,
      updated:  result.updated,
      tableSize: result.tableSize,
      sheetName: result.sheetName,
      anchorFound: result.anchorFound,
      message: 'Quotation updated: ' + result.updated + ' updated, ' + result.appended + ' added.'
    };
  } catch (e) {
    uq_err_('uq_submitUpdateQuotationDiamonds', e);
    throw e;
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// ──────────────────────────────────────────────────────────────────────────────
/** Core writer:
 *  - Opens the Quotation by URL (the client's **copy**, not the template)
 *  - Looks for named range QUOTE_DIAMONDS_ANCHOR; if missing, scans headers heuristically
 *  - Reads Diamonds table under the header, indexes by Cert#
 *  - Upserts rows (update existing by Cert#, otherwise append)
 *  - v1.2: also writes Ratio and Measurement columns when present in the Quotation header
 */
// ──────────────────────────────────────────────────────────────────────────────
function uq_writeDiamondsToQuote_(quotationUrl, records) {
  uq_log_('== uq_writeDiamondsToQuote_() ==');
  uq_log_('[QUOTE] URL:', quotationUrl);

  // Open the live **Quotation file** for this client (from 100_ "Quotation URL")
  const ssQuote = SpreadsheetApp.openByUrl(quotationUrl);
  uq_log_('[QUOTE] Opened:', ssQuote.getName(), 'ID:', ssQuote.getId());
  const sheets  = ssQuote.getSheets().map(s => s.getName());
  uq_log_('[QUOTE] Tabs:', JSON.stringify(sheets));

  // 1) Locate anchor (preferred: named range)
  let anchor = null, rc = null;
  try {
    rc = ssQuote.getRangeByName(UQ_ANCHOR_NAME_DIAMONDS);
    if (rc) {
      anchor = [rc.getSheet().getName(), rc.getA1Notation().replace(/^.*!/, '')];
      uq_log_('[QUOTE] Named range found:', UQ_ANCHOR_NAME_DIAMONDS, '→', rc.getA1Notation());
    } else {
      uq_log_('[QUOTE] Named range NOT found:', UQ_ANCHOR_NAME_DIAMONDS);
    }
  } catch (e) {
    uq_err_('getRangeByName', e);
  }

  // 2) If no named range, heuristically find a header cell with "Cert"/"Certificate"
  let sh = null, headerRow = -1, headerCol = -1;
  if (anchor) {
    sh = ssQuote.getSheetByName(anchor[0]);
    headerRow = rc.getRow();
    headerCol = rc.getColumn();
    uq_log_('[QUOTE] Using anchor → sheet:', sh.getName(), 'row:', headerRow, 'col:', headerCol);
  } else {
    const allSheets = ssQuote.getSheets();
    outer:
    for (let i = 0; i < allSheets.length; i++) {
      const s = allSheets[i];
      const r = Math.min(50, s.getLastRow());
      const c = Math.min(30, s.getLastColumn());
      if (r < 1 || c < 1) continue;
      const data = s.getRange(1, 1, r, c).getDisplayValues();
      for (let rr = 0; rr < r; rr++) {
        for (let cc = 0; cc < c; cc++) {
          const cell = String(data[rr][cc] || '');
          if (!cell) continue;
          const hit = UQ_DIAMONDS_HEADER_ALIASES.some(k => cell.toLowerCase().indexOf(k.toLowerCase()) >= 0);
          if (hit) {
            sh = s; headerRow = rr + 1; headerCol = cc + 1;
            uq_log_('[QUOTE] Heuristic header hit on sheet', sh.getName(), 'at', 'R'+headerRow+'C'+headerCol, 'cell:', cell);
            break outer;
          }
        }
      }
    }
  }
  if (!sh || headerRow < 1 || headerCol < 1) {
    throw new Error('Could not locate Diamonds table header in Quotation. Add named range "' + UQ_ANCHOR_NAME_DIAMONDS + '" to the header.');
  }

  // 3) Read existing table block under header
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const width   = Math.min(20, (lastCol - headerCol + 1));
  const startDataRow = headerRow + 1;
  const height  = Math.max(0, lastRow - headerRow);
  uq_log_('[QUOTE] Table window → sheet:', sh.getName(), 'headerRow:', headerRow, 'headerCol:', headerCol, 'height:', height, 'width:', width);

  const rng  = sh.getRange(headerRow, headerCol, height, width);
  const grid = rng.getDisplayValues();

  // Build header map (first row of the grid)
  const hdr = (grid[0] || []).map(h => String(h || '').trim());
  uq_log_('[QUOTE] Header row:', JSON.stringify(hdr));

  const colIdx = (nameCandidates) => {
    for (const k of nameCandidates) {
      const needle = k.toLowerCase().replace(/\s+/g,'');
      for (let j=0; j<hdr.length; j++) {
        const h = hdr[j].toLowerCase().replace(/\s+/g,'');
        if (h.indexOf(needle) >= 0) return j;
      }
    }
    return -1;
  };

  const idxCert = colIdx(['Certificate No', 'Cert #', 'Certificate', 'Cert']);
  if (idxCert < 0) throw new Error('Quotation Diamonds header must include a "Certificate No" (or similar) column.');
  uq_log_('[QUOTE] idxCert:', idxCert);

  // v1.2: resolve Ratio and Measurement column indexes in the Quotation header
  const idxRatio   = colIdx(UQ_QUOTE_RATIO_ALIASES);
  const idxMeasure = colIdx(UQ_QUOTE_MEASUREMENT_ALIASES);
  uq_log_('[QUOTE] idxRatio:', idxRatio, 'idxMeasure:', idxMeasure);

  // Build index of existing rows (grid[1..N])
  const existing = new Map();
  for (let r = 1; r < grid.length; r++) {
    const cert = String(grid[r][idxCert] || '').trim();
    if (!cert) continue;
    existing.set(cert.toLowerCase(), r);
  }
  uq_log_('[QUOTE] Existing rows:', existing.size);

  // Prepare a mutable write buffer
  const write = sh.getRange(headerRow, headerCol, height, width).getValues();

  let updated = 0, appended = 0;
  records.forEach(rec => {
    const key = rec.certNo.trim().toLowerCase();

    // Compose sparse output row
    const out = new Array(width).fill('');
    const put = (cols, val) => {
      const at = colIdx(Array.isArray(cols) ? cols : [cols]);
      if (at >= 0) out[at] = val;
    };

    put(['Certificate No','Cert #','Certificate','Cert'], rec.certNo);
    put(['LAB','Lab'],                                    rec.lab);
    put(['Shape'],                                        rec.shape);
    put(['Carat'],                                        rec.carat);
    put(['Color'],                                        rec.color);
    put(['Clarity'],                                      rec.clarity);
    put(['Vendor'],                                       rec.vendor);
    put(['Type','Stone Type'],                            rec.stoneType);
    put(['Order Status'],                                 rec.orderStatus);
    put(['Stone Status'],                                 rec.stoneStatus);
    // v1.2: Ratio and Measurement
    put(UQ_QUOTE_RATIO_ALIASES,       rec.ratio);
    put(UQ_QUOTE_MEASUREMENT_ALIASES, rec.measurement);

    if (existing.has(key)) {
      // Update in place
      const rr = existing.get(key);
      for (let j=0; j<width; j++) {
        if (out[j] !== '') write[rr][j] = out[j];
      }
      updated++;
      uq_log_('↻ update', rec.certNo, 'at grid row offset', rr);
    } else {
      // Append into first blank row within the buffered block if available
      let rr = -1;
      for (let r = 1; r < write.length; r++) {
        const isBlank = String(write[r][idxCert] || '').trim() === '';
        if (isBlank) { rr = r; break; }
      }
      if (rr >= 0) {
        for (let j=0; j<width; j++) write[rr][j] = out[j];
        uq_log_('＋ append-in-block', rec.certNo, 'at grid row offset', rr);
      } else {
        // Extend sheet by one row at the end of the current table window
        const newR = headerRow + (grid.length);
        sh.insertRowsAfter(headerRow + height - 1, 1);
        sh.getRange(newR, headerCol, 1, width).setValues([out]);
        uq_log_('＋ append-by-extension', rec.certNo, 'at sheet row', newR);
      }
      appended++;
    }
  });

  // Commit buffered writes
  sh.getRange(headerRow, headerCol, height, width).setValues(write);
  uq_log_('[QUOTE] commit: updated =', updated, 'appended =', appended);

  return {
    sheetName: sh.getName(),
    updated: updated,
    appended: appended,
    upserted: updated + appended,
    tableSize: Math.max(0, height - 1),
    anchorFound: !!anchor
  };
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
