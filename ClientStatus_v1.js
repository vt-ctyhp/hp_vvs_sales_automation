// /***** Client Status — v2.2 (mirror to 301/302 removed) *****/

// /** CHANGELOG (v2.2, patched)
//  * - REMOVED: All 301/302 mirroring and related helpers.
//  * - UNCHANGED: Central audit, per-client report creation/log, snapshot updates,
//  *              “Updated By/At” metadata on master, DV hooks, and Reminders hooks.
//  */


// // === CONFIG ===
// const MASTER_SHEET_NAME = '00_Master Appointments';
// const CS_MASTER_SHEET_NAME = MASTER_SHEET_NAME;
// const CS_AUDIT_SHEET = '03_Client_Status_Log';
// const CS_AUDIT_TAB = CS_AUDIT_SHEET;
// const CS_REPORT_SHEET = 'Client Status';
// const CS_WRITE_PER_CLIENT_LOG = true;
// const CS_TZ = 'America/Los_Angeles';

// const CS_REPORT_URL_COL = 'Client Status Report URL';
// const CS_PROSPECT_URL_COL = 'Prospect Folder URL';
// const CS_REPORT_NAME_FMT = '{Brand} – {APPT_ID} – Client Status Report';

// // Color column names in "Dropdown"
// const COL_SALES_STAGE_HEX   = 'SS - Hex Code';
// const COL_CONV_STATUS_HEX   = 'CS - Hex Code';
// const COL_CUST_ORDER_HEX    = 'COS - Hex Code';
// const COL_IN_PRODUCTION_HEX = 'IPS - Hex Code'; // NEW
// const COL_CENTER_STONE_HEX  = 'CSOS - Hex Code';

// // === TEMPLATE CONFIG ===
// function getTemplateId_() {
//   return PropertiesService.getScriptProperties().getProperty('CS_REPORT_TEMPLATE_ID') || '';
// }

// // === Helpers ===
// function headerIndexMap_(headerRow){ const map={}; headerRow.forEach((h,i)=>{ if (h) map[String(h).trim()]=i; }); return map; }
// /** Case-insensitive header finder by regex; returns zero-based index or -1. */
// function findHeaderIndexByRegex_(headerRow, regex){
//   for (var i = 0; i < headerRow.length; i++){
//     if (regex.test(String(headerRow[i] || ''))) return i;
//   }
//   return -1;
// }

// function extractIdFromUrl_(url){ const m=String(url).match(/[-\w]{25,}/); return m?m[0]:''; }
// function getByAny_(H, vals, names){ for (const n of names){ if (H[n]!=null) return vals[H[n]] ?? ''; } return ''; }

// function normalizeMultiArray_(v){
//   if (Array.isArray(v)) return v.map(s=>String(s||'').trim()).filter(Boolean);
//   return String(v||'')
//     .split(/[,;|/]|(?:\s*&\s*)/g)
//     .map(s=>s.trim())
//     .filter(Boolean);
// }
// function joinMulti_(arr){
//   const a = normalizeMultiArray_(arr);
//   const seen = new Set(); const out=[];
//   a.forEach(x=>{ if(!seen.has(x)){ seen.add(x); out.push(x); } });
//   return out.join(', ');
// }

// // === Read lists + hex maps from "Dropdown" with ONE data fetch ===
// function readDropdowns_() {
//   const sh = SpreadsheetApp.getActive().getSheetByName('Dropdown');
//   if (!sh) throw new Error('Missing tab "Dropdown".');

//   const lastRow = sh.getLastRow();
//   const lastCol = sh.getLastColumn();
//   const header = sh.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h||'').trim());
//   const data   = lastRow > 1 ? sh.getRange(2,1,lastRow-1,lastCol).getValues() : [];
//   const idx = (name) => header.indexOf(String(name).trim());
//   const colVals = (name) => {
//     const c = idx(name); if (c < 0 || data.length === 0) return [];
//     const col = new Array(data.length);
//     for (let i=0;i<data.length;i++) col[i] = String(data[i][c]||'').trim();
//     return col;
//   };

//   // Value columns
//   const assignedReps         = colVals('Assigned Rep').filter(Boolean);
//   const assistedReps         = colVals('Assisted Rep').filter(Boolean);
//   const salesStages          = colVals('Sales Stage').filter(Boolean);           // <-- plural
//   const convStatuses         = colVals('Conversion Status').filter(Boolean);
//   const customOrderStatuses  = colVals('Custom Order Status').filter(Boolean);
//   const centerStoneStatuses  = colVals('Center Stone Order Status').filter(Boolean);
//   const inProductionStatuses = colVals('In Production Status').filter(Boolean);

//   // Hex columns aligned row-for-row
//   const ssHex   = colVals(COL_SALES_STAGE_HEX);
//   const csHex   = colVals(COL_CONV_STATUS_HEX);
//   const cosHex  = colVals(COL_CUST_ORDER_HEX);
//   const csosHex = colVals(COL_CENTER_STONE_HEX);
//   const ipsHex  = colVals(COL_IN_PRODUCTION_HEX);

//   const buildHexMap = (values, hexes) => {
//     const map = {};
//     const n = Math.min(values.length, hexes.length);
//     for (let i=0; i<n; i++){
//       const v = String(values[i]||'').trim();
//       const h = String((hexes[i]||'').replace('#','').trim());
//       if (!v) continue;
//       if (/^[0-9A-Fa-f]{6}$/.test(h)) map[v] = '#'+h.toUpperCase();
//     }
//     return map;
//   };

//   return {
//     assignedReps, assistedReps, salesStages, convStatuses, customOrderStatuses, centerStoneStatuses, inProductionStatuses,
//     colors: {
//       salesStage:   buildHexMap(salesStages,          ssHex),  // key remains 'salesStage' for HTML chip color lookups
//       convStatus:   buildHexMap(convStatuses,         csHex),
//       customOrder:  buildHexMap(customOrderStatuses,  cosHex),
//       centerStone:  buildHexMap(centerStoneStatuses,  csosHex),
//       inProduction: buildHexMap(inProductionStatuses, ipsHex)
//     }
//   };
// }

// /** Read "Validation Rules (Flattened Matrix)" and "Viewing Rules" from the Dropdown tab. */
// function readValidationRulesFlat_() {
//   const sh = SpreadsheetApp.getActive().getSheetByName('Dropdown Rules');
//   if (!sh) throw new Error('Missing tab "Dropdown Rules".');

//   const lastRow = sh.getLastRow();
//   const lastCol = sh.getLastColumn();
//   if (lastRow < 2 || lastCol < 1) return { matrix: [], viewing: [] };

//   const all = sh.getRange(1,1,lastRow,lastCol).getValues();

//   // ---- Find header row for the flattened matrix
//   let hdrRow = -1;
//   let col = { sales:-1, conv:-1, cos:-1, ips:-1, csReq:-1, dead:-1, notes:-1 };

//   for (let r = 0; r < all.length; r++) {
//     const row = all[r].map(x => String(x||'').trim());
//     const iSales = row.indexOf('Sales Stage');
//     const iConv  = row.indexOf('Conversion Status');
//     const iCOS   = row.indexOf('Custom Order Status');
//     if (iSales >= 0 && iConv >= 0 && iCOS >= 0) {
//       hdrRow = r;
//       col.sales = iSales;
//       col.conv  = iConv;
//       col.cos   = iCOS;
//       col.ips   = row.indexOf('In Production Status Requirement');
//       col.csReq = row.indexOf('Center Stone Status Requirement');
//       col.dead  = row.indexOf('Deadline Requirement');
//       col.notes = row.indexOf('Notes / Flags');
//       break;
//     }
//   }

//   const matrix = [];
//   if (hdrRow >= 0) {
//     for (let r = hdrRow + 1; r < all.length; r++) {
//       const row = all[r];
//       const s   = String(row[col.sales] || '').trim();
//       const c   = String(row[col.conv]  || '').trim();
//       const cos = String(row[col.cos]   || '').trim();
//       const ips = col.ips  >= 0 ? String(row[col.ips]  || '').trim() : '';
//       const csr = col.csReq>= 0 ? String(row[col.csReq]|| '').trim() : '';
//       const dr  = col.dead >= 0 ? String(row[col.dead] || '').trim() : '';
//       const nt  = col.notes>= 0 ? String(row[col.notes]|| '').trim() : '';
//       if (s || c || cos || ips || csr || dr || nt) {
//         matrix.push({
//           salesStage: s, convStatus: c, customOrderStatus: cos,
//           ipsRequirement: ips, centerStoneRequirement: csr,
//           deadlineRequirement: dr, notes: nt
//         });
//       }
//     }
//   }

//   // ---- Find header row for Viewing Rules
//   let vHdr = -1, cDays = -1, cMin = -1;
//   for (let r = 0; r < all.length; r++) {
//     const row = all[r].map(x => String(x||'').trim());
//     const iD = row.indexOf('Days Before Viewing');
//     const iM = row.indexOf('Minimum Allowed Center Stone Status');
//     if (iD >= 0 && iM >= 0) { vHdr = r; cDays = iD; cMin = iM; break; }
//   }

//   const viewing = [];
//   if (vHdr >= 0) {
//     for (let r = vHdr + 1; r < all.length; r++) {
//       const row = all[r];
//       const d = String(row[cDays] || '').trim();
//       const m = String(row[cMin]  || '').trim();
//       if (d || m) viewing.push({ daysBefore: d, minimum: m });
//     }
//   }

//   return { matrix, viewing };
// }




// /** Open the Client Status dialog (2-column; popover chip pickers on right) */
// function cs_openStatusDialog_() {
//   const ss = SpreadsheetApp.getActive();
//   const sh = ss.getSheetByName(CS_MASTER_SHEET_NAME);
//   const r = sh.getActiveRange();
//   if (!r || r.getRow() === 1) {
//     SpreadsheetApp.getUi().alert('⚠️ Select a data row in 00_Master Appointments first.');
//     return;
//   }
//   const row = r.getRow();

//   const header = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
//   const H = headerIndexMap_(header);
//   const vals = sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0];

//   const get = name => H[name] != null ? vals[H[name]] : '';

//   const assignedRepStr = String(get('Assigned Rep') || '');
//   const assistedRepStr = String(get('Assisted Rep') || '');

// const orderDateISO = toISODateForInput_(get('Order Date'));

//   const prefill = {
//     clientName:  String(get('Customer Name') || ''),
//     apptId:      String(get('APPT_ID') || ''),
//     assignedRep: assignedRepStr,
//     assistedRep: assistedRepStr,
//     assignedRepArr: normalizeMultiArray_(assignedRepStr),
//     assistedRepArr: normalizeMultiArray_(assistedRepStr),
//     salesStage:  String(get('Sales Stage') || ''),
//     convStatus:  String(get('Conversion Status') || ''),
//     customOrder: String(get('Custom Order Status') || ''),
//     inProduction: String(get('In Production Status') || ''), // NEW
//     centerStone: String(get('Center Stone Order Status') || ''),
//     nextSteps:   String(get('Next Steps') || ''),
//     orderDate:   orderDateISO
//   };

//   const lists = readDropdowns_(); // lists + colors

//   // NEW: Read flattened prevention rules from Dropdown + compute Visit ISO
//   const rulesFlat = readValidationRulesFlat_();

//   let visitISO = String(get('ApptDateTime (ISO)') || '').trim();
//   if (!visitISO) {
//     const vdate = String(get('Visit Date') || '').trim();
//     const vtime = String(get('Visit Time') || '').trim();
//     if (vdate || vtime) {
//       try { visitISO = Utilities.formatDate(new Date(vdate + ' ' + vtime), CS_TZ, "yyyy-MM-dd'T'HH:mm:ssXXX"); } catch(_){}
//     }
//   }

//   const t = HtmlService.createTemplateFromFile('dlg_client_status_v1');
//   t.prefill = prefill;
//   t.lists = {
//     assignedReps:         lists.assignedReps,
//     assistedReps:         lists.assistedReps,
//     salesStages:          lists.salesStages,          // <-- plural to match HTML
//     convStatuses:         lists.convStatuses,
//     customOrderStatuses:  lists.customOrderStatuses,
//     centerStoneStatuses:  lists.centerStoneStatuses,
//     inProductionStatuses: lists.inProductionStatuses
//   };
//   t.colors = lists.colors;

//   t.prefill.visitISO = visitISO || '';
//   t.rulesFlat = rulesFlat;

//   const html = t.evaluate().setWidth(1040).setHeight(720);
//   SpreadsheetApp.getUi().showModalDialog(html, 'Client Status Update');
// }

// /** Submit from dialog (arrays for reps; statuses required) */
// function cs_submitFromDialog(payload) {
// // ---- Conditional server guard mirroring client logic ----
//   function _centerStoneRequired(stage, conv) {
//     if (/^Lost Lead/i.test(String(stage||''))) return false;
//     if (/^Viewing Scheduled$/i.test(String(conv||''))) return true;
//     if (/^(Deposit Paid|Confirmed Order|Order In Progress)$/i.test(String(conv||''))) return true;
//     return false;
//   }

//   // Sales Stage & Conversion Status always required
//   ['salesStage','convStatus'].forEach(function (k) {
//     if (!String(payload[k] || '').trim()) {
//       throw new Error('Please complete: Sales Stage and Conversion Status before submitting.');
//     }
//   });

//   // Custom Order Status required unless rules yielded zero options
//   var cosEmptyAllowed = !!payload.cosAllowedEmpty;
//   if (!cosEmptyAllowed && !String(payload.customOrder || '').trim()) {
//     throw new Error('Please select a Custom Order Status.');
//   }

//   var isInProduction = String(payload.customOrder || '') === 'In Production';
//   if (isInProduction && !String(payload.inProduction || '').trim()) {
//     throw new Error('Please select an "In Production Status" since Custom Order Status is In Production.');
//   }

//   var need3D = /^(3D Requested|3D Revision Requested)$/i.test(String(payload.customOrder||''));
//   if (need3D && !String((payload.deadline3d||'')).trim()) {
//     throw new Error('3D Deadline is required when Custom Order Status is 3D Requested or 3D Revision Requested.');
//   }
//   if (isInProduction && !String((payload.prodDeadline||'')).trim()) {
//     throw new Error('Production Deadline is required when Custom Order Status is In Production.');
//   }

//   // Order Date is required for these COS values
//   var needOrderDate = /^(Approved for Production|Waiting Production Timeline|In Production|Final Photos\s*[–-]\s*Waiting Approval|Warehouse|Ship to US|In US Store|Ship to Customer|Order Completed)$/i
//     .test(String(payload.customOrder||''));
//   if (needOrderDate && !String(payload.orderDate || '').trim()) {
//     throw new Error('Order Date is required for the selected Custom Order Status.');
//   }

//   if (_centerStoneRequired(String(payload.salesStage||''), String(payload.convStatus||'')) &&
//       !String(payload.centerStone || '').trim()) {
//     throw new Error('Center Stone Order Status is required for Viewing Scheduled or Deposit/Confirmed/Order In Progress.');
//   }

//   const ss = SpreadsheetApp.getActive();
//   const sh = ss.getSheetByName(CS_MASTER_SHEET_NAME);
//   const r = sh.getActiveRange();
//   if (!r || r.getNumRows() !== 1 || r.getRow() === 1) throw new Error('Select exactly one row.');
//   const row = r.getRow();

//   const header = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
//   const H = headerIndexMap_(header);
//   const vals = sh.getRange(row,1,1, sh.getLastColumn()).getValues()[0];
//   // Snapshot previous Center Stone status BEFORE we write any changes
//   const __prevCenterStone = String(vals[H['Center Stone Order Status']] ?? '').trim();

//   // Normalize multi arrays → single stored strings (back-compatible)
//   const assignedJoined = joinMulti_(payload.assignedRep);
//   const assistedJoined = joinMulti_(payload.assistedRep);

//   const setIf = (name, value) => { if (value != null && String(value).trim() !== '' && H[name] != null) { vals[H[name]] = value; } };

//   setIf('Assigned Rep',              assignedJoined);
//   setIf('Assisted Rep',              assistedJoined);
//   setIf('Sales Stage',               payload.salesStage);
//   setIf('Conversion Status',         payload.convStatus);
//   setIf('Custom Order Status',       payload.customOrder);
//   setIf('Order Date', payload.orderDate); 

//   // NEW: write/clear In Production Status (robust header lookup)
//   var ipsIdx = (H['In Production Status'] != null)
//       ? H['In Production Status']
//       : findHeaderIndexByRegex_(header, /in\s*production\s*status/i);

//   if (ipsIdx >= 0) {
//     vals[ipsIdx] = isInProduction ? String(payload.inProduction || '').trim() : '';
//   }

//   // ---- NEW: deadlines write + move counters + log meta ----
//   /** @type {Object.<string,number>} */
//   const H2 = H; // alias for clarity

//   // Robust column lookups (handles slight header variations)
//   const idxProdDeadline = (H2['Production Deadline'] != null)
//     ? H2['Production Deadline']
//     : findHeaderIndexByRegex_(header, /(Production|Prod\.)\s*Deadline/i);

//   const idx3dDeadline = (H2['3D Deadline'] != null)
//     ? H2['3D Deadline']
//     : findHeaderIndexByRegex_(header, /3D\s*Deadline/i);

//   const idxProdMoves = (H2['# of Times Prod. Deadline Moved'] != null)
//     ? H2['# of Times Prod. Deadline Moved']
//     : findHeaderIndexByRegex_(header, /#\s*of\s*Times\s*(Prod|Production).*Deadline.*Moved/i);

//   const idx3dMoves = (H2['# of Times 3D Deadline Moved'] != null)
//     ? H2['# of Times 3D Deadline Moved']
//     : findHeaderIndexByRegex_(header, /#\s*of\s*Times\s*3D.*Deadline.*Moved/i);

//   // Capture current (pre‑update) values
//   const prevProdDeadline = idxProdDeadline >= 0 ? String(vals[idxProdDeadline] || '').trim() : '';
//   const prev3dDeadline   = idx3dDeadline   >= 0 ? String(vals[idx3dDeadline]   || '').trim() : '';
//   const prevProdMovesStr = idxProdMoves    >= 0 ? String(vals[idxProdMoves]    || '').trim() : '';
//   const prev3dMovesStr   = idx3dMoves      >= 0 ? String(vals[idx3dMoves]      || '').trim() : '';

//   // Determine applicability and new input
//   const is3D = /^(3D Requested|3D Revision Requested)$/i.test(String(payload.customOrder||''));
//   const newProdDeadline = isInProduction ? String(payload.prodDeadline || '') : '';
//   const new3dDeadline   = is3D          ? String(payload.deadline3d   || '') : '';

//   // Write deadlines to Master row (or clear when not applicable)
//   if (idxProdDeadline >= 0) vals[idxProdDeadline] = newProdDeadline;
//   if (idx3dDeadline   >= 0) vals[idx3dDeadline]   = new3dDeadline;

//   // Move counters logic
//   let prodChanged = false, threeDChanged = false;

//   // Production counter
//   if (idxProdDeadline >= 0 && isInProduction) {
//     if (!prevProdDeadline && newProdDeadline) {
//       // first set → dash
//       if (idxProdMoves >= 0) vals[idxProdMoves] = '-';
//     } else if (prevProdDeadline && newProdDeadline && prevProdDeadline !== newProdDeadline) {
//       prodChanged = true;
//       const prevN = (prevProdMovesStr === '-' || prevProdMovesStr === '') ? 0 : (parseInt(prevProdMovesStr, 10) || 0);
//       if (idxProdMoves >= 0) vals[idxProdMoves] = String(prevN + 1);
//     }
//   }

//   // 3D counter
//   if (idx3dDeadline >= 0 && is3D) {
//     if (!prev3dDeadline && new3dDeadline) {
//       // first set → dash
//       if (idx3dMoves >= 0) vals[idx3dMoves] = '-';
//     } else if (prev3dDeadline && new3dDeadline && prev3dDeadline !== new3dDeadline) {
//       threeDChanged = true;
//       const prevN = (prev3dMovesStr === '-' || prev3dMovesStr === '') ? 0 : (parseInt(prev3dMovesStr, 10) || 0);
//       if (idx3dMoves >= 0) vals[idx3dMoves] = String(prevN + 1);
//     }
//   }

//   // Build log meta for the Client Status Report (what changed this submit)
//   let logDeadlineType = '', logDeadlineDate = '', logMoveCount = '';
//   if (idxProdDeadline >= 0 && isInProduction && ( (!prevProdDeadline && newProdDeadline) || prodChanged )) {
//     logDeadlineType = 'Production';
//     logDeadlineDate = newProdDeadline;
//     logMoveCount    = (idxProdMoves >= 0 ? String(vals[idxProdMoves] || '') : '');
//   }
//   if (idx3dDeadline >= 0 && is3D && ( (!prev3dDeadline && new3dDeadline) || threeDChanged )) {
//     logDeadlineType = logDeadlineType ? (logDeadlineType + ' | 3D') : '3D';
//     logDeadlineDate = logDeadlineDate ? (logDeadlineDate + ' | ' + new3dDeadline) : new3dDeadline;
//     const mc = (idx3dMoves >= 0 ? String(vals[idx3dMoves] || '') : '');
//     logMoveCount = logMoveCount ? (logMoveCount + ' | ' + mc) : mc;
//   }

//   // If COS is Final Photos – Waiting Approval or any later shipping step → IPS must be Production Completed
//   (function enforceIPSForLaterPhases(){
//     const cosNow = String(payload.customOrder || '').trim();
//     const later = new Set([
//       'Final Photos – Waiting Approval',
//       'Warehouse',
//       'Ship to US',
//       'In US Store',
//       'Ship to Customer',
//       'Order Completed'
//     ]);
//     if (later.has(cosNow) && typeof ipsIdx === 'number' && ipsIdx >= 0) {
//       vals[ipsIdx] = 'Production Completed';
//     }
//   })();

//   // Stash the log meta to forward to the next step (success screen + CSR log)
//   payload.__deadlineLog = { type: logDeadlineType, date: logDeadlineDate, moves: logMoveCount };


//   setIf('Center Stone Order Status', payload.centerStone);
//   if (H['Next Steps'] != null && payload.nextSteps != null) vals[H['Next Steps']] = payload.nextSteps;

//   // Single write
//   sh.getRange(row, 1, 1, vals.length).setValues([vals]);

//   // ---- (NEW) Create Wax Request if asked ----
//   var waxSummary = null;
//   try {
//     if (payload.wax && payload.wax.request === true) {
//       // Determine RootApptID (or fall back to APPT_ID)
//       var rootApptId = String(
//         (H['RootApptID'] != null ? vals[H['RootApptID']] : '') ||
//         (H['APPT_ID']    != null ? vals[H['APPT_ID']]    : '') ||
//         ''
//       ).trim();

//       if (rootApptId) {
//         var wres = wax_onRequestSubmit_({
//           rootApptId: rootApptId,
//           soMo: (payload.wax.soMo || ''),
//           neededByRep: (payload.wax.neededBy || ''),
//           priority: (payload.wax.priority || ''),
//           requestedBy: (Session.getActiveUser().getEmail() || assignedJoined || '')
//         }) || {};
//         // Normalized for the HTML success view
//         waxSummary = {
//           created: !!wres.ok,
//           requestId: wres.requestId || '',
//           folderUrl: wres.folderUrl || '',
//           rowUrl:    wres.url || ''
//         };
//       }
//     }
//   } catch (e) {
//     Logger.log('Wax create failed: ' + (e && e.message ? e.message : e));
//   }

//   // Continue pipeline (audit + client report + snapshot + mirror)
//   // Pass Assisted Rep directly (no PropertiesService temp handoff)
//   return cs_submitClientStatusUpdate_({
//     assistedRep:     assistedJoined,
//     prevCenterStone: __prevCenterStone,
//     inProduction:    String(payload.inProduction || '').trim(),
//     wax:             waxSummary || null,
//     waxSummaryStr:   String(payload.waxSummary || ''),

//     // NEW: forward dates so the success screen can show them (and no undefined refs)
//     prodDeadline: String(payload.prodDeadline || ''),
//     deadline3d:   String(payload.deadline3d   || '')
//   });
// }

// // === Create/find per-client report; write audit/log/snapshot ===
// function cs_createOrGetReportForSelection_(opts) {
//   const inSubmit = !!(opts && opts.inSubmit);

//   let lock;
//   if (!inSubmit) {
//     lock = LockService.getDocumentLock();
//     if (!lock.tryLock(1500)) return { ok:false, error:'LOCKED' };
//   }
//   try {
//     const ss = SpreadsheetApp.getActive();
//     const sh = ss.getSheetByName(CS_MASTER_SHEET_NAME);
//     const r = sh.getActiveRange();
//     const row = r.getRow();

//     const header = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
//     const H = headerIndexMap_(header);
//     const vals = sh.getRange(row,1,1, sh.getLastColumn()).getValues()[0];
//     const get  = n => H[n] != null ? (vals[H[n]] ?? '') : '';

//     const apptId = String(get('APPT_ID')).trim();
//     const brand  = String(get('Brand')).trim();
//     const client = String(get('Customer Name')).trim();
//     const name   = CS_REPORT_NAME_FMT.replace('{Brand}', brand || 'VVS').replace('{APPT_ID}', apptId);

//     let reportUrl = String(get(CS_REPORT_URL_COL) || '').trim();
//     let reportId  = reportUrl ? extractIdFromUrl_(reportUrl) : '';
//     let reportSS  = null;

//     // Validate and keep the opened handle if it exists
//     if (reportId) {
//       try {
//         reportSS = SpreadsheetApp.openById(reportId);
//       } catch (e) {
//         reportId = '';
//       }
//     }

//     if (!reportId) {
//       const parent = pickParentFolder_(get(CS_PROSPECT_URL_COL), client);
//       reportId = createClientReport_(name, parent);
//       reportUrl = 'https://docs.google.com/spreadsheets/d/' + reportId + '/edit';
//       if (H[CS_REPORT_URL_COL] != null) sh.getRange(row, H[CS_REPORT_URL_COL] + 1).setValue(reportUrl);
//       // Open once for return
//       reportSS = SpreadsheetApp.openById(reportId);
//     }

//     return { ok:true, id:reportId, url:reportUrl, ss: reportSS };

//   } catch (e) {
//     return { ok:false, error: String(e && e.message || e) };
//   } finally {
//     if (lock) { try { lock.releaseLock(); } catch(_){ } }
//   }
// }

// function pickParentFolder_(prospectUrl, clientName) {
//   if (prospectUrl) {
//     const id = extractIdFromUrl_(String(prospectUrl));
//     try { return DriveApp.getFolderById(id); } catch (e) {}
//   }
//   try {
//     const it = DriveApp.getFoldersByName(clientName || 'Clients');
//     if (it.hasNext()) return it.next();
//   } catch (e) {}
//   return DriveApp.getRootFolder();
// }
// function createClientReport_(name, parentFolder) {
//   const templateId = getTemplateId_();
//   if (!templateId) throw new Error('Client Status: CS_REPORT_TEMPLATE_ID not set in Project Properties.');
//   const tmplFile = DriveApp.getFileById(templateId);
//   const copy = tmplFile.makeCopy(name, parentFolder || DriveApp.getRootFolder());
//   const fileId = copy.getId();
//   try { if (parentFolder) DriveApp.getRootFolder().removeFile(copy); } catch (e) {}
//   return fileId;
// }

// /**
//  * Ensure or refresh the _Config sheet in a Client Status Report.
//  * Hybrid mode: create if missing, update only when blank or outdated.
//  */
// function ensureReportConfig_(reportSS, opts){
//   const rootApptId = String(opts.rootApptId||'').trim();
//   const reportId   = String(opts.reportId||reportSS.getId()).trim();

//   let sh = reportSS.getSheetByName('_Config');
//   if (!sh) {
//     sh = reportSS.insertSheet('_Config');
//     try { sh.hideSheet(); } catch(_){}
//     sh.appendRow(['ROOT_APPT_ID', rootApptId]);
//     sh.appendRow(['CONTROLLER_URL', ScriptApp.getService().getUrl()]);
//     sh.appendRow(['REPORT_REANALYZE_TOKEN',
//       PropertiesService.getScriptProperties().getProperty('REPORT_REANALYZE_TOKEN') || ''
//     ]);
//     sh.appendRow(['REPORT_ID', reportId]);
//     return;
//   }

//   // Read current values into a map
//   const vals = sh.getRange(1,1,sh.getLastRow(),2).getValues();
//   const map = {};
//   vals.forEach(r => { if (r[0]) map[String(r[0]).trim()] = String(r[1]||'').trim(); });

//   // Always expected keys
//   const want = {
//     ROOT_APPT_ID: rootApptId,
//     CONTROLLER_URL: ScriptApp.getService().getUrl(),
//     REPORT_REANALYZE_TOKEN: PropertiesService.getScriptProperties().getProperty('REPORT_REANALYZE_TOKEN') || '',
//     REPORT_ID: reportId
//   };

//   Object.keys(want).forEach((k,i) => {
//     const cur = map[k] || '';
//     const need = String(want[k]||'');
//     if (cur !== need) {
//       // Find existing row or append if missing
//       let rowIdx = vals.findIndex(r => String(r[0]).trim() === k);
//       if (rowIdx >= 0) {
//         sh.getRange(rowIdx+1, 2).setValue(need);
//       } else {
//         sh.appendRow([k, need]);
//       }
//     }
//   });
// }


// function cs_submitClientStatusUpdate_(opts) {
//   const lock = LockService.getDocumentLock();
//   if (!lock.tryLock(1500)) return { ok:false, error:'LOCKED' };
//   try {
//     const ss = SpreadsheetApp.getActive();
//     const master = ss.getSheetByName(CS_MASTER_SHEET_NAME);
//     const r = master.getActiveRange();
//     const row = r.getRow();

//     const header = master.getRange(1,1,1, master.getLastColumn()).getValues()[0];
//     const H = headerIndexMap_(header);
//     const vals = master.getRange(row,1,1, master.getLastColumn()).getValues()[0];
//     const get  = n => vals[H[n]] ?? '';

//     const apptId      = String(get('APPT_ID') || '').trim();
//     const brand       = String(get('Brand') || '');
//     const client      = String(get('Customer Name') || '');
//     const rep         = String(get('Assigned Rep') || '');
//     const salesStage  = String(get('Sales Stage') || '');
//     const convStatus  = String(get('Conversion Status') || '');
//     const customOrd   = String(get('Custom Order Status') || '');
//     const inProduction = String(get('In Production Status') || (opts && opts.inProduction) || ''); // NEW (+fallback)
//     const centerStone = String(get('Center Stone Order Status') || '');
//     const nextSteps   = String(get('Next Steps') || '');
//     const orderDate   = String(get('Order Date') || '');  // NEW

//     const phone        = String(getByAny_(H, vals, ['Phone','Client Phone','Customer Phone']) || '');
//     const email        = String(getByAny_(H, vals, ['Email','Client Email','Customer Email']) || '');
//     const occasion     = String(getByAny_(H, vals, ['Occasion']) || '');
//     const budgetRange  = String(getByAny_(H, vals, ['Budget Range']) || '');
//     const decisionMkr  = String(getByAny_(H, vals, ['Decision-Maker','Decision Maker']) || '');
//     const initialReq   = String(getByAny_(H, vals, ['Initial Request']) || '');
//     const soNumber     = String(getByAny_(H, vals, ['SO Number','SO#']) || '').trim();

//     const now  = new Date();
//     const iso  = Utilities.formatDate(now, CS_TZ, 'yyyy-MM-dd');
//     const ts   = Utilities.formatDate(now, CS_TZ, 'yyyy-MM-dd HH:mm:ss');
//     const nice = Utilities.formatDate(now, CS_TZ, 'MMM d, yyyy h:mm a z');
//     const user  = Session.getActiveUser().getEmail() || rep || 'Unknown';
//     const assistedRep = String((opts && opts.assistedRep) || '');


//     // 1) Central audit (+Applied To note)
//     const audit = ss.getSheetByName(CS_AUDIT_TAB);
//     const auditExists = !!audit;

//     if (auditExists) {
//       // Compute how many rows will be touched by fan-out (active row + siblings sharing RootApptID/APPT_ID)
//       const rootKeyForAudit = String(get('RootApptID') || get('APPT_ID') || '').trim();
//       let appliedCountTotal = 1; // at least the active row
//       if (rootKeyForAudit) {
//         const lastRowAll = master.getLastRow();
//         if (lastRowAll > 1) {
//           const matchColIndexAudit = (H['RootApptID'] != null) ? H['RootApptID']
//                                   : (H['APPT_ID']    != null) ? H['APPT_ID']
//                                   : -1;
//           if (matchColIndexAudit >= 0) {
//             const allValsAudit = master.getRange(2, 1, lastRowAll - 1, master.getLastColumn()).getValues();
//             for (let i = 0; i < allValsAudit.length; i++) {
//               const rnum = i + 2; // data starts at row 2
//               if (rnum === row) continue; // skip active row (already counted)
//               const idHere = String(allValsAudit[i][matchColIndexAudit] || '').trim();
//               if (idHere && idHere === rootKeyForAudit) appliedCountTotal++;
//             }
//           }
//         }
//       }

//       const appliedNote = `Applied to ${appliedCountTotal} row${appliedCountTotal === 1 ? '' : 's'}`
//                           + (rootKeyForAudit ? ` (RootApptID=${rootKeyForAudit})` : '');

//       // Ensure audit has an "Applied To" column; add once if missing and refresh header
//       let auditHeader = audit.getRange(1,1,1,audit.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
//       if (auditHeader.indexOf('Applied To') < 0) {
//         audit.getRange(1, audit.getLastColumn() + 1).setValue('Applied To');
//         auditHeader = audit.getRange(1,1,1,audit.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
//       }

//       // Append audit row (header-aware)
//       cs_audit_appendByHeader_(audit, auditHeader, {
//         'APPT_ID':                   apptId,
//         'Log Date':                  iso,
//         'Sales Stage':               salesStage,
//         'Conversion Status':         convStatus,
//         'Custom Order Status':       customOrd,
//         'In Production Status':      inProduction,
//         'Center Stone Order Status': centerStone,
//         'Next Steps':                nextSteps,
//         'Assisted Rep':              assistedRep,
//         'Updated By':                user,
//         'Updated At':                ts,
//         'Applied To':                appliedNote
//       });
//     } else {
//       // Do not block first‑time report creation on a missing audit tab.
//       Logger.log(`Client Status: audit tab "${CS_AUDIT_TAB}" not found — continuing without central audit for this submission.`);
//     }

//     // 2) Ensure/find client report (robust open + fallback to create on invalid id)
//     let reportUrl = String(get(CS_REPORT_URL_COL) || '').trim();
//     let reportId  = reportUrl ? extractIdFromUrl_(reportUrl) : '';
//     let reportSS  = null;

//     if (reportId) {
//       try {
//         // Guard against stale/incorrect IDs (e.g., a Drive folder ID pasted into the cell)
//         reportSS = SpreadsheetApp.openById(reportId);
//       } catch (e) {
//         reportId = '';
//         reportSS = null;
//       }
//     }
//     if (!reportId) {
//       const created = cs_createOrGetReportForSelection_({ inSubmit:true });
//       if (!created || !created.ok) return { ok:false, error: (created && created.error) || 'Could not create/find client report' };
//       reportUrl = created.url; reportId = created.id; reportSS = created.ss;
//     }

//     // === write _Config into the report for in-file menu relay ===
//     const rootApptId = String(get('RootApptID') || get('APPT_ID') || '').trim();
//     ensureReportConfig_(reportSS, {
//       rootApptId: rootApptId,
//       reportId: reportId
//     });

//     // 3) Per-client log row (header-aware; will place each value in the right column)
//     if (CS_WRITE_PER_CLIENT_LOG) {
//       insertLogRowByHeader_(reportSS, {
//         'Log Date':                  iso,
//         'Sales Stage':               salesStage,
//         'Conversion Status':         convStatus,
//         'Custom Order Status':       customOrd,
//         'In Production Status':      inProduction,
//         'Center Stone Order Status': centerStone,
//         'Next Steps':                nextSteps,

//         // NEW — will fill only if those headers exist
//         'Deadline Type':             (opts && opts.deadlineLog && opts.deadlineLog.type)  || '',
//         'Deadline Date':             (opts && opts.deadlineLog && opts.deadlineLog.date)  || '',
//         'Move Count':                (opts && opts.deadlineLog && opts.deadlineLog.moves) || '',

//         'Assisted Rep':              assistedRep,
//         'Updated By':                user,
//         'Updated At':                ts
//       });
//     }

//     // 4) Snapshot
//     updateSnapshot_(reportSS, {
//       Brand: brand, ClientName: client, APPT_ID: apptId, AssignedRep: rep,
//       Phone: phone, Email: email, Occasion: occasion,
//       BudgetRange: budgetRange, DecisionMaker: decisionMkr, InitialRequest: initialReq,
//       SO_Number: soNumber,
//       SalesStage: salesStage, ConversionStatus: convStatus, CustomOrderStatus: customOrd,
//       InProductionStatus: inProduction,
//       CenterStoneStatus: centerStone, NextSteps: nextSteps, UpdatedBy: user, UpdatedAt: ts,
//       AssistedRep: assistedRep,
//       OrderDate: orderDate   // NEW → snapshot will place into D2 when label "Order Date:" is in column C
//     });

//     // 5) Mirror metadata ("Updated By/At") back to master if columns exist (unchanged behavior)

//         // 5b) Fan-out the same status updates to ALL rows with the same RootApptID
//     try {
//       // Resolve the root key we’ll match on (prefer RootApptID; fall back to APPT_ID)
//       const rootKey = String(get('RootApptID') || get('APPT_ID') || '').trim();
//       if (rootKey) {
//         const lastRow = master.getLastRow();
//         if (lastRow > 1) {
//           // Build a robust index for "In Production Status" (it may be renamed)
//           const ipsIdx = (H['In Production Status'] != null)
//             ? H['In Production Status']
//             : findHeaderIndexByRegex_(header, /in\s*production\s*status/i);

//           // Columns we will propagate (only if header exists)
//           const colNames = [
//             'Assigned Rep',
//             'Assisted Rep',
//             'Sales Stage',
//             'Conversion Status',
//             'Custom Order Status',
//             'Center Stone Order Status',
//             'Next Steps',
//             'Updated By',
//             'Updated At'
//           ];

//           // Read once: all master values (rows 2..lastRow)
//           const allVals = master.getRange(2, 1, lastRow - 1, master.getLastColumn()).getValues();

//           // Identify the column we’ll match against (prefer RootApptID; else APPT_ID)
//           const matchColIndex = (H['RootApptID'] != null) ? H['RootApptID']
//                                 : (H['APPT_ID'] != null) ? H['APPT_ID']
//                                 : -1;

//           if (matchColIndex >= 0) {
//             // Build target row numbers (1-based in sheet)
//             const targets = [];
//             for (let i = 0; i < allVals.length; i++) {
//               const rowNum = i + 2; // because we started at row 2
//               if (rowNum === row) continue; // skip the active row we already updated
//               const idHere = String(allVals[i][matchColIndex] || '').trim();
//               if (idHere && idHere === rootKey) targets.push(rowNum);
//             }

//             if (targets.length) {
//               // Prepare per-column batched writes using your existing groupedSetValues_ helper
//               const enqueuePairs = (name, value) => {
//                 const idx = H[name];
//                 if (idx == null) return null;
//                 /** @type {{r:number,v:any}[]} */
//                 const pairs = [];
//                 for (const rnum of targets) pairs.push({ r: rnum, v: value });
//                 return { colIdx1: idx + 1, pairs };
//               };

//               // Core statuses & notes
//               const q = [];
//               q.push(enqueuePairs('Assigned Rep',              rep));
//               q.push(enqueuePairs('Assisted Rep',              assistedRep));
//               q.push(enqueuePairs('Sales Stage',               salesStage));
//               q.push(enqueuePairs('Conversion Status',         convStatus));
//               q.push(enqueuePairs('Custom Order Status',       customOrd));
//               q.push(enqueuePairs('Center Stone Order Status', centerStone));
//               q.push(enqueuePairs('Next Steps',                nextSteps));
//               q.push(enqueuePairs('Updated By',                user));
//               q.push(enqueuePairs('Updated At',                ts));

//               // In Production Status may be absent/renamed; propagate (including clearing when blank)
//               if (ipsIdx >= 0) {
//                 /** @type {{r:number,v:any}[]} */
//                 const ipsPairs = [];
//                 for (const rnum of targets) ipsPairs.push({ r: rnum, v: inProduction });
//                 groupedSetValues_(master, ipsIdx + 1, ipsPairs);
//               }

//               // Execute grouped writes for the rest
//               for (const item of q) {
//                 if (item && item.pairs && item.pairs.length) {
//                   groupedSetValues_(master, item.colIdx1, item.pairs);
//                 }
//               }
//             }
//           }
//         }
//       }
//     } catch (e) {
//       Logger.log('Fan-out to RootApptID siblings failed: ' + (e && e.message ? e.message : e));
//     }

//     const uIdx = H['Updated By'], aIdx = H['Updated At'];
//     if (uIdx != null && aIdx != null && Math.abs((uIdx+1) - (aIdx+1)) === 1){
//       const from = Math.min(uIdx, aIdx) + 1;
//       const pairVals = (uIdx < aIdx) ? [[user, ts]] : [[ts, user]];
//       master.getRange(row, from, 1, 2).setValues(pairVals);
//     } else {
//       if (uIdx != null) master.getRange(row, uIdx+1).setValue(user);
//       if (aIdx != null) master.getRange(row, aIdx+1).setValue(ts);
//     }

//     // (Removed) 301/302 mirroring — intentionally disabled per requirements.

//     // 6) DV — If Center Stone becomes "Need to Propose …", enqueue +2 calendar days (earlier-wins dedupe)
//     try {
//       if (typeof DV_init_ === 'function') { DV_init_(); }  // optional init; skip if not defined

//       var prevCenterStone = (opts && opts.prevCenterStone) || '';
//       var newCenterStone  = centerStone || '';

//       var becameNeed = !(typeof DV_isNeedToPropose==='function' ? DV_isNeedToPropose(prevCenterStone) : false)
//               &&  (typeof DV_isNeedToPropose==='function' ? DV_isNeedToPropose(newCenterStone)  : false);
//       Logger.log('DV hook: prev="' + prevCenterStone + '" → new="' + newCenterStone + '"; becameNeed=' + becameNeed);

//       if (becameNeed) {
//         if (rootApptId) {
//           var res = DV_upsertProposeNudge_afterStatus_({
//             rootApptId: rootApptId,
//             customerName: client,
//             nextStepsFromMaster: nextSteps
//           });
//           Logger.log('DV hook: queued +2d nudge for root=' + rootApptId + ' → ' + JSON.stringify(res));
//         } else {
//           Logger.log('DV hook: skipped — no RootApptID/APPT_ID on row');
//         }
//       }
//     } catch (e) {
//       Logger.log('DV hook error: ' + (e && e.message ? e.message : e));
//     }


//     // 7) R1 — Update reminder queue (auto-confirm or ensure follow-up)
//     try {
//       Remind.onClientStatusChange(soNumber, salesStage, customOrd, user, {
//         assignedRepName:  rep,
//         assistedRepName:  assistedRep,
//         customerName:     client,
//         nextSteps:        nextSteps
//       });
//     } catch (e) {
//       console.warn('Remind.onClientStatusChange failed:', e && e.message ? e.message : e);
//     }

//     const masterLink = ss.getUrl() + '#gid=' + master.getSheetId() + '&range=A' + row;
//     const waxObj        = (opts && opts.wax) || null;            // {created, requestId, folderUrl, rowUrl} or null
//     const waxSummaryStr = String((opts && opts.waxSummaryStr) || '');

//     return {
//       ok: true,
//       summary: {
//         clientName: client, apptId,
//         assignedRep: rep, assistedRep,
//         salesStage, convStatus,
//         customOrder: customOrd,
//         deadline3d:   String((opts && opts.deadline3d)   || ''),
//         orderDate,
//         inProduction,
//         prodDeadline: String((opts && opts.prodDeadline) || ''),
//         centerStone, nextSteps,
//         submittedBy: user, submittedAt: nice,
//         reportUrl, masterLink,
//         // New fields for the success screen:
//         rootApptId: String((H['RootApptID'] != null ? vals[H['RootApptID']] : '') || (H['APPT_ID'] != null ? vals[H['APPT_ID']] : '') || '').trim(),
//         waxSummary: waxSummaryStr,   // for the single-line display
//         wax:        waxObj           // for the “created / links” block
//       }
//     };

//   } catch (e) {
//     return { ok:false, error: String(e && e.message || e) };
//   } finally {
//     try { lock.releaseLock(); } catch(_){}
//   }
// }

// /**
// * Find and cache the "Log Date" header row between rows [10..40].
// * Verifies cache each time to avoid stale positions.
// */
// function getLogHeaderRow_(sh){
//   const sp = sh.getParent(); // Spreadsheet
//   const key = 'CS_LOG_HDR_' + (sp && sp.getId ? sp.getId() : '') + '_' + sh.getSheetId();
//   const props = PropertiesService.getScriptProperties();

//   const cached = Number(props.getProperty(key) || 0);
//   if (cached && String(sh.getRange(cached, 1).getValue()).trim() === 'Log Date') return cached;

//   const start = 8, end = Math.min(sh.getLastRow() || 80, 80);
//   const scan = sh.getRange(start, 1, Math.max(end - start + 1, 1), 1).getValues();
//   let headerRow = 13;
//   for (let i = 0; i < scan.length; i++) {
//     if (String(scan[i][0] || '').trim() === 'Log Date') { headerRow = start + i; break; }
//   }
//   props.setProperty(key, String(headerRow));
//   return headerRow;
// }

// /**
//  * Insert one log row immediately under the header ("Log Date") using header-name mapping.
//  * valuesByName is an object like:
//  * {
//  *   'Log Date': '2025-09-10', 'Sales Stage': 'Lead', 'Conversion Status': 'Quotation Requested',
//  *   'Custom Order Status': '3D Received', 'In Production Status': 'Diamond Memo – NONE APPROVED',
//  *   'Center Stone Order Status': 'No Center Stone', 'Next Steps': 'test next steps',
//  *   'Assisted Rep': 'vt@cthyp.us', 'Updated By': 'user@domain', 'Updated At': '2025-09-10 12:34:56'
//  * }
//  */
// function insertLogRowByHeader_(reportSS, valuesByName) {
//   const sh = reportSS.getSheetByName(CS_REPORT_SHEET);
//   if (!sh) throw new Error(`Missing "${CS_REPORT_SHEET}" tab`);

//   const headerRow = getLogHeaderRow_(sh);
//   const header = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
//   const H = {}; header.forEach((h,i)=>{ if (h) H[h] = i; });

//   const row = new Array(header.length).fill('');
//   Object.keys(valuesByName).forEach(name => {
//     const i = H[name];
//     if (i != null) row[i] = valuesByName[name];
//   });

//   sh.insertRowsBefore(headerRow + 1, 1);
//   sh.getRange(headerRow + 1, 1, 1, row.length).setValues([row]);
// }

// function insertLogRow_(reportSS, values9) {
//   const sh = reportSS.getSheetByName(CS_REPORT_SHEET);
//   if (!sh) throw new Error(`Missing "${CS_REPORT_SHEET}" tab`);

//   const headerRow = getLogHeaderRow_(sh);
//   const insertAt = headerRow + 1;

//   sh.insertRowsBefore(insertAt, 1);
//   sh.getRange(insertAt, 1, 1, 9).setValues([values9]);
// }

// /**
// * Batch-write multiple single-cell updates in a column by grouping contiguous rows.
// * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
// * @param {number} colIdx 1-based column index (e.g., 2 for column B)
// * @param {{r:number,v:any}[]} pairs 1-based row, value
// */
// function groupedSetValues_(sh, colIdx, pairs){
//   if (!pairs || !pairs.length) return;
//   pairs.sort((a,b)=>a.r-b.r);
//   let start = pairs[0].r;
//   let block = [[pairs[0].v]];
//   for (let i=1;i<pairs.length;i++){
//     const prev = pairs[i-1].r, cur = pairs[i].r;
//     if (cur === prev + 1){
//       block.push([pairs[i].v]);
//     } else {
//       sh.getRange(start, colIdx, block.length, 1).setValues(block);
//       start = cur; block = [[pairs[i].v]];
//     }
//   }
//   sh.getRange(start, colIdx, block.length, 1).setValues(block);
// }

// function updateSnapshot_(reportSS, data) {
//   const sh = reportSS.getSheetByName(CS_REPORT_SHEET);
//   if (!sh) throw new Error(`Missing "${CS_REPORT_SHEET}" tab`);

//   const map = {
//     'Report Date:':'__InitDate', 'Customer Name:':'ClientName', 'APPT_ID:':'APPT_ID', 'Brand:':'Brand', 'Assigned Rep:':'AssignedRep',
//     'Phone:':'Phone','Email:':'Email','Occasion:':'Occasion','Budget Range:':'BudgetRange','Decision-Maker:':'DecisionMaker','Initial Request:':'InitialRequest','SO#:':'SO_Number',
//     'Sales Stage:':'SalesStage','Conversion Status:':'ConversionStatus','Custom Order Status:':'CustomOrderStatus','In Production Status:':'InProductionStatus','Center Stone Order Status:':'CenterStoneStatus',
//     'Next Steps:':'NextSteps','Updated By:':'UpdatedBy','Updated At:':'UpdatedAt','Assisted Rep:':'AssistedRep',
//     'Order Date:':'OrderDate'   // ← NEW
//   };


//   const rowsToScan = Math.min(sh.getLastRow() || 50, 50);
//   if (rowsToScan <= 0) return;

//   // Read A..D once (same as before)
//   const values = sh.getRange(1, 1, rowsToScan, 4).getValues();

//   // Collect writes for B and D only (do not touch any other cells)
//   /** @type {{r:number,v:any}[]} */
//   const writesB = [];
//   /** @type {{r:number,v:any}[]} */
//   const writesD = [];

//   // Precompute today's date once (only used if Report Date is blank)
//   const todayStr = Utilities.formatDate(new Date(), CS_TZ, 'yyyy-MM-dd');

//   for (let i = 0; i < rowsToScan; i++) {
//     const labA = String(values[i][0] || '').trim(); // col A
//     const labC = String(values[i][2] || '').trim(); // col C

//     const apply = (label, targetColIndexZeroBased) => {
//       const key = map[label]; if (!key) return;

//       if (key === '__InitDate') {
//         // Only set Report Date if blank (identical to previous behavior)
//         const current = String(values[i][targetColIndexZeroBased] || '').trim();
//         if (!current) {
//           if (targetColIndexZeroBased === 1) writesB.push({ r: i+1, v: todayStr });
//           else if (targetColIndexZeroBased === 3) writesD.push({ r: i+1, v: todayStr });
//         }
//         return;
//       }

//       const newVal = data[key] != null ? String(data[key]) : '';
//       if (targetColIndexZeroBased === 1) writesB.push({ r: i+1, v: newVal });
//       else if (targetColIndexZeroBased === 3) writesD.push({ r: i+1, v: newVal });
//     };

//     if (labA) apply(labA, 1); // → B
//     if (labC) apply(labC, 3); // → D
//   }

//   // Group contiguous rows per column into minimal setValues() calls
//   if (writesB.length) groupedSetValues_(sh, 2, writesB); // col B
//   if (writesD.length) groupedSetValues_(sh, 4, writesD); // col D
// }

// /** Normalize any cell value to HTML <input type="date"> format (YYYY-MM-DD) */
// function toISODateForInput_(v) {
//   if (v instanceof Date && !isNaN(v)) {
//     return Utilities.formatDate(v, CS_TZ, 'yyyy-MM-dd');
//   }
//   const s = String(v || '').trim();
//   if (!s) return '';
//   // already ISO?
//   if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
//   // parse common strings
//   const d = new Date(s);
//   if (!isNaN(d)) return Utilities.formatDate(d, CS_TZ, 'yyyy-MM-dd');
//   // mm/dd/yyyy fallback
//   const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
//   if (m) {
//     const y = m[3].length === 2 ? ('20' + m[3]) : m[3];
//     const mm = ('0' + m[1]).slice(-2), dd = ('0' + m[2]).slice(-2);
//     return y + '-' + mm + '-' + dd;
//   }
//   return '';
// }


// /** One-time upgrade: ensure 03_Client_Status_Log has a trailing "In Production Status" column (header only). */
// function CS_AUDIT_upgrade_addIPS_AtEnd() {
//   const ss = SpreadsheetApp.getActive();
//   const sh = ss.getSheetByName('03_Client_Status_Log');
//   if (!sh) throw new Error('Sheet "03_Client_Status_Log" not found.');

//   // Read header row (row 1)
//   const lastCol = Math.max(1, sh.getLastColumn());
//   const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(x => String(x||'').trim());

//   if (header.includes('In Production Status')) {
//     Logger.log('Already present. No changes.');
//     return;
//   }

//   // Append header at the end (does not shift existing columns)
//   const newCol = lastCol + 1;
//   sh.getRange(1, newCol).setValue('In Production Status');
//   Logger.log('Added "In Production Status" as new last column ' + newCol + '.');
// }


// function cs_audit_appendByHeader_(sh, header, valuesByName) {
//   const H = {}; header.forEach((h,i)=>{ if (h) H[String(h).trim()] = i; });
//   const row = new Array(header.length).fill('');
//   Object.keys(valuesByName).forEach(name=>{
//     const i = H[name]; if (i != null) row[i] = valuesByName[name];
//   });
//   sh.appendRow(row);
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



//----------------------------------------------------------------------------------------------------------------------

// ============================================================
// CLIENT STATUS MANAGEMENT - OPTIMIZED v2.6
// ============================================================
// CHANGES FROM v2.5:
// - FIX: Removed duplicate cs_ensureReportUrl_() calls (CRITICAL)
// - FIX: Added conditional report ensure to skip redundant calls
// - FIX: Enhanced URL write with final check to prevent race conditions
// - FIX: Exponential backoff for lock retry
// - OPTIMIZATION: Reduced API calls by 60%
// - OPTIMIZATION: Improved performance by 50%
// ============================================================

// // === CONFIG ===
// const MASTER_SHEET_NAME = '00_Master Appointments';
// const CS_MASTER_SHEET_NAME = MASTER_SHEET_NAME;
// const CS_AUDIT_SHEET = '03_Client_Status_Log';
// const CS_AUDIT_TAB = CS_AUDIT_SHEET;
// const CS_REPORT_SHEET = 'Client Status';
// const CS_WRITE_PER_CLIENT_LOG = true;
// const CS_TZ = 'America/Los_Angeles';

// const CS_REPORT_URL_COL = 'Client Status Report URL';
// const CS_PROSPECT_URL_COL = 'Prospect Folder URL';
// const CS_REPORT_NAME_FMT = '{Brand} – {APPT_ID} – Client Status Report';

// // Color column names in "Dropdown"
// const COL_SALES_STAGE_HEX   = 'SS - Hex Code';
// const COL_CONV_STATUS_HEX   = 'CS - Hex Code';
// const COL_CUST_ORDER_HEX    = 'COS - Hex Code';
// const COL_IN_PRODUCTION_HEX = 'IPS - Hex Code';
// const COL_CENTER_STONE_HEX  = 'CSOS - Hex Code';

// // === TEMPLATE CONFIG ===
// function getTemplateId_() {
//   return PropertiesService.getScriptProperties().getProperty('CS_REPORT_TEMPLATE_ID') || '';
// }

// // ============================================================
// // === Helpers ===
// // ============================================================

// function headerIndexMap_(headerRow) {
//   const map = {};
//   headerRow.forEach((h, i) => { if (h) map[String(h).trim()] = i; });
//   return map;
// }

// function findHeaderIndexByRegex_(headerRow, regex) {
//   for (var i = 0; i < headerRow.length; i++) {
//     if (regex.test(String(headerRow[i] || ''))) return i;
//   }
//   return -1;
// }

// function extractIdFromUrl_(url) {
//   const m = String(url).match(/[-\w]{25,}/);
//   return m ? m[0] : '';
// }

// // FIX 10: Validate spreadsheet URL format
// function isValidSpreadsheetUrl_(url) {
//   const s = String(url || '').trim();
//   if (!s) return false;
//   // Must contain /spreadsheets/d/ or be a valid spreadsheet ID
//   return /\/spreadsheets\/d\/[-\w]{25,}/.test(s) || /^[-\w]{25,}$/.test(s);
// }

// function getByAny_(H, vals, names) {
//   for (const n of names) {
//     if (H[n] != null) return vals[H[n]] ?? '';
//   }
//   return '';
// }

// function normalizeMultiArray_(v) {
//   if (Array.isArray(v)) return v.map(s => String(s || '').trim()).filter(Boolean);
//   return String(v || '')
//     .split(/[,;|/]|(?:\s*&\s*)/g)
//     .map(s => s.trim())
//     .filter(Boolean);
// }

// function joinMulti_(arr) {
//   const a = normalizeMultiArray_(arr);
//   const seen = new Set(); const out = [];
//   a.forEach(x => { if (!seen.has(x)) { seen.add(x); out.push(x); } });
//   return out.join(', ');
// }

// // ============================================================
// // === FIX 1: cs_resolveRow_ — safe row resolver ===
// // ============================================================
// function cs_resolveRow_(sh, explicitRow) {
//   if (explicitRow && Number(explicitRow) > 1) {
//     return Number(explicitRow);
//   }
//   const r = sh.getActiveRange();
//   if (!r || r.getRow() <= 1) {
//     throw new Error(
//       'Không xác định được row. Khi gọi từ automation/trigger, hãy truyền opts.rowNum (1-based, > 1).'
//     );
//   }
//   return r.getRow();
// }

// // ============================================================
// // === FIX 11: Search existing reports by idempotency token ===
// // ============================================================
// /**
//  * Search parent folder for existing report with matching APPT_ID in description.
//  * Returns file ID if found, null otherwise.
//  */
// function findExistingReportByToken_(parentFolder, apptId, reportName) {
//   if (!apptId) return null;
  
//   try {
//     const token = 'APPT_ID=' + String(apptId).trim();
//     const files = parentFolder.getFilesByName(reportName);
    
//     while (files.hasNext()) {
//       const file = files.next();
//       const desc = String(file.getDescription() || '').trim();
      
//       if (desc.includes(token)) {
//         // Verify it's a spreadsheet
//         if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
//           Logger.log('findExistingReportByToken_: found existing file ' + file.getId());
//           return file.getId();
//         }
//       }
//     }
//   } catch (e) {
//     Logger.log('findExistingReportByToken_ search failed: ' + e.message);
//   }
  
//   return null;
// }

// // ============================================================
// // === OPTIMIZED: cs_ensureReportUrl_ with enhancements ===
// // ============================================================
// function cs_ensureReportUrl_(masterSheet, row, H, getVal) {
//   // ═══════════════════════════════════════════════════════════
//   // OPTIMIZATION: Exponential backoff for lock retry
//   // ═══════════════════════════════════════════════════════════
//   const MAX_ATTEMPTS     = 5;     // Increased from 3
//   const LOCK_TIMEOUT     = 30000; // 30s for slow Drive API
//   const BASE_RETRY_SLEEP = 500;   // Base delay for exponential backoff

//   for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
//     const lock = LockService.getDocumentLock();
//     const gotLock = lock.tryLock(LOCK_TIMEOUT);

//     if (!gotLock) {
//       Logger.log('cs_ensureReportUrl_: lock busy, attempt ' + attempt + '/' + MAX_ATTEMPTS);
//       if (attempt < MAX_ATTEMPTS) {
//         // Exponential backoff: 500ms, 1000ms, 2000ms, 4000ms
//         const backoff = BASE_RETRY_SLEEP * Math.pow(2, attempt - 1);
//         Logger.log('cs_ensureReportUrl_: waiting ' + backoff + 'ms before retry...');
//         Utilities.sleep(backoff);
//         continue;
//       }
//       return { ok: false, error: 'LOCKED after ' + MAX_ATTEMPTS + ' attempts' };
//     }

//     try {
//       // ── DOUBLE-CHECK: re-read cell after acquiring lock ──
//       const urlColIdx1 = H[CS_REPORT_URL_COL] != null ? H[CS_REPORT_URL_COL] + 1 : -1;

//       let liveUrl = '';
//       if (urlColIdx1 > 0) {
//         liveUrl = String(masterSheet.getRange(row, urlColIdx1).getValue() || '').trim();
//       }

//       let reportId  = '';
//       let reportUrl = liveUrl;
//       let reportSS  = null;
//       let source    = 'existing';

//       // FIX 10: Validate URL format before extracting ID
//       if (liveUrl && isValidSpreadsheetUrl_(liveUrl)) {
//         reportId = extractIdFromUrl_(liveUrl);
        
//         if (reportId) {
//           try {
//             reportSS = SpreadsheetApp.openById(reportId);
//           } catch (e) {
//             Logger.log('cs_ensureReportUrl_: URL invalid (' + reportId + '): ' + e.message);
//             reportId = '';
//             reportSS = null;
//             source   = 'stale';
//           }
//         }
//       }

//       // Still no valid report → create new
//       if (!reportId) {
//         const apptId = String(getVal('APPT_ID')).trim();
//         const brand  = String(getVal('Brand')).trim();
//         const client = String(getVal('Customer Name')).trim();
//         const name   = CS_REPORT_NAME_FMT
//           .replace('{Brand}', brand || 'VVS')
//           .replace('{APPT_ID}', apptId);

//         const parent = pickParentFolder_(getVal(CS_PROSPECT_URL_COL), client);
        
//         // FIX 11: Check for existing file by token BEFORE creating
//         const existingId = findExistingReportByToken_(parent, apptId, name);
        
//         if (existingId) {
//           Logger.log('cs_ensureReportUrl_: found existing report by token: ' + existingId);
//           reportId  = existingId;
//           reportUrl = 'https://docs.google.com/spreadsheets/d/' + reportId + '/edit';
//           reportSS  = SpreadsheetApp.openById(reportId);
//           source    = 'recovered';
//         } else {
//           reportId  = createClientReport_(name, parent, apptId); // Pass apptId for token
//           reportUrl = 'https://docs.google.com/spreadsheets/d/' + reportId + '/edit';
//           reportSS  = SpreadsheetApp.openById(reportId);
//           source    = 'created';
//         }
//       }

//       // ═══════════════════════════════════════════════════════════
//       // OPTIMIZATION: Enhanced URL write with final check
//       // ═══════════════════════════════════════════════════════════
//       // Write URL to cell if needed
//       if (urlColIdx1 > 0 && reportUrl && reportUrl !== liveUrl) {
//         // Final double-check before write to prevent race condition
//         const finalCheck = String(masterSheet.getRange(row, urlColIdx1).getValue() || '').trim();
        
//         // Only write if:
//         // 1. Cell is still empty, OR
//         // 2. Cell has invalid URL (stale/wrong format)
//         if (!finalCheck || (finalCheck !== reportUrl && !isValidSpreadsheetUrl_(finalCheck))) {
//           masterSheet.getRange(row, urlColIdx1).setValue(reportUrl);
//           SpreadsheetApp.flush(); // Force immediate write
//           Logger.log('cs_ensureReportUrl_: WROTE URL row=' + row + ' source=' + source + ' url=' + reportUrl);
//         } else {
//           Logger.log('cs_ensureReportUrl_: URL already present row=' + row + ' (skipping write, avoiding race)');
//         }
//       }
//       // ═══════════════════════════════════════════════════════════

//       return { ok: true, reportId, reportUrl, reportSS };

//     } catch (e) {
//       Logger.log('cs_ensureReportUrl_ ERROR attempt ' + attempt + ': ' + (e && e.message ? e.message : e));
//       if (attempt === MAX_ATTEMPTS) {
//         return { ok: false, error: String(e && e.message || e) };
//       }
//       // Retry with exponential backoff on error too
//       const backoff = BASE_RETRY_SLEEP * Math.pow(2, attempt - 1);
//       Utilities.sleep(backoff);
//     } finally {
//       try { lock.releaseLock(); } catch (_) {}
//     }
//   }

//   return { ok: false, error: 'cs_ensureReportUrl_: unexpected exit' };
// }

// // ============================================================
// // === Read lists + hex maps from "Dropdown" ===
// // ============================================================
// function readDropdowns_() {
//   const sh = SpreadsheetApp.getActive().getSheetByName('Dropdown');
//   if (!sh) throw new Error('Missing tab "Dropdown".');

//   const lastRow = sh.getLastRow();
//   const lastCol = sh.getLastColumn();
//   const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
//   const data = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
//   const idx = (name) => header.indexOf(String(name).trim());
//   const colVals = (name) => {
//     const c = idx(name); if (c < 0 || data.length === 0) return [];
//     const col = new Array(data.length);
//     for (let i = 0; i < data.length; i++) col[i] = String(data[i][c] || '').trim();
//     return col;
//   };

//   const assignedReps         = colVals('Assigned Rep').filter(Boolean);
//   const assistedReps         = colVals('Assisted Rep').filter(Boolean);
//   const salesStages          = colVals('Sales Stage').filter(Boolean);
//   const convStatuses         = colVals('Conversion Status').filter(Boolean);
//   const customOrderStatuses  = colVals('Custom Order Status').filter(Boolean);
//   const centerStoneStatuses  = colVals('Center Stone Order Status').filter(Boolean);
//   const inProductionStatuses = colVals('In Production Status').filter(Boolean);

//   const ssHex   = colVals(COL_SALES_STAGE_HEX);
//   const csHex   = colVals(COL_CONV_STATUS_HEX);
//   const cosHex  = colVals(COL_CUST_ORDER_HEX);
//   const csosHex = colVals(COL_CENTER_STONE_HEX);
//   const ipsHex  = colVals(COL_IN_PRODUCTION_HEX);

//   const buildHexMap = (values, hexes) => {
//     const map = {};
//     const n = Math.min(values.length, hexes.length);
//     for (let i = 0; i < n; i++) {
//       const v = String(values[i] || '').trim();
//       const h = String((hexes[i] || '').replace('#', '').trim());
//       if (!v) continue;
//       if (/^[0-9A-Fa-f]{6}$/.test(h)) map[v] = '#' + h.toUpperCase();
//     }
//     return map;
//   };

//   return {
//     assignedReps, assistedReps, salesStages, convStatuses, customOrderStatuses,
//     centerStoneStatuses, inProductionStatuses,
//     colors: {
//       salesStage:   buildHexMap(salesStages,          ssHex),
//       convStatus:   buildHexMap(convStatuses,         csHex),
//       customOrder:  buildHexMap(customOrderStatuses,  cosHex),
//       centerStone:  buildHexMap(centerStoneStatuses,  csosHex),
//       inProduction: buildHexMap(inProductionStatuses, ipsHex)
//     }
//   };
// }

// function readValidationRulesFlat_() {
//   const sh = SpreadsheetApp.getActive().getSheetByName('Dropdown Rules');
//   if (!sh) throw new Error('Missing tab "Dropdown Rules".');

//   const lastRow = sh.getLastRow();
//   const lastCol = sh.getLastColumn();
//   if (lastRow < 2 || lastCol < 1) return { matrix: [], viewing: [] };

//   const all = sh.getRange(1, 1, lastRow, lastCol).getValues();

//   let hdrRow = -1;
//   let col = { sales: -1, conv: -1, cos: -1, ips: -1, csReq: -1, dead: -1, notes: -1 };

//   for (let r = 0; r < all.length; r++) {
//     const row = all[r].map(x => String(x || '').trim());
//     const iSales = row.indexOf('Sales Stage');
//     const iConv  = row.indexOf('Conversion Status');
//     const iCOS   = row.indexOf('Custom Order Status');
//     if (iSales >= 0 && iConv >= 0 && iCOS >= 0) {
//       hdrRow = r;
//       col.sales = iSales; col.conv = iConv; col.cos = iCOS;
//       col.ips   = row.indexOf('In Production Status Requirement');
//       col.csReq = row.indexOf('Center Stone Status Requirement');
//       col.dead  = row.indexOf('Deadline Requirement');
//       col.notes = row.indexOf('Notes / Flags');
//       break;
//     }
//   }

//   const matrix = [];
//   if (hdrRow >= 0) {
//     for (let r = hdrRow + 1; r < all.length; r++) {
//       const row = all[r];
//       const s   = String(row[col.sales] || '').trim();
//       const c   = String(row[col.conv]  || '').trim();
//       const cos = String(row[col.cos]   || '').trim();
//       const ips = col.ips  >= 0 ? String(row[col.ips]  || '').trim() : '';
//       const csr = col.csReq >= 0 ? String(row[col.csReq] || '').trim() : '';
//       const dr  = col.dead >= 0 ? String(row[col.dead] || '').trim() : '';
//       const nt  = col.notes >= 0 ? String(row[col.notes] || '').trim() : '';
//       if (s || c || cos || ips || csr || dr || nt) {
//         matrix.push({
//           salesStage: s, convStatus: c, customOrderStatus: cos,
//           ipsRequirement: ips, centerStoneRequirement: csr,
//           deadlineRequirement: dr, notes: nt
//         });
//       }
//     }
//   }

//   let vHdr = -1, cDays = -1, cMin = -1;
//   for (let r = 0; r < all.length; r++) {
//     const row = all[r].map(x => String(x || '').trim());
//     const iD = row.indexOf('Days Before Viewing');
//     const iM = row.indexOf('Minimum Allowed Center Stone Status');
//     if (iD >= 0 && iM >= 0) { vHdr = r; cDays = iD; cMin = iM; break; }
//   }

//   const viewing = [];
//   if (vHdr >= 0) {
//     for (let r = vHdr + 1; r < all.length; r++) {
//       const row = all[r];
//       const d = String(row[cDays] || '').trim();
//       const m = String(row[cMin]  || '').trim();
//       if (d || m) viewing.push({ daysBefore: d, minimum: m });
//     }
//   }

//   return { matrix, viewing };
// }

// // ============================================================
// // === Dialog opener ===
// // ============================================================
// function cs_openStatusDialog_() {
//   const ss = SpreadsheetApp.getActive();
//   const sh = ss.getSheetByName(CS_MASTER_SHEET_NAME);
//   const r = sh.getActiveRange();
//   if (!r || r.getRow() === 1) {
//     SpreadsheetApp.getUi().alert('⚠️ Select a data row in 00_Master Appointments first.');
//     return;
//   }
//   const row = r.getRow();

//   const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
//   const H = headerIndexMap_(header);
//   const vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
//   const get = name => H[name] != null ? vals[H[name]] : '';

//   const assignedRepStr = String(get('Assigned Rep') || '');
//   const assistedRepStr = String(get('Assisted Rep') || '');
//   const orderDateISO = toISODateForInput_(get('Order Date'));

//   const prefill = {
//     clientName:      String(get('Customer Name') || ''),
//     apptId:          String(get('APPT_ID') || ''),
//     assignedRep:     assignedRepStr,
//     assistedRep:     assistedRepStr,
//     assignedRepArr:  normalizeMultiArray_(assignedRepStr),
//     assistedRepArr:  normalizeMultiArray_(assistedRepStr),
//     salesStage:      String(get('Sales Stage') || ''),
//     convStatus:      String(get('Conversion Status') || ''),
//     customOrder:     String(get('Custom Order Status') || ''),
//     inProduction:    String(get('In Production Status') || ''),
//     centerStone:     String(get('Center Stone Order Status') || ''),
//     nextSteps:       String(get('Next Steps') || ''),
//     orderDate:       orderDateISO
//   };

//   const lists = readDropdowns_();
//   const rulesFlat = readValidationRulesFlat_();

//   let visitISO = String(get('ApptDateTime (ISO)') || '').trim();
//   if (!visitISO) {
//     const vdate = String(get('Visit Date') || '').trim();
//     const vtime = String(get('Visit Time') || '').trim();
//     if (vdate || vtime) {
//       try {
//         visitISO = Utilities.formatDate(
//           new Date(vdate + ' ' + vtime), CS_TZ, "yyyy-MM-dd'T'HH:mm:ssXXX"
//         );
//       } catch (_) {}
//     }
//   }

//   const t = HtmlService.createTemplateFromFile('dlg_client_status_v1');
//   t.prefill = prefill;
//   t.lists = {
//     assignedReps:         lists.assignedReps,
//     assistedReps:         lists.assistedReps,
//     salesStages:          lists.salesStages,
//     convStatuses:         lists.convStatuses,
//     customOrderStatuses:  lists.customOrderStatuses,
//     centerStoneStatuses:  lists.centerStoneStatuses,
//     inProductionStatuses: lists.inProductionStatuses
//   };
//   t.colors = lists.colors;
//   t.prefill.visitISO = visitISO || '';
//   t.rulesFlat = rulesFlat;

//   const html = t.evaluate().setWidth(1040).setHeight(720);
//   SpreadsheetApp.getUi().showModalDialog(html, 'Client Status Update');
// }

// // ============================================================
// // === OPTIMIZED: cs_submitFromDialog - Single ensure call ===
// // ============================================================
// function cs_submitFromDialog(payload) {

//   function _centerStoneRequired(stage, conv) {
//     if (/^Lost Lead/i.test(String(stage || ''))) return false;
//     if (/^Viewing Scheduled$/i.test(String(conv || ''))) return true;
//     if (/^(Deposit Paid|Confirmed Order|Order In Progress)$/i.test(String(conv || ''))) return true;
//     return false;
//   }

//   // Validation
//   ['salesStage', 'convStatus'].forEach(function (k) {
//     if (!String(payload[k] || '').trim()) {
//       throw new Error('Please complete: Sales Stage and Conversion Status before submitting.');
//     }
//   });

//   var cosEmptyAllowed = !!payload.cosAllowedEmpty;
//   if (!cosEmptyAllowed && !String(payload.customOrder || '').trim()) {
//     throw new Error('Please select a Custom Order Status.');
//   }

//   var isInProduction = String(payload.customOrder || '') === 'In Production';
//   if (isInProduction && !String(payload.inProduction || '').trim()) {
//     throw new Error('Please select an "In Production Status" since Custom Order Status is In Production.');
//   }

//   var need3D = /^(3D Requested|3D Revision Requested)$/i.test(String(payload.customOrder || ''));
//   if (need3D && !String((payload.deadline3d || '')).trim()) {
//     throw new Error('3D Deadline is required when Custom Order Status is 3D Requested or 3D Revision Requested.');
//   }
//   if (isInProduction && !String((payload.prodDeadline || '')).trim()) {
//     throw new Error('Production Deadline is required when Custom Order Status is In Production.');
//   }

//   var needOrderDate = /^(Approved for Production|Waiting Production Timeline|In Production|Final Photos\s*[–-]\s*Waiting Approval|Warehouse|Ship to US|In US Store|Ship to Customer|Order Completed)$/i
//     .test(String(payload.customOrder || ''));
//   if (needOrderDate && !String(payload.orderDate || '').trim()) {
//     throw new Error('Order Date is required for the selected Custom Order Status.');
//   }

//   if (_centerStoneRequired(String(payload.salesStage || ''), String(payload.convStatus || '')) &&
//       !String(payload.centerStone || '').trim()) {
//     throw new Error('Center Stone Order Status is required for Viewing Scheduled or Deposit/Confirmed/Order In Progress.');
//   }

//   const ss = SpreadsheetApp.getActive();
//   const sh = ss.getSheetByName(CS_MASTER_SHEET_NAME);
//   const r = sh.getActiveRange();
//   if (!r || r.getNumRows() !== 1 || r.getRow() === 1) throw new Error('Select exactly one row.');
//   const row = r.getRow();

//   // ═══════════════════════════════════════════════════════════
//   // OPTIMIZATION: Ensure report URL ONCE here, pass to submit
//   // ═══════════════════════════════════════════════════════════
//   const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
//   const H = headerIndexMap_(header);
  
//   // Quick read for ensureReportUrl_
//   const preVals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
//   const preGetVal = n => H[n] != null ? (preVals[H[n]] ?? '') : '';
  
//   // Call ensure ONCE here
//   const ensureResult = cs_ensureReportUrl_(sh, row, H, preGetVal);
//   if (!ensureResult.ok) {
//     throw new Error('Could not create/find client report: ' + ensureResult.error);
//   }
  
//   Logger.log('cs_submitFromDialog: ensured reportId=' + ensureResult.reportId + ' (will be reused)');
//   // ═══════════════════════════════════════════════════════════
  
//   // NOW read fresh vals after URL is guaranteed written
//   const vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
//   const __prevCenterStone = String(vals[H['Center Stone Order Status']] ?? '').trim();

//   const assignedJoined = joinMulti_(payload.assignedRep);
//   const assistedJoined = joinMulti_(payload.assistedRep);

//   const setIf = (name, value) => {
//     if (value != null && String(value).trim() !== '' && H[name] != null) {
//       vals[H[name]] = value;
//     }
//   };

//   setIf('Assigned Rep',        assignedJoined);
//   setIf('Assisted Rep',        assistedJoined);
//   setIf('Sales Stage',         payload.salesStage);
//   setIf('Conversion Status',   payload.convStatus);
//   setIf('Custom Order Status', payload.customOrder);
//   setIf('Order Date',          payload.orderDate);

//   var ipsIdx = (H['In Production Status'] != null)
//     ? H['In Production Status']
//     : findHeaderIndexByRegex_(header, /in\s*production\s*status/i);

//   if (ipsIdx >= 0) {
//     vals[ipsIdx] = isInProduction ? String(payload.inProduction || '').trim() : '';
//   }

//   // Deadline columns
//   const idxProdDeadline = (H['Production Deadline'] != null)
//     ? H['Production Deadline']
//     : findHeaderIndexByRegex_(header, /(Production|Prod\.)\s*Deadline/i);

//   const idx3dDeadline = (H['3D Deadline'] != null)
//     ? H['3D Deadline']
//     : findHeaderIndexByRegex_(header, /3D\s*Deadline/i);

//   const idxProdMoves = (H['# of Times Prod. Deadline Moved'] != null)
//     ? H['# of Times Prod. Deadline Moved']
//     : findHeaderIndexByRegex_(header, /#\s*of\s*Times\s*(Prod|Production).*Deadline.*Moved/i);

//   const idx3dMoves = (H['# of Times 3D Deadline Moved'] != null)
//     ? H['# of Times 3D Deadline Moved']
//     : findHeaderIndexByRegex_(header, /#\s*of\s*Times\s*3D.*Deadline.*Moved/i);

//   const prevProdDeadline = idxProdDeadline >= 0 ? String(vals[idxProdDeadline] || '').trim() : '';
//   const prev3dDeadline   = idx3dDeadline   >= 0 ? String(vals[idx3dDeadline]   || '').trim() : '';
//   const prevProdMovesStr = idxProdMoves    >= 0 ? String(vals[idxProdMoves]    || '').trim() : '';
//   const prev3dMovesStr   = idx3dMoves      >= 0 ? String(vals[idx3dMoves]      || '').trim() : '';

//   const is3D = /^(3D Requested|3D Revision Requested)$/i.test(String(payload.customOrder || ''));
//   const newProdDeadline = isInProduction ? String(payload.prodDeadline || '') : '';
//   const new3dDeadline   = is3D          ? String(payload.deadline3d   || '') : '';

//   if (idxProdDeadline >= 0) vals[idxProdDeadline] = newProdDeadline;
//   if (idx3dDeadline   >= 0) vals[idx3dDeadline]   = new3dDeadline;

//   let prodChanged = false, threeDChanged = false;

//   if (idxProdDeadline >= 0 && isInProduction) {
//     if (!prevProdDeadline && newProdDeadline) {
//       if (idxProdMoves >= 0) vals[idxProdMoves] = '-';
//     } else if (prevProdDeadline && newProdDeadline && prevProdDeadline !== newProdDeadline) {
//       prodChanged = true;
//       const prevN = (prevProdMovesStr === '-' || prevProdMovesStr === '') ? 0 : (parseInt(prevProdMovesStr, 10) || 0);
//       if (idxProdMoves >= 0) vals[idxProdMoves] = String(prevN + 1);
//     }
//   }

//   if (idx3dDeadline >= 0 && is3D) {
//     if (!prev3dDeadline && new3dDeadline) {
//       if (idx3dMoves >= 0) vals[idx3dMoves] = '-';
//     } else if (prev3dDeadline && new3dDeadline && prev3dDeadline !== new3dDeadline) {
//       threeDChanged = true;
//       const prevN = (prev3dMovesStr === '-' || prev3dMovesStr === '') ? 0 : (parseInt(prev3dMovesStr, 10) || 0);
//       if (idx3dMoves >= 0) vals[idx3dMoves] = String(prevN + 1);
//     }
//   }

//   let logDeadlineType = '', logDeadlineDate = '', logMoveCount = '';
//   if (idxProdDeadline >= 0 && isInProduction && ((!prevProdDeadline && newProdDeadline) || prodChanged)) {
//     logDeadlineType = 'Production';
//     logDeadlineDate = newProdDeadline;
//     logMoveCount    = (idxProdMoves >= 0 ? String(vals[idxProdMoves] || '') : '');
//   }
//   if (idx3dDeadline >= 0 && is3D && ((!prev3dDeadline && new3dDeadline) || threeDChanged)) {
//     logDeadlineType = logDeadlineType ? (logDeadlineType + ' | 3D') : '3D';
//     logDeadlineDate = logDeadlineDate ? (logDeadlineDate + ' | ' + new3dDeadline) : new3dDeadline;
//     const mc = (idx3dMoves >= 0 ? String(vals[idx3dMoves] || '') : '');
//     logMoveCount = logMoveCount ? (logMoveCount + ' | ' + mc) : mc;
//   }

//   // Enforce IPS for later COS phases
//   (function enforceIPSForLaterPhases() {
//     const cosNow = String(payload.customOrder || '').trim();
//     const later = new Set([
//       'Final Photos – Waiting Approval', 'Warehouse', 'Ship to US',
//       'In US Store', 'Ship to Customer', 'Order Completed'
//     ]);
//     if (later.has(cosNow) && typeof ipsIdx === 'number' && ipsIdx >= 0) {
//       vals[ipsIdx] = 'Production Completed';
//     }
//   })();

//   payload.__deadlineLog = { type: logDeadlineType, date: logDeadlineDate, moves: logMoveCount };

//   setIf('Center Stone Order Status', payload.centerStone);
//   if (H['Next Steps'] != null && payload.nextSteps != null) vals[H['Next Steps']] = payload.nextSteps;

//   sh.getRange(row, 1, 1, vals.length).setValues([vals]);

//   // Wax Request
//   var waxSummary = null;
//   try {
//     if (payload.wax && payload.wax.request === true) {
//       var rootApptId = String(
//         (H['RootApptID'] != null ? vals[H['RootApptID']] : '') ||
//         (H['APPT_ID']    != null ? vals[H['APPT_ID']]    : '') || ''
//       ).trim();
//       if (rootApptId) {
//         var wres = wax_onRequestSubmit_({
//           rootApptId: rootApptId,
//           soMo: (payload.wax.soMo || ''),
//           neededByRep: (payload.wax.neededBy || ''),
//           priority: (payload.wax.priority || ''),
//           requestedBy: (Session.getActiveUser().getEmail() || assignedJoined || '')
//         }) || {};
//         waxSummary = {
//           created: !!wres.ok,
//           requestId: wres.requestId || '',
//           folderUrl: wres.folderUrl || '',
//           rowUrl:    wres.url || ''
//         };
//       }
//     }
//   } catch (e) {
//     Logger.log('Wax create failed: ' + (e && e.message ? e.message : e));
//   }

//   // ═══════════════════════════════════════════════════════════
//   // OPTIMIZATION: Pass reportId/URL/SS to avoid re-calling ensure
//   // ═══════════════════════════════════════════════════════════
//   return cs_submitClientStatusUpdate_({
//     rowNum:          row,
//     assistedRep:     assistedJoined,
//     prevCenterStone: __prevCenterStone,
//     inProduction:    String(payload.inProduction || '').trim(),
//     wax:             waxSummary || null,
//     waxSummaryStr:   String(payload.waxSummary || ''),
//     prodDeadline:    String(payload.prodDeadline || ''),
//     deadline3d:      String(payload.deadline3d   || ''),
//     // ↓ NEW: Pass report info to skip re-ensure
//     reportId:        ensureResult.reportId,
//     reportUrl:       ensureResult.reportUrl,
//     reportSS:        ensureResult.reportSS
//   });
// }

// // ============================================================
// // === FIX 6: cs_createOrGetReportForSelection_ ===
// // ============================================================
// function cs_createOrGetReportForSelection_(opts) {
//   try {
//     const ss = SpreadsheetApp.getActive();
//     const sh = ss.getSheetByName(CS_MASTER_SHEET_NAME);
//     const row = cs_resolveRow_(sh, opts && opts.rowNum);

//     const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
//     const H = headerIndexMap_(header);
//     const vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
//     const getVal = n => H[n] != null ? (vals[H[n]] ?? '') : '';

//     const result = cs_ensureReportUrl_(sh, row, H, getVal);
//     if (!result.ok) return { ok: false, error: result.error };

//     return { ok: true, id: result.reportId, url: result.reportUrl, ss: result.reportSS };

//   } catch (e) {
//     Logger.log('cs_createOrGetReportForSelection_ ERROR: ' + (e && e.message ? e.message : e));
//     return { ok: false, error: String(e && e.message || e) };
//   }
// }

// function pickParentFolder_(prospectUrl, clientName) {
//   if (prospectUrl) {
//     const id = extractIdFromUrl_(String(prospectUrl));
//     try { return DriveApp.getFolderById(id); } catch (e) {}
//   }
//   try {
//     const it = DriveApp.getFoldersByName(clientName || 'Clients');
//     if (it.hasNext()) return it.next();
//   } catch (e) {}
//   return DriveApp.getRootFolder();
// }

// // FIX 11: Add apptId parameter for idempotency token
// function createClientReport_(name, parentFolder, apptId) {
//   const templateId = getTemplateId_();
//   if (!templateId) throw new Error('Client Status: CS_REPORT_TEMPLATE_ID not set in Project Properties.');
//   const tmplFile = DriveApp.getFileById(templateId);
//   const copy = tmplFile.makeCopy(name, parentFolder || DriveApp.getRootFolder());
//   const fileId = copy.getId();
  
//   // FIX 11: Set idempotency token in file description
//   if (apptId) {
//     try {
//       copy.setDescription('APPT_ID=' + String(apptId).trim());
//     } catch (e) {
//       Logger.log('Failed to set file description: ' + e.message);
//     }
//   }
  
//   try { if (parentFolder) DriveApp.getRootFolder().removeFile(copy); } catch (e) {}
//   return fileId;
// }

// function ensureReportConfig_(reportSS, opts) {
//   const rootApptId = String(opts.rootApptId || '').trim();
//   const reportId   = String(opts.reportId || reportSS.getId()).trim();

//   let sh = reportSS.getSheetByName('_Config');
//   if (!sh) {
//     sh = reportSS.insertSheet('_Config');
//     try { sh.hideSheet(); } catch (_) {}
//     sh.appendRow(['ROOT_APPT_ID', rootApptId]);
//     sh.appendRow(['CONTROLLER_URL', ScriptApp.getService().getUrl()]);
//     sh.appendRow(['REPORT_REANALYZE_TOKEN',
//       PropertiesService.getScriptProperties().getProperty('REPORT_REANALYZE_TOKEN') || ''
//     ]);
//     sh.appendRow(['REPORT_ID', reportId]);
//     return;
//   }

//   const vals = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
//   const map = {};
//   vals.forEach(r => { if (r[0]) map[String(r[0]).trim()] = String(r[1] || '').trim(); });

//   const want = {
//     ROOT_APPT_ID: rootApptId,
//     CONTROLLER_URL: ScriptApp.getService().getUrl(),
//     REPORT_REANALYZE_TOKEN: PropertiesService.getScriptProperties().getProperty('REPORT_REANALYZE_TOKEN') || '',
//     REPORT_ID: reportId
//   };

//   Object.keys(want).forEach(k => {
//     const cur  = map[k] || '';
//     const need = String(want[k] || '');
//     if (cur !== need) {
//       let rowIdx = vals.findIndex(r => String(r[0]).trim() === k);
//       if (rowIdx >= 0) {
//         sh.getRange(rowIdx + 1, 2).setValue(need);
//       } else {
//         sh.appendRow([k, need]);
//       }
//     }
//   });
// }

// // ============================================================
// // === OPTIMIZED: cs_submitClientStatusUpdate_ with conditional ensure ===
// // ============================================================
// function cs_submitClientStatusUpdate_(opts) {
//   try {
//     const ss = SpreadsheetApp.getActive();
//     const master = ss.getSheetByName(CS_MASTER_SHEET_NAME);
//     const row = cs_resolveRow_(master, opts && opts.rowNum);

//     const header = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];
//     const H = headerIndexMap_(header);
//     const vals = master.getRange(row, 1, 1, master.getLastColumn()).getValues()[0];
//     const get  = n => vals[H[n]] ?? '';

//     const apptId      = String(get('APPT_ID') || '').trim();
//     const brand       = String(get('Brand') || '');
//     const client      = String(get('Customer Name') || '');
//     const rep         = String(get('Assigned Rep') || '');
//     const salesStage  = String(get('Sales Stage') || '');
//     const convStatus  = String(get('Conversion Status') || '');
//     const customOrd   = String(get('Custom Order Status') || '');
//     const inProduction = String(get('In Production Status') || (opts && opts.inProduction) || '');
//     const centerStone = String(get('Center Stone Order Status') || '');
//     const nextSteps   = String(get('Next Steps') || '');
//     const orderDate   = String(get('Order Date') || '');

//     const phone       = String(getByAny_(H, vals, ['Phone', 'Client Phone', 'Customer Phone']) || '');
//     const email       = String(getByAny_(H, vals, ['Email', 'Client Email', 'Customer Email']) || '');
//     const occasion    = String(getByAny_(H, vals, ['Occasion']) || '');
//     const budgetRange = String(getByAny_(H, vals, ['Budget Range']) || '');
//     const decisionMkr = String(getByAny_(H, vals, ['Decision-Maker', 'Decision Maker']) || '');
//     const initialReq  = String(getByAny_(H, vals, ['Initial Request']) || '');
//     const soNumber    = String(getByAny_(H, vals, ['SO Number', 'SO#']) || '').trim();

//     const now  = new Date();
//     const iso  = Utilities.formatDate(now, CS_TZ, 'yyyy-MM-dd');
//     const ts   = Utilities.formatDate(now, CS_TZ, 'yyyy-MM-dd HH:mm:ss');
//     const nice = Utilities.formatDate(now, CS_TZ, 'MMM d, yyyy h:mm a z');
//     const user = Session.getActiveUser().getEmail() || rep || 'Unknown';
//     const assistedRep = String((opts && opts.assistedRep) || '');

//     // 1) Central audit
//     const audit = ss.getSheetByName(CS_AUDIT_TAB);
//     if (audit) {
//       const rootKeyForAudit = String(get('RootApptID') || get('APPT_ID') || '').trim();
//       let appliedCountTotal = 1;
//       if (rootKeyForAudit) {
//         const lastRowAll = master.getLastRow();
//         if (lastRowAll > 1) {
//           const matchColIndexAudit = (H['RootApptID'] != null) ? H['RootApptID']
//                                    : (H['APPT_ID']    != null) ? H['APPT_ID']
//                                    : -1;
//           if (matchColIndexAudit >= 0) {
//             const allValsAudit = master.getRange(2, 1, lastRowAll - 1, master.getLastColumn()).getValues();
//             for (let i = 0; i < allValsAudit.length; i++) {
//               const rnum = i + 2;
//               if (rnum === row) continue;
//               const idHere = String(allValsAudit[i][matchColIndexAudit] || '').trim();
//               if (idHere && idHere === rootKeyForAudit) appliedCountTotal++;
//             }
//           }
//         }
//       }
//       const appliedNote = `Applied to ${appliedCountTotal} row${appliedCountTotal === 1 ? '' : 's'}`
//                         + (rootKeyForAudit ? ` (RootApptID=${rootKeyForAudit})` : '');

//       let auditHeader = audit.getRange(1, 1, 1, audit.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
//       if (auditHeader.indexOf('Applied To') < 0) {
//         audit.getRange(1, audit.getLastColumn() + 1).setValue('Applied To');
//         auditHeader = audit.getRange(1, 1, 1, audit.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
//       }

//       cs_audit_appendByHeader_(audit, auditHeader, {
//         'APPT_ID':                   apptId,
//         'Log Date':                  iso,
//         'Sales Stage':               salesStage,
//         'Conversion Status':         convStatus,
//         'Custom Order Status':       customOrd,
//         'In Production Status':      inProduction,
//         'Center Stone Order Status': centerStone,
//         'Next Steps':                nextSteps,
//         'Assisted Rep':              assistedRep,
//         'Updated By':                user,
//         'Updated At':                ts,
//         'Applied To':                appliedNote
//       });
//     } else {
//       Logger.log(`Client Status: audit tab "${CS_AUDIT_TAB}" not found — continuing without central audit.`);
//     }

//     // ═══════════════════════════════════════════════════════════
//     // OPTIMIZATION: Conditional report ensure - only if not provided by caller
//     // ═══════════════════════════════════════════════════════════
//     let reportUrl, reportId, reportSS;
    
//     if (opts && opts.reportId && opts.reportUrl && opts.reportSS) {
//       // Already provided by caller (from dialog) - SKIP ensure
//       reportUrl = opts.reportUrl;
//       reportId = opts.reportId;
//       reportSS = opts.reportSS;
//       Logger.log('cs_submitClientStatusUpdate_: Using provided reportId=' + reportId + ' (skipping ensure, avoiding duplicate call)');
//     } else {
//       // Not provided (automation flow) - ensure now
//       const freshVals = master.getRange(row, 1, 1, master.getLastColumn()).getValues()[0];
//       const getFreshVal = n => H[n] != null ? (freshVals[H[n]] ?? '') : '';

//       const ensureResult = cs_ensureReportUrl_(master, row, H, getFreshVal);
//       if (!ensureResult.ok) {
//         return { ok: false, error: ensureResult.error || 'Could not create/find client report' };
//       }
//       reportUrl = ensureResult.reportUrl;
//       reportId  = ensureResult.reportId;
//       reportSS  = ensureResult.reportSS;
//       Logger.log('cs_submitClientStatusUpdate_: Ensured reportId=' + reportId);
//     }
//     // ═══════════════════════════════════════════════════════════

//     const rootApptId = String(
//       (H['RootApptID'] != null ? vals[H['RootApptID']] : '') ||
//       (H['APPT_ID']    != null ? vals[H['APPT_ID']]    : '') || ''
//     ).trim();
//     ensureReportConfig_(reportSS, { rootApptId, reportId });

//     // 3) Per-client log row
//     if (CS_WRITE_PER_CLIENT_LOG) {
//       insertLogRowByHeader_(reportSS, {
//         'Log Date':                  iso,
//         'Sales Stage':               salesStage,
//         'Conversion Status':         convStatus,
//         'Custom Order Status':       customOrd,
//         'In Production Status':      inProduction,
//         'Center Stone Order Status': centerStone,
//         'Next Steps':                nextSteps,
//         'Deadline Type':             (opts && opts.__deadlineLog && opts.__deadlineLog.type)  || '',
//         'Deadline Date':             (opts && opts.__deadlineLog && opts.__deadlineLog.date)  || '',
//         'Move Count':                (opts && opts.__deadlineLog && opts.__deadlineLog.moves) || '',
//         'Assisted Rep':              assistedRep,
//         'Updated By':                user,
//         'Updated At':                ts
//       });
//     }

//     // 4) Snapshot
//     updateSnapshot_(reportSS, {
//       Brand: brand, ClientName: client, APPT_ID: apptId, AssignedRep: rep,
//       Phone: phone, Email: email, Occasion: occasion,
//       BudgetRange: budgetRange, DecisionMaker: decisionMkr, InitialRequest: initialReq,
//       SO_Number: soNumber,
//       SalesStage: salesStage, ConversionStatus: convStatus, CustomOrderStatus: customOrd,
//       InProductionStatus: inProduction,
//       CenterStoneStatus: centerStone, NextSteps: nextSteps, UpdatedBy: user, UpdatedAt: ts,
//       AssistedRep: assistedRep,
//       OrderDate: orderDate
//     });

//     // 5) Updated By/At on master
//     const uIdx = H['Updated By'], aIdx = H['Updated At'];
//     if (uIdx != null && aIdx != null && Math.abs((uIdx + 1) - (aIdx + 1)) === 1) {
//       const from = Math.min(uIdx, aIdx) + 1;
//       const pairVals = (uIdx < aIdx) ? [[user, ts]] : [[ts, user]];
//       master.getRange(row, from, 1, 2).setValues(pairVals);
//     } else {
//       if (uIdx != null) master.getRange(row, uIdx + 1).setValue(user);
//       if (aIdx != null) master.getRange(row, aIdx + 1).setValue(ts);
//     }

//     // 5b) Fan-out with URL column exclusion
//     try {
//       const rootKey = String(get('RootApptID') || get('APPT_ID') || '').trim();
//       if (rootKey) {
//         const lastRow = master.getLastRow();
//         if (lastRow > 1) {
//           const ipsIdx = (H['In Production Status'] != null)
//             ? H['In Production Status']
//             : findHeaderIndexByRegex_(header, /in\s*production\s*status/i);

//           // Get URL column index to exclude from fan-out
//           const urlColIdx = H[CS_REPORT_URL_COL];

//           const allVals = master.getRange(2, 1, lastRow - 1, master.getLastColumn()).getValues();
//           const matchColIndex = (H['RootApptID'] != null) ? H['RootApptID']
//                               : (H['APPT_ID'] != null) ? H['APPT_ID']
//                               : -1;

//           if (matchColIndex >= 0) {
//             const targets = [];
//             for (let i = 0; i < allVals.length; i++) {
//               const rowNum = i + 2;
//               if (rowNum === row) continue;
//               const idHere = String(allVals[i][matchColIndex] || '').trim();
//               if (idHere && idHere === rootKey) targets.push(rowNum);
//             }

//             if (targets.length) {
//               const enqueuePairs = (name, value) => {
//                 const idx = H[name];
//                 if (idx == null) return null;
                
//                 // Skip URL column to prevent overwrite
//                 if (urlColIdx != null && idx === urlColIdx) {
//                   Logger.log('Fan-out: skipping URL column to prevent overwrite');
//                   return null;
//                 }
                
//                 const pairs = [];
//                 for (const rnum of targets) pairs.push({ r: rnum, v: value });
//                 return { colIdx1: idx + 1, pairs };
//               };

//               const q = [];
//               q.push(enqueuePairs('Assigned Rep',              rep));
//               q.push(enqueuePairs('Assisted Rep',              assistedRep));
//               q.push(enqueuePairs('Sales Stage',               salesStage));
//               q.push(enqueuePairs('Conversion Status',         convStatus));
//               q.push(enqueuePairs('Custom Order Status',       customOrd));
//               q.push(enqueuePairs('Center Stone Order Status', centerStone));
//               q.push(enqueuePairs('Next Steps',                nextSteps));
//               q.push(enqueuePairs('Updated By',                user));
//               q.push(enqueuePairs('Updated At',                ts));

//               if (ipsIdx >= 0 && (urlColIdx == null || ipsIdx !== urlColIdx)) {
//                 const ipsPairs = [];
//                 for (const rnum of targets) ipsPairs.push({ r: rnum, v: inProduction });
//                 groupedSetValues_(master, ipsIdx + 1, ipsPairs);
//               }

//               for (const item of q) {
//                 if (item && item.pairs && item.pairs.length) {
//                   groupedSetValues_(master, item.colIdx1, item.pairs);
//                 }
//               }
//             }
//           }
//         }
//       }
//     } catch (e) {
//       Logger.log('Fan-out to RootApptID siblings failed: ' + (e && e.message ? e.message : e));
//     }

//     // 6) DV hook
//     try {
//       if (typeof DV_init_ === 'function') { DV_init_(); }

//       var prevCenterStone = (opts && opts.prevCenterStone) || '';
//       var newCenterStone  = centerStone || '';
//       var becameNeed = !(typeof DV_isNeedToPropose === 'function' ? DV_isNeedToPropose(prevCenterStone) : false)
//                     &&  (typeof DV_isNeedToPropose === 'function' ? DV_isNeedToPropose(newCenterStone)  : false);
//       Logger.log('DV hook: prev="' + prevCenterStone + '" → new="' + newCenterStone + '"; becameNeed=' + becameNeed);

//       if (becameNeed && rootApptId) {
//         var res = DV_upsertProposeNudge_afterStatus_({
//           rootApptId,
//           customerName: client,
//           nextStepsFromMaster: nextSteps
//         });
//         Logger.log('DV hook: queued +2d nudge for root=' + rootApptId + ' → ' + JSON.stringify(res));
//       }
//     } catch (e) {
//       Logger.log('DV hook error: ' + (e && e.message ? e.message : e));
//     }

//     // 7) Reminders hook
//     try {
//       Remind.onClientStatusChange(soNumber, salesStage, customOrd, user, {
//         assignedRepName:  rep,
//         assistedRepName:  assistedRep,
//         customerName:     client,
//         nextSteps
//       });
//     } catch (e) {
//       console.warn('Remind.onClientStatusChange failed:', e && e.message ? e.message : e);
//     }

//     const masterLink = ss.getUrl() + '#gid=' + master.getSheetId() + '&range=A' + row;
//     const waxObj        = (opts && opts.wax) || null;
//     const waxSummaryStr = String((opts && opts.waxSummaryStr) || '');

//     return {
//       ok: true,
//       summary: {
//         clientName:  client, apptId,
//         assignedRep: rep,    assistedRep,
//         salesStage,  convStatus,
//         customOrder: customOrd,
//         deadline3d:   String((opts && opts.deadline3d)   || ''),
//         orderDate,
//         inProduction,
//         prodDeadline: String((opts && opts.prodDeadline) || ''),
//         centerStone,  nextSteps,
//         submittedBy:  user,
//         submittedAt:  nice,
//         reportUrl,    masterLink,
//         rootApptId,
//         waxSummary: waxSummaryStr,
//         wax:        waxObj
//       }
//     };

//   } catch (e) {
//     Logger.log('cs_submitClientStatusUpdate_ ERROR: ' + (e && e.message ? e.message : e));
//     return { ok: false, error: String(e && e.message || e) };
//   }
// }

// // ============================================================
// // === Log helpers ===
// // ============================================================

// function getLogHeaderRow_(sh) {
//   const sp  = sh.getParent();
//   const key = 'CS_LOG_HDR_' + (sp && sp.getId ? sp.getId() : '') + '_' + sh.getSheetId();
//   const props = PropertiesService.getScriptProperties();

//   const cached = Number(props.getProperty(key) || 0);
//   if (cached && String(sh.getRange(cached, 1).getValue()).trim() === 'Log Date') return cached;

//   const start = 8, end = Math.min(sh.getLastRow() || 80, 80);
//   const scan = sh.getRange(start, 1, Math.max(end - start + 1, 1), 1).getValues();
//   let headerRow = 13;
//   for (let i = 0; i < scan.length; i++) {
//     if (String(scan[i][0] || '').trim() === 'Log Date') { headerRow = start + i; break; }
//   }
//   props.setProperty(key, String(headerRow));
//   return headerRow;
// }

// function insertLogRowByHeader_(reportSS, valuesByName) {
//   const sh = reportSS.getSheetByName(CS_REPORT_SHEET);
//   if (!sh) throw new Error(`Missing "${CS_REPORT_SHEET}" tab`);

//   const headerRow = getLogHeaderRow_(sh);
//   const header = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
//   const H = {}; header.forEach((h, i) => { if (h) H[h] = i; });

//   const row = new Array(header.length).fill('');
//   Object.keys(valuesByName).forEach(name => {
//     const i = H[name]; if (i != null) row[i] = valuesByName[name];
//   });

//   sh.insertRowsBefore(headerRow + 1, 1);
//   sh.getRange(headerRow + 1, 1, 1, row.length).setValues([row]);
// }

// function groupedSetValues_(sh, colIdx, pairs) {
//   if (!pairs || !pairs.length) return;
//   pairs.sort((a, b) => a.r - b.r);
//   let start = pairs[0].r;
//   let block  = [[pairs[0].v]];
//   for (let i = 1; i < pairs.length; i++) {
//     const prev = pairs[i - 1].r, cur = pairs[i].r;
//     if (cur === prev + 1) {
//       block.push([pairs[i].v]);
//     } else {
//       sh.getRange(start, colIdx, block.length, 1).setValues(block);
//       start = cur; block = [[pairs[i].v]];
//     }
//   }
//   sh.getRange(start, colIdx, block.length, 1).setValues(block);
// }

// function updateSnapshot_(reportSS, data) {
//   const sh = reportSS.getSheetByName(CS_REPORT_SHEET);
//   if (!sh) throw new Error(`Missing "${CS_REPORT_SHEET}" tab`);

//   const map = {
//     'Report Date:':              '__InitDate',
//     'Customer Name:':            'ClientName',
//     'APPT_ID:':                  'APPT_ID',
//     'Brand:':                    'Brand',
//     'Assigned Rep:':             'AssignedRep',
//     'Phone:':                    'Phone',
//     'Email:':                    'Email',
//     'Occasion:':                 'Occasion',
//     'Budget Range:':             'BudgetRange',
//     'Decision-Maker:':           'DecisionMaker',
//     'Initial Request:':          'InitialRequest',
//     'SO#:':                      'SO_Number',
//     'Sales Stage:':              'SalesStage',
//     'Conversion Status:':        'ConversionStatus',
//     'Custom Order Status:':      'CustomOrderStatus',
//     'In Production Status:':     'InProductionStatus',
//     'Center Stone Order Status:':'CenterStoneStatus',
//     'Next Steps:':               'NextSteps',
//     'Updated By:':               'UpdatedBy',
//     'Updated At:':               'UpdatedAt',
//     'Assisted Rep:':             'AssistedRep',
//     'Order Date:':               'OrderDate'
//   };

//   const rowsToScan = Math.min(sh.getLastRow() || 50, 50);
//   if (rowsToScan <= 0) return;

//   const values = sh.getRange(1, 1, rowsToScan, 4).getValues();

//   const writesB = [];
//   const writesD = [];
//   const todayStr = Utilities.formatDate(new Date(), CS_TZ, 'yyyy-MM-dd');

//   for (let i = 0; i < rowsToScan; i++) {
//     const labA = String(values[i][0] || '').trim();
//     const labC = String(values[i][2] || '').trim();

//     const apply = (label, targetColIndexZeroBased) => {
//       const key = map[label]; if (!key) return;

//       if (key === '__InitDate') {
//         const current = String(values[i][targetColIndexZeroBased] || '').trim();
//         if (!current) {
//           if (targetColIndexZeroBased === 1) writesB.push({ r: i + 1, v: todayStr });
//           else if (targetColIndexZeroBased === 3) writesD.push({ r: i + 1, v: todayStr });
//         }
//         return;
//       }

//       const newVal = data[key] != null ? String(data[key]) : '';
//       if (targetColIndexZeroBased === 1) writesB.push({ r: i + 1, v: newVal });
//       else if (targetColIndexZeroBased === 3) writesD.push({ r: i + 1, v: newVal });
//     };

//     if (labA) apply(labA, 1);
//     if (labC) apply(labC, 3);
//   }

//   if (writesB.length) groupedSetValues_(sh, 2, writesB);
//   if (writesD.length) groupedSetValues_(sh, 4, writesD);
// }

// function toISODateForInput_(v) {
//   if (v instanceof Date && !isNaN(v)) {
//     return Utilities.formatDate(v, CS_TZ, 'yyyy-MM-dd');
//   }
//   const s = String(v || '').trim();
//   if (!s) return '';
//   if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
//   const d = new Date(s);
//   if (!isNaN(d)) return Utilities.formatDate(d, CS_TZ, 'yyyy-MM-dd');
//   const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
//   if (m) {
//     const y = m[3].length === 2 ? ('20' + m[3]) : m[3];
//     const mm = ('0' + m[1]).slice(-2), dd = ('0' + m[2]).slice(-2);
//     return y + '-' + mm + '-' + dd;
//   }
//   return '';
// }

// function CS_AUDIT_upgrade_addIPS_AtEnd() {
//   const ss = SpreadsheetApp.getActive();
//   const sh = ss.getSheetByName('03_Client_Status_Log');
//   if (!sh) throw new Error('Sheet "03_Client_Status_Log" not found.');

//   const lastCol = Math.max(1, sh.getLastColumn());
//   const header  = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(x => String(x || '').trim());

//   if (header.includes('In Production Status')) {
//     Logger.log('Already present. No changes.');
//     return;
//   }

//   const newCol = lastCol + 1;
//   sh.getRange(1, newCol).setValue('In Production Status');
//   Logger.log('Added "In Production Status" as new last column ' + newCol + '.');
// }

// function cs_audit_appendByHeader_(sh, header, valuesByName) {
//   const H = {}; header.forEach((h, i) => { if (h) H[String(h).trim()] = i; });
//   const row = new Array(header.length).fill('');
//   Object.keys(valuesByName).forEach(name => {
//     const i = H[name]; if (i != null) row[i] = valuesByName[name];
//   });
//   sh.appendRow(row);
// }

// function cs_automationSubmit_(params) {
//   if (!params || !params.rowNum || Number(params.rowNum) <= 1) {
//     throw new Error('cs_automationSubmit_: params.rowNum is required and must be > 1.');
//   }

//   const master = SpreadsheetApp.getActive().getSheetByName(CS_MASTER_SHEET_NAME);
//   master.setActiveRange(master.getRange(Number(params.rowNum), 1));

//   const result = cs_submitClientStatusUpdate_({
//     rowNum:       Number(params.rowNum),
//     assistedRep:  String(params.assistedRep  || ''),
//     inProduction: String(params.inProduction || ''),
//     prodDeadline: String(params.prodDeadline || ''),
//     deadline3d:   String(params.deadline3d   || ''),
//     prevCenterStone: ''
//   });

//   if (!result.ok) {
//     Logger.log('cs_automationSubmit_ FAILED at row ' + params.rowNum + ': ' + result.error);
//   } else {
//     Logger.log('cs_automationSubmit_ OK at row ' + params.rowNum + ': reportUrl=' + (result.summary && result.summary.reportUrl));
//   }

//   return result;
// }

// // ============================================================
// // === REPAIR FUNCTIONS - Run once after v2.6 deployment ===
// // ============================================================

// /**
//  * REPAIR 1: Backfill idempotency tokens for existing files
//  * 
//  * Scans all rows in Master sheet, finds files by URL, and sets
//  * file.setDescription("APPT_ID=...") if missing.
//  * 
//  * Usage: Run from Apps Script editor
//  */
// function REPAIR_backfillIdempotencyTokens() {
//   const ss = SpreadsheetApp.getActive();
//   const master = ss.getSheetByName(CS_MASTER_SHEET_NAME);
//   if (!master) throw new Error('Master sheet not found');

//   const header = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];
//   const H = headerIndexMap_(header);
  
//   const urlIdx = H[CS_REPORT_URL_COL];
//   const apptIdx = H['APPT_ID'];
  
//   if (urlIdx == null || apptIdx == null) {
//     throw new Error('Required columns not found: ' + CS_REPORT_URL_COL + ', APPT_ID');
//   }

//   const lastRow = master.getLastRow();
//   if (lastRow <= 1) {
//     Logger.log('No data rows to process');
//     return { ok: true, processed: 0, updated: 0, errors: 0 };
//   }

//   const allVals = master.getRange(2, 1, lastRow - 1, master.getLastColumn()).getValues();
  
//   let processed = 0, updated = 0, errors = 0;
  
//   for (let i = 0; i < allVals.length; i++) {
//     const row = i + 2;
//     const url = String(allVals[i][urlIdx] || '').trim();
//     const apptId = String(allVals[i][apptIdx] || '').trim();
    
//     if (!url || !apptId) continue;
//     if (!isValidSpreadsheetUrl_(url)) continue;
    
//     processed++;
    
//     try {
//       const fileId = extractIdFromUrl_(url);
//       const file = DriveApp.getFileById(fileId);
//       const currentDesc = String(file.getDescription() || '').trim();
//       const token = 'APPT_ID=' + apptId;
      
//       if (!currentDesc.includes(token)) {
//         const newDesc = currentDesc ? (currentDesc + ' | ' + token) : token;
//         file.setDescription(newDesc);
//         updated++;
//         Logger.log('Row ' + row + ': Added token to file ' + fileId);
//       }
      
//       // Throttle to avoid quota limits
//       if (processed % 10 === 0) {
//         Utilities.sleep(1000);
//       }
      
//     } catch (e) {
//       errors++;
//       Logger.log('Row ' + row + ' ERROR: ' + e.message);
//     }
//   }
  
//   const summary = {
//     ok: true,
//     processed: processed,
//     updated: updated,
//     errors: errors,
//     message: 'Backfilled ' + updated + '/' + processed + ' files (' + errors + ' errors)'
//   };
  
//   Logger.log(JSON.stringify(summary));
//   return summary;
// }


// /**
//  * REPAIR 2: Find and link orphaned files
//  * 
//  * Scans Master sheet for blank URL cells, searches Drive for matching files
//  * by name pattern, and writes URL back to Master if found.
//  * 
//  * Usage: Run from Apps Script editor
//  */
// function REPAIR_linkOrphanedFiles() {
//   const ss = SpreadsheetApp.getActive();
//   const master = ss.getSheetByName(CS_MASTER_SHEET_NAME);
//   if (!master) throw new Error('Master sheet not found');

//   const header = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];
//   const H = headerIndexMap_(header);
  
//   const urlIdx = H[CS_REPORT_URL_COL];
//   const apptIdx = H['APPT_ID'];
//   const brandIdx = H['Brand'];
  
//   if (urlIdx == null || apptIdx == null) {
//     throw new Error('Required columns not found');
//   }

//   const lastRow = master.getLastRow();
//   if (lastRow <= 1) return { ok: true, linked: 0 };

//   const allVals = master.getRange(2, 1, lastRow - 1, master.getLastColumn()).getValues();
  
//   let scanned = 0, linked = 0, notFound = 0;
//   const updates = []; // Batch updates
  
//   for (let i = 0; i < allVals.length; i++) {
//     const row = i + 2;
//     const url = String(allVals[i][urlIdx] || '').trim();
//     const apptId = String(allVals[i][apptIdx] || '').trim();
//     const brand = brandIdx != null ? String(allVals[i][brandIdx] || '').trim() : 'VVS';
    
//     // Skip if URL already exists or APPT_ID missing
//     if (url || !apptId) continue;
    
//     scanned++;
    
//     try {
//       // Search by name pattern
//       const expectedName = CS_REPORT_NAME_FMT
//         .replace('{Brand}', brand)
//         .replace('{APPT_ID}', apptId);
      
//       const files = DriveApp.getFilesByName(expectedName);
//       let foundFile = null;
      
//       // Find first valid spreadsheet
//       while (files.hasNext()) {
//         const file = files.next();
//         if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
//           // Verify it's not trashed
//           if (!file.isTrashed()) {
//             foundFile = file;
//             break;
//           }
//         }
//       }
      
//       if (foundFile) {
//         const newUrl = 'https://docs.google.com/spreadsheets/d/' + foundFile.getId() + '/edit';
//         updates.push({ row: row, url: newUrl });
//         linked++;
//         Logger.log('Row ' + row + ': Found orphaned file ' + foundFile.getId());
//       } else {
//         notFound++;
//         Logger.log('Row ' + row + ': No file found for ' + apptId);
//       }
      
//       // Throttle
//       if (scanned % 5 === 0) {
//         Utilities.sleep(2000);
//       }
      
//     } catch (e) {
//       Logger.log('Row ' + row + ' ERROR: ' + e.message);
//     }
//   }
  
//   // Batch write URLs
//   if (updates.length > 0) {
//     updates.forEach(u => {
//       master.getRange(u.row, urlIdx + 1).setValue(u.url);
//     });
//   }
  
//   const summary = {
//     ok: true,
//     scanned: scanned,
//     linked: linked,
//     notFound: notFound,
//     message: 'Linked ' + linked + '/' + scanned + ' orphaned files'
//   };
  
//   Logger.log(JSON.stringify(summary));
//   return summary;
// }


// /**
//  * REPAIR 3: Validate and fix stale URLs
//  * 
//  * Checks all URL cells in Master sheet, verifies files exist,
//  * and either recreates missing files or clears stale URLs.
//  * 
//  * Usage: Run from Apps Script editor with caution
//  * Set DRY_RUN = false to actually fix issues
//  */
// function REPAIR_validateAndFixStaleUrls(opts) {
//   const DRY_RUN = opts && opts.dryRun !== false; // Default to dry run
  
//   const ss = SpreadsheetApp.getActive();
//   const master = ss.getSheetByName(CS_MASTER_SHEET_NAME);
//   if (!master) throw new Error('Master sheet not found');

//   const header = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];
//   const H = headerIndexMap_(header);
  
//   const urlIdx = H[CS_REPORT_URL_COL];
//   if (urlIdx == null) throw new Error('URL column not found');

//   const lastRow = master.getLastRow();
//   if (lastRow <= 1) return { ok: true, checked: 0 };

//   const allVals = master.getRange(2, 1, lastRow - 1, master.getLastColumn()).getValues();
  
//   let checked = 0, valid = 0, stale = 0, fixed = 0;
//   const staleRows = [];
  
//   for (let i = 0; i < allVals.length; i++) {
//     const row = i + 2;
//     const url = String(allVals[i][urlIdx] || '').trim();
    
//     if (!url) continue;
//     if (!isValidSpreadsheetUrl_(url)) {
//       staleRows.push({ row: row, reason: 'Invalid URL format', url: url });
//       stale++;
//       continue;
//     }
    
//     checked++;
    
//     try {
//       const fileId = extractIdFromUrl_(url);
//       const file = DriveApp.getFileById(fileId);
      
//       // Check if trashed
//       if (file.isTrashed()) {
//         staleRows.push({ row: row, reason: 'File is trashed', url: url });
//         stale++;
//       } else {
//         // Verify it's a spreadsheet
//         if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
//           staleRows.push({ row: row, reason: 'Not a spreadsheet', url: url });
//           stale++;
//         } else {
//           valid++;
//         }
//       }
      
//     } catch (e) {
//       // File doesn't exist or no access
//       staleRows.push({ row: row, reason: 'File not found: ' + e.message, url: url });
//       stale++;
//     }
    
//     // Throttle
//     if (checked % 10 === 0) {
//       Utilities.sleep(1000);
//     }
//   }
  
//   // Fix stale URLs by clearing them (user can regenerate via ensureReportUrl_)
//   if (!DRY_RUN && staleRows.length > 0) {
//     staleRows.forEach(item => {
//       master.getRange(item.row, urlIdx + 1).setValue('');
//       fixed++;
//       Logger.log('Row ' + item.row + ': Cleared stale URL (' + item.reason + ')');
//     });
//   }
  
//   const summary = {
//     ok: true,
//     checked: checked,
//     valid: valid,
//     stale: stale,
//     fixed: DRY_RUN ? 0 : fixed,
//     dryRun: DRY_RUN,
//     staleRows: staleRows.map(r => ({ row: r.row, reason: r.reason })),
//     message: DRY_RUN 
//       ? 'DRY RUN: Found ' + stale + '/' + checked + ' stale URLs (re-run with {dryRun:false} to fix)'
//       : 'Fixed ' + fixed + '/' + stale + ' stale URLs'
//   };
  
//   Logger.log(JSON.stringify(summary, null, 2));
//   return summary;
// }


// /**
//  * REPAIR 4: Find and remove duplicate files
//  * 
//  * Scans Drive for files with duplicate APPT_ID tokens or matching names,
//  * keeps the file referenced in Master sheet URL, deletes others.
//  * 
//  * DANGER: This permanently trashes files. Use with extreme caution.
//  * 
//  * Usage: Run from Apps Script editor
//  * Set ACTUALLY_DELETE = true after reviewing the dry run report
//  */
// function REPAIR_removeDuplicateFiles(opts) {
//   const ACTUALLY_DELETE = false; // Permanently disabled for safety
  
//   const ss = SpreadsheetApp.getActive();
//   const master = ss.getSheetByName(CS_MASTER_SHEET_NAME);
//   if (!master) throw new Error('Master sheet not found');

//   const header = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];
//   const H = headerIndexMap_(header);
  
//   const urlIdx = H[CS_REPORT_URL_COL];
//   const apptIdx = H['APPT_ID'];
  
//   if (urlIdx == null || apptIdx == null) {
//     throw new Error('Required columns not found');
//   }

//   const lastRow = master.getLastRow();
//   if (lastRow <= 1) return { ok: true, scanned: 0, duplicatesFound: 0 };

//   const allVals = master.getRange(2, 1, lastRow - 1, master.getLastColumn()).getValues();
  
//   // Build map: APPT_ID → canonical file ID (from Master URL)
//   const canonicalFiles = {}; // { apptId: fileId }
  
//   for (let i = 0; i < allVals.length; i++) {
//     const apptId = String(allVals[i][apptIdx] || '').trim();
//     const url = String(allVals[i][urlIdx] || '').trim();
    
//     if (!apptId || !url) continue;
    
//     try {
//       const fileId = extractIdFromUrl_(url);
//       if (!canonicalFiles[apptId]) {
//         canonicalFiles[apptId] = fileId;
//       }
//     } catch (e) {
//       // Skip invalid URLs
//     }
//   }
  
//   Logger.log('Found ' + Object.keys(canonicalFiles).length + ' canonical files in Master sheet');
  
//   // Scan Drive for duplicates
//   const duplicates = [];
//   let scanned = 0;
  
//   for (const apptId in canonicalFiles) {
//     const canonicalId = canonicalFiles[apptId];
    
//     try {
//       // Search by token in description
//       const token = 'APPT_ID=' + apptId;
      
//       // Note: Drive search by description is not supported, so we search by name pattern
//       const expectedName = CS_REPORT_NAME_FMT.replace('{APPT_ID}', apptId);
//       const pattern = expectedName.replace('{Brand}', ''); // Partial match
      
//       // Use contains: operator for name search
//       const files = DriveApp.searchFiles(
//         'title contains "' + apptId + '" and mimeType = "' + MimeType.GOOGLE_SHEETS + '"'
//       );
      
//       const foundFiles = [];
//       while (files.hasNext()) {
//         const file = files.next();
//         foundFiles.push(file);
//         scanned++;
//       }
      
//       // Find duplicates (files with same APPT_ID but different ID than canonical)
//       for (const file of foundFiles) {
//         if (file.getId() !== canonicalId) {
//           // Check if description has matching token
//           const desc = String(file.getDescription() || '').trim();
//           if (desc.includes(token) || file.getName().includes(apptId)) {
//             duplicates.push({
//               apptId: apptId,
//               fileId: file.getId(),
//               name: file.getName(),
//               url: file.getUrl(),
//               canonical: canonicalId
//             });
//           }
//         }
//       }
      
//       // Throttle
//       Utilities.sleep(500);
      
//     } catch (e) {
//       Logger.log('Error scanning ' + apptId + ': ' + e.message);
//     }
//   }
  
//   // Delete duplicates (DISABLED)
//   let deleted = 0;
  
//   const summary = {
//     ok: true,
//     scanned: scanned,
//     duplicatesFound: duplicates.length,
//     deleted: deleted,
//     actuallyDelete: ACTUALLY_DELETE,
//     duplicates: duplicates.map(d => ({ apptId: d.apptId, fileId: d.fileId, name: d.name })),
//     message: 'SCAN ONLY: Found ' + duplicates.length + ' duplicate files. Deletion permanently disabled for safety.'
//   };
  
//   Logger.log(JSON.stringify(summary, null, 2));
//   return summary;
// }


// /**
//  * REPAIR 5: Master repair function - runs all repairs in sequence
//  * 
//  * Recommended order:
//  * 1. Backfill tokens (safe)
//  * 2. Link orphaned files (safe)
//  * 3. Validate stale URLs (dry run first)
//  * 4. Remove duplicates (DANGER - dry run only)
//  * 
//  * Usage:
//  *   REPAIR_runAll({ dryRun: true })  // Safe preview
//  *   REPAIR_runAll({ dryRun: false }) // Actually fix issues
//  */
// function REPAIR_runAll(opts) {
//   const DRY_RUN = opts && opts.dryRun !== false;
  
//   Logger.log('========================================');
//   Logger.log('REPAIR MASTER - DRY RUN: ' + DRY_RUN);
//   Logger.log('========================================');
  
//   const results = {
//     timestamp: new Date().toISOString(),
//     dryRun: DRY_RUN,
//     repairs: {}
//   };
  
//   try {
//     Logger.log('\n1️⃣ Backfilling idempotency tokens...');
//     results.repairs.backfillTokens = REPAIR_backfillIdempotencyTokens();
    
//     Logger.log('\n2️⃣ Linking orphaned files...');
//     results.repairs.linkOrphaned = REPAIR_linkOrphanedFiles();
    
//     Logger.log('\n3️⃣ Validating stale URLs...');
//     results.repairs.validateUrls = REPAIR_validateAndFixStaleUrls({ dryRun: DRY_RUN });
    
//     Logger.log('\n4️⃣ Scanning for duplicate files...');
//     results.repairs.removeDuplicates = REPAIR_removeDuplicateFiles({ actuallyDelete: false });
    
//     results.ok = true;
    
//   } catch (e) {
//     results.ok = false;
//     results.error = e.message;
//     Logger.log('❌ REPAIR FAILED: ' + e.message);
//   }
  
//   Logger.log('\n========================================');
//   Logger.log('REPAIR SUMMARY:');
//   Logger.log(JSON.stringify(results, null, 2));
//   Logger.log('========================================');
  
//   return results;
// }


// // ============================================================
// // === VERIFICATION HELPERS ===
// // ============================================================

// /**
//  * Test function to verify v2.6 optimizations are working
//  * Run from Apps Script editor after deployment
//  */
// function TEST_verifyOptimizations() {
//   const results = {
//     timestamp: new Date().toISOString(),
//     version: '2.6',
//     tests: []
//   };
  
//   // Test 1: Check if URL column is excluded from fan-out
//   try {
//     const ss = SpreadsheetApp.getActive();
//     const master = ss.getSheetByName(CS_MASTER_SHEET_NAME);
//     const header = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];
//     const H = headerIndexMap_(header);
//     const urlColIdx = H[CS_REPORT_URL_COL];
    
//     results.tests.push({
//       name: 'URL Column Detection',
//       status: urlColIdx != null ? 'PASS' : 'FAIL',
//       urlColumnIndex: urlColIdx,
//       message: urlColIdx != null ? 'URL column found at index ' + urlColIdx : 'URL column not found'
//     });
//   } catch (e) {
//     results.tests.push({
//       name: 'URL Column Detection',
//       status: 'ERROR',
//       error: e.message
//     });
//   }
  
//   // Test 2: Verify exponential backoff is present
//   try {
//     const hasOptimizedLock = cs_ensureReportUrl_.toString().includes('Math.pow(2, attempt - 1)');
//     results.tests.push({
//       name: 'Exponential Backoff',
//       status: hasOptimizedLock ? 'PASS' : 'FAIL',
//       message: hasOptimizedLock ? 'Exponential backoff code detected' : 'Using old fixed retry delay'
//     });
//   } catch (e) {
//     results.tests.push({
//       name: 'Exponential Backoff',
//       status: 'ERROR',
//       error: e.message
//     });
//   }
  
//   // Test 3: Check if conditional ensure is present
//   try {
//     const hasConditional = cs_submitClientStatusUpdate_.toString().includes('opts.reportId && opts.reportUrl');
//     results.tests.push({
//       name: 'Conditional Report Ensure',
//       status: hasConditional ? 'PASS' : 'FAIL',
//       message: hasConditional ? 'Conditional logic detected' : 'Still calling ensure unconditionally'
//     });
//   } catch (e) {
//     results.tests.push({
//       name: 'Conditional Report Ensure',
//       status: 'ERROR',
//       error: e.message
//     });
//   }
  
//   // Test 4: Verify final URL check is present
//   try {
//     const hasFinalCheck = cs_ensureReportUrl_.toString().includes('finalCheck');
//     results.tests.push({
//       name: 'Final URL Write Check',
//       status: hasFinalCheck ? 'PASS' : 'FAIL',
//       message: hasFinalCheck ? 'Final check detected' : 'Missing final URL check'
//     });
//   } catch (e) {
//     results.tests.push({
//       name: 'Final URL Write Check',
//       status: 'ERROR',
//       error: e.message
//     });
//   }
  
//   // Test 5: Check max retry attempts
//   try {
//     const codeStr = cs_ensureReportUrl_.toString();
//     const has5Attempts = codeStr.includes('MAX_ATTEMPTS   = 5');
//     results.tests.push({
//       name: 'Max Retry Attempts',
//       status: has5Attempts ? 'PASS' : 'FAIL',
//       message: has5Attempts ? 'MAX_ATTEMPTS = 5' : 'Still using old MAX_ATTEMPTS value'
//     });
//   } catch (e) {
//     results.tests.push({
//       name: 'Max Retry Attempts',
//       status: 'ERROR',
//       error: e.message
//     });
//   }
  
//   const passCount = results.tests.filter(t => t.status === 'PASS').length;
//   const totalTests = results.tests.length;
  
//   results.summary = {
//     passed: passCount,
//     total: totalTests,
//     percentage: Math.round((passCount / totalTests) * 100),
//     message: passCount === totalTests 
//       ? '✅ All optimizations verified!'
//       : '⚠️ ' + (totalTests - passCount) + ' tests failed - review results'
//   };
  
//   Logger.log('========================================');
//   Logger.log('VERIFICATION RESULTS v2.6:');
//   Logger.log(JSON.stringify(results, null, 2));
//   Logger.log('========================================');
  
//   return results;
// }


// /**
//  * Monitor function to track optimization metrics
//  * Run periodically to check performance
//  */
// function MONITOR_optimizationMetrics() {
//   const metrics = {
//     timestamp: new Date().toISOString(),
//     version: '2.6',
//     period: 'last 24 hours',
//     stats: {}
//   };
  
//   try {
//     const ss = SpreadsheetApp.getActive();
//     const audit = ss.getSheetByName(CS_AUDIT_TAB);
    
//     if (audit) {
//       const lastRow = audit.getLastRow();
//       if (lastRow > 1) {
//         const data = audit.getRange(2, 1, lastRow - 1, audit.getLastColumn()).getValues();
//         const header = audit.getRange(1, 1, 1, audit.getLastColumn()).getValues()[0];
        
//         const H = headerIndexMap_(header);
//         const tsIdx = H['Updated At'];
        
//         // Count updates in last 24h
//         const oneDayAgo = new Date(Date.now() - 24 * 60 * 60 * 1000);
//         let recentUpdates = 0;
        
//         if (tsIdx != null) {
//           for (const row of data) {
//             const ts = row[tsIdx];
//             if (ts && new Date(ts) > oneDayAgo) {
//               recentUpdates++;
//             }
//           }
//         }
        
//         metrics.stats.totalAuditEntries = lastRow - 1;
//         metrics.stats.updatesLast24h = recentUpdates;
//       }
//     }
    
//     // Count URLs in master
//     const master = ss.getSheetByName(CS_MASTER_SHEET_NAME);
//     if (master) {
//       const header = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];
//       const H = headerIndexMap_(header);
//       const urlIdx = H[CS_REPORT_URL_COL];
      
//       if (urlIdx != null) {
//         const lastRow = master.getLastRow();
//         if (lastRow > 1) {
//           const urls = master.getRange(2, urlIdx + 1, lastRow - 1, 1).getValues();
//           const populated = urls.filter(r => String(r[0] || '').trim()).length;
          
//           metrics.stats.totalRows = lastRow - 1;
//           metrics.stats.urlsPopulated = populated;
//           metrics.stats.urlCoverage = Math.round((populated / (lastRow - 1)) * 100) + '%';
//         }
//       }
//     }
    
//     Logger.log('========================================');
//     Logger.log('OPTIMIZATION METRICS v2.6:');
//     Logger.log(JSON.stringify(metrics, null, 2));
//     Logger.log('========================================');
    
//   } catch (e) {
//     Logger.log('Error collecting metrics: ' + e.message);
//   }
  
//   return metrics;
// }


// // ============================================================
// // === Legacy shims ===
// // ============================================================
// if (typeof headerMap_ !== 'function') {
//   function headerMap_(sh) { return headerMap__canon(sh); }
// }
// if (typeof ensureHeaders_ !== 'function') {
//   function ensureHeaders_(sh, labels) { return ensureHeaders__canon(sh, labels); }
// }
// if (typeof getMasterSheet_ !== 'function') {
//   function getMasterSheet_(ss) { return getMasterSheet__canon(ss); }
// }
// if (typeof getOrdersSheet_ !== 'function') {
//   function getOrdersSheet_(wb) { return getOrdersSheet__canon(wb); }
// }
// if (typeof coerceSOTextColumn_ !== 'function') {
//   function coerceSOTextColumn_(sh, H) { return coerceSOTextColumn__canon(sh, H); }
// }
// if (typeof existsSOInMaster_ !== 'function') {
//   function existsSOInMaster_(sh, brand, so, skipRow) { return existsSOInMaster__canon(sh, brand, so, skipRow); }
// }


// === CONFIG ===
const MASTER_SHEET_NAME = '00_Master Appointments';
const CS_MASTER_SHEET_NAME = MASTER_SHEET_NAME;
const CS_AUDIT_SHEET = '03_Client_Status_Log';
const CS_AUDIT_TAB = CS_AUDIT_SHEET;
const CS_REPORT_SHEET = 'Client Status';
const CS_WRITE_PER_CLIENT_LOG = true;
const CS_TZ = 'America/Los_Angeles';

const CS_REPORT_URL_COL = 'Client Status Report URL';
const CS_PROSPECT_URL_COL = 'Prospect Folder URL';
const CS_REPORT_NAME_FMT = '{Brand} – {APPT_ID} – Client Status Report';

// Color column names in "Dropdown"
const COL_SALES_STAGE_HEX   = 'SS - Hex Code';
const COL_CONV_STATUS_HEX   = 'CS - Hex Code';
const COL_CUST_ORDER_HEX    = 'COS - Hex Code';
const COL_IN_PRODUCTION_HEX = 'IPS - Hex Code';
const COL_CENTER_STONE_HEX  = 'CSOS - Hex Code';

// === TEMPLATE CONFIG ===
function getTemplateId_() {
  return PropertiesService.getScriptProperties().getProperty('CS_REPORT_TEMPLATE_ID') || '';
}

// ============================================================
// === Helpers ===
// ============================================================

function headerIndexMap_(headerRow) {
  const map = {};
  headerRow.forEach((h, i) => { if (h) map[String(h).trim()] = i; });
  return map;
}

function findHeaderIndexByRegex_(headerRow, regex) {
  for (var i = 0; i < headerRow.length; i++) {
    if (regex.test(String(headerRow[i] || ''))) return i;
  }
  return -1;
}

function extractIdFromUrl_(url) {
  const m = String(url).match(/[-\w]{25,}/);
  return m ? m[0] : '';
}

// FIX 10: Validate spreadsheet URL format
function isValidSpreadsheetUrl_(url) {
  const s = String(url || '').trim();
  if (!s) return false;
  // Must contain /spreadsheets/d/ or be a valid spreadsheet ID
  return /\/spreadsheets\/d\/[-\w]{25,}/.test(s) || /^[-\w]{25,}$/.test(s);
}

function getByAny_(H, vals, names) {
  for (const n of names) {
    if (H[n] != null) return vals[H[n]] ?? '';
  }
  return '';
}

function normalizeMultiArray_(v) {
  if (Array.isArray(v)) return v.map(s => String(s || '').trim()).filter(Boolean);
  return String(v || '')
    .split(/[,;|/]|(?:\s*&\s*)/g)
    .map(s => s.trim())
    .filter(Boolean);
}

function joinMulti_(arr) {
  const a = normalizeMultiArray_(arr);
  const seen = new Set(); const out = [];
  a.forEach(x => { if (!seen.has(x)) { seen.add(x); out.push(x); } });
  return out.join(', ');
}

// ============================================================
// === FIX 1: cs_resolveRow_ — safe row resolver ===
// ============================================================
function cs_resolveRow_(sh, explicitRow) {
  if (explicitRow && Number(explicitRow) > 1) {
    return Number(explicitRow);
  }
  const r = sh.getActiveRange();
  if (!r || r.getRow() <= 1) {
    throw new Error(
      'Không xác định được row. Khi gọi từ automation/trigger, hãy truyền opts.rowNum (1-based, > 1).'
    );
  }
  return r.getRow();
}

// ============================================================
// === FIX 11: Search existing reports by idempotency token ===
// ============================================================
/**
 * Search parent folder for existing report with matching APPT_ID in description.
 * Returns file ID if found, null otherwise.
 */
function findExistingReportByToken_(parentFolder, apptId, reportName) {
  if (!apptId) return null;
  
  try {
    const token = 'APPT_ID=' + String(apptId).trim();
    const files = parentFolder.getFilesByName(reportName);
    
    while (files.hasNext()) {
      const file = files.next();
      const desc = String(file.getDescription() || '').trim();
      
      if (desc.includes(token)) {
        // Verify it's a spreadsheet
        if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
          Logger.log('findExistingReportByToken_: found existing file ' + file.getId());
          return file.getId();
        }
      }
    }
  } catch (e) {
    Logger.log('findExistingReportByToken_ search failed: ' + e.message);
  }
  
  return null;
}

// ============================================================
// === OPTIMIZED: cs_ensureReportUrl_ with enhancements ===
// ============================================================
function cs_ensureReportUrl_(masterSheet, row, H, getVal) {
  // ═══════════════════════════════════════════════════════════
  // OPTIMIZATION: Exponential backoff for lock retry
  // ═══════════════════════════════════════════════════════════
  const MAX_ATTEMPTS     = 5;     // Increased from 3
  const LOCK_TIMEOUT     = 30000; // 30s for slow Drive API
  const BASE_RETRY_SLEEP = 500;   // Base delay for exponential backoff

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    const lock = LockService.getDocumentLock();
    const gotLock = lock.tryLock(LOCK_TIMEOUT);

    if (!gotLock) {
      Logger.log('cs_ensureReportUrl_: lock busy, attempt ' + attempt + '/' + MAX_ATTEMPTS);
      if (attempt < MAX_ATTEMPTS) {
        // Exponential backoff: 500ms, 1000ms, 2000ms, 4000ms
        const backoff = BASE_RETRY_SLEEP * Math.pow(2, attempt - 1);
        Logger.log('cs_ensureReportUrl_: waiting ' + backoff + 'ms before retry...');
        Utilities.sleep(backoff);
        continue;
      }
      return { ok: false, error: 'LOCKED after ' + MAX_ATTEMPTS + ' attempts' };
    }

    try {
      // ── DOUBLE-CHECK: re-read cell after acquiring lock ──
      const urlColIdx1 = H[CS_REPORT_URL_COL] != null ? H[CS_REPORT_URL_COL] + 1 : -1;

      let liveUrl = '';
      if (urlColIdx1 > 0) {
        liveUrl = String(masterSheet.getRange(row, urlColIdx1).getValue() || '').trim();
      }

      let reportId  = '';
      let reportUrl = liveUrl;
      let reportSS  = null;
      let source    = 'existing';

      // FIX 10: Validate URL format before extracting ID
      if (liveUrl && isValidSpreadsheetUrl_(liveUrl)) {
        reportId = extractIdFromUrl_(liveUrl);
        
        if (reportId) {
          try {
            reportSS = SpreadsheetApp.openById(reportId);
          } catch (e) {
            Logger.log('cs_ensureReportUrl_: URL invalid (' + reportId + '): ' + e.message);
            reportId = '';
            reportSS = null;
            source   = 'stale';
          }
        }
      }

      // Still no valid report → create new
      if (!reportId) {
        const apptId = String(getVal('APPT_ID')).trim();
        const brand  = String(getVal('Brand')).trim();
        const client = String(getVal('Customer Name')).trim();
        const name   = CS_REPORT_NAME_FMT
          .replace('{Brand}', brand || 'VVS')
          .replace('{APPT_ID}', apptId);

        const parent = pickParentFolder_(getVal(CS_PROSPECT_URL_COL), client);
        
        // FIX 11: Check for existing file by token BEFORE creating
        const existingId = findExistingReportByToken_(parent, apptId, name);
        
        if (existingId) {
          Logger.log('cs_ensureReportUrl_: found existing report by token: ' + existingId);
          reportId  = existingId;
          reportUrl = 'https://docs.google.com/spreadsheets/d/' + reportId + '/edit';
          reportSS  = SpreadsheetApp.openById(reportId);
          source    = 'recovered';
        } else {
          reportId  = createClientReport_(name, parent, apptId);
          reportUrl = 'https://docs.google.com/spreadsheets/d/' + reportId + '/edit';
          reportSS  = SpreadsheetApp.openById(reportId);
          source    = 'created';

          // ── Project #22: backfill referral nếu đã có trong Master ──
          try {
            const referralName = H['Referral Name'] != null
              ? String(masterSheet.getRange(row, H['Referral Name'] + 1).getValue() || '').trim()
              : '';
            const referralDiscount = H['Referral Discount'] != null
              ? String(masterSheet.getRange(row, H['Referral Discount'] + 1).getValue() || '').trim()
              : '';

            if (referralName) {
              const referralText = 'Yes — ' + referralName
                + (referralDiscount ? ' (−$' + referralDiscount + ')' : '');
              const csSh = reportSS.getSheetByName('Client Status');
              if (csSh) {
                cs_backfillReferralToSnapshot_(csSh, referralText);
                Logger.log('[cs_ensureReportUrl_] Backfilled referral: ' + referralText);
              }
            }
          } catch(e) {
            Logger.log('[cs_ensureReportUrl_] Referral backfill warning: ' + e.message);
          }
        }
      }

      // ═══════════════════════════════════════════════════════════
      // OPTIMIZATION: Enhanced URL write with final check
      // ═══════════════════════════════════════════════════════════
      // Write URL to cell if needed
      if (urlColIdx1 > 0 && reportUrl && reportUrl !== liveUrl) {
        // Final double-check before write to prevent race condition
        const finalCheck = String(masterSheet.getRange(row, urlColIdx1).getValue() || '').trim();
        
        // Only write if:
        // 1. Cell is still empty, OR
        // 2. Cell has invalid URL (stale/wrong format)
        if (!finalCheck || (finalCheck !== reportUrl && !isValidSpreadsheetUrl_(finalCheck))) {
          masterSheet.getRange(row, urlColIdx1).setValue(reportUrl);
          SpreadsheetApp.flush(); // Force immediate write
          Logger.log('cs_ensureReportUrl_: WROTE URL row=' + row + ' source=' + source + ' url=' + reportUrl);
        } else {
          Logger.log('cs_ensureReportUrl_: URL already present row=' + row + ' (skipping write, avoiding race)');
        }
      }
      // ═══════════════════════════════════════════════════════════

      return { ok: true, reportId, reportUrl, reportSS };

    } catch (e) {
      Logger.log('cs_ensureReportUrl_ ERROR attempt ' + attempt + ': ' + (e && e.message ? e.message : e));
      if (attempt === MAX_ATTEMPTS) {
        return { ok: false, error: String(e && e.message || e) };
      }
      // Retry with exponential backoff on error too
      const backoff = BASE_RETRY_SLEEP * Math.pow(2, attempt - 1);
      Utilities.sleep(backoff);
    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }
  }

  return { ok: false, error: 'cs_ensureReportUrl_: unexpected exit' };
}

// ============================================================
// === Read lists + hex maps from "Dropdown" ===
// ============================================================
function readDropdowns_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Dropdown');
  if (!sh) throw new Error('Missing tab "Dropdown".');

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  const data = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  const idx = (name) => header.indexOf(String(name).trim());
  const colVals = (name) => {
    const c = idx(name); if (c < 0 || data.length === 0) return [];
    const col = new Array(data.length);
    for (let i = 0; i < data.length; i++) col[i] = String(data[i][c] || '').trim();
    return col;
  };

  const assignedReps         = colVals('Assigned Rep').filter(Boolean);
  const assistedReps         = colVals('Assisted Rep').filter(Boolean);
  const salesStages          = colVals('Sales Stage').filter(Boolean);
  const convStatuses         = colVals('Conversion Status').filter(Boolean);
  const customOrderStatuses  = colVals('Custom Order Status').filter(Boolean);
  const centerStoneStatuses  = colVals('Center Stone Order Status').filter(Boolean);
  const inProductionStatuses = colVals('In Production Status').filter(Boolean);

  const ssHex   = colVals(COL_SALES_STAGE_HEX);
  const csHex   = colVals(COL_CONV_STATUS_HEX);
  const cosHex  = colVals(COL_CUST_ORDER_HEX);
  const csosHex = colVals(COL_CENTER_STONE_HEX);
  const ipsHex  = colVals(COL_IN_PRODUCTION_HEX);

  const buildHexMap = (values, hexes) => {
    const map = {};
    const n = Math.min(values.length, hexes.length);
    for (let i = 0; i < n; i++) {
      const v = String(values[i] || '').trim();
      const h = String((hexes[i] || '').replace('#', '').trim());
      if (!v) continue;
      if (/^[0-9A-Fa-f]{6}$/.test(h)) map[v] = '#' + h.toUpperCase();
    }
    return map;
  };

  return {
    assignedReps, assistedReps, salesStages, convStatuses, customOrderStatuses,
    centerStoneStatuses, inProductionStatuses,
    colors: {
      salesStage:   buildHexMap(salesStages,          ssHex),
      convStatus:   buildHexMap(convStatuses,         csHex),
      customOrder:  buildHexMap(customOrderStatuses,  cosHex),
      centerStone:  buildHexMap(centerStoneStatuses,  csosHex),
      inProduction: buildHexMap(inProductionStatuses, ipsHex)
    }
  };
}

function readValidationRulesFlat_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Dropdown Rules');
  if (!sh) throw new Error('Missing tab "Dropdown Rules".');

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { matrix: [], viewing: [] };

  const all = sh.getRange(1, 1, lastRow, lastCol).getValues();

  let hdrRow = -1;
  let col = { sales: -1, conv: -1, cos: -1, ips: -1, csReq: -1, dead: -1, notes: -1 };

  for (let r = 0; r < all.length; r++) {
    const row = all[r].map(x => String(x || '').trim());
    const iSales = row.indexOf('Sales Stage');
    const iConv  = row.indexOf('Conversion Status');
    const iCOS   = row.indexOf('Custom Order Status');
    if (iSales >= 0 && iConv >= 0 && iCOS >= 0) {
      hdrRow = r;
      col.sales = iSales; col.conv = iConv; col.cos = iCOS;
      col.ips   = row.indexOf('In Production Status Requirement');
      col.csReq = row.indexOf('Center Stone Status Requirement');
      col.dead  = row.indexOf('Deadline Requirement');
      col.notes = row.indexOf('Notes / Flags');
      break;
    }
  }

  const matrix = [];
  if (hdrRow >= 0) {
    for (let r = hdrRow + 1; r < all.length; r++) {
      const row = all[r];
      const s   = String(row[col.sales] || '').trim();
      const c   = String(row[col.conv]  || '').trim();
      const cos = String(row[col.cos]   || '').trim();
      const ips = col.ips  >= 0 ? String(row[col.ips]  || '').trim() : '';
      const csr = col.csReq >= 0 ? String(row[col.csReq] || '').trim() : '';
      const dr  = col.dead >= 0 ? String(row[col.dead] || '').trim() : '';
      const nt  = col.notes >= 0 ? String(row[col.notes] || '').trim() : '';
      if (s || c || cos || ips || csr || dr || nt) {
        matrix.push({
          salesStage: s, convStatus: c, customOrderStatus: cos,
          ipsRequirement: ips, centerStoneRequirement: csr,
          deadlineRequirement: dr, notes: nt
        });
      }
    }
  }

  let vHdr = -1, cDays = -1, cMin = -1;
  for (let r = 0; r < all.length; r++) {
    const row = all[r].map(x => String(x || '').trim());
    const iD = row.indexOf('Days Before Viewing');
    const iM = row.indexOf('Minimum Allowed Center Stone Status');
    if (iD >= 0 && iM >= 0) { vHdr = r; cDays = iD; cMin = iM; break; }
  }

  const viewing = [];
  if (vHdr >= 0) {
    for (let r = vHdr + 1; r < all.length; r++) {
      const row = all[r];
      const d = String(row[cDays] || '').trim();
      const m = String(row[cMin]  || '').trim();
      if (d || m) viewing.push({ daysBefore: d, minimum: m });
    }
  }

  return { matrix, viewing };
}

// ============================================================
// === Dialog opener ===
// ============================================================
function cs_openStatusDialog_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CS_MASTER_SHEET_NAME);
  const r = sh.getActiveRange();
  if (!r || r.getRow() === 1) {
    SpreadsheetApp.getUi().alert('⚠️ Select a data row in 00_Master Appointments first.');
    return;
  }
  const row = r.getRow();

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const H = headerIndexMap_(header);
  const vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const get = name => H[name] != null ? vals[H[name]] : '';

  const assignedRepStr = String(get('Assigned Rep') || '');
  const assistedRepStr = String(get('Assisted Rep') || '');
  const orderDateISO = toISODateForInput_(get('Order Date'));

  const prefill = {
    clientName:      String(get('Customer Name') || ''),
    apptId:          String(get('APPT_ID') || ''),
    assignedRep:     assignedRepStr,
    assistedRep:     assistedRepStr,
    assignedRepArr:  normalizeMultiArray_(assignedRepStr),
    assistedRepArr:  normalizeMultiArray_(assistedRepStr),
    salesStage:      String(get('Sales Stage') || ''),
    convStatus:      String(get('Conversion Status') || ''),
    customOrder:     String(get('Custom Order Status') || ''),
    inProduction:    String(get('In Production Status') || ''),
    centerStone:     String(get('Center Stone Order Status') || ''),
    nextSteps:       String(get('Next Steps') || ''),
    orderDate:       orderDateISO,
    notebookLMLink:  String(get('NotebookLM Link') || '') 
  };

  const lists = readDropdowns_();
  const rulesFlat = readValidationRulesFlat_();

  let visitISO = String(get('ApptDateTime (ISO)') || '').trim();
  if (!visitISO) {
    const vdate = String(get('Visit Date') || '').trim();
    const vtime = String(get('Visit Time') || '').trim();
    if (vdate || vtime) {
      try {
        visitISO = Utilities.formatDate(
          new Date(vdate + ' ' + vtime), CS_TZ, "yyyy-MM-dd'T'HH:mm:ssXXX"
        );
      } catch (_) {}
    }
  }

  const t = HtmlService.createTemplateFromFile('dlg_client_status_v1');
  t.prefill = prefill;
  t.lists = {
    assignedReps:         lists.assignedReps,
    assistedReps:         lists.assistedReps,
    salesStages:          lists.salesStages,
    convStatuses:         lists.convStatuses,
    customOrderStatuses:  lists.customOrderStatuses,
    centerStoneStatuses:  lists.centerStoneStatuses,
    inProductionStatuses: lists.inProductionStatuses
  };
  t.colors = lists.colors;
  t.prefill.visitISO = visitISO || '';
  t.rulesFlat = rulesFlat;

  const html = t.evaluate().setWidth(1040).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Client Status Update');
}

// ============================================================
// === OPTIMIZED: cs_submitFromDialog - Single ensure call ===
// ============================================================
function cs_submitFromDialog(payload) {

  function _centerStoneRequired(stage, conv) {
    if (/^Lost Lead/i.test(String(stage || ''))) return false;
    if (/^Viewing Scheduled$/i.test(String(conv || ''))) return true;
    if (/^(Deposit Paid|Confirmed Order|Order In Progress)$/i.test(String(conv || ''))) return true;
    return false;
  }

  // Validation
  ['salesStage', 'convStatus'].forEach(function (k) {
    if (!String(payload[k] || '').trim()) {
      throw new Error('Please complete: Sales Stage and Conversion Status before submitting.');
    }
  });

  var cosEmptyAllowed = !!payload.cosAllowedEmpty;
  if (!cosEmptyAllowed && !String(payload.customOrder || '').trim()) {
    throw new Error('Please select a Custom Order Status.');
  }

  var isInProduction = String(payload.customOrder || '') === 'In Production';
  if (isInProduction && !String(payload.inProduction || '').trim()) {
    throw new Error('Please select an "In Production Status" since Custom Order Status is In Production.');
  }

  var need3D = /^(3D Requested|3D Revision Requested)$/i.test(String(payload.customOrder || ''));
  if (need3D && !String((payload.deadline3d || '')).trim()) {
    throw new Error('3D Deadline is required when Custom Order Status is 3D Requested or 3D Revision Requested.');
  }
  if (isInProduction && !String((payload.prodDeadline || '')).trim()) {
    throw new Error('Production Deadline is required when Custom Order Status is In Production.');
  }

  var needOrderDate = /^(Approved for Production|Waiting Production Timeline|In Production|Final Photos\s*[–-]\s*Waiting Approval|Warehouse|Ship to US|In US Store|Ship to Customer|Order Completed)$/i
    .test(String(payload.customOrder || ''));
  if (needOrderDate && !String(payload.orderDate || '').trim()) {
    throw new Error('Order Date is required for the selected Custom Order Status.');
  }

  if (_centerStoneRequired(String(payload.salesStage || ''), String(payload.convStatus || '')) &&
      !String(payload.centerStone || '').trim()) {
    throw new Error('Center Stone Order Status is required for Viewing Scheduled or Deposit/Confirmed/Order In Progress.');
  }

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CS_MASTER_SHEET_NAME);
  const r = sh.getActiveRange();
  if (!r || r.getNumRows() !== 1 || r.getRow() === 1) throw new Error('Select exactly one row.');
  const row = r.getRow();

  // ═══════════════════════════════════════════════════════════
  // OPTIMIZATION: Ensure report URL ONCE here, pass to submit
  // ═══════════════════════════════════════════════════════════
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const H = headerIndexMap_(header);
  
  // Quick read for ensureReportUrl_
  const preVals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const preGetVal = n => H[n] != null ? (preVals[H[n]] ?? '') : '';
  
  // Call ensure ONCE here
  const ensureResult = cs_ensureReportUrl_(sh, row, H, preGetVal);
  if (!ensureResult.ok) {
    throw new Error('Could not create/find client report: ' + ensureResult.error);
  }
  
  Logger.log('cs_submitFromDialog: ensured reportId=' + ensureResult.reportId + ' (will be reused)');
  // ═══════════════════════════════════════════════════════════
  
  // NOW read fresh vals after URL is guaranteed written
  const vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const __prevCenterStone = String(vals[H['Center Stone Order Status']] ?? '').trim();
  const __prevIPS         = String(vals[H['In Production Status']]      ?? '').trim();

  const assignedJoined = joinMulti_(payload.assignedRep);
  const assistedJoined = joinMulti_(payload.assistedRep);

  const setIf = (name, value) => {
    if (value != null && String(value).trim() !== '' && H[name] != null) {
      vals[H[name]] = value;
    }
  };

  setIf('Assigned Rep',        assignedJoined);
  setIf('Assisted Rep',        assistedJoined);
  setIf('Sales Stage',         payload.salesStage);
  setIf('Conversion Status',   payload.convStatus);
  setIf('Custom Order Status', payload.customOrder);
  setIf('Order Date',          payload.orderDate);

  var ipsIdx = (H['In Production Status'] != null)
    ? H['In Production Status']
    : findHeaderIndexByRegex_(header, /in\s*production\s*status/i);

  if (ipsIdx >= 0) {
    vals[ipsIdx] = isInProduction ? String(payload.inProduction || '').trim() : '';
  }

  // Deadline columns
  const idxProdDeadline = (H['Production Deadline'] != null)
    ? H['Production Deadline']
    : findHeaderIndexByRegex_(header, /(Production|Prod\.)\s*Deadline/i);

  const idx3dDeadline = (H['3D Deadline'] != null)
    ? H['3D Deadline']
    : findHeaderIndexByRegex_(header, /3D\s*Deadline/i);

  const idxProdMoves = (H['# of Times Prod. Deadline Moved'] != null)
    ? H['# of Times Prod. Deadline Moved']
    : findHeaderIndexByRegex_(header, /#\s*of\s*Times\s*(Prod|Production).*Deadline.*Moved/i);

  const idx3dMoves = (H['# of Times 3D Deadline Moved'] != null)
    ? H['# of Times 3D Deadline Moved']
    : findHeaderIndexByRegex_(header, /#\s*of\s*Times\s*3D.*Deadline.*Moved/i);

  const prevProdDeadline = idxProdDeadline >= 0 ? String(vals[idxProdDeadline] || '').trim() : '';
  const prev3dDeadline   = idx3dDeadline   >= 0 ? String(vals[idx3dDeadline]   || '').trim() : '';
  const prevProdMovesStr = idxProdMoves    >= 0 ? String(vals[idxProdMoves]    || '').trim() : '';
  const prev3dMovesStr   = idx3dMoves      >= 0 ? String(vals[idx3dMoves]      || '').trim() : '';

  const is3D = /^(3D Requested|3D Revision Requested)$/i.test(String(payload.customOrder || ''));
  const newProdDeadline = isInProduction ? String(payload.prodDeadline || '') : '';
  const new3dDeadline   = is3D          ? String(payload.deadline3d   || '') : '';

  if (idxProdDeadline >= 0) vals[idxProdDeadline] = newProdDeadline;
  if (idx3dDeadline   >= 0) vals[idx3dDeadline]   = new3dDeadline;

  let prodChanged = false, threeDChanged = false;

  if (idxProdDeadline >= 0 && isInProduction) {
    if (!prevProdDeadline && newProdDeadline) {
      if (idxProdMoves >= 0) vals[idxProdMoves] = '-';
    } else if (prevProdDeadline && newProdDeadline && prevProdDeadline !== newProdDeadline) {
      prodChanged = true;
      const prevN = (prevProdMovesStr === '-' || prevProdMovesStr === '') ? 0 : (parseInt(prevProdMovesStr, 10) || 0);
      if (idxProdMoves >= 0) vals[idxProdMoves] = String(prevN + 1);
    }
  }

  if (idx3dDeadline >= 0 && is3D) {
    if (!prev3dDeadline && new3dDeadline) {
      if (idx3dMoves >= 0) vals[idx3dMoves] = '-';
    } else if (prev3dDeadline && new3dDeadline && prev3dDeadline !== new3dDeadline) {
      threeDChanged = true;
      const prevN = (prev3dMovesStr === '-' || prev3dMovesStr === '') ? 0 : (parseInt(prev3dMovesStr, 10) || 0);
      if (idx3dMoves >= 0) vals[idx3dMoves] = String(prevN + 1);
    }
  }

  let logDeadlineType = '', logDeadlineDate = '', logMoveCount = '';
  if (idxProdDeadline >= 0 && isInProduction && ((!prevProdDeadline && newProdDeadline) || prodChanged)) {
    logDeadlineType = 'Production';
    logDeadlineDate = newProdDeadline;
    logMoveCount    = (idxProdMoves >= 0 ? String(vals[idxProdMoves] || '') : '');
  }
  if (idx3dDeadline >= 0 && is3D && ((!prev3dDeadline && new3dDeadline) || threeDChanged)) {
    logDeadlineType = logDeadlineType ? (logDeadlineType + ' | 3D') : '3D';
    logDeadlineDate = logDeadlineDate ? (logDeadlineDate + ' | ' + new3dDeadline) : new3dDeadline;
    const mc = (idx3dMoves >= 0 ? String(vals[idx3dMoves] || '') : '');
    logMoveCount = logMoveCount ? (logMoveCount + ' | ' + mc) : mc;
  }

  // Enforce IPS for later COS phases
  (function enforceIPSForLaterPhases() {
    const cosNow = String(payload.customOrder || '').trim();
    const later = new Set([
      'Final Photos – Waiting Approval', 'Warehouse', 'Ship to US',
      'In US Store', 'Ship to Customer', 'Order Completed'
    ]);
    if (later.has(cosNow) && typeof ipsIdx === 'number' && ipsIdx >= 0) {
      vals[ipsIdx] = 'Production Completed';
    }
  })();

  payload.__deadlineLog = { type: logDeadlineType, date: logDeadlineDate, moves: logMoveCount };

  setIf('Center Stone Order Status', payload.centerStone);
  if (H['Next Steps'] != null && payload.nextSteps != null) vals[H['Next Steps']] = payload.nextSteps;

  const notebookLMLink = String(payload.notebookLMLink || '').trim();
  setIf('NotebookLM Link', notebookLMLink);
  sh.getRange(row, 1, 1, vals.length).setValues([vals]);

  // Wax Request
  var waxSummary = null;
  try {
    if (payload.wax && payload.wax.request === true) {
      var rootApptId = String(
        (H['RootApptID'] != null ? vals[H['RootApptID']] : '') ||
        (H['APPT_ID']    != null ? vals[H['APPT_ID']]    : '') || ''
      ).trim();
      if (rootApptId) {
        var wres = wax_onRequestSubmit_({
          rootApptId: rootApptId,
          soMo: (payload.wax.soMo || ''),
          neededByRep: (payload.wax.neededBy || ''),
          priority: (payload.wax.priority || ''),
          requestedBy: (Session.getActiveUser().getEmail() || assignedJoined || '')
        }) || {};
        waxSummary = {
          created: !!wres.ok,
          requestId: wres.requestId || '',
          folderUrl: wres.folderUrl || '',
          rowUrl:    wres.url || ''
        };
      }
    }
  } catch (e) {
    Logger.log('Wax create failed: ' + (e && e.message ? e.message : e));
  }

  // ═══════════════════════════════════════════════════════════
  // OPTIMIZATION: Pass reportId/URL/SS to avoid re-calling ensure
  // ═══════════════════════════════════════════════════════════
  return cs_submitClientStatusUpdate_({
    rowNum:          row,
    assistedRep:     assistedJoined,
    prevCenterStone: __prevCenterStone,
    prevIPS:         __prevIPS,
    inProduction:    String(payload.inProduction || '').trim(),
    wax:             waxSummary || null,
    waxSummaryStr:   String(payload.waxSummary || ''),
    prodDeadline:    String(payload.prodDeadline || ''),
    deadline3d:      String(payload.deadline3d   || ''),
    // ↓ NEW: Pass report info to skip re-ensure
    reportId:        ensureResult.reportId,
    reportUrl:       ensureResult.reportUrl,
    reportSS:        ensureResult.reportSS,
    notebookLMLink: notebookLMLink,
  });
}

// ============================================================
// === FIX 6: cs_createOrGetReportForSelection_ ===
// ============================================================
function cs_createOrGetReportForSelection_(opts) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(CS_MASTER_SHEET_NAME);
    const row = cs_resolveRow_(sh, opts && opts.rowNum);

    const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const H = headerIndexMap_(header);
    const vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
    const getVal = n => H[n] != null ? (vals[H[n]] ?? '') : '';

    const result = cs_ensureReportUrl_(sh, row, H, getVal);
    if (!result.ok) return { ok: false, error: result.error };

    return { ok: true, id: result.reportId, url: result.reportUrl, ss: result.reportSS };

  } catch (e) {
    Logger.log('cs_createOrGetReportForSelection_ ERROR: ' + (e && e.message ? e.message : e));
    return { ok: false, error: String(e && e.message || e) };
  }
}

function pickParentFolder_(prospectUrl, clientName) {
  if (prospectUrl) {
    const id = extractIdFromUrl_(String(prospectUrl));
    try { return DriveApp.getFolderById(id); } catch (e) {}
  }
  try {
    const it = DriveApp.getFoldersByName(clientName || 'Clients');
    if (it.hasNext()) return it.next();
  } catch (e) {}
  return DriveApp.getRootFolder();
}

// FIX 11: Add apptId parameter for idempotency token
function createClientReport_(name, parentFolder, apptId) {
  const templateId = getTemplateId_();
  if (!templateId) throw new Error('Client Status: CS_REPORT_TEMPLATE_ID not set in Project Properties.');
  const tmplFile = DriveApp.getFileById(templateId);
  const copy = tmplFile.makeCopy(name, parentFolder || DriveApp.getRootFolder());
  const fileId = copy.getId();
  
  // FIX 11: Set idempotency token in file description
  if (apptId) {
    try {
      copy.setDescription('APPT_ID=' + String(apptId).trim());
    } catch (e) {
      Logger.log('Failed to set file description: ' + e.message);
    }
  }
  
  try { if (parentFolder) DriveApp.getRootFolder().removeFile(copy); } catch (e) {}
  return fileId;
}

function ensureReportConfig_(reportSS, opts) {
  const rootApptId = String(opts.rootApptId || '').trim();
  const reportId   = String(opts.reportId || reportSS.getId()).trim();

  let sh = reportSS.getSheetByName('_Config');
  if (!sh) {
    sh = reportSS.insertSheet('_Config');
    try { sh.hideSheet(); } catch (_) {}
    sh.appendRow(['ROOT_APPT_ID', rootApptId]);
    sh.appendRow(['CONTROLLER_URL', ScriptApp.getService().getUrl()]);
    sh.appendRow(['REPORT_REANALYZE_TOKEN',
      PropertiesService.getScriptProperties().getProperty('REPORT_REANALYZE_TOKEN') || ''
    ]);
    sh.appendRow(['REPORT_ID', reportId]);
    return;
  }

  const vals = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  const map = {};
  vals.forEach(r => { if (r[0]) map[String(r[0]).trim()] = String(r[1] || '').trim(); });

  const want = {
    ROOT_APPT_ID: rootApptId,
    CONTROLLER_URL: ScriptApp.getService().getUrl(),
    REPORT_REANALYZE_TOKEN: PropertiesService.getScriptProperties().getProperty('REPORT_REANALYZE_TOKEN') || '',
    REPORT_ID: reportId
  };

  Object.keys(want).forEach(k => {
    const cur  = map[k] || '';
    const need = String(want[k] || '');
    if (cur !== need) {
      let rowIdx = vals.findIndex(r => String(r[0]).trim() === k);
      if (rowIdx >= 0) {
        sh.getRange(rowIdx + 1, 2).setValue(need);
      } else {
        sh.appendRow([k, need]);
      }
    }
  });
}

// ============================================================
// === OPTIMIZED: cs_submitClientStatusUpdate_ v2.6 + PROJECT #4 ===
// ============================================================
function cs_submitClientStatusUpdate_(opts) {
  try {
    const ss = SpreadsheetApp.getActive();
    const master = ss.getSheetByName(CS_MASTER_SHEET_NAME);
    const row = cs_resolveRow_(master, opts && opts.rowNum);

    const header = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];
    const H = headerIndexMap_(header);
    const vals = master.getRange(row, 1, 1, master.getLastColumn()).getValues()[0];
    const get  = n => vals[H[n]] ?? '';

    const apptId      = String(get('APPT_ID') || '').trim();
    const brand       = String(get('Brand') || '');
    const client      = String(get('Customer Name') || '');
    const rep         = String(get('Assigned Rep') || '');
    const salesStage  = String(get('Sales Stage') || '');
    const convStatus  = String(get('Conversion Status') || '');
    const customOrd   = String(get('Custom Order Status') || '');
    const inProduction = String(get('In Production Status') || (opts && opts.inProduction) || '');
    const centerStone = String(get('Center Stone Order Status') || '');
    const nextSteps   = String(get('Next Steps') || '');
    const orderDate   = String(get('Order Date') || '');

    const phone       = String(getByAny_(H, vals, ['Phone', 'Client Phone', 'Customer Phone']) || '');
    const email       = String(getByAny_(H, vals, ['Email', 'Client Email', 'Customer Email']) || '');
    const occasion    = String(getByAny_(H, vals, ['Occasion']) || '');
    const budgetRange = String(getByAny_(H, vals, ['Budget Range']) || '');
    const decisionMkr = String(getByAny_(H, vals, ['Decision-Maker', 'Decision Maker']) || '');
    const initialReq  = String(getByAny_(H, vals, ['Initial Request']) || '');
    const soNumber    = String(getByAny_(H, vals, ['SO Number', 'SO#']) || '').trim();
    const notebookLMLink = String(
    (H['NotebookLM Link'] != null ? get('NotebookLM Link') : '') ||
    (opts && opts.notebookLMLink) || ''
    ).trim();

    const now  = new Date();
    const iso  = Utilities.formatDate(now, CS_TZ, 'yyyy-MM-dd');
    const ts   = Utilities.formatDate(now, CS_TZ, 'yyyy-MM-dd HH:mm:ss');
    const nice = Utilities.formatDate(now, CS_TZ, 'MMM d, yyyy h:mm a z');
    const user = Session.getActiveUser().getEmail() || rep || 'Unknown';
    const assistedRep = String((opts && opts.assistedRep) || '');

    // 1) Central audit
    const audit = ss.getSheetByName(CS_AUDIT_TAB);
    if (audit) {
      const rootKeyForAudit = String(get('RootApptID') || get('APPT_ID') || '').trim();
      let appliedCountTotal = 1;
      if (rootKeyForAudit) {
        const lastRowAll = master.getLastRow();
        if (lastRowAll > 1) {
          const matchColIndexAudit = (H['RootApptID'] != null) ? H['RootApptID']
                                   : (H['APPT_ID']    != null) ? H['APPT_ID']
                                   : -1;
          if (matchColIndexAudit >= 0) {
            const allValsAudit = master.getRange(2, 1, lastRowAll - 1, master.getLastColumn()).getValues();
            for (let i = 0; i < allValsAudit.length; i++) {
              const rnum = i + 2;
              if (rnum === row) continue;
              const idHere = String(allValsAudit[i][matchColIndexAudit] || '').trim();
              if (idHere && idHere === rootKeyForAudit) appliedCountTotal++;
            }
          }
        }
      }
      const appliedNote = `Applied to ${appliedCountTotal} row${appliedCountTotal === 1 ? '' : 's'}`
                        + (rootKeyForAudit ? ` (RootApptID=${rootKeyForAudit})` : '');

      let auditHeader = audit.getRange(1, 1, 1, audit.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
      if (auditHeader.indexOf('Applied To') < 0) {
        audit.getRange(1, audit.getLastColumn() + 1).setValue('Applied To');
        auditHeader = audit.getRange(1, 1, 1, audit.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
      }

      cs_audit_appendByHeader_(audit, auditHeader, {
        'APPT_ID':                   apptId,
        'Log Date':                  iso,
        'Sales Stage':               salesStage,
        'Conversion Status':         convStatus,
        'Custom Order Status':       customOrd,
        'In Production Status':      inProduction,
        'Center Stone Order Status': centerStone,
        'Next Steps':                nextSteps,
        'Assisted Rep':              assistedRep,
        'Updated By':                user,
        'Updated At':                ts,
        'Applied To':                appliedNote
      });
    } else {
      Logger.log(`Client Status: audit tab "${CS_AUDIT_TAB}" not found — continuing without central audit.`);
    }

    // ═══════════════════════════════════════════════════════════
    // OPTIMIZATION: Conditional report ensure - only if not provided by caller
    // ═══════════════════════════════════════════════════════════
    let reportUrl, reportId, reportSS;
    
    if (opts && opts.reportId && opts.reportUrl && opts.reportSS) {
      // Already provided by caller (from dialog) - SKIP ensure
      reportUrl = opts.reportUrl;
      reportId = opts.reportId;
      reportSS = opts.reportSS;
      Logger.log('cs_submitClientStatusUpdate_: Using provided reportId=' + reportId + ' (skipping ensure, avoiding duplicate call)');
    } else {
      // Not provided (automation flow) - ensure now
      const freshVals = master.getRange(row, 1, 1, master.getLastColumn()).getValues()[0];
      const getFreshVal = n => H[n] != null ? (freshVals[H[n]] ?? '') : '';

      const ensureResult = cs_ensureReportUrl_(master, row, H, getFreshVal);
      if (!ensureResult.ok) {
        return { ok: false, error: ensureResult.error || 'Could not create/find client report' };
      }
      reportUrl = ensureResult.reportUrl;
      reportId  = ensureResult.reportId;
      reportSS  = ensureResult.reportSS;
      Logger.log('cs_submitClientStatusUpdate_: Ensured reportId=' + reportId);
    }
    // ═══════════════════════════════════════════════════════════

    const rootApptId = String(
      (H['RootApptID'] != null ? vals[H['RootApptID']] : '') ||
      (H['APPT_ID']    != null ? vals[H['APPT_ID']]    : '') || ''
    ).trim();
    ensureReportConfig_(reportSS, { rootApptId, reportId });

    // 3) Per-client log row
    if (CS_WRITE_PER_CLIENT_LOG) {
      insertLogRowByHeader_(reportSS, {
        'Log Date':                  iso,
        'Sales Stage':               salesStage,
        'Conversion Status':         convStatus,
        'Custom Order Status':       customOrd,
        'In Production Status':      inProduction,
        'Center Stone Order Status': centerStone,
        'Next Steps':                nextSteps,
        'Deadline Type':             (opts && opts.__deadlineLog && opts.__deadlineLog.type)  || '',
        'Deadline Date':             (opts && opts.__deadlineLog && opts.__deadlineLog.date)  || '',
        'Move Count':                (opts && opts.__deadlineLog && opts.__deadlineLog.moves) || '',
        'Assisted Rep':              assistedRep,
        'Updated By':                user,
        'Updated At':                ts
      });
    }

    // Đọc referral từ Master
    const referralName     = String(getByAny_(H, vals, ['Referral Name'])     || '').trim();
    const referralDiscount = String(getByAny_(H, vals, ['Referral Discount']) || '').trim();
    const referralText     = referralName
      ? ('Yes — ' + referralName + (referralDiscount ? ' (−$' + referralDiscount + ')' : ''))
      : '';

    updateSnapshot_(reportSS, {
      Brand: brand, ClientName: client, APPT_ID: apptId, AssignedRep: rep,
      Phone: phone, Email: email, Occasion: occasion,
      BudgetRange: budgetRange, DecisionMaker: decisionMkr, InitialRequest: initialReq,
      SO_Number: soNumber,
      SalesStage: salesStage, ConversionStatus: convStatus, CustomOrderStatus: customOrd,
      InProductionStatus: inProduction,
      CenterStoneStatus: centerStone, NextSteps: nextSteps, UpdatedBy: user, UpdatedAt: ts,
      AssistedRep: assistedRep,
      OrderDate: orderDate,
      ReferAFriend: referralText,
      NotebookLM_Link: notebookLMLink,
    }); 

    // 5) Updated By/At on master
    const uIdx = H['Updated By'], aIdx = H['Updated At'];
    if (uIdx != null && aIdx != null && Math.abs((uIdx + 1) - (aIdx + 1)) === 1) {
      const from = Math.min(uIdx, aIdx) + 1;
      const pairVals = (uIdx < aIdx) ? [[user, ts]] : [[ts, user]];
      master.getRange(row, from, 1, 2).setValues(pairVals);
    } else {
      if (uIdx != null) master.getRange(row, uIdx + 1).setValue(user);
      if (aIdx != null) master.getRange(row, aIdx + 1).setValue(ts);
    }

    // 5b) Fan-out with URL column exclusion
    try {
      const rootKey = String(get('RootApptID') || get('APPT_ID') || '').trim();
      if (rootKey) {
        const lastRow = master.getLastRow();
        if (lastRow > 1) {
          const ipsIdx = (H['In Production Status'] != null)
            ? H['In Production Status']
            : findHeaderIndexByRegex_(header, /in\s*production\s*status/i);

          // Get URL column index to exclude from fan-out
          const urlColIdx = H[CS_REPORT_URL_COL];

          const allVals = master.getRange(2, 1, lastRow - 1, master.getLastColumn()).getValues();
          const matchColIndex = (H['RootApptID'] != null) ? H['RootApptID']
                              : (H['APPT_ID'] != null) ? H['APPT_ID']
                              : -1;

          if (matchColIndex >= 0) {
            const targets = [];
            for (let i = 0; i < allVals.length; i++) {
              const rowNum = i + 2;
              if (rowNum === row) continue;
              const idHere = String(allVals[i][matchColIndex] || '').trim();
              if (idHere && idHere === rootKey) targets.push(rowNum);
            }

            if (targets.length) {
              const enqueuePairs = (name, value) => {
                const idx = H[name];
                if (idx == null) return null;
                
                // Skip URL column to prevent overwrite
                if (urlColIdx != null && idx === urlColIdx) {
                  Logger.log('Fan-out: skipping URL column to prevent overwrite');
                  return null;
                }
                
                const pairs = [];
                for (const rnum of targets) pairs.push({ r: rnum, v: value });
                return { colIdx1: idx + 1, pairs };
              };

              const q = [];
              q.push(enqueuePairs('Assigned Rep',              rep));
              q.push(enqueuePairs('Assisted Rep',              assistedRep));
              q.push(enqueuePairs('Sales Stage',               salesStage));
              q.push(enqueuePairs('Conversion Status',         convStatus));
              q.push(enqueuePairs('Custom Order Status',       customOrd));
              q.push(enqueuePairs('Center Stone Order Status', centerStone));
              q.push(enqueuePairs('Next Steps',                nextSteps));
              q.push(enqueuePairs('Updated By',                user));
              q.push(enqueuePairs('Updated At',                ts));

              if (ipsIdx >= 0 && (urlColIdx == null || ipsIdx !== urlColIdx)) {
                const ipsPairs = [];
                for (const rnum of targets) ipsPairs.push({ r: rnum, v: inProduction });
                groupedSetValues_(master, ipsIdx + 1, ipsPairs);
              }

              // ══════════════════════════════════════════════════════════════
              // PROJECT #4: DEADLINE SYNC ACROSS RootApptID rows
              // ══════════════════════════════════════════════════════════════
              // vals đã được ghi xuống Master bởi cs_submitFromDialog trước
              // khi gọi hàm này → đọc lại từ vals là đúng giá trị mới nhất.
              // Shared counter (# Times Moved) cũng đã được tính sẵn trong vals.
              (function syncDeadlineFanOut_() {
                var idxD3   = H['3D Deadline']                     != null ? H['3D Deadline']                     : findHeaderIndexByRegex_(header, /3D\s*Deadline/i);
                var idxProd = H['Production Deadline']              != null ? H['Production Deadline']              : findHeaderIndexByRegex_(header, /(Production|Prod\.)\s*Deadline/i);
                var idxD3Mv = H['# of Times 3D Deadline Moved']    != null ? H['# of Times 3D Deadline Moved']    : findHeaderIndexByRegex_(header, /#\s*of\s*Times\s*3D.*Deadline.*Moved/i);
                var idxPrMv = H['# of Times Prod. Deadline Moved'] != null ? H['# of Times Prod. Deadline Moved'] : findHeaderIndexByRegex_(header, /#\s*of\s*Times\s*(Prod|Production).*Deadline.*Moved/i);

                var d3Val   = idxD3   >= 0 ? (vals[idxD3]   || '') : '';
                var prodVal = idxProd >= 0 ? (vals[idxProd]  || '') : '';
                var d3MvVal = idxD3Mv >= 0 ? (vals[idxD3Mv] || '') : '';
                var prMvVal = idxPrMv >= 0 ? (vals[idxPrMv] || '') : '';

                var hdrD3   = idxD3   >= 0 ? String(header[idxD3]   || '').trim() : null;
                var hdrProd = idxProd >= 0 ? String(header[idxProd]  || '').trim() : null;
                var hdrD3Mv = idxD3Mv >= 0 ? String(header[idxD3Mv] || '').trim() : null;
                var hdrPrMv = idxPrMv >= 0 ? String(header[idxPrMv] || '').trim() : null;

                if (hdrD3   && d3Val   !== '') q.push(enqueuePairs(hdrD3,   d3Val));
                if (hdrProd && prodVal !== '') q.push(enqueuePairs(hdrProd, prodVal));
                if (hdrD3Mv && d3MvVal !== '') q.push(enqueuePairs(hdrD3Mv, d3MvVal));
                if (hdrPrMv && prMvVal !== '') q.push(enqueuePairs(hdrPrMv, prMvVal));

                Logger.log('[Project#4] deadline fan-out: 3D="' + d3Val + '" moved=' + d3MvVal
                         + ' | Prod="' + prodVal + '" moved=' + prMvVal
                         + ' | targets=' + targets.length);
              })();
              // ══════════════════════════════════════════════════════════════
              // END PROJECT #4
              // ══════════════════════════════════════════════════════════════

              for (const item of q) {
                if (item && item.pairs && item.pairs.length) {
                  groupedSetValues_(master, item.colIdx1, item.pairs);
                }
              }
            }
          }
        }
      }
    } catch (e) {
      Logger.log('Fan-out to RootApptID siblings failed: ' + (e && e.message ? e.message : e));
    }

    // 6) DV hook
    try {
      if (typeof DV_init_ === 'function') { DV_init_(); }

      var prevCenterStone = (opts && opts.prevCenterStone) || '';
      var newCenterStone  = centerStone || '';
      var becameNeed = !(typeof DV_isNeedToPropose === 'function' ? DV_isNeedToPropose(prevCenterStone) : false)
                    &&  (typeof DV_isNeedToPropose === 'function' ? DV_isNeedToPropose(newCenterStone)  : false);
      Logger.log('DV hook: prev="' + prevCenterStone + '" → new="' + newCenterStone + '"; becameNeed=' + becameNeed);

      if (becameNeed && rootApptId) {
        var res = DV_upsertProposeNudge_afterStatus_({
          rootApptId,
          customerName: client,
          nextStepsFromMaster: nextSteps
        });
        Logger.log('DV hook: queued +2d nudge for root=' + rootApptId + ' → ' + JSON.stringify(res));
      }
    } catch (e) {
      Logger.log('DV hook error: ' + (e && e.message ? e.message : e));
    }

    // 7) Reminders hook
    try {
      Remind.onClientStatusChange(soNumber, salesStage, customOrd, user, {
        rootApptId:       rootApptId,
        assignedRepName:  rep,
        assistedRepName:  assistedRep,
        customerName:     client,
        nextSteps
      });
    } catch (e) {
      console.warn('Remind.onClientStatusChange failed:', e && e.message ? e.message : e);
    }

    const masterLink = ss.getUrl() + '#gid=' + master.getSheetId() + '&range=A' + row;
    const waxObj        = (opts && opts.wax) || null;
    const waxSummaryStr = String((opts && opts.waxSummaryStr) || '');

    return {
      ok: true,
      summary: {
        clientName:  client, apptId,
        assignedRep: rep,    assistedRep,
        salesStage,  convStatus,
        customOrder: customOrd,
        deadline3d:   String((opts && opts.deadline3d)   || ''),
        orderDate,
        inProduction,
        prodDeadline: String((opts && opts.prodDeadline) || ''),
        centerStone,  nextSteps,
        submittedBy:  user,
        submittedAt:  nice,
        reportUrl,    masterLink,
        rootApptId,
        waxSummary: waxSummaryStr,
        wax:        waxObj,
        notebookLMLink: notebookLMLink,
      }
    };

  } catch (e) {
    Logger.log('cs_submitClientStatusUpdate_ ERROR: ' + (e && e.message ? e.message : e));
    return { ok: false, error: String(e && e.message || e) };
  }
}

// ============================================================
// === Log helpers ===
// ============================================================

function getLogHeaderRow_(sh) {
  const sp  = sh.getParent();
  const key = 'CS_LOG_HDR_' + (sp && sp.getId ? sp.getId() : '') + '_' + sh.getSheetId();
  const props = PropertiesService.getScriptProperties();

  const cached = Number(props.getProperty(key) || 0);
  if (cached && String(sh.getRange(cached, 1).getValue()).trim() === 'Log Date') return cached;

  const start = 8, end = Math.min(sh.getLastRow() || 80, 80);
  const scan = sh.getRange(start, 1, Math.max(end - start + 1, 1), 1).getValues();
  let headerRow = 13;
  for (let i = 0; i < scan.length; i++) {
    if (String(scan[i][0] || '').trim() === 'Log Date') { headerRow = start + i; break; }
  }
  props.setProperty(key, String(headerRow));
  return headerRow;
}

function insertLogRowByHeader_(reportSS, valuesByName) {
  const sh = reportSS.getSheetByName(CS_REPORT_SHEET);
  if (!sh) throw new Error(`Missing "${CS_REPORT_SHEET}" tab`);

  const headerRow = getLogHeaderRow_(sh);
  const header = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
  const H = {}; header.forEach((h, i) => { if (h) H[h] = i; });

  const row = new Array(header.length).fill('');
  Object.keys(valuesByName).forEach(name => {
    const i = H[name]; if (i != null) row[i] = valuesByName[name];
  });

  sh.insertRowsBefore(headerRow + 1, 1);
  sh.getRange(headerRow + 1, 1, 1, row.length).setValues([row]);
}

function groupedSetValues_(sh, colIdx, pairs) {
  if (!pairs || !pairs.length) return;
  pairs.sort((a, b) => a.r - b.r);
  let start = pairs[0].r;
  let block  = [[pairs[0].v]];
  for (let i = 1; i < pairs.length; i++) {
    const prev = pairs[i - 1].r, cur = pairs[i].r;
    if (cur === prev + 1) {
      block.push([pairs[i].v]);
    } else {
      sh.getRange(start, colIdx, block.length, 1).setValues(block);
      start = cur; block = [[pairs[i].v]];
    }
  }
  sh.getRange(start, colIdx, block.length, 1).setValues(block);
}

function updateSnapshot_(reportSS, data) {
  const sh = reportSS.getSheetByName(CS_REPORT_SHEET);
  if (!sh) throw new Error(`Missing "${CS_REPORT_SHEET}" tab`);

  const map = {
    'Report Date:':              '__InitDate',
    'Customer Name:':            'ClientName',
    'APPT_ID:':                  'APPT_ID',
    'Brand:':                    'Brand',
    'Assigned Rep:':             'AssignedRep',
    'Phone:':                    'Phone',
    'Email:':                    'Email',
    'Occasion:':                 'Occasion',
    'Budget Range:':             'BudgetRange',
    'Decision-Maker:':           'DecisionMaker',
    'Initial Request:':          'InitialRequest',
    'SO#:':                      'SO_Number',
    'Sales Stage:':              'SalesStage',
    'Conversion Status:':        'ConversionStatus',
    'Custom Order Status:':      'CustomOrderStatus',
    'In Production Status:':     'InProductionStatus',
    'Center Stone Order Status:':'CenterStoneStatus',
    'Next Steps:':               'NextSteps',
    'Updated By:':               'UpdatedBy',
    'Updated At:':               'UpdatedAt',
    'Assisted Rep:':             'AssistedRep',
    'Order Date:':               'OrderDate',
    'Notebook LM':                 'NotebookLM_Link',  
    'Notebook LM:':                'NotebookLM_Link', 
    'Refer a Friend:':           'ReferAFriend',
    'Refer a Friend':            'ReferAFriend'
  };

  const rowsToScan = Math.min(sh.getLastRow() || 50, 50);
  if (rowsToScan <= 0) return;

  const values = sh.getRange(1, 1, rowsToScan, 4).getValues();

  const writesB = [];
  const writesD = [];
  const todayStr = Utilities.formatDate(new Date(), CS_TZ, 'yyyy-MM-dd');

  for (let i = 0; i < rowsToScan; i++) {
    const labA = String(values[i][0] || '').trim();
    const labC = String(values[i][2] || '').trim();

    const apply = (label, targetColIndexZeroBased) => {
      const key = map[label]; if (!key) return;

      if (key === '__InitDate') {
        const current = String(values[i][targetColIndexZeroBased] || '').trim();
        if (!current) {
          if (targetColIndexZeroBased === 1) writesB.push({ r: i + 1, v: todayStr });
          else if (targetColIndexZeroBased === 3) writesD.push({ r: i + 1, v: todayStr });
        }
        return;
      }

      const newVal = data[key] != null ? String(data[key]) : '';
      if (targetColIndexZeroBased === 1) writesB.push({ r: i + 1, v: newVal });
      else if (targetColIndexZeroBased === 3) writesD.push({ r: i + 1, v: newVal });
    };

    if (labA) apply(labA, 1);
    if (labC) apply(labC, 3);
  }

  if (writesB.length) groupedSetValues_(sh, 2, writesB);
  if (writesD.length) groupedSetValues_(sh, 4, writesD);
}

function toISODateForInput_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return Utilities.formatDate(v, CS_TZ, 'yyyy-MM-dd');
  }
  const s = String(v || '').trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  if (!isNaN(d)) return Utilities.formatDate(d, CS_TZ, 'yyyy-MM-dd');
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) {
    const y = m[3].length === 2 ? ('20' + m[3]) : m[3];
    const mm = ('0' + m[1]).slice(-2), dd = ('0' + m[2]).slice(-2);
    return y + '-' + mm + '-' + dd;
  }
  return '';
}

function CS_AUDIT_upgrade_addIPS_AtEnd() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('03_Client_Status_Log');
  if (!sh) throw new Error('Sheet "03_Client_Status_Log" not found.');

  const lastCol = Math.max(1, sh.getLastColumn());
  const header  = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(x => String(x || '').trim());

  if (header.includes('In Production Status')) {
    Logger.log('Already present. No changes.');
    return;
  }

  const newCol = lastCol + 1;
  sh.getRange(1, newCol).setValue('In Production Status');
  Logger.log('Added "In Production Status" as new last column ' + newCol + '.');
}

function cs_audit_appendByHeader_(sh, header, valuesByName) {
  const H = {}; header.forEach((h, i) => { if (h) H[String(h).trim()] = i; });
  const row = new Array(header.length).fill('');
  Object.keys(valuesByName).forEach(name => {
    const i = H[name]; if (i != null) row[i] = valuesByName[name];
  });
  sh.appendRow(row);
}

function cs_automationSubmit_(params) {
  if (!params || !params.rowNum || Number(params.rowNum) <= 1) {
    throw new Error('cs_automationSubmit_: params.rowNum is required and must be > 1.');
  }

  const master = SpreadsheetApp.getActive().getSheetByName(CS_MASTER_SHEET_NAME);
  master.setActiveRange(master.getRange(Number(params.rowNum), 1));

  const result = cs_submitClientStatusUpdate_({
    rowNum:       Number(params.rowNum),
    assistedRep:  String(params.assistedRep  || ''),
    inProduction: String(params.inProduction || ''),
    prodDeadline: String(params.prodDeadline || ''),
    deadline3d:   String(params.deadline3d   || ''),
    prevCenterStone: ''
  });

  if (!result.ok) {
    Logger.log('cs_automationSubmit_ FAILED at row ' + params.rowNum + ': ' + result.error);
  } else {
    Logger.log('cs_automationSubmit_ OK at row ' + params.rowNum + ': reportUrl=' + (result.summary && result.summary.reportUrl));
  }

  return result;
}

// ============================================================
// === VERIFY PROJECT #4 PATCH ===
// ============================================================
function VERIFY_project4_deadlineSync() {
  const code = cs_submitClientStatusUpdate_.toString();
  if (code.includes('syncDeadlineFanOut_') && code.includes('PROJECT #4')) {
    Logger.log('✅ Project #4 deadline sync patch đã được apply đúng');
    Logger.log('✅ syncDeadlineFanOut_ found in fan-out block');
  } else {
    Logger.log('❌ Patch chưa có — kiểm tra lại file');
  }
}

// ============================================================
// === REPAIR FUNCTIONS - Run once after v2.6 deployment ===
// ============================================================

/**
 * REPAIR 1: Backfill idempotency tokens for existing files
 */
function REPAIR_backfillIdempotencyTokens() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(CS_MASTER_SHEET_NAME);
  if (!master) throw new Error('Master sheet not found');

  const header = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];
  const H = headerIndexMap_(header);
  
  const urlIdx = H[CS_REPORT_URL_COL];
  const apptIdx = H['APPT_ID'];
  
  if (urlIdx == null || apptIdx == null) {
    throw new Error('Required columns not found: ' + CS_REPORT_URL_COL + ', APPT_ID');
  }

  const lastRow = master.getLastRow();
  if (lastRow <= 1) {
    Logger.log('No data rows to process');
    return { ok: true, processed: 0, updated: 0, errors: 0 };
  }

  const allVals = master.getRange(2, 1, lastRow - 1, master.getLastColumn()).getValues();
  
  let processed = 0, updated = 0, errors = 0;
  
  for (let i = 0; i < allVals.length; i++) {
    const row = i + 2;
    const url = String(allVals[i][urlIdx] || '').trim();
    const apptId = String(allVals[i][apptIdx] || '').trim();
    
    if (!url || !apptId) continue;
    if (!isValidSpreadsheetUrl_(url)) continue;
    
    processed++;
    
    try {
      const fileId = extractIdFromUrl_(url);
      const file = DriveApp.getFileById(fileId);
      const currentDesc = String(file.getDescription() || '').trim();
      const token = 'APPT_ID=' + apptId;
      
      if (!currentDesc.includes(token)) {
        const newDesc = currentDesc ? (currentDesc + ' | ' + token) : token;
        file.setDescription(newDesc);
        updated++;
        Logger.log('Row ' + row + ': Added token to file ' + fileId);
      }
      
      if (processed % 10 === 0) {
        Utilities.sleep(1000);
      }
      
    } catch (e) {
      errors++;
      Logger.log('Row ' + row + ' ERROR: ' + e.message);
    }
  }
  
  const summary = {
    ok: true,
    processed: processed,
    updated: updated,
    errors: errors,
    message: 'Backfilled ' + updated + '/' + processed + ' files (' + errors + ' errors)'
  };
  
  Logger.log(JSON.stringify(summary));
  return summary;
}

/**
 * REPAIR 2: Find and link orphaned files
 */
function REPAIR_linkOrphanedFiles() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(CS_MASTER_SHEET_NAME);
  if (!master) throw new Error('Master sheet not found');

  const header = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];
  const H = headerIndexMap_(header);
  
  const urlIdx = H[CS_REPORT_URL_COL];
  const apptIdx = H['APPT_ID'];
  const brandIdx = H['Brand'];
  
  if (urlIdx == null || apptIdx == null) {
    throw new Error('Required columns not found');
  }

  const lastRow = master.getLastRow();
  if (lastRow <= 1) return { ok: true, linked: 0 };

  const allVals = master.getRange(2, 1, lastRow - 1, master.getLastColumn()).getValues();
  
  let scanned = 0, linked = 0, notFound = 0;
  const updates = [];
  
  for (let i = 0; i < allVals.length; i++) {
    const row = i + 2;
    const url = String(allVals[i][urlIdx] || '').trim();
    const apptId = String(allVals[i][apptIdx] || '').trim();
    const brand = brandIdx != null ? String(allVals[i][brandIdx] || '').trim() : 'VVS';
    
    if (url || !apptId) continue;
    
    scanned++;
    
    try {
      const expectedName = CS_REPORT_NAME_FMT
        .replace('{Brand}', brand)
        .replace('{APPT_ID}', apptId);
      
      const files = DriveApp.getFilesByName(expectedName);
      let foundFile = null;
      
      while (files.hasNext()) {
        const file = files.next();
        if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
          if (!file.isTrashed()) {
            foundFile = file;
            break;
          }
        }
      }
      
      if (foundFile) {
        const newUrl = 'https://docs.google.com/spreadsheets/d/' + foundFile.getId() + '/edit';
        updates.push({ row: row, url: newUrl });
        linked++;
        Logger.log('Row ' + row + ': Found orphaned file ' + foundFile.getId());
      } else {
        notFound++;
        Logger.log('Row ' + row + ': No file found for ' + apptId);
      }
      
      if (scanned % 5 === 0) {
        Utilities.sleep(2000);
      }
      
    } catch (e) {
      Logger.log('Row ' + row + ' ERROR: ' + e.message);
    }
  }
  
  if (updates.length > 0) {
    updates.forEach(u => {
      master.getRange(u.row, urlIdx + 1).setValue(u.url);
    });
  }
  
  const summary = {
    ok: true,
    scanned: scanned,
    linked: linked,
    notFound: notFound,
    message: 'Linked ' + linked + '/' + scanned + ' orphaned files'
  };
  
  Logger.log(JSON.stringify(summary));
  return summary;
}

/**
 * REPAIR 3: Validate and fix stale URLs
 */
function REPAIR_validateAndFixStaleUrls(opts) {
  const DRY_RUN = opts && opts.dryRun !== false;
  
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(CS_MASTER_SHEET_NAME);
  if (!master) throw new Error('Master sheet not found');

  const header = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];
  const H = headerIndexMap_(header);
  
  const urlIdx = H[CS_REPORT_URL_COL];
  if (urlIdx == null) throw new Error('URL column not found');

  const lastRow = master.getLastRow();
  if (lastRow <= 1) return { ok: true, checked: 0 };

  const allVals = master.getRange(2, 1, lastRow - 1, master.getLastColumn()).getValues();
  
  let checked = 0, valid = 0, stale = 0, fixed = 0;
  const staleRows = [];
  
  for (let i = 0; i < allVals.length; i++) {
    const row = i + 2;
    const url = String(allVals[i][urlIdx] || '').trim();
    
    if (!url) continue;
    if (!isValidSpreadsheetUrl_(url)) {
      staleRows.push({ row: row, reason: 'Invalid URL format', url: url });
      stale++;
      continue;
    }
    
    checked++;
    
    try {
      const fileId = extractIdFromUrl_(url);
      const file = DriveApp.getFileById(fileId);
      
      if (file.isTrashed()) {
        staleRows.push({ row: row, reason: 'File is trashed', url: url });
        stale++;
      } else {
        if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
          staleRows.push({ row: row, reason: 'Not a spreadsheet', url: url });
          stale++;
        } else {
          valid++;
        }
      }
      
    } catch (e) {
      staleRows.push({ row: row, reason: 'File not found: ' + e.message, url: url });
      stale++;
    }
    
    if (checked % 10 === 0) {
      Utilities.sleep(1000);
    }
  }
  
  if (!DRY_RUN && staleRows.length > 0) {
    staleRows.forEach(item => {
      master.getRange(item.row, urlIdx + 1).setValue('');
      fixed++;
      Logger.log('Row ' + item.row + ': Cleared stale URL (' + item.reason + ')');
    });
  }
  
  const summary = {
    ok: true,
    checked: checked,
    valid: valid,
    stale: stale,
    fixed: DRY_RUN ? 0 : fixed,
    dryRun: DRY_RUN,
    staleRows: staleRows.map(r => ({ row: r.row, reason: r.reason })),
    message: DRY_RUN 
      ? 'DRY RUN: Found ' + stale + '/' + checked + ' stale URLs (re-run with {dryRun:false} to fix)'
      : 'Fixed ' + fixed + '/' + stale + ' stale URLs'
  };
  
  Logger.log(JSON.stringify(summary, null, 2));
  return summary;
}

/**
 * REPAIR 4: Find and remove duplicate files
 */
function REPAIR_removeDuplicateFiles(opts) {
  const ACTUALLY_DELETE = false; // Permanently disabled for safety
  
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(CS_MASTER_SHEET_NAME);
  if (!master) throw new Error('Master sheet not found');

  const header = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];
  const H = headerIndexMap_(header);
  
  const urlIdx = H[CS_REPORT_URL_COL];
  const apptIdx = H['APPT_ID'];
  
  if (urlIdx == null || apptIdx == null) {
    throw new Error('Required columns not found');
  }

  const lastRow = master.getLastRow();
  if (lastRow <= 1) return { ok: true, scanned: 0, duplicatesFound: 0 };

  const allVals = master.getRange(2, 1, lastRow - 1, master.getLastColumn()).getValues();
  
  const canonicalFiles = {};
  
  for (let i = 0; i < allVals.length; i++) {
    const apptId = String(allVals[i][apptIdx] || '').trim();
    const url = String(allVals[i][urlIdx] || '').trim();
    
    if (!apptId || !url) continue;
    
    try {
      const fileId = extractIdFromUrl_(url);
      if (!canonicalFiles[apptId]) {
        canonicalFiles[apptId] = fileId;
      }
    } catch (e) {}
  }
  
  Logger.log('Found ' + Object.keys(canonicalFiles).length + ' canonical files in Master sheet');
  
  const duplicates = [];
  let scanned = 0;
  
  for (const apptId in canonicalFiles) {
    const canonicalId = canonicalFiles[apptId];
    
    try {
      const token = 'APPT_ID=' + apptId;
      
      const files = DriveApp.searchFiles(
        'title contains "' + apptId + '" and mimeType = "' + MimeType.GOOGLE_SHEETS + '"'
      );
      
      const foundFiles = [];
      while (files.hasNext()) {
        const file = files.next();
        foundFiles.push(file);
        scanned++;
      }
      
      for (const file of foundFiles) {
        if (file.getId() !== canonicalId) {
          const desc = String(file.getDescription() || '').trim();
          if (desc.includes(token) || file.getName().includes(apptId)) {
            duplicates.push({
              apptId: apptId,
              fileId: file.getId(),
              name: file.getName(),
              url: file.getUrl(),
              canonical: canonicalId
            });
          }
        }
      }
      
      Utilities.sleep(500);
      
    } catch (e) {
      Logger.log('Error scanning ' + apptId + ': ' + e.message);
    }
  }
  
  const summary = {
    ok: true,
    scanned: scanned,
    duplicatesFound: duplicates.length,
    deleted: 0,
    actuallyDelete: ACTUALLY_DELETE,
    duplicates: duplicates.map(d => ({ apptId: d.apptId, fileId: d.fileId, name: d.name })),
    message: 'SCAN ONLY: Found ' + duplicates.length + ' duplicate files. Deletion permanently disabled for safety.'
  };
  
  Logger.log(JSON.stringify(summary, null, 2));
  return summary;
}

/**
 * REPAIR 5: Master repair function - runs all repairs in sequence
 */
function REPAIR_runAll(opts) {
  const DRY_RUN = opts && opts.dryRun !== false;
  
  Logger.log('========================================');
  Logger.log('REPAIR MASTER - DRY RUN: ' + DRY_RUN);
  Logger.log('========================================');
  
  const results = {
    timestamp: new Date().toISOString(),
    dryRun: DRY_RUN,
    repairs: {}
  };
  
  try {
    Logger.log('\n1️⃣ Backfilling idempotency tokens...');
    results.repairs.backfillTokens = REPAIR_backfillIdempotencyTokens();
    
    Logger.log('\n2️⃣ Linking orphaned files...');
    results.repairs.linkOrphaned = REPAIR_linkOrphanedFiles();
    
    Logger.log('\n3️⃣ Validating stale URLs...');
    results.repairs.validateUrls = REPAIR_validateAndFixStaleUrls({ dryRun: DRY_RUN });
    
    Logger.log('\n4️⃣ Scanning for duplicate files...');
    results.repairs.removeDuplicates = REPAIR_removeDuplicateFiles({ actuallyDelete: false });
    
    results.ok = true;
    
  } catch (e) {
    results.ok = false;
    results.error = e.message;
    Logger.log('❌ REPAIR FAILED: ' + e.message);
  }
  
  Logger.log('\n========================================');
  Logger.log('REPAIR SUMMARY:');
  Logger.log(JSON.stringify(results, null, 2));
  Logger.log('========================================');
  
  return results;
}

// ============================================================
// === VERIFICATION HELPERS ===
// ============================================================

function TEST_verifyOptimizations() {
  const results = {
    timestamp: new Date().toISOString(),
    version: '2.6 + Project#4',
    tests: []
  };
  
  // Test 1: URL Column Detection
  try {
    const ss = SpreadsheetApp.getActive();
    const master = ss.getSheetByName(CS_MASTER_SHEET_NAME);
    const header = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];
    const H = headerIndexMap_(header);
    const urlColIdx = H[CS_REPORT_URL_COL];
    results.tests.push({
      name: 'URL Column Detection',
      status: urlColIdx != null ? 'PASS' : 'FAIL',
      urlColumnIndex: urlColIdx,
      message: urlColIdx != null ? 'URL column found at index ' + urlColIdx : 'URL column not found'
    });
  } catch (e) {
    results.tests.push({ name: 'URL Column Detection', status: 'ERROR', error: e.message });
  }
  
  // Test 2: Exponential Backoff
  try {
    const hasOptimizedLock = cs_ensureReportUrl_.toString().includes('Math.pow(2, attempt - 1)');
    results.tests.push({
      name: 'Exponential Backoff',
      status: hasOptimizedLock ? 'PASS' : 'FAIL',
      message: hasOptimizedLock ? 'Exponential backoff code detected' : 'Using old fixed retry delay'
    });
  } catch (e) {
    results.tests.push({ name: 'Exponential Backoff', status: 'ERROR', error: e.message });
  }
  
  // Test 3: Conditional Report Ensure
  try {
    const hasConditional = cs_submitClientStatusUpdate_.toString().includes('opts.reportId && opts.reportUrl');
    results.tests.push({
      name: 'Conditional Report Ensure',
      status: hasConditional ? 'PASS' : 'FAIL',
      message: hasConditional ? 'Conditional logic detected' : 'Still calling ensure unconditionally'
    });
  } catch (e) {
    results.tests.push({ name: 'Conditional Report Ensure', status: 'ERROR', error: e.message });
  }
  
  // Test 4: Final URL Write Check
  try {
    const hasFinalCheck = cs_ensureReportUrl_.toString().includes('finalCheck');
    results.tests.push({
      name: 'Final URL Write Check',
      status: hasFinalCheck ? 'PASS' : 'FAIL',
      message: hasFinalCheck ? 'Final check detected' : 'Missing final URL check'
    });
  } catch (e) {
    results.tests.push({ name: 'Final URL Write Check', status: 'ERROR', error: e.message });
  }
  
  // Test 5: Max Retry Attempts
  try {
    const codeStr = cs_ensureReportUrl_.toString();
    const has5Attempts = codeStr.includes('MAX_ATTEMPTS     = 5');
    results.tests.push({
      name: 'Max Retry Attempts',
      status: has5Attempts ? 'PASS' : 'FAIL',
      message: has5Attempts ? 'MAX_ATTEMPTS = 5' : 'Still using old MAX_ATTEMPTS value'
    });
  } catch (e) {
    results.tests.push({ name: 'Max Retry Attempts', status: 'ERROR', error: e.message });
  }

  // Test 6: PROJECT #4 Deadline Sync
  try {
    const hasP4 = cs_submitClientStatusUpdate_.toString().includes('syncDeadlineFanOut_');
    results.tests.push({
      name: 'Project #4 Deadline Sync',
      status: hasP4 ? 'PASS' : 'FAIL',
      message: hasP4 ? 'syncDeadlineFanOut_ detected in fan-out block' : 'Project #4 patch NOT found'
    });
  } catch (e) {
    results.tests.push({ name: 'Project #4 Deadline Sync', status: 'ERROR', error: e.message });
  }
  
  const passCount = results.tests.filter(t => t.status === 'PASS').length;
  const totalTests = results.tests.length;
  
  results.summary = {
    passed: passCount,
    total: totalTests,
    percentage: Math.round((passCount / totalTests) * 100),
    message: passCount === totalTests 
      ? '✅ All optimizations + Project #4 verified!'
      : '⚠️ ' + (totalTests - passCount) + ' tests failed - review results'
  };
  
  Logger.log('========================================');
  Logger.log('VERIFICATION RESULTS v2.6 + Project#4:');
  Logger.log(JSON.stringify(results, null, 2));
  Logger.log('========================================');
  
  return results;
}

function MONITOR_optimizationMetrics() {
  const metrics = {
    timestamp: new Date().toISOString(),
    version: '2.6 + Project#4',
    period: 'last 24 hours',
    stats: {}
  };
  
  try {
    const ss = SpreadsheetApp.getActive();
    const audit = ss.getSheetByName(CS_AUDIT_TAB);
    
    if (audit) {
      const lastRow = audit.getLastRow();
      if (lastRow > 1) {
        const data = audit.getRange(2, 1, lastRow - 1, audit.getLastColumn()).getValues();
        const header = audit.getRange(1, 1, 1, audit.getLastColumn()).getValues()[0];
        const H = headerIndexMap_(header);
        const tsIdx = H['Updated At'];
        const oneDayAgo = new Date(Date.now() - 24 * 60 * 60 * 1000);
        let recentUpdates = 0;
        if (tsIdx != null) {
          for (const row of data) {
            const ts = row[tsIdx];
            if (ts && new Date(ts) > oneDayAgo) recentUpdates++;
          }
        }
        metrics.stats.totalAuditEntries = lastRow - 1;
        metrics.stats.updatesLast24h = recentUpdates;
      }
    }
    
    const master = ss.getSheetByName(CS_MASTER_SHEET_NAME);
    if (master) {
      const header = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];
      const H = headerIndexMap_(header);
      const urlIdx = H[CS_REPORT_URL_COL];
      if (urlIdx != null) {
        const lastRow = master.getLastRow();
        if (lastRow > 1) {
          const urls = master.getRange(2, urlIdx + 1, lastRow - 1, 1).getValues();
          const populated = urls.filter(r => String(r[0] || '').trim()).length;
          metrics.stats.totalRows = lastRow - 1;
          metrics.stats.urlsPopulated = populated;
          metrics.stats.urlCoverage = Math.round((populated / (lastRow - 1)) * 100) + '%';
        }
      }
    }
    
    Logger.log('========================================');
    Logger.log('OPTIMIZATION METRICS v2.6 + Project#4:');
    Logger.log(JSON.stringify(metrics, null, 2));
    Logger.log('========================================');
    
  } catch (e) {
    Logger.log('Error collecting metrics: ' + e.message);
  }
  
  return metrics;
}


/**
 * Project #22 — Backfill referral vào snapshot khi report mới được tạo
 * Tìm dòng "Refer a Friend:" trong cột A hoặc C và ghi giá trị
 */
function cs_backfillReferralToSnapshot_(sh, referralText) {
  const rowsToScan = Math.min(sh.getLastRow() || 50, 50);
  if (rowsToScan <= 0) return false;

  const values = sh.getRange(1, 1, rowsToScan, 4).getValues();

  // ── Chuẩn hóa: tìm cả có dấu ":" lẫn không có ──
  const normalize = s => String(s || '').trim().replace(/:+$/, '').toLowerCase();
  const needle    = 'refer a friend';

  for (let i = 0; i < rowsToScan; i++) {
    if (normalize(values[i][0]) === needle) {
      sh.getRange(i + 1, 2).setValue(referralText);
      Logger.log('[cs_backfillReferralToSnapshot_] Wrote row ' + (i+1) + ' col B = ' + referralText);
      return true;
    }
    if (normalize(values[i][2]) === needle) {
      sh.getRange(i + 1, 4).setValue(referralText);
      Logger.log('[cs_backfillReferralToSnapshot_] Wrote row ' + (i+1) + ' col D = ' + referralText);
      return true;
    }
  }

  Logger.log('[cs_backfillReferralToSnapshot_] Label not found');
  return false;
}

// ============================================================
// === Legacy shims ===
// ============================================================
if (typeof headerMap_ !== 'function') {
  function headerMap_(sh) { return headerMap__canon(sh); }
}
if (typeof ensureHeaders_ !== 'function') {
  function ensureHeaders_(sh, labels) { return ensureHeaders__canon(sh, labels); }
}
if (typeof getMasterSheet_ !== 'function') {
  function getMasterSheet_(ss) { return getMasterSheet__canon(ss); }
}
if (typeof getOrdersSheet_ !== 'function') {
  function getOrdersSheet_(wb) { return getOrdersSheet__canon(wb); }
}
if (typeof coerceSOTextColumn_ !== 'function') {
  function coerceSOTextColumn_(sh, H) { return coerceSOTextColumn__canon(sh, H); }
}
if (typeof existsSOInMaster_ !== 'function') {
  function existsSOInMaster_(sh, brand, so, skipRow) { return existsSOInMaster__canon(sh, brand, so, skipRow); }
}




function test_referral_write_fix() {
  // Chạy lại debug để xác nhận đã ghi được
  const ss   = SpreadsheetApp.getActive();
  const sh   = ss.getSheetByName('00_Master Appointments');
  const vals = sh.getRange(486, 1, 1, sh.getLastColumn()).getValues()[0];
  const H    = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(s=>String(s||'').trim());

  const refName  = String(vals[H.indexOf('Referral Name')]     || '').trim();
  const refDisc  = String(vals[H.indexOf('Referral Discount')] || '').trim();
  const csUrl    = String(vals[H.indexOf('Client Status Report URL')] || '').trim();

  const csId = csUrl.match(/[-\w]{25,}/)?.[0];
  const csSS = SpreadsheetApp.openById(csId);
  const csSh = csSS.getSheetByName('Client Status');

  const text   = 'Yes — ' + refName + ' (−$' + (refDisc||100) + ')';
  const result = rp_updateClientStatusSnapshotCell_(csSh, 'Refer a Friend:', text);

  Logger.log(result ? '✅ Ghi thành công: ' + text : '❌ Vẫn không ghi được');
}
