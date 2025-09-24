/***** Client Status — v2.2 (mirror to 301/302 removed) *****/

/** CHANGELOG (v2.2, patched)
 * - REMOVED: All 301/302 mirroring and related helpers.
 * - UNCHANGED: Central audit, per-client report creation/log, snapshot updates,
 *              “Updated By/At” metadata on master, DV hooks, and Reminders hooks.
 */


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
const COL_IN_PRODUCTION_HEX = 'IPS - Hex Code'; // NEW
const COL_CENTER_STONE_HEX  = 'CSOS - Hex Code';

// === TEMPLATE CONFIG ===
function getTemplateId_() {
  return PropertiesService.getScriptProperties().getProperty('CS_REPORT_TEMPLATE_ID') || '';
}

// === Helpers ===
function headerIndexMap_(headerRow){ const map={}; headerRow.forEach((h,i)=>{ if (h) map[String(h).trim()]=i; }); return map; }
/** Case-insensitive header finder by regex; returns zero-based index or -1. */
function findHeaderIndexByRegex_(headerRow, regex){
  for (var i = 0; i < headerRow.length; i++){
    if (regex.test(String(headerRow[i] || ''))) return i;
  }
  return -1;
}

function extractIdFromUrl_(url){ const m=String(url).match(/[-\w]{25,}/); return m?m[0]:''; }
function getByAny_(H, vals, names){ for (const n of names){ if (H[n]!=null) return vals[H[n]] ?? ''; } return ''; }

function normalizeMultiArray_(v){
  if (Array.isArray(v)) return v.map(s=>String(s||'').trim()).filter(Boolean);
  return String(v||'')
    .split(/[,;|/]|(?:\s*&\s*)/g)
    .map(s=>s.trim())
    .filter(Boolean);
}
function joinMulti_(arr){
  const a = normalizeMultiArray_(arr);
  const seen = new Set(); const out=[];
  a.forEach(x=>{ if(!seen.has(x)){ seen.add(x); out.push(x); } });
  return out.join(', ');
}

// === Read lists + hex maps from "Dropdown" with ONE data fetch ===
function readDropdowns_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Dropdown');
  if (!sh) throw new Error('Missing tab "Dropdown".');

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h||'').trim());
  const data   = lastRow > 1 ? sh.getRange(2,1,lastRow-1,lastCol).getValues() : [];
  const idx = (name) => header.indexOf(String(name).trim());
  const colVals = (name) => {
    const c = idx(name); if (c < 0 || data.length === 0) return [];
    const col = new Array(data.length);
    for (let i=0;i<data.length;i++) col[i] = String(data[i][c]||'').trim();
    return col;
  };

  // Value columns
  const assignedReps         = colVals('Assigned Rep').filter(Boolean);
  const assistedReps         = colVals('Assisted Rep').filter(Boolean);
  const salesStages          = colVals('Sales Stage').filter(Boolean);           // <-- plural
  const convStatuses         = colVals('Conversion Status').filter(Boolean);
  const customOrderStatuses  = colVals('Custom Order Status').filter(Boolean);
  const centerStoneStatuses  = colVals('Center Stone Order Status').filter(Boolean);
  const inProductionStatuses = colVals('In Production Status').filter(Boolean);

  // Hex columns aligned row-for-row
  const ssHex   = colVals(COL_SALES_STAGE_HEX);
  const csHex   = colVals(COL_CONV_STATUS_HEX);
  const cosHex  = colVals(COL_CUST_ORDER_HEX);
  const csosHex = colVals(COL_CENTER_STONE_HEX);
  const ipsHex  = colVals(COL_IN_PRODUCTION_HEX);

  const buildHexMap = (values, hexes) => {
    const map = {};
    const n = Math.min(values.length, hexes.length);
    for (let i=0; i<n; i++){
      const v = String(values[i]||'').trim();
      const h = String((hexes[i]||'').replace('#','').trim());
      if (!v) continue;
      if (/^[0-9A-Fa-f]{6}$/.test(h)) map[v] = '#'+h.toUpperCase();
    }
    return map;
  };

  return {
    assignedReps, assistedReps, salesStages, convStatuses, customOrderStatuses, centerStoneStatuses, inProductionStatuses,
    colors: {
      salesStage:   buildHexMap(salesStages,          ssHex),  // key remains 'salesStage' for HTML chip color lookups
      convStatus:   buildHexMap(convStatuses,         csHex),
      customOrder:  buildHexMap(customOrderStatuses,  cosHex),
      centerStone:  buildHexMap(centerStoneStatuses,  csosHex),
      inProduction: buildHexMap(inProductionStatuses, ipsHex)
    }
  };
}

/** Read "Validation Rules (Flattened Matrix)" and "Viewing Rules" from the Dropdown tab. */
function readValidationRulesFlat_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Dropdown Rules');
  if (!sh) throw new Error('Missing tab "Dropdown Rules".');

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { matrix: [], viewing: [] };

  const all = sh.getRange(1,1,lastRow,lastCol).getValues();

  // ---- Find header row for the flattened matrix
  let hdrRow = -1;
  let col = { sales:-1, conv:-1, cos:-1, ips:-1, csReq:-1, dead:-1, notes:-1 };

  for (let r = 0; r < all.length; r++) {
    const row = all[r].map(x => String(x||'').trim());
    const iSales = row.indexOf('Sales Stage');
    const iConv  = row.indexOf('Conversion Status');
    const iCOS   = row.indexOf('Custom Order Status');
    if (iSales >= 0 && iConv >= 0 && iCOS >= 0) {
      hdrRow = r;
      col.sales = iSales;
      col.conv  = iConv;
      col.cos   = iCOS;
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
      const csr = col.csReq>= 0 ? String(row[col.csReq]|| '').trim() : '';
      const dr  = col.dead >= 0 ? String(row[col.dead] || '').trim() : '';
      const nt  = col.notes>= 0 ? String(row[col.notes]|| '').trim() : '';
      if (s || c || cos || ips || csr || dr || nt) {
        matrix.push({
          salesStage: s, convStatus: c, customOrderStatus: cos,
          ipsRequirement: ips, centerStoneRequirement: csr,
          deadlineRequirement: dr, notes: nt
        });
      }
    }
  }

  // ---- Find header row for Viewing Rules
  let vHdr = -1, cDays = -1, cMin = -1;
  for (let r = 0; r < all.length; r++) {
    const row = all[r].map(x => String(x||'').trim());
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




/** Open the Client Status dialog (2-column; popover chip pickers on right) */
function cs_openStatusDialog_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CS_MASTER_SHEET_NAME);
  const r = sh.getActiveRange();
  if (!r || r.getRow() === 1) {
    SpreadsheetApp.getUi().alert('⚠️ Select a data row in 00_Master Appointments first.');
    return;
  }
  const row = r.getRow();

  const header = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
  const H = headerIndexMap_(header);
  const vals = sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0];

  const get = name => H[name] != null ? vals[H[name]] : '';

  const assignedRepStr = String(get('Assigned Rep') || '');
  const assistedRepStr = String(get('Assisted Rep') || '');

const orderDateISO = toISODateForInput_(get('Order Date'));

  const prefill = {
    clientName:  String(get('Customer Name') || ''),
    apptId:      String(get('APPT_ID') || ''),
    assignedRep: assignedRepStr,
    assistedRep: assistedRepStr,
    assignedRepArr: normalizeMultiArray_(assignedRepStr),
    assistedRepArr: normalizeMultiArray_(assistedRepStr),
    salesStage:  String(get('Sales Stage') || ''),
    convStatus:  String(get('Conversion Status') || ''),
    customOrder: String(get('Custom Order Status') || ''),
    inProduction: String(get('In Production Status') || ''), // NEW
    centerStone: String(get('Center Stone Order Status') || ''),
    nextSteps:   String(get('Next Steps') || ''),
    orderDate:   orderDateISO
  };

  const lists = readDropdowns_(); // lists + colors

  // NEW: Read flattened prevention rules from Dropdown + compute Visit ISO
  const rulesFlat = readValidationRulesFlat_();

  let visitISO = String(get('ApptDateTime (ISO)') || '').trim();
  if (!visitISO) {
    const vdate = String(get('Visit Date') || '').trim();
    const vtime = String(get('Visit Time') || '').trim();
    if (vdate || vtime) {
      try { visitISO = Utilities.formatDate(new Date(vdate + ' ' + vtime), CS_TZ, "yyyy-MM-dd'T'HH:mm:ssXXX"); } catch(_){}
    }
  }

  const t = HtmlService.createTemplateFromFile('dlg_client_status_v1');
  t.prefill = prefill;
  t.lists = {
    assignedReps:         lists.assignedReps,
    assistedReps:         lists.assistedReps,
    salesStages:          lists.salesStages,          // <-- plural to match HTML
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

/** Submit from dialog (arrays for reps; statuses required) */
function cs_submitFromDialog(payload) {
// ---- Conditional server guard mirroring client logic ----
  function _centerStoneRequired(stage, conv) {
    if (/^Lost Lead/i.test(String(stage||''))) return false;
    if (/^Viewing Scheduled$/i.test(String(conv||''))) return true;
    if (/^(Deposit Paid|Confirmed Order|Order In Progress)$/i.test(String(conv||''))) return true;
    return false;
  }

  // Sales Stage & Conversion Status always required
  ['salesStage','convStatus'].forEach(function (k) {
    if (!String(payload[k] || '').trim()) {
      throw new Error('Please complete: Sales Stage and Conversion Status before submitting.');
    }
  });

  // Custom Order Status required unless rules yielded zero options
  var cosEmptyAllowed = !!payload.cosAllowedEmpty;
  if (!cosEmptyAllowed && !String(payload.customOrder || '').trim()) {
    throw new Error('Please select a Custom Order Status.');
  }

  var isInProduction = String(payload.customOrder || '') === 'In Production';
  if (isInProduction && !String(payload.inProduction || '').trim()) {
    throw new Error('Please select an "In Production Status" since Custom Order Status is In Production.');
  }

  var need3D = /^(3D Requested|3D Revision Requested)$/i.test(String(payload.customOrder||''));
  if (need3D && !String((payload.deadline3d||'')).trim()) {
    throw new Error('3D Deadline is required when Custom Order Status is 3D Requested or 3D Revision Requested.');
  }
  if (isInProduction && !String((payload.prodDeadline||'')).trim()) {
    throw new Error('Production Deadline is required when Custom Order Status is In Production.');
  }

  // Order Date is required for these COS values
  var needOrderDate = /^(Approved for Production|Waiting Production Timeline|In Production|Final Photos\s*[–-]\s*Waiting Approval|Warehouse|Ship to US|In US Store|Ship to Customer|Order Completed)$/i
    .test(String(payload.customOrder||''));
  if (needOrderDate && !String(payload.orderDate || '').trim()) {
    throw new Error('Order Date is required for the selected Custom Order Status.');
  }

  if (_centerStoneRequired(String(payload.salesStage||''), String(payload.convStatus||'')) &&
      !String(payload.centerStone || '').trim()) {
    throw new Error('Center Stone Order Status is required for Viewing Scheduled or Deposit/Confirmed/Order In Progress.');
  }

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CS_MASTER_SHEET_NAME);
  const r = sh.getActiveRange();
  if (!r || r.getNumRows() !== 1 || r.getRow() === 1) throw new Error('Select exactly one row.');
  const row = r.getRow();

  const header = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
  const H = headerIndexMap_(header);
  const vals = sh.getRange(row,1,1, sh.getLastColumn()).getValues()[0];
  // Snapshot previous Center Stone status BEFORE we write any changes
  const __prevCenterStone = String(vals[H['Center Stone Order Status']] ?? '').trim();

  // Normalize multi arrays → single stored strings (back-compatible)
  const assignedJoined = joinMulti_(payload.assignedRep);
  const assistedJoined = joinMulti_(payload.assistedRep);

  const setIf = (name, value) => { if (value != null && String(value).trim() !== '' && H[name] != null) { vals[H[name]] = value; } };

  setIf('Assigned Rep',              assignedJoined);
  setIf('Assisted Rep',              assistedJoined);
  setIf('Sales Stage',               payload.salesStage);
  setIf('Conversion Status',         payload.convStatus);
  setIf('Custom Order Status',       payload.customOrder);
  setIf('Order Date', payload.orderDate); 

  // NEW: write/clear In Production Status (robust header lookup)
  var ipsIdx = (H['In Production Status'] != null)
      ? H['In Production Status']
      : findHeaderIndexByRegex_(header, /in\s*production\s*status/i);

  if (ipsIdx >= 0) {
    vals[ipsIdx] = isInProduction ? String(payload.inProduction || '').trim() : '';
  }

  // ---- NEW: deadlines write + move counters + log meta ----
  /** @type {Object.<string,number>} */
  const H2 = H; // alias for clarity

  // Robust column lookups (handles slight header variations)
  const idxProdDeadline = (H2['Production Deadline'] != null)
    ? H2['Production Deadline']
    : findHeaderIndexByRegex_(header, /(Production|Prod\.)\s*Deadline/i);

  const idx3dDeadline = (H2['3D Deadline'] != null)
    ? H2['3D Deadline']
    : findHeaderIndexByRegex_(header, /3D\s*Deadline/i);

  const idxProdMoves = (H2['# of Times Prod. Deadline Moved'] != null)
    ? H2['# of Times Prod. Deadline Moved']
    : findHeaderIndexByRegex_(header, /#\s*of\s*Times\s*(Prod|Production).*Deadline.*Moved/i);

  const idx3dMoves = (H2['# of Times 3D Deadline Moved'] != null)
    ? H2['# of Times 3D Deadline Moved']
    : findHeaderIndexByRegex_(header, /#\s*of\s*Times\s*3D.*Deadline.*Moved/i);

  // Capture current (pre‑update) values
  const prevProdDeadline = idxProdDeadline >= 0 ? String(vals[idxProdDeadline] || '').trim() : '';
  const prev3dDeadline   = idx3dDeadline   >= 0 ? String(vals[idx3dDeadline]   || '').trim() : '';
  const prevProdMovesStr = idxProdMoves    >= 0 ? String(vals[idxProdMoves]    || '').trim() : '';
  const prev3dMovesStr   = idx3dMoves      >= 0 ? String(vals[idx3dMoves]      || '').trim() : '';

  // Determine applicability and new input
  const is3D = /^(3D Requested|3D Revision Requested)$/i.test(String(payload.customOrder||''));
  const newProdDeadline = isInProduction ? String(payload.prodDeadline || '') : '';
  const new3dDeadline   = is3D          ? String(payload.deadline3d   || '') : '';

  // Write deadlines to Master row (or clear when not applicable)
  if (idxProdDeadline >= 0) vals[idxProdDeadline] = newProdDeadline;
  if (idx3dDeadline   >= 0) vals[idx3dDeadline]   = new3dDeadline;

  // Move counters logic
  let prodChanged = false, threeDChanged = false;

  // Production counter
  if (idxProdDeadline >= 0 && isInProduction) {
    if (!prevProdDeadline && newProdDeadline) {
      // first set → dash
      if (idxProdMoves >= 0) vals[idxProdMoves] = '-';
    } else if (prevProdDeadline && newProdDeadline && prevProdDeadline !== newProdDeadline) {
      prodChanged = true;
      const prevN = (prevProdMovesStr === '-' || prevProdMovesStr === '') ? 0 : (parseInt(prevProdMovesStr, 10) || 0);
      if (idxProdMoves >= 0) vals[idxProdMoves] = String(prevN + 1);
    }
  }

  // 3D counter
  if (idx3dDeadline >= 0 && is3D) {
    if (!prev3dDeadline && new3dDeadline) {
      // first set → dash
      if (idx3dMoves >= 0) vals[idx3dMoves] = '-';
    } else if (prev3dDeadline && new3dDeadline && prev3dDeadline !== new3dDeadline) {
      threeDChanged = true;
      const prevN = (prev3dMovesStr === '-' || prev3dMovesStr === '') ? 0 : (parseInt(prev3dMovesStr, 10) || 0);
      if (idx3dMoves >= 0) vals[idx3dMoves] = String(prevN + 1);
    }
  }

  // Build log meta for the Client Status Report (what changed this submit)
  let logDeadlineType = '', logDeadlineDate = '', logMoveCount = '';
  if (idxProdDeadline >= 0 && isInProduction && ( (!prevProdDeadline && newProdDeadline) || prodChanged )) {
    logDeadlineType = 'Production';
    logDeadlineDate = newProdDeadline;
    logMoveCount    = (idxProdMoves >= 0 ? String(vals[idxProdMoves] || '') : '');
  }
  if (idx3dDeadline >= 0 && is3D && ( (!prev3dDeadline && new3dDeadline) || threeDChanged )) {
    logDeadlineType = logDeadlineType ? (logDeadlineType + ' | 3D') : '3D';
    logDeadlineDate = logDeadlineDate ? (logDeadlineDate + ' | ' + new3dDeadline) : new3dDeadline;
    const mc = (idx3dMoves >= 0 ? String(vals[idx3dMoves] || '') : '');
    logMoveCount = logMoveCount ? (logMoveCount + ' | ' + mc) : mc;
  }

  // If COS is Final Photos – Waiting Approval or any later shipping step → IPS must be Production Completed
  (function enforceIPSForLaterPhases(){
    const cosNow = String(payload.customOrder || '').trim();
    const later = new Set([
      'Final Photos – Waiting Approval',
      'Warehouse',
      'Ship to US',
      'In US Store',
      'Ship to Customer',
      'Order Completed'
    ]);
    if (later.has(cosNow) && typeof ipsIdx === 'number' && ipsIdx >= 0) {
      vals[ipsIdx] = 'Production Completed';
    }
  })();

  // Stash the log meta to forward to the next step (success screen + CSR log)
  payload.__deadlineLog = { type: logDeadlineType, date: logDeadlineDate, moves: logMoveCount };


  setIf('Center Stone Order Status', payload.centerStone);
  if (H['Next Steps'] != null && payload.nextSteps != null) vals[H['Next Steps']] = payload.nextSteps;

  // Single write
  sh.getRange(row, 1, 1, vals.length).setValues([vals]);

  // ---- (NEW) Create Wax Request if asked ----
  var waxSummary = null;
  try {
    if (payload.wax && payload.wax.request === true) {
      // Determine RootApptID (or fall back to APPT_ID)
      var rootApptId = String(
        (H['RootApptID'] != null ? vals[H['RootApptID']] : '') ||
        (H['APPT_ID']    != null ? vals[H['APPT_ID']]    : '') ||
        ''
      ).trim();

      if (rootApptId) {
        var wres = wax_onRequestSubmit_({
          rootApptId: rootApptId,
          soMo: (payload.wax.soMo || ''),
          neededByRep: (payload.wax.neededBy || ''),
          priority: (payload.wax.priority || ''),
          requestedBy: (Session.getActiveUser().getEmail() || assignedJoined || '')
        }) || {};
        // Normalized for the HTML success view
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

  // Continue pipeline (audit + client report + snapshot + mirror)
  // Pass Assisted Rep directly (no PropertiesService temp handoff)
  return cs_submitClientStatusUpdate_({
    assistedRep:     assistedJoined,
    prevCenterStone: __prevCenterStone,
    inProduction:    String(payload.inProduction || '').trim(),
    wax:             waxSummary || null,
    waxSummaryStr:   String(payload.waxSummary || ''),

    // NEW: forward dates so the success screen can show them (and no undefined refs)
    prodDeadline: String(payload.prodDeadline || ''),
    deadline3d:   String(payload.deadline3d   || '')
  });
}

// === Create/find per-client report; write audit/log/snapshot ===
function cs_createOrGetReportForSelection_(opts) {
  const inSubmit = !!(opts && opts.inSubmit);

  let lock;
  if (!inSubmit) {
    lock = LockService.getDocumentLock();
    if (!lock.tryLock(1500)) return { ok:false, error:'LOCKED' };
  }
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(CS_MASTER_SHEET_NAME);
    const r = sh.getActiveRange();
    const row = r.getRow();

    const header = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
    const H = headerIndexMap_(header);
    const vals = sh.getRange(row,1,1, sh.getLastColumn()).getValues()[0];
    const get  = n => H[n] != null ? (vals[H[n]] ?? '') : '';

    const apptId = String(get('APPT_ID')).trim();
    const brand  = String(get('Brand')).trim();
    const client = String(get('Customer Name')).trim();
    const name   = CS_REPORT_NAME_FMT.replace('{Brand}', brand || 'VVS').replace('{APPT_ID}', apptId);

    let reportUrl = String(get(CS_REPORT_URL_COL) || '').trim();
    let reportId  = reportUrl ? extractIdFromUrl_(reportUrl) : '';
    let reportSS  = null;

    // Validate and keep the opened handle if it exists
    if (reportId) {
      try {
        reportSS = SpreadsheetApp.openById(reportId);
      } catch (e) {
        reportId = '';
      }
    }

    if (!reportId) {
      const parent = pickParentFolder_(get(CS_PROSPECT_URL_COL), client);
      reportId = createClientReport_(name, parent);
      reportUrl = 'https://docs.google.com/spreadsheets/d/' + reportId + '/edit';
      if (H[CS_REPORT_URL_COL] != null) sh.getRange(row, H[CS_REPORT_URL_COL] + 1).setValue(reportUrl);
      // Open once for return
      reportSS = SpreadsheetApp.openById(reportId);
    }

    return { ok:true, id:reportId, url:reportUrl, ss: reportSS };

  } catch (e) {
    return { ok:false, error: String(e && e.message || e) };
  } finally {
    if (lock) { try { lock.releaseLock(); } catch(_){ } }
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
function createClientReport_(name, parentFolder) {
  const templateId = getTemplateId_();
  if (!templateId) throw new Error('Client Status: CS_REPORT_TEMPLATE_ID not set in Project Properties.');
  const tmplFile = DriveApp.getFileById(templateId);
  const copy = tmplFile.makeCopy(name, parentFolder || DriveApp.getRootFolder());
  const fileId = copy.getId();
  try { if (parentFolder) DriveApp.getRootFolder().removeFile(copy); } catch (e) {}
  return fileId;
}

/**
 * Ensure or refresh the _Config sheet in a Client Status Report.
 * Hybrid mode: create if missing, update only when blank or outdated.
 */
function ensureReportConfig_(reportSS, opts){
  const rootApptId = String(opts.rootApptId||'').trim();
  const reportId   = String(opts.reportId||reportSS.getId()).trim();

  let sh = reportSS.getSheetByName('_Config');
  if (!sh) {
    sh = reportSS.insertSheet('_Config');
    try { sh.hideSheet(); } catch(_){}
    sh.appendRow(['ROOT_APPT_ID', rootApptId]);
    sh.appendRow(['CONTROLLER_URL', ScriptApp.getService().getUrl()]);
    sh.appendRow(['REPORT_REANALYZE_TOKEN',
      PropertiesService.getScriptProperties().getProperty('REPORT_REANALYZE_TOKEN') || ''
    ]);
    sh.appendRow(['REPORT_ID', reportId]);
    return;
  }

  // Read current values into a map
  const vals = sh.getRange(1,1,sh.getLastRow(),2).getValues();
  const map = {};
  vals.forEach(r => { if (r[0]) map[String(r[0]).trim()] = String(r[1]||'').trim(); });

  // Always expected keys
  const want = {
    ROOT_APPT_ID: rootApptId,
    CONTROLLER_URL: ScriptApp.getService().getUrl(),
    REPORT_REANALYZE_TOKEN: PropertiesService.getScriptProperties().getProperty('REPORT_REANALYZE_TOKEN') || '',
    REPORT_ID: reportId
  };

  Object.keys(want).forEach((k,i) => {
    const cur = map[k] || '';
    const need = String(want[k]||'');
    if (cur !== need) {
      // Find existing row or append if missing
      let rowIdx = vals.findIndex(r => String(r[0]).trim() === k);
      if (rowIdx >= 0) {
        sh.getRange(rowIdx+1, 2).setValue(need);
      } else {
        sh.appendRow([k, need]);
      }
    }
  });
}


function cs_submitClientStatusUpdate_(opts) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1500)) return { ok:false, error:'LOCKED' };
  try {
    const ss = SpreadsheetApp.getActive();
    const master = ss.getSheetByName(CS_MASTER_SHEET_NAME);
    const r = master.getActiveRange();
    const row = r.getRow();

    const header = master.getRange(1,1,1, master.getLastColumn()).getValues()[0];
    const H = headerIndexMap_(header);
    const vals = master.getRange(row,1,1, master.getLastColumn()).getValues()[0];
    const get  = n => vals[H[n]] ?? '';

    const apptId      = String(get('APPT_ID') || '').trim();
    const brand       = String(get('Brand') || '');
    const client      = String(get('Customer Name') || '');
    const rep         = String(get('Assigned Rep') || '');
    const salesStage  = String(get('Sales Stage') || '');
    const convStatus  = String(get('Conversion Status') || '');
    const customOrd   = String(get('Custom Order Status') || '');
    const inProduction = String(get('In Production Status') || (opts && opts.inProduction) || ''); // NEW (+fallback)
    const centerStone = String(get('Center Stone Order Status') || '');
    const nextSteps   = String(get('Next Steps') || '');
    const orderDate   = String(get('Order Date') || '');  // NEW

    const phone        = String(getByAny_(H, vals, ['Phone','Client Phone','Customer Phone']) || '');
    const email        = String(getByAny_(H, vals, ['Email','Client Email','Customer Email']) || '');
    const occasion     = String(getByAny_(H, vals, ['Occasion']) || '');
    const budgetRange  = String(getByAny_(H, vals, ['Budget Range']) || '');
    const decisionMkr  = String(getByAny_(H, vals, ['Decision-Maker','Decision Maker']) || '');
    const initialReq   = String(getByAny_(H, vals, ['Initial Request']) || '');
    const soNumber     = String(getByAny_(H, vals, ['SO Number','SO#']) || '').trim();

    const now  = new Date();
    const iso  = Utilities.formatDate(now, CS_TZ, 'yyyy-MM-dd');
    const ts   = Utilities.formatDate(now, CS_TZ, 'yyyy-MM-dd HH:mm:ss');
    const nice = Utilities.formatDate(now, CS_TZ, 'MMM d, yyyy h:mm a z');
    const user  = Session.getActiveUser().getEmail() || rep || 'Unknown';
    const assistedRep = String((opts && opts.assistedRep) || '');


    // 1) Central audit (+Applied To note)
    const audit = ss.getSheetByName(CS_AUDIT_TAB);
    const auditExists = !!audit;

    if (auditExists) {
      // Compute how many rows will be touched by fan-out (active row + siblings sharing RootApptID/APPT_ID)
      const rootKeyForAudit = String(get('RootApptID') || get('APPT_ID') || '').trim();
      let appliedCountTotal = 1; // at least the active row
      if (rootKeyForAudit) {
        const lastRowAll = master.getLastRow();
        if (lastRowAll > 1) {
          const matchColIndexAudit = (H['RootApptID'] != null) ? H['RootApptID']
                                  : (H['APPT_ID']    != null) ? H['APPT_ID']
                                  : -1;
          if (matchColIndexAudit >= 0) {
            const allValsAudit = master.getRange(2, 1, lastRowAll - 1, master.getLastColumn()).getValues();
            for (let i = 0; i < allValsAudit.length; i++) {
              const rnum = i + 2; // data starts at row 2
              if (rnum === row) continue; // skip active row (already counted)
              const idHere = String(allValsAudit[i][matchColIndexAudit] || '').trim();
              if (idHere && idHere === rootKeyForAudit) appliedCountTotal++;
            }
          }
        }
      }

      const appliedNote = `Applied to ${appliedCountTotal} row${appliedCountTotal === 1 ? '' : 's'}`
                          + (rootKeyForAudit ? ` (RootApptID=${rootKeyForAudit})` : '');

      // Ensure audit has an "Applied To" column; add once if missing and refresh header
      let auditHeader = audit.getRange(1,1,1,audit.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
      if (auditHeader.indexOf('Applied To') < 0) {
        audit.getRange(1, audit.getLastColumn() + 1).setValue('Applied To');
        auditHeader = audit.getRange(1,1,1,audit.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
      }

      // Append audit row (header-aware)
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
      // Do not block first‑time report creation on a missing audit tab.
      Logger.log(`Client Status: audit tab "${CS_AUDIT_TAB}" not found — continuing without central audit for this submission.`);
    }

    // 2) Ensure/find client report (robust open + fallback to create on invalid id)
    let reportUrl = String(get(CS_REPORT_URL_COL) || '').trim();
    let reportId  = reportUrl ? extractIdFromUrl_(reportUrl) : '';
    let reportSS  = null;

    if (reportId) {
      try {
        // Guard against stale/incorrect IDs (e.g., a Drive folder ID pasted into the cell)
        reportSS = SpreadsheetApp.openById(reportId);
      } catch (e) {
        reportId = '';
        reportSS = null;
      }
    }
    if (!reportId) {
      const created = cs_createOrGetReportForSelection_({ inSubmit:true });
      if (!created || !created.ok) return { ok:false, error: (created && created.error) || 'Could not create/find client report' };
      reportUrl = created.url; reportId = created.id; reportSS = created.ss;
    }

    // === write _Config into the report for in-file menu relay ===
    const rootApptId = String(get('RootApptID') || get('APPT_ID') || '').trim();
    ensureReportConfig_(reportSS, {
      rootApptId: rootApptId,
      reportId: reportId
    });

    // 3) Per-client log row (header-aware; will place each value in the right column)
    if (CS_WRITE_PER_CLIENT_LOG) {
      insertLogRowByHeader_(reportSS, {
        'Log Date':                  iso,
        'Sales Stage':               salesStage,
        'Conversion Status':         convStatus,
        'Custom Order Status':       customOrd,
        'In Production Status':      inProduction,
        'Center Stone Order Status': centerStone,
        'Next Steps':                nextSteps,

        // NEW — will fill only if those headers exist
        'Deadline Type':             (opts && opts.deadlineLog && opts.deadlineLog.type)  || '',
        'Deadline Date':             (opts && opts.deadlineLog && opts.deadlineLog.date)  || '',
        'Move Count':                (opts && opts.deadlineLog && opts.deadlineLog.moves) || '',

        'Assisted Rep':              assistedRep,
        'Updated By':                user,
        'Updated At':                ts
      });
    }

    // 4) Snapshot
    updateSnapshot_(reportSS, {
      Brand: brand, ClientName: client, APPT_ID: apptId, AssignedRep: rep,
      Phone: phone, Email: email, Occasion: occasion,
      BudgetRange: budgetRange, DecisionMaker: decisionMkr, InitialRequest: initialReq,
      SO_Number: soNumber,
      SalesStage: salesStage, ConversionStatus: convStatus, CustomOrderStatus: customOrd,
      InProductionStatus: inProduction,
      CenterStoneStatus: centerStone, NextSteps: nextSteps, UpdatedBy: user, UpdatedAt: ts,
      AssistedRep: assistedRep,
      OrderDate: orderDate   // NEW → snapshot will place into D2 when label "Order Date:" is in column C
    });

    // 5) Mirror metadata ("Updated By/At") back to master if columns exist (unchanged behavior)

        // 5b) Fan-out the same status updates to ALL rows with the same RootApptID
    try {
      // Resolve the root key we’ll match on (prefer RootApptID; fall back to APPT_ID)
      const rootKey = String(get('RootApptID') || get('APPT_ID') || '').trim();
      if (rootKey) {
        const lastRow = master.getLastRow();
        if (lastRow > 1) {
          // Build a robust index for "In Production Status" (it may be renamed)
          const ipsIdx = (H['In Production Status'] != null)
            ? H['In Production Status']
            : findHeaderIndexByRegex_(header, /in\s*production\s*status/i);

          // Columns we will propagate (only if header exists)
          const colNames = [
            'Assigned Rep',
            'Assisted Rep',
            'Sales Stage',
            'Conversion Status',
            'Custom Order Status',
            'Center Stone Order Status',
            'Next Steps',
            'Updated By',
            'Updated At'
          ];

          // Read once: all master values (rows 2..lastRow)
          const allVals = master.getRange(2, 1, lastRow - 1, master.getLastColumn()).getValues();

          // Identify the column we’ll match against (prefer RootApptID; else APPT_ID)
          const matchColIndex = (H['RootApptID'] != null) ? H['RootApptID']
                                : (H['APPT_ID'] != null) ? H['APPT_ID']
                                : -1;

          if (matchColIndex >= 0) {
            // Build target row numbers (1-based in sheet)
            const targets = [];
            for (let i = 0; i < allVals.length; i++) {
              const rowNum = i + 2; // because we started at row 2
              if (rowNum === row) continue; // skip the active row we already updated
              const idHere = String(allVals[i][matchColIndex] || '').trim();
              if (idHere && idHere === rootKey) targets.push(rowNum);
            }

            if (targets.length) {
              // Prepare per-column batched writes using your existing groupedSetValues_ helper
              const enqueuePairs = (name, value) => {
                const idx = H[name];
                if (idx == null) return null;
                /** @type {{r:number,v:any}[]} */
                const pairs = [];
                for (const rnum of targets) pairs.push({ r: rnum, v: value });
                return { colIdx1: idx + 1, pairs };
              };

              // Core statuses & notes
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

              // In Production Status may be absent/renamed; propagate (including clearing when blank)
              if (ipsIdx >= 0) {
                /** @type {{r:number,v:any}[]} */
                const ipsPairs = [];
                for (const rnum of targets) ipsPairs.push({ r: rnum, v: inProduction });
                groupedSetValues_(master, ipsIdx + 1, ipsPairs);
              }

              // Execute grouped writes for the rest
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

    const uIdx = H['Updated By'], aIdx = H['Updated At'];
    if (uIdx != null && aIdx != null && Math.abs((uIdx+1) - (aIdx+1)) === 1){
      const from = Math.min(uIdx, aIdx) + 1;
      const pairVals = (uIdx < aIdx) ? [[user, ts]] : [[ts, user]];
      master.getRange(row, from, 1, 2).setValues(pairVals);
    } else {
      if (uIdx != null) master.getRange(row, uIdx+1).setValue(user);
      if (aIdx != null) master.getRange(row, aIdx+1).setValue(ts);
    }

    // (Removed) 301/302 mirroring — intentionally disabled per requirements.

    // 6) DV — If Center Stone becomes "Need to Propose …", enqueue +2 calendar days (earlier-wins dedupe)
    try {
      if (typeof DV_init_ === 'function') { DV_init_(); }  // optional init; skip if not defined

      var prevCenterStone = (opts && opts.prevCenterStone) || '';
      var newCenterStone  = centerStone || '';

      var becameNeed = !(typeof DV_isNeedToPropose==='function' ? DV_isNeedToPropose(prevCenterStone) : false)
              &&  (typeof DV_isNeedToPropose==='function' ? DV_isNeedToPropose(newCenterStone)  : false);
      Logger.log('DV hook: prev="' + prevCenterStone + '" → new="' + newCenterStone + '"; becameNeed=' + becameNeed);

      if (becameNeed) {
        if (rootApptId) {
          var res = DV_upsertProposeNudge_afterStatus_({
            rootApptId: rootApptId,
            customerName: client,
            nextStepsFromMaster: nextSteps
          });
          Logger.log('DV hook: queued +2d nudge for root=' + rootApptId + ' → ' + JSON.stringify(res));
        } else {
          Logger.log('DV hook: skipped — no RootApptID/APPT_ID on row');
        }
      }
    } catch (e) {
      Logger.log('DV hook error: ' + (e && e.message ? e.message : e));
    }


    // 7) R1 — Update reminder queue (auto-confirm or ensure follow-up)
    try {
      Remind.onClientStatusChange(soNumber, salesStage, customOrd, user, {
        assignedRepName:  rep,
        assistedRepName:  assistedRep,
        customerName:     client,
        nextSteps:        nextSteps
      });
    } catch (e) {
      console.warn('Remind.onClientStatusChange failed:', e && e.message ? e.message : e);
    }

    const masterLink = ss.getUrl() + '#gid=' + master.getSheetId() + '&range=A' + row;
    const waxObj        = (opts && opts.wax) || null;            // {created, requestId, folderUrl, rowUrl} or null
    const waxSummaryStr = String((opts && opts.waxSummaryStr) || '');

    return {
      ok: true,
      summary: {
        clientName: client, apptId,
        assignedRep: rep, assistedRep,
        salesStage, convStatus,
        customOrder: customOrd,
        deadline3d:   String((opts && opts.deadline3d)   || ''),
        orderDate,
        inProduction,
        prodDeadline: String((opts && opts.prodDeadline) || ''),
        centerStone, nextSteps,
        submittedBy: user, submittedAt: nice,
        reportUrl, masterLink,
        // New fields for the success screen:
        rootApptId: String((H['RootApptID'] != null ? vals[H['RootApptID']] : '') || (H['APPT_ID'] != null ? vals[H['APPT_ID']] : '') || '').trim(),
        waxSummary: waxSummaryStr,   // for the single-line display
        wax:        waxObj           // for the “created / links” block
      }
    };

  } catch (e) {
    return { ok:false, error: String(e && e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

/**
* Find and cache the "Log Date" header row between rows [10..40].
* Verifies cache each time to avoid stale positions.
*/
function getLogHeaderRow_(sh){
  const sp = sh.getParent(); // Spreadsheet
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

/**
 * Insert one log row immediately under the header ("Log Date") using header-name mapping.
 * valuesByName is an object like:
 * {
 *   'Log Date': '2025-09-10', 'Sales Stage': 'Lead', 'Conversion Status': 'Quotation Requested',
 *   'Custom Order Status': '3D Received', 'In Production Status': 'Diamond Memo – NONE APPROVED',
 *   'Center Stone Order Status': 'No Center Stone', 'Next Steps': 'test next steps',
 *   'Assisted Rep': 'vt@cthyp.us', 'Updated By': 'user@domain', 'Updated At': '2025-09-10 12:34:56'
 * }
 */
function insertLogRowByHeader_(reportSS, valuesByName) {
  const sh = reportSS.getSheetByName(CS_REPORT_SHEET);
  if (!sh) throw new Error(`Missing "${CS_REPORT_SHEET}" tab`);

  const headerRow = getLogHeaderRow_(sh);
  const header = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
  const H = {}; header.forEach((h,i)=>{ if (h) H[h] = i; });

  const row = new Array(header.length).fill('');
  Object.keys(valuesByName).forEach(name => {
    const i = H[name];
    if (i != null) row[i] = valuesByName[name];
  });

  sh.insertRowsBefore(headerRow + 1, 1);
  sh.getRange(headerRow + 1, 1, 1, row.length).setValues([row]);
}

function insertLogRow_(reportSS, values9) {
  const sh = reportSS.getSheetByName(CS_REPORT_SHEET);
  if (!sh) throw new Error(`Missing "${CS_REPORT_SHEET}" tab`);

  const headerRow = getLogHeaderRow_(sh);
  const insertAt = headerRow + 1;

  sh.insertRowsBefore(insertAt, 1);
  sh.getRange(insertAt, 1, 1, 9).setValues([values9]);
}

/**
* Batch-write multiple single-cell updates in a column by grouping contiguous rows.
* @param {GoogleAppsScript.Spreadsheet.Sheet} sh
* @param {number} colIdx 1-based column index (e.g., 2 for column B)
* @param {{r:number,v:any}[]} pairs 1-based row, value
*/
function groupedSetValues_(sh, colIdx, pairs){
  if (!pairs || !pairs.length) return;
  pairs.sort((a,b)=>a.r-b.r);
  let start = pairs[0].r;
  let block = [[pairs[0].v]];
  for (let i=1;i<pairs.length;i++){
    const prev = pairs[i-1].r, cur = pairs[i].r;
    if (cur === prev + 1){
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
    'Report Date:':'__InitDate', 'Customer Name:':'ClientName', 'APPT_ID:':'APPT_ID', 'Brand:':'Brand', 'Assigned Rep:':'AssignedRep',
    'Phone:':'Phone','Email:':'Email','Occasion:':'Occasion','Budget Range:':'BudgetRange','Decision-Maker:':'DecisionMaker','Initial Request:':'InitialRequest','SO#:':'SO_Number',
    'Sales Stage:':'SalesStage','Conversion Status:':'ConversionStatus','Custom Order Status:':'CustomOrderStatus','In Production Status:':'InProductionStatus','Center Stone Order Status:':'CenterStoneStatus',
    'Next Steps:':'NextSteps','Updated By:':'UpdatedBy','Updated At:':'UpdatedAt','Assisted Rep:':'AssistedRep',
    'Order Date:':'OrderDate'   // ← NEW
  };


  const rowsToScan = Math.min(sh.getLastRow() || 50, 50);
  if (rowsToScan <= 0) return;

  // Read A..D once (same as before)
  const values = sh.getRange(1, 1, rowsToScan, 4).getValues();

  // Collect writes for B and D only (do not touch any other cells)
  /** @type {{r:number,v:any}[]} */
  const writesB = [];
  /** @type {{r:number,v:any}[]} */
  const writesD = [];

  // Precompute today's date once (only used if Report Date is blank)
  const todayStr = Utilities.formatDate(new Date(), CS_TZ, 'yyyy-MM-dd');

  for (let i = 0; i < rowsToScan; i++) {
    const labA = String(values[i][0] || '').trim(); // col A
    const labC = String(values[i][2] || '').trim(); // col C

    const apply = (label, targetColIndexZeroBased) => {
      const key = map[label]; if (!key) return;

      if (key === '__InitDate') {
        // Only set Report Date if blank (identical to previous behavior)
        const current = String(values[i][targetColIndexZeroBased] || '').trim();
        if (!current) {
          if (targetColIndexZeroBased === 1) writesB.push({ r: i+1, v: todayStr });
          else if (targetColIndexZeroBased === 3) writesD.push({ r: i+1, v: todayStr });
        }
        return;
      }

      const newVal = data[key] != null ? String(data[key]) : '';
      if (targetColIndexZeroBased === 1) writesB.push({ r: i+1, v: newVal });
      else if (targetColIndexZeroBased === 3) writesD.push({ r: i+1, v: newVal });
    };

    if (labA) apply(labA, 1); // → B
    if (labC) apply(labC, 3); // → D
  }

  // Group contiguous rows per column into minimal setValues() calls
  if (writesB.length) groupedSetValues_(sh, 2, writesB); // col B
  if (writesD.length) groupedSetValues_(sh, 4, writesD); // col D
}

/** Normalize any cell value to HTML <input type="date"> format (YYYY-MM-DD) */
function toISODateForInput_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return Utilities.formatDate(v, CS_TZ, 'yyyy-MM-dd');
  }
  const s = String(v || '').trim();
  if (!s) return '';
  // already ISO?
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // parse common strings
  const d = new Date(s);
  if (!isNaN(d)) return Utilities.formatDate(d, CS_TZ, 'yyyy-MM-dd');
  // mm/dd/yyyy fallback
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) {
    const y = m[3].length === 2 ? ('20' + m[3]) : m[3];
    const mm = ('0' + m[1]).slice(-2), dd = ('0' + m[2]).slice(-2);
    return y + '-' + mm + '-' + dd;
  }
  return '';
}


/** One-time upgrade: ensure 03_Client_Status_Log has a trailing "In Production Status" column (header only). */
function CS_AUDIT_upgrade_addIPS_AtEnd() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('03_Client_Status_Log');
  if (!sh) throw new Error('Sheet "03_Client_Status_Log" not found.');

  // Read header row (row 1)
  const lastCol = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(x => String(x||'').trim());

  if (header.includes('In Production Status')) {
    Logger.log('Already present. No changes.');
    return;
  }

  // Append header at the end (does not shift existing columns)
  const newCol = lastCol + 1;
  sh.getRange(1, newCol).setValue('In Production Status');
  Logger.log('Added "In Production Status" as new last column ' + newCol + '.');
}


function cs_audit_appendByHeader_(sh, header, valuesByName) {
  const H = {}; header.forEach((h,i)=>{ if (h) H[String(h).trim()] = i; });
  const row = new Array(header.length).fill('');
  Object.keys(valuesByName).forEach(name=>{
    const i = H[name]; if (i != null) row[i] = valuesByName[name];
  });
  sh.appendRow(row);
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



