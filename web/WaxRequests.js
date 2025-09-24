/** WaxRequests.gs — v1.3 (modular)
 *  - Source of truth sheet: "05_Wax_Requests"
 *  - Status options pulled live from Dropdown tab row family under label "Wax Print Status"
 *  - Mirrors summary back to "00_Master Appointments": Wax Print Status, Wax Deadline (Admin), Wax Request URL
 *  - Foldering: prefers "05-3D Folder" on master row; else scans "Order Folder" for a "05-3D" child; creates "Wax Requests" under 05-3D.
 *
 *  NOTE: Uses your existing helpers from Resolver.gs: SH, headers_, getCell_, setCell_, appendObj_, idFromUrl_, ensureArtifactsForRow_.
 *        Those are already defined in your project. :contentReference[oaicite:2]{index=2} :contentReference[oaicite:3]{index=3}
 */

var WAX = {
  SHEET: '05_Wax_Requests',
  DROPDOWN_TAB: 'Dropdown',
  DROPDOWN_LABEL: 'Wax Print Status',
  MASTER_STATUS_COL: 'Wax Print Status',
  MASTER_DEADLINE_COL: 'Wax Deadline (Admin)',
  MASTER_URL_COL: 'Wax Request URL'
};

// ==== WAX DEBUG HELPERS ====
var WAX_DEBUG = false; // flip false to mute
function WAX_LOG(){ if(!WAX_DEBUG) return; try{ Logger.log([].slice.call(arguments).join(' ')); }catch(e){ Logger.log(String(arguments[0])); } }
function WAX_LOG_OBJ(label,obj){ if(!WAX_DEBUG) return; try{ Logger.log(label+' :: '+JSON.stringify(obj,null,2)); }catch(e){ Logger.log(label+' :: (json fail)'); } }
// HTML can call these to write into server logs
function wax__trace(msg){ WAX_LOG('[WAX][client]', msg); }
function wax__traceObj(label,obj){ WAX_LOG_OBJ('[WAX][client] '+label, obj); }

// --- Ping used by the dialog to prove google.script.run works end‑to‑end
function wax__ping(tag) {
  WAX_LOG('[WAX][ping] <- ' + tag);
  return { ok: true, tag: tag, ts: new Date() };
}


// ---------- Sheet + header ----------
function wax_ensureSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(WAX.SHEET);
  if (!sh) sh = ss.insertSheet(WAX.SHEET);

  var HEADER = [
    'WaxRequestID',               // WAX-YYYYMMDD-###
    'RootApptID',                 // join key to Master/CSR
    'SO/MO Number',
    'Brand',
    'Customer Name',
    'Assigned Rep',
    'Assisted Rep',
    'Requested By',               // human name
    'Priority',                   // Low/Normal/High/Rush (rep)
    'Requested Date',             // timestamp
    'Needed By (Rep)',            // rep's ask
    'Wax Print Status',           // dropdown from Dropdown tab
    'Wax Deadline (Admin)',       // admin-committed date
    'Estimated Print Date',       // optional
    'Completed Print Date',
    'Status Notes',
    'Master Row Link',
    'Updated By',                 // last admin who edited
    'Updated At',                 // last edit timestamp
    'Days Until Admin Deadline',  // AdminDeadline − Today
    'Overdue?',                   // OVERDUE if past Admin deadline
    'Days Late vs Rep Request'    // Completed − NeededBy (Rep)
  ];

  if (sh.getLastColumn() === 0) {
    sh.getRange(1,1,1,HEADER.length).setValues([HEADER]);
    sh.setFrozenRows(1);
  } else {
    // normalize header in place (width + names)
    if (sh.getLastColumn() !== HEADER.length) {
      if (sh.getLastColumn() > HEADER.length) {
        sh.deleteColumns(HEADER.length + 1, sh.getLastColumn() - HEADER.length);
      } else {
        sh.insertColumnsAfter(sh.getLastColumn(), HEADER.length - sh.getLastColumn());
      }
    }
    sh.getRange(1,1,1,HEADER.length).setValues([HEADER]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function wax_statusOptions() {
  var sh = SpreadsheetApp.getActive().getSheetByName(WAX.DROPDOWN_TAB);
  if (!sh) return DEFAULTS.slice();

  // Read whole tab once, then find the anchor label anywhere (case-insensitive, NBSP-safe)
  var vals = sh.getDataRange().getDisplayValues();
  var rows = vals.length, cols = vals[0] ? vals[0].length : 0;

  function norm(s){ return String(s||'').replace(/\u00A0/g,' ').trim().toLowerCase(); }

  var want = norm(WAX.DROPDOWN_LABEL); // e.g., "wax print status"
  var rLabel = -1, cLabel = -1;
  for (var r = 0; r < rows; r++) {
    for (var c = 0; c < cols; c++) {
      if (norm(vals[r][c]) === want) { rLabel = r; cLabel = c; break; }
    }
    if (rLabel >= 0) break;
  }
  if (rLabel < 0) return DEFAULTS.slice();

  // Collect values BELOW the label, same column, until the first blank
  var out = [];
  for (var rr = rLabel + 1; rr < rows; rr++) {
    var v = String(vals[rr][cLabel] || '').replace(/\u00A0/g,' ').trim();
    if (!v) break;
    out.push(v);
  }
  return out.length ? out : DEFAULTS.slice();
}


function wax_nextId_() {
  var tz = (typeof CFG !== 'undefined' && CFG.TZ) ? CFG.TZ : 'America/Los_Angeles';
  var ymd = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  var sh = wax_ensureSheet_();
  var H = headers_(WAX.SHEET);
  var col = H['WaxRequestID'];
  var last = sh.getLastRow();
  if (last < 2 || !col) return 'WAX-' + ymd + '-001';
  var ids = sh.getRange(2, col, last - 1, 1).getValues().flat();
  var count = ids.filter(function(v){ return String(v||'').indexOf('WAX-' + ymd) === 0; }).length + 1;
  return 'WAX-' + ymd + '-' + String(count).padStart(3, '0');
}

// ---------- Master mirror ----------
function wax_ensureMasterCols_() {
  var s = SH(SHT.MASTER), H = headers_(SHT.MASTER);
  function ensure(colName){
    if (!H[colName]) s.getRange(1, s.getLastColumn() + 1).setValue(colName);
  }
  ensure(WAX.MASTER_STATUS_COL);
  ensure(WAX.MASTER_DEADLINE_COL);
  ensure(WAX.MASTER_URL_COL);
}

function wax_mirrorToMaster_(rootApptId, status, adminDeadline, reqUrl) {
  wax_ensureMasterCols_();
  var s = SH(SHT.MASTER), H = headers_(SHT.MASTER);

  // find row by RootApptID
  var c = H['RootApptID']; if (!c) return;
  var last = s.getLastRow(); if (last < 2) return;
  var vals = s.getRange(2, c, last - 1, 1).getValues().flat();
  var idx = vals.findIndex(function(v){ return String(v||'') === String(rootApptId||''); });
  if (idx < 0) return;
  var row = idx + 2;

  if (H[WAX.MASTER_STATUS_COL])   s.getRange(row, H[WAX.MASTER_STATUS_COL]).setValue(status || '');
  if (H[WAX.MASTER_DEADLINE_COL]) s.getRange(row, H[WAX.MASTER_DEADLINE_COL]).setValue(adminDeadline || '');
  if (H[WAX.MASTER_URL_COL])      s.getRange(row, H[WAX.MASTER_URL_COL]).setValue(reqUrl || '');
}

// ---------- Folders (prefer "05-3D") ----------
function wax_base3DFolderForRoot_(rootApptId) {
  // Read master row once
  var s = SH(SHT.MASTER), H = headers_(SHT.MASTER);
  var col = H['RootApptID']; if (!col) return null;
  var last = s.getLastRow(); if (last < 2) return null;
  var vals = s.getRange(2, col, last - 1, 1).getValues().flat();
  var idx = vals.findIndex(function(v){ return String(v||'') === String(rootApptId||''); });
  if (idx < 0) return null;
  var row = idx + 2;

  // Try direct "05-3D Folder" (URL cell)
  var f3dUrl = H['05-3D Folder'] ? s.getRange(row, H['05-3D Folder']).getDisplayValue() : '';
  if (f3dUrl) {
    try {
      var id = (typeof idFromUrl_ === 'function') ? idFromUrl_(f3dUrl) : (f3dUrl.match(/[-\w]{25,}/)||[])[0];
      return id ? DriveApp.getFolderById(id) : null;
    } catch(e) { /* fall through */ }
  }

  // Else scan "Order Folder" for a child named "05-3D"
  var orderUrl = H['Order Folder'] ? s.getRange(row, H['Order Folder']).getDisplayValue() : '';
  if (orderUrl) {
    try {
      var ofId = (typeof idFromUrl_ === 'function') ? idFromUrl_(orderUrl) : (orderUrl.match(/[-\w]{25,}/)||[])[0];
      var orderFolder = ofId ? DriveApp.getFolderById(ofId) : null;
      if (orderFolder) {
        var it = orderFolder.getFoldersByName('05-3D'); // exact match
        if (it.hasNext()) return it.next();
        // try a looser search (rare)
        var child = orderFolder.createFolder('05-3D');
        return child;
      }
    } catch(e) { /* fall through */ }
  }

  // Last resort: ensure client artifacts; then create a "05-3D" under client folder
  try {
    var mrow = row;
    ensureArtifactsForRow_(mrow); // your existing helper creates client/prospect scaffolding. :contentReference[oaicite:4]{index=4}
    var cfId = H['ClientFolderID'] ? s.getRange(mrow, H['ClientFolderID']).getValue() : '';
    var clientFolder = cfId ? DriveApp.getFolderById(cfId) : null;
    if (clientFolder) {
      var f = clientFolder.getFoldersByName('05-3D');
      return f.hasNext() ? f.next() : clientFolder.createFolder('05-3D');
    }
  } catch (e) {}
  return null;
}

function wax_getOrCreateFolder_(rootApptId) {
  var base = wax_base3DFolderForRoot_(rootApptId);
  if (!base) return null;
  var it = base.getFoldersByName('Wax Requests');
  return it.hasNext() ? it.next() : base.createFolder('Wax Requests');
}

// ---------- Metrics ----------
function wax_recomputeMetricsForRow_(rowIdx) {
  var sh = SH(WAX.SHEET), H = headers_(WAX.SHEET);
  function V(name){ return H[name] ? sh.getRange(rowIdx, H[name]).getValue() : ''; }
  function S(name, val){ if (H[name]) sh.getRange(rowIdx, H[name]).setValue(val); }
  function dOnly(x){ try{ var d=new Date(x); return new Date(d.getFullYear(),d.getMonth(),d.getDate()).getTime(); }catch(_){return NaN;} }

  var today = new Date();
  var status = String(V('Wax Print Status')||'');
  var adminDeadline = V('Wax Deadline (Admin)');
  var repNeed = V('Needed By (Rep)');
  var completed = V('Completed Print Date');

  // Days Until Admin Deadline
  var daysLeft = '';
  if (adminDeadline && !/completed|canceled/i.test(status)) {
    daysLeft = Math.round((dOnly(adminDeadline) - dOnly(today)) / 86400000);
  }
  S('Days Until Admin Deadline', daysLeft);

  // Overdue?
  var overdue = '';
  if (adminDeadline && !/completed|canceled/i.test(status)) {
    overdue = (dOnly(today) > dOnly(adminDeadline)) ? 'OVERDUE' : '';
  }
  S('Overdue?', overdue);

  // Days Late vs Rep Request
  var late = '';
  if (completed && repNeed) {
    late = Math.round((dOnly(completed) - dOnly(repNeed)) / 86400000);
  }
  S('Days Late vs Rep Request', late);
}

// called from dialog AND from CSR server flow
function wax_onRequestSubmit_(payload) {
  // payload: { rootApptId, soMo, neededByRep, priority, requestedBy }
  var sh = wax_ensureSheet_();
  var H = headers_(WAX.SHEET);
  var id = wax_nextId_();
  var now = new Date();

  // Find Master row for snapshot + link
  var s = SH(SHT.MASTER), MH = headers_(SHT.MASTER);

  var last = s.getLastRow();
  if (last < 2) throw new Error('Master has no data.');

  var want = String(payload.rootApptId || '').trim();
  if (!want) throw new Error('Missing root/appt id for wax request.');

  // --- search RootApptID first, then APPT_ID as fallback ---
  var colRoot = MH['RootApptID'] || 0;
  var colAppt = MH['APPT_ID']    || 0;

  var idx = -1;

  if (colRoot) {
    var valsRoot = s.getRange(2, colRoot, last - 1, 1).getValues().flat();
    idx = valsRoot.findIndex(function(v){ return String(v||'').trim() === want; });
  }

  if (idx < 0 && colAppt) {
    var valsAppt = s.getRange(2, colAppt, last - 1, 1).getValues().flat();
    idx = valsAppt.findIndex(function(v){ return String(v||'').trim() === want; });
  }

  if (idx < 0) throw new Error('Root/Appt ID not found in Master: ' + want);

  var mRow = idx + 2;

  var brand   = MH['Brand'] ? s.getRange(mRow, MH['Brand']).getDisplayValue() : '';
  var cust    = MH['Customer Name'] ? s.getRange(mRow, MH['Customer Name']).getDisplayValue() : '';
  var assigned= MH['Assigned Rep'] ? s.getRange(mRow, MH['Assigned Rep']).getDisplayValue() : '';
  var assisted= MH['Assisted Rep'] ? s.getRange(mRow, MH['Assisted Rep']).getDisplayValue() : '';

  // Master row link
  var ssId = SpreadsheetApp.getActive().getId();
  var gid  = s.getSheetId();
  var masterRowLink = 'https://docs.google.com/spreadsheets/d/' + ssId + '/edit#gid=' + gid + '&range=A' + mRow;

  var newRowIdx = appendObj_(WAX.SHEET, { 'WaxRequestID': id });
  function setW(name, val){ if (H[name]) SH(WAX.SHEET).getRange(newRowIdx, H[name]).setValue(val); }

  setW('RootApptID', payload.rootApptId || '');
  setW('SO/MO Number', payload.soMo || '');
  setW('Brand', brand);
  setW('Customer Name', cust);
  setW('Assigned Rep', assigned);
  setW('Assisted Rep', assisted);
  setW('Requested By', payload.requestedBy || '');
  setW('Priority', payload.priority || '');
  setW('Requested Date', now);
  setW('Needed By (Rep)', payload.neededByRep || '');

  // default status = first option OR "Requested"
  var opts = wax_statusOptions();
  setW('Wax Print Status', opts.length ? opts[0] : 'Wax Requested');

  setW('Wax Deadline (Admin)', '');
  setW('Estimated Print Date', '');
  setW('Completed Print Date', '');
  setW('Status Notes', '');
  setW('Master Row Link', masterRowLink);
  setW('Updated By', '');
  setW('Updated At', '');

  wax_recomputeMetricsForRow_(newRowIdx);

  // Mirror to Master
  var rowUrl = wax_directRowUrl_(WAX.SHEET, newRowIdx);
  wax_mirrorToMaster_(payload.rootApptId, (opts.length ? opts[0] : 'Requested'), '', rowUrl);

  // Return folder URL so UI can prompt user to upload spec
  var f = wax_getOrCreateFolder_(payload.rootApptId);
  return { ok:true, requestId: id, url: rowUrl, folderUrl: (f ? f.getUrl() : '') };
}

function wax_directRowUrl_(sheetName, rowIdx) {
  var ss = SpreadsheetApp.getActive();
  var sh = SH(sheetName);
  return 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + sh.getSheetId() + '&range=A' + rowIdx;
}


function wax_adminOpenDialog_() {
  const out = HtmlService
    .createHtmlOutputFromFile('WaxPendingDialog')
    .setWidth(1100).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(out, 'Wax — Pending Requests');
}



function wax_adminGetPendingData() {
  WAX_LOG('[WAX][fetch] called by dialog');
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(WAX.SHEET); // '05_Wax_Requests'
  if (!sh) { WAX_LOG('[WAX][fetch] sheet not found:', WAX.SHEET); throw new Error('Sheet not found: ' + WAX.SHEET); }

  var tz = (typeof CFG !== 'undefined' && CFG.TZ) ? CFG.TZ : (Session.getScriptTimeZone() || 'America/Los_Angeles');

  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  WAX_LOG('[WAX][fetch] START sheet=' + sh.getName() + ' gid=' + sh.getSheetId() + ' lastRow=' + lastRow + ' lastCol=' + lastCol);

  if (lastRow < 2) {
    WAX_LOG('[WAX][fetch] no data rows.');
    return { rows: [], statusOptions: wax_statusOptions() };
  }

  var header = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h){
    return String(h||'').replace(/\u00A0/g, ' ').trim();
  });

  // Build both an exact map and a normalized map
  var idx = {}; var idxNorm = {};
  function norm(s){ return String(s||'').toLowerCase().replace(/\u00A0/g,' ').replace(/[\s_\-()/]+/g,'').trim(); }
  header.forEach(function(h,i){
    if (h) {
      idx[h] = i;
      idxNorm[norm(h)] = i;
    }
  });

  // allow aliases/synonyms for columns that may vary
  var ALIASES = {
    'Needed By (Rep)': [
      'Needed By (Rep)',
      'Needed by (Rep)',
      'Rep Needed-By',
      'Rep Needed By',
      'Needed-By Rep',
      'Rep NeededBy'
    ]
  };

  // tolerant index finder: exact → alias list → normalized
  function iAny(name){
    if (idx[name] != null) return idx[name];
    var variants = ALIASES[name] || [name];
    for (var k=0; k<variants.length; k++){
      var v = variants[k];
      if (idx[v] != null) return idx[v];
      var n = norm(v);
      if (idxNorm[n] != null) return idxNorm[n];
    }
    throw new Error('Missing column in '+WAX.SHEET+': "'+name+'"');
  }

  // [DBG] header + key columns present/missing
  WAX_LOG_OBJ('[WAX][fetch] header', header);
  ['WaxRequestID','SO/MO Number','Customer Name','Priority','Wax Print Status','Wax Deadline (Admin)','Estimated Print Date','Completed Print Date','Status Notes','Master Row Link','Needed By (Rep)']
    .forEach(function(name){ WAX_LOG('[WAX][fetch] col['+name+']=' + (idx[name]!=null ? (idx[name]+1) : 'MISSING')); });

  function i(name){ if (idx[name] == null) throw new Error('Missing column in '+WAX.SHEET+': "'+name+'"'); return idx[name]; }

  // Read everything as display text (robust + JSON‑friendly)
  var data = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  WAX_LOG('[WAX][fetch] rows(all)=' + data.length);

  var stCol = i('Wax Print Status');

  var rows = data
    .filter(function(row){
      var st = String(row[stCol] || '').trim().toLowerCase();
      return !/(^|\s)(completed|canceled|cancelled)(\s|$)/.test(st);
    })
    .map(function(row){
      return {
        id:            row[i('WaxRequestID')],
        so:            row[i('SO/MO Number')],
        customer:      row[i('Customer Name')],
        priority:      row[i('Priority')],
        status:        row[stCol],
        repNeed:       row[iAny('Needed By (Rep)')],         // already display text
        adminDeadline: row[i('Wax Deadline (Admin)')],
        estPrint:      row[i('Estimated Print Date')],
        completed:     row[i('Completed Print Date')],
        notes:         row[i('Status Notes')],
        link:          row[i('Master Row Link')]
      };
    });
  WAX_LOG('[WAX][fetch] rows(pending)=' + rows.length);


  // [DBG] options pull
  var opts = wax_statusOptions();
  WAX_LOG_OBJ('[WAX][fetch] status options', opts);

  WAX_LOG('[WAX][fetch] returning rows=' + rows.length + ' opts=' + opts.length);
  return { rows: rows, statusOptions: opts };
}



function wax_adminCommitFromDialog(payload) {
  if (!payload || !Array.isArray(payload.updates) || !payload.updates.length) {
    throw new Error('No updates received.');
  }
  var sh = SH(WAX.SHEET), H = headers_(WAX.SHEET);
  var now = new Date(), me = (Session.getActiveUser().getEmail() || 'admin');

  // Index by WaxRequestID
  var colId = H['WaxRequestID'];
  var ids = sh.getRange(2, colId, Math.max(0, sh.getLastRow() - 1), 1).getValues().flat();

  var summaries = [];

  payload.updates.forEach(function(u){
    var id = String(u.id || '').trim(); if (!id) return;
    var idx = ids.findIndex(function(v){ return String(v||'') === id; });
    if (idx < 0) return;
    var row = idx + 2;

    function setW(name, val){ if (H[name]) sh.getRange(row, H[name]).setValue(val); }

    function dateOrBlank(s){
      if (!s) return '';
      // handle both "MM/DD/YYYY" and "YYYY-MM-DD"
      var iso = s.indexOf('-') > -1 ? s : (function(){
        var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (!m) return s;
        return m[3] + '-' + ('0'+m[1]).slice(-2) + '-' + ('0'+m[2]).slice(-2);
      })();
      return new Date(iso + 'T12:00:00');
    }

    // Coerce dates once, reuse for both write + metrics (midday avoids TZ edge cases)
    var adminDate    = dateOrBlank(u.adminDeadline || '');
    var estDate      = dateOrBlank(u.estPrint      || '');
    var completedDate= dateOrBlank(u.completed     || '');
    var statusNew    = String(u.status || '');

    setW('Wax Print Status', statusNew);
    setW('Wax Deadline (Admin)', adminDate);
    setW('Estimated Print Date', estDate);
    setW('Completed Print Date', completedDate);
    setW('Status Notes', String(u.notes || ''));
    setW('Updated By', me);
    setW('Updated At', now);

    // ---- Inline metrics (avoid extra reads)
    function dOnly(x){ try{ var d=new Date(x); return new Date(d.getFullYear(),d.getMonth(),d.getDate()).getTime(); }catch(_){ return NaN; } }

    // Days Until Admin Deadline & Overdue?
    var daysLeft = '';
    var overdue  = '';
    if (adminDate && !/(^|\s)(completed|canceled|cancelled)(\s|$)/.test(statusNew.toLowerCase())) {
      daysLeft = Math.round((dOnly(adminDate) - dOnly(new Date())) / 86400000);
      overdue  = (dOnly(new Date()) > dOnly(adminDate)) ? 'OVERDUE' : '';
    }
    setW('Days Until Admin Deadline', daysLeft);
    setW('Overdue?', overdue);

    // Days Late vs Rep Request (needs the sheet's Rep date, which we didn't change)
    var daysLate = '';
    if (completedDate && H['Needed By (Rep)']) {
      var repNeedCell = sh.getRange(row, H['Needed By (Rep)']).getValue(); // 1 quick read
      if (repNeedCell) daysLate = Math.round((dOnly(completedDate) - dOnly(repNeedCell)) / 86400000);
    }
    setW('Days Late vs Rep Request', daysLate);


    var root = H['RootApptID'] ? sh.getRange(row, H['RootApptID']).getValue() : '';
    var st   = H['Wax Print Status'] ? sh.getRange(row, H['Wax Print Status']).getValue() : '';
    var dl   = H['Wax Deadline (Admin)'] ? sh.getRange(row, H['Wax Deadline (Admin)']).getValue() : '';
    var url  = wax_directRowUrl_(WAX.SHEET, row);
    wax_mirrorToMaster_(String(root||''), String(st||''), dl || '', url);

    var so = H['SO/MO Number'] ? sh.getRange(row, H['SO/MO Number']).getDisplayValue() : '';
    summaries.push('• ' + id + ' — SO: ' + (so || 'n/a'));
  });

  return { ok: true, message: 'Updated ' + summaries.length + ' record(s):\n\n' + summaries.join('\n') };
}


// ---------- Menu hook (call from your onOpen) ----------
function addWaxAdminMenu_(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Wax Admin')
    .addItem('Pull Pending', 'wax_adminBuildPending_')
    .addItem('Commit Updates', 'wax_adminCommit_')
    .addToUi();     
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



function wax__debugFetchOnce(){
  try {
    var res = wax_adminGetPendingData();
    WAX_LOG('[WAX][runner] returned rows=' + (res.rows||[]).length + ' | opts=' + (res.statusOptions||[]).length);
  } catch (e) {
    WAX_LOG('[WAX][runner][ERROR] ' + (e && e.message ? e.message : e));
    throw e;
  }
}

function wax_adminOpenDialog_TEST() {
  var html = HtmlService.createHtmlOutput('<html><body><div id="x">boot</div><script>document.getElementById("x").textContent="JS ran";google.script.run.wax__trace("TEST dialog JS ran");</script></body></html>');
  SpreadsheetApp.getUi().showModalDialog(html, 'TEST');
}

