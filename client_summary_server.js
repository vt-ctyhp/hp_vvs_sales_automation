/** File: 04 - client_summary_server.gs
 * Client Summary — server + opener (no external dependencies).
 * Master tab is hardcoded to "00_Master Appointments".
 */

const CSU_MASTER_NAME = '00_Master Appointments';
const CSU_TZ = Session.getScriptTimeZone() || 'America/Los_Angeles';

/** ===== Menu opener ===== */
function openClientSummary() {
  // Use the canonical filename. If you keep a prefixed copy, uncomment the fallback.
  var html;
  try {
    html = HtmlService.createHtmlOutputFromFile('dlg_client_summary_v1');
  } catch (e) {
    throw new Error('Missing HTML file "dlg_client_summary_v1.html". Create it and try again.');
    // Fallback if you insist on a prefixed file name:
    // html = HtmlService.createHtmlOutputFromFile('04.1 - dlg_client_summary_v1');
  }
  html.setWidth(1040).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Client Summary');
}

/** ===== Internal helpers (no external dependencies) ===== */
function csu_getActiveRowContext_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CSU_MASTER_NAME);
  if (!sh) throw new Error('Missing sheet: ' + CSU_MASTER_NAME);

  const rg = sh.getActiveRange();
  if (!rg || rg.getRow() < 2 || rg.getSheet().getName() !== CSU_MASTER_NAME) {
    throw new Error('Select a data row on "' + CSU_MASTER_NAME + '" and try again.');
  }

  const row = rg.getRow();
  const lastCol = sh.getLastColumn();

  const header = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (h) { return String(h || '').trim(); });
  const H = {}; header.forEach(function (h, i) { if (h) H[h] = i; });

  // values (typed), display strings, and rich text (for links)
  const valRow = sh.getRange(row, 1, 1, lastCol);
  const vals = valRow.getValues()[0];
  const disp = valRow.getDisplayValues()[0];
  var rtv = null; try { rtv = valRow.getRichTextValues()[0]; } catch (_){}

  return {
    ss: ss, sh: sh, row: row, header: header, H: H, vals: vals, disp: disp, rtv: rtv,
    sheetId: sh.getSheetId(), baseUrl: ss.getUrl()
  };
}

function csu_getVal_(ctx, name) {
  return (ctx.H[name] != null) ? ctx.vals[ctx.H[name]] : '';
}

function csu_getDisp_(ctx, name) {
  return (ctx.H[name] != null) ? String(ctx.disp[ctx.H[name]] || '') : '';
}

function csu_linkOrText_(ctx, name) {
  if (ctx.H[name] == null) return '';
  try {
    const rt = ctx.rtv ? ctx.rtv[ctx.H[name]] : null;
    const u = rt && rt.getLinkUrl && rt.getLinkUrl();
    if (u) return u;
  } catch (_){}
  return String(ctx.disp[ctx.H[name]] || '');
}

function csu_digits_(s) { return String(s || '').replace(/\D/g, ''); }

function csu_combineDateTime_(dateVal, timeVal) {
  // Works with Google Sheets date (Date), Excel time fraction (Number), Date, or "h:mm AM/PM" string
  if (!(dateVal instanceof Date) || isNaN(dateVal)) return null;
  var d = new Date(dateVal.getTime());

  if (timeVal instanceof Date && !isNaN(timeVal)) {
    d.setHours(timeVal.getHours(), timeVal.getMinutes(), timeVal.getSeconds(), 0);
    return d;
  }
  if (typeof timeVal === 'number' && !isNaN(timeVal)) {
    var totalSec = Math.round(timeVal * 24 * 3600);
    var hh = Math.floor(totalSec / 3600);
    var mm = Math.floor((totalSec % 3600) / 60);
    var ss = totalSec % 60;
    d.setHours(hh, mm, ss, 0);
    return d;
  }
  if (typeof timeVal === 'string' && timeVal) {
    var m = timeVal.match(/^\s*(\d{1,2}):(\d{2})(?:\s*([APap][Mm]))?\s*$/);
    if (m) {
      var hh2 = parseInt(m[1], 10);
      var mm2 = parseInt(m[2], 10);
      var ampm = m[3] ? m[3].toUpperCase() : '';
      if (ampm === 'PM' && hh2 < 12) hh2 += 12;
      if (ampm === 'AM' && hh2 === 12) hh2 = 0;
      d.setHours(hh2, mm2, 0, 0);
      return d;
    }
  }
  // No time: return the date-only
  return d;
}

function csu_fmt_(d) {
  if (!d || isNaN(d)) return '';
  return Utilities.formatDate(d, CSU_TZ, 'EEE, MMM d, yyyy h:mm a');
}

function csu_masterRowLink_(ctx) {
  return ctx.baseUrl + '#gid=' + ctx.sheetId + '&range=A' + ctx.row;
}

/** ===== Server APIs (called from HTML) ===== */

/** Bootstrap: read the active row + build a summary object. */
function csu_bootstrap() {
  try {
    var ctx = csu_getActiveRowContext_();

    // Your exact columns:
    // RootApptID, SO#, Visit Date, Visit Time, Visit Type, Status,
    // Sales Stage, Conversion Status, Custom Order Status, Center Stone Order Status, Brand
    var apptId     = csu_getDisp_(ctx, 'RootApptID');
    var so         = String(csu_getDisp_(ctx, 'SO#') || '').replace(/^'/, '').trim();
    var visitDate  = csu_getVal_(ctx, 'Visit Date');
    var visitTime  = csu_getVal_(ctx, 'Visit Time');
    var visitType  = csu_getDisp_(ctx, 'Visit Type'); // (e.g., In Store / Virtual)
    var apptDt     = csu_combineDateTime_(visitDate, visitTime);
    var apptNice   = csu_fmt_(apptDt);

    var brand      = csu_getDisp_(ctx, 'Brand');

    // Optional columns (if they exist); will show "—" if absent
    var assignedRep  = csu_getDisp_(ctx, 'Assigned Rep');
    var assistedRep  = csu_getDisp_(ctx, 'Assisted Rep');
    var shortTag     = csu_getDisp_(ctx, 'Short Tag');
    var nextSteps    = csu_getDisp_(ctx, 'Next Steps');
    var customer     = csu_getDisp_(ctx, 'Customer Name') || csu_getDisp_(ctx, 'Customer') || csu_getDisp_(ctx, 'Client Name');

    var emailLower   = (csu_getDisp_(ctx, 'EmailLower') || csu_getDisp_(ctx, 'Email')).toLowerCase();
    var phoneNorm    = csu_digits_(csu_getDisp_(ctx, 'PhoneNorm') || csu_getDisp_(ctx, 'Phone'));

    // Statuses
    var salesStage   = csu_getDisp_(ctx, 'Sales Stage');
    var convStatus   = csu_getDisp_(ctx, 'Conversion Status');
    var customOrder  = csu_getDisp_(ctx, 'Custom Order Status');
    var centerStone  = csu_getDisp_(ctx, 'Center Stone Order Status');

    // Design preview (text on master, optional)
    var designBrief = csu_getDisp_(ctx, 'Design Request') || csu_getDisp_(ctx, '3D Design Request') ||
                      csu_getDisp_(ctx, 'Design Brief')   || csu_getDisp_(ctx, '3D Notes');

    // Quick links (prefer the hyperlink in the cell)
    function link(name) { return csu_linkOrText_(ctx, name); }
    var links = {
      masterRow:     csu_masterRowLink_(ctx),
      odooSO:        link('Odoo SO URL'),
      clientFolder:  link('Client Folder'),
      orderFolder:   link('Order Folder'),
      folder3d:      link('05-3D Folder'),
      intake00:      link('00-Intake'),
      tracker3d:     link('3D Tracker'),
      statusReport:  link('Client Status Report URL'),
      checklist:     link('Checklist URL'),
      quotation:     link('Quotation URL'),
      soShortcut:    link('SO Shortcut in Client')
    };

    // === Deadlines (3D & Production) ===
    // Flexible header lookups: adjust the alias lists to match your sheet if needed.
    function pickDate_(names){
      for (var i = 0; i < names.length; i++){
        var v = csu_getVal_(ctx, names[i]);
        if (v instanceof Date && !isNaN(v)) return v;
        var s = csu_getDisp_(ctx, names[i]);
        if (s){ var d = new Date(s); if (!isNaN(d)) return d; }
      }
      return null;
    }
    function pickCount_(names){
      for (var i = 0; i < names.length; i++){
        var v = csu_getVal_(ctx, names[i]);
        if (typeof v === 'number' && !isNaN(v)) return Math.max(0, Math.floor(v));
        var s = csu_getDisp_(ctx, names[i]);
        if (s){ var n = parseInt(csu_digits_(s), 10); if (!isNaN(n)) return Math.max(0, n); }
      }
      return 0;
    }

    var d3dDate  = pickDate_(['3D Deadline','3D Due','3D Due Date','3D Target Date']);
    var d3dMoves = pickCount_(['3D Deadline Moves','3D Moves','# times 3D deadline moved','3D Deadline Change Count']);

    var prodDate  = pickDate_(['Prod. Deadline','Production Deadline','Prod Deadline','Production Due','Production Due Date','Production Target Date']);
    var prodMoves = pickCount_(['Prod. Deadline Moves','Production Deadline Moves','Prod Moves','# times prod deadline moved','Production Change Count']);

    var d3dText  = (csu_fmt_(d3dDate)  || '—') + (d3dMoves > 0 ? (' · moved ' + d3dMoves + '×') : '');
    var prodText = (csu_fmt_(prodDate) || '—') + (prodMoves > 0 ? (' · moved ' + prodMoves + '×') : '');


    // Colors: if you have a "readDropdowns_" function elsewhere it will be used; otherwise chips render with default style.
    var colors = { salesStage:{}, convStatus:{}, customOrder:{}, centerStone:{} };
    try {
      if (typeof readDropdowns_ === 'function') {
        var dd = readDropdowns_();
        if (dd && dd.colors) colors = dd.colors;
      }
    } catch (_){}

    // Visit Type appended to the time if present
    if (visitType) apptNice = apptNice ? (apptNice + ' · ' + visitType) : visitType;

    return {
      ok: true,
      atGlance: {
        apptId: apptId,
        apptTime: apptNice,
        brand: brand,
        assignedRep: assignedRep,
        assistedRep: assistedRep,
        so: so,
        shortTag: shortTag,
        customer: customer,
        emailLower: emailLower,
        phoneNorm: phoneNorm,
        nextSteps: nextSteps
      },
      // NEW
      deadlines: { d3d: d3dText, prod: prodText },

      statuses: { salesStage: salesStage, convStatus: convStatus, customOrder: customOrder, centerStone: centerStone },
      designBrief: designBrief,
      colors: colors,
      links: links
    };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

/** Appointment history: all rows with the same EmailLower OR PhoneNorm. */
function csu_apptHistory() {
  try {
    var ctx = csu_getActiveRowContext_();
    var sh = ctx.sh, H = ctx.H, ss = ctx.ss;

    var emailLower = (csu_getDisp_(ctx, 'EmailLower') || csu_getDisp_(ctx, 'Email')).toLowerCase();
    var phoneNorm  = csu_digits_(csu_getDisp_(ctx, 'PhoneNorm') || csu_getDisp_(ctx, 'Phone'));

    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, items: [] };

    var lastCol = sh.getLastColumn();
    var dataVals = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var dataDisp = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

    function col(name) { return (H[name] != null) ? H[name] : -1; }

    var cApptID = col('RootApptID');
    var cEmail  = (H['EmailLower'] != null) ? H['EmailLower'] : (H['Email'] != null ? H['Email'] : -1);
    var cPhone  = (H['PhoneNorm'] != null) ? H['PhoneNorm'] : (H['Phone'] != null ? H['Phone'] : -1);
    var cBrand  = col('Brand');
    var cSO     = (H['SO#'] != null) ? H['SO#'] : -1;

    var cVD     = col('Visit Date');
    var cVT     = col('Visit Time');

    var cSS     = col('Sales Stage');
    var cCS     = col('Conversion Status');
    var cCOS    = col('Custom Order Status');
    var cCSOS   = col('Center Stone Order Status');

    var items = [];
    for (var i = 0; i < dataVals.length; i++) {
      var em = (cEmail >= 0 ? String(dataDisp[i][cEmail] || '').toLowerCase() : '');
      var ph = (cPhone >= 0 ? csu_digits_(dataDisp[i][cPhone] || '') : '');
      var hit = (emailLower && em && emailLower === em) || (phoneNorm && ph && phoneNorm === ph);
      if (!hit) continue;

      var vd = (cVD >= 0) ? dataVals[i][cVD] : '';
      var vt = (cVT >= 0) ? dataVals[i][cVT] : '';
      var dt = csu_combineDateTime_(vd, vt);

      var rowNum = i + 2;
      items.push({
        row: rowNum,
        link: ss.getUrl() + '#gid=' + sh.getSheetId() + '&range=A' + rowNum,
        apptId: (cApptID >= 0 ? dataDisp[i][cApptID] : ''),
        brand:  (cBrand >= 0 ? dataDisp[i][cBrand] : ''),
        so:     (cSO >= 0 ? String(dataDisp[i][cSO] || '').replace(/^'/, '') : ''),
        time:   csu_fmt_(dt),
        statuses: {
          salesStage:   (cSS >= 0 ? dataDisp[i][cSS] : ''),
          convStatus:   (cCS >= 0 ? dataDisp[i][cCS] : ''),
          customOrder:  (cCOS >= 0 ? dataDisp[i][cCOS] : ''),
          centerStone:  (cCSOS >= 0 ? dataDisp[i][cCSOS] : '')
        }
      });
    }

    // Sort newest first by computed time, then by row
    items.sort(function (a, b) {
      var da = new Date(a.time), db = new Date(b.time);
      if (!isNaN(db) && !isNaN(da)) return db - da;
      return (b.row || 0) - (a.row || 0);
    });

    return { ok: true, items: items };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e), items: [] };
  }
}

/** Lightweight design snapshot: links + best‑effort latest 3D spec from Tracker→Log. */
function csu_last3DSnapshot() {
  try {
    var ctx = csu_getActiveRowContext_();

    var so = String(csu_getDisp_(ctx, 'SO#') || '').replace(/^'/,'').trim();
    var trackerUrl = csu_linkOrText_(ctx, '3D Tracker');
    var folderUrl  = csu_linkOrText_(ctx, '05-3D Folder');

    var payload = { ok:true, trackerUrl: trackerUrl, folderUrl: folderUrl, latest:null, info:'' };
    if (!trackerUrl){ payload.info = 'No 3D tracker link on this row.'; return payload; }

    var id = _idFromUrlLoose_(trackerUrl);
    if (!id){ payload.info = 'Invalid 3D tracker link.'; return payload; }

    var ssT = SpreadsheetApp.openById(id);
    var shT = ssT.getSheetByName('Log') || ssT.getSheetByName('3D Revision Log') || ssT.getSheetByName('Revision Log');
    if (!shT || shT.getLastRow() < 2){ payload.info = 'Tracker has no log yet.'; return payload; }

    var lastCol = shT.getLastColumn();
    var header = shT.getRange(1,1,1,lastCol).getDisplayValues()[0].map(function(h){ return String(h||'').trim(); });
    var pos = {}; header.forEach(function(h,i){ if(h) pos[h]=i+1; });

    function col(names){ for (var i=0;i<names.length;i++){ var c=pos[names[i]]; if (c) return c; } return 0; }
    var cSO   = col(['SO#','SO Number','SO']);
    var cRev  = col(['Revision #','Revision','Rev #']);
    var cDate = col(['Log Date','Date','Updated At']);
    var cBy   = col(['Updated By','Designer','Logged By']);

    var n = shT.getLastRow() - 1; // data rows
    var pick = 0;

    if (cSO){
      var soVals  = shT.getRange(2, cSO,  n, 1).getDisplayValues();
      if (cRev){
        var revVals = shT.getRange(2, cRev, n, 1).getValues(); // numeric for max
        var maxRev = -1;
        for (var i=0;i<n;i++){
          var mSO = String(soVals[i][0]||'').trim();
          if (so && mSO === so){
            var rv = Number(revVals[i][0]||0);
            if (rv > maxRev){ maxRev = rv; pick = 2+i; }
          }
        }
      }
      if (!pick){ // fallback: last by order for this SO
        for (var j=n-1;j>=0;j--){
          var mSO2 = String(soVals[j][0]||'').trim();
          if (!so || mSO2 === so){ pick = 2+j; break; }
        }
      }
    } else {
      // No SO column → last non-empty row
      pick = shT.getLastRow();
    }

    if (!pick){ payload.info = 'No matching SO# in tracker.'; return payload; }

    var rowVals = shT.getRange(pick, 1, 1, lastCol).getDisplayValues()[0];
    function get(names){ for (var i=0;i<names.length;i++){ var idx=header.indexOf(names[i]); if (idx>=0) return rowVals[idx]; } return ''; }

    // Map common spec fields; extend aliases here if your tracker differs
    var specPairs = [
      ['Ring Style',        ['Ring Style','Style']],
      ['Metal',             ['Metal']],
      ['US Size',           ['US Size','Size']],
      ['Band Width (mm)',   ['Band Width (mm)','Band Width']],
      ['Center Type',       ['Center Type','Center Diamond Type','Center']],
      ['Shape',             ['Shape']],
      ['Diamond Dimension', ['Diamond Dimension','Center Dimensions']],
      ['Accent Type',       ['Accent Type','Accent Diamond Type']],
      ['Design Notes',      ['Design Notes','Notes']]
    ];

    var lis = [];
    for (var k=0;k<specPairs.length;k++){
      var label=specPairs[k][0], val=get(specPairs[k][1]);
      if (val) lis.push('<li><b>'+label+':</b> '+val+'</li>');
    }

    var revNo = cRev  ? shT.getRange(pick, cRev ).getDisplayValue() : '';
    var when  = cDate ? shT.getRange(pick, cDate).getDisplayValue() : '';
    var who   = cBy   ? shT.getRange(pick, cBy  ).getDisplayValue() : '';

    payload.latest = {
      revNo: revNo,
      when: when,
      who: who,
      html: (lis.length ? '<ul>'+lis.join('')+'</ul>' : '')
    };
    if (!payload.latest.html) payload.info = 'Latest log row found but no recognizable spec fields.';

    return payload;

  } catch (e) {
    return { ok:false, error: String(e && e.message || e) };
  }
}

/** Extract a Drive file ID from most Google URLs. */
function _idFromUrlLoose_(url){
  if (!url) return '';
  var m = String(url).match(/[-\w]{25,}/);
  return m ? m[0] : '';
}


/** Optional health check (handy during setup). */
function csu_ping() { return 'pong'; }

/**
 * Adapter so Client Summary can open the Client Status flow reliably.
 * It tries known openers in your ClientStatus module and falls back gracefully.
 */
function csu_openClientStatus_() {
  try {
    // Preferred openers (if defined in ClientStatus_v1.gs)
    if (typeof openClientStatus === 'function')        return openClientStatus();
    if (typeof openClientStatusUpdate === 'function')  return openClientStatusUpdate();

    // Legacy/step-based fallbacks (v1 scaffolding)
    if (typeof cs_openStatusDialog_ === 'function')    return cs_openStatusDialog_();
    if (typeof cs_showReadOnlyPing_ === 'function')    return cs_showReadOnlyPing_();

    // As a last resort, inform the user rather than silently failing
    SpreadsheetApp.getUi().alert('Client Status module not available: no opener function found.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error opening Client Status: ' + (e && e.message ? e.message : e));
  }
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



