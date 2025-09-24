/**
 * Start 3D Design – Configuration & Header Check
 * Purpose: Verify that required Script Properties are set, target sheets exist,
 * and critical headers are present (header-based lookups; resilient to column reorders).
 *
 * Run from Apps Script: select start3d_configCheck() ▶︎ Run
 * See logs under: View ▸ Logs
 */
function start3d_configCheck() {
  const sp = PropertiesService.getScriptProperties();

  // --- Required Script Properties (no guesses; set these explicitly) ---
  const required = {
    SS_3D_DESIGN_TRACKER_ID: null,                 // Spreadsheet ID of "3D Design Requests Tracker" (production, not template)
    TAB_3D_LOG: 'Log',                              // Tab name in tracker for the log (confirm exact casing)
    SS_SALES_INTAKE_ID: null,                       // Spreadsheet ID of "100_HPUSA_VVS - Sales Intake Tracking"
    TAB_100_DESIGN_REQUEST: '100_ Design Request',  // Tab name (confirm exact)
    TAB_00_MASTER_APPOINTMENTS: '00_Master Appointments', // Tab name (confirm exact)
    COL_100_SO: null,                               // Exact header text for SO# in 100_
    COL_100_DESCRIPTION: null,                      // Exact header text for Description in 100_
    COL_00_BRAND: null,                             // Exact header text for Brand in 00_
    URL_302_REPORT: null,                           // Viewer link to 302_[VVS] Sales Orders Report
    URL_3D_REVISION_LOG: null                       // Viewer link to the 3D Revision Log (where you want success screen to point)
  };

  // Report missing Script Properties
  const missingKeys = Object.keys(required).filter(k => !sp.getProperty(k));
  Logger.log('Start3D ▸ Config Check — Missing properties: %s', missingKeys.length ? missingKeys.join(', ') : 'none');
  if (missingKeys.length) {
    Logger.log('Set Script Properties under Project Settings ▸ Script properties, then re-run this check.');
    return;
  }

  // --- Open spreadsheets/tabs and validate headers ---
  const trackerId = sp.getProperty('SS_3D_DESIGN_TRACKER_ID');
  const trackerTab = sp.getProperty('TAB_3D_LOG');
  const intakeId  = sp.getProperty('SS_SALES_INTAKE_ID');
  const tab100    = sp.getProperty('TAB_100_DESIGN_REQUEST');
  const tab00     = sp.getProperty('TAB_00_MASTER_APPOINTMENTS');
  const hSO       = sp.getProperty('COL_100_SO');
  const hDesc     = sp.getProperty('COL_100_DESCRIPTION');
  const hBrand    = sp.getProperty('COL_00_BRAND');

  // 3D Tracker: existence + required headers
  const ssTracker = SpreadsheetApp.openById(trackerId);
  const shLog = ssTracker.getSheetByName(trackerTab);
  if (!shLog) throw new Error(`3D Tracker: sheet "${trackerTab}" not found`);

  const headers3D = _getHeaders_(shLog);
  Logger.log('3D Tracker headers (trimmed): %s', Array.from(headers3D.keys()).join(' | '));

  // Required headers in 3D tracker (trim-insensitive)
  const required3D = [
    'Timestamp','User','Action','Revision #','Mode',
    'Accent Type','Ring Style','Metal','US Size','Band Width (mm)',
    'Center Type','Shape','Diamond Dimension','Design Notes',
    'Short Tag','SO#','Brand','Odoo SO URL','Master Link'
  ];
  const missing3D = required3D.filter(h => !headers3D.has(h));
  if (missing3D.length) {
    Logger.log('WARNING ▸ 3D Tracker missing expected headers: %s', missing3D.join(', '));
  } else {
    Logger.log('3D Tracker header check: OK');
  }

  // Sales Intake: 100_ Design Request
  const ssIntake = SpreadsheetApp.openById(intakeId);
  const sh100 = ssIntake.getSheetByName(tab100);
  if (!sh100) throw new Error(`Sales Intake: sheet "${tab100}" not found`);
  const headers100 = _getHeaders_(sh100);
  Logger.log('100_ headers (trimmed): %s', Array.from(headers100.keys()).join(' | '));
  if (!headers100.has(hSO))   Logger.log('WARNING ▸ 100_ missing SO header "%s"', hSO);
  if (!headers100.has(hDesc)) Logger.log('WARNING ▸ 100_ missing Description header "%s"', hDesc);

  // Sales Intake: 00_Master Appointments
  const sh00 = ssIntake.getSheetByName(tab00);
  if (!sh00) throw new Error(`Sales Intake: sheet "${tab00}" not found`);
  const headers00 = _getHeaders_(sh00);
  Logger.log('00_Master headers (trimmed): %s', Array.from(headers00.keys()).join(' | '));
  if (!headers00.has(hBrand)) Logger.log('WARNING ▸ 00_Master missing Brand header "%s"', hBrand);

  Logger.log('Start3D ▸ Config Check — COMPLETE');
}

/**
 * Helper: read header row (row 1), return Map<headerTextTrimmed, 1-based column index>.
 */
function _getHeaders_(sh) {
  const values = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || [];
  const map = new Map();
  values.forEach((h, i) => {
    const t = (h == null ? '' : String(h)).trim();
    if (t) map.set(t, i + 1);
  });
  return map;
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



