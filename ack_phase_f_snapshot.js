/***** Phase F — Morning Snapshot Mode *****
 * Captures a frozen "expected set" (post schedule + Maria/Paul coverage) at ~8:30 AM PT
 * and writes it to 13_Morning_Snapshot; appends audit to 14_Snapshot_Log.
 * Also provides helpers to read today's snapshot and to create/remove a daily trigger.
 ****************************************************************/

// Fallback timezone if TIMEZONE not defined elsewhere
const TZ_SNAPSHOT = (typeof TIMEZONE !== 'undefined' && TIMEZONE) ? TIMEZONE : 'America/Los_Angeles';

/** Headers for snapshot sheets */
function snapshotHeaders13_() {
  return [
    'Snapshot Date','Captured At','RootApptID','Rep','Role',
    'Customer Name','Sales Stage','Custom Order Status','Updated By','Updated At',
    'Days Since Last Update','Client Status Report URL'
  ];
}
function snapshotHeaders14_() {
  return ['Snapshot Date','Captured At','RootApptID','Rep','Role'];
}

/** Take today's morning snapshot (manual or scheduled).
 *  - Reads 07 + 08
 *  - Applies schedule + Maria↔Paul coverage via computeExpectedSetsWithSchedule_()
 *  - Writes frozen expected set into 13_Morning_Snapshot (replacing prior content)
 *  - Appends audit rows to 14_Snapshot_Log
 */
function takeMorningSnapshot() {
  const ss = SpreadsheetApp.getActive();
  const tz = (typeof TIMEZONE !== 'undefined' && TIMEZONE) ? TIMEZONE : ss.getSpreadsheetTimeZone();

  const s07 = getSheetOrThrow_('07_Root_Index');
  const s08 = getSheetOrThrow_('08_Reps_Map');

  // canonical target sheets
  const sSnap = getSheetOrThrow_('13_Morning_Snapshot');
  const sLog  = getSheetOrThrow_('14_Snapshot_Log');

  // 1) Build lookups needed to determine Role for a (root,rep)
  const map08 = getObjects_(s08);
  const roleByRootRep = new Map();
  map08.forEach(r => {
    const root = String(r['RootApptID'] || '').trim();
    const rep  = String(r['Rep'] || '').trim();
    const role = String(r['Role (Assigned/Assisted)'] || '').trim();
    if (root && rep && role) roleByRootRep.set(`${root}||${rep}`, role);
  });

  // 2) Use the policy engine to classify and compute expected sets (policy + schedule + coverage)
  const idx07 = getObjects_(s07);
  const {
    scopeGroupByRoot,
    snapByRoot,
    expectedByRepDuty
  } = computeExpectedSetsWithPolicies_(idx07, map08);

  // 3) Prepare today keys
  const now = new Date();
  const dateKey    = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const capturedAt = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');

  // 4) Canonical header set for both 10 & 11 (must match exactly, width must match rows)
  const SNAP_HEADERS = [
    'Snapshot Date','Captured At','RootApptID','Rep','Role','Scope Group',
    'Customer Name','Sales Stage','Conversion Status','Custom Order Status',
    'Updated By','Updated At','Days Since Last Update','Client Status Report URL'
  ];

  // === write 10_Morning_Snapshot (REPLACE CONTENTS, force header) ===
  healSnapshotSheet_(sSnap, SNAP_HEADERS);  // clears contents, formats, banding, sets exact headers

  // Build rows in policy priority order (current expected on-duty pairs)
  const out = [];
  expectedByRepDuty.forEach((rootsSet, rep) => {
    rootsSet.forEach(root => {
      const r07  = snapByRoot.get(root) || {};
      const role = roleByRootRep.get(`${root}||${rep}`) || '';  // 'Assigned'|'Assisted'
      const group = scopeGroupByRoot.get(root) || '';
      out.push([
        dateKey, capturedAt, root, rep, role, group,
        String(r07['Customer Name'] || '').trim(),
        r07['Sales Stage'] || '',
        r07['Conversion Status'] || '',
        r07['Custom Order Status'] || '',
        r07['Updated By'] || '',
        r07['Updated At'] || '',
        r07['Days Since Last Update'] || '',
        r07['Client Status Report URL'] || ''
      ]);
    });
  });

  if (out.length) {
    sSnap.getRange(2, 1, out.length, SNAP_HEADERS.length).setValues(out);
    // datetime formatting
    const colCaptured = SNAP_HEADERS.indexOf('Captured At') + 1;
    const colUpdAt    = SNAP_HEADERS.indexOf('Updated At')   + 1;
    sSnap.getRange(2, colCaptured, out.length, 1).setNumberFormat('yyyy-mm-dd HH:mm:ss');
    sSnap.getRange(2, colUpdAt,    out.length, 1).setNumberFormat('yyyy-mm-dd HH:mm:ss');
  }

  // === append to 11_Snapshot_Log with same shape ===
  ensureSnapshotLogHeader_(sLog, SNAP_HEADERS);  // writes header if missing or wrong width
  if (out.length) {
    const start = sLog.getLastRow() + 1;
    sLog.getRange(start, 1, out.length, SNAP_HEADERS.length).setValues(out);
    const colCaptured = SNAP_HEADERS.indexOf('Captured At') + 1;
    sLog.getRange(start, colCaptured, out.length, 1).setNumberFormat('yyyy-mm-dd HH:mm:ss');
  }
}


/** Read today's snapshot from 13_Morning_Snapshot.
 * Returns:
 *  - snapExpectedByRep: Map<Rep, Set<Root>>
 *  - captureTime: Date or null if none for today
 */
function readTodaySnapshot_() {
  const ss = SpreadsheetApp.getActive();
  const tz = TZ_SNAPSHOT || ss.getSpreadsheetTimeZone();
  const todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  const sh13 = getSheetOrThrow_(SHEET_13);
  const rows = getObjects_(sh13);
  const snapExpectedByRep = new Map();
  let captureTime = null;

  rows.forEach(r => {
    const snapVal = r['Snapshot Date'];
    let snapKey = '';
    if (snapVal instanceof Date) {
      snapKey = Utilities.formatDate(snapVal, tz, 'yyyy-MM-dd');
    } else {
      snapKey = String(snapVal).slice(0,10); // fallback
    }
    if (snapKey !== todayKey) return;
    const rep = String(r['Rep'] || '').trim();
    const root = String(r['RootApptID'] || '').trim();
    const cap = toDateSafe_(r['Captured At']);
    if (cap && (!captureTime || cap.getTime() > captureTime.getTime())) captureTime = cap;

    if (!rep || !root) return;
    if (!snapExpectedByRep.has(rep)) snapExpectedByRep.set(rep, new Set());
    snapExpectedByRep.get(rep).add(root);
  });

  return { snapExpectedByRep, captureTime };
}

/** Ensure sheet exists with the given headers (repair if headers differ) */
function ensureSheetWithHeaders_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  const have = (sh.getLastColumn() ? sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0] : []);
  const cur = have.map(v => String(v || '').trim());
  const want = headers.map(h => String(h || '').trim());

  if (cur.join('\u0001') !== want.join('\u0001')) {
    sh.clearContents();
    sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Clear a snapshot sheet to exactly the given headers; remove old banding/merges/formats (robust) */
function healSnapshotSheet_(sh, headers) {
  if (!sh) throw new Error('healSnapshotSheet_: sheet is null');
  // 0) Break merges safely (on the used range only — avoids OOB)
  try {
    const used = sh.getDataRange(); // always at least 1x1
    used.breakApart();
  } catch (_) {}

  // 1) Clear contents & formats, and remove any banding
  sh.clearContents();
  sh.clearFormats();
  try {
    const bandings = sh.getBandings();
    if (bandings && bandings.length) bandings.forEach(b => { try { b.remove(); } catch (_) {} });
  } catch (_) {}

  // 2) Force exact column count = headers.length
  //    (guard for zero/negative counts to avoid out-of-bounds)
  const want = Math.max(1, Number(headers.length || 0));
  let have = Math.max(1, sh.getLastColumn());   // Sheets guarantees >=1 after clear

  if (have > want) {
    // Delete only when count > 0
    const delCnt = have - want;
    if (delCnt > 0) sh.deleteColumns(want + 1, delCnt);
    have = sh.getLastColumn();
  } else if (have < want) {
    const addCnt = want - have;
    // If there are 0/1 columns, insert starting at col 2 or at 1 as needed
    if (have <= 1) {
      // Ensure at least 1 column exists; then insert remaining after it
      if (have === 0) sh.insertColumns(1, 1);
      if (addCnt > 0) sh.insertColumnsAfter(1, addCnt);
    } else {
      sh.insertColumnsAfter(have, addCnt);
    }
    have = sh.getLastColumn();
  }

  // 3) Write canonical header across exactly 'want' columns
  sh.getRange(1, 1, 1, want).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
}

/** Ensure the log sheet has the same header set; repair width only (no content clear) */
function ensureSnapshotLogHeader_(sh, headers) {
  if (!sh) throw new Error('ensureSnapshotLogHeader_: sheet is null');

  const want = Math.max(1, Number(headers.length || 0));
  const have = Math.max(1, sh.getLastColumn());

  // If width mismatches, fix it first
  if (have > want) {
    const delCnt = have - want;
    if (delCnt > 0) sh.deleteColumns(want + 1, delCnt);
  } else if (have < want) {
    const addCnt = want - have;
    if (have <= 1) {
      if (have === 0) sh.insertColumns(1, 1);
      if (addCnt > 0) sh.insertColumnsAfter(1, addCnt);
    } else {
      sh.insertColumnsAfter(have, addCnt);
    }
  }

  // Read current header (if any)
  let headerNow = [];
  try { headerNow = sh.getRange(1, 1, 1, want).getValues()[0] || []; } catch (_) { headerNow = []; }

  const same = headerNow.length === want &&
               headerNow.every((v, i) => String(v || '').trim() === String(headers[i]));

  if (!same) {
    sh.getRange(1, 1, 1, want).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
}


function ack_healSnapshotSheetsOnce() {
  const s13 = getSheetOrThrow_('13_Morning_Snapshot');
  const s14 = getSheetOrThrow_('14_Snapshot_Log');

  const HDR = [
    'Snapshot Date','Captured At','RootApptID','Rep','Role','Scope Group',
    'Customer Name','Sales Stage','Conversion Status','Custom Order Status',
    'Updated By','Updated At','Days Since Last Update','Client Status Report URL'
  ];

  healSnapshotSheet_(s13, HDR);
  ensureSnapshotLogHeader_(s14, HDR);

  Logger.log('Healed 13_Morning_Snapshot and 14_Snapshot_Log headers/columns. Now re-run takeMorningSnapshot().');
}


/** Create a daily trigger at ~8:30 AM PT to run takeMorningSnapshot() */
function createSnapshotTrigger() {
  // Remove existing triggers for this handler first
  removeSnapshotTrigger();

  const tz = TZ_SNAPSHOT;
  const hourPT = 8;   // 8 AM
  const minute = 30;  // :30

  ScriptApp.newTrigger('takeMorningSnapshot')
    .timeBased()
    .atHour(hourPT)
    .nearMinute(minute)
    .everyDays(1)
    .create();

  SpreadsheetApp.getUi().alert('Daily snapshot trigger created for ~8:30 AM PT.\n(Uses the spreadsheet\'s timezone settings.)');
}

/** Remove the daily snapshot trigger */
function removeSnapshotTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'takeMorningSnapshot') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // no alert to keep it quiet if none existed
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



