/*****  Phase B — Data Pipes for Acknowledgement Workflows  *****
 * Builds:
 *  - 07_Root_Index (one row per RootApptID, canonical today state)
 *  - 08_Reps_Map (normalized Root ↔ Rep ↔ Role, with Include? logic)
 *  - Recompute Ack Status on 00_ from 06_Acknowledgement_Log (optional in Phase B)
 *
 * Assumptions:
 *  - 00_ sheet exists with the exact headers specified below.
 *  - 06/07/08 sheets exist (created in Phase A) — this script will overwrite their bodies.
 *  - Named ranges exist on 'Dropdown': Roster_Assigned_Reps, Roster_Assisted_Reps, Ack_Status_List.
 ****************************************************************/

//////////////////////////
// === CONFIG START === //
//////////////////////////

/** Sheet/tab names **/
const SHEET_00   = '00_Master Appointments';         // <- change if your 00_ tab has a different name
const SHEET_06   = '06_Acknowledgement_Log';
const SHEET_07   = '07_Root_Index';
const SHEET_08   = '08_Reps_Map';
const SHEET_09 = '09_Ack_Dashboard';
const SHEET_10 = '10_Roster_Schedule'; // schedule lives here
const SHEET_13 = '13_Morning_Snapshot';
const SHEET_14 = '14_Snapshot_Log';

// === CONFIG: Dropdowns ===
const DROPDOWN_SHEET = 'Dropdown';     // ← tab label in your spreadsheet
const NAMED_RANGE_ASSIGNED_REPS = 'AssignedReps';  // if you have a named range
const NAMED_RANGE_ASSISTED_REPS = 'AssistedReps';  // if you have a named range

const DROPDOWN_HEADERS = {
  ASSIGNED_REP:   'Assigned Rep',        // exact header text on Dropdown
  ASSIGNED_EMAIL: 'Assigned Rep Email',  // adjust if your sheet header differs
  ASSISTED_REP:   'Assisted Rep',
  ASSISTED_EMAIL: 'Assisted Rep Email'
};


/** Timezone & day boundary **/
const TIMEZONE   = 'America/Los_Angeles';

/** Exact header names on 00_ (copy/paste from your sheet) **/
const H = {
  ROOT_ID:                 'RootApptID',
  CUSTOMER:                'Customer Name',
  SALES_STAGE:             'Sales Stage',
  CONVERSION_STATUS:       'Conversion Status',
  CUSTOM_ORDER_STATUS:     'Custom Order Status',
  IN_PROD_STATUS:          'In Production Status',
  CS_ORDER_STATUS:         'Center Stone Order Status',
  NEXT_STEPS:              'Next Steps',
  ASSIGNED_REP:            'Assigned Rep',
  ASSISTED_REP:            'Assisted Rep',
  UPDATED_BY:              'Updated By',
  UPDATED_AT:              'Updated At',
  CLIENT_STATUS_URL:       'Client Status Report URL',
  ACK_STATUS:              'Ack Status' // summary column on 00_ (fully updated / follow up)
};

/** Named ranges on Dropdown **/
const NR = {
  ROSTER_ASSIGNED: 'Roster_Assigned_Reps',
  ROSTER_ASSISTED: 'Roster_Assisted_Reps',
  ACK_STATUS_LIST: 'Ack_Status_List'
};

/** Constants/labels **/
const LABELS = {
  FULLY_UPDATED:    'Fully Updated',
  NEEDS_FOLLOW_UP:  'Needs follow-up',
  INCLUDE_Y:        'Y',
  INCLUDE_N:        'N',
  ROLE_ASSIGNED:    'Assigned',
  ROLE_ASSISTED:    'Assisted'
};

/** Values in Sales Stage that exclude a root entirely from maps/queues **/
const EXCLUDE_SALES_STAGES = new Set(['Won', 'Lost Lead']);
const EXCLUDE_SALES_STAGES_LC = new Set(['won','lost lead']); // <- add this
const EXCLUDE_SALES_STAGES_FALLBACK = new Set(['Won','Lost Lead']);

/** The literal inclusion test for queues (Phase B/now): Custom Order Status == 'In Production' (case-insensitive) **/
const IN_PRODUCTION_LITERAL = 'In Production';

////////////////////////
// === CONFIG END === //
////////////////////////

/** ===========================
 * STYLE THEME (Ack + Reminders)
 * Base colors + two standard tints
 * =========================== */
const STYLE_THEME = {
  STAGE_COLORS: {
    // ACK section groups (Sales Stages)
    'Appointment':         '#AECBFA',
    'Viewing Scheduled':   '#FFD7C2',
    'Hot Lead':            '#D93025',
    'Follow-Up Required':  '#C5221F',
    'In Production':       '#C8E6C9',

    // Reminder section groups (by subsection header text)
    'Custom Order Update Needed': '#E6CFF2',
    'DV_URGENT_OTW_DAILY':        '#C5221F',
    'DV_URGENT':                  '#C5221F',

    // 👇 ADD these three aliases (exact strings) so any label variant resolves to peach
    'FollowUp':                   '#FFD7C2',
    'FOLLOWUP':                   '#FFD7C2',
    'Need to Follow-Up':          '#FFD7C2',

    // Fallbacks
    REMINDERS:                    '#C5221F',
    ACK_DEFAULT:                  '#AECBFA'
  },

  TINT_1: 0.86,
  TINT_2: 0.92
};

/** === Tiny color helpers (pure) === */
function _hexToRgb_(hex) {
  let s = String(hex || '').replace('#', '').trim();
  if (s.length === 3) s = s.split('').map(c => c + c).join('');
  const n = parseInt(s, 16);
  if (isNaN(n)) return { r: 240, g: 240, b: 240 }; // safe fallback
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
}
function _rgbToHex_(r, g, b) {
  const to2 = v => ('0' + Math.max(0, Math.min(255, Math.round(v))).toString(16)).slice(-2);
  return ('#' + to2(r) + to2(g) + to2(b)).toUpperCase();
}
/** Blend a → b by t in [0..1] */
function _blend_(hexA, hexB, t) {
  const a = _hexToRgb_(hexA), b = _hexToRgb_(hexB);
  const r = a.r * (1 - t) + b.r * t;
  const g = a.g * (1 - t) + b.g * t;
  const b2 = a.b * (1 - t) + b.b * t;
  return _rgbToHex_(r, g, b2);
}
/** Lighten a color by blending toward white using a factor (0..1). */
function _tint_(hex, factor) {
  return _blend_(hex, '#FFFFFF', factor);
}


function runAllPipes() {
  buildRootIndex();
  buildRepsMap();
}


/** ========================
 *   07_Root_Index builder
 *  ======================== */
function buildRootIndex() {
  const ss = SpreadsheetApp.getActive();
  const tz = TIMEZONE || ss.getSpreadsheetTimeZone();
  const s00 = getSheetOrThrow_(SHEET_00);
  const s07 = getSheetOrThrow_(SHEET_07);

  const rows00 = getObjects_(s00);
  if (!rows00.length) {
    clearAndWrite_(s07, rootIndexHeaders_(), []);
    return;
  }

  // Group rows by RootApptID
  const byRoot = new Map();
  rows00.forEach(r => {
    const root = String(r[H.ROOT_ID] || '').trim();
    if (!root) return;
    if (!byRoot.has(root)) byRoot.set(root, []);
    byRoot.get(root).push(r);
  });

  const out = [];
  const now = new Date();

  byRoot.forEach((rows, root) => {
    // Pick canonical row = max Updated At
    let canonical = rows[0];
    let maxAt = toDateSafe_(rows[0][H.UPDATED_AT]);
    rows.forEach(rr => {
      const at = toDateSafe_(rr[H.UPDATED_AT]);
      if (at && (!maxAt || at.getTime() > maxAt.getTime())) {
        maxAt = at;
        canonical = rr;
      }
    });

    // EXCLUDE Won / Lost Lead entirely from 07
    const salesStageCanon = String(canonical[H.SALES_STAGE] || '').trim();
    const salesStageLC = salesStageCanon.toLowerCase();
    if (EXCLUDE_SALES_STAGES.has(salesStageCanon) || EXCLUDE_SALES_STAGES_LC.has(salesStageLC)) {
      return; // skip this root entirely
    }

    const daysSince = (maxAt)
      ? Math.floor((now.getTime() - maxAt.getTime()) / 86400000)
      : '';

    out.push([
      root,
      canonical[H.CUSTOMER] || '',
      salesStageCanon,
      canonical[H.CONVERSION_STATUS] || '',
      canonical[H.CUSTOM_ORDER_STATUS] || '',
      canonical[H.IN_PROD_STATUS] || '',
      canonical[H.CS_ORDER_STATUS] || '',
      canonical[H.NEXT_STEPS] || '',
      canonical[H.UPDATED_BY] || '',
      maxAt || '',
      daysSince,
      canonical[H.CLIENT_STATUS_URL] || ''
    ]);
  });

  // Sort by Days Since Last Update desc, then Customer asc
  const IDX_DAYS = 10;
  const IDX_CUSTOMER = 1;
  out.sort((a,b) => {
    const d = (Number(b[IDX_DAYS])||0) - (Number(a[IDX_DAYS])||0);
    if (d !== 0) return d;
    const ca = (a[IDX_CUSTOMER]||'') + '';
    const cb = (b[IDX_CUSTOMER]||'') + '';
    return ca.localeCompare(cb);
  });

  clearAndWrite_(s07, rootIndexHeaders_(), out);

  // Formatting niceties
  const hdrs = rootIndexHeaders_();
  const colUpdatedAt = hdrs.indexOf('Updated At') + 1;
  if (colUpdatedAt > 0 && out.length) {
    s07.getRange(2, colUpdatedAt, out.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
}

/** =======================
 *   08_Reps_Map builder
 *  ======================= */
function buildRepsMap() {
  const ss = SpreadsheetApp.getActive();
  const s00 = getSheetOrThrow_(SHEET_00);
  const s07 = getSheetOrThrow_(SHEET_07);
  const s08 = getSheetOrThrow_(SHEET_08);

  const rows00 = getObjects_(s00);
  const idx07  = getObjects_(s07); // we’ll prefer canonical Sales Stage from 07 if available

  // Build quick lookups
  const salesStageByRoot = new Map();
  idx07.forEach(r => {
    const root = String(r['RootApptID'] || '').trim();
    if (root) salesStageByRoot.set(root, String(r['Sales Stage'] || '').trim());
  });

  // Roster normalization maps
  const norm = buildRosterNormalizer_();

  // Collect map rows: dedupe by (root|rep), role precedence Assigned > Assisted
  const mapByKey = new Map();

  rows00.forEach(r => {
    const root = String(r[H.ROOT_ID] || '').trim();
    if (!root) return;

    // Determine Sales Stage (prefer 07’s canonical if present)
    let salesStage = salesStageByRoot.get(root);
    if (!salesStage) salesStage = String(r[H.SALES_STAGE] || '').trim();

    // Parse multi-select reps
    const assignedNames = parseMultiNames_(r[H.ASSIGNED_REP]).map(n => norm(n));
    const assistedNames = parseMultiNames_(r[H.ASSISTED_REP]).map(n => norm(n));

    // Build candidate sets
    const assignedSet = new Set(assignedNames.filter(Boolean));
    const assistedSet = new Set(assistedNames.filter(Boolean));

    // Union both roles; role precedence: Assigned wins
    const all = new Set([...assignedSet, ...assistedSet]);
    all.forEach(rep => {
      const role = assignedSet.has(rep) ? LABELS.ROLE_ASSIGNED : LABELS.ROLE_ASSISTED;
      const key  = `${root}||${rep}`;
      const existing = mapByKey.get(key);

      if (!existing) {
        mapByKey.set(key, {
          root, rep, role,
          salesStage
        });
      } else {
        // Upgrade role to Assigned if needed
        if (existing.role !== LABELS.ROLE_ASSIGNED && role === LABELS.ROLE_ASSIGNED) {
          existing.role = LABELS.ROLE_ASSIGNED;
        }
        // Prefer canonical sales stage from 07 if any row had it
        if (!salesStageByRoot.has(root)) {
          existing.salesStage = salesStage;
        }
      }
    });
  });

  // Materialize with Include? and Last Sync
  const now = new Date();
  const out = [];
  [...mapByKey.values()].forEach(x => {
    const include = EXCLUDE_SALES_STAGES.has((x.salesStage || '').trim()) ? LABELS.INCLUDE_N : LABELS.INCLUDE_Y;
    out.push([
      x.root,
      x.rep,
      x.role,
      x.salesStage || '',
      include,
      now
    ]);
  });

  // Sort by Root, then Role (Assigned first), then Rep
  const ROLE_ORDER = { 'Assigned': 0, 'Assisted': 1 };
  out.sort((a,b) => {
    const r = String(a[0]).localeCompare(String(b[0]));
    if (r !== 0) return r;
    const ro = (ROLE_ORDER[a[2]] ?? 9) - (ROLE_ORDER[b[2]] ?? 9);
    if (ro !== 0) return ro;
    return String(a[1]).localeCompare(String(b[1]));
  });

  clearAndWrite_(s08, repsMapHeaders_(), out);

  // Format Last Sync as datetime
  if (out.length) {
    const colLastSync = repsMapHeaders_().indexOf('Last Sync') + 1;
    s08.getRange(2, colLastSync, out.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
}


/** =======================================================
 *  Recompute Ack Status on 00_ using 06_Acknowledgement_Log
 *  - Root is in-scope if Custom Order Status == 'In Production' AND Sales Stage not in {Won, Lost Lead}
 *  - Fully Updated only if every expected rep (Include?=Y in 08) has a 'Fully Updated' ack TODAY
 *  - Needs follow-up otherwise
 *  ======================================================= */
function recomputeAckStatusSummary() {
  const ss = SpreadsheetApp.getActive();
  const tz = (typeof TIMEZONE !== 'undefined' && TIMEZONE) ? TIMEZONE : ss.getSpreadsheetTimeZone();

  // Sheets we read/write
  const s00 = getSheetOrThrow_(SHEET_00);
  const s06 = getSheetOrThrow_(SHEET_06);
  const s07 = getSheetOrThrow_(SHEET_07);
  const s08 = getSheetOrThrow_(SHEET_08);

  // Pull data once
  const idx07 = getObjects_(s07);
  const map08 = getObjects_(s08);
  const log06 = getObjects_(s06);

  // Policies + today’s on‑duty sets (and roles)
  const {
    policies,
    scopeGroupByRoot,
    expectedByRootDuty,
    roleByRootRepDuty
  } = computeExpectedSetsWithPolicies_(idx07, map08);

  // Gate to only groups that the queues include (your rule: “rows pulled into the Ack queues”)
  const queueGroups = new Set(policies.filter(p => p.queueInclude).map(p => p.group));
  const mustAckByGroup = new Map(policies.map(p => [p.group, String(p.mustAck || 'ALL_ON_DUTY').toUpperCase()]));

  // Latest ACK status today per (root||rep)
  const todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const latestAckByRootRep = new Map(); // `${root}||${rep}` -> {status, ts}
  log06.forEach(r => {
    const root = String(r['RootApptID'] || '').trim();
    const rep  = String(r['Rep'] || '').trim();
    const status = String(r['Ack Status'] || '').trim();
    if (!root || !rep || !status) return;

    const logDateVal = r['Log Date'];
    const logDateKey = (logDateVal instanceof Date)
      ? Utilities.formatDate(logDateVal, tz, 'yyyy-MM-dd')
      : String(logDateVal || '').slice(0,10);
    if (logDateKey !== todayKey) return;

    const ts = toDateSafe_(r['Timestamp']) || new Date();
    const k = `${root}||${rep}`;
    const prev = latestAckByRootRep.get(k);
    if (!prev || ts.getTime() > prev.ts.getTime()) latestAckByRootRep.set(k, { status, ts });
  });

  // Build summary per root following MustAck
  const summaryByRoot = new Map();       // root -> 'Fully Updated' | 'Needs follow-up'
  const requiredExists = new Set();      // roots that have a non-empty required set today

  expectedByRootDuty.forEach((onDutySet, root) => {
    // Only roots that are in queue-enabled policy groups are summarized; others left blank
    const group = scopeGroupByRoot.get(root);
    if (!queueGroups.has(group)) return;

    // Compute required reps by MustAck
    const must = (mustAckByGroup.get(group) || 'ALL_ON_DUTY').replace(/[ \-]+/g,'_').toUpperCase();
    let required;
    if (must === 'ASSISTED_REPS_ONLY') {
      required = new Set([...onDutySet].filter(rep => roleByRootRepDuty.get(`${root}||${rep}`) === 'Assisted'));
    } else if (must === 'ASSIGNED_REPS_ONLY') {
      required = new Set([...onDutySet].filter(rep => roleByRootRepDuty.get(`${root}||${rep}`) === 'Assigned'));
    } else { // default: ALL_ON_DUTY
      required = onDutySet;
    }

    if (!required || required.size === 0) {
      // No required reps on duty today for this root → leave blank on 00_
      return;
    }
    requiredExists.add(root);

    // Aggregate today
    let allOK = true, anyNeeds = false;
    required.forEach(rep => {
      const rec = latestAckByRootRep.get(`${root}||${rep}`);
      if (!rec) { allOK = false; return; }
      const st = String(rec.status || '').trim();
      if (equalsIgnoreCase_(st, LABELS.NEEDS_FOLLOW_UP)) anyNeeds = true;
      if (!equalsIgnoreCase_(st, LABELS.FULLY_UPDATED))  allOK = false;
    });

    const summary = (allOK && !anyNeeds) ? LABELS.FULLY_UPDATED : LABELS.NEEDS_FOLLOW_UP;
    summaryByRoot.set(root, summary);
  });

  // Write back to 00_: required roots get summary (default to Needs follow-up if missing),
  // non-required or out-of-policy roots are blanked.
  const headers = getHeaders_(s00);
  const idxAck  = headers.indexOf(H.ACK_STATUS);
  const idxRoot = headers.indexOf(H.ROOT_ID);
  if (idxAck < 0 || idxRoot < 0) throw new Error('Ack Status or RootApptID column not found on 00_.');

  const dataRange = s00.getDataRange();
  const values = dataRange.getValues(); // includes header
  for (let i = 1; i < values.length; i++) {
    const root = String(values[i][idxRoot] || '').trim();
    if (!root) { values[i][idxAck] = ''; continue; }

    const group = scopeGroupByRoot.get(root);
    if (!queueGroups.has(group)) { values[i][idxAck] = ''; continue; }

    if (requiredExists.has(root)) {
      // If we computed a summary for this root, use it; otherwise, default to Needs follow-up
      values[i][idxAck] = summaryByRoot.has(root) ? summaryByRoot.get(root) : LABELS.NEEDS_FOLLOW_UP;
    } else {
      // No required reps on duty today for this root → blank
      values[i][idxAck] = '';
    }
  }
  dataRange.setValues(values);
}


/** =========================
 * Helpers & small utilities
 * ========================= */

function rootIndexHeaders_() {
  return [
    'RootApptID',
    'Customer Name',
    'Sales Stage',
    'Conversion Status',
    'Custom Order Status',
    'In Production Status',
    'Center Stone Order Status',
    'Next Steps',
    'Updated By',
    'Updated At',
    'Days Since Last Update',
    'Client Status Report URL'
  ];
}

function repsMapHeaders_() {
  return [
    'RootApptID',
    'Rep',
    'Role (Assigned/Assisted)',
    'Sales Stage',
    'Include? (Y/N)',
    'Last Sync'
  ];
}

/** Ensure sheet exists; throw if missing to prevent silent failure */
function getSheetOrThrow_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error(`Sheet "${name}" not found.`);
  return sh;
}

/** Read sheet rows as array of {header:value} objects */
function getObjects_(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (!values.length) return [];
  const headers = values.shift().map(h => String(h || '').trim());
  return values.map(row => {
    const o = {};
    headers.forEach((h, i) => o[h] = row[i]);
    return o;
  }).filter(o => Object.keys(o).length > 0);
}

function getHeaders_(sheet) {
  const range = sheet.getRange(1,1,1,sheet.getLastColumn());
  return range.getValues()[0].map(h => String(h || '').trim());
}

/** Clear body and write new data with headers */
function clearAndWrite_(sheet, headers, rows) {
  sheet.clearContents();
  if (!headers || !headers.length) return;
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  if (rows && rows.length) {
    sheet.getRange(2,1,rows.length,headers.length).setValues(rows);
  }
  sheet.setFrozenRows(1);
}

/** Multi-select parsing: splits on comma or newline; trims spaces */
function parseMultiNames_(cell) {
  if (!cell) return [];
  const txt = String(cell).trim();
  if (!txt) return [];
  return txt.split(/[,|\n]/).map(s => s.trim()).filter(Boolean);
}

/** Build case-insensitive normalizer using both roster named ranges; returns fn(name)->canonical */
function buildRosterNormalizer_() {
  const ss = SpreadsheetApp.getActive();

  const lowerToCanon = new Map();
  [NR.ROSTER_ASSIGNED, NR.ROSTER_ASSISTED].forEach(nr => {
    const range = ss.getRangeByName(nr);
    if (!range) return;
    const vals = range.getValues().flat().map(v => String(v || '').trim()).filter(Boolean);
    vals.forEach(v => lowerToCanon.set(v.toLowerCase(), v));
  });

  // Return normalizer function; if unknown, keep trimmed input
  return function normalize(name) {
    if (!name) return '';
    const key = String(name).trim().toLowerCase();
    return lowerToCanon.get(key) || String(name).trim();
  };
}

/** Safe date parsing */
function toDateSafe_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const n = Number(v);
  if (!isNaN(n) && n > 0) return new Date(n);
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function equalsIgnoreCase_(a, b) {
  return String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
}

/** ============ Optional: Log Appender (for Phase C use) ============ */
/** appendAckLog_ — appends a single ack row to 06_Acknowledgement_Log
 *  This is provided now so Phase C can call it.
 */
function appendAckLog_(payload) {
  // payload: {root, rep, role, ackStatus, ackNote, ackBy, snapshot}
  // snapshot fields should mirror 07_Root_Index columns at time of logging
  const s06 = getSheetOrThrow_(SHEET_06);
  const headers = getHeaders_(s06);
  const row = [];

  const tz = TIMEZONE || SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const now = new Date();
  const logDate = Utilities.formatDate(now, tz, 'yyyy-MM-dd'); // write as text date to avoid locale shift

  const H06 = {
    LOG_DATE: 'Log Date',
    TS: 'Timestamp',
    ROOT: 'RootApptID',
    REP: 'Rep',
    ROLE: 'Role',
    ACK_STATUS: 'Ack Status',
    ACK_NOTE: 'Ack Note',
    ACK_BY: 'Ack By (Email/Name)',
    CUSTOMER: 'Customer (at log)',
    SALES_STAGE: 'Sales Stage (at log)',
    CONVERSION: 'Conversion Status (at log)',
    COS: 'Custom Order Status (at log)',
    INPROD: 'In Production Status (at log)',
    CSOS: 'Center Stone Order Status (at log)',
    NEXT: 'Next Steps (at log)',
    UPD_BY: 'Last Updated By (at log)',
    UPD_AT: 'Last Updated At (at log)',
    URL: 'Client Status Report URL'
  };

  const obj = {};
  obj[H06.LOG_DATE] = logDate;
  obj[H06.TS]       = now;
  obj[H06.ROOT]     = payload.root || '';
  obj[H06.REP]      = payload.rep || '';
  obj[H06.ROLE]     = payload.role || '';
  obj[H06.ACK_STATUS]= payload.ackStatus || '';
  obj[H06.ACK_NOTE] = payload.ackNote || '';
  obj[H06.ACK_BY]   = payload.ackBy || '';

  const snap = payload.snapshot || {};
  obj[H06.CUSTOMER]   = snap['Customer Name'] || '';
  obj[H06.SALES_STAGE]= snap['Sales Stage'] || '';
  obj[H06.CONVERSION] = snap['Conversion Status'] || '';
  obj[H06.COS]        = snap['Custom Order Status'] || '';
  obj[H06.INPROD]     = snap['In Production Status'] || '';
  obj[H06.CSOS]       = snap['Center Stone Order Status'] || '';
  obj[H06.NEXT]       = snap['Next Steps'] || '';
  obj[H06.UPD_BY]     = snap['Updated By'] || '';
  obj[H06.UPD_AT]     = snap['Updated At'] || '';
  obj[H06.URL]        = snap['Client Status Report URL'] || '';

  const ordered = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  s06.appendRow(ordered);

  // Format Timestamp column
  const tsCol = headers.indexOf(H06.TS) + 1;
  const lastRow = s06.getLastRow();
  if (tsCol > 0) s06.getRange(lastRow, tsCol).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

/***** Phase C — Per‑Rep Queues + Submit Flow *****
 * Features:
 *  - Build Today’s per‑rep queue tabs (pending only)
 *  - My Queue (Detect Me) / Refresh / Submit
 *  - Submission appends to 06_Acknowledgement_Log and recomputes 00_ Ack Status
 *
 * Requires Phase B:
 *  - 06_Acknowledgement_Log, 07_Root_Index, 08_Reps_Map tabs exist
 *  - appendAckLog_(), recomputeAckStatusSummary(), and helpers are available
 ****************************************************************/

//////////////////////////
// === CONFIG START === //
//////////////////////////

// Queue sheets prefix
const QUEUE_PREFIX = 'Q_';




// ======== PUBLIC MENU ACTIONS (wired via onOpen in ack_pipes.gs) ========

/** Build per‑rep queue tabs for all reps with pending items today */
function buildTodaysQueuesAll() {
  const {expectedByRep, roleByRootRep, snapByRoot} = computeExpectedToday_();
  const reps = [...expectedByRep.keys()];

  if (!reps.length) {
    Logger.log('[ACK] No pending items today for any rep.'); // ⬅️ was SpreadsheetApp.getUi().alert(...)
    return;
  }

  reps.forEach(rep => {
    buildQueueForRep_(rep, expectedByRep.get(rep), snapByRoot);
  });

  Logger.log(`[ACK] Built queues for ${reps.length} reps.`); // ⬅️ was SpreadsheetApp.getUi().alert(...)
}

/** Detect current user and open their queue tab (creates if needed) */
function openMyQueue() {
  const rep = detectRepName_();
  if (!rep) return;
  openQueueForRep_(rep);
}

/** Detect current user and refresh their queue (rebuild pending only) */
function refreshMyQueue() {
  // Keep behavior: detect rep → build my ack queue
  var rep = detectRepName_();
  if (!rep) return;

  var payload = computeExpectedToday_(); // harmless; used by build below
  var roots = payload.expectedByRep.get(rep) || new Set();
  buildQueueForRep_(rep, roots, payload.snapByRoot);
}


/** Detect current user and submit their acks from the queue */
function submitMyQueue() {
  const rep = detectRepName_();
  if (!rep) return;
  submitQueueForRep_(rep);
}


// ==================== CORE BUILD / SUBMIT LOGIC ====================

/** Compute today's expected roots per rep, along with snapshots and roles */
function computeExpectedToday_() {
  const ss  = SpreadsheetApp.getActive();
  const tz  = (typeof TIMEZONE !== 'undefined' && TIMEZONE) ? TIMEZONE : ss.getSpreadsheetTimeZone();

  const s06 = getSheetOrThrow_('06_Acknowledgement_Log');
  const s07 = getSheetOrThrow_('07_Root_Index');
  const s08 = getSheetOrThrow_('08_Reps_Map');

  // Pull sheets once
  const idx07 = getObjects_(s07);
  const map08 = getObjects_(s08);
  const log06 = getObjects_(s06);

  // === NEW: use policy engine (all scope groups) + schedule/coverage ===
  // returns: { policies, inScope, scopeGroupByRoot, snapByRoot, nameByRoot,
  //            expectedByRootDuty, expectedByRepDuty, assignedGaps, assistedGaps }
  const {
    inScope,
    snapByRoot,
    expectedByRepDuty
  } = computeExpectedSetsWithPolicies_(idx07, map08);

  // Latest ack per rep for TODAY by root (to filter pending)
  const todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const latestAckByRootRep = new Map(); // key: root||rep -> {status, ts}

  log06.forEach(r => {
    const root = String(r['RootApptID'] || '').trim();
    const rep  = String(r['Rep'] || '').trim();
    const status = String(r['Ack Status'] || '').trim();
    const logDateVal = r['Log Date'];
    if (!root || !rep || !status) return;

    // ignore logs that are not today
    const logDateKey = (logDateVal instanceof Date)
      ? Utilities.formatDate(logDateVal, tz, 'yyyy-MM-dd')
      : String(logDateVal).slice(0,10);
    if (logDateKey !== todayKey) return;

    // ignore logs for roots that are not currently in-scope per policies
    if (!inScope.has(root)) return;

    const ts = toDateSafe_(r['Timestamp']) || new Date();
    const key = `${root}||${rep}`;
    const prev = latestAckByRootRep.get(key);
    if (!prev || ts.getTime() > prev.ts.getTime()) {
      latestAckByRootRep.set(key, { status, ts });
    }
  });

  // Build pending-only per-rep sets (hide roots already acknowledged today by that rep)
  const pendingByRep = new Map(); // rep -> Set(root)
  expectedByRepDuty.forEach((rootSet, rep) => {
    rootSet.forEach(root => {
      const key = `${root}||${rep}`;
      if (latestAckByRootRep.has(key)) return; // already acked today
      if (!pendingByRep.has(rep)) pendingByRep.set(rep, new Set());
      pendingByRep.get(rep).add(root);
    });
  });

  // Return shape expected by callers
  return {
    expectedByRep: pendingByRep,  // pending-only set per rep
    roleByRootRep: new Map(),     // not needed for rendering; keep for interface compatibility
    snapByRoot                      // used to build per-row details in the queue
  };
}


/** Build (or rebuild) a single rep’s queue sheet with pending items only */
function buildQueueForRep_(rep, rootsSet, snapByRoot) {
  const sh = ensureQueueSheet_(rep);
  const headers = queueHeaders_();

  // Expand the input set of roots into an array
  const roots = rootsSet ? [...rootsSet] : [];

  // If nothing pending, still show a clean header row (and exit)
  if (!roots.length) {
    sh.clearContents();
    sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // Optional: add a friendly empty-state row
    sh.getRange(2,1).setValue('— No pending items today —').setFontStyle('italic');
    // You can early-return here; no DV needed since there's no data rows
    return;
  }

  // Build the data rows from 07 snapshot for this rep's pending roots
  const rows = roots.map(root => {
    const r = snapByRoot.get(root) || {};
    return [
      root,
      r['Customer Name'] || '',
      r['Sales Stage'] || '',
      r['Conversion Status'] || '',
      r['Custom Order Status'] || '',
      r['In Production Status'] || '',
      r['Center Stone Order Status'] || '',
      r['Next Steps'] || '',
      r['Updated By'] || '',
      r['Updated At'] || '',
      r['Days Since Last Update'] || '',
      r['Client Status Report URL'] || '',
      '', // Ack Status (input)
      ''  // Ack Note (input)
    ];
  });

  // Grouped render (writes header + grouped blocks in policy priority order)
  renderQueueGroupedRows_(sh, headers, rows);

  // Formatting niceties
  sh.setFrozenRows(1);

  // Data validation for Ack Status (applies down the sheet; harmless on blank header rows)
  const ackCol = headers.indexOf('Ack Status') + 1;
  if (ackCol > 0) {
    // Determine how many rows currently exist and apply DV to the entire column below header
    const lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      applyAckStatusValidation_(sh, ackCol, lastRow - 1); // from row 2 to lastRow
    }
  }

  // Datetime formatting for Updated At column (if present)
  const colUpdatedAt = headers.indexOf('Updated At') + 1;
  if (colUpdatedAt > 0) {
    const lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      sh.getRange(2, colUpdatedAt, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    }
  }

  // Autosize for readability
  sh.autoResizeColumns(1, headers.length);

  // Ensure header wrap + center alignment is applied
  if (typeof _formatQueueHeaderAndColumns_ === 'function') {
    _formatQueueHeaderAndColumns_(sh);
  }

  // Apply stage-based alternating colors to ACK rows
  if (typeof styleAckSectionsWithTints_ === 'function') {
    styleAckSectionsWithTints_(sh);
  }
}


/** Submit all acks from a rep’s queue (rows with Ack Status filled) */
function submitQueueForRep_(rep) {
  const ss = SpreadsheetApp.getActive();
  const tz = TIMEZONE || ss.getSpreadsheetTimeZone();

  const sh = getQueueSheetOrNull_(rep);
  if (!sh) {
    SpreadsheetApp.getUi().alert(`No queue tab found for ${rep}. Build/Refresh first.`);
    return;
  }

  const headers = getHeaders_(sh);
  const data = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), headers.length).getValues();

  const idxRoot = headers.indexOf('RootApptID');
  const idxAck  = headers.indexOf('Ack Status');
  const idxNote = headers.indexOf('Ack Note');

  if (idxRoot < 0 || idxAck < 0 || idxNote < 0) {
    throw new Error('Queue headers missing RootApptID/Ack Status/Ack Note.');
  }

  // Lookups from 08 + 07 for role/snapshot at time of log
  const s08   = getSheetOrThrow_(SHEET_08);
  const s07   = getSheetOrThrow_(SHEET_07);
  const map08 = getObjects_(s08);
  const idx07 = getObjects_(s07);

  const roleByRootRep = new Map();
  map08.forEach(r => {
    const root    = String(r['RootApptID'] || '').trim();
    const include = String(r['Include? (Y/N)'] || r['Include?'] || '').trim().toUpperCase();
    if (include !== 'Y') return;
    const repName = String(r['Rep'] || '').trim();
    const role    = String(r['Role (Assigned/Assisted)'] || '').trim() || 'Assigned';
    if (root && repName) roleByRootRep.set(`${root}||${repName}`, role);
  });

  const snapByRoot = new Map();
  idx07.forEach(r => {
    const root = String(r['RootApptID'] || '').trim();
    if (root) snapByRoot.set(root, r);
  });

  // Validate & collect payloads
  const payloads = [];
  const missingNotes = []; // collect roots that chose Needs follow-up but no note

  for (let i = 0; i < data.length; i++) {
    const row  = data[i];
    const root = String(row[idxRoot] || '').trim();
    const ack  = String(row[idxAck]  || '').trim();
    const note = String(row[idxNote] || '').trim();

    if (!root) continue;     // skip empty row
    if (!ack)  continue;     // nothing selected → skip this row

    // Skip Reminder rows here; Phase‑2 handles them (and also logs to 06)
    var idxReminderId = headers.indexOf('Reminder ID');
    if (idxReminderId >= 0) {
      var remId = String(row[idxReminderId] || '').trim();
      if (remId) {
        continue; // do not treat Reminder rows as acks in the legacy submit
      }
    }

    // Enforce note if Needs follow-up
    if (equalsIgnoreCase_(ack, LABELS.NEEDS_FOLLOW_UP) && !note) {
      missingNotes.push(root);
      continue; // do not include this row in payloads
    }

    const role     = roleByRootRep.get(`${root}||${rep}`) || 'Assigned';
    const snapshot = snapByRoot.get(root) || {};
    const ackBy    = Session.getActiveUser().getEmail() || rep;

    payloads.push({ root, rep, role, ackStatus: ack, ackNote: note, ackBy, snapshot });
  }

  // If any rows were missing the required note, stop and tell the user (list the roots)
  if (missingNotes.length > 0) {
    const list = missingNotes.slice(0, 20).join('\n');
    const more = missingNotes.length > 20 ? `\n…and ${missingNotes.length - 20} more.` : '';
    SpreadsheetApp.getUi().alert(
      'Please enter a Note for every row marked "Needs follow-up".\n\nRootApptID(s):\n' + list + more
    );
    return; // abort without writing any logs
  }

  // Append logs
  payloads.forEach(p => appendAckLog_(p));

  // Recompute 00_ summaries and rebuild this rep’s queue (pending only)
  recomputeAckStatusSummary();

  const { expectedByRep, snapByRoot: snaps } = computeExpectedToday_();
  const newRoots = expectedByRep.get(rep) || new Set();
  buildQueueForRep_(rep, newRoots, snaps);

}



// ========================= DETECT / BUILD HELPERS =========================

/** Build a queue tab for the current user if not exists, return sheet */
function ensureQueueSheet_(rep) {
  const name = queueSheetNameForRep_(rep);
  let sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) {
    sh = SpreadsheetApp.getActive().insertSheet(name);
  }
  return sh;
}

function getQueueSheetOrNull_(rep) {
  const name = queueSheetNameForRep_(rep);
  return SpreadsheetApp.getActive().getSheetByName(name);
}

function queueSheetNameForRep_(rep) {
  return QUEUE_PREFIX + shortNameify_(rep);
}

function shortNameify_(name) {
  // "Wendy (PM)" -> "Wendy_PM"
  return String(name || '')
    .replace(/[^\p{L}\p{N}]+/gu, '_')
    .replace(/^_+|_+$/g, '')
    .replace(/_{2,}/g, '_')
    .slice(0, 90);
}

function queueHeaders_() {
  return [
    'RootApptID',
    'Customer Name',
    'Sales Stage',
    'Conversion Status',
    'Custom Order Status',
    'In Production Status',
    'Center Stone Order Status',
    'Next Steps',
    'Updated By',
    'Updated At',
    'Days Since Last Update',
    'Client Status Report URL',
    'Ack Status',   // DV from Ack_Status_List
    'Ack Note'      // required if Needs follow-up
  ];
}

function applyAckStatusValidation_(sheet, colIndex, numRows) {
  const ss = SpreadsheetApp.getActive();
  const rangeNamed = ss.getRangeByName(NR.ACK_STATUS_LIST);
  if (!rangeNamed) return; // silent if not present

  const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeNamed, true)
      .setAllowInvalid(false)
      .build();

  // Apply to [row 2 .. row 1+numRows]
  sheet.getRange(2, colIndex, Math.max(1,numRows), 1).setDataValidation(rule);
}

/** Detect current rep name by email via Dropdown; fallback: chooser over roster */
function detectRepName_() {
  const email = (Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  let rep = '';

  if (email) {
    rep = lookupRepNameByEmail_(email);
    if (rep) return rep;
  }

  // fallback: chooser
  const roster = listAllRosterNames_();
  if (!roster.length) {
    SpreadsheetApp.getUi().alert('No roster found. Please create Dropdown named ranges first.');
    return '';
  }

  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Select Your Name',
    'Enter your display name exactly as in the roster:\n\n' + roster.join(', '),
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return '';
  const input = (resp.getResponseText() || '').trim();
  // normalize against roster using Phase B normalizer
  const norm = buildRosterNormalizer_();
  const canon = norm(input);
  if (!roster.includes(canon)) {
    ui.alert('Name not found in roster. Please try again.');
    return '';
  }
  return canon;
}

/** Return union of both named roster lists (de‑duped) */
function listAllRosterNames_() {
  const ss = SpreadsheetApp.getActive();
  const out = new Set();

  [NR.ROSTER_ASSIGNED, NR.ROSTER_ASSISTED].forEach(nr => {
    const r = ss.getRangeByName(nr);
    if (!r) return;
    r.getValues().flat().forEach(v => {
      const s = String(v || '').trim();
      if (s) out.add(s);
    });
  });

  return [...out].sort((a,b) => a.localeCompare(b));
}

/** Lookup rep name by email on Dropdown sheet (prefers Assigned over Assisted if both match) */
function lookupRepNameByEmail_(emailLC) {
  const sh = SpreadsheetApp.getActive().getSheetByName(DROPDOWN_SHEET);
  if (!sh) return '';

  const values = sh.getDataRange().getValues();
  if (!values.length) return '';

  // header map
  const headers = values[0].map(h => String(h || '').trim());
  const col = {
    aRep:   headers.indexOf(DROPDOWN_HEADERS.ASSIGNED_REP),
    aMail:  headers.indexOf(DROPDOWN_HEADERS.ASSIGNED_EMAIL),
    sRep:   headers.indexOf(DROPDOWN_HEADERS.ASSISTED_REP),
    sMail:  headers.indexOf(DROPDOWN_HEADERS.ASSISTED_EMAIL)
  };

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const aMail = String(row[col.aMail] || '').trim().toLowerCase();
    if (aMail && aMail === emailLC) {
      const aRep = String(row[col.aRep] || '').trim();
      if (aRep) return aRep;
    }

    const sMail = String(row[col.sMail] || '').trim().toLowerCase();
    if (sMail && sMail === emailLC) {
      const sRep = String(row[col.sRep] || '').trim();
      if (sRep) return sRep;
    }
  }
  return '';
}

/** Open (and build if needed) a queue tab for a given rep */
function openQueueForRep_(rep) {
  const {expectedByRep, snapByRoot} = computeExpectedToday_();
  const roots = expectedByRep.get(rep) || new Set();
  buildQueueForRep_(rep, roots, snapByRoot);

  // Switch UI focus to that sheet
  const sh = SpreadsheetApp.getActive().getSheetByName(queueSheetNameForRep_(rep));
  if (sh) SpreadsheetApp.setActiveSheet(sh);
}

// ==== Master (00_) snapshot + lookup (cached per execution) ====

var __PH1_MASTER_CACHE = null;

function _getMasterSnapshot_() {
  if (__PH1_MASTER_CACHE) return __PH1_MASTER_CACHE;

  var sh = SpreadsheetApp.getActive().getSheetByName('00_Master Appointments');
  if (!sh) {
    __PH1_MASTER_CACHE = { headers: [], idx: {}, rows: [], soIdx: new Map(), custIdx: new Map() };
    return __PH1_MASTER_CACHE;
  }

  var values = sh.getDataRange().getDisplayValues();
  var headers = values.shift().map(function(h){ return String(h || '').trim(); });
  function H(label) {
    var key = String(label || '').trim();
    var i = headers.indexOf(key);
    return i >= 0 ? i : -1;
  }

  var IDX = {
    ROOT:   H('RootApptID'),
    SO:     H('SO#'),
    CUSTOMER: H('Customer Name'),
    SALES_STAGE: H('Sales Stage'),
    CONVERSION:  H('Conversion Status'),
    COS:     H('Custom Order Status'),
    INPROD:  H('In Production Status'),
    CSOS:    H('Center Stone Order Status'),
    NEXT:    H('Next Steps'),
    UPD_BY:  H('Updated By'),
    UPD_AT:  H('Updated At'),
    CSR_URL: H('Client Status Report URL'),
    VISIT_DATE: (function(){ var i = headers.indexOf('Visit Date'); if (i<0) i = headers.indexOf('Appt Date'); return i; })(),
    VISIT_TIME: (function(){ var i = headers.indexOf('Visit Time'); if (i<0) i = headers.indexOf('Appt Time'); return i; })()
  };

  var soIdx = new Map();     // key -> row array (latest by Updated At wins)
  var custIdx = new Map();   // lower-cased name -> row array (latest by Updated At wins)

  function toDate(v){
    if (v instanceof Date) return v;
    var n = Number(v); if (!isNaN(n) && n>0) return new Date(n);
    var d = new Date(String(v)); return isNaN(d) ? null : d;
  }
  function soKey(raw) {
    var s = String(raw == null ? '' : raw).trim().replace(/^'+/, '');
    if (!s) return '';
    s = s.replace(/^\s*SO#?/i, '').replace(/\s|\u00A0/g,'');
    var digits = s.replace(/\D/g,'');
    if (!digits) return '';
    return digits.length < 6 ? digits.padStart(6,'0') : digits.slice(-6);
  }

  values.forEach(function(r){
    var so = soKey(r[IDX.SO]);
    var cust = String(r[IDX.CUSTOMER] || '').trim();
    var updAt = toDate(r[IDX.UPD_AT]) || new Date(0);

    // choose latest by Updated At
    function maybePut(map, key) {
      if (!key) return;
      var rec = map.get(key);
      if (!rec || (toDate(rec[IDX.UPD_AT]) || new Date(0)) < updAt) {
        map.set(key, r);
      }
    }
    maybePut(soIdx, so);
    if (cust) maybePut(custIdx, cust.toLowerCase());
  });

  __PH1_MASTER_CACHE = { headers: headers, idx: IDX, rows: values, soIdx: soIdx, custIdx: custIdx };
  return __PH1_MASTER_CACHE;
}

function _lookupMasterBySoOrCustomer_(soPretty, customer, master) {
  master = master || _getMasterSnapshot_();

  function soKeyFromPretty(raw) {
    var s = String(raw || '').trim();
    // pretty form "##.####" → digits
    var digits = s.replace(/\D/g,'');
    if (!digits) return '';
    return digits.length < 6 ? digits.padStart(6,'0') : digits.slice(-6);
  }

  var rec = null;
  var key = soKeyFromPretty(soPretty);
  if (key && master.soIdx.has(key)) rec = master.soIdx.get(key);
  if (!rec && customer) {
    var k = String(customer).trim().toLowerCase();
    rec = master.custIdx.get(k) || null;
  }

  if (!rec) {
    return { root:'', customer: customer||'', salesStage:'', conversionStatus:'', cos:'', inProd:'', csos:'', updatedBy:'', updatedAt:'', csrUrl:'' };
  }

  var I = master.idx;
  return {
    root: rec[I.ROOT] || '',
    customer: rec[I.CUSTOMER] || customer || '',
    salesStage: rec[I.SALES_STAGE] || '',
    conversionStatus: rec[I.CONVERSION] || '',
    cos: rec[I.COS] || '',
    inProd: rec[I.INPROD] || '',
    csos: rec[I.CSOS] || '',
    updatedBy: rec[I.UPD_BY] || '',
    updatedAt: (function(v){ var d=(v instanceof Date)?v:new Date(String(v)); return isNaN(d)?'':d; })(rec[I.UPD_AT]),
    csrUrl: rec[I.CSR_URL] || ''
  };
}

// ==== Descriptive "Next Steps" text for Reminders ====

function _buildReminderNextStepsText_(type, firstDueDate, cosLabel, baseText) {
  var now = new Date();
  var MS = 24*60*60*1000;

  function daysBetween(a,b){ return Math.floor((a.getTime()-b.getTime())/MS); }
  function prettyDays(n){ return (n===1) ? '1 day ago' : (n+' days ago'); }

  // Default if nothing else
  var suffix = baseText ? (' — ' + baseText) : '';

  // COS: we back-calc event date as (firstDueDate − 2 days), per reminder policy. 2-day buffer per spec. 
  // (This preserves the existing engine’s “start after 2 days” behavior.)  :contentReference[oaicite:3]{index=3}
  if (type === 'COS') {
    var cos = String(cosLabel || '3D pending').trim();
    var eventDate = null;
    if (firstDueDate instanceof Date) {
      eventDate = new Date(firstDueDate.getTime() - 2*MS);
    } else if (firstDueDate) {
      var d = new Date(String(firstDueDate));
      if (!isNaN(d)) eventDate = new Date(d.getTime() - 2*MS);
    }
    var phr = cos;
    if (eventDate) {
      var n = Math.max(0, daysBetween(now, eventDate));
      phr += ' — requested ' + prettyDays(n);
    }
    return '[Reminder] ' + phr + suffix;
  }

  // FOLLOWUP: nudge with simple language; engine already decides due timing.  :contentReference[oaicite:4]{index=4}
  if (type === 'FOLLOWUP') {
    return '[Reminder] Follow-Up Required — log outreach and update Sales Stage' + suffix;
  }

  // DV flavors (if present in your queue): keep generic, we don’t recompute appointment deltas here.
  if (type === 'DV_URGENT' || type === 'DV_PROPOSE') {
    return (type === 'DV_URGENT')
      ? '[Reminder] Diamond Viewing — Urgent: stones not lined up' + suffix
      : '[Reminder] Diamond Viewing — Propose: send options' + suffix;
  }

  // Legacy 3D subtypes (if your queue still uses them)
  if (type === 'START3D' || type === 'ASSIGNSO' || type === 'REV3D') {
    return '[Reminder] 3D Task — ' + type + suffix;
  }

  // Fallback
  return '[Reminder] ' + (type || 'Task') + suffix;
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



