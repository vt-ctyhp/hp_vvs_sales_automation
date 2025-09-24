// @bundle: Ack pipes + dashboard + schedule + snapshot
/** Policy engine for acknowledgement scope groups.
 * Reads 12_Ack_Policies and classifies each root into first-match group.
 * Keeps today's knobs: MustAck, QueueInclude, SnapshotInclude, AckCadence, Coverage.
 */

const POLICY_SHEET = '12_Ack_Policies';

// Built-in *lightweight* aliases for robust matching (case/trim-insensitive).
// You can always add additional values in the policy's "Match Values" cell.
const ACK_ALIASES = {
  'Sales Stage': {
    'appointment': ['appt', 'appointment scheduled', 'booked appointment'], // safety
    'follow-up required': ['follow up required', 'follow-up', 'follow up']
  },
  'Conversion Status': {
    'viewing scheduled': ['viewing scheduled']
  },
  'Custom Order Status': {
    'in production': ['in-production', 'in prod', 'production']
  }
};

// --- tiny text utils ---
function _norm_(s) {
  return String(s || '')
    .replace(/\u2011|\u2013|\u2014/g, '-')   // normalize dashes
    .replace(/\s+/g, ' ')                    // collapse spaces
    .trim()
    .toLowerCase();
}
function _splitList_(s) {
  if (s == null) return [];
  return String(s).split(',').map(x => _norm_(x)).filter(Boolean);
}
function _expandAliasesFor_(colName, tokens) {
  const m = ACK_ALIASES[colName] || {};
  const out = new Set();
  tokens.forEach(t => {
    out.add(t);
    const aliasArr = m[t] || [];
    aliasArr.forEach(a => out.add(_norm_(a)));
  });
  return [...out];
}
function _matchesWithAliases_(actual, expectedList, colName) {
  const a = _norm_(actual);
  if (!a) return false;
  const expanded = _expandAliasesFor_(colName, expectedList);
  return expanded.includes(a);
}

// --- read policies from 12_Ack_Policies ---
function readAckPolicies_() {
  const sh = getSheetOrThrow_(POLICY_SHEET);
  const rows = getObjects_(sh);
  const out = [];
  rows.forEach(r => {
    if (_norm_(r['Enabled']) !== 'y') return;
    const priority = Number(r['Priority'] || 9999);
    const group = String(r['Group Name'] || '').trim();
    const col = String(r['Match Column'] || '').trim();
    const vals = _splitList_((r['Match Values (comma-sep)'] || ''));
    const mustAck = String(r['MustAck'] || 'ALL_ON_DUTY').trim();
    const qIncl = _norm_(r['QueueInclude']) === 'y';
    const sIncl = _norm_(r['SnapshotInclude']) === 'y';
    const cadence = String(r['AckCadence'] || 'DAILY').trim();
    const assistedCover = _norm_(r['Coverage Assisted Pairing']) === 'y';

    if (!group || !col || !vals.length) return; // skip malformed rows safely

    out.push({
      priority,
      group,
      matchColumn: col,
      matchValues: vals,
      mustAck,
      queueInclude: qIncl,
      snapshotInclude: sIncl,
      ackCadence: cadence,
      assistedCoverage: assistedCover
    });
  });
  out.sort((a,b) => (a.priority||9999) - (b.priority||9999));
  return out;
}

// --- pick the first matching group for a 07_Root_Index row ---
function classifyScopeGroupForRoot_(rootObj, policies) {
  for (const p of policies) {
    const col = p.matchColumn;
    const actual = rootObj[col];
    if (_matchesWithAliases_(actual, p.matchValues, col)) return p.group;
  }
  return null; // out of scope for acknowledgements
}

// --- classify all roots once; return maps for downstream callers ---
function classifyAllRoots_(idx07Objs, policies) {
  const inScope = new Set();
  const scopeGroupByRoot = new Map();
  const snapByRoot = new Map(); // same as before: canonical "today" row per root
  const nameByRoot = new Map();

  idx07Objs.forEach(r => {
    const root = String(r['RootApptID'] || '').trim();
    if (!root) return;
    const g = classifyScopeGroupForRoot_(r, policies);
    if (!g) return;            // not in any policy group → out of scope
    inScope.add(root);
    scopeGroupByRoot.set(root, g);
    snapByRoot.set(root, r);
    nameByRoot.set(root, String(r['Customer Name'] || '').trim());
  });
  return { inScope, scopeGroupByRoot, snapByRoot, nameByRoot };
}

/** Helper for callers that previously took a set of roots:
 * We keep schedule/coverage logic as-is by reusing computeExpectedSetsWithSchedule_(inScope, map08).
 */
function computeExpectedSetsWithPolicies_(idx07Objs, map08Objs) {
  // 1) Read policies + classify roots
  const policies = readAckPolicies_();
  const { inScope, scopeGroupByRoot, snapByRoot, nameByRoot } = classifyAllRoots_(idx07Objs, policies);

  // Build a quick lookup: group → MustAck (UPPERCASE; default ALL_ON_DUTY)
  const mustAckByGroup = new Map();
  policies.forEach(p => {
    const v = String(p.mustAck || 'ALL_ON_DUTY').trim().replace(/[ \-]+/g,'_').toUpperCase();
    // First policy hit for a group wins; later rows can be added but won’t override here.
    if (!mustAckByGroup.has(p.group)) mustAckByGroup.set(p.group, v);
  });

  // 2) Schedule + assisted coverage (role‑aware)
  const {
    expectedByRootDuty,        // Map<root, Set<rep>> (assigned+assisted after on‑duty + coverage)
    expectedByRepDuty,         // Map<rep, Set<root>>
    roleByRootRepDuty,         // Map<`${root}||${rep}`, 'Assigned'|'Assisted'>
    assignedGaps,
    assistedGaps
  } = computeExpectedSetsWithSchedule_(inScope, map08Objs);

  // 3) Apply MustAck filter per policy group
  //    Supported values:
  //      - ALL_ON_DUTY (default) → keep both Assigned & Assisted (current behavior)
  //      - ASSIGNED_REPS_ONLY    → keep only reps whose duty role is 'Assigned'
  //      - ASSISTED_REPS_ONLY    → keep only reps whose duty role is 'Assisted'
  const filteredByRoot = new Map();
  expectedByRootDuty.forEach((repSet, root) => {
    const group = scopeGroupByRoot.get(root) || '';
    const rule  = (mustAckByGroup.get(group) || 'ALL_ON_DUTY');

    // Fast exit if default behavior
    if (rule === 'ALL_ON_DUTY') {
      filteredByRoot.set(root, new Set(repSet));
      return;
    }

    const keep = new Set();
    repSet.forEach(rep => {
      const role = roleByRootRepDuty.get(`${root}||${rep}`) || 'Assigned';
      if (rule === 'ASSIGNED_REPS_ONLY' && role === 'Assigned') keep.add(rep);
      if (rule === 'ASSISTED_REPS_ONLY' && role === 'Assisted') keep.add(rep);
    });

    // If nothing matches the rule for this root, we intentionally leave it empty (no expected acks today).
    if (keep.size > 0) filteredByRoot.set(root, keep);
    else filteredByRoot.set(root, new Set()); // preserve presence with empty set
  });

  // 4) Invert (rep → roots) after filtering
  const filteredByRep = new Map();
  filteredByRoot.forEach((repSet, root) => {
    repSet.forEach(rep => {
      if (!filteredByRep.has(rep)) filteredByRep.set(rep, new Set());
      filteredByRep.get(rep).add(root);
    });
  });

  return {
    policies,
    inScope,
    scopeGroupByRoot,
    snapByRoot,
    nameByRoot,
    expectedByRootDuty: filteredByRoot,   // ← now MustAck‑aware
    expectedByRepDuty:  filteredByRep,    // ← now MustAck‑aware
    roleByRootRepDuty,                    // handy for UI/debug
    assignedGaps,
    assistedGaps
  };
}

// --- convenience for queue rendering: group ordering from policy priority ---
function getPolicyGroupOrder_() {
  return readAckPolicies_().map(p => p.group);
}


/** Render queue rows grouped by policy priority (one tab per rep).
 * - Clears previous content AND banding safely.
 * - Writes header + each group as a separate block with a one-line title.
 * - Applies alternating row banding per block (no overlap).
 */
function renderQueueGroupedRows_(sheet, headers, rows) {
  // Hard reset: contents + formats + any existing banding
  sheet.clearContents();
  sheet.clearFormats();
  try {
    const bandings = sheet.getBandings();
    if (bandings && bandings.length) bandings.forEach(b => { try { b.remove(); } catch(_) {} });
  } catch(_) {}


  // 1) Remove any existing alternating banding on this sheet to avoid overlap errors
  try {
    const bandings = sheet.getBandings();
    if (bandings && bandings.length) bandings.forEach(b => { try { b.remove(); } catch(_) {} });
  } catch (_) {
    // Safe: some older editors don’t support getBandings; just continue.
  }

  // 2) Always write the header
  if (!headers || !headers.length) return;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // If no data, courtesy empty state and bail
  if (!rows || !rows.length) {
    sheet.getRange(2, 1).setValue('— No pending items today —').setFontStyle('italic');
    sheet.autoResizeColumns(1, headers.length);
    return;
  }

  // 3) Build 07 index once to classify scope group
  const s07 = getSheetOrThrow_('07_Root_Index');
  const idx07 = getObjects_(s07);
  const byRoot = new Map();
  idx07.forEach(r => {
    const root = String(r['RootApptID'] || '').trim();
    if (root) byRoot.set(root, r);
  });

  // Load policies + priority order
  const policies = readAckPolicies_();
  const groupOrder = policies.map(p => p.group);

  // Column positions used for sorting/grouping
  const colRoot = headers.indexOf('RootApptID');
  const colDays = headers.indexOf('Days Since Last Update');
  const colCust = headers.indexOf('Customer Name');

  // 4) Bucket rows by scope group (first-match wins)
  const buckets = new Map(); // group -> array of row arrays
  groupOrder.forEach(g => buckets.set(g, [])); // initialize in policy order

  rows.forEach(row => {
    const root = colRoot >= 0 ? String(row[colRoot] || '').trim() : '';
    const r07  = root ? (byRoot.get(root) || {}) : {};
    const g    = classifyScopeGroupForRoot_(r07, policies);
    if (!g) return;             // out of scope -> skip
    if (!buckets.has(g)) buckets.set(g, []); // in case group was added later
    buckets.get(g).push(row);
  });

  // 5) Write group blocks in policy order; apply banding to each block only once
  let curRow = 2; // first data row after header
  groupOrder.forEach(g => {
    const arr = buckets.get(g) || [];
    if (!arr.length) return;

    // Sort within group: Days DESC, Customer ASC (when available)
    arr.sort((A,B) => {
      const dA = colDays >= 0 ? Number(A[colDays] || 0) : 0;
      const dB = colDays >= 0 ? Number(B[colDays] || 0) : 0;
      if (dB !== dA) return dB - dA;
      const cA = colCust >= 0 ? String(A[colCust] || '') : '';
      const cB = colCust >= 0 ? String(B[colCust] || '') : '';
      return cA.localeCompare(cB);
    });

    // Group title row (non-banded) — style across the full width
    const title = `— ${g} — (${arr.length})`;
    sheet.getRange(curRow, 1, 1, headers.length)
        .setValue(title)
        .mergeAcross()                // single merged cell looks cleaner; remove if you prefer not merged
        .setFontWeight('bold')
        .setFontStyle('italic')
        .setHorizontalAlignment('left')
        .setBackground('#bdbdbd');    // slightly darker grey for the label row
    curRow += 1;
    
    // Add a subtle bottom border to separate label from data
    sheet.getRange(curRow-1, 1, 1, headers.length)
        .setBorder(false, false, true, false, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

    // Data block
    const start = curRow;
    sheet.getRange(start, 1, arr.length, headers.length).setValues(arr);

    // Apply banding to the data block only (not the title row, not the header)
    try {
      // Apply alternating colors to the data block only (no header/footer styling)
      sheet.getRange(start, 1, arr.length, headers.length)
          .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, /*showHeader=*/false, /*showFooter=*/false);
    } catch (e) {
      // If a future overlap is detected for any reason, skip banding this block and continue
      // (Overlaps shouldn't occur because we cleared sheet-level banding up front.)
    }

    curRow = start + arr.length + 1; // +1 blank spacer between groups
  });

  // 6) Tidy
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Daily orchestrator for Acknowledgements
 * Runs in the correct order, with basic logging and guardrails.
 * Assumes all functions below already exist in your project:
 *   - buildRootIndex, buildRepsMap            (Phase B)
 *   - takeMorningSnapshot                     (Phase F, policy-powered)
 *   - buildTodaysQueuesAll                    (Phase C)
 *   - buildAckDashboard                       (Phase D)
 */

// timezone (matches your workbook)
const ACK_SCHED_TZ = (typeof TIMEZONE !== 'undefined' && TIMEZONE) ? TIMEZONE : 'America/Los_Angeles';

/** Run the full morning flow (one call) */
function ack_runMorningFlow() {
  const tz = ACK_SCHED_TZ;
  const start = new Date();
  Logger.log(`[ACK] Morning flow start @ ${Utilities.formatDate(start, tz, 'yyyy-MM-dd HH:mm:ss')}`);

  try {
    // 1) Refresh base tables (07, 08)
    Logger.log('[ACK] buildRootIndex()');
    buildRootIndex();

    Logger.log('[ACK] buildRepsMap()');
    buildRepsMap();

    // 2) Freeze denominators for today (policy + schedule)
    Logger.log('[ACK] takeMorningSnapshot()');
    takeMorningSnapshot();

    Logger.log('[ACK] buildTodaysQueuesAll()');
    buildTodaysQueuesAll();

    // 3b) Inject Reminders into every Q_<rep> tab (replace-mode; safe if none due)
    try {
      Logger.log('[ACK] injectRemindersForAllReps()');
      ack_injectRemindersForAllReps();
    } catch (e) {
      Logger.log('[ACK] injectRemindersForAllReps error: ' + (e && (e.stack || e.message) || e));
    }

    // 4) Refresh dashboard (live + snapshot)
    Logger.log('[ACK] buildAckDashboard()');
    buildAckDashboard();

    const end = new Date();
    Logger.log(`[ACK] Morning flow done in ~${Math.round((end - start)/1000)}s`);

  } catch (e) {
    Logger.log('[ACK] Morning flow ERROR: ' + (e && (e.stack || e.message) || e));
    // Optional: post a quick chat/email alert here if you use ops notifications
  }
}

/**
 * Inject Reminders into all Q_<rep> tabs (replace-mode).
 * Uses your existing:
 *   - _allQueueReps_()            // reps from 08_Reps_Map with Include?=Y
 *   - _injectRemindersAfterBuild_(rep) // does ensure + remove-old + insert-new + style
 *   - getRemindersInAckFlag_()    // honors your flag (off → no-op)
 */
function ack_injectRemindersForAllReps() {
  // Respect flag (same switch you already use in refreshMyQueueHybrid)
  if (typeof getRemindersInAckFlag_ === 'function' && !getRemindersInAckFlag_()) {
    Logger.log('[ACK] REMINDERS_IN_ACK is OFF — skipping global injection.');
    return;
  }

  // Pull the reps you already track in 08_Reps_Map (Include?=Y)
  var reps = (typeof _allQueueReps_ === 'function') ? _allQueueReps_() : [];
  if (!reps || !reps.length) {
    Logger.log('[ACK] No reps to inject (08_Reps_Map had none with Include?=Y).');
    return;
  }

  // For each rep, inject/replace the Reminders block at the top of Q_<rep>
  var ok = 0, fail = 0;
  reps.forEach(function(rep) {
    try {
      _injectRemindersAfterBuild_(rep);  // idempotent; ensures tab exists
      ok++;
    } catch (e) {
      fail++;
      Logger.log('[ACK] Reminder injection failed for ' + rep + ': ' + (e && (e.stack || e.message) || e));
    }
  });
  Logger.log('[ACK] Reminders injected for ' + ok + ' reps' + (fail ? (', failures=' + fail) : ''));
}


/** (Optional) Late-day dashboard rebuild */
function ack_lateDayDashboardRefresh() {
  try {
    Logger.log('[ACK] Late-day: buildAckDashboard()');
    buildAckDashboard();
  } catch (e) {
    Logger.log('[ACK] Late-day refresh ERROR: ' + (e && (e.stack || e.message) || e));
  }
}

/** One-time installer: create (or clean and re-create) the time-driven triggers */
function ack_installDailyTriggers() {
  // Clear existing triggers for these handlers (safe idempotent)
  const keep = new Set(['ack_runMorningFlow','ack_middayQueuesRefresh','ack_lateDayDashboardRefresh']);
  ScriptApp.getProjectTriggers().forEach(t => {
    if (keep.has(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });

  // Morning flow @ 8:25 AM PT (choose your time window)
  ScriptApp.newTrigger('ack_runMorningFlow')
           .timeBased()
           .atHour(8).nearMinute(25)        // runs between 8:25–8:30
           .everyDays(1)
           .create();

  // Optional midday queue refresh @ 1:00 PM PT
  ScriptApp.newTrigger('ack_middayQueuesRefresh')
           .timeBased()
           .atHour(13).nearMinute(0)
           .everyDays(1)
           .create();

  // Optional late-day dashboard rebuild @ 4:30 PM PT
  ScriptApp.newTrigger('ack_lateDayDashboardRefresh')
           .timeBased()
           .atHour(16).nearMinute(30)
           .everyDays(1)
           .create();

  Logger.log('[ACK] Triggers installed');
}

/** One-time: remove all ACK scheduler triggers (cleanup tool) */
function ack_removeAllAckSchedulerTriggers() {
  const handlers = new Set(['ack_runMorningFlow','ack_middayQueuesRefresh','ack_lateDayDashboardRefresh']);
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlers.has(t.getHandlerFunction())) { ScriptApp.deleteTrigger(t); removed++; }
  });
  Logger.log(`[ACK] Removed ${removed} ACK triggers`);
}

function ack_installHourlyDashboardTrigger() {
  // Idempotent cleanup
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === "refreshDashboardHourly") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // If you already have refreshDashboardHourly, we’ll use it.
  // If not, we provide a tiny fallback implementation below.
  ScriptApp.newTrigger("refreshDashboardHourly")
    .timeBased()
    .everyHours(1)
    .create();
}

// Optional: fallback if you don't already define this elsewhere
function refreshDashboardHourly() {
  try { buildAckDashboard(); } catch (e) { Logger.log("Hourly dashboard refresh error: " + e); }
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



