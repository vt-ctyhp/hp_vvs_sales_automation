/***** Phase D — Manager Dashboard (09_Ack_Dashboard) *****
 * Builds a real-time dashboard with:
 *  1) Compliance Today (by Rep)
 *  2) Missing Acks (today)
 *  3) Needs Follow-up (last N days)
 *  4) Stale Updates (Days Since Last Update >= threshold)
 *
 * Requires Phases B & C:
 *  - Sheets: 06_Acknowledgement_Log, 07_Root_Index, 08_Reps_Map, 09_Ack_Dashboard
 *  - Helpers: getObjects_, getSheetOrThrow_, clearAndWrite_, getHeaders_,
 *             toDateSafe_, equalsIgnoreCase_, buildRosterNormalizer_
 *  - Constants: TIMEZONE, IN_PRODUCTION_LITERAL, EXCLUDE_SALES_STAGES / _LC
 ****************************************************************/

function MASTER_SS_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  // fallback for container-bound or editor runs
  return SpreadsheetApp.getActive();
}

// === Dashboard colors ===
const DASH_HEADER_BG = '#e31c79'; // header background (VVS pink)
const DASH_HEADER_FG = '#ffffff'; // header text
const DASH_BAND_1    = '#ffffff'; // body stripe 1
const DASH_BAND_2    = '#fff0fb'; // body stripe 2

// Timezone
const TZ_DASH = typeof TIMEZONE !== 'undefined' ? TIMEZONE : 'America/Los_Angeles';

// Look-back window for Needs Follow-up panel (days)
const NEEDS_FOLLOWUP_LOOKBACK_DAYS = 7;

// Stale threshold (days since last real Client Status update)
const STALE_UPDATE_DAYS_THRESHOLD = 3;

// Fallbacks (only if not defined in other files)
const IN_PRODUCTION_LITERAL_FALLBACK = 'In Production';
const IN_PROD = (typeof IN_PRODUCTION_LITERAL !== 'undefined' ? IN_PRODUCTION_LITERAL : IN_PRODUCTION_LITERAL_FALLBACK);

/** Entry point: rebuild the entire 09_Ack_Dashboard (with formatting & optional filter views) */
function buildAckDashboard() {
  const ss = MASTER_SS_();
  const s09 = getSheetOrThrow_('09_Ack_Dashboard');

  // Base sheets
  const s07 = getSheetOrThrow_('07_Root_Index');
  const s08 = getSheetOrThrow_('08_Reps_Map');
  const s06 = getSheetOrThrow_('06_Acknowledgement_Log');

  // Snapshot sheets (your naming)
  const SNAP_TODAY_SHEET = '13_Morning_Snapshot';
  const SNAP_LOG_SHEET   = '14_Snapshot_Log';

  const tz = (typeof TIMEZONE !== 'undefined' && TIMEZONE) ? TIMEZONE : ss.getSpreadsheetTimeZone();
  const todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const now = new Date();

  // ---------- Pull canonical "today" snapshot (07) ----------
  const idx07 = getObjects_(s07);
  const map08 = getObjects_(s08);

  // Use policy engine to classify scope and build expected sets
  const {
    policies,
    inScope,
    scopeGroupByRoot,
    snapByRoot,
    nameByRoot,
    expectedByRootDuty,
    expectedByRepDuty,
    assignedGaps,
    assistedGaps
  } = computeExpectedSetsWithPolicies_(idx07, map08);




  // ---------- Read only a small window from big logs ----------
  // We need: all of TODAY (for compliance/missing) + last 7 days (for needs-FU & trailing completion)
  const LOG06_LOOKBACK_DAYS = Math.max(NEEDS_FOLLOWUP_LOOKBACK_DAYS, 7) + 1; // today + 7
  const log06 = getObjectsByDateWindow_(s06, 'Log Date', LOG06_LOOKBACK_DAYS);

  // Latest ACK today by (root, rep)
  const latestTodayByRootRep = new Map(); // key: root||rep -> {status, ts}
  log06.forEach(r => {
    const root = String(r['RootApptID'] || '').trim();
    const rep  = String(r['Rep'] || '').trim();
    const status = String(r['Ack Status'] || '').trim();
    const logDateVal = r['Log Date'];
    const tsVal = toDateSafe_(r['Timestamp']) || new Date();
    const dateKey = (logDateVal instanceof Date)
      ? Utilities.formatDate(logDateVal, tz, 'yyyy-MM-dd')
      : String(logDateVal || '').slice(0,10);

    if (dateKey !== todayKey || !inScope.has(root) || !rep || !status) return;
    const key = `${root}||${rep}`;
    const prev = latestTodayByRootRep.get(key);
    if (!prev || tsVal.getTime() > prev.ts.getTime()) latestTodayByRootRep.set(key, { status, ts: tsVal });
  });

  // ---------- Overload Today (Snapshot) by team ----------
  const snapTodaySheet = ss.getSheetByName(SNAP_TODAY_SHEET);
  const snapTodayObjs = snapTodaySheet ? getObjectsByDateWindow_(snapTodaySheet, 'Snapshot Date', 1) : [];
  const assignedCountTodayByRep = new Map();
  const assistedCountTodayByRep = new Map();

  snapTodayObjs.forEach(r => {
    const dVal = r['Snapshot Date'];
    const dateKeySnap = (dVal instanceof Date)
      ? Utilities.formatDate(dVal, tz, 'yyyy-MM-dd')
      : String(dVal || '').slice(0, 10);
    if (dateKeySnap !== todayKey) return;
    const rep = String(r['Rep'] || '').trim();
    const role = String(r['Role'] || '').trim();
    if (!rep || !role) return;
    if (equalsIgnoreCase_(role, 'Assigned')) {
      assignedCountTodayByRep.set(rep, (assignedCountTodayByRep.get(rep) || 0) + 1);
    } else if (equalsIgnoreCase_(role, 'Assisted')) {
      assistedCountTodayByRep.set(rep, (assistedCountTodayByRep.get(rep) || 0) + 1);
    }
  });

  const assignedCounts = [...assignedCountTodayByRep.values()].filter(n => n > 0);
  const assistedCounts = [...assistedCountTodayByRep.values()].filter(n => n > 0);
  const medianAssigned = (assignedCounts.length ? median_(assignedCounts) : 0);
  const medianAssisted = (assistedCounts.length ? median_(assistedCounts) : 0);

  const assignedOverloadRows = [...assignedCountTodayByRep.entries()]
    .filter(([,exp]) => exp > 0 && medianAssigned > 0)
    .map(([rep, exp]) => [rep, exp, medianAssigned, exp / medianAssigned])
    .sort((a,b) => (b[3] - a[3]) || (b[1] - a[1]) || String(a[0]).localeCompare(String(b[0])))
    .slice(0, 5);

  const assistedOverloadRows = [...assistedCountTodayByRep.entries()]
    .filter(([,exp]) => exp > 0 && medianAssisted > 0)
    .map(([rep, exp]) => [rep, exp, medianAssisted, exp / medianAssisted])
    .sort((a,b) => (b[3] - a[3]) || (b[1] - a[1]) || String(a[0]).localeCompare(String(b[0])))
    .slice(0, 5);

  // ---------- Compliance & Missing (real-time, on-duty only) ----------
  const complianceRows = [];
  const missingRows = [];
  [...expectedByRepDuty.keys()].sort().forEach(rep => {
    const roots = expectedByRepDuty.get(rep) || new Set();
    const expected = roots.size;
    let fully = 0, needs = 0, ackedAny = 0;

    roots.forEach(root => {
      const key = `${root}||${rep}`;
      const latest = latestTodayByRootRep.get(key);
      if (latest) {
        ackedAny++;
        const st = String(latest.status || '').trim();
        if (equalsIgnoreCase_(st, LABELS.FULLY_UPDATED)) fully++;
        else if (equalsIgnoreCase_(st, LABELS.NEEDS_FOLLOW_UP)) needs++;
      } else {
        const snap = snapByRoot.get(root) || {};
        missingRows.push([
          rep,
          root,
          String(snap['Customer Name'] || '').trim(),
          snap['Updated By'] || '',
          snap['Updated At'] || '',
          snap['Days Since Last Update'] || '',
          snap['Client Status Report URL'] || ''
        ]);
      }
    });

    const missing = expected - ackedAny;
    const pct = expected ? (fully / expected) : '';
    complianceRows.push([rep, expected, fully, needs, missing, pct]);
  });

  complianceRows.sort((a,b) => {
    const miss = (b[4]||0) - (a[4]||0);
    if (miss !== 0) return miss;
    const pa = (a[5] === '' ? -1 : a[5]);
    const pb = (b[5] === '' ? -1 : b[5]);
    if (pa !== pb) return pa - pb;
    return String(a[0]).localeCompare(String(b[0]));
  });

  // ---------- Snapshot Compliance panel (frozen denominator) ----------
  const { snapExpectedByRep, captureTime } = readTodaySnapshot_();
  const snapComplianceRows = [];
  if (snapExpectedByRep && snapExpectedByRep.size) {
    [...snapExpectedByRep.keys()].sort().forEach(rep => {
      const roots = snapExpectedByRep.get(rep) || new Set();
      const expected = roots.size;
      let fully = 0, needs = 0, ackedAny = 0;

      roots.forEach(root => {
        const key = `${root}||${rep}`;
        const latest = latestTodayByRootRep.get(key);
        if (latest) {
          ackedAny++;
          const st = String(latest.status || '').trim();
          if (equalsIgnoreCase_(st, LABELS.FULLY_UPDATED)) fully++;
          else if (equalsIgnoreCase_(st, LABELS.NEEDS_FOLLOW_UP)) needs++;
        }
      });

      const missing = expected - ackedAny;
      const pct = expected ? (fully / expected) : '';
      snapComplianceRows.push([rep, expected, fully, needs, missing, pct]);
    });

    snapComplianceRows.sort((a,b) => {
      const miss = (b[4]||0) - (a[4]||0);
      if (miss !== 0) return miss;
      const pa = (a[5] === '' ? -1 : a[5]);
      const pb = (b[5] === '' ? -1 : b[5]);
      if (pa !== pb) return pa - pb;
      return String(a[0]).localeCompare(String(b[0]));
    });
  }

  // ---------- Needs-FU (7d) and Stale ----------
  const cutoff = new Date(now.getTime() - NEEDS_FOLLOWUP_LOOKBACK_DAYS * 86400000);
  const needsFUByPairLatest = new Map();
  const latestFUtsByPair = new Map();

  log06.forEach(r => {
    const root = String(r['RootApptID'] || '').trim();
    const rep  = String(r['Rep'] || '').trim();
    const status = String(r['Ack Status'] || '').trim();
    const ts   = toDateSafe_(r['Timestamp']) || new Date();

    if (equalsIgnoreCase_(status, LABELS.FULLY_UPDATED)) {
      const k = `${root}||${rep}`;
      const prev = latestFUtsByPair.get(k);
      if (!prev || ts.getTime() > prev.getTime()) latestFUtsByPair.set(k, ts);
    }
    if (ts >= cutoff && equalsIgnoreCase_(status, LABELS.NEEDS_FOLLOW_UP) && rep) {
      const k = `${root}||${rep}`;
      const cur = needsFUByPairLatest.get(k);
      if (!cur || ts.getTime() > (toDateSafe_(cur['Timestamp']) || 0)) {
        needsFUByPairLatest.set(k, r);
      }
    }
  });

  const needsRows = [];
  needsFUByPairLatest.forEach(rec => {
    const root = String(rec['RootApptID'] || '').trim();
    const rep  = String(rec['Rep'] || '').trim();
    const note = String(rec['Ack Note'] || '').trim();
    const ts   = toDateSafe_(rec['Timestamp']) || '';
    const cust = String(rec['Customer (at log)'] || '').trim() || nameByRoot.get(root) || '';
    const cos  = String(rec['Custom Order Status (at log)'] || '').trim();
    const updAt= toDateSafe_(rec['Last Updated At (at log)']) || '';
    const url  = String(rec['Client Status Report URL'] || '').trim();

    const laterFU = latestFUtsByPair.get(`${root}||${rep}`);
    const resolved = laterFU && ts && laterFU.getTime() > ts.getTime();

    needsRows.push([rep, root, cust, ts, resolved ? 'Yes' : 'No', note, cos, updAt, url]);
  });

  needsRows.sort((a,b) => {
    const r = String(a[4]).localeCompare(String(b[4])); // 'No' < 'Yes'
    if (r !== 0) return r;
    const ta = toDateSafe_(a[3]) || new Date(0);
    const tb = toDateSafe_(b[3]) || new Date(0);
    if (tb.getTime() !== ta.getTime()) return tb.getTime() - ta.getTime();
    return String(a[0]).localeCompare(String(b[0]));
  });

  // Stale (in-scope only)
  const staleRows = [];
  snapByRoot.forEach((r, root) => {
    const days = Number(r['Days Since Last Update'] || 0);
    if (days >= STALE_UPDATE_DAYS_THRESHOLD) {
      staleRows.push([
        root,
        String(r['Customer Name'] || '').trim(),
        days,
        r['Updated By'] || '',
        r['Updated At'] || '',
        r['Client Status Report URL'] || ''
      ]);
    }
  });
  staleRows.sort((a,b) => {
    const d = (Number(b[2])||0) - (Number(a[2])||0);
    if (d !== 0) return d;
    return String(a[1]).localeCompare(String(b[1]));
  });

  // ---------- Trailing 7‑Day Completion (by team) ----------
  const snapLogSheet = ss.getSheetByName(SNAP_LOG_SHEET);
  const snapLogObjs = snapLogSheet ? getObjectsByDateWindow_(snapLogSheet, 'Snapshot Date', 7) : [];

  // Build 7-day date set
  const last7 = [];
  for (let i = 6; i >= 0; i--) {
    const d = new Date(now.getTime() - i*86400000);
    last7.push(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
  }
  const last7Set = new Set(last7);

  // Denominators by rep & role; also role per rep (rep is single-team)
  const expected7ByRep = new Map(); // rep -> count
  const repRole = new Map();        // rep -> 'Assigned'|'Assisted'
  const snapshotPairSet = new Set(); // `${date}||${root}||${rep}`

  snapLogObjs.forEach(r => {
    const d = String(r['Snapshot Date'] || '').trim();
    if (!last7Set.has(d)) return;
    const rep = String(r['Rep'] || '').trim();
    const role = String(r['Role'] || '').trim();
    const root = String(r['RootApptID'] || '').trim();
    if (!rep || !role || !root) return;

    repRole.set(rep, role); // single-team invariant
    expected7ByRep.set(rep, (expected7ByRep.get(rep) || 0) + 1);
    snapshotPairSet.add(`${d}||${root}||${rep}`);
  });

  // Numerators: Fully Updated acks that correspond to snapshot pairs that day
  const done7ByRep = new Map(); // rep -> count
  const countedAck = new Set(); // avoid double-count for same (day, root, rep)
  log06.forEach(r => {
    const rep = String(r['Rep'] || '').trim();
    const root = String(r['RootApptID'] || '').trim();
    const status = String(r['Ack Status'] || '').trim();
    const logDateVal = r['Log Date'];
    if (!rep || !root || !equalsIgnoreCase_(status, LABELS.FULLY_UPDATED)) return;

    const dateKey = (logDateVal instanceof Date)
      ? Utilities.formatDate(logDateVal, tz, 'yyyy-MM-dd')
      : String(logDateVal || '').slice(0,10);

    const key = `${dateKey}||${root}||${rep}`;
    if (!last7Set.has(dateKey)) return;
    if (!snapshotPairSet.has(key)) return;         // only count if rep owed it that day
    if (countedAck.has(key)) return;               // count once per (day, root, rep)

    countedAck.add(key);
    done7ByRep.set(rep, (done7ByRep.get(rep) || 0) + 1);
  });

  // Split into team tables
  const t7AssignedRows = [];
  const t7AssistedRows = [];
  expected7ByRep.forEach((exp, rep) => {
    if (!exp) return;
    const role = repRole.get(rep) || '';
    const done = (done7ByRep.get(rep) || 0);
    const pct = done / exp;
    const row = [rep, exp, done, pct];

    if (equalsIgnoreCase_(role, 'Assigned')) t7AssignedRows.push(row);
    else if (equalsIgnoreCase_(role, 'Assisted')) t7AssistedRows.push(row);
  });

  const sorter7 = (a,b) => {
    const p = (a[3]||0) - (b[3]||0);
    if (p !== 0) return p;
    const e = (b[1]||0) - (a[1]||0);
    if (e !== 0) return e;
    return String(a[0]).localeCompare(String(b[0]));
  };
  t7AssignedRows.sort(sorter7);
  t7AssistedRows.sort(sorter7);

  // ---------- Scope changes today (info only) ----------
  const rootsWithLogToday = new Set(
    log06
      .filter(r => {
        const d = r['Log Date'];
        const dk = (d instanceof Date) ? Utilities.formatDate(d, tz, 'yyyy-MM-dd')
                                       : String(d || '').slice(0,10);
        return dk === todayKey;
      })
      .map(r => String(r['RootApptID'] || '').trim())
      .filter(Boolean)
  );
  let scopeChangesToday = 0;
  rootsWithLogToday.forEach(root => { if (!inScope.has(root)) scopeChangesToday++; });

  // ---------- RENDER ----------
  s09.clearContents();

  // HARD RESET: wipe all formatting on the whole sheet,
  // remove banding + conditional formats + stray filter views.
  resetSheetFormatting_(s09);

  let row = 1;

  // Title
  s09.getRange(row,1).setValue('Ack Dashboard — Today').setFontWeight('bold');
  s09.getRange(row,2).setValue(Utilities.formatDate(new Date(), tz, 'EEE, MMM d, yyyy h:mm a'));
  s09.getRange(row,4).setValue('Scope changes today (logs on out-of-scope roots)').setFontWeight('bold');
  s09.getRange(row,5).setValue(scopeChangesToday);
  row += 2;

  // ===== Overload Today — Assigned (Top 5) =====
  s09.getRange(row,1).setValue('Overload Today — Assigned (Snapshot, Top 5)').setFontWeight('bold');
  row++;
  const aoHeaders = ['Rep','Expected (A)','Team Median (A)','AOI'];
  writeTable_(s09, row, 1, aoHeaders, assignedOverloadRows);
  if (assignedOverloadRows.length) s09.getRange(row+1, 4, assignedOverloadRows.length, 1).setNumberFormat('0.00');
  applyRowBandingSafe_(s09, row, 1,
    Math.max(1, assignedOverloadRows.length) + 1,
    aoHeaders.length,
    SpreadsheetApp.BandingTheme.LIGHT_GREY
  );
  const allCF = []; // collect CF rules; set once at the end
  if (assignedOverloadRows.length) {
    const rAOI = s09.getRange(row+1, 4, assignedOverloadRows.length, 1);
    allCF.push(
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(1.25).setBackground('#F8CBAD').setRanges([rAOI]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(1.10, 1.25).setBackground('#FFEB9C').setRanges([rAOI]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThanOrEqualTo(1.10).setBackground('#C6EFCE').setRanges([rAOI]).build()
    );
  }
  row += Math.max(2, assignedOverloadRows.length + 2);

  // ===== Overload Today — Assisted (Top 5) =====
  s09.getRange(row,1).setValue('Overload Today — Assisted (Snapshot, Top 5)').setFontWeight('bold');
  row++;
  const asoHeaders = ['Rep','Expected (Asst)','Team Median (Asst)','AsstOI'];
  writeTable_(s09, row, 1, asoHeaders, assistedOverloadRows);
  if (assistedOverloadRows.length) s09.getRange(row+1, 4, assistedOverloadRows.length, 1).setNumberFormat('0.00');
  applyRowBandingSafe_(s09, row, 1,
    Math.max(1, assistedOverloadRows.length) + 1,
    asoHeaders.length,
    SpreadsheetApp.BandingTheme.LIGHT_GREY
  );
  if (assistedOverloadRows.length) {
    const rAsstOI = s09.getRange(row+1, 4, assistedOverloadRows.length, 1);
    allCF.push(
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(1.25).setBackground('#F8CBAD').setRanges([rAsstOI]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(1.10, 1.25).setBackground('#FFEB9C').setRanges([rAsstOI]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThanOrEqualTo(1.10).setBackground('#C6EFCE').setRanges([rAsstOI]).build()
    );
  }
  row += Math.max(2, assistedOverloadRows.length + 2);

  // ===== Compliance — Real-time =====
  s09.getRange(row,1).setValue('Compliance Today (by Rep) — Real-time').setFontWeight('bold');
  row++;
  const compHeaders = ['Rep','Expected','Fully Updated','Needs follow-up','Missing','% Fully Updated'];
  writeTable_(s09, row, 1, compHeaders, complianceRows);
  if (complianceRows.length) s09.getRange(row+1, 6, complianceRows.length, 1).setNumberFormat('0.0%');
  applyRowBandingSafe_(s09, row, 1,
    Math.max(1, complianceRows.length) + 1,
    compHeaders.length,
    SpreadsheetApp.BandingTheme.LIGHT_GREY
  );

  // (Compliance CF already handled in your Phase D helper; omitting here to keep one CF set)
  row += Math.max(2, complianceRows.length + 2);

  // ===== Compliance — Snapshot =====
  if (snapComplianceRows.length) {
    const label = captureTime ? `Compliance (Snapshot @ ${Utilities.formatDate(captureTime, tz, 'h:mm a')})`
                              : 'Compliance (Snapshot — today)';
    s09.getRange(row,1).setValue(label).setFontWeight('bold');
    row++;
    writeTable_(s09, row, 1, compHeaders, snapComplianceRows);
    if (snapComplianceRows.length) s09.getRange(row+1, 6, snapComplianceRows.length, 1).setNumberFormat('0.0%');
    applyRowBandingSafe_(s09, row, 1,
      Math.max(1, snapComplianceRows.length) + 1,
      compHeaders.length,
      SpreadsheetApp.BandingTheme.LIGHT_GREY
    );
    row += Math.max(2, snapComplianceRows.length + 2);
  } else {
    s09.getRange(row,1).setValue('Compliance (Snapshot) — No snapshot taken today yet').setFontStyle('italic');
    row += 2;
  }

  // ===== Missing Acks (today) =====
  s09.getRange(row,1).setValue('Missing Acks (today)').setFontWeight('bold');
  row++;
  const missHeaders = ['Rep','RootApptID','Customer Name','Updated By','Updated At','Days Since Last Update','Client Status Report URL'];
  writeTable_(s09, row, 1, missHeaders, missingRows);
  if (missingRows.length) {
    const colUpdatedAt = missHeaders.indexOf('Updated At') + 1;
    s09.getRange(row+1, colUpdatedAt, missingRows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
  applyRowBandingSafe_(s09, row, 1,
    Math.max(1, missingRows.length) + 1,
    missHeaders.length,
    SpreadsheetApp.BandingTheme.LIGHT_GREY
  );
  row += Math.max(2, missingRows.length + 2);

  // ===== Needs Follow-up (last N days) =====
  s09.getRange(row,1).setValue(`Needs Follow-up (last ${NEEDS_FOLLOWUP_LOOKBACK_DAYS} days)`).setFontWeight('bold');
  row++;
  const needHeaders = ['Rep','RootApptID','Customer Name','Flagged At','Resolved?','Note','Custom Order Status (at log)','Last Updated At (at log)','Client Status Report URL'];
  writeTable_(s09, row, 1, needHeaders, needsRows);
  if (needsRows.length) {
    const colTs = needHeaders.indexOf('Flagged At') + 1;
    const colUpdAt = needHeaders.indexOf('Last Updated At (at log)') + 1;
    s09.getRange(row+1, colTs, needsRows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    s09.getRange(row+1, colUpdAt, needsRows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
  applyRowBandingSafe_(s09, row, 1,
    Math.max(1, needsRows.length) + 1,
    needHeaders.length,
    SpreadsheetApp.BandingTheme.LIGHT_GREY
  );
  row += Math.max(2, needsRows.length + 2);

  // ===== Stale Updates =====
  s09.getRange(row,1).setValue(`Stale Updates (≥ ${STALE_UPDATE_DAYS_THRESHOLD} days)`).setFontWeight('bold');
  row++;
  const staleHeaders = ['RootApptID','Customer Name','Days Since Last Update','Updated By','Updated At','Client Status Report URL'];
  writeTable_(s09, row, 1, staleHeaders, staleRows);
  if (staleRows.length) {
    const colUpdAt = staleHeaders.indexOf('Updated At') + 1;
    s09.getRange(row+1, colUpdAt, staleRows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
  applyRowBandingSafe_(s09, row, 1,
    Math.max(1, staleRows.length) + 1,
    staleHeaders.length,
    SpreadsheetApp.BandingTheme.LIGHT_GREY
  );
  row += Math.max(2, staleRows.length + 2);

  // ===== Trailing 7‑Day Completion — Assigned =====
  s09.getRange(row,1).setValue('Trailing 7-Day Completion — Assigned').setFontWeight('bold');
  row++;
  const t7Headers = ['Rep','7d Expected','7d Fully Updated','7d %'];
  writeTable_(s09, row, 1, t7Headers, t7AssignedRows);
  if (t7AssignedRows.length) s09.getRange(row+1, 4, t7AssignedRows.length, 1).setNumberFormat('0.0%');
  applyRowBandingSafe_(s09, row, 1,
    Math.max(1, t7AssignedRows.length) + 1,
    t7Headers.length,
    SpreadsheetApp.BandingTheme.LIGHT_GREY
  );
  if (t7AssignedRows.length) {
    const rPct = s09.getRange(row+1, 4, t7AssignedRows.length, 1);
    allCF.push(
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.90).setBackground('#C6EFCE').setRanges([rPct]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.80, 0.8999).setBackground('#FFEB9C').setRanges([rPct]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.80).setBackground('#F8CBAD').setRanges([rPct]).build()
    );
  }
  row += Math.max(2, t7AssignedRows.length + 2);

  // ===== Trailing 7‑Day Completion — Assisted =====
  s09.getRange(row,1).setValue('Trailing 7-Day Completion — Assisted').setFontWeight('bold');
  row++;
  writeTable_(s09, row, 1, t7Headers, t7AssistedRows);
  if (t7AssistedRows.length) s09.getRange(row+1, 4, t7AssistedRows.length, 1).setNumberFormat('0.0%');
  applyRowBandingSafe_(s09, row, 1,
    Math.max(1, t7AssistedRows.length) + 1,
    t7Headers.length,
    SpreadsheetApp.BandingTheme.LIGHT_GREY
  );
  if (t7AssistedRows.length) {
    const rPct = s09.getRange(row+1, 4, t7AssistedRows.length, 1);
    allCF.push(
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.85).setBackground('#C6EFCE').setRanges([rPct]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.75, 0.8499).setBackground('#FFEB9C').setRanges([rPct]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.75).setBackground('#F8CBAD').setRanges([rPct]).build()
    );
  }
  row += Math.max(2, t7AssistedRows.length + 2);

  // ===== Coverage Gaps =====
  const assignedGapRows = [];
  assignedGaps.forEach((obj, root) => {
    const snap = snapByRoot.get(root) || {};
    assignedGapRows.push([
      root,
      String(snap['Customer Name'] || '').trim(),
      (obj.assigned || []).join(', '),
      'Both assigned off'
    ]);
  });
  if (assignedGapRows.length) {
    s09.getRange(row,1).setValue('Assigned Coverage Gaps (today)').setFontWeight('bold');
    row++;
    const agHeaders = ['RootApptID','Customer Name','Assigned Reps','Reason'];
    writeTable_(s09, row, 1, agHeaders, assignedGapRows);
    applyRowBandingSafe_(s09, row, 1,
      Math.max(1, assignedGapRows.length) + 1,
      agHeaders.length,
      SpreadsheetApp.BandingTheme.LIGHT_GREY
    );
    row += Math.max(2, assignedGapRows.length + 2);
  }

  const assistedGapRows = [];
  assistedGaps.forEach((obj, root) => {
    const snap = snapByRoot.get(root) || {};
    assistedGapRows.push([
      root,
      String(snap['Customer Name'] || '').trim(),
      obj.pair || 'Maria & Paul',
      'Both assisted off'
    ]);
  });
  if (assistedGapRows.length) {
    s09.getRange(row,1).setValue('Assisted Coverage Gaps (today)').setFontWeight('bold');
    row++;
    const asgHeaders = ['RootApptID','Customer Name','Assisted Pair','Reason'];
    writeTable_(s09, row, 1, asgHeaders, assistedGapRows);
    applyRowBandingSafe_(s09, row, 1,
      Math.max(1, assistedGapRows.length) + 1,
      asgHeaders.length,
      SpreadsheetApp.BandingTheme.LIGHT_GREY
    );
    row += Math.max(2, assistedGapRows.length + 2);
  }

  // Set all CF once (faster than 3–4 separate calls)
  if (allCF.length) s09.setConditionalFormatRules(allCF);
}

/* small utility */
function median_(arr) {
  const a = arr.slice().sort((x,y) => x - y);
  const n = a.length;
  if (!n) return 0;
  const mid = Math.floor(n/2);
  return (n % 2) ? a[mid] : (a[mid-1] + a[mid]) / 2;
}

/** Utility to write a table with headers at (row,col) */
function writeTable_(sheet, row, col, headers, rows) {
  if (!headers || !headers.length) return;
  sheet.getRange(row, col, 1, headers.length).setValues([headers]).setFontWeight('bold');
  if (rows && rows.length) {
    sheet.getRange(row+1, col, rows.length, headers.length).setValues(rows);
  }
}


/** Apply conditional formatting and auto-resize (no banding here) */
function applyDashboardFormatting_(sheet, pos) {
  const rules = [];

  // Clear old banding only if you want to manage banding elsewhere
  // sheet.getBandings().forEach(b => b.remove()); // <-- leave commented; we band inline already

  // === Compliance: % Fully Updated (col 6), Missing (col 5) ===
  if (pos.compStart.rows > 0) {
    const compPctRange = sheet.getRange(pos.compStart.row + 1, pos.compStart.col + 5, pos.compStart.rows, 1); // % col
    const compMissingRange = sheet.getRange(pos.compStart.row + 1, pos.compStart.col + 4, pos.compStart.rows, 1); // Missing col

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(0.9).setBackground('#C6EFCE').setRanges([compPctRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(0.6, 0.9).setBackground('#FFEB9C').setRanges([compPctRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0.6).setBackground('#F8CBAD').setRanges([compPctRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0).setBackground('#F8CBAD').setRanges([compMissingRange]).build()
    );
  }

  // === Missing Acks: Days Since Last Update (col 6) ===
  if (pos.missStart.rows > 0) {
    const missDaysRange = sheet.getRange(pos.missStart.row + 1, pos.missStart.col + 5, pos.missStart.rows, 1);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(5).setBackground('#F8CBAD').setRanges([missDaysRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(3, 4).setBackground('#FFEB9C').setRanges([missDaysRange]).build()
    );
  }

  // === Needs Follow-up: Resolved? (col 5) ===
  if (pos.needStart.rows > 0) {
    const needResolvedRange = sheet.getRange(pos.needStart.row + 1, pos.needStart.col + 4, pos.needStart.rows, 1);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Yes').setBackground('#C6EFCE').setRanges([needResolvedRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('No').setBackground('#F8CBAD').setRanges([needResolvedRange]).build()
    );
  }

  // === Stale Updates: Days Since Last Update (col 3) ===
  if (pos.staleStart.rows > 0) {
    const staleDaysRange = sheet.getRange(pos.staleStart.row + 1, pos.staleStart.col + 2, pos.staleStart.rows, 1);
    const y = STALE_UPDATE_DAYS_THRESHOLD;
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(y + 2).setBackground('#F8CBAD').setRanges([staleDaysRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(y, y + 1).setBackground('#FFEB9C').setRanges([staleDaysRange]).build()
    );
  }

  sheet.setConditionalFormatRules(rules);

  // Auto-resize columns for readability
  const lastCol = Math.max(
    pos.compStart.col + pos.compStart.cols - 1,
    pos.missStart.col + pos.missStart.cols - 1,
    pos.needStart.col + pos.needStart.cols - 1,
    pos.staleStart.col + pos.staleStart.cols - 1
  );
  sheet.autoResizeColumns(1, lastCol);
}

function rangesOverlap_(a, b) {
  const aTop = a.getRow(), aLeft = a.getColumn();
  const aBottom = aTop + a.getNumRows() - 1;
  const aRight = aLeft + a.getNumColumns() - 1;

  const bTop = b.getRow(), bLeft = b.getColumn();
  const bBottom = bTop + b.getNumRows() - 1;
  const bRight = bLeft + b.getNumColumns() - 1;

  const rowsOverlap = !(aBottom < bTop || bBottom < aTop);
  const colsOverlap = !(aRight  < bLeft || bRight  < aLeft);
  return rowsOverlap && colsOverlap;
}

function applyRowBandingSafe_(sheet, startRow, startCol, numRows, numCols, theme) {
  // Normalize sizes
  const rows = Math.max(1, numRows | 0);
  const cols = Math.max(1, numCols | 0);
  const target = sheet.getRange(startRow, startCol, rows, cols);

  // --- helpers ---
  function paintHeader_() {
    const headerRange = sheet.getRange(startRow, startCol, 1, cols);
    headerRange
      .setBackground(DASH_HEADER_BG)  // header bg (pink)
      .setFontColor(DASH_HEADER_FG)   // header text (white)
      .setFontWeight('bold');         // header bold
  }
  function normalizeBody_() {
    const bodyRows = rows - 1;
    if (bodyRows <= 0) return;
    const bodyRange = sheet.getRange(startRow + 1, startCol, bodyRows, cols);
    // Reset body to defaults; keeps links blue by using null.
    bodyRange.setFontColor(null).setFontWeight('normal').setBackground(null);
  }
  function clearHeaderRowTail_() {
    // Clear any leftover formatting to the RIGHT of the header cells
    // (prevents the pink bar from continuing past the table).
    const lastCol = sheet.getLastColumn(); // used width on the sheet
    const tail = lastCol - (startCol + cols - 1);
    if (tail > 0) {
      sheet.getRange(startRow, startCol + cols, 1, tail)
           .setBackground(null)
           .setFontColor(null)
           .setFontWeight(null);
    }
  }

  // 0) If an exact band exists already, just restyle it (idempotent fast‑path)
  const exact = sheet.getBandings().find(b => {
    const r = b.getRange();
    return r.getRow() === startRow &&
           r.getColumn() === startCol &&
           r.getNumRows() === rows &&
           r.getNumColumns() === cols;
  });
  if (exact) {
    try {
      if (typeof exact.setHeaderRowColor === 'function') exact.setHeaderRowColor(DASH_HEADER_BG);
      if (typeof exact.setFirstBandColor  === 'function') exact.setFirstBandColor(DASH_BAND_1);
      if (typeof exact.setSecondBandColor === 'function') exact.setSecondBandColor(DASH_BAND_2);
    } catch (_) {}
    paintHeader_();
    normalizeBody_();
    clearHeaderRowTail_();
    return exact;
  }

  // 1) Remove any overlapping bandings first
  sheet.getBandings().forEach(b => { if (rangesOverlap_(target, b.getRange())) b.remove(); });

  // 2) Clear any leftover formatting WITHIN the table footprint
  target.setBackground(null).setFontColor(null).setFontWeight(null);

  // 3) Apply banding (retry once if Sheets complains about existing banding)
  let band;
  try {
    band = target.applyRowBanding(theme || SpreadsheetApp.BandingTheme.LIGHT_GREY);
  } catch (e) {
    if (String(e).toLowerCase().includes('alternating background colors')) {
      sheet.getBandings().forEach(b => b.remove());
      band = target.applyRowBanding(theme || SpreadsheetApp.BandingTheme.LIGHT_GREY);
    } else {
      throw e;
    }
  }

  // 4) Color the band (API sometimes returns void in older runtimes; no chaining)
  try {
    if (band && typeof band.setHeaderRowColor === 'function') band.setHeaderRowColor(DASH_HEADER_BG);
    if (band && typeof band.setFirstBandColor  === 'function') band.setFirstBandColor(DASH_BAND_1);
    if (band && typeof band.setSecondBandColor === 'function') band.setSecondBandColor(DASH_BAND_2);
  } catch (_) {}

  // 5) Final explicit styles
  paintHeader_();        // header: pink bg + white bold text
  normalizeBody_();      // body: default colors + normal weight
  clearHeaderRowTail_(); // clear any pink “tail” to the right of the table

  return band;
}


/** OPTIONAL: create separate Filter Views per section (Advanced Sheets API) */
function createDashboardFilterViews_(sheet, pos) {
  const ss = MASTER_SS_();
  const spreadsheetId = ss.getId();
  const sheetId = sheet.getSheetId();

  // Clear existing filter views on this sheet
  const meta = Sheets.Spreadsheets.get(spreadsheetId, {fields: 'sheets.properties,sheets.filterViews'});
  const current = (meta.sheets || []).find(s => s.properties && s.properties.sheetId === sheetId);
  const delReqs = [];
  if (current && current.filterViews) {
    current.filterViews.forEach(v => {
      delReqs.push({ deleteFilterView: { filterId: v.filterViewId } });
    });
  }

  // Helper: range object
  const r = (start) => ({
    sheetId,
    startRowIndex: start.row - 1,
    endRowIndex: start.row - 1 + (start.rows + 1), // +1 header
    startColumnIndex: start.col - 1,
    endColumnIndex: start.col - 1 + start.cols
  });

  // Add one filter view per table
  const addReqs = [
    {
      addFilterView: {
        filter: {
          title: 'Compliance Today (by Rep)',
          range: r(pos.compStart)
        }
      }
    },
    {
      addFilterView: {
        filter: {
          title: 'Missing Acks (today)',
          range: r(pos.missStart)
        }
      }
    },
    {
      addFilterView: {
        filter: {
          title: `Needs Follow-up (last ${NEEDS_FOLLOWUP_LOOKBACK_DAYS} days)`,
          range: r(pos.needStart)
        }
      }
    },
    {
      addFilterView: {
        filter: {
          title: `Stale Updates (≥ ${STALE_UPDATE_DAYS_THRESHOLD} days)`,
          range: r(pos.staleStart)
        }
      }
    }
  ];

  // Batch delete (if any) then add
  const requests = delReqs.concat(addReqs);
  if (requests.length) {
    Sheets.Spreadsheets.batchUpdate({requests}, spreadsheetId);
  }
}

/** Windowed reader for append-only logs:
 *  Reads only rows where DATE_HEADER >= (today - lookbackDays).
 *  Expects the date column to contain Date values or ISO-like strings.
 */
function getObjectsByDateWindow_(sheet, DATE_HEADER, lookbackDays) {
  const ss = MASTER_SS_();
  const tz = ss.getSpreadsheetTimeZone();
  const headers = getHeaders_(sheet);
  const idxDate = headers.indexOf(DATE_HEADER);
  if (idxDate < 0) throw new Error(`${sheet.getName()}: header not found → ${DATE_HEADER}`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const minTime = new Date().getTime() - (Math.max(1, lookbackDays) * 86400000);
  // Read ONLY the date column to find the first row in the window
  const dateVals = sheet.getRange(2, idxDate+1, lastRow-1, 1).getValues(); // rows 2..last
  let startOffset = dateVals.length; // default: nothing in window
  for (let i = dateVals.length - 1; i >= 0; i--) {
    const v = dateVals[i][0];
    const d = (v instanceof Date) ? v : (v ? new Date(v) : null);
    if (d && !isNaN(d) && d.getTime() >= minTime) startOffset = i; // keep pushing up
    else if (startOffset < dateVals.length) break; // we’ve crossed the boundary
  }

  const startRow = (startOffset < dateVals.length) ? (2 + startOffset) : (lastRow + 1);
  if (startRow > lastRow) return []; // nothing in window

  const numRows = lastRow - startRow + 1;
  const values = sheet.getRange(startRow, 1, numRows, headers.length).getValues();

  // Map to objects using headers
  const out = new Array(values.length);
  for (let r = 0; r < values.length; r++) {
    const obj = {};
    const row = values[r];
    for (let c = 0; c < headers.length; c++) obj[headers[c]] = row[c];
    out[r] = obj;
  }
  return out;
}

function resetSheetFormatting_(sheet) {
  // 1) Clear ALL formats on the full canvas (not just the used range)
  var maxR = sheet.getMaxRows(), maxC = sheet.getMaxColumns();
  sheet.getRange(1, 1, maxR, maxC).clear({ formatOnly: true });

  // 2) Remove banding & conditional formats
  try { sheet.getBandings().forEach(function(b){ b.remove(); }); } catch (_) {}
  try { sheet.setConditionalFormatRules([]); } catch (_) {}

  // 3) (Optional but thorough) Remove all Filter Views on this sheet
  try {
    var ss = sheet.getParent();
    var spreadsheetId = ss.getId();
    var meta = Sheets.Spreadsheets.get(spreadsheetId, {fields: 'sheets.properties,sheets.filterViews'});
    var cur = (meta.sheets || []).find(function(s){ return s.properties && s.properties.sheetId === sheet.getSheetId(); });
    if (cur && cur.filterViews && cur.filterViews.length) {
      var reqs = cur.filterViews.map(function(v){ return { deleteFilterView: { filterId: v.filterViewId } }; });
      Sheets.Spreadsheets.batchUpdate({ requests: reqs }, spreadsheetId);
    }
  } catch (_) {}
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




