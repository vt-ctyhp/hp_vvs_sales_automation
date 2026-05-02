/**
 * 12.12 — Reminders Heal (Verbose Instrumented, Resumable, Batched, Checkpointed)
 *
 * Behavior (unchanged from 12.11):
 *   • Rebuild/repair reminders from "00_Master Appointments":
 *       - COS while COS is in 3D‑pending set
 *       - Follow‑Up while Sales Stage = "Follow‑Up Required"
 *       - DV Propose Nudge (one per appointment day)
 *       - DV Urgent Daily for today if appt within next 7 days and not OTW/final
 *   • Reconcile queue: CONFIRMS stale COS/FU/DV_URGENT rows only
 *   • Normalize DV date fields to canonical PT strings (chunked)
 *
 * Instrumentation added:
 *   • Very detailed audit logging at every step (headers found, counts, offsets, batches, time left).
 *   • Explicit logging of resume trigger creation/removal and state saves.
 *   • Per‑phase candidate counts + first few samples for sanity.
 *
 * Safe/idempotent:
 *   • No schema changes. Header-by-name. No deletes.
 *   • Writes only via your existing Remind/DV helpers and queue confirm/normalize.
 *   • State kept in Script Properties. Execution-time aware. Resumable.
 */

var HEALX = Object.freeze({
  TZ: (typeof REMIND !== 'undefined' ? REMIND.TIMEZONE : 'America/Los_Angeles'),
  MASTER_SHEET: (typeof REMIND !== 'undefined' ? REMIND.ORDERS_SHEET_NAME : '00_Master Appointments'),
  QUEUE_SHEET:  (typeof REMIND !== 'undefined' ? REMIND.QUEUE_SHEET_NAME  : '04_Reminders_Queue'),
  AUDIT_SHEET:  '17_Reminders_Heal_Audit',

  // Batching (unchanged defaults)
  BATCH_SO:         60,   // P1A: COS/FU per batch
  BATCH_DV:         60,   // P1B/P1C: DV items per batch
  BATCH_Q:          200,  // P2: queue rows per chunk
  BATCH_NORM:       200,  // P3: normalization rows per chunk
  BATCH_ENRICH:    200,  // NEW: P1D — DV queue enrichment rows per chunk

  // Progress pings (unchanged)
  PROGRESS_PING_A:  60,
  PROGRESS_PING_B:  60,
  PROGRESS_PING_Q:  400,

  // Execution budget
  EXEC_MAX_MS:      330000,
  EXEC_STOP_EARLY:   40000,

  // Script Properties keys
  STATE_KEY:     'HEALX_STATE_V1',
  RUN_FLAG:      'HEALX_RUN_LOCK',
  RUN_ID_KEY:    'HEALX_RUN_ID',

  // Types eligible for auto-confirm in reconcile
  RECONCILE_TYPES: Object.freeze(['COS','FOLLOWUP','DV_URGENT_OTW_DAILY']),

  // DV detector
  DV_REGEX: /diamond\s*viewing/i,

  // Resume handler
  RESUME_FN: 'RemindersHeal_resume',

  // Debug property key
  DEBUG_PROP: 'HEALX_DEBUG',   // '1' = on (default), '0' = off
});

// === Manual mode gate (Script Property 'HEALX_MANUAL': '1' = manual, else auto) ===
function _hx__manualMode_() {
  try {
    var v = PropertiesService.getScriptProperties().getProperty('HEALX_MANUAL');
    return String(v || '').trim() === '1';
  } catch (_) { return false; }
}

/* =========================
 * Public commands
 * ========================= */

function RemindersHeal_start()  { RemindersHeal__runStep_({ resetIfIdle: true  }); }
function RemindersHeal_resume() { RemindersHeal__runStep_({ resetIfIdle: false }); }

function RemindersHeal_cancel() {
  var p = PropertiesService.getScriptProperties();
  p.deleteProperty(HEALX.STATE_KEY);
  p.deleteProperty(HEALX.RUN_FLAG);
  p.deleteProperty(HEALX.RUN_ID_KEY);
  _hx__killExistingResumeTriggers_();
  _hx__audit('INFO', 'Heal — CANCELLED by user. State cleared.');
}

function RemindersHeal_status() {
  var s = _hx__loadState_();
  Logger.log(JSON.stringify(s || {}, null, 2));
  _hx__audit('INFO', 'STATUS: ' + JSON.stringify(s || {}, null, 2));
  return s;
}

/* =========================
 * Runner (time‑boxed step)
 * ========================= */

function RemindersHeal__runStep_(opts) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    _hx__audit('WARN', 'Runner exit: another instance holds the lock (likely mid‑batch).');
    return;
  }

  var started = Date.now();
  try {
    var state = _hx__initOrLoadState_(opts && opts.resetIfIdle);
    var budget = {
      startedMs: started,
      deadlineMs: started + HEALX.EXEC_MAX_MS - HEALX.EXEC_STOP_EARLY
    };

    _hx__dbg('RUN', { phase: state.phase, offsets: {
      p1a: state.p1aOffset, p1b: state.p1bOffset, p1c: state.p1cOffset, p2: state.p2Offset, p3: state.p3Offset
    }, totals: {
      p1a: state.p1aTotal, p1b: state.p1bTotal, p1c: state.p1cTotal, p2: state.p2Total, p3: state.p3Total
    }, msLeft: budget.deadlineMs - Date.now() });

    if (state.phase === 'P1A')      _hx__phase1A_COS_FU_(state, budget);
    else if (state.phase === 'P1B') _hx__phase1B_DV_PROPOSE_(state, budget);
    else if (state.phase === 'P1C') _hx__phase1C_DV_URGENT_(state, budget);
    // NEW:
    else if (state.phase === 'P1D') _hx__phase1D_ENRICH_DV_REPS_(state, budget);
    // ---
    else if (state.phase === 'P2')  _hx__phase2_RECONCILE_(state, budget);
    else if (state.phase === 'P3')  _hx__phase3_NORMALIZE_(state, budget);
    else if (state.phase === 'DONE') {
      _hx__audit('INFO', 'Heal — DONE (nothing further).');
      _hx__clearState_();
      return;
    } else {
      state.phase = 'P1A';
      _hx__saveState_(state);
      _hx__audit('INFO', 'Heal — state missing phase; defaulted to P1A.');
      _hx__scheduleResume_(10);
      return;
    }

    // Persist state every step
    _hx__saveState_(state);

    if (state.phase === 'DONE') {
      _hx__audit('INFO', 'Heal — FINISHED all phases.');
      _hx__clearState_();
    } else {
      if (_hx__manualMode_()) {
        _hx__audit('INFO', 'Manual step complete — click "Run next batch" to continue. phase=' + state.phase);
      } else {
        _hx__scheduleResume_(10);
      }
    }

  } catch (e) {
    _hx__audit('ERROR', 'Heal crashed: ' + (e && e.stack ? e.stack : e));
    _hx__scheduleResume_(30); // retry later; keep state
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* =========================
 * Phase implementations
 * ========================= */

function _hx__phase1A_COS_FU_(state, budget) {
  _hx__audit('INFO', 'P1A — start (COS/FU). Time left ~' + (budget.deadlineMs - Date.now()) + 'ms');

  var M = _hx__readMaster_();
  // Build unique SO list (and count blank SO rows for logging)
  var mapSO = new Map();
  var blankSO = 0; // <-- NEW

  for (var i = 0; i < M.rows.length; i++) {
    var R = M.rows[i];
    if (!R.soKey) { blankSO++; continue; } // <-- NEW

    var prev = mapSO.get(R.soKey);
    if (!prev) {
      mapSO.set(R.soKey, R);
      continue;
    }

    // Helper: keep prev if it already has a value, else take from R
    function fill(prevVal, newVal) {
      return (prevVal && String(prevVal).trim()) ? prevVal : (newVal || '');
    }

    // Merge missing fields from newer row
    prev.assignedRepName  = fill(prev.assignedRepName,  R.assignedRepName);
    prev.assistedRepName  = fill(prev.assistedRepName,  R.assistedRepName);
    prev.assignedRepEmail = fill(prev.assignedRepEmail, R.assignedRepEmail);
    prev.assistedRepEmail = fill(prev.assistedRepEmail, R.assistedRepEmail);

    prev.customerName = fill(prev.customerName, R.customerName);
    prev.nextSteps    = fill(prev.nextSteps,    R.nextSteps);
    prev.salesStage   = fill(prev.salesStage,   R.salesStage);
    prev.cos          = fill(prev.cos,          R.cos);
    prev.csos         = fill(prev.csos,         R.csos);
    prev.rootId       = fill(prev.rootId,       R.rootId);
    prev.apptISO      = fill(prev.apptISO,      R.apptISO);
  }

  // Materialize final list
  var list = Array.from(mapSO.values());

  if (!state.p1aTotal) {
    state.p1aTotal = list.length;
    _hx__audit('INFO', 'P1A — unique SO count: ' + state.p1aTotal + ' (blank SO rows: ' + blankSO + ')'); // <-- now valid
    if (list.length) _hx__dbg('P1A_SAMPLE', { firstSO: (list[0].soPretty || list[0].soRaw), customer: list[0].customerName });
    _hx__saveState_(state);
  }


  var off = state.p1aOffset || 0;
  if (off >= list.length) {
    state.phase = 'P1B'; state.p1bOffset = 0; state.p1bTotal = 0;
    _hx__audit('INFO', 'P1A finished — processed ' + off + '/' + list.length + ' SO(s).');
    return;
  }

  var max = Math.min(off + HEALX.BATCH_SO, list.length);
  _hx__dbg('P1A_BATCH', { from: off, to: max, size: (max-off) });

  for (var i = off; i < max; i++) {
    var R = list[i];
    var opts = {
      assignedRepName:  R.assignedRepName || '',
      assistedRepName:  R.assistedRepName || '',
      assignedRepEmail: R.assignedRepEmail || '',
      assistedRepEmail: R.assistedRepEmail || '',
      customerName:     R.customerName || '',
      nextSteps:        R.nextSteps || ''
    };

    try {
      if (typeof Remind !== 'undefined' && typeof Remind.onClientStatusChange === 'function') {
        Remind.onClientStatusChange(R.soPretty || R.soRaw || '', R.salesStage || '', R.cos || '', '', opts);
      } else {
        if (_hx__is3DPendingCOS_(R.cos) && typeof Remind !== 'undefined' && typeof Remind.scheduleCOS === 'function') {
          Remind.scheduleCOS(R.soPretty || R.soRaw || '', opts, false);
        }
        if (_hx__eqCI(R.salesStage, 'Follow-Up Required') && typeof Remind !== 'undefined' && typeof Remind.ensureFollowUp === 'function') {
          Remind.ensureFollowUp(R.soPretty || R.soRaw || '', opts);
        }
      }
    } catch (e) {
      _hx__audit('WARN', 'P1A: upsert failed for SO ' + (R.soPretty || R.soRaw) + ' — ' + e);
    }

    if ((i + 1) % HEALX.PROGRESS_PING_A === 0) {
      _hx__audit('INFO', 'P1A progress: ' + (i + 1) + '/' + list.length + ' SO(s).');
    }
    if (Date.now() >= budget.deadlineMs) { i++; max = i; break; }
  }

  SpreadsheetApp.flush(); Utilities.sleep(120);
  state.p1aOffset = max;

  if (state.p1aOffset >= list.length) {
    state.phase = 'P1B'; state.p1bOffset = 0; state.p1bTotal = 0;
    _hx__audit('INFO', 'P1A complete: ' + list.length + ' SO(s).');
  }
}

function _hx__phase1B_DV_PROPOSE_(state, budget) {
  _hx__audit('INFO', 'P1B — start (DV PROPOSE). Time left ~' + (budget.deadlineMs - Date.now()) + 'ms');

  var M = _hx__readMaster_();
  var items = [];
  var dvRows = 0, dvWithISO = 0, dvWithRoot = 0;
  for (var i=0;i<M.rows.length;i++){
    var R = M.rows[i];
    if (R.isDV) dvRows++;
    if (!R.isDV || !R.apptISO || !R.rootId) continue;
    dvWithISO++; dvWithRoot += (R.rootId ? 1 : 0);
    items.push(R);
  }

  if (!state.p1bTotal) {
    state.p1bTotal = items.length;
    _hx__audit('INFO', 'P1B — DV rows: ' + dvRows + ', candidates: ' + items.length +
                      ' (withISO=' + dvWithISO + ', withRoot=' + dvWithRoot + ')');
    if (items.length) _hx__dbg('P1B_SAMPLE', { root: items[0].rootId, apptISO: items[0].apptISO, cust: items[0].customerName });
  }

  var off = state.p1bOffset || 0;
  if (off >= items.length) {
    state.phase = 'P1C'; state.p1cOffset = 0; state.p1cTotal = 0;
    _hx__audit('INFO', 'P1B finished — processed ' + off + '/' + items.length + ' appointment(s).');
    return;
  }

  var max = Math.min(off + HEALX.BATCH_DV, items.length);
  _hx__dbg('P1B_BATCH', { from: off, to: max, size: (max-off) });

  for (var i2 = off; i2 < max; i2++) {
    var R2 = items[i2];
    try {
      if (typeof DVQ_upsert_ === 'function' && typeof DVQ_id_ === 'function') {
        var apptDayKey = _hx__dayKeyFromISO_(R2.apptISO);
        var dueLocal   = _hx__localAt_(_hx__addDays_(new Date(R2.apptISO), -_hx__policy().PROPOSE_NUDGE_OFFSET_DAYS), 9, 0);
        DVQ_upsert_({
          id:   DVQ_id_(R2.rootId, _hx__remType().PROPOSE_NUDGE, apptDayKey),
          type: _hx__remType().PROPOSE_NUDGE,
          dueAt: dueLocal,
          customerName: R2.customerName || '',
          nextSteps: R2.nextSteps || '',
          dvNotes: 'Heal: 12 days before appointment',
          status: 'PENDING'
        });
      } else {
        _hx__audit('WARN', 'P1B: DV helpers missing for root ' + R2.rootId);
      }
    } catch (e) {
      _hx__audit('WARN', 'P1B: upsert failed for root ' + R2.rootId + ' — ' + e);
    }

    if ((i2 + 1) % HEALX.PROGRESS_PING_B === 0) {
      _hx__audit('INFO', 'P1B progress: ' + (i2 + 1) + '/' + items.length + ' appointment(s).');
    }
    if (Date.now() >= budget.deadlineMs) { i2++; max = i2; break; }
  }

  SpreadsheetApp.flush(); Utilities.sleep(120);
  state.p1bOffset = max;

  if (state.p1bOffset >= items.length) {
    state.phase = 'P1C'; state.p1cOffset = 0; state.p1cTotal = 0;
    _hx__audit('INFO', 'P1B complete: ' + items.length + ' appointment(s).');
  }
}

function _hx__phase1C_DV_URGENT_(state, budget) {
  _hx__audit('INFO', 'P1C — start (DV URGENT). Time left ~' + (budget.deadlineMs - Date.now()) + 'ms');

  var M = _hx__readMaster_();
  var items = [], canUrgent = 0, withinWin = 0, blockedByStatus = 0;
  for (var i=0;i<M.rows.length;i++){
    var R = M.rows[i];
    if (!R.isDV || !R.apptISO || !R.rootId) continue;
    canUrgent++;
    if (_hx__shouldCreateUrgentToday_(R)) {
      withinWin++; items.push(R);
    } else {
      blockedByStatus++;
    }
  }
  if (!state.p1cTotal) {
    state.p1cTotal = items.length;
    _hx__audit('INFO', 'P1C — urgent candidates=' + canUrgent + ', dueToday=' + items.length +
                      ' (withinWindow=' + withinWin + ', blockedByStatus=' + blockedByStatus + ')');
    if (items.length) _hx__dbg('P1C_SAMPLE', { root: items[0].rootId, apptISO: items[0].apptISO, cust: items[0].customerName });
  }

  var off = state.p1cOffset || 0;
  if (off >= items.length) {
    state.phase = 'P1D';  // was 'P2'
    state.p1dOffset = 2; state.p1dTotal = 0;
    _hx__audit('INFO', 'P1C finished — processed ' + off + '/' + items.length + ' urgent item(s).');
    return;
  }

  var max = Math.min(off + HEALX.BATCH_DV, items.length);
  _hx__dbg('P1C_BATCH', { from: off, to: max, size: (max-off) });

  for (var i3 = off; i3 < max; i3++) {
    var R3 = items[i3];
    try {
      if (typeof DV_upsertUrgentDaily_forToday_ === 'function') {
        DV_upsertUrgentDaily_forToday_({
          rootApptId: R3.rootId,
          customerName: R3.customerName || '',
          nextStepsFromMaster: R3.nextSteps || ''
        });
      } else {
        _hx__audit('WARN', 'P1C: DV helpers missing for root ' + R3.rootId);
      }
    } catch (e) {
      _hx__audit('WARN', 'P1C: urgent upsert failed for root ' + R3.rootId + ' — ' + e);
    }

    if ((i3 + 1) % HEALX.PROGRESS_PING_B === 0) {
      _hx__audit('INFO', 'P1C progress: ' + (i3 + 1) + '/' + items.length + ' urgent item(s).');
    }
    if (Date.now() >= budget.deadlineMs) { i3++; max = i3; break; }
  }

  SpreadsheetApp.flush(); Utilities.sleep(120);
  state.p1cOffset = max;

  if (state.p1cOffset >= items.length) {
    state.phase = 'P1D';  // was 'P2'
    state.p1dOffset = 2; state.p1dTotal = 0;
    _hx__audit('INFO', 'P1C complete: ' + items.length + ' urgent item(s).');
  }

}

function _hx__phase1D_ENRICH_DV_REPS_(state, budget) {
  _hx__audit('INFO', 'P1D — start (Enrich DV reps w/o SO). Time left ~' + (budget.deadlineMs - Date.now()) + 'ms');

  var q = _hx__readQueue_();
  if (!q) {
    _hx__audit('WARN', 'P1D: queue not found; skipping to P2.');
    state.phase = 'P2'; state.p2Offset = 2; state.p2Total = 0;
    return;
  }

  // Build two indexes from Master:
  //  - byRoot: exact RootApptID → best master row
  //  - byName: normalized customer name → best master row (with rep names/emails)
  var M = _hx__readMaster_();
  var byRoot = new Map();
  var byName = new Map();

  function keepBest(existing, candidate) {
    // Prefer a row that already has BOTH rep names; otherwise take the one that has more fields filled.
    function score(r) {
      var s = 0;
      if (r.assignedRepName) s += 2;
      if (r.assistedRepName) s += 2;
      if (r.assignedRepEmail) s += 1;
      if (r.assistedRepEmail) s += 1;
      if (r.apptISO) s += 1;
      return s;
    }
    if (!existing) return candidate;
    return (score(candidate) > score(existing)) ? candidate : existing;
  }

  for (var i = 0; i < M.rows.length; i++) {
    var R = M.rows[i];
    if (R.rootId) {
      byRoot.set(R.rootId, keepBest(byRoot.get(R.rootId), R));
    }
    if (R.customerName) {
      var nk = _hx__normName_(R.customerName);
      if (nk) byName.set(nk, keepBest(byName.get(nk), R));
    }
  }

  // How many rows to consider?
  if (!state.p1dTotal) {
    state.p1dTotal = Math.max(0, q.lastRow - 1);
    _hx__audit('INFO', 'P1D — queue rows to scan: ' + state.p1dTotal);
  }

  var start = state.p1dOffset || 2;
  if (start > q.lastRow) {
    state.phase = 'P2'; state.p2Offset = 2; state.p2Total = 0;
    _hx__audit('INFO', 'P1D finished — nothing to enrich. Moving to P2.');
    return;
  }

  var end = Math.min(start + HEALX.BATCH_ENRICH - 1, q.lastRow);
  _hx__dbg('P1D_CHUNK', { range: start + '-' + end });

  var updated = 0, scanned = 0;

  for (var rowI = start; rowI <= end; rowI++) {
    // Read one row (display vals so we can see what's blank)
    var row = q.sh.getRange(rowI, 1, 1, q.lastCol).getDisplayValues()[0];

    var typ = String(row[q.c.type-1] || '').toUpperCase().trim();
    if (typ.indexOf('DV_') !== 0) continue; // only DV rows

    var soDisp = String(row[q.c.soNumber-1] || '').trim();
    // If an SO exists and reps are already present, skip — P1A handled these.
    if (soDisp) continue;

    // If queue doesn't even have rep columns, there's nothing to set.
    if (!q.c.assignedRep && !q.c.assistedRep && !q.c.assignedEmail && !q.c.assistedEmail) continue;

    var assignedNow = q.c.assignedRep ? String(row[q.c.assignedRep-1] || '').trim() : '';
    var assistedNow = q.c.assistedRep ? String(row[q.c.assistedRep-1] || '').trim() : '';

    // Skip if both rep names are already filled
    if (assignedNow && assistedNow) continue;

    // Try RootApptID from the queue "id" (e.g., DV|<rootId>|...)
    var id  = String(row[q.c.id-1] || '').trim();
    var rootId = _hx__dvRootFromId_(id);

    var R = null;
    if (rootId && byRoot.has(rootId)) {
      R = byRoot.get(rootId);
    } else {
      // fallback: match by customer name
      var cust = String(row[q.c.customerName-1] || '').trim();
      var key  = _hx__normName_(cust);
      if (key && byName.has(key)) R = byName.get(key);
    }

    if (!R) continue; // nothing to enrich with

    // Prepare values to write (only into columns that exist)
    var writePairs = [];
    if (q.c.assignedRep   && R.assignedRepName)  writePairs.push([q.c.assignedRep,   R.assignedRepName]);
    if (q.c.assistedRep   && R.assistedRepName)  writePairs.push([q.c.assistedRep,   R.assistedRepName]);
    if (q.c.assignedEmail && R.assignedRepEmail) writePairs.push([q.c.assignedEmail, R.assignedRepEmail]);
    if (q.c.assistedEmail && R.assistedRepEmail) writePairs.push([q.c.assistedEmail, R.assistedRepEmail]);

    if (writePairs.length) {
      // Write each cell individually to avoid clobbering other fields
      for (var k = 0; k < writePairs.length; k++) {
        q.sh.getRange(rowI, writePairs[k][0]).setValue(writePairs[k][1]);
      }
      updated++;
      _hx__audit('INFO', 'P1D enriched row ' + rowI + ' (DV, no SO): ' +
                        (R.assignedRepName || '(assigned?)') + ' | ' + (R.assistedRepName || '(assisted?)'));
    }

    scanned++;
    if (Date.now() >= budget.deadlineMs) break;
  }

  SpreadsheetApp.flush(); Utilities.sleep(80);
  state.p1dOffset = end + 1;

  if (state.p1dOffset > q.lastRow) {
    state.phase = 'P2'; state.p2Offset = 2; state.p2Total = 0;
    _hx__audit('INFO', 'P1D chunk complete: updated=' + updated + ', scanned=' + scanned + '. Moving to P2.');
  } else {
    _hx__audit('INFO', 'P1D chunk: rows ' + start + '–' + end + ' (updated=' + updated + ', scanned=' + scanned + ').');
  }
}


function _hx__phase2_RECONCILE_(state, budget) {
  _hx__audit('INFO', 'P2 — start (Reconcile). Time left ~' + (budget.deadlineMs - Date.now()) + 'ms');

  var q = _hx__readQueue_(); 
  if (!q) {
    _hx__audit('WARN', 'P2: queue not found; skipping to P3.');
    state.phase = 'P3'; state.p3Offset = 2; state.p3Total = 0;
    return;
  }
  if (!state.p2Total) {
    state.p2Total = Math.max(0, q.lastRow - 1);
    _hx__audit('INFO', 'P2 — queue rows to scan: ' + state.p2Total);
  }

  var M = _hx__readMaster_();
  var bySO = new Map(), byRoot = new Map(), soCount = 0, rootCount = 0;
  for (var i=0;i<M.rows.length;i++){
    var R = M.rows[i];
    if (R.soKey) { bySO.set(R.soKey, R); soCount++; }
    if (R.rootId) { byRoot.set(R.rootId, R); rootCount++; }
  }
  _hx__dbg('P2_INDEX', { soKeys: soCount, rootIds: rootCount });

  var start = state.p2Offset || 2;
  if (start > q.lastRow) {
    state.phase = 'P3'; state.p3Offset = 2; state.p3Total = 0;
    _hx__audit('INFO', 'P2 finished — scanned ' + (q.lastRow - 1) + ' queue row(s).');
    return;
  }

  var end = Math.min(start + HEALX.BATCH_Q - 1, q.lastRow);
  _hx__dbg('P2_CHUNK', { range: start + '-' + end });

  var confirmed = 0, scanned = 0;

  for (var rowI = start; rowI <= end; rowI++) {
    var row = q.sh.getRange(rowI, 1, 1, q.lastCol).getDisplayValues()[0];
    var st = String(row[q.c.status-1] || '').trim().toUpperCase();
    if (st !== ((REMIND && REMIND.ST_PENDING) || 'PENDING') &&
        st !== ((REMIND && REMIND.ST_SNOOZED ) || 'SNOOZED')) {
      continue;
    }

    var typ = String(row[q.c.type-1] || '').trim().toUpperCase();
    if (HEALX.RECONCILE_TYPES.indexOf(typ) === -1) continue;

    scanned++;
    if (scanned % HEALX.PROGRESS_PING_Q === 0) {
      _hx__audit('INFO', 'P2 progress: scanned +' + scanned + ' rows in this chunk…');
    }

    var id  = String(row[q.c.id-1] || '').trim();
    var soR = String(row[q.c.soNumber-1] || '').trim();
    var soK = _hx__soKey(soR);

    if (typ === 'COS') {
      var R = soK ? bySO.get(soK) : null;
      var cos = R ? (R.cos || '') : '';
      if (!R || !_hx__is3DPendingCOS_(cos)) { _hx__confirmRow_(q, rowI, 'Auto-confirm (COS no longer pending)'); confirmed++; }
      continue;
    }
    if (typ === 'FOLLOWUP') {
      var R2 = soK ? bySO.get(soK) : null;
      var stg = R2 ? (R2.salesStage || '') : '';
      if (!R2 || !_hx__eqCI(stg, 'Follow-Up Required')) { _hx__confirmRow_(q, rowI, 'Auto-confirm (Stage not Follow‑Up Required)'); confirmed++; }
      continue;
    }
    if (typ === 'DV_URGENT_OTW_DAILY') {
      var rootId = _hx__dvRootFromId_(id);
      var R3 = rootId ? byRoot.get(rootId) : (soK ? bySO.get(soK) : null);
      if (!R3 || !_hx__shouldCreateUrgentToday_(R3)) { _hx__confirmRow_(q, rowI, 'Auto-confirm (DV urgent not required)'); confirmed++; }
      continue;
    }
  }

  SpreadsheetApp.flush(); Utilities.sleep(100);
  state.p2Offset = end + 1;

  if (state.p2Offset > q.lastRow) {
    state.phase = 'P3'; state.p3Offset = 2; state.p3Total = 0;
    _hx__audit('INFO', 'P2 chunk done: confirmed=' + confirmed + ', scanned=' + scanned + '. Moving to P3.');
  } else {
    _hx__audit('INFO', 'P2 chunk: rows ' + start + '–' + end + ' (confirmed=' + confirmed + ', scanned=' + scanned + ').');
  }
}

function _hx__phase3_NORMALIZE_(state, budget) {
  _hx__audit('INFO', 'P3 — start (Normalize DV dates). Time left ~' + (budget.deadlineMs - Date.now()) + 'ms');

  var q = _hx__readQueue_(); 
  if (!q) {
    state.phase = 'DONE';
    _hx__audit('INFO', 'P3 skipped — queue missing. DONE.');
    return;
  }
  if (!state.p3Total) {
    state.p3Total = Math.max(0, q.lastRow - 1);
    _hx__audit('INFO', 'P3 — queue rows to check: ' + state.p3Total);
  }

  var start = state.p3Offset || 2;
  if (start > q.lastRow) {
    state.phase = 'DONE';
    _hx__audit('INFO', 'P3 finished — nothing left. DONE.');
    return;
  }

  var end = Math.min(start + HEALX.BATCH_NORM - 1, q.lastRow);
  _hx__dbg('P3_CHUNK', { range: start + '-' + end });

  var normalized = 0;

  for (var rowI = start; rowI <= end; rowI++) {
    var row = q.sh.getRange(rowI, 1, 1, q.lastCol).getDisplayValues()[0];
    var typ = String(row[q.c.type-1] || '').trim().toUpperCase();
    if (typ.indexOf('DV_') !== 0) continue;

    var nextDisp  = String(row[q.c.nextDueAt-1]   || '').trim();
    var firstDisp = String(row[q.c.firstDueDate-1]|| '').trim();

    var nextDT = _hx__parseDateTimePT_(nextDisp);
    var firstD = _hx__parseDateOnlyPT_(firstDisp);

    var changed = false;
    if (nextDT){
      var canonNext = Utilities.formatDate(nextDT, HEALX.TZ, 'yyyy-MM-dd HH:mm:ss');
      if (canonNext !== nextDisp) {
        q.sh.getRange(rowI, q.c.nextDueAt).setValue(canonNext);
        changed = true;
      }
    }
    if (firstD){
      var canonFirst = Utilities.formatDate(firstD, HEALX.TZ, 'yyyy-MM-dd');
      if (canonFirst !== firstDisp) {
        q.sh.getRange(rowI, q.c.firstDueDate).setValue(canonFirst);
        changed = true;
      }
    }
    if (changed) normalized++;
    if (Date.now() >= budget.deadlineMs) break;
  }

  SpreadsheetApp.flush(); Utilities.sleep(80);
  state.p3Offset = end + 1;

  if (state.p3Offset > q.lastRow) {
    state.phase = 'DONE';
    _hx__audit('INFO', 'P3 chunk complete: normalized=' + normalized + '. DONE.');
  } else {
    _hx__audit('INFO', 'P3 chunk: rows ' + start + '–' + end + ' (normalized=' + normalized + ').');
  }
}

/* =========================
 * Core IO (Master / Queue)
 * ========================= */

function _hx__readMaster_() {
  var ss = SpreadsheetApp.getActive();

  // 1) Resolve Master tab
  var candidates = [
    HEALX.MASTER_SHEET,
    '00_Master Appointments',
    '00_Master',
    '00 – Master',
    '00 Master'
  ].filter(Boolean);

  var sh = null, tried = [];
  for (var i = 0; i < candidates.length; i++) {
    var name = String(candidates[i]).trim();
    if (!name) continue;
    var s = ss.getSheetByName(name);
    tried.push(name);
    if (s) { sh = s; break; }
  }
  _hx__dbg('MASTER_RESOLVE', { tried: tried, chosen: (sh && sh.getName()) || '(not found)' });

  if (!sh) {
    _hx__audit('ERROR', 'Master sheet not found. Tried: ' + tried.join(' | '));
    throw new Error('Missing Master sheet. Tried: ' + tried.join(' | '));
  }

  // 2) Read header + sizes
  var rg = sh.getDataRange().getDisplayValues();
  var headers = rg[0] || [];
  var lastRow = sh.getLastRow();
  _hx__audit('INFO', 'Master resolved: "' + sh.getName() + '" — rows=' + (lastRow - 1) + ', cols=' + headers.length);

  // 3) Header map (case‑sensitive, by‑name)
  var H = {}; headers.forEach(function(h,i){ var k=String(h||'').trim(); if (k) H[k]=i+1; });
  function pick(names){ for (var i=0;i<names.length;i++){ if (H[names[i]]) return H[names[i]]; } return null; }

  // 4) Column discovery with generous fallbacks
  var cSO     = pick(['SO#','SO','So Number','SO Number']);
  var cCust   = pick(['Customer Name','Customer']);
  var cAssign = pick(['Assigned Rep','Assigned']);
  var cAssist = pick(['Assisted Rep','Assisted']);
  var cAEmail = pick(['Assigned Rep Email']);
  var cSEmail = pick(['Assisted Rep Email']);
  var cStage  = pick(['Sales Stage']);
  var cCOS    = pick(['Custom Order Status']);
  var cCSOS   = pick(['Center Stone Order Status','Center Stone Status','CSOS','Diamond Memo Status','DV Status']);
  var cNext   = pick(['Next Steps','NextSteps','Next steps']);
  var cRoot   = pick(['RootApptID','RootApptId','APPT_ID','ApptID','APPT ID']);
  var cVisit  = pick(['Visit Type','VisitType','Appointment Type','Appt Type']);
  var cDate   = pick([
    'ApptDateTimeISO','ApptDateTime (ISO)','Appointment Date/Time ISO','Appointment DateTime (ISO)',
    'Appt Start ISO','Appt Start (ISO)','Appt Start','Appointment Start','Event Start ISO','EventStartISO',
    'Visit Date','Appt Date','Appointment Date','Date','Event Date','Start Date'
  ]);
  var cTime   = pick(['Visit Time','Appt Time','Appointment Time','Time','Event Time','Start Time']);

  _hx__dbg('MASTER_HEADERS', {
    SO: cSO, Customer: cCust, Assigned: cAssign, Assisted: cAssist,
    AssignedEmail: cAEmail, AssistedEmail: cSEmail,
    SalesStage: cStage, COS: cCOS, CSOS: cCSOS, NextSteps: cNext,
    RootAppt: cRoot, VisitType: cVisit, Date: cDate, Time: cTime,
    first6Headers: headers.slice(0,6)
  });

  if (!cSO)   _hx__audit('WARN','SO column not found — COS/FU may be skipped.');
  if (!cRoot) _hx__audit('WARN','RootApptID column not found — DV IDs may not bind.');

  var rows = [];
  if (lastRow < 2) {
    _hx__audit('WARN', 'Master has no data rows.');
    return { sh: sh, headers: headers, rows: rows };
  }

  // 5) Row read (BULK — one getValues(), compute in memory)
  var nonBlankSO = 0, dvRows = 0, dvWithISO = 0, withRoot = 0;

  var lastCol = sh.getLastColumn();
  var dataVals = (lastRow > 1) ? sh.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  function vAt(rowArr, c){ return c ? rowArr[c-1] : ''; }

  for (var i = 0; i < dataVals.length; i++) {
    var arr = dataVals[i];
    var rIndex = i + 2;

    var soRaw    = _hx__v(vAt(arr, cSO));
    var soKey    = _hx__soKey(soRaw);
    var soPretty = _hx__soPretty(soRaw);
    if (soKey) nonBlankSO++;

    var visitType = _hx__v(vAt(arr, cVisit));
    var isDV      = HEALX.DV_REGEX.test(String(visitType||''));
    if (isDV) dvRows++;

    var dateCell  = vAt(arr, cDate);
    var timeCell  = vAt(arr, cTime);
    var apptISO   = _hx__getApptISO_(dateCell, timeCell);
    if (isDV && apptISO) dvWithISO++;

    var rootId    = _hx__v(vAt(arr, cRoot));
    if (rootId) withRoot++;

    rows.push({
      row: rIndex,
      soRaw: soRaw, soKey: soKey, soPretty: soPretty,
      customerName: _hx__v(vAt(arr, cCust)),
      assignedRepName: _hx__v(vAt(arr, cAssign)),
      assistedRepName: _hx__v(vAt(arr, cAssist)),
      assignedRepEmail: _hx__v(vAt(arr, cAEmail)),
      assistedRepEmail: _hx__v(vAt(arr, cSEmail)),
      salesStage: _hx__v(vAt(arr, cStage)),
      cos: _hx__v(vAt(arr, cCOS)),
      csos: _hx__v(vAt(arr, cCSOS)) || _hx__v(vAt(arr, cCOS)),
      nextSteps: _hx__v(vAt(arr, cNext)),
      rootId: rootId,
      isDV: isDV,
      apptISO: apptISO
    });
  }


  var sampleSO = '';
  for (var i=0;i<rows.length;i++){ if (rows[i].soKey) { sampleSO = rows[i].soPretty || rows[i].soRaw; break; } }

  _hx__audit('INFO', 'Master scan summary — nonBlankSO=' + nonBlankSO +
                     ', dvRows=' + dvRows + ', dvWithISO=' + dvWithISO + ', withRoot=' + withRoot +
                     '; firstSO=' + (sampleSO||'(none)'));

  return { sh: sh, headers: headers, rows: rows };
}

function _hx__readQueue_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(HEALX.QUEUE_SHEET);
  if (!sh) { _hx__audit('WARN','Queue sheet "' + HEALX.QUEUE_SHEET + '" not found.'); return null; }

  var lastRow = sh.getLastRow();
  var lastCol = Math.max(1, sh.getLastColumn());
  var hdr = sh.getRange(1,1,1,lastCol).getDisplayValues()[0] || [];
  var C = {}; hdr.forEach(function(h,i){ var k=String(h||'').trim().toLowerCase(); if (k) C[k]=i+1; });

  function need(name){
    var c = C[String(name||'').trim().toLowerCase()];
    if (!c) throw new Error('04_Reminders_Queue missing header: ' + name);
    return c;
  }

  // helper: first header that exists (all keys are already lowercased in C)
  function any() {
    for (var i = 0; i < arguments.length; i++) {
      var k = String(arguments[i] || '').toLowerCase();
      if (C[k]) return C[k];
    }
    return null;
  }

  var cols = {
    id:            need('id'),
    type:          need('type'),
    status:        need('status'),
    soNumber:      C['sonumber'] || C['so number'] || C['so#'] || C['so'] || need('soNumber'),
    customerName:  C['customername'] || C['customer name'] || C['customer'] || need('customerName'),
    nextDueAt:     need('nextdueat'),
    firstDueDate:  need('firstduedate'),
    confirmedAt:   need('confirmedat'),
    confirmedBy:   need('confirmedby'),

    // ✅ Rep fields (support both camelCase and spaced variants)
    assignedRep:   any('assignedrepname','assigned rep name','assigned rep','assigned'),
    assistedRep:   any('assistedrepname','assisted rep name','assisted rep','assisted'),
    assignedEmail: any('assignedrepemail','assigned rep email','assigned email'),
    assistedEmail: any('assistedrepemail','assisted rep email','assisted email')
  };


  _hx__dbg('QUEUE_HEADERS', { lastRow: lastRow, lastCol: lastCol, resolved: cols, first6: hdr.slice(0,6) });
  return { sh: sh, lastRow: lastRow, lastCol: lastCol, c: cols };
}

/* =========================
 * Confirm & Normalize helpers
 * ========================= */

function _hx__confirmRow_(q, rowIdx, reason) {
  var nowStr = Utilities.formatDate(new Date(), HEALX.TZ, 'yyyy-MM-dd HH:mm:ss');
  q.sh.getRange(rowIdx, q.c.status).setValue((REMIND && REMIND.ST_CONFIRMED) || 'CONFIRMED');
  q.sh.getRange(rowIdx, q.c.confirmedAt).setValue(nowStr);
  q.sh.getRange(rowIdx, q.c.confirmedBy).setValue('system:heal');
  _hx__audit('CONFIRM', 'Row ' + rowIdx + ': ' + reason);
}

function _hx__parseDateTimePT_(s) {
  s = String(s||'').trim(); if (!s) return null;
  var m = s.match(/^(\d{4}-\d{2}-\d{2})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (m) return new Date(m[1] + 'T' + ('0'+m[2]).slice(-2) + ':' + m[3] + ':' + (m[4]||'00'));
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})\s*([AP]M)$/i);
  if (m) {
    var mo=+m[1], d=+m[2], y=+m[3], h=+m[4], mm=+m[5], ap=m[6];
    if (/pm/i.test(ap) && h<12) h+=12; if (/am/i.test(ap) && h===12) h=0;
    var localStr = Utilities.formatDate(new Date(Date.UTC(y, mo-1, d, h, mm, 0)), HEALX.TZ, 'yyyy/MM/dd HH:mm:ss');
    return new Date(localStr);
  }
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})$/);
  if (m) {
    var mo2=+m[1], d2=+m[2], y2=+m[3], h2=+m[4], mm2=+m[5];
    var localStr2 = Utilities.formatDate(new Date(Date.UTC(y2, mo2-1, d2, h2, mm2, 0)), HEALX.TZ, 'yyyy/MM/dd HH:mm:ss');
    return new Date(localStr2);
  }
  return null;
}
function _hx__parseDateOnlyPT_(s) {
  s = String(s||'').trim(); if (!s) return null;
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(m[1]+'-'+m[2]+'-'+m[3]+'T00:00:00');
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    var mo=+m[1], d=+m[2], y=+m[3];
    var localStr = Utilities.formatDate(new Date(Date.UTC(y, mo-1, d, 9, 0, 0)), HEALX.TZ, 'yyyy/MM/dd HH:mm:ss');
    return new Date(localStr);
  }
  return null;
}

/* =========================
 * State & Scheduling
 * ========================= */

function _hx__initOrLoadState_(resetIfIdle) {
  var p = PropertiesService.getScriptProperties();
  var s = _hx__loadState_();
  var now = new Date();

  if (!s || resetIfIdle) {
    s = {
      runId: Utilities.getUuid(),
      startedAt: Utilities.formatDate(now, HEALX.TZ, 'yyyy-MM-dd HH:mm:ss'),
      phase: 'P1A',
      p1aOffset: 0, p1aTotal: 0,
      p1bOffset: 0, p1bTotal: 0,
      p1cOffset: 0, p1cTotal: 0,
      // NEW:
      p1dOffset: 0, p1dTotal: 0,
      // ---
      p2Offset:  2, p2Total:  0,
      p3Offset:  2, p3Total:  0
    };
    _hx__audit('INFO', 'Heal — NEW run started. runId=' + s.runId);
    _hx__saveState_(s);
  }
  p.setProperty(HEALX.RUN_ID_KEY, s.runId);
  return s;
}
function _hx__saveState_(s) {
  var p = PropertiesService.getScriptProperties();
  p.setProperty(HEALX.STATE_KEY, JSON.stringify(s || {}));
  _hx__dbg('STATE_SAVED', s);
}
function _hx__loadState_() {
  var raw = PropertiesService.getScriptProperties().getProperty(HEALX.STATE_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (_){ return null; }
}
function _hx__clearState_() {
  var p = PropertiesService.getScriptProperties();
  p.deleteProperty(HEALX.STATE_KEY);
  p.deleteProperty(HEALX.RUN_FLAG);
  _hx__killExistingResumeTriggers_();
}

function _hx__scheduleResume_(seconds) {
  if (_hx__manualMode_()) {
    _hx__audit('INFO', 'Manual mode: resume NOT scheduled (click "Run next batch").');
    return;
  }
  _hx__killExistingResumeTriggers_();
  var ms = Math.max(2000, (seconds||10)*1000);
  ScriptApp.newTrigger(HEALX.RESUME_FN).timeBased().after(ms).create();
  _hx__audit('INFO', 'Scheduled resume: ' + HEALX.RESUME_FN + ' in ~' + Math.round(ms/1000) + 's');
}

function _hx__killExistingResumeTriggers_() {
  var ts = ScriptApp.getProjectTriggers();
  var killed = 0;
  for (var i=0;i<ts.length;i++){
    var fn = (ts[i].getHandlerFunction && ts[i].getHandlerFunction()) || '';
    if (fn === HEALX.RESUME_FN) { ScriptApp.deleteTrigger(ts[i]); killed++; }
  }
  if (killed) _hx__audit('INFO', 'Removed existing resume triggers: ' + killed);
}

/* =========================
 * Audit / Debug
 * ========================= */

function _hx__debugOn_() {
  try {
    var v = PropertiesService.getScriptProperties().getProperty(HEALX.DEBUG_PROP);
    return (v == null) ? true : String(v).trim() === '1';
  } catch (_){ return true; }
}

/** Always-on audit row (INFO/WARN/ERROR/CONFIRM). */
function _hx__audit(level, note) {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(HEALX.AUDIT_SHEET) || ss.insertSheet(HEALX.AUDIT_SHEET);
    var hdr = ['ts','level','note'];
    var cur = sh.getRange(1,1,1,hdr.length).getDisplayValues()[0] || [];
    var fix = false;
    for (var i=0;i<hdr.length;i++){ if (String(cur[i]||'').trim() !== hdr[i]) { cur[i]=hdr[i]; fix=true; } }
    if (fix) sh.getRange(1,1,1,hdr.length).setValues([cur]);
    sh.appendRow([Utilities.formatDate(new Date(), HEALX.TZ, 'yyyy-MM-dd HH:mm:ss'), level, String(note||'')]);
  } catch (_){}
}

/** Conditional, structured debug (JSON) — controlled by Script Property HEALX_DEBUG ('1' / '0'). */
function _hx__dbg(tag, obj) {
  if (!_hx__debugOn_()) return;
  try {
    _hx__audit('DEBUG', tag + ': ' + JSON.stringify(obj || {}));
  } catch (_){}
}

/* =========================
 * Policy / DV helpers
 * ========================= */

function _hx__is3DPendingCOS_(val) {
  try {
    var s = String(val || '').trim().toLowerCase();
    if (typeof REMIND !== 'undefined' && Array.isArray(REMIND.COS_3D_PENDING)) {
      for (var i=0;i<REMIND.COS_3D_PENDING.length;i++) {
        if (s === String(REMIND.COS_3D_PENDING[i]||'').trim().toLowerCase()) return true;
      }
    }
  } catch (_){}
  return false;
}
function _hx__shouldCreateUrgentToday_(R) {
  if (!R.apptISO) return false;
  var w = _hx__policy().URGENT_WINDOW_DAYS || 7;
  var today = _hx__localAt_(new Date(), 0, 0);
  var apptD = _hx__localAt_(new Date(R.apptISO), 0, 0);
  var days = Math.floor((apptD - today) / (24*60*60*1000));
  if (days < 0 || days > w) return false;

  if (typeof DV_shouldStopUrgentForStatus === 'function' && DV_shouldStopUrgentForStatus(R.csos || R.cos)) return false;
  if (typeof DV_shouldStopDailyForStatus  === 'function' && DV_shouldStopDailyForStatus (R.csos || R.cos)) return false;
  return true;
}
function _hx__remType() {
  return (typeof DV !== 'undefined' && DV.REMTYPE) ? DV.REMTYPE : {
    PROPOSE_NUDGE:    'DV_PROPOSE_NUDGE',
    URGENT_OTW_DAILY: 'DV_URGENT_OTW_DAILY'
  };
}
function _hx__policy() {
  var d = { PROPOSE_NUDGE_OFFSET_DAYS: 12, URGENT_WINDOW_DAYS: 7 };
  try { return (typeof DV !== 'undefined' && DV.POLICY) ? DV.POLICY : d; }
  catch (_){ return d; }
}
function _hx__dvRootFromId_(id) {
  var s = String(id || '');
  // Accept:  "DV|AP-20250921-001|..."  or  "DV/AP-20250921-001|..."
  var m = s.match(/^DV[|\/]([^|\/]+)/i);
  return m ? m[1] : '';
}

/* =========================
 * Utils (normalize / dates / SO)
 * ========================= */
function _hx__normName_(s) {
  s = String(s || '').toLowerCase();
  // collapse spaces, strip punctuation; keep letters/numbers/spaces only
  s = s.replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
  return s;
}

function _hx__v(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return v;
  return String(v == null ? '' : v).replace(/\u00A0/g, ' ').trim();
}
function _hx__eqCI(a,b){ return _hx__v(a).toLowerCase() === _hx__v(b).toLowerCase(); }

function _hx__soKey(raw) {
  var s = String(raw == null ? '' : raw).trim();
  if (!s) return '';
  s = s.replace(/^'+/, '').replace(/^\s*SO#?/i,'').replace(/\s|\u00A0/g,'');
  var digits = s.replace(/\D/g,''); if (!digits) return '';
  return (digits.length < 6) ? digits.padStart(6,'0') : digits.slice(-6);
}
function _hx__soPretty(raw){ var k=_hx__soKey(raw); return k ? (k.slice(0,2)+'.'+k.slice(2)) : ''; }

function _hx__getApptISO_(dateCell, timeCell) {
  if (dateCell) {
    if (Object.prototype.toString.call(dateCell) === '[object Date]' && !isNaN(dateCell)) {
      var y=dateCell.getFullYear(), m=dateCell.getMonth(), d=dateCell.getDate();
      var hh=9, mm=0;
      if (timeCell && Object.prototype.toString.call(timeCell) === '[object Date]' && !isNaN(timeCell)) {
        hh = Number(Utilities.formatDate(timeCell, HEALX.TZ, 'H'));
        mm = Number(Utilities.formatDate(timeCell, HEALX.TZ, 'm'));
      } else if (timeCell) {
        var mt = String(timeCell||'').trim().match(/^(\d{1,2})(?::(\d{2}))?\s*([ap]m)?$/i);
        if (mt) {
          hh = +mt[1]; mm = mt[2]?+mt[2]:0;
          var ap = (mt[3]||'').toLowerCase();
          if (ap === 'pm' && hh < 12) hh += 12;
          if (ap === 'am' && hh === 12) hh = 0;
        }
      }
      var localStr = Utilities.formatDate(new Date(Date.UTC(y, m, d, hh, mm, 0)), HEALX.TZ, 'yyyy/MM/dd HH:mm:ss');
      return new Date(localStr).toISOString();
    }
    var s = String(dateCell||'').trim(); var d2 = s ? new Date(s) : null;
    if (d2 && !isNaN(d2)) return d2.toISOString();
  }
  return '';
}
function _hx__dayKeyFromISO_(iso) {
  var d = new Date(iso);
  var Y = Utilities.formatDate(d, HEALX.TZ, 'yyyy');
  var M = Utilities.formatDate(d, HEALX.TZ, 'MM');
  var D = Utilities.formatDate(d, HEALX.TZ, 'dd');
  return Y + M + D;
}
function _hx__addDays_(d, n){ var t=new Date(d); t.setDate(t.getDate()+n); return t; }
function _hx__localAt_(d, hh, mm){
  var Y = Number(Utilities.formatDate(d, HEALX.TZ, 'yyyy'));
  var M = Number(Utilities.formatDate(d, HEALX.TZ, 'MM'));
  var D = Number(Utilities.formatDate(d, HEALX.TZ, 'dd'));
  return new Date(Y, M-1, D, hh||0, mm||0, 0, 0);
}

/** ========= Manual UI (no triggers) ========= */
//function onOpen() {
//  try {
//    var ui = SpreadsheetApp.getUi();
//    ui.createMenu('⛭ Heal Reminders (Manual)')
//      .addItem('Start / Reset (manual)', 'HealManual_startReset')
//      .addItem('Run next batch',        'HealManual_nextBatch')
//      .addSeparator()
//      .addSubMenu(
//        ui.createMenu('Jump to phase')
//          .addItem('→ P1A (COS/FU)',     'HealManual_jumpP1A')
//          .addItem('→ P1B (DV propose)', 'HealManual_jumpP1B')
//          .addItem('→ P1C (DV urgent)',  'HealManual_jumpP1C')
//          .addItem('→ P2 (Reconcile)',   'HealManual_jumpP2')
//          .addItem('→ P3 (Normalize)',   'HealManual_jumpP3')
//      )
//      .addSeparator()
//      .addItem('Status (log state)',    'RemindersHeal_status')
//      .addItem('Disable manual mode',   'HealManual_disable')
//      .addToUi();
//  } catch (_) {}
//}

/** Enable manual mode + reset state, then log and stop (no scheduling). */
function HealManual_startReset() {
  var p = PropertiesService.getScriptProperties();
  p.setProperty('HEALX_MANUAL','1');           // enable manual mode
  p.deleteProperty(HEALX.STATE_KEY);           // clear state
  _hx__audit('INFO', 'Manual mode ENABLED; state reset. Click "Run next batch" to begin (phase=P1A).');
  RemindersHeal_status();
}

/** One click = process the next tiny batch of the current phase. */
function HealManual_nextBatch() {
  PropertiesService.getScriptProperties().setProperty('HEALX_MANUAL','1');
  RemindersHeal__runStep_({ resetIfIdle: false });
}

/** Disable manual mode (go back to auto, if you later want to resume scheduling) */
function HealManual_disable() {
  PropertiesService.getScriptProperties().deleteProperty('HEALX_MANUAL');
  _hx__audit('INFO', 'Manual mode DISABLED. (Auto scheduling will resume on next start.)');
  RemindersHeal_status();
}

/** --------- Optional jumpers (sets phase then runs a batch) --------- */
function _jumpTo_(phase) {
  PropertiesService.getScriptProperties().setProperty('HEALX_MANUAL','1');
  var s = _hx__loadState_() || {};
  if (!s.runId) { s = _hx__initOrLoadState_(true); } // new run if none
  s.phase = phase;
  _hx__saveState_(s);
  _hx__audit('INFO', 'Jumped to phase: ' + phase + ' — click "Run next batch".');
}

function HealManual_jumpP1A(){ _jumpTo_('P1A'); }
function HealManual_jumpP1B(){ _jumpTo_('P1B'); }
function HealManual_jumpP1C(){ _jumpTo_('P1C'); }
function HealManual_jumpP2(){  _jumpTo_('P2');  }
function HealManual_jumpP3(){  _jumpTo_('P3');  }
