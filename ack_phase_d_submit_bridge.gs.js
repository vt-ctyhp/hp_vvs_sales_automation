// @bundle: Ack pipes + dashboard + schedule + snapshot
/***** Phase 2 — Unified Submit (ACK + Reminders)
 *  - Calls your original ACK submit (unchanged behavior)
 *  - Then processes "Reminder Action" rows in Q_<Rep> using Reminder ID
 *  - Updates 04_Reminders_Queue and appends audit rows to 15_Reminders_Log
 ****************************************************************/

// Public entry point — use this instead of the original Submit in your menu
function submitMyQueueUnified() {
  var rep = (typeof detectRepName_ === 'function') ? detectRepName_() : '';
  if (!rep) { SpreadsheetApp.getUi().alert('Could not detect your Rep name.'); return; }

  var ss = SpreadsheetApp.getActive();
  var sh = (typeof ensureQueueSheet_ === 'function') ? ensureQueueSheet_(rep) : ss.getSheetByName('Q_' + rep);
  if (!sh) { SpreadsheetApp.getUi().alert('Queue sheet not found for ' + rep); return; }

  // 0) PRE-COLLECT the user's Reminder actions (strict) BEFORE the legacy submit clears anything
  var preWork = _collectReminderActionsFromSheet_(sh);

  // 0b) SCRUB header/subsection hint text in the Action column so legacy submit doesn't count banners
  _scrubReminderSectionHintCells_(sh);

  // 1) Run existing ACK submit (unchanged behavior; it will now ignore banners; see Edit 3 for "skip Reminder rows")
  if (typeof submitMyQueue === 'function') {
    try { submitMyQueue(); } catch (e) { Logger.log('submitMyQueue() failed: ' + e); }
  }

  // 2) Apply Reminder actions captured in preWork → update 04_ + 15_ AND also log to 06_
  var out = _processReminderActionsForRep_(rep, preWork, { alsoLogTo06: true });
  recomputeAckStatusSummary(); // reflect reminder-driven logs in 00_ immediately

  // 3) Auto-refresh the owner’s queue (ACK + Reminders) so both lists remain visible
  try {
    if (typeof refreshMyQueueHybrid === 'function') {
      refreshMyQueueHybrid();                     // ← hybrid: ACKs + Reminders (date-based due)
    } else if (typeof refreshMyQueue === 'function') {
      refreshMyQueue();                           // ← fallback: ACK-only
    }
  } catch (e) {
    Logger.log('Hybrid refresh after submit failed: ' + e);
  }

  var msg = 'Submit complete.\n' +
            'Reminders processed: ' + out.processed + '\n' +
            (out.errors.length ? ('Errors:\n- ' + out.errors.join('\n- ')) : 'No reminder errors.');
  SpreadsheetApp.getUi().alert(msg);
}


// ===== Core processing =====

function _processReminderActionsForRep_(rep, preWork, opts) {
  opts = opts || {};
  var alsoLogTo06 = !!opts.alsoLogTo06;

  var ss = SpreadsheetApp.getActive();
  var sh = (typeof ensureQueueSheet_ === 'function') ? ensureQueueSheet_(rep) : ss.getSheetByName('Q_' + rep);
  try { _ensureReminderBridgeColumnsOnQueue_(sh); } catch (_) {}

  if (!sh) return { processed: 0, errors: ['Queue sheet not found for ' + rep] };

  var lc = sh.getLastColumn(), lr = sh.getLastRow();
  if (lr < 2) return { processed: 0, errors: [] };

  // Map Q_ headers (needed if we must collect work ourselves)
  var qHdr = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(function(s){ return String(s||'').trim(); });
  function Qi(name, alts){
    var want = _normHeader_(name);
    for (var i = 0; i < qHdr.length; i++) {
      if (_normHeader_(qHdr[i]) === want) return i + 1;
    }
    if (alts && alts.length) {
      for (var j = 0; j < alts.length; j++) {
        var w = _normHeader_(alts[j]);
        for (var k = 0; k < qHdr.length; k++) {
          if (_normHeader_(qHdr[k]) === w) return k + 1;
        }
      }
    }
    return -1;
  }

  // NEW: prefer the pre‑collected actions; the legacy submit rebuilds Q_ and drops bridge columns.
  var work = Array.isArray(preWork) ? preWork : _collectReminderActionsFromSheet_(sh);

  // If we already have work, do not depend on Q_ headers anymore.
  if (!work.length) {
    // No preWork → ensure the bridge columns exist, then try to collect once more.
    try { _ensureReminderBridgeColumnsOnQueue_(sh); } catch (_) {}

    // Re-read headers after auto‑heal and check for the ID column only now.
    lc   = sh.getLastColumn();
    qHdr = sh.getRange(1, 1, 1, lc).getDisplayValues()[0].map(function (s) { return String(s || '').trim(); });
    var cId = Qi('Reminder ID');
    if (cId === -1) {
      return { processed: 0, errors: ['Reminder ID column missing on Q_ sheet. Please Refresh My Queue and try again.'] };
    }

    // Try again now that columns are present.
    work = _collectReminderActionsFromSheet_(sh);
    if (!work.length) return { processed: 0, errors: [] };
  }

  // Validate inputs (Notes and Snooze until)
  var errors = [];
  function needsNote(a){ return a === 'Cancel' || a.indexOf('Snooze') === 0; }
  function isCustomSnooze(a){ return a === 'Snooze…'; }

  work.forEach(function(w){
    if (needsNote(w.action) && !String(w.note||'').trim()) {
      errors.push('Row ' + w.idx + ': Notes required for "' + w.action + '".');
    }
    if (isCustomSnooze(w.action)) {
      var dt = _parseDateStrict_(w.snoozeCell);
      if (!dt) errors.push('Row ' + w.idx + ': "Snooze Until" date/time required for Snooze….');
    }
  });
  if (errors.length) return { processed: 0, errors: errors };

  // Build id → 04_ row map
  var sQ = ss.getSheetByName('04_Reminders_Queue');
  if (!sQ) return { processed: 0, errors: ['04_Reminders_Queue not found.'] };
  var qLc = sQ.getLastColumn(), qLr = sQ.getLastRow();
  if (qLr < 2) return { processed: 0, errors: ['04_Reminders_Queue is empty.'] };

  var qHdrs = sQ.getRange(1,1,1,qLc).getDisplayValues()[0].map(function(s){ return String(s||'').trim(); });
  function QiQ(name){ var c = qHdrs.indexOf(name); return c<0 ? -1 : (c+1); }

  var cIdQ    = QiQ('id');
  var cTypeQ  = QiQ('type');
  var cSOQ    = QiQ('soNumber');
  var cStatQ  = QiQ('status');
  var cNextQ  = QiQ('nextDueAt');
  var cSnozQ  = QiQ('snoozeUntil');
  var cCfmAtQ = QiQ('confirmedAt');
  var cCfmByQ = QiQ('confirmedBy');
  var cCanQ   = QiQ('cancelReason');
  var cAdmAct = QiQ('lastAdminAction');
  var cAdmBy  = QiQ('lastAdminBy');

  if ([cIdQ,cStatQ].some(function(x){return x===-1;})) {
    return { processed: 0, errors: ['04_Reminders_Queue headers incomplete.'] };
  }

  var qVals = sQ.getRange(2,1,qLr-1,qLc).getValues();
  var rowById = new Map(); // id -> {rowIdx, arr}
  for (var i=0;i<qVals.length;i++){
    var id = String(qVals[i][cIdQ-1] || '').trim();
    if (id) rowById.set(id, {row: i+2, arr: qVals[i]});
  }

  var actor = _safeEmail_() || Session.getActiveUser().getEmail() || 'unknown';
  var tz = Session.getScriptTimeZone() || 'America/Los_Angeles';
  var now = new Date();

  // === Confirm-gate: verify 00_ state matches type-specific rules before allowing Confirm ===
  // Build quick 00_ lookup by RootApptID
  var s00 = ss.getSheetByName('00_Master Appointments');
  var byRoot00 = new Map();
  if (s00) {
    try {
      var objs00 = getObjects_(s00);
      objs00.forEach(function(r){
        var root = String(r['RootApptID'] || '').trim();
        if (root) byRoot00.set(root, r);
      });
    } catch(_) {}
  }

  // Center Stone Order Status gating for DV_PROPOSE
  // === CSOS allow-lists for DV confirms ===
  var CSOS_OK_FOR_DV_PROPOSE = new Set([
    'diamond memo – some on the way',
    'diamond memo – on the way',
    'diamond memo – some delivered',
    'diamond memo – delivered',
    'diamond viewing ready',
    'diamond deposit, confirmed order'
  ]);

  var CSOS_OK_FOR_DV_URGENT = new Set([
    'diamond memo – delivered',
    'diamond viewing ready',
    'diamond deposit, confirmed order'
  ]);

  // Normalize + membership test (default DENY)
  function _csosAllows_(raw, okSet) {
    var s = String(raw || '').trim().toLowerCase();
    if (!s) return false;
    // require positive inclusion in the chosen allow-list
    if (okSet && okSet.size) {
      // exact case-insensitive match
      if (okSet.has(s)) return true;
    }
    return false;
  }

  // Return true if underlying 00_ state satisfies Confirm for this type
  function _okConfirmForType_(typeLabel, root) {
    var t = _normType_(typeLabel);
    var o = byRoot00.get(String(root||'').trim());
    if (!o) return false; // be conservative if we can’t verify

    if (t === 'FOLLOW_UP') {
      var ssVal = String(o['Sales Stage'] || '').trim().toLowerCase();
      return ssVal && ssVal !== 'follow-up required';
    }
    if (t === 'DV_PROPOSE' || t === 'DV_PROPOSE_NUDGE') {
      var csosRawP = o['Center Stone Order Status'];
      return _csosAllows_(csosRawP, CSOS_OK_FOR_DV_PROPOSE);
    }

    if (t === 'DV_URGENT') {
      var csosRawU = o['Center Stone Order Status'];
      return _csosAllows_(csosRawU, CSOS_OK_FOR_DV_URGENT);
    }
    if (t === 'COS') {
      var cos = String(o['Custom Order Status'] || '').trim();
      var target = (typeof IN_PRODUCTION_LITERAL !== 'undefined' && IN_PRODUCTION_LITERAL) ? IN_PRODUCTION_LITERAL : 'In Production';
      return String(cos||'').trim().toLowerCase() === String(target).trim().toLowerCase();
    }
    return true; // other types: allow

  }

  // PASS: block any invalid Confirms BEFORE we write anything
  work.forEach(function(w){
    if (String(w.action || '').toLowerCase() !== 'confirm') return;
    var rec = rowById.get(w.id);
    if (!rec) { errors.push('Row ' + w.idx + ': Reminder ID not found in 04_.'); return; }
    var typeQ = (cTypeQ !== -1) ? String(rec.arr[cTypeQ-1] || '') : '';
    if (!_okConfirmForType_(typeQ, w.root)) {
      var tNorm = (typeof _normType_ === 'function') ? _normType_(typeQ) : String(typeQ || '').toUpperCase();
      var msg;

      if (tNorm === 'DV_PROPOSE') {
        msg = 'Cannot Confirm DV_PROPOSE — CSOS must be one of: '
            + 'Diamond Memo – SOME On the Way; Diamond Memo – On the Way; '
            + 'Diamond Memo – SOME Delivered; Diamond Memo – Delivered; '
            + 'Diamond Viewing Ready; Diamond Deposit, Confirmed Order.';
      } else if (tNorm === 'DV_URGENT') {
        msg = 'Cannot Confirm DV_URGENT — CSOS must be one of: '
            + 'Diamond Memo – Delivered; Diamond Viewing Ready; '
            + 'Diamond Deposit, Confirmed Order.';
      } else if (tNorm === 'FOLLOW_UP') {
        msg = 'Cannot Confirm Follow-Up — Sales Stage must move off “Follow-Up Required”.';
      } else if (tNorm === 'COS') {
        msg = 'Cannot Confirm COS — Custom Order Status must be “In Production”.';
      } else {
        msg = 'Cannot Confirm ' + (typeQ || 'Reminder') + ' — underlying 00_ status not satisfied. '
            + 'Use Snooze… with a reason.';
      }

      errors.push('Row ' + w.idx + ': ' + msg);
    }

  });
  if (errors.length) return { processed: 0, errors: errors };


  // Ensure 15_ log exists
  var sLog = ss.getSheetByName('15_Reminders_Log') || ss.insertSheet('15_Reminders_Log');
  if (sLog.getLastRow() < 1) {
    sLog.getRange(1,1,1,9).setValues([['ts','id','soNumber','type','action','by','note','snoozeUntil','nextDueAt']]);
    sLog.setFrozenRows(1);
  }

  var processed = 0;

  work.forEach(function(w){
    var rec = rowById.get(w.id);
    if (!rec) { errors.push('Row ' + w.idx + ': Reminder ID not found in 04_.'); return; }

    var newStatus = null, snoozeUntil = null, nextDue = null, cancelReason = '', confirmedAt = '', confirmedBy = '', lastAdminAction = '', lastAdminBy = actor;

    var aL = String(w.action||'').toLowerCase();
    if (aL === 'confirm') {
      newStatus = 'CONFIRMED';
      confirmedAt = now;
      confirmedBy = actor;
      lastAdminAction = 'CONFIRM';
    } else if (aL.indexOf('snooze') === 0) {
      newStatus = 'SNOOZED';
      if (aL === 'snooze 1 day') {
        snoozeUntil = _tomorrow0930_(tz);
      } else {
        snoozeUntil = _parseDateStrict_(w.snoozeCell); // validated earlier
      }
      nextDue = snoozeUntil;
      lastAdminAction = 'SNOOZE';
      cancelReason = '';
    } else if (aL === 'cancel') {
      newStatus = 'CANCELLED';
      cancelReason = w.note || 'Cancelled';
      lastAdminAction = 'CANCEL';
    } else {
      return; // unknown label
    }

    // Apply updates to 04_
    var rr = rec.row;
    if (newStatus != null) sQ.getRange(rr, cStatQ).setValue(newStatus);
    if (cSnozQ !== -1)     sQ.getRange(rr, cSnozQ).setValue(snoozeUntil);
    if (cNextQ !== -1)     sQ.getRange(rr, cNextQ).setValue(nextDue);
    if (cCfmAtQ !== -1)    sQ.getRange(rr, cCfmAtQ).setValue(confirmedAt);
    if (cCfmByQ !== -1)    sQ.getRange(rr, cCfmByQ).setValue(confirmedBy);
    if (cCanQ !== -1)      sQ.getRange(rr, cCanQ).setValue(cancelReason);
    if (cAdmAct !== -1)    sQ.getRange(rr, cAdmAct).setValue(lastAdminAction);
    if (cAdmBy  !== -1)    sQ.getRange(rr, cAdmBy).setValue(lastAdminBy);

    // Append audit in 15_
    var so = (cSOQ !== -1)   ? rec.arr[cSOQ-1]   : '';
    var tp = (cTypeQ !== -1) ? rec.arr[cTypeQ-1] : '';
    sLog.appendRow([
      Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'),
      w.id, so, tp, newStatus, actor, w.note || '',
      snoozeUntil || '', nextDue || ''
    ]);

    processed++;
  });

  // ALSO append these Reminder actions as Ack logs in 06_ (so both flows appear in 06)
  if (alsoLogTo06 && processed > 0) {
    _appendAcknowledgementLogsForReminders_(work, rep, tz);
  }

  return { processed: processed, errors: errors };
}


function _collectReminderActionsFromSheet_(sh) {
  var lc = sh.getLastColumn(), lr = sh.getLastRow();
  if (lr < 2) return [];

  var hdr = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(function(s){ return String(s||'').trim(); });
  function Qi(name, alts){
    var want = _normHeader_(name);
    for (var i = 0; i < hdr.length; i++) {
      if (_normHeader_(hdr[i]) === want) return i + 1;
    }
    if (alts && alts.length) {
      for (var j = 0; j < alts.length; j++) {
        var w = _normHeader_(alts[j]);
        for (var k = 0; k < hdr.length; k++) {
          if (_normHeader_(hdr[k]) === w) return k + 1;
        }
      }
    }
    return -1;
  }


  var cAction = Qi('Ack Status');            // reused for Reminder Action
  var cNote   = Qi('Ack Note');
  var cSnooz  = Qi('Reminder Snooze Until');
  var cId     = Qi('Reminder ID');
  var cRoot   = Qi('RootApptID');
  var cCust   = Qi('Customer Name');
  var cCSR    = Qi('Client Status Report URL');
  var cCOS    = Qi('Custom Order Status');
  var cUpdAt  = Qi('Updated At');

  if (cId === -1) return [];

  var vals = sh.getRange(2,1,lr-1,lc).getValues();
  var out = [];

  function normAction(a){
    var s = String(a||'').trim();
    var L = s.toLowerCase();
    if (L === 'confirm') return 'Confirm';
    if (L === 'cancel')  return 'Cancel';
    if (L === 'snooze 1 day') return 'Snooze 1 Day';
    if (L === 'snooze…' || L === 'snooze ...' || L === 'snooze ...') return 'Snooze…';
    return ''; // anything else (including "Action: Confirm / ...") is not a real action
  }

  for (var i=0;i<vals.length;i++){
    var row = vals[i];
    var id   = String(row[cId-1] || '').trim();
    if (!id) continue; // not a reminder row

    var actionRaw = row[cAction-1];
    var action = normAction(actionRaw);
    if (!action) continue; // user did not pick a real action

    // Skip section header rows like "— Diamond Viewing —" (col A starts with '— ') just in case
    var colA = String(row[0] || '');
    if (colA.indexOf('— ') === 0) continue;

    out.push({
      idx: 2 + i,
      id: id,
      action: action,
      note: String(row[cNote-1] || '').trim(),
      snoozeCell: row[cSnooz-1] || '',
      root: cRoot>0 ? String(row[cRoot-1]||'') : '',
      customer: cCust>0 ? String(row[cCust-1]||'') : '',
      csrUrl: cCSR>0 ? String(row[cCSR-1]||'') : '',
      cos: cCOS>0 ? String(row[cCOS-1]||'') : '',
      updatedAt: cUpdAt>0 ? row[cUpdAt-1] : ''
    });
  }
  return out;
}

function _scrubReminderSectionHintCells_(sh) {
  try {
    var lc = sh.getLastColumn(), lr = sh.getLastRow();
    if (lr < 2) return;
    var hdr = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(function(s){ return String(s||'').trim(); });
    function H(name){ 
      var want = _normHeader_(name);
      for (var i=0;i<hdr.length;i++){ if (_normHeader_(hdr[i]) === want) return i+1; }
      return -1;
    }
    var cAction = H('Ack Status');
    var cId     = H('Reminder ID');

    if (cAction <= 0 || cId <= 0) return;

    var rng = sh.getRange(2,1,lr-1,Math.max(cAction,cId));
    var vals = rng.getValues();
    var dirty = [];
    for (var i=0;i<vals.length;i++){
      var c1 = String(vals[i][0] || '');
      var id = String(vals[i][cId-1] || '').trim();
      var a  = String(vals[i][cAction-1] || '');
      var isSection = c1.indexOf('— ') === 0 && !id;
      var isHint = /Action:|Reminder Action/i.test(a);
      if (isSection && isHint) {
        vals[i][cAction-1] = ''; // scrub hint so legacy submit won't count it
        dirty.push(i);
      }
    }
    if (dirty.length) rng.setValues(vals);
  } catch(_) {}
}

function _appendAcknowledgementLogsForReminders_(work, rep, tz) {
  if (!work || !work.length) return;
  var ss = SpreadsheetApp.getActive();
  var s06 = ss.getSheetByName('06_Acknowledgement_Log') || ss.insertSheet('06_Acknowledgement_Log');
  var hdr = s06.getLastRow() ? s06.getRange(1,1,1,s06.getLastColumn()).getDisplayValues()[0]
                             : (function(){ s06.getRange(1,1,1,10).setValues([['Timestamp','Log Date','RootApptID','Rep','Ack Status','Ack Note','Customer (at log)','Custom Order Status (at log)','Last Updated At (at log)','Client Status Report URL']]); s06.setFrozenRows(1); return s06.getRange(1,1,1,s06.getLastColumn()).getDisplayValues()[0]; })();

  function hIdx(name, alts){
    var i = hdr.indexOf(name);
    if (i >= 0) return i;
    if (alts && alts.length) {
      for (var j=0;j<alts.length;j++){ var k = hdr.indexOf(alts[j]); if (k>=0) return k; }
    }
    return -1;
  }

  var iTS   = hIdx('Timestamp');
  var iLD   = hIdx('Log Date');
  var iRoot = hIdx('RootApptID', ['Root Appt ID','Root Appointment ID']);
  var iRep  = hIdx('Rep');
  var iStat = hIdx('Ack Status');
  var iNote = hIdx('Ack Note');
  var iCust = hIdx('Customer (at log)', ['Customer','Customer Name (at log)']);
  var iCOS  = hIdx('Custom Order Status (at log)', ['COS (at log)']);
  var iUpd  = hIdx('Last Updated At (at log)', ['Updated At (at log)']);
  var iURL  = hIdx('Client Status Report URL', ['CSR URL']);

  var now = new Date();
  var dateStr = Utilities.formatDate(now, tz || Session.getScriptTimeZone() || 'America/Los_Angeles', 'yyyy-MM-dd HH:mm:ss');
  var logDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  // Prepare rows in one shot
  var rows = work.map(function(w){
    var arr = new Array(hdr.length).fill('');
    if (iTS>=0)   arr[iTS] = dateStr;
    if (iLD>=0)   arr[iLD] = logDate;
    if (iRoot>=0) arr[iRoot] = w.root || '';
    if (iRep>=0)  arr[iRep] = rep || '';
    if (iStat>=0) {
      // Map Reminder action → canonical Ack status used by dashboards
      var L = String(w.action||'').toLowerCase();
      arr[iStat] = (L === 'confirm') ? (typeof LABELS !== 'undefined' ? LABELS.FULLY_UPDATED : 'Fully Updated')
                                     : (typeof LABELS !== 'undefined' ? LABELS.NEEDS_FOLLOW_UP : 'Needs follow-up');
    }
    if (iNote>=0) arr[iNote] = w.note || '';
    if (iCust>=0) arr[iCust] = w.customer || '';
    if (iCOS>=0)  arr[iCOS]  = w.cos || '';
    if (iUpd>=0)  arr[iUpd]  = w.updatedAt || '';
    if (iURL>=0)  arr[iURL]  = w.csrUrl || '';
    return arr;
  });

  if (rows.length) {
    var startRow = s06.getLastRow() + 1;
    s06.getRange(startRow, 1, rows.length, hdr.length).setValues(rows);
  }
}


// ===== utilities (Phase 2) =====

function _parseDateStrict_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return v;
  var d = new Date(String(v));
  return isNaN(d) ? null : d;
}
function _tomorrow0930_(tz) {
  var tzStr = tz || Session.getScriptTimeZone() || 'America/Los_Angeles';
  var now = new Date();
  // Base: tomorrow (always)
  var base = Utilities.formatDate(new Date(now.getTime() + 24*60*60*1000), tzStr, 'yyyy-MM-dd');
  var d = new Date(base + 'T09:30:00');
  d.setHours(9, 30, 0, 0);
  return d;
}


