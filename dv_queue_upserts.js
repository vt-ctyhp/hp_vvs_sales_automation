/** File: 12.1 - dv_queue_upserts.gs (v1.1, queue schema aligned)
 * Purpose: Safe upserts into 04_Reminders_Queue using EXISTING headers.
 * Uses: id, type, firstDueDate, nextDueAt, status, customerName, nextSteps, createdAt, createdBy
 * Policy: Calendar days; "earlier nextDueAt wins" when same id appears again.
 */

/** ── SO helpers shim (DV) — unify with Reminders_v1 ─────────────────────────
 * Canonical: _soKey_(raw) → '001293'  |  _soPretty_(raw) → '00.1293'
 * Back-compat: _canonSO_(raw) → _soKey_(raw)
 * If Reminders_v1 already defined them, we reuse; else we provide identical fallbacks.
 */


if (typeof _soKey_ !== 'function') {
  function _soKey_(raw) {
    let s = String(raw == null ? '' : raw).trim();
    if (!s) return '';
    s = s.replace(/^'+/, '');          // leading apostrophe from Sheets
    s = s.replace(/^\s*SO#?/i, '');    // optional "SO" label
    s = s.replace(/\s|\u00A0/g, '');   // spaces & NBSP
    const digits = s.replace(/\D/g, ''); // digits only
    if (!digits) return '';
    const d = digits.length < 6 ? digits.padStart(6,'0') : digits.slice(-6);
    return d; // '001293'
  }
}

if (typeof _soPretty_ !== 'function') {
  function _soPretty_(raw) {
    const k = _soKey_(raw);
    return k ? (k.slice(0,2) + '.' + k.slice(2)) : '';
  }
}

// Back-compat alias: some older code calls _canonSO_
if (typeof _canonSO_ !== 'function') {
  function _canonSO_(raw) { return _soKey_(raw); }
}


// ----- Sheet access + headers (no column additions here) -----

function DVQ_getQueueSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('04_Reminders_Queue');
  if (!sh) throw new Error('Missing sheet "04_Reminders_Queue".');
  return sh;
}

function DVQ_headerMap_(sh) {
  var lastCol = Math.max(1, sh.getLastColumn());
  var hdr = sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(function(h){ return String(h||'').trim(); });
  var H = {}; hdr.forEach(function(h,i){ if(h) H[h]=i+1; });
  return H;
}

/** Add a header once (append to far right) and return fresh header map */
function DVQ_getOrAddHeader_(sh, headerName){
  var H = DVQ_headerMap_(sh);
  if (!H[headerName]) {
    sh.getRange(1, sh.getLastColumn()+1).setValue(headerName);
    SpreadsheetApp.flush();
    H = DVQ_headerMap_(sh);
  }
  return H;
}

// Required queue columns we will use
function DVQ_requireHeaders_(H) {
  var need = ['id','type','firstDueDate','nextDueAt','status','customerName','nextSteps','createdAt','createdBy'];
  var missing = need.filter(function(n){ return !H[n]; });
  if (missing.length) {
    throw new Error('04_Reminders_Queue is missing required header(s): ' + missing.join(', ') + '. Please add them once, to the header row.');
  }
}

// ----- Time helpers (calendar-day policy) -----

function DVQ_fmtPT_(d) {
  var tz = Session.getScriptTimeZone?.() || 'America/Los_Angeles';
  return Utilities.formatDate(new Date(d), tz, 'yyyy-MM-dd HH:mm:ss');
}
function DVQ_fmtDateOnlyPT_(d) {
  var tz = Session.getScriptTimeZone?.() || 'America/Los_Angeles';
  return Utilities.formatDate(new Date(d), tz, 'yyyy-MM-dd');
}
function DVQ_toDate_(x) {
  if (!x) return null;
  if (Object.prototype.toString.call(x) === '[object Date]' && !isNaN(x)) return x;
  var s = String(x||'').trim();

  // YYYY-MM-DD HH:mm[:ss]
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (m) return new Date(m[1]+'-'+m[2]+'-'+m[3]+'T'+('0'+m[4]).slice(-2)+':'+m[5]+':'+(m[6]||'00'));

  // M/D/YYYY [h:mm] [AM|PM]
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})\s*([AP]M)?)?$/i);
  if (m) {
    var mo = +m[1], da = +m[2], y = +m[3];
    var hh = m[4] ? +m[4] : 9, mm = m[5] ? +m[5] : 0;
    var ap = m[6] || '';
    if (/pm/i.test(ap) && hh < 12) hh += 12;
    if (/am/i.test(ap) && hh === 12) hh = 0;
    var tz = Session.getScriptTimeZone?.() || 'America/Los_Angeles';
    var localStr = Utilities.formatDate(new Date(Date.UTC(y, mo-1, da, hh, mm, 0)), tz, 'yyyy/MM/dd HH:mm:ss');
    return new Date(localStr);
  }
  return null;
}


function DVQ_iso_(d){ return new Date(d).toISOString(); }
function DVQ_addDays_(d, days){ var dd=new Date(d); dd.setDate(dd.getDate()+days); return dd; }
function DVQ_setLocalTime_(d, hh, mm) {
  var tz = Session.getScriptTimeZone?.() || 'America/Los_Angeles';
  var y = Utilities.formatDate(d, tz, 'yyyy');
  var m = Utilities.formatDate(d, tz, 'MM');
  var da= Utilities.formatDate(d, tz, 'dd');
  return new Date(Number(y), Number(m)-1, Number(da), hh||9, mm||0, 0, 0);
}
function DVQ_todayKey_(){
  var tz = Session.getScriptTimeZone?.() || 'America/Los_Angeles';
  return Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
}

// ----- Core upsert (earlier "nextDueAt" wins for same id) -----

/**
 * Upsert a reminder row into 04_Reminders_Queue using existing columns.
 * - "id" is the dedupe key. If found, we keep the EARLIER of (existing.nextDueAt, newDue).
 * - "firstDueDate" is set once to the earliest seen for this id.
 * - "nextSteps" should come from 00_Master Appointments (human-authored).
 * - "dvNotes" stores system reasons like "Auto: 12 days before appointment".
 */
function DVQ_upsert_({id, type, dueAt, customerName, nextSteps, dvNotes, status, assignedRep, assistedRep}) {
  if (!id)   throw new Error('DVQ_upsert_: missing id');
  if (!type) throw new Error('DVQ_upsert_: missing type');
  if (!dueAt)throw new Error('DVQ_upsert_: missing dueAt');

  var sh = DVQ_getQueueSheet_();
  var H  = DVQ_headerMap_(sh);
  DVQ_requireHeaders_(H);

  // Ensure dvNotes column exists (added once, only if missing)
  H = DVQ_getOrAddHeader_(sh, 'dvNotes');

  // Ensure exact rep columns on 04_Reminders_Queue (only if we have values)
  var iAssignedRepName = 0, iAssistedRepName = 0;
  if (assignedRep || assistedRep) {
    H = DVQ_getOrAddHeader_(sh, 'assignedRepName');  // exact queue header
    H = DVQ_getOrAddHeader_(sh, 'assistedRepName');  // exact queue header
    iAssignedRepName = H['assignedRepName'] || 0;
    iAssistedRepName = H['assistedRepName'] || 0;
  } else {
    iAssignedRepName = H['assignedRepName'] || 0;
    iAssistedRepName = H['assistedRepName'] || 0;
  }


  var lastRow = sh.getLastRow();
  var iId = H['id'], iNext = H['nextDueAt'], iFirst = H['firstDueDate'], iStatus = H['status'];
  var foundRow = 0;

  if (lastRow >= 2 && iId) {
    var ids = sh.getRange(2, iId, lastRow-1, 1).getDisplayValues();
    for (var r=0;r<ids.length;r++){
      if (String(ids[r][0]||'').trim() === id) { foundRow = r+2; break; }
    }
  }

  var now = new Date();
  var due = new Date(dueAt);
  var statusVal = (String(status || '').trim().toUpperCase()) || 'PENDING';

  if (foundRow) {
    // Earlier nextDueAt wins; firstDueDate stays earliest
    var existingNext  = iNext  ? DVQ_toDate_(sh.getRange(foundRow, iNext ).getDisplayValue())  : null;
    var existingFirst = iFirst ? DVQ_toDate_(sh.getRange(foundRow, iFirst).getDisplayValue()) : null;

    var due = DVQ_toDate_(dueAt) || new Date(dueAt);
    var newNext  = (!existingNext  || due < existingNext)  ? due : existingNext;
    var newFirst = (!existingFirst || due < existingFirst) ? due : existingFirst;

    if (iNext)  sh.getRange(foundRow, iNext ).setValue(DVQ_fmtPT_(newNext));
    if (iFirst) sh.getRange(foundRow, iFirst).setValue(DVQ_fmtDateOnlyPT_(newFirst));

    if (H['type']) sh.getRange(foundRow, H['type']).setValue(type);
    if (H['customerName'] && customerName) sh.getRange(foundRow, H['customerName']).setValue(customerName);
    if (iStatus) {
      var st = String(sh.getRange(foundRow, iStatus).getDisplayValue()||'').trim();
      if (!st) sh.getRange(foundRow, iStatus).setValue(statusVal);
    }
    if (H['nextSteps'] && typeof nextSteps === 'string' && nextSteps.length) {
      sh.getRange(foundRow, H['nextSteps']).setValue(nextSteps);
    }
    if (H['dvNotes'] && typeof dvNotes === 'string' && dvNotes.length) {
      sh.getRange(foundRow, H['dvNotes']).setValue(dvNotes);
    }

    // Write reps if columns exist and values provided
    if (iAssignedRepName && assignedRep) {
      sh.getRange(foundRow, iAssignedRepName).setValue(String(assignedRep));
    }
    if (iAssistedRepName && assistedRep) {
      sh.getRange(foundRow, iAssistedRepName).setValue(String(assistedRep));
    }

  } else {
    // New row
    var rowArr = new Array(Math.max(1, sh.getLastColumn())).fill('');
    var set = function(c,v){ if (H[c]) rowArr[H[c]-1] = v; };

    set('id', id);
    set('type', type);
    set('firstDueDate', DVQ_fmtDateOnlyPT_(due));
    set('nextDueAt',   DVQ_fmtPT_(due));
    set('status', statusVal);
    if (customerName) set('customerName', customerName);
    if (typeof nextSteps === 'string' && nextSteps.length) set('nextSteps', nextSteps);
    if (typeof dvNotes === 'string'   && dvNotes.length)   set('dvNotes', dvNotes);
    if (H['createdAt']) set('createdAt', now);
    if (H['createdBy']) set('createdBy', 'system:dv');

    // Reps on new rows (respect exact queue headers)
    if (iAssignedRepName) rowArr[iAssignedRepName - 1] = String(assignedRep || '');
    if (iAssistedRepName) rowArr[iAssistedRepName - 1] = String(assistedRep || '');

    sh.getRange(sh.getLastRow()+1, 1, 1, rowArr.length).setValues([rowArr]);
  }

  SpreadsheetApp.flush();
  return { ok:true, id:id, row: (foundRow || sh.getLastRow()), nextDueAt: due.toString() };
}

// ----- DV-specific convenience wrappers (id format uses RootApptID) -----

function DVQ_id_(rootApptId, remType, optDayKey){
  // Example: DV|ROOT123|DV_PROPOSE_NUDGE   or   DV|ROOT123|DV_URGENT_OTW_DAILY|20250902
  var parts = ['DV', String(rootApptId||'').trim(), String(remType||'').trim()];
  if (optDayKey) parts.push(String(optDayKey));
  return parts.join('|');
}

/** 12 days BEFORE appointment */
function DV_upsertProposeNudge_forAppt_({rootApptId, apptDateTimeISO, customerName, nextStepsFromMaster, assignedRep, assistedRep}) {
  var tz = Session.getScriptTimeZone?.() || 'America/Los_Angeles';
  var appt = new Date(apptDateTimeISO);
  var apptDayKey = Utilities.formatDate(appt, tz, 'yyyyMMdd');
  var dueLocal = DVQ_setLocalTime_(DVQ_addDays_(appt, -DV.POLICY.PROPOSE_NUDGE_OFFSET_DAYS), 9, 0);

  return DVQ_upsert_({
    id:   DVQ_id_(rootApptId, DV.REMTYPE.PROPOSE_NUDGE, apptDayKey), // ← now includes date
    type: DV.REMTYPE.PROPOSE_NUDGE,
    dueAt: dueLocal,
    customerName: customerName || '',
    nextSteps: nextStepsFromMaster || '',
    dvNotes: 'Auto: 12 days before appointment',
    status: 'PENDING'
  });
}


/** 2 days AFTER setting NEED_TO_PROPOSE */
function DV_upsertProposeNudge_afterStatus_({rootApptId, customerName, nextStepsFromMaster}) {
  var dueLocal = DVQ_setLocalTime_(DVQ_addDays_(new Date(), DV.POLICY.PROPOSE_AFTER_STATUS_DAYS), 9, 0);
  return DVQ_upsert_({
    id:   DVQ_id_(rootApptId, DV.REMTYPE.PROPOSE_NUDGE),
    type: DV.REMTYPE.PROPOSE_NUDGE,
    dueAt: dueLocal,
    customerName: customerName || '',
    nextSteps: nextStepsFromMaster || '',            // ← from Master
    dvNotes: 'Auto: 2 days after NEED_TO_PROPOSE',   // ← system note
    status: 'PENDING'
  });
}


/** One daily URGENT for TODAY, deduped by yyyyMMdd in id */
function DV_upsertUrgentDaily_forToday_({rootApptId, customerName, nextStepsFromMaster}) {
  var dayKey = DVQ_todayKey_();
  var dueLocal = DVQ_setLocalTime_(new Date(), 9, 0);
  return DVQ_upsert_({
    id:   DVQ_id_(rootApptId, DV.REMTYPE.URGENT_OTW_DAILY, dayKey),
    type: DV.REMTYPE.URGENT_OTW_DAILY,
    dueAt: dueLocal,
    customerName: customerName || '',
    nextSteps: nextStepsFromMaster || '',                 // ← from Master
    dvNotes: 'Auto: daily urgent (within 7 days, not On the Way)', // ← system note
    status: 'PENDING'
  });
}


// ----- TEST (safe, far-future) -----

function DV__test_upsertQueue_v2b(){
  var root = 'TEST-ROOT-APPT-002';
  var apptIso = '2030-05-20T15:00:00-07:00'; // far future to avoid real send
  var ns = 'From Master: call vendor for 3 stones';

  // 12-days-before candidate (earlier)
  var r1 = DV_upsertProposeNudge_forAppt_({
    rootApptId: root,
    apptDateTimeISO: apptIso,
    customerName: 'Test Customer',
    nextStepsFromMaster: ns
  });

  // 2-days-after candidate (later) — earlier-wins should keep 12-days-before
  var r2 = DV_upsertProposeNudge_afterStatus_({
    rootApptId: root,
    customerName: 'Test Customer',
    nextStepsFromMaster: ns
  });

  // Daily urgent (today)
  var r3 = DV_upsertUrgentDaily_forToday_({
    rootApptId: root,
    customerName: 'Test Customer',
    nextStepsFromMaster: ns
  });

  Logger.log(JSON.stringify({r1:r1, r2:r2, r3:r3}, null, 2));
}

function DVQ__normalizeDatesOnce() {
  var sh = DVQ_getQueueSheet_(), H = DVQ_headerMap_(sh);
  var iType=H['type'], iNext=H['nextDueAt'], iFirst=H['firstDueDate'];
  if (!iType || !iNext || !iFirst) throw new Error('Queue headers missing.');
  var last = sh.getLastRow(); if (last < 2) return 0;

  var types = sh.getRange(2, iType, last-1, 1).getDisplayValues();
  var nexts = sh.getRange(2, iNext, last-1, 1).getDisplayValues();
  var firsts= sh.getRange(2, iFirst,last-1, 1).getDisplayValues();

  var changed = 0;
  for (var r=0; r<types.length; r++) {
    var t = String(types[r][0]||'').toUpperCase();
    if (t.indexOf('DV_') !== 0) continue;

    var nd = DVQ_toDate_(nexts[r][0]);
    var fd = DVQ_toDate_(firsts[r][0]);
    if (nd) sh.getRange(r+2, iNext ).setValue(DVQ_fmtPT_(nd));
    if (fd) sh.getRange(r+2, iFirst).setValue(DVQ_fmtDateOnlyPT_(fd));
    if (nd || fd) changed++;
  }
  return changed;
}

/** File: 12.3 - dv_backfill_reps.gs (v1.0)
 * Purpose: Backfill missing assignedRepName / assistedRepName on 04_Reminders_Queue
 *          using "Assigned Rep" / "Assisted Rep" from 00_Master Appointments, keyed by RootApptID.
 * Safety : Only fills BLANK cells; never overwrites existing values; DV-only; idempotent.
 */

var DV_BACKFILL = Object.freeze({
  QUEUE_NAME:  '04_Reminders_Queue',
  MASTER_NAME: '00_Master Appointments'
});

// ---------- Safe fallbacks (no-ops if already defined elsewhere) ----------
if (typeof DVQ_getQueueSheet_ !== 'function') {
  function DVQ_getQueueSheet_() {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(DV_BACKFILL.QUEUE_NAME);
    if (!sh) throw new Error('Missing sheet "' + DV_BACKFILL.QUEUE_NAME + '".');
    return sh;
  }
}
if (typeof DVQ_headerMap_ !== 'function') {
  function DVQ_headerMap_(sh) {
    var lastCol = Math.max(1, sh.getLastColumn());
    var hdr = sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(function(h){ return String(h||'').trim(); });
    var H = {}; hdr.forEach(function(h,i){ if(h) H[h]=i+1; });
    return H;
  }
}
if (typeof DVQ_getOrAddHeader_ !== 'function') {
  function DVQ_getOrAddHeader_(sh, headerName) {
    var H = DVQ_headerMap_(sh);
    if (!H[headerName]) {
      sh.getRange(1, sh.getLastColumn()+1).setValue(headerName);
      SpreadsheetApp.flush();
      H = DVQ_headerMap_(sh);
    }
    return H;
  }
}
if (typeof DV_M_headerMap_ !== 'function') {
  function DV_M_headerMap_(sh) {
    var lastCol = Math.max(1, sh.getLastColumn());
    var hdr = sh.getRange(1,1,1,lastCol).getDisplayValues()[0] || [];
    var H = {}; for (var i=0;i<hdr.length;i++){ var k=String(hdr[i]||'').trim(); if(k) H[k]=i+1; }
    return H;
  }
}
if (typeof DV_M_pick_ !== 'function') {
  function DV_M_pick_(H, sh, row, names) {
    for (var i=0;i<names.length;i++){
      var n = names[i];
      if (H[n]) { return sh.getRange(row, H[n]).getValue(); }
    }
    return '';
  }
}
if (typeof DV_M_getVisitType_ !== 'function') {
  function DV_M_getVisitType_(H, sh, row) {
    var v = DV_M_pick_(H, sh, row, ['Visit Type','VisitType','Appointment Type','Appt Type']);
    return String(v||'').trim();
  }
}
if (typeof DV_M_getAssignedRep_ !== 'function') {
  function DV_M_getAssignedRep_(H, sh, row) {
    var v = DV_M_pick_(H, sh, row, ['Assigned Rep']);
    return String(v || '').trim();
  }
}
if (typeof DV_M_getAssistedRep_ !== 'function') {
  function DV_M_getAssistedRep_(H, sh, row) {
    var v = DV_M_pick_(H, sh, row, ['Assisted Rep']);
    return String(v || '').trim();
  }
}
if (typeof DV_M_getRootApptId_ !== 'function') {
  function DV_M_getRootApptId_(H, sh, row) {
    var v = DV_M_pick_(H, sh, row, ['RootApptID','RootApptId','APPT_ID','ApptID','APPT ID']);
    return String(v||'').trim();
  }
}

// ---------- Helpers ----------
/** Parse root from id like: 'DV|ROOT123|DV_PROPOSE_NUDGE|20250902' */
function DVQ_parseRootApptIdFromId_(id) {
  var s = String(id||'').trim();
  if (!s) return '';
  var m = s.match(/^DV\|([^|]+)\|/i);
  if (m) return m[1].trim();
  var parts = s.split('|');
  return parts.length >= 2 ? String(parts[1]||'').trim() : '';
}

/** Build a map: rootApptId -> {assigned, assisted, isDV} from 00_Master Appointments. */
function DVQ_buildMasterRepIndex__byRootApptId_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(DV_BACKFILL.MASTER_NAME);
  if (!sh) throw new Error('Missing sheet "' + DV_BACKFILL.MASTER_NAME + '".');

  var H = DV_M_headerMap_(sh);
  var iRoot    = H['RootApptID'] || H['RootApptId'] || H['APPT_ID'] || H['ApptID'] || H['APPT ID'];
  var iAssigned= H['Assigned Rep'];
  var iAssisted= H['Assisted Rep'];
  var iVisit   = H['Visit Type'] || H['VisitType'] || H['Appointment Type'] || H['Appt Type'];

  var lastRow = sh.getLastRow();
  if (!iRoot || (!iAssigned && !iAssisted) || lastRow < 2) return {};

  var idxCols = [iRoot];
  if (iVisit)    idxCols.push(iVisit);
  if (iAssigned) idxCols.push(iAssigned);
  if (iAssisted) idxCols.push(iAssisted);
  idxCols.sort(function(a,b){return a-b;});
  var minC = idxCols[0], maxC = idxCols[idxCols.length-1];
  var width = maxC - minC + 1;
  var vals = sh.getRange(2, minC, lastRow-1, width).getDisplayValues();

  var index = {};
  for (var r=0; r<vals.length; r++) {
    var row = vals[r];
    var root     = String(row[iRoot - minC] || '').trim();
    if (!root) continue;

    var assigned = iAssigned ? String(row[iAssigned - minC] || '').trim() : '';
    var assisted = iAssisted ? String(row[iAssisted - minC] || '').trim() : '';
    var visit    = iVisit    ? String(row[iVisit    - minC] || '').trim() : '';
    var isDV     = /diamond\s*viewing/i.test(visit);

    var cur = index[root];
    if (!cur) {
      index[root] = { assigned: assigned, assisted: assisted, isDV: isDV };
      continue;
    }
    // Prefer DV rows; otherwise, fill missing pieces only
    if (isDV && !cur.isDV) {
      index[root] = {
        assigned: assigned || cur.assigned,
        assisted: assisted || cur.assisted,
        isDV: true
      };
      continue;
    }
    if (!cur.assigned && assigned) cur.assigned = assigned;
    if (!cur.assisted && assisted) cur.assisted = assisted;
  }
  return index;
}

/**
 * Backfill missing assignedRepName/assistedRepName on DV reminders.
 * opts:
 *   dryRun: true (default) -> logs/returns preview, no writes.
 *   limit : number of changes to apply (0 = no limit).
 */
function DVQ_backfillReps_safe_(opts) {
  opts = opts || {};
  var dry = (opts.dryRun !== false);     // default true
  var maxChanges = Math.max(0, opts.limit || 0);

  var qSh = DVQ_getQueueSheet_();
  var qH  = DVQ_headerMap_(qSh);

  var iId   = qH['id'];
  var iType = qH['type'] || 0;
  if (!iId || !iType) throw new Error('04_Reminders_Queue must have headers "id" and "type".');

  // Ensure target columns (only add in APPLY mode)
  var iAssignedName = qH['assignedRepName'] || 0;
  var iAssistedName = qH['assistedRepName'] || 0;
  if (!dry) {
    if (!iAssignedName) { qH = DVQ_getOrAddHeader_(qSh, 'assignedRepName'); iAssignedName = qH['assignedRepName']; }
    if (!iAssistedName) { qH = DVQ_getOrAddHeader_(qSh, 'assistedRepName'); iAssistedName = qH['assistedRepName']; }
  }

  var last = qSh.getLastRow();
  if (last < 2) return { ok:true, dryRun:dry, scanned:0, updated:0 };

  // Read minimal rectangle
  var needCols = [iId, iType];
  if (iAssignedName) needCols.push(iAssignedName);
  if (iAssistedName) needCols.push(iAssistedName);
  needCols.sort(function(a,b){return a-b;});
  var minC = needCols[0], maxC = needCols[needCols.length-1], width = maxC - minC + 1;

  var data = qSh.getRange(2, minC, last-1, width).getDisplayValues();

  // Build master index once
  var masterIdx = DVQ_buildMasterRepIndex__byRootApptId_();

  var updated = 0, scanned = 0, skipped = 0, noChange = 0, unresolved = 0;
  var preview = [];

  for (var r=0; r<data.length; r++) {
    var row = data[r];
    var typeStr = String(row[iType - minC] || '').toUpperCase();
    if (typeStr.indexOf('DV_') !== 0) { skipped++; continue; } // DV-only

    var id = String(row[iId - minC] || '').trim();
    if (!id) { skipped++; continue; }

    var root = DVQ_parseRootApptIdFromId_(id);
    if (!root) { unresolved++; continue; }

    var curAssigned = iAssignedName ? String(row[iAssignedName - minC] || '').trim() : '';
    var curAssisted = iAssistedName ? String(row[iAssistedName - minC] || '').trim() : '';

    var needAssigned = !curAssigned;
    var needAssisted = !curAssisted;
    if (!needAssigned && !needAssisted) { noChange++; continue; }

    var ref = masterIdx[root] || null;
    if (!ref || (!ref.assigned && !ref.assisted)) { unresolved++; continue; }

    var newAssigned = needAssigned ? (ref.assigned || '') : curAssigned;
    var newAssisted = needAssisted ? (ref.assisted || '') : curAssisted;

    if (dry) {
      preview.push({
        row: r+2,
        id: id,
        rootApptId: root,
        assignedRepName: needAssigned ? newAssigned : '(keep)',
        assistedRepName: needAssisted ? newAssisted : '(keep)'
      });
      updated++;
      if (maxChanges && updated >= maxChanges) break;
      continue;
    }

    // Apply changes only when we have non-empty values to fill
    if (needAssigned && newAssigned && iAssignedName) {
      qSh.getRange(r+2, iAssignedName).setValue(newAssigned);
    }
    if (needAssisted && newAssisted && iAssistedName) {
      qSh.getRange(r+2, iAssistedName).setValue(newAssisted);
    }
    updated++;
    if (maxChanges && updated >= maxChanges) break;
  }

  var summary = {
    ok: true,
    dryRun: dry,
    scanned: data.length,
    updated: updated,
    skipped_nonDV_or_noId: skipped,
    unchanged_already_had_values: noChange,
    unresolved_no_master_match_or_blank: unresolved,
    preview: dry ? preview : undefined
  };

  if (typeof remind__dbg === 'function') {
    remind__dbg('DVQ_backfillReps_safe_', summary);
  } else {
    Logger.log(JSON.stringify(summary, null, 2));
  }

  return summary;
}

// ---------- Convenience runners ----------
function DVQ__test_backfillReps_dryRun() {
  var res = DVQ_backfillReps_safe_({ dryRun: true });
  Logger.log(JSON.stringify(res, null, 2));
  SpreadsheetApp.getActive().toast('DV backfill DRY RUN complete. See Executions/Logs.', 'DV Backfill', 5);
}

function DVQ__test_backfillReps_dryRun_100() {
  var res = DVQ_backfillReps_safe_({ dryRun: true, limit: 100 });
  Logger.log(JSON.stringify(res, null, 2));
  SpreadsheetApp.getActive().toast('DV backfill DRY RUN (first 100) complete.', 'DV Backfill', 5);
}

function DVQ__run_backfillReps_apply() {
  var res = DVQ_backfillReps_safe_({ dryRun: false });
  Logger.log(JSON.stringify(res, null, 2));
  SpreadsheetApp.getActive().toast('DV backfill APPLIED. See Executions/Logs.', 'DV Backfill', 5);
}

function DVQ__run_backfillReps_apply_100() {
  var res = DVQ_backfillReps_safe_({ dryRun: false, limit: 100 });
  Logger.log(JSON.stringify(res, null, 2));
  SpreadsheetApp.getActive().toast('DV backfill APPLIED (first 100).', 'DV Backfill', 5);
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



