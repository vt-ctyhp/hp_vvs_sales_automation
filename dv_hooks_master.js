/** File: 12.2 - dv_hooks_master.gs (v1.0)
 * Purpose: Read 00_Master Appointments row, detect "Diamond Viewing",
 *          and enqueue the 12-days-before "Propose" nudge into 04_Reminders_Queue.
 * Notes  : Header-mapped, additive-only; uses nextSteps from Master and dvNotes for system reason.
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

// ---------- Header helpers ----------

function DV_M_headerMap_(sh) {
  var lastCol = Math.max(1, sh.getLastColumn());
  var hdr = sh.getRange(1,1,1,lastCol).getDisplayValues()[0] || [];
  var H = {}; for (var i=0;i<hdr.length;i++){ var k=String(hdr[i]||'').trim(); if (k) H[k]=i+1; }
  return H;
}

function DV_M_pick_(H, sh, row, names) {
  for (var i=0;i<names.length;i++){
    var n = names[i];
    if (H[n]) { return sh.getRange(row, H[n]).getValue(); }
  }
  return '';
}

function DV_M_getApptISO_(H, sh, row) {
  // Try known ISO-style columns first (back-compat)
  var cands = [
    'ApptDateTimeISO','ApptDateTime (ISO)','Appointment Date/Time ISO','Appointment DateTime (ISO)',
    'Appt Start ISO','Appt Start (ISO)','Appt Start','Appointment Start','Event Start ISO','EventStartISO'
  ];
  var v = DV_M_pick_(H, sh, row, cands);
  if (v) {
    if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return v.toISOString();
    var s = String(v||'').trim(); var d = s ? new Date(s) : null;
    if (d && !isNaN(d)) return d.toISOString();
  }

  // NEW: Fallback to Visit Date + Visit Time (what your sheet has)
  var dVal = DV_M_pick_(H, sh, row, ['Visit Date','Appt Date','Appointment Date','Date','Event Date','Start Date']);
  if (!dVal) return ''; // no usable date

  var tVal = DV_M_pick_(H, sh, row, ['Visit Time','Appt Time','Appointment Time','Time','Event Time','Start Time']);

  // Join into a real Date in project TZ, then return ISO string
  var iso = DV_M_joinDateTimeToISO_(dVal, tVal); // helper below
  return iso || '';
}

/** Join date + time cells into an ISO string (TZ-aware, safe on partials). */
function DV_M_joinDateTimeToISO_(dVal, tVal) {
  var tz = Session.getScriptTimeZone?.() || 'America/Los_Angeles';

  // --- Parse date (Date object or string) ---
  var y, mo, da;
  if (Object.prototype.toString.call(dVal) === '[object Date]' && !isNaN(dVal)) {
    y  = dVal.getFullYear(); mo = dVal.getMonth(); da = dVal.getDate();
  } else {
    var ds = String(dVal||'').trim();
    var mYMD = ds.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    var mMDY = ds.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (mYMD) { y=+mYMD[1]; mo=+mYMD[2]-1; da=+mYMD[3]; }
    else if (mMDY) { y=+mMDY[3]; mo=+mMDY[1]-1; da=+mMDY[2]; }
    else return '';
  }

  // --- Parse time (Date object or string; default to 9:00 if blank) ---
  var hh = 9, mm = 0;
  if (tVal) {
    if (Object.prototype.toString.call(tVal) === '[object Date]' && !isNaN(tVal)) {
      hh = Number(Utilities.formatDate(tVal, tz, 'H'));
      mm = Number(Utilities.formatDate(tVal, tz, 'm'));
    } else {
      var ts = String(tVal||'').trim();
      var mt = ts.match(/^(\d{1,2})(?::(\d{2}))?\s*([ap]m)?$/i);
      if (mt) {
        hh = +mt[1]; mm = mt[2] ? +mt[2] : 0;
        var ap = (mt[3]||'').toLowerCase();
        if (ap === 'pm' && hh < 12) hh += 12;
        if (ap === 'am' && hh === 12) hh = 0;
      }
    }
  }

  // Construct local time, then return ISO
  var localStr = Utilities.formatDate(new Date(Date.UTC(y, mo, da, hh, mm, 0)), tz, 'yyyy/MM/dd HH:mm:ss');
  return new Date(localStr).toISOString();
}


function DV_M_getRootApptId_(H, sh, row) {
  var v = DV_M_pick_(H, sh, row, ['RootApptID','RootApptId','APPT_ID','ApptID','APPT ID']);
  return String(v||'').trim();
}

function DV_M_getCustomerName_(H, sh, row) {
  var v = DV_M_pick_(H, sh, row, ['Customer Name','Customer','Name']);
  return String(v||'').trim();
}

function DV_M_getNextSteps_(H, sh, row) {
  var v = DV_M_pick_(H, sh, row, ['Next Steps','NextSteps','Next steps']);
  return String(v||'').trim();
}

function DV_M_getAssignedRep_(H, sh, row) {
  // Exact header on 00_Master Appointments
  var v = DV_M_pick_(H, sh, row, ['Assigned Rep']);
  return String(v || '').trim();
}

function DV_M_getAssistedRep_(H, sh, row) {
  // Exact header on 00_Master Appointments
  var v = DV_M_pick_(H, sh, row, ['Assisted Rep']);
  return String(v || '').trim();
}


function DV_M_getVisitType_(H, sh, row) {
  var v = DV_M_pick_(H, sh, row, ['Visit Type','VisitType','Appointment Type','Appt Type']);
  return String(v||'').trim();
}

// ---------- Core hook (row-create scenario) ----------

/**
 * Detects "Diamond Viewing" on the given row and enqueues the 12-days-before Propose nudge.
 * Returns a summary; throws only on structural errors (missing sheet/row).
 */
function DV_tryEnqueueOnCreate_({ sh, row, dryRun }) {
  if (!sh) throw new Error('DV_tryEnqueueOnCreate_: missing sheet');
  if (!row || row < 2) throw new Error('DV_tryEnqueueOnCreate_: invalid row');

  var H = DV_M_headerMap_(sh);
  if (!H['Visit Type'] && !H['VisitType'] && !H['Appointment Type'] && !H['Appt Type']) {
    return { ok:false, skipped:true, reason:'No Visit Type column' };
  }

  var visit = DV_M_getVisitType_(H, sh, row);
  var isDV = /diamond\s*viewing/i.test(String(visit||''));
  if (!isDV) return { ok:false, skipped:true, reason:'Visit Type is not Diamond Viewing' };

  // Collect data
  var apptIso = DV_M_getApptISO_(H, sh, row);
  if (!apptIso) return { ok:false, skipped:true, reason:'Missing/invalid ApptDateTime ISO' };

  var rootApptId = DV_M_getRootApptId_(H, sh, row);
  if (!rootApptId) return { ok:false, skipped:true, reason:'Missing RootApptID' };

  var customerName = DV_M_getCustomerName_(H, sh, row);
  var nextStepsFromMaster = DV_M_getNextSteps_(H, sh, row) || '';

  if (dryRun) {
    return {
      ok:true, dryRun:true, would: 'DV_PROPOSE_NUDGE',
      rootApptId: rootApptId, apptIso: apptIso,
      customerName: customerName, nextSteps: nextStepsFromMaster
    };
  }

  // Enqueue: 12 days BEFORE appointment; nextSteps comes from Master; dvNotes explains reason
  // Add appt day-key so each DV booking gets its own nudge (e.g., 20250301)
  var apptDayKey = Utilities.formatDate(new Date(apptIso), Session.getScriptTimeZone() || 'America/Los_Angeles', 'yyyyMMdd');

  // Enqueue: 12 days BEFORE that specific appointment
  var assignedRep = DV_M_getAssignedRep_(H, sh, row);
  var assistedRep  = DV_M_getAssistedRep_(H, sh, row);

  var res = DVQ_upsert_({
    id:   DVQ_id_(rootApptId, DV.REMTYPE.PROPOSE_NUDGE, apptDayKey),
    type: DV.REMTYPE.PROPOSE_NUDGE,
    dueAt: DVQ_setLocalTime_(DVQ_addDays_(new Date(apptIso), -DV.POLICY.PROPOSE_NUDGE_OFFSET_DAYS), 9, 0),
    customerName: customerName || '',
    nextSteps: nextStepsFromMaster || '',
    dvNotes: 'Auto: 12 days before appointment',
    status: 'PENDING',
    assignedRep: assignedRep,   // ← NEW
    assistedRep: assistedRep    // ← NEW
  });



  return { ok:true, enqueued: true, id: res.id, nextDueAt: res.nextDueAt };
}

// ---------- Manual tester (safe) ----------

/**
 * Select a data row in 00_Master Appointments and run this.
 * If Visit Type = "Diamond Viewing", it will upsert the 12-days-before DV_PROPOSE_NUDGE.
 * Uses your queue’s earlier-wins rule, so it won’t create duplicates.
 */
function DV__test_tryEnqueueOnCreate_fromActiveRow(){
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet "00_Master Appointments".');

  var r = ss.getActiveRange();
  if (!r || r.getSheet().getName() !== sh.getName() || r.getRow() < 2) {
    throw new Error('Select a data row on "00_Master Appointments" and try again.');
  }

  var row = r.getRow();
  var out = DV_tryEnqueueOnCreate_({ sh: sh, row: row, dryRun: false });
  Logger.log(JSON.stringify(out, null, 2));
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



