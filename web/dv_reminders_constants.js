/** File: 12 - dv_reminders_constants.gs (v1.0)
 * Purpose: Central constants + helpers for Diamond-Viewing reminders.
 * Notes : No side effects; safe to add.
 * Policy: Calendar-day math per spec.
 */

var DV = Object.freeze({
  REMTYPE: Object.freeze({
    PROPOSE_NUDGE:      'DV_PROPOSE_NUDGE',       // Single nudge to propose diamonds
    URGENT_OTW_DAILY:   'DV_URGENT_OTW_DAILY',    // Daily URGENT starting T-7 until some OTW
    REPLACEMENT_NUDGE:  'DV_REPLACEMENT_NUDGE'    // Gentle follow-up when SOME_OTW (suggest more proposals)
  }),
  TPL: Object.freeze({
    PROPOSE_NUDGE:      'TPL_DV_PROPOSE_NUDGE',
    URGENT_OTW_DAILY:   'TPL_DV_URGENT_OTW_DAILY',
    REPLACEMENT_NUDGE:  'TPL_DV_REPLACEMENT_NUDGE'
  }),
  MASTER_STATUS: Object.freeze({
    // Aliases the Client Status UI may set for "Need to Propose"
    NEED_TO_PROPOSE_ALIASES: [
      'Need to Propose Diamonds for viewing',
      'Need to Propose',
      'Need to Propose Diamonds'
    ],
    PROPOSED:       'Diamond Memo - Proposed',
    OTW:            'Diamond Memo - On the Way',
    SOME_OTW:       'Diamond Memo - SOME On the Way',
    NOT_APPROVED:   'Diamond Memo - NOT APPROVED'
  }),
  POLICY: Object.freeze({
    PROPOSE_NUDGE_OFFSET_DAYS: 12, // 12 days BEFORE appt
    PROPOSE_AFTER_STATUS_DAYS: 2,  // 2 days AFTER Need-to-Propose status
    URGENT_WINDOW_DAYS:        7   // Daily URGENT for appts within next 7 days
  })
});

function DV_diagSO__test() {
  const samples = ["SO 12.3456"," 001293","0.1293","00.1293","'001293","SO#00.1293"];
  samples.forEach(x => Logger.log(x + " -> key=" + _soKey_(x) + " pretty=" + _soPretty_(x)));
}

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

// Add more “done” states + aliases
function DV__eqCI_(a,b){ return String(a||'').trim().toLowerCase() === String(b||'').trim().toLowerCase(); }
function DV__anyEqCI_(s, arr){ s=String(s||'').trim(); for (var i=0;i<arr.length;i++){ if (DV__eqCI_(s,arr[i])) return true; } return false; }

(function(){
  var _old = DV_normalizeMasterStatus;
  DV_normalizeMasterStatus = function(status){
    var norm = _old(status);
    if (norm) return norm;

    var s = String(status||'').trim();

    // Primary labels
    if (DV__eqCI_(s, 'Diamond Memo - Delivered'))         return 'DELIVERED';
    if (DV__eqCI_(s, 'Diamond Viewing Ready'))            return 'VIEWING_READY';
    if (DV__eqCI_(s, 'Diamond Deposit, Confirmed Order')) return 'DEPOSIT_CONFIRMED';

    // Common aliases
    if (DV__anyEqCI_(s, ['Memo Delivered','Diamonds Delivered','Delivered']))        return 'DELIVERED';
    if (DV__anyEqCI_(s, ['Viewing Ready','Ready for Viewing','DV Ready']))           return 'VIEWING_READY';
    if (DV__anyEqCI_(s, ['Deposit Confirmed','Deposit Received - Order Confirmed'])) return 'DEPOSIT_CONFIRMED';

    return ''; // unrecognized
  };
})();

/** Stop daily DV reminders when Master hits any of these CSOS final states. */
function DV_shouldStopDailyForStatus(status) {
  var s = String(status || '').trim().toLowerCase();
  if (!s) return false;

  // Existing OTW rules are still valid end conditions:
  if (DV_shouldStopUrgentForStatus(status)) return true; // covers OTW / SOME_OTW

  // Your requested aliases for final/ready states:
  var STOP_ALIASES = [
    'diamond memo - delivered',
    'diamond viewing ready',
    'diamond deposit, confirmed order',
    'deposit confirmed',
    'viewing ready',
    'memo delivered'
  ];

  return STOP_ALIASES.indexOf(s) !== -1;
}

/** Return true if the raw status indicates "Need to Propose" (handles aliases) */
function DV_isNeedToPropose(status) {
  var s = String(status || '').trim();
  return DV.MASTER_STATUS.NEED_TO_PROPOSE_ALIASES.some(function(a){
    return a.toLowerCase() === s.toLowerCase();
  });
}

/** Normalize a Master status string into a compact token ('' if not recognized) */
function DV_normalizeMasterStatus(status) {
  var s = String(status || '').trim();
  if (!s) return '';
  var map = {};
  map[DV.MASTER_STATUS.PROPOSED.toLowerCase()]     = 'PROPOSED';
  map[DV.MASTER_STATUS.OTW.toLowerCase()]          = 'OTW';
  map[DV.MASTER_STATUS.SOME_OTW.toLowerCase()]     = 'SOME_OTW';
  map[DV.MASTER_STATUS.NOT_APPROVED.toLowerCase()] = 'NOT_APPROVED';
  if (map[s.toLowerCase()]) return map[s.toLowerCase()];
  if (DV_isNeedToPropose(s)) return 'NEED_TO_PROPOSE';
  return '';
}

/** Policy: stop URGENT dailies when any stone is OTW (SOME_OTW or OTW) */
function DV_shouldStopUrgentForStatus(status) {
  var norm = DV_normalizeMasterStatus(status);
  return norm === 'OTW' || norm === 'SOME_OTW';
}

/** Use your existing TEAM_CHAT_WEBHOOK property for routing DV messages */
function DV_getTeamChatWebhook_() {
  var p = PropertiesService.getScriptProperties().getProperty('TEAM_CHAT_WEBHOOK') || '';
  return String(p || '').trim();
}

/** Lightweight self-test you can run from the editor */
function DV__selfTest(){
  var log = Logger.log;
  log('DV constants loaded. RemTypes: %s', JSON.stringify(DV.REMTYPE));
  log('Templates: %s', JSON.stringify(DV.TPL));
  log('Policy: %s', JSON.stringify(DV.POLICY));
  log('Webhook (TEAM_CHAT_WEBHOOK): %s', DV_getTeamChatWebhook_() ? 'present' : 'MISSING');
  // Sanity checks
  log('Normalize("Diamond Memo - On the Way") => %s', DV_normalizeMasterStatus('Diamond Memo - On the Way'));
  log('isNeedToPropose("Need to Propose") => %s', DV_isNeedToPropose('Need to Propose'));
  return true;
}

// Add more “done” states + aliases
// (Keep existing MASTER_STATUS keys; this augments behavior without breaking)
function DV__eqCI_(a,b){ return String(a||'').trim().toLowerCase() === String(b||'').trim().toLowerCase(); }
function DV__anyEqCI_(s, arr){ s=String(s||'').trim(); for (var i=0;i<arr.length;i++){ if (DV__eqCI_(s,arr[i])) return true; } return false; }


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
