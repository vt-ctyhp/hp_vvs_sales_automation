// @bundle: Ack pipes + dashboard + schedule + snapshot
/***** Phase 1 — Reminders → Acknowledgement Queue (top section)
 *  - Leaves all Acknowledgements logic untouched; injects a Reminders block at the top.
 *  - Uses feature flag REMINDERS_IN_ACK (must be TRUE to activate).
 *  - Reads from 04_Reminders_Queue; filters only items due now for the active rep.
 *  - Groups by Reminder Type; adds inline Action dropdown (Confirm/Snooze 1 Day/Snooze…/Cancel).
 *  - Does NOT submit anything yet (Phase 2 handles submission + audit).
 ****************************************************************/


// === SHEETS / HEADERS we rely on ===
var SHEET_REMINDERS_Q = '04_Reminders_Queue'; // set in Phase 0 header guard


// We reuse the existing queue headers function if present; otherwise fallback
function _queueHeadersFallback_() {
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
    'Ack Status',            // used as the "Reminder Action" selector in top section
    'Ack Note',              // notes for both ACK + Reminder rows
    'Reminder Snooze Until', // new (Phase 2)
    'Reminder ID'            // new (Phase 2) — hidden
  ];
}


// Reminder headers on 04_ (from our Phase 0 guard)
var REM_H = {
  ID: 'id',
  SO: 'soNumber',
  TYPE: 'type',
  FIRST_DUE: 'firstDueDate',
  NEXT_DUE: 'nextDueAt',
  RECURRENCE: 'recurrence',
  STATUS: 'status',
  SNOOZE_UNTIL: 'snoozeUntil',
  ASSIGNED_NAME: 'assignedRepName',
  ASSIGNED_EMAIL: 'assignedRepEmail',
  ASSISTED_NAME: 'assistedRepName',
  ASSISTED_EMAIL: 'assistedRepEmail',
  CUSTOMER: 'customerName',
  NEXT_STEPS: 'nextSteps',
  CREATED_AT: 'createdAt',
  CREATED_BY: 'createdBy',
  CONFIRMED_AT: 'confirmedAt',
  CONFIRMED_BY: 'confirmedBy',
  LAST_SENT_AT: 'lastSentAt',
  ATTEMPTS: 'attempts',
  LAST_ERROR: 'lastError',
  CANCEL_REASON: 'cancelReason',
  LAST_ADMIN_ACTION: 'lastAdminAction',
  LAST_ADMIN_BY: 'lastAdminBy'
};

// Action choices for Reminders (Phase 1 UI only)
var REM_ACTIONS = ['Confirm', 'Snooze 1 Day', 'Snooze…', 'Cancel'];

// Preferred type display & ordering for grouping (highest urgency first)
var TYPE_LABEL = {
  'DV_URGENT': 'Diamond Viewing — URGENT',                       // dynamic "(Appt in less than X days)" added later
  'DV_PROPOSE': 'Diamond Viewing — Need to Propose',             // dynamic "(Appt in less than X days)" added later
  'DV_PROPOSE_NUDGE': 'Diamond Viewing — Need to Propose',       // treated same as DV_PROPOSE
  'FOLLOW_UP': 'Need to Follow-Up',
  'COS': 'Custom Order Update Needed',                           // see COS trigger note you and I aligned on
  'START3D': 'Check on 3D Design',
  'ASSIGNSO': 'Check on 3D Design',
  'REV3D': 'Check on 3D Design',
  'OTHER': 'Other'
};

// Group order: DV first, then 3D tasks, then Follow-Up, then COS, then Other
var TYPE_ORDER = ['DV_URGENT', 'DV_PROPOSE', 'DV_PROPOSE_NUDGE', 'START3D', 'ASSIGNSO', 'REV3D', 'FOLLOW_UP', 'COS', 'OTHER'];

// Types that are day-granular → show when dueDate ≤ today (ignore time-of-day)
var DAY_GRANULAR_TYPES = new Set([
  'FOLLOW_UP', 'FOLLOWUP',                 // 👈 include both spellings
  'DV_PROPOSE', 'DV_PROPOSE_NUDGE',
  'DV_URGENT', 'START3D', 'ASSIGNSO', 'REV3D', 'COS'
]);

// Helper
function _isDayGranular_(t) { return DAY_GRANULAR_TYPES.has(_normType_(t)); }

// Format a Date (or string) to yyyy-mm-dd in script TZ
function _dateKey_(d, tz) {
  var dd = _toDateSafe_(d);
  if (!dd) return '';
  var z = (typeof tz === 'string' && tz) ? tz :
          (Session.getScriptTimeZone ? Session.getScriptTimeZone() : SpreadsheetApp.getActive().getSpreadsheetTimeZone());
  return Utilities.formatDate(dd, z, 'yyyy-MM-dd');
}


// Best-effort: compute "days until appointment" from 00_Master.
// Falls back silently if date/time isn’t available (no UI break).
function _apptDaysOutForItem_(it) {
  try {
    if (typeof _getMasterSnapshot_ !== 'function') return null;
    var master = _getMasterSnapshot_();
    if (!master || !master.soIdx) return null;

    // Try by SO#, then by customer name
    var soPretty = String(it.soNumber || '').trim();
    var soKey = soPretty.replace(/\D/g, '');
    if (soKey && master.soIdx.has(soKey)) return _daysFromMasterRow_(master.soIdx.get(soKey), master.idx);

    var custLC = String(it.customer || '').trim().toLowerCase();
    if (custLC && master.custIdx && master.custIdx.has(custLC)) return _daysFromMasterRow_(master.custIdx.get(custLC), master.idx);

    return null;
  } catch (_) { return null; }
}

function _daysFromMasterRow_(row, IDX) {
  try {
    var d = row[IDX.VISIT_DATE], t = row[IDX.VISIT_TIME];
    if (!(d instanceof Date) && d) d = new Date(String(d));
    if (t && !(t instanceof Date)) t = new Date(String(t));
    var when = null;
    if (d instanceof Date && !isNaN(d)) {
      if (t instanceof Date && !isNaN(t)) {
        when = new Date(d.getFullYear(), d.getMonth(), d.getDate(), t.getHours(), t.getMinutes());
      } else {
        when = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      }
    }
    if (!when || isNaN(when)) return null;
    var ms = when.getTime() - (new Date()).getTime();
    return Math.ceil(ms / (24*60*60*1000)); // days until (can be 0/today or negative/past)
  } catch (_) { return null; }
}

// Final display label for a reminder item (adds "(Appt in less than X days)" for DV* when possible)
function _displayLabelFor_(it) {
  var t = String(it.type || '').toUpperCase();
  var base = (TYPE_LABEL && TYPE_LABEL[t]) ? TYPE_LABEL[t] : (it.reminder || t || 'Reminder');

  // Only add the days hint for DV types; safe fallback if not computable
  if (t.indexOf('DV_URGENT') === 0 || t.indexOf('DV_PROPOSE') === 0) {
    var d = _apptDaysOutForItem_(it);
    if (typeof d === 'number' && isFinite(d)) {
      if (d <= 0) return base + ' (Appt today)';
      var s = (d === 1 ? '1 day' : (d + ' days'));
      return base + ' (Appt in less than ' + s + ')';
    }
  }
  return base;
}



// === Public helpers you can run from the editor or menus ===
function getRemindersInAckFlag_(){ return true; }

/** Rebuild my Q_ tab as usual, then insert Reminders on top if flag TRUE. */
function refreshMyQueueHybrid() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) { SpreadsheetApp.getUi().alert('Please wait a moment and try again.'); return; }
  try {
    var rep = (typeof detectRepName_ === 'function') ? detectRepName_() : '';
    if (!rep) { SpreadsheetApp.getUi().alert('Could not detect your Rep name.'); return; }

    // 1) Build ack queue exactly as today
    if (typeof refreshMyQueue === 'function') {
      try { refreshMyQueue(); } catch (e) { Logger.log('refreshMyQueue() failed: ' + e); }
    } else if (typeof buildTodaysQueuesAll === 'function') {
      try { buildTodaysQueuesAll(); } catch (e2) { Logger.log('buildTodaysQueuesAll() failed: ' + e2); }
    }

    // 2) Inject Reminders on top (flag-gated, replace-mode active)
    _injectRemindersAfterBuild_(rep);
    
  } finally {
    lock.releaseLock();
  }
}


/** Rebuild all reps’ Q_ tabs as usual, then insert Reminders on top (flag TRUE). */
function buildTodaysQueuesAll_WithReminders() {
  // Bulk job: build ACK queues only. Reminders are injected on-demand by the owner refresh.
  if (typeof buildTodaysQueuesAll === 'function') buildTodaysQueuesAll();
}


// === Core: injects the Reminders section into Q_<rep> at row 2 ===
function _injectRemindersAfterBuild_(rep) {
  try {
    if (typeof getRemindersInAckFlag_ === 'function' && !getRemindersInAckFlag_()) {
      Logger.log('REMINDERS_IN_ACK is OFF — skipping injection.'); 
      return;
    }

    var ss = SpreadsheetApp.getActive();
    var sh = _ensureQueueSheet_(rep);
    _ensureReminderBridgeColumnsOnQueue_(sh);
    // Re-read headers from the sheet so we use the full live width (incl. new columns)
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn())
                    .getDisplayValues()[0]
                    .map(function (s) { return String(s || '').trim(); });

    if (!headers || !headers.length) throw new Error('Queue headers not available.');

    // Replace-mode: remove any previously injected REMINDERS block
    _removeExistingReminderSection_(sh, headers);

    // Find the Action + Notes columns (we reuse Ack Status / Ack Note)
    var colAction = headers.indexOf('Ack Status') + 1;
    var colNotes  = headers.indexOf('Ack Note') + 1;

    // Gather due reminders for this rep
    var reminders = _readDueRemindersForRep_(rep);
    if (!reminders.length) { Logger.log('No due reminders for ' + rep); return; }

    // Build a matrix of rows to insert (includes section headers + grouped blocks)
    var matrix = _buildReminderRowsMatrix_(reminders, headers);

    Logger.log('[DBG matrix] rows=%s, cols=%s; first reminder row ID sample="%s"',
      matrix.length, headers.length,
      (function(){
        for (var i=0;i<matrix.length;i++){
          // skip banner/subsection/spacer (column A starts with '—' or blank)
          if (String(matrix[i][0]||'').indexOf('—') === 0) continue;
          var idCol = headers.indexOf('Reminder ID');
          if (idCol >= 0) return String(matrix[i][idCol] || '');
        }
        return '(none)';
      })()
    );

    // Insert above existing data body (keep header row at row 1)
    sh.insertRowsBefore(2, matrix.length);
    sh.getRange(2, 1, matrix.length, headers.length).setValues(matrix);

    // Insert a pink ACKNOWLEDGEMENTS header row right after the Reminders block
    var ackHeaderRowIndex = 2 + matrix.length;
    sh.insertRowsBefore(ackHeaderRowIndex, 1);

    var ackHdr = new Array(headers.length).fill('');
    ackHdr[0] = '— ACKNOWLEDGEMENTS —';
    sh.getRange(ackHeaderRowIndex, 1, 1, headers.length).setValues([ackHdr]);

    // Paint it pink; the ACK styler will keep it pink
    sh.getRange(ackHeaderRowIndex, 1, 1, sh.getLastColumn())
      .setBackground('#FFD1DC')
      .setFontColor('#000000')
      .setFontWeight('bold')
      .setFontStyle('normal')
      .setVerticalAlignment('middle')
      .setFontSize(14);


    // Data validation: apply the Reminders action DV only to the rows that contain reminders
    // (Skip section header rows—these have blanks in the Action column.)
    var dv = SpreadsheetApp.newDataValidation()
      .requireValueInList(REM_ACTIONS, true)
      .setAllowInvalid(false)
      .setHelpText('Choose Confirm / Snooze 1 Day / Snooze… / Cancel for Reminders.')
      .build();

    // Find contiguous reminder blocks where Action is blank vs dropdown-needed
    var start = 2, end = 2 + matrix.length - 1;
    for (var r = start; r <= end; r++) {
      var actionCell = sh.getRange(r, colAction);
      var val = String(actionCell.getValue() || '').trim();
      // Our matrix writes blank in Action; DV should still be applied.
      // But skip rows that are section headers (we mark them by writing "—" in column 1).
      var c1 = String(sh.getRange(r, 1).getValue() || '');
      var isSectionRow = c1.startsWith('===') || c1.startsWith('— ');
      if (!isSectionRow) actionCell.setDataValidation(dv);
    }

    // Light styling for the inserted section
    _styleReminderSection_(sh, 2, matrix.length, colAction, colNotes);
    _formatQueueHeaderAndColumns_(sh);

    // Re-style ACK section now that Reminders were inserted above it
    if (typeof styleAckSectionsWithTints_ === 'function') {
      styleAckSectionsWithTints_(sh);
    }

  } catch (err) {
    Logger.log('injectReminders failed for ' + rep + ': ' + err);
  }
}


function _buildReminderRowsMatrix_(items, queueHeaders) {
  // Shared columns (we reuse ACK columns for the Reminder UI)
  var actionIdx = queueHeaders.indexOf('Ack Status');            // shared Action column
  var notesIdx  = queueHeaders.indexOf('Ack Note');              // shared Notes column

  var rows = [];

  // 1) Overall banner row at the very top
  rows.push(_bannerRow_(
    queueHeaders.length,
    '— REMINDERS (due now) — complete first —',
    'Reminder Action',
    actionIdx,
    notesIdx
  ));

  if (!items || !items.length) return rows; // defensive (caller already checks)

  // 2) Group all items by their (possibly unfamiliar) type
  var byType = {};
  items.forEach(function (it) {
    var t = String(it.type || '').toUpperCase();
    if (!byType[t]) byType[t] = [];
    byType[t].push(it);
  });

  // 3) Build a stable ordering:
  //    - Known types in TYPE_ORDER keep their priority
  //    - Unknown types are appended, alphabetical, so they still render
  var priority = TYPE_ORDER.reduce(function (m, k, i) { m[String(k).toUpperCase()] = i; return m; }, {});
  var presentTypes = Object.keys(byType);

  presentTypes.sort(function (a, b) {
    var pa = (priority[a] != null) ? priority[a] : 999;
    var pb = (priority[b] != null) ? priority[b] : 999;
    if (pa !== pb) return pa - pb;
    // same priority bucket → alphabetical for stability
    return a < b ? -1 : (a > b ? 1 : 0);
  });

  // 4) Render each present type as its own subsection, then its rows
  presentTypes.forEach(function (tkey, idx) {
    var arr = byType[tkey] || [];
    if (!arr.length) return;

    var label = (TYPE_LABEL && TYPE_LABEL[tkey]) ? TYPE_LABEL[tkey] : tkey;

    // Subsection header (light hint in the action/notes columns)
    rows.push(_subsectionRow_(
      queueHeaders.length,
      '— ' + label + ' —',
      'Action: Confirm / Snooze / Cancel',
      actionIdx,
      notesIdx
    ));

    // Actual reminder rows in the queue’s shape
    arr.forEach(function (it) {
      rows.push(_remRowToQueueShape_(it, queueHeaders.length));
    });

    // Spacer after each type block (including the last—harmless and keeps visual rhythm)
    rows.push(_spacerRow_(queueHeaders.length));
  });

  return rows;
}


// === Readers & builders ===

function _readDueRemindersForRep_(rep) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_REMINDERS_Q);
  Logger.log('[DBG 04_] file="%s"  tab="%s"  gid=%s',
    ss.getName(), sh ? sh.getName() : '(null)', sh ? sh.getSheetId() : '(null)');
  if (!sh) return [];

  var data = sh.getDataRange().getValues();
  if (!data || data.length < 2) { Logger.log('[DBG 04_] empty or no data rows'); return []; }

  var headers = data[0].map(function(h){ return String(h || '').trim(); });
  Logger.log('[DBG 04_ headers] count=%s  %s', headers.length, JSON.stringify(headers));

  // Robust, normalized header indices
  var cID           = _hIdx_(headers, REM_H.ID, ['ID','Id']);
  var cTYPE         = _hIdx_(headers, REM_H.TYPE);
  var cNEXT_DUE     = _hIdx_(headers, REM_H.NEXT_DUE);
  var cSNOOZE_UNTIL = _hIdx_(headers, REM_H.SNOOZE_UNTIL);
  var cSTATUS       = _hIdx_(headers, REM_H.STATUS);
  var cASSIGNED_N   = _hIdx_(headers, REM_H.ASSIGNED_NAME);
  var cASSIGNED_E   = _hIdx_(headers, REM_H.ASSIGNED_EMAIL);
  var cASSISTED_N   = _hIdx_(headers, REM_H.ASSISTED_NAME);
  var cASSISTED_E   = _hIdx_(headers, REM_H.ASSISTED_EMAIL);
  var cCUSTOMER     = _hIdx_(headers, REM_H.CUSTOMER);
  var cNEXT_STEPS   = _hIdx_(headers, REM_H.NEXT_STEPS);
  var cSO           = _hIdx_(headers, REM_H.SO);

  // NEW: pick up RootApptID (if present) + Reminder text/title (if present)
  var cROOT         = _hIdx_(headers, 'rootApptId', ['RootApptID','Root Appointment ID','Root Appt ID']);
  var cREMINDER_TXT = _hIdx_(headers, 'Reminder', ['Reminder Text','Reminder Title','Reminder Name']);

  Logger.log('[DBG 04_ index] id=%s  type=%s  nextDue=%s  status=%s  customer=%s  so=%s  root=%s  reminderTxt=%s',
             cID, cTYPE, cNEXT_DUE, cSTATUS, cCUSTOMER, cSO, cROOT, cREMINDER_TXT);

  var now = new Date();
  var myEmail = _safeEmail_();

  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    var status   = String(row[cSTATUS] || '').toUpperCase();
    var type     = String(row[cTYPE]   || '').toUpperCase();
    var nextDue  = _toDateSafe_(row[cNEXT_DUE]);
    var snoozeTo = _toDateSafe_(row[cSNOOZE_UNTIL]);

    // Eligibility: pending/snoozed and due (date-only for day-granular types)
    if (!nextDue) continue;
    if (status !== 'PENDING' && status !== 'SNOOZED') continue;

    var tz = (Session.getScriptTimeZone && Session.getScriptTimeZone()) || ss.getSpreadsheetTimeZone();
    var todayKey = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var typeCanon = _normType_(type);
    var isDay = _isDayGranular_(typeCanon);

    // SNOOZED: respect snoozeUntil first
    if (status === 'SNOOZED' && snoozeTo) {
      if (isDay) {
        // show on/after snooze date
        if (_dateKey_(snoozeTo, tz) > todayKey) continue;
      } else {
        // time-critical
        if (snoozeTo > now) continue;
      }
    } else {
      // PENDING path: check nextDueAt against today (or now for time-critical)
      if (isDay) {
        if (_dateKey_(nextDue, tz) > todayKey) continue; // due in future date
      } else {
        if (nextDue > now) continue;                      // not yet due by time
      }
    }

    // Routing: rep involved (assigned or assisted) by name/email
    var aName = String(row[cASSIGNED_N] || '').trim();
    var aMail = String(row[cASSIGNED_E] || '').trim().toLowerCase();
    var sName = String(row[cASSISTED_N] || '').trim();
    var sMail = String(row[cASSISTED_E] || '').trim().toLowerCase();

    // Support comma-separated reps (e.g., "Val, Khoa")
    var repMatchByName =
      _listHasExactCaseFold_(_csvTokens_(aName), rep) ||
      _listHasExactCaseFold_(_csvTokens_(sName), rep);

    // Also allow comma-separated emails if they ever occur
    var repMatchByMail = false;
    if (myEmail) {
      repMatchByMail =
        _listHasExactCaseFold_(_csvTokens_(aMail), myEmail) ||
        _listHasExactCaseFold_(_csvTokens_(sMail), myEmail);
    }

    if (!(repMatchByName || repMatchByMail)) continue;

    out.push({
      id: String(row[cID] || ''),
      type: type,
      customer: row[cCUSTOMER],
      nextDueAt: nextDue,
      status: status,
      note: row[cNEXT_STEPS] || '',
      soNumber: row[cSO] || '',
      root: String(cROOT >= 0 ? row[cROOT] : ''),
      reminder: String(cREMINDER_TXT >= 0 ? row[cREMINDER_TXT] : '')
    });
  }

  Logger.log('[DBG 04_ → reminders pre‑enrich for rep="%s"] kept=%s', rep, out.length);

  // Enrich with 07_Root_Index, then sort by type priority → due date asc
  out = _enrichRemindersFromIndex_(out);

  var order = TYPE_ORDER.reduce(function(m, k, i){ m[k] = i; return m; }, {});
  out.sort(function(a,b){
    var ta = order[a.type] != null ? order[a.type] : 999;
    var tb = order[b.type] != null ? order[b.type] : 999;
    if (ta !== tb) return ta - tb;
    return a.nextDueAt.getTime() - b.nextDueAt.getTime();
  });

  Logger.log('[DBG 04_ → reminders enriched for rep="%s"] kept=%s', rep, out.length);
  return out;
}

function _remRowToQueueShape_(it, width) {
  var MS = 24*60*60*1000;
  // "Days Since Last Update" should reflect the last client status update, not reminder due.
  var daysSince = '';
  if (typeof it.daysSinceUpdate === 'number' && isFinite(it.daysSinceUpdate)) {
    daysSince = it.daysSinceUpdate;
  } else if (it.updatedAt instanceof Date) {
    var d = Math.floor((new Date().getTime() - it.updatedAt.getTime()) / MS);
    daysSince = d >= 0 ? d : '';
  }

  // Next Steps: enriched text → [Descriptive] — [Note] — [Reminder]; fallback to raw note.
  var nextSteps = it.nextStepsRich || it.note || '';

  var row = [
    String(it.root || ''),             // RootApptID
    String(it.customer || ''),         // Customer Name
    String(it.salesStage || ''),       // Sales Stage
    String(it.conversionStatus || ''), // Conversion Status
    String(it.cos || ''),              // Custom Order Status
    String(it.inProd || ''),           // In Production Status
    String(it.csos || ''),             // Center Stone Order Status
    nextSteps,                         // Next Steps (descriptive + note + reminder)
    String(it.updatedBy || ''),        // Updated By (Last Updated By)
    it.updatedAt || '',                // Updated At (Last Updated At)
    daysSince,                         // Days Since Last Update
    String(it.csrUrl || ''),           // Client Status Report URL
    '',                                // Ack Status (Reminder Action DV lives here)
    '',                                // Ack Note
    '',                                // Reminder Snooze Until (user-entered when Snooze…)
    String(it.id || '')                // Reminder ID (hidden col)
  ];

  while (row.length < width) row.push('');
  return row;
}

// === REMINDERS COLOR PALETTE (PLACEHOLDERS — set your hex codes) ===
// Tip: Keep these #RRGGBB. You can leave any blank "" to skip that styling.

var REM_COLOR = {
  // Banner "— REMINDERS (due now) — complete first —"
  BANNER_BG: '#D32F2F',   // ← red background (edit to your hex)
  BANNER_FG: '#FFFFFF',   // ← white text

  // Section title rows (the lines like "— Diamond Viewing —", "— Follow‑Up —")
  SECTION_TITLE_FG: '#000000',    // text color for section titles
  SECTION_TITLE_ODD_BG: '#FDE2E2', // alt-1 title bg (edit)
  SECTION_TITLE_EVEN_BG: '#E7F0FE',// alt-2 title bg (edit)

  // Body rows (the actual reminder rows under each section)
  SECTION_BODY_FG: '#000000',     // text color for body rows
  SECTION_BODY_ODD_BG: '#FFF8F0', // alt-1 body bg (edit)
  SECTION_BODY_EVEN_BG: '#F6FBFF' // alt-2 body bg (edit)
};


// === Styling helpers (lightweight) ===
function _bannerRow_(width, title, rightHint, actionIdx, notesIdx) {
  var r = new Array(width).fill('');
  r[0] = title;
  if (actionIdx >= 0) r[actionIdx] = rightHint;
  if (notesIdx >= 0) r[notesIdx]   = 'Notes';
  return r;
}
function _subsectionRow_(width, title, rightHint, actionIdx, notesIdx) {
  var r = new Array(width).fill('');
  r[0] = title;
  if (actionIdx >= 0) r[actionIdx] = rightHint;
  if (notesIdx >= 0) r[notesIdx]   = 'Notes';
  return r;
}
function _spacerRow_(width) { return new Array(width).fill(''); }

function _styleReminderSection_(sh, startRow, numRows, colAction, colNotes) {
  try {
    if (!numRows || numRows <= 0) return;
    const lc = sh.getLastColumn();
    const all = sh.getRange(startRow, 1, numRows, lc).getDisplayValues();
    const hdr = sh.getRange(1, 1, 1, lc).getDisplayValues()[0].map(s => String(s || '').trim());
    const cId = hdr.indexOf('Reminder ID') + 1;

    // Remove any banding that could bleed colors into our spacer rows
    (sh.getBandings() || []).forEach(function(b){ try { b.remove(); } catch(_){} });

    // vertical centering for entire inserted block
    sh.getRange(startRow, 1, numRows, lc).setVerticalAlignment('middle');

    const T1 = STYLE_THEME.TINT_1 || 0.86;
    const T2 = STYLE_THEME.TINT_2 || 0.92;

    // Helper: choose a base color for a subsection title
    function baseForReminderSection(titleText) {
      const raw   = String(titleText || '').trim();
      const clean = raw.replace(/^—\s*/, '').replace(/\s*—\s*$/, '');

      // Direct map (exact or UPPER)
      if (STYLE_THEME.STAGE_COLORS[clean])             return STYLE_THEME.STAGE_COLORS[clean];
      if (STYLE_THEME.STAGE_COLORS[clean.toUpperCase()]) return STYLE_THEME.STAGE_COLORS[clean.toUpperCase()];

      // Hyphen/dash class: -, ‑, –, —, −
      const hy = '[\\s\\-\\u2010\\u2011\\u2012\\u2013\\u2014\\u2212]?';

      // FOLLOWUP / Need to Follow‑Up → peach
      if (new RegExp('^FOLLOW' + hy + 'UP$', 'i').test(clean) ||
          /need\s+to\s+follow[\s\-\u2010\u2011\u2012\u2013\u2014\u2212]?up/i.test(clean)) {
        return STYLE_THEME.STAGE_COLORS['FOLLOWUP'] || '#FFD7C2';
      }

      // Other known subsections
      if (/custom\s+order\s+update\s+needed/i.test(clean)) return STYLE_THEME.STAGE_COLORS['Custom Order Update Needed'];
      if (/dv_urgent_otw_daily/i.test(clean))             return STYLE_THEME.STAGE_COLORS['DV_URGENT_OTW_DAILY'];
      if (/diamond\s+viewing|urgent/i.test(clean))        return STYLE_THEME.STAGE_COLORS['DV_URGENT'];

      // Fallback = reminders banner red
      return STYLE_THEME.STAGE_COLORS.REMINDERS || '#C5221F';
    }


    // Pass 1: style the banner row (exact text match) and subsection title rows
    for (let i = 0; i < all.length; i++) {
      const r  = startRow + i;
      const c1 = String(all[i][0] || '');
      const isBanner = (c1.trim() === '— REMINDERS (due now) — complete first —');
      const isSectionRow = (!isBanner && c1.startsWith('— '));

      if (isBanner) {
        sh.getRange(r, 1, 1, lc)
          .setBackground(STYLE_THEME.STAGE_COLORS.REMINDERS || '#C5221F')
          .setFontColor('#FFFFFF')
          .setFontWeight('bold')
          .setFontStyle('normal')
          .setFontSize(14);    // ← add this
        continue;
      } else if (isSectionRow) {
        const base  = baseForReminderSection(c1);
        const clean = String(c1).replace(/^—\s*/, '').replace(/\s*—\s*$/, '');
        const needsWhite = /DV_URGENT|URGENT|DV_URGENT_OTW_DAILY/i.test(clean) ||
                           String(base).toUpperCase() === '#C5221F';

        sh.getRange(r, 1, 1, lc)
          .setBackground(base)
          .setFontStyle('italic')
          .setFontWeight('bold')
          .setFontColor(needsWhite ? '#FFFFFF' : '#000000');
      }
    }

    // Pass 2: tint body rows underneath each subsection title
    let currentBase = STYLE_THEME.STAGE_COLORS.REMINDERS || '#E53935';
    let flip = false;

    for (let i = 0; i < all.length; i++) {
      const r  = startRow + i;
      const c1 = String(all[i][0] || '');
      const isBanner = (c1.trim() === '— REMINDERS (due now) — complete first —');
      const isSectionRow = (!isBanner && c1.startsWith('— '));
      const isSpacer = (c1 === '');

      if (isSpacer) {
        // Explicitly white to avoid banding/theme tints showing up as blue
        sh.getRange(r, 1, 1, lc).setBackground('#FFFFFF');
        continue;
      }

      if (isSectionRow) {
        currentBase = baseForReminderSection(c1);
        flip = false; // reset alternation within the new block
        continue;
      }
      if (isBanner || isSpacer) continue;

      // body rows are those with a Reminder ID
      if (cId > 0) {
        const idVal = String(sh.getRange(r, cId).getDisplayValue() || '').trim();
        if (!idVal) continue;
      } else continue;

      flip = !flip;
      const bg = _tint_(currentBase, flip ? T1 : T2);
      sh.getRange(r, 1, 1, lc).setBackground(bg);
    }

    // keep hint styling in Action/Notes columns
    if (colAction > 0) sh.getRange(startRow, colAction, numRows, 1).setFontStyle('italic');
    if (colNotes  > 0) sh.getRange(startRow, colNotes,  numRows, 1).setFontStyle('italic');

  } catch (e) {
    Logger.log('Style (reminders) pass skipped: ' + e);
  }
}

function _formatQueueHeaderAndColumns_(sh) {
  try {
    var lc = sh.getLastColumn();
    var hdr = sh.getRange(1, 1, 1, lc).getDisplayValues()[0].map(function(s){ return String(s || '').trim(); });

    // 1) Header row: wrap + center the first (header) row
    sh.getRange(1, 1, 1, lc)
      .setWrap(true)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');


    // 2) Robust header lookups (supports alts + normalization)
    function colOf(name, alts) {
      var i = hdr.indexOf(name);
      if (i >= 0) return i + 1;
      if (alts && alts.length) {
        for (var j = 0; j < alts.length; j++) {
          var k = hdr.indexOf(alts[j]);
          if (k >= 0) return k + 1;
        }
      }
      var want = _normHeader_(name);
      for (var a = 0; a < hdr.length; a++) {
        if (_normHeader_(hdr[a]) === want) return a + 1;
      }
      return -1;
    }

        // === Uniform width + wrap for status columns ===
    var UNIFORM_STATUS_WIDTH_PX = 160; // adjust if you want narrower/wider

    var statusCols = [
      'Sales Stage',
      'Conversion Status',
      'Custom Order Status',
      'In Production Status',
      'Center Stone Order Status'
    ];

    statusCols.forEach(function(label){
      var c = colOf(label);
      if (c > 0) {
        // Same width for all five columns
        sh.setColumnWidth(c, UNIFORM_STATUS_WIDTH_PX);
        // Wrap text for the entire column (header + body)
        sh.getRange(1, c, sh.getMaxRows()).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      }
    });

    var cCSR  = colOf('Client Status Report URL', ['CSR URL']);
    var cNext = colOf('Next Steps');

    // --- NEW: Center + compact "Days Since Last Update"
    var cDays = colOf('Days Since Last Update', ['Days Since Last Update (calc)','Days Since Update']);
    if (cDays > 0) {
      var HALF_WIDTH_PX = Math.max(60, Math.floor(UNIFORM_STATUS_WIDTH_PX / 2)); // "half the size"
      // Center both header + body for this column
      sh.getRange(1, cDays, sh.getMaxRows(), 1)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      sh.setColumnWidth(cDays, HALF_WIDTH_PX);
    }

    // --- NEW: "Reminder Snooze Until" → Date-only validation + a bit wider
    var cSnooze = colOf('Reminder Snooze Until', ['Snooze Until','Snooze']);
    if (cSnooze > 0) {
      sh.setColumnWidth(cSnooze, 140); // a little bit wider for date UI
      var lr = sh.getLastRow();
      if (lr >= 2) {
        var dateDv = SpreadsheetApp.newDataValidation()
          .requireDate()
          .setAllowInvalid(false)
          .setHelpText('Pick a date (yyyy-mm-dd).')
          .build();
        sh.getRange(2, cSnooze, lr - 1, 1)
          .setDataValidation(dateDv)
          .setNumberFormat('yyyy-mm-dd');
      }
    }


    // 3) Wrap strategies
    if (cNext > 0) {
      sh.getRange(1, cNext, sh.getMaxRows()).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    }
    if (cCSR > 0) {
      sh.getRange(1, cCSR, sh.getMaxRows()).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    }

    // 4) Pin column widths so the sheet doesn't re-shrink each run.
    //    Tweak these to taste.
    var NEXT_STEPS_WIDTH_PX = 240; // ≈ half of a common 480px working width
    var CSR_URL_WIDTH_PX    = 140; // ≈ one‑third of a common 420px working width

    if (cNext > 0) sh.setColumnWidth(cNext, NEXT_STEPS_WIDTH_PX);
    if (cCSR  > 0) sh.setColumnWidth(cCSR,  CSR_URL_WIDTH_PX);

  } catch (e) {
    Logger.log('format header/columns skipped: ' + e);
  }
}

/** Style ACK rows with stage-based tints + middle vertical alignment */
function styleAckSectionsWithTints_(sh) {
  try {
    const lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (lr < 2) return;

    const hdr = sh.getRange(1, 1, 1, lc).getDisplayValues()[0].map(s => String(s || '').trim());

    // --- helper: tolerant column resolver
    const colOf = (name, alts) => {
      let i = hdr.indexOf(name); if (i >= 0) return i + 1;
      if (alts) for (const a of alts) { i = hdr.indexOf(a); if (i >= 0) return i + 1; }
      return -1;
    };
    const cRoot  = colOf('RootApptID', ['Root Appt ID','Root Appointment ID']);
    const cStage = colOf('Sales Stage');
    const cRemId = colOf('Reminder ID'); // may be hidden

    if (cRoot < 1 || cStage < 1) return;

    // === NEW: find where ACK styling should start (at/after “— ACKNOWLEDGEMENTS —”).
    // Prefer explicit marker; else compute the end of the Reminders block and begin after it.
    const colA = sh.getRange(2, 1, lr - 1, 1).getDisplayValues().map(r => String(r[0] || '').trim());
    let ackStart = colA.findIndex(v => v === '— ACKNOWLEDGEMENTS —');
    if (ackStart >= 0) {
      ackStart = 2 + ackStart;   // sheet row index
    } else {
      // fallback: find end of REMINDERS block using the same scan as the cleaner
      const bannerText = '— REMINDERS (due now) — complete first —';
      const startIdx = colA.findIndex(v => v === bannerText);
      if (startIdx >= 0) {
        let start = 2 + startIdx;
        const idCol = Math.max(1, (hdr.indexOf('Reminder ID') + 1));
        const scan = sh.getRange(start, 1, lr - start + 1, idCol).getDisplayValues();
        let end = start;
        for (let r = 0; r < scan.length; r++) {
          const a = String(scan[r][0] || '');
          const id = idCol ? String(scan[r][idCol - 1] || '').trim() : '';
          const isSection = a.indexOf('— ') === 0;
          const isSpacer  = a === '';
          const isRemRow  = !!id;
          if (isSection || isSpacer || isRemRow) {
            end = start + r;
          } else {
            break; // boundary reached
          }
        }
        ackStart = end + 1; // first row after the reminders block
      } else {
        ackStart = 2; // no reminders; style whole body as before
      }
    }

    // Vertical center the whole data area once
    sh.getRange(2, 1, lr - 1, lc).setVerticalAlignment('middle');

    // remove any banding that might mask our colors
    (sh.getBandings() || []).forEach(b => { try { b.remove(); } catch (_) {} });

    const T1 = STYLE_THEME.TINT_1 || 0.86;
    const T2 = STYLE_THEME.TINT_2 || 0.92;

    // Map ACK section header text → base color (null = not an ACK section)
    function resolveAckBase(titleRaw) {
      const t = String(titleRaw || '')
        .replace(/^—\s*/, '')
        .replace(/\s*—\s*$/, '')
        .replace(/\(\s*\d+\s*\)\s*$/, '')
        .trim()
        .toLowerCase();

      if (/^acknowledgements?$/.test(t)) return '#ec84b5'; // pink ACK header row
      if (/\bappointment\b/.test(t))     return STYLE_THEME.STAGE_COLORS['Appointment'];
      if (/viewing\s*scheduled/.test(t)) return STYLE_THEME.STAGE_COLORS['Viewing Scheduled'];
      if (/hot\s*lead/.test(t))          return STYLE_THEME.STAGE_COLORS['Hot Lead'];
      // accept -, -, –, —, − in “Follow-Up”
      if (/follow[\s\-\u2010\u2011\u2012\u2013\u2014\u2212]?up/.test(t))
        return STYLE_THEME.STAGE_COLORS['Follow-Up Required'] || '#C5221F';
      if (/in\s*production/.test(t))     return STYLE_THEME.STAGE_COLORS['In Production'];
      return null; // not an ACK section
    }

    // Read only once for performance
    const width = Math.max(cStage, cRemId, cRoot);
    const body  = sh.getRange(2, 1, lr - 1, width).getDisplayValues();

    let currentBase = STYLE_THEME.STAGE_COLORS.ACK_DEFAULT || '#AECBFA';
    let flip = false;

    for (let i = 0; i < body.length; i++) {
      const r  = 2 + i;
      if (r < ackStart) continue;               // ← KEY: never touch the Reminders block
      const c1 = String(body[i][0] || '');
      const isSectionRow = c1.startsWith('— ');
      const hasRoot      = String(body[i][(cRoot-1)] || '').trim() !== '';
      const isReminder   = (cRemId > 0) ? String(body[i][(cRemId-1)] || '').trim() !== '' : false;

      if (isSectionRow) {
        const base = resolveAckBase(c1);
        if (base) {
          currentBase = base;
          flip = false;
          sh.getRange(r, 1, 1, lc)    // style the ACK section header line
            .setBackground(currentBase)
            .setFontWeight('bold')
            .setFontStyle('normal')
            .setFontColor('#000000');
        }
        continue; // skip non-ACK section rows silently
      }

      // Only tint ACK data rows (skip reminder rows and empties)
      if (!hasRoot || isReminder) continue;

      flip = !flip;
      sh.getRange(r, 1, 1, lc).setBackground(_tint_(currentBase, flip ? T1 : T2));
    }

  } catch (e) {
    Logger.log('Style (ACK) pass skipped: ' + e);
  }
}


/* === Reminders: replace-mode cleaner (deletes any prior injected block) ===
 * Deletes the contiguous "REMINDERS (due now)" section starting at the first banner
 * until the first non-section, non-spacer, non-reminder row (data row).
 */
function _removeExistingReminderSection_(sh, knownHeaders) {
  try {
    var lr = sh.getLastRow();
    if (lr < 2) return;

    // Resolve headers (use provided or read live)
    var headers = Array.isArray(knownHeaders) && knownHeaders.length
      ? knownHeaders
      : sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function (s) { return String(s || '').trim(); });

    var idCol = headers.indexOf('Reminder ID') + 1; // hidden col added by bridge
    if (idCol < 1) idCol = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].indexOf('Reminder ID') + 1;

    // Find the banner "— REMINDERS (due now) — complete first —" in column A
    var colA = sh.getRange(2, 1, Math.max(lr - 1, 0), 1).getDisplayValues();
    var bannerText = '— REMINDERS (due now) — complete first —';
    var start = -1;
    for (var i = 0; i < colA.length; i++) {
      if (String(colA[i][0] || '').trim() === bannerText) { start = 2 + i; break; }
    }
    if (start === -1) return; // nothing to remove

    // Walk down from start to find the end of the block.
    // Inside the block if: column A starts with "— " (banner/subsection),
    // or the row is a blank spacer (col A empty),
    // or the row is a real reminder (has a Reminder ID).
    var maxRows = lr - start + 1;
    var scan = sh.getRange(start, 1, maxRows, Math.max(idCol, 1)).getDisplayValues();
    var end = start;
    for (var r = 0; r < scan.length; r++) {
      var a = String(scan[r][0] || '');
      var id = idCol > 0 ? String(scan[r][idCol - 1] || '').trim() : '';
      var isSectionRow = a.indexOf('— ') === 0;
      var isSpacer = a === '';
      var isReminderRow = !!id;
      if (isSectionRow || isSpacer || isReminderRow) {
        end = start + r; // still inside block
        continue;
      }
      break; // first non-section, non-spacer, non-reminder row = boundary
    }

    var count = end - start + 1;
    if (count > 0) sh.deleteRows(start, count);
  } catch (e) {
    Logger.log('[cleaner] skip: ' + e);
  }
}


// === Sheet/query utilities ===
function _ensureQueueSheet_(rep) {
  // use your existing helper if present
  if (typeof ensureQueueSheet_ === 'function') return ensureQueueSheet_(rep);
  // otherwise ensure the sheet exists
  var ss = SpreadsheetApp.getActive();
  var name = 'Q_' + String(rep || '').trim();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  // write queue headers if missing
  var headers = (typeof queueHeaders_ === 'function') ? queueHeaders_() : _queueHeadersFallback_();
  var haveCols = sh.getLastColumn();
  if (!haveCols) sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

function _allQueueReps_() {
  // Reads 08_Reps_Map for list of reps to update (de-duped)
  var ss = SpreadsheetApp.getActive();
  var s08 = ss.getSheetByName('08_Reps_Map');
  var reps = new Set();
  if (!s08) return [];
  var rows = s08.getDataRange().getValues();
  if (rows.length < 2) return [];
  var hdrs = rows[0].map(function(h){ return String(h || '').trim(); });
  var idxRep = hdrs.indexOf('Rep');
  var idxIncl = hdrs.indexOf('Include? (Y/N)');
  for (var i = 1; i < rows.length; i++) {
    var include = String(rows[i][idxIncl] || '').toUpperCase();
    if (include !== 'Y') continue;
    var rep = String(rows[i][idxRep] || '').trim();
    if (rep) reps.add(rep);
  }
  return Array.from(reps.values());
}

function _toDateSafe_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  var n = Number(v);
  if (!isNaN(n) && n > 0) return new Date(n);
  var d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function _csvTokens_(s) {
  // Split on comma (and Chinese comma) + optional semicolon; trim + de‑junk
  return String(s == null ? '' : s)
    .split(/[,\uFF0C;]+/)
    .map(function(t){ return String(t).replace(/[\u200B-\u200D\uFEFF\u00A0]/g,' ').trim(); })
    .filter(function(t){ return t.length > 0; });
}
function _listHasExactCaseFold_(arr, needle) {
  var n = String(needle || '').trim().toLowerCase();
  for (var i = 0; i < arr.length; i++) {
    if (String(arr[i] || '').trim().toLowerCase() === n) return true;
  }
  return false;
}

function _safeEmail_() {
  try { return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); }
  catch (e) { return ''; }
}

function _ensureReminderBridgeColumnsOnQueue_(sh) {
  // Build normalized header index (1-based)
  var lc = sh.getLastColumn();
  var hdrVals = sh.getRange(1, 1, 1, lc).getDisplayValues()[0];
  var normIdx = {};
  for (var i = 0; i < hdrVals.length; i++) {
    var n = _normHeader_(hdrVals[i]);
    if (!(n in normIdx)) normIdx[n] = i + 1; // 1-based column index
  }

  // Desired columns (normalized keys → canonical labels)
  var want = [
    { key: _normHeader_('Reminder Snooze Until'), label: 'Reminder Snooze Until' },
    { key: _normHeader_('Reminder ID'),          label: 'Reminder ID' }
  ];

  // Add any missing columns once, after current last column (idempotent)
  want.forEach(function(w) {
    if (!normIdx[w.key]) {
      sh.insertColumnAfter(lc);
      lc++;
      sh.getRange(1, lc).setValue(w.label);
      normIdx[w.key] = lc;
    }
  });

  // --- KEEP ONLY "Reminder ID" HIDDEN & PIN IT AT THE FAR RIGHT ---
  var idCol = normIdx[_normHeader_('Reminder ID')];
  var snoozeCol = normIdx[_normHeader_('Reminder Snooze Until')];

  // If the ID column is not the last used column, move it to the end once.
  if (idCol && idCol !== sh.getLastColumn()) {
    // Move the single "Reminder ID" column to the far right (last + 1)
    sh.moveColumns(sh.getRange(1, idCol, 1, 1), sh.getLastColumn() + 1);

    // Rebuild header index after the move
    lc = sh.getLastColumn();
    hdrVals = sh.getRange(1, 1, 1, lc).getDisplayValues()[0];
    normIdx = {};
    for (var j = 0; j < hdrVals.length; j++) normIdx[_normHeader_(hdrVals[j])] = j + 1;

    idCol = normIdx[_normHeader_('Reminder ID')];
    snoozeCol = normIdx[_normHeader_('Reminder Snooze Until')];
  }

  // Unhide everything to the right of the ID column (prevents cascading hide)
  var maxCols = sh.getMaxColumns();
  if (idCol && idCol < maxCols) sh.showColumns(idCol + 1, maxCols - idCol);

  // Normalize then hide exactly the one ID column.
  if (idCol) {
    sh.showColumns(idCol, 1); // ensure known state
    sh.hideColumns(idCol, 1); // hide exactly one col (no neighbor)
  }

  // QoL: make Snooze slightly wider and date-only (idempotent)
  if (snoozeCol) {
    sh.setColumnWidth(snoozeCol, 140);
    var lr2 = sh.getLastRow();
    if (lr2 >= 2) {
      var dv = SpreadsheetApp.newDataValidation()
        .requireDate()
        .setAllowInvalid(false)
        .setHelpText('Pick a date (yyyy-mm-dd).')
        .build();
      sh.getRange(2, snoozeCol, lr2 - 1, 1)
        .setDataValidation(dv)
        .setNumberFormat('yyyy-mm-dd');
    }
  }


}

// Robust header normalizer & finder (use everywhere we read headers)
function _normHeader_(s) {
  return String(s == null ? '' : s)
    .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, ' ') // zero-width, NBSP → space
    .replace(/\s+/g, ' ')                         // collapse spaces
    .trim()
    .toLowerCase();
}
function _hIdx_(headers, want, alts) {
  var target = _normHeader_(want);
  for (var i = 0; i < headers.length; i++) {
    if (_normHeader_(headers[i]) === target) return i;
  }
  if (alts && alts.length) {
    for (var j = 0; j < alts.length; j++) {
      var t = _normHeader_(alts[j]);
      for (var k = 0; k < headers.length; k++) {
        if (_normHeader_(headers[k]) === t) return k;
      }
    }
  }
  return -1;
}

// === NEW: Next DV appt lookup + friendly label builder =======================

/** Return the furthest-future Diamond Viewing DateTime for a RootApptID (or latest DV if all are past).
 * Looks at 00_Master Appointments with tolerant headers. Returns a Date or null.
 */
function _nextDiamondViewingDateForRoot_(root){
  if (!root) return null;
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) return null;

  // robust header read
  var hdr = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(function(s){ return String(s||'').trim(); });
  function idxOf(name, alts){
    var i = hdr.indexOf(name);
    if (i >= 0) return i;
    if (alts && alts.length) {
      for (var k=0;k<alts.length;k++){ var j = hdr.indexOf(alts[k]); if (j>=0) return j; }
    }
    // last resort: normalized match
    var want = _normHeader_(name);
    for (var a=0;a<hdr.length;a++){ if (_normHeader_(hdr[a]) === want) return a; }
    return -1;
  }

  var iRoot  = idxOf('RootApptID', ['Root Appt ID','Root Appointment ID','APPT_ID']);
  // visit type variations (use exact Diamond Viewing if present)
  var iType  = idxOf('Visit Type', ['Appt Type','Appointment Type','Type']);
  // date/time variations (prefer ISO if present)
  var iISO   = idxOf('ApptDateTime (ISO)', ['ApptDateTime','VisitDateTimeISO']);
  var iDate  = idxOf('Visit Date', ['Appt Date','Appointment Date','Date']);
  var iTime  = idxOf('Visit Time', ['Appt Time','Appointment Time','Time']);

  if (iRoot < 0) return null;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  var rng = sh.getRange(2, 1, lastRow-1, sh.getLastColumn()).getDisplayValues();
  var now = new Date();
  var dvs = [];

  for (var r=0; r<rng.length; r++){
    var row = rng[r];
    if (String(row[iRoot]||'').trim() !== String(root).trim()) continue;

    var vt = iType >=0 ? String(row[iType]||'').trim().toLowerCase() : '';
    if (!/diamond\s*viewing/i.test(vt)) continue;

    var dt = null;
    if (iISO >= 0 && row[iISO]) {
      var d = new Date(String(row[iISO]));
      if (!isNaN(d)) dt = d;
    }
    if (!dt) {
      var dStr = iDate >= 0 ? String(row[iDate]||'').trim() : '';
      var tStr = iTime >= 0 ? String(row[iTime]||'').trim() : '';
      if (dStr || tStr) {
        var d2 = new Date((dStr||'') + ' ' + (tStr||''));
        if (!isNaN(d2)) dt = d2;
      }
    }
    if (dt) dvs.push(dt);
  }

  if (!dvs.length) return null;

  // choose the most future DV (max date overall)
  dvs.sort(function(a,b){ return a.getTime() - b.getTime(); });
  var mostFuture = dvs[dvs.length - 1];

  // If that most-future date is still in the past (edge case), we still return it (caller handles phrasing).
  return mostFuture;
}

/** Build the friendly label (adds "(Appt in less than X days)" or "(Appt today)").
 * Keeps non-DV types as-is. For DV: uses NEXT DV from 00_ (furthest future).
 */
function _friendlyReminderLabel_(type, root){
  type = _normType_(type || '');
  var base = (TYPE_LABEL && TYPE_LABEL[type]) ? TYPE_LABEL[type] : type || 'Reminder';

  // Only DV types get the date suffix
  if (type !== 'DV_URGENT' && type !== 'DV_PROPOSE') return base;

  var dv = _nextDiamondViewingDateForRoot_(root);
  if (!dv) return base; // no info → just base

  var now = new Date();
  var MS = 24*60*60*1000;
  // ceil to keep “less than X days” feel
  var days = Math.ceil((dv.getTime() - now.getTime()) / MS);

  var suffix = '';
  if (days <= 0) {
    suffix = '(Appt today)';
  } else {
    suffix = '(Appt in less than ' + days + ' days)';
  }
  return base + ' ' + suffix;
}


/** Phase‑1 enrichment: join 04_Reminders_Queue items with 07_Root_Index. */
function _enrichRemindersFromIndex_(items) {
  try {
    if (!items || !items.length) return items;
    var ss = (typeof MASTER_SS_ === 'function') ? MASTER_SS_() : SpreadsheetApp.getActive();
    var s07 = (typeof getSheetOrThrow_ === 'function') ? getSheetOrThrow_('07_Root_Index') : ss.getSheetByName('07_Root_Index');
    if (!s07) { Logger.log('[enrich] 07_Root_Index missing — skip enrichment'); return items; }

    // Read objects from 07 using existing helper if available
    var rows = (typeof getObjects_ === 'function') ? getObjects_(s07) : (function(){
      var hdr = s07.getRange(1,1,1,s07.getLastColumn()).getDisplayValues()[0];
      var vals = s07.getRange(2,1,Math.max(0,s07.getLastRow()-1),hdr.length).getValues();
      return vals.map(function(v){ var o={}; for (var i=0;i<hdr.length;i++) o[hdr[i]]=v[i]; return o; });
    })();

    function norm(s){ return String(s||'').trim().toLowerCase(); }
    var byRoot = new Map();
    var bySO   = new Map();
    var byCust = new Map();

    rows.forEach(function(o){
      var root = norm(o['RootApptID'] || o['Root Appointment ID'] || o['Root Appt ID']);
      var so   = norm(o['SO Number'] || o['SO#'] || o['SO No'] || o['SO']);
      var cust = norm(o['Customer Name'] || o['Customer'] || o['Client Name']);
      if (root) byRoot.set(root, o);
      if (so)   bySO.set(so, o);
      if (cust && !byCust.has(cust)) byCust.set(cust, o); // first wins
    });

    var MS = 24*60*60*1000;

    items.forEach(function(it){
      var row = null;
      var rootKey = norm(it.root || '');
      var soKey   = norm(it.soNumber || '');
      var custKey = norm(it.customer || '');

      if (rootKey && byRoot.has(rootKey)) row = byRoot.get(rootKey);
      else if (soKey && bySO.has(soKey))  row = bySO.get(soKey);
      else if (custKey && byCust.has(custKey)) row = byCust.get(custKey);

      if (!row) return;

      // Map fields → reminder item
      it.root             = String(row['RootApptID'] || row['Root Appointment ID'] || row['Root Appt ID'] || it.root || '');
      it.salesStage       = row['Sales Stage'] || it.salesStage || '';
      it.conversionStatus = row['Conversion Status'] || it.conversionStatus || '';
      it.cos              = row['Custom Order Status'] || row['Custom Order Status (at log)'] || it.cos || '';
      it.inProd           = row['In Production Status'] || it.inProd || '';
      it.csos             = row['Center Stone Order Status'] || row['Center Stone Order Status (at log)'] || it.csos || '';
      it.updatedBy        = row['Updated By'] || row['Last Updated By'] || it.updatedBy || '';
      it.updatedAt        = _toDateSafe_(row['Updated At'] || row['Last Updated At'] || row['Last Updated At (at log)'] || it.updatedAt || '');
      var ds = Number(row['Days Since Last Update'] || row['Days Since Last Update (calc)'] || '');
      if (!isNaN(ds) && isFinite(ds)) it.daysSinceUpdate = ds;
      if ((it.daysSinceUpdate == null || it.daysSinceUpdate === '') && it.updatedAt instanceof Date) {
        it.daysSinceUpdate = Math.max(0, Math.floor((new Date().getTime() - it.updatedAt.getTime())/MS));
      }
      it.csrUrl           = row['Client Status Report URL'] || row['CSR URL'] || it.csrUrl || '';

      // Compose Next Steps rich text (revised):
      //   [Reminder: <friendly label>] — [Note: ...]
      // We intentionally DROP the leading COS/Sales Stage text to avoid repeating “3D Waiting Approval…”.
      var noteStr = String(it.note || '').trim();

      // NEW: always compute a friendly label (DV adds "(Appt ...)")
      var label = _friendlyReminderLabel_(it.type, it.root);

      var parts = [];
      // Put the reminder label first, clearly
      parts.push('Reminder: ' + label);

      // Keep your original note if present
      if (noteStr) parts.push('Note: ' + noteStr);

      it.nextStepsRich = parts.join(' — ');
    });

    return items;
  } catch (e) {
    Logger.log('[enrich] failed: ' + e);
    return items;
  }
}

function _normType_(t){
  t = String(t||'').toUpperCase().trim();
  // normalize separators
  t = t.replace(/[\s\-]+/g, '_');
  if (t.indexOf('DV_PROPOSE') === 0) return 'DV_PROPOSE';
  if (t.indexOf('DV_URGENT')  === 0) return 'DV_URGENT';
  if (t === 'FOLLOWUP' || t === 'FOLLOW_UP') return 'FOLLOW_UP';  // 👈 canonical
  return t;
}


function dbg__reminders_probe() {
  var rep = (typeof detectRepName_ === 'function') ? detectRepName_() : '(none)';
  var email = (function(){ try { return Session.getActiveUser().getEmail(); } catch(_) { return ''; } })();
  Logger.log('[probe] rep="%s" email="%s"', rep, email);

  var items = _readDueRemindersForRep_(rep);
  Logger.log('[probe] due items count=%s', items.length);
  items.slice(0, 5).forEach(function(it, i){
    Logger.log('#%s id=%s type=%s due=%s status=%s assigned="%s" customer="%s"',
      i+1, it.id, it.type, it.nextDueAt, it.status, (it.assigned || ''), it.customer || '');
  });
}
