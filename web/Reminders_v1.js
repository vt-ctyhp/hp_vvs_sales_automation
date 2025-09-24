/** 
 * 01 - Reminders_v1.gs (v3.0)
 * Engine for daily Google Chat reminders for Start 3D / Assign SO / 3D Revision / Follow-up.
 * Tailored for the '00_Master Appointments' tab and your R1 mapping (Rep Directory + Email helper columns).
 *
 * Safe-by-default: column-by-name resolution, idempotent queue, batched Chat posts, daily 9:30 schedule.
 */

/**
 * Dump all queue rows that match a given customer (case/space/NBSP tolerant).
 * Shows: row#, id, type, status, SO, customer.
 */
function remind__dumpQueueForCustomer(customerName){
  // inline normalizers (same behavior as in Reminders_v1)
  function normWS(s){ return String(s||'').replace(/\u00A0/g,' ').replace(/\s+/g,' ').trim(); }
  function normSO(s){ return normWS(s).replace(/^'+/, ''); }
  function eqCI(a,b){ return normWS(a).toLowerCase() === normWS(b).toLowerCase(); }

  const sh = SpreadsheetApp.getActive().getSheetByName('04_Reminders_Queue');
  if (!sh){ console.log('Queue sheet not found'); return; }

  const rg = sh.getDataRange().getDisplayValues();
  const headers = rg[0] || [];
  const rows = rg.slice(1);

  // simple 1-based header indexer
  const H = {};
  headers.forEach((h,i)=>{ const k=String(h||'').trim().toLowerCase(); if (k) H[k]=i+1; });

  const cSO   = H['so number'] || H['so#'] || H['sonumber'] || H['so'] || H['so_number'] || H['sonum'] || H['so number'];
  const cType = H['type'];
  const cSt   = H['status'];
  const cCust = H['customername'] || H['customer name'] || H['customer'];
  const cId   = H['id'];

  if (!cCust){ console.log('customerName column not found in queue header.'); return; }

  const key = normWS(customerName);
  let found = 0;
  for (let i=0;i<rows.length;i++){
    const r = rows[i];
    const rCust = normWS(r[cCust-1]);
    if (!eqCI(rCust, key)) continue;

    const rowIdx = i+2;
    const so     = cSO   ? normSO(r[cSO-1]) : '';
    const type   = cType ? r[cType-1] : '';
    const st     = cSt   ? r[cSt-1]   : '';
    const id     = cId   ? r[cId-1]   : '';
    console.log('#'+rowIdx, 'id='+id, 'type='+type, 'status='+st, 'so='+so, 'cust='+rCust);
    found++;
  }
  if (!found) console.log('No rows for customer:', customerName);
}

const REMIND = Object.freeze({
  ORDERS_SHEET_NAME: '00_Master Appointments',
  QUEUE_SHEET_NAME:  '04_Reminders_Queue',
  LOG_SHEET_NAME:    '15_Reminders_Log',
  REP_DIR_SHEET:     '05_RepDirectory',

  TIMEZONE:          'America/Los_Angeles',
  DAILY_HOUR:        9,
  DAILY_MINUTE:      30,

  // Script Properties keys
  PROP_TEAM_WEBHOOK:    'TEAM_CHAT_WEBHOOK',
  PROP_MANAGER_WEBHOOK: 'MANAGER_CHAT_WEBHOOK',
  PROP_DAILY_SENT_FOR:  'DAILY_SENT_FOR',    // YYYY-MM-DD string of last successful daily send

  // Reminder types
  TYPE_START3D:   'START3D',
  TYPE_ASSIGNSO:  'ASSIGNSO',
  TYPE_REV3D:     'REV3D',
  TYPE_FOLLOWUP:  'FOLLOWUP',

  // Queue statuses (internal only; does NOT alter business statuses)
  ST_PENDING:   'PENDING',
  ST_CONFIRMED: 'CONFIRMED',
  ST_CANCELLED: 'CANCELLED',
  ST_SNOOZED:   'SNOOZED',

  // Labels in Orders sheet (resolved by name, not index)
  COL_SO:                  'SO#',
  COL_CUSTOMER_NAME:       'Customer Name',
  COL_ASSIGNED_REP_NAME:   'Assigned Rep',
  COL_ASSISTED_REP_NAME:   'Assisted Rep',
  COL_ASSIGNED_REP_EMAIL:  'Assigned Rep Email',  // helper column added by setup
  COL_ASSISTED_REP_EMAIL:  'Assisted Rep Email',  // helper column added by setup
  COL_SALES_STAGE:         'Sales Stage',         // detect Follow-Up Required
  COL_CUSTOM_ORDER_STATUS: 'Custom Order Status', // detect 3D Received

  TYPE_COS: 'COS',  // unified "Custom Order Status Reminder"
  COS_3D_PENDING: [
    '3D Requested',
    '3D Revision Requested',
    '3D Waiting Approval',
    'Waiting Production Timeline',
    'Final Photos - Waiting Approval',
    'In US Store'
  ],
});

  /** Pretty-print a local timestamp 'YYYY-MM-DD HH:mm[:ss]' as 'Wed, Sep 3 at 9:30AM' (no TZ conversion). */
  function remind__prettyLocalTimestamp_(tsStr) {
    const s = String(tsStr || '').trim();
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (!m) return '';
    const y = +m[1], mo = +m[2]-1, d = +m[3], hh24 = +m[4], mm = +m[5];
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const days   = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    // Day-of-week is calendar-based (safe across TZ)
    const dow = new Date(y, mo, d).getDay();
    const ampm = hh24 >= 12 ? 'PM' : 'AM';
    const hh12 = ((hh24 + 11) % 12) + 1;
    const mm2  = (mm < 10 ? '0' : '') + mm;
    return `${days[dow]}, ${months[mo]} ${d} at ${hh12}:${mm2}${ampm}`;
  }

  /** Pretty-print a date 'YYYY-MM-DD' at 9:30AM PT as 'Wed, Sep 3 at 9:30AM' (no TZ conversion). */
  function remind__pretty930ForIsoDate_(isoDate) {
    const s = String(isoDate || '').trim();
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return '';
    const y = +m[1], mo = +m[2]-1, d = +m[3];
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const days   = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    const dow = new Date(y, mo, d).getDay();
    return `${days[dow]}, ${months[mo]} ${d} at 9:30AM`;
  }

// ---------- SO helpers (canonical key + pretty display) ----------
/**
 * Canonical key used for all comparisons / indexing.
 * Accepts '00.1293', "  SO#00.1293", 0.1293 (display), '001293', etc.
 * Returns 6 digits '001293' or '' if it can't safely parse.
 */
function _soKey_(raw) {
  let s = String(raw == null ? '' : raw).trim();
  if (!s) return '';
  s = s.replace(/^'+/, '');          // leading apostrophe from Sheets
  s = s.replace(/^\s*SO#?/i, '');    // optional "SO" label
  s = s.replace(/\s|\u00A0/g, '');   // spaces & NBSP
  const digits = s.replace(/\D/g, ''); // keep only digits
  if (!digits) return '';
  // If fewer than 6 digits, left‑pad (common case: '0.1293' -> '001293')
  // If more than 6, keep the last 6 (defensive; avoids harming unrelated prefixes)
  const d = digits.length < 6 ? digits.padStart(6,'0') : digits.slice(-6);
  return d;
}

/** Human display form ##.#### from a key or any raw value. */
function _soPretty_(raw) {
  const k = _soKey_(raw);
  return k ? (k.slice(0,2) + '.' + k.slice(2)) : '';
}

/** ── Back‑compat (Sept 2025): older code calls `_canonSO_()`; new code uses `_soKey_()`. */
function _canonSO_(raw) { 
  return _soKey_(raw);      // same canonical key your newer code uses
}

/** Equality on SOs (uses the canonical key). */
function _eqSO_(a, b) {
  const ka = _soKey_(a), kb = _soKey_(b);
  return !!ka && ka === kb;
}


/** Public API object. Keep methods as plain functions to be callable from triggers & other files. */
const Remind = (function() {

  // --- DEBUG helper ---------------------------------------------------------
  const __DBG_ON = true; // flip to false to mute all extra logs
  function _dbg() {
    if (!__DBG_ON) return;
    try { console.log.apply(console, arguments); } catch (e) {}
  }
  // --------------------------------------------------------------------------

  // ---------- Utilities ----------
  function _props() { return PropertiesService.getScriptProperties(); }

  function _now() { return new Date(); }

  // ---------- PT (America/Los_Angeles) time helpers ----------

  // Format a Date as 'yyyy-MM-dd HH:mm:ss' in PT.
  function _fmtPT_(d) {
    return Utilities.formatDate(d, REMIND.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  }

  // Return 'yyyy-MM-dd' for "today" in PT.
  function _todayIsoPT_() {
    return Utilities.formatDate(_now(), REMIND.TIMEZONE, 'yyyy-MM-dd');
  }

  // Return 'yyyy-MM-dd 09:30:00' in PT for a given date (Date | ISO | ms).
  function _pt930Str_(dateOrIso) {
    let iso;
    if (dateOrIso instanceof Date) {
      iso = Utilities.formatDate(dateOrIso, REMIND.TIMEZONE, 'yyyy-MM-dd');
    } else if (typeof dateOrIso === 'number' && isFinite(dateOrIso)) {
      iso = Utilities.formatDate(new Date(dateOrIso), REMIND.TIMEZONE, 'yyyy-MM-dd');
    } else {
      const s = String(dateOrIso || '').slice(0, 10);
      iso = /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : _todayIsoPT_();
    }
    return iso + ' 09:30:00';
  }

  // Return "now + addMinutes" formatted in PT as 'yyyy-MM-dd HH:mm:ss'.
  function _nowPTStr_(addMinutes) {
    const ms = (addMinutes || 0) * 60 * 1000;
    return _fmtPT_(new Date(_now().getTime() + ms));
  }

  // Canonicalize 'YYYY-MM-DD H:mm[:ss]' -> 'YYYY-MM-DD HH:mm:ss' for safe string comparison.
  function _canonPTDateTime_(s) {
    const t = String(s || '').trim();

    // 1) Canonical: YYYY-MM-DD HH:mm[:ss]
    let m = t.match(/^(\d{4}-\d{2}-\d{2})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (m) {
      const HH = ('0' + m[2]).slice(-2);
      const SS = m[4] || '00';
      return `${m[1]} ${HH}:${m[3]}:${SS}`;
    }

    // 2) Common Sheets display: M/D/YYYY h:mm AM/PM
    m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})\s*([AP]M)$/i);
    if (m) {
      let [, mo, d, y, h, mm, ap] = m;
      let H = parseInt(h, 10);
      if (/pm/i.test(ap) && H < 12) H += 12;
      if (/am/i.test(ap) && H === 12) H = 0;
      const iso = `${y}-${('0'+mo).slice(-2)}-${('0'+d).slice(-2)}`;
      return `${iso} ${('0'+H).slice(-2)}:${mm}:00`;
    }

    // 3) M/D/YYYY h:mm (24h)
    m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})$/);
    if (m) {
      const [, mo, d, y, H, mm] = m;
      const iso = `${y}-${('0'+mo).slice(-2)}-${('0'+d).slice(-2)}`;
      return `${iso} ${('0'+H).slice(-2)}:${mm}:00`;
    }

    // 4) M/D/YYYY (date-only → assume 09:00)
    m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) {
      const [, mo, d, y] = m;
      const iso = `${y}-${('0'+mo).slice(-2)}-${('0'+d).slice(-2)}`;
      return `${iso} 09:00:00`;
    }

    return '';
  }


  function _tzDate(y, m, d, hh, mm) {
    const tz = REMIND.TIMEZONE;
    // Construct in UTC then convert to TZ to avoid DST issues
    const dt = new Date(Date.UTC(y, m, d, hh, mm, 0, 0));
    return new Date(Utilities.formatDate(dt, tz, 'yyyy/MM/dd HH:mm:ss'));
  }

  function _todayInTZ() {
    const tz = REMIND.TIMEZONE;
    const now = _now();
    const y = Number(Utilities.formatDate(now, tz, 'yyyy'));
    const m = Number(Utilities.formatDate(now, tz, 'M')) - 1;
    const d = Number(Utilities.formatDate(now, tz, 'd'));
    return { y, m, d };
  }

  function _todayAt930() {
    const { y, m, d } = _todayInTZ();
    return _tzDate(y, m, d, REMIND.DAILY_HOUR, REMIND.DAILY_MINUTE);
    // Always 9:30 in PT
  }

  // Pretty "MM/DD at Sat, h:mm AM/PM" from Date/parseable
  function _prettyAppt_(raw) {
    if (!raw) return '';
    const tz = REMIND.TIMEZONE;
    const d  = (raw instanceof Date) ? raw : new Date(String(raw));
    if (isNaN(d)) return '';
    const mmdd = Utilities.formatDate(d, tz, 'MM/dd');
    const dow  = Utilities.formatDate(d, tz, 'EEE');
    const t    = Utilities.formatDate(d, tz, 'h:mm a');
    return `${mmdd} at ${dow}, ${t}`;
  }

  // Build a Date from separate Visit Date + Visit Time cells, then pretty-print.
  // From Visit Date + Visit Time cells → formatted string
  function _prettyApptFromVisitCells_(dateCell, timeCell) {
    const tz = REMIND.TIMEZONE;
    if (!dateCell) return '';

    let y, m, d;
    if (dateCell instanceof Date && !isNaN(dateCell)) {
      y = dateCell.getFullYear(); m = dateCell.getMonth(); d = dateCell.getDate();
    } else {
      const s = String(dateCell||'').trim();
      const mYMD = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      const mMDY = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
      if (mYMD) { y=+mYMD[1]; m=+mYMD[2]-1; d=+mYMD[3]; }
      else if (mMDY) { y=+mMDY[3]; m=+mMDY[1]-1; d=+mMDY[2]; }
      else return '';
    }

    let hh=null, mm=null;
    if (timeCell) {
      if (timeCell instanceof Date && !isNaN(timeCell)) {
        hh = Number(Utilities.formatDate(timeCell, tz, 'H'));
        mm = Number(Utilities.formatDate(timeCell, tz, 'm'));
      } else {
        const ts = String(timeCell||'').trim();
        const mt = ts.match(/^(\d{1,2})(?::(\d{2}))?\s*([ap]m)?$/i);
        if (mt) {
          hh = +mt[1]; mm = mt[2]?+mt[2]:0;
          const ap = (mt[3]||'').toLowerCase();
          if (ap === 'pm' && hh < 12) hh += 12;
          if (ap === 'am' && hh === 12) hh = 0;
        }
      }
    }
    const base = _tzDate(y, m, d, (hh==null?9:hh), (mm==null?0:mm));
    return _prettyAppt_(base);
  }


  function _dateOnlyStr(date) {
    return Utilities.formatDate(date, REMIND.TIMEZONE, 'yyyy-MM-dd');
  }

  function _dateDiffDays(a, b) {
    const MS = 24*60*60*1000;
    return Math.floor((a.getTime() - b.getTime()) / MS);
  }

  function _normalizeWS(val) {
    // Replace NBSP with space, collapse whitespace, trim
    return String(val || '').replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim();
  }
  function _normalize(val) {
    return _normalizeWS(val); // keep name for all existing callers
  }
  function _normalizeCI(val) {
    return _normalizeWS(val).toLowerCase();
  }
  function _equalsCI(a, b) {
    return _normalizeCI(a) === _normalizeCI(b);
  }


  function _uuid() {
    return Utilities.getUuid();
  }

  function _getSpreadsheet() {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  function _getSheetByName(name, createIfMissing=false) {
    const ss = _getSpreadsheet();
    let sh = ss.getSheetByName(name);
    if (!sh && createIfMissing) {
      sh = ss.insertSheet(name);
    }
    return sh;
  }

  function _ensureHeaders(sh, headers) {
    const range = sh.getRange(1,1,1,headers.length);
    const values = range.getValues()[0];
    let changed = false;
    headers.forEach((h, i) => {
      if (_normalize(values[i]) !== h) {
        values[i] = h;
        changed = true;
      }
    });
    if (changed) {
      range.setValues([values]);
    }
  }

  function _findHeaderIndexes(headerRow) {
    const map = {};
    headerRow.forEach((h, i) => {
      map[_normalize(h).toLowerCase()] = i+1; // 1-based
    });
    return {
      col: (label) => {
        const key = _normalize(label).toLowerCase();
        const idx = map[key];
        if (!idx) throw new Error('Column not found by header: ' + label);
        return idx;
      },
      tryCol: (label) => map[_normalize(label).toLowerCase()] || null
    };
  }

  function _colLetter(n) {
    let s = '';
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = (n - m - 1) / 26;
    }
    return s;
  }

  function _getGid(sh) { return sh.getSheetId(); }

  function _ordersData() {
    const sh = _getSheetByName(REMIND.ORDERS_SHEET_NAME);
    const vr = sh.getDataRange().getDisplayValues();
    const headers = vr[0];
    const idx = _findHeaderIndexes(headers);

    const soCol          = idx.col(REMIND.COL_SO);
    const custCol        = idx.col(REMIND.COL_CUSTOMER_NAME);
    const aRepNameCol    = idx.col(REMIND.COL_ASSIGNED_REP_NAME);
    const asRepNameCol   = idx.tryCol(REMIND.COL_ASSISTED_REP_NAME);
    const aRepEmailCol   = idx.tryCol(REMIND.COL_ASSIGNED_REP_EMAIL);
    const asRepEmailCol  = idx.tryCol(REMIND.COL_ASSISTED_REP_EMAIL);
    const salesStageCol  = idx.col(REMIND.COL_SALES_STAGE);
    const cosCol         = idx.col(REMIND.COL_CUSTOM_ORDER_STATUS);
    const nextStepsCol   = idx.tryCol('Next Steps');

    // NEW: try common headers for Center Stone Order Status (CSOS)
    const csosCol = idx.tryCol('Center Stone Order Status') ||
                    idx.tryCol('Center Stone Status') ||
                    idx.tryCol('CSOS') ||
                    idx.tryCol('Diamond Memo Status') ||
                    idx.tryCol('DV Status');

    // NEW: RootAppt + Visit Date/Time (plus optional unified ISO)
    const rootApptCol = idx.tryCol('RootApptID') || idx.tryCol('RootApptId') ||
                        idx.tryCol('APPT_ID')    || idx.tryCol('ApptID')     ||
                        idx.tryCol('APPT ID')    || null;

    const visitDateCol = idx.tryCol('Visit Date') || idx.tryCol('Appt Date') ||
                        idx.tryCol('Appointment Date') || idx.tryCol('Date') || null;

    const visitTimeCol = idx.tryCol('Visit Time') || idx.tryCol('Appt Time') ||
                        idx.tryCol('Appointment Time') || idx.tryCol('Time') || null;

    const apptIsoCol  = idx.tryCol('ApptDateTimeISO')           ||
                        idx.tryCol('ApptDateTime (ISO)')        ||
                        idx.tryCol('Appointment Date/Time ISO') ||
                        idx.tryCol('Appointment DateTime (ISO)')||
                        idx.tryCol('Appt Start ISO')            ||
                        idx.tryCol('Appt Start (ISO)')          || null;

    const rows = vr.slice(1);
    return {
      sh, headers, idx,
      soCol, custCol, aRepNameCol, asRepNameCol, aRepEmailCol, asRepEmailCol,
      salesStageCol, cosCol, nextStepsCol,
      csosCol, rootApptCol, visitDateCol, visitTimeCol, apptIsoCol,   // ← NEW
      rows, firstDataRow: 2
    };

  }
  // Normalize SO for matching & storage: trim + strip any leading apostrophe from Sheets
  function _normalizeSO(val) {
    return String(val || '').replace(/^'+/, '').trim();
  }

  // Inside the Remind IIFE — replace these two with the key-based versions
  function _buildSoIndex(data) {
    const map = new Map();
    data.rows.forEach((row, i) => {
      const key = _soKey_(row[data.soCol - 1]);
      if (key) map.set(key, data.firstDataRow + i);
    });
    return map;
  }

  function _ordersRowBySo(so, data, soIndex) {
    const r = soIndex.get(_soKey_(so));
    if (!r) return null;
    const rowVals = data.sh.getRange(r, 1, 1, data.headers.length).getDisplayValues()[0];
    return { rowNumber: r, rowVals, gid: _getGid(data.sh) };
  }

  function _buildRootIndex(data) {
    const map = new Map();
    if (!data.rootApptCol) return map;
    data.rows.forEach((row, i) => {
      const raw = row[data.rootApptCol - 1];
      const key = _normalize(raw);
      if (key) map.set(key, data.firstDataRow + i);
    });
    return map;
  }

  function _ordersRowByRoot(rootId, data, rootIndex) {
    if (!rootId || !data.rootApptCol) return null;
    const r = rootIndex.get(_normalize(rootId));
    if (!r) return null;
    const rowVals = data.sh.getRange(r, 1, 1, data.headers.length).getDisplayValues()[0];
    return { rowNumber: r, rowVals, gid: _getGid(data.sh) };
  }


  function _sheetUrlForRow(gid, row) {
    const ss = _getSpreadsheet();
    return 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + gid + '&range=' + row + ':' + row;
  }


  function _is3DPendingStatus_(val) {
    const s = String(val || '').trim().toLowerCase();
    return REMIND.COS_3D_PENDING.some(x => s === String(x || '').trim().toLowerCase());
  }


  // ---------- Queue & Log ----------
  const QHEAD = [
    'id','soNumber','type','firstDueDate','nextDueAt','recurrence','status','snoozeUntil',
    'assignedRepName','assignedRepEmail','assistedRepName','assistedRepEmail',
    'customerName','nextSteps',
    'createdAt','createdBy','confirmedAt','confirmedBy','lastSentAt','attempts','lastError',
    'cancelReason','lastAdminAction','lastAdminBy'  // ← NEW
  ];

  const LHEAD = ['ts','id','soNumber','type','action','by','note'];

  function _queueSheet() {
    const sh = _getSheetByName(REMIND.QUEUE_SHEET_NAME, true);
    _ensureHeaders(sh, QHEAD);
    return sh;
  }
  function _logSheet() {
    const sh = _getSheetByName(REMIND.LOG_SHEET_NAME, true);
    _ensureHeaders(sh, LHEAD);
    return sh;
  }

  function _log(id, so, type, action, by, note) {
    const sh = _logSheet();
    sh.appendRow([
      Utilities.formatDate(_now(), REMIND.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
      id || '', so || '', type || '', action || '', by || '', note || ''
    ]);
  }

  function _readQueue() {
    const sh = _queueSheet();
    const rg = sh.getDataRange().getDisplayValues();
    const headers = rg[0];
    const rows = rg.slice(1);
    const idx = _findHeaderIndexes(headers);
    return { sh, headers, rows, idx };
  }

  function _queueUpsert_(so, type, firstDueDate, createdBy, opts) {
    const q = _readQueue();
    const idCol = q.idx.col('id');
    const soCol = q.idx.col('soNumber');
    const typeCol = q.idx.col('type');
    const statusCol = q.idx.col('status');

    // Deduplicate logic:
    // - COS: if SO present → match by SO; if SO blank and customer present → match by customerName
    // - Legacy 3D (START3D/ASSIGNSO/REV3D): treat them as one group (back-compat)
    const similar3D = new Set([REMIND.TYPE_START3D, REMIND.TYPE_ASSIGNSO, REMIND.TYPE_REV3D]);
    let matchRow = null;

    const wantSOKey = _soKey_(so);
    const wantCust = _normalize(opts && opts.customerName);

    const custColIdx = (q.idx.tryCol ? q.idx.tryCol('customerName') : null) || q.idx.col('customerName');

    for (let i = 0; i < q.rows.length; i++) {
      const r = q.rows[i];
      const rStatus = _normalize(r[statusCol-1]);
      if (rStatus === REMIND.ST_CANCELLED || rStatus === REMIND.ST_CONFIRMED) continue;

      const rSoKey = _soKey_(r[soCol-1]);
      const rType  = _normalize(r[typeCol-1]);

      if (type === REMIND.TYPE_COS) {
        const rCust = _normalize(r[custColIdx-1]);
        if (wantSOKey && rSoKey && rType === REMIND.TYPE_COS && rSoKey === wantSOKey) { matchRow = i+2; break; }
        if (!wantSOKey && wantCust && rType === REMIND.TYPE_COS && _equalsCI(rCust, wantCust)) { matchRow = i+2; break; }
      } else if (similar3D.has(type)) {
        if (wantSOKey && rSoKey && similar3D.has(rType) && rSoKey === wantSOKey) { matchRow = i+2; break; }
      } else {
        if (wantSOKey && rSoKey && rType === type && rSoKey === wantSOKey) { matchRow = i+2; break; }
      }
    }

    const tzNow = _now();
    const nextDueAtStr = _pt930Str_(firstDueDate); // always 9:30
    const payload = {
      id: _uuid(),
      soNumber: _soPretty_(so),  // <— pretty, standardized
      type: type,
      firstDueDate: _dateOnlyStr(firstDueDate),
      nextDueAt: nextDueAtStr,
      recurrence: 'DAILY',
      status: REMIND.ST_PENDING,
      snoozeUntil: '',
      cancelReason: '',
      lastAdminAction: '',
      lastAdminBy: '',
      assignedRepName: _normalize(opts?.assignedRepName),
      assignedRepEmail: _normalize(opts?.assignedRepEmail),
      assistedRepName: _normalize(opts?.assistedRepName),
      assistedRepEmail: _normalize(opts?.assistedRepEmail),
      customerName: _normalize(opts?.customerName),
      nextSteps: _normalize(opts?.nextSteps),
      createdAt: Utilities.formatDate(tzNow, REMIND.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
      createdBy: _normalize(createdBy),
      confirmedAt: '',
      confirmedBy: '',
      lastSentAt: '',
      attempts: '0',
      lastError: ''
    };

    if (matchRow) {
      const sh = q.sh;
      const id = sh.getRange(matchRow, idCol).getValue();
      sh.getRange(matchRow, q.idx.col('nextDueAt')).setValue(payload.nextDueAt);
      sh.getRange(matchRow, q.idx.col('firstDueDate')).setValue(payload.firstDueDate);
      sh.getRange(matchRow, q.idx.col('status')).setValue(REMIND.ST_PENDING);
      if (payload.assignedRepName)  sh.getRange(matchRow, q.idx.col('assignedRepName')).setValue(payload.assignedRepName);
      if (payload.assignedRepEmail) sh.getRange(matchRow, q.idx.col('assignedRepEmail')).setValue(payload.assignedRepEmail);
      if (payload.assistedRepName)  sh.getRange(matchRow, q.idx.col('assistedRepName')).setValue(payload.assistedRepName);
      if (payload.assistedRepEmail) sh.getRange(matchRow, q.idx.col('assistedRepEmail')).setValue(payload.assistedRepEmail);
      if (payload.customerName)     sh.getRange(matchRow, q.idx.col('customerName')).setValue(payload.customerName);
      if (payload.nextSteps)        sh.getRange(matchRow, q.idx.col('nextSteps')).setValue(payload.nextSteps);

      // NEW: keep the visible SO in the standardized ##.#### format
      sh.getRange(matchRow, q.idx.col('soNumber')).setValue(payload.soNumber);
      _writePrettyRow_(q, matchRow, payload);                 

      _log(id, so, type, 'ENQUEUED', createdBy, 'Upserted (refreshed due date)');
      return id;
    } else {
      // Append
      q.sh.appendRow(QHEAD.map(h => payload[h]));
      const q2 = _readQueue();
      _writePrettyRow_(q2, q2.sh.getLastRow(), payload);

      _log(payload.id, so, type, 'ENQUEUED', createdBy, 'New');
      return payload.id;
    }
  } 

  function _setQueueStatusFor(so, types, newStatus, note, by) {
    const q = _readQueue();
    const soCol = q.idx.col('soNumber');
    const typeCol = q.idx.col('type');
    const statusCol = q.idx.col('status');

    const nowStr = Utilities.formatDate(_now(), REMIND.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    for (let i=0; i<q.rows.length; i++) {
      const r = q.rows[i];
      const rSo = _normalizeSO(r[soCol-1]);
      const rType = _normalize(r[typeCol-1]);
      const rStatus = _normalize(r[statusCol-1]);
     if (!_eqSO_(r[soCol-1], so)) continue;
      if (types.length && !types.includes(rType)) continue;
      if (rStatus === newStatus) continue;
      const rowIdx = i+2;
      q.sh.getRange(rowIdx, statusCol).setValue(newStatus);
      if (newStatus === REMIND.ST_CONFIRMED) {
        q.sh.getRange(rowIdx, q.idx.col('confirmedAt')).setValue(nowStr);
        q.sh.getRange(rowIdx, q.idx.col('confirmedBy')).setValue(by || '');
      }
      _log(q.sh.getRange(rowIdx, q.idx.col('id')).getValue(), so, rType, newStatus, by || '', note || '');
    }
  }

  function _at930(dateOnly) {
    // dateOnly is Date or string 'YYYY-MM-DD'
    let d = dateOnly;
    if (!(d instanceof Date)) d = new Date(dateOnly + 'T00:00:00');
    const y = d.getFullYear(), m = d.getMonth(), day = d.getDate();
    return _tzDate(y, m, day, REMIND.DAILY_HOUR, REMIND.DAILY_MINUTE);
  }

  function _addDays(dateOnly, add) {
    const d = new Date(dateOnly);
    d.setDate(d.getDate() + add);
    return d;
  }

  function _closeActiveCosFor_(so, customerName, note, by) {
    const q   = _readQueue();
    const soC = q.idx.col('soNumber');
    const tyC = q.idx.col('type');
    const stC = q.idx.col('status');
    const idC = q.idx.col('id');

    const custC = (q.idx.tryCol ? q.idx.tryCol('customerName') : null) || q.idx.col('customerName');

    const targetSO   = _normalizeSO(so);
    const targetCust = _normalize(customerName);

    const typesToClose = new Set([REMIND.TYPE_COS, REMIND.TYPE_START3D, REMIND.TYPE_ASSIGNSO, REMIND.TYPE_REV3D]);
    const nowStr = Utilities.formatDate(_now(), REMIND.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    for (let i = 0; i < q.rows.length; i++) {
      const r = q.rows[i];
      const rStatus = _normalize(r[stC-1]);
      if (rStatus === REMIND.ST_CONFIRMED || rStatus === REMIND.ST_CANCELLED) continue;

      const rType = _normalize(r[tyC-1]);
      if (!typesToClose.has(rType)) continue;

      const rSO   = _normalizeSO(r[soC-1]);
      const rCust = _normalize(r[custC-1]);

      // ✅ Close if:
      // 1) SO matches, OR
      // 2) SO is blank AND customer matches (only when we know the customer)
      const hit =
        (targetSO && _equalsCI(rSO, targetSO)) ||
        (targetSO && !rSO && targetCust && _equalsCI(rCust, targetCust)) ||
        (!targetSO && targetCust && _equalsCI(rCust, targetCust));

      if (!hit) continue;

      const rowIdx = i + 2;
      q.sh.getRange(rowIdx, stC).setValue(REMIND.ST_CONFIRMED);
      q.sh.getRange(rowIdx, q.idx.col('confirmedAt')).setValue(nowStr);
      q.sh.getRange(rowIdx, q.idx.col('confirmedBy')).setValue(by || '');
      _log(q.sh.getRange(rowIdx, idC).getValue(), rSO, rType, REMIND.ST_CONFIRMED, by || '', note || 'Closed (restart)');
    }
  }


  // ---------- Public enqueue APIs (lightweight hooks from existing code) ----------
  function scheduleStart3D(soNumber) {
    const firstDue = _addDays(_todayAt930(), 2); // +2 days
    return _queueUpsert_(soNumber, REMIND.TYPE_START3D, firstDue, _currentEditor_());
  }
  function scheduleAssignSO(soNumber) {
    const firstDue = _addDays(_todayAt930(), 2);
    return _queueUpsert_(soNumber, REMIND.TYPE_ASSIGNSO, firstDue, _currentEditor_());
  }
  function schedule3DRevision(soNumber) {
    const firstDue = _addDays(_todayAt930(), 2);
    return _queueUpsert_(soNumber, REMIND.TYPE_REV3D, firstDue, _currentEditor_());
  }

  function ensureFollowUp(soNumber, opts) {
    let next = _todayAt930();
    const now = _now();
    if (now.getTime() >= next.getTime()) next = _addDays(next, 1);
    return _queueUpsert_(soNumber, REMIND.TYPE_FOLLOWUP, next, _currentEditor_(), {
      assignedRepName:  opts && opts.assignedRepName,
      assistedRepName:  opts && opts.assistedRepName,
      assignedRepEmail: opts && opts.assignedRepEmail,
      assistedRepEmail: opts && opts.assistedRepEmail,
      customerName:     opts && opts.customerName,
      nextSteps:        opts && opts.nextSteps
    });
  }

  function onClientStatusChange(soNumber, newSalesStage, newCustomOrderStatus, editorEmail, opts) {
    // COS: Create/keep while status remains in the pending set; confirm when it leaves
    if (_is3DPendingStatus_(newCustomOrderStatus)) {
      // If the status is "3D Revision Requested", start a new cycle and end the previous.
      const restart = _equalsCI(newCustomOrderStatus, '3D Revision Requested');
      scheduleCOS(soNumber, opts || {}, restart);
    } else {
      _closeActiveCosFor_(soNumber, '', 'Auto-confirmed (Custom Order Status left pending set)', editorEmail || _currentEditor_());
    }

    // Follow-up: keep while Sales Stage == Follow-Up Required; else confirm
    if (_equalsCI(newSalesStage, 'Follow-Up Required')) {
      ensureFollowUp(soNumber, opts || {});
    } else {
      _setQueueStatusFor(
        soNumber,
        [REMIND.TYPE_FOLLOWUP],
        REMIND.ST_CONFIRMED,
        'Auto-confirmed (Sales Stage moved off Follow-Up Required)',
        editorEmail || _currentEditor_()
      );
    }
  }


  // ---------- Menu actions (Snooze / Cancel) ----------
  function snoozeForSO(soNumber, isoDateStr /* YYYY-MM-DD */) {
    const q = _readQueue();
    const soCol = q.idx.col('soNumber');
    const statusCol = q.idx.col('status');
    const snoozeCol = q.idx.col('snoozeUntil');

    const snoozeAt = _at930(new Date(isoDateStr));
    const snoozeStr = _pt930Str_(isoDateStr);

    for (let i=0; i<q.rows.length; i++) {
      const r = q.rows[i];
      const rSo = _normalizeSO(r[soCol-1]);

      const rStatus = _normalize(r[statusCol-1]);
      if (!_eqSO_(r[soCol-1], soNumber)) continue;
      if (rStatus === REMIND.ST_CONFIRMED || rStatus === REMIND.ST_CANCELLED) continue;
      const rowIdx = i+2;
      q.sh.getRange(rowIdx, statusCol).setValue(REMIND.ST_SNOOZED);
      q.sh.getRange(rowIdx, snoozeCol).setValue(snoozeStr);
      _log(q.sh.getRange(rowIdx, q.idx.col('id')).getValue(), soNumber, q.sh.getRange(rowIdx, q.idx.col('type')).getValue(), 'SNOOZED', _currentEditor_(), 'Until ' + snoozeStr);
    }
  }

  function cancelForSO(soNumber) {
    _setQueueStatusFor(soNumber, [], REMIND.ST_CANCELLED, 'Manually cancelled', _currentEditor_());
  }

  // ---- DV helpers ----
  const DV_TYPES = Object.freeze({
    PROPOSE_NUDGE: 'DV_PROPOSE_NUDGE',
    URGENT_DAILY:  'DV_URGENT_OTW_DAILY'
  }); // :contentReference[oaicite:7]{index=7}

  function _isDVType_(t) {
    const s = String(t || '').trim().toUpperCase();
    return s === DV_TYPES.PROPOSE_NUDGE || s === DV_TYPES.URGENT_DAILY;
  }

  // Parse 'DV|<ROOT>|<TYPE>[|YYYYMMDD]' → '<ROOT>'
  function _dvRootFromId_(id) {
    const p = String(id || '').split('|');
    return (p.length >= 3 && p[0] === 'DV') ? p[1] : '';
  }


  // ---------- DAILY SEND (9:30) ----------
  function runDailySend_() {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return;   // ← ensure single run

    try {
      const props = _props();
      const todayStr = _dateOnlyStr(_todayAt930());

      const orders = _ordersData();
      const soIndex = _buildSoIndex(orders);
      const rootIndex = _buildRootIndex(orders);

      let ctx = _readQueue();

      _autoConfirmFromSheet_(ctx, orders, soIndex);

      // Auto-confirm DV rows once based on CSOS stop states
      _autoConfirmDVFromMaster_(ctx, orders, soIndex, rootIndex);

      // Refresh context after any auto-confirms
      ctx = _readQueue();

      const now = _now();
      const due = _collectDue_(ctx, now);


      const teamMsg    = _buildTeamMessage_(ctx, due, orders, soIndex, rootIndex); // ← pass rootIndex
      const managerMsg = _buildManagerMessage_(due); // (we’ll adjust manager builder separately)
      console.log('Due count:', due.length, 'First due:', due[0] && JSON.stringify(due[0]));
      console.log('Team msg len:', (teamMsg||'').length, 'Mgr msg len:', (managerMsg||'').length);

      const sentAny = _sendTeamAndManager_(teamMsg, managerMsg);
      _afterSendUpdate_(ctx, due);

      if (sentAny) props.setProperty(REMIND.PROP_DAILY_SENT_FOR, todayStr);
    } finally {
      try { lock.releaseLock(); } catch (e) {}
    }
  }


  function _autoConfirmFromSheet_(ctx, orders, soIndex) {
    // Use the queue sheet's header index map
    const qIdx       = ctx.idx;
    const soCol      = qIdx.col('soNumber');
    const typeCol    = qIdx.col('type');
    const statusCol  = qIdx.col('status');
    const idCol      = qIdx.col('id');
    const confirmedAtCol = qIdx.col('confirmedAt');

    const salesStageIdx = orders.salesStageCol - 1;
    const cosIdx        = orders.cosCol - 1;

    for (let i = 0; i < ctx.rows.length; i++) {
      const r = ctx.rows[i];
      const rStatus = String(r[statusCol - 1] || '').trim();
      if (rStatus === REMIND.ST_CONFIRMED || rStatus === REMIND.ST_CANCELLED) continue;

      const so = String(r[soCol - 1] || '').trim();
      if (!so) continue;

      const orderRow = _ordersRowBySo(so, orders, soIndex);
      if (!orderRow) continue;

      const rowVals   = orderRow.rowVals;
      const salesStage = String(rowVals[salesStageIdx] || '').trim();
      const cos        = String(rowVals[cosIdx] || '').trim();
      const rType      = String(r[typeCol - 1] || '').trim();

    // ⛔ Hard guard: do not ever auto-confirm DV rows
      if (typeof _isDVType_ === 'function' && _isDVType_(rType)) {
        continue; // DV_* rows keep cycling daily until CSOS changes; never auto-confirm here
      }

    // COS items: confirm when COS leaves the pending set
      if (rType === REMIND.TYPE_COS && !_is3DPendingStatus_(cos)) {
        const rowIdx = i + 2;
        const id = ctx.sh.getRange(rowIdx, idCol).getValue();
        ctx.sh.getRange(rowIdx, statusCol).setValue(REMIND.ST_CONFIRMED);
        ctx.sh.getRange(rowIdx, confirmedAtCol)
              .setValue(Utilities.formatDate(_now(), REMIND.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'));
        _log(id, so, rType, REMIND.ST_CONFIRMED, '', 'Auto-confirm (COS left pending set)');
        continue;
      }


      // Auto-confirm Follow-up when Sales Stage moved off Follow-Up Required
      if (rType === REMIND.TYPE_FOLLOWUP && !_equalsCI(salesStage, 'Follow-Up Required')) {
        const rowIdx = i + 2;
        const id = ctx.sh.getRange(rowIdx, idCol).getValue();
        ctx.sh.getRange(rowIdx, statusCol).setValue(REMIND.ST_CONFIRMED);
        ctx.sh.getRange(rowIdx, confirmedAtCol)
              .setValue(Utilities.formatDate(_now(), REMIND.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'));
        _log(id, so, rType, REMIND.ST_CONFIRMED, '', 'Auto-confirm (Sales Stage moved off Follow-Up Required)');
      }
    }
  }

  function _autoConfirmDVFromMaster_(ctx, orders, soIndex, rootIndex) {
    if (!orders) return;
    // We need at least one column to read a DV status signal
    const statusColInOrders = orders.csosCol || orders.cosCol;
    if (!statusColInOrders) return;

    const qIdx   = ctx.idx;
    const idCol  = qIdx.col('id');
    const soCol  = qIdx.col('soNumber');
    const typeCol= qIdx.col('type');
    const stCol  = qIdx.col('status');
    const confAt = qIdx.col('confirmedAt');

    for (let i = 0; i < ctx.rows.length; i++) {
      const r    = ctx.rows[i];
      const rowI = i + 2;

      const status = String(r[stCol - 1] || '').trim().toUpperCase();
      if (status === REMIND.ST_CONFIRMED || status === REMIND.ST_CANCELLED) continue;

      const rType = String(r[typeCol - 1] || '').trim().toUpperCase();
      // Only DV rows
      if (typeof _isDVType_ === 'function' ? !_isDVType_(rType) : true) continue;

      const idVal = ctx.sh.getRange(rowI, idCol).getValue();
      const rootId = (typeof _dvRootFromId_ === 'function') ? _dvRootFromId_(idVal) : '';

      // Prefer RootApptID; else fall back to SO
      let ord = null;
      if (rootId && typeof _ordersRowByRoot === 'function') ord = _ordersRowByRoot(rootId, orders, rootIndex);
      if (!ord) {
        const so = String(r[soCol - 1] || '').trim();
        ord = (so ? _ordersRowBySo(so, orders, soIndex) : null);
      }
      if (!ord) continue;

      const rowVals = ord.rowVals;
      const csos = rowVals[statusColInOrders - 1];

      // Flip to CONFIRMED when CSOS hits Delivered / Viewing Ready / Deposit Confirmed / OTW (incl. SOME OTW)
      if (typeof DV_shouldStopDailyForStatus === 'function' && DV_shouldStopDailyForStatus(csos)) {
        const nowStr = Utilities.formatDate(new Date(), REMIND.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
        ctx.sh.getRange(rowI, stCol).setValue(REMIND.ST_CONFIRMED);
        ctx.sh.getRange(rowI, confAt).setValue(nowStr);
        _log(ctx.sh.getRange(rowI, idCol).getValue(),
            String(r[soCol-1]||''),
            rType,
            REMIND.ST_CONFIRMED,
            '',
            'Auto-confirm (CSOS reached Delivered/Viewing Ready/Deposit Confirmed/OTW)');
      }
    }
  }



  function _collectDue_(ctx, now) {
    const statusCol = ctx.idx.col('status');
    const nextDueCol = ctx.idx.col('nextDueAt');
    const snoozeCol = ctx.idx.col('snoozeUntil');
    const typeCol = ctx.idx.col('type');
    const soCol = ctx.idx.col('soNumber');
    const firstDueCol = ctx.idx.col('firstDueDate');

    const out = [];
    for (let i=0; i<ctx.rows.length; i++) {
      const r = ctx.rows[i];
      const st = _normalize(r[statusCol-1]).toUpperCase();
      if (st !== REMIND.ST_PENDING && st !== REMIND.ST_SNOOZED) continue;

      const nextDueAtStr = _normalize(r[nextDueCol-1]);
      if (!nextDueAtStr) continue;

      // 1‑minute grace; both strings are PT 'yyyy-MM-dd HH:mm:ss'
      const nowPT = _nowPTStr_(1); // already 'YYYY-MM-DD HH:mm:ss'
      const duePT = _canonPTDateTime_(nextDueAtStr);
      if (!duePT || nowPT < duePT) continue; // not yet due in PT

      // Snooze guard: if snoozed into the future in PT, skip
      const snoozeStr = _canonPTDateTime_(_normalize(r[snoozeCol-1]));
      if (st === REMIND.ST_SNOOZED && snoozeStr && nowPT < snoozeStr) continue;

      // >>> DEBUG: log each row that will be included in "due"
      const idDbg  = ctx.sh.getRange(i+2, ctx.idx.col('id')).getValue();
      const tyDbg  = _normalize(r[typeCol-1]);
      const soDbg  = _normalizeSO(r[soCol-1]) || '(pending)';
      _dbg('[DBG_DUE_INCL]', { row: i+2, id: idDbg, so: soDbg, type: tyDbg, status: st,
                              duePT, nowPT, snoozeStr });

      out.push({
        rowIdx: i+2,
        id: ctx.sh.getRange(i+2, ctx.idx.col('id')).getValue(),
        so: _normalizeSO(r[soCol-1]),
        type: _normalize(r[typeCol-1]),
        firstDueDate: _normalize(r[firstDueCol-1]) // YYYY-MM-DD
      });
    }
    return out;
  }

  /**
   * Build the Team message for due reminders.
   * - DV_* types are rendered in the CSOS section (never COS).
   * - 3+ days overdue items are listed ONLY in the Overdue section and excluded from other sections/counts.
   * - Includes lightweight debug logging that can be toggled with the Script Property: REMIND_DEBUG_TEAM = '1'
   */
  function _buildTeamMessage_(ctx, due, orders, soIndex, rootIndex) {
    if (!due || !due.length) return '';

    // ---- debug toggle + helper ----
    // Set Script Property "REMIND_DEBUG_TEAM" to '1' to enable (Project Settings → Script properties).
    const DBG_ON = (function () {
      try { return PropertiesService.getScriptProperties().getProperty('REMIND_DEBUG_TEAM') === '1'; }
      catch (e) { return false; }
    })();
    const _dbg = function (tag, payload) {
      if (!DBG_ON) return;
      try {
        if (payload === undefined) console.log('[' + tag + ']');
        else console.log('[' + tag + ']', payload);
      } catch (e) {}
    };

    _dbg('DBG_TEAM_START', { dueN: due.length });

    const gid = _getGid(orders.sh);
    const sections = { DUE_COS: [], DUE_CSOS: [], DUE_FOLLOW: [], OVERDUE_3PLUS: [] };
    const seenDV = new Set(); // guard to avoid duplicate DV bullets across roots

    // Column indexes in Orders (0-based for row arrays)
    const nIdx = {
      so:        orders.soCol - 1,
      cust:      orders.custCol - 1,
      aRepName:  orders.aRepNameCol - 1,
      asRepName: orders.asRepNameCol ? orders.asRepNameCol - 1 : null
    };

    // Column indexes in Queue (1-based for getRange)
    const qAssignedCol = ctx.idx.col('assignedRepName');
    const qAssistedCol = ctx.idx.col('assistedRepName');
    const qCustomerCol = ctx.idx.col('customerName');
    const qNextCol     = ctx.idx.col('nextSteps');

    const todayAt930 = _todayAt930();

    // Helper to format one COS/FU line consistently
    function fmtLine(opts) {
      const parts = [];
      parts.push(`• *SO#${opts.soOut}* — ${opts.customerOut}`);
      parts.push(`*Assigned:* ${opts.assignedOut}`);
      if (opts.assistedOut) parts.push(`*Assisted:* ${opts.assistedOut}`);
      if (opts.type === REMIND.TYPE_FOLLOWUP) {
        parts.push('_Stage:_ Follow‑Up Required');
      } else {
        parts.push(`_COS:_ ${opts.cosOut || 'in progress'}`);
      }
      if (opts.nextOut) parts.push(`_Next Steps:_ ${opts.nextOut}`);
      if (opts.link)    parts.push(`Open row: ${opts.link}`);
      if (opts.managerNote) parts.push('*(Manager notified)*');
      return parts.join(' — ');
    }

    for (let i = 0; i < due.length; i++) {
      const item = due[i];

      // Resolve Orders row by SO (may be null if SO is blank)
      const ordBySo = _ordersRowBySo(item.so, orders, soIndex);
      const row     = ordBySo ? ordBySo.rowVals   : null;
      const rowNum  = ordBySo ? ordBySo.rowNumber : null;
      const link    = rowNum ? _sheetUrlForRow(gid, rowNum) : '';

      // From Orders (preferred)
      const customerFromOrders = row ? _normalize(row[nIdx.cust]) : '';
      const assignedFromOrders = row ? _normalize(row[nIdx.aRepName]) : '';
      const assistedFromOrders = (row && nIdx.asRepName != null) ? _normalize(row[nIdx.asRepName]) : '';

      // From Queue (fallbacks)
      const assignedFromQueue = _normalize(ctx.sh.getRange(item.rowIdx, qAssignedCol).getDisplayValue());
      const assistedFromQueue = _normalize(ctx.sh.getRange(item.rowIdx, qAssistedCol).getDisplayValue());
      const customerFromQueue = _normalize(ctx.sh.getRange(item.rowIdx, qCustomerCol).getDisplayValue());

      // Display fields
      const assignedOut = assignedFromOrders || assignedFromQueue || '?';
      const assistedOut = assistedFromOrders || assistedFromQueue || '';
      const customerOut = customerFromOrders || customerFromQueue || '(Customer?)';
      const soOut       = item.so ? item.so : '(pending)';

      // Next Steps (Orders first, then Queue)
      let nextOut = '';
      if (orders.nextStepsCol && row) nextOut = _normalize(row[orders.nextStepsCol - 1]);
      if (!nextOut) nextOut = _normalize(ctx.sh.getRange(item.rowIdx, qNextCol).getDisplayValue());

      // Current COS from Orders (if available)
      const cosFromOrders = (row && orders.cosCol) ? _normalize(row[orders.cosCol - 1]) : '';

      // Overdue determination (based on firstDueDate @ 9:30)
      const firstDue    = new Date(item.firstDueDate + 'T00:00:00');
      const overdueDays = _dateDiffDays(todayAt930, _at930(firstDue));
      const isOverdue3p = overdueDays >= 3;

      // ---------- DV branch (CSOS) ----------
      if (typeof _isDVType_ === 'function' && _isDVType_(item.type)) {
        const rootId = (typeof _dvRootFromId_ === 'function') ? _dvRootFromId_(item.id) : '';
        const dvKey  = (rootId ? rootId : (item.so || '')) + '|' + String(item.type).toUpperCase();
        if (seenDV.has(dvKey)) {
          _dbg('DBG_TEAM_DUP_GUARD_DV', { id: item.id, dvKey });
          continue;
        }
        seenDV.add(dvKey);

        // Prefer locating Master row by RootApptID; else fall back to SO
        const ordByRoot = (rootId && orders.rootApptCol) ? _ordersRowByRoot(rootId, orders, rootIndex) : null;
        const best      = ordByRoot || ordBySo;
        const row2      = best ? best.rowVals   : null;
        const rowNum2   = best ? best.rowNumber : null;
        const link2     = rowNum2 ? _sheetUrlForRow(gid, rowNum2) : '';

        // Compose display fields for DV
        const soFromOrders = row2 ? _normalizeSO(row2[orders.soCol - 1]) : '';
        const soPretty     = _soPretty_(soFromOrders || item.so || '');

        const cust2   = row2 ? _normalize(row2[orders.custCol - 1]) : '';
        const assign2 = row2 ? _normalize(row2[orders.aRepNameCol - 1]) : '';
        const assist2 = (row2 && orders.asRepNameCol) ? _normalize(row2[orders.asRepNameCol - 1]) : '';

        const assignedOut2 = assign2 || assignedFromQueue || '?';
        const assistedOut2 = assist2 || assistedFromQueue || '';
        const customerOut2 = cust2   || customerFromQueue || '(Customer?)';

        // CSOS (DV status) + Next Steps
        const csosOut = (row2 && orders.csosCol) ? _normalize(row2[orders.csosCol - 1])
                    : (row2 && orders.cosCol)  ? _normalize(row2[orders.cosCol  - 1])
                    : '';

        let nextOut2 = '';
        if (orders.nextStepsCol && row2) nextOut2 = _normalize(row2[orders.nextStepsCol - 1]);
        if (!nextOut2) nextOut2 = nextOut;

        // Appointment time: ISO (if present) or Visit Date + Visit Time
        let prettyAppt = '';
        if (row2 && orders.apptIsoCol)  prettyAppt = _prettyAppt_(row2[orders.apptIsoCol - 1]);
        if (!prettyAppt && row2 && orders.visitDateCol) {
          const dateCell = row2[orders.visitDateCol - 1];
          const timeCell = orders.visitTimeCol ? row2[orders.visitTimeCol - 1] : '';
          prettyAppt = _prettyApptFromVisitCells_(dateCell, timeCell);
        }
        if (!prettyAppt) prettyAppt = '(appt time unavailable)';

        const lead = (String(item.type).toUpperCase() === 'DV_PROPOSE_NUDGE')
          ? `• Reminder: Need to propose diamonds for viewing on ${prettyAppt}`
          : `• *URGENT: Upcoming Diamond Viewing on ${prettyAppt}. No Diamonds Ordered.*`;

        const tail = [
          soPretty ? `*SO#${soPretty}*` : '',
          customerOut2,
          `*Assigned:* ${assignedOut2}`,
          assistedOut2 ? `*Assisted:* ${assistedOut2}` : '',
          csosOut ? `_CSOS:_ ${csosOut}` : '',
          nextOut2 ? `_Next Steps:_ ${nextOut2}` : '',
          link2 ? `Open row: ${link2}` : ''
        ].filter(Boolean).join(' — ');

        const msg = tail ? `${lead} — ${tail}` : lead;

        // 3+ days overdue takes precedence: only show in Overdue section
        if (isOverdue3p) {
          sections.OVERDUE_3PLUS.push(msg);
          _dbg('DBG_TEAM_PUSH_OVERDUE', { id: item.id, route: 'CSOS', days: overdueDays });
        } else {
          sections.DUE_CSOS.push(msg);
          _dbg('DBG_TEAM_PUSH_CSOS', { id: item.id, dvKey, days: overdueDays });
        }
        _dbg('DBG_TEAM_DV_BRANCH', { id: item.id, rootId, dvKey });
        continue; // ⛔ do not fall through to COS
      }

      // ---------- Non-DV branches ----------
      const line = fmtLine({
        type: item.type,
        soOut, customerOut, assignedOut, assistedOut,
        cosOut: cosFromOrders,
        nextOut, link,
        managerNote: isOverdue3p
      });

      if (item.type === REMIND.TYPE_FOLLOWUP) {
        if (isOverdue3p) {
          sections.OVERDUE_3PLUS.push(line);
          _dbg('DBG_TEAM_PUSH_OVERDUE', { id: item.id, route: 'FOLLOWUP', days: overdueDays });
        } else {
          sections.DUE_FOLLOW.push(line);
          _dbg('DBG_TEAM_PUSH_FOLLOWUP', { id: item.id, so: item.so });
        }
      } else {
        if (isOverdue3p) {
          sections.OVERDUE_3PLUS.push(line);
          _dbg('DBG_TEAM_PUSH_OVERDUE', { id: item.id, route: 'COS', days: overdueDays });
        } else {
          sections.DUE_COS.push(line);
          _dbg('DBG_TEAM_PUSH_COS', { id: item.id, so: item.so });
        }
      }
    } // end for

    // Compose final message
    const cosCount  = sections.DUE_COS.length;
    const csosCount = sections.DUE_CSOS.length;
    const fuCount   = sections.DUE_FOLLOW.length;
    const odCount   = sections.OVERDUE_3PLUS.length;

    _dbg('DBG_TEAM_COUNTS', { cosCount, csosCount, fuCount, overdue3p: odCount });

    const parts = [];
    parts.push('*Daily reminders — Team*');
    parts.push(`• Custom Order Status due: ${cosCount}`);
    parts.push(`• Center Stone Order Status due: ${csosCount}`);
    parts.push(`• Follow-ups due: ${fuCount}`);
    if (odCount) parts.push(`• Overdue (3+ days): ${odCount}`);

    if (cosCount) {
      parts.push('\n*Due today — Custom Order Status*');
      parts.push(sections.DUE_COS.join('\n'));
    }
    if (csosCount) {
      parts.push('\n*Due today — Center Stone Order Status*');
      parts.push(sections.DUE_CSOS.join('\n'));
    }
    if (fuCount) {
      parts.push('\n*Due today — Follow-ups*');
      parts.push(sections.DUE_FOLLOW.join('\n'));
    }
    if (odCount) {
      parts.push('\n*Overdue (3+ days) — escalated*');
      parts.push(sections.OVERDUE_3PLUS.join('\n'));
    }

    return parts.join('\n');
  }

  /**
   * Build the Manager digest message.
   * - Summarizes counts (Custom Order Status vs Follow-ups)
   * - Lists items overdue 3+ days (escalated), including SO#, Customer, Assigned/Assisted when available.
   * - Uses the queue for names/customer (no Orders dependency).
   */
  function _buildManagerMessage_(due) {
    if (!due || !due.length) return '';

    // Read queue once to resolve names/customer by row index
    const ctx = _readQueue();
    const qIdx = ctx.idx;
    const colCust     = (qIdx.tryCol ? qIdx.tryCol('customerName')     : null) || qIdx.col('customerName');
    const colAssigned = (qIdx.tryCol ? qIdx.tryCol('assignedRepName')  : null) || qIdx.col('assignedRepName');
    const colAssisted = (qIdx.tryCol ? qIdx.tryCol('assistedRepName')  : null) || qIdx.col('assistedRepName');

    const todayAt930 = _todayAt930();

    let cosCount = 0, csosCount = 0, fuCount = 0;
    const overdue = []; // {item, days, so, cust, assigned, assisted}

    due.forEach(item => {
      if (item.type === REMIND.TYPE_FOLLOWUP) {
        fuCount++;
      } else if (_isDVType_(item.type)) {
          const rootId = _dvRootFromId_(item.id);
          const dvKey  = (rootId ? rootId : (item.so || '')) + '|' + String(item.type).toUpperCase();
          _dbg('[DBG_TEAM_DV_BRANCH]', { id: item.id, rootId, dvKey });

        csosCount++;
      } else {
        cosCount++;
      }


      // Overdue if firstDueDate is 3+ days behind today
      const firstDue = new Date(item.firstDueDate + 'T00:00:00');
      const days = _dateDiffDays(todayAt930, _at930(firstDue));
      if (days >= 3) {
        // Pull details from the queue for this row
        let cust = '', assigned = '', assisted = '';
        try {
          cust     = _normalize(ctx.sh.getRange(item.rowIdx, colCust).getDisplayValue());
          assigned = _normalize(ctx.sh.getRange(item.rowIdx, colAssigned).getDisplayValue());
          assisted = _normalize(ctx.sh.getRange(item.rowIdx, colAssisted).getDisplayValue());
        } catch (e) {}
        overdue.push({
          item, days,
          so: item.so || '(pending)',
          cust: cust || 'Customer?',
          assigned, assisted
        });
      }
    });

    const lines = [];
      lines.push('*Daily reminders — Manager summary*');
      lines.push(`• Custom Order Status due: ${cosCount}`);
      lines.push(`• Center Stone Order Status due: ${csosCount}`);
      lines.push(`• Follow-ups due: ${fuCount}`);


    if (overdue.length) {
      lines.push('\n*Overdue (3+ days) — escalated*');
      overdue.forEach(o => {
        const typeLabel = (o.item.type === REMIND.TYPE_FOLLOWUP)
          ? 'Follow‑up'
          : (_isDVType_(o.item.type) ? 'Center Stone Order Status' : 'Custom Order Status');
        let line = `• SO#${o.so} — ${o.cust} — ${typeLabel} — ${o.days}d overdue`;

        if (o.assigned) line += ` — Assigned: ${o.assigned}`;
        if (o.assisted) line += ` — Assisted: ${o.assisted}`;
        lines.push(line);
      });
    }
    
    _dbg('[DBG_MGR_COUNTS]', { cosCount, csosCount, fuCount, overdueCount: overdue.length });
    return lines.join('\n');
  }


  function _sendTeamAndManager_(teamMsg, managerMsg) {
    const props = _props();
    const teamHook = props.getProperty(REMIND.PROP_TEAM_WEBHOOK);
    const mgrHook  = props.getProperty(REMIND.PROP_MANAGER_WEBHOOK);

    let sentAny = false;
    if (teamHook && teamMsg && _normalize(teamMsg).length) {
      _postToChat_(teamHook, teamMsg);
      sentAny = true;
    }
    if (mgrHook && managerMsg && _normalize(managerMsg).length) {
      _postToChat_(mgrHook, managerMsg);
      sentAny = true;
    }
    return sentAny;
  }

  function _postToChat_(webhookUrl, text) {
    const payload = JSON.stringify({ text: text });
    const res = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      payload: payload,
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code >= 300) {
      throw new Error('Chat webhook error ' + code + ': ' + res.getContentText());
    }
  }

  function _afterSendUpdate_(ctx, sentItems) {
    if (!sentItems.length) return;
    const nextDueCol     = ctx.idx.col('nextDueAt');
    const statusCol      = ctx.idx.col('status');
    const lastSentCol    = ctx.idx.col('lastSentAt');
    const attemptsCol    = ctx.idx.col('attempts');
    const typeCol        = ctx.idx.col('type');            // NEW
    const confirmedAtCol = ctx.idx.col('confirmedAt');     // NEW
    const confirmedByCol = ctx.idx.col('confirmedBy');     // NEW

    const nowStr = Utilities.formatDate(_now(), REMIND.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    sentItems.forEach(item => {
      const r = item.rowIdx;

      // POLICY (2025‑09‑03): DV URGENT stays active daily until CSOS stop‑state in Master.
      // Do not auto‑close here; fall through so the default reschedule runs.
      const thisType = String(ctx.sh.getRange(r, typeCol).getDisplayValue() || '').trim().toUpperCase();

      // Auto-close single DV Propose nudge after first send
      if (thisType === 'DV_PROPOSE_NUDGE') {
        ctx.sh.getRange(r, statusCol).setValue(REMIND.ST_CONFIRMED);
        ctx.sh.getRange(r, confirmedAtCol).setValue(nowStr);
        ctx.sh.getRange(r, confirmedByCol).setValue('system:auto');

        const attempts = Number(ctx.sh.getRange(r, attemptsCol).getValue() || 0) + 1;
        ctx.sh.getRange(r, lastSentCol).setValue(nowStr);
        ctx.sh.getRange(r, attemptsCol).setValue(attempts);

        _log(ctx.sh.getRange(r, ctx.idx.col('id')).getValue(), item.so, item.type, 'SENT', '', 'Auto-confirmed DV_PROPOSE_NUDGE');
        return; // do not reschedule this item
      }

      if (thisType === 'DV_URGENT_OTW_DAILY') {
        console.log('[DV] AFTER-SEND policy: persist URGENT_DAILY (reschedule to next day)');
        // No return; the generic advance-by-1-day logic below will run.
      }

      // existing advance-by-1-day logic …
      const cur = ctx.sh.getRange(r, statusCol).getValue();
      if (cur === REMIND.ST_PENDING || cur === REMIND.ST_SNOOZED) {
        const curNextStr = String(ctx.sh.getRange(r, nextDueCol).getDisplayValue() || '').trim();
        const baseIso    = curNextStr ? curNextStr.slice(0,10) : _todayIsoPT_();
        const baseDate   = new Date(baseIso + 'T00:00:00');
        const nextDate   = _addDays(baseDate, 1);
        ctx.sh.getRange(r, nextDueCol).setValue(_pt930Str_(nextDate));
        if (cur === REMIND.ST_SNOOZED) ctx.sh.getRange(r, statusCol).setValue(REMIND.ST_PENDING);
      }
      const attempts = Number(ctx.sh.getRange(r, attemptsCol).getValue() || 0) + 1;
      ctx.sh.getRange(r, lastSentCol).setValue(nowStr);
      ctx.sh.getRange(r, attemptsCol).setValue(attempts);
      _log(ctx.sh.getRange(r, ctx.idx.col('id')).getValue(), item.so, item.type, 'SENT', '', '');
    });
  }



    function _currentEditor_() {
      try {
        return Session.getActiveUser().getEmail() || '';
      } catch (e) {
        return '';
      }
    }

    /**
     * Create/update a Custom Order Status reminder for an SO.
     * If restart=true, close any prior active COS/3D rows that match this SO (or customer if SO is blank),
     * then create a fresh cycle due T+2 @ 9:30.
     * If restart=false, upsert by SO (or by customer if SO blank).
     */
    function scheduleCOS(soNumber, opts, restart /* boolean */) {
      const so   = _normalizeSO(soNumber);
      const cust = _normalize(opts && opts.customerName);

      if (restart) {
        // Close any COS/3D for this SO
        _closeActiveCosFor_(so, cust, 'Superseded by update (restart: SO)', _currentEditor_());

        // EXTRA GUARD: also close any blank-SO entries for this customer
        if (cust) {
          _closeActiveCosFor_('', cust, 'Superseded by update (restart: blank-SO by customer)', _currentEditor_());
        }
      }

      const firstDue = _addDays(_todayAt930(), 2);
      return _queueUpsert_(so, REMIND.TYPE_COS, firstDue, _currentEditor_(), {
        assignedRepName:  opts && opts.assignedRepName,
        assistedRepName:  opts && opts.assistedRepName,
        assignedRepEmail: opts && opts.assignedRepEmail,
        assistedRepEmail: opts && opts.assistedRepEmail,
        customerName:     cust,
        nextSteps:        opts && opts.nextSteps
      });
    }

    // Snooze by SO or (if SO blank) by Customer. Acts on all active items (PENDING/SNOOZED).
    function snoozeForTarget(soNumber, customerName, isoDateStr /* YYYY-MM-DD */) {
      const q = _readQueue();
      const soCol   = q.idx.col('soNumber');
      const stCol   = q.idx.col('status');
      const typeCol = q.idx.col('type');
      const snooCol = q.idx.col('snoozeUntil');
      const custCol = (q.idx.tryCol ? q.idx.tryCol('customerName') : null) || q.idx.col('customerName');
      const actCol  = q.idx.col('lastAdminAction');
      const byCol   = q.idx.col('lastAdminBy');

      const targetSO   = _normalizeSO(soNumber);
      const targetCust = _normalize(customerName);
      const snoozeStr  = _pt930Str_(isoDateStr);
      const who        = _currentEditor_();

      let changed = 0; // <-- needed

      for (let i = 0; i < q.rows.length; i++) {
        const r       = q.rows[i];
        const rStatus = _normalize(r[stCol-1]);
        if (rStatus === REMIND.ST_CONFIRMED || rStatus === REMIND.ST_CANCELLED) continue;

        const rSO   = _normalizeSO(r[soCol-1]);
        const rCust = _normalize(r[custCol-1]);
        const custMatch = targetCust && rCust && _equalsCI(rCust, targetCust);
        const soMatch   = _eqSO_(rSO, targetSO);

        const hit =
          (targetSO && soMatch) ||
          (targetSO && !rSO && custMatch) ||
          (!targetSO && custMatch);

        if (!hit) continue;

        const rowIdx = i + 2;
        q.sh.getRange(rowIdx, stCol).setValue(REMIND.ST_SNOOZED);
        q.sh.getRange(rowIdx, snooCol).setValue(snoozeStr);

        q.sh.getRange(rowIdx, actCol).setValue('SNOOZED until ' + snoozeStr);
        q.sh.getRange(rowIdx, byCol).setValue(who);

        _log(q.sh.getRange(rowIdx, q.idx.col('id')).getValue(),
            rSO || '(blankSO)',
            q.sh.getRange(rowIdx, typeCol).getValue(),
            'SNOOZED', who, 'Until ' + snoozeStr);

        changed++;
      }
      return changed; // so the UI can show how many rows were affected
    }

    // Cancel by SO or (if SO blank) by Customer. Optional types filter; reason required by UI.
    function cancelForTarget(soNumber, customerName, typesArray /* optional */, reason /* string */) {
      const q = _readQueue();
      const soCol   = q.idx.col('soNumber');
      const stCol   = q.idx.col('status');
      const typeCol = q.idx.col('type');
      const idCol   = q.idx.col('id');
      const custCol = (q.idx.tryCol ? q.idx.tryCol('customerName') : null) || q.idx.col('customerName');
      const reasonCol = q.idx.col('cancelReason');
      const actCol    = q.idx.col('lastAdminAction');
      const byCol     = q.idx.col('lastAdminBy');

      const targetSO   = _normalizeSO(soNumber);
      const targetCust = _normalize(customerName);
      const types      = Array.isArray(typesArray) && typesArray.length ? new Set(typesArray) : null;

      const nowStr = Utilities.formatDate(_now(), REMIND.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
      const who    = _currentEditor_();
      const why    = String(reason || '').trim();

      let changed = 0; // <-- count cancellations

      for (let i = 0; i < q.rows.length; i++) {
        const r       = q.rows[i];
        const rStatus = _normalize(r[stCol-1]);
        if (rStatus === REMIND.ST_CONFIRMED || rStatus === REMIND.ST_CANCELLED) continue;

        const rSO   = _normalizeSO(r[soCol-1]);
        const rType = _normalize(r[typeCol-1]);
        const rCust = _normalize(r[custCol-1]);

        if (types && !types.has(rType)) continue;

        const custMatch = targetCust && rCust && _equalsCI(rCust, targetCust);
        const soMatch   = _eqSO_(rSO, targetSO);

        const hit =
          (targetSO && soMatch) ||
          (targetSO && !rSO && custMatch) ||
          (!targetSO && custMatch);

        if (!hit) continue;

        const rowIdx = i + 2;
        q.sh.getRange(rowIdx, stCol).setValue(REMIND.ST_CANCELLED);
        q.sh.getRange(rowIdx, q.idx.col('confirmedAt')).setValue(nowStr);
        q.sh.getRange(rowIdx, q.idx.col('confirmedBy')).setValue(who);
        q.sh.getRange(rowIdx, reasonCol).setValue(why);
        q.sh.getRange(rowIdx, actCol).setValue('CANCELLED: ' + (why || '(no reason)'));
        q.sh.getRange(rowIdx, byCol).setValue(who);

        _log(q.sh.getRange(rowIdx, idCol).getValue(),
            rSO || '(blankSO)',
            rType,
            REMIND.ST_CANCELLED,
            who,
            why);

        changed++;
      }
      return changed;
    }


    // Bring any SNOOZED (and PENDING) rows for this order back into the daily cycle.
    // - status => PENDING
    // - snoozeUntil => '' (cleared)
    // - nextDueAt => next 9:30 AM PT (today if before 9:30, else tomorrow)
    // - lastAdminAction / lastAdminBy filled
    // - log 'UNSNOOZED'
    function unsnoozeForTarget(soNumber, customerName) {
      const q = _readQueue();
      const soCol   = q.idx.col('soNumber');
      const stCol   = q.idx.col('status');
      const typeCol = q.idx.col('type');
      const snooCol = q.idx.col('snoozeUntil');
      const nextCol = q.idx.col('nextDueAt');
      const custCol = (q.idx.tryCol ? q.idx.tryCol('customerName') : null) || q.idx.col('customerName');
      const actCol  = q.idx.col('lastAdminAction');
      const byCol   = q.idx.col('lastAdminBy');

      const targetSO   = _normalizeSO(soNumber);
      const targetCust = _normalize(customerName);

      // compute next 9:30 AM in project TZ
      const nowStrPT = _nowPTStr_(0);
      const todayIso = _todayIsoPT_();
      const today930 = _pt930Str_(todayIso);

      // If it's before 9:30 now → today 9:30; otherwise → tomorrow 9:30
      const todayDate = new Date(todayIso + 'T00:00:00');
      const nextDate  = (nowStrPT < today930) ? todayDate : _addDays(todayDate, 1);
      const nextStr   = _pt930Str_(nextDate);

      const who       = _currentEditor_();

      let changed = 0;
      for (let i = 0; i < q.rows.length; i++) {
        const r = q.rows[i];
        const rStatus = _normalize(r[stCol-1]);

        // Only touch active rows
        if (rStatus !== REMIND.ST_SNOOZED && rStatus !== REMIND.ST_PENDING) continue;

        const rSO   = _normalizeSO(r[soCol-1]);
        const rCust = _normalize(r[custCol-1]);
        const custMatch = targetCust && rCust && _equalsCI(rCust, targetCust);
        const soMatch   = _eqSO_(rSO, targetSO);

        const hit =
          (targetSO && soMatch) ||
          (targetSO && !rSO && custMatch) ||
          (!targetSO && custMatch);

        if (!hit) continue;

        const rowIdx = i + 2;

        // flip to active
        q.sh.getRange(rowIdx, stCol).setValue(REMIND.ST_PENDING);
        q.sh.getRange(rowIdx, snooCol).setValue('');
        q.sh.getRange(rowIdx, nextCol).setValue(nextStr);

        // breadcrumbs
        q.sh.getRange(rowIdx, actCol).setValue('UNSNOOZED (next: ' + nextStr + ')');
        q.sh.getRange(rowIdx, byCol).setValue(who);

        _log(q.sh.getRange(rowIdx, q.idx.col('id')).getValue(),
            rSO || '(blankSO)',
            q.sh.getRange(rowIdx, typeCol).getValue(),
            'UNSNOOZED', who, 'Next ' + nextStr);

        changed++;
      }

      return changed; // for UI feedback
    }



  // ---------- Exposed trigger runners (inside Remind IIFE) ----------
  function remindersDailyCron() {
    runDailySend_();
  }

  function remindersHourlySafetyNet() {
    // Only send if the daily 9:30 didn't run today
    const props = _props();
    const todayStr = _dateOnlyStr(_todayAt930());
    const last = props.getProperty(REMIND.PROP_DAILY_SENT_FOR);
    if (last === todayStr) return;
    runDailySend_();
  }

  // For manual testing from menu/script editor
  function runDailyNowForTesting() {
    runDailySend_();
  }

  // ---------- Public API (INSIDE the IIFE) ----------
  return {
    scheduleStart3D,
    scheduleAssignSO,
    schedule3DRevision,
    scheduleCOS,
    ensureFollowUp,
    onClientStatusChange,
    snoozeForSO,
    cancelForSO,
    unsnoozeForTarget,
    snoozeForTarget,
    cancelForTarget,
    remindersDailyCron,
    remindersHourlySafetyNet,
    runDailyNowForTesting
  };
})(); // <— CLOSE THE Remind IIFE EXACTLY ONCE HERE


function remind_menu_snoozeSelected() {
  const ui = SpreadsheetApp.getUi();
  let t;
  try { t = remind__getSelectionTarget_(); }
  catch (e) { ui.alert(e.message); return; }

  const tz = REMIND.TIMEZONE || Session.getScriptTimeZone?.() || 'America/Los_Angeles';
  const now = new Date();
  const hh = Number(Utilities.formatDate(now, tz, 'H'));
  const mm = Number(Utilities.formatDate(now, tz, 'm'));
  const isBefore930 = (hh < 9) || (hh === 9 && mm < 30);

  const base = new Date(now.getFullYear(), now.getMonth(), now.getDate() + (isBefore930 ? 0 : 1));
  const todayStr   = Utilities.formatDate(now,  tz, 'yyyy-MM-dd');
  const defaultStr = Utilities.formatDate(base, tz, 'yyyy-MM-dd');

  const tpl = HtmlService.createTemplateFromFile('dlg_reminders_snooze');
  const active = remind__getActiveSummaryForTarget(t.so, t.cust);
  tpl.payload = {
    so: t.so, cust: t.cust,
    label: t.display || t.label,
    tz, todayStr, defaultDateStr: defaultStr,
    active: active      // <— include active reminders
  };

  const html = tpl.evaluate().setWidth(420).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, 'Snooze');
}

// Small date parser for YYYY-MM-DD or MM/DD/YYYY
function tryParseDate_(s) {
  const t = String(s||'').trim();
  let m = t.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2])-1, Number(m[3]));
  m = t.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[1])-1, Number(m[2]));
  return null;
}

function remind_menu_cancelSelected() {
  const ui = SpreadsheetApp.getUi();
  let t;
  try { t = remind__getSelectionTarget_(); } catch (e) { ui.alert(e.message); return; }

  // NEW: bring current active reminders into the dialog
  const active = remind__getActiveSummaryForTarget(t.so, t.cust);

  const tpl = HtmlService.createTemplateFromFile('dlg_reminders_cancel');
  tpl.payload = {
    so: t.so,
    cust: t.cust,
    label: t.display || t.label,
    active: active                   // <-- pass to HTML
  };
  const html = tpl.evaluate().setWidth(420).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cancel reminders');
}



/** Validate & execute cancel from dialog (type dropdown + reason). */
function remind__cancelTarget_do(so, cust, choice, reason) {
  const t   = String(choice || '').toUpperCase();
  const why = String(reason || '').trim();
  if (!why) throw new Error('Please provide a reason to cancel.');

  let types = null; // BOTH by default
  if (t === 'COS') {
    types = [REMIND.TYPE_COS];
  } else if (t === 'FOLLOWUP' || t === 'FOLLOW-UP' || t === 'FU') {
    types = [REMIND.TYPE_FOLLOWUP];
  }

  // If a specific type was chosen, verify the target actually has that type active.
  if (types && types.length === 1) {
    const active = remind__getActiveSummaryForTarget(so, cust)
      .map(x => String(x.type || '').toUpperCase());
    if (!active.includes(types[0].toUpperCase())) {
      throw new Error('No "' + types[0] + '" reminder exists for this target.');
    }
  }

  // Fallback to current selection if payload is missing
  if (!so && !cust && typeof remind__getSelectionTarget_ === 'function') {
    try { const sel = remind__getSelectionTarget_(); so = so || sel.so; cust = cust || sel.cust; } catch (_){}
  }

  const changed = Remind.cancelForTarget(so, cust, types, why);
  return { ok: true, changed };
}


// Returns [{type, prettyNext, prettySnooze}] using the given tz for display (e.g., 'America/Los_Angeles').
function remind__getActiveSummaryForTarget(so, cust /* tz ignored; PT local */) {
  const ss  = SpreadsheetApp.getActive();
  const qsh = ss.getSheetByName('04_Reminders_Queue');
  if (!qsh) return [];

  const rg = qsh.getDataRange().getDisplayValues();
  const h  = rg[0] || [];
  const rows = rg.slice(1);
  const H = {}; h.forEach((x,i)=>{ const k=String(x||'').trim().toLowerCase(); if (k) H[k]=i+1; });

  const cSO   = H['sonumber'] || H['so number'] || H['so#'] || H['so'];
  const cCust = H['customername'] || H['customer name'] || H['customer'];
  const cType = H['type'], cStat = H['status'], cNxt = H['nextdueat'], cSnoo = H['snoozeuntil'];

  const norm = s => String(s||'').replace(/\u00A0/g,' ').replace(/\s+/g,' ').trim();
  const soKey     = _soKey_(so);
  const custLower = norm(cust).toLowerCase();

  const out = [];
  for (let i=0; i<rows.length; i++) {
    const st = norm(rows[i][cStat-1]).toUpperCase();
    if (st !== 'PENDING' && st !== 'SNOOZED') continue;

    const rKey  = cSO   ? _soKey_(rows[i][cSO-1]) : '';
    const rCust = cCust ? norm(rows[i][cCust-1])  : '';

    const soHit   = soKey && rKey && (rKey === soKey);
    const custHit = custLower && rCust && (rCust.toLowerCase() === custLower);

    const hit = soHit || (soKey && !rKey && custHit) || (!soKey && custHit);
    if (!hit) continue;

    out.push({
      type: rows[i][cType-1],
      prettyNext:   cNxt  ? remind__prettyLocalTimestamp_(rows[i][cNxt-1])  : '',
      prettySnooze: cSnoo ? remind__prettyLocalTimestamp_(rows[i][cSnoo-1]) : ''
    });
  }
  return out;
}




/** Validate & execute snooze from dialog. */
function remind__snoozeTarget_do(so, cust, isoDateStr) {
  if (!isoDateStr || !/^\d{4}-\d{2}-\d{2}$/.test(String(isoDateStr))) {
    throw new Error('Invalid date. Use the calendar.');
  }
  // Validate not in the past (cheap check vs today in project TZ)
  const tz = (typeof REMIND !== 'undefined' ? REMIND.TIMEZONE : null) || Session.getScriptTimeZone?.() || 'America/Los_Angeles';
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  if (isoDateStr < todayStr) throw new Error('Date must be today or later.');

  const changed = Remind.snoozeForTarget(so, cust, isoDateStr);
  if (!changed) throw new Error('No active reminders to snooze for this order/customer.');
  return { ok: true, changed, untilPretty: remind__pretty930ForIsoDate_(isoDateStr) };
}


/**
 * Resolve the currently selected data row into { so, cust, row, label }.
 * Throws a user-friendly error if no data row is selected.
 */
function remind__getSelectionTarget_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  const r  = ss.getActiveRange();

  if (!sh || !r || r.getSheet().getName() !== sh.getName() || r.getRow() < 2) {
    throw new Error('Please select a data row in "00_Master Appointments" first.');
  }

  const row  = r.getRow();
  const hdrs = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getDisplayValues()[0] || [];
  const H    = {}; hdrs.forEach((h,i)=>{ const k=String(h||'').trim(); if (k) H[k]=i+1; });

  const so   = H['SO#'] ? String(sh.getRange(row, H['SO#']).getDisplayValue()).replace(/^'+/, '').trim() : '';
  const cust = H['Customer Name']
              ? String(sh.getRange(row, H['Customer Name']).getDisplayValue()).trim()
              : (H['Customer'] ? String(sh.getRange(row, H['Customer']).getDisplayValue()).trim() : '');

  // old "label" kept for back-compat; new "display" shows both SO and Customer if available
  const label   = so ? ('SO#' + so) : (cust ? ('Customer: ' + cust) : '(unknown)');
  const display = (so && cust) ? ('SO#' + so + ' — ' + cust) : (so ? ('SO#' + so) : (cust || '(unknown)'));

  return { so, cust, row, label, display };
}

function remind_menu_snoozeTodaySelected() {
  const ui = SpreadsheetApp.getUi();
  let t; try { t = remind__getSelectionTarget_(); } catch (e) { ui.alert(e.message); return; }

  const tz = REMIND.TIMEZONE || Session.getScriptTimeZone?.() || 'America/Los_Angeles';
  const now = new Date();
  const hh = Number(Utilities.formatDate(now, tz, 'H'));
  const mm = Number(Utilities.formatDate(now, tz, 'm'));
  const isBefore930 = (hh < 9) || (hh === 9 && mm < 30);

  const base = new Date(now.getFullYear(), now.getMonth(), now.getDate() + (isBefore930 ? 0 : 1));
  const dateStr = Utilities.formatDate(base, tz, 'yyyy-MM-dd');
  const pretty  = Utilities.formatDate(new Date(Utilities.formatDate(new Date(Date.UTC(base.getFullYear(), base.getMonth(), base.getDate(), 9, 30, 0)), tz, 'yyyy/MM/dd HH:mm:ss')), tz, "EEE, MMM d 'at' h:mma");

  const n = Remind.snoozeForTarget(t.so, t.cust, dateStr);
  ui.alert(`Snoozed ${n} reminder(s) for ${t.display || t.label} until ${pretty}.`);
}


function remind_menu_unsnoozeNowSelected() {
  const ui = SpreadsheetApp.getUi();
  let t; try { t = remind__getSelectionTarget_(); } catch (e) { ui.alert(e.message); return; }

  const n = Remind.unsnoozeForTarget(t.so, t.cust);
  if (!n) { ui.alert('No SNOOZED or PENDING reminders found for ' + (t.display || t.label) + '.'); return; }
  ui.alert('Unsnoozed ' + n + ' reminder(s) for ' + (t.display || t.label) + '.');
}

// Put near the queue helpers
function _writePrettyRow_(q, rowIdx, payload){
  if (!payload) return; // no-op if nothing to write
  const pairs = [
    ['soNumber','soNumber'],
    ['nextDueAt','nextDueAt'],
    ['firstDueDate','firstDueDate'],
    ['status','status'],
    ['assignedRepName','assignedRepName'],
    ['assignedRepEmail','assignedRepEmail'],
    ['assistedRepName','assistedRepName'],
    ['assistedRepEmail','assistedRepEmail'],
    ['customerName','customerName'],
    ['nextSteps','nextSteps'],
  ];
  for (let i=0;i<pairs.length;i++){
    const key = pairs[i][0], col = pairs[i][1];
    const v = payload[key];
    if (v !== undefined && v !== null && v !== '') {
      q.sh.getRange(rowIdx, q.idx.col(col)).setValue(v);
    }
  }
}

function remind__debugCheckWebhooks() {
  const p = PropertiesService.getScriptProperties();
  const team = p.getProperty(REMIND.PROP_TEAM_WEBHOOK);
  const mgr  = p.getProperty(REMIND.PROP_MANAGER_WEBHOOK);
  console.log('TEAM_CHAT_WEBHOOK set?', !!team, 'length:', team ? team.length : 0);
  console.log('MANAGER_CHAT_WEBHOOK set?', !!mgr, 'length:', mgr ? mgr.length : 0);
  return { teamSet: !!team, managerSet: !!mgr };
}

/** Install/refresh the time-based triggers for daily Chat reminders. */
function remind__installTimeTriggers() {
  // Clean up any prior copies of our runners to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'Remind.remindersDailyCron' || fn === 'Remind.remindersHourlySafetyNet') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Daily @ 9:30 AM in your project TZ (REMIND.TIMEZONE = 'America/Los_Angeles')
  ScriptApp.newTrigger('Remind.remindersDailyCron')
    .timeBased()
    .atHour(9)                 // 9 AM hour
    .nearMinute(30)            // as close to :30 as Apps Script can schedule
    .everyDays(1)
    .inTimezone(REMIND.TIMEZONE)
    .create();

  // Hourly safety net: only sends if the daily didn’t mark today as sent
  ScriptApp.newTrigger('Remind.remindersHourlySafetyNet')
    .timeBased()
    .everyHours(1)
    .create();
}

/** Run this after installing webhooks to confirm end-to-end posting works. */
function remind__sendNowForTesting() {
  if (typeof Remind === 'undefined' || typeof Remind.runDailyNowForTesting !== 'function') {
    throw new Error('Remind API not loaded. Check IIFE initialization.');
  }
  Remind.runDailyNowForTesting();
}


function DV_M_getApptISO_(H, sh, row) {
  // First try ISO-style columns if present (back-compat)
  var cands = [
    'ApptDateTimeISO','ApptDateTime (ISO)','Appointment Date/Time ISO','Appointment DateTime (ISO)',
    'Appt Start ISO','Appt Start (ISO)','Appt Start','Appointment Start','Event Start ISO','EventStartISO'
  ];
  var v = DV_M_pick_(H, sh, row, cands);
  if (v) {
    if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return v.toISOString();
    var s = String(v||'').trim(), d = s ? new Date(s) : null;
    if (d && !isNaN(d)) return d.toISOString();
  }

  // Fallback: Visit Date + Visit Time (your current sheet shape)
  var dVal = DV_M_pick_(H, sh, row, ['Visit Date','Appt Date','Appointment Date','Date','Event Date','Start Date']);
  if (!dVal) return ''; // no date

  var tVal = DV_M_pick_(H, sh, row, ['Visit Time','Appt Time','Appointment Time','Time','Event Time','Start Time']);
  return DV_M_joinDateTimeToISO_(dVal, tVal) || '';
}

/** Join separate date + time cells into an ISO string in project TZ. */
function DV_M_joinDateTimeToISO_(dVal, tVal) {
  var tz = Session.getScriptTimeZone?.() || 'America/Los_Angeles';

  var y, mo, da;
  if (Object.prototype.toString.call(dVal) === '[object Date]' && !isNaN(dVal)) {
    y  = dVal.getFullYear(); mo = dVal.getMonth(); da = dVal.getDate();
  } else {
    var ds = String(dVal||'').trim();
    var mYMD = ds.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    var mMDY = ds.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (mYMD)      { y=+mYMD[1]; mo=+mYMD[2]-1; da=+mYMD[3]; }
    else if (mMDY) { y=+mMDY[3]; mo=+mMDY[1]-1; da=+mMDY[2]; }
    else return '';
  }

  var hh = 9, mm = 0; // default 9:00 if blank
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

  var localStr = Utilities.formatDate(new Date(Date.UTC(y, mo, da, hh, mm, 0)), tz, 'yyyy/MM/dd HH:mm:ss');
  return new Date(localStr).toISOString();
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



