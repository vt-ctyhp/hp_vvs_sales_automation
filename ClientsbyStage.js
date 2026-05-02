/** ============================================================================
 * Clients by Stage — "By Sales Stage (Latest Visit)"
 * ----------------------------------------------------------------------------
 * Creates a materialized, read-only rollup from 00_Master Appointments:
 *  - One row per client, using latest row overall (by Visit Date) for status/notes.
 *  - Sections by Sales Stage (fixed order).
 *  - Sorting inside each section: by most recent PAST Visit Date (oldest → newest).
 *    • If no past visit exists, sort by earliest UPCOMING visit at the end of the section.
 *  - Columns:
 *      RootApptID, Customer, Assigned Rep, Brand, Visit Date (past),
 *      Next Visit (Scheduled), Sales Stage, Conversion Status, Next Steps,
 *      SO#, Client Status Report (link), Open Master Row (link),
 *      [Deposit/Won enrichment if available]: First Deposit Date, First Deposit Amount,
 *      Order Total, Paid-to-Date, Outstanding Balance
 *
 * Refresh model:
 *  - Timed trigger every 30 minutes during business hours (Mon–Sat, 08:00–18:00).
 *  - On-demand via menu: Client Rollup → Refresh now.
 *
 * Safe & robust:
 *  - Leaves 00_Master Appointments untouched (system of record).
 *  - Avoids heavy formulas. One read → in-memory transform → one write.
 *  - Gracefully handles missing optional sheets/columns.
 * ============================================================================
 */

/** =============================== CONFIG ================================== */

const CONFIG = {
  MASTER_SHEET_NAME: '00_Master Appointments',

  // --- Choose your preferred short name here (see list in the message) ---
  ROLLUP_SHEET_NAME: '02_Clients by Stage',

  // Fixed section order
  STAGE_ORDER: [
    'Appointment',
    'Lead',
    'Hot Lead',
    'Follow-Up Required',
    'Deposit',
    'Won',
    'Lost Lead',
  ],

  // Business-hours gating for the timed trigger
  BUSINESS_HOURS: {
    // 0=Sun, 1=Mon, ... 6=Sat. Allow Mon–Sat.
    allowedDays: new Set([1, 2, 3, 4, 5, 6]),
    startHour: 8,   // inclusive 08:00
    endHour: 18,    // exclusive 18:00
  },

  // Timed trigger cadence (minutes)
  TIMED_TRIGGER_MINUTES: 30,

  // Optional data sources (best-effort lookups). Leave as-is.
  OPTIONAL_SOURCES: {
    // Accounts Receivable / Payments sheet: used to compute Paid-to-Date, First Deposit.
    // The code will try candidates in order and skip silently if none found.
    AR_SHEET_CANDIDATES: ['1) AR Master Data', 'AR Master Data', 'AR', 'AR_Data'],

    // Orders sheet: used to pick up Order Total (by SO#), if present.
    ORDERS_SHEET_CANDIDATES: ['1.2) ADM1 - Customer order', 'Customer order', 'Orders', 'ADM1 - Customer order'],

    // Quotation sheet (fallback if Order Total is stored there)
    QUOTES_SHEET_CANDIDATES: ['1.1) ADM1 - Quotations', 'Quotations', 'Quotes']
  },

  // Header aliases to make the code resilient to small naming drifts.
  HEADER_ALIASES: {
    RootApptID: ['rootapptid', 'root_appt_id', 'root appt id', 'root appointment id'],
    Customer: ['customer', 'client', 'customer name', 'name'],
    AssignedRep: ['assigned rep', 'assigned', 'rep', 'assignedto', 'assigned to', 'sales rep'],
    Brand: ['brand', 'company', 'brand/company', 'org'],

    VisitDate: ['visit date', 'visitdate', 'visit', 'appt date', 'appointment date', 'appointment'],
    SalesStage: ['sales stage', 'stage', 'salesstage'],
    ConversionStatus: ['conversion status', 'conversion', 'status'],
    NextSteps: ['next steps', 'next step', 'next action', 'follow-up notes', 'follow up notes', 'followup notes', 'notes'],

    SONumber: ['so#', 'so #', 'so number', 'so', 'sales order', 'so no', 'so no.', 'so num'],
    ClientStatusReportURL: ['client status report link', 'client status report', 'csr link', 'clientstatusreporturl', 'csr url'],

    Email: ['email', 'email address', 'customer email'],
    Phone: ['phone', 'phone number', 'customer phone', 'mobile', 'mobile phone']
  }
};

/** =============================== TRIGGERS ================================= */

function installOrUpdateTimedTrigger() {
  removeTimedTriggers();
  ScriptApp.newTrigger('timedRefreshHandler')
    .timeBased()
    .everyMinutes(CONFIG.TIMED_TRIGGER_MINUTES)
    .create();
  SpreadsheetApp.getActive().toast(
    'Timed trigger installed: every ' + CONFIG.TIMED_TRIGGER_MINUTES + ' minutes.',
    'Clients by Stage'
  );
}

function removeTimedTriggers() {
  const all = ScriptApp.getProjectTriggers();
  for (const t of all) {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'timedRefreshHandler') {
      ScriptApp.deleteTrigger(t);
    }
  }
  SpreadsheetApp.getActive().toast('Timed trigger(s) removed.', 'Clients by Stage');
}

function timedRefreshHandler() {
  const now = new Date();
  if (!isWithinBusinessHours_(now)) return; // Silent skip outside business hours
  refreshClientStageRollup();
}

function isWithinBusinessHours_(when) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const local = new Date(Utilities.formatDate(when, tz, "yyyy-MM-dd'T'HH:mm:ss"));
  const day = local.getDay(); // 0..6
  const hour = local.getHours(); // 0..23
  const bh = CONFIG.BUSINESS_HOURS;
  return bh.allowedDays.has(day) && hour >= bh.startHour && hour < bh.endHour;
}

/** ============================ PUBLIC ENTRYPOINT =========================== */

function refreshClientStageRollup() {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
  if (!src) throw new Error('Source sheet not found: ' + CONFIG.MASTER_SHEET_NAME);

  const data = src.getDataRange().getValues();
  if (!data || data.length < 2) {
    writeRollup_(ss, [], []); // write header only
    return;
  }

  const header = data[0].map(v => String(v || '').trim());
  const hmap = buildHeaderMap_(header, CONFIG.HEADER_ALIASES);

  // Guard: essential columns
  const required = ['VisitDate', 'SalesStage', 'Customer'];
  for (const key of required) {
    if (hmap[key] == null) {
      throw new Error(`Required column "${key}" is missing in ${CONFIG.MASTER_SHEET_NAME}.`);
    }
  }

  // Build AR (payments) and Orders maps (best effort)
  const soToPayments = buildPaymentsMapBySO_(ss); // { so: {paidToDate, firstDepositDate, firstDepositAmount} }
  const soToOrderTotal = buildOrderTotalsMapBySO_(ss); // { so: orderTotalNumber }

  const today = atMidnight_(new Date(), ss);

  // 1) Group rows by primary identity: RootApptID → Email → Phone → Name
  const groups = new Map(); // key => { rows: [ {i,rowObj} ] }
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const salesStage = normStr_(row[hmap.SalesStage]);
    if (!salesStage) continue; // exclude unstaged rows (per "Default")

    const key = pickPrimaryKey_(row, hmap);
    if (!key) continue;

    let g = groups.get(key);
    if (!g) { g = { rows: [] }; groups.set(key, g); }
    g.rows.push({ i: r + 1, row });
  }

  // 2) Reduce each group to a single representative row (latest overall by Visit Date),
  //    plus compute lastPastVisit and nextFutureVisit for sorting/UX.
  const entries = [];
  for (const [key, g] of groups) {
    if (!g.rows.length) continue;

    // Determine representative row: latest overall by VisitDate (including future)
    let rep = g.rows[0];
    let repDate = parseDate_(rep.row[hmap.VisitDate], ss);

    for (let k = 1; k < g.rows.length; k++) {
      const cand = g.rows[k];
      const d = parseDate_(cand.row[hmap.VisitDate], ss);
      if (dateGt_(d, repDate)) { rep = cand; repDate = d; }
    }

    // Compute lastPastVisit and nextFutureVisit from ALL rows in group
    let lastPast = null, nextFuture = null;
    for (const { row } of g.rows) {
      const d = parseDate_(row[hmap.VisitDate], ss);
      if (!d) continue;
      if (!dateGt_(d, today)) { // d <= today
        if (!lastPast || dateGt_(d, lastPast)) lastPast = d;
      } else { // future
        if (!nextFuture || dateLt_(d, nextFuture)) nextFuture = d;
      }
    }

    // Representative fields from the chosen rep row
    const rr = rep.row;
    const stageCanonical = canonicalStage_(rr[hmap.SalesStage]);

    if (!stageCanonical) continue; // unknown stage → skip

    // Compose fields
    const soVal = hmap.SONumber != null ? (rr[hmap.SONumber] || '') : '';
    const so = String(soVal || '').trim();

    // Financial enrichment (best effort)
    let paidToDate = null, firstDepDate = null, firstDepAmt = null, orderTotal = null, outstanding = null;
    if (so && (stageCanonical === 'Deposit' || stageCanonical === 'Won')) {
      const p = soToPayments[so];
      if (p) {
        paidToDate = toNumberOrNull_(p.paidToDate);
        firstDepDate = p.firstDepositDate || null;
        firstDepAmt = toNumberOrNull_(p.firstDepositAmount);
      }
      const ot = soToOrderTotal[so];
      if (ot != null) orderTotal = toNumberOrNull_(ot);

      if (orderTotal != null && paidToDate != null) {
        outstanding = orderTotal - paidToDate;
      }
    }

    const entry = {
      _section: stageCanonical, // for grouping & header
      _sortHasPast: !!lastPast, // true => sort before "future-only"
      _sortDate: lastPast ? lastPast : nextFuture, // used for ordering

      RootApptID: hmap.RootApptID != null ? String(rr[hmap.RootApptID] || '').trim() : '',
      Customer: String(rr[hmap.Customer]).trim(),
      AssignedRep: hmap.AssignedRep != null ? String(rr[hmap.AssignedRep] || '').trim() : '',
      Brand: hmap.Brand != null ? String(rr[hmap.Brand] || '').trim() : '',

      VisitDatePast: lastPast, // may be null
      NextVisitScheduled: nextFuture, // may be null

      SalesStage: stageCanonical,
      ConversionStatus: hmap.ConversionStatus != null ? String(rr[hmap.ConversionStatus] || '').trim() : '',
      NextSteps: hmap.NextSteps != null ? String(rr[hmap.NextSteps] || '').trim() : '',
      SONumber: so,

      ClientStatusReportURL: hmap.ClientStatusReportURL != null ? String(rr[hmap.ClientStatusReportURL] || '').trim() : '',
      _masterRowIndex: rep.i // for "Open Master Row" link
    };

    // Attach financials if relevant/available
    if (stageCanonical === 'Deposit' || stageCanonical === 'Won') {
      entry.FirstDepositDate = firstDepDate;
      entry.FirstDepositAmount = firstDepAmt;
      entry.OrderTotal = orderTotal;
      entry.PaidToDate = paidToDate;
      entry.OutstandingBalance = outstanding;
    }

    entries.push(entry);
  }

  // 3) Partition by stage, then sort inside each partition
  const stageBuckets = new Map();
  for (const stage of CONFIG.STAGE_ORDER) stageBuckets.set(stage, []);
  for (const e of entries) {
    const bucket = stageBuckets.get(e._section);
    if (bucket) bucket.push(e);
  }
  for (const stage of CONFIG.STAGE_ORDER) {
    const arr = stageBuckets.get(stage);
    arr.sort((a, b) => {
      // 1) Rows with a past visit come before "future-only"
      if (a._sortHasPast !== b._sortHasPast) return a._sortHasPast ? -1 : 1;
      // 2) Oldest → newest by the chosen sort date (past or future)
      const da = a._sortDate, db = b._sortDate;
      if (!da && !db) return 0;
      if (!da) return 1;
      if (!db) return -1;
      return da.getTime() - db.getTime();
    });
  }

  // 4) Flatten into output rows (with section headers)
  const headerOut = buildOutputHeader_();
  const outRows = [];
  const ssId = ss.getId();
  const masterSheetId = src.getSheetId();

  const pushSectionHeader = (stage) => {
    // Section header row (visual separator). Keep simple; no merges for stability.
    outRows.push([`— ${stage} —`].concat(Array(headerOut.length - 1).fill('')));
  };

  for (const stage of CONFIG.STAGE_ORDER) {
    const arr = stageBuckets.get(stage);
    if (!arr || !arr.length) continue;
    pushSectionHeader(stage);
    for (const e of arr) {
      const openMasterLinkFormula = makeOpenMasterLinkFormula_(ssId, masterSheetId, e._masterRowIndex);
      const csrLinkFormula = makeHyperlinkFormulaOrBlank_(e.ClientStatusReportURL, 'Open');

      const row = [
        e.RootApptID || '',
        e.Customer || '',
        e.AssignedRep || '',
        e.Brand || '',
        dateToCell_(e.VisitDatePast),
        dateToCell_(e.NextVisitScheduled),
        e.SalesStage || '',
        e.ConversionStatus || '',
        e.NextSteps || '',
        e.SONumber || '',

        csrLinkFormula || '',
        openMasterLinkFormula || '',
      ];

      // Deposit/Won enrichment columns (ensure consistent positions with header)
      const isMoneyStage = (e.SalesStage === 'Deposit' || e.SalesStage === 'Won');
      const firstDepDateCell = isMoneyStage ? dateToCell_(e.FirstDepositDate) : '';
      const firstDepAmtCell = isMoneyStage && e.FirstDepositAmount != null ? e.FirstDepositAmount : '';
      const orderTotalCell = isMoneyStage && e.OrderTotal != null ? e.OrderTotal : '';
      const ptdCell = isMoneyStage && e.PaidToDate != null ? e.PaidToDate : '';
      const outstandingCell = isMoneyStage && e.OutstandingBalance != null ? e.OutstandingBalance : '';

      row.push(firstDepDateCell, firstDepAmtCell, orderTotalCell, ptdCell, outstandingCell);
      outRows.push(row);
    }
    // Blank spacer row after each section
    outRows.push(Array(headerOut.length).fill(''));
  }

  writeRollup_(ss, headerOut, outRows);
  SpreadsheetApp.getActive().toast('Clients by Stage refreshed.', 'Clients by Stage', 5);
}

/** =============================== HELPERS ================================== */

// Build output header (kept in one place for clarity)
function buildOutputHeader_() {
  return [
    'RootApptID',
    'Customer',
    'Assigned Rep',
    'Brand',
    'Visit Date (Past)',
    'Next Visit (Scheduled)',
    'Sales Stage',
    'Conversion Status',
    'Next Steps',
    'SO#',
    'Client Status Report',
    'Open Master Row',
    'First Deposit Date',
    'First Deposit Amount',
    'Order Total',
    'Paid-to-Date',
    'Outstanding Balance'
  ];
}

// Create hyperlink formula to open the Master sheet at a given row
function makeOpenMasterLinkFormula_(spreadsheetId, masterSheetId, rowIndex) {
  if (!spreadsheetId || !masterSheetId || !rowIndex) return '';
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${masterSheetId}&range=A${rowIndex}`;
  return `=HYPERLINK("${url}","Open")`;
}

// Turn a URL into =HYPERLINK(...) formula if it looks valid
function makeHyperlinkFormulaOrBlank_(url, label) {
  const u = String(url || '').trim();
  if (!u || !/^https?:\/\//i.test(u)) return '';
  const safeLabel = (label || 'Open').replace(/"/g, '""');
  const safeUrl = u.replace(/"/g, '""');
  return `=HYPERLINK("${safeUrl}","${safeLabel}")`;
}

// Read headers and build a robust alias map -> column index
function buildHeaderMap_(headerRow, aliases) {
  const map = {};
  const normalized = headerRow.map(h => normalizeHeader_(h));
  const resolve = (key) => {
    const al = aliases[key];
    if (!al) return null;
    for (let c = 0; c < normalized.length; c++) {
      const h = normalized[c];
      if (!h) continue;
      // direct exact
      if (h === normalizeHeader_(key)) return c;
      // alias match
      for (const a of al) {
        if (h === normalizeHeader_(a)) return c;
      }
    }
    return null;
  };

  const keys = Object.keys(aliases);
  for (const k of keys) {
    const idx = resolve(k);
    if (idx != null) map[k] = idx;
  }
  return map;
}

function normalizeHeader_(s) {
  return String(s || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[^\w\s#]/g, '') // keep letters/digits/_/space/#
    .trim();
}

function normStr_(v) {
  return String(v == null ? '' : v).trim();
}

// Primary key per row: RootApptID → Email → Phone → Name
function pickPrimaryKey_(row, hmap) {
  const ra = hmap.RootApptID != null ? normStr_(row[hmap.RootApptID]) : '';
  if (ra) return 'root:' + ra;

  const em = hmap.Email != null ? normEmail_(row[hmap.Email]) : '';
  if (em) return 'email:' + em;

  const ph = hmap.Phone != null ? normPhone_(row[hmap.Phone]) : '';
  if (ph) return 'phone:' + ph;

  const nm = hmap.Customer != null ? normName_(row[hmap.Customer]) : '';
  if (nm) return 'name:' + nm;

  return null;
}

function normEmail_(v) {
  const s = String(v || '').trim().toLowerCase();
  return s || '';
}

function normPhone_(v) {
  const s = String(v || '').replace(/[^\d]/g, '');
  // Optional: strip leading country code 1 (US)
  return s.replace(/^1(?=\d{10}$)/, '');
}

function normName_(v) {
  return String(v || '').trim().toUpperCase();
}

function parseDate_(v, ss) {
  if (!v && v !== 0) return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return atMidnight_(v, ss);
  }
  const s = String(v).trim();
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d) ? null : atMidnight_(d, ss);
}

function atMidnight_(d, ss) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const y = Number(Utilities.formatDate(d, tz, 'yyyy'));
  const m = Number(Utilities.formatDate(d, tz, 'MM')) - 1;
  const day = Number(Utilities.formatDate(d, tz, 'dd'));
  return new Date(y, m, day);
}

function dateGt_(a, b) {
  if (!a && !b) return false;
  if (a && !b) return true;
  if (!a && b) return false;
  return a.getTime() > b.getTime();
}
function dateLt_(a, b) {
  if (!a && !b) return false;
  if (a && !b) return false;
  if (!a && b) return true;
  return a.getTime() < b.getTime();
}

function dateToCell_(d) {
  return d ? d : '';
}

function canonicalStage_(v) {
  const s = String(v || '').trim().toLowerCase();
  const simple = s.replace(/[^\w]/g, ''); // strip spaces/punct
  if (/^appointment$/.test(simple)) return 'Appointment';
  if (/^lead$/.test(simple)) return 'Lead';
  if (/^hotlead$/.test(simple)) return 'Hot Lead';
  if (/^followuprequired$/.test(simple) || /^followupreq$/.test(simple)) return 'Follow-Up Required';
  if (/^deposit$/.test(simple)) return 'Deposit';
  if (/^won$/.test(simple)) return 'Won';
  if (/^lostlead$/.test(simple) || /^lost$/.test(simple) || /^deadlead$/.test(simple)) return 'Lost Lead';
  return null;
}

/** ============================ OPTIONAL LOOKUPS ============================ */

// Build a map of SO# → payments summary { paidToDate, firstDepositDate, firstDepositAmount }
function buildPaymentsMapBySO_(ss) {
  const map = Object.create(null);
  const sheet = findFirstExistingSheet_(ss, CONFIG.OPTIONAL_SOURCES.AR_SHEET_CANDIDATES);
  if (!sheet) return map; // optional

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return map;

  const header = values[0].map(v => String(v || '').trim());
  const hnorm = header.map(normalizeHeader_);

  // Try to find SO and payment columns robustly
  const soIdx = findHeaderIndexByAliases_(header, ['SO Number', 'SO#', 'SO', 'Sales Order', 'SO No', 'SO No.']);
  if (soIdx == null) return map;

  const payDateIdxs = [];
  const payAmtIdxs = [];
  for (let c = 0; c < header.length; c++) {
    const h = header[c];
    if (/^payment date/i.test(h) || /date \d+$/i.test(h) && /payment/i.test(h)) payDateIdxs.push(c);
    if (/^payment amount/i.test(h) || /amount \d+$/i.test(h) && /payment/i.test(h)) payAmtIdxs.push(c);
  }

  // Fallback: also include any columns that look like "Deposit Date"/"Deposit Amount"
  for (let c = 0; c < header.length; c++) {
    const h = header[c];
    if (/deposit date/i.test(h)) payDateIdxs.push(c);
    if (/deposit amount/i.test(h)) payAmtIdxs.push(c);
  }

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const so = String(row[soIdx] || '').trim();
    if (!so) continue;

    let paid = 0;
    let firstDate = null;
    let firstAmt = null;

    // Collect payment pairs by index ordering
    const pairs = [];
    const n = Math.max(payDateIdxs.length, payAmtIdxs.length);
    for (let i = 0; i < n; i++) {
      const d = row[payDateIdxs[i]];
      const a = row[payAmtIdxs[i]];
      const dd = (d && Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d)) ? d : (d ? new Date(d) : null);
      const aa = toNumberOrNull_(a);
      if (dd && !isNaN(dd) && aa != null) {
        pairs.push({ d: dd, a: aa });
      }
    }

    // Sum all, pick earliest as "first deposit"
    for (const p of pairs) {
      paid += p.a;
      if (!firstDate || p.d < firstDate) {
        firstDate = p.d;
        firstAmt = p.a;
      }
    }

    if (paid || firstDate) {
      map[so] = {
        paidToDate: paid || null,
        firstDepositDate: firstDate ? atMidnight_(firstDate, ss) : null,
        firstDepositAmount: firstAmt || null
      };
    }
  }
  return map;
}

// Build a map of SO# → Order Total (best effort from Orders/Quotes sheets)
function buildOrderTotalsMapBySO_(ss) {
  const map = Object.create(null);

  const candidates = [
    ...CONFIG.OPTIONAL_SOURCES.ORDERS_SHEET_CANDIDATES,
    ...CONFIG.OPTIONAL_SOURCES.QUOTES_SHEET_CANDIDATES
  ];

  for (const name of candidates) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;

    const values = sh.getDataRange().getValues();
    if (!values || values.length < 2) continue;

    const header = values[0].map(v => String(v || '').trim());
    const soIdx = findHeaderIndexByAliases_(header, ['SO#', 'SO #', 'SO Number', 'SO', 'Sales Order', 'SO No', 'SO No.']);
    if (soIdx == null) continue;

    const totalIdx = findHeaderIndexByAliases_(header, ['Order Total', 'Quotation Amount', 'Total', 'Grand Total', 'Order Amount']);
    if (totalIdx == null) continue;

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const so = String(row[soIdx] || '').trim();
      if (!so) continue;

      const total = toNumberOrNull_(row[totalIdx]);
      if (total != null) map[so] = total;
    }
  }

  return map;
}

function toNumberOrNull_(v) {
  if (v == null || v === '') return null;
  const n = Number(v);
  return isFinite(n) ? n : null;
}

function findFirstExistingSheet_(ss, nameCandidates) {
  for (const name of nameCandidates) {
    const sh = ss.getSheetByName(name);
    if (sh) return sh;
  }
  return null;
}

function findHeaderIndexByAliases_(headerRow, nameCandidates) {
  const normalized = headerRow.map(h => normalizeHeader_(h));
  for (const cand of nameCandidates) {
    const n = normalizeHeader_(cand);
    const idx = normalized.indexOf(n);
    if (idx !== -1) return idx;
  }
  return null;
}

/** ================================ WRITE =================================== */

function writeRollup_(ss, headerOut, rows) {
  const sh = getOrCreateSheet_(ss, CONFIG.ROLLUP_SHEET_NAME);

  // Clear everything then write new content
  sh.clear({ contentsOnly: true });
  if (!headerOut.length) {
    sh.getRange(1, 1).setValue('No staged rows found.');
    return;
  }

  // Write header
  sh.getRange(1, 1, 1, headerOut.length).setValues([headerOut]);

  // Write body
  if (rows.length) {
    sh.getRange(2, 1, rows.length, headerOut.length).setValues(rows);
  }

  // === Apply consistent formatting & sizing (safe; values-only) ===
    applyRollupFormatting_(sh, 1 + rows.length, headerOut.length);
    SpreadsheetApp.getActive().toast('Client Stage Rollup refreshed.', 'Stage Rollup', 5);
}

function getOrCreateSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (sh) return sh;
  return ss.insertSheet(name);
}

// === UI/Formatting helper for the rollup (safe; values-only) =================
function applyRollupFormatting_(sh, totalRows, totalCols) {
  if (!sh || totalCols < 1) return;

  // ---- Theme knobs (edit freely) ----
  const THEME = {
    // ~3 lines at 10pt Arial. Increase to 58–62 if you want a tad more room.
    ROW_HEIGHT: 56,
    FONT_FAMILY: 'Arial',
    FONT_SIZE: 10,
    HEADER_BG: '#E8EAED',           // header row (row 1)
    SECTION_BG_FALLBACK: '#F3F4F6', // used if a stage name isn't recognized
    DATE_FORMAT: 'yyyy-mm-dd',
    MONEY_FORMAT: '#,##0.00',
    // Column widths for the 17 rollup columns (adjust any you like)
    COLUMN_WIDTHS: [140, 190, 160, 70, 110, 130, 140, 180, 520, 70, 120, 120, 120, 120, 120, 120, 130],

    // 🎨 Section header colors (close to your chips)
    STAGE_COLORS: {
      'Appointment': '#AECBFA',          // blue
      'Lead': '#FFD7C2',                 // peach
      'Hot Lead': '#D93025',             // red
      'Follow-Up Required': '#C5221F',   // deeper red
      'Deposit': '#C8E6C9',              // light green
      'Won': '#34A853',                  // green
      'Lost Lead': '#5F6368'             // dark grey
    },
    // Text color on the header strip for contrast
    STAGE_TEXT: {
      'Appointment': '#202124',
      'Lead': '#202124',
      'Hot Lead': '#FFFFFF',
      'Follow-Up Required': '#FFFFFF',
      'Deposit': '#202124',
      'Won': '#FFFFFF',
      'Lost Lead': '#FFFFFF'
    },
    // How much lighter should the table rows be vs. the header?
    // We compute two tints by blending the header color with white.
    // 0.86 and 0.92 give a pleasant “just lighter” effect.
    TINT_1: 0.86,
    TINT_2: 0.92
  };

  const lastRow = Math.max(1, totalRows);
  const lastCol = totalCols;

  // NEW: reset any stale formatting in the body (exclude row 1 header).
  // This avoids cases where a previous section header's bold/white text
  // sticks to the first data row after sections shift between refreshes.
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, lastCol).clearFormat();
  }

  // Freeze header, clear banding, uniform row height
  sh.setFrozenRows(1);
  sh.getBandings().forEach(b => b.remove());
  try { sh.setRowHeights(1, lastRow, THEME.ROW_HEIGHT); } catch (e) {}

  // Whole area base formatting
  const whole = sh.getRange(1, 1, lastRow, lastCol);
  whole
    .setFontFamily(THEME.FONT_FAMILY)
    .setFontSize(THEME.FONT_SIZE)
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // keep rows uniform

  // Header row styling
  const head = sh.getRange(1, 1, 1, lastCol);
  head.setFontWeight('bold').setBackground(THEME.HEADER_BG);

  // Column widths
  for (let c = 1; c <= Math.min(lastCol, THEME.COLUMN_WIDTHS.length); c++) {
    sh.setColumnWidth(c, THEME.COLUMN_WIDTHS[c - 1]);
  }

  // --- Build a quick header map for targeted formatting (dates/money/Next Steps)
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).toLowerCase());
  const idxNextSteps = headers.indexOf('next steps');
  const idxDatePast  = headers.indexOf('visit date (past)');
  const idxDateNext  = headers.indexOf('next visit (scheduled)');
  const idxDepDate   = headers.indexOf('first deposit date');
  const moneyCols = [
    headers.indexOf('first deposit amount'),
    headers.indexOf('order total'),
    headers.indexOf('paid-to-date'),
    headers.indexOf('outstanding balance')
  ].filter(i => i !== -1);

  // --- Color-coded section headers + per-section tinted banding
  if (lastRow > 1) {
    const aVals = sh.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0] || '').trim());
    const endColA1 = columnLetter_(lastCol);

    // Collect absolute header row indices like [2, 10, 15, ...]
    const headerRows = [];
    for (let i = 0; i < aVals.length; i++) {
      const v = aVals[i];
      if (v.startsWith('— ') && v.endsWith(' —')) headerRows.push(i + 2);
    }
    // For each header, color its strip and tint the block below it up to the next header.
    for (let h = 0; h < headerRows.length; h++) {
      const rHeader = headerRows[h];
      const nextHeader = (h + 1 < headerRows.length) ? headerRows[h + 1] : (lastRow + 1);

      // Extract stage from "— Stage —"
      const raw = sh.getRange(rHeader, 1).getDisplayValue().trim();
      const stage = raw.slice(2, -2).trim();

      const headerBg = THEME.STAGE_COLORS[stage] || THEME.SECTION_BG_FALLBACK;
      const headerFg = THEME.STAGE_TEXT[stage] || '#202124';

      // Paint the header strip
      sh.getRange(`A${rHeader}:${endColA1}${rHeader}`)
        .setBackground(headerBg)
        .setFontColor(headerFg)
        .setFontWeight('bold')
        .setHorizontalAlignment('left')
        .setBorder(true, null, null, null, null, null, '#D0D7DE', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

      // Compute data block (skip header row; skip trailing spacer row we create)
      let rStart = rHeader + 1;
      let rEnd   = nextHeader - 2; // -1 for spacer row, -1 to make inclusive
      if (rEnd < rStart) continue;

      // Build alternating backgrounds lighter than header
      const tint1 = blendWithWhite_(headerBg, THEME.TINT_1);
      const tint2 = blendWithWhite_(headerBg, THEME.TINT_2);
      const rows = rEnd - rStart + 1;

      const bgBlock = new Array(rows);
      for (let i = 0; i < rows; i++) {
        const color = (i % 2 === 0) ? tint1 : tint2;
        const rowArr = new Array(lastCol);
        rowArr.fill(color);
        bgBlock[i] = rowArr;
      }
      sh.getRange(rStart, 1, rows, lastCol).setBackgrounds(bgBlock);
    }
  }

  // Targeted number formats (dates & money)
  if (lastRow > 1) {
    const rowsCount = lastRow - 1;
    if (idxDatePast !== -1) sh.getRange(2, idxDatePast + 1, rowsCount, 1).setNumberFormat(THEME.DATE_FORMAT);
    if (idxDateNext !== -1) sh.getRange(2, idxDateNext + 1, rowsCount, 1).setNumberFormat(THEME.DATE_FORMAT);
    if (idxDepDate  !== -1) sh.getRange(2, idxDepDate  + 1, rowsCount, 1).setNumberFormat(THEME.DATE_FORMAT);
    for (const i of moneyCols) sh.getRange(2, i + 1, rowsCount, 1).setNumberFormat(THEME.MONEY_FORMAT).setHorizontalAlignment('right');

    // Next Steps: show the FIRST lines only (clip) and align to top
    if (idxNextSteps !== -1) {
      sh.getRange(2, idxNextSteps + 1, rowsCount, 1)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
        .setVerticalAlignment('top');
    }
  }
}

/* --- color utils --- */
function blendWithWhite_(hex, alpha) {
  // alpha is the weight of WHITE (0..1). 0.90 => 90% white + 10% color.
  const c = hexToRgb_(hex);
  const r = Math.round(255 * alpha + c.r * (1 - alpha));
  const g = Math.round(255 * alpha + c.g * (1 - alpha));
  const b = Math.round(255 * alpha + c.b * (1 - alpha));
  return rgbToHex_(r, g, b);
}
function hexToRgb_(hex) {
  const h = String(hex || '').replace('#', '');
  const n = h.length === 3 ? h.split('').map(x => x + x).join('') : h;
  const num = parseInt(n, 16);
  return { r: (num >> 16) & 255, g: (num >> 8) & 255, b: num & 255 };
}
function rgbToHex_(r, g, b) {
  const to2 = (n) => n.toString(16).padStart(2, '0');
  return `#${to2(r)}${to2(g)}${to2(b)}`;
}



// A1 helpers
function columnLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - 1) / 26 | 0; }
  return s;
}

