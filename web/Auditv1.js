/*****  AuditV1.gs — Master Appointments Data Quality Audit (v1)
 *  Creates menu: Audit → Run Master Audit (v1)
 *  Writes findings to a sheet called: "Audit Findings (v1)"
 *  Targets: 00_Master Appointments
 *  Policy: Minimum deposit = 20% of Order Total
 *****/

const AUDIT_CONFIG = {
  MASTER_SHEET_NAME: '00_Master Appointments',
  AUDIT_SHEET_NAME: 'Audit Findings (v1)',
  MIN_DEPOSIT_PCT: 0.20, // 20%
  // Statuses where financial fields MUST be present & coherent
  STATUSES_REQUIRE_FINANCIALS: new Set(['Deposit Paid', 'Order In Progress', 'Order Completed']),
  // Skip canceled or explicitly inactive rows
  EXCLUDE_IF_STATUS: new Set(['Canceled']),
  ACTIVE_COL_TREATED_NO: 'No', // rows with Active? = "No" are skipped
};

const FIELD = {
  APPT_ID: 'APPT_ID',
  ROOT_APPT_ID: 'RootApptID',
  SALES_STAGE: 'Sales Stage',
  CONV_STATUS: 'Conversion Status',
  ORDER_STATUS: 'Custom Order Status',
  ASSIGNED_REP: 'Assigned Rep',
  ASSISTED_REP: 'Assisted Rep',
  CUSTOMER_NAME: 'Customer Name',
  SO: 'SO#',
  ORDER_TOTAL: 'Order Total',
  PAID_TO_DATE: 'Paid-to-Date',
  REM_BAL: 'Remaining Balance',
  LAST_PAY_DATE: 'Last Payment Date',
  DEADLINE_3D: '3D Deadline',
  DEADLINE_PROD: 'Production Deadline',
  NEXT_STEPS: 'Next Steps',
  UPDATED_AT: 'Updated At',
  UPDATED_BY: 'Updated By',
  STATUS: 'Status',
  ACTIVE: 'Active?',
  QUOTE_URL: 'Quotation URL',
};

const OUTPUT_HEADERS = [
  'Timestamp',
  'APPT_ID',
  'RootApptID',
  'SO#',
  'Customer Name',
  'Sales Stage',
  'Conversion Status',
  'Custom Order Status',
  'Assigned Rep',      // <— renamed from Owner
  'Assisted Rep',      // <— added next to Assigned Rep
  'Issue',
  'Action',
  'Severity',
  'Go to',
  'Updated At',
  'Updated By'
];


function runMasterAuditV1() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(AUDIT_CONFIG.MASTER_SHEET_NAME);
  if (!master) throw new Error(`Sheet not found: ${AUDIT_CONFIG.MASTER_SHEET_NAME}`);

  // Read all data
  const range = master.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return writeAudit_([], ss);

  const headers = values[0].map(String);
  const H = headerIndexMap_(headers);

  // Validate required columns exist
  const required = [
    FIELD.APPT_ID, FIELD.ROOT_APPT_ID, FIELD.SALES_STAGE, FIELD.CONV_STATUS,
    FIELD.ORDER_STATUS, FIELD.ASSIGNED_REP, FIELD.ASSISTED_REP, FIELD.CUSTOMER_NAME,
    FIELD.SO, FIELD.ORDER_TOTAL, FIELD.PAID_TO_DATE, FIELD.REM_BAL, FIELD.LAST_PAY_DATE,
    FIELD.DEADLINE_3D, FIELD.DEADLINE_PROD, FIELD.NEXT_STEPS, FIELD.UPDATED_AT, FIELD.UPDATED_BY,
    FIELD.STATUS, FIELD.ACTIVE, FIELD.QUOTE_URL
  ];
  required.forEach(name => {
    if (!(name in H)) {
      // Not all are truly mandatory for every rule, so don't hard-fail; warn in logs.
      // Logger.log(`Warning: Column not found: ${name}`);
    }
  });

  const findings = [];
  const now = new Date();
  const sheetId = master.getSheetId();
  const baseUrl = ss.getUrl();

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    // Skip canceled or inactive rows
    const status = readString_(row, H[FIELD.STATUS]);
    const activeVal = readString_(row, H[FIELD.ACTIVE]);
    if (AUDIT_CONFIG.EXCLUDE_IF_STATUS.has(status)) continue;
    if (activeVal === AUDIT_CONFIG.ACTIVE_COL_TREATED_NO) continue;

    const conv = readString_(row, H[FIELD.CONV_STATUS]);
    const requireFinancials = AUDIT_CONFIG.STATUSES_REQUIRE_FINANCIALS.has(conv);

    const orderTotal = readNumber_(row, H[FIELD.ORDER_TOTAL]);
    const paid = readNumber_(row, H[FIELD.PAID_TO_DATE]);
    const remBal = readNumber_(row, H[FIELD.REM_BAL]);
    const lastPayDate = readAny_(row, H[FIELD.LAST_PAY_DATE]); // could be Date or blank
    const so = readString_(row, H[FIELD.SO]);
    const cust = readString_(row, H[FIELD.CUSTOMER_NAME]);
    const salesStage = readString_(row, H[FIELD.SALES_STAGE]);
    const orderStatus = readString_(row, H[FIELD.ORDER_STATUS]);
    const assignedRep = readString_(row, H[FIELD.ASSIGNED_REP]);
    const assistedRep = readString_(row, H[FIELD.ASSISTED_REP]);
    const owner = assignedRep || assistedRep || '(Unassigned)';
    const appt = readString_(row, H[FIELD.APPT_ID]);
    const root = readString_(row, H[FIELD.ROOT_APPT_ID]);
    const updatedAt = readAny_(row, H[FIELD.UPDATED_AT]);
    const updatedBy = readString_(row, H[FIELD.UPDATED_BY]);

    // Helper to push a finding row
    const pushFinding = (issue, action, severity, gotoColHeader) => {
      const cIndex = H[gotoColHeader];
      const link = cIndex != null
        ? makeCellLink_(baseUrl, sheetId, r + 1, cIndex + 1, `Go to ${gotoColHeader}`)
        : '';
      findings.push([
        now,
        appt,
        root,
        so,
        cust,
        salesStage,
        conv,
        orderStatus,
        assignedRep || '(Unassigned)',   // Assigned Rep
        assistedRep,                     // Assisted Rep
        issue,
        action,
        severity,
        link,
        updatedAt,
        updatedBy
      ]);
    };

    // ===== Rule set (v1) =====

    // A) Deposit & later: require financials
    if (requireFinancials) {
      // R1: Missing Order Total (Critical)
      if (!isFinite(orderTotal)) {
        pushFinding(
          'Missing Order Total',
          `Enter Order Total via Record Payments.`,
          'Critical',
          FIELD.ORDER_TOTAL
        );
      }

      // R2: Deposit Paid but no payment fields (Critical)
      if (conv === 'Deposit Paid') {
        if (!isFinite(paid) || !lastPayDate) {
          const which = [
            !isFinite(paid) ? 'Paid-to-Date' : null,
            !lastPayDate ? 'Last Payment Date' : null
          ].filter(Boolean).join(', ');
          pushFinding(
            `Deposit Paid but missing: ${which}`,
            'Record the deposit amount (Paid-to-Date) and set Last Payment Date.',
            'Critical',
            !isFinite(paid) ? FIELD.PAID_TO_DATE : FIELD.LAST_PAY_DATE
          );
        }
      }

      // R3: Completed but balance not zero (Critical)
      if (conv === 'Order Completed') {
        if (isFinite(remBal) && remBal > 0) {
          pushFinding(
            `Order Completed but Remaining Balance > 0 (${remBal})`,
            'Collect final payment or correct Paid-to-Date / Remaining Balance.',
            'Critical',
            FIELD.REM_BAL
          );
        }
      }

      // R4: Missing SO# for paid/production/completed (Critical)
      if (!so) {
        pushFinding(
          'Missing SO#',
          'Add the Sales Order number (SO#) to link payments and production.',
          'Critical',
          FIELD.SO
        );
      }

      // R5: Payment math mismatch (Warning): Order Total - Paid = Remaining
      if (isFinite(orderTotal) && isFinite(paid) && isFinite(remBal)) {
        const diff = round2_(orderTotal - paid - remBal);
        if (diff !== 0) {
          pushFinding(
            `Payment math mismatch (Order Total - Paid-to-Date - Remaining ≠ 0)`,
            'Correct Order Total / Paid-to-Date / Remaining Balance so they reconcile.',
            'Warning',
            FIELD.REM_BAL
          );
        }
      }

      // R6: Overpayment anomaly (Warning)
      if (isFinite(orderTotal) && isFinite(paid) && paid > orderTotal) {
        pushFinding(
          `Overpayment: Paid-to-Date > Order Total`,
          'Verify totals or correct Order Total.',
          'Warning',
          FIELD.PAID_TO_DATE
        );
      }

      // R7: Minimum deposit policy 20% (Critical)
      if (isFinite(orderTotal) && isFinite(paid)) {
        const minRequired = orderTotal * AUDIT_CONFIG.MIN_DEPOSIT_PCT;
        if (paid < minRequired) {
          pushFinding(
            `Deposit below policy (${formatMoney_(paid)} < ${formatMoney_(minRequired)})`,
            `Collect additional deposit (≥ ${Math.round(AUDIT_CONFIG.MIN_DEPOSIT_PCT * 100)}% of Order Total) or update totals if misrecorded.`,
            'Critical',
            FIELD.PAID_TO_DATE
          );
        }
      }
    }

    // B) Order In Progress — deadlines by sub-status (Warnings)
    if (conv === 'Order In Progress') {
      if (/3D/i.test(orderStatus)) {
        const d3 = readAny_(row, H[FIELD.DEADLINE_3D]);
        if (!d3) {
          pushFinding(
            'Missing 3D Deadline',
            'Enter 3D Deadline while in 3D-related stage.',
            'Warning',
            FIELD.DEADLINE_3D
          );
        }
      }
      if (/Prod|Production/i.test(orderStatus)) {
        const dp = readAny_(row, H[FIELD.DEADLINE_PROD]);
        if (!dp) {
          pushFinding(
            'Missing Production Deadline',
            'Enter Production Deadline while in production stage.',
            'Warning',
            FIELD.DEADLINE_PROD
          );
        }
      }
    }

    // C) Quotation hygiene (Warnings)
    if (conv === 'Quotation Requested' || conv === 'Quotation Sent') {
      const q = readString_(row, H[FIELD.QUOTE_URL]);
      if (!q) {
        pushFinding(
          'Missing Quotation URL',
          'Add Quotation URL to document what was sent.',
          'Warning',
          FIELD.QUOTE_URL
        );
      }
    }

    // D) Follow-up hygiene (Warnings)
    if (typeof conv === 'string' && conv.startsWith('Follow-Up')) {
      const ns = readString_(row, H[FIELD.NEXT_STEPS]);
      if (!ns) {
        pushFinding(
          'Follow-up status but Next Steps missing',
          'Fill Next Steps so anyone can pick up the thread.',
          'Warning',
          FIELD.NEXT_STEPS
        );
      }
    }
  }

  writeAudit_(findings, ss);
  SpreadsheetApp.getUi().alert(`Audit complete. Findings: ${findings.length}`);
}

/* ----------------- Helpers ----------------- */

function headerIndexMap_(headers) {
  const map = {};
  headers.forEach((h, i) => { map[h] = i; });
  return map;
}

function readString_(row, idx) {
  if (idx == null) return '';
  const v = row[idx];
  return (v == null) ? '' : String(v).trim();
}

function readNumber_(row, idx) {
  if (idx == null) return NaN;
  const v = row[idx];
  if (v === '' || v == null) return NaN;
  const n = (typeof v === 'number') ? v : Number(String(v).replace(/[^0-9.\-]/g, ''));
  return Number.isFinite(n) ? n : NaN;
}

function readAny_(row, idx) {
  if (idx == null) return '';
  return row[idx];
}

function round2_(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

function formatMoney_(n) {
  if (!Number.isFinite(n)) return '';
  return '$' + n.toFixed(2);
}

function makeCellLink_(fileUrl, sheetId, rowIdx1, colIdx1, label) {
  // Apps Script sheet links use #gid=<id>&range=A1
  const a1 = colToLetter_(colIdx1) + String(rowIdx1);
  const url = `${fileUrl}#gid=${sheetId}&range=${encodeURIComponent(a1)}`;
  // Return a =HYPERLINK() formula so it’s clickable in Sheets
  return `=HYPERLINK("${url}","${label}")`;
}

function colToLetter_(col) {
  // col is 1-based
  let temp = '', letter = '';
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

function writeAudit_(rows, ss) {
  const out = ss.getSheetByName(AUDIT_CONFIG.AUDIT_SHEET_NAME) ||
              ss.insertSheet(AUDIT_CONFIG.AUDIT_SHEET_NAME);

  // 0) Remove any existing BASIC filter & banding (avoid conflicts)
  try { const f = out.getFilter(); if (f) f.remove(); } catch (_) {}
  try { (out.getBandings() || []).forEach(b => b.remove()); } catch(_) {}

  // 1) Clear and write header
  out.clear();
  out.getRange(1, 1, 1, OUTPUT_HEADERS.length)
     .setValues([OUTPUT_HEADERS])
     .setFontWeight('bold')
     .setFontSize(11)
     .setVerticalAlignment('middle')
     .setHorizontalAlignment('center')
     .setBackground('#1F2937')   // slate-800
     .setFontColor('#FFFFFF')
     .setBorder(true, true, true, true, false, false, '#111827', SpreadsheetApp.BorderStyle.SOLID);

  // 2) Write body rows (if any)
  if (rows.length) {
    out.getRange(2, 1, rows.length, OUTPUT_HEADERS.length)
       .setValues(rows)
       .setFontSize(10)
       .setVerticalAlignment('middle');
  }

  // 3) Freeze header, compact heights
  out.setFrozenRows(1);
  const totalRows = Math.max(2, rows.length + 1);
  out.setRowHeights(1, 1, 26);
  if (rows.length) out.setRowHeights(2, rows.length, 22);

  // 4) Column sizing & wrapping
  const H = OUTPUT_HEADERS.reduce((m,h,i)=>{ m[h]=i+1; return m; }, {});
  [
    'Timestamp','APPT_ID','RootApptID','SO#',
    'Customer Name','Sales Stage','Conversion Status','Custom Order Status','Assigned Rep'
  ].forEach(h => { if (H[h]) out.autoResizeColumn(H[h]); });

  const setW = (name, w) => { if (H[name]) out.setColumnWidth(H[name], w); };
  setW('Assisted Rep', 140);
  setW('Issue',        320);
  setW('Action',       320);
  setW('Severity',     110);
  setW('Go to',        110);
  setW('Updated At',   140);
  setW('Updated By',   160);

  // Wrap long text in Issue/Action
  ['Issue','Action'].forEach(h => {
    if (H[h]) out.getRange(2, H[h], Math.max(1, rows.length), 1).setWrap(true);
  });

  // Make links easier to click in "Go to"
  if (H['Go to']) {
    out.getRange(2, H['Go to'], Math.max(1, rows.length), 1)
       .setHorizontalAlignment('left')
       .setFontStyle('italic');
  }

  // Center smaller columns
  ['Severity','SO#'].forEach(h => {
    if (H[h]) out.getRange(2, H[h], Math.max(1, rows.length), 1).setHorizontalAlignment('center');
  });

  // Date/time formats
  const dateFmt = 'yyyy-MM-dd HH:mm';
  if (H['Timestamp'])  out.getRange(2, H['Timestamp'],  Math.max(1, rows.length), 1).setNumberFormat(dateFmt);
  if (H['Updated At']) out.getRange(2, H['Updated At'], Math.max(1, rows.length), 1).setNumberFormat(dateFmt);

  // 5) Light zebra banding for readability (body only)
  if (rows.length) {
    const banding = out
      .getRange(2, 1, rows.length, OUTPUT_HEADERS.length)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    banding.setFirstRowColor('#FFFFFF');
    banding.setSecondRowColor('#F7F8FA');
  }

  // 6) Conditional formatting (row-level highlight by Severity)
  out.setConditionalFormatRules([]);
  if (rows.length && H['Severity']) {
    const allData = out.getRange(2, 1, rows.length, OUTPUT_HEADERS.length);
    const sevColLetter = (function colLetter_(i){
      let col=i, s=''; while (col>0){ const m=(col-1)%26; s=String.fromCharCode(65+m)+s; col=(col-1-m)/26; } return s;
    })(H['Severity']);
    const rules = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=$${sevColLetter}2="Critical"`)
        .setBackground('#FDECEA')
        .setFontColor('#B71C1C')
        .setRanges([allData])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=$${sevColLetter}2="Warning"`)
        .setBackground('#FFF8E1')
        .setFontColor('#8D6E00')
        .setRanges([allData])
        .build()
    ];
    out.setConditionalFormatRules(rules);
  }

  // 7) Thin dotted row separators (optional)
  if (rows.length) {
    out.getRange(2, 1, rows.length, OUTPUT_HEADERS.length)
       .setBorder(false, false, false, false, true, false, '#E5E7EB', SpreadsheetApp.BorderStyle.DOTTED);
  }

  // 8) ✅ Sort ONLY the data rows (row 2..N), not the header — do this BEFORE creating the filter
  try {
    if (rows.length) {
      const sortSpec = [];
      if (H['Severity'])     sortSpec.push({column: H['Severity'],     ascending: true});   // "Critical" < "Warning"
      if (H['Assigned Rep']) sortSpec.push({column: H['Assigned Rep'], ascending: true});
      if (H['Timestamp'])    sortSpec.push({column: H['Timestamp'],    ascending: false});
      if (sortSpec.length) {
        out.getRange(2, 1, rows.length, OUTPUT_HEADERS.length).sort(sortSpec);
      }
    }
  } catch (_) {}

  // 9) Create BASIC filter (after sorting)
  if (rows.length) {
    out.getRange(2, 1, rows.length, OUTPUT_HEADERS.length).createFilter();
  }
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
