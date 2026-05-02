// ============================================================
// payment_report_server.gs  —  Project #11
// 400 Payments Ledger  ·  Exact column mapping (A → AU)
// ============================================================

// ── Column index map (0-based) ───────────────────────────────
var PR_COL = {
  PAYMENT_ID:         0,   // A
  Brand:              1,   // B
  RootApptID:         2,   // C
  SO:                 3,   // D  (SO#)
  BasketID:           4,   // E
  DocType:            5,   // F
  DocNumber:          6,   // G
  PaymentDateTime:    7,   // H
  Method:             8,   // I
  Reference:          9,   // J
  Notes:              10,  // K
  AmountGross:        11,  // L
  FeePercent:         12,  // M
  FeeAmount:          13,  // N
  AmountNet:          14,  // O
  AllocatedToSO:      15,  // P
  LinesJSON:          16,  // Q
  Subtotal:           17,  // R
  DepositPaidToDate:  18,  // S
  Order_Total_SO:     19,  // T
  Balance_SO:         20,  // U
  PmtHistory:         21,  // V
  SubmittedBy:        22,  // W
  SubmittedDateTime:  23,  // X
  AnchorType:         24,  // Y
  PaidToDate_SO:      25,  // Z
  DocFileID:          26,  // AA
  DocPDFID:           27,  // AB
  DocURL:             28,  // AC
  PDFURL:             29,  // AD
  OrderTotalSet:      30,  // AE
  OrderTotalValue:    31,  // AF
  OrderTotalSource:   32,  // AG
  OrderTotalTarget:   33,  // AH
  OrderTotalOldValue: 34,  // AI
  ARShortcutID:       35,  // AJ
  ARShortcutURL:      36,  // AK
  DocStatus:          37,  // AL
  DocRole:            38,  // AM
  SupersedesDoc:      39,  // AN
  AppliesToDoc:       40,  // AO
  ReplacedByDoc:      41,  // AP
  TaxRate:            42,  // AQ
  TaxAmount:          43,  // AR
  InvoiceTotal:       44,  // AS
  BalanceDue:         45,  // AT
  TaxEnabled:         46   // AU
};
var PR_TOTAL_COLS = 47;

// ── Ledger helpers ───────────────────────────────────────────

function pr_getLedger_() {
  var id = (PropertiesService.getScriptProperties().getProperty('LEDGER_FILE_ID') || '').trim();
  if (!id) throw new Error('LEDGER_FILE_ID not set. Run Sales → Seed Config first.');
  return SpreadsheetApp.openById(id);
}

function pr_getPaymentsSheet_(ledger) {
  var sheets = ledger.getSheets();
  var hit = sheets.filter(function(s) { return /payment/i.test(s.getName()); })[0];
  return hit || sheets[0];
}

// ── Open dialog ──────────────────────────────────────────────

function openPaymentReportDialog() {
  var html = HtmlService.createHtmlOutputFromFile('dlg_payment_report_v1')
    .setWidth(1100).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, '💵 Payment Report');
}

// ── Filter options for dropdowns ─────────────────────────────

function pr_getFilterOptions() {
  try {
    var ledger = pr_getLedger_();
    var sh     = pr_getPaymentsSheet_(ledger);
    var last   = sh.getLastRow();
    if (last < 2) return { ok:true, brands:[], methods:[], docTypes:[], docStatuses:[], submittedBys:[] };

    var data = sh.getRange(2, 1, last - 1, PR_TOTAL_COLS).getValues();

    var distinct = function(colIdx) {
      var seen = {}, out = [];
      data.forEach(function(row) {
        var v = String(row[colIdx] || '').trim();
        if (v && !seen[v]) { seen[v] = true; out.push(v); }
      });
      return out.sort();
    };

    return {
      ok:           true,
      brands:       distinct(PR_COL.Brand),
      methods:      distinct(PR_COL.Method),
      docTypes:     distinct(PR_COL.DocType),
      docStatuses:  distinct(PR_COL.DocStatus),
      submittedBys: distinct(PR_COL.SubmittedBy)
    };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Customer name lookup (from 00_Master Appointments) ───────
// Key = APPT_ID (col A). Customer Name = col P.
// RootApptID in this sheet is a Calendly username, NOT an AP- ID.

function pr_buildCustomerMap_() {
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName('00_Master Appointments');
    if (!sh) { Logger.log('pr_buildCustomerMap_: sheet not found'); return {}; }

    var lastCol = sh.getLastColumn();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return {};

    var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var H = {};
    headers.forEach(function(h, i) { var k = String(h||'').trim(); if(k) H[k] = i; });

    // APPT_ID is col A (index 0) — primary key
    var apptCol = H['APPT_ID'] !== undefined ? H['APPT_ID'] : 0;

    // Customer Name is col P (index 15)
    var nameCol = H['Customer Name'] !== undefined ? H['Customer Name']
                : H['CustomerName']  !== undefined ? H['CustomerName']
                : H['Client Name']   !== undefined ? H['Client Name']  : 15;

    var data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var map  = {};

    data.forEach(function(row) {
      var id   = String(row[apptCol] || '').trim().replace(/^'/, '');
      var name = String(row[nameCol] || '').trim();
      if (id && name && !map[id]) map[id] = name;
    });

    Logger.log('pr_buildCustomerMap_: ' + Object.keys(map).length + ' entries, apptCol=' + apptCol + ' nameCol=' + nameCol);
    return map;
  } catch(e) {
    Logger.log('pr_buildCustomerMap_ error: ' + e.message);
    return {};
  }
}

// Lookup: exact match on APPT_ID, fallback case-insensitive
function pr_lookupCustomer_(customerMap, rawId) {
  if (!rawId) return '';
  rawId = String(rawId).trim().replace(/^'/, '');

  // 1. Exact match
  if (customerMap[rawId]) return customerMap[rawId];

  // 2. Case-insensitive fallback
  var lower = rawId.toLowerCase();
  var keys  = Object.keys(customerMap);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].toLowerCase() === lower) return customerMap[keys[i]] || '';
  }
  return '';
}

// ── Core report generator ────────────────────────────────────

function pr_generatePaymentReport(filters) {
  try {
    filters = filters || {};
    var ledger = pr_getLedger_();
    var sh     = pr_getPaymentsSheet_(ledger);
    var last   = sh.getLastRow();
    if (last < 2) return { ok: true, rows: [], summary: _pr_emptySummary_() };

    // Build customer name lookup ONCE before iterating rows
    var customerMap = pr_buildCustomerMap_();

    var data = sh.getRange(2, 1, last - 1, PR_TOTAL_COLS).getValues();

    var dateFrom = filters.dateFrom ? new Date(filters.dateFrom + 'T00:00:00') : null;
    var dateTo   = filters.dateTo   ? new Date(filters.dateTo   + 'T23:59:59') : null;

    var isActive = function(arr) { return arr && arr.length > 0; };
    var matchArr = function(arr, val) {
      var v = String(val || '').trim().toLowerCase();
      return arr.some(function(f){ return String(f).trim().toLowerCase() === v; });
    };

    var tz   = Session.getScriptTimeZone() || 'America/Los_Angeles';
    var fmtDT = function(v) {
      if (!v) return '';
      var d = v instanceof Date ? v : new Date(v);
      return isNaN(d) ? String(v) : Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm');
    };
    var fmtN = function(v) { return (v === '' || v == null) ? '' : (Number(v) || 0); };
    var fmtS = function(v) { return String(v == null ? '' : v).trim(); };

    var results = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];

      // Date filter on PaymentDateTime
      if (dateFrom || dateTo) {
        var raw = row[PR_COL.PaymentDateTime];
        var d   = raw instanceof Date ? raw : (raw ? new Date(raw) : null);
        if (!d || isNaN(d)) continue;
        if (dateFrom && d < dateFrom) continue;
        if (dateTo   && d > dateTo)   continue;
      }

      if (isActive(filters.brands)       && !matchArr(filters.brands,       row[PR_COL.Brand]))       continue;
      if (isActive(filters.methods)      && !matchArr(filters.methods,       row[PR_COL.Method]))      continue;
      if (isActive(filters.docTypes)     && !matchArr(filters.docTypes,      row[PR_COL.DocType]))     continue;
      if (isActive(filters.docStatuses)  && !matchArr(filters.docStatuses,   row[PR_COL.DocStatus]))   continue;
      if (isActive(filters.submittedBys) && !matchArr(filters.submittedBys,  row[PR_COL.SubmittedBy])) continue;

      results.push({
        ledgerRow:         i + 2,
        PAYMENT_ID:        fmtS(row[PR_COL.PAYMENT_ID]),
        Brand:             fmtS(row[PR_COL.Brand]),
        RootApptID:        fmtS(row[PR_COL.RootApptID]),
        CustomerName:      pr_lookupCustomer_(customerMap, fmtS(row[PR_COL.RootApptID])),
        SO:                fmtS(row[PR_COL.SO]),
        BasketID:          fmtS(row[PR_COL.BasketID]),
        DocType:           fmtS(row[PR_COL.DocType]),
        DocNumber:         fmtS(row[PR_COL.DocNumber]),
        PaymentDateTime:   fmtDT(row[PR_COL.PaymentDateTime]),
        Method:            fmtS(row[PR_COL.Method]),
        Reference:         fmtS(row[PR_COL.Reference]),
        Notes:             fmtS(row[PR_COL.Notes]),
        AmountGross:       fmtN(row[PR_COL.AmountGross]),
        FeePercent:        fmtN(row[PR_COL.FeePercent]),
        FeeAmount:         fmtN(row[PR_COL.FeeAmount]),
        AmountNet:         fmtN(row[PR_COL.AmountNet]),
        AllocatedToSO:     fmtS(row[PR_COL.AllocatedToSO]),
        Subtotal:          fmtN(row[PR_COL.Subtotal]),
        DepositPaidToDate: fmtN(row[PR_COL.DepositPaidToDate]),
        Order_Total_SO:    fmtN(row[PR_COL.Order_Total_SO]),
        Balance_SO:        fmtN(row[PR_COL.Balance_SO]),
        SubmittedBy:       fmtS(row[PR_COL.SubmittedBy]),
        SubmittedDateTime: fmtDT(row[PR_COL.SubmittedDateTime]),
        AnchorType:        fmtS(row[PR_COL.AnchorType]),
        PaidToDate_SO:     fmtN(row[PR_COL.PaidToDate_SO]),
        DocStatus:         fmtS(row[PR_COL.DocStatus]),
        DocRole:           fmtS(row[PR_COL.DocRole]),
        TaxRate:           fmtN(row[PR_COL.TaxRate]),
        TaxAmount:         fmtN(row[PR_COL.TaxAmount]),
        InvoiceTotal:      fmtN(row[PR_COL.InvoiceTotal]),
        BalanceDue:        fmtN(row[PR_COL.BalanceDue]),
        TaxEnabled:        fmtS(row[PR_COL.TaxEnabled]),
        DocURL:            fmtS(row[PR_COL.DocURL]),
        PDFURL:            fmtS(row[PR_COL.PDFURL])
      });
    }

    var sumK = function(k) { return results.reduce(function(a,r){ return a + (Number(r[k])||0); }, 0); };
    var summary = {
      total:             results.length,
      AmountGross:       sumK('AmountGross'),
      FeeAmount:         sumK('FeeAmount'),
      AmountNet:         sumK('AmountNet'),
      TaxAmount:         sumK('TaxAmount'),
      BalanceDue:        sumK('BalanceDue'),
      InvoiceTotal:      sumK('InvoiceTotal'),
      DepositPaidToDate: sumK('DepositPaidToDate'),
      Order_Total_SO:    sumK('Order_Total_SO')
    };

    return { ok: true, rows: results, summary: summary };
  } catch(e) {
    Logger.log('pr_generatePaymentReport: ' + e.stack);
    return { ok: false, error: e.message, rows: [], summary: _pr_emptySummary_() };
  }
}

function _pr_emptySummary_() {
  return { total:0, AmountGross:0, FeeAmount:0, AmountNet:0,
           TaxAmount:0, BalanceDue:0, InvoiceTotal:0, DepositPaidToDate:0, Order_Total_SO:0 };
}

// ── Utility for dialog: build filter label string ────────────

function pr_buildFilterLabel(filters) {
  var p = [];
  if (filters.dateFrom || filters.dateTo)
    p.push((filters.dateFrom||'?') + ' → ' + (filters.dateTo||'?'));
  if (filters.brands       && filters.brands.length)        p.push('Brand: '   + filters.brands.join(', '));
  if (filters.methods      && filters.methods.length)        p.push('Method: ' + filters.methods.join(', '));
  if (filters.docTypes     && filters.docTypes.length)       p.push('DocType: '+ filters.docTypes.join(', '));
  if (filters.docStatuses  && filters.docStatuses.length)    p.push('Status: ' + filters.docStatuses.join(', '));
  if (filters.submittedBys && filters.submittedBys.length)   p.push('By: '     + filters.submittedBys.join(', '));
  return p.length ? p.join('  ·  ') : 'All payments';
}

// ── Menu wiring (add to onOpen inside 💎 Sales) ──────────────
// .addItem('📊 Payment Report', 'openPaymentReportDialog')

// ── Debug: kiểm tra lookup 1 APPT_ID cụ thể ────────────────
function pr_debugCustomerLookup() {
  var TEST_ID = 'AP-20251210-003'; // ← đổi thành ID cần test
  var map    = pr_buildCustomerMap_();
  var result = pr_lookupCustomer_(map, TEST_ID);
  SpreadsheetApp.getUi().alert(
    'TEST_ID : ' + TEST_ID + '\n' +
    'Result  : ' + (result || '(not found)') + '\n' +
    'Map size: ' + Object.keys(map).length + ' entries'
  );
}