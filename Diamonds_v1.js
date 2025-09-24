/** 
 * Diamonds_v1.gs — Phase 1 (Dialog A: Propose Diamonds)
 * Robust, header-by-name writes with aliases; atomic; batch efficient.
 * Requires manual Script Properties (see setup section):
 *   - CFG_DIAMONDS_TRACKING_FILE_ID: Spreadsheet ID of 200_ (Diamonds& Gems Tracking)
 *   - CFG_DIAMONDS_TRACKING_MASTER_TAB (optional): defaults to "0. MASTER LG SHEET"
 *
 * Requires the following headers to already exist:
 *   On 100_ (00_Master Appointments):
 *     - Center Stone Order Status
 *     - RootApptID
 *     - Customer Name
 *     - Visit Date
 *     - Visit Time
 *     - Brand
 *     - Assigned Rep
 *     - DV Stones (JSON Lines)
 *     - DV Stones Summary
 *   On 200_ (0. MASTER LG SHEET):
 *     - Stone Status
 *     - Order Status
 *     - Ordered By
 *     - Requested By
 *     - Request Date
 *     - Purchased / Ordered Date
 *     - Assigned Rep
 *     - Customer Name
 *     - Customer Appt Time & Date
 *     - Company
 *     - Vendor
 *     - Stone Type
 *     - Shape
 *     - Carat
 *     - Color
 *     - Clarity
 *     - Measurements
 *     - L/W Ratio
 *     - LAB
 *     - Certificate No
 *     - Stone Decision (PO, Return)
 *     - Cut
 *     - Pol.
 *     - Sym.
 *     - Fluor.Intesity   // supports alias "Fluor.Intensity" in code
 *     - Fluor.Color
 *     - RootApptID
 *
 * Menu: Add a menu entry to open dp_openProposeDiamonds().
 */

// ------------------------------ PUBLIC ENTRYPOINTS ------------------------------

/**
 * Menu entrypoint to open the dialog (hook this into your Sales menu).
 */

var DP200_INSERT_AT_ROW = 3;


function dp_openProposeDiamonds() {
  var ui = SpreadsheetApp.getUi();
  var bootstrap;
  try {
    // This validates the sheet/row and builds all dropdowns + context
    bootstrap = dp_bootstrapForProposeDialog_();
  } catch (e) {
    ui.alert('💎 Propose Diamonds', (e && e.message ? e.message : String(e)) +
      '\n\nOpen 100_ (00_Master Appointments) and select a customer data row, then try again.', ui.ButtonSet.OK);
    return;
  }

  // Inject the payload directly into the HTML template
  var t = HtmlService.createTemplateFromFile('dlg_propose_diamonds_v1');
  t.bootstrap = bootstrap;

  var html = t.evaluate()
    .setWidth(1100)
    .setHeight(680);
  ui.showModalDialog(html, '💎 Propose Diamonds');
}



/**
 * Bootstrap payload for the dialog: 100_ row context + dropdowns.
 */
function dp_bootstrapForProposeDialog_() {
  var ctx = dp_getActiveMasterRowContext_(); // throws helpful error if not on valid row
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

  return {
    context: {
      sheetName: ctx.sheet.getName(),
      rowIndex: ctx.rowIndex,
      rootApptId: ctx.rootApptId,
      customerName: ctx.customerName,
      visitDateValue: ctx.visitDateValue,
      visitTimeValue: ctx.visitTimeValue,
      visitDateStr: ctx.visitDateStr,
      visitTimeStr: ctx.visitTimeStr,
      companyBrand: ctx.companyBrand || null,
      assignedRep: ctx.assignedRep || null,
      timezone: tz
    },
    dropdowns: dp_dropdowns_()
  };
}

/**
 * Submit handler: validates, appends to 200_, updates 100_ CSOS and local JSON-lines + Summary.
 * Returns a result object with counts and warnings for UI to display.
 */
function dp_submitProposals(payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(28 * 1000);
  try {
    if (!payload || !Array.isArray(payload.stones) || payload.stones.length === 0) {
      throw new Error('No stones were provided.');
    }

    var ctx = dp_getActiveMasterRowContext_();

    // 🔹 OPEN 200_ TARGET + HEADER MAP
    var target = dp_get200Sheet_();
    var sh200 = target.sheet;

    // 🔹 Use the two-row header map for 200_
    var hm200 = dp_headerMapFor200_(sh200);

    // 🔎 DEBUG: log where critical columns were found
    dp_debug_log200Mapping_(sh200, hm200, [
      'Vendor','Stone Type','Shape','Carat','Color','Clarity','Cut',
      'Assigned Rep','RootApptID','Customer Name','Customer Appt Time & Date'
    ]);

    var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

    var lastCol = sh200.getLastColumn();   // 🔹 FULL SHEET WIDTH

    // Compose "Customer Appt Time & Date"
    var composedApptDateTime = dp_composeApptDateTimeString_(ctx.visitDateValue, ctx.visitTimeValue, tz);

    // Build rows
    var toAppend = [];
    var certNosNew = [];
    payload.stones.forEach(function (s, i) {
      var validated = dp_validateStoneInput_(s, i);
      certNosNew.push(validated.certificateNo.trim());

      // 🔹 PASS lastCol so we build rows at full width
      var rowArray = dp_build200RowArray_({
        stone: validated,
        ctx: ctx,
        hm200: hm200,
        composedApptDateTime: composedApptDateTime,
        lastCol: lastCol
      });
      toAppend.push(rowArray);
    });

    // Duplicates (warning only)
    var dupWarnings = dp_warnDuplicateCerts_(sh200, hm200, certNosNew);

    (function dp_debug_previewWrite(){
      if (!toAppend.length) return;
      function c(k){ return dp_findHeaderIndex_(hm200, dp_aliases200_[k] || [k], false); }
      var f = toAppend[0];
      var preview = {
        Vendor: {col: c('Vendor'), val: f[c('Vendor')-1]},
        StoneType: {col: c('Stone Type'), val: f[c('Stone Type')-1]},
        Shape: {col: c('Shape'), val: f[c('Shape')-1]},
        Carat: {col: c('Carat'), val: f[c('Carat')-1]},
        Color: {col: c('Color'), val: f[c('Color')-1]},
        Clarity: {col: c('Clarity'), val: f[c('Clarity')-1]},
        Cut: {col: c('Cut'), val: f[c('Cut')-1]},
        AssignedRep: {col: c('Assigned Rep'), val: f[c('Assigned Rep')-1]},
        RootApptID: {col: c('RootApptID'), val: f[c('RootApptID')-1]},
        CustomerName: {col: c('Customer Name'), val: f[c('Customer Name')-1]},
        ApptTimeDate: {col: c('Customer Appt Time & Date'), val: f[c('Customer Appt Time & Date')-1]},
        Measurements: {col: c('Measurements'), val: f[c('Measurements')-1]}
      };
      Logger.log('dp_submitProposals PRE-WRITE preview (first row): ' + JSON.stringify(preview, null, 2));
    })();


    // 🔹 PREPEND at row 3 (insert rows BEFORE row 3), then write across full width
    var insertAt = DP200_INSERT_AT_ROW; // 3
    if (toAppend.length > 0) {
      sh200.insertRowsBefore(insertAt, toAppend.length);

      // (A) Copy formats/validations FIRST from the (now-shifted) template row
      try {
        var templateRowIdx = insertAt + toAppend.length; // original row 3 moved down
        if (sh200.getLastRow() >= templateRowIdx) {
          var tpl = sh200.getRange(templateRowIdx, 1, 1, lastCol);
          var dest = sh200.getRange(insertAt, 1, toAppend.length, lastCol);
          tpl.copyTo(dest, {formatOnly: true});
        }
      } catch (e) {
        Logger.log('dp_submitProposals format copy error: ' + (e && e.message ? e.message : e));
      }

      // (B) Now write values across full width
      var writeRange = sh200.getRange(insertAt, 1, toAppend.length, lastCol);
      writeRange.setValues(toAppend);

      // (C) Read-back verify a few key columns for the first row we wrote
      (function dp_debug_readBack(){
        try {
          var rb = sh200.getRange(insertAt, 1, 1, lastCol).getDisplayValues()[0];
          function c(k){ return dp_findHeaderIndex_(hm200, dp_aliases200_[k] || [k], false); }
          var snapshot = {
            Vendor: rb[c('Vendor')-1],
            StoneType: rb[c('Stone Type')-1],
            Shape: rb[c('Shape')-1],
            Carat: rb[c('Carat')-1],
            Color: rb[c('Color')-1],
            Clarity: rb[c('Clarity')-1],
            Cut: rb[c('Cut')-1],
            AssignedRep: rb[c('Assigned Rep')-1],
            RootApptID: rb[c('RootApptID')-1],
            CustomerName: rb[c('Customer Name')-1],
            ApptTimeDate: rb[c('Customer Appt Time & Date')-1]
          };
          Logger.log('dp_submitProposals POST-WRITE readback (row '+insertAt+'): ' + JSON.stringify(snapshot, null, 2));
        } catch (e) {
          Logger.log('dp_submitProposals readback error: ' + (e && e.message ? e.message : e));
        }
      })();
    }



    // Refresh counts and update 100_
    var resultCounts = dp_computeCountsForAppointment_(sh200, hm200, ctx.rootApptId);
    dp_update100AfterPropose_(ctx, payload.stones, resultCounts);

    dp_onCsosChanged_(ctx.rootApptId, 'Diamond Memo – Proposed');

    // 🔹 RETURN precise target info for the UI
    return {
      ok: true,
      added: toAppend.length,
      duplicates: dupWarnings,
      counts: resultCounts,
      message: 'Added ' + toAppend.length + ' stone(s) to 200_. Center Stone Order Status set to “Diamond Memo – Proposed.”',
      targetSpreadsheetUrl: target.ss.getUrl(),
      targetSpreadsheetName: target.ss.getName(),
      targetTabName: target.tab,
      appendedFirstRow: insertAt,
      appendedLastRow: (insertAt + toAppend.length - 1)
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}



// ------------------------------ INTERNAL HELPERS ------------------------------

function dp_dropdowns_() {
  return {
    vendors: ['Loupe360', 'LG', 'From In Stock', 'Other'],
    stoneTypes: ['Lab Diamond', 'Natural Diamond', 'Sapphire', 'Emerald', 'Ruby'],
    shapes: ['Round', 'Oval', 'Pear', 'Cushion', 'Emerald', 'Radiant', 'Princess', 'Marquise', 'Asscher', 'Heart', 'Other'],
    clarities: ['FL', 'IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
    labs: ['GIA', 'IGI', 'Other'],
    // Cut/Pol/Sym
    grades: ['EX', 'VG', 'G', 'F', 'P'],
    // Fluorescence
    fluorIntensity: ['None', 'Faint', 'Medium', 'Strong', 'Very Strong'],
    fluorColor: ['None', 'Blue', 'Yellow', 'Other'],
    colorsMain: ['D', 'E', 'F', 'G', 'H', 'I', 'Fancy'],
    fancyIntensity: ['Faint', 'Light', 'Fancy', 'Intense', 'Vivid', 'Deep', 'Dark'], // used only when Color=Fancy
    fancyColors: ['Yellow', 'Brown', 'Blue', 'Pink', 'Green', 'Orange', 'Purple', 'Red', 'Gray', 'Black', 'Other']
  };
}

/**
 * Returns {sheet, rowIndex, headerMap, <fields>} for the active row on 100_.
 * Throws a helpful error if user isn't on a valid 100_ row.
 */
function dp_getActiveMasterRowContext_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getActiveSheet();
  var range = sh.getActiveRange();
  if (!range) throw new Error('Select a client row on 100_ (00_Master Appointments) and try again.');

  // Be tolerant: if we’re not on the 100_ sheet, try to locate it by name
  var hm = dp_headerMap_(sh);
  var missing = [];
  ['Center Stone Order Status','RootApptID','Customer Name','Visit Date','Visit Time'].forEach(function(h){
    if (dp_findHeaderIndex_(hm, dp_aliases100_[h], false) < 0) missing.push(h);
  });
  if (missing.length) {
    throw new Error('Required column(s) not found on 100_: ' + missing.join(', ') + '. Please add them, then try again.');
  }

  var need100 = ['Center Stone Order Status','RootApptID','Customer Name','Visit Date','Visit Time','Brand','Assigned Rep']; // removed the two DV columns from detection
  var is100 = need100.every(function(h){ return dp_findHeaderIndex_(hm, dp_aliases100_[h], false) > -1; });
  if (!is100) {
    // Try to locate a sheet with "00_Master Appointments" in name
    var fallback = ss.getSheets().find(function(s){ return /00[_\s-]*Master\s*Appointments/i.test(s.getName()); });
    if (fallback) {
      sh = fallback;
      hm = dp_headerMap_(sh);
      // Update range to active row index 2 by default if the original selection wasn't in this sheet
      range = sh.getRange(Math.max(2, sh.getActiveRange() ? sh.getActiveRange().getRow() : 2), 1);
    } else {
      throw new Error('Open the “100_ (00_Master Appointments)” sheet, select a customer row, and try again.');
    }
  }

  var row = range.getRow();
  if (row <= 1) throw new Error('Please select a data row (not the header).');

  var get = function(h, required) { return dp_getCellByHeader_(sh, hm, row, dp_aliases100_[h], required); };

  var centerStoneStatus = get('Center Stone Order Status', true);
  var rootApptId        = String(get('RootApptID', true) || '').trim();
  var customerName      = String(get('Customer Name', true) || '').trim();
  var visitDateValue    = get('Visit Date', true);
  var visitTimeValue    = get('Visit Time', true);
  var companyBrand      = String(get('Brand', false) || '').trim();
  var assignedRep       = String(get('Assigned Rep', false) || '').trim();

  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var visitDateStr = dp_formatDateOnly_(visitDateValue, tz);
  var visitTimeStr = dp_formatTimeOnly_(visitTimeValue, tz);

  return {
    sheet: sh,
    rowIndex: row,
    headerMap: hm,
    centerStoneStatus: centerStoneStatus,
    rootApptId: rootApptId,
    customerName: customerName,
    visitDateValue: visitDateValue,
    visitTimeValue: visitTimeValue,
    visitDateStr: visitDateStr,
    visitTimeStr: visitTimeStr,
    companyBrand: companyBrand,
    assignedRep: assignedRep
  };
}

function dp_get200Sheet_() {
  var props = PropertiesService.getScriptProperties();
  var fileId = props.getProperty('CFG_DIAMONDS_TRACKING_FILE_ID');
  if (!fileId) {
    throw new Error('Missing Script Property: CFG_DIAMONDS_TRACKING_FILE_ID.\n' +
      'Set this to the Spreadsheet ID of 200_ (Diamonds & Gems Tracking).');
  }
  var tab = props.getProperty('CFG_DIAMONDS_TRACKING_MASTER_TAB') || '0. MASTER LG SHEET';
  var ss = SpreadsheetApp.openById(fileId);
  var sh = ss.getSheetByName(tab);
  if (!sh) {
    throw new Error('Tab "' + tab + '" not found in 200_. Check CFG_DIAMONDS_TRACKING_MASTER_TAB or the sheet name.');
  }
  return { ss: ss, sheet: sh, tab: tab };
}

// --- Header mapping / aliases ---

function dp_headerMap_(sheet) {
  var row = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  var byExact = {};
  var byNorm = {};
  for (var i=0;i<row.length;i++) {
    var h = String(row[i] || '').trim();
    if (!h) continue;
    byExact[h] = i+1; // 1-based
    var n = dp_norm_(h);
    if (!(n in byNorm)) byNorm[n] = i+1;
  }
  return { byExact: byExact, byNorm: byNorm };
}

// 200_ uses two header rows; combine row 1 and row 2 text per column.
var DP200_HEADER_ROWS = 2;

/**
 * Builds a header map for 200_ from the first two rows.
 * If both rows have text, we combine: "Row1 Row2".
 * If row1 is blank and row2 has text, we use row2.
 * Returns the same shape as dp_headerMap_: { byExact, byNorm }
 */
function dp_headerMapFor200_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headerRows = Math.min(DP200_HEADER_ROWS, Math.max(1, sheet.getMaxRows() >= 2 ? 2 : 1));
  var rows = sheet.getRange(1, 1, headerRows, lastCol).getDisplayValues();

  var byExact = {};
  var byNorm = {};

  for (var c = 0; c < lastCol; c++) {
    var h1 = (rows[0] && rows[0][c]) ? String(rows[0][c]).trim() : '';
    var h2 = (rows[1] && rows[1][c]) ? String(rows[1][c]).trim() : '';
    var h  = h1 && h2 ? (h1 + ' ' + h2) : (h1 || h2); // combine if both exist
    if (!h) continue;
    byExact[h] = c + 1; // 1-based col index
    var n = dp_norm_(h);
    if (!(n in byNorm)) byNorm[n] = c + 1;
  }
  return { byExact: byExact, byNorm: byNorm };
}


function dp_findHeaderIndex_(hm, aliases, required) {
  for (var i=0;i<aliases.length;i++) {
    var a = aliases[i];
    if (hm.byExact[a]) return hm.byExact[a];
    var n = dp_norm_(a);
    if (hm.byNorm[n]) return hm.byNorm[n];
  }
  if (required) {
    throw new Error('Required header not found: ' + aliases[0] + ' (with aliases: ' + aliases.join(', ') + ')');
  }
  return -1;
}

function dp_getCellByHeader_(sheet, hm, row, aliases, required) {
  var col = dp_findHeaderIndex_(hm, aliases, required);
  if (col < 0) return '';
  return sheet.getRange(row, col).getValue();
}

function dp_setCellByHeader_(sheet, hm, row, aliases, value) {
  var col = dp_findHeaderIndex_(hm, aliases, true);
  sheet.getRange(row, col).setValue(value);
}

function dp_norm_(s) {
  return String(s || '').toLowerCase().replace(/\s+/g,'').replace(/[^a-z0-9]/g,'');
}

var dp_aliases100_ = {
  'Center Stone Order Status': ['Center Stone Order Status','CSOS','CenterStoneOrderStatus'],
  'RootApptID': ['RootApptID','APPT_ID','RootApptId','RootApptID '],
  'Customer Name': ['Customer Name','Client Name','Customer'],
  'Visit Date': ['Visit Date','VisitDate','Visit_Date'],
  'Visit Time': ['Visit Time','VisitTime','Visit_Time'],
  'Brand': ['Brand','Company','Company Brand'],
  'Assigned Rep': ['Assigned Rep','Sales Rep','Assigned Sales Rep'],
  'DV Stones (JSON Lines)': ['DV Stones (JSON Lines)','DV Stones JSON Lines','DV Stones-JSON Lines'],
  'DV Stones Summary': ['DV Stones Summary','DV Stones- Summary','DV Stones- Summary ']
};

var dp_aliases200_ = {
  'Stone Status': ['Stone Status','StoneStatus'],
  'Order Status': ['Order Status','OrderStatus'],
  'Ordered By': ['Ordered By','OrderedBy'],
  'Requested By': ['Requested By','RequestedBy'],
  'Request Date': ['Request Date','RequestDate'],
  'Purchased / Ordered Date': ['Purchased / Ordered Date','Purchased/Ordered Date','PurchasedOrderedDate'],
  'Assigned Rep': ['Assigned Rep','Sales Rep','Assigned Sales Rep','AssignedRep'],
  'Customer Name': ['Customer Name','Client Name','Customer'],
  'Customer Appt Time & Date': ['Customer Appt Time & Date','Customer Appt Time&Date','CustomerApptTimeDate'],
  'Company': ['Company','Brand'],
  'Vendor': ['Vendor'],
  'Stone Type': ['Stone Type','StoneType'],
  'Shape': ['Shape'],
  'Carat': ['Carat'],
  'Color': ['Color'],
  'Clarity': ['Clarity'],
  'Measurements': ['Measurements','Measurement','Meas.','Meas'],
  'L/W Ratio': ['L/W Ratio','L-W Ratio','LW Ratio','LWRatio'],
  'LAB': ['LAB','Lab','Grading Lab'],
  'Certificate No': ['Certificate No','Cert #','Cert No','Certificate #','Certificate Number'],
  'Stone Decision (PO, Return)': ['Stone Decision (PO, Return)','Stone Decision\n(PO, Return)', 'Stone Decision','StoneDecision'],
  'Cut': ['Cut'],
  'Pol.': ['Pol.','Pol','Polish'],
  'Sym.': ['Sym.','Sym','Symmetry'],
  'Fluor.Intesity': ['Fluor.Intesity','Fluor.Intensity','Fluor Intensity'],
  'Fluor.Color': ['Fluor.Color','Fluor Color','Fluorescence Color'],
  'RootApptID': ['RootApptID','APPT_ID','RootApptId','RootApptID ']
};

// --- Validation & formatting ---

function dp_validateStoneInput_(s, idx) {
  var prefix = 'Row #' + (idx+1) + ': ';
  var must = function(cond, msg){ if(!cond) throw new Error(prefix + msg); };

  var stoneType = String(s.stoneType || '').trim();
  var shape     = String(s.shape || '').trim();
  var vendor    = String(s.vendor || '').trim();
  var color     = String(s.color || '').trim();
  var clarity   = String(s.clarity || '').trim();
  var lab       = String(s.lab || '').trim();
  var certNo    = String(s.certNo || '').trim();
  must(stoneType, 'Stone Type is required.');
  must(shape,     'Shape is required.');
  must(vendor,    'Vendor is required.');
  must(color,     'Color is required.');
  must(clarity,   'Clarity is required.');
  must(lab,       'LAB is required.');
  must(certNo,    'Certificate No is required.');

  // carat & lw ratio numeric
  var carat = Number(s.carat);
  must(!isNaN(carat) && carat > 0, 'Carat must be a positive number.');
  carat = Math.round(carat * 100) / 100; // 2 decimals

  var lwRatio = (s.lwRatio === '' || s.lwRatio === null || typeof s.lwRatio === 'undefined') ? '' : Number(s.lwRatio);
  if (lwRatio !== '') {
    must(!isNaN(lwRatio) && lwRatio > 0, 'L/W Ratio must be a positive number.');
    lwRatio = Math.round(lwRatio * 100) / 100;
  }

  // Fancy color extras (if Color=Fancy)
  var colorOut = color;
  if (color === 'Fancy') {
    var fColor = String(s.fancyColor || '').trim();
    var fInt   = String(s.fancyIntensity || '').trim();
    must(fColor, 'Fancy Color is required when Color = Fancy.');
    must(fInt,   'Fancy Intensity is required when Color = Fancy.');
    // Since 200_ does not have separate columns, encode both into Color field
    colorOut = 'Fancy (' + fColor + '; ' + fInt + ')';
  }

  // Optional advanced
  var measurements = (s.measurements || '').trim(); // free text
  var cut  = s.cut ? String(s.cut).trim() : '';
  var pol  = s.pol ? String(s.pol).trim() : '';
  var sym  = s.sym ? String(s.sym).trim() : '';
  var flI  = s.fluorIntensity ? String(s.fluorIntensity).trim() : '';
  var flC  = s.fluorColor ? String(s.fluorColor).trim() : '';

  return {
    stoneType: stoneType,
    shape: shape,
    vendor: vendor,
    color: colorOut,
    clarity: clarity,
    lab: lab,
    certificateNo: certNo,
    carat: carat,
    lwRatio: lwRatio === '' ? '' : lwRatio,
    measurements: measurements,
    cut: cut,
    pol: pol,
    sym: sym,
    fluorIntensity: flI,
    fluorColor: flC
  };
}

function dp_formatDateOnly_(d, tz) {
  if (Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d)) {
    return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  }
  // Try parsing if string
  var dd = new Date(d);
  if (!isNaN(dd)) return Utilities.formatDate(dd, tz, 'yyyy-MM-dd');
  return '';
}

function dp_formatTimeOnly_(t, tz) {
  if (Object.prototype.toString.call(t) === '[object Date]' && !isNaN(t)) {
    return Utilities.formatDate(t, tz, 'HH:mm');
  }
  var tt = new Date(t);
  if (!isNaN(tt)) return Utilities.formatDate(tt, tz, 'HH:mm');
  // Try simple "h:mm a"
  var m = String(t || '').trim().match(/^(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);
  if (m) {
    var h = Number(m[1]); var min = Number(m[2]); var ap = m[3];
    if (ap) {
      if (/PM/i.test(ap) && h < 12) h += 12;
      if (/AM/i.test(ap) && h === 12) h = 0;
    }
    var base = new Date();
    base.setHours(h, min, 0, 0);
    return Utilities.formatDate(base, tz, 'HH:mm');
  }
  return '';
}

/** Returns "yyyy-MM-dd HH:mm" or "" if unable to compose. */
function dp_composeApptDateTimeString_(dateVal, timeVal, tz) {
  var ds = dp_formatDateOnly_(dateVal, tz);
  var ts = dp_formatTimeOnly_(timeVal, tz);
  if (!ds || !ts) return '';
  return ds + ' ' + ts;
}

/** Construct a 200_ row (array) respecting live header length and aliases. */
function dp_build200RowArray_(opts) {
  var s = opts.stone, ctx = opts.ctx, hm200 = opts.hm200;
  var lastCol = opts.lastCol;                     // 🔹 full width
  var row = new Array(lastCol).fill('');          // 🔹 prefill blanks

  var put = function(headerKey, value, required){
    var col = dp_findHeaderIndex_(hm200, dp_aliases200_[headerKey], !!required);
    if (col > -1) row[col-1] = value;            // absolute column index
  };

  // Required workflow fields
  put('Stone Status', 'Diamond Viewing', true);
  put('Order Status', 'Proposing', true);
  put('Ordered By', ''); // blank at propose
  put('Requested By', dp_getCurrentUserEmail_());
  put('Request Date', new Date());
  put('Purchased / Ordered Date', '');

  // Appointment context
  put('Assigned Rep', ctx.assignedRep || '', true);
  put('Customer Name', ctx.customerName || '', true);
  put('Customer Appt Time & Date', opts.composedApptDateTime || '', true);
  put('RootApptID', ctx.rootApptId || '', true);
  put('Company', ctx.companyBrand ? ctx.companyBrand : 'N/A');

  // Stone fields
  put('Vendor', s.vendor || '', true);
  put('Stone Type', s.stoneType || '', true);
  put('Shape', s.shape || '', true);
  put('Carat', s.carat, true);
  put('Color', s.color || '', true);
  put('Clarity', s.clarity || '', true);
  put('LAB', s.lab || '', true);
  put('Certificate No', s.certificateNo || '', true);
  put('Measurements', s.measurements || '');
  put('L/W Ratio', s.lwRatio === '' ? '' : s.lwRatio);

  // Advanced (optional)
  put('Cut', s.cut || '');
  put('Pol.', s.pol || '');
  put('Sym.', s.sym || '');
  put('Fluor.Intesity', s.fluorIntensity || '');
  put('Fluor.Color', s.fluorColor || '');

  // Decision blank at propose
  put('Stone Decision (PO, Return)', '');

  // Fill any remaining undefined with blank strings to avoid "Cannot convert" errors
  for (var i=0; i<row.length; i++) if (typeof row[i] === 'undefined') row[i] = '';

  return row;
}

function dp_getCurrentUserEmail_() {
  try {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
  } catch (e) {
    return '';
  }
}

/** Return list of duplicate certNos that already exist in 200_. */
function dp_warnDuplicateCerts_(sh200, hm200, certNosNew) {
  var col = dp_findHeaderIndex_(hm200, dp_aliases200_['Certificate No'], true);
  var vals = sh200.getRange(2, col, Math.max(0, sh200.getLastRow()-1), 1).getDisplayValues().map(function(r){ return String(r[0] || '').trim().toLowerCase(); });
  var set = new Set(vals);
  var dups = [];
  certNosNew.forEach(function(c){
    var key = String(c || '').trim().toLowerCase();
    if (key && set.has(key)) dups.push(c);
  });
  return dups;
}

/** Compute summary counts for this RootApptID by reading 200_. */
// Counts stones for a RootApptID from 200_.
// Robust to multi-select chips ("Diamond Viewing, In Stock") and 2-row headers.
// Returns { proposing, onTheWay, notApproved, delivered, inStock, total }.
function dp_computeCountsForAppointment_(sh200, hm200, rootApptId) {
  var colRoot       = dp_findHeaderIndex_(hm200, dp_aliases200_['RootApptID'], true);
  var colOrder      = dp_findHeaderIndex_(hm200, dp_aliases200_['Order Status'], true);
  var colStoneStat  = dp_findHeaderIndex_(hm200, dp_aliases200_['Stone Status'], true);

  var last = sh200.getLastRow();
  if (last < 3) {
    return { proposing: 0, onTheWay: 0, notApproved: 0, delivered: 0, inStock: 0, total: 0 };
  }

  // 200_ has two header rows → start at row 3
  var rng = sh200.getRange(3, 1, last - 2, sh200.getLastColumn()).getDisplayValues();
  var idxRoot = colRoot - 1, idxOrder = colOrder - 1, idxStone = colStoneStat - 1;

  var want = String(rootApptId || '').trim();
  var counts = { proposing: 0, onTheWay: 0, notApproved: 0, delivered: 0, inStock: 0, total: 0 };

  rng.forEach(function (row) {
    if (String(row[idxRoot] || '').trim() !== want) return;

    var order   = String(row[idxOrder] || '').trim();
    var stoneSt = String(row[idxStone] || '').trim();

    counts.total++;

    if (/^Proposing$/i.test(order))        counts.proposing++;
    else if (/^On the Way$/i.test(order))  counts.onTheWay++;
    else if (/^Not Approved$/i.test(order))counts.notApproved++;
    else if (/^Delivered$/i.test(order))   counts.delivered++;

    if (stoneSt) {
      // Split on common delimiters (; , | •) and trim
      var tokens = stoneSt.split(/\s*[;,|]\s*|\s*•\s*/g)
                          .map(function(s){ return s.trim(); })
                          .filter(Boolean);
      if (tokens.some(function(t){ return /^in stock$/i.test(t); })) {
        counts.inStock++;
      }
    }
  });

  return counts;
}


/** Apply 100_ updates: CSOS, JSON-lines append, Summary. */
function dp_update100AfterPropose_(ctx, stones, countsAfter) {
  var sh = ctx.sheet;
  var hm = ctx.headerMap;

  // 1) Set CSOS = "Diamond Memo – Proposed"
  dp_setCellByHeader_(sh, hm, ctx.rowIndex, dp_aliases100_['Center Stone Order Status'], 'Diamond Memo – Proposed');

  // 2) Append JSON lines
  var colJson = dp_findHeaderIndex_(hm, dp_aliases100_['DV Stones (JSON Lines)'], true);
  var curJson = String(sh.getRange(ctx.rowIndex, colJson).getValue() || '').trim();
  var linesToAdd = stones.map(function(s){
    // Store a compact JSON per stone (lossless for later)
    var item = {
      ts: new Date().toISOString(),
      vendor: s.vendor,
      stoneType: s.stoneType,
      shape: s.shape,
      carat: Number(s.carat),
      color: s.color === 'Fancy' ? { main: 'Fancy', fancyColor: s.fancyColor, fancyIntensity: s.fancyIntensity } : s.color,
      clarity: s.clarity,
      lab: s.lab,
      certNo: s.certNo,
      measurements: s.measurements || '',
      lwRatio: (s.lwRatio === '' ? '' : Number(s.lwRatio))
    };
    return JSON.stringify(item);
  });
  var newJsonValue = curJson ? (curJson + '\n' + linesToAdd.join('\n')) : linesToAdd.join('\n');
  sh.getRange(ctx.rowIndex, colJson).setValue(newJsonValue);

  // 3) Update Summary badge (count snapshot)
  var colSum = dp_findHeaderIndex_(hm, dp_aliases100_['DV Stones Summary'], true);
  var summary = 'Proposed: ' + countsAfter.proposing + ' • On the Way: ' + countsAfter.onTheWay +
                ' • Not Approved: ' + countsAfter.notApproved + ' • In Stock: ' + countsAfter.inStock +
                ' • Total: ' + countsAfter.total;
  sh.getRange(ctx.rowIndex, colSum).setValue(summary);
}

/** Integration hook for reminders — safe no-op by default. */
function dp_onCsosChanged_(rootApptId, newStatus) {
  // If your DV/Reminders module exposes a function to confirm PROPOSE_NUDGE,
  // you can wire it here, e.g.:
  // try { DV_confirmProposeNudgeByRootApptId_(rootApptId); } catch(e) {}
  // This hook purposely swallows errors to avoid breaking dialog UX.
}

/**
 * TEMP: Show the 200_ Script Properties so you can confirm you're looking at the same file.
 * Add a temporary menu item to call this: .addItem('Diamonds — Debug target', 'dp_debug_show200Props')
 */
function dp_debug_show200Props() {
  var p = PropertiesService.getScriptProperties();
  var id = p.getProperty('CFG_DIAMONDS_TRACKING_FILE_ID') || '(not set)';
  var tab = p.getProperty('CFG_DIAMONDS_TRACKING_MASTER_TAB') || '0. MASTER LG SHEET';
  var msg = 'CFG_DIAMONDS_TRACKING_FILE_ID:\n' + id + '\n\n' +
            'CFG_DIAMONDS_TRACKING_MASTER_TAB:\n' + tab + '\n\n' +
            'If the success panel opens a different URL than you expect, update this property.';
  SpreadsheetApp.getUi().alert('Diamonds — 200_ Target', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Logs the actual 200_ header strings found for the given logical keys.
 * Helps confirm the two-row combiner is locking onto the expected columns.
 */
function dp_debug_log200Mapping_(sh200, hm200, keys) {
  try {
    var lastCol = sh200.getLastColumn();
    var row1 = sh200.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    var row2 = sh200.getRange(2, 1, 1, lastCol).getDisplayValues()[0];

    Logger.log('--- 200_ header row 1 (preview) ---');
    Logger.log(JSON.stringify(row1));
    Logger.log('--- 200_ header row 2 (preview) ---');
    Logger.log(JSON.stringify(row2));

    var report = {};
    keys.forEach(function(k){
      var idx = dp_findHeaderIndex_(hm200, dp_aliases200_[k] || [k], false);
      if (idx > -1) {
        report[k] = { col: idx, row1: row1[idx-1] || '', row2: row2[idx-1] || '' };
      } else {
        report[k] = { col: null, note: 'NOT FOUND' };
      }
    });
    Logger.log('--- 200_ header resolution ---');
    Logger.log(JSON.stringify(report, null, 2));
  } catch (e) {
    Logger.log('dp_debug_log200Mapping_ error: ' + (e && e.message ? e.message : e));
  }
}
function dp_debug_dump200CombinedHeadersOnce() {
  var target = dp_get200Sheet_();
  var sh200 = target.sheet;
  var lastCol = sh200.getLastColumn();
  var rows = sh200.getRange(1, 1, Math.min(2, sh200.getMaxRows() >= 2 ? 2 : 1), lastCol).getDisplayValues();
  var h1 = rows[0] || [], h2 = rows[1] || [];
  var combined = [];
  for (var c = 0; c < lastCol; c++) {
    var a = String(h1[c] || '').trim(), b = String(h2[c] || '').trim();
    combined.push(a && b ? (a + ' ' + b) : (a || b));
  }
  Logger.log('--- 200_ combined header (2 rows) ---');
  Logger.log(JSON.stringify(combined));
}

/** Locate ALL rows in 100_ with the given RootApptID. */
function dp_find100RowsByRootApptId_(rootApptId) {
  var ss = SpreadsheetApp.getActive();
  // Find the "100_ (00_Master Appointments)" sheet by name pattern
  var sh = ss.getSheets().find(function(s){
    return /00[_\s-]*Master\s*Appointments/i.test(s.getName());
  }) || ss.getActiveSheet(); // fallback

  var hm = dp_headerMap_(sh);
  // Ensure required columns exist
  var needed = ['RootApptID','Center Stone Order Status','DV Stones (JSON Lines)','DV Stones Summary'];
  var missing = needed.filter(function(h){ return dp_findHeaderIndex_(hm, dp_aliases100_[h], false) < 0; });
  if (missing.length) {
    throw new Error('100_ sheet is missing required columns: ' + missing.join(', '));
  }

  var colRoot = dp_findHeaderIndex_(hm, dp_aliases100_['RootApptID'], true);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getDisplayValues();

  var want = String(rootApptId || '').trim();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][colRoot - 1] || '').trim() === want) {
      out.push({ sheet: sh, rowIndex: i + 2, headerMap: hm });
    }
  }
  return out;
}


// ------------------------------ END Diamonds_v1.gs ------------------------------


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



