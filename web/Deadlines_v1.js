/**
 * Deadlines_v1.gs — v1.2 (2025-09-03)
 * ----------------------------------------------------
 * Dialog + server logic to record "3D Deadline" or "Production Deadline"
 * for the active row, increment "# of Times ... Moved" counters, and append
 * a log row into the client's "Client Status" sheet (tab "Client Status",
 * headers on row 14) using the URL in column "Client Status Report URL".
 *
 * MAIN sheet required headers (row 1; order agnostic):
 *   - 3D Deadline
 *   - # of Times 3D Deadline Moved
 *   - Production Deadline
 *   - # of Times Prod. Deadline Moved
 *   - Client Status Report URL
 * Optional (if present in main sheet, gets copied to log):
 *   - Assisted Rep
 *
 * CLIENT STATUS report requirements (target Sheet):
 *   - Tab: "Client Status"
 *   - Header row: 14
 *   - Columns (exact names, used as-is):
 *     Log Date | Sales Stage | Conversion Status | Custom Order Status |
 *     Center Stone Order Status | Next Steps | Deadline Type | Deadline Date |
 *     Move Count | Assisted Rep | Updated By | Updated At
 */

var Deadlines = (function () {
  'use strict';

  // ---- Canonical header names in MAIN sheet ----
  var HDR_3D_DEADLINE = '3D Deadline';
  var HDR_3D_MOVES = '# of Times 3D Deadline Moved';
  var HDR_PROD_DEADLINE = 'Production Deadline';
  var HDR_PROD_MOVES = '# of Times Prod. Deadline Moved';
  var HDR_CLIENT_STATUS_URL = 'Client Status Report URL';
  // Optional column; if present we copy it into the log
  var HDR_ASSISTED_REP_OPTIONAL = 'Assisted Rep';

  // ---- Optional summary context (column candidates in MAIN sheet) ----
  var HDR_CANDIDATES = {
    CUSTOMER_NAME: ['Customer Name','Customer','Client Name'],
    SO_NUMBER:     ['SO#','SO #','SO Number','SO No','Sales Order #'],
    ROOT_APPT_ID:  ['RootApptID','Root Appt ID','Root Appointment ID','RootAppt Id','rootApptID']
  };

  // ---- Client Status log spec (TARGET REPORT SHEET) ----
  var CLIENT_STATUS_TAB = 'Client Status';
  var CLIENT_STATUS_HEADER_ROW = 14;
  var CLIENT_STATUS_COLUMNS = [
    'Log Date',
    'Sales Stage',
    'Conversion Status',
    'Custom Order Status',
    'Center Stone Order Status',
    'Next Steps',
    'Deadline Type',
    'Deadline Date',
    'Move Count',
    'Assisted Rep',
    'Updated By',
    'Updated At'
  ];

  /**
   * Opens the Record Deadline dialog.
   */
  function showRecordDeadlineDialog() {
    Logger.log('[showRecordDeadlineDialog] Opening dialog.');
    var html = HtmlService.createTemplateFromFile('dlg_record_deadline_v1')
      .evaluate()
      .setWidth(460)
      .setHeight(320);
    SpreadsheetApp.getUi().showModalDialog(html, 'Record Deadline');
  }

  /**
   * Returns initial data for the dialog.
   */
  function getRecordDeadlineInit() {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getActiveSheet();
    var tz = ss.getSpreadsheetTimeZone();
    var rowIndex = sheet.getActiveCell() ? sheet.getActiveCell().getRow() : null;

    Logger.log('[getRecordDeadlineInit] SS: %s | Sheet: %s | Row: %s | TZ: %s',
               ss.getName(), sheet.getName(), rowIndex, tz);

    if (!rowIndex || rowIndex < 1) {
      throw new Error('Please select a row first.');
    }
    if (rowIndex === 1) {
      throw new Error('You selected the header row. Please select a data row.');
    }

    var headers = getHeaderMap_(sheet);
    assertHeadersPresent_(headers);

    var rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    var existing3d = getCellValueByHeader_(rowValues, headers, HDR_3D_DEADLINE);
    var existingProd = getCellValueByHeader_(rowValues, headers, HDR_PROD_DEADLINE);
    var clientStatusUrl = getCellValueByHeader_(rowValues, headers, HDR_CLIENT_STATUS_URL);

    Logger.log('[getRecordDeadlineInit] existing3D: %s | existingProd: %s | hasURL: %s',
               existing3d, existingProd, !!clientStatusUrl);

    return {
      timezone: tz,
      rowIndex: rowIndex,
      existing: {
        threeD: serializeDateForInput_(existing3d, tz),
        production: serializeDateForInput_(existingProd, tz)
      },
      hasClientStatusUrl: isNonEmpty_(clientStatusUrl),
      clientStatusUrl: clientStatusUrl || ''
    };
  }

  /**
   * Saves deadline for selected row and appends to Client Status log.
   * @param {{kind:('3D'|'PROD'), dateIso:string}} payload
   * @return {{ok:boolean, message:string, moveCount:number}}
   */
  function saveRecordDeadline(payload) {
    Logger.log('[saveRecordDeadline] Payload: %s', JSON.stringify(payload));
    var kind = (payload && payload.kind || '').toUpperCase();
    if (kind !== '3D' && kind !== 'PROD') {
      throw new Error('Invalid deadline type.');
    }
    var dateIso = (payload && payload.dateIso || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateIso)) {
      throw new Error('Please select a valid date.');
    }

    var ss = SpreadsheetApp.getActive();
    var tz = ss.getSpreadsheetTimeZone();
    var sheet = ss.getActiveSheet();
    var rowIndex = sheet.getActiveCell() ? sheet.getActiveCell().getRow() : null;

    Logger.log('[saveRecordDeadline] SS: %s | Sheet: %s | Row: %s | TZ: %s',
               ss.getName(), sheet.getName(), rowIndex, tz);

    if (!rowIndex || rowIndex < 2) {
      throw new Error('Please select a data row first.');
    }

    var headers = getHeaderMap_(sheet);
    assertHeadersPresent_(headers);

    var rowRange = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
    var rowValues = rowRange.getValues()[0];

    var clientStatusUrl = getCellValueByHeader_(rowValues, headers, HDR_CLIENT_STATUS_URL);
    if (!isNonEmpty_(clientStatusUrl)) {
      Logger.log('[saveRecordDeadline] Missing Client Status Report URL at row %s', rowIndex);
      throw new Error('Missing "Client Status Report URL" for this row. Please submit the Client Status report first, then record a deadline.');
    }

    var newDate = parseIsoDateAtMidnight_(dateIso, tz);
    var userEmail = getActiveUserEmail_() || Session.getEffectiveUser().getEmail() || 'unknown@user';

    var hdrDeadline = (kind === '3D') ? HDR_3D_DEADLINE : HDR_PROD_DEADLINE;
    var hdrMoves = (kind === '3D') ? HDR_3D_MOVES : HDR_PROD_MOVES;

    var oldDateVal = getCellValueByHeader_(rowValues, headers, hdrDeadline);
    var oldMovesVal = getCellValueByHeader_(rowValues, headers, hdrMoves);

    var oldDate = coerceToDate_(oldDateVal);
    var oldMoves = coerceToNumber_(oldMovesVal);
    if (isNaN(oldMoves) || oldMoves < 0) oldMoves = 0;

    var isChange = !datesEqual_(oldDate, newDate, tz);
    var newMoves = (oldDate ? (isChange ? (oldMoves + 1) : oldMoves) : 0);

    Logger.log('[saveRecordDeadline] Kind: %s | OldDate: %s | NewDate: %s | OldMoves: %s | NewMoves: %s | IsChange: %s',
               kind, oldDate, newDate, oldMoves, newMoves, isChange);

    // Write back to main sheet
    var deadlineCol = headers.find(hdrDeadline);
    var movesCol = headers.find(hdrMoves);

    Logger.log('[saveRecordDeadline] Writing main sheet: deadlineCol=%s, movesCol=%s', deadlineCol, movesCol);
    sheet.getRange(rowIndex, deadlineCol).setValue(newDate);
    sheet.getRange(rowIndex, movesCol).setValue(newMoves);
    SpreadsheetApp.flush();

    // Optional Assisted Rep from main sheet (defensive against missing constant)
    var assistedRep = '';
    var assistedCol = (typeof HDR_ASSISTED_REP_OPTIONAL !== 'undefined')
      ? headers.find(HDR_ASSISTED_REP_OPTIONAL)
      : headers.find('Assisted Rep');

    if (assistedCol) {
      assistedRep = rowValues[assistedCol - 1] || '';
      Logger.log('[saveRecordDeadline] Assisted Rep (optional): %s', assistedRep);
    } else {
      Logger.log('[saveRecordDeadline] Assisted Rep column not present in main sheet.');
    }

    // Build log info
    var logInfo = {
      kind: (kind === '3D' ? '3D Deadline' : 'Production Deadline'),
      newDate: newDate,
      moveCount: newMoves,
      assistedRep: assistedRep,
      updatedBy: userEmail,
      updatedAt: new Date(),
      spreadsheetName: ss.getName(),
      sheetName: sheet.getName(),
      rowIndex: rowIndex
    };

    Logger.log('[saveRecordDeadline] Logging to Client Status. URL: %s | Info: %s',
               clientStatusUrl, JSON.stringify({
                 kind: logInfo.kind,
                 newDate: logInfo.newDate,
                 moveCount: logInfo.moveCount,
                 assistedRep: logInfo.assistedRep,
                 updatedBy: logInfo.updatedBy
               }));

    appendToClientStatusReportStrict_(clientStatusUrl, logInfo);

    // Build human-readable summary
    var fmt = function(d){ return d ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : ''; };
    var summary = {
      deadlineType: (kind === '3D' ? '3D Deadline' : 'Production Deadline'),
      customerName: getMaybeRowValueByCandidates_(rowValues, headers, HDR_CANDIDATES.CUSTOMER_NAME),
      soNumber:     getMaybeRowValueByCandidates_(rowValues, headers, HDR_CANDIDATES.SO_NUMBER),
      rootApptId:   getMaybeRowValueByCandidates_(rowValues, headers, HDR_CANDIDATES.ROOT_APPT_ID),
      previousDeadline: fmt(oldDate),
      newDeadline:  fmt(newDate),
      moveCount:    newMoves
    };

    Logger.log('[saveRecordDeadline] Success summary -> Customer: %s | SO#: %s | RootApptID: %s | Prev: %s | New: %s | Moves: %s',
      summary.customerName, summary.soNumber, summary.rootApptId, summary.previousDeadline, summary.newDeadline, summary.moveCount);

    var msg = (isChange || !oldDate) ? (kind + ' saved.') : (kind + ' unchanged; no move counted.');
    return { ok: true, message: msg, moveCount: newMoves, summary: summary };
  }

  // -------------------- Client Status logging (STRICT) --------------------

  function appendToClientStatusReportStrict_(url, logInfo) {
    var fileId = extractFileIdFromUrl_(url);
    if (!fileId) {
      Logger.log('[appendToClientStatusReportStrict_] ERROR: cannot parse fileId from URL: %s', url);
      throw new Error('Could not parse file ID from Client Status Report URL.');
    }

    var file = DriveApp.getFileById(fileId);
    var mime = file.getMimeType();
    Logger.log('[appendToClientStatusReportStrict_] Target file: "%s" | Mime: %s', file.getName(), mime);

    if (mime !== MimeType.GOOGLE_SHEETS) {
      throw new Error('Client Status Report URL must point to a Google Sheet with tab "Client Status" and headers on row 14.');
    }

    var targetSS = SpreadsheetApp.openById(fileId);
    var sh = targetSS.getSheetByName(CLIENT_STATUS_TAB);
    if (!sh) {
      Logger.log('[appendToClientStatusReportStrict_] ERROR: Missing tab "%s"', CLIENT_STATUS_TAB);
      throw new Error('Client Status report is missing required tab "' + CLIENT_STATUS_TAB + '".');
    }

    // Read header row 14
    var lastCol = Math.max(sh.getLastColumn(), CLIENT_STATUS_COLUMNS.length);
    var headerVals = sh.getRange(CLIENT_STATUS_HEADER_ROW, 1, 1, lastCol).getValues()[0];
    var headerMap = createHeaderMapFromRow_(headerVals);

    // Validate required columns
    var missing = CLIENT_STATUS_COLUMNS.filter(function (name) { return !headerMap.find(name); });
    if (missing.length) {
      Logger.log('[appendToClientStatusReportStrict_] ERROR: Missing required columns in row 14: %s', missing.join(', '));
      throw new Error('Client Status tab "' + CLIENT_STATUS_TAB + '" is missing required column(s) in header row ' +
        CLIENT_STATUS_HEADER_ROW + ': ' + missing.join(', '));
    }

    // Next row to write
    var nextRow = Math.max(sh.getLastRow() + 1, CLIENT_STATUS_HEADER_ROW + 1);
    Logger.log('[appendToClientStatusReportStrict_] NextRow: %s (lastRow: %s)', nextRow, sh.getLastRow());

    // Build write set — exact columns specified
    var now = new Date();
    var logDate = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0); // date-only
    var writePairs = [
      ['Log Date', logDate],
      ['Deadline Type', logInfo.kind],
      ['Deadline Date', logInfo.newDate],
      ['Move Count', logInfo.moveCount],
      ['Assisted Rep', logInfo.assistedRep || ''],
      ['Updated By', logInfo.updatedBy],
      ['Updated At', logInfo.updatedAt]
    ];

    writePairs.forEach(function (p) {
      Logger.log('[appendToClientStatusReportStrict_] Writing "%s" => %s', p[0], p[1]);
      var colIndex = headerMap.find(p[0]);
      sh.getRange(nextRow, colIndex).setValue(p[1]);
    });

    Logger.log('[appendToClientStatusReportStrict_] Wrote %s fields to row %s on "%s" in %s',
      writePairs.length, nextRow, CLIENT_STATUS_TAB, targetSS.getName());
  }

  // -------------------- Helpers --------------------

  function getHeaderMap_(sheet) {
    var headerRow = 1;
    var lastCol = sheet.getLastColumn();
    var raw = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
    var names = [];
    var colIndexByName = {};

    for (var c = 0; c < raw.length; c++) {
      var name = String(raw[c] || '').trim();
      if (!name) continue;
      names.push(name);
      colIndexByName[name] = c + 1; // 1-based
    }

    function find(name) {
      if (!name) return null;
      if (colIndexByName[name]) return colIndexByName[name];
      var lower = String(name).toLowerCase();
      for (var k in colIndexByName) {
        if (String(k).toLowerCase() === lower) return colIndexByName[k];
      }
      return null;
    }

    Logger.log('[getHeaderMap_] Main sheet headers: %s', JSON.stringify(names));
    return { names: names, colIndexByName: colIndexByName, find: find };
  }

  function assertHeadersPresent_(headers) {
    var required = [
      '3D Deadline',
      '# of Times 3D Deadline Moved',
      'Production Deadline',
      '# of Times Prod. Deadline Moved',
      'Client Status Report URL'
    ];
    var missing = [];
    for (var i = 0; i < required.length; i++) {
      if (!headers.find(required[i])) missing.push(required[i]);
    }
    if (missing.length) {
      Logger.log('[assertHeadersPresent_] Missing in main sheet: %s', missing.join(', '));
      throw new Error('Missing required column(s): ' + missing.join(', ') + '. Please add these headers in row 1.');
    }
  }

  function getCellValueByHeader_(rowValues, headers, name) {
    var col = headers.find(name);
    return col ? rowValues[col - 1] : '';
  }

  function parseIsoDateAtMidnight_(iso, tz) {
    var parts = iso.split('-'); // YYYY-MM-DD
    var year = +parts[0], month = +parts[1], day = +parts[2];
    return new Date(year, month - 1, day, 0, 0, 0);
  }

  function serializeDateForInput_(value, tz) {
    var d = coerceToDate_(value);
    if (!d) return '';
    var yyyy = String(d.getFullYear());
    var mm = ('0' + (d.getMonth() + 1)).slice(-2);
    var dd = ('0' + d.getDate()).slice(-2);
    return yyyy + '-' + mm + '-' + dd;
  }

  function coerceToDate_(value) {
    if (!value) return null;
    if (Object.prototype.toString.call(value) === '[object Date]') return value;
    if (typeof value === 'number') return new Date(value);
    if (typeof value === 'string') {
      var iso = value.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (iso) return new Date(+iso[1], +iso[2] - 1, +iso[3], 0, 0, 0);
      var d = new Date(value);
      return isNaN(d.getTime()) ? null : d;
    }
    return null;
  }

  function coerceToNumber_(value) {
    if (typeof value === 'number') return value;
    if (typeof value === 'string') {
      var n = parseFloat(value.replace(/[^0-9.\-]/g, ''));
      return isNaN(n) ? NaN : n;
    }
    return NaN;
  }

  function datesEqual_(a, b, tz) {
    if (!a && !b) return true;
    if (!a || !b) return false;
    return a.getFullYear() === b.getFullYear() &&
           a.getMonth() === b.getMonth() &&
           a.getDate() === b.getDate();
  }

  function isNonEmpty_(s) {
    return !!(s && String(s).trim().length > 0);
  }

  function getActiveUserEmail_() {
    try { return Session.getActiveUser().getEmail(); } catch (e) {}
    try { return Session.getEffectiveUser().getEmail(); } catch (e) {}
    return '';
  }

  function extractFileIdFromUrl_(url) {
    if (!url) return '';
    var m = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    return m ? m[1] : '';
  }
  function findFirstHeaderIndex_(headers, candidateNames) {
    for (var i = 0; i < candidateNames.length; i++) {
      var idx = headers.find(candidateNames[i]);
      if (idx) return idx;
    }
    return null;
  }
  function getMaybeRowValueByCandidates_(rowValues, headers, candidateNames) {
    var idx = findFirstHeaderIndex_(headers, candidateNames);
    return idx ? (rowValues[idx - 1] || '') : '';
  }

  function createHeaderMapFromRow_(rowArray) {
    var names = [];
    var colIndexByName = {};
    for (var i = 0; i < rowArray.length; i++) {
      var name = String(rowArray[i] || '').trim();
      if (!name) continue;
      names.push(name);
      colIndexByName[name] = i + 1; // 1-based
    }
    function find(name) {
      if (!name) return null;
      if (colIndexByName[name]) return colIndexByName[name];
      var lower = String(name).toLowerCase();
      for (var k in colIndexByName) {
        if (String(k).toLowerCase() === lower) return colIndexByName[k];
      }
      return null;
    }
    Logger.log('[createHeaderMapFromRow_] Row %s headers: %s', CLIENT_STATUS_HEADER_ROW, JSON.stringify(names));
    return { names: names, colIndexByName: colIndexByName, find: find };
  }

  // Public API
  return {
    showRecordDeadlineDialog: showRecordDeadlineDialog,
    getRecordDeadlineInit: getRecordDeadlineInit,
    saveRecordDeadline: saveRecordDeadline
  };
})();

/**
 * -------------------- TOP-LEVEL WRAPPERS (needed for google.script.run) --------------------
 * These wrappers make the module functions callable from HTML Service.
 * Also add explicit Logger lines so you can see each execution in the Logs UI.
 */

function showRecordDeadlineDialog() {
  Logger.log('[PUBLIC:showRecordDeadlineDialog] Entry');
  return Deadlines.showRecordDeadlineDialog();
}

function getRecordDeadlineInit() {
  Logger.log('[PUBLIC:getRecordDeadlineInit] Entry');
  return Deadlines.getRecordDeadlineInit();
}

function saveRecordDeadline(payload) {
  Logger.log('[PUBLIC:saveRecordDeadline] Entry payload=%s', JSON.stringify(payload));
  return Deadlines.saveRecordDeadline(payload);
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



