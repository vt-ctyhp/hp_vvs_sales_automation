/**
 * Sales workflow sheet repository: spreadsheet access and low-level sheet helpers.
 */

function swSpreadsheet_() {
  var id = '';
  try {
    id = swTrim_(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || PropertiesService.getScriptProperties().getProperty('MASTER_FILE_ID'));
  } catch (_) {}
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (_) {}
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet and no SPREADSHEET_ID script property.');
  return ss;
}

function swRequireWorkflowReadSheets_(ss, options) {
  options = options || {};
  var names = [SW_SHEETS.TASKS, SW_SHEETS.CONFIG];
  if (options.templates !== false) names.push(SW_SHEETS.TEMPLATES);
  names.forEach(function (name) {
    swGetRequiredSheet_(ss, name);
  });
}

function swGetRequiredSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name + '. Run sw_setupSalesWorkflow first.');
  return sh;
}

function swEnsureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    var existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
    var col = existing.length;
    headers.forEach(function (h) {
      if (existing.indexOf(h) < 0) {
        col++;
        sh.getRange(1, col).setValue(h);
      }
    });
  }
  return sh;
}

function swStyleSheet_(sh) {
  try {
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#EFE8DD').setFontColor('#2A2725');
    sh.autoResizeColumns(1, Math.min(sh.getLastColumn(), 12));
  } catch (_) {}
}

function swReadSheetObjects_(sh) {
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < 1) return [];
  var values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var obj = { __rowNumber: i + 1 };
    var blank = true;
    for (var j = 0; j < headers.length; j++) {
      if (!headers[j]) continue;
      obj[headers[j]] = values[i][j];
      if (values[i][j] !== '') blank = false;
    }
    if (!blank) out.push(obj);
  }
  return out;
}

function swReadSheetObjectsExpectedHeaders_(sh, expectedHeaders) {
  expectedHeaders = expectedHeaders || [];
  if (!sh || !expectedHeaders.length) return [];

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var values;
  try {
    values = sh.getRange(1, 1, lastRow, expectedHeaders.length).getDisplayValues();
  } catch (_) {
    return swReadSheetObjects_(sh);
  }

  var actualHeaders = values[0].map(function (h) { return swTrim_(h); });
  for (var h = 0; h < expectedHeaders.length; h++) {
    if (swHeaderKey_(actualHeaders[h]) !== swHeaderKey_(expectedHeaders[h])) {
      return swReadSheetObjects_(sh);
    }
  }

  var out = [];
  for (var i = 1; i < values.length; i++) {
    var obj = { __rowNumber: i + 1 };
    var blank = true;
    for (var j = 0; j < expectedHeaders.length; j++) {
      obj[expectedHeaders[j]] = values[i][j];
      if (values[i][j] !== '') blank = false;
    }
    if (!blank) out.push(obj);
  }
  return out;
}

function swReadSelectedRows_(sh, startRow, rowCount, indexes, mode) {
  rowCount = Number(rowCount) || 0;
  if (!sh || rowCount <= 0) return [];
  var columns = swSelectedColumnIndexes_(indexes);
  if (!columns.length) return [];

  var out = [];
  for (var i = 0; i < rowCount; i++) out.push([]);
  var minCol = columns[0];
  var maxCol = columns[columns.length - 1];
  var span = maxCol - minCol + 1;
  var readDisplay = mode !== 'values';

  if (columns.length >= 30 || span <= (columns.length * 2) + 8) {
    var block = readDisplay
      ? sh.getRange(startRow, minCol + 1, rowCount, span).getDisplayValues()
      : sh.getRange(startRow, minCol + 1, rowCount, span).getValues();
    for (var r = 0; r < block.length; r++) {
      for (var c = 0; c < columns.length; c++) {
        out[r][columns[c]] = block[r][columns[c] - minCol];
      }
    }
    return out;
  }

  var groups = [];
  var current = null;
  columns.forEach(function (col) {
    if (!current || col > current.end + 3) {
      current = { start: col, end: col, columns: [col] };
      groups.push(current);
      return;
    }
    current.end = col;
    current.columns.push(col);
  });

  groups.forEach(function (group) {
    var width = group.end - group.start + 1;
    var values = readDisplay
      ? sh.getRange(startRow, group.start + 1, rowCount, width).getDisplayValues()
      : sh.getRange(startRow, group.start + 1, rowCount, width).getValues();
    for (var r = 0; r < values.length; r++) {
      for (var c = 0; c < group.columns.length; c++) {
        out[r][group.columns[c]] = values[r][group.columns[c] - group.start];
      }
    }
  });
  return out;
}

function swSelectedColumnIndexes_(indexes) {
  var seen = {};
  var out = [];
  (indexes || []).forEach(function (idx) {
    idx = Number(idx);
    if (!isFinite(idx) || idx < 0 || seen[idx]) return;
    seen[idx] = true;
    out.push(idx);
  });
  out.sort(function (a, b) { return a - b; });
  return out;
}
