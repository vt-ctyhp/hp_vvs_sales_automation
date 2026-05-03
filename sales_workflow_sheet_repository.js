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
