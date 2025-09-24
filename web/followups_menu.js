/**
 * 16 - followups_menu.gs
 * Adds a top-level "Follow-ups" menu with Snooze/Cancel actions for the active SO row.
 * Uses installable onOpen trigger (created by Setup.installTriggers), so we don't touch your existing onOpen.
 */

function followups_menu_onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Follow-ups')
    .addItem('Snooze reminders for this SO…', 'followups_menu_snoozeActiveSO')
    .addItem('Cancel reminders for this SO', 'followups_menu_cancelActiveSO')
    .addToUi();
}

function followups_menu_snoozeActiveSO() {
  const so = _activeRowSO_();
  if (!so) { SpreadsheetApp.getUi().alert('SO# not found on this row.'); return; }
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Snooze until (YYYY-MM-DD)', 'Example: 2025-10-15', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const iso = resp.getResponseText().trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) { ui.alert('Please use YYYY-MM-DD.'); return; }
  Remind.snoozeForSO(so, iso);
  ui.alert('Snoozed reminders for SO#' + so + ' until ' + iso + ' @ 9:30 AM PT.');
}

function followups_menu_cancelActiveSO() {
  const so = _activeRowSO_();
  if (!so) { SpreadsheetApp.getUi().alert('SO# not found on this row.'); return; }
  Remind.cancelForSO(so);
  SpreadsheetApp.getUi().alert('Cancelled active reminders for SO#' + so + '.');
}

function _activeRowSO_() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== REMIND.ORDERS_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('Switch to sheet: ' + REMIND.ORDERS_SHEET_NAME);
    return '';
  }
  const row = sh.getActiveRange().getRow();
  if (row <= 1) return '';
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  const soCol = headers.findIndex(h => (h||'').trim().toLowerCase() === REMIND.COL_SO.toLowerCase()) + 1;
  if (soCol <= 0) return '';
  return (sh.getRange(row, soCol).getDisplayValue() || '').trim();
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



