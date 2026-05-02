/** File: 02 - AssignSO_menu.gs
 * Purpose: Keep only the menu/openers for the Assign SO feature.
 * Notes: The server logic (getActiveMasterPreview, checkSOConflicts, saveAssignedSO)
 *        already exists and remains unchanged.
 */

/** Menu handler (Sales → Assign SO). */
function assignSO() {            // DO NOT rename; used by the Sales menu builder
  openAssignSO();
}

/** Opens the Assign SO dialog. */
function openAssignSO() {
  // Try to precompute preview; if user isn’t on a valid row, fall back gracefully
  let preview = null;
  try {
    preview = getActiveMasterPreview();
  } catch (e) {
    preview = null; // fall back to client call; preserves behavior
  }

  const t = HtmlService.createTemplateFromFile('dlg_assign_so_v1');
  t.designRequest = '';  // unchanged behavior
  t.preview = preview;   // NEW: inject preview if available

  const html = t.evaluate().setWidth(520).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'Assign Sales Order');
}

/** Opens Assign SO and (optionally) injects a prebuilt preview object. */
function openAssignSOWithPreview(preview) {
  const t = HtmlService.createTemplateFromFile('dlg_assign_so_v1');  // keep same HTML
  t.designRequest = '';    // unchanged behavior
  t.preview = preview || null;  // <— NEW: inject prebuilt preview if provided
  const html = t.evaluate().setWidth(520).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'Assign Sales Order');
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



