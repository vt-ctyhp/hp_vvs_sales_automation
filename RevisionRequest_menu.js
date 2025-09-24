/** File: 03 - RevisionRequest_menu.gs
 * Purpose: Keep only the menu/opener for the 3D Revision Request feature.
 * Notes: Server logic (getRevisionPrefill, save3DRevision, save3DRevisionCore_)
 *        already exists and remains unchanged.
 */


/** Menu handler (Sales → 3D Revision Request). */
function open3DRevision() {      // DO NOT rename; used by the Sales menu builder
  const html = HtmlService.createHtmlOutputFromFile('dlg_revision3d_v1')
    .setWidth(650).setHeight(680).setTitle('3D Revision Request');
  SpreadsheetApp.getUi().showModalDialog(html, '3D Revision Request');
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



