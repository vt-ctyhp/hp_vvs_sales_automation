/** Toggle debug on/off. Set to false after you’re done. */
const REMIND_DEBUG = true;

/** Core debug logger: Logger.log + append to 10_Automation_Log if present. */
function remind__dbg(tag, data) {
  if (!REMIND_DEBUG) return;
  try {
    const msg = tag + (data ? (' ' + JSON.stringify(data)) : '');
    Logger.log('[REMIND] ' + msg);
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('10_Automation_Log');
    if (sh) {
      sh.appendRow([
        Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Los_Angeles', 'yyyy-MM-dd HH:mm:ss'),
        'REMIND_DEBUG', tag, msg
      ]);
    }
  } catch (err) {
    // last resort: logger only
    Logger.log('[REMIND] debug write failed: ' + (err && err.message ? err.message : err));
  }
}

/** Quick sanity test you can run from the editor. */
function remind__debugSmokeTest() {
  remind__dbg('smoke', { now: new Date().toISOString() });
  SpreadsheetApp.getActive().toast('Wrote a debug line (check Executions / 10_Automation_Log).', 'Reminders', 5);
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



