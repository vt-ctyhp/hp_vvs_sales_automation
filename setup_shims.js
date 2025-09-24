/** Top-level wrappers so you can run setup from the toolbar */
function remind__runDailyNowForTesting(){ 
  return Remind.runDailyNowForTesting(); 
    }

function remind__dev_enqueueCOS(){ Remind.scheduleCOS('SO_TEST', {customerName:'Test Client', assignedRepName:'Rep A', nextSteps:'Ping R&D'}, false); }

// Paste anywhere (e.g., setup_shims.gs)
function remind__debugListCOSForSO__RUN() {
  remind__debugListCOSForSO('SO12345');  // put your SO here
}
function remind__dumpQueueForCustomer__RUN(){
  // TODO: replace with the exact customer name you want to inspect
  remind__dumpQueueForCustomer('Client Alpha');
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



