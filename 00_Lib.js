/* ---------- [HELPERS/PROP] ---------- */
function PROP_(k, def){ return PropertiesService.getScriptProperties().getProperty(k) || def || ''; }
function MASTER_SS_(){ return SpreadsheetApp.openById(PROP_('SPREADSHEET_ID')); }
function DEFAULT_TZ_(){ return PROP_('DEFAULT_TZ','America/Los_Angeles'); }
function OPENAI_MODEL_(){ return PROP_('OPENAI_MODEL','gpt-5'); }


// Drop-in: tolerant ID extraction from any Drive URL/string
function idFromAnyGoogleUrl_(s) {
  s = String(s || '').trim();
  // 1) Try the common /d/<id> pattern
  let m = s.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m) return m[1];
  // 2) Try query param id=<id>
  m = s.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  if (m) return m[1];
  // 3) Fallback: any 25+ char Drive-ish ID anywhere in the string
  m = s.match(/[-\w]{25,}/);
  return m ? m[0] : '';
}

// Lighten hex (0..1), returns hex
function lighten_(hex, amt){
  hex = (hex||'').replace('#','');
  if (hex.length!==6) return '#f0f3f7';
  const pct = Math.max(0, Math.min(1, amt||0.2));
  const to = (c)=> {
    const v = parseInt(c,16); const out = Math.round(v + (255 - v)*pct);
    return ('0'+out.toString(16)).slice(-2);
  };
  return '#' + to(hex.slice(0,2)) + to(hex.slice(2,4)) + to(hex.slice(4,6));
}


/* ---------- [UTIL] MD5 hex ---------- */
function md5Hex_(s){
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(s), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + ((b<0?b+256:b) & 0xff).toString(16)).slice(-2)).join('');
}


/* ---------- [OBJ UTILS] setByPath_ / mergeDeep_ ---------- */
function setByPath_(obj, path, value) {
  const parts = String(path||'').split('.').filter(Boolean);
  let cur = obj;
  for (let i=0;i<parts.length-1;i++){
    const k = parts[i];
    if (typeof cur[k] !== 'object' || cur[k] === null) cur[k] = {};
    cur = cur[k];
  }
  cur[parts[parts.length-1]] = value;
}

function mergeDeep_(base, patch){
  if (Array.isArray(base) || Array.isArray(patch)) return patch;
  if (typeof base !== 'object' || base === null) return patch;
  if (typeof patch !== 'object' || patch === null) return patch;
  const out = JSON.parse(JSON.stringify(base));
  Object.keys(patch).forEach(k=>{
    out[k] = (k in out) ? mergeDeep_(out[k], patch[k]) : patch[k];
  });
  return out;
}

/* ---------- [DRIVE/NEWEST] ---------- */
function newestByRegexInFolder_(folder, re){
  let newest=null, ts=0, it=folder.getFiles();
  while (it.hasNext()){
    const f = it.next();
    if (!re.test(f.getName())) continue;
    const t = (f.getLastUpdated ? f.getLastUpdated().getTime()
                                : f.getDateCreated().getTime());
    if (t>ts){ ts=t; newest=f; }
  }
  return newest;
}


/* ---------- [AP-FOLDER/LOOKUP] Find AP folder id via Master ---------- */
function getApFolderIdForRoot_(ss, rootApptId){
  const sh = ss.getSheetByName('00_Master Appointments');
  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
  const iRoot = hdr.indexOf('RootApptID');
  const iId   = hdr.indexOf('RootAppt Folder ID');     // NEW
  const iUrl  = hdr.indexOf('RootAppt Folder URL');    // existing
  if (iRoot < 0) throw new Error('RootApptID column not found on Master');

  const last = sh.getLastRow(); if (last < 2) throw new Error('Master is empty');
  const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  const row  = vals.find(r => String(r[iRoot]||'').trim() === String(rootApptId).trim());
  if (!row) throw new Error('RootApptID not found on Master: ' + rootApptId);

  if (iId >= 0 && row[iId]) return String(row[iId]).trim(); // direct ID
  if (iUrl >= 0 && row[iUrl]) return idFromAnyGoogleUrl_(String(row[iUrl]).trim());
  throw new Error('No RootAppt folder reference found for ' + rootApptId);
}

/** Find Client Status Report ID for a RootApptID from Master. */
function getReportIdForRoot_(rootApptId){
  const ss = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  const sh = ss.getSheetByName('00_Master Appointments');
  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
  const iRoot = hdr.indexOf('RootApptID');
  const iUrl  = hdr.indexOf('Client Status Report URL'); // column exists in your ClientStatus code
  if (iRoot < 0 || iUrl < 0) throw new Error('Missing RootApptID / Client Status Report URL');
  const last = sh.getLastRow(); if (last < 2) throw new Error('Master empty');
  const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][iRoot]||'').trim() === rootApptId){
      const reportUrl = String(vals[i][iUrl]||'').trim();
      const id = idFromAnyGoogleUrl_(reportUrl);
      if (!id) throw new Error('Bad Client Status Report URL for ' + rootApptId + ': ' + reportUrl);
      return id;
    }
  }
  throw new Error('RootApptID not found: '+rootApptId);
}




/** Single Web‑App entrypoint: routes to Upload vs AskController by content type / action */
function doPost(e) {
  var ct = (e && e.postData && e.postData.type) || '';
  var p  = (e && e.parameter) || {};

  // Ask/Chat JSON posts from the sidebar or external tools
  if (/^application\/json/i.test(ct) || p.action === 'chat' || p.action === 'apply_patch') {
    return AC_doPost_(e);
  }

  // Everything else (raw audio, multipart form-data, or octet-stream) → upload intake
  return doPost_UPLOAD_(e);
}

/** Return the deployed /exec URL for this web app. */
function WEBAPP_EXEC_URL_(){
  // Prefer an explicit Script Property if you pasted the exec URL there
  var prop = PropertiesService.getScriptProperties().getProperty('WEBAPP_EXEC_URL');
  if (prop && /^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/i.test(prop)) return prop.trim();
  // Fallback: try Deployment API (Apps Script) else last-resort to ScriptApp.getService().getUrl()
  return ScriptApp.getService().getUrl().replace(/\/dev(\b|$)/,'/exec');
}


/** Controller URL resolver: prefer pinned WEBAPP_EXEC_URL, else fall back to ScriptApp and force /exec */
function AC_controllerUrl_(){
  var pinned = PropertiesService.getScriptProperties().getProperty('WEBAPP_EXEC_URL');
  if (pinned && pinned.trim()) return pinned.trim();

  // Fallback: current script’s web-app URL, normalized to /exec if it happens to be /dev
  var u = (ScriptApp.getService().getUrl && ScriptApp.getService().getUrl()) || '';
  if (u.endsWith('/dev')) u = u.slice(0, -4) + 'exec';
  return u;
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



