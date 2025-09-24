/********** CONFIG **********/

const SHT = {
MASTER: '00_Master Appointments',
FORM_INBOX: '02_Form_Inbox',
LOG: '20_Automation_Log',
ERR: '90_Validation_Errors'
};

function PROP_(k, def){ return PropertiesService.getScriptProperties().getProperty(k) || def || ''; }

const CFG = {
 TZ: PROP_('DEFAULT_TZ','America/Los_Angeles'),
 HP_ROOT: PROP_('HP_CLIENTS_ROOT_ID',''),
 VVS_ROOT: PROP_('VVS_CLIENTS_ROOT_ID',''),
 INTAKE_TPL: PROP_('INTAKE_TEMPLATE_ID',''),
 DEBUG: /true/i.test(PROP_('DEBUG','false'))
};

// --- Lightweight profiler (add once for debugging) ---
var __t0 = 0, __last = 0;
function __startProfile(label){
  __t0 = Date.now(); __last = __t0;
  Logger.log('▶ ' + label + ' @ ' + new Date(__t0).toISOString());
}
function __mark(label){
  const now = Date.now();
  Logger.log('⏱ ' + label + '  +' + (now - __t0) + 'ms  (Δ' + (now - __last) + 'ms)');
  __last = now;
}
function __wrap(label, fn){                 // times a block
  const t = Date.now();
  Logger.log('→ ' + label);
  try { return fn(); }
  finally { Logger.log('← ' + label + '  ' + (Date.now()-t) + 'ms'); }
}

function debug_introspectHelpers(){
  Logger.log('_findMostRecentPriorRow.length = ' + _findMostRecentPriorRow.length);
  Logger.log('_currentRowToObj_.length = ' + _currentRowToObj_.length);
}


/********** SHEET HELPERS **********/
function SS(){ return SpreadsheetApp.getActive(); }
function SH(name){ const s=SS().getSheetByName(name); if(!s) throw new Error(`Missing sheet: ${name}`); return s; }
function headers_(name){ const s=SH(name); const arr=s.getRange(1,1,1,s.getLastColumn()).getValues()[0]; const map={}; arr.forEach((h,i)=>{ if(h) map[String(h).trim()]=i+1; }); return map; }
function setCell_(sheetName,row,colName,val){ const m=headers_(sheetName); const c=m[colName]; if(!c) throw new Error(`Column "${colName}" not found on ${sheetName}`); SH(sheetName).getRange(row,c).setValue(val); }
function getCell_(sheetName,row,colName){ const m=headers_(sheetName); const c=m[colName]; if(!c) return ''; return SH(sheetName).getRange(row,c).getValue(); }
function appendObj_(sheetName, obj){
  const s = SH(sheetName), H = headers_(sheetName);
  const rowArr = new Array(s.getLastColumn()).fill('');

  Object.keys(obj || {}).forEach(k => { if (H[k]) rowArr[H[k]-1] = obj[k]; });

  if (sheetName === SHT.MASTER) {
    // Use sentinel-based last-data-row so formulas below don’t push us down
    const r = nextDataRow_(sheetName, LASTROW_SENTINELS);
    s.getRange(r, 1, 1, rowArr.length).setValues([rowArr]);
    return r;
  } else {
    // For non-Master sheets (logs, errors), regular append is fine
    s.appendRow(rowArr);
    return s.getLastRow();
  }
}
function log_(action, details){ appendObj_(SHT.LOG, {'Timestamp': new Date(), 'Action': action, 'Details': typeof details==='string'?details:JSON.stringify(details)}); }
function err_(where, why, payload){ appendObj_(SHT.ERR, {'Timestamp': new Date(), 'Where': where, 'Why': why, 'Payload': JSON.stringify(payload||{})}); }


// Put near the top of Resolver.gs
function nvGet(nv, key){
if (nv[key] && nv[key][0] !== undefined) return nv[key][0];
const k = Object.keys(nv || {}).find(k => k && k.trim().toLowerCase() === key.trim().toLowerCase());
return k ? (nv[k][0] || '') : ''
;
}


function setOnce_(sheetName, row, colName, value){
const cur = getCell_(sheetName, row, colName);
if (!cur && value) setCell_(sheetName, row, colName, value);
}

function findMasterRowByUID_(uuid){
if (!uuid) return 0;
const s = SH(SHT.MASTER), m = headers_(SHT.MASTER);
const col = m['CalendlyEventUID']; if (!col) return 0;
const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS); if (last < 2) return 0;
const vals = s.getRange(2, col, last - 1, 1).getValues().flat();
const idx = vals.findIndex(v => String(v||'') === String(uuid));
return idx < 0 ? 0 : idx + 2;
}

/** Shared: build a stable contact key (prefer email; fallback phone) */
function _contactKey_(brand, vtype, emailLower, phoneNorm){
const b=(brand||'').toUpperCase().trim();
const t=(vtype||'').toLowerCase().trim();
const e=(emailLower||'').toLowerCase().trim();
const p=(phoneNorm||'').trim();
const id = e || p; // prefer email
return ['CANCEL', b, t, id].join(':');
}

/** Shared: remember the UID of the old (canceled) event for this contact */
function _rememberCancelUID_(brand, vtype, emailLower, phoneNorm, oldUid, ttlSec){
try{
const key = _contactKey_(brand, vtype, emailLower, phoneNorm);
CacheService.getScriptCache().put(key, String(oldUid||''), ttlSec || 7200); // 2h
return key;
}catch(_){ return ''; }
}

/** Shared: retrieve & delete any pending canceled UID for this contact */
function _popPendingCancelUID_(brand, vtype, emailLower, phoneNorm){
try{
const key = _contactKey_(brand, vtype, emailLower, phoneNorm);
const cache = CacheService.getScriptCache();
const uid = cache.get(key);
if (uid) cache.remove(key);
return uid || '';
}catch(_){ return ''; }
}

const RFLAGS = { REUSE_ARTIFACTS_FROM_PRIOR: true, PRIOR_LOOKBACK_DAYS: 0 };



// Build a stable “same person” key (adjust column names if yours differ)
function _samePersonKey(row) {
const e = String(row['EmailLower'] || '').trim().toLowerCase();
const p = String(row['PhoneNorm']  || '').trim();
return e || p ? (e + '|' + p) : '';
}


function findRecentCanceledRowByContact_(emailLower, phoneNorm, minutes=240){
const s=SH(SHT.MASTER), m=headers_(SHT.MASTER);
const colE=m['EmailLower'], colP=m['PhoneNorm'], colSta=m['Status'], colT=m['ApptDateTime (ISO)'];
if (!colE || !colP || !colSta) return 0;
const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS); if (last < 2) return 0;
const rows=s.getRange(2,1,last-1,s.getLastColumn()).getValues();
const now=Date.now(), win=minutes*60*1000;
for (let i=rows.length-1;i>=0;i--){
const r=rows[i];
const e=(r[colE-1]||'').toString().toLowerCase();
const p=(r[colP-1]||'').toString();
const sta=(r[colSta-1]||'').toString();
const t=colT ? (r[colT-1]? new Date(r[colT-1]).getTime() : now) : now;
const contactMatch = (emailLower && e===emailLower) || (phoneNorm && p===phoneNorm);
if (contactMatch && /canceled/i.test(sta) && (now - t) <= win) return i+2;
}
return 0;
}

// Find the most recent Canceled row for this contact, based on CanceledAt (preferred)
function findRecentCanceledByContactAt_(emailLower, phoneNorm, minutes){
const s = SH(SHT.MASTER), m = headers_(SHT.MASTER);
const colE=m['EmailLower'], colP=m['PhoneNorm'], colSta=m['Status'], colCA=m['CanceledAt'];
if (!colE || !colP || !colSta || !colCA) return 0;
const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS); if (last < 2) return 0;
const vals=s.getRange(2,1,last-1,s.getLastColumn()).getValues();
const now=Date.now(), win=(minutes||120)*60*1000; // default 120 minutes
for (let i=vals.length-1;i>=0;i--){
const r=vals[i];
const e=(r[colE-1]||'').toString().toLowerCase();
const p=(r[colP-1]||'').toString();
const sta=(r[colSta-1]||'').toString();
const ca = r[colCA-1] ? new Date(r[colCA-1]).getTime() : 0;
const matchContact = (emailLower && e===emailLower) || (phoneNorm && p===phoneNorm);
if (matchContact && /canceled/i.test(sta) && ca && (now - ca) <= win) return i+2;
}
return 0;
}

// --- Source normalization (simple map; extend as needed)
const SOURCE_MAP = {
'instagram':'Instagram',
'tiktok':'TikTok',
'facebook':'Facebook','fb':'Facebook',
'google':'Google','search':'Google','google ads':'Google',
'yelp':'Yelp',
'referral':'Referral','friend':'Referral',
};
function normSource_(raw){
const k = (raw||'').toString().trim().toLowerCase();
return SOURCE_MAP[k] || raw || '';
}
// --- Split full name
function splitName_(full){
const t = (full||'').toString().trim();
if (!t) return {first:'', last:''};
const parts = t.split(/\s+/);
return {first: parts[0], last: parts.slice(1).join(' ')};
}
// --- Safer default duration
const DEFAULT_DURATION_MIN = 30;



/********** NORMALIZERS **********/
function normEmail_(e){ return (e||'').toString().trim().toLowerCase(); }
function normPhone_(p){
if(!p) return ''; const d=(''+p).replace(/\D+/g,'');
if(d.length===10) return '+1'+d;
if(d.length===11 && d.startsWith('1')) return '+'+d;
return d?('+'+d):'';
}
function locToEnum_(loc){
const s=(loc||'').toString().toLowerCase();
if(/virtual|zoom|google meet|video/.test(s)) return 'Virtual';
if(/store|in[-\s]?store|in person|walk/.test(s)) return 'In Store';
return loc||'';
}
function brandFromCompany_(company){
const s=(company||'').toString().toUpperCase();
if (s.includes('VVS')) return 'VVS';
if (s.includes('HP')) return 'HPUSA';
return ''; // allow empty → can be corrected later
}
function parseBudget_(raw){
if(!raw) return {min:'',max:''};
const picks = (''+raw).split(';').map(s=>s.trim()).filter(Boolean);
if (picks.length!==1) return {min:'',max:''};
const m = picks[0].match(/\$?\s*([\d,]+)\s*[-–]\s*\$?\s*([\d,]+)/);
if(!m) return {min:'',max:''};
const toNum = s => Number(String(s).replace(/[^\d]/g,''))||'';
return {min: toNum(m[1]), max: toNum(m[2])};
}



/********** ID / APPT_ID **********/
function nextApptId_(iso){
const tz = CFG.TZ, dt = iso ? new Date(iso) : new Date();
const ymd = Utilities.formatDate(dt, tz, 'yyyyMMdd');
const s = SH(SHT.MASTER), m = headers_(SHT.MASTER);
const col = m['APPT_ID'] || 0;
if (!col) return `AP-${ymd}-001`;

const lastRow = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
let vals = [];
if (lastRow >= 2){
  vals = s.getRange(2, col, lastRow - 1, 1).getValues().flat().filter(Boolean);
}
const countToday = vals.filter(v => String(v).startsWith('AP-'+ymd)).length + 1;
return `AP-${ymd}-${String(countToday).padStart(3,'0')}`;
}

/********** DRIVE HELPERS **********/
function brandRoot_(brand){
if(brand==='HPUSA' && CFG.HP_ROOT) return DriveApp.getFolderById(CFG.HP_ROOT);
if(brand==='VVS'   && CFG.VVS_ROOT) return DriveApp.getFolderById(CFG.VVS_ROOT);
throw new Error(`No brand root configured for ${brand}`);
}
function getOrCreate_(parent, name){
const it=parent.getFoldersByName(name);
return it.hasNext()? it.next() : parent.createFolder(name);
}
function ensureClientFolder_(brand, customerName, phoneNorm, emailLower){
const root = brandRoot_(brand);
const safe = String(customerName || emailLower || phoneNorm || 'Unknown')
            .trim()
            .replace(/[\\/:*?"<>|]/g, '-');  // sanitize for folder names
const it = root.getFoldersByName(safe);
return it.hasNext() ? it.next() : root.createFolder(safe);
}

function ensureProspectFolder_(clientFolder, apptId){
const prospects = getOrCreate_(clientFolder, 'Prospects');
const name = `${apptId} (NO-SO-YET)`;
const it = prospects.getFoldersByName(name);
return it.hasNext()? it.next() : prospects.createFolder(name);
}
function cloneIntakeDoc_(destFolder, brand, apptId){
if(!CFG.INTAKE_TPL) return '';
const file = DriveApp.getFileById(CFG.INTAKE_TPL);
const copy = file.makeCopy(`${brand} – ${apptId} – Intake`, destFolder);
return copy.getUrl();
}

function _appendNote_(row, msg){
const prev = getCell_(SHT.MASTER,row,'Automation Notes') || '';
setCell_(SHT.MASTER,row,'Automation Notes', (prev?prev+'\n':'') + msg);
}

function ensureApptSubfolders_(rootApptId, apFolder) {
  ['01_Audio','02_Design','03_Transcripts','04_AI_Summaries','05_ChatLogs']
    .forEach(name => {
      const it = apFolder.getFoldersByName(name);
      if (!it.hasNext()) apFolder.createFolder(name);
    });
}

function _ensureApSubfoldersByFolderId_(apFolderId) {
  const apFolder = DriveApp.getFolderById(apFolderId);
  ['01_Audio','02_Design','03_Transcripts','04_AI_Summaries','05_ChatLogs'].forEach(name => {
    const it = apFolder.getFoldersByName(name);
    if (!it.hasNext()) apFolder.createFolder(name);
  });
  return apFolder;
}

function bootstrapApptFolder_(rowIdx) {
  // Prefer the robust, ID-writing path
  try {
    return bootstrapApFolderForRow_(rowIdx);
  } catch (e) {
    // Fallback to legacy behavior if needed
    const apId = getCell_(SHT.MASTER,rowIdx,'APPT_ID');
    if (!apId) return;
    const pfId = getCell_(SHT.MASTER,rowIdx,'ProspectFolderID');
    if (!pfId) return;

    const apFolder = DriveApp.getFolderById(pfId);
    ensureApptSubfolders_(apId, apFolder);
    setCell_(SHT.MASTER,rowIdx,'RootAppt Folder ID', apFolder.getId());
    return apFolder.getId();
  }
}

/***** Root Appt Folder bootstrap — v1.1 *****
 * Ensures [RootApptID] folder has numbered subfolders and writes its ID to Master.
 * Works whether [RootApptID] currently lives under Prospects or under 00-Intake;
 * we only rely on the folder's Drive ID (stable when moved).
 */

// === Sheet + headers helpers (reuses your pattern) ===
const _SP = PropertiesService.getScriptProperties();

function _openMaster_() {
  const id = _SP.getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Missing SPREADSHEET_ID script property');
  const ss = SpreadsheetApp.openById(id);
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet "00_Master Appointments"');
  return sh;
}

function _headers_(sh) {
  const row = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  return row.reduce((m,h,i)=>{ if(String(h).trim()) m[String(h).trim()] = i+1; return m; }, {});
}

// === Create subfolders inside an existing AP folder (by ID) ===
function _ensureApSubfoldersByFolderId_(apFolderId) {
  const apFolder = DriveApp.getFolderById(apFolderId);
  ['01_Audio', '02_Design', '03_Transcripts', '04_AI_Summaries'].forEach(name => {
    const it = apFolder.getFoldersByName(name);
    if (!it.hasNext()) apFolder.createFolder(name);
  });
  return apFolder;
}

// === If you want an Appointments home container for new AP folders (optional) ===
function _getAppointmentsHome_() {
  const homeId = _SP.getProperty('APPOINTMENTS_FOLDER_ID');
  if (homeId) return DriveApp.getFolderById(homeId);
  // create once if not configured
  const created = DriveApp.createFolder('[SYS] Appointments');
  _SP.setProperty('APPOINTMENTS_FOLDER_ID', created.getId());
  return created;
}

/**
 * Ensure an AP folder exists (by name) under the Appointments home, for cases where
 * you DON'T already have a prospect folder path created elsewhere.
 */
function _ensureApFolderUnderHome_(apId) {
  const home = _getAppointmentsHome_();
  const it = home.getFoldersByName(apId);
  return it.hasNext() ? it.next() : home.createFolder(apId);
}

/**
 * Write RootAppt Folder ID into Master for the given row.
 */
function _writeApFolderIdToMasterRow_(row, apFolderId) {
  const sh = _openMaster_();
  const H  = _headers_(sh);
  const colFid = H['RootAppt Folder ID'];
  if (!colFid) throw new Error('Missing "RootAppt Folder ID" column in Master');

  sh.getRange(row, colFid).setValue(apFolderId);

  const colNotes = H['Automation Notes'];
  if (colNotes) {
    const prev = sh.getRange(row, colNotes).getValue() || '';
    const add  = `AP folder set: https://drive.google.com/drive/folders/${apFolderId}`;
    sh.getRange(row, colNotes).setValue(prev ? (prev + '\n' + add) : add);
  }
}

/**
 * MAIN: Bootstrap AP folder for a given Master row.
 * Strategy:
 *   1) Get RootApptID from the row.
 *   2) Try to use an existing Prospect/Intake folder ID if your resolver already produced it
 *      (e.g., "ProspectFolderID" or similar column). If present, use that ID.
 *   3) Else, ensure a new AP folder under APPOINTMENTS_FOLDER_ID (optional container).
 *   4) Ensure numbered subfolders, then write "RootAppt Folder ID" back to Master.
 */
function bootstrapApFolderForRow_(row) {
  const sh = _openMaster_();
  const H  = _headers_(sh);

  const colApId = H['RootApptID'];
  if (!colApId) throw new Error('Missing "RootApptID" column in Master');

  const apId = String(sh.getRange(row, colApId).getValue() || '').trim();
  if (!/^AP-\d{8}-\d{3}$/i.test(apId)) {
    throw new Error('Invalid or empty RootApptID on row ' + row + ': ' + apId);
  }

  // 2) Prefer an existing resolver-created folder ID if you store one (adjust the column name if different)
  const prospectIdColNameCandidates = ['ProspectFolderID', 'RootAppt Folder ID', 'AP Folder ID'];
  let colExistingId = null;
  for (const name of prospectIdColNameCandidates) {
    if (H[name]) { colExistingId = H[name]; break; }
  }

  let apFolder;
  if (colExistingId) {
    const existing = String(sh.getRange(row, colExistingId).getValue() || '').trim();
    if (existing) {
      // Use existing folder ID (works even if it has been moved in Drive)
      apFolder = DriveApp.getFolderById(existing);
    }
  }

  // 3) If we still don't have a folder, create it under the Appointments home
  if (!apFolder) {
    apFolder = _ensureApFolderUnderHome_(apId);
  }

  // 4) Ensure numbered subfolders and write ID back
  _ensureApSubfoldersByFolderId_(apFolder.getId());
  _writeApFolderIdToMasterRow_(row, apFolder.getId());

  return apFolder.getId();
}

/***** Auto-bootstrap minute worker — v1.0 *****/

// Runs through the last N rows and bootstraps any row that has RootApptID but no RootAppt Folder ID yet.
function ensureBootstrapForRecentRows_() {
  const sh = _openMaster_();                  // <-- uses your existing helper (bound or Option A from earlier)
  const H  = _headers_(sh);
  const colApId = H['RootApptID'];
  const colFid  = H['RootAppt Folder ID'];
  if (!colApId || !colFid) {
    Logger.log('Missing required headers (RootApptID / RootAppt Folder ID)');
    return;
  }

  const last = sh.getLastRow();
  if (last < 2) return;

  // Limit scan to last N rows for speed; bump if you need.
  const N = Math.min(100, last - 1);
  const startRow = Math.max(2, last - N + 1);

  // Read necessary columns in one batch
  const apIds = sh.getRange(startRow, colApId, N, 1).getValues();
  const fids  = sh.getRange(startRow, colFid,  N, 1).getValues();

  let bootstrapped = 0;
  for (let i = 0; i < N; i++) {
    const row = startRow + i;
    const ap  = String(apIds[i][0] || '').trim();
    const fid = String(fids[i][0]  || '').trim();

    if (!ap) continue;              // skip if no AP-ID yet
    if (fid) continue;              // skip if already has folder ID

    try {
      const id = bootstrapApFolderForRow_(row);     // <-- the function you already added
      Logger.log(`Bootstrapped AP folder ${id} for row ${row}`);
      bootstrapped++;
    } catch (e) {
      Logger.log(`Row ${row}: bootstrap error: ${e && (e.message || e)}`);
      // continue to next row; safe and idempotent
    }
  }

  if (bootstrapped) Logger.log(`ensureBootstrapForRecentRows_: bootstrapped ${bootstrapped} row(s).`);
}

// Install a single minute-based trigger if none exists (mirrors your existing pattern)
function installBootstrapMinuteWorker() {
  const fn = 'ensureBootstrapForRecentRows_';
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === fn);
  if (!exists) {
    ScriptApp.newTrigger(fn).timeBased().everyMinutes(1).create();
    Logger.log('Installed minute worker for ensureBootstrapForRecentRows_()');
  } else {
    Logger.log('Minute worker already installed.');
  }
}


/********** MASTER MERGE **********/
// handy getter by row + header name
function rf(rowIdx, header){
return getCell_(SHT.MASTER, rowIdx, header);
}

function findMasterRowByEmailTime_(emailLower, iso){
if (!emailLower || !iso) return 0;
const s = SH(SHT.MASTER), m = headers_(SHT.MASTER);
const lastRow = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
const lastCol = s.getLastColumn();
if (lastRow < 2 || lastCol < 1) return 0;

const colE = m['EmailLower'], colT = m['ApptDateTime (ISO)'];
if (!colE || !colT) return 0;

const rows = s.getRange(2, 1, lastRow - 1, lastCol).getValues();
const ts = new Date(iso).getTime();
for (let i = 0; i < rows.length; i++){
const r = rows[i];
const e = (r[colE-1]||'').toString().toLowerCase();
const t = r[colT-1] ? new Date(r[colT-1]).getTime() : NaN;
if (e && e === emailLower && !isNaN(t) && Math.abs(t - ts) <= 24*3600*1000) return i + 2;
}
return 0;
}


function countVisits_(emailLower, phoneNorm){
  const s = SH(SHT.MASTER), H = headers_(SHT.MASTER);
  const colE  = H['EmailLower'] || 0;
  const colP  = H['PhoneNorm']  || 0;
  const colSta= H['Status']     || 0;
  const colAct= H['Active?']    || 0;

  if (!colE && !colP) return 1; // no way to match contact → first visit

  // Works with or without lastDataRow_ helper
  const last = (typeof lastDataRow_ === 'function')
    ? lastDataRow_(SHT.MASTER, LASTROW_SENTINELS)
    : s.getLastRow();

  const rows = (last < 2) ? [] : s.getRange(2, 1, last - 1, s.getLastColumn()).getValues();

  const isSameContact = (r) => {
    const e = colE ? String(r[colE-1]||'').toLowerCase() : '';
    const p = colP ? String(r[colP-1]||'')               : '';
    return (!!emailLower && e === emailLower) || (!!phoneNorm && p === phoneNorm);
  };

  const isCountable = (r) => {
    const status = colSta ? String(r[colSta-1]||'') : '';
    // Prefer explicit Active? column when present; otherwise fall back to Status
    const active = colAct ? /^yes$/i.test(String(r[colAct-1]||'')) : /scheduled|rescheduled/i.test(status);
    const completed = /completed/i.test(status);
    // Count only attended or still scheduled to attend
    return completed || active;
  };

  return rows.filter(r => isSameContact(r) && isCountable(r)).length + 1;
}


/***** Template selection *****/
function intakeTemplateIdForBrand_(brand){
const SP = PropertiesService.getScriptProperties();
const vvs = SP.getProperty('INTAKE_TEMPLATE_ID_VVS') || '';
const hp  = SP.getProperty('INTAKE_TEMPLATE_ID_HPUSA') || '';
const any = SP.getProperty('INTAKE_TEMPLATE_ID') || '';
if (brand === 'VVS' && vvs) return vvs;
if (brand === 'HPUSA' && hp) return hp;
return any;
}

/***** Build data for placeholders/block *****/
function buildIntakeData_(rowIdx){
const m = headers_(SHT.MASTER);
const s = SH(SHT.MASTER);
function val(h){ return (m[h]? s.getRange(rowIdx, m[h]).getValue() : '') || ''; }
const tz = CFG.TZ;
const iso = val('ApptDateTime (ISO)');
const date = iso ? new Date(iso) : null;
const apptDate = date ? Utilities.formatDate(date, tz, 'EEE, MMM d, yyyy') : (val('Visit Date')||'');
const apptTime = date ? Utilities.formatDate(date, tz, 'h:mm a') : (val('Visit Time')||'');
const apptDT   = date ? Utilities.formatDate(date, tz, 'EEE, MMM d, yyyy h:mm a z') : '';

const data = {
Brand:              val('Brand'),
Company:            val('Company') || val('Company (normalized)'),
CustomerName:       val('Customer Name'),
FirstName:          val('First Name'),
LastName:           val('Last Name'),
Phone:              val('Phone') || val('PhoneNorm'),
Email:              val('Email') || val('EmailLower'),
ApptDate:           apptDate,
ApptTime:           apptTime,
ApptDateTime:       apptDT,
Location:           val('Location'),
DiamondType:        val('Diamond Type'),
StyleNotes:         val('Style Notes'),
BudgetRange:        val('Budget Range'),
BudgetMin:          val('Budget Min'),
BudgetMax:          val('Budget Max'),
Source:             val('Source (normalized)') || val('Source'),
VisitNumber:        val('Visit #'),
ApptId:             val('APPT_ID'),
CalendlyEventUID:   val('CalendlyEventUID'),
FolderURL:          val('Client Folder'),
ProspectFolderURL:  (function(){
                      const pfId = val('ProspectFolderID');
                      try { return pfId ? DriveApp.getFolderById(pfId).getUrl() : ''; } catch(_){ return ''; }
                    })(),
Timestamp:          Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss'), // <- comma was missing here
RescheduledFromUID: val('RescheduledFromUID'),
RescheduledToUID:   val('RescheduledToUID'),
CanceledAt:         (function(){
                      const ca = val('CanceledAt');
                      try { return ca ? Utilities.formatDate(new Date(ca), CFG.TZ, 'yyyy-MM-dd HH:mm') : ''; }
                      catch(_){ return ca || ''; }
                    })()
};
return data;
}






/***** Fill {{Placeholders}} throughout the doc (body/header/footer) *****/
function fillIntakeDocPlaceholders_(docId, data){
const doc = DocumentApp.openById(docId);
const body = doc.getBody();
const header = doc.getHeader();
const footer = doc.getFooter();
const escape = s => s.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&');
Object.keys(data).forEach(k=>{
const pat = '\\{\\{\\s*' + escape(k) + '\\s*\\}\\}'; // matches {{Key}} with optional spaces
const val = data[k] == null ? '' : String(data[k]);
body.replaceText(pat, val);
if (header) header.replaceText(pat, val);
if (footer) footer.replaceText(pat, val);
});
doc.saveAndClose();
}

/***** Rebuild the Autofill Block at the bottom (no markers) *****/
function upsertAutofillBlock_(docId, data){
const doc = DocumentApp.openById(docId);
const body = doc.getBody();

// 1) Remove a previous autofill table if present (match by first-column "signature")
const signature = ['Appointment Date','Appointment Time','Customer','Diamond Type'];
for (let i = body.getNumChildren() - 1; i >= 0; i--){
const el = body.getChild(i);
if (el.getType() !== DocumentApp.ElementType.TABLE) continue;
const tbl = el.asTable();
if (tbl.getNumRows() < signature.length) continue;


let looksLikeAuto = true;
for (let r = 0; r < signature.length; r++){
  const cellText = (tbl.getCell(r, 0).getText() || '').trim();
  if (cellText !== signature[r]) { looksLikeAuto = false; break; }
}
if (looksLikeAuto){
  body.removeChild(tbl); // remove previous autofill block
  break;
}
}

// 2) Build the new rows
const rows = [];
const pushRow = (label, value) => {
const v = value == null ? '' : String(value);
if (v !== '') rows.push([label, v]);
};

pushRow('Appointment Date', data.ApptDate);
pushRow('Appointment Time', data.ApptTime);
pushRow('Customer',         data.CustomerName);
pushRow('Diamond Type',     data.DiamondType);
pushRow('Budget',           data.BudgetRange || ((data.BudgetMin||data.BudgetMax) ? `$${data.BudgetMin||''}–$${data.BudgetMax||''}` : ''));
pushRow('Location',         data.Location);
pushRow('Source',           data.Source);
pushRow('Visit #',          data.VisitNumber);
pushRow('Appt ID',          data.ApptId);
pushRow('Email',            data.Email);
pushRow('Phone',            data.Phone);

// Optional “Reschedule” section (only if present)
if (data.RescheduledFromUID || data.RescheduledToUID || data.CanceledAt){
rows.push(['', '']); // spacer
pushRow('Rescheduled From', data.RescheduledFromUID);
pushRow('Rescheduled To',   data.RescheduledToUID);
pushRow('Canceled At',      data.CanceledAt);
}

// 3) Append the new block at the very end (bottom of document)
if (rows.length){
const table = body.appendTable(rows);
table.setBorderWidth(0);
for (let r = 0; r < table.getNumRows(); r++){
  table.getRow(r).getCell(0).editAsText().setBold(true);
}
}

doc.saveAndClose();
}


/***** Create (or update) Intake Doc for a Master row *****/
function ensureAndFillIntakeDocForRow_(rowIdx){
const s = SH(SHT.MASTER), m = headers_(SHT.MASTER);
const brand = (m['Brand']? s.getRange(rowIdx, m['Brand']).getValue() : '') || '';
const apptId = (m['APPT_ID']? s.getRange(rowIdx, m['APPT_ID']).getValue() : '') || '';
let intakeUrl = m['IntakeDocURL']? s.getRange(rowIdx, m['IntakeDocURL']).getValue() : '';

// If there’s no Intake yet, clone the brand-specific template into the Prospect folder
if (!intakeUrl){
const pfId = m['ProspectFolderID']? s.getRange(rowIdx, m['ProspectFolderID']).getValue() : '';
if (!pfId) return; // prospect folder not created yet
const tplId = intakeTemplateIdForBrand_(brand);
if (!tplId) return;

const file = DriveApp.getFileById(tplId);
const dest = DriveApp.getFolderById(pfId);
const copy = file.makeCopy(`${brand} – ${apptId} – Intake`, dest);
intakeUrl = copy.getUrl();
setCell_(SHT.MASTER, rowIdx, 'IntakeDocURL', intakeUrl);
}

// Fill placeholders + (re)build Autofill Block every time
const id = intakeUrl.split('/d/')[1]?.split('/')[0] || '';
if (!id) return;
const data = buildIntakeData_(rowIdx);
fillIntakeDocPlaceholders_(id, data);
upsertAutofillBlock_(id, data);
}

/***** Create (or update) Quotation Sheet for a Master row *****/
function ensureAndFillChecklistDocForRow_(rowIdx){
const s = SH(SHT.MASTER), m = headers_(SHT.MASTER);
const brand  = (m['Brand']? s.getRange(rowIdx, m['Brand']).getValue() : '') || '';
const apptId = (m['APPT_ID']? s.getRange(rowIdx, m['APPT_ID']).getValue() : '') || '';
let url      = m['Checklist URL']? s.getRange(rowIdx, m['Checklist URL']).getValue() : '';

const tplId = checklistTemplateIdForBrand_(brand);
if (!tplId) return;             // no template configured → skip

// Create once (in Prospect folder)
if (!url){
const pfId = m['ProspectFolderID']? s.getRange(rowIdx, m['ProspectFolderID']).getValue() : '';
if (!pfId) return;
const file = DriveApp.getFileById(tplId);
const dest = DriveApp.getFolderById(pfId);
const copy = file.makeCopy(`${brand} – ${apptId} – Checklist`, dest);
url = copy.getUrl();
setCell_(SHT.MASTER, rowIdx, 'Checklist URL', url);
}

// Fill placeholders every time
const docId = idFromUrl_(url);
if (docId){
const data = buildIntakeData_(rowIdx);   // universal dataset you already have
fillIntakeDocPlaceholders_(docId, data); // reuse the Doc filler
}
}



/***** Create (or update) Quotation Sheet for a Master row *****/
function ensureAndFillQuotationForRow_(rowIdx){
const s = SH(SHT.MASTER), m = headers_(SHT.MASTER);
const brand  = (m['Brand']? s.getRange(rowIdx, m['Brand']).getValue() : '') || '';
const apptId = (m['APPT_ID']? s.getRange(rowIdx, m['APPT_ID']).getValue() : '') || '';
let url      = m['Quotation URL']? s.getRange(rowIdx, m['Quotation URL']).getValue() : '';

const tplId = quotationTemplateIdForBrand_(brand);
if (!tplId) return;             // no template configured → skip

// Create once (in Prospect folder)
if (!url){
const pfId = m['ProspectFolderID']? s.getRange(rowIdx, m['ProspectFolderID']).getValue() : '';
if (!pfId) return;
const file = DriveApp.getFileById(tplId);
const dest = DriveApp.getFolderById(pfId);
const copy = file.makeCopy(`${brand} – ${apptId} – Quotation`, dest);
url = copy.getUrl();
setCell_(SHT.MASTER, rowIdx, 'Quotation URL', url);
}

// Fill placeholders every time
const ssId = idFromUrl_(url);
if (ssId){
const data = buildIntakeData_(rowIdx);   // universal dataset you already have
fillSheetPlaceholders_(ssId, data);
}
}


/** ---------- ALIASES & SAFE CELL READERS ---------- **/
const COL_ALIASES = {
EmailLower: ['EmailLower','Email'],
PhoneNorm:  ['PhoneNorm','Phone'],
IntakeLink: ['IntakeDocURL','Intake URL','IntakeDoc Url','Intake Doc URL'],
ChecklistLink: ['Checklist URL','ChecklistURL','Checklist Link'],
QuotationLink: ['Quotation URL','QuotationURL','Quotation Link'],
ApptIso: ['ApptDateTime (ISO)','ApptDateTime(ISO)','ApptDateTime'],
Timestamp: ['Timestamp','Created At','CreatedAt']
};

// ---- Last-data-row sentinels (tune as needed) ----
const LASTROW_SENTINELS = ['APPT_ID','Customer Name','EmailLower','Timestamp'];

/** Visit-type → should we stamp "Appointment" stage? */
function _isConsultVisit_(vtRaw){
  const t = String(vtRaw || '').trim().toLowerCase();
  return t === 'appointment' || t === 'diamond viewing';
}

/**
 * Stamp Sales Stage = "Appointment" if Visit Type is Appointment/Diamond Viewing.
 * vtypeFromForm is preferred (since the row's Visit Type might not be written yet).
 */
function stampSalesStageIfConsult_(row, vtypeFromForm){
  const vt = (vtypeFromForm != null && vtypeFromForm !== '')
    ? String(vtypeFromForm).trim()
    : String(getCell_(SHT.MASTER, row, 'Visit Type') || '').trim();

  if (_isConsultVisit_(vt)) {
    setCell_(SHT.MASTER, row, 'Sales Stage', 'Appointment');
  }
}

/**
 * Return the last row that actually contains data in at least one of the
 * sentinel columns (ignores formulas that evaluate to "").
 * Always ≥ 1 (header row). If no data rows, returns 1.
 */
function lastDataRow_(sheetName, sentinels){
  const s = SH(sheetName), H = headers_(sheetName);
  const last = s.getLastRow();                        // may include formula-only rows
  if (last < 2) return 1;

  const cols = (sentinels||[]).map(name => H[name]).filter(Boolean);
  if (!cols.length) return last;                      // fallback when no sentinel columns exist

  let best = 1;
  for (let c of cols){
    const n = Math.max(0, last - 1);                  // data rows count
    if (n === 0) continue;
    const vals = s.getRange(2, c, n, 1).getValues();  // read 1 column
    for (let i = vals.length - 1; i >= 0; i--){
      const v = vals[i][0];
      if (v !== '' && String(v).trim() !== '') {      // treat "" and whitespace as empty
        best = Math.max(best, i + 2);                 // +2 => sheet row index
        break;                                        // found last in this column
      }
    }
  }
  return best;
}

/** Next available data row based on sentinels (min is row 2). */
function nextDataRow_(sheetName, sentinels){
  return Math.max(2, lastDataRow_(sheetName, sentinels) + 1);
}


function _firstHeaderIndex_(H, names){
for (let i=0;i<names.length;i++){
 const n = names[i];
 if (H[n]) return H[n];
}
return 0;
}








function _getByAliases_(sheet, row, H, names){
const c = _firstHeaderIndex_(H, names);
return c ? String(sheet.getRange(row, c).getValue() || '') : '';
}








/** ---------- CURRENT ROW to OBJ (use fallbacks) ---------- **/
function _currentRowToObj_(rowIdx){
const s = SH(SHT.MASTER), H = headers_(SHT.MASTER);
// prefer normalized; fallback to raw
const email = _getByAliases_(s, rowIdx, H, COL_ALIASES.EmailLower).trim().toLowerCase();
const phone = _getByAliases_(s, rowIdx, H, COL_ALIASES.PhoneNorm).trim();
const uid   = (H['CalendlyEventUID'] ? String(s.getRange(rowIdx, H['CalendlyEventUID']).getValue()||'') : '');
return { EmailLower: email, PhoneNorm: phone, CalendlyEventUID: uid };
}

/**
 * Find the most recent prior row for the same person (by EmailLower/PhoneNorm),
 * skipping the current row (curRowIdx) and any row with the same CalendlyEventUID.
 * Returns: { rowIndex, IntakeDocURL, ChecklistURL, QuotationURL } or null.
 */
function _findMostRecentPriorRow(cur, curRowIdx){
  const s = SH(SHT.MASTER), H = headers_(SHT.MASTER);
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  if (last < 2) return null;

  // Resolve column indexes (using your alias map where applicable)
  const cEmail = _firstHeaderIndex_(H, COL_ALIASES.EmailLower);
  const cPhone = _firstHeaderIndex_(H, COL_ALIASES.PhoneNorm);
  const cUID   = H['CalendlyEventUID'] || 0;
  const cISO   = _firstHeaderIndex_(H, COL_ALIASES.ApptIso);
  const cTS    = _firstHeaderIndex_(H, COL_ALIASES.Timestamp);
  const cInt   = _firstHeaderIndex_(H, COL_ALIASES.IntakeLink);
  const cChk   = _firstHeaderIndex_(H, COL_ALIASES.ChecklistLink);
  const cQuo   = _firstHeaderIndex_(H, COL_ALIASES.QuotationLink);

  const n = last - 1; // data rows (2..last)

  // Helper to pull one column as an array (or blanks if column missing)
  const pull = (col) => col
    ? s.getRange(2, col, n, 1).getValues().map(a => String(a[0] || ''))
    : Array(n).fill('');

  // Batch-read all needed columns once
  const emailVec   = pull(cEmail).map(v => v.trim().toLowerCase());
  const phoneVec   = pull(cPhone).map(v => v.trim());
  const uidVec     = pull(cUID).map(v => v.trim());
  const isoVec     = pull(cISO);
  const tsVec      = pull(cTS);
  const intakeVec  = pull(cInt);
  const checkVec   = pull(cChk);
  const quoteVec   = pull(cQuo);

  // Inputs we match against
  const wantEmail  = String(cur && cur.EmailLower || '').trim().toLowerCase();
  const wantPhone  = String(cur && cur.PhoneNorm  || '').trim();
  const curUID     = String(cur && cur.CalendlyEventUID || '').trim();
  const selfIdx    = Number(curRowIdx || 0); // 0 means "unknown"

  // Scan from bottom (most recent first)
  for (let i = n - 1; i >= 0; i--) {
    const r = i + 2;  // actual sheet row index

    // Hard self‑skip when we know the current row
    if (selfIdx && r === selfIdx) continue;

    // Skip exact UID match (same event)
    if (cUID && curUID && uidVec[i] && uidVec[i] === curUID) continue;

    // Same person?
    const same = (wantEmail && emailVec[i] === wantEmail) ||
                 (wantPhone && phoneVec[i] === wantPhone);
    if (!same) continue;

    // Optional lookback window
    if (RFLAGS && RFLAGS.PRIOR_LOOKBACK_DAYS > 0) {
      let t = 0;
      if (cISO) { const d = new Date(String(isoVec[i] || '')); if (!isNaN(d)) t = d.getTime(); }
      if (!t && cTS) { const d2 = new Date(String(tsVec[i] || '')); if (!isNaN(d2)) t = d2.getTime(); }
      if (t && (Date.now() - t) > RFLAGS.PRIOR_LOOKBACK_DAYS * 86400000) continue;
    }

    // Has any artifact? If yes, this is a valid prior row
    const intake = intakeVec[i] || '';
    const check  = checkVec[i]  || '';
    const quote  = quoteVec[i]  || '';
    if (intake || check || quote) {
      return { rowIndex: r, IntakeDocURL: intake, ChecklistURL: check, QuotationURL: quote };
    }
  }

  return null;
}


function ensureArtifactsForRow_(row){
const brand   = getCell_(SHT.MASTER,row,'Brand');
const apptId  = getCell_(SHT.MASTER,row,'APPT_ID');
if (!brand || !apptId) return;

// 1) Client folder (create if missing, backfill ID/link)
let clientFolder, cfId = getCell_(SHT.MASTER,row,'ClientFolderID');
if (cfId) { try { clientFolder = DriveApp.getFolderById(cfId); } catch(_){ clientFolder = null; } }
if (!clientFolder){
 const phoneNorm = getCell_(SHT.MASTER,row,'PhoneNorm');
 const emailLower= getCell_(SHT.MASTER,row,'EmailLower');
 const custName  = getCell_(SHT.MASTER,row,'Customer Name');
 clientFolder = ensureClientFolder_(brand, custName, phoneNorm, emailLower);
 setCell_(SHT.MASTER,row,'ClientFolderID', clientFolder.getId());
}
if (!getCell_(SHT.MASTER,row,'Client Folder')) {
 setCell_(SHT.MASTER,row,'Client Folder', clientFolder.getUrl());
}

// 2) Prospect folder (create if missing)
let prospectFolder, pfId = getCell_(SHT.MASTER,row,'ProspectFolderID');
if (pfId) { try { prospectFolder = DriveApp.getFolderById(pfId); } catch(_){ prospectFolder = null; } }
if (!prospectFolder){
prospectFolder = ensureProspectFolder_(clientFolder, apptId);
setCell_(SHT.MASTER,row,'ProspectFolderID', prospectFolder.getId());
}

// After ensuring Prospect folder:
bootstrapApFolderForRow_(row);   // use the robust writer that sets "RootAppt Folder ID"

// 2b) Reuse Intake / Checklist / Quotation links from the most recent prior visit (if any)
if (RFLAGS.REUSE_ARTIFACTS_FROM_PRIOR) {
try {
 const curObj = _currentRowToObj_(row);
 const prior  = _findMostRecentPriorRow(curObj, row);  // <— pass row here
 if (prior) {
   let reused = [];
   if (!getCell_(SHT.MASTER, row, 'IntakeDocURL') && prior.IntakeDocURL) {
     setCell_(SHT.MASTER, row, 'IntakeDocURL', prior.IntakeDocURL); reused.push('Intake');
   }
   if (!getCell_(SHT.MASTER, row, 'Checklist URL') && prior.ChecklistURL) {
     setCell_(SHT.MASTER, row, 'Checklist URL', prior.ChecklistURL); reused.push('Checklist');
   }
   if (!getCell_(SHT.MASTER, row, 'Quotation URL') && prior.QuotationURL) {
     setCell_(SHT.MASTER, row, 'Quotation URL', prior.QuotationURL); reused.push('Quotation');
   }
   if (reused.length) log_('REUSED_FROM_PRIOR', { row, priorRow: prior.rowIndex, reused });
 }
} catch (e) {
 err_('reuseArtifacts', e.message, { row });
}
}


// 3) Intake doc, Checklist, Quotation (brand-specific template) + autofill
ensureAndFillIntakeDocForRow_(row);
ensureAndFillChecklistDocForRow_(row);
ensureAndFillQuotationForRow_(row);

}


function idFromUrl_(url){
if(!url) return '';
const m = String(url).match(/\/d\/([a-zA-Z0-9\-_]+)/);
return m ? m[1] : '';
}


function checklistTemplateIdForBrand_(brand){
const SP = PropertiesService.getScriptProperties();
const vvs = SP.getProperty('CHECKLIST_TEMPLATE_ID_VVS') || '';
const hp  = SP.getProperty('CHECKLIST_TEMPLATE_ID_HPUSA') || '';
return brand === 'VVS' ? vvs : brand === 'HPUSA' ? hp : '';
}


function quotationTemplateIdForBrand_(brand){
const SP = PropertiesService.getScriptProperties();
const vvs = SP.getProperty('QUOTATION_TEMPLATE_ID_VVS') || '';
const hp  = SP.getProperty('QUOTATION_TEMPLATE_ID_HPUSA') || '';
return brand === 'VVS' ? vvs : brand === 'HPUSA' ? hp : '';
}
















function fillSheetPlaceholders_(spreadsheetId, data){
const ss = SpreadsheetApp.openById(spreadsheetId);
const keys = Object.keys(data);
ss.getSheets().forEach(sh => {
keys.forEach(k => {
  const val = data[k] == null ? '' : String(data[k]);
  const pat = '{{' + k + '}}';
  sh.createTextFinder(pat).useRegularExpression(false).replaceAllWith(val);
});
});
SpreadsheetApp.flush();
}


/********** RESOLVER ENTRY (FORM) **********/
function onFormSubmit(e){
  __mark('onFormSubmit: START');

try{
const nv = e && e.namedValues ? e.namedValues : {};

// 1) Read the Form answers (match your exact Form question titles)
const submittedAt = (nv['Timestamp']||[''])[0];     // auto-added by Google Forms

// Convert the Form's Timestamp to ISO in our TZ
const bookedAtISO = Utilities.formatDate(
  new Date(submittedAt || new Date()),
  CFG.TZ || 'America/Los_Angeles',
  "yyyy-MM-dd'T'HH:mm:ssXXX"
);

const company     = (nv['Company']||[''])[0];
const name        = (nv['Customer Name']||[''])[0];
const phone       = (nv['Phone']||[''])[0];
const email       = (nv['Email']||[''])[0];
const vtype       = (nv['Visit Type']||[''])[0];
const vdate       = (nv['Visit Date']||[''])[0];    // e.g., 8/27/2025
const vtime       = (nv['Visit Time']||[''])[0];    // e.g., 11:00:00 AM
const location    = (nv['Location']||[''])[0];
const budgetRaw   = (nv['Budget Range']||[''])[0];
const sourceRaw   = (nv['Source']||[''])[0];
const notes       = (nv['Style Notes']||[''])[0];
const calUID      = nvGet(nv, 'Admin: Calendly Event UID');
const diamondTypeQ = (nv['Diamond Type']||[''])[0];
const diamondTypeNorm = (() => {
  const s = (diamondTypeQ||'').toLowerCase();
  const hasLab = /lab/.test(s), hasNat = /natural/.test(s);
  return hasLab && hasNat ? 'Both' : hasLab ? 'Lab' : hasNat ? 'Natural' : '';
})();
 // 2) Canonical derivations
const brand       = brandFromCompany_(company);
const emailLower  = normEmail_(email);
const phoneNorm   = normPhone_(phone);

// Build ISO from Form date/time (keep your TZ)
const apptIso = (vdate || vtime)
  ? Utilities.formatDate(new Date(`${vdate} ${vtime}`), CFG.TZ, "yyyy-MM-dd'T'HH:mm:ssXXX")
  : '';

// Parse budget min/max only when a single range chosen
const {min,max} = parseBudget_(budgetRaw);

// Split name
const parts = splitName_(name);

Logger.log(JSON.stringify({
  company,
  name,
  emailLower,
  phoneNorm,
  vtype,
  vdate,
  vtime,
  location,
  calUID
}, null, 2));
__mark('parsed+normalized fields');

  // 3) Decide: reschedule vs new booking (never reuse the canceled row)
  let row = 0;
  let createdNow = false;

  // Prefer direct UID match (rare on true reschedule, but harmless)
  if (calUID) row = findMasterRowByUID_(calUID);

  // POP any pending canceled UID placed by the cancel webhook (most robust)
  let looksLikeReschedule = false, oldRow = 0, oldUID = '';
  const pendingOldUID = _popPendingCancelUID_(brand, vtype, emailLower, phoneNorm);
  if (!row && pendingOldUID){
    const r = findMasterRowByUID_(pendingOldUID);
    if (r){
      looksLikeReschedule = true;
      oldRow = r;
      oldUID = pendingOldUID;
    }
  }

  // If no cached UID, fall back to recent Canceled row by CanceledAt (your existing heuristic)
  if (!row && !looksLikeReschedule){
    const candRow = findRecentCanceledByContactAt_(emailLower, phoneNorm, /*minutes*/ 240);
    if (candRow){
      const normL = s => (s||'').toString().trim().toLowerCase();
      const normU = s => (s||'').toString().trim().toUpperCase();
      const sameType  = normL(getCell_(SHT.MASTER,candRow,'Visit Type')) === normL(vtype);
      const sameBrand = normU(getCell_(SHT.MASTER,candRow,'Brand'))      === normU(brand);
      if (sameType && sameBrand){
        looksLikeReschedule = true;
        oldRow = candRow;
        oldUID = getCell_(SHT.MASTER,candRow,'CalendlyEventUID') || '';
      }
    }
  }


  // If fallback found a row, only reuse it when it doesn’t already belong to a different UID
  if (row) {
    const existingUID = rf(row, 'CalendlyEventUID') || '';
    if (existingUID && calUID && existingUID !== calUID) {
      // likely a reschedule where cancel hasn’t landed yet → force a new row
      row = 0;
    }
  }

// 3.4 Create only if still not found
__mark('reschedule detection done; looksLikeReschedule=' + looksLikeReschedule + ', oldRow=' + oldRow + ', preRow=' + row);

if (!row){
  row = appendObj_(SHT.MASTER, {'APPT_ID': nextApptId_(apptIso)});
  setCell_(SHT.MASTER, row, 'Visit #', countVisits_(emailLower, phoneNorm));
  // stamp using the form value (robust even before the row’s Visit Type is written)
  stampSalesStageIfConsult_(row, vtype);
  createdNow = true;
}


// Write-once (only set if empty)
setOnce_(SHT.MASTER, row, 'Booked At (ISO)', bookedAtISO);


// >>> RootApptID + Active? (new/current row) — START
__mark('before ensureRootAndActiveForNewRow'); 

(function ensureRootAndActiveForNewRow(){
 try {
   const newAppt = getCell_(SHT.MASTER, row, 'APPT_ID') || '';


   if (looksLikeReschedule && oldRow) {
     // inherit root from old chain
     const oldRoot = getCell_(SHT.MASTER, oldRow, 'RootApptID') || '';
     const oldAppt = getCell_(SHT.MASTER, oldRow, 'APPT_ID')   || '';
     const root    = oldRoot || oldAppt || '';
     if (root) {
       setOnce_(SHT.MASTER, row,    'RootApptID', root);
       setOnce_(SHT.MASTER, oldRow, 'RootApptID', root); // backfill old if blank
     }
   } else {
     // NEW: same-person chaining (email/phone) even when NOT a reschedule
     try {
       const cur = {
         EmailLower: String(emailLower || '').trim().toLowerCase(),
         PhoneNorm:  String(phoneNorm  || '').trim(),
         CalendlyEventUID: String(calUID || '')
       };
       // Reuse helper; it ignores current row and skips UID matches
      __mark('ensureRoot: BEFORE prior-scan');
      const prior = _findMostRecentPriorRow(cur, row);
      __mark('ensureRoot: AFTER prior-scan ' + (prior && prior.rowIndex ? ('hit row ' + prior.rowIndex) : '(none)'));

       if (prior && prior.rowIndex) {
         const prevRoot = getCell_(SHT.MASTER, prior.rowIndex, 'RootApptID')
                       || getCell_(SHT.MASTER, prior.rowIndex, 'APPT_ID')
                       || '';
         if (prevRoot) setOnce_(SHT.MASTER, row, 'RootApptID', prevRoot);
       }
     } catch(_){}
   }
 } catch(e) {
   Logger.log('ensureRootAndActiveForNewRow error: ' + e.message);
 }
})();
__mark('after ensureRootAndActiveForNewRow'); 

 // <<< RootApptID + Active? (new/current row) — END

// 3.6 If this is a reschedule, mark + link the OLD row (idempotent, no self-link)
if (looksLikeReschedule && oldRow) {
  withScriptLock_(() => {
    // Only link when we have two distinct UIDs
    if (oldUID && calUID && oldUID !== calUID) {
      const already = (getCell_(SHT.MASTER, oldRow, 'RescheduledToUID') || '');
      if (!already) {
        setCell_(SHT.MASTER, oldRow, 'RescheduledToUID', calUID);   // old → new
      }
      const newFrom = (getCell_(SHT.MASTER, row, 'RescheduledFromUID') || '');
      if (!newFrom) {
        setCell_(SHT.MASTER, row, 'RescheduledFromUID', oldUID);    // new ← old
      }
    }

    // Status on the old row (don’t downgrade if it’s already Rescheduled)
    const curSta = getCell_(SHT.MASTER, oldRow, 'Status') || '';
    if (!/rescheduled/i.test(curSta)) {
      setCell_(SHT.MASTER, oldRow, 'Status', 'Rescheduled');
    }

    // Stamp a cancel time if it’s empty (for analytics consistency)
    if (!(getCell_(SHT.MASTER, oldRow, 'CanceledAt'))) {
      try { setCell_(SHT.MASTER, oldRow, 'CanceledAt', new Date()); } catch(_){}
    }


        // >>> Active? (old row) — START
       try {
         setCell_(SHT.MASTER, oldRow, 'Active?', 'No');
       } catch (_noActiveCol) { /* column may not exist yet; safe to ignore */ }
       // <<< Active? (old row) — END


    const prev = getCell_(SHT.MASTER, oldRow, 'Automation Notes') || '';
    setCell_(SHT.MASTER, oldRow, 'Automation Notes',
      (prev? prev+'\n':'') + `Rescheduled → ${calUID} @ ${new Date().toISOString()}`);
  });
}


// 4) Write authoritative + narrative + convenience fields
const updates = {
  // status/keys
  'Status': 'Scheduled',
  'Active?': 'Yes',  

  // brand/company
  'Brand': brand || getCell_(SHT.MASTER,row,'Brand') || '',
  'Company': company || getCell_(SHT.MASTER,row,'Company') || '',
  'Company (normalized)': brand || getCell_(SHT.MASTER,row,'Company (normalized)') || '',


  // client
  'Customer Name': name || getCell_(SHT.MASTER,row,'Customer Name') || '',
  'First Name': parts.first || getCell_(SHT.MASTER,row,'First Name') || '',
  'Last Name': parts.last || getCell_(SHT.MASTER,row,'Last Name') || '',
  'Phone': phone || getCell_(SHT.MASTER,row,'Phone') || '',
  'PhoneNorm': phoneNorm || getCell_(SHT.MASTER,row,'PhoneNorm') || '',
  'Email': email || getCell_(SHT.MASTER,row,'Email') || '',
  'EmailLower': emailLower || getCell_(SHT.MASTER,row,'EmailLower') || '',


  // scheduling
  'Visit Type': vtype || getCell_(SHT.MASTER,row,'Visit Type') || '',
  'Visit Date': vdate || getCell_(SHT.MASTER,row,'Visit Date') || '',
  'Visit Time': vtime || getCell_(SHT.MASTER,row,'Visit Time') || '',
  'ApptDateTime (ISO)': apptIso || getCell_(SHT.MASTER,row,'ApptDateTime (ISO)') || '',
  'Timezone': CFG.TZ,
  'Duration (min)': getCell_(SHT.MASTER,row,'Duration (min)') || DEFAULT_DURATION_MIN,
  'Location': locToEnum_(location) || getCell_(SHT.MASTER,row,'Location') || '',


  // calendly
  'CalendlyEventUID': calUID || getCell_(SHT.MASTER,row,'CalendlyEventUID') || '',
  'Diamond Type':     diamondTypeNorm || getCell_(SHT.MASTER,row,'Diamond Type') || '',


  // sales signals
  'Budget Range': budgetRaw || getCell_(SHT.MASTER,row,'Budget Range') || '',
  'Budget Min': min || getCell_(SHT.MASTER,row,'Budget Min') || '',
  'Budget Max': max || getCell_(SHT.MASTER,row,'Budget Max') || '',
  'Source': sourceRaw || getCell_(SHT.MASTER,row,'Source') || '',
  'Source (normalized)': normSource_(sourceRaw) || getCell_(SHT.MASTER,row,'Source (normalized)') || '',
  'Style Notes': notes || getCell_(SHT.MASTER,row,'Style Notes') || '',

  // meta
  'Timestamp': submittedAt || getCell_(SHT.MASTER,row,'Timestamp') || ''
};

// Optional: infer Diamond Type from notes
if (!getCell_(SHT.MASTER,row,'Diamond Type')) {
  const m = /preferred diamond type:\s*([^\n]+)/i.exec(notes || '');
  if (m && m[1]) updates['Diamond Type'] = m[1].trim();
}

// Commit all updates
Object.keys(updates).forEach(k=> setCell_(SHT.MASTER,row,k,updates[k]));
// Ensure final Sales Stage is correct based on Visit Type
stampSalesStageIfConsult_(row, vtype);

// Ensure Visit # is set even when we merged into a previously Canceled row
if (!getCell_(SHT.MASTER, row, 'Visit #')) {
  setCell_(SHT.MASTER, row, 'Visit #', countVisits_(emailLower, phoneNorm));
}

// After: setOnce_(SHT.MASTER, row, 'Booked At (ISO)', bookedAtISO);
  // Before: ensureArtifactsForRow_(row);

  (function ensureRootAfterWrites(){
    try {
      const newAppt = getCell_(SHT.MASTER, row, 'APPT_ID') || '';
      const curRoot = getCell_(SHT.MASTER, row, 'RootApptID') || '';

      // Build a search key from the just-written values (don’t read half-baked cells)
      const curObj = { EmailLower: emailLower, PhoneNorm: phoneNorm, CalendlyEventUID: calUID };
      const prior  = _findMostRecentPriorRow(curObj, row); // pass current row to skip self


      let desired = '';
      if (prior) {
        desired = getCell_(SHT.MASTER, prior.rowIndex, 'RootApptID') ||
                  getCell_(SHT.MASTER, prior.rowIndex, 'APPT_ID') || '';
      }
      if (!desired) desired = newAppt;

      // Allow overwrite when blank OR currently self-root (new chain)
      if (!curRoot || curRoot === newAppt) {
        if (desired) setCell_(SHT.MASTER, row, 'RootApptID', desired);
      }
      if (prior && !getCell_(SHT.MASTER, prior.rowIndex, 'RootApptID')) {
        setCell_(SHT.MASTER, prior.rowIndex, 'RootApptID', desired);
      }
    } catch(e){
      err_('ensureRootAfterWrites', e.message, { row });
    }
  })();

// --- DV: enqueue 12-days-before "Propose" nudge (idempotent; Diamond Viewing only) ---
try {
  DV_tryEnqueueOnCreate_({ sh: SH(SHT.MASTER), row: row, dryRun: false });
} catch (e) {
  Logger.log('DV_onCreate skipped: ' + e.message);
}
// --- end DV enqueue ---

// 5) Artifacts (robust)
ensureArtifactsForRow_(row);

// 5b) Ensure the Intake exists (if missing) and (re)fill it every time
ensureAndFillIntakeDocForRow_(row);

// Post notification
try {
  const debug = /true/i.test(PropertiesService.getScriptProperties().getProperty('DEBUG') || 'false');

  Logger.log(`[CHAT] gate check: DEBUG=${debug} createdNow=${createdNow} looksLikeReschedule=${looksLikeReschedule} oldRow=${oldRow} row=${row}`);

  if (!debug) {
    if (looksLikeReschedule && oldRow) {
      Logger.log(`[CHAT] sending RESCHEDULED card oldRow=${oldRow} newRow=${row}`);
      postRescheduledCard_(oldRow, row);
    } else if (createdNow) {
      Logger.log(`[CHAT] sending CREATED card row=${row}`);
      postIntakeCreatedCard_(row);
    } else {
      Logger.log(`[CHAT] not sending: neither reschedule nor createdNow`);
    }
  } else {
    Logger.log(`[CHAT] not sending: DEBUG=true`);
  }
} catch (ex) {
  err_('postNotify_', ex.message, { row, looksLikeReschedule, oldRow });
}


// 6) Notes
const prev = getCell_(SHT.MASTER,row,'Automation Notes') || '';
setCell_(SHT.MASTER,row,'Automation Notes', (prev?prev+'\n':'') + `Form merged @ ${new Date().toISOString()}`);

if (CFG.DEBUG) log_('FORM_MERGED', {row, emailLower, apptIso, brand});
}catch(ex){
err_('onFormSubmit', ex.message, {stack: ex.stack});
throw ex;
}
}

function withScriptLock_(fn){
const lock = LockService.getScriptLock();
if (!lock.tryLock(5000)) { return fn(false); }   // still run, just note no lock
try { return fn(true); } finally { lock.releaseLock(); }
}


/********** UTIL **********/


function ping_(){ return Utilities.formatDate(new Date(), CFG.TZ, "yyyy-MM-dd HH:mm:ss"); }


// Manual backfill: process all existing rows in 02_Form_Inbox (optional)
function backfillFromFormInbox_(){
const s=SH(SHT.FORM_INBOX), last=s.getLastRow();
if (last<2) return;
const m=headers_(SHT.FORM_INBOX);
const rows=s.getRange(2,1,last-1,s.getLastColumn()).getValues();
rows.forEach(r=>{
const nv = {};
Object.keys(m).forEach(h=> nv[h]=[r[m[h]-1]]);
onFormSubmit({namedValues: nv});
});
}


function debug_runResolverOnLastFormRow(){
  __startProfile('debug_runResolverOnLastFormRow');

  // Confirm which helper definitions are active
  Logger.log('_findMostRecentPriorRow.length = ' + _findMostRecentPriorRow.length);
  Logger.log('_currentRowToObj_.length = ' + _currentRowToObj_.length);

  const s = SH(SHT.FORM_INBOX), m = headers_(SHT.FORM_INBOX);
  const r = s.getLastRow(); if (r < 2) { Logger.log('No inbox rows'); return; }

  const vals = s.getRange(r,1,1,s.getLastColumn()).getValues()[0];
  const nv = {};
  Object.keys(m).forEach(h => nv[h] = [ vals[m[h]-1] ]);

  Logger.log('Inbox NV keys: ' + Object.keys(nv).join(', '));
  const __src = _findMostRecentPriorRow.toString();
  Logger.log('active prior-scan has markers? ' + (__src.indexOf('[prior-scan]') >= 0));
  Logger.log(__src.substring(0, 220));  // show the head of the active function


  __mark('calling onFormSubmit');
  onFormSubmit({ namedValues: nv });
  __mark('onFormSubmit returned');
}


// --- Chat helpers (simple card) ---
function chatWebhook_(){
const url = PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_ALL');
if (!url) throw new Error('Missing CHAT_WEBHOOK_ALL script property');
return url;
}

function _redactWebhook_(url){
  if (!url) return '(missing)';
  try {
    const end = url.slice(-12);
    return '…' + end; // show only the tail
  } catch(_) {
    return '(unprintable)';
  }
}

function debug_diagWebhookProperty(){
  const sp = PropertiesService.getScriptProperties();
  const url = sp.getProperty('CHAT_WEBHOOK_ALL');
  const dbg = sp.getProperty('DEBUG');
  Logger.log(`[CHAT] DEBUG prop = ${dbg == null ? '(unset)' : dbg}`);
  Logger.log(`[CHAT] CHAT_WEBHOOK_ALL present? ${!!url}  value=${_redactWebhook_(url)}`);
}

function debug_postPlainTextToChat(text){
  const url = PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_ALL');
  if (!url) throw new Error('CHAT_WEBHOOK_ALL missing');

  const payload = { text: String(text || 'Hello from Apps Script') };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  Logger.log(`[CHAT/TEXT] code=${res.getResponseCode()} body=${res.getContentText()}`);
}


// Build a small Cards V2 intake-created card from a Master row (strings only)
function buildIntakeCreatedCard_(rowIdx){
const s = SpreadsheetApp.getActive().getSheetByName('00_Master Appointments');
const H = s.getRange(1,1,1,s.getLastColumn()).getValues()[0]
         .reduce((m,h,i)=> (m[h]=i+1,m), {});
const S = (v) => v == null ? '' : String(v);           // <— force STRING
function V(h){ return H[h] ? s.getRange(rowIdx, H[h]).getValue() : ''; }


const brand = S((V('Brand') || V('Company') || '')).toUpperCase();
const title = brand === 'VVS' ? 'VVS Appointment Ready'
           : brand === 'HPUSA' ? 'HPUSA Appointment Ready'
           : 'New Appointment Ready';


const tz = PropertiesService.getScriptProperties().getProperty('DEFAULT_TZ') || 'America/Los_Angeles';
const iso = V('ApptDateTime (ISO)');
const dt  = iso ? Utilities.formatDate(new Date(iso), tz, 'EEE, MMM d, yyyy h:mm a z')
              : S((V('Visit Date') || '') + ' ' + (V('Visit Time') || '')).trim();

const customer   = S(V('Customer Name') || (S(V('First Name')) + ' ' + S(V('Last Name'))).trim());
const assigned   = S(V('Assigned Rep') || '(unassigned)');
const vtype      = S(V('Visit Type') || 'Appointment');
const budget     = S(V('Budget Range') || '');
const source     = S(V('Source (normalized)') || V('Source') || '');

const folderUrl    = S(V('Client Folder') || '');
const intakeUrl    = S(V('IntakeDocURL') || '');
const checklistUrl = S(V('Checklist URL') || '');

const widgets = [
{ decoratedText: { topLabel: 'Customer',          text: customer || '(unknown)' } },
{ decoratedText: { topLabel: 'Assigned Rep',      text: assigned } },
{ decoratedText: { topLabel: 'Visit Date & Time', text: dt || '(tbd)' } },
{ decoratedText: { topLabel: 'Visit Type',        text: vtype } },
{ decoratedText: { topLabel: 'Budget',            text: budget } },
{ decoratedText: { topLabel: 'Source',            text: source } },
];

const buttons = [];
if (folderUrl)    buttons.push({ text: 'OPEN FOLDER', onClick: { openLink: { url: folderUrl } } });
if (intakeUrl)    buttons.push({ text: 'INTAKE FORM', onClick: { openLink: { url: intakeUrl } } });
if (checklistUrl) buttons.push({ text: 'CHECKLIST',   onClick: { openLink: { url: checklistUrl } } });
if (buttons.length) widgets.push({ buttonList: { buttons } });

return {
cardsV2: [{
  cardId: 'intake-created',
  card: {
    header: { title },
    sections: [{ widgets }]
  }
}]
};
}


function buildRescheduledCard_(oldRowIdx, newRowIdx){
const s = SpreadsheetApp.getActive().getSheetByName('00_Master Appointments');
const H = s.getRange(1,1,1,s.getLastColumn()).getValues()[0]
         .reduce((m,h,i)=> (m[h]=i+1,m), {});
const S = (v) => v == null ? '' : String(v);
const V = (row, h) => H[h] ? s.getRange(row, H[h]).getValue() : '';

const tz   = PropertiesService.getScriptProperties().getProperty('DEFAULT_TZ') || 'America/Los_Angeles';
const brand= S((V(newRowIdx,'Brand') || V(newRowIdx,'Company') || '')).toUpperCase();
const title= brand==='VVS' ? 'VVS Appointment — Rescheduled'
         : brand==='HPUSA' ? 'HPUSA Appointment — Rescheduled'
         : 'Appointment — Rescheduled';

const oldISO = V(oldRowIdx,'ApptDateTime (ISO)');
const newISO = V(newRowIdx,'ApptDateTime (ISO)');
const oldDT  = oldISO ? Utilities.formatDate(new Date(oldISO), tz, 'EEE, MMM d, yyyy h:mm a z') : '(original time)';
const newDT  = newISO ? Utilities.formatDate(new Date(newISO), tz, 'EEE, MMM d, yyyy h:mm a z') : '(new time)';

const customer = S(V(newRowIdx,'Customer Name') || (S(V(newRowIdx,'First Name'))+' '+S(V(newRowIdx,'Last Name'))).trim());
const folderUrl    = S(V(newRowIdx,'Client Folder') || '');
const intakeUrl    = S(V(newRowIdx,'IntakeDocURL') || '');
const checklistUrl = S(V(newRowIdx,'Checklist URL') || '');

const widgets = [
{ decoratedText: { topLabel: 'Customer', text: customer || '(unknown)' } },
{ decoratedText: { topLabel: 'Old → New', text: oldDT + '  →  ' + newDT } },
];

const buttons = [];
if (folderUrl)    buttons.push({ text:'OPEN FOLDER', onClick:{ openLink:{ url: folderUrl } }});
if (intakeUrl)    buttons.push({ text:'INTAKE FORM', onClick:{ openLink:{ url: intakeUrl } }});
if (checklistUrl) buttons.push({ text:'CHECKLIST',   onClick:{ openLink:{ url: checklistUrl } }});
if (buttons.length) widgets.push({ buttonList:{ buttons } });

return {
cardsV2: [{
  cardId: 'intake-rescheduled',
  card: {
    header: { title },
    sections: [{ widgets }]
  }
}]
};
}

function postRescheduledCard_(oldRowIdx, newRowIdx){
  const payload = buildRescheduledCard_(oldRowIdx, newRowIdx);
  return _postChatPayload_('rescheduled', payload, { oldRowIdx, newRowIdx });
}

// Post the card
function postIntakeCreatedCard_(rowIdx){
  const payload = buildIntakeCreatedCard_(rowIdx);
  return _postChatPayload_('created', payload, { rowIdx });
}

/**
 * Low-level poster with rich logging of request/response.
 * Returns {code, body} or throws on hard failure.
 */
function _postChatPayload_(kind, payload, ctx){
  const sp = PropertiesService.getScriptProperties();
  const url = sp.getProperty('CHAT_WEBHOOK_ALL');
  const debugProp = sp.getProperty('DEBUG') || '(unset)';
  const json = JSON.stringify(payload);
  const sizeB = json.length;

  Logger.log(`[CHAT] kind=${kind} ctx=${JSON.stringify(ctx)} DEBUG=${debugProp}`);
  Logger.log(`[CHAT] webhook set? ${!!url}  url=${_redactWebhook_(url)}  payloadSize=${sizeB}B`);

  if (!url) {
    Logger.log(`[CHAT] ABORT: CHAT_WEBHOOK_ALL missing`);
    err_('chat_post', 'CHAT_WEBHOOK_ALL missing', { kind, ctx, sizeB });
    return { code: 0, body: 'missing webhook' };
  }

  // Show a safe preview of the payload (first ~600 chars)
  Logger.log(`[CHAT] payload head: ${json.substring(0, 600)}${sizeB>600?' …[trunc]':''}`);

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: json,
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    const body = (res.getContentText() || '').substr(0, 1000); // cap in logs
    Logger.log(`[CHAT] response code=${code}`);
    Logger.log(`[CHAT] response body: ${body}`);

    if (code !== 200) {
      err_('chat_post', `Non-200 response (${code})`, { kind, ctx, body });
    }
    return { code, body };
  } catch (e) {
    Logger.log(`[CHAT] EXCEPTION: ${e && e.message}`);
    err_('chat_post', e.message || 'exception', { kind, ctx, stack: e && e.stack });
    throw e;
  }
}

function debug_postCardForLastDataRow(){
  const r = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  if (r < 2) throw new Error('Master sheet has no data rows');
  postIntakeCreatedCard_(r);
}


function diagArtifacts() {
const SP = PropertiesService.getScriptProperties();
const need = {
DEFAULT_TZ: SP.getProperty('DEFAULT_TZ') || '(missing)',
VVS_CLIENTS_ROOT_ID: SP.getProperty('VVS_CLIENTS_ROOT_ID') || '(missing)',
HP_CLIENTS_ROOT_ID:  SP.getProperty('HP_CLIENTS_ROOT_ID')  || '(missing)',
INTAKE_TEMPLATE_ID_VVS: SP.getProperty('INTAKE_TEMPLATE_ID_VVS') || '(missing)',
INTAKE_TEMPLATE_ID_HPUSA: SP.getProperty('INTAKE_TEMPLATE_ID_HPUSA') || '(missing)',
CHECKLIST_TEMPLATE_ID_VVS: SP.getProperty('CHECKLIST_TEMPLATE_ID_VVS') || '(missing)',
CHECKLIST_TEMPLATE_ID_HPUSA: SP.getProperty('CHECKLIST_TEMPLATE_ID_HPUSA') || '(missing)',
QUOTATION_TEMPLATE_ID_VVS: SP.getProperty('QUOTATION_TEMPLATE_ID_VVS') || '(missing)',
QUOTATION_TEMPLATE_ID_HPUSA: SP.getProperty('QUOTATION_TEMPLATE_ID_HPUSA') || '(missing)',
};


const s = SpreadsheetApp.getActive().getSheetByName('00_Master Appointments');
const headers = s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(h=>String(h).trim());
const mustHaveCols = [
'Brand','APPT_ID','ClientFolderID','Client Folder','ProspectFolderID',
'IntakeDocURL','Checklist URL','Quotation URL','EmailLower','PhoneNorm'
];
const missingCols = mustHaveCols.filter(c => !headers.includes(c));

const r = s.getLastRow();
const brand = r>=2 ? s.getRange(r, headers.indexOf('Brand')+1).getValue() : '';
const apptId= r>=2 ? s.getRange(r, headers.indexOf('APPT_ID')+1).getValue() : '';


Logger.log(JSON.stringify({
scriptProperties: need,
missingColumns: missingCols,
lastRowBrand: brand, lastRowApptId: apptId
}, null, 2));
}

function debug_createArtifactsForLastRow() {
const s = SH(SHT.MASTER);
const r = s.getLastRow();
if (r < 2) { Logger.log('No data rows found.'); return; }

// Before state
const H = headers_(SHT.MASTER);
function val(h){ return H[h] ? s.getRange(r, H[h]).getValue() : ''; }
Logger.log('Before:', JSON.stringify({
row: r,
Brand: val('Brand'),
APPT_ID: val('APPT_ID'),
ClientFolderID: val('ClientFolderID'),
ClientFolder: val('Client Folder'),
ProspectFolderID: val('ProspectFolderID'),
IntakeDocURL: val('IntakeDocURL'),
ChecklistURL: val('Checklist URL'),
QuotationURL: val('Quotation URL'),
}, null, 2));

// Force creation/fill
ensureArtifactsForRow_(r);

// After state
Logger.log('After:', JSON.stringify({
row: r,
ClientFolderID: val('ClientFolderID'),
FolderURL: val('Client Folder'),
ProspectFolderID: val('ProspectFolderID'),
IntakeDocURL: val('IntakeDocURL'),
ChecklistURL: val('Checklist URL'),
QuotationURL: val('Quotation URL'),
}, null, 2));
}

function debug_showTZ(){
const tz = PropertiesService.getScriptProperties().getProperty('DEFAULT_TZ');
Logger.log('DEFAULT_TZ = ' + tz);
Logger.log('Now = ' + Utilities.formatDate(new Date(), tz || 'America/Los_Angeles', "yyyy-MM-dd HH:mm:ss z"));
}

function diagResolverMaster_() {
 const ui = SpreadsheetApp.getUi();
 const ss = SpreadsheetApp.getActive();

 // 1) Confirm Resolver’s constant
 const masterName = SHT && SHT.MASTER ? SHT.MASTER : '(SHT.MASTER not set)';
 const sh = ss.getSheetByName(masterName);

 if (!sh) {
   ui.alert('Resolver: Master sheet not found',
     'SHT.MASTER = "' + masterName + '" but no such tab exists.\n' +
     'Available tabs: ' + ss.getSheets().map(s => s.getName()).join(' | '),
     ui.ButtonSet.OK);
   return;
 }


 // 2) Header map (row 1)
 const H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]
           .reduce((m,h,i)=> (h && (m[String(h).trim()]=i+1), m), {});
 const need = ['APPT_ID','RootApptID','Brand','Customer Name','EmailLower','PhoneNorm',
               'CalendlyEventUID','Visit Type','ApptDateTime (ISO)','Folder URL',
               'ProspectFolderID','IntakeDocURL','Checklist URL','Quotation URL'];
 const missing = need.filter(h => !H[h]);


 // 3) Sample the last row (if any) — READ‑ONLY
 const last = sh.getLastRow();
 let sample = {};
 if (last >= 2) {
   const V = (h) => H[h] ? sh.getRange(last, H[h]).getValue() : '';
   sample = {
     rowIndex: last,
     APPT_ID: V('APPT_ID'),
     Brand: V('Brand'),
     EmailLower: V('EmailLower'),
     PhoneNorm: V('PhoneNorm'),
     'ApptDateTime (ISO)': V('ApptDateTime (ISO)'),
     'Client Folder': V('Client Folder'),
     ProspectFolderID: V('ProspectFolderID'),
     IntakeDocURL: V('IntakeDocURL'),
     'Checklist URL': V('Checklist URL'),
     'Quotation URL': V('Quotation URL')
   };
 }


 // 4) Report
 const ok = [
   `SHT.MASTER = "${masterName}"`,
   `Headers found: ${Object.keys(H).length}`,
   last >= 2 ? `Last data row: ${last}` : 'Last data row: (none)'
 ];
 const warn = missing.length ? ['Missing headers: ' + missing.join(', ')] : [];


 const msg =
   'Resolver Master Diagnostic\n\n' +
   '✅ ' + ok.join('\n✅ ') + '\n\n' +
   (warn.length ? '⚠️ ' + warn.join('\n⚠️ ') + '\n\n' : '') +
   (last >= 2 ? ('Sample (last row):\n' + JSON.stringify(sample, null, 2)) : 'Sample: (no data rows)');


 ui.alert(msg);
}


function migrateFolderUrlToClientFolder(){
 const s = SH(SHT.MASTER);
 let H = headers_(SHT.MASTER);


 // Ensure "Client Folder" column exists
 if (!H['Client Folder']) {
   s.getRange(1, s.getLastColumn()+1).setValue('Client Folder');
   H = headers_(SHT.MASTER);
 }


 const colClient = H['Client Folder'];
 const colOld    = H['Folder URL'];   // may be missing


 const last = s.getLastRow();
 if (last >= 2 && colOld) {
   const oldVals = s.getRange(2, colOld, last-1, 1).getValues();
   const newVals = s.getRange(2, colClient, last-1, 1).getValues();


   // copy only when Client Folder is blank
   for (let i=0;i<oldVals.length;i++){
     const src = String(oldVals[i][0]||'').trim();
     const dst = String(newVals[i][0]||'').trim();
     if (src && !dst) newVals[i][0] = src;
   }
   s.getRange(2, colClient, last-1, 1).setValues(newVals);


   // delete the “Folder URL” column
   s.deleteColumn(colOld);
 }


 SpreadsheetApp.getUi().alert('Migration complete: Client Folder set; "Folder URL" removed.');
}

function debug_bootstrapForLastRealRow() {
  const sh = _openMaster_();
  const last = sh.getLastRow();
  if (last < 2) { Logger.log('No data rows'); return; }

  // Find the last row that actually has a RootApptID
  const H  = _headers_(sh);
  const colApId = H['RootApptID'];
  if (!colApId) throw new Error('Missing "RootApptID" column');

  const vals = sh.getRange(2, colApId, last-1, 1).getValues(); // 2..last
  let lastRowWithAp = -1;
  for (let i = vals.length - 1; i >= 0; i--) {
    const v = String(vals[i][0] || '').trim();
    if (v) { lastRowWithAp = i + 2; break; } // +2 accounts for header + 0-index
  }
  if (lastRowWithAp === -1) { Logger.log('No RootApptID found'); return; }

  const id = bootstrapApFolderForRow_(lastRowWithAp);
  Logger.log('Bootstrapped AP folder ID: ' + id + ' for row ' + lastRowWithAp);
}


/**
 * Ensure every row in 00_Master Appointments has a RootAppt Folder ID.
 * - Scans all rows 2..last
 * - If RootApptID is present and RootAppt Folder ID is blank, calls bootstrapApFolderForRow_.
 */
function backfillAllRootApptFolders() {
  const ssId = PROP_('SPREADSHEET_ID');
  if (!ssId) throw new Error('SPREADSHEET_ID script property not set');
  const ss = SpreadsheetApp.openById(ssId);
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet "00_Master Appointments"');

  const H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
  const iApId = H.indexOf('RootApptID');
  const iFid  = H.indexOf('RootAppt Folder ID');
  if (iApId < 0 || iFid < 0) throw new Error('Missing RootApptID / RootAppt Folder ID columns');

  const last = sh.getLastRow();
  if (last < 2) { Logger.log('No data rows'); return; }

  const apIds = sh.getRange(2, iApId+1, last-1, 1).getValues().map(r => String(r[0]||'').trim());
  const fids  = sh.getRange(2, iFid+1, last-1, 1).getValues().map(r => String(r[0]||'').trim());

  let fixed = 0;
  for (let i=0; i<apIds.length; i++) {
    const row = i + 2;
    if (apIds[i] && !fids[i]) {
      try {
        const id = bootstrapApFolderForRow_(row);   // your existing folder creator
        Logger.log(`Row ${row}: created RootAppt folder ${id}`);
        fixed++;
      } catch(e) {
        Logger.log(`Row ${row}: bootstrap error: ${e && e.message}`);
      }
    }
  }
  Logger.log(`backfillAllRootApptFolders: bootstrapped ${fixed} missing rows`);
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

