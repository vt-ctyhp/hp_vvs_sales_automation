// ====================================================================
// RESOLVER.GS - FIXED VERSION
// ====================================================================
// Applied fixes:
// ✅ FIX #1: Phone resolution - standardized helper
// ✅ FIX #2: DateTime sync - prevent blank overwrites
// ✅ FIX #3: Artifact creation - race condition protection
// ✅ FIX #4: Repair worker - coordinated (automatic via cache lock)
// ✅ FIX #5: Customer chain continuity - reuse Prospect Folder for repeat customers
// ====================================================================

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

// --- Lightweight profiler ---
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
function __wrap(label, fn){
  const t = Date.now();
  Logger.log('→ ' + label);
  try { return fn(); }
  finally { Logger.log('← ' + label + '  ' + (Date.now()-t) + 'ms'); }
}
function debug_introspectHelpers(){
  Logger.log('_findMostRecentPriorRow.length = ' + _findMostRecentPriorRow.length);
  Logger.log('_currentRowToObj_.length = ' + _currentRowToObj_.length);
}
function SS(){ return SpreadsheetApp.getActive(); }
function SH(name){ const s=SS().getSheetByName(name); if(!s) throw new Error(`Missing sheet: ${name}`); return s; }
const _HEADER_CACHE_ = {};
function headers_(name){
  if (_HEADER_CACHE_[name]) return _HEADER_CACHE_[name];
  const s = SH(name);
  const arr = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const map = {};
  arr.forEach((h, i) => { if (h) map[String(h).trim()] = i + 1; });
  _HEADER_CACHE_[name] = map;
  return map;
}
function setCell_(sheetName,row,colName,val){ const m=headers_(sheetName); const c=m[colName]; if(!c) throw new Error(`Column "${colName}" not found on ${sheetName}`); SH(sheetName).getRange(row,c).setValue(val); }
function getCell_(sheetName,row,colName){ const m=headers_(sheetName); const c=m[colName]; if(!c) return ''; return SH(sheetName).getRange(row,c).getValue(); }
function appendObj_(sheetName, obj){
  const s = SH(sheetName), H = headers_(sheetName);
  const rowArr = new Array(s.getLastColumn()).fill('');

  Object.keys(obj || {}).forEach(k => { if (H[k]) rowArr[H[k]-1] = obj[k]; });

  if (sheetName === SHT.MASTER) {
    const r = nextDataRow_(sheetName, LASTROW_SENTINELS);
    s.getRange(r, 1, 1, rowArr.length).setValues([rowArr]);
    return r;
  } else {
    s.appendRow(rowArr);
    return s.getLastRow();
  }
}
function log_(action, details){ appendObj_(SHT.LOG, {'Timestamp': new Date(), 'Action': action, 'Details': typeof details==='string'?details:JSON.stringify(details)}); }
function err_(where, why, payload){
  try {
    appendObj_(SHT.ERR, {
      'Timestamp': new Date(),
      'Where': where,
      'Why': why,
      'Payload': JSON.stringify(payload||{})
    });
  } catch(e) {
    try {
      appendObj_(SHT.LOG, {
        'Timestamp': new Date(),
        'Action': '[ERROR] ' + where,
        'Details': why + ' | ' + JSON.stringify(payload||{})
      });
    } catch(_) {
      Logger.log('[ERR_FALLBACK] ' + where + ': ' + why);
    }
  }
}

function nvGet(nv, key){
  if (nv[key] && nv[key][0] !== undefined) return nv[key][0];
  const k = Object.keys(nv || {}).find(k => k && k.trim().toLowerCase() === key.trim().toLowerCase());
  return k ? (nv[k][0] || '') : '';
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

function findBestMasterRowByUID_(uuid){
  if (!uuid) return 0;
  const s = SH(SHT.MASTER), H = headers_(SHT.MASTER);
  const cUid = H['CalendlyEventUID'] || 0;
  if (!cUid) return 0;
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  if (last < 2) return 0;

  const rows = s.getRange(2, 1, last - 1, s.getLastColumn()).getValues();
  let bestRow = 0, bestScore = -1;
  for (let i = 0; i < rows.length; i++){
    const r = rows[i];
    if (String(r[cUid - 1] || '') !== String(uuid)) continue;
    const rowIndex = i + 2;
    const score = dedupeCanonicalRowScore_(r, H, rowIndex);
    if (score > bestScore) {
      bestScore = score;
      bestRow = rowIndex;
    }
  }
  return bestRow;
}

function dedupeNormKey_(value){
  return String(value == null ? '' : value).trim().toLowerCase().replace(/\s+/g, ' ');
}

function dedupeNormDateKey_(value){
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, CFG.TZ, 'yyyy-MM-dd');
  }
  const s = String(value).trim();
  const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3];
  const mdY = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s);
  if (mdY) {
    return mdY[3] + '-' + String(mdY[1]).padStart(2, '0') + '-' + String(mdY[2]).padStart(2, '0');
  }
  const dt = new Date(s);
  if (!isNaN(dt.getTime())) return Utilities.formatDate(dt, CFG.TZ, 'yyyy-MM-dd');
  return dedupeNormKey_(s);
}

function dedupeNormTimeKey_(value){
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, CFG.TZ, 'HH:mm');
  }
  const s = String(value).trim();
  const m12 = /^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)$/i.exec(s);
  if (m12) {
    let h = parseInt(m12[1], 10);
    const ap = m12[3].toUpperCase();
    if (ap === 'AM' && h === 12) h = 0;
    if (ap === 'PM' && h !== 12) h += 12;
    return String(h).padStart(2, '0') + ':' + m12[2];
  }
  const m24 = /^(\d{1,2}):(\d{2})(?::\d{2})?$/.exec(s);
  if (m24) return String(parseInt(m24[1], 10)).padStart(2, '0') + ':' + m24[2];
  const dt = new Date('2000-01-01 ' + s);
  if (!isNaN(dt.getTime())) return Utilities.formatDate(dt, CFG.TZ, 'HH:mm');
  return dedupeNormKey_(s);
}

function dedupeContactKey_(emailLower, phoneNorm){
  const email = normEmail_(emailLower);
  const phone = normPhone_(phoneNorm);
  if (email) return 'email:' + email;
  if (phone) return 'phone:' + phone;
  return '';
}

function dedupeAppointmentFingerprint_(brand, visitDate, visitTime, visitType, emailLower, phoneNorm){
  const contact = dedupeContactKey_(emailLower, phoneNorm);
  const d = dedupeNormDateKey_(visitDate);
  const t = dedupeNormTimeKey_(visitTime);
  if (!contact || !d || !t) return '';
  return [
    dedupeNormKey_(brand),
    d,
    t,
    dedupeNormKey_(visitType),
    contact
  ].join('|');
}

function dedupeIsCurrentRow_(row, H){
  const status = H['Status'] ? row[H['Status'] - 1] : '';
  const active = H['Active?'] ? row[H['Active?'] - 1] : '';
  const s = dedupeNormKey_(status);
  const a = dedupeNormKey_(active);
  if (/cancel|resched|duplicate|superseded|inactive/.test(s)) return false;
  if (a === 'yes' || a === 'true' || a === '1') return true;
  if (a === 'no' || a === 'false' || a === '0') return false;
  return true;
}

function dedupeCanonicalRowScore_(row, H, rowIndex){
  let score = 0;
  if (dedupeIsCurrentRow_(row, H)) score += 1000;
  if (H['CalendlyEventUID'] && row[H['CalendlyEventUID'] - 1]) score += 100;
  if (H['Visit #'] && row[H['Visit #'] - 1]) score += 80;
  if (H['APPT_ID'] && row[H['APPT_ID'] - 1]) score += 40;
  if (H['RootApptID'] && row[H['RootApptID'] - 1]) score += 20;
  return score + Math.min(rowIndex || 0, 99999) / 100000;
}

function dedupeFingerprintForMasterRow_(row, H){
  const email = H['EmailLower'] ? row[H['EmailLower'] - 1] : (H['Email'] ? row[H['Email'] - 1] : '');
  const phone = H['PhoneNorm'] ? row[H['PhoneNorm'] - 1] : (H['Phone'] ? row[H['Phone'] - 1] : '');
  return dedupeAppointmentFingerprint_(
    H['Brand'] ? row[H['Brand'] - 1] : '',
    H['Visit Date'] ? row[H['Visit Date'] - 1] : '',
    H['Visit Time'] ? row[H['Visit Time'] - 1] : '',
    H['Visit Type'] ? row[H['Visit Type'] - 1] : '',
    email,
    phone
  );
}

function findCurrentMasterRowByFingerprint_(brand, visitDate, visitTime, visitType, emailLower, phoneNorm, excludeRow){
  const wanted = dedupeAppointmentFingerprint_(brand, visitDate, visitTime, visitType, emailLower, phoneNorm);
  if (!wanted) return 0;
  const s = SH(SHT.MASTER), H = headers_(SHT.MASTER);
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  if (last < 2) return 0;

  const rows = s.getRange(2, 1, last - 1, s.getLastColumn()).getValues();
  let bestRow = 0, bestScore = -1;
  for (let i = 0; i < rows.length; i++){
    const rowIndex = i + 2;
    if (excludeRow && rowIndex === excludeRow) continue;
    const r = rows[i];
    if (!dedupeIsCurrentRow_(r, H)) continue;
    if (dedupeFingerprintForMasterRow_(r, H) !== wanted) continue;
    const score = dedupeCanonicalRowScore_(r, H, rowIndex);
    if (score > bestScore) {
      bestScore = score;
      bestRow = rowIndex;
    }
  }
  return bestRow;
}

/** Shared: build a stable contact key (prefer email; fallback phone) */
function _contactKey_(brand, vtype, emailLower, phoneNorm){
  const b=(brand||'').toUpperCase().trim();
  const t=(vtype||'').toLowerCase().trim();
  const e=(emailLower||'').toLowerCase().trim();
  const p=(phoneNorm||'').trim();
  const id = e || p;
  return ['CANCEL', b, t, id].join(':');
}

function _rememberCancelUID_(brand, vtype, emailLower, phoneNorm, oldUid, ttlSec){
  try{
    const key = _contactKey_(brand, vtype, emailLower, phoneNorm);
    CacheService.getScriptCache().put(key, String(oldUid||''), ttlSec || 7200);
    return key;
  }catch(_){ return ''; }
}

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

function findRecentCanceledByContactAt_(emailLower, phoneNorm, minutes){
  const s = SH(SHT.MASTER), m = headers_(SHT.MASTER);
  const colE=m['EmailLower'], colP=m['PhoneNorm'], colSta=m['Status'], colCA=m['CanceledAt'];
  if (!colE || !colP || !colSta || !colCA) return 0;
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS); if (last < 2) return 0;
  const vals=s.getRange(2,1,last-1,s.getLastColumn()).getValues();
  const now=Date.now(), win=(minutes||120)*60*1000;
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

function splitName_(full){
  const t = (full||'').toString().trim();
  if (!t) return {first:'', last:''};
  const parts = t.split(/\s+/);
  return {first: parts[0], last: parts.slice(1).join(' ')};
}

const DEFAULT_DURATION_MIN = 30;

/********** NORMALIZERS **********/
function normEmail_(e){ return (e||'').toString().trim().toLowerCase(); }
function normPhone_(p){
  if(!p) return '';
  const d=(''+p).replace(/\D+/g,'');
  if(d.length===10) return '+1'+d;
  if(d.length===11 && d.startsWith('1')) return '+'+d;
  return d?('+'+d):'';
}

// ====================================================================
// FIX #1: PHONE RESOLUTION - STANDARDIZED HELPER
// ====================================================================
function resolvePhoneForRow_(rawPhone, row, oldRow, isReschedule) {
  const raw = String(rawPhone || '').trim();
  if (raw) {
    const norm = normPhone_(raw);
    if (norm) return { raw, norm };
  }
  if (row) {
    try {
      const currRaw = getCell_(SHT.MASTER, row, 'Phone');
      if (currRaw) {
        const currNorm = getCell_(SHT.MASTER, row, 'PhoneNorm') || normPhone_(currRaw);
        if (currNorm) return { raw: currRaw, norm: currNorm };
      }
    } catch(_) {}
  }
  if (isReschedule && oldRow) {
    try {
      const oldRaw = getCell_(SHT.MASTER, oldRow, 'Phone');
      if (oldRaw) {
        const oldNorm = getCell_(SHT.MASTER, oldRow, 'PhoneNorm') || normPhone_(oldRaw);
        if (oldNorm) return { raw: oldRaw, norm: oldNorm };
      }
    } catch(_) {}
  }
  return { raw: '', norm: '' };
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
  return '';
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
              .replace(/[\\/:*?"<>|]/g, '-');
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
  const prev = getCell_(SHT.MASTER, row, 'Automation Notes') || '';

  // ✅ Fix regex: bắt cả " @ 2026-..." có space
  const stripTs = s => s.replace(/\s*@\s*[\d]{4}-[\d\-T:.Z]+$/, '').trim();
  const msgCore = stripTs(msg);

  if (prev.split('\n').some(line => stripTs(line) === msgCore)) return;

  setCell_(SHT.MASTER, row, 'Automation Notes', (prev ? prev + '\n' : '') + msg);
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
  try {
    return bootstrapApFolderForRow_(rowIdx);
  } catch (e) {
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

function _getAppointmentsHome_() {
  const homeId = _SP.getProperty('APPOINTMENTS_FOLDER_ID');
  if (homeId) return DriveApp.getFolderById(homeId);
  const created = DriveApp.createFolder('[SYS] Appointments');
  _SP.setProperty('APPOINTMENTS_FOLDER_ID', created.getId());
  return created;
}

function _ensureApFolderUnderHome_(apId) {
  const home = _getAppointmentsHome_();
  const it = home.getFoldersByName(apId);
  return it.hasNext() ? it.next() : home.createFolder(apId);
}

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

function bootstrapApFolderForRow_(row) {
  const sh = _openMaster_();
  const H  = _headers_(sh);

  const colApId = H['RootApptID'];
  if (!colApId) throw new Error('Missing "RootApptID" column in Master');

  const apId = String(sh.getRange(row, colApId).getValue() || '').trim();
  if (!/^AP-\d{8}-\d{3}$/i.test(apId)) {
    throw new Error('Invalid or empty RootApptID on row ' + row + ': ' + apId);
  }

  const prospectIdColNameCandidates = ['ProspectFolderID', 'RootAppt Folder ID', 'AP Folder ID'];
  let colExistingId = null;
  for (const name of prospectIdColNameCandidates) {
    if (H[name]) { colExistingId = H[name]; break; }
  }

  let apFolder;
  if (colExistingId) {
    const existing = String(sh.getRange(row, colExistingId).getValue() || '').trim();
    if (existing) {
      try {
        apFolder = DriveApp.getFolderById(existing);
      } catch(_) {}
    }
  }

  if (!apFolder) {
    apFolder = _ensureApFolderUnderHome_(apId);
  }

  _ensureApSubfoldersByFolderId_(apFolder.getId());
  _writeApFolderIdToMasterRow_(row, apFolder.getId());

  return apFolder.getId();
}

function ensureBootstrapForRecentRows_() {
  // ── FIX: ngăn chạy đồng thời khi trigger fire 2 lần ──
  const lock = LockService.getDocumentLock();
  const gotLock = lock.tryLock(5000);
  if (!gotLock) {
    Logger.log('ensureBootstrapForRecentRows_: already running, skipped.');
    return;
  }

  try {
    const sh = _openMaster_();
    const H  = _headers_(sh);
    const colApId = H['RootApptID'];
    const colFid  = H['RootAppt Folder ID'];
    const colPfId = H['ProspectFolderID'];

    if (!colApId || !colFid) {
      Logger.log('Missing required headers (RootApptID / RootAppt Folder ID)');
      return;
    }

    const last = sh.getLastRow();
    if (last < 2) return;

    const N = Math.min(5, last - 1);
    const startRow = Math.max(2, last - N + 1);

    const apIds = sh.getRange(startRow, colApId, N, 1).getValues();
    const fids  = sh.getRange(startRow, colFid,  N, 1).getValues();
    const pfIds = colPfId
      ? sh.getRange(startRow, colPfId, N, 1).getValues()
      : Array(N).fill(['']);

    let bootstrapped = 0;
    for (let i = 0; i < N; i++) {
      const row  = startRow + i;
      const ap   = String(apIds[i][0] || '').trim();
      const fid  = String(fids[i][0]  || '').trim();
      const pfid = String(pfIds[i][0] || '').trim();

      if (!ap)         continue;
      if (fid || pfid) continue;

      try {
        const id = bootstrapApFolderForRow_(row);
        Logger.log(`Bootstrapped AP folder ${id} for row ${row}`);
        bootstrapped++;
      } catch (e) {
        Logger.log(`Row ${row}: bootstrap error: ${e && (e.message || e)}`);
      }
    }

    if (bootstrapped) Logger.log(`ensureBootstrapForRecentRows_: bootstrapped ${bootstrapped} row(s).`);

  } finally {
    lock.releaseLock();
  }
}

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

  if (!colE && !colP) return 1;

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
    const active = colAct ? /^yes$/i.test(String(r[colAct-1]||'')) : /scheduled|rescheduled/i.test(status);
    const completed = /completed/i.test(status);
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

// ====================================================================
// FIX #2: DATETIME SYNC - PREVENT BLANK OVERWRITES
// ====================================================================
// function syncVisitDateTime_(row, vdate, vtime) {
//   const TZ = CFG.TZ || 'America/Los_Angeles';
  
//   if (!vdate || !vtime) {
//     Logger.log(`[syncDateTime] row ${row} skipped - incomplete data (vdate="${vdate}" vtime="${vtime}")`);
//     return '';
//   }
  
//   const oldISO   = getCell_(SHT.MASTER, row, 'ApptDateTime (ISO)') || '';
//   const oldVDate = getCell_(SHT.MASTER, row, 'Visit Date') || '';
//   const oldVTime = getCell_(SHT.MASTER, row, 'Visit Time') || '';
  
//   let newISO = '';
//   try {
//     const parts = vdate.split('/');
//     let dateStr = vdate;
    
//     if (parts.length === 3) {
//       const mm = parts[0].padStart(2, '0');
//       const dd = parts[1].padStart(2, '0');
//       const yyyy = parts[2];
//       dateStr = `${yyyy}-${mm}-${dd}`;
//     }
    
//     const combined = `${dateStr} ${vtime}`;
//     const dt = new Date(combined);
    
//     if (!isNaN(dt.getTime())) {
//       newISO = Utilities.formatDate(dt, TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");
//     }
//   } catch (e) {
//     err_('syncVisitDateTime_', 'Failed to parse date/time', { vdate, vtime, error: e.message });
//     return '';
//   }
  
//   const isReschedule = oldISO && newISO && oldISO !== newISO;
  
//   if (isReschedule) {
//     try {
//       const oldDT = Utilities.formatDate(new Date(oldISO), TZ, 'yyyy-MM-dd HH:mm:ss z');
//       const newDT = Utilities.formatDate(new Date(newISO), TZ, 'yyyy-MM-dd HH:mm:ss z');
//       Logger.log(`[syncDateTime] row ${row} RESCHEDULE detected:`);
//       Logger.log(`  OLD: ${oldDT}`);
//       Logger.log(`  NEW: ${newDT}`);
//     } catch (_) {}
//   }
  
//   if (newISO) {
//     setCell_(SHT.MASTER, row, 'ApptDateTime (ISO)', newISO);
//     setCell_(SHT.MASTER, row, 'Visit Date', vdate);
//     setCell_(SHT.MASTER, row, 'Visit Time', vtime);
//     Logger.log(`[syncDateTime] row ${row} synced: ${newISO}`);
//   }
  
//   return newISO;
// }
function syncVisitDateTime_(row, vdate, vtime) {
  const TZ = CFG.TZ || 'America/Los_Angeles';

  if (!vdate || !vtime) {
    Logger.log(`[syncDateTime] row ${row} skipped - incomplete data`);
    return '';
  }

  // ── FIX: Calendly gửi ISO object → parse trực tiếp ──
  let newISO = '';
  try {
    // Case 1: vdate là full ISO (Calendly format)
    // "2026-04-29T07:00:00.000Z" + "1899-12-30T22:30:00.000Z"
    const isISODate = /^\d{4}-\d{2}-\d{2}T/.test(String(vdate));
    const isISOTime = /^\d{4}-\d{2}-\d{2}T/.test(String(vtime));

    if (isISODate && isISOTime) {
      // Lấy phần date từ vdate, phần time từ vtime
      const datePart = new Date(vdate);
      const timePart = new Date(vtime);

      const combined = new Date(
        datePart.getUTCFullYear(),
        datePart.getUTCMonth(),
        datePart.getUTCDate(),
        timePart.getUTCHours(),
        timePart.getUTCMinutes(),
        0
      );
      newISO = Utilities.formatDate(combined, TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");

    // Case 2: format cũ MM/DD/YYYY + "2:30:00 PM"
    } else {
      const parts = String(vdate).split('/');
      let dateStr = vdate;
      if (parts.length === 3) {
        dateStr = `${parts[2]}-${parts[0].padStart(2,'0')}-${parts[1].padStart(2,'0')}`;
      }
      const dt = new Date(`${dateStr} ${vtime}`);
      if (!isNaN(dt.getTime())) {
        newISO = Utilities.formatDate(dt, TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");
      }
    }
  } catch(e) {
    Logger.log(`[syncDateTime] row ${row} parse error: ${e.message}`);
    return '';
  }

  if (newISO) {
    setCell_(SHT.MASTER, row, 'ApptDateTime (ISO)', newISO);
    setCell_(SHT.MASTER, row, 'Visit Date', vdate);
    setCell_(SHT.MASTER, row, 'Visit Time', vtime);
    Logger.log(`[syncDateTime] row ${row} synced: ${newISO}`);
  }

  return newISO;
}

function syncBookedAt_(row, submittedAt) {
  const TZ = CFG.TZ || 'America/Los_Angeles';
  try {
    const bookedAtISO = Utilities.formatDate(
      new Date(submittedAt || new Date()),
      TZ,
      "yyyy-MM-dd'T'HH:mm:ssXXX"
    );
    setOnce_(SHT.MASTER, row, 'Booked At (ISO)', bookedAtISO);
  } catch (e) {
    err_('syncBookedAt_', 'Failed to sync Booked At', { row, submittedAt, error: e.message });
  }
}

function buildIntakeData_(rowIdx){
  const m = headers_(SHT.MASTER);
  const s = SH(SHT.MASTER);
  function val(h){ return (m[h]? s.getRange(rowIdx, m[h]).getValue() : '') || ''; }

  const tz = CFG.TZ || 'America/Los_Angeles';
  const iso = val('ApptDateTime (ISO)');

  let apptDate = '';
  let apptTime = '';
  let apptDT = '';

  if (iso) {
    try {
      const date = new Date(iso);
      if (!isNaN(date.getTime())) {
        apptDate = Utilities.formatDate(date, tz, 'EEE, MMM d, yyyy');
        apptTime = Utilities.formatDate(date, tz, 'h:mm a');
        apptDT = Utilities.formatDate(date, tz, 'EEE, MMM d, yyyy h:mm a z');
      }
    } catch (e) {
      Logger.log(`[buildIntakeData] row ${rowIdx} ISO parse failed: ${iso}`);
    }
  }

  if (!apptDate && !apptTime) {
    apptDate = val('Visit Date') || '';
    apptTime = val('Visit Time') || '';
  }

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
    Timestamp:          Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss'),
    RescheduledFromUID: val('RescheduledFromUID'),
    RescheduledToUID:   val('RescheduledToUID'),
    CanceledAt:         (function(){
                          const ca = val('CanceledAt');
                          try { return ca ? Utilities.formatDate(new Date(ca), tz, 'yyyy-MM-dd HH:mm') : ''; }
                          catch(_){ return ca || ''; }
                        })()
  };
  return data;
}

function backfillMissingISO_() {
  const s = SH(SHT.MASTER);
  const H = headers_(SHT.MASTER);
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  
  if (last < 2) { Logger.log('No data rows'); return; }
  
  const colISO   = H['ApptDateTime (ISO)'] || 0;
  const colVDate = H['Visit Date'] || 0;
  const colVTime = H['Visit Time'] || 0;
  
  if (!colISO || !colVDate || !colVTime) { Logger.log('Missing required columns'); return; }
  
  const isoVals   = s.getRange(2, colISO, last-1, 1).getValues();
  const vdateVals = s.getRange(2, colVDate, last-1, 1).getValues();
  const vtimeVals = s.getRange(2, colVTime, last-1, 1).getValues();
  
  let fixed = 0;
  for (let i = 0; i < isoVals.length; i++) {
    const row = i + 2;
    const iso = String(isoVals[i][0] || '').trim();
    const vdate = String(vdateVals[i][0] || '').trim();
    const vtime = String(vtimeVals[i][0] || '').trim();
    
    if (!iso && vdate && vtime) {
      try {
        syncVisitDateTime_(row, vdate, vtime);
        fixed++;
        Logger.log(`Row ${row}: backfilled ISO from Visit Date/Time`);
      } catch (e) {
        Logger.log(`Row ${row}: backfill failed: ${e.message}`);
      }
    }
  }
  Logger.log(`backfillMissingISO_: fixed ${fixed} row(s)`);
}

function fillIntakeDocPlaceholders_(docId, data){
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const header = doc.getHeader();
  const footer = doc.getFooter();
  const escape = s => s.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&');
  Object.keys(data).forEach(k=>{
    const pat = '\\{\\{\\s*' + escape(k) + '\\s*\\}\\}';
    const val = data[k] == null ? '' : String(data[k]);
    body.replaceText(pat, val);
    if (header) header.replaceText(pat, val);
    if (footer) footer.replaceText(pat, val);
  });
  doc.saveAndClose();
}

function upsertAutofillBlock_(docId, data){
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

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
    if (looksLikeAuto){ body.removeChild(tbl); break; }
  }

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

  if (data.RescheduledFromUID || data.RescheduledToUID || data.CanceledAt){
    rows.push(['', '']);
    pushRow('Rescheduled From', data.RescheduledFromUID);
    pushRow('Rescheduled To',   data.RescheduledToUID);
    pushRow('Canceled At',      data.CanceledAt);
  }

  if (rows.length){
    const table = body.appendTable(rows);
    table.setBorderWidth(0);
    for (let r = 0; r < table.getNumRows(); r++){
      table.getRow(r).getCell(0).editAsText().setBold(true);
    }
  }

  doc.saveAndClose();
}

function ensureAndFillIntakeDocForRow_(rowIdx){
  const intakeUrl = getCell_(SHT.MASTER, rowIdx, 'IntakeDocURL');
  if (intakeUrl) {
    const docId = idFromUrl_(String(intakeUrl));
    if (!docId) return;
    const data = buildIntakeData_(rowIdx);
    try { fillIntakeDocPlaceholders_(docId, data); } catch(_){}
    try { upsertAutofillBlock_(docId, data); } catch(_){}
    return;
  }
  ensureArtifactsForRow_(rowIdx);
}

function ensureAndFillChecklistDocForRow_(rowIdx){
  const url = getCell_(SHT.MASTER, rowIdx, 'Checklist URL');
  if (url) {
    const docId = idFromUrl_(String(url));
    if (!docId) return;
    try { fillIntakeDocPlaceholders_(docId, buildIntakeData_(rowIdx)); } catch(_){}
    return;
  }
  ensureArtifactsForRow_(rowIdx);
}

function ensureAndFillQuotationForRow_(rowIdx){
  const url = getCell_(SHT.MASTER, rowIdx, 'Quotation URL');
  if (url) {
    const ssId = idFromUrl_(String(url));
    if (!ssId) return;
    try { fillSheetPlaceholders_(ssId, buildIntakeData_(rowIdx)); } catch(_){}
    return;
  }
  ensureArtifactsForRow_(rowIdx);
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

const LASTROW_SENTINELS = ['APPT_ID','Customer Name','EmailLower','Timestamp'];

function _isConsultVisit_(vtRaw){
  const t = String(vtRaw || '').trim().toLowerCase();
  return t === 'appointment' || t === 'diamond viewing';
}

function stampSalesStageIfConsult_(row, vtypeFromForm){
  const vt = (vtypeFromForm != null && vtypeFromForm !== '')
    ? String(vtypeFromForm).trim()
    : String(getCell_(SHT.MASTER, row, 'Visit Type') || '').trim();

  if (_isConsultVisit_(vt)) {
    setCell_(SHT.MASTER, row, 'Sales Stage', 'Appointment');
  }
}

function lastDataRow_(sheetName, sentinels){
  const s = SH(sheetName), H = headers_(sheetName);
  const last = s.getLastRow();
  if (last < 2) return 1;

  const cols = (sentinels||[]).map(name => H[name]).filter(Boolean);
  if (!cols.length) return last;

  let best = 1;
  for (let c of cols){
    const n = Math.max(0, last - 1);
    if (n === 0) continue;
    const vals = s.getRange(2, c, n, 1).getValues();
    for (let i = vals.length - 1; i >= 0; i--){
      const v = vals[i][0];
      if (v !== '' && String(v).trim() !== '') {
        best = Math.max(best, i + 2);
        break;
      }
    }
  }
  return best;
}

function nextDataRow_(sheetName, sentinels){
  const sheetActualLast = SH(sheetName).getLastRow();
  const sentinelLast    = lastDataRow_(sheetName, sentinels);
  
  // Luôn append SAU row cuối cùng thực tế của sheet
  // bất kể Walk-In hay Form booking
  return Math.max(sheetActualLast, sentinelLast, 1) + 1;
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

function _currentRowToObj_(rowIdx){
  const s = SH(SHT.MASTER), H = headers_(SHT.MASTER);
  const email = _getByAliases_(s, rowIdx, H, COL_ALIASES.EmailLower).trim().toLowerCase();
  const phone = _getByAliases_(s, rowIdx, H, COL_ALIASES.PhoneNorm).trim();
  const uid   = (H['CalendlyEventUID'] ? String(s.getRange(rowIdx, H['CalendlyEventUID']).getValue()||'') : '');
  return { EmailLower: email, PhoneNorm: phone, CalendlyEventUID: uid };
}

function _findMostRecentPriorRow(cur, curRowIdx){
  const s = SH(SHT.MASTER), H = headers_(SHT.MASTER);
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  if (last < 2) return null;

  const cEmail = _firstHeaderIndex_(H, COL_ALIASES.EmailLower);
  const cPhone = _firstHeaderIndex_(H, COL_ALIASES.PhoneNorm);
  const cUID   = H['CalendlyEventUID'] || 0;
  const cISO   = _firstHeaderIndex_(H, COL_ALIASES.ApptIso);
  const cTS    = _firstHeaderIndex_(H, COL_ALIASES.Timestamp);
  const cInt   = _firstHeaderIndex_(H, COL_ALIASES.IntakeLink);
  const cChk   = _firstHeaderIndex_(H, COL_ALIASES.ChecklistLink);
  const cQuo   = _firstHeaderIndex_(H, COL_ALIASES.QuotationLink);
  const cPfId  = H['ProspectFolderID'] || 0;

  const n = last - 1;
  const pull = (col) => col
    ? s.getRange(2, col, n, 1).getValues().map(a => String(a[0] || ''))
    : Array(n).fill('');

  const emailVec   = pull(cEmail).map(v => v.trim().toLowerCase());
  const phoneVec   = pull(cPhone).map(v => v.trim());
  const uidVec     = pull(cUID).map(v => v.trim());
  const isoVec     = pull(cISO);
  const tsVec      = pull(cTS);
  const intakeVec  = pull(cInt);
  const checkVec   = pull(cChk);
  const quoteVec   = pull(cQuo);
  const pfIdVec    = pull(cPfId);

  const wantEmail  = String(cur && cur.EmailLower || '').trim().toLowerCase();
  const wantPhone  = String(cur && cur.PhoneNorm  || '').trim();
  const curUID     = String(cur && cur.CalendlyEventUID || '').trim();
  const selfIdx    = Number(curRowIdx || 0);

  for (let i = n - 1; i >= 0; i--) {
    const r = i + 2;
    if (selfIdx && r === selfIdx) continue;
    if (cUID && curUID && uidVec[i] && uidVec[i] === curUID) continue;

    const same = (wantEmail && emailVec[i] === wantEmail) ||
                 (wantPhone && phoneVec[i] === wantPhone);
    if (!same) continue;

    // ✅ FIX: Return bất kỳ prior row cùng khách — dù chưa có artifact
    return {
      rowIndex:        r,
      IntakeDocURL:    intakeVec[i] || '',
      ChecklistURL:    checkVec[i]  || '',
      QuotationURL:    quoteVec[i]  || '',
      ProspectFolderID: pfIdVec[i] || ''
    };
  }

  return null;
}

// ====================================================================
// FIX #3: ARTIFACT CREATION - RACE CONDITION PROTECTION
// ====================================================================
function ensureArtifactsForRow_(row) {
  const lockKey = `artifact_lock_${row}`;
  const cache = CacheService.getScriptCache();
  
  if (cache.get(lockKey)) {
    Logger.log(`[artifact] row ${row} locked by another process - skipping`);
    return;
  }
  
  cache.put(lockKey, 'locked', 120);
  
  try {
    _ensureArtifactsForRowImpl_(row);
  } finally {
    cache.remove(lockKey);
  }
}

// ====================================================================
// FIX #5: CUSTOMER CHAIN CONTINUITY
// ====================================================================
/**
 * Resolve the Prospect Folder for a new appointment row.
 *
 * Priority order (highest → lowest):
 *   1. ProspectFolderID already set on THIS row (idempotent re-run)
 *   2. ProspectFolderID from the most-recent PRIOR row for same customer
 *      → This is the key fix: repeat customers reuse the SAME folder
 *   3. Create a brand-new Prospect Folder under the Client Folder
 *
 * Returns { folder: DriveFolder, pfId: string, isReused: boolean }
 */
function _resolveProspectFolder_(row, clientFolder, apptId, emailLower, phoneNorm, calUID) {
  // 1️⃣ Already set on this row
  const existingPfId = String(getCell_(SHT.MASTER, row, 'ProspectFolderID') || '').trim();
  if (existingPfId) {
    try {
      const f = DriveApp.getFolderById(existingPfId);
      Logger.log(`[prospectFolder] row ${row} → using existing pfId on this row: ${existingPfId}`);
      return { folder: f, pfId: existingPfId, isReused: false };
    } catch (_) {
      Logger.log(`[prospectFolder] row ${row} → existing pfId invalid, will search prior`);
    }
  }

  // 2️⃣ Check customer history for an existing Prospect Folder
  const cur = {
    EmailLower:       String(emailLower || '').trim().toLowerCase(),
    PhoneNorm:        String(phoneNorm  || '').trim(),
    CalendlyEventUID: String(calUID     || '').trim()
  };
  const prior = _findMostRecentPriorRow(cur, row);

  if (prior && prior.ProspectFolderID) {
    try {
      const f = DriveApp.getFolderById(prior.ProspectFolderID);
      Logger.log(`[prospectFolder] row ${row} → REUSING prior ProspectFolderID ${prior.ProspectFolderID} from row ${prior.rowIndex}`);
      return { folder: f, pfId: prior.ProspectFolderID, isReused: true, priorRow: prior.rowIndex };
    } catch (_) {
      Logger.log(`[prospectFolder] row ${row} → prior pfId invalid, will create new`);
    }
  }

  // 3️⃣ Create brand-new Prospect Folder
  const newFolder = ensureProspectFolder_(clientFolder, apptId);
  Logger.log(`[prospectFolder] row ${row} → created NEW prospect folder for ${apptId}`);
  return { folder: newFolder, pfId: newFolder.getId(), isReused: false };
}

/**
 * Internal implementation — artifact creation with customer chain continuity.
 *
 * FLOW (fixed):
 *   1. Resolve Client Folder (by brand + name/email/phone)
 *   2. *** Resolve Prospect Folder via customer history FIRST ***
 *      → If prior row exists with ProspectFolderID → reuse it
 *      → Only create new if truly first-time customer
 *   3. Inherit artifact URLs from prior row (if reusing folder)
 *   4. Create only MISSING artifacts (never duplicates)
 *   5. Atomic write to Master sheet
 *   6. Fill placeholders
 */
function _ensureArtifactsForRowImpl_(row) {
  Logger.log('=== ARTIFACT START row=' + row + ' ===');

  const s = SH(SHT.MASTER);
  const H = headers_(SHT.MASTER);
  const totalCols = s.getLastColumn();

  // ✅ Đọc toàn bộ row 1 lần duy nhất
  const rowSnap = s.getRange(row, 1, 1, totalCols).getValues()[0];
  const snap = (col) => H[col] ? String(rowSnap[H[col]-1] || '').trim() : '';

  const brand      = snap('Brand');
  const apptId     = snap('APPT_ID');
  const emailLower = snap('EmailLower');
  const phoneNorm  = snap('PhoneNorm');
  const calUID     = snap('CalendlyEventUID');

  if (!brand || !apptId) {
    Logger.log('Abort: missing brand or apptId');
    return;
  }

  let cfId      = snap('ClientFolderID');
  let cfUrl     = snap('Client Folder');
  let pfId      = snap('ProspectFolderID');
  let intakeUrl = snap('IntakeDocURL');
  let chkUrl    = snap('Checklist URL');
  let quoUrl    = snap('Quotation URL');

  const pending = {};

  // ── 1. CLIENT FOLDER ────────────────────────────────────────
  // Priority: (1) cfId on this row → (2) cfId from prior row same contact → (3) create new
  let clientFolder = null;
  if (cfId) {
    try { clientFolder = DriveApp.getFolderById(cfId); } catch (_) {}
  }

  // Check prior row for same customer (catches name changes like "babyboo" vs "test2704")
  if (!clientFolder) {
    try {
      const curObj = {
        EmailLower:       String(emailLower || '').trim().toLowerCase(),
        PhoneNorm:        String(phoneNorm  || '').trim(),
        CalendlyEventUID: String(calUID     || '').trim()
      };
      const prior = _findMostRecentPriorRow(curObj, row);
      if (prior && prior.rowIndex) {
        const priorCfId = String(getCell_(SHT.MASTER, prior.rowIndex, 'ClientFolderID') || '').trim();
        if (priorCfId) {
          try {
            clientFolder = DriveApp.getFolderById(priorCfId);
            cfId = priorCfId;
            pending['ClientFolderID'] = cfId;
            Logger.log('[clientFolder] REUSED from prior row ' + prior.rowIndex + ': ' + priorCfId);
          } catch(_) {
            Logger.log('[clientFolder] prior cfId invalid: ' + priorCfId);
          }
        }
      }
    } catch(_) {}
  }

  // Create new only if truly first-time customer
  if (!clientFolder) {
    const custName = snap('Customer Name');
    clientFolder = ensureClientFolder_(brand, custName, phoneNorm, emailLower);
    cfId = clientFolder.getId();
    pending['ClientFolderID'] = cfId;
    Logger.log('[clientFolder] CREATED NEW for: ' + snap('Customer Name'));
  }

  if (!cfUrl) {
    cfUrl = clientFolder.getUrl();
    pending['Client Folder'] = cfUrl;
  }

  // ── 2. PROSPECT FOLDER ──────────────────────────────────────
  const pfResult = _resolveProspectFolder_(
    row, clientFolder, apptId, emailLower, phoneNorm, calUID
  );
  pfId = pfResult.pfId;
  const prospectFolder = pfResult.folder;
  const isRepeatCustomer = pfResult.isReused;

  if (!snap('ProspectFolderID') || snap('ProspectFolderID') !== pfId) {
    pending['ProspectFolderID'] = pfId;
  }

  if (isRepeatCustomer) {
    Logger.log('[artifact] REPEAT CUSTOMER: reusing ProspectFolderID=' + pfId);
    _appendNote_(row, 'Repeat customer: reusing ProspectFolderID=' + pfId + ' @ ' + new Date().toISOString());
  }

  // ── 2b. KẾ THỪA ARTIFACT URLs TỪ PRIOR ROW ─────────────────
  // FIX: repeat customer → inherit URLs thay vì tạo file mới với apptId mới
  if (isRepeatCustomer && pfResult.priorRow) {
    try {
      const priorRow    = pfResult.priorRow;
      const priorIntake = getCell_(SHT.MASTER, priorRow, 'IntakeDocURL')  || '';
      const priorChk    = getCell_(SHT.MASTER, priorRow, 'Checklist URL') || '';
      const priorQuo    = getCell_(SHT.MASTER, priorRow, 'Quotation URL') || '';

      if (priorIntake && !intakeUrl) {
        intakeUrl = priorIntake;
        pending['IntakeDocURL'] = priorIntake;
        Logger.log('[artifact] Inherited IntakeDocURL from row ' + priorRow);
      }
      if (priorChk && !chkUrl) {
        chkUrl = priorChk;
        pending['Checklist URL'] = priorChk;
        Logger.log('[artifact] Inherited Checklist URL from row ' + priorRow);
      }
      if (priorQuo && !quoUrl) {
        quoUrl = priorQuo;
        pending['Quotation URL'] = priorQuo;
        Logger.log('[artifact] Inherited Quotation URL from row ' + priorRow);
      }
    } catch(e) {
      Logger.log('[artifact] inherit URLs error: ' + e.message);
    }
  }

  // ── 3. SCAN FOLDER 1 LẦN ────────────────────────────────────
  // Chỉ scan nếu còn thiếu URLs sau khi kế thừa từ prior row
  const folder = DriveApp.getFolderById(pfId);
  const existingFiles = {};
  if (!intakeUrl || !chkUrl || !quoUrl) {
    const fileIter = folder.getFiles();
    while (fileIter.hasNext()) {
      const f = fileIter.next();
      existingFiles[f.getName()] = f.getUrl();
    }
    Logger.log('[artifact] folder scan: ' + Object.keys(existingFiles).length + ' files found');
  } else {
    Logger.log('[artifact] folder scan skipped — all URLs inherited from prior row');
  }

  // ── 4. RECOVER OR CREATE ─────────────────────────────────────
  function recoverOrCreate(expectedName, existingUrl, templateFn, pendingKey) {
    if (existingUrl) return existingUrl;

    // ✅ Dùng cache scan thay vì gọi Drive lại
    if (existingFiles[expectedName]) {
      const url = existingFiles[expectedName];
      pending[pendingKey] = url;
      Logger.log('Recovered: ' + expectedName);
      return url;
    }

    const tplId = templateFn(brand);
    if (!tplId) return '';

    try {
      const copy = DriveApp.getFileById(tplId).makeCopy(expectedName, folder);
      const url = copy.getUrl();
      pending[pendingKey] = url;
      existingFiles[expectedName] = url; // update cache
      Logger.log('Created: ' + expectedName);
      return url;
    } catch (e) {
      Logger.log('ERROR creating ' + expectedName + ': ' + e.message);
      return '';
    }
  }

  intakeUrl = recoverOrCreate(
    brand + ' \u2013 ' + apptId + ' \u2013 Intake',
    intakeUrl, intakeTemplateIdForBrand_, 'IntakeDocURL'
  );
  chkUrl = recoverOrCreate(
    brand + ' \u2013 ' + apptId + ' \u2013 Checklist',
    chkUrl, checklistTemplateIdForBrand_, 'Checklist URL'
  );
  quoUrl = recoverOrCreate(
    brand + ' \u2013 ' + apptId + ' \u2013 Quotation',
    quoUrl, quotationTemplateIdForBrand_, 'Quotation URL'
  );

  // ── 5. ATOMIC WRITE (1 lần duy nhất) ─────────────────────────
  if (Object.keys(pending).length) {
    Logger.log('[artifact] writing ' + Object.keys(pending).length + ' fields: ' + Object.keys(pending).join(', '));
    _atomicWriteUrls_(row, pending);
  } else {
    Logger.log('[artifact] nothing to write - all fields already set');
  }

  // ── 6. GOOGLE SLIDES ─────────────────────────────────────────
  try {
    const slidesName = brand + ' \u2013 ' + apptId + ' \u2013 Slides';
    const clientFiles = {};
    const cf = clientFolder.getFiles();
    while (cf.hasNext()) {
      const f = cf.next();
      clientFiles[f.getName()] = f.getUrl();
    }
    if (!clientFiles[slidesName]) {
      const tplId = slidesTemplateIdForBrand_(brand);
      if (tplId) {
        const copy = DriveApp.getFileById(tplId).makeCopy(slidesName, clientFolder);
        const pres = SlidesApp.openById(copy.getId());
        const custName = snap('Customer Name');
        if (custName) _insertClientName_(pres.getSlides()[0], custName);
        _ensureTenBlankSlides_(pres);
        pres.saveAndClose();
        Logger.log('[Slides] Created: ' + slidesName);
      }
    } else {
      Logger.log('[Slides] Already exists: ' + slidesName);
    }
  } catch (e) {
    Logger.log('[Slides] ERROR: ' + e.message);
  }

  // ── 7. FILL PLACEHOLDERS ─────────────────────────────────────
  const hadNewFiles = Object.keys(pending).some(k =>
    k === 'IntakeDocURL' || k === 'Checklist URL' || k === 'Quotation URL'
  );

  if (!hadNewFiles) {
    Logger.log('[artifact] skip placeholder fill - no new files created');
    Logger.log('=== ARTIFACT END row=' + row + ' ===');
    return;
  }

  Logger.log('[artifact] filling placeholders for new files...');
  const data = buildIntakeData_(row);

  try {
      if (intakeUrl && pending['IntakeDocURL']) {
        const docId = idFromUrl_(intakeUrl);
        if (docId) {
          // ── Fill placeholders vào Doc trước ──────────────────
          fillIntakeDocPlaceholders_(docId, data);
          upsertAutofillBlock_(docId, data);
          Logger.log('[PDF] Intake Doc filled, exporting to PDF...');

          // ── Export PDF → ghi đè URL vào IntakeDocURL ─────────
          let destFolder;
          try { destFolder = DriveApp.getFolderById(pfId); } catch(_) {}

          if (destFolder) {
            const pdfUrl = exportIntakeDocToPdf_(docId, destFolder, brand, apptId);
            if (pdfUrl) {
              // Ghi PDF URL vào IntakeDocURL (thay thế Docs URL)
              pending['IntakeDocURL'] = pdfUrl;
              SH(SHT.MASTER).getRange(row, headers_(SHT.MASTER)['IntakeDocURL']).setValue(pdfUrl);
              SpreadsheetApp.flush();
              Logger.log('[PDF] ✅ IntakeDocURL updated to PDF URL');

              // ── Xóa Doc gốc sau khi đã có PDF ────────────────
              try {
                DriveApp.getFileById(docId).setTrashed(true);
                Logger.log('[PDF] Doc gốc đã xóa: ' + docId);
              } catch(e) {
                Logger.log('[PDF] Không xóa được Doc gốc: ' + e.message);
              }
            }
          }
        }
      }
    } catch (e) {
      Logger.log('[PDF] ERROR in intake flow: ' + e.message);
    }

  try {
    if (chkUrl && pending['Checklist URL']) {
      const id = idFromUrl_(chkUrl);
      if (id) fillIntakeDocPlaceholders_(id, data);
    }
  } catch (_) {}

  try {
    if (quoUrl && pending['Quotation URL']) {
      const id = idFromUrl_(quoUrl);
      if (id) fillSheetPlaceholders_(id, data);
    }
  } catch (_) {}

  Logger.log('=== ARTIFACT END row=' + row + ' ===');
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

function keep_(rawVal, col, row) {
  if (rawVal == null) return String(getCell_(SHT.MASTER, row, col) || '');
  const t = String(rawVal).trim();
  return t !== '' ? t : String(getCell_(SHT.MASTER, row, col) || '');
}

// ====================================================================
// MAIN RESOLVER - FORM SUBMIT HANDLER
// ====================================================================
function onFormSubmit(e){
  __mark('onFormSubmit: START');
  try{
    const nv = e && e.namedValues ? e.namedValues : {};

    // ============================================================
    // ✅ DEDUP GUARD: Chặn xử lý trùng trong 60 giây
    // Nguyên nhân: acuityPollAndSubmit ghi vào Form_Inbox
    // đồng thời trigger ON_FORM_SUBMIT cũng fire → xử lý 2 lần
    // ============================================================
    const _calUID = nvGet(nv, 'Admin: Calendly Event UID') || '';
    const _email  = (nv['Email'] || [''])[0] || '';
    const _ts     = String((nv['Timestamp'] || [''])[0] || '').replace(/\W/g, '');

    if (!_calUID && !_email) {
      Logger.log('[dedup] No UID/email - skip dedup guard, proceed normally');
    } else {
      const _dedupKey   = 'formsubmit_' + (_calUID || (_email + '_' + _ts));
      const _dedupCache = CacheService.getScriptCache();

      if (_dedupCache.get(_dedupKey)) {
        Logger.log('[dedup] DUPLICATE DETECTED - skipping. key=' + _dedupKey);
        return; // ← dừng tại đây, không xử lý lần 2
      }
      _dedupCache.put(_dedupKey, '1', 60); // lock 60 giây
      Logger.log('[dedup] key set: ' + _dedupKey);
    }
    // ============================================================

    const submittedAt = (nv['Timestamp']||[''])[0];
    const company     = (nv['Company']||[''])[0];
    const name        = (nv['Customer Name']||[''])[0];
    const phone =
    (nv['Phone']||[''])[0] ||
    (nv['Phone Number']||[''])[0] ||
    (nv["Partner's Phone Number"]||[''])[0] ||
    (nv['Send text messages to']||[''])[0] ||
    (nv['Phone Number']||[''])[0] ||
    (nv['Phone number']||[''])[0] ||
    (nv['Your Phone']||[''])[0] ||
    '';
    const email       = (nv['Email']||[''])[0];
    const vtype       = (nv['Visit Type']||[''])[0];
    const vdate =
    (nv['Visit Date']||[''])[0] ||
    (nv['Event Date']||[''])[0] ||
    (nv['Start Date']||[''])[0] ||
    (nv['Preferred Visit Date']||[''])[0] ||
    '';
    const vtime =
    (nv['Visit Time']||[''])[0] ||
    (nv['Event Time']||[''])[0] ||
    (nv['Start Time']||[''])[0] ||
    (nv['Preferred Visit Time']||[''])[0] ||
    '';
    const location    = (nv['Location']||[''])[0];
    const budgetRaw   = (nv['Budget Range']||[''])[0];
    const sourceRaw   = (nv['Source']||[''])[0];
    const notes       = (nv['Style Notes']||[''])[0];
    // Legacy form/header name. This is the external appointment occurrence UID:
    // Calendly UUID for old bookings, Acuity appointment/synthetic UID for new bookings.
    const calUID      = nvGet(nv, 'Admin: Calendly Event UID');
    const diamondTypeQ = (nv['Diamond Type']||[''])[0];
    
    const diamondTypeNorm = (() => {
      const s = (diamondTypeQ||'').toLowerCase();
      const hasLab = /lab/.test(s), hasNat = /natural/.test(s);
      return hasLab && hasNat ? 'Both' : hasLab ? 'Lab' : hasNat ? 'Natural' : '';
    })();

    const brand       = brandFromCompany_(company);
    const emailLower  = normEmail_(email);
    const phoneNorm   = normPhone_(phone);

    const {min,max} = parseBudget_(budgetRaw);
    const parts = splitName_(name);

    Logger.log(JSON.stringify({
      company, name, emailLower, phoneNorm, vtype, vdate, vtime, location, calUID
    }, null, 2));
    __mark('parsed+normalized fields');

    let row = 0;
    let createdNow = false;

    if (calUID) row = findBestMasterRowByUID_(calUID);

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

    if (!row && !looksLikeReschedule){
      const candRow = findRecentCanceledByContactAt_(emailLower, phoneNorm, 240);
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

    if (row) {
      const existingUID = rf(row, 'CalendlyEventUID') || '';
      if (existingUID && calUID && existingUID !== calUID) {
        row = 0;
      }
    }

    if (!row && !looksLikeReschedule){
      const fpRow = findCurrentMasterRowByFingerprint_(brand, vdate, vtime, vtype, emailLower, phoneNorm, 0);
      if (fpRow) {
        row = fpRow;
        Logger.log('[dedupe] fingerprint matched existing current row=' + row);
      }
    }

    __mark('reschedule detection done; looksLikeReschedule=' + looksLikeReschedule + ', oldRow=' + oldRow + ', preRow=' + row);

    if (!row){
      const tempIso = (vdate && vtime) 
        ? Utilities.formatDate(new Date(`${vdate} ${vtime}`), CFG.TZ, "yyyy-MM-dd'T'HH:mm:ssXXX") 
        : '';
      row = withScriptLock_(() => {
        let existingRow = calUID ? findBestMasterRowByUID_(calUID) : 0;
        if (!existingRow && !looksLikeReschedule) {
          existingRow = findCurrentMasterRowByFingerprint_(brand, vdate, vtime, vtype, emailLower, phoneNorm, 0);
        }
        if (existingRow) {
          Logger.log('[dedupe] lock recheck reused row=' + existingRow);
          return existingRow;
        }

        const visitNo = countVisits_(emailLower, phoneNorm);
        const newRow = appendObj_(SHT.MASTER, {
          'APPT_ID': nextApptId_(tempIso),
          'CalendlyEventUID': calUID || '',
          'Status': 'Scheduled',
          'Active?': 'Yes',
          'Brand': brand || '',
          'Company': company || '',
          'Customer Name': name || '',
          'Email': email || '',
          'EmailLower': emailLower || '',
          'Phone': phone || '',
          'PhoneNorm': phoneNorm || '',
          'Visit Date': vdate || '',
          'Visit Time': vtime || '',
          'Visit Type': vtype || '',
          'Visit #': visitNo,
          'Timestamp': submittedAt || ''
        });
        stampSalesStageIfConsult_(newRow, vtype);
        createdNow = true;
        return newRow;
      });
    }

    const phoneData = resolvePhoneForRow_(phone, row, oldRow, looksLikeReschedule);

    syncBookedAt_(row, submittedAt);

    __mark('before ensureRootAndActiveForNewRow'); 

    (function ensureRootAndActiveForNewRow(){
      try {
        const newAppt = getCell_(SHT.MASTER, row, 'APPT_ID') || '';

        if (looksLikeReschedule && oldRow) {
          const oldRoot = getCell_(SHT.MASTER, oldRow, 'RootApptID') || '';
          const oldAppt = getCell_(SHT.MASTER, oldRow, 'APPT_ID')   || '';
          const root    = oldRoot || oldAppt || '';
          if (root) {
            setOnce_(SHT.MASTER, row,    'RootApptID', root);
            setOnce_(SHT.MASTER, oldRow, 'RootApptID', root);
          }
        } else {
          try {
            const cur = {
              EmailLower: String(emailLower || '').trim().toLowerCase(),
              PhoneNorm:  String(phoneNorm  || '').trim(),
              CalendlyEventUID: String(calUID || '')
            };
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

    if (looksLikeReschedule && oldRow) {
      withScriptLock_(() => {
        if (oldUID && calUID && oldUID !== calUID) {
          const already = (getCell_(SHT.MASTER, oldRow, 'RescheduledToUID') || '');
          if (!already) {
            setCell_(SHT.MASTER, oldRow, 'RescheduledToUID', calUID);
          }
          const newFrom = (getCell_(SHT.MASTER, row, 'RescheduledFromUID') || '');
          if (!newFrom) {
            setCell_(SHT.MASTER, row, 'RescheduledFromUID', oldUID);
          }
        }

        const curSta = getCell_(SHT.MASTER, oldRow, 'Status') || '';
        if (!/rescheduled/i.test(curSta)) {
          setCell_(SHT.MASTER, oldRow, 'Status', 'Rescheduled');
        }

        if (!(getCell_(SHT.MASTER, oldRow, 'CanceledAt'))) {
          try { setCell_(SHT.MASTER, oldRow, 'CanceledAt', new Date()); } catch(_){}
        }

        try {
          setCell_(SHT.MASTER, oldRow, 'Active?', 'No');
        } catch (_noActiveCol) {}

        const prev = getCell_(SHT.MASTER, oldRow, 'Automation Notes') || '';
        setCell_(SHT.MASTER, oldRow, 'Automation Notes',
          (prev? prev+'\n':'') + `Rescheduled → ${calUID} @ ${new Date().toISOString()}`);
      });
    }

    const existingStatusForUpdate = getCell_(SHT.MASTER,row,'Status') || '';
    const existingActiveForUpdate = getCell_(SHT.MASTER,row,'Active?') || '';
    const updates = {
      'Status': existingStatusForUpdate || 'Scheduled',
      'Active?': existingActiveForUpdate || 'Yes',
      'Brand': brand || getCell_(SHT.MASTER,row,'Brand') || '',
      'Company': company || getCell_(SHT.MASTER,row,'Company') || '',
      'Company (normalized)': brand || getCell_(SHT.MASTER,row,'Company (normalized)') || '',
      'Customer Name': name || getCell_(SHT.MASTER,row,'Customer Name') || '',
      'First Name': parts.first || getCell_(SHT.MASTER,row,'First Name') || '',
      'Last Name': parts.last || getCell_(SHT.MASTER,row,'Last Name') || '',
      'Phone': phoneData.raw,
      'PhoneNorm': phoneData.norm,
      'Email': email || getCell_(SHT.MASTER,row,'Email') || '',
      'EmailLower': emailLower || getCell_(SHT.MASTER,row,'EmailLower') || '',
      'Visit Type': vtype || getCell_(SHT.MASTER,row,'Visit Type') || '',
      'Timezone': CFG.TZ,
      'Duration (min)': getCell_(SHT.MASTER,row,'Duration (min)') || DEFAULT_DURATION_MIN,
      'Location': locToEnum_(location) || getCell_(SHT.MASTER,row,'Location') || '',
      'CalendlyEventUID': getCell_(SHT.MASTER,row,'CalendlyEventUID') || calUID || '',
      'Diamond Type': diamondTypeNorm || getCell_(SHT.MASTER,row,'Diamond Type') || '',
      'Budget Range': budgetRaw || getCell_(SHT.MASTER,row,'Budget Range') || '',
      'Budget Min': min || getCell_(SHT.MASTER,row,'Budget Min') || '',
      'Budget Max': max || getCell_(SHT.MASTER,row,'Budget Max') || '',
      'Source': sourceRaw || getCell_(SHT.MASTER,row,'Source') || '',
      'Source (normalized)': normSource_(sourceRaw) || getCell_(SHT.MASTER,row,'Source (normalized)') || '',
      'Style Notes': notes || getCell_(SHT.MASTER,row,'Style Notes') || '',
      'Timestamp': submittedAt || getCell_(SHT.MASTER,row,'Timestamp') || ''
    };

    if (!getCell_(SHT.MASTER,row,'Diamond Type')) {
      const m = /preferred diamond type:\s*([^\n]+)/i.exec(notes || '');
      if (m && m[1]) updates['Diamond Type'] = m[1].trim();
    }

    // ✅ Batch write: ghi tất cả updates trong 1 API call
    (function batchWrite(updates) {
      const s = SH(SHT.MASTER);
      const H = headers_(SHT.MASTER);
      const totalCols = s.getLastColumn();
      const live = s.getRange(row, 1, 1, totalCols).getValues()[0];
      Object.keys(updates).forEach(k => {
        const c = H[k];
        if (c && updates[k] !== '' && updates[k] != null) live[c - 1] = updates[k];
      });
      s.getRange(row, 1, 1, totalCols).setValues([live]);
      SpreadsheetApp.flush();
      Logger.log('[batchWrite] wrote ' + Object.keys(updates).length + ' fields in 1 call');
    })(updates);

    if (vdate && vtime) {
      syncVisitDateTime_(row, vdate, vtime);
    } else {
      Logger.log(`[onFormSubmit] row ${row} - datetime sync skipped (vdate="${vdate}" vtime="${vtime}")`);
    }

    stampSalesStageIfConsult_(row, vtype);

    if (!getCell_(SHT.MASTER, row, 'Visit #')) {
      setCell_(SHT.MASTER, row, 'Visit #', countVisits_(emailLower, phoneNorm));
    }

    (function ensureRootAfterWrites(){
      try {
        const newAppt = getCell_(SHT.MASTER, row, 'APPT_ID') || '';
        const curRoot = getCell_(SHT.MASTER, row, 'RootApptID') || '';

        const curObj = { EmailLower: emailLower, PhoneNorm: phoneNorm, CalendlyEventUID: calUID };
        const prior  = _findMostRecentPriorRow(curObj, row);

        let desired = '';
        if (prior) {
          desired = getCell_(SHT.MASTER, prior.rowIndex, 'RootApptID') ||
                    getCell_(SHT.MASTER, prior.rowIndex, 'APPT_ID') || '';
        }
        if (!desired) desired = newAppt;

        if (desired && desired !== curRoot) {
          setCell_(SHT.MASTER, row, 'RootApptID', desired);
        }
        if (prior && !getCell_(SHT.MASTER, prior.rowIndex, 'RootApptID')) {
          setCell_(SHT.MASTER, prior.rowIndex, 'RootApptID', desired);
        }

        if (prior && prior.rowIndex) {
          try {
            const priorPfUrl = getCell_(SHT.MASTER, prior.rowIndex, 'PaymentsFolderURL');
            if (priorPfUrl) {
              setOnce_(SHT.MASTER, row, 'PaymentsFolderURL', priorPfUrl);
              Logger.log('[ensureRoot] Inherited PaymentsFolderURL from row '
                + prior.rowIndex + ': ' + priorPfUrl);
            }
          } catch(_) {}
        }

      } catch(e){
        err_('ensureRootAfterWrites', e.message, { row });
      }
    })();

    try {
      if (typeof DV_tryEnqueueOnCreate_ === 'function') {
        DV_tryEnqueueOnCreate_({ sh: SH(SHT.MASTER), row: row, dryRun: false });
      }
    } catch (e) {
      Logger.log('DV_onCreate skipped: ' + e.message);
    }

    ensureArtifactsForRow_(row);

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

    _appendNote_(row, `Form merged @ ${new Date().toISOString()}`);

    if (CFG.DEBUG) log_('FORM_MERGED', {row, emailLower, brand});
  }catch(ex){
    err_('onFormSubmit', ex.message, {stack: ex.stack});
    throw ex;
  }
}

function withScriptLock_(fn){
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try { return fn(true); } finally { lock.releaseLock(); }
}

function _atomicWriteUrls_(row, candidateUpdates) {
  if (!candidateUpdates || !Object.keys(candidateUpdates).length) return;

  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(5000);

  if (!gotLock) {
    Logger.log('[atomicWrite] Could not acquire lock - using direct write fallback for row ' + row);
    const H = headers_(SHT.MASTER);
    let written = 0;
    Object.keys(candidateUpdates).forEach(colName => {
      const c   = H[colName];
      const val = candidateUpdates[colName];
      if (!c || !val) return;
      const existing = String(SH(SHT.MASTER).getRange(row, c).getValue() || '').trim();
      if (!existing) {
        SH(SHT.MASTER).getRange(row, c).setValue(val);
        written++;
      }
    });
    SpreadsheetApp.flush();
    Logger.log('[atomicWrite] fallback wrote ' + written + ' fields for row ' + row);
    return;
  }
  try {
    const s    = SH(SHT.MASTER);
    const H    = headers_(SHT.MASTER);
    const totalCols = s.getLastColumn();

    const live = s.getRange(row, 1, 1, totalCols).getValues()[0];

    const skipped = [];
    const written = [];
    let changed = false;

    Object.keys(candidateUpdates).forEach(colName => {
      const c   = H[colName];
      const val = candidateUpdates[colName];
      if (!c || !val) return;

      const existing = String(live[c-1] || '').trim();
      if (existing) {
        skipped.push(colName);
        return;
      }
      live[c-1] = val;
      written.push(colName);
      changed = true;
    });

    if (changed) {
      s.getRange(row, 1, 1, live.length).setValues([live]);
      SpreadsheetApp.flush();
    }

    if (written.length) Logger.log(`[atomicWrite] row=${row} wrote: ${written.join(', ')}`);
    if (skipped.length) Logger.log(`[atomicWrite] row=${row} skipped (already set): ${skipped.join(', ')}`);

  } finally {
    lock.releaseLock();
  }
}

/********** UTIL **********/
function ping_(){ return Utilities.formatDate(new Date(), CFG.TZ, "yyyy-MM-dd HH:mm:ss"); }

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

  Logger.log('_findMostRecentPriorRow.length = ' + _findMostRecentPriorRow.length);
  Logger.log('_currentRowToObj_.length = ' + _currentRowToObj_.length);

  const s = SH(SHT.FORM_INBOX), m = headers_(SHT.FORM_INBOX);
  const r = s.getLastRow(); if (r < 2) { Logger.log('No inbox rows'); return; }

  const vals = s.getRange(r,1,1,s.getLastColumn()).getValues()[0];
  const nv = {};
  Object.keys(m).forEach(h => nv[h] = [ vals[m[h]-1] ]);

  Logger.log('Inbox NV keys: ' + Object.keys(nv).join(', '));

  __mark('calling onFormSubmit');
  onFormSubmit({ namedValues: nv });
  __mark('onFormSubmit returned');
}

// --- Chat helpers ---
function chatWebhook_(){
  const url = PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_ALL');
  if (!url) throw new Error('Missing CHAT_WEBHOOK_ALL script property');
  return url;
}

function _redactWebhook_(url){
  if (!url) return '(missing)';
  try { return '…' + url.slice(-12); } catch(_) { return '(unprintable)'; }
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

function buildIntakeCreatedCard_(rowIdx){
  const s = SpreadsheetApp.getActive().getSheetByName('00_Master Appointments');
  const H = s.getRange(1,1,1,s.getLastColumn()).getValues()[0]
           .reduce((m,h,i)=> (m[h]=i+1,m), {});
  const S = (v) => v == null ? '' : String(v);
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
      card: { header: { title }, sections: [{ widgets }] }
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
      card: { header: { title }, sections: [{ widgets }] }
    }]
  };
}

function postRescheduledCard_(oldRowIdx, newRowIdx){
  const payload = buildRescheduledCard_(oldRowIdx, newRowIdx);
  return _postChatPayload_('rescheduled', payload, { oldRowIdx, newRowIdx });
}

function postIntakeCreatedCard_(rowIdx){
  const payload = buildIntakeCreatedCard_(rowIdx);
  return _postChatPayload_('created', payload, { rowIdx });
}

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

  Logger.log(`[CHAT] payload head: ${json.substring(0, 600)}${sizeB>600?' …[trunc]':''}`);

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: json,
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    const body = (res.getContentText() || '').substr(0, 1000);
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

  ensureArtifactsForRow_(r);

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

  const masterName = SHT && SHT.MASTER ? SHT.MASTER : '(SHT.MASTER not set)';
  const sh = ss.getSheetByName(masterName);

  if (!sh) {
    ui.alert('Resolver: Master sheet not found',
      'SHT.MASTER = "' + masterName + '" but no such tab exists.\n' +
      'Available tabs: ' + ss.getSheets().map(s => s.getName()).join(' | '),
      ui.ButtonSet.OK);
    return;
  }

  const H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]
            .reduce((m,h,i)=> (h && (m[String(h).trim()]=i+1), m), {});
  const need = ['APPT_ID','RootApptID','Brand','Customer Name','EmailLower','PhoneNorm',
                'CalendlyEventUID','Visit Type','ApptDateTime (ISO)','Client Folder',
                'ProspectFolderID','IntakeDocURL','Checklist URL','Quotation URL'];
  const missing = need.filter(h => !H[h]);

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

  if (!H['Client Folder']) {
    s.getRange(1, s.getLastColumn()+1).setValue('Client Folder');
    H = headers_(SHT.MASTER);
  }

  const colClient = H['Client Folder'];
  const colOld    = H['Folder URL'];

  const last = s.getLastRow();
  if (last >= 2 && colOld) {
    const oldVals = s.getRange(2, colOld, last-1, 1).getValues();
    const newVals = s.getRange(2, colClient, last-1, 1).getValues();

    for (let i=0;i<oldVals.length;i++){
      const src = String(oldVals[i][0]||'').trim();
      const dst = String(newVals[i][0]||'').trim();
      if (src && !dst) newVals[i][0] = src;
    }
    s.getRange(2, colClient, last-1, 1).setValues(newVals);
    s.deleteColumn(colOld);
  }

  SpreadsheetApp.getUi().alert('Migration complete: Client Folder set; "Folder URL" removed.');
}

function debug_bootstrapForLastRealRow() {
  const sh = _openMaster_();
  const last = sh.getLastRow();
  if (last < 2) { Logger.log('No data rows'); return; }

  const H  = _headers_(sh);
  const colApId = H['RootApptID'];
  if (!colApId) throw new Error('Missing "RootApptID" column');

  const vals = sh.getRange(2, colApId, last-1, 1).getValues();
  let lastRowWithAp = -1;
  for (let i = vals.length - 1; i >= 0; i--) {
    const v = String(vals[i][0] || '').trim();
    if (v) { lastRowWithAp = i + 2; break; }
  }
  if (lastRowWithAp === -1) { Logger.log('No RootApptID found'); return; }

  const id = bootstrapApFolderForRow_(lastRowWithAp);
  Logger.log('Bootstrapped AP folder ID: ' + id + ' for row ' + lastRowWithAp);
}

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
        const id = bootstrapApFolderForRow_(row);
        Logger.log(`Row ${row}: created RootAppt folder ${id}`);
        fixed++;
      } catch(e) {
        Logger.log(`Row ${row}: bootstrap error: ${e && e.message}`);
      }
    }
  }
  Logger.log(`backfillAllRootApptFolders: bootstrapped ${fixed} missing rows`);
}

// ====================================================================
// FIX #4: URL REPAIR WORKER
// ====================================================================
function repairMissingUrls_() {
  const REPAIR_LOOKBACK_ROWS = 50;
  const REPAIR_WINDOW_HOURS  = 48;

  Logger.log('===== REPAIR START =====');

  const s    = SH(SHT.MASTER);
  const H    = headers_(SHT.MASTER);
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);

  Logger.log('Last data row = ' + last);

  if (last < 2) { Logger.log('No data rows'); return; }

  const cAppt    = H['APPT_ID']       || 0;
  const cBrand   = H['Brand']         || 0;
  const cIntake  = H['IntakeDocURL']  || 0;
  const cChk     = H['Checklist URL'] || 0;
  const cQuo     = H['Quotation URL'] || 0;
  const cTs      = H['Timestamp']     || 0;

  if (!cAppt || !cBrand) { Logger.log('Missing APPT_ID or Brand column → abort'); return; }

  const startRow = Math.max(2, last - REPAIR_LOOKBACK_ROWS + 1);
  const numRows  = last - startRow + 1;

  Logger.log('Scanning rows from ' + startRow + ' → ' + last);

  const colsToRead = [cAppt, cBrand, cIntake, cChk, cQuo, cTs].filter(Boolean);
  const maxCol = Math.max(...colsToRead);
  const block  = s.getRange(startRow, 1, numRows, maxCol).getValues();

  const now = Date.now();
  const windowMs = REPAIR_WINDOW_HOURS * 3600 * 1000;

  let repaired = 0;

  for (let i = 0; i < block.length; i++) {
    const sheetRow = startRow + i;
    const r = block[i];

    const appt  = cAppt  ? String(r[cAppt-1]  || '').trim() : '';
    const brand = cBrand ? String(r[cBrand-1] || '').trim() : '';
    if (!appt || !brand) continue;

    if (cTs) {
      const rawTs = r[cTs-1];
      const ts = rawTs ? new Date(rawTs).getTime() : 0;
      if (ts && (now - ts) > windowMs) continue;
    }

    const intakeVal = cIntake ? String(r[cIntake-1] || '').trim() : '';
    const chkVal    = cChk    ? String(r[cChk-1]    || '').trim() : '';
    const quoVal    = cQuo    ? String(r[cQuo-1]    || '').trim() : '';

    if (intakeVal && chkVal && quoVal) continue;

    // ✅ Fix: kiểm tra repair lock riêng để tránh chạy lại liên tục
    const repairLockKey = `repair_lock_${sheetRow}`;
    const repairCache = CacheService.getScriptCache();
    if (repairCache.get(repairLockKey)) {
      Logger.log('[row ' + sheetRow + '] repair already attempted recently - skipping');
      continue;
    }
    repairCache.put(repairLockKey, '1', 300); // ✅ không retry trong 5 phút

    Logger.log('[row ' + sheetRow + '] MISSING URL → healing');

    try {
      const repairLock = LockService.getUserLock();
      const gotRepairLock = repairLock.tryLock(3000);
      if (gotRepairLock) {
        try { ensureArtifactsForRow_(sheetRow); }
        finally { repairLock.releaseLock(); }
      } else {
        Logger.log('[repair] row ' + sheetRow + ' lock busy - will retry next run');
      }

      // ✅ Re-read từ sheet sau khi chạy xong (không dùng cache cũ)
      const afterIntake = getCell_(SHT.MASTER, sheetRow, 'IntakeDocURL');
      const afterChk    = getCell_(SHT.MASTER, sheetRow, 'Checklist URL');
      const afterQuo    = getCell_(SHT.MASTER, sheetRow, 'Quotation URL');

      Logger.log('[row ' + sheetRow + '] AFTER repair: intake=' + !!afterIntake + ' chk=' + !!afterChk + ' quo=' + !!afterQuo);

      // ✅ Chỉ tính là repaired nếu thực sự có URL sau khi chạy
      if (afterIntake || afterChk || afterQuo) repaired++;

    } catch (e) {
      Logger.log('[row ' + sheetRow + '] ERROR: ' + (e && e.message));
    }
  }

  Logger.log('===== REPAIR DONE | repaired ' + repaired + ' row(s) =====');
}

function installUrlRepairWorker() {
  const FN = 'repairMissingUrls_';
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === FN);
  if (!exists) {
    ScriptApp.newTrigger(FN).timeBased().everyMinutes(1).create();
    Logger.log('[repairWorker] trigger installed: every 1 minute');
  } else {
    Logger.log('[repairWorker] already installed');
  }
}

function uninstallUrlRepairWorker() {
  const FN = 'repairMissingUrls_';
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === FN)
    .forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('[repairWorker] trigger removed');
}

function TEST_REPAIR() { repairMissingUrls_(); }

// ====================================================================
// BACKFILL: Fix existing rows that have duplicate Prospect Folders
// ====================================================================
/**
 * Scans all Master rows and consolidates ProspectFolderIDs for each
 * unique customer (email or phone) to use the EARLIEST one.
 *
 * Safe to run multiple times (idempotent).
 * Run once after deploying FIX #5 to clean up historical duplicates.
 */
function backfillConsolidateProspectFolders_() {
  const s    = SH(SHT.MASTER);
  const H    = headers_(SHT.MASTER);
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);

  if (last < 2) { Logger.log('No data rows'); return; }

  const cEmail = H['EmailLower'] || 0;
  const cPhone = H['PhoneNorm']  || 0;
  const cPfId  = H['ProspectFolderID'] || 0;
  const cAppt  = H['APPT_ID'] || 0;

  if (!cPfId) { Logger.log('Missing ProspectFolderID column'); return; }

  const numRows = last - 1;
  const emailVec = cEmail ? s.getRange(2, cEmail, numRows, 1).getValues().map(r => String(r[0]||'').toLowerCase().trim()) : Array(numRows).fill('');
  const phoneVec = cPhone ? s.getRange(2, cPhone, numRows, 1).getValues().map(r => String(r[0]||'').trim())              : Array(numRows).fill('');
  const pfIdVec  = s.getRange(2, cPfId, numRows, 1).getValues().map(r => String(r[0]||'').trim());
  const apptVec  = cAppt  ? s.getRange(2, cAppt,  numRows, 1).getValues().map(r => String(r[0]||'').trim())              : Array(numRows).fill('');

  // Build map: contactKey → earliest ProspectFolderID
  const canonPfId = {};  // contactKey → pfId
  for (let i = 0; i < numRows; i++) {
    const email = emailVec[i];
    const phone = phoneVec[i];
    const pfId  = pfIdVec[i];
    if (!pfId) continue;

    const keys = [];
    if (email) keys.push('e:' + email);
    if (phone) keys.push('p:' + phone);

    keys.forEach(k => {
      if (!canonPfId[k]) canonPfId[k] = pfId; // first one wins (earliest row)
    });
  }

  // Second pass: stamp canonical pfId onto rows that have a different (or empty) pfId
  let fixed = 0;
  for (let i = 0; i < numRows; i++) {
    const row   = i + 2;
    const email = emailVec[i];
    const phone = phoneVec[i];
    const pfId  = pfIdVec[i];

    let canonical = '';
    if (email && canonPfId['e:' + email]) canonical = canonPfId['e:' + email];
    else if (phone && canonPfId['p:' + phone]) canonical = canonPfId['p:' + phone];

    if (canonical && canonical !== pfId) {
      try {
        // Verify the canonical folder still exists
        DriveApp.getFolderById(canonical);
        setCell_(SHT.MASTER, row, 'ProspectFolderID', canonical);
        Logger.log(`Row ${row} (${apptVec[i]}): ProspectFolderID updated ${pfId || '(empty)'} → ${canonical}`);
        fixed++;
      } catch (_) {
        Logger.log(`Row ${row}: canonical pfId ${canonical} unreachable — skipped`);
      }
    }
  }

  Logger.log(`backfillConsolidateProspectFolders_: fixed ${fixed} row(s)`);
}

// DEBUG/TEST FUNCTIONS
function debug_testSyncDateTime() {
  const row = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  if (row < 2) { Logger.log('No data rows'); return; }
  
  Logger.log('=== BEFORE ===');
  Logger.log('ISO: ' + getCell_(SHT.MASTER, row, 'ApptDateTime (ISO)'));
  Logger.log('Visit Date: ' + getCell_(SHT.MASTER, row, 'Visit Date'));
  Logger.log('Visit Time: ' + getCell_(SHT.MASTER, row, 'Visit Time'));
  
  syncVisitDateTime_(row, '12/25/2025', '2:30:00 PM');
  
  Logger.log('=== AFTER ===');
  Logger.log('ISO: ' + getCell_(SHT.MASTER, row, 'ApptDateTime (ISO)'));
  Logger.log('Visit Date: ' + getCell_(SHT.MASTER, row, 'Visit Date'));
  Logger.log('Visit Time: ' + getCell_(SHT.MASTER, row, 'Visit Time'));
}

function debug_compareBuildIntakeData() {
  const row = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  if (row < 2) { Logger.log('No data rows'); return; }
  
  Logger.log('=== buildIntakeData_ (ISO-first logic) ===');
  const data = buildIntakeData_(row);
  Logger.log(JSON.stringify({
    ApptDate: data.ApptDate,
    ApptTime: data.ApptTime,
    ApptDateTime: data.ApptDateTime
  }, null, 2));
}

// Legacy shims
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

// ====================================================================
// BACKUP RESOLVER
// ====================================================================

function backupResolveInbox_() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const inboxSh = ss.getSheetByName(SHT.FORM_INBOX);

  if (!inboxSh) {
    ui.alert('❌ Sheet không tìm thấy', `Sheet "${SHT.FORM_INBOX}" không tồn tại.`, ui.ButtonSet.OK);
    return;
  }

  const last = inboxSh.getLastRow();
  if (last < 2) {
    ui.alert('ℹ️ Không có dữ liệu', 'Sheet 02_Form_Inbox trống hoặc chỉ có header.', ui.ButtonSet.OK);
    return;
  }

  const H = headers_(SHT.FORM_INBOX);
  const totalCols = inboxSh.getLastColumn();
  const allRows = inboxSh.getRange(2, 1, last - 1, totalCols).getValues();
  const headers = inboxSh.getRange(1, 1, 1, totalCols).getValues()[0].map(h => String(h).trim());

  const unresolved = [];
  const masterSh   = SH(SHT.MASTER);
  const MH         = headers_(SHT.MASTER);
  const masterLast = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);

  let masterUIDs = [];
  if (MH['CalendlyEventUID'] && masterLast >= 2) {
    masterUIDs = masterSh
      .getRange(2, MH['CalendlyEventUID'], masterLast - 1, 1)
      .getValues().flat().map(v => String(v || '').trim()).filter(Boolean);
  }

  allRows.forEach((row, i) => {
    const inboxRow = i + 2;
    const nv = {};
    headers.forEach((h, ci) => { nv[h] = [row[ci]]; });

    const calUID = nvGet(nv, 'Admin: Calendly Event UID') || '';
    const emailRaw = (nv['Email'] || [''])[0];
    const emailLower = normEmail_(emailRaw);

    if (calUID && masterUIDs.includes(calUID)) return;

    if (!calUID && emailLower) {
      const vdate = (nv['Visit Date'] || nv['Event Date'] || [''])[0];
      const foundInMaster = findMasterRowByEmailTime_(emailLower, vdate);
      if (foundInMaster) return;
    }

    unresolved.push({ inboxRow, nv, calUID, emailLower });
  });

  if (!unresolved.length) {
    ui.alert('✅ Đã đồng bộ', 'Tất cả rows trong Form_Inbox đã được resolve lên Master.', ui.ButtonSet.OK);
    return;
  }

  const lines = unresolved.slice(0, 20).map(u => {
    const name  = (u.nv['Customer Name'] || [''])[0] || '(no name)';
    const vdate = (u.nv['Visit Date'] || u.nv['Event Date'] || [''])[0] || '';
    const uid   = u.calUID ? ` [${u.calUID.slice(0,12)}…]` : '';
    return `Row ${u.inboxRow}: ${name} | ${vdate}${uid}`;
  }).join('\n');

  const extraNote = unresolved.length > 20 ? `\n…và ${unresolved.length - 20} rows khác.` : '';

  const answer = ui.alert(
    `📋 ${unresolved.length} rows chưa resolve`,
    `Sẽ đẩy các rows sau lên Master:\n\n${lines}${extraNote}\n\nTiếp tục?`,
    ui.ButtonSet.YES_NO
  );

  if (answer !== ui.Button.YES) { Logger.log('[backupResolve] User cancelled.'); return; }

  let success = 0, failed = 0;
  const errors = [];

  unresolved.forEach(u => {
    try {
      onFormSubmit({ namedValues: u.nv });
      _markInboxRowProcessed_(inboxSh, H, u.inboxRow);
      success++;
      Logger.log(`[backupResolve] ✅ row ${u.inboxRow} → Master OK`);
    } catch (e) {
      failed++;
      errors.push(`Row ${u.inboxRow}: ${e.message}`);
      Logger.log(`[backupResolve] ❌ row ${u.inboxRow} ERROR: ${e.message}`);
      err_('backupResolveInbox_', e.message, { inboxRow: u.inboxRow, email: u.emailLower });
    }
  });

  const summary = `✅ Thành công: ${success}\n❌ Lỗi: ${failed}` +
    (errors.length ? '\n\nChi tiết lỗi:\n' + errors.slice(0, 5).join('\n') : '');

  ui.alert('Kết quả Backup Resolve', summary, ui.ButtonSet.OK);
}

function _markInboxRowProcessed_(sh, H, rowIdx) {
  let col = H['Resolved At'];
  if (!col) {
    const newCol = sh.getLastColumn() + 1;
    sh.getRange(1, newCol).setValue('Resolved At');
    H['Resolved At'] = newCol;
    col = newCol;
  }
  sh.getRange(rowIdx, col).setValue(new Date());
}

function backupResolveInboxRows_(fromRow, toRow) {
  const inboxSh = SH(SHT.FORM_INBOX);
  const H = headers_(SHT.FORM_INBOX);
  const totalCols = inboxSh.getLastColumn();
  const headers = inboxSh.getRange(1, 1, 1, totalCols).getValues()[0].map(h => String(h).trim());

  const from = Math.max(2, fromRow || 2);
  const to   = Math.min(inboxSh.getLastRow(), toRow || inboxSh.getLastRow());

  if (from > to) { Logger.log('[backupResolveRows] Invalid range: from=' + from + ' to=' + to); return; }

  Logger.log(`[backupResolveRows] Processing inbox rows ${from}–${to}`);

  const block = inboxSh.getRange(from, 1, to - from + 1, totalCols).getValues();
  let success = 0, failed = 0;

  block.forEach((row, i) => {
    const inboxRow = from + i;
    const nv = {};
    headers.forEach((h, ci) => { nv[h] = [row[ci]]; });

    try {
      onFormSubmit({ namedValues: nv });
      _markInboxRowProcessed_(inboxSh, H, inboxRow);
      success++;
      Logger.log(`  ✅ Row ${inboxRow} OK`);
    } catch (e) {
      failed++;
      Logger.log(`  ❌ Row ${inboxRow} ERROR: ${e.message}`);
      err_('backupResolveInboxRows_', e.message, { inboxRow });
    }
  });

  Logger.log(`[backupResolveRows] Done. ✅${success} ❌${failed}`);
}

function backupResolveInbox_DryRun() {
  const inboxSh = SH(SHT.FORM_INBOX);
  const last = inboxSh.getLastRow();
  if (last < 2) { Logger.log('Form_Inbox is empty'); return; }

  const H = headers_(SHT.FORM_INBOX);
  const totalCols = inboxSh.getLastColumn();
  const headers = inboxSh.getRange(1, 1, 1, totalCols).getValues()[0].map(h => String(h).trim());
  const allRows = inboxSh.getRange(2, 1, last - 1, totalCols).getValues();

  const masterSh   = SH(SHT.MASTER);
  const MH         = headers_(SHT.MASTER);
  const masterLast = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  let masterUIDs = [];
  if (MH['CalendlyEventUID'] && masterLast >= 2) {
    masterUIDs = masterSh
      .getRange(2, MH['CalendlyEventUID'], masterLast - 1, 1)
      .getValues().flat().map(v => String(v || '').trim()).filter(Boolean);
  }

  let unresolvedCount = 0;
  allRows.forEach((row, i) => {
    const inboxRow = i + 2;
    const nv = {};
    headers.forEach((h, ci) => { nv[h] = [row[ci]]; });

    const calUID     = nvGet(nv, 'Admin: Calendly Event UID') || '';
    const emailRaw   = (nv['Email'] || [''])[0];
    const emailLower = normEmail_(emailRaw);
    const name       = (nv['Customer Name'] || [''])[0] || '';
    const vdate      = (nv['Visit Date'] || nv['Event Date'] || [''])[0] || '';

    if (calUID && masterUIDs.includes(calUID)) {
      Logger.log(`Row ${inboxRow}: ✅ RESOLVED  ${name} | ${calUID}`);
      return;
    }
    if (!calUID && emailLower && findMasterRowByEmailTime_(emailLower, vdate)) {
      Logger.log(`Row ${inboxRow}: ✅ RESOLVED  ${name} | ${emailLower}`);
      return;
    }

    Logger.log(`Row ${inboxRow}: ❌ UNRESOLVED  ${name} | ${emailLower} | UID=${calUID || '(none)'} | ${vdate}`);
    unresolvedCount++;
  });

  Logger.log(`\n=== DRY RUN SUMMARY: ${unresolvedCount} unresolved / ${allRows.length} total ===`);
}

function addBackupResolverMenu() {
  SpreadsheetApp.getActive().addMenu('🔁 Backup Resolver', [
    { name: '▶ Resolve rows chưa xử lý (Auto detect)',    functionName: 'backupResolveInbox_'                  },
    { name: '🔍 Dry Run – Chỉ xem không push',            functionName: 'backupResolveInbox_DryRun'             },
    { name: '─────────────────────────────',              functionName: 'addBackupResolverMenu'                 },
    { name: '🔧 Backfill: Fix RootApptID',                functionName: 'backfillConsolidateRootApptIDs_'       },
    { name: '🔧 Backfill: Fix ProspectFolderID',          functionName: 'backfillConsolidateProspectFolders_'   },
    { name: '📊 Diag: Customer Chain Report',             functionName: 'diagCustomerChain_'                    },
  ]);
}

function resolver_onOpen_() {
  addBackupResolverMenu();
}

function runBackupRows() {
  backupResolveInboxRows_(441, 442);
}

function diagProtections() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  
  const sheetProtections = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  Logger.log('=== SHEET PROTECTIONS ===');
  sheetProtections.forEach((p, i) => {
    Logger.log(`[${i}] Description: ${p.getDescription()}`);
    Logger.log(`    Editors: ${p.getEditors().map(u => u.getEmail()).join(', ')}`);
    Logger.log(`    Unprotected ranges: ${p.getUnprotectedRanges().map(r => r.getA1Notation()).join(', ')}`);
  });

  const rangeProtections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  Logger.log('=== RANGE PROTECTIONS ===');
  rangeProtections.forEach((p, i) => {
    Logger.log(`[${i}] Range: ${p.getRange().getA1Notation()}`);
    Logger.log(`    Description: ${p.getDescription()}`);
    Logger.log(`    Editors: ${p.getEditors().map(u => u.getEmail()).join(', ')}`);
  });

  if (!sheetProtections.length && !rangeProtections.length) {
    Logger.log('Không có protection nào trên sheet này.');
  }
}

// ====================================================================
// FIX #6: ROOTAPPTID CHAIN CONSOLIDATION
// ====================================================================
/**
 * Backfill: gộp tất cả appointments của cùng 1 khách về dùng chung
 * RootApptID = APPT_ID của appointment ĐẦU TIÊN (earliest row).
 *
 * Root cause được fix:
 * - _findMostRecentPriorRow() cũ chỉ trả về prior row khi đã có artifacts
 * - Nếu APPT_001 chưa có IntakeURL → APPT_002 tự set RootApptID = APPT_002
 * - Dẫn đến broken chain: APPT_001/APPT_002/APPT_003 có 3 RootApptID khác nhau
 *
 * Safe to re-run (idempotent).
 */
function backfillConsolidateRootApptIDs_() {
  const s    = SH(SHT.MASTER);
  const H    = headers_(SHT.MASTER);
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);

  if (last < 2) { Logger.log('No data rows'); return; }

  const cEmail = H['EmailLower']  || 0;
  const cPhone = H['PhoneNorm']   || 0;
  const cAppt  = H['APPT_ID']     || 0;
  const cRoot  = H['RootApptID']  || 0;

  if (!cAppt || !cRoot) { Logger.log('Missing APPT_ID or RootApptID column'); return; }

  const n = last - 1;
  const emailVec = cEmail ? s.getRange(2,cEmail,n,1).getValues().map(r=>String(r[0]||'').toLowerCase().trim()) : Array(n).fill('');
  const phoneVec = cPhone ? s.getRange(2,cPhone,n,1).getValues().map(r=>String(r[0]||'').trim())              : Array(n).fill('');
  const apptVec  =          s.getRange(2,cAppt, n,1).getValues().map(r=>String(r[0]||'').trim());
  const rootVec  =          s.getRange(2,cRoot, n,1).getValues().map(r=>String(r[0]||'').trim());

  // Pass 1: tìm APPT_ID earliest (row thấp nhất) cho mỗi contact key
  const canonRoot = {};
  for (let i = 0; i < n; i++) {
    const appt  = apptVec[i];
    if (!appt) continue;
    const email = emailVec[i];
    const phone = phoneVec[i];
    const keys  = [];
    if (email) keys.push('e:' + email);
    if (phone) keys.push('p:' + phone);
    keys.forEach(k => {
      if (!canonRoot[k]) canonRoot[k] = appt; // first (earliest) wins
    });
  }

  // Pass 2: stamp canonical RootApptID lên rows sai hoặc trống
  let fixed = 0;
  for (let i = 0; i < n; i++) {
    const row  = i + 2;
    const appt = apptVec[i];
    if (!appt) continue;

    const email = emailVec[i];
    const phone = phoneVec[i];

    let canonical = '';
    if (email && canonRoot['e:' + email]) canonical = canonRoot['e:' + email];
    else if (phone && canonRoot['p:' + phone]) canonical = canonRoot['p:' + phone];
    if (!canonical) canonical = appt;

    const current = rootVec[i];
    if (current !== canonical) {
      setCell_(SHT.MASTER, row, 'RootApptID', canonical);
      Logger.log(`Row ${row} (${appt}): RootApptID  "${current || '(empty)'}"  →  "${canonical}"`);
      fixed++;
    }
  }

  Logger.log(`\nbackfillConsolidateRootApptIDs_: fixed ${fixed} row(s)`);
}


/**
 * Diagnostic: in customer chain report để verify
 * RootApptID + ProspectFolderID + Quotation URL đồng nhất cho repeat customers.
 *
 * Chạy TRƯỚC và SAU backfill để so sánh.
 */
function diagCustomerChain_() {
  const s    = SH(SHT.MASTER);
  const H    = headers_(SHT.MASTER);
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  if (last < 2) { Logger.log('No data rows'); return; }

  const cEmail = H['EmailLower']       || 0;
  const cPhone = H['PhoneNorm']        || 0;
  const cAppt  = H['APPT_ID']          || 0;
  const cRoot  = H['RootApptID']       || 0;
  const cPfId  = H['ProspectFolderID'] || 0;
  const cQuo   = H['Quotation URL']    || 0;
  const cName  = H['Customer Name']    || 0;
  const cSta   = H['Status']           || 0;

  const n = last - 1;
  const rows = s.getRange(2, 1, n, s.getLastColumn()).getValues();

  // Group by contact key
  const map = {};
  rows.forEach((r, i) => {
    const email = cEmail ? String(r[cEmail-1]||'').toLowerCase().trim() : '';
    const phone = cPhone ? String(r[cPhone-1]||'').trim() : '';
    const appt  = cAppt  ? String(r[cAppt-1] ||'').trim() : '';
    const root  = cRoot  ? String(r[cRoot-1] ||'').trim() : '';
    const pfId  = cPfId  ? String(r[cPfId-1] ||'').trim() : '';
    const quo   = cQuo   ? String(r[cQuo-1]  ||'').trim() : '';
    const name  = cName  ? String(r[cName-1] ||'').trim() : '';
    const sta   = cSta   ? String(r[cSta-1]  ||'').trim() : '';
    const key   = email || phone;
    if (!key || !appt) return;
    if (!map[key]) map[key] = [];
    map[key].push({ row: i+2, appt, root, pfId, quo, name, sta });
  });

  let totalRepeat = 0, issues = 0;

  Object.entries(map).forEach(([key, entries]) => {
    if (entries.length < 2) return;
    totalRepeat++;

    const name  = entries[0].name;
    const roots = [...new Set(entries.map(e => e.root).filter(Boolean))];
    const pfIds = [...new Set(entries.map(e => e.pfId).filter(Boolean))];
    const quos  = [...new Set(entries.map(e => e.quo).filter(Boolean))];

    const rootOk = roots.length <= 1;
    const pfOk   = pfIds.length <= 1;
    const quoOk  = quos.length <= 1;
    const allOk  = rootOk && pfOk && quoOk;

    if (!allOk) issues++;

    Logger.log(`\n${allOk ? '✅' : '❌'} 👤 ${name} (${key}) — ${entries.length} appts`);
    entries.forEach(e => {
      Logger.log(`   Row ${String(e.row).padStart(3)}: ${e.appt.padEnd(18)} root=${e.root||'?'}`);
    });
    if (!rootOk) Logger.log(`   ❌ RootApptID: ${roots.join(' / ')}`);
    if (!pfOk)   Logger.log(`   ❌ ProspectFolder: ${pfIds.length} khác nhau`);
    if (!quoOk)  Logger.log(`   ❌ Quotation URL: ${quos.length} khác nhau`);
  });

  Logger.log(`\n${'='.repeat(55)}`);
  Logger.log(`SUMMARY: ${totalRepeat - issues}/${totalRepeat} repeat customers OK`);
  if (issues > 0) Logger.log(`⚠️  ${issues} customer(s) cần chạy backfill`);
  else Logger.log(`🎉 Tất cả đều nhất quán — không cần fix gì thêm`);
  Logger.log('='.repeat(55));
}

// ====================================================================
// END OF RESOLVER.GS - FIXED VERSION (FIX #1 → #6)
// ====================================================================


// ============================================================
//  BƯỚC 1: Chạy function này để lấy Doc ID của 2 template mới
//  Sau đó copy ID vào BƯỚC 2 bên dưới
// ============================================================

function findNewTemplateIds() {
  // ← Điền đúng tên file Google Doc của bạn vào đây
  const HPUSA_FILE_NAME = 'Custom Jewelry Design';        // tên file HPUSA
  const VVS_FILE_NAME   = 'Custom Engagement Ring Design'; // tên file VVS

  Logger.log('===== TÌM TEMPLATE IDs =====');

  [
    { brand: 'HPUSA', name: HPUSA_FILE_NAME },
    { brand: 'VVS',   name: VVS_FILE_NAME   }
  ].forEach(({ brand, name }) => {
    const results = DriveApp.getFilesByName(name);
    let count = 0;

    while (results.hasNext()) {
      const file = results.next();
      count++;
      Logger.log(`\n[${brand}] Tìm thấy:`);
      Logger.log(`  Tên  : ${file.getName()}`);
      Logger.log(`  ID   : ${file.getId()}`);
      Logger.log(`  Link : ${file.getUrl()}`);
    }

    if (count === 0) Logger.log(`[${brand}] Không tìm thấy file tên "${name}"`);
    if (count > 1)   Logger.log(`[${brand}] ⚠️  Có ${count} file trùng tên — chọn đúng ID nhé`);
  });
}


// ============================================================
//  BƯỚC 2: Paste 2 ID vào đây rồi chạy function này
// ============================================================

// ============================================================
//  PROJECT #17 – AUTO GOOGLE SLIDES GENERATION (simplified)
//  Chỉ tạo file trong folder, không ghi URL vào sheet
// ============================================================


// ── 1. LẤY TEMPLATE ID THEO BRAND ───────────────────────────

function slidesTemplateIdForBrand_(brand) {
  const SP = PropertiesService.getScriptProperties();
  if (brand === 'HPUSA') return SP.getProperty('SLIDES_TEMPLATE_ID_HPUSA') || '';
  if (brand === 'VVS')   return SP.getProperty('SLIDES_TEMPLATE_ID_VVS')   || '';
  return '';
}


// ── 2. CORE FUNCTION ─────────────────────────────────────────

function generateSlidesForRow_(folder, data) {
  const brand  = String(data.Brand  || '').trim().toUpperCase();
  const apptId = String(data.ApptId || '').trim();  // ← phải có giá trị
  const name   = String(data.CustomerName || '').trim();

  if (!brand || !apptId) {
    Logger.log('[Slides] Abort: missing brand="' + brand + '" apptId="' + apptId + '"');
    return;
  }

  const tplId = slidesTemplateIdForBrand_(brand);
  if (!tplId) {
    Logger.log('[Slides] Không có template cho brand "' + brand + '"');
    return;
  }

  const fileName = brand + ' \u2013 ' + apptId + ' \u2013 Slides';

  const existing = folder.getFilesByName(fileName);
  if (existing.hasNext()) {
    Logger.log('[Slides] Đã tồn tại: ' + fileName);
    return;
  }

  const copy = DriveApp.getFileById(tplId).makeCopy(fileName, folder);
  const pres = SlidesApp.openById(copy.getId());

  if (name) _insertClientName_(pres.getSlides()[0], name);
  _ensureTenBlankSlides_(pres);

  pres.saveAndClose();
  Logger.log('[Slides] ✅ Đã tạo: ' + fileName);
}


// ── 3. INSERT TÊN KH VÀO WELCOME SLIDE ──────────────────────

function _insertClientName_(welcomeSlide, customerName) {
  if (!welcomeSlide) return;
  let replaced = false;

  welcomeSlide.getShapes().forEach(shape => {
    if (!shape.getText) return;
    const tf  = shape.getText();
    const txt = tf.asString();
    if (txt.includes('{{CustomerName}}')) {
      tf.replaceAllText('{{CustomerName}}', customerName);
      replaced = true;
    }
  });

  if (!replaced) {
    Logger.log('[Slides] ⚠️  Không tìm thấy {{CustomerName}} trên welcome slide');
  }
}

// ====================================================================
// PDF GENERATION - Auto export Intake Doc → PDF
// ====================================================================

/**
 * Export Google Doc sang PDF và lưu vào cùng folder với Intake Doc
 * Returns: URL của file PDF hoặc '' nếu thất bại
 */
function exportIntakeDocToPdf_(docId, destFolder, brand, apptId) {
  if (!docId || !destFolder) return '';

  const pdfName = brand + ' \u2013 ' + apptId + ' \u2013 Intake';

  // ── Xóa PDF cũ cùng tên nếu có ──────────────────────────────
  try {
    const oldIt = destFolder.getFilesByName(pdfName);
    while (oldIt.hasNext()) {
      oldIt.next().setTrashed(true);
      Logger.log('[PDF] Deleted old PDF: ' + pdfName);
    }
  } catch(_) {}

  let tempDocId = '';

  try {
    // ── BƯỚC 1: Copy sang Doc tạm (không có tabs) ────────────
    // Doc tạm sẽ chỉ có 1 tab duy nhất = nội dung chính
    const originalDoc  = DocumentApp.openById(docId);
    const originalBody = originalDoc.getBody();

    // Tạo doc tạm trong cùng folder
    const tempDoc  = DocumentApp.create('_temp_' + apptId);
    tempDocId      = tempDoc.getId();
    const tempBody = tempDoc.getBody();

    // Copy toàn bộ nội dung từ doc gốc sang doc tạm
    tempBody.clear();
    const numChildren = originalBody.getNumChildren();

    for (let i = 0; i < numChildren; i++) {
      const el   = originalBody.getChild(i);
      const type = el.getType();

      try {
        if (type === DocumentApp.ElementType.PARAGRAPH) {
          const para = el.asParagraph().copy();
          tempBody.appendParagraph(para);
        } else if (type === DocumentApp.ElementType.TABLE) {
          const tbl = el.asTable().copy();
          tempBody.appendTable(tbl);
        } else if (type === DocumentApp.ElementType.LIST_ITEM) {
          const li = el.asListItem().copy();
          tempBody.appendListItem(li);
        }
      } catch (copyErr) {
        Logger.log('[PDF] Skip element ' + i + ': ' + copyErr.message);
      }
    }

    // Copy page settings (margins, size)
    try {
      const style = {};
      style[DocumentApp.Attribute.MARGIN_TOP]    = originalBody.getMarginTop();
      style[DocumentApp.Attribute.MARGIN_BOTTOM] = originalBody.getMarginBottom();
      style[DocumentApp.Attribute.MARGIN_LEFT]   = originalBody.getMarginLeft();
      style[DocumentApp.Attribute.MARGIN_RIGHT]  = originalBody.getMarginRight();
      tempBody.setAttributes(style);
    } catch(_) {}

    tempDoc.saveAndClose();
    Logger.log('[PDF] Temp doc created: ' + tempDocId);

    // ── BƯỚC 2: Export temp doc → PDF ────────────────────────
    const exportUrl = 'https://docs.google.com/document/d/' + tempDocId + '/export?format=pdf';
    const token     = ScriptApp.getOAuthToken();

    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      Logger.log('[PDF] Export failed: HTTP ' + response.getResponseCode());
      return '';
    }

    // ── BƯỚC 3: Lưu PDF vào destFolder ───────────────────────
    const pdfBlob = response.getBlob().setName(pdfName);
    const pdfFile = destFolder.createFile(pdfBlob);
    Logger.log('[PDF] ✅ Created: ' + pdfName);
    return pdfFile.getUrl();

  } catch (e) {
    Logger.log('[PDF] ❌ ERROR: ' + e.message);
    return '';

  } finally {
    // ── BƯỚC 4: Xóa doc tạm dù thành công hay thất bại ───────
    if (tempDocId) {
      try {
        DriveApp.getFileById(tempDocId).setTrashed(true);
        Logger.log('[PDF] Temp doc deleted: ' + tempDocId);
      } catch(_) {}
    }
  }
}
/**
 * Tạo PDF cho 1 row cụ thể
 * Gọi sau khi Intake Doc đã được fill xong
 */
function generatePdfForRow_(row) {
  Logger.log('[PDF] generatePdfForRow_ row=' + row);

  const intakeUrl = getCell_(SHT.MASTER, row, 'IntakeDocURL');
  const pfId      = getCell_(SHT.MASTER, row, 'ProspectFolderID');
  const brand     = getCell_(SHT.MASTER, row, 'Brand')   || '';
  const apptId    = getCell_(SHT.MASTER, row, 'APPT_ID') || '';

  if (!intakeUrl) {
    Logger.log('[PDF] Skip: no IntakeDocURL on row ' + row);
    return '';
  }
  if (!pfId) {
    Logger.log('[PDF] Skip: no ProspectFolderID on row ' + row);
    return '';
  }
  if (!brand || !apptId) {
    Logger.log('[PDF] Skip: missing brand or apptId on row ' + row);
    return '';
  }

  // ── Check đã có PDF chưa ────────────────────────────────────
  const existingPdf = getCell_(SHT.MASTER, row, 'IntakePdfURL');
  if (existingPdf) {
    Logger.log('[PDF] Already set on row ' + row + ' → skip');
    return existingPdf;
  }

  const docId = idFromUrl_(intakeUrl);
  if (!docId) {
    Logger.log('[PDF] Cannot extract docId from: ' + intakeUrl);
    return '';
  }

  let destFolder;
  try {
    destFolder = DriveApp.getFolderById(pfId);
  } catch (e) {
    Logger.log('[PDF] Cannot access ProspectFolder: ' + e.message);
    return '';
  }

  const pdfUrl = exportIntakeDocToPdf_(docId, destFolder, brand, apptId);

  if (pdfUrl) {
    try {
      setCell_(SHT.MASTER, row, 'IntakePdfURL', pdfUrl);
      Logger.log('[PDF] ✅ URL saved to Master row ' + row);
    } catch (e) {
      Logger.log('[PDF] ERROR saving URL: ' + e.message);
    }
  }

  return pdfUrl;
}

// function exportIntakeDocToPdf_(docId, destFolder, brand, apptId) {
//   if (!docId || !destFolder) return '';

//   const pdfName = brand + ' \u2013 ' + apptId + ' \u2013 Intake';

//   try {
//     const docUrl  = 'https://docs.google.com/document/d/' + docId + '/export?format=pdf';
//     const token   = ScriptApp.getOAuthToken();

//     const response = UrlFetchApp.fetch(docUrl, {
//       headers: { 'Authorization': 'Bearer ' + token },
//       muteHttpExceptions: true
//     });

//     if (response.getResponseCode() !== 200) {
//       Logger.log('[PDF] Export failed: HTTP ' + response.getResponseCode());
//       return '';
//     }

//     // ── Xóa PDF cũ cùng tên nếu có ──────────────────────────
//     try {
//       const oldIt = destFolder.getFilesByName(pdfName);
//       while (oldIt.hasNext()) {
//         oldIt.next().setTrashed(true);
//         Logger.log('[PDF] Deleted old PDF: ' + pdfName);
//       }
//     } catch(_) {}

//     const pdfBlob = response.getBlob().setName(pdfName);
//     const pdfFile = destFolder.createFile(pdfBlob);
//     Logger.log('[PDF] ✅ Created: ' + pdfName);
//     return pdfFile.getUrl();

//   } catch (e) {
//     Logger.log('[PDF] ❌ ERROR: ' + e.message);
//     return '';
//   }
// }


// ── 4. ĐẢM BẢO 10 BLANK SLIDES SAU WELCOME ──────────────────

function _ensureTenBlankSlides_(pres) {
  const needed = 10 - (pres.getSlides().length - 1);
  for (let i = 0; i < needed; i++) {
    pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  }
  if (needed > 0) Logger.log(`[Slides] Đã thêm ${needed} blank slides`);
}


function dailyBackfillRootApptIDs_() {
  // Chỉ fix các row trong 30 ngày gần nhất để nhanh
  const s    = SH(SHT.MASTER);
  const H    = headers_(SHT.MASTER);
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  if (last < 2) return;

  const cEmail = H['EmailLower']  || 0;
  const cPhone = H['PhoneNorm']   || 0;
  const cAppt  = H['APPT_ID']     || 0;
  const cRoot  = H['RootApptID']  || 0;
  const cTS    = H['Timestamp']   || 0;

  if (!cAppt || !cRoot) return;

  const n = last - 1;
  const apptVec  = s.getRange(2,cAppt,n,1).getValues().map(r=>String(r[0]||'').trim());
  const emailVec = cEmail ? s.getRange(2,cEmail,n,1).getValues().map(r=>String(r[0]||'').toLowerCase().trim()) : Array(n).fill('');
  const phoneVec = cPhone ? s.getRange(2,cPhone,n,1).getValues().map(r=>String(r[0]||'').trim())              : Array(n).fill('');
  const rootVec  = s.getRange(2,cRoot,n,1).getValues().map(r=>String(r[0]||'').trim());

  // Tìm earliest APPT_ID cho mỗi contact
  const sorted = apptVec.map((a,i)=>({a,i})).filter(x=>x.a).sort((a,b)=>a.a.localeCompare(b.a));
  const canon = {};
  sorted.forEach(({a,i}) => {
    const e = emailVec[i], p = phoneVec[i];
    if (e && !canon['e:'+e]) canon['e:'+e] = a;
    if (p && !canon['p:'+p]) canon['p:'+p] = a;
  });

  let fixed = 0;
  for (let i=0;i<n;i++) {
    const row = i+2;
    if (!apptVec[i]) continue;
    const e = emailVec[i], p = phoneVec[i];
    const canonical = canon['e:'+e] || canon['p:'+p] || apptVec[i];
    if (rootVec[i] !== canonical) {
      setCell_(SHT.MASTER, row, 'RootApptID', canonical);
      fixed++;
    }
  }
  if (fixed) Logger.log('dailyBackfill: fixed ' + fixed + ' rows');
}

function installDailyBackfill() {
  const FN = 'dailyBackfillRootApptIDs_';
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === FN)) {
    ScriptApp.newTrigger(FN).timeBased().everyDays(1).atHour(3).create();
    Logger.log('Installed daily backfill trigger at 3am');
  }
}

function fix_changeTriggerTo2Min() {
  const FN = 'ensureBootstrapForRecentRows_';

  // Xóa trigger cũ (1 phút)
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === FN)
    .forEach(t => {
      ScriptApp.deleteTrigger(t);
      Logger.log('Đã xóa trigger cũ: ' + t.getUniqueId());
    });

  // Tạo trigger mới (2 phút)
  ScriptApp.newTrigger(FN)
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log('✅ Đã tạo trigger mới: ' + FN + ' chạy mỗi 2 phút');
}

function debug_diagRepeatCustomerRow() {
  const sh = SH(SHT.MASTER);
  const H  = headers_(SHT.MASTER);
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);

  // ── Tìm 2 row gần nhất cùng email/phone ──────────────────────
  const cEmail = H['EmailLower'] || 0;
  const cPhone = H['PhoneNorm']  || 0;
  const cAppt  = H['APPT_ID']    || 0;
  const cPfId  = H['ProspectFolderID'] || 0;
  const cInt   = H['IntakeDocURL']     || 0;
  const cChk   = H['Checklist URL']    || 0;
  const cQuo   = H['Quotation URL']    || 0;
  const cBrand = H['Brand']            || 0;

  if (last < 3) { Logger.log('Cần ít nhất 2 data rows'); return; }

  const rows = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();

  // Lấy row cuối cùng có APPT_ID
  let row2Idx = 0, row1Idx = 0;
  for (let i = rows.length - 1; i >= 0; i--) {
    const appt = cAppt ? String(rows[i][cAppt-1]||'').trim() : '';
    if (!appt) continue;
    if (!row2Idx) { row2Idx = i + 2; continue; }
    if (!row1Idx) {
      const e2 = cEmail ? String(rows[row2Idx-2][cEmail-1]||'').toLowerCase() : '';
      const e1 = cEmail ? String(rows[i][cEmail-1]||'').toLowerCase() : '';
      const p2 = cPhone ? String(rows[row2Idx-2][cPhone-1]||'') : '';
      const p1 = cPhone ? String(rows[i][cPhone-1]||'') : '';
      if ((e1 && e1 === e2) || (p1 && p1 === p2)) {
        row1Idx = i + 2;
        break;
      }
    }
  }

  if (!row1Idx) {
    Logger.log('❌ Không tìm thấy 2 row cùng khách — hãy điền row index thủ công bên dưới');
    return;
  }

  Logger.log('=== REPEAT CUSTOMER DIAGNOSIS ===');
  Logger.log('Row 1 (lần 1): ' + row1Idx);
  Logger.log('Row 2 (lần 2): ' + row2Idx);

  [row1Idx, row2Idx].forEach((r, idx) => {
    const label = idx === 0 ? 'LẦN 1' : 'LẦN 2';
    const appt  = cAppt  ? String(sh.getRange(r, cAppt).getValue()  || '') : '';
    const brand = cBrand ? String(sh.getRange(r, cBrand).getValue() || '') : '';
    const pfId  = cPfId  ? String(sh.getRange(r, cPfId).getValue()  || '') : '';
    const intake= cInt   ? String(sh.getRange(r, cInt).getValue()   || '') : '';
    const chk   = cChk   ? String(sh.getRange(r, cChk).getValue()   || '') : '';
    const quo   = cQuo   ? String(sh.getRange(r, cQuo).getValue()   || '') : '';

    Logger.log('\n--- ' + label + ' (row ' + r + ') ---');
    Logger.log('APPT_ID         : ' + appt);
    Logger.log('Brand           : ' + brand);
    Logger.log('ProspectFolderID: ' + (pfId ? pfId : '❌ TRỐNG'));
    Logger.log('IntakeDocURL    : ' + (intake ? '✅' : '❌ TRỐNG'));
    Logger.log('Checklist URL   : ' + (chk    ? '✅' : '❌ TRỐNG'));
    Logger.log('Quotation URL   : ' + (quo    ? '✅' : '❌ TRỐNG'));

    // ── Check folder Drive ──
    if (pfId) {
      try {
        const folder = DriveApp.getFolderById(pfId);
        Logger.log('Folder name     : ' + folder.getName());

        // List tất cả files trong folder
        const files = [];
        const it = folder.getFiles();
        while (it.hasNext()) {
          const f = it.next();
          files.push(f.getName());
        }
        Logger.log('Files in folder (' + files.length + '):');
        files.forEach(name => Logger.log('  📄 ' + name));

        // Check tên file kỳ vọng
        const expectedIntake  = brand + ' \u2013 ' + appt + ' \u2013 Intake';
        const expectedChk     = brand + ' \u2013 ' + appt + ' \u2013 Checklist';
        const expectedQuo     = brand + ' \u2013 ' + appt + ' \u2013 Quotation';
        Logger.log('Expected Intake  : "' + expectedIntake + '" → ' + (files.includes(expectedIntake)  ? '✅ found' : '❌ NOT FOUND'));
        Logger.log('Expected Checklist: "' + expectedChk   + '" → ' + (files.includes(expectedChk)    ? '✅ found' : '❌ NOT FOUND'));
        Logger.log('Expected Quotation: "' + expectedQuo   + '" → ' + (files.includes(expectedQuo)    ? '✅ found' : '❌ NOT FOUND'));

      } catch (e) {
        Logger.log('❌ Folder không truy cập được: ' + e.message);
      }
    }
  });

  Logger.log('\n=== ACTION ===');
  Logger.log('Nếu LẦN 2 thiếu URLs → chạy lệnh sau để force fix:');
  Logger.log('  ensureArtifactsForRow_(' + row2Idx + ')');
}

function debug_fixRow556() {
  const row = 556;
  
  Logger.log('=== FIX ROW ' + row + ' START ===');

  // Step 1: Lấy ProspectFolderID từ row 555 (prior row)
  const priorPfId  = getCell_(SHT.MASTER, 555, 'ProspectFolderID');
  const priorCfId  = getCell_(SHT.MASTER, 555, 'ClientFolderID');
  const priorCfUrl = getCell_(SHT.MASTER, 555, 'Client Folder');
  Logger.log('Prior ClientFolderID : ' + priorCfId);
  Logger.log('Prior Client Folder  : ' + priorCfUrl);
  Logger.log('Prior ProspectFolderID: ' + priorPfId);

  if (!priorPfId) {
    Logger.log('❌ Không lấy được prior ProspectFolderID');
    return;
  }

  // Step 2: Set ProspectFolderID cho row 556 trước
  setCell_(SHT.MASTER, row, 'ProspectFolderID', priorPfId);
  Logger.log('✅ ProspectFolderID đã set: ' + priorPfId);

  // Step 3: Scan folder tìm files của AP-20260531-002
  const folder = DriveApp.getFolderById(priorPfId);
  const brand  = getCell_(SHT.MASTER, row, 'Brand')   || 'HPUSA';
  const apptId = getCell_(SHT.MASTER, row, 'APPT_ID') || 'AP-20260531-002';

  const expectedIntake = brand + ' \u2013 ' + apptId + ' \u2013 Intake';
  const expectedChk    = brand + ' \u2013 ' + apptId + ' \u2013 Checklist';
  const expectedQuo    = brand + ' \u2013 ' + apptId + ' \u2013 Quotation';

  const fileMap = {};
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    fileMap[f.getName()] = f.getUrl();
  }

  const intakeUrl = fileMap[expectedIntake] || '';
  const chkUrl    = fileMap[expectedChk]    || '';
  const quoUrl    = fileMap[expectedQuo]    || '';

  Logger.log('Intake URL  : ' + (intakeUrl ? '✅' : '❌'));
  Logger.log('Checklist URL: ' + (chkUrl   ? '✅' : '❌'));
  Logger.log('Quotation URL: ' + (quoUrl   ? '✅' : '❌'));

  // Step 4: Ghi URLs vào sheet
  const pending = {};
  if (intakeUrl)  pending['IntakeDocURL']  = intakeUrl;
  if (chkUrl)     pending['Checklist URL'] = chkUrl;
  if (quoUrl)     pending['Quotation URL'] = quoUrl;
  if (priorCfId)  pending['ClientFolderID'] = priorCfId;
  if (priorCfUrl) pending['Client Folder']  = priorCfUrl;

  if (Object.keys(pending).length) {
    _atomicWriteUrls_(row, pending);

    // ✅ Verify thực tế sau khi ghi
    const checkIntake = getCell_(SHT.MASTER, row, 'IntakeDocURL');
    const checkChk    = getCell_(SHT.MASTER, row, 'Checklist URL');
    const checkQuo    = getCell_(SHT.MASTER, row, 'Quotation URL');

    if (checkIntake && checkChk && checkQuo) {
      Logger.log('✅ URLs đã ghi thành công');
    } else {
      Logger.log('❌ Ghi thất bại - chạy lại lần nữa');
    }
  } else {
    Logger.log('❌ Không tìm thấy files trong folder — chạy ensureArtifactsForRow_(' + row + ')');
  }

  // Step 5: Verify
  Logger.log('\n=== VERIFY ===');
  Logger.log('ProspectFolderID: ' + (getCell_(SHT.MASTER, row, 'ProspectFolderID') ? '✅' : '❌'));
  Logger.log('IntakeDocURL    : ' + (getCell_(SHT.MASTER, row, 'IntakeDocURL')     ? '✅' : '❌'));
  Logger.log('Checklist URL   : ' + (getCell_(SHT.MASTER, row, 'Checklist URL')    ? '✅' : '❌'));
  Logger.log('Quotation URL   : ' + (getCell_(SHT.MASTER, row, 'Quotation URL')    ? '✅' : '❌'));

  Logger.log('=== FIX ROW ' + row + ' END ===');
  Logger.log('Client Folder   : ' + (getCell_(SHT.MASTER, row, 'Client Folder')  ? '✅' : '❌'));
  Logger.log('ClientFolderID  : ' + (getCell_(SHT.MASTER, row, 'ClientFolderID') ? '✅' : '❌'));
}

function debug_testArtifactSpeed() {
  const row = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  if (row < 2) { Logger.log('No data'); return; }

  const t0 = Date.now();
  ensureArtifactsForRow_(row);
  Logger.log('⏱ Total time: ' + (Date.now() - t0) + 'ms');

  // Verify URLs đã được ghi
  Logger.log('IntakeDocURL  : ' + (getCell_(SHT.MASTER, row, 'IntakeDocURL')  ? '✅' : '❌'));
  Logger.log('Checklist URL : ' + (getCell_(SHT.MASTER, row, 'Checklist URL') ? '✅' : '❌'));
  Logger.log('Quotation URL : ' + (getCell_(SHT.MASTER, row, 'Quotation URL') ? '✅' : '❌'));
  Logger.log('Client Folder : ' + (getCell_(SHT.MASTER, row, 'Client Folder') ? '✅' : '❌'));
}

function debug_diagDuplicateRows() {
  const masterSh = SH(SHT.MASTER);
  const inboxSh  = SH(SHT.FORM_INBOX);
  const MH = headers_(SHT.MASTER);
  const IH = headers_(SHT.FORM_INBOX);

  // ── 1. Kiểm tra Form_Inbox có bao nhiêu row ──────────────────
  const inboxLast = inboxSh.getLastRow();
  Logger.log('=== FORM_INBOX ===');
  Logger.log('Total rows (kể header): ' + inboxLast);
  Logger.log('Data rows: ' + (inboxLast - 1));

  if (inboxLast >= 2) {
    const inboxData = inboxSh.getRange(2, 1, inboxLast - 1, inboxSh.getLastColumn()).getValues();
    inboxData.forEach((r, i) => {
      const ts    = IH['Timestamp']         ? String(r[IH['Timestamp']-1]         || '') : '';
      const email = IH['Email']             ? String(r[IH['Email']-1]             || '') : '';
      const uid   = IH['Admin: Calendly Event UID'] 
                    ? String(r[IH['Admin: Calendly Event UID']-1] || '') : '';
      Logger.log(`Inbox row ${i+2}: email=${email} | ts=${ts} | uid=${uid.slice(0,20)}`);
    });
  }

  // ── 2. Kiểm tra Master có bao nhiêu row trùng ────────────────
  const masterLast = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  Logger.log('\n=== MASTER (2 rows cuối) ===');

  const cEmail = MH['EmailLower'] || 0;
  const cPhone = MH['PhoneNorm']  || 0;
  const cAppt  = MH['APPT_ID']    || 0;
  const cUID   = MH['CalendlyEventUID'] || 0;
  const cTS    = MH['Timestamp']  || 0;
  const cSta   = MH['Status']     || 0;

  const checkRows = Math.min(5, masterLast - 1);
  const startRow  = masterLast - checkRows + 1;
  const block = masterSh.getRange(startRow, 1, checkRows, masterSh.getLastColumn()).getValues();

  block.forEach((r, i) => {
    const row   = startRow + i;
    const appt  = cAppt  ? String(r[cAppt-1]  || '') : '';
    const email = cEmail ? String(r[cEmail-1] || '') : '';
    const uid   = cUID   ? String(r[cUID-1]   || '') : '';
    const ts    = cTS    ? String(r[cTS-1]    || '') : '';
    const sta   = cSta   ? String(r[cSta-1]   || '') : '';
    Logger.log(`Master row ${row}: appt=${appt} | email=${email} | uid=${uid.slice(0,20)} | ts=${ts} | status=${sta}`);
  });

  // ── 3. Kiểm tra triggers đang active ─────────────────────────
  Logger.log('\n=== TRIGGERS ===');
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    Logger.log(`${t.getHandlerFunction()} | type=${t.getEventType()} | source=${t.getTriggerSource()}`);
  });

  // ── 4. Tìm duplicate CalendlyEventUID trong Master ───────────
  Logger.log('\n=== DUPLICATE UID CHECK ===');
  if (cUID && masterLast >= 2) {
    const uids = masterSh.getRange(2, cUID, masterLast - 1, 1).getValues().flat().map(v => String(v||'').trim());
    const seen = {}, dupes = [];
    uids.forEach((u, i) => {
      if (!u) return;
      if (seen[u]) dupes.push({ uid: u, rows: [seen[u], i + 2] });
      else seen[u] = i + 2;
    });
    if (dupes.length) {
      Logger.log('❌ Tìm thấy ' + dupes.length + ' duplicate UID:');
      dupes.forEach(d => Logger.log('  UID: ' + d.uid.slice(0,30) + ' | rows: ' + d.rows.join(', ')));
    } else {
      Logger.log('✅ Không có duplicate UID');
      Logger.log('→ Duplicate do trigger kép hoặc 2 form submission khác nhau');
    }
  }
}
function debug_testDedupGuard() {
  Logger.log('===== DEDUP GUARD TEST =====');

  const cache = CacheService.getScriptCache();

  // ── Test 1: Cùng UID → phải chặn lần 2 ──────────────────────
  const uid1 = 'test-uid-abc123';
  const key1 = 'formsubmit_' + uid1;
  cache.remove(key1); // reset

  const fakeNV1 = {
    'Timestamp':                     ['Mon Apr 21 2026 10:00:00'],
    'Email':                         ['khachA@gmail.com'],
    'Admin: Calendly Event UID':     [uid1],
    'Company':                       ['HPUSA'],
    'Customer Name':                 ['Test Khach A'],
    'Visit Type':                    ['Appointment'],
  };

  Logger.log('\n-- Test 1: Cùng UID fire 2 lần --');
  Logger.log('Lần 1 → expect: xử lý bình thường');
  // Giả lập check dedup
  const hit1a = !!cache.get(key1);
  if (!hit1a) { cache.put(key1, '1', 60); }
  Logger.log('  Cache hit: ' + hit1a + ' → ' + (hit1a ? '🛑 BLOCKED' : '✅ PASS THROUGH'));

  Logger.log('Lần 2 → expect: bị chặn');
  const hit1b = !!cache.get(key1);
  Logger.log('  Cache hit: ' + hit1b + ' → ' + (hit1b ? '🛑 BLOCKED ✅' : '❌ NOT BLOCKED'));

  cache.remove(key1);

  // ── Test 2: 2 khách khác nhau cùng lúc → cả 2 phải pass ─────
  const uid2a = 'test-uid-khachB';
  const uid2b = 'test-uid-khachC';
  const key2a = 'formsubmit_' + uid2a;
  const key2b = 'formsubmit_' + uid2b;
  cache.remove(key2a);
  cache.remove(key2b);

  Logger.log('\n-- Test 2: 2 khách khác nhau cùng lúc --');

  const hit2a = !!cache.get(key2a);
  if (!hit2a) cache.put(key2a, '1', 60);
  Logger.log('Khách B → Cache hit: ' + hit2a + ' → ' + (hit2a ? '🛑 BLOCKED' : '✅ PASS THROUGH'));

  const hit2b = !!cache.get(key2b);
  if (!hit2b) cache.put(key2b, '1', 60);
  Logger.log('Khách C → Cache hit: ' + hit2b + ' → ' + (hit2b ? '🛑 BLOCKED' : '✅ PASS THROUGH'));

  cache.remove(key2a);
  cache.remove(key2b);

  // ── Test 3: Không có UID, dùng email + timestamp ──────────────
  Logger.log('\n-- Test 3: Không có UID, dùng email + ts --');
  const email3 = 'walkin@gmail.com';
  const ts3    = 'MonApr2120261000';
  const key3   = 'formsubmit_' + email3 + '_' + ts3;
  cache.remove(key3);

  const hit3a = !!cache.get(key3);
  if (!hit3a) cache.put(key3, '1', 60);
  Logger.log('Walk-in lần 1 → ' + (hit3a ? '🛑 BLOCKED' : '✅ PASS THROUGH'));

  const hit3b = !!cache.get(key3);
  Logger.log('Walk-in lần 2 → ' + (hit3b ? '🛑 BLOCKED ✅' : '❌ NOT BLOCKED'));

  cache.remove(key3);

  // ── Test 4: Không có UID, không có email → skip dedup ────────
  Logger.log('\n-- Test 4: Không có UID + email → skip dedup --');
  Logger.log('Expect: không set cache, cho qua bình thường');
  Logger.log('→ ✅ Handled by: if (!_calUID && !_email) skip dedup');

  // ── Test 5: Verify duplicate rows trong Master ────────────────
  Logger.log('\n-- Test 5: Duplicate APPT_ID trong Master --');
  const masterSh = SH(SHT.MASTER);
  const H = headers_(SHT.MASTER);
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  const cAppt = H['APPT_ID'] || 0;

  if (cAppt && last >= 3) {
    const checkN   = Math.min(20, last - 1);
    const startRow = last - checkN + 1;
    const vals = masterSh.getRange(startRow, cAppt, checkN, 1)
                 .getValues().flat().map(v => String(v||'').trim());

    const seen = {}, dupes = [];
    vals.forEach((v, i) => {
      if (!v) return;
      const absRow = startRow + i;
      if (seen[v] !== undefined) {
        dupes.push({ appt: v, rows: [seen[v], absRow] });
      } else {
        seen[v] = absRow;
      }
    });

    if (dupes.length) {
      Logger.log('❌ Tìm thấy ' + dupes.length + ' duplicate APPT_ID trong ' + checkN + ' rows cuối:');
      dupes.forEach(d => Logger.log('  ' + d.appt + ' → rows ' + d.rows.join(' & ')));
    } else {
      Logger.log('✅ Không có duplicate APPT_ID trong ' + checkN + ' rows cuối');
    }
  }

  Logger.log('\n===== DEDUP TEST COMPLETE =====');
  Logger.log('Expected results:');
  Logger.log('  Test 1: Lần 1 PASS, Lần 2 BLOCKED ✅');
  Logger.log('  Test 2: Cả 2 khách PASS ✅');
  Logger.log('  Test 3: Lần 1 PASS, Lần 2 BLOCKED ✅');
  Logger.log('  Test 4: Skip dedup ✅');
  Logger.log('  Test 5: 0 duplicates ✅');
}

function debug_monitorDuplicates() {
  const masterSh = SH(SHT.MASTER);
  const H  = headers_(SHT.MASTER);
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);

  const cAppt  = H['APPT_ID']          || 0;
  const cEmail = H['EmailLower']        || 0;
  const cUID   = H['CalendlyEventUID']  || 0;
  const cTS    = H['Timestamp']         || 0;
  const cNotes = H['Automation Notes']  || 0;

  Logger.log('===== DUPLICATE MONITOR (last 30 rows) =====');

  const checkN   = Math.min(30, last - 1);
  const startRow = last - checkN + 1;
  const block    = masterSh.getRange(startRow, 1, checkN, masterSh.getLastColumn()).getValues();

  // Group theo UID
  const byUID   = {};
  const byAppt  = {};

  block.forEach((r, i) => {
    const row   = startRow + i;
    const appt  = cAppt  ? String(r[cAppt-1]  || '').trim() : '';
    const email = cEmail ? String(r[cEmail-1] || '').trim() : '';
    const uid   = cUID   ? String(r[cUID-1]   || '').trim() : '';
    const ts    = cTS    ? String(r[cTS-1]    || '').trim() : '';
    const notes = cNotes ? String(r[cNotes-1] || '').trim() : '';

    if (!appt) return;

    // Check duplicate APPT_ID
    if (!byAppt[appt]) byAppt[appt] = [];
    byAppt[appt].push({ row, email, uid, ts, notes });

    // Check duplicate UID
    if (uid) {
      if (!byUID[uid]) byUID[uid] = [];
      byUID[uid].push({ row, appt, email });
    }
  });

  // Report duplicate APPT_ID
  let hasDupe = false;
  Object.entries(byAppt).forEach(([appt, entries]) => {
    if (entries.length < 2) return;
    hasDupe = true;
    Logger.log('\n❌ DUPLICATE APPT_ID: ' + appt);
    entries.forEach(e => {
      Logger.log('  Row ' + e.row + ' | email=' + e.email + ' | uid=' + e.uid.slice(0,20));
    });
  });

  // Report duplicate UID
  Object.entries(byUID).forEach(([uid, entries]) => {
    if (entries.length < 2) return;
    hasDupe = true;
    Logger.log('\n❌ DUPLICATE UID: ' + uid.slice(0,30));
    entries.forEach(e => {
      Logger.log('  Row ' + e.row + ' | appt=' + e.appt + ' | email=' + e.email);
    });
  });

  if (!hasDupe) {
    Logger.log('✅ Không có duplicate trong 30 rows cuối — dedup guard hoạt động tốt');
  }

  Logger.log('\n===== MONITOR COMPLETE =====');
}

function debug_testPdfExport() {
  const row = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  if (row < 2) { Logger.log('No data rows'); return; }

  Logger.log('===== PDF EXPORT TEST row=' + row + ' =====');

  const intakeUrl = getCell_(SHT.MASTER, row, 'IntakeDocURL');
  const pfId      = getCell_(SHT.MASTER, row, 'ProspectFolderID');
  const brand     = getCell_(SHT.MASTER, row, 'Brand')   || '';
  const apptId    = getCell_(SHT.MASTER, row, 'APPT_ID') || '';

  Logger.log('Brand    : ' + (brand   || '❌ MISSING'));
  Logger.log('ApptId   : ' + (apptId  || '❌ MISSING'));
  Logger.log('PfId     : ' + (pfId    || '❌ MISSING'));
  Logger.log('IntakeURL: ' + (intakeUrl ? intakeUrl.slice(0, 60) + '...' : '❌ MISSING'));

  if (!intakeUrl || !pfId || !brand || !apptId) {
    Logger.log('❌ Missing required fields — abort');
    return;
  }

  // ── Kiểm tra URL type ────────────────────────────────────────
  const isPdf = intakeUrl.includes('drive.google.com/file');
  const isDoc = intakeUrl.includes('docs.google.com/document');
  Logger.log('URL type : ' + (isPdf ? '📄 PDF (already converted)' : isDoc ? '📝 Google Doc' : '❓ Unknown'));

  if (isPdf) {
    Logger.log('✅ Đã là PDF rồi — không cần convert. Test hoàn tất.');
    return;
  }
  if (!isDoc) {
    Logger.log('❌ URL không phải Google Doc — không thể export');
    return;
  }

  // ── Extract docId ────────────────────────────────────────────
  const docId = idFromUrl_(intakeUrl);
  if (!docId) { Logger.log('❌ Cannot extract docId from URL'); return; }
  Logger.log('DocId    : ' + docId);

  // ── Access folder ────────────────────────────────────────────
  let destFolder;
  try {
    destFolder = DriveApp.getFolderById(pfId);
    Logger.log('Folder   : ✅ ' + destFolder.getName());
  } catch (e) {
    Logger.log('❌ Cannot access folder: ' + e.message);
    return;
  }

  // ── Test export URL trực tiếp trước ─────────────────────────
  Logger.log('\n--- Testing export URL ---');
  try {
    const testUrl  = 'https://docs.google.com/document/d/' + docId + '/export?format=pdf&rm=minimal';
    const token    = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(testUrl, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const code = response.getResponseCode();
    Logger.log('HTTP response : ' + code + (code === 200 ? ' ✅' : ' ❌'));
    if (code === 200) {
      const blob = response.getBlob();
      Logger.log('PDF size     : ' + (blob.getBytes().length / 1024).toFixed(1) + ' KB');
      Logger.log('Content type : ' + blob.getContentType());
    } else {
      Logger.log('Response body: ' + response.getContentText().slice(0, 200));
      return;
    }
  } catch (e) {
    Logger.log('❌ Fetch error: ' + e.message);
    return;
  }

  // ── Thực hiện export PDF thật ────────────────────────────────
  Logger.log('\n--- Exporting PDF ---');
  const t0     = Date.now();
  const pdfUrl = exportIntakeDocToPdf_(docId, destFolder, brand, apptId);
  const ms     = Date.now() - t0;

  Logger.log('⏱ Export time: ' + ms + 'ms');

  if (!pdfUrl) {
    Logger.log('❌ PDF export returned empty URL');
    return;
  }

  Logger.log('PDF URL: ' + pdfUrl);

  // ── Ghi PDF URL vào IntakeDocURL ─────────────────────────────
  Logger.log('\n--- Updating IntakeDocURL in sheet ---');
  try {
    setCell_(SHT.MASTER, row, 'IntakeDocURL', pdfUrl);
    const saved = getCell_(SHT.MASTER, row, 'IntakeDocURL');
    Logger.log('Saved URL: ' + (saved === pdfUrl ? '✅ Match' : '❌ Mismatch'));
  } catch (e) {
    Logger.log('❌ Cannot save URL: ' + e.message);
  }

  // ── Xóa Doc gốc ──────────────────────────────────────────────
  Logger.log('\n--- Deleting original Doc ---');
  try {
    DriveApp.getFileById(docId).setTrashed(true);
    Logger.log('✅ Doc gốc đã xóa');
  } catch (e) {
    Logger.log('⚠️ Không xóa được Doc gốc: ' + e.message);
  }

  // ── Verify cuối ──────────────────────────────────────────────
  Logger.log('\n===== FINAL VERIFY =====');
  const finalUrl = getCell_(SHT.MASTER, row, 'IntakeDocURL');
  Logger.log('IntakeDocURL : ' + (finalUrl ? finalUrl.slice(0, 60) + '...' : '❌ EMPTY'));
  Logger.log('Is PDF       : ' + (finalUrl && finalUrl.includes('drive.google.com/file') ? '✅' : '❌'));
  Logger.log('Time total   : ' + ms + 'ms');
  Logger.log('===== TEST COMPLETE =====');
}

function debug_findRowWithDocUrl() {
  const s    = SH(SHT.MASTER);
  const H    = headers_(SHT.MASTER);
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  const cInt = H['IntakeDocURL'] || 0;
  const cApp = H['APPT_ID']     || 0;

  if (!cInt || last < 2) { Logger.log('No data'); return; }

  const vals = s.getRange(2, cInt, last-1, 1).getValues();
  const appts = cApp ? s.getRange(2, cApp, last-1, 1).getValues() : [];

  Logger.log('===== TÌM ROW CÓ DOC URL =====');
  let found = 0;
  for (let i = vals.length-1; i >= 0; i--) {
    const url   = String(vals[i][0] || '').trim();
    const appt  = cApp ? String(appts[i][0] || '') : '';
    const row   = i + 2;

    if (!url) continue;

    const isPdf = url.includes('drive.google.com/file');
    const isDoc = url.includes('docs.google.com/document');

    if (isDoc) {
      Logger.log('✅ Row ' + row + ' (' + appt + '): Google Doc URL');
      Logger.log('   ' + url.slice(0, 70));
      found++;
      if (found >= 3) break; // Lấy 3 row gần nhất
    }
  }

  if (!found) {
    Logger.log('⚠️ Không tìm thấy row nào còn Doc URL');
    Logger.log('→ Tất cả đã được convert sang PDF hoặc chưa có IntakeDocURL');
  }

  Logger.log('===== DONE =====');
}

function debug_testPdfNoBlanKPage() {
  // Tìm row có Doc URL còn sót
  const s    = SH(SHT.MASTER);
  const H    = headers_(SHT.MASTER);
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  const cInt = H['IntakeDocURL'] || 0;
  const cApp = H['APPT_ID']     || 0;
  const cBrand = H['Brand']     || 0;
  const cPfId  = H['ProspectFolderID'] || 0;

  // Tìm row mới nhất có Doc URL
  let testRow = 0;
  for (let i = last - 1; i >= 1; i--) {
    const url = cInt ? String(s.getRange(i+1, cInt).getValue() || '') : '';
    if (url.includes('docs.google.com/document')) {
      testRow = i + 1;
      break;
    }
  }

  if (!testRow) {
    Logger.log('⚠️ Không tìm thấy row có Doc URL');
    Logger.log('→ Tạo 1 form test mới để thử');
    return;
  }

  const intakeUrl = getCell_(SHT.MASTER, testRow, 'IntakeDocURL');
  const pfId      = getCell_(SHT.MASTER, testRow, 'ProspectFolderID');
  const brand     = getCell_(SHT.MASTER, testRow, 'Brand')   || '';
  const apptId    = getCell_(SHT.MASTER, testRow, 'APPT_ID') || '';

  Logger.log('===== PDF NO BLANK PAGE TEST =====');
  Logger.log('Row    : ' + testRow);
  Logger.log('Brand  : ' + brand);
  Logger.log('ApptId : ' + apptId);

  const docId = idFromUrl_(intakeUrl);
  if (!docId) { Logger.log('❌ Cannot extract docId'); return; }

  let destFolder;
  try {
    destFolder = DriveApp.getFolderById(pfId);
    Logger.log('Folder : ✅ ' + destFolder.getName());
  } catch(e) {
    Logger.log('❌ Folder error: ' + e.message);
    return;
  }

  const t0     = Date.now();
  const pdfUrl = exportIntakeDocToPdf_(docId, destFolder, brand, apptId);
  Logger.log('⏱ Time : ' + (Date.now() - t0) + 'ms');

  if (pdfUrl) {
    Logger.log('✅ PDF URL: ' + pdfUrl);
    Logger.log('→ Mở link và kiểm tra không còn trang "V2 W/O Carat Chart"');
  } else {
    Logger.log('❌ Export failed');
  }
}

/**
 * Backfill: copy PaymentsFolderURL từ row đầu tiên của mỗi khách
 * sang tất cả rows còn lại của cùng khách đó.
 * Chạy 1 lần sau khi deploy fix.
 */
function backfillInheritPaymentsFolderURL_() {
  const s    = SH(SHT.MASTER);
  const H    = headers_(SHT.MASTER);
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  if (last < 2) return;

  const cEmail = H['EmailLower']        || 0;
  const cPhone = H['PhoneNorm']         || 0;
  const cPfUrl = H['PaymentsFolderURL'] || 0;
  const cAppt  = H['APPT_ID']           || 0;

  if (!cPfUrl) { Logger.log('Missing PaymentsFolderURL column'); return; }

  const n        = last - 1;
  const emailVec = cEmail ? s.getRange(2,cEmail,n,1).getValues().map(r=>String(r[0]||'').toLowerCase().trim()) : Array(n).fill('');
  const phoneVec = cPhone ? s.getRange(2,cPhone,n,1).getValues().map(r=>String(r[0]||'').trim())              : Array(n).fill('');
  const pfUrlVec = s.getRange(2,cPfUrl,n,1).getValues().map(r=>String(r[0]||'').trim());
  const apptVec  = cAppt  ? s.getRange(2,cAppt, n,1).getValues().map(r=>String(r[0]||'').trim())              : Array(n).fill('');

  // Pass 1: tìm PaymentsFolderURL đầu tiên cho mỗi contact
  const canon = {};
  for (let i = 0; i < n; i++) {
    if (!pfUrlVec[i]) continue;
    const e = emailVec[i], p = phoneVec[i];
    if (e && !canon['e:'+e]) canon['e:'+e] = pfUrlVec[i];
    if (p && !canon['p:'+p]) canon['p:'+p] = pfUrlVec[i];
  }

  // Pass 2: stamp lên rows trống
  let fixed = 0;
  for (let i = 0; i < n; i++) {
    if (pfUrlVec[i]) continue; // đã có rồi
    const row = i + 2;
    const e = emailVec[i], p = phoneVec[i];
    const url = canon['e:'+e] || canon['p:'+p] || '';
    if (!url) continue;
    setCell_(SHT.MASTER, row, 'PaymentsFolderURL', url);
    Logger.log('Row ' + row + ' (' + (apptVec[i]||'?') + '): PaymentsFolderURL inherited → ' + url);
    fixed++;
  }
  Logger.log('backfillInheritPaymentsFolderURL_: fixed ' + fixed + ' rows');
}

function fixExistingRepeatCustomerRows() {
  const s    = SH(SHT.MASTER);
  const H    = headers_(SHT.MASTER);
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);
  if (last < 2) return;

  const cEmail  = H['EmailLower']        || 0;
  const cPhone  = H['PhoneNorm']         || 0;
  const cIntake = H['IntakeDocURL']      || 0;
  const cChk    = H['Checklist URL']     || 0;
  const cQuo    = H['Quotation URL']     || 0;
  const cAppt   = H['APPT_ID']           || 0;
  const cRoot   = H['RootApptID']        || 0;

  if (!cIntake) { Logger.log('Missing IntakeDocURL column'); return; }

  const n = last - 1;
  const rows = s.getRange(2, 1, n, s.getLastColumn()).getValues();

  // Tìm URLs đầu tiên cho mỗi contact
  const canonUrls = {}; // contactKey → { intake, chk, quo }

  rows.forEach((r, i) => {
    const email  = cEmail  ? String(r[cEmail-1] ||'').toLowerCase().trim() : '';
    const phone  = cPhone  ? String(r[cPhone-1] ||'').trim()               : '';
    const intake = cIntake ? String(r[cIntake-1]||'').trim()               : '';
    const chk    = cChk    ? String(r[cChk-1]   ||'').trim()               : '';
    const quo    = cQuo    ? String(r[cQuo-1]   ||'').trim()               : '';
    if (!intake && !chk && !quo) return;

    const keys = [];
    if (email) keys.push('e:'+email);
    if (phone) keys.push('p:'+phone);
    keys.forEach(k => {
      if (!canonUrls[k]) canonUrls[k] = { intake, chk, quo };
    });
  });

  // Stamp URLs lên rows trống
  let fixed = 0;
  rows.forEach((r, i) => {
    const row    = i + 2;
    const email  = cEmail  ? String(r[cEmail-1] ||'').toLowerCase().trim() : '';
    const phone  = cPhone  ? String(r[cPhone-1] ||'').trim()               : '';
    const intake = cIntake ? String(r[cIntake-1]||'').trim()               : '';
    const chk    = cChk    ? String(r[cChk-1]   ||'').trim()               : '';
    const quo    = cQuo    ? String(r[cQuo-1]   ||'').trim()               : '';
    if (intake && chk && quo) return; // đủ rồi

    const canon = canonUrls['e:'+email] || canonUrls['p:'+phone];
    if (!canon) return;

    const pending = {};
    if (!intake && canon.intake) pending['IntakeDocURL']  = canon.intake;
    if (!chk    && canon.chk)    pending['Checklist URL'] = canon.chk;
    if (!quo    && canon.quo)    pending['Quotation URL'] = canon.quo;

    if (Object.keys(pending).length) {
      _atomicWriteUrls_(row, pending);
      Logger.log('Row ' + row + ': fixed ' + Object.keys(pending).join(', '));
      fixed++;
    }
  });

  Logger.log('fixExistingRepeatCustomerRows: fixed ' + fixed + ' rows');
}
