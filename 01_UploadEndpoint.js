/***** Appointment Upload Endpoint *****
 * Expects a raw binary body (audio) with URL params for metadata:
 *   .../exec?token=...&root_appt_id=AP-YYYYMMDD-###&brand=VVS|HPUSA&rep_email=...&filename=...
 *
 * Current uploads are written directly to _AppointmentArtifacts by
 * sw_ingestRawAppointmentUpload_().
 */

// === script props ===
const UP_SP = PropertiesService.getScriptProperties();

function doPost_UPLOAD_(e) {
  if (typeof sw_ingestRawAppointmentUpload_ === 'function') {
    return sw_ingestRawAppointmentUpload_(e);
  }

  Logger.log('[UPLOAD] sw_ingestRawAppointmentUpload_ is unavailable; current upload workflow is not loaded.');
  return ContentService
    .createTextOutput('ACK (appointment upload handler unavailable)')
    .setMimeType(ContentService.MimeType.TEXT);
}

function processUploadQueue() {
  if (typeof sw_retireLegacyAppointmentTrigger_ === 'function') {
    return sw_retireLegacyAppointmentTrigger_('processUploadQueue');
  }
  Logger.log('processUploadQueue is retired. Current worker: sw_processAppointmentAutomation.');
}

function _hard(code, msg){ const e = new Error(msg); e.code = code; e.retry = false; return e; }

function _resolveApFolderId_(ss, rootApptId){
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw _hard('ERROR_BAD_SHEET', 'Missing sheet: 00_Master Appointments');

  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x=>String(x||'').trim());
  const iRoot   = hdr.indexOf('RootApptID');
  const iFolder = hdr.indexOf('RootAppt Folder ID');
  if (iRoot < 0 || iFolder < 0) {
    throw _hard('ERROR_MISSING_COLUMNS','Need columns RootApptID and RootAppt Folder ID');
  }

  const rowIdx = _findRowByRootApptId_(sh, iRoot, rootApptId); // 2..N or 0 if not found
  if (!rowIdx) throw _hard('ERROR_ROOT_NOT_FOUND', `RootApptID not in Master: ${rootApptId}`);

  // Try read once
  let id = String(sh.getRange(rowIdx, iFolder+1).getValue() || '').trim();
  if (id) return id;

  // Auto-heal: attempt bootstrap once (idempotent)
  try {
    const lock = LockService.getScriptLock();
    if (lock.tryLock(5000)) {
      try {
        // Call your bootstrapper on this row
        bootstrapApptFolder_(rowIdx);
      } finally {
        try { lock.releaseLock(); } catch(_){}
      }
    }
  } catch (e) {
    // annotate but keep going to re-read
  }

  // Re-read after bootstrap attempt
  id = String(sh.getRange(rowIdx, iFolder+1).getValue() || '').trim();
  if (id) return id;

  throw _hard('ERROR_NO_ROOT_FOLDER_ID', `RootAppt Folder ID still blank after bootstrap for ${rootApptId}`);
}

function _findRowByRootApptId_(sh, iRoot, rootApptId){
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const rng = sh.getRange(2, iRoot+1, last-1, 1).getValues(); // 2..N in that single column
  for (let i = 0; i < rng.length; i++){
    if (String(rng[i][0]||'').trim() === rootApptId) return i + 2; // sheet row index
  }
  return 0;
}

function setAudioStatusFor(apId, status){
  const ss  = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  const sh  = ss.getSheetByName('00_Master Appointments');
  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(s=>String(s||'').trim());
  const iRoot = hdr.indexOf('RootApptID');

  // find/create Audio Status col (tolerant aliases)
  let iAS = (function(){
    const aliases = ['Audio Status','AudioStatus','Audio status','Status (Audio)'];
    for (const a of aliases){ const i = hdr.indexOf(a); if (i >= 0) return i; }
    sh.insertColumnAfter(sh.getLastColumn());
    const col = sh.getLastColumn();
    sh.getRange(1,col).setValue('Audio Status');
    return col-1;
  })();

  if (iRoot < 0) throw new Error('RootApptID column not found.');
  const last = sh.getLastRow(); if (last < 2) return;

  for (let r=2; r<=last; r++){
    if (String(sh.getRange(r, iRoot+1).getValue()||'').trim() === String(apId).trim()){
      sh.getRange(r, iAS+1).setValue(status);
      return;
    }
  }
}

/** Set a single field on Master row (by RootApptID), creating the column if missing */
function setMasterFieldForRoot_(ss, rootApptId, headerName, value){
  const sh = ss.getSheetByName('00_Master Appointments'); if (!sh) return;
  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x=>String(x||'').trim());
  const iRoot = hdr.indexOf('RootApptID');
  let iCol = hdr.indexOf(headerName);
  if (iCol < 0) { sh.insertColumnAfter(sh.getLastColumn()); iCol = sh.getLastColumn()-1; sh.getRange(1,iCol+1).setValue(headerName); }
  if (iRoot < 0) return;
  const last = sh.getLastRow(); if (last < 2) return;
  for (let r = 2; r <= last; r++){
    if (String(sh.getRange(r, iRoot+1).getValue()||'').trim() === String(rootApptId).trim()){
      sh.getRange(r, iCol+1).setValue(value);
      return;
    }
  }
}

/*******************************
 * Manual Scribe/Strategist summary helpers.
 * Current automatic appointment summary processing lives in
 * sw_processAppointmentAutomation().
 *******************************/

function OPENAI_PROP_(k){ return PropertiesService.getScriptProperties().getProperty(k) || ''; }
const OPENAI_API_KEY_SUM = OPENAI_PROP_('OPENAI_API_KEY');  // reuse same key

/** Public entry: summarize the most recent transcript for a RootApptID (optional).
 *  If rootApptIdOpt is omitted, it summarizes the newest transcript found.
 */
function summarizeLatestTranscript(rootApptIdOpt) {
  if (!OPENAI_API_KEY_SUM) throw new Error('Missing OPENAI_API_KEY Script Property');

  const ss = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  const sh = ss.getSheetByName('00_Master Appointments');

  const HDR = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(s=>String(s||'').trim());
  const iRoot = HDR.indexOf('RootApptID');
  const iFld  = HDR.indexOf('RootAppt Folder ID');
  const iISO  = HDR.indexOf('ApptDateTime (ISO)');
  if (iRoot < 0 || iFld < 0) throw new Error('Missing RootApptID / RootAppt Folder ID columns');

  const last = sh.getLastRow(); if (last < 2) throw new Error('Master has no data rows');

  let rowIdx = 0, rootId = '', apISO = '';
  if (rootApptIdOpt){
    for (let r=2; r<=last; r++){
      if (String(sh.getRange(r,iRoot+1).getValue()||'').trim() === String(rootApptIdOpt).trim()){
        rowIdx = r; break;
      }
    }
    if (!rowIdx) throw new Error('RootApptID not found: '+rootApptIdOpt);
  } else {
    for (let r=last; r>=2; r--){ // newest first
      const fid = String(sh.getRange(r,iFld+1).getValue()||'').trim();
      if (fid){ rowIdx=r; break; }
    }
    if (!rowIdx) throw new Error('No rows with RootAppt Folder ID');
  }

  rootId = String(sh.getRange(rowIdx, iRoot+1).getValue()||'').trim();
  apISO  = iISO>=0 ? String(sh.getRange(rowIdx, iISO+1).getValue()||'') : '';

  const apId = String(sh.getRange(rowIdx, iFld+1).getValue()||'').trim();
  const ap  = DriveApp.getFolderById(apId);

  // 1) Find newest transcript .txt under 03_Transcripts
  const tFolderIt = ap.getFoldersByName('03_Transcripts');
  if (!tFolderIt.hasNext()) throw new Error('No 03_Transcripts folder for '+rootId);
  const tFolder = tFolderIt.next();

  let newest=null, newestTs=0;
  const it = tFolder.getFiles();
  while (it.hasNext()){
    const f = it.next();
    if (!/\.txt$/i.test(f.getName())) continue;
    const ts = f.getDateCreated().getTime();
    if (ts > newestTs){ newest=f; newestTs=ts; }
  }
  if (!newest) throw new Error('No transcript .txt found for '+rootId);

  const transcript = newest.getBlob().getDataAsString('UTF-8');
  // Build a Drive view URL for the newest transcript file
  const transcriptUrl = 'https://drive.google.com/file/d/' + newest.getId() + '/view';

  // 2) Build payload identical in spirit to your Terminal test
  const payload = buildSummarizerPayload_(transcript);

  // 3) Call OpenAI Responses API
  const resultObj = openAIResponses_(payload);
  const scribeNormalized = normalizeScribe_(resultObj);

  // --- MASTER-OWNED IDENTITY: inject name/phone/email from Master into Scribe ---
  try {
    const ssId = PROP_('SPREADSHEET_ID');
    const ms   = SpreadsheetApp.openById(ssId);
    const sh   = ms.getSheetByName('00_Master Appointments');
    const HDR  = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(s=>String(s||'').trim());
    const iRoot= HDR.indexOf('RootApptID');
    const iNm  = HDR.indexOf('Customer Name');
    const iPh  = HDR.indexOf('Phone');        // prefer Phone (display)
    const iPhN = HDR.indexOf('PhoneNorm');    // fallback
    const iEm  = HDR.indexOf('Email');        // prefer Email (display)
    const iEmL = HDR.indexOf('EmailLower');   // fallback

    // locate row for this root
    const last = sh.getLastRow();
    let rowIdx = 0;
    for (let r=2;r<=last;r++){
      if (String(sh.getRange(r,iRoot+1).getValue()||'').trim() === rootId){ rowIdx=r; break; }
    }

    if (rowIdx){
      const name  = iNm  >=0 ? String(sh.getRange(rowIdx,iNm+1 ).getValue()||'').trim() : '';
      const phone = iPh  >=0 ? String(sh.getRange(rowIdx,iPh+1 ).getValue()||'').trim()
                  : iPhN >=0 ? String(sh.getRange(rowIdx,iPhN+1).getValue()||'').trim() : '';
      const email = iEm  >=0 ? String(sh.getRange(rowIdx,iEm+1 ).getValue()||'').trim()
                  : iEmL >=0 ? String(sh.getRange(rowIdx,iEmL+1).getValue()||'').trim() : '';

      scribeNormalized.customer_profile = scribeNormalized.customer_profile || {};
      if (name)  scribeNormalized.customer_profile.customer_name = name;
      if (phone) scribeNormalized.customer_profile.phone = phone;
      if (email) scribeNormalized.customer_profile.email = email;
    }
  } catch(_){}


  // 4) Save JSON snapshot to 04_Summaries/
  const summaryUrl = saveSummaryJson_(ap, rootId, scribeNormalized);

  // 5) Upsert SYS_Consults by ConsultID = RootApptID + '|' + ApptISO (fallback to file ts)
  const isoForId = apISO || newest.getDateCreated().toISOString();
  const consultId = buildConsultId_(rootId, isoForId);
  upsertSYSConsults_(ss, consultId, rootId, scribeNormalized, summaryUrl);

  // after saveSummaryJson_ and SYS_Consults upsert succeed:
  try { runStrategistAnalysisForRoot(rootId); } catch (e) {
    Logger.log('Strategist skipped: ' + (e && e.message || e));
  }

  // 6) after you compute resultObj and save JSON (summaryUrl)
  setMasterFieldForRoot_(ss, rootId, 'Last ConsultID', consultId);
  const needs = needsReviewFromScribe_(scribeNormalized) ? 'TRUE' : 'FALSE';
  setMasterFieldForRoot_(ss, rootId, 'NeedsReview', needs);
  try { setAudioStatusFor(rootId, 'SUMMARIZED'); } catch(_) {}

  try {
    upsertClientSummaryTab_(rootId, scribeNormalized, apISO, transcriptUrl);
  } catch(e){ Logger.log('Summary tab write failed: ' + e.message); }

  try {
    const ss = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
    mirrorSummaryToMaster_(ss, rootId, scribeNormalized);
  } catch(e){ Logger.log('Master mirror failed: ' + e.message); }

  Logger.log('Summarized OK for %s → %s', rootId, summaryUrl);
  return { consultId, summaryUrl };
}

/******************************************************
 * Run Strategist analysis for a given RootApptID
 * - Loads latest Scribe JSON and Transcript
 * - Calls Strategist model
 * - Saves Strategist JSON and mirrors URL to Master
 ******************************************************/
function runStrategistAnalysisForRoot(rootApptId){
  if (!rootApptId) throw new Error('runStrategistAnalysisForRoot: missing rootApptId');

  const ss = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  const sh = ss.getSheetByName('00_Master Appointments');
  const HDR = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(s=>String(s||'').trim());
  const iRoot = HDR.indexOf('RootApptID');
  const iFld  = HDR.indexOf('RootAppt Folder ID');
  const iISO  = HDR.indexOf('ApptDateTime (ISO)');
  if (iRoot < 0 || iFld < 0) throw new Error('Missing RootApptID / RootAppt Folder ID columns');

  // locate row
  const last = sh.getLastRow(); if (last < 2) throw new Error('Master has no data rows');
  let rowIdx = 0;
  for (let r=2; r<=last; r++){
    if (String(sh.getRange(r, iRoot+1).getValue()||'').trim() === String(rootApptId).trim()){ rowIdx = r; break; }
  }
  if (!rowIdx) throw new Error('RootApptID not found: '+rootApptId);

  const apId = String(sh.getRange(rowIdx, iFld+1).getValue()||'').trim();
  const apISO      = (iISO>=0) ? String(sh.getRange(rowIdx, iISO+1).getValue()||'') : '';
  if (!apId) throw new Error('RootAppt Folder ID missing for '+rootApptId);
  const ap = DriveApp.getFolderById(apId);


  // 1) Load newest Scribe JSON from 04_Summaries
  const sFolderIt = ap.getFoldersByName('04_Summaries');
  if (!sFolderIt.hasNext()) throw new Error('No 04_Summaries folder for '+rootApptId);
  const sFolder = sFolderIt.next();

  let newestScribe = null, tsScribe = 0;
  const sIt = sFolder.getFiles();
  while (sIt.hasNext()){
    const f = sIt.next();
    if (!/__summary_.*\.json$/i.test(f.getName())) continue; // scribe json saved by saveSummaryJson_
    const t = f.getDateCreated().getTime();
    if (t > tsScribe){ newestScribe = f; tsScribe = t; }
  }
  if (!newestScribe) throw new Error('No Scribe summary JSON found for '+rootApptId);
  const scribeObj = JSON.parse(newestScribe.getBlob().getDataAsString('UTF-8'));

  // 2) Load newest transcript (optional but improves context)
  let transcript = '';
  const tFolderIt = ap.getFoldersByName('03_Transcripts');
  if (tFolderIt.hasNext()){
    const tFolder = tFolderIt.next();
    let newestTxt = null, tsTxt = 0;
    const tIt = tFolder.getFiles();
    while (tIt.hasNext()){
      const f = tIt.next();
      if (!/\.txt$/i.test(f.getName())) continue;
      const t = f.getDateCreated().getTime();
      if (t > tsTxt){ newestTxt = f; tsTxt = t; }
    }
    if (newestTxt){
      transcript = newestTxt.getBlob().getDataAsString('UTF-8');
    }
  }

  // === Step 2A — Generate MEMO (freeform, transcript REQUIRED) ===
  let memoPayload = buildStrategistMemoPayload_(scribeObj, transcript, '');
  try { memoPayload.meta = { __root: rootApptId, __apId: apId }; } catch(_){}
  const memoText = openAIResponses_TextOnly_(memoPayload);

  // Save memo (and debug copy)
  strat_writeDebug_(ap, rootApptId, 'memo_preview', memoText);
  const memoUrl = saveStrategistMemoText_(ap, rootApptId, memoText);

  // === Step 2B — Extract JSON from MEMO (strict schema) ===
  let extractPayload = buildStrategistExtractPayload_(memoText, scribeObj);
  try { extractPayload.meta = { __root: rootApptId, __apId: apId }; } catch(_){}
  const strategistObj = openAIResponses_(extractPayload);

  // Save extracted JSON + log
  strat_writeDebug_(ap, rootApptId, 'parsed_strategist', strategistObj);
  const strategistUrl = saveStrategistJson_(ap, rootApptId, strategistObj);

  Logger.log('Saved Strategist memo: ' + memoUrl);
  Logger.log('Saved Strategist JSON: ' + strategistUrl + ' keys=' + Object.keys(strategistObj||{}).join(','));
  return { strategistUrl, memoUrl };

} 

function enforceStrictRequired_(schema){
  const clone = JSON.parse(JSON.stringify(schema));
  (function walk(node){
    if (!node || typeof node !== 'object') return;
    if (node.type === 'object' && node.properties && typeof node.properties === 'object') {
      node.required = Object.keys(node.properties); // all keys required
      Object.values(node.properties).forEach(child => {
        if (child && child.type === 'object') walk(child);
        else if (child && child.type === 'array' && child.items && child.items.type === 'object') walk(child.items);
      });
    }
  })(clone);
  return clone;
}


/** Build ConsultID = RootApptID + '|' + ApptISO (iso compacted) */
function buildConsultId_(rootApptId, iso){
  const safeIso = String(iso||'').replace(/\s+/g,'').replace(/:/g,'').replace(/Z$/,'Z');
  return rootApptId + '|' + safeIso;
}

/** True if any *_confidence ≤ 0.69 (root or nested diamond_specs or arrays) */
function needsReview_(o){
  const TH = 0.69;
  function anyLow(v){
    if (!v) return false;
    if (typeof v === 'object'){
      for (const k in v){
        if (/confidence$/i.test(k) && typeof v[k] === 'number' && v[k] <= TH) return true;
        if (v[k] && typeof v[k] === 'object' && anyLow(v[k])) return true;
      }
    }
    return false;
  }
  return anyLow(o);
}

/** Upsert into SYS_Consults using the Scribe schema (normalized) */
function upsertSYSConsults_(ss, consultId, rootApptId, scribe, summaryUrl){
  const SHEET = 'SYS_Consults';
  const HEADERS = [
    // Identity / linkage
    'ConsultID','RootApptID','EventUUID','Brand','Rep','ApptISO',
    'ApId','AudioFolderId','DesignFolderId','TranscriptFileId','SummaryJsonFileId',

    // Customer profile (from Scribe.customer_profile)
    'CustomerName','Phone','Email','PartnerName','CommPrefs','DecisionMakers',

    // Budget / Timeline (+ conf from Scribe.conf)
    'Budget','Budget_conf','Timeline','Timeline_conf',

    // Diamond specs (+ diamond_conf)
    'Diamond_LabOrNatural','Diamond_Shape','Diamond_Carat','Diamond_Color',
    'Diamond_Clarity','Diamond_Ratio','Diamond_CutPolishSym','Diamond_conf',

    // Design specs
    'Design_RingSize','Design_BandWidthMM','Design_WeddingBandFit','Design_Engraving','Design_Notes',

    // Notes / lists
    'RapportNotes','NextSteps','DesignRefs',

    // System
    'NeedsReview','Audio Status','ConfirmedBy','ConfirmedAt','QuotationURL','CreatedAt','LastUpdatedAt'
  ];

  const sh = getOrCreateSheet_(ss, SHEET, HEADERS);
  const H  = shHeaderIndexMap1_(sh);

  // --- Master-derived Brand & Rep (single source of truth)
  const brandFromMaster = (typeof getBrandForRoot_ === 'function') ? getBrandForRoot_(rootApptId) : '';
  const repFromMaster   = (typeof getAssignedRepForRoot_ === 'function' ? getAssignedRepForRoot_(rootApptId) : '')
                        || (typeof getAssistedRepForRoot_ === 'function' ? getAssistedRepForRoot_(rootApptId) : '');

  // --- Drive IDs (best effort)
  let apId = '', audioFolderId = '', designFolderId = '', transcriptFileId = '', summaryFileId = '';
  try {
    apId = (typeof getApFolderIdForRoot_ === 'function')
      ? getApFolderIdForRoot_(ss, rootApptId)
      : _resolveApFolderId_(ss, rootApptId);
    const ap = DriveApp.getFolderById(apId);

    const af = ap.getFoldersByName('01_Audio');       if (af.hasNext()) audioFolderId  = af.next().getId();
    const df = ap.getFoldersByName('02_Design');      if (df.hasNext()) designFolderId = df.next().getId();
    const tf = ap.getFoldersByName('03_Transcripts'); if (tf.hasNext()){
      const tFolder = tf.next();
      let newest=null, ts=0, it=tFolder.getFiles();
      while (it.hasNext()){
        const f = it.next();
        if (!/\.txt$/i.test(f.getName())) continue;
        const tms = (f.getLastUpdated ? f.getLastUpdated() : f.getDateCreated()).getTime();
        if (tms>ts){ newest=f; ts=tms; }
      }
      if (newest) transcriptFileId = newest.getId();
    }
    if (summaryUrl) summaryFileId = String(summaryUrl).split('/d/')[1]?.split('/')[0] || '';
  } catch(_){}

  // --- Flatten Scribe according to your schema
  const flat = flattenScribeForSys_(scribe);

  const map = {
    // Identity
    'ConsultID': consultId,
    'RootApptID': rootApptId,
    'EventUUID': '',
    'Brand': brandFromMaster,
    'Rep': repFromMaster,
    'ApptISO': consultId.split('|')[1] || '',
    'ApId': apId,
    'AudioFolderId': audioFolderId,
    'DesignFolderId': designFolderId,
    'TranscriptFileId': transcriptFileId,
    'SummaryJsonFileId': summaryFileId,

    // Customer profile
    'CustomerName':  flat.customer_name,
    'Phone':         flat.phone,
    'Email':         flat.email,
    'PartnerName':   flat.partner_name,
    'CommPrefs':     flat.comm_prefs,
    'DecisionMakers':flat.decision_makers,

    // Budget/Timeline (+ conf)
    'Budget':        flat.budget,
    'Budget_conf':   flat.budget_conf,
    'Timeline':      flat.timeline,
    'Timeline_conf': flat.timeline_conf,

    // Diamond specs (+ conf)
    'Diamond_LabOrNatural':  flat.diamond_lab_or_natural,
    'Diamond_Shape':         flat.diamond_shape,
    'Diamond_Carat':         valNum(flat.diamond_carat),
    'Diamond_Color':         flat.diamond_color,
    'Diamond_Clarity':       flat.diamond_clarity,
    'Diamond_Ratio':         flat.diamond_ratio,
    'Diamond_CutPolishSym':  flat.diamond_cut_polish_sym,
    'Diamond_conf':          flat.diamond_conf,

    // Design specs
    'Design_RingSize':        flat.design_ring_size,
    'Design_BandWidthMM':     valNum(flat.design_band_width_mm),
    'Design_WeddingBandFit':  flat.design_wedding_band_fit,
    'Design_Engraving':       flat.design_engraving,
    'Design_Notes':           flat.design_notes,

    // Notes / lists
    'RapportNotes': flat.rapport_notes,
    'NextSteps':    flat.next_steps_text,
    'DesignRefs':   flat.design_refs_text,

    // System flags / timestamps
    'NeedsReview':  needsReviewFromScribe_(scribe) ? 'TRUE' : 'FALSE',
    'Audio Status': 'SUMMARIZED',
    'ConfirmedBy':'',
    'ConfirmedAt':'',
    'QuotationURL':'',
    'CreatedAt': new Date(),
    'LastUpdatedAt': new Date()
  };

  // --- Upsert by ConsultID
  const idCol = H['ConsultID'];
  const lastRow = sh.getLastRow();
  let foundRow = 0;
  if (lastRow >= 2){
    const ids = sh.getRange(2, idCol, lastRow-1, 1).getValues().flat();
    const idx = ids.findIndex(v => String(v||'') === consultId);
    if (idx >= 0) foundRow = idx + 2;
  }

  if (!foundRow){
    const row = new Array(sh.getLastColumn()).fill('');
    Object.keys(map).forEach(k => { if (H[k]) row[H[k]-1] = map[k]; });
    sh.appendRow(row);
  } else {
    Object.keys(map).forEach(k => { if (H[k]) sh.getRange(foundRow, H[k]).setValue(map[k]); });
    if (H['LastUpdatedAt']) sh.getRange(foundRow, H['LastUpdatedAt']).setValue(new Date());
  }
}

/** ————— helpers ————— */
function getOrCreateSheet_(ss, name, headers){
  let s = ss.getSheetByName(name);
  if (!s){ s = ss.insertSheet(name); }
  const have = s.getLastColumn() >= headers.length
    ? s.getRange(1,1,1,headers.length).getValues()[0].map(v=>String(v||'').trim())
    : [];
  if (have.length !== headers.length || have.some((v,i)=>v !== headers[i])){
    // reset header exactly
    if (s.getLastColumn() < headers.length){
      s.insertColumnsAfter(s.getLastColumn(), headers.length - s.getLastColumn());
    }
    s.getRange(1,1,1,headers.length).setValues([headers]);
    s.setFrozenRows(1);
  }
  return s;
}

function shHeaderIndexMap1_(sh){
  const row = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const map = {};
  row.forEach((h,i)=>{ if (String(h||'').trim()) map[String(h).trim()] = i+1; });
  return map;
}


function joinArr(a){ return Array.isArray(a) ? a.join(' • ') : ''; }
function safe(v){ return v == null ? '' : String(v); }
function valNum(v){ return (typeof v === 'number' || v === null) ? v : (v===''||v==null? '' : Number(v)); }

/** Convert Scribe JSON into flat fields for SYS_Consults. */
function flattenScribeForSys_(s){
  s = s || {};
  const cp   = s.customer_profile || {};
  const ds   = s.diamond_specs    || {};
  const de   = s.design_specs     || {};
  const conf = s.conf             || {};

  const nextStepsText = (s.next_steps || []).map(ns => {
    const owner = ns && ns.owner ? ns.owner : '';
    const task  = ns && ns.task  ? ns.task  : '';
    const due   = ns && ns.due_iso ? ' (due ' + String(ns.due_iso).split('T')[0] + ')' : '';
    const notes = ns && ns.notes ? ' — ' + ns.notes : '';
    return (owner || task) ? (owner + ': ' + task + due + notes) : '';
  }).filter(Boolean).join(' • ');

  const designRefsText = (s.design_refs || []).map(dr => {
    const name = dr && dr.name ? dr.name : '';
    const file = dr && dr.file ? dr.file : '';
    const desc = dr && dr.desc ? dr.desc : '';
    return [name, file, desc].filter(Boolean).join(' — ');
  }).filter(Boolean).join(' • ');

  return {
    // customer profile
    customer_name:  cp.customer_name || '',
    phone:          cp.phone || '',
    email:          cp.email || '',
    partner_name:   cp.partner_name || '',
    comm_prefs:     (cp.comm_prefs || []).join(', '),
    decision_makers:(cp.decision_makers || []).join(', '),

    // budget / timeline (+conf)
    budget:         s.budget || '',
    budget_conf:    (typeof conf.budget  === 'number') ? conf.budget  : '',
    timeline:       s.timeline || '',
    timeline_conf:  (typeof conf.timeline=== 'number') ? conf.timeline: '',

    // diamond specs (+conf)
    diamond_lab_or_natural: ds.lab_or_natural || '',
    diamond_shape:          ds.shape || '',
    diamond_carat:          (ds.carat==null ? '' : Number(ds.carat)),
    diamond_color:          ds.color || '',
    diamond_clarity:        ds.clarity || '',
    diamond_ratio:          ds.ratio || '',
    diamond_cut_polish_sym: ds.cut_polish_sym || '',
    diamond_conf:           (typeof conf.diamond=== 'number') ? conf.diamond : '',

    // design specs
    design_ring_size:        de.ring_size || '',
    design_band_width_mm:    (de.band_width_mm==null ? '' : Number(de.band_width_mm)),
    design_wedding_band_fit: de.wedding_band_fit || '',
    design_engraving:        de.engraving || '',
    design_notes:            de.design_notes || '',

    // notes / lists (flattened)
    rapport_notes: (s.rapport_notes || []).join(' • '),
    next_steps_text: nextStepsText,
    design_refs_text: designRefsText
  };
}

/** Needs-review if any confidence in Scribe.conf ≤ 0.69. */
function needsReviewFromScribe_(s){
  const c = (s && s.conf) || {};
  return [c.budget, c.timeline, c.diamond].some(v => typeof v === 'number' && v <= 0.69);
}


function processSummariesWorker(){
  if (typeof sw_retireLegacyAppointmentTrigger_ === 'function') {
    return sw_retireLegacyAppointmentTrigger_('processSummariesWorker');
  }
  Logger.log('processSummariesWorker is retired. Current worker: sw_processAppointmentAutomation.');
}

function confirmConsult_(consultId){
  const ss = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  const sh = ss.getSheetByName('SYS_Consults'); if (!sh) throw new Error('Missing SYS_Consults');
  const H = (function(row){ const m={}; row.forEach((h,i)=>m[String(h||'').trim()]=i+1); return m; })
           (sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]);
  const last = sh.getLastRow(); if (last<2) return;
  const idCol = H['ConsultID']; if (!idCol) throw new Error('ConsultID col missing');
  const ids = sh.getRange(2, idCol, last-1, 1).getValues().flat();
  const idx = ids.findIndex(v => String(v||'')===String(consultId));
  if (idx<0) throw new Error('ConsultID not found: '+consultId);
  const row = idx+2;

  if (H['NeedsReview']) sh.getRange(row, H['NeedsReview']).setValue('FALSE');
  if (H['ConfirmedBy']) sh.getRange(row, H['ConfirmedBy']).setValue(Session.getActiveUser().getEmail()||'rep');
  if (H['ConfirmedAt']) sh.getRange(row, H['ConfirmedAt']).setValue(new Date());

  // mirror to Master by RootApptID if you like (optional)
}

function diagConsultUpsert_(rootApptId){
  const ss = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  const sh = ss.getSheetByName('SYS_Consults'); if (!sh) { Logger.log('No SYS_Consults'); return; }
  const H = (function(row){ const m={}; row.forEach((h,i)=>m[String(h||'').trim()]=i+1); return m; })
           (sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]);
  const idCol = H['ConsultID'], rootCol = H['RootApptID'];
  const last = sh.getLastRow();
  if (last<2 || !idCol || !rootCol) { Logger.log('Missing headers or no rows'); return; }
  const roots = sh.getRange(2, rootCol, last-1, 1).getValues().flat();
  const matches = roots.map(String).map(v=>v.trim()).reduce((n,v,i)=> n + (v===rootApptId?1:0), 0);
  Logger.log('SYS_Consults rows for %s: %s', rootApptId, matches);
}

function mirrorSummaryToMaster_(ss, rootApptId, resultObj){
  const sh = ss.getSheetByName('00_Master Appointments');
  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
  const iRoot = hdr.indexOf('RootApptID');
  const iNext = hdr.indexOf('Next Steps');
  if (iRoot < 0 || iNext < 0) return;

  const last = sh.getLastRow(); if (last < 2) return;
  const vals = sh.getRange(2, iRoot+1, last-1, 1).getValues().flat();
  const idx = vals.findIndex(v=>String(v||'').trim()===String(rootApptId).trim());
  if (idx < 0) return;
  const row = idx + 2;

  // Build a concise “Next Steps” line from the first 1–3 followups
  const fu = (resultObj.next_steps||[]).slice(0,3).map(f=>{
    const due = f.due_iso ? ' (due ' + f.due_iso.split('T')[0] + ')' : '';
    return `${f.owner}: ${f.task}${due}`;
  });
  if (fu.length){
    const prev = String(sh.getRange(row, iNext+1).getValue()||'').trim();
    const text = fu.join(' • ');
    sh.getRange(row, iNext+1).setValue(prev ? (prev + '\n' + text) : text);
  }
}

function diag_clientReportUrlRead(rootApptId){
  const ssId = PROP_('SPREADSHEET_ID');
  Logger.log('SPREADSHEET_ID (Script Property) = %s', ssId);

  const ss = SpreadsheetApp.openById(ssId);
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet "00_Master Appointments"');

  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
  const iRoot = header.indexOf('RootApptID');
  const iURL  = header.indexOf('Client Status Report URL');

  Logger.log('Header indexes: RootApptID=%s  ClientStatusReportURL=%s', iRoot, iURL);
  if (iRoot < 0 || iURL < 0) throw new Error('Header not found exactly (check spelling/spaces)');

  const last = sh.getLastRow();
  const roots = sh.getRange(2, iRoot+1, Math.max(0,last-1), 1).getValues().flat().map(v=>String(v||'').trim());
  const idx = roots.findIndex(v => v === String(rootApptId).trim());

  Logger.log('Found root row index (0-based in data) = %s', idx);
  if (idx < 0) throw new Error('RootApptID not found on THIS master: ' + rootApptId);

  const row = idx + 2;
  const rawUrl = String(sh.getRange(row, iURL+1).getValue() || '');
  Logger.log('Cell raw URL length=%s  value="%s"', rawUrl.length, rawUrl);

  const id = idFromAnyGoogleUrl_(rawUrl);
  Logger.log('Parsed ID = "%s"', id || '(none)');

  if (!id) throw new Error('Cell is blank/invalid for row '+row+'. Paste a valid Sheets URL there.');
}

// === Brand helpers (place ABOVE upsertClientSummaryTab_) ===
function getBrandForRoot_(rootApptId){
  const ss = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  const sh = ss.getSheetByName('00_Master Appointments');
  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
  const iRoot = hdr.indexOf('RootApptID'), iBrand = hdr.indexOf('Brand');
  if (iRoot<0 || iBrand<0) return 'VVS';
  const last = sh.getLastRow(); if (last<2) return 'VVS';
  const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][iRoot]||'').trim() === String(rootApptId).trim()){
      const b = String(vals[i][iBrand]||'').trim().toUpperCase();
      return (b==='HPUSA'||b==='VVS') ? b : 'VVS';
    }
  }
  return 'VVS';
}

// Pantone 213c ≈ #D50057 for HPUSA; VVS = #FFD1DC
function brandAccentHex_(brand){
  const b = String(brand||'').toUpperCase();
  return (b==='HPUSA') ? '#D50057' : '#FFD1DC';
}

function getAssignedRepForRoot_(rootApptId){
  const ss = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  const sh = ss.getSheetByName('00_Master Appointments');
  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
  const iRoot = hdr.indexOf('RootApptID'), iRep = hdr.indexOf('Assigned Rep');
  if (iRoot<0 || iRep<0) return '';
  const last = sh.getLastRow(); if (last<2) return '';
  const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][iRoot]||'').trim() === String(rootApptId).trim()){
      return String(vals[i][iRep]||'').trim();
    }
  }
  return '';
}

function getAssistedRepForRoot_(rootApptId){
  const ss = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  const sh = ss.getSheetByName('00_Master Appointments');
  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
  const iRoot = hdr.indexOf('RootApptID'), iAssist = hdr.indexOf('Assisted Rep');
  if (iRoot<0 || iAssist<0) return '';
  const last = sh.getLastRow(); if (last<2) return '';
  const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][iRoot]||'').trim() === String(rootApptId).trim()){
      return String(vals[i][iAssist]||'').trim();
    }
  }
  return '';
}

function getApptIsoForRoot_(ss, root){
  const sh = ss.getSheetByName('00_Master Appointments');
  const H = (function(row){ const m={}; row.forEach((h,i)=>m[String(h||'').trim()]=i+1); return m; })
           (sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]);
  if (!H['RootApptID']) throw new Error('Missing RootApptID col');
  const last = sh.getLastRow(); if (last < 2) return '';
  for (let r=2;r<=last;r++){
    if (String(sh.getRange(r, H['RootApptID']).getValue()||'').trim() === String(root).trim()){
      const idx = H['ApptDateTime (ISO)'];
      return idx ? String(sh.getRange(r, idx).getValue()||'').trim() : '';
    }
  }
  return '';
}

function newestFileIn_(folder, extRegex){
  let newest=null, ts=0;
  const it = folder.getFiles();
  while (it.hasNext()){
    const f = it.next();
    if (extRegex && !extRegex.test(f.getName())) continue;
    const t = f.getDateCreated().getTime();
    if (t > ts){ newest=f; ts=t; }
  }
  return newest;
}

// Compact object printer for logs (truncates long strings).
function _brief_(obj, max=240) {
  try {
    const o = {};
    Object.keys(obj || {}).forEach(k => {
      const v = String(obj[k]);
      o[k] = v.length > max ? (v.slice(0, max) + ` …(${v.length} chars)`) : v;
    });
    return JSON.stringify(o);
  } catch (_){ return '(unprintable)'; }
}

// Safe byte-length of a post body (works for binary audio).
function _bodyLen_(e){
  try { return (e && e.postData && e.postData.getBytes && e.postData.getBytes().length) || 0; }
  catch(_){ return 0; }
}

function diag_chunkerEndToEnd() {
  const SP = PropertiesService.getScriptProperties();
  const url = SP.getProperty('CHUNKER_URL');
  if (!url) { Logger.log('CHUNKER_URL missing'); return; }

  // 1) DNS / reachability
  try {
    const r = UrlFetchApp.fetch(url + '/diag', { method:'get', muteHttpExceptions:true });
    Logger.log('DIAG: ' + r.getResponseCode() + ' ' + r.getContentText());
  } catch (e) {
    Logger.log('DIAG failed: ' + e);
    return;
  }

  // 2) Minimal POST shape check (no real file; server should 400 with JSON)
  try {
    const r2 = UrlFetchApp.fetch(url + '/chunk', {
      method:'post',
      contentType:'application/json',
      payload: JSON.stringify({ fileId:'ping', destFolderId:'ping', baseName:'ping', chunkSeconds:900 }),
      muteHttpExceptions:true
    });
    Logger.log('CHUNK ping: ' + r2.getResponseCode() + ' ' + r2.getContentText());
  } catch (e2) {
    Logger.log('CHUNK ping failed: ' + e2);
  }
}



function test_diag_clientReportUrlRead(){
diag_clientReportUrlRead('AP-20250910-001');
}

function test_rerenderClientSummaryTabForRoot(){
rerenderClientSummaryTabForRoot_('AP-20250907-003');
}

function test_summarizeLatestTranscript(){
summarizeLatestTranscript('AP-20250907-003');
}

function test_runStrategistAnalysisForRoot(){
runStrategistAnalysisForRoot('AP-20250907-003');
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
