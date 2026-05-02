/***** Upload Endpoint — v1.0 (ACK + Stage + Enqueue) *****
 * Expects a raw binary body (audio) with URL params for metadata:
 *   .../exec?token=...&root_appt_id=AP-YYYYMMDD-###&brand=VVS|HPUSA&rep_email=...&filename=...
 *
 * Script Properties required:
 *   STAGING_FOLDER_ID  -> your [SYS] Upload Staging folder
 *   UPLOAD_TOKEN       -> shared secret from Shortcut (e.g., consult_upload_1063)
 *
 * Writes a row to _upload_queue (PENDING) for a worker to process later.
 */

// === script props ===
const UP_SP = PropertiesService.getScriptProperties();

// === queue config (mirrors your Calendly queue pattern) ===
const UPLOAD_QUEUE_SHEET = '_upload_queue';
const UPLOAD_QUEUE_HEADERS = [
  'ts_enqueued',
  'root_appt_id',
  'brand',
  'rectype',
  'assigned_reps_csv',
  'status',
  'attempts',
  'file_metadata_json',
  'staging_url'
]; // similar shape to your webhook queue with status/attempts fields:contentReference[oaicite:1]{index=1}

function getOrCreateUploadQueue_() {
  const ss = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')
  );
  let sh = ss.getSheetByName(UPLOAD_QUEUE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(UPLOAD_QUEUE_SHEET);
    sh.getRange(1,1,1,UPLOAD_QUEUE_HEADERS.length).setValues([UPLOAD_QUEUE_HEADERS]);
    sh.setFrozenRows(1);
  }
  // self-heal header if someone edits it
  const have = sh.getRange(1,1,1,UPLOAD_QUEUE_HEADERS.length).getValues()[0].map(v=>String(v||'').trim());
  if (UPLOAD_QUEUE_HEADERS.some((h,i)=>h!==have[i])) {
    sh.getRange(1,1,1,UPLOAD_QUEUE_HEADERS.length).setValues([UPLOAD_QUEUE_HEADERS]);
  }
  return sh;
}

function _respOk_(msg) { return ContentService.createTextOutput(String(msg||'OK')); }
function _bad_(why)    { return ContentService.createTextOutput('BAD: ' + String(why||'')).setMimeType(ContentService.MimeType.TEXT); }


function doPost_UPLOAD_(e) {
  const SP = PropertiesService.getScriptProperties();
  const tz = Session.getScriptTimeZone() || 'America/Los_Angeles';

  // [TRACE] — request envelope
  const TRACE = Utilities.getUuid().slice(0, 8);
  const ct0   = e && e.postData ? (e.postData.type || '') : '';
  const len0  = _bodyLen_(e);
  const keys0 = Object.keys((e && e.parameter) || {});
  Logger.log('[UPLOAD %s] START ct=%s bytes=%s paramKeys=%s', TRACE, ct0, len0, keys0.join(',') || '(none)');
  try { Logger.log('[UPLOAD %s] e.parameter = %s',  TRACE, _brief_(e.parameter)); } catch(_){}
  try { Logger.log('[UPLOAD %s] e.parameters = %s', TRACE, _brief_(e.parameters)); } catch(_){}


  // --- quick auth ---
  const tokenParam = String(e?.parameter?.token || '').trim();
  const tokenWant  = String(SP.getProperty('UPLOAD_TOKEN') || '').trim();
  if (!tokenParam || tokenParam !== tokenWant) {
    return ContentService.createTextOutput('ACK (invalid token)');
  }

  // --- quick param sanity ---
  const apId   = String(e?.parameter?.root_appt_id || '').trim();
  let   fname  = String(e?.parameter?.filename || '').trim();
  if (!/^AP-\d{8}-\d{3}$/i.test(apId)) return ContentService.createTextOutput('ACK (bad/missing RootApptID)');
  if (!e?.postData)                    return ContentService.createTextOutput('ACK (empty body)');

  // --- capture + normalize new params (keep raw for diagnostics) ---
  (function(){
    const P = e && e.parameter ? e.parameter : {};

    // raw (exactly as received)
    const rectypeRaw = String(P.rectype || '').trim();
    const brandRaw   = String(P.brand   || '').trim();
    const repsRaw    = String(P.assigned_reps || '').trim();

    // [TRACE] raw values
    Logger.log('[UPLOAD %s] RAW rectype="%s" brand="%s" assigned_reps="%s"', TRACE, rectypeRaw, brandRaw, repsRaw);

    // visible log for 1st mile debugging (safe, small)
    try { Logger.log('UPLOAD PARAMS raw=%s', JSON.stringify({rectypeRaw, brandRaw, repsRaw})); } catch(_){}

    // normalize (defensive; accept "1", "2", "1. Consult", "2. Debrief", or "consult/debrief")
    const rc = rectypeRaw.trim().toLowerCase();          // "1", "2", "1. consult", "consult", etc.
    const stripped = rc.replace(/^\s*\d+[.)]?\s*/,'');   // remove leading "1.", "2)", etc.
    var rectype = (stripped === 'consult' || stripped === 'debrief')
      ? stripped
      : (rc === '1' ? 'consult' : rc === '2' ? 'debrief' : '');

    var brandNorm = brandRaw.toUpperCase();
    brandNorm = (brandNorm === 'HPUSA' || brandNorm === 'VVS') ? brandNorm : '';

    var repsCsv = repsRaw.split(',').map(s=>s.trim()).filter(Boolean).join(',');

    // light clamps
    if (rectype.length > 16)   rectype = rectype.slice(0,16);
    if (brandNorm.length > 8)  brandNorm = brandNorm.slice(0,8);
    if (repsCsv.length > 300)  repsCsv = repsCsv.slice(0,300);

    // [TRACE] normalized values
    Logger.log('[UPLOAD %s] NORM rectype="%s" brand="%s" repsCsv="%s"', TRACE, rectype, brandNorm, repsCsv);

    // attach both normalized + raw for downstream visibility
    e._rectype      = rectype;
    e._brand        = brandNorm;
    e._repsCsv      = repsCsv;
    e._rectype_raw  = rectypeRaw;
    e._brand_raw    = brandRaw;
    e._reps_raw     = repsRaw;
  })();


  // --- build blob once (RAW if Content-Type=audio/*; else octet-stream) ---
  const ct = e.postData.type || '';
  const bytes = e.postData.getBytes();             // single read
  if (!bytes || !bytes.length) return ContentService.createTextOutput('ACK (empty body)');

  // name & mime
  if (!/\.(mp3|mp4|m4a|wav|webm|mpeg|mpga)$/i.test(fname)) {
    const tsIso = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH-mm-ss");
    fname = (apId + '__' + tsIso + '.m4a');       // default
  }
  const mimeMap = { mp3:'audio/mpeg', mpeg:'audio/mpeg', mpga:'audio/mpeg', mp4:'audio/mp4', m4a:'audio/mp4', wav:'audio/wav', webm:'audio/webm' };
  const ext = fname.split('.').pop().toLowerCase();
  const mime = mimeMap[ext] || (/^audio\//i.test(ct) ? ct : 'application/octet-stream');
  const blob = Utilities.newBlob(bytes, mime, fname);

  // --- save to staging (single Drive call) ---
  const stagingId = SP.getProperty('STAGING_FOLDER_ID');
  if (!stagingId) return ContentService.createTextOutput('ACK (no staging folder configured)');
  const file = DriveApp.getFolderById(stagingId).createFile(blob);
  const stagingUrl = 'https://drive.google.com/file/d/' + file.getId() + '/view';

  // [TRACE] staging result
  Logger.log('[UPLOAD %s] STAGED id=%s name="%s" size=%s mime=%s url=%s',
             TRACE, file.getId(), file.getName(), file.getSize(), blob.getContentType(), stagingUrl);

  // --- enqueue (self-heal header, then append full row) ---
  try {
    const ss = SpreadsheetApp.openById(SP.getProperty('SPREADSHEET_ID'));
    const sh = ss.getSheetByName(UPLOAD_QUEUE_SHEET) || ss.insertSheet(UPLOAD_QUEUE_SHEET);

    (function ensureHeader(){
      const have = sh.getRange(1,1,1, Math.max(sh.getLastColumn(), UPLOAD_QUEUE_HEADERS.length))
                    .getValues()[0].map(v => String(v||'').trim());
      const need = UPLOAD_QUEUE_HEADERS;
      const differs = (have.length < need.length) || need.some((h,i) => h !== (have[i]||''));
      if (differs) {
        if (sh.getLastColumn() < need.length) {
          sh.insertColumnsAfter(sh.getLastColumn(), need.length - sh.getLastColumn());
        }
        sh.getRange(1,1,1,need.length).setValues([need]);
        sh.setFrozenRows(1);
      }
    })();

    const meta = {
      apId,
      filename: file.getName(),
      stagingFileId: file.getId(),
      stagingUrl,
      contentType: blob.getContentType(),
      size: file.getSize(),
      receivedAtIso: new Date().toISOString(),
      intakeMode: (/multipart\//i.test(ct) ? 'multipart' : 'raw'),
      // normalized
      brand:   e._brand || '',
      rectype: e._rectype || '',
      assigned_reps_csv: e._repsCsv || '',
      // raw (for diagnostics)
      brand_raw:   e._brand_raw || '',
      rectype_raw: e._rectype_raw || '',
      assigned_reps_raw: e._reps_raw || ''
    };

    // use normalized when present; else show raw in the visible columns (makes debugging immediate)
    const visBrand   = (e._brand   && e._brand.length)   ? e._brand   : (e._brand_raw   || '');
    const visRectype = (e._rectype && e._rectype.length) ? e._rectype : (e._rectype_raw || '');
    const visReps    = (e._repsCsv && e._repsCsv.length) ? e._repsCsv : (e._reps_raw    || '');

    // [TRACE] what we’re about to append
    Logger.log('[UPLOAD %s] VIS brand="%s" rectype="%s" reps="%s"', TRACE, visBrand, visRectype, visReps);

    const row = [
      new Date(),        // ts_enqueued
      apId,              // root_appt_id
      visBrand,          // brand
      visRectype,        // rectype
      visReps,           // assigned_reps_csv
      'PENDING',         // status
      0,                 // attempts
      JSON.stringify(meta),
      stagingUrl
    ];
    sh.appendRow(row);

    Logger.log('[UPLOAD %s] QUEUED → ss="%s" sheet="%s" (cols=%s) rowCount=%s',
        TRACE, ss.getName(), sh.getName(), sh.getLastColumn(), sh.getLastRow());

    // Optional echo mode for one-shot tests: add &debug=1 to your URL
    if (String(e?.parameter?.debug || '') === '1') {
      const out = {
        trace: TRACE,
        saw_params: e.parameter || {},
        normalized: { rectype: e._rectype, brand: e._brand, assigned_reps_csv: e._repsCsv },
        visible:    { rectype: visRectype, brand: visBrand, assigned_reps_csv: visReps },
        meta: meta,
        stagingUrl: stagingUrl
     };
      return ContentService.createTextOutput(JSON.stringify(out, null, 2))
        .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (qErr) {
    Logger.log('[UPLOAD %s] ENQUEUE ERROR: %s', TRACE, (qErr && (qErr.stack || qErr.message)) || qErr);
    // still ACK; iOS shouldn't hang
  }

  return ContentService.createTextOutput('ACK queued (file staged): ' + stagingUrl);
}

/** === CONFIG (reads existing Script Properties via your PROP_ helper) === */
const MASTER_SSID       = PROP_('SPREADSHEET_ID');
const STAGING_FOLDER_ID = PROP_('STAGING_FOLDER_ID');

function canonicalFolderId_(idOrFolder) {
  // Accept Folder object, URL, or raw id
  var id = (typeof idOrFolder === 'string')
    ? ((idOrFolder.match(/\/folders\/([^/?#]+)/) || [,''])[1] || idOrFolder)
    : idOrFolder.getId();

  var token = ScriptApp.getOAuthToken();
  var res = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(id) +
    '?fields=id,mimeType,shortcutDetails&supportsAllDrives=true',
    { method:'get', headers:{ Authorization:'Bearer ' + token }, muteHttpExceptions:true }
  );
  if (res.getResponseCode() !== 200) {
    throw new Error('Folder meta failed for id=' + id + ' → ' + res.getContentText());
  }
  var j = JSON.parse(res.getContentText());
  // If it’s a shortcut, use its targetId; else use id exactly as returned (may include "_")
  return (j.mimeType === 'application/vnd.google-apps.shortcut' && j.shortcutDetails && j.shortcutDetails.targetId)
    ? j.shortcutDetails.targetId
    : j.id;
}

/** Process queue with self-healing + clear error codes
 *  ALWAYS requests chunking by CHUNK_SECONDS (delegated to chunker).
 */
function processUploadQueue() {
  if (!MASTER_SSID || !STAGING_FOLDER_ID) {
    _queueAnnotateAll_('ERROR_NO_CONFIG', 'Missing SPREADSHEET_ID or STAGING_FOLDER_ID in Script Properties', /*retry*/false);
    return;
  }

  let ss;
  try { ss = SpreadsheetApp.openById(MASTER_SSID); }
  catch (e) {
    _queueAnnotateAll_('ERROR_OPEN_MASTER', 'Cannot open Master spreadsheet: ' + e, /*retry*/false);
    return;
  }

  const q = ss.getSheetByName('_upload_queue');
  if (!q || q.getLastRow() < 2) return;

  const rng = q.getDataRange();
  const values = rng.getValues();
  const H = _hdr_(values[0]);

  for (let r = 1; r < values.length; r++) {
    let row = values[r];
    let status = String(row[H.status] || '');
    if (status !== 'PENDING' && status !== 'RETRY') continue;

    // Mark as WORKING immediately so another execution can't pick this up.
    row[H.status] = 'WORKING';
    values[r] = row;
    // Write only this cell to avoid clobbering other rows
    q.getRange(r+1, H.status+1).setValue('WORKING');
    SpreadsheetApp.flush();

    const meta = _parseJsonSafe_(row[H.file_metadata_json]);

    try {
      // --- 1) Resolve RootAppt + destination folder
      const rootApptId = (meta.root_appt_id || meta.root_apptId || meta.root || meta.apId || row[H.root_appt_id] || '').trim();
      if (!rootApptId) throw _hard('ERROR_BAD_META', 'Missing root_appt_id in file_metadata_json');

      const apId = _resolveApFolderId_(ss, rootApptId); // may throw HARD if not found / missing
      const ap  = DriveApp.getFolderById(apId);
      const audioFolder = _getOrCreate_(ap, '01_Audio');

      // --- 2) Locate staged file (prefer ID, else fallback by filename in Staging)
      const file = _resolveStagedFile_(meta);
      if (!file) throw _soft('WAITING_FILE', 'Staged file not found yet (id/name lookup failed)');

      // Build deterministic base name after we have the file id
      const baseName = rootApptId + '__' + String(file.getId()).slice(0, 8);

      // --- 3) Move & stamp status
      file.moveTo(audioFolder);
      _updateMasterAudioStatus_(ss, rootApptId, 'RECEIVED');

      // --- 3b) ALWAYS request chunking by CHUNK_SECONDS; reuse parts if already present
      try {
        const movedFileId   = file.getId();
        const audioFolderId = canonicalFolderId_(audioFolder); // <-- EXACT id (underscore preserved if present)

        console.log(JSON.stringify({
          where: 'pre-chunk',
          movedFileId,
          audioFolderId,
          audioFolderUrl: audioFolder.getUrl()
        }));

        // --- Build a compact hybrid prompt + force language from Master row (used for BOTH branches)
        const mSh  = ss.getSheetByName('00_Master Appointments');
        const mHdr = mSh.getRange(1,1,1,mSh.getLastColumn()).getValues()[0].map(x=>String(x||'').trim());
        const idx  = (h)=> mHdr.indexOf(h);
        const rowIdx    = _findRowByRootApptId_(mSh, idx('RootApptID'), rootApptId);

        const custName   = rowIdx ? _s(mSh.getRange(rowIdx, idx('Customer Name')+1).getValue()) : '';
        const repName    = rowIdx ? _s(mSh.getRange(rowIdx, idx('Assigned Rep')+1).getValue())   : '';
        const styleNotes = rowIdx ? _s(mSh.getRange(rowIdx, idx('Style Notes')+1).getValue())    : '';
        const diamond    = rowIdx ? _s(mSh.getRange(rowIdx, idx('Diamond Type')+1).getValue())   : '';
        const brand      = String(meta.brand || (H.brand!=null? row[H.brand] : '') || '').trim();

        const extraTerms = []
          .concat(diamond ? [diamond] : [])
          .concat(styleNotes.split(/[,/;•|]/).map(s=>s.trim()).filter(Boolean))
          .slice(0, 15);

        const built = buildTranscribePromptV2_({
          extraTerms,
          names: [custName, repName].filter(Boolean),
          notes: brand ? `Brand: ${brand}` : '',
          languageHint: _guessLanguage_([custName, styleNotes].join(' '))
        });
        const prompt   = built.prompt;
        const language = built.language || 'en'; // <-- defined for both branches

        // If parts for this baseName already exist in 01_Audio, reuse them
        function _escapeRe(s){ return s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }
        const re = new RegExp('^' + _escapeRe(baseName) + '__part-\\d{3}\\.m4a$', 'i');
        let existingParts = [];
        (function scanExisting(){
          const it = audioFolder.getFiles();
          while (it.hasNext()){
            const f = it.next();
            if (re.test(f.getName())) existingParts.push({ fileId: f.getId(), name: f.getName() });
          }
        })();

        // Always ask the chunker to split by time; if audio is shorter than CHUNK_SECONDS, it should return null.
        let parts = existingParts.length ? existingParts : null;
        if (!parts) {
          // Pass {force:true} so backend chunks by time even if file size is small.
          parts = chunkIfTooBig_(movedFileId, audioFolderId, baseName, { force: true });
        }

        if (!parts) {
          // --- No parts returned (short recording): transcribe the single file
          const raw = transcribeWithOpenAI_(file, {
            model: 'gpt-4o-mini-transcribe',
            language,
            temperature: 0,
            prompt
          });
          const txt = normalizeTranscript_(raw);
          const url = saveTranscript_(ap, rootApptId, txt);

          // Light sanity check—keep but don't block success
          try {
            const sizeBytes = file.getSize();
            const estSecs   = Math.round((sizeBytes * 8) / (32 * 1000)); // ~32kbps heuristic
            const words     = String(txt||'').trim().split(/\s+/).length;
            if (estSecs > 1200 && words < estSecs * 0.5) { // >20 min but very few words → suspect truncation
              throw _soft('TRANSCRIBE_SUSPECT_TRUNCATION', 'Transcript shorter than expected; retrying with chunking');
            }
          } catch(_) {}

          _updateMasterAudioStatus_(ss, rootApptId, 'TRANSCRIBED');
          setMasterFieldForRoot_(ss, rootApptId, 'Has Transcript', 'TRUE');
          try { nudgeSummariesIfNone(); } catch(_){}
          row[H.file_metadata_json] = _annotateMeta_(row[H.file_metadata_json], {
            transcribed: true, transcriptUrl: url
          });

        } else {
          // --- We have chunked parts: transcribe each, then combine to one .txt
          let combined = [];
          for (let i = 0; i < parts.length; i++) {
            const p = parts[i];                              // {fileId, name}
            const pf = DriveApp.getFileById(p.fileId);       // the part file
            try {
              const partRaw = transcribeWithOpenAI_(pf, { model:'gpt-4o-mini-transcribe', language, temperature:0, prompt });
              const pt = normalizeTranscript_(partRaw);
              combined.push(`=== Part ${i+1} (${p.name}) ===\n` + pt.trim());
            } catch (eEach) {
              // If one part fails, mark for retry
              throw _soft('TRANSCRIBE_PART_RETRY', `Part ${i+1} failed: ` + eEach);
            }
          }

          // Join all parts into one final .txt and save under 03_Transcripts
          const finalTxt = combined.join('\n\n');
          const url = saveTranscript_(ap, rootApptId, finalTxt);

          _updateMasterAudioStatus_(ss, rootApptId, 'TRANSCRIBED');
          setMasterFieldForRoot_(ss, rootApptId, 'Has Transcript', 'TRUE');
          try { nudgeSummariesIfNone(); } catch(_) {}
          row[H.file_metadata_json] = _annotateMeta_(row[H.file_metadata_json], {
            transcribed: true, transcriptUrl: url, chunked: true, parts: parts.length
          });

          // (optional) You can trash the original huge file after successful chunking:
          // DriveApp.getFileById(movedFileId).setTrashed(true);
        }

        // At this point, a transcript .txt exists on Drive. The summarizer worker can run later.
      } catch (eTrans) {
        // Bubble a SOFT error so the outer catch marks RETRY and annotates the row
        throw _soft('TRANSCRIBE_RETRY', String(eTrans));
      }

        // --- 3c) NEW: Inline Scribe → Strategist → SYS_Consults (no waiting on minute worker)
      try {
        // Step 1: Scribe summary (strict JSON). Also sets Audio Status=SUMMARIZED, writes 04_Summaries, and upserts SYS_Consults.
        // Returns { consultId, summaryUrl }.
        const sum = summarizeLatestTranscript(rootApptId);

        // Step 2 (defensive): Strategist memo + strict JSON.
        // NOTE: Your summarizeLatestTranscript() already calls runStrategistAnalysisForRoot() internally.
        // We call it again defensively so older copies still get Strategist. Safe if idempotent.
        let strategistUrl = '';
        let memoUrl = '';
        try {
          const out = runStrategistAnalysisForRoot(rootApptId);
          strategistUrl = out && out.strategistUrl || '';
          memoUrl       = out && out.memoUrl || '';
        } catch (e2) {
          // If summarizeLatestTranscript already triggered Strategist, this second call may be redundant — ignore errors.
          Logger.log('Strategist inline skip/info: ' + (e2 && (e2.message || e2)));
        }

        // Mirror URLs to Master for convenience (best effort).
        try {
          if (sum && sum.summaryUrl)        setMasterFieldForRoot_(ss, rootApptId, 'Summary JSON URL',     sum.summaryUrl);
          if (strategistUrl)                setMasterFieldForRoot_(ss, rootApptId, 'Strategist JSON URL',  strategistUrl);
          try { setAudioStatusFor(rootApptId, 'SUMMARIZED'); } catch(_) {}
        } catch(_){}

        // Annotate queue metadata
        row[H.file_metadata_json] = _annotateMeta_(row[H.file_metadata_json], {
          summarized: true,
          consultId: (sum && sum.consultId) || '',
          summaryUrl: (sum && sum.summaryUrl) || '',
          strategistUrl: strategistUrl || '',
          memoUrl: memoUrl || ''
        });
      } catch (eSum) {
        // Non-fatal; transcript is saved and minute worker can still pick it up.
        Logger.log('Inline summarize failed (will rely on worker/trigger): ' + (eSum && (eSum.stack || eSum.message) || eSum));
        try { nudgeSummariesIfNone(); } catch(_){}
      }


      // --- 4) Done
      row[H.status]   = 'DONE';
      row[H.attempts] = Number(row[H.attempts] || 0) + 1;
      row[H.file_metadata_json] = _annotateMeta_(row[H.file_metadata_json], { ok:true, movedTo:'01_Audio' });
      values[r] = row;

    } catch (err) {
      const attempts = Number(row[H.attempts] || 0) + 1;
      const code = (err && err.code) || 'RETRY';
      const msg  = (err && err.message) || String(err);

      // decide whether to retry
      const retry = (err && err.retry !== false); // HARD errors set retry=false
      row[H.status]   = retry ? 'RETRY' : 'ERROR_GAVE_UP';
      row[H.attempts] = attempts;
      row[H.file_metadata_json] = _annotateMeta_((row[H.file_metadata_json] || ''), { error:code, message:msg, attempts });
      values[r] = row;
    }
  }

  rng.setValues(values);
}

/** ---------- helpers ---------- */
function _hdr_(hdr){
  const m = {};
  hdr.forEach((h,i)=> m[String(h||'').trim()] = i);
  // tolerant defaults for your header names
  if (!('file_metadata_json' in m)) m.file_metadata_json = m['file_metadata_json'] ?? 6;
  if (!('status' in m))            m.status = m['status'] ?? 4;
  if (!('attempts' in m))          m.attempts = m['attempts'] ?? 5;
  if (!('root_appt_id' in m))      m.root_appt_id = m['root_appt_id'] ?? 1;
  return m;
}
function _parseJsonSafe_(s){ try { return JSON.parse(String(s||'{}')); } catch(_){ return {}; } }
function _annotateMeta_(cell, extra){
  let o = _parseJsonSafe_(cell);
  o = Object.assign({}, o, { worker_ts:new Date().toISOString() }, extra||{});
  return JSON.stringify(o);
}
function _soft(code, msg){ const e = new Error(msg); e.code = code; e.retry = true;  return e; }
function _hard(code, msg){ const e = new Error(msg); e.code = code; e.retry = false; return e; }

function _resolveStagedFile_(meta){
  if (meta.stagingFileId) { try { return DriveApp.getFileById(String(meta.stagingFileId)); } catch(_) {} }
  const name = (meta.filename || '').trim();
  if (!name) return null;
  const staging = DriveApp.getFolderById(STAGING_FOLDER_ID);
  const it = staging.getFilesByName(name);
  let newest=null, t=0;
  while (it.hasNext()){ const f = it.next(); const ts = f.getDateCreated().getTime(); if (ts>t){ newest=f; t=ts; } }
  return newest;
}

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

function _getOrCreate_(parent, name){
  const it = parent.getFoldersByName(name);
  return it.hasNext()? it.next() : parent.createFolder(name);
}

function _updateMasterAudioStatus_(ss, rootApptId, status){
  const sh = ss.getSheetByName('00_Master Appointments'); if (!sh) return;
  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x=>String(x||'').trim());
  const iRoot = hdr.indexOf('RootApptID');

  // find or create the Audio Status column
  let iAS = (function(){
    const aliases = ['Audio Status','AudioStatus','Audio status','Status (Audio)'];
    for (const name of aliases){
      const idx = hdr.indexOf(name);
      if (idx >= 0) return idx;
    }
    // create a new column at the end if none exist
    sh.insertColumnAfter(sh.getLastColumn());
    const col = sh.getLastColumn();
    sh.getRange(1, col).setValue('Audio Status');
    return col-1; // 0-based
  })();

  if (iRoot < 0) return;
  const last = sh.getLastRow(); if (last < 2) return;

  // scan rows to find matching RootApptID and set status
  for (let r = 2; r <= last; r++){
    const v = String(sh.getRange(r, iRoot+1).getValue() || '').trim();
    if (v === String(rootApptId).trim()){
      sh.getRange(r, iAS+1).setValue(status); // setValue (single row) to avoid big-range overwrites
      return;
    }
  }
}


/** If config is missing or master can’t be opened, annotate all PENDING/RETRY and stop. */
function _queueAnnotateAll_(code, message, retry){
  try{
    const ss = MASTER_SSID ? SpreadsheetApp.openById(MASTER_SSID) : null;
    const sh = ss ? ss.getSheetByName('_upload_queue') : null;
    if (!sh) return;
    const rng = sh.getDataRange(); const vals = rng.getValues(); const H = _hdr_(vals[0]);
    for (let r=1; r<vals.length; r++){
      const st = String(vals[r][H.status]||'');
      if (st !== 'PENDING' && st !== 'RETRY') continue;
      vals[r][H.status]   = retry ? 'RETRY' : 'ERROR_GAVE_UP';
      vals[r][H.attempts] = Number(vals[r][H.attempts]||0) + 1;
      vals[r][H.file_metadata_json] = _annotateMeta_(vals[r][H.file_metadata_json], { error:code, message, attempts:vals[r][H.attempts] });
    }
    rng.setValues(vals);
  }catch(_){ /* best effort */ }
}

function debug_traceLastDoneUpload(){
  const ss = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  const q  = ss.getSheetByName('_upload_queue');
  if (!q) { Logger.log('No _upload_queue'); return; }
  const vals = q.getDataRange().getValues();
  const H = vals[0].reduce((m,h,i)=> (m[String(h||'').trim()] = i, m), {});
  // find last DONE
  let rDone = 0;
  for (let r = vals.length-1; r >= 1; r--){
    if (String(vals[r][H.status]||'') === 'DONE') { rDone = r; break; }
  }
  if (!rDone) { Logger.log('No DONE rows yet'); return; }

  const meta = (function(){ try { return JSON.parse(String(vals[rDone][H.file_metadata_json]||'{}')); } catch(_){ return {}; } })();
  const root = String(vals[rDone][H.root_appt_id] || meta.root_appt_id || '').trim();
  const fileName = String(meta.filename || '').trim();
  const apId = (function(){
    const sh = ss.getSheetByName('00_Master Appointments');
    const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x=>String(x||'').trim());
    const iRoot = hdr.indexOf('RootApptID');
    const iFld  = hdr.indexOf('RootAppt Folder ID');
    if (iRoot < 0 || iFld < 0) throw new Error('Missing RootApptID / RootAppt Folder ID columns');
    const last = sh.getLastRow();
    const rows = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
    for (let i=0;i<rows.length;i++){
      if (String(rows[i][iRoot]||'').trim() === root) return String(rows[i][iFld]||'').trim();
    }
    throw new Error('RootApptID not found in Master: '+root);
  })();

  const ap = DriveApp.getFolderById(apId);
  const audio = (function(){
    const it = ap.getFoldersByName('01_Audio');
    return it.hasNext() ? it.next() : null;
  })();
  if (!audio) { Logger.log('No 01_Audio folder under '+ap.getName()+' '+ap.getUrl()); return; }

  Logger.log('RootApptID: %s', root);
  Logger.log('01_Audio URL: %s', audio.getUrl());

  // verify file by name
  if (fileName){
    const it = audio.getFilesByName(fileName);
    let count = 0, fid = '';
    while (it.hasNext()){ const f = it.next(); count++; fid = f.getId(); }
    Logger.log('Found %s file(s) named "%s" in 01_Audio%s',
               count, fileName, count ? ' (example: https://drive.google.com/file/d/'+fid+'/view)' : '');
  } else {
    Logger.log('filename missing in metadata; listing first few files in 01_Audio:');
    const it = audio.getFiles(); let n=0;
    while (it.hasNext() && n<5){ const f = it.next(); Logger.log('- %s', f.getName()); n++; }
  }
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

/** === OpenAI config via Script Properties === */
function OPENAI_(k){ return PropertiesService.getScriptProperties().getProperty(k) || ''; }
const OPENAI_API_KEY = OPENAI_('OPENAI_API_KEY');  // <-- set this in Script Properties

function _multipart_(parts){
  const boundary = '----apps-script-boundary-' + Date.now();
  const blobs = [];
  const push = (s)=> blobs.push(Utilities.newBlob(s));
  const CRLF = '\r\n';
  parts.forEach(p=>{
    push('--'+boundary+CRLF);
    if(p.type==='file'){
      push('Content-Disposition: form-data; name="'+p.name+'"; filename="'+(p.filename||'file')+'"'+CRLF);
      push('Content-Type: '+(p.contentType||'application/octet-stream')+CRLF+CRLF);
      blobs.push(p.blob); push(CRLF);
    }else{
      push('Content-Disposition: form-data; name="'+p.name+'"'+CRLF+CRLF);
      push(String(p.value||'')+CRLF);
    }
  });
  push('--'+boundary+'--'+CRLF);
  return {
    payload: blobs.reduce((bytes, b) => bytes.concat(b.getBytes()), []),
    contentType: 'multipart/form-data; boundary='+boundary
  };
}

/**
 * Transcribe an audio file with OpenAI's transcription models.
 *
 * @param {GoogleAppsScript.Drive.File} file - A Drive file object (mp3, m4a, wav, etc.)
 * @param {Object} opts - Optional overrides {model, language, temperature, prompt}
 * @returns {string} transcript text
 */
function transcribeWithOpenAI_(file, opts) {
  const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) throw new Error('Missing OPENAI_API_KEY');

  // --- Normalize filename + MIME ---
  let name = String(file.getName() || '').replace(/"/g, '').trim();
  if (!/\.(mp3|mp4|m4a|wav|webm|mpeg|mpga)$/i.test(name)) {
    name = (name || 'audio') + '.m4a';
  }
  const ext = name.split('.').pop().toLowerCase();
  const mimeMap = {
    mp3:  'audio/mpeg',
    mpeg: 'audio/mpeg',
    mpga: 'audio/mpeg',
    mp4:  'audio/mp4',
    m4a:  'audio/mp4',
    wav:  'audio/wav',
    webm: 'audio/webm'
  };
  const mime = mimeMap[ext] || 'audio/mp4';

  // --- Build the blob with clean name + type ---
  const audioBlob = file.getBlob().setName(name).setContentType(mime);

  // --- Defaults & overrides ---
  const model       = (opts && opts.model)       || 'gpt-4o-mini-transcribe';
  const language    = (opts && opts.language)    || 'en';     // e.g. 'en', 'vi'
  const temperature = (opts && opts.temperature !== undefined) ? opts.temperature : 0;
  const prompt      = (opts && opts.prompt)      || '';

  // --- Call API ---
  const payload = {
    model,
    file: audioBlob,
    temperature
  };
  if (language) payload.language = language;
  if (prompt)   payload.prompt   = prompt;

  const res = UrlFetchApp.fetch('https://api.openai.com/v1/audio/transcriptions', {
    method: 'post',
    headers: { Authorization: 'Bearer ' + key },
    payload,
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code !== 200) {
    throw new Error('OpenAI ' + code + ': ' + text);
  }

  const body = JSON.parse(text);
  return String(body.text || '').trim();
}


/****************************
 * Transcribe Prompt Builder v2 (Hybrid)
 ****************************/

// === Fixed core: keep this short and stable (≈120–180 tokens) ===
const FIXED_CORE_PROMPT = [
  "You are a verbatim transcriber for luxury fine-jewelry sales calls.",
  "Rules:",
  "• Transcribe exactly what is spoken (no rewriting).",
  "• Use US jewelry terminology and spellings.",
  "• Prefer these exact forms & casing:",
  "  VVS Jewelry Co., Hung Phat USA (HPUSA), San Jose, 18K, 14K, GIA, IGI, VVS2, VS1, VS2, SI1, SI2, IF, pavé, halo, solitaire, bezel, cathedral, prong, knife-edge.",
  "• Diamonds: write carat (ct), color (D–Z), clarity (IF–I), cut, polish, symmetry, fluorescence.",
  "• Materials: platinum, white gold, yellow gold, rose gold.",
  "• Shapes: round brilliant, oval, cushion, emerald, radiant, pear, marquise, princess.",
  "• Phrases to disambiguate:",
  "  carat (diamond weight) ≠ carrot (vegetable).",
  "  pavé has an accent (pavé).",
  "  halo (ring style) ≠ hello.",
  "• If letters are spelled (e.g., I-G-I), keep them as uppercase abbreviations (IGI).",
  "• Output plain text only."
].join("\n");

// === Your existing glossaries (trimmed) ===
const BASE_GLOSSARY = [
  'VVS Jewelry Co.','Hung Phat USA','HPUSA','San Jose',
  'carat','clarity','color grade','cut grade','polish','symmetry','fluorescence',
  'VVS2','VS1','VS2','SI1','SI2','IF','GIA','IGI',
  'platinum','18K','14K','white gold','yellow gold','rose gold',
  'halo','hidden halo','solitaire','cathedral','bezel','channel set','prong','micro-pavé','pavé','knife-edge','tapered shank',
  'round brilliant','oval','cushion','emerald','radiant','pear','marquise','princess',
  'lab-grown diamond','natural diamond','HPUSA in-house','custom CAD','wax model'
];

const KNOWN_NAMES = ['Vivianne','Val','Summer','Wendy','Phụng Minh','Kris','Tường Vân','An'];

// Utilities
function _s(v){ return (v==null?'':String(v)).trim(); }
function _dedupKeepOrder(arr){
  const seen = new Set(); const out = [];
  for (const x of arr) { const k = x.toLowerCase(); if (!seen.has(k)) { seen.add(k); out.push(x); } }
  return out;
}
function _truncateWords(s, max){
  const w = s.split(/\s+/).filter(Boolean);
  return w.length <= max ? s : w.slice(0, max).join(' ');
}

function _words(s){ return String(s||'').split(/\s+/).filter(Boolean); }
function _truncateTokens(s, maxTokens){
  const w = _words(s);
  return (w.length <= maxTokens) ? s : w.slice(0, maxTokens).join(' ');
}


// Light language guess (your existing)
function _guessLanguage_(blob){
  return /[ăâđêôơưáàảãạắằẳẵặấầẩẫậéèẻẽẹếềểễệíìỉĩịóòỏõọốồổỗộớờởỡợúùủũụứừửữựýỳỷỹỵ]/i.test(blob) ? 'vi' : 'en';
}

/** Build hybrid prompt
 * opts = {
 *   extraTerms: []        // session-specific nouns/brands/models (5–15 max)
 *   names: []             // person names heard/expected (3–8 max)
 *   notes: ""             // a couple short style notes (<= 30 words)
 *   languageHint: "en"|"vi" // if you already know, pass it
 * }
 */
function buildTranscribePromptV2_(opts={}){
  const names = _dedupKeepOrder((opts.names||[]).concat(KNOWN_NAMES)).slice(0, 8);
  const terms = _dedupKeepOrder((opts.extraTerms||[]).concat(BASE_GLOSSARY)).slice(0, 40);

  // Tiny dynamic add-on. Keep this compact so it doesn’t drown the core.
  const dynamic = [
    names.length ? "Names to prefer: " + names.join(", ") + "." : "",
    terms.length ? "Domain words to prefer: " + terms.join(", ") + "." : "",
    _s(opts.notes) ? "Context notes: " + _truncateWords(_s(opts.notes), 30) : ""
  ].filter(Boolean).join("\n");

  const prompt = [FIXED_CORE_PROMPT, dynamic].join("\n\n");
  return {
    prompt,
    language: opts.languageHint || 'en'
  };
}

function normalizeTranscript_(txt){
  if (!txt) return '';

  // Case-sensitive brand/codes
  const hardRepls = [
    [/\bhp ?usa\b/gi, 'HPUSA'],
    [/\bigi\b/g, 'IGI'],
    [/\bgia\b/g, 'GIA'],
    [/\bvv?s2\b/gi, 'VVS2'],
  ];

  // Homophones/contextual
  const softRepls = [
    // carat vs carrot: only fix near diamond keywords
    [/\bcarrot(s?)\b(?=.{0,20}\b(diamond|ring|stone|ct|carat)\b)/gi, 'carat$1'],
    // halo vs hello when followed by ring words
    [/\bhello\b(?=\s+(setting|style|ring|design))/gi, 'halo'],
    // pavé vs pave; prefer pavé when around setting words
    [/\bpave\b(?=\s+(setting|band|style|micro))/gi, 'pavé'],
    // karat for gold (K) vs carat (ct) — if number + K nearby, use 14K/18K
    [/\b(14|18)\s?k(ar)?at\b/gi, (m,n)=> `${n}K`],
  ];

  for (const [re, val] of hardRepls) txt = txt.replace(re, val);
  for (const [re, val] of softRepls) txt = txt.replace(re, val);

  // Tidy repeated spaces and normalize common accents
  txt = txt.replace(/\s{2,}/g, ' ');
  return txt;
}


/**
 * Build a concise prompt for a given Master row.
 * @param {number} rowIdx 1-based row index in "00_Master Appointments"
 * @returns {string}
 */
function buildTranscriptionPromptForRow_(rowIdx){
  const sh = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID')).getSheetByName(TRANSCRIBE_SHEET);
  if (!sh || rowIdx < 2) return FIXED_CORE_PROMPT;

  const brand     = _s(_cellByHeader_(sh, rowIdx, 'Brand')) || 'VVS';
  const custName  = _s(_cellByHeader_(sh, rowIdx, 'Customer Name'));
  const rep       = _s(_cellByHeader_(sh, rowIdx, 'Assigned Rep'));
  const diamond   = _s(_cellByHeader_(sh, rowIdx, 'Diamond Type'));
  const styles    = _s(_cellByHeader_(sh, rowIdx, 'Style Notes'));

  const extra = []
    .concat(diamond ? [diamond] : [])
    .concat(styles.split(/[,/;•|]/).map(s=>s.trim()).filter(Boolean))
    .slice(0, 15);

  const { prompt } = buildTranscribePromptV2_({
    extraTerms: extra,
    names: [custName, rep].filter(Boolean),
    notes: brand
  });
  return prompt;
}


/**
 * Optional helper: build the full payload for a given row + audio blob.
 * You can call this, or just call buildTranscriptionPromptForRow_ yourself.
 */
function buildTranscriptionRequest_(rowIdx, audioBlob, apiKey, opts){
  const sh = SpreadsheetApp.getActive().getSheetByName(TRANSCRIBE_SHEET);
  const notes = _s(_cellByHeader_(sh, rowIdx, 'Style Notes'));
  const guessLang = _guessLanguage_([_s(_cellByHeader_(sh, rowIdx, 'Customer Name')), notes].join(' '));
  const language = (opts && opts.language) || guessLang;       // override if you know it
  const model = (opts && opts.model) || 'gpt-4o-mini-transcribe';
  const temperature = (opts && (opts.temperature !== undefined)) ? opts.temperature : 0;

  const prompt = buildTranscriptionPromptForRow_(rowIdx);

  return {
    url: 'https://api.openai.com/v1/audio/transcriptions',
    params: {
      method: 'post',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: {
        model,
        file: audioBlob,
        language,            // 'en' or 'vi' (or leave out to auto-detect)
        temperature,         // keep at 0 for stability
        prompt               // compact, high-signal vocabulary
      },
      muteHttpExceptions: true
    }
  };
}


function saveTranscript_(ap, rootApptId, text){
  const folder = ap.getFoldersByName('03_Transcripts').hasNext()
    ? ap.getFoldersByName('03_Transcripts').next()
    : ap.createFolder('03_Transcripts');
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH-mm");
  const file = folder.createFile(Utilities.newBlob(text, 'text/plain', rootApptId+'__'+ts+'__full.txt'));
  return file.getUrl();
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
 * Phase 3 — Summarizer Worker *
 * - Reads transcript .txt
 * - Calls OpenAI Responses API with strict JSON Schema
 * - Saves JSON to 04_Summaries/
 * - Upserts row in SYS_Consults by ConsultID
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


/** Phase-3 minute worker: summarize any TRANSCRIBED consults missing Summary JSON URL */
function processSummariesWorker(){
  const ss = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  const sh = ss.getSheetByName('00_Master Appointments');
  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(s=>String(s||'').trim());
  const idx = (h)=>hdr.indexOf(h);
  const iRoot = idx('RootApptID'), iAudio = idx('Audio Status'), iNotes = idx('Automation Notes');
  if (iRoot<0 || iAudio<0) return;

  const last = sh.getLastRow(); if (last<2) return;
  for (let r = 2; r <= last; r++) {
    const status = String(sh.getRange(r, iAudio+1).getValue() || '').trim();
    if (/^TRANSCRIBED$/i.test(status)) {
      const root = String(sh.getRange(r, iRoot+1).getValue() || '').trim();
      try {
        // Scribe → saves JSON → upserts SYS_Consults → runs Strategist → flips Audio Status to SUMMARIZED
        summarizeLatestTranscript(root);
      } catch (e) {
        if (iNotes >= 0) {
          const prev = String(sh.getRange(r, iNotes+1).getValue() || '');
          sh.getRange(r, iNotes+1).setValue(prev ? prev + '\n' + e : String(e));
        }
      }
    }
  }
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

function chunkIfTooBig_(fileId, destFolderId, baseName) {
  const url  = PropertiesService.getScriptProperties().getProperty('CHUNKER_URL');
  const secs = Number(PropertiesService.getScriptProperties().getProperty('CHUNK_SECONDS') || '900');
  const overlap = Number(PropertiesService.getScriptProperties().getProperty('CHUNK_OVERLAP_SECONDS') || '5');
  if (!url) throw new Error('Missing CHUNKER_URL');

  // --- Pre-warm the Render service to avoid cold-start "Address unavailable" ---
  // (Lightweight GET; ignore failures — this is only to wake the container.)
  try {
    UrlFetchApp.fetch(url + '/diag', {
      method: 'get',
      muteHttpExceptions: true
      // No Authorization header needed for /diag; it’s a health check
    });
  } catch (_) { /* best effort warm-up */ }

  const payload = { fileId, destFolderId, baseName, chunkSeconds: secs, overlapSeconds: overlap };
  const token   = ScriptApp.getOAuthToken();

  console.log(JSON.stringify({ CHUNKER_URL: url, fileId, destFolderId, baseName, secs }, null, 2));
  const res = UrlFetchApp.fetch(url + '/chunk', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    throw new Error('Chunker error ' + res.getResponseCode() + ': ' + res.getContentText());
  }
  const out = JSON.parse(res.getContentText());
  if (!out.ok) throw new Error('Chunker bad response: ' + res.getContentText());

  // Allow "no split" for short recordings.
  return (out.parts && out.parts.length) ? out.parts : null; // [{fileId, name}, ...] | null
}


function _test_transcribe_single(){
  // Put an existing small .m4a file ID from Drive here (the one in your staging row works too)
  const id = '1rmFCHJ0vwfaOnme2djLpPsKmQFO6tMCw';  // example from your row
  const txt = transcribeWithOpenAI_(DriveApp.getFileById(id));
  Logger.log(txt || '(empty)');
}

function debug_transcribeWithOpenAI_(file) {
  // --- Build normalized name + mime (same logic as prod fn) ---
  let name = String(file.getName() || '').replace(/"/g, '').trim();
  if (!/\.(mp3|mp4|m4a|wav|webm|mpeg|mpga)$/i.test(name)) {
    name = (name || 'audio') + '.m4a';
  }
  const ext = name.split('.').pop().toLowerCase();
  const mimeMap = {
    mp3:  'audio/mpeg',
    mpeg: 'audio/mpeg',
    mpga: 'audio/mpeg',
    mp4:  'audio/mp4',
    m4a:  'audio/mp4',
    wav:  'audio/wav',
    webm: 'audio/webm'
  };
  const mime = mimeMap[ext] || 'audio/mp4';

  const blob = file.getBlob().setName(name).setContentType(mime);
  const bytes = blob.getBytes();

  // --- Simple container sanity check (ftyp for MP4/M4A) ---
  let hasFtyp = false;
  for (let i = 0; i < Math.min(bytes.length - 3, 256); i++) {
    const str = String.fromCharCode(bytes[i], bytes[i+1], bytes[i+2], bytes[i+3]);
    if (str === 'ftyp') { hasFtyp = true; break; }
  }

  Logger.log(JSON.stringify({
    name: blob.getName(),
    mime: blob.getContentType(),
    size_bytes: bytes.length,
    hasFtyp
  }, null, 2));

  try {
    const text = transcribeWithOpenAI_(file);
    Logger.log('Transcript (first 200 chars): ' + text.substring(0, 200));
    return text;
  } catch (e) {
    Logger.log('Transcription error: ' + (e.stack || e));
    throw e;
  }
}

function runDebugOnce() {
  const file = DriveApp.getFileById('1MJi41RXNxyL8Y0qEuzZ1t8op3GTk2vZ7');
  debug_transcribeWithOpenAI_(file);
}

function debugLog_(msg) {
  try {
    const text = (typeof msg === 'string') ? msg : JSON.stringify(msg);
    const id = '1pH7LpnsoY0BKTut-BqN1E-szgCJO9ScdIBwCKu8nbm4'; // Log sheet ID
    const ss = SpreadsheetApp.openById(id);
    const sh = ss.getSheetByName('Log') || ss.insertSheet('Log');

    // prevent concurrent writes from multiple doPost calls
    const lock = LockService.getDocumentLock();
    lock.tryLock(3000);
    try {
      sh.appendRow([new Date(), text]);
      SpreadsheetApp.flush();
    } finally {
      try { lock.releaseLock(); } catch (_){}
    }
  } catch (err) {
    // never throw from logger
    console.log('debugLog_ failed: ' + (err && (err.stack || err.message) || err));
  }
}

function authorizeOnce() {
  // triggers Drive + Sheets scopes and prompts consent once
  const id = '1SZbiNyWbcjhbtXiJ8tVfCKWAYSPING2eCMm3fXmNulk'; // your Log sheet
  const ss = SpreadsheetApp.openById(id);
  const sh = ss.getSheetByName('Log') || ss.insertSheet('Log');
  sh.appendRow([new Date(), 'authorized']);
}

function logProps(){
  const p = PropertiesService.getScriptProperties();
  Logger.log('UPLOAD_TOKEN = "%s"', (p.getProperty('UPLOAD_TOKEN') || '(missing)'));
  Logger.log('WebAppURL    = %s', ScriptApp.getService().getUrl());
}

function diagChunkerAuth() {
  const url = PropertiesService.getScriptProperties().getProperty("CHUNKER_URL") + "/diag";
  const token = ScriptApp.getOAuthToken();
  const res = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { Authorization: "Bearer " + token }
  });
  Logger.log(res.getContentText());
}

function debug_createTinyInAudioFolder_supportsAllDrives() {
  var folderId = '1cszplDkzfQCOz-YKpJRCrJty5f21JRR';
  var token = ScriptApp.getOAuthToken();
  var boundary = 'b' + Date.now();

  var meta = { name: 'tiny.txt', parents: [folderId] };
  var body = Utilities.newBlob(
    [
      '--' + boundary,
      '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n',
      JSON.stringify(meta),
      '\r\n--' + boundary,
      '\r\nContent-Type: text/plain\r\n\r\n',
      'hello',
      '\r\n--' + boundary + '--'
    ].join('')
  ).getBytes();

  // ✅ NOTE supportsAllDrives=true
  var res = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true',
    {
      method: 'post',
      contentType: 'multipart/related; boundary=' + boundary,
      headers: { Authorization: 'Bearer ' + token },
      payload: body,
      muteHttpExceptions: true
    }
  );
  Logger.log(res.getResponseCode() + ' ' + res.getContentText());
}

function debug_probeAudioFolder() {
  // Paste the URL you opened for 01_Audio (the one that shows the trailing underscore)
  var url = 'https://drive.google.com/drive/folders/1cszplDkzfQCOz-YKpJRCrJty5f21JRR_';
  var idRaw = (url.match(/\/folders\/([^/?#]+)/) || [,''])[1];
  var idTrim = idRaw.replace(/_+$/,''); // try without trailing underscores

  Logger.log('Testing IDs:\n- raw   = %s\n- trim  = %s', idRaw, idTrim);

  function getMeta(id) {
    var token = ScriptApp.getOAuthToken();
    var res = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(id) +
      '?fields=id,name,mimeType,shortcutDetails,parents,driveId&supportsAllDrives=true',
      { method:'get', headers:{ Authorization:'Bearer ' + token }, muteHttpExceptions:true }
    );
    return { code: res.getResponseCode(), text: res.getContentText() };
  }

  var a = getMeta(idRaw);
  var b = getMeta(idTrim);
  Logger.log('RAW  %s %s', a.code, a.text);
  Logger.log('TRIM %s %s', b.code, b.text);

  // If one responds 200 and mimeType is shortcut, resolve to its targetId and test again
  function resolveIfShortcut(text) {
    try {
      var j = JSON.parse(text);
      if (j && j.mimeType === 'application/vnd.google-apps.shortcut' && j.shortcutDetails && j.shortcutDetails.targetId) {
        return j.shortcutDetails.targetId;
      }
    } catch (_){}
    return '';
  }

  var rawTarget  = resolveIfShortcut(a.text);
  var trimTarget = resolveIfShortcut(b.text);

  if (rawTarget) {
    var r2 = getMeta(rawTarget);
    Logger.log('RAW target (%s) -> %s %s', rawTarget, r2.code, r2.text);
  }
  if (trimTarget) {
    var t2 = getMeta(trimTarget);
    Logger.log('TRIM target (%s) -> %s %s', trimTarget, t2.code, t2.text);
  }
}

function installWorker() {
  const fn = 'processUploadQueue';
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === fn);
  if (!exists) ScriptApp.newTrigger(fn).timeBased().everyMinutes(1).create();
}

function bootstrapApptFolder_(rowIdx){
  if (typeof rowIdx !== 'number' || !isFinite(rowIdx) || rowIdx < 2) {
    throw new Error('bootstrapApptFolder_: invalid rowIdx = ' + rowIdx);
  }

  const ss = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet: 00_Master Appointments');

  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(s=>String(s||'').trim());
  const iRoot  = hdr.indexOf('RootApptID');
  const iFldId = hdr.indexOf('RootAppt Folder ID');
  if (iRoot < 0 || iFldId < 0) throw new Error('Missing RootApptID / RootAppt Folder ID columns');

  const rowVals = sh.getRange(rowIdx, 1, 1, sh.getLastColumn()).getValues()[0];
  const root    = String(rowVals[iRoot]||'').trim();

  // choose a parent folder (Prospect URL preferred, else Script Property)
  const iPros = hdr.indexOf('Prospect Folder URL');
  let parentFolder = null;
  if (iPros >= 0) {
    const prospectUrl = String(rowVals[iPros]||'').trim();
    const pid = (prospectUrl.match(/[-\w]{25,}/)||[])[0];
    if (pid) { try { parentFolder = DriveApp.getFolderById(pid); } catch(_) {} }
  }
  if (!parentFolder) {
    const rootId = PROP_('APPTS_ROOT_FOLDER_ID');
    if (!rootId) throw new Error('APPTS_ROOT_FOLDER_ID not configured and no Prospect Folder URL');
    parentFolder = DriveApp.getFolderById(rootId);
  }

  const iClient = hdr.indexOf('Customer Name');
  const client  = iClient>=0 ? String(rowVals[iClient]||'').trim() : '';
  const folderName = client ? `${root} — ${client}` : root;

  const ap = parentFolder.createFolder(folderName);
  ['01_Audio','02_Design','03_Transcripts','04_Summaries'].forEach(n => ap.createFolder(n));

  sh.getRange(rowIdx, iFldId+1).setValue(ap.getId());
}

function nudgeSummariesIfNone() {
  const fn = 'processSummariesWorker';
  const hasAny = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === fn);
  if (!hasAny) {
    // fire once ~5s later; the regular minute trigger will do the steady-state work
    ScriptApp.newTrigger(fn).timeBased().after(5000).create();
  }
}

function installSummariesMinuteWorker(){
  const fn = 'processSummariesWorker';
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === fn)) {
    ScriptApp.newTrigger(fn).timeBased().everyMinutes(1).create();
  }
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



