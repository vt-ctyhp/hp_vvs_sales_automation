/** true if DEBUG_STRATEGIST script property is truthy */
function STRAT_DEBUG_ON_() {
  return /^(1|true|yes)$/i.test(PROP_('DEBUG_STRATEGIST',''));
}

/** Write a JSON blob into 04_Summaries for forensic inspection. */
function strat_writeDebug_(ap, rootApptId, nameSuffix, objOrText) {
  if (!STRAT_DEBUG_ON_()) return '';
  try {
    const sf = ap.getFoldersByName('04_Summaries').hasNext()
      ? ap.getFoldersByName('04_Summaries').next()
      : ap.createFolder('04_Summaries');
    const ts = Utilities.formatDate(new Date(), DEFAULT_TZ_(), "yyyy-MM-dd'T'HH-mm-ss");
    const name = rootApptId + '__debug_' + nameSuffix + '_' + ts + '.json';
    const text = (typeof objOrText === 'string') ? objOrText : JSON.stringify(objOrText, null, 2);
    const f = sf.createFile(Utilities.newBlob(text, 'application/json', name));
    return f.getUrl();
  } catch(e) {
    Logger.log('strat_writeDebug_ failed: ' + (e && (e.stack || e.message) || e));
    return '';
  }
}

/** Basic sanity: does an object look like a real Strategist doc? */
function strat_hasExpectedKeys_(o) {
  if (!o || typeof o !== 'object') return false;
  const must = ['recommended_play','ask_now','today_action','executive_summary'];
  return must.every(k => k in o);
}


/* ---------- [SCRIBE/SAVE] Save corrected Scribe JSON ---------- */
function saveCorrectedScribeJson_(ap, rootApptId, obj){
  const it = ap.getFoldersByName('04_Summaries');
  const sf = it.hasNext() ? it.next() : ap.createFolder('04_Summaries');
  const ts = Utilities.formatDate(new Date(), DEFAULT_TZ_(), "yyyy-MM-dd'T'HH-mm");
  const name = rootApptId + '__summary_corrected_' + ts + '.json';
  const file = sf.createFile(
    Utilities.newBlob(JSON.stringify(obj, null, 2), 'application/json', name)
  );
  try { writeScribeLatestPointer_(ap, file.getId(), Date.now()); } catch(_){}
  return 'https://drive.google.com/file/d/' + file.getId() + '/view';
}


function writeScribeLatestPointer_(ap, fileId, version){
  try{
    const name = 'scribe.latest.json';
    const payload = { fileId: String(fileId || '').trim(), version: Number(version) || Date.now() };
    if (!payload.fileId) return;

    const it = ap.getFilesByName(name);
    if (it.hasNext()){
      it.next().setContent(JSON.stringify(payload, null, 2));
    } else {
      ap.createFile(Utilities.newBlob(JSON.stringify(payload, null, 2), 'application/json', name));
    }
  } catch(_){}
}

/** Save summary JSON into 04_Summaries and return the Drive URL */
function saveSummaryJson_(ap, rootApptId, obj){
  const sFolder = ap.getFoldersByName('04_Summaries').hasNext()
    ? ap.getFoldersByName('04_Summaries').next()
    : ap.createFolder('04_Summaries');
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH-mm");
  const name = rootApptId + '__summary_' + ts + '.json';
  const file = sFolder.createFile(Utilities.newBlob(JSON.stringify(obj, null, 2), 'application/json', name));
  try { writeScribeLatestPointer_(ap, file.getId(), Date.now()); } catch(_){}
  return file.getUrl();
}


/* ---------- [STRATEGIST] Fallback save if you don’t already have saveStrategistJson_ ---------- */
function saveStrategistJson_Fallback_(ap, rootApptId, strategistObj){
  const sf = ap.getFoldersByName('04_Summaries').hasNext()
    ? ap.getFoldersByName('04_Summaries').next()
    : ap.createFolder('04_Summaries');
  const ts = Utilities.formatDate(new Date(), DEFAULT_TZ_(), "yyyy-MM-dd'T'HH-mm");
  const name = `${rootApptId}__analysis_${ts}.json`;
  const file = sf.createFile(
    Utilities.newBlob(JSON.stringify(strategistObj, null, 2), 'application/json', name)
  );
  return 'https://drive.google.com/file/d/' + file.getId() + '/view';
}

/* ---------- [STRATEGIST] Fallback LLM call if buildStrategistPayload_ / openAIResponses_ not present ---------- */
function callOpenAI_StrategistFallback_(scribeObj, transcript, overrideMemo){
  const apiKey = PROP_('OPENAI_API_KEY','');
  const model  = OPENAI_MODEL_();
  if (!apiKey) throw new Error('OPENAI_API_KEY missing; cannot re-run Strategist');

  const sys = [
    "You are a senior luxury fine-jewelry strategist and closer.",
    "Return JSON ONLY with keys:",
    "  recommended_play (string), ask_now (string), today_action (string),",
    "  executive_summary (array of strings), viewing_lineup (array), viewing_strategy (array),",
    "  close_sequence (array), close_strategy (array), top_objections (array of {objection,reply}),",
    "  gaps_risks_struct (array of {risk,mitigation}), budget_alignment (object),",
    "  tradeoffs (array of {option_label,rationale,impact_on_look,risks,price_low,price_high}),",
    "  where_customer_stands_narrative (string)."
  ].join('\n');

  const user = JSON.stringify({
    scribe: scribeObj||{},
    transcript_summary: String(transcript||'').slice(0,1500),
    override_memo: String(overrideMemo||'')
  });

  const body = {
    model,
    response_format: { type: 'json_object' },
    messages: [
      {role:'system', content: sys},
      {role:'user',   content: user}
    ]
  };

  const resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method:'post', contentType:'application/json', muteHttpExceptions:true,
    headers:{ Authorization:'Bearer '+apiKey },
    payload: JSON.stringify(body)
  });
  if (resp.getResponseCode() !== 200){
    throw new Error('OpenAI Strategist error: ' + resp.getResponseCode() + ' ' + resp.getContentText());
  }
  const json = JSON.parse(resp.getContentText());
  let text = (json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content) || '{}';
  try { return JSON.parse(text); } catch(_) { return { recommended_play: text }; }
}

/* ---------- [STRATEGIST] Re-run pipeline with override ---------- */
function rerunStrategistWithOverride_(rootApptId, reportIdOpt, memoText, actorEmail){
  const lock = LockService.getDocumentLock();
  lock.waitLock(30 * 1000);
  try {
    const ms = MASTER_SS_();
    const apId = getApFolderIdForRoot_(ms, rootApptId);           // ← use apId consistently
    const ap   = DriveApp.getFolderById(apId);

    // Save/append memo
    const saved = saveOverrideMemo_(ap, rootApptId, memoText, actorEmail);

    // Load inputs
    const {scribeObj}  = loadLatestScribeJsonAndVersion_(ap);
    const transcript   = loadTranscriptSnippet_(ap);
    const overrideMemo = loadOverrideMemoText_(ap);

    // Apply structured corrections from THIS override memo → corrected Scribe
    let scribeForRun = scribeObj;
    let patchedPaths = [];
    try {
      const memoPatch = extractScribePatchFromMemo_(memoText);
      if (memoPatch && Object.keys(memoPatch).length){
        patchedPaths = flattenPatchPaths_(memoPatch);
        scribeForRun = mergeDeep_(scribeObj, memoPatch);
        scribeForRun = normalizeScribe_(scribeForRun);
        saveCorrectedScribeJson_(ap, rootApptId, scribeForRun);
      }
    } catch (e) {
      Logger.log('Patch-from-override failed: ' + (e && e.message ? e.message : e));
    }

    // === Step 2A — Generate MEMO (freeform) ===
    let memoPayload = buildStrategistMemoPayload_(scribeForRun, transcript, overrideMemo);
    try { memoPayload.meta = { __root: rootApptId, __apId: apId }; } catch(_){}
    const memoOutText = openAIResponses_TextOnly_(memoPayload);   // ← avoid redeclaring "memoText"
    strat_writeDebug_(ap, rootApptId, 'memo_generated', memoOutText);
    const memoUrl = saveStrategistMemoText_(ap, rootApptId, memoOutText);

    // === Step 2B — Extract JSON (strict schema) ===
    let extractPayload = buildStrategistExtractPayload_(memoOutText, scribeForRun);
    try { extractPayload.meta = { __root: rootApptId, __apId: apId }; } catch(_){}
    const strategistObj = openAIResponses_(extractPayload);

    // Save Strategist JSON
    const strategistUrl =
      (typeof saveStrategistJson_ === 'function')
        ? saveStrategistJson_(ap, rootApptId, strategistObj)
        : saveStrategistJson_Fallback_(ap, rootApptId, strategistObj);

    // Re-render Client Summary / Consult tab
    let transcriptUrl = '';
    const tFolder = ap.getFoldersByName('03_Transcripts');
    if (tFolder.hasNext()){
      const tf = tFolder.next();
      const txt = newestByRegexInFolder_(tf, /\.txt$/i);
      if (txt) transcriptUrl = 'https://drive.google.com/file/d/' + txt.getId() + '/view';
    }

    const apISO = (typeof getApptIsoForRoot_ === 'function')
      ? getApptIsoForRoot_(ms, rootApptId)
      : new Date().toISOString();

    const reportId = reportIdOpt || getReportIdForRoot_(rootApptId);
    if (!reportId) throw new Error('No Client Status Report ID found for ' + rootApptId);

    try {
      const rpt = SpreadsheetApp.openById(reportId);
      if (typeof ensureReportConfig_ === 'function') {
        try { ensureReportConfig_(rpt, { rootApptId, reportId }); } catch (e1) {
          try { ensureReportConfig_(rpt, rootApptId); } catch (e2) {}
        }
      }
    } catch(e) { /* non-fatal */ }

    if (typeof upsertClientSummaryTab_ === 'function') {
      upsertClientSummaryTab_(rootApptId, scribeForRun, apISO, transcriptUrl, strategistObj, { reportId });
    } else if (typeof rerenderClientSummaryTabForRoot_ === 'function') {
      rerenderClientSummaryTabForRoot_(rootApptId);
    } else {
      throw new Error('Renderer not available (missing upsertClientSummaryTab_ and rerenderClientSummaryTabForRoot_).');
    }

    try {
      const log = ensureOverridesLog_();
      log.appendRow([new Date(), rootApptId, actorEmail||'unknown', String(memoOutText||'').length, saved.aggregateFileId, saved.versionFileId, strategistUrl]);
    } catch(_){}

    return {
      ok: true,
      strategistUrl,
      overrideAggFileId: saved.aggregateFileId,
      patched: friendlyNamesForPaths_(patchedPaths),
      analysisOnly: !(patchedPaths && patchedPaths.length)
    };

  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

/** ---------- Last-mile sanitizer for /v1/responses payload ---------- */
function sanitizeForResponsesApi_(body){
  try {
    if (!body || typeof body !== 'object') return;

    // 1) Remove unsupported params (temperature is the culprit here)
    if ('temperature' in body) delete body.temperature;
    if (body.generation_config && 'temperature' in body.generation_config) {
      delete body.generation_config.temperature;
    }

    // 2A) Relax JSON Schema strictness for response_format.json_schema (old shape)
    if (body.response_format && body.response_format.type === 'json_schema') {
      body.response_format.json_schema = body.response_format.json_schema || {};
      body.response_format.json_schema.strict = false;
    }

    // 2B) Relax JSON Schema strictness for text.format (Responses "text" shape — used by Scribe)
    if (body.text && body.text.format &&
        String(body.text.format.type || '').toLowerCase() === 'json_schema') {
      body.text.format.strict = false;
    }

    // 3) Small cleanup that avoids weird API gripe on nulls
    if (body.max_output_tokens === null) delete body.max_output_tokens;

  } catch (_) {
    // Never throw from a sanitizer
  }
}


/** ---------- [UTIL] ensure text output for Responses API (across variants) ---------- */
/** Force text output for /v1/responses and strip unsupported fields. */
function ensureTextOutputFormat_(body){
  if (!body) return;

  // Force plain text aggregation for Responses API
  const hasText = body.text && body.text.format && String(body.text.format.type || '').toLowerCase() === 'text';
  if (!hasText) body.text = { format: { type: 'text' } };

  // ⚠️ Responses API does NOT accept top-level "response_format"
  if (Object.prototype.hasOwnProperty.call(body, 'response_format')) {
    try { delete body.response_format; } catch (_) { body.response_format = undefined; }
  }

  // Optional: if someone accidentally passed chat-style payload
  if (Object.prototype.hasOwnProperty.call(body, 'messages')) {
    try { delete body.messages; } catch (_) { body.messages = undefined; }
  }
}

/** ---------- [UTIL] Deep text pull from Responses body (handles nested "message" wrappers) ---------- */
function __extractTextFromResponsesBody__(body){
  try {
    // 0) Easy path
    if (typeof body.output_text === 'string' && body.output_text.trim()) {
      return body.output_text.trim();
    }

    // 1) Walk any nested content under body.output[*].content[*]
    const texts = [];
    function walk(node){
      if (!node) return;
      if (Array.isArray(node)) { node.forEach(walk); return; }
      if (typeof node !== 'object') return;

      // Common shapes
      if (typeof node.text === 'string' && node.text.trim()) {
        texts.push(node.text.trim());
      } else if (typeof node.content === 'string' && node.content.trim()) {
        texts.push(node.content.trim());
      }

      // If this is a "message" with its own content array, descend again
      if (node.type === 'message' && Array.isArray(node.content)) {
        walk(node.content);
      }

      // Generic descent into children that might hold nested parts
      if (node.content && Array.isArray(node.content)) walk(node.content);
      if (node.output && Array.isArray(node.output)) walk(node.output);
      if (node.parts && Array.isArray(node.parts)) walk(node.parts);
    }

    if (Array.isArray(body.output)) {
      body.output.forEach(o => walk(o && o.content));
    }

    if (texts.length) return texts.join('\n').trim();

    // 2) Chat Completions fallback (if someone swapped endpoints)
    if (body.choices && body.choices[0] && body.choices[0].message
        && typeof body.choices[0].message.content === 'string') {
      return body.choices[0].message.content.trim();
    }
  } catch(_) {}

  return '';
}


/** ---------- [EXTRACTOR] Robust JSON pull from Responses body ---------- */
function __extractJsonFromResponsesBody__(body) {
  // Try the new-style content containers first
  try {
    if (Array.isArray(body.output)) {
      for (var oi = 0; oi < body.output.length; oi++) {
        var out = body.output[oi];
        var content = (out && out.content) || [];
        for (var ci = 0; ci < content.length; ci++) {
          var c = content[ci];
          if (!c) continue;

          // 1) Native JSON blocks (cover several type names used over time)
          var t = String(c.type || '').toLowerCase();
          if ((t === 'json' || t === 'output_json' || t === 'json_schema') && c.json && typeof c.json === 'object') {
            return c.json;
          }

          // 2) Text blocks that *contain* JSON (possibly fenced)
          var txt = (typeof c.text === 'string' && c.text) ||
                    (typeof c.content === 'string' && c.content) || '';
          var parsed = __tryParseJsonFromText__(txt);
          if (parsed) return parsed;
        }
      }
    }
  } catch (_){ /* fall through */ }

  // 3) Convenience aggregator (some models set this)
  if (typeof body.output_text === 'string' && body.output_text.trim()) {
    var parsed2 = __tryParseJsonFromText__(body.output_text);
    if (parsed2) return parsed2;
  }

  // 4) Chat Completions shape (defensive)
  try {
    if (body.choices && body.choices[0] && body.choices[0].message) {
      var mtxt = String(body.choices[0].message.content || '');
      var parsed3 = __tryParseJsonFromText__(mtxt);
      if (parsed3) return parsed3;
    }
  } catch (_){}

  return null; // caller will decide what to do
}

/** Strip ```json fences and attempt increasingly tolerant JSON parses */
function __tryParseJsonFromText__(s) {
  if (!s || typeof s !== 'string') return null;
  var t = s.replace(/^[\s`]*```(?:json)?\s*/i, '')
           .replace(/```[\s`]*$/,'')
           .trim();

  // Straight parse
  try { return JSON.parse(t); } catch(_){}

  // First to last brace
  var first = t.indexOf('{'), last = t.lastIndexOf('}');
  if (first >= 0 && last > first) {
    var frag = t.slice(first, last + 1);
    try { return JSON.parse(frag); } catch(_){}
  }

  // Balanced-brace scan (first complete object)
  var start = -1, depth = 0;
  for (var i = 0; i < t.length; i++) {
    var ch = t[i];
    if (ch === '{') {
      if (depth === 0) start = i;
      depth++;
    } else if (ch === '}') {
      if (depth > 0) {
        depth--;
        if (depth === 0 && start >= 0) {
          var cand = t.slice(start, i + 1);
          try { return JSON.parse(cand); } catch(_){}
        }
      }
    }
  }
  return null;
}

/** ---------- [OPENAI] Responses API → TEXT ONLY (memo step) ---------- */
function openAIResponses_TextOnly_(payload){
  // Deep clone and strip meta
  var bodyForApi;
  try { bodyForApi = JSON.parse(JSON.stringify(payload||{})); } catch(_) { bodyForApi = payload; }
  if (bodyForApi && bodyForApi.meta) delete bodyForApi.meta;

  // Force text output and sanitize unsupported fields
  ensureTextOutputFormat_(bodyForApi);
  sanitizeForResponsesApi_(bodyForApi);

  const res  = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method:'post',
    contentType:'application/json',
    muteHttpExceptions:true,
    headers:{ Authorization:'Bearer ' + OPENAI_API_KEY_SUM },
    payload: JSON.stringify(bodyForApi)
  });

  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code !== 200) throw new Error('OpenAI Responses ' + code + ': ' + text);

  const body = JSON.parse(text);

  // Optional debug
  try {
    if (STRAT_DEBUG_ON_() && payload && payload.meta && payload.meta.__apId && payload.meta.__root){
      const ap = DriveApp.getFolderById(String(payload.meta.__apId));
      strat_writeDebug_(ap, String(payload.meta.__root), 'responses_raw_memo', body);
    }
  } catch(_) {}

  const out = __extractTextFromResponsesBody__(body);
  if (out && out.trim()) return out.trim();

  try {
    if (STRAT_DEBUG_ON_() && payload && payload.meta && payload.meta.__apId && payload.meta.__root){
      const ap = DriveApp.getFolderById(String(payload.meta.__apId));
      strat_writeDebug_(ap, String(payload.meta.__root), 'memo_no_text_envelope', body);
    }
  } catch(_){}
  throw new Error('No text returned from Responses (memo step).');
}


/** ---------- [OPENAI] Responses API → JSON object (strict, robust nested scan) ---------- */
function openAIResponses_(payload){
  // Strip meta before calling API
  var bodyForApi;
  try {
    bodyForApi = JSON.parse(JSON.stringify(payload||{}));
    if (bodyForApi && bodyForApi.meta) delete bodyForApi.meta;
  } catch(_) {
    bodyForApi = payload;
    if (bodyForApi && bodyForApi.meta) delete bodyForApi.meta;
  }

  // 🔧 last-mile sanitize — remove 'temperature', relax schema strictness
  sanitizeForResponsesApi_(bodyForApi);

  const res  = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method:'post', contentType:'application/json', muteHttpExceptions:true,
    headers:{ Authorization:'Bearer ' + OPENAI_API_KEY_SUM },
    payload: JSON.stringify(bodyForApi)
  });
  const code = res.getResponseCode();
  const raw  = res.getContentText();
  if (code !== 200) throw new Error('OpenAI Responses ' + code + ': ' + raw);

  const body = JSON.parse(raw);

  // Optional: debug raw envelope
  try {
    if (STRAT_DEBUG_ON_() && payload && payload.meta && payload.meta.__apId && payload.meta.__root){
      const ap = DriveApp.getFolderById(String(payload.meta.__apId));
      strat_writeDebug_(ap, String(payload.meta.__root), 'responses_raw', body);
    }
  } catch(_) {}

  // 1) Deep-scan for explicit JSON blocks anywhere under output[*].content[*]
  let foundJson = null;
  function walkForJson(node){
    if (!node || foundJson) return;
    if (Array.isArray(node)) { node.forEach(walkForJson); return; }
    if (typeof node !== 'object') return;

    if (node.type === 'json' && node.json && typeof node.json === 'object' && node.json !== null) {
      foundJson = node.json; return;
    }
    if (node.type === 'message' && Array.isArray(node.content)) walkForJson(node.content);

    if (node.content) walkForJson(node.content);
    if (node.output)  walkForJson(node.output);
    if (node.parts)   walkForJson(node.parts);
  }

  if (Array.isArray(body.output)) body.output.forEach(o => walkForJson(o && o.content));
  if (foundJson) return foundJson;

  // 2) JSON string in output_text → parse
  if (typeof body.output_text === 'string' && body.output_text.trim()) {
    try {
      const parsed = JSON.parse(body.output_text);
      if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) return parsed;
    } catch(_) {}
  }

  // 3) JSON string anywhere in any text node (including ```json fenced blocks)
  const allText = (function(){
    const t = __extractTextFromResponsesBody__(body);
    return t && t.trim() ? t : '';
  })();
  if (allText) {
    const cleaned = allText.replace(/^[\s`]*```json\s*/i,'').replace(/```[\s`]*$/,'').trim();
    try {
      const parsed = JSON.parse(cleaned);
      if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) return parsed;
    } catch(_) {
      const m = allText.match(/\{[\s\S]*\}/);
      if (m) {
        try {
          const parsed2 = JSON.parse(m[0]);
          if (parsed2 && typeof parsed2 === 'object' && !Array.isArray(parsed2)) return parsed2;
        } catch(_) {}
      }
    }
  }

  // 4) Fail with preview
  const preview = typeof raw === 'string' ? raw.slice(0, 800) : '';
  throw new Error('Could not extract JSON object from Responses payload. Preview: ' + preview);
}


/** ---------- [SAVE] Strategist JSON with guardrails ---------- */
function saveStrategistJson_(ap, rootApptId, strategistObj){
  // Unwrap if we accidentally received the Responses envelope
  if (strategistObj && strategistObj.object === 'response') {
    let txt = '';
    if (typeof strategistObj.output_text === 'string') {
      txt = strategistObj.output_text;
    } else if (Array.isArray(strategistObj.output)) {
      txt = strategistObj.output
        .map(o => (o && o.content ? o.content : []))
        .flat()
        .map(c => (c && (c.text || c.content) ? (c.text || c.content) : ''))
        .join('');
    }
    if (!txt && Array.isArray(strategistObj.output)) {
      const jsonBlock = strategistObj.output
        .map(o => (o && o.content ? o.content : []))
        .flat()
        .find(c => c && c.type === 'json' && c.json && typeof c.json === 'object');
      if (jsonBlock) strategistObj = jsonBlock.json;
    } else {
      const cleaned = String(txt||'').replace(/^[\s`]*```json\s*/i,'').replace(/```[\s`]*$/,'').trim();
      try { strategistObj = JSON.parse(cleaned || '{}'); } catch (_) { strategistObj = {}; }
    }
  }

  // Guard: don’t save a request payload by mistake
  const looksLikeRequestPayload =
    strategistObj && typeof strategistObj === 'object' &&
    ('instructions' in strategistObj) && ('input' in strategistObj) && ('text' in strategistObj);

  if (looksLikeRequestPayload) {
    try {
      if (STRAT_DEBUG_ON_()) strat_writeDebug_(ap, rootApptId, 'warning_saved_payload_instead_of_response', strategistObj);
    } catch(_){}
    throw new Error('Strategist result looks like a request payload; aborting save to avoid corrupting 04_Summaries.');
  }

  // Warn if shape isn’t what we expect
  if (!strat_hasExpectedKeys_(strategistObj) && STRAT_DEBUG_ON_()) {
    strat_writeDebug_(ap, rootApptId, 'warning_empty_before_save', strategistObj);
  }

  const folder = (function(){
    const it = ap.getFoldersByName('04_Summaries');
    return it.hasNext() ? it.next() : ap.createFolder('04_Summaries');
  })();
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH-mm");
  const file = folder.createFile(
    Utilities.newBlob(JSON.stringify(strategistObj, null, 2), 'application/json',
                      rootApptId + '__analysis_' + ts + '.json')
  );
  return file.getUrl();
}

/* ===========================================================
 * === Strategist memo payload (force plain text output)    ===
 * ===========================================================
 */
function buildStrategistMemoPayload_(scribeObj, transcript, overrideMemo){
  if (!transcript) throw new Error('Transcript is required for Strategist memo.');

  // Keep the same coaching intent as your current prompt
  const instructions = [
    "You are a luxury fine-jewelry strategist and closer with 50 years of experience.",
    "Goal: deliver the smartest, most candid analysis to help close this specific customer.",
    "Write ONE freeform memo to coach the sales rep. Do NOT output JSON. Do NOT use code fences.",
    "Be concrete, evidence-based, and deeply empathetic; map motivations, fears, hidden constraints, power dynamics, and decision levers.",
    "Lay out the exact close path: sequencing (what to do next, then after), emotional hooks, value anchors, and plausible concessions.",
    "Reference the provided transcript and scribe facts; never fabricate specifics you don’t have.",
    "If a REP DEBRIEF section is present, treat it as authoritative clarifications and prefer it over the live transcript when they conflict."
  ].join('\n');

  return {
    // === Memo 1 model selection ===
    model: 'gpt-5',                // force GPT‑5 for freeform analysis
    reasoning: { effort: 'low' },  // rein in reasoning-token spend, still high-quality

    instructions,
    input: [{
      role: "user",
      content: [
        { type: "input_text", text: "TRANSCRIPT (verbatim; required):" },
        { type: "input_text", text: String(transcript || '') },
        { type: "input_text", text: "SCRIBE FACTS (JSON; factual context):" },
        { type: "input_text", text: JSON.stringify(scribeObj || {}) },
        { type: "input_text", text: "REP OVERRIDE MEMO (if any):" },
        { type: "input_text", text: String(overrideMemo || '') }
      ]
    }],

    // Force Responses API to return plaintext (no JSON/schema here)
    text: { format: { type: "text" }, verbosity: "medium" },

    // Headroom for a solid memo without inviting runaway outputs
    max_output_tokens: 5000
  };
}



/* ===========================================================
 * === Summarizer payload (USE SCRIBE SCHEMA, not strategist) ===
 * ===========================================================
 */
function buildSummarizerPayload_(transcript){
  // ==== Scribe JSON Schema (facts only) ====
  const numberOrNull = {"type":["number","null"]};

  const schema = {
    "type": "object",
    "additionalProperties": false,
    "properties": {
      "customer_profile": {
        "type": "object",
        "additionalProperties": false,
        "properties": {
          "customer_name": { "type": "string" },
          "phone":         { "type": "string" },
          "email":         { "type": "string" },
          "partner_name":      { "type": "string" },
          "occasion_intent":   { "type": ["string","null"] },
          "comm_prefs":        { "type": "array", "items": { "type": "string" } },
          "decision_makers":   { "type": "array", "items": { "type": "string" } }
        }
      },
      "budget":   { "type": "string" },
      "timeline": { "type": "string" },

      "diamond_specs": {
        "type": "object",
        "additionalProperties": false,
        "properties": {
          "lab_or_natural": { "type": ["string","null"] },
          "shape":          { "type": ["string","null"] },
          "carat":          numberOrNull,
          "color":          { "type": ["string","null"] },
          "clarity":        { "type": ["string","null"] },
          "ratio":          { "type": ["string","null"] },
          "cut_polish_sym": { "type": ["string","null"] }
        }
      },

      "design_specs": {
        "type": "object",
        "additionalProperties": false,
        "properties": {
          "metal": { "type": ["string","null"] },
          "ring_size":        { "type": ["string","null"] },
          "band_width_mm":    numberOrNull,
          "wedding_band_fit": { "type": ["string","null"] },
          "engraving":        { "type": ["string","null"] },
          "design_notes":     { "type": ["string","null"] }
        }
      },

      "rapport_notes": { "type": "array", "items": { "type": "string" } },

      "next_steps": {
        "type": "array",
        "items": {
          "type": "object",
          "additionalProperties": false,
          "properties": {
            "owner":   { "type": "string" },
            "task":    { "type": "string" },
            "due_iso": { "type": ["string","null"] },
            "notes":   { "type": "string" }
          },
          "required": ["owner","task","due_iso","notes"]
        }
      },

      "design_refs": {
        "type": "array",
        "items": {
          "type": "object",
          "additionalProperties": false,
          "properties": {
            "name": { "type": "string" },
            "file": { "type": "string" },
            "desc": { "type": "string" }
          },
          "required": ["name","file","desc"]
        }
      },

      "conf": {
        "type": "object",
        "additionalProperties": false,
        "properties": {
          "budget":   numberOrNull,
          "timeline": numberOrNull,
          "diamond":  numberOrNull
        }
      }
    },
    "required": []
  };

const instr =
`You are a sales scribe for fine-jewelry consultations.
Return EXACTLY one JSON object that satisfies the provided JSON Schema.

Rules
- The input may contain two labeled sections:
  • "CONSULT TRANSCRIPT (verbatim)"
  • "REP DEBRIEF (authoritative clarifications & context)"
- When information conflicts, prefer the REP DEBRIEF (it’s the rep’s post‑consult clarification).
- Use ONLY facts explicitly stated in these sections.
- If unknown → set value = "" or null.
- Do NOT invent, infer, or guess details.
- Keep strings concise and sales-usable.

ROUTING RULES
- Metal (karat, color, or two-tone) → design_specs.metal as a short phrase. Do NOT put metal in design_specs.design_notes.
- Examples: "14k yellow gold", "18k two-tone (YG/WG)", "platinum", "undecided (YG vs Pt)".
- Diamond type → diamond_specs.lab_or_natural. If transcript says “lab-grown” or “lab diamond” → write "lab". If it says “natural diamond” or "natural" or "mined" → write "natural". If unclear → leave blank/null.`;

  return {
    model: (typeof OPENAI_MODEL_ === 'function') ? OPENAI_MODEL_() : 'gpt-5-mini',
    instructions: instr,
    input: [{
      role: "user",
      content: [
        { type: "input_text", text: "Parse the consultation and output ONLY one JSON object that matches the schema below." },
        { type: "input_text", text: String(transcript || '') }
      ]
    }],
    text: {
      format: {
        type: "json_schema",
        name: "ConsultSummary",
        strict: false,
        schema: schema                    // ← use the SCRIBE schema, not strategist
      }
    }
  };
}

/** Save the Strategist memo (freeform text) into 04_Summaries and return Drive URL. */
function saveStrategistMemoText_(ap, rootApptId, memoText){
  const sf = ap.getFoldersByName('04_Summaries').hasNext()
    ? ap.getFoldersByName('04_Summaries').next()
    : ap.createFolder('04_Summaries');
  const ts = Utilities.formatDate(new Date(), DEFAULT_TZ_(), "yyyy-MM-dd'T'HH-mm");
  const name = rootApptId + '__analysis_memo_' + ts + '.txt';
  const file = sf.createFile(Utilities.newBlob(String(memoText || ''), 'text/plain', name));
  return 'https://drive.google.com/file/d/' + file.getId() + '/view';
}


/** One-shot Strategist test for a root. Does NOT edit sheets. */
function test_debugStrategistOnce(rootApptId){
  if (!rootApptId) throw new Error('rootApptId required');

  // Resolve AP folder + inputs
  const ms = MASTER_SS_();
  const apId = getApFolderIdForRoot_(ms, rootApptId);
  const ap = DriveApp.getFolderById(apId);

  // newest scribe
  const sIt = ap.getFoldersByName('04_Summaries');
  if (!sIt.hasNext()) throw new Error('No 04_Summaries for '+rootApptId);
  const sf = sIt.next();
  const newest = (re) => {
    let best=null, ts=0, it=sf.getFiles();
    while (it.hasNext()){ const f=it.next(); if(!re.test(f.getName())) continue;
      const t=(f.getLastUpdated?f.getLastUpdated():f.getDateCreated()).getTime();
      if (t>ts){ ts=t; best=f; } }
    return best;
  };
  const scribeFile = newest(/__summary_.*\.json$/i) || newest(/__summary_corrected_.*\.json$/i);
  if (!scribeFile) throw new Error('No Scribe JSON found for '+rootApptId);
  const scribeObj = JSON.parse(scribeFile.getBlob().getDataAsString('UTF-8'));

  // newest transcript (optional)
  let transcript=''; const tfIt = ap.getFoldersByName('03_Transcripts');
  if (tfIt.hasNext()){
    const tf = tfIt.next();
    const txt = (function(){ let b=null,t=0,it=tf.getFiles();
      while(it.hasNext()){ const f=it.next(); if(!/\.txt$/i.test(f.getName())) continue;
        const tt=f.getDateCreated().getTime(); if (tt>t){ t=tt; b=f; } } return b; })();
    if (txt) transcript = txt.getBlob().getDataAsString('UTF-8');
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
  const url = saveStrategistJson_(ap, rootApptId, strategistObj);
  Logger.log('Saved Strategist memo: ' + memoUrl);
}

function testtest_debugStrategistOnce(){
  test_debugStrategistOnce('AP-20250907-003');
}

/* ===========================================================
 * === Strategist extract schema (strict + all required)    ===
 * ===========================================================
 */
const STRATEGIST_JSON_SCHEMA_V3 = {
  "type": "object",
  "additionalProperties": false,
  "properties": {
    "recommended_play": { "type": "string" },
    "ask_now":          { "type": "string" },
    "today_action":     { "type": "string" },

    "executive_summary": { "type": "array", "items": { "type": "string" } },

    "viewing_lineup":   { "type": "array", "items": { "type": "string" } },
    "viewing_strategy": { "type": "array", "items": { "type": "string" } },

    "close_sequence":   { "type": "array", "items": { "type": "string" } },
    "close_strategy":   { "type": "array", "items": { "type": "string" } },

    "top_objections": {
      "type": "array",
      "items": {
        "type": "object",
        "additionalProperties": false,
        "properties": {
          "objection": { "type": "string" },
          "reply":     { "type": "string" }
        },
        "required": ["objection","reply"]
      }
    },

    "where_customer_stands_narrative": { "type": "string" },
    "where_customer_stands": {
      "type": "array",
      "items": {
        "type": "object",
        "additionalProperties": false,
        "properties": {
          "point":          { "type": "string" },
          "evidence_level": { "type": "string", "enum": ["direct","strong_inference","weak_inference"] },
          "why":            { "type": "string" }
        },
        "required": ["point","evidence_level","why"]
      }
    },

    "client_priorities": {
      "type": "object",
      "additionalProperties": false,
      "properties": {
        "top_priorities":  { "type": "array", "items": { "type": "string" } },
        "non_negotiables": { "type": "array", "items": { "type": "string" } },
        "nice_to_haves":   { "type": "array", "items": { "type": "string" } }
      },
      "required": ["top_priorities","non_negotiables","nice_to_haves"]
    },

    "evidence_pins": { "type": "array", "items": { "type": "string" } }
  },
  "required": [
    "recommended_play",
    "ask_now",
    "today_action",
    "executive_summary",
    "viewing_lineup",
    "viewing_strategy",
    "close_sequence",
    "close_strategy",
    "top_objections",
    "where_customer_stands_narrative",
    "where_customer_stands",
    "client_priorities",
    "evidence_pins"
  ]
};

/** Step 2B — Build payload to EXTRACT structured JSON from the memo (strict schema). */
function buildStrategistExtractPayload_(memoText, scribeObj){
  // This step MUST return a strict JSON object; the schema lives in STRATEGIST_JSON_SCHEMA_V3.
  const instructions = [
    "You are a meticulous extractor. Return EXACTLY one JSON object matching the given JSON Schema.",
    "Sources: (1) the strategist MEMO (primary) (2) the Scribe facts (secondary).",
    "Rules:",
    "- All top-level keys in the schema are REQUIRED.",
    "- If there is no evidence:",
    "  • string fields → return \"\" (empty string).",
    "  • array fields  → return [].",
    "  • object fields → include the object with its required keys, each empty as above.",
    "- Do not invent facts. Keep bullets concise. Do not add keys."
  ].join('\n');

  return {
    // === Memo 2 model selection ===
    model: 'gpt-5-mini',  // fast & cheap extractor; perfect for schema fill

    instructions,
    input: [{
      role: "user",
      content: [
        { type: "input_text", text: "STRATEGIST MEMO (freeform):" },
        { type: "input_text", text: String(memoText || '') },
        { type: "input_text", text: "SCRIBE FACTS (JSON; for cross-check):" },
        { type: "input_text", text: JSON.stringify(scribeObj || {}) }
      ]
    }],
    text: {
      format: {
        type:  "json_schema",
        name:  "StrategistExtractV3",
        strict:true,
        schema: STRATEGIST_JSON_SCHEMA_V3
      }
    },
    max_output_tokens: 4000
  };
}

/** Locate latest CONSULT and DEBRIEF transcript .txt files and return texts + URLs. */
function __findTranscriptAndDebriefTexts__(ap){
  const out = { consultText:'', debriefText:'', combined:'', consultUrl:'', debriefUrl:'' };

  const tfIt = ap.getFoldersByName('03_Transcripts');
  if (!tfIt.hasNext()) return out;
  const tf = tfIt.next();

  function newestByRegex(re){
    let f=null, t=0, it=tf.getFiles();
    while (it.hasNext()){
      const x=it.next();
      if (!re.test(x.getName())) continue;
      const tt = (x.getLastUpdated?x.getLastUpdated():x.getDateCreated()).getTime();
      if (tt>t){ t=tt; f=x; }
    }
    return f;
  }

  // iOS Shortcut: name your .txt as ...__consult__YYYY...txt and ...__debrief__YYYY...txt
  const fConsult = newestByRegex(/__consult__.*\.txt$/i) || newestByRegex(/\.txt$/i); // fallback: any .txt
  const fDebrief = newestByRegex(/__debrief__.*\.txt$/i);

  if (fConsult){
    out.consultText = fConsult.getBlob().getDataAsString('UTF-8');
    out.consultUrl  = 'https://drive.google.com/file/d/'+fConsult.getId()+'/view';
  }
  if (fDebrief){
    out.debriefText = fDebrief.getBlob().getDataAsString('UTF-8');
    out.debriefUrl  = 'https://drive.google.com/file/d/'+fDebrief.getId()+'/view';
  }

  out.combined =
    (out.consultText ? '=== CONSULT TRANSCRIPT (verbatim) ===\n' + out.consultText.trim() + '\n\n' : '') +
    (out.debriefText ? '=== REP DEBRIEF (authoritative clarifications & context) ===\n' + out.debriefText.trim() : '');

  return out;
}


/***** ==================== DIAG CORE (No behavior change) ==================== *****/

/** Find latest consult + (optional) debrief transcript for a RootApptID. */
function __diag_findLatestTranscriptForRoot_(rootApptId){
  const ss = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  const sh = ss.getSheetByName('00_Master Appointments');
  const HDR = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(s=>String(s||'').trim());
  const iRoot = HDR.indexOf('RootApptID');
  const iFld  = HDR.indexOf('RootAppt Folder ID');
  const iISO  = HDR.indexOf('ApptDateTime (ISO)');
  if (iRoot<0 || iFld<0) throw new Error('Missing RootApptID / RootAppt Folder ID columns');

  const last = sh.getLastRow(); if (last<2) throw new Error('Master has no data rows');

  let rowIdx = 0;
  for (let r=2; r<=last; r++){
    if (String(sh.getRange(r, iRoot+1).getValue()||'').trim() === String(rootApptId).trim()){ rowIdx = r; break; }
  }
  if (!rowIdx) throw new Error('RootApptID not found: '+rootApptId);

  const apId = String(sh.getRange(rowIdx, iFld+1).getValue()||'').trim();
  const apISO= (iISO>=0) ? String(sh.getRange(rowIdx, iISO+1).getValue()||'') : '';
  if (!apId) throw new Error('RootAppt Folder ID missing for '+rootApptId);

  const ap = DriveApp.getFolderById(apId);

  const pair = __findTranscriptAndDebriefTexts__(ap);
  if (!pair.consultText && !pair.debriefText) {
    throw new Error('No transcript .txt found for '+rootApptId+' (looked for __consult__/__debrief__ in 03_Transcripts).');
  }

  // Prefer the combined string (has labeled sections)
  const transcript   = pair.combined || pair.consultText || '';
  const transcriptUrl= pair.consultUrl || '';
  const debriefUrl   = pair.debriefUrl || '';

  return { ss, sh, apId, apISO, ap, transcript, transcriptUrl, debriefUrl };
}


/** Recursively find all occurrences of a key (e.g., "temperature") in an object. */
function __diag_findOccurrencesOfKey_(obj, key){
  const out = [];
  (function walk(node, path){
    if (node && typeof node === 'object'){
      for (const k of Object.keys(node)){
        const p = path ? path + '.' + k : k;
        if (k === key) out.push({ path: p, value: node[k] });
        const v = node[k];
        if (v && typeof v === 'object') walk(v, p);
      }
    }
  })(obj, '');
  return out;
}

/** Extract the JSON Schema object from a Responses payload, supporting both shapes:
 *   - payload.response_format.json_schema.schema
 *   - payload.text.format.schema
 */
function __diag_extractSchemaFromResponsesPayload_(payload){
  let schema=null, path='';
  try {
    if (payload && payload.response_format && payload.response_format.json_schema && payload.response_format.json_schema.schema){
      schema = payload.response_format.json_schema.schema;
      path = 'response_format.json_schema.schema';
    } else if (payload && payload.text && payload.text.format && payload.text.format.schema){
      schema = payload.text.format.schema;
      path = 'text.format.schema';
    }
  } catch(_){}
  return { schema, path };
}

/** Validate OpenAI JSON Schema strictness rule:
 * Every object node with .properties must have .required listing **every** property key.
 * Returns { ok, report[], countObjects }.
 */
function __diag_validateRequiredArrays_(schema){
  const report = [];
  let objects = 0;

  function walk(node, path){
    if (!node || typeof node !== 'object') return;
    const isObj = node.type === 'object' && node.properties && typeof node.properties === 'object';
    if (isObj){
      objects++;
      const keys = Object.keys(node.properties);
      const req  = Array.isArray(node.required) ? node.required : null;
      const missingArr = !req;
      const missingKeys = req ? keys.filter(k => req.indexOf(k) === -1) : keys;
      if (missingArr || missingKeys.length){
        report.push({
          where: path || '(root)',
          note: missingArr ? 'required array is missing' : 'required missing keys',
          missing: missingKeys
        });
      }
      // Walk children
      for (const [k, child] of Object.entries(node.properties)){
        if (child && child.type === 'object') walk(child, (path?path+'.':'')+'properties.'+k);
        else if (child && child.type === 'array' && child.items && child.items.type === 'object'){
          walk(child.items, (path?path+'.':'')+'properties.'+k+'.items');
        }
      }
    } else {
      // Arrays / primitives—walk items if object
      if (node.type === 'array' && node.items && node.items.type === 'object'){
        walk(node.items, (path?path+'.':'')+'items');
      }
    }
  }
  walk(schema, '');

  return { ok: report.length === 0, report, countObjects: objects };
}

/** Pretty truncation */
function __diag_trunc_(s, n){ s = String(s||''); return s.length <= n ? s : (s.slice(0, n) + '…'); }

/** Build Scribe payload for a root (NO network). */
function diag_buildScribePayloadForRoot_(rootApptId){
  const ctx = __diag_findLatestTranscriptForRoot_(rootApptId);
  const payload = buildSummarizerPayload_(ctx.transcript); // your existing builder
  return { ctx, payload };
}

/** Preview + validate Scribe payload (no network). */
function test_scribe_payload_preview(rootApptId){
  const { ctx, payload } = diag_buildScribePayloadForRoot_(rootApptId);

  Logger.log('== Scribe Payload Preview for %s ==', rootApptId);
  Logger.log('- Transcript length: %s chars', (ctx.transcript||'').length);
  Logger.log('- Model: %s', String(payload && payload.model || '(none)'));

  // Temperature check (should not be present for models that don’t support it)
  const temps = __diag_findOccurrencesOfKey_(payload, 'temperature');
  Logger.log('- temperature occurrences: %s', temps.length);
  temps.forEach(t => Logger.log('  • %s = %s', t.path, JSON.stringify(t.value)));

  // Schema extraction + validation
  const { schema, path } = __diag_extractSchemaFromResponsesPayload_(payload);
  if (!schema){
    Logger.log('!! No JSON Schema found in payload (checked response_format.json_schema.schema and text.format.schema).');
  } else {
    Logger.log('- JSON Schema path: %s', path);
    const check = __diag_validateRequiredArrays_(schema);
    Logger.log('- Object nodes scanned: %s', check.countObjects);
    if (check.ok){
      Logger.log('✔ Schema passes strict required-array rule.');
    } else {
      Logger.log('✖ Schema has %s location(s) missing required[] or missing keys:', check.report.length);
      check.report.forEach((r,i) => {
        Logger.log('  [%s] at %s — %s — missing: %s',
          i+1, r.where, r.note, (r.missing||[]).join(', ') || '(none)');
      });
    }
  }

  // Small preview of payload (don’t flood logs)
  try {
    const safePreview = JSON.parse(JSON.stringify(payload));
    // Drop long transcript echo if builder embeds it
    if (safePreview && safePreview.input && typeof safePreview.input === 'string'){
      safePreview.input = __diag_trunc_(safePreview.input, 1000);
    }
    Logger.log('- Payload preview:\n%s', __diag_trunc_(JSON.stringify(safePreview), 4000));
  } catch(_){}

  return payload; // so you can inspect in Apps Script debugger
}

/** Actually call Scribe once (uses your existing openAIResponses_), catching & logging full error. */
function test_scribe_call_once(rootApptId){
  const { ctx, payload } = diag_buildScribePayloadForRoot_(rootApptId);

  // Preflight: show schema issues before hitting network
  const { schema, path } = __diag_extractSchemaFromResponsesPayload_(payload);
  if (schema){
    const check = __diag_validateRequiredArrays_(schema);
    if (!check.ok){
      Logger.log('⚠ Preflight schema issues detected (%s problems). OpenAI will likely 400:', check.report.length);
      check.report.forEach((r,i) => Logger.log('  [%s] %s — %s — missing: %s', i+1, r.where, r.note, (r.missing||[]).join(', ')));
    }
  } else {
    Logger.log('⚠ No JSON Schema found in payload. If the model expects one, this will 400.');
  }

  // Call
  try {
    const resultObj = openAIResponses_(payload); // your existing function
    Logger.log('✅ Scribe returned object keys: %s', Object.keys(resultObj||{}).join(', '));
    return resultObj;
  } catch (e){
    Logger.log('❌ Scribe call failed: %s', e && (e.stack || e.message) || e);
    // Helpful: dump the schema portion that OpenAI referenced
    if (schema){
      try {
        Logger.log('— Schema (first 2k chars) —\n%s', __diag_trunc_(JSON.stringify(schema), 2000));
      } catch(_){}
    }
    // Also show where temperature appears, if any
    const temps = __diag_findOccurrencesOfKey_(payload, 'temperature');
    if (temps.length){
      Logger.log('temperature present at: %s', temps.map(t=>t.path).join(' | '));
    }
    throw e; // rethrow so the execution log captures it clearly
  }
}

/** Locate where the strict flag lives in a Scribe payload and log it. */
function diag_logScribeStrictFlag(rootApptId){
  const { payload } = diag_buildScribePayloadForRoot_(rootApptId); // from earlier helper you ran
  let strictPath = '', strictVal = '(unset)';
  try {
    if (payload && payload.text && payload.text.format && typeof payload.text.format.strict !== 'undefined') {
      strictPath = 'text.format.strict';
      strictVal  = String(payload.text.format.strict);
    } else if (payload && payload.response_format && payload.response_format.json_schema) {
      strictPath = 'response_format.json_schema.strict';
      strictVal  = String(payload.response_format.json_schema.strict);
    }
  } catch(_){}
  Logger.log('Scribe strict flag: path=%s  value=%s', strictPath || '(not found)', strictVal);
  return { strictPath, strictVal, payload };
}

/** One-off DIAGNOSTIC call: force non-strict for this single test and call OpenAI. */
function test_scribe_call_non_strict_once(rootApptId){
  const { ctx, payload } = diag_buildScribePayloadForRoot_(rootApptId);
  // Flip strict=false ONLY in this diagnostic copy
  try {
    if (payload && payload.text && payload.text.format) payload.text.format.strict = false;
    if (payload && payload.response_format && payload.response_format.json_schema)
      payload.response_format.json_schema.strict = false;
  } catch(_){}

  // Call your existing Responses client – we want to see if removing strict alone unblocks it.
  try {
    const resultObj = openAIResponses_(payload);
    Logger.log('✅ Non-strict Scribe OK. Keys: %s', Object.keys(resultObj||{}).join(', '));
    return resultObj;
  } catch (e){
    Logger.log('❌ Non-strict Scribe still failed: %s', e && (e.stack || e.message) || e);
    throw e;
  }
}


function testtest_scribe_payload_preview(){
test_scribe_payload_preview('AP-20250907-003');
}

function test_diag_logScribeStrictFlag(){
diag_logScribeStrictFlag('AP-20250907-003');
}

function testtest_scribe_call_once(){
test_scribe_call_once('AP-20250907-003');
}
function testtest_scribe_call_non_strict_once(){
test_scribe_call_non_strict_once('AP-20250907-003');
}

function testtest_strategist_memo_only(){
test_strategist_memo_only('AP-20250907-003');
}

function testtest_strategist_extract_only(){
test_strategist_extract_only('AP-20250907-003', /*optional*/ memoText);
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



