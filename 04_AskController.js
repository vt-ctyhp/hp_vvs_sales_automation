/***** =========================================================
 *  100_ MASTER: Ask Strategist (Phase 1)
 *  File: AskController.gs   (new file)
 * ==========================================================**/
/* ---------- [MASTER/ROW] Fetch row data for a RootApptID ---------- */
function getMasterRowForRoot_(rootApptId){
  const ss = MASTER_SS_();
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet "00_Master Appointments"');
  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
  const iRoot = hdr.indexOf('RootApptID');
  if (iRoot < 0) throw new Error('Missing RootApptID column on Master');

  const last = sh.getLastRow(); if (last < 2) throw new Error('Master is empty');
  const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  const idx  = vals.findIndex(r => String(r[iRoot]||'').trim() === String(rootApptId).trim());
  if (idx < 0) return {row: -1, hdr, data: {}};

  const data = {}; hdr.forEach((h,i)=> data[h] = vals[idx][i]);
  return { row: idx+2, hdr, data };
}

/* ---------- [SCRIBE/LOAD + VERSION] Latest Scribe JSON & version key ---------- */
function loadLatestScribeJsonAndVersion_(apFolder){
  // Prefer pointer at AP root
  let versionKey = '';
  let scribeObj = {};
  let source = '';
  const it = apFolder.getFilesByName('scribe.latest.json');
  if (it.hasNext()){
    try{
      const ptr = JSON.parse(it.next().getBlob().getDataAsString('UTF-8'));
      versionKey = 'v' + (Number(ptr.version||0)||0);
      if (ptr.fileId){
        const f = DriveApp.getFileById(ptr.fileId);
        scribeObj = JSON.parse(f.getBlob().getDataAsString('UTF-8'));
        source = 'drive:' + f.getName();
        return {scribeObj, versionKey, source};
      }
    }catch(_){}
  }
  // Fallback to corrected or base in 04_Summaries
  const sFolder = apFolder.getFoldersByName('04_Summaries');
  if (sFolder.hasNext()){
    const sf = sFolder.next();
    const corr = newestByRegexInFolder_(sf, /__summary_corrected_.*\.json$/i);
    const base = newestByRegexInFolder_(sf, /__summary_.*\.json$/i);
    const f = corr || base;
    if (f){
      scribeObj = JSON.parse(f.getBlob().getDataAsString('UTF-8'));
      source = 'drive:' + f.getName();
            versionKey = (corr ? 'corrected:' : 'base:') + String((f.getLastUpdated ? f.getLastUpdated() : f.getDateCreated()).getTime());
    }
  }
  return {scribeObj, versionKey, source};
}

/* ---------- [TRANSCRIPT/SNIPPET] Short snippet (<= 1500 chars) ---------- */
/** Returns CONSULT transcript + REP DEBRIEF combined (trimmed), for memo/summarizer input. */
function loadTranscriptSnippet_(ap, maxCharsOpt){
  const pair = __findTranscriptAndDebriefTexts__(ap);
  const text = (pair.combined || '').trim();
  const MAX  = Math.max(1000, Number(maxCharsOpt || 10000));
  return text.length > MAX ? text.slice(0, MAX) : text;
}


/* ---------- [ASK/EVIDENCE] Build evidence pack + pins (cached) ---------- */
function buildAskEvidencePack_(rootApptId){
  const cache = CacheService.getScriptCache();
  const cacheKeyBase = 'ASK:EVID:' + rootApptId;

  const msRow = getMasterRowForRoot_(rootApptId);
  if (msRow.row < 0) throw new Error('RootApptID not found on Master: ' + rootApptId);

  const apId = getApFolderIdForRoot_(MASTER_SS_(), rootApptId);
  const apFolder = DriveApp.getFolderById(apId);
  const {scribeObj, versionKey, source} = loadLatestScribeJsonAndVersion_(ap);
  const transcript = loadTranscriptSnippet_(ap);
  const overrideMemo = loadOverrideMemoText_(ap);                   // NEW
  const overrideHash = md5Hex_(overrideMemo || '');                 // NEW

  const cacheKey = cacheKeyBase + ':' + (versionKey || 'nov') + ':' + overrideHash; // NEW

  // Try cache
  const hit = cache.get(cacheKey);
  if (hit){ try { return JSON.parse(hit);} catch(_){ } }

  const m = msRow.data;
  const prefs = (scribeObj && scribeObj.preferences) || {};

  const evidence = {
    master: {
      rootApptId: rootApptId,
      customerName: String(m['Customer Name']||''),
      brand: String(m['Brand']||''),
      visitType: String(m['Visit Type']||''),
      budgetRange: String(m['Budget Range']||''),
      styleNotes: String(m['Style Notes']||''),
      diamondType: String(m['Diamond Type']||'')
    },
    scribe: scribeObj || {},
    transcript_summary: transcript || '',
    override_memo: overrideMemo || '',                                // NEW
    meta: {
      scribeSource: source || '',
      scribeVersionKey: versionKey || '',
      overrideHash: overrideHash                                      // NEW
    }
  };

  // Pins for transparency
  const pins = [];
  const f = evidence.master || {};
  ['customerName','brand','visitType','budgetRange','styleNotes','diamondType'].forEach(k=>{ if (f[k]) pins.push('master:'+k); });
  if (prefs.shape)                         pins.push('scribe:/preferences/shape');
  if (prefs.ratio_range || prefs.ratio)    pins.push('scribe:/preferences/ratio_range');
  if (transcript)                          pins.push('transcript:summary');
  if (overrideMemo)                        pins.push('override:memo');          // NEW

  const out = { evidence, pins };
  cache.put(cacheKey, JSON.stringify(out), 600); // 10 minutes
  return out;
}


/* ---------- [ASK/MODEL] Call LLM or facts-only fallback ---------- */
function callAskStrategist_(question, evidence, pins){
  const apiKey = PROP_('OPENAI_API_KEY','');
  const model  = OPENAI_MODEL_();

  if (!apiKey){
    // Facts‑only fallback
    const f = evidence.master || {};
    const facts = [
      f.customerName && `Customer: ${f.customerName} [master:customerName]`,
      f.visitType && `Visit Type: ${f.visitType} [master:visitType]`,
      f.budgetRange && `Budget Range: ${f.budgetRange} [master:budgetRange]`,
      f.styleNotes && `Style Notes: ${f.styleNotes} [master:styleNotes]`,
      f.diamondType && `Diamond Type: ${f.diamondType} [master:diamondType]`,
    ].filter(Boolean);
    return {
      answer: "No model key configured; showing strict facts from Master only.",
      factsHtml: `<ul>${facts.map(x=>`<li>${x}</li>`).join('')}</ul>`,
      inferenceHtml: "None (read‑only).",
      confidence: "High (facts only)",
      pins
    };
  }

  const sys = [
    "You are a senior luxury fine‑jewelry strategist and closer.",
    "Tone: direct & candid, but kind.",
    "Rules:",
    "- Use ONLY the provided EVIDENCE (Master fields, Scribe JSON, transcript snippet).",
    "- If a fact is unknown, say 'Unknown' and propose the minimum next step to learn it.",
    "- Focus on actionable, tactically useful advice.",
    "- Return JSON ONLY with keys: direct_answer (string), reasoning (string), sources (array of strings), confidence (string).",
  ].join('\n');

  const user = [
    "QUESTION:",
    question,
    "",
    "EVIDENCE (JSON):",
    JSON.stringify(evidence),
  ].join('\n');

  const body = {
    model,
    // temperature not supported for gpt-5 chat; use default (1)
    response_format: { type: 'json_object' },   // <— add this
    messages: [
      {role:'system', content: sys},
      {role:'user', content: user}
    ]
  };

  const resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    muteHttpExceptions: true,
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(body)
  });

  if (resp.getResponseCode() !== 200){
    throw new Error('OpenAI error: ' + resp.getResponseCode() + ' ' + resp.getContentText());
  }

  const data = JSON.parse(resp.getContentText());
  const content = (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || '';

  // Try to parse JSON; if model returned text, attempt salvage
  let obj = null;
  try { obj = JSON.parse(content); } catch(_) {
    // Fallback: wrap as direct answer
    obj = { direct_answer: String(content).trim(), reasoning: "", sources: pins || [], confidence: "Med" };
  }

  return {
    answer: obj.direct_answer || "(no answer)",
    factsHtml: "", // optional in Phase 1
    inferenceHtml: obj.reasoning || "",
    confidence: obj.confidence || "—",
    pins: Array.isArray(obj.sources) && obj.sources.length ? obj.sources : (pins||[]),
    model: OPENAI_MODEL_(),             // ← NEW
    question: String(question||'')      // ← NEW (handy for logs)
  };
}

/* ---------- [WEBAPP/DOGET] Add ASK + ADOPT_OVERRIDE actions ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const op = String(p.op || p.action || p.a || '').toLowerCase();
  const root = String(p.root_appt_id || p.root || p.appt || '').trim();
  const token = String(p.token || '').trim();

  const needsToken = { reanalyze_from_report: true, reanalyze: true, override: true, adopt_override: true };
  const RE_TOKEN = PropertiesService.getScriptProperties().getProperty('REPORT_REANALYZE_TOKEN') || '';

  const json = (o) => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  const bad  = (msg, code) => json({ ok:false, error: code || 'BAD_REQUEST', message: String(msg||'') });

  try {
    // --- health check, used by the sidebar bootstrap ---
    if (op === 'ping') {
      const cfg = (function(){
        try {
          const ss = SpreadsheetApp.getActiveSpreadsheet();
          const sh = ss && ss.getSheetByName('_Config');
          if (!sh) return {};
          const vals = sh.getRange(1,1,sh.getLastRow(),2).getValues();
          const map = {};
          vals.forEach(r => { if (r[0]) map[String(r[0]).trim()] = String(r[1]||'').trim(); });
          return map;
        } catch(_) { return {}; }
      })();
      return json({
        ok: true,
        now: new Date().toISOString(),
        scriptId: (ScriptApp.getScriptId && ScriptApp.getScriptId()) || '',
        user: (Session.getActiveUser && Session.getActiveUser().getEmail()) || '',
        reportConfig: cfg
      });
    }

    if (!op) {
      return ContentService.createTextOutput(
        [
          'OK: AskController.doGet',
          'Try:',
          '  ?op=ping',
          '  ?op=reanalyze_from_report&root_appt_id=AP-YYYYMMDD-###&token=…',
          '  ?op=override&root_appt_id=AP-YYYYMMDD-###&memo=Text…&token=…',
          '  ?op=strategist&root_appt_id=AP-YYYYMMDD-###',
          '  ?op=summarize_latest[&root_appt_id=AP-YYYYMMDD-###]',
          '  ?op=rerender_tab&root_appt_id=AP-YYYYMMDD-###'
        ].join('\n')
      ).setMimeType(ContentService.MimeType.TEXT);
    }

    if (needsToken[op]) {
      if (!RE_TOKEN || !token || token !== RE_TOKEN) {
        return bad('Missing/invalid token for op=' + op, 'UNAUTHORIZED');
      }
    }

    // --- ROUTES ---
    if (op === 'reanalyze_from_report' || op === 'reanalyze') {
      if (!root) return bad('root_appt_id required', 'MISSING_ROOT');
      const info = consult_reanalyzeFromCorrections_(root);
      return json(Object.assign({ ok:true, op, root_appt_id: root }, info));
    }

    if (op === 'override' || op === 'adopt_override') {
      if (!root) return bad('root_appt_id required', 'MISSING_ROOT');
      const memo = String(p.memo || p.override_memo || p.m || '').trim();
      const reportId = String(p.report_id || p.rid || '').trim() || null;
      const actor = (Session.getActiveUser && Session.getActiveUser().getEmail()) || '';
      const out = rerunStrategistWithOverride_(root, reportId, memo, actor);
      return json(Object.assign({ ok:true, op, root_appt_id: root }, out));
    }

    if (op === 'strategist' || op === 'deep_dive') {
      if (!root) return bad('root_appt_id required', 'MISSING_ROOT');
      const r = runStrategistAnalysisForRoot(root);
      return json({ ok:true, op, root_appt_id: root, strategistUrl: r && r.strategistUrl });
    }

    if (op === 'summarize_latest' || op === 'summarize') {
      const r = summarizeLatestTranscript(root || null);
      return json(Object.assign({ ok:true, op, root_appt_id: root || '(auto)' }, r || {}));
    }

    if (op === 'rerender_tab' || op === 'render_tab' || op === 'render') {
      if (!root) return bad('root_appt_id required', 'MISSING_ROOT');
      rerenderClientSummaryTabForRoot_(root);
      return json({ ok:true, op, root_appt_id: root, message: 'Rendered' });
    }

    return bad('Unknown op=' + op, 'UNKNOWN_OP');

  } catch (err) {
    return json({ ok:false, error:'EXCEPTION', message: String(err && (err.stack || err.message) || err), strategistUrl: r && r.strategistUrl });
  }
}

/* ---------- [PATCH] Filter to whitelisted Scribe paths (incl. client_priorities) ---------- */
function filterPatchByAllowed_(patch){
  // Canonical allow-list (exact dotted Scribe paths)
  var allowed = new Set([
    // Profile
    'customer_profile.customer_name',
    'customer_profile.comm_prefs',
    'customer_profile.decision_makers',
    'customer_profile.partner_name',
    'customer_profile.occupation',
    'customer_profile.emotional_motivation',

    // Budget & time
    'budget_low',
    'budget_high',
    'timeline',

    // Diamond specs
    'diamond_specs.lab_or_natural',
    'diamond_specs.carat',
    'diamond_specs.color',
    'diamond_specs.clarity',
    'diamond_specs.cut',          // ← added
    'diamond_specs.ratio',

    // Design specs
    'design_specs.metal',
    'design_specs.design_notes',
    'design_specs.ring_size',     // ← added
    'design_specs.target_ratio',  // ← added
    'design_specs.band_width_mm',
    'design_specs.engraving',
    'design_specs.wedding_band_fit',
    'design_specs.cut_polish_sym',

    // Strategist-detailed style (optional but safe)
    'detailed_style.shape',
    'detailed_style.ratio',
    'detailed_style.band_width_mm',
    'detailed_style.engraving',

    // Client priorities (ALL editable)
    'client_priorities.top_priorities',
    'client_priorities.non_negotiables',
    'client_priorities.nice_to_haves',

    // Legacy alias
    'profile.diamond_type'
  ]);

  var aliasToCanonical = {
    'customer_name': 'customer_profile.customer_name',
    'profile.customer_name': 'customer_profile.customer_name',
    'design_notes': 'design_specs.design_notes'
  };

  var out = {};
  (function pull(obj, prefix){
    Object.keys(obj || {}).forEach(function(k){
      var v = obj[k];
      var path = prefix ? (prefix + '.' + k) : k;
      if (aliasToCanonical[path]) path = aliasToCanonical[path];

      if (v && typeof v === 'object' && !Array.isArray(v)) {
        pull(v, path);
      } else if (allowed.has(path)) {
        setByPath_(out, path, v);
      }
    });
  })(patch, '');
  return out;
}


/* ---------- [PATCH] Flatten leaf paths from a nested patch object ---------- */
function flattenPatchPaths_(patch){
  const out = [];
  (function walk(obj, prefix){
    Object.keys(obj || {}).forEach(k=>{
      const v = obj[k];
      const path = prefix ? (prefix + '.' + k) : k;
      if (v && typeof v === 'object' && !Array.isArray(v)) walk(v, path);
      else out.push(path);
    });
  })(patch, '');
  return out;
}

/* ---------- [PATCH] Map dotted paths to friendly sheet labels ---------- */
function friendlyNamesForPaths_(paths){
  const M = {
    'customer_profile.customer_name': 'Client Name',
    'diamond_specs.lab_or_natural':   'Diamond Type',
    'diamond_specs.carat':            'Target Carat(s)',
    'diamond_specs.color':            'Color Range',
    'diamond_specs.clarity':          'Clarity Range',
    'diamond_specs.ratio':            'Ratio',
    'design_specs.band_width_mm':     'Band Width (mm)',
    'design_specs.engraving':         'Engraving',
    'design_specs.wedding_band_fit':  'Wedding Band Fit',
    'design_specs.cut_polish_sym':    'Cut / Polish / Symmetry',
    'detailed_style.shape':           'Shape',
    'detailed_style.ratio':           'Ratio',
    'design_specs.design_notes':      'Design Notes',   // ← NEW
    'budget_low':                     'Budget Low (numeric)',
    'budget_high':                    'Budget High (numeric)',
    'timeline':                       'Timeline',
    'customer_profile.occupation':    'Occupation',
    'customer_profile.partner_name':  'Partner',
    'customer_profile.emotional_motivation': 'Emotional Motivation',
    'profile.diamond_type':           'Diamond Type (legacy)'
  };
  return (paths || []).map(p => M[p] || p);
}



/* ---------- [PATCH] Heuristic fallback (no API key) ---------- */
function applyMemoHeuristicsToScribe_(memo){
  const text = String(memo||'');
  const t = text.toLowerCase();
  const out = {};

  // diamond type (lab vs natural)
  if (/lab[-\s]?grown|wants\s+a\s+lab|lab\s+diamond/.test(t) || /not\s+natural/.test(t)) {
    setByPath_(out, 'diamond_specs.lab_or_natural', 'lab');
    setByPath_(out, 'profile.diamond_type', 'lab'); // legacy alias
  } else if (/\bnatural\b/.test(t)) {
    setByPath_(out, 'diamond_specs.lab_or_natural', 'natural');
    setByPath_(out, 'profile.diamond_type', 'natural');
  }

  // customer name (e.g., "customer name is Maria Reyes")
  const m = text.match(/customer\s*name\s*(is|:)\s*([^\n,.;]+)/i);
  if (m && m[2]) {
    const name = m[2].trim();
    if (name) {
      setByPath_(out, 'customer_profile.customer_name', name);
    }
  }

  // design notes (e.g., "design notes: bezel + vintage basket")
  const dn = text.match(/design\s*notes?\s*(is|:)\s*([^\n]+)/i);
  if (dn && dn[2]) {
    const notes = dn[2].trim();
    if (notes) setByPath_(out, 'design_specs.design_notes', notes);
  }


  return out;
}


/* ---------- [PATCH] LLM extractor (memo -> scribe patch) ---------- */
function extractScribePatchFromMemo_(memoText){
  memoText = String(memoText||'').trim();
  if (!memoText) return {};
  const apiKey = PROP_('OPENAI_API_KEY','');
  if (!apiKey) return applyMemoHeuristicsToScribe_(memoText);

  const sys = [
    'You convert a free-form override memo into a PATCH object for a Scribe JSON.',
    'Only include fields explicitly asserted; do not infer.',
    'Return a JSON OBJECT with any of these dotted paths only (do not include anything about assigned/assisted reps or other Master-only fields):',
    'customer_profile.customer_name (string)  // preferred',
    'customer_name (string)                    // alias accepted',
    'design_specs.design_notes (string)',      // NEW
    'diamond_specs.lab_or_natural ("lab" | "natural"),',
    'diamond_specs.carat (number), diamond_specs.color (string), diamond_specs.clarity (string), diamond_specs.ratio (string/number),',
    'design_specs.band_width_mm (number), design_specs.engraving (string), design_specs.wedding_band_fit (string), design_specs.cut_polish_sym (string),',
    'detailed_style.shape (string), detailed_style.ratio (string),',
    'budget_low (number), budget_high (number), timeline (string),',
    'customer_profile.occupation (string), customer_profile.partner_name (string), customer_profile.emotional_motivation (string),',
    'profile.diamond_type ("lab" | "natural").',
    'Return JSON ONLY.'
  ].join('\\n');


  const user = 'Memo:\n' + memoText + '\n\nReturn the patch object now.';
  const body = {
    model: OPENAI_MODEL_(),
    response_format: { type: 'json_object' },
    messages: [
      { role: 'system', content: sys },
      { role: 'user',   content: user }
    ]
  };

  const resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method:'post', contentType:'application/json', muteHttpExceptions:true,
    headers:{ Authorization:'Bearer '+apiKey },
    payload: JSON.stringify(body)
  });

  if (resp.getResponseCode() !== 200){
    return applyMemoHeuristicsToScribe_(memoText);
  }

  let text;
  try { text = JSON.parse(resp.getContentText()).choices[0].message.content; }
  catch(_) { text = '{}'; }

  let raw;
  try { raw = JSON.parse(text); } catch(_) { raw = {}; }

  return filterPatchByAllowed_(raw);
}

/* ---------- [OVERRIDE] Load aggregated memo (if any) ---------- */
function loadOverrideMemoText_(apFolder){
  try {
    const sf = apFolder.getFoldersByName('04_Summaries');
    if (!sf.hasNext()) return '';
    const folder = sf.next();
    const it = folder.getFilesByName('override_memos.md');       // single aggregate file
    if (it.hasNext()) return it.next().getBlob().getDataAsString('UTF-8');
  } catch(_){}
  return '';
}

/* ---------- [OVERRIDE] Append + version the memo ---------- */
function saveOverrideMemo_(apFolder, rootApptId, memoText, actorEmail){
  const sf = apFolder.getFoldersByName('04_Summaries').hasNext()
    ? apFolder.getFoldersByName('04_Summaries').next()
    : apFolder.createFolder('04_Summaries');

  const nowIso = Utilities.formatDate(new Date(), DEFAULT_TZ_(), "yyyy-MM-dd'T'HH:mm:ss");
  const header = `\n\n### ${nowIso} — ${actorEmail || 'unknown'}\n`;
  // Rotate aggregate file if it would exceed ~200 KB after append
  const ROTATE_BYTES = 200 * 1024;
  const chunk  = header + (String(memoText||'').trim()) + '\n';

  // 1) Append to aggregate file (with rotation) — ALWAYS keep a File handle
  const aggName = 'override_memos.md';
  let aggFile = null;

  // Read current content if the file exists
  const it = sf.getFilesByName(aggName);
  const current = it.hasNext() ? it.next() : null;
  const curText = current ? current.getBlob().getDataAsString('UTF-8') : '';
  const nextText = (curText || '') + chunk;

  if (nextText.length > ROTATE_BYTES) {
    // rotate archive snapshot containing prior+new
    const tsRot = Utilities.formatDate(new Date(), DEFAULT_TZ_(), "yyyy-MM-dd'T'HH-mm");
    sf.createFile('override_memos_' + tsRot + '.md', nextText, MimeType.PLAIN_TEXT);
    // reset primary aggregate to just the new chunk
    if (current) {
      current.setContent(chunk);
      aggFile = current;
    } else {
      aggFile = sf.createFile(aggName, chunk, MimeType.PLAIN_TEXT);
    }
  } else {
    if (current) {
      current.setContent(nextText);
      aggFile = current;
    } else {
      aggFile = sf.createFile(aggName, nextText, MimeType.PLAIN_TEXT);
    }
  }

  // 2) Versioned copy for audit
  const ts = Utilities.formatDate(new Date(), DEFAULT_TZ_(), "yyyy-MM-dd'T'HH-mm");
  const ver = sf.createFile(
    Utilities.newBlob(chunk, 'text/markdown', `${rootApptId}__override_${ts}.md`)
  );

  return {
    aggregateFileId: aggFile.getId(),
    versionFileId: ver.getId()
  };
}

/* ---------- [MASTER] Ensure override log on Master ---------- */
function ensureOverridesLog_(){
  const ss = MASTER_SS_();
  const name = '05_Strategist_Overrides_Log';
  let sh = ss.getSheetByName(name);
  if (!sh){
    sh = ss.insertSheet(name);
    sh.appendRow(['Timestamp','RootApptID','Actor','MemoChars','AggFileId','VersionFileId','StrategistUrl']);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------- [ASK LOG] Ensure log sheet exists inside the given report ---------- */
function ensureAskLogSheet_(reportId){
  const ss = SpreadsheetApp.openById(reportId);
  const NAME = '04_Ask-Strategist_Log';
  let sh = ss.getSheetByName(NAME);
  if (!sh){
    sh = ss.insertSheet(NAME);
    sh.appendRow([
      'Timestamp','Actor','RootApptID','Question','DirectAnswer (first 300)',
      'Confidence','Sources/Pins','Model','ScribeVersionKey','OverrideHash'
    ]);
    sh.setFrozenRows(1);
  } else {
    // Self-heal header (keeps existing data)
    const hdr = sh.getRange(1,1,1,10).getValues()[0].map(v=>String(v||'').trim());
    const want = [
      'Timestamp','Actor','RootApptID','Question','DirectAnswer (first 300)',
      'Confidence','Sources/Pins','Model','ScribeVersionKey','OverrideHash'
    ];
    if (hdr.some((h,i)=>h!==want[i])){
      sh.getRange(1,1,1,want.length).setValues([want]);
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

/* ---------- [ASK LOG] Append one row ---------- */
function appendAskLogRow_(reportId, payload){
  try{
    const sh = ensureAskLogSheet_(reportId);
    const clip = s => String(s||'').slice(0,300);                 // short preview
    const sources = (payload.pins || []).join(' • ');
    sh.appendRow([
      new Date(),
      String(payload.actor||''),
      String(payload.root||''),
      clip(payload.question||''),
      clip(payload.answer||''),
      String(payload.confidence||''),
      sources,
      String(payload.model||''),
      String(payload.scribeVersionKey||''),
      String(payload.overrideHash||'')
    ]);
  }catch(e){
    // never throw back to the webapp
    Logger.log('appendAskLogRow_ error: ' + (e && (e.stack || e.message) || e));
  }
}

/* ---------- [DEEP DIVE] Model call ---------- */
function callAskStrategistDeepDive_(question, evidence){
  const apiKey = PROP_('OPENAI_API_KEY','');
  const model  = 'gpt-5';
  if (!apiKey) throw new Error('OPENAI_API_KEY missing; cannot Deep Dive');

  const sys = [
    "You are a senior luxury fine‑jewelry strategist and coach for sales reps.",
    "Return JSON ONLY with keys:",
    "  coaching_scripts (array of 3–5 short bullets; each ≤ 22 words),",
    "  one_slide_title (string; ≤ 10 words),",
    "  one_slide_bullets (array of 4–6 bullets; each ≤ 12 words).",
    "Rules:",
    "- Use only the provided EVIDENCE (Master facts, Scribe facts, transcript snippet, override memo).",
    "- No speculation; if unknown, omit or say 'Unknown'.",
  ].join('\n');

  const user = JSON.stringify({
    question: String(question||''),
    evidence: evidence || {}
  });

  const body = {
    "model": "gpt-5",
    response_format: { type: 'json_object' },
    messages: [
      { role:'system', content: sys },
      { role:'user',   content: user }
    ]
  };

  const resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method:'post', contentType:'application/json', muteHttpExceptions:true,
    headers:{ Authorization:'Bearer '+apiKey },
    payload: JSON.stringify(body)
  });
  if (resp.getResponseCode() !== 200){
    throw new Error('OpenAI Deep Dive error: ' + resp.getResponseCode() + ' ' + resp.getContentText());
  }
  const json = JSON.parse(resp.getContentText());
  let text = (json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content) || '{}';
  try { return JSON.parse(text); } catch(_) {
    return { coaching_scripts: [text] };
  }
}

/* ---------- [DEEP DIVE] Build one-slide deck in 04_Summaries ---------- */
function createOneSlideDeck_(apFolder, rootApptId, payload){
  const title  = String(payload.one_slide_title || 'Consult — Key Moves').slice(0, 80);
  const bullets = Array.isArray(payload.one_slide_bullets) ? payload.one_slide_bullets.slice(0,6) : [];
  const scripts = Array.isArray(payload.coaching_scripts) ? payload.coaching_scripts.slice(0,5) : [];

  // Create deck, then move into 04_Summaries
  const deck = SlidesApp.create(rootApptId + ' — Deep Dive');
  const deckId = deck.getId();
  try{
    const summaries = (function(){
      const it = apFolder.getFoldersByName('04_Summaries');
      return it.hasNext() ? it.next() : apFolder.createFolder('04_Summaries');
    })();
    DriveApp.getFileById(deckId).moveTo(summaries);
  }catch(_){}

  // Use the first slide as the summary slide
  const slide = deck.getSlides()[0];
  slide.getPageElements().forEach(pe => { try{ pe.remove(); }catch(_){}}); // clear

  // Title
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 40, 30, 600, 60)
       .getText().setText(title).getTextStyle().setBold(true).setFontSize(28);

  // Bullets
  const body = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 40, 110, 600, 300).getText();
  bullets.forEach((b,i)=>{ if (b) body.appendParagraph('• ' + b); });

  // (Optional) Add scripts as speaker notes for the slide
  const notes = slide.getNotesPage().getSpeakerNotesShape().getText();
  if (scripts.length){
    notes.appendParagraph('Coaching scripts:');
    scripts.forEach(s => notes.appendParagraph('- ' + s));
  }

  return 'https://docs.google.com/presentation/d/' + deckId + '/edit';
}


/***** =========================================================
 *  AskController Web App API (for the Consult AI Sidebar)
 *  Actions:
 *    GET/POST ?action=bootstrap_sidebar&root_appt_id=...&report_id=...
 *    POST     { action:"apply_patch", token, root_appt_id, report_id, patch:{...} }
 *    POST     { action:"chat", token, root_appt_id, report_id, message, thread_id? }
 *  Security: write actions require REPORT_REANALYZE_TOKEN
 * ========================================================== */

// --- small response helpers (namespaced to avoid collisions) ---
function AC_json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function AC_err_(code, msg){ return AC_json_({ ok:false, error:code, message:String(msg||'') }); }
function AC_need_(p){ return AC_err_('MISSING_PARAM','Missing param: '+p); }

// ---- token helper ----
function AC_checkToken_(e){
  var have = String((e.parameter && e.parameter.token) || '').trim();
  if (!have && e.postData && /json/i.test(String(e.postData.type||''))){
    try{
      var j = JSON.parse(e.postData.contents||'{}');
      have = String(j.token||'').trim();
    }catch(_){}
  }
  var want = PROP_('REPORT_REANALYZE_TOKEN','');
  return have && want && have === want;
}

// ---- JSON POST body helper ----
function AC_body_(e){
  if (!e || !e.postData) return {};
  if (!/json/i.test(String(e.postData.type||''))) return {};
  try { return JSON.parse(e.postData.contents||'{}'); } catch(_){ return {}; }
}

// ---- load newest Scribe / Strategist from Drive ----
function AC_loadLatestArtifacts_(rootApptId){
  var ss  = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  var apId = getApFolderIdForRoot_(ss, rootApptId);
  var ap   = DriveApp.getFolderById(apId);
  var sumIt = ap.getFoldersByName('04_Summaries');
  var out = { scribeObj:null, scribeUrl:'', strategistObj:null, strategistUrl:'', apFolderUrl: ap.getUrl() };
  if (!sumIt.hasNext()) return out;

  var sf = sumIt.next();

  function newestByName_(re){
    var it = sf.getFiles(); var newest=null, ts=0;
    while (it.hasNext()){
      var f = it.next();
      if (!re.test(f.getName())) continue;
      var t = (f.getLastUpdated?f.getLastUpdated().getTime():f.getDateCreated().getTime());
      if (t>ts){ ts=t; newest=f; }
    }
    return newest;
  }

  var scribe = newestByName_(/__summary_.*\.json$/i) || newestByName_(/__summary_corrected_.*\.json$/i);
  var strat  = newestByName_(/__analysis_.*\.json$/i);
  if (scribe){
    out.scribeObj = JSON.parse(scribe.getBlob().getDataAsString('UTF-8'));
    out.scribeUrl = 'https://drive.google.com/file/d/' + scribe.getId() + '/view';
  }
  if (strat){
    out.strategistObj = JSON.parse(strat.getBlob().getDataAsString('UTF-8'));
    out.strategistUrl = 'https://drive.google.com/file/d/' + strat.getId() + '/view';
  }
  return out;
}

// ---- chat storage under AP root: 05_ChatLogs/ ----
function AC_chatFolder_(apFolder){
  var it = apFolder.getFoldersByName('05_ChatLogs');
  return it.hasNext()? it.next() : apFolder.createFolder('05_ChatLogs');
}
function AC_newThreadId_(){ return Utilities.getUuid().replace(/-/g,''); }
function AC_loadThread_(apFolder, threadId){
  var cf = AC_chatFolder_(apFolder);
  var it = cf.getFilesByName('chat_'+threadId+'.json');
  if (!it.hasNext()) return { thread_id: threadId, messages: [], created_at: new Date().toISOString() };
  var f = it.next();
  try { return JSON.parse(f.getBlob().getDataAsString('UTF-8')); }
  catch(_){ return { thread_id: threadId, messages: [] }; }
}
function AC_saveThread_(apFolder, thread){
  var cf = AC_chatFolder_(apFolder);
  var name = 'chat_'+thread.thread_id+'.json';
  var it = cf.getFilesByName(name);
  var blob = Utilities.newBlob(JSON.stringify(thread, null, 2), 'application/json', name);
  if (it.hasNext()){ it.next().setContent(blob.getDataAsString('UTF-8')); }
  else { cf.createFile(blob); }
}
function AC_pointerChatLatest_(apFolder, threadId){
  try {
    var name='chat.latest.json';
    var payload = { thread_id: String(threadId||'').trim(), version: Date.now() };
    if (!payload.thread_id) return;
    var it = apFolder.getFilesByName(name);
    if (it.hasNext()) it.next().setContent(JSON.stringify(payload, null, 2));
    else apFolder.createFile(Utilities.newBlob(JSON.stringify(payload, null, 2),'application/json',name));
  } catch(_) {}
}

// ---- OpenAI chat (Strategist assistant) ----
function AC_openAIChat_(messages, scribeObj, strategistObj){
  var key = PROP_('OPENAI_API_KEY','');
  if (!key) throw new Error('Missing OPENAI_API_KEY');

  var instr =
    "You are a senior fine-jewelry Strategist embedded in a client Consult tab.\n"+
    "Use FACTS from Scribe JSON and context from Strategist JSON. Be precise; no hallucinations.\n"+
    "When the user corrects facts, acknowledge, and focus replies accordingly. Keep answers concise.";

  var input = [
    { role:"system", content: instr },
    { role:"user",   content: "Scribe JSON (facts):\n"+JSON.stringify(scribeObj||{}) },
    { role:"user",   content: "Strategist JSON:\n"+JSON.stringify(strategistObj||{}) }
  ];
  // append prior conversation
  (messages||[]).forEach(function(m){
    input.push({ role: m.role, content: m.content });
  });

  var body = {
    model: "gpt-5",
    messages: input
  };

  var res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: { Authorization: 'Bearer '+key },
    payload: JSON.stringify(body)
  });
  if (res.getResponseCode() !== 200){
    throw new Error('OpenAI chat '+res.getResponseCode()+': '+res.getContentText());
  }
  var json = JSON.parse(res.getContentText());
  var txt  = (json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content) || '';
  return String(txt).trim();
}

// ---- apply a JSON patch into Scribe, rerun Strategist, re-render tab ----
function AC_applyPatchPipeline_(rootApptId, reportId, patchObj){
  var ss    = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  var apId  = getApFolderIdForRoot_(ss, rootApptId);
  var apFolder = DriveApp.getFolderById(apId);

  // --- merge patch into latest Scribe ---
  var latest = AC_loadLatestArtifacts_(rootApptId);
  var base   = latest.scribeObj || {};
  var merged = (typeof mergeDeep_ === 'function')
                 ? mergeDeep_(base, patchObj || {})
                 : Object.assign({}, base, patchObj || {});
  if (typeof normalizeScribe_ === 'function') merged = normalizeScribe_(merged);

  // --- 1) SAVE corrected Scribe (always) ---
  var correctedUrl = saveCorrectedScribeJson_(apFolder, rootApptId, merged);

  // --- 2) Strategist: best-effort (never block the pipeline) ---
  var strategistObj = null;
  var strategistUrl = '';
  var transcript    = '';

  try {
    // Try to load newest transcript text (best effort)
    try {
      var tf = apFolder.getFoldersByName('03_Transcripts');
      if (tf.hasNext()){
        var t = tf.next();
        var newestTxt = (function(){
          var it=t.getFiles(), best=null, ts=0;
          while (it.hasNext()){
            var f=it.next();
            if (!/\.txt$/i.test(f.getName())) continue;
            var tt=(f.getLastUpdated?f.getLastUpdated():f.getDateCreated()).getTime();
            if (tt>ts){ ts=tt; best=f; }
          }
          return best;
        })();
        if (newestTxt) transcript = newestTxt.getBlob().getDataAsString('UTF-8');
      }
    } catch (_ignore) {}

    // Memo → Extract (can throw if key/model missing)
    var memoPayload = buildStrategistMemoPayload_(merged, transcript || '', '');
    var memoText    = openAIResponses_TextOnly_(memoPayload);
    strat_writeDebug_(apFolder, rootApptId, 'memo_from_patch', memoText);

    var extractPayload = buildStrategistExtractPayload_(memoText, merged);
    strategistObj = openAIResponses_(extractPayload);
    strategistUrl = saveStrategistJson_(apFolder, rootApptId, strategistObj);

  } catch (e) {
    // Log but keep going; we still re-render the tab
    try { strat_writeDebug_(apFolder, rootApptId, 'strategist_error',
           (e && (e.stack || e.message)) || String(e)); } catch(_){}
  }

  // --- 3) Re-render the Consult tab (always) ---
  var apISO = (typeof getApptIsoForRoot_ === 'function')
                ? getApptIsoForRoot_(ss, rootApptId)
                : new Date().toISOString();

  upsertClientSummaryTab_(rootApptId, merged, apISO, '', strategistObj, { reportId: reportId });

  return {
    ok: true,
    correctedUrl: correctedUrl,
    strategistUrl: strategistUrl || '',
    strategistRan: !!strategistUrl
  };
}

// ---- ROUTER: GET + POST ----
/** WebApp controller for Consult AI sidebar (bootstrap/ping only). */
function AC_ping_(e) {
  try {
    var p = (e && e.parameter) || {};
    var op = String(p.op || 'ping').toLowerCase();

    if (op === 'ping') {
      var cfg = (function () {
        // Best-effort: read report config if present
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sh = ss ? ss.getSheetByName('_Config') : null;
        var map = {};
        if (sh) {
          var vals = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
          for (var i = 0; i < vals.length; i++) {
            var k = String(vals[i][0] || '').trim();
            var v = String(vals[i][1] || '').trim();
            if (k) map[k] = v;
          }
        }
        return map;
      })();

      var out = {
        ok: true,
        now: new Date().toISOString(),
        scriptId: ScriptApp.getScriptId ? ScriptApp.getScriptId() : '',
        user: (Session.getActiveUser && Session.getActiveUser().getEmail()) || '',
        reportConfig: cfg
      };
      return ContentService
        .createTextOutput(JSON.stringify(out))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Default response
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, message: 'Consult AI controller ready' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok:false, error: String(err && (err.message || err)) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function AC_doPost_(e){
  try{
    var ct = (e && e.postData && e.postData.type) || '';
    var body = AC_body_(e);
    var action = String(body.action||'').trim();
    if (!action) return AC_need_('action');
    if (!action) {
      // fallback: allow form posts ?action=chat
      action = String((e && e.parameter && e.parameter.action) || '').trim();
      if (!action) return AC_need_('action');
      // promote query params to body-ish if needed
      body.root_appt_id = body.root_appt_id || (e.parameter.root_appt_id || e.parameter.root || '');
      body.report_id    = body.report_id    || (e.parameter.report_id    || e.parameter.rid  || '');
      body.message      = body.message      || (e.parameter.message      || '');
      body.thread_id    = body.thread_id    || (e.parameter.thread_id    || '');
      body.token        = body.token        || (e.parameter.token        || '');
    }

    if (action === 'apply_patch'){
      if (!AC_checkToken_(e)) return AC_err_('AUTH','Bad token');
      var root = String(body.root_appt_id||'').trim();
      var rpt  = String(body.report_id||'').trim();
      var patch= (body.patch && typeof body.patch==='object') ? body.patch : {};
      if (!root) return AC_need_('root_appt_id');
      if (!rpt)  return AC_need_('report_id');

      var out = AC_applyPatchPipeline_(root, rpt, patch);
      return AC_json_(Object.assign({ ok:true }, out));
    }

    if (action === 'chat'){
      if (!AC_checkToken_(e)) return AC_err_('AUTH','Bad token');
      var root = String(body.root_appt_id||'').trim();
      var rpt  = String(body.report_id||'').trim();
      var msg  = String(body.message||'').trim();
      var threadId = String(body.thread_id||'').trim();
      if (!root) return AC_need_('root_appt_id');
      if (!rpt)  return AC_need_('report_id');
      if (!msg)  return AC_need_('message');

      var ss  = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
      var apId= getApFolderIdForRoot_(ss, root);
      var ap  = DriveApp.getFolderById(apId);

      var latest = AC_loadLatestArtifacts_(root);
      threadId = threadId || AC_newThreadId_();
      var thread = AC_loadThread_(ap, threadId);
      thread.messages.push({ role:'user', content: msg, at:new Date().toISOString() });

      var replyText = AC_openAIChat_(thread.messages, latest.scribeObj||{}, latest.strategistObj||{});
      thread.messages.push({ role:'assistant', content: replyText, at:new Date().toISOString() });
      AC_saveThread_(ap, thread);
      AC_pointerChatLatest_(ap, thread.thread_id);

      return AC_json_({ ok:true, thread_id: thread.thread_id, messages: thread.messages });
    }

    return AC_err_('BAD_ACTION','Unknown POST action: '+action);
  } catch (err){
    return AC_err_('POST_ERR', err && (err.stack||err.message) || err);
  }
}

/** Sidebar bootstrap: return root, newest links, and last chat thread (if any). */
function SB_bootstrap() {
  // Read ROOT_APPT_ID from the active report’s _Config sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = (function () {
    var map = {};
    try {
      var sh = ss.getSheetByName('_Config');
      if (!sh) return map;
      var vals = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
      vals.forEach(function (r) {
        var k = String(r[0] || '').trim();
        if (k) map[k] = String(r[1] || '').trim();
      });
    } catch (_) {}
    return map;
  })();

  var root = String(cfg.ROOT_APPT_ID || '').trim();
  if (!root) return { ok: false, error: 'NO_ROOT', message: 'ROOT_APPT_ID missing in _Config' };

  // Resolve AP folder
  var ms  = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  var apId= getApFolderIdForRoot_(ms, root);
  var ap  = DriveApp.getFolderById(apId);

  // Helpers to fetch newest file URLs
  function newestUrlIn_(folder, re) {
    var it = folder.getFiles(), newest = null, ts = 0;
    while (it.hasNext()) {
      var f = it.next();
      if (!re.test(f.getName())) continue;
      var t = (f.getLastUpdated ? f.getLastUpdated() : f.getDateCreated()).getTime();
      if (t > ts) { ts = t; newest = f; }
    }
    return newest ? ('https://drive.google.com/file/d/' + newest.getId() + '/view') : '';
  }

  // Newest transcript
  var transcript_url = '';
  var tIt = ap.getFoldersByName('03_Transcripts');
  if (tIt.hasNext()) transcript_url = newestUrlIn_(tIt.next(), /\.txt$/i);

  // Newest Scribe (prefer corrected), Strategist
  var scribe_url = '', strategist_url = '';
  var sIt = ap.getFoldersByName('04_Summaries');
  if (sIt.hasNext()) {
    var sf = sIt.next();
    // prefer corrected Scribe
    scribe_url = newestUrlIn_(sf, /__summary_corrected_.*\.json$/i) || newestUrlIn_(sf, /__summary_.*\.json$/i);
    strategist_url = newestUrlIn_(sf, /__analysis_.*\.json$/i);
  }

  // Latest chat thread id (optional continuity)
  var latest_thread_id = '';
  var cIt = ap.getFoldersByName('05_ChatLogs');
  if (cIt.hasNext()) {
    var cf = cIt.next();
    var p  = cf.getFilesByName('chat.latest.json');
    if (p.hasNext()) {
      try {
        var j = JSON.parse(p.next().getBlob().getDataAsString('UTF-8'));
        latest_thread_id = String(j.thread_id || '');
      } catch (_){}
    }
  }

  return {
    ok: true,
    root_appt_id: root,
    ap_folder_url: ap.getUrl(),
    transcript_url: transcript_url,
    scribe_url: scribe_url,
    strategist_url: strategist_url,
    latest_thread_id: latest_thread_id
  };
}

/** Sidebar → Apply: filter + route to the pipeline that saves Corrected Scribe, runs Strategist, renders tab. */
function SB_applyPatch(patch) {
  const startedAt = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ---- tiny logger (sheet-level) ----
  function _log(status, info){
    try{
      let sh = ss.getSheetByName('_Apply_Log');
      if (!sh) { sh = ss.insertSheet('_Apply_Log'); sh.appendRow(['Timestamp','Status','Info']); sh.setFrozenRows(1); }
      sh.appendRow([new Date(), status, String(info || '')]);
    }catch(_){}
  }

  try {
    _log('CLICK', 'sidebar SB_applyPatch invoked');

    // 1) Resolve root from _Config
    const cfg = (() => {
      const map = {};
      try {
        const sh = ss.getSheetByName('_Config');
        if (!sh) return map;
        const vals = sh.getRange(1,1,sh.getLastRow(),2).getValues();
        vals.forEach(r => { if (r[0]) map[String(r[0]).trim()] = String(r[1] || '').trim(); });
      } catch (_){}
      return map;
    })();

    const root = String(cfg.ROOT_APPT_ID || '').trim();
    const reportId = ss.getId();
    if (!root){
      _log('ERROR','NO_ROOT in _Config');
      return { ok:false, error:'NO_ROOT', message:'ROOT_APPT_ID missing in _Config' };
    }

    // 2) Filter to allowed paths
    const filtered = filterPatchByAllowed_(patch || {});
    const changed  = flattenPatchPaths_(filtered || {});
    if (!changed.length){
      _log('EMPTY_PATCH','No allowed fields changed');
      return { ok:false, error:'EMPTY_PATCH', message:'No allowed fields changed. Check Column C paths.' };
    }

    // 3) Apply
    const out = AC_applyPatchPipeline_(root, reportId, filtered);
    _log('APPLIED', JSON.stringify({ changed, outStatus: !!(out && out.ok) }));
    return Object.assign({ ok:true, changed_paths: changed, ms: (new Date()-startedAt) }, out);

  } catch (e) {
    _log('EXCEPTION', (e && (e.stack || e.message)) || e);
    return { ok:false, error:'EXCEPTION', message: String(e && (e.message || e)) };
  }
}


/** Sidebar → Live Chat send (continues the latest thread if provided). */
function SB_chatSend(threadId, message) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cfg = (function () {
      var map = {};
      try {
        var sh = ss.getSheetByName('_Config');
        if (sh) {
          var vals = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
          vals.forEach(function (r) { if (r[0]) map[String(r[0]).trim()] = String(r[1] || '').trim(); });
        }
      } catch (_){}
      return map;
    })();

    var root = String(cfg.ROOT_APPT_ID || '').trim();
    if (!root) return { ok: false, error: 'NO_ROOT', message: 'ROOT_APPT_ID missing in _Config' };

    var res = AC_chatCore_(root, ss.getId(), String(message || ''), String(threadId || ''));
    return Object.assign({ ok: true }, res);
  } catch (e) {
    return { ok: false, error: 'EXCEPTION', message: String(e && (e.message || e)) };
  }
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



