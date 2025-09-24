/***** 05_AskCore.gs — library-callable core (100_ project) *****/

/** Public: library function for chat.
 * @param {string} rootApptId
 * @param {string} reportId
 * @param {string} message
 * @param {string=} threadIdOpt
 * @return {{ok:boolean, thread_id:string, messages:Array<{role:string,content:string,at:string}>}}
 */
function AC_chatCore_(rootApptId, reportId, message, threadIdOpt) {
  if (!rootApptId) throw new Error('AC_chatCore_: missing rootApptId');
  if (!message)    throw new Error('AC_chatCore_: missing message');

  const ss  = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  const apId= getApFolderIdForRoot_(ss, rootApptId);
  const ap  = DriveApp.getFolderById(apId);

  // Context: newest Scribe + Strategist (best effort)
  const latest = AC_loadLatestArtifactsFromFolder_(ap);

  // Chat thread store under 05_ChatLogs/
  const threadId = threadIdOpt || AC_newThreadId_();
  const thread   = AC_loadThread_(ap, threadId);
  thread.messages.push({ role: 'user', content: String(message), at: new Date().toISOString() });

  const replyText = AC_openAIChat_(thread.messages, latest.scribeObj || {}, latest.strategistObj || {});
  thread.messages.push({ role: 'assistant', content: replyText, at: new Date().toISOString() });

  AC_saveThread_(ap, thread);
  AC_pointerChatLatest_(ap, thread.thread_id);

  return { ok:true, thread_id: thread.thread_id, messages: thread.messages };
}

/** Public: library function to apply a JSON patch to Scribe, then re-analyze & re-render tab. */
function AC_applyPatchCore_(rootApptId, reportId, patch) {
  if (!rootApptId) throw new Error('AC_applyPatchCore_: missing rootApptId');
  if (!patch || typeof patch !== 'object') throw new Error('AC_applyPatchCore_: patch must be object');

  const lock = LockService.getScriptLock();
  lock.waitLock(30 * 1000);
  try {
    const ms = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
    const apId = getApFolderIdForRoot_(ms, rootApptId);
    const ap   = DriveApp.getFolderById(apId);

    // 1) Load newest Scribe (prefer corrected)
    const { scribeObj, transcript, transcriptUrl } = AC_loadLatestArtifactsFromFolder_(ap);
    if (!scribeObj) throw new Error('No Scribe JSON found for ' + rootApptId);

    // 2) Merge patch → corrected Scribe
    const merged = normalizeScribe_(mergeDeep_(scribeObj, patch));
    const summaryUrl = saveCorrectedScribeJson_(ap, rootApptId, merged);

    // 3) Strategist memo → extract on corrected Scribe
    let memoPayload = buildStrategistMemoPayload_(merged, transcript || '', '');
    const memoText  = openAIResponses_TextOnly_(memoPayload);
    strat_writeDebug_(ap, rootApptId, 'memo_from_patch', memoText);

    let extractPayload = buildStrategistExtractPayload_(memoText, merged);
    const strategistObj = openAIResponses_(extractPayload);
    const strategistUrl = saveStrategistJson_(ap, rootApptId, strategistObj);

    // 4) Re-render the consult tab in the Client Status Report
    const apISO = getApptIsoForRoot_(ms, rootApptId) || new Date().toISOString();
    upsertClientSummaryTab_(rootApptId, merged, apISO, transcriptUrl, strategistObj, { reportId });

    return { ok:true, summaryUrl, strategistUrl };
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

/* ========= Helpers (self-contained) ========= */

function AC_loadLatestArtifactsFromFolder_(apFolder){
  // Newest Scribe JSON (corrected preferred)
  const sFolderIt = apFolder.getFoldersByName('04_Summaries');
  let scribeObj=null, strategistObj=null, transcript='', transcriptUrl='';
  if (sFolderIt.hasNext()){
    const sf = sFolderIt.next();
    const newest = function(re){
      let newest=null, t=0, it=sf.getFiles();
      while (it.hasNext()){
        const f = it.next();
        if (!re.test(f.getName())) continue;
        const ts = (f.getLastUpdated?f.getLastUpdated():f.getDateCreated()).getTime();
        if (ts>t){ t=ts; newest=f; }
      }
      return newest;
    };
    const corrected = newest(/__summary_corrected_.*\.json$/i);
    const base      = newest(/__summary_.*\.json$/i);
    const scribeFile= corrected || base;
    if (scribeFile) scribeObj = JSON.parse(scribeFile.getBlob().getDataAsString('UTF-8'));

    const strat = newest(/__analysis_.*\.json$/i);
    if (strat) strategistObj = JSON.parse(strat.getBlob().getDataAsString('UTF-8'));
  }

  // Newest transcript
  const tFolderIt = apFolder.getFoldersByName('03_Transcripts');
  if (tFolderIt.hasNext()){
    const tf = tFolderIt.next();
    let newest=null, t=0, it=tf.getFiles();
    while (it.hasNext()){
      const f = it.next();
      if (!/\.txt$/i.test(f.getName())) continue;
      const ts = f.getDateCreated().getTime();
      if (ts>t){ t=ts; newest=f; }
    }
    if (newest){
      transcript   = newest.getBlob().getDataAsString('UTF-8');
      transcriptUrl= 'https://drive.google.com/file/d/'+newest.getId()+'/view';
    }
  }
  return { scribeObj, strategistObj, transcript, transcriptUrl };
}

function AC_threadsFolder_(apFolder){
  const it = apFolder.getFoldersByName('05_ChatLogs');
  return it.hasNext() ? it.next() : apFolder.createFolder('05_ChatLogs');
}

function AC_loadThread_(apFolder, threadId){
  const cf = AC_threadsFolder_(apFolder);
  const it = cf.getFilesByName(threadId + '.json');
  if (it.hasNext()){
    const obj = JSON.parse(it.next().getBlob().getDataAsString('UTF-8'));
    obj.thread_id = obj.thread_id || threadId;
    obj.messages = Array.isArray(obj.messages) ? obj.messages : [];
    return obj;
  }
  return { thread_id: threadId, messages: [] };
}

function AC_saveThread_(apFolder, thread){
  const cf = AC_threadsFolder_(apFolder);
  const name = thread.thread_id + '.json';
  const blob = Utilities.newBlob(JSON.stringify(thread, null, 2), 'application/json', name);
  const it = cf.getFilesByName(name);
  if (it.hasNext()){ it.next().setContent(blob.getDataAsString('UTF-8')); }
  else { cf.createFile(blob); }
}

function AC_pointerChatLatest_(apFolder, threadId){
  try{
    const name = 'chat.latest.json';
    const payload = { thread_id: String(threadId||'') };
    if (!payload.thread_id) return;
    const cf = AC_threadsFolder_(apFolder);
    const it = cf.getFilesByName(name);
    if (it.hasNext()) it.next().setContent(JSON.stringify(payload, null, 2));
    else cf.createFile(Utilities.newBlob(JSON.stringify(payload,null,2),'application/json',name));
  }catch(_){}
}

function AC_openAIChat_(messages, scribeObj, strategistObj){
  const apiKey = OPENAI_PROP_('OPENAI_API_KEY') || PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY') || '';
  if (!apiKey) throw new Error('OPENAI_API_KEY missing');

  // Keep the last ~12 turns to control prompt size.
  const last = messages.slice(-24).map(m => ({role: m.role, content: m.content}));

  const sys = [
    'You are Consult AI, a sales assistant for fine-jewelry reps.',
    'Use ONLY facts in the provided Scribe (facts) and Strategist (analysis) JSON when answering.',
    'If the rep asks to change facts, produce clear instructions for what to patch OR confirm the applied patch result when we run it.',
    'Keep answers concise and actionable; include short checklists or bullets when useful.'
  ].join('\n');

  const context = [
    { role: 'system', content: sys },
    { role: 'system', content: 'SCRIBE_JSON:\n' + JSON.stringify(scribeObj || {}) },
    { role: 'system', content: 'STRATEGIST_JSON:\n' + JSON.stringify(strategistObj || {}) }
  ];

  const body = {
    model: 'gpt-5',
    // temperature not supported for this model; use default (1)
    messages: context.concat(last)
  };

  const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method:'post', contentType:'application/json', muteHttpExceptions:true,
    headers:{ Authorization:'Bearer '+apiKey },
    payload: JSON.stringify(body)
  });
  if (res.getResponseCode() !== 200){
    throw new Error('OpenAI Chat error: ' + res.getResponseCode() + ' ' + res.getContentText());
  }
  const json = JSON.parse(res.getContentText());
  return (json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content) || '';
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



