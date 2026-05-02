/**************************************************************
 * Knowledge Pack Exporter — v5 (Rules + Pragmas + Pairing)
 * - Assigns files to bundles by:
 *    (1) in-file pragma:  // @bundle: <Bundle Title>
 *    (2) filename regex rules (ordered)
 *    (3) catch-all bundle + Unassigned Report
 * - Pairs server/UI that belong together (e.g., start3d_server + dlg_start3d)
 * - REST-only Drive writes, backoff on 429/5xx
 *
 * Keep scopes you already have from v4.
 **************************************************************/

/** ---------- Util & HTTP ---------- **/
function _props_(){ return PropertiesService.getScriptProperties().getProperties(); }
function _tz_(){ return (_props_().DEFAULT_TZ || 'America/Los_Angeles').trim(); }
function _todayStamp_(){ return Utilities.formatDate(new Date(), _tz_(), 'yyyy-MM-dd'); }
function _idFromAnyGoogleUrl_(s){
  s = String(s || '').trim();
  let m = s.match(/\/d\/([a-zA-Z0-9_-]{20,})/); if (m) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]{20,})/); if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return s;
  return '';
}
function _http_(url, opt, label){
  let o = Object.assign({ muteHttpExceptions: true }, opt || {});
  if (!o.headers) o.headers = {};
  o.headers.Authorization = 'Bearer ' + ScriptApp.getOAuthToken();
  let lastText = '', lastCode = 0, delay = 250;
  for (let i = 0; i < 6; i++){
    const res = UrlFetchApp.fetch(url, o);
    const code = res.getResponseCode();
    const ok = code >= 200 && code < 300;
    lastCode = code; lastText = res.getContentText();
    if (ok) return res;
    if (code === 429 || (code >= 500 && code < 600)){ Utilities.sleep(delay); delay = Math.min(5000, delay * 2); continue; }
    break;
  }
  throw new Error((label || 'HTTP') + ' failed ' + lastCode + ': ' + lastText.slice(0, 500));
}

/** ---------- Drive (REST) ---------- **/
function driveGetOrCreateFolder_(parentIdOrRoot, name){
  const parentId = parentIdOrRoot && parentIdOrRoot !== '' ? parentIdOrRoot : 'root';
  const q = [
    "mimeType = 'application/vnd.google-apps.folder'",
    "trashed = false",
    "'" + parentId + "' in parents",
    "name = '" + name.replace(/'/g, "\\'") + "'"
  ].join(' and ');
  const listUrl = 'https://www.googleapis.com/drive/v3/files'
    + '?q=' + encodeURIComponent(q)
    + '&pageSize=10&supportsAllDrives=true&includeItemsFromAllDrives=true&fields=files(id,name)';
  const listRes = _http_(listUrl, { method: 'get' }, 'Drive list');
  const files = JSON.parse(listRes.getContentText()).files || [];
  if (files.length) return files[0];
  const meta = { name, parents: [parentId], mimeType: 'application/vnd.google-apps.folder' };
  const createRes = _http_(
    'https://www.googleapis.com/drive/v3/files?supportsAllDrives=true',
    { method: 'post', contentType: 'application/json; charset=UTF-8', payload: JSON.stringify(meta) },
    'Drive create folder'
  );
  return JSON.parse(createRes.getContentText());
}
function driveUploadText_(parentId, name, content, mime){
  const boundary = 'batch_' + Date.now();
  const meta = { name, parents: [parentId], mimeType: mime || 'text/plain' };
  const body =
    '--' + boundary + '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(meta) + '\r\n' +
    '--' + boundary + '\r\n' +
    'Content-Type: ' + (mime || 'text/plain') + '; charset=UTF-8\r\n\r\n' +
    (content || '') + '\r\n' +
    '--' + boundary + '--';
  const res = _http_(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true',
    { method: 'post', contentType: 'multipart/related; boundary=' + boundary, payload: body },
    'Drive upload'
  );
  return JSON.parse(res.getContentText());
}
function driveFolderUrl_(id){ return 'https://drive.google.com/drive/folders/' + id; }

/** ---------- Apps Script REST ---------- **/
function getProjectContentViaAPI_(scriptId){
  const url = 'https://script.googleapis.com/v1/projects/' + encodeURIComponent(scriptId) + '/content';
  const res = _http_(url, { method: 'get' }, 'Apps Script getContent');
  return JSON.parse(res.getContentText()); // {scriptId, files:[...]}
}

/** ---------- Sheet schema ---------- **/
function buildSheetSchema_(ssId){
  const out = {};
  const ss = SpreadsheetApp.openById(ssId);
  ss.getSheets().forEach(sh => {
    const lastCol = sh.getLastColumn();
    const headers = lastCol
      ? sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(h => String(h || '').trim())
      : [];
    out[sh.getName()] = { headers, approxRows: Math.max(0, sh.getLastRow() - 1) };
  });
  return out;
}

/** ---------- Bundle rules ----------
 * Order matters. First matching rule (or pragma) wins.
 * Title must match your existing bundle file names from v4 so Projects stay stable.
 */
const BUNDLES = [
  { title: 'Core & Config',                                out: '01 - Core & Config.txt',
    tests: [/^appsscript$/, /^00_Lib$/, /^ScriptProperties$/, /^99_KnowledgePack_Exporter$/] },

  { title: 'Calendly Webhook + Queue',                     out: '02 - Calendly Webhook + Queue.txt',
    tests: [/^Calendly/i, /^Webhook/i, /^cal.*queue/i, /^processCalendlyQueue$/] },

  { title: 'Resolver & Artifacts (folders, templates, intake/checklist/quotation)', out: '03 - Resolver & Artifacts.txt',
    tests: [/^Resolver$/] },

  { title: 'Upload → Workers → Summary Renderer → Ask Controller/Core', out: '04 - Conversations & Summaries Suite.txt',
    tests: [/^01_UploadEndpoint$/, /^02_Workers$/, /^03_SummaryRenderer$/, /^04_AskController$/, /^05_AskCore$/] },

  { title: 'Client Status (Server + UI)',                  out: '05 - Client Status (Server + UI).txt',
    tests: [/^ClientStatus_v1$/, /^dlg_client_status_v1$/] },

  { title: 'Client Summary (Server + UI)',                 out: '06 - Client Summary (Server + UI).txt',
    tests: [/^client_summary_server$/, /^dlg_client_summary_v1$/] },

  { title: 'Appointment Summary (Server + UI)',            out: '07 - Appointment Summary (Server + UI).txt',
    tests: [/^appt_summary_server$/, /^dlg_appt_summary_v1$/] },

  { title: 'Start 3D + Assign SO + config check (+ Wax Requests)', out: '08 - Sales & 3D — Start & Assign (Server + UI).txt',
    tests: [/^v1_sales_menu$/, /^start3d_server$/, /^AssignSO_menu$/, /^dlg_start3d_v1$/, /^dlg_assign_so_v1$/, /^cfg_start3d_check$/, /^WaxRequests$/] },

  { title: '3D Revision flow (server + dialog)',           out: '09 - Sales & 3D — Revisions (Server + UI).txt',
    tests: [/^RevisionRequest_menu$/, /^revision3d_server$/, /^dlg_revision3d_v1$/] },

  { title: 'Record Payment flow',                           out: '10 - Payments — Record (Server + UI).txt',
    tests: [/^Payments_v1$/, /^dlg_record_payment_v1$/] },

  { title: 'Payment Summary flow',                          out: '11 - Payments — Summary (Server + UI).txt',
    tests: [/^Payment_Summary_v1$/, /^dlg_payment_summary_v1$/] },

  { title: 'Operational reports + audit + deadlines tool',  out: '12 - Reports (Server + UI + Audit + Deadlines).txt',
    tests: [/^report_server$/, /^dlg_report_status_v1$/, /^dlg_report_reps_v1$/, /^Auditv1$/, /^Deadlines_v1$/, /^dlg_record_deadline_v1$/] },

  { title: 'Propose diamonds + Update quotation (settings + dialogs)', out: '13 - Diamonds — Propose & Update Quotation (Server + UI).txt',
    tests: [/^Diamonds_v1$/, /^dlg_propose_diamonds_v1$/, /^UpdateQuotation_v1$/, /^dlg_update_quote_diamonds_v1$/, /^UpdateQuotation_Settings_v1$/, /^dlg_update_quote_settings_v1$/] },

  { title: 'Order approvals',                               out: '14 - Diamonds — Order Approvals (Server + UI).txt',
    tests: [/^Diamonds_OrderApprove_v1$/, /^dlg_order_approve_diamonds_v1$/] },

  { title: 'Confirm delivery + stone decision dialogs',     out: '15 - Diamonds — Confirm Delivery & Stone Decisions (Server + UI).txt',
    tests: [/^Diamonds_ConfirmDelivery_v1$/, /^dlg_confirm_delivery_v1$/, /^Diamonds_StoneDecision_v1$/, /^dlg_stone_decision_v1$/] },

  { title: 'Ack pipes + dashboard + schedule + snapshot',   out: '16 - Acknowledgements Suite.txt',
    tests: [/^ack_pipes$/, /^ack_phase_d_dashboard$/, /^ack_phase_e_schedule$/, /^ack_phase_f_snapshot$/] },

  { title: 'Reminders + queues + DV hooks + shim + dialog + debug', out: '99 - Reminders & Follow-ups Suite.txt',
    tests: [/^Reminders_v1$/, /^followups_menu$/, /^setup_shims$/, /^dv_reminders_constants$/, /^dv_queue_upserts$/, /^dv_hooks_master$/, /^dlg_reminders_snooze$/, /^99_debug_utils$/] },
];

const CATCH_ALL_TITLE = 'Misc & Utilities (auto-collected)';
const CATCH_ALL_FILE  = '99 - Misc & Utilities.txt';

/** ---------- Pragmas & pairing ---------- **/
function readPragmaBundle_(src){
  // Look only at the first ~2KB for a pragma.
  const head = (src || '').slice(0, 2048);
  const m = head.match(/@bundle:\s*([^\n\r]+)/i);
  return m ? m[1].trim() : '';
}

// Pairs server/UI based on stems (start3d ↔ dlg_start3d_v1, etc.)
function stemsFor_(base){
  const b = String(base || '');
  // Known stems you use often:
  const stems = [];
  const m3d1 = b.match(/^(start3d|AssignSO|revision3d)/i);
  if (m3d1) stems.push(m3d1[1].toLowerCase());
  const mdia = b.match(/^Diamonds_(OrderApprove|ConfirmDelivery|StoneDecision)/i);
  if (mdia) stems.push('diamonds');
  // Generic stem: strip prefixes/suffixes
  stems.push(b.replace(/^dlg_/, '').replace(/_(server|menu|v1)$/i, '').toLowerCase());
  return Array.from(new Set(stems.filter(Boolean)));
}

/** ---------- Export (rules + pragmas) ---------- **/
function exportKnowledgePack(){
  const props = _props_();
  const stamp = _todayStamp_();

  // Destination
  const parentProp = _idFromAnyGoogleUrl_(props.PACK_PARENT_FOLDER_ID || '') || 'root';
  const packFolder = driveGetOrCreateFolder_(parentProp, `[PACK] ${stamp}`);
  const packId = packFolder.id;

  // 1) Fetch project sources
  const proj = getProjectContentViaAPI_(ScriptApp.getScriptId());
  const files = proj.files || [];
  const extOf = (t) => t === 'HTML' ? '.html' : t === 'SERVER_JS' ? '.gs' : t === 'JSON' ? '.json' : '.txt';

  // Index by name and gather metadata
  const byName = {};  // { name : { name, type, ext, source, pragma, stems[] } }
  files.forEach(f => {
    const name = f.name; // Apps Script "name" (no extension)
    const src  = f.source || '';
    byName[name] = {
      name, type: f.type, ext: extOf(f.type), source: src,
      pragma: readPragmaBundle_(src),
      stems: stemsFor_(name)
    };
  });

  // 2) FUNCTION_INDEX.json
  const funIndex = [];
  const fnNames = (src) => {
    const set = new Set();
    for (const m of src.matchAll(/(?:^|\s)function\s+([A-Za-z0-9_]+)\s*\(/g)) set.add(m[1]);
    for (const m of src.matchAll(/(?:const|let|var)\s+([A-Za-z0-9_]+)\s*=\s*function\s*\(/g)) set.add(m[1]);
    for (const m of src.matchAll(/(?:const|let|var)\s+([A-Za-z0-9_]+)\s*=\s*\([^)]*\)\s*=>/g)) set.add(m[1]);
    return Array.from(set).sort();
  };
  files.forEach(f => {
    const ext = extOf(f.type);
    const nameWithExt = `${f.name}${ext}`;
    const fns = fnNames(f.source || '');
    if (fns.length) funIndex.push({ file: nameWithExt, type: f.type, functions: fns });
  });
  driveUploadText_(packId, 'FUNCTION_INDEX.json', JSON.stringify(funIndex, null, 2), 'application/json');

  // 3) Assign to bundles: pragma → regex rules → catch-all
  const used = new Set();
  const bundleBuffers = new Map(); // outFile -> { title, items: [name.ext], sections: [text] }
  function addToBundle_(bundleOut, bundleTitle, nameExt, source){
    if (!bundleBuffers.has(bundleOut)){
      bundleBuffers.set(bundleOut, { title: bundleTitle, items: [], sections: [] });
    }
    const buf = bundleBuffers.get(bundleOut);
    buf.items.push(nameExt);
    buf.sections.push(`===== FILE: ${nameExt} =====\n` + (source || '') + (source && !source.endsWith('\n') ? '\n' : '') + '\n');
  }

  // First pass: pragma wins
  Object.keys(byName).forEach(k => {
    const f = byName[k];
    const nameExt = `${f.name}${f.ext}`;
    if (used.has(nameExt)) return;
    if (!f.pragma) return;
    // Find matching bundle by title
    const spec = BUNDLES.find(b => b.title.toLowerCase() === f.pragma.toLowerCase());
    if (!spec) return; // pragma points to unknown title → will be handled later
    addToBundle_(spec.out, spec.title, nameExt, f.source);
    used.add(nameExt);
  });

  // Second pass: regex rules
  Object.keys(byName).forEach(k => {
    const f = byName[k];
    const nameExt = `${f.name}${f.ext}`;
    if (used.has(nameExt)) return;
    const spec = BUNDLES.find(b => (b.tests || []).some(re => re.test(f.name)));
    if (spec){
      addToBundle_(spec.out, spec.title, nameExt, f.source);
      used.add(nameExt);
    }
  });

  // Third pass: pairing — pull obvious UI/server companions into the same bundle when their stem matches
  bundleBuffers.forEach((buf, outFile) => {
    // Collect stems already in this bundle
    const stemSet = new Set(buf.items.map(n => stemsFor_(n.replace(/\.[^.]+$/, ''))).flat());
    // Find unassigned files with same stem and pull them in
    Object.keys(byName).forEach(k => {
      const f = byName[k];
      const nameExt = `${f.name}${f.ext}`;
      if (used.has(nameExt)) return;
      const hasSharedStem = f.stems.some(s => stemSet.has(s));
      if (hasSharedStem){
        addToBundle_(outFile, buf.title, nameExt, f.source);
        used.add(nameExt);
      }
    });
  });

  // Fourth: write bundles
  bundleBuffers.forEach((buf, outFile) => {
    const header = `# Bundle: ${buf.title}\nWhat's inside: ${buf.items.join(', ')}\n\n`;
    const content = header + buf.sections.join('');
    driveUploadText_(packId, outFile, content, 'text/plain');
    Utilities.sleep(40);
  });

  // Fifth: catch-all + Unassigned Report
  const leftovers = [];
  const leftSections = [];
  Object.keys(byName).sort().forEach(k => {
    const f = byName[k];
    const nm = `${f.name}${f.ext}`;
    if (!used.has(nm)){
      leftovers.push(nm);
      leftSections.push(`===== FILE: ${nm} =====\n` + (f.source || '') + (f.source && !f.source.endsWith('\n') ? '\n' : '') + '\n');
      used.add(nm);
    }
  });
  if (leftSections.length){
    const content = `# Bundle: ${CATCH_ALL_TITLE}\nWhat's inside: ${leftovers.join(', ')}\n\n` + leftSections.join('');
    driveUploadText_(packId, CATCH_ALL_FILE, content, 'text/plain');

    // Human-readable report to make misses obvious
    const report = [
      '# Unassigned Report',
      '',
      'The following files did not match any explicit rule or pragma and were placed in the catch-all bundle:',
      '',
      ...leftovers.map(n => `- ${n}`),
      '',
      'To fix:',
      '1) Add a pragma at the top of the file, e.g.',
      '   `// @bundle: <exact bundle title>`',
      '2) Or extend the regex rules in BUNDLES for that area.',
      ''
    ].join('\n');
    driveUploadText_(packId, '00 - Unassigned Report.md', report, 'text/markdown');
  }

  // 6) Triggers snapshot
  let trigInfo = [];
  try {
    trigInfo = ScriptApp.getProjectTriggers().map(t => ({
      handlerFunction: t.getHandlerFunction(),
      eventType: String(t.getEventType()),
      triggerSource: String(t.getTriggerSource())
    }));
  } catch (e) {
    trigInfo = [{ error: 'No permission to read triggers (missing https://www.googleapis.com/auth/script.scriptapp)' }];
  }
  driveUploadText_(packId, 'TRIGGERS.json', JSON.stringify(trigInfo, null, 2), 'application/json');

  // 7) Script property KEYS
  const p = _props_();
  driveUploadText_(packId, 'SCRIPT_PROPERTY_KEYS.json', JSON.stringify(Object.keys(p).sort(), null, 2), 'application/json');

  // 8) Sheet schemas
  const schemas = {};
  if (p.SPREADSHEET_ID) { try { schemas['SPREADSHEET_ID'] = buildSheetSchema_(p.SPREADSHEET_ID); } catch (_){ } }
  const SHEET_KEY_RE = /(SHEET|SPREADSHEET|WORKBOOK|SS)_?ID/i;
  Object.keys(p).forEach(k => {
    if (k === 'SPREADSHEET_ID') return;
    if (!SHEET_KEY_RE.test(k)) return;
    const id = _idFromAnyGoogleUrl_(p[k]);
    if (!id) return;
    try { schemas[k] = buildSheetSchema_(id); } catch (_){}
  });
  driveUploadText_(packId, 'SHEETS_SCHEMA.json', JSON.stringify(schemas, null, 2), 'application/json');
  
  Logger.log('Pack ready: ' + driveFolderUrl_(packId));
}
