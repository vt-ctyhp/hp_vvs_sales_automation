/** File: start3d_server.gs — FINAL FULL REPLACEMENT (with safe sibling propagation)
 *  What this file provides:
 *   - start3d_init()                 : Step‑1 banner logic (SO / Design Request presence)
 *   - previewOdooPaste(p)            : Pretty “Copy into Odoo” text (fixed Mode)
 *   - getActiveMasterPreview()       : Prefill + context for Step‑3
 *   - checkSOConflicts(payload)      : SO uniqueness across 100_
 *   - saveAssignedSO(payload)        : Writes 100_ only; scaffolds SO folders; propagates SO to same RootApptID rows
 *   - append3DTrackerLog_(...)       : Header‑based logger with “Revision #”
 *
 *  Relies on your existing helpers in this project:
 *   - scaffoldOrderFolders(), moveApFolderToIntake_()
 *   - createClientShortcutForSO_()  (included below)
 */

// ---------- STEP 1: Init banner (SO / Design Request presence) ----------
function start3d_init(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet "00_Master Appointments".');

  const r = ss.getActiveRange();
  if (!r || r.getSheet().getName() !== sh.getName() || r.getRow() < 2){
    throw new Error('Select a data row on "00_Master Appointments" and try again.');
  }
  const row = r.getRow();

  // Batch-read header + the active row in two calls for speed.
  const lastCol = sh.getLastColumn();
  const header  = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h || '').trim());
  const rowVals = sh.getRange(row, 1, 1, lastCol).getDisplayValues()[0];

  // Header map (1-based)
  const H = {};
  for (let i = 0; i < header.length; i++) {
    const k = header[i];
    if (k) H[k] = i + 1;
  }
  const get = n => {
    const c = H[n];
    return c ? rowVals[c - 1] : '';
  };

  const brand = String(get('Brand') || '').toUpperCase().trim();
  const so    = String(get('SO#') || '').replace(/^'/, '').trim();
  const desc  = String(get('Design Request') || '').trim();

  return { ok: true, brand, so, hasSO: !!so, hasDesignRequest: !!desc };
}


// ---------- STEP 2: Pretty Odoo paste ----------
function previewOdooPaste(p){
  const lines = [];
  const add = (s)=>lines.push(s);
  add('—— 3D DESIGN REQUEST — START 3D / CREATE NEW SO ——');
  add('');
  add('SETTING');
  add('• Accent Diamond: ' + (p.AccentDiamondType || ''));
  add('• Ring Style    : ' + (p.RingStyle || ''));
  add('• Metal         : ' + (p.Metal || ''));
  add('• US Size       : ' + (p.USSize || ''));
  add('• Band Width    : ' + (p.BandWidthMM || '') + ' mm');
  add('');
  add('DESIGN NOTES');
  const notes = String(p.DesignNotes||'').trim();
  if (notes){ notes.split(/\r?\n/).forEach(s=> add('• ' + s.replace(/^\s*[-•]\s*/,'').trim())); }
  add('');
  add('CENTER STONE');
  add('• Type       : ' + (p.CenterDiamondType || ''));
  add('• Shape      : ' + (p.Shape || ''));
  add('• Dimension  : ' + (p.DiamondDimension || ''));
  add('');
  add('(Mode: Start 3D Design / Create New SO)');
  return lines.join('\n');
}

// ---------- Combined payload for Step 2 + Step 3 (one round-trip) ----------
function start3d_step2Payload(form){
  try {
    const paste = previewOdooPaste(form);
    const preview = getActiveMasterPreview();
    return { ok: true, paste, preview };
  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}


// ---------- Shared preview for Step 3 / success ----------
function getActiveMasterPreview(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet "00_Master Appointments".');

  const r = ss.getActiveRange();
  if (!r || r.getSheet().getName() !== sh.getName() || r.getRow() < 2){
    throw new Error('Select a data row on "00_Master Appointments" and try again.');
  }
  const row = r.getRow();

  const hdrs = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0] || [];
  const H = {}; hdrs.forEach((h,i)=>{ const k=String(h).trim(); if(k) H[k]=i+1; });

  const getVal  = n => H[n] ? sh.getRange(row, H[n]).getValue() : '';
  const getRich = n => {
    if (!H[n]) return '';
    const rng = sh.getRange(row, H[n]);
    const rtv = rng.getRichTextValue?.();
    return rtv?.getLinkUrl?.() || String(rng.getValue() || '');
  };

  const brand    = String(getVal('Brand') || '').toUpperCase();
  const soRaw    = String(getVal('SO#') || '').trim();
  const so       = soRaw.replace(/^'/,'');
  const odooUrl  = getRich('Odoo SO URL') || getRich('SO URL') || getRich('Odoo Link');

  const customer = getVal('Customer Name') || getVal('Customer') || getVal('Name');
  const email    = getVal('EmailLower') || getVal('Email');
  const phone    = getVal('PhoneNorm')  || getVal('Phone');
  const linkedAt = H['SO Linked At'] ? getVal('SO Linked At') : '';

  const hasLinkedSO = !!(so || linkedAt);
  const tz = Session.getScriptTimeZone?.() || 'America/Los_Angeles';
  const linkedAtDisplay = linkedAt ? Utilities.formatDate(new Date(linkedAt), tz, 'MMM d, yyyy h:mm a z') : '';

  const masterLink = ss.getUrl() + '#gid=' + sh.getSheetId() + '&range=A' + row;

  return {
    row, sheetId: sh.getSheetId(), headers: H,
    brand, so, odooUrl, customer, email, phone, masterLink,
    hasLinkedSO, existingBrand: brand, existingSo: so, existingLinkedAtDisplay: linkedAtDisplay
  };
}

function checkSOConflicts(payload){
  const brand = String(payload?.brand||'').toUpperCase();
  const so    = String(payload?.so||'').trim();
  if (!/^(HPUSA|VVS)$/.test(brand)) return { existsInMaster:false, existsInOrders:false };
  if (!/^\d{2}\.\d{4}$/.test(so))   return { existsInMaster:false, existsInOrders:false };

  // 100_ Master only
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) return { existsInMaster:false, existsInOrders:false };

  const r  = ss.getActiveRange();
  const activeRow = (r && r.getSheet().getName()===sh.getName()) ? r.getRow() : 0;

  // Build header map once, then read only the two needed columns
  const hdrs = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0] || [];
  const H = {}; hdrs.forEach((h,i)=>{ const k=String(h||'').trim(); if(k) H[k]=i+1; });

  let existsInMaster = false, masterHits = [];
  if (H['SO#'] && H['Brand'] && sh.getLastRow() >= 2){
    const rows = sh.getLastRow() - 1;
    const soVals    = sh.getRange(2, H['SO#'],   rows, 1).getDisplayValues();
    const brandVals = sh.getRange(2, H['Brand'], rows, 1).getValues();
    for (let i=0;i<rows;i++){
      const rr = i+2; if (rr === activeRow) continue;
      const b = String(brandVals[i][0]||'').toUpperCase().trim();
      const s = String(soVals[i][0]  ||'').replace(/^'/,'').trim();
      if (b===brand && s===so){
        existsInMaster = true;
        masterHits.push({row:rr, link:ss.getUrl()+'#gid='+sh.getSheetId()+'&range=A'+rr});
        break;
      }
    }
  }

  // Orders (301/302) removed — always false to preserve response shape
  return { existsInMaster, masterHits, existsInOrders:false, ordersLink:'', ordersLabel:'' };
}


// ---------- Save Assign SO (100_ + SO folders + client shortcut + 00‑Intake move + Tracker log + SAFE PROPAGATION) ----------
function saveAssignedSO(payload){
  const brand = String(payload?.brand || '').toUpperCase();
  const so    = String(payload?.so || '').trim();
  const url   = String(payload?.odooUrl || '').trim();
  const designRequest = String(payload?.designRequest || '').trim();
  const designForm    = payload?.designForm || {};
  let   shortTag      = String(payload?.shortTag || '').trim();
  const forceOverwrite= !!payload?.forceOverwrite; // also bypasses propagation safety
  const calledFromAssignSO = !(payload && payload.designForm); // Assign SO dialog sends no designForm

  // Strictness of sibling propagation:
  // - true  = only update siblings if (email OR phone OR name) matches
  // - false = update all siblings with the same RootApptID
  const REQUIRE_SECONDARY_MATCH = true;

  if (!/^(HPUSA|VVS)$/.test(brand)) throw new Error('Brand must be HPUSA or VVS');
  if (!/^\d{2}\.\d{4}$/.test(so))   throw new Error('SO# must be like 12.3456');
  if (!/^(https?:\/\/)?[^\s]+\.com(\/|\?|#|$)/i.test(url)) throw new Error('URL must be a .com address (protocol optional)');

  // Resolve active 100_ row early (needed by duplicate checks)
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet "00_Master Appointments".');

  const r = ss.getActiveRange();
  if (!r || r.getSheet().getName() !== sh.getName() || r.getRow() < 2) {
    throw new Error('Select the target data row in "00_Master Appointments".');
  }
  const row = r.getRow();

  // Block duplicates across DIFFERENT RootApptIDs (OK to repeat within same root)
  (function guardDuplicatesAcrossDifferentRoots(){
    if (forceOverwrite) return;
    const hdr = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getDisplayValues()[0].map(h=>String(h||'').trim());
    const Hmap = {}; hdr.forEach((h,i)=>{ if(h) Hmap[h]=i+1; });

    const cSO    = Hmap['SO#'];
    const cBrand = Hmap['Brand'];
    const cRoot  = Hmap['RootApptID'] || Hmap['APPT_ID']; // fallback if using APPT_ID
    if (!cSO || !cBrand || !cRoot) return;

    const last = sh.getLastRow();
    if (last < 2) return;

    const rows = last - 1;
    const soVals    = sh.getRange(2, cSO,    rows, 1).getDisplayValues();
    const brandVals = sh.getRange(2, cBrand, rows, 1).getValues();
    const rootVals  = sh.getRange(2, cRoot,  rows, 1).getDisplayValues();

    const targetSO    = String(so).trim();
    const targetBrand = String(brand).toUpperCase().trim();
    const currentRoot = String(sh.getRange(row, cRoot).getDisplayValue() || '').trim();

    for (let i=0; i<rows; i++){
      const rr = i + 2;
      if (rr === row) continue;
      const b = String(brandVals[i][0]||'').toUpperCase().trim();
      const s = String(soVals[i][0]  ||'').replace(/^'/,'').trim();
      if (b === targetBrand && s === targetSO){
        const rroot = String(rootVals[i][0]||'').trim();
        if (rroot !== currentRoot) {
          throw new Error('DUPLICATE_SO|' + JSON.stringify({ existsInMaster: true, reason: 'different_root' }));
        }
      }
    }
  })();

  // === Ensure required headers (label-based; safe if they already exist) ===
  const H = (function ensureHeaders(){
    const hdrs = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0]||[];
    const m={}; hdrs.forEach((h,i)=>{ const k=String(h||'').trim(); if(k) m[k]=i+1; });
    const need = ['Brand','SO#','Odoo SO URL','SO Linked At','Custom Order Status','Design Request',
                  'Short Tag','3D Tracker','Order Folder','05-3D Folder','Client Folder','SO Shortcut in Client','00-Intake',
                  'Customer Name','EmailLower','PhoneNorm','RootApptID','APPT_ID'];
    let appended=false;
    need.forEach(n=>{ if(!m[n]){ sh.getRange(1, sh.getLastColumn()+1).setValue(n); m[n]=sh.getLastColumn(); appended=true; }});
    if (appended){
      const hdrs2 = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]||[];
      const mm={}; hdrs2.forEach((h,i)=>{ const k=String(h||'').trim(); if(k) mm[k]=i+1; }); return mm;
    }
    return m;
  })();
  const set = (c,v)=> H[c] && sh.getRange(row, H[c]).setValue(v);
  const get = (c)=> H[c] ? sh.getRange(row, H[c]).getValue() : '';

  const customer   = H['Customer Name']? sh.getRange(row, H['Customer Name']).getValue(): (H['Customer']? sh.getRange(row, H['Customer']).getValue():'');
  const email      = H['EmailLower']? sh.getRange(row, H['EmailLower']).getValue(): (H['Email']? sh.getRange(row, H['Email']).getValue():'');
  const phone      = H['PhoneNorm']? sh.getRange(row, H['PhoneNorm']).getValue(): (H['Phone']? sh.getRange(row, H['Phone']).getValue():'');
  const rootApptId = H['RootApptID']? sh.getRange(row, H['RootApptID']).getValue(): (H['APPT_ID']? sh.getRange(row, H['APPT_ID']).getValue():'');

  // derive ShortTag from Step‑1 if blank
  if (!shortTag){
    const f = designForm || {}; const pre = [f.Shape, f.RingStyle].filter(Boolean).join(' ').trim();
    shortTag = pre ? _titleCase_(_truncate_(pre,24)) : '';
  }
  if (!shortTag && typeof generateShortTag_==='function') shortTag = generateShortTag_(designRequest);

  // Write master (SO as text on this row)
  sh.getRange(1, H['SO#'], sh.getMaxRows(), 1).setNumberFormat('@'); // force text format
  set('Brand', brand); set('SO#', '\''+so); set('Odoo SO URL', url);
  const linkedAt = new Date(); set('SO Linked At', linkedAt);
  if (!String(get('Custom Order Status')||'').trim()){
    set('Custom Order Status', '3D Requested'); // don’t stomp non-empty status
  }
  if (designRequest) set('Design Request', designRequest);
  if (shortTag)      set('Short Tag', shortTag);

  const masterLink = ss.getUrl() + '#gid=' + sh.getSheetId() + '&range=A' + row;

  // Drive scaffolding + tracker (single copy-init)
  let orderFolderLink='', threeDFolderLink='', trackerUrl='', trackerId='';
  let orderFolderId='', intakeFolderId='';
  try{
    const drv = scaffoldOrderFolders({ brand, so, customer, rootApptId, shortTag }); // existing helper
    orderFolderId   = drv.orderFolderId || '';
    orderFolderLink = drv.orderFolderLink || '';
    threeDFolderLink= drv.threeDFolderLink || '';
    intakeFolderId  = drv.intakeFolderId || '';
    if (orderFolderLink)  set('Order Folder', orderFolderLink);
    if (threeDFolderLink) set('05-3D Folder', threeDFolderLink);
    if (drv.intakeFolderLink) set('00-Intake', drv.intakeFolderLink);

    // Create client shortcut to SO folder (INSIDE client folder)
    try {
      const sc = createClientShortcutForSO_({ brand, so, orderFolderId, customer, email });
      if (sc && sc.clientFolderLink) set('Client Folder', sc.clientFolderLink);
      if (sc && sc.shortcutLink)     set('SO Shortcut in Client', sc.shortcutLink);
    } catch(e){ Logger.log('Client shortcut skipped: ' + e.message); }

    // Move prospect/intake docs into 00‑Intake under SO
    try {
      const movedUrl = moveApFolderToIntake_({ brand, customer, email, rootApptId, intakeFolderId }); // existing helper
      if (movedUrl) set('00-Intake', movedUrl);
    } catch(e){ Logger.log('Move Intake skipped: ' + e.message); }

    // Copy/init 3D Tracker once
    const tracker = copy3DTrackerToSO_(drv.threeDFolderId, brand, so); // expected {id,url}
    trackerId = tracker.id || ''; trackerUrl = tracker.url || '';
    if (trackerUrl) set('3D Tracker', trackerUrl);

    // Append Start‑3D log (header‑mapped; Revision # auto = 1)
    if (trackerId){
      append3DTrackerLog_({
        trackerId: trackerId,
        action: 'Start 3D Initiated',
        form: Object.assign({}, designForm, { Mode: 'Start 3D Design / Create New SO' }),
        brand, so, odooUrl: url, masterLink, shortTag
      });
    }
  }catch(e){ Logger.log('Drive/Tracker setup skipped: ' + e.message); }

  // === NEW: Propagate the SO to other rows with the SAME RootApptID (safe OR checks) ===
  const propSummary = propagateSOToSiblingRows_({
    sh, H, row, brand, so, url, linkedAt,
    requireSecondaryMatch: REQUIRE_SECONDARY_MATCH,
    forceOverwrite
  });

  SpreadsheetApp.flush();

  const tz = Session.getScriptTimeZone?.() || 'America/Los_Angeles';
  const linkedAtStr = Utilities.formatDate(linkedAt, tz, 'MMM d, yyyy h:mm a z');

  // === COS reminder for Assign SO / Start 3D (no restart) ===
  try {
    // SO (strip any leading apostrophe)
    var soNum = String((H['SO#'] ? sh.getRange(row, H['SO#']).getDisplayValue() : so) || '')
                  .replace(/^'+/, '').trim();

    // Customer (use the already-read variable if present, else read from sheet)
    var custName = String((typeof customer !== 'undefined' && customer) ?
                          customer :
                          (H['Customer Name'] ? sh.getRange(row, H['Customer Name']).getDisplayValue() :
                          (H['Customer'] ? sh.getRange(row, H['Customer']).getDisplayValue() : ''))).trim();

    // Assigned / Assisted / Next Steps
    var assignedRepName = String(H['Assigned Rep'] ? sh.getRange(row, H['Assigned Rep']).getDisplayValue() : '').trim();
    var assistedRepName = String(H['Assisted Rep'] ? sh.getRange(row, H['Assisted Rep']).getDisplayValue() : '').trim();
    var nextSteps       = String(H['Next Steps']   ? sh.getRange(row, H['Next Steps']).getDisplayValue()   : '').trim();

    if (soNum) {
      Remind.scheduleCOS(soNum, {
        customerName:     custName,
        assignedRepName:  assignedRepName,
        assistedRepName:  assistedRepName,
        nextSteps:        nextSteps
      }, false); // restart = false for Assign SO / Start 3D
    }
  } catch (e) {
    console.warn('Remind.scheduleCOS (Assign SO/Start3D) failed:', e && e.message ? e.message : e);
  }

  return {
    ok:true,
    summary:{
      brand, so, customer, email, phone,
      linkedAt: linkedAtStr, masterLink,
      orderFolderLink, threeDFolderLink,
      trackerUrl,
      shortTag,
      propagation: propSummary
    }
  };
}

// ---------- Sibling propagation (RootApptID OR safety) ----------
function propagateSOToSiblingRows_({ sh, H, row, brand, so, url, linkedAt, requireSecondaryMatch, forceOverwrite }){
  const last = sh.getLastRow();
  if (last < 3) return { updated:0, skipped:0, updatedRows:[], skippedRows:[] };

  const col = name => (H[name] ? H[name] : 0);
  const cSO     = col('SO#');
  const cURL    = col('Odoo SO URL') || col('SO URL') || col('Odoo Link');
  const cLinked = col('SO Linked At');
  const cStatus = col('Custom Order Status') || col('Order Status');
  const cRoot   = col('RootApptID') || col('APPT_ID');
  const cEmail  = col('EmailLower') || col('Email');
  const cPhone  = col('PhoneNorm')  || col('Phone');
  const cName   = col('Customer Name') || col('Customer') || col('Name');

  if (!cSO || !cURL || !cLinked || !cRoot) {
    return { updated:0, skipped:0, updatedRows:[], skippedRows:[], note:'Missing columns to propagate' };
  }

  const rows = last - 1;
  const headerCount = Math.max(1, sh.getLastColumn());
  // Read only needed columns for speed
  const rootVals  = sh.getRange(2, cRoot,  rows, 1).getDisplayValues();
  const emailVals = cEmail ? sh.getRange(2, cEmail, rows, 1).getDisplayValues() : null;
  const phoneVals = cPhone ? sh.getRange(2, cPhone, rows, 1).getDisplayValues() : null;
  const nameVals  = cName  ? sh.getRange(2, cName,  rows, 1).getDisplayValues() : null;

  const currentRoot = String(sh.getRange(row, cRoot).getDisplayValue() || '').trim();
  const meEmail = cEmail ? String(sh.getRange(row, cEmail).getDisplayValue()||'') : '';
  const mePhone = cPhone ? String(sh.getRange(row, cPhone).getDisplayValue()||'') : '';
  const meName  = cName  ? String(sh.getRange(row, cName ).getDisplayValue()||'') : '';

  const emailNormMe = normalizeEmail_(meEmail);
  const phoneNormMe = normalizePhoneDigits_(mePhone);
  const nameTokensMe= nameTokens_(meName);

  let updated=0, skipped=0;
  const updatedRows=[], skippedRows=[];

  for (let i=0;i<rows;i++){
    const rr = i + 2;
    if (rr === row) continue;

    const root = String(rootVals[i][0]||'').trim();
    if (!root || root !== currentRoot) continue; // only same-root siblings

    // Secondary OR checks
    const e2 = cEmail ? String(emailVals[i][0]||'') : '';
    const p2 = cPhone ? String(phoneVals[i][0]||'') : '';
    const n2 = cName  ? String(nameVals[i][0] ||'') : '';

    const emailOK = emailNormMe && normalizeEmail_(e2) && (normalizeEmail_(e2) === emailNormMe);
    const phoneOK = phoneMatch_(phoneNormMe, normalizePhoneDigits_(p2));
    const nameOK  = nameMatch_(nameTokensMe, nameTokens_(n2));

    const passes = forceOverwrite || !requireSecondaryMatch || (emailOK || phoneOK || nameOK);

    if (!passes){
      skipped++; skippedRows.push(rr); continue;
    }

    // Write SO, URL, LinkedAt to sibling row
    if (cSO)     sh.getRange(rr, cSO).setValue("'"+so);
    if (cURL)    sh.getRange(rr, cURL).setValue(url);
    if (cLinked) sh.getRange(rr, cLinked).setValue(linkedAt);

    // Only set status if blank — avoid stomping a later stage
    if (cStatus){
      const st = String(sh.getRange(rr, cStatus).getDisplayValue()||'').trim();
      if (!st) sh.getRange(rr, cStatus).setValue('3D Requested');
    }

    updated++; updatedRows.push(rr);
  }

  return { updated, skipped, updatedRows, skippedRows };
}

// ---------- Normalizers & matchers (robust but conservative) ----------
function normalizeEmail_(e){
  e = String(e||'').trim().toLowerCase();
  if (!e || e.indexOf('@') === -1) return '';
  const parts = e.split('@'); if (parts.length !== 2) return e;
  let local = parts[0], domain = parts[1];
  // strip +tag
  local = local.replace(/\+.*$/, '');
  // gmail: dot-insensitive
  if (domain === 'gmail.com' || domain === 'googlemail.com') {
    local = local.replace(/\./g, '');
    domain = 'gmail.com';
  }
  return local + '@' + domain;
}
function normalizePhoneDigits_(p){ return String(p||'').replace(/\D+/g,''); }
function phoneMatch_(aDigits, bDigits){
  if (!aDigits || !bDigits) return false;
  if (aDigits === bDigits) return true;
  // handle country code: compare last 7–10 digits
  const a10 = aDigits.slice(-10), b10 = bDigits.slice(-10);
  if (a10 && b10 && a10 === b10) return true;
  const a7 = aDigits.slice(-7), b7 = bDigits.slice(-7);
  return !!(a7 && b7 && a7 === b7);
}
function nameTokens_(s){
  s = String(s||'').toLowerCase();
  if (!s) return [];
  // remove punctuation, split into tokens, drop short tokens and honorifics
  const raw = s.replace(/[^a-z0-9\s]/g,' ').split(/\s+/).filter(Boolean);
  const stop = new Set(['mr','mrs','ms','miss','dr','jr','sr','ii','iii','iv']);
  return raw.filter(t => t.length >= 2 && !stop.has(t));
}
function nameMatch_(tokensA, tokensB){
  if (!tokensA.length || !tokensB.length) return false;
  // intersection count
  const setB = new Set(tokensB);
  let hits = 0; for (const t of tokensA){ if (setB.has(t)) hits++; }
  if (hits >= 2) return true; // strong signal: two token overlap
  // containment fallback for 1-token names, but require length >=5 to avoid "kevin"
  const a = tokensA.join(' '), b = tokensB.join(' ');
  if (a.length >= 5 && (b.indexOf(a) !== -1)) return true;
  if (b.length >= 5 && (a.indexOf(b) !== -1)) return true;
  return false;
}

// ---------- Tracker writer (HEADER‑BASED + REVISION #) ----------
function append3DTrackerLog_({ trackerId, action, form, brand, so, odooUrl, masterLink, shortTag }){
  if (!trackerId) throw new Error('append3DTrackerLog_: missing trackerId');
  const ssT = SpreadsheetApp.openById(trackerId);
  const shT = ssT.getSheetByName('Log') || ssT.insertSheet('Log');

  // Desired schema (order agnostic — we map by header)
  const desired = [
    'Timestamp','User','Action','Revision #',
    'Mode','Accent Type','Ring Style','Metal','US Size','Band Width (mm)',
    'Center Type','Shape','Diamond Dimension','Design Notes',
    'Short Tag','SO#','Brand','Odoo SO URL','Master Link'
  ];

  // Ensure headers exist; append missing headers to the right
  let lastCol = Math.max(1, shT.getLastColumn());
  let header = (shT.getLastRow() >= 1) ? (shT.getRange(1,1,1,lastCol).getValues()[0] || []).map(h=>String(h||'').trim()) : [];
  if (header.length === 0){ header = desired.slice(); shT.getRange(1,1,1,header.length).setValues([header]); }
  else {
    const missing = desired.filter(h => header.indexOf(h) === -1);
    if (missing.length){ shT.getRange(1, header.length+1, 1, missing.length).setValues([missing]); header = header.concat(missing); }
  }
  const pos = {}; header.forEach((h,i)=>{ pos[h]=i+1; });

  // Compute next Revision # for this SO
  let revNo = 1;
  try {
    const last = shT.getLastRow();
    if (last >= 2 && pos['SO#'] && pos['Revision #']){
      const rows = shT.getRange(2,1,last-1,Math.max(1,shT.getLastColumn())).getValues();
      const iSO = pos['SO#'] - 1, iREV = pos['Revision #'] - 1;
      let maxRev = 0;
      for (let i=0;i<rows.length;i++){
        const soPrev = String(rows[i][iSO]||'').trim();
        const rv = Number(rows[i][iREV] || 0);
        if (soPrev === so && rv > maxRev) maxRev = rv;
      }
      revNo = maxRev ? (maxRev + 1) : 1;
    }
  } catch(e){ Logger.log('Rev# compute skipped: ' + e.message); }

  const actor = Session.getActiveUser?.().getEmail?.() || Session.getEffectiveUser?.().getEmail?.() || '';
  const now   = new Date();

  const mapObj = {
    'Timestamp': now,
    'User': actor,
    'Action': action || 'Start 3D Initiated',
    'Revision #': revNo,
    'Mode': (form && form.Mode) || '',
    'Accent Type': (form && (form.AccentDiamondType || form['Accent Type'])) || '',
    'Ring Style' : (form && (form.RingStyle || form['Ring Style'])) || '',
    'Metal'      : (form && form.Metal) || '',
    'US Size'    : (form && (form.USSize || form['US Size'])) || '',
    'Band Width (mm)': (form && (form.BandWidthMM || form['Band Width (mm)'])) || '',
    'Center Type': (form && (form.CenterDiamondType || form['Center Type'])) || '',
    'Shape'      : (form && form.Shape) || '',
    'Diamond Dimension': (form && (form.DiamondDimension || form['Diamond Dimension'])) || '',
    'Design Notes': (form && (form.DesignNotes || form['Design Notes'])) || '',
    'Short Tag'  : shortTag || '',
    'SO#'        : so || '',
    'Brand'      : brand || '',
    'Odoo SO URL': odooUrl || '',
    'Master Link': masterLink || ''
  };

  // Build row aligned to header
  const row = new Array(header.length);
  for (let c=0;c<header.length;c++){ const key = header[c]; row[c] = mapObj.hasOwnProperty(key) ? mapObj[key] : ''; }
  shT.getRange(shT.getLastRow()+1, 1, 1, row.length).setValues([row]);
}

// ---------- Client shortcut helpers (INSIDE client folder) ----------
function getBrandClientsRootId_(brand){
  const props = PropertiesService.getScriptProperties();
  let id = '';
  if (brand === 'HPUSA') {
    id = (props.getProperty('HP_ClientsRootID') || props.getProperty('HP_CLIENTS_ROOT_ID') || '').trim();
  } else {
    id = (props.getProperty('VVS_ClientsRootID') || props.getProperty('VVS_CLIENTS_ROOT_ID') || '').trim();
  }
  if (!id) throw new Error('Missing Script Property: clients root for ' + brand);
  return id;
}
function safeFolderName_(s){ return String(s || '').replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim(); }
function findOrCreateClientFolder_({ brand, customer, email }) {
  const clientsRootId = getBrandClientsRootId_(brand);
  const parent = DriveApp.getFolderById(clientsRootId);
  const baseName = safeFolderName_(customer || '').trim();
  const name = baseName || safeFolderName_(email ? String(email).split('@')[0] : 'Client');
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) { const f = it.next(); return { id: f.getId(), link: f.getUrl(), existed: true }; }
  const created = parent.createFolder(name);
  return { id: created.getId(), link: created.getUrl(), existed: false };
}
function createDriveShortcut_({ targetId, parentId, name }){
  // Pre-check to avoid duplicates (Advanced Drive)
  try{
    const q = [
      "'" + parentId + "' in parents",
      "mimeType = 'application/vnd.google-apps.shortcut'",
      "trashed = false",
      "title = '" + String(name).replace(/'/g,"\\'") + "'"
    ].join(' and ');
    const res = Drive.Files.list({ q, maxResults: 5, corpora:'allDrives', includeTeamDriveItems:true, supportsAllDrives:true });
    if (res?.items?.length){
      const f = res.items[0];
      return {
        id: f.id,
        openLink: 'https://drive.google.com/drive/folders/' + targetId,
        shortcutFileLink: f.alternateLink || f.webViewLink
      };
    }
  }catch(e){ Logger.log('shortcut precheck skipped: ' + e.message); }
  const file = Drive.Files.insert({
    title: name,
    mimeType: 'application/vnd.google-apps.shortcut',
    parents: [{ id: parentId }],
    shortcutDetails: { targetId }
  }, null, { supportsAllDrives: true });
  return {
    id: file.id,
    openLink: 'https://drive.google.com/drive/folders/' + targetId,
    shortcutFileLink: file.alternateLink || file.webViewLink
  };
}
function createClientShortcutForSO_({ brand, so, orderFolderId, customer, email }){
  const client = findOrCreateClientFolder_({ brand, customer, email });
  const sc = createDriveShortcut_({ targetId: orderFolderId, parentId: client.id, name: `SO${so} (shortcut)` });
  return { clientFolderLink: client.link, shortcutLink: sc.openLink };
}

// ---------- Small utils ----------
function _titleCase_(s){ return String(s||'').toLowerCase().replace(/\b[\p{L}’']+/gu, w => w[0].toUpperCase()+w.slice(1)); }
function _truncate_(s,n){ s=String(s||''); return s.length<=n?s:s.slice(0,n).replace(/\s+\S*$/,'').trim(); }


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
