function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  // ⏰ Reminders submenu
  const remindersMenu = ui.createMenu('⏰ Reminders')
    .addItem('💤 Snooze selected order', 'remind_menu_snoozeSelected')
    .addItem('💤 Snooze today (1-click)', 'remind_menu_snoozeTodaySelected')
    .addItem('↩️ Unsnooze now (1-click)', 'remind_menu_unsnoozeNowSelected')
    .addItem('❌ Cancel selected order', 'remind_menu_cancelSelected');

  // 💎 Diamond Viewing submenu
  const diamondsMenu = ui.createMenu('💎 Diamond Viewing')
    .addItem('💎 Propose', 'dp_openProposeDiamonds')
    .addItem('💎 Order/Approve', 'dp_openOrderApproveDiamonds')
    .addItem('💎 Confirm Delivery', 'dp_openConfirmDeliveryDiamonds')
    .addItem('💎 Client Decisions', 'dp_openStoneDecisions');

  // 🧾 Update Quotation submenu
  const updateQuotationMenu = ui.createMenu('🧾 Update Quotation')
    .addItem('💎 Diamonds', 'uq_openUpdateQuotationDiamonds')
    .addItem('💍 Ring Setting', 'uq_openUpdateQuotationSettings');

  // 🕯️ Wax Print submenu
  const waxMenu = ui.createMenu('🕯️ Wax Print')
    .addItem('📥 Pull Pending', 'wax_adminOpenDialog_');

  // ✅ Acknowledgements submenu (Phase B + Phase C actions)
  // Requires functions from Phase B: runAllPipes, buildRootIndex, buildRepsMap, recomputeAckStatusSummary
  // And from Phase C: buildTodaysQueuesAll, openMyQueue, refreshMyQueue, submitMyQueue
  const ackMenu = ui.createMenu('✅ Acknowledgements')
    .addItem('🧰 Build Today’s Queues (All)', 'buildTodaysQueuesAll')
    .addItem('🙋 My Queue (Detect Me)', 'openMyQueue')
    .addItem('🔄 Refresh My Queue', 'refreshMyQueueHybrid')
    .addItem('✅ Submit My Queue', 'submitMyQueueUnified')
    .addItem('📊 Refresh 09_Ack_Dashboard', 'buildAckDashboard')
    .addSeparator()
    .addItem('🧮 Recompute 00 Ack Status (from log)', 'recomputeAckStatusSummary');

  // 💎 Sales — primary workflow
  ui.createMenu('💎 Sales')
    .addItem('Authorize Once', 'authorizeDriveOnce')
    .addItem('🚀 Start 3D Design', 'openStart3D')
    .addItem('🔗 Assign SO', 'assignSO')
    .addItem('♻️ 3D Revision Request', 'open3DRevision')
    .addItem('📆 Record Deadline', 'showRecordDeadlineDialog')
    .addSeparator()
    .addSubMenu(diamondsMenu)
    .addSubMenu(updateQuotationMenu)
    .addSubMenu(waxMenu)
    .addSeparator()
    .addItem('🧑‍💼 Client Status Update', 'cs_openStatusDialog_')
    .addItem('📁 Client Summary', 'openClientSummary')
    .addSeparator()
    .addItem('💵 Record Payment', 'openRecordPayment')
    .addItem('🧾 Payment Summary', 'openPaymentSummary')
    .addSeparator()
    .addSubMenu(remindersMenu)
    .addToUi();

  // 📈 Reports — quick access
  ui.createMenu('📈 Reports')
    .addItem('📋 By Status', 'openReportByStatus')
    .addItem('🙋 By Rep', 'openReportByRep')
    .addSubMenu(
      ui.createMenu('Quick PDF Exports')
        .addItem('Booked Appointment (Sales Stage)', 'report_menu_export_BookedAppointment')
        .addItem('Viewing Scheduled (Conversion)',   'report_menu_export_ViewingScheduled')
        .addItem('Deposit Paid (Conversion)',        'report_menu_export_DepositPaid')
        .addItem('In Production (Custom Order)',     'report_menu_export_InProduction')
    )
    .addSeparator()
    .addItem('📆 Appointment Summary', 'as_openAppointmentSummary')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('📊 Stage Rollup')
        .addItem('Refresh now', 'refreshClientStageRollup')
    )
    .addToUi();

  // 🧹 Audit — data quality checks
  ui.createMenu('🧹 Audit')
    .addItem('🔍 Run Master Audit (v1)', 'runMasterAuditV1')
    .addToUi();

  // ✅ Acknowledgements — top-level menu
  ackMenu.addToUi();
}


// Optional: makes the menu appear immediately when installed as an add-on
function onInstall(e) { onOpen(e);}

function authorizeDriveOnce() {
  // Touch Drive to request the Drive scope for this user
  DriveApp.getRootFolder().getName();
  SpreadsheetApp.getActive().toast('Drive access authorized. You can close this.');
}


function openRecordPayment() {
  try { rp_markActiveMasterRowIndex_(); } catch(_){}
  const html = HtmlService.createHtmlOutputFromFile('dlg_record_payment_v1').setWidth(980).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Record Payment');
}

function openPaymentSummary() {
  // Capture current selection like “Record Payment” does; fall back gracefully
  try { rp_markActiveMasterRowIndex_(); } catch (e) {
    try { rp_captureSelectionForPrefill_(); } catch (_) {}
  }
  const html = HtmlService.createHtmlOutputFromFile('dlg_payment_summary_v1')
                .setTitle('Payment Summary');
  SpreadsheetApp.getUi().showModalDialog(html, 'Payment Summary');
}

// Opens the "By Status" dialog (stub)
function openReportByStatus() {
  const t = HtmlService.createTemplateFromFile('dlg_report_status_v1');
  t.lists = report_getDropdownLists_();  // inject all four status lists (and reps, ignored here)
  const html = t.evaluate().setWidth(1150).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Report — By Status');
}


function openReportByRep() {
  const t = HtmlService.createTemplateFromFile('dlg_report_reps_v1');
  const lists = report_getDropdownLists_();           // must return .assignedReps and .assistedReps
  t.assignedReps = lists.assignedReps || [];
  t.assistedReps = (lists.assistedReps && lists.assistedReps.length) ? lists.assistedReps : (lists.assignedReps || []);
  const html = t.evaluate().setWidth(780).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Report — By Rep');
}

// Simple ping for client→server sanity check
function report_ping() {
  return 'pong';
}


function seedConfigProps_() {
const LEDGER_ID = '1omobZq0MOTB8l8IkLemlCibukm0Pphcxh04kSG0KQBQ'; // 400_Payments Ledger ID
const props = PropertiesService.getScriptProperties();
const seed = { LEDGER_FILE_ID: LEDGER_ID.trim() };
['HPUSA_301_FILE_ID','VVS_302_FILE_ID','VVS_ClientsRootID','HP_ClientsRootID','HPUSA_SO_ROOT_FOLDER_ID','VVS_SO_ROOT_FOLDER_ID']
  .forEach(k => { if (!props.getProperty(k)) seed[k] = ''; });




// Non-destructive: only set the ones that are missing.
// DO NOT pass 'true' to setProperties or you will wipe existing keys.
Object.keys(seed).forEach(k => {
  if (!props.getProperty(k)) props.setProperty(k, seed[k]);
});
SpreadsheetApp.getUi().alert('Config seeded (non-destructive). Existing keys preserved.');
SpreadsheetApp.getUi().alert('Config seeded: LEDGER_FILE_ID set.');
}


function healthCheck() {
const ui = SpreadsheetApp.getUi(), ok = [], warn = [], fail = [];
const ss = SpreadsheetApp.getActive();
ok.push(`Active: ${ss.getName()}`);


const MASTER_NAME = '00_Master Appointments'; // hardcoded for v1
const master = ss.getSheetByName(MASTER_NAME);
if (!master) {
  const tabs = ss.getSheets().map(s => s.getName()).join(' | ');
  fail.push(`Missing master sheet "${MASTER_NAME}". Available: ${tabs}`);
} else {
  const headers = (r=>{const m={}; r.forEach((h,i)=>{ if(h) m[String(h).trim()]=i+1; }); return m;})
    (master.getRange(1,1,1,master.getLastColumn()).getValues()[0]);
  const need = ['APPT_ID','RootApptID','Brand','Customer Name','EmailLower','PhoneNorm'];
  const miss = need.filter(h => !(h in headers));
  miss.length ? warn.push('Master missing headers: ' + miss.join(', '))
              : ok.push(`Master "${master.getName()}" OK.`);
}




const props = PropertiesService.getScriptProperties();
const ledgerId = (props.getProperty('LEDGER_FILE_ID') || '').trim();
if (!ledgerId) {
  fail.push('LEDGER_FILE_ID missing. Click Sales → Seed Config (temp).');
} else {
  try {
    const led = SpreadsheetApp.openById(ledgerId);
    ok.push('400_Payments Ledger reachable: ' + led.getName());
    if (!led.getSheetByName('Payments')) warn.push('Ledger has no "Payments" tab yet (okay for now).');
  } catch (e) {
    fail.push('Could not open 400_Payments Ledger by ID.');
  }
}




ui.alert(
  'Health Check (v1)\n\n' +
  (ok.length   ? '✅ Checks:\n- ' + ok.join('\n- ') + '\n\n' : '') +
  (warn.length ? '⚠️ Warnings:\n- ' + warn.join('\n- ') + '\n\n' : '') +
  (fail.length ? '❌ Failures:\n- ' + fail.join('\n- ') + '\n\n' : '') +
  ss.getUrl()
);
}




function setMasterSheetName_() {
const ui = SpreadsheetApp.getUi();
const ss = SpreadsheetApp.getActive();
// Offer common guesses to make it easy
const guesses = ['00_Master','00_Master Appointments','00 – Master','00 Master'];
const existing = ss.getSheets().map(s => s.getName());
const prompt = ui.prompt(
  'Set Master Sheet Name',
  'Type the exact tab name for your master sheet.\n\nExamples:\n- 00_Master\n- 00_Master Appointments\n\nAvailable tabs:\n' + existing.join(' | '),
  ui.ButtonSet.OK_CANCEL
);
if (prompt.getSelectedButton() !== ui.Button.OK) return;
const name = prompt.getResponseText().trim();
if (!name) { ui.alert('No name entered.'); return; }
if (!ss.getSheetByName(name)) { ui.alert('No tab named "' + name + '" in this file.'); return; }
PropertiesService.getScriptProperties().setProperty(CFG.MASTER_SHEET_NAME, name);
ui.alert('Saved. MASTER_SHEET_NAME = "' + name + '". Run Health Check (v1) next.');
}

function openStart3D(){
  const ui = SpreadsheetApp.getUi();
  // Try to precompute; if user isn’t on a valid row/sheet, don’t block the dialog.
  let bootstrap = null;
  try {
    bootstrap = start3d_init(); // returns {ok,brand,so,hasSO,hasDesignRequest}
  } catch (e) {
    bootstrap = null; // fall back to client-side call; preserves existing behavior
  }

  const t = HtmlService.createTemplateFromFile('dlg_start3d_v1'); // no ".html"
  t.BOOTSTRAP = bootstrap;

  const html = t.evaluate().setWidth(650).setHeight(400);
  ui.showModalDialog(html, 'Start 3D Design / Create New SO');
}


/**
* Called by the dialog when you click "Next".
* `form` is a plain object with the field values.
*/
function start3D_nextStep_(form) {
// ---- do whatever you need here (validate, write to sheet, open next dialog) ----
// For now, just log so we know it fired:
Logger.log('start3D_nextStep_ payload: ' + JSON.stringify(form));
return { ok: true, message: 'Captured form.' };
}




// Optional: lets the dialog send logs to Execution Log
function logClient_(x) { Logger.log('[CLIENT]', JSON.stringify(x)); }




/**
* Build the Odoo paste text for Step 2.
* Returns a STRING to the HTML success handler.
*/




function ping() { return 'pong'; }




// round‑trip echo for testing
function echoDesignForm_(payload) {
// payload is an object from the client
return { ok: true, received: payload };
}


function buildOdooPaste_(f){
// f = {accType, ringStyle, metal, size, band, notes, centerType, shape, dim}
const lines = (f.notes || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
const note1 = lines[0] || '';
const note2 = lines[1] || '';
const note3 = lines[2] || '';
return [
  'RING SETTING SPECS',
  `+ Diamond Type:  ${f.accType || ''}`,
  `+ Ring Style: ${f.ringStyle || ''}`,
  `+ Metal: ${f.metal || ''}`,
  `+ US Ring Size : ${f.size || ''}`,
  `+ Band Width: ${f.band || ''}`,
  'NOTE:',
  `1. ${note1}`,
  `2. ${note2}`,
  `3. ${note3}`,
  '',
  'STONE INFO',
  `+ Diamond Type:  ${f.centerType || ''}`,
  `+ Shape: ${f.shape || ''}`,
  `+ Diamond Dimension: ${f.dim || ''}`
].join('\n');
}


function generateOdooPaste_(payload){
return {
  ok: true,
  mode: payload.mode,
  odooPaste: buildOdooPaste_(payload)
};
}


function assignSO_legacyPrompts() {
const ui = SpreadsheetApp.getUi();




// 1) Gather inputs
const brand = ui.prompt('Link SO — Step 1/2', 'Brand? (HPUSA or VVS)', ui.ButtonSet.OK_CANCEL).getResponseText().trim();
if (!/^HPUSA|VVS$/i.test(brand)) { ui.alert('Please type exactly HPUSA or VVS.'); return; }




const soNum = ui.prompt('Link SO — Step 2/2', 'SO# (numbers only, e.g., 1234):', ui.ButtonSet.OK_CANCEL).getResponseText().trim();
if (!/^\d+$/.test(soNum)) { ui.alert('SO# should be digits only.'); return; }




const soUrl = ui.prompt('Optional', 'Paste Odoo SO URL (or leave blank):', ui.ButtonSet.OK_CANCEL).getResponseText().trim();
const key = ui.prompt('Find Row', 'Customer Email OR APPT_ID:', ui.ButtonSet.OK_CANCEL).getResponseText().trim();
if (!key) { ui.alert('Need an email or APPT_ID to find the row.'); return; }




// 2) Find the row in 00_Master Appointments (email match OR APPT_ID)
const ss = SpreadsheetApp.getActive();
const sh = ss.getSheetByName('00_Master Appointments');
const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
const H = headers.reduce((m,h,i)=>{ m[String(h).trim()] = i+1; return m; },{});




const last = sh.getLastRow(); if (last < 2) { ui.alert('No data rows.'); return; }
let row = 0;
const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
for (let i=0;i<vals.length;i++){
  const r = vals[i];
  const emailLower = (H['EmailLower'] ? (r[H['EmailLower']-1]||'') : '').toString().toLowerCase();
  const appt = (H['APPT_ID'] ? (r[H['APPT_ID']-1]||'') : '').toString();
  if (emailLower === key.toLowerCase() || appt === key) { row = i+2; break; }
}
if (!row) { ui.alert('Could not find a row by that Email or APPT_ID.'); return; }


// 3) Stamp SO fields (first test = write-only, no Drive ops)
function set(col,value){ if (H[col]) sh.getRange(row, H[col]).setValue(value); }
if (!H['SO#']) sh.getRange(1, sh.getLastColumn()+1).setValue('SO#'); // helper if header missing
if (!H['Odoo SO URL']) sh.getRange(1, sh.getLastColumn()+1).setValue('Odoo SO URL');
if (!H['SO Linked At']) sh.getRange(1, sh.getLastColumn()+1).setValue('SO Linked At');


// Rebuild header map if we added columns
if (!H['SO#'] || !H['Odoo SO URL'] || !H['SO Linked At']) {
  const hdrs = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  for (let i=0;i<hdrs.length;i++) H[String(hdrs[i]).trim()] = i+1;
}


set('SO#', soNum);
set('Odoo SO URL', soUrl);
set('Brand', String(brand).toUpperCase());
set('SO Linked At', new Date());


ui.alert('Linked! (SO stamped — no folders moved yet)');
}










function handleAssignSO(data) {
const brand = String(data && data.brand || '').toUpperCase();
const soNum = String(data && data.soNum  || '').trim();
const soUrl = String(data && data.soUrl  || '').trim();
const key   = String(data && data.key    || '').trim();


if (!/^(HPUSA|VVS)$/.test(brand)) throw new Error('Brand must be HPUSA or VVS');
if (!/^\d{2}\.\d{4}$/.test(soNum)) throw new Error('SO# must be like 12.3456');
if (!key)                         throw new Error('Need email or APPT_ID');




Logger.log(JSON.stringify({ brand, soNum, soUrl, key }, null, 2));
return 'assignSO captured (log-only).';
}


function ensureNoDuplicateSOInMaster_({ ss, sh, H, brand, so, currentRow }) {
// If headers aren't in place yet, skip the check safely.
if (!H['Brand'] || !H['SO#']) return;


const last = sh.getLastRow();
if (last < 2) return;


// Batch read columns for speed
const rows = last - 1;
const brandVals = sh.getRange(2, H['Brand'], rows, 1).getValues();
const soVals    = sh.getRange(2, H['SO#'], rows, 1).getValues();


const sheetId = sh.getSheetId();
const urlBase = ss.getUrl();


for (let i = 0; i < rows; i++) {
  const r = i + 2;
  if (r === currentRow) continue;
  const b = String(brandVals[i][0] || '').toUpperCase().trim();
  let s   = String(soVals[i][0] || '').trim().replace(/^'/, ''); // strip leading apostrophe
  if (b === brand && s === so) {
    const link = urlBase + '#gid=' + sheetId + '&range=A' + r;
    throw new Error('This SO is already linked on row ' + r + '. Open: ' + link);
  }
}
}


function upsertSOIndex_(entry) {
const ss = SpreadsheetApp.getActive();
const name = 'SO Index (v1)';
const hdrs = ['Brand','SO#','Odoo SO URL','Master Row','Linked At','Updated At'];




// Ensure tab exists
let sh = ss.getSheetByName(name);
if (!sh) {
  sh = ss.insertSheet(name);
  sh.getRange(1,1,1,hdrs.length).setValues([hdrs]);
}




// Header map
const firstRow = sh.getRange(1,1,1,Math.max(sh.getLastColumn(), hdrs.length)).getValues()[0] || [];
const H = firstRow.reduce((m,h,i)=>{ if(String(h).trim()) m[String(h).trim()] = i+1; return m; }, {});
// Ensure all headers present
let appended = false;
hdrs.forEach(h=>{
  if (!H[h]) {
    sh.getRange(1, sh.getLastColumn()+1).setValue(h);
    H[h] = sh.getLastColumn();
    appended = true;
  }
});
if (appended) {
  const row1 = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  Object.keys(H).forEach(k=>delete H[k]);
  row1.forEach((h,i)=>{ if(String(h).trim()) H[String(h).trim()] = i+1; });
}
















// Always keep SO# column as text
sh.getRange(1, H['SO#'], sh.getMaxRows(), 1).setNumberFormat('@');
















// Look for existing Brand + SO#
const last = sh.getLastRow();
let hitRow = 0;
if (last >= 2) {
  const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  for (let i=0;i<vals.length;i++){
    const brand = String(vals[i][H['Brand']-1]||'').toUpperCase();
    const so    = String(vals[i][H['SO#']-1]||'').replace(/^'/,''); // strip leading apostrophe if present
    if (brand === String(entry.brand).toUpperCase() && so === String(entry.so)) {
      hitRow = i+2;
      break;
    }
  }
}

const now = new Date();
const set = (r, col, val)=> sh.getRange(r, H[col]).setValue(val);

if (!hitRow) {
  // Append
  hitRow = sh.getLastRow() + 1;
  set(hitRow,'Brand', entry.brand);
  set(hitRow,'SO#',   '\'' + entry.so);
  set(hitRow,'Odoo SO URL', entry.url);
  set(hitRow,'Master Row',   entry.masterLink || '');
  set(hitRow,'Linked At', now);
} else {
  // Update
  set(hitRow,'Odoo SO URL', entry.url);
  if (entry.masterLink) set(hitRow,'Master Row', entry.masterLink);
  set(hitRow,'Updated At', now);
}
}


// Header synonym resolver (handles slight naming differences)
function resolveHeader_(H, candidates) {
for (const c of candidates) if (H[c]) return c;
return null;
}



function openAssignSOWithDesign_(formOrText) {
// Accept either a pre-built text or the form object; prefer buildOdooPaste_ if available
let designText = '';
if (typeof formOrText === 'string') {
  designText = formOrText;
} else if (formOrText && formOrText.odooPaste) {
  designText = String(formOrText.odooPaste || '');
} else {
  try { designText = buildOdooPaste_(formOrText || {}); } catch (e) { designText = ''; }
}
















const t = HtmlService.createTemplateFromFile('dlg_assign_so_v1'); // switch to templated HTML
t.designRequest = designText;
const html = t.evaluate().setWidth(520).setHeight(360).setTitle('Assign Sales Order');
SpreadsheetApp.getUi().showModalDialog(html, 'Assign Sales Order');
}




function scaffoldOrderFolders({ brand, so, customer, rootApptId, shortTag }) {
// Resolve root folder for this brand
const props = PropertiesService.getScriptProperties();
const rootKey = (brand === 'HPUSA') ? 'HPUSA_SO_ROOT_FOLDER_ID' : 'VVS_SO_ROOT_FOLDER_ID';
const rootId = (props.getProperty(rootKey) || '').trim();
if (!rootId) throw new Error(`Missing Script Property: ${rootKey}. Set it in Project Settings → Script properties.`);




const parent = DriveApp.getFolderById(rootId);




// Build names (Option A)
const sanitize = s => String(s || '').replace(/[\\/<>\":|?*]/g, '').trim();
const tagPart = shortTag ? ` — ${shortTag}` : '';
const modernName   = sanitize(`${brand}–SO${so}${tagPart}`); // e.g., HPUSA–SO12.3456 — Oval Solitaire
const modernNoTag  = sanitize(`${brand}–SO${so}`);           // e.g., HPUSA–SO12.3456
const legacyName   = sanitize(`SO${so}`);                    // e.g., SO12.3456




// Prefer existing modern-with-tag > modern-no-tag > legacy
let orderFolder = null;
let it = parent.getFoldersByName(modernName);
if (it.hasNext()) orderFolder = it.next();




if (!orderFolder) {
  it = parent.getFoldersByName(modernNoTag);
  if (it.hasNext()) orderFolder = it.next();
}
if (!orderFolder) {
  it = parent.getFoldersByName(legacyName);
  if (it.hasNext()) orderFolder = it.next();
}




// Create if none found (use modernName; if no tag, this equals modernNoTag)
if (!orderFolder) {
  orderFolder = parent.createFolder(modernName);
} else {
  // ---- Rename-once migration (only if no tag is currently in the name) ----
  // If we found legacy or no-tag name AND we have a shortTag now, try to rename once.
  const currentName = orderFolder.getName();
  const hasTagAlready = /—\s+.+$/.test(currentName); // em-dash followed by tag
  const targetName = modernName;
















  if (shortTag && !hasTagAlready && currentName !== targetName) {
    // If another folder already has the target name, use that one instead of renaming to avoid duplicates.
    const dup = parent.getFoldersByName(targetName);
    if (dup.hasNext()) {
      // Prefer the existing correctly named folder
      orderFolder = dup.next();
    } else {
      // Safe rename using DriveApp
      try {
        orderFolder.setName(targetName);
      } catch (e) {
        Logger.log('Folder rename skipped: ' + e.message);
        // Non-fatal: continue with the found folder
      }
    }
  }
}
















// Ensure required subfolders exist (SOP v1)
const need = ['00-Intake','04-Deposit','05-3D','09-ReadyForPickup','10-Completed'];
const child = {};
need.forEach(n => {
  const c = orderFolder.getFoldersByName(n);
  child[n] = c.hasNext() ? c.next() : orderFolder.createFolder(n);
});
















return {
  orderFolderId: orderFolder.getId(),
  orderFolderLink: orderFolder.getUrl(),
  threeDFolderId: child['05-3D'].getId(),
  threeDFolderLink: child['05-3D'].getUrl(),
  intakeFolderId:   child['00-Intake'].getId(),
  intakeFolderLink: child['00-Intake'].getUrl()
};
}

function generateShortTag_(designText) {
try {
  designText = String(designText || '');
  const grab = (label) => {
    const re = new RegExp('^\\s*' + label + '\\s*:\\s*(.+)$', 'mi');
    const m = re.exec(designText);
    return m ? m[1].trim() : '';
  };
  const shape = grab('Shape');
  const ring  = grab('Ring Style');

  // Build
  let tag = [shape, ring].filter(Boolean).join(' ').trim();

  // Safety: strip any leaked labels (defensive)
  tag = tag.replace(/\b(?:mode|accent diamond|ring style|metal|us size|band width|design notes|center stone|diamond type|shape|dimension)\s*:\s*/gi, '').trim();

  // Normalize whitespace
  tag = tag.replace(/\s{2,}/g, ' ');

  // Title case + truncate
  tag = truncate_(titleCase_(tag), 24);
  return tag;
} catch(_) { return ''; }
}

function titleCase_(s){ return String(s||'').toLowerCase().replace(/\\b[\\p{L}’']+/gu, w => w[0].toUpperCase() + w.slice(1)); }
function truncate_(s, n){ s = String(s||''); return s.length <= n ? s : s.slice(0, n).replace(/\\s+\\S*$/, '').trim(); }

function normalizeIdFromUrlOrId_(s){
s = String(s || '').trim();
const m = s.match(/[-\w]{25,}/);
return m ? m[0] : s;
}


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



function safeFolderName_(s){
return String(s || '')
  .replace(/[\\/:*?"<>|]/g, ' ')
  .replace(/\s+/g, ' ')
  .trim();
}


function findOrCreateClientFolder_({ brand, customer, email }) {
const clientsRootId = getBrandClientsRootId_(brand);
const parent = DriveApp.getFolderById(clientsRootId);

// Use EXACT folder name == Customer Name (sanitized). No email suffix.
const baseName = safeFolderName_(customer || '').trim();
const name = baseName || safeFolderName_(email ? String(email).split('@')[0] : 'Client');

// Exact match under the brand clients root
const it = parent.getFoldersByName(name);
if (it.hasNext()) {
  const f = it.next();
  return { id: f.getId(), link: f.getUrl(), existed: true };
}


// Create exactly "[Customer Name]"
const created = parent.createFolder(name);
return { id: created.getId(), link: created.getUrl(), existed: false };
}

function createDriveShortcut_({ targetId, parentId, name }){
// pre-check for existing shortcut with same name in that client folder
try {
  const q = [
    "'" + parentId + "' in parents",
    "mimeType = 'application/vnd.google-apps.shortcut'",
    "trashed = false",
    "title = '" + String(name).replace(/'/g,"\\'") + "'"
  ].join(' and ');
  const res = Drive.Files.list({
    q, maxResults: 5,
    corpora: 'allDrives', includeTeamDriveItems: true, supportsAllDrives: true
  });
  if (res?.items?.length) {
    const f = res.items[0];
    return {
      id: f.id,
      // what we store in sheets (opens the actual SO folder)
      openLink: 'https://drive.google.com/drive/folders/' + targetId,
      // preview of the shortcut file (not used in sheet)
      shortcutFileLink: f.alternateLink || f.webViewLink
    };
  }
} catch(e){ Logger.log('shortcut precheck skipped: ' + e.message); }
















// create
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
const sc = createDriveShortcut_({
  targetId: orderFolderId,
  parentId: client.id,
  name: `SO${so} (shortcut)`
});
return {
  clientFolderLink: client.link,
  shortcutLink: sc.openLink
};
}
































function seedClientRootsOnce(){
const SP = PropertiesService.getScriptProperties();
const ask = (label)=> Browser.inputBox(label + '1kNvsFO6lP722jjxTm7KWHMSOCJjlzGbY', Browser.Buttons.OK_CANCEL);
// Example: for now you can paste the same ID you already use for HPUSA clients root and VVS clients root.
const hp = ask('HPUSA clients root ID');
const vvs= ask('VVS clients root ID');
if (hp !== 'cancel') {
  SP.setProperty('HP_ClientsRootID', hp.trim());
  SP.setProperty('HP_CLIENTS_ROOT_ID', hp.trim());
}
if (vvs !== 'cancel') {
  SP.setProperty('VVS_ClientsRootID', vvs.trim());
  SP.setProperty('VVS_CLIENTS_ROOT_ID', vvs.trim());
}
SpreadsheetApp.getUi().alert('Client roots saved.');
}
















function debug_shortcutSmokeTest(){
const brand = 'HPUSA';
const customer = 'Eric Lee';
const so = '12.3456';
// Paste the REAL order folder ID that was created under the SO root:
const orderFolderId = '1kNvsFO6lP722jjxTm7KWHMSOCJjlzGbY';
















const res = createClientShortcutForSO_({ brand, so, orderFolderId, customer, email: '' });
Logger.log(JSON.stringify(res, null, 2));
}
















function driveFindChildFolderByName_(parentId, name) {
var q = "mimeType='application/vnd.google-apps.folder' and trashed=false and '" +
        parentId + "' in parents";
var items = (Drive.Files.list({ q:q, maxResults:200 }).items) || [];
name = (name || '').toLowerCase();
for (var i = 0; i < items.length; i++) {
  if ((items[i].title || '').toLowerCase() === name) return items[i];
}
return null;
}
















function driveEnsureChildFolder_(parentId, name) {
var f = driveFindChildFolderByName_(parentId, name);
if (f) return f;
return Drive.Files.insert({
  title: name,
  mimeType: 'application/vnd.google-apps.folder',
  parents: [{ id: parentId }]
});
}
















/**
* Move the AP-… prospect folder from the Client Folder into
* Brand-SO# / 00-Intake. Creates 00-Intake if missing.
* Returns {moved:boolean, id?:string, intakeId?:string, reason?:string}
*/
function moveProspectFolderToIntake_(rootApptId, clientFolderId, orderFolderId) {
try {
  if (!rootApptId || !clientFolderId || !orderFolderId) {
    return { moved:false, reason:'missing ids' };
  }
















  // Find the AP-… folder directly under the Client Folder
  var q = "mimeType='application/vnd.google-apps.folder' and trashed=false and '" +
          clientFolderId + "' in parents";
  var kids = (Drive.Files.list({ q:q, maxResults:200 }).items) || [];
















  // Match by exact AP prefix if available, otherwise any AP-YYYYMMDD-### name
  var esc = rootApptId.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
  var re  = new RegExp('^' + esc);           // begins with AP-2025…
  var reAlt = /^AP-\d{8}-\d{3}/;             // fallback pattern
















  var ap = null;
  for (var i = 0; i < kids.length; i++) {
    var nm = kids[i].title || '';
    if (re.test(nm) || reAlt.test(nm)) { ap = kids[i]; break; }
  }
  if (!ap) return { moved:false, reason:'prospect folder not found' };
















  // Ensure 00-Intake under the Order Folder
  var intake = driveEnsureChildFolder_(orderFolderId, '00-Intake');
















  // Already there?
  var parentIds = (ap.parents || []).map(function(p){ return p.id; });
  if (parentIds.indexOf(intake.id) !== -1) {
    return { moved:true, id: ap.id, intakeId: intake.id };
  }
















  // Move by changing parents (remove old, add 00-Intake)
  Drive.Files.update(
    {},
    ap.id,
    null,
    { addParents: intake.id, removeParents: parentIds.join(',') }
  );
















  // Keep original AP folder name (no rename)
  return { moved:true, id: ap.id, intakeId: intake.id };
} catch (err) {
  Logger.log('moveProspectFolderToIntake_ error: ' + err);
  return { moved:false, reason:String(err) };
}
}
















/** Extracts a Drive ID from either an ID or a URL. */
function _driveIdFrom_(idOrUrl) {
if (!idOrUrl) return null;
if (/^https?:\/\//i.test(idOrUrl)) {
  var m = String(idOrUrl).match(/[-\w]{25,}/);
  return m ? m[0] : null;
}
return String(idOrUrl);
}
















/** Get a Drive folder from ID or URL (null-safe). */
function _getFolder_(idOrUrl) {
var id = _driveIdFrom_((idOrUrl || ''));
return id ? DriveApp.getFolderById(id) : null;
}
















/** Ensure a child folder exists by name under parent; return it. */
function _ensureSubfolder_(parentFolder, name) {
var it = parentFolder.getFoldersByName(name);
return it.hasNext() ? it.next() : parentFolder.createFolder(name);
}
















/** Prefer names in order; return the first existing subfolder or null. */
function _findFirstSubfolderByNames_(parent, names) {
for (var i = 0; i < names.length; i++) {
  var it = parent.getFoldersByName(names[i]);
  if (it.hasNext()) return it.next();
}
return null;
}
















/** Move a folder under a new parent (remove old parent link). */
function _moveFolderUnder_(childFolder, newParent) {
newParent.addFolder(childFolder);
// Remove from all other parents so it truly “moves”
var parents = childFolder.getParents();
while (parents.hasNext()) {
  var p = parents.next();
  if (p.getId() !== newParent.getId()) {
    p.removeFolder(childFolder);
  }
}
}
















/**
* Merge all children of src into dst, then trash src.
* (Files/folders are re-parented so there’s no duplication.)
*/
function _mergeFolderInto_(src, dst) {
// files
var files = [];
var fit = src.getFiles();
while (fit.hasNext()) files.push(fit.next());
files.forEach(function(f) {
  dst.addFile(f);
  src.removeFile(f);
});
// subfolders
var subs = [];
var sit = src.getFolders();
while (sit.hasNext()) subs.push(sit.next());
subs.forEach(function(sf) {
  dst.addFolder(sf);
  src.removeFolder(sf);
});
src.setTrashed(true);
}
















/**
* Move the “prospect/AP” folder from the client folder into the Brand–SO# folder
* as “00-Intake”. If "00-Intake" already exists in the SO folder, merge into it.
*
* @param {string} clientFolderIdOrUrl  “Client Folder” (sheet 100) value
* @param {string} soFolderId           Brand–SO# folder id
* @return {string|null} intake folder URL or null if nothing moved
*/
function moveProspectToIntake_(clientFolderIdOrUrl, soFolderId) {
var client = _getFolder_(clientFolderIdOrUrl);
var soFolder = _getFolder_(soFolderId);
if (!client || !soFolder) return null;
















// Look for common names you’ve used
var prospect = _findFirstSubfolderByNames_(client, ['AP', '00-Prospect', '00_Prospect', 'Prospect']);
if (!prospect) return null; // nothing to move
















// Make/locate 00-Intake inside SO folder
var intake = _ensureSubfolder_(soFolder, '00-Intake');
















// If we just created intake and it’s empty, it’s slightly cheaper to move+rename:
// but to keep logic simple & consistent, always merge then clean up.
_mergeFolderInto_(prospect, intake);




return intake.getUrl();
}








function moveApFolderToIntake_({ brand, customer, email, rootApptId, intakeFolderId }) {
if (!intakeFolderId) return null;




// Find client folder under brand’s clients root
const client = findOrCreateClientFolder_({ brand, customer, email });
const clientFolder = DriveApp.getFolderById(client.id);








// Locate /Prospects under client
const prospectsIt = clientFolder.getFoldersByName('Prospects');
if (!prospectsIt.hasNext()) return null;
const prospects = prospectsIt.next();
















// Find matching AP folder
let ap = null;
const kids = prospects.getFolders();
while (kids.hasNext()) {
  const f = kids.next();
  const name = f.getName();
  if (!/^AP[-_]/i.test(name)) continue;
  if (rootApptId && name.indexOf(rootApptId) >= 0) { ap = f; break; }
  if (!ap && /\(NO-SO-YET\)\s*$/i.test(name)) ap = f;
}
if (!ap) return null;








// IDs
const apId = ap.getId();
const intakeId = intakeFolderId;








// Use Advanced Drive API for Shared Drives
try {
  const apMeta   = Drive.Files.get(apId,    { supportsAllDrives: true });
  const intakeMd = Drive.Files.get(intakeId,{ supportsAllDrives: true });




  const apDriveId     = apMeta.driveId || '';
  const intakeDriveId = intakeMd.driveId || '';




  if (apDriveId && intakeDriveId && apDriveId !== intakeDriveId) {
    // Cross-Shared-Drive move of a FOLDER is not allowed: create a shortcut in Intake instead.
    Drive.Files.insert({
      title: apMeta.title + ' (shortcut)',
      mimeType: 'application/vnd.google-apps.shortcut',
      parents: [{ id: intakeId }],
      shortcutDetails: { targetId: apId }
    }, null, { supportsAllDrives: true });
    Logger.log('Created shortcut in 00-Intake because folders cannot be moved across Shared Drives.');
    // Keep the Intake folder link in the UI; user can open shortcut from there.
    return DriveApp.getFolderById(intakeId).getUrl();
  }




  // Same Shared Drive → reparent with add/remove parents
  const parentIds = (apMeta.parents || []).map(function(p){ return p.id; });
  Drive.Files.update(
    {},                 // metadata unchanged
    apId,
    null,               // no media
    {
      addParents: intakeId,
      removeParents: parentIds.join(','),
      supportsAllDrives: true
    }
  );




  // Optional rename
  if (rootApptId) {
    const desired = rootApptId.startsWith('AP-') ? rootApptId : `AP-${rootApptId}`;
    try { ap.setName(desired); } catch (e) { Logger.log('AP rename skipped: ' + e.message); }
  }
















  return DriveApp.getFolderById(apId).getUrl(); // return moved AP folder URL
} catch (e) {
  Logger.log('Shared-drive move failed, falling back to simple shortcut: ' + e.message);
  try {
    // Fallback: always ensure there’s at least a shortcut in Intake
    Drive.Files.insert({
      title: ap.getName() + ' (shortcut)',
      mimeType: 'application/vnd.google-apps.shortcut',
      parents: [{ id: intakeId }],
      shortcutDetails: { targetId: ap.getId() }
    }, null, { supportsAllDrives: true });
    return DriveApp.getFolderById(intakeId).getUrl();
  } catch (ee) {
    Logger.log('Fallback shortcut failed: ' + ee.message);
    return null;
  }
}
}








function copy3DTrackerToSO_(threeDFolderId, brand, so) {
const templateId = (PropertiesService.getScriptProperties().getProperty('3D_TRACKER_TEMPLATE_ID') || '').trim();
if (!templateId) throw new Error('3D_TRACKER_TEMPLATE_ID not set');
















const folder = DriveApp.getFolderById(threeDFolderId);
const tmpl   = DriveApp.getFileById(templateId);
const name   = `${brand}-SO${so} – 3D Tracker`;
















// Reuse if already present
const existing = folder.getFilesByName(name);
const file = existing.hasNext() ? existing.next() : tmpl.makeCopy(name, folder);
















// Ensure the template has a "Log" tab
const ss = SpreadsheetApp.openById(file.getId());
const sh = ss.getSheetByName('Log');
if (!sh) throw new Error('3D Tracker template missing "Log" tab');
















// Hand back id+url so callers can store it
return { id: ss.getId(), url: ss.getUrl() };
}








function getRowBasics_() {
const ss = SpreadsheetApp.getActive();
const sh = ss.getSheetByName('00_Master Appointments');
const r  = ss.getActiveRange();
if (!r || r.getSheet().getName() !== sh.getName() || r.getRow() < 2) {
  throw new Error('Select a data row on "00_Master Appointments" first.');
}
const row = r.getRow();
const H = {};
sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]
  .forEach((h,i)=>{ const k=String(h).trim(); if(k) H[k]=i+1; });




const get = n => H[n] ? String(sh.getRange(row, H[n]).getValue()||'') : '';




const brand      = (get('Brand') || '').toUpperCase();
const so         = (get('SO#') || '').replace(/^'/,'').trim();
const odooUrl    = get('Odoo SO URL');
const trackerUrl = get('3D Tracker');
const threeDUrl  = get('05-3D Folder');     // <— NEW
const masterLink = ss.getUrl() + '#gid=' + sh.getSheetId() + '&range=A' + row;




return { row, brand, so, odooUrl, trackerUrl, threeDUrl, masterLink, sheetId: sh.getSheetId(), headers: H };
}




// Helper: extract a Drive file id from a URL
function getIdFromUrl_(u){
  var m = String(u || '').match(/[-\w]{25,}/);
  return m ? m[0] : '';
}




/**
 * Returns prefill values for the 3D Revision dialog, using the newest row in
 * the Tracker's "Log" sheet (if present). Non-fatal: falls back to blanks.
 * Shape matches the Step-1 form keys used by dlg_3d_revision_v1.html.
 */
function getRevisionPrefill(){
  // Row context (brand, so, urls, headers, etc.)
  var basics = getRowBasics_(); // throws if no active row
  var out = {
    ok: true,
    brand: basics.brand || '',
    so: basics.so || '',
    trackerUrl: basics.trackerUrl || '',
    odooUrl: basics.odooUrl || '',
    prefill: { // defaults
      Mode: 'REVISION',
      AccentDiamondType: '',
      RingStyle: '',
      Metal: '',
      USSize: '',
      BandWidthMM: '',
      CenterDiamondType: '',
      Shape: '',
      DiamondDimension: '',
      DesignNotes: ''
    }
  };




  var trackerId = getIdFromUrl_(basics.trackerUrl);
  if (!trackerId) return out; // no tracker yet; return blanks safely




  var ssT = SpreadsheetApp.openById(trackerId);
  var sh  = ssT.getSheetByName('Log') || ssT.getSheets()[0];
  var last = sh.getLastRow();
  if (last < 2) return out; // header only or empty




  // Build header map (robust to renamed labels)
  var hdrs = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0] || [];
  var H = {}; hdrs.forEach(function(h,i){ var k=String(h||'').trim(); if(k) H[k]=i+1; });




  // Fallback indices by the order we append in append3DTrackerLog_:
  // 1 Timestamp, 2 User, 3 Action,
  // 4 Mode, 5 AccentDiamondType, 6 RingStyle, 7 Metal, 8 USSize, 9 BandWidthMM,
  // 10 CenterDiamondType, 11 Shape, 12 DiamondDimension, 13 DesignNotes, ... :contentReference[oaicite:0]{index=0}
  function gv(name, idx){
    if (H[name]) return sh.getRange(last, H[name]).getValue();
    if (idx)     return sh.getRange(last, idx).getValue();
    return '';
  }




  var pf = {
    Mode:              String(gv('Mode', 4) || '') || 'REVISION',
    AccentDiamondType: String(gv('Accent Type', 0) || gv('AccentDiamondType', 5) || ''),
    RingStyle:         String(gv('RingStyle', 6) || gv('Ring Style', 6) || ''),
    Metal:             String(gv('Metal', 7) || ''),
    USSize:            String(gv('USSize', 8) || gv('US Size', 8) || ''),
    BandWidthMM:       String(gv('BandWidthMM', 9) || gv('Band Width (mm)', 9) || ''),
    CenterDiamondType: String(gv('Center Type', 0) || gv('CenterDiamondType', 10) || gv('Center Diamond Type', 10) || ''),
    Shape:             String(gv('Shape', 11) || ''),
    DiamondDimension:  String(gv('DiamondDimension', 12) || gv('Diamond Dimension', 12) || ''),
    DesignNotes:       String(gv('Design Notes', 13) || '')
  };




  // Normalize whitespace
  Object.keys(pf).forEach(function(k){ pf[k] = (pf[k] == null ? '' : String(pf[k])).trim(); });




  out.prefill = pf;
  return out;
}












function append3DTrackerLog_({ trackerId, action, form, brand, so, odooUrl, masterLink, shortTag }) {
if (!trackerId) throw new Error('No trackerId');
const ss = SpreadsheetApp.openById(trackerId);
const sh = ss.getSheetByName('Log') || ss.insertSheet('Log');
















const user = Session.getActiveUser?.().getEmail?.() || Session.getEffectiveUser()?.getEmail?.() || '';
















// `form` is the Step 1 payload (same keys you already use in Start 3D)
const row = [
  new Date(),                    // Timestamp
  user,                          // User
  action || 'Start 3D Linked',   // Action
  form?.Mode || '',              // Mode
  form?.AccentDiamondType || '',
  form?.RingStyle || '',
  form?.Metal || '',
  form?.USSize || '',
  form?.BandWidthMM || '',
  form?.CenterDiamondType || '',
  form?.Shape || '',
  form?.DiamondDimension || '',
  (form?.DesignNotes || ''),     // Design Notes
  shortTag || '',
  so || '',
  brand || '',
  odooUrl || '',
  masterLink || '',
  ''                             // Version (optional: fill later)
];
sh.appendRow(row);
}


// Exported: called by google.script.run from HTML
// ==== REPLACE save3DRevision(payload) WRAPPER ====
function save3DRevision(payload) {
try {
  Logger.log('save3DRevision keys: ' + Object.keys(payload || {}).join(','));
  const basics = getRowBasics_();                // includes headers + handy links
  // If you prefer to use the richer preview object, you could call getActiveMasterPreview() instead.
















  // Your existing tracker-write logic may already be running elsewhere.
  // The core helper below ONLY handles Master Odoo URL + response shape.
  return save3DRevisionCore_(payload, basics);
} catch (e) {
  console.error('save3DRevision error:', e && e.stack || e);
  throw e; // let the dialog failure handler show the message
}
}
































// 2) Rename your existing implementation to this (everything that was
//    below the recursive return stays exactly the same body):
// ==== REPLACE/ADD save3DRevisionCore_(payload, basics) ====
/**
* Appends a 3D revision row to the Tracker "Log" sheet and returns a JSON-safe summary.
* Expects: payload = { form: {...}, odooUrl?: string }, basics = row context (brand, so, links...)
*/
function save3DRevisionCore_(payload, basics) {
 try {
   // ---------- 0) Normalize inputs ----------
   basics = basics || {};
   payload = payload || {};
   var form = payload.form || {};
   if (!form.Mode) form.Mode = 'REVISION';


   var tz = (Session.getScriptTimeZone && Session.getScriptTimeZone()) || 'America/Los_Angeles';
   var toStr = function (v) {
     if (v === null || v === undefined) return '';
     if (v instanceof Date) return Utilities.formatDate(v, tz, 'MMM d, yyyy h:mm a z');
     return String(v);
   };
   var actor = (function () {
     try { return Session.getActiveUser().getEmail() || ''; } catch (_) { return ''; }
   })();

   // read helpers
   var getIdFromUrl = function (u) {
     if (!u) return '';
     var m = String(u).match(/\/d\/([a-zA-Z0-9-_]+)/);
     return m ? m[1] : '';
   };

   // ---------- 1) Resolve SO / brand / links ----------
   var soVal = form.SO || form['SO#'] || basics.so || basics.SO || basics['SO#'] || '';
   var brand = basics.brand || basics.Brand || form.Brand || '';
   var odooUrl = (payload.odooUrl || basics.odooUrl || basics.OdooUrl || '').toString();


   // Try to provide a direct link back to the master row if basics gives us enough info.
   var masterLink = basics.masterLink || basics.masterUrl || '';
   if (!masterLink) {
     try {
       var ss = SpreadsheetApp.getActive();
       var sh = ss.getActiveSheet();
       if (sh) {
         var row = sh.getActiveRange().getRow();
         var sheetId = sh.getSheetId();
         masterLink = ss.getUrl() + '#gid=' + sheetId + '&range=' + row + ':' + row;
       }
     } catch (_) {}
   }

   // ---------- 2) Resolve / ensure Tracker ----------
   var trackerId = basics.trackerId || getIdFromUrl(basics.trackerUrl);
   if (!trackerId && typeof ensure3DTrackerForSO_ === 'function') {
     trackerId = ensure3DTrackerForSO_(basics);
   }
   if (!trackerId) {
     return { ok: false, error: 'No 3D Revision Tracker found for this row (missing trackerId).' };
   }
   var ssT = SpreadsheetApp.openById(trackerId);
   var shT = ssT.getSheetByName('Log') || ssT.insertSheet('Log');

   // ---------- 3) Ensure header has all needed columns ----------
   // We support your requested columns + the 3D form fields.
   var desiredHeaders = [
     'Timestamp', 'User', 'Action', 'Revision #', 'Brand',
     'SO#', 'Odoo SO URL', 'Master Link',
     'Mode', 'Accent Type', 'Ring Style', 'Metal',
     'US Size', 'Band Width (mm)',
     'Center Type', 'Shape', 'Diamond Dimension',
     'Design Notes'
   ];

   var lastRow = shT.getLastRow();
   var lastCol = Math.max(1, shT.getLastColumn());
   var header = [];
   if (lastRow === 0) {
     header = desiredHeaders.slice();
     shT.getRange(1, 1, 1, header.length).setValues([header]);
     lastRow = 1;
     lastCol = header.length;
   } else {
     header = shT.getRange(1, 1, 1, lastCol).getValues()[0] || [];
     header = header.map(function (h) { return String(h || '').trim(); });








     // add any missing desired header(s) to the right
     var missing = desiredHeaders.filter(function (h) { return header.indexOf(h) === -1; });
     if (missing.length) {
       shT.getRange(1, header.length + 1, 1, missing.length).setValues([missing]);
       header = header.concat(missing);
     }
   }
   var pos = {}; header.forEach(function (h, i) { pos[h] = i + 1; });


  // ---------- 4.5) Compute Revision # for this SO ----------
   var revNo = 1;
   try {
     if (pos['SO#'] && pos['Action'] && lastRow >= 2 && soVal) {
       var rng = shT.getRange(2, 1, lastRow - 1, header.length).getValues();
       var iSO = pos['SO#'] - 1, iAC = pos['Action'] - 1;
       var priorCount = 0;
       for (var i = 0; i < rng.length; i++) {
         var soPrev = String(rng[i][iSO] || '').trim();
         var actPrev = String(rng[i][iAC] || '').trim().toLowerCase();
         if (soPrev === soVal && actPrev === '3d revision requested') priorCount++;
       }
       revNo = priorCount + 1;
     }
   } catch (_) {}


   // ---------- 4) Capture previous row object (for success table) ----------
   function rowObj(rowIdx) {
     if (!rowIdx || rowIdx < 2) return {};
     var vals = shT.getRange(rowIdx, 1, 1, header.length).getValues()[0] || [];
     var o = {};
     for (var i = 0; i < header.length; i++) o[header[i]] = toStr(vals[i]);
     return o;
   }
   var prevObj = rowObj(lastRow);

   // ---------- 5) Read values from form robustly (synonyms allowed) ----------
   // Helper to pick the first non-empty among a list of keys (supports labels and compacted keys)
   function pick(keys) {
     for (var i = 0; i < keys.length; i++) {
       var k = keys[i];
       if (form.hasOwnProperty(k) && form[k] != null && form[k] !== '') return form[k];
     }
     // try compact/normalized lookup (ignore case, spaces, punctuation)
     var norm = {};
     Object.keys(form).forEach(function (k) {
       var nk = String(k).toLowerCase().replace(/[^a-z0-9]+/g, '');
       norm[nk] = form[k];
     });
     for (var j = 0; j < keys.length; j++) {
       var nk2 = String(keys[j]).toLowerCase().replace(/[^a-z0-9]+/g, '');
       if (norm.hasOwnProperty(nk2) && norm[nk2] != null && norm[nk2] !== '') return norm[nk2];
     }
     return '';
   }


   var modeVal      = pick(['Mode']);
   var accentVal    = pick(['AccentDiamondType','Accent','Accent Type']);
   var ringStyleVal = pick(['RingStyle','Ring Style']);
   var metalVal     = pick(['Metal']);
   var usSizeVal    = pick(['US Size','USSize','Ring Size','Size','Size US']);
   var bandVal      = pick(['Band Width (mm)','BandWidthMM','Band Width','BandWidth']);
   var centerType   = pick(['CenterDiamondType','Center Diamond Type']);
   var shapeVal     = pick(['Shape']);
   var dimVal       = pick(['DiamondDimension','Dimension','Diamond Dimension']);
   var notesVal     = pick(['DesignNotes','Design Notes']);


   // ---------- 6) Build the row to write ----------
   var now = new Date();
   var rowObjToWrite = {
     'Timestamp': now,
     'User': actor,
     'Action': '3D Revision Requested',
     'Revision #': revNo,
     'Brand': brand,
     'SO#': soVal,
     'Odoo SO URL': odooUrl,
     'Master Link': masterLink,
     'Mode': modeVal,
     'Accent Type': accentVal,
     'Ring Style': ringStyleVal,
     'Metal': metalVal,
     'US Size': usSizeVal,
     'Band Width (mm)': bandVal,
     'Center Type': centerType,
     'Shape': shapeVal,
     'Diamond Dimension': dimVal,
     'Design Notes': notesVal
   };


   // Create the row array aligned to header
   var writeRow = new Array(header.length);
   for (var c = 0; c < header.length; c++) {
     var key = header[c];
     writeRow[c] = rowObjToWrite.hasOwnProperty(key) ? rowObjToWrite[key] : '';
   }


   // ---------- 7) Append + format timestamp as date & time ----------
   var writeR = lastRow + 1;
   shT.getRange(writeR, 1, 1, header.length).setValues([writeRow]);
   // Set a clear datetime format on the Timestamp cell (column 1)
   shT.getRange(writeR, 1).setNumberFormat('mmm d, yyyy h:mm am/pm');

   // ---------- 8) Prepare JSON-safe summary for success panel ----------
   var latest = rowObj(writeR); // this is already stringified via rowObj/toStr
   var nextHuman = {
    'Action':             latest['Action'] || '3D Revision Requested',
    'Revision #':         latest['Revision #'] || String(revNo),
    'SO#':                latest['SO#'] || '',
    'Mode':               latest['Mode'] || '',
    'Brand':              latest['Brand'] || '',
    'US Size':            latest['US Size'] || '',
    'Band Width (mm)':    latest['Band Width (mm)'] || '',

    // ✅ New schema first, old fallback second
    'Accent Type':        latest['Accent Type'] || latest['AccentDiamondType'] || '',
    'Ring Style':         latest['Ring Style'] || latest['RingStyle'] || '',
    'Metal':              latest['Metal'] || '',
    'Center Type':        latest['Center Type'] || latest['CenterDiamondType'] || '',
    'Shape':              latest['Shape'] || '',
    'Diamond Dimension':  latest['Diamond Dimension'] || latest['DiamondDimension'] || '',
    'Design Notes':       latest['Design Notes'] || latest['DesignNotes'] || '',

    'Odoo SO URL':        latest['Odoo SO URL'] || '',
    'Master Link':        latest['Master Link'] || '',
    'Timestamp':          latest['Timestamp'] || '',
    'User':               latest['User'] || ''
  };




  // --- NEW: Write Odoo “Copy & Paste” into 100_ + set status ---
  (function updateMasterDesignAndStatus_(){
    try {
      var ss = SpreadsheetApp.getActive();
      var sh = ss.getSheetByName('00_Master Appointments');
      if (!sh || !basics || !basics.row) return;

      var row = basics.row;
      var H   = basics.headers || {};

      // Ensure columns exist (adds if missing)
      ['Design Request','Custom Order Status'].forEach(function(name){
        if (!H[name]) {
          sh.getRange(1, sh.getLastColumn()+1).setValue(name);
          H[name] = sh.getLastColumn();
        }
      });

      // 1) Overwrite Design Request with the dialog’s formatted paste (if provided)
      var paste = (payload && payload.odooPaste) ? String(payload.odooPaste) : '';
      if (paste) sh.getRange(row, H['Design Request']).setValue(paste);

      // 2) Set Custom Order Status
      sh.getRange(row, H['Custom Order Status']).setValue('3D Revision Requested');
    } catch (e) {
      Logger.log('Master update (Design/Status) skipped: ' + e.message);
    }
  })();

  // [Removed] No 301/302 updates from Assign SO / 3D Revision path

   return {
     ok: true,
     trackerUrl: 'https://docs.google.com/spreadsheets/d/' + trackerId,
     odooUrl: odooUrl,
     submittedAt: toStr(now),
     prev: prevObj,   // previous log row (stringified)
     next: nextHuman,  // current row (human-friendly labels)
     soUpdate: soUpdate
   };

 } catch (e) {
   return { ok: false, error: (e && e.message) ? e.message : String(e) };
 }
}




/** === 301/302 updater (safe, non-throwing) === */
function _normalizeSO_(raw){
  // Accept "SO12.3456" or "12.3456" → "12.3456"
  var s = String(raw || '').trim();
  s = s.replace(/^SO/i, '').trim();
  return s;
}


function _logAutomationSafe_(msg){
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName('10_Automation_Log') || ss.insertSheet('10_Automation_Log');
    sh.appendRow([new Date(), '3D Revision', String(msg || '')]);
  } catch (_) { /* swallow */ }
}




/** Read “Dropdown” tab values by header name (row 1) */
function report_getDropdownLists_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Dropdown');
  if (!sh) throw new Error('Sheet not found: "Dropdown"');




  const values = sh.getDataRange().getValues();
  if (!values.length) return {};




  const headers = values[0].map(String);
  const rows = values.slice(1);




  const dedup = (arr) => {
    const seen = new Set(), out = [];
    for (const v of arr) {
      const s = (v == null ? '' : String(v)).trim();
      if (!s || seen.has(s)) continue;
      seen.add(s); out.push(s);
    }
    return out;
  };
  const col = (name) => {
    const i = headers.indexOf(name);
    if (i === -1) return [];
    return dedup(rows.map(r => r[i]));
  };

  const assignedReps  = col('Assigned Rep');
  const assistedReps0 = col('Assisted Rep');
  const assistedReps  = assistedReps0.length ? assistedReps0 : assignedReps;

  return {
    salesStage:        col('Sales Stage'),
    conversionStatus:  col('Conversion Status'),
    customOrderStatus: col('Custom Order Status'),
    centerStoneStatus: col('Center Stone Order Status'),
    assignedReps,
    assistedReps,
  };
}




/***** ===== Report v1 — Data + Filtering + Export ===== *****/
// Read master sheet
// Read master sheet with safe 60s cache (skips cache if too large)
function report_getMasterData_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Sheet not found: "00_Master Appointments"');

  const key = 'report:master:v1:' + sh.getSheetId() + ':' + sh.getLastRow() + 'x' + sh.getLastColumn();
  const cached = report_cacheGet_(key);
  if (cached && cached.headers && cached.rows) return cached;

  const values  = sh.getDataRange().getValues();
  if (values.length < 2) return { headers: [], rows: [] };

  const headers = values[0].map(String);
  const rows    = values.slice(1);

  const payload = { headers, rows };
  report_cachePut_(key, payload, 60); // silently skipped if too large
  return payload;
}


function report_headerIndex_(headers) {
  const H = {};
  headers.forEach((h, i) => H[String(h)] = i);
  return H;
}


// Shape result with Client Status Report URL (after Next Steps); remove Odoo/3D
function report_shapeResult_(headers, rows) {
  const cols = [
    'APPT_ID','Customer Name','Assigned Rep','Brand','SO#',
    'Sales Stage','Conversion Status','Custom Order Status','Center Stone Order Status',
    'Next Steps','Client Status Report URL'
  ];
  const idx = cols.map(c => headers.indexOf(c));
  const shaped = rows.map(r => idx.map(i => (i >= 0 ? r[i] : ''))); // safe if a column is missing
  return { headers: cols, rows: shaped, total: rows.length, previewLimit: 1000 };
}


// OR within each list, OR across lists; empty selections = show all
function report_runStatus(payload) {
  const data = report_getMasterData_();
  const H = report_headerIndex_(data.headers);


  const sets = {
    salesStage:        new Set(payload?.salesStage || []),
    conversionStatus:  new Set(payload?.conversionStatus || []),
    customOrderStatus: new Set(payload?.customOrderStatus || []),
    centerStoneStatus: new Set(payload?.centerStoneStatus || [])
  };


  const out = [];
  for (const r of data.rows) {
    const vStage  = String(r[H['Sales Stage']] || '').trim();
    const vConv   = String(r[H['Conversion Status']] || '').trim();
    const vCust   = String(r[H['Custom Order Status']] || '').trim();
    const vCenter = String(r[H['Center Stone Order Status']] || '').trim();


    // OR within each category; if no selections in a category, treat as TRUE (non-restrictive)
    const mStage  = sets.salesStage.size        ? sets.salesStage.has(vStage)      : true;
    const mConv   = sets.conversionStatus.size  ? sets.conversionStatus.has(vConv) : true;
    const mCust   = sets.customOrderStatus.size ? sets.customOrderStatus.has(vCust): true;
    const mCenter = sets.centerStoneStatus.size ? sets.centerStoneStatus.has(vCenter) : true;


    // AND across categories
    const match = (mStage && mConv && mCust && mCenter);
    if (match) out.push(r);
  }


  const shaped  = report_shapeResult_(data.headers, out);
  const summary = report_buildSummary_(data.headers, out);
  const warn    = out.length > 10000 ? 'Large result (>10,000 rows). Consider narrowing filters.' : '';


  if (payload && payload._mode === 'preview') {
    shaped.rows = shaped.rows.slice(0, Math.min(shaped.previewLimit || 1000, shaped.rows.length));
  }
  return Object.assign({}, shaped, { summary, warn });
}




// OR within each list, OR across Assigned/Assisted; empty selections = show all
function report_runReps(payload) {
  const data = report_getMasterData_();
  const H = report_headerIndex_(data.headers);




  const assignedSet = new Set((payload?.assigned || []).map(String));
  const assistedSet = new Set((payload?.assisted || []).map(String));
  const anySelected = (assignedSet.size + assistedSet.size) > 0;




  const out = [];
  for (const r of data.rows) {
    const assignedTokens = report_tokenizeMulti_(r[H['Assigned Rep']]);
    const assistedTokens = report_tokenizeMulti_(r[H['Assisted Rep']]);




    const assignedMatch = assignedSet.size ? assignedTokens.some(t => assignedSet.has(t)) : false;
    const assistedMatch = assistedSet.size ? assistedTokens.some(t => assistedSet.has(t)) : false;




    const match = anySelected ? (assignedMatch || assistedMatch) : true;
    if (match) out.push(r);
  }




  const shaped = report_shapeResult_(data.headers, out);     // FULL shape
  const summary = report_buildSummary_(data.headers, out);
  const warn = out.length > 10000 ? 'Large result (>10,000 rows). Consider narrowing filters.' : '';




  if (payload && payload._mode === 'preview') {
    shaped.rows = shaped.rows.slice(0, Math.min(shaped.previewLimit || 1000, shaped.rows.length));
  }
  return Object.assign({}, shaped, { summary, warn });
}




// Short label for export sheet name
function report_shortCriteria_(type, payload) {
  if (type === 'status') {
    const parts = [];
    for (const k of ['salesStage','conversionStatus','customOrderStatus','centerStoneStatus']) {
      if (Array.isArray(payload?.[k]) && payload[k].length) parts.push(payload[k].slice(0, 2).join('|'));
    }
    return parts.join(' • ') || 'All';
  }
  if (type === 'rep') {
    const A = (payload?.assigned || []).slice(0,2).join('|');
    const S = (payload?.assisted || []).slice(0,2).join('|');
    const parts = [];
    if (A) parts.push('A:' + A);
    if (S) parts.push('S:' + S);
    return parts.join(' • ') || 'All';
  }
  return 'All';
}


// Create new tab and write full results
function report_export(type, payload) {
  const ss = SpreadsheetApp.getActive();
  const res = (type === 'status') ? report_runStatus(payload)
            : (type === 'rep')    ? report_runReps(payload)
            : (function(){ throw new Error('Unknown report type: ' + type); })();

  const name = report_buildSheetName_(type, payload);   // safe <=100 chars
  const sh = ss.insertSheet(name);
  if (!sh) throw new Error('Failed to create export sheet');

  if (res.headers.length) sh.getRange(1, 1, 1, res.headers.length).setValues([res.headers]);
  if (res.rows.length)    sh.getRange(2, 1, res.rows.length, res.headers.length).setValues(res.rows);

  report_applyExportFormatting_(sh, res);               // from 5D
  const url = ss.getUrl() + '#gid=' + sh.getSheetId();

  // NEW: write summary on a second tab
  report_writeSummarySheet_(ss, name, res.summary);

  sh.activate();
  return { sheetName: name, rows: res.rows.length, columns: res.headers.length, url };
}


/** Split dropdown chips (multi-select) into clean tokens. Handles commas, newlines, bullets, etc. */
function report_tokenizeMulti_(v) {
  if (v == null) return [];
  return String(v)
    .replace(/\u00A0/g, ' ')          // NBSP → space
    .split(/[\n\r,;|•]+/g)            // common chip separators
    .map(s => s.trim())
    .filter(Boolean);
}

/** Build a sheet name <= 100 chars. Falls back to a compact name with a short hash. */
function report_buildSheetName_(type, payload) {
  const title = (type === 'status') ? 'By Status' : 'By Rep';
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH.mm');
  const crit = report_shortCriteria_(type, payload) || 'All';

  // First try: full name
  let name = `Report — ${title} — ${crit} — ${now}`;
  if (name.length <= 100) return name;

  // Compact fallback with short payload hash
  const hash = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(payload))
  ).slice(0, 4);
  const compact = `Report — ${title} — ${now} — ${hash}`;
  if (compact.length <= 100) return compact;

  // Final guard (extremely unlikely): hard trim to 100
  return compact.slice(0, 100);
}


/** Apply pretty formatting to a freshly-exported report sheet. */
// Clamp rows to ~2 lines by flattening "Next Steps" (export-only), clipping, and forcing height
function report_applyExportFormatting_(sh, res) {
  var lastRow = sh.getLastRow();
  var lastCol = res.headers.length;
  if (lastRow < 1 || lastCol < 1) return;


  // 0) Header + freeze
  sh.setFrozenRows(1);
  var header = sh.getRange(1, 1, 1, lastCol);
  header.setFontWeight('bold').setBackground('#f5f5f5');


  // 1) Column sizing
  sh.autoResizeColumns(1, lastCol);
  var nextIdx = res.headers.indexOf('Next Steps'); // 0-based index
  if (nextIdx !== -1) sh.setColumnWidth(nextIdx + 1, 320);


  // 2) EXPORT-ONLY NORMALIZATION: remove hard line breaks in "Next Steps"
  if (lastRow > 1 && nextIdx !== -1) {
    var rng = sh.getRange(2, nextIdx + 1, lastRow - 1, 1);
    var vals = rng.getValues(); // [[val],[val],...]
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i][0];
      if (v != null && v !== '') {
        // Replace CR/LF runs with a separator so it stays readable on one line
        vals[i][0] = String(v).replace(/\r\n|\r|\n/g, ' • ');
      }
    }
    rng.setValues(vals);
  }


  // 3) Clip ALL data cells so nothing can push row height
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, lastCol)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
      .setVerticalAlignment('top');
  }


  // 4) Commit before height ops
  SpreadsheetApp.flush();


  // 5) Force fixed height for DATA rows (≈ two lines)
  if (lastRow > 1) {
    var DATA_ROW_HEIGHT = 48; // change to 46 if you prefer
    var start = 2, remain = lastRow - 1, BATCH = 500;
    while (remain > 0) {
      var n = Math.min(BATCH, remain);
      sh.setRowHeights(start, n, DATA_ROW_HEIGHT);
      start += n; remain -= n;
    }
    SpreadsheetApp.flush();
  }


  // 6) Apply filter + banding last
  var full = sh.getRange(1, 1, lastRow, lastCol);
  if (sh.getFilter()) sh.getFilter().remove();
  full.createFilter();
  full.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);


  SpreadsheetApp.flush();
}




/** Count frequencies in an array of strings. Returns {value: count}. */
function report_countMap_(arr) {
  const m = {};
  for (const v0 of arr) {
    const v = (v0 == null ? '' : String(v0)).trim();
    const key = v || '(blank)';
    m[key] = (m[key] || 0) + 1;
  }
  return m;
}




/** Split multi-select dropdown cell to tokens (we added earlier). */
function report_tokenizeMulti_(v) {
  if (v == null) return [];
  return String(v)
    .replace(/\u00A0/g, ' ')        // NBSP → space
    .split(/[\n\r,;|•]+/g)          // common chip separators
    .map(s => s.trim())
    .filter(Boolean);
}




/** Build summary counts from filtered rows. */
function report_buildSummary_(headers, rows) {
  const H = {};
  headers.forEach((h, i) => H[h] = i);




  const salesStages   = [];
  const conversions   = [];
  const customs       = [];
  const centers       = [];
  const unifiedReps   = [];  // union of tokens in Assigned + Assisted per row (no dup per row)




  for (const r of rows) {
    salesStages.push(r[H['Sales Stage']]);
    conversions.push(r[H['Conversion Status']]);
    customs.push(r[H['Custom Order Status']]);
    centers.push(r[H['Center Stone Order Status']]);




    const assigned = report_tokenizeMulti_(r[H['Assigned Rep']]);
    const assisted = report_tokenizeMulti_(r[H['Assisted Rep']]);
    const bag = new Set([...assigned, ...assisted]);
    if (bag.size === 0) bag.add('(blank)');
    for (const rep of bag) unifiedReps.push(rep);
  }




  return {
    totalRows: rows.length,
    groups: {
      rep:                   report_countMap_(unifiedReps),                // Rep (Assigned or Assisted)
      salesStage:            report_countMap_(salesStages),
      conversionStatus:      report_countMap_(conversions),
      customOrderStatus:     report_countMap_(customs),
      centerStoneStatus:     report_countMap_(centers),
    }
  };
}




/** Build a safe name <=100 chars for the second (Summary) sheet. */
function report_buildSummarySheetNameFromBase_(baseName) {
  const name1 = baseName + ' — Summary';
  if (name1.length <= 100) return name1;
  const name2 = 'Summary — ' + baseName;
  if (name2.length <= 100) return name2;
  return name2.slice(0, 100);
}




/** Write the summary object to a new sheet as stacked sections. */
function report_writeSummarySheet_(ss, baseName, summary) {
  if (!summary || !summary.groups) return null;
  const name = report_buildSummarySheetNameFromBase_(baseName);
  const sh = ss.insertSheet(name);
  if (!sh) throw new Error('Failed to create summary sheet');




  // Build rows: each group as its own section
  const sections = [
    ['Rep (Assigned or Assisted)',      'rep'],
    ['Sales Stage',                     'salesStage'],
    ['Conversion Status',               'conversionStatus'],
    ['Custom Order Status',             'customOrderStatus'],
    ['Center Stone Order Status',       'centerStoneStatus'],
  ];




  let row = 1;
  for (const [label, key] of sections) {
    const map = summary.groups[key] || {};
    const entries = Object.keys(map).map(k => [k, map[k]]);
    // Sort desc by count, then alpha
    entries.sort((a,b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])));




    // Section header
    sh.getRange(row, 1).setValue(label).setFontWeight('bold');
    row += 1;
    // Table header
    sh.getRange(row, 1, 1, 2).setValues([['Value','Count']]).setFontWeight('bold').setBackground('#f5f5f5');
    row += 1;
    // Data
    if (entries.length) {
      sh.getRange(row, 1, entries.length, 2).setValues(entries);
      row += entries.length;
    }
    // Blank spacer row
    row += 1;
  }


  // Auto-resize & banding
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.autoResizeColumns(1, 2);
    sh.getRange(1, 1, lastRow, 2).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  }
  return { sheetName: name };
}


/** Lightweight cache accessors (100 KB per entry limit). */
function report_cache_()           { return CacheService.getUserCache(); }
function report_cacheGet_(key)     { try { const s = report_cache_().get(key); return s ? JSON.parse(s) : null; } catch(e) { return null; } }
function report_cachePut_(key, obj, ttlSec) {
  try {
    const json = JSON.stringify(obj);
    if (json.length <= 90000) report_cache_().put(key, json, ttlSec || 60); // stay below 100KB hard limit
  } catch (e) { /* noop */ }
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



