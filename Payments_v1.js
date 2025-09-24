/*** 02 - Payments_v1.gs — FINAL v8.3 (2025‑09-04)
     Full replacement to pair with “dlg_record_payment_v1.html — FINAL v3.4 (2025‑08‑29)”.
     Implements:
       • rp_init() robust prefill (SO- or APPT‑anchored) with fallback to 100_ for OT/PTD.
       • Financial Summary for RECEIPTS uses Lines Subtotal − Payment (per latest spec).
       • Invoices leave Remaining Balance unchanged by Requested Amount.
       • Doc placeholder filling fixed (reliable {{...}} replacement across body/tables).
       • 3D reader resilient to header variants.
       • Files: APPT → Client/04-Deposit; SO → Orders/SO#/04-Deposit. PaymentsFolderURL written when blank.
       • AR monthly shortcut: brand-based top folder (VVS → 20_AR, HPUSA → 21_AR).
       • Table renderer never double‑bullets “✧”.


     References: latest server/client snapshots and template. 
**************************************************************************/


/*** === CONFIG / CONSTANTS === ***/
// --- DEBUG SWITCH (toggle via Script Properties: RP_DEBUG = "true" / "false") ---
var RP_DEBUG = (function(){ 
  try { return PropertiesService.getScriptProperties().getProperty('RP_DEBUG') === 'true'; } 
  catch (_){ return false; } 
})();
function RP_LOG() { if (RP_DEBUG) Logger.log.apply(Logger, arguments); }
// --- Back-compat wrappers so older calls still work in v8.1 ---
function rp_log() { try { RP_LOG.apply(null, arguments); } catch(_) {} }
// Simple stopwatch: const done = rp_time('label'); ...; done();
function rp_time(label) {
  var t = Date.now();
  return function(){ RP_LOG('%s %dms', label || '⏱', Date.now() - t); };
}


const RP_MASTER_SHEET = '00_Master Appointments';


const RP_KEY_ALIASES = {
  LEDGER_FILE_ID: ['PAYMENTS_400_FILE_ID','LEDGER_FILE_ID','PAYMENTS_LEDGER_FILE_ID','PAYMENTS_FILE_ID','CFG_PAYMENTS_LEDGER_FILE_ID'],
  LEDGER_SHEET_NAME: ['PAYMENTS_SHEET_NAME','CFG_PAYMENTS_SHEET_NAME'],


  ORDERS_HPUSA_FILE_ID: ['HPUSA_301_FILE_ID','HPUSA_ORDERS_FILE_ID','CFG_HPUSA_ORDERS_FILE_ID'],
  ORDERS_VVS_FILE_ID:   ['VVS_302_FILE_ID','VVS_ORDERS_FILE_ID','CFG_VVS_ORDERS_FILE_ID'],
  ORDERS_TAB_COMMON:    ['301/302_TAB_NAME','ORDERS_TAB_NAME','CFG_ORDERS_TAB_NAME'],
  ORDERS_HPUSA_TAB:     ['HPUSA_301_TAB_NAME','CFG_HPUSA_ORDERS_TAB_NAME'],
  ORDERS_VVS_TAB:       ['VVS_302_TAB_NAME','CFG_VVS_ORDERS_TAB_NAME'],


  TEMPLATE_DEPOSIT_INVOICE_HPUSA: ['TEMPLATE_DEPOSIT_INVOICE_HPUSA','HPUSA_DI_TEMPLATE_ID','CFG_TEMPLATE_DEPOSIT_INVOICE_HPUSA'],
  TEMPLATE_DEPOSIT_RECEIPT_HPUSA: ['TEMPLATE_DEPOSIT_RECEIPT_HPUSA','HPUSA_DR_TEMPLATE_ID','CFG_TEMPLATE_DEPOSIT_RECEIPT_HPUSA'],
  TEMPLATE_SALES_INVOICE_HPUSA:   ['TEMPLATE_SALES_INVOICE_HPUSA','HPUSA_SI_TEMPLATE_ID','CFG_TEMPLATE_SALES_INVOICE_HPUSA'],
  TEMPLATE_SALES_RECEIPT_HPUSA:   ['TEMPLATE_SALES_RECEIPT_HPUSA','HPUSA_SR_TEMPLATE_ID','CFG_TEMPLATE_SALES_RECEIPT_HPUSA'],

  TEMPLATE_DEPOSIT_INVOICE_VVS: ['TEMPLATE_DEPOSIT_INVOICE_VVS','VVS_DI_TEMPLATE_ID','CFG_TEMPLATE_DEPOSIT_INVOICE_VVS'],
  TEMPLATE_DEPOSIT_RECEIPT_VVS: ['TEMPLATE_DEPOSIT_RECEIPT_VVS','VVS_DR_TEMPLATE_ID','CFG_TEMPLATE_DEPOSIT_RECEIPT_VVS'],
  TEMPLATE_SALES_INVOICE_VVS:   ['TEMPLATE_SALES_INVOICE_VS','TEMPLATE_SALES_INVOICE_VVS','VVS_SI_TEMPLATE_ID','CFG_TEMPLATE_SALES_INVOICE_VVS'],
  TEMPLATE_SALES_RECEIPT_VVS:   ['TEMPLATE_SALES_RECEIPT_VVS','VVS_SR_TEMPLATE_ID','CFG_TEMPLATE_SALES_RECEIPT_VVS'],

  AR_HPUSA_ROOT_ID: ['AR_HP_RootID','AR_HPUSA_ROOT_ID','CFG_AR_HPUSA_ROOT_ID'],
  AR_VVS_ROOT_ID:   ['AR_VVS_RootID','AR_VS_ROOT_ID','AR_VS_ROOT','AR_VVS_ROOT_ID','CFG_AR_VVS_ROOT_ID'],

  FEES_JSON:    ['PAYMENT_FEES_JSON','CFG_PAYMENT_FEES_JSON'],
  FEES_TAB_NAME:['PAYMENTS_FEES_TAB_NAME','CFG_PAYMENTS_FEES_TAB_NAME'],

  SO_RECEIPT_MASTER_AMOUNT: ['SO_RECEIPT_MASTER_AMOUNT','CFG_SO_RECEIPT_MASTER_AMOUNT'],

  HPUSA_SO_ROOT_FOLDER_ID: ['HPUSA_SO_ROOT_FOLDER_ID','CFG_HPUSA_SO_ROOT_FOLDER_ID'],
  VVS_SO_ROOT_FOLDER_ID:   ['VVS_SO_ROOT_FOLDER_ID','CFG_VVS_SO_ROOT_FOLDER_ID']
};

/** ===== Doc & Tax constants (shared with Payment Summary) ===== */
var RP_DOC_STATUS = { DRAFT:'DRAFT', ISSUED:'ISSUED', REPLACED:'REPLACED', VOID:'VOID' };
var RP_DOC_ROLE   = { DEPOSIT:'DEPOSIT', PROGRESS:'PROGRESS', FINAL:'FINAL', CREDIT:'CREDIT', PAYMENT_RECEIPT:'PAYMENT_RECEIPT', SALES_RECEIPT:'SALES_RECEIPT' };

function rp_prop_(k, d){  // reads Script Properties safely
  try { return PropertiesService.getScriptProperties().getProperty(k) || d || ''; }
  catch(_){ return d || ''; }
}

/** Templates (set these in Script Properties) */
var TEMPLATE_CM_VVS_ID = rp_prop_('TEMPLATE_CM_VVS_ID','');  // Credit Memo (VVS)
var TEMPLATE_CM_HP_ID  = rp_prop_('TEMPLATE_CM_HP_ID','');   // Credit Memo (HPUSA)

/** Tax defaults (from your spec) */
var TAX_RATE_DEFAULT = Number(rp_prop_('TAX_RATE_DEFAULT','0.09375'));  // 9.375%
var TAX_MODE_DEFAULT = rp_prop_('TAX_MODE_DEFAULT','TAX_INCLUDED');     // NO_TAX | TAX_INCLUDED | ADD_TAX
var TAX_ROUNDING     = rp_prop_('TAX_ROUNDING','INVOICE');              // INVOICE | LINE
var TAX_DECIMALS     = Number(rp_prop_('TAX_DECIMALS','2'));            // 2
var ALLOW_SALES_RECEIPT_PARTIAL = (rp_prop_('ALLOW_SALES_RECEIPT_PARTIAL','false') === 'true');


const RP_PROPS = PropertiesService.getScriptProperties();
const RP_TZ    = Session.getScriptTimeZone() || 'America/Los_Angeles';

function rp_propOneOf_(aliases, opt = {}) {
  const props = RP_PROPS;
  for (const k of (aliases || [])) {
    const v = props.getProperty(k);
    if (v != null && String(v).trim() !== '') return { key:k, value:v };
  }
  if (opt.required) throw new Error(`[Config] Missing property for ${opt.label || 'unnamed'}. Tried: ${aliases.join(', ')}.`);
  return { key:'', value:'' };
}


/*** === MENU / DIALOG OPENERS === ***/
function openRecordPayment() {
  RP_LOG('[openRecordPayment] opening dlg_record_payment_v1.html');
  try { rp_markActiveMasterRowIndex_(); } catch(_){}
  const html = HtmlService.createHtmlOutputFromFile('dlg_record_payment_v1').setWidth(980).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Record Payment');
}
function rp_ping(){ RP_LOG('[rp_ping]'); return 'pong'; }


/*** Pin the selected 100_ row for focus‑safe prefill ***/
function rp_markActiveMasterRowIndex_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RP_MASTER_SHEET);
  const rng = sh && sh.getActiveRange();
  const row = rng ? rng.getRow() : 0;
  if (row >= 2) CacheService.getUserCache().put('RP_ACTIVE_MASTER_ROW', String(row), 300);
  return row;
}


/*** === UTILS === ***/
function rp_money(n){ var v=Number(n||0); if(!isFinite(v)) v=0; var parts=v.toFixed(2).split('.'); parts[0]=parts[0].replace(/\B(?=(\d{3})+(?!\d))/g,','); return '$'+parts.join('.'); }
function rp_fmtDateYMD_(d){ return Utilities.formatDate(d, RP_TZ, 'yyyy-MM-dd'); }
function rp_fileIdFromUrl(url){ const s=String(url==null?'':url); let m=s.match(/\/d\/([-\w]{25,})/); if(m&&m[1]) return m[1]; m=s.match(/[?&]id=([-\w]{25,})/); if(m&&m[1]) return m[1]; m=s.match(/[-\w]{25,}/); return m?m[0]:''; }
/** Keep folder names safe & tidy */
function rp_sanitizeForFolder_(s) {
  return String(s || '')
    .trim()
    .replace(/[\\\/]+/g, '-')   // remove slashes
    .replace(/\s+/g, ' ')       // collapse spaces
    .replace(/^-+|-+$/g, '');   // trim leading/trailing dashes
}


function rp_soEq(a,b){ const sa=String(a==null?'':a).trim(), sb=String(b==null?'':b).trim(); if(sa===sb) return true; const na=Number(sa.replace(/[^\d.]/g,'')), nb=Number(sb.replace(/[^\d.]/g,'')); if(!isNaN(na)&&!isNaN(nb)) return Math.abs(na-nb)<1e-9; return false; }
function rp_headerMap(values) { const headers = (values && values[0]) || []; const map = {}; headers.forEach((h,i)=>{ map[String(h).trim()] = i; }); return map; }
function rp_hIndex_(headerRow) { const H = {}; (headerRow||[]).forEach((h,i)=>{ const k=String(h||'').trim(); if (k) H[k]=i+1; }); return H; }
function rp_pick(H, ...names) { for (const n of names) { if (H[n]) return H[n]; } return 0; }
function rp_pick0(map, ...names) { for (const n of names) { if (map[n] != null) return map[n]; } return -1; }
function rp_num_(v) {
  // strip EVERYTHING except digits, dot, and minus → handles $, spaces, commas, etc.
  const s = String(v == null ? '' : v).replace(/[^\d.\-]/g, '');
  const n = parseFloat(s);
  return isFinite(n) ? n : 0;
}


function rp_getHeaderRowCached_(sh) {
  const cache = CacheService.getUserCache();
  const key = 'HDR::' + sh.getParent().getId() + '::' + sh.getSheetId();
  const hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch(_){ } }
  const lc = sh.getLastColumn();
  const row = sh.getRange(1,1,1,lc).getValues()[0].map(v=>String(v).trim());
  try { cache.put(key, JSON.stringify(row), 120); } catch(_){}
  return row;
}


/*** === ACTIVE MASTER ROW === ***/
function rp_activeMasterRow() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RP_MASTER_SHEET);
  if (!sh) throw new Error(`Missing sheet "${RP_MASTER_SHEET}"`);
  const rng = sh.getActiveRange();
  if (!rng) throw new Error('No active range selected.');
  const row = rng.getRow();
  if (row === 1) throw new Error('Header row selected. Click a client row.');
  const lc = sh.getLastColumn();


  const header = sh.getRange(1,1,1,lc).getDisplayValues();
  const map = rp_headerMap(header);
  const rowVals = sh.getRange(row,1,1,lc).getDisplayValues()[0];


  const apptIdx = map['APPT_ID'] != null ? map['APPT_ID'] : (map['RootApptID'] != null ? map['RootApptID'] : map['Root Appt ID']);
  const custIdx = rp_pick0(map, 'Customer Name','Customer','Client Name','Client');
  const soIdx   = rp_pick0(map, 'SO#','SO','SO Number','Sales Order','Sales Order #');
  const trkIdx  = (map['3D Tracker'] != null ? map['3D Tracker'] : map['3D Log']);


  if (apptIdx == null) throw new Error('Missing "APPT_ID" / RootApptID column on 00_Master Appointments.');
  if (custIdx == null) throw new Error('Missing "Customer Name" column on 00_Master Appointments.');


  let trackerUrl = '';
  if (trkIdx != null) {
    trackerUrl = String(rowVals[trkIdx] || '').trim();
    if (!trackerUrl) {
      try {
        const rich = sh.getRange(row, trkIdx + 1).getRichTextValue();
        if (rich) {
          trackerUrl = rich.getLinkUrl() || '';
          if (!trackerUrl && rich.getRuns) {
            const runs = rich.getRuns();
            for (let i = 0; i < runs.length; i++) {
              const u = runs[i].getLinkUrl && runs[i].getLinkUrl();
              if (u) { trackerUrl = u; break; }
            }
          }
        }
      } catch (_) {}
    }
  }

  return {
    rowIndex: row,
    rootApptId: String(rowVals[apptIdx] || '').trim(),
    customerName: String(rowVals[custIdx] || '').trim(),
    soNumber: String((soIdx != null ? rowVals[soIdx] : '') || '').trim(),
    trackerUrl: trackerUrl,
    map,
    rowVals,
    sh
  };
}


/*** === ORDERS LOOKUPS === ***/
function rp_getOrdersTargets() {
  const hpFile = rp_propOneOf_(RP_KEY_ALIASES.ORDERS_HPUSA_FILE_ID, { label:'HPUSA Orders fileId' }).value;
  const vvsFile = rp_propOneOf_(RP_KEY_ALIASES.ORDERS_VVS_FILE_ID, { label:'VVS Orders fileId' }).value;
  const commonTab = rp_propOneOf_(RP_KEY_ALIASES.ORDERS_TAB_COMMON, { label:'Orders common tab name' }).value || '1. Sales';
  const hpTab = rp_propOneOf_(RP_KEY_ALIASES.ORDERS_HPUSA_TAB).value || commonTab;
  const vvsTab = rp_propOneOf_(RP_KEY_ALIASES.ORDERS_VVS_TAB).value  || commonTab;
  const out = [];
  if (hpFile) out.push({ brand:'HPUSA', fileId:hpFile, tabName:hpTab });
  if (vvsFile) out.push({ brand:'VVS',   fileId:vvsFile, tabName:vvsTab });
  return out;
}
function rp_lookupSOAcrossBrands(soNumber) {
  if (!soNumber) return null;
  const cache = CacheService.getUserCache();
  const key = 'SO_SNAP::' + String(soNumber).trim();
  const hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (_) {} }
  const targets = rp_getOrdersTargets();
  rp_log('[rp_lookupSOAcrossBrands] so=', soNumber, ' targets=', targets.map(t => t.brand + ':' + t.tabName).join(','));
  for (const t of targets) {
    const ss = SpreadsheetApp.openById(t.fileId);
    const sh = ss.getSheetByName(t.tabName);
    if (!sh) continue;
    const lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (lr < 2 || lc < 1) continue;

    const headers = sh.getRange(1,1,1,lc).getValues()[0].map(v => String(v).trim());
    const map = {}; headers.forEach((h,i) => map[h] = i);

    const soIdx   = map['SO#'] != null ? map['SO#'] : map['SO'];
    const otIdx   = map['Order Total'] != null ? map['Order Total'] : map['Order Total '];
    const ptdIdx  = rp_pick0(map, 'Paid-to-Date','Paid-To-Date','Paid to Date','Paid-to-date');
    const balIdx  = rp_pick0(map, 'Remaining Balance','Balance');
    const lpdIdx  = rp_pick0(map, 'Last Payment Date','LastPaymentDate');
    const pfIdx   = rp_pick0(map, 'PaymentsFolderURL');

    if (soIdx == null || otIdx == null || ptdIdx < 0) continue;

    const vals = sh.getRange(2,1,lr-1,lc).getValues();
    for (let i = 0; i < vals.length; i++) {
      const row = vals[i];
      if (rp_soEq(row[soIdx], soNumber)) {
        const snap = {
          brand: t.brand,
          sheetName: t.tabName,
          soNumber,
          orderTotal: row[otIdx] || '',
          paidToDate: row[ptdIdx] || '',
          balance: (balIdx >= 0 ? row[balIdx] : ''),
          lastPaymentDate: (lpdIdx >= 0 ? row[lpdIdx] : ''),
          paymentsFolderURL: (pfIdx >= 0 ? row[pfIdx] : '')
        };
        try { cache.put(key, JSON.stringify(snap), 60); } catch (_){}
        rp_log('[rp_lookupSOAcrossBrands] HIT brand=', t.brand, ' tab=', t.tabName, ' so=', soNumber);
        return snap;
      }
    }
  }
  return null;
}


/*** === PREFILL API === ***/
/*** === PREFILL API (100_/400_ only) === ***/
function rp_init() {
  Logger.log('[rp_init] start');
  const stop = rp_time && rp_time('[rp_init] total');

  try {
    // 1) Resolve the active 100_ row (keeps your “select row then open dialog” UX)
    const cache = CacheService.getUserCache();
    const cached = cache.get('RP_ACTIVE_MASTER_ROW');
    const forcedRow = cached ? Number(cached) : 0;
    const master = (forcedRow >= 2) ? rp_getMasterRowByIndex_(forcedRow) : rp_activeMasterRow();

    // 2) Read the active row (typed values) to pull OT/PTD from 100_ only
    const sh = master.sh, map = master.map, rowVals = master.rowVals;
    const rowValsRaw = sh.getRange(master.rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];

    const otM  = map['Order Total'] != null ? rp_num_(rowValsRaw[map['Order Total']]) : 0;
    const ptdM = (function(){
      const idx = rp_pick0(map, 'Paid-to-Date','Paid-To-Date','Paid to Date','Paid-to-date');
      return idx >= 0 ? rp_num_(rowValsRaw[idx]) : 0;
    })();

    // 3) Derive brand + anchor purely from 100_ (no 301/302 lookups)
    const brand      = (map['Brand'] != null) ? String(rowVals[map['Brand']] || '').trim() : '';
    const hasSO      = !!(master.soNumber && String(master.soNumber).trim());
    const anchorType = hasSO ? 'SO' : 'APPT';   // <-- key fix: honor the SO# on 100_

    // 4) Prefill numbers (OT/PTD/balance) purely from 100_
    const orderTotal = otM > 0 ? otM : 0;
    const paidToDate = ptdM > 0 ? ptdM : 0;
    const balance    = Math.max(0, orderTotal - paidToDate);

    // 5) Payments Folder: only from 100_ (column “PaymentsFolderURL” if present)
    const paymentsFolderURL = (function(){
      const pfIdx = rp_pick0(map, 'PaymentsFolderURL');
      return pfIdx >= 0 ? String(rowVals[pfIdx] || '').trim() : '';
    })();

    // 6) Last payment date (optional quality-of-life from 400_ only; no orders read)
    let lastPaymentDate = '';
    try {
      const prev = rp_prevPaymentsForAnchor_({
        anchorType: anchorType,
        rootApptId: master.rootApptId,
        soNumber:   master.soNumber,
        limit: 1
      });
      if (prev && prev.items && prev.items[0] && prev.items[0].date) lastPaymentDate = String(prev.items[0].date);
    } catch (_) {}

    // 7) Base payload to the UI
    const out = {
      anchorType,
      brand,
      rootApptId:      master.rootApptId,
      customerName:    master.customerName,
      soNumber:        master.soNumber || '',
      trackerUrl:      master.trackerUrl || '',
      orderTotal:      String(orderTotal || ''),
      paidToDate:      String(paidToDate || ''),
      balance:         String(balance),
      lastPaymentDate: lastPaymentDate,
      paymentsFolderURL,
      masterRowIndex:  master.rowIndex
    };

    // 8) “Saved Lines” (100_ first; else last from 400_ only)
    try {
      const mObj  = rp_getMasterRowByIndex_(out.masterRowIndex);
      var saved   = rp_readSavedLinesFromMaster_(mObj);
    } catch(_) {}

    if (!saved) {
      try {
        saved = rp_findLastSavedLinesForAnchor_({
          anchorType: out.anchorType,
          rootApptId: out.rootApptId,
          soNumber:   out.soNumber
        });
      } catch(_) {}
    }

    if (saved && saved.lines && saved.lines.length) {
      out.savedLines    = saved.lines;
      out.savedSubtotal = saved.subtotal || 0;
    }

    // 9) Done
    Logger.log('[rp_init] out: ' + JSON.stringify(out));
    if (stop) stop();
    Logger.log('[rp_init] end');
    return out;

  } catch (e) {
    Logger.log('[rp_init] ERROR: ' + (e && e.stack ? e.stack : e));
    throw e;
  } finally {
    Logger.log('[rp_init] end');
  }
}


function rp_listDocNumbersForAnchor({ anchorType, rootApptId, soNumber, limit } = {}) {
  const { sh } = rp_getLedgerTarget();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return [];

  const head = rp_getHeaderRowCached_(sh);
  const H = {}; head.forEach((h,i)=> H[h]=i);
  const cType = H['DocType'], cAppt = H['RootApptID'], cSO = H['SO#'];
  const cDoc  = (H['DocNumber'] != null ? H['DocNumber'] : H['Doc #']);
  if (cType == null || cAppt == null || cSO == null || cDoc == null) return [];

  const start = Math.max(2, lr - 1000);               // scan window
  const vals  = sh.getRange(start,1,lr-start+1,lc).getValues();
  const out   = [];

  for (let i = vals.length - 1; i >= 0; i--) {        // newest → oldest
    const r = vals[i];
    const match = String(anchorType||'').toUpperCase()==='SO'
      ? (String(r[cSO]||'').trim() === String(soNumber||'').trim())
      : (String(r[cAppt]||'').trim() === String(rootApptId||'').trim());
    if (!match) continue;
    const dn = String(r[cDoc]||'').trim();
    if (dn) out.push(dn);
    if (limit && out.length >= limit) break;
  }
  return out;
}



/*** === 3D SPEC HELPERS === ***/
function rp_getLatest3DFields(state) {
  try {
    if (!state || !(state.soNumber || state.rootApptId)) return { ok:false, reason:'BAD_STATE', spec:null };
    const res = rp_get3DSpecFromTracker(state.trackerUrl, state.soNumber, state.rootApptId);
    if (!res || !res.ok || !res.spec) return res || { ok:false, reason:'NO_SPEC', spec:null };
    const s = res.spec || {};
    const hasAny = !!((s.ringStyle && String(s.ringStyle).trim()) || (s.metalType && String(s.metalType).trim()) ||
                      (s.accentType && String(s.accentType).trim()) || (s.ringSize && String(s.ringSize).trim()) ||
                      (s.centerType && String(s.centerType).trim()) || (s.dimensions && String(s.dimensions).trim()));
    return hasAny ? { ok:true, reason:'OK', spec:s } : { ok:false, reason:'EMPTY_FIELDS', spec:null };
  } catch (e) {
    return { ok:false, reason:'EXCEPTION: ' + (e && e.message ? e.message : e), spec:null };
  }
}
function rp_get3DSpecFromTracker(trackerUrl, soNumber, rootApptId) {
  if (!trackerUrl) return { ok:false, reason:'NO_3D_TRACKER_URL', spec:null };
  const fileId = rp_fileIdFromUrl(trackerUrl);
  if (!fileId)   return { ok:false, reason:'BAD_TRACKER_URL', spec:null };


  const ss = SpreadsheetApp.openById(fileId);
  const sh = ss.getSheetByName('Log') || ss.getSheetByName('3D Log') || ss.getSheetByName('3D Revision Log');
  if (!sh) return { ok:false, reason:'LOG_TAB_NOT_FOUND', spec:null };


  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return { ok:false, reason:'LOG_EMPTY', spec:null };


  const headers = sh.getRange(1, 1, 1, lc).getValues()[0].map(v => String(v).trim());
  const H = {}; headers.forEach((h,i)=> H[h]=i);
  function pick(){ for (var i=0;i<arguments.length;i++){ if (H[arguments[i]] != null) return H[arguments[i]]; } return null; }


  const cTS = pick('Timestamp','Date','Time','Submitted At','Created At','Updated');
  const cSO = pick('SO#','SO','Sales Order','Sales Order #');
  const cAP = pick('Root Appt ID','RootApptID','APPT_ID','Appt ID','Root Appt','Appointment ID');


  const cStyle  = pick('Ring Style','Style');
  const cMetal  = pick('Metal Type','Metal','Metal (Type)');
  const cAccent = pick('Accent Diamond Type','Accent Type','Accent');
  const cSize   = pick('Ring Size','US Size');
  const cCenter = pick('Center Stone Type','Center Type');
  const cDims   = pick('Stone Dimensions (mm)','Center Stone Dimensions (mm)','Dimensions (mm)','Dimensions','Measurements (mm)','Measurements');


  const vals = sh.getRange(2, 1, lr-1, lc).getValues();
  let best = null;
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i];


    // --- match logic (unchanged) ---
    let match = false;
    if (cSO != null && soNumber) match = rp_soEq(r[cSO], soNumber);
    if (!match && cAP != null && rootApptId) {
      const a = String(r[cAP] || '').toUpperCase().replace(/[\u200B-\u200D\uFEFF]/g,'').trim();
      const b = String(rootApptId||'').toUpperCase().replace(/[\u200B-\u200D\uFEFF]/g,'').trim();
      match = !!a && !!b && (a === b || a.endsWith(b) || b.endsWith(a));
    }
    if (!match) continue;


    // --- timestamp scoring (unchanged) ---
    let t = 0;
    const v = r[cTS];
    if (v instanceof Date) t = v.getTime();
    else if (v) { const p = Date.parse(String(v)); if (!isNaN(p)) t = p; } else { t = (i+1); }


    // --- candidate extraction ---
    const candidate = {
      t, r,
      ringStyle:   cStyle != null  ? String(r[cStyle]  || '').trim() : '',
      metalType:   cMetal != null  ? String(r[cMetal]  || '').trim() : '',
      accentType:  cAccent != null ? String(r[cAccent] || '').trim() : '',
      ringSize:    cSize != null   ? String(r[cSize]   || '').trim() : '',
      centerType:  cCenter != null ? String(r[cCenter] || '').trim() : '',
      dimensions:  cDims != null   ? String(r[cDims]   || '').trim() : ''
    };


    // Prefer the first row with the maximum timestamp (same as stable-desc sort pick)
    if (!best || candidate.t > best.t) best = candidate;
  }
  if (!best) return { ok:false, reason:'NO_MATCH', spec:null };


  const out = {
    ringStyle:best.ringStyle,
    metalType:best.metalType,
    accentType:best.accentType,
    ringSize:best.ringSize,
    centerType:best.centerType,
    dimensions:best.dimensions
  };
  if (!out.ringStyle && !out.metalType && !out.accentType && !out.ringSize && !out.centerType && !out.dimensions) {
    return { ok:false, reason:'NO_FIELDS_FOR_MATCH', spec:null };
  }
  return { ok:true, reason:'OK', spec: out };
}


/*** === LEDGER TARGET === ***/
function rp_getLedgerTarget() {
  const fileRes = rp_propOneOf_(RP_KEY_ALIASES.LEDGER_FILE_ID, { required:true, label:'Payments Ledger File ID' });
  const sheetRes = rp_propOneOf_(RP_KEY_ALIASES.LEDGER_SHEET_NAME, { required:false, label:'Payments sheet name' });
  const fileId = fileRes.value;
  const sheetName = sheetRes.value || 'Payments';
  const ss = SpreadsheetApp.openById(fileId);
  rp_log('[rp_getLedgerTarget] fileId=', fileId, ' sheet=', sheetName);
  const sh = ss.getSheetByName(sheetName) || ss.getSheets()[0];
  return { ss, sh, resolved: { ledgerFileKey:fileRes.key, ledgerSheetKey:sheetRes.key || '(default: Payments)' } };
}
function rp_ensureHeaders_(sh, headersNeeded) {
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 1 || lc < 1) sh.getRange(1,1,1,1).setValue('');
  const current = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0];
  const map = {};
  for (let i = 0; i < current.length; i++) map[String(current[i]).trim()] = i;
  let cursor = current.length;
  headersNeeded.forEach(h => {
    if (map[h] == null) { sh.getRange(1, cursor+1).setValue(h); map[h] = cursor; cursor++; }
  });
  return map;
}


/*** === FEES === ***/
const RP_FEES_CACHE_KEY = 'PAYMENT_FEES::v1';
const RP_FEES_CACHE_TTL = 3600;
function rp_parseFeeCell_(v){ if(v==null||v==='') return 0; const s=String(v).trim(); if(/%/.test(s)){ const n=parseFloat(s.replace(/[^\d.\-]/g,'')); return isNaN(n)?0:n/100;} const n=parseFloat(s.replace(/[^\d.\-]/g,'')); if(isNaN(n)) return 0; return n>1 ? n/100 : Math.max(0,n); }
function rp_readFeesFromSheet_(){ try{ const p=PropertiesService.getScriptProperties(); const tab=p.getProperty('PAYMENTS_FEES_TAB_NAME')||p.getProperty('CFG_PAYMENTS_FEES_TAB_NAME')||'Current Fees'; const { sh }=rp_getLedgerTarget(); const ss=sh.getParent(); const s=ss.getSheetByName(tab); if(!s) return null; const lr=s.getLastRow(), lc=s.getLastColumn(); if(lr<2||lc<1) return null; const headers=s.getRange(1,1,1,lc).getValues()[0].map(v=>String(v).trim()); const hmap={}; headers.forEach((h,i)=>hmap[h]=i); const methodCol=hmap['Method']!=null ? hmap['Method'] : hmap['Payment Method']; const feeCol=hmap['Fee %']!=null ? hmap['Fee %'] : (hmap['Fee']!=null?hmap['Fee']:hmap['Percent']); if(methodCol==null||feeCol==null) return null; const vals=s.getRange(2,1,lr-1,lc).getValues(); const out={}; vals.forEach(row=>{ const m=String(row[methodCol]||'').trim(); if(!m) return; out[m] = rp_parseFeeCell_(row[feeCol]); }); return Object.keys(out).length ? out : null; } catch(_){ return null; } }
function rp_readFeesFromProp_(){ const p=PropertiesService.getScriptProperties(); const raw=p.getProperty('PAYMENT_FEES_JSON')||p.getProperty('CFG_PAYMENT_FEES_JSON'); if(!raw) return null; try{ const obj=JSON.parse(raw); const out={}; Object.keys(obj).forEach(k=>{ out[k]=rp_parseFeeCell_(obj[k]); }); return out; }catch(_){ return null; } }
function rp_getDefaultFees_(){ return {"Card":0.03,"Synchrony":0.06,"Wire":0,"Zelle":0,"Cash":0,"Check":0,"Other":0}; }
function rp_getFees(){ const cache=CacheService.getUserCache(); const cached=cache.get(RP_FEES_CACHE_KEY); if(cached){ try{ return JSON.parse(cached); }catch(_){}} let fees=rp_readFeesFromProp_(); if(!fees) fees=rp_readFeesFromSheet_(); if(!fees) fees=rp_getDefaultFees_(); try{ cache.put(RP_FEES_CACHE_KEY, JSON.stringify(fees), RP_FEES_CACHE_TTL); }catch(_){} return fees; }
function rp_refreshFeesCache(){ CacheService.getUserCache().remove(RP_FEES_CACHE_KEY); return 'Fees cache cleared.'; }


/*** === SUBMIT === ***/
function rp_amountForMasterOnSOReceipt_(pmt) {
  const prop = (PropertiesService.getScriptProperties().getProperty('SO_RECEIPT_MASTER_AMOUNT') ||
                PropertiesService.getScriptProperties().getProperty('CFG_SO_RECEIPT_MASTER_AMOUNT') || 'ALLOC').toUpperCase();
  return prop === 'GROSS' ? Number((pmt && pmt.amount) || 0) : Number((pmt && pmt.allocatedToSO) || 0);
}
function rp_submit(payload) {
  if (!payload) throw new Error('Empty submit payload.');
  const { anchorType, brand, rootApptId, soNumber, docType, lines, pmt } = payload;
  if (!docType) throw new Error('Doc Type is required.');
  if (!lines || !lines.length) throw new Error('At least one line is required.');


  const subtotal = lines.reduce((s, ln) => s + (Number(ln.qty||0) * Number(ln.amt||0)), 0);
  if (!(subtotal > 0)) throw new Error('Lines subtotal must be greater than 0.');


  const isReceipt = /Receipt/i.test(docType);
  const amountGross = isReceipt ? Number(pmt.amount||0) : 0;
  if (isReceipt && !(amountGross > 0)) throw new Error('Payment Amount is required for receipts.');


  const fees = rp_getFees();
  const feePct = isReceipt ? Number(fees[pmt.method] || 0) : 0;
  const feeAmt = isReceipt ? +(amountGross * feePct).toFixed(2) : 0;
  const amountNet = isReceipt ? +(amountGross - feeAmt).toFixed(2) : 0;


  // Allocation (UI disabled for now) → AllocatedToSO = Gross
  const allocToSO = isReceipt ? Number(pmt.amount || 0) : 0;


  const stamp = Utilities.formatDate(new Date(), RP_TZ, 'yyyyMMdd-HHmmss');
  const anchorKey = soNumber ? soNumber : rootApptId;
  const basketId = `BASK-${anchorKey}-${stamp}`;
  const paymentId = `PAY-${Utilities.getUuid()}`;


  const submittedAt = new Date();
  const submittedBy = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'Unknown';

  // === Phase 2 normalize (doc markers) ===
  var docStatus  = String(payload.docStatus || '').toUpperCase() || 'DRAFT';
  var docRole    = String(payload.docRole   || '').toUpperCase();
  var supersedes = String(payload.supersedes|| '').trim();
  var appliesTo  = String(payload.appliesTo || '').trim();

  // If the dialog leaves role blank, infer a reasonable default from docType
  if (!docRole) {
    var dt = String(payload.docType || '').toUpperCase();
    if (dt.indexOf('CREDIT')   >= 0) docRole = 'CREDIT';
    else if (dt.indexOf('PROGRESS') >= 0) docRole = 'PROGRESS';
    else if (dt.indexOf('DEPOSIT')  >= 0 && dt.indexOf('INVOICE') >= 0) docRole = 'DEPOSIT';
    else if (dt.indexOf('INVOICE')  >= 0) docRole = 'FINAL';
    else {
      // Receipts: treat as SALES_RECEIPT if there are lines, else PAYMENT_RECEIPT
      var hasLines = Array.isArray(payload.lines) && payload.lines.length;
      docRole = hasLines ? 'SALES_RECEIPT' : 'PAYMENT_RECEIPT';
    }
  }



  const rowObj = {
    'PAYMENT_ID': paymentId, 'Brand': brand || '', 'RootApptID': rootApptId || '', 'SO#': soNumber || '',
    'AnchorType': anchorType || '', 'BasketID': basketId, 'DocType': docType,
    'PaymentDateTime': isReceipt ? (pmt.dateTime || '') : '',
    'Method': isReceipt ? (pmt.method || '') : '', 'Reference': isReceipt ? (pmt.reference || '') : '', 'Notes': isReceipt ? (pmt.notes || '') : '',
    'AmountGross': amountGross, 'FeePercent': feePct, 'FeeAmount': feeAmt, 'AmountNet': amountNet, 'AllocatedToSO': allocToSO,
    'LinesJSON': JSON.stringify(lines), 'Subtotal': +subtotal.toFixed(2),
    'Order Total_SO': (payload.snapshots && payload.snapshots.orderTotal) ? String(payload.snapshots.orderTotal) : '',
    'Paid-To-Date_SO': (payload.snapshots && payload.snapshots.paidToDate) ? String(payload.snapshots.paidToDate) : '',
    'Balance_SO': (payload.snapshots && payload.snapshots.balance) ? String(payload.snapshots.balance) : '',
    'Submitted By': submittedBy, 'Submitted Date/Time': submittedAt
  };


  try {
    if ((!rowObj['Brand'] || !String(rowObj['Brand']).trim()) && anchorType === 'APPT') {
      const m = rp_findMasterRowByRootApptId_(rootApptId);
      if (m && m.map['Brand'] != null) { rowObj['Brand'] = String(m.rowVals[m.map['Brand']] || '').trim(); }
    }
  } catch(_){}


  const { sh } = rp_getLedgerTarget();
  const headersNeeded = Object.keys(rowObj);
  const map = rp_ensureHeaders_(sh, headersNeeded);
  const nextRow = sh.getLastRow() + 1;
  const rowArr = new Array(Math.max(...Object.values(map)) + 1).fill('');
  for (const [key, val] of Object.entries(rowObj)) { rowArr[map[key]] = val; }
  sh.getRange(nextRow, 1, 1, rowArr.length).setValues([rowArr]);

  // === Phase 2: ensure columns exist, then write markers ===
  // 1) Make sure the 4 headers exist (adds them to row 1 if missing)
  rp_ensureHeaders_(sh, ['DocStatus','DocRole','SupersedesDoc#','AppliesToDoc#']);

  // 2) Build a 1‑based header map, then pick the 4 columns
  var headerRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var H1 = rp_hIndex_(headerRow);                      // 1‑based indices
  var cDocStatus  = rp_pick(H1, 'DocStatus','Status'); // pick() returns 1‑based or 0 if absent
  var cDocRole    = rp_pick(H1, 'DocRole','Role');
  var cSupersedes = rp_pick(H1, 'SupersedesDoc#','Supersedes','Replaces','ReplacesDoc#');
  var cAppliesTo  = rp_pick(H1, 'AppliesToDoc#','Applies To','SettlesDoc#','Settles');

  // 3) Write values into the row we just added
  if (cDocStatus)  sh.getRange(nextRow, cDocStatus ).setValue(docStatus);
  if (cDocRole)    sh.getRange(nextRow, cDocRole   ).setValue(docRole);
  if (cSupersedes) sh.getRange(nextRow, cSupersedes).setValue(supersedes);
  if (cAppliesTo)  sh.getRange(nextRow, cAppliesTo ).setValue(appliesTo);


  // Order Total writes whenever the checkbox is selected (Invoices AND Receipts)
  const setFlag = !!(payload && payload.flags && payload.flags.setOrderTotal);
  if (setFlag) {
    if (anchorType === 'APPT') {
      try {
        const resMaster = rp_setOrderTotal_Master_Safe_(rootApptId, subtotal, true, payload.masterRowIndex);
        rp_auditOrderTotalOnLedger_(nextRow, {
          set: !!resMaster.updated,
          value: (resMaster.updated ? subtotal : ''),
          prev: resMaster.prev,
          source: docType,
          target: 'APPT'
        });
      } catch (e) { Logger.log('APPT OT write error: ' + (e && e.message ? e.message : e)); }
    } else if (anchorType === 'SO') {
      try {
        const resMaster2 = rp_setOrderTotal_Master_Safe_(rootApptId, subtotal, true, payload.masterRowIndex);
        rp_auditOrderTotalOnLedger_(nextRow, {
          set: !!resMaster2.updated,
          value: (resMaster2.updated ? subtotal : ''),
          prev: resMaster2.prev,
          source: docType,
          target: 'MASTER'
        });
      } catch (e) { Logger.log('SO OT write error: ' + (e && e.message ? e.message : e)); }
    }
  }


  // Persist Saved Lines JSON/Subtotal to 100_ (and 301/302 for SO) when OT is set
  if (setFlag) {
    try {
      rp_persistSavedLinesToMaster_({ masterRowIndex: payload.masterRowIndex, rootApptId, lines: payload.lines, subtotal });
    } catch (e) {
      Logger.log('Saved lines persist warning: ' + (e && e.message ? e.message : e));
    }
  }

  // Receipts write-back: 100_ only (+ supersedes handling)
  try {
    if (/receipt/i.test(docType)) {
      const when = (payload && payload.pmt && payload.pmt.dateTime) ? new Date(payload.pmt.dateTime) : new Date();

      // Base amount applied to 100_ for this receipt
      const amtForMaster =
        (anchorType === 'SO')
          ? rp_amountForMasterOnSOReceipt_(payload.pmt || { amount: amountGross, allocatedToSO: allocToSO })
          : amountGross;

      // If this receipt supersedes a prior receipt, void the old row and apply a NET delta to PTD
      let netAmtForMaster = amtForMaster;
      const supersedes = String(payload && payload.supersedes || '').trim();

      if (supersedes) {
        try {
          const sup = rp_findLedgerRowByDocNumber_(supersedes);
          if (sup) {
            const tOld = String(sup.rowVals[sup.H['DocType']] || '').toUpperCase();
            const statusOld = (sup.H['DocStatus'] != null ? String(sup.rowVals[sup.H['DocStatus']] || '') : '').toUpperCase().trim();

            // Ensure same anchor
            const sameAnchor = (function(){
              const aNew = String(anchorType || '').toUpperCase();
              if (aNew === 'SO') return rp_soEq(sup.rowVals[sup.H['SO#']], soNumber);
              return String(sup.rowVals[sup.H['RootApptID']] || '').trim() === String(rootApptId || '').trim();
            })();

            // Only affect PTD if superseded doc is a RECEIPT on the same anchor and not already VOID/REPLACED
            if (tOld.includes('RECEIPT') && sameAnchor && statusOld !== 'VOID' && statusOld !== 'REPLACED') {
              const prevApplied = rp_getAppliedAmtForMasterOnReceiptRow_(sup.rowVals, sup.H) || 0;
              netAmtForMaster = amtForMaster - prevApplied;
            
            // ... inside the `if (supersedes) { ... }` try block after we computed tOld/statusOld/sameAnchor
            if (tOld.includes('INVOICE') && sameAnchor && statusOld !== 'VOID' && statusOld !== 'REPLACED') {
              rp_updateLedgerRow_(sup.row, { 'DocStatus': 'REPLACED' });
            }

              // Flip old row's status → future sums/blocks ignore it
              rp_updateLedgerRow_(sup.row, { 'DocStatus': 'VOID' });
            }
          }
        } catch (e) {
          Logger.log('Supersedes handling warning: ' + (e && e.message ? e.message : e));
        }
      }

      if (payload.masterRowIndex) {
        rp_applyReceiptToMaster({ masterRowIndex: payload.masterRowIndex, amount: netAmtForMaster, when });
      }
    }
  } catch (e) { Logger.log('Receipt write-back warning: ' + (e && e.message ? e.message : e)); }


  // Refresh 100_ "Cash-in (Gross)" after receipt
  try {
    if (/receipt/i.test(docType)) {
      const mRow = Number(payload.masterRowIndex || 0);
      if (mRow >= 2 && rootApptId) { rp_updateMasterCashInGross_({ masterRowIndex: mRow, rootApptId: rootApptId }); }
    }
  } catch (e) { Logger.log('[Cash-in Gross] refresh skipped: ' + (e && e.message ? e.message : e)); }


  // First receipt flips Sales Stage to "Deposit"
  try {
    if (/receipt/i.test(docType)) {
      const count = rp_countReceiptsForAppt_(rootApptId);
      if (count === 1 && payload.masterRowIndex) rp_setSalesStageOnMaster_({ masterRowIndex: payload.masterRowIndex, value: 'Deposit', allowOverride: false });
    }
  } catch (e) { Logger.log('Sales Stage set skipped: ' + ((e && e.message) ? e.message : e)); }


  return { ok:true, paymentId, basketId, row: nextRow };
}


/*** === TEMPLATE / PLACEHOLDERS + TABLE RENDERING === ***/
function rp_docCodeFromDocType_(docType) {
  const t = String(docType || '').toLowerCase();
  if (t.includes('deposit') && t.includes('invoice')) return { code:'DI', family:'Deposit' };
  if (t.includes('deposit') && t.includes('receipt')) return { code:'DR', family:'Deposit' };
  if (t.includes('sales')   && t.includes('invoice')) return { code:'SI', family:'Sales'   };
  if (t.includes('sales')   && t.includes('receipt')) return { code:'SR', family:'Sales'   };
  return { code:'UNK', family:'Deposit' };
}
function rp_getTemplateIdFor(brand, docType) {
  const p = PropertiesService.getScriptProperties();

  // Normalize brand to the tokens your properties use
  const bRaw = String(brand || '').toUpperCase();
  const normBrand = /VVS/.test(bRaw) ? 'VVS'
                   : /HPUSA/.test(bRaw) ? 'HPUSA'
                   : bRaw.replace(/[^A-Z0-9]/g, '');

  // Normalize doc type to DEPOSIT_INVOICE, DEPOSIT_RECEIPT, SALES_INVOICE, SALES_RECEIPT
  const normType = String(docType || '')
    .toUpperCase().replace(/[^A-Z]/g, ' ')
    .replace(/\s+/g, ' ').trim().replace(/ /g, '_');

  const codeMap = { DEPOSIT_INVOICE:'DI', DEPOSIT_RECEIPT:'DR', SALES_INVOICE:'SI', SALES_RECEIPT:'SR' };
  const code = codeMap[normType];

  // 1) Prefer the concise per‑brand keys first (your current setup)
  const primary = [
    `${normBrand}_${code}_TEMPLATE_ID`,           // HPUSA_DI_TEMPLATE_ID, VVS_SI_TEMPLATE_ID, etc.
    `CFG_${normBrand}_${code}_TEMPLATE_ID`
  ];

  // 2) Then the canonical long keys
  const canonical = [
    `TEMPLATE_${normType}_${normBrand}`,          // TEMPLATE_DEPOSIT_INVOICE_HPUSA, etc.
    `CFG_TEMPLATE_${normType}_${normBrand}`
  ];

  // 3) Aliases from RP_KEY_ALIASES (covers older typos like *_VS vs *_VVS)
  const aliasKey = `TEMPLATE_${normType}_${normBrand}`;
  const aliasList = (RP_KEY_ALIASES && RP_KEY_ALIASES[aliasKey]) ? RP_KEY_ALIASES[aliasKey] : [];

  const keysToTry = [...primary, ...canonical, ...aliasList];

  for (const k of keysToTry) {
    const v = p.getProperty(k);
    if (v && String(v).trim()) return String(v).trim();
  }

  // 4) Last‑resort title search that includes the brand token
  const titles = [
    `[TEMPLATE] ${docType} -- ${normBrand}`,
    `[TEMPLATE] ${docType} — ${normBrand}`
  ];
  for (const name of titles) {
    const it = DriveApp.searchFiles(`title = "${name}"`);
    if (it.hasNext()) return it.next().getId();
  }

  throw new Error(
    `Template not found. Brand="${brand}" -> "${normBrand}", DocType="${docType}" -> "${normType}". ` +
    `Tried keys: ${keysToTry.join(', ')}`
  );
}

function rp_renderLinesText_(lines){
  const out = (lines||[]).map(ln => String(ln && ln.desc || '').trim()).filter(Boolean)
    .map(s => {
      const t = s.replace(/^\s+|\s+$/g,'');
      return t.startsWith('✧') ? t : ('✧ ' + t);
    });
  return out.join('\n');
}

/** Replace {{PLACEHOLDERS}} in a Doc (robust across body/tables). */
function rp_fillDocPlaceholders_(docId, replacements) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const hdr  = (doc.getHeader && doc.getHeader()) || null;
  const ftr  = (doc.getFooter && doc.getFooter()) || null;

  const repl = Object.assign({}, replacements || {});
  // Canonical cross-mapping (fill both old & new keys)
  if (repl.ORDER_TOTAL_SO != null && repl.ORDER_TOTAL == null) repl.ORDER_TOTAL = repl.ORDER_TOTAL_SO;
  if (repl.PAID_TO_DATE_BEFORE != null && repl.Paid_to_date == null) repl.Paid_to_date = repl.PAID_TO_DATE_BEFORE;
  if (repl.BALANCE_AFTER != null && repl.BALANCE == null) repl.BALANCE = repl.BALANCE_AFTER;
  if (repl.BALANCE_BEFORE != null && repl.BALANCE == null) repl.BALANCE = repl.BALANCE_BEFORE;
  if (repl.pmtId && !repl.PMT_ID) repl.PMT_ID = repl.pmtId;

  // ✅ Bridge common title-case placeholders used by older templates
  (function bridgeTitleCasePlaceholders() {
    function alias(srcKey, aliases) {
      if (repl[srcKey] == null) return;
      aliases.forEach(k => { if (repl[k] == null) repl[k] = repl[srcKey]; });
    }
    // Order Total
    alias('ORDER_TOTAL', ['Order Total', 'ORDER TOTAL']);
    // Paid-to-Date (previous payments under OT line)
    alias('Paid_to_date', ['Paid-To-Date', 'Paid to Date', 'PAID-TO-DATE', 'PAID TO DATE']);
    // Balance (either before or after)
    if (repl.BALANCE != null) alias('BALANCE', ['Balance', 'BALANCE DUE', 'Remaining Balance']);
    // Payment fields (if your templates use these labels)
    alias('PAYMENT_AMOUNT',    ['Payment Amount', 'Amount Paid']);
    alias('REQ_AMT', ['Requested Amount', 'REQUESTED_AMOUNT', 'REQ AMT']);
    alias('PAYMENT_METHOD',    ['Payment Method', 'Method']);
    alias('PAYMENT_REFERENCE', ['Payment Reference', 'Reference']);
  })();

  const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const sections = [body, hdr, ftr].filter(Boolean);

  Object.keys(repl).forEach(key => {
    const pat = '\\{\\{\\s*' + esc(key) + '\\s*\\}\\}';  // string pattern (regex text)
    const val = (repl[key] == null) ? '' : String(repl[key]);
    sections.forEach(sec => sec.replaceText(pat, val));
  });
  doc.saveAndClose();
}
function rp_deletePlaceholderLine_(docId, key){
  const doc = DocumentApp.openById(docId), body = doc.getBody();
  let r = body.findText('\\{\\{\\s*' + key + '\\s*\\}\\}');
  while (r) { const p = r.getElement().getParent().asParagraph(); p.removeFromParent(); r = body.findText('\\{\\{\\s*' + key + '\\s*\\}\\}'); }
  doc.saveAndClose();
}

// Remove just the table row (or the single paragraph line) that contains a placeholder key.
function rp_deletePlaceholderRowOrLine_(docId, key){
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const pat = '\\{\\{\\s*' + key + '\\s*\\}\\}';
  let range = body.findText(pat);
  let removedAny = false;
  while (range) {
    const el = range.getElement();
    // Walk up: Text -> Paragraph -> TableCell -> TableRow -> Table
    let node = el;
    let row = null;
    while (node && node.getParent) {
      try {
        if (node.getType && node.getType() === DocumentApp.ElementType.TABLE_ROW) {
          row = node.asTableRow();
          break;
        }
      } catch (_) {}
      node = node.getParent && node.getParent();
    }
    if (row) {
      // Only remove the row, not the entire table
      row.removeFromParent();
      removedAny = true;
    } else {
      // Fallback if the placeholder is not in a table
      try { el.getParent().asParagraph().removeFromParent(); removedAny = true; } catch(_){}
    }
    range = body.findText(pat);
  }
  doc.saveAndClose();
  return removedAny;
}


/**
 * Insert a "Previous Payments" row directly under the "Order Total" row
 * in the 2‑column summary table (labels | values). Idempotent:
 * - If a "Previous Payments" row already exists, does nothing.
 * - If the template uses a placeholder line elsewhere, also does nothing.
 */
function rp_insertPrevRowUnderOT_(docId, label, amountText) {
  const doc  = DocumentApp.openById(docId);
  const body = doc.getBody();


  // If a "Previous Payments" label already exists anywhere in the doc, skip.
  const existing = body.findText(/Previous Payments/i);
  if (existing) { doc.saveAndClose(); return false; }


  const norm = s => String(s || '').replace(/\s+/g, ' ').trim().toUpperCase();


  // Find the summary table and the row that has "Order Total" in its first cell
  let target = null;
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.TABLE) continue;
    const t = el.asTable();
    for (let r = 0; r < t.getNumRows(); r++) {
      const row = t.getRow(r);
      if (row.getNumCells() < 2) continue; // expect at least 2 columns
      const left = norm(row.getCell(0).getText());
      if (left.includes('ORDER TOTAL')) {
        target = { t, r }; break;
      }
    }
    if (target) break;
  }
  if (!target) { doc.saveAndClose(); return false; }


  // Insert new row right under "Order Total"
  const { t, r } = target;
  // If the immediate next row is already "Previous Payments", skip
  if (r + 1 < t.getNumRows() && norm(t.getRow(r + 1).getCell(0).getText()).includes('PREVIOUS PAYMENTS')) {
    doc.saveAndClose(); return false;
  }


  const cols = t.getRow(0).getNumCells();
  const nr = t.insertTableRow(r + 1);
  nr.appendTableCell(label);                   // first column: label
  for (let c = 1; c < cols - 1; c++) nr.appendTableCell('');  // keep middle columns (if any)
  nr.appendTableCell(String(amountText));      // last column: amount


  doc.saveAndClose();
  return true;
}




/** Find & fill DESCRIPTION/QUANTITY/AMOUNT table. Returns { usedTable: boolean }. */
function rp_fillItemsTable_(docId, lines) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const toUpper = s => String(s || '').trim().toUpperCase();
  let table = null;
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.TABLE) continue;
    const t = el.asTable();
    if (t.getNumRows() < 1 || t.getRow(0).getNumCells() < 3) continue;
    const h0 = toUpper(t.getRow(0).getCell(0).getText());
    const h1 = toUpper(t.getRow(0).getCell(1).getText());
    const h2 = toUpper(t.getRow(0).getCell(2).getText());
    if (h0.includes('DESCRIPTION') && (h1.includes('QTY') || h1.includes('QUANTITY')) && h2.includes('AMOUNT')) { table = t; break; }
  }
  if (!table) { doc.saveAndClose(); return { usedTable:false }; }


  // Insert each line directly under the header (row index 1), preserving any existing rows
  let insertAt = 1;  // 0 = header row; we start inserting at row 1


  (lines || []).forEach(ln => {
    const qty = (ln && ln.qty != null) ? ln.qty : '';
    const amt = Number(ln && ln.amt || 0);
    const total = Number(qty || 0) * amt;


    // Insert new row under header
    const row = table.insertTableRow(insertAt++);


    // Build 3 cells (Description | Quantity | Amount)
    // Keep description exactly as typed; DO NOT auto-prefix bullets
    const cDesc = row.appendTableCell(String((ln && ln.desc) || '').trim());
    const cQty  = row.appendTableCell(String(qty));
    const cAmt  = row.appendTableCell(rp_money(total));


    // Align: Description → LEFT; Quantity & Amount → RIGHT
    function setAlign(cell, align) {
      const n = cell.getNumChildren();
      for (let i = 0; i < n; i++) {
        const kid = cell.getChild(i);
        if (kid && kid.getType && kid.getType() === DocumentApp.ElementType.PARAGRAPH) {
          kid.asParagraph().setAlignment(align);
        }
      }
    }
    setAlign(cDesc, DocumentApp.HorizontalAlignment.LEFT);
    setAlign(cQty,  DocumentApp.HorizontalAlignment.RIGHT);
    setAlign(cAmt,  DocumentApp.HorizontalAlignment.RIGHT);




  });


  doc.saveAndClose();
  return { usedTable:true };
}




/** Format a payments list */
function rp_formatPaymentsList_(items){
  return (items||[]).map(it => {
    const dt = it.date || '';
    const amt = rp_money(Number(it.amount || 0));
    const mth = it.method || '';
    return `✧ ${dt} — ${amt} ${mth}`.trim();
  }).join('\n');
}

/** Prior receipts for anchor (before current ledger row) */
function rp_prevPaymentsForAnchor_({ anchorType, rootApptId, soNumber, beforeRow, limit } = {}) {
  const { sh } = rp_getLedgerTarget();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return { items: [] };

  const head = rp_getHeaderRowCached_(sh);
  const H = {}; head.forEach((h,i)=> H[h]=i);
  const cType = H['DocType'], cAppt = H['RootApptID'], cSO = H['SO#'], cGross = H['AmountGross'], cWhen = H['PaymentDateTime'], cMethod = H['Method'];
  const cStatus = (H['DocStatus'] != null ? H['DocStatus'] : H['Status']);
  if (cType == null || cAppt == null || cSO == null || cGross == null) return { items: [] };

  const start = rp_scanWindowStart_(lr);
  const vals = sh.getRange(start,1,lr-start+1,lc).getValues();
  const out = [];
  for (let i=0;i<vals.length;i++){
    const rowIndex = start + i; // fix: respect scan window
    if (beforeRow && rowIndex >= beforeRow) continue;

    const r = vals[i];
    const type = String(r[cType] || '').toLowerCase();
    if (!(type.includes('receipt'))) continue;

    const status = cStatus != null ? String(r[cStatus] || '').toUpperCase().trim() : '';
    if (status === 'VOID' || status === 'REPLACED' || status === 'DRAFT') continue;

    if (String(anchorType||'').toUpperCase() === 'SO') {
      if (!rp_soEq(r[cSO], soNumber)) continue;
    } else {
      if (String(r[cAppt] || '').trim() !== String(rootApptId || '').trim()) continue;
    }

    const whenRaw = r[cWhen];
    let when = null;
    if (whenRaw instanceof Date) when = whenRaw;
    else if (whenRaw) { const p = Date.parse(String(whenRaw)); if (!isNaN(p)) when = new Date(p); }

    out.push({ when, date: when ? rp_fmtDateYMD_(when) : '', amount: Number(r[cGross] || 0), method: String(r[cMethod] || '') });
  }
  out.sort((a,b)=> (b.when?b.when.getTime():0) - (a.when?a.when.getTime():0));
  return { items: (limit && limit>0) ? out.slice(0, limit) : out };
}

/*** --- Saved Lines helpers (100_ / 301-302 / 400_) --- ***/
function rp_scanWindowStart_(lr) {
  const p = PropertiesService.getScriptProperties();
  const win = Number(p.getProperty('LEDGER_SCAN_WINDOW') || 2000);
  return Math.max(2, lr - Math.max(200, win) + 1);
}

function rp_sanitizeDesc_(s){ return String(s||'').replace(/^\s*✧\s*/gm,'').trim(); }

function rp_readSavedLinesFromMaster_(m) {
  if (!m) return null;
  const sh = m.sh;
  let header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  let H = rp_hIndex_(header);
  let cJSON = rp_pick(H,'Saved Lines JSON','SavedLinesJSON');
  let cSub  = rp_pick(H,'Saved Subtotal','SavedSubtotal');
  if (!cJSON && !cSub) return null;
  const raw = cJSON ? sh.getRange(m.rowIndex, cJSON).getDisplayValue() : '';
  const subtotal = cSub ? Number(String(sh.getRange(m.rowIndex, cSub).getDisplayValue()).replace(/[^\d.\-]/g,'')) || 0 : 0;
  if (!raw) return null;
  try {
    const arr = JSON.parse(raw);
    const lines = (arr||[]).map(ln => ({ desc: rp_sanitizeDesc_(ln.desc), qty: Number(ln.qty)||0, amt: Number(ln.amt)||0 }));
    return { lines, subtotal };
  } catch(_){ return null; }
}

function rp_readSavedLinesFromOrders_(brand, soNumber) {
  const hit = rp_findSoRowInBrand_(brand, soNumber);
  if (!hit) return null;
  const sh = hit.sh;
  let header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  let H = rp_hIndex_(header);
  let cJSON = rp_pick(H,'Saved Lines JSON','SavedLinesJSON');
  let cSub  = rp_pick(H,'Saved Subtotal','SavedSubtotal');
  if (!cJSON && !cSub) return null;
  const raw = cJSON ? sh.getRange(hit.rowIndex, cJSON).getDisplayValue() : '';
  const subtotal = cSub ? Number(String(sh.getRange(hit.rowIndex, cSub).getDisplayValue()).replace(/[^\d.\-]/g,'')) || 0 : 0;
  if (!raw) return null;
  try {
    const arr = JSON.parse(raw);
    const lines = (arr||[]).map(ln => ({ desc: rp_sanitizeDesc_(ln.desc), qty: Number(ln.qty)||0, amt: Number(ln.amt)||0 }));
    return { lines, subtotal };
  } catch(_){ return null; }
}

function rp_findLastSavedLinesForAnchor_({anchorType, rootApptId, soNumber} = {}) {
  const { sh } = rp_getLedgerTarget();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return null;
  const head = sh.getRange(1,1,1,lc).getValues()[0].map(v=>String(v).trim());
  const H = {}; head.forEach((h,i)=> H[h]=i);
  const cType = H['DocType'], cAppt = H['RootApptID'], cSO = H['SO#'], cJSON = H['LinesJSON'], cSub = H['Subtotal'];
  if (cType == null || cAppt == null || cSO == null || cJSON == null) return null;
  const start = rp_scanWindowStart_(lr);
  const vals = sh.getRange(start,1,lr-start+1,lc).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    const r = vals[i];
    const isMatch = (String(anchorType||'').toUpperCase()==='SO')
      ? rp_soEq(r[cSO], soNumber)
      : (String(r[cAppt]||'').trim() === String(rootApptId||'').trim());
    if (!isMatch) continue;
    const raw = String(r[cJSON]||'').trim();
    if (!raw) continue;
    try {
      const arr = JSON.parse(raw);
      const lines = (arr||[]).map(ln => ({ desc: rp_sanitizeDesc_(ln.desc), qty: Number(ln.qty)||0, amt: Number(ln.amt)||0 }));
      const subtotal = cSub != null ? Number(r[cSub]||0) : 0;
      return { lines, subtotal };
    } catch(_){ /* next */ }
  }
  return null;
}


/** Generate Doc+PDF, fill placeholders + items table; return ids+links */
function rp_generateDocAndPdf_(brand, docType, destFolder, payload, ledgerMeta) {
  if (!destFolder) throw new Error('Destination folder not resolved.');
  const tmplId = rp_getTemplateIdFor(brand, docType);
  const now = new Date();
  const so = payload.soNumber || '';
  let docNumber, baseName;

  if (payload.anchorType === 'SO') {
    const version = rp_nextDocVersion_('SO', payload.rootApptId, so, docType);
    docNumber = rp_formatDocNumberForSO({ brand, so, docType, version, dateObj: now });
    baseName  = docNumber;
  } else {
    const seqDocNum = rp_buildDocNumber_(brand, so, docType);
    docNumber = seqDocNum;
    baseName  = `${brand}–SO${so}–${seqDocNum}_${rp_fmtDateYMD_(now)}`;
  }

  const tmpl   = DriveApp.getFileById(tmplId);
  const doc    = tmpl.makeCopy(baseName, destFolder);
  const docId  = doc.getId();

  // Try table render first
  const lines = (payload.lines || []).slice(0, 5);
  const tableRes = rp_fillItemsTable_(docId, lines);

  const linesSubtotalNum = lines.reduce((s, ln) => s + Number(ln.qty||0)*Number(ln.amt||0), 0);
  const linesSubtotal = rp_money(linesSubtotalNum);
  const num = v => Number(String(v == null ? '' : v).replace(/[^\d.\-]/g, '')) || 0;

  // SO snapshot BEFORE
  let orderTotalBefore = 0, paidBefore = 0, balBefore = 0;
  if (ledgerMeta && (ledgerMeta.orderTotal != null || ledgerMeta.paidToDate != null || ledgerMeta.balance != null)) {
    orderTotalBefore = num(ledgerMeta.orderTotal);
    paidBefore       = num(ledgerMeta.paidToDate);
    balBefore        = num(ledgerMeta.balance);
  } else if (payload && payload.snapshots) {
    orderTotalBefore = num(payload.snapshots.orderTotal);
    paidBefore       = num(payload.snapshots.paidToDate);
    balBefore        = num(payload.snapshots.balance);
  }

  // Payment (this doc)
  const isReceipt   = /Receipt/i.test(docType);
  const pmt         = payload.pmt || {};
  const amountGross = isReceipt ? num(pmt.amount) : 0;
  const allocToSO   = isReceipt ? num(pmt.allocatedToSO) : 0;

  // For SO-anchored RECEIPTS, show after = BEFORE − allocToSO (clamped ≥0 for docs)
  const paidAfter    = isReceipt ? (paidBefore + allocToSO) : paidBefore;
  const balAfterRaw  = isReceipt ? Math.max(0, orderTotalBefore - paidAfter) : balBefore;

  // Previous payments (before this row)
  const prevLimit = Number(PropertiesService.getScriptProperties().getProperty('PREV_PAYMENTS_LIMIT') || 10);
  const prev = rp_prevPaymentsForAnchor_({ anchorType: payload.anchorType, rootApptId: payload.rootApptId, soNumber: payload.soNumber, beforeRow: (ledgerMeta && ledgerMeta.ledgerRow) || null, limit: prevLimit });
  const prevItems = (prev.items || []).map(it => ({ date: it.date || '', amount: Number(it.amount || 0), method: it.method || '' }));
  const prevSumNum = prevItems.reduce((s, it) => s + (Number(it.amount) || 0), 0);
  const hasPrev    = (prevItems.length > 0) && (prevSumNum > 0);
  const prevLabel  = hasPrev ? 'Previous Payments' : '';
  const prevBlock  = hasPrev ? rp_formatPaymentsList_(prevItems) : '';


// ✅ Always display an Order Total on the doc (fallback to Lines Subtotal when Orders/Master OT is blank)
const orderTotalForDocNum = (orderTotalBefore > 0 ? orderTotalBefore : linesSubtotalNum);


  const repl = {
    'DOC_DATE'     : Utilities.formatDate(now, RP_TZ, 'MMM d, yyyy'),
    'CUSTOMER_NAME': payload.customerName || '',
    'ROOT_APPT_ID' : payload.rootApptId || '',
    'SO_NUMBER'    : so || '',
    'DOC_NUMBER'   : docNumber || '',
    'PMT_ID'       : payload.pmtId || '',


    'LINES'         : tableRes.usedTable ? '' : rp_renderLinesText_(lines),
    'LINES_SUBTOTAL': linesSubtotal,
    ...(() => { const m={}; for(let i=0;i<5;i++){ const ln=lines[i]||{}; m['DESC_'+(i+1)]=(tableRes.usedTable? '' : (ln.desc||'')); m['QTY_'+(i+1)]=tableRes.usedTable? '' : (ln.qty!=null?String(ln.qty):''); m['AMT_'+(i+1)]=tableRes.usedTable? '' : (ln.amt!=null?rp_money(num(ln.amt)):''); } return m; })(),


    // Treat 0 as a real value; only blank when truly null/undefined
    'ORDER_TOTAL_SO'      : (orderTotalBefore || orderTotalBefore === 0) ? rp_money(orderTotalBefore) : '',
    'PAID_TO_DATE_BEFORE' : (paidBefore       || paidBefore       === 0) ? rp_money(paidBefore)       : '',
    'BALANCE_BEFORE'      : (balBefore        || balBefore        === 0) ? rp_money(balBefore)        : '',


    'PAID_TO_DATE_AFTER'  : isReceipt ? rp_money(paidAfter)    : '',
    'BALANCE_AFTER'       : isReceipt ? rp_money(balAfterRaw)  : '',
    'PAYMENT_AMOUNT'      : isReceipt ? rp_money(amountGross)  : '',


    'PAYMENT_METHOD'     : isReceipt ? (pmt.method || '') : '',
    'PAYMENT_REFERENCE'  : isReceipt ? (pmt.reference || '') : '',
    'PAYMENT_NOTES'      : isReceipt ? (pmt.notes || '') : '',


    'PREVIOUS_PAYMENTS_LABEL': prevLabel,
    'PREVIOUS_PAYMENTS_BLOCK': prevBlock
  };
  
  // Insert right after: const repl = { ... }  and before: rp_fillDocPlaceholders_(docId, repl);
  (function reconcileOrderTotalsAndBalance(){
    const isReceipt = /Receipt/i.test(docType);
    const isInvoice = /Invoice/i.test(docType);

    // 1) Base OT fallback to Lines Subtotal when blank on snapshots
    const baseOT = (orderTotalBefore > 0 ? orderTotalBefore : linesSubtotalNum);

    // 2) Paid‑to‑date fallback to summed prior receipts if snapshot is blank
    const paidBeforeForMath = (paidBefore > 0 ? paidBefore : prevSumNum);

    // 3) Payment applied: Gross on APPT, Allocated on SO
    const payApplied = isReceipt
      ? ((String(payload.anchorType || '').toUpperCase() === 'SO') ? allocToSO : amountGross)
      : 0;

    const paidAfterNum = isReceipt ? (paidBeforeForMath + payApplied) : paidBeforeForMath;
    const balBeforeNum = Math.max(0, baseOT - paidBeforeForMath);
    const balAfterNum  = isReceipt ? Math.max(0, baseOT - paidAfterNum) : balBeforeNum;

    // 4) Override replacements to use the robust math everywhere
    repl.ORDER_TOTAL_SO      = rp_money(baseOT);
    repl.ORDER_TOTAL         = repl.ORDER_TOTAL_SO;

    repl.PAID_TO_DATE_BEFORE = rp_money(paidBeforeForMath);
    repl.BALANCE_BEFORE      = rp_money(balBeforeNum);

    if (isReceipt) {
      repl.PAID_TO_DATE_AFTER = rp_money(paidAfterNum);
      repl.BALANCE_AFTER      = rp_money(balAfterNum);
    }
    repl.BALANCE = isReceipt ? (repl.BALANCE_AFTER || '') : (repl.BALANCE_BEFORE || '');

    // 5) Invoice‑only: expose requested amount for {{REQ_AMT}}
    if (isInvoice) {
      const reqAmt = num(pmt.amount);
      repl.REQ_AMT = rp_money(reqAmt);
    }

    // 6) If balance is zero, blank the value and delete that row only (not the whole table)
    const EPS = 0.005;
    const balNumToShow = isReceipt ? balAfterNum : balBeforeNum;
    if (balNumToShow <= EPS) {
      // Try placeholder key and common aliases that may appear in templates
      ['BALANCE', 'Remaining Balance', 'BALANCE DUE', 'Balance'].forEach(k => {
        try { rp_deletePlaceholderRowOrLine_(docId, k); } catch(_) {}
      });
      // Ensure any stray {{BALANCE}} usages print as empty
      repl.BALANCE = '';
      repl.BALANCE_BEFORE = '';
      repl.BALANCE_AFTER = '';
    }
  })();

  if (payload.anchorType === 'SO') {
    repl.BALANCE_SO = /Receipt/i.test(docType) ? (repl.BALANCE_AFTER || '') : (repl.BALANCE_BEFORE || '');
  }

    // --- SHOW "Previous Payments" row under Order Total only on DI/DR/SI (not SR) ---
    // Also show it when the ledger shows previous payments even if snapshots are blank.
    const codeInfo = rp_docCodeFromDocType_(docType); // { code:'DI'|'DR'|'SI'|'SR' }
    const paidToDateForRow = (paidBefore > 0) ? paidBefore : prevSumNum; // fall back to ledger sum
    const showPrevRowUnderOT = (codeInfo.code !== 'SR') && (paidToDateForRow > 0);

    // Convenience aliases for updated naming
    repl.ORDER_TOTAL  = repl.ORDER_TOTAL_SO || '';
    repl.Paid_to_date = showPrevRowUnderOT ? rp_money(paidToDateForRow) : '';
    repl.BALANCE      = repl.BALANCE_AFTER || repl.BALANCE_BEFORE || '';

  // Fill placeholders
  rp_fillDocPlaceholders_(docId, repl);


  // If the template doesn't provide a dedicated {{Paid_to_date}} line,
  // insert a "Previous Payments" row directly under "Order Total".
  if (showPrevRowUnderOT) {
    rp_insertPrevRowUnderOT_(docId, 'Previous Payments', rp_money(paidToDateForRow));
  }

  // Remove the single-line OT-adjacent "Previous Payments {{Paid_to_date}}" row when not showing
  if (!showPrevRowUnderOT) {
    // This deletes the entire line/paragraph that contains {{Paid_to_date}} in the template.
    rp_deletePlaceholderLine_(docId, 'Paid_to_date');
  }

  // The separate previous-payments block behavior is unchanged
  if (!hasPrev) {
    rp_deletePlaceholderLine_(docId, 'PREVIOUS_PAYMENTS_LABEL');
    rp_deletePlaceholderLine_(docId, 'PREVIOUS_PAYMENTS_BLOCK');
  }


  Utilities.sleep(200);
  const pdfBlob = DriveApp.getFileById(docId).getAs('application/pdf');
  pdfBlob.setName(baseName + '.pdf');
  const pdfFile = destFolder.createFile(pdfBlob);
  const pdfId   = pdfFile.getId();


  return {
    docId, pdfId, docNumber,
    docUrl: 'https://docs.google.com/document/d/' + docId + '/edit',
    pdfUrl: 'https://drive.google.com/file/d/' + pdfId + '/view'
  };
}
function rp_nextSequenceFor_(soNumber, isReceipt) {
  const { sh } = rp_getLedgerTarget();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return 1;
  const headers = sh.getRange(1,1,1,lc).getValues()[0].map(String);
  const map = {}; headers.forEach((h,i)=> map[h.trim()] = i);
  const soCol = map['SO#'], typeCol = map['DocType'];
  if (soCol == null || typeCol == null) return 1;
  const vals = sh.getRange(2,1,lr-1,lc).getValues();
  let count = 0;
  for (const row of vals) {
    if (rp_soEq(row[soCol], soNumber)) {
      const dt = String(row[typeCol]||'');
      const fam = /Receipt/i.test(dt) ? 'Receipt' : 'Invoice';
      if ((isReceipt && fam==='Receipt') || (!isReceipt && fam==='Invoice')) count++;
    }
  }
  return count + 1;
}
function rp_buildDocNumber_(brand, soNumber, docType) {
  const isReceipt = /Receipt/i.test(docType);
  const seq = rp_nextSequenceFor_(soNumber, isReceipt);
  return isReceipt ? `Receipt_PM-${seq}` : `Invoice_v${seq}`;
}
function rp_formatDocNumberForSO({ brand, so, docType, version, dateObj } = {}) {
  if (!brand || !so || !docType) throw new Error('Usage: rp_formatDocNumberForSO({brand, so, docType, version, dateObj})');
  const info = rp_docCodeFromDocType_(docType);
  const v = Math.max(1, Number(version||1));
  const d = dateObj || new Date();
  const ymd = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return `${String(brand).toUpperCase()}–SO${String(so).trim()}–${info.code}_v${v}_${ymd}`;
}

/*** === LEDGER UPDATE === ***/
function rp_updateLedgerRow_(row, updates) {
  const { sh } = rp_getLedgerTarget();
  const lc = sh.getLastColumn();
  const headers = sh.getRange(1,1,1,lc).getValues()[0].map(v => String(v).trim());
  const map = {}; headers.forEach((h,i)=> map[h]=i);
  const newHeaders = [];
  Object.keys(updates).forEach(h => { if (map[h]==null) newHeaders.push(h); });
  if (newHeaders.length) {
    let cursor = headers.length;
    newHeaders.forEach(h => { sh.getRange(1, cursor+1).setValue(h); map[h]=cursor; cursor++; });
  }
  const arr = sh.getRange(row,1,1,Math.max(...Object.values(map))+1).getValues()[0];
  Object.entries(updates).forEach(([h,v]) => { arr[map[h]] = v; });
  sh.getRange(row,1,1,arr.length).setValues([arr]);
}

/*** === LEDGER LOOKUPS (by Doc#) === ***/
function rp_findLedgerRowByDocNumber_(docNumber) {
  const { sh } = rp_getLedgerTarget();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || !docNumber) return null;

  const head = rp_getHeaderRowCached_(sh);
  const H = {}; head.forEach((h,i)=> H[h]=i);
  const cDoc = (H['DocNumber'] != null ? H['DocNumber'] : H['Doc #']);
  if (cDoc == null) return null;

  const start = rp_scanWindowStart_(lr);
  const vals = sh.getRange(start, 1, lr - start + 1, lc).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    const r = vals[i];
    if (String(r[cDoc] || '').trim() === String(docNumber || '').trim()) {
      return { row: start + i, rowVals: r, H, sh };
    }
  }
  return null;
}

/** Amount from a receipt row that was actually applied to 100_ PTD */
function rp_getAppliedAmtForMasterOnReceiptRow_(rowVals, H) {
  const anchor = String(rowVals[H['AnchorType']] || '').toUpperCase();
  const gross  = Number(rowVals[H['AmountGross']] || 0);
  const alloc  = (H['AllocatedToSO'] != null) ? Number(rowVals[H['AllocatedToSO']] || 0) : 0;
  if (anchor === 'SO') {
    // Respect property SO_RECEIPT_MASTER_AMOUNT (GROSS vs ALLOC) same as submit path
    return rp_amountForMasterOnSOReceipt_({ amount: gross, allocatedToSO: alloc });
  }
  return gross; // APPT receipts apply the gross to 100_
}

/** After new doc number exists, back-link the old row with ReplacedByDoc# */
function rp_linkSupersession_(oldDocNumber, newDocNumber){
  try {
    if (!oldDocNumber || !newDocNumber) return;
    const hit = rp_findLedgerRowByDocNumber_(oldDocNumber);
    if (!hit) return;
    rp_updateLedgerRow_(hit.row, { 'ReplacedByDoc#': newDocNumber });
  } catch (_) {}
}


/*** === FOLDERS: resolve destination & PaymentsFolderURL === **/
function rp_ensureChildFolder_(parent, name){ const it=parent.getFoldersByName(name); return it.hasNext()?it.next():parent.createFolder(name); }
function rp_getOrdersTabName_() { const p = PropertiesService.getScriptProperties(); return p.getProperty('301/302_TAB_NAME') || p.getProperty('ORDERS_TAB_NAME') || '1. Sales'; }
function rp_findSoRowInBrand_(brand, soNumber) {
  if (!brand || !soNumber) return null;
  const entry = (function(){
    const props = PropertiesService.getScriptProperties();
    const hp = props.getProperty('HPUSA_301_FILE_ID') || props.getProperty('HPUSA_ORDERS_FILE_ID') || props.getProperty('CFG_HPUSA_ORDERS_FILE_ID') || '';
    const vvs = props.getProperty('VVS_302_FILE_ID')  || props.getProperty('VVS_ORDERS_FILE_ID')  || props.getProperty('CFG_VVS_ORDERS_FILE_ID') || '';
    if (String(brand).toUpperCase().includes('HPUSA') && hp) return { brand:'HPUSA', fileId:hp };
    if (String(brand).toUpperCase().includes('VVS')   && vvs) return { brand:'VVS',   fileId:vvs };
    return null;
  })();
  if (!entry) return null;
  const ss = SpreadsheetApp.openById(entry.fileId);
  const sh = ss.getSheetByName(rp_getOrdersTabName_()) || ss.getSheetByName('1. Sales') || ss.getSheets()[0];
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  const values = sh.getRange(1,1,lr,lc).getValues();
  const map = rp_headerMap(values);
  if (map['SO#'] == null) return null;
  for (let r = 1; r < values.length; r++) { if (rp_soEq(values[r][map['SO#']], soNumber)) return { sh, rowIndex: r+1, map, rowVals: values[r] }; }
  return null;
}


/**
* Resolve destination folder and write PaymentsFolderURL when blank.
*  - APPT: dest = Client Folder › 04-Deposit; write 100_.PaymentsFolderURL if blank.
*  - SO:   dest = Orders root (VVS/HPUSA) › SO{#} › 04-Deposit; write PaymentsFolderURL (100_ & 301/302) if blank to that Orders path.
*/
function rp_resolveDestAndClientFolders_(payload) {
  const wrote = { master:false, orders:false };
  // Fast-path: if caller passed a Payments Folder URL (from prefill), trust it.
  // This keeps doc-gen working even when 100_/301/302 pointers are momentarily blank.
  try {
    if (payload && payload.paymentsFolderURL) {
      const fId = rp_fileIdFromUrl(payload.paymentsFolderURL);
      if (fId) {
        const f = DriveApp.getFolderById(fId);
        if (f) return { dest: f, paymentsFolderUrl: f.getUrl(), wrote };
      }
    }
  } catch (_) { /* fall through to canonical resolution */ }

  if (!payload || !payload.rootApptId) throw new Error('Bad payload for folder resolution.');
  const m = payload.masterRowIndex ? rp_getMasterRowByIndex_(payload.masterRowIndex) : rp_findMasterRowByRootApptId_(payload.rootApptId);
  if (!m) throw new Error('Master row not found for RootApptID ' + payload.rootApptId);

  const clientFolderIdx = rp_pick0(m.map, 'Client Folder', 'ClientFolderURL', 'Customer Folder');
  if (clientFolderIdx < 0) throw new Error('Missing "Client Folder" column on 100_.');
  const clientFolderUrl = String(m.rowVals[clientFolderIdx] || '').trim();
  if (!clientFolderUrl) throw new Error('Client Folder URL is blank on this row.');
  const clientFolder = DriveApp.getFolderById(rp_fileIdFromUrl(clientFolderUrl));

  // APPT‑anchored
  if (String(payload.anchorType).toUpperCase() !== 'SO' || !payload.soNumber) {
    const clientPaymentsFolder = rp_ensureChildFolder_(clientFolder, '04-Deposit');
    const pfIdxM = rp_pick0(m.map, 'PaymentsFolderURL');
    if (pfIdxM >= 0) {
      const curM = String(m.rowVals[pfIdxM] || '').trim();
      if (!curM) { m.sh.getRange(m.rowIndex, pfIdxM + 1).setValue(clientPaymentsFolder.getUrl()); wrote.master = true; }
    }
    return { dest: clientPaymentsFolder, paymentsFolderUrl: clientPaymentsFolder.getUrl(), wrote };
  }

  // SO‑anchored: build folder under brand Orders root using properties + 100_ only (no 301/302 I/O)
  const brandNorm = String(payload.brand || '').toUpperCase().includes('VVS') ? 'VVS' : 'HPUSA';
  const rootId = (brandNorm === 'VVS'
    ? rp_propOneOf_(RP_KEY_ALIASES.VVS_SO_ROOT_FOLDER_ID).value
    : rp_propOneOf_(RP_KEY_ALIASES.HPUSA_SO_ROOT_FOLDER_ID).value) || '';
  if (!rootId) throw new Error('SO root not configured in Script Properties.');
  const root = DriveApp.getFolderById(rootId);
  const so = String(payload.soNumber || '').trim();
  // Short Tag from 100_ only
  const stIdxM = rp_pick0(m.map, 'Short Tag','ShortTag','SO Short Tag','SO Tag');
  const shortTag = stIdxM >= 0 ? rp_sanitizeForFolder_(String(m.rowVals[stIdxM] || '')) : '';
  const folderLabel = [brandNorm, `SO${so}`, shortTag].filter(Boolean).join('-');
  const soFolder = (function(){ const it = root.getFoldersByName(folderLabel); return it.hasNext()? it.next() : root.createFolder(folderLabel); })();
  const paymentsFolder = rp_ensureChildFolder_(soFolder, '04-Deposit');
  // Write PaymentsFolderURL to 100_ only (if blank)
  const pfIdxM = rp_pick0(m.map, 'PaymentsFolderURL');
  if (pfIdxM >= 0) {
    const curM = String(m.rowVals[pfIdxM] || '').trim();
    if (!curM) { m.sh.getRange(m.rowIndex, pfIdxM + 1).setValue(paymentsFolder.getUrl()); wrote.master = true; }
  }
  return { dest: paymentsFolder, paymentsFolderUrl: paymentsFolder.getUrl(), wrote };

}



/*** === DOC GENERATION + AR SHORTCUT === ***/
function rp_makeDocForPayment(ledgerRow, payload) {
  try {
    if (!payload || !payload.docType) return { ok:false, reason:'BAD_PAYLOAD', hint:'Missing payload or docType' };
    var docType = String(payload.docType);
    var anchorType = String(payload.anchorType || '');
    var brand = String(payload.brand || '').trim();

    // Attach PAYMENT_ID to payload for {{PMT_ID}}
    try {
      const { sh } = rp_getLedgerTarget();
      const lc = sh.getLastColumn();
      const head = sh.getRange(1,1,1,lc).getValues()[0].map(v=>String(v).trim());
      const H = {}; head.forEach((h,i)=> H[h]=i);
      const rowVals = sh.getRange(ledgerRow, 1, 1, lc).getValues()[0];
      payload.pmtId = String(rowVals[H['PAYMENT_ID']] || '');
      // Fallback: if brand is missing in payload (e.g., APPT-anchored row without Brand on 100_),
      // read it from the saved ledger row to avoid template lookup failures.
      if (!brand) {
        try { brand = String(rowVals[H['Brand']] || '').trim(); } catch (_){}
      }

    } catch(_){}

    var resolved = rp_resolveDestAndClientFolders_(payload);
    var destFolder = resolved.dest;

    if (!brand && anchorType === 'APPT') { try { brand = rp_brandFromMaster_(payload.rootApptId) || ''; } catch(_) {} }
    if (!brand) { return { ok:false, reason:'BRAND_NOT_FOUND_ON_MASTER', hint:'Brand is required to pick template' }; }

    var meta = {
      orderTotal: (payload.snapshots && payload.snapshots.orderTotal) || '',
      paidToDate: (payload.snapshots && payload.snapshots.paidToDate) || '',
      balance:    (payload.snapshots && payload.snapshots.balance)    || '',
      ledgerRow:  ledgerRow
    };

    var out = rp_generateDocAndPdf_(brand, docType, destFolder, payload, meta);
    if (!out || !out.docId || !out.pdfId) return { ok:false, reason:'DOC_GEN_RETURN_INVALID', hint:'Missing docId/pdfId from generator' };

    // Optional rename for APPT
    if (anchorType === 'APPT') {
      try {
        var ver = rp_nextDocVersion_('APPT', payload.rootApptId, '', docType);
        var shortBase = rp_makeApptFilename_(brand, payload.rootApptId, docType, ver, new Date());
        DriveApp.getFileById(out.docId).setName(shortBase);
        DriveApp.getFileById(out.pdfId).setName(shortBase + '.pdf');
        out.docUrl = 'https://docs.google.com/document/d/' + out.docId + '/edit';
        out.pdfUrl = 'https://drive.google.com/file/d/' + out.pdfId + '/view';
      } catch (e) { Logger.log('APPT rename failed: ' + ((e && e.message) ? e.message : e)); }
    }

    // Update ledger with doc fields
    try { rp_updateLedgerRow_(ledgerRow, { 'DocNumber': out.docNumber || '', 'DocFileID': out.docId || '', 'DocPDFID': out.pdfId || '', 'DocURL': out.docUrl || '', 'PDFURL': out.pdfUrl || '' }); }
    catch (e) { Logger.log('Ledger update (doc fields) failed: ' + ((e && e.message) ? e.message : e)); }

    // If this doc supersedes another, back-link old row to this Doc#
    try {
      if (payload && payload.supersedes && out && out.docNumber) {
        rp_linkSupersession_(payload.supersedes, out.docNumber);
      }
    } catch (e) { Logger.log('Supersession back-link warning: ' + ((e && e.message) ? e.message : e)); }

    // AR monthly shortcut (brand‑based: VVS→20_AR, HPUSA→21_AR), PDF only
    var arShortcutURL = '';
    try {
      var arMonthly = rp_ensureArMonthlyFolder_(brand, new Date());
      var ar = null;
      if (arMonthly) { ar = rp_createDriveShortcut_(arMonthly.getId(), out.pdfId, (out.docNumber || 'Doc') + '.pdf'); arShortcutURL = (ar && ar.url) || ''; }
      rp_updateLedgerRow_(ledgerRow, { 'ARShortcutID': (ar && ar.id) || '', 'ARShortcutURL': arShortcutURL });
    } catch (e) { Logger.log('AR shortcut error: ' + ((e && e.message) ? e.message : e)); }

    return { ok: true, row: ledgerRow, brand: brand, docNumber: out.docNumber || '', docId: out.docId, pdfId: out.pdfId, docUrl: out.docUrl, pdfUrl: out.pdfUrl, paymentsFolderURL: resolved.paymentsFolderUrl, wrote: resolved.wrote, arShortcutURL: arShortcutURL };
  } catch (e) {
    return { ok:false, reason:'UNCAUGHT', hint:(e && e.message) ? e.message : String(e) };
  }
}

/*** === ORDER / MASTER HELPERS === ***/
function rp_brandFromMaster_(rootApptId) {
  const m = rp_findMasterRowByRootApptId_(rootApptId);
  if (!m) return '';
  const idx = m.map['Brand'];
  return idx != null ? String(m.rowVals[idx] || '').trim() : '';
}
function rp_makeApptFilename_(brand, rootApptId, docType, version, when) {
  const d = when || new Date();
  const yyyy = d.getFullYear(), mm = String(d.getMonth()+1).padStart(2,'0'), dd = String(d.getDate()).padStart(2,'0');
  const info = rp_docCodeFromDocType_(docType);
  const safeBrand = String(brand || '').trim() || 'Brand';
  const ridRaw = String(rootApptId || '').trim();
  const ridPart = ridRaw.replace(/^A+/, '');
  const rid = `A${ridPart}`;
  const v = (version && version > 0) ? version : 1;
  return `${safeBrand}–${rid}–${info.code}_v${v}_${yyyy}-${mm}-${dd}`;
}
function rp_nextDocVersion_(anchorType, rootApptId, soNumber, docType) {
  const { sh } = rp_getLedgerTarget();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return 1;
  const values = sh.getRange(2,1,lr-1,lc).getValues();
  const head = rp_headerMap(sh.getRange(1,1,1,lc).getValues());
  let n = 0;
  for (const row of values) {
    const a = String(row[head['AnchorType']] || '');
    const d = String(row[head['DocType']] || '');
    const rid = String(row[head['RootApptID']] || '');
    const so  = String(row[head['SO#']] || '');
    if (a === anchorType && d === docType && (a==='APPT' ? (rid===String(rootApptId)) : (so===String(soNumber)))) { n++; }
  }
  return n+1;
}

function rp_createDriveShortcut_(parentFolderId, targetFileId, title) {
  // Verify the target exists and is visible across My Drive & Shared Drives.
  // A very short retry handles "just-created" eventual consistency.
  function assertTargetVisible() {
    try {
      Drive.Files.get(targetFileId, {
        supportsAllDrives: true,
        supportsTeamDrives: true,
        fields: 'id'
      });
    } catch (e) {
      Utilities.sleep(400);
      // Second attempt
      Drive.Files.get(targetFileId, {
        supportsAllDrives: true,
        supportsTeamDrives: true,
        fields: 'id'
      });
    }
  }
  
  assertTargetVisible();

  var resource = {
    title: title,
    mimeType: 'application/vnd.google-apps.shortcut',
    parents: [{ id: parentFolderId }],
    shortcutDetails: {
      targetId: targetFileId,
      // Supplying the target mime type can help certain Drive surfaces.
      targetMimeType: 'application/pdf'
    }
  };

  var file = Drive.Files.insert(
    resource,
    null,
    {
      supportsAllDrives: true,
      supportsTeamDrives: true,
      fields: 'id'
    }
  );

  return { id: file.id, url: 'https://drive.google.com/file/d/' + file.id + '/view' };
}

function rp_getMasterRowByIndex_(rowIndex) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RP_MASTER_SHEET);
  if (!sh) throw new Error(`Missing sheet "${RP_MASTER_SHEET}"`);


  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (rowIndex < 2 || rowIndex > lr) throw new Error('Invalid master row index: ' + rowIndex);


  // Read header + the one row we need
  const header = sh.getRange(1,1,1,lc).getDisplayValues();
  const map = rp_headerMap(header);
  const rowVals = sh.getRange(rowIndex,1,1,lc).getDisplayValues()[0];


  // Robust header lookups (alias‑tolerant)
  const apptIdx = rp_pick0(map, 'APPT_ID','RootApptID','Root Appt ID');
  const custIdx = rp_pick0(map, 'Customer Name','Customer','Client Name','Client');
  const soIdx   = rp_pick0(map, 'SO#','SO','SO Number','Sales Order','Sales Order #');
  const trkIdx  = rp_pick0(map, '3D Tracker','3D Log');


  if (apptIdx < 0) throw new Error('Missing "APPT_ID"/RootApptID column on ' + RP_MASTER_SHEET);
  if (custIdx < 0) throw new Error('Missing "Customer Name"/Customer column on ' + RP_MASTER_SHEET);


  // Resolve tracker URL (including rich‑text link)
  let trackerUrl = '';
  if (trkIdx >= 0) {
    trackerUrl = String(rowVals[trkIdx] || '').trim();
    if (!trackerUrl) {
      try {
        const rich = sh.getRange(rowIndex, trkIdx + 1).getRichTextValue();
        if (rich) {
          trackerUrl = rich.getLinkUrl() || '';
          if (!trackerUrl && rich.getRuns) {
            const runs = rich.getRuns();
            for (let i = 0; i < runs.length; i++) {
              const u = runs[i].getLinkUrl && runs[i].getLinkUrl();
              if (u) { trackerUrl = u; break; }
            }
          }
        }
      } catch (_) {}
    }
  }
  // Return the SAME shape as rp_activeMasterRow()
  return {
    rowIndex,
    rootApptId: String(rowVals[apptIdx] || '').trim(),
    customerName: String(rowVals[custIdx] || '').trim(),
    soNumber: String((soIdx >= 0 ? rowVals[soIdx] : '') || '').trim(),
    trackerUrl,
    map,
    rowVals,
    sh
  };
}


function rp_findMasterRowByRootApptId_(rootApptId) {
  if (!rootApptId) return null;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RP_MASTER_SHEET);
  if (!sh) throw new Error(`Missing sheet "${RP_MASTER_SHEET}"`);
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return null;
  const values = sh.getRange(1,1,lr,lc).getDisplayValues();
  const map = rp_headerMap(values);
  const idx = (map['RootApptID'] != null ? map['RootApptID'] : (map['Root Appt ID'] != null ? map['Root Appt ID'] : map['APPT_ID']));
  if (idx == null) throw new Error('Missing RootApptID / Root Appt ID / APPT_ID header on ' + RP_MASTER_SHEET);
  for (let r=1; r<values.length; r++){
    if (String(values[r][idx] || '').trim() === String(rootApptId).trim()) return { sh, rowIndex:r+1, map, rowVals: values[r] };
  }
  return null;
}

function rp_setOrderTotal_SO_(brand, soNumber, amount, allowOverride) {
  const soRow = rp_findSoRowInBrand_(brand, soNumber);
  if (!soRow) return { ok:false, reason:'SO_ROW_NOT_FOUND' };
  const idxOT = soRow.map['Order Total'];
  if (idxOT == null) return { ok:false, reason:'ORDER_TOTAL_COL_MISSING' };

  const cell = soRow.sh.getRange(soRow.rowIndex, idxOT + 1);
  const prev = Number(cell.getValue() || 0);
  const val  = Math.round(Number(amount || 0) * 100) / 100;

  if (!(val > 0)) return { ok:false, reason:'AMOUNT_NOT_POSITIVE', prev };
  if (prev > 0 && !allowOverride) return { ok:true, updated:false, prev, value:prev };

  cell.setValue(val);

  // Also recompute Remaining Balance on 301/302: RB = max(0, OT − PTD)
  try {
    const cPTD = rp_pick0(soRow.map, 'Paid-to-Date','Paid-To-Date','Paid to Date','Paid-to-date');
    const cBAL = rp_pick0(soRow.map, 'Remaining Balance','Balance');
    if (cPTD >= 0 && cBAL >= 0) {
      const ptd = rp_num_(soRow.sh.getRange(soRow.rowIndex, cPTD + 1).getValue());
      const rb  = Math.max(0, val - ptd);
      soRow.sh.getRange(soRow.rowIndex, cBAL + 1).setValue(rb);
    }
  } catch (_) {}
  return { ok:true, updated:true, prev, value:val };
}


function rp_setOrderTotal_Master_Safe_(rootApptId, value, allowOverride, masterRowIndex) {
  try {
    var val = Math.round(Number(value || 0) * 100) / 100;
    if (!(val > 0)) return { ok:false, updated:false, reason:'AMOUNT_NOT_POSITIVE' };


    // --- Resolve the target Master row (same logic/shape as before) ---
    var m;
    if (masterRowIndex) {
      m = rp_getMasterRowByIndex_(masterRowIndex);
    } else {
      m = rp_findMasterRowByRootApptId_(rootApptId);
      if (!m) return { ok:false, updated:false, reason:'MASTER_ROW_NOT_FOUND' };
    }


    // Exact same header we require for Master OT
    var cOT0 = m.map['Order Total'];
    if (cOT0 == null) return { ok:false, updated:false, reason:'ORDER_TOTAL_HEADER_MISSING' };


    // Optional columns for RB math
    var cPTD0 = rp_pick0(m.map, 'Paid-to-Date','Paid-To-Date','Paid to Date','Paid-to-date'); // 0-based
    var cRB0  = rp_pick0(m.map, 'Remaining Balance','Balance');                               // 0-based


    // --- Minimal read: fetch OT (prev) and PTD in a single block when PTD exists ---
    var prev, ptd = 0;
    if (cPTD0 >= 0) {
      var readMin  = Math.min(cOT0, cPTD0) + 1;                // 1-based col
      var readSpan = Math.max(cOT0, cPTD0) - Math.min(cOT0, cPTD0) + 1;
      var blk      = m.sh.getRange(m.rowIndex, readMin, 1, readSpan).getValues()[0];
      prev = blk[(cOT0 + 1)  - readMin];
      ptd  = rp_num_(blk[(cPTD0 + 1) - readMin]);
    } else {
      prev = m.sh.getRange(m.rowIndex, cOT0 + 1).getValue();
    }


    // Same override rule as before
    if (prev && !allowOverride) return { ok:true, updated:false, value:prev, prev:prev };


    // --- Prepare grouped writes: always OT; RB only if (PTD & RB exist) ---
    var targets = [{ col: cOT0 + 1, val: val }]; // 1-based
    if (cPTD0 >= 0 && cRB0 >= 0) {
      var newRB = Math.max(0, val - ptd);
      targets.push({ col: cRB0 + 1, val: newRB });
    }


    // Group contiguous targets → minimal setValues() calls
    targets.sort(function(a,b){ return a.col - b.col; });
    var runs = [], cur = null;
    for (var i = 0; i < targets.length; i++) {
      var t = targets[i];
      if (!cur) {
        cur = { start: t.col, vals: [t.val] };
      } else if (t.col === cur.start + cur.vals.length) {
        cur.vals.push(t.val);
      } else {
        runs.push(cur);
        cur = { start: t.col, vals: [t.val] };
      }
    }
    if (cur) runs.push(cur);


    // Execute writes
    for (var r = 0; r < runs.length; r++) {
      var run = runs[r];
      m.sh.getRange(m.rowIndex, run.start, 1, run.vals.length).setValues([run.vals]);
    }


    return { ok:true, updated:true, value:val, prev:prev };


  } catch (e) {
    return { ok:false, updated:false, reason: (e && e.message) ? e.message : String(e) };
  }
}


function rp_auditOrderTotalOnLedger_(row, info) {
  rp_updateLedgerRow_(row, {
    'OrderTotalSet': !!(info && info.set),
    'OrderTotalValue': (info && info.value != null) ? info.value : '',
    'OrderTotalSource': (info && info.source) || '',
    'OrderTotalTarget': (info && info.target) || '',
    'OrderTotalOldValue': (info && info.prev != null) ? info.prev : ''
  });
}


/*** === AR HELPERS === ***/
// Brand-based AR top: VVS → 20_AR, HPUSA → 21_AR
function rp_getArBrandRootId_(brand) {
  const isVVS = String(brand || '').toUpperCase().includes('VVS');
  if (isVVS)  return rp_propOneOf_(RP_KEY_ALIASES.AR_VVS_ROOT_ID,  { required:false, label:'AR VVS Root' }).value || '';
  else        return rp_propOneOf_(RP_KEY_ALIASES.AR_HPUSA_ROOT_ID,{ required:false, label:'AR HPUSA Root' }).value || '';
}
function rp_ensureArMonthlyFolder_(brand, when) {
  const rootId = rp_getArBrandRootId_(brand);
  if (!rootId) return null;
  const root = DriveApp.getFolderById(rootId);
  const topName = String(brand || '').toUpperCase().includes('VVS') ? '20_AR' : '21_AR';
  const yyyy = String((when || new Date()).getFullYear());
  const mm = String((when || new Date()).getMonth() + 1).padStart(2,'0');
  function ensure(parent, name){ const it=parent.getFoldersByName(name); return it.hasNext()?it.next():parent.createFolder(name); }
  const f1 = ensure(root, topName);
  const f2 = ensure(f1, yyyy);
  const f3 = ensure(f2, mm);
  return f3;
}


/*** === MASTER / ORDERS WRITE-BACKS === **/
/*** Persist Saved Lines JSON/Subtotal to 100_ and 301/302 ***/
function rp_persistSavedLinesToMaster_({ masterRowIndex, rootApptId, lines, subtotal } = {}) {
  const m = masterRowIndex ? rp_getMasterRowByIndex_(masterRowIndex) : rp_findMasterRowByRootApptId_(rootApptId);
  if (!m) throw new Error('Master row not found.');
  const sh = m.sh;
  let header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  let H = rp_hIndex_(header);
  let cJSON = rp_pick(H,'Saved Lines JSON','SavedLinesJSON');
  if (!cJSON) { sh.getRange(1, sh.getLastColumn()+1).setValue('Saved Lines JSON'); header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]; H = rp_hIndex_(header); cJSON = H['Saved Lines JSON']; }
  let cSub  = rp_pick(H,'Saved Subtotal','SavedSubtotal');
  if (!cSub) { sh.getRange(1, sh.getLastColumn()+1).setValue('Saved Subtotal'); header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]; H = rp_hIndex_(header); cSub = H['Saved Subtotal']; }
  const sanitized = (lines||[]).map(ln => ({ desc: rp_sanitizeDesc_(ln.desc), qty: Number(ln.qty)||0, amt: Number(ln.amt)||0 }));
  sh.getRange(m.rowIndex, cJSON).setValue(JSON.stringify(sanitized));
  sh.getRange(m.rowIndex, cSub).setValue(Math.round(Number(subtotal||0)*100)/100);
  return { ok:true };
}








function rp_persistSavedLinesToOrders_({ brand, soNumber, lines, subtotal } = {}) {
  const hit = rp_findSoRowInBrand_(brand, soNumber);
  if (!hit) return { ok:false, reason:'SO_NOT_FOUND' };
  const sh = hit.sh;
  let header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  let H = rp_hIndex_(header);
  let cJSON = rp_pick(H,'Saved Lines JSON','SavedLinesJSON');
  if (!cJSON) { sh.getRange(1, sh.getLastColumn()+1).setValue('Saved Lines JSON'); header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]; H = rp_hIndex_(header); cJSON = H['Saved Lines JSON']; }
  let cSub  = rp_pick(H,'Saved Subtotal','SavedSubtotal');
  if (!cSub) { sh.getRange(1, sh.getLastColumn()+1).setValue('Saved Subtotal'); header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]; H = rp_hIndex_(header); cSub = H['Saved Subtotal']; }
  const sanitized = (lines||[]).map(ln => ({ desc: rp_sanitizeDesc_(ln.desc), qty: Number(ln.qty)||0, amt: Number(ln.amt)||0 }));
  sh.getRange(hit.rowIndex, cJSON).setValue(JSON.stringify(sanitized));
  sh.getRange(hit.rowIndex, cSub).setValue(Math.round(Number(subtotal||0)*100)/100);
  return { ok:true };
}




function rp_applyReceiptToOrders({ brand, so, amount, when } = {}) {
  if (!brand || !so || !(amount > 0))
    throw new Error('Usage: rp_applyReceiptToOrders({brand:"VVS|HPUSA", so:"...", amount:100, when:new Date()})');


  const hit = rp_findSoRowInBrand_(brand, so);
  if (!hit) throw new Error('SO row not found for brand ' + brand + ', SO ' + so);


  const { sh, rowIndex, map } = hit;


  // Column indexes (0-based within the sheet header map)
  const cPTD = rp_pick0(map, 'Paid-to-Date','Paid-To-Date','Paid to Date','Paid-to-date');
  const cOT  = rp_pick0(map, 'Order Total','Order Total ');
  const cBAL = rp_pick0(map, 'Remaining Balance','Balance');
  const cLPD = rp_pick0(map, 'Last Payment Date','LastPaymentDate');
  if (cPTD < 0 || cOT < 0 || cBAL < 0 || cLPD < 0)
    throw new Error('Orders sheet missing PTD/OT/BAL/LPD headers');


  // ---- Single minimal read covering OT and PTD (never writes to OT) ----
  const readMin = Math.min(cPTD, cOT) + 1;                   // 1-based column
  const readMax = Math.max(cPTD, cOT) + 1;
  const readSpan = readMax - readMin + 1;
  const block = sh.getRange(rowIndex, readMin, 1, readSpan).getValues()[0];


  // Map 0-based header col → index inside the read block
  const fromBlock = (col0) => block[(col0 + 1) - readMin];


  const orderTotal = rp_num_(fromBlock(cOT));
  const paidToDate = rp_num_(fromBlock(cPTD));
  const newPaid = paidToDate + rp_num_(amount);
  const newBal  = Math.max(0, orderTotal - newPaid);
  const whenVal = when || new Date();


  // ---- Group the three write targets into contiguous runs (avoid touching OT) ----
  const targets = [
    { col: cPTD + 1, val: newPaid },      // 1-based columns
    { col: cBAL + 1, val: newBal },
    { col: cLPD + 1, val: whenVal }
  ].sort((a, b) => a.col - b.col);


  const runs = [];
  let cur = null;
  for (const t of targets) {
    if (!cur) {
      cur = { start: t.col, vals: [t.val] };
    } else if (t.col === cur.start + cur.vals.length) {
      cur.vals.push(t.val);
    } else {
      runs.push(cur);
      cur = { start: t.col, vals: [t.val] };
    }
  }
  if (cur) runs.push(cur);


  // Write each contiguous run with a single setValues()
  runs.forEach(r => {
    sh.getRange(rowIndex, r.start, 1, r.vals.length).setValues([r.vals]);
  });


  return { ok: true, row: rowIndex, newPaid, newBal };
}


function rp_applyReceiptToMaster({ masterRowIndex, amount, when } = {}) {
  if (!masterRowIndex || masterRowIndex < 2 || !amount)
    throw new Error('Usage: rp_applyReceiptToMaster({masterRowIndex: <row>, amount: 50, when:new Date()})');


  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet "00_Master Appointments".');


  // --- Ensure headers exist (PTD & LPD). Keep same behavior as before. ---
  let header = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0] || [];
  let H = rp_hIndex_(header);


  let cPTD = rp_pick(H, 'Paid-to-Date', 'Paid-To-Date', 'Paid to Date', 'Paid-to-date'); // 1-based col
  if (!cPTD) {
    sh.getRange(1, sh.getLastColumn() + 1).setValue('Paid-to-Date');
    header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    H = rp_hIndex_(header);
    cPTD = H['Paid-to-Date'];
  }


  let cLPD = rp_pick(H, 'Last Payment Date', 'LastPaymentDate');
  if (!cLPD) {
    sh.getRange(1, sh.getLastColumn() + 1).setValue('Last Payment Date');
    header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    H = rp_hIndex_(header);
    cLPD = H['Last Payment Date'];
  }


  // After possible header inserts, resolve optional OT/RB (same semantics as before)
  const H2 = rp_hIndex_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  const cRB = rp_pick(H2, 'Remaining Balance', 'Balance') || 0;         // 1-based (0 if missing)
  const cOT = rp_pick(H2, 'Order Total', 'Order Total ') || 0;


  const row = masterRowIndex;


  // --- Minimal reads: PTD (always), OT (only if present) ---
  let paidToDate = 0, orderTotal = 0;
  if (cOT) {
    const readMin  = Math.min(cPTD, cOT);
    const readSpan = Math.max(cPTD, cOT) - readMin + 1;
    const block    = sh.getRange(row, readMin, 1, readSpan).getValues()[0];
    paidToDate = rp_num_(block[cPTD - readMin]);
    orderTotal = rp_num_(block[cOT  - readMin]);
  } else {
    paidToDate = rp_num_(sh.getRange(row, cPTD).getValue());
  }


  const newPaid = paidToDate + rp_num_(amount);
  const whenVal = when || new Date();


  // --- Prepare targets: PTD + LPD (always). RB only if OT & RB exist. ---
  const targets = [
    { col: cPTD, val: newPaid },
    { col: cLPD, val: whenVal }
  ];


  let newBal;
  if (cOT && cRB) {
    newBal = Math.max(0, orderTotal - newPaid);
    targets.push({ col: cRB, val: newBal });
  }


  // --- Group contiguous columns into runs to minimize setValues() calls ---
  targets.sort((a, b) => a.col - b.col);
  const runs = [];
  let cur = null;
  for (const t of targets) {
    if (!cur) {
      cur = { start: t.col, vals: [t.val] };
    } else if (t.col === cur.start + cur.vals.length) {
      cur.vals.push(t.val);
    } else {
      runs.push(cur);
      cur = { start: t.col, vals: [t.val] };
    }
  }
  if (cur) runs.push(cur);


  runs.forEach(r => {
    sh.getRange(row, r.start, 1, r.vals.length).setValues([r.vals]);
  });


  if (cOT && cRB) return { ok: true, row, newPaid, newBal };
  return { ok: true, row, newPaid };
}


function rp_calcGrossCashInForAppt_(rootApptId) {
  if (!rootApptId) return 0;
  const { sh } = rp_getLedgerTarget();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return 0;
  const header = rp_getHeaderRowCached_(sh);
  const H = {}; header.forEach((h,i)=> H[h]=i);
  const cAppt = H['RootApptID'], cType = H['DocType'], cGross = H['AmountGross'];
  const cStatus = (H['DocStatus'] != null ? H['DocStatus'] : H['Status']);
  if (cAppt == null || cType == null || cGross == null) return 0;

  const start = rp_scanWindowStart_(lr);
  const vals = sh.getRange(start,1,lr-start+1,lc).getValues();
  let sum = 0;
  for (const r of vals) {
    const ap = String(r[cAppt]||'').trim();
    if (ap !== rootApptId) continue;
    const t  = String(r[cType]||'').toUpperCase();
    if (!(t.includes('RECEIPT') || t === 'DR' || t === 'SR')) continue;

    const status = cStatus != null ? String(r[cStatus] || '').toUpperCase().trim() : '';
    if (status === 'VOID' || status === 'REPLACED' || status === 'DRAFT') continue;

    sum += Number(r[cGross]||0);
  }
  return Math.round(sum * 100) / 100;
}

function rp_updateMasterCashInGross_({ masterRowIndex, rootApptId } = {}) {
  if (!rootApptId || !masterRowIndex) return 0;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet "00_Master Appointments".');
  const gross = rp_calcGrossCashInForAppt_(rootApptId);
  const header = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0] || [];
  const H = rp_hIndex_(header);
  let cGross = rp_pick(H, 'Cash-in (Gross)');
  if (!cGross) {
    sh.getRange(1, sh.getLastColumn()+1).setValue('Cash-in (Gross)');
    const h2 = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    cGross = rp_hIndex_(h2)['Cash-in (Gross)'];
  }
  sh.getRange(masterRowIndex, cGross).setValue(gross);
  return gross;
}
function rp_countReceiptsForAppt_(rootApptId) {
  if (!rootApptId) return 0;
  const { sh } = rp_getLedgerTarget();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return 0;
  const headers = rp_getHeaderRowCached_(sh);
  const H = {}; headers.forEach((h,i)=> H[h]=i);
  const cAppt = H['RootApptID'], cType = H['DocType'];
  const cStatus = (H['DocStatus'] != null ? H['DocStatus'] : H['Status']);
  if (cAppt == null || cType == null) return 0;

  const start = rp_scanWindowStart_(lr);
  const vals = sh.getRange(start,1,lr-start+1,lc).getValues();
  let n = 0;
  for (const r of vals) {
    if (String(r[cAppt]||'').trim() !== String(rootApptId).trim()) continue;
    const t = String(r[cType]||'').toUpperCase();
    if (!(t.includes('RECEIPT') || t === 'DR' || t === 'SR')) continue;

    const status = cStatus != null ? String(r[cStatus] || '').toUpperCase().trim() : '';
    if (status === 'VOID' || status === 'REPLACED' || status === 'DRAFT') continue;

    n++;
  }
  return n;
}

function rp_setSalesStageOnMaster_({ masterRowIndex, value, allowOverride } = {}) {
  if (!masterRowIndex || masterRowIndex < 2) return { ok:false, reason:'BAD_ROW' };
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) return { ok:false, reason:'MISSING_MASTER' };
  const lc = sh.getLastColumn();
  let header = sh.getRange(1,1,1,lc).getValues()[0] || [];
  let H = (function(hdr){ const m={}; hdr.forEach((h,i)=>{ const k=String(h||'').trim(); if (k) m[k]=i+1; }); return m; })(header);
  let cStage = H['Sales Stage'] || H['SalesStage'] || H['Stage'] || 0;
  if (!cStage) {
    sh.getRange(1, lc + 1).setValue('Sales Stage');
    header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    H = (function(hdr){ const m={}; hdr.forEach((h,i)=>{ const k=String(h||'').trim(); if (k) m[k]=i+1; }); return m; })(header);
    cStage = H['Sales Stage'];
  }
  const cell = sh.getRange(masterRowIndex, cStage);
  const cur = String(cell.getDisplayValue() || '').trim();
  if (!allowOverride && cur) return { ok:true, updated:false, prev:cur, value:cur };
  cell.setValue(value);
  return { ok:true, updated:true, value, prev:cur };
}

/*** === DEBUG / DIAGNOSTICS === ***/
function rp_debugConfig() {
  const show = (label, list) => {
    const hit = rp_propOneOf_(list || [], {label});
    return { label, tried:list, resolvedKey: hit.key || '(none)', valuePreview: hit.value ? (hit.value.slice(0,6) + '…') : '' };
  };
  const report = {
    ledgerFile:  show('LEDGER_FILE_ID', RP_KEY_ALIASES.LEDGER_FILE_ID),
    ledgerSheet: show('LEDGER_SHEET_NAME', RP_KEY_ALIASES.LEDGER_SHEET_NAME),
    hpOrders:    show('ORDERS_HPUSA_FILE_ID', RP_KEY_ALIASES.ORDERS_HPUSA_FILE_ID),
    vvsOrders:   show('ORDERS_VVS_FILE_ID', RP_KEY_ALIASES.ORDERS_VVS_FILE_ID),
    ordTab:      show('ORDERS_TAB_COMMON', RP_KEY_ALIASES.ORDERS_TAB_COMMON),
    hpTab:       show('ORDERS_HPUSA_TAB', RP_KEY_ALIASES.ORDERS_HPUSA_TAB),
    vvsTab:      show('ORDERS_VS_TAB', RP_KEY_ALIASES.ORDERS_VVS_TAB),
    arHP:        show('AR_HPUSA_ROOT_ID', RP_KEY_ALIASES.AR_HPUSA_ROOT_ID),
    arVVS:       show('AR_VVS_ROOT_ID', RP_KEY_ALIASES.AR_VVS_ROOT_ID),
    feesJson:    show('FEES_JSON', RP_KEY_ALIASES.FEES_JSON),
    feesTab:     show('FEES_TAB_NAME', RP_KEY_ALIASES.FEES_TAB_NAME)
  };
  Logger.log(JSON.stringify(report, null, 2));
  return report;
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



