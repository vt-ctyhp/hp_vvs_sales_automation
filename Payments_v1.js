/*** === CONFIG / CONSTANTS === ***/
var RP_DEBUG = (function(){
  try { return PropertiesService.getScriptProperties().getProperty('RP_DEBUG') === 'true'; }
  catch (_){ return false; }
})();
function RP_LOG() { if (RP_DEBUG) Logger.log.apply(Logger, arguments); }
function rp_log() { try { RP_LOG.apply(null, arguments); } catch(_) {} }
function rp_time(label) {
  var t = Date.now();
  return function(){ RP_LOG('%s %dms', label || '⏱', Date.now() - t); };
}

const RP_MASTER_SHEET = '00_Master Appointments';

function rp_getTaxRateForBrand(brand) {
  try { return rp_getTaxRate_(brand || ''); } catch(_) { return 0; }
}

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

var RP_DOC_STATUS = { DRAFT:'DRAFT', ISSUED:'ISSUED', REPLACED:'REPLACED', VOID:'VOID' };
var RP_DOC_ROLE   = { DEPOSIT:'DEPOSIT', PROGRESS:'PROGRESS', FINAL:'FINAL', CREDIT:'CREDIT', PAYMENT_RECEIPT:'PAYMENT_RECEIPT', SALES_RECEIPT:'SALES_RECEIPT' };

function rp_prop_(k, d){
  try { return PropertiesService.getScriptProperties().getProperty(k) || d || ''; }
  catch(_){ return d || ''; }
}

var TEMPLATE_CM_VVS_ID = rp_prop_('TEMPLATE_CM_VVS_ID','');
var TEMPLATE_CM_HP_ID  = rp_prop_('TEMPLATE_CM_HP_ID','');
var TAX_RATE_DEFAULT = Number(rp_prop_('TAX_RATE_DEFAULT','0.09375'));
var TAX_MODE_DEFAULT = rp_prop_('TAX_MODE_DEFAULT','TAX_INCLUDED');
var TAX_ROUNDING     = rp_prop_('TAX_ROUNDING','INVOICE');
var TAX_DECIMALS     = Number(rp_prop_('TAX_DECIMALS','2'));
var ALLOW_SALES_RECEIPT_PARTIAL = (rp_prop_('ALLOW_SALES_RECEIPT_PARTIAL','false') === 'true');

const RP_PROPS = PropertiesService.getScriptProperties();
const RP_TZ    = Session.getScriptTimeZone() || 'America/Los_Angeles';

function rp_propOneOf_(aliases, opt) {
  opt = opt || {};
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
  // ← THÊM: xóa cache cũ để tránh load nhầm row
  try { CacheService.getUserCache().remove('RP_ACTIVE_MASTER_ROW'); } catch(_){}
  try { rp_markActiveMasterRowIndex_(); } catch(_){}
  const html = HtmlService.createTemplateFromFile('dlg_record_payment_v1').evaluate().setWidth(980).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Record Payment');
}
function rp_ping(){ RP_LOG('[rp_ping]'); return 'pong'; }

function rp_markActiveMasterRowIndex_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RP_MASTER_SHEET);
  const rng = sh && sh.getActiveRange();
  const row = rng ? rng.getRow() : 0;
  if (row >= 2) CacheService.getUserCache().put('RP_ACTIVE_MASTER_ROW', String(row), 300);
  return row;
}

function rp_truncateDesc_(text) {
  const charsPerLine = 50;  // đo được từ template
  const maxLines     = 2;   // cho phép tối đa 2 dòng
  const maxChars     = charsPerLine * maxLines; // = 116 ký tự

  if (!text || text.length <= maxChars) return text || '';
  return text.slice(0, maxChars - 1).trim() + '…';
}


/*** === UTILS === ***/
function rp_money(n){ var v=Number(n||0); if(!isFinite(v)) v=0; var parts=v.toFixed(2).split('.'); parts[0]=parts[0].replace(/\B(?=(\d{3})+(?!\d))/g,','); return '$'+parts.join('.'); }
function rp_fmtDateYMD_(d){ return Utilities.formatDate(d, RP_TZ, 'yyyy-MM-dd'); }
function rp_fileIdFromUrl(url){ const s=String(url==null?'':url); let m=s.match(/\/d\/([-\w]{25,})/); if(m&&m[1]) return m[1]; m=s.match(/[?&]id=([-\w]{25,})/); if(m&&m[1]) return m[1]; m=s.match(/[-\w]{25,}/); return m?m[0]:''; }
function rp_sanitizeForFolder_(s) {
  return String(s || '').trim().replace(/[\\\/]+/g, '-').replace(/\s+/g, ' ').replace(/^-+|-+$/g, '');
}
function rp_soEq(a,b){ const sa=String(a==null?'':a).trim(), sb=String(b==null?'':b).trim(); if(sa===sb) return true; const na=Number(sa.replace(/[^\d.]/g,'')), nb=Number(sb.replace(/[^\d.]/g,'')); if(!isNaN(na)&&!isNaN(nb)) return Math.abs(na-nb)<1e-9; return false; }
function rp_headerMap(values) { const headers = (values && values[0]) || []; const map = {}; headers.forEach((h,i)=>{ map[String(h).trim()] = i; }); return map; }
function rp_hIndex_(headerRow) { const H = {}; (headerRow||[]).forEach((h,i)=>{ const k=String(h||'').trim(); if (k) H[k]=i+1; }); return H; }
function rp_pick(H, ...names) { for (const n of names) { if (H[n]) return H[n]; } return 0; }
function rp_pick0(map, ...names) { for (const n of names) { if (map[n] != null) return map[n]; } return -1; }
function rp_num_(v) {
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
  return { rowIndex: row, rootApptId: String(rowVals[apptIdx] || '').trim(), customerName: String(rowVals[custIdx] || '').trim(), soNumber: String((soIdx != null ? rowVals[soIdx] : '') || '').trim(), trackerUrl, map, rowVals, sh };
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
        const snap = { brand: t.brand, sheetName: t.tabName, soNumber, orderTotal: row[otIdx] || '', paidToDate: row[ptdIdx] || '', balance: (balIdx >= 0 ? row[balIdx] : ''), lastPaymentDate: (lpdIdx >= 0 ? row[lpdIdx] : ''), paymentsFolderURL: (pfIdx >= 0 ? row[pfIdx] : '') };
        try { cache.put(key, JSON.stringify(snap), 60); } catch (_){}
        return snap;
      }
    }
  }
  return null;
}


/*** === PREFILL API === ***/
// function rp_init() {
//   Logger.log('[rp_init] start');
//   const stop = rp_time && rp_time('[rp_init] total');
//   let prevTaxEnabled = true; 
//   try {
//     const cache = CacheService.getUserCache();
//     const cached = cache.get('RP_ACTIVE_MASTER_ROW');
//     const forcedRow = cached ? Number(cached) : 0;
//     const master = (forcedRow >= 2) ? rp_getMasterRowByIndex_(forcedRow) : rp_activeMasterRow();
//     const sh = master.sh, map = master.map, rowVals = master.rowVals;
//     const rowValsRaw = sh.getRange(master.rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
//     const otM  = map['Order Total'] != null ? rp_num_(rowValsRaw[map['Order Total']]) : 0;
//     const ptdM = (function(){ const idx = rp_pick0(map, 'Paid-to-Date','Paid-To-Date','Paid to Date','Paid-to-date'); return idx >= 0 ? rp_num_(rowValsRaw[idx]) : 0; })();
//     const brand      = (map['Brand'] != null) ? String(rowVals[map['Brand']] || '').trim() : '';
//     const hasSO      = !!(master.soNumber && String(master.soNumber).trim());
//     const anchorType = hasSO ? 'SO' : 'APPT';
//     const orderTotalRaw = otM > 0 ? otM : 0;
//     const paidToDate    = ptdM > 0 ? ptdM : 0;  
//     let orderTotal = orderTotalRaw;
//     const { sh: ledgerSh } = rp_getLedgerTarget();
//     const lr = ledgerSh.getLastRow(), lc = ledgerSh.getLastColumn();
//     if (lr >= 2) {
//       const head = ledgerSh.getRange(1,1,1,lc).getValues()[0].map(v=>String(v).trim());
//       const H = {}; head.forEach((h,i)=> H[h]=i);
//       const cAppt = H['RootApptID'], cSO = H['SO#'];
//       const cTax  = H['TaxEnabled'];  // cột được ghi bởi rp_submit v8.7
//       if (cTax != null && (cAppt != null || cSO != null)) {
//         const start = Math.max(2, lr - 500);
//         const vals  = ledgerSh.getRange(start,1,lr-start+1,lc).getValues();
//         for (let i = vals.length - 1; i >= 0; i--) {
//           const r = vals[i];
//           const match = hasSO
//             ? rp_soEq(r[cSO], master.soNumber)
//             : String(r[cAppt]||'').trim() === String(master.rootApptId||'').trim();
//           if (match) {
//             // Lấy row đầu tiên (mới nhất) khớp anchor
//             const raw = r[cTax];
//             // TaxEnabled ghi là true/false hoặc TRUE/FALSE string
//             prevTaxEnabled = (raw === false || String(raw).toLowerCase() === 'false') ? false : true;
//             break;
//           }
//         }
//       }
//     }
//     try {
//       if (orderTotalRaw > 0 && brand) {
//         const taxRateInit = prevTaxEnabled ? (() => { try { return rp_getTaxRate_(brand); } catch(_){ return 0; } })() : 0;
//         if (orderTotalRaw > 0 && taxRateInit > 0) {
//           orderTotal = rp_round2(orderTotalRaw * (1 + taxRateInit));
//         }
//       }
//     } catch(_) {}
//     const balance = Math.max(0, orderTotal - paidToDate);
//     const paymentsFolderURL = (function(){ const pfIdx = rp_pick0(map, 'PaymentsFolderURL'); return pfIdx >= 0 ? String(rowVals[pfIdx] || '').trim() : ''; })();
//     let lastPaymentDate = '';
//     let prevPayments    = [];
//     try {
//       const prev = rp_prevPaymentsForAnchor_({ anchorType, rootApptId: master.rootApptId, soNumber: master.soNumber, limit: 10 });
//       if (prev && prev.items && prev.items.length) {
//         lastPaymentDate = String(prev.items[0].date || '');
//         prevPayments    = prev.items.map(it => ({ date: it.date || '', amount: Number(it.amount || 0), method: it.method || '', docNumber: it.docNumber || '' }));
//       }
//     } catch (_) {}
//     const out = {
//       anchorType, brand, rootApptId: master.rootApptId,
//       customerName: master.customerName, soNumber: master.soNumber || '',
//       trackerUrl: master.trackerUrl || '',
//       orderTotal: String(orderTotal || ''),
//       paidToDate: String(paidToDate || ''),
//       balance: String(balance),
//       lastPaymentDate, prevPayments, paymentsFolderURL,
//       masterRowIndex: master.rowIndex,
//       taxEnabled: prevTaxEnabled,   // ← THÊM DÒNG NÀY
//     };
//     let saved = null;
//     try { const mObj = rp_getMasterRowByIndex_(out.masterRowIndex); saved = rp_readSavedLinesFromMaster_(mObj); } catch(_) {}
//     if (!saved) { try { saved = rp_findLastSavedLinesForAnchor_({ anchorType: out.anchorType, rootApptId: out.rootApptId, soNumber: out.soNumber }); } catch(_) {} }
//     if (saved && saved.lines && saved.lines.length) { out.savedLines = saved.lines; out.savedSubtotal = saved.subtotal || 0; }
//     Logger.log('[rp_init] out: ' + JSON.stringify(out));
//     if (stop) stop();

//     // THÊM: đọc phone + email từ 100_ Master
//     const phoneIdx = rp_pick0(map, 'Phone', 'Phone Number', 'Tel', 'Mobile');
//     const emailIdx = rp_pick0(map, 'Email', 'Email Address', 'E-mail');
//     out.phone = phoneIdx >= 0 ? String(rowVals[phoneIdx] || '').trim() : '';
//     out.email = emailIdx >= 0 ? String(rowVals[emailIdx] || '').trim() : '';


//     return out;
//   } catch (e) {
//     Logger.log('[rp_init] ERROR: ' + (e && e.stack ? e.stack : e));
//     throw e;
//   }
  
// }

function rp_init() {
  Logger.log('[rp_init] start');
  const stop = rp_time && rp_time('[rp_init] total');

  // v8.8 FIX: theo dõi xem có tìm thấy bản ghi TaxEnabled không
  let prevTaxEnabled = true;
  let rp_foundTaxRecord = false;

  try {
    const cache = CacheService.getUserCache();
    const cached = cache.get('RP_ACTIVE_MASTER_ROW');
    const forcedRow = cached ? Number(cached) : 0;
    const master = (forcedRow >= 2) ? rp_getMasterRowByIndex_(forcedRow) : rp_activeMasterRow();
    const sh = master.sh, map = master.map, rowVals = master.rowVals;
    const rowValsRaw = sh.getRange(master.rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
    const otM  = map['Order Total'] != null ? rp_num_(rowValsRaw[map['Order Total']]) : 0;
    const ptdM = (function(){ const idx = rp_pick0(map, 'Paid-to-Date','Paid-To-Date','Paid to Date','Paid-to-date'); return idx >= 0 ? rp_num_(rowValsRaw[idx]) : 0; })();
    const brand      = (map['Brand'] != null) ? String(rowVals[map['Brand']] || '').trim() : '';
    const hasSO      = !!(master.soNumber && String(master.soNumber).trim());
    const anchorType = hasSO ? 'SO' : 'APPT';
    const orderTotalRaw = otM > 0 ? otM : 0;
    const paidToDate    = ptdM > 0 ? ptdM : 0;
    let orderTotal = orderTotalRaw;

    const { sh: ledgerSh } = rp_getLedgerTarget();
    const lr = ledgerSh.getLastRow(), lc = ledgerSh.getLastColumn();
    if (lr >= 2) {
      const head = ledgerSh.getRange(1,1,1,lc).getValues()[0].map(v=>String(v).trim());
      const H = {}; head.forEach((h,i)=> H[h]=i);
      const cAppt = H['RootApptID'], cSO = H['SO#'];
      const cTax  = H['TaxEnabled'];
      if (cTax != null && (cAppt != null || cSO != null)) {
        const start = Math.max(2, lr - 500);
        const vals  = ledgerSh.getRange(start,1,lr-start+1,lc).getValues();
        for (let i = vals.length - 1; i >= 0; i--) {
          const r = vals[i];
          const match = hasSO
            ? rp_soEq(r[cSO], master.soNumber)
            : String(r[cAppt]||'').trim() === String(master.rootApptId||'').trim();
          if (match) {
            const raw = r[cTax];
            // Chỉ dùng nếu ô KHÔNG rỗng
            if (raw !== '' && raw !== null && raw !== undefined) {
              prevTaxEnabled = (raw === false || String(raw).toLowerCase() === 'false') ? false : true;
              rp_foundTaxRecord = true;
            }
            break;
          }
        }
      }
    }

    // ★ v8.8 SỬA CHÍNH:
    // Nếu không tìm thấy bản ghi TaxEnabled (dữ liệu cũ trước khi có tính năng thuế)
    // VÀ khách đã có thanh toán trước đó (paidToDate > 0)
    // → Tổng đơn hàng đã bao gồm thuế rồi, KHÔNG cộng thuế thêm lần nữa
    if (!rp_foundTaxRecord && ptdM > 0) {
      prevTaxEnabled = false;
    }

    try {
      if (orderTotalRaw > 0 && brand) {
        const taxRateInit = prevTaxEnabled ? (() => { try { return rp_getTaxRate_(brand); } catch(_){ return 0; } })() : 0;
        if (orderTotalRaw > 0 && taxRateInit > 0) {
          orderTotal = rp_round2(orderTotalRaw * (1 + taxRateInit));
        }
      }
    } catch(_) {}

    const balance = Math.max(0, orderTotal - paidToDate);
    const paymentsFolderURL = (function(){ const pfIdx = rp_pick0(map, 'PaymentsFolderURL'); return pfIdx >= 0 ? String(rowVals[pfIdx] || '').trim() : ''; })();
    let lastPaymentDate = '';
    let prevPayments    = [];
    try {
      const prev = rp_prevPaymentsForAnchor_({ anchorType, rootApptId: master.rootApptId, soNumber: master.soNumber, limit: 10 });
      if (prev && prev.items && prev.items.length) {
        lastPaymentDate = String(prev.items[0].date || '');
        prevPayments    = prev.items.map(it => ({ date: it.date || '', amount: Number(it.amount || 0), method: it.method || '', docNumber: it.docNumber || '' }));
      }
    } catch (_) {}

    const out = {
      anchorType, brand, rootApptId: master.rootApptId,
      customerName: master.customerName, soNumber: master.soNumber || '',
      trackerUrl: master.trackerUrl || '',
      orderTotal: String(orderTotal || ''),
      paidToDate: String(paidToDate || ''),
      balance: String(balance),
      lastPaymentDate, prevPayments, paymentsFolderURL,
      masterRowIndex: master.rowIndex,
      taxEnabled: prevTaxEnabled,
    };

    let saved = null;
    try { const mObj = rp_getMasterRowByIndex_(out.masterRowIndex); saved = rp_readSavedLinesFromMaster_(mObj); } catch(_) {}
    if (!saved) { try { saved = rp_findLastSavedLinesForAnchor_({ anchorType: out.anchorType, rootApptId: out.rootApptId, soNumber: out.soNumber }); } catch(_) {} }
    // ... existing savedLines block ...
    if (saved && saved.lines && saved.lines.length) {
      out.savedLines    = saved.lines;
      out.savedSubtotal = saved.subtotal || 0;
    }

    // ★ FIX v8.9 CORRECTED: Ưu tiên InvoiceTotal từ ledger (đã tính đúng referral + tax)
    // Chỉ fallback sang savedSubtotal nếu không có InvoiceTotal nào trong ledger
    if (!orderTotalRaw) {
      try {
        const lastIT = rp_getLastInvoiceTotalFromLedger_({
          hasSO,
          soNumber:   master.soNumber,
          rootApptId: master.rootApptId
        });

        if (lastIT > 0) {
          // ✅ Dùng InvoiceTotal từ ledger — đã có tax + referral
          out.orderTotal = String(lastIT);
          out.balance    = String(Math.max(0, lastIT - paidToDate));
          Logger.log('[rp_init] v8.10: InvoiceTotal from ledger=%s → OT=%s balance=%s',
            lastIT, out.orderTotal, out.balance);
        } 
        else if (out.savedSubtotal > 0) {
          // Fallback: không có ledger row nào → estimate từ savedSubtotal
          // Lưu ý: savedSubtotal KHÔNG có referral discount, chỉ dùng khi không còn cách nào khác
          const fallbackTaxRate = prevTaxEnabled ? rp_getTaxRate_(brand) : 0;
          const derivedOT       = rp_round2(out.savedSubtotal * (1 + fallbackTaxRate));
          out.orderTotal = String(derivedOT);
          out.balance    = String(Math.max(0, derivedOT - paidToDate));
          Logger.log('[rp_init] FIX v8.9: OT blank, no ledger row → fallback savedSubtotal=%s → OT=%s',
            out.savedSubtotal, derivedOT);
        }
      } catch(e) {
        Logger.log('[rp_init] FIX v8.9 error: ' + e.message);
      }
    }

    Logger.log('[rp_init] out: ' + JSON.stringify(out));
    if (stop) stop();

    // đọc phone + email từ Master
    const phoneIdx = rp_pick0(map, 'Phone', 'Phone Number', 'Tel', 'Mobile');
    const emailIdx = rp_pick0(map, 'Email', 'Email Address', 'E-mail');
    out.phone = phoneIdx >= 0 ? String(rowVals[phoneIdx] || '').trim() : '';
    out.email = emailIdx >= 0 ? String(rowVals[emailIdx] || '').trim() : '';

    // Project #22: kiểm tra referral đã dùng chưa
    try {
      const ref = rp_isReferralAlreadyUsed_({
        anchorType: out.anchorType,
        rootApptId: out.rootApptId,
        soNumber:   out.soNumber
      });
      out.referralUsed     = ref.used;
      out.referralName     = ref.name     || '';
      out.referralDiscount = ref.discount || 0;
    } catch(_) {
      out.referralUsed     = false;
      out.referralName     = '';
      out.referralDiscount = 0;
    }

    return out;


  } catch (e) {
    Logger.log('[rp_init] ERROR: ' + (e && e.stack ? e.stack : e));
    throw e;
  }
}


/**
 * Project #21 — Gọi riêng từ dialog sau khi prefill xong
 * KHÔNG nhúng vào rp_init() để tránh ảnh hưởng các trigger khác
 */
function rp_checkHasSalesInvoice(payload) {
  try {
    const { anchorType, rootApptId, soNumber } = payload || {};
    const result = rp_checkInvoiceBeforeReceipt_({
      anchorType, rootApptId, soNumber,
      docType: 'Sales Receipt'
    });
    return { ok: true, hasSalesInvoice: result.ok };
  } catch(e) {
    Logger.log('[rp_checkHasSalesInvoice] ERROR: ' + e.message);
    return { ok: true, hasSalesInvoice: false };
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
  const start = Math.max(2, lr - 1000);
  const vals  = sh.getRange(start,1,lr-start+1,lc).getValues();
  const out   = [];
  for (let i = vals.length - 1; i >= 0; i--) {
    const r = vals[i];
    const match = String(anchorType||'').toUpperCase()==='SO' ? (String(r[cSO]||'').trim() === String(soNumber||'').trim()) : (String(r[cAppt]||'').trim() === String(rootApptId||'').trim());
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
    const hasAny = !!((s.ringStyle&&String(s.ringStyle).trim())||(s.metalType&&String(s.metalType).trim())||(s.accentType&&String(s.accentType).trim())||(s.ringSize&&String(s.ringSize).trim())||(s.centerType&&String(s.centerType).trim())||(s.dimensions&&String(s.dimensions).trim()));
    return hasAny ? { ok:true, reason:'OK', spec:s } : { ok:false, reason:'EMPTY_FIELDS', spec:null };
  } catch (e) { return { ok:false, reason:'EXCEPTION: ' + (e && e.message ? e.message : e), spec:null }; }
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
    let match = false;
    if (cSO != null && soNumber) match = rp_soEq(r[cSO], soNumber);
    if (!match && cAP != null && rootApptId) {
      const a = String(r[cAP] || '').toUpperCase().replace(/[\u200B-\u200D\uFEFF]/g,'').trim();
      const b = String(rootApptId||'').toUpperCase().replace(/[\u200B-\u200D\uFEFF]/g,'').trim();
      match = !!a && !!b && (a === b || a.endsWith(b) || b.endsWith(a));
    }
    if (!match) continue;
    let t = 0;
    const v = r[cTS];
    if (v instanceof Date) t = v.getTime();
    else if (v) { const p = Date.parse(String(v)); if (!isNaN(p)) t = p; } else { t = (i+1); }
    const candidate = { t, r, ringStyle: cStyle!=null?String(r[cStyle]||'').trim():'', metalType: cMetal!=null?String(r[cMetal]||'').trim():'', accentType: cAccent!=null?String(r[cAccent]||'').trim():'', ringSize: cSize!=null?String(r[cSize]||'').trim():'', centerType: cCenter!=null?String(r[cCenter]||'').trim():'', dimensions: cDims!=null?String(r[cDims]||'').trim():'' };
    if (!best || candidate.t > best.t) best = candidate;
  }
  if (!best) return { ok:false, reason:'NO_MATCH', spec:null };
  const out = { ringStyle:best.ringStyle, metalType:best.metalType, accentType:best.accentType, ringSize:best.ringSize, centerType:best.centerType, dimensions:best.dimensions };
  if (!out.ringStyle && !out.metalType && !out.accentType && !out.ringSize && !out.centerType && !out.dimensions) return { ok:false, reason:'NO_FIELDS_FOR_MATCH', spec:null };
  return { ok:true, reason:'OK', spec: out };
}


/*** === LEDGER TARGET === ***/
function rp_getLedgerTarget() {
  const fileRes = rp_propOneOf_(RP_KEY_ALIASES.LEDGER_FILE_ID, { required:true, label:'Payments Ledger File ID' });
  const sheetRes = rp_propOneOf_(RP_KEY_ALIASES.LEDGER_SHEET_NAME, { required:false, label:'Payments sheet name' });
  const fileId = fileRes.value;
  const sheetName = sheetRes.value || 'Payments';
  const ss = SpreadsheetApp.openById(fileId);
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
  headersNeeded.forEach(h => { if (map[h] == null) { sh.getRange(1, cursor+1).setValue(h); map[h] = cursor; cursor++; } });
  return map;
}


/*** === FEES === ***/
const RP_FEES_CACHE_KEY = 'PAYMENT_FEES::v1';
const RP_FEES_CACHE_TTL = 3600;
function rp_parseFeeCell_(v){ if(v==null||v==='') return 0; const s=String(v).trim(); if(/%/.test(s)){ const n=parseFloat(s.replace(/[^\d.\-]/g,'')); return isNaN(n)?0:n/100;} const n=parseFloat(s.replace(/[^\d.\-]/g,'')); if(isNaN(n)) return 0; return n>1 ? n/100 : Math.max(0,n); }
function rp_readFeesFromSheet_(){ try{ const p=PropertiesService.getScriptProperties(); const tab=p.getProperty('PAYMENTS_FEES_TAB_NAME')||p.getProperty('CFG_PAYMENTS_FEES_TAB_NAME')||'Current Fees'; const { sh }=rp_getLedgerTarget(); const ss=sh.getParent(); const s=ss.getSheetByName(tab); if(!s) return null; const lr=s.getLastRow(), lc=s.getLastColumn(); if(lr<2||lc<1) return null; const headers=s.getRange(1,1,1,lc).getValues()[0].map(v=>String(v).trim()); const hmap={}; headers.forEach((h,i)=>hmap[h]=i); const methodCol=hmap['Method']!=null ? hmap['Method'] : hmap['Payment Method']; const feeCol=hmap['Fee %']!=null ? hmap['Fee %'] : (hmap['Fee']!=null?hmap['Fee']:hmap['Percent']); if(methodCol==null||feeCol==null) return null; const vals=s.getRange(2,1,lr-1,lc).getValues(); const out={}; vals.forEach(row=>{ const m=String(row[methodCol]||'').trim(); if(!m) return; out[m] = rp_parseFeeCell_(row[feeCol]); }); return Object.keys(out).length ? out : null; } catch(_){ return null; } }
function rp_readFeesFromProp_(){ const p=PropertiesService.getScriptProperties(); const raw=p.getProperty('PAYMENT_FEES_JSON')||p.getProperty('CFG_PAYMENT_FEES_JSON'); if(!raw) return null; try{ const obj=JSON.parse(raw); const out={}; Object.keys(obj).forEach(k=>{ out[k]=rp_parseFeeCell_(obj[k]); }); return out; }catch(_){ return null; } }
function rp_getDefaultFees_(){ return {"Credit Card":0.03,"Synchrony":0.06,"Wire":0,"Zelle":0,"Cash":0,"Check":0,"Other":0}; }
function rp_getFees(){ const cache=CacheService.getUserCache(); const cached=cache.get(RP_FEES_CACHE_KEY); if(cached){ try{ return JSON.parse(cached); }catch(_){}} let fees=rp_readFeesFromProp_(); if(!fees) fees=rp_readFeesFromSheet_(); if(!fees) fees=rp_getDefaultFees_(); try{ cache.put(RP_FEES_CACHE_KEY, JSON.stringify(fees), RP_FEES_CACHE_TTL); }catch(_){} return fees; }
function rp_refreshFeesCache(){ CacheService.getUserCache().remove(RP_FEES_CACHE_KEY); return 'Fees cache cleared.'; }


/*** === SUBMIT === ***/
function rp_amountForMasterOnSOReceipt_(pmt) {
  const prop = (PropertiesService.getScriptProperties().getProperty('SO_RECEIPT_MASTER_AMOUNT') || PropertiesService.getScriptProperties().getProperty('CFG_SO_RECEIPT_MASTER_AMOUNT') || 'ALLOC').toUpperCase();
  return prop === 'GROSS' ? Number((pmt && pmt.amount) || 0) : Number((pmt && pmt.allocatedToSO) || 0);
}

function rp_submit(payload) {
  if (!payload) throw new Error('Empty submit payload.');
  const { anchorType, brand, rootApptId, soNumber, docType, lines, pmt, customerName } = payload;
  if (!docType) throw new Error('Doc Type is required.');
  if (!lines || !lines.length) throw new Error('At least one line is required.');

  const subtotal = lines.reduce((s, ln) => s + (Number(ln.qty||0) * Number(ln.amt||0)), 0);
  if (!(subtotal > 0)) throw new Error('Lines subtotal must be greater than 0.');

  const isReceipt = /Receipt/i.test(docType);
  const amountGross = isReceipt ? Number(pmt.amount||0) : 0;
  if (isReceipt && !(amountGross > 0)) throw new Error('Payment Amount is required for receipts.');


  if (/sales\s*receipt/i.test(docType)) {
    const prereq = rp_checkInvoiceBeforeReceipt_({ anchorType, rootApptId, soNumber, docType });
    if (!prereq.ok) {
      throw new Error(prereq.message || 'Vui lòng tạo Sales Invoice trước khi tạo Sales Receipt.');
    }
  }

  // ===== TAX ENGINE =====
  // v8.7: taxEnabled flag — default true, only false if user explicitly unchecked
  const taxEnabled         = (payload.taxEnabled !== false);
  const taxRate            = taxEnabled ? rp_getTaxRate_(brand) : 0;
  const referralDiscount   = (payload.referralEnabled && payload.referralDiscount)
                             ? Math.max(0, Number(payload.referralDiscount || 0)) : 0;
  const discountedSubtotal = Math.max(0, subtotal - referralDiscount);
  const taxAmount          = rp_round2(discountedSubtotal * taxRate);
  const invoiceTotal       = rp_round2(discountedSubtotal + taxAmount);
  // const balanceDue         = rp_round2(invoiceTotal - (isReceipt ? Number(pmt.amount||0) : 0));
  const snapshotPtd        = (payload.snapshots && payload.snapshots.paidToDate)
                           ? Number(payload.snapshots.paidToDate) : 0;
  const balanceDue         = isReceipt
                            ? rp_round2(invoiceTotal - Number(pmt.amount || 0))
                            : rp_round2(invoiceTotal - snapshotPtd);

  const fees = rp_getFees();
  const feePct = isReceipt ? Number(fees[pmt.method] || 0) : 0;
  const feeAmt = isReceipt ? +(amountGross * feePct).toFixed(2) : 0;
  const amountNet = isReceipt ? +(amountGross - feeAmt).toFixed(2) : 0;
  const allocToSO = isReceipt ? Number(pmt.amount || 0) : 0;

  const stamp = Utilities.formatDate(new Date(), RP_TZ, 'yyyyMMdd-HHmmss');
  const anchorKey = soNumber ? soNumber : rootApptId;
  const basketId = `BASK-${anchorKey}-${stamp}`;
  const paymentId = `PAY-${Utilities.getUuid()}`;
  const submittedAt = new Date();
  const submittedBy = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'Unknown';

  var docStatus  = String(payload.docStatus || '').toUpperCase() || 'DRAFT';
  var docRole    = String(payload.docRole   || '').toUpperCase();
  var supersedes = String(payload.supersedes|| '').trim();
  var appliesTo  = String(payload.appliesTo || '').trim();

  if (!docRole) {
    var dt = String(payload.docType || '').toUpperCase();
    if      (dt.indexOf('CREDIT')   >= 0) docRole = 'CREDIT';
    else if (dt.indexOf('PROGRESS') >= 0) docRole = 'PROGRESS';
    else if (dt.indexOf('DEPOSIT')  >= 0 && dt.indexOf('INVOICE') >= 0) docRole = 'DEPOSIT';
    else if (dt.indexOf('DEPOSIT')  >= 0 && dt.indexOf('RECEIPT') >= 0) docRole = 'DEPOSIT';
    else if (dt.indexOf('INVOICE')  >= 0) docRole = 'FINAL';
    else {
      var hasLines = Array.isArray(payload.lines) && payload.lines.length;
      docRole = hasLines ? 'SALES_RECEIPT' : 'PAYMENT_RECEIPT';
    }
  }

const rowObj = {
    'PAYMENT_ID': paymentId, 'Brand': brand || '', 'RootApptID': rootApptId || '', 'SO#': soNumber || '',
    'Customer Name': customerName || '',
    'TaxEnabled': taxEnabled,
    'AnchorType': anchorType || '', 'BasketID': basketId, 'DocType': docType,
    'PaymentDateTime': isReceipt ? (pmt.dateTime || '') : '',
    'Method': isReceipt ? (pmt.method || '') : '', 'Reference': isReceipt ? (pmt.reference || '') : '', 'Notes': isReceipt ? (pmt.notes || '') : '',
    'AmountGross': amountGross, 'FeePercent': feePct, 'FeeAmount': feeAmt, 'AmountNet': amountNet, 'AllocatedToSO': allocToSO,
    'LinesJSON': JSON.stringify(lines),
    'Subtotal': +subtotal.toFixed(2),
    'ReferralDiscount': referralDiscount,
    'DiscountedSubtotal': +discountedSubtotal.toFixed(2),
    'TaxRate': taxRate,
    'TaxAmount': taxAmount,
    'InvoiceTotal': invoiceTotal,
    'BalanceDue': balanceDue,
    'Order Total_SO': (payload.snapshots && payload.snapshots.orderTotal) ? String(payload.snapshots.orderTotal) : '',
    'Paid-To-Date_SO': (payload.snapshots && payload.snapshots.paidToDate) ? String(payload.snapshots.paidToDate) : '',
    'Balance_SO': (() => {
      const snapBal = (payload.snapshots && payload.snapshots.balance)
                      ? Number(payload.snapshots.balance) : 0;
      // Receipt → balance sau khi trả; Invoice → balance trước (chưa trả)
      return String(isReceipt ? Math.max(0, snapBal - amountGross) : snapBal);
    })(),
    'Submitted By': submittedBy, 'Submitted Date/Time': submittedAt,
    'ReferralEnabled': !!(payload.referralEnabled),
    'ReferralName':    String(payload.referralName  || ''),
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

  rp_ensureHeaders_(sh, ['DocStatus','DocRole','SupersedesDoc#','AppliesToDoc#']);
  var headerRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var H1 = rp_hIndex_(headerRow);
  var cDocStatus  = rp_pick(H1, 'DocStatus','Status');
  var cDocRole    = rp_pick(H1, 'DocRole','Role');
  var cSupersedes = rp_pick(H1, 'SupersedesDoc#','Supersedes','Replaces','ReplacesDoc#');
  var cAppliesTo  = rp_pick(H1, 'AppliesToDoc#','Applies To','SettlesDoc#','Settles');
  if (cDocStatus)  sh.getRange(nextRow, cDocStatus ).setValue(docStatus);
  if (cDocRole)    sh.getRange(nextRow, cDocRole   ).setValue(docRole);
  if (cSupersedes) sh.getRange(nextRow, cSupersedes).setValue(supersedes);
  if (cAppliesTo)  sh.getRange(nextRow, cAppliesTo ).setValue(appliesTo);

  const setFlag = !!(payload && payload.flags && payload.flags.setOrderTotal);
  if (setFlag) {
    // v8.8 FIX: ghi invoiceTotal (có thuế) thay vì subtotal thuần
    const orderTotalToWrite = invoiceTotal;
    if (anchorType === 'APPT') {
      try {
        const resMaster = rp_setOrderTotal_Master_Safe_(rootApptId, orderTotalToWrite, true, payload.masterRowIndex);
        rp_auditOrderTotalOnLedger_(nextRow, { set: !!resMaster.updated, value: (resMaster.updated ? orderTotalToWrite : ''), prev: resMaster.prev, source: docType, target: 'APPT' });
      } catch (e) { Logger.log('APPT OT write error: ' + (e && e.message ? e.message : e)); }
    } else if (anchorType === 'SO') {
      try {
        const resMaster2 = rp_setOrderTotal_Master_Safe_(rootApptId, orderTotalToWrite, true, payload.masterRowIndex);
        rp_auditOrderTotalOnLedger_(nextRow, { set: !!resMaster2.updated, value: (resMaster2.updated ? orderTotalToWrite : ''), prev: resMaster2.prev, source: docType, target: 'MASTER' });
      } catch (e) { Logger.log('SO OT write error: ' + (e && e.message ? e.message : e)); }
    }
  }
  if (setFlag) { try { rp_persistSavedLinesToMaster_({ masterRowIndex: payload.masterRowIndex, rootApptId, lines: payload.lines, subtotal }); } catch (e) { Logger.log('Saved lines persist warning: ' + (e && e.message ? e.message : e)); } }

  try {
    if (/receipt/i.test(docType)) {
      const when = (payload && payload.pmt && payload.pmt.dateTime) ? new Date(payload.pmt.dateTime) : new Date();
      const amtForMaster = (anchorType === 'SO') ? rp_amountForMasterOnSOReceipt_(payload.pmt || { amount: amountGross, allocatedToSO: allocToSO }) : amountGross;
      let netAmtForMaster = amtForMaster;
      const supersedesDoc = String(payload && payload.supersedes || '').trim();
      if (supersedesDoc) {
        try {
          const sup = rp_findLedgerRowByDocNumber_(supersedesDoc);
          if (sup) {
            const tOld = String(sup.rowVals[sup.H['DocType']] || '').toUpperCase();
            const statusOld = (sup.H['DocStatus'] != null ? String(sup.rowVals[sup.H['DocStatus']] || '') : '').toUpperCase().trim();
            const sameAnchor = (function(){ const aNew = String(anchorType || '').toUpperCase(); if (aNew === 'SO') return rp_soEq(sup.rowVals[sup.H['SO#']], soNumber); return String(sup.rowVals[sup.H['RootApptID']] || '').trim() === String(rootApptId || '').trim(); })();
            if (tOld.includes('RECEIPT') && sameAnchor && statusOld !== 'VOID' && statusOld !== 'REPLACED') { const prevApplied = rp_getAppliedAmtForMasterOnReceiptRow_(sup.rowVals, sup.H) || 0; netAmtForMaster = amtForMaster - prevApplied; rp_updateLedgerRow_(sup.row, { 'DocStatus': 'VOID' }); }
            if (tOld.includes('INVOICE') && sameAnchor && statusOld !== 'VOID' && statusOld !== 'REPLACED') { rp_updateLedgerRow_(sup.row, { 'DocStatus': 'REPLACED' }); }
          }
        } catch (e) { Logger.log('Supersedes handling warning: ' + (e && e.message ? e.message : e)); }
      }
      if (payload.masterRowIndex) { rp_applyReceiptToMaster({ masterRowIndex: payload.masterRowIndex, amount: netAmtForMaster, when }); }
    }
  } catch (e) { Logger.log('Receipt write-back warning: ' + (e && e.message ? e.message : e)); }

  try { if (/receipt/i.test(docType)) { const mRow = Number(payload.masterRowIndex || 0); if (mRow >= 2 && rootApptId) { rp_updateMasterCashInGross_({ masterRowIndex: mRow, rootApptId }); } } } catch (e) { Logger.log('[Cash-in Gross] refresh skipped: ' + (e && e.message ? e.message : e)); }
  try { if (/receipt/i.test(docType)) { const count = rp_countReceiptsForAppt_(rootApptId); if (count === 1 && payload.masterRowIndex) rp_setSalesStageOnMaster_({ masterRowIndex: payload.masterRowIndex, value: 'Deposit', allowOverride: false }); } } catch (e) { Logger.log('Sales Stage set skipped: ' + ((e && e.message) ? e.message : e)); }

  // Project #22: ghi referral vào Client Status Report nếu có
  if (payload.referralEnabled && payload.referralName) {
    try {
      rp_applyReferralToClientStatus_({
        masterRowIndex: payload.masterRowIndex,
        rootApptId,
        referralName:     payload.referralName,
        referralDiscount: payload.referralDiscount || 100,
        submittedAt
      });
    } catch(e) {
      Logger.log('[Project #22] Referral write-back warning: ' + (e && e.message ? e.message : e));
    }
  }

  try { if (typeof swInvalidatePaymentReadModelsAfterWrite_ === 'function') swInvalidatePaymentReadModelsAfterWrite_(null, 'Payment submitted'); } catch (_) {}
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

// function rp_getTemplateIdFor(brand, docType) {
//   const p = PropertiesService.getScriptProperties();
//   const bRaw = String(brand || '').toUpperCase();
//   const normBrand = /VVS/.test(bRaw) ? 'VVS' : /HPUSA/.test(bRaw) ? 'HPUSA' : bRaw.replace(/[^A-Z0-9]/g, '');
//   const normType = String(docType || '').toUpperCase().replace(/[^A-Z]/g, ' ').replace(/\s+/g, ' ').trim().replace(/ /g, '_');
//   const codeMap = { DEPOSIT_INVOICE:'DI', DEPOSIT_RECEIPT:'DR', SALES_INVOICE:'SI', SALES_RECEIPT:'SR' };
//   const code = codeMap[normType];
//   const primary = [`${normBrand}_${code}_TEMPLATE_ID`, `CFG_${normBrand}_${code}_TEMPLATE_ID`];
//   const canonical = [`TEMPLATE_${normType}_${normBrand}`, `CFG_TEMPLATE_${normType}_${normBrand}`];
//   const aliasKey = `TEMPLATE_${normType}_${normBrand}`;
//   const aliasList = (RP_KEY_ALIASES && RP_KEY_ALIASES[aliasKey]) ? RP_KEY_ALIASES[aliasKey] : [];
//   const keysToTry = [...primary, ...canonical, ...aliasList];
//   for (const k of keysToTry) { const v = p.getProperty(k); if (v && String(v).trim()) return String(v).trim(); }
//   const titles = [`[TEMPLATE] ${docType} -- ${normBrand}`, `[TEMPLATE] ${docType} — ${normBrand}`];
//   for (const name of titles) { const it = DriveApp.searchFiles(`title = "${name}"`); if (it.hasNext()) return it.next().getId(); }
//   throw new Error(`Template not found. Brand="${brand}" -> "${normBrand}", DocType="${docType}" -> "${normType}". Tried keys: ${keysToTry.join(', ')}`);
// }

function rp_getTemplateIdFor(brand, docType, taxEnabled) {
  const p = PropertiesService.getScriptProperties();

  const bRaw = String(brand || '').toUpperCase();
  const normBrand = /VVS/.test(bRaw) ? 'VVS'
                  : /HPUSA/.test(bRaw) ? 'HPUSA'
                  : bRaw.replace(/[^A-Z0-9]/g, '');

  const normType = String(docType || '').toUpperCase()
    .replace(/[^A-Z]/g, ' ').replace(/\s+/g, ' ').trim().replace(/ /g, '_');

  const codeMap = {
    DEPOSIT_INVOICE: 'DI', DEPOSIT_RECEIPT: 'DR',
    SALES_INVOICE:   'SI', SALES_RECEIPT:   'SR'
  };
  const code = codeMap[normType];
  if (!code) throw new Error(`Unknown docType: "${docType}".`);

  // ★ HPUSA → single template cho tất cả 4 loại
  if (normBrand === 'HPUSA') {
    const key = `HPUSA_${code}_TEMPLATE_ID`;
    const v = p.getProperty(key);
    if (v && String(v).trim()) {
      RP_LOG('[rp_getTemplateIdFor] HPUSA single → key=%s', key);
      return String(v).trim();
    }
    throw new Error(`HPUSA template not found for "${code}". Key: ${key}`);
  }

  // VVS → TAX / NOTAX
  const useTax = (taxEnabled !== false);
  const taxSuffix = useTax ? 'TAX' : 'NOTAX';

  const keysToTry = [
    `${normBrand}_${code}_${taxSuffix}_TEMPLATE_ID`,
    `CFG_${normBrand}_${code}_${taxSuffix}_TEMPLATE_ID`,
    `${normBrand}_${code}_TEMPLATE_ID`,
    `CFG_${normBrand}_${code}_TEMPLATE_ID`,
    ...((RP_KEY_ALIASES[`TEMPLATE_${normType}_${normBrand}`] || [])),
    `TEMPLATE_${normType}_${normBrand}`,
    `CFG_TEMPLATE_${normType}_${normBrand}`,
  ];

  for (const k of keysToTry) {
    if (!k) continue;
    const v = p.getProperty(k);
    if (v && String(v).trim()) {
      RP_LOG('[rp_getTemplateIdFor] %s tax=%s → key=%s', normBrand, useTax, k);
      return String(v).trim();
    }
  }

  throw new Error(
    `Template not found. Brand="${normBrand}", DocType="${normType}", Tax=${useTax}.\n` +
    `Tried: ${keysToTry.filter(Boolean).join(', ')}`
  );
}

function rp_renderLinesText_(lines){
  const out = (lines||[]).map(ln => String(ln && ln.desc || '').trim()).filter(Boolean).map(s => { const t = s.replace(/^\s+|\s+$/g,''); return t.startsWith('✧') ? t : ('✧ ' + t); });
  return out.join('\n');
}

function rp_fillDocPlaceholders_(docId, replacements) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const hdr  = (doc.getHeader && doc.getHeader()) || null;
  const ftr  = (doc.getFooter && doc.getFooter()) || null;
  const repl = Object.assign({}, replacements || {});
  if (repl.ORDER_TOTAL_SO != null && repl.ORDER_TOTAL == null) repl.ORDER_TOTAL = repl.ORDER_TOTAL_SO;
  if (repl.PAID_TO_DATE_BEFORE != null && repl.Paid_to_date == null) repl.Paid_to_date = repl.PAID_TO_DATE_BEFORE;
  if (repl.BALANCE_AFTER != null && repl.BALANCE == null) repl.BALANCE = repl.BALANCE_AFTER;
  if (repl.BALANCE_BEFORE != null && repl.BALANCE == null) repl.BALANCE = repl.BALANCE_BEFORE;
  if (repl.pmtId && !repl.PMT_ID) repl.PMT_ID = repl.pmtId;
  (function bridgeTitleCasePlaceholders() {
    function alias(srcKey, aliases) { if (repl[srcKey] == null) return; aliases.forEach(k => { if (repl[k] == null) repl[k] = repl[srcKey]; }); }
    alias('ORDER_TOTAL', ['Order Total', 'ORDER TOTAL']);
    alias('Paid_to_date', ['Paid-To-Date', 'Paid to Date', 'PAID-TO-DATE', 'PAID TO DATE']);
    if (repl.BALANCE != null) alias('BALANCE', ['BALANCE_SO']);
    alias('PAYMENT_AMOUNT',    ['Payment Amount', 'Amount Paid']);
    alias('REQ_AMT', ['Requested Amount', 'REQUESTED_AMOUNT', 'REQ AMT']);
    alias('PAYMENT_METHOD',    ['Payment Method', 'Method']);
    alias('PAYMENT_REFERENCE', ['Payment Reference', 'Reference']);
    alias('LINES_SUBTOTAL', ['SUBTOTAL', 'Sub Total', 'SUB_TOTAL', 'Lines Subtotal', 'LINES SUBTOTAL']);
    alias('REQ_AMT', ['Requested Amount', 'REQUESTED_AMOUNT', 'REQ AMT','Deposit Amount', 'DEPOSIT_AMOUNT']);
    alias('BALANCE_BEFORE', ['BALANCE_DUE','Balance Due']);
  })();
  const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const sections = [body, hdr, ftr].filter(Boolean);
  Object.keys(repl).forEach(key => {
    const pat = '\\{\\{\\s*' + esc(key) + '\\s*\\}\\}';
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

function rp_deletePlaceholderRowOrLine_(docId, key){
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const pat = '\\{\\{\\s*' + key + '\\s*\\}\\}';
  let range = body.findText(pat);
  let removedAny = false;
  while (range) {
    const el = range.getElement();
    let node = el;
    let row = null;
    while (node && node.getParent) {
      try { if (node.getType && node.getType() === DocumentApp.ElementType.TABLE_ROW) { row = node.asTableRow(); break; } } catch (_) {}
      node = node.getParent && node.getParent();
    }
    if (row) { row.removeFromParent(); removedAny = true; }
    else { try { el.getParent().asParagraph().removeFromParent(); removedAny = true; } catch(_){} }
    range = body.findText(pat);
  }
  doc.saveAndClose();
  return removedAny;
}

function rp_insertPrevRowUnderOT_(docId, label, amountText) {
  const doc  = DocumentApp.openById(docId);
  const body = doc.getBody();
  const existing = body.findText(/Previous Payments/i);
  if (existing) { doc.saveAndClose(); return false; }
  const norm = s => String(s || '').replace(/\s+/g, ' ').trim().toUpperCase();
  let target = null;
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.TABLE) continue;
    const t = el.asTable();
    for (let r = 0; r < t.getNumRows(); r++) {
      const row = t.getRow(r);
      if (row.getNumCells() < 2) continue;
      const left = norm(row.getCell(0).getText());
      if (left.includes('ORDER TOTAL')) { target = { t, r }; break; }
    }
    if (target) break;
  }
  if (!target) { doc.saveAndClose(); return false; }
  const { t, r } = target;
  if (r + 1 < t.getNumRows() && norm(t.getRow(r + 1).getCell(0).getText()).includes('PREVIOUS PAYMENTS')) { doc.saveAndClose(); return false; }
  const cols = t.getRow(0).getNumCells();
  const nr = t.insertTableRow(r + 1);
  nr.appendTableCell(label);
  for (let c = 1; c < cols - 1; c++) nr.appendTableCell('');
  nr.appendTableCell(String(amountText));
  doc.saveAndClose();
  return true;
}

function rp_fillItemsTable_(docId, lines) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const toUpper = s => String(s || '').trim().toUpperCase();
  const SUMMARY_KEYWORDS = ['SUB TOTAL','SUBTOTAL','SALES TAX','INVOICE TOTAL','ORDER TOTAL','DEPOSIT PAID','DEPOSIT AMOUNT','BALANCE','PAYMENT AMOUNT','AMOUNT PAID','TOTAL'];
  const isSummaryRow = (row) => { try { const cell0 = toUpper(row.getCell(0).getText()); return SUMMARY_KEYWORDS.some(k => cell0.includes(k)); } catch(_) { return false; } };
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
  if (!table) { doc.saveAndClose(); return { usedTable: false }; }
  function setCellText(cell, text, align, bold) {
    try { while (cell.getNumChildren() > 0) { cell.removeChild(cell.getChild(0)); } } catch(_) {}

    // ★ FIX: truncate cứng tại 120 ký tự để tránh giãn dòng
    const maxLen = 120;
    const displayText = (text && text.length > maxLen)
      ? text.slice(0, maxLen).trim() + '…'
      : (text || '');

    const para = cell.appendParagraph(displayText);
    if (align) para.setAlignment(align);
    try {
      const t = para.editAsText();

      // Auto-scale font size theo độ dài
      const len = displayText.length;
      const fontSize = len > 90 ? 6.5
                    : len > 60 ? 7
                    : 8;

      t.setFontSize(fontSize);
      if (bold) t.setBold(true);
    } catch(_) {}
    try { cell.setPaddingTop(3); cell.setPaddingBottom(3); cell.setPaddingLeft(5); cell.setPaddingRight(5); } catch(_) {}
  }
  let lastDataRow = table.getNumRows() - 1;
  for (let r = 1; r < table.getNumRows(); r++) { if (isSummaryRow(table.getRow(r))) { lastDataRow = r - 1; break; } }
  const dataRows    = lastDataRow;
  const linesToFill = (lines || []).slice(0, dataRows);
  for (let i = 0; i < dataRows; i++) {
    const row = table.getRow(i + 1);
    const ln  = linesToFill[i];
    if (ln) {
      const qty   = (ln.qty != null) ? ln.qty : '';
      const amt   = Number(ln.amt || 0);
      const total = Number(qty || 0) * amt;
      setCellText(row.getCell(0), rp_truncateDesc_(String(ln.desc || '').trim()), DocumentApp.HorizontalAlignment.LEFT, true);
      setCellText(row.getCell(1), String(qty),                  DocumentApp.HorizontalAlignment.RIGHT, true);
      setCellText(row.getCell(2), rp_money(total),              DocumentApp.HorizontalAlignment.RIGHT, true);
    } else {
      const cell0text = row.getCell(0).getText() || '';
      const cell2text = row.getCell(2).getText() || '';
      if (cell0text.includes('{{') || cell2text.includes('{{')) continue;
      setCellText(row.getCell(0), '', DocumentApp.HorizontalAlignment.LEFT,  false);
      setCellText(row.getCell(1), '', DocumentApp.HorizontalAlignment.RIGHT, false);
      setCellText(row.getCell(2), '', DocumentApp.HorizontalAlignment.RIGHT, false);
    }
  }
  doc.saveAndClose();
  return { usedTable: true };
}

function rp_formatPaymentsList_(items){
  return (items||[]).map(it => {
    const dt  = it.date || '';
    const amt = rp_money(Number(it.amount || 0));
    const raw = String(it.method || '').trim();
    const mth = /^zelle/i.test(raw) ? 'Zelle' : /^credit card$|^card$/i.test(raw) ? 'Credit Card' : raw;
    return `✧ ${dt} — ${amt} ${mth}`.trim();
  }).join('\n');
}

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
    const rowIndex = start + i;
    if (beforeRow && rowIndex >= beforeRow) continue;
    const r = vals[i];
    const type = String(r[cType] || '').toLowerCase();
    if (!(type.includes('receipt'))) continue;
    const status = cStatus != null ? String(r[cStatus] || '').toUpperCase().trim() : '';
    if (status === 'VOID' || status === 'REPLACED' || status === 'DRAFT') continue;
    if (String(anchorType||'').toUpperCase() === 'SO') { if (!rp_soEq(r[cSO], soNumber)) continue; }
    else { if (String(r[cAppt] || '').trim() !== String(rootApptId || '').trim()) continue; }
    const whenRaw = r[cWhen];
    let when = null;
    if (whenRaw instanceof Date) when = whenRaw;
    else if (whenRaw) { const p = Date.parse(String(whenRaw)); if (!isNaN(p)) when = new Date(p); }
    const cPayId = H['PAYMENT_ID'] != null ? H['PAYMENT_ID'] : H['PaymentId'];
    out.push({ when, date: when ? rp_fmtDateYMD_(when) : '', amount: Number(r[cGross] || 0), method: String(r[cMethod] || ''), paymentId: cPayId != null ? String(r[cPayId] || '') : '' });
  }
  out.sort((a,b)=> (b.when?b.when.getTime():0) - (a.when?a.when.getTime():0));
  const seen = new Set();
  const deduped = out.filter(it => { const key = it.paymentId || `${it.date}|${it.amount}|${it.method}`; if (seen.has(key)) return false; seen.add(key); return true; });
  return { items: (limit && limit>0) ? deduped.slice(0, limit) : deduped };
}

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
  try { const arr = JSON.parse(raw); const lines = (arr||[]).map(ln => ({ desc: rp_sanitizeDesc_(ln.desc), qty: Number(ln.qty)||0, amt: Number(ln.amt)||0 })); return { lines, subtotal }; } catch(_){ return null; }
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
    const isMatch = (String(anchorType||'').toUpperCase()==='SO') ? rp_soEq(r[cSO], soNumber) : (String(r[cAppt]||'').trim() === String(rootApptId||'').trim());
    if (!isMatch) continue;
    const raw = String(r[cJSON]||'').trim();
    if (!raw) continue;
    try { const arr = JSON.parse(raw); const lines = (arr||[]).map(ln => ({ desc: rp_sanitizeDesc_(ln.desc), qty: Number(ln.qty)||0, amt: Number(ln.amt)||0 })); const subtotal = cSub != null ? Number(r[cSub]||0) : 0; return { lines, subtotal }; } catch(_){ }
  }
  return null;
}


/** Generate Doc+PDF — v8.7: taxEnabled fully respected */
function rp_generateDocAndPdf_(brand, docType, destFolder, payload, ledgerMeta) {
  if (!destFolder) throw new Error('Destination folder not resolved.');
  const taxEnabled = (payload.taxEnabled !== false);
  const tmplId = rp_getTemplateIdFor(brand, docType, taxEnabled);
  const now = new Date();
  const so = payload.soNumber || '';
  let docNumber, baseName;

  if (payload.anchorType === 'SO') {
    const version = rp_nextDocVersion_(payload.anchorType, payload.rootApptId, so, docType);
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

  const lines = (payload.lines || []).slice(0, 5);
  const tableRes = rp_fillItemsTable_(docId, lines);

  // Khai báo linesSubtotalNum TRƯỚC khi dùng
  const linesSubtotalNum = lines.reduce((s, ln) => s + Number(ln.qty||0)*Number(ln.amt||0), 0);
  const linesSubtotal    = rp_money(linesSubtotalNum);

  // Project #22: referral discount
  const referralDiscount   = (payload.referralEnabled && payload.referralDiscount)
                             ? Math.max(0, Number(payload.referralDiscount || 0)) : 0;
  const discountedSubtotal = Math.max(0, linesSubtotalNum - referralDiscount);

  const num = v => Number(String(v == null ? '' : v).replace(/[^\d.\-]/g, '')) || 0;

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

  const isReceipt   = /Receipt/i.test(docType);
  const pmt         = payload.pmt || {};
  const amountGross = isReceipt ? num(pmt.amount) : 0;
  const allocToSO   = isReceipt ? num(pmt.allocatedToSO) : 0;

  // v8.7 FIX: taxEnabled fully respected — taxRate = 0 when disabled
  // const taxEnabled   = (payload.taxEnabled !== false);
const taxRate      = taxEnabled ? rp_getTaxRate_(brand) : 0;
  const taxAmount    = rp_round2(discountedSubtotal * taxRate);
  const invoiceTotal = rp_round2(discountedSubtotal + taxAmount);

  const feeAmt       = isReceipt ? rp_round2(amountGross * (rp_getFees()[pmt.method] || 0)) : 0;
  const amountNet    = isReceipt ? rp_round2(amountGross - feeAmt) : 0;

  // v8.7 FIX: orderTotalWithTax recalculated correctly when taxEnabled = false
  const orderTotalWithTax = referralDiscount > 0
    ? invoiceTotal   // ← khi có discount, luôn dùng invoiceTotal (đã trừ discount)
    : (!taxEnabled
        ? invoiceTotal
        : (orderTotalBefore > 0 ? orderTotalBefore : invoiceTotal));

  const paidAfter    = isReceipt ? (paidBefore + (allocToSO || amountGross)) : paidBefore;
  const balAfterRaw  = isReceipt ? Math.max(0, orderTotalWithTax - paidAfter) : Math.max(0, orderTotalWithTax - paidBefore);

  const prevLimit = Number(PropertiesService.getScriptProperties().getProperty('PREV_PAYMENTS_LIMIT') || 10);
  const prev = rp_prevPaymentsForAnchor_({ anchorType: payload.anchorType, rootApptId: payload.rootApptId, soNumber: payload.soNumber, beforeRow: null, limit: prevLimit });
  const prevItems = (prev.items || []).map(it => ({ date: it.date || '', amount: Number(it.amount || 0), method: it.method || '', docNumber: it.docNumber || '' }));
  const prevSumNum = prevItems.reduce((s, it) => s + (Number(it.amount) || 0), 0);
  const hasPrev    = (prevItems.length > 0) && (prevSumNum > 0);
  const prevLabel  = hasPrev ? 'Previous Payments' : '';
  const prevBlock  = hasPrev ? rp_formatPaymentsList_(prevItems) : '';

  const methodDisplay = (() => {
    const m = String(pmt.method || '').trim();
    if (/^zelle/i.test(m)) return 'Zelle';
    if (/^credit card$|^card$/i.test(m)) return 'Credit Card';
    return m;
  })();

  const repl = {
    'DOC_DATE'     : Utilities.formatDate(now, RP_TZ, 'MMM d, yyyy'),
    'CUSTOMER_NAME': payload.customerName || '',
    'ROOT_APPT_ID' : payload.rootApptId || '',
    'SO_NUMBER'    : so || '',
    'DOC_NUMBER'   : docNumber || '',
    'PMT_ID'       : payload.pmtId || '',
    'LINES'        : tableRes.usedTable ? '' : rp_renderLinesText_(lines),
    'LINES_SUBTOTAL': rp_money(discountedSubtotal),
    ...(() => { const m={}; for(let i=0;i<5;i++){ const ln=lines[i]||{}; m['DESC_'+(i+1)]=(tableRes.usedTable? '' : (ln.desc||'')); m['QTY_'+(i+1)]=tableRes.usedTable? '' : (ln.qty!=null?String(ln.qty):''); m['AMT_'+(i+1)]=tableRes.usedTable? '' : (ln.amt!=null?rp_money(num(ln.amt)):''); } return m; })(),
    'ORDER_TOTAL_SO'      : (orderTotalBefore || orderTotalBefore === 0) ? rp_money(orderTotalBefore) : '',
    'PAID_TO_DATE_BEFORE' : (paidBefore       || paidBefore       === 0) ? rp_money(paidBefore)       : '',
    'BALANCE_BEFORE'      : (balBefore        || balBefore        === 0) ? rp_money(balBefore)        : '',
    'PAID_TO_DATE_AFTER'  : isReceipt ? rp_money(paidAfter)   : '',
    'BALANCE_AFTER'       : isReceipt ? rp_money(balAfterRaw) : '',
    'PAYMENT_AMOUNT'      : isReceipt ? rp_money(amountGross) : rp_money(paidBefore > 0 ? paidBefore : prevSumNum),
    'PAYMENT_METHOD'      : isReceipt ? methodDisplay : '',
    'PAYMENT_REFERENCE'   : isReceipt ? (pmt.reference || '') : '',
    'PAYMENT_NOTES'       : isReceipt ? (pmt.notes || '') : '',
    'PREVIOUS_PAYMENTS_LABEL': prevLabel,
    'PREVIOUS_PAYMENTS_BLOCK': prevBlock,
    'TAX_AMOUNT':    rp_money(taxAmount),
    'INVOICE_TOTAL': rp_money(invoiceTotal),
    'FEE_AMOUNT':    rp_money(feeAmt),
    'AMOUNT_NET':    rp_money(amountNet),
    'BALANCE_DUE':   isReceipt ? rp_money(rp_round2(invoiceTotal - amountGross)) : '',
    'PHONE':         payload.phone || '',
    'EMAIL':         payload.email || '',
    'REFERRAL_NAME':     (payload.referralEnabled && payload.referralName) ? payload.referralName : '',
    'REFERRAL_DISCOUNT': (payload.referralEnabled && payload.referralDiscount) ? rp_money(payload.referralDiscount) : '',
    ...(() => {
      const selected = isReceipt ? String(pmt.method || '').trim().toLowerCase() : '';
      const chk = (method) => {
        const m = method.toLowerCase();
        if (m === 'credit card') return (selected === 'credit card' || selected === 'card') ? '☑' : '☐';
        if (m === 'zelle') return selected.startsWith('zelle') ? '☑' : '☐';
        return selected === m ? '☑' : '☐';
      };
      return { 'CHK_CASH': chk('Cash'), 'CHK_ZELLE': chk('Zelle'), 'CHK_CHECK': chk('Check'), 'CHK_TRADE_IN': chk('Trade-in'), 'CHK_BANK_WIRE': chk('Bank Wire'), 'CHK_CREDIT_CARD': chk('Credit Card'), 'CHK_DEPOSIT_CREDIT': chk('Deposit Credit'), 'CHK_OTHER': chk('Other') };
    })()
  };

  // Reconcile Order Totals and Balance
(function reconcileOrderTotalsAndBalance(){
    const isReceipt = /Receipt/i.test(docType);
    const isInvoice = /Invoice/i.test(docType);
    const baseOT = orderTotalWithTax;
    const payApplied = isReceipt ? ((String(payload.anchorType || '').toUpperCase() === 'SO') ? allocToSO : amountGross) : 0;
    const paidBeforeForMath = (paidBefore > 0 ? paidBefore : Math.max(0, prevSumNum - payApplied));
    const paidAfterNum = isReceipt ? (paidBeforeForMath + payApplied) : paidBeforeForMath;
    const balBeforeNum = Math.max(0, baseOT - paidBeforeForMath);
    const balAfterNum  = isReceipt ? Math.max(0, baseOT - paidAfterNum) : balBeforeNum;

    repl.ORDER_TOTAL_SO      = rp_money(baseOT);
    repl.ORDER_TOTAL         = repl.ORDER_TOTAL_SO;
    repl.PAID_TO_DATE_BEFORE = rp_money(paidBeforeForMath);
    repl.BALANCE_BEFORE      = rp_money(balBeforeNum);

    if (isReceipt) {
      repl.PAID_TO_DATE_AFTER = rp_money(paidAfterNum);
      repl.BALANCE_AFTER      = rp_money(balAfterNum);
      repl.BALANCE_DUE        = rp_money(balAfterNum);
    }

    if (/Sales\s*Receipt/i.test(docType)) {
      repl.PAYMENT_AMOUNT = rp_money(balBeforeNum);
    }
    repl.BALANCE = isReceipt ? (repl.BALANCE_AFTER || '') : (repl.BALANCE_BEFORE || '');

    // if (isInvoice) {
    //   const reqAmt = num(pmt.amount);
    //   repl.REQ_AMT = rp_money(reqAmt);
    //   const paidSoFar = paidBeforeForMath > 0 ? paidBeforeForMath : prevSumNum;
    //   repl.PAYMENT_AMOUNT  = rp_money(paidSoFar);
    //   repl.DEPOSIT_AMOUNT  = rp_money(paidSoFar);
    //   repl.BALANCE_DUE     = rp_money(balBeforeNum);
    // }
    if (isInvoice) {
      const reqAmt = num(pmt.amount);

      // ★ FIX: DEPOSIT_AMOUNT = số tiền user nhập, không phải paidSoFar
      const depositToShow = reqAmt > 0 ? reqAmt
        : (paidBeforeForMath > 0 ? paidBeforeForMath : prevSumNum);

      repl.REQ_AMT         = rp_money(reqAmt > 0 ? reqAmt : 0);
      repl.PAYMENT_AMOUNT  = rp_money(depositToShow);
      repl.DEPOSIT_AMOUNT  = rp_money(depositToShow);

      // BALANCE_DUE = Total − Deposit Amount
      const balAfterDeposit = Math.max(0, balBeforeNum - (reqAmt > 0 ? reqAmt : 0));
      repl.BALANCE_DUE = rp_money(balAfterDeposit);
    }

    const EPS = 0.005;
    const balNumToShow = isReceipt ? balAfterNum : balBeforeNum;
    if (balNumToShow <= EPS) {
      repl.BALANCE = '$0.00'; repl.BALANCE_DUE = '$0.00';
      repl.BALANCE_BEFORE = '$0.00'; repl.BALANCE_AFTER = '$0.00';
    }
    if (/Sales\s*Invoice/i.test(docType)) {
      repl.PAYMENT_AMOUNT = rp_money(balBeforeNum);
      repl.BALANCE_DUE    = rp_money(balBeforeNum);
    }

    // v8.7: khi taxEnabled = false → set $0.00, xóa row SAU KHI fill (xem bên dưới)
    if (!taxEnabled) {
      repl.TAX_AMOUNT    = '$0.00';
      repl.INVOICE_TOTAL = rp_money(discountedSubtotal);  // ← $325 (đúng)
    }
    // if (!taxEnabled) {
    //   delete repl.TAX_AMOUNT;          // KHÔNG fill → placeholder còn nguyên trong doc
    //   repl.INVOICE_TOTAL = linesSubtotal;  // Total = Subtotal khi không có thuế
    // }
  })();

  if (payload.anchorType === 'SO') {
    repl.BALANCE_SO = /Receipt/i.test(docType) ? (repl.BALANCE_AFTER || '') : (repl.BALANCE_BEFORE || '');
  }

  const codeInfo = rp_docCodeFromDocType_(docType);
  const paidToDateForRow = (paidBefore > 0) ? paidBefore : prevSumNum;
  const showPrevRowUnderOT = (codeInfo.code !== 'SR') && (paidToDateForRow > 0);

  repl.ORDER_TOTAL  = repl.ORDER_TOTAL_SO || '';
  repl.Paid_to_date = showPrevRowUnderOT ? rp_money(paidToDateForRow) : '';
  repl.BALANCE      = repl.BALANCE_AFTER || repl.BALANCE_BEFORE || '';

  // ===== FILL PLACEHOLDERS =====
  rp_fillDocPlaceholders_(docId, repl);

  // ★ UPDATE: HPUSA DR/SR dùng template duy nhất → KHÔNG xóa dòng tax,
  // chỉ fill $0.00 (đã được set trong reconcileOrderTotalsAndBalance)
  // Các trường hợp khác vẫn xóa dòng như cũ
  const bNorm = String(brand || '').toUpperCase();
  const dtNorm = String(docType || '').toUpperCase().replace(/[^A-Z]/g,' ').replace(/\s+/g,' ').trim().replace(/ /g,'_');
  const isHPUSA_Single = /HPUSA/.test(bNorm) && (dtNorm === 'DEPOSIT_RECEIPT' || dtNorm === 'SALES_RECEIPT');

  if (!taxEnabled && !isHPUSA_Single) {
    try { rp_deletePlaceholderRowOrLine_(docId, 'TAX_AMOUNT'); } catch(_) {}
    try { rp_deletePlaceholderRowOrLine_(docId, 'TAX_LABEL');  } catch(_) {}
  }
  // HPUSA DR/SR: TAX_AMOUNT đã được fill = '$0.00' → không cần làm gì thêm

  if (showPrevRowUnderOT) { rp_insertPrevRowUnderOT_(docId, 'Previous Payments', rp_money(paidToDateForRow)); }
  if (!showPrevRowUnderOT) { rp_deletePlaceholderLine_(docId, 'Paid_to_date'); }
  if (!hasPrev) { rp_deletePlaceholderLine_(docId, 'PREVIOUS_PAYMENTS_LABEL'); rp_deletePlaceholderLine_(docId, 'PREVIOUS_PAYMENTS_BLOCK'); }

  Utilities.sleep(200);
  const pdfBlob = DriveApp.getFileById(docId).getAs('application/pdf');
  pdfBlob.setName(baseName + '.pdf');
  const pdfFile = destFolder.createFile(pdfBlob);
  const pdfId   = pdfFile.getId();

  return { docId, pdfId, docNumber, docUrl: 'https://docs.google.com/document/d/' + docId + '/edit', pdfUrl: 'https://drive.google.com/file/d/' + pdfId + '/view' };
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
  for (const row of vals) { if (rp_soEq(row[soCol], soNumber)) { const dt = String(row[typeCol]||''); const fam = /Receipt/i.test(dt) ? 'Receipt' : 'Invoice'; if ((isReceipt && fam==='Receipt') || (!isReceipt && fam==='Invoice')) count++; } }
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
  if (newHeaders.length) { let cursor = headers.length; newHeaders.forEach(h => { sh.getRange(1, cursor+1).setValue(h); map[h]=cursor; cursor++; }); }
  const arr = sh.getRange(row,1,1,Math.max(...Object.values(map))+1).getValues()[0];
  Object.entries(updates).forEach(([h,v]) => { arr[map[h]] = v; });
  sh.getRange(row,1,1,arr.length).setValues([arr]);
}

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
  for (let i = vals.length - 1; i >= 0; i--) { const r = vals[i]; if (String(r[cDoc] || '').trim() === String(docNumber || '').trim()) { return { row: start + i, rowVals: r, H, sh }; } }
  return null;
}

function rp_getAppliedAmtForMasterOnReceiptRow_(rowVals, H) {
  const anchor = String(rowVals[H['AnchorType']] || '').toUpperCase();
  const gross  = Number(rowVals[H['AmountGross']] || 0);
  const alloc  = (H['AllocatedToSO'] != null) ? Number(rowVals[H['AllocatedToSO']] || 0) : 0;
  if (anchor === 'SO') { return rp_amountForMasterOnSOReceipt_({ amount: gross, allocatedToSO: alloc }); }
  return gross;
}

function rp_linkSupersession_(oldDocNumber, newDocNumber){
  try { if (!oldDocNumber || !newDocNumber) return; const hit = rp_findLedgerRowByDocNumber_(oldDocNumber); if (!hit) return; rp_updateLedgerRow_(hit.row, { 'ReplacedByDoc#': newDocNumber }); } catch (_) {}
}


/*** === FOLDERS === ***/
function rp_ensureChildFolder_(parent, name){ const it=parent.getFoldersByName(name); return it.hasNext()?it.next():parent.createFolder(name); }
function rp_getOrdersTabName_() { const p = PropertiesService.getScriptProperties(); return p.getProperty('301/302_TAB_NAME') || p.getProperty('ORDERS_TAB_NAME') || '1. Sales'; }

function rp_findSoRowInBrand_(brand, soNumber) {
  if (!brand || !soNumber) return null;
  const entry = (function(){ const props = PropertiesService.getScriptProperties(); const hp = props.getProperty('HPUSA_301_FILE_ID') || props.getProperty('HPUSA_ORDERS_FILE_ID') || props.getProperty('CFG_HPUSA_ORDERS_FILE_ID') || ''; const vvs = props.getProperty('VVS_302_FILE_ID')  || props.getProperty('VVS_ORDERS_FILE_ID')  || props.getProperty('CFG_VVS_ORDERS_FILE_ID') || ''; if (String(brand).toUpperCase().includes('HPUSA') && hp) return { brand:'HPUSA', fileId:hp }; if (String(brand).toUpperCase().includes('VVS')   && vvs) return { brand:'VVS',   fileId:vvs }; return null; })();
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

function rp_resolveDestAndClientFolders_(payload) {
  const wrote = { master:false, orders:false };
  try { if (payload && payload.paymentsFolderURL) { const fId = rp_fileIdFromUrl(payload.paymentsFolderURL); if (fId) { const f = DriveApp.getFolderById(fId); if (f) return { dest: f, paymentsFolderUrl: f.getUrl(), wrote }; } } } catch (_) {}
  // if (!payload || !payload.rootApptId) throw new Error('Bad payload for folder resolution.');
  if (!payload || (!payload.rootApptId && !payload.masterRowIndex)) throw new Error(
    'Bad payload for folder resolution. rootApptId="' + (payload && payload.rootApptId) + '" masterRowIndex="' + (payload && payload.masterRowIndex) + '"'
  );
  const m = payload.masterRowIndex ? rp_getMasterRowByIndex_(payload.masterRowIndex) : rp_findMasterRowByRootApptId_(payload.rootApptId);
  if (!m) throw new Error('Master row not found for RootApptID ' + payload.rootApptId);
  const clientFolderIdx = rp_pick0(m.map, 'Client Folder', 'ClientFolderURL', 'Customer Folder');
  if (clientFolderIdx < 0) throw new Error('Missing "Client Folder" column on 100_.');
  const clientFolderUrl = String(m.rowVals[clientFolderIdx] || '').trim();
  if (!clientFolderUrl) throw new Error('Client Folder URL is blank on this row.');
  const clientFolder = DriveApp.getFolderById(rp_fileIdFromUrl(clientFolderUrl));
  if (String(payload.anchorType).toUpperCase() !== 'SO' || !payload.soNumber) {
    const clientPaymentsFolder = rp_ensureChildFolder_(clientFolder, '04-Deposit');
    const pfIdxM = rp_pick0(m.map, 'PaymentsFolderURL');
    if (pfIdxM >= 0) { const curM = String(m.rowVals[pfIdxM] || '').trim(); if (!curM) { m.sh.getRange(m.rowIndex, pfIdxM + 1).setValue(clientPaymentsFolder.getUrl()); wrote.master = true; } }
    return { dest: clientPaymentsFolder, paymentsFolderUrl: clientPaymentsFolder.getUrl(), wrote };
  }
  const brandNorm = String(payload.brand || '').toUpperCase().includes('VVS') ? 'VVS' : 'HPUSA';
  const rootId = (brandNorm === 'VVS' ? rp_propOneOf_(RP_KEY_ALIASES.VVS_SO_ROOT_FOLDER_ID).value : rp_propOneOf_(RP_KEY_ALIASES.HPUSA_SO_ROOT_FOLDER_ID).value) || '';
  if (!rootId) throw new Error('SO root not configured in Script Properties.');
  const root = DriveApp.getFolderById(rootId);
  const so = String(payload.soNumber || '').trim();
  const stIdxM = rp_pick0(m.map, 'Short Tag','ShortTag','SO Short Tag','SO Tag');
  const shortTag = stIdxM >= 0 ? rp_sanitizeForFolder_(String(m.rowVals[stIdxM] || '')) : '';
  const folderLabel = [brandNorm, `SO${so}`, shortTag].filter(Boolean).join('-');
  const soFolder = (function(){ const it = root.getFoldersByName(folderLabel); return it.hasNext()? it.next() : root.createFolder(folderLabel); })();
  const paymentsFolder = rp_ensureChildFolder_(soFolder, '04-Deposit');
  const pfIdxM = rp_pick0(m.map, 'PaymentsFolderURL');
  if (pfIdxM >= 0) { const curM = String(m.rowVals[pfIdxM] || '').trim(); if (!curM) { m.sh.getRange(m.rowIndex, pfIdxM + 1).setValue(paymentsFolder.getUrl()); wrote.master = true; } }
  return { dest: paymentsFolder, paymentsFolderUrl: paymentsFolder.getUrl(), wrote };
}


/*** === DOC GENERATION + AR SHORTCUT === ***/
function rp_makeDocForPayment(ledgerRow, payload) {
  try {
    if (!payload || !payload.docType) return { ok:false, reason:'BAD_PAYLOAD', hint:'Missing payload or docType' };
    var docType = String(payload.docType);
    var anchorType = String(payload.anchorType || '');
    var brand = String(payload.brand || '').trim();
    try {
      const { sh } = rp_getLedgerTarget();
      const lc = sh.getLastColumn();
      const head = sh.getRange(1,1,1,lc).getValues()[0].map(v=>String(v).trim());
      const H = {}; head.forEach((h,i)=> H[h]=i);
      const rowVals = sh.getRange(ledgerRow, 1, 1, lc).getValues()[0];
      payload.pmtId = String(rowVals[H['PAYMENT_ID']] || '');
      if (!brand) { try { brand = String(rowVals[H['Brand']] || '').trim(); } catch (_){} }
    } catch(_){}
    var resolved = rp_resolveDestAndClientFolders_(payload);
    var destFolder = resolved.dest;
    if (!brand && anchorType === 'APPT') { try { brand = rp_brandFromMaster_(payload.rootApptId) || ''; } catch(_) {} }
    if (!brand) { return { ok:false, reason:'BRAND_NOT_FOUND_ON_MASTER', hint:'Brand is required to pick template' }; }
    var meta = { orderTotal: (payload.snapshots && payload.snapshots.orderTotal) || '', paidToDate: (payload.snapshots && payload.snapshots.paidToDate) || '', balance: (payload.snapshots && payload.snapshots.balance) || '', ledgerRow };
    var out = rp_generateDocAndPdf_(brand, docType, destFolder, payload, meta);
    if (!out || !out.docId || !out.pdfId) return { ok:false, reason:'DOC_GEN_RETURN_INVALID', hint:'Missing docId/pdfId from generator' };
    if (anchorType === 'APPT') {
      try {
        var ver = rp_nextDocVersion_(anchorType, payload.rootApptId, '', docType);
        var shortBase = rp_makeApptFilename_(brand, payload.rootApptId, docType, ver, new Date());
        DriveApp.getFileById(out.docId).setName(shortBase);
        DriveApp.getFileById(out.pdfId).setName(shortBase + '.pdf');
        out.docUrl = 'https://docs.google.com/document/d/' + out.docId + '/edit';
        out.pdfUrl = 'https://drive.google.com/file/d/' + out.pdfId + '/view';
      } catch (e) { Logger.log('APPT rename failed: ' + ((e && e.message) ? e.message : e)); }
    }
    try { rp_updateLedgerRow_(ledgerRow, { 'DocNumber': out.docNumber || '', 'DocFileID': out.docId || '', 'DocPDFID': out.pdfId || '', 'DocURL': out.docUrl || '', 'PDFURL': out.pdfUrl || '' }); } catch (e) { Logger.log('Ledger update (doc fields) failed: ' + ((e && e.message) ? e.message : e)); }
    try { if (payload && payload.supersedes && out && out.docNumber) { rp_linkSupersession_(payload.supersedes, out.docNumber); } } catch (e) { Logger.log('Supersession back-link warning: ' + ((e && e.message) ? e.message : e)); }
    var arShortcutURL = '';
    try {
      var arMonthly = rp_ensureArMonthlyFolder_(brand, new Date());
      var ar = null;
      if (arMonthly) { ar = rp_createDriveShortcut_(arMonthly.getId(), out.pdfId, (out.docNumber || 'Doc') + '.pdf'); arShortcutURL = (ar && ar.url) || ''; }
      rp_updateLedgerRow_(ledgerRow, { 'ARShortcutID': (ar && ar.id) || '', 'ARShortcutURL': arShortcutURL });
    } catch (e) { Logger.log('AR shortcut error: ' + ((e && e.message) ? e.message : e)); }
    return { ok: true, row: ledgerRow, brand, docNumber: out.docNumber || '', docId: out.docId, pdfId: out.pdfId, docUrl: out.docUrl, pdfUrl: out.pdfUrl, paymentsFolderURL: resolved.paymentsFolderUrl, wrote: resolved.wrote, arShortcutURL };
  } catch (e) { return { ok:false, reason:'UNCAUGHT', hint:(e && e.message) ? e.message : String(e) }; }
}


/*** === ORDER / MASTER HELPERS === ***/
function rp_brandFromMaster_(rootApptId) { const m = rp_findMasterRowByRootApptId_(rootApptId); if (!m) return ''; const idx = m.map['Brand']; return idx != null ? String(m.rowVals[idx] || '').trim() : ''; }

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
  for (const row of values) { const a = String(row[head['AnchorType']] || ''); const d = String(row[head['DocType']] || ''); const rid = String(row[head['RootApptID']] || ''); const so  = String(row[head['SO#']] || ''); if (a === anchorType && d === docType && (a==='APPT' ? (rid===String(rootApptId)) : (so===String(soNumber)))) { n++; } }
  return n+1;
}

function rp_createDriveShortcut_(parentFolderId, targetFileId, title) {
  function assertTargetVisible() { try { Drive.Files.get(targetFileId, { supportsAllDrives: true, supportsTeamDrives: true, fields: 'id' }); } catch (e) { Utilities.sleep(400); Drive.Files.get(targetFileId, { supportsAllDrives: true, supportsTeamDrives: true, fields: 'id' }); } }
  assertTargetVisible();
  var resource = { title, mimeType: 'application/vnd.google-apps.shortcut', parents: [{ id: parentFolderId }], shortcutDetails: { targetId: targetFileId, targetMimeType: 'application/pdf' } };
  var file = Drive.Files.insert(resource, null, { supportsAllDrives: true, supportsTeamDrives: true, fields: 'id' });
  return { id: file.id, url: 'https://drive.google.com/file/d/' + file.id + '/view' };
}

function rp_getMasterRowByIndex_(rowIndex) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RP_MASTER_SHEET);
  if (!sh) throw new Error(`Missing sheet "${RP_MASTER_SHEET}"`);
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (rowIndex < 2 || rowIndex > lr) throw new Error('Invalid master row index: ' + rowIndex);
  const header = sh.getRange(1,1,1,lc).getDisplayValues();
  const map = rp_headerMap(header);
  const rowVals = sh.getRange(rowIndex,1,1,lc).getDisplayValues()[0];
  const apptIdx = rp_pick0(map, 'APPT_ID','RootApptID','Root Appt ID');
  const custIdx = rp_pick0(map, 'Customer Name','Customer','Client Name','Client');
  const soIdx   = rp_pick0(map, 'SO#','SO','SO Number','Sales Order','Sales Order #');
  const trkIdx  = rp_pick0(map, '3D Tracker','3D Log');
  if (apptIdx < 0) throw new Error('Missing "APPT_ID"/RootApptID column on ' + RP_MASTER_SHEET);
  if (custIdx < 0) throw new Error('Missing "Customer Name"/Customer column on ' + RP_MASTER_SHEET);
  let trackerUrl = '';
  if (trkIdx >= 0) { trackerUrl = String(rowVals[trkIdx] || '').trim(); if (!trackerUrl) { try { const rich = sh.getRange(rowIndex, trkIdx + 1).getRichTextValue(); if (rich) { trackerUrl = rich.getLinkUrl() || ''; if (!trackerUrl && rich.getRuns) { const runs = rich.getRuns(); for (let i = 0; i < runs.length; i++) { const u = runs[i].getLinkUrl && runs[i].getLinkUrl(); if (u) { trackerUrl = u; break; } } } } } catch (_) {} } }
  return { rowIndex, rootApptId: String(rowVals[apptIdx] || '').trim(), customerName: String(rowVals[custIdx] || '').trim(), soNumber: String((soIdx >= 0 ? rowVals[soIdx] : '') || '').trim(), trackerUrl, map, rowVals, sh };
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
  for (let r=1; r<values.length; r++){ if (String(values[r][idx] || '').trim() === String(rootApptId).trim()) return { sh, rowIndex:r+1, map, rowVals: values[r] }; }
  return null;
}

function rp_setOrderTotal_Master_Safe_(rootApptId, value, allowOverride, masterRowIndex) {
  try {
    var val = Math.round(Number(value || 0) * 100) / 100;
    if (!(val > 0)) return { ok:false, updated:false, reason:'AMOUNT_NOT_POSITIVE' };
    var m;
    if (masterRowIndex) { m = rp_getMasterRowByIndex_(masterRowIndex); } else { m = rp_findMasterRowByRootApptId_(rootApptId); if (!m) return { ok:false, updated:false, reason:'MASTER_ROW_NOT_FOUND' }; }
    var cOT0 = m.map['Order Total'];
    if (cOT0 == null) return { ok:false, updated:false, reason:'ORDER_TOTAL_HEADER_MISSING' };
    var cPTD0 = rp_pick0(m.map, 'Paid-to-Date','Paid-To-Date','Paid to Date','Paid-to-date');
    var cRB0  = rp_pick0(m.map, 'Remaining Balance','Balance');
    var prev, ptd = 0;
    if (cPTD0 >= 0) { var readMin = Math.min(cOT0, cPTD0) + 1; var readSpan = Math.max(cOT0, cPTD0) - Math.min(cOT0, cPTD0) + 1; var blk = m.sh.getRange(m.rowIndex, readMin, 1, readSpan).getValues()[0]; prev = blk[(cOT0 + 1)  - readMin]; ptd  = rp_num_(blk[(cPTD0 + 1) - readMin]); } else { prev = m.sh.getRange(m.rowIndex, cOT0 + 1).getValue(); }
    if (prev && !allowOverride) return { ok:true, updated:false, value:prev, prev };
    var targets = [{ col: cOT0 + 1, val }];
    if (cPTD0 >= 0 && cRB0 >= 0) { var newRB = Math.max(0, val - ptd); targets.push({ col: cRB0 + 1, val: newRB }); }
    targets.sort(function(a,b){ return a.col - b.col; });
    var runs = [], cur = null;
    for (var i = 0; i < targets.length; i++) { var t = targets[i]; if (!cur) { cur = { start: t.col, vals: [t.val] }; } else if (t.col === cur.start + cur.vals.length) { cur.vals.push(t.val); } else { runs.push(cur); cur = { start: t.col, vals: [t.val] }; } }
    if (cur) runs.push(cur);
    for (var r = 0; r < runs.length; r++) { var run = runs[r]; m.sh.getRange(m.rowIndex, run.start, 1, run.vals.length).setValues([run.vals]); }
    return { ok:true, updated:true, value:val, prev };
  } catch (e) { return { ok:false, updated:false, reason: (e && e.message) ? e.message : String(e) }; }
}

function rp_auditOrderTotalOnLedger_(row, info) {
  rp_updateLedgerRow_(row, { 'OrderTotalSet': !!(info && info.set), 'OrderTotalValue': (info && info.value != null) ? info.value : '', 'OrderTotalSource': (info && info.source) || '', 'OrderTotalTarget': (info && info.target) || '', 'OrderTotalOldValue': (info && info.prev != null) ? info.prev : '' });
}


/*** === AR HELPERS === ***/
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
  const f1 = ensure(root, topName); const f2 = ensure(f1, yyyy); const f3 = ensure(f2, mm);
  return f3;
}


/*** === MASTER / ORDERS WRITE-BACKS === ***/
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

function rp_applyReceiptToMaster({ masterRowIndex, amount, when } = {}) {
  if (!masterRowIndex || masterRowIndex < 2 || !amount) throw new Error('Usage: rp_applyReceiptToMaster({masterRowIndex: <row>, amount: 50, when:new Date()})');
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet "00_Master Appointments".');
  let header = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0] || [];
  let H = rp_hIndex_(header);
  let cPTD = rp_pick(H, 'Paid-to-Date', 'Paid-To-Date', 'Paid to Date', 'Paid-to-date');
  if (!cPTD) { sh.getRange(1, sh.getLastColumn() + 1).setValue('Paid-to-Date'); header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]; H = rp_hIndex_(header); cPTD = H['Paid-to-Date']; }
  let cLPD = rp_pick(H, 'Last Payment Date', 'LastPaymentDate');
  if (!cLPD) { sh.getRange(1, sh.getLastColumn() + 1).setValue('Last Payment Date'); header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]; H = rp_hIndex_(header); cLPD = H['Last Payment Date']; }
  const H2 = rp_hIndex_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  const cRB = rp_pick(H2, 'Remaining Balance', 'Balance') || 0;
  const cOT = rp_pick(H2, 'Order Total', 'Order Total ') || 0;
  const row = masterRowIndex;
  let paidToDate = 0, orderTotal = 0;
  if (cOT) { const readMin = Math.min(cPTD, cOT); const readSpan = Math.max(cPTD, cOT) - readMin + 1; const block = sh.getRange(row, readMin, 1, readSpan).getValues()[0]; paidToDate = rp_num_(block[cPTD - readMin]); orderTotal = rp_num_(block[cOT  - readMin]); } else { paidToDate = rp_num_(sh.getRange(row, cPTD).getValue()); }
  const newPaid = paidToDate + rp_num_(amount);
  const whenVal = when || new Date();
  const targets = [{ col: cPTD, val: newPaid }, { col: cLPD, val: whenVal }];
  let newBal;
  if (cOT && cRB) { newBal = Math.max(0, orderTotal - newPaid); targets.push({ col: cRB, val: newBal }); }
  targets.sort((a, b) => a.col - b.col);
  const runs = []; let cur = null;
  for (const t of targets) { if (!cur) { cur = { start: t.col, vals: [t.val] }; } else if (t.col === cur.start + cur.vals.length) { cur.vals.push(t.val); } else { runs.push(cur); cur = { start: t.col, vals: [t.val] }; } }
  if (cur) runs.push(cur);
  runs.forEach(r => { sh.getRange(row, r.start, 1, r.vals.length).setValues([r.vals]); });
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
    const ap = String(r[cAppt]||'').trim(); if (ap !== rootApptId) continue;
    const t  = String(r[cType]||'').toUpperCase(); if (!(t.includes('RECEIPT') || t === 'DR' || t === 'SR')) continue;
    const status = cStatus != null ? String(r[cStatus] || '').toUpperCase().trim() : ''; if (status === 'VOID' || status === 'REPLACED' || status === 'DRAFT') continue;
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
  if (!cGross) { sh.getRange(1, sh.getLastColumn()+1).setValue('Cash-in (Gross)'); const h2 = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]; cGross = rp_hIndex_(h2)['Cash-in (Gross)']; }
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
    const t = String(r[cType]||'').toUpperCase(); if (!(t.includes('RECEIPT') || t === 'DR' || t === 'SR')) continue;
    const status = cStatus != null ? String(r[cStatus] || '').toUpperCase().trim() : ''; if (status === 'VOID' || status === 'REPLACED' || status === 'DRAFT') continue;
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
  if (!cStage) { sh.getRange(1, lc + 1).setValue('Sales Stage'); header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]; H = (function(hdr){ const m={}; hdr.forEach((h,i)=>{ const k=String(h||'').trim(); if (k) m[k]=i+1; }); return m; })(header); cStage = H['Sales Stage']; }
  const cell = sh.getRange(masterRowIndex, cStage);
  const cur = String(cell.getDisplayValue() || '').trim();
  if (!allowOverride && cur) return { ok:true, updated:false, prev:cur, value:cur };
  cell.setValue(value);
  return { ok:true, updated:true, value, prev:cur };
}


/*** === TAX + FEE HELPERS === ***/
function rp_round2(n){ return Math.round(Number(n) * 100) / 100; }

function rp_getTaxRate_(brand){
  try {
    const { ss } = rp_getLedgerTarget();
    const sh = ss.getSheetByName('Current Tax');
    if (!sh) { Logger.log('rp_getTaxRate_: sheet "Current Tax" not found. Returning 0.'); return 0; }
    const lr = sh.getLastRow();
    if (lr < 2) return 0;
    const data = sh.getRange(2, 1, lr - 1, 2).getValues();
    const brandNorm = String(brand || '').trim().toLowerCase();
    for (let i = 0; i < data.length; i++){ if (String(data[i][0]).trim().toLowerCase() === brandNorm){ return Number(data[i][1]) / 100; } }
    Logger.log('rp_getTaxRate_: brand [' + brand + '] not found. Returning 0.');
    return 0;
  } catch(e){ Logger.log('rp_getTaxRate_ error: ' + e.message); return 0; }
}


/*** === DEBUG / DIAGNOSTICS === ***/
function rp_debugConfig() {
  const show = (label, list) => { const hit = rp_propOneOf_(list || [], {label}); return { label, tried:list, resolvedKey: hit.key || '(none)', valuePreview: hit.value ? (hit.value.slice(0,6) + '…') : '' }; };
  const report = { ledgerFile: show('LEDGER_FILE_ID', RP_KEY_ALIASES.LEDGER_FILE_ID), ledgerSheet: show('LEDGER_SHEET_NAME', RP_KEY_ALIASES.LEDGER_SHEET_NAME), hpOrders: show('ORDERS_HPUSA_FILE_ID', RP_KEY_ALIASES.ORDERS_HPUSA_FILE_ID), vvsOrders: show('ORDERS_VVS_FILE_ID', RP_KEY_ALIASES.ORDERS_VVS_FILE_ID), ordTab: show('ORDERS_TAB_COMMON', RP_KEY_ALIASES.ORDERS_TAB_COMMON), hpTab: show('ORDERS_HPUSA_TAB', RP_KEY_ALIASES.ORDERS_HPUSA_TAB), vvsTab: show('ORDERS_VS_TAB', RP_KEY_ALIASES.ORDERS_VVS_TAB), arHP: show('AR_HPUSA_ROOT_ID', RP_KEY_ALIASES.AR_HPUSA_ROOT_ID), arVVS: show('AR_VVS_ROOT_ID', RP_KEY_ALIASES.AR_VVS_ROOT_ID), feesJson: show('FEES_JSON', RP_KEY_ALIASES.FEES_JSON), feesTab: show('FEES_TAB_NAME', RP_KEY_ALIASES.FEES_TAB_NAME) };
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

function updateHPUSATemplates() {
  const p = PropertiesService.getScriptProperties();
  p.setProperty('HPUSA_SR_TEMPLATE_ID', '1cmFtxmTQ2skVYCD9IwPATJ96qZQkDLWay1hYxV5r9Bg');
  p.setProperty('HPUSA_DR_TEMPLATE_ID', '1PkK4acFY4XWm6ZmqmPyim-cBGeek6v7ruDDaNdugg3w');
  Logger.log('SR: ' + p.getProperty('HPUSA_SR_TEMPLATE_ID'));
  Logger.log('DR: ' + p.getProperty('HPUSA_DR_TEMPLATE_ID'));
}

function updateVVSTemplates() {
  const p = PropertiesService.getScriptProperties();
  p.setProperty('VVS_SR_TEMPLATE_ID', '1oK83QSRhMMGAZawTWUm6G9ibUP7ovi8MXv_0sSn36EE');
  p.setProperty('VVS_DR_TEMPLATE_ID', '1hkYkS3Vk2hPM7WzganZtDNWqvBPsbN7dr0EFEqi_o4s');
  p.setProperty('VVS_DI_TEMPLATE_ID', '1hU59NKqZ_ffAe3kIRataes3NaXDDknCizvU6uzKWOSE');
  p.setProperty('VVS_SI_TEMPLATE_ID', '1PnUBVWeMwJZHSjppoT4tpsCoJSRQUjfz_bBbbiArKIg');
  Logger.log('Done VVS templates.');
}

function getHPUSAInvoiceTemplateIds() {
  const diId = rp_getTemplateIdFor('HPUSA', 'Deposit Invoice');
  const siId = rp_getTemplateIdFor('HPUSA', 'Sales Invoice');

  Logger.log('Deposit Invoice HPUSA: ' + diId);
  Logger.log('Sales Invoice HPUSA:   ' + siId);

  return { depositInvoice: diId, salesInvoice: siId };
}



/*** === TEMPLATE ID MANAGEMENT (v8.8) === ***/
function rp_setAllTemplateIds(config) {
  const p = PropertiesService.getScriptProperties();
  const codes = ['DR','SR','DI','SI'];
  const suffixes = ['TAX','NOTAX'];
  const brands = Object.keys(config || {});
  const set = {};
  const missing = [];

  for (const brand of brands) {
    const bConf = config[brand] || {};
    for (const code of codes) {
      for (const sfx of suffixes) {
        const mapKey = `${code}_${sfx}`;
        const propKey = `${brand}_${mapKey}_TEMPLATE_ID`;
        const val = bConf[mapKey];
        if (val && String(val).trim()) {
          set[propKey] = String(val).trim();
        } else {
          missing.push(propKey);
        }
      }
    }
  }

  if (Object.keys(set).length) {
    p.setProperties(set);
    Logger.log('[rp_setAllTemplateIds] Set %d keys:\n%s', Object.keys(set).length, JSON.stringify(set, null, 2));
  }
  if (missing.length) {
    Logger.log('[rp_setAllTemplateIds] Missing (skipped): %s', missing.join(', '));
  }
  return { set: Object.keys(set), missing };
}

function rp_checkTemplateIds() {
  const p = PropertiesService.getScriptProperties();
  const brands = ['HPUSA','VVS'];
  const codes  = ['DR','SR','DI','SI'];
  const sfxs   = ['TAX','NOTAX'];
  const report = { ok:[], missing:[] };

  for (const brand of brands) {
    for (const code of codes) {
      for (const sfx of sfxs) {
        const key = `${brand}_${code}_${sfx}_TEMPLATE_ID`;
        const val = p.getProperty(key);
        if (val && val.trim()) report.ok.push(key);
        else report.missing.push(key);
      }
    }
  }

  Logger.log('✅ Configured (%d):\n%s', report.ok.length, report.ok.join('\n'));
  Logger.log('❌ Missing (%d):\n%s', report.missing.length, report.missing.join('\n'));
  return report;
}

/*** === PROJECT #21: INVOICE-BEFORE-RECEIPT GATE === ***/

function rp_checkInvoiceBeforeReceipt_({ anchorType, rootApptId, soNumber, docType } = {}) {
  const dt = String(docType || '').toLowerCase().replace(/\s+/g, ' ').trim();

  // Chỉ kiểm tra Sales Receipt — Deposit Receipt được miễn
  if (!(dt.includes('sales') && dt.includes('receipt'))) {
    return { ok: true };
  }

  const { sh } = rp_getLedgerTarget();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();

  if (lr < 2) {
    return {
      ok: false,
      reason: 'NO_INVOICE_FOUND',
      message: 'No records found in the ledger.\n\nPlease create a Sales Invoice before generating a Sales Receipt.'
    };  
  }

  const head = rp_getHeaderRowCached_(sh);
  const H = {}; head.forEach((h, i) => H[h] = i);

  const cType   = H['DocType'];
  const cAppt   = H['RootApptID'];
  const cSO     = H['SO#'];
  const cStatus = H['DocStatus'] != null ? H['DocStatus'] : H['Status'];

  if (cType == null) {
    return {
      ok: false,
      reason: 'LEDGER_SCHEMA_ERROR',
      message: 'Lỗi cấu trúc ledger: không tìm thấy cột DocType.'
    };
  }

  const start = rp_scanWindowStart_(lr);
  const vals  = sh.getRange(start, 1, lr - start + 1, lc).getValues();

  for (let i = 0; i < vals.length; i++) {
    const r = vals[i];
    const rowDocType = String(r[cType] || '').toLowerCase().replace(/\s+/g, ' ').trim();

    if (!(rowDocType.includes('sales') && rowDocType.includes('invoice'))) continue;

    // Bỏ qua nếu VOID hoặc DRAFT (Draft chưa chính thức, không được tính)
    if (cStatus != null) {
      const status = String(r[cStatus] || '').toUpperCase().trim();
      if (status === 'VOID' || status === 'DRAFT') continue; // ← THÊM DRAFT
    }

    const isMatch = String(anchorType || '').toUpperCase() === 'SO'
      ? rp_soEq(r[cSO], soNumber)
      : String(r[cAppt] || '').trim() === String(rootApptId || '').trim();

    if (isMatch) {
      return { ok: true, matchedDocType: r[cType] };
    }
  }

  return {
    ok: false,
    reason: 'NO_INVOICE_FOUND',
    message:
      '⚠️ Unable to create Sales Receipt.\n\n' +
      'A valid Sales Invoice (Status: Issued) is required before creating a Sales Receipt.\n\n' +
      'If you have a Draft Invoice, please re-open it and change Status to "Issue now" first.'
      // ↑ Hướng dẫn cụ thể: Draft không đủ, phải Issue
  };
}


function rp_validateDocTypePrerequisite(payload) {
  try {
    const { anchorType, rootApptId, soNumber, docType } = payload || {};
    return rp_checkInvoiceBeforeReceipt_({ anchorType, rootApptId, soNumber, docType });
  } catch (e) {
    Logger.log('[rp_validateDocTypePrerequisite] ERROR: ' + (e && e.stack ? e.stack : e));
    return { ok: false, reason: 'SERVER_ERROR', message: e.message || String(e) };
  }
}

/*** === PROJECT #22: REFER A FRIEND === ***/

/**
 * Ghi thông tin referral vào Master sheet (00_Master Appointments)
 * Tìm hoặc tạo các cột: Referral Name, Referral Discount, Referral Date
 */
function rp_applyReferralToClientStatus_({ masterRowIndex, rootApptId, referralName, referralDiscount, submittedAt } = {}) {
  if (!masterRowIndex && !rootApptId) {
    Logger.log('[rp_applyReferralToClientStatus_] No anchor provided, skipping.');
    return { ok: false, reason: 'NO_ANCHOR' };
  }

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet "00_Master Appointments".');

  // Xác định row
  let rowIndex = masterRowIndex;
  if (!rowIndex || rowIndex < 2) {
    const m = rp_findMasterRowByRootApptId_(rootApptId);
    if (!m) throw new Error('Master row not found for RootApptID: ' + rootApptId);
    rowIndex = m.rowIndex;
  }

  // Đọc header hiện tại
  let header = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
  let H = rp_hIndex_(header);

  // Tạo cột nếu chưa có
  function ensureCol(colName) {
    let c = H[colName];
    if (!c) {
      const nextCol = sh.getLastColumn() + 1;
      sh.getRange(1, nextCol).setValue(colName);
      header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      H = rp_hIndex_(header);
      c = H[colName];
    }
    return c;
  }

  const cReferralName     = ensureCol('Referral Name');
  const cReferralDiscount = ensureCol('Referral Discount');
  const cReferralDate     = ensureCol('Referral Date');

  // Ghi giá trị
  const dateVal = submittedAt instanceof Date ? submittedAt : new Date();
  sh.getRange(rowIndex, cReferralName).setValue(String(referralName || '').trim());
  sh.getRange(rowIndex, cReferralDiscount).setValue(Number(referralDiscount || 100));
  sh.getRange(rowIndex, cReferralDate).setValue(dateVal);

  Logger.log('[rp_applyReferralToClientStatus_] Wrote referral: row=' + rowIndex + ' name=' + referralName + ' discount=' + referralDiscount);

  // ── Sync lên Client Status Report ──
  try {

    const m = masterRowIndex
      ? rp_getMasterRowByIndex_(masterRowIndex)
      : rp_findMasterRowByRootApptId_(rootApptId);

    if (!m) throw new Error('Master row not found for CS sync');

    const csUrlIdx = rp_pick0(m.sh
      ? rp_hIndex_(m.sh.getRange(1,1,1,m.sh.getLastColumn()).getValues()[0])
      : {}, 'Client Status Report URL');

    const csUrl = csUrlIdx >= 0
      ? String(m.sh.getRange(m.rowIndex, csUrlIdx).getValue() || '').trim()
      : '';

    if (csUrl) {
      const csId = rp_fileIdFromUrl(csUrl);
      if (csId) {
        const csSS  = SpreadsheetApp.openById(csId);
        const csSh  = csSS.getSheetByName('Client Status');
        if (csSh) {
          const referralText = referralName
            ? ('Yes — ' + referralName + ' (−$' + Number(referralDiscount||100) + ')')
            : 'Yes';
          rp_updateClientStatusSnapshotCell_(csSh, 'Refer a Friend:', referralText);
          Logger.log('[rp_applyReferralToClientStatus_] Synced to Client Status Report');
        }
      }
    }
  } catch(e) {
    Logger.log('[rp_applyReferralToClientStatus_] CS sync warning: ' + e.message);
  }

  return { ok: true, rowIndex, referralName, referralDiscount };
}


/**
 * Đọc thông tin referral hiện tại của 1 row — dùng cho rp_init() nếu cần prefill
 */
function rp_getReferralForMasterRow_(masterRowIndex) {
  if (!masterRowIndex || masterRowIndex < 2) return null;
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('00_Master Appointments');
    if (!sh) return null;
    const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const H = rp_hIndex_(header);
    const cName = H['Referral Name'], cDisc = H['Referral Discount'], cDate = H['Referral Date'];
    if (!cName) return null;
    const name = String(sh.getRange(masterRowIndex, cName).getValue() || '').trim();
    const disc = cDisc ? Number(sh.getRange(masterRowIndex, cDisc).getValue() || 0) : 0;
    const date = cDate ? sh.getRange(masterRowIndex, cDate).getValue() : null;
    return name ? { referralName: name, referralDiscount: disc, referralDate: date } : null;
  } catch(_) { return null; }
}

/**
 * Kiểm tra referral đã được dùng cho anchor này chưa
 * Scan ledger tìm row có ReferralEnabled = true cùng anchor
 */
function rp_isReferralAlreadyUsed_({ anchorType, rootApptId, soNumber } = {}) {
  try {
    const { sh } = rp_getLedgerTarget();
    const lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (lr < 2) return { used: false };

    const head = rp_getHeaderRowCached_(sh);
    const H = {}; head.forEach((h, i) => H[h] = i);

    const cAppt  = H['RootApptID'];
    const cSO    = H['SO#'];
    const cRef   = H['ReferralEnabled'];
    const cName  = H['ReferralName'];
    const cDisc  = H['ReferralDiscount'];
    const cStatus= H['DocStatus'] != null ? H['DocStatus'] : H['Status'];

    if (cRef == null) return { used: false };

    const start = rp_scanWindowStart_(lr);
    const vals  = sh.getRange(start, 1, lr - start + 1, lc).getValues();

    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];

      // Bỏ qua row VOID
      if (cStatus != null) {
        const status = String(r[cStatus] || '').toUpperCase().trim();
        if (status === 'VOID') continue;
      }

      // Kiểm tra ReferralEnabled = true
      const refVal = r[cRef];
      if (refVal !== true && String(refVal).toLowerCase() !== 'true') continue;

      // Kiểm tra cùng anchor
      const isMatch = String(anchorType || '').toUpperCase() === 'SO'
        ? rp_soEq(r[cSO], soNumber)
        : String(r[cAppt] || '').trim() === String(rootApptId || '').trim();

      if (isMatch) {
        return {
          used:     true,
          name:     cName != null ? String(r[cName]  || '') : '',
          discount: cDisc != null ? Number(r[cDisc]  || 0)  : 100
        };
      }
    }

    return { used: false };
  } catch(e) {
    Logger.log('[rp_isReferralAlreadyUsed_] ERROR: ' + e.message);
    return { used: false };
  }
}

/**
 * Tìm dòng có label trong cột A hoặc C của Client Status sheet
 * và ghi giá trị vào cột B hoặc D tương ứng
 */
function rp_updateClientStatusSnapshotCell_(sh, label, value) {
  const rowsToScan = Math.min(sh.getLastRow() || 50, 50);
  if (rowsToScan <= 0) return false;

  const values = sh.getRange(1, 1, rowsToScan, 4).getValues();

  // ── Chuẩn hóa: tìm cả có dấu ":" lẫn không có ──
  const normalize = s => String(s || '').trim().replace(/:+$/, '').toLowerCase();
  const needle    = normalize(label);

  for (let i = 0; i < rowsToScan; i++) {
    const labA = normalize(values[i][0]);
    const labC = normalize(values[i][2]);

    if (labA === needle) {
      sh.getRange(i + 1, 2).setValue(value);
      Logger.log('[rp_updateClientStatusSnapshotCell_] Wrote "' + label + '" → row ' + (i+1) + ' col B = ' + value);
      return true;
    }
    if (labC === needle) {
      sh.getRange(i + 1, 4).setValue(value);
      Logger.log('[rp_updateClientStatusSnapshotCell_] Wrote "' + label + '" → row ' + (i+1) + ' col D = ' + value);
      return true;
    }
  }

  Logger.log('[rp_updateClientStatusSnapshotCell_] Label "' + label + '" not found');
  return false;
}

/*** === DIAGNOSTIC v8.8 === ***/
function rp_diagnoseSetup() {
  const results = [];
  const log = (label, status, detail) => {
    results.push({ label, status, detail });
    Logger.log('[%s] %s — %s', status === 'OK' ? '✅' : '❌', label, detail);
  };

  // 1. Kiểm tra Ledger file
  try {
    const { ss, sh, resolved } = rp_getLedgerTarget();
    log('Ledger File', 'OK', 'File ID resolved via key: ' + resolved.ledgerFileKey);
    log('Ledger Sheet', sh ? 'OK' : 'WARN',
        sh ? 'Sheet "' + sh.getName() + '" found (key: ' + resolved.ledgerSheetKey + ')' : 'Sheet not found — using first sheet');
    const lr = sh.getLastRow();
    log('Ledger Rows', 'OK', lr + ' rows total');

    // Kiểm tra các cột bắt buộc
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v).trim());
    const required = ['PAYMENT_ID', 'Brand', 'RootApptID', 'SO#', 'DocType', 'AmountGross'];
    const missing = required.filter(h => !head.includes(h));
    log('Ledger Headers', missing.length === 0 ? 'OK' : 'WARN',
        missing.length === 0 ? 'All required columns present' : 'Missing: ' + missing.join(', '));
  } catch (e) {
    log('Ledger File', 'ERROR', e.message);
  }

  // 2. Kiểm tra Master sheet
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(RP_MASTER_SHEET);
    log('Master Sheet', sh ? 'OK' : 'ERROR',
        sh ? '"' + RP_MASTER_SHEET + '" found, ' + sh.getLastRow() + ' rows' : 'Sheet not found!');
  } catch (e) {
    log('Master Sheet', 'ERROR', e.message);
  }

  // 3. Kiểm tra Template IDs cho HPUSA và VVS
  const p = PropertiesService.getScriptProperties();
  const brands = ['HPUSA', 'VVS'];
  const codes  = ['DR', 'SR', 'DI', 'SI'];
  const sfxs   = ['TAX', 'NOTAX'];
  for (const brand of brands) {
    for (const code of codes) {
      for (const sfx of sfxs) {
        const key = `${brand}_${code}_${sfx}_TEMPLATE_ID`;
        const val = p.getProperty(key);
        log('Template: ' + key, val ? 'OK' : 'MISSING', val ? val.slice(0, 12) + '…' : '(not set)');
      }
    }
  }

  // 4. Kiểm tra Orders file IDs
  const ordersKeys = [
    ['HPUSA Orders', RP_KEY_ALIASES.ORDERS_HPUSA_FILE_ID],
    ['VVS Orders',   RP_KEY_ALIASES.ORDERS_VVS_FILE_ID],
  ];
  for (const [label, aliases] of ordersKeys) {
    const res = rp_propOneOf_(aliases);
    log(label + ' File', res.value ? 'OK' : 'MISSING',
        res.value ? 'key=' + res.key + ' id=' + res.value.slice(0,12)+'…' : 'Not configured');
  }

  // 5. Kiểm tra SO Root Folders
  const folderKeys = [
    ['HPUSA SO Root', RP_KEY_ALIASES.HPUSA_SO_ROOT_FOLDER_ID],
    ['VVS SO Root',   RP_KEY_ALIASES.VVS_SO_ROOT_FOLDER_ID],
  ];
  for (const [label, aliases] of folderKeys) {
    const res = rp_propOneOf_(aliases);
    if (res.value) {
      try {
        const f = DriveApp.getFolderById(res.value);
        log(label, 'OK', f.getName());
      } catch(e) {
        log(label, 'ERROR', 'Cannot open folder: ' + e.message);
      }
    } else {
      log(label, 'MISSING', 'Not configured');
    }
  }

  // 6. Kiểm tra Tax sheet
  try {
    const { ss } = rp_getLedgerTarget();
    const taxSh = ss.getSheetByName('Current Tax');
    if (taxSh) {
      const rows = taxSh.getRange(2, 1, Math.max(1, taxSh.getLastRow()-1), 2).getValues();
      log('Current Tax Sheet', 'OK', rows.map(r => r[0]+':'+r[1]+'%').join(', '));
    } else {
      log('Current Tax Sheet', 'WARN', 'Sheet "Current Tax" not found — tax = 0%');
    }
  } catch(e) {
    log('Current Tax Sheet', 'ERROR', e.message);
  }

  const errors   = results.filter(r => r.status === 'ERROR');
  const warnings = results.filter(r => r.status === 'WARN' || r.status === 'MISSING');
  Logger.log('\n========== SUMMARY ==========');
  Logger.log('✅ OK:      %d', results.filter(r => r.status === 'OK').length);
  Logger.log('⚠️  WARN:   %d', warnings.length);
  Logger.log('❌ ERRORS: %d', errors.length);
  if (errors.length)   Logger.log('ERRORS:\n' + errors.map(r => '  • ' + r.label + ': ' + r.detail).join('\n'));
  if (warnings.length) Logger.log('WARNINGS:\n' + warnings.map(r => '  • ' + r.label + ': ' + r.detail).join('\n'));

  return results;
}

function rp_getLastInvoiceTotalFromLedger_({ hasSO, soNumber, rootApptId } = {}) {
  const { sh } = rp_getLedgerTarget();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return 0;

  const head = rp_getHeaderRowCached_(sh);
  const H = {}; head.forEach((h, i) => H[h] = i);

  const cAppt     = H['RootApptID'];
  const cSO       = H['SO#'];
  const cInvTotal = H['InvoiceTotal'];
  const cStatus   = H['DocStatus'] != null ? H['DocStatus'] : H['Status'];

  if (cInvTotal == null) return 0;

  const start = rp_scanWindowStart_(lr);
  const vals  = sh.getRange(start, 1, lr - start + 1, lc).getValues();

  for (let i = vals.length - 1; i >= 0; i--) {
    const r = vals[i];

    const match = hasSO
      ? rp_soEq(r[cSO], soNumber)
      : String(r[cAppt] || '').trim() === String(rootApptId || '').trim();
    if (!match) continue;

    // Bỏ qua VOID / REPLACED / DRAFT
    if (cStatus != null) {
      const status = String(r[cStatus] || '').toUpperCase().trim();
      if (status === 'VOID' || status === 'REPLACED' || status === 'DRAFT') continue;
    }

    const invTotal = Number(r[cInvTotal] || 0);
    if (invTotal > 0) return invTotal;
  }
  return 0;
}


function rp_getTemplateLinks() {
  const p = PropertiesService.getScriptProperties();
  const brands = ['HPUSA', 'VVS'];
  const codes  = ['DR', 'SR', 'DI', 'SI'];
  const sfxs   = ['TAX', 'NOTAX'];

  const LABELS = {
    DR: 'Deposit Receipt',
    SR: 'Sales Receipt',
    DI: 'Deposit Invoice',
    SI: 'Sales Invoice',
    TAX:   'w/ Tax',
    NOTAX: 'No Tax'
  };

  const results = [];

  for (const brand of brands) {
    for (const code of codes) {
      for (const sfx of sfxs) {
        const key = `${brand}_${code}_${sfx}_TEMPLATE_ID`;
        const fileId = p.getProperty(key);

        const entry = {
          brand,
          docType:  LABELS[code],
          taxMode:  LABELS[sfx],
          propKey:  key,
          fileId:   fileId || '',
          fileName: '',
          url:      ''
        };

        if (fileId && fileId.trim()) {
          try {
            const file = DriveApp.getFileById(fileId.trim());
            entry.fileName = file.getName();
            entry.url      = `https://docs.google.com/document/d/${fileId.trim()}/edit`;
          } catch (e) {
            entry.fileName = '⚠️ Cannot access file: ' + e.message;
          }
        } else {
          entry.fileName = '❌ Not configured';
        }

        results.push(entry);
        Logger.log('[%s] %s | %s | key=%s | id=%s | name=%s',
          brand, LABELS[code], LABELS[sfx], key,
          fileId || '(none)', entry.fileName
        );
      }
    }
  }

  // In bảng tóm tắt theo brand
  for (const brand of brands) {
    Logger.log('\n══════════ %s ══════════', brand);
    const rows = results.filter(r => r.brand === brand);
    for (const r of rows) {
      Logger.log('  %-20s %-8s → %s', r.docType, r.taxMode,
        r.url ? r.url : r.fileName
      );
    }
  }

  return results;
}

function setupAllTemplates() {
  const p = PropertiesService.getScriptProperties();

  // ★ HPUSA — 4 template đơn (không TAX/NOTAX, 1 template xử lý cả 2)
  p.setProperty('HPUSA_DI_TEMPLATE_ID', '11O2dHWzCzkjWPXBAMRMVkxoqfB9l1r59UPSAD8d3IVo');
  p.setProperty('HPUSA_DR_TEMPLATE_ID', '16Vwynd45r_Qtn-iHqooU8Hg2H4cdN308KCDGNB-zy8Y');
  p.setProperty('HPUSA_SI_TEMPLATE_ID', '1uE6Z02YaXtnQjJuuUq1eETvPggGuW3tIE3xRYqiYBA8');
  p.setProperty('HPUSA_SR_TEMPLATE_ID', '1Wjp0z3vTXUthMFbW90v2w1BgdnA3AZUCLPW2S6eFnCA');

  // VVS — giữ nguyên TAX/NOTAX
  rp_setAllTemplateIds({
    VVS: {
      DR_TAX:   '1hkYkS3Vk2hPM7WzganZtDNWqvBPsbN7dr0EFEqi_o4s',
      DR_NOTAX: '1VlZi7Ztn9tEPg8O4QlP2WvVkm1ZGTNLnVBFL1wi639o',
      SR_TAX:   '1oK83QSRhMMGAZawTWUm6G9ibUP7ovi8MXv_0sSn36EE',
      SR_NOTAX: '1Fy04ddyQw5vKQdZzTXLH9SLmPsea6WcAQPzOaf0szfw',
      DI_TAX:   '1hU59NKqZ_ffAe3kIRataes3NaXDDknCizvU6uzKWOSE',
      DI_NOTAX: '1IXXfnkzgrHt7d2-Oz0Rm0RqBuDeKjLB9Y_eYk9-j3dg',
      SI_TAX:   '1PnUBVWeMwJZHSjppoT4tpsCoJSRQUjfz_bBbbiArKIg',
      SI_NOTAX: '1YP1b0Lt2IcO04lB-ZwmE5suuGV8VgFrj45vWKkZaw-w',
    }
  });

  Logger.log('HPUSA_DI: ' + p.getProperty('HPUSA_DI_TEMPLATE_ID'));
  Logger.log('HPUSA_DR: ' + p.getProperty('HPUSA_DR_TEMPLATE_ID'));
  Logger.log('HPUSA_SI: ' + p.getProperty('HPUSA_SI_TEMPLATE_ID'));
  Logger.log('HPUSA_SR: ' + p.getProperty('HPUSA_SR_TEMPLATE_ID'));
}

function debugDocTabs() {
  // Thay bằng ID của 1 trong 4 template HPUSA
  const docId = '11O2dHWzCzkjWPXBAMRMVkxoqfB9l1r59UPSAD8d3IVo'; // DI template
  
  const doc = DocumentApp.openById(docId);
  
  // Kiểm tra tabs
  try {
    const tabs = doc.getTabs();
    Logger.log('Số tabs: ' + tabs.length);
    tabs.forEach((tab, i) => {
      Logger.log('Tab ' + i + ': id=' + tab.getId() + ' title=' + tab.getTitle());
    });
  } catch(e) {
    Logger.log('getTabs() error: ' + e.message);
  }
  
  // Thử export với các tab_id khác nhau
  const token = ScriptApp.getOAuthToken();
  const variants = ['t.0', 't.1', '0', '1'];
  
  for (const tabId of variants) {
    const url = 'https://docs.google.com/document/d/' + docId 
      + '/export?format=pdf&tab_id=' + tabId;
    try {
      const resp = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
      Logger.log('tab_id=' + tabId + ' → HTTP ' + resp.getResponseCode() 
        + ' size=' + resp.getBlob().getBytes().length);
    } catch(e) {
      Logger.log('tab_id=' + tabId + ' → ERROR: ' + e.message);
    }
  }
}

function setNewHPUSATemplateIds() {
  const p = PropertiesService.getScriptProperties();

  // ── Paste 4 ID mới vào đây ───────────────────────────────────
  const NEW_IDS = {
    HPUSA_DI_TEMPLATE_ID: '1VQzVvb2jVJWDdWEykjMqWkJNoISY5rGcaqA0XGf-V4E',  // Deposit Invoice
    HPUSA_DR_TEMPLATE_ID: '1hp8OW3MtLng4HP8-np__BdDB7kF-jCNPWVnnCwN_4Gk',  // Deposit Receipt
    HPUSA_SI_TEMPLATE_ID: '19cF01yMfjT2BHW1H_JKG9bXNNvrw5FOfbXiBz424tBE',  // Sales Invoice
    HPUSA_SR_TEMPLATE_ID: '1tPJvLtrS_6ByRM6IiUrgQvpKyeriJHXX4j5OmFuD3Uo',  // Sales Receipt
  };

  Logger.log('===== UPDATE HPUSA TEMPLATE IDs =====');

  Object.entries(NEW_IDS).forEach(([key, newId]) => {
    if (!newId || newId.includes('PASTE')) {
      Logger.log('⏭ ' + key + ': skip (chưa điền ID)');
      return;
    }
    const oldId = p.getProperty(key) || '(none)';
    p.setProperty(key, newId.trim());
    Logger.log('✅ ' + key);
    Logger.log('   OLD: ' + oldId);
    Logger.log('   NEW: ' + newId.trim());
  });

  // ── Verify sau khi set ───────────────────────────────────────
  Logger.log('\n===== VERIFY =====');
  const token = ScriptApp.getOAuthToken();

  ['DI','DR','SI','SR'].forEach(code => {
    const key   = 'HPUSA_' + code + '_TEMPLATE_ID';
    const docId = p.getProperty(key) || '';
    if (!docId) { Logger.log('❌ ' + key + ': not set'); return; }

    // Kiểm tra file accessible
    try {
      const file = DriveApp.getFileById(docId);
      Logger.log('📄 HPUSA ' + code + ': "' + file.getName() + '"');
    } catch(e) {
      Logger.log('❌ HPUSA ' + code + ': cannot access → ' + e.message);
      return;
    }

    // Kiểm tra tabs
    try {
      const tabs = DocumentApp.openById(docId).getTabs();
      if (tabs.length === 0) {
        Logger.log('   Tabs: ✅ No tabs (clean)');
      } else {
        Logger.log('   Tabs: ⚠️ ' + tabs.length + ' tab(s) → ' +
          tabs.map(t => '"' + t.getTitle() + '"').join(', ') +
          ' — cần remove tabs');
      }
    } catch(e) {
      Logger.log('   Tabs: ✅ No tabs feature (clean)');
    }

    // Test export PDF
    try {
      const url  = 'https://docs.google.com/document/d/' + docId + '/export?format=pdf';
      const resp = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });
      const kb = (resp.getBlob().getBytes().length / 1024).toFixed(0);
      Logger.log('   PDF: HTTP ' + resp.getResponseCode() + ' / ' + kb + ' KB ' +
        (resp.getResponseCode() === 200 ? '✅' : '❌'));
    } catch(e) {
      Logger.log('   PDF: ❌ ' + e.message);
    }
  });

  Logger.log('\n===== DONE =====');
}

function rp_measureDescColumn() {
  // Thay bằng ID của 1 doc đã được generate (không phải template)
  const docId = '1GTsUOIi1bxyxUlYwWESQJGVdDE_unVfH-FaKnr37rCg';
  const doc   = DocumentApp.openById(docId);
  const body  = doc.getBody();

  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.TABLE) continue;
    const t = el.asTable();
    if (t.getNumRows() < 1 || t.getRow(0).getNumCells() < 3) continue;
    const h0 = String(t.getRow(0).getCell(0).getText()).toUpperCase();
    if (!h0.includes('DESCRIPTION')) continue;

    const cell = t.getRow(0).getCell(0); // header cell
    const colWidth = cell.getWidth();    // đơn vị: points
    const fontSize = 8;                  // font size bạn đang dùng

    // Ước tính: 1 pt font ≈ 0.5–0.6 pt width mỗi ký tự (Cardo serif)
    const charWidthPt   = fontSize * 0.55;
    const charsPerLine  = Math.floor(colWidth / charWidthPt);

    Logger.log('Column width: %s pt', colWidth);
    Logger.log('Font size: %s pt', fontSize);
    Logger.log('Estimated chars per line: %s', charsPerLine);
    return { colWidth, fontSize, charsPerLine };
  }

  Logger.log('Description column not found.');
  return null;
}

/**
 * rp_hardResetApptPayments — v2
 * THAY THẾ hàm cũ cùng tên trong Payments.gs
 *
 * Thay đổi so với v1:
 * - BƯỚC 2: XÓA HẲN rows khỏi Ledger (file 400) thay vì chỉ set VOID
 *   Rows được xóa từ dưới lên trên để tránh index shift
 * - Giữ nguyên tất cả các bước khác (Master reset, Drive delete)
 */
function rp_hardResetApptPayments({ rootApptId, masterRowIndex, deletePdfFiles = false } = {}) {
  if (!rootApptId && !masterRowIndex) throw new Error('Cần rootApptId hoặc masterRowIndex');

  const results = {
    ledger:  { deleted: 0, rows: [] },
    master:  { cleared: [] },
    drive:   { deleted: 0, errors: [] },
  };

  // ══════════════════════════════════════════════
  // BƯỚC 1 — Tìm master row
  // ══════════════════════════════════════════════
  let masterRow;
  try {
    masterRow = masterRowIndex
      ? rp_getMasterRowByIndex_(masterRowIndex)
      : rp_findMasterRowByRootApptId_(rootApptId);
    if (!masterRow) throw new Error('Không tìm thấy master row');
    if (!rootApptId)    rootApptId    = masterRow.rootApptId;
    if (!masterRowIndex) masterRowIndex = masterRow.rowIndex;
  } catch (e) {
    throw new Error('[rp_hardResetApptPayments] ' + e.message);
  }

  Logger.log('[reset v2] RootApptID=%s masterRow=%s', rootApptId, masterRowIndex);

  // ══════════════════════════════════════════════
  // BƯỚC 2 — Xóa hẳn rows trên Ledger + thu thập PDF IDs
  // ══════════════════════════════════════════════
  const pdfFileIds = [];
  try {
    const { sh } = rp_getLedgerTarget();
    const lr = sh.getLastRow(), lc = sh.getLastColumn();

    if (lr >= 2) {
      const head = rp_getHeaderRowCached_(sh);
      const H = {}; head.forEach((h, i) => H[h] = i);

      const cAppt  = H['RootApptID'];
      const cPdfId = H['DocPDFID'];
      const cDocId = H['DocFileID'];

      if (cAppt == null) throw new Error('Không tìm thấy cột RootApptID trên Ledger');

      const start = rp_scanWindowStart_(lr);
      const vals  = sh.getRange(start, 1, lr - start + 1, lc).getValues();

      // Thu thập row numbers cần xóa (các row khớp rootApptId)
      // Lưu ý: index trong vals là 0-based, row thực tế = start + i
      const rowsToDelete = []; // row numbers thực tế trên sheet (1-based)

      for (let i = 0; i < vals.length; i++) {
        const r = vals[i];
        const appt = String(r[cAppt] || '').trim();
        if (appt !== String(rootApptId).trim()) continue;

        const rowNum = start + i; // row number thực tế trên sheet
        rowsToDelete.push(rowNum);
        results.ledger.rows.push(rowNum);

        // Thu thập PDF/Doc IDs để xóa trên Drive nếu cần
        if (cPdfId != null && r[cPdfId]) pdfFileIds.push(String(r[cPdfId]));
        if (cDocId != null && r[cDocId]) pdfFileIds.push(String(r[cDocId]));
      }

      // XÓA ROWS TỪ DƯỚI LÊN TRÊN để tránh index shift
      // (nếu xóa từ trên xuống, các row phía dưới sẽ bị lệch số)
      rowsToDelete.sort((a, b) => b - a); // sắp xếp giảm dần
      for (const rowNum of rowsToDelete) {
        try {
          sh.deleteRow(rowNum);
          results.ledger.deleted++;
          Logger.log('[reset v2] Deleted ledger row: %s', rowNum);
        } catch (e) {
          Logger.log('[reset v2] Failed to delete row %s: %s', rowNum, e.message);
        }
      }

      // Xóa cache header sau khi đã xóa rows để tránh cache stale
      try {
        const cache = CacheService.getUserCache();
        const key = 'HDR::' + sh.getParent().getId() + '::' + sh.getSheetId();
        cache.remove(key);
      } catch (_) {}
    }
  } catch (e) {
    Logger.log('[reset v2] Ledger error: ' + e.message);
  }

  // ══════════════════════════════════════════════
  // BƯỚC 3 — Xóa file PDF + Doc trên Drive (tuỳ chọn)
  // ══════════════════════════════════════════════
  if (deletePdfFiles && pdfFileIds.length) {
    const seen = new Set();
    pdfFileIds.forEach(fid => {
      if (!fid || seen.has(fid)) return;
      seen.add(fid);
      try {
        DriveApp.getFileById(fid).setTrashed(true);
        results.drive.deleted++;
        Logger.log('[reset v2] Trashed file: ' + fid);
      } catch (e) {
        results.drive.errors.push(fid + ': ' + e.message);
      }
    });
  }

  // ══════════════════════════════════════════════
  // BƯỚC 4 — Reset Master sheet
  // ══════════════════════════════════════════════
  try {
    const sh  = masterRow.sh;
    const H   = rp_hIndex_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);

    const cPTD           = rp_pick(H, 'Paid-to-Date', 'Paid-To-Date', 'Paid to Date', 'Paid-to-date');
    const cLPD           = rp_pick(H, 'Last Payment Date', 'LastPaymentDate');
    const cRB            = rp_pick(H, 'Remaining Balance', 'Balance');
    const cOT            = rp_pick(H, 'Order Total');
    const cCash          = rp_pick(H, 'Cash-in (Gross)');
    const cStage         = rp_pick(H, 'Sales Stage', 'SalesStage', 'Stage');
    const cSavedLines    = rp_pick(H, 'Saved Lines JSON', 'SavedLinesJSON', 'Saved Lines');
    const cOrderLines    = rp_pick(H, 'Order Lines (JSON)', 'OrderLines');
    const cLinesSubtotal = rp_pick(H, 'Lines Subtotal (Saved)', 'LinesSubtotal');
    const cSavedSubtotal = rp_pick(H, 'Saved Subtotal', 'SavedSubtotal');
    const cFolder        = rp_pick0(masterRow.map, 'PaymentsFolderURL');

    const masterUpdates = [
      { col: cPTD,           val: 0,  label: 'Paid-to-Date' },
      { col: cLPD,           val: '', label: 'Last Payment Date' },
      { col: cRB,            val: 0,  label: 'Remaining Balance' },
      { col: cOT,            val: 0,  label: 'Order Total' },
      { col: cCash,          val: 0,  label: 'Cash-in (Gross)' },
      { col: cStage,         val: '', label: 'Sales Stage' },
      { col: cSavedLines,    val: '', label: 'Saved Lines JSON' },
      { col: cOrderLines,    val: '', label: 'Order Lines (JSON)' },
      { col: cLinesSubtotal, val: '', label: 'Lines Subtotal (Saved)' },
      { col: cSavedSubtotal, val: '', label: 'Saved Subtotal' },
    ];

    if (cFolder >= 0) {
      masterUpdates.push({ col: cFolder + 1, val: '', label: 'PaymentsFolderURL' });
    }

    masterUpdates.forEach(u => {
      if (!u.col) return;
      sh.getRange(masterRowIndex, u.col).setValue(u.val);
      results.master.cleared.push(u.label);
    });

  } catch (e) {
    Logger.log('[reset v2] Master error: ' + e.message);
  }

  Logger.log('[reset v2] DONE: %s', JSON.stringify(results));
  try { if (typeof swInvalidatePaymentReadModelsAfterWrite_ === 'function') swInvalidatePaymentReadModelsAfterWrite_(null, 'Payment hard reset'); } catch (_) {}
  return results;
}

function runReset() {
  const result = rp_hardResetApptPayments({
    rootApptId:    'AP-20260308-002',  // ← ID ĐÚNG
    masterRowIndex: 396,
    deletePdfFiles: true,
  });

  Logger.log('Voided ledger rows: '  + result.ledger.voided);
  Logger.log('Master cleared: '      + result.master.cleared.join(', '));
  Logger.log('Drive files deleted: ' + result.drive.deleted);
}

function rp_findApptByCustomerName() {
  const customerName = 'Ana luisa Dela vega'; // ← tên khách

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  const vals = sh.getRange(1, 1, lr, lc).getDisplayValues();
  const map = {};
  vals[0].forEach((h, i) => map[String(h).trim()] = i);

  const cCust  = map['Customer Name'] ?? map['Customer'];
  const cAppt  = map['RootApptID'] ?? map['APPT_ID'];
  const cOT    = map['Order Total'];

  for (let r = 1; r < vals.length; r++) {
    const name = String(vals[r][cCust] || '').trim().toLowerCase();
    if (name.includes('ana') && name.includes('dela')) {
      Logger.log('Row=%s RootApptID=%s OT=%s Customer=%s',
        r+1,
        vals[r][cAppt],
        vals[r][cOT],
        vals[r][cCust]
      );
    }
  }
}

/**
 * rp_resetFromDialog — wrapper được gọi từ dialog UI
 * Cho phép reset toàn bộ dữ liệu thanh toán của 1 row từ giao diện Record Payment
 *
 * @param {Object} payload
 *   - rootApptId      {string}  RootApptID của row cần reset
 *   - masterRowIndex  {number}  Index của row trên sheet 00_Master Appointments
 *   - deletePdfFiles  {boolean} Có xóa file PDF/Doc trên Drive không (mặc định false)
 *   - confirmToken    {string}  Phải bằng "CONFIRM_RESET" — tránh gọi nhầm
 *
 * @returns {{ ok: boolean, voided: number, cleared: string[], deleted: number, error?: string }}
 */
function rp_resetFromDialog(payload) {
  try {
    const { rootApptId, masterRowIndex, deletePdfFiles, confirmToken } = payload || {};

    // Bảo vệ: bắt buộc phải có token xác nhận từ phía client
    if (confirmToken !== 'CONFIRM_RESET') {
      return { ok: false, error: 'Missing confirmation token.' };
    }

    if (!rootApptId && !masterRowIndex) {
      return { ok: false, error: 'Missing rootApptId or masterRowIndex.' };
    }

    Logger.log('[rp_resetFromDialog] rootApptId=%s masterRowIndex=%s deletePdf=%s',
      rootApptId, masterRowIndex, !!deletePdfFiles);

    const result = rp_hardResetApptPayments({
      rootApptId:     String(rootApptId || '').trim(),
      masterRowIndex: Number(masterRowIndex) || 0,
      deletePdfFiles: !!deletePdfFiles
    });

    return {
      ok:      true,
      voided:  result.ledger  && result.ledger.voided   || 0,
      cleared: result.master  && result.master.cleared  || [],
      deleted: result.drive   && result.drive.deleted   || 0,
      driveErrors: result.drive && result.drive.errors  || []
    };

  } catch (e) {
    Logger.log('[rp_resetFromDialog] ERROR: ' + (e && e.stack ? e.stack : e));
    return { ok: false, error: e.message || String(e) };
  }
}
