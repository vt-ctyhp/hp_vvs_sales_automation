/** 00_Canon.gs — Canonical terms & label-based helpers (v1)
 *  - One source of truth for headers & property keys
 *  - Robust sheet/column resolution by LABEL (no A1 literals)
 *  - Fast batched reads; tolerant synonyms; single duplicate SO guard
 */

/* ===== Canonical dictionary ===== */
const CANON = {
  PROPS: {
    MASTER_SHEET_NAME: 'MASTER_SHEET_NAME',
    ORDERS_TAB_NAME:   '301/302_TAB_NAME',
    HP_ORDERS_FILE:    'HPUSA_301_FILE_ID',
    VVS_ORDERS_FILE:   'VVS_302_FILE_ID',
    LEDGER_FILE_ID:    'LEDGER_FILE_ID',
    HP_CLIENTS_ROOT:   'HP_CLIENTS_ROOT_ID',
    VVS_CLIENTS_ROOT:  'VVS_CLIENTS_ROOT_ID',
    HP_SO_ROOT:        'HPUSA_SO_ROOT_FOLDER_ID',
    VVS_SO_ROOT:       'VVS_SO_ROOT_FOLDER_ID',
  },
  SHEETS: {
    MASTER_FALLBACKS: ['00_Master Appointments','00_Master','00 – Master','00 Master']
  },
  HEADERS: {
    so:               'SO#',
    brand:            'Brand',
    odooUrl:          'Odoo SO URL',
    linkedAt:         'SO Linked At',
    orderStatus:      'Custom Order Status',  // canonical everywhere
    designRequest:    'Design Request',
    shortTag:         'Short Tag',
    threeDTracker:    '3D Tracker',
    orderFolder:      'Order Folder',
    threeDFolder:     '05-3D Folder',
    clientFolder:     'Client Folder',
    soShortcutClient: 'SO Shortcut in Client',
    intakeFolder:     '00-Intake',
    apptId:           'APPT_ID',
    rootApptId:       'RootApptID',
    customer:         'Customer Name',
    email:            'EmailLower',
    phone:            'PhoneNorm',
    assignedRep:      'Assigned Rep',
    assistedRep:      'Assisted Rep'
  },
  // Accepted labels → resolve into the canonical names above
  SYN: {
    'SO#': ['SO Number','SO No','SO'],
    'Odoo SO URL': ['SO URL','Odoo Link'],
    'Custom Order Status': ['Order Status','Status'],
    'Customer Name': ['Customer','Name'],
    'EmailLower': ['Email','Customer Email'],
    'PhoneNorm': ['Phone','Customer Phone'],
    'RootApptID': ['Appt ID','ApptID','Appointment ID'],
    'Assigned Rep': ['Sales Rep','Rep','Owner'],
    '05-3D Folder': ['05 - 3D Folder','3D Folder'],
    '00-Intake': ['00 Intake','Intake Folder'],
    'Client Folder': ['Client Folder URL'],
    'SO Shortcut in Client': ['SO Shortcut','SO Shortcut in Client Folder'],
    '3D Tracker': ['3D Log','3D Sheet','3D Tracker URL']
  }
};

/* ===== Script Properties ===== */
function _prop_(k, def){ 
  const v = PropertiesService.getScriptProperties().getProperty(k); 
  return (v==null||v==='') ? (def||'') : v; 
}

/* ===== Sheets & headers ===== */
function getMasterSheet_(ss){
  const names = [ _prop_(CANON.PROPS.MASTER_SHEET_NAME, ''), ...CANON.SHEETS.MASTER_FALLBACKS ];
  for (const n of names){
    if (!n) continue;
    const sh = ss.getSheetByName(n);
    if (sh) return sh;
  }
  throw new Error('Could not find Master sheet. Tried: ' + names.filter(Boolean).join(' | '));
}

function headerMap_(sh){
  const last = Math.max(1, sh.getLastColumn());
  const hdr = sh.getRange(1,1,1,last).getDisplayValues()[0] || [];
  const H = {};
  hdr.forEach((h,i)=>{ const k = String(h||'').trim(); if (k) H[k] = i+1; });
  return H; // 1-based column numbers
}

function _synonymsFor_(canonLabel){
  return [canonLabel].concat(CANON.SYN[canonLabel] || []);
}

function headerIndexByCanon_(H, canonLabel){
  for (const name of _synonymsFor_(canonLabel)){
    if (H[name]) return H[name];
  }
  return 0; // not found
}

function ensureHeaders_(sh, canonLabels){
  const H = headerMap_(sh);
  let appended = false;
  canonLabels.forEach(canon => {
    const present = headerIndexByCanon_(H, canon);
    if (!present){
      sh.getRange(1, sh.getLastColumn()+1).setValue(canon);
      appended = true;
    }
  });
  return appended ? headerMap_(sh) : H;
}

function getCellByCanon_(sh, row, H, canonLabel){
  const col = headerIndexByCanon_(H, canonLabel);
  return col ? sh.getRange(row, col).getValue() : '';
}

function setCellByCanon_(sh, row, H, canonLabel, value){
  const col = headerIndexByCanon_(H, canonLabel);
  if (col) sh.getRange(row, col).setValue(value);
}

function coerceSOTextColumn_(sh, H){
  const col = headerIndexByCanon_(H, CANON.HEADERS.so);
  if (col) sh.getRange(1, col, sh.getMaxRows(), 1).setNumberFormat('@');
}

/* ===== Orders book & tab ===== */
function ordersFileIdForBrand_(brand){
  const b = String(brand||'').toUpperCase();
  const key = (b === 'HPUSA') ? CANON.PROPS.HP_ORDERS_FILE : CANON.PROPS.VVS_ORDERS_FILE;
  const id = _prop_(key, '').trim();
  if (!id) throw new Error('Missing Script Property: ' + key);
  return id;
}

function getOrdersSheet_(wb){
  const pinned = _prop_(CANON.PROPS.ORDERS_TAB_NAME, '').trim();
  if (pinned){
    const sh = wb.getSheetByName(pinned);
    if (sh) return sh;
  }
  // Fallback: fuzzy finder if property not set or tab renamed
  const hit = wb.getSheets().find(s => /(^|\b)(sales|customer\s*order|orders)(\b|$)/i.test(s.getName()));
  if (hit) return hit;
  throw new Error('Orders tab not found. Tried property "' + CANON.PROPS.ORDERS_TAB_NAME + '" or fuzzy match.');
}

/* ===== Business rules ===== */
function normalizeSON_(so){
  return String(so||'').replace(/^'+/,'').trim();
}

function existsSOInMaster_(sh, brand, so, skipRow){
  const H  = headerMap_(sh);
  const iBrand = headerIndexByCanon_(H, CANON.HEADERS.brand);
  const iSO    = headerIndexByCanon_(H, CANON.HEADERS.so);
  if (!iBrand || !iSO) return false;

  const last = sh.getLastRow();
  if (last < 2) return false;
  const rows = last - 1;
  const brandVals = sh.getRange(2, iBrand, rows, 1).getValues();
  const soVals    = sh.getRange(2, iSO,    rows, 1).getDisplayValues();
  const targetB = String(brand||'').toUpperCase().trim();
  const targetS = normalizeSON_(so);

  for (let i=0;i<rows;i++){
    const r = i+2;
    if (r === skipRow) continue;
    const b = String(brandVals[i][0]||'').toUpperCase().trim();
    const s = normalizeSON_(soVals[i][0]||'');
    if (b === targetB && s === targetS) return true;
  }
  return false;
}

/* ===== Utilities ===== */
function colToA1_(n){ let s=''; while(n>0){ n--; s=String.fromCharCode(65+(n%26))+s; n = Math.floor(n/26); } return s; }

/* ===== Canon public surface (prefixed to avoid collisions) ===== */
function headerMap__canon(sh){ return headerMap_(sh); }
function ensureHeaders__canon(sh, labels){ return ensureHeaders_(sh, labels); }
function getMasterSheet__canon(ss){ return getMasterSheet_(ss); }
function getOrdersSheet__canon(wb){ return getOrdersSheet_(wb); }
function coerceSOTextColumn__canon(sh, H){ return coerceSOTextColumn_(sh, H); }
function existsSOInMaster__canon(sh, brand, so, skipRow){ return existsSOInMaster_(sh, brand, so, skipRow); }

// --- Canon aliases (short names for direct calls) ---
const HMAP  = headerMap__canon;
const ENS   = ensureHeaders__canon;
const MAST  = getMasterSheet__canon;
const ORD   = getOrdersSheet__canon;
const SOFMT = coerceSOTextColumn__canon;
const SODUP = existsSOInMaster__canon;
