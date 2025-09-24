/** File: 03 - revision3d_server.gs — v1.1 (no‑risk optimizations)
 * Purpose: Identical behavior; faster I/O via batched row reads and targeted tracker reads.
 *
 * Provides:
 *  - open3DRevision()
 *  - rev3d_init()            // batched Master row read; targeted Tracker scan
 *  - previewRevOdooPaste()
 *  - submit3DRevision()      // batched Master row read; unchanged logging & summary
 *
 * Dependencies (unchanged):
 *  - append3DTrackerLog_(...)
 *  - copy3DTrackerToSO_(...)
 */

// ---------- Open dialog ----------
function open3DRevision(){
  const html = HtmlService.createHtmlOutputFromFile('dlg_revision3d_v1')
    .setTitle('3D Revision Request')
    .setWidth(650)
    .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, '3D Revision Request');
}

// ---------- Bootstrap (fast) ----------
function rev3d_init(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet "00_Master Appointments".');

  const r = ss.getActiveRange();
  if (!r || r.getSheet().getName() !== sh.getName() || r.getRow() < 2){
    throw new Error('Select a data row on "00_Master Appointments" and try again.');
  }
  const row = r.getRow();

  // Header + single-row snapshot (values + rich text) in batched calls
  const lastCol = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1,1,1,lastCol).getValues()[0] || [];
  const H = {}; header.forEach((h,i)=>{ const k=String(h||'').trim(); if(k) H[k]=i+1; });

  const rowRange = sh.getRange(row,1,1,lastCol);
  const rowVals  = rowRange.getValues()[0] || [];
  const rowRich  = rowRange.getRichTextValues()[0] || [];

  const idx = name => H[name] ? (H[name]-1) : -1;
  const getPlain = name => { const i=idx(name); return i>=0 ? rowVals[i] : ''; };
  const getLink  = name => {
    const i=idx(name); if (i<0) return '';
    const rtv = rowRich[i];
    if (rtv && typeof rtv.getLinkUrl === 'function'){
      const u = rtv.getLinkUrl();
      if (u) return String(u);
    }
    const v = rowVals[i];
    return v != null ? String(v) : '';
  };

  const brand  = String(getPlain('Brand') || '').toUpperCase();
  const soRaw  = String(getPlain('SO#') || '').trim();
  const so     = soRaw.replace(/^'/,'');
  const odoo   = getLink('Odoo SO URL') || getLink('SO URL') || getLink('Odoo Link') || '';
  const cust   = getPlain('Customer Name') || getPlain('Customer') || getPlain('Name') || '';
  const email  = getPlain('EmailLower') || getPlain('Email') || '';
  const phone  = getPlain('PhoneNorm')  || getPlain('Phone') || '';
  const threeDFolderLink = getPlain('05-3D Folder') || '';
  const orderFolderLink  = getPlain('Order Folder') || '';
  const trackerUrl       = getPlain('3D Tracker') || '';
  const trackerId        = _idFromUrlLoose_(trackerUrl);
  const masterLink       = ss.getUrl() + '#gid=' + sh.getSheetId() + '&range=A' + row;

  // Orders link (same logic)
  let ordersLink='', ordersLabel='';
  try {
    if (brand && so){
      const book = _getSOBookForBrand_(brand);
      ordersLabel = book.ok ? book.label : (brand==='VVS' ? '302_[VVS] Sales Order Report' : '301_[HPUSA] Sales Order Report');
      if (book.ok){
        const res = _findSORowByValue_(book.sh, so);
        if (res.ok) ordersLink = book.ss.getUrl() + '#gid=' + book.sheetId + '&range=A' + res.row;
      }
    }
  } catch(_){}

  // Prefill from Tracker→Log (targeted scan to pick a single row)
  let lastForm = null, lastRevNo = '';
  if (trackerId && brand && so){
    try {
      const ssT = SpreadsheetApp.openById(trackerId);
      const shT = ssT.getSheetByName('Log');
      if (shT && shT.getLastRow() >= 2){
        const tLastCol = Math.max(1, shT.getLastColumn());
        const tHeader  = shT.getRange(1,1,1,tLastCol).getValues()[0] || [];
        const pos = {}; tHeader.forEach((h,i)=>{ pos[String(h||'').trim()] = i+1; });

        const soCol = pos['SO#'] ? (pos['SO#']) : -1;
        const rvCol = pos['Revision #'] ? (pos['Revision #']) : -1;

        const nRows = shT.getLastRow() - 1; // excluding header
        let pickRow = -1;

        if (soCol > 0){
          // Read just the SO# column (2..last)
          const soVals = shT.getRange(2, soCol, nRows, 1).getValues();
          if (rvCol > 0){
            // Also read Revision # and choose highest revision for this SO
            const revVals = shT.getRange(2, rvCol, nRows, 1).getValues();
            let maxRev = -1;
            for (let i=0;i<nRows;i++){
              const mSO = String(soVals[i][0]||'').trim();
              if (mSO === so){
                const rv = Number(revVals[i][0]||0);
                if (rv > maxRev){ maxRev = rv; pickRow = 2 + i; }
              }
            }
            if (pickRow > 0) lastRevNo = String(shT.getRange(pickRow, rvCol).getValue() || '');
          } else {
            // No Revision # column → fallback: last by order (scan upward)
            for (let i=nRows-1;i>=0;i--){
              const mSO = String(soVals[i][0]||'').trim();
              if (mSO === so){ pickRow = 2 + i; break; }
            }
          }
        }

        if (pickRow > 0){
          // Fetch only the selected row (full width once; cheap)
          const rowT = shT.getRange(pickRow, 1, 1, tLastCol).getValues()[0] || [];
          lastForm = {
            AccentDiamondType:   _firstFromNames_(rowT, tHeader, ['Accent Type']),
            RingStyle:           _firstFromNames_(rowT, tHeader, ['Ring Style']),
            Metal:               _firstFromNames_(rowT, tHeader, ['Metal']),
            USSize:              _firstFromNames_(rowT, tHeader, ['US Size']),
            BandWidthMM:         _firstFromNames_(rowT, tHeader, ['Band Width (mm)']),
            CenterDiamondType:   _firstFromNames_(rowT, tHeader, ['Center Type']),
            Shape:               _firstFromNames_(rowT, tHeader, ['Shape']),
            DiamondDimension:    _firstFromNames_(rowT, tHeader, ['Diamond Dimension']),
            DesignNotes:         _firstFromNames_(rowT, tHeader, ['Design Notes'])
          };
        }
      }
    } catch(e){
      Logger.log('rev3d_init: tracker read failed: ' + e.message);
    }
  }

  return {
    ok: true,
    hasSO: !!so,
    hasTracker: !!trackerId,
    brand, so, odooUrl: odoo,
    customer: cust, email, phone,
    masterLink,
    threeDFolderLink, orderFolderLink,
    trackerUrl,
    lastRevNo: String(lastRevNo||''),
    lastForm: lastForm || null
  };
}

// ---------- Build "Copy into Odoo" text (unchanged) ----------
function previewRevOdooPaste(p){
  const lines = [];
  const add = s => lines.push(s);
  add('—— 3D REVISION REQUEST ——');
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
  if (notes){ notes.split(/\r?\n/).forEach(s => add('• ' + s.replace(/^\s*[-•]\s*/,'').trim())); }
  add('');
  add('CENTER STONE');
  add('• Type       : ' + (p.CenterDiamondType || ''));
  add('• Shape      : ' + (p.Shape || ''));
  add('• Dimension  : ' + (p.DiamondDimension || ''));
  add('');
  add('(Mode: 3D Revision Request)');
  return lines.join('\n');
}

// ---------- Submit (identical outputs; batched reads only) ----------
function submit3DRevision(payload){
  const form = payload && payload.form ? payload.form : {};
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet "00_Master Appointments".');

  const r = ss.getActiveRange();
  if (!r || r.getSheet().getName() !== sh.getName() || r.getRow() < 2){
    throw new Error('Select a data row on "00_Master Appointments" and try again.');
  }
  const row = r.getRow();

  // Header + single-row snapshot (batched)
  const lastCol = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1,1,1,lastCol).getValues()[0] || [];
  const H = {}; header.forEach((h,i)=>{ const k=String(h||'').trim(); if (k) H[k]=i+1; });

  const rowRange = sh.getRange(row,1,1,lastCol);
  const rowVals  = rowRange.getValues()[0] || [];
  const rowRich  = rowRange.getRichTextValues()[0] || [];

  const idx = name => H[name] ? (H[name]-1) : -1;
  const getPlain = name => { const i=idx(name); return i>=0 ? rowVals[i] : ''; };
  const getLink  = name => {
    const i=idx(name); if (i<0) return '';
    const rtv = rowRich[i];
    if (rtv && typeof rtv.getLinkUrl === 'function'){
      const u = rtv.getLinkUrl();
      if (u) return String(u);
    }
    const v = rowVals[i];
    return v != null ? String(v) : '';
  };

  const brand  = String(getPlain('Brand') || '').toUpperCase();
  const so     = String(getPlain('SO#') || '').replace(/^'/,'').trim();
  const odoo   = getLink('Odoo SO URL') || getLink('SO URL') || getLink('Odoo Link') || '';
  let trackerUrl = String(getPlain('3D Tracker') || '');
  let trackerId  = _idFromUrlLoose_(trackerUrl); // may be blank

  if (!brand || !so) throw new Error('Missing Brand / SO#. Please assign SO# first from “Assign SO”.');

  // If tracker missing and folder exists, attempt to create (best-effort)
  if (!trackerId){
    try {
      const folderLink = getPlain('05-3D Folder') || '';
      const folderId   = _idFromUrlLoose_(folderLink);
      if (folderId && typeof copy3DTrackerToSO_ === 'function'){
        const tr = copy3DTrackerToSO_(folderId, brand, so); // {id,url}
        if (tr && tr.id){ trackerId = tr.id; trackerUrl = tr.url || trackerUrl; }
      }
    } catch(e){ Logger.log('submit3DRevision: tracker create skipped: ' + e.message); }
  }
  if (!trackerId) throw new Error('Missing 3D Tracker for this SO. Run Start 3D / Assign SO first to scaffold the tracker.');

  // Short tag: derive from Shape + RingStyle (no master write; only log)
  const shortTag = _titleCase_( _truncate_([form.Shape, form.RingStyle].filter(Boolean).join(' '), 24) );

  // Append tracker log using your shared helper (unchanged)
  if (typeof append3DTrackerLog_ !== 'function'){
    throw new Error('Missing function append3DTrackerLog_. Please keep Start 3D server file enabled.');
  }

  const masterLink = ss.getUrl() + '#gid=' + sh.getSheetId() + '&range=A' + row;
  append3DTrackerLog_({
    trackerId: trackerId,
    action: '3D Revision Requested',
    form: Object.assign({}, form, { Mode: '3D Revision Request' }),
    brand, so, odooUrl: odoo, masterLink, shortTag
  });

  // Orders link for summary (unchanged logic)
  let ordersLink='', ordersLabel='';
  try{
    const book = _getSOBookForBrand_(brand);
    ordersLabel = book.ok ? book.label : (brand==='VVS' ? '302_[VVS] Sales Order Report' : '301_[HPUSA] Sales Order Report');
    if (book.ok){
      const res = _findSORowByValue_(book.sh, so);
      if (res.ok) ordersLink = book.ss.getUrl() + '#gid=' + book.sheetId + '&range=A' + res.row;
    }
  }catch(_){}

  SpreadsheetApp.flush(); // keep flush to preserve current timing expectations

  // R1 — New 3D Revision → restart the Custom Order Status reminder cycle
  try {
    var soNum = String((typeof so !== 'undefined' ? so :
                      (payload && (payload.so || payload.soNumber || payload.SO_Number)) ||
                      '')).trim();
    if (!soNum && typeof row !== 'undefined') {
      var soCol = header.indexOf('SO#') + 1;
      if (soCol > 0) soNum = String(sh.getRange(row, soCol).getDisplayValue()).trim();
    }

    // Also get customer name for fallback close when older COS was created pre‑SO
    var custName = '';
    var custCol = header.indexOf('Customer Name') + 1;
    if (custCol > 0) {
      custName = String(sh.getRange(row, custCol).getDisplayValue()).trim();
    }

    Remind.scheduleCOS(soNum, { customerName: custName }, true); // restart=true
  } catch (e) {
    console.warn('Remind.scheduleCOS (3D Revision) failed:', e && e.message ? e.message : e);
  }

  return {
    ok: true,
    summary: {
      brand, so,
      odooUrl: odoo,
      trackerUrl,
      masterLink,
      orderFolderLink: (H['Order Folder'] ? getPlain('Order Folder') : ''),
      threeDFolderLink: (H['05-3D Folder'] ? getPlain('05-3D Folder') : '')
    }
  };
}

// ---------- Small helpers (unchanged behaviors) ----------
function _idFromUrlLoose_(url){
  if (!url) return '';
  const m = String(url).match(/\/d\/([a-zA-Z0-9\-_]+)/) || String(url).match(/[-\/]folders\/([a-zA-Z0-9\-_]+)/);
  return m ? m[1] : '';
}
function _firstFromNames_(rowArr, headerArr, names){
  // Equivalent to prior _firstNonEmpty_ but works from [row values + header names]
  const pos = {}; headerArr.forEach((h,i)=>{ pos[String(h||'').trim()] = i; });
  for (let i=0;i<names.length;i++){
    const idx = pos[names[i]];
    if (idx != null && idx >= 0){
      const v = rowArr[idx];
      if (v !== '' && v != null) return v;
    }
  }
  return '';
}
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



