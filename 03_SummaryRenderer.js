/** Create/replace the Client Summary tab (clean layout).
 * - Section 1 (a–f) = Organized Client Information (from Scribe JSON)
 * - Section 2 (a–d) = In-Depth Analysis (Strategist)
 * - Lead Scores, Followups, Low-Confidence, Provenance
 * Backwards compatible:
 *   upsertClientSummaryTab_(rootApptId, scribeObj, apptIsoOpt, transcriptUrlOpt [, strategistObjOpt])
 */
function upsertClientSummaryTab_(rootApptId, scribeObj, apptIsoOpt, transcriptUrlOpt, strategistObjOpt /*, optsOpt */) {
  
  // ---------- small utils ----------
  const tz = PropertiesService.getScriptProperties().getProperty('DEFAULT_TZ') || 'America/Los_Angeles';

  const fmtMoneyRange = (low, high) => {
    const n = (x) => (x == null || x === '') ? '' : '$' + Number(x).toLocaleString();
    if (low == null && high == null) return '';
    return (low != null && high != null) ? (n(low) + ' – ' + n(high)) : (n(low != null ? low : high));
  };

  const linkOrBlank = (u, label) =>
    (u && String(u).trim())
      ? '=HYPERLINK("' + String(u).trim() + '","' + String((label || u)).trim() + '")'
      : '';
  const newestByRegex_ = (folder, re) => {
    let newest = null, ts = 0, it = folder.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (!re.test(f.getName())) continue;
      const t = f.getDateCreated().getTime();
      if (t > ts) {
        ts = t;
        newest = f;
      }
    }
    return newest;
  };

  // Treat empty string as a valid, intentional value; only fall back on null/undefined
  function firstDefined_() {
    for (var i = 0; i < arguments.length; i++) {
      var v = arguments[i];
      if (v !== undefined && v !== null) return v;
    }
    return '';
  }

  // --------- AUTO-LOAD Scribe if caller passed blank/empty ----------
  try {
    const needsScribe =
      !scribeObj ||
      (typeof scribeObj === 'object' && Object.keys(scribeObj).length === 0);

    if (needsScribe) {
      const ms = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
      const apId = getApFolderIdForRoot_(ms, rootApptId);
      if (apId) {
        const ap = DriveApp.getFolderById(apId);
        const sIt = ap.getFoldersByName('04_Summaries');
        if (sIt.hasNext()) {
          const sFolder = sIt.next();
          // Prefer corrected, else base
          const pickNewest = (re) => {
            let newest = null, t = 0, it = sFolder.getFiles();
            while (it.hasNext()){
              const f = it.next();
              if (!re.test(f.getName())) continue;
              const ts = (f.getLastUpdated ? f.getLastUpdated().getTime()
                                           : f.getDateCreated().getTime());
              if (ts > t){ t = ts; newest = f; }
            }
            return newest;
          };
          const corrected = pickNewest(/__summary_corrected_.*\.json$/i);
          const base      = pickNewest(/__summary_.*\.json$/i);
          const f = corrected || base;
          if (f) {
            scribeObj = JSON.parse(f.getBlob().getDataAsString('UTF-8'));
          }
        }
      }
    }
  } catch(_) { /* best-effort auto-load; keep going */ }

  // Ensure Scribe is normalized and durability concerns are removed
  try { scribeObj = normalizeScribe_(scribeObj); } catch(_) {}

  // ---------- Transcript URL (only what was passed in) ----------
  let transcriptUrl = String(transcriptUrlOpt || '');

  // ---------- Report + brand ----------
  const opts = (arguments.length >= 6 && typeof arguments[5] === 'object') ? (arguments[5] || {}) : {};
  const reportId = String(opts.reportId || '').trim() || getReportIdForRoot_(rootApptId);
  if (!reportId) throw new Error('No Client Status Report ID found for ' + rootApptId);
  const rpt = SpreadsheetApp.openById(reportId);
  try { ensureReportConfig_(rpt, { rootApptId, reportId }); } catch(_) {}
  const brand = getBrandForRoot_(rootApptId);
  const accent = brandAccentHex_(brand);
  const when = apptIsoOpt ? Utilities.formatDate(new Date(apptIsoOpt), tz, 'yyyy-MM-dd') : Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const tabName = 'Consult — ' + when;
  const sh = rpt.getSheetByName(tabName) || rpt.insertSheet(tabName);
  sh.clear();
  sh.setHiddenGridlines(true);
  try { sh.setTabColor(accent); } catch (_) { }

  // ---------- Read Master quick fields ----------
  const master = (function fetchMasterForRoot_() {
    const id = PROP_('SPREADSHEET_ID');
    if (!id) return {};
    const ss = SpreadsheetApp.openById(id);
    const s = ss.getSheetByName('00_Master Appointments');
    if (!s) return {};
    const hdr = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
    const iRoot = hdr.indexOf('RootApptID');
    if (iRoot < 0) return {};
    const last = s.getLastRow();
    if (last < 2) return {};
    const vals = s.getRange(2, 1, last - 1, s.getLastColumn()).getValues();
    const rIdx = vals.findIndex(r => String(r[iRoot] || '').trim() === String(rootApptId).trim());
    if (rIdx < 0) return {};
    const row = vals[rIdx];
    const V = (name, alts) => {
      let i = hdr.indexOf(name);
      if (i < 0 && alts) for (const a of alts) {
        const j = hdr.indexOf(a);
        if (j >= 0) { i = j; break; }
      }
      return i >= 0 ? String(row[i] || '').trim() : '';
    };
    return {
      folderUrl: V('Folder URL', ['Client Folder URL']),
      prospectUrl: V('Prospect Folder URL'),
      intakeUrl: V('IntakeDocURL', ['Intake URL']),
      checklistUrl: V('Checklist URL'),
      quotationUrl: V('Quotation URL'),
      summaryJsonUrl: V('Summary JSON URL'),
      strategistJsonUrl: V('Strategist JSON URL'),
      apptId: V('APPT_ID'),
      visitType: V('Visit Type'),
      location: V('Location'),
      assignedRep: V('Assigned Rep'),
      assistedRep: V('Assisted Rep'),
      customerPhone: V('Phone', ['PhoneNorm']),
      customerEmail: V('Email', ['EmailLower'])
    };
  })();

  // ---------- Load Strategist JSON if not provided ----------
  let strategist = strategistObjOpt || null;
  let strategistJsonDriveUrl = master.strategistJsonUrl || '';
  if (!strategist) {
    try {
      const ss = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
      const apId = getApFolderIdForRoot_(ss, rootApptId);
      if (apId) {
        const ap = DriveApp.getFolderById(apId);
        const sIt = ap.getFoldersByName('04_Summaries');
        if (sIt.hasNext()) {
          const sf = sIt.next();
          const f = newestByRegex_(sf, /__analysis_.*\.json$/i);
          if (f) {
            strategist = JSON.parse(f.getBlob().getDataAsString('UTF-8'));
            strategistJsonDriveUrl = 'https://drive.google.com/file/d/' + f.getId() + '/view';
          }
        }
      }
    } catch (_) { }
  }

  // ---------- Sequential layout using one row ----------
  let row = 1;

  // Section 1
  row = renderSection1_({
    sh, row, scribeObj, master, brand, accent, when, transcriptUrl, fmtMoneyRange,
    strategistOpt: strategist // <— add this
  });

  // Section 2
  row = renderSection2_({ sh, startRow: row, strategist, accent, rootApptId });

  // Followups (split layout: A:B and E:G)
  sh.getRange(row, 1, 1, 2).merge().setValue('Followups').setFontWeight('bold');
  row++;

  const repName = getAssignedRepForRoot_(rootApptId);
  const assistedRep = (typeof getAssistedRepForRoot_ === 'function') ? getAssistedRepForRoot_(rootApptId) : (master.assistedRep || '');

  let fu = (scribeObj.followups || []).map(f => {
    const owner = f.owner || '';
    const due = f.due_iso || '';
    const task = f.task || '';
    const notes = f.notes || '';
    let assignedTo = '';
    if (/^customer$/i.test(owner)) assignedTo = [repName, assistedRep].filter(Boolean).join(' & ');
    return [owner, due, task, notes, assignedTo];
  });
  if (!fu.length) fu = [['', '', '(none)', '', '']];

  // A:B — Owner, Due
  const leftRange  = sh.getRange(row, 1, fu.length, 2).setValues(fu.map(r => [r[0], r[1]]));
  styleTable_(leftRange, ['Owner', 'Due (ISO)'], { wrapCols: [], accentHex: accent });

  // E:G — Task, Notes, Assigned To (start at column E)
  const rightRange = sh.getRange(row, 5, fu.length, 3).setValues(fu.map(r => [r[2], r[3], r[4]]));
  styleTable_(rightRange, ['Task', 'Notes', 'Assigned To'], { wrapCols: [1, 2, 3], accentHex: accent });

  row += fu.length + 1;

  // Provenance footer
  const RENDERER_VERSION = '1.6.1';
  const prov = [
    ['Model Version (Scribe)', 'gpt-5-mini'],
    ['Model Version (Strategist)', 'gpt-5'],
    ['Rendered At', Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss z")],
    ['Renderer', RENDERER_VERSION]
  ];
  const pRange = sh.getRange(row, 1, prov.length, 2).setValues(prov);
  styleLabelValueBlock_(pRange, accent);
  row += prov.length + 1;

  // Sheet cosmetics (left block A–C; spacer D; right block E..)
  sh.setColumnWidths(1, 1, 220); // A: left labels
  sh.setColumnWidths(2, 1, 400); // B: left values
  sh.setColumnWidths(3, 1, 120); // C: hidden JSON paths (smaller font set in renderSection1_)
  sh.setColumnWidths(4, 1, 24);  // D: gutter / spacer
  sh.setColumnWidths(5, 1, 110); // E: right labels — half width (Req #4)
  sh.setColumnWidths(6, 1, 240); // F: right values
  // (G–I used by narrative merge; default widths are fine)
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setVerticalAlignment('top').setWrap(true);
  try { sh.autoResizeColumns(2, 1); } catch (_) { }

  // === DROP-IN (inside upsertClientSummaryTab_) ===
  function renderSection1_({ sh, row, scribeObj, master, brand, accent, when, transcriptUrl, fmtMoneyRange, strategistOpt }) {
    // Build low-confidence path set once (existing helper)
    const lowSet = listLowConfidencePathsFromConf_(scribeObj, 0.69);

    // Helper: write Label(A)/Value(B)/Path(C) and queue highlight if low-confidence
    const pending = []; // rows in this block that need highlight on B
    function writeLVWithPath_(sheet, r, label, value, jsonPath) {
      sheet.getRange(r, 1).setValue(label);
      sheet.getRange(r, 2).setValue(value);
      sheet.getRange(r, 3).setValue(jsonPath || '');
      if (jsonPath && lowSet.has(jsonPath)) pending.push(r);
    }
    function applyBlockStyleAndHighlights_(startRow) {
      const len = row - startRow;
      if (len <= 0) return;
      const rng = sh.getRange(startRow, 1, len, 2);
      styleLabelValueBlock_(rng, accent);
      while (pending.length) {
        const r = pending.pop();
        sh.getRange(r, 2).setBackground('#fff9c4'); // pale yellow
      }
    }

    // ---- 1. Organized Client Information (merge A:B) ----
    const clientName =
      (scribeObj?.customer_profile?.customer_name) ||
      (scribeObj?.customer_name) ||
      (scribeObj?.customer?.name) ||
      (master?.customerName) ||
      '';

    sh.getRange(row, 1, 1, 2).merge()
      .setValue('1. Organized Client Information' + (clientName ? (' — ' + clientName) : ''))
      .setFontWeight('bold').setFontSize(12);
    row += 2;

    // a. Customer Profile (merge A:B)
    sh.getRange(row, 1, 1, 2).merge().setValue('a. Customer Profile').setFontWeight('bold'); row++;
    const profileStart = row;

    // Design Notes for context
    const ds = scribeObj?.design_specs || {};
    const notes = ds.design_notes || [ds.basket_design, ds.prong_type, ds.cut_polish_sym].filter(Boolean).join(' • ') || '';

    writeLVWithPath_(sh, row++, 'Client Name', clientName, 'customer_profile.customer_name');
    writeLVWithPath_(sh, row++, 'Design Notes', notes, 'design_specs.design_notes');

    writeLVWithPath_(sh, row++, 'Consult Date', when, '');
    writeLVWithPath_(sh, row++, 'Brand', brand, '');
    writeLVWithPath_(sh, row++, 'Visit Type', master.visitType || '', '');
    writeLVWithPath_(sh, row++, 'Assigned Rep / Assisted', [master.assignedRep, master.assistedRep].filter(Boolean).join(' / '), '');
    writeLVWithPath_(sh, row++, 'Contact Info', [master.customerPhone, master.customerEmail].filter(Boolean).join(' · '), '');

    const cp = (scribeObj && scribeObj.customer_profile) || {};
    writeLVWithPath_(sh, row++, 'Occupation', (cp.occupation || ''), 'customer_profile.occupation');
    writeLVWithPath_(sh, row++, 'Partner', (cp.partner_name || ''), 'customer_profile.partner_name');
    writeLVWithPath_(sh, row++, 'Emotional Motivation', (cp.emotional_motivation || ''), 'customer_profile.emotional_motivation');

    // Priorities moved INTO Profile
    const cp2 = scribeObj.client_priorities || {};

    // (Comm prefs & Decision makers were already moved in earlier patch)
    writeLVWithPath_(sh, row++, 'Communication Preferences', (cp.comm_prefs || []).join(' • '), 'customer_profile.comm_prefs');
    writeLVWithPath_(sh, row++, 'Decision Maker(s)', (cp.decision_makers || []).join(' • '), 'customer_profile.decision_makers');

    applyBlockStyleAndHighlights_(profileStart);
    row += 1;

    // Key Links (merge A:B)
    let apRootUrl = '';
    try {
      const ms0 = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
      const apId0 = (typeof getApFolderIdForRoot_ === 'function') ? getApFolderIdForRoot_(ms0, rootApptId) : '';
      if (apId0) apRootUrl = 'https://drive.google.com/drive/folders/' + apId0;
    } catch (_) { }

    // Gather newest consult/debrief + saved JSONs
    let consultLink = '', debriefLink = '', scribeLink = '', strategistLink = '';
    try {
      const ms0 = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
      const apId0 = getApFolderIdForRoot_(ms0, rootApptId);
      if (apId0){
        const ap0 = DriveApp.getFolderById(apId0);
        const pair0 = __findTranscriptAndDebriefTexts__(ap0);
        consultLink = pair0.consultUrl || '';
        debriefLink = pair0.debriefUrl || '';

        const summariesIt = ap0.getFoldersByName('04_Summaries');
        if (summariesIt.hasNext()){
          const sFolder = summariesIt.next();
          const newest = (re) => {
            let f=null, ts=0, it=sFolder.getFiles();
            while (it.hasNext()){
              const x = it.next();
              if (!re.test(x.getName())) continue;
              const t = (x.getLastUpdated ? x.getLastUpdated().getTime()
                                          : x.getDateCreated().getTime());
              if (t > ts){ ts=t; f=x; }
            }
            return f ? 'https://drive.google.com/file/d/' + f.getId() + '/view' : '';
          };
          scribeLink     = newest(/__summary_.*\.json$/i);
          strategistLink = newest(/__analysis_.*\.json$/i);
        }
      }
    } catch(_) {}

    const driveLinks = [];
    if (apRootUrl) driveLinks.push(['RootAppt Folder', apRootUrl]);
    if (master.folderUrl)   driveLinks.push(['Client Folder', master.folderUrl]);
    if (master.prospectUrl) driveLinks.push(['Prospect Folder', master.prospectUrl]);
    if (master.intakeUrl)   driveLinks.push(['Intake Doc', master.intakeUrl]);
    if (master.checklistUrl)driveLinks.push(['Checklist', master.checklistUrl]);
    if (master.quotationUrl)driveLinks.push(['Quotation', master.quotationUrl]);
    if (consultLink) driveLinks.push(['Consult Transcript (.txt)', consultLink]);
    if (debriefLink) driveLinks.push(['Rep Debrief (.txt)', debriefLink]);
    if (scribeLink)  driveLinks.push(['Scribe JSON', scribeLink]);
    if (strategistLink) driveLinks.push(['Strategist JSON', strategistLink]);

    sh.getRange(row, 1, 1, 2).merge().setValue('Key Links').setFontWeight('bold'); row++;
    if (driveLinks.length) {
      const vals = driveLinks.map(([label, url]) => [label, '=HYPERLINK("' + url + '","' + label + '")']);
      const linkRange = sh.getRange(row, 1, vals.length, 2).setValues(vals);
      styleLabelValueBlock_(linkRange, accent);
      row += vals.length + 1;
    } else {
      const rng2 = sh.getRange(row, 1, 1, 2).setValues([['Key Links', '']]);
      styleLabelValueBlock_(rng2, accent);
      row += 2;
    }

    // b. Detailed Style Preferences (merge A:B) — *** Durability Concerns removed ***
    sh.getRange(row, 1, 1, 2).merge().setValue('b. Detailed Style Preferences').setFontWeight('bold'); row++;
    const styleStart = row;
    const d = scribeObj.diamond_specs || {};
    const dst = (strategistOpt && strategistOpt.detailed_style) || {};

    writeLVWithPath_(sh, row++, 'Shape',       (dst.shape || d.shape || ''),                  'diamond_specs.shape');
    writeLVWithPath_(sh, row++, 'Ratio',       (d.ratio ?? ds.target_ratio ?? dst.ratio ?? ''), 'diamond_specs.ratio');
    writeLVWithPath_(sh, row++, 'Band Width (mm)', (ds.band_width_mm ?? dst.band_width_mm ?? ''), 'design_specs.band_width_mm');
    writeLVWithPath_(sh, row++, 'Metal',       (ds.metal || ''),                               'design_specs.metal');
    writeLVWithPath_(sh, row++, 'Engraving',   (ds.engraving ?? dst.engraving ?? ''),          'design_specs.engraving');

    writeLVWithPath_(sh, row++, 'Wedding Band Fit', (ds.wedding_band_fit || ''), 'design_specs.wedding_band_fit');

    applyBlockStyleAndHighlights_(styleStart);
    row += 1;

    // c. Specs Requested (merge A:B)
    sh.getRange(row, 1, 1, 2).merge().setValue('c. Specs Requested').setFontWeight('bold'); row++;
    const specStart = row;
    writeLVWithPath_(sh, row++, 'Diamond Type', (d.lab_or_natural || ''), 'diamond_specs.lab_or_natural');
    writeLVWithPath_(sh, row++, 'Target Carat(s)', (d.carat != null ? Number(d.carat) : ''), 'diamond_specs.carat');
    writeLVWithPath_(sh, row++, 'Color Range', (d.color || ''), 'diamond_specs.color');
    writeLVWithPath_(sh, row++, 'Clarity Range', (d.clarity || ''), 'diamond_specs.clarity');
    writeLVWithPath_(sh, row++, 'Cut / Polish / Symmetry', (d.cut_polish_sym || ds.cut_polish_sym || ''), 'diamond_specs.cut_polish_sym');
    applyBlockStyleAndHighlights_(specStart);
    row += 2;

    // d. Budget (merge A:B)
    sh.getRange(row, 1, 1, 2).merge().setValue('d. Budget').setFontWeight('bold'); row++;
    const budgetStart = row;
    writeLVWithPath_(sh, row++, 'Budget Window', fmtMoneyRange(scribeObj.budget_low, scribeObj.budget_high), '');
    writeLVWithPath_(sh, row++, 'Budget Low (numeric)', (scribeObj.budget_low != null ? Number(scribeObj.budget_low) : ''), 'budget_low');
    writeLVWithPath_(sh, row++, 'Budget High (numeric)', (scribeObj.budget_high != null ? Number(scribeObj.budget_high) : ''), 'budget_high');
    applyBlockStyleAndHighlights_(budgetStart);
    row += 1;

    // e. Timeline (merge A:B) — with Occasion / Intent moved here
    sh.getRange(row, 1, 1, 2).merge().setValue('e. Timeline').setFontWeight('bold'); row++;
    const timelineStart = row;
    writeLVWithPath_(sh, row++, 'Needed By', (scribeObj.timeline || ''), 'timeline');
    writeLVWithPath_(sh, row++, 'Occasion / Intent', (cp.occasion_intent || ''), 'customer_profile.occasion_intent');
    applyBlockStyleAndHighlights_(timelineStart);
    row += 1;

    // Smaller font for hidden Column C (JSON paths)
    try {
      sh.getRange(1, 3, sh.getMaxRows(), 1).setFontSize(8);
      sh.hideColumns(3);
    } catch (_) {}

    return row + 1;
  }


}

function setByPath_(obj, path, value){
  const parts = String(path||'').split('.').filter(Boolean);
  if (!parts.length) return;
  let cur = obj;
  for (let i=0; i<parts.length-1; i++){
    const k = parts[i];
    if (!cur[k] || typeof cur[k] !== 'object' || Array.isArray(cur[k])) cur[k] = {};
    cur = cur[k];
  }
  cur[parts[parts.length-1]] = value;
}


/* ===== Formatting helpers ===== */

/** Bold left labels, light gray background on label col, thin borders. */
function styleLabelValueBlock_(rng, accentHex){
  const sh = rng.getSheet();
  const rows = rng.getNumRows();
  const labelCol = rng.offset(0,0,rows,1);
  const valueCol = rng.offset(0,1,rows,1);
  const bg = (typeof lighten_ === 'function')
    ? lighten_(accentHex || '#f0f3f7', 0.80)
    : '#e9edf3';
  labelCol.setFontWeight('bold').setBackground(bg).setFontColor('#000').setHorizontalAlignment('left');
  valueCol.setBackground('#ffffff');
  outline_(sh, rng);
}

// --- NEW: rectangle overlap helper
function _rangesOverlap_(a, b) {
  const ar = a.getRow(), ac = a.getColumn(), anr = a.getNumRows(), anc = a.getNumColumns();
  const br = b.getRow(), bc = b.getColumn(), bnr = b.getNumRows(), bnc = b.getNumColumns();
  const aMaxR = ar + anr - 1, aMaxC = ac + anc - 1;
  const bMaxR = br + bnr - 1, bMaxC = bc + bnc - 1;
  return !(aMaxR < br || bMaxR < ar || aMaxC < bc || bMaxC < ac);
}

// --- NEW: clear any banding that touches rng
function _clearBandingOnRange_(rng) {
  const sh = rng.getSheet();
  const bands = sh.getBandings(); // all bandings on the sheet
  if (!bands || !bands.length) return;
  const toRemove = [];
  bands.forEach(b => {
    try {
      const br = b.getRange();
      if (_rangesOverlap_(rng, br)) toRemove.push(b);
    } catch (_) {}
  });
  toRemove.forEach(b => { try { b.remove(); } catch(_) {} });
}

/** Zebra stripes using Row Banding (idempotent). */
function zebra_(rng, accentHex){
  if (!rng) return;                           // 🔒 guard
  if (rng.getNumRows() <= 0 || rng.getNumColumns() <= 0) return;  // 🔒 guard

  _clearBandingOnRange_(rng);
  const band = rng.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  try {
    const base = (accentHex || '#f0f3f7');
    const light1 = (typeof lighten_ === 'function') ? lighten_(base, 0.96) : '#f7f9fb';
    const light2 = (typeof lighten_ === 'function') ? lighten_(base, 0.90) : '#eef3f7';
    band.setHeaderRowColor(null).setFirstRowColor(light1).setSecondRowColor(light2);
  } catch(_){}
}

/** Draw light borders around a range. */
function outline_(sh, rng){
  if (!rng) return;                           // 🔒 guard
  if (rng.getNumRows() <= 0 || rng.getNumColumns() <= 0) return;  // 🔒 guard
  sh.getRange(rng.getRow(), rng.getColumn(), rng.getNumRows(), rng.getNumColumns())
    .setBorder(true,true,true,true,true,true,'#e0e0e0',SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Style a tabular block:
 * - adds a header row just above the data block (if not present)
 * - sets borders, zebra, numeric formatting for money/percent
 * - wraps selected cols
 */
function styleTable_(dataRange, headerLabels, opts){
  // 🔒 guard invalid range
  if (!dataRange) return;

  const sh = dataRange.getSheet();
  const nR = dataRange.getNumRows(), nC = dataRange.getNumColumns();
  // 🔒 nothing to style if zero-sized
  if (nR <= 0 || nC <= 0) return;

  // 🔧 clamp header row to 1 (row 0 is invalid)
  const headerRowIdx = Math.max(1, dataRange.getRow() - 1);
  const headerRow = sh.getRange(headerRowIdx, dataRange.getColumn(), 1, nC);

  // clear banding before re-applying
  _clearBandingOnRange_(headerRow);
  _clearBandingOnRange_(dataRange);

  headerRow.setValues([ headerLabels.slice(0,nC).concat(new Array(Math.max(0,nC-headerLabels.length)).fill('')) ]);
  const headerBg = lighten_ ? lighten_((opts && opts.accentHex) || '#f0f3f7', 0.80) : '#e9edf3';
  headerRow.setFontWeight('bold').setBackground(headerBg);
  outline_(sh, headerRow);

  outline_(sh, dataRange);
  zebra_(dataRange, (opts && opts.accentHex) || '#f0f3f7');

  if (opts && opts.money && opts.numCols){
    opts.numCols.forEach(c=>{
      if (c>=1 && c<=nC) dataRange.offset(0, c-1, nR, 1).setNumberFormat('$#,##0');
    });
  }
  if (opts && opts.wrapCols){
    opts.wrapCols.forEach(c=>{
      if (c>=1 && c<=nC) dataRange.offset(0, c-1, nR, 1).setWrap(true);
    });
  }
}



/** Apply a simple 0→100 green heat bar on a (single-column) numeric range. */
function addHeatRule0to100_(sh, rng){
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint('#1b5e20')
    .setGradientMidpointWithValue('#a5d6a7', SpreadsheetApp.InterpolationType.NUMBER, '50')
    .setGradientMinpoint('#f1f8e9')
    .setRanges([rng])
    .build();
  const rules = sh.getConditionalFormatRules().concat(rule);
  sh.setConditionalFormatRules(rules);
}

/** Simple titled list section: title (bold) + items column with zebra. Returns next free row. */
function titledList_(sh, startRow, title, items, col=4){
  sh.getRange(startRow,col).setValue(title).setFontWeight('bold');
  startRow++;
  const arr = (items && items.length) ? items.map(x=>[x]) : [['(none)']];
  const rng = sh.getRange(startRow+1,col,arr.length,1).setValues(arr);
  zebra_(rng);
  outline_(sh, rng);
  return startRow + arr.length + 1;
}

function renderExecutiveSummary_({ sh, row, strategist, accent }) {
  if (!strategist || !Array.isArray(strategist.executive_summary) || !strategist.executive_summary.length) return row;
  sh.getRange(row,1).setValue('Executive Summary').setFontWeight('bold'); row++;
  const items = strategist.executive_summary.map(b => [b]);
  const r = sh.getRange(row,1,items.length,1).setValues(items);
  zebra_(r, accent); outline_(sh, r);
  return row + items.length + 1;
}

/** Re-render the Client Summary tab from the latest saved JSON (no OpenAI call). */
function rerenderClientSummaryTabForRoot_(rootApptId){
  const ss = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));

  // Resolve appointment folder
  const apId = getApFolderIdForRoot_(ss, rootApptId);
  if (!apId) throw new Error('No RootAppt Folder ID for '+rootApptId);
  const ap = DriveApp.getFolderById(apId);

  // --- Load newest Scribe JSON (required for Section 1) ---
  const sfIt = ap.getFoldersByName('04_Summaries');
  if (!sfIt.hasNext()) throw new Error('No 04_Summaries for '+rootApptId);
  const sf = sfIt.next();

  function newestByRegex_(folder, re){
    let newest=null, ts=0, it=folder.getFiles();
    while (it.hasNext()){
      const f = it.next();
      if (!re.test(f.getName())) continue;
      const t=f.getDateCreated().getTime();
      if (t>ts){ ts=t; newest=f; }
    }
    return newest;
  }

  const scribeFile = newestByRegex_(sf, /__summary_.*\.json$/i);
  if (!scribeFile) throw new Error('No Scribe summary JSON (__summary_*.json) for '+rootApptId);
  const scribeObj = JSON.parse(scribeFile.getBlob().getDataAsString('UTF-8'));

  // --- Load newest Strategist JSON (optional Section 2) ---
  const strategistFile = newestByRegex_(sf, /__analysis_.*\.json$/i);
  const strategistObj = strategistFile
    ? JSON.parse(strategistFile.getBlob().getDataAsString('UTF-8'))
    : null;

  // Get newest transcript URL (for link)
  let transcriptUrl = '';
  const tfIt = ap.getFoldersByName('03_Transcripts');
  if (tfIt.hasNext()){
    const tf = tfIt.next();
    const txt = newestByRegex_(tf, /\.txt$/i);
    if (txt) transcriptUrl = 'https://drive.google.com/file/d/'+txt.getId()+'/view';
  }

  // Date for tab name: prefer Master’s Appt ISO; else Scribe file ts
  const apISO = getApptIsoForRoot_(ss, rootApptId) || new Date().toISOString();

  // Always resolve the Client Status Report ID from Master
  const reportId = getReportIdForRoot_(rootApptId);

  upsertClientSummaryTab_(rootApptId, scribeObj, apISO, transcriptUrl, strategistObj, { reportId });
  Logger.log('Re-rendered summary tab for %s from %s (+ strategist: %s)',
             rootApptId, scribeFile.getName(), strategistFile ? strategistFile.getName() : 'none');
}

function renderOrganizedClientInfoSection_(sh, startRow, scribe, master, accentHex){
  const A = (label, value)=> [[label, value==null?'':String(value)]];

  // ===== 1.a Customer Profile =====
  sh.getRange(startRow,1).setValue('1. Organized Client Information').setFontWeight('bold').setFontSize(12);
  startRow += 2;
  sh.getRange(startRow,1).setValue('a. Customer Profile').setFontWeight('bold'); startRow++;

  const prof = []
    .concat(A('Client Name', scribe.customer_name||''))
    .concat(A('Consult Date', Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'America/Los_Angeles','yyyy-MM-dd')))
    .concat(A('Brand', getBrandForRoot_(String(sh.getParent().getName()).split('—')[0] || '') || ''))
    .concat(A('Visit Type', master.visitType||''))
    .concat(A('Assigned Rep / Assisted Rep', [master.assignedRep, master.assistedRep].filter(Boolean).join(' / ')))
    .concat(A('Contact Info', [master.customerPhone, master.customerEmail].filter(Boolean).join(' · ')));

  let rng = sh.getRange(startRow,1,prof.length,2).setValues(prof);
  styleLabelValueBlock_(rng, accentHex);
  startRow += prof.length + 1;

  // Resolve RootAppt Folder URL for Key Links (inline block)
  let apRootUrl = '';
  try {
    const ss0 = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
    const apId = (typeof getApFolderIdForRoot_ === 'function') ? getApFolderIdForRoot_(ss0, String(sh.getParent().getName()).split('—')[0] ? rootApptId : rootApptId) : '';
    if (apId) apRootUrl = 'https://drive.google.com/drive/folders/' + apId;
  } catch(_) {}

  // Key Links row (inline list)
  const links = [
    apRootUrl && `=HYPERLINK("${apRootUrl}","RootAppt Folder")`,
    master.folderUrl && `=HYPERLINK("${master.folderUrl}","Client Folder")`,
    master.prospectUrl && `=HYPERLINK("${master.prospectUrl}","Prospect Folder")`,
    master.intakeUrl && `=HYPERLINK("${master.intakeUrl}","Intake Doc")`,
    master.checklistUrl && `=HYPERLINK("${master.checklistUrl}","Checklist")`,
    master.quotationUrl && `=HYPERLINK("${master.quotationUrl}","Quotation")`,
    `=HYPERLINK("${(function(u){return u||''})(master.summaryJsonUrl)}","Summary JSON")`
  ].filter(Boolean).join('  •  ');
  rng = sh.getRange(startRow,1,1,2).setValues([['Key Links', links]]); 
  styleLabelValueBlock_(rng, accentHex);
  startRow += 2;

  // ===== 1.b Style Preferences =====
  sh.getRange(startRow,1).setValue('b. Style Preferences').setFontWeight('bold'); startRow++;
  const styleRows = [
    ['Ring Style', scribe.ring_style||''],
    ['Inspiration', (scribe.design_refs||[]).map(r=>r.file?`=HYPERLINK("${r.file}","${r.name||'ref'}")`:(r.name||'')).join(' • ')]
  ];
  rng = sh.getRange(startRow,1,styleRows.length,2).setValues(styleRows);
  styleLabelValueBlock_(rng, accentHex);
  startRow += styleRows.length + 1;

  // ===== 1.c Specs Requested =====
  sh.getRange(startRow,1).setValue('c. Specs Requested').setFontWeight('bold'); startRow++;

  // Design Specs table
  const ds = scribe.design_specs || {};
  const dsRows = [
    ['Ring Design', [ds.basket_design, ds.prong_type, ds.cut_polish_sym].filter(Boolean).join(' • ')],
    ['Wedding band fit', ds.wedding_band_fit || ''],
    ['Engraving', ds.engraving || ''],
    ['Ring Size', ds.ring_size || ''],
    ['Band Width (mm)', ds.band_width_mm!=null ? Number(ds.band_width_mm) : ''],
    ['Target Ratio', ds.target_ratio || '']
  ];
  rng = sh.getRange(startRow,1,dsRows.length,2).setValues(dsRows);
  styleLabelValueBlock_(rng, accentHex);
  startRow += dsRows.length + 1;

  // Diamond Specs table
  const d = scribe.diamond_specs || {};
  const dRows = [
    ['Diamond Type', d.lab_or_natural || ''],
    ['Shape(s)', d.cut || ''],
    ['Target Carat(s)', d.carat!=null ? Number(d.carat) : ''],
    ['Color Range', d.color || ''],
    ['Clarity Range', d.clarity || ''],
    ['Cut / Polish / Symmetry', d.cut_polish_sym || (scribe.design_specs && scribe.design_specs.cut_polish_sym) || ''],
    ['Ratio', d.ratio || '']
  ];
  rng = sh.getRange(startRow,1,dRows.length,2).setValues(dRows);
  styleLabelValueBlock_(rng, accentHex);
  startRow += dRows.length + 2;

  // ===== 1.d Budget =====
  sh.getRange(startRow,1).setValue('d. Budget').setFontWeight('bold'); startRow++;
  const fmtMoney=(v)=> (v==null||v==='')?'':('$'+Number(v).toLocaleString());
  const budgetRows = [
    ['Budget Window', (function(l,h){ 
      if (l==null&&h==null) return '';
      if (l!=null&&h!=null) return fmtMoney(l)+' – '+fmtMoney(h);
      return fmtMoney(l!=null?l:h);
    })(scribe.budget_low, scribe.budget_high)],
    ['Budget Notes', (scribe.summary_md||'').match(/budget[^.\n]*[.\n]/i)?.[0] || '']
  ];
  rng = sh.getRange(startRow,1,budgetRows.length,2).setValues(budgetRows);
  styleLabelValueBlock_(rng, accentHex);
  startRow += budgetRows.length + 1;

  // ===== 1.e Timeline =====
  sh.getRange(startRow,1).setValue('e. Timeline').setFontWeight('bold'); startRow++;
  const tlRows = [
    ['Deadline (ISO)', scribe.timeline || ''],
    ['Occasion / Intent', (scribe.client_priorities && scribe.client_priorities.occasion_intent) || '']
  ];
  rng = sh.getRange(startRow,1,tlRows.length,2).setValues(tlRows);
  styleLabelValueBlock_(rng, accentHex);
  startRow += tlRows.length + 1;

  // ===== 1.f Client Priorities =====
  sh.getRange(startRow,1).setValue('f. Client Priorities').setFontWeight('bold'); startRow++;
  const cp = scribe.client_priorities || {};
  const listBlock = (title, arr)=>{
    sh.getRange(startRow,1).setValue(title).setFontStyle('italic'); startRow++;
    const items = (arr&&arr.length?arr:['(none)']).map(x=>[x]);
    const lr = sh.getRange(startRow,1,items.length,1).setValues(items);
    zebra_(lr, accentHex); outline_(sh, lr);
    startRow += items.length + 1;
  };
  listBlock('Top priorities', cp.top_priorities);
  listBlock('Non-negotiables', cp.non_negotiables);
  listBlock('Nice-to-haves', cp.nice_to_haves);
  listBlock('Durability concerns', cp.durability_concerns);
  listBlock('Communication preferences', cp.comm_prefs);
  listBlock('Decision makers', cp.decision_makers);

  return startRow + 1;
}

function renderBriefBanner_({ sh, row, strategist, accent }) {
  if (!strategist) return row;
  const banner = [
    ['Recommended Play', strategist.recommended_play || '(none)'],
    ['Ask Now',          strategist.ask_now || '(none)'],
    ['Today’s Action',   strategist.today_action || '(none)']
  ];
  const rng = sh.getRange(row,1,banner.length,2).setValues(banner);
  styleLabelValueBlock_(rng, accent);
  return row + banner.length + 1;
}

/**
 * Section 2 — In-Depth Analysis (fixed position at D1)
 * L1: Brief banner (Play / Ask Now / Today)
 * L2: 5-minute plan (Exec Summary, Lineup, Close Sequence, Objections, Risks, Budget, Trade-offs)
 * L3: Deep Narrative (180–250 words) — merged across D..H
 *
 * Params:
 *   sh:         Sheet
 *   startRow:   (ignored; we anchor at row=1)
 *   strategist: Strategist JSON object (nullable)
 *   accent:     hex color string for accents
 *   rootApptId: for the stub message when strategist JSON is missing
 *
 * Returns: next free row (number)
 */
function renderSection2_({ sh, startRow, strategist, accent, rootApptId }) {
  const COL = 5;        // E
  let row = 1;          // E1 anchor (top-right)

  // How many columns the bullet BODY spans (keeps lines roomy)
  const LIST_MERGE_COLS = 2; // E..H

  const fmtMoney = (v) => (v==null || v==='') ? '' : ('$'+Number(v).toLocaleString());
  const fmtPercentMaybe = (v) => {
    if (v==null || v==='') return '';
    const n = Number(v);
    return (n <= 1 && n >= 0) ? (Math.round(n*1000)/10)+'%' : (Math.round(n*10)/10)+'%';
  };

  // Title helper: merge E:F for all titles
  const setHdr = (label, size=12) => {
    sh.getRange(row, COL, 1, 2).merge().setValue(label).setFontWeight('bold').setFontSize(size);
    row += 2;
  };
  // List helper with merged BODY rows; Title row merges E:F
  const titledListAt = (title, items) => {
    sh.getRange(row, COL, 1, 2).merge().setValue(title).setFontWeight('bold'); row++;
    const arr = (items && items.length) ? items.map(x => [x]) : [['(none)']];
    sh.getRange(row, COL, arr.length, 1).setValues(arr);
    const mergeRange = sh.getRange(row, COL, arr.length, LIST_MERGE_COLS);
    mergeRange.mergeAcross().setWrap(true);
    zebra_(mergeRange, accent);
    outline_(sh, mergeRange);
    row += arr.length + 1;
  };
  const tableAt = (data, headerLabels, opts) => {
    if (!data || !data.length) return;
    const rng = sh.getRange(row, COL, data.length, data[0].length).setValues(data);
    styleTable_(rng, headerLabels, Object.assign({ accentHex: accent }, opts || {}));
    row += data.length + 1;
  };
  const labelValueAt = (pairs) => {
    const rng = sh.getRange(row, COL, pairs.length, 2).setValues(pairs);
    styleLabelValueBlock_(rng, accent);
    row += pairs.length + 1;
  };

  // ===== Title =====
  setHdr('2. In-Depth Analysis');

  // Stub if strategist missing
  if (!strategist) {
    const stub = [[
      '(Strategist analysis not found yet — run runStrategistAnalysisForRoot("' + rootApptId + '"))',
      '', '', ''
    ]];
    const stubRange = sh.getRange(row, COL, 1, LIST_MERGE_COLS).setValues(stub);
    stubRange.mergeAcross().setWrap(true);
    zebra_(stubRange, accent); outline_(stubRange);
    row += 3;
    return row;
  }

  // ===== LAYER 1 — Brief Banner =====
  labelValueAt([
    ['Recommended Play', strategist.recommended_play || '(none)'],
    ['Ask Now',          strategist.ask_now || '(none)'],
    ['Today’s Action',   strategist.today_action || '(none)']
  ]);

  // ===== LAYER 2 — 5-minute Plan =====
  // Executive Summary
  if (Array.isArray(strategist.executive_summary) && strategist.executive_summary.length) {
    sh.getRange(row, COL, 1, 2).merge().setValue('Executive Summary').setFontWeight('bold'); row++;
    const bullets = strategist.executive_summary.map(b => [b]);
    sh.getRange(row, COL, bullets.length, 1).setValues(bullets);
    const mergeRange = sh.getRange(row, COL, bullets.length, LIST_MERGE_COLS);
    mergeRange.mergeAcross().setWrap(true);
    zebra_(mergeRange, accent); outline_(mergeRange);
    row += bullets.length + 1;
  }

  // Viewing Lineup + Strategy
  if (Array.isArray(strategist.viewing_lineup) && strategist.viewing_lineup.length) {
    titledListAt('b. Diamond Viewing Strategy — Lineup (concise)', strategist.viewing_lineup);
  }

  // Close Sequence + Strategy
  if (Array.isArray(strategist.close_sequence) && strategist.close_sequence.length) {
    titledListAt('c. Strategic Approach to Close — Sequence', strategist.close_sequence);
  }

  // Objections & Replies (title merges E:F)
  if (Array.isArray(strategist.top_objections) && strategist.top_objections.length) {
    sh.getRange(row, COL, 1, 2).merge().setValue('Objections & Replies').setFontWeight('bold'); row++;
    const rows = strategist.top_objections.map(o => [o.objection || '', o.reply || '']);
    tableAt(rows, ['Objection','Reply (verbatim line included)'], { wrapCols:[2] });
  }

  // === Client Priorities for Closing (from Strategist) ===
  if (strategist && strategist.client_priorities) {
    const pr = strategist.client_priorities || {};
    // Top priorities
    sh.getRange(row, COL, 1, 2).merge().setValue('Client Priorities — Top').setFontWeight('bold'); row++;
    const topArr = (pr.top_priorities || []).map(x=>[x]).concat(pr.top_priorities && pr.top_priorities.length ? [] : [['(none)']]);
    const topRng = sh.getRange(row, COL, topArr.length, 1).setValues(topArr);
    const topMerge = sh.getRange(row, COL, topArr.length, LIST_MERGE_COLS); topMerge.mergeAcross().setWrap(true);
    zebra_(topMerge, accent); outline_(topMerge);
    row += topArr.length + 1;

    // Non-negotiables
    sh.getRange(row, COL, 1, 2).merge().setValue('Client Priorities — Non-negotiables').setFontWeight('bold'); row++;
    const nnArr = (pr.non_negotiables || []).map(x=>[x]).concat(pr.non_negotiables && pr.non_negotiables.length ? [] : [['(none)']]);
    const nnRng = sh.getRange(row, COL, nnArr.length, 1).setValues(nnArr);
    const nnMerge = sh.getRange(row, COL, nnArr.length, LIST_MERGE_COLS); nnMerge.mergeAcross().setWrap(true);
    zebra_(nnMerge, accent); outline_(nnMerge);
    row += nnArr.length + 1;

    // Nice-to-haves
    sh.getRange(row, COL, 1, 2).merge().setValue('Client Priorities — Nice-to-haves').setFontWeight('bold'); row++;
    const nhArr = (pr.nice_to_haves || []).map(x=>[x]).concat(pr.nice_to_haves && pr.nice_to_haves.length ? [] : [['(none)']]);
    const nhRng = sh.getRange(row, COL, nhArr.length, 1).setValues(nhArr);
    const nhMerge = sh.getRange(row, COL, nhArr.length, LIST_MERGE_COLS); nhMerge.mergeAcross().setWrap(true);
    zebra_(nhMerge, accent); outline_(nhMerge);
    row += nhArr.length + 1;
  }

  // ===== LAYER 3 — Deep Narrative (title merges E:F) =====
  sh.getRange(row, COL, 1, 2).merge().setValue('Deep Narrative (180–250 words)').setFontWeight('bold'); row++;
  const narrative = String(
    strategist.where_customer_stands_narrative ||
    strategist.free_response || ''
  ).trim();

  if (narrative) {
    const narRange = sh.getRange(row, COL, 1, 2); // E..F only
    narRange.merge().setValue(narrative).setWrap(true);
    outline_(sh, narRange);
    row += 2;
  } else {
    const r = sh.getRange(row, COL, 1, 2).setValues([['(none)']]);
    r.mergeAcross().setWrap(true);
    zebra_(r, accent); outline_(r);
    row += 2;
  }

  return row;
}

/** Canonicalize a diamond type string to 'lab' | 'natural' | '' (no IGI/GIA heuristics) */
function canonicalizeDiamondType_(s){
  const x = String(s||'').toLowerCase();

  // Lab
  if (/\blab[-\s]?grown\b/.test(x)) return 'lab';
  if (/\blab\s*diamond\b/.test(x))  return 'lab';

  // Natural
  if (/\bnatural(\s*diamond)?\b/.test(x)) return 'natural';
  if (/\bmined\b/.test(x))                return 'natural';

  // Unknown
  return '';
}


function normalizeScribe_(obj){
  obj = obj && typeof obj === 'object' ? obj : {};

  // Ensure nested shape
  obj.customer_profile = obj.customer_profile && typeof obj.customer_profile === 'object'
    ? obj.customer_profile : {};

  // 1) Mirror customer_name if it's only present at the top (backward support)
  if (typeof obj.customer_name === 'string' && obj.customer_name.trim() && !obj.customer_profile.customer_name){
    obj.customer_profile.customer_name = obj.customer_name.trim();
  }

  // 2) Move any lingering priorities into customer_profile if present (old structure)
  //    (We only move these three; the ranked priorities moved to Strategist, not Scribe.)
  if (obj.client_priorities && typeof obj.client_priorities === 'object'){
    const cp = obj.client_priorities;
    if (cp.occasion_intent && !obj.customer_profile.occasion_intent){
      obj.customer_profile.occasion_intent = cp.occasion_intent;
    }
    if (Array.isArray(cp.comm_prefs) && !obj.customer_profile.comm_prefs){
      obj.customer_profile.comm_prefs = cp.comm_prefs;
    }
    if (Array.isArray(cp.decision_makers) && !obj.customer_profile.decision_makers){
      obj.customer_profile.decision_makers = cp.decision_makers;
    }
    // Remove the old container to avoid confusion (the ranked lists now belong in Strategist)
    try { delete obj.client_priorities; } catch(_){}
  }

  // --- DIAMOND TYPE canonicalization (only exact phrases, no brand heuristics) ---
    if (!obj.diamond_specs) obj.diamond_specs = {};
    let dt = obj.diamond_specs.lab_or_natural;
    const canon = canonicalizeDiamondType_(dt);
    if (canon) {
      obj.diamond_specs.lab_or_natural = canon;
    } else if (dt === '' || dt == null) {
      obj.diamond_specs.lab_or_natural = null; // leave null if unclear
    }

  // 3) Conf block sanity (keep as numbers 0..1 if present)
  if (obj.conf && typeof obj.conf === 'object'){
    ['budget','timeline','diamond'].forEach(k=>{
      if (obj.conf[k] != null) {
        const v = Number(obj.conf[k]);
        if (isFinite(v)) obj.conf[k] = Math.max(0, Math.min(1, v));
        else obj.conf[k] = null;
      }
    });
  }

  return obj;
}




/** Writes a label/value at A/B and a hidden JSON path in C (used for corrections). */
function writeLVWithPath_(sh, row, label, value, jsonPath){
  sh.getRange(row,1).setValue(label);           // A: label
  sh.getRange(row,2).setValue(value);           // B: value (editable)
  sh.getRange(row,3).setValue(jsonPath||'');    // C: hidden path
}


/** Return a Set of JSON paths whose sibling *_confidence <= TH (default 0.69). */
function listLowConfidencePaths_(obj, TH = 0.69) {
  const out = [];
  (function walk(path, o) {
    if (!o || typeof o !== 'object') return;
    for (const k in o) {
      const v = o[k];
      if (/_confidence$/i.test(k) && typeof v === 'number' && v <= TH) {
        const base = k.replace(/_confidence$/i, '');
        out.push(path ? (path + '.' + base) : base);
      } else if (v && typeof v === 'object') {
        walk(path ? (path + '.' + k) : k, v);
      }
    }
  })('', obj || {});
  return new Set(Array.from(new Set(out)));
}

function listLowConfidencePathsFromConf_(s, TH=0.69){
  const out = new Set();
  if (s && s.conf){
    if (typeof s.conf.budget  === 'number' && s.conf.budget  <= TH){ out.add('budget_low'); out.add('budget_high'); }
    if (typeof s.conf.timeline=== 'number' && s.conf.timeline<= TH){ out.add('timeline'); }
    if (typeof s.conf.diamond === 'number' && s.conf.diamond <= TH){
      ['diamond_specs.lab_or_natural','diamond_specs.carat','diamond_specs.color',
       'diamond_specs.clarity','diamond_specs.ratio','diamond_specs.cut_polish_sym']
       .forEach(p => out.add(p));
    }
  }
  return out;
}



function ensureReportConfig_(reportSS, opts){
  const rootApptId = String(opts.rootApptId||'').trim();
  const reportId   = String(opts.reportId||reportSS.getId()).trim();
  let sh = reportSS.getSheetByName('_Config');
  if (!sh) {
    sh = reportSS.insertSheet('_Config');
    try { sh.hideSheet(); } catch(_){}
    sh.appendRow(['ROOT_APPT_ID', rootApptId]);
    sh.appendRow(['CONTROLLER_URL', WEBAPP_EXEC_URL_()]);
    sh.appendRow(['REPORT_REANALYZE_TOKEN',
      PropertiesService.getScriptProperties().getProperty('REPORT_REANALYZE_TOKEN') || ''
    ]);
    sh.appendRow(['REPORT_ID', reportId]);
    return;
  }
  const vals = sh.getRange(1,1,sh.getLastRow(),2).getValues();
  const map = {}; vals.forEach(r => { if (r[0]) map[String(r[0]).trim()] = String(r[1]||'').trim(); });
  const want = {
    ROOT_APPT_ID: rootApptId,
    CONTROLLER_URL: WEBAPP_EXEC_URL_(),
    REPORT_REANALYZE_TOKEN: PropertiesService.getScriptProperties().getProperty('REPORT_REANALYZE_TOKEN') || '',
    REPORT_ID: reportId
  };
  Object.keys(want).forEach(k=>{
    const cur = map[k] || '';
    const need = String(want[k]||'');
    if (cur !== need){
      const rowIdx = vals.findIndex(r => String(r[0]).trim() === k);
      if (rowIdx >= 0) sh.getRange(rowIdx+1, 2).setValue(need);
      else sh.appendRow([k, need]);
    }
  });
}

function collectCorrectionsFromTab_(sh){
  const last = sh.getLastRow(); if (last < 2) return {};
  const vals = sh.getRange(1,1,last,3).getValues(); // A=label, B=value, C=jsonPath
  const patch = {};

  const toNumberIfNumericLabel = (label, val) => {
    const numHints = /Carat|Width|Low|High|mm|\(numeric\)/i;
    if (!numHints.test(label)) return val;
    const n = Number(String(val).replace(/,/g,''));
    return isNaN(n) ? val : n;
  };

  const splitIfListPath = (path, val) => {
    // Treat these paths as " • " joined lists in the sheet UI
    const listPaths = new Set([
      'customer_profile.comm_prefs',
      'customer_profile.decision_makers'
    ]);

    // (Back-compat: if you still have any client_priorities rows on older tabs, leave this.)
    if (/^client_priorities\./.test(path)) {
      if (typeof val === 'string') {
        return val.split('•').map(s=>s.trim()).filter(Boolean);
      }
    }

    if (listPaths.has(path) && typeof val === 'string') {
      return val.split('•').map(s=>s.trim()).filter(Boolean);
    }

    return val;
  };

  for (let r = 0; r < vals.length; r++){
    const label = String(vals[r][0]||'').trim();
    const raw   = vals[r][1];
    const path  = String(vals[r][2]||'').trim();
    if (!path) continue;

    // 1) Coerce numeric by label hint
    let out = toNumberIfNumericLabel(label, raw);
    // 2) Split list fields to arrays by " • "
    out = splitIfListPath(path, out);

    // 3) Write into nested patch
    setByPath_(patch, path, out);
  }
  return patch;
}


/**
 * Re-analyze a consult using rep corrections from the currently-open report tab.
 * - Reads col B values where col C has dotted JSON paths
 * - Merges into the newest Scribe JSON → saves a *Corrected* Scribe JSON (new file)
 * - Regenerates Strategist JSON
 * - Re-renders the Client Summary / Consult tab
 *
 * @param {string} rootApptId
 * @param {GoogleAppsScript.Spreadsheet.Sheet=} consultSheetOpt   optional: a specific sheet/tab to read corrections from
 */
function consult_reanalyzeFromCorrections_(rootApptId, consultSheetOpt) {
  if (!rootApptId) throw new Error('consult_reanalyzeFromCorrections_: missing rootApptId');

  // --- tiny fallback: collectCorrectionsFromTab_ if you don’t already have it defined
  if (typeof collectCorrectionsFromTab_ !== 'function') {
    this.collectCorrectionsFromTab_ = function (sh) {
      const last = sh.getLastRow(); if (last < 2) return {};
      const values = sh.getRange(1,1,last,3).getValues(); // A=label, B=value(editable), C=jsonPath
      const patch = {};
      for (let r=0; r<values.length; r++){
        const label = String(values[r][0]||'').trim();
        const val   = values[r][1];
        const path  = String(values[r][2]||'').trim();
        if (!path) continue;

        let out = val;
        // Numeric hint
        if (/Carat|Width|Low|High|mm|\(numeric\)/i.test(label)){
          const num = Number(val);
          if (!isNaN(num)) out = num;
        }
        // Split durability concerns " • " into array
        if (/Durability Concerns/i.test(label) && typeof val === 'string'){
          const parts = val.split('•').map(s=>s.trim()).filter(Boolean);
          out = parts;
        }

        if (typeof setByPath_ === 'function') {
          setByPath_(patch, path, out);
        } else {
          // micro setter if you don’t have setByPath_
          const parts = path.split('.').map(s=>s.trim()).filter(Boolean);
          let cur = patch;
          for (let i=0;i<parts.length-1;i++){
            const p = parts[i];
            if (!cur[p] || typeof cur[p] !== 'object') cur[p] = {};
            cur = cur[p];
          }
          cur[parts[parts.length-1]] = out;
        }
      }
      return patch;
    };
  }

  // --- Resolve Master + folder
  const ss = SpreadsheetApp.openById(PROP_('SPREADSHEET_ID'));
  const apId = getApFolderIdForRoot_(ss, rootApptId);
  if (!apId) throw new Error('No RootAppt Folder ID for ' + rootApptId);
  const ap = DriveApp.getFolderById(apId);

  // --- Load the newest Scribe JSON (prefer corrected; else base)
  const sFolderIt = ap.getFoldersByName('04_Summaries');
  if (!sFolderIt.hasNext()) throw new Error('No 04_Summaries for ' + rootApptId);
  const sFolder = sFolderIt.next();

  function newestByRegexInFolder_(folder, re){
    let newest=null, ts=0, it=folder.getFiles();
    while (it.hasNext()){
      const f = it.next();
      if (!re.test(f.getName())) continue;
      const t=f.getDateCreated().getTime();
      if (t>ts){ ts=t; newest=f; }
    }
    return newest;
  }

  const scribeCorrected = newestByRegexInFolder_(sFolder, /__summary_corrected_.*\.json$/i);
  const scribeBase      = newestByRegexInFolder_(sFolder, /__summary_.*\.json$/i);
  const scribeFile      = scribeCorrected || scribeBase;
  if (!scribeFile) throw new Error('No Scribe JSON found for ' + rootApptId);

  const scribeObj = JSON.parse(scribeFile.getBlob().getDataAsString('UTF-8'));

  // --- Optional transcript text (best effort)
  let transcript = '';
  let transcriptUrl = '';
  const tFolderIt = ap.getFoldersByName('03_Transcripts');
  if (tFolderIt.hasNext()){
    const tf = tFolderIt.next();
    const txt = newestByRegexInFolder_(tf, /\.txt$/i);
    if (txt){
      transcript = txt.getBlob().getDataAsString('UTF-8');
      transcriptUrl = 'https://drive.google.com/file/d/' + txt.getId() + '/view';
    }
  }

  // --- Read corrections from the active report tab (or supplied tab)
  let report;
  try { report = SpreadsheetApp.openById(getReportIdForRoot_(rootApptId)); } catch(_){}
  const sourceSheet = consultSheetOpt || (report ? report.getActiveSheet() : null);
  if (!sourceSheet) throw new Error('No consult sheet to read corrections from.');

  // Build patch from the sheet and filter to allowed dotted paths
  const patchRaw = collectCorrectionsFromTab_(sourceSheet);
  const patch    = filterPatchByAllowed_(patchRaw);   // <-- whitelist to allowed Scribe paths

  // --- Merge: base Scribe + patch → corrected Scribe
  let corrected;
  if (typeof mergeDeep_ === 'function') {
    corrected = mergeDeep_(scribeObj, patch);
  } else {
    // minimal deep merge...
    corrected = JSON.parse(JSON.stringify(scribeObj));
    (function merge(a, b){
      if (b == null) return;
      if (Array.isArray(a) || Array.isArray(b)) return b;
      if (typeof b !== 'object') return b;
      Object.keys(b).forEach(k=>{
        if (a && typeof a === 'object' && k in a) {
          if (typeof a[k] === 'object' && typeof b[k] === 'object' && !Array.isArray(a[k]) && !Array.isArray(b[k])){
            a[k] = merge(a[k], b[k]);
          } else {
            a[k] = b[k];
          }
        } else {
          if (!a || typeof a !== 'object') a = {};
          a[k] = b[k];
        }
      });
      return a;
    })(corrected, patch);
  }
  corrected = normalizeScribe_(corrected);

  // --- Save corrected Scribe JSON (NEW file)
  const correctedUrl = saveCorrectedScribeJson_(ap, rootApptId, corrected);

  // 3) Strategist memo → extract on corrected Scribe
  let memoPayload    = buildStrategistMemoPayload_(corrected, transcript || '', '');
  const memoText  = openAIResponses_TextOnly_(memoPayload);
  strat_writeDebug_(ap, rootApptId, 'memo_from_patch', memoText);

  let extractPayload = buildStrategistExtractPayload_(memoText, corrected);
  const strategistObj = openAIResponses_(extractPayload);
  const strategistUrl = saveStrategistJson_(ap, rootApptId, strategistObj);

  // Ensure we pass the *same* report we are editing
  const reportId = report ? report.getId() : '';

  // --- Re-render the consult tab (right-hand analysis at D1)
  const apISO = getApptIsoForRoot_(ss, rootApptId) || new Date().toISOString();
  upsertClientSummaryTab_(rootApptId, corrected, apISO, transcriptUrl, strategistObj, { reportId });
}


function migrateTranscriptUrlToHasTranscript() {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing 00_Master Appointments');

  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h||'').trim());
  const iUrl = header.indexOf('Transcript URL');
  let iHas = header.indexOf('Has Transcript');

  // If neither exists, just create Has Transcript and stop.
  if (iUrl < 0 && iHas < 0) {
    sh.insertColumnAfter(sh.getLastColumn());
    const col = sh.getLastColumn();
    sh.getRange(1, col).setValue('Has Transcript');
    return;
  }

  // Ensure Has Transcript column exists
  if (iHas < 0) {
    sh.insertColumnAfter(sh.getLastColumn());
    iHas = sh.getLastColumn()-1;
    sh.getRange(1, iHas+1).setValue('Has Transcript');
  }

  // If Transcript URL exists, convert its non-empty rows to TRUE, else FALSE
  const last = sh.getLastRow();
  if (last >= 2 && iUrl >= 0) {
    const urls = sh.getRange(2, iUrl+1, last-1, 1).getValues().flat();
    const out  = urls.map(v => [String(v||'').trim() ? 'TRUE' : 'FALSE']);
    sh.getRange(2, iHas+1, out.length, 1).setValues(out);
    // Optional: clear old URL column header to mark legacy
    sh.getRange(1, iUrl+1).setValue('Transcript URL (legacy)');
  }
}


function test_upsertClientSummaryTab(){
  // Auto-loads latest Scribe + Strategist from Drive and re-renders the tab.
  rerenderClientSummaryTabForRoot_('AP-20250907-003');
}

function test_chat_wire() {
  const root = 'AP-20250907-001';
  const reportId = getReportIdForRoot_(root);
  const out = AC_chatCore_(root, reportId, 'Ping from wire test');
  Logger.log(JSON.stringify(out, null, 2));
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



