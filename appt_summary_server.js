/*** 16 - appt_summary_server.gs (v1.2)
     Appointment Summary (Date Range) — "00_Master Appointments" as sole source of truth

     UI: as_openAppointmentSummary() opens dialog (dlg_appt_summary_v1.html)
     Data: single batched read from 00_Master Appointments, minimal span, cached header map.

     Columns returned:
       Visit Date, RootApptID, Customer, Phone, Email, Visit Type, Visit #, SO#, Brand,
       Sales Stage, Conversion Status, Custom Order Status, Center Stone Order Status
***/

const AS_CFG = Object.freeze({
  SHEET_NAME: '00_Master Appointments',
  HEADERS: [
    'RootApptID',
    'SO#',
    'Customer',
    'Phone',
    'Email',
    'Visit Type',
    'Visit #',
    'Assigned Rep',
    'Assisted Rep',
    'Sales Stage',
    'Conversion Status',
    'Custom Order Status',
    'Center Stone Order Status',
    'Visit Date',
    'Brand'
  ],

  CACHE_TTL_SEC: 600 // 10 minutes
});

// Bump this to invalidate header-index cache safely when headers change
const AS_HDR_CACHE_VER = 'b1';


function as_openAppointmentSummary() {
  const t = HtmlService.createTemplateFromFile('dlg_appt_summary_v1');
  try {
    t.BOOTSTRAP = as_bootstrap(); // structured object
  } catch (e) {
    // Ensure the template still renders and shows a clear error
    t.BOOTSTRAP = { ok: false, error: e && e.message ? e.message : String(e) };
  }
  const html = t.evaluate().setWidth(1050).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Appointment Summary');
}


/** Bootstrap: min/max visit dates + default last-30-day range. */
function as_bootstrap() {
  try {
    const { sheet, headerRow, idx } = as_getSheetAndIndex_();
    const vDateCol = idx['visit date'];
    if (!vDateCol) return { ok: false, error: 'Visit Date column not found on "00_Master Appointments".' };

    const lastRow = sheet.getLastRow();
    const todayISO = as_toISO_(new Date());

    if (lastRow <= headerRow) {
      return {
        ok: true,
        minISO: todayISO,
        maxISO: todayISO,
        defStartISO: todayISO,
        defEndISO: todayISO,
        rowCount: 0,
        brands: []
      };
    }

    const firstDataRow = headerRow + 1;
    const dateVals = sheet.getRange(firstDataRow, vDateCol, lastRow - headerRow, 1).getValues();

    let minMs = Number.POSITIVE_INFINITY, maxMs = Number.NEGATIVE_INFINITY, count = 0;
    for (const [d] of dateVals) {
      const ms = as_coerceDateMs_(d);
      if (ms == null) continue;
      if (ms < minMs) minMs = ms;
      if (ms > maxMs) maxMs = ms;
      count++;
    }

    const today = new Date();
    const defEnd = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const defStart = new Date(defEnd); defStart.setDate(defStart.getDate() - 30);

    const minISO = isFinite(minMs) ? as_toISO_(new Date(minMs)) : todayISO;
    const maxISO = isFinite(maxMs) ? as_toISO_(new Date(maxMs)) : todayISO;

    // Build unique Brand list (blank‑safe, sorted)
    let brands = [];
    const bCol = idx['brand'];
    if (bCol && lastRow > headerRow) {
      const bVals = sheet.getRange(firstDataRow, bCol, lastRow - headerRow, 1).getValues();
      const set = new Set();
      for (const [b] of bVals) {
        const s = String(b || '').trim();
        if (s) set.add(s);
      }
      brands = [...set].sort((a, b) => a.localeCompare(b));
    }

    // Build unique lists for additional filters
    const uniqFromCol = (col) => {
      if (!col || lastRow <= headerRow) return [];
      const vals = sheet.getRange(firstDataRow, col, lastRow - headerRow, 1).getValues();
      const set = new Set();
      for (const [v] of vals) {
        const s = String(v || '').trim();
        if (s) set.add(s);
      }
      return [...set].sort((a, b) => a.localeCompare(b));
    };

    const stages       = uniqFromCol(idx['sales stage']);
    const conversions  = uniqFromCol(idx['conversion status']);
    const customOrders = uniqFromCol(idx['custom order status']);
    const centerStones = uniqFromCol(idx['center stone order status']);


    return {
      ok: true,
      minISO,
      maxISO,
      defStartISO: as_toISO_(as_clampDate_(defStart, minMs, maxMs)),
      defEndISO:   as_toISO_(as_clampDate_(defEnd,   minMs, maxMs)),
      rowCount: count,
      brands,
      stages,
      conversions,
      customOrders,
      centerStones
    };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}

/** Run summary between [startISO, endISO] inclusive. */
function as_runAppointmentSummary(startISO, endISO, filters) {
  try {
    const { sheet, headerRow, idx, colSpan } = as_getSheetAndIndex_(true);
    if (!idx['visit date']) return { ok: false, error: 'Visit Date column not found on "00_Master Appointments".' };

    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return { ok: true, rows: [], total: 0 };

    const startMs = as_coerceDateMs_(startISO);
    const endMs = as_coerceDateMs_(endISO);
    if (startMs == null || endMs == null) return { ok: false, error: 'Invalid start/end date.' };

    const firstDataRow = headerRow + 1;
    const rng = sheet.getRange(firstDataRow, colSpan.minCol, lastRow - headerRow, colSpan.width);
    const values = rng.getValues();

    const ixInSpan = (hName) => idx[hName] ? (idx[hName] - colSpan.minCol) : null;
    const iVisitDate = ixInSpan('visit date');

    const wanted = {
      root: ixInSpan('rootapptid'),
      so:   ixInSpan('so#'),
      customer: ixInSpan('customer'),
      phone:    ixInSpan('phone'),
      email:    ixInSpan('email'),
      vtype:    ixInSpan('visit type'),
      vnum:     ixInSpan('visit #'),
      assigned: ixInSpan('assigned rep'),
      assisted: ixInSpan('assisted rep'),
      stage:    ixInSpan('sales stage'),
      conv:     ixInSpan('conversion status'),
      cos:      ixInSpan('custom order status'),
      csos:     ixInSpan('center stone order status'),
      brand:    ixInSpan('brand')
    };

    const F = filters || {};
    // Back-compat: if caller passed just an array, treat it as Brand filters
    if (Array.isArray(F)) F = { brands: F };

    const toSet = (arr) => new Set((arr || [])
      .map(s => String(s).toLowerCase().trim())
      .filter(Boolean));

    const brandsSet       = toSet(F.brands);
    const stagesSet       = toSet(F.stages);
    const conversionsSet  = toSet(F.conversions);
    const customOrdersSet = toSet(F.customOrders);
    const centerStonesSet = toSet(F.centerStones);

    const out = [];

    for (let r = 0; r < values.length; r++) {
      const row = values[r];
      const vMs = as_coerceDateMs_(iVisitDate != null ? row[iVisitDate] : null);
    
      if (vMs == null || vMs < startMs || vMs > endMs) continue;

      // Apply Brand filter
      if (brandsSet.size) {
        const v = String(as_val(row, wanted.brand) || '').toLowerCase().trim();
        if (!v || !brandsSet.has(v)) continue;
      }
      // Sales Stage
      if (stagesSet.size) {
        const v = String(as_val(row, wanted.stage) || '').toLowerCase().trim();
        if (!v || !stagesSet.has(v)) continue;
      }
      // Conversion
      if (conversionsSet.size) {
        const v = String(as_val(row, wanted.conv) || '').toLowerCase().trim();
        if (!v || !conversionsSet.has(v)) continue;
      }
      // Custom Order
      if (customOrdersSet.size) {
        const v = String(as_val(row, wanted.cos) || '').toLowerCase().trim();
        if (!v || !customOrdersSet.has(v)) continue;
      }
      // Center Stone
      if (centerStonesSet.size) {
        const v = String(as_val(row, wanted.csos) || '').toLowerCase().trim();
        if (!v || !centerStonesSet.has(v)) continue;
      }

      out.push({
        VisitDateISO: as_toISO_(new Date(vMs)),
        RootApptID:   as_val(row, wanted.root),
        Customer:     as_val(row, wanted.customer),
        Phone:        as_val(row, wanted.phone),
        Email:        as_val(row, wanted.email),
        VisitType:    as_val(row, wanted.vtype),
        VisitNum:     as_val(row, wanted.vnum),
        SO:           as_val(row, wanted.so),
        Brand:        as_val(row, wanted.brand),
        AssignedRep:  as_val(row, wanted.assigned),
        AssistedRep:  as_val(row, wanted.assisted),
        SalesStage:   as_val(row, wanted.stage),
        Conversion:   as_val(row, wanted.conv),
        CustomOrder:  as_val(row, wanted.cos),
        CenterStone:  as_val(row, wanted.csos)
      });
    }

    out.sort((a, b) => {
      // 1) Visit Date (ISO strings are comparable safely)
      if (a.VisitDateISO < b.VisitDateISO) return -1;
      if (a.VisitDateISO > b.VisitDateISO) return 1;

      // 2) Brand (case-insensitive; blanks last)
      const ab = (a.Brand || '').toString().toLowerCase();
      const bb = (b.Brand || '').toString().toLowerCase();
      if (ab && bb) {
        const cmp = ab.localeCompare(bb);
        if (cmp !== 0) return cmp;
      } else if (ab && !bb) return -1;
      else if (!ab && bb) return 1;

      // 3) RootApptID
      if ((a.RootApptID || '') < (b.RootApptID || '')) return -1;
      if ((a.RootApptID || '') > (b.RootApptID || '')) return 1;

      // 4) Visit #
      return Number(a.VisitNum || 0) - Number(b.VisitNum || 0);
    });

    return { ok: true, rows: out, total: out.length };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}

/**
 * Build a PDF for the given date range + filters (brands, stages, conversions, etc.).
 * Reuses the same single-read summarizer for integrity/performance.
 * Returns {ok:true, fileName, bytesBase64, mimeType}.
 */
function as_exportAppointmentSummaryPdf(startISO, endISO, filters) {
  try {
    // 1) Compute rows via the same fast path
    const res = as_runAppointmentSummary(startISO, endISO, filters);
    if (!res || !res.ok) return res;
    const rows = res.rows || [];

    // 2) Local HTML escaper (do NOT rely on any client code)
    const htmlEsc = (s) => String(s == null ? '' : s)
      .replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));

    // 3) Summaries for filter meta line
    const F = filters || {};
    const summarize = (arr, labelAll) => (arr && arr.length) ? arr.join(', ') : labelAll;

    const metaHtml = [
      `Date range: <b>${htmlEsc(startISO)}</b> → <b>${htmlEsc(endISO)}</b>`,
      `Brand(s): <b>${htmlEsc(summarize(F.brands,       'All'))}</b>`,
      `Stage(s): <b>${htmlEsc(summarize(F.stages,       'All'))}</b>`,
      `Conversion: <b>${htmlEsc(summarize(F.conversions,'All'))}</b>`,
      `Custom Order: <b>${htmlEsc(summarize(F.customOrders,'All'))}</b>`,
      `Center Stone: <b>${htmlEsc(summarize(F.centerStones,'All'))}</b>`,
      `Rows: <b>${rows.length}</b>`
    ].join('&nbsp;&nbsp;•&nbsp;&nbsp;');

    // 4) HTML → PDF
    const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    @page { size: Letter landscape; margin: 12mm; }
    body { font: 12px/1.45 system-ui,-apple-system,Segoe UI,Roboto,Arial; color:#111; }
    h1 { font-size: 16px; margin: 0 0 6px; }
    .meta { font-size: 12px; color:#555; margin: 0 0 10px; }
    table { width: 100%; border-collapse: collapse; font-size: 11px; }
    thead th { background:#f4f4f4; border:1px solid #ddd; text-align:left; padding:6px 8px; }
    tbody td { border:1px solid #e5e5e5; padding:6px 8px; vertical-align:top; }
    .right { text-align:right; }
  </style>
</head>
<body>
  <h1>Appointment Summary</h1>
  <div class="meta">${metaHtml}</div>
  <table>
    <thead>
      <tr>
        <th>Visit Date</th>
        <th>RootApptID</th>
        <th>Customer</th>
        <th>Phone</th>
        <th>Email</th>
        <th>Visit Type</th>
        <th>Visit #</th>
        <th>SO#</th>
        <th>Brand</th>
        <th>Assigned Rep</th>
        <th>Assisted Rep</th>
        <th>Sales Stage</th>
        <th>Conversion</th>
        <th>Custom Order</th>
        <th>Center Stone</th>
      </tr>
    </thead>
    <tbody>
      ${rows.map(r => `
        <tr>
          <td>${htmlEsc(r.VisitDateISO)}</td>
          <td>${htmlEsc(r.RootApptID)}</td>
          <td>${htmlEsc(r.Customer)}</td>
          <td>${htmlEsc(r.Phone)}</td>
          <td>${htmlEsc(r.Email)}</td>
          <td>${htmlEsc(r.VisitType)}</td>
          <td class="right">${htmlEsc(r.VisitNum)}</td>
          <td>${htmlEsc(r.SO)}</td>
          <td>${htmlEsc(r.Brand)}</td>
          <td>${htmlEsc(r.AssignedRep)}</td>
          <td>${htmlEsc(r.AssistedRep)}</td>
          <td>${htmlEsc(r.SalesStage)}</td>
          <td>${htmlEsc(r.Conversion)}</td>
          <td>${htmlEsc(r.CustomOrder)}</td>
          <td>${htmlEsc(r.CenterStone)}</td>
        </tr>`).join('')}
    </tbody>
  </table>
</body>
</html>`;

    const blob = Utilities.newBlob(html, 'text/html', 'summary.html').getAs('application/pdf');

    return {
      ok: true,
      fileName: `Appointment_Summary_${startISO}_to_${endISO}.pdf`,
      bytesBase64: Utilities.base64Encode(blob.getBytes()),
      mimeType: blob.getContentType()
    };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}



/* ---------------------------- Internals (helpers) ---------------------------- */

function as_getSheetAndIndex_(computeSpan) {
  const ss = SpreadsheetApp.getActive();
  const sheet = as_getTargetSheet_(ss);

  const cache = CacheService.getScriptCache();
  const cacheKey = `as_hdr_${sheet.getSheetId()}_${AS_HDR_CACHE_VER}`; // faster + safe if renamed later

  let cached = cache.get(cacheKey), headerRow, idx;

  if (cached) {
    const parsed = JSON.parse(cached);
    headerRow = parsed.headerRow;
    idx = parsed.idx;
  } else {
    headerRow = as_findHeaderRow_(sheet) || 1;
    const headerValues = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    idx = {};
    for (let c = 0; c < headerValues.length; c++) {
      const raw = String(headerValues[c] || '').trim();
      if (!raw) continue;
      idx[as_fuzzyHeaderKey_(as_norm_(raw))] = c + 1; // 1-based
    }
    cache.put(cacheKey, JSON.stringify({ headerRow, idx }), AS_CFG.CACHE_TTL_SEC);
  }

  if (!computeSpan) return { sheet, headerRow, idx };

  // Compute contiguous min..max span among required columns that exist
  const need = AS_CFG.HEADERS.map(h => as_norm_(h));
  let minCol = Number.POSITIVE_INFINITY, maxCol = 0;
  for (const want of need) {
    const col = idx[want];
    if (col) {
      if (col < minCol) minCol = col;
      if (col > maxCol) maxCol = col;
    }
  }
  if (!isFinite(minCol) || maxCol === 0) {
    const vCol = idx['visit date'];
    if (!vCol) throw new Error('Required columns not found.');
    minCol = vCol; maxCol = vCol;
  }
  return { sheet, headerRow, idx, colSpan: { minCol, maxCol, width: (maxCol - minCol + 1) } };
}

function as_getTargetSheet_(ss) {
  const active = ss.getActiveSheet();
  if (active && active.getName() === AS_CFG.SHEET_NAME) return active;
  const byName = ss.getSheetByName(AS_CFG.SHEET_NAME);
  if (byName) return byName;
  throw new Error(`Sheet "${AS_CFG.SHEET_NAME}" not found. Open it and try again.`);
}

function as_findHeaderRow_(sheet) {
  const maxProbe = Math.min(10, sheet.getLastRow());
  for (let r = 1; r <= maxProbe; r++) {
    const rowVals = sheet.getRange(r, 1, 1, 10).getValues()[0];
    if (rowVals.some(v => String(v || '').trim() !== '')) return r;
  }
  return 1;
}

function as_norm_(s) {
  return String(s || '').toLowerCase().trim();
}

function as_fuzzyHeaderKey_(key) {
  // Canonical → variants (all lowercased)
  const map = [
    ['rootapptid', ['rootapptid','root appt id','root appointment id','root id']],
    ['so#', ['so#','so #','so number','so no','so']],
    ['customer', ['customer','customer name','client','client name','name']],
    ['phone', ['phone','phone number','tel','telephone']],
    ['email', ['email','e-mail','mail']],
    ['visit type', ['visit type','type','appt type','appointment type']],
    ['visit #', ['visit #','visit number','visit no','visit','visit#']],
    ['assigned rep', ['assigned rep','assigned','primary rep','rep']],
    ['assisted rep', ['assisted rep','assistant rep','assisted','helper rep']],
    ['sales stage', ['sales stage','stage','pipeline stage','sales status']],
    ['conversion status', ['conversion status','conversion','converted']],
    ['custom order status', ['custom order status','co status','custom status']],
    ['center stone order status', ['center stone order status','center stone status','cs order status']],
    ['visit date', ['visit date','appointment date','appt date','date']],
    ['brand', ['brand','company','store','business unit','division']]
  ];
  for (const [canon, variants] of map) {
    if (variants.includes(key)) return canon;
  }
  return key;
}

function as_coerceDateMs_(d) {
  if (d == null || d === '') return null;
  if (Object.prototype.toString.call(d) === '[object Date]') {
    if (isNaN(d.getTime())) return null;
    return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
  }
  const s = String(d).trim();
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (m) {
    const dt = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return isNaN(dt.getTime()) ? null : new Date(dt.getFullYear(), dt.getMonth(), dt.getDate()).getTime();
  }
  const dt = new Date(s);
  if (isNaN(dt.getTime())) return null;
  return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate()).getTime();
}

function as_toISO_(d) {
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

function as_clampDate_(d, minMs, maxMs) {
  if (!isFinite(minMs) || !isFinite(maxMs)) return d;
  const ms = new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
  if (ms < minMs) return new Date(minMs);
  if (ms > maxMs) return new Date(maxMs);
  return d;
}

function as_val(row, ix) {
  if (ix == null) return '';
  const v = row[ix];
  if (v == null) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return as_toISO_(v);
  return String(v);
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



