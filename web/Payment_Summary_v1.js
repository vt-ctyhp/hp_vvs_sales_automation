/** PaymentSummary_v1.gs — v1.1 (fix data pull for Date/Time, DocType, etc.)
 * Reads 400_Payments → Payments using the same headers written by rp_submit().
 * No writes. Safe to add to project alongside Payments_v1.gs.
 *
 * Depends on:
 *   - rp_init() → anchor prefill (SO/APPT, order totals, PTD, folder URL)
 *   - rp_getLedgerTarget() → open the ledger and resolve the “Payments” sheet
 */

// Precompiled, module-wide (micro-alloc savings in hot paths)
var PS_RE_DOC_ID = /[-\w]{25,}/;     // Google file id
var PS_RE_NUM_SCRUB = /[^\d.\-]/g;   // numeric scrub


// ---------- PUBLIC API ----------
function ps_init() {
  try {
    // Use the same prefill you use for Record Payment
    var prefill = (typeof rp_init === 'function') ? rp_init() : null;
    if (!prefill) return { ok:false, error:'Prefill failed. No active row?' };

    var anchor = {
      anchorType: prefill.anchorType || (prefill.brand ? 'SO' : 'APPT'),
      soNumber:   String(prefill.soNumber || prefill.SO || prefill['SO#'] || '').trim(),
      rootApptId: String(prefill.rootApptId || prefill.rootApptID || prefill.APPT_ID || '').trim()
    };

    var history = ps_fetchHistoryForAnchor_(anchor);
    return { ok:true, prefill: prefill, history: history };
  } catch (e) {
    return { ok:false, error: e && e.message ? e.message : String(e) };
  }
}

// ---------- CORE ----------
function ps_fetchHistoryForAnchor_(anchor) {
  var tgt = (typeof rp_getLedgerTarget === 'function') ? rp_getLedgerTarget() : null;
  var sh = tgt && (tgt.sh || tgt.sheet);
  if (!sh) return { entries:[], totals:{}, warn:'Ledger not reachable.' };

  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return { entries:[], totals:{}, warn:'No rows yet.' };

  // Header map (1-based)
  var header = sh.getRange(1,1,1,lc).getValues()[0] || [];
  var H = ps_headerMap_(header);

  // Pick columns using robust variants (match what rp_submit writes)
  var cSO        = ps_pick_(H, ['SO#','SO','SO Number','Sales Order','Sales Order #']);
  var cAppt      = ps_pick_(H, ['RootApptID','APPT_ID','Root Appt ID','Root Appt','Appointment ID']);
  // NEW: semantic & relationship columns
  var cDocStatus   = ps_pick_(H, ['DocStatus','Status']);
  var cDocRole     = ps_pick_(H, ['DocRole','Role']);
  var cSupersedes  = ps_pick_(H, ['SupersedesDoc#','Supersedes','Replaces','ReplacesDoc#']);
  var cAppliesTo   = ps_pick_(H, ['AppliesToDoc#','Applies To','SettlesDoc#','Settles']);
  var cDocType   = ps_pick_(H, ['DocType','Doc Type','Document Type','Type']);
  var cDocNo     = ps_pick_(H, ['Doc #','Doc Number','DocNumber']);
  var cLinesJSON = ps_pick_(H, ['LinesJSON','Lines JSON','Line Items JSON','Lines']);

  var cSubtotal  = ps_pick_(H, ['Subtotal','Lines Subtotal','LinesSubtotal']);                // rp_submit → "Subtotal"
  var cMethod    = ps_pick_(H, ['Method','Payment Method']);                                  // rp_submit → "Method"
  var cRef       = ps_pick_(H, ['Reference','Ref','Auth Code','Check #','Check Number']);
  var cNotes     = ps_pick_(H, ['Notes','Note']);
  var cPDF       = ps_pick_(H, ['PDF URL','PDF Link','PDF','PDFURL','Pdf URL']);
  var cDocURL    = ps_pick_(H, ['Doc URL','Doc Link','Document URL','Google Doc URL','Doc','DocURL']);
  var cGross     = ps_pick_(H, ['AmountGross','Payment Amount','Amount','Total Paid']);       // rp_submit → "AmountGross"
  var cPayDT     = ps_pick_(H, ['PaymentDateTime','Payment Date/Time','Payment Date','Payment Timestamp']); // rp_submit → "PaymentDateTime"
  var cSubmitted = ps_pick_(H, ['Submitted Date/Time','Submitted At','Submitted','Date/Time','Date','Timestamp','Created At','CreatedAt']); // rp_submit → "Submitted Date/Time"
  var cDueDate   = ps_pick_(H, ['DueDate','Due Date','InvoiceDueDate','Invoice Due Date','DocDueDate','Doc Due Date']);
  var cGroupId   = ps_pick_(H, ['InvoiceGroupID','InvoiceGroup','Invoice Group ID','Invoice Group','GroupID','Group ID','BasketID']);

  // Read all data
  var vals = sh.getRange(2,1,lr-1,lc).getValues();
  var entries = [];

  var so = ps_clean_(String(anchor.soNumber || ''));
  var appt = ps_clean_(String(anchor.rootApptId || ''));

  var soEq = (typeof rp_soEq === 'function') ? rp_soEq : ps_soEq_;

  // Cache timezone once per execution (used for dateDisplay)
  var tz = Session.getScriptTimeZone();

  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];

    var rSO   = cSO     ? r[cSO-1]   : '';
    var rAppt = cAppt   ? r[cAppt-1] : '';

    var match = false;
    if (so && rSO && soEq(String(rSO), so)) match = true;
    if (!match && appt && rAppt && ps_clean_(String(rAppt)) === appt) match = true;
    if (!match) continue;

    // Prefer explicit PaymentDateTime (receipts), otherwise Submitted Date/Time (all docs)
    var whenAny = (cPayDT && r[cPayDT-1]) ? r[cPayDT-1] : (cSubmitted ? r[cSubmitted-1] : null);
    var when = ps_parseDate_(whenAny);

    var docType = cDocType ? String(r[cDocType-1] || '').trim() : '';
    if (!docType) docType = ps_guessDocType_(r, cGross); // tiny fallback

    var docNo = cDocNo ? String(r[cDocNo-1] || '').trim() : '';
    if (!docNo) {
      // Use BasketID or PAYMENT_ID if available (covered in cDocNo pick above). If still blank, try deriving from Doc URL filename.
      if (cDocURL && r[cDocURL-1]) {
        var s = String(r[cDocURL-1] || '');
        var m = s.match(PS_RE_DOC_ID); // doc id
        docNo = m && m[0] ? m[0].slice(-6).toUpperCase() : '';
      }
    }

    var lines = ps_parseLines_(cLinesJSON ? r[cLinesJSON-1] : '');
    var subtotal = ps_num_(cSubtotal ? r[cSubtotal-1] : 0);
    if (!(subtotal > 0) && lines.length) {
      subtotal = lines.reduce(function(s, L){ return s + (ps_num_(L.qty||1) * ps_num_(L.amt||0)); }, 0);
    }

    var payment = ps_num_(cGross ? r[cGross-1] : 0); // non-zero for receipts
    var method  = cMethod ? String(r[cMethod-1] || '').trim() : '';
    var ref     = cRef    ? String(r[cRef-1]    || '').trim() : '';
    var notes   = cNotes  ? String(r[cNotes-1]  || '').trim() : '';
    var pdfUrl  = cPDF    ? String(r[cPDF-1]    || '').trim() : '';
    var docUrl  = cDocURL ? String(r[cDocURL-1] || '').trim() : '';


    // NEW: read status/role/relations; default status to ISSUED for legacy rows
    var docStatus  = cDocStatus  ? String(r[cDocStatus-1]  || '').trim().toUpperCase() : '';
    if (!docStatus) docStatus = 'ISSUED';
    var docRole    = cDocRole    ? String(r[cDocRole-1]    || '').trim().toUpperCase() : '';
    var supersedes = cSupersedes ? String(r[cSupersedes-1] || '').trim()               : '';
    var appliesTo  = cAppliesTo  ? String(r[cAppliesTo-1]  || '').trim()               : '';

    var due = cDueDate ? ps_parseDate_(r[cDueDate-1]) : null;
    var grpId = cGroupId ? String(r[cGroupId-1] || '').trim() : '';
    entries.push({
      when: when ? when.toISOString() : '',
      dateDisplay: when ? Utilities.formatDate(when, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '',
      docType: docType,
      docNumber: docNo,
      method: method,
      reference: ref,
      notes: notes,
      pdfUrl: pdfUrl,
      docUrl: docUrl,
      lines: lines,
      linesSubtotal: subtotal,
      payment: payment,

      // NEW
      docStatus: docStatus,
      docRole:   docRole,
      supersedes: supersedes,
      appliesTo:  appliesTo,
      dueDate: due ? due.toISOString() : '',
      dueDateDisplay: due ? Utilities.formatDate(due, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      invoiceGroupId: grpId,
      soNumber: rSO ? String(rSO).trim() : ''
    });

  }

  // Sort Oldest→Newest (numeric epoch; identical ordering for ISO dates)
  // Use a stable secondary key (original index) to preserve order when dates tie.
  var _INF = -9007199254740991; // sentinel for blank dates
  var keyed = entries.map(function(e, i){
    return { k: e.when ? Date.parse(e.when) : _INF, i: i, e: e };
  });
  keyed.sort(function(a, b){ return (a.k - b.k) || (a.i - b.i); });
  entries = keyed.map(function(x){ return x.e; });

  // Totals — only ISSUED docs contribute
  var issued = entries.filter(function(e){
    var s = String(e.docStatus || 'ISSUED').toUpperCase();
    return s === 'ISSUED';
  });

  var totals = {
    invoicesLinesSubtotal: issued.filter(function(e){ return /invoice/i.test(e.docType); })
                                .reduce(function(t,e){ return t + (e.linesSubtotal || 0); }, 0),
    totalPayments: issued.filter(function(e){ return /receipt/i.test(e.docType); })
                        .reduce(function(t,e){ return t + (e.payment || 0); }, 0),
    byMethod: (function(){
      var m = {};
      issued.filter(function(e){ return /receipt/i.test(e.docType) && e.method; })
            .forEach(function(e){ m[e.method] = (m[e.method]||0) + (e.payment||0); });
      return m;
    })()
  };
  totals.netLinesMinusPayments = Math.max(0, ps_num_(totals.invoicesLinesSubtotal) - ps_num_(totals.totalPayments));

  return { entries: entries, totals: totals };
}

// ---------- helpers ----------
function ps_headerMap_(hdrs){
  var H = {}; for (var i=0;i<hdrs.length;i++){ var k=String(hdrs[i]||'').trim(); if(k) H[k]=i+1; } return H;
}
function ps_pick_(H, names){ for (var i=0;i<names.length;i++){ if (H[names[i]]) return H[names[i]]; } return 0; }
function ps_val_(row, H, names){ var c=ps_pick_(H,names); return c? row[c-1] : ''; }
function ps_num_(v){
  if (typeof rp_num_ === 'function') return rp_num_(v);
  var s = String(v == null ? '' : v).replace(PS_RE_NUM_SCRUB,'');
  var n = parseFloat(s);
  return isFinite(n) ? n : 0;
}
function ps_soEq_(a,b){
  var A = ps_clean_(a), B = ps_clean_(b);
  if (A === B) return true;
  var na = Number(A.replace(/[^\d.]/g,'')), nb = Number(B.replace(/[^\d.]/g,''));
  return isFinite(na) && isFinite(nb) && Math.abs(na-nb) < 1e-9;
}
function ps_clean_(s){ return String(s==null?'':s).toUpperCase().replace(/[\u200B-\u200D\uFEFF]/g,'').trim(); }

function ps_parseLines_(raw){
  if (!raw) return [];
  if (Array.isArray(raw)) {
    return raw.map(function(x){
      return { desc: String(x.desc||x.description||'').trim(),
               qty:  ps_num_(x.qty||x.quantity||1),
               amt:  ps_num_(x.amt||x.amount||0) };
    }).filter(function(x){ return x.desc || x.amt; });
  }
  if (typeof raw === 'string') {
    var s = raw.trim();
    var c = s.charAt(0);
    if (c === '[' || c === '{') {
      try {
        var parsed = JSON.parse(s);
        if (Array.isArray(parsed)) {
          return parsed.map(function(x){
            return { desc: String(x.desc||x.description||'').trim(),
                     qty:  ps_num_(x.qty||x.quantity||1),
                     amt:  ps_num_(x.amt||x.amount||0) };
          }).filter(function(x){ return x.desc || x.amt; });
        }
      } catch(_){ /* fall through to plain-text line */ }
    }
  }
  return [{ desc: String(raw), qty: 1, amt: 0 }];
}

function ps_parseDate_(v){
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return v;
  var d = new Date(String(v)); return isNaN(d) ? null : d;
}
function ps_guessDocType_(row, cGross){
  var amt = ps_num_(cGross ? row[cGross-1] : 0);
  return (amt > 0) ? 'Receipt' : 'Invoice';
}

// =============== Export PDF ===============
function ps_exportPdf(){
  try {
    var prefill = (typeof rp_init === 'function') ? rp_init() : null;
    if (!prefill) throw new Error('No active row / prefill failed.');

    var anchor = {
      anchorType: prefill.anchorType || (prefill.brand ? 'SO' : 'APPT'),
      soNumber:   String(prefill.soNumber || prefill.SO || prefill['SO#'] || '').trim(),
      rootApptId: String(prefill.rootApptId || prefill.rootApptID || prefill.APPT_ID || '').trim()
    };

    var hist = ps_fetchHistoryForAnchor_(anchor);

    var tz = Session.getScriptTimeZone();
    var stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH-mm');
    var anchorLabel = (anchor.soNumber || anchor.rootApptId || prefill.customerName || 'Unknown');
    var title = 'Payment Summary — ' + anchorLabel + ' — ' + stamp;

    // Build a Google Doc → export to PDF
    var doc = DocumentApp.create(title);
    var body = doc.getBody();
    body.appendParagraph('Payment Summary').setHeading(DocumentApp.ParagraphHeading.HEADING1);

    // Prefill / header info
    var kv = [
      ['Anchor', anchor.anchorType],
      ['Brand', prefill.brand || ''],
      ['Customer', prefill.customerName || prefill.Customer || ''],
      ['SO#', anchor.soNumber || ''],
      ['Order Total', ps_currency_(prefill.orderTotal)],
      ['Paid-To-Date', ps_currency_(prefill.paidToDate)],
      ['Remaining (OT − PTD)', ps_currency_(Math.max(0, (ps_num_(prefill.orderTotal)||0) - (ps_num_(prefill.paidToDate)||0)))]
    ];
    var t0 = body.appendTable(kv.map(function(r){ return [r[0], String(r[1]||'')]; }));
    for (var i=0; i<t0.getNumRows(); i++) t0.getCell(i,0).editAsText().setBold(true);

    body.appendParagraph(' ');
    var tbl = body.appendTable();
    var hdr = tbl.appendTableRow();
    ['Date/Time','Type','Status','Role','Doc #','Lines Subtotal','Payment','Method','Reference','Links','PDF','Doc'].forEach(function(h){
      hdr.appendTableCell(h).editAsText().setBold(true);
    });

    (hist.entries || []).forEach(function(e){
      var row = tbl.appendTableRow();
      row.appendTableCell(e.dateDisplay || '');
      row.appendTableCell(e.docType || '');
      row.appendTableCell(e.docStatus || '');
      row.appendTableCell(e.docRole || '');
      row.appendTableCell(e.docNumber || '');
      row.appendTableCell(ps_currency_(e.linesSubtotal || 0));
      row.appendTableCell(e.payment ? ps_currency_(e.payment) : '');
      row.appendTableCell(e.method || '');
      row.appendTableCell(e.reference || '');

      // PDF cell with link if available
      var cpdf = row.appendTableCell(e.pdfUrl ? 'Open' : '');
      if (e.pdfUrl) {
        var p = cpdf.getChild(0).asParagraph().editAsText();
        p.setText('Open'); p.setLinkUrl(0, 3, e.pdfUrl); // link the word "Open"
      }

      // Doc cell with link if available
      var cdoc = row.appendTableCell(e.docUrl ? 'Open' : '');
      if (e.docUrl) {
        var q = cdoc.getChild(0).asParagraph().editAsText();
        q.setText('Open'); q.setLinkUrl(0, 3, e.docUrl);
      }
    });

    body.appendParagraph(' ');
    var t2 = body.appendTable([
      ['Invoice Lines Subtotal', ps_currency_(hist.totals && hist.totals.invoicesLinesSubtotal || 0)],
      ['Total Payments',        ps_currency_(hist.totals && hist.totals.totalPayments || 0)],
      ['Net (Lines − Payments)',ps_currency_(hist.totals && hist.totals.netLinesMinusPayments || 0)]
    ]);
    for (var j=0; j<t2.getNumRows(); j++) t2.getCell(j,0).editAsText().setBold(true);

    doc.saveAndClose();
    var blob = DriveApp.getFileById(doc.getId()).getAs('application/pdf');

    // Save into Payments Folder when available
    var folderId = ps_parseFolderIdFromUrl_(prefill.paymentsFolderURL || '');
    var folder;
    try { folder = folderId ? DriveApp.getFolderById(folderId) : null; } catch(e){ folder = null; }
    var pdf = (folder || DriveApp.getRootFolder()).createFile(blob).setName(title + '.pdf');

    // Clean up the temp Doc
    try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch (_){}

    return { ok:true, url: pdf.getUrl(), fileId: pdf.getId() };
  } catch (e) {
    return { ok:false, error: e && e.message ? e.message : String(e) };
  }
}

// ---- tiny helpers (add once; skip if you already have them) ----
function ps_parseFolderIdFromUrl_(url){
  if (!url) return '';
  var m = String(url).match(/\/folders\/([a-zA-Z0-9_-]{10,})/);
  return m ? m[1] : '';
}
function ps_currency_(n){
  var v = Number(n); if (!isFinite(v)) v = 0;
  return '$' + v.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
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



