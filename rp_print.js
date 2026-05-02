
function rp_openPrintReceipt(payload) {
  try {
    const p   = payload || {};
    const pmt = p.pmt   || {};

    const dt   = String(p.docType || '').toLowerCase();
    const isSR = dt.includes('sales')   && dt.includes('receipt');
    const isDR = dt.includes('deposit') && dt.includes('receipt');

    if (!isSR && !isDR) {
      throw new Error('[rp_openPrintReceipt] Not a supported receipt type: ' + p.docType);
    }

    const tmplFile = isSR ? 'print_sr_8x10_hpusa' : 'print_dr_8x10_hpusa';

    Logger.log('[rp_openPrintReceipt] docType="%s" taxEnabled=%s → template=%s',
               p.docType, p.taxEnabled, tmplFile);

    function money(n) {
      const v = isFinite(Number(n)) ? Number(n) : 0;
      const parts = Math.abs(v).toFixed(2).split('.');
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
      return (v < 0 ? '-$' : '$') + parts.join('.');
    }

    const taxEnabled = (p.taxEnabled !== false);
    const brand      = String(p.brand || '').trim();
    let   taxRate    = 0;
    if (taxEnabled && brand) {
      try { taxRate = rp_getTaxRate_(brand); } catch (_) {}
    }

    const lines       = (p.lines || []).slice(0, 6);
    const rawSubtotal = lines.reduce(
      (s, ln) => s + (Number(ln.qty || 0) * Number(ln.amt || 0)), 0
    );

    const referralDiscount   = (p.referralEnabled && p.referralDiscount)
                              ? Math.max(0, Number(p.referralDiscount || 0)) : 0;
    const subtotalNum        = Math.max(0, rawSubtotal - referralDiscount);
    const taxNum             = Math.round(subtotalNum * taxRate * 100) / 100;
    const totalNum           = Math.round((subtotalNum + taxNum) * 100) / 100;

    const snap          = p.snapshots  || {};
    const amountPaid    = Number(pmt.amount || 0);
    const orderTotal    = totalNum > 0 ? totalNum : Number(snap.orderTotal || 0);
    const paidToDate    = Number(snap.paidToDate || 0);
    const balanceDueNum = Math.max(0, orderTotal - paidToDate - amountPaid);

    let docDate = '';
    try {
      const raw = pmt.dateTime || '';
      const d   = raw ? new Date(raw) : new Date();
      docDate   = Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    } catch (_) {}

    let previousPaymentsBlock = '';
    try {
      if (Array.isArray(p.prevPayments) && p.prevPayments.length > 0) {
        previousPaymentsBlock = p.prevPayments.map(function(it) {
          const dt  = it.date   || '';
          const amt = money(Number(it.amount || 0));
          const raw = String(it.method || '').trim();
          const m   = /^zelle/i.test(raw) ? 'Zelle' : raw;  // ★ normalize Zelle
          return '✧ ' + dt + ' — ' + amt + (m ? ' ' + m : '');
        }).join('\n');
      }
    } catch (_) {}

    function truncateDesc(text) {
      const charsPerLine = 40;
      const maxLines     = 2;
      const maxChars     = charsPerLine * maxLines;
      if (!text || text.length <= maxChars) return text || '';
      return text.slice(0, maxChars - 1).trim() + '…';
    }

    function lineDesc(i) {
      const ln = lines[i];
      if (!ln || !ln.desc) return '';
      const raw = String(ln.desc).replace(/^\s*✧\s*/u, '').trim();
      return truncateDesc(raw);
    }
    function lineQty(i) {
      const ln = lines[i];
      return (ln && ln.qty != null && String(ln.qty).trim() !== '') ? String(ln.qty) : '';
    }
    function lineAmt(i) {
      const ln = lines[i];
      if (!ln || !ln.qty || !ln.amt) return '';
      return money(Number(ln.qty) * Number(ln.amt));
    }

    // ★ UPDATE: bỏ checkbox, chỉ dùng text PAYMENT_METHOD
    const methodDisplay = (() => {
      const m = String(pmt.method || '').trim();
      if (/^zelle/i.test(m))               return 'Zelle';
      if (/^credit card$|^card$/i.test(m)) return 'Credit Card';
      return m;
    })();

    const tmpl = HtmlService.createTemplateFromFile(tmplFile);

    tmpl.customerName          = p.customerName || '';
    tmpl.docDate               = docDate;
    tmpl.phone                 = p.phone        || '';
    tmpl.email                 = p.email        || '';
    tmpl.paymentNotes          = pmt.notes      || '';
    tmpl.soNumber              = p.soNumber     || '';

    for (let i = 0; i < 6; i++) {
      const n = i + 1;
      tmpl['line' + n + '_desc'] = lineDesc(i);
      tmpl['line' + n + '_qty']  = lineQty(i);
      tmpl['line' + n + '_amt']  = lineAmt(i);
    }

    tmpl.previousPaymentsBlock = previousPaymentsBlock;
    tmpl.paymentsLabel         = previousPaymentsBlock ? 'PAYMENTS' : '';
    tmpl.linesSubtotal         = money(subtotalNum);
    tmpl.referralDiscount      = referralDiscount > 0 ? ('− ' + money(referralDiscount)) : '';
    tmpl.discountedSubtotal    = referralDiscount > 0 ? money(subtotalNum) : '';
    tmpl.taxAmount             = money(taxNum);
    tmpl.invoiceTotal          = money(totalNum);
    tmpl.paymentAmount         = money(amountPaid);
    tmpl.balanceDue            = money(balanceDueNum);
    tmpl.paymentMethod         = methodDisplay;  // ★ thay thế chk_*

    return tmpl.evaluate().getContent();

  } catch (e) {
    Logger.log('[rp_openPrintReceipt] ERROR: ' + (e && e.stack ? e.stack : e));
    throw e;
  }
}