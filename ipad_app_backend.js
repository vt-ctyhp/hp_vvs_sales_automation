/*** ipad_app.gs — v2.2  (Project #20 · sync P21 + P22)
 * ══════════════════════════════════════════════════════════════════════
 * iPad Web App — HPUSA & VVS dual-brand support
 *
 * CHANGES v2.1 → v2.2:
 *   • Project #21 — ipad_loadRecord() returns hasSalesInvoice
 *     (pre-checks Payments Ledger for a non-VOID Sales Invoice)
 *   • Project #21 — ipad_checkHasSalesInvoice() public wrapper added
 *     (called from HTML after prefill, NOT inside loadRecord to avoid
 *      slowing down triggers — same pattern as dlg_record_payment_v1)
 *   • Project #22 — ipad_loadRecord() returns referralUsed / referralName
 *     / referralDiscount (reads from rp_isReferralAlreadyUsed_)
 *   • ipad_submit() unchanged — rp_submit() already enforces P21 guard
 *     and P22 discount logic (added to Payments_v1.gs)
 *
 * CHANGES v2.0 → v2.1:
 *   • ipad_loadRecord() returns taxEnabled (read from Payments Ledger)
 *   • ipad_getLedgerSheet_() helper added
 *
 * CHANGES v1 → v2.0:
 *   • ipad_searchCustomers() filters by brand
 *   • ipad_loadRecord() returns brand
 *   • ipad_submitDiamondViewing() accepts brand param
 *   • Print: HPUSA → rp_openPrintReceipt() | VVS → no 8×10
 * ══════════════════════════════════════════════════════════════════════
 */

// ── Entry point ────────────────────────────────────────────────────────
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('ipad_app')
    .setTitle('HP & VVS — Receipt Generator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ═══════════════════════════════════════════════════════════════════════
// BRAND CONFIG
// ═══════════════════════════════════════════════════════════════════════

function ipad_getBrandConfig() {
  return {
    HPUSA: {
      key:       'HPUSA',
      label:     'Hung Phat USA',
      short:     'HP',
      color:     '#E0006A',
      textColor: '#fff',
      hasPrint:  true,
    },
    VVS: {
      key:       'VVS',
      label:     'VVS Jewelry',
      short:     'VVS',
      color:     '#1a1a2e',
      textColor: '#d4af37',
      hasPrint:  false,
    },
  };
}

// ═══════════════════════════════════════════════════════════════════════
// SEARCH & LOOKUP
// ═══════════════════════════════════════════════════════════════════════

/**
 * Search customers — optionally filtered by brand.
 * @param {string} query  — name / SO# / Appt ID
 * @param {string} brand  — 'HPUSA' | 'VVS' | '' (all)
 */
function ipad_searchCustomers(query, brand) {
  if (!query || String(query).trim().length < 2) return [];
  const q       = String(query).trim().toLowerCase();
  const bFilter = brand ? String(brand).trim().toUpperCase() : '';

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RP_MASTER_SHEET);
  if (!sh) return [];

  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return [];

  const header = sh.getRange(1, 1, 1, lc).getDisplayValues()[0];
  const map    = rp_headerMap([header]);

  const custIdx  = rp_pick0(map, 'Customer Name', 'Customer', 'Client Name', 'Client');
  const apptIdx  = rp_pick0(map, 'APPT_ID', 'RootApptID', 'Root Appt ID');
  const soIdx    = rp_pick0(map, 'SO#', 'SO', 'SO Number', 'Sales Order');
  const phoneIdx = rp_pick0(map, 'Phone', 'Phone Number', 'Tel', 'Mobile');
  const emailIdx = rp_pick0(map, 'Email', 'Email Address', 'E-mail');
  const brandIdx = rp_pick0(map, 'Brand');
  const otIdx    = rp_pick0(map, 'Order Total');
  const ptdIdx   = rp_pick0(map, 'Paid-to-Date', 'Paid-To-Date', 'Paid to Date');
  const pfIdx    = rp_pick0(map, 'PaymentsFolderURL');

  const vals    = sh.getRange(2, 1, lr - 1, lc).getDisplayValues();
  const results = [];

  for (let i = 0; i < vals.length && results.length < 25; i++) {
    const row      = vals[i];
    const name     = custIdx  >= 0 ? String(row[custIdx]  || '').trim() : '';
    const so       = soIdx    >= 0 ? String(row[soIdx]    || '').trim() : '';
    const appt     = apptIdx  >= 0 ? String(row[apptIdx]  || '').trim() : '';
    const rowBrand = brandIdx >= 0 ? String(row[brandIdx] || '').trim().toUpperCase() : '';

    if (!name && !so && !appt) continue;

    // Brand filter
    if (bFilter) {
      if (bFilter === 'HPUSA' && !rowBrand.includes('HPUSA')) continue;
      if (bFilter === 'VVS'   && !rowBrand.includes('VVS'))   continue;
    }

    // // Text search
    // if (
    //   name.toLowerCase().includes(q) ||
    //   so.toLowerCase().includes(q)   ||
    //   appt.toLowerCase().includes(q)
    // ) {
    // Text search — name / SO# / Appt ID / Phone
    const phoneQ = phoneIdx >= 0 ? String(row[phoneIdx] || '').trim().toLowerCase().replace(/\D/g,'') : '';
    const qDigits = q.replace(/\D/g,'');
    const phoneMatch = phoneQ && qDigits.length >= 4 && phoneQ.includes(qDigits);

    if (
      name.toLowerCase().includes(q) ||
      so.toLowerCase().includes(q)   ||
      appt.toLowerCase().includes(q) ||
      (phoneIdx >= 0 && String(row[phoneIdx] || '').toLowerCase().includes(q)) ||
      phoneMatch
    ) {
      results.push({
        rowIndex:          i + 2,
        customerName:      name,
        soNumber:          so,
        rootApptId:        appt,
        brand:             brandIdx >= 0 ? String(row[brandIdx] || '').trim() : '',
        phone:             phoneIdx >= 0 ? String(row[phoneIdx] || '').trim() : '',
        email:             emailIdx >= 0 ? String(row[emailIdx] || '').trim() : '',
        orderTotal:        otIdx  >= 0   ? String(row[otIdx]    || '') : '',
        paidToDate:        ptdIdx >= 0   ? String(row[ptdIdx]   || '') : '',
        paymentsFolderURL: pfIdx  >= 0   ? String(row[pfIdx]    || '') : '',
        anchorType:        so ? 'SO' : 'APPT',
      });
    }
  }

  // Chỉ dedup rows clone thật sự: cùng rootApptId + cùng orderTotal
  // Rows cùng khách nhưng khác visit (khác payment) → GIỮ NGUYÊN cả 2
  const dedupKey = new Set();
  return results.filter(r => {
    if (!r.rootApptId) return true;
    const k = r.rootApptId + '|' + (r.orderTotal || '0');
    if (dedupKey.has(k)) return false;
    dedupKey.add(k);
    return true;
  });
}

/**
 * Load full record by Master row index.
 * Brand-aware: fetches tax rate + taxEnabled + hasSalesInvoice + referral status.
 *
 * P21: hasSalesInvoice — pre-checked so HTML can cache instantly.
 * P22: referralUsed / referralName / referralDiscount — for lock UI.
 */
// function ipad_loadRecord(rowIndex) {
//   try {
//     rowIndex = Number(rowIndex);
//     if (rowIndex < 2) throw new Error('Invalid row index');

//     const m     = rp_getMasterRowByIndex_(rowIndex);
//     const brand = (m.map['Brand'] != null)
//                   ? String(m.rowVals[m.map['Brand']] || '').trim()
//                   : '';
//     const hasSO      = !!(m.soNumber && String(m.soNumber).trim());
//     const anchorType = hasSO ? 'SO' : 'APPT';

//     // Tax rate for THIS brand
//     let taxRate = 0;
//     try { taxRate = brand ? rp_getTaxRate_(brand) : 0; } catch(_) {}

//     // ── Đọc giá trị raw từ Master sheet ─────────────────────────────
//     const ptdIdx = rp_pick0(m.map, 'Paid-to-Date','Paid-To-Date','Paid to Date','Paid-to-date');
//     const otIdx  = m.map['Order Total'] != null ? m.map['Order Total'] : -1;
//     const rbIdx  = rp_pick0(m.map, 'Remaining Balance','Balance');

//     const ss2 = SpreadsheetApp.getActive().getSheetByName(RP_MASTER_SHEET);
//     const raw = ss2.getRange(rowIndex, 1, 1, ss2.getLastColumn()).getValues()[0];

//     const paidToDate    = ptdIdx >= 0 ? rp_num_(raw[ptdIdx]) : 0;
//     const orderTotalRaw = otIdx  >= 0 ? rp_num_(raw[otIdx])  : 0;

//     // ── FIX #1 + #4: orderTotal & balance ───────────────────────────
//     // Sheet đã lưu invoiceTotal (đã có tax). KHÔNG nhân thêm taxRate.
//     // Ưu tiên đọc Remaining Balance trực tiếp từ cột RB của sheet.
//     let orderTotal = orderTotalRaw;
//     let balance;

//     if (orderTotal > 0 && rbIdx >= 0) {
//       // Đọc thẳng từ sheet — chính xác nhất, tránh tính lại
//       balance = Math.max(0, rp_num_(raw[rbIdx]));
//     } else {
//       balance = Math.max(0, orderTotal - paidToDate);
//     }

//     // FIX #4: Fallback InvoiceTotal từ ledger khi OT blank (mirror rp_init v8.9)
//     if (!orderTotal) {
//       try {
//         const lastIT = rp_getLastInvoiceTotalFromLedger_({
//           hasSO, soNumber: m.soNumber, rootApptId: m.rootApptId
//         });
//         if (lastIT > 0) {
//           orderTotal = lastIT;
//           balance    = Math.max(0, lastIT - paidToDate);
//           Logger.log('[ipad_loadRecord] InvoiceTotal fallback=%s', lastIT);
//         }
//       } catch(_) {}
//     }

//     // ── FIX #2 + #3 + #5: taxEnabled — mirror rp_init() v8.8 ────────
//     // Dùng rp_getLedgerTarget() thay vì ipad_getLedgerSheet_() (undefined)
//     // Scan từ LAST row ngược lên, dừng ở row khớp đầu tiên (giống rp_init)
//     let taxEnabled      = true;
//     let foundTaxRecord  = false;

//     try {
//       const { sh: lSh } = rp_getLedgerTarget(); // FIX #5: đúng hàm
//       const lLr = lSh.getLastRow();
//       const lLc = lSh.getLastColumn();

//       if (lLr >= 2) {
//         const lHdr = lSh.getRange(1, 1, 1, lLc).getValues()[0]
//                         .map(v => String(v).trim());
//         const LH = {}; lHdr.forEach((h, i) => LH[h] = i);

//         const cTax    = LH['TaxEnabled'];
//         const cAppt   = LH['RootApptID'];
//         const cSO     = LH['SO#'];

//         if (cTax != null) {
//           const start = Math.max(2, lLr - 500);
//           const lVals = lSh.getRange(start, 1, lLr - start + 1, lLc).getValues();

//           // FIX #2: Scan từ CUỐI lên, dừng ở match đầu tiên — giống rp_init()
//           for (let i = lVals.length - 1; i >= 0; i--) {
//             const r = lVals[i];
//             const match = hasSO
//               ? rp_soEq(r[cSO], m.soNumber)
//               : String(r[cAppt] || '').trim() === String(m.rootApptId || '').trim();

//             if (match) {
//               const rawTax = r[cTax];
//               // Chỉ dùng nếu ô KHÔNG rỗng — giống rp_init() v8.8
//               if (rawTax !== '' && rawTax !== null && rawTax !== undefined) {
//                 taxEnabled     = !(rawTax === false || String(rawTax).toLowerCase() === 'false');
//                 foundTaxRecord = true;
//               }
//               break; // Dừng ngay sau row match đầu tiên
//             }
//           }
//         }
//       }
//     } catch (te) {
//       Logger.log('[ipad_loadRecord taxEnabled] ERROR: ' + te.message);
//     }

//     // FIX #3: Guard legacy data — mirror rp_init() v8.8
//     // Nếu không tìm thấy record TaxEnabled VÀ đã có thanh toán
//     // → OT cũ đã bao gồm tax → đặt taxEnabled = false tránh double-tax
//     if (!foundTaxRecord && paidToDate > 0) {
//       taxEnabled = false;
//     }

//     // ── Previous payments ────────────────────────────────────────────
//     let prevPayments = [];
//     try {
//       const prev = rp_prevPaymentsForAnchor_({
//         anchorType, rootApptId: m.rootApptId,
//         soNumber: m.soNumber, limit: 10,
//       });
//       if (prev && prev.items) {
//         prevPayments = prev.items.map(it => ({
//           date: it.date || '', amount: Number(it.amount || 0),
//           method: it.method || '', docNumber: it.docNumber || '',
//         }));
//       }
//     } catch(_) {}

//     // ── Extra columns ────────────────────────────────────────────────
//     const phoneIdx = rp_pick0(m.map, 'Phone','Phone Number','Tel','Mobile');
//     const emailIdx = rp_pick0(m.map, 'Email','Email Address','E-mail');
//     const pfIdx    = rp_pick0(m.map, 'PaymentsFolderURL');

//     // ── Saved lines ──────────────────────────────────────────────────
//     let savedLines = [];
//     try {
//       const saved = rp_readSavedLinesFromMaster_(m) ||
//                     rp_findLastSavedLinesForAnchor_({
//                       anchorType, rootApptId: m.rootApptId, soNumber: m.soNumber
//                     });
//       if (saved && saved.lines) savedLines = saved.lines;
//     } catch(_) {}

//     // ── Project #21: hasSalesInvoice ────────────────────────────────
//     let hasSalesInvoice = false;
//     try {
//       const invCheck = rp_checkInvoiceBeforeReceipt_({
//         anchorType,
//         rootApptId: m.rootApptId || '',
//         soNumber:   m.soNumber   || '',
//         docType:    'Sales Receipt',
//       });
//       hasSalesInvoice = invCheck.ok;
//     } catch(_) { hasSalesInvoice = false; }

//     // ── Project #22: referral already used? ─────────────────────────
//     let referralUsed = false, referralNameVal = '', referralDiscountVal = 0;
//     try {
//       const refCheck = rp_isReferralAlreadyUsed_({
//         anchorType,
//         rootApptId: m.rootApptId || '',
//         soNumber:   m.soNumber   || '',
//       });
//       referralUsed        = refCheck.used;
//       referralNameVal     = refCheck.name     || '';
//       referralDiscountVal = refCheck.discount || 0;
//     } catch(_) {}

//     Logger.log('[ipad_loadRecord] row=%s OT=%s PTD=%s BAL=%s taxEnabled=%s foundRecord=%s',
//       rowIndex, orderTotal, paidToDate, balance, taxEnabled, foundTaxRecord);

//     return {
//       ok: true,
//       rowIndex, anchorType, brand,
//       hasPrint: brand.toUpperCase().includes('HPUSA'),
//       taxEnabled,
//       // P21
//       hasSalesInvoice,
//       // P22
//       referralUsed,
//       referralName:     referralNameVal,
//       referralDiscount: referralDiscountVal,
//       // core
//       customerName:      m.customerName,
//       soNumber:          m.soNumber || '',
//       rootApptId:        m.rootApptId || '',
//       trackerUrl:        m.trackerUrl || '',
//       phone:             phoneIdx >= 0 ? String(m.rowVals[phoneIdx] || '').trim() : '',
//       email:             emailIdx >= 0 ? String(m.rowVals[emailIdx] || '').trim() : '',
//       orderTotal:        String(orderTotal || ''),
//       paidToDate:        String(paidToDate || ''),
//       balance:           String(balance),
//       taxRate,
//       paymentsFolderURL: pfIdx >= 0 ? String(m.rowVals[pfIdx] || '').trim() : '',
//       prevPayments,
//       savedLines,
//     };

//   } catch (e) {
//     Logger.log('[ipad_loadRecord] ERROR: ' + e.message);
//     return { ok: false, error: e.message };
//   }
// }

// ═══════════════════════════════════════════════════════════════════════
// SUBMIT
// ═══════════════════════════════════════════════════════════════════════

/**
 * Full submit: ledger → doc/PDF → 8×10 print HTML (HPUSA only).
 * P21 guard and P22 discount logic are enforced inside rp_submit().
 */
function ipad_submit(payload) {
  try {
    const submitRes = rp_submit(payload);
    if (!submitRes || !submitRes.ok) {
      return { ok: false, error: 'rp_submit failed', detail: JSON.stringify(submitRes) };
    }

    const docRes = rp_makeDocForPayment(submitRes.row, payload);
    if (!docRes || !docRes.ok) {
      return {
        ok:     false,
        error:  docRes && docRes.hint   ? docRes.hint   : 'Doc generation failed',
        reason: docRes && docRes.reason ? docRes.reason : '',
      };
    }

    // ── Lấy prevPayments TRƯỚC — cần cho cả printHtml và return ──  
    let prevPayments = [];
    try {
      const prev = rp_prevPaymentsForAnchor_({
        anchorType: payload.anchorType || 'APPT',
        rootApptId: payload.rootApptId || '',
        soNumber:   payload.soNumber   || '',
        limit: 10,
      });
      if (prev && prev.items) {
        prevPayments = prev.items.map(it => ({
          date:      it.date      || '',
          amount:    Number(it.amount || 0),
          method:    it.method    || '',
          docNumber: it.docNumber || '',
        }));
      }
    } catch(e) {
      Logger.log('[ipad_submit] prevPayments error: ' + e.message);
    }

    // ── 8×10 print — HPUSA receipts only (SAU khi có prevPayments) ──
    let printHtml = '';
    const brandUp   = String(payload.brand || '').toUpperCase();
    const isHPUSA   = brandUp.includes('HPUSA');
    const isReceipt = /Receipt/i.test(payload.docType || '');
    if (isHPUSA && isReceipt) {
      try {
        printHtml = rp_openPrintReceipt(
          Object.assign({}, payload, {
            docNumber:    docRes.docNumber || '',
            prevPayments: prevPayments,      // ← truyền vào đây
          })
        );
      } catch (pe) {
        Logger.log('[ipad_submit] print warning: ' + pe.message);
      }
    }

    // ── Tính balance mới nhất ──
    const pmtAmount = payload.pmt ? Number(payload.pmt.amount || 0) : 0;
    const pmtMethod = payload.pmt ? String(payload.pmt.method || '') : '';
    const snapPtd   = payload.snapshots ? Number(payload.snapshots.paidToDate || 0) : 0;
    const snapBal   = payload.snapshots ? Number(payload.snapshots.balance    || 0) : 0;

    const isReceiptDoc = /Receipt/i.test(payload.docType || '');
    const newPaidToDate = isReceiptDoc ? (snapPtd + pmtAmount) : snapPtd;
    const newBalance    = isReceiptDoc ? Math.max(0, snapBal - pmtAmount) : snapBal;

    return {
      ok:                true,
      brand:             String(payload.brand || ''),
      docNumber:         docRes.docNumber         || '',
      docUrl:            docRes.docUrl            || '',
      pdfUrl:            docRes.pdfUrl            || '',
      paymentsFolderURL: (payload.paymentsFolderURL && payload.paymentsFolderURL.includes('drive.google.com'))
                              ? payload.paymentsFolderURL
                              : (docRes.paymentsFolderURL || ''),
            intakeDocURL:   payload.intakeDocURL   || docRes.intakeDocURL   || '',
            checklistURL:   payload.checklistURL   || docRes.checklistURL   || '',
            quotationURL:   payload.quotationURL   || docRes.quotationURL   || '',
            arShortcutURL:     docRes.arShortcutURL     || '',
      printHtml,
      // ── THÊM MỚI ──────────────────────────────────────
      paymentSummary: {
        docType:     payload.docType || '',
        amount:      pmtAmount,
        method:      pmtMethod,
        paidToDate:  newPaidToDate,
        balance:     newBalance,
      },
      prevPayments: prevPayments,
    };

  } catch (e) {
    Logger.log('[ipad_submit] ERROR: ' + (e && e.stack ? e.stack : e));
    return { ok: false, error: e.message || String(e) };
  }
}

/**
 * Quick $25 Diamond Viewing Deposit — works for both brands.
 * taxEnabled = false (fixed-price viewing deposit).
 */
function ipad_submitDiamondViewing(params) {
  const p = params || {};
  const isInvoice = /Invoice/i.test(p.dvDocType || 'Deposit Invoice');

  const payload = {
    anchorType:    p.anchorType    || 'APPT',
    brand:         p.brand         || 'HPUSA',
    rootApptId:    p.rootApptId    || '',
    soNumber:      p.soNumber      || '',
    customerName:  p.customerName  || '',
    phone:         p.phone         || '',
    email:         p.email         || '',
    docType: p.dvDocType || 'Deposit Invoice',
    taxEnabled:    false,
    referralEnabled:  false,
    referralName:     '',
    referralDiscount: 0,
    lines: [{ desc: 'Diamond Viewing Deposit', qty: 1, amt: 25 }],
    pmt: {
      amount:    25,
      dateTime:  (p.pmt && p.pmt.dateTime)  || '',
      method:    isInvoice ? '' : ((p.pmt && p.pmt.method) || ''),
      reference: (p.pmt && p.pmt.reference) || '',
      notes:     (p.pmt && p.pmt.notes)     || 'Diamond Viewing Deposit — $25',
      allocatedToSO: 0,
    },
    flags:     { setOrderTotal: false },
    snapshots: { orderTotal: 25, paidToDate: 0, balance: 25 },
    docStatus:  'ISSUED',
    docRole:    'DEPOSIT',
    supersedes: '', appliesTo: '',
    masterRowIndex:    p.rowIndex          || 0,
    trackerUrl:        p.trackerUrl        || '',
    paymentsFolderURL: p.paymentsFolderURL || '',
  };

  return ipad_submit(payload);
}

// ── Helpers exposed to client ──────────────────────────────────────────

function ipad_getTaxRate(brand) {
  try { return rp_getTaxRate_(brand || ''); } catch(_) { return 0; }
}

function ipad_listDocNumbers(params) {
  try { return rp_listDocNumbersForAnchor(params || {}); } catch(_) { return []; }
}

/**
 * Project #21 — public wrapper called from iPad HTML after prefill.
 * Kept separate from ipad_loadRecord() so it does NOT run during
 * background triggers (same isolation pattern as dlg_record_payment_v1).
 *
 * @param {{anchorType,rootApptId,soNumber}} payload
 * @returns {{ok:boolean, hasSalesInvoice:boolean}}
 */
function ipad_checkHasSalesInvoice(payload) {
  try {
    const result = rp_checkInvoiceBeforeReceipt_({
      anchorType: (payload && payload.anchorType) || 'APPT',
      rootApptId: (payload && payload.rootApptId) || '',
      soNumber:   (payload && payload.soNumber)   || '',
      docType:    'Sales Receipt',
    });
    return { ok: true, hasSalesInvoice: result.ok };
  } catch(e) {
    Logger.log('[ipad_checkHasSalesInvoice] ERROR: ' + e.message);
    return { ok: true, hasSalesInvoice: false };
  }
}

// function ipad_submitIntake(payload) {
//   try {
//     const p = payload || {};
//     const tz = Session.getScriptTimeZone();

//     // Format date MM/DD/YYYY
//     let visitDateStr = '';
//     if (p.date) {
//       const d = new Date(p.date + 'T12:00:00');
//       visitDateStr = Utilities.formatDate(d, tz, 'MM/dd/yyyy');
//     }

//     // Format time "2:30 PM"
//     let visitTimeStr = '';
//     if (p.time) {
//       const tp = p.time.split(':');
//       let h = parseInt(tp[0], 10);
//       const m = tp[1] || '00';
//       const ap = h >= 12 ? 'PM' : 'AM';
//       h = h % 12 || 12;
//       visitTimeStr = h + ':' + m + ' ' + ap;
//     }

//     const diamonds = Array.isArray(p.diamond) ? p.diamond : (p.diamond ? [p.diamond] : []);
//     const budgets  = Array.isArray(p.budget)  ? p.budget  : (p.budget  ? [p.budget]  : []);
//     const sources  = Array.isArray(p.source)  ? p.source  : (p.source  ? [p.source]  : []);
//     const now = new Date();

//     const namedValues = {
//       'Timestamp':                 [Utilities.formatDate(now, tz, 'M/d/yyyy H:mm:ss')],
//       'Company':                   [p.company   || ''],
//       'Customer Name':             [p.name      || ''],
//       'Phone':                     [p.phone     || ''],
//       'Email':                     [p.email     || ''],
//       'Visit Type':                [p.visitType || 'Walk-In'],
//       'Visit Date':                [visitDateStr],
//       'Visit Time':                [visitTimeStr],
//       'Location':                  [p.location  || 'In Store'],
//       'Diamond Type':              [diamonds.join(', ')],
//       'Budget Range':              [budgets.join(', ')],
//       'Source':                    [sources.join(', ')],
//       'Style Notes':               [p.notes || ''],
//       'Admin: Calendly Event UID': [p.uid   || ''],
//     };

//     // ── Ghi vào 02_Form_Inbox ──────────────────────────────────────
//     const ss = SpreadsheetApp.getActive();
//     const sh = ss.getSheetByName('02_Form_Inbox');
//     if (!sh) throw new Error('Sheet "02_Form_Inbox" not found');

//     const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
//     const rowData = headers.map(h => {
//       const val = namedValues[h];
//       return val ? val[0] : '';
//     });

//     sh.appendRow(rowData);
//     const newRow = sh.getLastRow();
//     Logger.log('[ipad_submitIntake] Appended to 02_Form_Inbox row=' + newRow);

//     // ── Gọi onFormSubmit ───────────────────────────────────────────
//     onFormSubmit({
//       namedValues: namedValues,
//       range: sh.getRange(newRow, 1, 1, headers.length),
//       values: rowData,
//     });

//     Logger.log('[ipad_submitIntake] onFormSubmit OK | name=' + (p.name||''));

//     // ── Tìm masterRowIndex vừa được tạo ──
//     let masterRowIndex = 0;
//     try {
//       const ss2 = SpreadsheetApp.getActive();
//       const mSh = ss2.getSheetByName(RP_MASTER_SHEET);
//       if (mSh) {
//         const lr2 = mSh.getLastRow(), lc2 = mSh.getLastColumn();
//         if (lr2 >= 2) {
//           const hdr2  = mSh.getRange(1,1,1,lc2).getDisplayValues()[0];
//           const map2  = rp_headerMap([hdr2]);
//           const nIdx  = rp_pick0(map2,'Customer Name','Customer','Client Name');
//           const vals2 = mSh.getRange(2,1,lr2-1,lc2).getDisplayValues();
//           const target = String(p.name||'').trim().toLowerCase();
//           for (let i = vals2.length - 1; i >= 0; i--) {
//             const rowName = nIdx >= 0 ? String(vals2[i][nIdx]||'').trim().toLowerCase() : '';
//             if (rowName && rowName === target) {
//               masterRowIndex = i + 2;
//               break;
//             }
//           }
//         }
//       }
//     } catch(me) {
//       Logger.log('[ipad_submitIntake] masterRowIndex error: ' + me.message);
//     }

//     return {
//       ok:             true,
//       row:            newRow,
//       masterRowIndex: masterRowIndex,
//       customerName:   p.name    || '',
//       brand:          p.company || 'HPUSA',
//     };

//   } catch(e) {
//     Logger.log('[ipad_submitIntake] ERROR: ' + e.message + '\n' + (e.stack||''));
//     return { ok: false, error: e.message };
//   }
// }