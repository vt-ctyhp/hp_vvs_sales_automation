/*** ipad_app_root_linking.gs — v1.0
 * ══════════════════════════════════════════════════════════════════════
 * Root Appointment ID — Contact Matching & Auto-Linking
 *
 * ANSWERS TO 4 QUESTIONS:
 *
 *   Q1. Match priority:
 *       Email (exact, case-insensitive) > Phone (digits-only, ≥10 digits)
 *       Rationale: email is more globally unique than phone.
 *       BOTH matching = definitive. ONE matching = accepted with logging.
 *       Edge case: if email matches person A but phone matches person B
 *       → email wins. Staff gets a console warning to verify manually.
 *
 *   Q2. Auto-linking:
 *       ipad_submitIntake() calls ipad_findExistingRootByContact_()
 *       BEFORE onFormSubmit. If a match is found, the existing
 *       rootApptId is injected into namedValues['Admin: Calendly Event UID']
 *       (the field onFormSubmit already reads to skip creating a new root).
 *
 *   Q3. Dedup existing data:
 *       ipad_findDuplicateContacts() — admin utility, returns conflict list.
 *       ipad_relinkRecord() — merges one row into another root ID.
 *       Call ipad_findDuplicateContacts() manually from Apps Script editor
 *       to audit, then call ipad_relinkRecord() to fix specific rows.
 *
 *   Q4. Load behavior when selecting customer:
 *       ipad_loadRecord() now aggregates financial data (OT, PTD, balance,
 *       prevPayments) across ALL master rows sharing the same rootApptId.
 *       The primary row (the one you clicked) is still the canonical row
 *       for contact info, brand, trackerUrl, etc.
 *
 * ══════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════
// Q1 — CONTACT MATCHING HELPERS
// ═══════════════════════════════════════════════════════════════════════

/** Strip everything except digits */
function ipad_normalizePhone_(phone) {
  return String(phone || '').replace(/\D/g, '');
}

/** Lowercase + trim */
function ipad_normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

/**
 * Find an existing master row by email (priority 1) or phone (priority 2).
 *
 * Returns:
 *   { rootApptId, rowIndex, customerName, brand, matchedBy }
 *   OR null if no match found.
 *
 * matchedBy: 'email' | 'phone' | 'both'
 *
 * @param {string} email
 * @param {string} phone
 * @param {string} brand  — 'HPUSA' | 'VVS' | '' for any
 */
function ipad_findExistingRootByContact_(email, phone, brand) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RP_MASTER_SHEET);
  if (!sh) return null;

  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return null;

  const header = sh.getRange(1, 1, 1, lc).getDisplayValues()[0];
  const map    = rp_headerMap([header]);

  const emailIdx = rp_pick0(map, 'Email', 'Email Address', 'E-mail');
  const phoneIdx = rp_pick0(map, 'Phone', 'Phone Number', 'Tel', 'Mobile');
  const apptIdx  = rp_pick0(map, 'APPT_ID', 'RootApptID', 'Root Appt ID');
  const nameIdx  = rp_pick0(map, 'Customer Name', 'Customer', 'Client Name');
  const brandIdx = rp_pick0(map, 'Brand');

  // Can't link without a root column
  if (apptIdx < 0) {
    Logger.log('[ipad_findExistingRootByContact_] No RootApptID column — skipping');
    return null;
  }

  const normEmail = ipad_normalizeEmail_(email);
  const normPhone = ipad_normalizePhone_(phone);
  const hasEmail  = normEmail.length > 3 && normEmail.includes('@');
  const hasPhone  = normPhone.length >= 10;

  if (!hasEmail && !hasPhone) return null;

  const bFilter = brand ? String(brand).trim().toUpperCase() : '';
  const vals    = sh.getRange(2, 1, lr - 1, lc).getDisplayValues();

  let emailMatch = null;
  let phoneMatch = null;

  for (let i = 0; i < vals.length; i++) {
    const row     = vals[i];
    const rootId  = String(row[apptIdx] || '').trim();
    if (!rootId) continue;                                   // Skip rows with no root

    const rowBrand = brandIdx >= 0 ? String(row[brandIdx] || '').trim().toUpperCase() : '';
    if (bFilter && !rowBrand.includes(bFilter)) continue;   // Brand filter

    const rowEmail = emailIdx >= 0 ? ipad_normalizeEmail_(row[emailIdx])   : '';
    const rowPhone = phoneIdx >= 0 ? ipad_normalizePhone_(row[phoneIdx])   : '';
    const name     = nameIdx  >= 0 ? String(row[nameIdx] || '').trim()     : '';
    const brandVal = brandIdx >= 0 ? String(row[brandIdx] || '').trim()    : '';

    const candidate = { rootApptId: rootId, rowIndex: i + 2, customerName: name, brand: brandVal };

    if (hasEmail && rowEmail && rowEmail === normEmail) {
      emailMatch = emailMatch || Object.assign({}, candidate, { matchedBy: 'email' });
    }
    if (hasPhone && rowPhone && rowPhone === normPhone) {
      phoneMatch = phoneMatch || Object.assign({}, candidate, { matchedBy: 'phone' });
    }
  }

  // Conflict guard: email→person A but phone→person B → trust email, warn
  if (emailMatch && phoneMatch && emailMatch.rootApptId !== phoneMatch.rootApptId) {
    Logger.log('[ipad_findExistingRootByContact_] CONFLICT: email→%s phone→%s — trusting email',
      emailMatch.rootApptId, phoneMatch.rootApptId);
    return Object.assign({}, emailMatch, { matchedBy: 'email', conflictNote: 'phone matched different root' });
  }

  // Both match same root → highest confidence
  if (emailMatch && phoneMatch && emailMatch.rootApptId === phoneMatch.rootApptId) {
    return Object.assign({}, emailMatch, { matchedBy: 'both' });
  }

  return emailMatch || phoneMatch || null;
}


// ═══════════════════════════════════════════════════════════════════════
// Q2 — AUTO-LINKING: Updated ipad_submitIntake()
// ═══════════════════════════════════════════════════════════════════════

/**
 * Submit intake form — now with auto-linking.
 *
 * CHANGE vs v1: Before calling onFormSubmit, checks if email or phone
 * already exists in the master sheet. If found, injects the existing
 * rootApptId into namedValues['Admin: Calendly Event UID'] so
 * onFormSubmit skips creating a new root.
 *
 * Returns { ok, row, masterRowIndex, customerName, brand,
 *           linked, linkedRoot, linkedName, linkedMatchedBy }
 */
// function ipad_submitIntake(payload) {
//   try {
//     const p  = payload || {};
//     const tz = Session.getScriptTimeZone();

//     // ── Date / Time formatting ─────────────────────────────────────
//     let visitDateStr = '';
//     if (p.date) {
//       const d = new Date(p.date + 'T12:00:00');
//       visitDateStr = Utilities.formatDate(d, tz, 'MM/dd/yyyy');
//     }
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
//     const now      = new Date();

//     // ── Q2: Auto-link — check if contact already exists ───────────
//     let linkedRoot    = '';
//     let linkedName    = '';
//     let linkedMatchBy = '';
//     let isLinked      = false;

//     // Only auto-link if no UID was already provided by staff
//     const staffUID = String(p.uid || '').trim();
//     if (!staffUID && (p.email || p.phone)) {
//       const existing = ipad_findExistingRootByContact_(p.email, p.phone, p.company || '');
//       if (existing && existing.rootApptId) {
//         linkedRoot    = existing.rootApptId;
//         linkedName    = existing.customerName;
//         linkedMatchBy = existing.matchedBy;
//         isLinked      = true;

//         // ── Lấy PaymentsFolderURL từ row cũ để tái sử dụng ──────────
//         try {
//           const ss  = SpreadsheetApp.getActive();
//           const sh  = ss.getSheetByName(RP_MASTER_SHEET);
//           const lc  = sh.getLastColumn();
//           const hdr = sh.getRange(1,1,1,lc).getDisplayValues()[0];
//           const map = rp_headerMap([hdr]);
//           const pfIdx = rp_pick0(map, 'PaymentsFolderURL');
//           if (pfIdx >= 0 && existing.rowIndex) {
//             const existingRow = sh.getRange(existing.rowIndex, 1, 1, lc).getDisplayValues()[0];
//             const existingFolderURL = String(existingRow[pfIdx] || '').trim();
//             if (existingFolderURL) {
//               // Ghi vào namedValues để onFormSubmit nhận được
//               namedValues['PaymentsFolderURL'] = [existingFolderURL];
//               Logger.log('[ipad_submitIntake] Reusing existing folder: %s', existingFolderURL);
//             }
//           }
//         } catch(fe) {
//           Logger.log('[ipad_submitIntake] folder reuse error: ' + fe.message);
//         }
//       }
//     }

//     // ── Build namedValues ──────────────────────────────────────────
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
//       // Inject existing rootApptId so onFormSubmit uses it, not a new UID
//       'Admin: Calendly Event UID': [isLinked ? linkedRoot : staffUID],
//     };

//     // ── Append to 02_Form_Inbox ────────────────────────────────────
//     const ss = SpreadsheetApp.getActive();
//     const sh = ss.getSheetByName('02_Form_Inbox');
//     if (!sh) throw new Error('Sheet "02_Form_Inbox" not found');

//     const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
//     const rowData = headers.map(h => { const v = namedValues[h]; return v ? v[0] : ''; });

//     sh.appendRow(rowData);
//     const newRow = sh.getLastRow();
//     Logger.log('[ipad_submitIntake] Appended row=%s | linked=%s', newRow, isLinked);

//     // ── Trigger form processing ────────────────────────────────────
//     onFormSubmit({
//       namedValues: namedValues,
//       range: sh.getRange(newRow, 1, 1, headers.length),
//       values: rowData,
//     });

//     Logger.log('[ipad_submitIntake] onFormSubmit OK | name=%s', p.name || '');

//     // ── Find masterRowIndex ────────────────────────────────────────
//     let masterRowIndex = 0;
//     try {
//       // If linked, look up by rootApptId (most reliable)
//       if (isLinked && linkedRoot) {
//         masterRowIndex = ipad_findRowByRootApptId_(linkedRoot) || 0;
//       }
//       // Fallback: search by name from bottom up
//       if (!masterRowIndex) {
//         const ss2 = SpreadsheetApp.getActive();
//         const mSh = ss2.getSheetByName(RP_MASTER_SHEET);
//         if (mSh) {
//           const lr2 = mSh.getLastRow(), lc2 = mSh.getLastColumn();
//           if (lr2 >= 2) {
//             const hdr2   = mSh.getRange(1,1,1,lc2).getDisplayValues()[0];
//             const map2   = rp_headerMap([hdr2]);
//             const nIdx   = rp_pick0(map2,'Customer Name','Customer','Client Name');
//             const vals2  = mSh.getRange(2,1,lr2-1,lc2).getDisplayValues();
//             const target = String(p.name||'').trim().toLowerCase();
//             for (let i = vals2.length - 1; i >= 0; i--) {
//               const rowName = nIdx >= 0 ? String(vals2[i][nIdx]||'').trim().toLowerCase() : '';
//               if (rowName && rowName === target) { masterRowIndex = i + 2; break; }
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
//       // Linking info (for optional UI feedback)
//       linked:           isLinked,
//       linkedRoot:       linkedRoot,
//       linkedName:       linkedName,
//       linkedMatchedBy:  linkedMatchBy,
//     };

//   } catch(e) {
//     Logger.log('[ipad_submitIntake] ERROR: ' + e.message + '\n' + (e.stack||''));
//     return { ok: false, error: e.message };
//   }
// }


// ══════════════════════════════════════════════════════════════════════
// FAST INTAKE SUBMIT — Phase 1 sync (~300ms) + Phase 2 async (trigger)
// ══════════════════════════════════════════════════════════════════════

/**
 * Phase 1: Ghi row tối thiểu vào Master → return rowIndex NGAY (~300ms)
 * Phase 2: Queue payload → time-trigger xử lý nặng trong nền
 *
 * Auto-linking vẫn chạy ở Phase 1 (chỉ đọc sheet, không ghi nhiều)
 * onFormSubmit() nặng được dời sang processIntakeQueue()
 */
function ipad_submitIntake(payload) {
  try {
    const p = payload || {};

    // ── Auto-link check (chỉ đọc) ─────────────────────────────────
    let linkedRoot = '', linkedName = '', linkedMatchBy = '', isLinked = false;
    const staffUID = String(p.uid || '').trim();
    if (!staffUID && (p.email || p.phone)) {
      const existing = ipad_findExistingRootByContact_(p.email, p.phone, p.company || '');
      if (existing && existing.rootApptId) {
        linkedRoot    = existing.rootApptId;
        linkedName    = existing.customerName;
        linkedMatchBy = existing.matchedBy;
        isLinked      = true;
      }
    }
    const resolvedUID = isLinked ? linkedRoot : staffUID;
    const brand = String(p.company || p.brand || 'HPUSA')
                    .toUpperCase().includes('VVS') ? 'VVS' : 'HPUSA';

    // ── Chỉ ghi vào queue, KHÔNG ghi vào Master ───────────────────
    const queueId = _intake_queueOnly_(p, resolvedUID, brand);

    return {
      ok:             true,
      masterRowIndex: 0,          // chưa có — client nhận biết qua tempId
      tempId:         queueId,    // ID tạm để track
      rootApptId:     resolvedUID || '',
      customerName:   p.name  || '',
      brand:          brand,
      status:         'QUEUED',
      linked:         isLinked,
      linkedRoot:     linkedRoot,
      linkedName:     linkedName,
      linkedMatchedBy: linkedMatchBy,
      queueRow: queueId,
    };

  } catch(e) {
    Logger.log('[ipad_submitIntake] ERROR: ' + e.message);
    return { ok: false, error: e.message };
  }
}


/**
 * Chỉ ghi vào _IntakeQueue — không động vào Master sheet.
 * Return queueId để client track.
 */
function _intake_queueOnly_(p, resolvedUID, brand) {
  const ss    = SpreadsheetApp.getActive();
  let   bufSh = ss.getSheetByName('_IntakeQueue');

  if (!bufSh) {
    bufSh = ss.insertSheet('_IntakeQueue');
    bufSh.appendRow(['QueuedAt','Status','MasterRowIndex','ResolvedUID','Brand','Payload','ProcessedAt','Error']);
    bufSh.setFrozenRows(1);
    bufSh.hideSheet();
  }
  if (bufSh.getLastRow() === 0) {
    bufSh.appendRow(['QueuedAt','Status','MasterRowIndex','ResolvedUID','Brand','Payload','ProcessedAt','Error']);
    bufSh.setFrozenRows(1);
  }

  bufSh.appendRow([
    new Date(),       // col 1: QueuedAt
    'PENDING',        // col 2: Status
    0,                // col 3: MasterRowIndex (chưa biết)
    resolvedUID || '',// col 4: ResolvedUID
    brand,            // col 5: Brand
    JSON.stringify(p),// col 6: Payload
    '',               // col 7: ProcessedAt
    '',               // col 8: Error
  ]);

  const newRow = bufSh.getLastRow();  // ← số dòng thật trong sheet
  Logger.log('[_intake_queueOnly_] Queued row=%s brand=%s name=%s', newRow, brand, p.name || '');
  return newRow;  // ← return NUMBER, không phải string ID
}


/**
 * Ghi những cột tối thiểu để ipad_loadRecord() chạy được ngay.
 * Không tạo folder, không gọi onFormSubmit.
 *
 * @param {object}  p            payload từ client
 * @param {string}  resolvedUID  UID thật (staffUID hoặc linkedRoot)
 * @param {string}  existingRoot nếu linked, dùng root này thay vì tạo mới
 */
function _intake_writeMinimalMasterRow_(p, resolvedUID, existingRoot) {
  const ss  = SpreadsheetApp.getActive();
  const mSh = ss.getSheetByName(RP_MASTER_SHEET);
  const tz  = Session.getScriptTimeZone();
  const now = new Date();

  const brand = String(p.company || p.brand || 'HPUSA')
                  .toUpperCase().includes('VVS') ? 'VVS' : 'HPUSA';

  // Dùng root cũ nếu linked, ngược lại tạo temp ID
  const rootApptId = existingRoot ||
    (resolvedUID && !resolvedUID.startsWith('WALKIN-') ? resolvedUID : '') ||
    ('WALKIN-' + Utilities.formatDate(now, tz, 'yyyyMMdd-HHmmss'));

  // Nếu linked và đã có PaymentsFolderURL, kế thừa lại
  let existingFolderURL = '';
  if (existingRoot) {
    try {
      const lc  = mSh.getLastColumn();
      const hdr = mSh.getRange(1,1,1,lc).getDisplayValues()[0];
      const map = rp_headerMap([hdr]);
      const apIdx = rp_pick0(map,'APPT_ID','RootApptID','Root Appt ID');
      const pfIdx = rp_pick0(map,'PaymentsFolderURL');
      if (apIdx >= 0 && pfIdx >= 0 && mSh.getLastRow() >= 2) {
        const vals = mSh.getRange(2,1,mSh.getLastRow()-1,lc).getDisplayValues();
        for (let i = 0; i < vals.length; i++) {
          if (String(vals[i][apIdx]||'').trim() === existingRoot) {
            const u = String(vals[i][pfIdx]||'').trim();
            if (u) { existingFolderURL = u; break; }
          }
        }
      }
    } catch(_) {}
  }

  const minimal = {
    'Customer Name':     p.name  || '',
    'Phone':             p.phone || '',
    'Email':             p.email || '',
    'Brand':             brand,
    'APPT_ID':           rootApptId,
    'RootApptID':        rootApptId,
    'Order Total':       '',
    'Paid-to-Date':      0,
    'Remaining Balance': 0,
    'PaymentsFolderURL': existingFolderURL,
    '_IntakeStatus':     'PENDING',
    '_QueuedAt':         Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'),
  };

  const lc  = mSh.getLastColumn();
  const hdr = mSh.getRange(1,1,1,lc).getValues()[0].map(v => String(v).trim());
  const row = hdr.map(h => minimal[h] !== undefined ? minimal[h] : '');

  mSh.appendRow(row);
  const rowIndex = mSh.getLastRow();

  Logger.log('[_intake_writeMinimalMasterRow_] row=%s rootId=%s linked=%s folder=%s',
    rowIndex, rootApptId, !!existingRoot, existingFolderURL ? 'yes' : 'no');

  return { rowIndex, rootApptId, customerName: p.name || '', brand };
}


/**
 * Ghi vào sheet buffer _IntakeQueue.
 * Time-trigger processIntakeQueue() sẽ đọc và gọi onFormSubmit thật.
 */
function _intake_queueForBackground_(p, masterRowIndex, resolvedUID) {
  const ss    = SpreadsheetApp.getActive();
  let   bufSh = ss.getSheetByName('_IntakeQueue');

  if (!bufSh) {
    bufSh = ss.insertSheet('_IntakeQueue');
    bufSh.appendRow(['QueuedAt','Status','MasterRowIndex','ResolvedUID','Payload','ProcessedAt','Error']);
    bufSh.setFrozenRows(1);
    bufSh.hideSheet();
  }

  // Đảm bảo header tồn tại nếu sheet đã có nhưng chưa có header
  if (bufSh.getLastRow() === 0) {
    bufSh.appendRow(['QueuedAt','Status','MasterRowIndex','ResolvedUID','Payload','ProcessedAt','Error']);
    bufSh.setFrozenRows(1);
  }

  bufSh.appendRow([
    new Date(),
    'PENDING',
    masterRowIndex,
    resolvedUID || '',
    JSON.stringify(p),
    '',
    '',
  ]);
}


/**
 * Poll endpoint — client gọi mỗi 15 giây để check folder đã tạo xong chưa.
 * Chỉ đọc 2 cell → cực nhẹ.
 */
function ipad_checkFolderReady(rowIndex) {
  try {
    rowIndex = Number(rowIndex);
    if (rowIndex < 2) return null;

    const ss  = SpreadsheetApp.getActive();
    const mSh = ss.getSheetByName(RP_MASTER_SHEET);
    const lc  = mSh.getLastColumn();
    const hdr = mSh.getRange(1,1,1,lc).getValues()[0].map(v => String(v).trim());

    const pfIdx  = hdr.indexOf('PaymentsFolderURL');
    const staIdx = hdr.indexOf('_IntakeStatus');
    if (pfIdx < 0) return null;

    const row = mSh.getRange(rowIndex, 1, 1, lc).getValues()[0];
    const url = String(row[pfIdx] || '').trim();
    if (!url) return null;

    return {
      paymentsFolderURL: url,
      status: staIdx >= 0 ? String(row[staIdx] || '').trim() : 'DONE',
    };
  } catch(e) {
    return null;
  }
}
/**
 * Find master row index by rootApptId. Returns rowIndex (≥2) or 0.
 */
function ipad_findRowByRootApptId_(rootApptId) {
  if (!rootApptId) return 0;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RP_MASTER_SHEET);
  if (!sh) return 0;
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return 0;
  const header = sh.getRange(1,1,1,lc).getDisplayValues()[0];
  const map    = rp_headerMap([header]);
  const apptIdx = rp_pick0(map,'APPT_ID','RootApptID','Root Appt ID');
  if (apptIdx < 0) return 0;
  const vals = sh.getRange(2,1,lr-1,lc).getDisplayValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][apptIdx]||'').trim() === String(rootApptId).trim()) return i + 2;
  }
  return 0;
}


// ═══════════════════════════════════════════════════════════════════════
// Q4 — LOAD RECORD WITH ROOT AGGREGATION
// ═══════════════════════════════════════════════════════════════════════

/**
 * Aggregated load: reads ONE canonical row for metadata but sums
 * financial data across ALL rows sharing the same rootApptId.
 *
 * Drop-in replacement for the existing ipad_loadRecord().
 * Adds: allLinkedRows — array of { rowIndex, customerName } for UI display.
 */
function ipad_loadRecord(rowIndex) {
  try {
    rowIndex = Number(rowIndex);
    if (rowIndex < 2) throw new Error('Invalid row index');

    // ── Load primary row ──────────────────────────────────────────
    const m     = rp_getMasterRowByIndex_(rowIndex);
    const brand = (m.map['Brand'] != null)
                  ? String(m.rowVals[m.map['Brand']] || '').trim() : '';
    const hasSO      = !!(m.soNumber && String(m.soNumber).trim());
    const anchorType = hasSO ? 'SO' : 'APPT';

    let taxRate = 0;
    try { taxRate = brand ? rp_getTaxRate_(brand) : 0; } catch(_) {}

    const ptdIdx = rp_pick0(m.map, 'Paid-to-Date','Paid-To-Date','Paid to Date','Paid-to-date');
    const otIdx  = m.map['Order Total'] != null ? m.map['Order Total'] : -1;
    const rbIdx  = rp_pick0(m.map, 'Remaining Balance','Balance');

    const ss2  = SpreadsheetApp.getActive().getSheetByName(RP_MASTER_SHEET);
    const raw  = ss2.getRange(rowIndex, 1, 1, ss2.getLastColumn()).getValues()[0];

    let paidToDate   = ptdIdx >= 0 ? rp_num_(raw[ptdIdx]) : 0;
    let orderTotal   = otIdx  >= 0 ? rp_num_(raw[otIdx])  : 0;

    // ── Q4: Aggregate across all rows with same rootApptId ─────────
    const rootApptId = m.rootApptId || '';
    let allLinkedRows = [];

    if (rootApptId) {
      const aggResult = ipad_aggregateByRoot_(rootApptId, brand, rowIndex);
      if (aggResult.rowCount > 1) {
        // Use aggregated financials only if multiple linked rows found
        paidToDate = aggResult.totalPaidToDate;
        // OT: use highest value found (most recent invoice total)
        if (aggResult.maxOrderTotal > orderTotal) orderTotal = aggResult.maxOrderTotal;
        allLinkedRows = aggResult.rows;
        Logger.log('[ipad_loadRecord] Aggregated %s rows for root=%s OT=%s PTD=%s',
          aggResult.rowCount, rootApptId, orderTotal, paidToDate);
      }
    }

    // ── Balance ────────────────────────────────────────────────────
    let balance;
    if (orderTotal > 0 && rbIdx >= 0 && allLinkedRows.length <= 1) {
      balance = Math.max(0, rp_num_(raw[rbIdx]));
    } else {
      balance = Math.max(0, orderTotal - paidToDate);
    }

    // ── OT fallback from ledger ────────────────────────────────────
    if (!orderTotal) {
      try {
        const lastIT = rp_getLastInvoiceTotalFromLedger_({ hasSO, soNumber: m.soNumber, rootApptId: m.rootApptId });
        if (lastIT > 0) { orderTotal = lastIT; balance = Math.max(0, lastIT - paidToDate); }
      } catch(_) {}
    }

    // ── taxEnabled (mirror existing logic) ────────────────────────
    let taxEnabled = true, foundTaxRecord = false;
    try {
      const { sh: lSh } = rp_getLedgerTarget();
      const lLr = lSh.getLastRow(), lLc = lSh.getLastColumn();
      if (lLr >= 2) {
        const lHdr = lSh.getRange(1,1,1,lLc).getValues()[0].map(v=>String(v).trim());
        const LH = {}; lHdr.forEach((h,i)=>LH[h]=i);
        const cTax = LH['TaxEnabled'], cAppt = LH['RootApptID'], cSO = LH['SO#'];
        if (cTax != null) {
          const start = Math.max(2, lLr - 500);
          const lVals = lSh.getRange(start,1,lLr-start+1,lLc).getValues();
          for (let i = lVals.length-1; i >= 0; i--) {
            const r = lVals[i];
            const match = hasSO ? rp_soEq(r[cSO],m.soNumber) : String(r[cAppt]||'').trim()===String(m.rootApptId||'').trim();
            if (match) {
              const rawTax = r[cTax];
              if (rawTax !== '' && rawTax !== null && rawTax !== undefined) {
                taxEnabled = !(rawTax===false||String(rawTax).toLowerCase()==='false');
                foundTaxRecord = true;
              }
              break;
            }
          }
        }
      }
    } catch(te) { Logger.log('[ipad_loadRecord taxEnabled] ' + te.message); }

    if (!foundTaxRecord && paidToDate > 0) taxEnabled = false;

    // ── Previous payments ──────────────────────────────────────────
    let prevPayments = [];
    try {
      const prev = rp_prevPaymentsForAnchor_({ anchorType, rootApptId: m.rootApptId, soNumber: m.soNumber, limit: 10 });
      if (prev && prev.items) {
        prevPayments = prev.items.map(it => ({ date: it.date||'', amount: Number(it.amount||0), method: it.method||'', docNumber: it.docNumber||'' }));
      }
    } catch(_) {}

    // ── Extra columns ──────────────────────────────────────────────
    const phoneIdx = rp_pick0(m.map,'Phone','Phone Number','Tel','Mobile');
    const emailIdx = rp_pick0(m.map,'Email','Email Address','E-mail');
    const pfIdx        = rp_pick0(m.map, 'PaymentsFolderURL');
    const intakeIdx    = rp_pick0(m.map, 'IntakeDocURL');
    const checklistIdx = rp_pick0(m.map, 'ChecklistURL', 'Checklist URL');
    const quotationIdx = rp_pick0(m.map, 'QuotationURL', 'Quotation URL');

    // ── Kế thừa IntakeDocURL / ChecklistURL / QuotationURL ─────
    let intakeDocURL    = intakeIdx    >= 0 ? String(m.rowVals[intakeIdx]    || '').trim() : '';
    let checklistURL    = checklistIdx >= 0 ? String(m.rowVals[checklistIdx] || '').trim() : '';
    let quotationURL    = quotationIdx >= 0 ? String(m.rowVals[quotationIdx] || '').trim() : '';

    if (!intakeDocURL || !checklistURL || !quotationURL) {
      try {
        const _ss3 = SpreadsheetApp.getActive();
        const _sh3 = _ss3.getSheetByName(RP_MASTER_SHEET);
        const _lr3 = _sh3.getLastRow(), _lc3 = _sh3.getLastColumn();
        if (_lr3 >= 2 && rootApptId) {
          const _hdr3   = _sh3.getRange(1,1,1,_lc3).getDisplayValues()[0];
          const _map3   = rp_headerMap([_hdr3]);
          const _apIdx3 = rp_pick0(_map3,'APPT_ID','RootApptID','Root Appt ID');
          const _inIdx3 = rp_pick0(_map3,'IntakeDocURL');
          const _chIdx3 = rp_pick0(_map3,'ChecklistURL','Checklist URL');
          const _quIdx3 = rp_pick0(_map3,'QuotationURL','Quotation URL');
          if (_apIdx3 >= 0) {
            const _vals3 = _sh3.getRange(2,1,_lr3-1,_lc3).getDisplayValues();
            for (let _i3 = 0; _i3 < _vals3.length; _i3++) {
              if (String(_vals3[_i3][_apIdx3]||'').trim() !== rootApptId) continue;
              if (!intakeDocURL && _inIdx3 >= 0) {
                const _u = String(_vals3[_i3][_inIdx3]||'').trim();
                if (_u) { intakeDocURL = _u; Logger.log('[ipad_loadRecord] IntakeDocURL inherited from row '+(_i3+2)+': '+_u); }
              }
              if (!checklistURL && _chIdx3 >= 0) {
                const _u = String(_vals3[_i3][_chIdx3]||'').trim();
                if (_u) { checklistURL = _u; Logger.log('[ipad_loadRecord] ChecklistURL inherited from row '+(_i3+2)+': '+_u); }
              }
              if (!quotationURL && _quIdx3 >= 0) {
                const _u = String(_vals3[_i3][_quIdx3]||'').trim();
                if (_u) { quotationURL = _u; Logger.log('[ipad_loadRecord] QuotationURL inherited from row '+(_i3+2)+': '+_u); }
              }
              if (intakeDocURL && checklistURL && quotationURL) break;
            }
          }
        }
      } catch(_) {}
    }

    // ── Saved lines ────────────────────────────────────────────────
    let savedLines = [];
    try {
      const saved = rp_readSavedLinesFromMaster_(m) || rp_findLastSavedLinesForAnchor_({ anchorType, rootApptId: m.rootApptId, soNumber: m.soNumber });
      if (saved && saved.lines) savedLines = saved.lines;
    } catch(_) {}

    // ── P21: hasSalesInvoice ───────────────────────────────────────
    let hasSalesInvoice = false;
    try {
      const invCheck = rp_checkInvoiceBeforeReceipt_({ anchorType, rootApptId: m.rootApptId||'', soNumber: m.soNumber||'', docType: 'Sales Receipt' });
      hasSalesInvoice = invCheck.ok;
    } catch(_) {}

    // ── P22: referral ──────────────────────────────────────────────
    let referralUsed = false, referralNameVal = '', referralDiscountVal = 0;
    try {
      const refCheck = rp_isReferralAlreadyUsed_({ anchorType, rootApptId: m.rootApptId||'', soNumber: m.soNumber||'' });
      referralUsed = refCheck.used; referralNameVal = refCheck.name||''; referralDiscountVal = refCheck.discount||0;
    } catch(_) {}

    Logger.log('[ipad_loadRecord] row=%s OT=%s PTD=%s BAL=%s taxEnabled=%s linkedRows=%s',
      rowIndex, orderTotal, paidToDate, balance, taxEnabled, allLinkedRows.length);

    return {
      ok: true,
      rowIndex, anchorType, brand,
      hasPrint: brand.toUpperCase().includes('HPUSA'),
      taxEnabled,
      hasSalesInvoice,
      referralUsed, referralName: referralNameVal, referralDiscount: referralDiscountVal,
      customerName:      m.customerName,
      soNumber:          m.soNumber || '',
      rootApptId:        m.rootApptId || '',
      trackerUrl:        m.trackerUrl || '',
      phone:             phoneIdx >= 0 ? String(m.rowVals[phoneIdx]||'').trim() : '',
      email:             emailIdx >= 0 ? String(m.rowVals[emailIdx]||'').trim() : '',
      orderTotal:        String(orderTotal || ''),
      paidToDate:        String(paidToDate || ''),
      balance:           String(balance),
      taxRate,
      paymentsFolderURL: (function(){
        // Ưu tiên URL từ row hiện tại
        const ownUrl = pfIdx >= 0 ? String(m.rowVals[pfIdx]||'').trim() : '';
        if (ownUrl) return ownUrl;

        // Nếu trống → tìm trong tất cả rows cùng rootApptId
        if (!rootApptId) return '';
        try {
          const ss2b = SpreadsheetApp.getActive();
          const sh2b = ss2b.getSheetByName(RP_MASTER_SHEET);
          const lr2b = sh2b.getLastRow(), lc2b = sh2b.getLastColumn();
          if (lr2b < 2) return '';

          const hdr2b = sh2b.getRange(1,1,1,lc2b).getDisplayValues()[0];
          const map2b = rp_headerMap([hdr2b]);
          const apptIdx2b = rp_pick0(map2b,'APPT_ID','RootApptID','Root Appt ID');
          const pfIdx2b   = rp_pick0(map2b,'PaymentsFolderURL');
          if (apptIdx2b < 0 || pfIdx2b < 0) return '';

          const vals2b = sh2b.getRange(2,1,lr2b-1,lc2b).getDisplayValues();
          for (let i = 0; i < vals2b.length; i++) {
            if (String(vals2b[i][apptIdx2b]||'').trim() !== rootApptId) continue;
            const u = String(vals2b[i][pfIdx2b]||'').trim();
            if (u) {
              Logger.log('[ipad_loadRecord] PaymentsFolderURL inherited from row '+(i+2)+': '+u);
              return u;
            }
          }
        } catch(_) {}
        return '';
      })(),
      intakeDocURL,
      checklistURL,
      quotationURL,
      prevPayments,
      savedLines,
      // Q4: linked rows for UI badge
      allLinkedRows,
    };

  } catch(e) {
    Logger.log('[ipad_loadRecord] ERROR: ' + e.message);
    return { ok: false, error: e.message };
  }
}

/**
 * Aggregate financial data for all master rows sharing a rootApptId.
 *
 * @param {string} rootApptId
 * @param {string} brand
 * @param {number} primaryRowIndex  — excluded from "other rows" list
 * @returns {{ rowCount, totalPaidToDate, maxOrderTotal, rows[] }}
 */
function ipad_aggregateByRoot_(rootApptId, brand, primaryRowIndex) {
  const result = { rowCount: 0, totalPaidToDate: 0, maxOrderTotal: 0, rows: [] };
  if (!rootApptId) return result;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RP_MASTER_SHEET);
  if (!sh) return result;

  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return result;

  const header = sh.getRange(1,1,1,lc).getDisplayValues()[0];
  const map    = rp_headerMap([header]);

  const apptIdx  = rp_pick0(map,'APPT_ID','RootApptID','Root Appt ID');
  const nameIdx  = rp_pick0(map,'Customer Name','Customer','Client Name');
  const brandIdx = rp_pick0(map,'Brand');
  const ptdIdx   = rp_pick0(map,'Paid-to-Date','Paid-To-Date','Paid to Date','Paid-to-date');
  const otIdx    = map['Order Total'] != null ? map['Order Total'] : -1;

  if (apptIdx < 0) return result;

  const bFilter = brand ? String(brand).trim().toUpperCase() : '';
  const vals = sh.getRange(2,1,lr-1,lc).getValues();

  for (let i = 0; i < vals.length; i++) {
    const row     = vals[i];
    const rowRoot = String(row[apptIdx] || '').trim();
    if (rowRoot !== rootApptId) continue;

    const rowBrand = brandIdx >= 0 ? String(row[brandIdx]||'').trim().toUpperCase() : '';
    if (bFilter && !rowBrand.includes(bFilter)) continue;

    const rowIdx  = i + 2;
    const ptd     = ptdIdx >= 0 ? rp_num_(row[ptdIdx]) : 0;
    const ot      = otIdx  >= 0 ? rp_num_(row[otIdx])  : 0;
    const name    = nameIdx >= 0 ? String(row[nameIdx]||'').trim() : '';

    result.rowCount++;
    result.totalPaidToDate += ptd;
    if (ot > result.maxOrderTotal) result.maxOrderTotal = ot;

    result.rows.push({
      rowIndex: rowIdx,
      customerName: name,
      isPrimary: rowIdx === primaryRowIndex,
    });
  }

  return result;
}


// ═══════════════════════════════════════════════════════════════════════
// Q3 — DEDUP UTILITIES (run manually from Apps Script editor)
// ═══════════════════════════════════════════════════════════════════════

/**
 * Scans the master sheet for rows sharing email or phone.
 * Outputs a report to Logger — run from Apps Script editor.
 *
 * Usage:
 *   Open Apps Script editor → run ipad_findDuplicateContacts()
 *   Review the Execution Log
 *
 * @returns {Array} list of conflict groups for further processing
 */
function ipad_findDuplicateContacts() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RP_MASTER_SHEET);
  if (!sh) { Logger.log('Sheet not found'); return []; }

  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return [];

  const header = sh.getRange(1,1,1,lc).getDisplayValues()[0];
  const map    = rp_headerMap([header]);

  const nameIdx  = rp_pick0(map,'Customer Name','Customer','Client Name');
  const emailIdx = rp_pick0(map,'Email','Email Address','E-mail');
  const phoneIdx = rp_pick0(map,'Phone','Phone Number','Tel','Mobile');
  const apptIdx  = rp_pick0(map,'APPT_ID','RootApptID','Root Appt ID');
  const brandIdx = rp_pick0(map,'Brand');

  const vals = sh.getRange(2,1,lr-1,lc).getDisplayValues();

  // Build lookup maps
  const byEmail = {}, byPhone = {};

  for (let i = 0; i < vals.length; i++) {
    const row   = vals[i];
    const ri    = i + 2;
    const name  = nameIdx  >= 0 ? String(row[nameIdx] ||'').trim() : '';
    const email = emailIdx >= 0 ? ipad_normalizeEmail_(row[emailIdx])  : '';
    const phone = phoneIdx >= 0 ? ipad_normalizePhone_(row[phoneIdx])  : '';
    const root  = apptIdx  >= 0 ? String(row[apptIdx] ||'').trim() : '';
    const brand = brandIdx >= 0 ? String(row[brandIdx]||'').trim() : '';

    const entry = { rowIndex: ri, name, root, brand };

    if (email.includes('@')) {
      if (!byEmail[email]) byEmail[email] = [];
      byEmail[email].push(entry);
    }
    if (phone.length >= 10) {
      if (!byPhone[phone]) byPhone[phone] = [];
      byPhone[phone].push(entry);
    }
  }

  const conflicts = [];

  // Email conflicts
  for (const [email, rows] of Object.entries(byEmail)) {
    if (rows.length < 2) continue;
    const roots = [...new Set(rows.map(r => r.root).filter(Boolean))];
    if (roots.length > 1) {
      Logger.log('[DEDUP EMAIL] %s → %s rows, %s different roots: %s',
        email, rows.length, roots.length,
        rows.map(r=>`row${r.rowIndex} "${r.name}" root:${r.root}`).join(' | '));
      conflicts.push({ type: 'email', key: email, rows, roots });
    }
  }

  // Phone conflicts
  for (const [phone, rows] of Object.entries(byPhone)) {
    if (rows.length < 2) continue;
    const roots = [...new Set(rows.map(r => r.root).filter(Boolean))];
    if (roots.length > 1) {
      Logger.log('[DEDUP PHONE] %s → %s rows, %s different roots: %s',
        phone, rows.length, roots.length,
        rows.map(r=>`row${r.rowIndex} "${r.name}" root:${r.root}`).join(' | '));
      conflicts.push({ type: 'phone', key: phone, rows, roots });
    }
  }

  Logger.log('[ipad_findDuplicateContacts] Total conflicts: %s', conflicts.length);
  if (conflicts.length === 0) Logger.log('✅ No duplicate contacts found.');

  return conflicts;
}

/**
 * Re-link a master row to a different rootApptId.
 *
 * Use after ipad_findDuplicateContacts() to fix specific rows.
 *
 * Example (run in Apps Script editor):
 *   // "Yasmin & Lupe" (row 45) should share root with "Yasmin Gonzalez" (row 12)
 *   ipad_relinkRecord({ rowIndex: 45, newRootApptId: 'abc123-ROOT-ID-OF-YASMIN' });
 *
 * @param {{ rowIndex: number, newRootApptId: string, dryRun?: boolean }} params
 */
function ipad_relinkRecord(params) {
  const p        = params || {};
  const rowIndex = Number(p.rowIndex);
  const newRoot  = String(p.newRootApptId || '').trim();
  const dryRun   = !!p.dryRun;

  if (rowIndex < 2 || !newRoot) {
    Logger.log('[ipad_relinkRecord] Invalid params: rowIndex=%s newRoot=%s', rowIndex, newRoot);
    return { ok: false, error: 'rowIndex < 2 or newRootApptId empty' };
  }

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RP_MASTER_SHEET);
  if (!sh) return { ok: false, error: 'Master sheet not found' };

  const lc     = sh.getLastColumn();
  const header = sh.getRange(1,1,1,lc).getDisplayValues()[0];
  const map    = rp_headerMap([header]);
  const apptIdx = rp_pick0(map,'APPT_ID','RootApptID','Root Appt ID');

  if (apptIdx < 0) return { ok: false, error: 'No RootApptID column found' };

  const row     = sh.getRange(rowIndex, 1, 1, lc).getDisplayValues()[0];
  const oldRoot = String(row[apptIdx] || '').trim();
  const name    = (() => { const ni = rp_pick0(map,'Customer Name','Customer'); return ni >= 0 ? String(row[ni]||'').trim() : '?'; })();

  Logger.log('[ipad_relinkRecord] %s: row=%s "%s" oldRoot=%s → newRoot=%s dryRun=%s',
    dryRun ? 'DRY RUN' : 'EXECUTING', rowIndex, name, oldRoot, newRoot, dryRun);

  if (dryRun) return { ok: true, dryRun: true, rowIndex, name, oldRoot, newRoot };

  // Write new root to the cell
  sh.getRange(rowIndex, apptIdx + 1).setValue(newRoot);
  SpreadsheetApp.flush();

  Logger.log('[ipad_relinkRecord] ✅ Done. row=%s "%s" now has root=%s', rowIndex, name, newRoot);
  return { ok: true, rowIndex, name, oldRoot, newRoot };
}

/**
 * Convenience: relink multiple rows at once.
 *
 * Example:
 *   ipad_relinkBatch([
 *     { rowIndex: 45, newRootApptId: 'abc123' },
 *     { rowIndex: 78, newRootApptId: 'abc123' },
 *   ]);
 */
function ipad_relinkBatch(items, dryRun) {
  const results = (items || []).map(item =>
    ipad_relinkRecord(Object.assign({}, item, { dryRun: !!dryRun }))
  );
  Logger.log('[ipad_relinkBatch] %s items processed. dryRun=%s', results.length, !!dryRun);
  return results;
}

// ══════════════════════════════════════════════════════════════════════
// TIME-TRIGGER — processIntakeQueue
// Cài đặt: Apps Script Editor → Triggers → Add Trigger
//   Function: processIntakeQueue | Time-based | Minute timer | Every 1 minute
// ══════════════════════════════════════════════════════════════════════

function processIntakeQueue() {
  const ss    = SpreadsheetApp.getActive();
  const bufSh = ss.getSheetByName('_IntakeQueue');
  if (!bufSh || bufSh.getLastRow() < 2) return;

  // ── Lock để tránh 2 trigger chạy đồng thời ───────────────────
  const lock = LockService.getScriptLock();
  try { lock.waitLock(5000); } catch(_) {
    Logger.log('[processIntakeQueue] Could not acquire lock — skipping this run');
    return;
  }

  try {
    const numRows = bufSh.getLastRow() - 1;
    const data    = bufSh.getRange(2, 1, numRows, 8).getValues();
    const tz      = Session.getScriptTimeZone();

    for (let i = 0; i < data.length; i++) {
      const row    = data[i];
      const status = String(row[1] || '').trim();

      // ── Chỉ xử lý PENDING, bỏ qua RUNNING / DONE / ERROR ──────
      if (status !== 'PENDING') continue;

      const sheetRow    = i + 2;
      const resolvedUID = String(row[3] || '').trim();
      const payloadJson = String(row[5] || '{}');

      // Đánh dấu RUNNING ngay — nhả lock sau đó
      bufSh.getRange(sheetRow, 2).setValue('RUNNING');
      SpreadsheetApp.flush();
      lock.releaseLock();

      try {
        const p = JSON.parse(payloadJson);

        let visitDateStr = '';
        if (p.date) {
          try { visitDateStr = Utilities.formatDate(new Date(p.date + 'T12:00:00'), tz, 'MM/dd/yyyy'); }
          catch(_) { visitDateStr = p.date || ''; }
        }
        let visitTimeStr = '';
        if (p.time) {
          const tp = p.time.split(':');
          let h = parseInt(tp[0], 10);
          const m = tp[1] || '00';
          const ap = h >= 12 ? 'PM' : 'AM';
          h = h % 12 || 12;
          visitTimeStr = h + ':' + m + ' ' + ap;
        }

        const diamonds = Array.isArray(p.diamond) ? p.diamond : (p.diamond?[p.diamond]:[]);
        const budgets  = Array.isArray(p.budget)  ? p.budget  : (p.budget ?[p.budget] :[]);
        const sources  = Array.isArray(p.source)  ? p.source  : (p.source ?[p.source] :[]);
        const now      = new Date();

        const namedValues = {
          'Timestamp':                 [Utilities.formatDate(now, tz, 'M/d/yyyy H:mm:ss')],
          'Company':                   [p.company   || row[4] || ''],
          'Customer Name':             [p.name      || ''],
          'Phone':                     [p.phone     || ''],
          'Email':                     [p.email     || ''],
          'Visit Type':                [p.visitType || 'Walk-In'],
          'Visit Date':                [visitDateStr],
          'Visit Time':                [visitTimeStr],
          'Location':                  [p.location  || 'In Store'],
          'Diamond Type':              [diamonds.join(', ')],
          'Budget Range':              [budgets.join(', ')],
          'Source':                    [sources.join(', ')],
          'Style Notes':               [p.notes || ''],
          'Admin: Calendly Event UID': [resolvedUID || ''],
        };

        const inboxSh = ss.getSheetByName('02_Form_Inbox');
        if (!inboxSh) throw new Error('Sheet "02_Form_Inbox" not found');

        const headers = inboxSh.getRange(1,1,1,inboxSh.getLastColumn()).getValues()[0];
        const rowData = headers.map(h => { const v=namedValues[h]; return v?v[0]:''; });
        inboxSh.appendRow(rowData);
        const newInboxRow = inboxSh.getLastRow();

        const mSh        = ss.getSheetByName(RP_MASTER_SHEET);
        const rowsBefore  = mSh ? mSh.getLastRow() : 0;

        onFormSubmit({
          namedValues: namedValues,
          range:       inboxSh.getRange(newInboxRow, 1, 1, headers.length),
          values:      rowData,
        });
        SpreadsheetApp.flush();

        const rowsAfter    = mSh ? mSh.getLastRow() : 0;
        const newRowsAdded = rowsAfter - rowsBefore;

        Logger.log('[processIntakeQueue] sheetRow=%s before=%s after=%s newRows=%s',
          sheetRow, rowsBefore, rowsAfter, newRowsAdded);

        // masterRowIndex = 0 vì architecture mới không ghi minimal row
        if (newRowsAdded > 0 && mSh) {
          // Không có minimal row để merge → chỉ cần lấy row mới nhất
          const newMasterRow = rowsBefore + 1;
          bufSh.getRange(sheetRow, 3).setValue(newMasterRow);
        }

        bufSh.getRange(sheetRow, 2).setValue('DONE');
        bufSh.getRange(sheetRow, 7).setValue(new Date());
        Logger.log('[processIntakeQueue] DONE sheetRow=%s name=%s', sheetRow, p.name||'');

      } catch(e) {
        bufSh.getRange(sheetRow, 2).setValue('ERROR');
        bufSh.getRange(sheetRow, 8).setValue(e.message);
        Logger.log('[processIntakeQueue] ERROR sheetRow=%s: %s', sheetRow, e.message);
      }

      // Re-acquire lock cho item tiếp theo
      try { lock.waitLock(5000); } catch(_) {
        Logger.log('[processIntakeQueue] Lost lock — stopping this run');
        return;
      }
    }

  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

/**
 * Merge data từ các row mới (onFormSubmit tạo) vào minimal row,
 * rồi xóa sạch các row mới thừa.
 *
 * @param {Sheet}  mSh            Master sheet
 * @param {number} minimalRowIdx  Row minimal đã có (do Phase 1 tạo)
 * @param {number} firstNewRow    Row đầu tiên onFormSubmit thêm vào
 * @param {number} lastNewRow     Row cuối cùng onFormSubmit thêm vào
 */
function _intake_mergeAndCleanup_(mSh, minimalRowIdx, firstNewRow, lastNewRow) {
  try {
    const lc  = mSh.getLastColumn();
    const hdr = mSh.getRange(1, 1, 1, lc).getValues()[0].map(v => String(v).trim());

    const staIdx = hdr.indexOf('_IntakeStatus');
    const skipCols = new Set(['_IntakeStatus', '_QueuedAt']);

    // ── Đọc minimal row hiện tại ─────────────────────────────────
    const minVals = mSh.getRange(minimalRowIdx, 1, 1, lc).getValues()[0];

    // ── Lấy row đầy đủ nhất trong các row mới ────────────────────
    // Ưu tiên row nào có PaymentsFolderURL
    const pfIdx = hdr.indexOf('PaymentsFolderURL');
    let bestNewRow = firstNewRow;
    let bestVals   = mSh.getRange(firstNewRow, 1, 1, lc).getValues()[0];

    for (let r = firstNewRow + 1; r <= lastNewRow; r++) {
      const rVals = mSh.getRange(r, 1, 1, lc).getValues()[0];
      const hasFolder = pfIdx >= 0 && String(rVals[pfIdx] || '').includes('drive.google.com');
      if (hasFolder) {
        bestNewRow = r;
        bestVals   = rVals;
        break;
      }
    }

    Logger.log('[_intake_mergeAndCleanup_] minimalRow=%s bestNewRow=%s', minimalRowIdx, bestNewRow);

    // ── Merge: copy bestVals → minimal row (bỏ qua skip cols) ────
    const merged = hdr.map(function(colName, idx) {
      if (skipCols.has(colName)) return minVals[idx];
      const newVal = bestVals[idx];
      // Chỉ ghi đè nếu giá trị mới có ý nghĩa
      if (newVal === '' || newVal === null || newVal === undefined) return minVals[idx];
      return newVal;
    });

    mSh.getRange(minimalRowIdx, 1, 1, lc).setValues([merged]);
    if (staIdx >= 0) mSh.getRange(minimalRowIdx, staIdx + 1).setValue('DONE');
    SpreadsheetApp.flush();

    Logger.log('[_intake_mergeAndCleanup_] ✅ Merged data into minimalRow=%s', minimalRowIdx);

    // ── Xóa tất cả các row mới (từ dưới lên để index không lệch) ─
    for (let r = lastNewRow; r >= firstNewRow; r--) {
      mSh.deleteRow(r);
      Logger.log('[_intake_mergeAndCleanup_] Deleted duplicate row=%s', r);
    }
    SpreadsheetApp.flush();

    Logger.log('[_intake_mergeAndCleanup_] ✅ Deleted %s duplicate row(s)', lastNewRow - firstNewRow + 1);

  } catch(e) {
    Logger.log('[_intake_mergeAndCleanup_] ERROR: ' + e.message);
  }
}


/**
 * Fallback: dùng khi onFormSubmit không thêm row mới
 * (trường hợp update existing row thay vì append)
 */
function _intake_patchMasterRow_(minimalRowIndex, customerName) {
  try {
    const ss  = SpreadsheetApp.getActive();
    const mSh = ss.getSheetByName(RP_MASTER_SHEET);
    const lc  = mSh.getLastColumn();
    const hdr = mSh.getRange(1, 1, 1, lc).getValues()[0].map(v => String(v).trim());

    const staIdx = hdr.indexOf('_IntakeStatus');
    if (staIdx >= 0) {
      mSh.getRange(minimalRowIndex, staIdx + 1).setValue('DONE');
    }

    // Thử tìm và merge nếu có row cùng tên với folder URL
    const nIdx  = hdr.indexOf('Customer Name');
    const pfIdx = hdr.indexOf('PaymentsFolderURL');
    if (nIdx < 0 || pfIdx < 0) return;

    const lr   = mSh.getLastRow();
    const vals = mSh.getRange(2, 1, lr - 1, lc).getValues();

    for (let i = vals.length - 1; i >= 0; i--) {
      const rowNum = i + 2;
      if (rowNum === minimalRowIndex) continue;
      const name   = String(vals[i][nIdx]  || '').trim().toLowerCase();
      const folder = String(vals[i][pfIdx] || '').trim();
      if (name === customerName.toLowerCase() && folder.includes('drive.google.com')) {
        // Dùng _intake_mergeAndCleanup_ để merge và xóa
        _intake_mergeAndCleanup_(mSh, minimalRowIndex, rowNum, rowNum);
        return;
      }
    }

    Logger.log('[_intake_patchMasterRow_] No duplicate found — only DONE marked');
  } catch(e) {
    Logger.log('[_intake_patchMasterRow_] ERROR: ' + e.message);
  }
}

/**
 * Cài time-trigger cho processIntakeQueue — chạy 1 lần duy nhất.
 * Gọi thủ công từ Apps Script Editor: chọn function này → Run
 *
 * An toàn: kiểm tra trùng trước khi tạo, không tạo duplicate.
 */
function installIntakeQueueTrigger() {
  const FUNC_NAME = 'processIntakeQueue';

  // Kiểm tra trigger đã tồn tại chưa
  const existing = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === FUNC_NAME);

  if (existing.length > 0) {
    Logger.log('[installIntakeQueueTrigger] Trigger đã tồn tại (%s cái) — bỏ qua.', existing.length);
    return { ok: true, created: false, message: 'Trigger already exists' };
  }

  // Tạo mới
  ScriptApp.newTrigger(FUNC_NAME)
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log('[installIntakeQueueTrigger] ✅ Trigger đã được tạo: %s — mỗi 1 phút.', FUNC_NAME);
  return { ok: true, created: true, message: 'Trigger created successfully' };
}


/**
 * Xóa trigger processIntakeQueue nếu cần tắt.
 * Gọi thủ công khi muốn disable background processing.
 */
function removeIntakeQueueTrigger() {
  const FUNC_NAME = 'processIntakeQueue';

  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === FUNC_NAME);

  if (triggers.length === 0) {
    Logger.log('[removeIntakeQueueTrigger] Không tìm thấy trigger nào cho %s.', FUNC_NAME);
    return { ok: true, removed: 0 };
  }

  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('[removeIntakeQueueTrigger] ✅ Đã xóa %s trigger.', triggers.length);
  return { ok: true, removed: triggers.length };
}


/**
 * Kiểm tra trạng thái trigger — xem đã cài chưa, chạy được không.
 * Gọi bất kỳ lúc nào để debug.
 */
function checkIntakeQueueTrigger() {
  const FUNC_NAME = 'processIntakeQueue';

  const all      = ScriptApp.getProjectTriggers();
  const matching = all.filter(t => t.getHandlerFunction() === FUNC_NAME);

  if (matching.length === 0) {
    Logger.log('[checkIntakeQueueTrigger] ❌ Chưa cài trigger cho %s.', FUNC_NAME);
    Logger.log('→ Chạy installIntakeQueueTrigger() để cài.');
    return { installed: false };
  }

  matching.forEach(t => {
    Logger.log('[checkIntakeQueueTrigger] ✅ Trigger: %s | Source: %s | Type: %s',
      t.getHandlerFunction(),
      t.getTriggerSource(),
      t.getEventType()
    );
  });

  // Kiểm tra thêm sheet buffer
  const ss    = SpreadsheetApp.getActive();
  const bufSh = ss.getSheetByName('_IntakeQueue');
  if (!bufSh) {
    Logger.log('[checkIntakeQueueTrigger] ⚠ Sheet _IntakeQueue chưa tồn tại — sẽ tự tạo khi submit đầu tiên.');
  } else {
    const pending = bufSh.getLastRow() < 2 ? 0 :
      bufSh.getRange(2, 2, bufSh.getLastRow() - 1, 1)
           .getValues()
           .filter(r => r[0] === 'PENDING').length;
    Logger.log('[checkIntakeQueueTrigger] Sheet _IntakeQueue: %s rows pending.', pending);
  }

  return { installed: true, count: matching.length };
}

/**
 * Đảm bảo PaymentsFolderURL tồn tại cho row.
 * Nếu chưa có → chủ động gọi rp_ensurePaymentsFolder_() tạo ngay.
 * Dùng khi user muốn submit payment trước khi trigger kịp chạy.
 *
 * @param {number} rowIndex
 * @returns {{ ok, paymentsFolderURL, created }}
 */
function ipad_ensureFolderReady(rowIndex) {
  try {
    rowIndex = Number(rowIndex);
    if (rowIndex < 2) return { ok: false, error: 'Invalid rowIndex' };

    const ss  = SpreadsheetApp.getActive();
    const mSh = ss.getSheetByName(RP_MASTER_SHEET);
    const lc  = mSh.getLastColumn();
    const hdr = mSh.getRange(1, 1, 1, lc).getValues()[0].map(v => String(v).trim());

    const pfIdx  = hdr.indexOf('PaymentsFolderURL');
    const apIdx  = hdr.indexOf('RootApptID') >= 0
                   ? hdr.indexOf('RootApptID') : hdr.indexOf('APPT_ID');
    const nIdx   = hdr.indexOf('Customer Name');
    const bIdx   = hdr.indexOf('Brand');

    if (pfIdx < 0) return { ok: false, error: 'No PaymentsFolderURL column' };

    const row = mSh.getRange(rowIndex, 1, 1, lc).getValues()[0];

    // ── Trường hợp 1: URL đã có → return ngay ─────────────────────
    const existingURL = String(row[pfIdx] || '').trim();
    if (existingURL && existingURL.includes('drive.google.com')) {
      Logger.log('[ipad_ensureFolderReady] Already has folder: %s', existingURL);
      return { ok: true, paymentsFolderURL: existingURL, created: false };
    }

    // ── Trường hợp 2: Chưa có → tìm từ rows cùng rootApptId ──────
    const rootApptId = apIdx >= 0 ? String(row[apIdx] || '').trim() : '';
    if (rootApptId && mSh.getLastRow() >= 2) {
      const vals = mSh.getRange(2, 1, mSh.getLastRow() - 1, lc).getValues();
      for (let i = 0; i < vals.length; i++) {
        if (i + 2 === rowIndex) continue;
        const sameRoot = apIdx >= 0
          && String(vals[i][apIdx] || '').trim() === rootApptId;
        if (!sameRoot) continue;
        const u = pfIdx >= 0 ? String(vals[i][pfIdx] || '').trim() : '';
        if (u && u.includes('drive.google.com')) {
          // Patch luôn vào row hiện tại
          mSh.getRange(rowIndex, pfIdx + 1).setValue(u);
          Logger.log('[ipad_ensureFolderReady] Inherited folder from row %s: %s', i + 2, u);
          return { ok: true, paymentsFolderURL: u, created: false };
        }
      }
    }

    // ── Trường hợp 3: Không có → tạo folder mới ngay lập tức ─────
    const customerName = nIdx >= 0 ? String(row[nIdx] || '').trim() : 'Walk-in';
    const brand        = bIdx >= 0 ? String(row[bIdx] || '').trim() : 'HPUSA';

    Logger.log('[ipad_ensureFolderReady] Creating folder now for row=%s name=%s', rowIndex, customerName);

    // Gọi hàm tạo folder sẵn có trong hệ thống
    const folderURL = rp_ensurePaymentsFolder_({
      customerName: customerName,
      brand:        brand,
      rootApptId:   rootApptId,
    });

    if (folderURL && folderURL.includes('drive.google.com')) {
      // Ghi URL vào cột
      mSh.getRange(rowIndex, pfIdx + 1).setValue(folderURL);
      // Đánh dấu intake status DONE luôn
      const staIdx = hdr.indexOf('_IntakeStatus');
      if (staIdx >= 0) mSh.getRange(rowIndex, staIdx + 1).setValue('DONE');

      Logger.log('[ipad_ensureFolderReady] Created folder: %s', folderURL);
      return { ok: true, paymentsFolderURL: folderURL, created: true };
    }

    // Nếu rp_ensurePaymentsFolder_ không tồn tại hoặc fail → return ok
    // nhưng không có URL (backend sẽ tự tạo khi rp_makeDocForPayment chạy)
    Logger.log('[ipad_ensureFolderReady] Could not create folder — will fallback to backend');
    return { ok: true, paymentsFolderURL: '', created: false };

  } catch(e) {
    Logger.log('[ipad_ensureFolderReady] ERROR: ' + e.message);
    // Không throw — trả ok:true để không block user
    return { ok: true, paymentsFolderURL: '', created: false };
  }
}

/**
 * Chạy full intake synchronously — gọi khi user click "Generate".
 * Thay thế cho _ensureAndLoadRecord khi masterRowIndex = 0.
 *
 * Flow:
 *   1. Tìm queue row PENDING
 *   2. Gọi onFormSubmit thật → tạo Master row
 *   3. Tìm rowIndex vừa tạo
 *   4. Update queue row = DONE
 *   5. Return { ok, masterRowIndex, paymentsFolderURL }
 */
function ipad_runIntakeNow(queueRow) {
  try {
    queueRow = Number(queueRow);
    if (!queueRow || queueRow < 2) return { ok: false, error: 'Invalid queueRow: ' + queueRow };

    const ss    = SpreadsheetApp.getActive();
    const bufSh = ss.getSheetByName('_IntakeQueue');
    if (!bufSh) return { ok: false, error: 'No queue sheet' };

    // ── Dùng lock để tránh race với processIntakeQueue ────────────
    const lock = LockService.getScriptLock();
    try { lock.waitLock(8000); } catch(_) {
      return { ok: false, error: 'Could not acquire lock — try again' };
    }

    let masterRowIndex = 0;
    let folderURL      = '';

    try {
      const rowData     = bufSh.getRange(queueRow, 1, 1, 8).getValues()[0];
      const status      = String(rowData[1] || '').trim();
      const resolvedUID = String(rowData[3] || '').trim();
      const payloadJson = String(rowData[5] || '{}');

      // Nếu đã DONE hoặc RUNNING (trigger đang xử lý) → đợi kết quả
      if (status === 'DONE') {
        const doneIdx = Number(rowData[2] || 0);
        if (doneIdx >= 2) {
          const mSh   = ss.getSheetByName(RP_MASTER_SHEET);
          const lc    = mSh.getLastColumn();
          const hdr   = mSh.getRange(1,1,1,lc).getValues()[0].map(v=>String(v).trim());
          const pfIdx = hdr.indexOf('PaymentsFolderURL');
          const r     = mSh.getRange(doneIdx,1,1,lc).getValues()[0];
          lock.releaseLock();
          return {
            ok: true, masterRowIndex: doneIdx, alreadyDone: true,
            paymentsFolderURL: pfIdx >= 0 ? String(r[pfIdx]||'').trim() : '',
          };
        }
      }

      if (status === 'RUNNING') {
        lock.releaseLock();
        return { ok: false, error: 'Already being processed — please wait a moment and retry' };
      }

      if (status !== 'PENDING') {
        lock.releaseLock();
        return { ok: false, error: 'Queue row status is: ' + status };
      }

      // ── Đánh dấu RUNNING NGAY — trong khi còn giữ lock ──────────
      // processIntakeQueue sẽ skip row này
      bufSh.getRange(queueRow, 2).setValue('RUNNING');
      SpreadsheetApp.flush();

    } finally {
      // Nhả lock sớm — chỉ cần lock khi đánh dấu RUNNING
      try { lock.releaseLock(); } catch(_) {}
    }

    // ── Từ đây chạy bình thường — không cần lock ─────────────────
    const rowData2    = bufSh.getRange(queueRow, 1, 1, 8).getValues()[0];
    const resolvedUID = String(rowData2[3] || '').trim();
    const payloadJson = String(rowData2[5] || '{}');
    const p           = JSON.parse(payloadJson);
    const tz          = Session.getScriptTimeZone();
    const now         = new Date();

    let visitDateStr = '';
    if (p.date) {
      try { visitDateStr = Utilities.formatDate(new Date(p.date+'T12:00:00'), tz, 'MM/dd/yyyy'); }
      catch(_) { visitDateStr = p.date || ''; }
    }
    let visitTimeStr = '';
    if (p.time) {
      const tp = p.time.split(':');
      let h = parseInt(tp[0], 10);
      const m2 = tp[1] || '00';
      const ap = h >= 12 ? 'PM' : 'AM';
      h = h % 12 || 12;
      visitTimeStr = h + ':' + m2 + ' ' + ap;
    }

    const diamonds = Array.isArray(p.diamond) ? p.diamond : (p.diamond?[p.diamond]:[]);
    const budgets  = Array.isArray(p.budget)  ? p.budget  : (p.budget ?[p.budget] :[]);
    const sources  = Array.isArray(p.source)  ? p.source  : (p.source ?[p.source] :[]);

    const namedValues = {
      'Timestamp':                 [Utilities.formatDate(now, tz, 'M/d/yyyy H:mm:ss')],
      'Company':                   [p.company   || rowData2[4] || ''],
      'Customer Name':             [p.name      || ''],
      'Phone':                     [p.phone     || ''],
      'Email':                     [p.email     || ''],
      'Visit Type':                [p.visitType || 'Walk-In'],
      'Visit Date':                [visitDateStr],
      'Visit Time':                [visitTimeStr],
      'Location':                  [p.location  || 'In Store'],
      'Diamond Type':              [diamonds.join(', ')],
      'Budget Range':              [budgets.join(', ')],
      'Source':                    [sources.join(', ')],
      'Style Notes':               [p.notes || ''],
      'Admin: Calendly Event UID': [resolvedUID || ''],
    };

    const inboxSh = ss.getSheetByName('02_Form_Inbox');
    if (!inboxSh) throw new Error('02_Form_Inbox not found');

    const headers = inboxSh.getRange(1,1,1,inboxSh.getLastColumn()).getValues()[0];
    const rowArr  = headers.map(h => { const v=namedValues[h]; return v?v[0]:''; });
    inboxSh.appendRow(rowArr);
    const inboxRow = inboxSh.getLastRow();

    const mSh      = ss.getSheetByName(RP_MASTER_SHEET);
    const lrBefore = mSh ? mSh.getLastRow() : 0;

    onFormSubmit({
      namedValues: namedValues,
      range:       inboxSh.getRange(inboxRow, 1, 1, headers.length),
      values:      rowArr,
    });
    SpreadsheetApp.flush();

    // ── Tìm masterRowIndex ────────────────────────────────────────
    const lrAfter = mSh ? mSh.getLastRow() : 0;
    if (lrAfter > lrBefore) {
      masterRowIndex = lrBefore + 1;
    } else {
      const lc2  = mSh.getLastColumn();
      const hdr2 = mSh.getRange(1,1,1,lc2).getValues()[0].map(v=>String(v).trim());
      const nIdx = hdr2.indexOf('Customer Name');
      if (nIdx >= 0 && mSh.getLastRow() >= 2) {
        const vals2  = mSh.getRange(2,1,mSh.getLastRow()-1,lc2).getValues();
        const target = String(p.name||'').trim().toLowerCase();
        for (let i = vals2.length-1; i >= 0; i--) {
          if (String(vals2[i][nIdx]||'').trim().toLowerCase() === target) {
            masterRowIndex = i + 2; break;
          }
        }
      }
    }

    // ── PaymentsFolderURL ─────────────────────────────────────────
    if (masterRowIndex >= 2) {
      const lc3   = mSh.getLastColumn();
      const hdr3  = mSh.getRange(1,1,1,lc3).getValues()[0].map(v=>String(v).trim());
      const pfIdx = hdr3.indexOf('PaymentsFolderURL');
      if (pfIdx >= 0)
        folderURL = String(mSh.getRange(masterRowIndex,1,1,lc3).getValues()[0][pfIdx]||'').trim();
    }

    // ── Đánh dấu DONE ─────────────────────────────────────────────
    bufSh.getRange(queueRow, 2).setValue('DONE');
    bufSh.getRange(queueRow, 3).setValue(masterRowIndex);
    bufSh.getRange(queueRow, 7).setValue(new Date());

    Logger.log('[ipad_runIntakeNow] DONE queueRow=%s masterRow=%s', queueRow, masterRowIndex);
    return { ok: true, masterRowIndex, paymentsFolderURL: folderURL };

  } catch(e) {
    // Nếu lỗi → đặt lại PENDING để trigger retry được
    try {
      const ss2    = SpreadsheetApp.getActive();
      const bufSh2 = ss2.getSheetByName('_IntakeQueue');
      if (bufSh2) {
        bufSh2.getRange(queueRow, 2).setValue('PENDING');
        bufSh2.getRange(queueRow, 8).setValue(e.message);
      }
    } catch(_) {}
    Logger.log('[ipad_runIntakeNow] ERROR: ' + e.message);
    return { ok: false, error: e.message };
  }
}


/**
 * Poll — client gọi để check queue đã DONE chưa.
 */
function ipad_checkQueueStatus(queueRow) {
  try {
    queueRow = Number(queueRow);
    const ss    = SpreadsheetApp.getActive();
    const bufSh = ss.getSheetByName('_IntakeQueue');
    if (!bufSh || queueRow < 2) return { status: 'UNKNOWN' };

    const row = bufSh.getRange(queueRow, 1, 1, 8).getValues()[0];
    const status         = String(row[1] || '').trim();
    const masterRowIndex = Number(row[2] || 0);

    if (status !== 'DONE' || masterRowIndex < 2) {
      return { status };
    }

    // DONE → lấy folder URL
    const mSh = ss.getSheetByName(RP_MASTER_SHEET);
    const lc  = mSh.getLastColumn();
    const hdr = mSh.getRange(1,1,1,lc).getValues()[0].map(v=>String(v).trim());
    const pfIdx = hdr.indexOf('PaymentsFolderURL');
    let folderURL = '';
    if (pfIdx >= 0 && masterRowIndex >= 2) {
      folderURL = String(mSh.getRange(masterRowIndex,1,1,lc).getValues()[0][pfIdx]||'').trim();
    }

    return { status: 'DONE', masterRowIndex, paymentsFolderURL: folderURL };
  } catch(e) {
    return { status: 'ERROR', error: e.message };
  }
}