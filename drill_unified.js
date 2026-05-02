// /** ===========================================================================
//  *  PROJECT #18 — Unified Drill-Down Table (sheet: Drill_Unified)
//  *
//  *  Business Requirement:
//  *    Thay vì nhiều bảng nhỏ rời rạc trong Drill_KPI, tạo 1 bảng duy nhất
//  *    chứa toàn bộ thông tin appointment + deposit + order trong cùng 1 chỗ.
//  *
//  *  Columns:
//  *    Route Appointment ID | Customer Name | Date of Appointment |
//  *    Appointment Type | Appointment Status | No-Show? | Completed? |
//  *    Deposit Amount | Deposit Made? | Order Total | Sales Rep | Brand
//  *
//  *  Filter: Quick Picker trên 00_Dashboard (B1 preset, B2/D2 date, H1 brand, H2 rep)
//  *
//  *  Cách tích hợp:
//  *    1. Thêm vào runOnceToBuildAll():  safeCall_(buildUnifiedDrillDown_);
//  *    2. Hàm tự chạy theo hourly trigger qua refreshDashboardHourly()
//  * =========================================================================== */

// // ─── Sheet name ──────────────────────────────────────────────────────────────
// const SH_UNIFIED = 'Drill_Unified';

// // ─── Styling ─────────────────────────────────────────────────────────────────
// const UNI_HEADER_BG   = '#1565C0';   // Blue đậm
// const UNI_HEADER_FG   = '#FFFFFF';
// const UNI_SUBHEAD_BG  = '#E3F2FD';   // Blue nhạt cho filter info
// const UNI_ODD_BG      = '#FFFFFF';
// const UNI_EVEN_BG     = '#F5F9FF';
// const UNI_YES_BG      = '#C8E6C9';   // Xanh lá — Yes / Completed / Deposit made
// const UNI_YES_FG      = '#1B5E20';
// const UNI_NO_BG       = '#FFCDD2';   // Đỏ nhạt — No-Show
// const UNI_NO_FG       = '#B71C1C';
// const UNI_WARN_BG     = '#FFF9C4';   // Vàng — Cancelled / Rescheduled
// const UNI_WARN_FG     = '#F57F17';
// const UNI_SCHED_BG    = '#E8EAF6';   // Tím nhạt — Scheduled
// const UNI_SCHED_FG    = '#283593';
// const UNI_BORDER      = '#BBDEFB';


// // =============================================================================
// // ENTRY POINT
// // =============================================================================

// /**
//  * Main function — build Drill_Unified sheet.
//  *
//  * Thêm dòng này vào runOnceToBuildAll() trong dashboard.gs:
//  *   safeCall_(buildUnifiedDrillDown_);
//  */
// function buildUnifiedDrillDown_() {
//   const ss   = SpreadsheetApp.getActive();
//   const dash = ss.getSheetByName(SH_DASH);
//   if (!dash) throw new Error('Missing sheet: ' + SH_DASH);

//   // ── 1. Đọc filters từ Quick Picker ────────────────────────────────────────
//   const start       = asDate_(dash.getRange(CELL_DATE_S).getValue());
//   const end         = asDate_(dash.getRange(CELL_DATE_E).getValue());
//   const brandFilter = String(dash.getRange(CELL_BRAND).getValue() || '').trim();
//   const repFilter   = String(dash.getRange(CELL_REP).getValue()   || '').trim();

//   if (!(start instanceof Date) || !(end instanceof Date)) {
//     Logger.log('[Unified] Invalid date range — skipping');
//     return;
//   }

//   // ── 2. Load dữ liệu ───────────────────────────────────────────────────────
//   const masterSh  = ss.getSheetByName(SH_MASTER);
//   const metricsSh = ss.getSheetByName(SH_METRICS);
//   if (!masterSh || !metricsSh) {
//     Logger.log('[Unified] Missing Master or Metrics sheet — skipping');
//     return;
//   }

//   const masterVals  = masterSh.getDataRange().getValues();
//   const masterH     = masterVals.shift().map(x => String(x || '').trim());

//   const metricsVals = metricsSh.getDataRange().getValues();
//   const metricsH    = metricsVals.shift().map(x => String(x || '').trim());
//   const xi          = makeIdx_(metricsH);

//   // Payment map: root → { d: Date, amt: number, so: string }
//   let firstPayMap = new Map();
//   try {
//     firstPayMap = fetchFirstPaymentMapFromPayments_();
//   } catch (e) {
//     Logger.log('[Unified] Could not load payment map: ' + e.message);
//   }

//   // ── 3. Build dataset ───────────────────────────────────────────────────────
//   const rows = _buildUnifiedRows_(
//     start, end, brandFilter, repFilter,
//     masterVals, masterH,
//     metricsVals, xi,
//     firstPayMap
//   );

//   // ── 4. Ghi sheet ──────────────────────────────────────────────────────────
//   _writeUnifiedSheet_(ss, rows, start, end, brandFilter, repFilter);

//   // ── 5. Thêm drill link trên 00_Dashboard ──────────────────────────────────
//   _addUnifiedLinkToDashboard_(dash, ss);

//   Logger.log('[Unified] Done. Rows written: ' + rows.length);
// }


// // =============================================================================
// // BUILD ROWS — join Master + Metrics + Payments
// // =============================================================================

// function _buildUnifiedRows_(
//   start, end, brandFilter, repFilter,
//   masterVals, masterH,
//   metricsVals, xi,
//   firstPayMap
// ) {
//   // ── Master column indexes ──────────────────────────────────────────────────
//   const iRoot   = findCol_(masterH, ['RootApptID','Root Appt ID','ROOT','Root_ID']);
//   const iBrand  = findCol_(masterH, ['Brand']);
//   const iRep    = findCol_(masterH, ['Assigned Rep','AssignedRep','Rep','Sales Rep']);
//   const iVType  = findCol_(masterH, ['Visit Type','VisitType','Type']);
//   const iVDate  = findCol_(masterH, ['Visit Date','Visit_Date','Appt Date','Appointment Date']);
//   const iStatus = findCol_(masterH, ['Status','Appt Status','Appointment Status'], false);
//   const iCust   = findCol_(masterH, ['Customer Name','Customer','Client Name'],    false);

//   // ── Helpers ────────────────────────────────────────────────────────────────
//   const inWin   = (d) => { const t = asDate_(d); return t && t >= start && t <= end; };
//   const matchBR = (r) =>
//     (!brandFilter || String(r[iBrand] || '').trim() === brandFilter) &&
//     (!repFilter   || String(r[iRep]   || '').trim() === repFilter);

//   // ── Metrics lookup: root → metrics row ────────────────────────────────────
//   const metricsMap = new Map();
//   metricsVals.forEach(r => {
//     const root = String(r[xi['RootApptID']] || '').trim();
//     if (root && !metricsMap.has(root)) metricsMap.set(root, r);
//   });

//   // ── Build output ──────────────────────────────────────────────────────────
//   const out = [];

//   masterVals.forEach(r => {
//     // Filter: phải trong date window và match Brand/Rep
//     if (!inWin(r[iVDate]) || !matchBR(r)) return;

//     const root     = String(r[iRoot]   || '').trim();
//     const brand    = String(r[iBrand]  || '').trim();
//     const rep      = String(r[iRep]    || '').trim();
//     const vtype    = String(r[iVType]  || '').trim();
//     const vdate    = r[iVDate];
//     const status   = iStatus >= 0 ? String(r[iStatus] || '').trim() : '';
//     const custName = iCust   >= 0 ? String(r[iCust]   || '').trim() : '';

//     // ── Derived status flags ─────────────────────────────────────────────────
//     const statusLow  = status.toLowerCase();
//     const isNoShow   = /no[-\s]?show/i.test(status);
//     const isCompleted= /completed?/i.test(status);

//     // ── Order Total từ Metrics ────────────────────────────────────────────────
//     let orderTotal = '';
//     const mRow = metricsMap.get(root);
//     if (mRow) {
//       const ot = mRow[xi['Order Total']];
//       if (ot !== '' && ot !== null && isFinite(Number(ot)) && Number(ot) > 0) {
//         orderTotal = Number(ot);
//       }
//     }

//     // ── Deposit Amount từ Payment Map ─────────────────────────────────────────
//     let depositAmt  = '';
//     let depositMade = 'No';
//     const fp = firstPayMap.get(root);
//     if (fp && fp.amt && Number(fp.amt) > 0) {
//       depositAmt  = Number(fp.amt);
//       depositMade = 'Yes';
//     }

//     out.push([
//       root,         // A: Route Appointment ID
//       custName,     // B: Customer Name
//       vdate,        // C: Date of Appointment
//       vtype,        // D: Appointment Type
//       status,       // E: Appointment Status
//       isNoShow ? 'Yes' : 'No',        // F: No-Show?
//       isCompleted ? 'Yes' : 'No',     // G: Completed?
//       depositAmt,   // H: Deposit Amount
//       depositMade,  // I: Deposit Made?
//       orderTotal,   // J: Order Total
//       rep,          // K: Sales Rep
//       brand         // L: Brand
//     ]);
//   });

//   // Sắp xếp: Date tăng dần → Rep → Customer
//   out.sort((a, b) => {
//     const da = asDate_(a[2]), db = asDate_(b[2]);
//     if (da && db && da.getTime() !== db.getTime()) return da - db;
//     const repCmp = String(a[10]).localeCompare(String(b[10]));
//     if (repCmp !== 0) return repCmp;
//     return String(a[1]).localeCompare(String(b[1]));
//   });

//   return out;
// }


// // =============================================================================
// // WRITE SHEET — ghi dữ liệu + format toàn bộ
// // =============================================================================

// function _writeUnifiedSheet_(ss, rows, start, end, brandFilter, repFilter) {
//   const tz      = Session.getScriptTimeZone() || 'GMT';
//   const fmtDate = (d) => (d instanceof Date)
//     ? Utilities.formatDate(d, tz, 'yyyy-MM-dd')
//     : String(d || '');

//   // ── Lấy hoặc tạo sheet ────────────────────────────────────────────────────
//   // Xoa va tao lai sheet de tranh moi loi merge/filter/banding
//   let sh = ss.getSheetByName(SH_UNIFIED);
//   if (sh) {
//     // Ghi nho vi tri de insert lai cung cho
//     const sheetIndex = sh.getIndex();
//     ss.deleteSheet(sh);
//     sh = ss.insertSheet(SH_UNIFIED, sheetIndex - 1);
//   } else {
//     sh = ss.insertSheet(SH_UNIFIED);
//   }
//   sh.setHiddenGridlines(false);

//   // ── Column headers ─────────────────────────────────────────────────────────
//   const HEADERS = [
//     'Route Appointment ID',
//     'Customer Name',
//     'Date of Appointment',
//     'Appointment Type',
//     'Appointment Status',
//     'No-Show?',
//     'Completed?',
//     'Deposit Amount',
//     'Deposit Made?',
//     'Order Total',
//     'Sales Rep',
//     'Brand'
//   ];
//   const NUM_COLS = HEADERS.length; // 12

//   // ── Row positions ──────────────────────────────────────────────────────────
//   const TITLE_ROW  = 1;
//   const FILTER_ROW = 2;
//   const STATS_ROW  = 3;
//   const HEADER_ROW = 5;
//   const DATA_START = 6;

//   // ── Title ──────────────────────────────────────────────────────────────────
//   sh.getRange(TITLE_ROW, 1, 1, NUM_COLS).merge()
//     .setValue('📊  Unified Drill-Down — Appointments × Deposits × Orders')
//     .setBackground(UNI_HEADER_BG)
//     .setFontColor(UNI_HEADER_FG)
//     .setFontWeight('bold')
//     .setFontSize(14)
//     .setHorizontalAlignment('left')
//     .setVerticalAlignment('middle');
//   sh.setRowHeight(TITLE_ROW, 36);

//   // ── Filter info bar ────────────────────────────────────────────────────────
//   const periodLabel = fmtDate(start) + '  →  ' + fmtDate(end);
//   const filterText  =
//     '📅 Period: ' + periodLabel +
//     (brandFilter ? '     🏷 Brand: ' + brandFilter : '     🏷 Brand: All') +
//     (repFilter   ? '     👤 Rep: ' + repFilter      : '     👤 Rep: All');

//   sh.getRange(FILTER_ROW, 1, 1, NUM_COLS).merge()
//     .setValue(filterText)
//     .setBackground(UNI_SUBHEAD_BG)
//     .setFontColor('#0D47A1')
//     .setFontSize(10)
//     .setFontWeight('normal')
//     .setHorizontalAlignment('left')
//     .setVerticalAlignment('middle');
//   sh.setRowHeight(FILTER_ROW, 24);

//   // ── Summary stats ──────────────────────────────────────────────────────────
//   const totalAppts     = rows.length;
//   const noShowCount    = rows.filter(r => r[5] === 'Yes').length;
//   const completedCount = rows.filter(r => r[6] === 'Yes').length;
//   const depositCount   = rows.filter(r => r[8] === 'Yes').length;
//   const totalDeposit   = rows.reduce((s, r) => s + (Number(r[7]) || 0), 0);
//   const totalOrder     = rows.reduce((s, r) => s + (Number(r[9]) || 0), 0);
//   const noShowRate     = totalAppts ? ((noShowCount / totalAppts) * 100).toFixed(1) + '%' : 'N/A';

//   const statsText =
//     'Total: ' + totalAppts + ' appts' +
//     '     ✅ Completed: ' + completedCount +
//     '     ❌ No-Show: ' + noShowCount + ' (' + noShowRate + ')' +
//     '     💰 Deposits: ' + depositCount +
//     '     Deposit $: $' + Math.round(totalDeposit).toLocaleString('en-US') +
//     '     Order $: $' + Math.round(totalOrder).toLocaleString('en-US');

//   sh.getRange(STATS_ROW, 1, 1, NUM_COLS).merge()
//     .setValue(statsText)
//     .setBackground('#FFFFFF')
//     .setFontColor('#37474F')
//     .setFontSize(10)
//     .setFontStyle('italic')
//     .setHorizontalAlignment('left')
//     .setVerticalAlignment('middle');
//   sh.setRowHeight(STATS_ROW, 22);

//   // Row 4: khoảng trắng
//   sh.setRowHeight(4, 8);

//   // ── Header row ─────────────────────────────────────────────────────────────
//   const hdrRange = sh.getRange(HEADER_ROW, 1, 1, NUM_COLS);
//   hdrRange.setValues([HEADERS])
//     .setBackground(UNI_HEADER_BG)
//     .setFontColor(UNI_HEADER_FG)
//     .setFontWeight('bold')
//     .setFontSize(10)
//     .setHorizontalAlignment('center')
//     .setVerticalAlignment('middle')
//     .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
//   sh.setRowHeight(HEADER_ROW, 30);
//   sh.setFrozenRows(HEADER_ROW);
//   // Note: setFrozenColumns removed — conflicts with merged title/filter cells

//   // ── Data ───────────────────────────────────────────────────────────────────
//   if (rows.length === 0) {
//     sh.getRange(DATA_START, 1, 1, NUM_COLS).merge()
//       .setValue('No data found for the selected filters.')
//       .setFontStyle('italic')
//       .setFontColor('#888888')
//       .setHorizontalAlignment('center');
//   } else {
//     // Chuẩn bị display values (date format)
//     const displayRows = rows.map(r => {
//       const d = asDate_(r[2]);
//       return [
//         r[0],                                            // RootApptID
//         r[1],                                            // Customer
//         d ? fmtDate(d) : (r[2] ? String(r[2]) : ''),    // Date formatted
//         r[3],                                            // Type
//         r[4],                                            // Status
//         r[5],                                            // No-Show?
//         r[6],                                            // Completed?
//         r[7],                                            // Deposit $
//         r[8],                                            // Deposit Made?
//         r[9],                                            // Order Total
//         r[10],                                           // Sales Rep
//         r[11]                                            // Brand
//       ];
//     });

//     const dataRange = sh.getRange(DATA_START, 1, rows.length, NUM_COLS);
//     dataRange.setValues(displayRows);

//     // ── Apply row-by-row formatting ──────────────────────────────────────────
//     // Batch để tránh quota
//     const bgArr    = [];
//     const fgArr    = [];
//     const fwArr    = [];
//     const halignH  = [];   // Horizontal alignment per row
//     const halignD  = [];
//     const halignN  = [];
//     const halignC  = [];
//     const halignDM = [];
//     const halignOT = [];

//     // Pre-build color arrays
//     const statusColBg  = [];  // col E (index 4)
//     const statusColFg  = [];
//     const noShowColBg  = [];  // col F (index 5)
//     const noShowColFg  = [];
//     const compColBg    = [];  // col G (index 6)
//     const compColFg    = [];
//     const depMadeColBg = [];  // col I (index 8)
//     const depMadeColFg = [];

//     rows.forEach((r, i) => {
//       const isEven   = (i % 2 === 0);
//       const baseBg   = isEven ? UNI_ODD_BG : UNI_EVEN_BG;

//       // Row background (base)
//       const rowBg = new Array(NUM_COLS).fill(baseBg);
//       const rowFg = new Array(NUM_COLS).fill('#212121');
//       const rowFw = new Array(NUM_COLS).fill('normal');

//       bgArr.push(rowBg);
//       fgArr.push(rowFg);
//       fwArr.push(rowFw);

//       // Status column (E = col 5 = index 4)
//       const statusLow = String(r[4] || '').toLowerCase();
//       let sBg = baseBg, sFg = '#212121';
//       if (/no[-\s]?show/.test(statusLow))         { sBg = UNI_NO_BG;   sFg = UNI_NO_FG;   }
//       else if (/completed?/.test(statusLow))       { sBg = UNI_YES_BG;  sFg = UNI_YES_FG;  }
//       else if (/cancel/.test(statusLow))           { sBg = UNI_WARN_BG; sFg = UNI_WARN_FG; }
//       else if (/scheduled|rescheduled/.test(statusLow)) { sBg = UNI_SCHED_BG; sFg = UNI_SCHED_FG; }
//       statusColBg.push([sBg]);
//       statusColFg.push([sFg]);

//       // No-Show? column (F = col 6 = index 5)
//       noShowColBg.push([r[5] === 'Yes' ? UNI_NO_BG   : baseBg]);
//       noShowColFg.push([r[5] === 'Yes' ? UNI_NO_FG   : '#212121']);

//       // Completed? column (G = col 7 = index 6)
//       compColBg.push([r[6] === 'Yes' ? UNI_YES_BG  : baseBg]);
//       compColFg.push([r[6] === 'Yes' ? UNI_YES_FG  : '#212121']);

//       // Deposit Made? column (I = col 9 = index 8)
//       depMadeColBg.push([r[8] === 'Yes' ? UNI_YES_BG : baseBg]);
//       depMadeColFg.push([r[8] === 'Yes' ? UNI_YES_FG : '#888888']);
//     });

//     // Apply base colors to whole range
//     dataRange.setBackgrounds(bgArr).setFontColors(fgArr).setFontWeights(fwArr);

//     // Apply status-specific colors column by column
//     const n = rows.length;
//     sh.getRange(DATA_START, 5, n, 1).setBackgrounds(statusColBg).setFontColors(statusColFg).setFontWeights(statusColFg.map(c => c[0] !== '#212121' ? ['bold'] : ['normal']));
//     sh.getRange(DATA_START, 6, n, 1).setBackgrounds(noShowColBg).setFontColors(noShowColFg);
//     sh.getRange(DATA_START, 7, n, 1).setBackgrounds(compColBg).setFontColors(compColFg);
//     sh.getRange(DATA_START, 9, n, 1).setBackgrounds(depMadeColBg).setFontColors(depMadeColFg);

//     // Number formats
//     sh.getRange(DATA_START, 8, n, 1).setNumberFormat('$#,##0');   // Deposit Amount
//     sh.getRange(DATA_START, 10, n, 1).setNumberFormat('$#,##0');  // Order Total

//     // Alignment
//     sh.getRange(DATA_START, 1, n, NUM_COLS).setHorizontalAlignment('left').setVerticalAlignment('middle');
//     sh.getRange(DATA_START, 3, n, 1).setHorizontalAlignment('center'); // Date
//     sh.getRange(DATA_START, 6, n, 2).setHorizontalAlignment('center'); // No-Show / Completed
//     sh.getRange(DATA_START, 8, n, 3).setHorizontalAlignment('right');  // Deposit/Made/Order
//     sh.getRange(DATA_START, 12, n, 1).setHorizontalAlignment('center'); // Brand

//     // Row heights
//     sh.setRowHeightsForced(DATA_START, n, 22);

//     // Border toàn bảng
//     sh.getRange(HEADER_ROW, 1, n + 1, NUM_COLS)
//       .setBorder(true, true, true, true, true, true,
//         UNI_BORDER, SpreadsheetApp.BorderStyle.SOLID);

//     // Filter
//     sh.getRange(HEADER_ROW, 1, n + 1, NUM_COLS).createFilter();
//   }

//   // ── Column widths ──────────────────────────────────────────────────────────
//   sh.setColumnWidth(1,  160);  // RootApptID
//   sh.setColumnWidth(2,  160);  // Customer Name
//   sh.setColumnWidth(3,  130);  // Date
//   sh.setColumnWidth(4,  140);  // Appointment Type
//   sh.setColumnWidth(5,  130);  // Status
//   sh.setColumnWidth(6,   80);  // No-Show?
//   sh.setColumnWidth(7,   90);  // Completed?
//   sh.setColumnWidth(8,  110);  // Deposit Amount
//   sh.setColumnWidth(9,  110);  // Deposit Made?
//   sh.setColumnWidth(10, 110);  // Order Total
//   sh.setColumnWidth(11, 140);  // Sales Rep
//   sh.setColumnWidth(12,  90);  // Brand

//   SpreadsheetApp.flush();
// }




// // =============================================================================
// // DASHBOARD LINK — đặt tại A3 (row trống, không có merged cells)
// // =============================================================================

// function _addUnifiedLinkToDashboard_(dash, ss) {
//   try {
//     const unifiedSh = ss.getSheetByName(SH_UNIFIED);
//     if (!unifiedSh) return;
//     const gid = unifiedSh.getSheetId();

//     // Đặt link bên dưới các charts — tìm last row có content rồi cộng thêm 2
//     const lastRow = dash.getLastRow();
//     const linkRow = lastRow + 2;

//     const linkCell = dash.getRange(linkRow, 1);
//     linkCell
//       .setFormula('=HYPERLINK("#gid=' + gid + '","🔗  Open Unified Drill-Down Table →")')
//       .setBackground('#1565C0')
//       .setFontColor('#FFFFFF')
//       .setFontWeight('bold')
//       .setFontSize(11)
//       .setHorizontalAlignment('left')
//       .setVerticalAlignment('middle');

//     dash.setRowHeight(linkRow, 30);
//     dash.setColumnWidth(1, 280);

//     // Label nhỏ bên cạnh
//     dash.getRange(linkRow, 2)
//       .setValue('← Click to view full appointment + deposit + order details')
//       .setFontColor('#555555')
//       .setFontSize(10)
//       .setFontStyle('italic')
//       .setVerticalAlignment('middle');

//     Logger.log('[Unified] Dashboard link added at row ' + linkRow);

//     Logger.log('[Unified] Dashboard link added at A3');
//   } catch (e) {
//     Logger.log('[Unified] Could not add dashboard link: ' + e.message);
//   }
// }

// // =============================================================================
// // INTEGRATION — Patch vào runOnceToBuildAll()
// // =============================================================================
// //
// //  Trong file dashboard.gs, tìm function runOnceToBuildAll() và thêm dòng:
// //
// //    function runOnceToBuildAll() {
// //      safeCall_(ensureDashboardLayout_);
// //      safeCall_(buildMetricsView_);
// //      safeCall_(writeDashboard_);
// //      safeCall_(snapshotKpisForHistory_);
// //      safeCall_(buildUnifiedDrillDown_);   // ← THÊM DÒNG NÀY
// //    }
// //
// // =============================================================================


// // =============================================================================
// // MANUAL TRIGGER — có thể chạy thẳng từ Apps Script editor để test
// // =============================================================================

// function runBuildUnifiedDrillDown() {
//   buildUnifiedDrillDown_();
//   SpreadsheetApp.getUi().alert(
//     '✅ Unified Drill-Down đã được tạo!',
//     'Sheet "' + SH_UNIFIED + '" đã được cập nhật với dữ liệu mới nhất.\n\n' +
//     'Vui lòng mở tab "Drill_Unified" để xem.',
//     SpreadsheetApp.getUi().ButtonSet.OK
//   );
// }



/** ===========================================================================
 *  PROJECT #18 — Unified Drill-Down Table (sheet: Drill_Unified)
 *
 *  Business Requirement:
 *    Thay vì nhiều bảng nhỏ rời rạc trong Drill_KPI, tạo 1 bảng duy nhất
 *    chứa toàn bộ thông tin appointment + deposit + order trong cùng 1 chỗ.
 *
 *  Columns:
 *    Route Appointment ID | Customer Name | Date of Appointment |
 *    Appointment Type | Appointment Status | No-Show? | Completed? |
 *    Deposit Amount | Deposit Made? | Order Total | Sales Rep | Brand
 *
 *  Filter: Quick Picker trên 00_Dashboard (B1 preset, B2/D2 date, H1 brand, H2 rep)
 *
 *  Cách tích hợp:
 *    1. Thêm vào runOnceToBuildAll():  safeCall_(buildUnifiedDrillDown_);
 *    2. Hàm tự chạy theo hourly trigger qua refreshDashboardHourly()
 * =========================================================================== */

// ─── Sheet name ──────────────────────────────────────────────────────────────
const SH_UNIFIED = 'Drill_Unified';

// ─── Styling ─────────────────────────────────────────────────────────────────
const UNI_HEADER_BG   = '#1565C0';   // Blue đậm
const UNI_HEADER_FG   = '#FFFFFF';
const UNI_SUBHEAD_BG  = '#E3F2FD';   // Blue nhạt cho filter info
const UNI_ODD_BG      = '#FFFFFF';
const UNI_EVEN_BG     = '#F5F9FF';
const UNI_YES_BG      = '#C8E6C9';   // Xanh lá — Yes / Completed / Deposit made
const UNI_YES_FG      = '#1B5E20';
const UNI_NO_BG       = '#FFCDD2';   // Đỏ nhạt — No-Show
const UNI_NO_FG       = '#B71C1C';
const UNI_WARN_BG     = '#FFF9C4';   // Vàng — Cancelled / Rescheduled
const UNI_WARN_FG     = '#F57F17';
const UNI_SCHED_BG    = '#E8EAF6';   // Tím nhạt — Scheduled
const UNI_SCHED_FG    = '#283593';
const UNI_BORDER      = '#BBDEFB';


// =============================================================================
// ENTRY POINT
// =============================================================================

/**
 * Main function — build Drill_Unified sheet.
 *
 * Thêm dòng này vào runOnceToBuildAll() trong dashboard.gs:
 *   safeCall_(buildUnifiedDrillDown_);
 */
function buildUnifiedDrillDown_() {
  const ss   = SpreadsheetApp.getActive();
  const dash = ss.getSheetByName(SH_DASH);
  if (!dash) throw new Error('Missing sheet: ' + SH_DASH);

  // ── 1. Đọc filters từ Quick Picker ────────────────────────────────────────
  const start       = asDate_(dash.getRange(CELL_DATE_S).getValue());
  const end         = asDate_(dash.getRange(CELL_DATE_E).getValue());
  const brandFilter = String(dash.getRange(CELL_BRAND).getValue() || '').trim();
  const repFilter   = String(dash.getRange(CELL_REP).getValue()   || '').trim();

  if (!(start instanceof Date) || !(end instanceof Date)) {
    Logger.log('[Unified] Invalid date range — skipping');
    return;
  }

  // ── 2. Load dữ liệu ───────────────────────────────────────────────────────
  const masterSh  = ss.getSheetByName(SH_MASTER);
  const metricsSh = ss.getSheetByName(SH_METRICS);
  if (!masterSh || !metricsSh) {
    Logger.log('[Unified] Missing Master or Metrics sheet — skipping');
    return;
  }

  const masterVals  = masterSh.getDataRange().getValues();
  const masterH     = masterVals.shift().map(x => String(x || '').trim());

  const metricsVals = metricsSh.getDataRange().getValues();
  const metricsH    = metricsVals.shift().map(x => String(x || '').trim());
  const xi          = makeIdx_(metricsH);

  // Payment map: root → { d: Date, amt: number, so: string }
  let firstPayMap = new Map();
  try {
    firstPayMap = fetchFirstPaymentMapFromPayments_();
  } catch (e) {
    Logger.log('[Unified] Could not load payment map: ' + e.message);
  }

  // ── 3. Build dataset ───────────────────────────────────────────────────────
  const rows = _buildUnifiedRows_(
    start, end, brandFilter, repFilter,
    masterVals, masterH,
    metricsVals, xi,
    firstPayMap
  );

  // ── 4. Ghi sheet ──────────────────────────────────────────────────────────
  _writeUnifiedSheet_(ss, rows, start, end, brandFilter, repFilter);



  Logger.log('[Unified] Done. Rows written: ' + rows.length);
}


// =============================================================================
// BUILD ROWS — join Master + Metrics + Payments
// =============================================================================

function _buildUnifiedRows_(
  start, end, brandFilter, repFilter,
  masterVals, masterH,
  metricsVals, xi,
  firstPayMap
) {
  // ── Master column indexes ──────────────────────────────────────────────────
  const iRoot   = findCol_(masterH, ['RootApptID','Root Appt ID','ROOT','Root_ID']);
  const iBrand  = findCol_(masterH, ['Brand']);
  const iRep    = findCol_(masterH, ['Assigned Rep','AssignedRep','Rep','Sales Rep']);
  const iVType  = findCol_(masterH, ['Visit Type','VisitType','Type']);
  const iVDate  = findCol_(masterH, ['Visit Date','Visit_Date','Appt Date','Appointment Date']);
  const iStatus = findCol_(masterH, ['Status','Appt Status','Appointment Status'], false);
  const iCust   = findCol_(masterH, ['Customer Name','Customer','Client Name'],    false);

  // ── Helpers ────────────────────────────────────────────────────────────────
  const inWin   = (d) => { const t = asDate_(d); return t && t >= start && t <= end; };
  const matchBR = (r) =>
    (!brandFilter || String(r[iBrand] || '').trim() === brandFilter) &&
    (!repFilter   || String(r[iRep]   || '').trim() === repFilter);

  // ── Metrics lookup: root → metrics row ────────────────────────────────────
  const metricsMap = new Map();
  metricsVals.forEach(r => {
    const root = String(r[xi['RootApptID']] || '').trim();
    if (root && !metricsMap.has(root)) metricsMap.set(root, r);
  });

  // ── Build output ──────────────────────────────────────────────────────────
  const out = [];

  masterVals.forEach(r => {
    // Filter: phải trong date window và match Brand/Rep
    if (!inWin(r[iVDate]) || !matchBR(r)) return;

    const root     = String(r[iRoot]   || '').trim();
    const brand    = String(r[iBrand]  || '').trim();
    const rep      = String(r[iRep]    || '').trim();
    const vtype    = String(r[iVType]  || '').trim();
    const vdate    = r[iVDate];
    const status   = iStatus >= 0 ? String(r[iStatus] || '').trim() : '';
    const custName = iCust   >= 0 ? String(r[iCust]   || '').trim() : '';

    // ── Derived status flags ─────────────────────────────────────────────────
    const statusLow  = status.toLowerCase();
    const isNoShow   = /no[-\s]?show/i.test(status);
    const isCompleted= /completed?/i.test(status);

    // ── Order Total từ Metrics ────────────────────────────────────────────────
    let orderTotal = '';
    const mRow = metricsMap.get(root);
    if (mRow) {
      const ot = mRow[xi['Order Total']];
      if (ot !== '' && ot !== null && isFinite(Number(ot)) && Number(ot) > 0) {
        orderTotal = Number(ot);
      }
    }

    // ── Deposit Amount từ Payment Map ─────────────────────────────────────────
    let depositAmt  = '';
    let depositMade = 'No';
    const fp = firstPayMap.get(root);
    if (fp && fp.amt && Number(fp.amt) > 0) {
      depositAmt  = Number(fp.amt);
      depositMade = 'Yes';
    }

    out.push([
      root,         // A: Route Appointment ID
      custName,     // B: Customer Name
      vdate,        // C: Date of Appointment
      vtype,        // D: Appointment Type
      status,       // E: Appointment Status
      isNoShow ? 'Yes' : 'No',        // F: No-Show?
      isCompleted ? 'Yes' : 'No',     // G: Completed?
      depositAmt,   // H: Deposit Amount
      depositMade,  // I: Deposit Made?
      orderTotal,   // J: Order Total
      rep,          // K: Sales Rep
      brand         // L: Brand
    ]);
  });

  // Sắp xếp: Date tăng dần → Rep → Customer
  out.sort((a, b) => {
    const da = asDate_(a[2]), db = asDate_(b[2]);
    if (da && db && da.getTime() !== db.getTime()) return da - db;
    const repCmp = String(a[10]).localeCompare(String(b[10]));
    if (repCmp !== 0) return repCmp;
    return String(a[1]).localeCompare(String(b[1]));
  });

  return out;
}


// =============================================================================
// WRITE SHEET — ghi dữ liệu + format toàn bộ
// =============================================================================

function _writeUnifiedSheet_(ss, rows, start, end, brandFilter, repFilter) {
  const tz      = Session.getScriptTimeZone() || 'GMT';
  const fmtDate = (d) => (d instanceof Date)
    ? Utilities.formatDate(d, tz, 'yyyy-MM-dd')
    : String(d || '');

  // ── Lấy hoặc tạo sheet ────────────────────────────────────────────────────
  // Xoa va tao lai sheet de tranh moi loi merge/filter/banding
  let sh = ss.getSheetByName(SH_UNIFIED);
  if (sh) {
    // Ghi nho vi tri de insert lai cung cho
    const sheetIndex = sh.getIndex();
    ss.deleteSheet(sh);
    sh = ss.insertSheet(SH_UNIFIED, sheetIndex - 1);
  } else {
    sh = ss.insertSheet(SH_UNIFIED);
  }
  sh.setHiddenGridlines(false);

  // ── Column headers ─────────────────────────────────────────────────────────
  const HEADERS = [
    'Route Appointment ID',
    'Customer Name',
    'Date of Appointment',
    'Appointment Type',
    'Appointment Status',
    'No-Show?',
    'Completed?',
    'Deposit Amount',
    'Deposit Made?',
    'Order Total',
    'Sales Rep',
    'Brand'
  ];
  const NUM_COLS = HEADERS.length; // 12

  // ── Row positions ──────────────────────────────────────────────────────────
  const TITLE_ROW  = 1;
  const FILTER_ROW = 2;
  const STATS_ROW  = 3;
  const HEADER_ROW = 5;
  const DATA_START = 6;

  // ── Title ──────────────────────────────────────────────────────────────────
  sh.getRange(TITLE_ROW, 1, 1, NUM_COLS).merge()
    .setValue('📊  Unified Drill-Down — Appointments × Deposits × Orders')
    .setBackground(UNI_HEADER_BG)
    .setFontColor(UNI_HEADER_FG)
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sh.setRowHeight(TITLE_ROW, 36);

  // ── Filter info bar ────────────────────────────────────────────────────────
  const periodLabel = fmtDate(start) + '  →  ' + fmtDate(end);
  const filterText  =
    '📅 Period: ' + periodLabel +
    (brandFilter ? '     🏷 Brand: ' + brandFilter : '     🏷 Brand: All') +
    (repFilter   ? '     👤 Rep: ' + repFilter      : '     👤 Rep: All');

  sh.getRange(FILTER_ROW, 1, 1, NUM_COLS).merge()
    .setValue(filterText)
    .setBackground(UNI_SUBHEAD_BG)
    .setFontColor('#0D47A1')
    .setFontSize(10)
    .setFontWeight('normal')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sh.setRowHeight(FILTER_ROW, 24);

  // ── Summary stats ──────────────────────────────────────────────────────────
  const totalAppts     = rows.length;
  const noShowCount    = rows.filter(r => r[5] === 'Yes').length;
  const completedCount = rows.filter(r => r[6] === 'Yes').length;
  const depositCount   = rows.filter(r => r[8] === 'Yes').length;
  const totalDeposit   = rows.reduce((s, r) => s + (Number(r[7]) || 0), 0);
  const totalOrder     = rows.reduce((s, r) => s + (Number(r[9]) || 0), 0);
  const noShowRate     = totalAppts ? ((noShowCount / totalAppts) * 100).toFixed(1) + '%' : 'N/A';

  const statsText =
    'Total: ' + totalAppts + ' appts' +
    '     ✅ Completed: ' + completedCount +
    '     ❌ No-Show: ' + noShowCount + ' (' + noShowRate + ')' +
    '     💰 Deposits: ' + depositCount +
    '     Deposit $: $' + Math.round(totalDeposit).toLocaleString('en-US') +
    '     Order $: $' + Math.round(totalOrder).toLocaleString('en-US');

  sh.getRange(STATS_ROW, 1, 1, NUM_COLS).merge()
    .setValue(statsText)
    .setBackground('#FFFFFF')
    .setFontColor('#37474F')
    .setFontSize(10)
    .setFontStyle('italic')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sh.setRowHeight(STATS_ROW, 22);

  // Row 4: khoảng trắng
  sh.setRowHeight(4, 8);

  // ── Header row ─────────────────────────────────────────────────────────────
  const hdrRange = sh.getRange(HEADER_ROW, 1, 1, NUM_COLS);
  hdrRange.setValues([HEADERS])
    .setBackground(UNI_HEADER_BG)
    .setFontColor(UNI_HEADER_FG)
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sh.setRowHeight(HEADER_ROW, 30);
  sh.setFrozenRows(HEADER_ROW);
  // Note: setFrozenColumns removed — conflicts with merged title/filter cells

  // ── Data ───────────────────────────────────────────────────────────────────
  if (rows.length === 0) {
    sh.getRange(DATA_START, 1, 1, NUM_COLS).merge()
      .setValue('No data found for the selected filters.')
      .setFontStyle('italic')
      .setFontColor('#888888')
      .setHorizontalAlignment('center');
  } else {
    // Chuẩn bị display values (date format)
    const displayRows = rows.map(r => {
      const d = asDate_(r[2]);
      return [
        r[0],                                            // RootApptID
        r[1],                                            // Customer
        d ? fmtDate(d) : (r[2] ? String(r[2]) : ''),    // Date formatted
        r[3],                                            // Type
        r[4],                                            // Status
        r[5],                                            // No-Show?
        r[6],                                            // Completed?
        r[7],                                            // Deposit $
        r[8],                                            // Deposit Made?
        r[9],                                            // Order Total
        r[10],                                           // Sales Rep
        r[11]                                            // Brand
      ];
    });

    const dataRange = sh.getRange(DATA_START, 1, rows.length, NUM_COLS);
    dataRange.setValues(displayRows);

    // ── Apply row-by-row formatting ──────────────────────────────────────────
    // Batch để tránh quota
    const bgArr    = [];
    const fgArr    = [];
    const fwArr    = [];
    const halignH  = [];   // Horizontal alignment per row
    const halignD  = [];
    const halignN  = [];
    const halignC  = [];
    const halignDM = [];
    const halignOT = [];

    // Pre-build color arrays
    const statusColBg  = [];  // col E (index 4)
    const statusColFg  = [];
    const noShowColBg  = [];  // col F (index 5)
    const noShowColFg  = [];
    const compColBg    = [];  // col G (index 6)
    const compColFg    = [];
    const depMadeColBg = [];  // col I (index 8)
    const depMadeColFg = [];

    rows.forEach((r, i) => {
      const isEven   = (i % 2 === 0);
      const baseBg   = isEven ? UNI_ODD_BG : UNI_EVEN_BG;

      // Row background (base)
      const rowBg = new Array(NUM_COLS).fill(baseBg);
      const rowFg = new Array(NUM_COLS).fill('#212121');
      const rowFw = new Array(NUM_COLS).fill('normal');

      bgArr.push(rowBg);
      fgArr.push(rowFg);
      fwArr.push(rowFw);

      // Status column (E = col 5 = index 4)
      const statusLow = String(r[4] || '').toLowerCase();
      let sBg = baseBg, sFg = '#212121';
      if (/no[-\s]?show/.test(statusLow))         { sBg = UNI_NO_BG;   sFg = UNI_NO_FG;   }
      else if (/completed?/.test(statusLow))       { sBg = UNI_YES_BG;  sFg = UNI_YES_FG;  }
      else if (/cancel/.test(statusLow))           { sBg = UNI_WARN_BG; sFg = UNI_WARN_FG; }
      else if (/scheduled|rescheduled/.test(statusLow)) { sBg = UNI_SCHED_BG; sFg = UNI_SCHED_FG; }
      statusColBg.push([sBg]);
      statusColFg.push([sFg]);

      // No-Show? column (F = col 6 = index 5)
      noShowColBg.push([r[5] === 'Yes' ? UNI_NO_BG   : baseBg]);
      noShowColFg.push([r[5] === 'Yes' ? UNI_NO_FG   : '#212121']);

      // Completed? column (G = col 7 = index 6)
      compColBg.push([r[6] === 'Yes' ? UNI_YES_BG  : baseBg]);
      compColFg.push([r[6] === 'Yes' ? UNI_YES_FG  : '#212121']);

      // Deposit Made? column (I = col 9 = index 8)
      depMadeColBg.push([r[8] === 'Yes' ? UNI_YES_BG : baseBg]);
      depMadeColFg.push([r[8] === 'Yes' ? UNI_YES_FG : '#888888']);
    });

    // Apply base colors to whole range
    dataRange.setBackgrounds(bgArr).setFontColors(fgArr).setFontWeights(fwArr);

    // Apply status-specific colors column by column
    const n = rows.length;
    sh.getRange(DATA_START, 5, n, 1).setBackgrounds(statusColBg).setFontColors(statusColFg).setFontWeights(statusColFg.map(c => c[0] !== '#212121' ? ['bold'] : ['normal']));
    sh.getRange(DATA_START, 6, n, 1).setBackgrounds(noShowColBg).setFontColors(noShowColFg);
    sh.getRange(DATA_START, 7, n, 1).setBackgrounds(compColBg).setFontColors(compColFg);
    sh.getRange(DATA_START, 9, n, 1).setBackgrounds(depMadeColBg).setFontColors(depMadeColFg);

    // Number formats
    sh.getRange(DATA_START, 8, n, 1).setNumberFormat('$#,##0');   // Deposit Amount
    sh.getRange(DATA_START, 10, n, 1).setNumberFormat('$#,##0');  // Order Total

    // Alignment
    sh.getRange(DATA_START, 1, n, NUM_COLS).setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.getRange(DATA_START, 3, n, 1).setHorizontalAlignment('center'); // Date
    sh.getRange(DATA_START, 6, n, 2).setHorizontalAlignment('center'); // No-Show / Completed
    sh.getRange(DATA_START, 8, n, 3).setHorizontalAlignment('right');  // Deposit/Made/Order
    sh.getRange(DATA_START, 12, n, 1).setHorizontalAlignment('center'); // Brand

    // Row heights
    sh.setRowHeightsForced(DATA_START, n, 22);

    // Border toàn bảng
    sh.getRange(HEADER_ROW, 1, n + 1, NUM_COLS)
      .setBorder(true, true, true, true, true, true,
        UNI_BORDER, SpreadsheetApp.BorderStyle.SOLID);

    // Filter
    sh.getRange(HEADER_ROW, 1, n + 1, NUM_COLS).createFilter();
  }

  // ── Column widths ──────────────────────────────────────────────────────────
  sh.setColumnWidth(1,  160);  // RootApptID
  sh.setColumnWidth(2,  160);  // Customer Name
  sh.setColumnWidth(3,  130);  // Date
  sh.setColumnWidth(4,  140);  // Appointment Type
  sh.setColumnWidth(5,  130);  // Status
  sh.setColumnWidth(6,   80);  // No-Show?
  sh.setColumnWidth(7,   90);  // Completed?
  sh.setColumnWidth(8,  110);  // Deposit Amount
  sh.setColumnWidth(9,  110);  // Deposit Made?
  sh.setColumnWidth(10, 110);  // Order Total
  sh.setColumnWidth(11, 140);  // Sales Rep
  sh.setColumnWidth(12,  90);  // Brand

  SpreadsheetApp.flush();
}





// =============================================================================
// INTEGRATION — Patch vào runOnceToBuildAll()
// =============================================================================
//
//  Trong file dashboard.gs, tìm function runOnceToBuildAll() và thêm dòng:
//
//    function runOnceToBuildAll() {
//      safeCall_(ensureDashboardLayout_);
//      safeCall_(buildMetricsView_);
//      safeCall_(writeDashboard_);
//      safeCall_(snapshotKpisForHistory_);
//      safeCall_(buildUnifiedDrillDown_);   // ← THÊM DÒNG NÀY
//    }
//
// =============================================================================


// =============================================================================
// MANUAL TRIGGER — có thể chạy thẳng từ Apps Script editor để test
// =============================================================================

function runBuildUnifiedDrillDown() {
  buildUnifiedDrillDown_();
  SpreadsheetApp.getUi().alert(
    '✅ Unified Drill-Down đã được tạo!',
    'Sheet "' + SH_UNIFIED + '" đã được cập nhật với dữ liệu mới nhất.\n\n' +
    'Vui lòng mở tab "Drill_Unified" để xem.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}