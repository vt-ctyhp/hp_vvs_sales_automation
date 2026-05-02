/** ============================================================================
 * PROJECT #15 — Production Stage Expansion
 * ---------------------------------------------------------------------------
 * Mở rộng tracking "In Production" với 5 sub-stage chi tiết:
 *   DBXM · Cooling · Hot · Preparing Forecasting · On Hold
 *
 * KIẾN TRÚC (đúng — không thêm cột vào Master):
 *   Custom Order Status  = high-level  ("In Production")       ← giữ nguyên
 *   In Production Status = sub-stage   (DBXM / Cooling / …)   ← mở rộng thêm 5 giá trị
 *   Timestamp stage      = đọc từ 03_Client_Status_Log         ← KHÔNG thêm cột Master
 *
 * RULE:
 *   Rule 1 → Stage chỉ active khi Custom Order Status = "In Production"
 *   Rule 2 → Mỗi order chỉ có 1 stage tại một thời điểm
 *   Rule 3 → History log tự động qua 03_Client_Status_Log (đã có sẵn)
 *
 * TRÌNH TỰ CÀI ĐẶT:
 *   Bước 1 → P15_setup()        – thêm 5 stage vào sheet Dropdown
 *   Bước 2 → Patch ClientStatus_v1.gs  (xem project15_patches_guide.gs)
 *   Bước 3 → Patch dashboard.gs        (xem project15_patches_guide.gs)
 *   Bước 4 → buildMetricsView_()       – rebuild 100_Metrics_View
 *   Bước 5 → runOnceToBuildAll()       – refresh 00_Dashboard
 *   Bước 6 → P15_verify()              – xác nhận mọi thứ đúng
 * ============================================================================ */


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  CONSTANTS                                                               ║
// ╚══════════════════════════════════════════════════════════════════════════╝

const P15_STAGES = [
  'DBXM',
  'Cooling',
  'Hot',
  'Preparing Forecasting',
  'On Hold'
];

const P15_IPS_COL            = 'In Production Status'; // cột đã có trong Master
const P15_ON_HOLD_ALERT_DAYS = 3;                      // cảnh báo On Hold > N ngày

/**
 * KPI card specs cho Dashboard — Production Stage Breakdown.
 * On Hold: upGood:false → tăng = đỏ (bottleneck).
 */
const KPI_CARDS_PRODUCTION = P15_STAGES.map(s => ({
  key:    'prod_' + s.toLowerCase().replace(/\s+/g, '_'),
  label:  s,
  fmt:    '0',
  upGood: s !== 'On Hold',
  drill:  'prod_stage_' + s.toLowerCase().replace(/\s+/g, '_')
}));


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  BƯỚC 1 — Thêm stage vào sheet Dropdown (idempotent)                    ║
// ╚══════════════════════════════════════════════════════════════════════════╝

function P15_setup() {
  const ss   = SpreadsheetApp.getActive();
  const drop = ss.getSheetByName('Dropdown');
  if (!drop) throw new Error('P15_setup: Không tìm thấy sheet "Dropdown".');

  const lastCol = drop.getLastColumn();
  const lastRow = drop.getLastRow();
  const header  = drop.getRange(1, 1, 1, lastCol).getValues()[0]
                      .map(h => String(h || '').trim());

  let ipsColIdx = header.findIndex(h => /in\s*production\s*status/i.test(h));
  if (ipsColIdx < 0) {
    ipsColIdx = lastCol;
    drop.getRange(1, ipsColIdx + 1).setValue('In Production Status');
    Logger.log('P15_setup: Tạo cột "In Production Status" tại cột ' + (ipsColIdx + 1));
  }

  const ipsCol1     = ipsColIdx + 1;
  const existingVals = new Set();
  if (lastRow > 1) {
    drop.getRange(2, ipsCol1, lastRow - 1, 1).getValues()
        .forEach(r => { const v = String(r[0] || '').trim(); if (v) existingVals.add(v); });
  }

  const toAdd = P15_STAGES.filter(s => !existingVals.has(s));
  if (!toAdd.length) {
    SpreadsheetApp.getUi().alert(
      '✅ P15 Bước 1: Tất cả production stage đã có trong Dropdown.\n\n'
      + 'Tiếp theo → áp dụng patches (Bước 2–3).'
    );
    return;
  }

  drop.getRange(Math.max(2, lastRow + 1), ipsCol1, toAdd.length, 1)
      .setValues(toAdd.map(s => [s]));

  SpreadsheetApp.getUi().alert(
    '✅ P15 Bước 1 Hoàn tất!\n\nĐã thêm vào "In Production Status":\n• '
    + toAdd.join('\n• ')
    + '\n\nTiếp theo → áp dụng patch GS (Bước 2).'
  );
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  ĐỌC LỊCH SỬ STAGE TỪ 03_Client_Status_Log                             ║
// ║  Thay thế hoàn toàn cho timestamp column trong Master                   ║
// ╚══════════════════════════════════════════════════════════════════════════╝

/**
 * Đọc 03_Client_Status_Log → Map<RootApptID, {ips, updatedAt}>.
 * Mỗi entry = lần cập nhật In Production Status gần nhất của từng Root.
 * Dùng để tính "Days in Stage" và phát hiện On Hold lâu.
 *
 * @returns {Map<string, {ips:string, updatedAt:Date|null}>}
 */
function P15_buildLatestIPSMapFromLog() {
  const ss  = SpreadsheetApp.getActive();
  const log = ss.getSheetByName('03_Client_Status_Log');
  if (!log) {
    Logger.log('P15_buildLatestIPSMapFromLog: Không tìm thấy "03_Client_Status_Log".');
    return new Map();
  }

  const lastRow = log.getLastRow();
  if (lastRow < 2) return new Map();

  const data   = log.getDataRange().getValues();
  const header = data[0].map(h => String(h || '').trim());

  const iAppt  = _p15FindCol(header, ['APPT_ID', 'Appt ID', 'APPTID']);
  const iIPS   = _p15FindCol(header, ['In Production Status', 'IPS']);
  const iUpdAt = _p15FindCol(header, ['Updated At', 'UpdatedAt', 'Timestamp']);

  if (iAppt < 0 || iIPS < 0) {
    Logger.log('P15_buildLatestIPSMapFromLog: Thiếu cột APPT_ID hoặc IPS trong log.');
    return new Map();
  }

  const apptToRoot = _p15BuildApptToRootMap();
  const result     = new Map(); // root → { ips, updatedAt }

  // Duyệt từ cuối lên (mới nhất trước)
  for (let r = data.length - 1; r >= 1; r--) {
    const row    = data[r];
    const apptId = String(row[iAppt] || '').trim();
    const ips    = String(row[iIPS]  || '').trim();
    if (!apptId || !ips) continue;

    const root = apptToRoot.get(apptId) || apptId;
    if (!result.has(root)) {
      const updatedAt = (iUpdAt >= 0 && row[iUpdAt] instanceof Date)
                         ? row[iUpdAt] : null;
      result.set(root, { ips, updatedAt });
    }
  }

  return result;
}

/**
 * Tính số ngày đơn hàng đang ở stage hiện tại.
 * @param {Date|null} stageUpdatedAt  Timestamp từ log
 * @param {Date}      today
 * @returns {number|''}
 */
function P15_daysInCurrentStage(stageUpdatedAt, today) {
  if (!(stageUpdatedAt instanceof Date) || isNaN(stageUpdatedAt)) return '';
  return Math.max(0, Math.floor(
    (today.getTime() - stageUpdatedAt.getTime()) / 86400000
  ));
}

function _p15BuildApptToRootMap() {
  const ss     = SpreadsheetApp.getActive();
  const master = ss.getSheetByName('00_Master Appointments');
  if (!master) return new Map();
  const data   = master.getDataRange().getValues();
  const header = data[0].map(h => String(h || '').trim());
  const iAppt  = _p15FindCol(header, ['APPT_ID', 'Appt ID', 'APPTID']);
  const iRoot  = _p15FindCol(header, ['RootApptID', 'Root Appt ID', 'ROOT']);
  if (iAppt < 0 || iRoot < 0) return new Map();
  const map = new Map();
  for (let r = 1; r < data.length; r++) {
    const appt = String(data[r][iAppt] || '').trim();
    const root = String(data[r][iRoot] || '').trim();
    if (appt && root && !map.has(appt)) map.set(appt, root);
  }
  return map;
}

function _p15FindCol(header, aliases) {
  const norm = header.map(h => String(h || '').trim().toLowerCase());
  for (const a of aliases) {
    const i = norm.indexOf(a.toLowerCase());
    if (i >= 0) return i;
  }
  return -1;
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  PRODUCTION BREAKDOWN — cho computeKpis_() trong dashboard.gs           ║
// ╚══════════════════════════════════════════════════════════════════════════╝

/**
 * Đếm đơn hàng theo production sub-stage từ 100_Metrics_View.
 * @returns {Object}  { prod_dbxm:N, prod_cooling:N, ... }
 */
function P15_computeProductionBreakdown(metrics, xi, brand, rep) {
  const result = {};
  KPI_CARDS_PRODUCTION.forEach(card => { result[card.key] = 0; });

  const cosIdx   = xi['Custom Order Status'];
  const ipsIdx   = xi[P15_IPS_COL];
  const brandIdx = xi['Brand'];
  const repIdx   = xi['Assigned Rep'];

  if (cosIdx == null || ipsIdx == null) {
    Logger.log(
      'P15_computeProductionBreakdown: Thiếu cột COS hoặc IPS trong Metrics View. '
      + 'Áp dụng dashboard.gs patches + buildMetricsView_().'
    );
    return result;
  }

  for (const r of metrics) {
    if (brand && String(r[brandIdx] || '').trim() !== brand) continue;
    if (rep   && String(r[repIdx]   || '').trim() !== rep)   continue;
    if (String(r[cosIdx] || '').trim() !== 'In Production')  continue;

    const ips = String(r[ipsIdx] || '').trim();
    const key = 'prod_' + ips.toLowerCase().replace(/\s+/g, '_');
    if (key in result) result[key]++;
  }

  return result;
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  ON HOLD ALERT ENGINE                                                    ║
// ╚══════════════════════════════════════════════════════════════════════════╝

/**
 * Trả về danh sách đơn hàng On Hold > P15_ON_HOLD_ALERT_DAYS ngày.
 * Timestamp lấy từ 03_Client_Status_Log.
 */
function P15_getOnHoldAlerts(metrics, xi, brand, rep, asOf) {
  const cosIdx      = xi['Custom Order Status'];
  const ipsIdx      = xi[P15_IPS_COL];
  const brandIdx    = xi['Brand'];
  const repIdx      = xi['Assigned Rep'];
  const rootIdx     = xi['RootApptID'];
  const custIdx     = xi['Customer Name'];
  const deadlineIdx = xi['Prod Deadline'];

  if (cosIdx == null || ipsIdx == null) return [];

  const latestIPSMap = P15_buildLatestIPSMapFromLog(); // đọc từ log 1 lần
  const asOfMs       = (asOf instanceof Date ? asOf : new Date()).getTime();
  const alerts       = [];

  for (const r of metrics) {
    if (brand && String(r[brandIdx] || '').trim() !== brand) continue;
    if (rep   && String(r[repIdx]   || '').trim() !== rep)   continue;
    if (String(r[cosIdx] || '').trim() !== 'In Production')  continue;
    if (String(r[ipsIdx] || '').trim() !== 'On Hold')        continue;

    const root     = String(r[rootIdx] || '');
    const logEntry = latestIPSMap.get(root);
    let daysInStage = null;
    if (logEntry && logEntry.updatedAt instanceof Date) {
      daysInStage = Math.max(0,
        Math.floor((asOfMs - logEntry.updatedAt.getTime()) / 86400000)
      );
    }

    if (daysInStage == null || daysInStage >= P15_ON_HOLD_ALERT_DAYS) {
      alerts.push({
        root,
        customer:    String(r[custIdx]  || ''),
        daysInStage,
        rep:         String(r[repIdx]   || ''),
        prodDeadline:(deadlineIdx != null && r[deadlineIdx] instanceof Date)
                      ? r[deadlineIdx] : null
      });
    }
  }

  alerts.sort((a, b) => (b.daysInStage || 0) - (a.daysInStage || 0));
  return alerts;
}

/**
 * Entry point cho timed trigger.
 * Cài trigger: Apps Script → Triggers → P15_alertOnHoldOrders → time-based (1–4 giờ).
 */
function P15_alertOnHoldOrders() {
  const ss        = SpreadsheetApp.getActive();
  const metricsSh = ss.getSheetByName('100_Metrics_View');
  if (!metricsSh) { Logger.log('P15_alertOnHoldOrders: Không tìm thấy 100_Metrics_View.'); return; }

  const data = metricsSh.getDataRange().getValues();
  if (data.length < 2) return;

  const xi      = makeIdx_(data[0].map(h => String(h || '').trim()));
  const metrics = data.slice(1);
  const alerts  = P15_getOnHoldAlerts(metrics, xi, '', '', new Date());

  if (!alerts.length) {
    Logger.log('P15: ✅ Không có đơn nào On Hold quá ' + P15_ON_HOLD_ALERT_DAYS + ' ngày.');
    return;
  }

  const tz    = (typeof CS_TZ !== 'undefined' ? CS_TZ : 'America/Los_Angeles');
  const lines = alerts.map(a =>
    '• ' + a.customer + ' (' + a.root + ')'
    + ' | Rep: ' + a.rep
    + ' | On Hold: ' + (a.daysInStage != null ? a.daysInStage + ' ngày' : 'không rõ')
    + (a.prodDeadline ? ' | Deadline: ' + Utilities.formatDate(a.prodDeadline, tz, 'yyyy-MM-dd') : '')
  );

  Logger.log('⚠️ Đơn On Hold > ' + P15_ON_HOLD_ALERT_DAYS + ' ngày:\n' + lines.join('\n'));

  // Bật email: bỏ comment 2 dòng dưới
  // const email = Session.getActiveUser().getEmail();
  // if (email) GmailApp.sendEmail(email, '⚠️ On Hold Alert', lines.join('\n'));
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  DRILL DATA — dùng trong rebuildKpiDrill_() của dashboard.gs            ║
// ╚══════════════════════════════════════════════════════════════════════════╝

/**
 * Tạo rows drill cho từng production stage.
 * Timestamp "Stage Updated At" lấy từ 03_Client_Status_Log.
 */
function P15_buildProductionDrillRows(metrics, xi, stage, brand, rep) {
  const cosIdx      = xi['Custom Order Status'];
  const ipsIdx      = xi[P15_IPS_COL];
  const brandIdx    = xi['Brand'];
  const repIdx      = xi['Assigned Rep'];
  const rootIdx     = xi['RootApptID'];
  const custIdx     = xi['Customer Name'];
  const totalIdx    = xi['Order Total'];
  const deadlineIdx = xi['Prod Deadline'];

  if (cosIdx == null || ipsIdx == null) return [];

  const latestIPSMap = P15_buildLatestIPSMapFromLog();
  const now          = new Date();

  return metrics
    .filter(r => {
      if (brand && String(r[brandIdx] || '').trim() !== brand) return false;
      if (rep   && String(r[repIdx]   || '').trim() !== rep)   return false;
      return String(r[cosIdx] || '').trim() === 'In Production'
          && String(r[ipsIdx] || '').trim() === stage;
    })
    .map(r => {
      const root     = String(r[rootIdx] || '');
      const logEntry = latestIPSMap.get(root);
      const updDate  = (logEntry && logEntry.updatedAt instanceof Date) ? logEntry.updatedAt : null;
      const days     = P15_daysInCurrentStage(updDate, now);
      return [
        root,
        String(r[custIdx] || ''),
        String(r[repIdx]  || ''),
        stage,
        updDate || '',
        days,
        (totalIdx    != null ? r[totalIdx]    || '' : ''),
        (deadlineIdx != null ? r[deadlineIdx] || '' : '')
      ];
    })
    .sort((a, b) => {
      const da = typeof a[5] === 'number' ? a[5] : -1;
      const db = typeof b[5] === 'number' ? b[5] : -1;
      return db - da; // lâu nhất lên đầu
    });
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  VERIFY                                                                  ║
// ╚══════════════════════════════════════════════════════════════════════════╝

function P15_verify() {
  const ss = SpreadsheetApp.getActive();
  const issues = [], checks = [];

  // 1. Dropdown
  const drop = ss.getSheetByName('Dropdown');
  if (!drop) {
    issues.push('❌ Không tìm thấy sheet "Dropdown"');
  } else {
    const lastRow = drop.getLastRow(), lastCol = drop.getLastColumn();
    const header  = drop.getRange(1, 1, 1, lastCol).getValues()[0];
    const ipsIdx  = header.findIndex(h => /in\s*production\s*status/i.test(String(h || '')));
    if (ipsIdx < 0) {
      issues.push('❌ Thiếu cột "In Production Status" trong Dropdown → chạy P15_setup()');
    } else {
      const existing = new Set(
        lastRow > 1
          ? drop.getRange(2, ipsIdx + 1, lastRow - 1, 1).getValues()
                .flat().map(v => String(v || '').trim()).filter(Boolean)
          : []
      );
      P15_STAGES.forEach(s => {
        if (existing.has(s)) checks.push('✅ Stage có trong Dropdown: ' + s);
        else                  issues.push('❌ Stage THIẾU: ' + s + ' → chạy P15_setup()');
      });
    }
  }

  // 2. Master KHÔNG có cột timestamp (đúng thiết kế)
  const master = ss.getSheetByName('00_Master Appointments');
  if (master) {
    const mHdr = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0]
                       .map(h => String(h || '').trim());
    checks.push(mHdr.includes(P15_IPS_COL)
      ? '✅ Master có cột "' + P15_IPS_COL + '"'
      : '❌ Master thiếu cột "' + P15_IPS_COL + '"');
    checks.push(!mHdr.includes('Production Stage Updated At')
      ? '✅ Master gọn — không có cột timestamp thừa'
      : '⚠️  Phát hiện cột timestamp trong Master — không cần, có thể xóa');
  }

  // 3. 03_Client_Status_Log
  const log = ss.getSheetByName('03_Client_Status_Log');
  if (!log) {
    issues.push('⚠️  "03_Client_Status_Log" không tìm thấy → Days in Stage sẽ không hoạt động');
  } else {
    const logHdr = log.getRange(1, 1, 1, log.getLastColumn()).getValues()[0]
                      .map(h => String(h || '').trim().toLowerCase());
    ['appt_id', 'in production status', 'updated at'].forEach(col => {
      if (logHdr.some(h => h.includes(col.split(' ')[0])))
        checks.push('✅ Log có cột: ' + col);
      else
        issues.push('⚠️  Log thiếu cột: ' + col);
    });
  }

  // 4. 100_Metrics_View
  const metricsSh = ss.getSheetByName('100_Metrics_View');
  if (!metricsSh) {
    issues.push('⚠️  100_Metrics_View chưa có → áp dụng dashboard patches + buildMetricsView_()');
  } else {
    const mxHdr = metricsSh.getRange(1, 1, 1, metricsSh.getLastColumn()).getValues()[0]
                            .map(h => String(h || '').trim());
    ['Custom Order Status', P15_IPS_COL].forEach(col => {
      if (mxHdr.includes(col)) checks.push('✅ Metrics View có cột: ' + col);
      else issues.push('⚠️  Metrics View thiếu cột: ' + col + ' → áp dụng patch D1–D4');
    });
  }

  // 5. KPI cards
  try {
    checks.push('✅ KPI_CARDS_PRODUCTION: ' + KPI_CARDS_PRODUCTION.length + ' cards');
  } catch (e) {
    issues.push('❌ KPI_CARDS_PRODUCTION lỗi: ' + e.message);
  }

  const allLines = [...issues, ...checks];
  const status   = issues.length === 0
    ? '✅ Tất cả pass — P15 đã cài đặt đầy đủ!'
    : '⚠️ ' + issues.length + ' vấn đề cần xử lý';

  Logger.log('═══ P15_verify ═══\n' + allLines.join('\n') + '\n\n' + status);
  SpreadsheetApp.getUi().alert('P15 Kiểm tra\n\n' + allLines.join('\n') + '\n\n' + status);
  return { ok: issues.length === 0, issues, checks };
}

