/** ===========================================================================
 *  VVS / HPUSA — Dashboard (clean, compact, no 99_* tabs)
 *  - Builds 100_Metrics_View (one row per RootApptID; durations, flags, weights)
 *  - Renders compact KPI cards on 00_Dashboard (header at A4; cards from A5)
 *  - Maintains a hidden KPI history block (00_Dashboard!AA5:AP) for sparklines
 *  - Uses in‑memory First Payment + 3D Status (no 99_* staging sheets)
 *  - Stage Weights live in 00_Dashboard!AS30:AT60 (or fallback defaults)
 *  - Orchestrator + hourly trigger entrypoint friendly
 *  =========================================================================== */

/** === SHEET NAMES (edit only if your tabs differ) ========================= */
const SH_MASTER   = '00_Master Appointments';
const SH_STATUSLOG= '03_Client_Status_Log';
const SH_ROOTIDX  = '07_Root_Index';
const SH_REPSMAP  = '08_Reps_Map';
const SH_METRICS  = '100_Metrics_View';
const SH_DASH     = '00_Dashboard';

/** === FILTER CELLS (top of 00_Dashboard) ================================== */
// Visual layout:
//  A2: "Date range" label
//  B1:D1: preset dropdown (merged)            → CELL_PRESET (top-left = B1)
//  B2: start date                              → CELL_DATE_S
//  C2: "→" arrow (visual only)
//  D2: end date                                → CELL_DATE_E
//  G1: "Brand" label  + H1:I1 merged input     → CELL_BRAND (top-left = I1)
//  G2: "Sales rep"   + H2:I2 merged input      → CELL_REP   (top-left = I2)
const CELL_PRESET = 'B1';   // merged B1:D1 (validation & reads use B1)
const CELL_DATE_S = 'B2';
const CELL_DATE_E = 'D2';
const CELL_BRAND  = 'H1';   // merged H1:I1 (validation & reads use H1)
const CELL_REP    = 'H2';   // merged H2:I2 (validation & reads use H2)

/** === HISTORY BLOCK (00_Dashboard hidden area) ============================ */
const HIST_START_COL   = 27; // AA
const HIST_HEADER_ROW  = 5;
const HIST_HEADERS = [
  'Date','Brand','Rep',
  'Appointments','Diamond Viewings',
  'Deposits (first-time)','Median 1st Appt→Dep (days)','Average Order Value','Weighted Pipeline',
  'No-touch >48h','DV no deposit >7d','3D wait >3d','3D overdue','Production overdue','Ops escalation'
];

/** === DEPOSIT RULES (exclude tiny DV/3D holds) ============================ */
// Treat FIRST deposit as the first valid receipt whose Net > MIN_FIRST_DEPOSIT_NET.
// Any receipt with Net <= MIN_FIRST_DEPOSIT_NET is ignored for "first deposit" logic.
const MIN_FIRST_DEPOSIT_NET = 25; // dollars; set 20 or 25 to your policy


/** === KPI CARDS — Weekly Scheduling ====================================== */
// A) Created This Week (inflow)
const KPI_CARDS_CREATED = [
  { key:'bookingsCreated',   label:'Bookings Created',    fmt:'0', upGood:true,  drill:'inf_bookingsCreated' },
  { key:'reschedulesCreated',label:'Reschedules Created', fmt:'0', upGood:false, drill:'inf_reschedulesCreated' },
  { key:'cancelsCreated',    label:'Cancels',             fmt:'0', upGood:false, drill:'inf_cancelsCreated' }
];

// B) On the Calendar This Week (what’s scheduled to happen)
const KPI_CARDS_ONCAL = [
  { key:'uniqueCustomersScheduled', label:'Unique Customers (Scheduled)', fmt:'0', upGood:true,  drill:'uniq_customersSch' },
  { key:'apptsScheduled',          label:'Appointments (Scheduled)',     fmt:'0', upGood:true,  drill:'cal_appts' },
  { key:'consultationsScheduled',  label:'Consultations',                fmt:'0', upGood:true,  drill:'cal_consults' },
  { key:'dvsScheduled',            label:'Diamond Viewings',             fmt:'0', upGood:true,  drill:'cal_dvs' }
];

// C) Changes to This Week’s Calendar (removals)
const KPI_CARDS_CHANGES = [
  { key:'reschedOffThisWeek',     label:'Rescheduled Off This Week', fmt:'0', upGood:false, drill:'chg_reschedOff' },
  { key:'cancelledFromThisWeek',  label:'Cancelled From This Week',  fmt:'0', upGood:false, drill:'chg_cancelledFrom' }
];

/** === KPI CARD SPECS (row groupings) ===================================== */
const KPI_CARDS_APPTS = [
  { key:'totalCustomers',    label:'Total Customers',             fmt:'0',      upGood:true,  drill:'appts_total_customers' },
  { key:'totalAppointments', label:'Total Appointments',          fmt:'0',      upGood:true,  drill:'appts_total_appointments' },
  { key:'consultations',     label:'Consultations',               fmt:'0',      upGood:true,  drill:'appts_consultations' },
  { key:'dvsInWindow',       label:'Diamond Viewings',            fmt:'0',      upGood:true,  drill:'appts_dv' },
  { key:'cohortSecondAppt',  label:'Cohort → 2nd',                fmt:'0',      upGood:true,  drill:'appts_cohort2nd' },
];

const KPI_CARDS_PAYMENTS = [
  { key:'depositsInWindow',    label:'Deposits (first-time)',       fmt:'0',      upGood:true,  drill:'pay_firstDeposits' },
  { key:'medianDvToDepDays',   label:'Median 1st Appt→Dep (days)',  fmt:'0',      upGood:false, drill:'pay_apptToDep' },
  { key:'aov',                 label:'Average Order Value',         fmt:'$#,##0', upGood:true,  drill:'pay_aov' },
  { key:'firstDepositsAmount', label:'First-time Deposits $',       fmt:'$#,##0', upGood:true,  drill:'pay_firstDeposits' },
  { key:'totalPaymentsAmount', label:'All Payments $',              fmt:'$#,##0', upGood:true,  drill:'pay_allDeposits' }
  // Removed: Weighted Pipeline (still available in drill, just not shown as a card)
];

const KPI_CARDS_RISK = [
  { key:'noTouch48', label:'No‑touch >48h',       fmt:'0', upGood:false, drill:'risk_noTouch48' },
  { key:'dvNoDep7',  label:'DV no deposit >7d',   fmt:'0', upGood:false, drill:'risk_dvNoDep7'  },
  { key:'wait3d',    label:'3D wait >3d',         fmt:'0', upGood:false, drill:'risk_3dWait'    },
  { key:'overdue3d', label:'3D overdue',          fmt:'0', upGood:false, drill:'risk_3dOverdue' },
  { key:'overduePr', label:'Production overdue',  fmt:'0', upGood:false, drill:'risk_prOverdue' },
  { key:'escal',     label:'Ops escalation',      fmt:'0', upGood:false, drill:'risk_escal'     },
];

// helpers (put near the top of the file)

function periodLabel_(start, end, presetCell) {
  const dash = SpreadsheetApp.getActive().getSheetByName(SH_DASH);
  const preset = String(dash.getRange(presetCell).getValue() || '').trim();
  const fmt = d => Utilities.formatDate(d, Session.getScriptTimeZone() || 'GMT', 'yyyy-MM-dd');
  if (!preset || /custom/i.test(preset)) return `${fmt(start)} → ${fmt(end)}`;
  return preset; // e.g., "This Week (Mon–Sun)", "This Month", etc.
}

function mergeIfNeeded_(sh, a1) {
  const r = sh.getRange(a1);
  // already merged exactly as desired?
  if (r.isPartOfMerge()) {
    const m = r.getMergedRanges()[0];
    if (m && m.getA1Notation() === a1) return; // correct; keep it
    m.breakApart(); // wrong region -> fix it
  }
  r.merge();
}
function unmergeIfPresent_(sh, a1) {
  const r = sh.getRange(a1);
  if (r.isPartOfMerge()) r.breakApart();
}


/** === ORCHESTRATION ======================================================= */
function runOnceToBuildAll() {
  safeCall_(ensureDashboardLayout_);
  safeCall_(buildMetricsView_);
  safeCall_(writeDashboard_);
  safeCall_(snapshotKpisForHistory_);
}
function refreshDashboardHourly() {
  runOnceToBuildAll();
}
function safeCall_(fn) {
  try { fn(); } catch (e) { console.error(`[safeCall] ${fn.name}: ${e && e.stack || e}`); }
}



/** === DASHBOARD LAYOUT (header at A4; cards from A5) ====================== */
function ensureDashboardLayout_() {
  const ss = SpreadsheetApp.getActive();
  const dash = ss.getSheetByName(SH_DASH) || ss.insertSheet(SH_DASH);

  // ── Labels ───────────────────────────────────────────────────────────────
  // A1: header label for the quick-pick row
  dash.getRange('A1')
    .setValue('⚡️ Quick Picker')
    .setBackground('#F6F8FB')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  dash.getRange('A2').setValue('📅 Date range');
  dash.getRange('G1').setValue('🏷️ Brand');
  dash.getRange('G2').setValue('👤 Sales rep');

  // ── Break any previous merges to make this idempotent ────────────────────
  // NEW (only change what’s necessary)
  unmergeIfPresent_(dash, 'E1:F1');   // legacy quick-pick merge we no longer use
  unmergeIfPresent_(dash, 'E2:F2');   // legacy rep merge we no longer use

  mergeIfNeeded_(dash, 'B1:D1');     // preset (top-left = B1)
  mergeIfNeeded_(dash, 'H1:I1');     // Brand input (top-left = I1)
  mergeIfNeeded_(dash, 'H2:I2');     // Rep input   (top-left = I2)


  // ── Merge to match the screenshot ────────────────────────────────────────
  dash.getRange('B1:D1').merge();   // preset
  dash.getRange('H1:I1').merge();   // Brand picker (visual width)
  dash.getRange('H2:I2').merge();   // Rep picker (visual width)

  // ── Initialize dates in B2 / D2 if empty ─────────────────────────────────
  if (!dash.getRange(CELL_DATE_S).getValue() || !dash.getRange(CELL_DATE_E).getValue()) {
    const mon = mondayOfWeek_(new Date());
    const sun = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate()+6);
    dash.getRange(CELL_DATE_S).setValue(mon).setNumberFormat('yyyy-mm-dd');
    dash.getRange(CELL_DATE_E).setValue(sun).setNumberFormat('yyyy-mm-dd');
  }

  // Visual arrow in C2
  dash.getRange('C2').setValue('→').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setFontColor('#9AA5B1').setFontWeight('bold');

  // ── Notes / hints ────────────────────────────────────────────────────────
  dash.getRange(CELL_BRAND).setNote('Leave blank for all brands.');
  dash.getRange(CELL_REP).setNote('Leave blank for all reps.');

  // ── Build dropdowns ──────────────────────────────────────────────────────
  ensurePeriodPreset_();      // writes validation to B1 (CELL_PRESET)
  ensureBrandRepDropdowns_(); // writes validations to H1 (brand) and H2 (rep)

  // ── Styling ──────────────────────────────────────────────────────────────
  dash.getRangeList(['A2','G1','G2'])
    .setBackground('#F6F8FB')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  dash.getRangeList([CELL_PRESET, CELL_DATE_S, CELL_DATE_E, CELL_BRAND, CELL_REP])
    .setBackground('#FFFFFF')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  // Border + sizes for A1:I2 (exact block shown)
  dash.getRange('A1:D2').setBorder(true,true,true,true,true,true,'#D9DEE8',SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange('G1:I2').setBorder(true,true,true,true,true,true,'#D9DEE8',SpreadsheetApp.BorderStyle.SOLID);
  dash.setRowHeights(1, 2, 26);

  // Date formats for inputs
  dash.getRange(CELL_DATE_S).setNumberFormat('yyyy-mm-dd');
  dash.getRange(CELL_DATE_E).setNumberFormat('yyyy-mm-dd');

  // Column widths (approximate the screenshot)
  dash.setColumnWidth(1, 110);  // A  label
  dash.setColumnWidth(2, 120);  // B  start
  dash.setColumnWidth(3, 36);   // C  arrow
  dash.setColumnWidth(4, 120);  // D  end
  dash.setColumnWidth(5, 24);   // E  spacer
  dash.setColumnWidth(6, 24);   // F  spacer
  dash.setColumnWidth(7, 120);  // G  right-side labels
  dash.setColumnWidth(8, 24);   // H  small spacer (merged with I visually)
  dash.setColumnWidth(9, 150);  // I  brand/rep inputs (anchor cells)

  // Keep the rest
  setCardGridColumnWidths_();
  dash.setHiddenGridlines(true);
  ensureHistoryBlock_();
  ensureStageWeightConfigBlock_();
}


/** === BUILD METRICS VIEW (no 99_* tabs; in-memory joins) ================== */
function buildMetricsView_() {
  const ss = SpreadsheetApp.getActive();
  const stageWeights = getStageWeightMap_();  // from config or defaults

  // --- Load master
  const mSh = ss.getSheetByName(SH_MASTER);
  const M = mSh.getDataRange().getValues(); const MH = M.shift().map(x=>String(x||'').trim());
  const iRoot = findCol_(MH, ['RootApptID','Root Appt ID','ROOT','Root_ID']);
  const iBrand= findCol_(MH, ['Brand']);
  const iRep  = findCol_(MH, ['Assigned Rep','AssignedRep','Rep','Sales Rep']);
  const iVisitType = findCol_(MH, ['Visit Type','VisitType','Type']);
  const iVisitDate = findCol_(MH, ['Visit Date','Visit_Date','Appt Date','Appointment Date']);
  const iSO    = findCol_(MH, ['SO#','SO Number','SO No','Sales Order #'], false);
  const iOTot  = findCol_(MH, ['Order Total','Total','SO Total'], false);
  const iPaid  = findCol_(MH, ['Paid-to-Date','Paid to Date','Paid'], false);
  const iRemain= findCol_(MH, ['Remaining Balance','Remain','Balance'], false);
  const iLastPay = findCol_(MH, ['Last Payment Date','Last Payment','Last Paid At'], false);
  const i3DDue = findCol_(MH, ['3D Deadline','3D Due','3D Due Date'], false);
  const i3DMoves = findCol_(MH, ['# of Times 3D Deadline Moved','3D Deadline Moves','# 3D Deadline Moves'], false);
  const iProdDue = findCol_(MH, ['Production Deadline','Prod Deadline'], false);
  const iProdMoves = findCol_(MH, ['# of Times Prod. Deadline Moved','Prod Deadline Moves','# Prod Deadline Moves'], false);
  const iStage = findCol_(MH, ['Sales Stage','Stage'], false);
  const iConv  = findCol_(MH, ['Conversion Status','Status'], false);
  const iSource= findCol_(MH, ['Source (normalized)','Source Normalized','Source'], false);
  const iBMin  = findCol_(MH, ['Budget Min','Budget (Min)','BudgetMin'], false);
  const iBMax  = findCol_(MH, ['Budget Max','Budget (Max)','BudgetMax'], false);
  const iPhone = findCol_(MH, ['Phone','Phone Number'], false);
  const iEmail = findCol_(MH, ['Email','Email Address'], false);
  const iOrderDate = findCol_(MH, ['Order Date','SO Date','Sales Order Date'], false);
  const iApptId = findCol_(MH, ['APPT_ID','Appt ID','APPTID','Appointment ID'], false);
  if (iRoot < 0 || iVisitDate < 0) throw new Error('Master must have RootApptID and Visit Date');
  const iCust = findCol_(MH, ['Customer Name','Customer','Client Name'], false);


  // --- Aggregate per root
  const per = new Map();
  const getAgg = (root) => {
    let o = per.get(root);
    if (!o) { o = {
      root, brand:null, rep:null, firstVisit:null, firstDV:null, visitCount:0,
      so:'', orderTotal:null, paidToDate:null, remain:null, lastPay:null,
      d3Due:null, d3Moves:0, prodDue:null, prodMoves:0,
      stage:null, conv:null, source:null, bMin:null, bMax:null,
      phone:null, email:null, orderDate:null, apptIds:new Set()
    }; per.set(root, o); }
    return o;
  };
  for (const r of M) {
    const root = r[iRoot]; if (!root) continue; const o = getAgg(root);
    const vDate = asDate_(r[iVisitDate]); const vType = iVisitType>=0 ? String(r[iVisitType]||'').trim().toLowerCase() : '';
    if (vDate) { o.visitCount++; if (!o.firstVisit || vDate < o.firstVisit) o.firstVisit = vDate; if (vType==='diamond viewing') { if (!o.firstDV || vDate < o.firstDV) o.firstDV = vDate; } }
    if (iBrand>=0 && r[iBrand]) o.brand = r[iBrand];
    if (iRep>=0 && r[iRep]) o.rep = r[iRep];
    if (iSO>=0 && r[iSO]) o.so = r[iSO];
    if (iOTot>=0 && isNum_(r[iOTot])) o.orderTotal = num_(r[iOTot]);
    if (iPaid>=0 && isNum_(r[iPaid])) o.paidToDate = num_(r[iPaid]);
    if (iRemain>=0 && isNum_(r[iRemain])) o.remain = num_(r[iRemain]);
    if (iLastPay>=0 && r[iLastPay] instanceof Date) o.lastPay = r[iLastPay];
    if (i3DDue>=0 && r[i3DDue] instanceof Date) o.d3Due = r[i3DDue];
    if (i3DMoves>=0 && isNum_(r[i3DMoves])) o.d3Moves = Math.max(o.d3Moves, num_(r[i3DMoves]));
    if (iProdDue>=0 && r[iProdDue] instanceof Date) o.prodDue = r[iProdDue];
    if (iProdMoves>=0 && isNum_(r[iProdMoves])) o.prodMoves = Math.max(o.prodMoves, num_(r[iProdMoves]));
    if (iStage>=0 && r[iStage]) o.stage = r[iStage];
    if (iConv>=0 && r[iConv]) o.conv = r[iConv];
    if (iSource>=0 && r[iSource]) o.source = r[iSource];
    if (iBMin>=0 && isNum_(r[iBMin])) o.bMin = num_(r[iBMin]);
    if (iBMax>=0 && isNum_(r[iBMax])) o.bMax = num_(r[iBMax]);
    if (iPhone>=0 && r[iPhone]) o.phone = r[iPhone];
    if (iEmail>=0 && r[iEmail]) o.email = r[iEmail];
    if (iOrderDate>=0 && r[iOrderDate] instanceof Date) o.orderDate = r[iOrderDate];
    if (iApptId>=0 && r[iApptId]) o.apptIds.add(r[iApptId]);
    if (iCust>=0 && r[iCust] && !o.custName) o.custName = r[iCust];

  }

  // --- Reps Map (Assigned / Assisted)
  const repsSh = ss.getSheetByName(SH_REPSMAP);
  const R = repsSh.getDataRange().getValues(); const RH = R.shift().map(x=>String(x||'').trim());
  const riRoot = findCol_(RH, ['RootApptID','Root Appt ID','ROOT','Root_ID']);
  const riRep  = findCol_(RH, ['Rep','Rep Name','Assigned Rep','Sales Rep','Name']);
  const riRole = findCol_(RH, ['Role','Rep Role']);
  const riInc  = findCol_(RH, ['Include?','Include','Use?'], false);
  const assignedByRoot = new Map(); const assistedByRoot = new Map();
  for (const r of R) {
    const root = r[riRoot], rep = r[riRep], role = String(r[riRole]||'').trim().toLowerCase();
    const inc  = riInc>=0 ? String(r[riInc]||'').trim().toUpperCase() : 'Y';
    if (!root || !rep || inc!=='Y') continue;
    if (role==='assigned') assignedByRoot.set(root, rep);
    else if (role==='assisted') { const arr = assistedByRoot.get(root)||[]; if (!arr.includes(rep)) arr.push(rep); assistedByRoot.set(root, arr); }
  }

  // --- Root Index (last touch)
  const rootIdx = ss.getSheetByName(SH_ROOTIDX);
  const RI = rootIdx.getDataRange().getValues(); const RIH = RI.shift().map(x=>String(x||'').trim());
  const rRoot = findCol_(RIH, ['RootApptID','Root Appt ID','ROOT','Root_ID']);
  const rUpd  = findCol_(RIH, ['Updated At','UpdatedAt','Last Updated'], false);
  const lastTouchByRoot = new Map();
  for (const r of RI) { const root=r[rRoot], d=rUpd>=0? r[rUpd]: null; if (root && d instanceof Date) lastTouchByRoot.set(root,d); }

  // --- External maps (in memory; no 99_* tabs)
  const firstPayByRoot = fetchFirstPaymentMapFromPayments_(); // Map<root,{d,amt,so}>
  const d3ByRoot       = fetch3DStatusMapAsMap_();            // Map<root,{req,res,pending}>

  // --- Build output rows
  const today = startOfDay_(new Date());
  const out = ss.getSheetByName(SH_METRICS) || ss.insertSheet(SH_METRICS);
  out.clear();
  const headers = [
    'RootApptID','Customer Name','Assigned Rep','Assisted Reps','Brand','First Visit Date','First DV Date',
    'Visit Count','Has 2nd Appt?','SO# (Active)','Deposit Made?','Deposit Date (First Pay)',
    'Lead→Deposit (biz days)','DV→Deposit (biz days)','Order Total','Paid-to-Date','Remaining Balance',
    'Order Date','Order Age (biz days)','Last Touch (Root Index)','No-touch >48h?','DV no deposit >7d?',
    'Last 3D Request Date','3D Pending?','3D Request Age (days)','3D Wait >3d?','3D Deadline','3D Overdue?',
    '# 3D Deadline Moves','Prod Deadline','Prod Overdue?','# Prod Deadline Moves','Sales Stage','Conversion Status',
    'Source (normalized)','Budget Min','Budget Max','Budget Band','Budget Midpoint','Active (≤90 days)',
    'Stage Weight','Weighted Pipeline Value','Ops Escalation?'
  ];
  out.appendRow(headers);

  const rows=[];
  for (const [root,o] of per.entries()) {
    const assigned = assignedByRoot.get(root) || o.rep || '';
    const assisted = (assistedByRoot.get(root)||[]).join(', ');

    const payInfo = firstPayByRoot.get(root) || { d:null, amt:null, so:'' };
    // Guard on amount too, even though the map is already filtered — keeps intent explicit
    const depositDate = (isFinite(Number(payInfo.amt)) && Number(payInfo.amt) > MIN_FIRST_DEPOSIT_NET) ? payInfo.d : null;
    const depositMade = (isFinite(Number(payInfo.amt)) && Number(payInfo.amt) > MIN_FIRST_DEPOSIT_NET);

    const leadToDep = (o.firstVisit && depositDate) ? businessDaysInclusive_(o.firstVisit, depositDate) : '';
    const dvToDep   = (o.firstDV && depositDate) ? businessDaysInclusive_(o.firstDV, depositDate) : '';

    const lastTouch = lastTouchByRoot.get(root) || o.lastPay || o.firstVisit || null;
    const noTouch48 = lastTouch ? ((new Date().getTime()-lastTouch.getTime())/(1000*3600) > 48) : false;
    const dvNoDep7  = (o.firstDV && !depositMade) ? ((today - startOfDay_(o.firstDV))/(1000*3600*24) > 7) : false;

    const d3 = d3ByRoot.get(root) || {req:null, res:null, pending:false};
    const d3ReqDate = d3.req; const d3Pending = !!d3.pending;
    const d3ReqAge  = d3ReqDate ? Math.floor((today - startOfDay_(d3ReqDate))/(1000*3600*24)) : '';
    const d3Wait3   = d3Pending && d3ReqAge!=='' ? d3ReqAge > 3 : false;

    const d3Overdue = (o.d3Due instanceof Date) ? (startOfDay_(o.d3Due) < today) : false;
    const prOverdue = (o.prodDue instanceof Date) ? (startOfDay_(o.prodDue) < today) : false;

    const budgetBand = computeBudgetBand_(o.bMax);
    const budgetMid  = (isNum_(o.bMin) && isNum_(o.bMax)) ? (o.bMin + o.bMax)/2 : '';
    const active90   = lastTouch ? ((new Date().getTime() - lastTouch.getTime())/(1000*3600*24) <= 90) : false;

    const stageNorm  = normalizeStage_(o.stage);
    const stageWeight= stageWeights[stageNorm] ?? 0;
    const valueForWeight = (isNum_(o.orderTotal) && o.orderTotal>0) ? o.orderTotal : (isNum_(budgetMid) ? budgetMid : 0);
    const weightedPipeline = active90 ? (stageWeight * valueForWeight) : 0;

    const orderAge = (o.orderDate instanceof Date) ? businessDaysInclusive_(o.orderDate, today) : '';
    const opsEscalation = (((o.d3Moves||0) >= 2) || ((o.prodMoves||0) >= 2)) && (isNum_(orderAge) ? orderAge > 28 : false);

    rows.push([
      root, (o.custName||''), assigned, assisted, (o.brand||''), (o.firstVisit||''), (o.firstDV||''),
      o.visitCount, o.visitCount>=2, o.so||'',
      depositMade, (depositDate||''),
      leadToDep, dvToDep,
      (isNum_(o.orderTotal)?o.orderTotal:''), (isNum_(o.paidToDate)?o.paidToDate:''), (isNum_(o.remain)?o.remain:''),
      (o.orderDate||''), orderAge, (lastTouch||''), noTouch48, dvNoDep7,
      (d3ReqDate||''), d3Pending, d3ReqAge, d3Wait3,
      (o.d3Due||''), d3Overdue, (o.d3Moves||0),
      (o.prodDue||''), prOverdue, (o.prodMoves||0),
      (o.stage||''), (o.conv||''), (o.source||''),
      (isNum_(o.bMin)?o.bMin:''), (isNum_(o.bMax)?o.bMax:''), budgetBand, (isNum_(budgetMid)?budgetMid:''),
      active90, stageWeight, weightedPipeline, opsEscalation
    ]);
  }
  rows.sort((a,b)=>String(a[0]).localeCompare(String(b[0])));
  if (rows.length) out.getRange(2,1,rows.length,headers.length).setValues(rows);
  out.setFrozenRows(1);

  try {
    // Dates
    out.getRange('F2:F').setNumberFormat('yyyy-mm-dd');        // First Visit Date
    out.getRange('G2:G').setNumberFormat('yyyy-mm-dd');        // First DV Date
    out.getRange('L2:L').setNumberFormat('yyyy-mm-dd');        // Deposit Date (First Pay)
    out.getRange('R2:R').setNumberFormat('yyyy-mm-dd');        // Order Date
    out.getRange('T2:T').setNumberFormat('yyyy-mm-dd hh:mm');  // Last Touch (Root Index)
    out.getRange('W2:W').setNumberFormat('yyyy-mm-dd');        // Last 3D Request Date
    out.getRange('AA2:AA').setNumberFormat('yyyy-mm-dd');      // 3D Deadline
    out.getRange('AD2:AD').setNumberFormat('yyyy-mm-dd');      // Prod Deadline

    // Money
    out.getRange('O2:O').setNumberFormat('$#,##0');            // Order Total
    out.getRange('P2:P').setNumberFormat('$#,##0');            // Paid-to-Date
    out.getRange('Q2:Q').setNumberFormat('$#,##0');            // Remaining Balance
    out.getRange('AM2:AM').setNumberFormat('$#,##0');          // Budget Midpoint
    out.getRange('AP2:AP').setNumberFormat('$#,##0');          // Weighted Pipeline Value

    // Weights & counts / durations
    out.getRange('AO2:AO').setNumberFormat('0.00');            // Stage Weight
    out.getRange('M2:N').setNumberFormat('0');                 // Lead→Dep / DV→Dep (biz days)
    out.getRange('S2:S').setNumberFormat('0');                 // Order Age (biz days)
    out.getRange('Y2:Y').setNumberFormat('0');                 // 3D Request Age (days)

    // Optional: booleans as text to avoid TRUE/FALSE coercion surprises
    // out.getRange('I2:I').setNumberFormat('@');              // Has 2nd Appt?
    // out.getRange('K2:K').setNumberFormat('@');              // Deposit Made?
    // out.getRange('U2:U').setNumberFormat('@');              // No-touch >48h?
    // ...etc.
  } catch(_) {}
}

/** === WRITE DASHBOARD (3 rows; compact cards; with drill anchors) ========= */
function writeDashboard_() {
  const ss = SpreadsheetApp.getActive();
  const dash = ss.getSheetByName(SH_DASH);
  if (!dash) throw new Error(`Missing sheet ${SH_DASH}`);
  ensureBrandRepDropdowns_();     // build dropdowns from data
  setCardGridColumnWidths_();     // half-width cards + quarter gaps

  // Inputs
  const start = asDate_(dash.getRange(CELL_DATE_S).getValue()); // B2
  const end   = asDate_(dash.getRange(CELL_DATE_E).getValue()); // D2
  const brandFilter = String(dash.getRange(CELL_BRAND).getValue() || '').trim(); // H1
  const repFilter   = String(dash.getRange(CELL_REP).getValue()   || '').trim(); // H2
  if (!(start instanceof Date) || !(end instanceof Date)) throw new Error('Window Start/End must be dates in B2/D2.');

  // Prev window (same length)
  const days = Math.round((startOfDay_(end) - startOfDay_(start)) / 86400000) + 1;
  const prevStart = new Date(start.getFullYear(), start.getMonth(), start.getDate() - days);
  const prevEnd   = new Date(end.getFullYear(),   end.getMonth(),   end.getDate()   - days);

  // Data
  const master = ss.getSheetByName(SH_MASTER).getDataRange().getValues();
  const mH = master.shift().map(x=>String(x||'').trim());
  const metrics = ss.getSheetByName(SH_METRICS).getDataRange().getValues();
  const xH = metrics.shift().map(x=>String(x||'').trim());

  // KPIs (current & previous)
  const cur = computeKpis_(start, end, brandFilter, repFilter, master, mH, metrics, xH);
  const prv = computeKpis_(prevStart, prevEnd, brandFilter, repFilter, master, mH, metrics, xH);

  // SCHEDULING KPIs (Created / On-Calendar / Changes)
  const schedCur = computeScheduleKpis_(start, end, brandFilter, repFilter, master, mH);
  const schedPrv = computeScheduleKpis_(prevStart, prevEnd, brandFilter, repFilter, master, mH);

  // Rebuild drill sheet first so we can link to anchors
  const anchors = rebuildKpiDrill_(start, end, brandFilter, repFilter, master, mH, metrics, xH);

  // (NEW) Get gid for the drill sheet so HYPERLINKs don’t throw
  const drillSheet = ss.getSheetByName('Drill_KPI');
  const drillGid = drillSheet ? drillSheet.getSheetId() : null;

  // Clear visible canvas
  dash.getRange('A4:Y80').clearContent().clearFormat();

  // Row 40: Charts (pipeline, weekly deposits, order flow)
  writeChartsRow4to15({
    start, end, brand: brandFilter, rep: repFilter,
    master, mH, metrics, xH,
    anchors, drillGid
  });

  // Optional label if you still want it above cards
  // dash.getRange('A4').setValue('KPI Snapshot').setFontWeight('bold').setFontSize(12);

  // === Weekly Scheduling KPIs ================================================
  // Helper to advance the writing cursor: 1 header row + N card rows + 1 blank
  const adv = (specLen) => 1 + Math.ceil(specLen / 10) * (4 + 1) + 1;
  const label = periodLabel_(start, end, CELL_PRESET);

  // Start at row 4 (header row). Cards will begin at row 5.
  let cardRow = 4;

  // A) Created This Week (inflow)
  dash.getRange(cardRow, 1).setValue(`Created — ${label}`).setFontWeight('bold').setFontSize(12);
  renderScorecardsCompact_({
    sheet: dash, startRow: cardRow + 1, startCol: 1,
    cardsPerRow: 10, cardWidth: 2, cardHeight: 4, hGap: 1, vGap: 1,
    spec: KPI_CARDS_CREATED, cur: schedCur, prv: schedPrv, addSparkline: false, anchors
  });
  cardRow += adv(KPI_CARDS_CREATED.length);

  // B) On the Calendar This Week
  dash.getRange(cardRow, 1).setValue(`On the Calendar — ${label}`).setFontWeight('bold').setFontSize(12);
  renderScorecardsCompact_({
    sheet: dash, startRow: cardRow + 1, startCol: 1,
    cardsPerRow: 10, cardWidth: 2, cardHeight: 4, hGap: 1, vGap: 1,
    spec: KPI_CARDS_ONCAL, cur: schedCur, prv: schedPrv, addSparkline: false, anchors
  });
  cardRow += adv(KPI_CARDS_ONCAL.length);

  // C) Changes to This Week’s Calendar
  dash.getRange(cardRow, 1).setValue(`Changes — ${label}`).setFontWeight('bold').setFontSize(12);
  renderScorecardsCompact_({
    sheet: dash, startRow: cardRow + 1, startCol: 1,
    cardsPerRow: 10, cardWidth: 2, cardHeight: 4, hGap: 1, vGap: 1,
    spec: KPI_CARDS_CHANGES, cur: schedCur, prv: schedPrv, addSparkline: false, anchors
  });
  cardRow += adv(KPI_CARDS_CHANGES.length);

  // D) Payments Snapshot
  dash.getRange(cardRow, 1).setValue('Payments Snapshot').setFontWeight('bold').setFontSize(12);
  renderScorecardsCompact_({
    sheet: dash, startRow: cardRow + 1, startCol: 1,
    cardsPerRow: 10, cardWidth: 2, cardHeight: 4, hGap: 1, vGap: 1,
    spec: KPI_CARDS_PAYMENTS, cur, prv, addSparkline: false, anchors
  });
  cardRow += adv(KPI_CARDS_PAYMENTS.length);

  // E) At‑Risk Snapshot (directly under Payments Snapshot)
  dash.getRange(cardRow, 1).setValue(`At‑Risk Snapshot (as of ${Utilities.formatDate(end, Session.getScriptTimeZone()||'GMT','yyyy-MM-dd')})`)
  .setFontWeight('bold').setFontSize(12);
  renderScorecardsCompact_({
    sheet: dash, startRow: cardRow + 1, startCol: 1,
    cardsPerRow: 10, cardWidth: 2, cardHeight: 4, hGap: 1, vGap: 1,
    spec: KPI_CARDS_RISK, cur, prv, addSparkline: false, anchors
  });
  cardRow += adv(KPI_CARDS_RISK.length);

}

/** === KPI computation (date-window based, as-of = window end) ============ */
function computeKpis_(winStart, winEnd, brandFilter, repFilter, master, masterHeader, metrics, metricsHeader) {
  const iRoot = findCol_(masterHeader, ['RootApptID','Root Appt ID','ROOT','Root_ID']);
  const iBrand= findCol_(masterHeader, ['Brand']);
  const iRep  = findCol_(masterHeader, ['Assigned Rep','AssignedRep','Rep','Sales Rep']);
  const iVType= findCol_(masterHeader, ['Visit Type','VisitType','Type']);
  const iVDate= findCol_(masterHeader, ['Visit Date','Visit_Date','Appt Date','Appointment Date']);

  const asOf = winEnd;

  const vtOf     = (r) => String(r[iVType]||'').trim().toLowerCase();
  const inRange  = (d) => { const t=asDate_(d); return t && t>=winStart && t<=winEnd; };
  const byBrandRep = (r) =>
    (!brandFilter || String(r[iBrand]||'').trim()===brandFilter) &&
    (!repFilter   || String(r[iRep]  ||'').trim()===repFilter);

  // === Appointment row counts (from master rows in range) ===
  const rowsInRange = master.filter(r => inRange(r[iVDate]) && byBrandRep(r));
  const consultations   = rowsInRange.filter(r => vtOf(r)==='appointment').length;
  const dvsInWindow     = rowsInRange.filter(r => vtOf(r)==='diamond viewing').length;
  const totalAppointments = consultations + dvsInWindow;
  const totalCustomers  = uniq_(rowsInRange.map(r => r[iRoot])).length;

  // === Brand/Rep filter set for ledger lookups (roots) =====================
  const allowedRootsSet = new Set(master.filter(r => byBrandRep(r)).map(r => r[iRoot]));

  // === Metrics header index + helpers =====================================
  const xi = makeIdx_(metricsHeader);
  const isBrand = (row) => !brandFilter || String(row[xi['Brand']]).trim() === brandFilter;
  const isRep   = (row) => !repFilter   || String(row[xi['Assigned Rep']]).trim() === repFilter;

  // === Cohort (History only: count roots whose FIRST visit is in window and have ≥2 visits) ===
  const cohortSecondAppt = metrics.filter(r =>
    isBrand(r) && isRep(r) &&
    r[xi['First Visit Date']] instanceof Date &&
    r[xi['First Visit Date']] >= winStart && r[xi['First Visit Date']] <= winEnd &&
    truthy(r[xi['Has 2nd Appt?']])
  ).length;

  // === First-time deposits in window ======================================
  const firstDepRowsInWin = metrics.filter(r => {
    if (!(isBrand(r) && isRep(r))) return false;
    const k = r[xi['Deposit Date (First Pay)']];
    return k instanceof Date && k >= winStart && k <= winEnd;
  });
  const depositsInWindow = firstDepRowsInWin.length;

  // AOV among first-time deposits
  const aov = avg_(
    firstDepRowsInWin
      .map(r => Number(r[xi['Order Total']]))
      .filter(n => isFinite(n) && n > 0)
  );

  // Sum of Order Total among first-time deposits
  const firstDepositsSum = sum_(
    firstDepRowsInWin.map(r => Number(r[xi['Order Total']]) || 0)
  );

  // Median 1st Appt→Dep in CALENDAR days among first-time deposits (allow 0-day)
  const diffsDays = firstDepRowsInWin.map(r => {
    const fv  = r[xi['First Visit Date']];           // << switch from First DV to First Visit
    const dep = r[xi['Deposit Date (First Pay)']];
    if (!(fv instanceof Date) || !(dep instanceof Date)) return NaN;
    return Math.max(0, Math.round(daysBetween_(dep, fv))); // 0,1,2,...
  }).filter(n => isFinite(n) && n >= 0);
  const medianDvToDepDays = median_(diffsDays);       // keep same variable name used by cards

  // First-time deposit AMOUNT ($) — from ledger (first valid receipt per root)
  const firstPayMap = fetchFirstPaymentMapFromPayments_(); // Map<root,{d,amt,so}>
  const firstDepositsAmount = sum_(
    firstDepRowsInWin.map(r => {
      const root = r[xi['RootApptID']];
      const fp = firstPayMap.get(root);
      return fp ? Number(fp.amt) || 0 : 0;
    })
  );

  // === Median 1st Appt→Dep in BUSINESS days (kept for History block) ======
  const medianDvToDep = median_(                               // keep name used later
    metrics
      .filter(r => isBrand(r) && isRep(r) && isActiveAsOf_(r, xi, asOf))
      .map(r => Number(r[xi['Lead→Deposit (biz days)']]))      // << switch to Lead→Deposit
      .filter(n => isFinite(n) && n >= 0)
  );

  // === Total deposits (all receipts) in window (ledger) ====================
  const totalDeposits = countAllDepositsInWindow_(winStart, winEnd, allowedRootsSet);

  // Total payments AMOUNT (all receipts) in window (ledger)
  const totalPaymentsAmount = sumAllDepositsAmountInWindow_(winStart, winEnd, allowedRootsSet);

  // === Weighted pipeline (kept for History/chart) ==========================
  const stageWeights = getStageWeightMap_();
  const weighted = sum_(
    metrics.filter(r => isBrand(r) && isRep(r))
           .map(r => computeWeightedForRow_(r, xi, asOf, stageWeights))
  );

  // === Risk (unchanged) ====================================================
  // No‑touch >48h (exclude Won/Lost Lead; remove 90‑day active limit)
  const noTouch48 = metrics.filter(r => {
    if (!(isBrand(r) && isRep(r))) return false;
    const stageRaw = r[xi['Sales Stage']];
    if (isWonOrLostStage_(stageRaw)) return false;        // ← NEW: exclude Won & Lost Lead
    return lastTouchHoursAgo_(r, xi, asOf) > 48;          // ← NEW: no isActiveAsOf_ check
  }).length;

  const dvNoDep7  = metrics.filter(r => {
    if (!(isBrand(r) && isRep(r) && isActiveAsOf_(r, xi, asOf))) return false;
    const dv = r[xi['First DV Date']];
    const depMade = truthy(r[xi['Deposit Made?']]);
    if (!(dv instanceof Date) || depMade) return false;
    return daysBetween_(asOf, dv) > 7;
  }).length;

  const wait3d = metrics.filter(r => {
    if (!(isBrand(r) && isRep(r) && isActiveAsOf_(r, xi, asOf))) return false;
    const req = r[xi['Last 3D Request Date']];
    const pending = truthy(r[xi['3D Pending?']]);
    return (req instanceof Date) && pending && daysBetween_(asOf, req) > 3;
  }).length;

  const overdue3d = metrics.filter(r => {
    if (!(isBrand(r) && isRep(r) && isActiveAsOf_(r, xi, asOf))) return false;
    const due     = r[xi['3D Deadline']];
    const pending = truthy(r[xi['3D Pending?']]);
    // only overdue if still pending and due date already passed
    return pending && (due instanceof Date) && startOfDay_(due) < startOfDay_(asOf);
  }).length;

  const overduePr = metrics.filter(r => {
    if (!(isBrand(r) && isRep(r) && isActiveAsOf_(r, xi, asOf))) return false;
    const due = r[xi['Prod Deadline']];
    return (due instanceof Date) && startOfDay_(due) < startOfDay_(asOf);
  }).length;

  const escal = metrics.filter(r => {
    if (!(isBrand(r) && isRep(r) && isActiveAsOf_(r, xi, asOf))) return false;
    const moves = Number(r[xi['# 3D Deadline Moves']]) || 0;
    const pmoves= Number(r[xi['# Prod Deadline Moves']]) || 0;
    const orderDate = r[xi['Order Date']];
    const orderAge = orderDate instanceof Date ? businessDaysInclusive_(orderDate, asOf) : 0;
    return (moves>=2 || pmoves>=2) && orderAge>28;
  }).length;

  return {
    // appointments
    totalCustomers, totalAppointments, consultations, dvsInWindow, cohortSecondAppt,

    // payments (for cards)
    depositsInWindow,
    firstDepositsSum,          // (kept for any downstream use)
    firstDepositsAmount,       // NEW: sum of first deposit $ in window
    medianDvToDepDays,         // NEW: calendar days (used by card)
    aov,
    totalDeposits,             // count of all receipts in window (kept)
    totalPaymentsAmount,       // NEW: $ sum of all receipts in window

    // keep these for History / charts
    medianDvToDep,             // business days (history)
    weighted,

    // risk
    noTouch48, dvNoDep7, wait3d, overdue3d, overduePr, escal
  };
}


/** helpers used above */
function isActiveAsOf_(row, xi, asOf) {
  if (!(asOf instanceof Date)) return false;                       // ← guard
  const last = row[xi['Last Touch (Root Index)']];
  return last instanceof Date ? (daysBetween_(asOf, last) <= 90) : false;
}

function lastTouchHoursAgo_(row, xi, asOf) {
  const last = row[xi['Last Touch (Root Index)']]; if (!(last instanceof Date)) return 99999;
  return (asOf.getTime() - last.getTime())/3600000;
}

/** Condensed 2×4 card renderer with optional sparkline + drill icon */
function renderScorecardsCompact_({
  sheet, startRow, startCol, cardsPerRow, cardWidth, cardHeight, hGap, vGap,
  spec, cur, prv, addSparkline, anchors
}) {
  const bg='#ffffff', border='#e0e0e0', titleColor='#616161', good='#2e7d32', bad='#c62828', neutral='#9e9e9e';

  // Get the drill sheet gid once, then reuse
  const drillSheet = SpreadsheetApp.getActive().getSheetByName('Drill_KPI');
  const drillGid = drillSheet ? drillSheet.getSheetId() : null;

  spec.forEach((kpi, idx) => {
    const rBlock = Math.floor(idx / cardsPerRow);
    const cBlock = idx % cardsPerRow;
    const r0 = startRow + rBlock * (cardHeight + vGap);
    const c0 = startCol + cBlock * (cardWidth + hGap);

    // Frame
    sheet.getRange(r0, c0, cardHeight, cardWidth)
      .setBackground(bg)
      .setBorder(true,true,true,true,false,false,border,SpreadsheetApp.BorderStyle.SOLID);

    // Title (merge over all but the last column so the icon cell stays unmerged)
    const titleCols = Math.max(1, cardWidth - 1);
    const titleRange = sheet.getRange(r0, c0, 1, titleCols).merge()
      .setValue(kpi.label)
      .setFontWeight('bold')
      .setFontSize(9)
      .setFontColor(titleColor)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW); // ← overflow label

    // 🔎 in the dedicated (unmerged) top-right cell (ensure no note on this cell)
    const iconCell = sheet.getRange(r0, c0 + cardWidth - 1, 1, 1).clearContent().clearNote();
    if (drillGid && anchors && anchors[kpi.drill]) {
      iconCell
        .setFormula(`=HYPERLINK("#gid=${drillGid}&range=${anchors[kpi.drill]}","🔎")`)
        .setHorizontalAlignment('right').setVerticalAlignment('middle').setFontSize(10);
    } else {
      iconCell.setValue('').setHorizontalAlignment('right').setVerticalAlignment('middle');
    }

    // Tooltip — ONLY on the label cell(s)
    const notes = kpiNotes_();
    if (notes[kpi.label]) titleRange.setNote(notes[kpi.label]);

    // Values + deltas
    const curVal = normNumber_(cur[kpi.key]);
    const prvVal = normNumber_(prv[kpi.key]);
    const delta = (isFiniteNumber_(curVal) && isFiniteNumber_(prvVal)) ? (curVal - prvVal) : null;
    const pct   = (delta !== null && prvVal !== 0) ? (delta / prvVal) : null;
    const arrow = (pct === null || Math.abs(pct) < 0.005) ? '•' : (delta >= 0 ? '▲' : '▼');
    const color = (pct === null || Math.abs(pct) < 0.005) ? neutral : ((delta >= 0) === !!kpi.upGood ? good : bad);

    // Row 2: current
    sheet.getRange(r0+1, c0, 1, cardWidth).merge()
      .setValue(curVal === '' ? '—' : curVal).setNumberFormat(kpi.fmt)
      .setFontSize(18).setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('bottom');

    // Row 3: delta / prev
    sheet.getRange(r0+2, c0, 1, 1)
      .setValue(`${arrow} ${delta===null?'—':fmtDelta_(delta,kpi.fmt)}${pct===null?'':(' '+fmtPct_(pct))}`)
      .setFontSize(9).setFontWeight('bold').setFontColor(color).setHorizontalAlignment('left').setVerticalAlignment('middle');

    sheet.getRange(r0+2, c0+1, 1, 1)
      .setValue(prvVal==='' ? 'Prev: —' : ('Prev: '+prvVal))
      .setNumberFormat(kpi.fmt).setFontSize(9).setFontColor(neutral).setHorizontalAlignment('right').setVerticalAlignment('middle');

    // Row 4: sparkline
    if (addSparkline) {
      const f = sparklineFormulaForKpi_(kpi.label);
      sheet.getRange(r0+3, c0, 1, cardWidth).merge()
        .setFormula(f).setHorizontalAlignment('left').setVerticalAlignment('middle');
    }
  });
}


/** === KPI HISTORY (00_Dashboard!AA5:AQ) =================================== */
function ensureHistoryBlock_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_DASH);
  const range = sh.getRange(HIST_HEADER_ROW, HIST_START_COL, 1, HIST_HEADERS.length);
  range.setValues([HIST_HEADERS]).setFontWeight('bold').setFontSize(9).setFontColor('#616161');
}

function snapshotKpisForHistory_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_DASH);
  ensureHistoryBlock_();

  const end   = asDate_(sh.getRange(CELL_DATE_E).getValue()); // D2
  const brand = String(sh.getRange(CELL_BRAND).getValue()||''); // H1
  const rep   = String(sh.getRange(CELL_REP).getValue()  ||''); // H2
  const start = asDate_(sh.getRange(CELL_DATE_S).getValue());   // B2

  // Recompute KPIs for snapshot (same filters)
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(SH_MASTER).getDataRange().getValues(); const mH = master.shift().map(x=>String(x||'').trim());
  const metrics= ss.getSheetByName(SH_METRICS).getDataRange().getValues(); const xH = metrics.shift().map(x=>String(x||'').trim());
  const k = computeKpis_(start, end, brand, rep, master, mH, metrics, xH, end);

  // Check if snapshot (end, brand, rep) already exists
  const firstDataRow = HIST_HEADER_ROW + 1;
  const lastRow = sh.getLastRow();
  const rowsToCheck = Math.max(0, lastRow - firstDataRow + 1);
  if (rowsToCheck > 0) {
    const trio = sh.getRange(firstDataRow, HIST_START_COL, rowsToCheck, 3).getValues(); // Date|Brand|Rep
    const exists = trio.some(r =>
      r[0] instanceof Date &&
      startOfDay_(r[0]).getTime() === startOfDay_(end).getTime() &&
      String(r[1]||'') === brand &&
      String(r[2]||'') === rep
    );
    if (exists) return;
  }
  if (!(start instanceof Date) || !(end instanceof Date)) {
    const mon = mondayOfWeek_(new Date());
    const sun = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate()+6);
    sh.getRange(CELL_DATE_S).setValue(mon).setNumberFormat('yyyy-mm-dd'); // B2
    sh.getRange(CELL_DATE_E).setValue(sun).setNumberFormat('yyyy-mm-dd'); // D2
  }

  // Append new row (pad to AA-1 with blanks)
  const row = [
    end, brand, rep,
    k.totalAppointments, k.dvsInWindow,
    k.depositsInWindow, k.medianDvToDepDays, k.aov, k.weighted,
    k.noTouch48, k.dvNoDep7, k.wait3d, k.overdue3d, k.overduePr, k.escal
  ];

  sh.appendRow(new Array(HIST_START_COL-1).fill(''));
  const newRow = sh.getLastRow();
  sh.getRange(newRow, HIST_START_COL, 1, HIST_HEADERS.length).setValues([row]);
  sh.getRange(newRow, HIST_START_COL).setNumberFormat('yyyy-mm-dd');
}

function sparklineFormulaForKpi_(kpiLabel) {
  const startL = columnToLetter(HIST_START_COL);
  const endL   = columnToLetter(HIST_START_COL + HIST_HEADERS.length - 1);
  const brandL = columnToLetter(HIST_START_COL + 1);
  const repL   = columnToLetter(HIST_START_COL + 2);
  return `
=IFERROR(
  SPARKLINE(
    TAKE(
      FILTER(
        INDEX($${startL}$6:$${endL}, 0, MATCH("${kpiLabel}", $${startL}$5:$${endL}$5, 0)),
        ($${brandL}$6:$${brandL}=$${CELL_BRAND})+($${CELL_BRAND}=""),
        ($${repL}$6:$${repL}=$${CELL_REP})+($${CELL_REP}="")
      ),
      12
    ),
    {"charttype","line";"linewidth",1}
  ),
"")`.trim();
}

/** === STAGE WEIGHTS CONFIG (00_Dashboard!AS30:AT60) ======================= */
function ensureStageWeightConfigBlock_() {
  const dash = SpreadsheetApp.getActive().getSheetByName(SH_DASH);
  const rng = dash.getRange('AS30:AT36');
  const vals = rng.getValues();
  const empty = vals.flat().every(v => v === '' || v === null);
  if (empty) {
    rng.clearContent();
    const rows = [
      ['Stage','Weight'],
      ['LEAD',               0.10],
      ['HOT LEAD',           0.20],
      ['CONSULT',            0.30],
      ['DIAMOND VIEWING',    0.50],
      ['DEPOSIT',            0.90],
      ['ORDER COMPLETED',    0.00],
    ];
    dash.getRange('AS30:AT36').offset(0,0,rows.length,2).setValues(rows);
  }
}

function getStageWeightMap_() {
  const dash = SpreadsheetApp.getActive().getSheetByName(SH_DASH);
  const range = dash.getRange('AS30:AT60').getValues().filter(r => (r[0]||'').toString().trim());
  const map = {};
  for (const r of range) {
    const name = String(r[0]||'').trim().toUpperCase();
    const w = Number(r[1]);
    if (name && isFinite(w)) map[name] = w;
  }
  if (Object.keys(map).length===0) {
    return { 'LEAD':0.10, 'HOT LEAD':0.20, 'CONSULT':0.30, 'DIAMOND VIEWING':0.50, 'DEPOSIT':0.90, 'ORDER COMPLETED':0.00 };
  }
  return map;
}

/** === EXTERNAL MAPS (no 99_* tabs) ======================================== */
function fetchFirstPaymentMapFromPayments_() {
  const LEDGER_FILE_ID = PropertiesService.getScriptProperties().getProperty('PAYMENTS_400_FILE_ID');
  if (!LEDGER_FILE_ID) throw new Error('Missing script property: PAYMENTS_400_FILE_ID');
  const lss = SpreadsheetApp.openById(LEDGER_FILE_ID);
  const pay = lss.getSheetByName('Payments');
  if (!pay) throw new Error('Payments tab not found in 400 ledger.');

  const data = pay.getDataRange().getValues(); if (data.length < 2) return new Map();
  const header = data.shift().map(h => String(h||'').trim());
  const iRoot = findCol_(header, ['RootApptID','Root Appt ID','ROOT','Root_ID']);
  const iSO   = findCol_(header, ['SO#','SO Number','SO No','Sales Order #'], false);
  const iWhen = findCol_(header, ['PaymentDateTime','Payment DateTime','Payment Date','Paid At']);
  const iDoc  = findCol_(header, ['DocType','Document Type']);
  const iNet  = findCol_(header, ['AmountNet',' AmountNet','Net','Net Amount']);
  const iStat = findCol_(header, ['DocStatus','Status'], false);

  const firstByRoot = new Map();
  for (const r of data) {
    const root = r[iRoot], when = r[iWhen], doc = String(r[iDoc]||''), net = Number(r[iNet])||0;
    const stat = (iStat >= 0 ? String(r[iStat]||'') : '');
    if (!root || !(when instanceof Date)) continue;
    if (!/receipt/i.test(doc)) continue;
    if (!(net > MIN_FIRST_DEPOSIT_NET)) continue; // ignore tiny DV/3D holds
    if (/void|reversed|cancel/i.test(stat)) continue;
    const prev = firstByRoot.get(root);
    if (!prev || when < prev.d) firstByRoot.set(root, { d: when, amt: net, so: (iSO>=0 ? r[iSO] : '') });
  }
  return firstByRoot; // Map<root,{d,amt,so}>
}

function fetch3DStatusMapAsMap_() {
  const ss = SpreadsheetApp.getActive();

  // --- Master: map APPT_ID → RootApptID
  const master = ss.getSheetByName(SH_MASTER);
  const M = master.getDataRange().getValues(); const MH = M.shift().map(x=>String(x||'').trim());
  const iApptM = findCol_(MH, ['APPT_ID','Appt ID','APPTID','Appointment ID']);
  const iRootM = findCol_(MH, ['RootApptID','Root Appt ID','ROOT','Root_ID']);
  const apptToRoot = new Map();
  for (const r of M) {
    const appt = r[iApptM], root = r[iRootM];
    if (appt && root && !apptToRoot.has(appt)) apptToRoot.set(appt, root);
  }

  // --- Status Log: scan chronologically
  const log = ss.getSheetByName(SH_STATUSLOG);
  const L = log.getDataRange().getValues(); const LH = L.shift().map(x=>String(x||'').trim());
  const iApptL = findCol_(LH, ['APPT_ID','Appt ID','APPTID','Appointment ID']);
  const iCus   = findCol_(LH, ['Custom Order Status','Status','Order Status']);
  const iWhen  = findCol_(LH, ['Updated At','UpdatedAt','Updated Time','Timestamp']);

  const rows = [];
  for (const r of L) {
    const appt = r[iApptL];
    const root = apptToRoot.get(appt);
    const st   = String(r[iCus]||'').trim().toLowerCase();
    const when = r[iWhen] instanceof Date ? r[iWhen] : null;
    if (!root || !st || !when) continue;
    rows.push({ root, status: st, when });
  }
  rows.sort((a,b)=>a.when - b.when);

  // Phrase patterns (edit here if your wording changes)
  // Requests: open work awaiting delivery/response
  const D3_REQUEST_PATTERNS = /^(3d requested|3d revision requested)\b/;
  // Resolved: any state indicating 3D was delivered and/or moved past approval
  // Feel free to add variants you actually use in the log.
  const D3_RESOLVE_PATTERNS = /^(3d received|3d waiting approval|3d approved|approved for production|waiting production timeline|in production)\b/;

  // --- Build last request / last resolution (then derive pending)
  const state = new Map(); // root -> { req:Date|null, res:Date|null, pending:boolean }

  // Walk chronologically (L is already sorted above)
  for (const x of rows) {
    const cur = state.get(x.root) || { req: null, res: null, pending: false };
    if (D3_REQUEST_PATTERNS.test(x.status)) cur.req = x.when;
    if (D3_RESOLVE_PATTERNS.test(x.status)) cur.res = x.when;
    state.set(x.root, cur);
  }

  // Derive pending from the most-recent pair:
  // pending iff there is a request and either no resolve or resolve < request.
  for (const cur of state.values()) {
    cur.pending = !!(cur.req && (!cur.res || cur.res < cur.req));
  }

  return state;
}


/** === HELPERS ============================================================== */
// Count all receipt-type deposits in window (any payment, not just first-time).
function countAllDepositsInWindow_(winStart, winEnd, allowedRootsSet) {
  const LEDGER_FILE_ID = PropertiesService.getScriptProperties().getProperty('PAYMENTS_400_FILE_ID');
  if (!LEDGER_FILE_ID) return 0; // fail-soft if not configured
  const lss = SpreadsheetApp.openById(LEDGER_FILE_ID);
  const pay = lss.getSheetByName('Payments');
  if (!pay) return 0;

  const data = pay.getDataRange().getValues();
  if (data.length < 2) return 0;

  const iRoot = findCol_(data[0], ['RootApptID','Root Appt ID','ROOT','Root_ID']);
  const iWhen = findCol_(data[0], ['PaymentDateTime','Payment DateTime','Payment Date','Paid At']);
  const iDoc  = findCol_(data[0], ['DocType','Document Type']);
  const iNet  = findCol_(data[0], ['AmountNet',' AmountNet','Net','Net Amount']);
  const iStat = findCol_(data[0], ['DocStatus','Status'], false);

  let cnt = 0;
  for (let r = 1; r < data.length; r++) {
    const row  = data[r];
    const root = row[iRoot];
    const when = row[iWhen];
    const doc  = String(row[iDoc]||'');
    const net  = Number(row[iNet])||0;
    const stat = (iStat >= 0 ? String(row[iStat]||'') : '');

    if (!root || !(when instanceof Date)) continue;
    if (!allowedRootsSet.has(root)) continue;
    if (!/receipt/i.test(doc)) continue;
    if (!(net > 0)) continue;
    if (/void|reversed|cancel/i.test(stat)) continue;
    if (when < winStart || when > winEnd) continue;

    cnt++;
  }
  return cnt;
}

function sumAllDepositsAmountInWindow_(winStart, winEnd, allowedRootsSet) {
  const LEDGER_FILE_ID = PropertiesService.getScriptProperties().getProperty('PAYMENTS_400_FILE_ID');
  if (!LEDGER_FILE_ID) return 0; // fail-soft
  const lss = SpreadsheetApp.openById(LEDGER_FILE_ID);
  const pay = lss.getSheetByName('Payments');
  if (!pay) return 0;

  const data = pay.getDataRange().getValues();
  if (data.length < 2) return 0;

  const header = data[0].map(h => String(h||'').trim());
  const iRoot = findCol_(header, ['RootApptID','Root Appt ID','ROOT','Root_ID']);
  const iWhen = findCol_(header, ['PaymentDateTime','Payment DateTime','Payment Date','Paid At']);
  const iDoc  = findCol_(header, ['DocType','Document Type']);
  const iNet  = findCol_(header, ['AmountNet',' AmountNet','Net','Net Amount']);
  const iStat = findCol_(header, ['DocStatus','Status'], false);

  let sum = 0;
  for (let r = 1; r < data.length; r++) {
    const row  = data[r];
    const root = row[iRoot];
    const when = row[iWhen];
    const doc  = String(row[iDoc]||'');
    const net  = Number(row[iNet])||0;
    const stat = (iStat >= 0 ? String(row[iStat]||'') : '');
    if (!root || !(when instanceof Date)) continue;
    if (!allowedRootsSet.has(root)) continue;
    if (!/receipt/i.test(doc)) continue;
    if (!(net > 0)) continue;
    if (/void|reversed|cancel/i.test(stat)) continue;
    if (when < winStart || when > winEnd) continue;
    sum += net;
  }
  return sum;
}


function makeIdx_(headers) {
  const m = {}; headers.forEach((h,i)=>{ m[String(h).trim()] = i; }); return m;
}
function findCol_(headers, candidates, required=true) {
  const norm = headers.map(h=>String(h||'').trim().toLowerCase());
  for (const want of candidates) {
    const w = String(want).trim().toLowerCase();
    const i = norm.indexOf(w); if (i >= 0) return i;
  }
  for (let i=0;i<norm.length;i++) {
    for (const want of candidates) {
      const w = String(want).trim().toLowerCase();
      if (norm[i].indexOf(w) >= 0) return i;
    }
  }
  if (required) throw new Error('Missing column: ' + candidates.join(' / '));
  return -1;
}
function asDate_(v){ return v instanceof Date ? v : (v ? new Date(v) : null); }
function startOfDay_(d){ return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function mondayOfWeek_(d) {
  const t = startOfDay_(d), day = t.getDay(), diff = (day === 0 ? -6 : 1 - day);
  return new Date(t.getFullYear(), t.getMonth(), t.getDate() + diff);
}
function isNum_(v){ return v!=='' && v!==null && !isNaN(Number(v)); }
function num_(v){ return Number(v); }
function uniq_(a){ return Array.from(new Set(a)); }
function truthy(v){ return v===true || v===1 || String(v).toLowerCase()==='true'; }
/** Business days inclusive (NETWORKDAYS-like, weekends only) */
function businessDaysInclusive_(start, end) {
  if (!(start instanceof Date) || !(end instanceof Date)) return '';
  const s = startOfDay_(start), e = startOfDay_(end);
  if (e < s) return 0;
  let days = 0, d = new Date(s);
  while (d <= e) {
    const wd = d.getDay();
    if (wd !== 0 && wd !== 6) days++;
    d = new Date(d.getFullYear(), d.getMonth(), d.getDate()+1);
  }
  return days;
}

function normalizeStage_(v) {
  if (!v) return '';
  const s = String(v).trim().toUpperCase();

  // Collapse all appointment-like statuses into APPOINTMENT
  // (e.g., "Appointment", "Consultation", "Consult", "Diamond Viewing" etc.)
  if (s.includes('APPOINT') || s.includes('CONSULT') || s.includes('DIAMOND')) return 'APPOINTMENT';

  if (s.includes('HOT')) return 'HOT LEAD';                 // Hot lead
  if (s.includes('DEPOSIT')) return 'DEPOSIT';              // Deposit
  if (s.includes('ORDER') && s.includes('COMPLET')) return 'ORDER COMPLETED'; // Order Completed
  if (s.includes('LEAD')) return 'LEAD';                    // plain Lead

  // Anything else is not part of our simplified taxonomy
  return '';
}


function computeBudgetBand_(bMax) {
  if (!isNum_(bMax)) return '';
  const x = Number(bMax);
  if (x < 3000) return '<3k';
  if (x <= 5000) return '3–5k';
  if (x <= 8000) return '5–8k';
  if (x <= 12000) return '8–12k';
  return '12k+';
}
function sum_(arr) { return (arr || []).reduce((s, v) => s + (Number(v) || 0), 0); }
function avg_(arr) {
  const nums = (arr || []).map(Number).filter(n => isFinite(n));
  return nums.length ? sum_(nums) / nums.length : '';
}
function median_(arr) {
  const nums = (arr || []).map(Number).filter(n => isFinite(n)).sort((a,b)=>a-b);
  const n = nums.length; if (!n) return '';
  return n % 2 ? nums[(n-1)/2] : (nums[n/2 - 1] + nums[n/2]) / 2;
}
function columnToLetter(col) {
  let temp = '', letter = ''; while (col > 0) {
    temp = (col - 1) % 26; letter = String.fromCharCode(temp + 65) + letter; col = (col - temp - 1) / 26;
  } return letter;
}
function normNumber_(v){ if(v===''||v===null||v===undefined)return ''; const n=Number(v); return isFinite(n)?n:''; }
function isFiniteNumber_(v){ return typeof v==='number' && isFinite(v); }
function fmtPct_(p){ if(p===null)return''; const s=p>=0?'+':''; return `${s}${(p*100).toFixed(1)}%`; }
function fmtDelta_(d,fmt){ const s=d>=0?'+':''; if(/\$/.test(fmt)) return `${s}$${Math.round(Math.abs(d)).toLocaleString()}`; return `${s}${Math.round(d).toLocaleString()}`; }

function daysBetween_(a,b){ return (startOfDay_(a)-startOfDay_(b))/86400000; }

function computeWeightedForRow_(row, xi, asOfDate, stageWeights) {
  const stage = String(row[xi['Sales Stage']]||'').trim();
  const stageNorm = normalizeStage_(stage);
  const w = stageWeights[stageNorm] ?? 0;

  const orderTotal = Number(row[xi['Order Total']]) || 0;
  const bMin = Number(row[xi['Budget Min']]) || NaN;
  const bMax = Number(row[xi['Budget Max']]) || NaN;
  const budgetMid = (isFinite(bMin) && isFinite(bMax)) ? (bMin + bMax)/2 : 0;
  const valueForWeight = orderTotal > 0 ? orderTotal : budgetMid;

  const lastTouch = row[xi['Last Touch (Root Index)']];
  const active90 = lastTouch instanceof Date ? (daysBetween_(asOfDate, lastTouch) <= 90) : false;

  return active90 ? (w * valueForWeight) : 0;
}

function applyCardColumnWidths_({sheet, startCol, cardsPerRow, cardWidthCols, gutterCols, cardColPx}) {
  const totalCols = cardsPerRow * (cardWidthCols + gutterCols) - gutterCols; // no gutter after last card
  const gutterPx = Math.max(1, Math.round(cardColPx / 2)); // ¼ of card (2 columns * W) -> W/2

  for (let i = 0; i < totalCols; i++) {
    const inBlock = i % (cardWidthCols + gutterCols);
    const col = startCol + i;
    if (inBlock < cardWidthCols) {
      sheet.setColumnWidth(col, cardColPx);
    } else {
      sheet.setColumnWidth(col, gutterPx);
    }
  }
}

/**
 * Weekly scheduling KPIs (Created vs On-Calendar vs Changes vs Unique).
 * Window: winStart..winEnd. Brand/Rep filters applied on each row.
 */
function computeScheduleKpis_(winStart, winEnd, brandFilter, repFilter, master, mH) {
  const iRoot   = findCol_(mH, ['RootApptID','Root Appt ID','ROOT','Root_ID']);
  const iBrand  = findCol_(mH, ['Brand']);
  const iRep    = findCol_(mH, ['Assigned Rep','AssignedRep','Rep','Sales Rep']);
  const iVType  = findCol_(mH, ['Visit Type','VisitType','Type']);
  const iVDate  = findCol_(mH, ['Visit Date','Visit_Date','Appt Date','Appointment Date']);
  const iStatus = findCol_(mH, ['Status','Appt Status','Appointment Status'], false);

  const iBookedAt   = findCol_(mH, ['Booked At (ISO)','Booked At','BookedAt','Created At','CreatedAt'], false);
  const iCanceledAt = findCol_(mH, ['CanceledAt','CancelledAt','Canceled At','Cancelled At'], false);
  const iResTo      = findCol_(mH, ['RescheduledToUID','Rescheduled To UID','ReschedToUID','Rescheduled To'], false);
  const iResFrom    = findCol_(mH, ['RescheduledFromUID','Rescheduled From UID','ReschedFromUID','Rescheduled From'], false);

  const iCust = findCol_(mH, ['Customer Name','Customer','Client Name'], false);

  const inWinDate   = (d) => { const t=asDate_(d); return t && t>=winStart && t<=winEnd; };
  const matchesBR   = (r) => (!brandFilter || String(r[iBrand]||'').trim()===brandFilter) &&
                              (!repFilter   || String(r[iRep]  ||'').trim()===repFilter);

  const vt         = (r) => String(r[iVType]||'').trim().toLowerCase();
  const isApptType = (r) => ['appointment','diamond viewing'].includes(vt(r));
  const isConsult  = (r) => vt(r)==='appointment';
  const isDV       = (r) => vt(r)==='diamond viewing';
  const isScheduled= (r) => !iStatus ? true : (String(r[iStatus]||'').trim().toLowerCase()==='scheduled');

  // --- A) Created This Week (by Booked At / CanceledAt)
  let bookingsCreated=0, reschedulesCreated=0, cancelsCreated=0;

  if (iBookedAt>=0) {
    for (const r of master) {
      if (!matchesBR(r)) continue;
      if (!inWinDate(r[iBookedAt])) continue;
      if (!isScheduled(r)) continue;
      const fromUID = iResFrom>=0 ? String(r[iResFrom]||'').trim() : '';
      if (fromUID) reschedulesCreated++;
      else bookingsCreated++;
    }
  }
  if (iCanceledAt>=0) {
    for (const r of master) {
      if (!matchesBR(r)) continue;
      const cAt = r[iCanceledAt];
      if (!inWinDate(cAt)) continue;
      const toUID = iResTo>=0 ? String(r[iResTo]||'').trim() : '';
      if (!toUID) cancelsCreated++; // exclude reschedules (those have RescheduledToUID)
    }
  }

  // --- B) On the Calendar This Week (by Visit Date, Status=Scheduled)
  const rowsSchedInWin = master.filter(r => matchesBR(r) && inWinDate(r[iVDate]) && isScheduled(r) && isApptType(r));
  const apptsScheduled         = rowsSchedInWin.length;
  const consultationsScheduled = rowsSchedInWin.filter(isConsult).length;
  const dvsScheduled           = rowsSchedInWin.filter(isDV).length;

  // --- C) Changes to This Week’s Calendar (removals from this week)
  let reschedOffThisWeek=0, cancelledFromThisWeek=0;
  const rowsWithVisitInWin = master.filter(r => matchesBR(r) && inWinDate(r[iVDate]));
  for (const r of rowsWithVisitInWin) {
    const toUID = iResTo>=0 ? String(r[iResTo]||'').trim() : '';
    const cAt   = iCanceledAt>=0 ? r[iCanceledAt] : null;
    if (toUID && cAt instanceof Date) reschedOffThisWeek++; // original slot vacated this week due to reschedule
    if (!toUID && (cAt instanceof Date)) cancelledFromThisWeek++; // true cancel of a slot in this week
  }

  // --- D) Unique Customers (Scheduled)
  const roots = new Set(rowsSchedInWin.map(r => r[iRoot]));
  const uniqueCustomersScheduled = roots.size;

  return {
    bookingsCreated, reschedulesCreated, cancelsCreated,
    apptsScheduled, consultationsScheduled, dvsScheduled,
    reschedOffThisWeek, cancelledFromThisWeek,
    uniqueCustomersScheduled
  };
}


function kpiNotes_() {
  return {
    'Total Appointments': 'Count of ALL appointment rows whose Visit Date falls inside the window (after Brand/Rep filters).',
    'Consultations': 'Unique RootApptIDs that had a Consultation visit inside the window. (Type contains “consult”, excluding Diamond Viewing.)',
    'Diamond Viewings': 'Unique RootApptIDs that had a Diamond Viewing visit inside the window.',
    'Deposits (first-time)': 'Count of Roots whose FIRST deposit date is inside the window.',
    'Median 1st Appt→Dep (days)': 'Among roots with a FIRST deposit in the window, the median CALENDAR days from First Appointment to that deposit (0 means same-day).',
    'Average Order Value': 'Average “Order Total” among Roots whose FIRST deposit date is inside the window.',
    'Weighted Pipeline': 'As-of window end: sum over ACTIVE (≤90d) Roots of [Stage Weight × Value], where Value is Order Total if present, else Budget Midpoint.',
    'No‑touch >48h': 'As-of window end: Roots (excluding Sales Stage = Won or Lost Lead) whose last touch is more than 48 hours old.',
    'DV no deposit >7d': 'As-of window end: ACTIVE Roots that had a DV but no deposit, and the DV is >7 days old.',
    '3D wait >3d': 'As-of window end: ACTIVE Roots whose latest 3D request is pending and older than 3 days.',
    '3D overdue': 'As-of window end: ACTIVE Roots with a 3D deadline before the window end date.',
    'Production overdue': 'As-of window end: ACTIVE Roots with a production deadline before the window end date.',
    'Ops escalation': 'As-of window end: ACTIVE Roots with ≥2 deadline moves (3D or Production) and order age >28 business days.',
    'Bookings Created': 'Booked At (ISO) in window; Status=Scheduled; RescheduledFromUID blank (true new bookings).',
    'Reschedules Created': 'Booked At (ISO) of the new timeslot in window; Status=Scheduled; RescheduledFromUID filled.',
    'Cancels': 'CanceledAt in window, excluding reschedules (RescheduledToUID blank).',
    'Appointments (Scheduled)': 'Visit Date in window; Status=Scheduled; Visit Type in {Appointment, Diamond Viewing}.',
    'Consultations': 'Visit Date in window; Status=Scheduled; Visit Type=Appointment.',
    'Diamond Viewings': 'Visit Date in window; Status=Scheduled; Visit Type=Diamond Viewing.',
    'Rescheduled Off This Week': 'Original Visit Date in window; row has RescheduledToUID and CanceledAt (slot vacated due to reschedule).',
    'Cancelled From This Week': 'Original Visit Date in window; CanceledAt present; RescheduledToUID blank (true cancel).',
    'Unique Customers (Scheduled)': 'Distinct RootApptID among scheduled Consultations + Diamond Viewings in the window.',
    'First-time Deposits $': 'Sum of net amounts for each customer’s first valid receipt where the first deposit date is inside the window (void/reversed excluded).',
    'All Payments $': 'Sum of net amounts of all valid receipt rows whose Paid At is inside the window (void/reversed excluded).',
  };
}


function writeCohortDebug() {
  const ss = SpreadsheetApp.getActive();
  const dash = ss.getSheetByName(SH_DASH);
  const start = asDate_(dash.getRange(CELL_DATE_S).getValue()); // B2
  const end   = asDate_(dash.getRange(CELL_DATE_E).getValue()); // D2
  const brand = String(dash.getRange(CELL_BRAND).getValue()||'').trim(); // H1
  const rep   = String(dash.getRange(CELL_REP).getValue()  ||'').trim(); // H2

  const sh = ss.getSheetByName(SH_METRICS);
  const data = sh.getDataRange().getValues();
  const H = {}; data[0].forEach((h,i)=> H[String(h).trim()] = i);
  const rows = data.slice(1);

  const isBrand = r => !brand || String(r[H['Brand']]||'').trim() === brand;
  const isRep   = r => !rep   || String(r[H['Assigned Rep']]||'').trim() === rep;

  const list = rows.filter(r => {
    const fv = r[H['First Visit Date']];
    const vc = Number(r[H['Visit Count']]||0);
    return isBrand(r) && isRep(r) && fv instanceof Date && fv >= start && fv <= end && vc >= 2;
  }).map(r => [
    r[H['RootApptID']], r[H['Customer Name']], r[H['First Visit Date']], r[H['Visit Count']],
    r[H['Assigned Rep']], r[H['Brand']]
  ]);

  const out = ss.getSheetByName('Debug_Cohort2nd') || ss.insertSheet('Debug_Cohort2nd');
  out.clear();
  out.appendRow(['RootApptID','Customer Name','First Visit','Visit Count','Assigned Rep','Brand']);
  if (list.length) out.getRange(2,1,list.length,6).setValues(list);
  out.setFrozenRows(1);
}

/** === Drill sheet ========================================================= */
function rebuildKpiDrill_(start, end, brand, rep, master, mH, metrics, xH) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Drill_KPI') || ss.insertSheet('Drill_KPI');
  sh.clear();

  // Title + filters
  sh.appendRow(['Drill‑down for', Utilities.formatDate(start,'GMT','yyyy-MM-dd'), '→', Utilities.formatDate(end,'GMT','yyyy-MM-dd'),
                'Brand:', brand || '(all)', 'Rep:', rep || '(all)']);
  sh.getRange(1,1,1,8).setFontWeight('bold');

  // Helper indexes
  const iRoot = findCol_(mH, ['RootApptID','Root Appt ID','ROOT','Root_ID']);
  const iBrand= findCol_(mH, ['Brand']); const iRep=findCol_(mH, ['Assigned Rep','AssignedRep','Rep','Sales Rep']);
  const iVType= findCol_(mH, ['Visit Type','VisitType','Type']);
  const iVDate= findCol_(mH, ['Visit Date','Visit_Date','Appt Date','Appointment Date']);
  const iStatus = findCol_(mH, ['Status','Appt Status','Appointment Status'], false);
  const iBookedAt   = findCol_(mH, ['Booked At (ISO)','Booked At','BookedAt','Created At','CreatedAt'], false);
  const iCanceledAt = findCol_(mH, ['CanceledAt','CancelledAt','Canceled At','Cancelled At'], false);
  const iResTo      = findCol_(mH, ['RescheduledToUID','Rescheduled To UID','ReschedToUID','Rescheduled To'], false);
  const iResFrom    = findCol_(mH, ['RescheduledFromUID','Rescheduled From UID','ReschedFromUID','Rescheduled From'], false);
  const iCust       = findCol_(mH, ['Customer Name','Customer','Client Name'], false);
  const xi = makeIdx_(xH);


  // Shared "as-of" and active helper for all drill sections (used by Weighted + Risk)
  const asOf = end;
  const isActive = (r) => {
    const last = r[xi['Last Touch (Root Index)']];
    return last instanceof Date ? (daysBetween_(asOf, last) <= 90) : false;
  };

  const matchBR = (r) => (!brand || String(r[iBrand]||'').trim()===brand)&&(!rep||String(r[iRep]||'').trim()===rep);
  const inWin   = (d) => { const t=asDate_(d); return t && t>=start && t<=end; };

  // Keep rows that really count in the window (exclude cancels, no-shows, and rescheduled-off slots)
  const isFunnelRow = (r) => {
    if (!matchBR(r) || !inWin(r[iVDate])) return false;

    const s = iStatus < 0 ? '' : String(r[iStatus]||'').trim().toLowerCase();
    const cancelled  = /cancel/.test(s);            // canceled / cancelled
    const noShow     = /no[-\s]?show/.test(s);      // no show
    const reschedOff = (iResTo >= 0 && String(r[iResTo]||'').trim()!=='') &&
                      (iCanceledAt >= 0 && (r[iCanceledAt] instanceof Date));

    return !cancelled && !noShow && !reschedOff;
  };

  const anchors = {}; let row = 3;

  // --- appointments from master (rows in window)
  const inRange = master.filter(r=> inWin(r[iVDate]) && matchBR(r));

  const addSection = (title, key, rows, headers) => {
    sh.getRange(row,1).setValue(title).setFontWeight('bold').setFontSize(11);
    row++;
    sh.getRange(row,1,1,headers.length).setValues([headers]).setFontWeight('bold');
    row++;
    if (rows.length) sh.getRange(row,1,rows.length,headers.length).setValues(rows);
    anchors[key] = `A${row-2}`; // anchor at section header
    row += Math.max(1, rows.length) + 2;
  };

  // Total Customers (unique roots with any visit) — include Customer
  {
    const uniqueRoots = uniq_(inRange.map(r=>r[iRoot]));
    const rows = uniqueRoots.map(root => {
      const first = inRange.find(r=>r[iRoot]===root);
      const cust  = (iCust >= 0 && first) ? first[iCust] : '';
      return [root, cust, first ? first[iVDate] : '', first ? first[iVType] : '', first ? first[iRep] : '', first ? first[iBrand] : ''];
    });
    addSection(
      'Total Customers (unique roots with any visit in window)',
      'appts_total_customers',
      rows,
      ['RootApptID','Customer','First Visit in Window','First Visit Type','Rep','Brand']
    );
  }

  // Total Appointments (all 'Appointment' or 'Diamond Viewing' rows)
  {
    const rows = inRange
      .filter(r => ['appointment','diamond viewing'].includes(String(r[iVType]||'').trim().toLowerCase()))
      .map(r => [r[iRoot], (iCust>=0 ? r[iCust] : ''), r[iVDate], r[iVType], r[iRep], r[iBrand]]);
    addSection('Total Appointments (rows)', 'appts_total_appointments',
           rows, ['RootApptID','Customer','Visit Date','Visit Type','Rep','Brand']);
  }

  // Consultations
  {
    const rows = inRange
      .filter(r => String(r[iVType]||'').trim().toLowerCase()==='appointment')
      .map(r => [r[iRoot], (iCust>=0 ? r[iCust] : ''), r[iVDate], r[iVType], r[iRep], r[iBrand]]);
    addSection('Consultations (Appointment rows)', 'appts_consultations',
           rows, ['RootApptID','Customer','Visit Date','Visit Type','Rep','Brand']);
  }

  // Diamond Viewings
  {
    const rows = inRange
      .filter(r => String(r[iVType]||'').trim().toLowerCase()==='diamond viewing')
      .map(r => [r[iRoot], (iCust>=0 ? r[iCust] : ''), r[iVDate], r[iVType], r[iRep], r[iBrand]]);
    addSection('Diamond Viewings (rows)', 'appts_dv',
           rows, ['RootApptID','Customer','Visit Date','Visit Type','Rep','Brand']);
  }

// === Created This Week (inflow) ============================================
if (iBookedAt>=0) {
  // Bookings Created (Status=Scheduled AND RescheduledFromUID blank) by Booked At
  {
    const rows = master.filter(r => matchBR(r) &&
      (iStatus<0 || String(r[iStatus]||'').trim().toLowerCase()==='scheduled') &&
      (!iResFrom || String(r[iResFrom]||'').trim()==='') &&
      inWin(r[iBookedAt])
    ).map(r => [r[iRoot], (iCust>=0?r[iCust]:''), r[iBookedAt], r[iVDate], r[iVType], r[iRep], r[iBrand]]);
    addSection('Created This Week — Bookings (by Booked At)', 'inf_bookingsCreated',
               rows, ['RootApptID','Customer','Booked At','Visit Date','Visit Type','Rep','Brand']);
  }

  // Reschedules Created (Status=Scheduled AND RescheduledFromUID filled) by Booked At of NEW slot
  {
    const rows = master.filter(r => matchBR(r) &&
      (iStatus<0 || String(r[iStatus]||'').trim().toLowerCase()==='scheduled') &&
      (iResFrom>=0 && String(r[iResFrom]||'').trim()!=='') &&
      inWin(r[iBookedAt])
    ).map(r => [r[iRoot], (iCust>=0?r[iCust]:''), r[iBookedAt], r[iResFrom], r[iVDate], r[iVType], r[iRep], r[iBrand]]);
    addSection('Created This Week — Reschedules (by Booked At of new time)', 'inf_reschedulesCreated',
               rows, ['RootApptID','Customer','Booked At','RescheduledFromUID','New Visit Date','Visit Type','Rep','Brand']);
  }
}

// Cancels Created (CanceledAt; exclude reschedules)
if (iCanceledAt>=0) {
  const rows = master.filter(r => matchBR(r) &&
    inWin(r[iCanceledAt]) &&
    (!iResTo || String(r[iResTo]||'').trim()==='')
  ).map(r => [r[iRoot], (iCust>=0?r[iCust]:''), r[iCanceledAt], r[iVDate], (iStatus>=0?r[iStatus]:''), r[iVType], r[iRep], r[iBrand]]);
  addSection('Created This Week — Cancels (by CanceledAt)', 'inf_cancelsCreated',
             rows, ['RootApptID','Customer','CanceledAt','Original Visit Date','Status','Visit Type','Rep','Brand']);
}

// === On the Calendar This Week (Status=Scheduled) ==========================
{
  const schedRows = master.filter(r => matchBR(r) && inWin(r[iVDate]) &&
    (iStatus<0 || String(r[iStatus]||'').trim().toLowerCase()==='scheduled') &&
    (String(r[iVType]||'').trim()!=='')
  );

  // Appointments (Scheduled) = consult + DV
  {
    const rows = schedRows
      .filter(r => ['appointment','diamond viewing'].includes(String(r[iVType]||'').trim().toLowerCase()))
      .map(r => [r[iRoot], (iCust>=0?r[iCust]:''), r[iVDate], r[iVType], r[iRep], r[iBrand]]);
    addSection('On Calendar — Appointments (Scheduled)', 'cal_appts',
               rows, ['RootApptID','Customer','Visit Date','Visit Type','Rep','Brand']);
  }

  // Consultations (Scheduled)
  {
    const rows = schedRows
      .filter(r => String(r[iVType]||'').trim().toLowerCase()==='appointment')
      .map(r => [r[iRoot], (iCust>=0?r[iCust]:''), r[iVDate], r[iVType], r[iRep], r[iBrand]]);
    addSection('On Calendar — Consultations (Scheduled)', 'cal_consults',
               rows, ['RootApptID','Customer','Visit Date','Visit Type','Rep','Brand']);
  }

  // Diamond Viewings (Scheduled)
  {
    const rows = schedRows
      .filter(r => String(r[iVType]||'').trim().toLowerCase()==='diamond viewing')
      .map(r => [r[iRoot], (iCust>=0?r[iCust]:''), r[iVDate], r[iVType], r[iRep], r[iBrand]]);
    addSection('On Calendar — Diamond Viewings (Scheduled)', 'cal_dvs',
               rows, ['RootApptID','Customer','Visit Date','Visit Type','Rep','Brand']);
  }
}

  // === Changes to This Week’s Calendar (removals) ============================
  {
    // Rescheduled Off This Week (original slot vacated) — require RescheduledToUID & CanceledAt on the original row; Visit Date is in window
    const rowsReschedOff = master.filter(r => matchBR(r) && inWin(r[iVDate]) &&
      (iResTo>=0 && String(r[iResTo]||'').trim()!=='') &&
      (iCanceledAt>=0 && (r[iCanceledAt] instanceof Date))
    ).map(r => [r[iRoot], (iCust>=0?r[iCust]:''), r[iVDate], r[iCanceledAt], r[iResTo], (iStatus>=0?r[iStatus]:''), r[iVType], r[iRep], r[iBrand]]);
    addSection('Changes — Rescheduled Off This Week', 'chg_reschedOff',
              rowsReschedOff, ['RootApptID','Customer','Original Visit Date','CanceledAt','RescheduledToUID','Status','Visit Type','Rep','Brand']);

    // Cancelled From This Week (true cancels; exclude reschedules) — Visit Date in window
    const rowsCancelledFrom = master.filter(r => matchBR(r) && inWin(r[iVDate]) &&
      (iCanceledAt>=0 && (r[iCanceledAt] instanceof Date)) &&
      (!iResTo || String(r[iResTo]||'').trim()==='')
    ).map(r => [r[iRoot], (iCust>=0?r[iCust]:''), r[iVDate], r[iCanceledAt], (iStatus>=0?r[iStatus]:''), r[iVType], r[iRep], r[iBrand]]);
    addSection('Changes — Cancelled From This Week', 'chg_cancelledFrom',
              rowsCancelledFrom, ['RootApptID','Customer','Visit Date','CanceledAt','Status','Visit Type','Rep','Brand']);
  }

  // === Unique Customers (Scheduled) =========================================
  {
    const uniqRoots = new Set(master.filter(r => matchBR(r) && inWin(r[iVDate]) &&
      (iStatus<0 || String(r[iStatus]||'').trim().toLowerCase()==='scheduled') &&
      ['appointment','diamond viewing'].includes(String(r[iVType]||'').trim().toLowerCase())
    ).map(r => r[iRoot]));

    // Show a simple roster of those unique customers
    const rows = [];
    for (const root of uniqRoots) {
      const first = master.find(r => r[iRoot]===root && matchBR(r));
      rows.push([root, (iCust>=0 && first ? first[iCust] : ''), (first ? first[iRep] : ''), (first ? first[iBrand] : '')]);
    }
    addSection('Unique Customers (Scheduled)', 'uniq_customersSch',
              rows, ['RootApptID','Customer','Rep','Brand']);
  }

  // Payments sections
  {
    // Map<RootApptID, {d: Date, amt: number, so: string}>
    const firstPayMap = fetchFirstPaymentMapFromPayments_();

    const firstDeps = metrics
      .filter(r =>
        (!brand || String(r[xi['Brand']]).trim()===brand) &&
        (!rep   || String(r[xi['Assigned Rep']]).trim()===rep) &&
        r[xi['Deposit Date (First Pay)']] instanceof Date &&
        r[xi['Deposit Date (First Pay)']] >= start &&
        r[xi['Deposit Date (First Pay)']] <= end
      )
      .map(r => {
        const root = r[xi['RootApptID']];
        const fp   = firstPayMap.get(root);
        const amt  = fp ? Number(fp.amt) || 0 : '';
        return [
          root,
          r[xi['Customer Name']],
          r[xi['Deposit Date (First Pay)']],
          amt,                              // ← First Deposit Amount ($) from ledger
          r[xi['Order Total']]
        ];
      });

    addSection(
      'Deposits (first-time) in window',
      'pay_firstDeposits',
      firstDeps,
      ['RootApptID','Customer','First Deposit Date','First Deposit Amount ($)','Order Total']
    );
  }


  // Median 1st Appt→Dep (days) — first-time deposits in window
  {
    const firstDepRowsInWin = metrics.filter(r => {
      if (brand && String(r[xi['Brand']]).trim() !== brand) return false;
      if (rep   && String(r[xi['Assigned Rep']]).trim() !== rep) return false;
      const d = r[xi['Deposit Date (First Pay)']];
      return d instanceof Date && d >= start && d <= end;
    });

    const rows = firstDepRowsInWin.map(r => {
      const fv  = r[xi['First Visit Date']];
      const dep = r[xi['Deposit Date (First Pay)']];
      const days = (fv instanceof Date && dep instanceof Date)
        ? Math.max(0, Math.round(daysBetween_(dep, fv)))
        : '';
      return [r[xi['RootApptID']], r[xi['Customer Name']], fv, dep, days, r[xi['Order Total']]];
    });

    addSection('Median 1st Appt→Dep (days) — first-time deposits in window', 'pay_apptToDep',
              rows, ['RootApptID','Customer','First Visit','First Deposit','Appt→Dep (days)','Order Total']);
  }

  // Average Order Value — first-time deposits in window
  {
    const rows = metrics.filter(r => {
      if (brand && String(r[xi['Brand']]).trim() !== brand) return false;
      if (rep   && String(r[xi['Assigned Rep']]).trim() !== rep) return false;
      const d = r[xi['Deposit Date (First Pay)']];
      return d instanceof Date && d >= start && d <= end;
    }).map(r => [r[xi['RootApptID']], r[xi['Customer Name']], r[xi['Deposit Date (First Pay)']], r[xi['Order Total']]]);

    addSection('Average Order Value — first‑time deposits in window', 'pay_aov',
               rows, ['RootApptID','Customer','First Deposit Date','Order Total']);
  }

  // Deposits (all) — any receipt in the window (reads Payments ledger)
  {
    const LEDGER_FILE_ID = PropertiesService.getScriptProperties().getProperty('PAYMENTS_400_FILE_ID');
    if (LEDGER_FILE_ID) {
      const lss = SpreadsheetApp.openById(LEDGER_FILE_ID);
      const pay = lss.getSheetByName('Payments');
      if (pay) {
        const data = pay.getDataRange().getValues();
        const payH = data.shift().map(h => String(h||'').trim());
        const iRootP = findCol_(payH, ['RootApptID','Root Appt ID','ROOT','Root_ID']);
        const iWhenP = findCol_(payH, ['PaymentDateTime','Payment DateTime','Payment Date','Paid At']);
        const iDocP  = findCol_(payH, ['DocType','Document Type']);
        const iNetP  = findCol_(payH, ['AmountNet',' AmountNet','Net','Net Amount']);
        const iStatP = findCol_(payH, ['DocStatus','Status'], false);

        // restrict to current Brand/Rep's roots
        const allowedRoots = new Set(master.filter(r => matchBR(r)).map(r => r[iRoot]));

        // NEW: map RootApptID → Customer (first non-blank from Master)
        const customerByRoot = new Map();
        if (iCust >= 0) {
          for (const mr of master) {
            const root = mr[iRoot];
            const cust = mr[iCust];
            if (root && cust && !customerByRoot.has(root)) customerByRoot.set(root, cust);
          }
        }

        const rows = [];
        for (const r of data) {
          const root = r[iRootP];
          const when = r[iWhenP];
          const doc  = String(r[iDocP]||'');
          const net  = Number(r[iNetP])||0;
          const stat = iStatP>=0 ? String(r[iStatP]||'') : '';
          if (!root || !(when instanceof Date)) continue;
          if (!allowedRoots.has(root)) continue;
          if (!/receipt/i.test(doc)) continue;
          if (!(net > 0)) continue;
          if (/void|reversed|cancel/i.test(stat)) continue;
          if (when < start || when > end) continue;

          const cust = customerByRoot.get(root) || '';
          rows.push([root, cust, when, net, doc]); // ← Customer column added
        }
        // sort by Payment DateTime (now column index 2)
        rows.sort((a,b)=> a[2] - b[2]);

        addSection(
          'Deposits (all) — receipts in window',
          'pay_allDeposits',
          rows,
          ['RootApptID','Customer','Payment DateTime','AmountNet','DocType'] // ← header updated
        );
      }
    }
  }


  // Weighted pipeline (as-of window end) — include First Visit Date + Order Date
  {
    const stageWeights = getStageWeightMap_();
    const rowsWeighted = metrics
      .filter(r => {
        if (brand && String(r[xi['Brand']]).trim() !== brand) return false;
        if (rep   && String(r[xi['Assigned Rep']]).trim() !== rep) return false;
        // only rows active as-of window end (same definition used elsewhere)
        return isActive(r);
      })
      .map(r => {
        const root       = r[xi['RootApptID']];
        const customer   = r[xi['Customer Name']];
        const firstVisit = r[xi['First Visit Date']];
        const orderDate  = r[xi['Order Date']];
        const stageRaw   = String(r[xi['Sales Stage']] || '').trim();
        const stageNorm  = normalizeStage_(stageRaw);
        const w          = stageWeights[stageNorm] ?? 0;

        const orderTotal = Number(r[xi['Order Total']]) || 0;
        const bMin       = Number(r[xi['Budget Min']]);   // may be NaN
        const bMax       = Number(r[xi['Budget Max']]);   // may be NaN
        const budgetMid  = (isFinite(bMin) && isFinite(bMax)) ? (bMin + bMax)/2 : 0;
        const valueBasis = orderTotal > 0 ? orderTotal : budgetMid;
        const weighted   = w * valueBasis;

        const assigned   = r[xi['Assigned Rep']];
        const brandVal   = r[xi['Brand']];

        return [root, customer, firstVisit, orderDate, stageRaw, w, valueBasis, weighted, assigned, brandVal];
      })
      // keep only meaningful contributions
      .filter(row => Number(row[7]) > 0)
      // sort by Weighted desc
      .sort((a, b) => Number(b[7]) - Number(a[7]));

    addSection(
      'Weighted Pipeline (as-of window end)',
      'pay_weighted',
      rowsWeighted,
      ['RootApptID','Customer','First Visit','Order Date','Sales Stage','Stage Weight','Value Basis','Weighted','Rep','Brand']
    );
  }

  // Risk sections (examples) — uses shared asOf/isActive defined above

  const rowsNoTouch48 = metrics.filter(r => {
    if (brand && String(r[xi['Brand']]).trim() !== brand) return false;
    if (rep   && String(r[xi['Assigned Rep']]).trim() !== rep) return false;
    const stageRaw = r[xi['Sales Stage']];
    if (isWonOrLostStage_(stageRaw)) return false;        // ← NEW: exclude Won & Lost Lead
    return lastTouchHoursAgo_(r, xi, asOf) > 48;          // ← NEW: no 90‑day active limit
  }).map(r => [r[xi['RootApptID']], r[xi['Customer Name']], r[xi['Last Touch (Root Index)']], r[xi['Assigned Rep']]]);

  addSection('No‑touch >48h (as‑of window end)', 'risk_noTouch48', rowsNoTouch48, ['RootApptID','Customer','Last Touch','Rep']);

  // DV no deposit >7d
  const rowsDvNoDep7 = metrics.filter(r => {
    if (!isActive(r)) return false;
    if (brand && String(r[xi['Brand']]).trim() !== brand) return false;
    if (rep   && String(r[xi['Assigned Rep']]).trim() !== rep) return false;
    const dv = r[xi['First DV Date']];
    const depMade = truthy(r[xi['Deposit Made?']]);
    return (dv instanceof Date) && !depMade && daysBetween_(asOf, dv) > 7;
  }).map(r => [r[xi['RootApptID']], r[xi['Customer Name']], r[xi['First DV Date']], r[xi['Assigned Rep']]]);
  addSection('DV no deposit >7d (as-of window end)', 'risk_dvNoDep7', rowsDvNoDep7, ['RootApptID','Customer','First DV','Rep']);

  // 3D wait >3d (pending requests older than 3 days)
  {
    const rows3dWait = metrics.filter(r => {
      if (!isActive(r)) return false;
      if (brand && String(r[xi['Brand']]).trim() !== brand) return false;
      if (rep   && String(r[xi['Assigned Rep']]).trim() !== rep) return false;
      const req = r[xi['Last 3D Request Date']];
      const pending = truthy(r[xi['3D Pending?']]);
      return (req instanceof Date) && pending && daysBetween_(asOf, req) > 3;
    }).map(r => {
      const req = r[xi['Last 3D Request Date']];
      // Days over threshold (vs 3 days), clamp to >= 1
      const over = Math.max(1, daysBetween_(asOf, req) - 3);
      return [r[xi['RootApptID']], r[xi['Customer Name']], req, over, r[xi['Assigned Rep']]];
    });

    addSection(
      '3D wait >3d (as-of window end)',
      'risk_3dWait',
      rows3dWait,
      ['RootApptID','Customer','Last 3D Request','Days Over (3D wait)','Rep']
    );
  }

  // 3D overdue (deadline before as-of)
  {
    const rows3dOverdue = metrics.filter(r => {
      if (!isActive(r)) return false;
      if (brand && String(r[xi['Brand']]).trim() !== brand) return false;
      if (rep   && String(r[xi['Assigned Rep']]).trim() !== rep) return false;
      const due = r[xi['3D Deadline']];
      const pending = truthy(r[xi['3D Pending?']]); // only count if still pending
      return pending && (due instanceof Date) && startOfDay_(due) < startOfDay_(asOf);
    }).map(r => {
      const due = r[xi['3D Deadline']];
      const over = daysBetween_(asOf, due); // calendar days past due
      return [r[xi['RootApptID']], r[xi['Customer Name']], due, over, r[xi['Assigned Rep']]];
    });

    addSection(
      '3D overdue (as-of window end)',
      'risk_3dOverdue',
      rows3dOverdue,
      ['RootApptID','Customer','3D Deadline','Days Overdue','Rep']
    );
  }

  // Production overdue (deadline before as-of)
  {
    const rowsPrOverdue = metrics.filter(r => {
      if (!isActive(r)) return false;
      if (brand && String(r[xi['Brand']]).trim() !== brand) return false;
      if (rep   && String(r[xi['Assigned Rep']]).trim() !== rep) return false;
      const due = r[xi['Prod Deadline']];
      return (due instanceof Date) && startOfDay_(due) < startOfDay_(asOf);
    }).map(r => {
      const due = r[xi['Prod Deadline']];
      const over = daysBetween_(asOf, due);
      return [r[xi['RootApptID']], r[xi['Customer Name']], due, over, r[xi['Assigned Rep']]];
    });

    addSection(
      'Production overdue (as-of window end)',
      'risk_prOverdue',
      rowsPrOverdue,
      ['RootApptID','Customer','Prod Deadline','Days Overdue','Rep']
    );
  }

  // Ops escalation (≥2 deadline moves and order age >28 biz days)
  {
    const rowsEscal = metrics.filter(r => {
      if (!isActive(r)) return false;
      if (brand && String(r[xi['Brand']]).trim() !== brand) return false;
      if (rep   && String(r[xi['Assigned Rep']]).trim() !== rep) return false;
      const moves3d = Number(r[xi['# 3D Deadline Moves']]) || 0;
      const movesPr = Number(r[xi['# Prod Deadline Moves']]) || 0;
      const orderDate = r[xi['Order Date']];
      const orderAgeBiz = (orderDate instanceof Date) ? businessDaysInclusive_(orderDate, asOf) : 0;
      return (moves3d >= 2 || movesPr >= 2) && orderAgeBiz > 28;
    }).map(r => {
      const d3Due   = r[xi['3D Deadline']];
      const prDue   = r[xi['Prod Deadline']];
      const d3Over  = (d3Due instanceof Date && startOfDay_(d3Due) < startOfDay_(asOf)) ? daysBetween_(asOf, d3Due) : 0;
      const prOver  = (prDue instanceof Date && startOfDay_(prDue) < startOfDay_(asOf)) ? daysBetween_(asOf, prDue) : 0;
      const daysOver = Math.max(d3Over, prOver); // use whichever is most overdue (0 if neither)
      return [
        r[xi['RootApptID']],
        r[xi['Customer Name']],
        r[xi['# 3D Deadline Moves']],
        r[xi['# Prod Deadline Moves']],
        daysOver || '',            // blank if not overdue
        r[xi['Assigned Rep']]
      ];
    });

    addSection(
      'Ops escalation (as-of window end)',
      'risk_escal',
      rowsEscal,
      ['RootApptID','Customer','# 3D Moves','# Prod Moves','Days Overdue','Rep']
    );
  }

  // --- Order Funnel (Window Cohort) — step lists ------------------------------
  {
    // Reuse master/mH filters and helpers already declared in this function:
    const inWin   = (d) => { const t=asDate_(d); return t && t>=start && t<=end; };
    const matchBR = (r) => (!brand || String(r[iBrand]||'').trim()===brand) &&
                          (!rep   || String(r[iRep]  ||'').trim()===rep);
    const vt      = (r) => String(r[iVType]||'').trim().toLowerCase();
    const isSched = (r) => !iStatus || String(r[iStatus]||'').trim().toLowerCase()==='scheduled';

    // Step 1: Consultations in window
    const consultRows = master.filter(r => isFunnelRow(r) && vt(r)==='appointment');
    const consultRoots = new Set(consultRows.map(r => r[iRoot]));
    const consultList = [...consultRoots].map(root => {
      const first = consultRows.find(r => r[iRoot]===root);
      return [root, (iCust>=0?first[iCust]:''), first[iVDate], first[iRep], first[iBrand]];
    });
    addSection('Order Funnel — Step 1: Consultations (window)', 'chart_flow_s1',
              consultList, ['RootApptID','Customer','Consult Date','Rep','Brand']);

    // Step 2: DVs in window among the consult cohort
    const dvRows      = master.filter(r => isFunnelRow(r) && vt(r)==='diamond viewing');
    const dvRootsInWin = new Set(dvRows.map(r => r[iRoot]));
    const step2Roots = [...consultRoots].filter(root => dvRootsInWin.has(root));
    const dvList = step2Roots.map(root => {
      const first = dvRows.find(r => r[iRoot]===root);
      return [root, (iCust>=0?first[iCust]:''), first[iVDate], first[iRep], first[iBrand]];
    });
    addSection('Order Funnel — Step 2: Diamond Viewings (window)', 'chart_flow_s2',
              dvList, ['RootApptID','Customer','DV Date','Rep','Brand']);

    // Step 3: First deposits in window among the DV cohort
    const xi = makeIdx_(xH);
    const depDateCol = xi['Deposit Date (First Pay)'];
    const rootCol    = xi['RootApptID'];
    const brandCol   = xi['Brand'];
    const repCol     = xi['Assigned Rep'];
    const depRowsInWin = metrics.filter(r =>
      (!brand || String(r[brandCol]||'').trim()===brand) &&
      (!rep   || String(r[repCol]  ||'').trim()===rep) &&
      r[depDateCol] instanceof Date &&
      r[depDateCol] >= start && r[depDateCol] <= end
    );
    const depRootsInWin = new Set(depRowsInWin.map(r => r[rootCol]));
    const step3Roots = new Set(step2Roots.filter(root => depRootsInWin.has(root)));
    const depList = depRowsInWin
      .filter(r => step3Roots.has(r[rootCol]))
      .map(r => [r[xi['RootApptID']], r[xi['Customer Name']], r[depDateCol], r[xi['Order Total']]]);
    addSection('Order Funnel — Step 3: First Deposits (>$'+MIN_FIRST_DEPOSIT_NET+') in window',
              'chart_flow_s3',
              depList, ['RootApptID','Customer','First Deposit Date','Order Total']);

    // Step 4: Orders Completed in window among the deposit cohort
    const orderDateCol = xi['Order Date'];
    const orderList = metrics.filter(r =>
      step3Roots.has(r[xi['RootApptID']]) &&
      r[orderDateCol] instanceof Date &&
      r[orderDateCol] >= start && r[orderDateCol] <= end
    ).map(r => [r[xi['RootApptID']], r[xi['Customer Name']], r[orderDateCol], r[xi['Order Total']]]);
    addSection('Order Funnel — Step 4: Orders Completed (window)', 'chart_flow_s4',
              orderList, ['RootApptID','Customer','Order Date','Order Total']);
  }
  
  // === Order Funnel (All-Time) — step lists ==================================
  {
    // Sets (all-time)
    const consultRootsAll = new Set(
      master.filter(r => matchBR(r) && String(r[iVType]||'').trim().toLowerCase()==='appointment')
            .map(r => r[iRoot]).filter(Boolean)
    );
    const dvRootsAll = new Set(
      master.filter(r => matchBR(r) && String(r[iVType]||'').trim().toLowerCase()==='diamond viewing')
            .map(r => r[iRoot]).filter(Boolean)
    );
    const step2RootsAll = new Set([...consultRootsAll].filter(root => dvRootsAll.has(root)));

    const depCol   = xi['Deposit Date (First Pay)'];
    const orderCol = xi['Order Date'];
    const rootCol  = xi['RootApptID'];
    const brandCol = xi['Brand'];
    const repCol   = xi['Assigned Rep'];

    const depRootsAll = new Set(
      metrics.filter(r =>
        (!brand || String(r[brandCol]||'').trim()===brand) &&
        (!rep   || String(r[repCol]  ||'').trim()===rep) &&
        (r[depCol] instanceof Date)
      ).map(r => r[rootCol]).filter(Boolean)
    );

    // First deposits among those who consulted (same rule as the chart)
    const step3RootsAll = new Set([...consultRootsAll].filter(root => depRootsAll.has(root)));

    const orderRootsAll = new Set(
      metrics.filter(r =>
        (!brand || String(r[brandCol]||'').trim()===brand) &&
        (!rep   || String(r[repCol]  ||'').trim()===rep) &&
        (r[orderCol] instanceof Date)
      ).map(r => r[rootCol]).filter(Boolean)
    );
    const step4RootsAll = new Set([...step3RootsAll].filter(root => orderRootsAll.has(root)));

    // Helper to pick the first date of a given type in MASTER
    const firstRowByType = (root, wantType) => {
      const rows = master.filter(r => r[iRoot]===root &&
        String(r[iVType]||'').trim().toLowerCase()===wantType);
      rows.sort((a,b)=> asDate_(a[iVDate]) - asDate_(b[iVDate]));
      return rows[0] || null;
    };

    // S1: Consultations (all-time)
    {
      const rows = [...consultRootsAll].map(root => {
        const fr = firstRowByType(root, 'appointment');
        return [root, (iCust>=0 && fr ? fr[iCust] : ''), (fr ? fr[iVDate] : ''), (fr ? fr[iRep] : ''), (fr ? fr[iBrand] : '')];
      });
      addSection('Order Funnel — All-Time: Consultations', 'chart_flow_all_s1',
                rows, ['RootApptID','Customer','First Consult Date','Rep','Brand']);
    }

    // S2: DVs (all-time) among the consult cohort
    {
      const rows = [...step2RootsAll].map(root => {
        const fr = firstRowByType(root, 'diamond viewing');
        return [root, (iCust>=0 && fr ? fr[iCust] : ''), (fr ? fr[iVDate] : ''), (fr ? fr[iRep] : ''), (fr ? fr[iBrand] : '')];
      });
      addSection('Order Funnel — All-Time: Diamond Viewings', 'chart_flow_all_s2',
                rows, ['RootApptID','Customer','First DV Date','Rep','Brand']);
    }

    // S3: First deposits (> threshold) among those who consulted
    {
      const rows = metrics.filter(r =>
        step3RootsAll.has(r[rootCol]) &&
        (!brand || String(r[brandCol]||'').trim()===brand) &&
        (!rep   || String(r[repCol]  ||'').trim()===rep) &&
        (r[depCol] instanceof Date)
      ).map(r => [r[rootCol], r[xi['Customer Name']], r[depCol], r[xi['Order Total']]]);
      addSection('Order Funnel — All-Time: First Deposits (>' + MIN_FIRST_DEPOSIT_NET + ')', 'chart_flow_all_s3',
                rows, ['RootApptID','Customer','First Deposit Date','Order Total']);
    }

    // S4: Orders Completed (all-time) among the deposit cohort
    {
      const rows = metrics.filter(r =>
        step3RootsAll.has(r[rootCol]) &&
        (!brand || String(r[brandCol]||'').trim()===brand) &&
        (!rep   || String(r[repCol]  ||'').trim()===rep) &&
        (r[orderCol] instanceof Date)
      ).map(r => [r[rootCol], r[xi['Customer Name']], r[orderCol], r[xi['Order Total']]]);
      addSection('Order Funnel — All-Time: Orders Completed', 'chart_flow_all_s4',
                rows, ['RootApptID','Customer','Order Date','Order Total']);
    }
  }

  // Fit columns + a single filter for convenience
  if (sh.getFilter()) sh.getFilter().remove();
  sh.autoResizeColumns(1, Math.min(8, sh.getLastColumn()));
  sh.getRange(1,1,Math.max(2,sh.getLastRow()), Math.max(5, sh.getLastColumn())).createFilter();

  // Named range for easier HYPERLINK gid reference
  ss.setActiveSheet(sh);
  sh.setName('Drill_KPI'); // ensure name stable
  ss.setActiveSheet(SpreadsheetApp.getActive().getSheetByName(SH_DASH));
  return anchors;
}


/** Make gap columns ≈ 1/4 of card column width for nicer spacing */
function applyCardColumnLayout_(sheet, startCol, cardsPerRow, cardWidth, hGap) {
  const baseWidth = 120;   // px per "card" column
  const gapWidth  = 30;    // px ≈ 1/4
  const totalCols = cardsPerRow * cardWidth + (cardsPerRow - 1) * hGap;
  for (let i = 0; i < totalCols; i++) {
    const col = startCol + i;
    // decide if this column is a gap:
    const inBlock = i % (cardWidth + hGap);
    const isGap = (inBlock >= cardWidth);
    sheet.setColumnWidth(col, isGap ? gapWidth : baseWidth);
  }
}

/** Half-width cards with 1/4-width gaps — but make the 2nd (icon) col narrower */
function setCardGridColumnWidths_() {
  const dash = SpreadsheetApp.getActive().getSheetByName(SH_DASH);

  // Layout used by writeDashboard_:
  // - startCol = A (1)
  // - 10 cards per row
  // - each card = 2 columns (col1=label, col2=icon)
  // - plus 1 gap col after each card
  const startCol = 1, cardsPerRow = 10, cardWidthCols = 2, hGapCols = 1;

  // Previous total per block = 120 + 120 + 30 = 270px. Keep overall width.
  // New: label col wider, icon col about half of label.
  const labelColPx = 130;
  const iconColPx  = 80;     // ~half of label
  const gapColPx   = 30;

  const block = cardWidthCols + hGapCols;           // 3 columns per card-block
  const totalCols = cardsPerRow * block - hGapCols; // no gap after last card

  for (let i = 0; i < totalCols; i++) {
    const col = startCol + i;
    const inBlock = i % block;

    if (inBlock === 0) {
      // first col in the card (label)
      dash.setColumnWidth(col, labelColPx);
    } else if (inBlock === 1) {
      // second col in the card (🔎 icon / drill)
      dash.setColumnWidth(col, iconColPx);
    } else {
      // gap column
      dash.setColumnWidth(col, gapColPx);
    }
  }
}

/** Build dropdowns for Brand (H1) and Rep (H2) from metrics data */
/** Build dropdowns for Brand (H1) and Rep (H2) */
function ensureBrandRepDropdowns_() {
  const ss = SpreadsheetApp.getActive();
  const dash = ss.getSheetByName(SH_DASH);

  // --- Brand dropdown: static list
  const brands = ['', 'HPUSA', 'VVS'];  // '' = allow blank = all brands
  const dvBrand = SpreadsheetApp.newDataValidation()
    .requireValueInList(brands, true)
    .build();
  dash.getRange(CELL_BRAND).setDataValidation(dvBrand);

  // --- Rep dropdown: from "Dropdown" tab, column "Assigned Rep"
  const DROPDOWN_SHEET = 'Dropdown';
  const repValues = (() => {
    const sh = ss.getSheetByName(DROPDOWN_SHEET);
    if (!sh) return [];

    const vals = sh.getDataRange().getValues();
    if (!vals.length) return [];

    // locate header "Assigned Rep" (case/substring tolerant)
    const H = vals[0].map(h => String(h || '').trim().toLowerCase());
    let idx = H.indexOf('assigned rep');
    if (idx < 0) {
      // fallback: find a header that contains both words
      idx = H.findIndex(h => h.includes('assigned') && h.includes('rep'));
    }
    if (idx < 0) return [];

    // unique, trimmed, non-blank reps from that column
    const uniq = new Set();
    for (let r = 1; r < vals.length; r++) {
      const v = String(vals[r][idx] || '').trim();
      if (v) uniq.add(v);
    }
    return Array.from(uniq).sort((a,b)=>a.localeCompare(b));
  })();

  const dvRep = SpreadsheetApp.newDataValidation()
    .requireValueInList(['', ...repValues], true) // '' = all reps
    .build();
  dash.getRange(CELL_REP).setDataValidation(dvRep);

  // Notes (optional)
  dash.getRange(CELL_BRAND).setNote('Choose HPUSA or VVS. Leave blank for all.');
  dash.getRange(CELL_REP).setNote('Pick from Dropdown!Assigned Rep. Leave blank for all.');

  // Clean up any legacy validations we no longer use
  // (safe no-ops if they don’t exist)
  dash.getRange('B2').clearDataValidations(); // date inputs stay dates
  dash.getRange('E1').clearDataValidations(); // old preset cell, if used in the past
  dash.getRange('E2').clearDataValidations();
}


/** Period preset in B1 (merged B1:D1) */
function ensurePeriodPreset_() {
  const dash = SpreadsheetApp.getActive().getSheetByName(SH_DASH);
  const list = ['This Week (Mon–Sun)','Last Week (Mon–Sun)','This Month','Last Month','Quarter to Date','Year to Date','Custom'];
  const dv = SpreadsheetApp.newDataValidation().requireValueInList(list, true).build();
  dash.getRange(CELL_PRESET).setDataValidation(dv);  // B1
  if (!String(dash.getRange(CELL_PRESET).getValue()||'')) {
    dash.getRange(CELL_PRESET).setValue('This Week (Mon–Sun)'); // show something by default
  }
}

function applyPreset_(preset) {
  const dash = SpreadsheetApp.getActive().getSheetByName(SH_DASH);
  const today = startOfDay_(new Date());
  const p = String(preset||'').toLowerCase();
  const mon = mondayOfWeek_(today);
  const sun = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate()+6);

  let s=null, e=null;
  if (p.includes('this week'))        { s=mon; e=sun; }
  else if (p.includes('last week'))   { const m2=new Date(mon.getFullYear(),mon.getMonth(),mon.getDate()-7); s=m2; e=new Date(m2.getFullYear(),m2.getMonth(),m2.getDate()+6); }
  else if (p.includes('this month'))  { s=new Date(today.getFullYear(),today.getMonth(),1); e=new Date(today.getFullYear(),today.getMonth()+1,0); }
  else if (p.includes('last month'))  { const lmEnd=new Date(today.getFullYear(),today.getMonth(),0); s=new Date(lmEnd.getFullYear(),lmEnd.getMonth(),1); e=lmEnd; }
  else if (p.includes('quarter to date')) { const q=Math.floor(today.getMonth()/3); s=new Date(today.getFullYear(),q*3,1); e=today; }
  else if (p.includes('year to date'))    { s=new Date(today.getFullYear(),0,1); e=today; }
  else return;

  dash.getRange(CELL_DATE_S).setValue(s).setNumberFormat('yyyy-mm-dd'); // B2
  dash.getRange(CELL_DATE_E).setValue(e).setNumberFormat('yyyy-mm-dd'); // D2
}

/** Hook the preset to onEdit */
function onEdit(e){
  try{
    const r=e.range, sh=r.getSheet(); if(sh.getName()!==SH_DASH) return;
    const a1=r.getA1Notation();
    if(a1===CELL_PRESET){ const v=String(r.getValue()||'').trim(); if(v && v.toLowerCase()!=='custom') applyPreset_(v); }
    if(a1===CELL_DATE_S || a1===CELL_DATE_E){ if (typeof updateQuickPickCaption_==='function') updateQuickPickCaption_(); }
  }catch(_){}
}


function byBRMetrics_(r, xi, brand, rep) {
  return (!brand || String(r[xi['Brand']]).trim() === brand) &&
         (!rep   || String(r[xi['Assigned Rep']]).trim() === rep);
}



/** Write a small table for charts; returns the full range used by the chart. */
function writeChartTable_(sh, startCol, startRow, headers, rows) {
  const colL = columnToLetter(startCol);
  const endCol = columnToLetter(startCol + headers.length - 1);
  const n = Math.max(1, rows.length);
  const target = sh.getRange(startRow, startCol, n + 1, headers.length);
  target.clearContent();
  sh.getRange(startRow, startCol, 1, headers.length).setValues([headers]).setFontWeight('bold').setFontSize(9);
  sh.getRange(startRow + 1, startCol, n, headers.length).setValues(rows);
  // Return the full data range including header for chart.addRange()
  return sh.getRange(startRow, startCol, n + 1, headers.length);
}


/** Sum payments (receipts) in ledger within window for a set of allowed roots */
function sumPaymentsInWindowForRoots_(winStart, winEnd, allowedRootsSet) {
  const LEDGER_FILE_ID = PropertiesService.getScriptProperties().getProperty('PAYMENTS_400_FILE_ID');
  if (!LEDGER_FILE_ID) return 0;
  const lss = SpreadsheetApp.openById(LEDGER_FILE_ID);
  const sh = lss.getSheetByName('Payments');
  if (!sh) return 0;

  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return 0;
  const H = {}; vals[0].forEach((h,i)=> H[String(h||'').trim()] = i);

  const cRoot = H['RootApptID'] ?? H['Root Appt ID'] ?? H['APPT_ID'];
  const cWhen = H['PaymentDateTime'] ?? H['Payment DateTime'] ?? H['Payment Date'] ?? H['Paid At'];
  const cDoc  = H['DocType'] ?? H['Document Type'] ?? H['Type'];
  const cNet  = H['AmountNet'] ?? H['Net'] ?? H['Net Amount'] ?? H['Amount'];
  const cStat = H['DocStatus'] ?? H['Status'];

  let sum = 0;
  for (let i = 1; i < vals.length; i++) {
    const row  = vals[i];
    const root = row[cRoot];
    const when = row[cWhen];
    const doc  = String(row[cDoc]||'');
    const net  = Number(row[cNet])||0;
    const stat = (cStat != null ? String(row[cStat]||'') : '');
    if (!root || !allowedRootsSet.has(root)) continue;
    if (!(when instanceof Date)) continue;
    if (when < winStart || when > winEnd) continue;
    if (!/receipt/i.test(doc)) continue;
    if (!(net > 0)) continue;
    if (/void|reversed|cancel/i.test(stat)) continue;
    sum += net;
  }
  return sum;
}

/** === CHARTS (Row 4) ===================================================== */
/**
 * Builds chart data in hidden columns (BA:BM…) and inserts 4 embedded charts:
 *  - Pipeline by Stage (bars=$ weighted + line=count; as-of window end)
 *  - Weekly First Deposits (line=count + bars=$ sums; last N weeks)
 *  - Order Flow (Window) — counts
 *  - Totals in Window — first deposits (#) + payments ($)
 */
const CHART_ANCHOR_ROW = 41;           // top row for charts
const CHART_ANCHOR_COL = 1;            // A
const CHART_DATA_START_COL = 53;       // BA
const CHART_TITLES = {
  pipeline: 'Pipeline by Stage (Count)',
  weekly:   'Weekly Activity & First Deposits (# + $)', 
  flow: 'Order Funnel (Current Period)',
  flowAll:  'Order Funnel (All-Time)', 
  totals:   'Totals in Window'
};

function writeChartsRow4to15({ start, end, brand, rep, master, mH, metrics, xH, anchors, drillGid }) {
  const ss = SpreadsheetApp.getActive();
  const dash = ss.getSheetByName(SH_DASH);

  // Resolve the drill sheet gid once for hyperlinks under charts
  const drillSheet = ss.getSheetByName('Drill_KPI');

  const xi = makeIdx_(xH);
  const stageWeights = getStageWeightMap_();

  // 1) Build data tables (hidden)
  const rPipeline = upsertPipelineByStageData_(dash, metrics, xi, stageWeights, end, brand, rep);
  const rWeekly   = upsertWeeklyActivityAndDepositsData_(dash, master, mH, metrics, xi, end, brand, rep, /*weeks*/12);
  const rFlow     = upsertOrderFunnelData_(dash, master, mH, metrics, xi, start, end, brand, rep);
  const rFlowAll  = upsertOrderFunnelAllTimeData_(dash, master, mH, metrics, xi, brand, rep); // moved up

  // 2) Remove prior charts that use our hidden BA:CB block (idempotent)
  removeChartsByTitles_(dash, [CHART_TITLES.pipeline, CHART_TITLES.weekly, CHART_TITLES.flow, CHART_TITLES.flow, CHART_TITLES.flowAll, CHART_TITLES.totals]);

  // 3) Dimensions
  const halfWidthPx = 650;
  const fullWidthPx = 1320;
  const topHeightPx = 260;
  const bottomHeightPx = 300;

  // 4) Insert charts
  // Clear any old charts in the target region
  const DASH = getSheetOrThrow_('00_Dashboard');

  // === Section Title ===
  DASH.getRange(40, 1).setValue('📊 Sales Pipeline Analysis — Current Stage Distribution')
      .setFontWeight('bold')
      .setFontSize(12)
      .setFontColor('#333333');


  // A) Pipeline by Stage — COLUMN (counts only, with data labels)
  {
    // Drop the header row so “Stage” is not treated as a category value
    // and only keep the first 2 columns [Stage, Count].
    const dataNoHeader = rPipeline.offset(1, 0, rPipeline.getNumRows() - 1, 2);

    const b = dash.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataNoHeader)      // data only
      .setNumHeaders(0)            // <-- no header row in the range above
      .setOption('title', CHART_TITLES.pipeline)
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { title: 'Stage' })
      .setOption('vAxis', { title: 'Count (#)', format: '0' })
      // <<< SHOW LABELS ON BARS >>>
      .setOption('series', { 0: { dataLabel: 'value' } })   // Sheets “data labels” for first series
      .setOption('width', halfWidthPx)
      .setOption('height', topHeightPx)
      .setPosition(CHART_ANCHOR_ROW, CHART_ANCHOR_COL, 0, 0);
    dash.insertChart(b.build());
  }

  // Under Pipeline chart
  if (drillGid && anchors && anchors['pay_weighted']) {
    dash.getRange(CHART_ANCHOR_ROW + 13, CHART_ANCHOR_COL, 1, 2)
      .merge()
      .setFormula(`=HYPERLINK("#gid=${drillGid}&range=${anchors['pay_weighted']}","🔎 Drill")`)
      .setHorizontalAlignment('left').setFontSize(10);
  }

  // B) Weekly Activity & First Deposits — COMBO (lines = counts, bars = $)
  {
    const b = dash.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(rWeekly)                  // includes header row
      .setNumHeaders(1)                   // <-- ensures "Use row 4 as headers"
      .setOption('title', CHART_TITLES.weekly)
      .setOption('legend', { position: 'right' })
      .setOption('hAxis', { title: 'Week (Mon start)' })
      .setOption('vAxes', {
        0: { title: 'Counts (#)',               format: '0'      },
        1: { title: 'First Deposit Sum ($)',    format: '$#,##0' }
      })
      .setOption('series', {
        0: { type: 'line', targetAxisIndex: 0, pointSize: 4, lineWidth: 2 }, // Consultations
        1: { type: 'line', targetAxisIndex: 0, pointSize: 4, lineWidth: 2 }, // Diamond Viewings
        2: { type: 'line', targetAxisIndex: 0, pointSize: 4, lineWidth: 2 }, // First Deposits (#)
        3: { type: 'bars', targetAxisIndex: 1 }                               // First Deposit Sum ($)
      })
      .setOption('width', halfWidthPx)
      .setOption('height', topHeightPx)
      .setPosition(CHART_ANCHOR_ROW, CHART_ANCHOR_COL + 9, 0, 0);
    dash.insertChart(b.build());

    // Optional: a single drill link under the chart to show the first‑time deposit rows
    if (drillGid && anchors && anchors['pay_firstDeposits']) {
      dash.getRange(CHART_ANCHOR_ROW + 13, CHART_ANCHOR_COL + 9, 1, 3)
        .merge()
        .setFormula(`=HYPERLINK("#gid=${drillGid}&range=${anchors['pay_firstDeposits']}",
          "🔎 Drill: show first‑time deposit rows")`)
        .setHorizontalAlignment('left')
        .setFontSize(10);
    }
  }

  // C) Order Flow (Window) — BAR (counts)
  {
    const b = dash.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(rFlow) // [Step, Count]
      .setNumHeaders(1)   // <-- tells Sheets to treat row 4 as headers
      .setOption('series', { 0: { dataLabel: 'value' } })
      .setOption('title', CHART_TITLES.flow)
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { title: 'Count (#)', format: '0' })
      .setOption('vAxis', { title: '' })        // no “Stage” title
      .setOption('width', halfWidthPx)
      .setOption('height', bottomHeightPx)
      .setPosition(CHART_ANCHOR_ROW + 15, CHART_ANCHOR_COL, 0, 0);
    dash.insertChart(b.build());
  }

  // Under Order Funnel (Window) chart
  if (drillGid && anchors && anchors['chart_flow_s1']) {
    dash.getRange(CHART_ANCHOR_ROW + 30, CHART_ANCHOR_COL, 1, 2)
      .merge()
      .setFormula(`=HYPERLINK("#gid=${drillGid}&range=${anchors['chart_flow_s1']}","🔎 Drill")`)
      .setHorizontalAlignment('left').setFontSize(10);
  }

  // D) Order Funnel (All-Time) — BAR (counts)  [moved into bottom-right slot]
  {
    const b = dash.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(rFlowAll)                 // includes header row
      .setNumHeaders(1)
      .setOption('series', { 0: { dataLabel: 'value' } })
      .setOption('title', CHART_TITLES.flowAll)
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { title: 'Count (#)', format: '0' })
      .setOption('vAxis', { title: '' })
      .setOption('width', halfWidthPx)
      .setOption('height', bottomHeightPx)
      .setPosition(CHART_ANCHOR_ROW + 15, CHART_ANCHOR_COL + 9, 0, 0);
    dash.insertChart(b.build());

    // Drill link under the all-time funnel → all-time step lists
    if (drillGid && anchors && anchors['chart_flow_all_s1']) {
      dash.getRange(CHART_ANCHOR_ROW + 30, CHART_ANCHOR_COL + 9, 1, 3)
        .merge()
        .setFormula(`=HYPERLINK("#gid=${drillGid}&range=${anchors['chart_flow_all_s1']}",
          "🔎 Drill: all-time step lists")`)
        .setHorizontalAlignment('left')
        .setFontSize(10);
    }
  }

}

/** Normalize stage and check for Lost Lead (case/variant tolerant) */
function isLostLeadStage_(stage) {
  const s = String(stage || '').trim().toLowerCase();
  return s === 'lost lead' || s === 'lost' || s.indexOf('lost lead') >= 0;
}

/** Also treat explicit "Won" as non-risk for No‑touch */
function isWonStage_(stage) {
  return /\bwon\b/i.test(String(stage || ''));
}
/** Convenience: Won OR Lost Lead */
function isWonOrLostStage_(stage) {
  return isWonStage_(stage) || isLostLeadStage_(stage);
}

/** Build a Set of RootApptIDs that pass brand/rep and Sales Stage != Lost Lead. */
function allowedRootsByStage_(metrics, xi, brand, rep) {
  const set = new Set();
  for (const r of metrics) {
    if (brand && String(r[xi['Brand']]||'').trim() !== brand) continue;
    if (rep   && String(r[xi['Assigned Rep']]||'').trim() !== rep) continue;
    const st = String(r[xi['Sales Stage']]||'').trim();
    if (isLostLeadStage_(st)) continue;
    const root = r[xi['RootApptID']]; if (root) set.add(root);
  }
  return set;
}

/** Stage-weighted value (no Lost Lead filter here; the caller filters rows) */
function weightedValueForRowNoLostLead_(row, xi, stageWeights) {
  const stageNorm = normalizeStage_(String(row[xi['Sales Stage']]||'').trim());
  const w = stageWeights[stageNorm] ?? 0;
  const orderTotal = Number(row[xi['Order Total']]) || 0;
  const bMin = Number(row[xi['Budget Min']]);   // may be NaN
  const bMax = Number(row[xi['Budget Max']]);   // may be NaN
  const budgetMid = (isFinite(bMin) && isFinite(bMax)) ? (bMin + bMax)/2 : 0;
  const valueForWeight = orderTotal > 0 ? orderTotal : budgetMid;
  return w * valueForWeight;
}

// Pipeline by Stage (as-of C1): COUNT ONLY; restrict to 5 stages; Stage != Lost Lead
function upsertPipelineByStageData_(dash, metrics, xi, stageWeights, asOf, brand, rep) {
  const cnts = new Map();
  const desiredOrder = ['APPOINTMENT','LEAD','HOT LEAD','DEPOSIT','ORDER COMPLETED'];
  const allowedSet = new Set(desiredOrder);

  for (const r of metrics) {
    if (brand && String(r[xi['Brand']]||'').trim() !== brand) continue;
    if (rep   && String(r[xi['Assigned Rep']]||'').trim() !== rep) continue;

    const stageRaw = String(r[xi['Sales Stage']]||'').trim();
    if (isLostLeadStage_(stageRaw)) continue;

    const stage = normalizeStage_(stageRaw);
    if (!allowedSet.has(stage)) continue;

    cnts.set(stage, (cnts.get(stage)||0) + 1);
  }

  // Table: [Stage, Count]
  const headers = ['Stage','Count (#)'];
  const rows = [];
  for (const s of desiredOrder) if (cnts.has(s)) rows.push([s, cnts.get(s)]);
  if (!rows.length) rows.push(['(no data)', 0]);

  const col = CHART_DATA_START_COL;      // BA
  const startRow = 4;
  const n = rows.length;
  const rng = dash.getRange(startRow, col, n + 1, headers.length);
  rng.clearContent();
  dash.getRange(startRow, col, 1, headers.length).setValues([headers]).setFontWeight('bold').setFontSize(9);
  dash.getRange(startRow + 1, col, n, headers.length).setValues(rows);
  dash.getRange(startRow + 1, col + 1, n, 1).setNumberFormat('0'); // integer

  return dash.getRange(startRow, col, n + 1, headers.length);
}

/** Weekly Activity + First Deposits (last N weeks)
 * Builds a table: [Week, Consultations (#), Diamond Viewings (#), First Deposits (#), First Deposit Sum ($)]
 * - Consultations / DVs use MASTER rows with Status=Scheduled and Visit Date in that week.
 * - First Deposits (#) uses metrics' "Deposit Date (First Pay)" (already thresholded > MIN_FIRST_DEPOSIT_NET).
 * - First Deposit Sum ($) uses the ledger’s first valid receipt per root (same week as first-deposit date).
 */
function upsertWeeklyActivityAndDepositsData_(dash, master, mH, metrics, xi, asOf, brand, rep, weeks) {
  const tz = Session.getScriptTimeZone() || 'GMT';
  const weekKey = (d) => Utilities.formatDate(mondayOfWeek_(d), tz, 'yyyy-MM-dd');

  // Prepare bins for the last N weeks (keys are Mondays of those weeks)
  const bins = new Map();
  const newestMon = mondayOfWeek_(new Date(asOf));
  for (let i = weeks - 1; i >= 0; i--) {
    const d = new Date(newestMon.getFullYear(), newestMon.getMonth(), newestMon.getDate() - 7 * i);
    bins.set(Utilities.formatDate(d, tz, 'yyyy-MM-dd'), { consults: 0, dvs: 0, depCnt: 0, depSum: 0 });
  }

  // --- MASTER: count scheduled Consultations / Diamond Viewings by Visit Date week
  const iBrand  = findCol_(mH, ['Brand']);
  const iRep    = findCol_(mH, ['Assigned Rep','AssignedRep','Rep','Sales Rep']);
  const iVType  = findCol_(mH, ['Visit Type','VisitType','Type']);
  const iVDate  = findCol_(mH, ['Visit Date','Visit_Date','Appt Date','Appointment Date']);
  const iStatus = findCol_(mH, ['Status','Appt Status','Appointment Status'], false);

  const matchesBR = (r) =>
    (!brand || String(r[iBrand]||'').trim()===brand) &&
    (!rep   || String(r[iRep]  ||'').trim()===rep);

  const isScheduled = (r) => iStatus < 0 ? true : (String(r[iStatus]||'').trim().toLowerCase() === 'scheduled');
  const vt          = (r) => String(r[iVType]||'').trim().toLowerCase();

  for (const r of master) {
    if (!matchesBR(r) || !isScheduled(r)) continue;
    const vd = asDate_(r[iVDate]); if (!(vd instanceof Date)) continue;
    const k = weekKey(vd); if (!bins.has(k)) continue;

    const type = vt(r);
    if (type === 'appointment')            bins.get(k).consults++;
    else if (type === 'diamond viewing')   bins.get(k).dvs++;
  }

  // --- METRICS + LEDGER: first‑time deposits by week (count + $)
  const depCol   = xi['Deposit Date (First Pay)'];
  const brandCol = xi['Brand'];
  const repCol   = xi['Assigned Rep'];
  const rootCol  = xi['RootApptID'];

  const firstPayMap = fetchFirstPaymentMapFromPayments_(); // Map<root,{d,amt,so}>

  for (const r of metrics) {
    if (brand && String(r[brandCol]||'').trim() !== brand) continue;
    if (rep   && String(r[repCol]  ||'').trim() !== rep)   continue;

    const dep = r[depCol];
    if (!(dep instanceof Date)) continue;

    const k = weekKey(dep);
    if (!bins.has(k)) continue;

    bins.get(k).depCnt++;

    // Sum $ from ledger map, only if same (Mon‑start) week as the deposit date.
    const root = r[rootCol];
    const fp = firstPayMap.get(root);
    if (fp && fp.d instanceof Date && mondayOfWeek_(fp.d).getTime() === mondayOfWeek_(dep).getTime()) {
      bins.get(k).depSum += Number(fp.amt) || 0;
    }
  }

  // Build the small table
  const headers = ['Week','Consultations (#)','Diamond Viewings (#)','First Deposits (#)','First Deposit Sum ($)'];
  const rows = Array.from(bins.entries())
    .sort((a,b)=> a[0].localeCompare(b[0]))
    .map(([k,v]) => [k, v.consults, v.dvs, v.depCnt, v.depSum]);

  // Write to the hidden block (BA..)
  const col = CHART_DATA_START_COL + 4;  // BE
  const startRow = 4;
  const n = rows.length;
  const rng = dash.getRange(startRow, col, n + 1, headers.length);
  rng.clearContent();
  dash.getRange(startRow, col, 1, headers.length).setValues([headers]).setFontWeight('bold').setFontSize(9);
  dash.getRange(startRow + 1, col, n, headers.length).setValues(rows);

  // Formats
  dash.getRange(startRow + 1, col + 1, n, 3).setNumberFormat('0');       // counts
  dash.getRange(startRow + 1, col + 4, n, 1).setNumberFormat('$#,##0');  // sum $

  return dash.getRange(startRow, col, n + 1, headers.length); // include header row
}

/** Order Funnel (Window Cohort): [Step, Count]
 * Cohort logic: start with roots that had a Consultation (Visit Type='Appointment') in the window.
 * Subsequent steps are strict subsets where the relevant event ALSO occurs in the same window.
 * Deposits use the "first real deposit" (Net > MIN_FIRST_DEPOSIT_NET) via metrics' First Deposit Date.
 */
function upsertOrderFunnelData_(dash, master, mH, metrics, xi, winStart, winEnd, brand, rep) {
  // --- master indexes
  const iRoot   = findCol_(mH, ['RootApptID','Root Appt ID','ROOT','Root_ID']);
  const iBrand  = findCol_(mH, ['Brand']);
  const iRep    = findCol_(mH, ['Assigned Rep','AssignedRep','Rep','Sales Rep']);
  const iVType  = findCol_(mH, ['Visit Type','VisitType','Type']);
  const iVDate  = findCol_(mH, ['Visit Date','Visit_Date','Appt Date','Appointment Date']);
  const iStatus = findCol_(mH, ['Status','Appt Status','Appointment Status'], false);
  const iResTo      = findCol_(mH, ['RescheduledToUID','Rescheduled To UID','ReschedToUID','Rescheduled To'], false);
  const iCanceledAt = findCol_(mH, ['CanceledAt','CancelledAt','Canceled At','Cancelled At'], false);

  const matchesBR = (r) =>
    (!brand || String(r[iBrand]||'').trim()===brand) &&
    (!rep   || String(r[iRep]  ||'').trim()===rep);

  const inRange   = (d) => { const t=asDate_(d); return t && t>=winStart && t<=winEnd; };
  const isSched   = (r) => !iStatus || String(r[iStatus]||'').trim().toLowerCase()==='scheduled';
  const vt        = (r) => String(r[iVType]||'').trim().toLowerCase();

  // Keep rows that count toward “actually happened/active” in the window
  const isFunnelRow = (r) => {
    if (!matchesBR(r) || !inRange(r[iVDate])) return false;

    const s = iStatus < 0 ? '' : String(r[iStatus]||'').trim().toLowerCase();
    const cancelled  = /cancel/.test(s);                       // cancel/cancelled/canceled
    const noShow     = /no[-\s]?show/.test(s);                 // no show
    const reschedOff = (iResTo>=0 && String(r[iResTo]||'').trim()!=='') &&
                      (iCanceledAt>=0 && (r[iCanceledAt] instanceof Date));

    // Include Scheduled/Completed/Done/Checked-in/etc.; exclude cancels, no-shows, and vacated (rescheduled off) slots
    return !cancelled && !noShow && !reschedOff;
  };

  // --- Step 1: Consultations cohort (unique roots with a Consultation row in window)
  const consultRoots = new Set(
    master
      .filter(r => isFunnelRow(r) && vt(r)==='appointment')
      .map(r => r[iRoot])
      .filter(Boolean)
  );

  // --- Step 2: DVs among those consultations (DV row ALSO in window)
  const dvRootsInWin = new Set(
    master
      .filter(r => isFunnelRow(r) && vt(r)==='diamond viewing')
      .map(r => r[iRoot])
      .filter(Boolean)
  );
  const step2Roots = new Set([...consultRoots].filter(root => dvRootsInWin.has(root)));

  // --- Step 3: First deposits among DV cohort (metrics' first-deposit date in window; already thresholded > MIN_FIRST_DEPOSIT_NET)
  const depDateCol = xi['Deposit Date (First Pay)'];
  const rootCol    = xi['RootApptID'];
  const brandCol   = xi['Brand'];
  const repCol     = xi['Assigned Rep'];
  const depRootsInWin = new Set(
    metrics
      .filter(r =>
        (!brand || String(r[brandCol]||'').trim()===brand) &&
        (!rep   || String(r[repCol]  ||'').trim()===rep) &&
        r[depDateCol] instanceof Date &&
        r[depDateCol] >= winStart && r[depDateCol] <= winEnd
      )
      .map(r => r[rootCol])
      .filter(Boolean)
  );
  // ✅ new rule: First Deposits among those who consulted (no DV requirement)
  const step3Roots = new Set([...consultRoots].filter(root => depRootsInWin.has(root)));

  // --- Step 4: Orders completed among deposit cohort (Order Date in window)
  const orderDateCol = xi['Order Date'];
  const orderRootsInWin = new Set(
    metrics
      .filter(r =>
        (!brand || String(r[brandCol]||'').trim()===brand) &&
        (!rep   || String(r[repCol]  ||'').trim()===rep) &&
        r[orderDateCol] instanceof Date &&
        r[orderDateCol] >= winStart && r[orderDateCol] <= winEnd
      )
      .map(r => r[rootCol])
      .filter(Boolean)
  );
  const step4Roots = new Set([...step3Roots].filter(root => orderRootsInWin.has(root)));

  // --- Build the tiny table for the chart
  const headers = ['Step', 'Unique Roots (#)'];
  const rows = [
    ['Consultations',            consultRoots.size],
    ['Diamond Viewings',         step2Roots.size],
    ['First Deposits (> $' + MIN_FIRST_DEPOSIT_NET + ')', step3Roots.size],
    ['Orders Completed',         step4Roots.size]
  ];

  const col = CHART_DATA_START_COL + 9; // BJ (same slot as before)
  const startRow = 4;
  const n = rows.length;
  const rng = dash.getRange(startRow, col, n + 1, headers.length);
  rng.clearContent();
  dash.getRange(startRow, col, 1, headers.length)
     .setValues([headers]).setFontWeight('bold').setFontSize(9);
  dash.getRange(startRow + 1, col, n, headers.length).setValues(rows);
  dash.getRange(startRow + 1, col + 1, n, 1).setNumberFormat('0'); // integer

  return dash.getRange(startRow, col, n + 1, headers.length);
}

/** Order Funnel (Historical, all-time): [Step, Count]
 * Steps are unique Roots across all time. No window filters.
 * Deposit uses first valid receipt > MIN_FIRST_DEPOSIT_NET (same map as metrics).
 */
function upsertOrderFunnelAllTimeData_(dash, master, mH, metrics, xi, brand, rep) {
  const iRoot = findCol_(mH, ['RootApptID','Root Appt ID','ROOT','Root_ID']);
  const iBrand= findCol_(mH, ['Brand']);
  const iRep  = findCol_(mH, ['Assigned Rep','AssignedRep','Rep','Sales Rep']);
  const iVType= findCol_(mH, ['Visit Type','VisitType','Type']);

  const matchBR = (r) => (!brand || String(r[iBrand]||'').trim()===brand) &&
                         (!rep   || String(r[iRep]  ||'').trim()===rep);
  const vt      = (r) => String(r[iVType]||'').trim().toLowerCase();

  const consultRoots = new Set(master.filter(r => matchBR(r) && vt(r)==='appointment').map(r => r[iRoot]).filter(Boolean));
  const dvRoots      = new Set(master.filter(r => matchBR(r) && vt(r)==='diamond viewing').map(r => r[iRoot]).filter(Boolean));
  const step2Roots   = new Set([...consultRoots].filter(root => dvRoots.has(root)));

  const depDateCol = xi['Deposit Date (First Pay)'];
  const brandCol   = xi['Brand'];
  const repCol     = xi['Assigned Rep'];
  const depRoots   = new Set(metrics
    .filter(r => (!brand || String(r[brandCol]||'').trim()===brand) &&
                 (!rep   || String(r[repCol]  ||'').trim()===rep) &&
                 r[depDateCol] instanceof Date)
    .map(r => r[xi['RootApptID']])
    .filter(Boolean));

  // ✅ new rule: First Deposits among those who consulted (no DV requirement)
  const step3Roots = new Set([...consultRoots].filter(root => depRoots.has(root)));

  const orderCol   = xi['Order Date'];
  const orderRoots = new Set(metrics
    .filter(r => (!brand || String(r[brandCol]||'').trim()===brand) &&
                 (!rep   || String(r[repCol]  ||'').trim()===rep) &&
                 r[orderCol] instanceof Date)
    .map(r => r[xi['RootApptID']])
    .filter(Boolean));
  const step4Roots = new Set([...step3Roots].filter(root => orderRoots.has(root)));

  const headers = ['Step','Unique Roots (#)'];
  const rows = [
    ['Consultations (all-time)', consultRoots.size],
    ['Diamond Viewings (all-time)', step2Roots.size],
    ['First Deposits from Consultations (> $' + MIN_FIRST_DEPOSIT_NET + ')', step3Roots.size],
    ['Orders Completed (all-time)', step4Roots.size]
  ];

  const col = CHART_DATA_START_COL + 17; // put it right of other tables
  const startRow = 4;
  const n = rows.length;
  const rng = dash.getRange(startRow, col, n + 1, headers.length);
  rng.clearContent();
  dash.getRange(startRow, col, 1, headers.length).setValues([headers]).setFontWeight('bold').setFontSize(9);
  dash.getRange(startRow + 1, col, n, headers.length).setValues(rows);
  dash.getRange(startRow + 1, col + 1, n, 1).setNumberFormat('0');
  return dash.getRange(startRow, col, n + 1, headers.length);
}


/** Totals in Window (brand/rep filters only): [Metric, Value] */
function upsertTotalsWindowData_(dash, master, mH, metrics, xi, winStart, winEnd, brand, rep) {
  // Brand/Rep allowed roots (no Lost Lead filter here on purpose)
  const rootsBR = new Set(master
    .filter(r =>
      (!brand || String(r[findCol_(mH,['Brand'])]||'').trim()===brand) &&
      (!rep   || String(r[findCol_(mH,['Assigned Rep','AssignedRep','Rep','Sales Rep'])]||'').trim()===rep))
    .map(r => r[findCol_(mH,['RootApptID','Root Appt ID','ROOT','Root_ID'])])
  );

  // First-time deposits in window (count)
  let firstDeposits = 0;
  for (const r of metrics) {
    const root = r[xi['RootApptID']];
    const fd   = r[xi['Deposit Date (First Pay)']];
    if (!rootsBR.has(root)) continue;
    if (fd instanceof Date && fd >= winStart && fd <= winEnd) firstDeposits++;
  }

  // Payments (all valid receipts) in window (sum $)
  const paymentsSum = sumPaymentsInWindowForRoots_(winStart, winEnd, rootsBR);

  const headers = ['Metric','Value'];
  const rows = [
    ['First Deposits (#)', firstDeposits],
    ['Payments (Sum $)',   paymentsSum]
  ];

  const col = CHART_DATA_START_COL + 13; // BN
  const startRow = 4;
  const n = rows.length;
  const rng = dash.getRange(startRow, col, n + 1, headers.length);
  rng.clearContent();
  dash.getRange(startRow, col, 1, headers.length).setValues([headers]).setFontWeight('bold').setFontSize(9);
  dash.getRange(startRow + 1, col, n, headers.length).setValues(rows);

  // Formats per row
  dash.getRange(startRow + 1, col + 1, 1, 1).setNumberFormat('0');       // First Deposits (#)
  dash.getRange(startRow + 2, col + 1, 1, 1).setNumberFormat('$#,##0');  // Payments (Sum $)

  return dash.getRange(startRow, col, n + 1, headers.length);
}

/** Remove charts that use our hidden data block (BA:CB) */
function removeChartsByTitles_(sh, titles) {
  const charts = sh.getCharts();
  const FIRST = CHART_DATA_START_COL;  // 53 (BA)
  const LAST  = 80;                    // (extend if you add tables further right)
  for (const c of charts) {
    try {
      const ranges = c.getRanges ? c.getRanges() : [];
      const usesBlock = ranges.some(r =>
        r.getSheet().getName() === sh.getName() &&
        r.getColumn() >= FIRST && (r.getColumn() + r.getNumColumns() - 1) <= LAST
      );
      if (usesBlock) sh.removeChart(c);
    } catch (_) {}
  }
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


