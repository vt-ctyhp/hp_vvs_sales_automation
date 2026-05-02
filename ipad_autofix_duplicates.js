/*** ipad_autofix_duplicates.gs — v1.0
 * ══════════════════════════════════════════════════════════════════════
 * Batch auto-fix: 178 duplicate Root Appointment ID conflicts
 *
 * HOW TO USE:
 *   Step 1 → Run ipad_autofix_DryRun()          (đọc log, không ghi gì)
 *   Step 2 → Run ipad_autofix_Commit()           (ghi thật vào sheet)
 *   Step 3 → Run ipad_findDuplicateContacts()    (kiểm tra còn lại)
 *
 * STRATEGY:
 *   • Nhóm rows theo email → sau đó phone
 *   • Mỗi nhóm: chọn rootApptId SỚM NHẤT làm canonical
 *     (format AP-YYYYMMDD-NNN → sort chữ = sort theo ngày)
 *   • Relink tất cả rows còn lại về canonical root đó
 *
 * SKIP LIST: email dùng chung bởi nhiều khách khác nhau hoặc test data
 * MANUAL REVIEW: script log cảnh báo nhưng không tự sửa
 * IDEMPOTENT: chạy lại nhiều lần vẫn an toàn
 * ══════════════════════════════════════════════════════════════════════
 */

// ── Emails có NHIỀU KHÁCH KHÁC NHAU dùng chung → KHÔNG tự merge ───────
var SKIP_EMAILS = [
  'vvsjewelco@gmail.com',      // Kenny Huynh + Erika & Enrique + Charlene (email cửa hàng)
  'sales@ctyhp.com',            // Angela & Isaiah + Sam & Evan + Dorothy & Brandon
  'paulpasaoa2@gmail.com',      // Records test
  'adrian@hungphatusa.com',     // Records test staff
  'thanhvu@ctyhp.vn',           // Records test AI
  'majoshuabenitez@gmail.com',  // TEST MARIA
];

// ── Emails cần xem lại tay trước khi merge ────────────────────────────
var MANUAL_REVIEW_EMAILS = [
  'frenda@gmail.com', // "Frenda Chek (2)" vs "Frenda Chek's Mom" — cùng người hay khác?
];


// ═══════════════════════════════════════════════════════════════════════
// STEP 1 — DRY RUN
// ═══════════════════════════════════════════════════════════════════════
function ipad_autofix_DryRun() {
  Logger.log('══ DRY RUN — chỉ đọc log, không ghi gì ══');
  var plan = ipad_autofix_buildPlan_();
  ipad_autofix_logPlan_(plan);
  Logger.log('══ DRY RUN xong. %s relinks trong %s nhóm. Chạy ipad_autofix_Commit() để apply. ══',
    plan.totalRelinks, plan.groups.length);
}


// ═══════════════════════════════════════════════════════════════════════
// STEP 2 — COMMIT
// ═══════════════════════════════════════════════════════════════════════
function ipad_autofix_Commit() {
  var plan = ipad_autofix_buildPlan_();
  Logger.log('[Commit] %s relinks trong %s nhóm…', plan.totalRelinks, plan.groups.length);

  var ss      = SpreadsheetApp.getActive();
  var sh      = ss.getSheetByName(RP_MASTER_SHEET);
  var lc      = sh.getLastColumn();
  var hdr     = sh.getRange(1, 1, 1, lc).getDisplayValues()[0];
  var colMap  = rp_headerMap([hdr]);
  var apptIdx = rp_pick0(colMap, 'APPT_ID', 'RootApptID', 'Root Appt ID');

  if (apptIdx < 0) {
    Logger.log('❌ Không tìm thấy cột RootApptID — dừng.');
    return;
  }
  var apptCol = apptIdx + 1; // 1-indexed

  var applied = 0, alreadyOk = 0, errors = 0;

  for (var g = 0; g < plan.groups.length; g++) {
    var ops = plan.groups[g].ops;
    for (var o = 0; o < ops.length; o++) {
      var op = ops[o];
      if (op.alreadyCorrect) { alreadyOk++; continue; }
      try {
        sh.getRange(op.rowIndex, apptCol).setValue(op.newRoot);
        Logger.log('  ✅ row%s "%s"  %s → %s',
          op.rowIndex, op.name, op.oldRoot || '(trống)', op.newRoot);
        applied++;
      } catch (e) {
        Logger.log('  ❌ row%s FAILED: %s', op.rowIndex, e.message);
        errors++;
      }
    }
  }

  SpreadsheetApp.flush();
  Logger.log('══ COMMIT xong: ✅ Đã ghi %s | ⏭ Đã đúng %s | ❌ Lỗi %s ══',
    applied, alreadyOk, errors);
  Logger.log('Chạy ipad_findDuplicateContacts() để kiểm tra conflicts còn lại.');
}


// ═══════════════════════════════════════════════════════════════════════
// BUILD PLAN
// ═══════════════════════════════════════════════════════════════════════
function ipad_autofix_buildPlan_() {
  var ss  = SpreadsheetApp.getActive();
  var sh  = ss.getSheetByName(RP_MASTER_SHEET);
  var lr  = sh.getLastRow(), lc = sh.getLastColumn();
  var hdr = sh.getRange(1, 1, 1, lc).getDisplayValues()[0];
  var map = rp_headerMap([hdr]);

  var nameIdx  = rp_pick0(map, 'Customer Name', 'Customer', 'Client Name');
  var emailIdx = rp_pick0(map, 'Email', 'Email Address', 'E-mail');
  var phoneIdx = rp_pick0(map, 'Phone', 'Phone Number', 'Tel', 'Mobile');
  var apptIdx  = rp_pick0(map, 'APPT_ID', 'RootApptID', 'Root Appt ID');
  var brandIdx = rp_pick0(map, 'Brand');

  var vals = sh.getRange(2, 1, lr - 1, lc).getDisplayValues();

  // Index tất cả rows
  var byEmail = {}, byPhone = {};
  for (var i = 0; i < vals.length; i++) {
    var row   = vals[i];
    var ri    = i + 2;
    var name  = nameIdx  >= 0 ? String(row[nameIdx]  || '').trim() : '';
    var email = emailIdx >= 0 ? ipad_normalizeEmail_(row[emailIdx])  : '';
    var phone = phoneIdx >= 0 ? ipad_normalizePhone_(row[phoneIdx])  : '';
    var root  = apptIdx  >= 0 ? String(row[apptIdx]  || '').trim() : '';
    var brand = brandIdx >= 0 ? String(row[brandIdx] || '').trim() : '';
    if (!name) continue;

    var info = { rowIndex: ri, name: name, root: root, brand: brand, email: email, phone: phone };
    if (email && email.includes('@')) {
      if (!byEmail[email]) byEmail[email] = [];
      byEmail[email].push(info);
    }
    if (phone && phone.length >= 10) {
      if (!byPhone[phone]) byPhone[phone] = [];
      byPhone[phone].push(info);
    }
  }

  var groups = [], processedRows = {}, totalRelinks = 0;

  // --- EMAIL GROUPS ---
  Object.keys(byEmail).forEach(function(email) {
    var rows = byEmail[email];
    if (rows.length < 2) return;

    if (SKIP_EMAILS.indexOf(email) >= 0) {
      Logger.log('[SKIP] %s → %s rows (email trong danh sách bỏ qua)', email, rows.length);
      return;
    }
    if (MANUAL_REVIEW_EMAILS.indexOf(email) >= 0) {
      Logger.log('[MANUAL REVIEW ⚠] %s → cần xem tay:\n  %s', email,
        rows.map(function(r){ return 'row'+r.rowIndex+' "'+r.name+'" root:'+r.root; }).join('\n  '));
      return;
    }

    var roots = rows.map(function(r){ return r.root; }).filter(Boolean);
    var uniqueRoots = ipad_unique_(roots);
    if (uniqueRoots.length <= 1) return; // đã đồng nhất

    var canonical = ipad_pickEarliestRoot_(roots);
    var ops = [];
    rows.forEach(function(r) {
      processedRows[r.rowIndex] = canonical;
      var already = (r.root === canonical);
      ops.push({ rowIndex: r.rowIndex, name: r.name, oldRoot: r.root, newRoot: canonical, alreadyCorrect: already });
      if (!already) totalRelinks++;
    });
    groups.push({ type: 'email', key: email, canonical: canonical, ops: ops });
  });

  // --- PHONE GROUPS (chỉ rows chưa được email group xử lý) ---
  Object.keys(byPhone).forEach(function(phone) {
    var rows2 = byPhone[phone];
    if (rows2.length < 2) return;

    var unhandled = rows2.filter(function(r){ return !(r.rowIndex in processedRows); });
    if (unhandled.length < 2) return;

    var roots2 = unhandled.map(function(r){ return r.root; }).filter(Boolean);
    var uniqueRoots2 = ipad_unique_(roots2);
    if (uniqueRoots2.length <= 1) return;

    // Kiểm tra tên — nếu quá khác nhau → cần xem tay
    if (!ipad_namesLikelySamePerson_(unhandled.map(function(r){ return r.name; }))) {
      Logger.log('[PHONE SKIP - tên khác nhau ⚠] %s → cần xem tay:\n  %s',
        phone,
        unhandled.map(function(r){ return 'row'+r.rowIndex+' "'+r.name+'" root:'+r.root; }).join('\n  '));
      return;
    }

    var canonical2 = ipad_pickEarliestRoot_(roots2);
    var ops2 = [];
    unhandled.forEach(function(r) {
      var already = (r.root === canonical2);
      ops2.push({ rowIndex: r.rowIndex, name: r.name, oldRoot: r.root, newRoot: canonical2, alreadyCorrect: already });
      if (!already) totalRelinks++;
    });
    groups.push({ type: 'phone', key: phone, canonical: canonical2, ops: ops2 });
  });

  return { groups: groups, totalRelinks: totalRelinks };
}


// ═══════════════════════════════════════════════════════════════════════
// LOG PLAN
// ═══════════════════════════════════════════════════════════════════════
function ipad_autofix_logPlan_(plan) {
  Logger.log('═ KẾ HOẠCH: %s groups, %s relinks ═\n', plan.groups.length, plan.totalRelinks);
  plan.groups.forEach(function(group) {
    var changes = group.ops.filter(function(o){ return !o.alreadyCorrect; });
    if (changes.length === 0) return;
    Logger.log('[%s] %s  →  canonical: %s  (%s relinks)',
      group.type.toUpperCase(), group.key, group.canonical, changes.length);
    group.ops.forEach(function(op) {
      if (op.alreadyCorrect) {
        Logger.log('    ⏭  row%s "%s" — đã đúng', op.rowIndex, op.name);
      } else {
        Logger.log('    ✏️  row%s "%s"  [%s] → [%s]',
          op.rowIndex, op.name, op.oldRoot || 'trống', op.newRoot);
      }
    });
  });
}


// ═══════════════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════════════

function ipad_pickEarliestRoot_(roots) {
  var valid = roots.filter(function(r){ return r && r.length > 0; });
  if (!valid.length) return '';
  valid.sort();
  return valid[0];
}

function ipad_unique_(arr) {
  var seen = {}, out = [];
  for (var i = 0; i < arr.length; i++) {
    if (arr[i] && !(arr[i] in seen)) { seen[arr[i]] = 1; out.push(arr[i]); }
  }
  return out;
}

function ipad_namesLikelySamePerson_(names) {
  if (names.length < 2) return true;
  var firstWords = names.map(function(n){
    return String(n || '').trim().toLowerCase().split(/[\s&\/]+/)[0];
  }).filter(function(w){ return w.length > 1; });
  if (firstWords.length < 2) return true;
  var counts = {};
  firstWords.forEach(function(w){ counts[w] = (counts[w] || 0) + 1; });
  for (var w in counts) { if (counts[w] >= 2) return true; }
  return false;
}


// ═══════════════════════════════════════════════════════════════════════
// INSPECT HELPER — xem chi tiết 1 email trước khi quyết định
// Gọi thủ công từ editor: ipad_autofix_inspectEmail('abc@gmail.com')
// ═══════════════════════════════════════════════════════════════════════
function ipad_autofix_inspectEmail(email) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RP_MASTER_SHEET);
  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  var hdr = sh.getRange(1,1,1,lc).getDisplayValues()[0];
  var map = rp_headerMap([hdr]);

  var nameIdx  = rp_pick0(map,'Customer Name','Customer','Client Name');
  var emailIdx = rp_pick0(map,'Email','Email Address','E-mail');
  var phoneIdx = rp_pick0(map,'Phone','Phone Number','Tel','Mobile');
  var apptIdx  = rp_pick0(map,'APPT_ID','RootApptID','Root Appt ID');
  var brandIdx = rp_pick0(map,'Brand');
  var ptdIdx   = rp_pick0(map,'Paid-to-Date','Paid-To-Date','Paid to Date');
  var otIdx    = map['Order Total'] != null ? map['Order Total'] : -1;

  var vals = sh.getRange(2,1,lr-1,lc).getDisplayValues();
  var normTarget = ipad_normalizeEmail_(email);
  Logger.log('══ INSPECT: %s ══', email);
  var found = 0;
  for (var i = 0; i < vals.length; i++) {
    var row = vals[i];
    if (emailIdx < 0) continue;
    if (ipad_normalizeEmail_(row[emailIdx]) !== normTarget) continue;
    found++;
    Logger.log('  row%s | %-8s | "%-30s" | root: %-22s | OT: %-10s | PTD: %-10s | phone: %s',
      i+2,
      brandIdx >= 0 ? String(row[brandIdx]||'').trim().substring(0,8) : '?',
      nameIdx  >= 0 ? String(row[nameIdx] ||'').trim() : '?',
      apptIdx  >= 0 ? String(row[apptIdx] ||'').trim() : '?',
      otIdx    >= 0 ? String(row[otIdx]   ||'').trim() : '?',
      ptdIdx   >= 0 ? String(row[ptdIdx]  ||'').trim() : '?',
      phoneIdx >= 0 ? String(row[phoneIdx]||'').trim() : '?');
  }
  if (!found) Logger.log('  (không tìm thấy)');
}


// ═══════════════════════════════════════════════════════════════════════
// MANUAL FIX — dùng sau khi auto-fix để xử lý các trường hợp đặc biệt
// ═══════════════════════════════════════════════════════════════════════

/**
 * Fix tay một nhóm cụ thể.
 * Ví dụ:
 *   ipad_autofix_manualFix({ rowsToRelink:[193,195], newRoot:'AP-20251108-002' })
 */
function ipad_autofix_manualFix(params) {
  var rows    = params.rowsToRelink || [];
  var newRoot = String(params.newRoot || '').trim();
  var dryRun  = !!params.dryRun;
  if (!newRoot || !rows.length) { Logger.log('[manualFix] Thiếu params'); return; }
  Logger.log('[manualFix] %s | %s rows | dryRun=%s', newRoot, rows.length, dryRun);
  rows.forEach(function(ri) {
    var r = ipad_relinkRecord({ rowIndex: ri, newRootApptId: newRoot, dryRun: dryRun });
    if (r.ok) Logger.log('  %s row%s "%s": %s → %s', dryRun?'[DRY]':'✅', r.rowIndex, r.name, r.oldRoot, r.newRoot);
    else      Logger.log('  ❌ row%s: %s', ri, r.error);
  });
}


// ── Các fix đặc biệt (chạy SAU ipad_autofix_Commit) ───────────────────

/** vvsjewelco@gmail.com: chỉ merge Kenny Huynh rows 192,193,195 */
function ipad_fix_KennyHuynh() {
  Logger.log('Merge Kenny Huynh (rows 192,193,195) → AP-20251108-002');
  ipad_autofix_manualFix({ rowsToRelink:[192,193,195], newRoot:'AP-20251108-002' });
  Logger.log('Erika & Enrique (row194) và Charlene (row225,226) giữ nguyên root riêng.');
}

/** Rachelle & Eddie (row 438) → về root của Rachelle AP-20250913-007 */
function ipad_fix_RachelleEddie() {
  Logger.log('Merge "Rachelle & Eddie" row438 → root Rachelle AP-20250913-007');
  ipad_autofix_manualFix({ rowsToRelink:[438], newRoot:'AP-20250913-007' });
}

/** Inspect Kevin Liao phone conflict trước khi quyết định */
function ipad_fix_KevinLiao_inspect() {
  ipad_autofix_inspectEmail('kevinaliao@gmail.com');
  Logger.log('Phone: row544 "Kevin & Celine Liao" vs row547 "Kevin liao" — xem log rồi quyết định.');
  Logger.log('Nếu cùng người: ipad_autofix_manualFix({ rowsToRelink:[544], newRoot:"AP-20260314-003" })');
}

// FIX 1: Jerad Mariscal row518 chưa được merge
function fix_JeradMariscal_row518() {
  ipad_relinkRecord({ rowIndex: 518, newRootApptId: 'AP-20260327-003' });
}

// FIX 2: "Kcirde & Mikaela Lacsa" row545 → về root của Mikaela
function fix_MikaelaKcirde_row545() {
  ipad_relinkRecord({ rowIndex: 545, newRootApptId: 'AP-20260404-001' });
}

// FIX 3: "Kevin & Celine Liao" row544 → về root của Kevin liao
// CHỈ chạy nếu staff xác nhận đây là cùng một cặp/người
function fix_KevinCelineLiao_row544() {
  ipad_relinkRecord({ rowIndex: 544, newRootApptId: 'AP-20260314-003' });
}

// FIX 4: Charlene row226 → về root của Charlene row225
function fix_Charlene_row226() {
  ipad_relinkRecord({ rowIndex: 226, newRootApptId: 'AP-20251125-003' });
}
