// ============================================================
//  PROJECT #17 – SLIDES GENERATION (final)
//  Vị trí : ClientFolder (Julia/) — ngang cấp Prospects
//  Tên file: "HPUSA Julia Slides" / "VVS Julia Slides"
//  Logic  : 1 file duy nhất per client, không tạo lại nếu đã có
// ============================================================


// ── 1. LẤY TEMPLATE ID THEO BRAND ───────────────────────────

function slidesTemplateIdForBrand_(brand) {
  const SP = PropertiesService.getScriptProperties();
  if (brand === 'HPUSA') return SP.getProperty('SLIDES_TEMPLATE_ID_HPUSA') || '';
  if (brand === 'VVS')   return SP.getProperty('SLIDES_TEMPLATE_ID_VVS')   || '';
  return '';
}


// ── 2. CORE FUNCTION ─────────────────────────────────────────
// clientFolder = Julia/ (folder của KH, ngang cấp Prospects)

function generateSlidesForRow_(clientFolder, data) {
  const brand = String(data.Brand        || '').trim().toUpperCase();
  const name  = String(data.CustomerName || '').trim();

  if (!brand || !name) {
    Logger.log('[Slides] Thiếu Brand hoặc CustomerName — bỏ qua');
    return;
  }

  // Tên file theo đúng format trong ảnh: "HPUSA Julia Slides"
  const fileName = `${brand} ${name} Slides`;

  const tplId = slidesTemplateIdForBrand_(brand);
  if (!tplId) {
    Logger.log(`[Slides] Không có template cho brand "${brand}" — bỏ qua`);
    return;
  }

  // 1 file per client — không tạo lại nếu đã có
  const existing = clientFolder.getFilesByName(fileName);
  if (existing.hasNext()) {
    Logger.log(`[Slides] Đã tồn tại: "${fileName}" — bỏ qua`);
    return;
  }

  // Copy template vào ClientFolder
  const copy = DriveApp.getFileById(tplId).makeCopy(fileName, clientFolder);
  const pres = SlidesApp.openById(copy.getId());

  // Điền tên KH vào welcome slide
  if (name) _insertClientName_(pres.getSlides()[0], name);

  // Đảm bảo đủ 10 blank slides
  _ensureTenBlankSlides_(pres);

  pres.saveAndClose();
  Logger.log(`[Slides] ✅ Đã tạo: "${fileName}" trong "${clientFolder.getName()}"`);
}


// ── 3. INSERT TÊN KH ─────────────────────────────────────────

function _insertClientName_(welcomeSlide, customerName) {
  if (!welcomeSlide) return;
  let replaced = false;
  welcomeSlide.getShapes().forEach(shape => {
    if (!shape.getText) return;
    const tf = shape.getText();
    if (tf.asString().includes('{{CustomerName}}')) {
      tf.replaceAllText('{{CustomerName}}', customerName);
      replaced = true;
    }
  });
  if (!replaced) Logger.log('[Slides] ⚠️  Không tìm thấy {{CustomerName}} trên welcome slide');
}


// ── 4. ĐẢM BẢO 10 BLANK SLIDES ──────────────────────────────

function _ensureTenBlankSlides_(pres) {
  const needed = 10 - (pres.getSlides().length - 1);
  for (let i = 0; i < needed; i++) {
    pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  }
  if (needed > 0) Logger.log(`[Slides] Đã thêm ${needed} blank slides`);
}


// ── 5. TÍCH HỢP VÀO _ensureArtifactsForRowImpl_() ────────────
//
//  Tìm đoạn "── 4. ATOMIC WRITE" trong resolver.gs
//  Thêm đoạn này NGAY TRƯỚC nó:
//
//  // Google Slides (Project #17) — lưu vào ClientFolder
//  try {
//    if (clientFolder) {
//      generateSlidesForRow_(clientFolder, buildIntakeData_(row));
//    }
//  } catch (e) {
//    Logger.log('[Slides] ERROR: ' + e.message);
//  }
//
//  Lưu ý: biến "clientFolder" đã có sẵn trong _ensureArtifactsForRowImpl_
//  vì nó được tạo ở phần "── 1. CLIENT FOLDER"
// ─────────────────────────────────────────────────────────────


// ── 6. SET TEMPLATE IDs ──────────────────────────────────────

function project17_setTemplateIds() {
  const SP = PropertiesService.getScriptProperties();

  SP.setProperty('SLIDES_TEMPLATE_ID_HPUSA', '1xyjqqORgy2d9m2bniWreGqXedHlK5xG6SrUo-iU2vIM');
  SP.setProperty('SLIDES_TEMPLATE_ID_VVS',   '1GzBBcxYxKQjhXlIJEHcA8oQjppYWYoB0Rbk0mwUE6Bs');

  ['SLIDES_TEMPLATE_ID_HPUSA', 'SLIDES_TEMPLATE_ID_VVS'].forEach(key => {
    const id = SP.getProperty(key);
    Logger.log(`🔍 ${key} → id="${id}" | length=${id ? id.length : 'null'}`); // <-- thêm dòng này
    try {
      Logger.log(`✅ ${key}: "${DriveApp.getFileById(id).getName()}"`);
    } catch (e) {
      Logger.log(`❌ ${key}: ${e.message}`); // <-- in lỗi gốc
    }
  });
}


// ── 7. TEST: 1 row cụ thể ────────────────────────────────────

function project17_testSingleRow() {
  const TARGET_ROW = 2; // ← đổi số hàng muốn test

  const s    = SH(SHT.MASTER);
  const H    = headers_(SHT.MASTER);
  const data = buildIntakeData_(TARGET_ROW);

  const cfId = H['ClientFolderID']
    ? String(s.getRange(TARGET_ROW, H['ClientFolderID']).getValue() || '').trim()
    : '';

  if (!cfId) {
    Logger.log('❌ Chưa có ClientFolderID — chạy ensureArtifactsForRow_ trước');
    return;
  }

  generateSlidesForRow_(DriveApp.getFolderById(cfId), data);
}


// ── 8. BACKFILL: tạo slides cho client chưa có ───────────────

function project17_backfillSlides() {
  const s    = SH(SHT.MASTER);
  const H    = headers_(SHT.MASTER);
  const last = lastDataRow_(SHT.MASTER, LASTROW_SENTINELS);

  const colBrand = H['Brand']          || 0;
  const colAppt  = H['APPT_ID']        || 0;
  const colCfId  = H['ClientFolderID'] || 0;
  const colName  = H['Customer Name']  || 0;

  if (!colBrand || !colCfId || !colName) {
    Logger.log('❌ Thiếu cột Brand / ClientFolderID / Customer Name');
    return;
  }

  // Track theo ClientFolderID để không lặp lại cùng 1 client
  const processed = new Set();
  let done = 0, skipped = 0, errors = 0;

  for (let row = 2; row <= last; row++) {
    const brand = String(s.getRange(row, colBrand).getValue() || '').trim().toUpperCase();
    const cfId  = String(s.getRange(row, colCfId).getValue()  || '').trim();
    const name  = String(s.getRange(row, colName).getValue()  || '').trim();

    if (!brand || !cfId || !name) { skipped++; continue; }
    if (brand !== 'HPUSA' && brand !== 'VVS') { skipped++; continue; }

    // Bỏ qua nếu client này đã được xử lý rồi (repeat customer)
    const clientKey = `${brand}|${cfId}`;
    if (processed.has(clientKey)) { skipped++; continue; }
    processed.add(clientKey);

    try {
      const clientFolder = DriveApp.getFolderById(cfId);
      const fileName     = `${brand} ${name} Slides`;

      if (clientFolder.getFilesByName(fileName).hasNext()) {
        Logger.log(`⏭️  Đã có: "${fileName}"`);
        skipped++;
        continue;
      }

      generateSlidesForRow_(clientFolder, buildIntakeData_(row));
      done++;
      Utilities.sleep(1500);
    } catch (e) {
      errors++;
      Logger.log(`❌ Row ${row}: ${e.message}`);
    }
  }

  Logger.log(`\n── BACKFILL XONG ──`);
  Logger.log(`  Đã tạo : ${done}`);
  Logger.log(`  Bỏ qua : ${skipped}`);
  Logger.log(`  Lỗi    : ${errors}`);
}