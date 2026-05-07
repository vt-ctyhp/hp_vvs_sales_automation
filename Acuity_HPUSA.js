
// ─── CONFIG ───────────────────────────────────────────────────────────────────

const ACUITY_CFG = {
  BASE_URL: 'https://acuityscheduling.com/api/v1',
  COMPANY:  'HPUSA',
  TZ:       'America/Los_Angeles',

  VISIT_TYPE_MAP: {
    'In-Person Custom Design Consultation': 'Appointment',
    'Diamond Viewing Appointment':          'Diamond Viewing',
    'Diamond Viewing':                      'Diamond Viewing',
  },

  BUDGET_MAP: {
    '$1000 - $5000':     '$1,000 - $5,000',
    '$1,000 - $5,000':   '$1,000 - $5,000',
    'Under $5,000':      '$1,000 - $5,000',
    'Under $5000':       '$1,000 - $5,000',
    '$5000 - $10000':    '$5,001 - $10,000',
    '$5,000 - $10,000':  '$5,001 - $10,000',
    '$5001 - $10000':    '$5,001 - $10,000',
    '$10000 - $15000':   '$10,001 - $15,000',
    '$10,000 - $15,000': '$10,001 - $15,000',
    '$10000- $15000':    '$10,001 - $15,000',
    '$10000-$15000':     '$10,001 - $15,000',
    '$10000- $20000':    '$15,001 - $20,000',
    '$15000 - $20000':   '$15,001 - $20,000',
    '$15,000 - $20,000': '$15,001 - $20,000',
    '$15001 - $20000':   '$15,001 - $20,000',
    '$15001-$20000':     '$15,001 - $20,000',
    '$15000-$20000':     '$15,001 - $20,000',
    '$20000+':           '$20,001+',
    '$20,000+':          '$20,001+',
    '$20001+':           '$20,001+',
  },

  SOURCE_MAP: {
    'instagram': 'Instagram',
    'ig':        'Instagram',
    'tiktok':    'Tiktok',
    'tik tok':   'Tiktok',
    'facebook':  'Facebook',
    'fb':        'Facebook',
    'google':    'Google',
    'yelp':      'Yelp',
    'referral':  'Referral',
    'friend':    'Referral',
    'reddit':    'Reddit',
    'shopify':   'Shopify',
  },

  DIAMOND_MAP: {
    'natural diamond': 'Natural Diamond',
    'natural':         'Natural Diamond',
    'lab diamond':     'Lab Diamond',
    'lab-grown':       'Lab Diamond',
    'lab grown':       'Lab Diamond',
    'lab':             'Lab Diamond',
  },
};

var ACUITY_ACTIVE_EXISTING_CURSOR_PROP = 'ACUITY_ACTIVE_EXISTING_CURSOR';
var ACUITY_CANCELED_EXISTING_CURSOR_PROP = 'ACUITY_CANCELED_EXISTING_CURSOR';
var ACUITY_ACTIVE_EXISTING_BATCH_SIZE = 8;
var ACUITY_CANCELED_EXISTING_BATCH_SIZE = 8;

// ─── TRIGGER MANAGEMENT ───────────────────────────────────────────────────────

function installAcuityTrigger() {
  if (typeof sw_installBackgroundOrchestratorTrigger === 'function') {
    return sw_installBackgroundOrchestratorTrigger();
  }
  Logger.log('Background orchestrator unavailable; Acuity trigger not installed.');
  return { ok: false, error: 'sw_installBackgroundOrchestratorTrigger unavailable' };
}

function uninstallAcuityTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'acuityPollAndSubmit')
    .forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('Trigger removed');
}

// ─── MAIN POLLER ──────────────────────────────────────────────────────────────

function acuityPollAndSubmit(e) {
  const redirected = typeof swOrchRedirectLegacyTrigger_ === 'function'
    ? swOrchRedirectLegacyTrigger_('acuityPollAndSubmit', e)
    : null;
  if (redirected) return redirected;

  const SP   = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('acuityPollAndSubmit: locked by another run, skipping');
    return { ok: true, skipped: true, reason: 'LOCK_BUSY' };
  }

  try {
    const userId = SP.getProperty('ACUITY_USER_ID');
    const apiKey = SP.getProperty('ACUITY_API_KEY');
    const formId = SP.getProperty('FORM_ID');
    if (!userId || !apiKey || !formId) {
      throw new Error('Missing Script Properties: ACUITY_USER_ID / ACUITY_API_KEY / FORM_ID');
    }

    // Load UIDs đã có trên Master (cache 55s)
    const existingUIDs = acuityGetMasterUIDs_();
    Logger.log('Existing UIDs on Master: ' + existingUIDs.size);
    const cancelableUIDs = acuityGetCancelableMasterUIDs_();

    const lists = acuityFetchAppointmentLists_(userId, apiKey);
    const activeList = lists.activeList || [];
    const canceledList = lists.canceledList || [];
    Logger.log('Fetched: active=' + activeList.length + ' canceled=' + canceledList.length);

    let submitted = 0, rescheduled = 0, edited = 0, canceled = 0, skipped = 0, errors = 0;
    let checkedExisting = 0, checkedCanceled = 0;
    const activeExisting = [];
    const newActive = [];
    const canceledExisting = [];

    activeList.forEach(function (appt) {
      const uid = String(appt.id);
      if (existingUIDs.has(uid)) {
        activeExisting.push(appt);
        return;
      }
      if (SP.getProperty('ACUITY:DONE:' + uid)) {
        skipped++;
        return;
      }
      newActive.push(appt);
    });

    canceledList.forEach(function (appt) {
      const uid = String(appt.id);
      if (cancelableUIDs.has(uid)) canceledExisting.push(appt);
      else skipped++;
    });

    for (const apptRef of newActive) {
      try {
        const appt = acuityFetchAppointmentDetail_(userId, apiKey, apptRef);
        const uid = String(appt.id);
        const fieldMap = acuityToFormFieldMap_(appt);
        acuitySubmitToForm_(formId, fieldMap);
        SP.setProperty('ACUITY:DONE:' + uid, '1');
        existingUIDs.add(uid);
        CacheService.getScriptCache().remove('MASTER_UIDS_CACHE');
        submitted++;
        Logger.log('✅ Submitted: ' + uid + ' | ' + appt.firstName + ' ' + appt.lastName);
      } catch (err) {
        errors++;
        Logger.log('❌ Error new appt ' + (apptRef && apptRef.id) + ': ' + (err && err.message || err));
      }
    }

    if (!submitted) {
      const activeBatch = acuityRotatingBatch_(activeExisting, ACUITY_ACTIVE_EXISTING_CURSOR_PROP, ACUITY_ACTIVE_EXISTING_BATCH_SIZE);
      for (const apptRef of activeBatch.items) {
        try {
          const appt = acuityFetchAppointmentDetail_(userId, apiKey, apptRef);
          const result = acuityHandleExisting_(appt, formId);
          checkedExisting++;
          if (result === 'rescheduled' || result === 'edited') {
            CacheService.getScriptCache().remove('MASTER_UIDS_CACHE');
          }
          if (result === 'rescheduled') {
            rescheduled++;
            break;
          }
          if (result === 'edited') edited++;
        } catch (err) {
          errors++;
          Logger.log('❌ Error existing appt ' + (apptRef && apptRef.id) + ': ' + (err && err.message || err));
        }
      }
    }

    if (!submitted && !rescheduled) {
      const canceledBatch = acuityRotatingBatch_(canceledExisting, ACUITY_CANCELED_EXISTING_CURSOR_PROP, ACUITY_CANCELED_EXISTING_BATCH_SIZE);
      for (const appt of canceledBatch.items) {
        try {
          const uid = String(appt.id);
          checkedCanceled++;
          if (acuityCancelOnMaster_(uid)) {
            CacheService.getScriptCache().remove('MASTER_UIDS_CACHE');
            canceled++;
          } else {
            skipped++;
          }
        } catch (err) {
          errors++;
          Logger.log('❌ Error canceled appt ' + (appt && appt.id) + ': ' + (err && err.message || err));
        }
      }
    }

    SP.setProperty('ACUITY_LAST_FETCH', new Date().toISOString());
    skipped += Math.max(0, activeExisting.length - checkedExisting) + Math.max(0, canceledExisting.length - checkedCanceled);
    Logger.log('Done — submitted=' + submitted + ' rescheduled=' + rescheduled + ' edited=' + edited + ' canceled=' + canceled + ' checkedExisting=' + checkedExisting + '/' + activeExisting.length + ' checkedCanceled=' + checkedCanceled + '/' + canceledExisting.length + ' skipped=' + skipped + ' errors=' + errors);
    return {
      ok: errors === 0,
      submitted: submitted,
      rescheduled: rescheduled,
      edited: edited,
      canceled: canceled,
      skipped: skipped,
      errors: errors,
      formSubmitted: submitted + rescheduled,
      checkedExisting: checkedExisting,
      existingCandidates: activeExisting.length,
      checkedCanceled: checkedCanceled,
      canceledCandidates: canceledExisting.length,
      deferredExisting: Math.max(0, activeExisting.length - checkedExisting),
      deferredCanceled: Math.max(0, canceledExisting.length - checkedCanceled),
      checkedAt: new Date().toISOString()
    };

  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

// ─── ACUITY API ───────────────────────────────────────────────────────────────

function acuityFetchAppointments_(userId, apiKey) {
  const lists = acuityFetchAppointmentLists_(userId, apiKey);
  const active = lists.activeList.map(a => acuityFetchAppointmentDetail_(userId, apiKey, a));
  const allAppts = [...active, ...lists.canceledList];
  const seen = new Set();
  return allAppts.filter(a => {
    const key = String(a.id);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function acuityFetchAppointmentLists_(userId, apiKey) {
  const todayStart = new Date();
  todayStart.setHours(0, 0, 0, 0);
  const pastStart = new Date(todayStart.getTime() - 7  * 24 * 3600 * 1000);
  const future    = new Date(todayStart.getTime() + 60 * 24 * 3600 * 1000);

  const params = [
    'minDate=' + encodeURIComponent(Utilities.formatDate(pastStart, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'")),
    'maxDate=' + encodeURIComponent(Utilities.formatDate(future,    'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'")),
    'max=100',
  ];

  const auth    = 'Basic ' + Utilities.base64Encode(userId + ':' + apiKey);
  const baseUrl = ACUITY_CFG.BASE_URL + '/appointments?' + params.join('&');

  // Fetch active list
  const respActive = UrlFetchApp.fetch(baseUrl, {
    method: 'get', headers: { 'Authorization': auth }, muteHttpExceptions: true,
  });
  if (respActive.getResponseCode() !== 200) {
    throw new Error('Acuity API error ' + respActive.getResponseCode() + ': ' + respActive.getContentText());
  }
  const activeList = JSON.parse(respActive.getContentText());

  // Fetch canceled list
  const respCanceled = UrlFetchApp.fetch(baseUrl + '&canceled=true', {
    method: 'get', headers: { 'Authorization': auth }, muteHttpExceptions: true,
  });
  if (respCanceled.getResponseCode() !== 200) {
    throw new Error('Acuity API (canceled) error ' + respCanceled.getResponseCode());
  }
  const canceledList = JSON.parse(respCanceled.getContentText());
  canceledList.forEach(a => { a.canceled = true; });

  Logger.log('Active: ' + activeList.length + ' | Canceled: ' + canceledList.length);
  return { activeList: activeList, canceledList: canceledList };
}

function acuityFetchAppointmentDetail_(userId, apiKey, appt) {
  const auth = 'Basic ' + Utilities.base64Encode(userId + ':' + apiKey);
  try {
    const r = UrlFetchApp.fetch(
      ACUITY_CFG.BASE_URL + '/appointments/' + appt.id,
      { method: 'get', headers: { 'Authorization': auth }, muteHttpExceptions: true }
    );
    if (r.getResponseCode() === 200) return JSON.parse(r.getContentText());
  } catch(_) {}
  return appt;
}

function acuityRotatingBatch_(items, cursorProp, batchSize) {
  items = items || [];
  batchSize = Math.max(1, Number(batchSize || 1) || 1);
  if (items.length <= batchSize) {
    try { PropertiesService.getScriptProperties().setProperty(cursorProp, '0'); } catch (_) {}
    return { items: items.slice(), start: 0, next: 0, total: items.length };
  }
  const props = PropertiesService.getScriptProperties();
  const start = Math.max(0, Number(props.getProperty(cursorProp) || 0) || 0) % items.length;
  const out = [];
  for (let i = 0; i < batchSize; i++) {
    out.push(items[(start + i) % items.length]);
  }
  const next = (start + batchSize) % items.length;
  props.setProperty(cursorProp, String(next));
  return { items: out, start: start, next: next, total: items.length };
}

// ─── MASTER HELPERS ───────────────────────────────────────────────────────────

function acuityGetMasterUIDs_() {
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'MASTER_UIDS_CACHE';
  const cached   = cache.get(cacheKey);

  if (cached) {
    Logger.log('UIDs from cache');
    return new Set(JSON.parse(cached));
  }

  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getSheetByName('00_Master Appointments');
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = hdr.indexOf('CalendlyEventUID') + 1;
  if (!col) return new Set();

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return new Set();

  const uids = sh.getRange(2, col, lastRow - 1, 1)
    .getValues().flat()
    .map(v => String(v).trim())
    .filter(Boolean);

  try { cache.put(cacheKey, JSON.stringify(uids), 55); } catch(_) {}
  Logger.log('UIDs from Master (cached 55s): ' + uids.length);

  return new Set(uids);
}

function acuityGetCancelableMasterUIDs_() {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getSheetByName('00_Master Appointments');
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const H   = {};
  hdr.forEach((h, i) => { if (h) H[String(h).trim()] = i + 1; });

  if (!H['CalendlyEventUID'] || !H['Status']) return new Set();

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return new Set();

  const uids = sh.getRange(2, H['CalendlyEventUID'], lastRow - 1, 1).getValues();
  const statuses = sh.getRange(2, H['Status'], lastRow - 1, 1).getValues();
  const cancelable = [];

  for (let i = 0; i < uids.length; i++) {
    const uid = String(uids[i][0] || '').trim();
    if (!uid) continue;
    const status = String(statuses[i][0] || '').trim();
    if (/canceled|rescheduled/i.test(status)) continue;
    cancelable.push(uid);
    const rescheduleIndex = uid.indexOf('_R');
    if (rescheduleIndex > 0) cancelable.push(uid.slice(0, rescheduleIndex));
  }

  return new Set(cancelable);
}

function acuityCancelOnMaster_(acuityUid) {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getSheetByName('00_Master Appointments');
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const H   = {};
  hdr.forEach((h, i) => { if (h) H[String(h).trim()] = i + 1; });

  if (!H['CalendlyEventUID'] || !H['Status']) return false;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;

  const uids = sh.getRange(2, H['CalendlyEventUID'], lastRow - 1, 1).getValues();

  // Pass 1: tìm dòng _R... active mới nhất
  let masterRow = 0;
  for (let i = 0; i < uids.length; i++) {
    const rowUid = String(uids[i][0]).trim();
    if (!rowUid.startsWith(acuityUid + '_R')) continue;
    const st = String(sh.getRange(i + 2, H['Status']).getValue() || '').trim();
    if (/canceled|rescheduled/i.test(st)) continue;
    masterRow = i + 2; // không break → lấy dòng cuối cùng
  }

  // Pass 2: nếu không có _R → tìm dòng gốc
  if (!masterRow) {
    for (let i = 0; i < uids.length; i++) {
      const rowUid = String(uids[i][0]).trim();
      if (rowUid !== acuityUid) continue;
      const st = String(sh.getRange(i + 2, H['Status']).getValue() || '').trim();
      if (/canceled|rescheduled/i.test(st)) continue;
      masterRow = i + 2;
      break;
    }
  }

  if (!masterRow) {
    Logger.log('Cancel: no active row found for uid=' + acuityUid);
    return false;
  }

  const curStatus = String(sh.getRange(masterRow, H['Status']).getValue() || '').trim();
  if (/canceled/i.test(curStatus)) {
    Logger.log('Already canceled, skip: row=' + masterRow + ' uid=' + acuityUid);
    return false;
  }

  sh.getRange(masterRow, H['Status']).setValue('Canceled');
  if (H['Active?'])          sh.getRange(masterRow, H['Active?']).setValue('No');
  if (H['CanceledAt'] && !sh.getRange(masterRow, H['CanceledAt']).getValue()) {
    sh.getRange(masterRow, H['CanceledAt']).setValue(new Date());
  }
  if (H['Automation Notes']) {
    const prev = sh.getRange(masterRow, H['Automation Notes']).getValue() || '';
    const note = 'Canceled via Acuity @ ' + new Date().toISOString();
    sh.getRange(masterRow, H['Automation Notes']).setValue(prev ? prev + '\n' + note : note);
  }
  try {
    if (typeof swInboxLogAppointmentScheduleChangeFromRows_ === 'function') {
      swInboxLogAppointmentScheduleChangeFromRows_(swSpreadsheet_(), 'APPOINTMENT_CANCELED', masterRow, 0);
    }
  } catch (inboxErr) {
    Logger.log('Inbox cancel notification failed: ' + (inboxErr && inboxErr.message || inboxErr));
  }
  Logger.log('Canceled on Master: row=' + masterRow + ' uid=' + acuityUid);
  return true;
}

// ─── NORMALIZE ────────────────────────────────────────────────────────────────

function acuityToFormFieldMap_(appt) {
  const forms    = appt.forms || [];
  const formData = {};
  for (const form of forms) {
    for (const field of (form.values || [])) {
      if (field.name && field.value !== undefined) {
        formData[field.name.trim()] = field.value;
      }
    }
  }

  const firstName = String(appt.firstName || '').trim();
  const lastName  = String(appt.lastName  || '').trim();
  const name      = [firstName, lastName].filter(Boolean).join(' ');
  const email     = String(appt.email || '').trim();
  const phone     = acuityExtractPhone_(appt, formData);

  const apptTypeName = String(appt.type || '').trim();
  const visitType    = ACUITY_CFG.VISIT_TYPE_MAP[apptTypeName] || 'Appointment';

  const startISO  = appt.datetime || appt.date || '';
  const startDate = startISO ? new Date(startISO) : null;
  const visitDate = startDate || null;
  const visitTime = startDate
    ? Utilities.formatDate(startDate, ACUITY_CFG.TZ, 'HH:mm')
    : '';

  const location    = acuityExtractLocation_(appt);
  const budgetRaw   = acuityFindAnswer_(formData, ['What is your preferred price range?', 'preferred price range', 'price range', 'budget']);
  const sourceRaw   = acuityFindAnswer_(formData, ['How did you hear about us?', 'hear about us', 'source']);
  const diamondRaw  = acuityFindAnswer_(formData, ['What is your preferred diamond type?', 'preferred diamond type', 'diamond type']);
  const designNotes = acuityFindAnswer_(formData, ['Do you have a design in mind?', 'design in mind', 'ring design', 'style notes']);
  const diamondLink = acuityFindAnswer_(formData, ['diamond link', 'link']);

  const budgetNorm  = acuityNormalizeBudget_(budgetRaw);
  const sourceNorm  = acuityNormalizeSource_(sourceRaw);
  const diamondNorm = acuityNormalizeDiamond_(diamondRaw);

  const styleNotes = [
    designNotes,
    diamondLink ? 'Diamond link: ' + diamondLink : '',
  ].filter(Boolean).join('\n\n');

  return {
    'Company':                   ACUITY_CFG.COMPANY,
    'Customer Name':              name,
    'Phone':                      phone,
    'Email':                      email,
    'Visit Type':                 visitType,
    'Visit Date':                 visitDate,
    'Visit Time':                 visitTime,
    'Location':                   location,
    'Diamond Type':               diamondNorm  ? [diamondNorm]  : [],
    'Budget Range':               budgetNorm   ? [budgetNorm]   : [],
    'Source':                     sourceNorm   ? [sourceNorm]   : ['Did not disclose'],
    'Style Notes':                styleNotes,
    'Admin: Calendly Event UID':  String(appt.id || ''),
  };
}

function acuityStableRescheduleUid_(appt) {
  const base = String(appt && appt.id || '').trim();
  const startISO = (appt && (appt.datetime || appt.date)) || '';
  const dt = startISO ? new Date(startISO) : null;
  const stamp = dt && !isNaN(dt.getTime())
    ? Utilities.formatDate(dt, ACUITY_CFG.TZ, 'yyyyMMddHHmmss')
    : 'unknown';
  return base ? base + '_R' + stamp : '';
}

function acuityNormEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function acuityNormPhone_(value) {
  let d = String(value || '').replace(/\D+/g, '');
  if (d.length > 10 && d[0] === '1') d = d.slice(1);
  return d.length >= 7 ? d : '';
}

// ─── FIELD EXTRACT HELPERS ────────────────────────────────────────────────────

function acuityExtractPhone_(appt, formData) {
  for (const c of [appt.phone, appt.smsNumber, appt.textReminderNumber]) {
    const clean = acuityCleanPhone_(c);
    if (clean) return clean;
  }
  for (const key of ['Send text messages to', 'Phone Number', 'Phone', 'Mobile', "Partner's Phone Number"]) {
    if (formData[key]) {
      const clean = acuityCleanPhone_(formData[key]);
      if (clean) return clean;
    }
  }
  return '';
}

function acuityCleanPhone_(val) {
  if (!val) return '';
  const digits = String(val).trim().replace(/[^\d+]/g, '');
  if (!digits) return '';
  if (/^\+\d{11,}$/.test(digits)) return digits;
  if (/^\d{10}$/.test(digits))    return '+1' + digits;
  if (/^1\d{10}$/.test(digits))   return '+' + digits;
  return digits;
}

function acuityExtractLocation_(appt) {
  const s = String(appt.location || '').toLowerCase();
  return /virtual|zoom|phone|video|google meet/.test(s) ? 'Virtual' : 'In Store';
}

function acuityFindAnswer_(formData, keys) {
  for (const key of keys) {
    if (formData[key] !== undefined) return String(formData[key] || '').trim();
    const keyLower = key.toLowerCase();
    for (const k of Object.keys(formData)) {
      if (k.toLowerCase().includes(keyLower) || keyLower.includes(k.toLowerCase())) {
        return String(formData[k] || '').trim();
      }
    }
  }
  return '';
}

function acuityNormalizeBudget_(raw) {
  if (!raw) return '';
  const s = raw.trim().replace(/\s*-\s*/g, ' - ').replace(/\s+/g, ' ');
  if (ACUITY_CFG.BUDGET_MAP[s])          return ACUITY_CFG.BUDGET_MAP[s];
  if (ACUITY_CFG.BUDGET_MAP[raw.trim()]) return ACUITY_CFG.BUDGET_MAP[raw.trim()];

  const nums = raw.replace(/[^\d]/g, ' ').trim().split(/\s+/)
    .map(Number).filter(n => n > 0).sort((a, b) => a - b);
  if (!nums.length) return raw.trim();

  const low  = nums[0];
  const high = nums[nums.length - 1];
  if (high <= 5000)                                    return '$1,000 - $5,000';
  if (high <= 10000 || (low >= 5000  && high <= 10000)) return '$5,001 - $10,000';
  if (high <= 15000 || (low >= 10000 && high <= 15000)) return '$10,001 - $15,000';
  if (high <= 20000 || (low >= 15000 && high <= 20000)) return '$15,001 - $20,000';
  if (low >= 20000 || raw.includes('+'))                return '$20,001+';
  return raw.trim();
}

function acuityNormalizeSource_(raw) {
  if (!raw) return '';
  const lower = raw.trim().toLowerCase();
  if (ACUITY_CFG.SOURCE_MAP[lower]) return ACUITY_CFG.SOURCE_MAP[lower];
  for (const [k, v] of Object.entries(ACUITY_CFG.SOURCE_MAP)) {
    if (lower.includes(k)) return v;
  }
  return raw.trim();
}

function acuityNormalizeDiamond_(raw) {
  if (!raw) return '';
  const lower = raw.trim().toLowerCase();
  if (ACUITY_CFG.DIAMOND_MAP[lower]) return ACUITY_CFG.DIAMOND_MAP[lower];
  for (const [k, v] of Object.entries(ACUITY_CFG.DIAMOND_MAP)) {
    if (lower.includes(k)) return v;
  }
  return '';
}

// ─── FORM SUBMIT ──────────────────────────────────────────────────────────────

function acuitySubmitToForm_(formId, fieldMap) {
  const form  = FormApp.openById(formId);
  const items = form.getItems();
  const resp  = form.createResponse();
  let count   = 0;

  for (const item of items) {
    const title = (item.getTitle() || '').trim();
    if (!(title in fieldMap)) continue;

    const val = fieldMap[title];
    if (val === null || val === undefined || val === '') continue;
    if (Array.isArray(val) && val.length === 0) continue;

    try {
      switch (item.getType()) {

        case FormApp.ItemType.TEXT:
          resp.withItemResponse(item.asTextItem().createResponse(String(val)));
          count++; break;

        case FormApp.ItemType.PARAGRAPH_TEXT:
          resp.withItemResponse(item.asParagraphTextItem().createResponse(String(val)));
          count++; break;

        case FormApp.ItemType.DATE:
          if (val instanceof Date && !isNaN(val)) {
            resp.withItemResponse(item.asDateItem().createResponse(val));
            count++;
          }
          break;

        case FormApp.ItemType.TIME: {
          const m = /^(\d{1,2}):(\d{2})$/.exec(String(val));
          if (m) {
            resp.withItemResponse(item.asTimeItem().createResponse(parseInt(m[1], 10), parseInt(m[2], 10)));
            count++;
          }
          break;
        }

        case FormApp.ItemType.LIST: {
          const li      = item.asListItem();
          const choices = li.getChoices();
          const target  = String(Array.isArray(val) ? val[0] : val).trim().toLowerCase();
          const match   = choices.find(c =>
            c.getValue().trim().toLowerCase() === target ||
            c.getValue().trim().toLowerCase().includes(target) ||
            target.includes(c.getValue().trim().toLowerCase())
          );
          if (match) { resp.withItemResponse(li.createResponse(match.getValue())); count++; }
          else Logger.log('No LIST match for "' + title + '" → "' + target + '"');
          break;
        }

        case FormApp.ItemType.CHECKBOX: {
          const cb      = item.asCheckboxItem();
          const choices = cb.getChoices();
          const answers = Array.isArray(val) ? val : [val];
          const matched = answers.map(ans => {
            const ansLower = String(ans).trim().toLowerCase();
            return choices.find(c =>
              c.getValue().trim().toLowerCase() === ansLower ||
              c.getValue().trim().toLowerCase().includes(ansLower) ||
              ansLower.includes(c.getValue().trim().toLowerCase())
            );
          }).filter(Boolean).map(c => c.getValue());

          if (matched.length) { resp.withItemResponse(cb.createResponse(matched)); count++; }
          else Logger.log('No CHECKBOX match for "' + title + '" → ' + JSON.stringify(answers));
          break;
        }

        case FormApp.ItemType.MULTIPLE_CHOICE: {
          const mc      = item.asMultipleChoiceItem();
          const choices = mc.getChoices();
          const target  = String(Array.isArray(val) ? val[0] : val).trim().toLowerCase();
          const match   = choices.find(c => c.getValue().trim().toLowerCase() === target);
          if (match) { resp.withItemResponse(mc.createResponse(match.getValue())); count++; }
          break;
        }

        default:
          Logger.log('Skipped unsupported type for "' + title + '"');
      }
    } catch (err) {
      Logger.log('Error filling "' + title + '": ' + (err && err.message || err));
    }
  }

  resp.submit();
  Logger.log('Form submitted: ' + count + ' fields filled');
}


function acuityHandleExisting_(appt, formId) {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getSheetByName('00_Master Appointments');
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const H   = {};
  hdr.forEach((h, i) => { if (h) H[String(h).trim()] = i + 1; });

  if (!H['CalendlyEventUID']) return 'unchanged';

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 'unchanged';

  // uid gốc: "1684912230" → match cả "1684912230" và "1684912230_R..."
  const uids = sh.getRange(2, H['CalendlyEventUID'], lastRow - 1, 1).getValues();
  let masterRow = 0;

  // Pass 1: tìm dòng _R... (reschedule mới nhất) — không bị Rescheduled
  for (let i = 0; i < uids.length; i++) {
    const rowUid = String(uids[i][0]).trim();
    if (!rowUid.startsWith(String(appt.id) + '_R')) continue;
    if (H['Status']) {
      const st = String(sh.getRange(i + 2, H['Status']).getValue() || '').trim();
      if (/rescheduled|canceled/i.test(st)) continue;
    }
    masterRow = i + 2;
    // Không break — lấy dòng cuối cùng (mới nhất)
  }

  // Pass 2: nếu không có dòng _R → tìm dòng gốc không bị Rescheduled
  if (!masterRow) {
    for (let i = 0; i < uids.length; i++) {
      const rowUid = String(uids[i][0]).trim();
      if (rowUid !== String(appt.id)) continue;
      if (H['Status']) {
        const st = String(sh.getRange(i + 2, H['Status']).getValue() || '').trim();
        if (/rescheduled|canceled/i.test(st)) continue;
      }
      masterRow = i + 2;
      break;
    }
  }

  Logger.log('acuityHandleExisting_ uid=' + appt.id + ' masterRow=' + masterRow);
  if (!masterRow) return 'unchanged';

  const fieldMap = acuityToFormFieldMap_(appt);

  // ── Detect reschedule: so sánh Visit Date/Time ───────────────────
  let isReschedule = false;
  if (fieldMap['Visit Date'] instanceof Date && H['Visit Date'] && H['Visit Time']) {
    const newDate = Utilities.formatDate(fieldMap['Visit Date'], ACUITY_CFG.TZ, 'M/d/yyyy');
    const curDate = String(sh.getRange(masterRow, H['Visit Date']).getDisplayValue() || '').trim();
    const newTime = fieldMap['Visit Time']; // "11:00" (24h)
    const curTimeRaw = String(sh.getRange(masterRow, H['Visit Time']).getDisplayValue() || '').trim();

    // Normalize Master time về HH:mm để compare
    // "11:00:00 AM" → "11:00" | "1:00:00 PM" → "13:00"
    function parseToHHMM_(t) {
      const m12 = /^(\d{1,2}):(\d{2}):\d{2}\s*(AM|PM)$/i.exec(t);
      if (m12) {
        let h = parseInt(m12[1], 10);
        const min = m12[2];
        const ap  = m12[3].toUpperCase();
        if (ap === 'AM' && h === 12) h = 0;
        if (ap === 'PM' && h !== 12) h += 12;
        return String(h).padStart(2,'0') + ':' + min;
      }
      const m24 = /^(\d{1,2}):(\d{2})$/.exec(t);
      if (m24) return String(parseInt(m24[1],10)).padStart(2,'0') + ':' + m24[2];
      return t;
    }

    const curTime = parseToHHMM_(curTimeRaw); // normalize về "11:00"
    Logger.log('Date compare: cur="' + curDate + '" new="' + newDate + '" | Time: cur="' + curTime + '" (raw="' + curTimeRaw + '") new="' + newTime + '"');
    if (newDate !== curDate || newTime !== curTime) isReschedule = true;
  }

  // ── RESCHEDULE ───────────────────────────────────────────────────
  // ── RESCHEDULE ───────────────────────────────────────────────────
  if (isReschedule) {
    // Check xem đã Rescheduled chưa → tránh submit lặp
    const curStatus = String(sh.getRange(masterRow, H['Status']).getValue() || '').trim();
    if (/rescheduled/i.test(curStatus)) {
      Logger.log('Already rescheduled on Master, skip: uid=' + appt.id);
      return 'unchanged';
    }

    const oldUid = H['CalendlyEventUID']
      ? String(sh.getRange(masterRow, H['CalendlyEventUID']).getValue() || appt.id || '').trim()
      : String(appt.id || '').trim();
    const newUid = acuityStableRescheduleUid_(appt);
    if (!newUid) {
      Logger.log('Reschedule skipped: could not build stable UID for appt=' + appt.id);
      return 'unchanged';
    }

    // Đánh dấu dòng cũ
    sh.getRange(masterRow, H['Status']).setValue('Rescheduled');
    if (H['Active?'])          sh.getRange(masterRow, H['Active?']).setValue('No');
    if (H['CanceledAt'] && !sh.getRange(masterRow, H['CanceledAt']).getValue()) {
      sh.getRange(masterRow, H['CanceledAt']).setValue(new Date());
    }
    if (H['RescheduledToUID'] && !sh.getRange(masterRow, H['RescheduledToUID']).getValue()) {
      sh.getRange(masterRow, H['RescheduledToUID']).setValue(newUid);
    }
    if (H['Automation Notes']) {
      const prev = sh.getRange(masterRow, H['Automation Notes']).getValue() || '';
      const note = 'Rescheduled via Acuity to ' + newUid + ' @ ' + new Date().toISOString();
      sh.getRange(masterRow, H['Automation Notes']).setValue(prev ? prev + '\n' + note : note);
    }
    Logger.log('Marked Rescheduled: row=' + masterRow + ' uid=' + appt.id);

    try {
      if (typeof _rememberCancelUID_ === 'function') {
        _rememberCancelUID_(
          ACUITY_CFG.COMPANY,
          fieldMap['Visit Type'],
          acuityNormEmail_(fieldMap['Email']),
          acuityNormPhone_(fieldMap['Phone']),
          oldUid,
          7200
        );
      }
    } catch (e) {
      Logger.log('Acuity reschedule link cache skipped: ' + (e && e.message ? e.message : e));
    }

    // Submit dòng mới với UID mới
    const newFieldMap = Object.assign({}, fieldMap, { 'Admin: Calendly Event UID': newUid });
    acuitySubmitToForm_(formId, newFieldMap);


    Logger.log('✅ Reschedule submitted newUid=' + newUid);
    return 'rescheduled';
  }

  // ── EDIT INFO: update dòng cũ tại chỗ ───────────────────────────
  const changes = [];

  if (fieldMap['Phone'] && H['Phone']) {
    const newVal = fieldMap['Phone'];
    const curVal = String(sh.getRange(masterRow, H['Phone']).getValue() || '').trim();
    Logger.log('Phone: cur="' + curVal + '" new="' + newVal + '" match=' + (curVal === newVal));
    if (newVal && newVal !== curVal) {
      sh.getRange(masterRow, H['Phone']).setValue(newVal);
      changes.push('Phone: ' + curVal + ' → ' + newVal);
    }
  }

  if (H['Diamond Type'] && fieldMap['Diamond Type'] && fieldMap['Diamond Type'].length) {
    const newVal = fieldMap['Diamond Type'][0];
    const curVal = String(sh.getRange(masterRow, H['Diamond Type']).getValue() || '').trim();
    Logger.log('Diamond: cur="' + curVal + '" new="' + newVal + '" match=' + (curVal === newVal));
    if (newVal && newVal !== curVal) {
      sh.getRange(masterRow, H['Diamond Type']).setValue(newVal);
      changes.push('Diamond Type: ' + curVal + ' → ' + newVal);
    }
  }

  if (H['Budget Range'] && fieldMap['Budget Range'] && fieldMap['Budget Range'].length) {
    const newVal = fieldMap['Budget Range'][0];
    const curVal = String(sh.getRange(masterRow, H['Budget Range']).getValue() || '').trim();
    Logger.log('Budget: cur="' + curVal + '" new="' + newVal + '" match=' + (curVal === newVal));
    if (newVal && newVal !== curVal) {
      sh.getRange(masterRow, H['Budget Range']).setValue(newVal);
      changes.push('Budget Range: ' + curVal + ' → ' + newVal);
    }
  }

  if (H['Source'] && fieldMap['Source'] && fieldMap['Source'].length) {
    const newVal = fieldMap['Source'][0];
    const curVal = String(sh.getRange(masterRow, H['Source']).getValue() || '').trim();
    Logger.log('Source: cur="' + curVal + '" new="' + newVal + '" match=' + (curVal === newVal));
    if (newVal && newVal !== curVal && newVal !== 'Did not disclose') {
      sh.getRange(masterRow, H['Source']).setValue(newVal);
      changes.push('Source: ' + curVal + ' → ' + newVal);
    }
  }

  if (H['Style Notes'] && fieldMap['Style Notes']) {
    const newVal = fieldMap['Style Notes'];
    const curVal = String(sh.getRange(masterRow, H['Style Notes']).getValue() || '').trim();
    Logger.log('StyleNotes: cur="' + curVal + '" new="' + newVal + '" match=' + (curVal === newVal));
    if (newVal && newVal !== curVal) {
      sh.getRange(masterRow, H['Style Notes']).setValue(newVal);
      changes.push('Style Notes updated');
    }
  }
  
  if (changes.length) {
    if (H['Automation Notes']) {
      const prev = sh.getRange(masterRow, H['Automation Notes']).getValue() || '';
      const note = 'Edited via Acuity @ ' + new Date().toISOString() + '\n' + changes.join('\n');
      sh.getRange(masterRow, H['Automation Notes']).setValue(prev ? prev + '\n' + note : note);
    }
    Logger.log('✏️ Edited row=' + masterRow + ' uid=' + appt.id + ' | ' + changes.join(' | '));
    return 'edited';
  }

  return 'unchanged';
}

function debugMasterRow() {
  const SP     = PropertiesService.getScriptProperties();
  const userId = SP.getProperty('ACUITY_USER_ID');
  const apiKey = SP.getProperty('ACUITY_API_KEY');

  const appts        = acuityFetchAppointments_(userId, apiKey);
  const existingUIDs = acuityGetMasterUIDs_();

  const TARGET_UID = '1684912230'; // ← uid gốc

  const appt = appts.find(a => String(a.id) === TARGET_UID);
  if (!appt) { Logger.log('Not found'); return; }

  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getSheetByName('00_Master Appointments');
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const H   = {};
  hdr.forEach((h, i) => { if (h) H[String(h).trim()] = i + 1; });

  const lastRow = sh.getLastRow();
  const uids    = sh.getRange(2, H['CalendlyEventUID'], lastRow - 1, 1).getValues();

  Logger.log('Searching for uid=' + TARGET_UID);
  Logger.log('Total rows: ' + uids.length);

  // Log tất cả rows liên quan
  for (let i = 0; i < uids.length; i++) {
    const rowUid = String(uids[i][0]).trim();
    if (!rowUid.includes(TARGET_UID)) continue;
    const row    = i + 2;
    const status = H['Status'] ? String(sh.getRange(row, H['Status']).getValue() || '').trim() : 'N/A';
    Logger.log('row=' + row + ' uid="' + rowUid + '" status="' + status + '"');
  }

  // Log existingUIDs có chứa _R không
  Logger.log('existingUIDs has "' + TARGET_UID + '": ' + existingUIDs.has(TARGET_UID));
  const rUid = TARGET_UID + '_R1775706825904'; // thay bằng _R uid thật
  Logger.log('existingUIDs has rUid: ' + existingUIDs.has(rUid));
  
  // List tất cả UIDs trong existingUIDs có chứa TARGET_UID
  Logger.log('=== All matching UIDs in existingUIDs ===');
  existingUIDs.forEach(u => {
    if (u.includes(TARGET_UID)) Logger.log('  "' + u + '"');
  });
}


function debugCompareFields() {
  const SP     = PropertiesService.getScriptProperties();
  const userId = SP.getProperty('ACUITY_USER_ID');
  const apiKey = SP.getProperty('ACUITY_API_KEY');

  // Fetch trực tiếp appointment bằng ID — không bị cache
  const url  = 'https://acuityscheduling.com/api/v1/appointments/1684912230';
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(userId + ':' + apiKey) },
    muteHttpExceptions: true,
  });
  const appt = JSON.parse(resp.getContentText());

  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getSheetByName('00_Master Appointments');
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const H   = {};
  hdr.forEach((h, i) => { if (h) H[String(h).trim()] = i + 1; });

  const masterRow  = 528;
  const fieldMap   = acuityToFormFieldMap_(appt);

  const phone   = String(sh.getRange(masterRow, H['Phone']).getValue() || '').trim();
  const diamond = String(sh.getRange(masterRow, H['Diamond Type']).getValue() || '').trim();
  const budget  = String(sh.getRange(masterRow, H['Budget Range']).getValue() || '').trim();
  const source  = String(sh.getRange(masterRow, H['Source']).getValue() || '').trim();
  const notes   = String(sh.getRange(masterRow, H['Style Notes']).getValue() || '').trim();

  Logger.log('Phone:   Master="' + phone   + '" Acuity="' + fieldMap['Phone'] + '" match=' + (phone === fieldMap['Phone']));
  Logger.log('Diamond: Master="' + diamond + '" Acuity="' + (fieldMap['Diamond Type'][0]||'') + '" match=' + (diamond === (fieldMap['Diamond Type'][0]||'')));
  Logger.log('Budget:  Master="' + budget  + '" Acuity="' + (fieldMap['Budget Range'][0]||'') + '" match=' + (budget === (fieldMap['Budget Range'][0]||'')));
  Logger.log('Source:  Master="' + source  + '" Acuity="' + (fieldMap['Source'][0]||'') + '" match=' + (source === (fieldMap['Source'][0]||'')));
  Logger.log('Notes:   Master="' + notes   + '" Acuity="' + fieldMap['Style Notes'] + '" match=' + (notes === fieldMap['Style Notes']));
}

function debugAcuityRawData() {
  const SP     = PropertiesService.getScriptProperties();
  const userId = SP.getProperty('ACUITY_USER_ID');
  const apiKey = SP.getProperty('ACUITY_API_KEY');

  // Fetch trực tiếp 1 appointment bằng ID
  const url  = 'https://acuityscheduling.com/api/v1/appointments/1684912230';
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(userId + ':' + apiKey) },
    muteHttpExceptions: true,
  });

  const data = JSON.parse(resp.getContentText());
  Logger.log('firstName: ' + data.firstName);
  Logger.log('datetime: ' + data.datetime);
  
  // Log form values
  (data.forms || []).forEach(f => {
    (f.values || []).forEach(v => {
      if (v.value) Logger.log(v.name + ': ' + v.value);
    });
  });
}

function TEST_editInfo() {
  const SP     = PropertiesService.getScriptProperties();
  const userId = SP.getProperty('ACUITY_USER_ID');
  const apiKey = SP.getProperty('ACUITY_API_KEY');
  const formId = SP.getProperty('FORM_ID');

  // Fetch trực tiếp bằng ID
  const url  = 'https://acuityscheduling.com/api/v1/appointments/1684912230';
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(userId + ':' + apiKey) },
    muteHttpExceptions: true,
  });
  const appt = JSON.parse(resp.getContentText());

  Logger.log('Testing Edit Info for: ' + appt.firstName + ' uid=' + appt.id);
  const result = acuityHandleExisting_(appt, formId);
  Logger.log('Result: ' + result);
}

function findFormId() {
  // Cách 1: Lấy từ Script Properties
  const SP = PropertiesService.getScriptProperties();
  const allProps = SP.getProperties();
  Logger.log('=== Script Properties ===');
  Object.keys(allProps).forEach(k => Logger.log(`${k} = ${allProps[k]}`));
}

function clearAcuityDoneKeys() {
  const SP   = PropertiesService.getScriptProperties();
  const all  = SP.getProperties();
  const keys = Object.keys(all).filter(k => k.startsWith('ACUITY:'));
  
  Logger.log('Found ' + keys.length + ' ACUITY: keys');
  keys.forEach(k => Logger.log('  ' + k + ' = ' + all[k]));
  
  keys.forEach(k => SP.deleteProperty(k));
  Logger.log('Deleted ' + keys.length + ' keys');
}
