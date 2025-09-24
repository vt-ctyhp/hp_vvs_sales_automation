/**************************************************************
 * One-time Fixer — Recompute "Visit #" safely (v1)
 * 
 * What it does:
 * - Reads 00_Master Appointments
 * - Recomputes "Visit #" per contact in chronological order
 *   using ONLY rows that represent a real visit:
 *     • Completed  → counts
 *     • Scheduled AND Active? != "No"  → counts
 *   (Canceled, Rescheduled, No-Show, and Active? = "No" DON'T count)
 * - Writes ONLY the "Visit #" column, and only if the value changed
 * - Creates a backup tab with diffs so you can undo
 *
 * How to run:
 * 1) Paste this file into your Apps Script project.
 * 2) Select fixVisitNumbersOnce() and Run.
 * 3) Check the "Backup – Visit # (YYYY-MM-DD HH:mm)" tab for a diff log.
 *
 * To rollback:
 * - Run undoVisitNumbersFromLatestBackup()
 *
 **************************************************************/

function fixVisitNumbersOnce() {
  const SP = PropertiesService.getScriptProperties();
  const MASTER_SSID = (SP.getProperty('SPREADSHEET_ID') || '').trim();
  const ss = MASTER_SSID ? SpreadsheetApp.openById(MASTER_SSID)
                         : SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet: 00_Master Appointments');

  const H = headerMap_(sh);

  // Required columns (tolerant, but we must at least find Visit # and a way to sort + identify contact)
  const colVisitNo = idx_(H, ['Visit #','Visit#','Visit Number']);
  if (!colVisitNo) throw new Error('Missing "Visit #" column.');

  const colStatus  = idx_(H, ['Status']);
  const colActive  = idx_(H, ['Active?','Active ?']);
  const colEmail   = idx_(H, ['EmailLower','Email Lower','Email lower']);
  const colPhone   = idx_(H, ['PhoneNorm','Phone Norm','Phone normalized','Phone (norm)']);
  const colIso     = idx_(H, ['ApptDateTime (ISO)','ApptDateTime(ISO)','ApptDateTime ISO']);
  const colVDate   = idx_(H, ['Visit Date','VisitDate']);
  const colVTime   = idx_(H, ['Visit Time','VisitTime']);
  const colName    = idx_(H, ['Customer Name','Client Name','Customer']);
  const colRootId  = idx_(H, ['RootApptID','APPT_ID','Appt ID','RootApptId']);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return SpreadsheetApp.getUi().alert('No data rows to process.');

  // Read once for speed
  const numRows = lastRow - 1;
  const rowVals = sh.getRange(2, 1, numRows, sh.getLastColumn()).getValues();

  // Build per-row objects (only what's needed)
  const rows = [];
  for (let i = 0; i < numRows; i++) {
    const r = rowVals[i];

    // Identify contact (prefer email, else phone). If neither exists, we SKIP changing Visit # for safety.
    const email   = colEmail ? String(r[colEmail-1] || '').toLowerCase().trim() : '';
    const phone   = colPhone ? String(r[colPhone-1] || '').trim()               : '';
    const contactKey = email ? `e:${email}` : (phone ? `p:${phone}` : '');

    const status  = colStatus ? String(r[colStatus-1] || '').trim() : '';
    const activeV = colActive ? String(r[colActive-1] || '').trim() : '';
    const vnoOld  = String(r[colVisitNo-1] ?? '').trim();
    const name    = colName ? String(r[colName-1] || '').trim() : '';
    const root    = colRootId ? String(r[colRootId-1] || '').trim() : '';

    // Build sortable datetime (ISO preferred; else Visit Date + Visit Time; else fallback by row order)
    const ts = composeDateTime_(
      colIso ? r[colIso-1] : null,
      colVDate ? r[colVDate-1] : null,
      colVTime ? r[colVTime-1] : null
    );

    rows.push({
      sheetRow: i + 2,             // 1-based row index in sheet
      contactKey,                   // "e:<email>" or "p:<phone>" or ""
      status,
      activeV,                      // "Yes"/"No"/""
      vnoOld,
      ts: ts || new Date(2000,0,1,0,0,0,0).getTime() + i, // stable fallback order
      name, root
    });
  }

  // Group by contactKey; rows with no contactKey are left untouched
  const groups = new Map();
  for (const r of rows) {
    if (!r.contactKey) continue;      // skip unidentifiable rows for safety
    if (!groups.has(r.contactKey)) groups.set(r.contactKey, []);
    groups.get(r.contactKey).push(r);
  }

  // Compute new Visit # per group (chronological)
  const newByRow = new Map();   // sheetRow -> newValue (string or '')
  for (const [key, list] of groups.entries()) {
    // sort by time (earliest first)
    list.sort((a,b) => a.ts - b.ts);

    let count = 0;
    for (const r of list) {
      const { countable, shouldHaveNumber } = classifyCountable_(r.status, r.activeV);
      if (shouldHaveNumber) {
        // number this row as (count of prior countable) + 1
        count += 1;
        newByRow.set(r.sheetRow, String(count));
      } else {
        // not countable => blank the Visit #
        newByRow.set(r.sheetRow, '');
      }
    }
  }

  // Prepare backup of changed rows only
  const changes = [];
  for (const r of rows) {
    if (!newByRow.has(r.sheetRow)) continue; // not in a contact group -> skip for safety
    const newV = newByRow.get(r.sheetRow);
    if (String(newV) !== String(r.vnoOld)) {
      changes.push({
        sheetRow: r.sheetRow,
        root: r.root,
        name: r.name,
        oldV: r.vnoOld,
        newV: newV,
        status: r.status,
        activeV: r.activeV
      });
    }
  }

  // If nothing to change, exit quietly
  if (!changes.length) {
    SpreadsheetApp.getUi().alert('Visit # is already consistent. No changes made.');
    return;
  }

  // Backup first (timestamped tab with diffs)
  const backupName = createBackupSheet_(ss, sh, changes);

  // Apply updates — only to Visit # column (row-by-row to avoid touching other cells)
  // (One-time run: row-by-row is fine; still reasonably fast for typical sizes.)
  for (const c of changes) {
    sh.getRange(c.sheetRow, colVisitNo).setValue(c.newV);
  }

  SpreadsheetApp.getUi().alert(
    'Visit # fix complete.\n' +
    `Changed rows: ${changes.length}\n` +
    `Backup tab: ${backupName}`
  );
}


/** UNDO: restore Visit # from the most recent backup tab created by this script. */
function undoVisitNumbersFromLatestBackup() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet: 00_Master Appointments');

  // Find the latest backup tab by prefix
  const tabs = ss.getSheets().map(s => s.getName());
  const backups = tabs
    .filter(n => /^Backup\s–\sVisit\s#\s\(/.test(n)) // "Backup – Visit # (YYYY-MM-DD HH:mm)"
    .sort(); // lexicographic sort puts latest last due to timestamp in name

  if (!backups.length) {
    SpreadsheetApp.getUi().alert('No backup tabs found.');
    return;
  }
  const backupName = backups[backups.length - 1];
  const bsh = ss.getSheetByName(backupName);

  const BH = headerMap_(bsh);
  const colRow  = idx_(BH, ['Sheet Row']);
  const colOld  = idx_(BH, ['Old Visit #']);
  if (!colRow || !colOld) throw new Error('Backup sheet missing required columns.');

  const last = bsh.getLastRow();
  if (last < 2) {
    SpreadsheetApp.getUi().alert('Backup tab is empty.');
    return;
  }

  const data = bsh.getRange(2, 1, last-1, bsh.getLastColumn()).getValues();
  const MH   = headerMap_(sh);
  const colVisitNo = idx_(MH, ['Visit #','Visit#','Visit Number']);
  if (!colVisitNo) throw new Error('Master missing "Visit #" column.');

  let restored = 0;
  data.forEach(r => {
    const rowIdx = Number(r[colRow-1]);
    const oldVal = r[colOld-1];
    if (rowIdx && (rowIdx >= 2) && rowIdx <= sh.getLastRow()) {
      sh.getRange(rowIdx, colVisitNo).setValue(oldVal);
      restored++;
    }
  });

  SpreadsheetApp.getUi().alert(`Restored Visit # from backup "${backupName}". Rows restored: ${restored}.`);
}


/* ------------------------ Helpers ------------------------ */

function headerMap_(sheet) {
  const hdr = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  const H = {};
  hdr.forEach((h,i) => { const k = String(h||'').trim(); if (k) H[k] = i+1; });
  return H;
}

function idx_(H, names) {
  for (const n of names) if (H[n]) return H[n];
  // tolerant fallback: normalized compare (lowercase, remove spaces and punctuation)
  const norm = o => String(o||'').toLowerCase().replace(/[^a-z0-9]/g,'');
  const want = names.map(norm);
  for (const k of Object.keys(H)) if (want.includes(norm(k))) return H[k];
  return 0;
}

function classifyCountable_(statusRaw, activeRaw) {
  const s = String(statusRaw || '').toLowerCase();
  const a = String(activeRaw || '').toLowerCase();

  const isCompleted   = /completed/.test(s);
  const isCanceled    = /cancell?ed/.test(s);
  const isRescheduled = /rescheduled/.test(s);
  const isNoShow      = /no[-\s]?show/.test(s);
  const isScheduled   = /scheduled/.test(s) && !isRescheduled; // exclude explicit "Rescheduled"

  // Active? column (if present) is authoritative: "No" means do not count.
  // If Active? is blank/missing, we fall back to Status.
  const activeYes = (a === 'yes') || (a === 'y') || (a === 'true');
  const activeNo  = (a === 'no')  || (a === 'n') || (a === 'false');

  const countable =
    isCompleted ||
    (isScheduled && !activeNo); // scheduled and not explicitly inactivated

  const shouldHaveNumber =
    countable && !isCanceled && !isRescheduled && !isNoShow;

  return { countable, shouldHaveNumber };
}

function composeDateTime_(isoVal, dateVal, timeVal) {
  // Prefer ISO value if valid
  if (isoVal) {
    const d = (isoVal instanceof Date) ? isoVal : new Date(isoVal);
    if (!isNaN(d)) return d.getTime();
  }
  // Compose from date + time if possible
  let d0 = null, t0 = null;
  if (dateVal) d0 = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
  if (timeVal) t0 = (timeVal instanceof Date) ? timeVal : new Date(timeVal);
  if (d0 && !isNaN(d0)) {
    const y = d0.getFullYear(), m = d0.getMonth(), d = d0.getDate();
    let hh = 0, mm = 0, ss = 0;
    if (t0 && !isNaN(t0)) { hh = t0.getHours(); mm = t0.getMinutes(); ss = t0.getSeconds(); }
    const out = new Date(y, m, d, hh, mm, ss, 0);
    if (!isNaN(out)) return out.getTime();
  }
  return null;
}

function createBackupSheet_(ss, masterSheet, changes) {
  const name = 'Backup – Visit # (' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") + ')';
  const sh = ss.insertSheet(name);
  const hdr = ['Sheet Row','RootApptID','Customer Name','Old Visit #','New Visit #','Status','Active?'];
  sh.getRange(1,1,1,hdr.length).setValues([hdr]);
  sh.setFrozenRows(1);

  const rows = changes.map(c => [c.sheetRow, c.root, c.name, c.oldV, c.newV, c.status, c.activeV]);
  if (rows.length) sh.getRange(2,1,rows.length,hdr.length).setValues(rows);

  // Make it visually read‑only-ish: protect editable range (optional; user can remove)
  try {
    const p = sh.protect().setDescription('Backup (auto-created by fixer)');
    const me = Session.getEffectiveUser();
    p.addEditor(me).removeEditors(p.getEditors().filter(e => e.getEmail() !== me.getEmail()));
    p.setWarningOnly(true); // allow edits but show a warning
  } catch(_){}

  return name;
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
