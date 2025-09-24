/***** Phase E — Schedule & Coverage helpers (10_Roster_Schedule only) *****/

/** Read 10_Roster_Schedule and return today's on-duty set + coverage partner map */
function getScheduleToday_() {
  const ss = SpreadsheetApp.getActive();
  const sh = getSheetOrThrow_(SHEET_10);
  const tz = (typeof TIMEZONE !== 'undefined' && TIMEZONE) ? TIMEZONE : ss.getSpreadsheetTimeZone();

  const dow = Utilities.formatDate(new Date(), tz, 'EEE'); // Mon/Tue/...
  const data = sh.getDataRange().getValues();
  if (!data.length) return { onDuty: new Set(), coveragePartner: new Map(), enabled: new Set() };

  const headers = data[0].map(h => String(h || '').trim());
  const col = {
    rep: headers.indexOf('Rep'),
    Mon: headers.indexOf('Mon'),
    Tue: headers.indexOf('Tue'),
    Wed: headers.indexOf('Wed'),
    Thu: headers.indexOf('Thu'),
    Fri: headers.indexOf('Fri'),
    Sat: headers.indexOf('Sat'),
    Sun: headers.indexOf('Sun'),
    covEnabled: headers.indexOf('Assisted Coverage Enabled?'),
    covPartner: headers.indexOf('Assisted Coverage Partner')
  };

  const dayColByEEE = {Mon: col.Mon, Tue: col.Tue, Wed: col.Wed, Thu: col.Thu, Fri: col.Fri, Sat: col.Sat, Sun: col.Sun};
  const dayCol = dayColByEEE[dow] ?? col.Mon; // fallback to Mon if something odd

  const norm = buildRosterNormalizer_ ? buildRosterNormalizer_() : (s => String(s || '').trim());

  const onDuty = new Set();
  const coveragePartner = new Map();  // rep -> partner
  const enabled = new Set();          // reps with Assisted Coverage Enabled? = Y

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const repRaw = row[col.rep];
    const rep = norm(repRaw);
    if (!rep) continue;

    const worksToday = isTruthyY_(row[dayCol]);
    if (worksToday) onDuty.add(rep);

    const en = isTruthyY_(row[col.covEnabled]);
    const partner = norm(row[col.covPartner]);
    if (en && partner && partner !== rep) {
      enabled.add(rep);
      coveragePartner.set(rep, partner);
    }
  }
  return { onDuty, coveragePartner, enabled, tz };
}

/** Compute on-duty expected sets with Maria↔Paul assisted-coverage applied.
 * Inputs:
 *  - inScope: Set of RootApptID currently in scope (from 07)
 *  - map08: rows from 08_Reps_Map
 * Returns:
 *  - expectedByRootDuty: Map<RootApptID, Set<Rep>>
 *  - expectedByRepDuty:  Map<Rep, Set<RootApptID>>
 *  - roleByRootRepDuty:  Map<`${root}||${rep}`, 'Assigned'|'Assisted'>
 *  - assignedGaps: Map<RootApptID, {assigned:Array<string>}>
 *  - assistedGaps: Map<RootApptID, {pair:string}>   // e.g., 'Maria & Paul'
 */
function computeExpectedSetsWithSchedule_(inScope, map08) {
  const { onDuty, coveragePartner, enabled, tz } = getScheduleToday_();
  const expectedAllByRoot = new Map(); // root -> Map<rep,role>
  const assignedListByRoot = new Map(); // root -> Set(assigned reps)
  const assistedListByRoot = new Map(); // root -> Set(assisted reps)

  // Build expected-all (pre-schedule) from 08 where Include? = Y and root in-scope
  map08.forEach(r => {
    const root = String(r['RootApptID'] || '').trim();
    if (!root || !inScope.has(root)) return;

    const include = String(r['Include? (Y/N)'] || r['Include?'] || '').trim().toUpperCase();
    if (include !== 'Y') return;

    const rep = String(r['Rep'] || '').trim();
    const role = String(r['Role (Assigned/Assisted)'] || '').trim() || 'Assigned';

    if (!expectedAllByRoot.has(root)) expectedAllByRoot.set(root, new Map());
    if (!assignedListByRoot.has(root)) assignedListByRoot.set(root, new Set());
    if (!assistedListByRoot.has(root)) assistedListByRoot.set(root, new Set());

    const cur = expectedAllByRoot.get(root).get(rep);
    // Prefer Assigned if there is any conflict
    const roleToSet = (cur === 'Assigned' || role === 'Assigned') ? 'Assigned' : 'Assisted';
    expectedAllByRoot.get(root).set(rep, roleToSet);

    if (roleToSet === 'Assigned') assignedListByRoot.get(root).add(rep);
    else assistedListByRoot.get(root).add(rep);
  });

  const expectedByRootDuty = new Map();
  const roleByRootRepDuty = new Map();
  const assignedGaps = new Map();
  const assistedGaps = new Map();

  // For each root, filter to on-duty; apply Maria↔Paul assisted-coverage
  expectedAllByRoot.forEach((repRoleMap, root) => {
    const dutySet = new Set();

    // 1) Assigned on-duty
    const assignedAll = assignedListByRoot.get(root) || new Set();
    let assignedOnDutyCount = 0;
    assignedAll.forEach(rep => {
      if (onDuty.has(rep)) {
        dutySet.add(rep);
        roleByRootRepDuty.set(`${root}||${rep}`, 'Assigned');
        assignedOnDutyCount++;
      }
    });

    // If there are assigned reps but none are on-duty → Assigned Coverage Gap
    if (assignedAll.size > 0 && assignedOnDutyCount === 0) {
      assignedGaps.set(root, { assigned: [...assignedAll].sort() });
    }

    // 2) Assisted on-duty
    const assistedAll = assistedListByRoot.get(root) || new Set();
    const assistedOnDuty = new Set();
    assistedAll.forEach(rep => {
      if (onDuty.has(rep)) {
        dutySet.add(rep);
        roleByRootRepDuty.set(`${root}||${rep}`, 'Assisted');
        assistedOnDuty.add(rep);
      }
    });

    // 3) Assisted coverage (only if coverage enabled for that rep)
    // If Paul is assisted on this root and OFF → route to Maria if Maria ON (and vice versa)
    assistedAll.forEach(rep => {
      if (onDuty.has(rep)) return; // already covered
      if (!enabled.has(rep)) return; // coverage not enabled for this assisted rep

      const partner = coveragePartner.get(rep); // e.g., Paul->Maria
      if (!partner) return;

      const bothOff = !onDuty.has(rep) && !onDuty.has(partner);
      if (bothOff) {
        // Assisted Coverage Gap (both Maria & Paul off)
        assistedGaps.set(root, { pair: `${rep} & ${partner}` });
        return;
      }
      if (onDuty.has(partner)) {
        // Route to partner for today
        dutySet.add(partner);
        // If partner is also an assigned rep on this root, keep Assigned role; else Assisted
        const originalPartnerRole = repRoleMap.get(partner);
        const role = (originalPartnerRole === 'Assigned') ? 'Assigned' : 'Assisted';
        roleByRootRepDuty.set(`${root}||${partner}`, role);
      }
    });

    if (dutySet.size > 0) {
      expectedByRootDuty.set(root, dutySet);
    }
  });

  // Invert to rep -> set of roots
  const expectedByRepDuty = new Map();
  expectedByRootDuty.forEach((repSet, root) => {
    repSet.forEach(rep => {
      if (!expectedByRepDuty.has(rep)) expectedByRepDuty.set(rep, new Set());
      expectedByRepDuty.get(rep).add(root);
    });
  });

  return { expectedByRootDuty, expectedByRepDuty, roleByRootRepDuty, assignedGaps, assistedGaps, tz };
}

/** Helper: treat Y/Yes/True/1 as true */
function isTruthyY_(v) {
  const s = String(v || '').trim().toLowerCase();
  return s === 'y' || s === 'yes' || s === 'true' || s === '1';
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




