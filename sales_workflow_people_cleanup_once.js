/**
 * TEMPORARY one-time people cleanup.
 *
 * Intended lifecycle:
 * 1. Run sw_oncePreviewPeopleCleanup(authToken).
 * 2. Review _SW_PeopleCleanupPlan.
 * 3. Run sw_onceApplyPeopleCleanup(authToken).
 * 4. Verify workflow/audit output.
 * 5. Remove this file/functions from source control.
 */

var SW_ONCE_PEOPLE_CLEANUP_PLAN_SHEET_ = '_SW_PeopleCleanupPlan';
var SW_ONCE_PEOPLE_CLEANUP_BACKUP_SHEET_ = '_SW_PeopleCleanupBackup';
var SW_ONCE_RETIRED_REP_QUALIFICATIONS_BACKUP_SHEET_ = '_SW_Retired_RepQualifications';
var SW_ONCE_REP_QUALIFICATIONS_SHEET_ = 'Rep Qualifications';
var SW_ONCE_PEOPLE_CLEANUP_CUTOFF_ = '2026-04-01';

function sw_oncePreviewPeopleCleanup(authToken) {
  return swTimed_('sw_oncePreviewPeopleCleanup', function () {
    var ss = swSpreadsheet_();
    var user = swAuthUserForApi_(ss, authToken);
    if (!user.isAdmin) throw new Error('Admin access required.');
    swEnsureEmployeeScheduleSheets_(ss);
    var plan = swOnceBuildPeopleCleanupPlan_(ss, false);
    swOnceWritePeopleCleanupPlan_(ss, plan);
    return {
      ok: true,
      dryRun: true,
      planSheet: SW_ONCE_PEOPLE_CLEANUP_PLAN_SHEET_,
      generatedAt: plan.generatedAt,
      summary: plan.summary,
      generatedPasswords: []
    };
  });
}

function sw_oncePreviewPeopleCleanupRun() {
  var result = sw_oncePreviewPeopleCleanup('');
  Logger.log(swStringify_(result, 2));
  return result;
}

function sw_onceApplyPeopleCleanup(authToken) {
  return swTimed_('sw_onceApplyPeopleCleanup', function () {
    var ss = swSpreadsheet_();
    var user = swAuthUserForApi_(ss, authToken);
    if (!user.isAdmin) throw new Error('Admin access required.');
    swEnsureEmployeeScheduleSheets_(ss);
    var plan = swOnceBuildPeopleCleanupPlan_(ss, false);
    swOnceWritePeopleCleanupPlan_(ss, plan);
    var result = {
      ok: true,
      appliedAt: swIso_(new Date()),
      planSheet: SW_ONCE_PEOPLE_CLEANUP_PLAN_SHEET_,
      backupSheet: SW_ONCE_PEOPLE_CLEANUP_BACKUP_SHEET_,
      retiredRepQualificationsBackupSheet: SW_ONCE_RETIRED_REP_QUALIFICATIONS_BACKUP_SHEET_,
      generatedPasswords: [],
      actions: [],
      warnings: []
    };
    swOnceApplyUsers_(ss, user, result);
    swOnceRewriteRoster_(ss, user, result);
    swOnceStandardizeDropdownIdentity_(ss, result);
    swOnceCleanRecentAppointmentOwners_(ss, result);
    swOnceBackupAndDeleteRepQualifications_(ss, result);
    try { if (typeof swClearAssignmentOptionsMemoryCache_ === 'function') swClearAssignmentOptionsMemoryCache_(ss); } catch (_) {}
    try { CacheService.getScriptCache().remove('sw:assignmentOptions:v1:' + ss.getId()); } catch (_) {}
    try { result.generation = sw_generateSalesWorkflowTasks(); } catch (err) { result.generation = { ok: false, error: swTrim_(err && err.message || err) }; }
    result.audit = sw_adminAuditWorkflowPeopleData(authToken);
    return result;
  });
}

function sw_onceApplyPeopleCleanupRun() {
  var result = sw_onceApplyPeopleCleanup('');
  Logger.log(swStringify_(result, 2));
  return result;
}

function sw_onceAuditWorkflowPeopleDataRun() {
  var result = sw_adminAuditWorkflowPeopleData('');
  Logger.log(swStringify_(result, 2));
  return result;
}

function sw_onceClearDropdownIdentityAfterCleanup(authToken) {
  return swTimed_('sw_onceClearDropdownIdentityAfterCleanup', function () {
    var ss = swSpreadsheet_();
    var user = swAuthUserForApi_(ss, authToken);
    if (!user.isAdmin) throw new Error('Admin access required.');
    var result = { ok: true, cleared: 0, actions: [], warnings: [] };
    var sh = ss.getSheetByName(SW_SHEETS.DROPDOWN);
    if (!sh || sh.getLastRow() < 2) return result;
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
    var identityCols = swDropdownIdentityColumnIndexes_(headers);
    if (!identityCols.length) return result;
    var values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();
    identityCols.forEach(function (col) {
      for (var r = 0; r < values.length; r++) {
        var oldValue = values[r][col];
        if (!swTrim_(oldValue)) continue;
        swOnceBackupCell_(ss, SW_SHEETS.DROPDOWN, r + 2, col + 1, oldValue, '', 'clear legacy Dropdown identity');
        sh.getRange(r + 2, col + 1).clearContent();
        result.cleared++;
      }
    });
    result.actions.push('Cleared ' + result.cleared + ' legacy Dropdown identity cell(s).');
    return result;
  });
}

function sw_onceClearDropdownIdentityAfterCleanupRun() {
  var result = sw_onceClearDropdownIdentityAfterCleanup('');
  Logger.log(swStringify_(result, 2));
  return result;
}

function sw_onceDeleteRetiredPeopleCleanupArtifacts(authToken) {
  return swTimed_('sw_onceDeleteRetiredPeopleCleanupArtifacts', function () {
    var ss = swSpreadsheet_();
    var user = swAuthUserForApi_(ss, authToken);
    if (!user.isAdmin) throw new Error('Admin access required.');
    var result = { ok: true, deletedSheets: [], skippedSheets: [] };
    [
      SW_ONCE_PEOPLE_CLEANUP_PLAN_SHEET_,
      SW_ONCE_PEOPLE_CLEANUP_BACKUP_SHEET_,
      SW_ONCE_RETIRED_REP_QUALIFICATIONS_BACKUP_SHEET_
    ].forEach(function (name) {
      var sh = ss.getSheetByName(name);
      if (!sh) {
        result.skippedSheets.push(name);
        return;
      }
      if (ss.getSheets().length <= 1) throw new Error('Cannot delete the last remaining sheet.');
      ss.deleteSheet(sh);
      result.deletedSheets.push(name);
    });
    return result;
  });
}

function sw_onceDeleteRetiredPeopleCleanupArtifactsRun() {
  var result = sw_onceDeleteRetiredPeopleCleanupArtifacts('');
  Logger.log(swStringify_(result, 2));
  return result;
}

function swOnceTruthPeople_() {
  return [
    { name: 'Val', email: 'val@ctyhp.com', role: SW_OWNER_ROLES.SALES_REP, active: true, lab: true, natural: 'None', general: true, defaultJoc: 'Paul' },
    { name: 'Lyn', email: 'lyn@ctyhp.com', role: SW_OWNER_ROLES.SALES_REP, active: true, lab: true, natural: 'None', general: true, defaultJoc: 'Mark' },
    { name: 'Wendy', email: 'phungminh@ctyhp.com', role: SW_OWNER_ROLES.SALES_REP, active: true, lab: true, natural: 'Backup', general: true, defaultJoc: 'Mark' },
    { name: 'Kris', email: 'tuongvan@ctyhp.com', role: SW_OWNER_ROLES.SALES_REP, active: true, lab: true, natural: 'Backup', general: true, defaultJoc: 'Mark' },
    { name: 'An Vo', email: 'hoaan@ctyhp.com', role: SW_OWNER_ROLES.SALES_REP, active: true, lab: true, natural: 'Primary', general: true, defaultJoc: 'Mark' },
    { name: 'Paul', email: 'os003@ctyhp.com', role: SW_OWNER_ROLES.JOC, active: true, lab: false, natural: 'None', general: false, defaultJoc: '' },
    { name: 'Mark', email: 'oc002@ctyhp.com', role: SW_OWNER_ROLES.JOC, active: true, lab: false, natural: 'None', general: false, defaultJoc: '' },
    { name: 'Maria', email: 'maria@ctyhp.com', role: SW_OWNER_ROLES.JOC, active: false, lab: false, natural: 'None', general: false, defaultJoc: '' }
  ];
}

function swOnceBuildPeopleCleanupPlan_(ss) {
  var plan = {
    generatedAt: swIso_(new Date()),
    rows: [],
    summary: {
      usersToCreate: 0,
      usersToUpdate: 0,
      rosterRewriteRows: swOnceTruthPeople_().length,
      appointmentCellsToUpdate: 0,
      appointmentRowsSkippedMultiOwner: 0,
      dropdownIdentityCellsToStandardize: 0,
      repQualificationsWillRetire: !!ss.getSheetByName(SW_ONCE_REP_QUALIFICATIONS_SHEET_)
    }
  };
  swOncePlanUsers_(ss, plan);
  swOncePlanDropdownIdentity_(ss, plan);
  swOncePlanRecentAppointmentOwners_(ss, plan);
  plan.rows.push(swOncePlanRow_('ROSTER_REWRITE', '10_Roster_Schedule', '', '', '', '', 'Rewrite roster to canonical active staff, Maria inactive, skills in roster', 'READY'));
  if (plan.summary.repQualificationsWillRetire) {
    plan.rows.push(swOncePlanRow_('RETIRE_TAB', SW_ONCE_REP_QUALIFICATIONS_SHEET_, '', '', '', '', 'Back up then delete retired Rep Qualifications tab', 'READY'));
  }
  return plan;
}

function swOncePlanUsers_(ss, plan) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.USERS, SW_AUTH_USER_HEADERS);
  var rows = swReadSheetObjectsExpectedHeaders_(sh, SW_AUTH_USER_HEADERS);
  var byEmail = {};
  rows.forEach(function (row) { byEmail[swNormEmail_(row['Email'])] = row; });
  swOnceTruthPeople_().forEach(function (person) {
    var existing = byEmail[person.email];
    var action = existing ? 'UPDATE_USER' : 'CREATE_USER';
    if (existing) plan.summary.usersToUpdate++;
    else plan.summary.usersToCreate++;
    plan.rows.push(swOncePlanRow_(action, SW_SHEETS.USERS, existing ? existing.__rowNumber : '', 'Email/Name/Roles/Active', existing ? existing['Name'] + ' | ' + existing['Roles'] + ' | ' + existing['Active?'] : '', person.name + ' | ' + person.role + ' | ' + (person.active ? 'Y' : 'N'), existing ? 'Preserve existing password when present' : (person.active ? 'Generate temporary password' : 'Inactive historical user, no password'), 'READY'));
  });
}

function swOncePlanDropdownIdentity_(ss, plan) {
  var sh = ss.getSheetByName(SW_SHEETS.DROPDOWN);
  if (!sh || sh.getLastRow() < 2) return;
  var values = sh.getDataRange().getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  [
    swPickIndex_(H, ['Assigned Rep', 'Client Advisor']),
    swPickIndex_(H, ['Assisted Rep', 'Assistant Rep', 'JOC'])
  ].forEach(function (col) {
    if (col < 0) return;
    for (var r = 1; r < values.length; r++) {
      var oldValue = values[r][col];
      var next = swOnceCanonicalIdentityList_(oldValue);
      if (next !== swTrim_(oldValue)) {
        plan.summary.dropdownIdentityCellsToStandardize++;
        plan.rows.push(swOncePlanRow_('STANDARDIZE_DROPDOWN_IDENTITY', SW_SHEETS.DROPDOWN, r + 1, headers[col], oldValue, next, 'Canonical workflow identity field only', 'READY'));
      }
    }
  });
}

function swOncePlanRecentAppointmentOwners_(ss, plan) {
  var sh = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!sh || sh.getLastRow() < 2) return;
  var headers = swEnsureMasterOwnerHeaders_(sh);
  var allHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(allHeaders);
  var dateCol = swPickIndex_(H, ['Visit Date', 'Appointment Date', 'Date']) + 1;
  if (dateCol <= 0) return;
  var display = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var truth = swOnceTruthIndex_();
  for (var i = 0; i < display.length; i++) {
    var dateKey = swOnceDateKey_(values[i][dateCol - 1] || display[i][dateCol - 1]);
    if (!dateKey || dateKey < SW_ONCE_PEOPLE_CLEANUP_CUTOFF_) continue;
    var rowNumber = i + 2;
    var assigned = swTrim_(display[i][headers.assignedRep - 1]);
    var assisted = swTrim_(display[i][headers.assistedRep - 1]);
    var ownerPlan = swOnceAppointmentOwnerPlan_(truth, assigned, assisted);
    ownerPlan.actions.forEach(function (action) {
      plan.summary.appointmentCellsToUpdate++;
      plan.rows.push(swOncePlanRow_('CLEAN_APPOINTMENT_OWNER', SW_SHEETS.MASTER, rowNumber, action.column, action.oldValue, action.newValue, action.reason, 'READY'));
    });
    if (ownerPlan.skips.length) plan.summary.appointmentRowsSkippedMultiOwner++;
    ownerPlan.skips.forEach(function (skip) {
      plan.rows.push(swOncePlanRow_('SKIP_APPOINTMENT_OWNER', SW_SHEETS.MASTER, rowNumber, skip.column, skip.value, '', skip.reason, 'SKIPPED'));
    });
  }
}

function swOnceApplyUsers_(ss, actor, result) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.USERS, SW_AUTH_USER_HEADERS);
  var rows = swReadSheetObjectsExpectedHeaders_(sh, SW_AUTH_USER_HEADERS);
  swOnceAssertNoBlockingUserConflicts_(rows);
  var byEmail = {};
  rows.forEach(function (row) { byEmail[swNormEmail_(row['Email'])] = row; });
  swOnceTruthPeople_().forEach(function (person) {
    var existing = byEmail[person.email] || null;
    var rowNumber = existing && existing.__rowNumber ? existing.__rowNumber : sh.getLastRow() + 1;
    var password = '';
    var salt = existing ? existing['Password Salt'] : '';
    var hash = existing ? existing['Password Hash'] : '';
    var temporary = existing ? existing['Temporary Password?'] : '';
    if (person.active && (!salt || !hash)) {
      password = swAuthGeneratedPassword_();
      salt = swAuthNewSalt_();
      hash = swAuthHash_(password, salt);
      temporary = 'Y';
      result.generatedPasswords.push({ name: person.name, email: person.email, password: password });
    }
    var next = {
      'Email': person.email,
      'Name': person.name,
      'Roles': person.role,
      'Active?': person.active ? 'Y' : 'N',
      'Password Salt': person.active ? salt : '',
      'Password Hash': person.active ? hash : '',
      'Temporary Password?': person.active ? (temporary || 'N') : 'N',
      'Last Login At': existing ? existing['Last Login At'] : '',
      'Notes': existing && existing['Notes'] ? existing['Notes'] : 'Created/standardized by one-time people cleanup.'
    };
    var oldLabel = existing ? existing['Name'] + ' | ' + existing['Roles'] + ' | ' + existing['Active?'] : '';
    var newLabel = next['Name'] + ' | ' + next['Roles'] + ' | ' + next['Active?'];
    if (oldLabel !== newLabel) swOnceBackupCell_(ss, SW_SHEETS.USERS, rowNumber, 1, oldLabel, newLabel, 'upsert canonical workflow user');
    sh.getRange(rowNumber, 1, 1, SW_AUTH_USER_HEADERS.length).setValues([SW_AUTH_USER_HEADERS.map(function (h) { return next[h] == null ? '' : next[h]; })]);
    swAuthClearUserCaches_(ss, person.email);
    result.actions.push((existing ? 'Updated ' : 'Created ') + person.name + ' <' + person.email + '>');
  });
}

function swOnceAssertNoBlockingUserConflicts_(rows) {
  var truth = swOnceTruthIndex_();
  (rows || []).forEach(function (row) {
    var email = swNormEmail_(row['Email']);
    var name = swOnceCanonicalIdentity_(row['Name']);
    if (!email || !name || !swWorkflowUserActive_(row)) return;
    var truthByName = truth.byName[swNorm_(name)];
    if (truthByName && truthByName.email !== email) {
      throw new Error('Active workflow user conflict for "' + name + '": existing email ' + email + ' conflicts with cleanup email ' + truthByName.email + '.');
    }
  });
}

function swOnceRewriteRoster_(ss, actor, result) {
  var sh = swEnsureEmployeeScheduleSheets_(ss).roster;
  var oldValues = sh.getDataRange().getDisplayValues();
  swOnceBackupBlock_(ss, SW_SHEETS.ROSTER, oldValues, 'pre-cleanup roster backup');
  var rosterRows = swReadEmployeeRosterRows_(ss);
  var oldByName = {};
  var oldByEmail = {};
  rosterRows.forEach(function (row) {
    var canonical = swOnceCanonicalIdentity_(row.name);
    if (canonical) oldByName[swNorm_(canonical)] = row;
    if (row.email) oldByEmail[row.email] = row;
  });
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
  var now = swIso_(new Date());
  var actorLabel = actor ? (actor.name || actor.email || '') : '';
  var rows = swOnceTruthPeople_().map(function (person) {
    var existing = oldByEmail[person.email] || oldByName[swNorm_(person.name)] || null;
    var days = existing && existing.days ? existing.days : swOnceDefaultDaysForPerson_(person);
    var valuesByHeaderKey = {
      rep: person.name,
      name: person.name,
      teammember: person.name,
      email: person.email,
      repemail: person.email,
      role: person.role,
      roles: person.role,
      active: person.active ? 'Y' : 'N',
      mon: days.Mon ? 'Y' : 'N',
      tue: days.Tue ? 'Y' : 'N',
      wed: days.Wed ? 'Y' : 'N',
      thu: days.Thu ? 'Y' : 'N',
      fri: days.Fri ? 'Y' : 'N',
      sat: days.Sat ? 'Y' : 'N',
      sun: days.Sun ? 'Y' : 'N',
      defaultjoc: person.defaultJoc || '',
      linkedjoc: person.defaultJoc || '',
      jocpartner: person.defaultJoc || '',
      assistedcoverageenabled: 'Y',
      coverageenabled: 'Y',
      assistedcoveragepartner: '',
      coveragepartner: '',
      labdiamond: person.lab ? 'Y' : 'N',
      naturaldiamond: swNormalizeNaturalSkill_(person.natural || 'None'),
      generalappointment: person.general ? 'Y' : 'N',
      skillnotes: '',
      updatedat: now,
      updatedby: actorLabel
    };
    return headers.map(function (header) {
      var key = swHeaderKey_(header);
      return valuesByHeaderKey[key] == null ? '' : valuesByHeaderKey[key];
    });
  });
  var oldRows = Math.max(0, sh.getLastRow() - 1);
  if (oldRows > 0) sh.getRange(2, 1, oldRows, sh.getLastColumn()).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  result.actions.push('Rewrote 10_Roster_Schedule with ' + rows.length + ' canonical roster row(s).');
}

function swOnceStandardizeDropdownIdentity_(ss, result) {
  var sh = ss.getSheetByName(SW_SHEETS.DROPDOWN);
  if (!sh || sh.getLastRow() < 2) return;
  var values = sh.getDataRange().getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  [
    swPickIndex_(H, ['Assigned Rep', 'Client Advisor']),
    swPickIndex_(H, ['Assisted Rep', 'Assistant Rep', 'JOC'])
  ].forEach(function (col) {
    if (col < 0) return;
    for (var r = 1; r < values.length; r++) {
      var oldValue = values[r][col];
      var next = swOnceCanonicalIdentityList_(oldValue);
      if (next === swTrim_(oldValue)) continue;
      swOnceSetCell_(ss, sh, r + 1, col + 1, oldValue, next, 'standardize legacy identity name');
      result.actions.push('Standardized Dropdown row ' + (r + 1) + ' ' + headers[col] + ': ' + oldValue + ' -> ' + next);
    }
  });
}

function swOnceCleanRecentAppointmentOwners_(ss, result) {
  var sh = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!sh || sh.getLastRow() < 2) return;
  var headers = swEnsureMasterOwnerHeaders_(sh);
  var allHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(allHeaders);
  var dateCol = swPickIndex_(H, ['Visit Date', 'Appointment Date', 'Date']) + 1;
  if (dateCol <= 0) return;
  var display = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var truth = swOnceTruthIndex_();
  var changedRows = 0;
  var skippedRows = 0;
  for (var i = 0; i < display.length; i++) {
    var dateKey = swOnceDateKey_(values[i][dateCol - 1] || display[i][dateCol - 1]);
    if (!dateKey || dateKey < SW_ONCE_PEOPLE_CLEANUP_CUTOFF_) continue;
    var rowNumber = i + 2;
    var assigned = swTrim_(display[i][headers.assignedRep - 1]);
    var assisted = swTrim_(display[i][headers.assistedRep - 1]);
    var ownerPlan = swOnceAppointmentOwnerPlan_(truth, assigned, assisted);
    if (ownerPlan.skips.length) skippedRows++;
    ownerPlan.actions.forEach(function (action) {
      var col = action.column === 'Assigned Rep' ? headers.assignedRep :
        action.column === 'Assigned Rep Email' ? headers.assignedRepEmail :
        action.column === 'Assisted Rep' ? headers.assistedRep :
        headers.assistedRepEmail;
      swOnceSetCell_(ss, sh, rowNumber, col, action.oldValue, action.newValue, action.reason);
      changedRows++;
    });
  }
  result.actions.push('Cleaned ' + changedRows + ' recent appointment owner cell(s); skipped ' + skippedRows + ' multi-owner row(s).');
}

function swOnceBackupAndDeleteRepQualifications_(ss, result) {
  var sh = ss.getSheetByName(SW_ONCE_REP_QUALIFICATIONS_SHEET_);
  if (!sh) {
    result.warnings.push('Rep Qualifications tab was already absent.');
    return;
  }
  var values = sh.getDataRange().getDisplayValues();
  var backup = ss.getSheetByName(SW_ONCE_RETIRED_REP_QUALIFICATIONS_BACKUP_SHEET_) || ss.insertSheet(SW_ONCE_RETIRED_REP_QUALIFICATIONS_BACKUP_SHEET_);
  backup.clearContents();
  if (values.length && values[0].length) {
    backup.getRange(1, 1, values.length, values[0].length).setValues(values);
  }
  swStyleSheet_(backup);
  if (ss.getSheets().length <= 1) throw new Error('Cannot delete the last remaining sheet.');
  ss.deleteSheet(sh);
  result.actions.push('Backed up and deleted Rep Qualifications.');
}

function swOnceAppointmentOwnerPlan_(truth, assigned, assisted) {
  var out = { actions: [], skips: [] };
  var advisorParts = swOnceSplitOwnerList_(assigned);
  var jocParts = swOnceSplitOwnerList_(assisted);
  if (advisorParts.length > 1 || jocParts.length > 1) {
    if (advisorParts.length > 1) out.skips.push({ column: 'Assigned Rep', value: assigned, reason: 'Multi-owner Client Advisor row skipped without changes.' });
    if (jocParts.length > 1) out.skips.push({ column: 'Assisted Rep', value: assisted, reason: 'Multi-owner JOC row skipped without changes.' });
    return out;
  }
  var advisor = advisorParts.length === 1 ? swOnceTruthByName_(truth, advisorParts[0], SW_OWNER_ROLES.SALES_REP) : null;
  if (advisor) {
    if (assigned !== advisor.name) out.actions.push({ column: 'Assigned Rep', oldValue: assigned, newValue: advisor.name, reason: 'canonical Client Advisor name' });
    out.actions.push({ column: 'Assigned Rep Email', oldValue: '', newValue: advisor.email, reason: 'canonical Client Advisor email' });
  } else if (assigned) {
    out.skips.push({ column: 'Assigned Rep', value: assigned, reason: 'Client Advisor is not a single active canonical advisor.' });
  }

  if (jocParts.length === 1) {
    var joc = swOnceTruthByName_(truth, jocParts[0], SW_OWNER_ROLES.JOC);
    if (joc && joc.active) {
      if (assisted !== joc.name) out.actions.push({ column: 'Assisted Rep', oldValue: assisted, newValue: joc.name, reason: 'canonical JOC name' });
      out.actions.push({ column: 'Assisted Rep Email', oldValue: '', newValue: joc.email, reason: 'canonical JOC email' });
    } else if (joc && !joc.active && advisor && advisor.defaultJoc) {
      var replacement = swOnceTruthByName_(truth, advisor.defaultJoc, SW_OWNER_ROLES.JOC);
      if (replacement && replacement.active) {
        out.actions.push({ column: 'Assisted Rep', oldValue: assisted, newValue: replacement.name, reason: 'replace inactive Maria with advisor default JOC' });
        out.actions.push({ column: 'Assisted Rep Email', oldValue: '', newValue: replacement.email, reason: 'advisor default JOC email' });
      }
    } else {
      out.skips.push({ column: 'Assisted Rep', value: assisted, reason: 'JOC is not a single active canonical JOC.' });
    }
  }
  return out;
}

function swOnceTruthIndex_() {
  var byName = {};
  var byEmail = {};
  swOnceTruthPeople_().forEach(function (person) {
    byName[swNorm_(person.name)] = person;
    byEmail[person.email] = person;
  });
  return { byName: byName, byEmail: byEmail };
}

function swOnceTruthByName_(truth, name, role) {
  var canonical = swOnceCanonicalIdentity_(name);
  var person = truth.byName[swNorm_(canonical || name)] || null;
  if (!person || !swWorkflowRoleMatches_(person.role, role)) return null;
  return person;
}

function swOnceCanonicalIdentityList_(value) {
  var parts = swOnceSplitOwnerList_(value);
  if (!parts.length) return '';
  return parts.map(function (part) {
    return swOnceCanonicalIdentity_(part) || swTrim_(part);
  }).join(', ');
}

function swOnceCanonicalIdentity_(value) {
  var key = swNorm_(value);
  var aliases = {
    'kris (tv)': 'Kris',
    'wendy (pm)': 'Wendy',
    'an vo (ha)': 'An Vo',
    'lyn ngoc': 'Lyn',
    'mark': 'Mark'
  };
  return aliases[key] || swTrim_(value);
}

function swOnceSplitOwnerList_(value) {
  value = swTrim_(value || '');
  if (!value) return [];
  return value.split(',').map(function (part) { return swTrim_(part); }).filter(Boolean);
}

function swOnceDefaultDaysForPerson_(person) {
  if (person.name === 'Maria') return { Mon: false, Tue: false, Wed: false, Thu: false, Fri: false, Sat: false, Sun: false };
  return { Mon: true, Tue: true, Wed: true, Thu: true, Fri: true, Sat: false, Sun: false };
}

function swOnceDateKey_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, swTimezone_(), 'yyyy-MM-dd');
  if (typeof value === 'number' && isFinite(value) && value > 0) {
    var ms = Math.round((value - 25569) * 86400 * 1000);
    return Utilities.formatDate(new Date(ms), swTimezone_(), 'yyyy-MM-dd');
  }
  return swScheduleDateKey_(value);
}

function swOncePlanRow_(type, sheetName, rowNumber, columnName, oldValue, newValue, reason, status) {
  return {
    type: type,
    sheetName: sheetName,
    rowNumber: rowNumber || '',
    columnName: columnName || '',
    oldValue: oldValue == null ? '' : oldValue,
    newValue: newValue == null ? '' : newValue,
    reason: reason || '',
    status: status || 'READY'
  };
}

function swOnceWritePeopleCleanupPlan_(ss, plan) {
  var sh = ss.getSheetByName(SW_ONCE_PEOPLE_CLEANUP_PLAN_SHEET_) || ss.insertSheet(SW_ONCE_PEOPLE_CLEANUP_PLAN_SHEET_);
  sh.clearContents();
  var headers = ['Generated At', 'Type', 'Sheet', 'Row', 'Column', 'Old Value', 'New Value', 'Reason', 'Status'];
  var rows = (plan.rows || []).map(function (row) {
    return [plan.generatedAt, row.type, row.sheetName, row.rowNumber, row.columnName, row.oldValue, row.newValue, row.reason, row.status];
  });
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  swStyleSheet_(sh);
}

function swOnceSetCell_(ss, sh, row, col, oldValue, newValue, reason) {
  oldValue = oldValue == null ? '' : oldValue;
  newValue = newValue == null ? '' : newValue;
  var current = sh.getRange(row, col).getDisplayValue();
  if (swTrim_(current) === swTrim_(newValue)) return false;
  swOnceBackupCell_(ss, sh.getName(), row, col, current || oldValue, newValue, reason);
  sh.getRange(row, col).setValue(newValue);
  return true;
}

function swOnceBackupCell_(ss, sheetName, row, col, oldValue, newValue, reason) {
  var sh = ss.getSheetByName(SW_ONCE_PEOPLE_CLEANUP_BACKUP_SHEET_) || ss.insertSheet(SW_ONCE_PEOPLE_CLEANUP_BACKUP_SHEET_);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 8).setValues([['Backed Up At', 'Sheet', 'Row', 'Column', 'A1', 'Old Value', 'New Value', 'Reason']]);
  }
  sh.appendRow([swIso_(new Date()), sheetName, row, col, swOnceA1_(row, col), oldValue == null ? '' : oldValue, newValue == null ? '' : newValue, reason || '']);
}

function swOnceBackupBlock_(ss, sheetName, values, reason) {
  var sh = ss.getSheetByName(SW_ONCE_PEOPLE_CLEANUP_BACKUP_SHEET_) || ss.insertSheet(SW_ONCE_PEOPLE_CLEANUP_BACKUP_SHEET_);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 8).setValues([['Backed Up At', 'Sheet', 'Row', 'Column', 'A1', 'Old Value', 'New Value', 'Reason']]);
  }
  var now = swIso_(new Date());
  var rows = [];
  (values || []).forEach(function (row, r) {
    (row || []).forEach(function (value, c) {
      if (value === '') return;
      rows.push([now, sheetName, r + 1, c + 1, swOnceA1_(r + 1, c + 1), value, '', reason || '']);
    });
  });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
}

function swOnceA1_(row, col) {
  var s = '';
  while (col > 0) {
    var m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - m) / 26);
  }
  return s + row;
}
