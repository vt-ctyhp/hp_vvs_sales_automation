/**
 * Sales workflow data repository: context, source appointment reads, roster, and schedule indexes.
 */

var SW_APPOINTMENT_ROOT_ROW_CACHE_SECONDS = 5 * 60;

function swReadRosterAvailabilityIndex_(ss) {
  var out = { exists: false, schemaOk: false, byName: {}, byEmail: {} };
  var roster = ss.getSheetByName(SW_SHEETS.ROSTER);
  var rosterRows = swReadEmployeeRosterRows_(ss);
  var rosterByEmail = {};
  var rosterByName = {};
  rosterRows.forEach(function (row) {
    if (row.email && !rosterByEmail[row.email]) rosterByEmail[row.email] = row;
    if (row.name && !rosterByName[swNorm_(row.name)]) rosterByName[swNorm_(row.name)] = row;
  });
  var users = swReadCanonicalWorkflowPeople_(ss, { schedulableOnly: true, includeInactive: true });
  if (!roster && !users.length) return out;
  out.exists = !!(roster || users.length);

  out.schemaOk = true;
  if (roster) {
    var values = roster.getDataRange().getDisplayValues();
    var headers = values[0].map(function (h) { return swTrim_(h); });
    var H = swHeaderMapFromArray_(headers);
    var repCol = swPickIndex_(H, ['Rep', 'Name', 'Team Member']);
    var dayCols = {
      Sun: swPickIndex_(H, ['Sun']),
      Mon: swPickIndex_(H, ['Mon']),
      Tue: swPickIndex_(H, ['Tue']),
      Wed: swPickIndex_(H, ['Wed']),
      Thu: swPickIndex_(H, ['Thu']),
      Fri: swPickIndex_(H, ['Fri']),
      Sat: swPickIndex_(H, ['Sat'])
    };
    out.schemaOk = repCol >= 0 && Object.keys(dayCols).some(function (day) { return dayCols[day] >= 0; });
    if (!out.schemaOk) return out;
  }

  users.forEach(function (user) {
    var rosterRow = (user.email && rosterByEmail[user.email]) || rosterByName[swNorm_(user.name)] || null;
    var row = {
      name: user.name,
      email: user.email,
      role: user.scheduleRole,
      active: user.active !== false && (!rosterRow || rosterRow.active !== false),
      defaultJoc: rosterRow ? rosterRow.defaultJoc || '' : '',
      coverageEnabled: !rosterRow || rosterRow.coverageEnabled !== false,
      coveragePartner: rosterRow ? rosterRow.coveragePartner || '' : '',
      days: rosterRow && rosterRow.days ? rosterRow.days : swDefaultEmployeeScheduleDays_(),
      skills: rosterRow && rosterRow.skills ? rosterRow.skills : swDefaultRepSkills_(),
      skillNotes: rosterRow ? rosterRow.skillNotes || '' : ''
    };
    out.byName[swNorm_(row.name)] = row;
    if (row.email) out.byEmail[row.email] = row;
  });
  return out;
}

function swReadScheduleChangesIndex_(ss) {
  var out = { byNameDate: {}, byEmailDate: {} };
  var sh = ss.getSheetByName(SW_SHEETS.SCHEDULE_CHANGES);
  if (!sh || sh.getLastRow() < 2) return out;

  var values = sh.getDataRange().getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var nameCol = swPickIndex_(H, ['Rep Name', 'Rep', 'Name']);
  var emailCol = swPickIndex_(H, ['Email', 'Rep Email']);
  var roleCol = swPickIndex_(H, ['Role', 'Roles']);
  var dateCol = swPickIndex_(H, ['Change Date', 'Date']);
  var typeCol = swPickIndex_(H, ['Change Type', 'Status', 'Override Status']);
  var fromCol = swPickIndex_(H, ['Available From', 'From']);
  var untilCol = swPickIndex_(H, ['Available Until', 'Until']);
  var notesCol = swPickIndex_(H, ['Notes', 'Note']);
  if (nameCol < 0 || dateCol < 0) return out;

  for (var i = 1; i < values.length; i++) {
    var name = swNorm_(values[i][nameCol]);
    var date = swScheduleDateKey_(values[i][dateCol]);
    if (!name || !date) continue;
    out.byNameDate[name + '|' + date] = {
      name: swTrim_(values[i][nameCol]),
      email: emailCol >= 0 ? swNormEmail_(values[i][emailCol]) : '',
      role: roleCol >= 0 ? swTrim_(values[i][roleCol]) : '',
      date: date,
      changeType: typeCol >= 0 ? swTrim_(values[i][typeCol]) : 'Full-day off',
      availableFrom: fromCol >= 0 ? swTrim_(values[i][fromCol]) : '',
      availableUntil: untilCol >= 0 ? swTrim_(values[i][untilCol]) : '',
      notes: notesCol >= 0 ? swTrim_(values[i][notesCol]) : '',
      rowNumber: i + 1
    };
    if (out.byNameDate[name + '|' + date].email) {
      out.byEmailDate[out.byNameDate[name + '|' + date].email + '|' + date] = out.byNameDate[name + '|' + date];
    }
  }
  return out;
}

function swEnsureEmployeeScheduleSheets_(ss) {
  var roster = swEnsureSheet_(ss, SW_SHEETS.ROSTER, SW_EMPLOYEE_SCHEDULE_HEADERS);
  var changes = swEnsureSheet_(ss, SW_SHEETS.SCHEDULE_CHANGES, SW_SCHEDULE_CHANGE_HEADERS);
  swStyleSheet_(roster);
  swStyleSheet_(changes);
  return { roster: roster, changes: changes };
}

function swReadEmployeeScheduleAdminData_(ss) {
  swEnsureEmployeeScheduleSheets_(ss);
  var config = swReadConfig_(ss, true);
  var people = swReadEmployeeSchedulePeople_(ss);
  return {
    ok: true,
    generatedAt: swIso_(new Date()),
    days: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],
    changeTypes: ['Working', 'Full-day off', 'Late / partial day', 'PTO', 'Sick', 'Vacation'],
    roleOptions: [SW_OWNER_ROLES.SALES_REP, SW_OWNER_ROLES.JOC],
    naturalOptions: ['None', 'Primary', 'Backup'],
    settings: {
      clientAdvisorRoundRobin: swTruthy_(swConfigValue_(config, 'SYSTEM', 'CLIENT_ADVISOR_ROUND_ROBIN', 'N'))
    },
    today: swEmployeeScheduleToday_(ss, people),
    people: people,
    changes: swReadEmployeeScheduleChanges_(ss)
  };
}

function swReadEmployeeSchedulePeople_(ss) {
  var rosterByEmail = {};
  var rosterByName = {};
  swReadEmployeeRosterRows_(ss).forEach(function (row) {
    if (row.email && !rosterByEmail[row.email]) rosterByEmail[row.email] = row;
    if (row.name && !rosterByName[swNorm_(row.name)]) rosterByName[swNorm_(row.name)] = row;
  });

  var people = [];
  var seenEmails = {};
  var seenRosterRows = {};
  swReadCanonicalWorkflowPeople_(ss, { schedulableOnly: true, includeInactive: true }).forEach(function (user) {
    var roster = (user.email && rosterByEmail[user.email]) || rosterByName[swNorm_(user.name)] || null;
    var row = {
      rowNumber: roster ? roster.rowNumber : 0,
      name: user.name,
      email: user.email,
      role: user.scheduleRole,
      active: user.active !== false && (!roster || roster.active !== false),
      userActive: user.active !== false,
      rosterActive: !roster || roster.active !== false,
      identityStatus: 'canonical',
      defaultJoc: roster ? roster.defaultJoc || '' : '',
      coverageEnabled: !roster || roster.coverageEnabled !== false,
      coveragePartner: roster ? roster.coveragePartner || '' : '',
      days: roster && roster.days ? roster.days : swDefaultEmployeeScheduleDays_(),
      skills: roster && roster.skills ? roster.skills : swDefaultRepSkills_(),
      skillNotes: roster ? roster.skillNotes || '' : ''
    };
    people.push(row);
    if (row.email) seenEmails[row.email] = true;
    if (roster && roster.rowNumber) seenRosterRows[roster.rowNumber] = true;
  });

  swReadEmployeeRosterRows_(ss).forEach(function (row) {
    if (row.rowNumber && seenRosterRows[row.rowNumber]) return;
    if (row.email && seenEmails[row.email]) return;
    if (!row.role && row.active === false) return;
    people.push(swMergeObjects_(row, {
      active: false,
      userActive: false,
      rosterActive: row.active !== false,
      identityStatus: 'orphan',
      skills: row.skills || swDefaultRepSkills_(),
      skillNotes: row.skillNotes || 'No matching active workflow user.'
    }));
  });

  return people.sort(function (a, b) {
    var ar = swEmployeePrimaryRoleRank_(a.role);
    var br = swEmployeePrimaryRoleRank_(b.role);
    if (ar !== br) return ar - br;
    return String(a.name || '').localeCompare(String(b.name || ''));
  });
}

function swReadEmployeeRosterRows_(ss) {
  var sh = ss.getSheetByName(SW_SHEETS.ROSTER);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getDataRange().getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var C = {
    name: swPickIndex_(H, ['Rep', 'Name', 'Team Member']),
    email: swPickIndex_(H, ['Email', 'Rep Email']),
    role: swPickIndex_(H, ['Role', 'Roles']),
    active: swPickIndex_(H, ['Active?', 'Active']),
    defaultJoc: swPickIndex_(H, ['Default JOC', 'Linked JOC', 'JOC Partner']),
    coverageEnabled: swPickIndex_(H, ['Assisted Coverage Enabled?', 'Coverage Enabled?']),
    coveragePartner: swPickIndex_(H, ['Assisted Coverage Partner', 'Coverage Partner']),
    lab: swPickIndex_(H, ['Lab Diamond', 'Lab']),
    natural: swPickIndex_(H, ['Natural Diamond', 'Natural']),
    general: swPickIndex_(H, ['General Appointment', 'General']),
    skillNotes: swPickIndex_(H, ['Skill Notes', 'Skills Notes', 'Notes'])
  };
  var days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  var dayCols = {};
  days.forEach(function (day) { dayCols[day] = swPickIndex_(H, [day]); });
  if (C.name < 0) return [];

  var out = [];
  for (var i = 1; i < values.length; i++) {
    var name = swTrim_(values[i][C.name]);
    if (!name) continue;
    var row = {
      rowNumber: i + 1,
      name: name,
      email: C.email >= 0 ? swNormEmail_(values[i][C.email]) : '',
      role: C.role >= 0 ? swNormalizeEmployeeRoleList_(values[i][C.role]) : '',
      active: C.active < 0 || !swTrim_(values[i][C.active]) || swTruthy_(values[i][C.active]),
      defaultJoc: C.defaultJoc >= 0 ? swTrim_(values[i][C.defaultJoc]) : '',
      coverageEnabled: C.coverageEnabled < 0 || !swTrim_(values[i][C.coverageEnabled]) || swTruthy_(values[i][C.coverageEnabled]),
      coveragePartner: C.coveragePartner >= 0 ? swTrim_(values[i][C.coveragePartner]) : '',
      skills: {
        labDiamond: C.lab >= 0 ? swTruthy_(values[i][C.lab]) : false,
        naturalDiamond: C.natural >= 0 ? swNormalizeNaturalSkill_(values[i][C.natural]) : 'None',
        generalAppointment: C.general < 0 || !swTrim_(values[i][C.general]) || swTruthy_(values[i][C.general])
      },
      skillNotes: C.skillNotes >= 0 ? swTrim_(values[i][C.skillNotes]) : '',
      days: {}
    };
    days.forEach(function (day) {
      row.days[day] = dayCols[day] >= 0 ? swTruthy_(values[i][dayCols[day]]) : (day !== 'Sat' && day !== 'Sun');
    });
    out.push(row);
  }
  return out;
}

function swReadCanonicalWorkflowPeople_(ss, options) {
  options = options || {};
  var out = [];
  var rows = [];
  try {
    rows = swAuthReadPublicUserRowsCached_(ss) || [];
  } catch (_) {
    var sh = ss.getSheetByName(SW_SHEETS.USERS);
    rows = sh ? swReadSheetObjectsExpectedHeaders_(sh, SW_AUTH_USER_HEADERS).map(function (row) {
      return swAuthPublicUserRow_(row);
    }) : [];
  }
  rows.forEach(function (row) {
    row = row || {};
    var email = swNormEmail_(row.email || row['Email']);
    if (!email) return;
    var name = swTrim_(row.name || row['Name']) || email;
    var roles = swAuthRoles_(row.roles || row['Roles']);
    var scheduleRole = swWorkflowSchedulableRoleList_(roles);
    var active = swWorkflowUserActive_(row);
    if (options.activeOnly && !active) return;
    if (options.schedulableOnly && !scheduleRole) return;
    out.push({
      email: email,
      name: name,
      roles: roles,
      role: roles.join(','),
      scheduleRole: scheduleRole,
      active: active,
      passwordSet: !!(row.passwordSet || row['Password Hash'])
    });
  });
  return out;
}

function swCanonicalWorkflowPeopleIndex_(ss, options) {
  var users = swReadCanonicalWorkflowPeople_(ss, options || {});
  var out = {
    users: users,
    byEmail: {},
    byName: {},
    activeJocByName: {},
    activeJocByEmail: {}
  };
  users.forEach(function (user) {
    if (user.email && !out.byEmail[user.email]) out.byEmail[user.email] = user;
    if (user.name && !out.byName[swNorm_(user.name)]) out.byName[swNorm_(user.name)] = user;
    if (user.active !== false && swWorkflowUserHasSchedulableRole_(user, SW_OWNER_ROLES.JOC)) {
      out.activeJocByName[swNorm_(user.name)] = user;
      out.activeJocByEmail[user.email] = user;
    }
  });
  return out;
}

function swWorkflowSchedulableRoleList_(roles) {
  var out = [];
  roles = Array.isArray(roles) ? roles : swAuthRoles_(roles || '');
  if (swAuthHasRole_(roles, SW_OWNER_ROLES.SALES_REP)) out.push(SW_OWNER_ROLES.SALES_REP);
  if (swAuthHasRole_(roles, SW_OWNER_ROLES.JOC)) out.push(SW_OWNER_ROLES.JOC);
  return out.join(',');
}

function swWorkflowUserActive_(row) {
  row = row || {};
  var value = row.active;
  if (value == null || value === '') value = row['Active?'];
  return value == null || swTrim_(value) === '' || swTruthy_(value);
}

function swWorkflowUserHasSchedulableRole_(user, role) {
  return swNormalizeEmployeeRoleList_((user && (user.scheduleRole || user.role || user.roles)) || '').split(',').some(function (item) {
    return swWorkflowRoleMatches_(item, role);
  });
}

function swDefaultEmployeeScheduleDays_() {
  return { Mon: true, Tue: true, Wed: true, Thu: true, Fri: true, Sat: false, Sun: false };
}

function swDefaultRepSkills_() {
  return { labDiamond: false, naturalDiamond: 'None', generalAppointment: true };
}

function swNormalizeEmployeeRoleList_(value) {
  var roles = [];
  String(value || '').split(/[,\n;]/).forEach(function (role) {
    var normalized = swNormalizeEmployeeRole_(role);
    if (normalized && roles.indexOf(normalized) < 0) roles.push(normalized);
  });
  return roles.join(',');
}

function swNormalizeEmployeeRole_(role) {
  if (swWorkflowRoleMatches_(role, SW_OWNER_ROLES.SALES_REP)) return SW_OWNER_ROLES.SALES_REP;
  if (swWorkflowRoleMatches_(role, SW_OWNER_ROLES.JOC)) return SW_OWNER_ROLES.JOC;
  return '';
}

function swEmployeeHasRole_(person, role) {
  var target = swNormalizeEmployeeRole_(role);
  if (!target) return false;
  return swNormalizeEmployeeRoleList_((person && person.role) || '').split(',').some(function (item) {
    return swWorkflowRoleMatches_(item, target);
  });
}

function swEmployeePrimaryRoleRank_(role) {
  if (String(role || '').indexOf(SW_OWNER_ROLES.JOC) >= 0) return 1;
  if (String(role || '').indexOf(SW_OWNER_ROLES.SALES_REP) >= 0) return 2;
  return 9;
}

function swNormalizeNaturalSkill_(value) {
  var s = swNorm_(value);
  if (s === 'primary') return 'Primary';
  if (s === 'backup') return 'Backup';
  if (s === 'y' || s === 'yes' || s === 'true' || s === '1') return 'Primary';
  return 'None';
}

function swReadEmployeeScheduleChanges_(ss) {
  var sh = ss.getSheetByName(SW_SHEETS.SCHEDULE_CHANGES);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getDataRange().getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var C = {
    name: swPickIndex_(H, ['Rep Name', 'Rep', 'Name']),
    email: swPickIndex_(H, ['Email', 'Rep Email']),
    role: swPickIndex_(H, ['Role', 'Roles']),
    date: swPickIndex_(H, ['Change Date', 'Date']),
    type: swPickIndex_(H, ['Change Type', 'Status', 'Override Status']),
    from: swPickIndex_(H, ['Available From', 'From']),
    until: swPickIndex_(H, ['Available Until', 'Until']),
    notes: swPickIndex_(H, ['Notes', 'Note'])
  };
  if (C.name < 0 || C.date < 0) return [];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var name = swTrim_(values[i][C.name]);
    var date = swScheduleDateKey_(values[i][C.date]);
    if (!name || !date) continue;
    out.push({
      rowNumber: i + 1,
      name: name,
      email: C.email >= 0 ? swNormEmail_(values[i][C.email]) : '',
      role: C.role >= 0 ? swNormalizeEmployeeRoleList_(values[i][C.role]) : '',
      date: date,
      changeType: C.type >= 0 ? swTrim_(values[i][C.type]) || 'Full-day off' : 'Full-day off',
      availableFrom: C.from >= 0 ? swTrim_(values[i][C.from]) : '',
      availableUntil: C.until >= 0 ? swTrim_(values[i][C.until]) : '',
      notes: C.notes >= 0 ? swTrim_(values[i][C.notes]) : ''
    });
  }
  out.sort(function (a, b) {
    return String(a.date || '').localeCompare(String(b.date || '')) ||
      String(a.name || '').localeCompare(String(b.name || ''));
  });
  return out;
}

function swEmployeeScheduleToday_(ss, people) {
  var today = new Date();
  var ctx = {
    rosterIndex: swReadRosterAvailabilityIndex_(ss),
    scheduleChangesIndex: swReadScheduleChangesIndex_(ss)
  };
  return (people || []).filter(function (person) {
    return person && person.active !== false && typeof swAvailabilityFor_ === 'function' &&
      swAvailabilityFor_(ss, person.name, today, ctx).available;
  }).map(function (person) {
    var override = swScheduleOverride_(ss, person.name, today, ctx);
    return {
      name: person.name,
      role: person.role || '',
      availableFrom: override ? override.availableFrom || '' : '',
      availableUntil: override ? override.availableUntil || '' : ''
    };
  });
}

function swCanonicalizeEmployeeScheduleRowsForWrite_(ss, people) {
  var index = swCanonicalWorkflowPeopleIndex_(ss, { schedulableOnly: true, includeInactive: true });
  var activeJocs = swCanonicalWorkflowPeopleIndex_(ss, { schedulableOnly: true, activeOnly: true });
  var out = [];
  var seenEmails = {};
  (people || []).forEach(function (person) {
    person = person || {};
    var email = swNormEmail_(person.email || '');
    var name = swTrim_(person.name || '');
    var user = email ? index.byEmail[email] : null;
    if (!user && name) user = index.byName[swNorm_(name)] || null;
    if (!email) throw new Error('Roster row for "' + (name || 'unnamed person') + '" is missing a workflow user email.');
    if (!user) throw new Error('Roster row for "' + (name || email) + '" does not match an active workflow user.');
    if (seenEmails[user.email]) throw new Error('Duplicate roster row for workflow user email: ' + user.email);
    seenEmails[user.email] = true;
    var role = user.scheduleRole || swWorkflowSchedulableRoleList_(user.roles);
    if (!role) throw new Error('Workflow user "' + user.name + '" does not have Client Advisor or JOC access.');
    var defaultJoc = swCanonicalJocNameForScheduleWrite_(activeJocs, person.defaultJoc, user);
    var coveragePartner = swCanonicalJocNameForScheduleWrite_(activeJocs, person.coveragePartner, user);
    var days = person.days || {};
    out.push({
      name: user.name,
      email: user.email,
      role: role,
      active: user.active !== false && swTruthy_(person.active == null ? 'Y' : person.active),
      days: {
        Mon: swTruthy_(days.Mon),
        Tue: swTruthy_(days.Tue),
        Wed: swTruthy_(days.Wed),
        Thu: swTruthy_(days.Thu),
        Fri: swTruthy_(days.Fri),
        Sat: swTruthy_(days.Sat),
        Sun: swTruthy_(days.Sun)
      },
      defaultJoc: defaultJoc,
      coverageEnabled: person.coverageEnabled == null ? true : swTruthy_(person.coverageEnabled),
      coveragePartner: coveragePartner,
      skills: person.skills || swDefaultRepSkills_(),
      skillNotes: person.skillNotes || ''
    });
  });
  return out;
}

function swCanonicalJocNameForScheduleWrite_(index, value, owner) {
  var raw = swTrim_(value || '');
  if (!raw) return '';
  var user = index.activeJocByName[swNorm_(raw)] || index.activeJocByEmail[swNormEmail_(raw)] || null;
  if (!user) throw new Error('JOC routing value "' + raw + '" must be an active workflow user with JOC access.');
  if (owner && user.email && owner.email && user.email === owner.email) {
    throw new Error('JOC routing value for "' + owner.name + '" cannot point to the same person.');
  }
  return user.name;
}

function swScheduleDateKey_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, swTimezone_(), 'yyyy-MM-dd');
  var s = swTrim_(value);
  if (!s) return '';
  var iso = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
  if (iso) return [iso[1], swPad2_(iso[2]), swPad2_(iso[3])].join('-');
  var us = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s);
  if (us) return [us[3], swPad2_(us[1]), swPad2_(us[2])].join('-');
  return swDateKey_(value);
}

function swBuildIdentityContext_(ss, readOnly) {
  var config = swReadConfig_(ss, readOnly);
  var peopleIndex = swReadPeopleIndex_(ss, config);
  var admins = swReadAdminsFromConfig_(config);
  return {
    tz: swTimezone_(),
    config: config,
    peopleIndex: peopleIndex,
    assistedRoster: peopleIndex.assistedRoster,
    admins: admins,
    lookbackDays: Number(swConfigValue_(config, 'SYSTEM', 'WORKFLOW_LOOKBACK_DAYS', '14')) || 14,
    futureDays: Number(swConfigValue_(config, 'SYSTEM', 'WORKFLOW_FUTURE_DAYS', '365')) || 365
  };
}

function swBuildTaskDetailContext_(ss, readOnly) {
  var ctx = swBuildIdentityContext_(ss, readOnly);
  ctx.templates = swReadTemplates_(ss, readOnly);
  return ctx;
}

function swBuildTaskDetailReadContext_(ss, readOnly) {
  var user = swCurrentUserConfigOnly_(ss, readOnly);
  if (user.isAdmin) {
    return {
      user: user,
      templates: swReadTemplates_(ss, readOnly),
      lightweight: true
    };
  }

  var ctx = swBuildTaskDetailContext_(ss, readOnly);
  ctx.user = swCurrentUser_(ss, ctx);
  ctx.lightweight = false;
  return ctx;
}

function swBuildContext_(ss, readOnly) {
  var ctx = swBuildTaskDetailContext_(ss, readOnly);
  ctx.rosterIndex = swReadRosterAvailabilityIndex_(ss);
  ctx.scheduleChangesIndex = swReadScheduleChangesIndex_(ss);
  ctx.waxIndex = swReadWaxRequestIndex_(ss);
  return ctx;
}

function swReadWaxRequestIndex_(ss) {
  var out = { byRoot: {}, activeByRoot: {}, needsUpdateByRoot: {}, statusOptions: [] };
  var sheetName = (typeof WAX !== 'undefined' && WAX.SHEET) ? WAX.SHEET : '05_Wax_Requests';
  var sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < 1) return out;

  var values = sh.getDataRange().getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var C = {
    id: swPickIndex_(H, ['WaxRequestID']),
    root: swPickIndex_(H, ['RootApptID']),
    so: swPickIndex_(H, ['SO/MO Number', 'SO Number', 'SO#']),
    customer: swPickIndex_(H, ['Customer Name']),
    priority: swPickIndex_(H, ['Priority']),
    status: swPickIndex_(H, ['Wax Print Status']),
    repNeed: swPickIndex_(H, ['Needed By (Rep)', 'Needed by (Rep)', 'Rep Needed By']),
    adminDeadline: swPickIndex_(H, ['Wax Deadline (Admin)', 'Wax Admin Deadline']),
    estPrint: swPickIndex_(H, ['Estimated Print Date']),
    completed: swPickIndex_(H, ['Completed Print Date']),
    notes: swPickIndex_(H, ['Status Notes']),
    link: swPickIndex_(H, ['Master Row Link'])
  };
  var now = new Date();
  var todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0, 0).getTime();

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var root = swTrim_(swCell_(row, C.root));
    if (!root) continue;
    var status = swTrim_(swCell_(row, C.status));
    var statusNorm = swNorm_(status);
    var active = !/(^|\s)(completed|canceled|cancelled)(\s|$)/.test(statusNorm);
    var adminDeadline = swTrim_(swCell_(row, C.adminDeadline));
    var adminMs = adminDeadline ? swDateValue_(adminDeadline) : 0;
    var needsUpdate = active && (!status || !adminDeadline || (adminMs && adminMs < todayStart));
    var rowUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + sh.getSheetId() + '&range=A' + (i + 1);
    var item = {
      id: swTrim_(swCell_(row, C.id)),
      root: root,
      so: swTrim_(swCell_(row, C.so)),
      customerName: swTrim_(swCell_(row, C.customer)),
      priority: swTrim_(swCell_(row, C.priority)),
      status: status,
      repNeed: swTrim_(swCell_(row, C.repNeed)),
      adminDeadline: adminDeadline,
      estPrint: swTrim_(swCell_(row, C.estPrint)),
      completed: swTrim_(swCell_(row, C.completed)),
      notes: swTrim_(swCell_(row, C.notes)),
      link: swTrim_(swCell_(row, C.link)) || rowUrl,
      rowUrl: rowUrl,
      active: active,
      needsUpdate: needsUpdate
    };
    if (!out.byRoot[root]) out.byRoot[root] = [];
    out.byRoot[root].push(item);
    if (active) {
      if (!out.activeByRoot[root]) out.activeByRoot[root] = [];
      out.activeByRoot[root].push(item);
    }
    if (needsUpdate) {
      if (!out.needsUpdateByRoot[root]) out.needsUpdateByRoot[root] = [];
      out.needsUpdateByRoot[root].push(item);
    }
  }
  try {
    if (typeof wax_statusOptions === 'function') out.statusOptions = wax_statusOptions();
  } catch (_) {}
  return out;
}

function swReadAppointments_(ss) {
  var sh = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!sh) throw new Error('Missing sheet: ' + SW_SHEETS.MASTER);
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  var headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
  var idx = swAppointmentColumnIndex_(headers);
  var indexes = swAppointmentColumnIndexes_(idx);
  var rowCount = lastRow - 1;
  var values = swReadSelectedRows_(sh, 2, rowCount, indexes, 'values');
  var display = swReadSelectedRows_(sh, 2, rowCount, indexes, 'display');

  var out = [];
  for (var i = 0; i < values.length; i++) {
    out.push(swAppointmentRecordFromRows_(display[i], values[i], idx, i + 2));
  }
  swCacheAppointmentRootRows_(ss, out);
  return out;
}

function swReadAppointmentsForRoot_(ss, rootApptId) {
  var want = swTrim_(rootApptId);
  if (!want) return [];
  var sh = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!sh) throw new Error('Missing sheet: ' + SW_SHEETS.MASTER);
  if (sh.getLastRow() < 2 || sh.getLastColumn() < 1) return [];

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
  var idx = swAppointmentColumnIndex_(headers);
  if (idx.root < 0 && idx.appt < 0) {
    return swReadAppointments_(ss).filter(function (rec) {
      return swTrim_(rec.root) === want || swTrim_(rec.appt) === want;
    });
  }

  var rowCount = lastRow - 1;
  var indexes = swAppointmentColumnIndexes_(idx);
  var cachedRows = swCachedAppointmentRootRows_(ss, want);
  if (cachedRows && cachedRows.length) {
    var cachedOut = [];
    cachedRows.forEach(function (rowNumber) {
      if (rowNumber < 2 || rowNumber > lastRow) return;
      var values = swReadSelectedRows_(sh, rowNumber, 1, indexes, 'values')[0] || [];
      var display = swReadSelectedRows_(sh, rowNumber, 1, indexes, 'display')[0] || [];
      var rec = swAppointmentRecordFromRows_(display, values, idx, rowNumber);
      if (swTrim_(rec.root) === want || swTrim_(rec.appt) === want) cachedOut.push(rec);
    });
    if (cachedOut.length) return cachedOut;
  }

  var roots = idx.root >= 0 ? sh.getRange(2, idx.root + 1, rowCount, 1).getDisplayValues() : [];
  var appts = idx.appt >= 0 ? sh.getRange(2, idx.appt + 1, rowCount, 1).getDisplayValues() : [];
  var out = [];
  for (var i = 0; i < rowCount; i++) {
    var root = roots.length ? swTrim_(roots[i][0]) : '';
    var appt = appts.length ? swTrim_(appts[i][0]) : '';
    if (root !== want && appt !== want) continue;
    var rowNumber = i + 2;
    var values = swReadSelectedRows_(sh, rowNumber, 1, indexes, 'values')[0] || [];
    var display = swReadSelectedRows_(sh, rowNumber, 1, indexes, 'display')[0] || [];
    out.push(swAppointmentRecordFromRows_(display, values, idx, rowNumber));
  }
  return out;
}

function swAppointmentColumnIndexes_(idx) {
  var out = [];
  Object.keys(idx || {}).forEach(function (key) {
    var col = Number(idx[key]);
    if (isFinite(col) && col >= 0) out.push(col);
  });
  return out;
}

function swCacheAppointmentRootRows_(ss, appointments) {
  var map = {};
  (appointments || []).forEach(function (rec) {
    if (!rec || !rec.row) return;
    [rec.root, rec.appt].forEach(function (id) {
      id = swTrim_(id);
      if (!id) return;
      if (!map[id]) map[id] = [];
      if (map[id].indexOf(rec.row) < 0) map[id].push(rec.row);
    });
  });
  try {
    var payload = swStringify_({ rowsById: map, cachedAt: swIso_(new Date()) });
    if (payload.length < 90000) CacheService.getScriptCache().put(swAppointmentRootRowCacheKey_(ss), payload, SW_APPOINTMENT_ROOT_ROW_CACHE_SECONDS);
  } catch (_) {}
}

function swCachedAppointmentRootRows_(ss, rootApptId) {
  try {
    var cached = CacheService.getScriptCache().get(swAppointmentRootRowCacheKey_(ss));
    var parsed = cached ? swParseJson_(cached, null) : null;
    var rows = parsed && parsed.rowsById ? parsed.rowsById[swTrim_(rootApptId)] : null;
    return Array.isArray(rows) ? rows.map(function (row) { return Number(row); }).filter(function (row) { return isFinite(row); }) : null;
  } catch (_) {}
  return null;
}

function swAppointmentRootRowCacheKey_(ss) {
  return 'sw:appointmentRootRows:v1:' + ss.getId();
}

function swAppointmentColumnIndex_(headers) {
  var H = swHeaderMapFromArray_(headers);
  return {
    appt: swPickIndex_(H, ['APPT_ID', 'Appt ID', 'Appointment ID']),
    root: swPickIndex_(H, ['RootApptID', 'Root Appt ID', 'Root Appointment ID']),
    uid: swPickIndex_(H, ['CalendlyEventUID', 'Calendly Event UID', 'Admin: Calendly Event UID', 'Acuity ID', 'UID']),
    name: swPickIndex_(H, ['Customer Name', 'Customer', 'Client Name', 'Name']),
    emailLower: swPickIndex_(H, ['EmailLower', 'Email Lower']),
    email: swPickIndex_(H, ['Email', 'Email Address', 'E-mail']),
    phoneNorm: swPickIndex_(H, ['PhoneNorm', 'Phone Norm']),
    phone: swPickIndex_(H, ['Phone', 'Phone Number', 'Mobile', 'Tel']),
    brand: swPickIndex_(H, ['Brand', 'Company']),
    bookedAt: swPickIndex_(H, ['Booked At (ISO)', 'Booked At', 'BookedAt', 'Created At', 'CreatedAt'], false),
    canceledAt: swPickIndex_(H, ['CanceledAt', 'CancelledAt', 'Canceled At', 'Cancelled At'], false),
    rescheduledFromUid: swPickIndex_(H, ['RescheduledFromUID', 'Rescheduled From UID', 'ReschedFromUID', 'Rescheduled From'], false),
    rescheduledToUid: swPickIndex_(H, ['RescheduledToUID', 'Rescheduled To UID', 'ReschedToUID', 'Rescheduled To'], false),
    visitDate: swPickIndex_(H, ['Visit Date', 'Appointment Date', 'Date']),
    visitTime: swPickIndex_(H, ['Visit Time', 'Appointment Time', 'Time']),
    visitType: swPickIndex_(H, ['Visit Type', 'Appointment Type']),
    diamondType: swPickIndex_(H, ['Diamond Type', 'Stone Type', 'Center Stone Type']),
    status: swPickIndex_(H, ['Status']),
    active: swPickIndex_(H, ['Active?', 'Active', 'Is Active']),
    assignedRep: swPickIndex_(H, ['Client Advisor', 'Assigned Rep', 'Rep', 'Owner']),
    assignedRepEmail: swPickIndex_(H, ['Client Advisor Email', 'Assigned Rep Email', 'Rep Email', 'Owner Email']),
    assistedRep: swPickIndex_(H, ['Assisted Rep', 'Assistant Rep']),
    assistedRepEmail: swPickIndex_(H, ['Assisted Rep Email', 'Assistant Rep Email']),
    clientFolder: swPickIndex_(H, ['Client Folder', 'ClientFolderURL', 'Client Folder URL']),
    reportUrl: swPickIndex_(H, ['Client Status Report URL', 'Report URL']),
    quotationUrl: swPickIndex_(H, ['Quotation URL', 'QuotationURL', 'Quote URL']),
    tracker3d: swPickIndex_(H, ['3D Tracker', '3D Log', '3D Tracker URL']),
    salesStage: swPickIndex_(H, ['Sales Stage']),
    convStatus: swPickIndex_(H, ['Conversion Status']),
    customOrder: swPickIndex_(H, ['Custom Order Status']),
    inProduction: swPickIndex_(H, ['In Production Status']),
    nextSteps: swPickIndex_(H, ['Next Steps']),
    designRequest: swPickIndex_(H, ['Design Request']),
    deadline3d: swPickIndex_(H, ['3D Deadline']),
    productionDeadline: swPickIndex_(H, ['Production Deadline', 'Prod. Deadline']),
    waxStatus: swPickIndex_(H, ['Wax Print Status']),
    waxDeadlineAdmin: swPickIndex_(H, ['Wax Deadline (Admin)', 'Wax Admin Deadline']),
    waxRequestUrl: swPickIndex_(H, ['Wax Request URL']),
    centerStoneStatus: swPickIndex_(H, ['Center Stone Order Status', 'Center Stone Status', 'CSOS', 'Diamond Memo Status', 'DV Status']),
    dvStonesJson: swPickIndex_(H, ['DV Stones (JSON Lines)', 'DV Stones JSON Lines', 'DV Stones-JSON Lines']),
    dvStonesSummary: swPickIndex_(H, ['DV Stones Summary', 'DV Stones- Summary']),
    dvCustomerLookingFor: swPickIndex_(H, ['DV Customer Looking For', 'Diamond Customer Looking For', 'Customer Diamond Requirements']),
    dvVarietyStrategy: swPickIndex_(H, ['DV Variety Strategy', 'Diamond Variety Strategy']),
    dvCustomerRequirementsJson: swPickIndex_(H, ['DV Customer Requirements (JSON)', 'DV Customer Requirements JSON', 'Customer Diamond Requirements JSON']),
    so: swPickIndex_(H, ['SO#', 'SO #', 'SO']),
    orderFolder: swPickIndex_(H, ['Order Folder', '05-3D Folder']),
    source: swPickIndex_(H, ['Source (normalized)', 'Source Normalized', 'Source', 'Lead Source']),
    budgetMin: swPickIndex_(H, ['Budget Min', 'Budget (Min)', 'BudgetMin']),
    budgetMax: swPickIndex_(H, ['Budget Max', 'Budget (Max)', 'BudgetMax', 'Budget']),
    orderTotal: swPickIndex_(H, ['Order Total', 'OrderTotal', 'Order Total Value', 'Order_Total_SO', 'SO Total']),
    paidToDate: swPickIndex_(H, ['Paid-to-Date', 'Paid to Date', 'PaidToDate', 'Paid']),
    remainingBalance: swPickIndex_(H, ['Remaining Balance', 'Balance', 'Balance_SO', 'Balance Due']),
    lastPaymentDate: swPickIndex_(H, ['Last Payment Date', 'LastPaymentDate', 'Last Paid At']),
    orderDate: swPickIndex_(H, ['Order Date', 'SO Date', 'Sales Order Date']),
    updatedAt: swPickIndex_(H, ['Updated At', 'Last Updated At', 'Last Updated', 'UpdatedAt', 'Updated At (ISO)']),
    deadline3dMoves: swPickIndex_(H, ['# of Times 3D Deadline Moved', '3D Deadline Moves', '# 3D Deadline Moves']),
    productionDeadlineMoves: swPickIndex_(H, ['# of Times Prod. Deadline Moved', 'Prod Deadline Moves', '# Prod Deadline Moves'])
  };
}

function swAppointmentRecordFromRows_(drow, vrow, idx, rowNumber) {
  var rec = {
    row: rowNumber,
    appt: swTrim_(swCell_(drow, idx.appt)),
    root: swTrim_(swCell_(drow, idx.root)),
    uid: swTrim_(swCell_(drow, idx.uid)),
    name: swTrim_(swCell_(drow, idx.name)),
    email: swNormEmail_(swCell_(drow, idx.emailLower) || swCell_(drow, idx.email)),
    phone: swNormPhone_(swCell_(drow, idx.phoneNorm) || swCell_(drow, idx.phone)),
    brand: swTrim_(swCell_(drow, idx.brand)),
    bookedAt: swTrim_(swCell_(drow, idx.bookedAt)),
    bookedAtRaw: swCell_(vrow, idx.bookedAt),
    canceledAt: swTrim_(swCell_(drow, idx.canceledAt)),
    canceledAtRaw: swCell_(vrow, idx.canceledAt),
    rescheduledFromUid: swTrim_(swCell_(drow, idx.rescheduledFromUid)),
    rescheduledToUid: swTrim_(swCell_(drow, idx.rescheduledToUid)),
    visitDate: swTrim_(swCell_(drow, idx.visitDate)),
    visitTime: swFormatAppointmentTime_(swCell_(drow, idx.visitTime), swCell_(vrow, idx.visitTime)),
    visitType: swTrim_(swCell_(drow, idx.visitType)),
    diamondType: swTrim_(swCell_(drow, idx.diamondType)),
    visitDateRaw: swCell_(vrow, idx.visitDate),
    visitTimeRaw: swCell_(vrow, idx.visitTime),
    status: swTrim_(swCell_(drow, idx.status)),
    active: swTrim_(swCell_(drow, idx.active)),
    assignedRep: swTrim_(swCell_(drow, idx.assignedRep)),
    assignedRepEmail: swNormEmail_(swCell_(drow, idx.assignedRepEmail)),
    assistedRep: swTrim_(swCell_(drow, idx.assistedRep)),
    assistedRepEmail: swNormEmail_(swCell_(drow, idx.assistedRepEmail)),
    clientFolder: swTrim_(swCell_(drow, idx.clientFolder)),
    reportUrl: swTrim_(swCell_(drow, idx.reportUrl)),
    quotationUrl: swTrim_(swCell_(drow, idx.quotationUrl)),
    tracker3dUrl: swTrim_(swCell_(drow, idx.tracker3d)),
    salesStage: swTrim_(swCell_(drow, idx.salesStage)),
    convStatus: swTrim_(swCell_(drow, idx.convStatus)),
    customOrder: swTrim_(swCell_(drow, idx.customOrder)),
    inProduction: swTrim_(swCell_(drow, idx.inProduction)),
    nextSteps: swTrim_(swCell_(drow, idx.nextSteps)),
    designRequest: swTrim_(swCell_(drow, idx.designRequest)),
    deadline3d: swTrim_(swCell_(drow, idx.deadline3d)),
    productionDeadline: swTrim_(swCell_(drow, idx.productionDeadline)),
    waxStatus: swTrim_(swCell_(drow, idx.waxStatus)),
    waxDeadlineAdmin: swTrim_(swCell_(drow, idx.waxDeadlineAdmin)),
    waxRequestUrl: swTrim_(swCell_(drow, idx.waxRequestUrl)),
    centerStoneStatus: swTrim_(swCell_(drow, idx.centerStoneStatus)),
    dvStonesJson: swTrim_(swCell_(drow, idx.dvStonesJson)),
    dvStonesSummary: swTrim_(swCell_(drow, idx.dvStonesSummary)),
    dvCustomerLookingFor: swTrim_(swCell_(drow, idx.dvCustomerLookingFor)),
    dvVarietyStrategy: swTrim_(swCell_(drow, idx.dvVarietyStrategy)),
    dvCustomerRequirementsJson: swTrim_(swCell_(drow, idx.dvCustomerRequirementsJson)),
    so: swTrim_(swCell_(drow, idx.so)),
    orderFolder: swTrim_(swCell_(drow, idx.orderFolder)),
    source: swTrim_(swCell_(drow, idx.source)),
    budgetMin: swTrim_(swCell_(drow, idx.budgetMin)),
    budgetMax: swTrim_(swCell_(drow, idx.budgetMax)),
    orderTotal: swTrim_(swCell_(drow, idx.orderTotal)),
    paidToDate: swTrim_(swCell_(drow, idx.paidToDate)),
    remainingBalance: swTrim_(swCell_(drow, idx.remainingBalance)),
    lastPaymentDate: swTrim_(swCell_(drow, idx.lastPaymentDate)),
    lastPaymentDateRaw: swCell_(vrow, idx.lastPaymentDate),
    orderDate: swTrim_(swCell_(drow, idx.orderDate)),
    orderDateRaw: swCell_(vrow, idx.orderDate),
    updatedAt: swTrim_(swCell_(drow, idx.updatedAt)),
    updatedAtRaw: swCell_(vrow, idx.updatedAt),
    deadline3dMoves: swTrim_(swCell_(drow, idx.deadline3dMoves)),
    productionDeadlineMoves: swTrim_(swCell_(drow, idx.productionDeadlineMoves))
  };
  rec.root = rec.root || rec.appt;
  rec.statusNorm = swNorm_(rec.status);
  rec.activeNorm = swNorm_(rec.active);
  return rec;
}
