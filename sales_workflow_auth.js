/**
 * Sales workflow email/password auth for the HtmlService dashboard.
 *
 * _SalesWorkflowUsers stores salted password hashes and role names. Session
 * tokens live in CacheService so raw passwords never land in workflow config.
 */

var SW_AUTH_SESSION_SECONDS = 6 * 60 * 60;
var SW_AUTH_USER_CACHE_SECONDS = 5 * 60;

function sw_login(email, password, options) {
  return swTimed_('sw_login', function () {
    var mark = typeof swStepTimer_ === 'function'
      ? swStepTimer_('sw_login')
      : function () {};
    options = options || {};
    var ss = swSpreadsheet_();
    mark('spreadsheet');
    email = swNormEmail_(email);
    password = String(password || '');
    mark('normalize');
    if (!email || !password) throw new Error('Email and password are required.');

    var row = swAuthFindUserRowForLogin_(ss, email);
    mark('userLookup', { found: !!row });
    if (!row || !swWorkflowUserActive_(row)) throw new Error('Login is not active for this email.');
    if (!row['Password Salt'] || !row['Password Hash']) throw new Error('Password is not set for this email.');
    var expected = swAuthHash_(password, row['Password Salt']);
    mark('passwordHash');
    if (expected !== row['Password Hash']) throw new Error('Email or password is incorrect.');

    var user = swAuthUserFromRow_(row);
    var token = swAuthNewToken_();
    mark('token');
    CacheService.getScriptCache().put(swAuthCacheKey_(token), swStringify_({
      email: user.email,
      name: user.name,
      roles: user.roles,
      user: user,
      issuedAt: swIso_(new Date())
    }), SW_AUTH_SESSION_SECONDS);
    mark('sessionCache');
    swAuthCacheApiUser_(ss, user);
    mark('apiUserCache');
    var out = {
      ok: true,
      token: token,
      user: user,
      expiresInSeconds: SW_AUTH_SESSION_SECONDS
    };
    if (options.includeBootstrap && typeof swBuildBootstrapResponse_ === 'function') {
      swRequireWorkflowReadSheets_(ss, { templates: false });
      mark('bootstrapRequiredSheets');
      var bootstrapMark = typeof swStepTimer_ === 'function'
        ? swStepTimer_('sw_loginBootstrap')
        : function () {};
      out.bootstrap = swBuildBootstrapResponse_(ss, user, bootstrapMark);
      mark('bootstrap');
    }
    return out;
  });
}

function sw_logout(token) {
  token = swTrim_(token);
  if (token) CacheService.getScriptCache().remove(swAuthCacheKey_(token));
  return { ok: true };
}

function sw_openWorkflowUserDialog() {
  swEnsureSheet_(swSpreadsheet_(), SW_SHEETS.USERS, SW_AUTH_USER_HEADERS);
  var html = HtmlService.createHtmlOutputFromFile('dlg_sales_workflow_users')
    .setWidth(560)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sales Workflow Users');
}

function sw_adminSetWorkflowPassword(email, password, name, roles) {
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var activeUsers = swAuthActiveUserCount_(ss);
  if (activeUsers > 0) {
    var googleUser = swCurrentUser_(ss, swBuildIdentityContext_(ss, true));
    if (!googleUser.isAdmin) throw new Error('Admin access required to set workflow passwords.');
  }

  var out = swAuthSetWorkflowPassword_(ss, {
    email: email,
    password: password,
    name: name,
    roles: roles,
    active: 'Y',
    temporary: 'Y'
  });
  var rosterLink = swEnsureOrSyncRosterForWorkflowUser_(ss, out, swSystemUser_());
  out.rosterLinked = rosterLink.linked;
  out.rosterRowNumber = rosterLink.rowNumber;
  out.warnings = rosterLink.warnings || [];
  try { if (typeof swClearAssignmentOptionsMemoryCache_ === 'function') swClearAssignmentOptionsMemoryCache_(ss); } catch (_) {}
  try { CacheService.getScriptCache().remove('sw:assignmentOptions:v1:' + ss.getId()); } catch (_) {}
  return out;
}

function sw_adminListWorkflowUsers(authToken) {
  var ss = swSpreadsheet_();
  var user = swAuthUserForApi_(ss, authToken);
  if (!user.isAdmin) throw new Error('Admin access required.');
  var rows = swAuthReadPublicUserRowsCached_(ss);
  return {
    ok: true,
    user: user,
    roleOptions: swAuthRoleOptions_(),
    users: rows
  };
}

function sw_adminUpsertWorkflowUser(authToken, data) {
  if (typeof authToken === 'object' && data == null) {
    data = authToken;
    authToken = '';
  }
  data = data || {};
  var ss = swSpreadsheet_();
  swEnsureSheet_(ss, SW_SHEETS.USERS, SW_AUTH_USER_HEADERS);
  var user = swAuthUserForApi_(ss, authToken);
  if (!user.isAdmin) throw new Error('Admin access required.');

  var password = String(data.password || '');
  var generated = false;
  if (!password) {
    password = swAuthGeneratedPassword_();
    generated = true;
  }
  var out = swAuthSetWorkflowPassword_(ss, {
    email: data.email,
    password: password,
    name: data.name,
    roles: data.roles,
    active: data.active == null ? 'Y' : data.active,
    temporary: generated ? 'Y' : (data.temporary || 'N'),
    notes: data.notes
  });
  var rosterLink = swEnsureOrSyncRosterForWorkflowUser_(ss, out, user);
  out.password = password;
  out.generatedPassword = generated;
  out.rosterLinked = rosterLink.linked;
  out.rosterRowNumber = rosterLink.rowNumber;
  out.warnings = rosterLink.warnings || [];
  try { if (typeof swClearAssignmentOptionsMemoryCache_ === 'function') swClearAssignmentOptionsMemoryCache_(ss); } catch (_) {}
  try { CacheService.getScriptCache().remove('sw:assignmentOptions:v1:' + ss.getId()); } catch (_) {}
  return out;
}

function sw_oneTimeGrantVtAdminAccess() {
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var email = 'vt@ctyhp.us';
  var existing = swAuthFindUserRow_(ss, email);
  if (existing && swTruthy_(existing['Active?'] || '') && existing['Password Salt'] && existing['Password Hash']) {
    Logger.log('SW_BOOTSTRAP_ADMIN_ALREADY_SET ' + JSON.stringify({
      ok: true,
      email: email,
      roles: existing['Roles'] || '',
      message: 'Admin login already exists; password was not reset.'
    }));
    return {
      ok: true,
      email: email,
      roles: existing['Roles'] || '',
      passwordCreated: false,
      message: 'Admin login already exists; password was not reset.'
    };
  }

  var password = swAuthGeneratedPassword_();
  var out = sw_adminSetWorkflowPassword(email, password, 'VT', 'Admin');
  Logger.log('SW_BOOTSTRAP_ADMIN_CREATED ' + JSON.stringify({
    ok: true,
    email: email,
    roles: out.roles,
    password: password,
    note: 'Use this password for dashboard login. Store it safely; raw passwords are not saved in the sheet.'
  }));
  return {
    ok: true,
    email: email,
    roles: out.roles,
    password: password,
    passwordCreated: true
  };
}

function swAuthUserForApi_(ss, token, ctx) {
  token = swTrim_(token);
  if (token) return swCurrentUserFromAuthToken_(ss, token);
  return swCurrentUser_(ss, ctx || swBuildIdentityContext_(ss, true));
}

function swCurrentUserFromAuthToken_(ss, token) {
  token = swTrim_(token);
  if (!token) throw new Error('Login required.');
  var cached = CacheService.getScriptCache().get(swAuthCacheKey_(token));
  if (!cached) throw new Error('Session expired. Please sign in again.');
  var session = swParseJson_(cached, null);
  if (!session || !session.email) throw new Error('Session expired. Please sign in again.');

  var row = swAuthFindUserRowReadOnly_(ss, session.email);
  if (!row || !swWorkflowUserActive_(row)) throw new Error('Login is no longer active.');
  var apiUser = swAuthUserFromRow_(row);
  swAuthCacheApiUser_(ss, apiUser);
  return apiUser;
}

function swAuthUserFromSession_(session) {
  session = session || {};
  if (session.user && session.user.email && Array.isArray(session.user.roles)) return session.user;
  var email = swNormEmail_(session.email || '');
  if (!email) return null;
  var roles = Array.isArray(session.roles) ? session.roles : swAuthRoles_(session.roles || '');
  return {
    email: email,
    name: swTrim_(session.name || '') || email,
    roles: roles,
    isAdmin: swAuthHasRole_(roles, 'Admin'),
    isJoc: swAuthHasRole_(roles, 'JOC'),
    isRep: swAuthHasRole_(roles, SW_OWNER_ROLES.SALES_REP),
    isDiamondOrderAdmin: swAuthHasRole_(roles, SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN),
    isDiamondOrderAssistant: swAuthHasRole_(roles, SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT)
  };
}

function swAuthUserFromRow_(row) {
  row = row || {};
  var roles = swAuthRoles_(row['Roles']);
  return {
    email: swNormEmail_(row['Email']),
    name: swTrim_(row['Name']) || swNormEmail_(row['Email']),
    roles: roles,
    isAdmin: swAuthHasRole_(roles, 'Admin'),
    isJoc: swAuthHasRole_(roles, 'JOC'),
    isRep: swAuthHasRole_(roles, SW_OWNER_ROLES.SALES_REP),
    isDiamondOrderAdmin: swAuthHasRole_(roles, SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN),
    isDiamondOrderAssistant: swAuthHasRole_(roles, SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT)
  };
}

function swAuthPublicUserRow_(row) {
  return {
    email: swNormEmail_(row['Email']),
    name: swTrim_(row['Name']),
    roles: row['Roles'] || '',
    active: row['Active?'] || '',
    temporaryPassword: row['Temporary Password?'] || '',
    passwordSet: !!row['Password Hash'],
    lastLoginAt: row['Last Login At'] || '',
    notes: row['Notes'] || ''
  };
}

function swAuthCachedApiUser_(ss, email) {
  try {
    var cached = CacheService.getScriptCache().get(swAuthUserCacheKey_(ss, email));
    var user = cached ? swParseJson_(cached, null) : null;
    if (user && user.email && Array.isArray(user.roles)) return user;
  } catch (_) {}
  return null;
}

function swAuthCachedLoginRow_(ss, email) {
  email = swNormEmail_(email);
  if (!email) return null;
  try {
    var cached = CacheService.getScriptCache().get(swAuthLoginRowCacheKey_(ss, email));
    var row = cached ? swParseJson_(cached, null) : null;
    if (row && swNormEmail_(row['Email']) === email) return row;
  } catch (_) {}
  return null;
}

function swAuthCacheApiUser_(ss, user) {
  if (!user || !user.email) return;
  try {
    CacheService.getScriptCache().put(swAuthUserCacheKey_(ss, user.email), swStringify_(user), SW_AUTH_USER_CACHE_SECONDS);
  } catch (_) {}
}

function swAuthCacheLoginRow_(ss, row) {
  row = row || {};
  var email = swNormEmail_(row['Email']);
  if (!email) return;
  try {
    CacheService.getScriptCache().put(swAuthLoginRowCacheKey_(ss, email), swStringify_(row), SW_AUTH_USER_CACHE_SECONDS);
  } catch (_) {}
}

function swAuthReadPublicUserRowsCached_(ss) {
  var key = swAuthUserListCacheKey_(ss);
  try {
    var cached = CacheService.getScriptCache().get(key);
    var rows = cached ? swParseJson_(cached, null) : null;
    if (Array.isArray(rows)) return rows;
  } catch (_) {}

  return swAuthCachePublicUserRowsFromAuthRows_(ss, swAuthReadUserRows_(ss, true));
}

function swAuthPublicUserForEmailCached_(ss, email) {
  email = swNormEmail_(email);
  if (!email) return null;
  var rows = swAuthReadPublicUserRowsCached_(ss);
  for (var i = 0; i < rows.length; i++) {
    if (swNormEmail_(rows[i].email) === email) return rows[i];
  }
  return null;
}

function swAuthCachePublicUserRowsFromAuthRows_(ss, rows) {
  var out = (rows || []).map(function (row) {
    return swAuthPublicUserRow_(row);
  });
  swAuthPutPublicUserRowsCache_(ss, out);
  return out;
}

function swAuthPutPublicUserRowsCache_(ss, rows) {
  try {
    var key = swAuthUserListCacheKey_(ss);
    var payload = swStringify_(rows || []);
    if (payload.length < 90000) CacheService.getScriptCache().put(key, payload, SW_AUTH_USER_CACHE_SECONDS);
  } catch (_) {}
}

function swAuthClearUserCaches_(ss, email) {
  try {
    var cache = CacheService.getScriptCache();
    cache.remove(swAuthUserListCacheKey_(ss));
    email = swNormEmail_(email);
    if (email) {
      cache.remove(swAuthUserCacheKey_(ss, email));
      cache.remove(swAuthLoginRowCacheKey_(ss, email));
    }
  } catch (_) {}
}

function swAuthUserCacheKey_(ss, email) {
  return 'sw:apiUser:v1:' + ss.getId() + ':' + swNormEmail_(email);
}

function swAuthLoginRowCacheKey_(ss, email) {
  return 'sw:loginRow:v1:' + ss.getId() + ':' + swNormEmail_(email);
}

function swAuthUserListCacheKey_(ss) {
  return 'sw:userList:v1:' + ss.getId();
}

function swAuthSetWorkflowPassword_(ss, options) {
  options = options || {};
  var email = swNormEmail_(options.email);
  var password = String(options.password || '');
  if (!email) throw new Error('Email is required.');
  if (password.length < 8) throw new Error('Password must be at least 8 characters.');

  var sh = swEnsureSheet_(ss, SW_SHEETS.USERS, SW_AUTH_USER_HEADERS);
  var found = swAuthFindUserRow_(ss, email);
  var salt = swAuthNewSalt_();
  var hash = swAuthHash_(password, salt);
  var roles = swAuthRolesForWrite_(options.roles || (found && found['Roles']) || SW_OWNER_ROLES.SALES_REP);
  var active = swTruthy_(options.active == null ? 'Y' : options.active);
  var name = swTrim_(options.name) || (found && found['Name']) || email;
  swValidateWorkflowUserIdentityForWrite_(ss, {
    email: email,
    name: name,
    roles: roles,
    active: active,
    rowNumber: found && found.__rowNumber
  });
  var next = {
    'Email': email,
    'Name': name,
    'Roles': roles,
    'Active?': active ? 'Y' : 'N',
    'Password Salt': salt,
    'Password Hash': hash,
    'Temporary Password?': swTruthy_(options.temporary || '') ? 'Y' : 'N',
    'Last Login At': found ? found['Last Login At'] : '',
    'Notes': options.notes != null ? swTrim_(options.notes) : (found ? found['Notes'] : '')
  };

  if (found && found.__rowNumber) {
    sh.getRange(found.__rowNumber, 1, 1, SW_AUTH_USER_HEADERS.length).setValues([SW_AUTH_USER_HEADERS.map(function (h) {
      return next[h] == null ? '' : next[h];
    })]);
  } else {
    sh.appendRow(SW_AUTH_USER_HEADERS.map(function (h) {
      return next[h] == null ? '' : next[h];
    }));
  }
  swAuthClearUserCaches_(ss, email);
  return {
    ok: true,
    email: email,
    name: next['Name'],
    roles: next['Roles'],
    active: next['Active?']
  };
}

function swValidateWorkflowUserIdentityForWrite_(ss, options) {
  options = options || {};
  var email = swNormEmail_(options.email || '');
  var name = swTrim_(options.name || '');
  var active = options.active !== false;
  var roles = swAuthRoles_(options.roles || '');
  var scheduleRole = swWorkflowSchedulableRoleList_(roles);
  var rowNumber = Number(options.rowNumber) || 0;
  var rows = swAuthReadUserRows_(ss, false);
  var duplicateActiveEmail = [];
  var duplicateActiveName = [];
  rows.forEach(function (row) {
    if (!row || !row.__rowNumber || row.__rowNumber === rowNumber) return;
    var rowEmail = swNormEmail_(row['Email']);
    var rowName = swTrim_(row['Name']);
    var rowActive = swWorkflowUserActive_(row);
    var rowScheduleRole = swWorkflowSchedulableRoleList_(swAuthRoles_(row['Roles']));
    if (active && rowActive && rowEmail && rowEmail === email) {
      duplicateActiveEmail.push(rowEmail);
    }
    if (active && scheduleRole && rowActive && rowScheduleRole && name && swNorm_(rowName) === swNorm_(name)) {
      duplicateActiveName.push(rowName + (rowEmail ? ' <' + rowEmail + '>' : ''));
    }
  });
  if (duplicateActiveEmail.length) throw new Error('Another active workflow user already uses email: ' + email);
  if (duplicateActiveName.length) {
    throw new Error('Another active Client Advisor/JOC already uses name "' + name + '": ' + duplicateActiveName.join(', '));
  }
}

function swEnsureOrSyncRosterForWorkflowUser_(ss, userData, actor) {
  var warnings = [];
  userData = userData || {};
  var email = swNormEmail_(userData.email || userData.Email || '');
  var name = swTrim_(userData.name || userData.Name || '') || email;
  var roles = swAuthRoles_(userData.roles || userData.Roles || '');
  var scheduleRole = swWorkflowSchedulableRoleList_(roles);
  var activeValue = userData.active;
  if (activeValue == null || activeValue === '') activeValue = userData['Active?'];
  var active = swWorkflowUserActive_({ active: activeValue });
  if (!email) return { linked: false, rowNumber: 0, warnings: ['No workflow user email to link.'] };

  var sh = swEnsureEmployeeScheduleSheets_(ss).roster;
  var headers = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getDisplayValues()[0].map(function (h) {
    return swTrim_(h);
  });
  var H = swHeaderMapFromArray_(headers);
  var C = {
    name: swPickIndex_(H, ['Rep', 'Name', 'Team Member']) + 1,
    email: swPickIndex_(H, ['Email', 'Rep Email']) + 1,
    role: swPickIndex_(H, ['Role', 'Roles']) + 1,
    active: swPickIndex_(H, ['Active?', 'Active']) + 1,
    updatedAt: swPickIndex_(H, ['Updated At']) + 1,
    updatedBy: swPickIndex_(H, ['Updated By']) + 1
  };
  var rowNumber = swFindRosterRowForWorkflowUser_(sh, email, name);
  if (!scheduleRole && !rowNumber) {
    return { linked: false, rowNumber: 0, warnings: warnings };
  }
  if (!rowNumber) {
    rowNumber = sh.getLastRow() + 1;
    var defaults = swDefaultEmployeeScheduleDays_();
    var newValuesByHeader = {
      rep: name,
      name: name,
      teammember: name,
      email: email,
      repemail: email,
      role: scheduleRole,
      roles: scheduleRole,
      active: active && scheduleRole ? 'Y' : 'N',
      mon: defaults.Mon ? 'Y' : 'N',
      tue: defaults.Tue ? 'Y' : 'N',
      wed: defaults.Wed ? 'Y' : 'N',
      thu: defaults.Thu ? 'Y' : 'N',
      fri: defaults.Fri ? 'Y' : 'N',
      sat: defaults.Sat ? 'Y' : 'N',
      sun: defaults.Sun ? 'Y' : 'N',
      defaultjoc: '',
      linkedjoc: '',
      jocpartner: '',
      assistedcoverageenabled: 'Y',
      coverageenabled: 'Y',
      assistedcoveragepartner: '',
      coveragepartner: '',
      labdiamond: 'None',
      naturaldiamond: 'None',
      generalappointment: 'None',
      skillnotes: '',
      updatedat: swIso_(new Date()),
      updatedby: actor ? (actor.name || actor.email || '') : ''
    };
    sh.getRange(rowNumber, 1, 1, headers.length).setValues([headers.map(function (header) {
      var key = swHeaderKey_(header);
      return newValuesByHeader[key] == null ? '' : newValuesByHeader[key];
    })]);
    return { linked: !!scheduleRole, rowNumber: rowNumber, warnings: warnings };
  }

  if (C.name > 0) sh.getRange(rowNumber, C.name).setValue(name);
  if (C.email > 0) sh.getRange(rowNumber, C.email).setValue(email);
  if (C.role > 0) sh.getRange(rowNumber, C.role).setValue(scheduleRole);
  if (C.active > 0) sh.getRange(rowNumber, C.active).setValue(active && scheduleRole ? 'Y' : 'N');
  if (C.updatedAt > 0) sh.getRange(rowNumber, C.updatedAt).setValue(swIso_(new Date()));
  if (C.updatedBy > 0) sh.getRange(rowNumber, C.updatedBy).setValue(actor ? (actor.name || actor.email || '') : '');
  if (!scheduleRole) warnings.push('Roster row was marked inactive because this user no longer has Client Advisor or JOC access.');
  return { linked: !!scheduleRole, rowNumber: rowNumber, warnings: warnings };
}

function swFindRosterRowForWorkflowUser_(sh, email, name) {
  if (!sh || sh.getLastRow() < 2) return 0;
  email = swNormEmail_(email);
  var targetName = swNorm_(name);
  var values = sh.getDataRange().getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var nameCol = swPickIndex_(H, ['Rep', 'Name', 'Team Member']);
  var emailCol = swPickIndex_(H, ['Email', 'Rep Email']);
  var nameMatches = [];
  for (var i = 1; i < values.length; i++) {
    if (emailCol >= 0 && email && swNormEmail_(values[i][emailCol]) === email) return i + 1;
    if (nameCol >= 0 && targetName && swNorm_(values[i][nameCol]) === targetName) nameMatches.push(i + 1);
  }
  if (nameMatches.length > 1) throw new Error('Multiple roster rows match "' + name + '". Clean up duplicate roster names before saving this user.');
  return nameMatches[0] || 0;
}

function swNormalizeWorkflowUserRoleLabels_(ss) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.USERS, SW_AUTH_USER_HEADERS);
  var rows = swAuthReadUserRows_(ss, false);
  var roleCol = SW_AUTH_USER_HEADERS.indexOf('Roles') + 1;
  if (roleCol <= 0) return 0;
  var updated = 0;
  rows.forEach(function (row) {
    if (!row.__rowNumber) return;
    var current = swTrim_(row['Roles']);
    var next = swAuthRolesForWrite_(current);
    if (current === next) return;
    sh.getRange(row.__rowNumber, roleCol).setValue(next);
    updated++;
  });
  return updated;
}

function swAuthRoleOptions_() {
  return [
    { value: SW_OWNER_ROLES.SALES_REP, label: 'Client Advisor', description: 'Can see tasks assigned to their email/name.' },
    { value: 'JOC', label: 'JOC', description: 'Can see assigned JOC work and claim JOC coverage tasks.' },
    { value: SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN, label: 'Diamond Order Admin', description: 'Can complete diamond order, delivery, and bulk return tasks.' },
    { value: SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT, label: 'Diamond Order Assistant', description: 'Can complete tracking and return tasks.' },
    { value: 'Admin', label: 'Admin', description: 'Can see all tasks and manage workflow users.' }
  ];
}

function swAuthRolesForWrite_(roles) {
  var allowed = {};
  swAuthRoleOptions_().forEach(function (role) {
    allowed[swAuthRoleKey_(role.value)] = role.value;
  });
  var out = [];
  swAuthRoles_(Array.isArray(roles) ? roles.join(',') : roles).forEach(function (role) {
    var canonical = allowed[swAuthRoleKey_(swAuthCanonicalRole_(role))];
    if (canonical && out.indexOf(canonical) < 0) out.push(canonical);
  });
  if (!out.length) out.push(SW_OWNER_ROLES.SALES_REP);
  return out.join(',');
}

function swAuthRoles_(value) {
  var out = [];
  String(value || '').split(/[,\n;]/).forEach(function (role) {
    role = swAuthCanonicalRole_(role);
    if (role) out.push(role);
  });
  return out;
}

function swAuthHasRole_(roles, role) {
  var target = swAuthRoleKey_(swAuthCanonicalRole_(role));
  return (roles || []).some(function (r) {
    return swAuthRoleKey_(swAuthCanonicalRole_(r)) === target;
  });
}

function swAuthCanonicalRole_(role) {
  role = swTrim_(role);
  if (!role) return '';
  var key = swAuthRoleKey_(role);
  var aliases = {
    advisor: SW_OWNER_ROLES.SALES_REP,
    clientadvisor: SW_OWNER_ROLES.SALES_REP,
    clientadvisors: SW_OWNER_ROLES.SALES_REP,
    sales: SW_OWNER_ROLES.SALES_REP,
    salesrep: SW_OWNER_ROLES.SALES_REP,
    rep: SW_OWNER_ROLES.SALES_REP,
    joc: 'JOC',
    admin: 'Admin',
    administrator: 'Admin',
    diamondadmin: SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN,
    diamondorderadmin: SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN,
    diamondordersadmin: SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN,
    diamondassistant: SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT,
    diamondorderassistant: SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT,
    diamondordersassistant: SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT
  };
  if (aliases[key]) return aliases[key];
  for (var i = 0; i < swAuthRoleOptions_().length; i++) {
    var option = swAuthRoleOptions_()[i];
    if (swAuthRoleKey_(option.value) === key || swAuthRoleKey_(option.label) === key) return option.value;
  }
  return role;
}

function swAuthRoleKey_(role) {
  return swWorkflowRoleKey_(role);
}

function swAuthFindUserRowForLogin_(ss, email) {
  email = swNormEmail_(email);
  if (!email) return null;
  var cached = swAuthCachedLoginRow_(ss, email);
  if (cached) return cached;
  var row = swAuthFindUserRowReadOnly_(ss, email) || swAuthFindUserRow_(ss, email);
  if (row) swAuthCacheLoginRow_(ss, row);
  return row;
}

function swAuthFindUserRow_(ss, email) {
  email = swNormEmail_(email);
  if (!email) return null;
  var sh = swEnsureSheet_(ss, SW_SHEETS.USERS, SW_AUTH_USER_HEADERS);
  var rows = swAuthReadUserRows_(ss, false);
  swAuthCachePublicUserRowsFromAuthRows_(ss, rows);
  for (var i = 0; i < rows.length; i++) {
    if (swNormEmail_(rows[i]['Email']) === email) return rows[i];
  }
  return null;
}

function swAuthFindUserRowReadOnly_(ss, email) {
  email = swNormEmail_(email);
  if (!email) return null;
  var rows = swAuthReadUserRows_(ss, true);
  swAuthCachePublicUserRowsFromAuthRows_(ss, rows);
  for (var i = 0; i < rows.length; i++) {
    if (swNormEmail_(rows[i]['Email']) === email) return rows[i];
  }
  return null;
}

function swAuthReadUserRows_(ss, readOnly) {
  var sh = readOnly ? ss.getSheetByName(SW_SHEETS.USERS) : swEnsureSheet_(ss, SW_SHEETS.USERS, SW_AUTH_USER_HEADERS);
  if (!sh) return [];
  return swReadSheetObjectsExpectedHeaders_(sh, SW_AUTH_USER_HEADERS);
}

function swAuthRolesForEmail_(ss, email) {
  var row = swAuthPublicUserForEmailCached_(ss, email);
  if (!row || !swWorkflowUserActive_(row)) return [];
  return swAuthRoles_(row.roles);
}

function swAuthActiveUserCount_(ss) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.USERS, SW_AUTH_USER_HEADERS);
  return swReadSheetObjectsExpectedHeaders_(sh, SW_AUTH_USER_HEADERS).filter(function (row) {
    return swNormEmail_(row['Email']) && swWorkflowUserActive_(row);
  }).length;
}

function swAuthWriteLastLogin_(ss, email) {
  var row = swAuthFindUserRow_(ss, email);
  if (!row || !row.__rowNumber) return;
  var sh = swEnsureSheet_(ss, SW_SHEETS.USERS, SW_AUTH_USER_HEADERS);
  var col = SW_AUTH_USER_HEADERS.indexOf('Last Login At') + 1;
  if (col > 0) sh.getRange(row.__rowNumber, col).setValue(swIso_(new Date()));
}

function swAuthHash_(password, salt) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt || '') + '|' + String(password || ''));
  return Utilities.base64Encode(bytes);
}

function swAuthNewSalt_() {
  return swAuthRandom_('salt');
}

function swAuthNewToken_() {
  return swAuthRandom_('sw');
}

function swAuthGeneratedPassword_() {
  return swAuthRandom_('pw').replace(/[^A-Za-z0-9]/g, '').slice(0, 18);
}

function swAuthRandom_(prefix) {
  var uuid = Utilities.getUuid ? Utilities.getUuid() : String(new Date().getTime()) + Math.random();
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, prefix + '|' + uuid + '|' + new Date().getTime());
  return prefix + '_' + Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, '');
}

function swAuthCacheKey_(token) {
  return 'sales_workflow_session_' + swTrim_(token);
}
