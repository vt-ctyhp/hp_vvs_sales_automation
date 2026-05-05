/**
 * Sales workflow email/password auth for the HtmlService dashboard.
 *
 * _SalesWorkflowUsers stores salted password hashes and role names. Session
 * tokens live in CacheService so raw passwords never land in workflow config.
 */

var SW_AUTH_SESSION_SECONDS = 6 * 60 * 60;
var SW_AUTH_USER_CACHE_SECONDS = 5 * 60;

function sw_login(email, password, options) {
  options = options || {};
  var ss = swSpreadsheet_();
  email = swNormEmail_(email);
  password = String(password || '');
  if (!email || !password) throw new Error('Email and password are required.');

  var row = swAuthFindUserRow_(ss, email);
  if (!row || !swTruthy_(row['Active?'] || '')) throw new Error('Login is not active for this email.');
  if (!row['Password Salt'] || !row['Password Hash']) throw new Error('Password is not set for this email.');
  var expected = swAuthHash_(password, row['Password Salt']);
  if (expected !== row['Password Hash']) throw new Error('Email or password is incorrect.');

  var user = swAuthUserFromRow_(row);
  var token = swAuthNewToken_();
  CacheService.getScriptCache().put(swAuthCacheKey_(token), swStringify_({
    email: user.email,
    name: user.name,
    roles: user.roles,
    user: user,
    issuedAt: swIso_(new Date())
  }), SW_AUTH_SESSION_SECONDS);
  swAuthCacheApiUser_(ss, user);
  var out = {
    ok: true,
    token: token,
    user: user,
    expiresInSeconds: SW_AUTH_SESSION_SECONDS
  };
  if (options.includeBootstrap && typeof swBuildBootstrapResponse_ === 'function') {
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var mark = typeof swStepTimer_ === 'function'
      ? swStepTimer_('sw_loginBootstrap')
      : function () {};
    mark('requiredSheets');
    out.bootstrap = swBuildBootstrapResponse_(ss, user, mark);
  }
  return out;
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
  out.password = password;
  out.generatedPassword = generated;
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

  var sessionUser = swAuthUserFromSession_(session);
  if (sessionUser) return sessionUser;

  var apiUser = swAuthCachedApiUser_(ss, session.email);
  if (apiUser) return apiUser;

  var row = swAuthFindUserRowReadOnly_(ss, session.email);
  if (!row || !swTruthy_(row['Active?'] || '')) throw new Error('Login is no longer active.');
  apiUser = swAuthUserFromRow_(row);
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

function swAuthCacheApiUser_(ss, user) {
  if (!user || !user.email) return;
  try {
    CacheService.getScriptCache().put(swAuthUserCacheKey_(ss, user.email), swStringify_(user), SW_AUTH_USER_CACHE_SECONDS);
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
    if (email) cache.remove(swAuthUserCacheKey_(ss, email));
  } catch (_) {}
}

function swAuthUserCacheKey_(ss, email) {
  return 'sw:apiUser:v1:' + ss.getId() + ':' + swNormEmail_(email);
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
  var next = {
    'Email': email,
    'Name': swTrim_(options.name) || (found && found['Name']) || email,
    'Roles': roles,
    'Active?': swTruthy_(options.active == null ? 'Y' : options.active) ? 'Y' : 'N',
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
  var row = swAuthFindUserRowReadOnly_(ss, email);
  if (!row || !swTruthy_(row['Active?'] || '')) return [];
  return swAuthRoles_(row['Roles']);
}

function swAuthActiveUserCount_(ss) {
  var sh = swEnsureSheet_(ss, SW_SHEETS.USERS, SW_AUTH_USER_HEADERS);
  return swReadSheetObjectsExpectedHeaders_(sh, SW_AUTH_USER_HEADERS).filter(function (row) {
    return swNormEmail_(row['Email']) && swTruthy_(row['Active?'] || '');
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
