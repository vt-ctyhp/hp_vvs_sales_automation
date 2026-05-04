/**
 * Sales workflow email/password auth for the HtmlService dashboard.
 *
 * _SalesWorkflowUsers stores salted password hashes and role names. Session
 * tokens live in CacheService so raw passwords never land in workflow config.
 */

var SW_AUTH_SESSION_SECONDS = 6 * 60 * 60;

function sw_login(email, password) {
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
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
    issuedAt: swIso_(new Date())
  }), SW_AUTH_SESSION_SECONDS);
  swAuthWriteLastLogin_(ss, email);
  return {
    ok: true,
    token: token,
    user: user,
    expiresInSeconds: SW_AUTH_SESSION_SECONDS
  };
}

function sw_logout(token) {
  token = swTrim_(token);
  if (token) CacheService.getScriptCache().remove(swAuthCacheKey_(token));
  return { ok: true };
}

function sw_adminSetWorkflowPassword(email, password, name, roles) {
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  var activeUsers = swAuthActiveUserCount_(ss);
  if (activeUsers > 0) {
    var googleUser = swCurrentUser_(ss, swBuildIdentityContext_(ss, true));
    if (!googleUser.isAdmin) throw new Error('Admin access required to set workflow passwords.');
  }

  email = swNormEmail_(email);
  password = String(password || '');
  if (!email) throw new Error('Email is required.');
  if (password.length < 8) throw new Error('Password must be at least 8 characters.');

  var sh = swEnsureSheet_(ss, SW_SHEETS.USERS, SW_AUTH_USER_HEADERS);
  var found = swAuthFindUserRow_(ss, email);
  var salt = swAuthNewSalt_();
  var hash = swAuthHash_(password, salt);
  var next = {
    'Email': email,
    'Name': swTrim_(name) || (found && found['Name']) || email,
    'Roles': swTrim_(roles) || (found && found['Roles']) || 'staff',
    'Active?': 'Y',
    'Password Salt': salt,
    'Password Hash': hash,
    'Temporary Password?': 'Y',
    'Last Login At': found ? found['Last Login At'] : '',
    'Notes': found ? found['Notes'] : ''
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
  return { ok: true, email: email, roles: next['Roles'] };
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

  var row = swAuthFindUserRow_(ss, session.email);
  if (!row || !swTruthy_(row['Active?'] || '')) throw new Error('Login is no longer active.');
  return swAuthUserFromRow_(row);
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
    isRep: true,
    isDiamondOrderAdmin: swAuthHasRole_(roles, SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN),
    isDiamondOrderAssistant: swAuthHasRole_(roles, SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT)
  };
}

function swAuthRoles_(value) {
  var out = [];
  String(value || '').split(/[,\n;]/).forEach(function (role) {
    role = swTrim_(role);
    if (role) out.push(role);
  });
  return out;
}

function swAuthHasRole_(roles, role) {
  var target = swNorm_(role);
  return (roles || []).some(function (r) { return swNorm_(r) === target; });
}

function swAuthFindUserRow_(ss, email) {
  email = swNormEmail_(email);
  if (!email) return null;
  var sh = swEnsureSheet_(ss, SW_SHEETS.USERS, SW_AUTH_USER_HEADERS);
  var rows = swReadSheetObjectsExpectedHeaders_(sh, SW_AUTH_USER_HEADERS);
  for (var i = 0; i < rows.length; i++) {
    if (swNormEmail_(rows[i]['Email']) === email) return rows[i];
  }
  return null;
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
