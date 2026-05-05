/**
 * Sales workflow identity repository: current user, permissions, config, people, and templates.
 */

function swCurrentUser_(ss, ctx) {
  var email = '';
  try { email = swNormEmail_(Session.getActiveUser().getEmail()); } catch (_) {}
  ctx = ctx || {};
  var config = ctx.config || swReadConfig_(ss);
  var assistedRoster = ctx.assistedRoster || swReadAssistedRoster_(ss);
  var admins = ctx.admins || swReadAdminsFromConfig_(config);
  var peopleIndex = ctx.peopleIndex || swReadPeopleIndex_(ss, config);
  var authRoles = email ? swAuthRolesForEmail_(ss, email) : [];
  var name = email ? (peopleIndex.nameByEmail[email] || '') : '';
  if (!name && email && !ctx.peopleIndex) name = swLookupNameByEmail_(ss, email);
  if (!name && email) name = email;
  var isAdmin = admins.length === 0 || admins.indexOf(email) >= 0 || swUserHasConfigRole_(config, email, 'Admin') || swAuthHasRole_(authRoles, 'Admin');
  var isJoc = swUserMatchesRoster_(assistedRoster, name, email) || swUserHasConfigRole_(config, email, 'JOC') || swAuthHasRole_(authRoles, 'JOC');
  return {
    email: email,
    name: name,
    isAdmin: isAdmin,
    isJoc: isJoc,
    isRep: swAuthHasRole_(authRoles, SW_OWNER_ROLES.SALES_REP),
    isDiamondOrderAdmin: swUserHasConfigRole_(config, email, SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN) || swAuthHasRole_(authRoles, SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN),
    isDiamondOrderAssistant: swUserHasConfigRole_(config, email, SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT) || swAuthHasRole_(authRoles, SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT)
  };
}

function swCurrentUserConfigOnly_(ss, readOnly) {
  var email = '';
  try { email = swNormEmail_(Session.getActiveUser().getEmail()); } catch (_) {}
  var config = swReadConfig_(ss, readOnly);
  var admins = swReadAdminsFromConfig_(config);
  var authUser = email && typeof swAuthPublicUserForEmailCached_ === 'function'
    ? swAuthPublicUserForEmailCached_(ss, email)
    : null;
  var authRoles = authUser && swWorkflowUserActive_(authUser) ? swAuthRoles_(authUser.roles) : (email ? swAuthRolesForEmail_(ss, email) : []);
  var name = '';
  for (var i = 0; i < config.length; i++) {
    if (email && swNormEmail_(config[i]['Email']) === email) {
      name = swTrim_(config[i]['Name'] || config[i]['Key']);
      break;
    }
  }
  if (!name && authUser && authUser.name) name = authUser.name;
  if (!name && email) name = email;
  var isAdmin = admins.length === 0 || admins.indexOf(email) >= 0 || swUserHasConfigRole_(config, email, 'Admin') || swAuthHasRole_(authRoles, 'Admin');
  var isJoc = swUserHasConfigRole_(config, email, 'JOC') || swAuthHasRole_(authRoles, 'JOC');
  return {
    email: email,
    name: name,
    isAdmin: isAdmin,
    isJoc: isJoc,
    isRep: swAuthHasRole_(authRoles, SW_OWNER_ROLES.SALES_REP),
    isDiamondOrderAdmin: swUserHasConfigRole_(config, email, SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN) || swAuthHasRole_(authRoles, SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN),
    isDiamondOrderAssistant: swUserHasConfigRole_(config, email, SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT) || swAuthHasRole_(authRoles, SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT)
  };
}

function swCurrentUserForTaskListView_(ss, view, readOnly) {
  view = view || 'mine';
  if (view === 'admin' || view === 'coverage') {
    var user = swCurrentUserConfigOnly_(ss, readOnly);
    if (view === 'admin' && user.isAdmin) return user;
    if (view === 'coverage' && (user.isAdmin || user.isJoc)) return user;
  }
  return swCurrentUser_(ss, swBuildIdentityContext_(ss, readOnly));
}

function swBuildBootstrapUser_(ss, readOnly) {
  var user = swCurrentUserConfigOnly_(ss, readOnly);
  if (user.isAdmin) return { user: user, lightweight: true };
  return {
    user: swCurrentUser_(ss, swBuildIdentityContext_(ss, readOnly)),
    lightweight: false
  };
}

function swSystemUser_() {
  return { name: 'System', email: '', isAdmin: true, isJoc: false };
}

function swTaskOwnedByUser_(task, user) {
  if (swTaskRoleOwnedByUser_(task, user)) return true;
  if (!swTaskNamedOwnerApplies_(task, user)) return false;
  return swTaskExplicitlyOwnedByUser_(task, user);
}

function swTaskExplicitlyOwnedByUser_(task, user) {
  if (!task || !user) return false;
  var email = swNormEmail_(user.email);
  if (email && swNormEmail_(task.currentOwnerEmail) === email) return true;
  if (swNorm_(user.name) && swNorm_(task.currentOwner) === swNorm_(user.name)) return true;
  return false;
}

function swTaskRoleOwnedByUser_(task, user) {
  if (!task || !user) return false;
  var role = swNorm_(task.ownerRole);
  if (role === swNorm_(SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN)) return !!user.isDiamondOrderAdmin;
  if (role === swNorm_(SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT)) {
    return !!user.isDiamondOrderAssistant || (!!user.isDiamondOrderAdmin && task.taskType === SW_TASKS.DIAMOND_RETURN);
  }
  return false;
}

function swTaskNamedOwnerApplies_(task, user) {
  if (!task || !user) return false;
  var role = swNorm_(task.ownerRole);
  if (role === swNorm_(SW_OWNER_ROLES.DIAMOND_ORDER_ADMIN) ||
      role === swNorm_(SW_OWNER_ROLES.DIAMOND_ORDER_ASSISTANT)) {
    return false;
  }
  if (role === swNorm_(SW_OWNER_ROLES.JOC)) return !!user.isJoc || !!user.isAdmin;
  if (swWorkflowRoleMatches_(task.ownerRole, SW_OWNER_ROLES.SALES_REP)) {
    return !!user.isRep || !!user.isAdmin || swTaskExplicitlyOwnedByUser_(task, user);
  }
  return !!user.isRep || !!user.isJoc || !!user.isAdmin || swTaskExplicitlyOwnedByUser_(task, user);
}

function swCanViewTask_(task, user) {
  if (user.isAdmin) return true;
  if (swJocCanUseCleanupCampaignTask_(task, user)) return true;
  if (swTaskOwnedByUser_(task, user)) return true;
  if (swCanClaimTask_(task, user)) return true;
  return false;
}

function swCanActOnTask_(task, user) {
  if (user.isAdmin) return true;
  if (swJocCanUseCleanupCampaignTask_(task, user)) return true;
  return swTaskOwnedByUser_(task, user);
}

function swJocCanUseCleanupCampaignTask_(task, user) {
  return !!(user && user.isJoc) &&
    swIsCleanupCampaignTask_(task) &&
    swWorkflowRoleMatches_(task.ownerRole, SW_OWNER_ROLES.JOC);
}

function swIsCleanupCampaignTask_(task) {
  return !!task &&
    typeof swIsDataCleanupTaskType_ === 'function' &&
    swIsDataCleanupTaskType_(task.taskType) &&
    swNorm_(task.lifecycleStage) === swNorm_('Cleanup Campaign');
}

function swCanClaimTask_(task, user) {
  if (!swTaskPendingLike_(task, new Date().getTime())) return false;
  if (!(user.isJoc || user.isAdmin)) return false;
  if (task.ownerRole !== 'JOC') return false;
  if (swTaskOwnedByUser_(task, user)) return false;
  return swNorm_(task.currentOwner) === swNorm_('JOC Coverage');
}

function swReadConfig_(ss, readOnly) {
  if (readOnly) {
    var cached = swReadConfigCache_(ss);
    if (cached) return cached;
  }
  var sh = readOnly
    ? swGetRequiredSheet_(ss, SW_SHEETS.CONFIG)
    : swEnsureSheet_(ss, SW_SHEETS.CONFIG, SW_CONFIG_HEADERS);
  var rows = swReadSheetObjectsExpectedHeaders_(sh, SW_CONFIG_HEADERS);
  if (readOnly) swPutConfigCache_(ss, rows);
  return rows;
}

function swReadConfigCache_(ss) {
  try {
    var cached = CacheService.getScriptCache().get(swConfigCacheKey_(ss));
    var rows = cached ? swParseJson_(cached, null) : null;
    if (Array.isArray(rows)) return rows;
  } catch (_) {}
  return null;
}

function swPutConfigCache_(ss, rows) {
  try {
    var payload = swStringify_(rows || []);
    if (payload.length < 90000) CacheService.getScriptCache().put(swConfigCacheKey_(ss), payload, 10 * 60);
  } catch (_) {}
}

function swClearConfigCache_(ss) {
  try { CacheService.getScriptCache().remove(swConfigCacheKey_(ss)); } catch (_) {}
}

function swConfigCacheKey_(ss) {
  return 'sw:config:v1:' + ss.getId();
}

function swReadAdmins_(ss) {
  return swReadAdminsFromConfig_(swReadConfig_(ss));
}

function swReadAdminsFromConfig_(config) {
  var emails = [];
  (config || []).forEach(function (r) {
    if (swNorm_(r['Section']) === 'system' && swNorm_(r['Key']) === 'adminemails') {
      String(r['Value'] || '').split(/[,\n;]/).forEach(function (e) {
        e = swNormEmail_(e);
        if (e) emails.push(e);
      });
    }
    if (swNorm_(r['Role']) === 'admin' && swTruthy_(r['Active?'] || 'Y')) {
      var em = swNormEmail_(r['Email']);
      if (em) emails.push(em);
    }
  });
  return swUnique_(emails);
}

function swReadPeopleIndex_(ss, config) {
  var out = {
    nameByEmail: {},
    emailByName: {},
    assistedRoster: []
  };
  swReadCanonicalWorkflowPeople_(ss, { activeOnly: true }).forEach(function (user) {
    if (user.email && user.name) {
      out.nameByEmail[user.email] = out.nameByEmail[user.email] || user.name;
      out.emailByName[swNorm_(user.name)] = out.emailByName[swNorm_(user.name)] || user.email;
    }
    if (swWorkflowUserHasSchedulableRole_(user, SW_OWNER_ROLES.JOC)) {
      out.assistedRoster.push({ name: user.name, email: user.email });
    }
  });

  return out;
}

function swReadAssistedRoster_(ss) {
  var out = [];
  var seen = {};
  swReadCanonicalWorkflowPeople_(ss, { schedulableOnly: true, activeOnly: true }).forEach(function (user) {
    if (!swWorkflowUserHasSchedulableRole_(user, SW_OWNER_ROLES.JOC)) return;
    var name = user.name;
    var email = user.email;
    var key = swNorm_(name) + '|' + email;
    if (!name || seen[key]) return;
    seen[key] = true;
    out.push({ name: name, email: email });
  });
  return out;
}

function swUniqueNumberList_(values) {
  var seen = {};
  var out = [];
  values.forEach(function (v) {
    v = Number(v);
    if (isNaN(v) || seen[v]) return;
    seen[v] = true;
    out.push(v);
  });
  return out;
}

function swReadTemplates_(ss, readOnly) {
  var cacheKey = '';
  if (readOnly) {
    try {
      cacheKey = 'sw:templates:v1:' + ss.getId();
      var cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) return swParseJson_(cached, {});
    } catch (_) {}
  }
  var sh = readOnly
    ? swGetRequiredSheet_(ss, SW_SHEETS.TEMPLATES)
    : swEnsureSheet_(ss, SW_SHEETS.TEMPLATES, SW_TEMPLATE_HEADERS);
  var rows = swReadSheetObjectsExpectedHeaders_(sh, SW_TEMPLATE_HEADERS);
  var out = {};
  rows.forEach(function (r) {
    var type = swTrim_(r['Task Type']);
    if (!type) return;
    out[type] = swEffectiveTemplateForTaskType_(type, {
      taskTitle: r['Task Title'] || type,
      instructions: r['Instructions'] || '',
      template: r['Template'] || '',
      attachmentLabel: r['Attachment Label'] || '',
      attachmentUrl: r['Attachment URL'] || '',
      checklistJson: r['Checklist JSON'] || '',
      primaryAction: r['Primary Action'] || 'Complete'
    });
  });
  if (cacheKey) {
    try {
      var json = JSON.stringify(out);
      if (json.length <= 90000) CacheService.getScriptCache().put(cacheKey, json, 300);
    } catch (_) {}
  }
  return out;
}

function swTemplateForType_(ss, taskType) {
  return swReadTemplates_(ss)[taskType] || swDefaultTemplate_(taskType);
}

function swDefaultTemplate_(taskType) {
  var all = swDefaultTemplates_();
  for (var i = 0; i < all.length; i++) {
    if (all[i][0] === taskType) {
      return swEffectiveTemplateForTaskType_(taskType, {
        taskTitle: all[i][1],
        instructions: all[i][2],
        template: all[i][3],
        attachmentLabel: all[i][4],
        attachmentUrl: all[i][5],
        checklistJson: all[i][6],
        primaryAction: all[i][7]
      });
    }
  }
  return { taskTitle: taskType, instructions: '', template: '', checklistJson: '', primaryAction: 'Complete' };
}
