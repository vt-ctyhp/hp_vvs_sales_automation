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
  var name = email ? (peopleIndex.nameByEmail[email] || '') : '';
  if (!name && email && !ctx.peopleIndex) name = swLookupNameByEmail_(ss, email);
  if (!name && email) name = email;
  var isAdmin = admins.length === 0 || admins.indexOf(email) >= 0 || swUserHasConfigRole_(config, email, 'Admin');
  var isJoc = swUserMatchesRoster_(assistedRoster, name, email) || swUserHasConfigRole_(config, email, 'JOC');
  return {
    email: email,
    name: name,
    isAdmin: isAdmin,
    isJoc: isJoc,
    isRep: !!name
  };
}

function swSystemUser_() {
  return { name: 'System', email: '', isAdmin: true, isJoc: false };
}

function swTaskOwnedByUser_(task, user) {
  var email = swNormEmail_(user.email);
  if (email && swNormEmail_(task.currentOwnerEmail) === email) return true;
  if (swNorm_(user.name) && swNorm_(task.currentOwner) === swNorm_(user.name)) return true;
  return false;
}

function swCanViewTask_(task, user) {
  if (user.isAdmin) return true;
  if (swTaskOwnedByUser_(task, user)) return true;
  if (swCanClaimTask_(task, user)) return true;
  return false;
}

function swCanActOnTask_(task, user) {
  if (user.isAdmin) return true;
  return swTaskOwnedByUser_(task, user);
}

function swCanClaimTask_(task, user) {
  if (task.status !== SW_STATUSES.PENDING) return false;
  if (!(user.isJoc || user.isAdmin)) return false;
  if (task.ownerRole !== 'JOC') return false;
  if (swTaskOwnedByUser_(task, user)) return false;
  return !!task.coverageReason || swNorm_(task.currentOwner) === swNorm_('JOC Coverage');
}

function swReadConfig_(ss, readOnly) {
  var sh = readOnly
    ? swGetRequiredSheet_(ss, SW_SHEETS.CONFIG)
    : swEnsureSheet_(ss, SW_SHEETS.CONFIG, SW_CONFIG_HEADERS);
  return swReadSheetObjects_(sh);
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
  var mark = swStepTimer_('swReadPeopleIndex');
  var out = {
    nameByEmail: {},
    emailByName: {},
    assistedRoster: []
  };
  var sh = ss.getSheetByName(SW_SHEETS.DROPDOWN);
  if (sh && sh.getLastRow() >= 2) {
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    var headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
    var H = swHeaderMapFromArray_(headers);
    mark('headers', { lastRow: lastRow, lastCol: lastCol });
    var pairs = [
      [swPickIndex_(H, ['Assigned Rep']), swPickIndex_(H, ['Assigned Rep Email'])],
      [swPickIndex_(H, ['Assisted Rep']), swPickIndex_(H, ['Assisted Rep Email'])]
    ];
    var assistedNameCol = swPickIndex_(H, ['Assisted Rep', 'Assistant Rep']);
    var assistedEmailCol = swPickIndex_(H, ['Assisted Rep Email', 'Assistant Rep Email']);
    var neededCols = [];
    pairs.forEach(function (pair) {
      if (pair[0] >= 0) neededCols.push(pair[0]);
      if (pair[1] >= 0) neededCols.push(pair[1]);
    });
    if (assistedNameCol >= 0) neededCols.push(assistedNameCol);
    if (assistedEmailCol >= 0) neededCols.push(assistedEmailCol);
    var seenAssisted = {};
    if (neededCols.length) {
      var minCol = Math.min.apply(null, neededCols);
      var maxCol = Math.max.apply(null, neededCols);
      var uniqueCols = swUniqueNumberList_(neededCols);
      var readSparse = (maxCol - minCol + 1) > uniqueCols.length + 2;
      var values = readSparse ? [] : sh.getRange(2, minCol + 1, lastRow - 1, maxCol - minCol + 1).getDisplayValues();
      var sparseValues = {};
      if (readSparse) {
        uniqueCols.forEach(function (col) {
          sparseValues[col] = sh.getRange(2, col + 1, lastRow - 1, 1).getDisplayValues();
        });
      }
      mark('dropdownRead', {
        readSparse: readSparse,
        rows: lastRow - 1,
        columns: uniqueCols.length,
        minCol: minCol + 1,
        maxCol: maxCol + 1
      });
      var dropdownCell = function (row, originalCol) {
        if (originalCol < 0) return '';
        if (readSparse) return sparseValues[originalCol] && sparseValues[originalCol][row] ? sparseValues[originalCol][row][0] : '';
        return originalCol >= 0 ? row[originalCol - minCol] : '';
      };

      for (var i = 0; i < lastRow - 1; i++) {
        var row = readSparse ? i : values[i];
        pairs.forEach(function (pair) {
          var nameCol = pair[0];
          var emailCol = pair[1];
          if (nameCol < 0 || emailCol < 0) return;
          var name = swTrim_(dropdownCell(row, nameCol));
          var email = swNormEmail_(dropdownCell(row, emailCol));
          if (name && email) {
            out.emailByName[swNorm_(name)] = out.emailByName[swNorm_(name)] || email;
            out.nameByEmail[email] = out.nameByEmail[email] || name;
          }
        });

        if (assistedNameCol >= 0) {
          var assistedName = swTrim_(dropdownCell(row, assistedNameCol));
          var assistedEmail = assistedEmailCol >= 0 ? swNormEmail_(dropdownCell(row, assistedEmailCol)) : '';
          var assistedKey = swNorm_(assistedName) + '|' + assistedEmail;
          if (assistedName && !seenAssisted[assistedKey]) {
            seenAssisted[assistedKey] = true;
            out.assistedRoster.push({ name: assistedName, email: assistedEmail });
          }
        }
      }
      mark('dropdownIndex', {
        nameByEmail: Object.keys(out.nameByEmail || {}).length,
        emailByName: Object.keys(out.emailByName || {}).length,
        assistedRoster: (out.assistedRoster || []).length
      });
    }
  }

  (config || []).forEach(function (row) {
    var email = swNormEmail_(row['Email']);
    var name = swTrim_(row['Name'] || row['Key']);
    if (email && name) {
      out.nameByEmail[email] = out.nameByEmail[email] || name;
      out.emailByName[swNorm_(name)] = out.emailByName[swNorm_(name)] || email;
    }
  });
  mark('configRows', {
    configRows: (config || []).length,
    nameByEmail: Object.keys(out.nameByEmail || {}).length,
    emailByName: Object.keys(out.emailByName || {}).length,
    assistedRoster: (out.assistedRoster || []).length
  });

  return out;
}

function swReadAssistedRoster_(ss) {
  var sh = ss.getSheetByName(SW_SHEETS.DROPDOWN);
  var out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var values = sh.getDataRange().getDisplayValues();
  var headers = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(headers);
  var nameCol = swPickIndex_(H, ['Assisted Rep', 'Assistant Rep']);
  var emailCol = swPickIndex_(H, ['Assisted Rep Email', 'Assistant Rep Email']);
  if (nameCol < 0) return out;
  var seen = {};
  for (var i = 1; i < values.length; i++) {
    var name = swTrim_(values[i][nameCol]);
    var email = emailCol >= 0 ? swNormEmail_(values[i][emailCol]) : '';
    var key = swNorm_(name) + '|' + email;
    if (!name || seen[key]) continue;
    seen[key] = true;
    out.push({ name: name, email: email });
  }
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
  var sh = readOnly
    ? swGetRequiredSheet_(ss, SW_SHEETS.TEMPLATES)
    : swEnsureSheet_(ss, SW_SHEETS.TEMPLATES, SW_TEMPLATE_HEADERS);
  var rows = swReadSheetObjects_(sh);
  var out = {};
  rows.forEach(function (r) {
    var type = swTrim_(r['Task Type']);
    if (!type) return;
    out[type] = {
      taskTitle: r['Task Title'] || type,
      instructions: r['Instructions'] || '',
      template: r['Template'] || '',
      attachmentLabel: r['Attachment Label'] || '',
      attachmentUrl: r['Attachment URL'] || '',
      checklistJson: r['Checklist JSON'] || '',
      primaryAction: r['Primary Action'] || 'Complete'
    };
  });
  return out;
}

function swTemplateForType_(ss, taskType) {
  return swReadTemplates_(ss)[taskType] || swDefaultTemplate_(taskType);
}

function swDefaultTemplate_(taskType) {
  var all = swDefaultTemplates_();
  for (var i = 0; i < all.length; i++) {
    if (all[i][0] === taskType) {
      return {
        taskTitle: all[i][1],
        instructions: all[i][2],
        template: all[i][3],
        attachmentLabel: all[i][4],
        attachmentUrl: all[i][5],
        checklistJson: all[i][6],
        primaryAction: all[i][7]
      };
    }
  }
  return { taskTitle: taskType, instructions: '', template: '', checklistJson: '', primaryAction: 'Complete' };
}
