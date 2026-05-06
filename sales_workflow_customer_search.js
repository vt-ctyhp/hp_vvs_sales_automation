/**
 * Sales Workflow customer search: dashboard kanban lookup and direct ops actions.
 */

var SW_CUSTOMER_SEARCH_MAX_CARDS_PER_COLUMN = 30;
var SW_CUSTOMER_SEARCH_LOG_SCAN_ROWS = 500;
var SW_CUSTOMER_SEARCH_READ_MODEL_CACHE_SECONDS = 10 * 60;
var SW_CUSTOMER_SEARCH_DETAIL_CACHE_SECONDS = 10 * 60;
var SW_CUSTOMER_SEARCH_DETAIL_INDEX_SHARDS = 32;
var SW_CUSTOMER_SEARCH_READ_MODEL_MEMORY_CACHE_ = {};
var SW_CUSTOMER_SEARCH_INITIAL_PAYLOAD_MEMORY_CACHE_ = {};
var SW_CUSTOMER_SEARCH_DETAIL_INDEX_MEMORY_CACHE_ = {};
var SW_CUSTOMER_SEARCH_PAYMENT_HISTORY_MEMORY_CACHE_ = {};
var SW_CUSTOMER_SEARCH_RECENT_LOG_MEMORY_CACHE_ = {};

function sw_searchCustomers(authToken, query, filters) {
  return swTimed_('sw_searchCustomers', function () {
    var mark = swStepTimer_('sw_searchCustomers');
    var ss = swSpreadsheet_();
    mark('spreadsheet');
    swRequireWorkflowReadSheets_(ss, { templates: false });
    mark('requiredSheets');
    var user = swAuthUserForApi_(ss, authToken);
    mark('identity');
    swRequireCustomerSearchUser_(user);

    filters = swCustomerSearchNormalizeFilters_(filters);
    query = swTrim_(query || filters.query || '');

    var config = swReadConfig_(ss, true);
    mark('config');
    var readModelStatus = typeof swCustomerReadModelStatus_ === 'function'
      ? swCustomerReadModelStatus_(ss)
      : null;
    var initialPayload = swCustomerSearchInitialPayloadForRequest_(ss, user, query, filters, readModelStatus);
    if (initialPayload) {
      mark('rows', {
        source: 'customerSearchInitialPayloadCache',
        rows: initialPayload.sourceRows || 0,
        ageSeconds: initialPayload.ageSeconds || 0
      });
      mark('filter', {
        source: 'customerSearchInitialPayloadCache',
        rows: initialPayload.filteredRows || 0,
        query: false
      });
      mark('cards', {
        source: 'customerSearchInitialPayloadCache',
        cards: initialPayload.cards || 0,
        hiddenCards: initialPayload.hiddenCards || 0
      });
      return initialPayload.payload;
    }

    var readModel = swCustomerSearchReadModelRows_(ss, config, readModelStatus);
    if (readModel && readModel.ok) {
      mark('rows', {
        source: readModel.source || 'customerReadModel',
        rows: readModel.rows ? readModel.rows.length : 0,
        ageSeconds: readModel.ageSeconds || 0
      });
      var readModelPayload = swCustomerSearchBuildPayload_(ss, user, query, filters, readModel.rows || [], readModel.source || 'customerReadModel', mark);
      swMaybeCacheCustomerSearchInitialPayload_(ss, user, query, filters, readModel.rows || [], readModelStatus || readModel.status, readModelPayload);
      return readModelPayload;
    }

    var appointments = swReadAppointments_(ss);
    mark('rows', {
      source: 'appointments',
      rows: appointments.length,
      fallbackReason: readModel ? readModel.fallbackReason || '' : ''
    });
    return swCustomerSearchBuildPayload_(ss, user, query, filters, appointments, 'appointments', mark);
  });
}

function swCustomerSearchBuildPayload_(ss, user, query, filters, sourceRows, source, mark) {
  mark = mark || function () {};
  sourceRows = sourceRows || [];
  swCustomerSearchApplyDefaultOwnerFilters_(sourceRows, user, filters);
  var rows = swCustomerSearchFilteredRows_(sourceRows, query, filters);
  mark('filter', { source: source || '', rows: rows.length, query: !!query });

  var kanban = swCustomerSearchKanbanFromRows_(ss, rows, source);
  mark('cards', { source: source || '', cards: kanban.cards, hiddenCards: kanban.hiddenCards });

  return {
    ok: true,
    generatedAt: swIso_(new Date()),
    query: query,
    source: source || '',
    filters: swCustomerSearchPublicFilters_(filters),
    filterOptions: swCustomerSearchFilterOptions_(sourceRows),
    kanban: {
      columns: kanban.columns
    }
  };
}

function swCustomerSearchInitialPayloadForRequest_(ss, user, query, filters, status) {
  if (!swCanUseCustomerSearchInitialPayloadCache_(user, query, filters)) return null;
  if (!(status && status.fresh)) return null;
  var key = swCustomerSearchInitialPayloadCacheKey_(ss);
  var expectedVersion = typeof SW_READ_MODEL_VERSION !== 'undefined' ? SW_READ_MODEL_VERSION : '';
  var cached = null;
  try {
    var memory = SW_CUSTOMER_SEARCH_INITIAL_PAYLOAD_MEMORY_CACHE_[key];
    if (memory && memory.expiresAt > new Date().getTime() &&
        memory.modelBuiltAt === status.builtAt &&
        memory.version === expectedVersion) {
      cached = memory.payload || null;
    }
  } catch (_) {}
  if (!cached) {
    cached = swCustomerSearchInitialPayloadCacheGet_(key);
    if (!cached ||
        cached.modelBuiltAt !== status.builtAt ||
        cached.version !== expectedVersion ||
        !(cached.payload && cached.payload.ok)) {
      return null;
    }
    try {
      SW_CUSTOMER_SEARCH_INITIAL_PAYLOAD_MEMORY_CACHE_[key] = {
        expiresAt: new Date().getTime() + SW_CUSTOMER_SEARCH_READ_MODEL_CACHE_SECONDS * 1000,
        version: cached.version || '',
        modelBuiltAt: cached.modelBuiltAt || '',
        payload: cached
      };
    } catch (_) {}
  }
  cached.payload.generatedAt = swIso_(new Date());
  cached.ageSeconds = status.ageSeconds || 0;
  return cached;
}

function swCanUseCustomerSearchInitialPayloadCache_(user, query, filters) {
  filters = filters || {};
  return !!(user && user.isAdmin) &&
    !swTrim_(query) &&
    filters.activeOnly !== false &&
    !swTrim_(filters.brand) &&
    !swTrim_(filters.clientAdvisor) &&
    !swTrim_(filters.joc) &&
    !filters.defaultOwner &&
    !filters.defaultAdvisor &&
    !filters.defaultJoc;
}

function swMaybeCacheCustomerSearchInitialPayload_(ss, user, query, filters, rows, status, payload) {
  if (!swCanUseCustomerSearchInitialPayloadCache_(user, query, filters)) return null;
  if (!(status && status.fresh)) return null;
  try {
    return swCacheCustomerSearchInitialPayload_(ss, rows || [], status, payload);
  } catch (_) {}
  return null;
}

function swCacheCustomerSearchInitialPayload_(ss, rows, status, payload) {
  if (!(status && status.builtAt)) return { ok: false, reason: 'missingStatus' };
  rows = rows || [];
  if (!(payload && payload.ok)) {
    payload = swCustomerSearchBuildPayload_(
      ss,
      { isAdmin: true },
      '',
      swCustomerSearchNormalizeFilters_({ activeOnly: true }),
      rows,
      'customerReadModelCache',
      function () {}
    );
  }
  var columns = payload && payload.kanban && payload.kanban.columns ? payload.kanban.columns : [];
  var filteredRows = columns.reduce(function (sum, col) {
    return sum + Number(col.count || 0);
  }, 0);
  var cards = columns.reduce(function (sum, col) {
    return sum + ((col.cards || []).length);
  }, 0);
  var hiddenCards = columns.reduce(function (sum, col) {
    return sum + Number(col.hiddenCount || 0);
  }, 0);
  var cachePayload = {
    cachedAt: swIso_(new Date()),
    version: typeof SW_READ_MODEL_VERSION !== 'undefined' ? SW_READ_MODEL_VERSION : '',
    modelBuiltAt: status.builtAt || '',
    sourceRows: rows.length,
    filteredRows: filteredRows,
    cards: cards,
    hiddenCards: hiddenCards,
    payload: payload
  };
  var key = swCustomerSearchInitialPayloadCacheKey_(ss);
  try {
    SW_CUSTOMER_SEARCH_INITIAL_PAYLOAD_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_CUSTOMER_SEARCH_READ_MODEL_CACHE_SECONDS * 1000,
      version: cachePayload.version,
      modelBuiltAt: cachePayload.modelBuiltAt,
      payload: cachePayload
    };
  } catch (_) {}
  var result = swCustomerSearchInitialPayloadCachePut_(key, cachePayload) || {};
  return {
    ok: result.ok !== false,
    chunks: result.chunks || 0,
    bytes: result.bytes || 0,
    reason: result.reason || ''
  };
}

function swCustomerSearchKanbanFromRows_(ss, rows, source) {
  var groups = swAdminDashboardRowsByRoot_(rows);
  var aiBriefByRoot = Object.keys(groups).length && typeof swAppointmentAiBriefIndex_ === 'function'
    && source !== 'customerReadModelSheet' && source !== 'customerReadModelCache'
    ? swAppointmentAiBriefIndex_(ss)
    : {};
  var master = ss.getSheetByName(SW_SHEETS.MASTER);
  var masterGid = master ? master.getSheetId() : '';
  var columnsByKey = {};
  SW_ADMIN_DASHBOARD_COLUMNS.forEach(function (col) {
    columnsByKey[col.key] = { key: col.key, label: col.label, count: 0, cards: [], hiddenCount: 0 };
  });

  Object.keys(groups).forEach(function (root) {
    var rootRows = groups[root];
    var active = rootRows.filter(function (rec) { return swIsAppointmentActive_(rec); });
    var rec = swAdminDashboardLatestRow_(active.length ? active : rootRows);
    if (!rec) return;
    var stage = swCustomerSearchStageForRecord_(rec, rootRows);
    var column = columnsByKey[stage.key] || columnsByKey.lead;
    column.count++;
    if (column.cards.length < SW_CUSTOMER_SEARCH_MAX_CARDS_PER_COLUMN) {
      var card = swCustomerSearchCardForRecord_(ss, masterGid, root, rec, rootRows, stage, aiBriefByRoot[root]);
      column.cards.push(swCustomerSearchKanbanCardForList_(card));
    } else {
      column.hiddenCount++;
    }
  });

  return {
    columns: SW_ADMIN_DASHBOARD_COLUMNS.map(function (col) { return columnsByKey[col.key]; }),
    cards: SW_ADMIN_DASHBOARD_COLUMNS.reduce(function (sum, col) {
      return sum + ((columnsByKey[col.key] && columnsByKey[col.key].cards) || []).length;
    }, 0),
    hiddenCards: SW_ADMIN_DASHBOARD_COLUMNS.reduce(function (sum, col) {
      return sum + Number((columnsByKey[col.key] && columnsByKey[col.key].hiddenCount) || 0);
    }, 0)
  };
}

function swCustomerSearchStageForRecord_(rec, rootRows) {
  if (rec && rec.__customerReadModel && rec.stageKey) {
    return {
      key: rec.stageKey,
      label: rec.stageLabel || swCustomerSearchStageLabel_(rec.stageKey)
    };
  }
  return swAdminDashboardPipelineStage_(rec, rootRows);
}

function swCustomerSearchStageLabel_(key) {
  key = swTrim_(key);
  for (var i = 0; i < SW_ADMIN_DASHBOARD_COLUMNS.length; i++) {
    if (SW_ADMIN_DASHBOARD_COLUMNS[i].key === key) return SW_ADMIN_DASHBOARD_COLUMNS[i].label;
  }
  return key || 'Lead';
}

function swCustomerSearchCardForRecord_(ss, masterGid, root, rec, rootRows, stage, aiBrief) {
  if (rec && rec.__customerReadModel) {
    return swCustomerSearchCardFromReadModel_(ss, masterGid, rec, stage);
  }
  return swCustomerSearchCard_(ss, masterGid, root, rec, rootRows, stage, aiBrief);
}

function swCustomerSearchKanbanCardForList_(card) {
  card = card || {};
  return {
    root: card.root || '',
    appt: card.appt || '',
    customerName: card.customerName || '',
    brand: card.brand || '',
    clientAdvisor: card.clientAdvisor || '',
    joc: card.joc || '',
    nextVisit: card.nextVisit || '',
    lastVisit: card.lastVisit || '',
    salesStage: card.salesStage || '',
    conversionStatus: card.conversionStatus || '',
    customOrderStatus: card.customOrderStatus || '',
    inProductionStatus: card.inProductionStatus || '',
    so: card.so || '',
    deadline3d: card.deadline3d || '',
    waxStatus: card.waxStatus || '',
    badges: card.badges || []
  };
}

function sw_getCustomerSearchDetail(authToken, rootApptId) {
  return swTimed_('sw_getCustomerSearchDetail', function () {
    var mark = swStepTimer_('sw_getCustomerSearchDetail');
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    mark('requiredSheets');
    var user = swAuthUserForApi_(ss, authToken);
    mark('identity');
    swRequireCustomerSearchUser_(user);
    var out = swCustomerSearchDetailPayload_(ss, user, rootApptId);
    mark('payload', {
      appointments: out.appointments ? out.appointments.length : 0,
      tasks: out.tasks ? out.tasks.length : 0,
      logs: out.logs ? out.logs.length : 0
    });
    return out;
  });
}

function sw_customerSearchUpdateStatus(authToken, rootApptId, payload) {
  return swTimed_('sw_customerSearchUpdateStatus', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken);
    swRequireCustomerSearchUser_(user);
    payload = payload || {};

    var lock = LockService.getDocumentLock() || LockService.getScriptLock();
    lock.waitLock(28000);
    try {
      var target = swCustomerSearchResolveRoot_(ss, rootApptId);
      var rowNum = swCustomerSearchValidateResolvedRow_(ss, target, rootApptId);
      var result = cs_submitFromDialogForRow_(rowNum, {
        assignedRep: target.rec.assignedRep || '',
        assistedRep: target.rec.assistedRep || '',
        salesStage: swTrim_(payload.salesStage),
        convStatus: swTrim_(payload.convStatus),
        customOrder: swTrim_(payload.customOrder),
        cosAllowedEmpty: !swTrim_(payload.customOrder),
        inProduction: swTrim_(payload.inProduction),
        centerStone: swTrim_(payload.centerStone),
        nextSteps: swTrim_(payload.nextSteps),
        orderDate: swTrim_(payload.orderDate),
        deadline3d: swTrim_(payload.deadline3d),
        prodDeadline: swTrim_(payload.prodDeadline),
        wax: null,
        waxSummary: '',
        notebookLMLink: swTrim_(payload.notebookLMLink)
      }, ss);
      if (result && result.ok === false) throw new Error(result.error || 'Client status update failed.');
      swCustomerSearchLog_(ss, 'CUSTOMER_SEARCH_STATUS_UPDATE', target.rec, user, payload, result);
      swCustomerSearchInvalidateReadModel_(ss, 'Customer Search status update changed appointment data.');
    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }
    return swCustomerSearchDetailPayload_(ss, user, rootApptId);
  });
}

function sw_customerSearchUpdate3DDeadline(authToken, rootApptId, payload) {
  return swTimed_('sw_customerSearchUpdate3DDeadline', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken);
    swRequireCustomerSearchUser_(user);
    payload = payload || {};
    var deadline = swTrim_(payload.deadline3d || payload.dateIso || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(deadline)) throw new Error('Select a valid 3D deadline.');

    var lock = LockService.getDocumentLock() || LockService.getScriptLock();
    lock.waitLock(28000);
    try {
      var target = swCustomerSearchActivateRoot_(ss, rootApptId);
      var result = (typeof Deadlines !== 'undefined' && Deadlines.saveRecordDeadline)
        ? Deadlines.saveRecordDeadline({ kind: '3D', dateIso: deadline })
        : saveRecordDeadline({ kind: '3D', dateIso: deadline });
      if (result && result.ok === false) throw new Error(result.error || '3D deadline update failed.');
      swCustomerSearchLog_(ss, 'CUSTOMER_SEARCH_3D_DEADLINE', target.rec, user, { deadline3d: deadline, note: swTrim_(payload.note) }, result);
      swCustomerSearchInvalidateReadModel_(ss, 'Customer Search 3D deadline update changed appointment data.');
    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }
    return swCustomerSearchDetailPayload_(ss, user, rootApptId);
  });
}

function sw_customerSearchSubmit3DRevision(authToken, rootApptId, payload) {
  return swTimed_('sw_customerSearchSubmit3DRevision', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken);
    swRequireCustomerSearchUser_(user);
    payload = payload || {};
    var form = payload.form || payload;
    if (!swTrim_(form.DesignNotes)) throw new Error('Enter revision design notes.');

    var lock = LockService.getDocumentLock() || LockService.getScriptLock();
    lock.waitLock(28000);
    try {
      var target = swCustomerSearchActivateRoot_(ss, rootApptId);
      var result = submit3DRevision({ form: form });
      if (result && result.ok === false) throw new Error(result.error || '3D revision failed.');
      swCustomerSearchLog_(ss, 'CUSTOMER_SEARCH_3D_REVISION', target.rec, user, form, result);
      swCustomerSearchInvalidateReadModel_(ss, 'Customer Search 3D revision changed customer workflow data.');
    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }
    return swCustomerSearchDetailPayload_(ss, user, rootApptId);
  });
}

function sw_customerSearchRequestWax(authToken, rootApptId, payload) {
  return swTimed_('sw_customerSearchRequestWax', function () {
    var ss = swSpreadsheet_();
    swRequireWorkflowReadSheets_(ss, { templates: false });
    var user = swAuthUserForApi_(ss, authToken);
    swRequireCustomerSearchUser_(user);
    payload = payload || {};
    if (!swTrim_(payload.soMo)) throw new Error('Enter the SO/MO number for the wax request.');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(swTrim_(payload.neededByRep))) throw new Error('Select Needed By (Rep).');
    if (!swTrim_(payload.priority)) throw new Error('Select wax priority.');

    var lock = LockService.getDocumentLock() || LockService.getScriptLock();
    lock.waitLock(28000);
    try {
      var target = swCustomerSearchResolveRoot_(ss, rootApptId);
      var result = wax_onRequestSubmit_({
        rootApptId: target.root,
        soMo: swTrim_(payload.soMo),
        neededByRep: swTrim_(payload.neededByRep),
        priority: swTrim_(payload.priority),
        requestedBy: (user && (user.email || user.name)) || ''
      });
      if (result && result.ok === false) throw new Error(result.error || 'Wax request failed.');
      swCustomerSearchLog_(ss, 'CUSTOMER_SEARCH_WAX_REQUEST', target.rec, user, payload, result);
      swCustomerSearchInvalidateReadModel_(ss, 'Customer Search wax request changed customer workflow data.');
    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }
    return swCustomerSearchDetailPayload_(ss, user, rootApptId);
  });
}

function swRequireCustomerSearchUser_(user) {
  if (!(user && (user.isAdmin || user.isJoc || user.isRep))) {
    throw new Error('Customer Search access requires Client Advisor, JOC, or Admin role.');
  }
}

function swCustomerSearchInvalidateReadModel_(ss, reason) {
  try {
    if (typeof swMarkWorkflowReadModelsStale_ === 'function') {
      swMarkWorkflowReadModelsStale_(ss, reason || 'Customer Search changed customer data.', 'customers');
      return;
    }
    if (typeof swInvalidateCustomerSearchReadModelCache_ === 'function') {
      swInvalidateCustomerSearchReadModelCache_(ss);
    }
  } catch (_) {}
}

function swCustomerSearchNormalizeFilters_(filters) {
  filters = filters || {};
  var activeRaw = filters.activeOnly;
  var activeOnly = !(activeRaw === false || String(activeRaw || '').toLowerCase() === 'false');
  var defaultAdvisor = filters.defaultAdvisor === true || String(filters.defaultAdvisor || '').toLowerCase() === 'true';
  var defaultJoc = filters.defaultJoc === true || String(filters.defaultJoc || '').toLowerCase() === 'true';
  var defaultOwner = filters.defaultOwner === true || String(filters.defaultOwner || '').toLowerCase() === 'true';
  return {
    query: swTrim_(filters.query || ''),
    brand: swTrim_(filters.brand || ''),
    clientAdvisor: swTrim_(filters.clientAdvisor || ''),
    joc: swTrim_(filters.joc || ''),
    defaultAdvisor: defaultAdvisor,
    defaultJoc: defaultJoc,
    defaultOwner: defaultOwner || defaultAdvisor || defaultJoc,
    activeOnly: activeOnly
  };
}

function swCustomerSearchPublicFilters_(filters) {
  return {
    query: filters.query || '',
    brand: filters.brand || '',
    clientAdvisor: filters.clientAdvisor || '',
    joc: filters.joc || '',
    defaultAdvisor: false,
    defaultJoc: false,
    defaultOwner: false,
    activeOnly: filters.activeOnly !== false
  };
}

function swCustomerSearchReadModelRows_(ss, config, knownStatus) {
  if (typeof swCustomerReadModelServingEnabled_ === 'function' &&
      !swCustomerReadModelServingEnabled_(config || [])) {
    return { ok: false, fallbackReason: 'disabled' };
  }
  if (typeof swCustomerReadModelStatus_ !== 'function') {
    return { ok: false, fallbackReason: 'statusUnavailable' };
  }

  try {
    var status = knownStatus || swCustomerReadModelStatus_(ss);
    if (!status.fresh) {
      return {
        ok: false,
        fallbackReason: status.reason || 'notFresh',
        actualVersion: status.actualVersion || '',
        expectedVersion: status.expectedVersion || '',
        ageSeconds: status.ageSeconds || 0
      };
    }

    var cachedRows = swReadCachedCustomerSearchReadModelRows_(ss, status);
    if (cachedRows) {
      return {
        ok: true,
        source: 'customerReadModelCache',
        rows: cachedRows,
        ageSeconds: status.ageSeconds || 0,
        status: status
      };
    }

    var sh = ss.getSheetByName(SW_SHEETS.READ_MODEL_CUSTOMERS);
    var rows = swReadSheetObjectsExpectedHeaders_(sh, SW_CUSTOMER_READ_MODEL_HEADERS)
      .map(swCustomerSearchReadModelRecord_)
      .filter(function (rec) { return !!rec.root; });
    swCacheCustomerSearchReadModelRows_(ss, rows, status);
    return {
      ok: true,
      source: 'customerReadModelSheet',
      rows: rows,
      ageSeconds: status.ageSeconds || 0,
      status: status
    };
  } catch (err) {
    try {
      Logger.log('SW_CUSTOMER_SEARCH_READ_MODEL_FALLBACK ' + JSON.stringify({
        reason: err && err.message ? err.message : String(err)
      }));
    } catch (_) {}
    return {
      ok: false,
      fallbackReason: err && err.message ? err.message : String(err)
    };
  }
}

function swCustomerSearchReadModelRecord_(row) {
  row = row || {};
  var active = swTrim_(row['Active?']) === 'Y';
  return {
    __customerReadModel: true,
    row: Number(row['Master Row'] || 0) || '',
    root: swTrim_(row['RootApptID']),
    appt: swTrim_(row['Latest APPT_ID']),
    name: swTrim_(row['Customer Name']),
    email: swNormEmail_(row['Email']),
    phone: swNormPhone_(row['Phone']),
    brand: swTrim_(row['Brand']),
    assignedRep: swTrim_(row['Client Advisor']),
    assignedRepEmail: swNormEmail_(row['Client Advisor Email']),
    assistedRep: swTrim_(row['JOC']),
    assistedRepEmail: swNormEmail_(row['JOC Email']),
    visitDate: swTrim_(row['Latest Visit Date']),
    visitTime: swFormatAppointmentTime_(row['Latest Visit Time'] || ''),
    visitType: swTrim_(row['Latest Visit Type']),
    nextVisit: swTrim_(row['Next Visit']),
    lastVisit: swTrim_(row['Last Visit']),
    appointmentCount: Number(row['Appointment Count'] || 0) || 0,
    activeAppointmentCount: Number(row['Active Appointment Count'] || 0) || 0,
    active: active ? 'Yes' : 'No',
    activeNorm: active ? 'yes' : 'no',
    status: '',
    statusNorm: '',
    stageKey: swTrim_(row['Stage Key']) || 'lead',
    stageLabel: swTrim_(row['Stage Label']),
    salesStage: swTrim_(row['Sales Stage']),
    convStatus: swTrim_(row['Conversion Status']),
    customOrder: swTrim_(row['Custom Order Status']),
    inProduction: swTrim_(row['In Production Status']),
    centerStoneStatus: swTrim_(row['Center Stone Status']),
    so: swTrim_(row['SO#']),
    orderTotal: swTrim_(row['Order Total']),
    paidToDate: swTrim_(row['Paid-to-Date']),
    remainingBalance: swTrim_(row['Remaining Balance']),
    lastPaymentDate: swTrim_(row['Last Payment Date']),
    quotationUrl: swTrim_(row['Quotation URL']),
    clientFolder: swTrim_(row['Client Folder']),
    reportUrl: swTrim_(row['Client Status Report URL']),
    tracker3dUrl: swTrim_(row['3D Tracker URL']),
    deadline3d: swTrim_(row['3D Deadline']),
    productionDeadline: swTrim_(row['Production Deadline']),
    waxStatus: swTrim_(row['Wax Print Status']),
    waxDeadlineAdmin: swTrim_(row['Wax Deadline (Admin)']),
    dvStonesSummary: swTrim_(row['DV Stones Summary']),
    nextSteps: swTrim_(row['Next Steps']),
    updatedAt: swTrim_(row['Updated At']),
    hasAiBrief: swTrim_(row['AI Brief?']) === 'Y',
    reviewFlagCount: Number(row['Review Flag Count'] || 0) || 0,
    latestAiBriefUpdatedAt: swTrim_(row['Latest AI Brief Updated At']),
    sourceRows: swCustomerSearchParseSourceRows_(row['Source Rows JSON']),
    searchText: swTrim_(row['Search Text'])
  };
}

function swCustomerSearchParseSourceRows_(value) {
  var parsed = swParseJson_(value, []);
  if (!Array.isArray(parsed)) return [];
  var seen = {};
  var out = [];
  parsed.forEach(function (row) {
    row = Number(row);
    if (!isFinite(row) || row < 2 || seen[row]) return;
    seen[row] = true;
    out.push(row);
  });
  return out;
}

function swCustomerSearchCardFromReadModel_(ss, masterGid, rec, stage) {
  var paid = swAdminDashboardNumberOrBlank_(rec.paidToDate);
  var balance = swAdminDashboardNumberOrBlank_(rec.remainingBalance);
  var orderTotal = swAdminDashboardNumberOrBlank_(rec.orderTotal);
  var card = {
    root: rec.root || '',
    appt: rec.appt || '',
    row: rec.row || '',
    customerName: rec.name || '',
    brand: rec.brand || '',
    clientAdvisor: rec.assignedRep || '',
    joc: rec.assistedRep || '',
    visitType: rec.visitType || '',
    visitDate: rec.visitDate || '',
    visitTime: rec.visitTime || '',
    nextVisit: rec.nextVisit || '',
    lastVisit: rec.lastVisit || '',
    salesStage: rec.salesStage || '',
    conversionStatus: rec.convStatus || '',
    customOrderStatus: rec.customOrder || '',
    inProductionStatus: rec.inProduction || '',
    stageKey: stage.key,
    stageLabel: stage.label,
    so: rec.so || '',
    nextSteps: rec.nextSteps || '',
    paymentCount: 0,
    paidNet: paid === '' ? 0 : paid,
    balanceDue: balance === 0 ? 0 : (balance || ''),
    orderTotal: orderTotal === 0 ? 0 : (orderTotal || ''),
    lastPaymentDate: rec.lastPaymentDate || '',
    source: 'customerReadModel',
    remainingBalance: rec.remainingBalance || '',
    updatedAt: rec.updatedAt || '',
    clientFolder: rec.clientFolder || '',
    reportUrl: rec.reportUrl || '',
    quotationUrl: rec.quotationUrl || '',
    tracker3dUrl: rec.tracker3dUrl || '',
    masterUrl: masterGid && rec.row ? ('https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + masterGid + '&range=A' + rec.row) : '',
    email: rec.email || '',
    phone: rec.phone || '',
    deadline3d: rec.deadline3d || '',
    productionDeadline: rec.productionDeadline || '',
    waxStatus: rec.waxStatus || '',
    waxDeadlineAdmin: rec.waxDeadlineAdmin || '',
    centerStoneStatus: rec.centerStoneStatus || '',
    badges: swCustomerSearchBadges_(rec),
    hasAiBrief: !!rec.hasAiBrief,
    reviewFlagCount: Number(rec.reviewFlagCount || 0),
    latestAiBriefUpdatedAt: rec.latestAiBriefUpdatedAt || ''
  };
  if (card.hasAiBrief) {
    card.badges = card.badges || [];
    card.badges.push('AI Brief');
    if (card.reviewFlagCount) card.badges.push('Flags: ' + card.reviewFlagCount);
  }
  return card;
}

function swReadCachedCustomerSearchReadModelRows_(ss, status) {
  var key = swCustomerSearchReadModelCacheKey_(ss);
  try {
    var memory = SW_CUSTOMER_SEARCH_READ_MODEL_MEMORY_CACHE_[key];
    if (memory && memory.expiresAt > new Date().getTime() &&
        memory.modelBuiltAt === status.builtAt) {
      return memory.rows || [];
    }
  } catch (_) {}

  var payload = swTaskListCacheGet_(key);
  if (!payload || payload.modelBuiltAt !== status.builtAt || !payload.rows) return null;
  try {
    SW_CUSTOMER_SEARCH_READ_MODEL_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_CUSTOMER_SEARCH_READ_MODEL_CACHE_SECONDS * 1000,
      modelBuiltAt: payload.modelBuiltAt || '',
      rows: payload.rows || []
    };
  } catch (_) {}
  return payload.rows || [];
}

function swCacheCustomerSearchReadModelRows_(ss, rows, status) {
  var key = swCustomerSearchReadModelCacheKey_(ss);
  var payload = {
    cachedAt: swIso_(new Date()),
    modelBuiltAt: status && status.builtAt || '',
    rows: rows || []
  };
  try {
    SW_CUSTOMER_SEARCH_READ_MODEL_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_CUSTOMER_SEARCH_READ_MODEL_CACHE_SECONDS * 1000,
      modelBuiltAt: payload.modelBuiltAt,
      rows: payload.rows
    };
  } catch (_) {}
  var result = swTaskListCachePut_(key, payload);
  try { swCacheCustomerSearchInitialPayload_(ss, rows || [], status); } catch (_) {}
  return result;
}

function swInvalidateCustomerSearchReadModelCache_(ss) {
  var key = swCustomerSearchReadModelCacheKey_(ss);
  var payloadKey = swCustomerSearchInitialPayloadCacheKey_(ss);
  try { delete SW_CUSTOMER_SEARCH_READ_MODEL_MEMORY_CACHE_[key]; } catch (_) {}
  try { delete SW_CUSTOMER_SEARCH_INITIAL_PAYLOAD_MEMORY_CACHE_[payloadKey]; } catch (_) {}
  try { swTaskListCacheRemove_(key); } catch (_) {}
  try { swCustomerSearchInitialPayloadCacheRemove_(payloadKey); } catch (_) {}
  try { swInvalidateCustomerSearchDetailCaches_(ss); } catch (_) {}
}

function swCustomerSearchReadModelCacheKey_(ss) {
  return 'sw:customerSearchReadModel:v1:' + ss.getId();
}

function swCustomerSearchInitialPayloadCacheKey_(ss) {
  return 'sw:customerSearchInitialPayload:v2:' + ss.getId();
}

function swCustomerSearchInitialPayloadCacheGet_(key) {
  try {
    var text = CacheService.getScriptCache().get(key);
    if (text) return swUnpackCustomerSearchInitialPayload_(swParseJson_(text, null));
  } catch (_) {}
  return swUnpackCustomerSearchInitialPayload_(swTaskListCacheGet_(key));
}

function swCustomerSearchInitialPayloadCachePut_(key, payload) {
  var cachePayload = swPackCustomerSearchInitialPayload_(payload);
  try {
    var text = swStringify_(cachePayload);
    if (text.length <= 90000) {
      CacheService.getScriptCache().put(key, text, SW_CUSTOMER_SEARCH_READ_MODEL_CACHE_SECONDS);
      return { ok: true, chunks: 1, bytes: text.length };
    }
  } catch (_) {}
  return swTaskListCachePut_(key, cachePayload);
}

function swCustomerSearchInitialPayloadCacheRemove_(key) {
  try { CacheService.getScriptCache().remove(key); } catch (_) {}
  try { swTaskListCacheRemove_(key); } catch (_) {}
}

function swPackCustomerSearchInitialPayload_(payload) {
  payload = payload || {};
  var response = payload.payload || {};
  var columns = response && response.kanban && response.kanban.columns ? response.kanban.columns : [];
  var dict = [];
  var dictByValue = {};
  function put(value) {
    value = String(value || '');
    if (!value) return 0;
    if (dictByValue[value]) return dictByValue[value];
    dict.push(value);
    dictByValue[value] = dict.length;
    return dict.length;
  }
  function packList(list) {
    return (list || []).map(put);
  }
  var filterOptions = response.filterOptions || {};
  return {
    z: 2,
    d: dict,
    ca: payload.cachedAt || '',
    v: payload.version || '',
    m: payload.modelBuiltAt || '',
    sr: Number(payload.sourceRows || 0),
    fr: Number(payload.filteredRows || 0),
    c: Number(payload.cards || 0),
    h: Number(payload.hiddenCards || 0),
    p: {
      ok: response.ok === false ? 0 : 1,
      g: response.generatedAt || '',
      q: put(response.query || ''),
      s: put(response.source || ''),
      f: response.filters || {},
      o: [
        packList(filterOptions.brands || []),
        packList(filterOptions.clientAdvisors || []),
        packList(filterOptions.jocs || [])
      ],
      k: columns.map(function (col) { return swPackCustomerSearchInitialColumn_(col, put); })
    }
  };
}

function swUnpackCustomerSearchInitialPayload_(payload) {
  if (!payload) return null;
  if (payload.payload && payload.payload.kanban) return payload;
  if (payload.z === 2 && payload.p && payload.p.k) return swUnpackCustomerSearchInitialPayloadV2_(payload);
  if (!(payload.p && payload.p.k)) return payload;
  return {
    cachedAt: payload.ca || '',
    version: payload.v || '',
    modelBuiltAt: payload.m || '',
    sourceRows: Number(payload.sr || 0),
    filteredRows: Number(payload.fr || 0),
    cards: Number(payload.c || 0),
    hiddenCards: Number(payload.h || 0),
    payload: {
      ok: payload.p.ok !== 0,
      generatedAt: payload.p.g || '',
      query: payload.p.q || '',
      source: payload.p.s || '',
      filters: payload.p.f || {},
      filterOptions: payload.p.o || {},
      kanban: {
        columns: (payload.p.k || []).map(swUnpackCustomerSearchInitialColumn_)
      }
    }
  };
}

function swUnpackCustomerSearchInitialPayloadV2_(payload) {
  var dict = payload.d || [];
  function get(value) {
    value = Number(value || 0);
    return value > 0 ? (dict[value - 1] || '') : '';
  }
  function unpackList(list) {
    return (list || []).map(get).filter(Boolean);
  }
  var options = payload.p.o || [];
  return {
    cachedAt: payload.ca || '',
    version: payload.v || '',
    modelBuiltAt: payload.m || '',
    sourceRows: Number(payload.sr || 0),
    filteredRows: Number(payload.fr || 0),
    cards: Number(payload.c || 0),
    hiddenCards: Number(payload.h || 0),
    payload: {
      ok: payload.p.ok !== 0,
      generatedAt: payload.p.g || '',
      query: get(payload.p.q),
      source: get(payload.p.s),
      filters: payload.p.f || {},
      filterOptions: {
        brands: unpackList(options[0] || []),
        clientAdvisors: unpackList(options[1] || []),
        jocs: unpackList(options[2] || [])
      },
      kanban: {
        columns: (payload.p.k || []).map(function (col) {
          return swUnpackCustomerSearchInitialColumnV2_(col, get);
        })
      }
    }
  };
}

function swPackCustomerSearchInitialColumn_(col, put) {
  col = col || {};
  put = put || function (value) { return value || ''; };
  return [
    put(col.key || ''),
    put(col.label || ''),
    Number(col.count || 0),
    Number(col.hiddenCount || 0),
    (col.cards || []).map(function (card) { return swPackCustomerSearchInitialCard_(card, put); })
  ];
}

function swUnpackCustomerSearchInitialColumn_(col) {
  col = col || [];
  return {
    key: col[0] || '',
    label: col[1] || '',
    count: Number(col[2] || 0),
    hiddenCount: Number(col[3] || 0),
    cards: (col[4] || []).map(swUnpackCustomerSearchInitialCard_)
  };
}

function swUnpackCustomerSearchInitialColumnV2_(col, get) {
  col = col || [];
  return {
    key: get(col[0]),
    label: get(col[1]),
    count: Number(col[2] || 0),
    hiddenCount: Number(col[3] || 0),
    cards: (col[4] || []).map(function (card) {
      return swUnpackCustomerSearchInitialCardV2_(card, get);
    })
  };
}

function swPackCustomerSearchInitialCard_(card, put) {
  card = card || {};
  put = put || function (value) { return value || ''; };
  return [
    put(card.root || ''),
    put(card.appt || ''),
    put(card.customerName || ''),
    put(card.brand || ''),
    put(card.clientAdvisor || ''),
    put(card.joc || ''),
    put(card.nextVisit || ''),
    put(card.lastVisit || ''),
    put(card.salesStage || ''),
    put(card.conversionStatus || ''),
    put(card.customOrderStatus || ''),
    put(card.inProductionStatus || ''),
    put(card.so || ''),
    put(card.deadline3d || ''),
    put(card.waxStatus || ''),
    (card.badges || []).map(put)
  ];
}

function swUnpackCustomerSearchInitialCard_(card) {
  card = card || [];
  return {
    root: card[0] || '',
    appt: card[1] || '',
    customerName: card[2] || '',
    brand: card[3] || '',
    clientAdvisor: card[4] || '',
    joc: card[5] || '',
    nextVisit: card[6] || '',
    lastVisit: card[7] || '',
    salesStage: card[8] || '',
    conversionStatus: card[9] || '',
    customOrderStatus: card[10] || '',
    inProductionStatus: card[11] || '',
    so: card[12] || '',
    deadline3d: card[13] || '',
    waxStatus: card[14] || '',
    badges: card[15] || []
  };
}

function swUnpackCustomerSearchInitialCardV2_(card, get) {
  card = card || [];
  return {
    root: get(card[0]),
    appt: get(card[1]),
    customerName: get(card[2]),
    brand: get(card[3]),
    clientAdvisor: get(card[4]),
    joc: get(card[5]),
    nextVisit: get(card[6]),
    lastVisit: get(card[7]),
    salesStage: get(card[8]),
    conversionStatus: get(card[9]),
    customOrderStatus: get(card[10]),
    inProductionStatus: get(card[11]),
    so: get(card[12]),
    deadline3d: get(card[13]),
    waxStatus: get(card[14]),
    badges: (card[15] || []).map(get).filter(Boolean)
  };
}

function swCacheCustomerSearchDetailIndex_(ss, rows, status) {
  var shards = {};
  var recordCount = 0;
  var keyCount = 0;
  (rows || []).forEach(function (rec) {
    if (!rec || !rec.root) return;
    var record = swCustomerSearchDetailIndexRecord_(rec);
    recordCount++;
    [rec.root, rec.appt].forEach(function (id) {
      id = swTrim_(id);
      if (!id) return;
      var shard = swCustomerSearchDetailIndexShard_(id);
      if (!shards[shard]) {
        shards[shard] = {
          cachedAt: swIso_(new Date()),
          version: typeof SW_READ_MODEL_VERSION !== 'undefined' ? SW_READ_MODEL_VERSION : '',
          modelBuiltAt: status && status.builtAt || '',
          records: {}
        };
      }
      shards[shard].records[id] = record;
      keyCount++;
    });
  });
  var meta = {
    cachedAt: swIso_(new Date()),
    version: typeof SW_READ_MODEL_VERSION !== 'undefined' ? SW_READ_MODEL_VERSION : '',
    modelBuiltAt: status && status.builtAt || '',
    shards: SW_CUSTOMER_SEARCH_DETAIL_INDEX_SHARDS,
    records: recordCount,
    keys: keyCount
  };
  var totalChunks = 0;
  var totalBytes = 0;
  var ok = true;
  var reason = '';
  Object.keys(shards).forEach(function (shard) {
    var shardKey = swCustomerSearchDetailIndexShardCacheKey_(ss, shard);
    var payload = shards[shard];
    try {
      SW_CUSTOMER_SEARCH_DETAIL_INDEX_MEMORY_CACHE_[shardKey] = {
        expiresAt: new Date().getTime() + SW_CUSTOMER_SEARCH_DETAIL_CACHE_SECONDS * 1000,
        version: payload.version,
        modelBuiltAt: payload.modelBuiltAt,
        payload: payload
      };
    } catch (_) {}
    var shardResult = swCustomerSearchDetailCachePut_(shardKey, payload) || {};
    totalChunks += shardResult.chunks || 0;
    totalBytes += shardResult.bytes || 0;
    if (shardResult.ok === false) {
      ok = false;
      if (!reason) reason = shardResult.reason || 'detailShardCacheFailed';
    }
  });
  var metaResult = swCustomerSearchDetailCachePut_(swCustomerSearchDetailIndexCacheKey_(ss), meta) || {};
  totalChunks += metaResult.chunks || 0;
  totalBytes += metaResult.bytes || 0;
  if (metaResult.ok === false) {
    ok = false;
    if (!reason) reason = metaResult.reason || 'detailIndexMetaCacheFailed';
  }
  return {
    ok: ok,
    reason: reason,
    chunks: totalChunks,
    bytes: totalBytes,
    records: recordCount,
    keys: keyCount
  };
}

function swCustomerSearchDetailIndexRecord_(rec) {
  rec = rec || {};
  return {
    __customerReadModel: true,
    row: rec.row || '',
    root: rec.root || '',
    appt: rec.appt || '',
    name: rec.name || '',
    email: rec.email || '',
    phone: rec.phone || '',
    brand: rec.brand || '',
    assignedRep: rec.assignedRep || '',
    assignedRepEmail: rec.assignedRepEmail || '',
    assistedRep: rec.assistedRep || '',
    assistedRepEmail: rec.assistedRepEmail || '',
    visitDate: rec.visitDate || '',
    visitTime: rec.visitTime || '',
    visitType: rec.visitType || '',
    nextVisit: rec.nextVisit || '',
    lastVisit: rec.lastVisit || '',
    appointmentCount: rec.appointmentCount || 0,
    activeAppointmentCount: rec.activeAppointmentCount || 0,
    active: rec.active || '',
    activeNorm: rec.activeNorm || '',
    status: rec.status || '',
    statusNorm: rec.statusNorm || '',
    stageKey: rec.stageKey || '',
    stageLabel: rec.stageLabel || '',
    salesStage: rec.salesStage || '',
    convStatus: rec.convStatus || '',
    customOrder: rec.customOrder || '',
    inProduction: rec.inProduction || '',
    centerStoneStatus: rec.centerStoneStatus || '',
    so: rec.so || '',
    orderTotal: rec.orderTotal || '',
    paidToDate: rec.paidToDate || '',
    remainingBalance: rec.remainingBalance || '',
    lastPaymentDate: rec.lastPaymentDate || '',
    quotationUrl: rec.quotationUrl || '',
    clientFolder: rec.clientFolder || '',
    reportUrl: rec.reportUrl || '',
    tracker3dUrl: rec.tracker3dUrl || '',
    deadline3d: rec.deadline3d || '',
    productionDeadline: rec.productionDeadline || '',
    waxStatus: rec.waxStatus || '',
    waxDeadlineAdmin: rec.waxDeadlineAdmin || '',
    dvStonesSummary: rec.dvStonesSummary || '',
    nextSteps: rec.nextSteps || '',
    updatedAt: rec.updatedAt || '',
    hasAiBrief: !!rec.hasAiBrief,
    reviewFlagCount: Number(rec.reviewFlagCount || 0),
    latestAiBriefUpdatedAt: rec.latestAiBriefUpdatedAt || '',
    sourceRows: rec.sourceRows || []
  };
}

function swReadCachedCustomerSearchDetailRecord_(ss, status, rootApptId) {
  var id = swTrim_(rootApptId);
  if (!id) return null;
  var key = swCustomerSearchDetailIndexShardCacheKey_(ss, swCustomerSearchDetailIndexShard_(id));
  var expectedVersion = typeof SW_READ_MODEL_VERSION !== 'undefined' ? SW_READ_MODEL_VERSION : '';
  var payload = null;
  try {
    var memory = SW_CUSTOMER_SEARCH_DETAIL_INDEX_MEMORY_CACHE_[key];
    if (memory && memory.expiresAt > new Date().getTime() &&
        memory.modelBuiltAt === status.builtAt &&
        memory.version === expectedVersion) {
      payload = memory.payload || null;
    }
  } catch (_) {}
  if (!payload) {
    payload = swCustomerSearchDetailCacheGet_(key);
    if (!payload ||
        payload.modelBuiltAt !== status.builtAt ||
        payload.version !== expectedVersion ||
        !payload.records) {
      return null;
    }
    try {
      SW_CUSTOMER_SEARCH_DETAIL_INDEX_MEMORY_CACHE_[key] = {
        expiresAt: new Date().getTime() + SW_CUSTOMER_SEARCH_DETAIL_CACHE_SECONDS * 1000,
        version: payload.version || '',
        modelBuiltAt: payload.modelBuiltAt || '',
        payload: payload
      };
    } catch (_) {}
  }
  return payload.records[id] || null;
}

function swCustomerSearchDetailIndexCacheKey_(ss) {
  return 'sw:customerSearchDetailIndex:v2:' + ss.getId() + ':meta';
}

function swCustomerSearchDetailIndexShardCacheKey_(ss, shard) {
  return 'sw:customerSearchDetailIndex:v2:' + ss.getId() + ':' + shard;
}

function swCustomerSearchDetailCacheGet_(key) {
  try {
    var text = CacheService.getScriptCache().get(key);
    if (text) return swParseJson_(text, null);
  } catch (_) {}
  return swTaskListCacheGet_(key);
}

function swCustomerSearchDetailCachePut_(key, payload) {
  try {
    var text = swStringify_(payload);
    if (text.length <= 90000) {
      CacheService.getScriptCache().put(key, text, SW_CUSTOMER_SEARCH_DETAIL_CACHE_SECONDS);
      return { ok: true, chunks: 1, bytes: text.length };
    }
  } catch (_) {}
  return swTaskListCachePut_(key, payload);
}

function swCustomerSearchDetailCacheRemove_(key) {
  try { CacheService.getScriptCache().remove(key); } catch (_) {}
  try { swTaskListCacheRemove_(key); } catch (_) {}
}

function swCustomerSearchDetailIndexShard_(id) {
  id = swTrim_(id);
  var hash = 0;
  for (var i = 0; i < id.length; i++) {
    hash = ((hash * 31) + id.charCodeAt(i)) % 2147483647;
  }
  return Math.abs(hash) % SW_CUSTOMER_SEARCH_DETAIL_INDEX_SHARDS;
}

function swCustomerSearchFilteredRows_(appointments, query, filters) {
  query = swTrim_(query);
  var q = swNorm_(query);
  var qPhone = swNormPhone_(query);
  var baseRows = (appointments || []).filter(function (rec) {
    if (filters.activeOnly && !swIsAppointmentActive_(rec)) return false;
    if (filters.brand && swNorm_(rec.brand) !== swNorm_(filters.brand)) return false;
    if (filters.clientAdvisor && !swCustomerSearchAdvisorMatches_(rec.assignedRep, filters.clientAdvisor)) return false;
    if (filters.joc && swNorm_(rec.assistedRep) !== swNorm_(filters.joc)) return false;
    return true;
  });
  if (!q) return baseRows;

  var matchedRoots = {};
  baseRows.forEach(function (rec) {
    if (!swCustomerSearchRecordMatches_(rec, q, qPhone)) return;
    var root = swTrim_(rec.root || rec.appt);
    if (root) matchedRoots[root] = true;
  });
  return baseRows.filter(function (rec) {
    return !!matchedRoots[swTrim_(rec.root || rec.appt)];
  });
}

function swCustomerSearchRecordMatches_(rec, q, qPhone) {
  if (rec.searchText && swNorm_(rec.searchText).indexOf(q) >= 0) return true;
  var fields = [
    rec.name, rec.email, rec.phone, rec.root, rec.appt, rec.so, rec.brand,
    rec.assignedRep, rec.assistedRep, rec.visitType, rec.visitDate,
    rec.salesStage, rec.convStatus, rec.customOrder, rec.nextSteps
  ];
  var text = swNorm_(fields.join(' '));
  if (text.indexOf(q) >= 0) return true;
  if (qPhone && swNormPhone_(rec.phone).indexOf(qPhone) >= 0) return true;
  return false;
}

function swCustomerSearchFilterOptions_(appointments) {
  var brands = [];
  var advisors = [];
  var jocs = [];
  (appointments || []).forEach(function (rec) {
    if (!swIsAppointmentActive_(rec)) return;
    if (rec.brand) brands.push(rec.brand);
    swCustomerSearchAdvisorParts_(rec.assignedRep).forEach(function (advisor) { advisors.push(advisor); });
    if (rec.assistedRep) jocs.push(rec.assistedRep);
  });
  return {
    brands: swUnique_(brands).sort(),
    clientAdvisors: swUnique_(advisors).sort(),
    jocs: swUnique_(jocs).sort()
  };
}

function swCustomerSearchAdvisorParts_(value) {
  var seen = {};
  var out = [];
  String(value || '').split(/\s*(?:\/|,|;|\+|&|\band\b)\s*/i).forEach(function (part) {
    part = swTrim_(part);
    var key = swNorm_(part);
    if (!part || seen[key]) return;
    seen[key] = true;
    out.push(part);
  });
  return out;
}

function swCustomerSearchAdvisorMatches_(value, filter) {
  var want = swNorm_(filter);
  if (!want) return true;
  var parts = swCustomerSearchAdvisorParts_(value);
  for (var i = 0; i < parts.length; i++) {
    if (swNorm_(parts[i]) === want) return true;
  }
  return swNorm_(value).indexOf(want) >= 0;
}

function swCustomerSearchApplyDefaultOwnerFilters_(appointments, user, filters) {
  if (!filters || !filters.defaultOwner || !user || user.isAdmin) return;
  if (user.isJoc && !filters.joc) {
    filters.joc = swCustomerSearchDefaultJoc_(appointments, user);
    return;
  }
  if (user.isRep && !filters.clientAdvisor) {
    filters.clientAdvisor = swCustomerSearchDefaultAdvisor_(appointments, user);
  }
}

function swCustomerSearchDefaultAdvisor_(appointments, user) {
  user = user || {};
  if (!user.isRep) return '';
  var candidates = swCustomerSearchUserAdvisorCandidates_(user);
  if (!candidates.length) return '';
  var candidateMap = {};
  candidates.forEach(function (candidate) {
    var key = swNorm_(candidate);
    if (key) candidateMap[key] = true;
  });
  var active = (appointments || []).filter(function (rec) { return swIsAppointmentActive_(rec); });
  for (var i = 0; i < active.length; i++) {
    var parts = swCustomerSearchAdvisorParts_(active[i].assignedRep);
    for (var j = 0; j < parts.length; j++) {
      if (candidateMap[swNorm_(parts[j])]) return parts[j];
    }
  }
  return '';
}

function swCustomerSearchDefaultJoc_(appointments, user) {
  user = user || {};
  if (!user.isJoc) return '';
  var candidates = swCustomerSearchUserAdvisorCandidates_(user);
  if (!candidates.length) return '';
  var candidateMap = {};
  candidates.forEach(function (candidate) {
    var key = swNorm_(candidate);
    if (key) candidateMap[key] = true;
  });
  var active = (appointments || []).filter(function (rec) { return swIsAppointmentActive_(rec); });
  for (var i = 0; i < active.length; i++) {
    var joc = swTrim_(active[i].assistedRep);
    if (joc && candidateMap[swNorm_(joc)]) return joc;
  }
  return '';
}

function swCustomerSearchUserAdvisorCandidates_(user) {
  var out = [];
  var name = swTrim_(user && user.name);
  if (name) {
    out.push(name);
    var nameParts = name.split(/\s+/).filter(Boolean);
    if (nameParts.length) out.push(nameParts[0]);
  }
  var email = swNormEmail_(user && user.email);
  if (email) {
    var local = email.split('@')[0].replace(/[._-]+/g, ' ');
    if (local) {
      out.push(local);
      var localParts = local.split(/\s+/).filter(Boolean);
      if (localParts.length) out.push(localParts[0]);
    }
  }
  return swUnique_(out);
}

function swCustomerSearchCard_(ss, masterGid, root, rec, rootRows, stage, aiBrief) {
  var card = swAdminDashboardCustomerCard_(ss, masterGid, root, rec, rootRows, stage, { byRoot: {}, bySo: {} });
  card.email = rec.email || '';
  card.phone = rec.phone || '';
  card.deadline3d = rec.deadline3d || '';
  card.productionDeadline = rec.productionDeadline || '';
  card.waxStatus = rec.waxStatus || '';
  card.waxDeadlineAdmin = rec.waxDeadlineAdmin || '';
  card.centerStoneStatus = rec.centerStoneStatus || '';
  card.badges = swCustomerSearchBadges_(rec);
  swCustomerSearchApplyAiBriefCompact_(card, aiBrief);
  return card;
}

function swCustomerSearchApplyAiBriefCompact_(card, aiBrief) {
  var compact = typeof swAppointmentAiBriefCompact_ === 'function'
    ? swAppointmentAiBriefCompact_(aiBrief)
    : { hasAiBrief: false, reviewFlagCount: 0, latestAiBriefUpdatedAt: '' };
  card.hasAiBrief = !!compact.hasAiBrief;
  card.reviewFlagCount = Number(compact.reviewFlagCount || 0);
  card.latestAiBriefUpdatedAt = compact.latestAiBriefUpdatedAt || '';
  if (!card.hasAiBrief) return card;
  card.badges = card.badges || [];
  card.badges.push('AI Brief');
  if (card.reviewFlagCount) card.badges.push('Flags: ' + card.reviewFlagCount);
  return card;
}

function swCustomerSearchBadges_(rec) {
  var badges = [];
  if (!swTrim_(rec.assignedRep)) badges.push('Missing Advisor');
  if (!swTrim_(rec.assistedRep)) badges.push('Missing JOC');
  if (/3d revision/i.test(swTrim_(rec.customOrder))) badges.push('3D Revision');
  var d3 = swCustomerSearchDateMs_(rec.deadline3d);
  if (d3 && d3 < swCustomerSearchTodayMs_() && /3d/i.test(swTrim_(rec.customOrder))) badges.push('3D Overdue');
  var waxDeadline = swCustomerSearchDateMs_(rec.waxDeadlineAdmin);
  if (waxDeadline && waxDeadline < swCustomerSearchTodayMs_() && !/complete|cancel/i.test(swTrim_(rec.waxStatus))) badges.push('Wax Issue');
  return badges;
}

function swCustomerSearchDetailPayload_(ss, user, rootApptId) {
  var mark = swStepTimer_('swCustomerSearchDetailPayload');
  var config = swReadConfig_(ss, true);
  mark('config');
  var target = swCustomerSearchResolveRootForDetail_(ss, rootApptId, config);
  mark('resolveRoot', { rows: target.rows.length, source: target.source || '' });
  var master = ss.getSheetByName(SW_SHEETS.MASTER);
  var masterGid = master ? master.getSheetId() : '';
  var aiBrief = swCustomerSearchDetailAiBrief_(ss, target);
  var card = swCustomerSearchDetailCard_(ss, masterGid, target, aiBrief);
  mark('card', { source: target.readModelRec ? 'customerReadModel' : 'appointments' });
  var paymentResult = swCustomerSearchPaymentHistory_(ss, target.root, card.so, 12);
  swCustomerSearchApplyPaymentSummary_(card, paymentResult.rows || [], target.rec);
  mark('payments', { payments: paymentResult.rows ? paymentResult.rows.length : 0, source: paymentResult.source || '' });
  var now = new Date().getTime();
  var rootTasks = typeof swReadTaskListForRoot_ === 'function'
    ? swReadTaskListForRoot_(ss, target.root)
    : (swReadTaskListState_(ss, true).tasks || []).filter(function (t) {
      return swTrim_(t.root) === target.root || swTrim_(t.appt) === target.root;
    });
  var tasks = (rootTasks || []).filter(function (t) {
    return t.status !== SW_STATUSES.COMPLETED;
  }).map(function (t) {
    return swPublicTask_(t, now);
  });
  mark('tasks', { tasks: tasks.length });
  var logs = swCustomerSearchRecentLogs_(ss, target.root);
  mark('logs', { logs: logs.length, source: logs.source || '' });
  var formOptions = swTaskFormOptions_(ss, { taskType: SW_TASKS.POST_CONSULT_STATUS });
  mark('formOptions', { groups: formOptions ? Object.keys(formOptions).length : 0 });

  return {
    ok: true,
    generatedAt: swIso_(new Date()),
    user: user,
    root: target.root,
    card: card,
    aiBrief: aiBrief,
    appointments: target.rows.map(swCustomerSearchPublicAppointment_),
    tasks: tasks,
    logs: logs,
    paymentHistory: paymentResult.rows || [],
    paymentHistoryUnavailable: paymentResult.unavailable || '',
    formOptions: formOptions,
    actions: {
      updateStatus: true,
      update3dDeadline: true,
      submit3dRevision: true,
      requestWax: true
    }
  };
}

function swCustomerSearchDetailCard_(ss, masterGid, target, aiBrief) {
  target = target || {};
  if (target.readModelRec) {
    return swCustomerSearchCardFromReadModel_(
      ss,
      masterGid,
      target.readModelRec,
      swCustomerSearchStageForRecord_(target.readModelRec, target.rows)
    );
  }
  var stage = swAdminDashboardPipelineStage_(target.rec, target.rows);
  return swCustomerSearchCard_(ss, masterGid, target.root, target.rec, target.rows, stage, aiBrief);
}

function swCustomerSearchDetailAiBrief_(ss, target) {
  target = target || {};
  if (target.readModelRec && !target.readModelRec.hasAiBrief) return null;
  return typeof swAppointmentAiBriefForRoot_ === 'function'
    ? swAppointmentAiBriefForRoot_(ss, target.root)
    : null;
}

function swCustomerSearchPaymentHistory_(ss, root, so, limit) {
  var cached = swCustomerSearchCachedPaymentHistory_(ss, root, so, limit);
  if (cached) return cached;

  var warnings = [];
  if (typeof swAdminDashboardReadPaymentReceiptRows_ !== 'function') {
    return { rows: [], unavailable: 'Payments ledger helper is unavailable.' };
  }

  var source = swAdminDashboardReadPaymentReceiptRows_(warnings);
  var values = source && source.rows ? source.rows : [];
  var wantRoot = swAdminDashboardCleanId_(root);
  var wantSo = swAdminDashboardCleanId_(so);
  var seen = {};
  var rows = [];

  values.forEach(function (row) {
    var rowRoot = swAdminDashboardCleanId_(row.root || '');
    var rowSo = swAdminDashboardCleanId_(row.so || '');
    if (!(wantRoot && rowRoot === wantRoot) && !(wantSo && rowSo === wantSo)) return;

    var when = new Date(Number(row.whenMs || 0));
    if (isNaN(when.getTime())) return;
    var key = row.paymentId || row.docNumber || [swAdminDashboardDateKey_(when), row.net, row.gross, row.method, rowSo, rowRoot].join('|');
    if (seen[key]) return;
    seen[key] = true;

    rows.push({
      root: rowRoot,
      so: rowSo,
      paymentId: row.paymentId || '',
      docType: row.docType || 'Receipt',
      docNumber: row.docNumber || '',
      method: row.method || '',
      date: swAdminDashboardDateKey_(when),
      whenMs: when.getTime(),
      amountNet: swAdminDashboardNumber_(row.net),
      amountGross: swAdminDashboardNumber_(row.gross === '' || row.gross == null ? row.net : row.gross),
      balanceDue: row.balance === '' || row.balance == null ? '' : swAdminDashboardNumber_(row.balance),
      orderTotal: row.orderTotal === '' || row.orderTotal == null ? '' : swAdminDashboardNumber_(row.orderTotal)
    });
  });

  rows.sort(function (a, b) { return Number(b.whenMs || 0) - Number(a.whenMs || 0); });
  if (limit && limit > 0) rows = rows.slice(0, limit);
  return {
    rows: rows,
    unavailable: warnings.length ? warnings.join(' ') : '',
    source: source && source.cacheHit ? 'paymentReceiptCache' : 'paymentReceiptRows'
  };
}

function swCustomerSearchCachedPaymentHistory_(ss, root, so, limit) {
  var payload = swCustomerSearchPaymentHistoryIndex_(ss);
  if (!(payload && payload.byKey)) return null;
  var keys = [swAdminDashboardCleanId_(root), swAdminDashboardCleanId_(so)].filter(Boolean);
  if (!keys.length) return null;
  var seen = {};
  var rows = [];
  keys.forEach(function (key) {
    (payload.byKey[key] || []).forEach(function (row) {
      var rowKey = row.paymentId || row.docNumber || [row.date, row.amountNet, row.amountGross, row.method, row.so, row.root].join('|');
      if (seen[rowKey]) return;
      seen[rowKey] = true;
      rows.push(row);
    });
  });
  rows.sort(function (a, b) { return Number(b.whenMs || 0) - Number(a.whenMs || 0); });
  if (limit && limit > 0) rows = rows.slice(0, limit);
  return {
    rows: rows,
    unavailable: payload.unavailable || '',
    source: 'customerPaymentHistoryCache'
  };
}

function swCustomerSearchPaymentHistoryIndex_(ss) {
  var key = swCustomerSearchPaymentHistoryIndexCacheKey_(ss);
  try {
    var memory = SW_CUSTOMER_SEARCH_PAYMENT_HISTORY_MEMORY_CACHE_[key];
    if (memory && memory.expiresAt > new Date().getTime()) return memory.payload || null;
  } catch (_) {}
  var payload = swTaskListCacheGet_(key);
  if (!payload) return null;
  try {
    SW_CUSTOMER_SEARCH_PAYMENT_HISTORY_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_CUSTOMER_SEARCH_DETAIL_CACHE_SECONDS * 1000,
      payload: payload
    };
  } catch (_) {}
  return payload;
}

function swCacheCustomerSearchPaymentHistoryIndex_(ss) {
  var warnings = [];
  if (typeof swAdminDashboardReadPaymentReceiptRows_ !== 'function') {
    return { ok: false, reason: 'paymentsHelperUnavailable', keys: 0, rows: 0 };
  }
  var source = swAdminDashboardReadPaymentReceiptRows_(warnings);
  var values = source && source.rows ? source.rows : [];
  var byKey = {};
  values.forEach(function (row) {
    var formatted = swCustomerSearchPaymentHistoryRowFromReceipt_(row);
    if (!formatted) return;
    [formatted.root, formatted.so].forEach(function (key) {
      key = swAdminDashboardCleanId_(key);
      if (!key) return;
      if (!byKey[key]) byKey[key] = [];
      byKey[key].push(formatted);
    });
  });
  Object.keys(byKey).forEach(function (key) {
    var seen = {};
    byKey[key] = byKey[key].filter(function (row) {
      var rowKey = row.paymentId || row.docNumber || [row.date, row.amountNet, row.amountGross, row.method, row.so, row.root].join('|');
      if (seen[rowKey]) return false;
      seen[rowKey] = true;
      return true;
    }).sort(function (a, b) {
      return Number(b.whenMs || 0) - Number(a.whenMs || 0);
    }).slice(0, 25);
  });
  var payload = {
    cachedAt: swIso_(new Date()),
    unavailable: warnings.length ? warnings.join(' ') : '',
    byKey: byKey
  };
  var cacheKey = swCustomerSearchPaymentHistoryIndexCacheKey_(ss);
  var cacheResult = swTaskListCachePut_(cacheKey, payload) || {};
  try {
    SW_CUSTOMER_SEARCH_PAYMENT_HISTORY_MEMORY_CACHE_[cacheKey] = {
      expiresAt: new Date().getTime() + SW_CUSTOMER_SEARCH_DETAIL_CACHE_SECONDS * 1000,
      payload: payload
    };
  } catch (_) {}
  return {
    ok: cacheResult.ok !== false,
    keys: Object.keys(byKey).length,
    rows: values.length,
    chunks: cacheResult.chunks || 0,
    bytes: cacheResult.bytes || 0,
    reason: cacheResult.reason || ''
  };
}

function swCustomerSearchPaymentHistoryRowFromReceipt_(row) {
  row = row || {};
  var when = new Date(Number(row.whenMs || 0));
  if (isNaN(when.getTime())) return null;
  return {
    root: swAdminDashboardCleanId_(row.root || ''),
    so: swAdminDashboardCleanId_(row.so || ''),
    paymentId: row.paymentId || '',
    docType: row.docType || 'Receipt',
    docNumber: row.docNumber || '',
    method: row.method || '',
    date: swAdminDashboardDateKey_(when),
    whenMs: when.getTime(),
    amountNet: swAdminDashboardNumber_(row.net),
    amountGross: swAdminDashboardNumber_(row.gross === '' || row.gross == null ? row.net : row.gross),
    balanceDue: row.balance === '' || row.balance == null ? '' : swAdminDashboardNumber_(row.balance),
    orderTotal: row.orderTotal === '' || row.orderTotal == null ? '' : swAdminDashboardNumber_(row.orderTotal)
  };
}

function swCustomerSearchPaymentHistoryIndexCacheKey_(ss) {
  return 'sw:customerPaymentHistory:v1:' + ss.getId();
}

function swCustomerSearchApplyPaymentSummary_(card, paymentRows, rec) {
  paymentRows = paymentRows || [];
  rec = rec || {};
  var paidNet = 0;
  paymentRows.forEach(function (row) {
    paidNet += Number(row.amountNet || 0);
  });

  var latest = paymentRows.length ? paymentRows[0] : null;
  var recPaid = swAdminDashboardNumberOrBlank_(rec.paidToDate);
  var recBalance = swAdminDashboardNumberOrBlank_(rec.remainingBalance);
  var recOrderTotal = swAdminDashboardNumberOrBlank_(rec.orderTotal);

  card.paymentCount = paymentRows.length;
  card.paidNet = paymentRows.length ? paidNet : (recPaid === '' ? 0 : recPaid);
  card.balanceDue = latest && latest.balanceDue !== '' ? latest.balanceDue : (recBalance === '' ? '' : recBalance);
  card.orderTotal = latest && latest.orderTotal !== '' ? latest.orderTotal : (recOrderTotal === '' ? '' : recOrderTotal);
  card.lastPaymentDate = latest ? latest.date : (rec.lastPaymentDate || '');
}

function swCustomerSearchPublicAppointment_(rec) {
  return {
    row: rec.row || '',
    root: rec.root || '',
    appt: rec.appt || '',
    customerName: rec.name || '',
    brand: rec.brand || '',
    visitDate: rec.visitDate || '',
    visitTime: rec.visitTime || '',
    visitType: rec.visitType || '',
    status: rec.status || '',
    assignedRep: rec.assignedRep || '',
    assistedRep: rec.assistedRep || '',
    so: rec.so || '',
    salesStage: rec.salesStage || '',
    conversionStatus: rec.convStatus || '',
    customOrderStatus: rec.customOrder || '',
    active: swIsAppointmentActive_(rec)
  };
}

function swCustomerSearchRecentLogs_(ss, root) {
  var cached = swCustomerSearchCachedRecentLogs_(ss, root);
  if (cached) return cached;
  return swCustomerSearchRecentLogsFromSheet_(ss, root);
}

function swCustomerSearchCachedRecentLogs_(ss, root) {
  var payload = swCustomerSearchRecentLogIndex_(ss);
  if (!(payload && payload.byKey)) return null;
  var logs = payload.byKey[swTrim_(root)] || [];
  logs = logs.slice(0, 20);
  logs.source = 'customerRecentLogCache';
  return logs;
}

function swCustomerSearchRecentLogIndex_(ss) {
  var key = swCustomerSearchRecentLogIndexCacheKey_(ss);
  try {
    var memory = SW_CUSTOMER_SEARCH_RECENT_LOG_MEMORY_CACHE_[key];
    if (memory && memory.expiresAt > new Date().getTime()) return memory.payload || null;
  } catch (_) {}
  var payload = swTaskListCacheGet_(key);
  if (!payload) return null;
  try {
    SW_CUSTOMER_SEARCH_RECENT_LOG_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_CUSTOMER_SEARCH_DETAIL_CACHE_SECONDS * 1000,
      payload: payload
    };
  } catch (_) {}
  return payload;
}

function swCacheCustomerSearchRecentLogIndex_(ss) {
  var sh = ss.getSheetByName(SW_SHEETS.LOG);
  var byKey = {};
  if (sh && sh.getLastRow() >= 2) {
    var last = sh.getLastRow();
    var start = Math.max(2, last - SW_CUSTOMER_SEARCH_LOG_SCAN_ROWS + 1);
    var values = sh.getRange(start, 1, last - start + 1, Math.min(sh.getLastColumn(), SW_LOG_HEADERS.length)).getDisplayValues();
    for (var i = values.length - 1; i >= 0; i--) {
      var log = swCustomerSearchLogFromRow_(values[i]);
      [log.root, log.appt].forEach(function (key) {
        key = swTrim_(key);
        if (!key) return;
        if (!byKey[key]) byKey[key] = [];
        if (byKey[key].length < 20) byKey[key].push(log);
      });
    }
  }
  var payload = { cachedAt: swIso_(new Date()), byKey: byKey };
  var cacheResult = swTaskListCachePut_(swCustomerSearchRecentLogIndexCacheKey_(ss), payload) || {};
  try {
    SW_CUSTOMER_SEARCH_RECENT_LOG_MEMORY_CACHE_[swCustomerSearchRecentLogIndexCacheKey_(ss)] = {
      expiresAt: new Date().getTime() + SW_CUSTOMER_SEARCH_DETAIL_CACHE_SECONDS * 1000,
      payload: payload
    };
  } catch (_) {}
  return {
    ok: cacheResult.ok !== false,
    keys: Object.keys(byKey).length,
    chunks: cacheResult.chunks || 0,
    bytes: cacheResult.bytes || 0,
    reason: cacheResult.reason || ''
  };
}

function swCustomerSearchRecentLogsFromSheet_(ss, root) {
  var sh = ss.getSheetByName(SW_SHEETS.LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  var last = sh.getLastRow();
  var start = Math.max(2, last - SW_CUSTOMER_SEARCH_LOG_SCAN_ROWS + 1);
  var values = sh.getRange(start, 1, last - start + 1, Math.min(sh.getLastColumn(), SW_LOG_HEADERS.length)).getDisplayValues();
  var out = [];
  for (var i = values.length - 1; i >= 0 && out.length < 20; i--) {
    var row = values[i];
    if (swTrim_(row[3]) !== root && swTrim_(row[4]) !== root) continue;
    out.push(swCustomerSearchLogFromRow_(row));
  }
  out.source = 'statusLogRows';
  return out;
}

function swCustomerSearchLogFromRow_(row) {
  row = row || [];
  return {
    eventAt: row[0] || '',
    eventType: row[1] || '',
    taskId: row[2] || '',
    root: row[3] || '',
    appt: row[4] || '',
    taskType: row[5] || '',
    actorName: row[6] || '',
    actorEmail: row[7] || '',
    status: row[10] || '',
    detailsJson: row[11] || ''
  };
}

function swPrewarmCustomerSearchDetailCaches_(ss) {
  var started = new Date().getTime();
  var out = {
    ok: true,
    paymentKeys: 0,
    paymentRows: 0,
    paymentBytes: 0,
    logKeys: 0,
    logBytes: 0,
    formOptionGroups: 0,
    ms: 0,
    error: ''
  };
  try {
    var payment = swCacheCustomerSearchPaymentHistoryIndex_(ss);
    out.paymentKeys = payment.keys || 0;
    out.paymentRows = payment.rows || 0;
    out.paymentBytes = payment.bytes || 0;
    if (payment.ok === false) {
      out.ok = false;
      out.error = payment.reason || 'paymentCacheFailed';
    }
  } catch (payErr) {
    out.ok = false;
    out.error = payErr && payErr.message ? payErr.message : String(payErr);
  }
  try {
    var logs = swCacheCustomerSearchRecentLogIndex_(ss);
    out.logKeys = logs.keys || 0;
    out.logBytes = logs.bytes || 0;
    if (logs.ok === false && !out.error) {
      out.ok = false;
      out.error = logs.reason || 'logCacheFailed';
    }
  } catch (logErr) {
    out.ok = false;
    if (!out.error) out.error = logErr && logErr.message ? logErr.message : String(logErr);
  }
  try {
    var formOptions = swTaskFormOptions_(ss, { taskType: SW_TASKS.POST_CONSULT_STATUS });
    out.formOptionGroups = formOptions ? Object.keys(formOptions).length : 0;
    if (typeof swIsDataCleanupTaskType_ === 'function' && SW_TASKS && SW_TASKS.DATA_CLEANUP_REVIEW) {
      var cleanupOptions = swTaskFormOptions_(ss, { taskType: SW_TASKS.DATA_CLEANUP_REVIEW });
      out.formOptionGroups += cleanupOptions ? Object.keys(cleanupOptions).length : 0;
    }
  } catch (formErr) {
    out.ok = false;
    if (!out.error) out.error = formErr && formErr.message ? formErr.message : String(formErr);
  }
  out.ms = new Date().getTime() - started;
  return out;
}

function swInvalidateCustomerSearchDetailCaches_(ss) {
  var detailIndexKey = swCustomerSearchDetailIndexCacheKey_(ss);
  var paymentKey = swCustomerSearchPaymentHistoryIndexCacheKey_(ss);
  var logKey = swCustomerSearchRecentLogIndexCacheKey_(ss);
  try { delete SW_CUSTOMER_SEARCH_DETAIL_INDEX_MEMORY_CACHE_[detailIndexKey]; } catch (_) {}
  try { delete SW_CUSTOMER_SEARCH_PAYMENT_HISTORY_MEMORY_CACHE_[paymentKey]; } catch (_) {}
  try { delete SW_CUSTOMER_SEARCH_RECENT_LOG_MEMORY_CACHE_[logKey]; } catch (_) {}
  swCustomerSearchDetailCacheRemove_(detailIndexKey);
  for (var shard = 0; shard < SW_CUSTOMER_SEARCH_DETAIL_INDEX_SHARDS; shard++) {
    var shardKey = swCustomerSearchDetailIndexShardCacheKey_(ss, shard);
    try { delete SW_CUSTOMER_SEARCH_DETAIL_INDEX_MEMORY_CACHE_[shardKey]; } catch (_) {}
    swCustomerSearchDetailCacheRemove_(shardKey);
  }
  try { swTaskListCacheRemove_(paymentKey); } catch (_) {}
  try { swTaskListCacheRemove_(logKey); } catch (_) {}
}

function swCustomerSearchRecentLogIndexCacheKey_(ss) {
  return 'sw:customerRecentLogs:v1:' + ss.getId();
}

function swCustomerSearchResolveRootForDetail_(ss, rootApptId, config) {
  var target = swCustomerSearchResolveRootFromReadModel_(ss, rootApptId, config);
  if (target) return target;
  target = swCustomerSearchResolveRoot_(ss, rootApptId);
  target.source = 'appointments';
  return target;
}

function swCustomerSearchResolveRootFromReadModel_(ss, rootApptId, config) {
  var want = swTrim_(rootApptId);
  if (!want) throw new Error('Missing customer/root id.');
  if (typeof swCustomerReadModelServingEnabled_ === 'function' &&
      !swCustomerReadModelServingEnabled_(config || [])) {
    return null;
  }
  if (typeof swCustomerReadModelStatus_ === 'function') {
    var status = swCustomerReadModelStatus_(ss);
    if (status && status.fresh) {
      var indexedRec = swReadCachedCustomerSearchDetailRecord_(ss, status, want);
      if (indexedRec) return swCustomerSearchDetailTargetFromReadModelRec_(ss, indexedRec, 'customerDetailIndexCache');
    }
  }
  var readModel = swCustomerSearchReadModelRows_(ss, config || []);
  if (!(readModel && readModel.ok && readModel.rows && readModel.rows.length)) return null;

  var rec = null;
  for (var i = 0; i < readModel.rows.length; i++) {
    var row = readModel.rows[i];
    if (swTrim_(row.root) === want || swTrim_(row.appt) === want) {
      rec = row;
      break;
    }
  }
  if (!rec) return null;
  return swCustomerSearchDetailTargetFromReadModelRec_(ss, rec, readModel.source || 'customerReadModel');
}

function swCustomerSearchDetailTargetFromReadModelRec_(ss, rec, source) {
  var root = swTrim_(rec.root || rec.appt || '');
  var sourceRows = rec.sourceRows || [];
  var rows = sourceRows.length > 1
    ? swCustomerSearchReadAppointmentRows_(ss, sourceRows, root)
    : [];
  if (!rows.length) rows = [swCustomerSearchAppointmentLikeFromReadModel_(rec)];
  var active = rows.filter(function (row) { return swIsAppointmentActive_(row); });
  var latest = rows.length
    ? swAdminDashboardLatestRow_(active.length ? active : rows)
    : swCustomerSearchAppointmentLikeFromReadModel_(rec);
  return {
    root: root,
    rec: latest,
    rows: rows.length ? rows : [latest],
    readModelRec: rec,
    source: source || 'customerReadModel'
  };
}

function swCustomerSearchReadAppointmentRows_(ss, rowNumbers, root) {
  rowNumbers = (rowNumbers || []).map(function (row) { return Number(row); })
    .filter(function (row) { return isFinite(row) && row >= 2; });
  if (!rowNumbers.length) return [];

  var sh = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!sh) throw new Error('Missing sheet: ' + SW_SHEETS.MASTER);
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  var seen = {};
  rowNumbers = rowNumbers.filter(function (row) {
    if (row > lastRow || seen[row]) return false;
    seen[row] = true;
    return true;
  }).sort(function (a, b) { return a - b; });
  if (!rowNumbers.length) return [];

  var headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (h) { return swTrim_(h); });
  var idx = swAppointmentColumnIndex_(headers);
  var indexes = swAppointmentColumnIndexes_(idx);
  var want = swTrim_(root);
  var out = [];
  rowNumbers.forEach(function (rowNumber) {
    var values = swReadSelectedRows_(sh, rowNumber, 1, indexes, 'values')[0] || [];
    var display = swReadSelectedRows_(sh, rowNumber, 1, indexes, 'display')[0] || [];
    var rec = swAppointmentRecordFromRows_(display, values, idx, rowNumber);
    if (!want || swTrim_(rec.root) === want || swTrim_(rec.appt) === want) out.push(rec);
  });
  return out;
}

function swCustomerSearchAppointmentLikeFromReadModel_(rec) {
  rec = rec || {};
  return {
    row: rec.row || '',
    root: rec.root || '',
    appt: rec.appt || '',
    name: rec.name || '',
    email: rec.email || '',
    phone: rec.phone || '',
    brand: rec.brand || '',
    visitDate: rec.visitDate || '',
    visitTime: rec.visitTime || '',
    visitType: rec.visitType || '',
    status: rec.activeNorm === 'yes' || rec.active === 'Yes' || rec.active === true ? 'Active' : '',
    statusNorm: '',
    active: rec.active || '',
    activeNorm: rec.activeNorm || '',
    assignedRep: rec.assignedRep || '',
    assignedRepEmail: rec.assignedRepEmail || '',
    assistedRep: rec.assistedRep || '',
    assistedRepEmail: rec.assistedRepEmail || '',
    so: rec.so || '',
    orderTotal: rec.orderTotal || '',
    paidToDate: rec.paidToDate || '',
    remainingBalance: rec.remainingBalance || '',
    lastPaymentDate: rec.lastPaymentDate || '',
    quotationUrl: rec.quotationUrl || '',
    clientFolder: rec.clientFolder || '',
    reportUrl: rec.reportUrl || '',
    tracker3dUrl: rec.tracker3dUrl || '',
    deadline3d: rec.deadline3d || '',
    productionDeadline: rec.productionDeadline || '',
    waxStatus: rec.waxStatus || '',
    waxDeadlineAdmin: rec.waxDeadlineAdmin || '',
    salesStage: rec.salesStage || '',
    convStatus: rec.convStatus || '',
    customOrder: rec.customOrder || '',
    inProduction: rec.inProduction || '',
    centerStoneStatus: rec.centerStoneStatus || '',
    nextSteps: rec.nextSteps || ''
  };
}

function swCustomerSearchResolveRoot_(ss, rootApptId) {
  var want = swTrim_(rootApptId);
  if (!want) throw new Error('Missing customer/root id.');
  var rows = typeof swReadAppointmentsForRoot_ === 'function'
    ? swReadAppointmentsForRoot_(ss, want)
    : swReadAppointments_(ss).filter(function (rec) {
      return swTrim_(rec.root) === want || swTrim_(rec.appt) === want;
    });
  if (!rows.length) throw new Error('Customer not found: ' + want);
  var active = rows.filter(function (rec) { return swIsAppointmentActive_(rec); });
  var rec = swAdminDashboardLatestRow_(active.length ? active : rows);
  var root = swTrim_(rec.root || rec.appt || want);
  return { root: root, rec: rec, rows: rows };
}

function swCustomerSearchActivateRoot_(ss, rootApptId) {
  var target = swCustomerSearchResolveRoot_(ss, rootApptId);
  var sh = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!sh || !target.rec.row) throw new Error('Could not resolve Master row for this customer.');
  ss.setActiveSheet(sh);
  ss.setActiveRange(sh.getRange(target.rec.row, 1));
  return target;
}

function swCustomerSearchValidateResolvedRow_(ss, target, rootApptId) {
  var row = Number(target && target.rec && target.rec.row || 0);
  var sh = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!sh || row <= 1) throw new Error('Could not resolve Master row for this customer.');

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  var H = swHeaderMapFromArray_(headers);
  var cRoot = swPickIndex_(H, ['RootApptID', 'Root Appt ID', 'ROOT', 'Root_ID']);
  var cAppt = swPickIndex_(H, ['APPT_ID', 'Appt ID', 'APPTID', 'Appointment ID']);
  var values = sh.getRange(row, 1, 1, sh.getLastColumn()).getDisplayValues()[0];

  var wantRoot = swTrim_(target.root || rootApptId);
  var wantAppt = swTrim_((target.rec && target.rec.appt) || rootApptId);
  var rowRoot = cRoot >= 0 ? swTrim_(values[cRoot]) : '';
  var rowAppt = cAppt >= 0 ? swTrim_(values[cAppt]) : '';
  var matches = (rowRoot && (rowRoot === wantRoot || rowRoot === rootApptId)) ||
    (rowAppt && (rowAppt === wantRoot || rowAppt === wantAppt || rowAppt === rootApptId));

  if (!matches) throw new Error('Resolved Master row no longer matches this customer. Refresh customer detail and try again.');
  return row;
}

function swCustomerSearchLog_(ss, eventType, rec, user, payload, result) {
  swAppendTaskLog_(ss, eventType, {
    taskId: '',
    root: rec.root || rec.appt || '',
    appt: rec.appt || '',
    taskType: eventType,
    status: SW_STATUSES.COMPLETED
  }, user, rec.assignedRep || '', rec.assignedRep || '', {
    payload: payload || {},
    result: result || {}
  });
}

function swCustomerSearchDateMs_(value) {
  value = swTrim_(value);
  if (!value) return 0;
  var iso = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(value);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])).getTime();
  var mdy = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/.exec(value);
  if (mdy) {
    var y = Number(mdy[3]);
    if (y < 100) y += 2000;
    return new Date(y, Number(mdy[1]) - 1, Number(mdy[2])).getTime();
  }
  var d = new Date(value);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function swCustomerSearchTodayMs_() {
  var today = new Date();
  return new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
}
