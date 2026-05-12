/**
 * One-time test customer data cleanup.
 *
 * Dry run:
 *   sw_previewTestDataCleanupOnce()
 *
 * Apply:
 *   sw_applyTestDataCleanupOnce({ confirmationToken: '<token from preview>' })
 *
 * This deletes source rows only. Generated _SW_* read-model tabs are left alone
 * and are invalidated/rebuilt after source rows are deleted.
 */

var SW_TEST_DATA_CLEANUP_PREVIEW_PROPERTY_ = 'SW_TEST_DATA_CLEANUP_ONCE_LAST_PREVIEW';
var SW_TEST_DATA_CLEANUP_PREVIEW_MAX_AGE_MS_ = 4 * 60 * 60 * 1000;
var SW_TEST_DATA_CLEANUP_REASON_ = 'One-time test customer data cleanup';
var SW_TEST_DATA_CLEANUP_PREVIEW_TAB_NAME_ = 'TestDataCleanup_Preview';
var SW_EXACT_MASTER_CUSTOMER_CLEANUP_PREVIEW_TAB_NAME_ = 'ExactMasterCustomerCleanup_Preview';
var SW_EXACT_MASTER_CUSTOMER_CLEANUP_DEFAULT_NAMES_ = [
  'testdemo3',
  'testdemo34',
  'testdemo5',
  'Adrian Test',
  'Vivianne Tran',
  'test05052 test05052',
  'Test Booking Paul',
  'CodexLive AcuityFlow'
];

var SW_TEST_DATA_CLEANUP_WORKFLOW_SOURCE_ORDER_ = [
  '00_Master Appointments',
  '02_Form_Inbox',
  '_ExternalBookingEvents',
  '_IntakeQueue',
  '_SalesTaskQueue',
  '_SalesTaskLog',
  '_AppointmentArtifacts',
  '_SalesDataCleanup',
  '03_Client_Status_Log',
  '05_Wax_Requests',
  '07_Root_Index'
];

var SW_TEST_DATA_CLEANUP_READ_MODEL_SHEETS_ = {
  '_SW_TaskReadModel': true,
  '_SW_CustomerReadModel': true,
  '_SW_DiamondReadModel': true,
  '_SW_DiamondRootReadModel': true,
  '_SW_AppointmentReadModel': true,
  '_SW_CalendarMonthReadModel': true,
  '_SW_PaymentReadModel': true,
  '_SW_AdminDashboardReadModel': true,
  '_SW_ReadModelMeta': true
};

function sw_previewTestDataCleanupOnce(options) {
  options = options || {};
  swRequireTestDataCleanupAdmin_(options);
  var plan = swBuildTestDataCleanupPlan_(options);
  swStoreTestDataCleanupPreview_(plan);
  swLogTestDataCleanupPlan_(plan, 'SW_TEST_DATA_CLEANUP_PREVIEW');
  return swPublicTestDataCleanupResult_(plan);
}

function sw_applyTestDataCleanupOnce(options) {
  options = swNormalizeTestDataCleanupApplyOptions_(options);
  swRequireTestDataCleanupAdmin_(options);

  var lock = LockService.getScriptLock();
  lock.waitLock(Number(options.lockWaitMs || 30000));
  try {
    var plan = swBuildTestDataCleanupPlan_(options);
    var validation = swValidateTestDataCleanupPreview_(plan, options);
    plan.previewValidation = validation;
    if (!validation.ok) {
      plan.ok = false;
      plan.errors.push({ type: 'preview', message: validation.message });
      swLogTestDataCleanupPlan_(plan, 'SW_TEST_DATA_CLEANUP_APPLY_BLOCKED');
      return swPublicTestDataCleanupResult_(plan);
    }

    var confirmation = swConfirmTestDataCleanupApply_(plan, options);
    plan.confirmation = confirmation;
    if (!confirmation.ok) {
      plan.ok = false;
      plan.errors.push({ type: 'confirmation', message: confirmation.message || 'Apply was not confirmed.' });
      swLogTestDataCleanupPlan_(plan, 'SW_TEST_DATA_CLEANUP_APPLY_CANCELLED');
      return swPublicTestDataCleanupResult_(plan);
    }

    if (plan.totalCandidateRows > 0) {
      plan.deletedRows = swApplyTestDataCleanupRows_(plan, options);
      plan.deletedCount = plan.deletedRows.length;
      plan.invalidation = swInvalidateAfterTestDataCleanup_(plan, options);
      swRecordTestDataCleanupResultFailures_(plan, plan.invalidation, 'invalidation');
      if (options.rebuildReadModels !== false) {
        plan.readModelRebuild = swRebuildAfterTestDataCleanup_();
        swRecordTestDataCleanupResultFailures_(plan, plan.readModelRebuild, 'readModelRebuild');
      }
    }

    plan.ok = plan.errors.length === 0;
    swLogTestDataCleanupPlan_(plan, 'SW_TEST_DATA_CLEANUP_APPLY');
    return swPublicTestDataCleanupResult_(plan);
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function sw_runTestDataCleanupOnceFromMenu() {
  var ui = SpreadsheetApp.getUi();
  var preview = null;
  try {
    preview = sw_previewTestDataCleanupOnce();
    if (!preview || !preview.ok) {
      throw new Error((preview && preview.summary && preview.summary.errors && preview.summary.errors[0] && preview.summary.errors[0].message) ||
        'Preview could not be built.');
    }

    if (preview.totalCandidateRows === 0) {
      ui.alert('Test data cleanup', 'No matching test data rows were found. Nothing to delete.', ui.ButtonSet.OK);
      return preview;
    }

    swWriteTestDataCleanupCandidatesTab_(preview);
    swOpenTestDataCleanupCandidatesSheet_(preview);
    if (!swPreviewTestDataCleanupForMenu_(ui, preview)) {
      return preview;
    }

    var prompt = ui.prompt(
      'Pre-commit test data cleanup',
      'Type this token to run cleanup:\n' + preview.confirmationToken +
      '\n\n(You already reviewed candidate rows in the previous step.)',
      ui.ButtonSet.OK_CANCEL
    );

    if (prompt.getSelectedButton() !== ui.Button.OK) {
      ui.alert('Test data cleanup', 'Cleanup was canceled by user.', ui.ButtonSet.OK);
      return preview;
    }

    var response = swTestDataCleanupTrim_(prompt.getResponseText());
    if (response !== preview.confirmationToken) {
      ui.alert('Test data cleanup', 'Token mismatch. No rows were deleted.', ui.ButtonSet.OK);
      return preview;
    }

    var result = sw_applyTestDataCleanupOnce({ confirmationToken: response });
    if (!result || !result.ok) {
      var firstErr = (result && result.errors && result.errors[0] && result.errors[0].message) || 'Cleanup reported errors.';
      ui.alert('Test data cleanup', 'Cleanup did not complete cleanly. ' + firstErr, ui.ButtonSet.OK);
      return result;
    }

    ui.alert(
      'Test data cleanup',
      'Completed cleanup.\n\nDeleted rows: ' + (result.deletedCount || 0) +
      '\nDownstream invalidation: ' + (result.invalidation ? 'done' : 'skipped') +
      '\nRead-model rebuild: ' +
      ((result.readModelRebuild && result.readModelRebuild.ok) ? 'done' : 'skipped'),
      ui.ButtonSet.OK
    );
    return result;
  } catch (err) {
    ui.alert('Test data cleanup error', (err && err.message) ? err.message : String(err), ui.ButtonSet.OK);
    return preview && preview.ok === false ? preview : { ok: false, errors: [{ type: 'menu', message: err && err.message ? err.message : String(err) }] };
  }
}

function sw_previewTestDataCleanupOnceFromMenu() {
  var ui = SpreadsheetApp.getUi();
  var preview = null;
  try {
    preview = sw_previewTestDataCleanupOnce();
    if (!preview || preview.ok === false) {
      throw new Error('Unable to build candidate preview.');
    }
    if (preview.totalCandidateRows === 0) {
      ui.alert('Test data cleanup', 'No matching test rows found.', ui.ButtonSet.OK);
      return preview;
    }
    swWriteTestDataCleanupCandidatesTab_(preview);
    swOpenTestDataCleanupCandidatesSheet_(preview);
    swPreviewTestDataCleanupForMenu_(ui, preview);
    return preview;
  } catch (err) {
    ui.alert('Test data cleanup', (err && err.message) ? err.message : String(err), ui.ButtonSet.OK);
    return { ok: false, errors: [{ type: 'menu', message: err && err.message ? err.message : String(err) }] };
  }
}

function sw_previewExactMasterCustomerCleanup(options) {
  options = swExactMasterCustomerCleanupOptions_(options);
  var preview = sw_previewTestDataCleanupOnce(options);
  if (!preview || preview.ok === false) return preview || { ok: false, message: 'preview failed' };

  swWriteTestDataCleanupCandidatesTab_(preview, SW_EXACT_MASTER_CUSTOMER_CLEANUP_PREVIEW_TAB_NAME_);
  var apiRead = swReadTestDataCleanupPreviewViaSheetsApi_(preview);
  var audit = swAuditExactMasterCustomerCleanupPreviewRows_(apiRead, options.exactMasterCustomerNames);

  return {
    ok: true,
    apply: false,
    createdAt: preview.createdAt,
    workflowSpreadsheetId: preview.workflowSpreadsheetId,
    workflowSpreadsheetName: preview.workflowSpreadsheetName,
    requestedCustomerNames: options.exactMasterCustomerNames,
    totalCandidateRows: preview.totalCandidateRows,
    confirmationToken: preview.confirmationToken,
    fingerprint: preview.fingerprint,
    reviewSheet: preview.reviewSheet || null,
    sourceSummary: preview.summary ? preview.summary.sources : [],
    skippedSources: preview.summary ? preview.summary.skippedSources : [],
    warnings: preview.summary ? preview.summary.warnings : [],
    errors: preview.summary ? preview.summary.errors : [],
    matchTypeCounts: audit.matchTypeCounts,
    requestedNameCounts: audit.requestedNameCounts,
    directMasterRows: audit.directMasterRows,
    previewRows: audit.previewRows,
    sheetsApi: {
      ok: apiRead.ok,
      status: apiRead.status,
      range: apiRead.range,
      rowCount: apiRead.rowCount,
      error: apiRead.error || ''
    }
  };
}

function sw_previewExactMasterCustomerCleanupJson(options) {
  return JSON.stringify(sw_previewExactMasterCustomerCleanup(options || {}), null, 2);
}

function sw_applyExactMasterCustomerCleanup(options) {
  options = swExactMasterCustomerCleanupOptions_(options);
  return sw_applyTestDataCleanupOnce(options);
}

function swPreviewTestDataCleanupForMenu_(ui, preview) {
  var sheetName = (preview && preview.reviewSheet && preview.reviewSheet.sheetName) || SW_TEST_DATA_CLEANUP_PREVIEW_TAB_NAME_;
  var rowCount = (preview && preview.reviewSheet && preview.reviewSheet.rowCount) || 0;
  ui.alert(
    'Pre-commit candidate inspection',
    'All candidate rows were written to tab: ' + sheetName +
    ' (' + rowCount + ' row(s)).\nPlease review this sheet before confirming cleanup.',
    ui.ButtonSet.OK
  );
  var confirm = ui.alert(
    'Pre-commit',
    'Proceed to cleanup execution after inspecting these candidates?',
    ui.ButtonSet.YES_NO
  );
  return confirm === ui.Button.YES;
}

function swOpenTestDataCleanupCandidatesSheet_(preview) {
  var review = preview && preview.reviewSheet ? preview.reviewSheet : {};
  var sheetName = review.sheetName || SW_TEST_DATA_CLEANUP_PREVIEW_TAB_NAME_;
  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  var sh = null;
  try {
    sh = ss.getSheetByName(sheetName);
  } catch (_) {}
  if (!sh) return null;
  try {
    ss.setActiveSheet(sh);
    ss.setActiveRange(sh.getRange(1, 1));
  } catch (_) {}
  return {
    sheetName: sheetName,
    rowCount: review.rowCount || 0,
    url: review.url || (ss.getUrl() + '#gid=' + sh.getSheetId())
  };
}

function swTestDataCleanupMenuSummary_(preview) {
  var lines = [
    'Pre-commit preview for one-time test data cleanup',
    'Total matched rows: ' + (preview.totalCandidateRows || 0),
    'Workflow file: ' + (preview.workflowSpreadsheetName || preview.workflowSpreadsheetId || '(unknown)'),
    ''
  ];
  var sources = (preview.summary && preview.summary.sources) || [];
  lines.push('Source candidates:');
  for (var i = 0; i < sources.length; i++) {
    var source = sources[i];
    lines.push('  ' + source.workbookKey + ' / ' + source.sheetName + ': ' + source.candidates + ' row(s)');
  }
  if (!sources.length) lines.push('  (none)');
  lines.push('', 'Top candidate sample:');
  var sampleCount = Math.min(15, (preview.candidates || []).length);
  var sample = (preview.candidates || []).slice(0, sampleCount);
  for (var j = 0; j < sample.length; j++) {
    var row = sample[j];
    var identity = swTestDataCleanupCandidateDisplayIdentity_(row);
    var match = swTestDataCleanupTrim_(row.matchedField && row.matchedValue ? (row.matchedField + ': ' + row.matchedValue) : (row.matchedField || row.reason || 'match'));
    lines.push(
      '  - ' + (j + 1) + ') ' + row.sheetName + ' | row ' + row.rowNumber +
      ' | Name: ' + (identity.customerName || '(unknown)') +
      ' | Email: ' + (identity.email || '(unknown)') +
      ' | ' + match
    );
  }
  if ((preview.candidates || []).length > sampleCount) {
    lines.push('  ... and ' + ((preview.candidates.length - sampleCount) + ' more'));
  }
  return lines.join('\n');
}

function swWriteTestDataCleanupCandidatesTab_(preview, tabName) {
  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  tabName = swTestDataCleanupTrim_(tabName || SW_TEST_DATA_CLEANUP_PREVIEW_TAB_NAME_);
  var sh = ss.getSheetByName(tabName);
  if (!sh) {
    sh = ss.insertSheet(tabName);
  } else {
    sh.clear();
  }

  if (!preview || !preview.candidates || !preview.candidates.length) {
    sh.getRange(1, 1).setValue('No test-data cleanup candidates found.');
    preview.reviewSheet = {
      ok: true,
      sheetName: tabName,
      rowCount: 0,
      url: ss.getUrl() + '#gid=' + sh.getSheetId()
    };
    try { SpreadsheetApp.flush(); } catch (_) {}
    return preview.reviewSheet;
  }

  var headers = [
    'Created At',
    'Workflow Spreadsheet',
    'Workbook Key',
    'Spreadsheet',
    'Sheet',
    'Row',
    'Match Type',
    'Matched Field',
    'Matched Value',
    'Reason',
    'Customer Name',
    'Email',
    'Phone',
    'RootApptID',
    'APPT_ID',
    'CalendlyEventUID',
    'TaskID',
    'SO#',
    'Brand',
    'Visit Date',
    'Status',
    'Payment ID',
    'Doc Number',
    'Seed Match Type',
    'Seed Matched Field',
    'Seed Matched Value',
    'Seed Reason'
  ];
  var now = swTestDataCleanupIso_(new Date());
  var rows = (preview.candidates || []).map(function (row) {
    var identity = swTestDataCleanupCandidateDisplayIdentity_(row);
    return [
      now,
      preview.workflowSpreadsheetName || preview.workflowSpreadsheetId || '',
      row.workbookKey || '',
      row.spreadsheetName || '',
      row.sheetName || '',
      row.rowNumber || '',
      row.matchType || '',
      row.matchedField || '',
      row.matchedValue || '',
      row.reason || '',
      identity.customerName || '',
      identity.email || '',
      row.phone || '',
      row.root || '',
      row.appt || '',
      row.uid || '',
      row.taskId || '',
      row.so || '',
      row.brand || '',
      row.visitDate || '',
      row.status || '',
      row.paymentId || '',
      row.docNumber || '',
      row.seedMatchType || '',
      row.seedMatchedField || '',
      row.seedMatchedValue || '',
      row.seedReason || ''
    ];
  });
  if (rows.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sh.autoResizeColumns(1, headers.length);
    sh.setFrozenRows(1);
  }

  preview.reviewSheet = {
    ok: true,
    sheetName: tabName,
    rowCount: rows.length,
    url: ss.getUrl() + '#gid=' + sh.getSheetId()
  };
  try { SpreadsheetApp.flush(); } catch (_) {}
  return preview.reviewSheet;
}

function sw_runTestDataCleanupPreviewAudit(options) {
  options = options || {};
  var preview = sw_previewTestDataCleanupOnce(options);
  if (!preview || preview.ok === false) return preview || { ok: false, message: 'preview failed' };

  swWriteTestDataCleanupCandidatesTab_(preview);
  var apiRead = swReadTestDataCleanupPreviewViaSheetsApi_(preview);
  var audit = swAuditTestDataCleanupPreviewRows_(preview, apiRead);

  return {
    ok: true,
    createdAt: preview.createdAt,
    workflowSpreadsheetId: preview.workflowSpreadsheetId,
    workflowSpreadsheetName: preview.workflowSpreadsheetName,
    totalCandidateRows: preview.totalCandidateRows,
    reviewSheet: preview.reviewSheet || null,
    sourceSummary: preview.summary ? preview.summary.sources : [],
    matchTypeCounts: audit.matchTypeCounts,
    apTarget: audit.apTarget,
    expectedTaskIds: audit.expectedTaskIds,
    suspiciousDirectRows: audit.suspiciousDirectRows,
    suspiciousLinkedRows: audit.suspiciousLinkedRows,
    previewRows: audit.previewRows,
    sheetsApi: {
      ok: apiRead.ok,
      status: apiRead.status,
      range: apiRead.range,
      rowCount: apiRead.rowCount,
      error: apiRead.error || ''
    }
  };
}

function sw_traceTestDemoCustomerCleanupPreview(options) {
  options = options || {};
  swRequireTestDataCleanupAdmin_(options);

  var names = options.names || ['testdemo3', 'testdemo34', 'testdemo5'];
  var needles = names.map(function (name) { return swTestDataCleanupTrim_(name); }).filter(Boolean);
  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  var workflowTargetBase = {
    spreadsheetId: swTestDataCleanupSpreadsheetId_(ss),
    spreadsheetName: ss.getName()
  };
  var masterTarget = [{
    spreadsheetId: workflowTargetBase.spreadsheetId,
    spreadsheetName: workflowTargetBase.spreadsheetName,
    sheetName: '00_Master Appointments'
  }];
  var sourceTargets = SW_TEST_DATA_CLEANUP_WORKFLOW_SOURCE_ORDER_.map(function (sheetName) {
    return {
      spreadsheetId: workflowTargetBase.spreadsheetId,
      spreadsheetName: workflowTargetBase.spreadsheetName,
      sheetName: sheetName
    };
  });
  var paymentTarget = swTestDataCleanupPaymentLedgerApiTarget_();
  if (paymentTarget) sourceTargets.push(paymentTarget);

  var masterMentions = swSheetsApiSearchTargets_(masterTarget, needles, {
    maxMatchesPerNeedle: Number(options.maxMatchesPerNeedle || 100)
  });
  var nameMentions = swSheetsApiSearchTargets_(sourceTargets, needles, {
    maxMatchesPerNeedle: Number(options.maxMatchesPerNeedle || 100)
  });

  var preview = sw_runTestDataCleanupPreviewAudit(options.previewOptions || {});
  var demoPreviewRows = swFilterTestDemoCleanupPreviewRows_(preview.previewRows || [], needles);
  var linkedNeedles = swCollectTestDemoCleanupLinkedNeedles_(demoPreviewRows);
  var linkedMentions = linkedNeedles.length
    ? swSheetsApiSearchTargets_(sourceTargets, linkedNeedles, {
        maxMatchesPerNeedle: Number(options.maxLinkedMatchesPerNeedle || 100)
      })
    : {};

  return {
    ok: true,
    names: needles,
    masterAppointments: masterMentions,
    nameMentions: nameMentions,
    linkedNeedles: linkedNeedles,
    linkedMentions: linkedMentions,
    preview: {
      ok: preview.ok,
      totalCandidateRows: preview.totalCandidateRows,
      reviewSheet: preview.reviewSheet,
      sourceSummary: preview.sourceSummary,
      matchTypeCounts: preview.matchTypeCounts,
      suspiciousDirectRows: preview.suspiciousDirectRows,
      suspiciousLinkedRows: preview.suspiciousLinkedRows,
      sheetsApi: preview.sheetsApi,
      demoRows: demoPreviewRows
    }
  };
}

function sw_traceTestDemoCustomerCleanupPreviewSummaryJson(options) {
  var result = sw_traceTestDemoCustomerCleanupPreview(options || {});
  function counts(searchResult) {
    var out = {};
    Object.keys(searchResult || {}).sort().forEach(function (key) {
      out[key] = (searchResult[key] && searchResult[key].count) || 0;
    });
    return out;
  }
  function rows(items) {
    return (items || []).map(function (row) {
      return {
        sheet: row.sheet || '',
        row: row.row || '',
        matchType: row.matchType || '',
        matchedField: row.matchedField || '',
        matchedValue: row.matchedValue || '',
        customerName: row.customerName || '',
        email: row.email || '',
        root: row.root || '',
        appt: row.appt || '',
        taskId: row.taskId || '',
        seedMatchType: row.seedMatchType || '',
        seedMatchedField: row.seedMatchedField || '',
        seedMatchedValue: row.seedMatchedValue || ''
      };
    });
  }
  return JSON.stringify({
    ok: result.ok,
    names: result.names,
    masterAppointmentMentionCounts: counts(result.masterAppointments),
    nameMentionCounts: counts(result.nameMentions),
    linkedNeedles: result.linkedNeedles,
    linkedMentionCounts: counts(result.linkedMentions),
    preview: {
      ok: result.preview.ok,
      totalCandidateRows: result.preview.totalCandidateRows,
      reviewSheet: result.preview.reviewSheet,
      sourceSummary: result.preview.sourceSummary,
      matchTypeCounts: result.preview.matchTypeCounts,
      suspiciousDirectRows: result.preview.suspiciousDirectRows,
      suspiciousLinkedRows: result.preview.suspiciousLinkedRows,
      sheetsApi: result.preview.sheetsApi,
      rows: rows(result.preview.demoRows)
    }
  }, null, 2);
}

function sw_verifyTestDemoCleanupPostApplyJson(options) {
  options = options || {};
  swRequireTestDataCleanupAdmin_(options);

  var names = options.names || ['testdemo3', 'testdemo34', 'testdemo5'];
  var needles = names.map(function (name) { return swTestDataCleanupTrim_(name); }).filter(Boolean);
  var linkedNeedles = options.linkedNeedles || [
    'AP-20260505-003',
    'AP-20260505-005',
    'AP-20260505-006',
    'SW|AP-20260505-003|AP-20260505-003|ASSIGN_APPOINTMENT',
    'SW|AP-20260505-003|AP-20260505-003|SEND_HYBRID_WELCOME',
    'SW|AP-20260505-003|AP-20260505-003|REVIEW_APPOINTMENT',
    'SW|AP-20260505-003|AP-20260505-003|APPOINTMENT_DAY_CHECKLIST',
    'SW|AP-20260505-005|AP-20260505-005|ASSIGN_APPOINTMENT',
    'SW|AP-20260505-005|AP-20260505-005|SEND_HYBRID_WELCOME',
    'SW|AP-20260505-005|AP-20260505-005|REVIEW_APPOINTMENT',
    'SW|AP-20260505-005|AP-20260505-005|APPOINTMENT_DAY_CHECKLIST',
    'SW|AP-20260505-003|AP-20260505-006|ASSIGN_APPOINTMENT',
    'SW|AP-20260505-003|AP-20260505-006|SEND_HYBRID_WELCOME',
    'SW|AP-20260505-003|AP-20260505-006|REVIEW_APPOINTMENT',
    'SW|AP-20260505-003|AP-20260505-006|APPOINTMENT_DAY_CHECKLIST',
    'SW|AP-20260505-003|AP-20260505-003|SEND_WELCOME',
    'SW|AP-20260505-003|AP-20260505-003|SEND_MAP_INSTRUCTIONS',
    'SW|AP-20260505-005|AP-20260505-005|SEND_WELCOME',
    'SW|AP-20260505-005|AP-20260505-005|SEND_MAP_INSTRUCTIONS',
    'SW|AP-20260505-003|AP-20260505-006|SEND_WELCOME',
    'SW|AP-20260505-003|AP-20260505-006|SEND_MAP_INSTRUCTIONS'
  ];

  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  var workflowTargetBase = {
    spreadsheetId: swTestDataCleanupSpreadsheetId_(ss),
    spreadsheetName: ss.getName()
  };
  var master = swReadTestDemoMasterAppointmentRowsViaSheetsApi_(workflowTargetBase.spreadsheetId, needles);
  var downstreamTargets = SW_TEST_DATA_CLEANUP_WORKFLOW_SOURCE_ORDER_.filter(function (sheetName) {
    return sheetName !== '00_Master Appointments';
  }).map(function (sheetName) {
    return {
      spreadsheetId: workflowTargetBase.spreadsheetId,
      spreadsheetName: workflowTargetBase.spreadsheetName,
      sheetName: sheetName
    };
  });
  var paymentTarget = swTestDataCleanupPaymentLedgerApiTarget_();
  if (paymentTarget) downstreamTargets.push(paymentTarget);
  var downstreamNameMentions = swSheetsApiSearchTargets_(downstreamTargets, needles, {
    maxMatchesPerNeedle: Number(options.maxMatchesPerNeedle || 50)
  });
  var downstreamLinkedMentions = swSheetsApiSearchTargets_(downstreamTargets, linkedNeedles, {
    maxMatchesPerNeedle: Number(options.maxMatchesPerNeedle || 50)
  });
  var preview = sw_runTestDataCleanupPreviewAudit(options.previewOptions || {});

  return JSON.stringify({
    ok: true,
    masterAppointments: master,
    downstreamNameMentionCounts: swTestDataCleanupSearchCounts_(downstreamNameMentions),
    downstreamLinkedMentionCounts: swTestDataCleanupSearchCounts_(downstreamLinkedMentions),
    preview: {
      ok: preview.ok,
      totalCandidateRows: preview.totalCandidateRows,
      sourceSummary: preview.sourceSummary,
      matchTypeCounts: preview.matchTypeCounts,
      suspiciousDirectRows: preview.suspiciousDirectRows,
      suspiciousLinkedRows: preview.suspiciousLinkedRows,
      sheetsApi: preview.sheetsApi
    }
  }, null, 2);
}

function sw_verifyTestDemoCleanupPostApplyDetailJson(options) {
  options = options || {};
  swRequireTestDataCleanupAdmin_(options);

  var names = options.names || ['testdemo3', 'testdemo34', 'testdemo5'];
  var needles = names.map(function (name) { return swTestDataCleanupTrim_(name); }).filter(Boolean);
  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  var workflowTargetBase = {
    spreadsheetId: swTestDataCleanupSpreadsheetId_(ss),
    spreadsheetName: ss.getName()
  };
  var downstreamTargets = SW_TEST_DATA_CLEANUP_WORKFLOW_SOURCE_ORDER_.filter(function (sheetName) {
    return sheetName !== '00_Master Appointments';
  }).map(function (sheetName) {
    return {
      spreadsheetId: workflowTargetBase.spreadsheetId,
      spreadsheetName: workflowTargetBase.spreadsheetName,
      sheetName: sheetName
    };
  });
  var paymentTarget = swTestDataCleanupPaymentLedgerApiTarget_();
  if (paymentTarget) downstreamTargets.push(paymentTarget);
  var downstreamNameMentions = swSheetsApiSearchTargets_(downstreamTargets, needles, {
    maxMatchesPerNeedle: Number(options.maxMatchesPerNeedle || 50)
  });
  return JSON.stringify({
    ok: true,
    downstreamNameMentions: downstreamNameMentions
  }, null, 2);
}

function sw_deleteResidualTestDemoIntakeQueueRowsJson(options) {
  options = options || {};
  swRequireTestDataCleanupAdmin_(options);

  var names = options.names || ['testdemo3', 'testdemo34', 'testdemo5'];
  var needles = names.map(function (name) { return swTestDataCleanupTrim_(name); }).filter(Boolean);
  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  var target = {
    spreadsheetId: swTestDataCleanupSpreadsheetId_(ss),
    spreadsheetName: ss.getName(),
    sheetName: '_IntakeQueue'
  };
  var mentions = swSheetsApiSearchTargets_([target], needles, {
    maxMatchesPerNeedle: Number(options.maxMatchesPerNeedle || 50)
  });
  var rowSet = {};
  Object.keys(mentions || {}).forEach(function (needle) {
    var matches = (mentions[needle] && mentions[needle].matches) || [];
    for (var i = 0; i < matches.length; i++) {
      var rowNo = Number(matches[i].row);
      if (rowNo > 1) rowSet[rowNo] = true;
    }
  });
  var rows = Object.keys(rowSet).map(Number).sort(function (a, b) { return b - a; });
  var sh = ss.getSheetByName('_IntakeQueue');
  if (!sh) throw new Error('Missing sheet: _IntakeQueue');
  for (var r = 0; r < rows.length; r++) sh.deleteRow(rows[r]);
  try { SpreadsheetApp.flush(); } catch (_) {}
  return JSON.stringify({
    ok: true,
    sheetName: '_IntakeQueue',
    deletedRowsDescending: rows,
    deletedCount: rows.length,
    mentions: mentions
  }, null, 2);
}

function swReadTestDataCleanupPreviewViaSheetsApi_(preview) {
  preview = preview || {};
  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  var ssId = swTestDataCleanupSpreadsheetId_(ss);
  var sheetName = (preview.reviewSheet && preview.reviewSheet.sheetName) || SW_TEST_DATA_CLEANUP_PREVIEW_TAB_NAME_;
  var a1 = "'" + sheetName.replace(/'/g, "''") + "'!A1:AA";
  try {
    if (typeof Sheets === 'undefined' || !Sheets || !Sheets.Spreadsheets || !Sheets.Spreadsheets.Values) {
      return {
        ok: false,
        status: 0,
        range: '',
        rowCount: 0,
        values: [],
        error: 'Sheets advanced service is not available in this Apps Script project.'
      };
    }
    var res = Sheets.Spreadsheets.Values.get(ssId, a1, { majorDimension: 'ROWS' });
    var values = (res && res.values) || [];
    return {
      ok: true,
      status: 200,
      range: (res && res.range) || '',
      rowCount: Math.max(0, values.length - 1),
      values: values,
      error: ''
    };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      range: '',
      rowCount: 0,
      values: [],
      error: err && err.message ? err.message : String(err)
    };
  }
}

function swAuditTestDataCleanupPreviewRows_(preview, apiRead) {
  var values = (apiRead && apiRead.values) || [];
  var headers = values.length ? values[0] : [];
  var rows = values.length > 1 ? values.slice(1) : [];
  var H = {};
  for (var i = 0; i < headers.length; i++) H[String(headers[i] || '')] = i;

  function cell(row, name) {
    var idx = H[name];
    return idx == null ? '' : swTestDataCleanupTrim_(row[idx]);
  }

  var matchTypeCounts = {};
  var apNeedle = 'AP-20260504-002';
  var apRows = [];
  var expectedTasks = [
    'SW|AP-20260504-002|AP-20260504-002|ASSIGN_APPOINTMENT',
    'SW|AP-20260504-002|AP-20260504-002|SEND_HYBRID_WELCOME',
    'SW|AP-20260504-002|AP-20260504-002|REVIEW_APPOINTMENT',
    'SW|AP-20260504-002|AP-20260504-002|APPOINTMENT_DAY_CHECKLIST'
  ];
  var expectedTaskFound = {};
  for (var e = 0; e < expectedTasks.length; e++) expectedTaskFound[expectedTasks[e]] = false;

  var suspiciousDirectRows = [];
  var suspiciousLinkedRows = [];
  var previewRows = [];
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    var matchType = cell(row, 'Match Type');
    var matchedValue = cell(row, 'Matched Value');
    var customer = cell(row, 'Customer Name');
    var email = cell(row, 'Email');
    var sheet = cell(row, 'Sheet');
    var rowNo = cell(row, 'Row');
    var root = cell(row, 'RootApptID');
    var appt = cell(row, 'APPT_ID');
    var taskId = cell(row, 'TaskID');
    var seedMatchType = cell(row, 'Seed Match Type');
    var seedMatchedField = cell(row, 'Seed Matched Field');
    var seedMatchedValue = cell(row, 'Seed Matched Value');
    var seedReason = cell(row, 'Seed Reason');

    previewRows.push({
      sheet: sheet,
      row: rowNo,
      matchType: matchType,
      matchedField: cell(row, 'Matched Field'),
      matchedValue: matchedValue,
      reason: cell(row, 'Reason'),
      customerName: customer,
      email: email,
      root: root,
      appt: appt,
      taskId: taskId,
      seedMatchType: seedMatchType,
      seedMatchedField: seedMatchedField,
      seedMatchedValue: seedMatchedValue,
      seedReason: seedReason
    });

    matchTypeCounts[matchType || '(blank)'] = (matchTypeCounts[matchType || '(blank)'] || 0) + 1;

    var bundle = [root, appt, taskId].join('|');
    if (bundle.indexOf(apNeedle) >= 0) {
      apRows.push({
        sheet: sheet,
        row: rowNo,
        matchType: matchType,
        matchedValue: matchedValue,
        customerName: customer,
        email: email,
        root: root,
        appt: appt,
        taskId: taskId
      });
    }
    if (expectedTaskFound[taskId] != null) expectedTaskFound[taskId] = true;

    var isDirect = /^direct/.test(matchType);
    var looksTestName = customer ? swTestDataCleanupTextLooksTest_(customer, 'strict') : false;
    var looksTestEmail = email ? swTestDataCleanupEmailLooksTest_(email, 'strict') : false;
    var looksTestMatch = matchedValue ? swTestDataCleanupTextLooksTest_(matchedValue, 'strict') : false;
    var looksTestIdentity = looksTestName || looksTestEmail || looksTestMatch;

    if (isDirect && !looksTestIdentity && suspiciousDirectRows.length < 80) {
      suspiciousDirectRows.push({
        sheet: sheet,
        row: rowNo,
        matchType: matchType,
        matchedValue: matchedValue,
        customerName: customer,
        email: email,
        taskId: taskId,
        root: root,
        appt: appt
      });
    }

    var isKeyLinked = /^(root|appt|uid|taskId|so)$/.test(matchType);
    var idLooksTest = swTestDataCleanupTextLooksTest_(root, 'strict') || swTestDataCleanupTextLooksTest_(appt, 'strict') || swTestDataCleanupTextLooksTest_(taskId, 'strict');
    if (isKeyLinked && !looksTestIdentity && !idLooksTest && suspiciousLinkedRows.length < 120) {
      suspiciousLinkedRows.push({
        sheet: sheet,
        row: rowNo,
        matchType: matchType,
        matchedValue: matchedValue,
        customerName: customer,
        email: email,
        taskId: taskId,
        root: root,
        appt: appt
      });
    }
  }

  var expectedTaskIds = expectedTasks.map(function (taskId) {
    return { taskId: taskId, found: !!expectedTaskFound[taskId] };
  });

  return {
    matchTypeCounts: matchTypeCounts,
    apTarget: {
      rootOrAppt: apNeedle,
      found: apRows.length > 0,
      rows: apRows
    },
    expectedTaskIds: expectedTaskIds,
    suspiciousDirectRows: suspiciousDirectRows,
    suspiciousLinkedRows: suspiciousLinkedRows
    ,
    previewRows: previewRows
  };
}

function swAuditExactMasterCustomerCleanupPreviewRows_(apiRead, names) {
  var values = (apiRead && apiRead.values) || [];
  var headers = values.length ? values[0] : [];
  var rows = values.length > 1 ? values.slice(1) : [];
  var H = {};
  for (var i = 0; i < headers.length; i++) H[String(headers[i] || '')] = i;
  var requested = swExactMasterCustomerCleanupNameSet_(names);
  var requestedCounts = {};
  (names || []).forEach(function (name) {
    requestedCounts[name] = 0;
  });

  function cell(row, name) {
    var idx = H[name];
    return idx == null ? '' : swTestDataCleanupTrim_(row[idx]);
  }

  var matchTypeCounts = {};
  var directMasterRows = [];
  var previewRows = [];
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r] || [];
    var item = {
      sheet: cell(row, 'Sheet'),
      row: cell(row, 'Row'),
      matchType: cell(row, 'Match Type'),
      matchedField: cell(row, 'Matched Field'),
      matchedValue: cell(row, 'Matched Value'),
      reason: cell(row, 'Reason'),
      customerName: cell(row, 'Customer Name'),
      email: cell(row, 'Email'),
      root: cell(row, 'RootApptID'),
      appt: cell(row, 'APPT_ID'),
      uid: cell(row, 'CalendlyEventUID'),
      taskId: cell(row, 'TaskID'),
      so: cell(row, 'SO#'),
      seedMatchType: cell(row, 'Seed Match Type'),
      seedMatchedField: cell(row, 'Seed Matched Field'),
      seedMatchedValue: cell(row, 'Seed Matched Value'),
      seedReason: cell(row, 'Seed Reason')
    };
    previewRows.push(item);
    matchTypeCounts[item.matchType || '(blank)'] = (matchTypeCounts[item.matchType || '(blank)'] || 0) + 1;

    var nameKey = swExactMasterCustomerCleanupNameKey_(item.customerName || item.matchedValue || item.seedMatchedValue);
    var requestedName = requested[nameKey];
    if (requestedName) requestedCounts[requestedName] = (requestedCounts[requestedName] || 0) + 1;

    if (item.sheet === '00_Master Appointments' && item.matchType === 'directMasterCustomerName') {
      directMasterRows.push(item);
    }
  }

  return {
    matchTypeCounts: matchTypeCounts,
    requestedNameCounts: requestedCounts,
    directMasterRows: directMasterRows,
    previewRows: previewRows
  };
}

function swExactMasterCustomerCleanupOptions_(options) {
  if (Array.isArray(options)) options = { names: options };
  options = options || {};
  var out = {};
  Object.keys(options).forEach(function (key) { out[key] = options[key]; });
  var names = out.exactMasterCustomerNames || out.customerNames || out.names || SW_EXACT_MASTER_CUSTOMER_CLEANUP_DEFAULT_NAMES_;
  out.exactMasterCustomerNames = swUniqueStrings_(names);
  out.seedMode = 'exactMasterCustomerNames';
  out.matchMode = 'exactMasterCustomerNames';
  return out;
}

function swExactMasterCustomerCleanupNameSet_(names) {
  var out = {};
  (names || []).forEach(function (name) {
    name = swTestDataCleanupTrim_(name);
    var key = swExactMasterCustomerCleanupNameKey_(name);
    if (key) out[key] = name;
  });
  return out;
}

function swExactMasterCustomerCleanupNameKey_(value) {
  return swTestDataCleanupNorm_(value).replace(/\s+/g, ' ');
}

function sw_inspectTestDataNeedle(options) {
  options = options || {};
  swRequireTestDataCleanupAdmin_(options);

  var needle = swTestDataCleanupTrim_(options.needle || 'AP-20260504-002');
  var taskIds = options.taskIds || [
    'SW|AP-20260504-002|AP-20260504-002|ASSIGN_APPOINTMENT',
    'SW|AP-20260504-002|AP-20260504-002|SEND_HYBRID_WELCOME',
    'SW|AP-20260504-002|AP-20260504-002|REVIEW_APPOINTMENT',
    'SW|AP-20260504-002|AP-20260504-002|APPOINTMENT_DAY_CHECKLIST'
  ];
  var taskSet = {};
  for (var t = 0; t < taskIds.length; t++) taskSet[taskIds[t]] = true;

  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  var sheetNames = swWorkflowTestDataCleanupSheetNames_(options);
  var outRows = [];

  for (var i = 0; i < sheetNames.length; i++) {
    var name = sheetNames[i];
    var sh = ss.getSheetByName(name);
    if (!sh) continue;
    var lr = sh.getLastRow();
    var lc = sh.getLastColumn();
    if (lr < 2 || lc < 1) continue;

    var headerRows = 1;
    var headerInfo = swReadTestDataCleanupHeaders_(sh, headerRows, lc);
    var columns = swTestDataCleanupColumns_(headerInfo);
    var values = sh.getRange(2, 1, lr - 1, lc).getDisplayValues();
    for (var r = 0; r < values.length; r++) {
      var row = values[r];
      var rowNo = r + 2;
      var rec = swTestDataCleanupRecordFromRow_(row, columns);
      var joined = row.join('\u001f');
      var hasNeedle = needle && joined.indexOf(needle) >= 0;
      var hasTask = rec.taskId && taskSet[rec.taskId];
      if (!hasNeedle && !hasTask) continue;

      var target = { sheetName: name };
      var direct = swTestDataCleanupDirectMatch_(rec, target, options);
      outRows.push({
        sheetName: name,
        rowNumber: rowNo,
        root: rec.root,
        appt: rec.appt,
        taskId: rec.taskId,
        customerName: rec.customerName,
        email: rec.email,
        so: rec.so,
        payloadJsonSnippet: swTestDataCleanupTrim_(rec.payloadJson).slice(0, 180),
        directMatched: !!direct.matched,
        directMatchType: direct.matchType || '',
        directMatchedField: direct.matchedField || '',
        directMatchedValue: direct.matchedValue || '',
        directReason: direct.reason || '',
        containsNeedleRaw: hasNeedle,
        containsTaskId: !!hasTask
      });
    }
  }

  outRows.sort(function (a, b) {
    if (a.sheetName !== b.sheetName) return a.sheetName < b.sheetName ? -1 : 1;
    return a.rowNumber - b.rowNumber;
  });

  return {
    ok: true,
    needle: needle,
    expectedTaskIds: taskIds,
    foundRows: outRows,
    foundCount: outRows.length
  };
}

function sw_inspectNeedleAllSheets(options) {
  options = options || {};
  swRequireTestDataCleanupAdmin_(options);
  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  var needle = swTestDataCleanupTrim_(options.needle || 'AP-20260504-002');
  var descNeedle = swTestDataCleanupTrim_(options.descNeedle || '"desc":"Test"');
  var out = [];
  var sheets = ss.getSheets();

  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var lr = sh.getLastRow();
    var lc = sh.getLastColumn();
    if (lr < 1 || lc < 1) continue;
    var values = sh.getRange(1, 1, lr, lc).getDisplayValues();
    for (var r = 0; r < values.length; r++) {
      var row = values[r];
      var joined = row.join('\u001f');
      if (needle && joined.indexOf(needle) < 0) continue;
      var hasDescTest = descNeedle ? joined.indexOf(descNeedle) >= 0 : false;
      var sampleCells = [];
      for (var c = 0; c < row.length; c++) {
        var cell = swTestDataCleanupTrim_(row[c]);
        if (!cell) continue;
        if (cell.indexOf(needle) >= 0 || (descNeedle && cell.indexOf(descNeedle) >= 0)) {
          sampleCells.push({
            col: c + 1,
            value: cell.slice(0, 180)
          });
          if (sampleCells.length >= 6) break;
        }
      }
      out.push({
        sheetName: sh.getName(),
        rowNumber: r + 1,
        hasDescTest: hasDescTest,
        sampleCells: sampleCells
      });
      if (out.length >= 300) {
        return { ok: true, needle: needle, descNeedle: descNeedle, rows: out, truncated: true };
      }
    }
  }

  return { ok: true, needle: needle, descNeedle: descNeedle, rows: out, truncated: false };
}

function sw_inspectSheetRowCells(options) {
  options = options || {};
  swRequireTestDataCleanupAdmin_(options);
  var sheetName = swTestDataCleanupTrim_(options.sheetName || '');
  var rowNumber = Number(options.rowNumber || 1);
  if (!sheetName || !isFinite(rowNumber) || rowNumber < 1) {
    return { ok: false, message: 'sheetName and valid rowNumber are required.' };
  }
  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return { ok: false, message: 'sheet not found: ' + sheetName };
  var lc = sh.getLastColumn();
  var headerRows = rowNumber >= 3 ? 2 : 1;
  var headerInfo = swReadTestDataCleanupHeaders_(sh, headerRows, lc);
  var row = sh.getRange(rowNumber, 1, 1, lc).getDisplayValues()[0];

  var cells = [];
  for (var c = 0; c < row.length; c++) {
    var value = swTestDataCleanupTrim_(row[c]);
    if (!value) continue;
    cells.push({
      col: c + 1,
      header: headerInfo.headers[c] || '',
      value: value
    });
  }
  return {
    ok: true,
    sheetName: sheetName,
    rowNumber: rowNumber,
    lastColumn: lc,
    cells: cells
  };
}

function sw_verifyTestDataCleanupPostApply(options) {
  options = options || {};
  swRequireTestDataCleanupAdmin_(options);

  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  var workflowId = swTestDataCleanupSpreadsheetId_(ss);
  var workflowTabs = [
    '00_Master Appointments',
    '02_Form_Inbox',
    '_ExternalBookingEvents',
    '_IntakeQueue',
    '_SalesTaskQueue',
    '_SalesTaskLog',
    '_AppointmentArtifacts',
    '_SalesDataCleanup',
    '03_Client_Status_Log',
    '05_Wax_Requests',
    '07_Root_Index'
  ];
  var sourceTargets = workflowTabs.map(function (sheetName) {
    return { spreadsheetId: workflowId, spreadsheetName: swTestDataCleanupSpreadsheetName_(ss), sheetName: sheetName };
  });

  var paymentTarget = swTestDataCleanupPaymentLedgerApiTarget_();
  if (paymentTarget) sourceTargets.push(paymentTarget);

  var deletedNeedles = [
    'AP-20260504-002',
    'E2E_CLEANUP_20260506_175656_1FB8C3A0',
    'test.cleanup.e2e.20260506_175656.1fb8c3a0@example.com'
  ];
  var realCustomerNeedles = [
    'Trang Nguyen',
    'trangn103@gmail.com',
    'Dat',
    'trungdat.lee@gmail.com',
    'Jacqueline Truong',
    'jacqueline1367@gmail.com'
  ];

  return {
    ok: true,
    sourceDeletedNeedles: swSheetsApiSearchTargets_(sourceTargets, deletedNeedles, { maxMatches: 120 }),
    sourceRealCustomerNeedles: swSheetsApiSearchTargets_(sourceTargets, realCustomerNeedles, { maxMatches: 120 }),
    targets: sourceTargets.map(function (t) {
      return { spreadsheetName: t.spreadsheetName, sheetName: t.sheetName };
    })
  };
}

function swTestDataCleanupPaymentLedgerApiTarget_() {
  try {
    var target = null;
    if (typeof rp_getLedgerTarget === 'function') {
      target = rp_getLedgerTarget();
    } else if (typeof pr_getLedger_ === 'function' && typeof pr_getPaymentsSheet_ === 'function') {
      var ledger = pr_getLedger_();
      target = { ss: ledger, sh: pr_getPaymentsSheet_(ledger) };
    }
    if (!target || !target.sh) return null;
    var ledgerSafety = swPaymentLedgerTargetSheetIsSafe_(target);
    if (!ledgerSafety.ok) return null;
    var ss = target.ss || target.sh.getParent();
    return {
      spreadsheetId: swTestDataCleanupSpreadsheetId_(ss),
      spreadsheetName: swTestDataCleanupSpreadsheetName_(ss),
      sheetName: target.sh.getName()
    };
  } catch (_) {}
  return null;
}

function swSheetsApiSearchTargets_(targets, needles, options) {
  options = options || {};
  var maxMatches = Number(options.maxMatches || 100);
  var out = {};
  needles.forEach(function (needle) {
    out[needle] = { count: 0, matches: [] };
  });
  for (var i = 0; i < targets.length; i++) {
    var target = targets[i];
    var values = swSheetsApiReadValues_(target.spreadsheetId, target.sheetName, 'A1:ZZ');
    if (!values.ok) {
      Object.keys(out).forEach(function (needle) {
        out[needle].matches.push({
          spreadsheetName: target.spreadsheetName,
          sheetName: target.sheetName,
          error: values.error || 'read failed'
        });
      });
      continue;
    }
    for (var r = 0; r < values.values.length; r++) {
      var row = values.values[r] || [];
      for (var c = 0; c < row.length; c++) {
        var text = swTestDataCleanupTrim_(row[c]);
        if (!text) continue;
        needles.forEach(function (needle) {
          if (text.indexOf(needle) < 0) return;
          out[needle].count++;
          if (out[needle].matches.length < maxMatches) {
            out[needle].matches.push({
              spreadsheetName: target.spreadsheetName,
              sheetName: target.sheetName,
              row: r + 1,
              col: c + 1,
              value: text.slice(0, 180)
            });
          }
        });
      }
    }
  }
  return out;
}

function swSheetsApiReadValues_(spreadsheetId, sheetName, range) {
  var a1 = "'" + String(sheetName || '').replace(/'/g, "''") + "'!" + (range || 'A1:ZZ');
  try {
    if (typeof Sheets === 'undefined' || !Sheets || !Sheets.Spreadsheets || !Sheets.Spreadsheets.Values) {
      return { ok: false, values: [], error: 'Sheets advanced service unavailable' };
    }
    var res = Sheets.Spreadsheets.Values.get(spreadsheetId, a1, { majorDimension: 'ROWS' });
    return { ok: true, values: (res && res.values) || [], range: (res && res.range) || '' };
  } catch (err) {
    return { ok: false, values: [], error: err && err.message ? err.message : String(err) };
  }
}

function swTestDataCleanupCandidateDisplayIdentity_(row) {
  var customer = swTestDataCleanupTrim_(row && row.customerName || '');
  var email = swTestDataCleanupTrim_(row && row.email || '');

  if (!customer && row && row.matchedCustomerName) {
    customer = swTestDataCleanupTrim_(row.matchedCustomerName);
  }
  if (!email && row && row.matchedEmail) {
    email = swTestDataCleanupTrim_(row.matchedEmail);
  }
  if (!customer && row && row.matchType === 'directName' && row.matchedValue) {
    customer = swTestDataCleanupTrim_(row.matchedValue);
  }
  if (!email && row && row.matchType === 'directEmail' && row.matchedValue) {
    email = swTestDataCleanupTrim_(row.matchedValue);
  }
  if (!email && row && row.reason) {
    var inferred = swTestDataCleanupNormalizeMatchReason_(row.reason);
    if (inferred && swTestDataCleanupLooksLikeEmailValue_(inferred)) {
      email = inferred;
    }
  }

  return {
    customerName: customer,
    email: email
  };
}

function swTestDataCleanupNormalizeMatchReason_(reason) {
  var match = String(reason || '').match(/'([^']+)'/);
  if (match && match[1]) return match[1];
  var emailParts = String(reason || '').match(/([\w.+-]+@[\w.-]+\.[A-Za-z]{2,})/);
  if (emailParts && emailParts[1]) return emailParts[1];
  return '';
}

function swBuildTestDataCleanupPlan_(options) {
  options = options || {};
  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  var plan = {
    ok: true,
    apply: options.apply === true,
    createdAt: swTestDataCleanupIso_(new Date()),
    workflowSpreadsheetId: swTestDataCleanupSpreadsheetId_(ss),
    workflowSpreadsheetName: swTestDataCleanupSpreadsheetName_(ss),
    confirmationToken: '',
    fingerprint: '',
    totalCandidateRows: 0,
    deletedCount: 0,
    keyset: swNewTestDataCleanupKeyset_(),
    sources: [],
    skippedSources: [],
    candidates: [],
    deletedRows: [],
    invalidation: null,
    readModelRebuild: null,
    warnings: [],
    errors: []
  };

  swAddTestDataCleanupWarning_(plan, 'Generated _SW_* read-model tabs are not row-deleted. They are invalidated/rebuilt after source cleanup.');
  swAddTestDataCleanupWarning_(plan, 'Payment ledger cleanup uses exact RootApptID/APPT_ID/SO matches from the workflow keyset, not name-only matching.');
  swAddTestDataCleanupWarning_(plan, 'Drive files and folders referenced by artifact rows are not trashed by this function.');

  // Pass 1: only direct test matches seed the keyset.
  swScanAllTestDataCleanupSources_(ss, plan, swTestDataCleanupPassOptions_(options, 'directOnly'));
  // Pass 2: with keyset finalized from direct seeds, pull all downstream linked rows.
  swScanAllTestDataCleanupSources_(ss, plan, swTestDataCleanupPassOptions_(options, 'keyOnly'));

  plan.candidates.sort(swCompareTestDataCleanupCandidates_);
  plan.totalCandidateRows = plan.candidates.length;
  plan.fingerprint = swTestDataCleanupFingerprint_(plan);
  plan.confirmationToken = 'DELETE_TEST_DATA_' + plan.fingerprint.slice(0, 10).toUpperCase();
  plan.summary = swTestDataCleanupSummary_(plan);
  return plan;
}

function swScanAllTestDataCleanupSources_(ss, plan, options) {
  swScanWorkflowTestDataCleanupSources_(ss, plan, options);
  swScanExternalBookingQueueForTestDataCleanup_(ss, plan, options);
  swScanPaymentLedgerForTestDataCleanup_(plan, options);
  swScanDiamond200ForTestDataCleanup_(plan, options);
}

function swTestDataCleanupPassOptions_(options, passMode) {
  var out = {};
  Object.keys(options || {}).forEach(function (key) { out[key] = options[key]; });
  out._scanPassMode = passMode || 'mixed';
  out._recordSkippedSources = passMode !== 'keyOnly';
  return out;
}

function swScanExternalBookingQueueForTestDataCleanup_(workflowSs, plan, options) {
  options = options || {};
  if (options.externalBookingQueue === false) return;
  var id = '';
  try {
    var props = PropertiesService.getScriptProperties();
    id = swTestDataCleanupTrim_(props.getProperty('HPAPP_ACUITY_QUEUE_SPREADSHEET_ID') ||
      props.getProperty('EXTERNAL_BOOKING_QUEUE_SPREADSHEET_ID') || '');
  } catch (_) {}
  if (!id) return;
  if (workflowSs && id === swTestDataCleanupSpreadsheetId_(workflowSs)) return;

  var ss = null;
  try { ss = SpreadsheetApp.openById(id); } catch (err) {
    swAddTestDataCleanupSkippedSource_(plan, {
      workbookKey: 'hpappAcuityQueue',
      sheetName: '_ExternalBookingEvents',
      reason: 'open failed: ' + (err && err.message ? err.message : String(err))
    }, options);
    return;
  }
  var sh = ss.getSheetByName('_ExternalBookingEvents');
  if (!sh) {
    swAddTestDataCleanupSkippedSource_(plan, { workbookKey: 'hpappAcuityQueue', sheetName: '_ExternalBookingEvents', reason: 'missing sheet' }, options);
    return;
  }
  swScanSheetForTestDataCleanup_({
    workbookKey: 'hpappAcuityQueue',
    workbookLabel: 'HPAPP Acuity queue',
    spreadsheet: ss,
    sheet: sh,
    sheetName: sh.getName(),
    headerRows: 1,
    dataStartRow: 2,
    allowDirectTestMatch: true,
    exactKeyOnly: false,
    deleteOrder: swTestDataCleanupDeleteOrder_('_ExternalBookingEvents', 'workflow')
  }, plan, options);
}

function swRequireTestDataCleanupAdmin_(options) {
  options = options || {};
  var token = swTestDataCleanupTrim_(options.authToken || options.token || '');
  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  if (typeof swAuthUserForApi_ !== 'function') {
    throw new Error('Admin access required, but workflow auth helpers are unavailable.');
  }
  var user = swAuthUserForApi_(ss, token);
  if (!user || !user.isAdmin) throw new Error('Admin access required for one-time test data cleanup.');
  return user;
}

function swScanWorkflowTestDataCleanupSources_(ss, plan, options) {
  var sheetNames = swWorkflowTestDataCleanupSheetNames_(options);
  for (var i = 0; i < sheetNames.length; i++) {
    var name = sheetNames[i];
    if (SW_TEST_DATA_CLEANUP_READ_MODEL_SHEETS_[name]) {
      swAddTestDataCleanupSkippedSource_(plan, { workbookKey: 'workflow', sheetName: name, reason: 'generated read model' }, options);
      continue;
    }
    var sh = ss.getSheetByName(name);
    if (!sh) {
      swAddTestDataCleanupSkippedSource_(plan, { workbookKey: 'workflow', sheetName: name, reason: 'missing sheet' }, options);
      continue;
    }
    var target = {
      workbookKey: 'workflow',
      workbookLabel: 'Workflow workbook',
      spreadsheet: ss,
      sheet: sh,
      sheetName: sh.getName(),
      headerRows: 1,
      dataStartRow: 2,
      allowDirectTestMatch: swSheetAllowsDirectTestMatch_(name),
      exactKeyOnly: name === '_SalesTaskLog' || name === '07_Root_Index' || name === '_AppointmentArtifacts',
      deleteOrder: swTestDataCleanupDeleteOrder_(name, 'workflow')
    };
    swScanSheetForTestDataCleanup_(target, plan, options);
  }
}

function swSheetAllowsDirectTestMatch_(sheetName) {
  sheetName = swTestDataCleanupTrim_(sheetName);
  return sheetName === '00_Master Appointments' ||
    sheetName === '02_Form_Inbox' ||
    sheetName === '_ExternalBookingEvents' ||
    sheetName === '_IntakeQueue' ||
    sheetName === '_SalesDataCleanup' ||
    sheetName === '03_Client_Status_Log';
}

function swWorkflowTestDataCleanupSheetNames_(options) {
  var names = SW_TEST_DATA_CLEANUP_WORKFLOW_SOURCE_ORDER_.slice();
  if (typeof SW_SHEETS !== 'undefined') {
    names = [
      SW_SHEETS.MASTER || '00_Master Appointments',
      '02_Form_Inbox',
      SW_SHEETS.EXTERNAL_BOOKING_EVENTS || '_ExternalBookingEvents',
      '_IntakeQueue',
      SW_SHEETS.TASKS || '_SalesTaskQueue',
      SW_SHEETS.LOG || '_SalesTaskLog',
      SW_SHEETS.APPOINTMENT_ARTIFACTS || '_AppointmentArtifacts',
      SW_SHEETS.DATA_CLEANUP || '_SalesDataCleanup',
      '03_Client_Status_Log',
      '05_Wax_Requests',
      '07_Root_Index'
    ];
  }
  if (options && options.workflowSheets && options.workflowSheets.length) {
    names = options.workflowSheets.slice();
  }
  return swUniqueStrings_(names);
}

function swScanPaymentLedgerForTestDataCleanup_(plan, options) {
  if (options && options.includePayments === false) return;
  var target = null;
  try {
    if (typeof rp_getLedgerTarget === 'function') {
      target = rp_getLedgerTarget();
    } else if (typeof pr_getLedger_ === 'function' && typeof pr_getPaymentsSheet_ === 'function') {
      var ledger = pr_getLedger_();
      target = { ss: ledger, sh: pr_getPaymentsSheet_(ledger), resolved: { ledgerFileKey: 'pr_getLedger_', ledgerSheetKey: 'pr_getPaymentsSheet_' } };
    }
  } catch (err) {
    swAddTestDataCleanupSkippedSource_(plan, { workbookKey: 'paymentsLedger', reason: err && err.message ? err.message : String(err) }, options);
    return;
  }
  if (!target || !target.sh) {
    swAddTestDataCleanupSkippedSource_(plan, { workbookKey: 'paymentsLedger', reason: 'payments ledger helper unavailable' }, options);
    return;
  }
  var ledgerSafety = swPaymentLedgerTargetSheetIsSafe_(target);
  if (!ledgerSafety.ok) {
    swAddTestDataCleanupSkippedSource_(plan, { workbookKey: 'paymentsLedger', sheetName: target.sh.getName(), reason: ledgerSafety.reason }, options);
    return;
  }

  swScanSheetForTestDataCleanup_({
    workbookKey: 'paymentsLedger',
    workbookLabel: 'Payments ledger',
    spreadsheet: target.ss || target.sh.getParent(),
    sheet: target.sh,
    sheetName: target.sh.getName(),
    headerRows: 1,
    dataStartRow: 2,
    allowDirectTestMatch: false,
    exactKeyOnly: true,
    deleteOrder: 10,
    note: target.resolved || {}
  }, plan, options);
}

function swPaymentLedgerTargetSheetIsSafe_(target) {
  var actual = '';
  try { actual = target && target.sh ? target.sh.getName() : ''; } catch (_) {}
  if (!actual) return { ok: false, reason: 'payments ledger sheet name unavailable' };

  var expected = 'Payments';
  var configured = false;
  try {
    if (typeof RP_KEY_ALIASES !== 'undefined' && RP_KEY_ALIASES.LEDGER_SHEET_NAME && typeof rp_propOneOf_ === 'function') {
      var sheetRes = rp_propOneOf_(RP_KEY_ALIASES.LEDGER_SHEET_NAME, { required: false, label: 'Payments sheet name' });
      if (sheetRes && sheetRes.value) {
        expected = String(sheetRes.value).trim() || expected;
        configured = true;
      }
    }
  } catch (_) {}

  if (configured && actual !== expected) {
    return { ok: false, reason: 'configured payments sheet "' + expected + '" was not resolved; refusing fallback tab "' + actual + '"' };
  }
  if (!configured && actual !== expected && !/payment/i.test(actual)) {
    return { ok: false, reason: 'default payments tab was not found; refusing fallback tab "' + actual + '"' };
  }
  return { ok: true };
}

function swScanDiamond200ForTestDataCleanup_(plan, options) {
  if (options && options.includeDiamonds === false) return;
  var target = null;
  try {
    if (typeof swDiamond200Target_ === 'function') target = swDiamond200Target_();
  } catch (err) {
    swAddTestDataCleanupSkippedSource_(plan, { workbookKey: 'diamonds200', reason: err && err.message ? err.message : String(err) }, options);
    return;
  }
  if (!target || !target.sheet) {
    swAddTestDataCleanupSkippedSource_(plan, { workbookKey: 'diamonds200', reason: 'diamond workbook helper unavailable' }, options);
    return;
  }

  swScanSheetForTestDataCleanup_({
    workbookKey: 'diamonds200',
    workbookLabel: '200 diamond workbook',
    spreadsheet: target.ss || target.sheet.getParent(),
    sheet: target.sheet,
    sheetName: target.sheet.getName(),
    headerRows: 2,
    dataStartRow: 3,
    allowDirectTestMatch: false,
    exactKeyOnly: true,
    deleteOrder: 20,
    note: { tab: target.tab || target.sheet.getName() }
  }, plan, options);
}

function swScanSheetForTestDataCleanup_(target, plan, options) {
  var sh = target.sheet;
  var lr = sh.getLastRow();
  var lc = sh.getLastColumn();
  var sourceInfo = {
    workbookKey: target.workbookKey,
    workbookLabel: target.workbookLabel,
    spreadsheetId: swTestDataCleanupSpreadsheetId_(target.spreadsheet),
    spreadsheetName: swTestDataCleanupSpreadsheetName_(target.spreadsheet),
    sheetName: target.sheetName,
    sheetId: sh.getSheetId(),
    lastRow: lr,
    lastColumn: lc,
    note: target.note || {}
  };
  swAddTestDataCleanupSource_(plan, sourceInfo);

  if (lr < target.dataStartRow || lc < 1) return;

  var headerInfo = swReadTestDataCleanupHeaders_(sh, target.headerRows || 1, lc);
  var columns = swTestDataCleanupColumns_(headerInfo);
  var usableColumns = swTestDataCleanupUsableIndexes_(columns);
  if (!usableColumns.length) {
    swAddTestDataCleanupSkippedSource_(plan, { workbookKey: target.workbookKey, sheetName: target.sheetName, reason: 'no usable key/name/email columns' }, options);
    return;
  }

  var rowCount = lr - target.dataStartRow + 1;
  var values = sh.getRange(target.dataStartRow, 1, rowCount, lc).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    var rowNumber = target.dataStartRow + i;
    var row = values[i];
    var rec = swTestDataCleanupRecordFromRow_(row, columns);
    var match = swTestDataCleanupMatchRecord_(rec, target, plan.keyset, options);
    if (!match.matched) continue;

    var candidate = swTestDataCleanupCandidate_(target, sourceInfo, rowNumber, rec, match);
    if (swAddTestDataCleanupCandidate_(plan, candidate)) {
      swAddRecordToTestDataCleanupKeyset_(plan.keyset, rec, match);
    }
  }
}

function swAddTestDataCleanupSource_(plan, sourceInfo) {
  if (!plan._sourceKeys) plan._sourceKeys = {};
  var key = [sourceInfo.workbookKey, sourceInfo.spreadsheetId, sourceInfo.sheetId].join('|');
  if (plan._sourceKeys[key]) return;
  plan._sourceKeys[key] = true;
  plan.sources.push(sourceInfo);
}

function swAddTestDataCleanupSkippedSource_(plan, item, options) {
  if (options && options._recordSkippedSources === false) return;
  if (!plan._skippedSourceKeys) plan._skippedSourceKeys = {};
  var key = [item.workbookKey || '', item.sheetName || '', item.reason || ''].join('|');
  if (plan._skippedSourceKeys[key]) return;
  plan._skippedSourceKeys[key] = true;
  plan.skippedSources.push(item);
}

function swReadTestDataCleanupHeaders_(sh, headerRows, lastColumn) {
  var rows = sh.getRange(1, 1, Math.max(1, headerRows), lastColumn).getDisplayValues();
  var headers = [];
  for (var c = 0; c < lastColumn; c++) {
    var pieces = [];
    for (var r = 0; r < rows.length; r++) {
      var piece = swTestDataCleanupTrim_(rows[r][c]);
      if (piece) pieces.push(piece);
    }
    headers.push(pieces.join(' '));
  }
  return {
    headers: headers,
    map: swTestDataCleanupHeaderMap_(headers)
  };
}

function swTestDataCleanupColumns_(headerInfo) {
  var H = headerInfo.map;
  return {
    root: swTestDataCleanupPick_(H, ['RootApptID', 'Root Appt ID', 'Root Appointment ID', 'ROOT', 'Root_ID']),
    appt: swTestDataCleanupPick_(H, ['APPT_ID', 'Appt ID', 'Appointment ID']),
    uid: swTestDataCleanupPick_(H, ['CalendlyEventUID', 'Calendly Event UID', 'Admin: Calendly Event UID', 'Acuity ID', 'UID', 'External Booking UID', 'ProviderAppointmentID', 'Provider Appointment ID']),
    taskId: swTestDataCleanupPick_(H, ['TaskID', 'Task ID']),
    name: swTestDataCleanupPick_(H, ['Customer Name', 'Customer', 'Client Name', 'Name', 'Full Name', 'Contact Name', 'Lead Name', 'Guest Name', 'Booked By']),
    nameFirst: swTestDataCleanupPick_(H, ['First Name', 'First', 'Given Name', 'Customer First Name', 'Lead First Name', 'Contact First Name']),
    nameLast: swTestDataCleanupPick_(H, ['Last Name', 'Last', 'Surname', 'Family Name', 'Customer Last Name', 'Lead Last Name', 'Contact Last Name']),
    emailLower: swTestDataCleanupPick_(H, ['EmailLower', 'Email Lower']),
    email: swTestDataCleanupPick_(H, ['Email', 'Email Address', 'E-mail', 'Customer Email', 'Client Email', 'Contact Email', 'Customer Email Address', 'Contact Email Address', 'User Email']),
    phone: swTestDataCleanupPick_(H, ['PhoneNorm', 'Phone Norm', 'Phone', 'Phone Number', 'Mobile', 'Tel']),
    so: swTestDataCleanupPick_(H, ['SO#', 'SO #', 'SO', 'SO Number', 'Sales Order', 'Sales Order #']),
    brand: swTestDataCleanupPick_(H, ['Brand', 'Company']),
    visitDate: swTestDataCleanupPick_(H, ['Visit Date', 'Appointment Date', 'Date', 'PaymentDateTime', 'Payment DateTime']),
    status: swTestDataCleanupPick_(H, ['Status', 'DocStatus', 'Doc Status']),
    automationNotes: swTestDataCleanupPick_(H, ['Automation Notes', 'Automation Note']),
    paymentId: swTestDataCleanupPick_(H, ['PAYMENT_ID', 'Payment ID', 'PaymentId']),
    docNumber: swTestDataCleanupPick_(H, ['DocNumber', 'Doc #', 'Document Number']),
    payloadJson: swTestDataCleanupPick_(H, [
      'Payload JSON',
      'Payload',
      'Request JSON',
      'Raw JSON',
      'RawPayloadJSON',
      'JSON',
      'Event JSON',
      'Saved Lines JSON',
      'Saved Line JSON',
      'Line Items JSON',
      'Items JSON',
      'Lines JSON',
      'SavedLinesJSON'
    ])
  };
}

function swTestDataCleanupUsableIndexes_(columns) {
  var out = [];
  Object.keys(columns || {}).forEach(function (key) {
    var idx = Number(columns[key]);
    if (isFinite(idx) && idx >= 0 && out.indexOf(idx) < 0) out.push(idx);
  });
  return out;
}

function swTestDataCleanupRecordFromRow_(row, columns) {
  var combinedName = swTestDataCleanupBuildNameFromParts_(
    swTestDataCleanupCell_(row, columns.nameFirst),
    swTestDataCleanupCell_(row, columns.nameLast)
  );
  var rec = {
    root: swTestDataCleanupCleanId_(swTestDataCleanupCell_(row, columns.root)),
    appt: swTestDataCleanupCleanId_(swTestDataCleanupCell_(row, columns.appt)),
    uid: swTestDataCleanupCleanId_(swTestDataCleanupCell_(row, columns.uid)),
    taskId: swTestDataCleanupCleanId_(swTestDataCleanupCell_(row, columns.taskId)),
    customerName: swTestDataCleanupTrim_(swTestDataCleanupCell_(row, columns.name) || combinedName),
    email: swTestDataCleanupNormEmail_(swTestDataCleanupCell_(row, columns.emailLower) || swTestDataCleanupCell_(row, columns.email)),
    phone: swTestDataCleanupNormPhone_(swTestDataCleanupCell_(row, columns.phone)),
    so: swTestDataCleanupCleanId_(swTestDataCleanupCell_(row, columns.so)),
    brand: swTestDataCleanupTrim_(swTestDataCleanupCell_(row, columns.brand)),
    visitDate: swTestDataCleanupTrim_(swTestDataCleanupCell_(row, columns.visitDate)),
    status: swTestDataCleanupTrim_(swTestDataCleanupCell_(row, columns.status)),
    automationNotes: swTestDataCleanupTrim_(swTestDataCleanupCell_(row, columns.automationNotes)),
    paymentId: swTestDataCleanupCleanId_(swTestDataCleanupCell_(row, columns.paymentId)),
    docNumber: swTestDataCleanupCleanId_(swTestDataCleanupCell_(row, columns.docNumber)),
    payloadJson: swTestDataCleanupTrim_(swTestDataCleanupCell_(row, columns.payloadJson) || swTestDataCleanupFindJsonInRow_(row))
  };
  swMergePayloadIntoTestDataCleanupRecord_(rec);
  swAttachTestDataCleanupRecordIdentityFallbacks_(rec, row, columns);
  return rec;
}

function swTestDataCleanupFindJsonInRow_(row) {
  if (!row) return '';
  for (var i = 0; i < row.length; i++) {
    var value = swTestDataCleanupTrim_(row[i]);
    if (!value || value.length < 2) continue;
    var first = value.charAt(0);
    if (first !== '{' && first !== '[') continue;
    try {
      JSON.parse(value);
      return value;
    } catch (_) {}
  }
  return '';
}

function swTestDataCleanupBuildNameFromParts_(firstName, lastName) {
  firstName = swTestDataCleanupTrim_(firstName);
  lastName = swTestDataCleanupTrim_(lastName);
  if (!firstName && !lastName) return '';
  return (firstName + ' ' + lastName).trim();
}

function swAttachTestDataCleanupRecordIdentityFallbacks_(rec, row, columns) {
  rec.customerName = swTestDataCleanupTrim_(rec.customerName || '');
  rec.email = swTestDataCleanupNormEmail_(rec.email || '');

  if (!rec.customerName) {
    rec.customerName = swTestDataCleanupBuildNameFromParts_(
      swTestDataCleanupCell_(row, columns.nameFirst),
      swTestDataCleanupCell_(row, columns.nameLast)
    );
  }
  if (!rec.email) {
    var emailFromRow = swTestDataCleanupFindEmailInRow_(row);
    if (emailFromRow) rec.email = emailFromRow;
  }
  return rec;
}

function swTestDataCleanupFindEmailInRow_(row) {
  if (!row) return '';
  for (var i = 0; i < row.length; i++) {
    var email = swTestDataCleanupLooksLikeEmailValue_(swTestDataCleanupTrim_(row[i]));
    if (email) return email;
  }
  return '';
}

function swMergePayloadIntoTestDataCleanupRecord_(rec) {
  if (!rec.payloadJson) return;
  var payload = swParseTestDataCleanupPayload_(rec.payloadJson);
  if (!payload) return;

  rec.root = rec.root || swTestDataCleanupCleanId_(payload.root);
  rec.appt = rec.appt || swTestDataCleanupCleanId_(payload.appt);
  rec.uid = rec.uid || swTestDataCleanupCleanId_(payload.uid);
  rec.customerName = rec.customerName || swTestDataCleanupTrim_(payload.customerName);
  rec.email = rec.email || swTestDataCleanupNormEmail_(payload.email);
  rec.phone = rec.phone || swTestDataCleanupNormPhone_(payload.phone);
  rec.so = rec.so || swTestDataCleanupCleanId_(payload.so);
}

function swParseTestDataCleanupPayload_(text) {
  var parsed = null;
  try { parsed = JSON.parse(text); } catch (_) {}
  if (!parsed) return null;

  var out = {};
  swWalkTestDataCleanupPayload_(parsed, out, 0);
  return out;
}

function swWalkTestDataCleanupPayload_(value, out, depth) {
  if (depth > 6 || value == null) return;
  if (Array.isArray(value)) {
    for (var i = 0; i < value.length; i++) swWalkTestDataCleanupPayload_(value[i], out, depth + 1);
    return;
  }
  if (typeof value !== 'object') return;

  Object.keys(value).forEach(function (key) {
    var v = value[key];
    var hk = swTestDataCleanupHeaderKey_(key);
    if (typeof v !== 'object' || v == null) {
      if (!out.root && (hk === 'rootapptid' || hk === 'rootappt' || hk === 'root')) out.root = v;
      if (!out.appt && (hk === 'apptid' || hk === 'appointmentid')) out.appt = v;
      if (!out.uid && (hk === 'calendlyeventuid' || hk === 'externalbookinguid' || hk === 'uid')) out.uid = v;
      if (!out.customerName && (hk === 'customername' || hk === 'clientname' || hk === 'name' || hk === 'fullname')) out.customerName = v;
      if (!out.customerName && (hk === 'contactname' || hk === 'leadname' || hk === 'bookedby')) out.customerName = v;
      if (!out.email && (hk === 'email' || hk === 'emaillower' || hk === 'emailaddress')) out.email = v;
      if (!out.email && (hk === 'customeremail' || hk === 'clientemail' || hk === 'contactemail' || hk === 'leademail' || hk === 'bookedbyemail')) out.email = v;
      if (!out.phone && (hk === 'phone' || hk === 'phonenumber' || hk === 'phonenorm')) out.phone = v;
      if (!out.so && (hk === 'so' || hk === 'sonumber' || hk === 'salesorder' || hk === 'salesordernumber')) out.so = v;

      if (!out.customerName && swTestDataCleanupLooksLikeNameField_(hk) && swTestDataCleanupLooksLikeNameValue_(v)) out.customerName = v;
      if (!out.email && swTestDataCleanupLooksLikeEmailField_(hk) && swTestDataCleanupLooksLikeEmailValue_(v)) out.email = v;
    }
    swWalkTestDataCleanupPayload_(v, out, depth + 1);
  });
}

function swTestDataCleanupMatchRecord_(rec, target, keyset, options) {
  var passMode = swTestDataCleanupTrim_(options && options._scanPassMode || 'mixed');
  if (passMode === 'directOnly') {
    if (target.allowDirectTestMatch !== false && !target.exactKeyOnly) {
      return swTestDataCleanupDirectMatch_(rec, target, options);
    }
    return { matched: false };
  }
  if (passMode === 'keyOnly') {
    return swTestDataCleanupExactKeyMatch_(rec, keyset);
  }

  var keyMatch = swTestDataCleanupExactKeyMatch_(rec, keyset);
  if (keyMatch.matched) return keyMatch;

  if (target.exactKeyOnly) return { matched: false };

  if (target.allowDirectTestMatch !== false) {
    var direct = swTestDataCleanupDirectMatch_(rec, target, options);
    if (direct.matched) return direct;
  }

  return { matched: false };
}

function swTestDataCleanupExactKeyMatch_(rec, keyset) {
  var taskIdentity = swTestDataCleanupLookupTestDataCleanupKeysetIdentity_(keyset.taskIds, rec.taskId);
  if (rec.taskId && taskIdentity.found) return {
    matched: true,
    matchType: 'taskId',
    matchedField: 'TaskID',
    matchedValue: rec.taskId,
    matchedCustomerName: taskIdentity.customerName,
    matchedEmail: taskIdentity.email,
    seedMatchType: taskIdentity.seedMatchType,
    seedMatchedField: taskIdentity.seedMatchedField,
    seedMatchedValue: taskIdentity.seedMatchedValue,
    seedReason: taskIdentity.seedReason,
    reason: 'exact task ID from test data keyset'
  };
  var rootIdentity = swTestDataCleanupLookupTestDataCleanupKeysetIdentity_(keyset.roots, rec.root);
  if (rec.root && rootIdentity.found) return {
    matched: true,
    matchType: 'root',
    matchedField: 'RootApptID',
    matchedValue: rec.root,
    matchedCustomerName: rootIdentity.customerName,
    matchedEmail: rootIdentity.email,
    seedMatchType: rootIdentity.seedMatchType,
    seedMatchedField: rootIdentity.seedMatchedField,
    seedMatchedValue: rootIdentity.seedMatchedValue,
    seedReason: rootIdentity.seedReason,
    reason: 'exact root ID from test data keyset'
  };
  var apptIdentity = swTestDataCleanupLookupTestDataCleanupKeysetIdentity_(keyset.appts, rec.appt);
  if (rec.appt && apptIdentity.found) return {
    matched: true,
    matchType: 'appt',
    matchedField: 'APPT_ID',
    matchedValue: rec.appt,
    matchedCustomerName: apptIdentity.customerName,
    matchedEmail: apptIdentity.email,
    seedMatchType: apptIdentity.seedMatchType,
    seedMatchedField: apptIdentity.seedMatchedField,
    seedMatchedValue: apptIdentity.seedMatchedValue,
    seedReason: apptIdentity.seedReason,
    reason: 'exact appointment ID from test data keyset'
  };
  var uidIdentity = swTestDataCleanupLookupTestDataCleanupKeysetIdentity_(keyset.uids, rec.uid);
  if (rec.uid && uidIdentity.found) return {
    matched: true,
    matchType: 'uid',
    matchedField: 'CalendlyEventUID',
    matchedValue: rec.uid,
    matchedCustomerName: uidIdentity.customerName,
    matchedEmail: uidIdentity.email,
    seedMatchType: uidIdentity.seedMatchType,
    seedMatchedField: uidIdentity.seedMatchedField,
    seedMatchedValue: uidIdentity.seedMatchedValue,
    seedReason: uidIdentity.seedReason,
    reason: 'exact booking UID from test data keyset'
  };
  var soIdentity = swTestDataCleanupLookupTestDataCleanupKeysetIdentity_(keyset.sos, rec.so);
  if (rec.so && soIdentity.found) return {
    matched: true,
    matchType: 'so',
    matchedField: 'SO#',
    matchedValue: rec.so,
    matchedCustomerName: soIdentity.customerName,
    matchedEmail: soIdentity.email,
    seedMatchType: soIdentity.seedMatchType,
    seedMatchedField: soIdentity.seedMatchedField,
    seedMatchedValue: soIdentity.seedMatchedValue,
    seedReason: soIdentity.seedReason,
    reason: 'exact SO from test data keyset'
  };
  return { matched: false };
}

function swTestDataCleanupLookupTestDataCleanupKeysetIdentity_(set, key) {
  if (!set || !key) return { found: false };
  var entry = set[key];
  if (!entry) return { found: false };
  if (entry === true) return { found: true };
  return {
    found: true,
    customerName: swTestDataCleanupTrim_(entry.customerName || ''),
    email: swTestDataCleanupNormEmail_(entry.email || ''),
    seedMatchType: swTestDataCleanupTrim_(entry.seedMatchType || ''),
    seedMatchedField: swTestDataCleanupTrim_(entry.seedMatchedField || ''),
    seedMatchedValue: swTestDataCleanupTrim_(entry.seedMatchedValue || ''),
    seedReason: swTestDataCleanupTrim_(entry.seedReason || '')
  };
}

function swTestDataCleanupDirectMatch_(rec, target, options) {
  if (swTestDataCleanupExactMasterCustomerNameMode_(options)) {
    return swTestDataCleanupExactMasterCustomerNameDirectMatch_(rec, target, options);
  }
  if (swTestDataCleanupRecordMarkedRemoved_(rec)) return { matched: false };
  var name = swTestDataCleanupTrim_(rec.customerName);
  var email = swTestDataCleanupNormEmail_(rec.email);
  var mode = swTestDataCleanupNorm_(options && options.matchMode || 'strict');
  var payloadMatch = swTestDataCleanupPayloadDirectMatch_(rec.payloadJson, target, mode);

  if (name && swTestDataCleanupTextLooksTest_(name, mode)) {
    return { matched: true, matchType: 'directName', matchedField: 'Customer Name', matchedValue: name, reason: 'customer name contains a test-data token' };
  }
  if (email && swTestDataCleanupEmailLooksTest_(email, mode)) {
    return { matched: true, matchType: 'directEmail', matchedField: 'Email', matchedValue: email, reason: 'email contains a test-data token' };
  }
  if (payloadMatch) {
    return { matched: true, matchType: 'directPayload', matchedField: 'Payload JSON', matchedValue: payloadMatch, reason: 'payload JSON contains a test-data token' };
  }
  return { matched: false };
}

function swTestDataCleanupExactMasterCustomerNameMode_(options) {
  var mode = swTestDataCleanupHeaderKey_((options && (options.seedMode || options.matchMode || options.cleanupMode)) || '');
  return mode === 'exactmastercustomernames' || mode === 'exactmastercustomername';
}

function swTestDataCleanupExactMasterCustomerNameDirectMatch_(rec, target, options) {
  if (swTestDataCleanupRecordMarkedRemoved_(rec)) return { matched: false };
  if (!swTestDataCleanupIsMasterAppointmentsSheet_(target && target.sheetName)) return { matched: false };
  var name = swTestDataCleanupTrim_(rec && rec.customerName);
  if (!name) return { matched: false };
  var requested = swExactMasterCustomerCleanupNameSet_((options && options.exactMasterCustomerNames) || []);
  var requestedName = requested[swExactMasterCustomerCleanupNameKey_(name)];
  if (!requestedName) return { matched: false };
  return {
    matched: true,
    matchType: 'directMasterCustomerName',
    matchedField: 'Customer Name',
    matchedValue: name,
    reason: 'Customer Name exactly matched requested cleanup list in 00_Master Appointments'
  };
}

function swTestDataCleanupIsMasterAppointmentsSheet_(sheetName) {
  var actual = swTestDataCleanupHeaderKey_(sheetName);
  var configured = '';
  try {
    configured = typeof SW_SHEETS !== 'undefined' && SW_SHEETS.MASTER ? SW_SHEETS.MASTER : '';
  } catch (_) {}
  return actual === swTestDataCleanupHeaderKey_(configured || '00_Master Appointments') ||
    actual === swTestDataCleanupHeaderKey_('Tab00Master appointments') ||
    actual === swTestDataCleanupHeaderKey_('Tab00Master');
}

function swTestDataCleanupRecordMarkedRemoved_(rec) {
  var notes = swTestDataCleanupNorm_(rec && rec.automationNotes);
  return notes.indexOf('removed due to test data cleanup') >= 0;
}

function swTestDataCleanupPayloadDirectMatch_(payloadJson, target, mode) {
  var raw = swTestDataCleanupTrim_(payloadJson);
  if (!raw) return '';
  if (!swTestDataCleanupSheetAllowsPayloadDirectMatch_(target && target.sheetName)) return '';
  var parsed = null;
  try {
    parsed = JSON.parse(raw);
  } catch (_) {
    var fallback = swTestDataCleanupPayloadRegexFallbackMatch_(raw);
    return fallback ? swTestDataCleanupPayloadMatchSnippet_('desc', fallback) : '';
  }
  return swTestDataCleanupPayloadValueMatch_(parsed, mode, '', 0);
}

function swTestDataCleanupSheetAllowsPayloadDirectMatch_(sheetName) {
  sheetName = swTestDataCleanupTrim_(sheetName);
  return sheetName === '02_Form_Inbox' ||
    sheetName === '00_Master Appointments' ||
    sheetName === '_ExternalBookingEvents' ||
    sheetName === '_IntakeQueue' ||
    sheetName === '_SalesDataCleanup';
}

function swTestDataCleanupPayloadRegexFallbackMatch_(raw) {
  var m = String(raw || '').match(/["']desc(?:ription)?["']\s*:\s*["']\s*(test|testing|tester|sample|dummy|fake)\s*[0-9_-]*\s*["']/i);
  return m && m[1] ? m[1] : '';
}

function swTestDataCleanupPayloadValueMatch_(value, mode, keyHint, depth) {
  if (depth > 8 || value == null) return '';

  if (Array.isArray(value)) {
    for (var i = 0; i < value.length; i++) {
      var arrHit = swTestDataCleanupPayloadValueMatch_(value[i], mode, keyHint, depth + 1);
      if (arrHit) return arrHit;
    }
    return '';
  }

  var t = typeof value;
  if (t === 'object') {
    var keys = Object.keys(value);
    for (var k = 0; k < keys.length; k++) {
      var key = keys[k];
      var hk = swTestDataCleanupHeaderKey_(key);
      var child = value[key];
      var hit = swTestDataCleanupPayloadPrimitiveMatch_(hk, child, mode);
      if (hit) return hit;
      var nested = swTestDataCleanupPayloadValueMatch_(child, mode, hk, depth + 1);
      if (nested) return nested;
    }
    return '';
  }

  return swTestDataCleanupPayloadPrimitiveMatch_(keyHint, value, mode);
}

function swTestDataCleanupPayloadPrimitiveMatch_(key, value, mode) {
  if (value == null) return '';

  if (typeof value === 'boolean') {
    if (value === true && swTestDataCleanupPayloadTestFlagKey_(key)) {
      return swTestDataCleanupPayloadMatchSnippet_(key || 'testFlag', 'true');
    }
    return '';
  }

  var text = swTestDataCleanupTrim_(value);
  if (!text) return '';
  if (!swTestDataCleanupPayloadFieldAllowsTestToken_(key)) return '';
  if (!swTestDataCleanupPayloadFieldIsHighConfidence_(key)) return '';
  if (!swTestDataCleanupTextLooksTest_(text, mode)) return '';
  return swTestDataCleanupPayloadMatchSnippet_(key || 'value', text);
}

function swTestDataCleanupPayloadFieldAllowsTestToken_(key) {
  var hk = swTestDataCleanupHeaderKey_(key);
  if (!hk) return true;
  if (/(email|mail|phone|mobile|tel|rootapptid|apptid|appointmentid|calendlyeventuid|uid|taskid|so|sonumber|salesorder|paymentid|docnumber)/.test(hk)) {
    return false;
  }
  return true;
}

function swTestDataCleanupPayloadFieldIsHighConfidence_(key) {
  var hk = swTestDataCleanupHeaderKey_(key);
  if (!hk) return false;
  return /(desc|description|item|itemname|product|service|lineitem|label|title|name)/.test(hk);
}

function swTestDataCleanupPayloadTestFlagKey_(key) {
  var hk = swTestDataCleanupHeaderKey_(key);
  if (!hk) return false;
  return hk === 'istest' || hk === 'test' || hk === 'testmode' || hk === 'istesting';
}

function swTestDataCleanupPayloadMatchSnippet_(key, text) {
  var clean = swTestDataCleanupTrim_(text);
  if (clean.length > 100) clean = clean.slice(0, 97) + '...';
  return String(key || 'value') + ': ' + clean;
}

function swTestDataCleanupTextLooksTest_(value, mode) {
  var text = swTestDataCleanupNorm_(value);
  if (!text) return false;
  if (swTestDataCleanupExplicitCustomerNameLooksTest_(text)) return true;
  if (mode === 'contains') return /test|testing|tester|sample|dummy|fake/.test(text);
  var spaced = text.replace(/[^a-z0-9]+/g, ' ');
  return /(^| )(test|testing|tester|testclient|testcustomer|sample|dummy|fake)([0-9]*)( |$)/.test(spaced);
}

function swTestDataCleanupExplicitCustomerNameLooksTest_(value) {
  var compact = swTestDataCleanupNorm_(value).replace(/[^a-z0-9]+/g, '');
  return compact === 'testdemo3' || compact === 'testdemo34' || compact === 'testdemo5';
}

function swFilterTestDemoCleanupPreviewRows_(rows, names) {
  var explicit = {};
  (names || []).forEach(function (name) {
    var compact = swTestDataCleanupNorm_(name).replace(/[^a-z0-9]+/g, '');
    if (compact) explicit[compact] = true;
  });
  var seedIds = {};
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || {};
    var compactCustomer = swTestDataCleanupNorm_(row.customerName).replace(/[^a-z0-9]+/g, '');
    var compactMatched = swTestDataCleanupNorm_(row.matchedValue).replace(/[^a-z0-9]+/g, '');
    var compactSeed = swTestDataCleanupNorm_(row.seedMatchedValue).replace(/[^a-z0-9]+/g, '');
    var directDemo = !!(explicit[compactCustomer] || explicit[compactMatched] || explicit[compactSeed]);
    if (directDemo) {
      swSetIfValue_(seedIds, row.root);
      swSetIfValue_(seedIds, row.appt);
      swSetIfValue_(seedIds, row.taskId);
      swSetIfValue_(seedIds, row.so);
    }
  }
  for (var j = 0; j < rows.length; j++) {
    var candidate = rows[j] || {};
    var compactName = swTestDataCleanupNorm_(candidate.customerName).replace(/[^a-z0-9]+/g, '');
    var compactValue = swTestDataCleanupNorm_(candidate.matchedValue).replace(/[^a-z0-9]+/g, '');
    var compactSeedValue = swTestDataCleanupNorm_(candidate.seedMatchedValue).replace(/[^a-z0-9]+/g, '');
    if (explicit[compactName] || explicit[compactValue] || explicit[compactSeedValue] ||
        (candidate.root && seedIds[candidate.root]) ||
        (candidate.appt && seedIds[candidate.appt]) ||
        (candidate.taskId && seedIds[candidate.taskId]) ||
        (candidate.so && seedIds[candidate.so])) {
      out.push(candidate);
    }
  }
  return out;
}

function swCollectTestDemoCleanupLinkedNeedles_(rows) {
  var seen = {};
  var out = [];
  function add(value) {
    value = swTestDataCleanupTrim_(value);
    if (!value || seen[value]) return;
    seen[value] = true;
    out.push(value);
  }
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || {};
    add(row.root);
    add(row.appt);
    add(row.taskId);
    add(row.so);
  }
  return out;
}

function swApplyTestDataCleanupRows_(plan, options) {
  var candidates = (plan && plan.candidates) || [];
  if (swTestDataCleanupExactMasterCustomerNameMode_(options)) {
    return swDeleteTestDataCleanupRows_(plan);
  }
  var masterRows = [];
  var deleteRows = [];
  for (var i = 0; i < candidates.length; i++) {
    var row = candidates[i];
    if (row && row.workbookKey === 'workflow' && row.sheetName === '00_Master Appointments') {
      masterRows.push(row);
    } else {
      deleteRows.push(row);
    }
  }

  var markedRows = swMarkMasterAppointmentsTestDataCleanup_(masterRows);
  var deletedRows = [];
  if (deleteRows.length) {
    var deletePlan = {};
    Object.keys(plan || {}).forEach(function (key) { deletePlan[key] = plan[key]; });
    deletePlan.candidates = deleteRows;
    deletedRows = swDeleteTestDataCleanupRows_(deletePlan);
  }
  return markedRows.concat(deletedRows);
}

function swMarkMasterAppointmentsTestDataCleanup_(rows) {
  rows = rows || [];
  if (!rows.length) return [];

  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  var sh = ss.getSheetByName('00_Master Appointments');
  if (!sh) throw new Error('Missing sheet: 00_Master Appointments');

  var noteCol = swFindTestDataCleanupAutomationNotesColumn_(sh, 1);
  var stamp = swTestDataCleanupIso_(new Date());
  var marker = 'Removed due to test data cleanup on ' + stamp + '.';
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var candidate = rows[i];
    var rowNumber = Number(candidate && candidate.rowNumber);
    if (!rowNumber || rowNumber < 2) continue;
    var cell = sh.getRange(rowNumber, noteCol);
    var current = swTestDataCleanupTrim_(cell.getDisplayValue());
    if (current.indexOf('Removed due to test data cleanup') < 0) {
      cell.setValue(current ? (current + '\n' + marker) : marker);
    }
    var marked = {};
    Object.keys(candidate || {}).forEach(function (key) { marked[key] = candidate[key]; });
    marked.action = 'markedAutomationNotes';
    marked.reason = (marked.reason ? marked.reason + '; ' : '') + 'retained Master Appointment row and marked Automation Notes';
    out.push(marked);
  }
  try { SpreadsheetApp.flush(); } catch (_) {}
  return out;
}

function swFindTestDataCleanupAutomationNotesColumn_(sh, headerRow) {
  var headers = sh.getRange(headerRow || 1, 1, 1, sh.getLastColumn()).getDisplayValues()[0] || [];
  for (var i = 0; i < headers.length; i++) {
    var key = swTestDataCleanupHeaderKey_(headers[i]);
    if (key === 'automationnotes' || key === 'automationnote') return i + 1;
  }
  throw new Error('Could not find Automation Notes column on 00_Master Appointments.');
}

function swReadTestDemoMasterAppointmentRowsViaSheetsApi_(spreadsheetId, names) {
  var values = swSheetsApiReadValues_(spreadsheetId, '00_Master Appointments', 'A1:ZZ');
  var rows = (values && values.values) || [];
  var headers = rows.length ? rows[0] : [];
  var H = {};
  for (var h = 0; h < headers.length; h++) H[swTestDataCleanupHeaderKey_(headers[h])] = h;
  var wanted = {};
  (names || []).forEach(function (name) {
    var compact = swTestDataCleanupNorm_(name).replace(/[^a-z0-9]+/g, '');
    if (compact) wanted[compact] = true;
  });
  var out = [];
  for (var r = 1; r < rows.length; r++) {
    var row = rows[r] || [];
    var customerName = swTestDataCleanupTrim_(row[H.customername]);
    var compactName = swTestDataCleanupNorm_(customerName).replace(/[^a-z0-9]+/g, '');
    if (!wanted[compactName]) continue;
    out.push({
      row: r + 1,
      customerName: customerName,
      email: swTestDataCleanupTrim_(row[H.email]),
      root: swTestDataCleanupTrim_(row[H.rootapptid]),
      appt: swTestDataCleanupTrim_(row[H.apptid]),
      automationNotes: swTestDataCleanupTrim_(row[H.automationnotes])
    });
  }
  return out;
}

function swTestDataCleanupSearchCounts_(searchResult) {
  var out = {};
  Object.keys(searchResult || {}).sort().forEach(function (key) {
    out[key] = (searchResult[key] && searchResult[key].count) || 0;
  });
  return out;
}

function swTestDataCleanupEmailLooksTest_(email, mode) {
  email = swTestDataCleanupNormEmail_(email);
  if (!email) return false;
  if (mode === 'contains') return /test|testing|tester|sample|dummy|fake/.test(email);

  var parts = email.split('@');
  var local = parts[0] || '';
  var domain = parts.slice(1).join('@') || '';
  var token = /(^|[._+\-])(test|testing|tester|testclient|testcustomer|sample|dummy|fake)([0-9]*)([._+\-]|$)/;
  return token.test(local) || token.test(domain);
}

function swTestDataCleanupCandidate_(target, sourceInfo, rowNumber, rec, match) {
  return {
    workbookKey: target.workbookKey,
    workbookLabel: target.workbookLabel,
    spreadsheetId: sourceInfo.spreadsheetId,
    spreadsheetName: sourceInfo.spreadsheetName,
    sheetName: target.sheetName,
    sheetId: target.sheet.getSheetId(),
    rowNumber: rowNumber,
    deleteOrder: target.deleteOrder || 100,
    root: rec.root || '',
    appt: rec.appt || '',
    uid: rec.uid || '',
    taskId: rec.taskId || '',
    so: rec.so || '',
    customerName: rec.customerName || '',
    email: rec.email || '',
    matchedCustomerName: swTestDataCleanupTrim_(match.matchedCustomerName || rec.customerName || ''),
    matchedEmail: swTestDataCleanupNormEmail_(match.matchedEmail || rec.email || ''),
    seedMatchType: swTestDataCleanupTrim_(match.seedMatchType || match.matchType || ''),
    seedMatchedField: swTestDataCleanupTrim_(match.seedMatchedField || match.matchedField || ''),
    seedMatchedValue: swTestDataCleanupTrim_(match.seedMatchedValue || match.matchedValue || ''),
    seedReason: swTestDataCleanupTrim_(match.seedReason || match.reason || ''),
    phone: rec.phone || '',
    brand: rec.brand || '',
    visitDate: rec.visitDate || '',
    status: rec.status || '',
    paymentId: rec.paymentId || '',
    docNumber: rec.docNumber || '',
    matchType: match.matchType || '',
    matchedField: match.matchedField || '',
    matchedValue: match.matchedValue || '',
    reason: match.reason || '',
    headerRows: target.headerRows || 1,
    rowFingerprint: swTestDataCleanupRowFingerprint_(rec)
  };
}

function swAddTestDataCleanupCandidate_(plan, candidate) {
  var key = [
    candidate.workbookKey,
    candidate.spreadsheetId,
    candidate.sheetId,
    candidate.rowNumber
  ].join('|');
  if (!plan._candidateKeys) plan._candidateKeys = {};
  if (plan._candidateKeys[key]) return false;
  plan._candidateKeys[key] = true;
  plan.candidates.push(candidate);
  return true;
}

function swAddRecordToTestDataCleanupKeyset_(keyset, rec, match) {
  match = match || {};
  if (!swTestDataCleanupMatchSeedsKeyset_(match)) return;
  var identity = swTestDataCleanupRecordIdentityForKeyset_(rec, match);
  if (!swTestDataCleanupExactMasterCustomerNameSeed_(match)) {
    swStoreTestDataCleanupKeysetIdentity_(keyset.roots, rec.root, identity);
  }
  swStoreTestDataCleanupKeysetIdentity_(keyset.appts, rec.appt, identity);
  swStoreTestDataCleanupKeysetIdentity_(keyset.uids, rec.uid, identity);
  swStoreTestDataCleanupKeysetIdentity_(keyset.taskIds, rec.taskId, identity);
  if (match.matchType !== 'directName' && match.matchType !== 'directEmail' && match.matchType !== 'directPayload') {
    swStoreTestDataCleanupKeysetIdentity_(keyset.sos, rec.so, identity);
  }
  swSetIfValue_(keyset.emails, rec.email);
  swSetIfValue_(keyset.phones, rec.phone);
}

function swTestDataCleanupMatchSeedsKeyset_(match) {
  var t = swTestDataCleanupTrim_(match && match.matchType);
  return t === 'directName' || t === 'directEmail' || t === 'directPayload' || t === 'directMasterCustomerName';
}

function swTestDataCleanupExactMasterCustomerNameSeed_(match) {
  return swTestDataCleanupTrim_(match && match.matchType) === 'directMasterCustomerName';
}

function swStoreTestDataCleanupKeysetIdentity_(set, key, identity) {
  if (!set || !key || !identity) return;
  if (!set[key]) {
    set[key] = identity;
    return;
  }
  if (set[key] === true) {
    set[key] = identity;
    return;
  }
  if (!set[key].customerName && identity.customerName) set[key].customerName = identity.customerName;
  if (!set[key].email && identity.email) set[key].email = identity.email;
}

function swTestDataCleanupRecordIdentityForKeyset_(rec, match) {
  match = match || {};
  return {
    customerName: swTestDataCleanupTrim_(rec && rec.customerName),
    email: swTestDataCleanupNormEmail_(rec && rec.email),
    seedMatchType: swTestDataCleanupTrim_(match.matchType),
    seedMatchedField: swTestDataCleanupTrim_(match.matchedField),
    seedMatchedValue: swTestDataCleanupTrim_(match.matchedValue),
    seedReason: swTestDataCleanupTrim_(match.reason)
  };
}

function swNewTestDataCleanupKeyset_() {
  return {
    roots: {},
    appts: {},
    uids: {},
    taskIds: {},
    sos: {},
    emails: {},
    phones: {}
  };
}

function swSetIfValue_(set, value) {
  value = swTestDataCleanupTrim_(value);
  if (value) set[value] = true;
}

function swDeleteTestDataCleanupRows_(plan) {
  var deleted = [];
  var groups = {};
  for (var i = 0; i < plan.candidates.length; i++) {
    var c = plan.candidates[i];
    var key = [c.workbookKey, c.spreadsheetId, c.sheetId].join('|');
    if (!groups[key]) groups[key] = { sample: c, rows: [] };
    groups[key].rows.push(c);
  }

  var groupList = Object.keys(groups).map(function (key) { return groups[key]; });
  groupList.sort(function (a, b) {
    var ao = a.sample.deleteOrder || 100;
    var bo = b.sample.deleteOrder || 100;
    if (ao !== bo) return ao - bo;
    return String(a.sample.sheetName).localeCompare(String(b.sample.sheetName));
  });

  for (var g = 0; g < groupList.length; g++) {
    var group = groupList[g];
    var sh = swOpenTestDataCleanupSheet_(group.sample);
    if (!sh) {
      plan.errors.push({ type: 'sheet', sheetName: group.sample.sheetName, message: 'Sheet unavailable at apply time.' });
      continue;
    }

    group.rows.sort(function (a, b) { return b.rowNumber - a.rowNumber; });
    for (var r = 0; r < group.rows.length; r++) {
      var row = group.rows[r];
      try {
        var recheck = swCurrentTestDataCleanupRowMatchesCandidate_(sh, row);
        if (!recheck.ok) {
          plan.errors.push({
            type: 'rowRevalidation',
            workbookKey: row.workbookKey,
            sheetName: row.sheetName,
            rowNumber: row.rowNumber,
            message: recheck.message || 'Current row no longer matches previewed candidate.'
          });
          continue;
        }
        sh.deleteRow(row.rowNumber);
        deleted.push(row);
      } catch (err) {
        plan.errors.push({
          type: 'row',
          workbookKey: row.workbookKey,
          sheetName: row.sheetName,
          rowNumber: row.rowNumber,
          message: err && err.message ? err.message : String(err)
        });
      }
    }
  }
  return deleted;
}

function swCurrentTestDataCleanupRowMatchesCandidate_(sh, candidate) {
  try {
    if (!sh || candidate.rowNumber < 1 || candidate.rowNumber > sh.getLastRow()) {
      return { ok: false, message: 'Candidate row is outside the current sheet range.' };
    }
    var lc = sh.getLastColumn();
    var headerInfo = swReadTestDataCleanupHeaders_(sh, candidate.headerRows || 1, lc);
    var columns = swTestDataCleanupColumns_(headerInfo);
    var row = sh.getRange(candidate.rowNumber, 1, 1, lc).getDisplayValues()[0];
    var rec = swTestDataCleanupRecordFromRow_(row, columns);
    var currentFingerprint = swTestDataCleanupRowFingerprint_(rec);
    if (currentFingerprint !== candidate.rowFingerprint) {
      return { ok: false, message: 'Candidate row fingerprint changed since preview.' };
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err && err.message ? err.message : String(err) };
  }
}

function swOpenTestDataCleanupSheet_(candidate) {
  var ss = null;
  try {
    if (candidate.workbookKey === 'workflow') {
      ss = swTestDataCleanupWorkflowSpreadsheet_();
    } else {
      ss = SpreadsheetApp.openById(candidate.spreadsheetId);
    }
    return ss.getSheetByName(candidate.sheetName);
  } catch (_) {}
  return null;
}

function swValidateTestDataCleanupPreview_(plan, options) {
  if (plan.totalCandidateRows === 0) return { ok: true, empty: true };

  var preview = swReadTestDataCleanupPreview_();
  if (!preview) {
    return { ok: false, message: 'Run sw_previewTestDataCleanupOnce() before apply.' };
  }
  if (preview.workflowSpreadsheetId !== plan.workflowSpreadsheetId) {
    return { ok: false, message: 'Preview was created for a different workflow spreadsheet.' };
  }
  if (preview.fingerprint !== plan.fingerprint) {
    return { ok: false, message: 'Current candidate rows differ from the stored preview. Run preview again before apply.' };
  }
  var previewAt = swTestDataCleanupDateMs_(preview.createdAt);
  if (previewAt && new Date().getTime() - previewAt > SW_TEST_DATA_CLEANUP_PREVIEW_MAX_AGE_MS_) {
    return { ok: false, message: 'Stored preview is older than four hours. Run preview again before apply.' };
  }
  if (options.confirmationToken && options.confirmationToken !== preview.confirmationToken) {
    return { ok: false, message: 'Confirmation token does not match the stored preview token.' };
  }
  return { ok: true, previewCreatedAt: preview.createdAt };
}

function swConfirmTestDataCleanupApply_(plan, options) {
  if (plan.totalCandidateRows === 0) return { ok: true, skipped: true, message: 'No candidate rows.' };
  if (options.confirmationToken && options.confirmationToken === plan.confirmationToken) {
    return { ok: true, source: 'options.confirmationToken' };
  }

  try {
    var ui = SpreadsheetApp.getUi();
    var preview = plan.candidates.slice(0, 12).map(function (row) {
      return row.workbookKey + ' / ' + row.sheetName + ' row ' + row.rowNumber + ': ' +
        (row.customerName || row.email || row.root || row.appt || row.so || row.taskId || '(matched row)');
    }).join('\n');
    var more = plan.totalCandidateRows > 12 ? '\n...and ' + (plan.totalCandidateRows - 12) + ' more row(s)' : '';
    var prompt = ui.prompt(
      'Confirm test data cleanup',
      'This will permanently delete ' + plan.totalCandidateRows + ' source row(s), bottom-up.\n\n' +
        preview + more + '\n\n' +
        'Type this token to apply:\n' + plan.confirmationToken,
      ui.ButtonSet.OK_CANCEL
    );
    var button = prompt.getSelectedButton();
    var response = swTestDataCleanupTrim_(prompt.getResponseText());
    if (button === ui.Button.OK && response === plan.confirmationToken) {
      return { ok: true, source: 'ui.prompt' };
    }
    return { ok: false, source: 'ui.prompt', message: 'Confirmation token was not entered.' };
  } catch (_) {}

  return { ok: false, source: 'unavailable', message: 'No UI confirmation was available and no confirmationToken option was provided.' };
}

function swNormalizeTestDataCleanupApplyOptions_(options) {
  if (typeof options === 'string') return { apply: true, confirmationToken: options };
  options = options || {};
  options.apply = true;
  return options;
}

function swStoreTestDataCleanupPreview_(plan) {
  var payload = {
    workflowSpreadsheetId: plan.workflowSpreadsheetId,
    workflowSpreadsheetName: plan.workflowSpreadsheetName,
    fingerprint: plan.fingerprint,
    confirmationToken: plan.confirmationToken,
    totalCandidateRows: plan.totalCandidateRows,
    summary: plan.summary,
    createdAt: plan.createdAt
  };
  try {
    PropertiesService.getDocumentProperties().setProperty(
      SW_TEST_DATA_CLEANUP_PREVIEW_PROPERTY_,
      JSON.stringify(payload)
    );
  } catch (err) {
    plan.errors.push({ type: 'previewStore', message: err && err.message ? err.message : String(err) });
  }
}

function swReadTestDataCleanupPreview_() {
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty(SW_TEST_DATA_CLEANUP_PREVIEW_PROPERTY_);
    return raw ? JSON.parse(raw) : null;
  } catch (_) {}
  return null;
}

function swInvalidateAfterTestDataCleanup_(plan, options) {
  var ss = swTestDataCleanupWorkflowSpreadsheet_();
  var out = {
    appointment: null,
    payment: null,
    diamond: null,
    taskList: null,
    taskDashboard: null,
    customerSearch: null,
    artifactRoots: 0,
    broad: null
  };
  var reason = SW_TEST_DATA_CLEANUP_REASON_;
  var touched = swTestDataCleanupTouchedWorkbooks_(plan.deletedRows || []);

  try {
    if (typeof swInvalidateAppointmentReadModelsAfterWrite_ === 'function') {
      swInvalidateAppointmentReadModelsAfterWrite_(ss, reason);
      out.appointment = { ok: true, helper: 'swInvalidateAppointmentReadModelsAfterWrite_' };
    }
  } catch (err1) {
    out.appointment = { ok: false, error: err1 && err1.message ? err1.message : String(err1) };
  }

  if (touched.paymentsLedger) {
    try {
      if (typeof swInvalidatePaymentReadModelsAfterWrite_ === 'function') {
        swInvalidatePaymentReadModelsAfterWrite_(ss, reason);
        out.payment = { ok: true, helper: 'swInvalidatePaymentReadModelsAfterWrite_' };
      }
    } catch (err2) {
      out.payment = { ok: false, error: err2 && err2.message ? err2.message : String(err2) };
    }
  }

  if (touched.diamonds200) {
    try {
      if (typeof swInvalidateDiamondReadModelsAfterWrite_ === 'function') {
        swInvalidateDiamondReadModelsAfterWrite_(ss, reason);
        out.diamond = { ok: true, helper: 'swInvalidateDiamondReadModelsAfterWrite_' };
      }
    } catch (err3) {
      out.diamond = { ok: false, error: err3 && err3.message ? err3.message : String(err3) };
    }
  }

  try {
    if (typeof swInvalidateTaskListCache_ === 'function') {
      swInvalidateTaskListCache_(ss);
      out.taskList = { ok: true, helper: 'swInvalidateTaskListCache_' };
    }
  } catch (err4) {
    out.taskList = { ok: false, error: err4 && err4.message ? err4.message : String(err4) };
  }

  try {
    if (typeof swInvalidateTaskDashboardProjectionCache_ === 'function') {
      swInvalidateTaskDashboardProjectionCache_(ss);
      out.taskDashboard = { ok: true, helper: 'swInvalidateTaskDashboardProjectionCache_' };
    }
  } catch (err5) {
    out.taskDashboard = { ok: false, error: err5 && err5.message ? err5.message : String(err5) };
  }

  try {
    if (typeof swInvalidateCustomerSearchReadModelCache_ === 'function') {
      swInvalidateCustomerSearchReadModelCache_(ss);
      out.customerSearch = { ok: true, helper: 'swInvalidateCustomerSearchReadModelCache_' };
    }
  } catch (err6) {
    out.customerSearch = { ok: false, error: err6 && err6.message ? err6.message : String(err6) };
  }

  if (typeof swInvalidateAppointmentArtifactRowsForRoot_ === 'function') {
    var roots = swUniqueStrings_(Object.keys(plan.keyset.roots || {}));
    roots.forEach(function (root) {
      try {
        swInvalidateAppointmentArtifactRowsForRoot_(ss, root);
        out.artifactRoots++;
      } catch (_) {}
    });
  }

  try {
    if (typeof swMarkWorkflowReadModelsStale_ === 'function') {
      swMarkWorkflowReadModelsStale_(ss, reason);
      out.broad = { ok: true, helper: 'swMarkWorkflowReadModelsStale_' };
    }
  } catch (err7) {
    out.broad = { ok: false, error: err7 && err7.message ? err7.message : String(err7) };
  }

  return out;
}

function swRebuildAfterTestDataCleanup_() {
  try {
    if (typeof sw_rebuildWorkflowReadModels === 'function') {
      return sw_rebuildWorkflowReadModels({ reason: SW_TEST_DATA_CLEANUP_REASON_ });
    }
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
  return { ok: true, skipped: true, reason: 'sw_rebuildWorkflowReadModels unavailable' };
}

function swTestDataCleanupTouchedWorkbooks_(rows) {
  var out = {};
  (rows || []).forEach(function (row) {
    out[row.workbookKey] = true;
  });
  return out;
}

function swTestDataCleanupDeleteOrder_(sheetName, workbookKey) {
  if (workbookKey === 'paymentsLedger') return 10;
  if (workbookKey === 'diamonds200') return 20;
  if (sheetName === '_AppointmentArtifacts') return 30;
  if (sheetName === '_SalesTaskQueue') return 40;
  if (sheetName === '_SalesTaskLog') return 45;
  if (sheetName === '_IntakeQueue') return 50;
  if (sheetName === '02_Form_Inbox') return 60;
  if (sheetName === '_SalesDataCleanup') return 65;
  if (sheetName === '03_Client_Status_Log') return 70;
  if (sheetName === '05_Wax_Requests') return 72;
  if (sheetName === '07_Root_Index') return 74;
  if (sheetName === '00_Master Appointments') return 90;
  return 80;
}

function swTestDataCleanupSummary_(plan) {
  var bySource = {};
  (plan.candidates || []).forEach(function (row) {
    var key = row.workbookKey + ':' + row.sheetName;
    if (!bySource[key]) bySource[key] = { workbookKey: row.workbookKey, sheetName: row.sheetName, candidates: 0, deleted: 0 };
    bySource[key].candidates++;
  });
  (plan.deletedRows || []).forEach(function (row) {
    var key = row.workbookKey + ':' + row.sheetName;
    if (!bySource[key]) bySource[key] = { workbookKey: row.workbookKey, sheetName: row.sheetName, candidates: 0, deleted: 0 };
    bySource[key].deleted++;
  });
  return {
    totalCandidateRows: plan.totalCandidateRows || 0,
    deletedCount: plan.deletedCount || 0,
    sources: Object.keys(bySource).sort().map(function (key) { return bySource[key]; }),
    keyset: swPublicTestDataCleanupKeysetCounts_(plan.keyset),
    skippedSources: plan.skippedSources || [],
    warnings: plan.warnings || [],
    errors: plan.errors || []
  };
}

function swPublicTestDataCleanupResult_(plan) {
  plan.summary = swTestDataCleanupSummary_(plan);
  return {
    ok: plan.ok !== false,
    apply: plan.apply === true,
    createdAt: plan.createdAt,
    workflowSpreadsheetId: plan.workflowSpreadsheetId,
    workflowSpreadsheetName: plan.workflowSpreadsheetName,
    totalCandidateRows: plan.totalCandidateRows,
    deletedCount: plan.deletedCount || 0,
    confirmationToken: plan.confirmationToken,
    fingerprint: plan.fingerprint,
    summary: plan.summary,
    candidates: (plan.candidates || []).map(swPublicTestDataCleanupCandidate_),
    deletedRows: (plan.deletedRows || []).map(swPublicTestDataCleanupCandidate_),
    previewValidation: plan.previewValidation || null,
    confirmation: plan.confirmation || null,
    invalidation: plan.invalidation || null,
    readModelRebuild: plan.readModelRebuild || null
  };
}

function swPublicTestDataCleanupCandidate_(row) {
  return {
    workbookKey: row.workbookKey,
    spreadsheetName: row.spreadsheetName,
    sheetName: row.sheetName,
    rowNumber: row.rowNumber,
    root: row.root,
    appt: row.appt,
    so: row.so,
    taskId: row.taskId,
    matchedCustomerName: row.matchedCustomerName || '',
    matchedEmail: row.matchedEmail || '',
    seedMatchType: row.seedMatchType || '',
    seedMatchedField: row.seedMatchedField || '',
    seedMatchedValue: row.seedMatchedValue || '',
    seedReason: row.seedReason || '',
    matchType: row.matchType || '',
    customerName: row.customerName,
    email: row.email,
    brand: row.brand,
    visitDate: row.visitDate,
    status: row.status,
    paymentId: row.paymentId,
    docNumber: row.docNumber,
    matchedField: row.matchedField,
    matchedValue: row.matchedValue,
    reason: row.reason
  };
}

function swTestDataCleanupLooksLikeNameField_(key) {
  return /(name|customer|client|lead|contact|guest|booked|payer)/i.test(key);
}

function swTestDataCleanupLooksLikeNameValue_(value) {
  value = swTestDataCleanupTrim_(value);
  if (!value) return false;
  if (value.length > 80) return false;
  if (!/[a-z]/i.test(value)) return false;
  if (/@/.test(value)) return false;
  if (/^[0-9\s._\-+,#]+$/.test(value)) return false;
  return true;
}

function swTestDataCleanupLooksLikeEmailField_(key) {
  return /email|mail/.test(key);
}

function swTestDataCleanupLooksLikeEmailValue_(value) {
  value = swTestDataCleanupNormEmail_(value);
  if (!value) return '';
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value)) return '';
  return value;
}

function swPublicTestDataCleanupKeysetCounts_(keyset) {
  return {
    roots: Object.keys(keyset.roots || {}).length,
    appts: Object.keys(keyset.appts || {}).length,
    uids: Object.keys(keyset.uids || {}).length,
    taskIds: Object.keys(keyset.taskIds || {}).length,
    sos: Object.keys(keyset.sos || {}).length,
    emails: Object.keys(keyset.emails || {}).length,
    phones: Object.keys(keyset.phones || {}).length
  };
}

function swLogTestDataCleanupPlan_(plan, label) {
  try {
    Logger.log(label + ' ' + JSON.stringify(swPublicTestDataCleanupResult_(plan)));
  } catch (_) {}
}

function swRecordTestDataCleanupResultFailures_(plan, result, type) {
  if (!swTestDataCleanupResultHasFailure_(result)) return;
  plan.errors.push({
    type: type || 'postDelete',
    message: 'Post-delete ' + (type || 'operation') + ' reported a failure.',
    result: result
  });
}

function swTestDataCleanupResultHasFailure_(value) {
  if (!value || typeof value !== 'object') return false;
  if (value.ok === false) return true;
  if (Array.isArray(value)) {
    for (var i = 0; i < value.length; i++) {
      if (swTestDataCleanupResultHasFailure_(value[i])) return true;
    }
    return false;
  }
  var keys = Object.keys(value);
  for (var k = 0; k < keys.length; k++) {
    if (swTestDataCleanupResultHasFailure_(value[keys[k]])) return true;
  }
  return false;
}

function swTestDataCleanupFingerprint_(plan) {
  var rows = (plan.candidates || []).map(function (row) {
    return [
      row.workbookKey,
      row.spreadsheetId,
      row.sheetId,
      row.rowNumber,
      row.root,
      row.appt,
      row.uid,
      row.taskId,
      row.so,
      swTestDataCleanupNorm_(row.customerName),
      swTestDataCleanupNormEmail_(row.email),
      row.paymentId,
      row.docNumber
    ].join('|');
  });
  return swTestDataCleanupHash_(rows.join('\n'));
}

function swTestDataCleanupRowFingerprint_(rec) {
  rec = rec || {};
  return swTestDataCleanupHash_([
    rec.root || '',
    rec.appt || '',
    rec.uid || '',
    rec.taskId || '',
    rec.so || '',
    swTestDataCleanupNorm_(rec.customerName || ''),
    swTestDataCleanupNormEmail_(rec.email || ''),
    rec.paymentId || '',
    rec.docNumber || ''
  ].join('|'));
}

function swTestDataCleanupHash_(text) {
  text = String(text || '');
  try {
    var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
    return bytes.map(function (b) {
      var v = b < 0 ? b + 256 : b;
      return ('0' + v.toString(16)).slice(-2);
    }).join('');
  } catch (_) {}

  var hash = 0;
  for (var i = 0; i < text.length; i++) {
    hash = ((hash << 5) - hash) + text.charCodeAt(i);
    hash |= 0;
  }
  return ('00000000' + (hash >>> 0).toString(16)).slice(-8);
}

function swCompareTestDataCleanupCandidates_(a, b) {
  if (a.deleteOrder !== b.deleteOrder) return a.deleteOrder - b.deleteOrder;
  if (a.workbookKey !== b.workbookKey) return a.workbookKey < b.workbookKey ? -1 : 1;
  if (a.sheetName !== b.sheetName) return a.sheetName < b.sheetName ? -1 : 1;
  return a.rowNumber - b.rowNumber;
}

function swTestDataCleanupHeaderMap_(headers) {
  var map = {};
  (headers || []).forEach(function (header, idx) {
    var raw = swTestDataCleanupTrim_(header);
    if (!raw) return;
    if (map[raw] == null) map[raw] = idx;
    var key = swTestDataCleanupHeaderKey_(raw);
    if (map[key] == null) map[key] = idx;
  });
  return map;
}

function swTestDataCleanupPick_(map, aliases) {
  for (var i = 0; i < aliases.length; i++) {
    var raw = aliases[i];
    if (map[raw] != null) return map[raw];
    var key = swTestDataCleanupHeaderKey_(raw);
    if (map[key] != null) return map[key];
  }
  return -1;
}

function swTestDataCleanupCell_(row, idx) {
  idx = Number(idx);
  return isFinite(idx) && idx >= 0 ? row[idx] : '';
}

function swTestDataCleanupWorkflowSpreadsheet_() {
  if (typeof swSpreadsheet_ === 'function') return swSpreadsheet_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet and swSpreadsheet_ is unavailable.');
  return ss;
}

function swTestDataCleanupSpreadsheetId_(ss) {
  try { return ss.getId(); } catch (_) {}
  return '';
}

function swTestDataCleanupSpreadsheetName_(ss) {
  try { return ss.getName(); } catch (_) {}
  return '';
}

function swAddTestDataCleanupWarning_(plan, message) {
  if (!plan.warnings) plan.warnings = [];
  if (plan.warnings.indexOf(message) < 0) plan.warnings.push(message);
}

function swUniqueStrings_(values) {
  var seen = {};
  var out = [];
  (values || []).forEach(function (value) {
    value = swTestDataCleanupTrim_(value);
    if (!value || seen[value]) return;
    seen[value] = true;
    out.push(value);
  });
  return out;
}

function swTestDataCleanupHeaderKey_(value) {
  return swTestDataCleanupNorm_(value).replace(/[^a-z0-9]+/g, '');
}

function swTestDataCleanupNorm_(value) {
  return swTestDataCleanupTrim_(value).toLowerCase();
}

function swTestDataCleanupTrim_(value) {
  return String(value == null ? '' : value).trim();
}

function swTestDataCleanupNormEmail_(value) {
  return swTestDataCleanupTrim_(value).toLowerCase();
}

function swTestDataCleanupNormPhone_(value) {
  return swTestDataCleanupTrim_(value).replace(/\D/g, '');
}

function swTestDataCleanupCleanId_(value) {
  return swTestDataCleanupTrim_(value).replace(/^'/, '');
}

function swTestDataCleanupIso_(date) {
  try {
    if (typeof swIso_ === 'function') return swIso_(date);
  } catch (_) {}
  return date && date.toISOString ? date.toISOString() : String(date || '');
}

function swTestDataCleanupDateMs_(value) {
  if (!value) return 0;
  var d = new Date(value);
  var ms = d.getTime();
  return isNaN(ms) ? 0 : ms;
}

var SW_PAYMENT_TEST_DATA_CLEANUP_PREVIEW_PROPERTY_ = 'SW_PAYMENT_TEST_DATA_CLEANUP_ONCE_LAST_PREVIEW';
var SW_PAYMENT_TEST_DATA_CLEANUP_PREVIEW_TAB_NAME_ = 'PaymentTestDataCleanup_Preview';
var SW_PAYMENT_TEST_DATA_CLEANUP_PREVIEW_MAX_AGE_MS_ = 4 * 60 * 60 * 1000;

function sw_previewPaymentLedgerTestDataCleanupOnce(options) {
  options = options || {};
  swRequireTestDataCleanupAdmin_(options);
  var plan = swBuildPaymentLedgerTestDataCleanupPlan_(options);
  swWritePaymentLedgerTestDataCleanupPreviewTab_(plan);
  swStorePaymentLedgerTestDataCleanupPreview_(plan);
  return swPublicPaymentLedgerTestDataCleanupResult_(plan);
}

function sw_applyPaymentLedgerTestDataCleanupOnce(options) {
  options = options || {};
  swRequireTestDataCleanupAdmin_(options);
  var token = swTestDataCleanupTrim_(options.confirmationToken || '');
  var plan = swBuildPaymentLedgerTestDataCleanupPlan_(options);
  var validation = swValidatePaymentLedgerTestDataCleanupPreview_(plan, token);
  plan.previewValidation = validation;
  if (!validation.ok) {
    plan.ok = false;
    plan.errors.push({ type: 'preview', message: validation.message });
    return swPublicPaymentLedgerTestDataCleanupResult_(plan);
  }
  var ledger = swPaymentLedgerCleanupTarget_();
  var rows = (plan.candidates || []).map(function (row) { return Number(row.rowNumber); }).filter(Boolean).sort(function (a, b) { return b - a; });
  for (var i = 0; i < rows.length; i++) ledger.sheet.deleteRow(rows[i]);
  try { SpreadsheetApp.flush(); } catch (_) {}
  plan.apply = true;
  plan.deletedCount = rows.length;
  plan.deletedRows = plan.candidates.slice();
  return swPublicPaymentLedgerTestDataCleanupResult_(plan);
}

function sw_inspectPaymentLedgerTestDataNeedlesJson(options) {
  options = options || {};
  swRequireTestDataCleanupAdmin_(options);
  var target = swTestDataCleanupPaymentLedgerApiTarget_();
  if (!target) throw new Error('Payment ledger target could not be resolved.');
  var needles = options.needles || [
    'test',
    'fake',
    'dummy',
    'sample',
    'testdemo3',
    'testdemo34',
    'testdemo5',
    'AP-20260504-002',
    'AP-20260505-003',
    'AP-20260505-005',
    'AP-20260505-006',
    'E2E_CLEANUP_20260506_175656_1FB8C3A0'
  ];
  var mentions = swSheetsApiSearchTargets_([target], needles, {
    maxMatchesPerNeedle: Number(options.maxMatchesPerNeedle || 100)
  });
  return JSON.stringify({ ok: true, target: target, mentions: mentions }, null, 2);
}

function swBuildPaymentLedgerTestDataCleanupPlan_(options) {
  options = options || {};
  var target = swTestDataCleanupPaymentLedgerApiTarget_();
  if (!target) throw new Error('Payment ledger target could not be resolved.');
  var apiRead = swSheetsApiReadValues_(target.spreadsheetId, target.sheetName, 'A1:ZZ');
  if (!apiRead || apiRead.ok === false) {
    throw new Error('Unable to read payment ledger through Sheets API: ' + ((apiRead && apiRead.error) || 'unknown error'));
  }
  var values = apiRead.values || [];
  var headers = values.length ? values[0] : [];
  var map = swPaymentLedgerCleanupHeaderMap_(headers);
  var known = swPaymentLedgerKnownTestIdentitySet_(options);
  var plan = {
    ok: true,
    apply: false,
    createdAt: swTestDataCleanupIso_(new Date()),
    spreadsheetId: target.spreadsheetId,
    spreadsheetName: target.spreadsheetName,
    sheetName: target.sheetName,
    apiRead: {
      ok: true,
      status: apiRead.status || 200,
      range: apiRead.range || '',
      rowCount: Math.max(0, values.length - 1)
    },
    candidates: [],
    deletedRows: [],
    deletedCount: 0,
    errors: [],
    warnings: [
      'Preview uses the live 400 Payments ledger via the Sheets API.',
      'Cleanup candidates are limited to exact known test identities, strict test-like customer/email identity fields, or exact test appointment/payment identifiers.',
      'Free-text notes and memo fields do not seed cleanup candidates.'
    ]
  };

  for (var r = 1; r < values.length; r++) {
    var row = values[r] || [];
    var rec = swPaymentLedgerCleanupRecordFromRow_(row, map);
    var match = swPaymentLedgerTestDataMatch_(rec, row, known, options);
    if (!match.matched) continue;
    plan.candidates.push(swPaymentLedgerTestDataCandidate_(target, r + 1, rec, match));
  }

  plan.totalCandidateRows = plan.candidates.length;
  plan.fingerprint = swPaymentLedgerTestDataFingerprint_(plan);
  plan.confirmationToken = 'DELETE_PAYMENT_TEST_DATA_' + plan.fingerprint.slice(0, 10).toUpperCase();
  return plan;
}

function swPaymentLedgerCleanupHeaderMap_(headers) {
  var out = {};
  for (var i = 0; i < (headers || []).length; i++) {
    var raw = swTestDataCleanupTrim_(headers[i]);
    var key = swTestDataCleanupHeaderKey_(raw);
    if (key && out[key] == null) out[key] = i;
  }
  return out;
}

function swPaymentLedgerCleanupPick_(map, names) {
  for (var i = 0; i < names.length; i++) {
    var key = swTestDataCleanupHeaderKey_(names[i]);
    if (map[key] != null) return map[key];
  }
  return -1;
}

function swPaymentLedgerCleanupCell_(row, idx) {
  idx = Number(idx);
  if (!isFinite(idx) || idx < 0) return '';
  return swTestDataCleanupTrim_(row[idx]);
}

function swPaymentLedgerCleanupRecordFromRow_(row, map) {
  var cols = {
    root: swPaymentLedgerCleanupPick_(map, ['RootApptID', 'Root Appt ID', 'Root Appointment ID', 'ROOT']),
    appt: swPaymentLedgerCleanupPick_(map, ['APPT_ID', 'Appt ID', 'Appointment ID']),
    so: swPaymentLedgerCleanupPick_(map, ['SO#', 'SO #', 'SO', 'SO Number', 'Sales Order', 'Sales Order #']),
    paymentId: swPaymentLedgerCleanupPick_(map, ['PAYMENT_ID', 'Payment ID', 'PaymentId', 'Payment Id', 'ID']),
    docNumber: swPaymentLedgerCleanupPick_(map, ['DocNumber', 'Doc Number', 'Doc #', 'Document Number']),
    name: swPaymentLedgerCleanupPick_(map, ['Customer Name', 'Customer', 'Client Name', 'Name', 'Payer', 'Payer Name', 'Contact Name']),
    email: swPaymentLedgerCleanupPick_(map, ['Email', 'Email Address', 'Customer Email', 'Client Email', 'Payer Email', 'Contact Email']),
    phone: swPaymentLedgerCleanupPick_(map, ['Phone', 'Phone Number', 'PhoneNorm', 'Phone Norm', 'Mobile']),
    amount: swPaymentLedgerCleanupPick_(map, ['Amount', 'Payment Amount', 'Paid Amount', 'Total', 'Gross Amount']),
    date: swPaymentLedgerCleanupPick_(map, ['PaymentDateTime', 'Payment DateTime', 'Payment Date', 'Date', 'Created At']),
    status: swPaymentLedgerCleanupPick_(map, ['Status', 'Payment Status', 'DocStatus', 'Doc Status']),
    description: swPaymentLedgerCleanupPick_(map, ['Description', 'Desc', 'Line Item', 'Line Item Description', 'Item', 'Product', 'Service']),
    memo: swPaymentLedgerCleanupPick_(map, ['Memo', 'Note', 'Notes', 'Internal Notes', 'Payment Notes'])
  };
  var payloadJson = swTestDataCleanupFindJsonInRow_(row);
  var rec = {
    root: swTestDataCleanupCleanId_(swPaymentLedgerCleanupCell_(row, cols.root)),
    appt: swTestDataCleanupCleanId_(swPaymentLedgerCleanupCell_(row, cols.appt)),
    so: swTestDataCleanupCleanId_(swPaymentLedgerCleanupCell_(row, cols.so)),
    paymentId: swTestDataCleanupCleanId_(swPaymentLedgerCleanupCell_(row, cols.paymentId)),
    docNumber: swTestDataCleanupCleanId_(swPaymentLedgerCleanupCell_(row, cols.docNumber)),
    customerName: swTestDataCleanupTrim_(swPaymentLedgerCleanupCell_(row, cols.name)),
    email: swTestDataCleanupNormEmail_(swPaymentLedgerCleanupCell_(row, cols.email)),
    phone: swTestDataCleanupNormPhone_(swPaymentLedgerCleanupCell_(row, cols.phone)),
    amount: swTestDataCleanupTrim_(swPaymentLedgerCleanupCell_(row, cols.amount)),
    paymentDate: swTestDataCleanupTrim_(swPaymentLedgerCleanupCell_(row, cols.date)),
    status: swTestDataCleanupTrim_(swPaymentLedgerCleanupCell_(row, cols.status)),
    description: swTestDataCleanupTrim_(swPaymentLedgerCleanupCell_(row, cols.description)),
    memo: swTestDataCleanupTrim_(swPaymentLedgerCleanupCell_(row, cols.memo)),
    payloadJson: payloadJson
  };
  swMergePayloadIntoTestDataCleanupRecord_(rec);
  if (!rec.email) rec.email = swTestDataCleanupFindEmailInRow_(row);
  return rec;
}

function swPaymentLedgerKnownTestIdentitySet_(options) {
  var set = {
    names: {},
    emails: {},
    roots: {},
    appts: {},
    paymentIds: {},
    docNumbers: {},
    sos: {}
  };
  function add(bucket, value) {
    value = swTestDataCleanupTrim_(value);
    if (!value) return;
    if (bucket === 'emails') value = swTestDataCleanupNormEmail_(value);
    set[bucket][value] = true;
  }
  [
    'testdemo3',
    'testdemo34',
    'testdemo5',
    'Test Customer Cleanup E2E 20260506_175656 1FB8C3A0'
  ].forEach(function (v) { add('names', v); });
  [
    'testdemo3@gmail.com',
    'testdemo34@gmail.com',
    'testdemo5@gmail.com',
    'test.cleanup.e2e.20260506_175656.1fb8c3a0@example.com'
  ].forEach(function (v) { add('emails', v); });
  [
    'AP-20260504-002',
    'AP-20260505-003',
    'AP-20260505-005',
    'AP-20260505-006',
    'E2E_CLEANUP_20260506_175656_1FB8C3A0'
  ].forEach(function (v) { add('roots', v); add('appts', v); });

  var extra = options.extraTestIdentities || {};
  Object.keys(extra.names || {}).forEach(function (v) { if (extra.names[v]) add('names', v); });
  Object.keys(extra.emails || {}).forEach(function (v) { if (extra.emails[v]) add('emails', v); });
  Object.keys(extra.roots || {}).forEach(function (v) { if (extra.roots[v]) add('roots', v); });
  Object.keys(extra.appts || {}).forEach(function (v) { if (extra.appts[v]) add('appts', v); });
  Object.keys(extra.paymentIds || {}).forEach(function (v) { if (extra.paymentIds[v]) add('paymentIds', v); });
  Object.keys(extra.docNumbers || {}).forEach(function (v) { if (extra.docNumbers[v]) add('docNumbers', v); });
  return set;
}

function swPaymentLedgerTestDataMatch_(rec, row, known, options) {
  var name = swTestDataCleanupTrim_(rec.customerName);
  var email = swTestDataCleanupNormEmail_(rec.email);
  var compactName = swTestDataCleanupNorm_(name).replace(/[^a-z0-9]+/g, '');
  var emailLocal = email ? email.split('@')[0] : '';
  var desc = swTestDataCleanupTrim_(rec.description);
  if (name && known.names[name]) return swPaymentLedgerMatch_('directName', 'Customer Name', name, 'customer name is an exact known test customer');
  if (email && known.emails[email]) return swPaymentLedgerMatch_('directEmail', 'Email', email, 'email is an exact known test customer email');
  if (rec.root && known.roots[rec.root]) return swPaymentLedgerMatch_('root', 'RootApptID', rec.root, 'root appointment ID is a known test appointment');
  if (rec.appt && known.appts[rec.appt]) return swPaymentLedgerMatch_('appt', 'APPT_ID', rec.appt, 'appointment ID is a known test appointment');
  if (rec.paymentId && known.paymentIds[rec.paymentId]) return swPaymentLedgerMatch_('paymentId', 'PAYMENT_ID', rec.paymentId, 'payment ID is a known test payment');
  if (rec.docNumber && known.docNumbers[rec.docNumber]) return swPaymentLedgerMatch_('docNumber', 'DocNumber', rec.docNumber, 'document number is a known test payment');
  if (name && swTestDataCleanupTextLooksTest_(name, 'strict')) return swPaymentLedgerMatch_('directName', 'Customer Name', name, 'customer name contains a strict test-data token');
  if (email && swTestDataCleanupEmailLooksTest_(email, 'strict')) return swPaymentLedgerMatch_('directEmail', 'Email', email, 'email contains a strict test-data token');
  if (compactName && /^test[a-z0-9]*$/.test(compactName)) return swPaymentLedgerMatch_('directName', 'Customer Name', name, 'customer name begins with test in an identity field');
  if (emailLocal && /^test[a-z0-9._+-]*$/.test(emailLocal)) return swPaymentLedgerMatch_('directEmail', 'Email', email, 'email local-part begins with test');
  if (desc && /(^|[^a-z0-9])(test|testing|tester|sample|dummy|fake)([0-9_-]*)([^a-z0-9]|$)/i.test(desc)) {
    return swPaymentLedgerMatch_('directDescription', 'Description', desc.slice(0, 120), 'payment description contains a test-data token');
  }
  return { matched: false };
}

function swPaymentLedgerMatch_(type, field, value, reason) {
  return { matched: true, matchType: type, matchedField: field, matchedValue: swTestDataCleanupTrim_(value), reason: reason };
}

function swPaymentLedgerTestDataCandidate_(target, rowNumber, rec, match) {
  return {
    spreadsheetId: target.spreadsheetId,
    spreadsheetName: target.spreadsheetName,
    sheetName: target.sheetName,
    rowNumber: rowNumber,
    matchType: match.matchType,
    matchedField: match.matchedField,
    matchedValue: match.matchedValue,
    reason: match.reason,
    customerName: rec.customerName || '',
    email: rec.email || '',
    phone: rec.phone || '',
    root: rec.root || '',
    appt: rec.appt || '',
    so: rec.so || '',
    paymentId: rec.paymentId || '',
    docNumber: rec.docNumber || '',
    paymentDate: rec.paymentDate || '',
    amount: rec.amount || '',
    status: rec.status || '',
    description: rec.description || '',
    memo: rec.memo || ''
  };
}

function swWritePaymentLedgerTestDataCleanupPreviewTab_(plan) {
  var ledger = swPaymentLedgerCleanupTarget_();
  var ss = ledger.spreadsheet;
  var sh = ss.getSheetByName(SW_PAYMENT_TEST_DATA_CLEANUP_PREVIEW_TAB_NAME_);
  if (!sh) sh = ss.insertSheet(SW_PAYMENT_TEST_DATA_CLEANUP_PREVIEW_TAB_NAME_);
  else sh.clear();
  if (!plan.candidates.length) {
    sh.getRange(1, 1).setValue('No payment test-data cleanup candidates found.');
    try { SpreadsheetApp.flush(); } catch (_) {}
    plan.reviewSheet = { ok: true, sheetName: SW_PAYMENT_TEST_DATA_CLEANUP_PREVIEW_TAB_NAME_, rowCount: 0, url: ss.getUrl() + '#gid=' + sh.getSheetId() };
    return plan.reviewSheet;
  }
  var headers = [
    'Created At', 'Spreadsheet', 'Sheet', 'Row', 'Match Type', 'Matched Field', 'Matched Value', 'Reason',
    'Customer Name', 'Email', 'Phone', 'RootApptID', 'APPT_ID', 'SO#', 'Payment ID', 'Doc Number',
    'Payment Date', 'Amount', 'Status', 'Description', 'Memo'
  ];
  var now = swTestDataCleanupIso_(new Date());
  var rows = plan.candidates.map(function (c) {
    return [now, c.spreadsheetName, c.sheetName, c.rowNumber, c.matchType, c.matchedField, c.matchedValue, c.reason,
      c.customerName, c.email, c.phone, c.root, c.appt, c.so, c.paymentId, c.docNumber, c.paymentDate, c.amount,
      c.status, c.description, c.memo];
  });
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  try { SpreadsheetApp.flush(); } catch (_) {}
  plan.reviewSheet = { ok: true, sheetName: SW_PAYMENT_TEST_DATA_CLEANUP_PREVIEW_TAB_NAME_, rowCount: rows.length, url: ss.getUrl() + '#gid=' + sh.getSheetId() };
  return plan.reviewSheet;
}

function swPaymentLedgerCleanupTarget_() {
  var target = null;
  if (typeof rp_getLedgerTarget === 'function') target = rp_getLedgerTarget();
  if ((!target || !target.sh) && typeof pr_getLedger_ === 'function' && typeof pr_getPaymentsSheet_ === 'function') {
    var ss = pr_getLedger_();
    target = { ss: ss, sh: pr_getPaymentsSheet_(ss) };
  }
  if (!target || !target.sh) throw new Error('Could not resolve 400 Payments ledger sheet.');
  return { spreadsheet: target.ss || target.sh.getParent(), sheet: target.sh };
}

function swStorePaymentLedgerTestDataCleanupPreview_(plan) {
  PropertiesService.getScriptProperties().setProperty(SW_PAYMENT_TEST_DATA_CLEANUP_PREVIEW_PROPERTY_, JSON.stringify({
    createdAtMs: new Date().getTime(),
    fingerprint: plan.fingerprint,
    confirmationToken: plan.confirmationToken,
    totalCandidateRows: plan.totalCandidateRows
  }));
}

function swReadPaymentLedgerTestDataCleanupPreview_() {
  var raw = PropertiesService.getScriptProperties().getProperty(SW_PAYMENT_TEST_DATA_CLEANUP_PREVIEW_PROPERTY_);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (_) { return null; }
}

function swValidatePaymentLedgerTestDataCleanupPreview_(plan, token) {
  var stored = swReadPaymentLedgerTestDataCleanupPreview_();
  if (!stored) return { ok: false, message: 'Run sw_previewPaymentLedgerTestDataCleanupOnce() before apply.' };
  if (stored.createdAtMs && new Date().getTime() - stored.createdAtMs > SW_PAYMENT_TEST_DATA_CLEANUP_PREVIEW_MAX_AGE_MS_) {
    return { ok: false, message: 'Stored payment cleanup preview is older than 4 hours. Re-run preview.' };
  }
  if (stored.fingerprint !== plan.fingerprint) return { ok: false, message: 'Payment cleanup preview fingerprint changed. Re-run preview.' };
  if (!token || token !== stored.confirmationToken) return { ok: false, message: 'Confirmation token mismatch.' };
  return { ok: true, message: 'Preview matches and token was confirmed.' };
}

function swPaymentLedgerTestDataFingerprint_(plan) {
  var payload = (plan.candidates || []).map(function (c) {
    return [c.sheetName, c.rowNumber, c.matchType, c.matchedField, c.matchedValue, c.paymentId, c.docNumber, c.root, c.appt, c.customerName, c.email].join('|');
  }).join('\n');
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, payload || 'empty-payment-cleanup');
  return digest.map(function (b) { return ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2); }).join('');
}

function swPublicPaymentLedgerTestDataCleanupResult_(plan) {
  var counts = {};
  (plan.candidates || []).forEach(function (c) { counts[c.matchType] = (counts[c.matchType] || 0) + 1; });
  return {
    ok: plan.ok !== false,
    apply: plan.apply === true,
    createdAt: plan.createdAt,
    spreadsheetId: plan.spreadsheetId,
    spreadsheetName: plan.spreadsheetName,
    sheetName: plan.sheetName,
    apiRead: plan.apiRead,
    totalCandidateRows: plan.totalCandidateRows || 0,
    deletedCount: plan.deletedCount || 0,
    confirmationToken: plan.confirmationToken,
    fingerprint: plan.fingerprint,
    reviewSheet: plan.reviewSheet || null,
    matchTypeCounts: counts,
    warnings: plan.warnings || [],
    errors: plan.errors || [],
    candidates: plan.candidates || [],
    deletedRows: plan.deletedRows || [],
    previewValidation: plan.previewValidation || null
  };
}
