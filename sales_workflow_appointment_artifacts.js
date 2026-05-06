/**
 * Appointment artifacts: upload tracking, AssemblyAI transcription, OpenAI summaries,
 * Client Advisor approval, and JOC handoff state.
 */

var SW_APPOINTMENT_ARTIFACT_SHEET = '_AppointmentArtifacts';
var SW_APPOINTMENT_ARTIFACT_HEADERS = [
  'ArtifactID',
  'RootApptID',
  'APPT_ID',
  'TaskID',
  'Artifact Type',
  'Workflow Stage',
  'Original Filename',
  'Canonical Filename',
  'Mime Type',
  'Size Bytes',
  'Drive File ID',
  'Drive URL',
  'Folder ID',
  'Uploaded By',
  'Uploaded By Email',
  'Uploaded At',
  'Assembly Upload URL',
  'Assembly Source',
  'Assembly Transcript ID',
  'Assembly Status',
  'Transcript Doc ID',
  'Transcript Doc URL',
  'Summary JSON File ID',
  'Summary JSON URL',
  'Summary Doc ID',
  'Summary Doc URL',
  'Sales Brief',
  'Client Follow-Up Draft',
  'Review Flags',
  'Summary Snapshot JSON',
  'Attempts',
  'Next Poll At',
  'Last Error',
  'Approved By',
  'Approved By Email',
  'Approved At',
  'JOC Handoff At',
  'Updated At'
];

var SW_ARTIFACT_TYPES = {
  APPOINTMENT_RECORDING: 'APPOINTMENT_RECORDING',
  CLIENT_ADVISOR_RECAP: 'CLIENT_ADVISOR_RECAP',
  DIAMOND_VIEWING_RECORDING: 'DIAMOND_VIEWING_RECORDING',
  CLIENT_INTAKE: 'CLIENT_INTAKE'
};

var SW_ARTIFACT_STAGES = {
  UPLOADED: 'UPLOADED',
  TRANSCRIPTION_QUEUED: 'TRANSCRIPTION_QUEUED',
  TRANSCRIBING: 'TRANSCRIBING',
  TRANSCRIPT_READY: 'TRANSCRIPT_READY',
  SUMMARY_QUEUED: 'SUMMARY_QUEUED',
  SUMMARY_READY: 'SUMMARY_READY',
  REVIEW_PENDING: 'REVIEW_PENDING',
  APPROVED: 'APPROVED',
  JOC_HANDOFF: 'JOC_HANDOFF',
  DUPLICATE_SKIPPED: 'DUPLICATE_SKIPPED',
  ERROR: 'ERROR'
};

var SW_ARTIFACT_UPLOAD_FIELDS = [
  { field: 'appointmentRecording', driveUrlField: 'appointmentRecordingDriveUrl', type: SW_ARTIFACT_TYPES.APPOINTMENT_RECORDING },
  { field: 'advisorRecap', driveUrlField: 'advisorRecapDriveUrl', type: SW_ARTIFACT_TYPES.CLIENT_ADVISOR_RECAP },
  { field: 'diamondViewingRecording', driveUrlField: 'diamondViewingRecordingDriveUrl', type: SW_ARTIFACT_TYPES.DIAMOND_VIEWING_RECORDING },
  { field: 'intakeMaterial', driveUrlField: 'intakeMaterialDriveUrl', type: SW_ARTIFACT_TYPES.CLIENT_INTAKE }
];
var SW_APPOINTMENT_ARTIFACT_ROOT_CACHE_SECONDS = 2 * 60;
var SW_APPOINTMENT_AI_BRIEF_CACHE_SECONDS = 10 * 60;
var SW_PUBLIC_APPOINTMENT_ARTIFACT_INDEX_CACHE_SECONDS = 10 * 60;
var SW_APPOINTMENT_FOLDER_ID_CACHE_SECONDS = 60 * 60;
var SW_APPOINTMENT_UPLOAD_FOLDER_CACHE_SECONDS = 60 * 60;
var SW_APPOINTMENT_ARTIFACT_ROOT_MEMORY_CACHE_ = {};
var SW_APPOINTMENT_AI_BRIEF_MEMORY_CACHE_ = {};
var SW_PUBLIC_APPOINTMENT_ARTIFACT_INDEX_MEMORY_CACHE_ = {};
var SW_APPOINTMENT_ROOT_FOLDER_MEMORY_CACHE_ = {};
var SW_APPOINTMENT_UPLOAD_FOLDER_MEMORY_CACHE_ = {};

function swEnsureAppointmentArtifactsSheet_(ss) {
  var sh = ss.getSheetByName(SW_APPOINTMENT_ARTIFACT_SHEET);
  if (!sh) sh = ss.insertSheet(SW_APPOINTMENT_ARTIFACT_SHEET);
  var changed = false;
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, SW_APPOINTMENT_ARTIFACT_HEADERS.length).setValues([SW_APPOINTMENT_ARTIFACT_HEADERS]);
    changed = true;
  } else if (swAppointmentArtifactsSheetNeedsMigration_(sh)) {
    swMigrateAppointmentArtifactsSheet_(sh);
    changed = true;
  }
  if (changed) swStyleSheet_(sh);
  return sh;
}

function swAppointmentArtifactsSheetNeedsMigration_(sh) {
  if (sh.getLastColumn() !== SW_APPOINTMENT_ARTIFACT_HEADERS.length) return true;
  var headers = sh.getRange(1, 1, 1, SW_APPOINTMENT_ARTIFACT_HEADERS.length).getDisplayValues()[0];
  for (var i = 0; i < SW_APPOINTMENT_ARTIFACT_HEADERS.length; i++) {
    if (swHeaderKey_(headers[i]) !== swHeaderKey_(SW_APPOINTMENT_ARTIFACT_HEADERS[i])) return true;
  }
  return false;
}

function swMigrateAppointmentArtifactsSheet_(sh) {
  var lastRow = Math.max(sh.getLastRow(), 1);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var currentHeaders = values[0].map(function (h) { return swTrim_(h); });
  var H = swHeaderMapFromArray_(currentHeaders);
  var aliases = {};
  var nextValues = values.map(function (_, rowIndex) {
    return SW_APPOINTMENT_ARTIFACT_HEADERS.map(function (header) {
      if (rowIndex === 0) return header;
      var sourceHeaders = aliases[header] || [header];
      for (var i = 0; i < sourceHeaders.length; i++) {
        var idx = H[sourceHeaders[i]];
        if (idx == null) idx = H[swHeaderKey_(sourceHeaders[i])];
        if (idx != null && values[rowIndex][idx] !== '') return values[rowIndex][idx];
      }
      return '';
    });
  });
  if (sh.getMaxColumns() < SW_APPOINTMENT_ARTIFACT_HEADERS.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), SW_APPOINTMENT_ARTIFACT_HEADERS.length - sh.getMaxColumns());
  }
  sh.getRange(1, 1, nextValues.length, SW_APPOINTMENT_ARTIFACT_HEADERS.length).setValues(nextValues);
  var extraCols = sh.getLastColumn() - SW_APPOINTMENT_ARTIFACT_HEADERS.length;
  if (extraCols > 0) sh.deleteColumns(SW_APPOINTMENT_ARTIFACT_HEADERS.length + 1, extraCols);
}

function sw_cleanupAppointmentArtifactsSchema() {
  var ss = swSpreadsheet_();
  var sh = swEnsureAppointmentArtifactsSheet_(ss);
  return {
    ok: true,
    sheet: SW_APPOINTMENT_ARTIFACT_SHEET,
    columns: sh.getLastColumn(),
    headers: SW_APPOINTMENT_ARTIFACT_HEADERS
  };
}

function swArtifactHeaderMap_(sh) {
  var values = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), SW_APPOINTMENT_ARTIFACT_HEADERS.length))
    .getDisplayValues()[0];
  var map = {};
  values.forEach(function (h, i) {
    h = swTrim_(h);
    if (h) map[h] = i + 1;
  });
  return map;
}

function swReadAppointmentArtifactRows_(ss) {
  var sh = swEnsureAppointmentArtifactsSheet_(ss);
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = swReadSheetObjectsExpectedHeaders_(sh, SW_APPOINTMENT_ARTIFACT_HEADERS);
  return rows.map(function (row) {
    row.rowNumber = row.__rowNumber || row.rowNumber || 0;
    return row;
  });
}

function swAppointmentArtifactRowsForRoot_(ss, rootApptId) {
  var root = swTrim_(rootApptId);
  if (!root) return [];
  var cached = swCachedAppointmentArtifactRowsForRoot_(ss, root);
  if (cached !== null) return cached;

  var sh = swEnsureAppointmentArtifactsSheet_(ss);
  if (!sh || sh.getLastRow() < 2) return [];
  var rootCol = SW_APPOINTMENT_ARTIFACT_HEADERS.indexOf('RootApptID') + 1;
  if (swHeaderKey_(sh.getRange(1, rootCol).getDisplayValue()) !== swHeaderKey_('RootApptID')) {
    var fallbackRows = swReadAppointmentArtifactRows_(ss).filter(function (row) {
      return swTrim_(row['RootApptID']) === root;
    });
    swCacheAppointmentArtifactRowsForRoot_(ss, root, fallbackRows);
    return fallbackRows;
  }
  var rowCount = sh.getLastRow() - 1;
  var width = Math.min(sh.getLastColumn(), SW_APPOINTMENT_ARTIFACT_HEADERS.length);
  var roots = sh.getRange(2, rootCol, rowCount, 1).getDisplayValues();
  var out = [];
  for (var i = 0; i < roots.length; i++) {
    if (swTrim_(roots[i][0]) !== root) continue;
    var rowNumber = i + 2;
    out.push(swAppointmentArtifactRowAtNumber_(sh, rowNumber, width));
  }
  swCacheAppointmentArtifactRowsForRoot_(ss, root, out);
  return out;
}

function swAppointmentArtifactRowAtNumber_(sh, rowNumber, width) {
  var values = sh.getRange(rowNumber, 1, 1, width).getDisplayValues()[0];
  var row = { __rowNumber: rowNumber, rowNumber: rowNumber };
  for (var j = 0; j < SW_APPOINTMENT_ARTIFACT_HEADERS.length; j++) {
    row[SW_APPOINTMENT_ARTIFACT_HEADERS[j]] = j < values.length ? values[j] : '';
  }
  return row;
}

function swCachedAppointmentArtifactRowsForRoot_(ss, root) {
  var key = swAppointmentArtifactRootCacheKey_(ss, root);
  try {
    var memory = SW_APPOINTMENT_ARTIFACT_ROOT_MEMORY_CACHE_[key];
    if (memory && memory.expiresAt > new Date().getTime()) return memory.rows || [];
  } catch (_) {}

  try {
    var cached = CacheService.getScriptCache().get(key);
    if (cached == null) return null;
    var rows = swParseJson_(cached, []);
    SW_APPOINTMENT_ARTIFACT_ROOT_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_APPOINTMENT_ARTIFACT_ROOT_CACHE_SECONDS * 1000,
      rows: rows || []
    };
    return rows || [];
  } catch (_) {}
  return null;
}

function swCacheAppointmentArtifactRowsForRoot_(ss, root, rows) {
  var key = swAppointmentArtifactRootCacheKey_(ss, root);
  rows = rows || [];
  try {
    SW_APPOINTMENT_ARTIFACT_ROOT_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_APPOINTMENT_ARTIFACT_ROOT_CACHE_SECONDS * 1000,
      rows: rows
    };
  } catch (_) {}
  try {
    var text = swStringify_(rows);
    if (text.length < 90000) CacheService.getScriptCache().put(key, text, SW_APPOINTMENT_ARTIFACT_ROOT_CACHE_SECONDS);
  } catch (_) {}
}

function swInvalidateAppointmentArtifactRowsForRoot_(ss, root) {
  root = swTrim_(root);
  if (!root) return;
  var key = swAppointmentArtifactRootCacheKey_(ss, root);
  try { delete SW_APPOINTMENT_ARTIFACT_ROOT_MEMORY_CACHE_[key]; } catch (_) {}
  try { CacheService.getScriptCache().remove(key); } catch (_) {}
  try { swInvalidateAppointmentAiBriefCache_(ss, root); } catch (_) {}
  try { swInvalidatePublicAppointmentArtifactsIndex_(ss); } catch (_) {}
}

function swAppointmentArtifactRootCacheKey_(ss, root) {
  return 'sw:artifactRoot:v1:' + ss.getId() + ':' + encodeURIComponent(root);
}

function swPublicAppointmentArtifacts_(ss, rootApptId) {
  var root = swTrim_(rootApptId);
  var cached = swCachedPublicAppointmentArtifactsForRoot_(ss, root);
  if (cached !== null) return cached;
  var out = swAppointmentArtifactRowsForRoot_(ss, root).map(swPublicAppointmentArtifactFromRow_).sort(swPublicAppointmentArtifactSort_);
  return out;
}

function swPublicAppointmentArtifactFromRow_(row) {
  row = row || {};
  return {
    artifactId: row['ArtifactID'] || '',
    rootApptId: row['RootApptID'] || '',
    apptId: row['APPT_ID'] || '',
    artifactType: row['Artifact Type'] || '',
    typeLabel: swArtifactTypeLabel_(row['Artifact Type']),
    workflowStage: row['Workflow Stage'] || '',
    stageLabel: swArtifactStageLabel_(row['Workflow Stage']),
    originalFilename: row['Original Filename'] || '',
    canonicalFilename: row['Canonical Filename'] || '',
    mimeType: row['Mime Type'] || '',
    sizeBytes: row['Size Bytes'] || '',
    driveUrl: row['Drive URL'] || '',
    transcriptDocUrl: row['Transcript Doc URL'] || '',
    summaryDocUrl: row['Summary Doc URL'] || '',
    summaryJsonUrl: row['Summary JSON URL'] || '',
    uploadedBy: row['Uploaded By'] || '',
    uploadedAt: row['Uploaded At'] || '',
    assemblySource: row['Assembly Source'] || '',
    assemblyStatus: row['Assembly Status'] || '',
    lastError: row['Last Error'] || '',
    updatedAt: row['Updated At'] || ''
  };
}

function swPublicAppointmentArtifactSort_(a, b) {
    return String(b.uploadedAt || b.updatedAt || '').localeCompare(String(a.uploadedAt || a.updatedAt || ''));
}

function swPublicAppointmentArtifactsIndexKey_(ss) {
  return 'sw:publicAppointmentArtifacts:v1:' + ss.getId();
}

function swCachePublicAppointmentArtifactsIndex_(ss) {
  return swCachePublicAppointmentArtifactsIndexFromRows_(ss, swReadAppointmentArtifactRows_(ss));
}

function swCachePublicAppointmentArtifactsIndexFromRows_(ss, rows) {
  var key = swPublicAppointmentArtifactsIndexKey_(ss);
  var byRoot = {};
  (rows || []).forEach(function (row) {
    var root = swTrim_(row['RootApptID']);
    if (!root) return;
    if (!byRoot[root]) byRoot[root] = [];
    byRoot[root].push(swPublicAppointmentArtifactFromRow_(row));
  });
  Object.keys(byRoot).forEach(function (root) {
    byRoot[root].sort(swPublicAppointmentArtifactSort_);
  });
  var payload = {
    cachedAt: swIso_(new Date()),
    byRoot: byRoot
  };
  try {
    SW_PUBLIC_APPOINTMENT_ARTIFACT_INDEX_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_PUBLIC_APPOINTMENT_ARTIFACT_INDEX_CACHE_SECONDS * 1000,
      byRoot: byRoot
    };
  } catch (_) {}
  if (typeof swTaskListCachePut_ === 'function') return swTaskListCachePut_(key, payload);
  return { ok: false, reason: 'chunkCacheUnavailable', chunks: 0, bytes: 0 };
}

function swCachedPublicAppointmentArtifactsForRoot_(ss, root) {
  root = swTrim_(root);
  if (!root) return [];
  var key = swPublicAppointmentArtifactsIndexKey_(ss);
  var now = new Date().getTime();
  try {
    var memory = SW_PUBLIC_APPOINTMENT_ARTIFACT_INDEX_MEMORY_CACHE_[key];
    if (memory && memory.expiresAt > now && memory.byRoot) {
      return memory.byRoot[root] || [];
    }
  } catch (_) {}
  try {
    var payload = typeof swTaskListCacheGet_ === 'function' ? swTaskListCacheGet_(key) : null;
    if (!payload || !payload.byRoot) return null;
    SW_PUBLIC_APPOINTMENT_ARTIFACT_INDEX_MEMORY_CACHE_[key] = {
      expiresAt: now + SW_PUBLIC_APPOINTMENT_ARTIFACT_INDEX_CACHE_SECONDS * 1000,
      byRoot: payload.byRoot || {}
    };
    return payload.byRoot[root] || [];
  } catch (_) {}
  return null;
}

function swInvalidatePublicAppointmentArtifactsIndex_(ss) {
  var key = swPublicAppointmentArtifactsIndexKey_(ss);
  try { delete SW_PUBLIC_APPOINTMENT_ARTIFACT_INDEX_MEMORY_CACHE_[key]; } catch (_) {}
  try { if (typeof swTaskListCacheRemove_ === 'function') swTaskListCacheRemove_(key); } catch (_) {}
}

function swArtifactTypeLabel_(type) {
  var map = {};
  map[SW_ARTIFACT_TYPES.APPOINTMENT_RECORDING] = 'Initial Consult Recording';
  map[SW_ARTIFACT_TYPES.CLIENT_ADVISOR_RECAP] = 'Client Advisor Recap';
  map[SW_ARTIFACT_TYPES.DIAMOND_VIEWING_RECORDING] = 'Diamond Viewing Recording';
  map[SW_ARTIFACT_TYPES.CLIENT_INTAKE] = 'Client Intake / Photo';
  return map[type] || swTrim_(type);
}

function swArtifactStageLabel_(stage) {
  return swTrim_(stage).replace(/_/g, ' ').toLowerCase().replace(/\b\w/g, function (c) {
    return c.toUpperCase();
  });
}

function swArtifactStageRank_(stage) {
  var rank = {};
  rank[SW_ARTIFACT_STAGES.UPLOADED] = 10;
  rank[SW_ARTIFACT_STAGES.TRANSCRIPTION_QUEUED] = 20;
  rank[SW_ARTIFACT_STAGES.TRANSCRIBING] = 30;
  rank[SW_ARTIFACT_STAGES.TRANSCRIPT_READY] = 40;
  rank[SW_ARTIFACT_STAGES.SUMMARY_QUEUED] = 50;
  rank[SW_ARTIFACT_STAGES.SUMMARY_READY] = 60;
  rank[SW_ARTIFACT_STAGES.REVIEW_PENDING] = 70;
  rank[SW_ARTIFACT_STAGES.APPROVED] = 80;
  rank[SW_ARTIFACT_STAGES.JOC_HANDOFF] = 90;
  rank[SW_ARTIFACT_STAGES.DUPLICATE_SKIPPED] = -2;
  rank[SW_ARTIFACT_STAGES.ERROR] = -1;
  return rank[stage] || 0;
}

function swArtifactNeedsTranscription_(type) {
  return type === SW_ARTIFACT_TYPES.APPOINTMENT_RECORDING ||
    type === SW_ARTIFACT_TYPES.DIAMOND_VIEWING_RECORDING;
}

function swAppointmentHasPrimaryRecording_(ss, rootApptId) {
  return swAppointmentArtifactRowsForRoot_(ss, rootApptId).some(function (row) {
    return row['Artifact Type'] === SW_ARTIFACT_TYPES.APPOINTMENT_RECORDING &&
      row['Workflow Stage'] !== SW_ARTIFACT_STAGES.ERROR &&
      swTrim_(row['Drive File ID']);
  });
}

function swPrimarySummaryArtifactForRoot_(ss, rootApptId) {
  var readyRank = swArtifactStageRank_(SW_ARTIFACT_STAGES.SUMMARY_READY);
  var rows = swAppointmentArtifactRowsForRoot_(ss, rootApptId).filter(function (row) {
    return row['Artifact Type'] === SW_ARTIFACT_TYPES.APPOINTMENT_RECORDING &&
      swArtifactStageRank_(row['Workflow Stage']) >= readyRank &&
      swTrim_(row['Client Follow-Up Draft']);
  });
  if (!rows.length) return null;
  rows.sort(function (a, b) {
    var aPrimary = a['Artifact Type'] === SW_ARTIFACT_TYPES.APPOINTMENT_RECORDING ? 1 : 0;
    var bPrimary = b['Artifact Type'] === SW_ARTIFACT_TYPES.APPOINTMENT_RECORDING ? 1 : 0;
    if (aPrimary !== bPrimary) return bPrimary - aPrimary;
    return String(b['Updated At'] || b['Uploaded At'] || '').localeCompare(String(a['Updated At'] || a['Uploaded At'] || ''));
  });
  return rows[0];
}

function swSummaryExtraForRoot_(ss, rootApptId) {
  var row = swPrimarySummaryArtifactForRoot_(ss, rootApptId);
  if (!row) return { ready: false };
  return swSummaryExtraFromArtifactRow_(row);
}

function swAppointmentAiBriefForRoot_(ss, rootApptId) {
  var cached = swCachedAppointmentAiBrief_(ss, rootApptId);
  if (cached) return cached.hasAiBrief ? cached : null;
  var summary = swSummaryExtraForRoot_(ss, rootApptId);
  var brief = summary && summary.ready ? swAppointmentAiBriefFull_(summary) : null;
  swCacheAppointmentAiBrief_(ss, rootApptId, brief);
  return brief;
}

function swAppointmentSummaryIndex_(ss) {
  var readyRank = swArtifactStageRank_(SW_ARTIFACT_STAGES.SUMMARY_READY);
  var byRoot = {};
  var rows = swReadAppointmentArtifactRows_(ss);
  try { swCachePublicAppointmentArtifactsIndexFromRows_(ss, rows); } catch (_) {}
  rows.forEach(function (row) {
    if (row['Artifact Type'] !== SW_ARTIFACT_TYPES.APPOINTMENT_RECORDING) return;
    if (swArtifactStageRank_(row['Workflow Stage']) < readyRank) return;
    if (!swTrim_(row['Client Follow-Up Draft'])) return;
    var root = swTrim_(row['RootApptID']);
    if (!root) return;
    if (!byRoot[root] || swCompareSummaryArtifacts_(row, byRoot[root]) < 0) byRoot[root] = row;
  });
  var out = {};
  Object.keys(byRoot).forEach(function (root) {
    out[root] = swSummaryExtraFromArtifactRow_(byRoot[root]);
  });
  return out;
}

function swAppointmentAiBriefIndex_(ss) {
  var summaries = swAppointmentSummaryIndex_(ss);
  var out = {};
  Object.keys(summaries || {}).forEach(function (root) {
    var brief = swAppointmentAiBriefFull_(summaries[root]);
    if (brief && brief.hasAiBrief) {
      out[root] = brief;
      swCacheAppointmentAiBrief_(ss, root, brief);
    }
  });
  return out;
}

function swAppointmentAiBriefCacheKey_(ss, root) {
  return 'sw:appointmentAiBrief:v1:' + ss.getId() + ':' + encodeURIComponent(swTrim_(root));
}

function swCachedAppointmentAiBrief_(ss, root) {
  root = swTrim_(root);
  if (!root) return null;
  var key = swAppointmentAiBriefCacheKey_(ss, root);
  try {
    var memory = SW_APPOINTMENT_AI_BRIEF_MEMORY_CACHE_[key];
    if (memory && memory.expiresAt > new Date().getTime()) return memory.value || { hasAiBrief: false };
  } catch (_) {}
  try {
    var cached = CacheService.getScriptCache().get(key);
    if (!cached) return null;
    var value = swParseJson_(cached, { hasAiBrief: false });
    SW_APPOINTMENT_AI_BRIEF_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_APPOINTMENT_AI_BRIEF_CACHE_SECONDS * 1000,
      value: value || { hasAiBrief: false }
    };
    return value || { hasAiBrief: false };
  } catch (_) {}
  return null;
}

function swCacheAppointmentAiBrief_(ss, root, brief) {
  root = swTrim_(root);
  if (!root) return;
  var key = swAppointmentAiBriefCacheKey_(ss, root);
  var value = brief && brief.hasAiBrief ? brief : { hasAiBrief: false };
  try {
    SW_APPOINTMENT_AI_BRIEF_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_APPOINTMENT_AI_BRIEF_CACHE_SECONDS * 1000,
      value: value
    };
  } catch (_) {}
  try {
    var text = swStringify_(value);
    if (text.length <= 90000) CacheService.getScriptCache().put(key, text, SW_APPOINTMENT_AI_BRIEF_CACHE_SECONDS);
  } catch (_) {}
}

function swInvalidateAppointmentAiBriefCache_(ss, root) {
  root = swTrim_(root);
  if (!root) return;
  var key = swAppointmentAiBriefCacheKey_(ss, root);
  try { delete SW_APPOINTMENT_AI_BRIEF_MEMORY_CACHE_[key]; } catch (_) {}
  try { CacheService.getScriptCache().remove(key); } catch (_) {}
}

function swAppointmentAiBriefFull_(summary) {
  if (!(summary && summary.ready)) return null;
  var flags = swAppointmentReviewFlagsFromValue_(summary.reviewFlags);
  return {
    hasAiBrief: true,
    ready: true,
    artifactId: summary.artifactId || '',
    workflowStage: summary.workflowStage || '',
    workflowStageLabel: swArtifactStageLabel_(summary.workflowStage || ''),
    salesBrief: summary.salesBrief || '',
    reviewFlags: flags,
    reviewFlagCount: flags.length,
    clientFollowUpDraft: summary.clientFollowUpDraft || '',
    transcriptDocUrl: summary.transcriptDocUrl || '',
    summaryDocUrl: summary.summaryDocUrl || '',
    summaryJsonUrl: summary.summaryJsonUrl || '',
    latestAiBriefUpdatedAt: summary.updatedAt || summary.uploadedAt || ''
  };
}

function swAppointmentAiBriefCompact_(brief) {
  if (!(brief && brief.hasAiBrief)) {
    return { hasAiBrief: false, reviewFlagCount: 0, latestAiBriefUpdatedAt: '' };
  }
  return {
    hasAiBrief: true,
    reviewFlagCount: Number(brief.reviewFlagCount || 0),
    latestAiBriefUpdatedAt: brief.latestAiBriefUpdatedAt || ''
  };
}

function swAppointmentReviewFlagsFromValue_(value) {
  if (Array.isArray(value)) {
    return value.map(swTrim_).filter(Boolean);
  }
  value = swTrim_(value);
  if (!value) return [];
  var parsed = swParseJson_(value, null);
  if (Array.isArray(parsed)) return parsed.map(swTrim_).filter(Boolean);
  return value.split(/\r?\n/).map(swTrim_).filter(Boolean);
}

function swCompareSummaryArtifacts_(a, b) {
  var aPrimary = a['Artifact Type'] === SW_ARTIFACT_TYPES.APPOINTMENT_RECORDING ? 1 : 0;
  var bPrimary = b['Artifact Type'] === SW_ARTIFACT_TYPES.APPOINTMENT_RECORDING ? 1 : 0;
  if (aPrimary !== bPrimary) return bPrimary - aPrimary;
  return String(b['Updated At'] || b['Uploaded At'] || '').localeCompare(String(a['Updated At'] || a['Uploaded At'] || ''));
}

function swSummaryExtraFromArtifactRow_(row) {
  return {
    ready: true,
    artifactId: row['ArtifactID'] || '',
    workflowStage: row['Workflow Stage'] || '',
    transcriptDocUrl: row['Transcript Doc URL'] || '',
    summaryDocUrl: row['Summary Doc URL'] || '',
    summaryJsonUrl: row['Summary JSON URL'] || '',
    salesBrief: row['Sales Brief'] || '',
    reviewFlags: row['Review Flags'] || '',
    clientFollowUpDraft: row['Client Follow-Up Draft'] || '',
    recapDraft: row['Client Follow-Up Draft'] || '',
    approvedText: row['Client Follow-Up Draft'] || '',
    uploadedAt: row['Uploaded At'] || '',
    updatedAt: row['Updated At'] || ''
  };
}

function sw_uploadAppointmentArtifacts(form) {
  var ss = swSpreadsheet_();
  sw_setupSalesWorkflow();
  form = form || {};
  var user = swAuthUserForApi_(ss, form.swAuthToken || form.authToken || '');
  var taskId = swTrim_(form.taskId || '');
  var task = taskId ? swGetTaskById_(ss, taskId) : null;
  if (task && !swCanActOnTask_(task, user)) throw new Error('You are not the current owner for this appointment task.');

  var root = swTrim_(form.rootApptId || (task && task.root) || '');
  if (!root) throw new Error('Missing RootApptID for appointment upload.');

  var created = [];
  SW_ARTIFACT_UPLOAD_FIELDS.forEach(function (slot) {
    var blobs = swFormBlobs_(form[slot.field]);
    blobs.forEach(function (blob) {
      created.push(swCreateAppointmentArtifactFromBlob_(ss, root, taskId, slot.type, blob, user, {}));
    });
    swFormTextValues_(form[slot.driveUrlField]).forEach(function (driveUrl) {
      created.push(swCreateAppointmentArtifactFromDriveFile_(ss, root, taskId, slot.type, driveUrl, user, {}));
    });
  });

  if (!created.length) throw new Error('Choose at least one file or paste a Drive file link.');
  return {
    ok: true,
    uploaded: created.length,
    artifacts: swPublicAppointmentArtifacts_(ss, root)
  };
}

function swFormBlobs_(value) {
  if (!value) return [];
  var values = Array.isArray(value) ? value : [value];
  return values.filter(function (blob) {
    return blob && typeof blob.getBytes === 'function' && blob.getBytes().length;
  });
}

function swFormTextValues_(value) {
  if (!value) return [];
  var values = Array.isArray(value) ? value : [value];
  var out = [];
  values.forEach(function (item) {
    String(item || '').split(/\n+/).forEach(function (line) {
      line = swTrim_(line);
      if (line) out.push(line);
    });
  });
  return out;
}

function swCreateAppointmentArtifactFromBlob_(ss, rootApptId, taskId, artifactType, blob, user, options) {
  options = options || {};
  var now = new Date();
  var folders = swEnsureAppointmentFolderForRoot_(ss, rootApptId);
  var originalName = swTrim_(options.filename || (blob.getName && blob.getName()) || 'upload');
  var canonicalName = swArtifactCanonicalName_(rootApptId, artifactType, originalName, now);
  var targetFolder = swArtifactTargetFolder_(folders, artifactType);
  var bytes = blob.getBytes();
  var uploadBlob = blob.copyBlob ? blob.copyBlob() : Utilities.newBlob(bytes, blob.getContentType(), originalName);
  uploadBlob.setName(canonicalName);
  var file = targetFolder.createFile(uploadBlob);
  var appt = swAppointmentRecordForRoot_(ss, rootApptId);
  var needsTranscription = swArtifactNeedsTranscription_(artifactType);
  var stage = needsTranscription ? SW_ARTIFACT_STAGES.TRANSCRIPTION_QUEUED : SW_ARTIFACT_STAGES.UPLOADED;
  var record = {
    'ArtifactID': 'ART-' + Utilities.getUuid(),
    'RootApptID': rootApptId,
    'APPT_ID': appt && appt.appt ? appt.appt : '',
    'TaskID': taskId || '',
    'Artifact Type': artifactType,
    'Workflow Stage': stage,
    'Original Filename': originalName,
    'Canonical Filename': canonicalName,
    'Mime Type': blob.getContentType ? blob.getContentType() : '',
    'Size Bytes': bytes.length,
    'Drive File ID': file.getId(),
    'Drive URL': file.getUrl(),
    'Folder ID': targetFolder.getId(),
    'Uploaded By': user.name || user.email || 'System',
    'Uploaded By Email': user.email || '',
    'Uploaded At': swIso_(now),
    'Attempts': 0,
    'Next Poll At': needsTranscription ? swIso_(now) : '',
    'Last Error': '',
    'Updated At': swIso_(now)
  };
  swAppendAppointmentArtifactRow_(ss, record);
  return record;
}

function swCreateAppointmentArtifactFromDriveFile_(ss, rootApptId, taskId, artifactType, driveUrl, user, options) {
  options = options || {};
  var fileId = swDriveFileIdFromUrl_(driveUrl);
  if (!fileId) throw new Error('Could not read a Drive file ID from: ' + driveUrl);
  var sourceFile;
  try {
    sourceFile = DriveApp.getFileById(fileId);
  } catch (err) {
    throw new Error('Could not open Drive file. Confirm the link is shared with this Apps Script account: ' + driveUrl);
  }
  var now = new Date();
  var folders = swEnsureAppointmentFolderForRoot_(ss, rootApptId);
  var originalName = swTrim_(options.filename || sourceFile.getName() || 'drive-file');
  var canonicalName = swArtifactCanonicalName_(rootApptId, artifactType, originalName, now);
  var targetFolder = swArtifactTargetFolder_(folders, artifactType);
  var file = sourceFile.makeCopy(canonicalName, targetFolder);
  var appt = swAppointmentRecordForRoot_(ss, rootApptId);
  var needsTranscription = swArtifactNeedsTranscription_(artifactType);
  var stage = needsTranscription ? SW_ARTIFACT_STAGES.TRANSCRIPTION_QUEUED : SW_ARTIFACT_STAGES.UPLOADED;
  var record = {
    'ArtifactID': 'ART-' + Utilities.getUuid(),
    'RootApptID': rootApptId,
    'APPT_ID': appt && appt.appt ? appt.appt : '',
    'TaskID': taskId || '',
    'Artifact Type': artifactType,
    'Workflow Stage': stage,
    'Original Filename': originalName,
    'Canonical Filename': canonicalName,
    'Mime Type': file.getMimeType ? file.getMimeType() : '',
    'Size Bytes': file.getSize ? file.getSize() : '',
    'Drive File ID': file.getId(),
    'Drive URL': file.getUrl(),
    'Folder ID': targetFolder.getId(),
    'Uploaded By': user.name || user.email || 'System',
    'Uploaded By Email': user.email || '',
    'Uploaded At': swIso_(now),
    'Attempts': 0,
    'Next Poll At': needsTranscription ? swIso_(now) : '',
    'Last Error': '',
    'Updated At': swIso_(now)
  };
  swAppendAppointmentArtifactRow_(ss, record);
  return record;
}

function swAppendAppointmentArtifactRow_(ss, record) {
  var sh = swEnsureAppointmentArtifactsSheet_(ss);
  var H = swArtifactHeaderMap_(sh);
  var values = new Array(sh.getLastColumn()).fill('');
  Object.keys(H).forEach(function (header) {
    if (record[header] != null) values[H[header] - 1] = record[header];
  });
  sh.getRange(sh.getLastRow() + 1, 1, 1, values.length).setValues([values]);
  swInvalidateAppointmentArtifactRowsForRoot_(ss, record && record['RootApptID']);
}

function swPatchAppointmentArtifactRow_(ss, row, patch) {
  if (!row || !row.rowNumber) return;
  var sh = swEnsureAppointmentArtifactsSheet_(ss);
  var H = swArtifactHeaderMap_(sh);
  var values = sh.getRange(row.rowNumber, 1, 1, sh.getLastColumn()).getValues()[0];
  Object.keys(patch || {}).forEach(function (header) {
    var col = H[header];
    if (!col) return;
    values[col - 1] = patch[header] == null ? '' : patch[header];
  });
  sh.getRange(row.rowNumber, 1, 1, values.length).setValues([values]);
  swInvalidateAppointmentArtifactRowsForRoot_(ss, row['RootApptID']);
  if (patch && patch['RootApptID'] && patch['RootApptID'] !== row['RootApptID']) {
    swInvalidateAppointmentArtifactRowsForRoot_(ss, patch['RootApptID']);
  }
}

function swDriveFileIdFromUrl_(urlOrId) {
  if (typeof idFromAnyGoogleUrl_ === 'function') return idFromAnyGoogleUrl_(urlOrId);
  var s = swTrim_(urlOrId);
  var m = s.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  if (m) return m[1];
  m = s.match(/[-\w]{25,}/);
  return m ? m[0] : '';
}

function swResolveAppointmentRootFolderId_(ss, rootApptId) {
  var root = swTrim_(rootApptId);
  if (!root) return '';
  var cached = swCachedAppointmentRootFolderId_(ss, root);
  if (cached) return cached;

  var folderId = '';
  try {
    if (typeof getApFolderIdForRoot_ === 'function') folderId = getApFolderIdForRoot_(ss, root);
  } catch (_) {}
  if (!folderId) {
    try {
      if (typeof _resolveApFolderId_ === 'function') folderId = _resolveApFolderId_(ss, root);
    } catch (_) {}
  }
  folderId = swTrim_(folderId);
  if (folderId) swCacheAppointmentRootFolderId_(ss, root, folderId);
  return folderId;
}

function swCachedAppointmentRootFolderId_(ss, root) {
  var key = swAppointmentRootFolderCacheKey_(ss, root);
  try {
    var memory = SW_APPOINTMENT_ROOT_FOLDER_MEMORY_CACHE_[key];
    if (memory && memory.expiresAt > new Date().getTime()) return memory.folderId || '';
  } catch (_) {}

  try {
    var cached = CacheService.getScriptCache().get(key);
    if (!cached) return '';
    SW_APPOINTMENT_ROOT_FOLDER_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_APPOINTMENT_FOLDER_ID_CACHE_SECONDS * 1000,
      folderId: cached
    };
    return cached;
  } catch (_) {}
  return '';
}

function swCacheAppointmentRootFolderId_(ss, root, folderId) {
  folderId = swTrim_(folderId);
  if (!folderId) return;
  var key = swAppointmentRootFolderCacheKey_(ss, root);
  try {
    SW_APPOINTMENT_ROOT_FOLDER_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_APPOINTMENT_FOLDER_ID_CACHE_SECONDS * 1000,
      folderId: folderId
    };
  } catch (_) {}
  try {
    CacheService.getScriptCache().put(key, folderId, SW_APPOINTMENT_FOLDER_ID_CACHE_SECONDS);
  } catch (_) {}
}

function swAppointmentRootFolderCacheKey_(ss, root) {
  return 'sw:appointmentRootFolderId:v1:' + ss.getId() + ':' + encodeURIComponent(swTrim_(root));
}

function swAppointmentDriveDropFolderInfoForRoot_(ss, rootApptId, artifactType) {
  var root = swTrim_(rootApptId);
  if (!root) throw new Error('Missing RootApptID for appointment upload folder.');
  var type = swNormalizeDriveUploadArtifactType_(artifactType);
  var cached = swCachedAppointmentUploadFolderInfo_(ss, root, type);
  if (cached && cached.folderId && cached.url) return cached;

  var folder = swResolveAppointmentDriveDropFolder_(ss, root, type);
  var folderId = folder.getId();
  var info = {
    ok: true,
    rootApptId: root,
    artifactType: type,
    folderId: folderId,
    url: swDriveFolderUrlForId_(folderId)
  };
  swCacheAppointmentUploadFolderInfo_(ss, root, type, info);
  return info;
}

function swResolveAppointmentDriveDropFolder_(ss, rootApptId, artifactType) {
  var rootFolderId = swResolveAppointmentRootFolderId_(ss, rootApptId);
  if (!rootFolderId) throw new Error('No client appointment folder found for ' + rootApptId + '.');
  var rootFolder = DriveApp.getFolderById(rootFolderId);
  if (artifactType === SW_ARTIFACT_TYPES.APPOINTMENT_RECORDING) {
    return swGetOrCreateSubfolder_(swGetOrCreateSubfolder_(rootFolder, '01_Audio'), 'Initial Consult Recordings');
  }
  if (artifactType === SW_ARTIFACT_TYPES.DIAMOND_VIEWING_RECORDING) {
    return swGetOrCreateSubfolder_(swGetOrCreateSubfolder_(rootFolder, '01_Audio'), 'Diamond Viewing Recordings');
  }
  if (artifactType === SW_ARTIFACT_TYPES.CLIENT_ADVISOR_RECAP) {
    return swGetOrCreateSubfolder_(swGetOrCreateSubfolder_(rootFolder, '02_Materials'), 'Client Advisor Recaps');
  }
  return swArtifactNeedsTranscription_(artifactType)
    ? swGetOrCreateSubfolder_(rootFolder, '01_Audio')
    : swGetOrCreateSubfolder_(rootFolder, '02_Materials');
}

function swCachedAppointmentUploadFolderInfo_(ss, root, artifactType) {
  var key = swAppointmentUploadFolderCacheKey_(ss, root, artifactType);
  try {
    var memory = SW_APPOINTMENT_UPLOAD_FOLDER_MEMORY_CACHE_[key];
    if (memory && memory.expiresAt > new Date().getTime()) return memory.info || null;
  } catch (_) {}

  try {
    var cached = CacheService.getScriptCache().get(key);
    if (!cached) return null;
    var info = swParseJson_(cached, null);
    if (!info || !info.folderId || !info.url) return null;
    SW_APPOINTMENT_UPLOAD_FOLDER_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_APPOINTMENT_UPLOAD_FOLDER_CACHE_SECONDS * 1000,
      info: info
    };
    return info;
  } catch (_) {}
  return null;
}

function swCacheAppointmentUploadFolderInfo_(ss, root, artifactType, info) {
  if (!info || !info.folderId || !info.url) return;
  var key = swAppointmentUploadFolderCacheKey_(ss, root, artifactType);
  try {
    SW_APPOINTMENT_UPLOAD_FOLDER_MEMORY_CACHE_[key] = {
      expiresAt: new Date().getTime() + SW_APPOINTMENT_UPLOAD_FOLDER_CACHE_SECONDS * 1000,
      info: info
    };
  } catch (_) {}
  try {
    CacheService.getScriptCache().put(key, swStringify_(info), SW_APPOINTMENT_UPLOAD_FOLDER_CACHE_SECONDS);
  } catch (_) {}
}

function swAppointmentUploadFolderCacheKey_(ss, root, artifactType) {
  return 'sw:appointmentUploadFolder:v1:' + ss.getId() + ':' +
    encodeURIComponent(swTrim_(root)) + ':' + encodeURIComponent(swTrim_(artifactType));
}

function swDriveFolderUrlForId_(folderId) {
  folderId = swTrim_(folderId);
  return folderId ? 'https://drive.google.com/drive/folders/' + encodeURIComponent(folderId) : '';
}

function swEnsureAppointmentFolderForRoot_(ss, rootApptId) {
  var folderId = swResolveAppointmentRootFolderId_(ss, rootApptId);
  if (!folderId) throw new Error('No client appointment folder found for ' + rootApptId + '.');
  var ap = DriveApp.getFolderById(folderId);
  return {
    root: ap,
    audio: swGetOrCreateSubfolder_(ap, '01_Audio'),
    materials: swGetOrCreateSubfolder_(ap, '02_Materials'),
    transcripts: swGetOrCreateSubfolder_(ap, '03_Transcripts'),
    summaries: swGetOrCreateSubfolder_(ap, '04_Summaries')
  };
}

function swGetOrCreateSubfolder_(folder, name) {
  var it = folder.getFoldersByName(name);
  return it.hasNext() ? it.next() : folder.createFolder(name);
}

function swArtifactTargetFolder_(folders, artifactType) {
  if (swArtifactNeedsTranscription_(artifactType)) return folders.audio;
  return folders.materials;
}

function swArtifactDriveDropFolder_(folders, artifactType) {
  if (artifactType === SW_ARTIFACT_TYPES.APPOINTMENT_RECORDING) {
    return swGetOrCreateSubfolder_(folders.audio, 'Initial Consult Recordings');
  }
  if (artifactType === SW_ARTIFACT_TYPES.DIAMOND_VIEWING_RECORDING) {
    return swGetOrCreateSubfolder_(folders.audio, 'Diamond Viewing Recordings');
  }
  if (artifactType === SW_ARTIFACT_TYPES.CLIENT_ADVISOR_RECAP) {
    return swGetOrCreateSubfolder_(folders.materials, 'Client Advisor Recaps');
  }
  return swArtifactTargetFolder_(folders, artifactType);
}

function swDriveUploadArtifactTypes_() {
  return [
    SW_ARTIFACT_TYPES.APPOINTMENT_RECORDING,
    SW_ARTIFACT_TYPES.DIAMOND_VIEWING_RECORDING,
    SW_ARTIFACT_TYPES.CLIENT_ADVISOR_RECAP
  ];
}

function swNormalizeDriveUploadArtifactType_(artifactType) {
  var type = swTrim_(artifactType);
  return swDriveUploadArtifactTypes_().indexOf(type) >= 0 ? type : SW_ARTIFACT_TYPES.APPOINTMENT_RECORDING;
}

function swSyncAppointmentDriveUploads_(ss, rootApptId, taskId, user) {
  var root = swTrim_(rootApptId);
  if (!root) throw new Error('Missing RootApptID for Drive upload sync.');
  var folders = swEnsureAppointmentFolderForRoot_(ss, root);
  var existing = {};
  swAppointmentArtifactRowsForRoot_(ss, root).forEach(function (row) {
    var fileId = swTrim_(row['Drive File ID']);
    if (fileId) existing[fileId] = true;
  });

  var created = [];
  swDriveUploadArtifactTypes_().forEach(function (artifactType) {
    var folder = swArtifactDriveDropFolder_(folders, artifactType);
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      if (!file || (file.isTrashed && file.isTrashed())) continue;
      if (existing[file.getId()]) continue;
      created.push(swRegisterAppointmentDriveFile_(ss, root, taskId, artifactType, file, user, { rename: true }));
      existing[file.getId()] = true;
    }
  });
  return created;
}

function swRegisterAppointmentDriveFile_(ss, rootApptId, taskId, artifactType, file, user, options) {
  options = options || {};
  var now = new Date();
  var parents = file.getParents ? file.getParents() : null;
  var folderId = parents && parents.hasNext() ? parents.next().getId() : '';
  var originalName = swTrim_(options.originalName || file.getName() || 'drive-file');
  var canonicalName = swIsCanonicalArtifactFilename_(rootApptId, artifactType, originalName)
    ? originalName
    : swArtifactCanonicalName_(rootApptId, artifactType, originalName, now);
  if (options.rename !== false && canonicalName !== originalName) {
    file.setName(canonicalName);
  }
  var appt = swAppointmentRecordForRoot_(ss, rootApptId);
  var needsTranscription = swArtifactNeedsTranscription_(artifactType);
  var stage = needsTranscription ? SW_ARTIFACT_STAGES.TRANSCRIPTION_QUEUED : SW_ARTIFACT_STAGES.UPLOADED;
  var record = {
    'ArtifactID': 'ART-' + Utilities.getUuid(),
    'RootApptID': rootApptId,
    'APPT_ID': appt && appt.appt ? appt.appt : '',
    'TaskID': taskId || '',
    'Artifact Type': artifactType,
    'Workflow Stage': stage,
    'Original Filename': originalName,
    'Canonical Filename': canonicalName,
    'Mime Type': file.getMimeType ? file.getMimeType() : '',
    'Size Bytes': file.getSize ? file.getSize() : '',
    'Drive File ID': file.getId(),
    'Drive URL': file.getUrl(),
    'Folder ID': folderId,
    'Uploaded By': user.name || user.email || 'System',
    'Uploaded By Email': user.email || '',
    'Uploaded At': swIso_(now),
    'Attempts': 0,
    'Next Poll At': needsTranscription ? swIso_(now) : '',
    'Last Error': '',
    'Updated At': swIso_(now)
  };
  swAppendAppointmentArtifactRow_(ss, record);
  return record;
}

function swIsCanonicalArtifactFilename_(rootApptId, artifactType, filename) {
  var prefix = swSafeFilePart_(rootApptId) + '__' + swSafeFilePart_(artifactType.toLowerCase()) + '__';
  return String(filename || '').indexOf(prefix) === 0;
}

function swArtifactCanonicalName_(rootApptId, artifactType, originalName, date) {
  var tz = swTimezone_();
  var stamp = Utilities.formatDate(date || new Date(), tz, 'yyyyMMdd-HHmmss');
  return [
    swSafeFilePart_(rootApptId),
    swSafeFilePart_(artifactType.toLowerCase()),
    stamp,
    swSafeFilename_(originalName || 'upload')
  ].join('__');
}

function swSafeFilePart_(value) {
  return swTrim_(value).replace(/[^A-Za-z0-9._-]+/g, '-').replace(/-+/g, '-').replace(/^-|-$/g, '') || 'value';
}

function swSafeFilename_(value) {
  var out = swTrim_(value).replace(/[\\/:*?"<>|#%{}~&]+/g, '-').replace(/\s+/g, '-');
  out = out.replace(/-+/g, '-').replace(/^-|-$/g, '');
  return out.slice(0, 120) || 'upload';
}

function swAppointmentRecordForRoot_(ss, rootApptId) {
  var rows = swReadAppointments_(ss);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].root === rootApptId || rows[i].appt === rootApptId) return rows[i];
  }
  return null;
}

function sw_ingestRawAppointmentUpload_(e) {
  try {
    var params = (e && e.parameter) || {};
    var tokenParam = swTrim_(params.token || '');
    var tokenWant = swTrim_(PropertiesService.getScriptProperties().getProperty('UPLOAD_TOKEN') || '');
    if (!tokenParam || tokenParam !== tokenWant) return ContentService.createTextOutput('ACK (invalid token)');
    if (!e || !e.postData || !e.postData.getBytes) return ContentService.createTextOutput('ACK (empty body)');

    var root = swTrim_(params.root_appt_id || params.rootApptId || '');
    if (!root) return ContentService.createTextOutput('ACK (missing RootApptID)');
    var filename = swTrim_(params.filename || (root + '__upload.m4a'));
    var mime = swTrim_(e.postData.type || '') || swMimeForFilename_(filename) || 'application/octet-stream';
    var type = swLegacyArtifactType_(params.rectype || params.type || '');
    var blob = Utilities.newBlob(e.postData.getBytes(), mime, filename);
    var ss = swSpreadsheet_();
    sw_setupSalesWorkflow();
    var user = {
      name: swTrim_(params.rep_name || params.uploaded_by || 'External Upload'),
      email: swNormEmail_(params.rep_email || ''),
      isAdmin: true
    };
    var created = swCreateAppointmentArtifactFromBlob_(ss, root, '', type, blob, user, { filename: filename });
    return ContentService.createTextOutput('ACK (queued ' + created['ArtifactID'] + ')');
  } catch (err) {
    return ContentService.createTextOutput('ACK (error: ' + swTrim_(err && err.message || err) + ')');
  }
}

function swLegacyArtifactType_(value) {
  var v = swNorm_(value).replace(/^\d+[.)]?\s*/, '');
  if (v === 'debrief' || v.indexOf('recap') >= 0) return SW_ARTIFACT_TYPES.CLIENT_ADVISOR_RECAP;
  if (v.indexOf('diamond') >= 0) return SW_ARTIFACT_TYPES.DIAMOND_VIEWING_RECORDING;
  return SW_ARTIFACT_TYPES.APPOINTMENT_RECORDING;
}

function swMimeForFilename_(filename) {
  var ext = String(filename || '').split('.').pop().toLowerCase();
  var map = {
    mp3: 'audio/mpeg',
    mpeg: 'audio/mpeg',
    mpga: 'audio/mpeg',
    mp4: 'audio/mp4',
    m4a: 'audio/mp4',
    wav: 'audio/wav',
    webm: 'audio/webm',
    mov: 'video/quicktime'
  };
  return map[ext] || '';
}

function swHandleAppointmentCompletion_(ss, task, data, user) {
  if (!task || task.taskType !== SW_TASKS.CHECKLIST) return null;
  var outcome = swTrim_(data.appointmentOutcome || '');
  if (!outcome) outcome = 'Completed';
  var root = task.root || task.appt || '';
  var rowsUpdated = swWriteAppointmentOutcomeToMaster_(ss, root, outcome);
  return {
    rootApptId: root,
    outcome: outcome,
    rowsUpdated: rowsUpdated,
    actor: user.name || user.email || ''
  };
}

function swWriteAppointmentOutcomeToMaster_(ss, rootApptId, outcome) {
  var sh = ss.getSheetByName(SW_SHEETS.MASTER);
  if (!sh) throw new Error('Missing sheet: ' + SW_SHEETS.MASTER);
  var headers = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getDisplayValues()[0].map(swTrim_);
  var H = swHeaderMapFromArray_(headers);
  var rootCol = swPickIndex_(H, ['RootApptID', 'Root Appt ID', 'APPT_ID']) + 1;
  if (!rootCol) throw new Error('Missing RootApptID/APPT_ID on Master.');
  var statusCol = swEnsureMasterColumnByAliases_(sh, headers, ['Status'], 'Status');
  var activeCol = swEnsureMasterColumnByAliases_(sh, headers, ['Active?', 'Active', 'Is Active'], 'Active?');
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var values = sh.getRange(2, rootCol, last - 1, 1).getDisplayValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    if (swTrim_(values[i][0]) === rootApptId) rows.push(i + 2);
  }
  rows.forEach(function (rowNumber) {
    if (swIsNoShowOutcome_(outcome)) {
      sh.getRange(rowNumber, statusCol).setValue('No Show');
      sh.getRange(rowNumber, activeCol).setValue('No');
    } else {
      sh.getRange(rowNumber, statusCol).setValue('Completed');
    }
  });
  return rows.length;
}

function swEnsureMasterColumnByAliases_(sh, headers, aliases, canonical) {
  var H = swHeaderMapFromArray_(headers);
  var idx = swPickIndex_(H, aliases);
  if (idx >= 0) return idx + 1;
  var col = sh.getLastColumn() + 1;
  sh.getRange(1, col).setValue(canonical);
  headers.push(canonical);
  return col;
}

function swAppointmentOutcomeForRoot_(state, rec) {
  var payload = swTaskPayload_(state, swTaskId_(rec, SW_TASKS.CHECKLIST));
  return swTrim_(swDeepValue_(payload, ['completion', 'appointmentOutcome']));
}

function swIsNoShowOutcome_(outcome) {
  return swNorm_(outcome).replace(/[-_]+/g, ' ') === 'no show';
}

function swMarkAppointmentSummaryApproved_(ss, rootApptId, approvedText, user) {
  var row = swPrimarySummaryArtifactForRoot_(ss, rootApptId);
  if (!row) return null;
  var now = swIso_(new Date());
  swPatchAppointmentArtifactRow_(ss, row, {
    'Workflow Stage': SW_ARTIFACT_STAGES.APPROVED,
    'Client Follow-Up Draft': approvedText || row['Client Follow-Up Draft'] || '',
    'Approved By': user.name || user.email || '',
    'Approved By Email': user.email || '',
    'Approved At': now,
    'Updated At': now
  });
  return { artifactId: row['ArtifactID'], stage: SW_ARTIFACT_STAGES.APPROVED };
}

function swMarkAppointmentJocHandoff_(ss, rootApptId, user) {
  var row = swPrimarySummaryArtifactForRoot_(ss, rootApptId);
  if (!row) return null;
  var now = swIso_(new Date());
  swPatchAppointmentArtifactRow_(ss, row, {
    'Workflow Stage': SW_ARTIFACT_STAGES.JOC_HANDOFF,
    'JOC Handoff At': now,
    'Updated At': now
  });
  return { artifactId: row['ArtifactID'], stage: SW_ARTIFACT_STAGES.JOC_HANDOFF };
}

function sw_processAppointmentAutomation(e) {
  var redirected = typeof swOrchRedirectLegacyTrigger_ === 'function'
    ? swOrchRedirectLegacyTrigger_('sw_processAppointmentAutomation', e)
    : null;
  if (redirected) return redirected;
  var options = e || {};

  return swTimed_('sw_processAppointmentAutomation', function () {
    var lock = LockService.getScriptLock();
    var lockWaitMs = Math.max(0, Number(options.lockWaitMs || 30000) || 0);
    var locked = false;
    try {
      locked = lock.tryLock(lockWaitMs);
    } catch (_) {
      locked = false;
    }
    if (!locked) {
      return {
        ok: true,
        skipped: true,
        reason: 'LOCK_BUSY',
        lockWaitMs: lockWaitMs,
        processed: 0,
        errors: 0,
        generatedTasks: false,
        at: swIso_(new Date())
      };
    }
    var refreshTasks = false;
    var summary = {
      ok: true,
      processed: 0,
      errors: 0,
      generatedTasks: false,
      at: swIso_(new Date())
    };
    try {
      sw_setupSalesWorkflow();
      var ss = swSpreadsheet_();
      var now = new Date();
      var rows = swReadAppointmentArtifactRows_(ss);
      for (var i = 0; i < rows.length && summary.processed < 8; i++) {
        var row = rows[i];
        if (!swArtifactReadyForWorker_(row, now)) continue;
        try {
          var result = swProcessAppointmentArtifact_(ss, row, now);
          if (result && result.refreshTasks) refreshTasks = true;
          summary.processed++;
        } catch (err) {
          summary.errors++;
          swHandleArtifactWorkerError_(ss, row, err);
        }
      }
      if (refreshTasks) {
        if (options.deferTaskGeneration) {
          summary.taskGenerationDeferred = true;
        } else {
          try {
            sw_generateSalesWorkflowTasks();
            summary.generatedTasks = true;
          } catch (genErr) {
            summary.errors++;
            summary.generationError = swTrim_(genErr && genErr.message || genErr);
          }
        }
      }
      return summary;
    } finally {
      if (locked) {
        try { lock.releaseLock(); } catch (_) {}
      }
    }
  });
}

function swArtifactReadyForWorker_(row, now) {
  var stage = row['Workflow Stage'];
  if (stage === SW_ARTIFACT_STAGES.ERROR) {
    return swArtifactCanRetryDriveSharingError_(row) || swArtifactCanRetryOpenAISummaryError_(row);
  }
  if (stage === SW_ARTIFACT_STAGES.DUPLICATE_SKIPPED ||
      stage === SW_ARTIFACT_STAGES.REVIEW_PENDING ||
      stage === SW_ARTIFACT_STAGES.APPROVED || stage === SW_ARTIFACT_STAGES.JOC_HANDOFF) return false;
  if (!swArtifactNeedsTranscription_(row['Artifact Type'])) return false;
  var due = swTrim_(row['Next Poll At']);
  if (!due) return true;
  var d = new Date(due);
  return isNaN(d.getTime()) || d.getTime() <= now.getTime();
}

function swProcessAppointmentArtifact_(ss, row, now) {
  var stage = row['Workflow Stage'];
  if (stage === SW_ARTIFACT_STAGES.ERROR && swArtifactCanRetryDriveSharingError_(row)) {
    swPatchAppointmentArtifactRow_(ss, row, {
      'Workflow Stage': SW_ARTIFACT_STAGES.TRANSCRIPTION_QUEUED,
      'Attempts': 0,
      'Next Poll At': swIso_(now),
      'Last Error': '',
      'Updated At': swIso_(now)
    });
    row['Workflow Stage'] = SW_ARTIFACT_STAGES.TRANSCRIPTION_QUEUED;
    row['Attempts'] = 0;
    row['Last Error'] = '';
    stage = row['Workflow Stage'];
  }
  if (stage === SW_ARTIFACT_STAGES.ERROR && swArtifactCanRetryOpenAISummaryError_(row)) {
    swPatchAppointmentArtifactRow_(ss, row, {
      'Workflow Stage': SW_ARTIFACT_STAGES.SUMMARY_QUEUED,
      'Attempts': 0,
      'Next Poll At': swIso_(now),
      'Last Error': '',
      'Updated At': swIso_(now)
    });
    row['Workflow Stage'] = SW_ARTIFACT_STAGES.SUMMARY_QUEUED;
    row['Attempts'] = 0;
    row['Last Error'] = '';
    stage = row['Workflow Stage'];
  }
  if (stage === SW_ARTIFACT_STAGES.UPLOADED || stage === SW_ARTIFACT_STAGES.TRANSCRIPTION_QUEUED) {
    swStartAssemblyTranscription_(ss, row, now);
    return { refreshTasks: false };
  }
  if (stage === SW_ARTIFACT_STAGES.TRANSCRIBING) {
    return swPollAssemblyTranscription_(ss, row, now);
  }
  if (stage === SW_ARTIFACT_STAGES.TRANSCRIPT_READY) {
    swPatchAppointmentArtifactRow_(ss, row, {
      'Workflow Stage': SW_ARTIFACT_STAGES.SUMMARY_QUEUED,
      'Next Poll At': swIso_(now),
      'Updated At': swIso_(now)
    });
    return { refreshTasks: false };
  }
  if (stage === SW_ARTIFACT_STAGES.SUMMARY_QUEUED) {
    var duplicate = swDuplicatePrimaryRecordingSummary_(ss, row);
    if (duplicate) {
      swPatchAppointmentArtifactRow_(ss, row, {
        'Workflow Stage': SW_ARTIFACT_STAGES.DUPLICATE_SKIPPED,
        'Next Poll At': '',
        'Last Error': '',
        'Updated At': swIso_(now)
      });
      Logger.log('Skipped duplicate appointment recording artifact ' + row['ArtifactID'] + '; primary is ' + duplicate.primaryArtifactId + '.');
      return { refreshTasks: false };
    }
    swGenerateAppointmentSummary_(ss, row, now);
    return { refreshTasks: true };
  }
  if (stage === SW_ARTIFACT_STAGES.SUMMARY_READY) {
    swPatchAppointmentArtifactRow_(ss, row, {
      'Workflow Stage': SW_ARTIFACT_STAGES.REVIEW_PENDING,
      'Updated At': swIso_(now)
    });
    return { refreshTasks: true };
  }
  return { refreshTasks: false };
}

function swDuplicatePrimaryRecordingSummary_(ss, row) {
  if (!row || row['Artifact Type'] !== SW_ARTIFACT_TYPES.APPOINTMENT_RECORDING) return null;
  var currentId = swTrim_(row['ArtifactID']);
  var root = swTrim_(row['RootApptID']);
  if (!currentId || !root) return null;
  var rows = swAppointmentArtifactRowsForRoot_(ss, root).filter(function (candidate) {
    if (candidate['Artifact Type'] !== SW_ARTIFACT_TYPES.APPOINTMENT_RECORDING) return false;
    if (candidate['Workflow Stage'] === SW_ARTIFACT_STAGES.ERROR ||
        candidate['Workflow Stage'] === SW_ARTIFACT_STAGES.DUPLICATE_SKIPPED) return false;
    return swTrim_(candidate['Transcript Doc ID']) || swArtifactStageRank_(candidate['Workflow Stage']) >= swArtifactStageRank_(SW_ARTIFACT_STAGES.SUMMARY_QUEUED);
  });
  if (rows.length <= 1) return null;

  var completed = rows.filter(function (candidate) {
    return swTrim_(candidate['Summary Doc ID']) ||
      swArtifactStageRank_(candidate['Workflow Stage']) >= swArtifactStageRank_(SW_ARTIFACT_STAGES.SUMMARY_READY);
  });
  var primary = (completed.length ? completed : rows).sort(swComparePrimaryRecordingArtifacts_)[0];
  var primaryId = swTrim_(primary && primary['ArtifactID']);
  return primaryId && primaryId !== currentId ? { primaryArtifactId: primaryId } : null;
}

function swComparePrimaryRecordingArtifacts_(a, b) {
  var aTime = swArtifactTimeMs_(a);
  var bTime = swArtifactTimeMs_(b);
  if (aTime !== bTime) return aTime - bTime;
  return String(a['ArtifactID'] || '').localeCompare(String(b['ArtifactID'] || ''));
}

function swArtifactTimeMs_(row) {
  var d = new Date(row['Uploaded At'] || row['Updated At'] || '');
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function swArtifactCanRetryDriveSharingError_(row) {
  if (!swArtifactNeedsTranscription_(row['Artifact Type'])) return false;
  if (swTrim_(row['Assembly Transcript ID'])) return false;
  if (!swTrim_(row['Drive File ID'])) return false;
  var lastError = String(row['Last Error'] || '');
  var oldSharingError = lastError.indexOf('Large recording cannot be shared') >= 0;
  var newSharingError = lastError.indexOf('Drive link sharing was denied') >= 0;
  if (!oldSharingError && !newSharingError) return false;
  if (oldSharingError) return true;
  var size = swBytesNumber_(row['Size Bytes']);
  return !size || size <= swAssemblyAppsScriptUploadMaxBytes_();
}

function swArtifactCanRetryOpenAISummaryError_(row) {
  if (!swArtifactNeedsTranscription_(row['Artifact Type'])) return false;
  if (!swTrim_(row['Transcript Doc ID'])) return false;
  if (swTrim_(row['Summary Doc ID'])) return false;
  var lastError = String(row['Last Error'] || '');
  return lastError.indexOf('OpenAI response incomplete') >= 0 &&
    lastError.indexOf('max_output_tokens') >= 0;
}

function swStartAssemblyTranscription_(ss, row, now) {
  var file = DriveApp.getFileById(row['Drive File ID']);
  var audioSource = swAssemblyAudioSourceForDriveFile_(file);
  var transcriptOne = swAssemblySubmitTranscript_(audioSource.audioUrl);
  swPatchAppointmentArtifactRow_(ss, row, {
    'Workflow Stage': SW_ARTIFACT_STAGES.TRANSCRIBING,
    'Assembly Upload URL': audioSource.audioUrl,
    'Assembly Source': audioSource.source,
    'Assembly Transcript ID': transcriptOne.id || '',
    'Assembly Status': transcriptOne.status || 'queued',
    'Attempts': 0,
    'Next Poll At': swIso_(swDateAddHours_(now, 5 / 60)),
    'Last Error': '',
    'Updated At': swIso_(now)
  });
}

function swPollAssemblyTranscription_(ss, row, now) {
  var id = swTrim_(row['Assembly Transcript ID']);
  if (!id) throw new Error('Missing AssemblyAI transcript job ID.');
  var transcript = swAssemblyGetTranscript_(id);
  var status = swNorm_(transcript.status);
  if (status === 'completed') {
    var text = swAssemblyTranscriptText_(transcript);
    swSaveTranscriptAndQueueSummary_(ss, row, text, now);
    return { refreshTasks: false };
  }
  if (status === 'error') {
    var err = new Error(transcript.error || 'AssemblyAI transcription failed.');
    err.terminal = true;
    throw err;
  }
  swPatchAppointmentArtifactRow_(ss, row, {
    'Assembly Status': transcript.status || 'processing',
    'Next Poll At': swIso_(swDateAddHours_(now, 5 / 60)),
    'Updated At': swIso_(now)
  });
  return { refreshTasks: false };
}

function swSaveTranscriptAndQueueSummary_(ss, row, transcriptText, now) {
  if (!swTrim_(transcriptText)) throw new Error('AssemblyAI returned an empty transcript.');
  var folders = swEnsureAppointmentFolderForRoot_(ss, row['RootApptID']);
  var doc = swCreateGoogleDocInFolder_(
    folders.transcripts,
    swTrim_(row['RootApptID']) + '__transcript__' + Utilities.formatDate(now, swTimezone_(), 'yyyyMMdd-HHmmss'),
    'Appointment Transcript',
    [
      'RootApptID: ' + row['RootApptID'],
      'Artifact: ' + swArtifactTypeLabel_(row['Artifact Type']),
      'Source file: ' + row['Canonical Filename']
    ],
    transcriptText
  );
  swPatchAppointmentArtifactRow_(ss, row, {
    'Workflow Stage': SW_ARTIFACT_STAGES.SUMMARY_QUEUED,
    'Assembly Status': 'completed',
    'Transcript Doc ID': doc.id,
    'Transcript Doc URL': doc.url,
    'Attempts': 0,
    'Next Poll At': swIso_(now),
    'Last Error': '',
    'Updated At': swIso_(now)
  });
}

function swGenerateAppointmentSummary_(ss, row, now) {
  var transcriptText = swReadGoogleDocText_(row['Transcript Doc ID']);
  if (!swTrim_(transcriptText)) throw new Error('Transcript document is empty.');
  var appointment = swAppointmentRecordForRoot_(ss, row['RootApptID']) || {};
  var result = swOpenAIAppointmentFollowUpDraft_(transcriptText, row, appointment);
  var normalized = swNormalizeAppointmentSummary_(result);
  var folders = swEnsureAppointmentFolderForRoot_(ss, row['RootApptID']);
  var baseName = row['RootApptID'] + '__client_follow_up_draft__' + Utilities.formatDate(now, swTimezone_(), 'yyyyMMdd-HHmmss');
  var jsonFile = folders.summaries.createFile(baseName + '.json', JSON.stringify(normalized, null, 2), 'application/json');
  var summaryDoc = swCreateGoogleDocInFolder_(
    folders.summaries,
    baseName,
    'AI Client Follow-Up Draft',
    [
      'RootApptID: ' + row['RootApptID'],
      'Customer: ' + (appointment.name || ''),
      'Client Advisor: ' + (appointment.assignedRep || '')
    ],
    swReadableAppointmentSummary_(normalized)
  );
  swPatchAppointmentArtifactRow_(ss, row, {
    'Workflow Stage': SW_ARTIFACT_STAGES.SUMMARY_READY,
    'Summary JSON File ID': jsonFile.getId(),
    'Summary JSON URL': jsonFile.getUrl(),
    'Summary Doc ID': summaryDoc.id,
    'Summary Doc URL': summaryDoc.url,
    'Sales Brief': normalized.salesBrief,
    'Client Follow-Up Draft': normalized.clientFollowUpDraft,
    'Review Flags': (normalized.reviewFlags || []).join('\n'),
    'Summary Snapshot JSON': JSON.stringify(normalized),
    'Attempts': 0,
    'Next Poll At': swIso_(now),
    'Last Error': '',
    'Updated At': swIso_(now)
  });
}

function swAssemblyApiKey_() {
  var key = swTrim_(PropertiesService.getScriptProperties().getProperty('ASSEMBLYAI_API_KEY'));
  if (!key) {
    var err = new Error('Missing ASSEMBLYAI_API_KEY in Script Properties.');
    err.terminal = true;
    throw err;
  }
  return key;
}

function swAssemblyBaseUrl_() {
  return swTrim_(PropertiesService.getScriptProperties().getProperty('ASSEMBLYAI_BASE_URL')) || 'https://api.assemblyai.com';
}

function swAssemblyAudioSourceForDriveFile_(file) {
  var size = Number(file.getSize && file.getSize()) || 0;
  if (!size || size <= swAssemblyAppsScriptUploadMaxBytes_()) {
    var uploadOne = swAssemblyUploadDriveFile_(file);
    return {
      audioUrl: uploadOne.upload_url || uploadOne.uploadUrl || '',
      source: 'ASSEMBLYAI_UPLOAD'
    };
  }
  return {
    audioUrl: swAssemblyDriveDownloadUrlForFile_(file, size),
    source: 'DRIVE_DIRECT_DOWNLOAD'
  };
}

function swAssemblyAppsScriptUploadMaxBytes_() {
  return 49 * 1000 * 1000;
}

function swAssemblyDriveDownloadUrlForFile_(file, size) {
  swEnsureAssemblyFileCanBeShared_(file);
  var sharingErrors = [];
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    sharingErrors.push('DriveApp: ' + (err && err.message ? err.message : err));
    try {
      swCreateDriveAnyoneWithLinkPermission_(file.getId());
    } catch (apiErr) {
      sharingErrors.push('Drive API: ' + (apiErr && apiErr.message ? apiErr.message : apiErr));
      var message = 'Drive link sharing was denied for this ' + swFormatBytes_(size) + ' audio file. The workflow only tried to share the audio file itself, not the client folder. Apps Script can send recordings up to about 49 MB directly to AssemblyAI; larger files need anyone-with-link sharing allowed for the audio file or need to be compressed below 49 MB. ' + sharingErrors.join(' | ');
      var e = new Error(message);
      e.terminal = true;
      throw e;
    }
  }
  return 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(file.getId());
}

function swEnsureAssemblyFileCanBeShared_(file) {
  var mime = swNorm_(file.getMimeType ? file.getMimeType() : '');
  var name = swNorm_(file.getName ? file.getName() : '');
  var audioVideo = mime.indexOf('audio/') === 0 || mime.indexOf('video/') === 0 ||
    /\.(m4a|mp3|mp4|mov|wav|aac|webm|mpeg|mpga)$/i.test(name);
  if (audioVideo) return;
  var err = new Error('Only audio/video appointment files can be shared externally for AssemblyAI transcription.');
  err.terminal = true;
  throw err;
}

function swCreateDriveAnyoneWithLinkPermission_(fileId) {
  var url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) +
    '/permissions?supportsAllDrives=true&sendNotificationEmail=false&fields=id,type,role,allowFileDiscovery';
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    payload: JSON.stringify({
      role: 'reader',
      type: 'anyone',
      allowFileDiscovery: false
    }),
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  if (code >= 200 && code < 300) return true;
  throw new Error('HTTP ' + code + ' ' + response.getContentText().slice(0, 500));
}

function swBytesNumber_(value) {
  var n = Number(String(value || '').replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

function swFormatBytes_(bytes) {
  bytes = Number(bytes) || 0;
  if (bytes >= 1000000000) return (bytes / 1000000000).toFixed(1) + ' GB';
  if (bytes >= 1000000) return (bytes / 1000000).toFixed(1) + ' MB';
  if (bytes >= 1000) return (bytes / 1000).toFixed(1) + ' KB';
  return bytes + ' B';
}

function swAssemblyUploadDriveFile_(file) {
  var response = UrlFetchApp.fetch(swAssemblyBaseUrl_() + '/v2/upload', {
    method: 'post',
    headers: {
      Authorization: swAssemblyApiKey_(),
      'Content-Type': 'application/octet-stream'
    },
    payload: file.getBlob().getBytes(),
    muteHttpExceptions: true
  });
  return swFetchJsonOrThrow_(response, 'AssemblyAI upload');
}

function swAssemblySubmitTranscript_(uploadUrl) {
  if (!uploadUrl) throw new Error('AssemblyAI upload did not return upload_url.');
  var body = {
    audio_url: uploadUrl,
    speech_models: ['universal-3-pro', 'universal-2'],
    language_detection: true,
    speaker_labels: true,
    format_text: true,
    punctuate: true
  };
  var response = UrlFetchApp.fetch(swAssemblyBaseUrl_() + '/v2/transcript', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: swAssemblyApiKey_() },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  return swFetchJsonOrThrow_(response, 'AssemblyAI transcript submit');
}

function swAssemblyGetTranscript_(transcriptId) {
  var response = UrlFetchApp.fetch(swAssemblyBaseUrl_() + '/v2/transcript/' + encodeURIComponent(transcriptId), {
    method: 'get',
    headers: { Authorization: swAssemblyApiKey_() },
    muteHttpExceptions: true
  });
  return swFetchJsonOrThrow_(response, 'AssemblyAI transcript poll');
}

function swAssemblyTranscriptText_(transcript) {
  if (transcript && transcript.utterances && transcript.utterances.length) {
    return transcript.utterances.map(function (u) {
      return 'Speaker ' + (u.speaker || '') + ': ' + swTrim_(u.text || '');
    }).join('\n\n');
  }
  return swTrim_(transcript && transcript.text || '');
}

function swFetchJsonOrThrow_(response, label) {
  var code = response.getResponseCode();
  var text = response.getContentText() || '';
  var body = {};
  try { body = JSON.parse(text); } catch (_) {}
  if (code >= 200 && code < 300) return body;
  var err = new Error(label + ' failed (' + code + '): ' + (body.error || body.message || text).toString().slice(0, 500));
  err.terminal = code >= 400 && code < 500 && code !== 429;
  throw err;
}

function swOpenAIAppointmentFollowUpDraft_(transcriptText, artifact, appointment) {
  var key = swTrim_(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'));
  if (!key) {
    var err = new Error('Missing OPENAI_API_KEY in Script Properties.');
    err.terminal = true;
    throw err;
  }
  var model = swTrim_(PropertiesService.getScriptProperties().getProperty('OPENAI_APPOINTMENT_SUMMARY_MODEL')) ||
    'gpt-5.4-mini';
  var clientFollowUpPrompt = "You are a warm, caring jewelry consultant writing a follow-up text to a client after their engagement ring consultation. This is an emotional, exciting moment in their life — write like someone who genuinely shared in that excitement with them.\n\nUsing the transcript, write a text message that:\n- Opens with their name and a warm callback to a specific fun or meaningful moment from the visit\n- Recaps what resonated with them (styles, stones, details) in a way that shows you were truly listening\n- States next steps clearly but naturally — woven in, not listed\n- Leaves the door open for any changes with zero pressure\n- Feels like a text from a trusted friend who happens to be an expert, not a sales rep\n\nUnder 180 words. No bullet points in the output. No generic openers like \"It was so great meeting you.\" Make every sentence earn its place. Return only the message.";
  var salesBriefPrompt = 'Write a compact internal Sales Brief for the Client Advisor and JOC. Focus only on sales-useful details supported by the transcript: the emotional callback, what resonated, specific style/stone/design preferences, stated timing or budget signals, concerns or open questions, and what the team should remember before the next touch. Keep it concise, practical, and non-repetitive. Do not write a transcript summary.';
  var reviewFlagsPrompt = 'Return Review Flags only when a human should double-check something before sending or acting: unclear next step, possible contradiction, missing consultant/client name, unsupported inference, sensitive wording, or low-certainty detail. Use short actionable phrases. Return an empty array when there are no real flags.';
  var body = {
    model: model,
    input: [
      {
        role: 'system',
        content: 'Generate one structured post-appointment AI review package for a jewelry sales workflow. Use only transcript-supported facts. Do not include markdown, labels, or bullets in clientFollowUpDraft.'
      },
      {
        role: 'user',
        content: [
          'Return JSON with exactly these fields: clientFollowUpDraft, salesBrief, reviewFlags.',
          '',
          'clientFollowUpDraft prompt:',
          clientFollowUpPrompt,
          '',
          'salesBrief prompt:',
          salesBriefPrompt,
          '',
          'reviewFlags prompt:',
          reviewFlagsPrompt,
          '',
          'Appointment context:',
          JSON.stringify({
            rootApptId: artifact['RootApptID'] || '',
            customerName: appointment.name || '',
            brand: appointment.brand || '',
            visitDate: appointment.visitDate || '',
            visitTime: appointment.visitTime || '',
            visitType: appointment.visitType || '',
            clientAdvisor: appointment.assignedRep || '',
            joc: appointment.assistedRep || '',
            existingNextSteps: appointment.nextSteps || '',
            designRequest: appointment.designRequest || '',
            diamondRequirements: appointment.dvCustomerLookingFor || appointment.dvCustomerRequirementsJson || ''
          }, null, 2),
          '',
          'Consultation transcript:',
          transcriptText
        ].join('\n')
      }
    ],
    max_output_tokens: 1200,
    text: {
      format: {
        type: 'json_schema',
        name: 'appointment_follow_up_review',
        strict: true,
        schema: swAppointmentSummarySchema_()
      }
    }
  };
  var response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  var parsed = swFetchJsonOrThrow_(response, 'OpenAI client follow-up draft');
  if (parsed.usage) Logger.log('SW_OPENAI_APPOINTMENT_FOLLOW_UP_USAGE ' + JSON.stringify(parsed.usage));
  if (parsed.status === 'incomplete') throw new Error('OpenAI response incomplete: ' + swStringify_(parsed.incomplete_details || {}));
  var json = swExtractOpenAIJson_(parsed);
  if (!json) throw new Error('OpenAI did not return valid follow-up review JSON.');
  return json;
}

function swAppointmentSummarySchema_() {
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      clientFollowUpDraft: { type: 'string' },
      salesBrief: { type: 'string' },
      reviewFlags: { type: 'array', items: { type: 'string' } }
    },
    required: ['clientFollowUpDraft', 'salesBrief', 'reviewFlags']
  };
}

function swExtractOpenAIJson_(body) {
  if (typeof __extractJsonFromResponsesBody__ === 'function') {
    var extracted = __extractJsonFromResponsesBody__(body);
    if (extracted) return extracted;
  }
  if (body && typeof body.output_text === 'string') {
    try { return JSON.parse(body.output_text); } catch (_) {}
  }
  try {
    var output = body.output || [];
    for (var i = 0; i < output.length; i++) {
      var content = output[i].content || [];
      for (var j = 0; j < content.length; j++) {
        if (content[j].json) return content[j].json;
        if (content[j].text) {
          try { return JSON.parse(content[j].text); } catch (_) {}
        }
      }
    }
  } catch (_) {}
  return null;
}

function swNormalizeAppointmentSummary_(result) {
  result = result || {};
  var reviewFlags = result.reviewFlags || [];
  if (!Array.isArray(reviewFlags)) reviewFlags = reviewFlags ? [String(reviewFlags)] : [];
  return {
    clientFollowUpDraft: swTrim_(result.clientFollowUpDraft),
    salesBrief: swTrim_(result.salesBrief),
    reviewFlags: reviewFlags.map(swTrim_).filter(Boolean)
  };
}

function swReadableAppointmentSummary_(summary) {
  var lines = [];
  if (summary.clientFollowUpDraft) {
    lines.push('Client-Facing Follow-Up Draft', summary.clientFollowUpDraft);
  }
  if (summary.salesBrief) {
    lines.push('', 'Sales Brief', summary.salesBrief);
  }
  if ((summary.reviewFlags || []).length) {
    lines.push('', 'Review Flags', (summary.reviewFlags || []).join('\n'));
  }
  return lines.join('\n').replace(/^\n+/, '');
}

function swCreateGoogleDocInFolder_(folder, title, heading, metadataLines, bodyText) {
  var doc = DocumentApp.create(title);
  var body = doc.getBody();
  body.clear();
  body.appendParagraph(heading || title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  (metadataLines || []).filter(Boolean).forEach(function (line) {
    body.appendParagraph(line).setHeading(DocumentApp.ParagraphHeading.NORMAL);
  });
  body.appendParagraph('');
  String(bodyText || '').split('\n').forEach(function (line) {
    body.appendParagraph(line);
  });
  doc.saveAndClose();
  var file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch (_) {}
  return { id: doc.getId(), url: doc.getUrl() };
}

function swReadGoogleDocText_(docId) {
  if (!docId) return '';
  return DocumentApp.openById(docId).getBody().getText();
}

function swHandleArtifactWorkerError_(ss, row, err) {
  var now = new Date();
  var attempts = Number(row['Attempts'] || 0) + 1;
  var terminal = !!(err && err.terminal) || attempts >= 3;
  swPatchAppointmentArtifactRow_(ss, row, {
    'Workflow Stage': terminal ? SW_ARTIFACT_STAGES.ERROR : row['Workflow Stage'],
    'Attempts': attempts,
    'Next Poll At': terminal ? '' : swIso_(swDateAddHours_(now, Math.min(30, Math.pow(2, attempts) * 5) / 60)),
    'Last Error': swTrim_(err && err.message || err).slice(0, 1000),
    'Updated At': swIso_(now)
  });
}

function sw_installAppointmentAutomationTriggers() {
  if (typeof sw_installBackgroundOrchestratorTrigger === 'function') {
    var result = sw_installBackgroundOrchestratorTrigger();
    result.message = 'Installed 5-minute background orchestrator and removed retired appointment background workers.';
    return result;
  }
  return { ok: false, error: 'sw_installBackgroundOrchestratorTrigger unavailable' };
}

function sw_setAppointmentAutomationScriptProperties(assemblyAiApiKey, openAiApiKey, options) {
  options = options || {};
  var props = {};
  var assemblyKey = swTrim_(assemblyAiApiKey);
  var openAiKey = swTrim_(openAiApiKey);
  if (!assemblyKey) throw new Error('ASSEMBLYAI_API_KEY is required.');
  if (!openAiKey) throw new Error('OPENAI_API_KEY is required.');

  props.ASSEMBLYAI_API_KEY = assemblyKey;
  props.OPENAI_API_KEY = openAiKey;
  props.ASSEMBLYAI_BASE_URL = swTrim_(options.assemblyAiBaseUrl) || 'https://api.assemblyai.com';

  var model = swTrim_(options.openAiAppointmentSummaryModel || options.openAiModel || '');
  if (model) props.OPENAI_APPOINTMENT_SUMMARY_MODEL = model;

  PropertiesService.getScriptProperties().setProperties(props, false);
  return {
    ok: true,
    updated: Object.keys(props),
    message: 'Appointment automation Script Properties updated.'
  };
}

function sw_promptSetAppointmentAutomationScriptProperties() {
  var ui = SpreadsheetApp.getUi();
  var assemblyPrompt = ui.prompt(
    'Set AssemblyAI API Key',
    'Paste the AssemblyAI API key. It will be stored in Script Properties, not in the spreadsheet.',
    ui.ButtonSet.OK_CANCEL
  );
  if (assemblyPrompt.getSelectedButton() !== ui.Button.OK) {
    return { ok: false, cancelled: true, message: 'Cancelled before setting AssemblyAI API key.' };
  }

  var openAiPrompt = ui.prompt(
    'Set OpenAI API Key',
    'Paste the OpenAI API key. It will be stored in Script Properties, not in the spreadsheet.',
    ui.ButtonSet.OK_CANCEL
  );
  if (openAiPrompt.getSelectedButton() !== ui.Button.OK) {
    return { ok: false, cancelled: true, message: 'Cancelled before setting OpenAI API key.' };
  }

  var result = sw_setAppointmentAutomationScriptProperties(
    assemblyPrompt.getResponseText(),
    openAiPrompt.getResponseText(),
    { assemblyAiBaseUrl: 'https://api.assemblyai.com' }
  );
  ui.alert('Appointment automation properties updated.');
  return result;
}

function sw_checkAppointmentAutomationScriptProperties() {
  var props = PropertiesService.getScriptProperties();
  return {
    ok: true,
    assemblyAiApiKeySet: !!swTrim_(props.getProperty('ASSEMBLYAI_API_KEY')),
    openAiApiKeySet: !!swTrim_(props.getProperty('OPENAI_API_KEY')),
    assemblyAiBaseUrl: swTrim_(props.getProperty('ASSEMBLYAI_BASE_URL')) || 'https://api.assemblyai.com',
    openAiAppointmentSummaryModelSet: !!swTrim_(props.getProperty('OPENAI_APPOINTMENT_SUMMARY_MODEL'))
  };
}

function sw_showAppointmentAutomationScriptPropertiesStatus() {
  var status = sw_checkAppointmentAutomationScriptProperties();
  SpreadsheetApp.getUi().alert(
    'Appointment Automation API Keys',
    [
      'AssemblyAI API key: ' + (status.assemblyAiApiKeySet ? 'set' : 'missing'),
      'OpenAI API key: ' + (status.openAiApiKeySet ? 'set' : 'missing'),
      'AssemblyAI base URL: ' + status.assemblyAiBaseUrl,
      'OpenAI summary model override: ' + (status.openAiAppointmentSummaryModelSet ? 'set' : 'not set')
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return status;
}

/**
 * Local-only helper. Paste keys into this function in the Apps Script editor,
 * run it once, then remove the pasted keys before saving/pushing code.
 */
function sw_setAppointmentAutomationScriptPropertiesLocalExample() {
  return sw_setAppointmentAutomationScriptProperties(
    'PASTE_ASSEMBLYAI_API_KEY_HERE',
    'PASTE_OPENAI_API_KEY_HERE',
    {
      assemblyAiBaseUrl: 'https://api.assemblyai.com'
    }
  );
}
