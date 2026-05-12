function zz_generateTasksForAppointmentOnce(options) {
  options = options || {};
  var wantedAppt = String(options.apptId || '').trim();
  var wantedUid = String(options.uid || '').trim();
  if (!wantedAppt && !wantedUid) throw new Error('Pass apptId or uid.');

  var lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(Number(options.lockWaitMs || 10000))) {
    return { ok: false, skipped: true, reason: 'LOCK_BUSY' };
  }

  try {
    sw_setupSalesWorkflow();
    var ss = swSpreadsheet_();
    var ctx = swBuildContext_(ss, true);
    ctx.appointmentSummaryByRoot = typeof swAppointmentSummaryIndex_ === 'function' ? swAppointmentSummaryIndex_(ss) : {};
    var now = new Date();
    ctx.now = now;
    var masterRows = swReadAppointments_(ss);
    swPrepareClientAdvisorRoundRobin_(ss, ctx, masterRows);
    var rec = null;
    masterRows.some(function (row) {
      var appt = String(row.appt || '').trim();
      var uid = String(row.uid || '').trim();
      if ((wantedAppt && appt === wantedAppt) || (wantedUid && uid === wantedUid)) {
        rec = row;
        return true;
      }
      return false;
    });
    if (!rec) return { ok: false, reason: 'APPOINTMENT_NOT_FOUND', apptId: wantedAppt, uid: wantedUid };

    var taskState = swReadTaskState_(ss);
    swBeginDeferredTaskWrites_(ss, taskState);
    var summary = {
      ok: true,
      generatedAt: swIso_(now),
      targeted: true,
      scannedAppointments: 1,
      created: 0,
      updated: 0,
      blocked: 0,
      skippedOld: 0,
      systemCompleted: 0
    };
    try {
      if (!swIsWorkflowRelevant_(rec, now, ctx)) {
        summary.skippedOld++;
      } else if (!swIsAppointmentActive_(rec)) {
        summary.blocked += swBlockTasksForAppointment_(ss, taskState, rec, SW_INACTIVE_APPOINTMENT_BLOCK_REASON);
      } else {
        swMaybeAutoAssignClientAdvisor_(ss, ctx, rec, summary);
        swGenerateTasksForAppointment_(ss, taskState, ctx, rec, now, summary);
      }
    } finally {
      swFlushDeferredTaskWrites_(ss, taskState);
    }

    if ((summary.autoAssignedClientAdvisors || 0) ||
        (summary.autoReassignedClientAdvisors || 0) ||
        (summary.autoLinkedJocFromAdvisor || 0)) {
      try {
        if (typeof swInvalidateAppointmentReadModelsAfterWrite_ === 'function') {
          swInvalidateAppointmentReadModelsAfterWrite_(ss, 'Targeted appointment owner/task generation');
        }
      } catch (_) {}
    }

    var master = ss.getSheetByName(SW_SHEETS.MASTER);
    var headers = master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0];
    var H = swHeaderMapFromArray_(headers);
    var rowNumber = Number(rec.row || 0);
    var rowValues = rowNumber ? master.getRange(rowNumber, 1, 1, master.getLastColumn()).getDisplayValues()[0] : [];
    var pick = function (names) {
      var idx = swPickIndex_(H, names);
      return idx >= 0 ? rowValues[idx] || '' : '';
    };
    summary.master = {
      row: rowNumber,
      apptId: pick(['APPT_ID']),
      uid: pick(['CalendlyEventUID']),
      assignedRep: pick(['Assigned Rep', 'Client Advisor']),
      assignedRepEmail: pick(['Assigned Rep Email', 'Client Advisor Email']),
      assistedRep: pick(['Assisted Rep', 'JOC']),
      assistedRepEmail: pick(['Assisted Rep Email', 'JOC Email'])
    };
    return summary;
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}
