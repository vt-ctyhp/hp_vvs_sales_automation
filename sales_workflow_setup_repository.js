/**
 * Sales workflow setup repository: seed data, migrations, and default workflow templates.
 */

function swSeedConfig_(sh) {
  swMigrateConfigRows_(sh);
  var rows = [
    ['SYSTEM', 'FEATURE_ENABLED', 'Y', '', '', '', 'Y', '', 'Set N to pause workflow generation.'],
    ['SYSTEM', 'ADMIN_EMAILS', '', '', '', '', 'Y', '', 'Comma-separated manager/admin emails. Blank means all users can administer during setup.'],
    ['SYSTEM', 'MAP_LINK_VVS', '', '', '', '', 'Y', '', 'Map or instructions link for VVS appointments.'],
    ['SYSTEM', 'MAP_LINK_HUNG_PHAT', '', '', '', '', 'Y', '', 'Map or instructions link for Hung Phat / HPUSA appointments.'],
    ['SYSTEM', 'LOCATION_MSG_VVS', '', '', '', '', 'Y', '', 'Store/location message for VVS map/instructions templates.'],
    ['SYSTEM', 'LOCATION_MSG_HUNG_PHAT', '', '', '', '', 'Y', '', 'Store/location message for Hung Phat / HPUSA map/instructions templates.'],
    ['SYSTEM', 'WELCOME_MSG_VVS', '', '', '', '', 'Y', '', 'Welcome to Your Ring Journey text message for VVS appointments.'],
    ['SYSTEM', 'WELCOME_MSG_HUNG_PHAT', '', '', '', '', 'Y', '', 'Welcome to Your Ring Journey text message for Hung Phat / HPUSA appointments.'],
    ['SYSTEM', 'HYBRID_MSG_VVS', '{{welcomeMessage}}\n\n{{locationMsg}}\n\n{{welcomeImageUrl}}', '', '', '', 'Y', '', 'Hybrid welcome + instructions message for VVS appointments. Supports {{welcomeImageUrl}}, {{mapLink}}, {{welcomeMessage}}, and {{locationMsg}}.'],
    ['SYSTEM', 'HYBRID_MSG_HUNG_PHAT', '{{welcomeMessage}}\n\n{{locationMsg}}\n\n{{welcomeImageUrl}}', '', '', '', 'Y', '', 'Hybrid welcome + instructions message for Hung Phat / HPUSA appointments. Supports {{welcomeImageUrl}}, {{mapLink}}, {{welcomeMessage}}, and {{locationMsg}}.'],
    ['SYSTEM', 'WELCOME_IMAGE_VVS', '', '', '', '', 'Y', '', 'Welcome to Your Ring Journey image URL for VVS appointments.'],
    ['SYSTEM', 'WELCOME_IMAGE_HUNG_PHAT', '', '', '', '', 'Y', '', 'Welcome to Your Ring Journey image URL for Hung Phat / HPUSA appointments.'],
    ['SYSTEM', 'WORKFLOW_LOOKBACK_DAYS', '14', '', '', '', 'Y', '', 'Do not generate new tasks for appointments older than this.'],
    ['SYSTEM', 'WORKFLOW_FUTURE_DAYS', '365', '', '', '', 'Y', '', 'Generate upcoming appointment workflow tasks through this many days out.'],
    ['SYSTEM', 'JOC_OWNER_SOURCE', '00_Master Appointments: Assisted Rep', '', '', '', 'Y', '', 'JOC ownership comes from each appointment row, not primary/backup config.'],
    ['SYSTEM', 'DIAMOND_TRACKING_SOURCE', '200_', '', '', '', 'Y', '', 'Diamond tracking ETA/status source of truth. Sales Workflow payloads only cache snapshots for cards.'],
    ['SYSTEM', 'DIAMOND_RETURN_WINDOW_DAYS', '30', '', '', '', 'Y', '', 'Diamonds marked return are due back this many days after Purchased / Ordered Date.'],
    ['SYSTEM', 'DIAMOND_RETURN_WARNING_DAYS', '7', '', '', '', 'Y', '', 'Show return tasks this many days before the 30-day return deadline.'],
    ['SYSTEM', 'SHARED_DIAMOND_ORDER_ADMIN_QUEUE', 'Diamond Order Admin', '', '', '', 'Y', '', 'Display label for DIAMOND_ORDER_ADMIN role-owned tasks.'],
    ['SYSTEM', 'SHARED_DIAMOND_ORDER_ASSISTANT_QUEUE', 'Diamond Order Assistant', '', '', '', 'Y', '', 'Display label for DIAMOND_ORDER_ASSISTANT role-owned tasks.'],
    ['SYSTEM', 'CLIENT_ADVISOR_ROUND_ROBIN', 'N', '', '', '', 'Y', '', 'Set Y to assign or reassign Client Advisors by schedule, skills, and round robin during queue refresh.'],
    ['SYSTEM', 'DATA_CLEANUP_ENABLED', 'Y', '', '', '', 'Y', '', 'Set N to pause stale customer data cleanup task generation.'],
    ['SYSTEM', 'DATA_CLEANUP_STALE_DAYS', '30', '', '', '', 'Y', '', 'Open Lead / Hot Lead / Follow-Up customers become cleanup candidates after this many days without a meaningful touch.'],
    ['SYSTEM', 'DATA_CLEANUP_CAMPAIGN_ID', 'ONE_TIME_2026_05', '', '', '', 'Y', '', 'Identifier for the one-time cleanup campaign.'],
    ['SYSTEM', 'DATA_CLEANUP_CAMPAIGN_TAB_ENABLED', 'Y', '', '', '', 'Y', '', 'Show the temporary Cleanup dashboard tab for unresolved one-time campaign cases. The generator disables this after the campaign is resolved.'],
    ['SYSTEM', 'READ_MODEL_ENABLED', 'N', '', '', '', 'Y', '', 'Shadow read-model flag. Phase 1 builds generated _SW_* tabs but does not serve app reads from them.'],
    ['SYSTEM', 'READ_MODEL_SERVE_TASKS', 'Y', '', '', '', 'Y', '', 'Use fresh _SW_TaskReadModel rows for dashboard bootstrap and queue views. Falls back to _SalesTaskQueue when stale or missing.'],
    ['SYSTEM', 'READ_MODEL_SERVE_CUSTOMERS', 'Y', '', '', '', 'Y', '', 'Use fresh _SW_CustomerReadModel and packed customer caches for customer search. Falls back to 00_Master Appointments when stale or missing.'],
    ['SYSTEM', 'READ_MODEL_SERVE_DIAMONDS', 'Y', '', '', '', 'Y', '', 'Use fresh _SW_DiamondReadModel rows for diamond dashboard APIs. Falls back to 200_ when stale or missing.'],
    ['SYSTEM', 'READ_MODEL_SERVE_APPOINTMENTS', 'Y', '', '', '', 'Y', '', 'Use fresh _SW_AppointmentReadModel and _SW_CalendarMonthReadModel rows for appointment/calendar reads. Falls back to 00_Master Appointments when stale or missing.'],
    ['SYSTEM', 'READ_MODEL_SERVE_PAYMENTS', 'Y', '', '', '', 'Y', '', 'Use fresh _SW_PaymentReadModel rows for dashboard payment reads. Falls back to the payment ledger when stale or missing.'],
    ['SYSTEM', 'READ_MODEL_SERVE_ADMIN', 'Y', '', '', '', 'Y', '', 'Use precomputed _SW_AdminDashboardReadModel payloads for standard unfiltered admin dashboard windows. Falls back to live aggregation when stale or filtered.'],
    ['SYSTEM', 'READ_MODEL_TTL_SECONDS', '600', '', '', '', 'Y', '', 'Freshness window for generated workflow read models. The refresh trigger still targets every 5 minutes.'],
    ['USER', 'ADMIN_1', '', 'Admin', '', '', 'Y', '1', 'Optional admin row.'],
    ['USER', 'DIAMOND_ORDER_ADMIN_1', '', 'DIAMOND_ORDER_ADMIN', '', '', 'Y', '1', 'Legacy config row only; dashboard access now comes from _SalesWorkflowUsers role DIAMOND_ORDER_ADMIN.'],
    ['USER', 'DIAMOND_ORDER_ASSISTANT_1', '', 'DIAMOND_ORDER_ASSISTANT', '', '', 'Y', '1', 'Legacy config row only; dashboard access now comes from _SalesWorkflowUsers role DIAMOND_ORDER_ASSISTANT.'],
    ['SYSTEM', 'SHARED_JOC_QUEUE', 'JOC Coverage', '', '', '', 'Y', '', 'Used when no scheduled JOC is available.']
  ];
  swAppendMissingConfigRows_(sh, rows);
}

function swMigrateConfigRows_(sh) {
  if (sh.getLastRow() < 2) return;
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, SW_CONFIG_HEADERS.length).getDisplayValues();
  var byKey = {};
  values.forEach(function (row, i) {
    byKey[swNorm_(row[0]) + '|' + swNorm_(row[1])] = i + 2;
  });

  var renames = {
    locationlabelvvs: 'LOCATION_MSG_VVS',
    locationlabelhungphat: 'LOCATION_MSG_HUNG_PHAT'
  };

  for (var r = values.length - 1; r >= 0; r--) {
    var rowIndex = r + 2;
    var section = swNorm_(values[r][0]);
    var key = swHeaderKey_(values[r][1]);
    if (section !== 'system') continue;
    if (key === 'maplink' || key === 'locationlabel') {
      sh.deleteRow(rowIndex);
      continue;
    }
    if (renames[key]) {
      var targetKey = swNorm_('SYSTEM') + '|' + swNorm_(renames[key]);
      var targetRow = byKey[targetKey];
      if (targetRow && targetRow !== rowIndex) {
        if (!swTrim_(sh.getRange(targetRow, 3).getDisplayValue()) && swTrim_(values[r][2])) {
          sh.getRange(targetRow, 3).setValue(values[r][2]);
        }
        sh.deleteRow(rowIndex);
      } else {
        sh.getRange(rowIndex, 2).setValue(renames[key]);
      }
    }
  }
}

function swAppendMissingConfigRows_(sh, rows) {
  var existing = {};
  if (sh.getLastRow() > 1) {
    var values = sh.getRange(2, 1, sh.getLastRow() - 1, SW_CONFIG_HEADERS.length).getDisplayValues();
    values.forEach(function (row) {
      var key = swNorm_(row[0]) + '|' + swNorm_(row[1]);
      if (key !== '|') existing[key] = true;
    });
  }
  var append = rows.filter(function (row) {
    return !existing[swNorm_(row[0]) + '|' + swNorm_(row[1])];
  });
  if (append.length) {
    sh.getRange(sh.getLastRow() + 1, 1, append.length, SW_CONFIG_HEADERS.length).setValues(append);
  }
}

function swSeedTemplates_(sh) {
  swMigrateTemplateRows_(sh);
  var rows = swDefaultTemplates_();
  swAppendMissingTemplateRows_(sh, rows);
  swUpdateManagedDiamondTemplateRows_(sh, rows);
}

function swAppendMissingTemplateRows_(sh, rows) {
  var existing = {};
  if (sh.getLastRow() > 1) {
    var values = sh.getRange(2, 1, sh.getLastRow() - 1, SW_TEMPLATE_HEADERS.length).getDisplayValues();
    values.forEach(function (row) {
      var taskType = swTrim_(row[0]);
      if (taskType) existing[taskType] = true;
    });
  }
  var append = rows.filter(function (row) {
    return !existing[swTrim_(row[0])];
  });
  if (append.length) {
    sh.getRange(sh.getLastRow() + 1, 1, append.length, SW_TEMPLATE_HEADERS.length).setValues(append);
  }
}

function swMigrateTemplateRows_(sh) {
  if (sh.getLastRow() < 2) return;
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, SW_TEMPLATE_HEADERS.length).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rowIndex = i + 2;
    var taskType = swTrim_(row[0]);
    if (taskType === SW_TASKS.WELCOME) {
      var welcomeTemplate = String(row[3] || '');
      if (swShouldUseDefaultWelcomeTemplate_(welcomeTemplate)) {
        sh.getRange(rowIndex, 4).setValue('{{welcomeMessage}}');
      }
      if (String(row[5] || '').indexOf('welcomeImageUrl') < 0) {
        sh.getRange(rowIndex, 5).setValue('Welcome Journey Image');
        sh.getRange(rowIndex, 6).setValue('{{welcomeImageUrl}}');
      }
    }
    if (taskType === SW_TASKS.MAP) {
      if (swShouldUseDefaultMapTemplate_(String(row[3] || ''))) {
        sh.getRange(rowIndex, 4).setValue('{{locationMsg}}');
      }
      if (String(row[5] || '').indexOf('mapLink') < 0) {
        sh.getRange(rowIndex, 5).setValue('Map / Instructions');
        sh.getRange(rowIndex, 6).setValue('{{mapLink}}');
      }
    }
    if (taskType === SW_TASKS.HYBRID) {
      if (swShouldUseDefaultHybridTemplate_(String(row[3] || ''))) {
        sh.getRange(rowIndex, 4).setValue('{{hybridMessage}}');
      }
      if (String(row[5] || '').indexOf('mapLink') < 0) {
        sh.getRange(rowIndex, 5).setValue('Map / Instructions');
        sh.getRange(rowIndex, 6).setValue('{{mapLink}}');
      }
    }
    if (taskType === SW_TASKS.CHECKLIST) {
      var checklistText = String(row[6] || '');
      if (checklistText.indexOf('appointment_context_confirmed') < 0 || checklistText.indexOf('uploaded_recap') >= 0 || checklistText.indexOf('recorded_appointment') >= 0) {
        sh.getRange(rowIndex, 2).setValue('Complete Appointment');
        sh.getRange(rowIndex, 3).setValue('Select the appointment outcome, upload the appointment materials, and confirm the handoff context. Completed appointments require an appointment recording before this task can close.');
        sh.getRange(rowIndex, 7).setValue('[{"id":"appointment_context_confirmed","label":"Appointment notes, next steps, and handoff context are captured","required":true},{"id":"client_materials_collected","label":"Relevant intake photos, recap notes, or viewing materials are ready for upload","required":false},{"id":"physical_handoff_complete","label":"Physical appointment wrap-up items are handled","required":false}]');
      }
    }
    if (taskType === SW_TASKS.PROCESS) {
      sh.getRange(rowIndex, 2).setValue('Legacy Process Appointment Data');
      sh.getRange(rowIndex, 3).setValue('Retired for new appointments. New completed appointments use the AssemblyAI/OpenAI appointment artifact workflow.');
      sh.getRange(rowIndex, 8).setValue('Legacy Submit');
    }
    if (taskType === SW_TASKS.APPROVE) {
      sh.getRange(rowIndex, 2).setValue('Review Client Follow-Up Draft');
      sh.getRange(rowIndex, 3).setValue('Review the Sales Brief, edit the AI-generated client-facing follow-up message, then approve it for JOC.');
      sh.getRange(rowIndex, 4).setValue('{{clientFollowUpDraft}}');
      sh.getRange(rowIndex, 8).setValue('Approve for JOC');
    }
    if (taskType === SW_TASKS.FINAL) {
      sh.getRange(rowIndex, 2).setValue('Send Approved Follow-Up');
      sh.getRange(rowIndex, 3).setValue('Send the approved client-facing follow-up message. Use the concise handoff context only if needed; the transcript is not required for final sending.');
      sh.getRange(rowIndex, 4).setValue('{{approvedText}}');
      sh.getRange(rowIndex, 8).setValue('Mark Sent');
    }
  }
}

function swUpdateManagedDiamondTemplateRows_(sh, defaults) {
  if (sh.getLastRow() < 2) return;
  var managed = {};
  [
    SW_TASKS.DIAMOND_PROPOSE,
    SW_TASKS.DIAMOND_QUOTE,
    SW_TASKS.DIAMOND_ORDER,
    SW_TASKS.DIAMOND_TRACK,
    SW_TASKS.DIAMOND_RETURN,
    SW_TASKS.DIAMOND_ORDER_ACK_REP,
    SW_TASKS.DIAMOND_ORDER_ACK_JOC
  ].forEach(function (taskType) { managed[taskType] = true; });

  var byType = {};
  (defaults || []).forEach(function (row) {
    if (managed[row[0]]) byType[row[0]] = row;
  });
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, SW_TEMPLATE_HEADERS.length).getDisplayValues();
  values.forEach(function (row, i) {
    var taskType = swTrim_(row[0]);
    var desired = byType[taskType];
    if (!desired) return;
    var rowIndex = i + 2;
    sh.getRange(rowIndex, 2).setValue(desired[1]);
    sh.getRange(rowIndex, 3).setValue(desired[2]);
    if (!swTrim_(row[3])) sh.getRange(rowIndex, 4).setValue(desired[3]);
    if (!swTrim_(row[4])) sh.getRange(rowIndex, 5).setValue(desired[4]);
    if (!swTrim_(row[5])) sh.getRange(rowIndex, 6).setValue(desired[5]);
    sh.getRange(rowIndex, 7).setValue(desired[6]);
    sh.getRange(rowIndex, 8).setValue(desired[7]);
  });
}

function swDefaultTemplates_() {
  return [
    [SW_TASKS.ASSIGN, 'Assign Appointment', 'System-owned assignment record. No manual action needed.', '', '', '', '', 'Assigned'],
    [SW_TASKS.WELCOME, 'Send Welcome to Your Ring Journey Text', 'Send the brand-specific welcome message and welcome image, then mark it sent.', '{{welcomeMessage}}', 'Welcome Journey Image', '{{welcomeImageUrl}}', '', 'Mark Sent'],
    [SW_TASKS.HYBRID, 'Send Hybrid Welcome + Instructions', 'Appointment is within 24 hours. Send the brand-specific hybrid welcome and instructions message.', '{{hybridMessage}}', 'Map / Instructions', '{{mapLink}}', '', 'Mark Sent'],
    [SW_TASKS.MAP, 'Send Map & Instructions', 'Send the map and appointment instructions.', '{{locationMsg}}', 'Map / Instructions', '{{mapLink}}', '', 'Mark Sent'],
    [SW_TASKS.REVIEW, 'Review Appointment Folder', 'Review the intake form, inspiration images, and customer folder before the appointment.', '', 'Client Folder', '{{clientFolder}}', '', 'Acknowledged & Reviewed'],
    [SW_TASKS.CHECKLIST, 'Complete Appointment', 'Select the appointment outcome, upload the appointment materials, and confirm the handoff context. Completed appointments require an appointment recording before this task can close.', '', '', '', '[{"id":"appointment_context_confirmed","label":"Appointment notes, next steps, and handoff context are captured","required":true},{"id":"client_materials_collected","label":"Relevant intake photos, recap notes, or viewing materials are ready for upload","required":false},{"id":"physical_handoff_complete","label":"Physical appointment wrap-up items are handled","required":false}]', 'Complete Appointment'],
    [SW_TASKS.PROCESS, 'Legacy Process Appointment Data', 'Retired for new appointments. New completed appointments use the AssemblyAI/OpenAI appointment artifact workflow.', '', 'Client Folder', '{{clientFolder}}', '', 'Legacy Submit'],
    [SW_TASKS.APPROVE, 'Review Client Follow-Up Draft', 'Review the Sales Brief, edit the AI-generated client-facing follow-up message, then approve it for JOC.', '{{clientFollowUpDraft}}', '', '', '', 'Approve for JOC'],
    [SW_TASKS.FINAL, 'Send Approved Follow-Up', 'Send the approved client-facing follow-up message. Use the concise handoff context only if needed; the transcript is not required for final sending.', '{{approvedText}}', '', '', '', 'Mark Sent'],
    [SW_TASKS.POST_CONSULT_STATUS, 'Post-Consult Client Status Update', 'JOC owns the first post-consult operational checkpoint. Update the client status, record next steps, and decide whether 3D or wax work is needed. If 3D is not needed, enter the reason in the task form.', 'Customer: {{customerName}}\nAppointment: {{appointmentDateTime}}\nClient Advisor: {{assignedRep}}\nCurrent SO: {{soNumber}}\n3D deadline: {{deadline3d}}\nWax status: {{waxStatus}}', 'Client Status Report', '{{reportUrl}}', '', 'Submit Client Status'],
    [SW_TASKS.START_3D, 'Start 3D Design', 'Start the 3D design from this dashboard task using the same Start 3D / Assign SO workflow. If 3D is not needed, mark No 3D Needed with a reason.', 'Customer: {{customerName}}\nBrand: {{brandRaw}}\nDesign request: {{designRequest}}\nNext steps: {{nextSteps}}\nClient folder: {{clientFolder}}', 'Client Folder', '{{clientFolder}}', '', 'Start 3D'],
    [SW_TASKS.RECORD_3D_DEADLINE, 'Record 3D Deadline', 'Record the 3D deadline the day after Start 3D. If the deadline cannot be obtained today, snooze this task with a reason so it is not counted late until the snooze date.', 'Customer: {{customerName}}\nSO: {{soNumber}}\n3D tracker: {{tracker3dUrl}}\nCurrent 3D deadline: {{deadline3d}}', '3D Tracker', '{{tracker3dUrl}}', '', 'Save 3D Deadline'],
    [SW_TASKS.REQUEST_WAX, 'Request Wax Print', 'Create the wax request from the dashboard so the Wax queue, Master mirror fields, and request folder stay aligned.', 'Customer: {{customerName}}\nSO/MO: {{soNumber}}\nWax status: {{waxStatus}}\nNext steps: {{nextSteps}}', '', '', '', 'Create Wax Request'],
    [SW_TASKS.UPDATE_WAX, 'Update Wax Request', 'Update open wax requests that are missing an admin deadline/status or are past their admin deadline.', 'Customer: {{customerName}}\nOpen wax requests: {{waxRequestSummary}}', 'Wax Request', '{{waxRequestUrl}}', '', 'Update Wax'],
    [SW_TASKS.DIAMOND_PROPOSE, 'Propose Diamonds for Viewing', 'For Diamond Viewing appointments, capture the structured customer requirements, then propose stones in the dashboard. Check the In-Stock Diamonds tab first if store inventory may work for the appointment date. Completing this task writes the customer requirements to Sheet 100, validates stones, inserts them into 200_, updates Sheet 100 diamond counts/status, and keeps the appointment context aligned.', 'Customer: {{customerName}}\nAppointment: {{appointmentDate}} {{appointmentTime}}\nDiamond status: {{diamondSummary}}\nLatest safe proposal target: {{diamondProposalTarget}}', '200_ Diamond Tracker', '{{diamondTrackerUrl}}', '[{"id":"reviewed_customer_needs","label":"Captured customer requirements, deciding factor, and variety strategy in this task","required":true},{"id":"checked_in_stock_options","label":"Checked in-stock diamonds and return dates where relevant","required":true},{"id":"entered_proposed_stones","label":"Entered proposed diamonds in this dashboard task","required":true},{"id":"confirmed_writeback","label":"Ready to write requirements to Sheet 100 and proposed diamonds to 200_","required":true}]', 'Submit Proposed Diamonds'],
    [SW_TASKS.DIAMOND_QUOTE, 'Prepare Diamond Viewing Quotation', 'Fill the quotation sheet and complete price research against the structured customer requirements from Sheet 100. Use the refresh buttons on the task card if 200_ diamond data or 3D tracker details changed after the quote was created.', 'Customer requirements:\n{{diamondCustomerRequirements}}\n\nQuotation: {{quotationUrl}}\n3D Tracker: {{tracker3dUrl}}\nDiamond tracker: {{diamondTrackerUrl}}\nDiamond status: {{diamondSummary}}', 'Quotation Sheet', '{{quotationUrl}}', '[{"id":"quote_link_checked","label":"Quotation sheet link opens correctly","required":true},{"id":"requirements_reviewed","label":"Reviewed customer requirements and variety strategy from Sheet 100","required":true},{"id":"diamonds_refreshed","label":"Diamond options refreshed from 200_ or confirmed current","required":true},{"id":"settings_refreshed","label":"3D setting details refreshed or confirmed current","required":true},{"id":"price_research_done","label":"Price research completed in quotation sheet","required":true}]', 'Quotation Ready'],
    [SW_TASKS.DIAMOND_ORDER, 'Order Diamonds', 'Review each proposed diamond against the structured customer requirements from Sheet 100. Select On the Way or Not Approved for every proposed stone; completion writes order status/date to 200_ and creates acknowledgement tasks for JOC and the client advisor.', 'Customer requirements:\n{{diamondCustomerRequirements}}\n\nPending proposed stones: {{diamondActionSummary}}\nTracker: {{diamondTrackerUrl}}', '200_ Diamond Tracker', '{{diamondTrackerUrl}}', '[{"id":"reviewed_customer_requirements","label":"Reviewed proposed diamonds against the Sheet 100 customer requirements","required":true},{"id":"selected_every_stone","label":"Selected On the Way or Not Approved for every proposed diamond","required":true},{"id":"confirmed_order_date","label":"Confirmed Purchased / Ordered Date is accurate","required":true},{"id":"ready_to_notify_team","label":"Ready for JOC and client advisor to acknowledge ordered diamonds","required":true}]', 'Confirm Order Updates'],
    [SW_TASKS.DIAMOND_TRACK, 'Track Diamond ETA', 'Check shipping/tracking for on-the-way diamonds. Enter ETA and status here; the 200_ tracker remains the source of truth. Late or concerning ETAs create rep/JOC alert tasks.', 'On-the-way stones: {{diamondActionSummary}}\nAppointment: {{appointmentDate}} {{appointmentTime}}', '200_ Diamond Tracker', '{{diamondTrackerUrl}}', '[{"id":"checked_tracking","label":"Checked diamond tracking/vendor ETA","required":true},{"id":"entered_eta","label":"Entered ETA/status on this task","required":true}]', 'Save ETA'],
    [SW_TASKS.DIAMOND_DELIVERY, 'Confirm Diamond Delivery', 'Confirm ordered diamonds were received and mark delivered/in stock. Return deadline must be based on Purchased / Ordered Date + 30 days.', 'Awaiting delivery: {{diamondActionSummary}}\nTracker: {{diamondTrackerUrl}}', '200_ Diamond Tracker', '{{diamondTrackerUrl}}', '[{"id":"confirmed_received","label":"Confirmed diamonds were received","required":true},{"id":"updated_delivery_status","label":"Updated delivery status in 200_","required":true},{"id":"return_deadline_checked","label":"Verified return due date is order date plus 30 days","required":true}]', 'Delivery Confirmed'],
    [SW_TASKS.DIAMOND_DECISIONS, 'Record Diamond Decisions', 'After viewing, mark each stone as Purchase or Return. Reconfirm diamond dimensions against the latest 3D tracker details and send manufacturing the confirmed dimensions using the generated message.', 'Decision context: {{diamondActionSummary}}\nManufacturing message:\n{{manufacturingMessage}}', '3D Tracker', '{{tracker3dUrl}}', '[{"id":"marked_purchase_return","label":"Marked Purchase/Return decisions in 200_","required":true},{"id":"confirmed_dimensions","label":"Confirmed diamond dimensions against latest 3D tracker","required":true},{"id":"manufacturing_messaged","label":"Sent manufacturing the confirmed dimensions message","required":true}]', 'Decisions Recorded'],
    [SW_TASKS.DIAMOND_RETURN, 'Return Diamonds', 'Review diamonds marked Return or not purchased. Diamonds must be returned within 30 days of Purchased / Ordered Date, not delivery date. Completing this task marks the listed 200_ rows as Return in Progress.', 'Return queue: {{diamondActionSummary}}\nTracker: {{diamondTrackerUrl}}', '200_ Diamond Tracker', '{{diamondTrackerUrl}}', '[{"id":"checked_due_dates","label":"Checked return due dates against Purchased / Ordered Date","required":true},{"id":"ready_to_return","label":"Ready to move listed stones into Return in Progress","required":true},{"id":"escalated_blockers","label":"Escalated any blocker before completing","required":true}]', 'Mark Return In Progress'],
    [SW_TASKS.DIAMOND_ORDER_ACK_REP, 'Acknowledge Diamonds Ordered', 'Client Advisor acknowledgement that diamonds were ordered. Review ordered stones against the customer requirements, ETA risk if present, and customer communication plan.', 'Customer requirements:\n{{diamondCustomerRequirements}}\n\nOrdered diamonds: {{diamondActionSummary}}\nAppointment: {{appointmentDate}} {{appointmentTime}}\nTracker: {{diamondTrackerUrl}}', '200_ Diamond Tracker', '{{diamondTrackerUrl}}', '[{"id":"reviewed_ordered_stones","label":"Reviewed which diamonds were ordered","required":true},{"id":"checked_customer_impact","label":"Checked customer requirements, appointment impact, and next step","required":true}]', 'Acknowledged'],
    [SW_TASKS.DIAMOND_ORDER_ACK_JOC, 'Acknowledge Diamonds Ordered for Quote', 'JOC acknowledgement that diamonds were ordered. Check the quotation plan against the customer requirements and update notes if ordered stones change quote assumptions.', 'Customer requirements:\n{{diamondCustomerRequirements}}\n\nOrdered diamonds: {{diamondActionSummary}}\nQuotation: {{quotationUrl}}\nTracker: {{diamondTrackerUrl}}', 'Quotation Sheet', '{{quotationUrl}}', '[{"id":"reviewed_ordered_stones","label":"Reviewed ordered diamonds against the quotation plan and requirements","required":true},{"id":"updated_quote_notes","label":"Updated quotation notes or confirmed no change needed","required":true}]', 'Acknowledged'],
    [SW_TASKS.DIAMOND_ETA_REP, 'Review Diamond ETA Risk', 'Diamond ETA/status needs Client Advisor review because it is late or concerning for the Diamond Viewing appointment.', 'ETA issue: {{diamondEtaIssue}}\nAppointment: {{appointmentDate}} {{appointmentTime}}\nTracker: {{diamondTrackerUrl}}', '200_ Diamond Tracker', '{{diamondTrackerUrl}}', '[{"id":"reviewed_eta_risk","label":"Reviewed ETA risk and customer impact","required":true},{"id":"coordinated_next_step","label":"Coordinated next step with JOC/order team","required":true}]', 'Risk Reviewed'],
    [SW_TASKS.DIAMOND_ETA_JOC, 'Review Diamond ETA for Quotation', 'Diamond ETA/status needs JOC review because it is late or concerning for the Diamond Viewing appointment.', 'ETA issue: {{diamondEtaIssue}}\nQuotation: {{quotationUrl}}\nTracker: {{diamondTrackerUrl}}', 'Quotation Sheet', '{{quotationUrl}}', '[{"id":"reviewed_eta_risk","label":"Reviewed ETA risk against quotation plan","required":true},{"id":"updated_quote_or_notes","label":"Updated quotation/notes if ETA changes options","required":true}]', 'Risk Reviewed'],
    [SW_TASKS.DATA_CLEANUP_REVIEW, 'Review Stale Customer Data', 'Review this stale customer profile, verify current ownership and operational status, then submit the proposed cleanup update. The record is not changed until the paired Client Advisor/JOC confirmation is complete.', 'Customer: {{customerName}}\nClient Advisor: {{assignedRep}}\nJOC: {{assistedRep}}\nCurrent stage: {{salesStage}}\nConversion: {{conversionStatus}}\nNext steps: {{nextSteps}}', 'Client Status Report', '{{reportUrl}}', '', 'Submit Cleanup Proposal'],
    [SW_TASKS.DATA_CLEANUP_CONFIRM, 'Confirm Customer Data Cleanup', 'Review the proposed cleanup update from the other owner. Confirm only when the customer record should be updated; return it if anything needs correction.', 'Customer: {{customerName}}\nProposed by: {{cleanupProposedBy}}\nProposed update: {{cleanupProposalSummary}}', 'Client Status Report', '{{reportUrl}}', '[{"id":"reviewed_proposal","label":"Reviewed the proposed cleanup against current customer context","required":true},{"id":"ready_to_confirm_or_return","label":"Ready to confirm or return with a reason","required":true}]', 'Submit Confirmation'],
    [SW_TASKS.DATA_CLEANUP_REVISE, 'Revise Customer Data Cleanup', 'The paired reviewer returned this cleanup update. Review their reason, revise the proposed customer status/data, and resubmit for confirmation.', 'Customer: {{customerName}}\nReturned reason: {{cleanupReturnReason}}\nPrevious proposal: {{cleanupProposalSummary}}', 'Client Status Report', '{{reportUrl}}', '', 'Resubmit Cleanup']
  ];
}

function swSeedAuthUsers_(sh) {
  if (sh.getLastRow() > 1) return;
  sh.getRange(2, 1, 1, SW_AUTH_USER_HEADERS.length).setValues([[
    'admin@example.com',
    'Admin Placeholder',
    'Admin',
    'N',
    '',
    '',
    'Y',
    '',
    'Replace with real users by running sw_adminSetWorkflowPassword(email, password, name, roles).'
  ]]);
}
