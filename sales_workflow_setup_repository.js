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
      if (String(row[3] || '').indexOf('welcomeMessage') < 0) {
        sh.getRange(rowIndex, 4).setValue('{{welcomeMessage}}');
      }
      if (String(row[5] || '').indexOf('welcomeImageUrl') < 0) {
        sh.getRange(rowIndex, 5).setValue('Welcome Journey Image');
        sh.getRange(rowIndex, 6).setValue('{{welcomeImageUrl}}');
      }
    }
    if (taskType === SW_TASKS.MAP && String(row[3] || '').indexOf('locationMsg') < 0) {
      sh.getRange(rowIndex, 4).setValue('{{locationMsg}}\n{{mapLink}}');
    }
    if (taskType === SW_TASKS.HYBRID && String(row[3] || '').indexOf('locationMsg') < 0) {
      sh.getRange(rowIndex, 4).setValue(String(row[3] || '') + '\n\n{{locationMsg}}\n{{mapLink}}');
    }
  }
}

function swDefaultTemplates_() {
  return [
    [SW_TASKS.ASSIGN, 'Assign Appointment', 'System-owned assignment record. No manual action needed.', '', '', '', '', 'Assigned'],
    [SW_TASKS.WELCOME, 'Send Welcome to Your Ring Journey Text', 'Send the brand-specific welcome message and welcome image, then mark it sent.', '{{welcomeMessage}}', 'Welcome Journey Image', '{{welcomeImageUrl}}', '', 'Mark Sent'],
    [SW_TASKS.HYBRID, 'Send Hybrid Welcome + Instructions', 'Appointment is within 24 hours. Send the combined welcome and instructions.', 'Hi {{customerName}}, we are looking forward to seeing you {{appointmentDate}} at {{appointmentTime}}. Please review the map/instructions before you arrive. Your stylist is {{assignedRep}}.\n\n{{locationMsg}}\n{{mapLink}}', 'Map / Instructions', '{{mapLink}}', '', 'Mark Sent'],
    [SW_TASKS.MAP, 'Send Map & Instructions', 'Send the map and appointment instructions.', '{{locationMsg}}\n{{mapLink}}', 'Map / Instructions', '{{mapLink}}', '', 'Mark Sent'],
    [SW_TASKS.REVIEW, 'Review Appointment Folder', 'Review the intake form, inspiration images, and customer folder before the appointment.', '', 'Client Folder', '{{clientFolder}}', '', 'Acknowledged & Reviewed'],
    [SW_TASKS.CHECKLIST, 'Appointment Day Checklist', 'Complete each appointment-day item before marking complete.', '', '', '', '[{"id":"printed_intake","label":"Printed intake form","required":true},{"id":"recorded_appointment","label":"Recorded appointment","required":true},{"id":"uploaded_recap","label":"Uploaded recap","required":true},{"id":"uploaded_photos","label":"Uploaded intake photos","required":true},{"id":"goody_bag","label":"Gave goody bag","required":true}]', 'Complete Checklist'],
    [SW_TASKS.PROCESS, 'Process Appointment Data', 'Upload the recording, generate the recap draft, and submit it here.', '', 'Client Folder', '{{clientFolder}}', '', 'Submit Recap Draft'],
    [SW_TASKS.APPROVE, 'Approve/Edit Recap Message', 'Review the JOC recap draft. Edit if needed, then finalize.', '{{recapDraft}}', '', '', '', 'Finalized'],
    [SW_TASKS.FINAL, 'Send Final Recap Text', 'Send the finalized recap message, then mark it sent.', '{{approvedText}}', '', '', '', 'Mark Sent'],
    [SW_TASKS.DIAMOND_PROPOSE, 'Propose Diamonds for Viewing', 'For Diamond Viewing appointments, propose stones immediately. Use the 200_ tracker as the diamond source of truth and keep the customer appointment/date context aligned.', 'Customer: {{customerName}}\nAppointment: {{appointmentDate}} {{appointmentTime}}\nDiamond status: {{diamondSummary}}\nLatest safe proposal target: {{diamondProposalTarget}}', '200_ Diamond Tracker', '{{diamondTrackerUrl}}', '[{"id":"reviewed_customer_needs","label":"Reviewed customer preferences and appointment context","required":true},{"id":"proposed_stones","label":"Proposed stones in 200_ diamond tracker","required":true},{"id":"notified_joc","label":"Notified JOC that proposed stones are ready for quotation research","required":true}]', 'Mark Proposed'],
    [SW_TASKS.DIAMOND_QUOTE, 'Prepare Diamond Viewing Quotation', 'Fill the quotation sheet and complete price research. Use the refresh buttons on the task card if 200_ diamond data or 3D tracker details changed after the quote was created.', 'Quotation: {{quotationUrl}}\n3D Tracker: {{tracker3dUrl}}\nDiamond tracker: {{diamondTrackerUrl}}\nDiamond status: {{diamondSummary}}', 'Quotation Sheet', '{{quotationUrl}}', '[{"id":"quote_link_checked","label":"Quotation sheet link opens correctly","required":true},{"id":"diamonds_refreshed","label":"Diamond options refreshed from 200_ or confirmed current","required":true},{"id":"settings_refreshed","label":"3D setting details refreshed or confirmed current","required":true},{"id":"price_research_done","label":"Price research completed in quotation sheet","required":true}]', 'Quotation Ready'],
    [SW_TASKS.DIAMOND_ORDER, 'Order Diamonds', 'Review proposed stones and mark each approved stone as On the Way or Not Approved in 200_. This task belongs to the diamond order admin.', 'Pending proposed stones: {{diamondActionSummary}}\nTracker: {{diamondTrackerUrl}}', '200_ Diamond Tracker', '{{diamondTrackerUrl}}', '[{"id":"reviewed_proposed_stones","label":"Reviewed proposed stones","required":true},{"id":"updated_order_status","label":"Updated Order Status in 200_","required":true},{"id":"recorded_order_date","label":"Recorded Purchased / Ordered Date for ordered stones","required":true}]', 'Order Updated'],
    [SW_TASKS.DIAMOND_TRACK, 'Track Diamond ETA', 'Check shipping/tracking for on-the-way diamonds. Enter ETA and status here; the 200_ tracker remains the source of truth. Late or concerning ETAs create rep/JOC alert tasks.', 'On-the-way stones: {{diamondActionSummary}}\nAppointment: {{appointmentDate}} {{appointmentTime}}', '200_ Diamond Tracker', '{{diamondTrackerUrl}}', '[{"id":"checked_tracking","label":"Checked diamond tracking/vendor ETA","required":true},{"id":"entered_eta","label":"Entered ETA/status on this task","required":true}]', 'Save ETA'],
    [SW_TASKS.DIAMOND_DELIVERY, 'Confirm Diamond Delivery', 'Confirm ordered diamonds were received and mark delivered/in stock. Return deadline must be based on Purchased / Ordered Date + 30 days.', 'Awaiting delivery: {{diamondActionSummary}}\nTracker: {{diamondTrackerUrl}}', '200_ Diamond Tracker', '{{diamondTrackerUrl}}', '[{"id":"confirmed_received","label":"Confirmed diamonds were received","required":true},{"id":"updated_delivery_status","label":"Updated delivery status in 200_","required":true},{"id":"return_deadline_checked","label":"Verified return due date is order date plus 30 days","required":true}]', 'Delivery Confirmed'],
    [SW_TASKS.DIAMOND_DECISIONS, 'Record Diamond Decisions', 'After viewing, mark each stone as Purchase or Return. Reconfirm diamond dimensions against the latest 3D tracker details and send manufacturing the confirmed dimensions using the generated message.', 'Decision context: {{diamondActionSummary}}\nManufacturing message:\n{{manufacturingMessage}}', '3D Tracker', '{{tracker3dUrl}}', '[{"id":"marked_purchase_return","label":"Marked Purchase/Return decisions in 200_","required":true},{"id":"confirmed_dimensions","label":"Confirmed diamond dimensions against latest 3D tracker","required":true},{"id":"manufacturing_messaged","label":"Sent manufacturing the confirmed dimensions message","required":true}]', 'Decisions Recorded'],
    [SW_TASKS.DIAMOND_RETURN, 'Return Diamonds', 'Review diamonds marked Return or not purchased. Diamonds must be returned within 30 days of Purchased / Ordered Date, not delivery date.', 'Return queue: {{diamondActionSummary}}\nTracker: {{diamondTrackerUrl}}', '200_ Diamond Tracker', '{{diamondTrackerUrl}}', '[{"id":"checked_due_dates","label":"Checked return due dates","required":true},{"id":"returned_unwanted_stones","label":"Returned all due unwanted stones or escalated blocker","required":true},{"id":"updated_tracker","label":"Updated 200_ with return status/notes","required":true}]', 'Return Updated'],
    [SW_TASKS.DIAMOND_ETA_REP, 'Review Diamond ETA Risk', 'Diamond ETA/status needs assigned-rep review because it is late or concerning for the Diamond Viewing appointment.', 'ETA issue: {{diamondEtaIssue}}\nAppointment: {{appointmentDate}} {{appointmentTime}}\nTracker: {{diamondTrackerUrl}}', '200_ Diamond Tracker', '{{diamondTrackerUrl}}', '[{"id":"reviewed_eta_risk","label":"Reviewed ETA risk and customer impact","required":true},{"id":"coordinated_next_step","label":"Coordinated next step with JOC/order team","required":true}]', 'Risk Reviewed'],
    [SW_TASKS.DIAMOND_ETA_JOC, 'Review Diamond ETA for Quotation', 'Diamond ETA/status needs JOC review because it is late or concerning for the Diamond Viewing appointment.', 'ETA issue: {{diamondEtaIssue}}\nQuotation: {{quotationUrl}}\nTracker: {{diamondTrackerUrl}}', 'Quotation Sheet', '{{quotationUrl}}', '[{"id":"reviewed_eta_risk","label":"Reviewed ETA risk against quotation plan","required":true},{"id":"updated_quote_or_notes","label":"Updated quotation/notes if ETA changes options","required":true}]', 'Risk Reviewed']
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
