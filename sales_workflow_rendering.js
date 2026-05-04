/**
 * Sales workflow rendering: template data, attachments, missing fields, and completion validation.
 */

function swValidateCompletion_(ss, task, data) {
  var template = swTemplateForType_(ss, task.taskType);
  var payload = swParseJson_(task.payloadJson, {});
  var renderData = swRenderDataForTask_(task, payload);
  var missingTemplateFields = swMissingFieldsForTask_(task, template, renderData);
  if (missingTemplateFields.length) {
    throw new Error('Missing template fields before completion: ' + missingTemplateFields.join(', '));
  }

  var checklist = swParseJson_(template.checklistJson, []);
  if (checklist && checklist.length) {
    var checked = data.checklist || {};
    var missing = [];
    checklist.forEach(function (item) {
      if (item.required !== false && !checked[item.id]) missing.push(item.label || item.id);
    });
    if (missing.length) throw new Error('Complete required checklist items: ' + missing.join(', '));
  }
  if (task.taskType === SW_TASKS.PROCESS && !swTrim_(data.recapText)) {
    throw new Error('Enter the recap draft before completing this task.');
  }
  if (task.taskType === SW_TASKS.APPROVE && !swTrim_(data.approvedText)) {
    throw new Error('Enter the approved recap text before finalizing.');
  }
  if (task.taskType === SW_TASKS.DIAMOND_TRACK && !swTrim_(data.trackingEta) && !swTrim_(data.trackingStatus)) {
    throw new Error('Enter a tracking ETA or tracking status before saving this task.');
  }
  if (task.taskType === SW_TASKS.DIAMOND_ORDER && (!data.diamondOrderDecisions || !data.diamondOrderDecisions.length)) {
    throw new Error('Select at least one diamond order decision before completing this task.');
  }
  if (task.taskType === SW_TASKS.DIAMOND_DECISIONS && (!data.diamondDecisions || !data.diamondDecisions.length)) {
    throw new Error('Select at least one Purchase/Return decision before completing this task.');
  }
}

function swRenderDataForTask_(task, payload) {
  payload = payload || {};
  var appt = payload.appointment || {};
  var extra = payload.extra || {};
  var completion = payload.completion || {};
  return {
    customerName: task.customerName || appt.customerName || '',
    brand: task.brand || appt.brand || '',
    appointmentDate: task.visitDate || appt.visitDate || '',
    appointmentTime: task.visitTime || appt.visitTime || '',
    visitType: task.visitType || appt.visitType || '',
    assignedRep: appt.assignedRep || '',
    assignedRepEmail: appt.assignedRepEmail || '',
    assistedRep: appt.assistedRep || '',
    assistedRepEmail: appt.assistedRepEmail || '',
    clientFolder: appt.clientFolder || '',
    reportUrl: appt.reportUrl || '',
    quotationUrl: extra.quotationUrl || appt.quotationUrl || '',
    tracker3dUrl: extra.tracker3dUrl || appt.tracker3dUrl || '',
    diamondTrackerUrl: extra.diamondTrackerUrl || '',
    diamondSummary: extra.diamondSummary || appt.dvStonesSummary || '',
    diamondProposalTarget: extra.diamondProposalTarget || '',
    diamondActionSummary: extra.diamondActionSummary || '',
    diamondEtaIssue: extra.diamondEtaIssue || '',
    manufacturingMessage: extra.manufacturingMessage || '',
    mapLink: extra.mapLink || '',
    locationMsg: extra.locationMsg || '',
    welcomeMessage: extra.welcomeMessage || '',
    welcomeImageUrl: extra.welcomeImageUrl || '',
    recapDraft: extra.recapDraft || completion.recapText || '',
    approvedText: extra.approvedText || completion.approvedText || extra.recapDraft || ''
  };
}

function swAttachmentsForTask_(task, template, data) {
  var out = [];
  var primaryUrl = template.attachmentUrl ? swRenderTemplate_(template.attachmentUrl, data) : '';
  var primaryLabel = template.attachmentLabel ? swRenderTemplate_(template.attachmentLabel, data) : '';
  swPushAttachment_(out, primaryLabel, primaryUrl);

  if (task.taskType === SW_TASKS.WELCOME || task.taskType === SW_TASKS.HYBRID) {
    swPushAttachment_(out, 'Welcome Journey Image', data.welcomeImageUrl || '');
  }
  if (swIsDiamondTaskType_(task.taskType)) {
    swPushAttachment_(out, 'Quotation Sheet', data.quotationUrl || '');
    swPushAttachment_(out, '3D Tracker', data.tracker3dUrl || '');
    swPushAttachment_(out, '200_ Diamond Tracker', data.diamondTrackerUrl || '');
  }
  return out;
}

function swPushAttachment_(out, label, url) {
  url = swTrim_(url);
  if (!url) return;
  for (var i = 0; i < out.length; i++) {
    if (swTrim_(out[i].url) === url) return;
  }
  out.push({ label: swTrim_(label) || url, url: url });
}

function swIsDiamondTaskType_(taskType) {
  return [
    SW_TASKS.DIAMOND_PROPOSE,
    SW_TASKS.DIAMOND_QUOTE,
    SW_TASKS.DIAMOND_ORDER,
    SW_TASKS.DIAMOND_TRACK,
    SW_TASKS.DIAMOND_DELIVERY,
    SW_TASKS.DIAMOND_DECISIONS,
    SW_TASKS.DIAMOND_RETURN,
    SW_TASKS.DIAMOND_ETA_REP,
    SW_TASKS.DIAMOND_ETA_JOC
  ].indexOf(taskType) >= 0;
}

function swRenderTemplate_(template, data) {
  return String(template || '').replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g, function (_, key) {
    return data[key] == null ? '' : String(data[key]);
  });
}

function swMissingFieldsForTask_(task, template, data) {
  var text = [template.template || '', template.attachmentUrl || ''].join('\n');
  if (task.taskType === SW_TASKS.WELCOME) {
    text += '\n{{welcomeMessage}}\n{{welcomeImageUrl}}';
  }
  if (task.taskType === SW_TASKS.HYBRID) {
    text += '\n{{mapLink}}\n{{locationMsg}}\n{{welcomeImageUrl}}';
  }
  return swMissingTemplateFields_(text, data);
}

function swMissingTemplateFields_(template, data) {
  var missing = {};
  String(template || '').replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g, function (_, key) {
    if (!data[key]) missing[key] = true;
    return '';
  });
  return Object.keys(missing).sort();
}
