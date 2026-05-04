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
  if (task.taskType === SW_TASKS.DIAMOND_PROPOSE && (!data.proposalStones || !data.proposalStones.length)) {
    throw new Error('Add at least one proposed diamond before completing this task.');
  }
  if (task.taskType === SW_TASKS.DIAMOND_PROPOSE) {
    var missingRequirements = swDiamondCustomerRequirementsMissing_(data.customerDiamondRequirements || {});
    if (missingRequirements.length) {
      throw new Error('Complete customer diamond requirements: ' + missingRequirements.join(', ') + '.');
    }
  }
  if (task.taskType === SW_TASKS.DIAMOND_ORDER && (!data.diamondOrderDecisions || !data.diamondOrderDecisions.length)) {
    throw new Error('Select at least one diamond order decision before completing this task.');
  }
  if (task.taskType === SW_TASKS.DIAMOND_ORDER) {
    var missingOrderDecision = (data.diamondOrderDecisions || []).filter(function (item) {
      return item && item.rowIndex && !swTrim_(item.decision);
    });
    if (missingOrderDecision.length) {
      throw new Error('Select On the Way or Not Approved for every proposed diamond before completing this task.');
    }
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
  var rawBrand = task.brand || appt.brand || '';
  return {
    customerName: task.customerName || appt.customerName || '',
    brand: swTemplateBrandName_(rawBrand),
    brandRaw: rawBrand,
    appointmentDate: task.visitDate || appt.visitDate || '',
    appointmentTime: task.visitTime || appt.visitTime || '',
    appointmentDateTime: [task.visitDate || appt.visitDate || '', task.visitTime || appt.visitTime || ''].filter(Boolean).join(' '),
    visitType: task.visitType || appt.visitType || '',
    clientAdvisor: appt.assignedRep || '',
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
    diamondCustomerRequirements: extra.diamondCustomerRequirements || appt.dvCustomerLookingFor || 'Not captured yet.',
    diamondVarietyStrategy: extra.diamondVarietyStrategy || appt.dvVarietyStrategy || 'Not captured yet.',
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

function swTemplateBrandName_(brand) {
  var b = swHeaderKey_(brand);
  if (!b) return '';
  if (b.indexOf('vvs') >= 0) return 'VVS Jewelry Co.';
  if (b.indexOf('hung') >= 0 || b.indexOf('phat') >= 0 || b.indexOf('hpusa') >= 0 || b === 'hp') return 'Hung Phat';
  return swTrim_(brand);
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
    SW_TASKS.DIAMOND_ORDER_ACK_REP,
    SW_TASKS.DIAMOND_ORDER_ACK_JOC,
    SW_TASKS.DIAMOND_ETA_REP,
    SW_TASKS.DIAMOND_ETA_JOC
  ].indexOf(taskType) >= 0;
}

function swRenderTemplate_(template, data) {
  return swRenderTemplatePasses_(template, data, 4);
}

function swRenderTemplatePasses_(template, data, maxPasses) {
  var out = String(template || '');
  var passes = Number(maxPasses) || 1;
  for (var i = 0; i < passes; i++) {
    var changed = false;
    out = out.replace(/\{\{\s*([^{}]+?)\s*\}\}/g, function (match, rawKey) {
      var value = swTemplateValue_(data, rawKey);
      changed = true;
      if (value == null) return '';
      return String(value);
    });
    if (!changed || out.indexOf('{{') < 0) break;
  }
  return out;
}

function swTemplateValue_(data, rawKey) {
  if (!data) return null;
  var key = swTrim_(rawKey);
  if (!key) return null;
  if (Object.prototype.hasOwnProperty.call(data, key)) return data[key];

  var alias = swTemplateAliasKey_(key);
  if (alias && Object.prototype.hasOwnProperty.call(data, alias)) return data[alias];
  return null;
}

function swTemplateAliasKey_(key) {
  var norm = swHeaderKey_(key);
  var aliases = {
    customername: 'customerName',
    customer: 'customerName',
    clientname: 'customerName',
    client: 'customerName',
    name: 'customerName',
    appointmentdate: 'appointmentDate',
    apptdate: 'appointmentDate',
    visitdate: 'appointmentDate',
    date: 'appointmentDate',
    appointmenttime: 'appointmentTime',
    appttime: 'appointmentTime',
    visittime: 'appointmentTime',
    time: 'appointmentTime',
    appointmentdatetime: 'appointmentDateTime',
    appointmentdateandtime: 'appointmentDateTime',
    apptdatetime: 'appointmentDateTime',
    visitdatetime: 'appointmentDateTime',
    brand: 'brand',
    company: 'brand',
    brandname: 'brand',
    brandraw: 'brandRaw',
    rawbrand: 'brandRaw',
    visittype: 'visitType',
    appointmenttype: 'visitType',
    clientadvisor: 'clientAdvisor',
    clientadviser: 'clientAdvisor',
    advisor: 'clientAdvisor',
    adviser: 'clientAdvisor',
    assignedrep: 'assignedRep',
    stylist: 'assignedRep',
    assignedrepemail: 'assignedRepEmail',
    assistedrep: 'assistedRep',
    assistedrepemail: 'assistedRepEmail'
  };
  return aliases[norm] || null;
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
  swScanMissingTemplateFields_(template, data, missing, 0);
  return Object.keys(missing).sort();
}

function swScanMissingTemplateFields_(template, data, missing, depth) {
  if (depth > 4) return;
  String(template || '').replace(/\{\{\s*([^{}]+?)\s*\}\}/g, function (_, key) {
    var value = swTemplateValue_(data, key);
    if (!value) {
      missing[swTrim_(key)] = true;
    } else if (String(value).indexOf('{{') >= 0) {
      swScanMissingTemplateFields_(value, data, missing, depth + 1);
    }
    return '';
  });
}
