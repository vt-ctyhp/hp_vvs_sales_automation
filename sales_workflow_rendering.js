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
  var isNoShowChecklist = task.taskType === SW_TASKS.CHECKLIST &&
    typeof swIsNoShowOutcome_ === 'function' &&
    swIsNoShowOutcome_(data.appointmentOutcome || '');
  if (checklist && checklist.length && !isNoShowChecklist) {
    var checked = data.checklist || {};
    var missing = [];
    checklist.forEach(function (item) {
      if (item.required !== false && !checked[item.id]) missing.push(item.label || item.id);
    });
    if (missing.length) throw new Error('Complete required checklist items: ' + missing.join(', '));
  }
  if (task.taskType === SW_TASKS.CHECKLIST) {
    var outcome = swTrim_(data.appointmentOutcome || '');
    if (outcome !== 'Completed' && outcome !== 'No Show') {
      throw new Error('Select Completed or No Show before completing the appointment.');
    }
    if (outcome === 'Completed' && typeof swAppointmentHasPrimaryRecording_ === 'function' &&
        !swAppointmentHasPrimaryRecording_(ss, task.root || task.appt || '')) {
      throw new Error('Upload the appointment recording before completing a completed appointment.');
    }
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
  if (typeof swValidatePostConsultCompletion_ === 'function') {
    swValidatePostConsultCompletion_(task, data);
  }
  if (typeof swValidateDataCleanupCompletion_ === 'function') {
    swValidateDataCleanupCompletion_(task, data);
  }
}

function swRenderDataForTask_(task, payload) {
  payload = payload || {};
  var appt = payload.appointment || {};
  var extra = payload.extra || {};
  var completion = payload.completion || {};
  var rawBrand = task.brand || appt.brand || '';
  var visitTime = swFormatAppointmentTime_(task.visitTime || appt.visitTime || '');
  var hybridMessage = extra.hybridMessage || [extra.welcomeMessage, extra.locationMsg].filter(function (value) {
    return swTrim_(value);
  }).join('\n\n');
  return {
    customerName: task.customerName || appt.customerName || '',
    brand: swTemplateBrandName_(rawBrand),
    brandRaw: rawBrand,
    appointmentDate: task.visitDate || appt.visitDate || '',
    appointmentTime: visitTime,
    appointmentDateTime: [task.visitDate || appt.visitDate || '', visitTime].filter(Boolean).join(' '),
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
    soNumber: extra.soNumber || appt.so || 'Not assigned yet',
    salesStage: extra.salesStage || appt.salesStage || '',
    conversionStatus: extra.convStatus || appt.convStatus || '',
    convStatus: extra.convStatus || appt.convStatus || '',
    customOrder: extra.customOrder || appt.customOrder || '',
    nextSteps: extra.nextSteps || appt.nextSteps || 'Not captured yet',
    designRequest: extra.designRequest || appt.designRequest || 'Not captured yet',
    deadline3d: extra.deadline3d || appt.deadline3d || 'Not recorded yet',
    productionDeadline: extra.productionDeadline || appt.productionDeadline || 'Not recorded yet',
    waxStatus: extra.waxStatus || appt.waxStatus || 'No active wax request',
    waxDeadlineAdmin: extra.waxDeadlineAdmin || appt.waxDeadlineAdmin || 'Not recorded yet',
    waxRequestUrl: extra.waxRequestUrl || appt.waxRequestUrl || '',
    waxRequestSummary: extra.waxRequestSummary || 'No open wax requests',
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
    hybridMessage: hybridMessage,
    welcomeImageUrl: extra.welcomeImageUrl || '',
    recapDraft: extra.recapDraft || completion.recapText || '',
    approvedText: extra.approvedText || completion.approvedText || extra.recapDraft || '',
    artifactId: extra.artifactId || '',
    workflowStage: extra.workflowStage || '',
    transcriptDocUrl: extra.transcriptDocUrl || '',
    summaryDocUrl: extra.summaryDocUrl || '',
    summaryJsonUrl: extra.summaryJsonUrl || '',
    salesBrief: extra.salesBrief || '',
    reviewFlags: extra.reviewFlags || '',
    clientFollowUpDraft: extra.clientFollowUpDraft || extra.recapDraft || '',
    cleanupProposedBy: swDeepValue_(extra, ['cleanupCase', 'proposedBy']) || '',
    cleanupReturnReason: swDeepValue_(extra, ['cleanupCase', 'returnReason']) || '',
    cleanupProposalSummary: typeof swDataCleanupProposalSummary_ === 'function'
      ? swDataCleanupProposalSummary_(swDeepValue_(extra, ['cleanupCase', 'proposal']) || {})
      : ''
  };
}

function swEffectiveTemplateForTaskType_(taskType, template) {
  template = template || {};
  var out = {
    taskTitle: template.taskTitle || taskType,
    instructions: template.instructions || '',
    template: template.template || '',
    attachmentLabel: template.attachmentLabel || '',
    attachmentUrl: template.attachmentUrl || '',
    checklistJson: template.checklistJson || '',
    primaryAction: template.primaryAction || 'Complete'
  };
  if (taskType === SW_TASKS.WELCOME) {
    if (swShouldUseDefaultWelcomeTemplate_(out.template)) out.template = '{{welcomeMessage}}';
    if (String(out.attachmentUrl || '').indexOf('welcomeImageUrl') < 0) {
      out.attachmentLabel = 'Welcome Journey Image';
      out.attachmentUrl = '{{welcomeImageUrl}}';
    }
  }
  if (taskType === SW_TASKS.MAP) {
    if (swShouldUseDefaultMapTemplate_(out.template)) out.template = '{{locationMsg}}';
    if (String(out.attachmentUrl || '').indexOf('mapLink') < 0) {
      out.attachmentLabel = 'Map / Instructions';
      out.attachmentUrl = '{{mapLink}}';
    }
  }
  if (taskType === SW_TASKS.HYBRID) {
    if (swShouldUseDefaultHybridTemplate_(out.template)) out.template = '{{hybridMessage}}';
    if (String(out.attachmentUrl || '').indexOf('mapLink') < 0) {
      out.attachmentLabel = 'Map / Instructions';
      out.attachmentUrl = '{{mapLink}}';
    }
  }
  return out;
}

function swShouldUseDefaultWelcomeTemplate_(template) {
  var text = String(template || '');
  if (!swTrim_(text)) return true;
  if (text.indexOf('welcomeMessage') < 0) return true;
  if (text.indexOf('welcomeImageUrl') >= 0) return true;
  return swContainsGoogleDriveLink_(text);
}

function swShouldUseDefaultMapTemplate_(template) {
  var text = String(template || '');
  if (!swTrim_(text)) return true;
  if (text.indexOf('locationMsg') < 0) return true;
  if (text.indexOf('mapLink') >= 0) return true;
  return swContainsGoogleDriveLink_(text);
}

function swShouldUseDefaultHybridTemplate_(template) {
  var text = String(template || '');
  if (!swTrim_(text)) return true;
  if (text.indexOf('welcomeImageUrl') >= 0) return true;
  if (text.indexOf('mapLink') >= 0) return true;
  if (swContainsGoogleDriveLink_(text)) return true;
  if (text.indexOf('hybridMessage') >= 0 && text.indexOf('locationMsg') >= 0) return true;
  if (text.indexOf('hybridMessage') >= 0) return false;
  if (swIsLegacyDefaultHybridTemplate_(text)) return true;
  if (/we are looking forward to seeing you/i.test(text)) return true;
  return false;
}

function swIsLegacyDefaultHybridTemplate_(template) {
  var compact = String(template || '').replace(/\s+/g, '');
  return compact === '{{welcomeMessage}}{{locationMsg}}';
}

function swContainsGoogleDriveLink_(text) {
  return /\b(?:https?:\/\/)?(?:drive|docs)\.google\.com\//i.test(String(text || ''));
}

function swDefaultHybridMessageTemplate_() {
  return '{{welcomeMessage}}\n\n{{locationMsg}}';
}

function swShouldUseDefaultHybridConfigValue_(value) {
  var text = String(value || '');
  if (!swTrim_(text)) return true;
  if (text.indexOf('welcomeImageUrl') >= 0) return true;
  if (text.indexOf('mapLink') >= 0) return true;
  return swContainsGoogleDriveLink_(text);
}

function swRenderedCopyableTemplateForTask_(task, template, data) {
  var rendered = template && template.template ? swRenderTemplate_(template.template, data) : '';
  if (!swIsClientMessageTaskType_(task && task.taskType)) return rendered;
  return swStripGoogleDriveLinks_(rendered);
}

function swIsClientMessageTaskType_(taskType) {
  return [
    SW_TASKS.WELCOME,
    SW_TASKS.HYBRID,
    SW_TASKS.MAP,
    SW_TASKS.FINAL
  ].indexOf(taskType) >= 0;
}

function swStripGoogleDriveLinks_(text) {
  var out = String(text || '');
  out = out.replace(/\b(?:https?:\/\/)?(?:drive|docs)\.google\.com\/[^\s<>"')]+/gi, '');
  out = out.replace(/[ \t]+\n/g, '\n');
  out = out.replace(/\n{3,}/g, '\n\n');
  return swTrim_(out);
}

function swAttachmentsForTask_(task, template, data) {
  var out = [];
  var primaryUrl = template.attachmentUrl ? swRenderTemplate_(template.attachmentUrl, data) : '';
  var primaryLabel = template.attachmentLabel ? swRenderTemplate_(template.attachmentLabel, data) : '';
  swPushAttachment_(out, primaryLabel, primaryUrl);

  if (task.taskType === SW_TASKS.MAP || task.taskType === SW_TASKS.HYBRID) {
    swPushAttachment_(out, 'Map / Instructions', data.mapLink || '');
  }
  if (task.taskType === SW_TASKS.WELCOME || task.taskType === SW_TASKS.HYBRID) {
    swPushAttachment_(out, 'Welcome Journey Image', data.welcomeImageUrl || '');
  }
  if (swIsDiamondTaskType_(task.taskType)) {
    swPushAttachment_(out, 'Quotation Sheet', data.quotationUrl || '');
    swPushAttachment_(out, '3D Tracker', data.tracker3dUrl || '');
    swPushAttachment_(out, '200_ Diamond Tracker', data.diamondTrackerUrl || '');
  }
  if (task.taskType === SW_TASKS.APPROVE) {
    swPushAttachment_(out, 'AI Follow-Up Doc', data.summaryDocUrl || '');
    swPushAttachment_(out, 'Transcript Doc', data.transcriptDocUrl || '');
  }
  if (task.taskType === SW_TASKS.FINAL) {
    swPushAttachment_(out, 'AI Follow-Up Doc', data.summaryDocUrl || '');
    swPushAttachment_(out, 'Transcript Doc', data.transcriptDocUrl || '');
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
    text += '\n{{mapLink}}\n{{hybridMessage}}\n{{welcomeImageUrl}}';
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
