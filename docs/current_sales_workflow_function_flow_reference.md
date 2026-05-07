# Current Sales Workflow Function And Data Flow Reference

## 1. Purpose

This document explains the current Sales Workflow structure before the domain-tab migration. It is a current-state map of the functions, UI surfaces, triggers, sheets, and data flows that need to stay intact while `00_Master Appointments` is gradually replaced as the operational source of truth.

Use this document to answer:

- which function currently serves a user action;
- which function reads or writes which sheet;
- which workflow depends on `00_Master Appointments`;
- which functions must be rewired to a future domain adapter;
- which smoke tests should be run after changing a flow.

This is not a proposal for the new infrastructure. It documents the structure that exists now.

## 2. Current Source Tabs And External Stores

| Store | Current Role |
|---|---|
| `00_Master Appointments` | Main operational appointment/customer/root table. Most current workflows read from it, and many write customer, appointment, status, 3D, payment, owner, and compatibility fields into it. |
| `02_Form_Inbox` | Raw form intake staging source for old/current form submit paths. |
| `03_Client_Status_Log` | Client status history used by admin health and status flows. |
| `05_Wax_Requests` | Wax request source table. Wax status is mirrored back to Master. |
| `07_Root_Index` | Root last-touch/health support for admin dashboard. |
| `_SalesTaskQueue` | Workflow task source of truth. |
| `_SalesTaskLog` | Task lifecycle event log. |
| `_SalesWorkflowUsers` | Login users, roles, active state, and canonical workflow identities. |
| `_SalesWorkflowConfig` | Feature flags, config values, and read-model serving flags. |
| `_SalesWorkflowTemplates` | Task templates and copyable messages. |
| `_SalesDataCleanup` | Temporary stale-customer cleanup campaign source. |
| `_AppointmentArtifacts` | Appointment recordings, transcript, summary, AI brief, upload, and Drive artifact metadata. |
| `10_Roster_Schedule` | Employee schedule extension data. |
| `Schedule Changes` | One-off schedule overrides. |
| `_SW_*ReadModel` tabs | Derived hidden serving tabs for task, customer, diamond, appointment, calendar, payment, and admin dashboard reads. |
| 200 stones workbook | Per-stone diamond inventory, proposal, order, tracking, return, and Loupe360 data. |
| Payment ledger workbook | Individual payment, invoice, receipt, and document records. |
| External order/quote/3D tracker workbooks | Operational order, quotation, and tracker data. |
| Drive folders | Customer folders, appointment folders, uploads, generated docs, PDFs, transcripts, and summaries. |
| Acuity API | External appointment booking/status source. |
| Google Form | Current bridge from Acuity/form intake into `onFormSubmit`. |

## 3. Runtime Surfaces

### 3.1 Dashboard Web App

Files:

- `WebApp.js` serves the HtmlService app.
- `Index.html` contains the dashboard UI and client-side `serverCall(...)` wrapper.
- `sales_workflow_api.js` exposes most dashboard callable functions.
- `sales_workflow_customer_search.js` and `sales_workflow_admin_dashboard.js` expose specialized dashboard APIs.

Current dashboard views:

- My Queue
- Calendar
- Customer Search
- Customer Pipeline / Admin Dashboard
- In-Stock Diamonds
- Diamond Tracking
- Bulk Returns
- JOC Coverage
- Admin Review
- Cleanup
- Schedules
- Manage Users

### 3.2 Sheet Menus And Dialogs

Files:

- `v1_sales_menu.js`
- `dlg_*.html`
- `start3d_server.js`
- `revision3d_server.js`
- `ClientStatus_v1.js`
- `Deadlines_v1.js`
- `Payments_v1.js`
- `Payment_Summary_v1.js`
- `WaxRequests.js`
- diamond dialog files

These flows often depend on the active selected Master row and use sheet dialogs rather than the dashboard token flow.

### 3.3 Background Jobs

Files:

- `sales_workflow_orchestrator.js`
- `Acuity_HPUSA.js`
- `Acuity_HPUSA_Lable.gs.js`
- `sales_workflow_appointment_artifacts.js`
- `sales_workflow_read_model.js`
- `Resolver.js`

The main background trigger is `sw_backgroundOrchestrator`, installed every 5 minutes. It serializes Acuity polling, Acuity label sync, appointment automation, task generation, URL repair, and read-model rebuilds.

### 3.4 External Upload And AI Surfaces

Files:

- `01_UploadEndpoint.js`
- `02_Workers.js`
- `03_SummaryRenderer.js`
- `04_AskController.js`
- `05_AskCore.js`
- `sales_workflow_appointment_artifacts.js`

These handle appointment recording uploads, transcript processing, OpenAI follow-up summary generation, client summary tabs, Ask Controller, and artifact references. They still use `RootApptID` and Master-derived folder/report data in several places.

## 4. Dashboard UI Call Map

| UI Action | Client Function | Server Function | Current Reads | Current Writes |
|---|---|---|---|---|
| App bootstrap | `bootstrap` | `sw_getBootstrap` | auth/users, config, task dashboard projection or task read model/queue | none |
| Login | login handler | `sw_login(...includeBootstrap)` | `_SalesWorkflowUsers`, then bootstrap reads | auth session cache |
| Logout | logout handler | `sw_logout` | auth/session | auth session cache |
| Queue tab | `loadTasks` | `sw_getMyTasks` | task dashboard projection, task read model, `_SalesTaskQueue` fallback | none |
| Task detail | `openTask` | `sw_getTaskDetail` | `_SalesTaskQueue`, templates, config, appointment row, artifacts, AI brief, task-specific context | none |
| Complete task | `completeTask` | `sw_completeTask` | task row, task payload, templates, source context | `_SalesTaskQueue`, `_SalesTaskLog`, plus task-specific source writes |
| Snooze task | `snoozeTask` | `sw_snoozeTask` | task row | `_SalesTaskQueue`, `_SalesTaskLog` |
| Claim task | `claimTask` | `sw_claimTask` | task row, current user | `_SalesTaskQueue`, `_SalesTaskLog` |
| Admin reassign task | admin detail action | `sw_adminReassignTask` | task row, user | `_SalesTaskQueue`, `_SalesTaskLog` |
| Admin owner assignment | admin detail action | `sw_adminAssignAppointmentOwners` | task row, Master root rows, workflow users | `00_Master Appointments`, `_SalesTaskQueue`, `_SalesTaskLog` |
| Block/unblock task | admin detail action | `sw_adminBlockTask`, `sw_adminUnblockTask` | task row | `_SalesTaskQueue`, `_SalesTaskLog` |
| Refresh queue | global button | `sw_generateSalesWorkflowTasks` | Master, users, config, templates, wax, diamond tracker, schedules | `_SalesTaskQueue`, `_SalesTaskLog`, sometimes Master owner assignment |
| Customer Search | `loadCustomerSearch` | `sw_searchCustomers` | customer read model/cache or Master fallback | none |
| Customer detail | `openCustomerSearchDetail` | `sw_getCustomerSearchDetail` | customer detail cache/read model/Master, payments, logs, artifacts, form options | none |
| Customer status save | `saveCustomerSearchStatus` | `sw_customerSearchUpdateStatus` | Master root/active row | Master status fields, status log/read-model invalidation |
| 3D deadline save | `saveCustomerSearch3DDeadline` | `sw_customerSearchUpdate3DDeadline` | Master root/active row | Master deadline fields, status log/read-model invalidation |
| 3D revision submit | `submitCustomerSearch3DRevision` | `sw_customerSearchSubmit3DRevision` | Master active row, 3D tracker/order data | 3D tracker log, Master fields as applicable |
| Wax request | `submitCustomerSearchWaxRequest` | `sw_customerSearchRequestWax` | Master/root/order mini context | `05_Wax_Requests`, Master wax mirror |
| Calendar month | `loadCalendar` | `sw_getCalendarAppointments` | calendar read model or Master fallback, AI brief index | none |
| Calendar AI brief | `loadCalendarAppointmentAiBrief` | `sw_getAppointmentAiBrief` | `_AppointmentArtifacts` and AI brief cache | none |
| Admin Dashboard | `loadAdminDashboard` | `sw_getAdminDashboard` | admin read model or Master/tasks/payments/root index/status log | none |
| Schedules | `loadEmployeeSchedules` | `sw_adminGetEmployeeSchedules` | `_SalesWorkflowUsers`, `10_Roster_Schedule`, `Schedule Changes` | none |
| Save schedules | schedule editor | `sw_adminSaveEmployeeSchedules` | users/schedule rows | `10_Roster_Schedule` |
| Save schedule override | schedule editor | `sw_adminUpsertScheduleChange` | users/schedule rows | `Schedule Changes` |
| Delete schedule override | schedule editor | `sw_adminDeleteScheduleChange` | schedule rows | `Schedule Changes` |
| Manage users | `openUserAdminPanel` | `sw_adminListWorkflowUsers` | `_SalesWorkflowUsers` | none |
| Save user | user admin panel | `sw_adminUpsertWorkflowUser` | users/roster | `_SalesWorkflowUsers`, `10_Roster_Schedule` link |
| Appointment upload folder | task detail button | `sw_getAppointmentUploadFolder` | task, Master root folder, Drive | Drive folders/cache |
| Sync appointment uploads | task detail button | `sw_syncAppointmentDriveUploads` | Drive drop folders, artifacts | `_AppointmentArtifacts` |
| In-stock diamonds | `loadInStockDiamonds` | `sw_getInStockDiamonds` | diamond read model or 200 workbook | none |
| Assign in-stock diamond | diamond action | `sw_assignInStockDiamond` | 200 workbook, selected root | 200 workbook, read-model invalidation |
| Loupe360 preview/apply | upload panel | `sw_previewLoupe360DiamondSync`, `sw_applyLoupe360DiamondSync` | uploaded sheet, 200 workbook | 200 workbook, temp Drive file cleanup |
| Diamond Tracking | `loadDiamondTracking` | `sw_getDiamondTrackingDashboard` | diamond read model or 200 workbook | none |
| Bulk Returns | `loadBulkReturns` | `sw_getBulkReturnCandidates` | diamond read model or 200 workbook | none |
| Submit bulk return | bulk return action | `sw_bulkMarkDiamondsReturnInProgress` | selected 200 rows | 200 workbook, task log/read-model invalidation |

## 5. Public Sales Workflow API Functions

### 5.1 Setup And Generation

| Function | Current Role | Current Reads | Current Writes |
|---|---|---|---|
| `sw_setupSalesWorkflow` | Ensures workflow sheets, headers, styling, config, templates, users, cleanup, artifacts, roster, and schedule changes exist. | existing sheets/config | `_SalesTaskQueue`, `_SalesTaskLog`, `_SalesWorkflowConfig`, `_SalesWorkflowTemplates`, `_SalesWorkflowUsers`, `_SalesDataCleanup`, `_AppointmentArtifacts`, `10_Roster_Schedule`, `Schedule Changes`, Master diamond requirement headers |
| `sw_generateSalesWorkflowTasks` | Locked public task-generation entrypoint. Redirects legacy triggers to orchestrator when appropriate. | Master, task queue, config, templates, users, roster, schedule changes, wax index, diamond tracker | `_SalesTaskQueue`, `_SalesTaskLog`, possible Master owner assignment |
| `swGenerateSalesWorkflowTasksUnlocked_` | Internal generator body. Sets up context, reads Master appointments, prepares round robin, scans each appointment, generates/updates tasks, and flushes deferred writes. | same as above | same as above |
| `sw_tryGenerateSalesWorkflowTasksAfterSubmit_` | Best-effort short-lock queue refresh used as rollback when async task-generation request is disabled or cannot be recorded. | same as generator | same as generator if lock acquired |
| `sw_installSalesWorkflowTriggers` | Installs the single background orchestrator trigger. | script triggers | Apps Script triggers |
| `sw_auditDuplicateTasks` | Logs duplicate task audit. | `_SalesTaskQueue` | none |
| `sw_cleanupDuplicateTasksDryRun` | Plans duplicate cleanup without mutation. | `_SalesTaskQueue` | none |
| `sw_cleanupDuplicateTasksApply` | Blocks duplicate pending task rows. | `_SalesTaskQueue` | `_SalesTaskQueue`, `_SalesTaskLog` |

### 5.2 Read-Only Dashboard APIs

| Function | Current Role | Current Reads | Current Writes |
|---|---|---|---|
| `sw_getBootstrap` | Returns user, visible views, My Queue tasks, and counts. | auth/session, users, config, task dashboard projection or task read model/queue | none |
| `sw_getMyTasks` | Returns tasks for `mine`, `cleanup`, `coverage`, or `admin` view. | users, config, task dashboard projection or task read model/queue | none |
| `sw_getTaskDetail` | Returns task detail, rendered template, attachments, missing fields, appointment payload, artifacts, upload folders, and AI brief when relevant. | task row/cache, templates, config, Master payload, artifacts | none |
| `sw_getCalendarAppointments` | Returns active future appointments for one month. | calendar read model/cache or Master, AI brief index | none |
| `sw_getAppointmentAiBrief` | Returns AI brief for a root. | `_AppointmentArtifacts`, AI brief cache | none |
| `sw_getDiamondTrackingDashboard` | Returns diamond tracking dashboard rows/stats. | diamond read model/cache or 200 workbook | none |
| `sw_getInStockDiamonds` | Returns available in-stock diamonds for proposal planning/assignment. | diamond read model/cache or 200 workbook | none |
| `sw_getBulkReturnCandidates` | Returns return-eligible diamond rows. | diamond read model/cache or 200 workbook | none |
| `sw_adminGetTasks` | Returns admin task rows by filter. | `_SalesTaskQueue` | none |
| `sw_adminGetEmployeeSchedules` | Returns users, roster, schedule changes, config options. | `_SalesWorkflowUsers`, `10_Roster_Schedule`, `Schedule Changes` | none |
| `sw_adminAuditWorkflowPeopleData` | Reports identity/schedule/dropdown/owner data quality issues. | users, roster, schedule changes, dropdown, Master | none |
| `sw_reviewDiamondWorkflowSetup` | Reviews diamond workflow configuration/templates/setup. | config, templates, 200 setup | none |
| `sw_measureSalesWorkflowSpeed` | Benchmarks major dashboard endpoints. | many, read-only except caches/logs | cache/log side effects only |
| `sw_measureSalesWorkflowStartupSpeed` | Benchmarks startup/bootstrap path. | web app/bootstrap | cache/log side effects only |
| `sw_diagnoseTaskVisibilityForOwner` | Explains why tasks are or are not visible to an owner. | users/tasks | none |
| `sw_testSalesWorkflowDryRun` | Dry-run workflow test helper. | workflow sources | none |

### 5.3 Mutating Dashboard APIs

| Function | Current Role | Current Reads | Current Writes |
|---|---|---|---|
| `sw_completeTask` | Validates and completes a task, runs task-specific writeback, logs completion, then requests async queue generation through the orchestrator. | task row, user, template, task payload, task-specific source rows | `_SalesTaskQueue`, `_SalesTaskLog`, plus possible Master, 200, wax, artifact, cleanup writes, async generation request marker |
| `sw_acknowledgeTask` | Marks data as acknowledged by delegating to `sw_completeTask`. | same as completion | same as completion |
| `sw_snoozeTask` | Snoozes a task until a date with a reason. | task row, user | `_SalesTaskQueue`, `_SalesTaskLog` |
| `sw_claimTask` | Claims a coverage/shared task. | task row, user | `_SalesTaskQueue`, `_SalesTaskLog` |
| `sw_adminReassignTask` | Admin reassigns one task. | task row, user | `_SalesTaskQueue`, `_SalesTaskLog` |
| `sw_adminAssignAppointmentOwners` | Admin assigns Client Advisor/JOC owners for all Master rows sharing a root, then refreshes tasks. | task row, Master root rows, users | `00_Master Appointments`, `_SalesTaskQueue`, `_SalesTaskLog` |
| `sw_adminBlockTask` | Admin blocks task. | task row | `_SalesTaskQueue`, `_SalesTaskLog` |
| `sw_adminUnblockTask` | Admin restores blocked task to pending. | task row | `_SalesTaskQueue`, `_SalesTaskLog` |
| `sw_logTemplateCopied` | Records that a user copied a task template. | task row/user | `_SalesTaskLog` |
| `sw_logClientLoadTiming` | Logs browser load timing payload. | user/session | Apps Script logs only |
| `sw_adminSaveEmployeeSchedules` | Saves canonical schedule rows. | users/schedules | `10_Roster_Schedule` |
| `sw_adminUpsertScheduleChange` | Adds/updates one schedule override. | users/schedules | `Schedule Changes` |
| `sw_adminDeleteScheduleChange` | Deletes one schedule override. | users/schedules | `Schedule Changes` |
| `sw_adminMigrateWorkflowPeople` | Migrates workflow people identity data from legacy sources. | users, roster, dropdown, Master owners | users/roster/dropdown depending on options |
| `sw_bulkMarkDiamondsReturnInProgress` | Marks selected diamonds as return in progress. | 200 workbook selected rows | 200 workbook, task log/read-model invalidation |
| `sw_getAppointmentUploadFolder` | Ensures/returns Drive upload folder for a task/root/artifact type. | task, Master root folder, Drive | Drive folder/cache |
| `sw_syncAppointmentDriveUploads` | Registers files from Drive drop folders into artifact rows. | task, Drive folders/files | `_AppointmentArtifacts` |

## 6. Booking Intake And Appointment Creation Flow

### 6.1 New Acuity Booking

Current flow:

1. A verified relay receives Acuity `scheduled`, `rescheduled`, and `canceled` webhooks and appends rows to `_ExternalBookingEvents`.
2. `sw_backgroundOrchestrator` calls `sw_processExternalBookingEvents` before the polling fallback.
3. `sw_processExternalBookingEvents` processes pending Acuity rows only, fetches appointment detail by ID, and writes queue status/result metadata.
4. `scheduled` events submit the existing Google Form bridge when the Acuity UID is not already in Master/Form Inbox.
5. `rescheduled` events reconcile the existing Acuity UID in Master through `acuityHandleExisting_`, mark the old row inactive/rescheduled, and submit the synthetic `_R...` form row.
6. `canceled` events call `acuityCancelOnMaster_`.
7. If the webhook queue has no pending form-producing work, `sw_backgroundOrchestrator` calls `acuityPollAndSubmit` as the fallback.
8. `acuityPollAndSubmit` loads Acuity credentials and form ID from Script Properties.
9. It reads existing UIDs from `00_Master Appointments` using `acuityGetMasterUIDs_`.
10. It fetches active and canceled appointment lists from Acuity.
11. For active appointments not already present, it fetches detail with `acuityFetchAppointmentDetail_`.
12. `acuityToFormFieldMap_` normalizes Acuity fields into Google Form question titles.
13. `acuitySubmitToForm_` submits a Google Form response.
14. The Form submit trigger runs `onFormSubmit`.
15. `onFormSubmit` parses form named values, dedupes by UID/email/timestamp, resolves/reschedules/dedupes against Master, and appends or updates a Master row.
16. `onFormSubmit` creates or repairs folders/docs, stamps sales stage for consults, ensures root IDs, enqueues DV hooks, posts Chat notifications, and appends automation notes.
17. The orchestrator defers downstream work briefly when Acuity submitted forms so the form-triggered writes can drain.

Primary functions:

- `acuityPollAndSubmit`
- `sw_processExternalBookingEvents`
- `acuityFetchAppointmentLists_`
- `acuityFetchAppointmentDetail_`
- `acuityToFormFieldMap_`
- `acuitySubmitToForm_`
- `onFormSubmit`
- `findBestMasterRowByUID_`
- `findCurrentMasterRowByFingerprint_`
- `_findMostRecentPriorRow`
- `ensureArtifactsForRow`
- `repairMissingUrls_`

Current writes:

- Google Form response
- `_ExternalBookingEvents`
- `00_Master Appointments`
- Drive folders/docs
- `02_Form_Inbox` when the Form captures raw responses
- automation logs

### 6.2 Existing Acuity Booking Edit

Current flow:

1. Webhook-backed events are handled first through `sw_processExternalBookingEvents`.
2. The polling fallback rotates through existing active Acuity appointments before new active appointments, so reschedules/edits are reconciled before a current active Acuity row can be treated as a brand-new booking.
3. `acuityHandleExisting_` finds the Master row by `CalendlyEventUID`, preferring newest `_R...` reschedule row.
4. It compares Acuity date/time to Master date/time.
5. If date/time changed, the flow becomes a reschedule.
6. If contact/profile fields changed, it updates the existing Master row in place for fields such as `Phone`, `Diamond Type`, `Budget Range`, `Source`, and `Style Notes`.
7. It appends an "Edited via Acuity" automation note.

Current writes:

- `00_Master Appointments`
- Google Form response only if the edit is treated as a reschedule

### 6.3 Acuity Reschedule

Current flow:

1. `acuityHandleExisting_` detects date/time mismatch.
2. It marks the old Master row `Status = Rescheduled`, `Active? = No`, sets `CanceledAt`, and writes `RescheduledToUID`.
3. It builds a stable synthetic UID with `acuityStableRescheduleUid_`.
4. It submits a new Google Form response with that synthetic UID.
5. `onFormSubmit` creates or resolves the new Master row.
6. `onFormSubmit` inherits the old root and owners with `inheritOwnersForReschedule_`.
7. It sets the old row `RescheduledToUID` and the new row `RescheduledFromUID`.

Current writes:

- old row in `00_Master Appointments`
- new row in `00_Master Appointments`
- Google Form response
- automation notes

### 6.4 Acuity Cancellation And Label Sync

Cancellation:

- `acuityPollAndSubmit` checks canceled Acuity appointments.
- `acuityCancelOnMaster_` finds the active Master row by UID or latest `_R...` UID.
- It writes `Status = Canceled`, `Active? = No`, and automation notes.

Label sync:

- `acuityLabelSync` fetches future Acuity appointments.
- `labelToStatus_` maps labels like `completed`, `confirmed`, `no-show`, and `canceled`.
- `findMasterEntry_` resolves the Master row by UID or latest reschedule UID.
- It updates Master `Status` and automation notes.

Current writes:

- `00_Master Appointments`

## 7. Background Orchestration Flow

Primary function:

- `sw_backgroundOrchestrator`

Current job order:

1. Acquire script/property lease.
2. Check intake drain state so form submissions are not overlapped by dependent jobs.
3. Run `sw_processExternalBookingEvents`.
4. If webhook-backed Acuity processing submitted forms/reschedules, mark intake drain and exit early.
5. Run `acuityPollAndSubmit`.
6. If Acuity polling submitted forms/reschedules, mark intake drain and exit early.
7. Run `acuityLabelSync`.
8. Run `sw_processAppointmentAutomation`.
9. Run `sw_generateSalesWorkflowTasks` if a pending async task-generation request exists, or if automation did not already generate tasks and hourly cadence is due.
10. Run `repairMissingUrls_` hourly.
11. Run `sw_rebuildWorkflowReadModels` when read models are due.
12. Save state and release lease.

Install/status functions:

- `sw_installBackgroundOrchestratorTrigger`
- `sw_removeBackgroundWorkerTriggers_`
- `sw_removeBackgroundOrchestratorTrigger`
- `sw_getBackgroundOrchestratorStatus`
- `sw_clearBackgroundOrchestratorLease`

Current protection behavior:

- Uses a lease to prevent overlap.
- Removes retired background worker triggers so Acuity, label sync, automation, generation, repair, and read-model rebuilds run through one orchestrator.
- Defers downstream work after intake submissions to avoid racing `onFormSubmit`.
- Honors async task-generation requests from task completion on the next safe orchestrator run, while preserving the manual synchronous Refresh Queue path.

## 8. Task Generation Flow

Primary function:

- `sw_generateSalesWorkflowTasks`

Current flow:

1. `sw_setupSalesWorkflow` ensures infrastructure sheets.
2. `swBuildContext_` reads config, templates, canonical workflow people, roster/schedule data, wax index, and supporting indexes.
3. `swReadAppointments_` reads selected columns from `00_Master Appointments`.
4. `swPrepareClientAdvisorRoundRobin_` prepares auto-assignment load data.
5. `swReadTaskState_` reads `_SalesTaskQueue`.
6. For each appointment:
   - skip rows with no root/appt;
   - skip old/out-of-window rows with `swIsWorkflowRelevant_`;
   - block tasks when `swIsAppointmentActive_` is false;
   - maybe auto-assign Client Advisor/JOC to Master;
   - generate core appointment tasks;
   - generate diamond tasks;
   - generate post-consult tasks;
   - generate data cleanup tasks.
7. Deferred task writes are flushed to `_SalesTaskQueue` and `_SalesTaskLog`.

Core task generation functions:

- `swGenerateTasksForAppointment_`
- `swBuildTask_`
- `swUpsertTask_`
- `swBlockTasksForAppointment_`
- `swResolveOwner_`
- `swResolveJocOwner_`
- `swMaybeAutoAssignClientAdvisor_`
- `swWriteAppointmentOwnerAssignmentToMaster_`
- `swGenerateDiamondWorkflowTasks_`
- `swGeneratePostConsultTasks_`
- `swGenerateDataCleanupTasks_`

Core appointment task sequence:

| Task | Owner | Created When |
|---|---|---|
| `ASSIGN_APPOINTMENT` | System | every relevant appointment; auto-completed |
| `SEND_HYBRID_WELCOME` | JOC | appointment is within the hybrid window |
| `SEND_WELCOME` | JOC | appointment is farther out |
| `SEND_MAP_INSTRUCTIONS` | JOC | appointment has visit time, due 48h before |
| `REVIEW_APPOINTMENT` | Client Advisor | appointment has visit time, due 24h before |
| `APPOINTMENT_DAY_CHECKLIST` | Client Advisor | appointment day |
| `APPROVE_RECAP_MESSAGE` | Client Advisor | checklist complete, appointment not no-show, AI summary ready |
| `SEND_FINAL_RECAP` | JOC | recap approved |

Post-consult task sequence:

| Task | Owner | Created When |
|---|---|---|
| `POST_CONSULT_CLIENT_STATUS` | JOC | appointment day checklist is completed |
| `START_3D_DESIGN` | JOC | status complete, 3D needed, SO not present |
| `RECORD_3D_DEADLINE` | JOC | 3D started and no deadline present |
| `REQUEST_WAX_PRINT` | JOC | status/start says wax needed and no active wax request |
| `UPDATE_WAX_REQUEST` | JOC | wax request needs update |

Diamond task sequence:

| Task | Owner | Created When |
|---|---|---|
| `PROPOSE_DIAMONDS` | Client Advisor | diamond viewing workflow active |
| `PREPARE_DV_QUOTATION` | JOC | diamond viewing workflow active |
| `ORDER_DIAMONDS` | Diamond Order Admin | proposed/orderable stones need order review |
| `TRACK_DIAMONDS` | Diamond Order Assistant | ordered stones need tracking |
| `CONFIRM_DIAMOND_DELIVERY` | Diamond Order Admin | delivered confirmation is needed |
| `ACK_DIAMONDS_ORDERED_ASSIGNED_REP` | Client Advisor | ordered acknowledgement is needed |
| `ACK_DIAMONDS_ORDERED_JOC` | JOC | ordered acknowledgement is needed |
| `RECORD_DIAMOND_DECISIONS` | JOC | decisions are due |
| `RETURN_DIAMONDS` | Diamond Order Assistant | return is due |
| `REVIEW_DIAMOND_ETA_ASSIGNED_REP` | Client Advisor | ETA risk exists |
| `REVIEW_DIAMOND_ETA_JOC` | JOC | ETA risk exists |

## 9. Task Detail And Completion Flow

### 9.1 Task Detail

Current flow:

1. UI calls `sw_getTaskDetail`.
2. Server authenticates the user and checks access.
3. It loads the task via cache/index/read model/queue.
4. It reads the template for the task type.
5. It renders template data with `swRenderDataForTask_`.
6. It renders copyable text with `swRenderedCopyableTemplateForTask_`.
7. It builds attachments with `swAttachmentsForTask_`.
8. It checks missing fields with `swMissingFieldsForTask_`.
9. For appointment checklist tasks, it may include appointment artifacts and upload folder links.
10. For recap approval/final tasks, it includes AI brief data.

Key functions:

- `sw_getTaskDetail`
- `swGetTaskById_`
- `swRenderDataForTask_`
- `swRenderedCopyableTemplateForTask_`
- `swAttachmentsForTask_`
- `swMissingFieldsForTask_`
- `swTaskDetailAppointmentAiBrief_`
- `swPublicAppointmentArtifacts_`
- `swCachedAppointmentUploadFoldersForTask_`

### 9.2 Task Completion

Current flow:

1. UI calls `sw_completeTask`.
2. Server authenticates the user and checks `swCanActOnTask_`.
3. Server rejects non-pending/non-due-snoozed tasks.
4. `swValidateCompletion_` validates general and task-specific required fields.
5. Task-specific adapters run before the task row is completed:
   - `swDiamondHandleTaskCompletion_`
   - `swHandlePostConsultTaskCompletion_`
   - `swHandleDataCleanupTaskCompletion_`
   - `swHandleAppointmentCompletion_`
   - `swMarkAppointmentSummaryApproved_`
   - `swMarkAppointmentJocHandoff_`
6. The task payload stores completion data, rendered template, attachments, actor, timestamp, and adapter results.
7. Task row is marked `Completed`.
8. A `COMPLETE` event is appended to `_SalesTaskLog`.
9. Read models are invalidated for appointment/source changes.
10. `sw_requestSalesWorkflowTaskGenerationAfterSubmit_` records a bounded async generation request for the background orchestrator. If `SW_COMPLETE_TASK_ASYNC_GENERATION=N`, or the request cannot be recorded, it falls back to `sw_tryGenerateSalesWorkflowTasksAfterSubmit_`.

Current writes:

- `_SalesTaskQueue`
- `_SalesTaskLog`
- task-specific source writes listed below

Task-specific writeback summary:

| Task Family | Adapter | Current Writes |
|---|---|---|
| Appointment checklist/no-show | `swHandleAppointmentCompletion_` | Master appointment outcome fields, `_AppointmentArtifacts` state |
| Approve recap | `swMarkAppointmentSummaryApproved_` | `_AppointmentArtifacts` summary approval fields |
| Send final recap | `swMarkAppointmentJocHandoff_` | `_AppointmentArtifacts` handoff fields |
| Client status | `swCompleteClientStatusTask_` -> `cs_submitFromDialog` | Master status fields, client status report/log |
| Start 3D | `swCompleteStart3DTask_` -> `saveAssignedSO` | Master SO/3D/order fields, order folders, tracker log |
| 3D deadline | `swComplete3DDeadlineTask_` -> `saveRecordDeadline` | Master 3D deadline/deadline move count |
| Wax request | `swCompleteWaxRequestTask_` -> `wax_onRequestSubmit_` | `05_Wax_Requests`, Master wax mirror |
| Wax update | `swCompleteWaxUpdateTask_` -> `wax_adminCommitFromDialog` | `05_Wax_Requests`, Master wax mirror |
| Diamond proposal/order/tracking/delivery/decision/return | `swDiamondHandleTaskCompletion_` | 200 workbook, Master DV requirement fields when proposal requirements are captured |
| Data cleanup | `swHandleDataCleanupTaskCompletion_` | `_SalesDataCleanup`, Master if admin confirmation applies proposal |

## 10. Customer Search And Customer Detail Flow

Primary file:

- `sales_workflow_customer_search.js`

Read flow:

1. UI calls `sw_searchCustomers`.
2. Server checks user role.
3. Filters are normalized with `swCustomerSearchNormalizeFilters_`.
4. If allowed and fresh, rows are read from `_SW_CustomerReadModel` cache/sheet.
5. Otherwise rows fall back to `swReadAppointments_` from Master.
6. Default owner filters are applied for Client Advisor/JOC users.
7. `swCustomerSearchKanbanFromRows_` groups rows by root and stage.
8. Cards are built with `swCustomerSearchCardFromReadModel_` or `swCustomerSearchCard_`.
9. The result includes list rows, Kanban columns, filters, filter options, source, and timing.

Detail flow:

1. UI calls `sw_getCustomerSearchDetail`.
2. Server resolves the root through read-model detail cache or Master root rows.
3. It builds a detail card, appointments list, payment history, recent logs, form options, AI brief, and action context.
4. It returns the detail payload to the customer drawer.

Mutating actions:

| Function | Current Purpose | Current Writes |
|---|---|---|
| `sw_customerSearchUpdateStatus` | Updates client status from customer detail. | Master via `cs_submitClientStatusUpdate_`/related helpers, status log, customer read-model invalidation |
| `sw_customerSearchUpdate3DDeadline` | Updates 3D deadline from customer detail. | Master deadline fields, status/log, customer read-model invalidation |
| `sw_customerSearchSubmit3DRevision` | Submits 3D revision from customer detail. | 3D tracker log and Master/order fields as applicable |
| `sw_customerSearchRequestWax` | Creates wax request from customer detail. | `05_Wax_Requests`, Master wax mirror, customer read-model invalidation |

Important caches:

- customer search read-model row cache;
- customer initial payload cache;
- customer detail shard cache;
- payment history index cache;
- recent task log index cache.

## 11. Admin Dashboard Flow

Primary file:

- `sales_workflow_admin_dashboard.js`

Current flow:

1. UI calls `sw_getAdminDashboard`.
2. Server checks admin access.
3. Filters/window presets are normalized.
4. If eligible and fresh, `_SW_AdminDashboardReadModel` is used.
5. Otherwise the fallback builds payload from:
   - Master appointments;
   - task state;
   - payment receipts;
   - root index;
   - status log;
   - config/stage weights.
6. The response includes metrics, health, pipeline, lead sources, advisor scorecard, top deals, receivables, and filter options.

Key functions:

- `swAdminDashboardBuildPayload_`
- `swAdminDashboardMetrics_`
- `swAdminDashboardHealthContext_`
- `swAdminDashboardHealth_`
- `swAdminDashboardPipelineStage_`
- `swAdminDashboardCustomerCard_`
- `swAdminDashboardReadPayments_`
- `swAdminDashboardReadRootIndex_`
- `swAdminDashboardReadStatusLog_`

Current writes:

- none from dashboard read itself;
- payment/root/status caches may be updated.

## 12. Calendar Flow

Current flow:

1. UI calls `sw_getCalendarAppointments(monthKey)`.
2. Server authenticates user.
3. If fresh, it returns `_SW_CalendarMonthReadModel` data.
4. Otherwise it scans Master appointment rows.
5. Rows are filtered to active, future, matching month appointments.
6. The payload includes appointment ID, root, customer, brand, date/time, owners, status, folder/report/quote/tracker links, DV flag, and compact AI brief state.

Current writes:

- none.

Related AI brief function:

- `sw_getAppointmentAiBrief` reads summary/artifact data from `_AppointmentArtifacts`.

## 13. Client Status, 3D Deadline, Start 3D, And Revision Dialogs

### 13.1 Client Status

Primary functions:

- `cs_openStatusDialog_`
- `cs_submitFromDialog`
- `cs_submitFromDialogForRow_`
- `cs_submitClientStatusUpdate_`
- `cs_createOrGetReportForSelection_`
- `cs_ensureReportUrl_`
- `cs_automationSubmit_`

Current flow:

1. Dialog resolves active row or explicit row.
2. It reads dropdowns and validation rules.
3. It creates or opens the client status report workbook if needed.
4. It writes selected status fields to Master.
5. It updates the client report snapshot/log.
6. It appends audit/status log rows.

Current Master-owned facts written:

- Sales Stage
- Conversion Status
- Custom Order Status
- In Production Status
- Center Stone Status
- Next Steps
- Order Date
- 3D Deadline and production deadline fields when included
- report URL/related snapshot fields

### 13.2 3D Deadline

Primary functions:

- `showRecordDeadlineDialog`
- `getRecordDeadlineInit`
- `saveRecordDeadline`

Current flow:

1. Dialog reads active Master row context.
2. User selects deadline kind/date.
3. `saveRecordDeadline` writes deadline fields on Master.
4. It increments deadline move count where appropriate.

Current writes:

- `00_Master Appointments`

### 13.3 Start 3D / Assign SO

Primary functions:

- `start3d_init`
- `previewOdooPaste`
- `start3d_step2Payload`
- `getActiveMasterPreview`
- `checkSOConflicts`
- `saveAssignedSO`
- `propagateSOToSiblingRows_`
- `append3DTrackerLog_`

Current flow:

1. Dialog initializes from active Master row.
2. User enters SO/Odoo/design fields.
3. `checkSOConflicts` validates against Master/order data.
4. `saveAssignedSO` writes SO/order/3D fields to Master.
5. It creates/links order/client folders and shortcuts.
6. It propagates SO fields to sibling Master rows when allowed.
7. It appends a tracker log.

Current writes:

- `00_Master Appointments`
- Drive folders/shortcuts
- 3D tracker workbook/log

### 13.4 3D Revision

Primary functions:

- `open3DRevision`
- `rev3d_init`
- `previewRevOdooPaste`
- `submit3DRevision`

Current flow:

1. Dialog loads active Master row/order/tracker context.
2. User enters revision form.
3. Revision payload is appended to tracker/log.
4. Master/order status fields may be updated as applicable.

Current writes:

- 3D tracker workbook/log
- `00_Master Appointments` where applicable

## 14. Wax Flow

Primary file:

- `WaxRequests.js`

Primary functions:

- `wax_ensureSheet_`
- `wax_statusOptions`
- `wax_onRequestSubmit_`
- `wax_adminGetPendingData`
- `wax_adminCommitFromDialog`
- `wax_mirrorToMaster_`
- `wax_getOrCreateFolder_`
- `wax_recomputeMetricsForRow_`

Current flow:

1. A dashboard/customer detail/task action calls `wax_onRequestSubmit_`, or an admin opens the pending wax dialog.
2. `wax_ensureSheet_` ensures `05_Wax_Requests`.
3. New request rows are appended with root/SO/date/priority/requester context.
4. Wax folders may be created from the 3D/order folder context.
5. Wax status/admin deadline/request URL are mirrored to Master.
6. Admin update dialog commits status/deadline updates and mirrors the current status back to Master.

Current writes:

- `05_Wax_Requests`
- `00_Master Appointments`
- Drive folders

## 15. Diamond Workflow Flow

Primary files:

- `sales_workflow_diamonds.js`
- `sales_workflow_diamond_sync.js`
- diamond dialog files

Current dashboard/read functions:

- `sw_getInStockDiamonds`
- `sw_getDiamondTrackingDashboard`
- `sw_getBulkReturnCandidates`
- `sw_refreshDiamondQuoteFromTracking`
- `sw_refreshDiamondQuoteFrom3D`
- `sw_refreshDiamondQuoteAll`

Current mutating functions:

- `sw_assignInStockDiamond`
- `sw_bulkMarkDiamondsReturnInProgress`
- `sw_previewLoupe360DiamondSync`
- `sw_applyLoupe360DiamondSync`
- `dp_submitProposals`
- `dp_submitOrderApprovals`
- `dp_submitConfirmDelivery`
- `dp_submitStoneDecisions`
- diamond task completion adapters in `sales_workflow_diamonds.js`

Current flow:

1. Task generation detects diamond viewing appointments from Master.
2. Diamond tasks are generated using Master appointment payload and 200 workbook rows.
3. Proposal workspace captures customer requirements and proposed stones.
4. Requirements are written to Master DV columns.
5. Per-stone proposal/order/tracking/delivery/decision/return fields are written to the 200 workbook.
6. Read models mirror dashboard fields into `_SW_DiamondReadModel` and `_SW_DiamondRootReadModel`.
7. Quote refresh writes selected diamond rows/settings into quotation workbooks.

Current writes:

- 200 workbook
- Master DV customer requirement columns
- quote workbook
- task queue/log
- read-model invalidation

## 16. Payment Flow

Primary files:

- `Payments_v1.js`
- `Payment_Summary_v1.js`
- `dlg_record_payment_v1.html`

Primary functions:

- `openRecordPayment`
- `rp_init`
- `rp_validateDocTypePrerequisite`
- `rp_checkHasSalesInvoice`
- `rp_listDocNumbersForAnchor`
- `rp_getLatest3DFields`
- `rp_submit`
- `rp_makeDocForPayment`
- `rp_resetFromDialog`
- `rp_getLedgerTarget`
- `rp_generateDocAndPdf_`
- `rp_updateLedgerRow_`
- `rp_persistSavedLinesToMaster_`
- `rp_applyReceiptToMaster`
- `rp_updateMasterCashInGross_`
- `rp_setSalesStageOnMaster_`
- `rp_applyReferralToClientStatus_`
- `ps_init`
- `ps_fetchHistoryForAnchor_`
- `ps_exportPdf`

Current flow:

1. Payment dialog opens from active Master row, dashboard action, or iPad app context.
2. `rp_init` resolves appointment/SO/customer context, tax rate, payment history, saved lines, and folder targets.
3. User validates invoice/receipt prerequisites.
4. `rp_submit` appends or updates the payment ledger row.
5. `rp_makeDocForPayment` creates invoice/receipt docs and PDFs.
6. Ledger row is updated with generated document links.
7. Appointment receipts update Master paid/cash-in/payment summary fields.
8. Order total and saved quote lines may be persisted to Master.
9. Sales stage and referral snapshot may be updated.
10. Payment read models are invalidated after payment writes.

Current writes:

- payment ledger workbook
- Drive docs/PDFs/folders
- `00_Master Appointments`
- client status report snapshot for referral fields

## 17. Appointment Artifacts And AI Summary Flow

Primary file:

- `sales_workflow_appointment_artifacts.js`

Primary functions:

- `swEnsureAppointmentArtifactsSheet_`
- `sw_uploadAppointmentArtifacts`
- `sw_getAppointmentUploadFolder`
- `sw_syncAppointmentDriveUploads`
- `sw_ingestRawAppointmentUpload_`
- `sw_processAppointmentAutomation`
- `swProcessAppointmentArtifact_`
- `swStartAssemblyTranscription_`
- `swPollAssemblyTranscription_`
- `swSaveTranscriptAndQueueSummary_`
- `swGenerateAppointmentSummary_`
- `swAppointmentAiBriefForRoot_`
- `swAppointmentSummaryIndex_`
- `swHandleAppointmentCompletion_`
- `swWriteAppointmentOutcomeToMaster_`

Current flow:

1. Appointment checklist requires recording upload when appointment outcome is completed.
2. Upload can happen via dashboard form, Drive drop folder sync, or raw upload endpoint.
3. Artifact rows are appended to `_AppointmentArtifacts`.
4. Background automation processes ready artifacts.
5. Audio/video files are shared or uploaded for AssemblyAI transcription.
6. Transcript text is saved to Drive and artifact row.
7. OpenAI summary/follow-up draft is generated.
8. Transcript doc, summary doc, summary JSON, review flags, sales brief, and client follow-up draft are stored as artifact metadata.
9. Task generation creates recap approval when checklist is complete and summary is ready.
10. Approve/final recap tasks mark approval/handoff fields on artifacts.

Current writes:

- `_AppointmentArtifacts`
- Drive docs/files
- `00_Master Appointments` appointment outcome fields
- read-model/cache invalidation

## 18. Read Models And Caches

Primary files:

- `sales_workflow_read_model.js`
- `sales_workflow_infrastructure_read_models.js`
- `sales_workflow_task_repository.js`
- `sales_workflow_customer_search.js`

Primary functions:

- `sw_rebuildWorkflowReadModels`
- `sw_getWorkflowReadModelStatus`
- `sw_invalidateWorkflowReadModels`
- `swMarkWorkflowReadModelsStale_`
- `swBuildTaskReadModel_`
- `swBuildCustomerReadModel_`
- `swBuildDiamondReadModels_`
- `swBuildAppointmentReadModels_`
- `swBuildPaymentReadModel_`
- `swBuildAdminDashboardReadModel_`
- `swTryReadTaskListStateFromReadModel_`
- `swTryGetCalendarAppointmentsFromReadModel_`
- `swTryGetInStockDiamondsFromReadModel_`
- `swTryGetDiamondTrackingDashboardFromReadModel_`
- `swTryReadAdminDashboardFromReadModel_`

Current flow:

1. Orchestrator or manual call runs `sw_rebuildWorkflowReadModels`.
2. Builder reads current source sheets:
   - Master for task/customer/appointment/calendar/admin inputs;
   - `_SalesTaskQueue` for task read model;
   - 200 workbook for diamond read models;
   - payment ledger for payment read model;
   - status/root/task auxiliaries for admin dashboard.
3. Hidden `_SW_*ReadModel` tabs are written.
4. Metadata is written to `_SW_ReadModelMeta`.
5. CacheService entries are prewarmed for hot paths.
6. Dashboard endpoints use fresh read models first and fall back to source sheets when stale/missing.

Current invalidation behavior:

- task queue writes invalidate task list/detail caches and task dashboard projections;
- customer search writes invalidate customer read model caches and detail caches;
- diamond writes invalidate diamond read model caches;
- appointment writes invalidate appointment/calendar/admin caches;
- payment writes invalidate payment/admin caches.

## 19. Identity, Users, Schedules, And Ownership

Primary functions:

- `sw_adminListWorkflowUsers`
- `sw_adminUpsertWorkflowUser`
- `sw_adminGetEmployeeSchedules`
- `sw_adminSaveEmployeeSchedules`
- `sw_adminUpsertScheduleChange`
- `sw_adminDeleteScheduleChange`
- `swReadCanonicalWorkflowPeople_`
- `swCanonicalWorkflowPeopleIndex_`
- `swReadEmployeeScheduleAdminData_`
- `swReadRosterAvailabilityIndex_`
- `swReadScheduleChangesIndex_`
- `swResolveOwner_`
- `swResolveJocOwner_`
- `swAvailabilityFor_`
- `swScheduleOverride_`

Current flow:

1. Users are stored in `_SalesWorkflowUsers`.
2. Schedulable user extension data is stored in `10_Roster_Schedule`.
3. One-off availability overrides are stored in `Schedule Changes`.
4. Task generation resolves Client Advisor and JOC owners from Master appointment fields against active canonical users.
5. JOC coverage tries intended JOC, coverage partner, available JOC, then shared coverage queue.
6. Admin owner assignment writes selected owners back to all Master rows for the root and refreshes tasks.

Current writes:

- `_SalesWorkflowUsers`
- `10_Roster_Schedule`
- `Schedule Changes`
- `00_Master Appointments` owner fields
- task queue/log after refresh

## 20. Data Cleanup Campaign Flow

Primary file:

- `sales_workflow_data_cleanup.js`

Current flow:

1. Task generation identifies stale/dirty customer rows.
2. Review tasks are created for Client Advisor and/or JOC.
3. User submits proposed cleanup.
4. Admin confirmation task is created.
5. Admin applies or returns changes.
6. Applied changes write back to Master and update cleanup case state.

Primary task types:

- `CUSTOMER_DATA_CLEANUP_REVIEW`
- `CUSTOMER_DATA_CLEANUP_CONFIRM`
- `CUSTOMER_DATA_CLEANUP_REVISE`

Current writes:

- `_SalesDataCleanup`
- `_SalesTaskQueue`
- `_SalesTaskLog`
- `00_Master Appointments` when applied

## 21. Reminders Flow

Primary file:

- `Reminders_v1.js`

Current role:

- Maintains reminder queue/log behavior for follow-up reminders, COS reminders, DV reminders, snooze/cancel actions, and time-trigger sending.

Primary functions:

- `remind__installTimeTriggers`
- `remind__sendNowForTesting`
- `remind_menu_snoozeSelected`
- `remind_menu_cancelSelected`
- `remind__cancelTarget_do`
- `remind__snoozeTarget_do`
- `remind__getActiveSummaryForTarget`
- `remind__dedupeActiveCosRemindersPreview`
- `remind__dedupeActiveCosReminders`

Current writes:

- reminder queue/log tabs
- selected reminder rows

Current reads:

- Master/customer/order fields;
- reminder queue/log tabs;
- sometimes DV/order state.

## 22. External Reports, Uploads, And Ask Controller

Primary files:

- `01_UploadEndpoint.js`
- `02_Workers.js`
- `03_SummaryRenderer.js`
- `04_AskController.js`
- `05_AskCore.js`

Current role:

- These support appointment summary uploads, transcript/scribe/strategist processing, client summary tabs, and Ask Controller chat/patch workflows.

Important functions:

| Function | Role |
|---|---|
| `doPost_UPLOAD_` | Raw upload endpoint for appointment files. |
| `processUploadQueue` | Legacy upload queue processor. |
| `summarizeLatestTranscript` | Summarizes latest transcript for a root. |
| `runStrategistAnalysisForRoot` | Runs strategist analysis for one root. |
| `upsertSYSConsults_` | Upserts consult summary rows. |
| `mirrorSummaryToMaster_` | Mirrors selected summary output back to Master. |
| `upsertClientSummaryTab_` | Creates/updates client summary report tab. |
| `rerenderClientSummaryTabForRoot_` | Rerenders an existing client summary tab. |
| `askControllerDoGet_` | Serves Ask Controller UI. |
| `AC_doPost_` | Handles Ask Controller API posts. |
| `AC_chatCore_` | Runs chat against latest artifacts. |
| `AC_applyPatchCore_` | Applies structured patch flow. |

Current reads/writes:

- reads Master row/folder/report fields by `RootApptID`;
- reads/writes Drive artifacts;
- writes summary/report outputs;
- may mirror summary fields back to Master.

## 23. Current Master Write Hotspots

The following current functions are important migration touch points because they write operational facts into `00_Master Appointments`:

| Function/Flow | Current Master Writes |
|---|---|
| `onFormSubmit` | new appointment rows, reschedule links, root IDs, contact/profile fields, status/active/booked/visit fields |
| `acuityHandleExisting_` | phone, diamond type, budget, source, style notes, reschedule status/links |
| `acuityCancelOnMaster_` | cancellation status and active state |
| `acuityLabelSync` | appointment status from Acuity labels |
| `swWriteAppointmentOwnerAssignmentToMaster_` | auto-assigned Client Advisor/JOC |
| `sw_adminAssignAppointmentOwners` | admin owner assignment for root rows |
| `cs_submitFromDialog` / `cs_submitClientStatusUpdate_` | client status, next steps, deadlines, report snapshot fields |
| `saveRecordDeadline` | 3D/production deadline fields and move counts |
| `saveAssignedSO` | SO/Odoo/order/3D tracker/folder/design fields |
| `submit3DRevision` | 3D revision/order status fields where applicable |
| `wax_mirrorToMaster_` | wax status/admin deadline/request URL |
| `rp_submit` and payment helpers | order total, saved lines, cash-in, paid/balance, sales stage, referral snapshot |
| `swDiamondWriteCustomerRequirements_` | DV customer requirements and strategy fields |
| `swHandleAppointmentCompletion_` / `swWriteAppointmentOutcomeToMaster_` | appointment outcome/no-show/completed state |
| data cleanup confirmation | selected customer/profile/status corrections |
| artifact/report helpers | folder/report/transcript/summary URL compatibility fields |
| `repairMissingUrls_` | missing folder/doc/report URL repair |

## 24. Migration-Sensitive Current Read Hotspots

These functions currently rely on Master reads and need adapters or projections during migration:

| Function/Flow | Why It Matters |
|---|---|
| `swReadAppointments_` | Central source for task generation, customer read model, admin dashboard fallback, calendar fallback. |
| `swReadAppointmentsForRoot_` | Root-targeted customer/detail/task context read. |
| `swAppointmentRecordForRoot_` | Artifact/AI summary context by root. |
| `sw_getCalendarAppointments` fallback | Calendar active appointment set. |
| `sw_searchCustomers` fallback | Customer search/card source. |
| `sw_getCustomerSearchDetail` fallback | Detailed customer drawer. |
| `sw_getAdminDashboard` fallback | Metrics/health/pipeline source. |
| `rp_findMasterRowByRootApptId_` | Payment context and writeback target. |
| `cs_resolveRow_` | Client status active row resolution. |
| `start3d_init` / `getActiveMasterPreview` | Start 3D selected row context. |
| `onFormSubmit` dedupe/root resolution helpers | Appointment/root creation and reschedule detection. |
| `swGenerateSalesWorkflowTasksUnlocked_` | All workflow task generation depends on appointment records. |

## 25. Minimum Smoke Tests For Current Flow Preservation

Run these after any infrastructure change:

- Acuity new booking creates/resolves a row and appears in Calendar.
- Acuity reschedule keeps the same root and marks old appointment inactive.
- Acuity cancellation/label sync updates appointment status.
- Login/bootstrap loads My Queue and visible views.
- Queue refresh generates expected tasks from a known appointment.
- Open task detail for core, post-consult, diamond, and appointment checklist tasks.
- Complete appointment checklist with recording requirement behavior.
- AI summary readiness creates recap approval task.
- Complete client status task.
- Complete Start 3D task.
- Record 3D deadline.
- Submit customer search status update.
- Submit customer search 3D deadline update.
- Submit 3D revision.
- Submit wax request and admin wax update.
- Customer Search list and detail show updated facts.
- Calendar month shows active appointment set.
- Admin Dashboard metrics/pipeline load.
- Record payment, create doc/PDF, and verify ledger/Master rollups.
- Diamond proposal/order/tracking/delivery/decision/return flows update 200 and dashboards.
- Bulk return updates selected rows.
- Owner assignment updates task routing.
- Schedule change affects assignment on next queue refresh.
- Appointment upload folder and Drive sync register artifacts.
- Read-model rebuild completes and dashboard endpoints use fresh models.
