# Sales Workflow Domain Data Infrastructure PRD

## 1. Purpose

Rebuild the Sales Workflow data infrastructure so customer, appointment, order, task, diamond, payment, artifact, and dashboard data are stored in smaller domain-owned tabs keyed by `RootApptID`.

The goal is to reduce full-sheet scans, shorten write locks, make dashboard reads faster, and lower drift risk while preserving all existing user workflows during migration.

## 2. Background

The current workflow relies heavily on `00_Master Appointments` as a wide operational table. Many functions read or write overlapping slices of that tab for unrelated workflows:

- appointment lifecycle and current appointment state;
- customer identity and contact information;
- client status and next-step fields;
- 3D start, deadline, revisions, and tracker references;
- diamond viewing summaries;
- payment rollups;
- owner assignment;
- external booking intake from Acuity/Calendly, currently resolved through form submit and Master rows;
- dashboard/customer-search/card fields.

This creates several problems:

- unrelated workflows contend on one large tab;
- small updates often require reading broad appointment records;
- dashboard views repeat similar customer slices;
- derived facts can drift from source facts;
- it is hard to know which UI views need refresh after a write.

The new infrastructure separates canonical facts by domain while keeping `00_Master Appointments` as a derived compatibility projection during migration.

## 3. Goals

- Store each canonical fact in exactly one source tab.
- Key customer-level domains by `RootApptID`.
- Preserve appointment-level event history by `APPT_ID` and link each event to its `RootApptID`.
- Serve dashboards through small read slices instead of broad source scans.
- Allow targeted reads and targeted writes for common operations.
- Keep user workflows safe when multiple people update the same customer in different workflow areas.
- Keep existing UI and function flows working throughout migration.
- Rewire external booking intake so Acuity/Calendly/Form submissions resolve into domain tabs instead of canonical Master writes.
- Add test gates after every migration phase so data-flow breakages are caught early.

## 4. Non-Goals

- Do not replace Google Sheets or Apps Script in this migration.
- Do not remove existing user-facing workflow features.
- Do not rewrite the dashboard UI as part of the data migration unless a small adapter change is required.
- Do not remove the existing Google Form/Form Inbox intake path until the new booking adapter is proven against live Acuity/Calendly cases.
- Do not move external client workbooks, quotation workbooks, 3D tracker workbooks, Drive folders, or the 200 stones workbook into the 100 workbook.
- Do not make `00_Master Appointments` canonical after migration. It remains compatibility output only.

## 5. Users And Workflows

Primary users:

- Client Advisors updating customer status, 3D deadlines, diamond requirements, payments, and cleanup items.
- JOCs processing appointments, recaps, 3D starts, wax requests, and coverage queues.
- Diamond order users proposing, ordering, tracking, delivering, deciding, and returning diamonds.
- Admins reviewing dashboards, users, schedules, assignments, health, and pipeline state.

Concurrent workflow examples that must stay safe:

- one user starts 3D while another updates client status;
- one user records a 3D deadline while another snoozes the deadline task;
- one user records payment while another opens the customer detail drawer;
- one user updates diamond viewing requirements while another proposes stones;
- admin changes appointment owner while task generation refreshes queues.

## 6. Product Requirements

### 6.1 Canonical Ownership

Every fact must have one canonical owner.

- Canonical tabs are the only write targets for their facts.
- Read models, caches, dashboard slices, and `00_Master Appointments` may duplicate facts only as derived projections.
- Derived projections must be rebuildable from canonical sources.
- Mutation functions must write through domain adapters, not directly to projection tabs.

### 6.2 Root-Keyed Data Model

All customer-level domains must be keyed by `RootApptID`.

Appointment event rows must include both:

- `APPT_ID`, the event/appointment identifier;
- `RootApptID`, the customer/root grouping identifier.

Domain tabs should include:

- `RootApptID`;
- `Version`;
- `Updated At`;
- `Updated By`;
- `Updated By Email`;
- domain-specific fields.

History tabs should be append-only unless a repair script is explicitly run.

### 6.3 Short Critical Sections

Apps Script has spreadsheet-level locking behavior rather than true per-tab locks, so this migration must implement short logical locks:

- read broad context outside the lock when safe;
- acquire lock only for version check and write;
- write the smallest target range possible;
- release lock immediately;
- mark affected read slices stale after the canonical write;
- rebuild projections outside user-facing write locks.

Optimistic version checks are required on root-owned tabs where concurrent edits are likely.

### 6.4 Dashboard Read Slices

Dashboards must read from purpose-built slices first, then fall back to targeted canonical reads if projections are stale.

Shared slices:

| Slice | Primary UI | Purpose |
|---|---|---|
| `TaskListSlice` | My Queue, Cleanup, Coverage, Admin Review | small task cards and counts |
| `TaskDetailSlice` | task drawer | task row plus task-specific root mini data |
| `CustomerCardSlice` | Customer Search, Kanban, Admin Pipeline | list/card fields only |
| `CustomerRootDetailSlice` | customer detail drawer, calendar detail expansion, admin pipeline detail, task customer panel | shared full root detail |
| `CalendarMonthSlice` | Calendar | month appointment cards |
| `AppointmentBriefSlice` | Calendar event detail and AI brief | appointment links, notes, artifact brief |
| `AdminHealthSlice` | Admin Dashboard | metrics, health, receivables, pipeline summaries |
| `DiamondInventorySlice` | In-Stock Diamonds | available stones |
| `DiamondTrackingSlice` | Diamond Tracking, Bulk Returns | ETA, return, tracking state |
| `PaymentSummarySlice` | customer detail, payment dialog, receivables | paid-to-date, balance, payment history |
| `FormOptionsSlice` | action forms and dialogs | dropdown/config options |

The repeated detailed customer panel must use one shared root detail service. Customer Search detail, Calendar expanded customer detail, Admin Pipeline detail, and task customer panels must not each assemble different full customer payloads.

### 6.5 Drift Control

Each write must declare which derived slices become stale.

An invalidation rule means: when a canonical source changes, mark the affected read models/caches/projections as no longer trusted.

Invalidation does not always mean immediate rebuild. The system may:

- rebuild only the affected root slice;
- serve a targeted canonical read for that root;
- let the background orchestrator rebuild stale projections;
- return the refreshed slice directly from the write endpoint.

### 6.6 External Booking Intake

Customer bookings from Acuity, Calendly, the Google Form, and iPad/manual intake must resolve through one booking intake adapter.

Current behavior to replace:

- `acuityPollAndSubmit` fetches Acuity appointments and submits new bookings to the configured Google Form.
- `onFormSubmit` resolves the form payload into `00_Master Appointments`.
- `acuityHandleExisting_` edits/reschedules existing Master rows.
- `acuityCancelOnMaster_` cancels existing Master rows.
- `acuityLabelSync` updates Master status from Acuity labels.
- legacy field `CalendlyEventUID` stores both old Calendly UIDs and current Acuity appointment/synthetic reschedule UIDs.

New behavior:

- raw provider/form payloads remain immutable intake evidence;
- normalized appointment facts write to `01_AppointmentEvents`;
- current root pointer writes to `02_RootAppointments`;
- stable customer identity/contact fields write to `03_CustomerInfo`;
- initial sales/customer workflow defaults write to `04_ClientStatus` only when the intake creates a new root or changes a status fact;
- DV requirement fields write to `06_DiamondViewing` only when the booking is a Diamond Viewing or includes DV requirements;
- `00_Master Appointments` receives booking fields only through the compatibility projection.

Provider UID rules:

- introduce normalized `External Booking UID`, `Booking Provider`, and `Provider Appointment ID` fields in appointment-domain reads;
- keep legacy `CalendlyEventUID` only as a compatibility/projection alias;
- dedupe new intake by `(Booking Provider, External Booking UID)`;
- reschedules create a new appointment event with the same `RootApptID`, mark the old event `Rescheduled`, and update `02_RootAppointments.Current APPT_ID`;
- cancellations and label/status sync update the matching appointment event, not customer identity or order domains.

Root resolution order:

1. Existing exact external UID match.
2. Reschedule link from provider/synthetic UID.
3. Recent same contact, same brand, same visit type, same appointment fingerprint.
4. Most recent prior root by normalized email/phone.
5. New root using the new `APPT_ID` as initial `RootApptID`.

Booking intake writes must invalidate calendar, appointment read models, customer cards, admin dashboard, task generation inputs, and the Master compatibility projection.

## 7. Canonical Domain Tabs

### 7.1 `01_AppointmentEvents`

Canonical owner for appointment-event facts.

Primary key:

- `APPT_ID`.

Required link:

- `RootApptID`.

Owns:

- appointment date/time;
- visit type;
- appointment active/cancel/reschedule state;
- booked/canceled/rescheduled timestamps;
- appointment source metadata;
- booking provider;
- provider appointment ID;
- external booking UID;
- reschedule from/to UID links;
- provider label/status sync state;
- appointment-specific intake answers and style notes;
- appointment-level URLs that belong to the appointment event.

Does not own:

- customer identity;
- current customer status;
- 3D/order state;
- payment totals;
- task state.

### 7.2 `02_RootAppointments`

Canonical owner for root-level appointment pointers.

Primary key:

- `RootApptID`.

Owns:

- current active `APPT_ID`;
- latest appointment pointer;
- root lifecycle/currentness;
- root merge/split metadata if needed later;
- root version.

Purpose:

- let root-keyed workflows resolve the active appointment without scanning appointment history.

### 7.2.1 `01_AppointmentEventHistory`

Append-only history for appointment event intake, edits, reschedules, cancellations, and provider label/status sync.

Primary identifiers:

- generated history row ID;
- `APPT_ID`;
- `RootApptID`;
- `External Booking UID`;
- `Booking Provider`.

Owns:

- previous appointment status;
- new appointment status;
- previous visit date/time;
- new visit date/time;
- reschedule/cancel reason when available;
- provider sync source;
- raw provider status/label summary;
- actor or automation source;
- timestamp.

This tab records how appointment event facts changed. The current appointment event state remains in `01_AppointmentEvents`.

### 7.3 `03_CustomerInfo`

Canonical owner for customer identity and stable customer attributes.

Primary key:

- `RootApptID`.

Owns:

- customer name;
- phone/email;
- brand;
- lead source and stable booking-origin profile fields;
- Client Advisor/JOC owner names and emails;
- customer folder/report URLs;
- stable lead/customer identity fields.

Does not own:

- sales stage;
- 3D deadline;
- payment balance;
- task assignment state.

### 7.4 `04_ClientStatus`

Canonical owner for current client workflow state.

Primary key:

- `RootApptID`.

Owns:

- sales stage;
- conversion status;
- custom order status;
- in-production status;
- center stone status;
- next steps;
- order date;
- `3D Deadline`;
- `3D Deadline Move Count`;
- `3D Deadline Updated At`;
- `3D Deadline Updated By`;
- client status version.

3D deadline rule:

- the current 3D deadline and current delay count live here;
- deadline change history lives in `04_ClientStatusHistory`;
- task snooze/delay state for a deadline task lives only in `_SalesTaskQueue`.

### 7.5 `04_ClientStatusHistory`

Append-only history for client status and 3D deadline changes.

Primary identifiers:

- generated history row ID;
- `RootApptID`;
- optional `TaskID`;
- optional `APPT_ID`.

Owns:

- previous value;
- new value;
- changed field;
- reason/note;
- actor;
- timestamp;
- source function/UI.

### 7.6 `05_Order3D`

Canonical owner for current 3D/order setup state.

Primary key:

- `RootApptID`.

Owns:

- SO number;
- Odoo URL;
- SO linked state;
- design request link/state;
- 3D tracker URL;
- order folder links;
- current 3D workflow state not owned by client status.

Does not own:

- current 3D deadline;
- payment totals;
- task state.

### 7.7 `05_Order3DHistory`

Append-only history for 3D starts, revisions, tracker changes, and order setup changes.

Owns:

- revision submissions;
- old/new tracker/order fields;
- actor/source;
- task ID when applicable;
- timestamp.

### 7.8 `05_Wax_Requests`

Existing canonical owner for wax request workflow.

Owns:

- wax request rows;
- wax status;
- request metadata;
- admin pending/commit data.

Customer dashboard wax badges are derived from this source.

### 7.9 `06_DiamondViewing`

Canonical owner for root-level diamond viewing requirements and workflow state.

Primary key:

- `RootApptID`.

Owns:

- customer requirements brief;
- variety strategy;
- DV customer looking-for summary;
- root-level DV workflow status;
- selected center-stone summary if it is customer-level, not per-stone.

Does not own:

- per-stone cert/status/tracking fields, which remain in the 200 stones workbook.

### 7.10 `07_OrderFinance`

Canonical owner for root-level finance summary facts.

Primary key:

- `RootApptID`.

Owns:

- order total;
- saved quote subtotal/lines summary;
- finance version;
- derived paid-to-date/balance snapshot from payment ledger.

Payment ledger remains canonical for individual payments, receipts, voids, and payment history.

### 7.11 Existing Canonical Sources Kept

| Source | Continues To Own |
|---|---|
| `02_Form_Inbox` | immutable raw form intake submissions only |
| `_SalesTaskQueue` | task state, snooze state, task owner, task completion |
| `_SalesTaskLog` | task lifecycle log |
| `_AppointmentArtifacts` | Drive/upload/artifact/AI brief metadata |
| `_SalesWorkflowUsers` | login users, roles, active state |
| `10_Roster_Schedule` | schedule extension data |
| `Schedule Changes` | one-off schedule overrides |
| `200_` stones workbook | per-stone diamond facts |
| payment ledger workbook | individual payment facts |
| external 3D/quote/client workbooks | their own operational records |

## 8. Compatibility Projection

`00_Master Appointments` remains available during migration as a derived compatibility projection.

Rules:

- no new canonical writes should be added to `00_Master Appointments`;
- existing consumers may read it until migrated;
- projection rebuild must be deterministic from domain tabs and existing external canonical sources;
- projection rows should preserve legacy headers and row references required by unmigrated functions;
- projection write drift must be detectable with reconciliation tests.

## 9. UI And Function Flow Map

### 9.1 Dashboard Surfaces

| UI Surface | Client Function | Server Function | Reads | Writes |
|---|---|---|---|---|
| Login/bootstrap | `bootstrap`, `sw_login(...includeBootstrap)` | `sw_getBootstrap`, `sw_login` | user identity, visible views, `TaskListSlice` counts | none |
| My Queue/Cleanup/Coverage/Admin Review | `loadTasks` | `sw_getMyTasks` | `TaskListSlice` | none |
| Task detail drawer | `openTask` | `sw_getTaskDetail` | `TaskDetailSlice`, task-specific root mini slices | none |
| Complete/snooze/claim task | `completeTask`, `snoozeTask`, `claimTask` | `sw_completeTask`, `sw_snoozeTask`, `sw_claimTask` | target task row, task context | `_SalesTaskQueue`, `_SalesTaskLog`, domain adapter when completion changes canonical facts |
| Customer Search/Kanban | `loadCustomerSearch` | `sw_searchCustomers` | `CustomerCardSlice` | none |
| Customer detail card click | `openCustomerSearchDetail` | `sw_getCustomerSearchDetail` or successor shared endpoint | `CustomerRootDetailSlice` | none |
| Customer status update | `saveCustomerSearchStatus` | `sw_customerSearchUpdateStatus` | root status/version | `04_ClientStatus`, `04_ClientStatusHistory` |
| 3D deadline update | `saveCustomerSearch3DDeadline` | `sw_customerSearchUpdate3DDeadline` | root status/version | `04_ClientStatus`, `04_ClientStatusHistory` |
| 3D revision | `submitCustomerSearch3DRevision` | `sw_customerSearchSubmit3DRevision` | `05_Order3D`, root status/version | `05_Order3DHistory`, optional `04_ClientStatus` status change |
| Wax request | `submitCustomerSearchWaxRequest` | `sw_customerSearchRequestWax` | customer/order mini slice | `05_Wax_Requests` |
| Calendar month | `loadCalendar` | `sw_getCalendarAppointments` | `CalendarMonthSlice` | none |
| Calendar event detail | `renderCalendarAppointmentDetail`, `loadCalendarAppointmentAiBrief` | `sw_getAppointmentAiBrief` | loaded calendar row, `AppointmentBriefSlice` | none |
| Admin Dashboard/Pipeline | `loadAdminDashboard` | `sw_getAdminDashboard` | `AdminHealthSlice`, `CustomerCardSlice`, payment/task projections | none |
| Employee schedules | `loadEmployeeSchedules` | `sw_adminGetEmployeeSchedules` | users, roster, schedule changes | none |
| Schedule edits | schedule save/upsert/delete handlers | `sw_adminSaveEmployeeSchedules`, `sw_adminUpsertScheduleChange`, `sw_adminDeleteScheduleChange` | schedule row/version | schedule/user tabs |
| User admin | `openUserAdminPanel` | `sw_adminListWorkflowUsers`, `sw_adminUpsertWorkflowUser` | workflow users | `_SalesWorkflowUsers` |
| Appointment uploads | upload/sync buttons in task detail | `sw_getAppointmentUploadFolder`, `sw_syncAppointmentDriveUploads` | `_AppointmentArtifacts`, Drive state | `_AppointmentArtifacts`, Drive metadata |
| In-stock diamonds | `loadInStockDiamonds` | `sw_getInStockDiamonds` | `DiamondInventorySlice` | none |
| Assign in-stock diamond | diamond assignment action | `sw_assignInStockDiamond` | selected diamond/root mini slice | `200_`, projection invalidation |
| Diamond Tracking | `loadDiamondTracking` | `sw_getDiamondTrackingDashboard` | `DiamondTrackingSlice` | none |
| Bulk Returns | `loadBulkReturns`, submit | `sw_getBulkReturnCandidates`, `sw_bulkMarkDiamondsReturnInProgress` | `DiamondTrackingSlice` | `200_`, task log/projection invalidation |
| Payment dialog/iPad payment | `rp_init`, payment submit flow | `rp_submit`, `rp_getLatest3DFields`, related `rp_*` | `PaymentSummarySlice`, `05_Order3D`, customer mini slice | payment ledger, `07_OrderFinance`, optional `04_ClientStatus` |

### 9.2 External Booking And Background Intake Flows

| Flow | Current Function(s) | New Domain Writes | Notes |
|---|---|---|---|
| New Acuity booking | `acuityPollAndSubmit` -> `acuitySubmitToForm_` -> `onFormSubmit` | `01_AppointmentEvents`, `02_RootAppointments`, `03_CustomerInfo`, optional `04_ClientStatus`, optional `06_DiamondViewing`, `_AppointmentArtifacts` | Form submission may remain as raw intake evidence during migration, but Master is not canonical. |
| Existing booking edit | `acuityHandleExisting_` | `01_AppointmentEvents`, `01_AppointmentEventHistory`, affected customer/DV domain only for changed fields | Phone/contact edits go to `03_CustomerInfo`; appointment time/type edits stay in appointment domain. |
| Reschedule | `acuityHandleExisting_`, `onFormSubmit` reschedule logic | old `01_AppointmentEvents` row marked `Rescheduled`; new `01_AppointmentEvents` row appended; `02_RootAppointments.Current APPT_ID` updated | New event keeps the same `RootApptID`; projection preserves legacy `RescheduledFromUID`/`RescheduledToUID`. |
| Cancellation | `acuityCancelOnMaster_` | `01_AppointmentEvents`, `01_AppointmentEventHistory`, `02_RootAppointments` if current pointer changes | Does not mutate customer identity or order domains. |
| Provider label/status sync | `acuityLabelSync` | `01_AppointmentEvents`, `01_AppointmentEventHistory` | Status labels like Completed, Confirmed, No-Show, and Canceled are appointment-event facts. |
| Direct form/Calendly/manual intake | `onFormSubmit` | same booking adapter as Acuity | `CalendlyEventUID` remains accepted as an input alias for `External Booking UID`. |

### 9.3 Shared Customer Detail Shape

`CustomerRootDetailSlice(rootApptId, mode)` is the common detail contract.

Modes:

| Mode | Used By | Included Sections |
|---|---|---|
| `card` | Customer Search, Kanban, Admin Pipeline cards | identity, owners, current appointment summary, status badges, SO, 3D deadline, wax badge, balance |
| `standard` | Calendar detail/customer preview | `card` plus appointment links, AI brief, current workflow state |
| `full` | customer detail drawer | all sections and action form options |
| `taskMini` | task detail drawer | only fields required by the task type |

Sections:

| Section | Canonical Source |
|---|---|
| `identity` | `03_CustomerInfo` |
| `currentAppointment` | `02_RootAppointments`, `01_AppointmentEvents` |
| `status` | `04_ClientStatus` |
| `statusHistory` | `04_ClientStatusHistory` |
| `order3d` | `05_Order3D` |
| `order3dHistory` | `05_Order3DHistory` |
| `diamondViewing` | `06_DiamondViewing`, root summary from 200 |
| `wax` | `05_Wax_Requests` |
| `finance` | `07_OrderFinance`, payment ledger projection |
| `artifacts` | `_AppointmentArtifacts` |
| `tasks` | `_SalesTaskQueue` filtered by root |
| `recentActivity` | `_SalesTaskLog` plus domain history logs |

## 10. Invalidation Rules

| Canonical Write | Invalidate |
|---|---|
| `01_AppointmentEvents` / `02_RootAppointments` | calendar, customer cards, admin dashboard, task detail root mini-cache |
| `01_AppointmentEventHistory` | appointment brief, customer detail recent activity/history |
| `02_Form_Inbox` raw intake append only | no dashboard invalidation until the booking adapter resolves the row |
| `03_CustomerInfo` | customer cards, customer detail, admin dashboard, calendar labels |
| `04_ClientStatus` | customer cards, customer detail, admin dashboard, reminders/tasks |
| `04_ClientStatusHistory` | customer detail recent activity/history |
| `05_Order3D` / `05_Order3DHistory` | customer cards, customer detail, payment dialogs, reminders/tasks |
| `05_Wax_Requests` | customer detail, customer cards if wax badge/status changes |
| `06_DiamondViewing` / `200_` root DV changes | diamond views, customer detail, customer cards if DV badge/status changes |
| `07_OrderFinance` / payment ledger | payment summary, customer cards, customer detail, admin receivables |
| `_SalesTaskQueue` | task lists, task detail, customer detail task section |
| `_AppointmentArtifacts` | appointment brief, AI brief, customer detail artifacts |
| users/schedules | bootstrap, task generation, assignment/admin views |

Every domain adapter must emit:

- changed domain;
- changed `RootApptID`;
- changed fields;
- actor;
- previous version;
- next version;
- invalidated slices.

## 11. Concurrency And Write Behavior

### 11.1 Write Adapter Contract

Each domain write adapter should:

1. Normalize and validate input.
2. Resolve `RootApptID` and any current `APPT_ID`.
3. Read current domain row and version.
4. Acquire the shortest possible lock.
5. Re-read version inside the lock.
6. Reject or merge on version conflict according to domain policy.
7. Write only the target domain range.
8. Append history if required.
9. Release lock.
10. Invalidate affected slices.
11. Return the refreshed slice required by the UI.

### 11.2 Conflict Policy

Default policy:

- non-overlapping domain writes can both succeed;
- overlapping writes to the same root and same domain require version checks;
- if two users update different fields in `04_ClientStatus`, merge if both were based on the same latest version and fields do not conflict;
- if both update the same field, reject the second write with a stale-data message and return the latest `CustomerRootDetailSlice`;
- append-only history writes do not conflict unless the parent current-state write fails.

Examples:

- Start 3D writes `05_Order3D`; client status update writes `04_ClientStatus`. Both can succeed independently.
- 3D deadline update writes `04_ClientStatus`; client status update also writes `04_ClientStatus`. Merge only if changed fields differ.
- Snoozing a 3D deadline task writes `_SalesTaskQueue`; changing the actual 3D deadline writes `04_ClientStatus`. These are separate facts.

## 12. Migration And Testing Plan

### Phase 0: Baseline Flow Map And Benchmarks

Implementation:

- Freeze the current function-to-sheet map.
- Record all dashboard UI entrypoints and their server functions.
- Run baseline benchmarks for bootstrap, queue, task detail, customer search, customer detail, calendar, admin dashboard, diamond views, and payments.
- Add missing step-level timing logs only where needed for migration validation.

Test gate:

- `sw_measureWorkflowReadModelBuildSpeed()` completes successfully.
- `sw_measureSalesWorkflowSpeed()` completes successfully.
- Manual smoke test covers login, task open, customer search, customer detail, calendar, admin dashboard, diamonds, and payments.
- Baseline output is saved with cache state noted: cold, warm, stale, or expired.

Exit criteria:

- every critical UI surface has a known read source and write path;
- current performance and row counts are documented.

### Phase 1: Create Domain Sheet Registry And Adapters In Read-Only Mode

Implementation:

- Add constants/header definitions for new domain tabs.
- Create setup/repair functions that create tabs and headers only.
- Add read-only domain repository functions that can read from new tabs if present but still fall back to `00_Master Appointments`.
- Add a `RootApptID` resolver shared by all adapters.
- Add a booking intake adapter interface that can accept Acuity, Calendly, Form, and iPad/manual payloads without writing domains yet.
- Add reconciliation utilities comparing current Master-derived values to domain-tab values.

Test gate:

- Setup function can run twice without changing existing data incorrectly.
- No production write function writes to new domain tabs yet.
- Reconciliation reports empty or expected gaps because domain tabs are not backfilled.
- Existing UI smoke tests still pass with reads served from existing sources.

Exit criteria:

- domain tabs exist and are safe;
- adapter interfaces are callable without changing behavior.

### Phase 2: Shadow Backfill From Current Sources

Implementation:

- Backfill `01_AppointmentEvents`, `01_AppointmentEventHistory` where reliable, `02_RootAppointments`, `03_CustomerInfo`, `04_ClientStatus`, `05_Order3D`, `06_DiamondViewing`, and `07_OrderFinance` from current canonical sources.
- Backfill external booking identity from legacy `CalendlyEventUID`, `RescheduledFromUID`, `RescheduledToUID`, appointment status, and `02_Form_Inbox` where available.
- Backfill history tabs only where reliable history exists.
- Keep `00_Master Appointments` as the active read/write source.
- Build reconciliation reports for every mapped field.

Test gate:

- Row counts match expected root and appointment counts.
- Required keys have no blanks: `RootApptID`, current `APPT_ID` pointer where applicable.
- No duplicate `RootApptID` in one-row-per-root domain tabs.
- No duplicate active `(Booking Provider, External Booking UID)` in appointment events.
- Sample at least 25 roots across active, canceled, rescheduled, 3D, DV, wax, and payment cases.
- Sample at least 10 external booking cases across new, edited, rescheduled, canceled, and label-synced appointments.
- Reconciliation mismatches are classified as mapping bug, source data quality issue, or intentional derived difference.

Exit criteria:

- backfilled tabs match current production state for all required fields;
- unresolved mismatches have documented remediation.

### Phase 3: Compatibility Projection Builder

Implementation:

- Build a deterministic projection that can regenerate legacy `00_Master Appointments` fields from domain tabs and external canonical sources.
- Run it in dry-run/report mode first.
- Compare projected values against existing `00_Master Appointments`.
- Confirm projected legacy `CalendlyEventUID`, `RescheduledFromUID`, `RescheduledToUID`, `Status`, and `Active?` values from appointment domains.
- Add projection metadata: source versions, built at, row count, mismatch count.

Test gate:

- Projection dry-run produces no unexpected row loss.
- Critical legacy fields match current Master for sampled roots.
- Existing read models can still rebuild from existing Master.
- Projection report identifies all fields that remain legacy-only.

Exit criteria:

- team can prove Master can become derived before any write cutover.

### Phase 4: Read Cutover For List/Card Views

Implementation:

- Update customer search and admin pipeline list/card reads to use `CustomerCardSlice`.
- Update calendar month reads to use `CalendarMonthSlice`.
- Keep detail reads and writes on current paths.
- Keep fallback to existing read models or Master when domain slices are stale/missing.

Test gate:

- Customer Search returns same roots/cards before and after cutover for standard filters.
- Calendar month returns same active appointment set.
- Admin Dashboard card/pipeline counts match baseline.
- Benchmark proves list/card endpoints use read slices and do not full-scan Master on the warm path.

Exit criteria:

- list/card views are served from small projections without changing user-visible results.

### Phase 5: Shared Customer Detail Slice

Implementation:

- Create or adapt the customer detail endpoint so it returns `CustomerRootDetailSlice`.
- Make Customer Search detail use the shared detail builder.
- Make Calendar expanded customer detail and Admin Pipeline card detail use the same shared detail builder when those surfaces need full customer information.
- Keep task detail on task-specific mini reads unless full customer detail is explicitly opened.

Test gate:

- Customer detail before/after payload contains the same user-visible sections.
- Calendar event detail remains fast and does not fetch full customer detail unless requested.
- Admin Pipeline detail and Customer Search detail show the same root facts.
- Customer detail benchmark improves or remains acceptable with source/fallback noted.

Exit criteria:

- there is one shared root detail payload path for detailed customer information.

### Phase 6: Write Cutover By Domain

Cut over one domain at a time. Do not cut over all writes in one release.

#### 6A: Appointment Intake, External Booking Status, And Root Pointers

Implementation:

- Route `onFormSubmit`, `acuityPollAndSubmit`, `acuityHandleExisting_`, `acuityCancelOnMaster_`, and `acuityLabelSync` through the booking intake adapter.
- Write appointment event facts to `01_AppointmentEvents` and `01_AppointmentEventHistory`.
- Write current root pointers to `02_RootAppointments`.
- Write contact/profile fields to `03_CustomerInfo` only when the incoming data changes customer identity/profile facts.
- Keep `02_Form_Inbox` as raw immutable intake evidence.
- Keep Master updated only through projection/compatibility paths.

Test gate:

- New Acuity/Calendly/form booking creates one appointment event, one root pointer, and one customer info row when the customer is new.
- Existing booking edit updates appointment/contact domains without creating duplicate roots.
- Reschedule marks the old event inactive/rescheduled, creates a new event with the same root, and updates the current pointer.
- Cancellation updates appointment event status and active state without changing customer/order domains.
- Label sync updates appointment status from provider labels and appends history.
- Existing dashboard calendar, customer cards, task generation, and admin booking metrics reflect the intake after invalidation/projection.
- Legacy `CalendlyEventUID` consumers still work from projection during migration.

#### 6B: Client Status And 3D Deadline

Implementation:

- Route `sw_customerSearchUpdateStatus`, `sw_customerSearchUpdate3DDeadline`, `saveRecordDeadline`, and status-dialog deadline branches through `04_ClientStatus`.
- Append status/deadline changes to `04_ClientStatusHistory`.
- Return refreshed `CustomerRootDetailSlice`.
- Keep task snooze state in `_SalesTaskQueue`.

Test gate:

- Updating client status changes only `04_ClientStatus` and history.
- Updating 3D deadline changes only `04_ClientStatus` current deadline fields and history.
- Snoozing deadline task does not change actual deadline.
- Customer card, detail, admin pipeline, reminders/tasks show the new deadline after invalidation/rebuild.
- Concurrent status/deadline updates either merge or return a stale-data response.

#### 6C: 3D Start, SO, And Revisions

Implementation:

- Route `saveAssignedSO`, `swCompleteStart3DTask_`, start-3D flows, and revision submit flows through `05_Order3D` and `05_Order3DHistory`.
- Only update `04_ClientStatus` when current client workflow status actually changes.

Test gate:

- Starting 3D writes order fields to `05_Order3D`.
- Revision history is append-only.
- Payment dialog can still resolve latest SO/3D fields.
- Customer cards and details show SO/tracker state after invalidation.

#### 6D: Customer Identity And Owners

Implementation:

- Route owner assignment and customer identity edits through `03_CustomerInfo`.
- Task generation reads owners from the new domain adapter.
- Master projection keeps legacy owner columns updated for unmigrated reads.

Test gate:

- Admin owner assignment changes customer owner once and task ownership refreshes.
- Client Advisor/JOC filters in Customer Search still work.
- Calendar labels and task cards show updated owners.

#### 6E: Diamond Viewing Root State

Implementation:

- Route root-level DV requirements/status through `06_DiamondViewing`.
- Keep per-stone rows in 200.
- Derive DV customer/card badges from root DV plus 200 summary.

Test gate:

- Proposal workspace loads the same customer requirements.
- Diamond task cards include the same DV brief.
- Per-stone edits still write only to 200.
- Customer detail DV section matches prior behavior.

#### 6F: Finance Summary

Implementation:

- Route root-level order total/quote summary through `07_OrderFinance`.
- Keep individual payments in the payment ledger.
- Payment writes update ledger first, then refresh finance summary/projection.

Test gate:

- Recording payment updates ledger and customer/payment summaries.
- Paid-to-date and balance match payment ledger for sampled roots.
- Admin receivables match baseline totals.
- Payment dialog can reset/replace/print without losing links.

Exit criteria for Phase 6:

- each domain can be cut over, tested, and rolled back independently;
- no canonical fact is written to both Master and a new domain tab.

### Phase 7: Projection-Only Master

Implementation:

- Protect or gate direct writes to `00_Master Appointments`.
- Move remaining Master consumers to domain adapters or read models.
- Rebuild Master projection after canonical writes or by orchestrator.
- Add drift detector comparing projection to domain state.

Test gate:

- Direct write audit shows no production function mutates Master as a canonical source.
- All dashboard smoke tests pass with Master treated as read-only derived output.
- Projection rebuild is deterministic across two consecutive runs.
- Drift detector reports zero unexpected mismatches after common workflow writes.

Exit criteria:

- `00_Master Appointments` is compatibility output only.

### Phase 8: Cleanup And Hardening

Implementation:

- Remove obsolete fallback branches after confidence window.
- Retire redundant Master-derived helper paths.
- Tighten invalidation to affected roots/models.
- Update docs and runbooks.
- Add admin repair/rebuild tools for each domain.

Test gate:

- Full benchmark suite meets performance targets.
- Full smoke suite passes without Master canonical writes.
- Reconciliation reports are clean.
- Background orchestrator rebuilds stale models without user-visible failures.

Exit criteria:

- new infrastructure is the default production path;
- legacy compatibility remains only where explicitly needed.

## 13. Test Suite Requirements

### 13.1 Automated/Scripted Checks

Required checks after each phase:

- domain setup idempotency;
- key uniqueness by domain;
- external booking UID uniqueness and reschedule linkage;
- required key completeness;
- reconciliation against current production values;
- read-model rebuild success;
- benchmark success;
- stale model fallback logging;
- invalidation metadata correctness.

### 13.2 Manual Smoke Tests

Run after each phase:

- login/bootstrap;
- My Queue load;
- open task detail;
- complete a low-risk test task;
- snooze and unsnooze a task;
- Customer Search list and filters;
- customer detail open;
- client status update;
- 3D deadline update;
- calendar month and event detail;
- new booking intake, existing booking edit, reschedule, cancellation, and provider label sync on safe test appointments;
- admin dashboard/pipeline;
- in-stock diamonds;
- diamond tracking;
- bulk returns if applicable;
- payment dialog init and test payment in a safe test root;
- schedule/user admin if the phase touches ownership or users.

### 13.3 Data Reconciliation Tests

For every mapped field:

- source canonical value;
- projected Master value;
- read-model value;
- UI payload value.

The test should report:

- exact matches;
- formatting-only differences;
- missing source;
- missing projection;
- conflicting canonical ownership;
- stale projection.

### 13.4 Concurrency Tests

Run using controlled test roots:

- simultaneous client status and 3D start;
- simultaneous 3D deadline and task snooze;
- simultaneous owner assignment and queue refresh;
- simultaneous booking intake and calendar/admin dashboard load;
- simultaneous payment submit and customer detail open;
- simultaneous diamond root update and diamond tracking dashboard load.

Expected result:

- non-overlapping domains both succeed;
- overlapping same-field writes return stale-data response for the later writer;
- append-only histories remain complete;
- dashboard slices refresh or report stale fallback.

## 14. Performance Targets

Targets for warm read-model/cache path on typical Apps Script runtime:

| Operation | Target |
|---|---|
| `sw_getBootstrap` | no slower than current baseline; preferred under 2 seconds |
| `sw_getMyTasks` | under 1 second for common views |
| `sw_searchCustomers` | under 1.5 seconds warm |
| `sw_getCustomerSearchDetail` / shared detail | under 1.5 seconds warm |
| `sw_getCalendarAppointments` | under 1 second warm |
| `sw_getAdminDashboard` | under 2 seconds warm |
| write actions | lock held only for version check and target write |

Benchmarks must record:

- source path used;
- row counts;
- cache/read-model age;
- fallback reason;
- slowest step;
- invalidated models after write.

## 15. Rollback Strategy

Each phase must be independently reversible.

Rollback defaults:

- read cutover rollback: disable domain/read-slice serving flag and fall back to existing read models/Master;
- write cutover rollback: route the affected domain adapter back to legacy write path before new writes resume;
- projection rollback: keep existing Master untouched until projection-only phase is accepted;
- read-model rollback: mark new read models disabled and rebuild legacy `_SW_*` read models.

Never delete legacy data during rollout. Cleanup occurs only after Phase 8 acceptance.

## 16. Operational Monitoring

Add or preserve logs for:

- domain write adapter name;
- root ID;
- changed fields;
- lock wait time;
- lock held time;
- previous and next version;
- invalidated slices;
- read source path;
- fallback reason;
- reconciliation mismatch counts;
- projection build status.

Admin health should surface:

- stale read models;
- projection drift count;
- failed backfills;
- invalid domain key duplicates;
- last successful domain reconciliation;
- last projection build time.

## 17. Acceptance Criteria

The migration is complete when:

- every canonical fact has exactly one source tab;
- external booking intake writes operational facts to appointment/customer/root domains rather than directly to Master;
- `00_Master Appointments` is projection-only;
- all dashboard list views read from small slices or fresh read models;
- detailed customer surfaces reuse `CustomerRootDetailSlice`;
- each write adapter invalidates only affected slices;
- concurrent non-overlapping workflows can succeed without clobbering each other;
- stale overlapping writes return a clear stale-data response;
- reconciliation shows no unexpected drift;
- full smoke tests pass;
- benchmark logs prove targeted read/write paths are used.

## 18. Assumptions

- `RootApptID` is the stable customer/root key for customer-level workflow data.
- `APPT_ID` remains the stable appointment-event key.
- `CalendlyEventUID` remains a legacy compatibility alias for normalized external booking UIDs during migration.
- The 200 stones workbook remains canonical for per-stone data.
- Payment ledger remains canonical for individual payment rows.
- Existing `_SW_*ReadModel` infrastructure remains useful and will be extended rather than discarded.
- The background orchestrator remains the preferred place to rebuild broader projections.
- Small user-facing writes should not trigger full projection rebuilds inside the lock.
