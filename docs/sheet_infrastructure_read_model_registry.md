# Sheet Infrastructure Read-Model Registry

Current implementation keeps operational source data in its owning workbook and
serves expensive dashboard reads from hidden `_SW_*ReadModel` tabs in the 100
workbook when those models are fresh.

## 100 Master Appointments Workbook

Source/operational tabs:
- `00_Master Appointments`: legacy appointment rows, customer/root compatibility state, owner assignments, payment rollups, diamond summary compatibility columns.
- `02_Form_Inbox`: intake submissions.
- `03_Client_Status_Log`: status history used by admin health.
- `05_Wax_Requests`: wax workflow.
- `07_Root_Index`: last-touch/root health support.
- `10_Roster_Schedule`, `Schedule Changes`, `Assignment Log`, `Daily Availability Cache`: scheduling and assignment support.
- `Dropdown`, `Dropdown Rules`, `Rep Qualifications`: validation/config compatibility.
- `_SalesWorkflowUsers`, `_SalesWorkflowConfig`, `_SalesWorkflowTemplates`: workflow config.
- `_SalesTaskQueue`, `_SalesTaskLog`, `_SalesDataCleanup`, `_AppointmentArtifacts`: workflow tasks/details.

Read-only serving tabs:
- `_SW_TaskReadModel`
- `_SW_CustomerReadModel`
- `_SW_DiamondReadModel`
- `_SW_DiamondRootReadModel`
- `_SW_AppointmentReadModel`
- `_SW_CalendarMonthReadModel`
- `_SW_PaymentReadModel`
- `_SW_AdminDashboardReadModel`
- `_SW_ReadModelMeta`

Legacy/model tabs retained until all consumers are migrated:
- `_Model_CurrentCustomers`, `_Model_AppointmentEvents`, `_Model_DataQuality`, `_IntakeQueue`, `04_Reminders_Queue`, `15_Reminders_Log`, `Log`, `20_Automation_Log`, `90_Validation_Errors`.

## 200 Stones Workbook

Source tab:
- `0. MASTER LG SHEET`

Primary readers:
- Diamond task generation.
- In-stock diamond dashboard.
- Diamond tracking dashboard.
- Bulk return picker.
- Quote refresh and diamond task detail snapshots.
- Loupe360 sync.

Primary writers:
- Proposal submission.
- Order approval.
- Delivery confirmation.
- Stone decisions.
- Return/bulk return.
- Tracking updates.
- Loupe360 sync.
- In-stock assignment.

Serving optimization:
- 200 remains the writable source of truth.
- `_SW_DiamondReadModel` and `_SW_DiamondRootReadModel` mirror only dashboard/workflow fields into 100.

## Payment Ledger Workbook

Source tabs:
- `Payments`
- `Current Fees`

Primary readers:
- Admin dashboard payment metrics.
- Customer detail payment history caches.
- Payment report and previous-payment lookup.

Primary writers:
- `rp_submit`
- payment document replacement/voiding
- `rp_hardResetApptPayments`

Serving optimization:
- Payment ledger remains the writable source of truth.
- `_SW_PaymentReadModel` mirrors active receipt rows needed by dashboard reads.

## External Client Workbooks

Operational targets:
- Quotation workbooks: quote table/named range for diamond quote refresh.
- 3D tracker workbooks: `Log`, `3D Log`, or `3D Revision Log`.
- Client Status reports: `Client Status`.

These remain external write targets. They are not consolidated into 100.

