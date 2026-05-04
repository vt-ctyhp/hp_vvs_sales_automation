# Diamond Viewing Sales Workflow Setup

## Summary
The Sales Workflow web app now includes Diamond Viewing tasks in the main queue and requires email/password login for browser users. Existing Apps Script triggers can still run under Apps Script identity.

The workflow keeps ownership editable in `_SalesWorkflowConfig`, message/task copy editable in `_SalesWorkflowTemplates`, and password hashes in `_SalesWorkflowUsers`.

## Sheets
- `_SalesWorkflowConfig`: workflow settings, admin users, JOC ownership, and diamond order role ownership.
- `_SalesWorkflowTemplates`: task titles, instructions, checklist JSON, attachments, and primary action labels.
- `_SalesWorkflowUsers`: email/password login users with salted password hashes. Do not store raw passwords.
- `_SalesTaskQueue`: generated task rows and payload snapshots.
- 200_ diamond tracker: source of truth for stone status, order status, tracking ETA, decisions, and return due dates.

## Ownership Rows
Run `sw_setupSalesWorkflow()` to seed the system/config rows. Diamond order access is role-based from `_SalesWorkflowUsers`; you do not need to fill Name or Email on `_SalesWorkflowConfig` for these roles.

- `USER | DIAMOND_ORDER_ADMIN_1 | Role=DIAMOND_ORDER_ADMIN`
- `USER | DIAMOND_ORDER_ASSISTANT_1 | Role=DIAMOND_ORDER_ASSISTANT`
- `SYSTEM | SHARED_DIAMOND_ORDER_ADMIN_QUEUE`
- `SYSTEM | SHARED_DIAMOND_ORDER_ASSISTANT_QUEUE`

The seeded diamond config rows can remain blank. Assigned rep and JOC ownership still come from the appointment row (`Assigned Rep`, `Assisted Rep`) and existing roster/config mapping.

## Login Users
For the first admin user, run this no-argument function from the Apps Script editor function dropdown:

```javascript
sw_oneTimeGrantVtAdminAccess();
```

It creates `vt@ctyhp.us` with the `Admin` role and logs the generated password under `SW_BOOTSTRAP_ADMIN_CREATED`. If that login already exists with a password, it does not reset it.

Create or reset a dashboard login with:

```javascript
sw_adminSetWorkflowPassword(
  'person@example.com',
  'temporary-password-here',
  'Person Name',
  'Admin,JOC,DIAMOND_ORDER_ADMIN,DIAMOND_ORDER_ASSISTANT'
);
```

Use only the roles that person needs. A diamond order admin usually needs `DIAMOND_ORDER_ADMIN`; a diamond order assistant usually needs `DIAMOND_ORDER_ASSISTANT`.

Role-only users do not inherit sales rep queues just because their login email appears in the `Dropdown` tab. Add `SALES_REP` only when that login should also receive assigned-rep tasks.

The `Temporary Password?` column is informational only. The dashboard does not force a password change on first login; users can continue using the password assigned to them until an admin resets it.

Admins can also manage users from:
- Google Sheets menu: `Sales > Manage workflow users`
- Sales Workflow dashboard: `Manage Users`

Both surfaces write to `_SalesWorkflowUsers`, support role checkboxes, and either auto-generate a password or use the password typed by the admin.

## Generated Diamond Tasks
- `PROPOSE_DIAMONDS`: assigned rep proposes stones immediately when a Diamond Viewing appointment is in workflow.
- `PREPARE_DV_QUOTATION`: JOC fills quotation, performs price research, and can refresh quotation data from 200_ and latest 3D tracker.
- `ORDER_DIAMONDS`: diamond order admin marks proposed stones as `On the Way` or `Not Approved`. This appears only when 200_ has at least one matching `Proposing` stone.
- `TRACK_DIAMONDS`: diamond order assistant writes tracking ETA/status to 200_. This appears only when 200_ has at least one matching `On the Way` stone.
- `CONFIRM_DIAMOND_DELIVERY`: diamond order admin confirms receipt. This appears only when 200_ has at least one matching `On the Way` stone.
- `RECORD_DIAMOND_DECISIONS`: JOC marks Purchase/Return, confirms dimensions against 3D tracker, and copies the manufacturing message. This appears after a matching stone is delivered or in stock.
- `RETURN_DIAMONDS`: diamond order assistant/admin reviews diamonds due to return. This appears when return-decision stones are due soon or overdue.
- ETA risk tasks go to assigned rep and JOC only when tracking is late or concerning.

## Return Deadline Rule
Return due date is now based on:

```text
Purchased / Ordered Date + 30 days
```

It is intentionally not based on delivery date or memo date.

## Review Setup
Run this read-only function any time you need to audit setup:

```javascript
sw_reviewDiamondWorkflowSetup();
```

It returns and logs:
- required workflow sheet existence and row counts
- diamond role rows from `_SalesWorkflowConfig`
- diamond task template presence from `_SalesWorkflowTemplates`
- login users without exposing password hashes
- 200_ diamond tracker availability

## Operational Notes
- 200_ remains the source of truth for diamond tracking and ETA. Task payloads only cache small snapshots for performance.
- If new 200_ diamond data or 3D tracker data arrives after the task was generated, use the quotation task buttons to refresh diamonds, 3D settings, or both.
- Legacy `04_Reminders_Queue` Diamond Viewing reminders can remain as fallback while the Sales Workflow tasks are verified.
