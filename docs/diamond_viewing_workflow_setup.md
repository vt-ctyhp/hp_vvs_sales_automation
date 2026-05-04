# Diamond Viewing Sales Workflow Setup

## Summary
The Sales Workflow web app now includes Diamond Viewing tasks in the main queue and requires email/password login for browser users. Existing Apps Script triggers can still run under Apps Script identity.

The workflow keeps ownership editable in `_SalesWorkflowConfig`, message/task copy editable in `_SalesWorkflowTemplates`, and password hashes in `_SalesWorkflowUsers`.

## Sheets
- `_SalesWorkflowConfig`: workflow settings, admin users, JOC ownership, and diamond order role ownership.
- `_SalesWorkflowTemplates`: task titles, instructions, checklist JSON, attachments, and primary action labels.
- `_SalesWorkflowUsers`: email/password login users with salted password hashes. Do not store raw passwords.
- `_SalesTaskQueue`: generated task rows and payload snapshots.
- 200_ diamond tracker: source of truth for stone status, order status, tracking ETA/status, decisions, and return due dates. If `Tracking ETA` or `Tracking Status` are missing, the Track Diamonds task creates them when saved.

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
- `PROPOSE_DIAMONDS`: assigned rep enters proposed stones directly in the dashboard. Completion runs the same Sheet 100 proposal flow: validates stones, inserts them into 200_, updates Sheet 100 diamond counts/status, and keeps appointment context tied to the customer.
- `PREPARE_DV_QUOTATION`: JOC fills quotation, performs price research, and can refresh quotation data from 200_ and latest 3D tracker.
- `ORDER_DIAMONDS`: diamond order admin selects `On the Way` or `Not Approved` for every proposed stone. Completion writes order status/date to 200_, then the generator creates acknowledgement tasks for the assigned rep and JOC.
- `ACK_DIAMONDS_ORDERED_ASSIGNED_REP`: assigned rep acknowledges which diamonds were ordered and checks customer impact.
- `ACK_DIAMONDS_ORDERED_JOC`: JOC acknowledges ordered diamonds and updates quotation notes if assumptions changed.
- `TRACK_DIAMONDS`: diamond order assistant writes tracking ETA/status to 200_. This appears only when 200_ has at least one matching `On the Way` stone.
- `CONFIRM_DIAMOND_DELIVERY`: diamond order admin confirms receipt. This appears only when 200_ has at least one matching `On the Way` stone.
- `RECORD_DIAMOND_DECISIONS`: JOC marks Purchase/Return, confirms dimensions against 3D tracker, and copies the manufacturing message. This appears after a matching stone is delivered or in stock.
- `RETURN_DIAMONDS`: diamond order assistant/admin reviews diamonds due to return. Completion marks the listed 200_ rows as `Return in Progress` and writes return notes.
- ETA risk tasks go to assigned rep and JOC only when tracking is late or concerning.

## Dashboard Views
- `My Queue`: role-filtered task cards for the signed-in user.
- `Calendar`: monthly appointment view for all users.
- `In-Stock Diamonds`: read-only store inventory view for all users. It shows diamonds currently marked delivered/in stock in 200_, excludes already purchased/returned stones, and surfaces the return due date so reps can decide whether a store stone will still be available around a Diamond Viewing appointment.
- `Diamond Tracking`: diamond order admin/assistant/admin view showing on-the-way diamonds, delivered diamonds, return items, missing ETAs, delayed/cancelled/unavailable tracking, and returns due soon. It reads from 200_ and highlights issues first.
- `Bulk Returns`: diamond order admin/admin view for selecting multiple eligible delivered/in-stock stones from 200_ and marking them `Return in Progress` for one bulk shipment. The action also writes the stone decision as `Return` and appends shared return notes.

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
- The Order Diamonds task does not ask users to manually update 200_; submitting the dashboard task performs the 200_ writeback through the existing diamond order approval function.
- The Propose Diamonds task does not duplicate a separate workflow; it uses the same validation and writeback path as the current Sheet 100 Propose Diamonds dialog.
- Before proposing diamonds, reps can open `In-Stock Diamonds`, compare the Diamond Viewing date against each stone's return due date, copy the stock details, then enter the chosen stone into the Propose Diamonds task with vendor `From In Stock`.
- Diamond order admins can use `Bulk Returns` when multiple stones are being shipped back together, instead of completing one return task at a time.
- If new 200_ diamond data or 3D tracker data arrives after the task was generated, use the quotation task buttons to refresh diamonds, 3D settings, or both.
- Legacy `04_Reminders_Queue` Diamond Viewing reminders can remain as fallback while the Sales Workflow tasks are verified.
