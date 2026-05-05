# User Home Dashboard Specification

## 1. Purpose & Scope
- Deliver a single landing page entry point that surfaces identity, KPIs, recent activity, and actionable queues for Client Advisors.
- Provide global navigation to the core application modules and shortcuts for common creation flows.
- Surface actionable alerts (e.g., 3-day 3D checks due) and high-level metrics without enabling heavy editing on the dashboard itself.

## 2. Routing & Access
- Route `/` renders the dashboard; unauthenticated visitors are redirected to `/login`.
- Primary navigation destinations:
  - `/customers`
  - `/orders`
  - `/revisions`
  - `/payments`
  - `/reports`
  - `/jobs` *(admin only)*
  - `/admin` *(admin only)*
- Role visibility:
  - `client_advisor`: hides Jobs/Admin links and quick actions reserved for admins.
  - `admin`: sees all sections, including Jobs runner and audit/health links.
- Front end must hide unauthorized controls, and server endpoints must re-check role on protected actions.

## 3. Data Fetching Contracts
Execute the following authenticated requests in parallel when the dashboard mounts. Handle partial failures gracefully (see §9).

1. **Current User**
   - `GET /api/me` → `{ id, email, role }`.
2. **KPIs**
   - `GET /api/reports/kpis?start=<iso>&end=<iso>` where the client supplies the window (today/7-day/30-day/etc.).
   - Response provides counts/sums that populate KPI cards.
3. **Action Queues**
   - `GET /api/orders/due?type=3d_check&limit=10` for 3-day 3D reviews.
   - `GET /api/orders/awaiting_payment?limit=10` for unpaid orders.
   - Optional: `GET /api/orders/in_production?limit=10`.
   - If specialized endpoints are unavailable, fall back to `GET /api/sales-orders?status=` filtering client-side.
4. **Recent Activity**
   - Preferred: `GET /api/audit?limit=20` returning latest edits/documents/payments.
   - Fallback: `GET /api/sales-orders?from=<iso>&to=<iso>&limit=20` and `GET /api/payments?from=<iso>&to=<iso>&limit=20`.
5. **Counts for Navigation Badges (optional)**
   - `GET /api/sales-orders?status=lead&countOnly=true` and similar endpoints for other statuses.

All requests must return HTTP 200 or a structured error body. Auth failures (401/403) must clear the token and redirect to `/login`.

## 4. Layout Structure

### A. Top Bar
- Left: application logo/name ("VVS App").
- Center: global search input (see §5).
- Right: user menu showing email and role with links: Profile (placeholder) and Sign out (clears token and redirects to `/login`).

### B. KPI Cards
- Render 3–6 KPI cards (e.g., New Leads, Hot Leads, Deposits Taken, In Production, Shipped, Payments Received).
- Each card displays count/sum for the selected time window and links to the filtered module view.
- Provide ARIA labels and adequate contrast for accessibility.

### C. Quick Actions
- Buttons: `+ New Customer` → `/customers?new=1`, `+ New Sales Order` → `/orders?new=1`, `Record Payment` → `/payments?new=1`, `Upload Revision` → `/revisions?new=1`.
- Optional admin-only `Run Jobs Now` button that triggers `POST /api/jobs/run`.

### D. Action Queues
- Two-column lists limited to 10 entries each.
  - **Due 3D Checks**: show SO code, customer, started_at, days since start. Row actions: Open SO, Mark Reviewed (if API exists), optional Snooze.
  - **Awaiting Payment**: show SO code, customer, balance due, last activity. Row actions: Record Payment, Open SO.
- Empty states: "All caught up—no items due."

### E. Recent Activity
- Display up to 20 audit entries (timestamp, actor, action, entity) with deep links. Fallback to recent sales orders/payments if audit feed unavailable.
- Include "View all" link pointing to `/admin/audit` (admins) or module pages.

### F. Footer
- Show app version (from build/config), server time, and admin-only link to system health readout.

## 5. Global Search
- Text input submits parallel requests:
  - `GET /api/customers?query=<q>&limit=5`
  - `GET /api/sales-orders?query=<q>&limit=5`
  - Optional: `GET /api/payments?query=<q>&limit=5`
- Render dropdown grouping results by entity type with keyboard navigation (arrow keys, Enter). Escape closes dropdown.
- Selecting a result navigates to the detail page or list with prefilled filter when multiple matches exist.
- Display loading spinner while requests in-flight and handle empty results with a friendly message.

## 6. State Management & Caching
- Kick off all primary data requests on mount using the current auth token.
- Cache the last successful response in memory to avoid flashing skeletons when returning to the dashboard quickly.
- Optionally persist the selected KPI time window in `localStorage`.

## 7. Time Window Selector
- Provide preset options: Today, Last 7 Days, Last 30 Days, This Month, Custom.
- Selector positioned near KPI cards; changes debounced before refetching KPIs.
- Show "No data in this window" for empty KPI responses.

## 8. Loading & Empty States
- Use skeleton placeholders for KPI cards, action queues, and recent activity while fetching.
- Empty states for queues and activity feed match tone described in §4.
- Global search dropdown shows spinner until queries resolve.

## 9. Error Handling & Telemetry
- Each widget (KPIs, queues, activity, search) shows inline error banner with a Retry button if its request fails.
- Logging:
  - On dashboard load, emit `{ userId, role, ts, kpiWindow }` to telemetry endpoint.
  - Log retries/failures with widget name, endpoint, status code (no PII beyond IDs).
- Authentication failures clear token and redirect to `/login`.

## 10. Performance Targets
- First paint ≤ 1.0s and populated KPIs ≤ 2.0s with cached auth token on typical hardware.
- API response targets: KPIs ≤ 200ms, queues ≤ 250ms, recent activity ≤ 250ms for requested payload sizes.
- Avoid N+1 server queries; batch related data when possible.

## 11. Accessibility & Security
- All controls reachable via keyboard; ensure ARIA labels for quick actions and KPI cards.
- Maintain WCAG AA contrast for text/badges.
- Attach `Authorization: Bearer <token>` to requests; avoid embedding secrets in front-end bundle.
- Sanitize/escape text in activity feed to prevent XSS.

## 12. Acceptance Criteria Checklist
- [ ] `/` displays current user identity and role-appropriate content.
- [ ] KPI cards respect selected time window and link to filtered module views.
- [ ] Quick action buttons navigate to appropriate creation flows.
- [ ] 3D check and awaiting payment queues display up to 10 items with correct links/actions.
- [ ] Recent activity feed surfaces the latest 20 items with accurate timestamps and links.
- [ ] Global search locates customers/orders and navigates via Enter/click.
- [ ] Role-based hiding works for Client Advisor vs. admin; server enforces authorization.
- [ ] Widget-level error handling with retry keeps the rest of the dashboard functional.
- [ ] Performance targets satisfied on sample dataset (~1k records).

## 13. Post-v1 Enhancements
- Pin favorite reports per user.
- Personal task list from Jobs queue filtered to "assigned to me".
- Calendar/reminder mini panel once integration exists.
- Export KPIs as downloadable CSV snapshot.

## 14. Sales Workflow Web App Notes
- The current Apps Script Sales Workflow dashboard uses an email/password login screen before showing the task queue.
- Login users, active status, role access, and schedulable staff identity are stored in `_SalesWorkflowUsers`.
- Diamond Order Admin and Diamond Order Assistant access is role-based; `_SalesWorkflowConfig` name/email cells are not required for those queues.
- Client Advisor access is stored and displayed as `Client Advisor`; legacy `SALES_REP` values are still accepted as an alias and normalized during setup.
- Client Advisor and JOC task ownership resolve appointment owner names/emails against active canonical users in `_SalesWorkflowUsers`. Unresolved Client Advisor owners route to Admin Review; unavailable JOC owners route through JOC Coverage.
- The dashboard has a shared `Calendar` tab for all users. It shows active upcoming appointments by month from `00_Master Appointments`, with appointment links and Client Advisor/JOC details in the side panel.
- The dashboard has a shared `In-Stock Diamonds` tab for all users. It reads 200_ and shows currently delivered/in-stock diamonds with return due dates for proposal planning. It supports filtering by shape, carat size range, color, and clarity. Its healthy return-date bucket is labeled `Available > 7d`.
- Diamond order admin/assistant/admin users also see a `Diamond Tracking` tab. It reads 200_ tracking ETA/status, highlights missing or concerning ETAs, and surfaces return-deadline issues.
- Diamond order admin/admin users also see a `Bulk Returns` tab. It reads eligible delivered/in-stock 200_ rows, supports selecting multiple stones for one return shipment, then marks selected rows `Return in Progress`, writes `Return`, and appends shared return notes.
- The `PROPOSE_DIAMONDS` task captures a structured customer requirements brief before proposed stones: summary, stone type/shape, carat/color/clarity ranges, ratio and budget notes, primary deciding factor, and variety focus checkboxes. It stores the brief in Sheet 100 columns `DV Customer Looking For`, `DV Variety Strategy`, and `DV Customer Requirements (JSON)`.
- Opening a `PROPOSE_DIAMONDS` task from the queue uses a full-width three-step proposal workspace rather than the standard side detail panel: customer requirements, find/add stones, then review/submit. The workspace keeps the existing queue navigation/header and existing task completion/writeback path.
- The proposal workspace's inventory matching uses the existing `sw_getInStockDiamonds` web-app API, filters the returned in-stock `200_` rows against the entered requirements, and can import selected matches into the proposal stone fields. Match rows show the diamond `Stone Type` so Client Advisors can see whether each matched stone is lab or natural before adding it. Pricing is intentionally omitted until the workflow has reliable price data.
- JOC, diamond order admin, and Client Advisor diamond task cards receive the Sheet 100 customer requirements in their generated payloads so quote/order decisions are reviewed against the same brief.
- Admins add/update Client Advisors and JOCs from `Sales > Manage workflow users` in Sheets or `Manage Users` in the dashboard, with either auto-generated or admin-entered passwords. Saving a schedulable user links or creates that user's `10_Roster_Schedule` extension row by email.
- Admins manage schedules from the dashboard `Schedules` tab. It edits weekly working days, Client Advisor skills, linked JOC, and backup JOC in `10_Roster_Schedule`; one-off overrides stay in `Schedule Changes`; names, emails, roles, and active login status remain canonical in `_SalesWorkflowUsers`.
- Client Advisor auto-assignment can be enabled from the `Schedules` tab. When enabled, queue refresh assigns or reassigns active appointments by working-day availability, lab/natural/general qualification, and round-robin load. The selected Client Advisor's linked JOC is written to the appointment as the paired JOC.
- Schedule changes support full-day off and partial-day availability windows. Saving a schedule or override refreshes the workflow so affected appointments/tasks are reassigned.
- JOC coverage now preserves named queues: if the linked/intended JOC is unavailable, the workflow tries that JOC's coverage partner, then another working JOC, then the shared `JOC Coverage` queue.
- Diamond Viewing workflow setup, ownership, templates, and review steps are documented in [`diamond_viewing_workflow_setup.md`](diamond_viewing_workflow_setup.md).

## 15. Current High-Level Workflow

### A. Setup, Access, and Bootstrap
- Admin setup creates or repairs the workflow sheets: `_SalesTaskQueue`, `_SalesTaskLog`, `_SalesWorkflowConfig`, `_SalesWorkflowTemplates`, and `_SalesWorkflowUsers`.
- The Apps Script web app opens to a password login screen. Successful login returns a workflow token and current user profile from `_SalesWorkflowUsers`.
- The browser stores the token in `sessionStorage` and calls `sw_getBootstrap`.
- `sw_getBootstrap` authenticates the token, reads the current task queue, calculates role-visible views, and returns:
  - current user and role flags;
  - My Queue tasks;
  - counts for My Queue, JOC Coverage, and Admin Review;
  - visibility for Calendar, In-Stock Diamonds, Diamond Tracking, Bulk Returns, Admin Dashboard, and Admin Review.

### B. Task Generation and Queue Refresh
- `sw_generateSalesWorkflowTasks` is the central queue builder. It reads appointments from `00_Master Appointments`, supporting config/template sheets, canonical workflow users, roster/schedule extension data, wax state, and diamond tracker data.
- `sw_installSalesWorkflowTriggers` installs:
  - hourly queue refresh;
  - 5-minute appointment automation.
- Admins can also manually refresh the queue from the dashboard.
- Task generation is idempotent by `TaskID`. Existing pending tasks are updated in place; completed tasks and claimed tasks are not reassigned by normal generation.
- Every create, assignment, completion, snooze, claim, block, unblock, and bulk return action is logged to `_SalesTaskLog`.

### C. Ownership Model
- System tasks are auto-completed and used as dependency anchors.
- Client Advisor tasks resolve from `Client Advisor` or legacy `Assigned Rep` on `00_Master Appointments`; the owner must match one active canonical workflow user by email or unique name.
- JOC tasks resolve from `Assisted Rep` on `00_Master Appointments`; the owner must match one active canonical JOC user by email or unique name. If there is no assisted rep or the assisted rep is unavailable, the task goes to `JOC Coverage`.
- JOC Coverage tasks are visible to JOC users and admins and can be claimed.
- Diamond Order Admin and Diamond Order Assistant tasks are shared role queues. Users with `DIAMOND_ORDER_ADMIN` or `DIAMOND_ORDER_ASSISTANT` in `_SalesWorkflowUsers` see the relevant role-owned tasks.
- Admins can assign appointment owners from the task detail panel. Saving requires active canonical Client Advisor/JOC users, writes canonical names/emails back to all `00_Master Appointments` rows with the same RootApptID, and then refreshes workflow ownership.
- Admins can also reassign, block, or unblock individual tasks from Admin Review.

### D. Dashboard Navigation Groups
- **Work**: My Queue for assigned tasks, Calendar for upcoming appointments, and JOC Coverage for shared coverage work.
- **Customers**: Customer Search for finding/updating customer records and Customer Pipeline for the admin kanban/health dashboard.
- **Diamonds**: Diamond Inventory, Diamond Tracking, and Bulk Returns as the diamond operations trio.
- **Admin**: Schedules, Admin Review, Manage Users, and the temporary Cleanup campaign.
- **Refresh Queue** remains a global action because it reruns task generation rather than opening a page.

### E. Dashboard Views
- **My Queue**: due pending tasks owned by the current user, including role-owned diamond tasks.
- **Calendar**: active upcoming appointments by month, visible to all workflow users.
- **Customer Search**: active customer lookup, filters, detail review, and selected status/deadline/wax actions. Client Advisors default to their own Client Advisor filter; JOC users default to their own JOC filter.
- **In-Stock Diamonds**: delivered/in-stock diamonds from `200_`, visible to all workflow users for proposal planning.
- **Diamond Tracking**: shipment, ETA, return, and tracking issues from `200_`, visible to admins and diamond order roles.
- **Bulk Returns**: return-eligible in-stock diamonds from `200_`, visible to admins and Diamond Order Admin users.
- **JOC Coverage**: unclaimed or coverage-routed JOC tasks, visible to JOC users and admins.
- **Admin Dashboard / Customer Pipeline**: weekly metrics and customer pipeline, visible to admins.
- **Admin Review**: all non-completed workflow tasks, visible to admins.
- **Manage Users**: add/update workflow users, roles, and passwords, visible to admins.
- **Cleanup**: temporary one-time stale customer cleanup campaign tab, visible while `DATA_CLEANUP_CAMPAIGN_TAB_ENABLED = Y`. Client Advisors see their assigned cleanup work; admins see all cleanup campaign tasks. JOC-side cleanup work routes through the shared `JOC Coverage` queue so any active JOC can claim and submit it. After campaign cases are resolved, future stale customers flow into the normal queue rather than this tab.

### F. People Data Cleanup APIs
- `sw_adminAuditWorkflowPeopleData(authToken)` compares `_SalesWorkflowUsers`, `10_Roster_Schedule`, `Schedule Changes`, `Dropdown`, and active appointment owners. It reports duplicate active emails, duplicate active Client Advisor/JOC names, orphan extension rows, Dropdown-only identities, and appointment owners that do not map to active canonical users.
- `sw_adminMigrateWorkflowPeople(authToken, { dryRun, clearDropdownIdentityData })` defaults to `dryRun: true`. Mutating mode links extension rows by exact email first, then unique normalized name; migrates default JOC pairings from `Dropdown`; optionally backs up and clears legacy identity cells from `Dropdown`; then refreshes workflow tasks.
- The old `Rep Qualifications` tab has been retired; current skill and routing data lives in `10_Roster_Schedule` and `_SalesWorkflowUsers`.

### G. Standard Task Execution
- A user opens a task card to load task detail, rendered instructions, copyable message/template text, links/attachments, checklists, and task-specific controls.
- The user may copy a template, snooze a task, claim a coverage task, or complete the task.
- Completion validates required fields and checklists, runs any task-specific writeback adapter, marks the task completed, logs the event, and immediately runs task generation again so downstream tasks appear.
- Snoozed tasks are hidden from active queues until the snooze date and do not count as late during the snooze window.

### H. Core Appointment Workflow
- When an appointment enters the workflow window, the system creates an auto-completed assignment task.
- If the appointment is within 24 hours, JOC gets a Hybrid Welcome + Instructions task.
- If the appointment is farther out, JOC gets a Welcome task now and a Map & Instructions task 48 hours before the appointment.
- The assigned Client Advisor gets a Review Appointment task 24 hours before the appointment and an Appointment Day Checklist task on the appointment day.
- After the checklist is completed, JOC gets Process Appointment Data.
- After JOC submits the recap draft, the Client Advisor gets Approve/Edit Recap Message.
- After approval, JOC gets Send Final Recap Text.

### I. Post-Consult Operations
- After the appointment checklist is complete, JOC gets Post-Consult Client Status Update.
- That task writes the current client status and captures whether 3D and/or wax work is needed.
- If 3D is needed and no SO/tracker exists yet, JOC gets Start 3D Design.
- If 3D has started but the 3D deadline is missing, JOC gets Record 3D Deadline for the next day at 9:30am.
- If wax is needed and no active wax request exists, JOC gets Request Wax Print.
- If existing wax requests need status or deadline updates, JOC gets Update Wax Request.

### J. Diamond Viewing Workflow
- For Diamond Viewing appointments, the assigned Client Advisor gets Propose Diamonds and JOC gets Prepare Diamond Viewing Quotation.
- Propose Diamonds captures the structured customer requirements brief and proposed stones, then writes requirements to Sheet 100 and proposed diamonds to `200_`.
- If `200_` has proposed stones, Diamond Order Admin gets Order Diamonds.
- When diamonds are marked On the Way, Diamond Order Assistant gets Track Diamond ETA, Diamond Order Admin gets Confirm Diamond Delivery, and the assigned Client Advisor/JOC get acknowledgement tasks.
- If delivered or in-stock stones exist, JOC gets Record Diamond Decisions on or after the viewing date.
- If stones need return and are within the return-warning window, Diamond Order Assistant gets Return Diamonds.
- If ETA risk is detected, both assigned Client Advisor and JOC get ETA risk review tasks.
- Bulk Returns lets admins or Diamond Order Admin users mark multiple eligible `200_` rows as `Return in Progress` in one shipment.

### K. Admin and Diagnostics
- Admin Dashboard aggregates appointment metrics, payments when configured, customer pipeline columns, and admin-visible open task health.
- Diagnostic helpers exist for setup review, task visibility troubleshooting, speed benchmarking, duplicate cleanup, and duplicate-safe generation testing.

### L. Customer Data Cleanup
- Setup creates `_SalesDataCleanup` plus Master cleanup columns: `Lost Lead Reason`, `Lost Lead Reason Notes`, `Data Cleanup Reviewed At`, and `Data Cleanup Confirmed At`.
- Config rows control the workflow: `DATA_CLEANUP_ENABLED`, `DATA_CLEANUP_STALE_DAYS` (default 30), `DATA_CLEANUP_CAMPAIGN_ID`, and `DATA_CLEANUP_CAMPAIGN_TAB_ENABLED`.
- Generation creates cleanup cases for active Lead / Hot Lead / Follow-Up customer roots with no meaningful touch for 30+ calendar days, excluding Won/Lost Lead and any root with an unresolved cleanup case.
- During the one-time campaign, the Client Advisor side routes to the assigned Client Advisor and the JOC side routes to shared `JOC Coverage`, not the appointment's assigned JOC. The first submitter proposes the update; the opposite side receives confirmation. Customer data writes back only after one JOC-side submitter and one assigned Client Advisor-side submitter have participated.
- Customer records are updated only after second-person confirmation. Returned confirmations create a revision task for the proposer and do not write back.
- After unresolved campaign cases reach zero, the campaign tab is disabled by config. Ongoing stale customers continue to create cleanup tasks in the regular queue.

## 16. Ownership and Functionality Gaps

### A. Ownership Gaps
- **Spec owner for REST vs Apps Script architecture**: sections 2-13 describe REST routes, KPI cards, global search, recent activity, and quick actions, but the current implementation is an Apps Script `google.script.run` dashboard. Decide whether the REST dashboard is future scope or replace those sections with the Apps Script contract.
- **Workflow people data owner**: `_SalesWorkflowUsers` is now canonical for schedulable staff, but one operational owner should still run the audit/migration tools and keep appointment owner names/emails aligned with active canonical users.
- **Unassigned Client Advisor task owner**: missing Client Advisor/legacy `Assigned Rep` routes tasks to `Admin Review` with `UNASSIGNED_REP`, but there is no sales coverage queue or claim path equivalent to JOC Coverage.
- **JOC coverage remediation owner**: JOC users can claim coverage tasks, but missing assisted rep, missing schedule data, or out-of-office routing still needs a defined admin/JOC process for fixing the source data.
- **Shared diamond role queue owner**: Diamond Order Admin and Assistant tasks are role-owned shared queues. There is no claim/lock workflow to show who is actively working a shared diamond task before completion.
- **Queue freshness owner**: tasks are generated hourly and manually, and the dashboard exposes this as Refresh Queue. Define who owns immediate refresh when `200_`, wax, or appointment data changes.
- **Source-system schema owner**: `200_`, Sheet 100, the 3D tracker, wax tracker, and external payments ledger each have columns/functions the dashboard expects. Missing-column warnings exist in some areas, but ownership for schema drift is not explicit.
- **Claimed/completed task reassignment policy**: generation intentionally does not reassign completed or claimed tasks. The policy is sound, but the business rule for owner changes after claim/completion should be explicit.

### B. Functionality Gaps
- **REST homepage features are not implemented in the Apps Script dashboard**: KPI cards, global search, quick creation shortcuts, recent activity feed, footer/version metadata, and per-widget retry are still aspirational unless the REST dashboard is built.
- **Search is missing**: there is no dashboard-wide search across appointments, customers, SOs, payments, or diamond records.
- **Recent activity is missing**: `_SalesTaskLog` exists, but there is no user-facing recent activity feed.
- **Payment and production queues are incomplete**: Admin Dashboard can read payment metrics when configured, but there are no Client Advisor-facing Awaiting Payment, In Production, or 3-day 3D check widgets as described in the original dashboard spec.
- **Widget-level resilience is limited**: the web app has global status/error handling, but most views do not independently preserve partial results or provide widget-level retry controls.
- **Shared-task concurrency is weak**: task completion and shared role queues do not appear to use a per-task lock/claim requirement, so two users can potentially act on the same role-owned task at nearly the same time.
- **Push notifications/escalations are not part of the dashboard**: overdue and risk tasks appear in queues, but there is no built-in email/Slack/push escalation from this dashboard layer.
- **Automated regression coverage is thin**: diagnostics and dry-run helpers exist, but there is no clear automated test suite for role visibility, generation idempotency, post-consult writebacks, diamond writebacks, or dashboard rendering.
- **Performance targets are not tied to acceptance checks**: `sw_measureSalesWorkflowSpeed` measures server-side operations, but the spec's first-paint/populated-widget targets are not currently enforced.
