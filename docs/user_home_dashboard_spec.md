# User Home Dashboard Specification

## 1. Purpose & Scope
- Deliver a single landing page entry point that surfaces identity, KPIs, recent activity, and actionable queues for staff.
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
  - `staff`: hides Jobs/Admin links and quick actions reserved for admins.
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
- [ ] Role-based hiding works for staff vs. admin; server enforces authorization.
- [ ] Widget-level error handling with retry keeps the rest of the dashboard functional.
- [ ] Performance targets satisfied on sample dataset (~1k records).

## 13. Post-v1 Enhancements
- Pin favorite reports per user.
- Personal task list from Jobs queue filtered to "assigned to me".
- Calendar/reminder mini panel once integration exists.
- Export KPIs as downloadable CSV snapshot.
