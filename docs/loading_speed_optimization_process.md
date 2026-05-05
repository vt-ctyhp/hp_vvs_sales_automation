# Loading Speed Optimization Process

## Purpose

This document describes the method we used to pinpoint Sales Workflow loading slowdowns methodically instead of guessing. The goal is to look at the system from multiple angles, instrument each phase, compare cold and warm runs, and work on the largest measured bottlenecks first.

The same process applies to startup, queue loading, customer search, customer detail, calendar, admin dashboard, diamond views, and any future workflow page.

## Operating Principles

- Measure before changing code.
- Separate server-side time from client-side time.
- Separate cold cache, warm cache, stale cache, and expired cache behavior.
- Log every major step with enough context to identify the data source, fallback path, row count, cache age, and elapsed time.
- Optimize the largest measured bottleneck first.
- Keep each phase small enough that one benchmark can prove whether it helped.
- Prefer cached read models, projections, and targeted reads over repeated full-sheet scans.
- Preserve correctness first: cache invalidation, permissions, and workflow writes must remain reliable.

## System Surfaces To Inspect

### 1. Apps Script Project

Review the deployed Apps Script project directly:

- Top-level callable functions exposed to the web app.
- Functions selected in the Apps Script editor dropdown.
- Installed triggers.
- Time-driven triggers for queue refreshes, appointment automation, read-model rebuilds, and cleanup jobs.
- Web app deployment version and whether the tested URL points to the latest deployment.
- Apps Script quotas, execution limits, and logs.

Questions to answer:

- Which function actually serves the page or workflow being tested?
- Which triggers mutate source data or invalidate caches?
- Which function should be run manually before benchmarking?
- Are stale triggers rebuilding data too often, too rarely, or at the wrong time?
- Is the benchmark testing the intended deployment version?

### 2. Server-Side Loading

Server-side time is the time spent inside Apps Script functions before data returns to the browser. We measure this with structured logs such as:

- `SW_TIMING_STEP`
- `SW_TIMING`
- `SW_BENCHMARK_STEP`
- `SW_BENCHMARK_SUMMARY`
- `SW_READ_MODEL_REBUILD`
- `SW_READ_MODEL_BENCHMARK`

Each server endpoint should log:

- operation name;
- step name;
- elapsed milliseconds for the step;
- cumulative elapsed time;
- source used, such as `taskDashboardProjection`, `taskQueue`, `customerReadModelCache`, or source sheet;
- row counts;
- cache age;
- fallback reason;
- relevant user role context.

This lets us distinguish a slow endpoint from a fast endpoint waiting on one specific slow substep.

### 3. Client-Side Loading

Client-side time includes:

- HTML template generation;
- browser parse and render time;
- login screen render;
- app shell render;
- `google.script.run` request timing;
- loading indicators and screen transitions;
- DOM updates after server responses;
- repeated client rendering of large card lists.

The app should track:

- HTML payload size;
- whether login shell and app shell are present;
- first visible UI;
- time from click to loading state;
- time from request start to response;
- time from response to painted UI;
- repeated render cost for large lists or panels.

Server logs alone cannot prove the user experience is fast. A fast server response can still feel slow if the browser builds too much DOM, blocks on large HTML, or rerenders too much.

### 4. Google Sheet Infrastructure

For a Sheets-backed Apps Script app, the spreadsheet is part of the runtime architecture. Review:

- source tabs;
- read-model tabs;
- metadata tabs;
- task/log tabs;
- config/template tabs;
- roster and workflow-user tabs;
- cache/projection sheets if any;
- row counts and column counts;
- volatile formulas;
- data validation ranges;
- hidden dependency tabs;
- cells updated by workflow actions.

Key questions:

- Which tabs are source of truth?
- Which tabs are derived read models?
- Which tabs are read on every page load?
- Which reads use `getDisplayValues`, `getValues`, or row-by-row calls?
- Which code reads an entire sheet when only one row or one root customer is needed?
- Which writes invalidate which derived models?

The largest Apps Script slowdowns often come from repeated sheet reads, full-sheet scans, and many small range calls.

## Benchmark Phases

### Phase 0: Map The Runtime

Document the current workflow before optimizing:

- page or action being tested;
- callable Apps Script functions;
- client event that triggers the server call;
- server function chain;
- source sheets read;
- sheets written;
- caches used;
- triggers that rebuild or invalidate data.

Output:

- a function map;
- a data-flow map;
- a list of benchmark functions to run;
- expected cache prerequisites.

### Phase 1: Baseline End-To-End Timing

Run the broad benchmark before changing code.

For Sales Workflow, the standard sequence is:

```js
sw_measureWorkflowReadModelBuildSpeed()
sw_measureSalesWorkflowSpeed()
```

Record:

- total read-model rebuild time;
- web app HTML size and generation time;
- returning-session bootstrap time;
- warm bootstrap time;
- queue view times;
- task detail times;
- customer search time;
- customer detail time;
- calendar time;
- diamond view times;
- admin dashboard time;
- slowest operations list.

Do not optimize yet. First identify the top two or three slowest user-visible paths.

### Phase 2: Add Step-Level Server Logs

For each slow operation, add or confirm step logs around:

- spreadsheet open;
- required sheet check;
- identity/auth;
- config read;
- source data read;
- read-model/cache read;
- filtering;
- grouping/indexing;
- card/list construction;
- detail payload construction;
- secondary data reads such as payments, logs, files, folders, or form options;
- response formatting.

Good log examples:

```text
SW_TIMING_STEP {"operation":"sw_searchCustomers","step":"rows","ms":576,"extra":{"source":"customerReadModelCache","rows":334}}
SW_TIMING_STEP {"operation":"sw_searchCustomers","step":"cards","ms":915,"extra":{"source":"customerReadModelCache","cards":161,"hiddenCards":153}}
SW_TIMING_STEP {"operation":"sw_getCustomerSearchDetail","step":"payload","ms":766,"extra":{"appointments":1,"tasks":1,"logs":0}}
```

The important part is that the log explains not only how long something took, but why that code path was used.

### Phase 3: Separate Cold, Warm, Stale, And Expired Cache Behavior

Run benchmarks in four states:

- no useful cache;
- freshly rebuilt read models;
- warm in-memory/cache service reads;
- stale or expired read models.

Record the source and fallback reason:

- `taskDashboardProjection`
- `taskReadModelCache`
- `taskReadModelSheet`
- `taskQueue`
- `customerReadModelCache`
- `customerReadModelSheet`
- `appointments`
- `versionMismatch`
- `status:STALE`
- `expired`
- `invalidated`

This prevents false conclusions. A slow run may be caused by a stale model fallback, not by the optimized path.

### Phase 4: Optimize Data Access

Prioritize data-access changes before UI polish.

Typical sequence:

1. Replace full-sheet reads with read models for list pages.
2. Replace repeated list filtering with projections for common dashboard views.
3. Add targeted caches for detail pages.
4. Add row indexes for root/customer/task lookup.
5. Avoid reading appointment/payment/log sheets when cache contains the needed payload.
6. Avoid building objects that will not be returned to the client.
7. Avoid repeated form option or assignment option reads.

Examples from the Sales Workflow work:

- Bootstrap moved from full task queue reads to task dashboard projections.
- Customer Search moved from appointment reads to `customerReadModelCache`.
- Customer detail moved from full customer reconstruction to `customerDetailIndexCache`.
- Payment and recent log data were prewarmed into targeted customer detail caches.
- Hidden customer-search cards were counted without fully assembling discarded card objects.

### Phase 5: Optimize Server Payload Construction

Once data reads are faster, inspect server object construction:

- card building;
- computed badges;
- date parsing;
- payment summaries;
- status summaries;
- attachment lists;
- folder/file lookups;
- action permissions;
- form option groups;
- repeated normalization.

Look for N+1 patterns:

- one Drive call per card;
- one sheet lookup per task;
- one form-options read per detail;
- one payment scan per customer;
- repeated parsing of large JSON payloads.

Fix by batching, indexing, prewarming, or skipping work for hidden/collapsed content.

### Phase 6: Optimize Client Rendering

After server timings are acceptable, measure browser cost:

- how many cards or rows are inserted;
- whether hidden results are rendered;
- whether large panels rerender on small state changes;
- whether loading text appears immediately;
- whether HTML payload size is growing;
- whether client-side filtering duplicates server work.

Client-side optimizations include:

- render only visible cards;
- keep detail panels mounted only when needed;
- avoid full-page rerenders;
- debounce search filters;
- preserve last successful content while refreshing;
- show loading state before server call;
- reduce HTML template size.

### Phase 7: Verify With Focused And Broad Benchmarks

After each change, run:

- focused benchmark for the changed path;
- broad workflow benchmark to catch regressions;
- at least one warm run;
- one stale/expired-cache interpretation if relevant.

For each change, record:

- before timing;
- after timing;
- source/fallback path;
- row counts;
- remaining largest slow step;
- whether deployment was updated.

Do not call a change successful only because total time improved once. Confirm that the expected step improved and the endpoint is still using the intended source.

## Data Flow Review

### Upstream Sources

Identify all source-of-truth inputs:

- `00_Master Appointments`;
- `_SalesTaskQueue`;
- `_SalesTaskLog`;
- `_SalesWorkflowUsers`;
- `_SalesWorkflowConfig`;
- `_SalesWorkflowTemplates`;
- `10_Roster_Schedule`;
- `Schedule Changes`;
- diamond inventory/tracking sheet;
- payment/receipt source;
- Drive folders and artifacts.

For each source, document:

- who writes it;
- which functions read it;
- which workflows depend on it;
- which read models or caches derive from it;
- what invalidates downstream data.

### Derived Models And Caches

Derived data should be explicit:

- task read model;
- customer read model;
- task dashboard projections;
- customer search row cache;
- customer detail index cache;
- payment history cache;
- recent log cache;
- form option cache;
- admin dashboard health indexes.

For each derived model, document:

- builder function;
- storage location;
- TTL;
- version string;
- invalidation function;
- fallback behavior;
- benchmark fields proving it is used.

### Downstream Consumers

Map each page/action to its data:

- startup/bootstrap;
- My Queue;
- Cleanup;
- Coverage;
- Admin Review;
- Task Detail;
- Customer Search;
- Customer Detail;
- Calendar;
- In-Stock Diamonds;
- Diamond Tracking;
- Bulk Returns;
- Admin Dashboard;
- Manage Users;
- Schedules.

This makes it possible to choose optimizations that help multiple pages at once.

## Trigger Review

Audit Apps Script triggers before and after read-model work:

- installed handler name;
- trigger type;
- cadence;
- owner account;
- last run behavior;
- expected output or mutation;
- whether it can overlap with manual actions;
- whether it marks read models stale;
- whether it rebuilds derived models.

Important trigger risks:

- stale read models immediately after a workflow write;
- long-running rebuilds overlapping with benchmarks;
- multiple installed copies of the same trigger;
- trigger running under an account with different permissions;
- rebuild cadence shorter than realistic TTL need;
- rebuild cadence longer than user tolerance for stale data.

## Optimization Department Checklist

### Server-Side Apps Script

- Add `SW_TIMING_STEP` around every expensive section.
- Remove full-sheet reads from request-time paths when read models exist.
- Use one batch read instead of repeated range reads.
- Cache expensive shared data for the request.
- Avoid repeated auth/config reads inside nested helpers.
- Avoid Drive calls inside card loops.
- Avoid JSON parsing large payloads repeatedly.

### Client-Side App

- Measure request time separately from render time.
- Avoid rendering hidden/collapsed results.
- Keep loading states immediate and specific.
- Avoid rebuilding the whole shell for detail-panel changes.
- Debounce filters and searches.
- Keep payloads small enough for fast serialization and DOM updates.

### Google Sheets Infrastructure

- Keep source tabs normalized enough for indexed reads.
- Move repeated dashboard reads into read-model tabs.
- Store source row numbers in read models for targeted detail reads.
- Keep metadata with `Status`, `Built At`, `Expires At`, `Invalidated At`, and `Version`.
- Avoid volatile formulas in tabs read on every request.
- Avoid unnecessary formatting or formulas across thousands of unused rows.

### Workflow And Data Movement

- Define which writes invalidate tasks versus customers.
- Keep invalidation targeted: task changes should not always invalidate customer data.
- Rebuild expensive models on schedule or explicit admin action, not inside every user request.
- Prewarm high-traffic detail caches after read-model rebuild.
- Keep fallback paths correct but visible in logs.

## Prioritization Framework

Rank targets by:

- user frequency;
- total elapsed time;
- slowest step size;
- number of users affected;
- whether it blocks first load;
- whether it affects multiple pages;
- implementation risk;
- cache correctness risk.

Preferred order:

1. Startup/bootstrap.
2. Primary queue views.
3. Most common detail open.
4. Search/list pages.
5. Admin dashboard and reporting.
6. Lower-frequency operational views.

## Standard Performance Report Template

Use this format after each benchmark:

```text
Date/time:
Deployment/version:
Cache state:
Read model status:

Top slow operations:
1.
2.
3.

Target operation:
Before:
After:
Largest remaining step:
Source/fallback:
Rows/cards returned:

Change made:
Risk:
Next recommended target:
```

## Success Criteria

A performance phase is successful when:

- benchmark logs show the intended source path;
- the target step improves materially;
- no protected fallback path broke;
- no stale-cache behavior is hidden;
- broad benchmark has no new critical regression;
- deployment version is recorded;
- the next bottleneck is clear.

## Current Lessons Learned

- Read-model rebuild speed is not the same as request speed. A rebuild can be expensive if it makes frequent user requests much faster.
- A stale read model can make startup look slow because the app falls back to source sheets.
- Cache source labels are essential. Without `source` and `fallbackReason`, a timing number is ambiguous.
- Customer detail improved most when we stopped reconstructing it from full source sheets and used targeted detail, payment, and recent-log caches.
- Search/list pages can still be slow after data reads are cached if the server constructs many cards that the client never receives.
- The best next target is usually visible in the benchmark summary, but the cause is visible only in step-level logs.
