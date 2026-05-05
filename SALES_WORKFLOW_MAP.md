# Sales Workflow Map

This map is intentionally conservative. It documents the current task queue
shape so future changes can stay small, observable, and low risk.

## Public API Layer

`sales_workflow_api.js` contains the global functions called by the HtmlService
UI, triggers, and admins. Keep these function names stable because the deployed
web app, buttons, and triggers call them directly.

Read-only UI calls:

- `sw_getBootstrap`
- `sw_getMyTasks`
- `sw_adminGetTasks`
- `sw_getTaskDetail`

Mutating task actions:

- `sw_completeTask`
- `sw_acknowledgeTask`
- `sw_claimTask`
- `sw_adminReassignTask`
- `sw_adminBlockTask`
- `sw_adminUnblockTask`
- `sw_logTemplateCopied`

Setup and generation:

- `sw_setupSalesWorkflow`
- `sw_generateSalesWorkflowTasks`
- `sw_refreshTaskOwners` (deprecated compatibility shim for old triggers)
- `sw_installSalesWorkflowTriggers`

## UI Layer

`Index.html` is the task queue web app UI served by `WebApp.js`.

The UI calls the public API through `google.script.run`. The first load should
use `sw_getBootstrap`, which returns the current user, view permissions, counts,
and initial My Queue tasks in one server round trip.

## Generation Layer

`sales_workflow_generation.js` owns task creation, task IDs, dependency
relationships, due-date decisions, owner resolution, appointment relevance, and
appointment active/current checks.

Do not change task IDs, lifecycle rules, owner rules, dependency rules, or due
date rules in readability-only refactors.

## Repository And Sheet IO Layer

`sales_workflow_repository.js` owns spreadsheet access, setup/seed/migration
helpers, task row conversion, task list reads, task writes, task logs, user
identity, config, templates, roster, and low-level sheet helpers.

Sheet headers are part of the contract:

- `_SalesTaskQueue`
- `_SalesTaskLog`
- `_SalesWorkflowConfig`
- `_SalesWorkflowTemplates`

Do not rename, reorder, or remove headers unless there is a dedicated migration
plan and manual verification window.

## Rendering And Utilities

`sales_workflow_rendering.js` owns task detail rendering, attachments, missing
field checks, and completion validation.

`sales_workflow_utils.js` owns normalization, date/time helpers, JSON helpers,
config lookup helpers, identity matching helpers, and timing logs.

Keep `SW_TIMING` logs for broad visibility. Use `SW_TIMING_STEP` only while
diagnosing specific slow paths.
