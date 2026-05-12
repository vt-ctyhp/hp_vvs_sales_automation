# Acuity + Calendly Webhook Setup and Troubleshooting Runbook

## Purpose
This document explains the exact current booking webhook setup so a developer can maintain and troubleshoot it without guessing.

Scope:
- HP Acuity webhook relay + queue + Apps Script processing
- Existing Calendly webhook system (kept separate by design)
- How to verify end-to-end behavior
- What to check first when something breaks

---

## 1) Current Architecture (Important)

### Calendly and Acuity are **not** the same pipeline.

1. **Calendly** runs in a separate Apps Script project with its own `doPost`, queue, and worker.
2. **Acuity** uses a Go relay endpoint (`POST /api/webhooks/acuity`) that verifies signature and appends rows to `_ExternalBookingEvents`.
3. Main workflow Apps Script then processes `_ExternalBookingEvents` via `sw_processExternalBookingEvents`.

This separation is intentional to minimize risk to the live Calendly flow.

---

## 2) Critical IDs and URLs

### HP Acuity queue spreadsheet
- Spreadsheet ID: `1gWXpPXkuoNdK9I5KyPqDYD3Re0Mbn3ouDxpCjxJ_5MA`
- Full URL:  
  `https://docs.google.com/spreadsheets/d/1gWXpPXkuoNdK9I5KyPqDYD3Re0Mbn3ouDxpCjxJ_5MA/edit?gid=327751876#gid=327751876`
- Queue tab name: `_ExternalBookingEvents`

### Calendly script (separate project)
- Script URL:  
  `https://script.google.com/d/1UkzFC045Wlkpb6SSCHmZyOFNcm0tdn6h07pnic2Kyw-56odiASu7Z9xA/edit`

### Acuity relay endpoint path
- Path: `/api/webhooks/acuity`
- Full webhook URL in Acuity = `<your deployed vvsapp base URL>/api/webhooks/acuity`

---

## 3) Configuration Reference

### 3.1 Go relay environment variables (server)

Required:
- `HPAPP_ACUITY_WEBHOOK_SECRET`
- `HPAPP_ACUITY_QUEUE_SPREADSHEET_ID`
- `VVSAPP_GOOGLE_SERVICE_ACCOUNT_JSON` **or** `VVSAPP_GOOGLE_SERVICE_ACCOUNT_FILE`

Optional:
- `HPAPP_EXTERNAL_BOOKING_QUEUE_RANGE` (default is `'_ExternalBookingEvents'!A:P`)

Backward-compat fallbacks in code:
- `VVSAPP_ACUITY_WEBHOOK_SECRET` / `VVSAPP_ACUITY_API_KEY` (only used if `HPAPP_ACUITY_WEBHOOK_SECRET` is empty)
- `VVSAPP_BOOKING_SPREADSHEET_ID`
- `VVSAPP_EXTERNAL_BOOKING_QUEUE_RANGE`

### 3.2 Apps Script Script Properties (main workflow project)

Used by Acuity queue processor:
- `FORM_ID`
- `ACUITY_USER_ID`
- `ACUITY_API_KEY`
- `HPAPP_ACUITY_QUEUE_SPREADSHEET_ID` (and compat alias `EXTERNAL_BOOKING_QUEUE_SPREADSHEET_ID`)

Note:
- `HPAPP_ACUITY_WEBHOOK_SECRET` is primarily for relay signature verification on the server side.
- It can exist in Apps Script properties for visibility/operational consistency, but Apps Script processor does not verify webhook signatures itself.

---

## 4) Queue Schema (`_ExternalBookingEvents`)

Columns in order:
1. `ReceivedAt`
2. `Provider`
3. `Action`
4. `ProviderAppointmentID`
5. `CalendarID`
6. `AppointmentTypeID`
7. `RawPayloadJSON`
8. `SignatureVerified`
9. `Status`
10. `Attempts`
11. `ProcessedAt`
12. `ResolvedUID`
13. `MasterRow`
14. `ResultJSON`
15. `Error`
16. `TestRunID`

---

## 5) End-to-End Flow

### 5.1 Acuity `scheduled`
1. Acuity sends webhook to relay.
2. Relay verifies `x-acuity-signature` with `HPAPP_ACUITY_WEBHOOK_SECRET`.
3. Relay appends a `PENDING` row to `_ExternalBookingEvents` with `Provider=acuity`.
4. Apps Script orchestrator calls `sw_processExternalBookingEvents`.
5. Processor fetches Acuity appointment detail (`ACUITY_USER_ID` + `ACUITY_API_KEY`), maps to form fields, submits Google Form.
6. Existing `onFormSubmit` flow writes/updates `00_Master Appointments`.

### 5.2 Acuity `rescheduled`
1. Same intake/queue process as above.
2. Processor reconciles old vs new visit time, marks prior row rescheduled/inactive as needed.
3. New active appointment path is submitted/resolved.
4. Linking fields (`RescheduledToUID`, `RescheduledFromUID`, etc.) are updated through existing logic.

### 5.3 Acuity `canceled`
1. Same intake/queue process.
2. Processor calls existing Acuity cancel-on-master behavior.
3. Active row is marked canceled/inactive; queue row status becomes `DONE` or `DONE_NO_ROW`.

### 5.4 Calendly (unchanged)
- Calendly webhook continues in its own Apps Script project and queue.
- It is intentionally not merged into this main repo flow.

---

## 6) Status Values You Will See

Common queue `Status` values:
- `PENDING`: waiting to process
- `RETRY`: temporary failure, will retry
- `ERROR_GAVE_UP`: failed too many times
- `DONE`: success
- `SKIPPED_DUP`: duplicate/already handled
- `SKIPPED_IGNORED_ACTION`: ignored action (example: `changed`)
- `DONE_NO_ROW`: cancel event had no active master row to close
- `DONE_NO_PRIOR`: reschedule had no prior row, handled as scheduled
- `DONE_NO_CHANGE`: existing row checked but no update needed

---

## 7) Trigger and Processing Cadence

- `sw_backgroundOrchestrator` is the single time-based worker (every 5 minutes).
- It runs `sw_processExternalBookingEvents` before fallback `acuityPollAndSubmit`.
- This order ensures webhook-backed events are handled first.

---

## 8) Verified Working Snapshot (May 6, 2026 PT)

Checks completed:
1. Relay unit tests passed:
   - accepts valid signature
   - rejects invalid signature
   - ignores `changed`
2. Queue spreadsheet configuration confirmed:
   - points to `[SYS] HPAPP Acuity Webhook Queue`
   - tab `_ExternalBookingEvents`
3. Synthetic queue test rows injected and processed successfully:
   - `changed` -> `SKIPPED_IGNORED_ACTION`
   - `canceled` (no matching master row) -> `DONE_NO_ROW`
4. Orchestrator trigger confirmed installed and running.

What this proves:
- Code path and queue processor are operational.

What this does not prove yet:
- Live Acuity production webhooks are reaching your deployed relay URL (needs one real webhook event test).

---

## 9) Live Verification Checklist (Production)

1. In Acuity webhook settings, set URL to:
   - `<deployed-base-url>/api/webhooks/acuity`
2. Create a real test appointment in Acuity.
3. Confirm queue row appears in `_ExternalBookingEvents` with `Provider=acuity`.
4. Wait one orchestrator cycle (up to 5 minutes) or run processor manually.
5. Confirm queue row leaves `PENDING` and reaches `DONE`/expected status.
6. Confirm downstream rows in:
   - `02_Form_Inbox`
   - `00_Master Appointments`
7. Run one reschedule and one cancel test the same way.

---

## 10) Fast Troubleshooting Playbook

### Symptom A: No row appears in `_ExternalBookingEvents`
Check:
1. Acuity webhook URL correctness.
2. Relay server is up and reachable.
3. `HPAPP_ACUITY_WEBHOOK_SECRET` on relay matches Acuity webhook secret exactly.
4. Service account permission on queue spreadsheet (Editor access).

### Symptom B: Row appears but stays `PENDING`
Check:
1. Orchestrator trigger exists and is healthy.
2. `sw_processExternalBookingEvents` execution logs.
3. Queue row `Provider` is `acuity` and `SignatureVerified` is true-like.

### Symptom C: Rows become `RETRY` or `ERROR_GAVE_UP`
Check:
1. `Error` column in queue row.
2. Missing Apps Script properties:
   - `FORM_ID`
   - `ACUITY_USER_ID`
   - `ACUITY_API_KEY`
   - queue spreadsheet ID property
3. Acuity API auth/access failures.
4. Form schema mismatch for required questions/options.

### Symptom D: Duplicates or unexpected skipped rows
Check:
1. Existing UID markers in Script Properties (`ACUITY:DONE:*`, `ACUITY:CANCELED:*`).
2. `SKIPPED_DUP` in queue status indicates dedupe worked.
3. Confirm event ordering for reschedule/cancel around same appointment.

---

## 11) High-Value Functions (Apps Script)

Operational:
- `sw_processExternalBookingEvents(options)`
- `sw_getHpAppAcuityQueueConfig()`
- `sw_getBackgroundOrchestratorStatus()`

Testing:
- `sw_testInjectExternalBookingEvent(options)`
- `sw_clearExternalBookingTestFlags(options)`

Cleanup:
- `sw_previewTestDataCleanupOnce()`
- `sw_applyTestDataCleanupOnce({ confirmationToken })`

---

## 12) Source Code Locations

Main queue processor:
- `sales_workflow_external_booking_events.js`

Orchestrator:
- `sales_workflow_orchestrator.js`

Queue schema constants:
- `sales_workflow_constants.js`

Acuity poller and form mapping:
- `Acuity_HPUSA.js`

Relay endpoint and signature verification:
- `internal/server/server.go`

Sheets append client (relay -> queue):
- `internal/server/google_sheets.go`

Config/env loading:
- `internal/config/config.go`

---

## 13) Security Notes

- Do not share raw webhook secret, Acuity API key, or service account private key in logs or docs.
- Keep `HPAPP_ACUITY_WEBHOOK_SECRET` consistent between Acuity webhook config and relay environment.
- Keep queue spreadsheet permission limited to required service accounts/users.
