/** Minimal no-op guards for retired installable triggers owned by other users.
 *
 * Apps Script only lets the installing account delete its own triggers. These
 * functions intentionally do nothing so old time-based triggers can fire
 * without rebuilding retired tabs or running legacy assignment logic.
 */

function runDailySetup() {
  Logger.log('SW_RETIRED_TRIGGER_NOOP runDailySetup');
}

function assignTodayAppointments() {
  Logger.log('SW_RETIRED_TRIGGER_NOOP assignTodayAppointments');
}

function timedRefreshHandler() {
  Logger.log('SW_RETIRED_TRIGGER_NOOP timedRefreshHandler');
}

function refreshClientStageRollup() {
  Logger.log('SW_RETIRED_TRIGGER_NOOP refreshClientStageRollup');
}
