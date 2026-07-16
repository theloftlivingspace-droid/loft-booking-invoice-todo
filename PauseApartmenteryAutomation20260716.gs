/**
 * PauseApartmenteryAutomation20260716.gs
 * -----------------------------------------------------------------------
 * Removes the runApartmenteryAutomation hourly trigger (installed by
 * installApartmenteryAutomationTrigger in ApartmenteryAutomation.gs).
 *
 * Reason: auditAllApartmenteryBookingIds20260716 found 84 of 132 rows in
 * Sheet1 have a WRONG "Apartmentery Booking ID" (a real id belonging to a
 * different booking) and 19 more have none at all. Since
 * autoCreateApartmenteryInvoicesAndReceipts() uses that column's value to
 * create invoices against apartmentery, letting the hourly trigger keep
 * running risks creating more invoices against the wrong booking while
 * we're still investigating the blast radius. Pause first, resume only
 * after the Sheet1 column is corrected and the invoice impact is checked.
 *
 * Run this once from the Apps Script editor. Safe to run more than once
 * (no-op if the trigger is already gone). To resume automation later,
 * call installApartmenteryAutomationTrigger() again (in
 * ApartmenteryAutomation.gs) once everything's confirmed clean.
 */
function pauseApartmenteryAutomation20260716() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'runApartmenteryAutomation') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  Logger.log(removed > 0
    ? `Removed ${removed} runApartmenteryAutomation trigger(s). Automation is now PAUSED — ` +
      `no more bookings or invoices will be auto-created until installApartmenteryAutomationTrigger() is called again.`
    : `No runApartmenteryAutomation trigger found — already paused (or was never installed).`);
  return { removed: removed };
}
