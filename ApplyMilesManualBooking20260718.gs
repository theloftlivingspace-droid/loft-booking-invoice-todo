/**
 * ApplyMilesManualBooking20260718.gs
 * -----------------------------------------------------------------------
 * Nathan created ABB-milesconse-20260609's apartmentery booking manually
 * (id 327142, https://apartmentery.com/user/branch/6801/unit/163864/
 * booking/327142) after 2 automated attempts kept hitting a same-day-
 * turnover collision even though the outgoing booking's endDate shrink
 * appeared to succeed both times (see DiagnoseMilesCollision20260718.gs
 * — never got to run since Nathan resolved it manually first, root cause
 * still unconfirmed).
 *
 * Writes the id to Sheet1 and marks the row done, same as
 * autoCreateApartmenteryBookings does after a successful automated
 * create, so this row stops showing up in getBookingToAdd_ and doesn't
 * get retried next hourly run.
 *
 * HOW TO RUN: Apps Script editor ▶ applyMilesManualBooking20260718 ▶ Run
 * ▶ read log.
 */

function applyMilesManualBooking20260718() {
  const resId = 'ABB-milesconse-20260609';
  const bookingId = '327142';

  const writeResult = setApartmenteryBookingId_(resId, bookingId);
  if (!writeResult.ok) {
    Logger.log(`FAILED to write ${bookingId} for ${resId}: ${writeResult.error} — needs manual look, this shouldn't conflict with anything.`);
    return;
  }
  setBookingDone(resId, true);
  Logger.log(`DONE — ${resId}: apartmentery bookingId set to ${bookingId} (manually created by Nathan), marked done.`);
}
