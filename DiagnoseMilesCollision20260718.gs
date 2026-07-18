/**
 * DiagnoseMilesCollision20260718.gs
 * -----------------------------------------------------------------------
 * autoCreateApartmenteryBookings ran twice (6:39pm and 6:40pm) for
 * ABB-milesconse-20260609 (room 204, 2026-06-09 → 2026-06-16). Both times
 * it logged the same-day-turnover shrink of the outgoing booking
 * (321078, ABB-milesconse-20260603) to endDate=2026-06-08 as SUCCEEDING
 * (updateApartmenteryBookingEndDateForRoom didn't throw — it only throws
 * on a non-3xx response), then immediately hit apartmentery's "การจองนี้
 * ชนกับการจองอื่น" (booking collision) error creating the new booking
 * anyway. Two tries a minute apart ruling out simple async lag.
 *
 * READ-ONLY. Checks two things directly:
 *   1. What 321078's endDate actually reads as on apartmentery RIGHT NOW
 *      (did the shrink really persist, or did apartmentery accept the
 *      POST but silently keep the old value — e.g. because startDate
 *      wasn't included correctly, or some other field in the resubmitted
 *      form failed silently)
 *   2. Every event in room 204's calendar with a start or end date
 *      anywhere near 2026-06-09 (±3 days), in case the true colliding
 *      booking isn't 321078 at all — a leftover/duplicate booking
 *      neither Sheet1 nor the earlier audit caught.
 *
 * HOW TO RUN: Apps Script editor ▶ diagnoseMilesCollision20260718 ▶ Run ▶
 * read log.
 */

function diagnoseMilesCollision20260718() {
  const room = '204';
  const unit = getApartmenteryUnitForRoom(room);
  if (!unit) {
    Logger.log(`Room ${room} not found in ROOM_TO_UNIT_ID — aborting.`);
    return;
  }

  Logger.log(`--- Current apartmentery state of booking 321078 (should be ABB-milesconse-20260603, shrunk endDate expected 2026-06-08) ---`);
  try {
    const state = _getApartmenteryBookingEditFormState_(unit.branchId, unit.unitId, '321078');
    Logger.log(`  startDate=${state.startDate} endDate=${state.endDate} customerName="${state.customerName}" ` +
      `${state.endDate === '2026-06-08' ? '[SHRINK PERSISTED — matches expected 2026-06-08]' : '[SHRINK DID NOT PERSIST — still showing old value, expected 2026-06-08]'}`);
  } catch (e) {
    Logger.log(`  Could not load booking 321078's edit form: ${e.message}`);
  }

  Logger.log(`--- All room ${room} calendar events with start or end near 2026-06-09 (±3 days) ---`);
  const path = `/user/branch/${unit.branchId}/unit/${unit.unitId}/booking`;
  const response = _apartmenteryFetch_(path, { method: 'get' });
  const html = response.getContentText();
  const blockRe = /\{\s*title:\s*'((?:[^'\\]|\\.)*)'[\s\S]*?start:\s*'([^']*)'[\s\S]*?end:\s*'([^']*)'[\s\S]*?url:\s*'([^']*)'\s*\}/g;
  let m;
  const windowStart = new Date('2026-06-06T00:00:00Z');
  const windowEnd = new Date('2026-06-12T00:00:00Z');
  let found = 0;
  while ((m = blockRe.exec(html)) !== null) {
    const start = _apartmenteryCalendarDateToIso_(m[2]);
    const end = m[3] ? _apartmenteryCalendarDateToIso_(m[3]) : '';
    const startDt = new Date(start + 'T00:00:00Z');
    if (startDt >= windowStart && startDt <= windowEnd) {
      const idMatch = m[4].match(/\/booking\/(\d+)/);
      Logger.log(`  bookingId=${idMatch ? idMatch[1] : '?'} title="${m[1]}" start=${start} end=${end || '(none captured)'}`);
      found++;
    }
  }
  if (found === 0) Logger.log(`  No events found in that window via the 'end:' capture — the calendar events may not include an end field in the raw HTML block (only start+url were used elsewhere in this codebase). See raw dump below instead.`);

  Logger.log(`--- Same window, using the same start+url-only pattern the rest of the codebase relies on (in case 'end:' isn't actually present in the HTML) ---`);
  const blockRe2 = /\{\s*title:\s*'((?:[^'\\]|\\.)*)'[\s\S]*?start:\s*'([^']*)'[\s\S]*?url:\s*'([^']*)'\s*\}/g;
  let m2;
  while ((m2 = blockRe2.exec(html)) !== null) {
    const start = _apartmenteryCalendarDateToIso_(m2[2]);
    const startDt = new Date(start + 'T00:00:00Z');
    if (startDt >= windowStart && startDt <= windowEnd) {
      const idMatch = m2[3].match(/\/booking\/(\d+)/);
      Logger.log(`  bookingId=${idMatch ? idMatch[1] : '?'} title="${m2[1]}" start=${start}`);
    }
  }
}
