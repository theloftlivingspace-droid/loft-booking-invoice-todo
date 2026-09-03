/**
 * DiagnoseSyncErrors20260903.gs
 * -----------------------------------------------------------------------
 * Investigates the two errors the automation reported on 2026-09-03,
 * both plausibly downstream of the dataRows/email_log duplicate-resId
 * incident fixed the same day in hotel-line-bot (see that repo's commits
 * ab54546, 0e0d8b1, bbfeb34 and this repo's memory notes).
 *
 * CASE A — THANAPORNPAN BUKBOON, room 204:
 *   setApartmenteryBookingId_ refused to write bookingId 321937 for
 *   TRP-thanapornp-20260609 because it's already linked to a different
 *   resId, TRP-thanapornp-20260616, in Sheet1. resIds here are built from
 *   guestKey + the DATE THE EMAIL ARRIVED, not a stable reservation
 *   reference — so if Little Hotelier/Trip.com sent two separate emails
 *   for the same underlying reservation on two different calendar days,
 *   this system would create two distinct Sheet1 rows for one real
 *   booking. This function pulls both rows from Sheet1 and prints them
 *   side by side so it's obvious whether they're the same reservation
 *   (guest+checkIn+checkOut+room all match → dup, safe to delete the
 *   0609 row) or genuinely different bookings that happen to collide
 *   (needs the real fix elsewhere).
 *
 * CASE B — AR Nonthanan, room 108, 2026-05-11 → 2026-05-12:
 *   Apartmentery itself rejected booking creation with HTTP 400
 *   ("การจองนี้ชนกับการจองอื่น") — a real overlap on apartmentery's own
 *   calendar, not just a Sheet1 issue. This function dumps every event
 *   apartmentery's room 108 calendar shows in a window around those
 *   dates, so we can see whether there are two overlapping entries there
 *   (leftover from the EXP-arnonthana-20260511-R20260511 duplicate that
 *   was cleaned out of Sheet1/email_log, but not necessarily ever
 *   created on apartmentery — or possibly WAS created there and needs
 *   deleting directly on apartmentery.com).
 *
 * READ-ONLY. Does not write to Sheet1 or apartmentery.
 *
 * HOW TO RUN: Apps Script editor ▶ diagnoseSyncErrors20260903 ▶ Run ▶
 * read log (View ▶ Logs, or Ctrl+Enter).
 */

function diagnoseSyncErrors20260903() {
  _diagnoseCaseA_Thanapornp20260903_();
  Logger.log('');
  _diagnoseCaseB_ArNonthanan20260903_();
}

function _diagnoseCaseA_Thanapornp20260903_() {
  Logger.log('=== CASE A: THANAPORNPAN BUKBOON — TRP-thanapornp-20260609 vs -20260616 ===');
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  const data = src.getDataRange().getValues();
  const header = data[0];
  const idx = indexMap_(header, ['ResId', 'ชื่อแขก', 'เช็คอิน', 'เช็คเอาท์', 'เลขห้อง', APARTMENTERY_BOOKING_ID_COL_HEADER]);

  const targets = ['TRP-thanapornp-20260609', 'TRP-thanapornp-20260616'];
  const rowsFound = [];
  for (let i = 1; i < data.length; i++) {
    const resId = String(data[i][idx.ResId] || '').trim();
    if (targets.indexOf(resId) === -1) continue;
    const row = {
      resId: resId,
      guest: data[i][idx['ชื่อแขก']],
      checkIn: data[i][idx['เช็คอิน']],
      checkOut: data[i][idx['เช็คเอาท์']],
      room: data[i][idx['เลขห้อง']],
      bookingId: data[i][idx[APARTMENTERY_BOOKING_ID_COL_HEADER]],
      sheetRow: i + 1
    };
    rowsFound.push(row);
    Logger.log(`  Sheet1 row ${row.sheetRow}: resId=${row.resId} guest="${row.guest}" ` +
      `checkIn=${row.checkIn} checkOut=${row.checkOut} room="${row.room}" bookingId=${row.bookingId || '(none)'}`);
  }

  if (rowsFound.length === 2) {
    const [a, b] = rowsFound;
    const sameGuest = String(a.guest).trim() === String(b.guest).trim();
    const sameStay = String(a.checkIn) === String(b.checkIn) && String(a.checkOut) === String(b.checkOut);
    const sameRoom = String(a.room).trim() === String(b.room).trim();
    if (sameGuest && sameStay && sameRoom) {
      Logger.log(`  → SAME RESERVATION (guest/dates/room all match). This is a duplicate row from two ` +
        `separate emails for one booking, not two real reservations. Safe to delete the row WITHOUT ` +
        `a bookingId (${(a.bookingId ? b : a).resId}) from Sheet1 — keep the one Apartmentery already ` +
        `has linked (${(a.bookingId ? a : b).resId}, bookingId=${a.bookingId || b.bookingId}).`);
    } else {
      Logger.log(`  → NOT an obvious duplicate (guest match=${sameGuest}, stay match=${sameStay}, ` +
        `room match=${sameRoom}). These may be two genuinely different reservations that happen to ` +
        `collide on the apartmentery lookup — do not delete either row without checking apartmentery ` +
        `directly (see CASE B pattern below, applied to room 204 / bookingId 321937 / unit ` +
        `${JSON.stringify(getApartmenteryUnitForRoom('204'))}).`);
    }
  } else {
    Logger.log(`  → Expected both resIds in Sheet1, found ${rowsFound.length}. If only one row exists now, ` +
      `someone may have already cleaned this up — re-check whether the Apartmentery error is stale.`);
  }
}

function _diagnoseCaseB_ArNonthanan20260903_() {
  Logger.log('=== CASE B: AR Nonthanan — room 108 calendar around 2026-05-11 → 2026-05-12 ===');
  const room = '108';
  const unit = getApartmenteryUnitForRoom(room);
  if (!unit) {
    Logger.log(`  Room ${room} not found in ROOM_TO_UNIT_ID — aborting.`);
    return;
  }

  const path = `/user/branch/${unit.branchId}/unit/${unit.unitId}/booking`;
  const response = _apartmenteryFetch_(path, { method: 'get' });
  const html = response.getContentText();

  const windowStart = new Date('2026-05-08T00:00:00Z');
  const windowEnd = new Date('2026-05-15T00:00:00Z');

  Logger.log(`  Events with start+end captured, ${windowStart.toISOString().slice(0,10)}..${windowEnd.toISOString().slice(0,10)}:`);
  const blockRe = /\{\s*title:\s*'((?:[^'\\]|\\.)*)'[\s\S]*?start:\s*'([^']*)'[\s\S]*?end:\s*'([^']*)'[\s\S]*?url:\s*'([^']*)'\s*\}/g;
  let m, found = 0;
  while ((m = blockRe.exec(html)) !== null) {
    const start = _apartmenteryCalendarDateToIso_(m[2]);
    const end = m[3] ? _apartmenteryCalendarDateToIso_(m[3]) : '';
    const startDt = new Date(start + 'T00:00:00Z');
    if (startDt >= windowStart && startDt <= windowEnd) {
      const idMatch = m[4].match(/\/booking\/(\d+)/);
      Logger.log(`    bookingId=${idMatch ? idMatch[1] : '?'} title="${m[1]}" start=${start} end=${end || '(none)'}`);
      found++;
    }
  }

  Logger.log(`  Same window, start+url-only fallback pattern (in case 'end:' isn't in the raw HTML):`);
  const blockRe2 = /\{\s*title:\s*'((?:[^'\\]|\\.)*)'[\s\S]*?start:\s*'([^']*)'[\s\S]*?url:\s*'([^']*)'\s*\}/g;
  let m2, found2 = 0;
  while ((m2 = blockRe2.exec(html)) !== null) {
    const start = _apartmenteryCalendarDateToIso_(m2[2]);
    const startDt = new Date(start + 'T00:00:00Z');
    if (startDt >= windowStart && startDt <= windowEnd) {
      const idMatch = m2[3].match(/\/booking\/(\d+)/);
      Logger.log(`    bookingId=${idMatch ? idMatch[1] : '?'} title="${m2[1]}" start=${start}`);
      found2++;
    }
  }

  if (found === 0 && found2 === 0) {
    Logger.log(`  → No events found in this window at all. That means apartmentery has NO booking for ` +
      `room 108 covering 2026-05-11/12 right now — so either the HTTP 400 was itself stale/cached, or ` +
      `whatever it collided with sits outside this ±3-4 day window (widen windowStart/windowEnd above and re-run).`);
  } else if (found > 1 || found2 > 1) {
    Logger.log(`  → MULTIPLE events found in this window — likely the real overlap apartmentery is ` +
      `rejecting against. Compare titles/bookingIds above: if one is a leftover from the ` +
      `EXP-arnonthana-20260511-R20260511 duplicate, delete IT directly on apartmentery.com (Sheet1/` +
      `email_log cleanup alone does not remove anything already created on apartmentery's side).`);
  } else {
    Logger.log(`  → Exactly one event found — that's presumably the legitimate AR Nonthanan booking. ` +
      `If the 400 is still happening on retry, the collision may be against a booking in an adjacent ` +
      `room's calendar being misread, or a stale cache on apartmentery's end — re-run this after a few ` +
      `minutes before assuming it's a leftover duplicate.`);
  }
}
