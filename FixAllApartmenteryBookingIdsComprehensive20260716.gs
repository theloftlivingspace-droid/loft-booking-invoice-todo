/**
 * FixAllApartmenteryBookingIdsComprehensive20260716.gs
 * -----------------------------------------------------------------------
 * The earlier one-off scripts today (checkInvoiceImpact20260716,
 * fixDuplicateApartmenteryIds20260716) worked from a small HARDCODED list
 * — 11 and 18 resIds respectively, pasted in manually from a partial read
 * of the audit log. The full audit (auditAllApartmenteryBookingIds /
 * auditAllApartmenteryBookingIds20260716) actually found 99 "wrong room"
 * rows plus 9 guest-name mismatches — this covers ALL of them, driven
 * live from Sheet1 + apartmentery each time it runs, so nothing has to be
 * copy-pasted or kept in sync by hand.
 *
 * Root cause (already fixed going forward in ApartmenteryClient.gs,
 * commit d674e49): the booking-creation fallback used to match a newly
 * created booking's id by guest name ALONE when apartmentery's redirect
 * didn't include it — with no date check, it could grab a different
 * booking's id whenever names/rooms repeated (very likely during the
 * ~130-booking historical backfill). This corrupted the Apartmentery
 * Booking ID column across many unrelated rows, chaining forward from
 * ~2026-03-27 through ~2026-07-04.
 *
 * PART 1 — checkInvoiceImpactComprehensive20260716() [READ-ONLY]
 *   For every mismatched row, checks whether the WRONG (currently-stored)
 *   bookingId already shows signs of having an invoice attached on
 *   apartmentery. If so, our automation likely created a real invoice
 *   against the WRONG guest's booking — flag for manual review before
 *   fixing anything, since correcting the ID afterward does NOT undo an
 *   already-created wrong invoice.
 *
 * PART 2 — fixAllApartmenteryBookingIdsComprehensive20260716() [WRITES]
 *   For every mismatched row, independently re-looks-up the correct id
 *   via findApartmenteryBookingIdForRoomByGuest_ (date-verified, doesn't
 *   care what's currently stored) and overwrites Sheet1 only on a
 *   confident match that differs from what's stored. Never guesses —
 *   unresolved rows are only logged.
 *
 * Run PART 1 first. Review its output. Only run PART 2 after Nathan has
 * reviewed (or accepted the risk on) whatever PART 1 flags.
 *
 * HOW TO RUN: Apps Script editor ▶ pick the function ▶ Run ▶ read log.
 */

/**
 * Shared by both parts: builds the full list of mismatched rows straight
 * from Sheet1 + a fresh pull of every room's apartmentery calendar. Same
 * logic as auditAllApartmenteryBookingIds, factored out so both the
 * impact-check and the fix draw from one live source instead of two
 * separately-maintained copies.
 */
function _getAllMismatchedApartmenteryRows20260716_() {
  const byBookingId = {};
  Object.keys(ROOM_TO_UNIT_ID).forEach(room => {
    const unit = getApartmenteryUnitForRoom(room);
    if (!unit) return;
    const path = `/user/branch/${unit.branchId}/unit/${unit.unitId}/booking`;
    let html;
    try {
      const response = _apartmenteryFetch_(path, { method: 'get' });
      html = response.getContentText();
    } catch (e) {
      Logger.log(`_getAllMismatchedApartmenteryRows20260716_: FAILED to fetch calendar for room ${room}: ${e.message}`);
      return;
    }
    const blockRe = /\{\s*title:\s*'((?:[^'\\]|\\.)*)'[\s\S]*?start:\s*'([^']*)'[\s\S]*?url:\s*'([^']*)'\s*\}/g;
    let m;
    while ((m = blockRe.exec(html)) !== null) {
      const idMatch = m[3].match(/\/booking\/(\d+)/);
      if (!idMatch) continue;
      const bookingId = idMatch[1];
      if (!byBookingId[bookingId]) {
        byBookingId[bookingId] = { room: room, title: m[1], start: _apartmenteryCalendarDateToIso_(m[2]) };
      }
    }
  });

  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName('Sheet1');
  const data = src.getDataRange().getValues();
  const header = data[0];
  const idx = indexMap_(header, ['ResId', 'เลขห้อง', 'ชื่อแขก', 'เช็คอิน', APARTMENTERY_BOOKING_ID_COL_HEADER]);

  const mismatched = [];
  for (let i = 1; i < data.length; i++) {
    const resId = String(data[i][idx.ResId] || '').trim();
    const storedId = String(data[i][idx[APARTMENTERY_BOOKING_ID_COL_HEADER]] || '').trim();
    if (!resId || !storedId) continue;
    const roomRaw = String(data[i][idx['เลขห้อง']] || '').trim();
    const guest = String(data[i][idx['ชื่อแขก']] || '').trim();
    const checkin = formatCellDate_(data[i][idx['เช็คอิน']]);
    const expectedRoom = (roomRaw.match(/^\d+/) || [''])[0];

    const found = byBookingId[storedId];
    const isDead = !found;
    const isWrongRoom = found && found.room !== expectedRoom;
    const isGuestMismatch = found && found.room === expectedRoom && guest && !_namesMatchIgnoringOrder_(guest, found.title);
    if (isDead || isWrongRoom || isGuestMismatch) {
      mismatched.push({
        resId, room: roomRaw, expectedRoom, guest, checkin, storedId,
        foundUnderRoom: found ? found.room : null,
        foundTitle: found ? found.title : null,
        kind: isDead ? 'DEAD' : (isWrongRoom ? 'WRONG_ROOM' : 'GUEST_MISMATCH')
      });
    }
  }
  return mismatched;
}

function checkInvoiceImpactComprehensive20260716() {
  const mismatched = _getAllMismatchedApartmenteryRows20260716_();
  Logger.log(`checkInvoiceImpactComprehensive20260716: ${mismatched.length} mismatched row(s) to check.`);

  const flagged = [];
  mismatched.forEach(row => {
    // Only WRONG_ROOM / GUEST_MISMATCH have a real (but wrong) booking to
    // check — DEAD ids don't point anywhere, so there's nothing to have
    // attached an invoice to.
    if (row.kind === 'DEAD') return;
    const unit = getApartmenteryUnitForRoom(row.foundUnderRoom);
    if (!unit) return;
    const path = `/user/branch/${unit.branchId}/unit/${unit.unitId}/booking/${row.storedId}/edit`;
    try {
      const response = _apartmenteryFetch_(path, { method: 'get' });
      if (response.getResponseCode() !== 200) {
        Logger.log(`${row.resId} (wrong id ${row.storedId}, real room ${row.foundUnderRoom}): HTTP ${response.getResponseCode()} — skipping.`);
        return;
      }
      const html = response.getContentText();
      const invoiceLinks = (html.match(/\/invoice\/\d+/g) || []);
      if (invoiceLinks.length > 0) {
        Logger.log(`⚠️ ${row.resId} (Sheet1 guest "${row.guest}", room ${row.expectedRoom}) stored WRONG id ` +
          `${row.storedId} which really belongs to "${row.foundTitle}" (room ${row.foundUnderRoom}) — ` +
          `that booking HAS invoice link(s): ${JSON.stringify(invoiceLinks)}. Possible wrong-guest invoice — review manually.`);
        flagged.push(Object.assign({ invoiceLinks }, row));
      }
    } catch (err) {
      if (isApartmenterySessionExpiredError(err)) { Logger.log('SESSION EXPIRED — stopping.'); throw err; }
      Logger.log(`${row.resId}: error checking invoice impact — ${err.message}`);
    }
  });

  Logger.log(`checkInvoiceImpactComprehensive20260716: ${flagged.length} row(s) flagged with possible wrong-guest invoices.`);
  Logger.log(JSON.stringify(flagged, null, 2));
  return flagged;
}

function fixAllApartmenteryBookingIdsComprehensive20260716() {
  const mismatched = _getAllMismatchedApartmenteryRows20260716_();
  Logger.log(`fixAllApartmenteryBookingIdsComprehensive20260716: ${mismatched.length} mismatched row(s) to fix.`);

  const report = { fixed: [], unresolved: [], unchanged: [] };
  mismatched.forEach(row => {
    let foundId = null;
    try {
      foundId = findApartmenteryBookingIdForRoomByGuest_(row.room, row.guest, row.checkin);
    } catch (err) {
      if (isApartmenterySessionExpiredError(err)) { Logger.log('SESSION EXPIRED — stopping. Refresh APARTMENTERY_SESSION and re-run.'); throw err; }
      Logger.log(`lookup failed for ${row.resId} (${row.guest}, ${row.room}, ${row.checkin}): ${err.message}`);
      report.unresolved.push({ resId: row.resId, guest: row.guest, reason: err.message });
      return;
    }
    if (!foundId) {
      Logger.log(`NO MATCH for ${row.resId} (${row.guest}, ${row.room}, ${row.checkin}) — current stored id: "${row.storedId}". Check apartmentery manually.`);
      report.unresolved.push({ resId: row.resId, guest: row.guest, currentId: row.storedId });
      return;
    }
    if (foundId === row.storedId) {
      report.unchanged.push({ resId: row.resId, guest: row.guest, id: row.storedId });
      return;
    }
    // allowOverwriteConflict: true — this sweep independently re-verifies
    // foundId against apartmentery's live calendar (room + guest + date),
    // so it's authoritative even if another not-yet-processed row in this
    // same sweep still holds foundId as its (wrong) stored value. The
    // uniqueness guard in setApartmenteryBookingId_ exists to stop NEW
    // wrong ids from the creation/recovery path, not to block a verified
    // correction here.
    setApartmenteryBookingId_(row.resId, foundId, { allowOverwriteConflict: true });
    Logger.log(`FIXED — ${row.resId} (${row.guest}, ${row.room}): "${row.storedId}" -> "${foundId}"`);
    report.fixed.push({ resId: row.resId, guest: row.guest, room: row.room, from: row.storedId, to: foundId });
  });

  Logger.log('fixAllApartmenteryBookingIdsComprehensive20260716 report: ' + JSON.stringify(report, null, 2));
  Logger.log(`SUMMARY: fixed=${report.fixed.length} unresolved=${report.unresolved.length} unchanged=${report.unchanged.length}`);
  return report;
}
