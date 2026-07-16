/**
 * AuditAllApartmenteryBookingIds20260716.gs
 * -----------------------------------------------------------------------
 * READ-ONLY audit — does NOT write to Sheet1 or apartmentery.
 *
 * Room 214's raw calendar (see DebugRoom214Calendar20260716) proved that
 * stale/wrong Apartmentery Booking ID values exist in Sheet1 beyond the 9
 * pairs originally found (e.g. Stanley Modjadji's stored id 325354 is
 * actually Luke Faisal's — Stanley's real id is 324791). These are
 * leftovers from before today's guest-matching/date-format fixes landed.
 *
 * This fetches each distinct room's apartmentery calendar ONCE, then
 * checks every non-cancelled Sheet1 row for that room: does a calendar
 * event exist whose title matches the guest name (word-order independent)
 * AND whose start date equals the row's checkin? Reports MATCH / MISMATCH
 * (stored id differs from the one found) / NOT_FOUND (no matching
 * calendar entry at all — could mean booking never created, or a name/date
 * mismatch worth a manual look) for every row. Nothing is written back —
 * review the report, then we decide together what (if anything) to fix.
 *
 * Run debugRoom214Calendar20260716-style manually from the Apps Script
 * editor; check the execution log for the full report.
 */
function auditAllApartmenteryBookingIds20260716() {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  if (!src) { Logger.log('Sheet not found: ' + SRC_BOOKING_SHEET); return; }

  const data = src.getDataRange().getValues();
  const header = data[0];
  const idx = indexMap_(header, ['เลขห้อง', 'ชื่อแขก', 'เช็คอิน', 'ResId', APARTMENTERY_BOOKING_ID_COL_HEADER]);

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const room = String(data[i][idx['เลขห้อง']] || '').trim();
    const guest = String(data[i][idx['ชื่อแขก']] || '').trim();
    const checkinRaw = data[i][idx['เช็คอิน']];
    const resId = String(data[i][idx['ResId']] || '').trim();
    const storedId = String(data[i][idx[APARTMENTERY_BOOKING_ID_COL_HEADER]] || '').trim();
    if (!resId || !room) continue;
    if (/ยกเลิก|cancel/i.test(room)) continue; // cancelled rows never need a real id — skip
    const checkin = checkinRaw instanceof Date
      ? Utilities.formatDate(checkinRaw, 'Asia/Bangkok', 'yyyy-MM-dd')
      : String(checkinRaw || '').trim();
    rows.push({ room: room, guest: guest, checkin: checkin, resId: resId, storedId: storedId });
  }

  Logger.log(`auditAllApartmenteryBookingIds20260716: ${rows.length} non-cancelled rows to check.`);

  // Group by room number so each room's calendar is fetched exactly once.
  const byRoomNum = {};
  rows.forEach(r => {
    const rn = roomNum_(r.room);
    if (!byRoomNum[rn]) byRoomNum[rn] = [];
    byRoomNum[rn].push(r);
  });

  const report = { match: [], mismatch: [], notFound: [], noCalendarAccess: [] };

  Object.keys(byRoomNum).forEach(rn => {
    const unit = getApartmenteryUnitForRoom(rn);
    if (!unit) {
      Logger.log(`Room ${rn} not in ROOM_TO_UNIT_ID — skipping ${byRoomNum[rn].length} row(s).`);
      byRoomNum[rn].forEach(r => report.noCalendarAccess.push(r.resId));
      return;
    }

    let events = [];
    try {
      const path = `/user/branch/${unit.branchId}/unit/${unit.unitId}/booking`;
      const response = _apartmenteryFetch_(path, { method: 'get' });
      if (response.getResponseCode() !== 200) {
        Logger.log(`Room ${rn}: calendar fetch returned HTTP ${response.getResponseCode()} — skipping.`);
        byRoomNum[rn].forEach(r => report.noCalendarAccess.push(r.resId));
        return;
      }
      const html = response.getContentText();
      const blockRe = /\{\s*title:\s*'((?:[^'\\]|\\.)*)'[\s\S]*?start:\s*'([^']*)'[\s\S]*?url:\s*'([^']*)'\s*\}/g;
      let m;
      while ((m = blockRe.exec(html)) !== null) {
        const idMatch = m[3].match(/\/booking\/(\d+)/);
        if (!idMatch) continue;
        events.push({ title: m[1], isoStart: _apartmenteryCalendarDateToIso_(m[2]), id: idMatch[1] });
      }
    } catch (err) {
      if (isApartmenterySessionExpiredError(err)) {
        Logger.log('SESSION EXPIRED — stopping audit. Refresh APARTMENTERY_SESSION and re-run.');
        throw err;
      }
      Logger.log(`Room ${rn}: fetch error (${err.message}) — skipping.`);
      byRoomNum[rn].forEach(r => report.noCalendarAccess.push(r.resId));
      return;
    }

    byRoomNum[rn].forEach(r => {
      const exact = events.find(e => e.title.indexOf(r.guest) !== -1 && e.isoStart === r.checkin);
      const loose = exact || events.find(e => _namesMatchIgnoringOrder_(r.guest, e.title) && e.isoStart === r.checkin);
      if (!loose) {
        Logger.log(`NOT_FOUND — ${r.resId} (${r.guest}, room ${r.room}, checkin ${r.checkin}): ` +
          `no calendar entry matches. Currently stored: "${r.storedId}".`);
        report.notFound.push({ resId: r.resId, guest: r.guest, room: r.room, checkin: r.checkin, storedId: r.storedId });
        return;
      }
      if (loose.id === r.storedId) {
        report.match.push(r.resId);
      } else {
        Logger.log(`MISMATCH — ${r.resId} (${r.guest}, room ${r.room}, checkin ${r.checkin}): ` +
          `stored "${r.storedId}" but calendar says "${loose.id}" (title: "${loose.title}").`);
        report.mismatch.push({ resId: r.resId, guest: r.guest, room: r.room, checkin: r.checkin,
          storedId: r.storedId, realId: loose.id });
      }
    });
  });

  Logger.log('=== SUMMARY ===');
  Logger.log(`match: ${report.match.length}, mismatch: ${report.mismatch.length}, ` +
    `notFound: ${report.notFound.length}, noCalendarAccess: ${report.noCalendarAccess.length}`);
  Logger.log('Full mismatch list: ' + JSON.stringify(report.mismatch, null, 2));
  Logger.log('Full notFound list: ' + JSON.stringify(report.notFound, null, 2));
  return report;
}
