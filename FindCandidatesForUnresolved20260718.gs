/**
 * FindCandidatesForUnresolved20260718.gs
 * -----------------------------------------------------------------------
 * fixAllApartmenteryBookingIdsComprehensive20260716's lookup
 * (findApartmenteryBookingIdForRoomByGuest_) only searches the ONE room
 * Sheet1 says the guest should be in, and only accepts an EXACT checkin-
 * date match. All 14 remaining mismatches from the 2026-07-18 audit came
 * back "NO MATCH" under that rule — meaning each guest's real apartmentery
 * booking is either sitting under a different room than Sheet1 expects
 * (consistent with the WRONG_ROOM pattern the audit found: ids get
 * cross-assigned between rooms), or was made/cancelled on a slightly
 * different date than what's stored in Sheet1.
 *
 * READ-ONLY. Pulls every room's calendar ONCE (not per-guest — 11 fetches
 * total instead of 11x14), then for each unresolved row does a loose
 * name-only search across ALL rooms/dates and prints every candidate
 * found, flagging whether that candidate's bookingId is already claimed
 * by a different resId elsewhere in Sheet1 (so a claimed one is NOT a
 * safe pick — it belongs to that other, already-fixed booking).
 *
 * Nothing here writes anything. After reviewing the candidates, apply
 * confirmed-safe ones the same way 2026-07-16's applyManuallyVerifiedIds
 * scripts did — one resId/id pair at a time, only for candidates that are
 * (a) unclaimed and (b) the single obvious match for that guest.
 *
 * HOW TO RUN: Apps Script editor ▶ findCandidatesForUnresolved20260718 ▶
 * Run ▶ read log.
 */

const UNRESOLVED_20260718_ = [
  { resId: 'ABB-e585a82c20-20260219', guest: '全, 桂珍', room: '103', checkin: '2026-02-19', storedId: '305941' },
  { resId: 'ABB-e5a698e88a-20260403', guest: '妘芮 林 Yunjui Lin', room: '103', checkin: '2026-04-03', storedId: '312166' },
  { resId: 'ABB-johnzambra-20260428', guest: 'John Zambrana', room: '108', checkin: '2026-04-28', storedId: '315889' },
  { resId: 'ABB-premmehta-20260501', guest: 'Prem Mehta', room: '108', checkin: '2026-05-01', storedId: '315890' },
  { resId: 'ABB-lataviaant-20260512', guest: "La'Tavia Antrice", room: '214', checkin: '2026-05-12', storedId: '317341' },
  { resId: 'ABB-kgotlellom-20260528', guest: 'Kgotlello Masemola', room: '203', checkin: '2026-05-27', storedId: '317454' },
  { resId: 'TRP-arayaratta-20260518', guest: 'Araya Rattanabamrung', room: '103', checkin: '2026-05-18', storedId: '317749' },
  { resId: 'ABB-dogukankan-20260531', guest: 'Dogukan Kaner', room: '214', checkin: '2026-05-31', storedId: '321157' },
  { resId: 'ABB-errolcox-20260608', guest: 'Errol Cox', room: '210', checkin: '2026-06-08', storedId: '321078' },
  { resId: 'ABB-milesconse-20260609', guest: 'Miles Consengco', room: '204', checkin: '2026-06-09', storedId: '321937' },
  { resId: 'ABB-saeidmickm-20260610', guest: 'Saeid Mick Momtahan', room: '108', checkin: '2026-06-10', storedId: '322314' },
  { resId: 'TRP-thanapornp-20260616', guest: 'THANAPORNPAN BUKBOON', room: '204', checkin: '2026-06-18', storedId: '322188' },
  { resId: 'TRP-pornpawitb-20260616', guest: 'Pornpawit Boon', room: '103', checkin: '2026-06-16', storedId: '322519' },
  { resId: 'ABB-syeedryan-20260717', guest: 'Syeed Ryan', room: '300', checkin: '2026-07-24', storedId: '326968' }
];

/**
 * Pulls every room's calendar once and returns a flat list of all events:
 * { bookingId, room, title, start }.
 */
function _pullAllApartmenteryEvents20260718_() {
  const events = [];
  Object.keys(ROOM_TO_UNIT_ID).forEach(room => {
    const unit = getApartmenteryUnitForRoom(room);
    if (!unit) return;
    const path = `/user/branch/${unit.branchId}/unit/${unit.unitId}/booking`;
    let html;
    try {
      const response = _apartmenteryFetch_(path, { method: 'get' });
      html = response.getContentText();
    } catch (e) {
      Logger.log(`_pullAllApartmenteryEvents20260718_: FAILED to fetch calendar for room ${room}: ${e.message}`);
      return;
    }
    const blockRe = /\{\s*title:\s*'((?:[^'\\]|\\.)*)'[\s\S]*?start:\s*'([^']*)'[\s\S]*?url:\s*'([^']*)'\s*\}/g;
    let m;
    while ((m = blockRe.exec(html)) !== null) {
      const idMatch = m[3].match(/\/booking\/(\d+)/);
      if (!idMatch) continue;
      events.push({
        bookingId: idMatch[1],
        room: room,
        title: m[1],
        start: _apartmenteryCalendarDateToIso_(m[2])
      });
    }
  });
  Logger.log(`_pullAllApartmenteryEvents20260718_: pulled ${events.length} total events across ${Object.keys(ROOM_TO_UNIT_ID).length} rooms.`);
  return events;
}

/**
 * Builds a bookingId -> resId map from Sheet1's current (post-fix) state,
 * so we can tell whether a candidate is already claimed by a different,
 * presumably-correct booking.
 */
function _buildBookingIdToResIdMap20260718_() {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName('Sheet1');
  const data = src.getDataRange().getValues();
  const header = data[0];
  const idx = indexMap_(header, ['ResId', APARTMENTERY_BOOKING_ID_COL_HEADER]);
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const resId = String(data[i][idx.ResId] || '').trim();
    const bid = String(data[i][idx[APARTMENTERY_BOOKING_ID_COL_HEADER]] || '').trim();
    if (resId && bid) map[bid] = resId;
  }
  return map;
}

function findCandidatesForUnresolved20260718() {
  const events = _pullAllApartmenteryEvents20260718_();
  const bookingIdToResId = _buildBookingIdToResIdMap20260718_();

  const results = [];
  UNRESOLVED_20260718_.forEach(row => {
    const candidates = events
      .filter(e => _namesMatchIgnoringOrder_(row.guest, e.title))
      .map(e => Object.assign({}, e, {
        claimedBy: (bookingIdToResId[e.bookingId] && bookingIdToResId[e.bookingId] !== row.resId)
          ? bookingIdToResId[e.bookingId]
          : null,
        dateMatchesSheet1: e.start === row.checkin
      }));

    Logger.log(`--- ${row.resId} (${row.guest}, expected room ${row.room}, Sheet1 checkin ${row.checkin}, stored id ${row.storedId}) ---`);
    if (candidates.length === 0) {
      Logger.log(`  NO CANDIDATES — name doesn't appear anywhere in any room's current calendar. Either the booking predates apartmentery's visible window, was deleted, or the guest name in Sheet1 doesn't match apartmentery at all. Check apartmentery manually.`);
    } else {
      candidates.forEach(c => {
        Logger.log(`  candidate bookingId=${c.bookingId} room=${c.room} title="${c.title}" start=${c.start}` +
          `${c.dateMatchesSheet1 ? ' [DATE MATCHES SHEET1]' : ' [date differs from Sheet1]'}` +
          `${c.claimedBy ? ` [ALREADY CLAIMED by ${c.claimedBy} — NOT safe to reuse]` : ' [unclaimed]'}`);
      });
    }
    results.push({ resId: row.resId, guest: row.guest, expectedRoom: row.room, sheetCheckin: row.checkin, storedId: row.storedId, candidates });
  });

  const safe = results.filter(r => r.candidates.filter(c => !c.claimedBy).length === 1);
  Logger.log(`=== SUMMARY ===`);
  Logger.log(`${results.length} unresolved rows checked. ${safe.length} have exactly ONE unclaimed candidate (likely safe to apply after a quick sanity check).`);
  Logger.log(JSON.stringify(results, null, 2));
  return results;
}
