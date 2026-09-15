/**
 * AddJadeSalazarExtension20260915.gs
 * -----------------------------------------------------------------------
 * Adds Jade Salazar's 2nd stay to Sheet1 (Master Sheet, SOURCE_SHEET_ID —
 * shared with payout-income-log).
 *
 * Context (from chat with Nathan, 2026-09-15):
 *   - Stay 1 already in Sheet1: ABB-jadesalaza-20260818, room 204 Elegance,
 *     2026-08-26 → 2026-08-30 (existing row — this script does NOT touch it).
 *   - Stay 2 (this script): guest moved rooms for 2026-08-30 → 2026-09-01
 *     (2 nights). Showed up in payout-income-log as a 2nd SCB payout item
 *     (NET ฿1,259.57, part of a combined ฿10,800.16 batch) that the
 *     BookingInvoiceTodo dashboard couldn't resolve a room for — flagged
 *     "⚠️ ไม่ทราบห้อง (204, 205)", both false-positive candidates from
 *     cr:-key (date+room) coincidence with unrelated bookings (Jeff Sun
 *     checking into 204 the same day; Supa Rungrueangsorakarn checking out
 *     of 205 the same day) — see BookingInvoiceTodo.tsx:374-392 for the
 *     general bug pattern. Nathan confirmed by hand: guest actually moved
 *     to room 203.
 *   - This is NOT a mid-stay move — RoomMove.gs's moveGuestRoom_ requires
 *     effectiveDate strictly between the original booking's checkin/
 *     checkout, but here moveDate (08-30) === stay 1's checkout exactly.
 *     So this is written as an ordinary new Sheet1 row instead. Once this
 *     row exists, the normal hourly autoCreateApartmenteryBookings pass
 *     creates the Apartmentery booking/invoice/receipt automatically — no
 *     separate trigger needed.
 *   - No real Airbnb conf code was available for this 2nd payout item at
 *     the time of this chat, so resId is synthetic: stay 1's resId + an
 *     -EXT- suffix + the new checkin date (same convention payout-income-
 *     log's AIRBNB_EXTENSIONS / resolveAirbnbBid() suffixing uses
 *     elsewhere). If the real conf code turns up later, rename this
 *     resId by hand in Sheet1 — nothing downstream keys off this exact
 *     string except this script's own dedupe check below.
 *
 * HOW TO RUN: Apps Script editor ▶ addJadeSalazarExtension20260915 ▶
 * Run ▶ read log. Safe to re-run — dedupes on resId, and re-checks room
 * availability every time rather than assuming the first run's check
 * still holds.
 */

const JADE_EXT_20260915_ = {
  resId:    'ABB-jadesalaza-20260818-EXT-20260830',
  room:     '203',   // confirmed by Nathan in chat — guest moved 204 → 203
  guest:    'Jade Salazar',
  checkIn:  '2026-08-30',
  checkOut: '2026-09-01',
  channel:  'Airbnb',
  note:     'ย้ายห้องจาก 204 (stay 2, ต่อจาก ABB-jadesalaza-20260818) — เพิ่มโดยสคริปต์นี้ 2026-09-15',
};

function addJadeSalazarExtension20260915() {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  if (!src) { Logger.log('Sheet1 not found'); return; }

  const data = src.getDataRange().getValues();
  const header = data[0];
  const idx = indexMap_(header, [
    'ResId', 'เลขห้อง', 'ชื่อแขก', 'เช็คอิน', 'เช็คเอาท์', 'Channel', 'Note',
    APARTMENTERY_BOOKING_ID_COL_HEADER,
  ]);
  const required = ['ResId', 'เลขห้อง', 'ชื่อแขก', 'เช็คอิน', 'เช็คเอาท์'];
  const missing = required.filter(function (k) { return idx[k] < 0; });
  if (missing.length) {
    Logger.log('Missing required Sheet1 columns: ' + missing.join(', '));
    return;
  }

  // Dedupe — safe to re-run.
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.ResId] || '').trim() === JADE_EXT_20260915_.resId) {
      Logger.log('Row already exists (resId ' + JADE_EXT_20260915_.resId + ') — no write, skipping.');
      return;
    }
  }

  // Reuses RoomMove.gs's availability check (same script project, global
  // scope) so this gets the exact same conflict logic a real room-move
  // would use, instead of re-implementing it slightly differently here.
  const availability = isRoomAvailable_(
    src, JADE_EXT_20260915_.room, JADE_EXT_20260915_.checkIn,
    JADE_EXT_20260915_.checkOut, JADE_EXT_20260915_.resId
  );
  if (!availability.available) {
    Logger.log('room 203 NOT available ' + JADE_EXT_20260915_.checkIn + ' – ' +
      JADE_EXT_20260915_.checkOut + ' — conflicts with ' +
      JSON.stringify(availability.conflict) + '. Row NOT written — check manually.');
    return;
  }

  const roomLabel = (typeof roomLabel_ === 'function')
    ? roomLabel_(JADE_EXT_20260915_.room)
    : JADE_EXT_20260915_.room + ' Allure'; // fallback if RoomMove.gs somehow isn't loaded

  const newRow = new Array(header.length).fill('');
  newRow[idx.ResId] = JADE_EXT_20260915_.resId;
  newRow[idx['เลขห้อง']] = roomLabel;
  newRow[idx['ชื่อแขก']] = JADE_EXT_20260915_.guest;
  newRow[idx['เช็คอิน']] = JADE_EXT_20260915_.checkIn;
  newRow[idx['เช็คเอาท์']] = JADE_EXT_20260915_.checkOut;
  if (idx.Channel >= 0) newRow[idx.Channel] = JADE_EXT_20260915_.channel;
  if (idx.Note >= 0) newRow[idx.Note] = JADE_EXT_20260915_.note;
  if (idx[APARTMENTERY_BOOKING_ID_COL_HEADER] >= 0) {
    // Must get its own, fresh id via autoCreateApartmenteryBookings —
    // never carry a stale id forward.
    newRow[idx[APARTMENTERY_BOOKING_ID_COL_HEADER]] = '';
  }

  const newRowIndex = src.getLastRow() + 1;
  src.getRange(newRowIndex, 1, 1, newRow.length).setValues([newRow]);

  if (typeof triggerStyleSheet1_ === 'function') triggerStyleSheet1_();

  Logger.log('Added row ' + newRowIndex + ': ' + JADE_EXT_20260915_.resId + ' — room ' +
    roomLabel + ', ' + JADE_EXT_20260915_.checkIn + ' → ' + JADE_EXT_20260915_.checkOut +
    '. Apartmentery booking/invoice/receipt will be created by the normal hourly automation pass.');
}
