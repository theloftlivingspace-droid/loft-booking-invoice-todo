/**
 * LinkPhotographyAdjustment20260723.gs
 * -----------------------------------------------------------------------
 * ONE-OFF: link the 2026-07-23 Airbnb "Paid Photography Adjustment"
 * (-฿2,862.44, ห้อง 300) to an EXISTING apartmentery booking instead of
 * auto-creating a new one.
 *
 * Why not auto-create: the adjustment row's dates (inherited from the
 * payout email's own date range, 2026-07-22 → 2026-07-31 / 9 nights) are
 * J Barber's stay dates for room 214, not room 300 — room 300's actual
 * current booking is Syeed Ryan, 2026-07-28 → 2026-08-01. Auto-creating a
 * room-300 booking for 07-22→07-31 would overlap Syeed Ryan's real stay
 * (07-28 falls inside that range) and either error out or corrupt room
 * 300's apartmentery calendar. autoCreateApartmenteryBookings() has no
 * logic to detect/reuse an existing booking in that case — it just
 * checks whether *this resId* already has a bookingId, and creates a new
 * one if not. So this needs to be linked manually, once.
 *
 * What this does instead: adds an invoice on the EXISTING Syeed Ryan
 * booking (Apartmentery Booking ID 326999, room 300, unitId 163861),
 * with rentalPrice=0 and the photography charge recorded as an
 * "other charge" line item (other1.desc/other1.price) — i.e. ค่าใช้จ่าย
 * อื่นๆ, not ค่าเช่า. Per Nathan's hard rule, this only ADDS a new
 * invoice; it never touches/deletes any existing invoice or receipt.
 *
 * Call once via the Apps Script editor (select
 * linkPhotographyAdjustment20260723 in the function dropdown, Run), then
 * check the Logger output / apartmentery UI, then delete this file.
 * -----------------------------------------------------------------------
 */
function linkPhotographyAdjustment20260723() {
  const branchId = APARTMENTERY_BRANCH_ID;         // '6801'
  const unitId = ROOM_TO_UNIT_ID['300'];            // '163861', room 300 Luxury
  const bookingId = '326999';                       // Syeed Ryan, 2026-07-28 -> 2026-08-01

  const result = createApartmenteryInvoice(
    branchId,
    unitId,
    bookingId,
    '0',            // rentalPrice: 0 — this is not room rent
    '2026-07-23',   // dateToPay: matches the Airbnb payout batch date
    [{
      desc: 'Photography Adjustment (Airbnb payout 2026-07-23, Batch THB 5613.97)',
      price: '2862.44'
    }]
  );

  Logger.log('linkPhotographyAdjustment20260723: created invoice ' + JSON.stringify(result) +
    ' on booking ' + bookingId + ' (room 300, unit ' + unitId + ')');
  return result;
}
