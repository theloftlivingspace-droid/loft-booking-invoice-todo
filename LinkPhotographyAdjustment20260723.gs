/**
 * LinkPhotographyAdjustment20260723.gs
 * -----------------------------------------------------------------------
 * ONE-OFF: link the 2026-07-23 Airbnb "Paid Photography Adjustment"
 * (-฿2,862.44) to an EXISTING apartmentery booking instead of
 * auto-creating a new one.
 *
 * Why not auto-create: the adjustment line has no confCode/room of its
 * own — the email never says which room the photography was for.
 * autoCreateApartmenteryBookings() has no logic to detect/reuse an
 * existing booking for a case like this — it just checks whether *this
 * resId* already has a bookingId, and creates a new one if not, which
 * would need real dates we don't have. So this needs to be linked
 * manually, once.
 *
 * Which booking: per the batch-room auto-pick rule now in
 * payout-income-log (pick the first room resolved from the same
 * batch/amount sent together), that's J Barber — the first guest in the
 * same ฿5,613.97 batch with a resolved room. Not Syeed Ryan/room 300,
 * which an earlier one-off pass in this file targeted before the
 * auto-pick rule existed.
 *
 * What this does: adds an invoice on J Barber's EXISTING booking
 * (Apartmentery Booking ID 327170, room 214, unitId 163867), with
 * rentalPrice=0 and the photography charge recorded as an "other
 * charge" line item (other1.desc/other1.price) — i.e. ค่าใช้จ่ายอื่นๆ,
 * not ค่าเช่า. Per Nathan's hard rule, this only ADDS a new invoice; it
 * never touches/deletes any existing invoice or receipt.
 *
 * Call once via the Apps Script editor (select
 * linkPhotographyAdjustment20260723 in the function dropdown, Run), then
 * check the Logger output / apartmentery UI, then delete this file.
 * -----------------------------------------------------------------------
 */
function linkPhotographyAdjustment20260723() {
  const branchId = APARTMENTERY_BRANCH_ID;         // '6801'
  const unitId = ROOM_TO_UNIT_ID['214'];            // '163867', room 214 Legacy
  const bookingId = '327170';                       // J Barber, 2026-07-22 -> 2026-07-31

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
    ' on booking ' + bookingId + ' (room 214, unit ' + unitId + ')');
  return result;
}
