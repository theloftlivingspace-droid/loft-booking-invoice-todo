/**
 * FixPhotographyAdjustment20260723.gs
 * -----------------------------------------------------------------------
 * ONE-OFF: replaces the wrong invoice created by
 * linkPhotographyAdjustment20260723() (LinkPhotographyAdjustment20260723.gs,
 * now deleted — see git history).
 *
 * What was wrong (per Nathan, from the invoice screenshot,
 * IVB06801260724003): the ฿2,862.44 Photography Adjustment is a deduction
 * Airbnb took OUT of the payout, not a charge on top of it — it should
 * post as a NEGATIVE amount. The invoice also showed an unwanted
 * "ค่าเช่า (Rental) 0.00" row.
 *
 * Fix:
 *  - other1.price sent as '-2862.44' (negative) instead of '2862.44'.
 *  - rentalPrice sent as '' (empty) instead of '0' — an explicit "0"
 *    string still renders a ค่าเช่า row in Apartmentery's template; empty
 *    is untested but is the only lever this form exposes to try to
 *    suppress it. VERIFY the resulting invoice — if the ค่าเช่า row still
 *    appears with empty (or invoice/add 500s on empty rentalPrice),
 *    Apartmentery's invoice template unconditionally renders that row
 *    and it can't be removed via this endpoint; report back and we'll
 *    look at whether Apartmentery has a non-rental document type instead.
 *
 * Per Nathan's hard rule this only ADDS a new invoice — it never
 * touches/deletes the wrong one. Delete/void IVB06801260724003 by hand
 * in the Apartmentery UI BEFORE running this, so the booking isn't left
 * with both.
 *
 * Call once via the Apps Script editor (select
 * fixPhotographyAdjustment20260723 in the function dropdown, Run), check
 * the Logger output / Apartmentery UI, then delete this file (and
 * LinkPhotographyAdjustment20260723.gs, which already ran).
 * -----------------------------------------------------------------------
 */
function fixPhotographyAdjustment20260723() {
  const branchId = APARTMENTERY_BRANCH_ID;         // '6801'
  const unitId = ROOM_TO_UNIT_ID['214'];            // '163867', room 214 Legacy
  const bookingId = '327170';                       // J Barber, 2026-07-22 -> 2026-07-31

  const result = createApartmenteryInvoice(
    branchId,
    unitId,
    bookingId,
    '',             // rentalPrice: empty — try to suppress the ค่าเช่า row (was '0')
    '2026-07-23',   // dateToPay: matches the Airbnb payout batch date
    [{
      desc: 'Photography Adjustment (Airbnb payout 2026-07-23, Batch THB 5613.97)',
      price: '-2862.44'   // negative: deducted from payout, not charged
    }]
  );

  Logger.log('fixPhotographyAdjustment20260723: created invoice ' + JSON.stringify(result) +
    ' on booking ' + bookingId + ' (room 214, unit ' + unitId + ')');
  return result;
}
