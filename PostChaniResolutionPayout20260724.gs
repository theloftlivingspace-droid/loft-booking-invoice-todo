/**
 * PostChaniResolutionPayout20260724.gs
 * -----------------------------------------------------------------------
 * ONE-OFF: posts the ฿700.14 "Resolution Payout" (2026-07-24, conf
 * HM53FQMSC9, room 363) to Apartmentery. This is a separate top-up payout
 * on top of Chani's main stay (฿12,024.85, already transferred/invoiced
 * earlier) — same confCode, same room, different bid
 * (ABB-HM53FQMSC9-RES-20260724).
 *
 * UNVERIFIED: Chani's Sheet1 booking is marked "363 Mycondo ยกเลิก"
 * (cancelled) — checkin === checkout (2026-07-23), bookingId 326899.
 * Unlike the Photography Adjustment case, this booking is cancelled, and
 * I have no way to check from here whether Apartmentery's invoice/add
 * endpoint accepts a new invoice against a cancelled booking, or whether
 * booking 326899 even still exists as postable. If createApartmenteryInvoice
 * throws or the resulting invoice looks wrong (e.g. rejected, or attached
 * to the wrong booking), report back — this may need a different approach
 * entirely (e.g. recording as host-side income with no guest invoice, if
 * that's how Apartmentery models cancellation resolution payouts).
 *
 * Positive amount (+700.14) — this is money paid TO the host, not a
 * deduction, so unlike the Photography Adjustment it's a positive
 * otherCharge line, not negative. rentalPrice sent as '' per the same
 * lesson learned there, to try to suppress the unwanted ค่าเช่า 0.00 row.
 *
 * Call once via the Apps Script editor (select
 * postChaniResolutionPayout20260724 in the function dropdown, Run), check
 * the Logger output / Apartmentery UI, then delete this file.
 * -----------------------------------------------------------------------
 */
function postChaniResolutionPayout20260724() {
  const branchId = APARTMENTERY_BRANCH_ID;          // '6801'
  const unitId = ROOM_TO_UNIT_ID['363'];             // '164250', room 363 Mycondo
  const bookingId = '326899';                        // Chani Boran, cancelled 2026-07-23
  const paidDate = '2026-07-24';                     // matches the Resolution Payout batch date

  const invoiceResult = createApartmenteryInvoice(
    branchId,
    unitId,
    bookingId,
    '',             // rentalPrice: empty — try to suppress the ค่าเช่า row
    paidDate,
    [{
      desc: 'Resolution Payout (Airbnb, 2026-07-24, conf HM53FQMSC9)',
      price: '700.14'   // positive: money paid to host, not a deduction
    }]
  );
  Logger.log('postChaniResolutionPayout20260724: created invoice ' + JSON.stringify(invoiceResult) +
    ' on booking ' + bookingId + ' (room 363, unit ' + unitId + ')');

  try {
    const receiptResult = createApartmenteryReceipt(
      branchId, unitId, bookingId, invoiceResult.invoiceId, paidDate, 'transfer'
    );
    Logger.log('postChaniResolutionPayout20260724: created receipt ' + JSON.stringify(receiptResult));
    return { invoiceId: invoiceResult.invoiceId, receiptId: receiptResult.receiptId,
      receiptLocation: receiptResult.location };
  } catch (err) {
    throw new Error(
      `Invoice ${invoiceResult.invoiceId} was created successfully, but receipt creation ` +
      `failed: ${err.message}. Call createApartmenteryReceipt('${branchId}', '${unitId}', ` +
      `'${bookingId}', '${invoiceResult.invoiceId}', '${paidDate}', 'transfer') directly to ` +
      `retry — do not re-run postChaniResolutionPayout20260724, or it will duplicate the invoice.`
    );
  }
}
