/**
 * One-off manual test of the full apartmentery flow: create a real booking,
 * then create a real invoice + receipt for it. Run this directly from the
 * Apps Script editor (select testApartmenteryFullFlow_ in the function
 * dropdown, then Run).
 *
 * THIS CREATES REAL DATA on apartmentery — a real booking, invoice, and
 * receipt for room 205 — there is no automated cancel/delete function in
 * this codebase, so once you're done verifying the result, go delete the
 * test booking manually in the apartmentery web UI. The guest name and
 * note are prefixed "TEST DELETE ME" so it's easy to spot.
 *
 * Before running: double check room 205's calendar in apartmentery for
 * 2026-09-01 to 2026-09-03 isn't already booked (this script picked that
 * date range blind, just to avoid the same-day-turnover dates already
 * known to be occupied around today's date).
 */
function testApartmenteryFullFlow_() {
  const room = '205';
  const startDate = '2026-09-01';
  const endDate = '2026-09-03';
  const testGuestName = 'TEST DELETE ME - Claude Test';
  const testRentalPrice = 1000;

  Logger.log(`[TEST] Step 1: creating booking for room ${room}, ${startDate} to ${endDate}`);

  const bookingResult = createApartmenteryBookingForRoom(room, {
    startDate: startDate,
    endDate: endDate,
    guestName: testGuestName,
    note: 'TEST DELETE ME - automated flow test'
  });

  if (bookingResult && bookingResult.skipped) {
    Logger.log(`[TEST] Booking step skipped: ${bookingResult.reason}`);
    return { step: 'booking', skipped: true, reason: bookingResult.reason };
  }

  Logger.log(`[TEST] Booking created: bookingId=${bookingResult.bookingId}, location=${bookingResult.location}`);

  Logger.log(`[TEST] Step 2: creating invoice + receipt for bookingId=${bookingResult.bookingId}`);

  const receiptResult = processPayoutToReceiptForRoom(
    room, bookingResult.bookingId, testRentalPrice, startDate
  );

  if (receiptResult && receiptResult.skipped) {
    Logger.log(`[TEST] Invoice/receipt step skipped: ${receiptResult.reason}`);
    return { step: 'invoice_receipt', skipped: true, reason: receiptResult.reason, bookingId: bookingResult.bookingId };
  }

  Logger.log(`[TEST] Invoice + receipt created: invoiceId=${receiptResult.invoiceId}, ` +
    `receiptId=${receiptResult.receiptId}, receiptLocation=${receiptResult.receiptLocation}`);

  const summary = {
    bookingId: bookingResult.bookingId,
    invoiceId: receiptResult.invoiceId,
    receiptId: receiptResult.receiptId
  };

  Logger.log('[TEST] FULL FLOW SUCCEEDED: ' + JSON.stringify(summary) +
    ' — remember to delete this test booking in apartmentery manually.');

  return summary;
}
