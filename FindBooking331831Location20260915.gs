/**
 * FindBooking331831Location20260915.gs
 * -----------------------------------------------------------------------
 * READ-ONLY. GET requests only, no writes/deletes/creates.
 *
 * Context: the dashboard's quick-link for this pendingMatch item pointed
 * to unit 163864 (room 204) + booking 331831, but that URL was only ever
 * a UI-constructed guess (getApartmenteryBookingUrl() builds it from
 * whatever room the dashboard's matching logic guessed at render time —
 * see the-loft-admin/src/BookingInvoiceTodo.tsx:25-31). It was never
 * verified against apartmentery.com itself. Nathan tried opening it
 * directly and got "Invalid branch/unit/booking" — so 331831 is NOT
 * under unit 163864 after all. (The room-203 calendar screenshot showing
 * nothing on 08-30/08-31 stands separately and is still real evidence —
 * it just doesn't tell us where 331831 actually is, only that it isn't
 * visibly in 203.)
 *
 * This script tries booking 331831's /edit page (a GET — same read used
 * internally by updateApartmenteryBookingEndDate etc., see
 * _getApartmenteryBookingEditFormState_ in ApartmenteryClient.gs) against
 * every known unit in ROOM_TO_UNIT_ID and reports which one (if any)
 * returns HTTP 200, plus the guest name / dates on that page so we can
 * confirm it's really Jade's booking and not a coincidental id collision.
 *
 * HOW TO RUN: Apps Script editor ▶ findBooking331831Location20260915 ▶
 * Run ▶ read log.
 */

function findBooking331831Location20260915() {
  const bookingId = '331831';
  const branchId = APARTMENTERY_BRANCH_ID;
  let found = false;

  Object.keys(ROOM_TO_UNIT_ID).forEach(function (roomNum) {
    const unitId = ROOM_TO_UNIT_ID[roomNum];
    const path = `/user/branch/${branchId}/unit/${unitId}/booking/${bookingId}/edit`;
    let response;
    try {
      response = _apartmenteryFetch_(path, { method: 'get' });
    } catch (err) {
      Logger.log(`room ${roomNum} (unit ${unitId}): fetch error — ${err.message}`);
      return;
    }
    const code = response.getResponseCode();
    if (code === 200) {
      found = true;
      const html = response.getContentText();
      const customerName = _extractInputValue_(html, 'customerName');
      const startDate = _extractInputValue_(html, 'startDate');
      const endDate = _extractInputValue_(html, 'endDate');
      Logger.log(`>>> FOUND: room ${roomNum} (unit ${unitId}) — HTTP 200`);
      Logger.log(`    customerName: ${customerName}`);
      Logger.log(`    startDate: ${startDate}  endDate: ${endDate}`);
    } else {
      Logger.log(`room ${roomNum} (unit ${unitId}): HTTP ${code} — not here`);
    }
  });

  if (!found) {
    Logger.log('Booking 331831 was not found under ANY known unit. Possibilities: ' +
      'it was deleted/moved manually since the id was recorded in Sheet1, the id itself ' +
      'is stale/wrong, or it lives under a unit not in ROOM_TO_UNIT_ID (e.g. 363 Mycondo, ' +
      'or a room closed for renovation and removed from the map). Worth checking ' +
      'apartmentery.com\'s search/customer lookup directly for "Jade Salazar" as a next step.');
  }
}
