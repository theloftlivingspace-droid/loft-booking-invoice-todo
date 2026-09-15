/**
 * CheckJade331831Status20260915.gs
 * -----------------------------------------------------------------------
 * READ-ONLY. Does not write, delete, or create anything on Apartmentery
 * or in Sheet1. Run this BEFORE any delete/recreate of booking 331831 so
 * Nathan knows whether an invoice/receipt has already been issued for it.
 *
 * Context: booking 331831 (resId ABB-jadesalaza-20260830) was created on
 * Apartmentery under unit 163864 = room 204, but Sheet1 (and the actual
 * apartmentery.com calendar for room 203, confirmed by screenshot
 * 2026-09-15) says the guest was in room 203. Before deleting the wrong
 * booking and letting autoCreateApartmenteryBookings recreate it under
 * the correct room, we need to know if invoice 2850271 already has a
 * receipt (paid) against it, since deleteApartmenteryBooking_ only
 * removes the booking — it does not know about or touch receipts.
 *
 * HOW TO RUN: Apps Script editor ▶ checkJade331831Status20260915 ▶
 * Run ▶ read log.
 */

function checkJade331831Status20260915() {
  const resId = 'ABB-jadesalaza-20260830';
  const aptBookingId = getApartmenteryBookingId_(resId);
  Logger.log('resId: ' + resId);
  Logger.log('Apartmentery Booking ID on file: ' + aptBookingId);

  if (!aptBookingId) {
    Logger.log('No Apartmentery Booking ID stored for this resId in Sheet1 — nothing to check further here.');
    return;
  }

  const hasInvoice = isBookingIdInvoiced_(aptBookingId);
  Logger.log('Has an invoice recorded against this bookingId (invoice_apt_ids_v1)? ' + hasInvoice);

  if (hasInvoice) {
    const map = getProp_(PROP_KEY_INVOICE_APT_IDS);
    Object.keys(map).forEach(function (invoiceKey) {
      const stored = String(map[invoiceKey] || '');
      const parts = stored.split(':');
      if (parts[0] === String(aptBookingId).trim()) {
        Logger.log('  invoiceKey: ' + invoiceKey + '  ->  stored value: ' + stored +
          '  (format is "{aptBookingId}:{invoiceId}")');
      }
    });
    Logger.log('NOTE: this only tells us an invoice WAS CREATED. It does not tell us if a ' +
      'receipt/payment was recorded against it — deleteApartmenteryBooking_ and ' +
      'createApartmenteryReceipt are separate calls with no shared log this script can read. ' +
      'Please also check the invoice page on apartmentery.com directly ' +
      '(https://apartmentery.com/user/branch/6801/unit/163864/booking/331831/invoice/2850271) ' +
      'for a "paid" / receipt status before deleting.');
  } else {
    Logger.log('No invoice on record for this bookingId in our tracking — but the link Nathan ' +
      'shared (invoice/2850271) suggests one exists on Apartmentery itself. Please check that ' +
      'page directly before deleting, since this script only sees what our own automation has ' +
      'recorded, not what a human may have created manually on Apartmentery.');
  }
}
