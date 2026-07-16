/**
 * CheckInvoiceImpact20260716.gs
 * -----------------------------------------------------------------------
 * READ-ONLY — does not write to Sheet1 or apartmentery, does not create
 * or modify anything. Only GETs booking detail pages to look.
 *
 * For each mismatched row found by auditAllApartmenteryBookingIds20260716,
 * fetches the WRONG (currently-stored) booking's detail page —
 * /user/branch/{branchId}/unit/{unitId}/booking/{storedId} — and checks
 * whether it already has an invoice/receipt attached. If it does, that
 * invoice was very likely created against the wrong guest's booking
 * (since our own automation is the only thing that creates invoices, and
 * it always uses whatever bookingId was in Sheet1 at the time — which we
 * now know was wrong for this row).
 *
 * This is best-effort text scanning (the page's exact invoice markup
 * isn't documented anywhere in this codebase yet), so treat "hasInvoicey
 * signal: true" as "needs a human to open this booking and look", not as
 * a confirmed fact.
 *
 * PASTE THE MISMATCH LIST from auditAllApartmenteryBookingIds20260716's
 * log into MISMATCHES below before running (kept as a separate step
 * rather than re-deriving it here, so you can review/trim the list
 * first — e.g. skip old 2026-02/03 rows that predate any invoice
 * automation entirely).
 */
function checkInvoiceImpact20260716() {
  // Paste {resId, room, storedId} entries here — only ones worth checking
  // (e.g. checkin date on/after whenever apartmentery invoice automation
  // went live; no point checking Feb 2026 rows if invoices weren't
  // automated yet back then).
  const MISMATCHES = [
    { resId: 'ABB-avtodagdel-20260331', room: '103 Elegance', storedId: '311988' },
    { resId: 'ABB-dietherman-20260409', room: '103 Elegance', storedId: '313282' },
    { resId: 'ABB-nikisokolo-20260414', room: '103 Elegance', storedId: '313283' },
    { resId: 'ABB-armanakbar-20260501', room: '103 Elegance', storedId: '315890' },
    { resId: 'BKC-natthaphon-20260513', room: '103 Elegance', storedId: '317011' },
    { resId: 'TRP-aomsublaos-20260519', room: '103 Elegance', storedId: '322715' },
    { resId: 'DBK-haotingyan-20260523', room: '103 Elegance', storedId: '326564' },
    { resId: 'ABB-nicklasche-20260606', room: '103 Elegance', storedId: '322992' },
    { resId: 'TRP-pornpawitb-20260612', room: '103 Elegance', storedId: '322713' },
    { resId: 'ABB-jaybrillan-20260622', room: '103 Elegance', storedId: '324662' },
    { resId: 'ABB-kionasincl-20260630', room: '103 Elegance', storedId: '325628' }
    // Add more from the mismatch list as needed — kept short here to prove
    // out the approach on room 103 first before running it against all 84.
  ];

  const results = [];
  MISMATCHES.forEach(row => {
    const unit = getApartmenteryUnitForRoom(row.room);
    if (!unit) {
      Logger.log(`${row.resId}: room "${row.room}" not in ROOM_TO_UNIT_ID — skipping.`);
      return;
    }
    // Confirmed 2026-07-16: plain .../booking/{id} returns HTTP 400 — the
    // working GET path for an existing booking's page is .../booking/{id}/edit
    // (same path _getApartmenteryBookingEditFormState_ already uses).
    const path = `/user/branch/${unit.branchId}/unit/${unit.unitId}/booking/${row.storedId}/edit`;
    try {
      const response = _apartmenteryFetch_(path, { method: 'get' });
      if (response.getResponseCode() !== 200) {
        Logger.log(`${row.resId} (wrong id ${row.storedId}): HTTP ${response.getResponseCode()} — skipping.`);
        results.push({ resId: row.resId, storedId: row.storedId, status: 'fetch_failed' });
        return;
      }
      const html = response.getContentText();
      // Best-effort signal: any link/reference to /invoice/{n} or the word
      // "receipt" on the booking's own page suggests one was created.
      const invoiceLinks = (html.match(/\/invoice\/\d+/g) || []);
      const hasReceiptWord = /receipt/i.test(html);
      Logger.log(`${row.resId} (wrong id ${row.storedId}): invoiceLinks=${JSON.stringify(invoiceLinks)}, ` +
        `mentionsReceipt=${hasReceiptWord}`);
      results.push({
        resId: row.resId,
        storedId: row.storedId,
        invoiceLinksFound: invoiceLinks.length,
        invoiceLinks: invoiceLinks,
        mentionsReceipt: hasReceiptWord
      });
    } catch (err) {
      if (isApartmenterySessionExpiredError(err)) {
        Logger.log('SESSION EXPIRED — stopping.');
        throw err;
      }
      Logger.log(`${row.resId} (wrong id ${row.storedId}): error ${err.message}`);
      results.push({ resId: row.resId, storedId: row.storedId, status: 'error', error: err.message });
    }
  });

  Logger.log('checkInvoiceImpact20260716 results: ' + JSON.stringify(results, null, 2));
  return results;
}
