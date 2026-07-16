/**
 * CheckInvoiceImpact20260716.gs
 * -----------------------------------------------------------------------
 * READ-ONLY — does not write to Sheet1 or apartmentery, does not create
 * or modify anything. Only GETs calendar/booking pages to look.
 *
 * For each mismatched row found by auditAllApartmenteryBookingIds20260716,
 * fetches the WRONG (currently-stored) booking's detail page and checks
 * whether it already has an invoice/receipt attached. If it does, that
 * invoice was very likely created against the wrong guest's booking
 * (since our own automation is the only thing that creates invoices, and
 * it always uses whatever bookingId was in Sheet1 at the time — which we
 * now know was wrong for this row).
 *
 * 2026-07-16 fix: apartmentery booking IDs are NOT scoped to a single
 * unit — they're a global sequence across every room. Opening a stored
 * id under the mismatched row's OWN room (e.g. checking id 311988 under
 * room 103's unit) returns HTTP 400 whenever that id actually belongs to
 * a different room, which is exactly what happened for every row below.
 * Fix: build a global bookingId -> {room, title, isoStart} index FIRST
 * by fetching every room's calendar once (there are only 10 rooms — see
 * ROOM_TO_UNIT_ID), then look up each storedId's REAL room from that
 * index before fetching its /edit page.
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
  // automated yet back then). `room` is kept only for logging context —
  // it is NOT used to pick which unit to fetch storedId from anymore.
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
    // out the approach first before running it against all 84.
  ];

  const bookingIndex = _buildGlobalApartmenteryBookingIndex20260716_();
  Logger.log(`Global booking index built: ${Object.keys(bookingIndex).length} booking id(s) across ` +
    `${Object.keys(ROOM_TO_UNIT_ID).length} room(s).`);

  const results = [];
  MISMATCHES.forEach(row => {
    const hit = bookingIndex[row.storedId];
    if (!hit) {
      Logger.log(`${row.resId} (wrong id ${row.storedId}): not found in ANY room's calendar — skipping.`);
      results.push({ resId: row.resId, storedId: row.storedId, status: 'id_not_found_in_any_room' });
      return;
    }
    // Use the id's REAL room/unit from the index — not row.room (the
    // mismatched Sheet1 row's own room), which is what caused the 400s.
    const unit = getApartmenteryUnitForRoom(hit.room);
    if (!unit) {
      Logger.log(`${row.resId} (wrong id ${row.storedId}): real room "${hit.room}" not in ROOM_TO_UNIT_ID — skipping.`);
      results.push({ resId: row.resId, storedId: row.storedId, status: 'real_room_not_in_map', realRoom: hit.room });
      return;
    }
    const path = `/user/branch/${unit.branchId}/unit/${unit.unitId}/booking/${row.storedId}/edit`;
    try {
      const response = _apartmenteryFetch_(path, { method: 'get' });
      if (response.getResponseCode() !== 200) {
        Logger.log(`${row.resId} (wrong id ${row.storedId}, real room ${hit.room}): ` +
          `HTTP ${response.getResponseCode()} — skipping.`);
        results.push({ resId: row.resId, storedId: row.storedId, realRoom: hit.room, status: 'fetch_failed' });
        return;
      }
      const html = response.getContentText();
      // Best-effort signal: any link/reference to /invoice/{n} or the word
      // "receipt" on the booking's own page suggests one was created.
      const invoiceLinks = (html.match(/\/invoice\/\d+/g) || []);
      const hasReceiptWord = /receipt/i.test(html);
      Logger.log(`${row.resId} (wrong id ${row.storedId}, real room ${hit.room}, real guest "${hit.title}"): ` +
        `invoiceLinks=${JSON.stringify(invoiceLinks)}, mentionsReceipt=${hasReceiptWord}`);
      results.push({
        resId: row.resId,
        storedId: row.storedId,
        realRoom: hit.room,
        realGuestTitle: hit.title,
        invoiceLinksFound: invoiceLinks.length,
        invoiceLinks: invoiceLinks,
        mentionsReceipt: hasReceiptWord
      });
    } catch (err) {
      if (isApartmenterySessionExpiredError(err)) {
        Logger.log('SESSION EXPIRED — stopping.');
        throw err;
      }
      Logger.log(`${row.resId} (wrong id ${row.storedId}, real room ${hit.room}): error ${err.message}`);
      results.push({ resId: row.resId, storedId: row.storedId, realRoom: hit.room, status: 'error', error: err.message });
    }
  });

  Logger.log('checkInvoiceImpact20260716 results: ' + JSON.stringify(results, null, 2));
  return results;
}

/**
 * Fetches every room's apartmentery calendar exactly once and returns a
 * flat map of bookingId -> { room, title, isoStart }, covering every room
 * in ROOM_TO_UNIT_ID. Booking ids are a single global sequence shared
 * across all units, so a stored id can belong to ANY room — this index
 * lets callers look up an id's real room instead of assuming it's the
 * same room as whatever Sheet1 row the id happened to be stored on.
 *
 * Same calendar-parsing regex/date-conversion as
 * auditAllApartmenteryBookingIds20260716, reused here so both scripts
 * stay in sync if apartmentery's markup ever changes.
 */
function _buildGlobalApartmenteryBookingIndex20260716_() {
  const index = {};
  Object.keys(ROOM_TO_UNIT_ID).forEach(roomNum => {
    const unit = getApartmenteryUnitForRoom(roomNum);
    if (!unit) return; // can't happen — roomNum comes straight from the map — but stay defensive.

    let html;
    try {
      const path = `/user/branch/${unit.branchId}/unit/${unit.unitId}/booking`;
      const response = _apartmenteryFetch_(path, { method: 'get' });
      if (response.getResponseCode() !== 200) {
        Logger.log(`Room ${roomNum}: calendar fetch returned HTTP ${response.getResponseCode()} — skipping for index.`);
        return;
      }
      html = response.getContentText();
    } catch (err) {
      if (isApartmenterySessionExpiredError(err)) {
        Logger.log('SESSION EXPIRED while building global booking index — refresh APARTMENTERY_SESSION and re-run.');
        throw err;
      }
      Logger.log(`Room ${roomNum}: fetch error (${err.message}) — skipping for index.`);
      return;
    }

    const blockRe = /\{\s*title:\s*'((?:[^'\\]|\\.)*)'[\s\S]*?start:\s*'([^']*)'[\s\S]*?url:\s*'([^']*)'\s*\}/g;
    let m;
    let countForRoom = 0;
    while ((m = blockRe.exec(html)) !== null) {
      const idMatch = m[3].match(/\/booking\/(\d+)/);
      if (!idMatch) continue;
      const id = idMatch[1];
      // If the same id somehow shows up on two rooms' calendars, keep the
      // first one found and log the collision rather than silently
      // overwriting — worth a human look, but shouldn't happen since ids
      // are meant to be globally unique.
      if (index[id]) {
        Logger.log(`Booking id ${id} appears on BOTH room ${index[id].room} and room ${roomNum} — keeping ${index[id].room}.`);
        continue;
      }
      index[id] = { room: roomNum, title: m[1], isoStart: _apartmenteryCalendarDateToIso_(m[2]) };
      countForRoom++;
    }
    Logger.log(`Room ${roomNum}: indexed ${countForRoom} booking id(s).`);
  });
  return index;
}
