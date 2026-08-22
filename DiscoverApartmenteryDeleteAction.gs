/**
 * DiscoverApartmenteryDeleteAction.gs
 * -----------------------------------------------------------------------
 * READ-ONLY diagnostic. Does NOT delete anything — only fetches pages
 * (GET requests) already known to work via _apartmenteryFetch_ and scans
 * the returned HTML for signals of a delete/cancel/destroy action:
 * forms, links, buttons, data-* attributes, onclick handlers, CSRF token
 * field names.
 *
 * WHY THIS EXISTS: ApartmenteryClient.gs has createApartmenteryBooking,
 * updateApartmenteryBookingEndDate, createApartmenteryInvoice,
 * createApartmenteryReceipt — but genuinely no delete function anywhere
 * (checked the full 1139-line file). Writing a delete call without seeing
 * a real request first means guessing the URL/method/form fields against
 * production, which is a bad idea. This script finds the real thing
 * instead — no browser devtools needed, since Apps Script already has the
 * live session cookie stored in Script Properties.
 *
 * I (Claude) cannot run this myself — apartmentery.com is not in my
 * sandbox's allowed network domains, and APARTMENTERY_SESSION lives only
 * in your Script Properties, not in the repo. You need to run this from
 * the Apps Script editor.
 *
 * HOW TO USE:
 *   1. Add this file to the loft-booking-invoice-todo Apps Script project
 *      (same project ApartmenteryClient.gs is already in).
 *   2. Pick a bookingId with NO invoice yet — the safest thing to
 *      inspect, since nothing protected by the hard rule is at risk if
 *      something unexpected happens. Get one via:
 *        - Sheet1's "Apartmentery Booking ID" column, cross-checked
 *          against invoice_apt_ids_v1 having no entry for it (i.e. exactly
 *          the isBookingIdInvoiced_() check in RoomMove.gs returning
 *          false), OR
 *        - just look in apartmentery.com's calendar for a booking you
 *          know hasn't been invoiced yet.
 *   3. In the Apps Script editor, select discoverBookingDeleteAction_ in
 *      the function dropdown... actually it takes arguments, so instead
 *      run the wrapper below: edit TEST_ROOM / TEST_BOOKING_ID at the top
 *      of runDiscoverDeleteActionForTestBooking(), then run THAT function.
 *   4. Check View > Logs (or View > Executions) for the report. Send me
 *      what it finds — especially any URL containing "delete", "destroy",
 *      "cancel", or "remove", and any CSRF/token field name+value it
 *      lists — and I'll wire up the real delete function.
 *   5. This makes GET requests only. It does not submit any form. It is
 *      safe to run against a real, live booking.
 * -----------------------------------------------------------------------
 */

// EDIT THESE TWO before running runDiscoverDeleteActionForTestBooking():
const TEST_ROOM = '205'; // any room in ROOM_TO_UNIT_ID
const TEST_BOOKING_ID = 'REPLACE_ME'; // an apartmentery bookingId with no invoice yet

function runDiscoverDeleteActionForTestBooking() {
  if (TEST_BOOKING_ID === 'REPLACE_ME') {
    Logger.log('Edit TEST_ROOM and TEST_BOOKING_ID at the top of this file first, then re-run.');
    return;
  }
  const report = discoverBookingDeleteAction_(TEST_ROOM, TEST_BOOKING_ID);
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

/**
 * Fetches every page we can reasonably reach for a given booking and scans
 * each for delete-related signals. Returns a report object; also logs it.
 */
function discoverBookingDeleteAction_(roomRaw, bookingId) {
  const unit = getApartmenteryUnitForRoom(roomRaw);
  if (!unit) {
    return { ok: false, error: `no unit mapping for room ${roomRaw} in ROOM_TO_UNIT_ID` };
  }
  const { branchId, unitId } = unit;

  const pagesToCheck = [
    { label: 'booking edit form', path: `/user/branch/${branchId}/unit/${unitId}/booking/${bookingId}/edit` },
    // Some apps only show a delete control on the "show"/detail page, not
    // the edit form — try without /edit too.
    { label: 'booking detail (no /edit)', path: `/user/branch/${branchId}/unit/${unitId}/booking/${bookingId}` },
    // The unit's own calendar/list view often has a delete icon per row —
    // that markup can reveal the action even if the two pages above don't.
    { label: 'unit calendar/list', path: `/user/branch/${branchId}/unit/${unitId}` },
  ];

  const report = { ok: true, roomRaw, branchId, unitId, bookingId, pages: [] };

  pagesToCheck.forEach(page => {
    try {
      const response = _apartmenteryFetch_(page.path, { method: 'get' });
      const code = response.getResponseCode();
      const html = code === 200 ? response.getContentText() : '';
      const findings = code === 200 ? _scanHtmlForDeleteSignals_(html, bookingId) : null;
      report.pages.push({
        label: page.label,
        path: page.path,
        httpStatus: code,
        findings: findings,
      });
    } catch (err) {
      report.pages.push({ label: page.label, path: page.path, error: err.message });
    }
  });

  return report;
}

/**
 * Regex-based scan — Apps Script has no real DOM parser, so this looks
 * for textual patterns rather than parsing the HTML properly. Good enough
 * to point at candidate lines; not meant to be a robust HTML parser.
 */
function _scanHtmlForDeleteSignals_(html, bookingId) {
  const findings = {
    keywordHits: [],       // raw lines mentioning delete/destroy/cancel/remove near this bookingId
    formActionsWithVerbs: [], // <form action="..."> where action contains a delete-ish verb
    linksWithVerbs: [],     // <a href="..."> same idea
    dataMethodDelete: [],   // data-method="delete" (common Rails-ish convention some Play apps copy)
    csrfFields: [],         // hidden input fields that look like csrf tokens (name + value)
    thaiDeleteWords: [],    // "ลบ" appearing near a link/button/form
  };

  const verbPattern = /(delete|destroy|remove|cancel)/i;

  // Split into lines for readable line-based matches in the report —
  // apartmentery's HTML may be minified onto few lines, so this is
  // approximate; treat line numbers as "roughly where", not exact.
  const lines = html.split(/\n/);
  lines.forEach((line, i) => {
    if (verbPattern.test(line) && line.indexOf(bookingId) !== -1) {
      findings.keywordHits.push({ line: i + 1, text: line.trim().substring(0, 300) });
    }
    if (/ลบ/.test(line) && (/<a|<button|<form/.test(line))) {
      findings.thaiDeleteWords.push({ line: i + 1, text: line.trim().substring(0, 300) });
    }
  });

  // <form ... action="...">
  const formRe = /<form[^>]*action=["']([^"']+)["'][^>]*>/gi;
  let m;
  while ((m = formRe.exec(html)) !== null) {
    if (verbPattern.test(m[1])) {
      findings.formActionsWithVerbs.push(m[1]);
    }
  }

  // <a ... href="...">
  const linkRe = /<a[^>]*href=["']([^"']+)["'][^>]*>/gi;
  while ((m = linkRe.exec(html)) !== null) {
    if (verbPattern.test(m[1])) {
      findings.linksWithVerbs.push(m[1]);
    }
  }

  // data-method="delete" (or DELETE)
  const dataMethodRe = /data-method=["']delete["']/gi;
  while ((m = dataMethodRe.exec(html)) !== null) {
    const start = Math.max(0, m.index - 200);
    findings.dataMethodDelete.push(html.substring(start, m.index + 40).trim());
  }

  // CSRF-ish hidden fields: <input type="hidden" name="csrfToken" value="...">
  // (or "_csrf", "authenticityToken" — apartmentery is on Play Framework,
  // whose default field name is literally "csrfToken", confirmed by
  // _extractPlayErrorMessage_'s comments elsewhere in this project, but
  // checking a few common names in case it differs on this page)
  ['csrfToken', '_csrf', 'authenticityToken', 'csrf_token'].forEach(fieldName => {
    const re = new RegExp(`<input[^>]*name=["']${fieldName}["'][^>]*value=["']([^"']*)["']`, 'i');
    const match = html.match(re);
    if (match) {
      findings.csrfFields.push({ field: fieldName, value: match[1] });
    }
  });

  return findings;
}

/**
 * Mobile-friendly wrapper — called via GET ?action=discoverDeleteAction&resId=...
 * on the live "checkinout"/"todo" webapp URL. Looks up room + Apartmentery
 * bookingId from the resId automatically (Nathan only needs to type one
 * value on his phone), and refuses to run against a booking that already
 * has an invoice — the whole point of this diagnostic is a SAFE, pre-invoice
 * booking; running it against an invoiced one isn't dangerous (still
 * read-only) but isn't what it's for, so it's blocked to avoid confusion.
 */
function discoverDeleteActionByResId_(resId) {
  if (!resId) return { ok: false, error: 'resId required — add &resId=... to the URL' };

  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  const booking = findBookingByResId_(src, resId);
  if (!booking) return { ok: false, error: `resId not found: ${resId}` };

  const aptBookingId = getApartmenteryBookingId_(resId);
  if (!aptBookingId) {
    return { ok: false, error: `resId ${resId} has no Apartmentery bookingId yet — nothing to inspect. Pick one that already shows an id in the "Apartmentery Booking ID" column of Sheet1.` };
  }
  if (isBookingIdInvoiced_(aptBookingId)) {
    return { ok: false, error: `resId ${resId} (bookingId ${aptBookingId}) already has an invoice — pick a resId with no invoice yet for this diagnostic.` };
  }

  return discoverBookingDeleteAction_(roomNum_(booking.room), aptBookingId);
}
