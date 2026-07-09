/**
 * ApartmenteryClient.gs
 * -----------------------------------------------------------------------
 * Automates invoice → receipt creation on apartmentery.com, triggered
 * from the payout-matching success path in loft-booking-invoice-todo.
 *
 * SESSION HANDLING (important):
 *   apartmentery's /login form is protected by invisible reCAPTCHA v3,
 *   which requires running Google's JS in a real browser context to
 *   produce a valid token. Apps Script's UrlFetchApp cannot do this, so
 *   automated re-login is NOT possible here.
 *
 *   Instead: the session cookie (PLAY_SESSION) is stored in Script
 *   Properties and reused across runs. When a request detects the
 *   session has expired, it stops immediately and sends a LINE alert
 *   asking you to log in manually and paste the fresh cookie back in.
 *
 * SETUP (one-time):
 *   1. Log in to apartmentery.com in your browser as normal.
 *   2. Open DevTools > Network, find any request to apartmentery.com,
 *      copy the full "Cookie" header value (or just the PLAY_SESSION=...
 *      part).
 *   3. In the Apps Script editor: Project Settings > Script Properties
 *      > add APARTMENTERY_SESSION = PLAY_SESSION=xxxxx...
 *   4. Wire _notifyLineSessionFailure_() below to your existing
 *      linePush helper (see TODO inside).
 *
 * WHEN YOU GET THE "SESSION EXPIRED" LINE ALERT:
 *   Repeat steps 1-3 above to refresh APARTMENTERY_SESSION, then
 *   whatever booking/invoice/receipt failed can be safely re-run —
 *   nothing partial gets left behind silently (see error handling notes
 *   in each function).
 * -----------------------------------------------------------------------
 */

const APARTMENTERY_BASE = 'https://apartmentery.com';

// All 11 rooms live under the same apartmentery branch (The Loft Living Space Co.).
// unitId per room, confirmed from /user/branch/6801/unit on 2026-07-09.
const APARTMENTERY_BRANCH_ID = '6801';
const ROOM_TO_UNIT_ID = {
  '103': '163863', // Elegance
  '108': '163862', // Retro
  '113': '163868', // Legacy
  '203': '163866', // Allure
  '204': '163864', // Elegance
  '205': '163865', // Allure
  '209': '193723', // Radiance
  '210': '193724', // Radiance
  '214': '163867', // Legacy
  '300': '163861', // Luxury
  '363': '164250'  // Mycondo (shown as "8/363 B" in apartmentery's UI)
};

/**
 * Looks up { branchId, unitId } for a room number as it appears in
 * Sheet1's เลขห้อง column (e.g. "103", "103 Elegance", "204 Elegance ยกเลิก").
 * Returns null if the room isn't in the map — caller should treat that as
 * "skip automation, handle this booking manually" rather than throwing.
 */
function getApartmenteryUnitForRoom(roomRaw) {
  const match = String(roomRaw || '').trim().match(/^(\d+)/);
  if (!match) return null;
  const roomNum = match[1];
  const unitId = ROOM_TO_UNIT_ID[roomNum];
  if (!unitId) return null;
  return { branchId: APARTMENTERY_BRANCH_ID, unitId: unitId };
}

function _getStoredSessionCookie_() {
  const cookie = PropertiesService.getScriptProperties().getProperty('APARTMENTERY_SESSION');
  if (!cookie) {
    throw new Error(
      'APARTMENTERY_SESSION not set in Script Properties. ' +
      'Log in to apartmentery.com manually, copy the PLAY_SESSION cookie, ' +
      'and set it in Project Settings > Script Properties.'
    );
  }
  return cookie;
}

/**
 * Pulls the actual exception title/message out of a Play framework error
 * page. Play's default error page buries this well past the <head>/CSS,
 * inside <h1>/<h2> tags — logging the first ~1500 chars of raw HTML just
 * gets you boilerplate styling, not the useful part. Falls back to a
 * larger raw slice if the expected structure isn't found.
 */
function _extractPlayErrorMessage_(html) {
  html = String(html || '');
  const strip = s => s.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();

  // A real validation rejection (HTTP 400) re-renders the normal page with
  // a Bootstrap alert box for the actual message — h1/h2 on that page is
  // just the ordinary page title, not useful. Check alert boxes first.
  const alertMatches = [...html.matchAll(/<div[^>]*class="[^"]*alert[^"]*"[^>]*>([\s\S]*?)<\/div>/gi)];
  const alertText = alertMatches.map(m => strip(m[1])).filter(Boolean).join(' | ');
  if (alertText) return alertText;

  // Otherwise this is likely Play's generic crash page (HTTP 500) — h1/h2
  // there hold the actual exception title/message.
  const h1Match = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
  const h2Match = html.match(/<h2[^>]*>([\s\S]*?)<\/h2>/i);
  if (h1Match || h2Match) {
    const title = h1Match ? strip(h1Match[1]) : '';
    const detail = h2Match ? strip(h2Match[1]) : '';
    return [title, detail].filter(Boolean).join(' — ');
  }

  // Structure didn't match what we expected — fall back to a bigger raw
  // slice so there's still something to look at.
  return 'no alert/h1/h2 found; raw (first 3000 chars): ' + html.slice(0, 3000);
}

/**
 * The unit's booking-listing page embeds all its bookings as a fullCalendar
 * events array in inline JS, e.g.:
 *   { title: 'Guest Name / Airbnb', start: '...', end: '...',
 *     url: '/user/branch/6801/unit/163865/booking/326192' }
 * This is the only place a newly-created booking's ID shows up when the
 * create-booking redirect just points back to this listing page instead of
 * the booking itself. Finds the entry whose title contains guestName and
 * returns its numeric booking ID, or null if none matches.
 */
function _findBookingIdByGuestName_(html, guestName) {
  html = String(html || '');
  const blockRe = /\{\s*title:\s*'((?:[^'\\]|\\.)*)'[\s\S]*?url:\s*'([^']*)'\s*\}/g;
  let m;
  let lastMatchId = null;
  while ((m = blockRe.exec(html)) !== null) {
    const title = m[1];
    const url = m[2];
    if (title.indexOf(guestName) !== -1) {
      const idMatch = url.match(/\/booking\/(\d+)/);
      if (idMatch) lastMatchId = idMatch[1]; // keep the last (most recent) match
    }
  }
  return lastMatchId;
}

/**
 * Checks whether a response indicates we got bounced to the login page
 * (session expired / invalid) rather than the page we actually asked for.
 */
function _isLoginPage_(response) {
  const code = response.getResponseCode();
  if (code >= 300 && code < 400) {
    const location = response.getHeaders()['Location'] || '';
    if (location.indexOf('/login') !== -1) return true;
  }
  const body = response.getContentText();
  return body.indexOf('class="form-signin"') !== -1 && body.indexOf('recaptchaToken') !== -1;
}

/**
 * Core request wrapper: attaches the stored session cookie and detects
 * expiry. On expiry it sends a LINE alert and throws — it does NOT try
 * to re-login (see header comment for why).
 */
function _apartmenteryFetch_(path, options) {
  options = options || {};
  const cookie = _getStoredSessionCookie_();

  const fetchOptions = Object.assign({}, options, {
    headers: Object.assign({}, options.headers, { Cookie: cookie }),
    followRedirects: false,
    muteHttpExceptions: true
  });

  const response = UrlFetchApp.fetch(APARTMENTERY_BASE + path, fetchOptions);

  if (_isLoginPage_(response)) {
    const msg =
      `Apartmentery session expired while calling ${path}. ` +
      `Log in manually and update APARTMENTERY_SESSION in Script Properties.`;
    _notifyLineSessionFailure_(msg);
    // Prefixed so callers (e.g. batch automation loops) can detect this
    // specific failure mode and stop immediately instead of retrying every
    // remaining item against a session that is definitely still dead.
    throw new Error('SESSION_EXPIRED: ' + msg);
  }

  return response;
}

/** True if an error thrown by _apartmenteryFetch_ was a session-expiry, not some other failure. */
function isApartmenterySessionExpiredError(err) {
  return !!(err && err.message && err.message.indexOf('SESSION_EXPIRED:') === 0);
}

/**
 * Sends a LINE alert when the stored session has expired.
 * Reuses hotel-line-bot's existing /api/send-maid-note endpoint (same
 * BOT_URL / ADMIN_TOKEN Script Properties pattern as cancelBooking_ in
 * loft-booking-invoice-todo) rather than adding a new bot endpoint —
 * "note" is free text, so it doubles fine as a generic admin alert.
 */
function _notifyLineSessionFailure_(detail) {
  try {
    const props = PropertiesService.getScriptProperties();
    const botUrl = props.getProperty('BOT_URL') || 'https://hotel-line-bot.onrender.com';
    const adminToken = props.getProperty('ADMIN_TOKEN') || 'apt2025@secret';
    UrlFetchApp.fetch(botUrl + '/api/send-maid-note', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ note: '⚠️ Apartmentery session หมดอายุ\n' + detail }),
      headers: { 'x-admin-token': adminToken },
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('Failed to send LINE alert: ' + e.message);
  }
}

// -----------------------------------------------------------------------
// Public API
// -----------------------------------------------------------------------

// -----------------------------------------------------------------------
// Phase 1: Booking creation
// -----------------------------------------------------------------------

/**
 * Creates a new booking on apartmentery and returns its bookingId.
 * Always creates a NEW customer record (customerType=new) — apartmentery
 * OTA guests are essentially always distinct per stay, and matching against
 * existing customers risks silently attaching a booking to the wrong person.
 * If you later want repeat-guest matching, that needs a deliberate lookup
 * step (search /user/customer by name/email) before this call, not a guess
 * baked into this function.
 *
 * Deliberately does NOT set the `reminder` checkbox — apartmentery's own
 * auto-invoice-reminder flow is not used here. Invoice/receipt creation is
 * driven entirely by processPayoutToReceiptForRoom() once a payout is
 * matched, so leaving apartmentery's reminder system off avoids a second,
 * uncoordinated invoice-creation path.
 *
 * @param {string} branchId
 * @param {string} unitId
 * @param {Object} opts
 * @param {string} opts.startDate      YYYY-MM-DD (check-in date)
 * @param {string} [opts.endDate]      YYYY-MM-DD, omit for open-ended stays
 * @param {string} opts.guestName      Required — becomes customerName
 * @param {string} [opts.guestMobile]
 * @param {string} [opts.guestEmail]
 * @param {string} [opts.note]         Booking-level note (e.g. "Airbnb ABB-XXXX")
 * @param {string} [opts.customerNote] Customer-level note
 */
function createApartmenteryBooking(branchId, unitId, opts) {
  if (!opts || !opts.startDate || !opts.guestName) {
    throw new Error('createApartmenteryBooking requires at least startDate and guestName.');
  }

  const path = `/user/branch/${branchId}/unit/${unitId}/booking`;

  // The `reminder` checkbox's sibling fields (remindEvery, reminderFrequency,
  // remindOnDayInMonth, etc.) are never removed from the booking form's DOM —
  // they're only hidden with CSS when the checkbox is unticked — so a real
  // browser submits them on every booking regardless of the checkbox state.
  // Confirmed 2026-07-09: omitting them entirely (as this function did
  // before) causes apartmentery's server to 500 on every booking. Sending
  // them with the form's own defaults, while still leaving `reminder` itself
  // unset, reproduces an "unticked checkbox" submission exactly — apartmentery's
  // own reminder/auto-invoice system stays off, matching the original intent
  // described below, but the request no longer 500s.
  const startDateObj = new Date(opts.startDate + 'T00:00:00Z');
  const dayOfMonth = String(startDateObj.getUTCDate());
  const monthOfYear = String(startDateObj.getUTCMonth() + 1);

  const payload = {
    startDate: opts.startDate,
    endDate: opts.endDate || '',
    note: opts.note || '',
    customerType: 'new',
    'customerName': opts.guestName,
    'customerMobileNo': opts.guestMobile || '',
    'customerIdNo': '',
    'customerEmail': opts.guestEmail || '',
    'customerNote': opts.customerNote || '',
    // reminder itself intentionally omitted — see function comment above
    remindEvery: '1',
    reminderFrequency: 'monthly',
    remindOnDayInMonth: dayOfMonth,
    remindOnDayInWeek: '1',
    remindOnDayInMonthInYear: dayOfMonth,
    remindOnMonthInYear: monthOfYear,
    remindInvoiceDayBefore: '5'
  };

  const response = _apartmenteryFetch_(path, { method: 'post', payload: payload });

  const code = response.getResponseCode();
  if (code >= 300 && code < 400) {
    const location = response.getHeaders()['Location'] || '';
    const match = location.match(/\/booking\/(\d+)/);
    if (match) {
      return { bookingId: match[1], location: location };
    }

    // Confirmed 2026-07-09: a successful booking creation redirects to the
    // unit's plain booking-listing page (e.g. ".../unit/163865/booking"),
    // not to the new booking's own URL — there's no ID in the Location
    // header at all. GET that listing page and pull the ID out of its
    // embedded fullCalendar events array by matching on guest name instead.
    if (/\/booking\/?$/.test(location)) {
      const listingResponse = _apartmenteryFetch_(location, { method: 'get' });
      const bookingId = _findBookingIdByGuestName_(listingResponse.getContentText(), opts.guestName);
      if (bookingId) {
        return { bookingId: bookingId, location: location };
      }
      Logger.log(`createApartmenteryBooking: redirected to listing page (${location}) but ` +
        `couldn't find a booking for guest "${opts.guestName}" in its calendar events — ` +
        `it may have been created under a slightly different title. Check apartmentery manually.`);
    }

    Logger.log(`createApartmenteryBooking: got redirect (HTTP ${code}) but Location header ` +
      `didn't match expected pattern. Raw Location: "${location}"`);
  }

  // Diagnostic logging — payload sent + the actual error message from
  // Play's error page (title/CSS alone told us nothing useful), so a
  // non-redirect response can actually be debugged instead of guessed at.
  Logger.log('createApartmenteryBooking FAILED — payload sent: ' + JSON.stringify(payload));
  Logger.log('createApartmenteryBooking FAILED — response code ' + code + ', extracted error: ' +
    _extractPlayErrorMessage_(response.getContentText()));

  throw new Error(
    `Booking creation for unit ${unitId} (guest ${opts.guestName}) did not redirect ` +
    `as expected (HTTP ${code}). Response may indicate a validation error — inspect manually.`
  );
}

/**
 * Convenience wrapper: same as createApartmenteryBooking, but takes the
 * room number as it appears in Sheet1's เลขห้อง column. Returns
 * { skipped: true, reason } instead of throwing if the room isn't in
 * ROOM_TO_UNIT_ID.
 *
 * @param {string} roomRaw
 * @param {Object} opts  Same shape as createApartmenteryBooking's opts.
 */
function createApartmenteryBookingForRoom(roomRaw, opts) {
  const unit = getApartmenteryUnitForRoom(roomRaw);
  if (!unit) {
    return {
      skipped: true,
      reason: `Room "${roomRaw}" not found in ROOM_TO_UNIT_ID — add it to the map ` +
              `in ApartmenteryClient.gs, or create the booking manually for now.`
    };
  }
  return createApartmenteryBooking(unit.branchId, unit.unitId, opts);
}

// -----------------------------------------------------------------------
// Phase 2/3: Invoice + receipt (see below)
// -----------------------------------------------------------------------

/**
 * Creates an invoice for a booking and returns its invoiceId.
 *
 * @param {string} branchId
 * @param {string} unitId
 * @param {string} bookingId
 * @param {number} rentalPrice   Amount matched from the SCB payout.
 * @param {string} [dateToPayStr] YYYY-MM-DD, defaults to today.
 */
function createApartmenteryInvoice(branchId, unitId, bookingId, rentalPrice, dateToPayStr) {
  const dateToPay = dateToPayStr || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');

  const path = `/user/branch/${branchId}/unit/${unitId}/booking/${bookingId}/invoice/add`;
  const payload = {
    dateToPay: dateToPay,
    rentalPrice: String(rentalPrice),
    electType: 'no',
    waterType: 'no',
    'invoiceNote.type': 'default'
  };

  const response = _apartmenteryFetch_(path, { method: 'post', payload: payload });

  // Successful submit normally redirects (302) to the invoice detail page,
  // e.g. Location: /user/branch/.../booking/.../invoice/{invoiceId}
  const code = response.getResponseCode();
  if (code >= 300 && code < 400) {
    const location = response.getHeaders()['Location'] || '';
    const match = location.match(/\/invoice\/(\d+)/);
    if (match) {
      return { invoiceId: match[1], location: location };
    }
    Logger.log(`createApartmenteryInvoice: got redirect (HTTP ${code}) but Location header ` +
      `didn't match expected pattern. Raw Location: "${location}"`);
  }

  Logger.log('createApartmenteryInvoice FAILED — payload sent: ' + JSON.stringify(payload));
  Logger.log('createApartmenteryInvoice FAILED — response code ' + code + ', extracted error: ' +
    _extractPlayErrorMessage_(response.getContentText()));

  throw new Error(
    `Invoice creation for booking ${bookingId} did not redirect as expected ` +
    `(HTTP ${code}). Response may indicate a validation error — inspect manually.`
  );
}

/**
 * Creates a receipt for an existing invoice.
 * Pre-fills electPrice/waterPrice/vat/withholding from the invoice itself
 * (GET the receipt/add form first) rather than recomputing them.
 *
 * @param {string} branchId
 * @param {string} unitId
 * @param {string} bookingId
 * @param {string} invoiceId
 * @param {string} [paidDateStr] YYYY-MM-DD, defaults to today.
 * @param {string} [paymentMethod] '' | 'transfer' | 'cash' | 'card' | 'check'
 */
function createApartmenteryReceipt(branchId, unitId, bookingId, invoiceId, paidDateStr, paymentMethod) {
  const paidDate = paidDateStr || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  const method = paymentMethod || 'transfer'; // default: matches SCB-payout trigger context

  const formPath = `/user/branch/${branchId}/unit/${unitId}/booking/${bookingId}/invoice/${invoiceId}/receipt/add`;

  // Step 1: GET the form to read prefilled values.
  const formResponse = _apartmenteryFetch_(formPath, { method: 'get' });
  const html = formResponse.getContentText();

  const prefill = {
    rentalPrice: _extractInputValue_(html, 'rentalPrice'),
    electPrice: _extractInputValue_(html, 'electPrice'),
    waterPrice: _extractInputValue_(html, 'waterPrice'),
    vat: _extractInputValue_(html, 'vat'),
    withholding: _extractInputValue_(html, 'withholding')
  };

  // Step 2: POST back with prefilled amounts + paidDate + paymentMethod.
  const payload = {
    paidDate: paidDate,
    rentalPrice: prefill.rentalPrice || '',
    electPrice: prefill.electPrice || '',
    waterPrice: prefill.waterPrice || '',
    vat: prefill.vat || '',
    withholding: prefill.withholding || '',
    paymentMethod: method,
    useReceiptNote: 'true'
  };

  const response = _apartmenteryFetch_(formPath, { method: 'post', payload: payload });

  const code = response.getResponseCode();
  if (code >= 300 && code < 400) {
    const location = response.getHeaders()['Location'] || '';
    const match = location.match(/\/receipt\/(\d+)/);
    return { receiptCreated: true, location: location, receiptId: match ? match[1] : null };
  }

  Logger.log('createApartmenteryReceipt FAILED — payload sent: ' + JSON.stringify(payload));
  Logger.log('createApartmenteryReceipt FAILED — response code ' + code + ', extracted error: ' +
    _extractPlayErrorMessage_(response.getContentText()));

  throw new Error(
    `Receipt creation for invoice ${invoiceId} did not redirect as expected ` +
    `(HTTP ${code}). Inspect manually.`
  );
}

/** Simple regex-based value extractor for <input id="X" ... value="Y"> */
function _extractInputValue_(html, fieldId) {
  const re = new RegExp(`id="${fieldId}"[^>]*value="([^"]*)"`);
  const match = html.match(re);
  return match ? match[1] : '';
}

/**
 * Full chain: create invoice, then create receipt for it.
 * Call this from the payout-match-success path in loft-booking-invoice-todo.
 *
 * If the invoice step succeeds but the receipt step fails (e.g. session
 * expires in between), the error thrown includes the invoiceId so you
 * can re-run just createApartmenteryReceipt() manually afterward instead
 * of duplicating the invoice.
 *
 * @param {string} branchId
 * @param {string} unitId
 * @param {string} bookingId
 * @param {number} rentalPrice   Matched payout amount.
 * @param {string} [paidDateStr] Defaults to today.
 */
/**
 * Convenience wrapper: same as processPayoutToReceipt, but takes the room
 * number as it appears in Sheet1's เลขห้อง column instead of raw
 * branchId/unitId. Returns { skipped: true, reason } instead of throwing
 * if the room isn't in ROOM_TO_UNIT_ID (e.g. a typo, or a room added after
 * this map was last updated) — the caller should log/notify rather than
 * halt the whole payout-matching run over one unresolved room.
 *
 * @param {string} roomRaw       e.g. "113" or "204 Elegance ยกเลิก"
 * @param {string} bookingId     apartmentery booking ID (from Sheet1's
 *                               Apartmentery Booking ID column, once that
 *                               column exists — see design discussion)
 * @param {number} rentalPrice
 * @param {string} [paidDateStr]
 */
function processPayoutToReceiptForRoom(roomRaw, bookingId, rentalPrice, paidDateStr) {
  const unit = getApartmenteryUnitForRoom(roomRaw);
  if (!unit) {
    return {
      skipped: true,
      reason: `Room "${roomRaw}" not found in ROOM_TO_UNIT_ID — add it to the map ` +
              `in ApartmenteryClient.gs, or create the invoice/receipt manually for now.`
    };
  }
  return processPayoutToReceipt(unit.branchId, unit.unitId, bookingId, rentalPrice, paidDateStr);
}

function processPayoutToReceipt(branchId, unitId, bookingId, rentalPrice, paidDateStr) {
  const invoiceResult = createApartmenteryInvoice(branchId, unitId, bookingId, rentalPrice, paidDateStr);

  try {
    const receiptResult = createApartmenteryReceipt(
      branchId, unitId, bookingId, invoiceResult.invoiceId, paidDateStr, 'transfer'
    );
    return {
      invoiceId: invoiceResult.invoiceId,
      receiptId: receiptResult.receiptId,
      receiptLocation: receiptResult.location
    };
  } catch (err) {
    throw new Error(
      `Invoice ${invoiceResult.invoiceId} was created successfully, but receipt ` +
      `creation failed: ${err.message}. Once the session issue is resolved, call ` +
      `createApartmenteryReceipt('${branchId}', '${unitId}', '${bookingId}', ` +
      `'${invoiceResult.invoiceId}') directly — do not re-run processPayoutToReceipt, ` +
      `or it will create a duplicate invoice.`
    );
  }
}
