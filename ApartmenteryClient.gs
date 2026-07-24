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
 * Uses hotel-line-bot's /api/send-admin-alert endpoint (same BOT_URL /
 * ADMIN_TOKEN Script Properties pattern as cancelBooking_ in
 * loft-booking-invoice-todo) — sends 1:1 straight to Nathan (ADMIN_USER)
 * instead of the maid group, since a session-expiry is a technical/admin
 * issue, not something the housekeeping group needs to see.
 */
function _notifyLineSessionFailure_(detail) {
  try {
    const props = PropertiesService.getScriptProperties();
    const botUrl = props.getProperty('BOT_URL') || 'https://hotel-line-bot.onrender.com';
    const adminToken = props.getProperty('ADMIN_TOKEN') || 'apt2025@secret';
    UrlFetchApp.fetch(botUrl + '/api/send-admin-alert', {
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
/** Collapses whitespace/tabs/newlines and lowercases, for name comparison. */
function _normalizeGuestNameForMatch_(name) {
  return String(name || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

/**
 * apartmentery's customer <option> labels look like "Name / Channel " or
 * "Name (phone or ID number)" or sometimes both, or neither (plain "Name").
 * Strips the trailing "(...)" and " / Channel" parts to get just the name,
 * so it can be compared against a guest name pulled from Sheet1.
 */
function _customerOptionCoreName_(label) {
  let s = String(label || '')
    .replace(/&#x27;/g, "'")
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/\s+/g, ' ')
    .trim();
  s = s.replace(/\s*\([^)]*\)\s*$/, '').trim(); // trailing "(phone/ID)"
  const slashIdx = s.lastIndexOf(' / ');
  if (slashIdx !== -1) s = s.slice(0, slashIdx).trim(); // trailing " / Channel"
  return s;
}

/** Extracts every {id, label} from the booking form's <select id="customerId">. */
function _extractCustomerOptions_(html) {
  const selectMatch = html.match(/<select[^>]*id="customerId"[\s\S]*?<\/select>/);
  if (!selectMatch) return [];
  const optRe = /<option\s+value="([^"]*)"[^>]*>([\s\S]*?)<\/option>/g;
  const options = [];
  let m;
  while ((m = optRe.exec(selectMatch[0])) !== null) {
    options.push({ id: m[1], label: m[2] });
  }
  return options;
}

/**
 * Looks for an exact (whitespace/case-insensitive) name match against
 * apartmentery's existing customer list, so a returning guest can be
 * booked against their real customerId instead of creating a duplicate.
 * Returns one of:
 *   { status: 'found', id }        — exactly one existing customer matches
 *   { status: 'ambiguous', ids }   — 2+ different customers share the name
 *   { status: 'none' }             — no match
 * Deliberately exact-match only (no fuzzy matching) — a wrong guess here
 * silently attaches a booking to the wrong person's customer record.
 */
function _findExistingApartmenteryCustomerId_(html, guestName) {
  // guestName arrives as "Name / Channel" (e.g. "Syeed Ryan / Airbnb") from
  // every automated caller, but each dropdown option's label is stripped of
  // its " / Channel" suffix by _customerOptionCoreName_ before comparison.
  // Without stripping it here too, the target never matches any option —
  // confirmed 2026-07-17: this silently broke existing-customer matching
  // for every repeat guest, not just same-day back-to-back bookings, since
  // "new" customerType submissions were always sent as "existing customer
  // not found" even when the guest's record clearly existed.
  const target = _normalizeGuestNameForMatch_(_customerOptionCoreName_(guestName));
  if (!target) return { status: 'none' };
  const matchIds = new Set();
  _extractCustomerOptions_(html).forEach(opt => {
    if (_normalizeGuestNameForMatch_(_customerOptionCoreName_(opt.label)) === target) {
      matchIds.add(opt.id);
    }
  });
  if (matchIds.size === 0) return { status: 'none' };
  if (matchIds.size === 1) return { status: 'found', id: Array.from(matchIds)[0] };
  return { status: 'ambiguous', ids: Array.from(matchIds) };
}

function createApartmenteryBooking(branchId, unitId, opts) {
  if (!opts || !opts.startDate || !opts.guestName) {
    throw new Error('createApartmenteryBooking requires at least startDate and guestName.');
  }

  const path = `/user/branch/${branchId}/unit/${unitId}/booking`;

  // Repeat-guest matching: apartmentery's own duplicate-customer validation
  // rejects a "new" customer submission whose name matches an existing
  // customer record that has blank idNo/mobile (confirmed 2026-07-09 — this
  // broke a real automation run for a genuinely returning guest). Look the
  // guest up against the form's own customer list first; if there's a
  // single unambiguous name match, book against that existing customerId
  // instead of creating a new record. An ambiguous match (2+ different
  // people share the exact name) or no match at all falls through to
  // creating a new customer as before — guessing wrong on an ambiguous
  // match would silently attach the booking to the wrong person, which is
  // worse than occasionally hitting apartmentery's duplicate-name
  // validation error (which is now visible in the log via
  // _extractPlayErrorMessage_ below, instead of a generic 500).
  let existingCustomerId = null;
  try {
    const formResponse = _apartmenteryFetch_(path, { method: 'get' });
    if (formResponse.getResponseCode() === 200) {
      const match = _findExistingApartmenteryCustomerId_(formResponse.getContentText(), opts.guestName);
      if (match.status === 'found') {
        existingCustomerId = match.id;
        Logger.log(`createApartmenteryBooking: matched guest "${opts.guestName}" to existing ` +
          `apartmentery customerId ${existingCustomerId} — booking as existing customer.`);
      } else if (match.status === 'ambiguous') {
        Logger.log(`createApartmenteryBooking: guest "${opts.guestName}" matches ${match.ids.length} ` +
          `different existing customer records (ids ${match.ids.join(', ')}) — too ambiguous to pick ` +
          `automatically, creating as a new customer instead.`);
      }
    }
  } catch (err) {
    if (isApartmenterySessionExpiredError(err)) throw err; // don't swallow — caller needs to stop
    Logger.log(`createApartmenteryBooking: couldn't check for an existing customer match for ` +
      `"${opts.guestName}" (${err.message}) — proceeding as a new customer.`);
  }

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
    customerType: existingCustomerId ? 'existing' : 'new',
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
  // customerId is only meaningful (and only read server-side) when
  // customerType is 'existing' — leave it out entirely for new customers,
  // matching how this already worked before repeat-guest matching existed.
  if (existingCustomerId) payload.customerId = existingCustomerId;

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
    //
    // Confirmed 2026-07-16: matching by guest name ALONE (the old
    // _findBookingIdByGuestName_, with no date check) picked up a
    // DIFFERENT booking's id whenever the same guest/room already had
    // another stay in the calendar, or another guest's title happened to
    // contain this guest's name as a substring — "lastMatchId" isn't
    // guaranteed to be the one just created. This silently wrote a wrong,
    // already-used bookingId into Sheet1, producing duplicate Apartmentery
    // Booking ID values across unrelated resIds. Use the same
    // date-verified matcher the collision-recovery path already relies on
    // (_findBookingIdByGuestNameAndDate_) so this only ever accepts a
    // booking whose calendar start date also matches opts.startDate.
    if (/\/booking\/?$/.test(location)) {
      const listingResponse = _apartmenteryFetch_(location, { method: 'get' });
      const bookingId = _findBookingIdByGuestNameAndDate_(listingResponse.getContentText(), opts.guestName, opts.startDate);
      if (bookingId) {
        return { bookingId: bookingId, location: location };
      }
      Logger.log(`createApartmenteryBooking: redirected to listing page (${location}) but ` +
        `couldn't find a booking for guest "${opts.guestName}" starting "${opts.startDate}" in its ` +
        `calendar events — it may have been created under a slightly different title, or the ` +
        `calendar page hasn't caught up yet. Check apartmentery manually.`);
    }

    Logger.log(`createApartmenteryBooking: got redirect (HTTP ${code}) but Location header ` +
      `didn't match expected pattern. Raw Location: "${location}"`);
  }

  // Diagnostic logging — payload sent + the actual error message from
  // Play's error page (title/CSS alone told us nothing useful), so a
  // non-redirect response can actually be debugged instead of guessed at.
  const extractedError = _extractPlayErrorMessage_(response.getContentText());
  Logger.log('createApartmenteryBooking FAILED — payload sent: ' + JSON.stringify(payload));
  Logger.log('createApartmenteryBooking FAILED — response code ' + code + ', extracted error: ' + extractedError);

  const err = new Error(
    `Booking creation for unit ${unitId} (guest ${opts.guestName}) did not redirect ` +
    `as expected (HTTP ${code}). Response may indicate a validation error — inspect manually.`
  );
  err.apartmenteryError = extractedError;
  // "การจองนี้ชนกับการจองอื่น" = this booking collides with an existing one on the
  // same unit/dates — near-always means the booking was already created manually
  // in apartmentery before automation existed, not a real scheduling conflict.
  err.isCollision = /ชนกับการจองอื่น/.test(extractedError);
  throw err;
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
// Same-day-turnover workaround: shrink an existing booking's endDate
// -----------------------------------------------------------------------

/**
 * Reads the current values of every field the booking-edit form submits,
 * so a targeted update (e.g. shortening endDate) can round-trip every
 * other field unchanged instead of guessing/blanking them. Same "the
 * browser always submits every field" reasoning as createApartmenteryBooking
 * / createApartmenteryInvoice above — the edit form has the same kind of
 * always-present fields, so read them back from the form's own
 * server-rendered current values rather than re-deriving them.
 */
function _getApartmenteryBookingEditFormState_(branchId, unitId, bookingId) {
  const path = `/user/branch/${branchId}/unit/${unitId}/booking/${bookingId}/edit`;
  const response = _apartmenteryFetch_(path, { method: 'get' });
  if (response.getResponseCode() !== 200) {
    throw new Error(`Could not load edit form for booking ${bookingId} (HTTP ${response.getResponseCode()}).`);
  }
  const html = response.getContentText();

  return {
    startDate: _extractInputValue_(html, 'startDate'),
    endDate: _extractInputValue_(html, 'endDate'),
    note: _extractTextareaValue_(html, 'note'),
    customerType: _extractCheckedRadioValue_(html, 'customerType') || 'existing',
    customerId: _extractSelectedOptionValue_(html, 'customerId'),
    customerName: _extractInputValue_(html, 'customerName'),
    customerMobileNo: _extractInputValue_(html, 'customerMobileNo'),
    customerIdNo: _extractInputValue_(html, 'customerIdNo'),
    customerEmail: _extractInputValue_(html, 'customerEmail'),
    customerNote: _extractTextareaValue_(html, 'customerNote'),
    reminderChecked: _extractCheckboxChecked_(html, 'reminder'),
    remindEvery: _extractInputValue_(html, 'remindEvery'),
    reminderFrequency: _extractCheckedRadioValue_(html, 'reminderFrequency') || 'monthly',
    remindOnDayInMonth: _extractSelectedOptionValue_(html, 'remindOnDayInMonth'),
    remindOnDayInWeek: _extractSelectedOptionValue_(html, 'remindOnDayInWeek'),
    remindOnDayInMonthInYear: _extractSelectedOptionValue_(html, 'remindOnDayInMonthInYear'),
    remindOnMonthInYear: _extractSelectedOptionValue_(html, 'remindOnMonthInYear'),
    remindInvoiceDayBefore: _extractSelectedOptionValue_(html, 'remindInvoiceDayBefore')
  };
}

/**
 * Changes a booking's endDate on apartmentery, leaving every other field
 * (customer, note, reminder settings) exactly as it currently is — reads
 * the live edit form first (see _getApartmenteryBookingEditFormState_)
 * and resubmits its own values back, with only endDate swapped.
 *
 * Purpose: apartmentery refuses to create a new booking whose startDate
 * equals another booking's endDate on the same unit — confirmed
 * 2026-07-09 to be a hard rule on apartmentery's side (reproduced with a
 * real browser submission too, not just this script), no way around it
 * via a different payload. Nathan's existing manual fix is to shorten the
 * outgoing booking's endDate by 1 day before creating the new one — this
 * automates that step. See autoCreateApartmenteryBookings in
 * ApartmenteryAutomation.gs for where it's called.
 *
 * @param {string} branchId
 * @param {string} unitId
 * @param {string} bookingId
 * @param {string} newEndDate  YYYY-MM-DD
 */
function updateApartmenteryBookingEndDate(branchId, unitId, bookingId, newEndDate) {
  const state = _getApartmenteryBookingEditFormState_(branchId, unitId, bookingId);

  const path = `/user/branch/${branchId}/unit/${unitId}/booking/${bookingId}/edit`;
  const payload = {
    startDate: state.startDate,
    endDate: newEndDate,
    note: state.note,
    customerType: state.customerType,
    customerId: state.customerId,
    customerName: state.customerName,
    customerMobileNo: state.customerMobileNo,
    customerIdNo: state.customerIdNo,
    customerEmail: state.customerEmail,
    customerNote: state.customerNote,
    remindEvery: state.remindEvery,
    reminderFrequency: state.reminderFrequency,
    remindOnDayInMonth: state.remindOnDayInMonth,
    remindOnDayInWeek: state.remindOnDayInWeek,
    remindOnDayInMonthInYear: state.remindOnDayInMonthInYear,
    remindOnMonthInYear: state.remindOnMonthInYear,
    remindInvoiceDayBefore: state.remindInvoiceDayBefore
  };
  // reminder is a real <input type=checkbox> — only include it when it's
  // actually checked, matching how a real browser omits unchecked checkboxes.
  if (state.reminderChecked) payload.reminder = 'true';

  const response = _apartmenteryFetch_(path, { method: 'post', payload: payload });

  const code = response.getResponseCode();
  if (code >= 300 && code < 400) {
    return { ok: true, bookingId: bookingId, oldEndDate: state.endDate, newEndDate: newEndDate };
  }

  Logger.log('updateApartmenteryBookingEndDate FAILED — payload sent: ' + JSON.stringify(payload));
  Logger.log('updateApartmenteryBookingEndDate FAILED — response code ' + code + ', extracted error: ' +
    _extractPlayErrorMessage_(response.getContentText()));

  throw new Error(
    `Updating endDate for booking ${bookingId} to ${newEndDate} did not redirect as expected ` +
    `(HTTP ${code}). Response may indicate a validation error — inspect manually.`
  );
}

/**
 * Same source data as _findBookingIdByGuestName_ (the unit's calendar
 * events array), but additionally requires the event's start date to
 * exactly match checkinDate before returning a match. Used to recover
 * the apartmentery bookingId for a booking that was created manually in
 * apartmentery *before* the automation existed — those never got an
 * "Apartmentery Booking ID" written back to Sheet1, so getApartmenteryBookingId_
 * comes back empty even though the booking is really there.
 *
 * Deliberately returns null instead of guessing when there's no exact
 * date match (e.g. a returning guest with an older, unrelated stay on
 * the same unit) — silently attaching this update to the wrong stay
 * would corrupt someone else's booking dates, which is worse than just
 * falling back to "sync manually this time".
 */
/**
 * Converts apartmentery's calendar date format ('09 Jul 2026') to
 * 'YYYY-MM-DD' so it can be compared against Sheet1's checkin dates.
 * Returns null if the string doesn't match the expected format.
 */
function _apartmenteryCalendarDateToIso_(dateStr) {
  const months = { Jan:'01', Feb:'02', Mar:'03', Apr:'04', May:'05', Jun:'06',
    Jul:'07', Aug:'08', Sep:'09', Oct:'10', Nov:'11', Dec:'12' };
  const m = String(dateStr || '').trim().match(/^(\d{1,2}) (\w{3}) (\d{4})$/);
  if (!m || !months[m[2]]) return null;
  return `${m[3]}-${months[m[2]]}-${m[1].padStart(2, '0')}`;
}

/**
 * True if every word in guestName also appears as a word in title,
 * ignoring case, punctuation, and word order. Handles Sheet1 storing some
 * guests as "Lastname, Firstname" while apartmentery's title has
 * "Firstname Lastname / Channel" (confirmed 2026-07-12: this order swap
 * was the reason ~18 collisions still failed to recover after the date-
 * format fix — e.g. Sheet1 "Lebedev, Egor" vs apartmentery "Egor Lebedev").
 */
function _namesMatchIgnoringOrder_(guestName, title) {
  const normalize = s => String(s || '')
    .toLowerCase()
    .replace(/[,\/]/g, ' ')
    .split(/\s+/)
    .filter(Boolean);
  const guestWords = normalize(guestName);
  if (guestWords.length === 0) return false;
  const titleWords = new Set(normalize(title));
  return guestWords.every(w => titleWords.has(w));
}

function _findBookingIdByGuestNameAndDate_(html, guestName, checkinDate) {
  html = String(html || '');
  const blockRe = /\{\s*title:\s*'((?:[^'\\]|\\.)*)'[\s\S]*?start:\s*'([^']*)'[\s\S]*?url:\s*'([^']*)'\s*\}/g;
  let m;
  const exactMatches = [];
  const looseMatches = [];
  while ((m = blockRe.exec(html)) !== null) {
    const title = m[1];
    // Confirmed 2026-07-12 (debugApartmenteryUnitCalendar): this page
    // renders dates as 'DD Mon YYYY' (e.g. '09 Jul 2026'), not ISO —
    // convert before comparing, a plain slice(0,10) against an ISO
    // checkinDate silently matched nothing for every single booking.
    const start = _apartmenteryCalendarDateToIso_(m[2]);
    const url = m[3];
    const idMatch = url.match(/\/booking\/(\d+)/);
    if (!idMatch) continue;
    if (title.indexOf(guestName) !== -1) {
      exactMatches.push({ id: idMatch[1], start: start });
    } else if (_namesMatchIgnoringOrder_(guestName, title)) {
      looseMatches.push({ id: idMatch[1], start: start });
    }
  }
  const exact = exactMatches.find(x => x.start === checkinDate)
    || looseMatches.find(x => x.start === checkinDate);
  return exact ? exact.id : null;
}

/**
 * Convenience wrapper: looks up a room's apartmentery calendar and tries
 * to find a booking matching guestName + checkinDate exactly. Returns
 * null if the room isn't mapped, the page can't be fetched, or there's
 * no exact match — never guesses.
 */
function findApartmenteryBookingIdForRoomByGuest_(roomRaw, guestName, checkinDate) {
  const unit = getApartmenteryUnitForRoom(roomRaw);
  if (!unit) return null;
  const path = `/user/branch/${unit.branchId}/unit/${unit.unitId}/booking`;
  const response = _apartmenteryFetch_(path, { method: 'get' });
  if (response.getResponseCode() !== 200) return null;
  return _findBookingIdByGuestNameAndDate_(response.getContentText(), guestName, checkinDate);
}

/**
 * Convenience wrapper taking a Sheet1 room string instead of raw
 * branchId/unitId, matching the *ForRoom naming pattern used elsewhere.
 */
function updateApartmenteryBookingEndDateForRoom(roomRaw, bookingId, newEndDate) {
  const unit = getApartmenteryUnitForRoom(roomRaw);
  if (!unit) {
    return {
      skipped: true,
      reason: `Room "${roomRaw}" not found in ROOM_TO_UNIT_ID.`
    };
  }
  return updateApartmenteryBookingEndDate(unit.branchId, unit.unitId, bookingId, newEndDate);
}

/**
 * Re-fetches a unit's calendar listing page — the same GET a browser
 * does whenever the calendar view is shown or refreshed.
 *
 * Why: 2026-07-18, a same-day-turnover shrink (updateApartmenteryBookingEndDateForRoom)
 * got a normal 3xx success response, and diagnoseMilesCollision20260718
 * confirmed afterward that the new endDate really had persisted — yet
 * the very next createApartmenteryBooking call for the incoming booking
 * still got apartmentery's "การจองนี้ชนกับการจองอื่น" collision error,
 * twice, a minute apart. Nathan confirmed doing the same shrink-then-add
 * manually in the browser never hits this — so it isn't apartmentery's
 * data actually being stale, it's specifically that manual use always
 * loads the calendar page (via clicking into the unit) between the two
 * writes and automation never does. Untested theory: apartmentery's
 * collision check on the "add booking" form may run against a value it
 * cached from the same session's last calendar-page load, not the
 * database directly — so re-loading the calendar page right after the
 * shrink (this function), before attempting the create, should make the
 * add-booking form's collision check see the shrunk endDate the same way
 * a real browser session would.
 *
 * Read-only, side-effect-free — if this theory is wrong, all it costs is
 * one extra GET request.
 */
function refreshApartmenteryUnitCalendarForRoom(roomRaw) {
  const unit = getApartmenteryUnitForRoom(roomRaw);
  if (!unit) return;
  const path = `/user/branch/${unit.branchId}/unit/${unit.unitId}/booking`;
  _apartmenteryFetch_(path, { method: 'get' });
}

/**
 * Creates an invoice for a booking and returns its invoiceId.
 *
 * @param {string} branchId
 * @param {string} unitId
 * @param {string} bookingId
 * @param {number} rentalPrice   Amount matched from the SCB payout.
 * @param {string} [dateToPayStr] YYYY-MM-DD, defaults to today.
 */
function createApartmenteryInvoice(branchId, unitId, bookingId, rentalPrice, dateToPayStr, otherCharges) {
  const dateToPay = dateToPayStr || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');

  const path = `/user/branch/${branchId}/unit/${unitId}/booking/${bookingId}/invoice/add`;

  // Same root cause as createApartmenteryBooking above: the invoice form's
  // elect/water/other-charge/withholding/promptpay fields are plain
  // select/input elements that a real browser always submits, even when
  // hidden by CSS (electType/waterType 'no', empty other1-7 rows, etc).
  // Confirmed 2026-07-09: sending only 5 fields (as before) causes a 500 on
  // every invoice. Sending the form's own defaults for all of them keeps
  // the original intent (no electric/water charge, no extra line items,
  // no withholding, no PromptPay) but reproduces a real form submission.
  //
  // Two fields are NOT empty-string defaults, confirmed from the form's
  // rendered HTML (view-source, 2026-07-09):
  //   - electPricePerUnit has value="10.0" hardcoded in the <input>.
  //   - withholdingPercent.value is a <select> with no `selected` option,
  //     so a real browser submits the first <option> ("3") by default.
  // Sending "" for withholdingPercent.value instead of "3" is the likely
  // cause of the follow-up 500 after the first field-completeness fix —
  // the server probably fails to parse "" as the expected numeric enum.
  // addVat / withholdingPercent.apply / promptPayId.apply are real
  // <input type=checkbox> elements, unchecked by default, so real
  // browsers omit them entirely from the submit — correctly not sent here.
  const payload = {
    dateToPay: dateToPay,
    rentalPrice: String(rentalPrice),
    electType: 'no',
    electFromStr: '',
    electToStr: '',
    electPricePerUnit: '10.0',
    waterType: 'no',
    waterFromStr: '',
    waterToStr: '',
    waterPricePerUnit: '',
    'other1.desc': '', 'other1.price': '',
    'other2.desc': '', 'other2.price': '',
    'other3.desc': '', 'other3.price': '',
    'other4.desc': '', 'other4.price': '',
    'other5.desc': '', 'other5.price': '',
    'other6.desc': '', 'other6.price': '',
    'other7.desc': '', 'other7.price': '',
    'withholdingPercent.value': '3',
    'promptPayId.value': '',
    'invoiceNote.type': 'default',
    'invoiceNote.value': ''
  };

  // ค่าใช้จ่ายอื่นๆ (other1-7) — e.g. Photography Adjustment ที่ Airbnb หักจาก
  // payout ตรงๆ ไม่เกี่ยวกับค่าเช่า: บันทึกเป็น other-charge line item แยกจาก
  // rentalPrice แทนที่จะไปปนกับค่าเช่าห้อง หรือสร้าง booking หลอกขึ้นมาใหม่
  (otherCharges || []).slice(0, 7).forEach(function (item, idx) {
    payload['other' + (idx + 1) + '.desc'] = item.desc || '';
    payload['other' + (idx + 1) + '.price'] = item.price != null ? String(item.price) : '';
  });

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

  Logger.log(`createApartmenteryInvoice FAILED — branchId=${branchId} unitId=${unitId} bookingId=${bookingId} path=${path}`);
  Logger.log('createApartmenteryInvoice FAILED — payload sent: ' + JSON.stringify(payload));
  Logger.log('createApartmenteryInvoice FAILED — response code ' + code + ', extracted error: ' +
    _extractPlayErrorMessage_(response.getContentText()));

  // Diagnostic only, best-effort: if invoice/add rejects the branch/unit/
  // booking combo, check whether the *same* combo's booking-edit page loads
  // fine. If it does, the problem is specific to the invoice/add route (e.g.
  // booking status, an existing invoice already on this booking) rather than
  // branchId/unitId/bookingId being genuinely wrong for each other. If the
  // edit page ALSO fails the same way, the booking really doesn't belong to
  // this branch/unit — most likely stale bookingId or the room was
  // reassigned after this booking was created in apartmentery.
  try {
    const editPath = `/user/branch/${branchId}/unit/${unitId}/booking/${bookingId}/edit`;
    const editResponse = _apartmenteryFetch_(editPath, { method: 'get' });
    const editCode = editResponse.getResponseCode();
    if (editCode === 200) {
      Logger.log(`createApartmenteryInvoice DIAGNOSTIC — booking ${bookingId} edit page loads fine ` +
        `(HTTP 200) under branchId=${branchId} unitId=${unitId}. So the branch/unit/booking combo IS ` +
        `valid — the invoice/add rejection is specific to that route (check for an existing invoice ` +
        `already on this booking, or a booking status issue).`);
    } else {
      Logger.log(`createApartmenteryInvoice DIAGNOSTIC — booking ${bookingId} edit page ALSO failed ` +
        `(HTTP ${editCode}) under branchId=${branchId} unitId=${unitId}: ` +
        _extractPlayErrorMessage_(editResponse.getContentText()) +
        ` — this branch/unit/booking combo looks genuinely wrong (stale bookingId, or the booking's ` +
        `room was reassigned after it was created in apartmentery).`);
    }
  } catch (diagErr) {
    if (!isApartmenterySessionExpiredError(diagErr)) {
      Logger.log(`createApartmenteryInvoice DIAGNOSTIC — edit-page check itself failed: ${diagErr.message}`);
    }
  }

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
 * Extracts the value="" of whichever <input type="radio" name="X" ...
 * checked> is currently checked in a radio group. Returns '' if none
 * checked (shouldn't happen for apartmentery's forms — every radio group
 * used here always has a default-checked option).
 */
function _extractCheckedRadioValue_(html, name) {
  const re = new RegExp(`<input[^>]*name="${name}"[^>]*value="([^"]*)"[^>]*checked`, 'g');
  let m;
  let last = '';
  while ((m = re.exec(html)) !== null) last = m[1];
  return last;
}

/**
 * Extracts the value="" of the <option ... selected> inside <select id="X">.
 */
function _extractSelectedOptionValue_(html, selectId) {
  const selectRe = new RegExp(`<select[^>]*id="${selectId}"[\\s\\S]*?</select>`);
  const selectMatch = html.match(selectRe);
  if (!selectMatch) return '';
  const optMatch = selectMatch[0].match(/<option\s+value="([^"]*)"[^>]*selected/);
  return optMatch ? optMatch[1] : '';
}

/** True if <input type="checkbox" id="X" ... checked> is present and checked. */
function _extractCheckboxChecked_(html, fieldId) {
  return new RegExp(`id="${fieldId}"[^>]*checked`).test(html);
}

/** Extracts the text content of <textarea id="X" name="X">...</textarea>. */
function _extractTextareaValue_(html, fieldId) {
  const m = html.match(new RegExp(`id="${fieldId}"[^>]*>([\\s\\S]*?)</textarea>`));
  return m ? m[1].trim() : '';
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
