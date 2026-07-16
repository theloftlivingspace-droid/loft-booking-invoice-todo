/**
 * AirbnbReservationToSheet1.gs
 * -----------------------------------------------------------------------
 * Fills the "363 Mycondo has no Little Hotelier" gap.
 *
 * Every other room's bookings land in Sheet1 via Little Hotelier sync.
 * Room 363 isn't on Little Hotelier at all, so its Airbnb bookings never
 * showed up in Sheet1 — which meant autoCreateApartmenteryBookings() in
 * ApartmenteryAutomation.gs (which already has 363 → unit 164250 mapped
 * in ROOM_TO_UNIT_ID) never had anything to pick up for that room.
 *
 * This file closes that gap by parsing Airbnb's "Reservation confirmed"
 * HOST email — sent the moment a guest books, not the payout email that
 * arrives much later — and appending a matching row into Sheet1 using
 * the exact same column layout getBookingToAdd_() already reads.
 *
 * Once the row exists in Sheet1, nothing else needs to change: the
 * existing hourly ApartmenteryAutomation.gs run picks it up and creates
 * the booking on apartmentery.com automatically.
 *
 * SCOPE: only the two Airbnb listings that are both actually room 363 /
 * Mycondo — "Private apartment best location in Bangkok" (id 17444947)
 * and "Cosy apartment downtown Bangkok" (id 18163498). Every other
 * Airbnb listing is ignored here — those already come through Little
 * Hotelier.
 *
 * SETUP:
 *   1. Add this file to the loft-booking-invoice-todo Apps Script project
 *      (same project as Code.gs / ApartmenteryAutomation.gs — it reuses
 *      SOURCE_SHEET_ID, SRC_BOOKING_SHEET, and indexMap_ from Code.gs).
 *   2. Run syncAirbnb363Reservations() once manually to backfill, check
 *      the execution log for how many rows it added.
 *   3. Run setupAirbnb363SyncTrigger() once to wire it to run hourly.
 * -----------------------------------------------------------------------
 */

const AIRBNB_363_LISTING_IDS = [
  '17444947', // "Private apartment best location in Bangkok" == room 363 / Mycondo (listing A)
  '18163498', // "Cosy apartment downtown Bangkok" == room 363 / Mycondo (listing B)
];
const AIRBNB_363_ROOM = '363';
const AIRBNB_363_SEARCH_QUERY = 'from:automated@airbnb.com subject:"Reservation confirmed" newer_than:90d';

/**
 * Main entry point. Searches Gmail for Airbnb "Reservation confirmed"
 * emails belonging to the 363/Mycondo listing, and appends any not
 * already present in Sheet1 (deduped by ResId = 'ABB-' + confirmation code).
 */
function syncAirbnb363Reservations() {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const sheet = ss.getSheetByName(SRC_BOOKING_SHEET);
  if (!sheet) throw new Error('ไม่พบชีต: ' + SRC_BOOKING_SHEET);

  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const numCols = header.length;
  const idx = indexMap_(header, ['เลขห้อง', 'ชื่อแขก', 'เช็คอิน', 'เช็คเอาท์', 'Channel', 'ResId', 'Note']);

  ['เลขห้อง', 'ชื่อแขก', 'เช็คอิน', 'เช็คเอาท์', 'Channel', 'ResId'].forEach(function (k) {
    if (idx[k] < 0) throw new Error('Sheet1 ไม่มีคอลัมน์ที่ต้องใช้: ' + k);
  });

  // Existing ResIds already in Sheet1 → dedupe key, so re-running this
  // (e.g. via the hourly trigger) never creates duplicate rows.
  const existingResIds = {};
  for (let r = 1; r < data.length; r++) {
    const v = String(data[r][idx['ResId']] || '').trim();
    if (v) existingResIds[v] = true;
  }

  const threads = GmailApp.search(AIRBNB_363_SEARCH_QUERY, 0, 50);
  const newRows = [];

  threads.forEach(function (thread) {
    thread.getMessages().forEach(function (msg) {
      const parsed = parseAirbnbReservationEmail_(msg);
      if (!parsed) return;

      const resId = 'ABB-' + parsed.confCode;
      if (existingResIds[resId]) return;
      existingResIds[resId] = true; // guard against dupes within the same run too

      const row = new Array(numCols).fill('');
      row[idx['เลขห้อง']] = AIRBNB_363_ROOM;
      row[idx['ชื่อแขก']] = parsed.guest;
      row[idx['เช็คอิน']] = parsed.checkin;
      row[idx['เช็คเอาท์']] = parsed.checkout;
      row[idx['Channel']] = 'Airbnb';
      row[idx['ResId']] = resId;
      if (idx['Note'] >= 0) row[idx['Note']] = 'auto: 363 Mycondo (' + parsed.confCode + ')';
      newRows.push(row);
    });
  });

  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, numCols).setValues(newRows);
  }
  Logger.log('syncAirbnb363Reservations: เพิ่ม ' + newRows.length + ' booking ใหม่เข้า Sheet1');
  return newRows.length;
}

/**
 * Returns {guest, checkin, checkout, confCode} for a 363/Mycondo Airbnb
 * "Reservation confirmed" email, or null if this message doesn't match
 * (wrong listing, or the expected fields weren't found).
 */
function parseAirbnbReservationEmail_(msg) {
  const raw = msg.getPlainBody();
  if (!raw) return null;

  // Airbnb sometimes sends quoted-printable-encoded plain bodies — same
  // cleanup parseAirbnbEmail() in payout-income-log already relies on.
  let text = raw.replace(/=\r?\n/g, '');
  text = decodeQP_(text);
  text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  // Identify the listing by its stable Airbnb room id in the listing URL
  // — not by title, since the listing title can be edited later.
  const isOurListing = AIRBNB_363_LISTING_IDS.some(function (id) {
    return text.indexOf('airbnb.com/rooms/' + id) >= 0;
  });
  if (!isOurListing) return null;

  const confMatch = text.match(/CONFIRMATION CODE\s*\r?\n\s*([A-Z0-9]{6,14})/);
  if (!confMatch) return null;
  const confCode = confMatch[1];

  // "Check-in      Checkout\n              \nThu, Jul 23   Mon, Aug 24"
  const dateLineMatch = text.match(
    /Check-in\s+Checkout[\s\S]{0,120}?\n\s*([A-Za-z]{3}, [A-Za-z]{3} \d{1,2})\s+([A-Za-z]{3}, [A-Za-z]{3} \d{1,2})/
  );
  if (!dateLineMatch) return null;

  const emailDate = msg.getDate();
  const checkin = resolveAirbnbEmailDate_(dateLineMatch[1], emailDate);
  const checkout = resolveAirbnbEmailDate_(dateLineMatch[2], emailDate);
  if (!checkin || !checkout) return null;

  // Guest name — most reliably pulled from the subject line:
  // "Reservation confirmed - Chani Boran arrives Jul 23"
  const subj = msg.getSubject() || '';
  const guestMatch = subj.match(/Reservation confirmed - (.+?) arrives/);
  const guest = guestMatch ? guestMatch[1].trim() : '';
  if (!guest) return null;

  return { guest: guest, checkin: checkin, checkout: checkout, confCode: confCode };
}

/**
 * "Thu, Jul 23" + the email's own Date header → "2026-07-23".
 * The email body never states the year, so it's inferred from the
 * email's send date: if the parsed month/day would land more than ~60
 * days in the past relative to the email itself, it must mean next year
 * (e.g. a reservation email sent in December for a January stay).
 */
function resolveAirbnbEmailDate_(str, emailDate) {
  const m = str.match(/([A-Za-z]{3}), ([A-Za-z]{3}) (\d{1,2})/);
  if (!m) return '';
  const MONTHS = { Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11 };
  const mo = MONTHS[m[2]];
  if (mo === undefined) return '';
  const day = parseInt(m[3], 10);

  const emailYear = emailDate.getFullYear();
  let candidate = new Date(emailYear, mo, day);
  const diffDays = (candidate.getTime() - emailDate.getTime()) / 86400000;
  if (diffDays < -60) candidate = new Date(emailYear + 1, mo, day);

  return Utilities.formatDate(candidate, 'Asia/Bangkok', 'yyyy-MM-dd');
}

/** Local copy of payout-income-log's decodeQP — separate Apps Script project, so duplicated here. */
function decodeQP_(s) {
  return s.replace(/((?:=[0-9A-Fa-f]{2})+)/g, function (match) {
    try {
      const bytes = match.split('=').filter(Boolean).map(function (h) { return parseInt(h, 16); });
      return Utilities.newBlob(bytes).getDataAsString('UTF-8');
    } catch (e) { return match; }
  });
}

/**
 * One-time setup: wire syncAirbnb363Reservations() to an hourly trigger.
 * Safe to run once; re-running adds a duplicate trigger, so check
 * Apps Script editor ▶ Triggers first if unsure whether it's already set up.
 */
function setupAirbnb363SyncTrigger() {
  ScriptApp.newTrigger('syncAirbnb363Reservations')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('Trigger created — syncAirbnb363Reservations จะรันทุกชั่วโมง');
}
