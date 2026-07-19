/**
 * ApartmenteryAutomation.gs
 * -----------------------------------------------------------------------
 * Fully-automated bridge between loft-booking-invoice-todo's existing
 * "things that still need doing" lists (getBookingToAdd_, getInvoiceToCreate_)
 * and apartmentery.com (via ApartmenteryClient.gs — must be added to this
 * same Apps Script project for these functions to work).
 *
 * WHAT THIS RUNS AUTOMATICALLY (no per-item approval):
 *   1. autoCreateApartmenteryBookings()
 *      For every booking in getBookingToAdd_() not yet marked done, and
 *      whose room resolves via ROOM_TO_UNIT_ID: creates the booking on
 *      apartmentery, stores the returned apartmentery bookingId back into
 *      Sheet1 (new "Apartmentery Booking ID" column), and calls
 *      setBookingDone(resId, true).
 *
 *   2. autoCreateApartmenteryInvoicesAndReceipts()
 *      For every payout in getInvoiceToCreate_() not yet marked done: looks
 *      up the matching Sheet1 row via the SAME matchKeys_ fuzzy-join already
 *      used elsewhere in this project, reads that row's apartmentery
 *      bookingId, then creates the invoice + receipt (net amount = the
 *      matched payout amount, paymentMethod=transfer), and calls
 *      setInvoiceDone(invoiceKey, true).
 *
 * WHAT IT SKIPS (leaves undone, does NOT mark done, does NOT alert per item):
 *   - Rooms not in ROOM_TO_UNIT_ID (e.g. a new room added after this file
 *     was last updated) — add the room to ApartmenteryClient.gs's
 *     ROOM_TO_UNIT_ID map, then it'll pick up next run.
 *   - Invoice items whose matching Sheet1 booking has no apartmentery
 *     bookingId yet (booking automation hasn't caught up, or is itself
 *     skipped for the reason above).
 *   - Invoice items where room resolution itself failed upstream
 *     (getInvoiceToCreate_ already flags these as "⚠️ ไม่ทราบห้อง (...)").
 *
 * WHAT STOPS THE WHOLE RUN IMMEDIATELY:
 *   - Apartmentery session expiry (SESSION_EXPIRED from ApartmenteryClient).
 *     One LINE alert is sent (not one per remaining item) and the run
 *     exits — every other item stays undone until the session cookie is
 *     refreshed, at which point the next run picks up exactly where it
 *     left off (nothing here is order-dependent).
 *
 * SETUP:
 *   1. Add ApartmenteryClient.gs to this same Apps Script project.
 *   2. Set APARTMENTERY_SESSION in Script Properties (see that file's header).
 *   3. Run addApartmenteryBookingIdColumnIfMissing_() once manually to add
 *      the tracking column to Sheet1 (idempotent — safe to run more than once).
 *   4. Wire runApartmenteryAutomation() to a time trigger, e.g. hourly:
 *      Apps Script editor ▶ Triggers ▶ Add Trigger ▶ runApartmenteryAutomation
 *      ▶ Time-driven ▶ Hour timer ▶ Every hour.
 * -----------------------------------------------------------------------
 */

/**
 * Wrapper — addApartmenteryBookingIdColumnIfMissing_() ends in "_" so Apps
 * Script hides it from the "select function to run" dropdown. Call this
 * instead when running manually from the editor.
 */
function runAddApartmenteryBookingIdColumn() {
  addApartmenteryBookingIdColumnIfMissing_();
  Logger.log('Done — column "' + APARTMENTERY_BOOKING_ID_COL_HEADER + '" is present on Sheet1.');
}

const APARTMENTERY_BOOKING_ID_COL_HEADER = 'Apartmentery Booking ID';

/**
 * dateStr - 1 day, both YYYY-MM-DD. Built in UTC (not Asia/Bangkok) so the
 * subtraction never shifts across a DST-less, whole-day boundary weirdly —
 * this is pure calendar-date arithmetic, no time-of-day involved.
 */
function _dateMinusOneDay_(dateStr) {
  const d = new Date(dateStr + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() - 1);
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

/**
 * One-time setup: run this ONCE from the Apps Script editor to wire
 * runApartmenteryAutomation() to an hourly trigger. This was previously
 * a manual "go do this yourself" step in the header comment above and
 * was never actually done — confirmed 2026-07-10 via the Triggers tab
 * only showing triggerHotelJob19, nothing for runApartmenteryAutomation.
 * That's why new bookings never reached apartmentery automatically:
 * the function itself was fine, it just had no trigger calling it.
 *
 * Also runs addApartmenteryBookingIdColumnIfMissing_() as part of setup
 * (idempotent, safe even if the column already exists).
 *
 * Safe to re-run: removes any existing runApartmenteryAutomation
 * trigger first, so running this twice won't create duplicates.
 */
function installApartmenteryAutomationTrigger() {
  addApartmenteryBookingIdColumnIfMissing_();

  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'runApartmenteryAutomation') {
      ScriptApp.deleteTrigger(t);
    }
  }
  ScriptApp.newTrigger('runApartmenteryAutomation')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('[installApartmenteryAutomationTrigger] Trigger installed — runApartmenteryAutomation will run every hour from now on.');
}

/**
 * Adds the "Apartmentery Booking ID" column to Sheet1 if it isn't there
 * yet. Idempotent — safe to call on every run (cheap no-op if present).
 */
function addApartmenteryBookingIdColumnIfMissing_() {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  if (!src) throw new Error('ไม่พบชีต: ' + SRC_BOOKING_SHEET);
  const header = src.getRange(1, 1, 1, src.getLastColumn()).getValues()[0];
  if (header.indexOf(APARTMENTERY_BOOKING_ID_COL_HEADER) >= 0) return; // already present
  const nextCol = src.getLastColumn() + 1;
  src.getRange(1, nextCol).setValue(APARTMENTERY_BOOKING_ID_COL_HEADER);
}

/** Reads the apartmentery bookingId already stored for a given Sheet1 ResId, or '' if none. */
function getApartmenteryBookingId_(resId) {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  if (!src) return '';
  const data = src.getDataRange().getValues();
  const header = data[0];
  const idx = indexMap_(header, ['ResId', APARTMENTERY_BOOKING_ID_COL_HEADER]);
  if (idx.ResId < 0 || idx[APARTMENTERY_BOOKING_ID_COL_HEADER] < 0) return '';
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.ResId] || '').trim() === resId) {
      return String(data[i][idx[APARTMENTERY_BOOKING_ID_COL_HEADER]] || '').trim();
    }
  }
  return '';
}

/**
 * Scans Sheet1's Apartmentery Booking ID column for any OTHER resId
 * already holding this exact bookingId. A given apartmentery bookingId
 * must map to exactly one resId — anything else means one of the two
 * rows is (or is about to become) wrong. Returns the conflicting resId,
 * or null if the bookingId is free (or only used by excludeResId itself).
 *
 * Added 2026-07-18 after the booking-id derangement bug resurfaced a
 * second time (96 mismatched rows found by auditAllApartmenteryBookingIds,
 * including brand-new July bookings created well after the first
 * comprehensive fix on 2026-07-16) — the guest+date matcher in
 * createApartmenteryBooking's fallback path can still silently return a
 * wrong (but exact-match-satisfying) id, e.g. when apartmentery's calendar
 * page hasn't yet indexed the just-created booking. Checking uniqueness
 * at write-time stops a wrong id from ever landing in Sheet1, instead of
 * only being caught later by a manual audit sweep.
 */
function _findResIdUsingApartmenteryBookingId_(apartmenteryBookingId, excludeResId) {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  if (!src) return null;
  const data = src.getDataRange().getValues();
  const header = data[0];
  const idx = indexMap_(header, ['ResId', APARTMENTERY_BOOKING_ID_COL_HEADER]);
  if (idx.ResId < 0 || idx[APARTMENTERY_BOOKING_ID_COL_HEADER] < 0) return null;
  const target = String(apartmenteryBookingId || '').trim();
  if (!target) return null;
  for (let i = 1; i < data.length; i++) {
    const rResId = String(data[i][idx.ResId] || '').trim();
    if (!rResId || rResId === excludeResId) continue;
    if (String(data[i][idx[APARTMENTERY_BOOKING_ID_COL_HEADER]] || '').trim() === target) {
      return rResId;
    }
  }
  return null;
}

/**
 * Writes the apartmentery bookingId for a given Sheet1 ResId.
 *
 * Guards against the same bookingId ever being written to two different
 * resIds — a bookingId on apartmentery belongs to exactly one booking,
 * so this must be true in Sheet1 too. Pass `{ allowOverwriteConflict: true }`
 * only from a corrective sweep (e.g. fixAllApartmenteryBookingIdsComprehensive)
 * that has independently re-verified the id against apartmentery's live
 * calendar — never from the primary creation/recovery path, since that's
 * exactly where a wrong id first gets introduced.
 */
function setApartmenteryBookingId_(resId, apartmenteryBookingId, opts) {
  opts = opts || {};
  if (!opts.allowOverwriteConflict) {
    const conflictResId = _findResIdUsingApartmenteryBookingId_(apartmenteryBookingId, resId);
    if (conflictResId) {
      const msg = `refused to write bookingId ${apartmenteryBookingId} for ${resId} — ` +
        `already assigned to a different resId (${conflictResId}). This lookup is likely ` +
        `wrong (apartmentery calendar indexing lag or a stale match) — leaving ${resId} ` +
        `without a bookingId so the next run retries instead of corrupting Sheet1.`;
      Logger.log('setApartmenteryBookingId_: ' + msg);
      return { ok: false, error: msg, conflictResId: conflictResId };
    }
  }
  addApartmenteryBookingIdColumnIfMissing_();
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  const data = src.getDataRange().getValues();
  const header = data[0];
  const idx = indexMap_(header, ['ResId', APARTMENTERY_BOOKING_ID_COL_HEADER]);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.ResId] || '').trim() === resId) {
      src.getRange(i + 1, idx[APARTMENTERY_BOOKING_ID_COL_HEADER] + 1).setValue(apartmenteryBookingId);
      return { ok: true };
    }
  }
  return { ok: false, error: 'resId not found: ' + resId };
}

/**
 * Phase 1 automation: create any missing apartmentery bookings.
 * Returns a summary object for logging — call this from
 * runApartmenteryAutomation() rather than a trigger directly, so both
 * phases share one run and one session-expiry stop.
 */
/**
 * Self-heal pass: booking_done_v1 (this automation's "already created in
 * apartmentery" flag) is the SAME Script Properties value the-loft-admin
 * dashboard's manual todo checkbox writes to via doGet_'s setBookingDone
 * action (Code.gs). Ticking that checkbox for an unrelated reason — or any
 * other bug that marks a resId done before it actually has a bookingId —
 * permanently hides that row from autoCreateApartmenteryBookings, since the
 * `if (b.done) continue` skip happens before anything checks whether an
 * apartmentery bookingId actually exists. Confirmed 2026-07-16: 9 rows
 * (Natphatsorn wongwai / EXP-natphatsorn-20260711 among them) were stuck
 * exactly this way — done=true, bookingId column empty — which is also why
 * the invoice-matching phase later cross-matched Natphatsorn's payout to a
 * different guest's booking in the same room (see the guest-name guard in
 * autoCreateApartmenteryInvoicesAndReceipts).
 *
 * Un-mark any non-cancelled row that's done but has no bookingId, so the
 * normal creation loop below picks it back up this run. Safe to call every
 * run: a correctly-done row always has a bookingId and is untouched.
 */
function unstickBookingDoneWithoutId_(items) {
  let count = 0;
  items.forEach(b => {
    if (!b.done) return;
    if (/ยกเลิก|cancel/i.test(b.room)) return; // cancelled stays never need one — leave as-is
    if (getApartmenteryBookingId_(b.resId)) return; // correctly done
    Logger.log(`unstickBookingDoneWithoutId_: ${b.resId} (${b.room}) was marked done with no ` +
      `apartmentery bookingId — un-marking so it gets retried this run.`);
    setBookingDone(b.resId, false);
    b.done = false; // so the loop below picks it up in this same pass, not just next run
    count++;
  });
  if (count > 0) {
    Logger.log(`unstickBookingDoneWithoutId_: un-stuck ${count} row(s).`);
  }
  return count;
}

/**
 * Mirror image of unstickBookingDoneWithoutId_: a row that already HAS an
 * apartmentery bookingId but is still marked done=false. Confirmed
 * 2026-07-16 — backfillMissingApartmenteryBookings() (and the collision
 * recovery path) wrote the bookingId column but never called
 * setBookingDone(), so 12 rows sat in Booking To Add forever with a
 * complete bookingId. Auto-tick them here every run so any future path
 * that sets a bookingId without also marking done self-heals instead of
 * requiring another manual investigation.
 */
function tickBookingDoneWithId_(items) {
  let count = 0;
  items.forEach(b => {
    if (b.done) return;
    if (!getApartmenteryBookingId_(b.resId)) return;
    Logger.log(`tickBookingDoneWithId_: ${b.resId} (${b.room}) has an apartmentery bookingId ` +
      `but was still marked done=false — marking done.`);
    setBookingDone(b.resId, true);
    b.done = true;
    count++;
  });
  if (count > 0) {
    Logger.log(`tickBookingDoneWithId_: ticked ${count} row(s).`);
  }
  return count;
}

function autoCreateApartmenteryBookings() {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  const items = getBookingToAdd_(ss, todayStr);

  unstickBookingDoneWithoutId_(items);
  tickBookingDoneWithId_(items);

  Logger.log(`autoCreateApartmenteryBookings: ${items.length} total rows from getBookingToAdd_ — ` +
    JSON.stringify(items.map(b => ({ resId: b.resId, room: b.room, guest: b.guest, done: b.done }))));

  const result = { created: 0, skipped: 0, sessionExpired: false, errors: [] };

  // apartmentery refuses to create a booking whose startDate equals another
  // booking's checkout date on the same room (same-day turnover) — confirmed
  // 2026-07-09, this is the cause of the generic HTTP 500 "Oops, an error
  // occurred" page apartmentery returns for these, with no way around it via
  // a different payload (reproduced the same block with a real browser
  // submission too). Nathan's existing manual workaround is to shorten the
  // outgoing booking's endDate by 1 day on apartmentery before creating the
  // new one — automate that instead of just skipping. Build a per-room map
  // of checkout-date -> resId up front so we know which outgoing booking to
  // shrink for each turnover we hit below.
  const outgoingByRoom = {};
  // Back-edge of the same problem: the NEW booking's own checkout can equal
  // ANOTHER booking's checkin (e.g. Mike 20->22, J Barber checks in 22).
  // Shrinking the outgoing side doesn't help here since there's no existing
  // apartmentery booking to shrink — instead we send the new booking with an
  // endDate 1 day short of the real checkout, so it doesn't collide with the
  // next guest's start date. The real checkout in Sheet1 is untouched.
  const incomingByRoom = {};
  items.forEach(x => {
    if (/ยกเลิก|cancel/i.test(x.room)) return; // cancelled stays don't occupy the room
    if (!x.checkout) return;
    const rn = roomNum_(x.room);
    if (!outgoingByRoom[rn]) outgoingByRoom[rn] = {};
    outgoingByRoom[rn][x.checkout] = x.resId;
    if (!x.checkin) return;
    if (!incomingByRoom[rn]) incomingByRoom[rn] = {};
    incomingByRoom[rn][x.checkin] = x.resId;
  });

  for (const b of items) {
    if (b.done) { Logger.log(`skip ${b.resId} (${b.room}): already marked done`); continue; }
    // Cancelled bookings ("204 Elegance ยกเลิก") never need an apartmentery booking.
    if (/ยกเลิก|cancel/i.test(b.room)) { Logger.log(`skip ${b.resId} (${b.room}): cancelled`); continue; }
    if (getApartmenteryBookingId_(b.resId)) { Logger.log(`skip ${b.resId} (${b.room}): already has apartmentery bookingId`); continue; }

    const roomOutgoing = outgoingByRoom[roomNum_(b.room)];
    const outgoingResId = roomOutgoing && roomOutgoing[b.checkin];
    if (outgoingResId && outgoingResId !== b.resId) {
      const outgoingAptId = getApartmenteryBookingId_(outgoingResId);
      if (!outgoingAptId) {
        // Outgoing booking hasn't been created on apartmentery yet, so there's
        // nothing to shrink — creating the new one would still 500. Skip for
        // now; next run picks this up once the outgoing booking exists.
        Logger.log(`skip ${b.resId} (${b.room}): same-day turnover with ${outgoingResId}, ` +
          `but that booking has no apartmentery bookingId yet — nothing to shrink. Will retry next run.`);
        result.skipped++;
        continue;
      }
      try {
        const newEndDate = _dateMinusOneDay_(b.checkin);
        Logger.log(`same-day turnover: shrinking ${outgoingResId} (${b.room}) apartmentery ` +
          `bookingId ${outgoingAptId} endDate to ${newEndDate} before creating ${b.resId}`);
        updateApartmenteryBookingEndDateForRoom(b.room, outgoingAptId, newEndDate);
        // Mimic a browser reloading the calendar view between the shrink
        // and the add — see refreshApartmenteryUnitCalendarForRoom's
        // comment in ApartmenteryClient.gs for why.
        refreshApartmenteryUnitCalendarForRoom(b.room);
      } catch (err) {
        if (isApartmenterySessionExpiredError(err)) {
          Logger.log(`SESSION EXPIRED while shrinking endDate for ${outgoingResId} (${b.room}): ${err.message}`);
          result.sessionExpired = true;
          break; // stop the whole run — see file header
        }
        Logger.log(`skip ${b.resId} (${b.room}): failed to shrink outgoing booking ${outgoingResId}'s ` +
          `endDate — ${err.message}`);
        result.skipped++;
        result.errors.push({ resId: b.resId, room: b.room, error: err.message });
        continue;
      }
    }

    // Back-edge check: does another booking in this room check in exactly
    // when this one checks out? If so, apartmentery will refuse to create
    // this booking against the full range — shrink the endDate we SEND by
    // 1 day (Sheet1's real checkout is untouched).
    const roomIncoming = incomingByRoom[roomNum_(b.room)];
    const incomingResId = b.checkout && roomIncoming && roomIncoming[b.checkout];
    const effectiveEndDate = (incomingResId && incomingResId !== b.resId)
      ? _dateMinusOneDay_(b.checkout)
      : (b.checkout || '');
    if (incomingResId && incomingResId !== b.resId) {
      Logger.log(`same-day turnover (back edge): ${b.resId} (${b.room}) checkout ${b.checkout} ` +
        `matches ${incomingResId}'s checkin — sending apartmentery endDate ${effectiveEndDate} instead`);
    }

    Logger.log(`attempting booking for ${b.resId} room ${b.room} guest ${b.guest}`);

    try {
      // Apartmentery customerName follows Nathan's original manual naming
      // convention "Guest Name / Channel" (e.g. "Pranee Antov / Expedia")
      // — the automation was previously sending just the bare guest name.
      const guestNameWithChannel = b.channel ? `${b.guest} / ${b.channel}` : b.guest;

      const created = createApartmenteryBookingForRoom(b.room, {
        startDate: b.checkin,
        endDate: effectiveEndDate,
        guestName: guestNameWithChannel,
        note: `${b.channel} ${b.resId}`.trim()
      });

      if (created && created.skipped) {
        Logger.log(`skip ${b.resId} (${b.room}): ${created.reason}`);
        result.skipped++;
        continue;
      }

      const writeResult = setApartmenteryBookingId_(b.resId, created.bookingId);
      if (!writeResult.ok) {
        // Uniqueness guard refused the write (bookingId already belongs to
        // another resId) — apartmentery DID create a real booking for this
        // guest just now, but our follow-up guest+date lookup grabbed the
        // wrong id for it (see setApartmenteryBookingId_ comment). Leave
        // b.done unset so this resId is retried next run; log loudly since
        // there's now a real, un-tracked booking on apartmentery for this
        // guest that needs manual linking.
        Logger.log(`created a real apartmentery booking for ${b.resId} (${b.room}, ${b.guest}) but ` +
          `could not record its id (${created.bookingId}) — ${writeResult.error} Needs manual lookup ` +
          `on apartmentery to find the correct bookingId and link it by hand.`);
        result.skipped++;
        result.errors.push({ resId: b.resId, guest: b.guest, room: b.room, error: writeResult.error });
        continue;
      }
      setBookingDone(b.resId, true);
      Logger.log(`created ${b.resId} (${b.room}) -> apartmentery bookingId ${created.bookingId}`);
      result.created++;

    } catch (err) {
      if (isApartmenterySessionExpiredError(err)) {
        Logger.log(`SESSION EXPIRED while creating booking for ${b.resId} (${b.room}): ${err.message}`);
        result.sessionExpired = true;
        break; // stop the whole run — see file header
      }

      if (err.isCollision) {
        // Already exists in apartmentery (created manually, or by a prior
        // run that errored after the create but before setApartmenteryBookingId_
        // was recorded) — recover its bookingId from the unit's calendar
        // instead of leaving this booking stuck forever on every hourly run.
        // Ported from backfillMissingApartmenteryBookings() — search by the
        // RAW guest name, not guestNameWithChannel, since a manually-entered
        // booking's title in apartmentery almost certainly doesn't have our
        // "/ Airbnb" / "/ Booking" suffix appended.
        let recoveredId = null;
        try {
          recoveredId = findApartmenteryBookingIdForRoomByGuest_(b.room, b.guest, b.checkin);
        } catch (lookupErr) {
          if (isApartmenterySessionExpiredError(lookupErr)) {
            Logger.log(`SESSION EXPIRED while recovering bookingId for ${b.resId} (${b.room}): ${lookupErr.message}`);
            result.sessionExpired = true;
            break;
          }
          Logger.log(`recovery lookup failed for ${b.resId} (${b.room}): ${lookupErr.message}`);
        }

        if (recoveredId) {
          const writeResult = setApartmenteryBookingId_(b.resId, recoveredId);
          if (!writeResult.ok) {
            Logger.log(`collision-recovery for ${b.resId} (${b.room}, ${b.guest}) found bookingId ` +
              `${recoveredId}, but ${writeResult.error} Needs manual review.`);
            result.skipped++;
            result.errors.push({ resId: b.resId, guest: b.guest, room: b.room, error: writeResult.error });
            continue;
          }
          setBookingDone(b.resId, true);
          Logger.log(`recovered ${b.resId} (${b.room}) -> existing apartmentery bookingId ${recoveredId} (was already in apartmentery, not newly created)`);
          result.created++; // counts toward "now has a bookingId", same outcome for the sheet
          result.recovered = (result.recovered || 0) + 1;
          continue;
        }

        Logger.log(`skip ${b.resId} (${b.room}): collision reported but no exact guest+date match found in unit's calendar.` +
          `${(outgoingResId && outgoingResId !== b.resId) ? ' This already tried the same-day-turnover shrink + calendar-refresh — if it\'s still colliding after that, it likely needs to be created manually on apartmentery.' : ' Not a same-day-turnover case — this collision is against some other booking, inspect manually.'}` +
          ` (${err.apartmenteryError})`);
        result.skipped++;
        result.errors.push({ resId: b.resId, guest: b.guest, room: b.room, error: `collision, but couldn't recover bookingId — ${err.apartmenteryError}` });
        continue;
      }

      Logger.log(`ERROR creating booking for ${b.resId} (${b.room}): ${err.message}`);
      result.errors.push({ resId: b.resId, guest: b.guest, room: b.room, error: err.apartmenteryError || err.message });
      // Non-session errors (e.g. one bad row) don't stop the batch —
      // continue so one problem booking doesn't block everything else.
    }
  }

  Logger.log('autoCreateApartmenteryBookings result: ' + JSON.stringify(result));
  return result;
}

/**
 * Phase 2/3 automation: create any missing invoices + receipts for
 * matched payouts. Joins invoice items back to their Sheet1 booking row
 * (and therefore its apartmentery bookingId) using the same matchKeys_
 * fuzzy-match already used elsewhere in this project.
 */
function autoCreateApartmenteryInvoicesAndReceipts() {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');

  const bookingItems = getBookingToAdd_(ss, todayStr);
  const invoiceItems = getInvoiceToCreate_(ss, todayStr);

  // Build matchKey -> {resId, guest} index from booking items (only ones
  // that actually have an apartmentery bookingId already — no point
  // matching to a booking apartmentery doesn't know about yet).
  const keyToResId = {};
  bookingItems.forEach(b => {
    const aptId = getApartmenteryBookingId_(b.resId);
    if (!aptId) return;
    (b.matchKeys || []).forEach(k => { if (!keyToResId[k]) keyToResId[k] = { resId: b.resId, guest: b.guest }; });
  });

  const result = { created: 0, skipped: 0, sessionExpired: false, errors: [] };

  for (const inv of invoiceItems) {
    if (inv.done) continue;
    if (String(inv.room).indexOf('ไม่ทราบห้อง') >= 0) { result.skipped++; continue; } // unresolved room, needs manual review

    let resId = null;
    for (const k of (inv.matchKeys || [])) {
      const candidate = keyToResId[k];
      if (!candidate) continue;
      // The 'cr:' key only encodes room + a ±4-day window around checkin —
      // no guest name at all. Two different guests' stays in the same room
      // only ~6 days apart can both fall inside each other's window (e.g.
      // one checks out 07-15, the next checks in 07-17), so a bare 'cr:'
      // hit can point at a completely different person's booking. Confirmed
      // 2026-07-16: Natphatsorn wongwai's payout (room 300, checkin 07-11)
      // resolved to Livio Castelli's booking 321720 (room 300, checkin
      // 07-17) this way, because Natphatsorn's own resId had no bookingId
      // yet and so wasn't in this index under her own name key. Guard
      // against this by requiring the candidate's guest name to actually
      // resemble the invoice item's guest before accepting the match.
      if (!_namesMatchIgnoringOrder_(inv.guest, candidate.guest) &&
          !_namesMatchIgnoringOrder_(candidate.guest, inv.guest)) {
        continue;
      }
      resId = candidate.resId;
      break;
    }
    if (!resId) { result.skipped++; continue; } // no linked apartmentery booking yet

    const aptBookingId = getApartmenteryBookingId_(resId);
    if (!aptBookingId) { result.skipped++; continue; }

    try {
      const outcome = processPayoutToReceiptForRoom(inv.room, aptBookingId, inv.net, todayStr);
      if (outcome && outcome.skipped) {
        result.skipped++;
        continue;
      }
      setInvoiceDone(inv.invoiceKey, true);
      result.created++;

    } catch (err) {
      if (isApartmenterySessionExpiredError(err)) {
        result.sessionExpired = true;
        break;
      }
      result.errors.push({ invoiceKey: inv.invoiceKey, guest: inv.guest, room: inv.room, resId: resId, aptBookingId: aptBookingId, error: err.message });
    }
  }

  return result;
}

/**
 * Backfill for EXISTING rows that already have no apartmentery bookingId
 * but will NEVER be picked up by autoCreateApartmenteryBookings(), because
 * that function skips any resId already marked done in booking_done_v1 —
 * and that flag predates apartmentery automation, so basically all of
 * Feb–Jul 2026's history is already marked done there for unrelated
 * reasons (general booking-todo UI), even though none of them ever got
 * an apartmentery booking created. Confirmed 2026-07-12: 130+ rows in
 * Sheet1 have an empty "Apartmentery Booking ID" column.
 *
 * This is the SAME logic as autoCreateApartmenteryBookings() (same-day
 * turnover handling included) minus the `if (b.done) continue` skip, and
 * it deliberately does NOT call setBookingDone() — that flag is shared
 * with the unrelated booking-todo UI and shouldn't be touched by this.
 *
 * Safe to re-run / resumable: only ever touches rows that (a) aren't
 * cancelled and (b) don't already have an Apartmentery Booking ID, so a
 * run that stops early (time budget or session expiry) picks up exactly
 * where it left off next time.
 *
 * Has a ~5 min runtime budget to stay under Apps Script's 6-min execution
 * limit for ~130 sequential HTTP calls to apartmentery.com — logs how many
 * are left if it has to stop early; just run it again to continue.
 */
function backfillMissingApartmenteryBookings() {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  const items = getBookingToAdd_(ss, todayStr);
  const startTime = Date.now();
  const MAX_RUNTIME_MS = 5 * 60 * 1000;

  Logger.log(`backfillMissingApartmenteryBookings: ${items.length} total rows from getBookingToAdd_`);

  const result = { created: 0, skipped: 0, sessionExpired: false, timedOut: false, errors: [] };

  // Same same-day-turnover map as autoCreateApartmenteryBookings — built
  // from ALL items (not just the ones missing an id) so a turnover against
  // an already-created booking still resolves correctly.
  const outgoingByRoom = {};
  // Back-edge mirror of the same check — see comment in
  // autoCreateApartmenteryBookings for why this is needed.
  const incomingByRoom = {};
  items.forEach(x => {
    if (/ยกเลิก|cancel/i.test(x.room)) return;
    if (!x.checkout) return;
    const rn = roomNum_(x.room);
    if (!outgoingByRoom[rn]) outgoingByRoom[rn] = {};
    outgoingByRoom[rn][x.checkout] = x.resId;
    if (!x.checkin) return;
    if (!incomingByRoom[rn]) incomingByRoom[rn] = {};
    incomingByRoom[rn][x.checkin] = x.resId;
  });

  for (const b of items) {
    if (Date.now() - startTime > MAX_RUNTIME_MS) {
      Logger.log('backfillMissingApartmenteryBookings: stopping early (time budget) — re-run to continue with the rest.');
      result.timedOut = true;
      break;
    }

    if (!b.resId) continue;
    if (/ยกเลิก|cancel/i.test(b.room)) { Logger.log(`skip ${b.resId} (${b.room}): cancelled`); continue; }
    if (getApartmenteryBookingId_(b.resId)) { continue; } // already backfilled or created normally

    const roomOutgoing = outgoingByRoom[roomNum_(b.room)];
    const outgoingResId = roomOutgoing && roomOutgoing[b.checkin];
    if (outgoingResId && outgoingResId !== b.resId) {
      const outgoingAptId = getApartmenteryBookingId_(outgoingResId);
      if (!outgoingAptId) {
        Logger.log(`skip ${b.resId} (${b.room}): same-day turnover with ${outgoingResId}, but that booking has no apartmentery bookingId yet — will retry next run.`);
        result.skipped++;
        continue;
      }
      try {
        const newEndDate = _dateMinusOneDay_(b.checkin);
        Logger.log(`same-day turnover: shrinking ${outgoingResId} (${b.room}) apartmentery bookingId ${outgoingAptId} endDate to ${newEndDate} before creating ${b.resId}`);
        updateApartmenteryBookingEndDateForRoom(b.room, outgoingAptId, newEndDate);
        refreshApartmenteryUnitCalendarForRoom(b.room);
      } catch (err) {
        if (isApartmenterySessionExpiredError(err)) {
          Logger.log(`SESSION EXPIRED while shrinking endDate for ${outgoingResId} (${b.room}): ${err.message}`);
          result.sessionExpired = true;
          break;
        }
        Logger.log(`skip ${b.resId} (${b.room}): failed to shrink outgoing booking ${outgoingResId}'s endDate — ${err.message}`);
        result.skipped++;
        result.errors.push({ resId: b.resId, room: b.room, error: err.message });
        continue;
      }
    }

    const roomIncoming = incomingByRoom[roomNum_(b.room)];
    const incomingResId = b.checkout && roomIncoming && roomIncoming[b.checkout];
    const effectiveEndDate = (incomingResId && incomingResId !== b.resId)
      ? _dateMinusOneDay_(b.checkout)
      : (b.checkout || '');
    if (incomingResId && incomingResId !== b.resId) {
      Logger.log(`[backfill] same-day turnover (back edge): ${b.resId} (${b.room}) checkout ${b.checkout} ` +
        `matches ${incomingResId}'s checkin — sending apartmentery endDate ${effectiveEndDate} instead`);
    }

    Logger.log(`[backfill] attempting booking for ${b.resId} room ${b.room} guest ${b.guest}`);

    try {
      const guestNameWithChannel = b.channel ? `${b.guest} / ${b.channel}` : b.guest;
      const created = createApartmenteryBookingForRoom(b.room, {
        startDate: b.checkin,
        endDate: effectiveEndDate,
        guestName: guestNameWithChannel,
        note: `${b.channel} ${b.resId}`.trim()
      });

      if (created && created.skipped) {
        Logger.log(`skip ${b.resId} (${b.room}): ${created.reason}`);
        result.skipped++;
        continue;
      }

      const writeResult = setApartmenteryBookingId_(b.resId, created.bookingId);
      if (!writeResult.ok) {
        Logger.log(`[backfill] created a real apartmentery booking for ${b.resId} (${b.room}, ${b.guest}) ` +
          `but could not record its id (${created.bookingId}) — ${writeResult.error} Needs manual lookup.`);
        result.skipped++;
        result.errors.push({ resId: b.resId, guest: b.guest, room: b.room, error: writeResult.error });
        continue;
      }
      setBookingDone(b.resId, true);
      Logger.log(`[backfill] created ${b.resId} (${b.room}) -> apartmentery bookingId ${created.bookingId}`);
      result.created++;

    } catch (err) {
      if (isApartmenterySessionExpiredError(err)) {
        Logger.log(`SESSION EXPIRED while creating booking for ${b.resId} (${b.room}): ${err.message}`);
        result.sessionExpired = true;
        break;
      }

      if (err.isCollision) {
        // Already exists in apartmentery (created manually pre-automation) —
        // recover its bookingId from the unit's calendar instead of erroring.
        // NOTE: search by the RAW guest name, not guestNameWithChannel — a
        // manually-entered booking's title in apartmentery almost certainly
        // doesn't have our "/ Airbnb" / "/ Booking" suffix appended, so
        // searching for the suffixed string found zero matches on the first
        // attempt (confirmed 2026-07-12: 0/122 recovered) even though the
        // bookings clearly exist (that's what caused the collision at all).
        // title.indexOf(guestName) is a substring match, so the raw name
        // still matches titles that DO have a suffix too — strictly safer.
        let recoveredId = null;
        try {
          recoveredId = findApartmenteryBookingIdForRoomByGuest_(b.room, b.guest, b.checkin);
        } catch (lookupErr) {
          if (isApartmenterySessionExpiredError(lookupErr)) {
            Logger.log(`SESSION EXPIRED while recovering bookingId for ${b.resId} (${b.room}): ${lookupErr.message}`);
            result.sessionExpired = true;
            break;
          }
          Logger.log(`recovery lookup failed for ${b.resId} (${b.room}): ${lookupErr.message}`);
        }

        if (recoveredId) {
          const writeResult = setApartmenteryBookingId_(b.resId, recoveredId);
          if (!writeResult.ok) {
            Logger.log(`[backfill] collision-recovery for ${b.resId} (${b.room}, ${b.guest}) found bookingId ` +
              `${recoveredId}, but ${writeResult.error} Needs manual review.`);
            result.skipped++;
            result.errors.push({ resId: b.resId, guest: b.guest, room: b.room, error: writeResult.error });
            continue;
          }
          setBookingDone(b.resId, true);
          Logger.log(`[backfill] recovered ${b.resId} (${b.room}) -> existing apartmentery bookingId ${recoveredId} (was already in apartmentery, not newly created)`);
          result.created++; // counts toward "now has a bookingId", same outcome for the sheet
          result.recovered = (result.recovered || 0) + 1;
          continue;
        }

        Logger.log(`skip ${b.resId} (${b.room}): collision reported but no exact guest+date match found in unit's calendar.` +
          `${(outgoingResId && outgoingResId !== b.resId) ? ' This already tried the same-day-turnover shrink + calendar-refresh — if it\'s still colliding after that, it likely needs to be created manually on apartmentery.' : ' Not a same-day-turnover case — this collision is against some other booking, inspect manually.'}` +
          ` (${err.apartmenteryError})`);
        result.skipped++;
        result.errors.push({ resId: b.resId, guest: b.guest, room: b.room, error: `collision, but couldn't recover bookingId — ${err.apartmenteryError}` });
        continue;
      }

      Logger.log(`ERROR creating booking for ${b.resId} (${b.room}): ${err.message} — ${err.apartmenteryError || ''}`);
      result.errors.push({ resId: b.resId, guest: b.guest, room: b.room, error: err.apartmenteryError || err.message });
    }
  }

  Logger.log('backfillMissingApartmenteryBookings result: ' + JSON.stringify(result));
  return result;
}

/**
 * DIAGNOSTIC ONLY — not part of the automation flow. Run manually when
 * debugging why createApartmenteryBooking keeps reporting collisions that
 * findApartmenteryBookingIdForRoomByGuest_ can't recover (confirmed
 * 2026-07-12: 0 recoveries across 3 backfill attempts, even after fixing
 * the guest-name-suffix bug). Fetches the unit's booking-listing page and
 * logs every {title, start, end, url} event exactly as embedded in the
 * page's fullCalendar array, so we can see what's ACTUALLY occupying the
 * unit's calendar instead of guessing further.
 *
 * Usage: run with a Sheet1-style room string, e.g.
 *   debugApartmenteryUnitCalendar('205 Allure')
 */
/** No-argument wrapper for the Run button — GAS's editor can't pass parameters. */
/**
 * READ-ONLY. Confirmed 2026-07-16: Nathan reported apartmentery bookings
 * looking room-mismatched shortly after a manual runApartmenteryAutomation()
 * run. That run's invoice phase errored cleanly for one stale bookingId
 * (Andrea Mastropietro / 326665 — apartmentery rejected it outright), but
 * a bookingId that's stale WITHOUT being outright invalid — i.e. it still
 * resolves, just to a *different* booking than the one on this row — would
 * NOT have errored. It would have silently created an invoice against
 * whatever booking that ID now points to. This function does not write
 * anything anywhere; it only reports. Do not "fix" anything based on its
 * output without Nathan reviewing first, since apartmentery itself may
 * already be wrong for reasons upstream of this sheet.
 *
 * For every room, pulls the full apartmentery calendar (bookingId -> title
 * + start date). For every Sheet1 row with a recorded Apartmentery Booking
 * ID, checks:
 *   - does that ID appear under its OWN room's calendar? (expected)
 *   - does it appear under a DIFFERENT room's calendar instead? (the
 *     "ห้องมั่ว" case — flagged loudly)
 *   - does it not appear in ANY room's calendar at all? (dead ID, same
 *     class of problem as Andrea's, just not yet hit by the invoice phase)
 *   - if it appears in the right room, does the guest name on that
 *     apartmentery event roughly match the guest name in Sheet1?
 */
/**
 * Applies only the stragglers from findCandidatesForUnresolved20260716
 * that were cross-checked by hand against the 94 rows
 * fixAllApartmenteryBookingIdsComprehensive20260716 already fixed, to
 * confirm the candidate id isn't already claimed by a different resId.
 * The other stragglers (Avto Dagdelen 04-05, Miles Consengco 06-09,
 * Pornpawit Boon 06-16) had their only candidate already claimed by a
 * sibling resId for the same guest's other stay — meaning that specific
 * stay most likely was never created in apartmentery at all, not just
 * mis-recorded. Those need a human to check apartmentery directly, not
 * an automatic overwrite.
 */
function applyConfirmedSafeStragglers20260716() {
  const rows = [
    { resId: 'ABB-johnzambra-20260428', id: '314905' },
    { resId: 'ABB-premmehta-20260501',  id: '315888' },
    { resId: 'ABB-kgotlellom-20260528', id: '317758' },
    { resId: 'ABB-saeidmickm-20260610', id: '321719' },
    { resId: 'ABB-errolcox-20260608',   id: '321724' } // not 100% cross-checked — see log after running
  ];
  rows.forEach(r => {
    const before = getApartmenteryBookingId_(r.resId);
    setApartmenteryBookingId_(r.resId, r.id);
    Logger.log(`applyConfirmedSafeStragglers20260716: ${r.resId} "${before}" -> "${r.id}"`);
  });
  Logger.log('Done. Re-run auditAllApartmenteryBookingIds to confirm these now show OK and nothing new broke.');
}

function findCandidatesForUnresolved20260716() {
  // The 11 real stragglers left after fixAllApartmenteryBookingIdsComprehensive20260716
  // (excludes the 3 cancelled-booking rows, which are harmless — automation
  // already skips any room containing "ยกเลิก"/"cancel").
  const rows = [
    { resId: 'ABB-e585a82c20-20260219', room: '103', guest: '全, 桂珍' },
    { resId: 'ABB-e5a698e88a-20260403', room: '103', guest: '妘芮 林 Yunjui Lin' },
    { resId: 'ABB-avtodagdel-20260405', room: '203', guest: 'Avto Dagdelen' },
    { resId: 'ABB-johnzambra-20260428', room: '108', guest: 'John Zambrana' },
    { resId: 'ABB-premmehta-20260501',  room: '108', guest: 'Prem Mehta' },
    { resId: 'ABB-lataviaant-20260512', room: '214', guest: "La'Tavia Antrice" },
    { resId: 'ABB-kgotlellom-20260528', room: '203', guest: 'Kgotlello Masemola' },
    { resId: 'ABB-errolcox-20260608',   room: '210', guest: 'Errol Cox' },
    { resId: 'ABB-milesconse-20260609', room: '204', guest: 'Miles Consengco' },
    { resId: 'ABB-saeidmickm-20260610', room: '108', guest: 'Saeid Mick Momtahan' },
    { resId: 'TRP-pornpawitb-20260616', room: '103', guest: 'Pornpawit Boon' }
  ];

  const roomsNeeded = [...new Set(rows.map(r => r.room))];
  const calendarByRoom = {};
  roomsNeeded.forEach(room => {
    const unit = getApartmenteryUnitForRoom(room);
    if (!unit) return;
    const path = `/user/branch/${unit.branchId}/unit/${unit.unitId}/booking`;
    const response = _apartmenteryFetch_(path, { method: 'get' });
    const html = response.getContentText();
    const blockRe = /\{\s*title:\s*'((?:[^'\\]|\\.)*)'[\s\S]*?start:\s*'([^']*)'[\s\S]*?url:\s*'([^']*)'\s*\}/g;
    let m;
    const events = [];
    while ((m = blockRe.exec(html)) !== null) {
      const idMatch = m[3].match(/\/booking\/(\d+)/);
      events.push({ title: m[1], start: _apartmenteryCalendarDateToIso_(m[2]), id: idMatch ? idMatch[1] : null });
    }
    calendarByRoom[room] = events;
  });

  rows.forEach(r => {
    const events = calendarByRoom[r.room] || [];
    const currentId = getApartmenteryBookingId_(r.resId);
    // Loose word-overlap match, ignoring date — surfaces every plausible
    // candidate in that room regardless of which date it's actually on.
    const candidates = events.filter(e => _namesMatchIgnoringOrder_(r.guest, e.title));
    Logger.log(`--- ${r.resId} (guest "${r.guest}", room ${r.room}, currently stored: ${currentId}) ---`);
    if (candidates.length === 0) {
      Logger.log(`  no name match found anywhere in room ${r.room}'s calendar at all — guest may be in a different room, or spelled very differently on apartmentery.`);
    } else {
      candidates.forEach(c => Logger.log(`  candidate: id=${c.id} start=${c.start} title="${c.title}"`));
    }
  });
}

function auditAllApartmenteryBookingIds() {
  Logger.log('auditAllApartmenteryBookingIds: READ-ONLY — pulling calendars for all rooms...');

  // Step 1: pull every room's calendar, build bookingId -> {room, title, start}
  const byBookingId = {};
  const roomEventCounts = {};
  Object.keys(ROOM_TO_UNIT_ID).forEach(room => {
    const unit = getApartmenteryUnitForRoom(room);
    if (!unit) return;
    const path = `/user/branch/${unit.branchId}/unit/${unit.unitId}/booking`;
    let html;
    try {
      const response = _apartmenteryFetch_(path, { method: 'get' });
      html = response.getContentText();
    } catch (e) {
      Logger.log(`auditAllApartmenteryBookingIds: FAILED to fetch calendar for room ${room}: ${e.message}`);
      return;
    }
    const blockRe = /\{\s*title:\s*'((?:[^'\\]|\\.)*)'[\s\S]*?start:\s*'([^']*)'[\s\S]*?url:\s*'([^']*)'\s*\}/g;
    let m;
    let count = 0;
    while ((m = blockRe.exec(html)) !== null) {
      const title = m[1];
      const start = _apartmenteryCalendarDateToIso_(m[2]);
      const idMatch = m[3].match(/\/booking\/(\d+)/);
      if (!idMatch) continue;
      const bookingId = idMatch[1];
      count++;
      // A bookingId can legitimately appear more than once per room's page
      // (e.g. rendered on multiple months) — keep the first, they should
      // be identical anyway.
      if (!byBookingId[bookingId]) {
        byBookingId[bookingId] = { room: room, title: title, start: start };
      }
    }
    roomEventCounts[room] = count;
    Logger.log(`auditAllApartmenteryBookingIds: room ${room} (unit ${unit.unitId}) — ${count} calendar events`);
  });
  Logger.log(`auditAllApartmenteryBookingIds: pulled ${Object.keys(byBookingId).length} unique bookingIds across all rooms.`);

  // Step 2: walk every Sheet1 row with a recorded bookingId and cross-check.
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName('Sheet1');
  const data = src.getDataRange().getValues();
  const header = data[0];
  const idx = indexMap_(header, ['ResId', 'เลขห้อง', 'ชื่อแขก', APARTMENTERY_BOOKING_ID_COL_HEADER]);
  if (idx.ResId < 0 || idx[APARTMENTERY_BOOKING_ID_COL_HEADER] < 0) {
    Logger.log('auditAllApartmenteryBookingIds: required columns not found — aborting.');
    return;
  }

  let okCount = 0, wrongRoomCount = 0, deadCount = 0, guestMismatchCount = 0;
  const wrongRoom = [], dead = [], guestMismatch = [];

  for (let i = 1; i < data.length; i++) {
    const resId = String(data[i][idx.ResId] || '').trim();
    const bookingId = String(data[i][idx[APARTMENTERY_BOOKING_ID_COL_HEADER]] || '').trim();
    if (!resId || !bookingId) continue;
    const roomRaw = idx['เลขห้อง'] >= 0 ? String(data[i][idx['เลขห้อง']] || '').trim() : '';
    const guest = idx['ชื่อแขก'] >= 0 ? String(data[i][idx['ชื่อแขก']] || '').trim() : '';
    const expectedRoom = (roomRaw.match(/^\d+/) || [''])[0];

    const found = byBookingId[bookingId];
    if (!found) {
      deadCount++;
      dead.push(`resId=${resId} room=${roomRaw} guest="${guest}" bookingId=${bookingId} — NOT FOUND in any room's calendar`);
      continue;
    }
    if (found.room !== expectedRoom) {
      wrongRoomCount++;
      wrongRoom.push(`resId=${resId} guest="${guest}" bookingId=${bookingId} — Sheet1 says room ${expectedRoom}, but apartmentery has this bookingId under room ${found.room} as "${found.title}" (${found.start})`);
      continue;
    }
    if (guest && !_namesMatchIgnoringOrder_(guest, found.title)) {
      guestMismatchCount++;
      guestMismatch.push(`resId=${resId} room=${expectedRoom} bookingId=${bookingId} — Sheet1 guest "${guest}" vs apartmentery title "${found.title}"`);
      continue;
    }
    okCount++;
  }

  Logger.log('=== auditAllApartmenteryBookingIds SUMMARY ===');
  Logger.log(`OK: ${okCount}`);
  Logger.log(`WRONG ROOM (bookingId lives under a different room than Sheet1 expects): ${wrongRoomCount}`);
  Logger.log(`DEAD (bookingId not found in any room's calendar): ${deadCount}`);
  Logger.log(`GUEST NAME MISMATCH (right room, but guest name doesn't match): ${guestMismatchCount}`);

  if (wrongRoom.length) {
    Logger.log('--- WRONG ROOM DETAIL ---');
    wrongRoom.forEach(l => Logger.log(l));
  }
  if (dead.length) {
    Logger.log('--- DEAD DETAIL ---');
    dead.forEach(l => Logger.log(l));
  }
  if (guestMismatch.length) {
    Logger.log('--- GUEST MISMATCH DETAIL ---');
    guestMismatch.forEach(l => Logger.log(l));
  }

  return { okCount, wrongRoomCount, deadCount, guestMismatchCount };
}

function debugApartmenteryUnitCalendar205() {
  return debugApartmenteryUnitCalendar('205 Allure');
}
function debugApartmenteryUnitCalendar204() {
  return debugApartmenteryUnitCalendar('204 Elegance');
}
function debugApartmenteryUnitCalendar203() {
  return debugApartmenteryUnitCalendar('203 Allure');
}
function debugApartmenteryUnitCalendar103() {
  return debugApartmenteryUnitCalendar('103 Elegance');
}

function debugApartmenteryUnitCalendar(roomRaw, monthFilter) {
  const unit = getApartmenteryUnitForRoom(roomRaw);
  if (!unit) {
    Logger.log(`debugApartmenteryUnitCalendar: room "${roomRaw}" not found in ROOM_TO_UNIT_ID.`);
    return;
  }
  const path = `/user/branch/${unit.branchId}/unit/${unit.unitId}/booking`;
  const response = _apartmenteryFetch_(path, { method: 'get' });
  Logger.log(`debugApartmenteryUnitCalendar: GET ${path} -> HTTP ${response.getResponseCode()}`);

  const html = response.getContentText();
  // Same pattern as _findBookingIdByGuestName_ / _findBookingIdByGuestNameAndDate_
  // (already proven to match this page's structure) — title/start/url only,
  // since requiring an 'end' field here too risked a false "0 events" result
  // if events don't reliably include one in this exact position.
  const blockRe = /\{\s*title:\s*'((?:[^'\\]|\\.)*)'[\s\S]*?start:\s*'([^']*)'[\s\S]*?url:\s*'([^']*)'\s*\}/g;
  let m;
  const events = [];
  while ((m = blockRe.exec(html)) !== null) {
    events.push({ title: m[1], start: m[2], url: m[3] });
  }
  Logger.log(`debugApartmenteryUnitCalendar: found ${events.length} total events for unit ${unit.unitId}`);

  // Default filter: Feb-Apr 2026, since that's usually the range worth
  // inspecting manually — pass monthFilter=null to see everything instead.
  const filtered = monthFilter === null ? events : events.filter(e => {
    const iso = _apartmenteryCalendarDateToIso_(e.start);
    return iso && iso >= '2026-02-01' && iso < '2026-04-06';
  });
  Logger.log(`debugApartmenteryUnitCalendar: showing ${filtered.length} events (Feb 1 - Apr 5 2026 unless monthFilter=null was passed):`);
  filtered.forEach((e, i) => Logger.log(`  [${i}] title="${e.title}" start=${e.start} url=${e.url}`));

  if (events.length === 0) {
    Logger.log('debugApartmenteryUnitCalendar: no events matched the expected regex — logging first 3000 chars of raw HTML instead, the page structure may have changed:');
    Logger.log(html.slice(0, 3000));
  }

  return events;
}

/**
 * Single entry point — run this from a time-driven trigger.
 * Runs booking creation first (so same-run invoices can find a freshly
 * created apartmentery bookingId), then invoice+receipt creation.
 * Logs a summary; only sends a LINE alert if something needs human
 * attention (errors, or session expiry — which ApartmenteryClient already
 * alerts on directly, so it isn't duplicated here).
 */
function runApartmenteryAutomation() {
  const bookingResult = autoCreateApartmenteryBookings();
  // Session already expired during phase 1 — no point attempting phase 2.
  const invoiceResult = bookingResult.sessionExpired
    ? { created: 0, skipped: 0, sessionExpired: true, errors: [] }
    : autoCreateApartmenteryInvoicesAndReceipts();

  Logger.log('Apartmentery automation run: ' + JSON.stringify({ bookingResult, invoiceResult }));

  const allErrors = [...bookingResult.errors, ...invoiceResult.errors];
  if (allErrors.length > 0) {
    try {
      const props = PropertiesService.getScriptProperties();
      const botUrl = props.getProperty('BOT_URL') || 'https://hotel-line-bot.onrender.com';
      const adminToken = props.getProperty('ADMIN_TOKEN') || 'apt2025@secret';
      const summary = allErrors
        .map(e => `- ${e.guest || ''} ${e.room || e.resId || e.invoiceKey}: ${e.error}`)
        .join('\n');
      UrlFetchApp.fetch(botUrl + '/api/send-admin-alert', {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ note: `⚠️ Apartmentery automation มีรายการที่ error ${allErrors.length} รายการ:\n${summary}` }),
        headers: { 'x-admin-token': adminToken },
        muteHttpExceptions: true
      });
    } catch (e) {
      Logger.log('Failed to send error summary LINE alert: ' + e.message);
    }
  }

  return { bookingResult, invoiceResult };
}
