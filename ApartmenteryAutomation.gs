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

/** Writes the apartmentery bookingId for a given Sheet1 ResId. */
function setApartmenteryBookingId_(resId, apartmenteryBookingId) {
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
function autoCreateApartmenteryBookings() {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  const items = getBookingToAdd_(ss, todayStr);

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
  items.forEach(x => {
    if (/ยกเลิก|cancel/i.test(x.room)) return; // cancelled stays don't occupy the room
    if (!x.checkout) return;
    const rn = roomNum_(x.room);
    if (!outgoingByRoom[rn]) outgoingByRoom[rn] = {};
    outgoingByRoom[rn][x.checkout] = x.resId;
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

    Logger.log(`attempting booking for ${b.resId} room ${b.room} guest ${b.guest}`);

    try {
      // Apartmentery customerName follows Nathan's original manual naming
      // convention "Guest Name / Channel" (e.g. "Pranee Antov / Expedia")
      // — the automation was previously sending just the bare guest name.
      const guestNameWithChannel = b.channel ? `${b.guest} / ${b.channel}` : b.guest;

      const created = createApartmenteryBookingForRoom(b.room, {
        startDate: b.checkin,
        endDate: b.checkout || '',
        guestName: guestNameWithChannel,
        note: `${b.channel} ${b.resId}`.trim()
      });

      if (created && created.skipped) {
        Logger.log(`skip ${b.resId} (${b.room}): ${created.reason}`);
        result.skipped++;
        continue;
      }

      setApartmenteryBookingId_(b.resId, created.bookingId);
      setBookingDone(b.resId, true);
      Logger.log(`created ${b.resId} (${b.room}) -> apartmentery bookingId ${created.bookingId}`);
      result.created++;

    } catch (err) {
      if (isApartmenterySessionExpiredError(err)) {
        Logger.log(`SESSION EXPIRED while creating booking for ${b.resId} (${b.room}): ${err.message}`);
        result.sessionExpired = true;
        break; // stop the whole run — see file header
      }
      Logger.log(`ERROR creating booking for ${b.resId} (${b.room}): ${err.message}`);
      result.errors.push({ resId: b.resId, guest: b.guest, room: b.room, error: err.message });
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

  // Build matchKey -> resId index from booking items (only ones that
  // actually have an apartmentery bookingId already — no point matching
  // to a booking apartmentery doesn't know about yet).
  const keyToResId = {};
  bookingItems.forEach(b => {
    const aptId = getApartmenteryBookingId_(b.resId);
    if (!aptId) return;
    (b.matchKeys || []).forEach(k => { if (!keyToResId[k]) keyToResId[k] = b.resId; });
  });

  const result = { created: 0, skipped: 0, sessionExpired: false, errors: [] };

  for (const inv of invoiceItems) {
    if (inv.done) continue;
    if (String(inv.room).indexOf('ไม่ทราบห้อง') >= 0) { result.skipped++; continue; } // unresolved room, needs manual review

    let resId = null;
    for (const k of (inv.matchKeys || [])) {
      if (keyToResId[k]) { resId = keyToResId[k]; break; }
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
      result.errors.push({ invoiceKey: inv.invoiceKey, guest: inv.guest, room: inv.room, error: err.message });
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
  items.forEach(x => {
    if (/ยกเลิก|cancel/i.test(x.room)) return;
    if (!x.checkout) return;
    const rn = roomNum_(x.room);
    if (!outgoingByRoom[rn]) outgoingByRoom[rn] = {};
    outgoingByRoom[rn][x.checkout] = x.resId;
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

    Logger.log(`[backfill] attempting booking for ${b.resId} room ${b.room} guest ${b.guest}`);

    try {
      const guestNameWithChannel = b.channel ? `${b.guest} / ${b.channel}` : b.guest;
      const created = createApartmenteryBookingForRoom(b.room, {
        startDate: b.checkin,
        endDate: b.checkout || '',
        guestName: guestNameWithChannel,
        note: `${b.channel} ${b.resId}`.trim()
      });

      if (created && created.skipped) {
        Logger.log(`skip ${b.resId} (${b.room}): ${created.reason}`);
        result.skipped++;
        continue;
      }

      setApartmenteryBookingId_(b.resId, created.bookingId);
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
        const guestNameWithChannel = b.channel ? `${b.guest} / ${b.channel}` : b.guest;
        let recoveredId = null;
        try {
          recoveredId = findApartmenteryBookingIdForRoomByGuest_(b.room, guestNameWithChannel, b.checkin);
        } catch (lookupErr) {
          if (isApartmenterySessionExpiredError(lookupErr)) {
            Logger.log(`SESSION EXPIRED while recovering bookingId for ${b.resId} (${b.room}): ${lookupErr.message}`);
            result.sessionExpired = true;
            break;
          }
          Logger.log(`recovery lookup failed for ${b.resId} (${b.room}): ${lookupErr.message}`);
        }

        if (recoveredId) {
          setApartmenteryBookingId_(b.resId, recoveredId);
          Logger.log(`[backfill] recovered ${b.resId} (${b.room}) -> existing apartmentery bookingId ${recoveredId} (was already in apartmentery, not newly created)`);
          result.created++; // counts toward "now has a bookingId", same outcome for the sheet
          result.recovered = (result.recovered || 0) + 1;
          continue;
        }

        Logger.log(`skip ${b.resId} (${b.room}): collision reported but no exact guest+date match found in unit's calendar — inspect manually. (${err.apartmenteryError})`);
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
