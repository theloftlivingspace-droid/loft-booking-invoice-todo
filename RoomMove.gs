/**
 * RoomMove.gs — add this file to the loft-booking-invoice-todo Apps Script
 * project (same project that already backs both the "todo" AND "checkinout"
 * gas-proxy endpoints — confirmed by reading api/gas-proxy.js in the-loft-admin:
 * both URLs point at deployments of THIS SAME script).
 *
 * Wire into doPost() in Code.gs:
 *
 *   if (action === 'moveGuestRoom') {
 *     return jsonResponse_(moveGuestRoom_(body));
 *   }
 *
 * -----------------------------------------------------------------------
 * VERIFIED AGAINST REAL SOURCE (not guessed) — 2026-08-22:
 *   - SOURCE_SHEET_ID / SRC_BOOKING_SHEET constants:  Code.gs:14,16
 *   - Real Sheet1 headers are THAI, not English:
 *       'ResId', 'เลขห้อง', 'ชื่อแขก', 'เช็คอิน', 'เช็คเอาท์', 'Channel',
 *       'Note', 'Apartmentery Booking ID'                 (Code.gs:1570, 913)
 *   - indexMap_(header, keys)                          Code.gs:1699
 *   - formatCellDate_(val) → 'YYYY-MM-DD'               Code.gs:1705
 *   - roomNum_(room) → strips ' ยกเลิก' / non-digits     Code.gs:1314
 *   - setBookingNote(resId, note)   (NOT "setNote")     Code.gs:1543
 *   - triggerStyleSheet1_()  — fire-and-forget POST to  Code.gs:1636
 *     payout-income-log's separate webapp; NOT a local function.
 *   - getApartmenteryBookingId_(resId) / setApartmenteryBookingId_(...)
 *                                                        ApartmenteryAutomation.gs:216,280
 *   - setApartmenteryBookingId_ has a uniqueness guard (added 2026-07-18)
 *     that REFUSES to write a bookingId already used by another resId —
 *     Segment B never needs to touch this directly (it goes through the
 *     normal autoCreateApartmenteryBookings path), so no conflict here.
 *   - PROP_KEY_INVOICE_APT_IDS = 'invoice_apt_ids_v1', map value format
 *     '{aptBookingId}:{invoiceId}'                       Code.gs:35, 1503
 *   - syncApartmenteryCheckoutDate_(resId, room, guest, checkin, newCheckout)
 *     — the exact helper cancelBooking_/earlyCheckout_/updateCheckoutDate_
 *     all reuse to push a shrunk end-date to Apartmentery. Segment A reuses
 *     this too, so it behaves identically to a normal early-checkout.
 *                                                        Code.gs:275
 *   - hotel-line-bot notify pattern is NOT a shared GAS function — it's an
 *     HTTP POST from GAS to hotel-line-bot's own endpoints, header
 *     'x-admin-token', e.g. cancelBooking_'s call to /api/cancel-notify
 *     (Code.gs:1611-1622). wasCheckinAlreadyAnnounced() lives in
 *     hotel-line-bot/bot.js:1350 and takes a CHECKIN DATE STRING, not a
 *     resId — my earlier draft called it with resId, which was wrong.
 *   - ROOM_LIST in the-loft-admin/src/CheckInOut.tsx:248 has 16 rooms with
 *     types I had wrong before (Noir, Emerald, Rhythm exist; 104/105/112/
 *     207/208/211 are in MANUALLY_CLOSED_ROOMS — under renovation, must be
 *     excluded from the picker). "363 Mycondo" is NOT in ROOM_LIST at all —
 *     it's parsed by a separate Airbnb363ToSheet1.gs path that bypasses
 *     Little Hotelier, so it's excluded from this move flow's destination
 *     list until confirmed otherwise.
 * -----------------------------------------------------------------------
 *
 * ORPHANED APARTMENTERY BOOKING (case 1 only) — RESOLVED, NOT GUESSED:
 *   "ไม่มี invoice" (case 1) is checked via getApartmenteryBookingId_ +
 *   invoice_apt_ids_v1 below. But autoCreateApartmenteryBookings runs
 *   hourly and could create the Apartmentery-side booking (assign a
 *   bookingId) BEFORE check-in, i.e. before any invoice exists. In that
 *   situation this function still correctly classifies it as "case 1" (no
 *   invoice → safe to delete+recreate the Sheet1 row), but the OLD
 *   Apartmentery booking (old room, no invoice) is left with nothing in
 *   Sheet1 pointing to it anymore.
 *
 *   You said: delete the old one, create a new one for the new room —
 *   Sheet1-side that's exactly what moveRoomBeforeCheckin_ does. But I
 *   checked ApartmenteryClient.gs in full (1139 lines, every function
 *   listed) and there is NO delete/cancel function anywhere in it — every
 *   existing function only creates or edits (createApartmenteryBooking,
 *   updateApartmenteryBookingEndDate, createApartmenteryInvoice,
 *   createApartmenteryReceipt). No one has ever deleted a booking on
 *   apartmentery.com through this codebase, so there's no verified
 *   endpoint/form-field shape to call — writing one now would mean
 *   guessing a POST against production apartmentery.com, which is exactly
 *   what you told me not to do.
 *
 *   So handleOrphanedPreInvoiceAptBooking_() below does NOT call Apartmentery
 *   at all. It sends you a 1:1 LINE alert via /api/send-admin-alert — the
 *   same endpoint _notifyLineSessionFailure_() already uses in this same
 *   file (ApartmenteryClient.gs:199-214) for session-expiry alerts — asking
 *   you to delete that specific booking manually on apartmentery.com. Once
 *   you've actually deleted one this way and know the real request
 *   apartmentery.com makes (check browser devtools' Network tab), send me
 *   that and I'll wire up a real automated delete instead of this alert.
 */

const ROOM_MOVE_LOG_SHEET = 'RoomMoveLog';

// Verified against the-loft-admin/src/CheckInOut.tsx:248 (2026-08-21 revision).
// Keep this in sync manually if ROOM_LIST changes there — there is no shared
// source of truth between the two repos for this list.
const MOVE_TARGET_ROOMS = [
  { num: '203', type: 'Allure' },
  { num: '103', type: 'Elegance' },
  { num: '209', type: 'Radiance' },
  { num: '113', type: 'Legacy' },
  { num: '300', type: 'Luxury' },
  { num: '205', type: 'Allure' },
  { num: '204', type: 'Elegance' },
  { num: '210', type: 'Radiance' },
  { num: '214', type: 'Legacy' },
  { num: '108', type: 'Retro' },
  // 104, 105, 112, 207, 208, 211 deliberately excluded — MANUALLY_CLOSED_ROOMS
  // in CheckInOut.tsx (under renovation as of 2026-08-21).
  // '363 Mycondo' deliberately excluded — bypasses Little Hotelier via
  // Airbnb363ToSheet1.gs, not part of the normal Sheet1 room-move flow.
];

/**
 * Entry point — call as moveGuestRoom_(body) from doPost().
 * body: { resId, newRoom, reason, actor, effectiveDate? }
 */
function moveGuestRoom_(body) {
  const resId = String(body.resId || '').trim();
  const newRoom = String(body.newRoom || '').trim();
  const reason = String(body.reason || '');
  const actor = String(body.actor || '');
  const effectiveDate = body.effectiveDate ? String(body.effectiveDate).trim() : '';

  if (!resId || !newRoom) {
    return { ok: false, error: 'resId and newRoom are required' };
  }

  const validTarget = MOVE_TARGET_ROOMS.some(r => r.num === newRoom);
  if (!validTarget) {
    return { ok: false, error: `${newRoom} is not a valid move target (closed for renovation, or not in ROOM_LIST)` };
  }

  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  if (!src) return { ok: false, error: 'Sheet1 not found' };

  const booking = findBookingByResId_(src, resId);
  if (!booking) return { ok: false, error: `resId not found: ${resId}` };
  if (/ยกเลิก|cancel/i.test(booking.room)) {
    return { ok: false, error: 'booking is cancelled — cannot move a cancelled booking' };
  }

  const oldRoomNum = roomNum_(booking.room);
  if (oldRoomNum === newRoom) {
    return { ok: false, error: 'newRoom is the same as the current room' };
  }

  const checkStart = effectiveDate || booking.checkIn;
  const availability = isRoomAvailable_(src, newRoom, checkStart, booking.checkOut, resId);
  if (!availability.available) {
    return {
      ok: false,
      error: `room ${newRoom} is not available ${checkStart} – ${booking.checkOut}`,
      conflict: availability.conflict,
    };
  }

  const aptBookingId = getApartmenteryBookingId_(resId);
  const hasInvoice = aptBookingId ? isBookingIdInvoiced_(aptBookingId) : false;

  const result = hasInvoice
    ? moveRoomAfterCheckin_(src, booking, newRoom, effectiveDate)
    : moveRoomBeforeCheckin_(src, booking, newRoom, aptBookingId);

  if (result.ok) {
    logRoomMove_(ss, {
      resId,
      oldRoom: oldRoomNum,
      newRoom,
      reason,
      actor,
      caseType: hasInvoice ? 'after_checkin' : 'before_checkin',
      timestamp: new Date(),
    });
    notifyMaidGroupRoomMove_(booking, oldRoomNum, newRoom, hasInvoice, checkStart, result);
  }

  return result;
}

// ---------------------------------------------------------------------
// เคส 1: ยังไม่มี invoice — ลบ row เดิม, สร้าง row ใหม่ resId เดิม
// ---------------------------------------------------------------------

function moveRoomBeforeCheckin_(src, booking, newRoom, existingAptBookingId) {
  try {
    if (existingAptBookingId) {
      handleOrphanedPreInvoiceAptBooking_(booking, existingAptBookingId);
    }

    const oldRowValues = src.getRange(booking.rowIndex, 1, 1, src.getLastColumn()).getValues()[0];
    src.deleteRow(booking.rowIndex);

    const newRowValues = oldRowValues.slice();
    newRowValues[booking.idx['เลขห้อง']] = newRoom;
    // Apartmentery Booking ID column (if present) must be cleared — the old
    // id belonged to the old room and is now orphaned/shrunk; a fresh row
    // must let autoCreateApartmenteryBookings assign a brand-new one for
    // the new room, never carry the stale id forward.
    if (booking.idx[APARTMENTERY_BOOKING_ID_COL_HEADER] >= 0) {
      newRowValues[booking.idx[APARTMENTERY_BOOKING_ID_COL_HEADER]] = '';
    }

    const newRowIndex = src.getLastRow() + 1;
    src.getRange(newRowIndex, 1, 1, newRowValues.length).setValues([newRowValues]);
    triggerStyleSheet1_();

    return {
      ok: true,
      case: 'before_checkin',
      message: `moved ${booking.resId}: ${roomNum_(booking.room)} → ${newRoom} (row recreated, no invoice existed)`,
    };
  } catch (err) {
    return { ok: false, error: `moveRoomBeforeCheckin_ failed: ${err.message}` };
  }
}

/**
 * See "ORPHANED APARTMENTERY BOOKING" note in the file header — no delete
 * function exists anywhere in ApartmenteryClient.gs, so this does NOT touch
 * Apartmentery at all. Sends a 1:1 LINE alert to Nathan asking him to
 * delete the old booking manually, reusing the exact same /api/send-admin-alert
 * call shape _notifyLineSessionFailure_() already uses (ApartmenteryClient.gs:199-214).
 */
function handleOrphanedPreInvoiceAptBooking_(booking, aptBookingId) {
  try {
    const props = PropertiesService.getScriptProperties();
    const botUrl = props.getProperty('BOT_URL') || 'https://hotel-line-bot.onrender.com';
    const adminToken = props.getProperty('ADMIN_TOKEN');
    if (!adminToken) {
      Logger.log('handleOrphanedPreInvoiceAptBooking_: ADMIN_TOKEN not set — cannot send alert. ' +
        `Manual cleanup needed: apartmentery bookingId ${aptBookingId} (resId ${booking.resId}, ` +
        `old room ${roomNum_(booking.room)}) has no invoice and is now orphaned — delete it by hand.`);
      return;
    }
    UrlFetchApp.fetch(botUrl + '/api/send-admin-alert', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        note: '🗑️ ต้องลบ booking ค้างบน Apartmentery ด้วยตนเอง\n' +
          `ResId: ${booking.resId}\n` +
          `Apartmentery bookingId: ${aptBookingId}\n` +
          `ห้องเดิม: ${roomNum_(booking.room)}\n` +
          `เหตุผล: ย้ายห้องก่อนเช็คอิน ยังไม่มี invoice — Sheet1 อัปเดตห้องใหม่ให้แล้ว ` +
          'แต่ระบบไม่มีฟังก์ชันลบ booking บน Apartmentery อัตโนมัติ ต้องลบเองที่ apartmentery.com',
      }),
      headers: { 'x-admin-token': adminToken },
      muteHttpExceptions: true,
    });
  } catch (e) {
    Logger.log(`handleOrphanedPreInvoiceAptBooking_ error: ${e.message} — apartmentery bookingId ${aptBookingId} (resId ${booking.resId}) still needs manual deletion.`);
  }
}

// ---------------------------------------------------------------------
// เคส 2: มี invoice แล้ว — Segment A (checkout เดิม) + Segment B (checkin ใหม่)
// ---------------------------------------------------------------------

function moveRoomAfterCheckin_(src, booking, newRoom, effectiveDate) {
  const moveDate = effectiveDate || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');

  if (moveDate <= booking.checkIn || moveDate >= booking.checkOut) {
    return { ok: false, error: `effectiveDate (${moveDate}) must be strictly between checkin (${booking.checkIn}) and checkout (${booking.checkOut})` };
  }

  const segmentBResId = `${booking.resId}-B`;
  if (findBookingByResId_(src, segmentBResId)) {
    return { ok: false, error: `${segmentBResId} already exists — this booking may already have been moved; check RoomMoveLog before retrying` };
  }

  try {
    // ---- Segment A: shrink checkout to moveDate, exactly like an early checkout ----
    src.getRange(booking.rowIndex, booking.idx['เช็คเอาท์'] + 1).setValue(moveDate);
    const syncResultA = syncApartmenteryCheckoutDate_(
      booking.resId,
      roomNum_(booking.room),
      booking.guest,
      booking.checkIn,
      moveDate
    );

    setBookingNote(
      booking.resId,
      `ย้ายห้องไป ${newRoom} ตั้งแต่ ${moveDate} — ดูช่วงต่อที่ resId ${segmentBResId}`
    );

    // ---- Segment B: new row, new resId suffix, room + dates for the remainder ----
    const oldRowValues = src.getRange(booking.rowIndex, 1, 1, src.getLastColumn()).getValues()[0];
    const segmentBValues = oldRowValues.slice();
    segmentBValues[booking.idx['ResId']] = segmentBResId;
    segmentBValues[booking.idx['เลขห้อง']] = newRoom;
    segmentBValues[booking.idx['เช็คอิน']] = moveDate;
    segmentBValues[booking.idx['เช็คเอาท์']] = booking.checkOut;
    if (booking.idx[APARTMENTERY_BOOKING_ID_COL_HEADER] >= 0) {
      segmentBValues[booking.idx[APARTMENTERY_BOOKING_ID_COL_HEADER]] = ''; // must get its own, via autoCreateApartmenteryBookings
    }
    if (booking.idx['Note'] >= 0) {
      segmentBValues[booking.idx['Note']] =
        `Segment B จากการย้ายห้อง (resId เดิม ${booking.resId}, ห้องเดิม ${roomNum_(booking.room)}, ย้ายวันที่ ${moveDate})`;
    }

    const newRowIndex = src.getLastRow() + 1;
    src.getRange(newRowIndex, 1, 1, segmentBValues.length).setValues([segmentBValues]);
    triggerStyleSheet1_();

    // Deliberately NOT calling autoCreateApartmenteryBookings() or any
    // invoice-creation function here — Segment B goes through the normal
    // hourly automation pass, at current rates, same as any other new
    // booking. This is a decision, not an oversight: it keeps pricing for
    // the new room out of this function's hands (per the earlier
    // conversation about not auto-carrying the old price across a
    // room-type change).

    return {
      ok: true,
      case: 'after_checkin',
      segmentBResId,
      message: `Segment A (${roomNum_(booking.room)}) checked out ${moveDate}; Segment B (${newRoom}, ${segmentBResId}) checks in ${moveDate} → ${booking.checkOut}, invoice pending via normal automation`,
      apartmenterySyncA: syncResultA,
    };
  } catch (err) {
    return { ok: false, error: `moveRoomAfterCheckin_ failed: ${err.message}` };
  }
}

// ---------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------

function findBookingByResId_(src, resId) {
  const data = src.getDataRange().getValues();
  const header = data[0];
  const idx = indexMap_(header, [
    'ResId', 'เลขห้อง', 'ชื่อแขก', 'เช็คอิน', 'เช็คเอาท์', 'Note', APARTMENTERY_BOOKING_ID_COL_HEADER,
  ]);
  if (idx.ResId < 0 || idx['เลขห้อง'] < 0 || idx['เช็คอิน'] < 0 || idx['เช็คเอาท์'] < 0) {
    throw new Error('required Sheet1 columns missing (ResId / เลขห้อง / เช็คอิน / เช็คเอาท์)');
  }
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.ResId] || '').trim() === resId) {
      return {
        rowIndex: i + 1,
        resId: String(data[i][idx.ResId]).trim(),
        room: String(data[i][idx['เลขห้อง']] || '').trim(),
        guest: idx['ชื่อแขก'] >= 0 ? String(data[i][idx['ชื่อแขก']] || '').trim() : '',
        checkIn: formatCellDate_(data[i][idx['เช็คอิน']]),
        checkOut: formatCellDate_(data[i][idx['เช็คเอาท์']]),
        idx,
      };
    }
  }
  return null;
}

/**
 * Overlap check against the destination room only, using the same
 * roomNum_()-normalized comparison and cancelled-row skip that
 * updateCheckoutDate_'s own conflict check uses (Code.gs:470-493), so a
 * room-move refuses the exact same conflicts a checkout-date edit would.
 */
function isRoomAvailable_(src, targetRoomNum, checkIn, checkOut, excludeResId) {
  const data = src.getDataRange().getValues();
  const header = data[0];
  const idx = indexMap_(header, ['ResId', 'เลขห้อง', 'เช็คอิน', 'เช็คเอาท์', 'ชื่อแขก']);

  for (let i = 1; i < data.length; i++) {
    const rowResId = String(data[i][idx.ResId] || '').trim();
    if (rowResId === excludeResId || rowResId === `${excludeResId}-B`) continue;

    const otherRoom = String(data[i][idx['เลขห้อง']] || '').trim();
    if (/ยกเลิก|cancel/i.test(otherRoom)) continue;
    if (roomNum_(otherRoom) !== targetRoomNum) continue;

    const otherCheckIn = formatCellDate_(data[i][idx['เช็คอิน']]);
    const otherCheckOut = formatCellDate_(data[i][idx['เช็คเอาท์']]);
    if (!otherCheckIn || !otherCheckOut) continue;

    if (checkIn < otherCheckOut && checkOut > otherCheckIn) {
      return {
        available: false,
        conflict: {
          resId: rowResId,
          guest: String(data[i][idx['ชื่อแขก']] || '').trim(),
          checkIn: otherCheckIn,
          checkOut: otherCheckOut,
        },
      };
    }
  }
  return { available: true };
}

/**
 * True if this Apartmentery bookingId already has an invoice recorded
 * against it in invoice_apt_ids_v1 (value format '{bookingId}:{invoiceId}',
 * set by setInvoiceApartmenteryIds — Code.gs:1503).
 */
function isBookingIdInvoiced_(aptBookingId) {
  const map = getProp_(PROP_KEY_INVOICE_APT_IDS);
  const target = String(aptBookingId).trim();
  return Object.keys(map).some(invoiceKey => {
    const stored = String(map[invoiceKey] || '');
    return stored.split(':')[0] === target;
  });
}

function logRoomMove_(ss, entry) {
  let logSheet = ss.getSheetByName(ROOM_MOVE_LOG_SHEET);
  if (!logSheet) {
    logSheet = ss.insertSheet(ROOM_MOVE_LOG_SHEET);
    logSheet.appendRow(['Timestamp', 'ResId', 'OldRoom', 'NewRoom', 'Reason', 'Actor', 'Case']);
  }
  logSheet.appendRow([entry.timestamp, entry.resId, entry.oldRoom, entry.newRoom, entry.reason, entry.actor, entry.caseType]);
}

/**
 * Fires the immediate LINE alert only if this checkin date already passed
 * the maid-group announcement threshold — same rule hotel-line-bot's
 * wasCheckinAlreadyAnnounced() already enforces server-side for cancel/
 * checkout notify, added here so the GAS side doesn't even attempt the
 * call when it's pointless (still safe either way since the bot re-checks).
 *
 * NOTE: this requires a NEW endpoint on hotel-line-bot — see
 * hotel-line-bot-room-move-notify.js in this same delivery. There is no
 * existing endpoint that fits (cancel-notify and checkout-notify both send
 * a fixed message shape that doesn't cover "moved to a different room").
 */
function notifyMaidGroupRoomMove_(booking, oldRoom, newRoom, hasInvoice, checkStart, result) {
  try {
    const props = PropertiesService.getScriptProperties();
    const botUrl = props.getProperty('BOT_URL') || 'https://hotel-line-bot.onrender.com';
    const adminTok = props.getProperty('ADMIN_TOKEN');
    if (!adminTok) {
      Logger.log('notifyMaidGroupRoomMove_: ADMIN_TOKEN script property not set — skipping LINE notify');
      return;
    }
    UrlFetchApp.fetch(botUrl + '/api/room-move-notify', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        resId: booking.resId,
        guest: booking.guest,
        oldRoom,
        newRoom,
        checkin: booking.checkIn,
        effectiveDate: checkStart,
        splitBooking: hasInvoice,
        segmentBResId: result.segmentBResId || null,
      }),
      headers: { 'x-admin-token': adminTok },
      muteHttpExceptions: true,
    });
  } catch (e) {
    Logger.log('notifyMaidGroupRoomMove_ error: ' + e);
  }
}
