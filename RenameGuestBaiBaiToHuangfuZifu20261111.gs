/**
 * RenameGuestBaiBaiToHuangfuZifu20261111.gs
 * -----------------------------------------------------------------------
 * Booking.com reservation #5062908920 (room "Loft", 10→20 Nov 2026) had its
 * guest name modified on 2026-11-11 07:44 (Little Hotelier "Reservation
 * Modification" email): "Bai Bai"  →  "Huangfu Zifu".
 *
 * The Gmail pipeline only creates rows from New Reservation emails, so the
 * rename never reached our data. This one-off updates:
 *   1. Sheet1 column "ชื่อแขก"  (ResId is intentionally left unchanged —
 *      it is derived from the original name and other tabs key off it)
 *   2. The Apartmentery booking's customerName ("Huangfu Zifu / <Channel>")
 *
 * Safe to re-run (idempotent). Aborts without writing if it can't find
 * exactly ONE matching Sheet1 row.
 *
 * HOW TO RUN: Apps Script editor ▶ renameGuestBaiBaiToHuangfuZifu20261111 ▶ Run
 * ▶ read the log.
 */

function renameGuestBaiBaiToHuangfuZifu20261111() {
  const OLD_NAME = 'bai bai';
  const NEW_NAME = 'Huangfu Zifu';
  const CHECKIN  = '2026-11-10';
  const CHECKOUT = '2026-11-20';

  const ss  = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  const data = src.getDataRange().getValues();
  const idx = indexMap_(data[0], ['ResId', 'เลขห้อง', 'ชื่อแขก', 'เช็คอิน', 'เช็คเอาท์', 'Channel', 'Apartmentery Booking ID']);
  if (idx.ResId < 0 || idx['ชื่อแขก'] < 0 || idx['เช็คอิน'] < 0) {
    Logger.log('ABORT — required Sheet1 columns not found');
    return;
  }

  const norm = s => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
  const matches = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const guest = norm(r[idx['ชื่อแขก']]);
    const ci = formatCellDate_(r[idx['เช็คอิน']]);
    const co = formatCellDate_(r[idx['เช็คเอาท์']]);
    const ch = String(r[idx.Channel] || '');
    if ((guest === OLD_NAME || guest === norm(NEW_NAME)) && ci === CHECKIN && co === CHECKOUT && /booking/i.test(ch)) {
      matches.push(i);
    }
  }
  if (matches.length !== 1) {
    Logger.log(`ABORT — expected exactly 1 matching row (Bai Bai / ${CHECKIN}→${CHECKOUT} / Booking.com), found ${matches.length}. Nothing written.`);
    return;
  }

  const i = matches[0];
  const r = data[i];
  const resId   = String(r[idx.ResId] || '').trim();
  const room    = String(r[idx['เลขห้อง']] || '').trim();
  const channel = String(r[idx.Channel] || '').trim();
  const oldGuest = String(r[idx['ชื่อแขก']] || '').trim();
  let aptId = idx['Apartmentery Booking ID'] >= 0 ? String(r[idx['Apartmentery Booking ID']] || '').trim() : '';

  // 1) Sheet1
  if (oldGuest !== NEW_NAME) {
    src.getRange(i + 1, idx['ชื่อแขก'] + 1).setValue(NEW_NAME);
    Logger.log(`Sheet1 row ${i + 1} (${resId}, room ${room}): "${oldGuest}" → "${NEW_NAME}"`);
  } else {
    Logger.log(`Sheet1 row ${i + 1} (${resId}) already reads "${NEW_NAME}"`);
  }
  triggerStyleSheet1_();

  // 2) Apartmentery
  try {
    if (!aptId) {
      aptId = findApartmenteryBookingIdForRoomByGuest_(room, oldGuest, CHECKIN) || '';
      if (aptId) setApartmenteryBookingId_(resId, aptId);
    }
    if (!aptId) {
      Logger.log('Apartmentery: no booking id found yet — nothing to rename there (the hourly automation will create it under the new name).');
    } else {
      const unit = getApartmenteryUnitForRoom(room);
      if (!unit) throw new Error('no Apartmentery unit for room ' + room);
      const newCustomer = channel ? `${NEW_NAME} / ${channel}` : NEW_NAME;
      const res = updateApartmenteryBookingCustomerName(unit.branchId, unit.unitId, aptId, newCustomer);
      Logger.log(`Apartmentery booking ${aptId}: "${res.oldCustomerName}" → "${res.newCustomerName}"`);
    }
  } catch (e) {
    Logger.log('Apartmentery rename FAILED — change the customer name by hand: ' + e);
  }

  invalidateDashboardCache_();
  invalidateRoomStatusCache_();
  Logger.log('DONE');
}
