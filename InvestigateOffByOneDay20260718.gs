/**
 * InvestigateOffByOneDay20260718.gs
 * -----------------------------------------------------------------------
 * Tests Nathan's hypothesis for the ±1-day offset found and fixed in
 * applyConfirmedSafeCandidates20260718: that it's the same mechanism as
 * the same-day-turnover buffer (autoCreateApartmenteryBookings /
 * backfillMissingApartmenteryBookings shrink the OUTGOING booking's
 * apartmentery endDate by 1 day so a same-day-checkin new booking doesn't
 * 500) — specifically, that for these 5 the NEW booking's own startDate
 * got shifted by 1 day instead of (or in addition to) shrinking the
 * previous booking.
 *
 * Important: reading the actual code (ApartmenteryAutomation.gs
 * autoCreateApartmenteryBookings, ApartmenteryClient.gs
 * createApartmenteryBooking) shows the new booking's startDate is always
 * sent as-is (opts.startDate = b.checkin, never shifted) — only the
 * OUTGOING booking's endDate is ever touched. So if this hypothesis is
 * right, the shift would have to be happening on apartmentery's SERVER
 * SIDE silently (not in our code) when it receives a startDate colliding
 * with another booking. This script checks for that pattern directly
 * instead of assuming: for each of the 5 fixed resIds, look for a
 * same-room booking whose Sheet1 checkout matches this resId's Sheet1
 * checkin (the same signal autoCreateApartmenteryBookings uses to detect
 * a turnover) and report what's found.
 *
 * READ-ONLY.
 *
 * HOW TO RUN: Apps Script editor ▶ investigateOffByOneDay20260718 ▶ Run ▶
 * read log.
 */

const OFFBYONE_20260718_ = [
  { resId: 'ABB-johnzambra-20260428', guest: 'John Zambrana', room: '108', sheetCheckin: '2026-04-28', apartmenteryStart: '2026-04-27' },
  { resId: 'ABB-premmehta-20260501', guest: 'Prem Mehta', room: '108', sheetCheckin: '2026-05-01', apartmenteryStart: '2026-05-02' },
  { resId: 'ABB-kgotlellom-20260528', guest: 'Kgotlello Masemola', room: '203', sheetCheckin: '2026-05-27', apartmenteryStart: '2026-05-28' },
  { resId: 'ABB-errolcox-20260608', guest: 'Errol Cox', room: '210', sheetCheckin: '2026-06-08', apartmenteryStart: '2026-06-09' },
  { resId: 'ABB-saeidmickm-20260610', guest: 'Saeid Mick Momtahan', room: '108', sheetCheckin: '2026-06-10', apartmenteryStart: '2026-06-09' }
];

function investigateOffByOneDay20260718() {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName('Sheet1');
  const data = src.getDataRange().getValues();
  const header = data[0];
  const idx = indexMap_(header, ['ResId', 'เลขห้อง', 'ชื่อแขก', 'เช็คอิน', 'เช็คเอาท์', APARTMENTERY_BOOKING_ID_COL_HEADER]);

  // Build every row keyed by resId, plus a per-room list of all
  // checkin/checkout pairs so we can look for a turnover partner.
  const allRows = [];
  for (let i = 1; i < data.length; i++) {
    const resId = String(data[i][idx.ResId] || '').trim();
    if (!resId) continue;
    allRows.push({
      resId,
      room: String(data[i][idx['เลขห้อง']] || '').trim(),
      guest: String(data[i][idx['ชื่อแขก']] || '').trim(),
      checkin: formatCellDate_(data[i][idx['เช็คอิน']]),
      checkout: formatCellDate_(data[i][idx['เช็คเอาท์']]),
      aptId: String(data[i][idx[APARTMENTERY_BOOKING_ID_COL_HEADER]] || '').trim()
    });
  }

  OFFBYONE_20260718_.forEach(row => {
    Logger.log(`--- ${row.resId} (${row.guest}, room ${row.room}) — Sheet1 checkin=${row.sheetCheckin}, apartmentery actual start=${row.apartmenteryStart} ---`);

    const currentRow = allRows.find(r => r.resId === row.resId);
    if (currentRow) {
      Logger.log(`  Sheet1 NOW: checkin=${currentRow.checkin} checkout=${currentRow.checkout} room=${currentRow.room} — ` +
        `${currentRow.checkin === row.sheetCheckin ? 'unchanged since audit' : 'CHANGED since the audit ran (was ' + row.sheetCheckin + ')'}`);
    } else {
      Logger.log(`  Sheet1 row for this resId not found (deleted or resId changed?).`);
    }

    // Turnover partner: same room, checkout === this booking's checkin.
    const rn = roomNum_(row.room);
    const turnoverPartner = allRows.find(r => r.resId !== row.resId && roomNum_(r.room) === rn && r.checkout === row.sheetCheckin);
    if (turnoverPartner) {
      Logger.log(`  TURNOVER PARTNER FOUND: ${turnoverPartner.resId} (${turnoverPartner.guest}) checks out ` +
        `${turnoverPartner.checkout} same day this booking checks in — this room DID have a same-day turnover. ` +
        `Outgoing booking's apartmentery id: ${turnoverPartner.aptId || '(none)'}.`);
      if (turnoverPartner.aptId) {
        try {
          const unit = getApartmenteryUnitForRoom(row.room);
          const state = _getApartmenteryBookingEditFormState_(unit.branchId, unit.unitId, turnoverPartner.aptId);
          Logger.log(`  Outgoing booking ${turnoverPartner.aptId}'s CURRENT apartmentery endDate: ${state.endDate} ` +
            `(expected shrunk value would be ${_dateMinusOneDay_(row.sheetCheckin)})`);
        } catch (e) {
          Logger.log(`  Could not read outgoing booking's current apartmentery state: ${e.message}`);
        }
      }
    } else {
      Logger.log(`  NO turnover partner in Sheet1 — no other booking in room ${row.room} checks out on ${row.sheetCheckin}. ` +
        `The same-day-turnover shrink code path would never have run for this booking, so that's likely NOT the cause here.`);
    }
  });

  Logger.log(`=== investigateOffByOneDay20260718 done — read the per-row detail above ===`);
}
