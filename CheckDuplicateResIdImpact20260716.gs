/**
 * CheckDuplicateResIdImpact20260716.gs
 * -----------------------------------------------------------------------
 * READ-ONLY — no writes anywhere. Diagnostic only.
 *
 * Sheet1 rows 26 and 28 (ABB-solgalmesp-20260327 "Sol Galmes Pons" and
 * ABB-galmespons-20260327 "Galmes Pons, Sol") are the same physical
 * booking (room 203 Allure, checkin 2026-03-27, same OTA) entered twice
 * under two different resIds. Both currently hold a WRONG Apartmentery
 * Booking ID (310885 and 311262 respectively) — the one real booking on
 * apartmentery's calendar for this room/date is 310881.
 *
 * Before deciding which resId to keep and which to retire, this checks:
 *   1. booking_done_v1 — which resId(s) are marked as having an
 *      apartmentery booking created (and what id is stored right now).
 *   2. invoice_done_v1 — whether an invoice/receipt has already been
 *      created against EITHER of the (wrong) stored bookingIds. If one
 *      of them already has money reconciled, that's the resId to keep —
 *      changing its bookingId later just corrects which apartmentery
 *      record it points to; abandoning it would orphan a real invoice.
 *   3. Payout_Income_Log — any row matching guest "Sol Galmes Pons" /
 *      "Galmes Pons" (either name order) to see status, Conf. Code, and
 *      Booking ID actually used for matching, and whether that Booking ID
 *      lines up with resId #1's or #2's stored Apartmentery Booking ID.
 *
 * HOW TO RUN: Apps Script editor ▶ checkDuplicateResIdImpact20260716 ▶ Run.
 */
function checkDuplicateResIdImpact20260716() {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const RES_IDS = ['ABB-solgalmesp-20260327', 'ABB-galmespons-20260327'];

  // --- 1. Sheet1: pull each resId's row directly ---
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  const data = src.getDataRange().getValues();
  const header = data[0];
  const idx = indexMap_(header, ['เลขห้อง', 'ชื่อแขก', 'เช็คอิน', 'เช็คเอาท์', 'Channel', 'ResId',
    APARTMENTERY_BOOKING_ID_COL_HEADER, 'วันจอง']);

  const sheetRows = {};
  for (let i = 1; i < data.length; i++) {
    const resId = String(data[i][idx.ResId] || '').trim();
    if (RES_IDS.includes(resId)) {
      sheetRows[resId] = {
        rowNum: i + 1,
        room: data[i][idx.เลขห้อง],
        guest: data[i][idx.ชื่อแขก],
        checkin: formatCellDate_(data[i][idx.เช็คอิน]),
        checkout: formatCellDate_(data[i][idx.เช็คเอาท์]),
        channel: data[i][idx.Channel],
        storedAptId: String(data[i][idx[APARTMENTERY_BOOKING_ID_COL_HEADER]] || '').trim(),
        bookedDate: formatCellDate_(data[i][idx.วันจอง]),
      };
    }
  }
  Logger.log('=== Sheet1 rows ===');
  Logger.log(JSON.stringify(sheetRows, null, 2));

  // --- 2. booking_done_v1 / invoice_done_v1 flags, keyed by resId's stored apt id ---
  const bookingDoneMap = getProp_(PROP_KEY_BOOKING_DONE);
  const invoiceDoneMap = getProp_(PROP_KEY_INVOICE_DONE);

  Logger.log('=== Flag check per resId ===');
  RES_IDS.forEach(resId => {
    const row = sheetRows[resId];
    if (!row) { Logger.log(`${resId}: NOT FOUND in Sheet1 (may have been edited already)`); return; }
    const bookingDone = !!bookingDoneMap[resId];
    const aptId = row.storedAptId;
    // invoice_done_v1 keys are usually the bare apartmentery bookingId (see
    // getInvoiceToCreate_ invoiceKey construction), occasionally
    // bookingId#confCode(#n) for split multi-guest payouts — check both the
    // bare id and any compound key that starts with it.
    const exactInvoiceDone = !!invoiceDoneMap[aptId];
    const compoundMatches = Object.keys(invoiceDoneMap).filter(k => k.indexOf(aptId + '#') === 0);
    Logger.log(`${resId}: booking_done_v1=${bookingDone}, storedAptId=${aptId}, ` +
      `invoice_done_v1[${aptId}]=${exactInvoiceDone}, compound keys matching=${JSON.stringify(compoundMatches)}`);
  });

  // --- 3. Payout_Income_Log: find any row mentioning this guest ---
  const payoutSheet = ss.getSheetByName(SRC_PAYOUT_SHEET);
  const pData = payoutSheet.getDataRange().getValues();
  const pHeader = pData[0];
  const pIdx = indexMap_(pHeader, ['วันที่ตรวจพบ', 'OTA', 'Booking ID', 'Conf. Code', 'ชื่อแขก', 'ห้อง',
    'เช็คอิน', 'เช็คเอาท์', 'สถานะ', 'หมายเหตุ']);

  const nameNeedles = ['sol galmes', 'galmes pons'];
  const payoutHits = [];
  for (let i = 1; i < pData.length; i++) {
    const guestField = String(pData[i][pIdx.ชื่อแขก] || '');
    const notesField = String(pData[i][pIdx.หมายเหตุ] || '');
    const haystack = (guestField + ' ' + notesField).toLowerCase();
    if (nameNeedles.some(n => haystack.indexOf(n) !== -1)) {
      payoutHits.push({
        rowNum: i + 1,
        detected: formatCellDate_(pData[i][pIdx['วันที่ตรวจพบ']]),
        ota: pData[i][pIdx.OTA],
        payoutBookingId: String(pData[i][pIdx['Booking ID']] || '').trim(),
        confCode: pData[i][pIdx['Conf. Code']],
        guest: guestField,
        room: pData[i][pIdx.ห้อง],
        checkin: formatCellDate_(pData[i][pIdx.เช็คอิน]),
        checkout: formatCellDate_(pData[i][pIdx.เช็คเอาท์]),
        status: pData[i][pIdx.สถานะ],
        notes: notesField,
      });
    }
  }
  Logger.log('=== Payout_Income_Log matches (name contains "Sol Galmes" or "Galmes Pons") ===');
  Logger.log(JSON.stringify(payoutHits, null, 2));

  return { sheetRows, payoutHits };
}
