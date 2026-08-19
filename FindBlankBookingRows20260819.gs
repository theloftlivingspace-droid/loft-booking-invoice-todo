/**
 * FindBlankBookingRows20260819.gs
 * -----------------------------------------------------------------------
 * READ-ONLY audit — does NOT write to Sheet1.
 *
 * The admin dashboard's Booking & Invoice To-Do list showed a card with
 * "NaN" / "Invalid Date 'N" for check-in and check-out. Root cause (traced
 * in the-loft-admin's BookingInvoiceTodo.tsx): the date block does
 * `new Date(item.checkin).getDate()` etc. — if checkin/checkout is an
 * empty string, that yields NaN/"Invalid Date". getBookingToAdd_() in
 * Code.gs keeps a row as long as ANY cell is non-empty (r.join('').trim()
 * !== ''), so a row with e.g. a ResId or Note but blank เช็คอิน/เช็คเอาท์
 * still surfaces in the to-do list as a broken-looking card.
 *
 * This scans every row in Sheet1 and reports ones that are "present" but
 * missing เช็คอิน and/or เช็คเอาท์ (the ones that would render as NaN),
 * plus a broader check for any row missing ชื่อแขก, เลขห้อง, or ResId too,
 * so we can see the full shape of the bad row before deciding what to
 * fix (delete row, fill in the missing dates, etc.) — nothing is changed
 * here.
 *
 * Run findBlankBookingRows20260819() manually from the Apps Script editor;
 * check the execution log for the report. Row numbers are 1-indexed to
 * match what you'd see in the Sheet1 UI (header = row 1).
 */
function findBlankBookingRows20260819() {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  if (!src) { Logger.log('Sheet not found: ' + SRC_BOOKING_SHEET); return; }

  const data = src.getDataRange().getValues();
  const header = data[0];
  const idx = indexMap_(header, ['เลขห้อง', 'ชื่อแขก', 'เช็คอิน', 'เช็คเอาท์', 'Channel', 'ResId', 'Note', 'Apartmentery Booking ID']);

  const missing = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowJoined = row.join('').trim();
    if (rowJoined === '') continue; // fully empty row — not what surfaces in the to-do list, skip

    const room    = String(row[idx['เลขห้อง']] || '').trim();
    const guest   = String(row[idx['ชื่อแขก']] || '').trim();
    const checkin = row[idx['เช็คอิน']];
    const checkout = row[idx['เช็คเอาท์']];
    const resId   = String(row[idx['ResId']] || '').trim();
    const channel = String(row[idx['Channel']] || '').trim();
    const note    = String(row[idx['Note']] || '').trim();

    const checkinEmpty  = checkin === '' || checkin === null || checkin === undefined;
    const checkoutEmpty = checkout === '' || checkout === null || checkout === undefined;

    if (checkinEmpty || checkoutEmpty || !guest || !room || !resId) {
      missing.push({
        sheetRow: i + 1, // 1-indexed, matches Sheet1 UI row number
        room: room || '(blank)',
        guest: guest || '(blank)',
        checkin: checkinEmpty ? '(blank)' : String(checkin),
        checkout: checkoutEmpty ? '(blank)' : String(checkout),
        resId: resId || '(blank)',
        channel: channel || '(blank)',
        note: note || '(blank)',
      });
    }
  }

  Logger.log(`findBlankBookingRows20260819: ${data.length - 1} data row(s) scanned, ${missing.length} incomplete row(s) found.`);
  missing.forEach(r => {
    Logger.log(
      `Sheet1 row ${r.sheetRow} — room="${r.room}" guest="${r.guest}" checkin="${r.checkin}" ` +
      `checkout="${r.checkout}" resId="${r.resId}" channel="${r.channel}" note="${r.note}"`
    );
  });
  Logger.log('Full list: ' + JSON.stringify(missing, null, 2));
  return missing;
}
