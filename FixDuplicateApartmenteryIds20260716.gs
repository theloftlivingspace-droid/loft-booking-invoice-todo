/**
 * FixDuplicateApartmenteryIds20260716.gs
 * -----------------------------------------------------------------------
 * One-off fix for the 9 resId pairs discovered 2026-07-16 sharing the same
 * (wrong, for one side) Apartmentery Booking ID in Sheet1 — caused by the
 * name-only listing-page fallback bug fixed in ApartmenteryClient.gs
 * (commit d674e49). This does NOT run automatically — must be run manually
 * from the Apps Script editor, since it needs the live APARTMENTERY_SESSION
 * cookie already stored in this project's Script Properties (not available
 * to Claude outside this environment).
 *
 * For each row below, re-looks-up the real bookingId via
 * findApartmenteryBookingIdForRoomByGuest_ (date-verified match) and, if a
 * confident match is found AND it differs from what's currently stored,
 * overwrites Sheet1's Apartmentery Booking ID column for that resId.
 * Anything that can't be confidently resolved is only logged — never
 * guessed — for manual review in apartmentery itself.
 *
 * HOW TO RUN:
 *   Apps Script editor ▶ select fixDuplicateApartmenteryIds20260716 from
 *   the function dropdown ▶ Run. Check the execution log afterward.
 */
function fixDuplicateApartmenteryIds20260716() {
  const rows = [
    // pair: 325684
    { resId: 'OTH-chawsukhai-20260701', room: '204 Elegance', guest: 'CHAW SU KHAING',          checkin: '2026-07-01' },
    { resId: 'ABB-andreamast-20260704', room: '204 Elegance', guest: 'Andrea Mastropietro',      checkin: '2026-07-15' },
    // pair: 325792 (both cancelled — should end up with NO bookingId; see note below)
    { resId: 'BKC-murtazahus-20260701', room: '205 Allure ยกเลิก', guest: 'Murtaza Hussain',     checkin: '2026-07-01' },
    { resId: 'ABB-e4bdb0e9a1-20260705', room: '204 Elegance ยกเลิก', guest: '佰顺 王',            checkin: '2026-07-05' },
    // pair: 326107
    { resId: 'ABB-niccojosel-20260702', room: '214 Legacy', guest: 'Nicco Joselito Tan',         checkin: '2026-07-04' },
    { resId: 'ABB-ameerahmcn-20260708', room: '214 Legacy', guest: 'Ameerah McNeill',            checkin: '2026-07-08' },
    // pair: 326166
    { resId: 'ABB-lukefaisal-20260702', room: '214 Legacy', guest: 'Luke Faisal',                checkin: '2026-07-02' },
    { resId: 'BKC-yeabsirate-20260708', room: '205 Allure', guest: 'Yeabsira Tefera',            checkin: '2026-07-09' },
    // pair: 326163
    { resId: 'ABB-leoyang-20260703',    room: '108 Retro',  guest: 'Leo Yang',                   checkin: '2026-07-03' },
    { resId: 'BKC-watcharapo-20260709', room: '209 Radiance ยกเลิก', guest: 'Watcharaporn Chaisura', checkin: '2026-07-09' },
    // pair: 326228 (Moritz Reinhold cancelled — should end up with NO bookingId)
    { resId: 'ABB-moritzrein-20260704', room: '204 Elegance ยกเลิก', guest: 'Moritz Reinhold',   checkin: '2026-07-04' },
    { resId: 'TRP-eaintphooh-20260709', room: '204 Elegance', guest: 'Eaint Phoo Htet',          checkin: '2026-07-10' },
    // pair: 326530
    { resId: 'ABB-henrycarde-20260704', room: '209 Radiance', guest: 'Henry Cardenas Paspuezan', checkin: '2026-07-05' },
    { resId: 'EXP-natphatsorn-20260711', room: '300 Luxury',  guest: 'Natphatsorn wongwai',      checkin: '2026-07-11' },
    // pair: 326550
    { resId: 'ABB-crystalesp-20260704', room: '210 Radiance', guest: 'Crystal Espinoza',         checkin: '2026-07-04' },
    { resId: 'ABB-zhgggtr-20260712',    room: '209 Radiance', guest: 'Z Hgggtr',                 checkin: '2026-07-12' },
    // pair: 326815
    { resId: 'EXP-praneeanto-20260711', room: '214 Legacy',   guest: 'Pranee Antov',             checkin: '2026-07-18' },
    { resId: 'ABB-elijahbras-20260715', room: '113 Legacy',   guest: 'Elijah Brasil',            checkin: '2026-07-17' }
  ];

  const report = { fixed: [], unresolved: [], unchanged: [] };

  rows.forEach(r => {
    const currentId = getApartmenteryBookingId_(r.resId);
    let foundId = null;
    try {
      foundId = findApartmenteryBookingIdForRoomByGuest_(r.room, r.guest, r.checkin);
    } catch (err) {
      if (isApartmenterySessionExpiredError(err)) {
        Logger.log('SESSION EXPIRED — stopping. Refresh APARTMENTERY_SESSION and re-run.');
        throw err;
      }
      Logger.log(`lookup failed for ${r.resId} (${r.guest}, ${r.room}, ${r.checkin}): ${err.message}`);
      report.unresolved.push({ resId: r.resId, guest: r.guest, reason: err.message });
      return;
    }

    if (!foundId) {
      Logger.log(`NO MATCH for ${r.resId} (${r.guest}, ${r.room}, ${r.checkin}) — current stored id: "${currentId}". ` +
        `Check apartmentery manually; a cancelled booking may legitimately have no calendar entry.`);
      report.unresolved.push({ resId: r.resId, guest: r.guest, currentId: currentId });
      return;
    }

    if (foundId === currentId) {
      Logger.log(`OK — ${r.resId} (${r.guest}) already correct: ${currentId}`);
      report.unchanged.push({ resId: r.resId, guest: r.guest, id: currentId });
      return;
    }

    setApartmenteryBookingId_(r.resId, foundId);
    Logger.log(`FIXED — ${r.resId} (${r.guest}, ${r.room}): "${currentId}" -> "${foundId}"`);
    report.fixed.push({ resId: r.resId, guest: r.guest, from: currentId, to: foundId });
  });

  Logger.log('fixDuplicateApartmenteryIds20260716 report: ' + JSON.stringify(report, null, 2));
  return report;
}
