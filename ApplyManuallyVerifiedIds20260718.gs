/**
 * ApplyManuallyVerifiedIds20260718.gs
 * -----------------------------------------------------------------------
 * Nathan checked apartmentery directly for the remaining unresolved rows
 * from the 2026-07-18 audit and confirmed:
 *   - THANAPORNPAN BUKBOON (TRP-thanapornp-20260616): correct id is
 *     321937 — currently sitting wrongly on ABB-milesconse-20260609
 *     (Miles Consengco's 06-09 stay row).
 *   - Miles Consengco has exactly 2 real bookings on apartmentery
 *     (https://apartmentery.com/.../booking/321078 and .../booking/322992).
 *     321078 already correctly belongs to a THIRD Miles Consengco resId,
 *     ABB-milesconse-20260603 (a 06-03 stay not in the original 14-row
 *     mismatch list — it was already fine). 322992 already correctly
 *     belongs to ABB-milesconse-20260621. That means Sheet1's THIRD Miles
 *     Consengco resId, ABB-milesconse-20260609 (06-09), has NO matching
 *     apartmentery booking at all — needs a human call (duplicate/
 *     mis-parsed resId of one of the other two stays, or a real third
 *     stay that was never created on apartmentery?), not something to
 *     guess at here. This script only CLEARS the wrong id (321937) from
 *     that row so THANAPORNPAN's fix can go through — deliberately does
 *     NOT write any replacement id for Miles's 06-09 row.
 *   - 妘芮 林 Yunjui Lin (ABB-e5a698e88a-20260403): 312166 was already
 *     correct — the earlier "guest name mismatch" was a false positive
 *     from the name matcher requiring every word to match (Sheet1's extra
 *     "林" token isn't in apartmentery's title). No write needed.
 *   - 全, 桂珍 (ABB-e585a82c20-20260219): 305941 was already correct —
 *     apartmentery just displays the guest under a differently-
 *     transliterated name ("Gao Boan"). No write needed.
 *   - Araya Rattanabamrung, Dogukan Kaner: no booking exists in
 *     apartmentery at all for either — expected, both are cancelled
 *     (ยกเลิก) so were never supposed to get one. No write needed.
 *
 * Order: clear Miles's 06-09 row first (frees 321937), THEN write
 * THANAPORNPAN's fix (uses the default uniqueness guard — will correctly
 * refuse if something unexpected still holds 321937).
 *
 * HOW TO RUN: Apps Script editor ▶ applyManuallyVerifiedIds20260718 ▶
 * Run ▶ read log.
 */

const MANUALLY_VERIFIED_20260718_ = [
  { resId: 'TRP-thanapornp-20260616', guest: 'THANAPORNPAN BUKBOON', to: '321937' }
];

// Cleared, not reassigned — see file header. No confirmed correct
// replacement id for this row yet; needs Nathan to confirm whether it's
// a real 3rd Miles Consengco stay or a duplicate resId first.
const CLEAR_ONLY_20260718_ = [
  { resId: 'ABB-milesconse-20260609', guest: 'Miles Consengco (06-09 stay)' }
];

function applyManuallyVerifiedIds20260718() {
  const report = { cleared: [], applied: [], failed: [] };

  CLEAR_ONLY_20260718_.forEach(row => {
    const before = getApartmenteryBookingId_(row.resId);
    // Empty target is never claimed by anyone, so the uniqueness guard
    // always passes here — this just blanks the column.
    const writeResult = setApartmenteryBookingId_(row.resId, '');
    if (!writeResult.ok) {
      Logger.log(`FAILED to clear ${row.resId} (${row.guest}): ${writeResult.error}`);
      report.failed.push(Object.assign({ error: writeResult.error, before }, row));
      return;
    }
    Logger.log(`CLEARED — ${row.resId} (${row.guest}): "${before}" -> "" (needs manual decision — see file header)`);
    report.cleared.push(Object.assign({ before }, row));
  });

  MANUALLY_VERIFIED_20260718_.forEach(fix => {
    const before = getApartmenteryBookingId_(fix.resId);
    const writeResult = setApartmenteryBookingId_(fix.resId, fix.to);
    if (!writeResult.ok) {
      Logger.log(`FAILED — ${fix.resId} (${fix.guest}): could not write ${fix.to} (was "${before}") — ` +
        `${writeResult.error} Nathan manually confirmed ${fix.to} on apartmentery, so this conflict means ` +
        `the OTHER resId currently holding ${fix.to} needs a look — is it a real second stay for this guest, ` +
        `or a stale duplicate that should be cleared?`);
      report.failed.push(Object.assign({ error: writeResult.error, before }, fix));
      return;
    }
    Logger.log(`FIXED — ${fix.resId} (${fix.guest}): "${before}" -> "${fix.to}" (manually verified on apartmentery by Nathan)`);
    report.applied.push(Object.assign({ before }, fix));
  });

  Logger.log(`SUMMARY: cleared=${report.cleared.length} applied=${report.applied.length} failed=${report.failed.length}`);
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}
