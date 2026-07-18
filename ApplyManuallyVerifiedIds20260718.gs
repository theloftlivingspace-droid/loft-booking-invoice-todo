/**
 * ApplyManuallyVerifiedIds20260718.gs
 * -----------------------------------------------------------------------
 * Nathan checked apartmentery directly for the remaining unresolved rows
 * from the 2026-07-18 audit and confirmed:
 *   - Miles Consengco has exactly 2 real bookings on apartmentery:
 *     https://apartmentery.com/.../booking/321078 and .../booking/322992.
 *     322992 (start 2026-06-21) was already correctly stored against
 *     ABB-milesconse-20260621 (his other, later stay) — untouched. That
 *     leaves 321078 as the only remaining candidate for
 *     ABB-milesconse-20260609 (his 06-09 stay), even though its
 *     apartmentery start date (2026-06-03) doesn't exactly match Sheet1's
 *     checkin — since apartmentery confirms only 2 Miles Consengco
 *     bookings exist total and the other one is already spoken for, this
 *     has to be it. (321078 is also the id that used to be wrongly stored
 *     against Errol Cox before the 2026-07-18 fix — freed once that was
 *     corrected.)
 *   - THANAPORNPAN BUKBOON (TRP-thanapornp-20260616): correct id is 321937
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
 * Fixes are applied in dependency order: Miles Consengco's row currently
 * holds 321937 (which is actually THANAPORNPAN's real booking), so that
 * id has to be freed from Miles's row before it can be written to
 * THANAPORNPAN's row. Uses the DEFAULT uniqueness guard (no
 * allowOverwriteConflict) — if Miles's target (322992) turns out to
 * already be claimed by the OTHER Miles Consengco resId (ABB-milesconse-
 * 20260621, a repeat stay), the guard will refuse and log it instead of
 * silently creating a new collision; that needs a manual look at whether
 * 20260621 is a genuine second stay or a stale duplicate.
 *
 * HOW TO RUN: Apps Script editor ▶ applyManuallyVerifiedIds20260718 ▶
 * Run ▶ read log.
 */

const MANUALLY_VERIFIED_20260718_ = [
  { resId: 'ABB-milesconse-20260609', guest: 'Miles Consengco', to: '321078' },
  { resId: 'TRP-thanapornp-20260616', guest: 'THANAPORNPAN BUKBOON', to: '321937' }
];

function applyManuallyVerifiedIds20260718() {
  const report = { applied: [], failed: [] };
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

  Logger.log(`SUMMARY: applied=${report.applied.length} failed=${report.failed.length}`);
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}
