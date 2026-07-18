/**
 * ApplyConfirmedSafeCandidates20260718.gs
 * -----------------------------------------------------------------------
 * From the 14 rows findCandidatesForUnresolved20260718 couldn't
 * auto-resolve (exact-date lookup failed), 5 had exactly ONE unclaimed
 * candidate: same room, exact name match, unclaimed by any other resId —
 * differing from Sheet1's stored checkin by exactly ONE day in every
 * single case (direction varies: some +1, some -1 vs apartmentery's
 * recorded start — not a single consistent timezone shift, but the
 * ±1-day pattern recurring across 5 unrelated bookings is still a strong
 * signal of a systematic date-boundary bug somewhere in the matching or
 * booking-creation path, not 5 coincidences — worth a follow-up
 * root-cause look, separate from this data fix.
 *
 * allowOverwriteConflict is NOT needed here — none of these candidates
 * are claimed by another resId, so the uniqueness guard in
 * setApartmenteryBookingId_ will pass normally.
 *
 * HOW TO RUN: Apps Script editor ▶ applyConfirmedSafeCandidates20260718 ▶
 * Run ▶ read log.
 */

const CONFIRMED_SAFE_20260718_ = [
  { resId: 'ABB-johnzambra-20260428', guest: 'John Zambrana', from: '315889', to: '314905' },
  { resId: 'ABB-premmehta-20260501', guest: 'Prem Mehta', from: '315890', to: '315888' },
  { resId: 'ABB-kgotlellom-20260528', guest: 'Kgotlello Masemola', from: '317454', to: '317758' },
  { resId: 'ABB-errolcox-20260608', guest: 'Errol Cox', from: '321078', to: '321724' },
  { resId: 'ABB-saeidmickm-20260610', guest: 'Saeid Mick Momtahan', from: '322314', to: '321719' }
];

function applyConfirmedSafeCandidates20260718() {
  const report = { applied: [], failed: [] };
  CONFIRMED_SAFE_20260718_.forEach(fix => {
    const writeResult = setApartmenteryBookingId_(fix.resId, fix.to);
    if (!writeResult.ok) {
      Logger.log(`FAILED — ${fix.resId} (${fix.guest}): could not write ${fix.to} — ${writeResult.error} Needs manual review, something changed since findCandidatesForUnresolved20260718 ran.`);
      report.failed.push(Object.assign({ error: writeResult.error }, fix));
      return;
    }
    Logger.log(`FIXED — ${fix.resId} (${fix.guest}): "${fix.from}" -> "${fix.to}"`);
    report.applied.push(fix);
  });

  Logger.log(`SUMMARY: applied=${report.applied.length} failed=${report.failed.length}`);
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}
