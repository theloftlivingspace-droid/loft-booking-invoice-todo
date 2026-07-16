/**
 * ApplyManuallyVerifiedIdsRound2_20260716.gs
 * -----------------------------------------------------------------------
 * Final 2 corrections, manually verified by Nathan directly on
 * apartmentery.com:
 *   - Dogukan Kaner (room 214, checkin 2026-05-31): real id 320478
 *   - Araya Rattanabamrung (room 103, checkin 2026-05-18): real id 317749
 *     — confirmed this is the same booking as "aomsub laosonti" on
 *     apartmentery's calendar (Trip.com account-holder name differs from
 *     the actual guest name in Sheet1); NOT a collision with
 *     TRP-aomsublaos-20260519, which is a separate booking.
 */
function applyManuallyVerifiedIdsRound2_20260716() {
  const rows = [
    { resId: 'ABB-dogukankan-20260531', guest: 'Dogukan Kaner',         id: '320478' },
    { resId: 'TRP-arayaratta-20260518', guest: 'Araya Rattanabamrung',  id: '317749' }
  ];

  const report = [];
  rows.forEach(r => {
    const before = getApartmenteryBookingId_(r.resId);
    setApartmenteryBookingId_(r.resId, r.id);
    Logger.log(`applyManuallyVerifiedIdsRound2_20260716: ${r.resId} (${r.guest}) "${before}" -> "${r.id}"`);
    report.push({ resId: r.resId, guest: r.guest, from: before, to: r.id });
  });

  Logger.log('Done. Re-run auditAllApartmenteryBookingIds to confirm 0 real problems remain (the 3 romanization-only name mismatches — 全,桂珍 / Yunjui Lin / La\'Tavia Antrice — are expected and fine).');
  return report;
}
