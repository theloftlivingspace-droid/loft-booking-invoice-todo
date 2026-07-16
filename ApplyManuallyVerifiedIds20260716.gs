/**
 * ApplyManuallyVerifiedIds20260716.gs
 * -----------------------------------------------------------------------
 * Applies the 9 remaining Apartmentery Booking ID corrections that Nathan
 * looked up himself directly on apartmentery.com (not via the automated
 * name/date matcher, which couldn't resolve these with confidence — see
 * auditAllApartmenteryBookingIds' GUEST_MISMATCH / WRONG_ROOM output from
 * earlier today). Writes each value exactly as given, no lookup involved.
 *
 * Run once from the Apps Script editor, then re-run
 * auditAllApartmenteryBookingIds to confirm the mismatch count drops to 0
 * and nothing new collides.
 */
function applyManuallyVerifiedIds20260716() {
  const rows = [
    { resId: 'ABB-e585a82c20-20260219', guest: '全, 桂珍',              id: '305941' },
    { resId: 'ABB-e5a698e88a-20260403', guest: '妘芮 林 Yunjui Lin',     id: '312166' },
    { resId: 'ABB-avtodagdel-20260405', guest: 'Avto Dagdelen',         id: '312165' },
    { resId: 'ABB-lataviaant-20260512', guest: "La'Tavia Antrice",      id: '317138' },
    { resId: 'ABB-errolcox-20260608',   guest: 'Errol Cox',             id: '321724' },
    { resId: 'ABB-milesconse-20260609', guest: 'Miles Consengco',       id: '321078' },
    { resId: 'TRP-thanapornp-20260616', guest: 'THANAPORNPAN BUKBOON',  id: '321937' },
    { resId: 'TRP-arayaratta-20260518', guest: 'Araya Rattanabamrung',  id: '317749' },
    { resId: 'TRP-pornpawitb-20260616', guest: 'Pornpawit Boon',        id: '322279' }
  ];

  const report = [];
  rows.forEach(r => {
    const before = getApartmenteryBookingId_(r.resId);
    setApartmenteryBookingId_(r.resId, r.id);
    Logger.log(`applyManuallyVerifiedIds20260716: ${r.resId} (${r.guest}) "${before}" -> "${r.id}"`);
    report.push({ resId: r.resId, guest: r.guest, from: before, to: r.id });
  });

  Logger.log('Done — all 9 rows written. Re-run auditAllApartmenteryBookingIds to confirm 0 mismatches remain.');
  return report;
}
