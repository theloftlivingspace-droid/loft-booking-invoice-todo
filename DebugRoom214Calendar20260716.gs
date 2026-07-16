/**
 * DebugRoom214Calendar20260716.gs
 * -----------------------------------------------------------------------
 * READ-ONLY diagnostic — does not write anything to Sheet1 or apartmentery.
 * Dumps every {title, start, url} calendar event block found on room 214's
 * (unit 163867) apartmentery booking-listing page, so we can eyeball which
 * entry is really Nicco Joselito Tan's (checkin 2026-07-04) vs which one
 * the automated matcher grabbed (325355 — currently also stored against
 * Supasuta Landreaugrasmuck in room 300, checkin 2026-06-30). Run this from
 * the Apps Script editor and paste the log back.
 */
function debugRoom214Calendar20260716() {
  const unit = getApartmenteryUnitForRoom('214');
  if (!unit) {
    Logger.log('Room 214 not found in ROOM_TO_UNIT_ID.');
    return;
  }
  const path = `/user/branch/${unit.branchId}/unit/${unit.unitId}/booking`;
  const response = _apartmenteryFetch_(path, { method: 'get' });
  Logger.log(`GET ${path} -> HTTP ${response.getResponseCode()}`);
  if (response.getResponseCode() !== 200) {
    Logger.log('Non-200 response, first 2000 chars: ' + response.getContentText().slice(0, 2000));
    return;
  }

  const html = response.getContentText();
  const blockRe = /\{\s*title:\s*'((?:[^'\\]|\\.)*)'[\s\S]*?start:\s*'([^']*)'[\s\S]*?url:\s*'([^']*)'\s*\}/g;
  let m;
  const events = [];
  while ((m = blockRe.exec(html)) !== null) {
    const title = m[1];
    const start = m[2];
    const url = m[3];
    const idMatch = url.match(/\/booking\/(\d+)/);
    events.push({ title: title, start: start, isoStart: _apartmenteryCalendarDateToIso_(start), id: idMatch ? idMatch[1] : null });
  }

  Logger.log(`Found ${events.length} calendar event blocks on room 214's listing page:`);
  events.forEach(e => Logger.log(JSON.stringify(e)));

  // Specifically highlight anything matching "Nicco" or matching the
  // suspect id 325355, so it's easy to spot without reading the whole dump.
  const relevant = events.filter(e =>
    /nicco/i.test(e.title) || e.id === '325355'
  );
  Logger.log('Rows relevant to Nicco Joselito Tan / id 325355: ' + JSON.stringify(relevant, null, 2));
}
