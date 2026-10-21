/**
 * AllDocsWarmTrigger.gs
 * -----------------------------------------------------------------------
 * Keeps the getAllDocs_ cache warm so the "View Documents" button on the
 * Check-in/out page never waits on a cold full-Drive scan (which can exceed
 * the 20s Vercel proxy limit and made the docs list silently fail to load).
 *
 * Hourly is enough: upload/delete patch the cache in place (Code.gs), so the
 * trigger only catches files changed directly in Drive. Kept hourly rather
 * than every few minutes to stay well inside Apps Script's daily trigger
 * runtime quota (the scan itself can take tens of seconds).
 *
 * Setup:
 *   Apps Script editor ▶ select installAllDocsWarmTrigger ▶ Run (once).
 *   Safe to re-run: removes any existing warmAllDocsCache_ trigger first.
 *   (Also run warmAllDocsCache_ once by hand to fill the cache immediately.)
 */
function installAllDocsWarmTrigger() {
  for (const t of ScriptApp.getProjectTriggers()) {
    if (t.getHandlerFunction() === 'warmAllDocsCache_') ScriptApp.deleteTrigger(t);
  }
  ScriptApp.newTrigger('warmAllDocsCache_').timeBased().everyHours(1).create();
  Logger.log('[installAllDocsWarmTrigger] Installed — warmAllDocsCache_ every hour.');
}
