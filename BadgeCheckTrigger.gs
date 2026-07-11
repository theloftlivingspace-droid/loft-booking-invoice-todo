/**
 * BadgeCheckTrigger.gs
 * -----------------------------------------------------------------------
 * External backstop for hotel-line-bot's internal */5 * * * * node-cron
 * job (runBadgeCheck → Web Push that updates the iOS home-screen badge).
 *
 * Why this exists:
 *   Same root cause as HotelJobTrigger.gs's 19:00 miss on 2026-07-09:
 *   hotel-line-bot runs on Render Free tier, and an in-process node-cron
 *   is not a reliable wall-clock guarantee on an instance that can sleep
 *   or restart (UptimeRobot pinging /health doesn't stop a cron job from
 *   silently dropping after a restart). Symptom Nathan saw: the badge
 *   only updates when he opens the app himself (which runs the
 *   client-side setForegroundBadge() fallback in AdminDailyDashboard.tsx)
 *   — meaning the server-side push was not firing on its own schedule.
 *
 * What this does:
 *   Every 10 minutes, calls POST /push/badge-check-now directly — the
 *   exact same runBadgeCheck() the internal cron calls. That function
 *   already no-ops if the count hasn't changed (compares against
 *   push_badge_last_count in Redis), so calling it more often than
 *   necessary is harmless and won't spam pushes.
 *
 * Setup:
 *   Apps Script editor ▶ select installBadgeCheckTrigger ▶ Run (once).
 *   Safe to re-run: removes any existing runBadgeCheckBackstop trigger
 *   first, so running this twice won't create duplicates.
 */
function installBadgeCheckTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'runBadgeCheckBackstop') {
      ScriptApp.deleteTrigger(t);
    }
  }
  ScriptApp.newTrigger('runBadgeCheckBackstop')
    .timeBased()
    .everyMinutes(10)
    .create();
  Logger.log('[installBadgeCheckTrigger] Trigger installed — runBadgeCheckBackstop every 10 min.');
}

function runBadgeCheckBackstop() {
  const props = PropertiesService.getScriptProperties();
  const botUrl = props.getProperty('BOT_URL') || 'https://hotel-line-bot.onrender.com';

  try {
    const response = UrlFetchApp.fetch(botUrl + '/push/badge-check-now', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({}),
      muteHttpExceptions: true
    });
    const code = response.getResponseCode();
    if (code === 200) {
      Logger.log('[runBadgeCheckBackstop] OK: ' + response.getContentText());
    } else {
      Logger.log('[runBadgeCheckBackstop] FAILED (' + code + '): ' + response.getContentText());
      _alertBadgeCheckFailure_('hotel-line-bot ตอบกลับ ' + code + ': ' + response.getContentText());
    }
  } catch (e) {
    Logger.log('[runBadgeCheckBackstop] EXCEPTION: ' + e.message);
    _alertBadgeCheckFailure_('เรียก /push/badge-check-now ไม่สำเร็จ: ' + e.message);
  }
}

/**
 * Alerts Nathan directly via LINE only when the backstop itself fails
 * (e.g. Render instance down entirely) — not on every normal run, since
 * this fires every 10 min and a normal "count unchanged, no push sent"
 * result is not something that needs an alert.
 */
function _alertBadgeCheckFailure_(detail) {
  try {
    const props = PropertiesService.getScriptProperties();
    const botUrl = props.getProperty('BOT_URL') || 'https://hotel-line-bot.onrender.com';
    const adminToken = props.getProperty('ADMIN_TOKEN') || 'apt2025@secret';
    UrlFetchApp.fetch(botUrl + '/api/send-admin-alert', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ note: '⚠️ Badge check backstop ล้มเหลว\n' + detail }),
      headers: { 'x-admin-token': adminToken },
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('[runBadgeCheckBackstop] Failed to send failure alert: ' + e.message);
  }
}
