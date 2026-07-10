/**
 * HotelJobTrigger.gs
 * -----------------------------------------------------------------------
 * External backstop for hotel-line-bot's internal 19:00 node-cron job
 * (runHotelJob → daily check-in/out summary to the maid LINE group).
 *
 * Why this exists:
 *   hotel-line-bot runs on Render Free tier, which spins the instance
 *   down after ~15 min of inactivity. UptimeRobot pings /health every
 *   5 min, but on 2026-07-09 the 19:00 cron never fired (confirmed via
 *   Render logs — no "เริ่มส่งสรุปแม่บ้าน..." log line at all around
 *   19:00, jumping straight from 18:30 to 19:30). Root cause not fully
 *   pinned down (sleep vs. cron-registration drop after a restart), but
 *   either way an external trigger with a real wall-clock guarantee is
 *   more reliable than trusting an in-process cron on a instance that
 *   can sleep.
 *
 * What this does:
 *   Calls hotel-line-bot's /api/test-hotel-job endpoint (added
 *   2026-07-09, runs the exact same runHotelJob() the internal cron
 *   uses) via a POST request. The HTTP request itself also serves as
 *   an extra wake-up ping for the Render instance, on top of
 *   UptimeRobot.
 *
 * Setup:
 *   1. Reuses BOT_URL / ADMIN_TOKEN Script Properties already set for
 *      ApartmenteryClient.gs in this same GAS project — no new
 *      properties needed.
 *   2. Apps Script editor ▶ Triggers ▶ Add Trigger:
 *        Function: triggerHotelJob19
 *        Event source: Time-driven
 *        Type: Day timer
 *        Time: 7pm to 8pm (pick a specific time close to 19:00,
 *              e.g. 7:05pm, to run slightly after the internal cron
 *              in case that one succeeds first)
 *   3. Keep the internal node-cron running too — this is a backstop,
 *      not a replacement. If both fire, the maid group just gets the
 *      same summary twice; harmless but worth knowing. If you'd rather
 *      avoid the duplicate, remove the internal cron.schedule(CRON_SCHED, ...)
 *      call in hotel-line-bot/bot.js once this trigger is confirmed
 *      reliable for a few days.
 *
 * On failure:
 *   Logs to Apps Script's own execution log (Executions ▶ this run)
 *   AND sends a LINE alert straight to ADMIN_USER (Nathan) via
 *   /api/send-admin-alert, so a failure here doesn't go unnoticed
 *   the way the silent 19:00 miss did.
 */
function triggerHotelJob19() {
  const props = PropertiesService.getScriptProperties();
  const botUrl = props.getProperty('BOT_URL') || 'https://hotel-line-bot.onrender.com';
  const adminToken = props.getProperty('ADMIN_TOKEN') || 'apt2025@secret';

  try {
    const response = UrlFetchApp.fetch(botUrl + '/api/test-hotel-job', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({}),
      headers: { 'x-admin-token': adminToken },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const body = response.getContentText();

    if (code === 200) {
      Logger.log('[triggerHotelJob19] OK: ' + body);
    } else {
      Logger.log('[triggerHotelJob19] FAILED (' + code + '): ' + body);
      _alertHotelJobFailure_('hotel-line-bot ตอบกลับ ' + code + ': ' + body);
    }
  } catch (e) {
    Logger.log('[triggerHotelJob19] EXCEPTION: ' + e.message);
    _alertHotelJobFailure_('เรียก /api/test-hotel-job ไม่สำเร็จ: ' + e.message);
  }
}

/**
 * Sends a 1:1 LINE alert to ADMIN_USER (Nathan) via hotel-line-bot's
 * /api/send-admin-alert — same endpoint _notifyLineSessionFailure_()
 * in ApartmenteryClient.gs uses, kept separate here so a failure in
 * one doesn't depend on the other's internal helper.
 */
function _alertHotelJobFailure_(detail) {
  try {
    const props = PropertiesService.getScriptProperties();
    const botUrl = props.getProperty('BOT_URL') || 'https://hotel-line-bot.onrender.com';
    const adminToken = props.getProperty('ADMIN_TOKEN') || 'apt2025@secret';
    UrlFetchApp.fetch(botUrl + '/api/send-admin-alert', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ note: '⚠️ 19:00 hotel job (GAS backstop) พลาดด้วย\n' + detail }),
      headers: { 'x-admin-token': adminToken },
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('[triggerHotelJob19] Failed to send failure alert: ' + e.message);
  }
}
