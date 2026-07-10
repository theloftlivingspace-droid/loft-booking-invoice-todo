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
 *   First checks /api/hotel-job-status (GET) — hotel-line-bot records
 *   a Redis flag (hotel_job_last_run) whenever runHotelJob() succeeds,
 *   whether triggered by its own cron or by this backstop. If today's
 *   flag is already set, this function does nothing (Render's cron
 *   worked fine, no need to duplicate). Only if it's NOT set does it
 *   call /api/test-hotel-job (POST) — the same runHotelJob() the
 *   internal cron uses — as a fallback. This way the maid group only
 *   ever gets the summary once per day, from whichever path fires
 *   first.
 *
 * Setup:
 *   1. Reuses BOT_URL / ADMIN_TOKEN Script Properties already set for
 *      ApartmenteryClient.gs in this same GAS project — no new
 *      properties needed.
 *   2. Apps Script editor ▶ Triggers ▶ Add Trigger:
 *        Function: triggerHotelJob19
 *        Event source: Time-driven
 *        Type: Day timer
 *        Time: 7pm to 8pm (pick a specific time a few minutes after
 *              19:00, e.g. 7:10pm, to give Render's own cron a chance
 *              to run first)
 *   3. Keep the internal node-cron running too — this is a backstop
 *      only, and won't duplicate the message thanks to the status
 *      check above.
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

  // Step 1: check whether Render's own 19:00 cron already ran successfully
  // today. Only fall back to firing the job ourselves if it didn't —
  // this is a backstop, not a duplicate sender.
  try {
    const statusResp = UrlFetchApp.fetch(botUrl + '/api/hotel-job-status', {
      method: 'get',
      headers: { 'x-admin-token': adminToken },
      muteHttpExceptions: true
    });
    const statusCode = statusResp.getResponseCode();
    const statusBody = JSON.parse(statusResp.getContentText());

    if (statusCode === 200 && statusBody.ranToday) {
      Logger.log('[triggerHotelJob19] Render cron already ran today (' + statusBody.lastRun + ') — skipping, no action needed.');
      return;
    }
    Logger.log('[triggerHotelJob19] Render cron has NOT run today yet (lastRun=' + statusBody.lastRun + ') — firing backstop.');
  } catch (e) {
    // Status check itself failed (e.g. instance asleep, Redis down) —
    // treat as "didn't run" and fire the backstop anyway, better safe
    // than silently missing the summary again.
    Logger.log('[triggerHotelJob19] Status check failed (' + e.message + ') — assuming it did not run, firing backstop anyway.');
  }

  // Step 2: fire the job ourselves.
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
      Logger.log('[triggerHotelJob19] Backstop OK: ' + body);
    } else {
      Logger.log('[triggerHotelJob19] Backstop FAILED (' + code + '): ' + body);
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
