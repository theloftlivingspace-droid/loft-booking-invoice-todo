/**
 * ============================================================
 *  The Loft Living Space — Booking & Invoice To-Do Webapp
 * ------------------------------------------------------------
 *  v2 — improved cross-sheet matching
 *   - allNameParts(): ดึงทุก word ≥3 chars จากชื่อ (รองรับ "Last, First", "FIRST LAST", Thai)
 *   - roomNum(): แยกเลขห้อง 3 หลักออกจาก "205 Allure", "113 Legacy" ฯลฯ
 *   - makeMatchKeys(): สร้าง key ทั้งแบบ name|date และ date|room พร้อม ±2 วัน window
 *   - copyText ใน Booking card: copy "ชื่อแขก / Channel" ครบ
 * ============================================================
 */

// ===== CONFIG =====
const SOURCE_SHEET_ID = '1XbTJLhecql_HNqyE80Hc6h30A2_elIxliudF4e6Rlz0';

const SRC_BOOKING_SHEET = 'Sheet1';
const SRC_PAYOUT_SHEET = 'Payout_Income_Log';

const PAYOUT_STATUSES_FOR_INVOICE = [
  'โอนแล้ว',
  'โอนแล้ว (Resolution Payout)',
  '✅ Matched - Airbnb payout',
  '✅ Matched - Booking.com remittance',
  '✅ Matched - Expedia remittance',
  '✅ Matched - Trip.com settlement',
];

const PROP_KEY_BOOKING_DONE = 'booking_done_v1';
const PROP_KEY_INVOICE_DONE = 'invoice_done_v1';
const PROP_KEY_BOOKING_SEEN = 'booking_seen_v1';
const PROP_KEY_INVOICE_SEEN = 'invoice_seen_v1';

/* ============================================================
 *  Web app entry point
 *  GET ?action=getData        → JSON API (for Vercel/React)
 *  GET ?action=setBookingDone&id=X&done=true  → JSON
 *  GET ?action=setInvoiceDone&id=X&done=true  → JSON
 *  GET (no action)            → serve HTML webapp
 * ============================================================ */
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'getData') {
    return jsonResponse_(getDashboardData());
  }

  if (action === 'setBookingDone') {
    const id   = e.parameter.id   || '';
    const done = e.parameter.done === 'true';
    setBookingDone(id, done);
    return jsonResponse_({ ok: true });
  }

  if (action === 'setInvoiceDone') {
    const id   = e.parameter.id   || '';
    const done = e.parameter.done === 'true';
    setInvoiceDone(id, done);
    return jsonResponse_({ ok: true });
  }

  // Default: serve HTML webapp
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('The Loft — Booking & Invoice To-Do')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function jsonResponse_(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

/* ============================================================
 *  Data loading
 * ============================================================ */
function getDashboardData() {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const todayStr = formatDateYMD_(new Date());
  return {
    today: todayStr,
    booking: getBookingToAdd_(ss, todayStr),
    invoice: getInvoiceToCreate_(ss, todayStr),
  };
}

function getBookingToAdd_(ss, todayStr) {
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  if (!src) throw new Error('ไม่พบชีต: ' + SRC_BOOKING_SHEET);

  const data = src.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1).filter(r => r.join('').trim() !== '');

  const idx = indexMap_(header, ['เลขห้อง', 'ชื่อแขก', 'เช็คอิน', 'เช็คเอาท์', 'Channel', 'ResId', 'Note']);

  const doneMap = getProp_(PROP_KEY_BOOKING_DONE);
  const seenMap = getProp_(PROP_KEY_BOOKING_SEEN);
  let seenChanged = false;

  const out = rows.map(r => {
    const resId   = String(r[idx.ResId] || '').trim();
    const guest   = String(r[idx['ชื่อแขก']] || '').trim();
    const checkin = formatCellDate_(r[idx['เช็คอิน']]);
    const room    = String(r[idx['เลขห้อง']] || '').trim();
    const channel = String(r[idx.Channel] || '').trim();

    let firstSeen = seenMap[resId];
    let isNewToday = false;
    if (!firstSeen) {
      firstSeen = todayStr; seenMap[resId] = todayStr;
      seenChanged = true; isNewToday = true;
    }

    return {
      resId, room, guest, checkin,
      checkout: formatCellDate_(r[idx['เช็คเอาท์']]),
      channel,
      note: String(r[idx.Note] || ''),
      firstSeen, isNewToday,
      done: !!doneMap[resId],
      matchKeys: makeMatchKeys_(guest, checkin, room),
    };
  });

  if (seenChanged) setProp_(PROP_KEY_BOOKING_SEEN, seenMap);

  // Dedupe bookings that are the same stay parsed twice under different resIds
  // (e.g. guest name parsed as "Sol Galmes Pons" vs "Galmes Pons, Sol" → two different resIds,
  // same room + same checkin + same checkout). Keep the one with the earlier firstSeen
  // (i.e. detected first), or alphabetically-first resId as a tiebreaker.
  const dedupMap = {};
  out.forEach(b => {
    const dupKey = roomNum_(b.room) + '|' + b.checkin + '|' + b.checkout;
    const existing = dedupMap[dupKey];
    if (!existing) {
      dedupMap[dupKey] = b;
    } else if (b.firstSeen < existing.firstSeen || (b.firstSeen === existing.firstSeen && b.resId < existing.resId)) {
      dedupMap[dupKey] = b;
    }
  });
  const deduped = Object.values(dedupMap);

  deduped.reverse();
  return deduped;
}

function getInvoiceToCreate_(ss, todayStr) {
  const src = ss.getSheetByName(SRC_PAYOUT_SHEET);
  if (!src) throw new Error('ไม่พบชีต: ' + SRC_PAYOUT_SHEET);

  const data = src.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1).filter(r => r.join('').trim() !== '');

  const idx = indexMap_(header, [
    'วันที่ตรวจพบ', 'OTA', 'Booking ID', 'Conf. Code', 'ชื่อแขก', 'ห้อง',
    'เช็คอิน', 'เช็คเอาท์', 'คืน', 'ยอดรวม (THB)', 'Commission (THB)', 'NET (THB)', 'สถานะ', 'หมายเหตุ',
  ]);

  const matchedConfCodes = new Set();
  rows.forEach(r => {
    const status = String(r[idx.สถานะ] || '').trim();
    if (status.includes('Matched')) {
      String(r[idx['Conf. Code']] || '').split(',').forEach(c => matchedConfCodes.add(c.trim()));
    }
  });

  // Build set of bookingIds that have a summary row (Conf. Code contains comma)
  const summaryBookingIds = new Set();
  rows.forEach(r => {
    const confCode  = String(r[idx['Conf. Code']] || '').trim();
    const bookingId = String(r[idx['Booking ID']] || '').trim();
    if (confCode.includes(',') && bookingId) summaryBookingIds.add(bookingId);
  });

  const filtered = rows.filter(r => {
    const status    = String(r[idx.สถานะ] || '').trim();
    const confCode  = String(r[idx['Conf. Code']] || '').trim();
    const note      = String(r[idx.หมายเหตุ] || '').trim();
    const bookingId = String(r[idx['Booking ID']] || '').trim();
    if (note.startsWith('↳')) return false;
    if (!PAYOUT_STATUSES_FOR_INVOICE.includes(status)) return false;
    if (!status.includes('Matched') && matchedConfCodes.has(confCode)) return false;
    // Skip single-conf sub-rows when a multi-conf summary row exists for same bookingId
    if (!confCode.includes(',') && summaryBookingIds.has(bookingId)) return false;
    return true;
  });

  const doneMap = getProp_(PROP_KEY_INVOICE_DONE);
  const seenMap = getProp_(PROP_KEY_INVOICE_SEEN);
  let seenChanged = false;
  const out = [];

  filtered.forEach(r => {
    const bookingId    = String(r[idx['Booking ID']] || '').trim();
    const detectedDate = formatCellDate_(r[idx['วันที่ตรวจพบ']]);
    const room         = String(r[idx.ห้อง] || '');
    const checkin      = formatCellDate_(r[idx.เช็คอิน]);
    const checkout     = formatCellDate_(r[idx.เช็คเอาท์]);
    const nights       = r[idx.คืน] || '';
    const ota          = String(r[idx.OTA] || '');
    const status       = String(r[idx.สถานะ] || '');
    const totalNet     = r[idx['NET (THB)']] || '';
    const rawConfCode  = String(r[idx['Conf. Code']] || '');
    const rawGuestField= String(r[idx.ชื่อแขก] || '');
    const notes        = String(r[idx.หมายเหตุ] || '');

    const firstGuest    = rawGuestField.split(',')[0].trim();
    const firstConfCode = rawConfCode.split(',')[0].trim();

    // parse ทุก "Guest(conf) NET ฿amount" entry จาก notes
    // ทุก case: conf เดียว, conf ต่างกัน, conf ซ้ำ (split payout เช่น Nihel 81.17 + 2638.54)
    const subPattern = /([^|]+?)\(([^)]+)\)\s*NET\s+฿([\d,]+\.?\d*)/g;
    const subs = [];
    let m;
    while ((m = subPattern.exec(notes)) !== null) {
      subs.push({ guest: m[1].trim(), confCode: m[2].trim(), net: parseFloat(m[3].replace(/,/g,'')) });
    }

    const entries = subs.length > 0 ? subs : [{ guest: firstGuest, confCode: firstConfCode, net: totalNet }];

    // invoiceKey: ถ้ามีหลาย entries และ conf ซ้ำ ใส่ index กำกับ (#0, #1)
    const confCount = {};
    entries.forEach(e => { confCount[e.confCode] = (confCount[e.confCode] || 0) + 1; });
    const confIdx = {};

    entries.forEach((entry, i) => {
      const hasDupeConf = confCount[entry.confCode] > 1;
      const ci2 = confIdx[entry.confCode] = (confIdx[entry.confCode] || 0);
      confIdx[entry.confCode]++;
      const invoiceKey = entries.length > 1
        ? bookingId + '#' + entry.confCode + (hasDupeConf ? '#' + ci2 : '')
        : bookingId;
      let firstSeen = seenMap[invoiceKey];
      let isNewSeen = false;
      if (!firstSeen) {
        firstSeen = todayStr; seenMap[invoiceKey] = todayStr;
        seenChanged = true; isNewSeen = true;
      }
      out.push({
        invoiceKey, bookingId, room,
        guest: entry.guest || firstGuest,
        checkin, checkout, nights,
        net: entries.length > 1 ? entry.net : totalNet,
        isSplitFromMulti: entries.length > 1,
        splitIndex: entries.length > 1 ? (i + 1) : null,
        splitTotal: entries.length > 1 ? entries.length : null,
        groupNet: entries.length > 1 ? totalNet : null,
        ota, status, detectedDate,
        detectedToday: detectedDate === todayStr,
        firstSeen, isNewInList: isNewSeen,
        done: !!doneMap[invoiceKey],
        matchKeys: makeMatchKeys_(entry.guest || firstGuest, checkin, room),
      });
    });
  });

  if (seenChanged) setProp_(PROP_KEY_INVOICE_SEEN, seenMap);
  out.sort((a, b) => {
    if (a.detectedDate !== b.detectedDate) return a.detectedDate < b.detectedDate ? 1 : -1;
    return a.guest < b.guest ? -1 : 1;
  });
  return out;
}

/* ============================================================
 *  Matching key builder — v2
 *  สร้าง key หลายแบบ:
 *    "n:{namePart}|{date}"  — name word + checkin (±2 วัน)
 *    "cr:{date}|{roomNum}"  — checkin + room number (±2 วัน)
 * ============================================================ */
function makeMatchKeys_(guest, checkin, room) {
  const parts = allNameParts_(guest);   // all words ≥3 chars
  const rn    = roomNum_(room);         // 3-digit room number
  const ci    = String(checkin || '').trim().substring(0, 10);
  const dates = ciDates_(ci);           // ci ±2 days

  const keys = [];
  parts.forEach(p => {
    dates.forEach(dt => keys.push('n:' + p + '|' + dt));
  });
  if (rn) {
    dates.forEach(dt => keys.push('cr:' + dt + '|' + rn));
  }
  return keys;
}

function allNameParts_(raw) {
  raw = String(raw || '').trim();
  return raw.split(/[\s,\/\\]+/)
    .map(p => p.toLowerCase().replace(/[^a-z0-9ก-๙]/g, ''))
    .filter(p => p.length >= 3);
}

function roomNum_(room) {
  const m = String(room || '').match(/\b(\d{3})\b/);
  return m ? m[1] : String(room || '').replace(/[^0-9]/g, '').substring(0, 3);
}

function ciDates_(ci) {
  const dates = [ci];
  if (!ci || ci.length < 10) return dates;
  try {
    const d = new Date(ci + 'T00:00:00Z');
    for (let delta = -2; delta <= 2; delta++) {
      if (delta === 0) continue;
      const d2 = new Date(d.getTime() + delta * 86400000);
      dates.push(d2.toISOString().substring(0, 10));
    }
  } catch (e) {}
  return dates;
}

/* ============================================================
 *  Checkbox state mutation
 * ============================================================ */
function setBookingDone(resId, done) {
  const map = getProp_(PROP_KEY_BOOKING_DONE);
  if (done) map[resId] = true; else delete map[resId];
  setProp_(PROP_KEY_BOOKING_DONE, map);
  return true;
}

function setInvoiceDone(invoiceKey, done) {
  const map = getProp_(PROP_KEY_INVOICE_DONE);
  if (done) map[invoiceKey] = true; else delete map[invoiceKey];
  setProp_(PROP_KEY_INVOICE_DONE, map);
  return true;
}

/* ============================================================
 *  Helpers
 * ============================================================ */
function indexMap_(header, keys) {
  const map = {};
  keys.forEach(k => { map[k] = header.indexOf(k); });
  return map;
}

function formatCellDate_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') return formatDateYMD_(val);
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // ISO string with T (e.g. "2026-03-26T17:00:00.000Z") → use local Bangkok date
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) return formatDateYMD_(d);
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return formatDateYMD_(d);
  return s;
}

function formatDateYMD_(d) {
  return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
}

function normalizeCode_(s) {
  return String(s || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
}

/* ============================================================
 *  GitHub integration helpers
 * ============================================================ */
function setupGithubToken() {
  PropertiesService.getScriptProperties().setProperty('GITHUB_TOKEN', 'PASTE_YOUR_TOKEN_HERE');
  Logger.log('✅ GITHUB_TOKEN set');
}

function testGithubToken() {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  const res = UrlFetchApp.fetch('https://api.github.com/repos/theloftlivingspace-droid/loft-booking-invoice-todo', {
    headers: { Authorization: 'token ' + token }
  });
  Logger.log(res.getContentText().slice(0, 300));
}

function pushToGithub() {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('ไม่พบ GITHUB_TOKEN — ใส่ใน Script Properties ก่อน');

  const REPO = 'theloftlivingspace-droid/loft-booking-invoice-todo';
  const BRANCH = 'main';
  const API = 'https://api.github.com';
  const headers = {
    'Authorization': 'token ' + token,
    'Content-Type': 'application/json',
    'User-Agent': 'Apps-Script-Pusher'
  };

  function ghFetch(method, path, data) {
    const res = UrlFetchApp.fetch(API + path, {
      method, headers,
      payload: data ? JSON.stringify(data) : undefined,
      muteHttpExceptions: true
    });
    const json = JSON.parse(res.getContentText());
    if (res.getResponseCode() >= 400) throw new Error(path + ': ' + JSON.stringify(json).slice(0, 200));
    return json;
  }

  function makeBlob(content) {
    const encoded = Utilities.base64Encode(Utilities.newBlob(content, 'text/plain', 'f').getBytes());
    return ghFetch('post', '/repos/' + REPO + '/git/blobs', { content: encoded, encoding: 'base64' }).sha;
  }

  const scriptId = ScriptApp.getScriptId();
  const exportUrl = 'https://www.googleapis.com/drive/v3/files/' + scriptId + '/export?mimeType=application/vnd.google-apps.script%2Bjson';
  const exportRes = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (exportRes.getResponseCode() !== 200) throw new Error('Export failed: ' + exportRes.getContentText().slice(0, 200));

  const files = JSON.parse(exportRes.getContentText()).files || [];
  const treeItems = files
    .filter(f => ['server_js','html','json'].includes(f.type))
    .map(f => {
      const path = f.name === 'appsscript' ? 'appsscript.json' : f.name + (f.type === 'html' ? '.html' : f.type === 'json' ? '.json' : '.gs');
      const sha = makeBlob(f.source);
      Logger.log('📄 ' + path + ' → ' + sha.slice(0,8));
      return { path, mode: '100644', type: 'blob', sha };
    });

  const refRes = UrlFetchApp.fetch(API + '/repos/' + REPO + '/git/ref/heads/' + BRANCH, { headers });
  const latestSha = JSON.parse(refRes.getContentText()).object.sha;
  const commitRes = UrlFetchApp.fetch(API + '/repos/' + REPO + '/git/commits/' + latestSha, { headers });
  const baseTree = JSON.parse(commitRes.getContentText()).tree.sha;

  const newTreeRes = UrlFetchApp.fetch(API + '/repos/' + REPO + '/git/trees', {
    method: 'post', headers,
    payload: JSON.stringify({ base_tree: baseTree, tree: treeItems })
  });
  const newTreeSha = JSON.parse(newTreeRes.getContentText()).sha;

  const now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm');
  const newCommitRes = UrlFetchApp.fetch(API + '/repos/' + REPO + '/git/commits', {
    method: 'post', headers,
    payload: JSON.stringify({ message: 'Auto-push from Apps Script ' + now, tree: newTreeSha, parents: [latestSha] })
  });
  const newCommitSha = JSON.parse(newCommitRes.getContentText()).sha;

  UrlFetchApp.fetch(API + '/repos/' + REPO + '/git/refs/heads/' + BRANCH, {
    method: 'patch', headers,
    payload: JSON.stringify({ sha: newCommitSha, force: false })
  });

  Logger.log('✅ Pushed ' + treeItems.length + ' files to GitHub: ' + newCommitSha);
  return newCommitSha;
}

function getProp_(key) {
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  return raw ? JSON.parse(raw) : {};
}

function setProp_(key, obj) {
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(obj));
}
