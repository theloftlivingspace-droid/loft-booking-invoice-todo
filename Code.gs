/**
 * ============================================================
 *  The Loft Living Space — Booking & Invoice To-Do Webapp
 * ------------------------------------------------------------
 *  Standalone Apps Script project (แยกจาก payout-income-log)
 *
 *  - "Booking To Add"   : booking ทั้งหมดจาก Sheet1 ที่ admin
 *                          ต้องเพิ่มใน Apartmentery
 *  - "Invoice To Create": payout ที่โอนแล้วจาก Payout_Income_Log
 *                          ที่ admin ต้องออกใบแจ้งหนี้ใน Apartmentery
 *
 *  - อ่านข้อมูลจาก SOURCE_SHEET_ID แบบ read-only เท่านั้น
 *  - เขียนเฉพาะ checkbox state ลง PropertiesService ของ project นี้
 *    ไม่แตะ sheet ต้นทางเลย
 *
 *  Setup:
 *   1) clasp create --type webapp --title "The Loft - Booking Invoice ToDo"
 *   2) ลบ Code.gs ที่ clasp สร้างมา แทนด้วยไฟล์นี้
 *   3) เพิ่มไฟล์ Index.html (ดูไฟล์คู่กัน)
 *   4) clasp push
 *   5) Deploy > New deployment > Web app
 *      - Execute as: Me
 *      - Who has access: Anyone with the link
 * ============================================================
 */

// ===== CONFIG =====
const SOURCE_SHEET_ID = '1XbTJLhecql_HNqyE80Hc6h30A2_elIxliudF4e6Rlz0';

const SRC_BOOKING_SHEET = 'Sheet1';
const SRC_PAYOUT_SHEET = 'Payout_Income_Log';

// สถานะใน Payout_Income_Log ที่นับว่า "โอนแล้ว" และต้องออกใบแจ้งหนี้
// ใช้ prefix match แทน exact match เพื่อรองรับทุก OTA
const PAYOUT_STATUS_PREFIXES = [
  '✅ Matched',          // Airbnb payout, Trip.com settlement, Expedia remittance ทุกแบบ
  'โอนแล้ว',            // โอนแล้ว / โอนแล้ว (Resolution Payout)
  'PrePaid',            // Booking.com prepaid
];

// คีย์ที่ใช้เก็บ checkbox state + snapshot ใน Script Properties
const PROP_KEY_BOOKING_DONE = 'booking_done_v1';      // { resId: true }
const PROP_KEY_INVOICE_DONE = 'invoice_done_v1';      // { invoiceKey: true }
const PROP_KEY_BOOKING_SEEN = 'booking_seen_v1';      // { resId: 'yyyy-MM-dd' (first seen) }
const PROP_KEY_INVOICE_SEEN = 'invoice_seen_v1';      // { invoiceKey: 'yyyy-MM-dd' (first seen) }

/* ============================================================
 *  Web app entry point
 * ============================================================ */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('The Loft — Booking & Invoice To-Do')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ============================================================
 *  Data loading (called from client via google.script.run)
 * ============================================================ */

/**
 * คืนข้อมูลทั้งสองแท็บ พร้อม flag "ใหม่วันนี้" / "ตรวจพบวันนี้"
 * และสถานะ checkbox ที่บันทึกไว้
 */
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
    const resId = String(r[idx.ResId] || '').trim();

    let firstSeen = seenMap[resId];
    let isNewToday = false;
    if (!firstSeen) {
      firstSeen = todayStr;
      seenMap[resId] = todayStr;
      seenChanged = true;
      isNewToday = true;
    }

    const guest = r[idx['ชื่อแขก']] || '';
    const checkin = formatCellDate_(r[idx['เช็คอิน']]);
    const room = String(r[idx['เลขห้อง']] || '').trim();

    return {
      resId: resId,
      room: room,
      guest: guest,
      checkin: checkin,
      checkout: formatCellDate_(r[idx['เช็คเอาท์']]),
      channel: r[idx.Channel] || '',
      note: r[idx.Note] || '',
      firstSeen: firstSeen,
      isNewToday: isNewToday,
      done: !!doneMap[resId],
      // matching keys (multi-strategy)
      confCode: normalizeCode_(resId),            // unlikely to match Payout HM codes
      firstCheckinKey: firstNameCheckinKey_(guest, checkin),  // firstName + checkin (best cross-sheet key)
      checkinRoomKey: checkinRoomKey_(checkin, room),         // checkin + room (fallback)
    };
  });

  if (seenChanged) setProp_(PROP_KEY_BOOKING_SEEN, seenMap);

  out.reverse();
  return out;
}

function getInvoiceToCreate_(ss, todayStr) {
  const src = ss.getSheetByName(SRC_PAYOUT_SHEET);
  if (!src) throw new Error('ไม่พบชีต: ' + SRC_PAYOUT_SHEET);

  const data = src.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1).filter(r => r.join('').trim() !== '');

  // หา column index แบบ flexible — รองรับทั้ง header เก่าและใหม่
  const idx = indexMap_(header, [
    'วันที่ตรวจพบ', 'OTA', 'Booking ID', 'Conf. Code', 'ชื่อแขก', 'ห้อง',
    'เช็คอิน', 'เช็คเอาท์', 'คืน', 'ยอดรวม (THB)', 'Commission (THB)', 'NET (THB)', 'สถานะ', 'หมายเหตุ',
  ]);
  // fallback column names (v22 และก่อนหน้า)
  if (idx['ห้อง'] < 0) idx['ห้อง'] = header.findIndex(h => /ห้อง|เลขห้อง/i.test(String(h)));
  if (idx['คืน'] < 0)  idx['คืน']  = header.findIndex(h => /คืน|จำนวนคืน/i.test(String(h)));
  if (idx['NET (THB)'] < 0) idx['NET (THB)'] = header.findIndex(h => /NET/i.test(String(h)));

  const isPayableStatus = s => PAYOUT_STATUS_PREFIXES.some(p => s.startsWith(p));

  // Logic:
  // - row "✅ Matched" = SCB batch row → เอา (แยก sub-entries จาก notes)
  // - row "โอนแล้ว" ที่ Conf Code ไม่มี SCB Matched → เอา
  // - row "โอนแล้ว" ที่ Conf Code มี SCB Matched อยู่แล้ว → ข้ามไป (ซ้ำ)
  // - row "↳" (sub-row SCB) → ข้ามไปเสมอ

  // เก็บ Conf Codes ที่มี SCB Matched row แล้ว (ป้องกัน duplicate)
  const matchedConfCodes = new Set();
  rows.forEach(r => {
    const status = String(r[idx['สถานะ']] || '').trim();
    if (status.includes('Matched')) {
      String(r[idx['Conf. Code']] || '').split(',').forEach(c => {
        const t = c.trim(); if (t) matchedConfCodes.add(t);
      });
    }
  });

  const filtered = rows.filter(r => {
    const status  = String(r[idx['สถานะ']]  || '').trim();
    const confCode= String(r[idx['Conf. Code']] || '').trim();
    const note    = String(r[idx['หมายเหตุ']] || '').trim();
    // ข้าม sub-rows (↳) เสมอ
    if (note.startsWith('↳')) return false;
    // ต้องเป็น status ที่ "จ่ายแล้ว"
    if (!isPayableStatus(status)) return false;
    // ถ้าเป็น row ปกติ (ไม่ใช่ Matched) แต่ conf นี้มี Matched row อยู่แล้ว → ข้าม
    if (!status.includes('Matched') && confCode && matchedConfCodes.has(confCode)) return false;
    return true;
  });

  const doneMap = getProp_(PROP_KEY_INVOICE_DONE);
  const seenMap = getProp_(PROP_KEY_INVOICE_SEEN);
  let seenChanged = false;

  const out = [];

  filtered.forEach(r => {
    const bookingId = String(r[idx['Booking ID']] || '').trim();
    const detectedDate = formatCellDate_(r[idx['วันที่ตรวจพบ']]);
    const room = r[idx.ห้อง] || '';
    const checkin = formatCellDate_(r[idx.เช็คอิน]);
    const checkout = formatCellDate_(r[idx.เช็คเอาท์]);
    const nights = r[idx.คืน] || '';
    const ota = r[idx.OTA] || '';
    const status = r[idx.สถานะ] || '';
    const totalNet = r[idx['NET (THB)']] || '';
    const rawConfCode = String(r[idx['Conf. Code']] || '');
    const rawGuestField = String(r[idx.ชื่อแขก] || '');

    // Parse sub-entries จาก หมายเหตุ
    // รองรับ 2 format:
    //   format A (✅ total row): "✅ Airbnb payout | GuestName(CONFCODE) NET ฿1234.56 | ..."
    //   format B (↳ sub-row ของ SCB): "↳ GuestName(CONFCODE) NET ฿1234.56"
    const notes = String(r[idx['หมายเหตุ']] || '');
    // pattern: ชื่อแขก ตามด้วย (CONFCODE) NET ฿amount
    // หลีกเลี่ยง Unicode char ใน character class — ใช้ negated pipe/paren/newline แทน
    const subPattern = /([^|()\n]+?)\(([^)]*)\)\s*NET\s+[\u0E3F]([\d,]+\.?\d*)/g;
    const subs = [];
    let m;
    while ((m = subPattern.exec(notes)) !== null) {
      const guestClean = m[1].trim().replace(/^[\s|]+/, '').trim();
      if (!guestClean) continue;
      subs.push({
        guest: guestClean,
        confCode: normalizeCode_(m[2]),
        net: parseFloat(m[3].replace(/,/g, '')),
      });
    }
    // fallback: ถ้า notes มี ฿ แต่ regex ไม่จับ → ลอง split by | แล้ว parse ทีละ segment
    if (!subs.length && notes.includes('฿')) {
      notes.split('|').forEach(seg => {
        seg = seg.trim();
        const sm = seg.match(/([^(]+)\(([^)]*)\)\s*NET\s+฿([\d,]+\.?\d*)/);
        if (sm) {
          const g = sm[1].trim().replace(/^[✅↳\s]+/, '').trim();
          if (g) subs.push({ guest: g, confCode: normalizeCode_(sm[2]), net: parseFloat(sm[3].replace(/,/g,'')) });
        }
      });
    }

    // กรณี guest มีหลายชื่อซ้ำ (merged row) ให้เอาชื่อแรกเป็น default
    const firstGuest = rawGuestField.split(',')[0].trim();
    const firstConfCode = rawConfCode.split(',')[0].trim();

    // ถ้ามีหลาย sub-entry (>1) → แยกเป็นหลาย card ทีละรายการ
    // ถ้ามี sub-entry เดียวหรือไม่มี → ใช้ค่าจาก row หลัก (1 card)
    const entries = subs.length > 1 ? subs : [{
      guest: firstGuest,
      confCode: normalizeCode_(firstConfCode),
      net: totalNet,
    }];

    entries.forEach((entry, i) => {
      // invoiceKey ต้อง unique ต่อ entry (สำหรับ checkbox done/seen state)
      const invoiceKey = entries.length > 1
        ? bookingId + '#' + (entry.confCode || i)
        : bookingId;

      let firstSeen = seenMap[invoiceKey];
      let isNewSeen = false;
      if (!firstSeen) {
        firstSeen = todayStr;
        seenMap[invoiceKey] = todayStr;
        seenChanged = true;
        isNewSeen = true;
      }

      out.push({
        invoiceKey: invoiceKey,
        bookingId: bookingId,
        room: room,
        guest: entry.guest || firstGuest,
        checkin: checkin,
        checkout: checkout,
        nights: nights,
        net: entries.length > 1 ? entry.net : totalNet,
        isSplitFromMulti: entries.length > 1,
        splitIndex: entries.length > 1 ? (i + 1) : null,
        splitTotal: entries.length > 1 ? entries.length : null,
        groupNet: entries.length > 1 ? totalNet : null, // ยอดรวมของ batch เดิม
        ota: ota,
        status: status,
        detectedDate: detectedDate,
        detectedToday: detectedDate === todayStr,
        firstSeen: firstSeen,
        isNewInList: isNewSeen,
        done: !!doneMap[invoiceKey],
        // matching keys (multi-strategy)
        confCode: entry.confCode || normalizeCode_(firstConfCode),
        firstCheckinKey: firstNameCheckinKey_(entry.guest || firstGuest, checkin),
        checkinRoomKey: checkinRoomKey_(checkin, room),
      });
    });
  });

  if (seenChanged) setProp_(PROP_KEY_INVOICE_SEEN, seenMap);

  // เรียงตามวันที่เงินเข้า (detectedDate) ใหม่สุดบนสุด
  out.sort((a, b) => {
    if (a.detectedDate !== b.detectedDate) return a.detectedDate < b.detectedDate ? 1 : -1;
    return a.guest < b.guest ? -1 : 1;
  });
  return out;
}

/* ============================================================
 *  Checkbox state mutation (only touches this script's own
 *  Properties store — never the source spreadsheet)
 * ============================================================ */

function setBookingDone(resId, done) {
  const map = getProp_(PROP_KEY_BOOKING_DONE);
  if (done) {
    map[resId] = true;
  } else {
    delete map[resId];
  }
  setProp_(PROP_KEY_BOOKING_DONE, map);
  return true;
}

function setInvoiceDone(invoiceKey, done) {
  const map = getProp_(PROP_KEY_INVOICE_DONE);
  if (done) {
    map[invoiceKey] = true;
  } else {
    delete map[invoiceKey];
  }
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
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return formatDateYMD_(val);
  }
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  if (!isNaN(d.getTime())) return formatDateYMD_(d);
  return s;
}

function formatDateYMD_(d) {
  return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
}

// ตัด whitespace / ตัวอักษรพิเศษ และทำเป็น uppercase สำหรับเทียบ ResId / Conf. Code
function normalizeCode_(s) {
  return String(s || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
}

// firstName + checkin: ดึงชื่อแรก (before space/comma) lowercase ไม่มีอักขระพิเศษ
// ใช้จับคู่ข้ามชีตได้ดีที่สุด เช่น "Nihel Ben Naceur" → "nihel" + "2026-05-03"
function firstNameCheckinKey_(guest, checkin) {
  const raw = String(guest || '').trim();
  // รองรับทั้ง "Last, First" และ "First Last"
  const parts = raw.split(/[\s,]+/);
  // ถ้ามี comma → format "Last, First" → เอา parts[1] (First), ไม่งั้นเอา parts[0]
  const firstName = (raw.includes(',') && parts.length > 1 ? parts[1] : parts[0]) || '';
  const fn = firstName.toLowerCase().replace(/[^a-z0-9ก-๙]/g, '');
  return fn + '|' + String(checkin || '').trim();
}

// checkin + roomNumber: fallback เผื่อชื่อพิมพ์ต่างกัน
function checkinRoomKey_(checkin, room) {
  const r = String(room || '').trim().replace(/\s+/g, '').toLowerCase();
  return String(checkin || '').trim() + '|' + r;
}

/* ============================================================
 *  GitHub integration helpers
 * ============================================================ */

/**
 * รันครั้งเดียวจาก Apps Script editor เพื่อ set GITHUB_TOKEN
 * ไม่ต้อง deploy — รันจาก ▶ Run ใน editor ได้เลย
 */
function setupGithubToken() {
  PropertiesService.getScriptProperties().setProperty(
    'GITHUB_TOKEN',
    'PASTE_YOUR_TOKEN_HERE'
  );
  Logger.log('✅ GITHUB_TOKEN set');
}

/**
 * ทดสอบว่า token ใช้ได้ไหม — รันจาก editor
 */
function testGithubToken() {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  const res = UrlFetchApp.fetch('https://api.github.com/repos/theloftlivingspace-droid/loft-booking-invoice-todo', {
    headers: { Authorization: 'token ' + token }
  });
  Logger.log(res.getContentText().slice(0, 300));
}

/**
 * Push ไฟล์ทั้งหมดจาก Apps Script project นี้ขึ้น GitHub
 * รันจาก Apps Script editor ▶ Run ได้เลย (ไม่ต้อง deploy)
 */
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
      method: method,
      headers: headers,
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

  // Export project as JSON via Drive API
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
      return { path: path, mode: '100644', type: 'blob', sha: sha };
    });

  // Get latest commit + tree
  const refRes = UrlFetchApp.fetch(API + '/repos/' + REPO + '/git/ref/heads/' + BRANCH, { headers: headers });
  const latestSha = JSON.parse(refRes.getContentText()).object.sha;
  const commitRes = UrlFetchApp.fetch(API + '/repos/' + REPO + '/git/commits/' + latestSha, { headers: headers });
  const baseTree = JSON.parse(commitRes.getContentText()).tree.sha;

  // Create new tree
  const newTreeRes = UrlFetchApp.fetch(API + '/repos/' + REPO + '/git/trees', {
    method: 'post',
    headers: headers,
    payload: JSON.stringify({ base_tree: baseTree, tree: treeItems })
  });
  const newTreeSha = JSON.parse(newTreeRes.getContentText()).sha;

  // Create commit
  const now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm');
  const newCommitRes = UrlFetchApp.fetch(API + '/repos/' + REPO + '/git/commits', {
    method: 'post',
    headers: headers,
    payload: JSON.stringify({
      message: 'Auto-push from Apps Script ' + now,
      tree: newTreeSha,
      parents: [latestSha]
    })
  });
  const newCommitSha = JSON.parse(newCommitRes.getContentText()).sha;

  // Update ref
  UrlFetchApp.fetch(API + '/repos/' + REPO + '/git/refs/heads/' + BRANCH, {
    method: 'patch',
    headers: headers,
    payload: JSON.stringify({ sha: newCommitSha, force: false })
  });

  Logger.log('✅ Pushed ' + treeItems.length + ' files to GitHub: ' + newCommitSha);
  return newCommitSha;
}

// Script Properties helpers — ใช้ PropertiesService ของ project นี้
// (ไม่เกี่ยวกับ source spreadsheet เลย)
function getProp_(key) {
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  return raw ? JSON.parse(raw) : {};
}

function setProp_(key, obj) {
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(obj));
}
