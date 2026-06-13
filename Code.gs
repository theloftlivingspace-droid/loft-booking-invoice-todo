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
const PAYOUT_STATUSES_FOR_INVOICE = [
  'โอนแล้ว',
  'โอนแล้ว (Resolution Payout)',
];

// คีย์ที่ใช้เก็บ checkbox state + snapshot ใน Script Properties
const PROP_KEY_BOOKING_DONE = 'booking_done_v1';      // { resId: true }
const PROP_KEY_INVOICE_DONE = 'invoice_done_v1';      // { bookingId: true }
const PROP_KEY_BOOKING_SEEN = 'booking_seen_v1';      // { resId: 'yyyy-MM-dd' (first seen) }
const PROP_KEY_INVOICE_SEEN = 'invoice_seen_v1';      // { bookingId: 'yyyy-MM-dd' (first seen) }

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

    return {
      resId: resId,
      room: r[idx['เลขห้อง']] || '',
      guest: r[idx['ชื่อแขก']] || '',
      checkin: formatCellDate_(r[idx['เช็คอิน']]),
      checkout: formatCellDate_(r[idx['เช็คเอาท์']]),
      channel: r[idx.Channel] || '',
      note: r[idx.Note] || '',
      firstSeen: firstSeen,
      isNewToday: isNewToday,
      done: !!doneMap[resId],
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

  const idx = indexMap_(header, [
    'วันที่ตรวจพบ', 'OTA', 'Booking ID', 'Conf. Code', 'ชื่อแขก', 'ห้อง',
    'เช็คอิน', 'เช็คเอาท์', 'คืน', 'ยอดรวม (THB)', 'Commission (THB)', 'NET (THB)', 'สถานะ', 'หมายเหตุ',
  ]);

  // เอาเฉพาะ row ✅ Matched (row รวม) และ row โอนแล้ว ปกติ — ไม่เอา row ↳ ย่อย
  const filtered = rows.filter(r => {
    const status = String(r[idx.สถานะ] || '').trim();
    const note = String(r[idx.หมายเหตุ] || '').trim();
    if (note.startsWith('↳')) return false; // row ย่อย ข้ามไป
    return PAYOUT_STATUSES_FOR_INVOICE.indexOf(status) !== -1;
  });

  const doneMap = getProp_(PROP_KEY_INVOICE_DONE);
  const seenMap = getProp_(PROP_KEY_INVOICE_SEEN);
  let seenChanged = false;

  const out = filtered.map(r => {
    const bookingId = String(r[idx['Booking ID']] || '').trim();
    const detectedDate = formatCellDate_(r[idx['วันที่ตรวจพบ']]);

    let firstSeen = seenMap[bookingId];
    let isNewSeen = false;
    if (!firstSeen) {
      firstSeen = todayStr;
      seenMap[bookingId] = todayStr;
      seenChanged = true;
      isNewSeen = true;
    }

    // Parse sub-NETs from หมายเหตุ e.g. "✅ Airbnb payout | Nihel(HMCTA5TJ35) NET ฿81.17 | Nihel(HMCTA5TJ35) NET ฿2638.54"
    const notes = String(r[idx.หมายเหตุ] || '');
    const netMatches = notes.match(/NET\s+฿([\d,]+\.?\d*)/g) || [];
    const netSubs = netMatches.map(m => parseFloat(m.replace(/NET\s+฿/, '').replace(/,/g, '')));
    const totalNet = r[idx['NET (THB)']] || '';

    return {
      bookingId: bookingId,
      room: r[idx.ห้อง] || '',
      guest: r[idx.ชื่อแขก] || '',
      checkin: formatCellDate_(r[idx.เช็คอิน]),
      checkout: formatCellDate_(r[idx.เช็คเอาท์]),
      nights: r[idx.คืน] || '',
      net: totalNet,
      netSubs: netSubs,
      ota: r[idx.OTA] || '',
      status: r[idx.สถานะ] || '',
      detectedDate: detectedDate,
      detectedToday: detectedDate === todayStr,
      firstSeen: firstSeen,
      isNewInList: isNewSeen,
      done: !!doneMap[bookingId],
    };
  });

  if (seenChanged) setProp_(PROP_KEY_INVOICE_SEEN, seenMap);

  out.sort((a, b) => {
    if (a.firstSeen !== b.firstSeen) return a.firstSeen < b.firstSeen ? 1 : -1;
    return a.detectedDate < b.detectedDate ? 1 : -1;
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

function setInvoiceDone(bookingId, done) {
  const map = getProp_(PROP_KEY_INVOICE_DONE);
  if (done) {
    map[bookingId] = true;
  } else {
    delete map[bookingId];
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

// Script Properties helpers — ใช้ PropertiesService ของ project นี้
// (ไม่เกี่ยวกับ source spreadsheet เลย)
function getProp_(key) {
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  return raw ? JSON.parse(raw) : {};
}

function setProp_(key, obj) {
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(obj));
}
