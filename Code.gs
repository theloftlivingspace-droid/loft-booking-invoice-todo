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
  '✅ Matched - Direct/Extranet',
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
 *  GET ?action=setNote&id=X&note=TEXT         → JSON (write Note col in Sheet1)
 *  GET (no action)            → serve HTML webapp
 * ============================================================ */
function doGet(e) {
  try {
    return doGet_(e);
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err), stack: err && err.stack ? String(err.stack).substring(0, 500) : '' });
  }
}

/* ============================================================
 *  POST entry point — used for binary/file uploads
 *  POST body (JSON): { action: 'uploadDoc', room, checkin, resId, fileName, mimeType, base64Data }
 *  POST body (JSON): { action: 'deleteDoc', fileId }
 *  GET  ?action=getAllDocs    → JSON map of resId -> DocFile[]
 * ============================================================ */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'uploadDoc') {
      return jsonResponse_(uploadDoc_(body));
    }
    if (action === 'deleteDoc') {
      return jsonResponse_(deleteDoc_(body));
    }
    if (action === 'markCheckedIn') {
      return jsonResponse_(markCheckedIn_(body));
    }
    if (action === 'earlyCheckout') {
      return jsonResponse_(earlyCheckout_(body));
    }
    if (action === 'updateCheckout') {
      return jsonResponse_(updateCheckoutDate_(body));
    }

    return jsonResponse_({ ok: false, error: 'Unknown POST action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err), stack: err && err.stack ? String(err.stack).substring(0, 500) : '' });
  }
}

// "Loft Documents" root folder — contains ONE SUBFOLDER PER BOOKING, named
// exactly "{roomNum}_{checkin}_{resId}" (matching folderKey() in CheckInOut.tsx).
// This is the folder structure that was actually in use all along; the flat
// single-folder + separate 'Docs' sheet scheme below it (in git history) was
// a wrong re-implementation that never saw any of the pre-existing documents.
const DOCS_ROOT_FOLDER_ID = '1fc3X-hmf1tUyCxTAbG6HxAN0qzrNyj4H'; // "Loft Documents"
const DOCS_SHEET_NAME = 'Docs'; // kept only as a supplementary upload log, not the source of truth

function getDocsRootFolder_() {
  return DriveApp.getFolderById(DOCS_ROOT_FOLDER_ID);
}

function docsKey_(room, checkin, resId) {
  return String(room || '') + '_' + formatCellDate_(checkin) + '_' + (String(resId) || 'noid');
}

function getOrCreateBookingDocsFolder_(room, checkin, resId) {
  const root = getDocsRootFolder_();
  const key = docsKey_(room, checkin, resId);
  const existing = root.getFoldersByName(key);
  if (existing.hasNext()) return existing.next();
  return root.createFolder(key);
}

function fileToDocFile_(f) {
  const fileId = f.getId();
  return {
    fileId: fileId,
    fileName: f.getName(),
    mimeType: f.getMimeType(),
    url: 'https://drive.google.com/file/d/' + fileId + '/view',
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + fileId,
    previewUrl: 'https://drive.google.com/file/d/' + fileId + '/preview',
    uploadedAt: f.getDateCreated().toISOString(),
  };
}

function getOrCreateDocsSheet_() {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  let sheet = ss.getSheetByName(DOCS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DOCS_SHEET_NAME);
    sheet.appendRow(['ResId', 'Room', 'Checkin', 'FileId', 'FileName', 'MimeType', 'UploadedAt']);
  }
  return sheet;
}

function uploadDoc_(body) {
  const room = body.room || '';
  const checkin = body.checkin || '';
  const resId = body.resId || '';
  const fileName = body.fileName || 'document';
  const mimeType = body.mimeType || 'application/octet-stream';
  const base64Data = body.base64Data || '';

  if (!resId || !base64Data) {
    return { ok: false, error: 'Missing resId or base64Data' };
  }

  let blob;
  try {
    const decoded = Utilities.base64Decode(base64Data);
    blob = Utilities.newBlob(decoded, mimeType, fileName);
  } catch (err) {
    return { ok: false, error: 'Invalid base64 data: ' + String(err) };
  }

  const folder = getOrCreateBookingDocsFolder_(room, checkin, resId);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  const uploadedAt = new Date().toISOString();

  // supplementary log only — not read back by getAllDocs_
  try {
    getOrCreateDocsSheet_().appendRow([resId, room, checkin, fileId, fileName, mimeType, uploadedAt]);
  } catch (e) { /* non-fatal */ }

  return {
    ok: true,
    fileId: fileId,
    fileName: fileName,
    mimeType: mimeType,
    url: 'https://drive.google.com/file/d/' + fileId + '/view',
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + fileId,
    previewUrl: 'https://drive.google.com/file/d/' + fileId + '/preview',
    uploadedAt: uploadedAt,
  };
}

function deleteDoc_(body) {
  const fileId = body.fileId || '';
  if (!fileId) return { ok: false, error: 'Missing fileId' };

  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (err) {
    return { ok: false, error: 'Could not delete file: ' + String(err) };
  }

  try {
    const sheet = getOrCreateDocsSheet_();
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][3] === fileId) { sheet.deleteRow(i + 1); break; }
    }
  } catch (e) { /* non-fatal */ }

  return { ok: true };
}

function getAllDocs_() {
  const root = getDocsRootFolder_();
  const subfolders = root.getFolders();
  const docs = {}; // "{room}_{checkin}_{resId}" -> DocFile[]  (matches folderKey() in CheckInOut.tsx)

  while (subfolders.hasNext()) {
    const sub = subfolders.next();
    const key = sub.getName();
    const files = sub.getFiles();
    const list = [];
    while (files.hasNext()) {
      list.push(fileToDocFile_(files.next()));
    }
    if (list.length) docs[key] = list;
  }

  return { ok: true, docs: docs };
}

/* ============================================================
 *  Check-in / Check-out persisted status
 * ------------------------------------------------------------
 *  Previously the frontend only tracked this in browser localStorage,
 *  so a check-in marked on the staff's phone never showed up on the
 *  admin's own device — nothing was ever written to a shared sheet.
 *  This sheet is now the single source of truth, read by getRoomStatus_().
 * ============================================================ */
const STATUS_SHEET_NAME = 'CheckStatus';

function getOrCreateStatusSheet_() {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  let sheet = ss.getSheetByName(STATUS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(STATUS_SHEET_NAME);
    sheet.appendRow(['ResId', 'CheckedInAt', 'CheckedOutAt', 'IsEarlyCheckout', 'NewCheckoutDate']);
  }
  return sheet;
}

function findStatusRow_(sheet, resId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === resId) return i + 1; // 1-indexed sheet row
  }
  return -1;
}

function markCheckedIn_(body) {
  const resId = String(body.resId || '');
  if (!resId) return { ok: false, error: 'Missing resId' };

  const sheet = getOrCreateStatusSheet_();
  const row = findStatusRow_(sheet, resId);
  const now = new Date().toISOString();
  if (row === -1) {
    sheet.appendRow([resId, now, '', '', '']);
  } else {
    sheet.getRange(row, 2).setValue(now);
  }
  return { ok: true, resId: resId, checkedInAt: now };
}

function earlyCheckout_(body) {
  const resId = String(body.resId || '');
  if (!resId) return { ok: false, error: 'Missing resId' };

  const now = new Date().toISOString();
  const isEarly = !!body.isEarly;
  const todayBKK = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  const newCheckout = String(body.newCheckout || todayBKK);

  // 1) Log to the CheckStatus sheet (existing behavior — per-device/audit trail)
  const sheet = getOrCreateStatusSheet_();
  const row = findStatusRow_(sheet, resId);
  if (row === -1) {
    sheet.appendRow([resId, '', now, isEarly ? 'TRUE' : 'FALSE', newCheckout]);
  } else {
    sheet.getRange(row, 3).setValue(now);
    sheet.getRange(row, 4).setValue(isEarly ? 'TRUE' : 'FALSE');
    sheet.getRange(row, 5).setValue(newCheckout);
  }

  // 2) Also update the real checkout date in Sheet1 (Loft_Reservations_Master)
  //    so occupancy/availability reflects the checkout immediately, not just
  //    the CheckStatus log.
  var sheet1Updated = false;
  try {
    const ss  = SpreadsheetApp.openById(SOURCE_SHEET_ID);
    const src = ss.getSheetByName(SRC_BOOKING_SHEET);
    if (src) {
      const data   = src.getDataRange().getValues();
      const header = data[0];
      const idx    = indexMap_(header, ['ResId', 'เลขห้อง', 'เช็คเอาท์', 'ชื่อแขก', 'เช็คอิน']);
      if (idx.ResId >= 0 && idx['เช็คเอาท์'] >= 0) {
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][idx.ResId] || '').trim() === resId) {
            src.getRange(i + 1, idx['เช็คเอาท์'] + 1).setValue(newCheckout);
            sheet1Updated = true;

            // แจ้งกลุ่มแม่บ้านผ่าน LINE bot ว่ามี checkout ก่อนกำหนด
            try {
              var props    = PropertiesService.getScriptProperties();
              var botUrl   = props.getProperty('BOT_URL')   || 'https://hotel-line-bot.onrender.com';
              var adminTok = props.getProperty('ADMIN_TOKEN') || 'apt2025@secret';
              var roomNum  = idx['เลขห้อง'] >= 0 ? String(data[i][idx['เลขห้อง']] || '').trim() : '';
              var guest    = idx['ชื่อแขก']  >= 0 ? String(data[i][idx['ชื่อแขก']]  || '').trim() : '';
              var checkin  = idx['เช็คอิน']  >= 0 ? String(data[i][idx['เช็คอิน']]  || '').trim() : '';
              UrlFetchApp.fetch(botUrl + '/api/checkout-notify', {
                method: 'post',
                contentType: 'application/json',
                payload: JSON.stringify({ room: roomNum, guest: guest, checkin: checkin, checkout: newCheckout }),
                headers: { 'x-admin-token': adminTok },
                muteHttpExceptions: true
              });
            } catch (e) { Logger.log('checkout-notify LINE error: ' + e); }

            break;
          }
        }
      }
    }
    if (sheet1Updated) triggerStyleSheet1_();
  } catch (e) {
    Logger.log('earlyCheckout_ Sheet1 update error: ' + e);
  }

  return { ok: true, resId: resId, checkedOutAt: now, newCheckout: newCheckout, sheet1Updated: sheet1Updated };
}

/**
 * Manual "edit checkout date" action — for cases where Little Hotelier
 * changes a guest's checkout date (e.g. an extended stay) but does NOT
 * send a modification email, so the Gmail-parsing pipeline never sees it.
 * Nathan enters the new checkout date by hand in the-loft-admin, which
 * calls this to:
 *   1) update the เช็คเอาท์ cell in Sheet1 — the source of truth that
 *      every other view (room-status grid, KPI cards, CheckInOut) reads.
 *   2) if the booking already has an Apartmentery bookingId, push the new
 *      end date there too, reusing updateApartmenteryBookingEndDateForRoom
 *      — the same "set end date" call the same-day-turnover buffer logic
 *      already uses to shrink bookings; it works equally well to extend.
 *   3) refuse the change if it would overlap another booking already on
 *      the same room — that's a real conflict Nathan has to resolve with
 *      the guest/OTA first, not something to silently paper over.
 * LINE notification to the maid group is fired from the client after a
 * successful save (same pattern as setNote / saveNote in CheckInOut.tsx),
 * not from here.
 */
function updateCheckoutDate_(body) {
  const resId = String(body.resId || '').trim();
  const newCheckout = String(body.newCheckout || '').trim();
  if (!resId) return { ok: false, error: 'resId required' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(newCheckout)) return { ok: false, error: 'newCheckout must be YYYY-MM-DD' };

  const ss  = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  if (!src) return { ok: false, error: 'Sheet1 not found' };

  const data   = src.getDataRange().getValues();
  const header = data[0];
  const idx    = indexMap_(header, ['ResId', 'เลขห้อง', 'เช็คเอาท์', 'ชื่อแขก', 'เช็คอิน']);
  if (idx.ResId < 0 || idx['เช็คเอาท์'] < 0) return { ok: false, error: 'required columns not found' };

  var rowIdx = -1, room = '', guest = '', checkin = '', oldCheckout = '';
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idx.ResId] || '').trim() === resId) {
      rowIdx = i;
      room = String(data[i][idx['เลขห้อง']] || '').trim();
      guest = idx['ชื่อแขก'] >= 0 ? String(data[i][idx['ชื่อแขก']] || '').trim() : '';
      checkin = idx['เช็คอิน'] >= 0 ? formatCellDate_(data[i][idx['เช็คอิน']]) : '';
      oldCheckout = formatCellDate_(data[i][idx['เช็คเอาท์']]);
      break;
    }
  }
  if (rowIdx === -1) return { ok: false, error: 'resId not found: ' + resId };
  if (/ยกเลิก|cancel/i.test(room)) return { ok: false, error: 'booking is cancelled' };
  if (newCheckout === oldCheckout) return { ok: false, error: 'newCheckout is the same as current checkout' };

  // Conflict check — only matters when extending: another booking on the
  // same room could already start somewhere inside [oldCheckout, newCheckout).
  if (newCheckout > oldCheckout) {
    const rn = roomNum_(room);
    for (var j = 1; j < data.length; j++) {
      if (j === rowIdx) continue;
      const otherRoom = String(data[j][idx['เลขห้อง']] || '').trim();
      if (/ยกเลิก|cancel/i.test(otherRoom)) continue;
      if (roomNum_(otherRoom) !== rn) continue;
      const otherCheckin = idx['เช็คอิน'] >= 0 ? formatCellDate_(data[j][idx['เช็คอิน']]) : '';
      if (!otherCheckin) continue;
      if (otherCheckin >= oldCheckout && otherCheckin < newCheckout) {
        return {
          ok: false,
          error: 'conflict',
          conflict: {
            resId: String(data[j][idx.ResId] || '').trim(),
            guest: idx['ชื่อแขก'] >= 0 ? String(data[j][idx['ชื่อแขก']] || '').trim() : '',
            checkin: otherCheckin,
          }
        };
      }
    }
  }

  src.getRange(rowIdx + 1, idx['เช็คเอาท์'] + 1).setValue(newCheckout);
  triggerStyleSheet1_();

  const result = {
    ok: true, resId: resId, room: room, guest: guest, checkin: checkin,
    oldCheckout: oldCheckout, newCheckout: newCheckout, apartmenterySynced: false
  };

  try {
    var aptId = getApartmenteryBookingId_(resId);
    if (!aptId) {
      // No bookingId recorded — could just be a booking the automation
      // hasn't created yet (next hourly run will pick it up with the new
      // date), OR one that was added to apartmentery manually before the
      // automation existed and so never got its ID written back to
      // Sheet1. Try to recover the latter case by matching guest name +
      // exact checkin date on the room's apartmentery calendar.
      var foundId = findApartmenteryBookingIdForRoomByGuest_(room, guest, checkin);
      if (foundId) {
        aptId = foundId;
        setApartmenteryBookingId_(resId, foundId);
        result.apartmenteryBackfilled = true;
      }
    }
    if (aptId) {
      const r = updateApartmenteryBookingEndDateForRoom(room, aptId, newCheckout);
      if (r && r.skipped) {
        result.apartmenteryNote = r.reason;
      } else {
        result.apartmenterySynced = true;
      }
    } else {
      result.apartmenteryNote = 'no apartmentery bookingId yet — nothing to sync';
    }
  } catch (e) {
    if (isApartmenterySessionExpiredError(e)) {
      result.apartmenteryNote = 'Apartmentery session expired — update the date there manually';
    } else {
      result.apartmenteryNote = 'Apartmentery sync failed: ' + e;
    }
  }

  return result;
}

function getCheckStatusMap_() {
  const sheet = getOrCreateStatusSheet_();
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const resId = String(data[i][0] || '');
    if (!resId) continue;
    map[resId] = {
      checkedInAt: data[i][1] || '',
      checkedOutAt: data[i][2] || '',
    };
  }
  return map;
}

/* ============================================================
 *  TEMP DEBUG — inspect the real "Loft Documents" root: how many
 *  per-booking subfolders exist, and whether any files are sitting
 *  loose directly in the root (stray uploads from before this fix).
 *  Safe/read-only. Can be removed later.
 * ============================================================ */
function debugScanDocsFolder_() {
  const root = getDocsRootFolder_();

  const subfolders = root.getFolders();
  const subfolderInfo = [];
  let totalFilesInSubfolders = 0;
  while (subfolders.hasNext()) {
    const sub = subfolders.next();
    const files = sub.getFiles();
    let count = 0;
    while (files.hasNext()) { files.next(); count++; }
    totalFilesInSubfolders += count;
    subfolderInfo.push({ name: sub.getName(), fileCount: count });
  }

  const rootFiles = root.getFiles();
  const strayFilesInRoot = [];
  while (rootFiles.hasNext()) {
    const f = rootFiles.next();
    strayFilesInRoot.push({
      fileId: f.getId(),
      fileName: f.getName(),
      createdAt: f.getDateCreated().toISOString(),
      sizeBytes: f.getSize(),
    });
  }

  return {
    ok: true,
    rootFolderId: root.getId(),
    rootFolderUrl: root.getUrl(),
    totalSubfolders: subfolderInfo.length,
    totalFilesInSubfolders: totalFilesInSubfolders,
    subfolders: subfolderInfo,
    strayFilesInRoot: strayFilesInRoot, // uploaded before the folder-structure fix — need migrating
  };
}

/* One-time cleanup: move any files sitting loose in the root folder into
 * their correct per-booking subfolder, using the Docs sheet log (which
 * still recorded resId/room/checkin for every upload) to know where each
 * stray file belongs. Safe to call multiple times — already-migrated
 * files won't be in the root anymore. */
function migrateStrayRootFiles_() {
  const root = getDocsRootFolder_();
  const sheet = getOrCreateDocsSheet_();
  const data = sheet.getDataRange().getValues();

  const infoByFileId = {};
  for (let i = 1; i < data.length; i++) {
    const [resId, room, checkin, fileId] = data[i];
    if (fileId) infoByFileId[fileId] = { resId, room, checkin };
  }

  const rootFiles = root.getFiles();
  const moved = [];
  const skipped = [];
  while (rootFiles.hasNext()) {
    const f = rootFiles.next();
    const fileId = f.getId();
    const info = infoByFileId[fileId];
    if (!info) { skipped.push({ fileId: fileId, fileName: f.getName(), reason: 'no matching Docs sheet row' }); continue; }

    const target = getOrCreateBookingDocsFolder_(info.room, info.checkin, info.resId);
    target.addFile(f);
    root.removeFile(f);
    moved.push({ fileId: fileId, fileName: f.getName(), movedTo: target.getName() });
  }

  return { ok: true, moved: moved, skipped: skipped };
}

function doGet_(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'getData') {
    return jsonResponse_(getDashboardData());
  }

  if (action === 'getRoomStatus') {
    return jsonResponse_(getRoomStatus_());
  }

  if (action === 'getAllDocs') {
    return jsonResponse_(getAllDocs_());
  }

  if (action === 'debugScanDocsFolder') {
    return jsonResponse_(debugScanDocsFolder_());
  }

  if (action === 'debugMigrateStrayFiles') {
    return jsonResponse_(migrateStrayRootFiles_());
  }

  if (action === 'setBookingDone') {
    const id   = e.parameter.id   || '';
    const done = e.parameter.done === 'true';
    setBookingDone(id, done);
    triggerStyleSheet1_();
    return jsonResponse_({ ok: true });
  }

  if (action === 'setInvoiceDone') {
    const id   = e.parameter.id   || '';
    const done = e.parameter.done === 'true';
    setInvoiceDone(id, done);
    triggerStyleSheet1_();
    return jsonResponse_({ ok: true });
  }

  if (action === 'setNote') {
    const id   = e.parameter.id   || '';
    const note = e.parameter.note || '';
    const result = setBookingNote(id, note);
    return jsonResponse_(result);
  }

  if (action === 'cancelBooking') {
    const id = e.parameter.id || '';
    const result = cancelBooking_(id);
    return jsonResponse_(result);
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
    pendingMatch: getPendingMatchPayouts_(ss),
  };
}

/* ============================================================
 *  Pending-match payouts — bank/OTA amounts already received but
 *  not yet reconciled to a booking (status not in
 *  PAYOUT_STATUSES_FOR_INVOICE and not marked ✅). Previously these
 *  only appeared in the Bank_Ledger sheet tab, invisible from the
 *  admin dashboard — money could arrive and go unnoticed unless the
 *  sheet was opened directly.
 * ============================================================ */
function getPendingMatchPayouts_(ss) {
  const src = ss.getSheetByName(SRC_PAYOUT_SHEET);
  if (!src) throw new Error('ไม่พบชีต: ' + SRC_PAYOUT_SHEET);

  const data = src.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1).filter(r => r.join('').trim() !== '');

  const idx = indexMap_(header, [
    'วันที่ตรวจพบ', 'OTA', 'Booking ID', 'Conf. Code', 'ชื่อแขก', 'ห้อง',
    'เช็คอิน', 'เช็คเอาท์', 'ยอดรวม (THB)', 'NET (THB)', 'สถานะ', 'หมายเหตุ',
  ]);

  const out = rows.filter(r => {
    const ota    = String(r[idx.OTA] || '').trim();
    const notes  = String(r[idx.หมายเหตุ] || '').trim();
    const status = String(r[idx.สถานะ] || '').trim();
    const bid    = String(r[idx['Booking ID']] || '').trim();
    if (!ota) return false;
    if (/^\d/.test(ota) || bid === 'THB') return false;   // summary/footer rows
    if (notes.startsWith('↳')) return false;               // matched sub-rows
    if (status.startsWith('✅')) return false;              // already matched
    if (PAYOUT_STATUSES_FOR_INVOICE.includes(status)) return false; // already matched
    return true;
  }).map(r => ({
    ota: String(r[idx.OTA] || ''),
    guest: String(r[idx.ชื่อแขก] || ''),
    room: String(r[idx.ห้อง] || ''),
    detectedDate: formatCellDate_(r[idx['วันที่ตรวจพบ']]),
    checkin: formatCellDate_(r[idx.เช็คอิน]),
    checkout: formatCellDate_(r[idx.เช็คเอาท์]),
    net: r[idx['NET (THB)']] || r[idx['ยอดรวม (THB)']] || '',
    status: String(r[idx.สถานะ] || '') || 'รอ match',
    note: String(r[idx.หมายเหตุ] || ''),
  }));

  out.sort((a, b) => (a.detectedDate < b.detectedDate ? 1 : -1));
  return out;
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

  // Build a lookup index of all bookings (name parts → list of {room, checkin}) from Sheet1.
  // Used to resolve the REAL room of each guest in a multi-room invoice payout, since the
  // "ห้อง" field in Payout_Income_Log only ever stores the combined room list (e.g. "108, 204, 300")
  // for every guest in the payout, with no indication of who stayed in which room.
  const bookingIndex_ = buildBookingLookupIndex_(ss);

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

  // Build subCiMap: conf code → {ci, co, nights} จาก single-conf rows ที่ถูก filter ออก
  // เพราะ summaryBookingIds (multi-conf total row มีอยู่แล้ว)
  // rows เหล่านี้มี ci/co จริงของแต่ละ guest — ต่างจาก total row ที่ ci/co เป็นของ guest แรก
  const subCiMap = {};
  rows.forEach(r => {
    const conf     = String(r[idx['Conf. Code']] || '').trim();
    const bookingId = String(r[idx['Booking ID']] || '').trim();
    // เฉพาะ single-conf rows ที่เป็นส่วนหนึ่งของ multi-guest payout
    if (!conf || conf.includes(',')) return;
    if (!summaryBookingIds.has(bookingId)) return;
    const ci  = formatCellDate_(r[idx.เช็คอิน]);
    const co  = formatCellDate_(r[idx.เช็คเอาท์]);
    const nts = r[idx.คืน] || '';
    if (ci) subCiMap[conf] = { ci, co, nights: nts };
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
    // รองรับเครื่องหมายลบทั้งก่อน฿ ("Adjustment -฿22.07") และหลัง฿ ("NET ฿-22.07"
    // ซึ่งเป็นรูปแบบจริงที่ parseAirbnbEmail เขียนลง sheet) — ไม่งั้น sub entry ที่เป็น
    // adjustment ลบจะหายไปจาก parsing เงียบๆ ทำให้ split ไม่ครบ — bug 2026-07-04 Nihel
    const subPattern = /([^|]+?)\(([^)]+)\)\s*(?:Adjustment\s+)?(?:NET\s+)?(-)?\s*฿(-)?([\d,]+\.?\d*)/g;
    const subs = [];
    let m;
    while ((m = subPattern.exec(notes)) !== null) {
      const _conf = m[2].trim();
      const _sub = subCiMap[_conf] || {};
      const _sign = (m[3] === '-' || m[4] === '-') ? -1 : 1;
      subs.push({ guest: m[1].trim(), confCode: _conf, net: _sign * parseFloat(m[5].replace(/,/g,'')), ci: _sub.ci || '', co: _sub.co || '', nights: _sub.nights || '' });
    }

    const entries = subs.length > 0 ? subs : [{ guest: firstGuest, confCode: firstConfCode, net: totalNet }];

    // เมื่อมีหลาย entries (multi-guest ใน 1 payout) และ room field มีหลายห้อง (comma-separated)
    // ห้ามเดาว่า room ตัวที่ i ตรงกับ entry ตัวที่ i ตามลำดับ — พบว่าลำดับใน room field
    // ไม่ได้สัมพันธ์กับลำดับ guest ใน notes เสมอไป (บางครั้งตรง บางครั้งสลับ)
    // วิธีที่แม่นยำกว่า: ค้นหา room จริงของ guest นั้นจาก booking sheet (Sheet1) โดยตรง
    // ด้วยชื่อ + checkin ใกล้เคียง (±3 วัน) ถ้าหาไม่เจอ fallback เป็น roomList ทั้งหมด
    // (กว้างกว่าเดิม แต่ยังดีกว่าเดาผิด)
    const roomList = room.split(',').map(r => r.trim()).filter(Boolean);
    function findRoomForGuest(guestName, entryCi) {
      // ใช้ checkin ของ entry นั้นๆ (จาก sub-row) ถ้ามี ไม่งั้นใช้ว่าง
      const found = lookupRoomFromIndex_(bookingIndex_, guestName, entryCi || '', roomList);
      if (found) return found;
      // หาไม่เจอใน Sheet1 (เช่น booking เก่าที่ถูกลบหลัง checkout) —
      // ห้ามคืน room string รวม (เช่น "363, 203") เพราะจะดู "ลิงค์ผิดห้อง"
      // ให้ flag ชัดเจนว่าไม่ทราบห้องแทน เพื่อให้ผู้ใช้ตรวจมือ
      return '⚠️ ไม่ทราบห้อง (' + room + ')';
    }

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
      const entryGuest = entry.guest || firstGuest;
      // ใช้ ci/co/nights ของ entry (จาก sub-row) ถ้ามี ไม่งั้น fallback total row
      const entryCheckin  = (entries.length > 1 && entry.ci)      ? entry.ci      : checkin;
      const entryCheckout = (entries.length > 1 && entry.co)      ? entry.co      : checkout;
      const entryNights   = (entries.length > 1 && entry.nights)  ? entry.nights  : nights;
      // Look up the real room whenever we have >1 possible room to disambiguate
      // (multi-guest payout, see comment above), OR whenever the parsed "ห้อง"
      // field itself is unusable (e.g. literal "?" — parser couldn't read it
      // off the OTA email at all). Previously this fallback only ran for
      // multi-guest split entries, so a single-guest row with room="?" had no
      // way to ever resolve — it just stayed "?" forever, even when a clear
      // name+date match existed in Sheet1 (bug 2026-07-06: 佰顺's non-refundable
      // cancellation payout, conf HMFTY4YTTK, stuck at "ห้อง ?" despite the
      // cancelled booking ABB-e4bdb0e9a1-20260705 being an exact name+date match).
      const roomNeedsLookup = (entries.length > 1 && roomList.length > 1) || !roomNum_(room);
      const entryRoom = roomNeedsLookup ? findRoomForGuest(entryGuest, entryCheckin) : room;
      out.push({
        invoiceKey, bookingId, room: entryRoom,
        guest: entryGuest,
        checkin: entryCheckin, checkout: entryCheckout, nights: entryNights,
        net: entries.length > 1 ? entry.net : totalNet,
        isSplitFromMulti: entries.length > 1,
        splitIndex: entries.length > 1 ? (i + 1) : null,
        splitTotal: entries.length > 1 ? entries.length : null,
        groupNet: entries.length > 1 ? totalNet : null,
        ota, status, detectedDate,
        detectedToday: detectedDate === todayStr,
        firstSeen, isNewInList: isNewSeen,
        done: !!doneMap[invoiceKey],
        matchKeys: makeMatchKeys_(entryGuest, entryCheckin, entryRoom),
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
 *  Room status — full Sheet1 dump for Check-in/Out PMS view
 *  (ไม่ filter/dedupe เหมือน getBookingToAdd_ เพราะ tab นี้ต้องเห็นทุก
 *   stay ที่ checked-in อยู่ หรือกำลังจะเข้า ไม่ใช่แค่ booking ที่ยังไม่ add invoice)
 * ============================================================ */
function getRoomStatus_() {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  if (!src) throw new Error('ไม่พบชีต: ' + SRC_BOOKING_SHEET);

  const data = src.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1).filter(r => r.join('').trim() !== '');
  const idx = indexMap_(header, ['เลขห้อง', 'ชื่อแขก', 'เช็คอิน', 'เช็คเอาท์', 'Channel', 'ResId', 'Note']);
  const statusMap = getCheckStatusMap_();

  const stays = rows.map(r => {
    const resId = String(r[idx.ResId] || '').trim();
    const st = statusMap[resId] || {};
    return {
      room:         String(r[idx['เลขห้อง']] || '').trim(),
      guest:        String(r[idx['ชื่อแขก']] || '').trim(),
      checkin:      formatCellDate_(r[idx['เช็คอิน']]),
      checkout:     formatCellDate_(r[idx['เช็คเอาท์']]),
      channel:      String(r[idx.Channel] || '').trim(),
      resId:        resId,
      note:         String(r[idx.Note] || '').trim(),
      checkedInAt:  st.checkedInAt || '',
      checkedOutAt: st.checkedOutAt || '',
    };
  }).filter(s => s.checkin && s.checkout);

  return { today: formatDateYMD_(new Date()), stays };
}

/* ============================================================
 *  Booking lookup index — resolves real room for multi-room invoices
 * ------------------------------------------------------------
 *  Payout_Income_Log stores one combined "ห้อง" field (e.g. "108, 204, 300")
 *  for ALL guests in a multi-room payout — there's no per-guest room data there.
 *  To resolve which guest stayed in which room, we look up the REAL booking
 *  record from Sheet1 by name + checkin proximity, restricted to rooms that
 *  are actually listed in this invoice's room field (extra safety net).
 * ============================================================ */
function buildBookingLookupIndex_(ss) {
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  if (!src) return {};
  const data = src.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1).filter(r => r.join('').trim() !== '');
  const idx = indexMap_(header, ['เลขห้อง', 'ชื่อแขก', 'เช็คอิน', 'เช็คเอาท์']);

  const index = {}; // namePart -> [{room, checkin}]
  rows.forEach(r => {
    const guest = String(r[idx['ชื่อแขก']] || '').trim();
    const room  = String(r[idx['เลขห้อง']] || '').trim();
    const rn    = roomNum_(room);
    const checkin = formatCellDate_(r[idx['เช็คอิน']]);
    if (!rn || !checkin) return;
    allNameParts_(guest).forEach(p => {
      if (!index[p]) index[p] = [];
      index[p].push({ room: rn, checkin });
    });
  });
  return index;
}

function lookupRoomFromIndex_(index, guestName, invoiceCheckin, allowedRoomList) {
  const parts = allNameParts_(guestName);
  const allowedNums = allowedRoomList.map(roomNum_).filter(Boolean);
  let best = null, bestDist = Infinity;

  parts.forEach(p => {
    const candidates = index[p] || [];
    candidates.forEach(c => {
      // Only consider rooms that are actually part of this invoice's room list —
      // never assign a room the invoice didn't even mention.
      if (allowedNums.length && allowedNums.indexOf(c.room) === -1) return;
      // ถ้า invoiceCheckin ว่าง (multi-guest total row) → match by name+room only
      if (!invoiceCheckin) {
        if (best === null) best = c.room;
        return;
      }
      const dist = Math.abs(daysDiff_(invoiceCheckin, c.checkin));
      if (dist <= 3 && dist < bestDist) {
        bestDist = dist;
        best = c.room;
      }
    });
  });
  return best;
}

function daysDiff_(a, b) {
  try {
    const da = new Date(a + 'T00:00:00Z');
    const db = new Date(b + 'T00:00:00Z');
    return (da.getTime() - db.getTime()) / 86400000;
  } catch (e) {
    return 999;
  }
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
    .map(p => p.toLowerCase().replace(/[^a-z0-9ก-๙\u4e00-\u9fff\u3400-\u4dbf]/g, ''))
    .filter(p => {
      if (!p) return false;
      const isCjk = /[\u4e00-\u9fff\u3400-\u4dbf]/.test(p);
      return isCjk ? p.length >= 2 : p.length >= 3;
    });
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
    for (let delta = -4; delta <= 4; delta++) {
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
  safeSetBookingDoneMap_(map);
  return true;
}

/**
 * booking_done_v1 is a single Script Properties value (9KB limit) that has
 * accumulated every resId ever marked done since the automation started,
 * with nothing ever pruned. Once it gets close to 9KB, setProperty() throws
 * and setBookingDone() was failing silently — the apartmentery booking gets
 * created (bookingId written to the sheet) but the row never gets marked
 * done, and it stays stuck forever because autoCreateApartmenteryBookings()
 * skips any resId that already has a bookingId before it would ever retry
 * setBookingDone(). Root-caused 2026-07-12 via the ABB-zhgggtr-20260712 row.
 *
 * Fix: if the write fails, prune the oldest entries (by the trailing
 * YYYYMMDD in the resId, when present) until it fits, then retry once.
 */
function safeSetBookingDoneMap_(map) {
  try {
    setProp_(PROP_KEY_BOOKING_DONE, map);
  } catch (err) {
    Logger.log(`safeSetBookingDoneMap_: setProp_ failed (${err.message}) — pruning booking_done_v1 and retrying`);
    pruneBookingDoneMap_(map);
    setProp_(PROP_KEY_BOOKING_DONE, map);
  }
}

function pruneBookingDoneMap_(map) {
  const SAFE_BYTES = 8000; // headroom under the ~9216 byte Script Properties limit
  const dated = Object.keys(map).map(key => {
    const m = key.match(/(\d{8})$/); // resIds mostly end in a YYYYMMDD date
    return { key: key, date: m ? m[1] : '00000000' }; // undated keys pruned first
  });
  dated.sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : 0); // oldest first
  let removed = 0;
  let i = 0;
  while (Utilities.newBlob(JSON.stringify(map)).getBytes().length > SAFE_BYTES && i < dated.length) {
    delete map[dated[i].key];
    removed++;
    i++;
  }
  Logger.log(`pruneBookingDoneMap_: removed ${removed} old entries from booking_done_v1`);
}

/**
 * Repair tool for rows already stuck by the bug above: any resId that has
 * an apartmentery bookingId written to Sheet1 (booking was actually created)
 * but is missing from booking_done_v1 (the done-write silently failed) gets
 * marked done now. Safe to re-run any time — run once from the Apps Script
 * editor after this deploy to fix already-stuck rows like ABB-zhgggtr-20260712.
 */
// Public wrapper — repairStuckDoneFlags_ ends in _ so Apps Script treats it
// as private and hides it from the editor's "Select function" Run dropdown.
// Run THIS one instead from the dropdown/Run button.
function runRepairStuckDoneFlags() {
  repairStuckDoneFlags_();
}

function repairStuckDoneFlags_() {
  const ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  if (!src) { Logger.log('repairStuckDoneFlags_: Sheet1 not found'); return; }
  const data = src.getDataRange().getValues();
  const header = data[0];
  const idx = indexMap_(header, ['ResId', APARTMENTERY_BOOKING_ID_COL_HEADER]);
  if (idx.ResId < 0 || idx[APARTMENTERY_BOOKING_ID_COL_HEADER] < 0) {
    Logger.log('repairStuckDoneFlags_: required columns not found');
    return;
  }
  const doneMap = getProp_(PROP_KEY_BOOKING_DONE);
  let fixed = 0;
  for (let i = 1; i < data.length; i++) {
    const resId = String(data[i][idx.ResId] || '').trim();
    const bookingId = String(data[i][idx[APARTMENTERY_BOOKING_ID_COL_HEADER]] || '').trim();
    if (resId && bookingId && !doneMap[resId]) {
      doneMap[resId] = true;
      fixed++;
      Logger.log(`repairStuckDoneFlags_: marking ${resId} done (has apartmentery bookingId ${bookingId})`);
    }
  }
  if (fixed > 0) safeSetBookingDoneMap_(doneMap);
  Logger.log(`repairStuckDoneFlags_: fixed ${fixed} stuck row(s)`);
}

function setInvoiceDone(invoiceKey, done) {
  const map = getProp_(PROP_KEY_INVOICE_DONE);
  if (done) map[invoiceKey] = true; else delete map[invoiceKey];
  safeSetInvoiceDoneMap_(map);
  return true;
}

/**
 * invoice_done_v1 is subject to the exact same 9KB Script Properties limit
 * as booking_done_v1 (see safeSetBookingDoneMap_ above), but this map is
 * written from inside autoCreateApartmenteryInvoicesAndReceipts() AFTER
 * processPayoutToReceiptForRoom() has already created the real invoice +
 * receipt in Apartmentery. If setProp_ throws here, the exception is caught
 * by the outer try/catch as a generic error — the invoiceKey never gets
 * marked done, so the next automation run will see it as still pending and
 * call processPayoutToReceiptForRoom() again, creating a DUPLICATE invoice.
 * This is worse than the booking_done_v1 case (which only stalls the
 * checkbox) — added 2026-07-15 after room 204 / Eaint Phoo Htet (booking
 * 326414) was found stuck in "ค้างสร้าง invoice".
 */
function safeSetInvoiceDoneMap_(map) {
  try {
    setProp_(PROP_KEY_INVOICE_DONE, map);
  } catch (err) {
    Logger.log(`safeSetInvoiceDoneMap_: setProp_ failed (${err.message}) — pruning invoice_done_v1 and retrying`);
    pruneInvoiceDoneMap_(map);
    setProp_(PROP_KEY_INVOICE_DONE, map);
  }
}

function pruneInvoiceDoneMap_(map) {
  const SAFE_BYTES = 8000; // headroom under the ~9216 byte Script Properties limit
  const dated = Object.keys(map).map(key => {
    // invoiceKey is usually a bare apartmentery bookingId (numeric) or
    // "bookingId#confCode(#n)" for split multi-guest payouts — neither
    // carries a reliable date, so fall back to numeric bookingId order
    // (older Apartmentery bookingIds are numerically smaller) as a proxy
    // for "oldest first", same intent as pruneBookingDoneMap_'s date sort.
    const bidPart = String(key).split('#')[0];
    const n = parseInt(bidPart, 10);
    return { key: key, sortKey: isNaN(n) ? 0 : n };
  });
  dated.sort((a, b) => a.sortKey - b.sortKey);
  let removed = 0;
  let i = 0;
  while (Utilities.newBlob(JSON.stringify(map)).getBytes().length > SAFE_BYTES && i < dated.length) {
    delete map[dated[i].key];
    removed++;
    i++;
  }
  Logger.log(`pruneInvoiceDoneMap_: removed ${removed} old entries from invoice_done_v1`);
}

/**
 * Repair tool mirroring runRepairStuckDoneFlags(), for invoice_done_v1.
 * Marks any invoiceKey done if a real invoice already exists for it in
 * Apartmentery (avoids re-marking things that were never actually created).
 * Run this manually once from the Apps Script editor if invoices are
 * suspected stuck due to the bug above — check Apartmentery directly first
 * to confirm which bookingIds already have an invoice before running.
 */
function runRepairStuckInvoiceDoneFlags(knownCreatedInvoiceKeys) {
  const doneMap = getProp_(PROP_KEY_INVOICE_DONE);
  let fixed = 0;
  (knownCreatedInvoiceKeys || []).forEach(key => {
    if (!doneMap[key]) { doneMap[key] = true; fixed++; }
  });
  if (fixed > 0) safeSetInvoiceDoneMap_(doneMap);
  Logger.log(`runRepairStuckInvoiceDoneFlags: fixed ${fixed} stuck invoice key(s)`);
}

function setBookingNote(resId, note) {
  if (!resId) return { ok: false, error: 'resId required' };
  const ss  = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  if (!src) return { ok: false, error: 'Sheet1 not found' };
  const data   = src.getDataRange().getValues();
  const header = data[0];
  const idx    = indexMap_(header, ['ResId', 'Note']);
  if (idx.ResId < 0) return { ok: false, error: 'ResId column not found' };
  if (idx.Note  < 0) return { ok: false, error: 'Note column not found' };
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idx.ResId] || '').trim() === resId) {
      src.getRange(i + 1, idx.Note + 1).setValue(note);
      triggerStyleSheet1_();
      return { ok: true };
    }
  }
  return { ok: false, error: 'resId not found: ' + resId };
}

function cancelBooking_(resId) {
  if (!resId) return { ok: false, error: 'resId required' };
  const ss  = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  const src = ss.getSheetByName(SRC_BOOKING_SHEET);
  if (!src) return { ok: false, error: 'Sheet1 not found' };
  const data   = src.getDataRange().getValues();
  const header = data[0];
  const idx    = indexMap_(header, ['ResId', 'เลขห้อง', 'เช็คเอาท์', 'ชื่อแขก', 'เช็คอิน']);
  if (idx.ResId < 0) return { ok: false, error: 'ResId column not found' };
  if (idx['เลขห้อง'] < 0) return { ok: false, error: 'เลขห้อง column not found' };
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idx.ResId] || '').trim() === resId) {
      const currentRoom = String(data[i][idx['เลขห้อง']] || '').trim();
      // ถ้า mark ยกเลิกแล้วให้ข้าม (idempotent)
      if (/ยกเลิก|cancel/i.test(currentRoom)) return { ok: true, alreadyCancelled: true };
      const newRoom = currentRoom + ' ยกเลิก';
      // เปลี่ยน checkout เป็นวันนี้ (Bangkok) เพื่อปลดล็อคห้องให้จองใหม่ได้ทันที
      var todayBKK = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
      src.getRange(i + 1, idx['เลขห้อง'] + 1).setValue(newRoom);
      if (idx['เช็คเอาท์'] >= 0) {
        src.getRange(i + 1, idx['เช็คเอาท์'] + 1).setValue(todayBKK);
      }
      triggerStyleSheet1_();
      // แจ้งกลุ่มแม่บ้านผ่าน LINE bot
      try {
        var props     = PropertiesService.getScriptProperties();
        var botUrl    = props.getProperty('BOT_URL')   || 'https://hotel-line-bot.onrender.com';
        var adminTok  = props.getProperty('ADMIN_TOKEN') || 'apt2025@secret';
        var guest     = String(data[i][idx['ชื่อแขก']] || '').trim();
        var checkin   = idx['เช็คอิน']  >= 0 ? String(data[i][idx['เช็คอิน']]  || '').trim() : '';
        UrlFetchApp.fetch(botUrl + '/api/cancel-notify', {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify({ room: currentRoom, guest: guest, checkin: checkin, checkout: todayBKK }),
          headers: { 'x-admin-token': adminTok },
          muteHttpExceptions: true
        });
      } catch(e) { Logger.log('LINE notify error: ' + e); }
      return { ok: true, room: newRoom, checkoutUpdated: todayBKK };
    }
  }
  return { ok: false, error: 'resId not found: ' + resId };
}

/* ============================================================
 *  Trigger styleSheet1 on payout-income-log GAS (fire-and-forget)
 * ============================================================ */
function triggerStyleSheet1_() {
  try {
    var PAYOUT_GAS_URL = 'https://script.google.com/macros/s/AKfycbyAP9Z_pIlKrXv9AOXwDhY0wNVSSFL0vU8VuH0SssFyxretRyt9CJNjxVZOLN3eFjs/exec';
    UrlFetchApp.fetch(PAYOUT_GAS_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ action: 'styleSheet1' }),
      muteHttpExceptions: true,
      followRedirects: true,
    });
  } catch (e) {
    Logger.log('triggerStyleSheet1_ error (non-fatal): ' + e);
  }
}

/* ============================================================
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
// setupGithubToken() ถูกลบออกแล้ว — ฟังก์ชันนี้เขียนทับ GITHUB_TOKEN ด้วย placeholder
// string ทุกครั้งที่ถูกรัน (โดยตั้งใจหรือไม่ตั้งใจ) ทำให้ token จริงที่ตั้งไว้หายไป
// ตั้งค่า GITHUB_TOKEN ผ่าน Project Settings → Script Properties โดยตรงเท่านั้น

function testGithubToken() {
  const token = PropertiesService.getScriptProperties().getProperty('GH_TOKEN_V2');
  Logger.log('Token length: ' + (token ? token.length : 'null'));
  Logger.log('Token first 10: ' + (token ? token.substring(0,10) : 'null'));
  Logger.log('Token last 5: ' + (token ? token.substring(token.length-5) : 'null'));
  Logger.log('Has whitespace: ' + (token ? /\s/.test(token) : 'null'));
  const res = UrlFetchApp.fetch('https://api.github.com/repos/theloftlivingspace-droid/loft-booking-invoice-todo', {
    headers: { Authorization: 'token ' + token },
    muteHttpExceptions: true
  });
  Logger.log('Status: ' + res.getResponseCode());
  Logger.log(res.getContentText().slice(0, 300));
}

function pushToGithub() {
  const token = PropertiesService.getScriptProperties().getProperty('GH_TOKEN_V2');
  if (!token) throw new Error('ไม่พบ GH_TOKEN_V2 — ใส่ใน Script Properties ก่อน');

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

/* ============================================================
 *  pullFromGithub — ดึง Code.gs + Index.html จาก GitHub
 *  แล้ว overwrite ไฟล์ใน Apps Script project นี้ทันที
 *
 *  รันจาก Apps Script editor ▶ Run > pullFromGithub
 *  ไม่ต้อง deploy — เห็นผลทันทีหลัง reload web app
 * ============================================================ */
function pullFromGithub() {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('ไม่พบ GITHUB_TOKEN — รัน setupGithubToken() ก่อน');

  const REPO   = 'theloftlivingspace-droid/loft-booking-invoice-todo';
  const BRANCH = 'main';
  const API    = 'https://api.github.com';
  const ghHeaders = { Authorization: 'token ' + token, 'User-Agent': 'Apps-Script-Puller' };

  function ghGet(path) {
    const res = UrlFetchApp.fetch(API + path, { headers: ghHeaders, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('GitHub GET ' + path + ' → ' + res.getResponseCode());
    return JSON.parse(res.getContentText());
  }

  // ดึง file list จาก tree ล่าสุด
  const ref    = ghGet('/repos/' + REPO + '/git/ref/heads/' + BRANCH);
  const commit = ghGet('/repos/' + REPO + '/git/commits/' + ref.object.sha);
  const tree   = ghGet('/repos/' + REPO + '/git/trees/' + commit.tree.sha);

  // Export project เพื่อดู file list ปัจจุบัน
  const scriptId = ScriptApp.getScriptId();
  const exportUrl = 'https://www.googleapis.com/drive/v3/files/' + scriptId +
    '/export?mimeType=application/vnd.google-apps.script%2Bjson';
  const exportRes = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (exportRes.getResponseCode() !== 200)
    throw new Error('Export failed: ' + exportRes.getContentText().slice(0, 200));

  const project = JSON.parse(exportRes.getContentText());
  const files   = project.files || [];

  // Map GitHub tree → { filename: content }
  const ghFiles = {};
  tree.tree.forEach(function(item) {
    if (item.type !== 'blob') return;
    // map: Code.gs → server_js "Code", Index.html → html "Index"
    const name = item.path;
    if (name === 'Code.gs' || name === 'Index.html') {
      const blob = ghGet('/repos/' + REPO + '/git/blobs/' + item.sha);
      ghFiles[name] = Utilities.newBlob(
        Utilities.base64Decode(blob.content.replace(/[\s]/g, '')),
        'text/plain'
      ).getDataAsString();
    }
  });

  Logger.log('📥 ดึงจาก GitHub: ' + Object.keys(ghFiles).join(', '));

  // Overwrite files ใน project
  files.forEach(function(f) {
    if (f.type === 'server_js' && f.name === 'Code' && ghFiles['Code.gs']) {
      f.source = ghFiles['Code.gs'];
      Logger.log('✅ Updated: Code.gs (' + f.source.length + ' chars)');
    }
    if (f.type === 'html' && f.name === 'Index' && ghFiles['Index.html']) {
      f.source = ghFiles['Index.html'];
      Logger.log('✅ Updated: Index.html (' + f.source.length + ' chars)');
    }
  });

  // Push กลับเข้า Apps Script via Drive API
  const importUrl = 'https://www.googleapis.com/upload/drive/v3/files/' + scriptId +
    '?uploadType=media&mimeType=application/vnd.google-apps.script%2Bjson';
  const importRes = UrlFetchApp.fetch(importUrl, {
    method: 'patch',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      'Content-Type': 'application/vnd.google-apps.script+json'
    },
    payload: JSON.stringify(project),
    muteHttpExceptions: true
  });

  if (importRes.getResponseCode() !== 200)
    throw new Error('Import failed: ' + importRes.getContentText().slice(0, 300));

  Logger.log('🎉 pullFromGithub สำเร็จ! Reload web app เพื่อเห็นผล');
  return 'OK';
}
