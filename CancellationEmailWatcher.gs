/* ============================================================
 *  CancellationEmailWatcher.gs
 *
 *  ตรวจอีเมล cancellation จาก OTA (Airbnb / Booking.com / Expedia /
 *  Trip.com / Little Hotelier direct) ใน theloftlivingspace@gmail.com
 *  อัตโนมัติผ่าน time-driven trigger แล้ว:
 *    - match ได้ชัวร์ (findRoom_ คืนแถวเดียวไม่กำกวม) → เรียก cancelBooking_()
 *      ตรงๆ (ฟังก์ชันเดิมที่ปุ่ม ✕ ใน UI ใช้อยู่ — idempotent, sync ครบ
 *      Sheet1/Apartmentery/LINE maid-notify/payout matching)
 *    - match ไม่ชัวร์ / ไม่พบ → ส่ง LINE แจ้ง admin ให้ยืนยันมือ ไม่แตะ Sheet1
 *
 *  ตั้ง trigger ครั้งแรก: รัน installCancellationEmailTrigger() หนึ่งครั้ง
 *  จาก Apps Script editor (ต้องรันมือ เพราะ deploy จาก git ตั้ง trigger
 *  ให้อัตโนมัติไม่ได้)
 *
 *  Matching logic (findRoom_/normG_) ก๊อปมาจาก payout-income-log/Code.gs
 *  เพื่อใช้ตรรกะเดียวกับตัว matcher ที่ผ่านการแก้บั๊กมาแล้วหลายรอบ — ถ้าแก้
 *  ตัวต้นทางใน payout-income-log อย่าลืม sync มาที่นี่ด้วย
 * ============================================================ */

var CANCEL_WATCHER_LABEL = 'CancelBot/Processed';

// ── Setup: รันครั้งเดียวจาก editor ──────────────────────────────
function installCancellationEmailTrigger() {
  // ลบ trigger เก่าของฟังก์ชันนี้ก่อน กันซ้ำ
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'checkCancellationEmails_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('checkCancellationEmails_')
    .timeBased()
    .everyMinutes(30)
    .create();
  Logger.log('ตั้ง trigger checkCancellationEmails_ ทุก 30 นาทีแล้ว');
}

// ── Main entry (เรียกจาก time trigger) ──────────────────────────
function checkCancellationEmails_() {
  var label = GmailApp.getUserLabelByName(CANCEL_WATCHER_LABEL) ||
    GmailApp.createLabel(CANCEL_WATCHER_LABEL);

  var query = '(from:automated@airbnb.com OR from:no-reply@app.littlehotelier.com ' +
    'OR from:auto_reservation@trip.com OR from:noreply_htl@trip.com) ' +
    '-label:"' + CANCEL_WATCHER_LABEL + '" newer_than:14d';

  var threads = GmailApp.search(query, 0, 50);
  Logger.log('checkCancellationEmails_: พบ ' + threads.length + ' thread(s)');

  threads.forEach(function (thread) {
    thread.getMessages().forEach(function (msg) {
      try {
        var msgObj = {
          subject: msg.getSubject(),
          sender: msg.getFrom(),
          date: msg.getDate(),
          htmlBody: msg.getBody(),
          plaintextBody: msg.getPlainBody(),
        };
        var parsed = parseCancellationEmail_(msgObj);
        if (parsed) handleParsedCancellation_(parsed);
      } catch (e) {
        Logger.log('checkCancellationEmails_ error on message: ' + e);
      }
    });
    thread.addLabel(label);
  });
}

function handleParsedCancellation_(parsed) {
  Logger.log('parsed cancellation: ' + JSON.stringify(parsed));

  if (!parsed.isActualCancellation) {
    Logger.log('ไม่ใช่การยกเลิกจริง (เช่น Trip.com fee-waiver ไม่สำเร็จ) — ข้าม');
    return;
  }
  if (!parsed.guestName) {
    notifyAdminUncertainCancel_(parsed, 'ไม่พบชื่อแขกในอีเมล — ต้องเช็คมือ');
    return;
  }

  var byGuestAll = buildGuestIndexFromSheet1_();
  var ci = parsed.checkIn ? new Date(parsed.checkIn) : null;
  var match = findBookingEntry_(parsed.guestName, ci, byGuestAll);

  if (match && match.resId) {
    var result = cancelBooking_(match.resId);
    Logger.log('auto-cancel ผ่าน cancelBooking_(' + match.resId + '): ' + JSON.stringify(result));
    if (!result.ok) {
      notifyAdminUncertainCancel_(parsed, 'match ห้อง ' + match.room + ' ได้ แต่ cancelBooking_ error: ' + result.error);
    }
  } else {
    notifyAdminUncertainCancel_(parsed, 'match booking ไม่ชัวร์ (ชื่อ/วันที่ใกล้เคียงหลายแถว หรือไม่พบเลย)');
  }
}

function notifyAdminUncertainCancel_(parsed, reason) {
  var note = '⚠️ อีเมลยกเลิกจาก ' + parsed.ota + ' — ต้องเช็คมือ\n' +
    'แขก: ' + (parsed.guestName || '-') + '\n' +
    'เช็คอิน: ' + (parsed.checkIn || '-') + ' / เช็คเอาท์: ' + (parsed.checkOut || '-') + '\n' +
    'เลขอ้างอิง OTA: ' + (parsed.bookingRef || '-') + '\n' +
    'เหตุผล: ' + reason + '\n' +
    'Subject: ' + parsed.raw.subject;
  try {
    var props    = PropertiesService.getScriptProperties();
    var botUrl   = props.getProperty('BOT_URL')    || 'https://hotel-line-bot.onrender.com';
    var adminTok = props.getProperty('ADMIN_TOKEN') || 'apt2025@secret';
    UrlFetchApp.fetch(botUrl + '/api/send-admin-alert', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ note: note }),
      headers: { 'x-admin-token': adminTok },
      muteHttpExceptions: true,
    });
  } catch (e) {
    Logger.log('notifyAdminUncertainCancel_ LINE error: ' + e);
  }
}

/* ============================================================
 *  Sheet1 guest index + matcher
 *  (findRoom_/normG_ ก๊อปมาจาก payout-income-log/Code.gs — ปรับให้
 *  คืน {room, resId, ci} ทั้ง entry แทนที่จะคืนแค่เลขห้อง เพราะ
 *  cancelBooking_ ต้องใช้ resId ไม่ใช่เลขห้อง)
 * ============================================================ */

function buildGuestIndexFromSheet1_() {
  var ss  = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  var src = ss.getSheetByName(SRC_BOOKING_SHEET);
  var data = src.getDataRange().getValues();
  var header = data[0];
  var idx = indexMap_(header, ['ResId', 'เลขห้อง', 'ชื่อแขก', 'เช็คอิน']);

  var byGuestAll = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var roomRaw = String(row[idx['เลขห้อง']] || '').trim();
    var guest   = String(row[idx['ชื่อแขก']] || '').trim();
    var resId   = String(row[idx.ResId] || '').trim();
    if (!roomRaw || !guest || !resId) continue;
    var isCancelled = /ยกเลิก|cancel/i.test(roomRaw);
    if (isCancelled) continue; // อย่า match กับแถวที่ยกเลิกไปแล้ว (กันยิง cancelBooking_ ซ้ำผ่านทาง fuzzy match ผิดแถว)
    var gk = normG_(guest);
    var ci = row[idx['เช็คอิน']] ? new Date(row[idx['เช็คอิน']]) : null;
    if (!byGuestAll[gk]) byGuestAll[gk] = [];
    byGuestAll[gk].push({ room: roomRaw, resId: resId, ci: ci });
  }
  return byGuestAll;
}

function normG_(s) {
  var words = s.toString().toLowerCase()
    .replace(/[,\/\\]+/g, ' ')
    .replace(/\s+/g, ' ').trim()
    .split(' ');
  words.sort();
  return words.join(' ');
}

// เหมือน findRoom() ใน payout-income-log ทุกประการ ยกเว้นคืน entry
// ทั้งก้อน ({room, resId, ci}) แทนที่จะคืนแค่ room string
function findBookingEntry_(guestRaw, ci, byGuest) {
  var CI_WINDOW_EXACT = 3 * 86400000;
  var CI_WINDOW_FUZZY = 5 * 86400000;
  var CI_WINDOW_DATE  = 1 * 86400000;

  var gk = normG_(guestRaw);

  // 1. Exact normalized name match
  if (byGuest[gk]) {
    var cands = byGuest[gk];
    if (ci) {
      var dc = cands.filter(function (c) {
        return c.ci && Math.abs(ci.getTime() - c.ci.getTime()) <= CI_WINDOW_EXACT;
      });
      return dc.length ? dc[0] : null; // มี ci แต่ไม่มี candidate ตรงช่วง → ห้ามเดา
    }
    return cands.length === 1 ? cands[0] : null; // ไม่มี ci และชื่อซ้ำหลายแถว → ไม่ชัวร์
  }

  // 1.5 CJK / substring match
  if (/[\u3400-\u9FFF]/.test(gk)) {
    var gkNoSpace = gk.replace(/\s+/g, '');
    var cjkCands = [];
    Object.keys(byGuest).forEach(function (k) {
      var kNoSpace = k.replace(/\s+/g, '');
      if (kNoSpace && gkNoSpace && (kNoSpace.indexOf(gkNoSpace) >= 0 || gkNoSpace.indexOf(kNoSpace) >= 0)) {
        cjkCands = cjkCands.concat(byGuest[k]);
      }
    });
    if (cjkCands.length) {
      if (ci) {
        var dcCjk = cjkCands.filter(function (c) {
          return c.ci && Math.abs(ci.getTime() - c.ci.getTime()) <= CI_WINDOW_EXACT;
        });
        if (dcCjk.length) return dcCjk[0];
      } else if (cjkCands.length === 1) {
        return cjkCands[0];
      }
    }
  }

  // 2. Fuzzy word match
  var parts = gk.split(' ').filter(function (p) { return p.length > 2; });
  if (parts.length) {
    var best = null, bestScore = 0;
    Object.keys(byGuest).forEach(function (k) {
      var kWords = k.split(' ');
      var score = 0;
      parts.forEach(function (p) {
        if (kWords.indexOf(p) >= 0) score += p.length;
      });
      if (score > bestScore) { bestScore = score; best = k; }
    });
    var longPart = parts.some(function (p) { return p.length >= 5; });
    var minScore = longPart ? 5 : 8;
    if (bestScore >= minScore && best) {
      var cands2 = byGuest[best];
      if (ci) {
        var dc2 = cands2.filter(function (c) {
          return c.ci && Math.abs(ci.getTime() - c.ci.getTime()) <= CI_WINDOW_FUZZY;
        });
        return dc2.length ? dc2[0] : null;
      }
      return cands2.length === 1 ? cands2[0] : null;
    }
  }

  // หมายเหตุ: ต้นฉบับ (payout-income-log/findRoom) มี step 3 "date-only fallback"
  // ที่จับคู่ได้แม้ชื่อไม่ตรงกันเลย ขอแค่ check-in วันเดียวกันแบบ unique ทั้งระบบ —
  // เหมาะกับ payout matching (ผิดแล้วแก้ทีหลังได้) แต่เสี่ยงเกินไปสำหรับที่นี่ เพราะ
  // ผลคือยกเลิก booking จริงแบบ irreversible ถ้าชื่อไม่ match เลยตั้งแต่ step 1/1.5/2
  // ให้ถือว่า "ไม่ชัวร์" แล้วส่งแจ้ง admin แทนเสมอ ไม่ auto-cancel ด้วยวันที่อย่างเดียว
  return null;
}

/* ============================================================
 *  Email parser (ported จาก parseCancellationEmail.js — เนื้อหาเดียวกัน
 *  ทุกประการ แค่แปลง module.exports → GAS global function + stripHtml_
 *  ใช้ GAS string methods แทน Node)
 * ============================================================ */

function stripHtml_(html) {
  if (!html) return '';
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&#39;|&rsquo;/gi, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

function toIsoDate_(d, m, y) {
  var months = {
    jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06',
    jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12',
  };
  var mm = /^\d+$/.test(m) ? String(m).padStart(2, '0') : months[String(m).slice(0, 3).toLowerCase()];
  if (!mm) return null;
  return y + '-' + mm + '-' + String(d).padStart(2, '0');
}

function parseAirbnb_(message) {
  var subject = message.subject || '';
  var text = stripHtml_(message.htmlBody) || message.plaintextBody || '';

  var refMatch = subject.match(/Reservation\s+([A-Z0-9]{8,12})/i) || text.match(/reservation\s+([A-Z0-9]{8,12})/i);
  var nameMatch = text.match(/your guest\s+([A-Za-z' -]+?)\s+(?:had to cancel|canceled|cancelled)/i);

  var checkIn = null, checkOut = null;
  var rangeMatch = subject.match(/for\s+([A-Za-z]{3})\s+(\d{1,2})\s*[–-]\s*(?:([A-Za-z]{3})\s+)?(\d{1,2}),\s*(\d{4})/);
  if (rangeMatch) {
    checkIn = toIsoDate_(rangeMatch[2], rangeMatch[1], rangeMatch[5]);
    checkOut = toIsoDate_(rangeMatch[4], rangeMatch[3] || rangeMatch[1], rangeMatch[5]);
  }

  return {
    ota: 'airbnb',
    isActualCancellation: true,
    bookingRef: refMatch ? refMatch[1] : null,
    guestName: nameMatch ? nameMatch[1].trim() : null,
    checkIn: checkIn,
    checkOut: checkOut,
    cancellationFee: null,
  };
}

function parseLittleHotelierChannel_(message, ota) {
  var text = stripHtml_(message.htmlBody) || message.plaintextBody || '';

  var refMatch = text.match(/Booking Confirmation Id:\s*(\S+)/i);
  var nameMatch = text.match(/Reservation Cancellation\s+([A-Za-z' .-]+?)\s+The loft/i);
  var checkInMatch = text.match(/Check-in:\s*(\d{2})-([A-Za-z]{3})-(\d{4})/i);
  var checkOutMatch = text.match(/Check-out:\s*(\d{2})-([A-Za-z]{3})-(\d{4})/i);
  var feeMatch = text.match(/Cancellation Fee:\s*([\d,.]+)\s*THB/i);

  return {
    ota: ota,
    isActualCancellation: true,
    bookingRef: refMatch ? refMatch[1] : null,
    guestName: nameMatch ? nameMatch[1].trim() : null,
    checkIn: checkInMatch ? toIsoDate_(checkInMatch[1], checkInMatch[2], checkInMatch[3]) : null,
    checkOut: checkOutMatch ? toIsoDate_(checkOutMatch[1], checkOutMatch[2], checkOutMatch[3]) : null,
    cancellationFee: feeMatch ? parseFloat(feeMatch[1].replace(/,/g, '')) : null,
  };
}

function parseLittleHotelierDirect_(message) {
  var text = stripHtml_(message.htmlBody) || message.plaintextBody || '';
  var nameMatch = text.match(/Reservation Cancellation\s+([A-Za-z' .-]+?)\s+has cancelled/i);
  var rangeMatch = text.match(/for\s+(\d{1,2})\s+([A-Za-z]+)\s+to\s+(\d{1,2})\s+([A-Za-z]+)/i);

  var checkIn = null, checkOut = null;
  if (rangeMatch) {
    var emailYear = message.date ? new Date(message.date).getFullYear() : new Date().getFullYear();
    checkIn = toIsoDate_(rangeMatch[1], rangeMatch[2], emailYear);
    checkOut = toIsoDate_(rangeMatch[3], rangeMatch[4], emailYear);
  }

  return {
    ota: 'little_hotelier_direct',
    isActualCancellation: true,
    bookingRef: null,
    guestName: nameMatch ? nameMatch[1].trim() : null,
    checkIn: checkIn,
    checkOut: checkOut,
    cancellationFee: null,
  };
}

function parseTripCom_(message) {
  var subject = message.subject || '';
  var text = stripHtml_(message.htmlBody) || message.plaintextBody || '';

  var isActualCancellation = /cancellation request accepted/i.test(subject) && !/fee waiver/i.test(subject);

  var refMatch = subject.match(/booking no\.?\s*#?(\d+)#?/i) || text.match(/[Rr]eservation no\.\s*(\d+)/);
  var nameMatch = text.match(/Guest Name\s+([A-Z' .\/-]+?)\s+Booking Amount/i);
  var feeMatch = text.match(/[Aa]pplicable cancellation fees\s+THB\s*([\d,.]+)/i);
  var stayMatch = text.match(/Staying period\s+([A-Za-z]{3})\s+(\d{1,2}),\s*(\d{4})\s*-\s*([A-Za-z]{3})\s+(\d{1,2}),\s*(\d{4})/);

  var checkIn = null, checkOut = null;
  if (stayMatch) {
    checkIn = toIsoDate_(stayMatch[2], stayMatch[1], stayMatch[3]);
    checkOut = toIsoDate_(stayMatch[5], stayMatch[4], stayMatch[6]);
  }

  return {
    ota: 'trip.com',
    isActualCancellation: isActualCancellation,
    bookingRef: refMatch ? refMatch[1] : null,
    guestName: nameMatch ? nameMatch[1].trim() : null,
    checkIn: checkIn,
    checkOut: checkOut,
    cancellationFee: feeMatch ? parseFloat(feeMatch[1].replace(/,/g, '')) : null,
  };
}

function parseCancellationEmail_(message) {
  var sender = (message.sender || '').toLowerCase();
  var subject = message.subject || '';
  var parsed = null;

  if (sender.indexOf('automated@airbnb.com') >= 0 && /cancel/i.test(subject)) {
    parsed = parseAirbnb_(message);
  } else if (sender.indexOf('no-reply@app.littlehotelier.com') >= 0) {
    if (/booking\.com cancellation/i.test(subject)) {
      parsed = parseLittleHotelierChannel_(message, 'booking.com');
    } else if (/expedia cancellation/i.test(subject)) {
      parsed = parseLittleHotelierChannel_(message, 'expedia');
    } else if (/reservation cancellation/i.test(subject)) {
      parsed = parseLittleHotelierDirect_(message);
    }
  } else if (sender.indexOf('trip.com') >= 0) {
    if (/cancellation request accepted/i.test(subject) || /fee waiver request (failed|unsuccessful)/i.test(subject)) {
      parsed = parseTripCom_(message);
    }
  }

  if (!parsed) return null;

  return Object.assign(parsed, {
    raw: { subject: subject, sender: message.sender, date: message.date || null },
  });
}
