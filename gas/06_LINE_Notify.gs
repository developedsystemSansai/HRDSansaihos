// ============================================================
// === PART 2: LINE Messaging API — ระบบแจ้งเตือน HRD ===
// ============================================================
//
// วิธีใช้ครั้งแรก (ทำตามลำดับ):
//   STEP 1 → รัน LINE_setup()       ตรวจสอบ Token + Bot info
//   STEP 2 → User add Bot → Bot ตอบ User ID กลับใน LINE
//   STEP 3 → Admin ผูก User ID กับชื่อพนักงานในหน้าเว็บ (แท็บพนักงาน)
//   STEP 4 → รัน LINE_sendTest()    ทดสอบส่งจริง
//   STEP 5 → รัน LINE_setTrigger()  ตั้งส่งอัตโนมัติทุกวัน
//
// กรณีที่จะส่งข้อความ LINE:
//   1. อัตโนมัติ (Trigger) — ครบกำหนด LINE_AFTER_DAYS วันหลังอบรม + ยังไม่ส่งสรุป
//      → ส่งซ้ำทุกวันจนกว่าจะมีสรุป (ใช้ >= ไม่ใช่ == เพื่อไม่พลาด)
//   2. Admin กดส่งด้วยตนเอง จากหน้าเว็บ (sendDocReminderManual)
//   3. User add Bot → Bot ตอบ User ID กลับอัตโนมัติ (doPost Webhook)
//   4. ทดสอบ (LINE_sendTest) — รันจาก Script Editor
// ============================================================

// ── ค่าคงที่ LINE (ปรับได้ที่นี่) ─────────────────────────
var LINE_AFTER_DAYS   = 7;   // ส่งแจ้งเตือนเมื่อผ่านมา >= N วันหลังสิ้นสุดอบรม
var LINE_NOTIFY_HOUR  = 8;   // เวลาที่ Trigger ทำงาน (8 = 08:00 น.)
var LINE_TEST_USER_ID = 'U9174b59332c4a96bbcfd9a79fb1579c5'; // User ID สำหรับทดสอบ

// URL ของ Web App (ใส่ใน Flex Message)
var APP_URL = 'https://script.google.com/macros/s/AKfycbwtnfYjlEAloUtZt5cvblr_6hUAgaVx4hYzff86RmkjjDRgGFbIM_43jnJnFkIn4eq8eQ/exec';

// ─────────────────────────────────────────────────────────

// ============================================================
// === _getLineToken() — ดึง Channel Access Token จาก Sheet ===
// ============================================================
function _getLineToken() {
  try {
    var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    var cfg = ss.getSheetByName('ตั้งค่าระบบ');
    if (!cfg) {
      Logger.log('⚠️ _getLineToken: ไม่พบ Sheet "ตั้งค่าระบบ"');
      return '';
    }
    var rows = cfg.getDataRange().getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === 'LINE Channel Access Token') {
        var t = String(rows[i][1]).trim();
        if (!t) Logger.log('⚠️ _getLineToken: พบแถวแต่ค่าว่าง — กรุณาใส่ Token');
        return t;
      }
    }
    Logger.log('⚠️ _getLineToken: ไม่พบแถว "LINE Channel Access Token" ใน Sheet ตั้งค่าระบบ');
  } catch(e) {
    Logger.log('❌ _getLineToken error: ' + e);
  }
  return '';
}

// ============================================================
// === _linePush(userId, messageObj) — ส่ง Push Message ===
// คืน { ok:bool, code:int, body:string }
// ============================================================
function _linePush(userId, messageObj) {
  var token  = _getLineToken();
  var result = { ok: false, code: 0, body: '' };

  if (!token) {
    result.body = 'ไม่พบ Channel Access Token — กรุณาใส่ใน Sheet "ตั้งค่าระบบ"';
    Logger.log('❌ _linePush: ' + result.body);
    return result;
  }
  if (!userId || !/^U[0-9a-f]{32}$/i.test(userId)) {
    result.body = 'LINE User ID ไม่ถูกรูปแบบ: "' + userId + '"';
    Logger.log('❌ _linePush: ' + result.body);
    return result;
  }

  try {
    var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method            : 'post',
      contentType       : 'application/json',
      headers           : { 'Authorization': 'Bearer ' + token },
      payload           : JSON.stringify({ to: userId, messages: [messageObj] }),
      muteHttpExceptions: true
    });
    result.code = res.getResponseCode();
    result.body = res.getContentText();
    result.ok   = (result.code === 200);

    if (!result.ok) {
      Logger.log('❌ LINE Push HTTP ' + result.code + ': ' + result.body.substring(0, 200));
      if (result.code === 400) Logger.log('   → User ยังไม่ได้ add Bot หรือ User ID ผิด');
      if (result.code === 401) Logger.log('   → Token ผิดหรือหมดอายุ');
      if (result.code === 429) Logger.log('   → Rate limit — ลองใหม่ในอีกสักครู่');
    }
  } catch(e) {
    result.body = e.toString();
    Logger.log('❌ _linePush exception: ' + e);
  }
  return result;
}

// ============================================================
// === _lineReply(replyToken, messageObj) — ตอบกลับ Webhook ===
// ============================================================
function _lineReply(replyToken, messageObj) {
  var token = _getLineToken();
  if (!token || !replyToken) return;
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method            : 'post',
      contentType       : 'application/json',
      headers           : { 'Authorization': 'Bearer ' + token },
      payload           : JSON.stringify({ replyToken: replyToken, messages: [messageObj] }),
      muteHttpExceptions: true
    });
  } catch(e) {
    Logger.log('❌ _lineReply error: ' + e);
  }
}

// ============================================================
// === _findEmpLineId(name, empSheet) — ค้นหา LINE User ID ===
// คืน User ID ถ้าสถานะแจ้งเตือน = "เปิดใช้งาน" มิฉะนั้นคืน ''
// ============================================================
function _findEmpLineId(name, empSheet) {
  try {
    var d  = empSheet.getDataRange().getValues();
    var h  = d[0];
    var nc = h.indexOf('ชื่อ-สกุล');
    var lc = h.indexOf('LINE User ID');
    var sc = h.indexOf('สถานะแจ้งเตือน');
    if (nc < 0 || lc < 0) return '';
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][nc] || '').trim() === name) {
        var st = sc >= 0 ? String(d[i][sc] || '').trim() : 'เปิดใช้งาน';
        if (st !== 'เปิดใช้งาน') return ''; // ปิดแจ้งเตือน
        return String(d[i][lc] || '').trim();
      }
    }
  } catch(e) {
    Logger.log('❌ _findEmpLineId error: ' + e);
  }
  return '';
}

// ============================================================
// === _writeLineLog() — บันทึก Log การส่ง LINE ===
// ============================================================
function _writeLineLog(logSheet, type, name, topic, result) {
  if (!logSheet) return;
  try {
    var now = new Date();
    logSheet.appendRow([
      Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy'),
      Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm:ss'),
      type,
      name,
      (topic || '').substring(0, 100),
      result.ok ? 'ส่งแล้ว' : 'ส่งไม่สำเร็จ',
      result.code,
      result.ok ? '' : (result.body || '').substring(0, 200)
    ]);
  } catch(e) {
    Logger.log('⚠️ _writeLineLog error: ' + e);
  }
}

// ============================================================
// === doPost — Webhook รับ Event จาก LINE ===
// เมื่อ User add Bot → ตอบ User ID กลับทันที
// ============================================================
function doPost(e) {
  var out = ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
  try {
    var body   = JSON.parse(e.postData.contents);
    var events = body.events || [];
    events.forEach(function(ev) {
      if (!ev.source || !ev.source.userId) return;
      Logger.log('📥 Webhook: type=' + ev.type + ' userId=' + ev.source.userId);

      // ตอบเฉพาะ follow event (User add Bot ครั้งแรก)
      if (ev.type !== 'follow') return;

      if (ev.replyToken) {
        _lineReply(ev.replyToken, {
          type: 'text',
          text: '🏥 ระบบแจ้งเตือน HRD โรงพยาบาลสันทราย\n' +
                '━━━━━━━━━━━━━━━━━━━━\n' +
                '🆔 LINE User ID ของคุณ:\n' +
                ev.source.userId + '\n' +
                '━━━━━━━━━━━━━━━━━━━━\n' +
                '📝 กรุณาแจ้ง User ID นี้ให้เจ้าหน้าที่ HRD\n' +
                'เพื่อลงทะเบียนรับการแจ้งเตือนในระบบ\n' +
                '━━━━━━━━━━━━━━━━━━━━\n' +
                'ขอบคุณที่ใช้บริการ 🙏'
        });
        Logger.log('✅ ตอบ User ID กลับ follow event: ' + ev.source.userId);
      }
    });
  } catch(err) {
    Logger.log('❌ doPost error: ' + err);
  }
  return out;
}

// ============================================================
// === _buildDocReminderMsg() — Flex Message แจ้งส่งเอกสาร ===
// ใช้กับทั้ง Trigger อัตโนมัติ และ Admin กดส่งเอง
// ============================================================
function _buildDocReminderMsg(name, records) {
  // records = array ของ { topic, endDate, daysAfter }
  // ถ้าส่งมาแค่ 1 record ก็ยังใช้ได้

  var itemContents = [];
  records.forEach(function(r, idx) {
    if (idx > 0) itemContents.push({ type: 'separator', margin: 'sm' });
    itemContents.push({
      type: 'box', layout: 'vertical',
      paddingAll: '10px', backgroundColor: '#fffbeb', cornerRadius: '6px',
      contents: [
        {
          type: 'box', layout: 'baseline', spacing: 'sm',
          contents: [
            { type: 'text', text: (idx + 1) + '.', size: 'sm', color: '#d97706', weight: 'bold', flex: 0 },
            { type: 'text', text: r.topic || 'การอบรม/ประชุม', size: 'sm', wrap: true, color: '#1e293b', weight: 'bold', flex: 5 }
          ]
        },
        {
          type: 'text',
          text: '📅 สิ้นสุด ' + _fmtDateTh(r.endDate) + '  (ครบกำหนดมา ' + r.daysAfter + ' วัน)',
          size: 'xs', color: '#64748b', margin: 'xs', wrap: true
        }
      ]
    });
  });

  var countText = records.length > 1
    ? 'รายการที่ยังไม่ส่งสรุป (' + records.length + ' รายการ)'
    : 'รายการที่ยังไม่ส่งสรุป';

  return {
    type    : 'flex',
    altText : '⚠️ [HRD] ' + name + ' กรุณาส่งสรุปอบรม ' + records.length + ' รายการ',
    contents: {
      type: 'bubble', size: 'mega',
      header: {
        type: 'box', layout: 'vertical',
        backgroundColor: '#d97706', paddingAll: '18px',
        contents: [
          { type: 'text', text: '⚠️ แจ้งเตือนส่งสรุปอบรม',
            color: '#ffffff', size: 'lg', weight: 'bold' },
          { type: 'text', text: 'ครบ ' + LINE_AFTER_DAYS + ' วันหลังสิ้นสุดอบรม',
            color: '#fef3c7', size: 'sm', margin: 'sm' }
        ]
      },
      body: {
        type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '18px',
        contents: [
          {
            type: 'box', layout: 'baseline', spacing: 'sm',
            contents: [
              { type: 'text', text: '👤', size: 'sm', flex: 0 },
              { type: 'text', text: name, size: 'sm', weight: 'bold', wrap: true, flex: 5 }
            ]
          },
          { type: 'separator', margin: 'md' },
          { type: 'text', text: countText,
            size: 'sm', weight: 'bold', color: '#1e293b', margin: 'md' }
        ].concat(itemContents).concat([
          { type: 'separator', margin: 'md' },
          { type: 'text',
            text: '📎 กรอกสรุปผลได้ที่ลิ้งค์ด้านล่าง',
            size: 'xs', color: '#059669', wrap: true, weight: 'bold', margin: 'md' },
          { type: 'text', text: APP_URL,
            size: 'xxs', color: '#2563eb', wrap: true, decoration: 'underline',
            action: { type: 'uri', label: 'เปิดลิ้งค์', uri: APP_URL } }
        ])
      },
      footer: {
        type: 'box', layout: 'vertical', paddingAll: '12px',
        contents: [
          { type: 'text', text: 'กรุณาดำเนินการโดยเร็ว ขอบคุณค่ะ 🙏',
            size: 'xs', color: '#d97706', align: 'center' },
          { type: 'text',
            text: 'ส่งโดย HRD | ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy') + ' น.',
            size: 'xxs', color: '#9ca3af', align: 'center', margin: 'xs' }
        ]
      }
    }
  };
}

// ============================================================
// === dailyNotificationCheck() — Trigger รายวัน (รวมระบบเดียว) ===
//
// ตรวจสอบทั้ง 2 Sheet แล้วส่ง 1 ข้อความต่อคน (รวมทุกรายการค้าง)
// เงื่อนไข: daysAfter >= LINE_AFTER_DAYS AND Summary ว่าง AND มี LINE ID
// ============================================================
function dailyNotificationCheck() {
  Logger.log('══════════════════════════════════════════════');
  Logger.log('  dailyNotificationCheck START ' +
             Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm'));
  Logger.log('══════════════════════════════════════════════');

  var token = _getLineToken();
  if (!token) {
    Logger.log('❌ ไม่พบ Channel Access Token — หยุดทำงาน');
    return;
  }

  // ── ดึงรายการที่ครบ LINE_AFTER_DAYS วันพอดี ──────────────
  var result = getFollowupStatus(LINE_AFTER_DAYS);
  if (!result.success) {
    Logger.log('❌ getFollowupStatus: ' + result.message);
    return;
  }

  var pending = result.pending;
  if (pending.length === 0) {
    Logger.log('✅ ไม่มีรายการที่ครบ ' + LINE_AFTER_DAYS + ' วันวันนี้ — ไม่ส่ง LINE');
    return;
  }

  Logger.log('📋 พบรายการที่ครบ 7 วัน: ' + pending.length + ' รายการ จาก ' + _countUniqueNames(pending) + ' คน');

  var ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  var empSh  = ss.getSheetByName('ข้อมูลพนักงาน');
  var logSh  = ss.getSheetByName('Log_การแจ้งเตือน');
  var sent = 0, skipped = 0;

  // จัดกลุ่มตามชื่อ เพื่อส่ง 1 ข้อความต่อคน
  var byName = {};
  pending.forEach(function(r) {
    if (!byName[r.name]) byName[r.name] = [];
    byName[r.name].push(r);
  });

  Object.keys(byName).forEach(function(name) {
    var records = byName[name];

    // ค้นหา LINE User ID
    var uid = '';
    if (empSh) uid = _findEmpLineId(name, empSh);
    if (!uid) {
      Logger.log('⚠️ ไม่พบ LINE ID หรือปิดแจ้งเตือน: ' + name);
      skipped++;
      // บันทึก log ว่าข้ามไป
      _writeLineLog(logSh, 'อัตโนมัติ-ไม่มี LINE ID',
        name, records.map(function(r){return r.topic;}).join(', '),
        { ok: false, code: 0, body: 'ไม่พบ LINE User ID ในระบบ' });
      return;
    }

    var msg    = _buildDocReminderMsg(name, records);
    var result2 = _linePush(uid, msg);

    Logger.log((result2.ok ? '✅' : '❌') + ' → ' + name +
               ' (' + records.length + ' รายการ) HTTP ' + result2.code);
    if (result2.ok) sent++;

    // บันทึก log ทุก record
    records.forEach(function(rec) {
      _writeLineLog(logSh, 'อัตโนมัติ-แจ้งเตือนครั้งเดียว', name, rec.topic, result2);
    });
  });

  if (logSh) SpreadsheetApp.flush();

  Logger.log('══════════════════════════════════════════════');
  Logger.log('  dailyNotificationCheck DONE:');
  Logger.log('  ส่งสำเร็จ=' + sent + ' คน | ข้าม=' + skipped + ' คน');
  Logger.log('══════════════════════════════════════════════');
}

function _countUniqueNames(arr) {
  var seen = {};
  arr.forEach(function(r) { seen[r.name] = true; });
  return Object.keys(seen).length;
}

// ============================================================
// === getFollowupStatus(afterDays) — ตรวจสถานะรายการค้าง ===
// คืน { success, pending[], notFound[], checkedAt }
// ============================================================
function getFollowupStatus(afterDays) {
  afterDays = (typeof afterDays === 'number') ? afterDays : LINE_AFTER_DAYS;

  try {
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var trainSh = ss.getSheetByName('อบรม ปีงบประมาณ 2569');
    var regSh   = ss.getSheetByName('Registrations');

    if (!trainSh) return { success: false, message: 'ไม่พบ Sheet "อบรม ปีงบประมาณ 2569"' };
    if (!regSh)   return { success: false, message: 'ไม่พบ Sheet "Registrations"' };

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    // ── อ่าน Registrations ──
    var regData = regSh.getDataRange().getValues();
    var regHdr  = regData[0].map(function(h) { return String(h || '').trim(); });
    var rColName    = getColIdx(regHdr, 'Full Name',  'ชื่อ-สกุล',       4);
    var rColTopic   = getColIdx(regHdr, 'Topic',      'หัวข้อ/หลักสูตร', 9);
    // หมายเหตุ: ไม่เช็ค Summary — ส่งแจ้งเตือนทุกคนที่ครบกำหนด ไม่ว่าจะส่งสรุปหรือยัง

    // ── อ่าน อบรม ปีงบประมาณ 2569 ──
    var trainData = trainSh.getDataRange().getValues();
    var trainHdr  = trainData[0].map(function(h) { return String(h || '').trim(); });
    var trainHdrNorm = trainHdr.map(function(h) {
      return (HEADER_ALIASES && HEADER_ALIASES[h]) ? HEADER_ALIASES[h] : h;
    });

    var tColName  = _findCol(trainHdrNorm, trainHdr, ['ชื่อ-สกุล','ชื่อ สกุล','ชื่อ','Full Name']);
    var tColTopic = _findCol(trainHdrNorm, trainHdr, ['เรื่อง/หลักสูตร','หลักสูตร','เรื่อง','Topic']);
    var tColEnd   = _findCol(trainHdrNorm, trainHdr, ['วันที่สิ้นสุด','ถึงวันที่','End Date']);
    var tColStart = _findCol(trainHdrNorm, trainHdr, ['วันที่เริ่ม','วันที่','Date']);

    if (tColName  < 0) return { success: false, message: 'ไม่พบคอลัมน์ "ชื่อ-สกุล" ใน Sheet อบรมฯ' };
    if (tColTopic < 0) return { success: false, message: 'ไม่พบคอลัมน์ "เรื่อง/หลักสูตร" ใน Sheet อบรมฯ' };
    if (tColEnd   < 0) return { success: false, message: 'ไม่พบคอลัมน์ "วันที่สิ้นสุด" ใน Sheet อบรมฯ' };

    var pending  = [];
    var notFound = [];

    for (var ti = 1; ti < trainData.length; ti++) {
      var tRow   = trainData[ti];
      var tName  = String(tRow[tColName]  || '').trim();
      var tTopic = String(tRow[tColTopic] || '').trim();
      if (!tName || !tTopic) continue;

      // ตรวจวันที่สิ้นสุด
      var endRaw = tRow[tColEnd];
      if (!endRaw) continue;
      var endDate = endRaw instanceof Date ? new Date(endRaw) : new Date(endRaw);
      if (isNaN(endDate.getTime())) continue;
      endDate.setHours(0, 0, 0, 0);

      var daysAfterN = Math.round((today - endDate) / 86400000);
      if (daysAfterN !== afterDays) continue; // ส่งเฉพาะวันที่ครบกำหนดพอดี

      var endDateStr   = Utilities.formatDate(endDate, 'Asia/Bangkok', 'yyyy-MM-dd');
      var startDateStr = '';
      if (tColStart >= 0 && tRow[tColStart]) {
        var sd = tRow[tColStart] instanceof Date ? tRow[tColStart] : new Date(tRow[tColStart]);
        if (!isNaN(sd.getTime()))
          startDateStr = Utilities.formatDate(sd, 'Asia/Bangkok', 'yyyy-MM-dd');
      }

      // ค้นหาใน Registrations (เช็คแค่ ชื่อ + เรื่อง — ไม่สนใจ Summary)
      var matchRowIdx = -1;
      var tNameLower  = tName.toLowerCase();
      var tTopicLower = tTopic.toLowerCase();

      for (var ri = 1; ri < regData.length; ri++) {
        var rName  = String(regData[ri][rColName]  || '').trim().toLowerCase();
        var rTopic = String(regData[ri][rColTopic] || '').trim().toLowerCase();
        if (rName === tNameLower && rTopic === tTopicLower) {
          matchRowIdx = ri;
          break;
        }
      }

      var record = {
        trainRowIdx : ti + 1,
        name        : tName,
        topic       : tTopic,
        startDate   : startDateStr,
        endDate     : endDateStr,
        daysAfter   : daysAfterN,
        regRowIdx   : matchRowIdx > 0 ? matchRowIdx + 1 : null
      };

      // ทุก record ที่ครบกำหนด → pending (ส่งแจ้งเตือน) ไม่ว่าจะลงทะเบียนใน Reg แล้วหรือยัง
      if (matchRowIdx < 0) {
        notFound.push(record); // หาใน Registrations ไม่เจอ (log ไว้เท่านั้น)
      } else {
        pending.push(record);  // พบใน Registrations → แจ้งเตือน
      }
    }

    Logger.log('✅ getFollowupStatus: pending=' + pending.length +
               ' | notFound=' + notFound.length);

    return {
      success   : true,
      pending   : pending,
      notFound  : notFound,
      checkedAt : Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm')
    };

  } catch(e) {
    Logger.log('❌ getFollowupStatus: ' + e);
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// === getFollowupPendingReport(token, afterDays) — เรียกจาก Admin ===
// ============================================================
function getFollowupPendingReport(token, afterDays) {
  if (!verifyAdminToken(token) && !verifyAdminTokenPrevHour(token)) {
    return { success: false, message: 'ไม่มีสิทธิ์ กรุณา Login ใหม่' };
  }
  return getFollowupStatus(afterDays || LINE_AFTER_DAYS);
}

// ============================================================
// === sendDocReminderManual(payload, token) — Admin กดส่งเอง ===
// ============================================================
function sendDocReminderManual(payload, token) {
  if (!verifyAdminToken(token)) {
    return { success: false, message: 'ไม่มีสิทธิ์ กรุณา Login ใหม่' };
  }
  try {
    if (!payload || !payload.lineId || !payload.fullName) {
      return { success: false, message: 'ข้อมูลไม่ครบ (lineId / fullName)' };
    }

    var lineId   = String(payload.lineId   || '').trim();
    var fullName = String(payload.fullName || '').trim();

    if (!/^U[0-9a-f]{32}$/i.test(lineId)) {
      return { success: false,
               message: 'LINE User ID ไม่ถูกรูปแบบ: ' + lineId +
                        '\nUser ID ต้องขึ้นต้นด้วย U ตามด้วยตัวอักษร/ตัวเลข 32 ตัว' };
    }

    // ดึงหลักสูตรล่าสุดของพนักงาน (ใช้สำหรับ Flex Message)
    var latestCourse = _getLatestCourse(fullName);
    var records = [{
      topic    : latestCourse || payload.course || 'การอบรม/ประชุม',
      endDate  : payload.endDate || '',
      daysAfter: payload.daysAfter || 0
    }];

    var flexMsg = _buildDocReminderMsg(fullName, records);
    var result  = _linePush(lineId, flexMsg);

    // บันทึก Log
    var logSh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Log_การแจ้งเตือน');
    records.forEach(function(rec) {
      _writeLineLog(logSh, 'Admin-แจ้งส่งสรุปด้วยตนเอง', fullName, rec.topic, result);
    });
    if (logSh) SpreadsheetApp.flush();

    Logger.log((result.ok ? '✅' : '❌') + ' sendDocReminderManual → ' + fullName + ' HTTP ' + result.code);
    if (!result.ok) {
      return { success: false,
               message: 'LINE API ตอบกลับ HTTP ' + result.code + ': ' + result.body.substring(0, 150) };
    }
    return { success: true };

  } catch(e) {
    Logger.log('❌ sendDocReminderManual ERROR: ' + e);
    return { success: false, message: e.toString() };
  }
}

// helper: ดึงหลักสูตรล่าสุดของพนักงาน
function _getLatestCourse(fullName) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TRAINING_REQUEST_SHEET);
    if (!sheet || sheet.getLastRow() <= 1) return '';
    var data  = sheet.getDataRange().getValues();
    var hdr   = data[0].map(function(h) { return String(h || '').trim(); });
    var cName  = hdr.indexOf('ชื่อ-สกุล');
    var cTopic = hdr.indexOf('เรื่อง/หลักสูตร');
    var cTs    = hdr.indexOf('Timestamp');
    var best = null, bestTs = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][cName] || '').trim() !== fullName) continue;
      var ts = (data[i][cTs] instanceof Date) ? data[i][cTs].getTime()
             : (data[i][cTs] ? new Date(data[i][cTs]).getTime() : 0);
      if (ts > bestTs) { bestTs = ts; best = data[i]; }
    }
    return best && cTopic >= 0 ? String(best[cTopic] || '').trim() : '';
  } catch(e) { return ''; }
}

// ============================================================
// === LINE Trigger Setup Functions ===
// ============================================================

/** ตั้ง Trigger รายวัน (รันครั้งเดียวพอ) */
function LINE_setTrigger() {
  // ลบ trigger เดิมทุกตัวที่เกี่ยวกับ LINE ก่อน (ป้องกันซ้ำ)
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'dailyNotificationCheck' || fn === 'dailyFollowupCheck') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('dailyNotificationCheck')
    .timeBased().atHour(LINE_NOTIFY_HOUR).everyDays(1).inTimezone('Asia/Bangkok').create();
  Logger.log('✅ Trigger ตั้งแล้ว — dailyNotificationCheck ทุกวัน ' + LINE_NOTIFY_HOUR + ':00 น.');
  Logger.log('   (ยกเลิก trigger เก่าทั้งหมดที่ซ้ำกันแล้ว)');
}

/** ยกเลิก Trigger */
function LINE_removeTrigger() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'dailyNotificationCheck' || fn === 'dailyFollowupCheck') {
      ScriptApp.deleteTrigger(t); n++;
    }
  });
  Logger.log('✅ ลบ Trigger แล้ว ' + n + ' รายการ');
}

// ============================================================
// === LINE_setup() — STEP 1: ตรวจสอบระบบ ===
// ============================================================
function LINE_setup() {
  Logger.log('══════════════════════════════════════════════');
  Logger.log('  LINE System — Setup Check');
  Logger.log('══════════════════════════════════════════════');

  var token = _getLineToken();
  if (!token) {
    Logger.log('❌ TOKEN: ไม่พบ');
    Logger.log('   → เปิด Sheet "ตั้งค่าระบบ" แล้วใส่ Channel Access Token');
    Logger.log('   → ได้จาก LINE Developers Console > Channel > Messaging API > Channel access token');
    return;
  }
  Logger.log('✅ TOKEN: พบแล้ว (' + token.length + ' ตัวอักษร)');

  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/info', {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    Logger.log('❌ BOT INFO: HTTP ' + res.getResponseCode() + ' — Token ผิดหรือหมดอายุ');
    return;
  }
  var bot = JSON.parse(res.getContentText());
  Logger.log('✅ BOT: ' + bot.displayName + ' (@' + (bot.basicId || '?') + ')');
  Logger.log('   chatMode: ' + bot.chatMode +
             (bot.chatMode === 'bot' ? ' ✅' : ' ⚠️ ต้องเปลี่ยนเป็น "bot mode" ใน LINE Developers'));

  var url = '';
  try { url = ScriptApp.getService().getUrl(); } catch(e) {}
  Logger.log(url ? '✅ Webhook URL: ' + url : '⚠️ ยังไม่ได้ Deploy Web App');

  Logger.log('──────────────────────────────────────────────');
  Logger.log('ขั้นต่อไป:');
  Logger.log('  1. User ส่งข้อความหา Bot (ครั้งแรก) → Bot ตอบ User ID');
  Logger.log('  2. Admin ผูก User ID กับชื่อพนักงานในหน้าเว็บ');
  Logger.log('  3. รัน LINE_sendTest() เพื่อทดสอบส่งจริง');
  Logger.log('  4. รัน LINE_setTrigger() เพื่อเปิดอัตโนมัติ');
  Logger.log('══════════════════════════════════════════════');
}

// ============================================================
// === LINE_sendTest() — STEP 4: ทดสอบส่งข้อความจริง ===
// ============================================================
function LINE_sendTest() {
  if (!LINE_TEST_USER_ID) {
    Logger.log('❌ ยังไม่ได้ใส่ LINE_TEST_USER_ID ในโค้ด'); return;
  }
  Logger.log('กำลังส่งทดสอบ → ' + LINE_TEST_USER_ID);
  var r = _linePush(LINE_TEST_USER_ID, {
    type: 'flex',
    altText: '✅ ทดสอบระบบแจ้งเตือน HRD สันทราย',
    contents: {
      type: 'bubble',
      header: {
        type: 'box', layout: 'vertical', backgroundColor: '#059669', paddingAll: '16px',
        contents: [{
          type: 'text', text: '✅ ระบบแจ้งเตือน HRD พร้อมแล้ว!',
          color: '#ffffff', size: 'lg', weight: 'bold'
        }]
      },
      body: {
        type: 'box', layout: 'vertical', paddingAll: '16px',
        contents: [
          { type: 'text', text: 'โรงพยาบาลสันทราย', size: 'md', weight: 'bold' },
          { type: 'text', text: 'ระบบแจ้งเตือนการอบรม HRD ทำงานปกติ 🎉',
            size: 'sm', color: '#666666', margin: 'sm' },
          { type: 'text',
            text: 'ทดสอบเมื่อ ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') + ' น.',
            size: 'xs', color: '#94a3b8', margin: 'sm' }
        ]
      }
    }
  });
  if (r.ok) {
    Logger.log('✅ ส่งสำเร็จ! HTTP 200 — ระบบพร้อมใช้งาน');
    Logger.log('   → รัน LINE_setTrigger() เพื่อเปิดส่งอัตโนมัติ');
  } else {
    Logger.log('❌ HTTP ' + r.code + ': ' + r.body);
    if (r.code === 400) Logger.log('   → User ยังไม่ได้ add Bot หรือ User ID ผิด');
    if (r.code === 401) Logger.log('   → Token ผิดหรือหมดอายุ');
  }
}

// ============================================================
// === testFollowupStatus() — ดูรายงานใน Logger โดยไม่ส่ง LINE ===
// ============================================================
function testFollowupStatus() {
  var result = getFollowupStatus(LINE_AFTER_DAYS);
  if (!result.success) { Logger.log('❌ ' + result.message); return; }

  Logger.log('══════════════════════════════════════════════════');
  Logger.log('  รายงานติดตามสรุปอบรม ณ ' + result.checkedAt);
  Logger.log('  เกณฑ์: ครบกำหนด ' + LINE_AFTER_DAYS + ' วันพอดีหลังสิ้นสุด (ส่งครั้งเดียว)');
  Logger.log('══════════════════════════════════════════════════');

  Logger.log('\n📋  รายการที่ครบ ' + LINE_AFTER_DAYS + ' วัน (แจ้งเตือนวันนี้): ' + result.pending.length + ' รายการ');
  result.pending.forEach(function(r, i) {
    Logger.log('  ' + (i+1) + '. [Train แถว ' + r.trainRowIdx + ']');
    Logger.log('     ชื่อ   : ' + r.name);
    Logger.log('     เรื่อง : ' + r.topic);
    Logger.log('     สิ้นสุด: ' + r.endDate + '  (ค้างมา ' + r.daysAfter + ' วัน)');
  });

  Logger.log('\n❓  หาใน Registrations ไม่เจอ: ' + result.notFound.length + ' รายการ');
  result.notFound.forEach(function(r, i) {
    Logger.log('  ' + (i+1) + '. ' + r.name + ' | ' + r.topic +
               '  (Train แถว ' + r.trainRowIdx + ')');
  });
  Logger.log('══════════════════════════════════════════════════');
}

// ============================================================
// === getSystemConfig() — อ่าน config เดิม (backward compat) ===
// ============================================================
function getSystemConfig(configSheet) {
  if (!configSheet) {
    configSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ตั้งค่าระบบ');
  }
  var data = configSheet ? configSheet.getDataRange().getValues() : [];
  var cfg  = { beforeDays: 7, afterDays: LINE_AFTER_DAYS, notifyHour: LINE_NOTIFY_HOUR,
               lineToken: '', driveFolderId: '' };
  for (var i = 1; i < data.length; i++) {
    var k = String(data[i][0] || '').trim(), v = String(data[i][1] || '').trim();
    if (k === 'แจ้งเตือนก่อนอบรม (วัน)')       cfg.beforeDays    = parseInt(v) || 7;
    if (k === 'แจ้งเตือนหลังอบรม (วัน)')        cfg.afterDays     = parseInt(v) || LINE_AFTER_DAYS;
    if (k === 'เวลาส่งแจ้งเตือน (ชั่วโมง)')     cfg.notifyHour    = parseInt(v) || LINE_NOTIFY_HOUR;
    if (k === 'LINE Channel Access Token')       cfg.lineToken     = v;
    if (k === 'Google Drive Folder ID')          cfg.driveFolderId = v;
  }
  return cfg;
}

// helper (ใช้ร่วมกับ getFollowupStatus)
function _findCol(normHdr, rawHdr, candidates) {
  for (var c = 0; c < candidates.length; c++) {
    var idx = normHdr.indexOf(candidates[c]);
    if (idx >= 0) return idx;
    idx = rawHdr.indexOf(candidates[c]);
    if (idx >= 0) return idx;
  }
  return -1;
}

function _fmtDateTh(ds) {
  if (!ds) return '-';
  try {
    var d = new Date(ds);
    if (isNaN(d.getTime())) return ds;
    var m = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.',
             'ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
    return d.getDate() + ' ' + m[d.getMonth()] + ' ' + (d.getFullYear() + 543);
  } catch(e) { return ds; }
}
