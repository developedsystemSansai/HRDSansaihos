// === PART 7: บันทึกขออนุมัติ + อัพโหลดไฟล์ Drive ===
// ============================================

// ชื่อ Sheet สำหรับบันทึกขออนุมัติ — ชี้ไปที่ Sheet ที่มีอยู่แล้ว
var TRAINING_REQUEST_SHEET = 'อบรม ปีงบประมาณ 2569';

/**
 * Map ข้อมูลลง Sheet โดยอ่าน header แถวแรกแล้ว match ชื่อคอลัมน์อัตโนมัติ
 * ถ้า Sheet ว่างเปล่า (ไม่มี header) จะสร้าง header มาตรฐานให้
 * ถ้ามี header อยู่แล้ว จะ append ต่อท้ายโดย map ตาม column name
 */
var STANDARD_HEADERS = [
  'ลำดับ',
  'Timestamp',
  'ชื่อ-สกุล',
  'ตำแหน่ง',
  'ระดับ',
  'กลุ่มภารกิจ',
  'กลุ่มงาน',
  'เลขที่ ชม.5',
  'เลขที่บันทึก',
  'เรื่อง/หลักสูตร',
  'วันที่เริ่ม',
  'วันที่สิ้นสุด',
  'สถานะหลักสูตร',
  'ประเภทการอบรม',
  'สถานที่',
  'จังหวัด',
  'หน่วยงานผู้จัด',
  'ค่าลงทะเบียน',
  'ค่าเบี้ยเลี้ยง',
  'ค่าที่พัก',
  'ค่าพาหนะ',
  'เงินโครงการ',
  'รวมค่าใช้จ่าย',
  'แหล่งงบประมาณ',
  'ชื่อไฟล์แนบ',
  'URL ไฟล์แนบ',
  'เอกสารสรุปผล',
  'สัญญา',
  'วุฒิบัตร',
  'ใบเสร็จ',
  'บันทึก ชม.5',
  'บันทึกขออนุมัติ',
  'บันทึกความ+ชม.5',
  'LINE ID',
  'ผู้บันทึก',
  'วันที่บันทึก',
  'หมายเหตุ'
];

// Alias map — ชื่อ header ที่อาจพบใน Sheet เดิม → ชื่อมาตรฐานที่ระบบใช้
// เพิ่ม alias ที่นี่ถ้า Sheet มี header ชื่อต่างออกไป
var HEADER_ALIASES = {
  'ชื่อ'              : 'ชื่อ-สกุล',
  'ชื่อ สกุล'         : 'ชื่อ-สกุล',
  'Full Name'         : 'ชื่อ-สกุล',
  'ชื่อผู้บันทึก'     : 'ผู้บันทึก',
  'หลักสูตร'          : 'เรื่อง/หลักสูตร',
  'เรื่อง'            : 'เรื่อง/หลักสูตร',
  'Topic'             : 'เรื่อง/หลักสูตร',
  'วันที่'            : 'วันที่เริ่ม',
  'Date'              : 'วันที่เริ่ม',
  'กลุ่มงาน/ฝ่าย'    : 'กลุ่มงาน',
  'ฝ่าย'             : 'กลุ่มงาน',
  'Department'        : 'กลุ่มงาน',
  'Position'          : 'ตำแหน่ง',
  'สถานที่จัด'        : 'สถานที่',
  'สถานที่อบรม'       : 'สถานที่',
  'Location'          : 'สถานที่',
  'สถานที่จัด/จังหวัด': 'จังหวัด',
  'ค่าใช้จ่ายรวม'     : 'รวมค่าใช้จ่าย',
  'รวม'               : 'รวมค่าใช้จ่าย',
  'Total Cost'        : 'รวมค่าใช้จ่าย',
  'งบประมาณ'          : 'แหล่งงบประมาณ',
  'ไฟล์แนบ'           : 'ชื่อไฟล์แนบ',
  'เอกสารแนบ'         : 'ชื่อไฟล์แนบ',
  'ค่าเบี้ยเลี้ยงเดินทาง': 'ค่าเบี้ยเลี้ยง',
  'เลขที่ ชม.5 '      : 'เลขที่ ชม.5',
  'เลขที่บันทึก\nข้อความขออนุมัติ': 'เลขที่บันทึก',
  'สถานะหลักสูตร*'    : 'สถานะหลักสูตร',
  'ประเภทการอบรม**'   : 'ประเภทการอบรม',
  'เอกสารสรุปผลการอบรม': 'เอกสารสรุปผล',
  'สัญญาอบรม'         : 'สัญญา'
};

/**
 * อ่าน header ของ Sheet — ถ้าแถวแรกว่างเปล่าให้สร้าง STANDARD_HEADERS
 * รองรับ alias (ชื่อ header ทางเลือก) อัตโนมัติ
 * ถ้า Sheet มี header อยู่แล้ว จะเพิ่ม column ที่ขาดต่อท้ายอัตโนมัติ
 * คืน { sheet, headers, headerMap }
 */
function getOrInitTrainingSheet() {
  var ss    = getOrCreateSheet();
  var sheet = ss.getSheetByName(TRAINING_REQUEST_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(TRAINING_REQUEST_SHEET);
  }

  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();

  var rawHeaders = [];
  if (lastRow >= 1 && lastCol >= 1) {
    rawHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
      return String(h === null || h === undefined ? '' : h).trim();
    });
  }

  var nonEmpty = rawHeaders.filter(function(h) { return h !== ''; });

  if (nonEmpty.length === 0) {
    // Sheet ว่าง — สร้าง header มาตรฐาน
    sheet.getRange(1, 1, 1, STANDARD_HEADERS.length)
      .setValues([STANDARD_HEADERS])
      .setFontWeight('bold')
      .setBackground('#1e3a8a')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    rawHeaders = STANDARD_HEADERS.slice();
    Logger.log('\u2705 สร้าง header ใหม่ใน Sheet: ' + TRAINING_REQUEST_SHEET);
  } else {
    Logger.log('\u2705 อ่าน ' + nonEmpty.length + ' headers จาก Sheet: ' + nonEmpty.slice(0,5).join(' | ') + '...');

    // ---- เพิ่ม column ที่ขาดโดยอัตโนมัติ (ไม่ลบ column เดิม) ----
    // สร้าง map ชื่อ header ที่มีอยู่ (รองรับ alias ด้วย)
    var existingNames = {};
    for (var ei = 0; ei < rawHeaders.length; ei++) {
      var eh = rawHeaders[ei];
      if (eh) {
        existingNames[eh] = true;
        // ถ้า header เป็น alias ให้นับ standard name ด้วย
        if (HEADER_ALIASES[eh]) existingNames[HEADER_ALIASES[eh]] = true;
      }
    }

    var colsToAdd = [];
    for (var si = 0; si < STANDARD_HEADERS.length; si++) {
      var sh = STANDARD_HEADERS[si];
      if (!existingNames[sh]) {
        colsToAdd.push(sh);
      }
    }

    if (colsToAdd.length > 0) {
      var startAddCol = lastCol + 1; // ต่อท้าย column ที่มีอยู่
      // เพิ่ม header ทีละ column
      for (var ci = 0; ci < colsToAdd.length; ci++) {
        var newColIdx = startAddCol + ci;
        var cell = sheet.getRange(1, newColIdx);
        cell.setValue(colsToAdd[ci])
            .setFontWeight('bold')
            .setBackground('#1e3a8a')
            .setFontColor('#ffffff');
        rawHeaders.push(colsToAdd[ci]);
      }
      SpreadsheetApp.flush();
      Logger.log('\u2705 เพิ่ม ' + colsToAdd.length + ' columns ใหม่: ' + colsToAdd.join(', '));
    }
  }

  // สร้าง headerMap รองรับทั้ง exact match และ alias
  var headerMap = {};
  for (var i = 0; i < rawHeaders.length; i++) {
    var h = rawHeaders[i];
    if (!h) continue;
    // ใส่ชื่อจริง
    headerMap[h] = i;
    // ถ้าชื่อนี้มี alias reverse (alias → standard) ให้ใส่ standard ด้วย
    if (HEADER_ALIASES[h] !== undefined && headerMap[HEADER_ALIASES[h]] === undefined) {
      headerMap[HEADER_ALIASES[h]] = i;
      Logger.log('  alias: "' + h + '" \u2192 "' + HEADER_ALIASES[h] + '" at col ' + (i+1));
    }
  }

  return { sheet: sheet, headers: rawHeaders, headerMap: headerMap };
}

/**
 * ฟังก์ชัน debug: แสดง header จริงของ Sheet อบรม + ตรวจสอบ column ที่ขาด
 * รันจาก Apps Script Editor เพื่อ diagnose ปัญหา
 */
function debugAndFixTrainingSheet() {
  var info = getOrInitTrainingSheet(); // จะ auto-add column ที่ขาด
  var sheet = info.sheet;
  Logger.log('=== debugAndFixTrainingSheet ===');
  Logger.log('Sheet: ' + TRAINING_REQUEST_SHEET);
  Logger.log('Total rows: ' + sheet.getLastRow());
  Logger.log('Total cols: ' + sheet.getLastColumn());
  Logger.log('Header count: ' + info.headers.length);
  Logger.log('--- Headers ---');
  info.headers.forEach(function(h, i) {
    var found = info.headerMap[h] !== undefined ? '\u2705' : '\u274c';
    Logger.log('  [' + (i+1) + '] ' + found + ' "' + h + '"');
  });
  var missing = STANDARD_HEADERS.filter(function(h) {
    return info.headerMap[h] === undefined;
  });
  if (missing.length > 0) {
    Logger.log('--- \u274c Missing standard headers ---');
    missing.forEach(function(h) { Logger.log('  - "' + h + '"'); });
  } else {
    Logger.log('--- \u2705 All standard headers present ---');
  }
  // แสดง 3 แถวแรก
  if (sheet.getLastRow() > 1) {
    var sample = sheet.getRange(2, 1, Math.min(3, sheet.getLastRow()-1), sheet.getLastColumn()).getValues();
    Logger.log('--- Sample data (first 3 rows) ---');
    sample.forEach(function(row, ri) {
      var preview = info.headers.slice(0, Math.min(6, info.headers.length)).map(function(h, ci) {
        return '"' + h + '"=' + (row[ci] !== '' ? row[ci] : '(ว่าง)');
      }).join(' | ');
      Logger.log('  Row ' + (ri+2) + ': ' + preview);
    });
  }
  Logger.log('=== เสร็จสิ้น ===');
}

/**
 * สร้าง row array ตาม header ของ Sheet จริง
 * ใช้ headerMap เพื่อ match ชื่อคอลัมน์
 */
function buildRowFromData(data, fileNames, fileUrls, headerMap, headerLength, sequenceNum) {
  var now = new Date();
  // ข้อมูลที่ต้องการใส่ — key = header name, value = ค่า
  var dataMap = {
    'ลำดับ'           : sequenceNum || '',
    'Timestamp'       : now,
    'ชื่อ-สกุล'      : String(data.name       || '').trim(),
    'ตำแหน่ง'        : String(data.position   || '').trim(),
    'ระดับ'           : String(data.level      || '').trim(),
    'กลุ่มภารกิจ'    : String(data.mission    || '').trim(),
    'กลุ่มงาน'       : String(data.dept       || '').trim(),
    'เลขที่ ชม.5'    : String(data.hm5        || '').trim(),
    'เลขที่บันทึก'   : String(data.memo       || '').trim(),
    'เรื่อง/หลักสูตร': String(data.topic      || '').trim(),
    'วันที่เริ่ม'    : data.startDate ? new Date(data.startDate) : '',
    'วันที่สิ้นสุด'  : data.endDate   ? new Date(data.endDate)   : '',
    'สถานะหลักสูตร'  : String(data.status     || '').trim(),
    'ประเภทการอบรม'  : String(data.type       || '').trim(),
    'สถานที่'        : String(data.venue      || '').trim(),
    'จังหวัด'        : String(data.province   || '').trim(),
    'หน่วยงานผู้จัด' : String(data.organizer  || '').trim(),
    'ค่าลงทะเบียน'   : parseFloat(data.costReg)    || 0,
    'ค่าเบี้ยเลี้ยง'  : parseFloat(data.costTravel) || 0,
    'ค่าที่พัก'       : parseFloat(data.costHotel)  || 0,
    'ค่าพาหนะ'        : parseFloat(data.costTrans)  || 0,
    'เงินโครงการ'     : parseFloat(data.costProj)   || 0,
    'รวมค่าใช้จ่าย'  : parseFloat(data.totalCost)  || 0,
    'แหล่งงบประมาณ'  : String(data.budget     || '').trim(),
    'ชื่อไฟล์แนบ'    : fileNames,
    'URL ไฟล์แนบ'    : fileUrls,
    'เอกสารสรุปผล'   : data.docSummary  ? '\u2713' : '',
    'สัญญา'          : data.docContract ? '\u2713' : '',
    'วุฒิบัตร'       : data.docCert     ? '\u2713' : '',
    'ใบเสร็จ'         : data.docReceipt  ? '\u2713' : '',
    'บันทึก ชม.5'    : data.docHm5      ? '\u2713' : '',
    'บันทึกขออนุมัติ': data.docApproval ? '\u2713' : '',
    'บันทึกความ+ชม.5': data.docCombined ? '\u2713' : '',
    'LINE ID'         : String(data.lineId    || '').trim(),
    'ผู้บันทึก'      : String(data.recorder  || '').trim(),
    'วันที่บันทึก'   : data.recDate ? new Date(data.recDate) : now,
    'หมายเหตุ'        : String(data.note      || '').trim()
  };

  // สร้าง row ตาม header ที่มีอยู่จริงใน Sheet
  var row = new Array(headerLength).fill('');
  for (var key in dataMap) {
    if (headerMap.hasOwnProperty(key)) {
      row[headerMap[key]] = dataMap[key];
    }
  }
  return row;
}

/**
 * อัพโหลดไฟล์ base64 ไปยัง Google Drive Folder
 * คืน { name, url, id, size } ต่อไฟล์
 */
/**
 * อัพโหลดไฟล์ไปยัง Google Drive โดยใช้ Drive REST API ผ่าน UrlFetchApp
 * วิธีนี้ไม่ต้องพึ่ง DriveApp scope จาก Web App — ใช้ ScriptApp.getOAuthToken() แทน
 * ซึ่งจะได้ token จาก scope ที่ appsscript.json กำหนดไว้
 */
function saveFilesToDrive(files) {
  var results = [];
  if (!files || files.length === 0) return results;

  // ดึง OAuth token จาก Script runtime (ใช้ scope ที่ประกาศใน appsscript.json)
  var token = ScriptApp.getOAuthToken();

  // ตรวจสอบว่าโฟลเดอร์มีอยู่จริง (ใช้ Drive API แทน DriveApp)
  var folderCheck = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files/' + ATTACHMENTS_FOLDER_ID +
    '?fields=id,name,mimeType',
    { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
  );
  if (folderCheck.getResponseCode() !== 200) {
    Logger.log('\u274c ไม่พบโฟลเดอร์ id=' + ATTACHMENTS_FOLDER_ID +
               ' HTTP=' + folderCheck.getResponseCode());
    throw new Error('ไม่พบโฟลเดอร์ Drive: ' + folderCheck.getContentText().substring(0, 100));
  }
  var folderName = JSON.parse(folderCheck.getContentText()).name;
  Logger.log('\u2705 โฟลเดอร์: ' + folderName + ' (id=' + ATTACHMENTS_FOLDER_ID + ')');

  for (var i = 0; i < files.length; i++) {
    var f = files[i];
    try {
      if (!f || !f.name || !f.base64) {
        Logger.log('\u274c ข้ามไฟล์ index=' + i + ' (ข้อมูลไม่ครบ)');
        continue;
      }

      // แยก base64 และ mimeType จาก data URI
      var raw      = f.base64.indexOf(',') !== -1 ? f.base64.split(',').pop() : f.base64;
      var mimeType = f.mimeType || 'application/octet-stream';
      if ((!f.mimeType || f.mimeType === 'application/octet-stream') &&
          f.base64.indexOf('data:') === 0) {
        var mm = f.base64.match(/^data:([^;]+);/);
        if (mm) mimeType = mm[1];
      }

      var bytes = Utilities.base64Decode(raw);

      // สร้างไฟล์ผ่าน Drive REST API (multipart upload)
      var boundary = '-------HRD_BOUNDARY_' + Date.now();
      var metadata = JSON.stringify({
        name: f.name,
        parents: [ATTACHMENTS_FOLDER_ID],
        mimeType: mimeType
      });

      // สร้าง multipart body
      var metaPart = '--' + boundary + '\r\n' +
                     'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
                     metadata + '\r\n';
      var dataPart = '--' + boundary + '\r\n' +
                     'Content-Type: ' + mimeType + '\r\n' +
                     'Content-Transfer-Encoding: base64\r\n\r\n' +
                     raw + '\r\n';
      var closing  = '--' + boundary + '--';

      var bodyBlob = Utilities.newBlob(metaPart + dataPart + closing)
                              .setContentType('multipart/related; boundary="' + boundary + '"');

      var uploadResp = UrlFetchApp.fetch(
        'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink',
        {
          method: 'POST',
          headers: { 'Authorization': 'Bearer ' + token },
          contentType: 'multipart/related; boundary="' + boundary + '"',
          payload: bodyBlob.getBytes(),
          muteHttpExceptions: true
        }
      );

      var uploadCode = uploadResp.getResponseCode();
      if (uploadCode !== 200) {
        throw new Error('Upload HTTP ' + uploadCode + ': ' + uploadResp.getContentText().substring(0, 200));
      }

      var uploadData = JSON.parse(uploadResp.getContentText());
      var driveId    = uploadData.id;
      var driveUrl   = 'https://drive.google.com/file/d/' + driveId + '/view?usp=sharing';

      // ตั้งสิทธิ์ให้ทุกคนที่มี link ดูได้
      UrlFetchApp.fetch(
        'https://www.googleapis.com/drive/v3/files/' + driveId + '/permissions',
        {
          method: 'POST',
          headers: { 'Authorization': 'Bearer ' + token,
                     'Content-Type': 'application/json' },
          payload: JSON.stringify({ role: 'reader', type: 'anyone' }),
          muteHttpExceptions: true
        }
      );

      results.push({ name: f.name, url: driveUrl, id: driveId, size: bytes.length });
      Logger.log('\u2705 อัพโหลดสำเร็จ: ' + f.name + ' | id=' + driveId);

    } catch (err) {
      var fname = (f && f.name) ? f.name : ('file_' + i);
      Logger.log('\u274c อัพโหลดล้มเหลว: ' + fname + ' — ' + err.toString());
      results.push({ name: fname, url: '', id: '', error: err.toString() });
    }
  }
  return results;
}

/**
 * บันทึกคำขออบรมลง Sheet + อัพโหลดไฟล์แนบ
 * (ปรับปรุงให้เริ่มเลขลำดับที่ 215 และรันต่อไปเรื่อยๆ)
 */
function saveTrainingRequest(data, files) {
  try {
    var info      = getOrInitTrainingSheet();
    var sheet     = info.sheet;
    var headerMap = info.headerMap;
    var hLen      = info.headers.length;

    Logger.log('=== saveTrainingRequest START ===');
    Logger.log('Sheet: ' + TRAINING_REQUEST_SHEET + ' | headers: ' + hLen);
    Logger.log('headerMap URL ไฟล์แนบ = ' + headerMap['URL ไฟล์แนบ']);
    Logger.log('headerMap ชื่อไฟล์แนบ = ' + headerMap['ชื่อไฟล์แนบ']);

    // อัพโหลดไฟล์ก่อน
    var fileResults = [];
    var fileNames   = '';
    var fileUrls    = '';
    if (files && files.length > 0) {
      Logger.log('กำลังอัพโหลด ' + files.length + ' ไฟล์...');
      fileResults = saveFilesToDrive(files);
      var okList   = fileResults.filter(function(r) { return r.url && r.url.length > 10; });
      var failList = fileResults.filter(function(r) { return !r.url || r.url.length <= 10; });
      fileNames = fileResults.map(function(r) { return r.name; }).join(', ');
      fileUrls  = okList.map(function(r) { return r.url; }).join('\n');
      Logger.log('✅ Upload: ' + okList.length + '/' + fileResults.length + ' สำเร็จ');
      okList.forEach(function(r) { Logger.log('  - ' + r.name + ' → ' + r.url); });
      if (failList.length > 0) {
        Logger.log('❌ ล้มเหลว: ' + failList.map(function(r){ return r.name + '(' + (r.error||'?') + ')'; }).join(', '));
      }
    }

    // คำนวณลำดับถัดไป: scan ทุกแถวหาค่า max แล้ว +1
    var sequenceNum = 1;
    var seqColIdx2 = headerMap['ลำดับ'];
    if (seqColIdx2 !== undefined && sheet.getLastRow() > 1) {
      var seqVals = sheet.getRange(2, seqColIdx2 + 1, sheet.getLastRow() - 1, 1).getValues();
      for (var sv = 0; sv < seqVals.length; sv++) {
        var n = parseInt(seqVals[sv][0], 10);
        if (!isNaN(n) && n >= sequenceNum) sequenceNum = n + 1;
      }
    }
    Logger.log('sequenceNum = ' + sequenceNum);

    // เก็บ newRow ก่อน append
    var newRow = sheet.getLastRow() + 1;
    Logger.log('newRow = ' + newRow);

    // บันทึกแถว (ใส่ fileUrls เป็น plain text ก่อนเพื่อให้แถวสมบูรณ์)
    var row = buildRowFromData(data, fileNames, fileUrls, headerMap, hLen, sequenceNum);
    sheet.appendRow(row);
    SpreadsheetApp.flush();
    Logger.log('✅ appendRow สำเร็จ');

    // ===== หา column index ของ "URL ไฟล์แนบ" อย่างน่าเชื่อถือ =====
    // วิธีที่ 1: จาก headerMap
    var urlColIdx = headerMap['URL ไฟล์แนบ'];
    // วิธีที่ 2 (fallback): scan header row จาก Sheet โดยตรง
    if (urlColIdx === undefined) {
      Logger.log('⚠️ headerMap ไม่พบ "URL ไฟล์แนบ" — scan header row โดยตรง...');
      var lastCol = sheet.getLastColumn();
      if (lastCol > 0) {
        var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        for (var hi = 0; hi < headerRow.length; hi++) {
          var hVal = String(headerRow[hi] || '').trim();
          if (hVal === 'URL ไฟล์แนบ') {
            urlColIdx = hi;
            Logger.log('✅ พบ "URL ไฟล์แนบ" จาก scan ที่ col index ' + hi + ' (col ' + (hi+1) + ')');
            break;
          }
        }
      }
    }

    Logger.log('urlColIdx (0-based) = ' + urlColIdx);

    // ===== ใส่ HYPERLINK formula =====
    var okFiles = fileResults.filter(function(r) { return r.url && r.url.length > 10; });
    Logger.log('okFiles.length = ' + okFiles.length);

    if (urlColIdx !== undefined && urlColIdx !== null && okFiles.length > 0) {
      var sheetCol = urlColIdx + 1; // 1-based สำหรับ getRange
      var urlCell  = sheet.getRange(newRow, sheetCol);
      // เคลียร์ค่าเดิมก่อน (อาจมี plain text จาก appendRow)
      urlCell.clearContent();

      var formula;
      if (okFiles.length === 1) {
        var safeName = okFiles[0].name.replace(/"/g, '').replace(/'/g, '');
        formula = '=HYPERLINK("' + okFiles[0].url + '","' + safeName + '")';
      } else {
        var safeFirst = okFiles[0].name.replace(/"/g, '').replace(/'/g, '');
        var label     = safeFirst + ' (+' + (okFiles.length - 1) + ')';
        formula       = '=HYPERLINK("' + okFiles[0].url + '","' + label + '")';
        // ใส่ URL ทั้งหมดใน note
        var noteText = okFiles.map(function(r) { return r.name + '\n' + r.url; }).join('\n\n');
        urlCell.setNote(noteText);
      }

      Logger.log('formula = ' + formula);
      urlCell.setFormula(formula);
      SpreadsheetApp.flush();

      // ตรวจสอบว่า formula ถูกใส่จริงไหม
      var checkVal = sheet.getRange(newRow, sheetCol).getFormula();
      Logger.log('✅ formula ใน cell หลัง set: "' + checkVal + '"');
    } else {
      Logger.log('⚠️ ข้าม HYPERLINK: urlColIdx=' + urlColIdx + ', okFiles=' + okFiles.length);
      if (fileUrls) {
        // fallback: ใส่ URL plain text
        if (urlColIdx !== undefined) {
          sheet.getRange(newRow, urlColIdx + 1).setValue(fileUrls);
          SpreadsheetApp.flush();
          Logger.log('  → ใส่ URL plain text แทน');
        }
      }
    }

    // ===== อัปเดต column "ชื่อไฟล์แนบ" =====
    var nameColIdx = headerMap['ชื่อไฟล์แนบ'];
    if (nameColIdx === undefined) {
      // fallback scan
      var lastColN = sheet.getLastColumn();
      if (lastColN > 0) {
        var hRowN = sheet.getRange(1, 1, 1, lastColN).getValues()[0];
        for (var hn = 0; hn < hRowN.length; hn++) {
          if (String(hRowN[hn]||'').trim() === 'ชื่อไฟล์แนบ') {
            nameColIdx = hn;
            break;
          }
        }
      }
    }
    if (nameColIdx !== undefined && fileNames) {
      sheet.getRange(newRow, nameColIdx + 1).setValue(fileNames);
      SpreadsheetApp.flush();
      Logger.log('✅ ใส่ชื่อไฟล์แนบ: ' + fileNames);
    }

    Logger.log('=== saveTrainingRequest SUCCESS row=' + newRow + ' ===');
    return {
      success    : true,
      message    : 'บันทึกสำเร็จ',
      fileResults: fileResults,
      fileCount  : fileResults.length
    };

  } catch (error) {
    Logger.log('❌ saveTrainingRequest ERROR: ' + error.toString());
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

/**
 * อัปเดต (แก้ไขย้อนหลัง) คำขออบรมที่บันทึกไว้แล้ว
 * รองรับ: แก้เลขที่ ชม.5 / บันทึก, ค่าใช้จ่าย, สถานะเอกสาร, แนบไฟล์เพิ่มเติม, หมายเหตุ
 */
function updateTrainingRequest(updates, files) {
  try {
    var info      = getOrInitTrainingSheet();
    var sheet     = info.sheet;
    var hMap      = info.headerMap;

    Logger.log('=== updateTrainingRequest START rowIndex=' + updates.rowIndex + ' ===');

    // ค้นหาแถวที่ถูกต้อง — ใช้ rowIndex ที่ส่งมาถ้ามี มิฉะนั้น scan ด้วย timestamp/id
    var targetRow = -1;

    if (updates.rowIndex && updates.rowIndex >= 2) {
      targetRow = updates.rowIndex;
      Logger.log('ใช้ rowIndex โดยตรง: ' + targetRow);
    } else {
      // fallback: scan ทุกแถว หา timestamp ที่ตรงกับ savedAt
      var allData = sheet.getDataRange().getValues();
      var tsColIdx = hMap['Timestamp'];
      var nameColIdx2 = hMap['ชื่อ-สกุล'];
      for (var i = 1; i < allData.length; i++) {
        var rowName = String(allData[i][nameColIdx2] || '').trim();
        if (rowName === String(updates.name || '').trim()) {
          targetRow = i + 1;
          break;
        }
      }
      Logger.log('scan หาแถว: ' + targetRow);
    }

    if (targetRow < 2) {
      return { success: false, message: 'ไม่พบแถวที่ต้องการแก้ไข' };
    }

    // Helper: เขียนค่าลง cell ถ้า column มีอยู่
    var setCell = function(headerName, value) {
      var ci = hMap[headerName];
      if (ci !== undefined) {
        sheet.getRange(targetRow, ci + 1).setValue(value);
      }
    };

    // อัปเดตเลขที่ ชม.5 และบันทึก
    if (updates.hm5  !== undefined) setCell('เลขที่ ชม.5',    updates.hm5);
    if (updates.memo !== undefined) setCell('เลขที่บันทึก',   updates.memo);

    // อัปเดตค่าใช้จ่าย
    if (updates.costReg    !== undefined) setCell('ค่าลงทะเบียน',   updates.costReg);
    if (updates.costTravel !== undefined) setCell('ค่าเบี้ยเลี้ยง',   updates.costTravel);
    if (updates.costHotel  !== undefined) setCell('ค่าที่พัก',        updates.costHotel);
    if (updates.costTrans  !== undefined) setCell('ค่าพาหนะ',         updates.costTrans);
    if (updates.costProj   !== undefined) setCell('เงินโครงการ',      updates.costProj);
    if (updates.totalCost  !== undefined) setCell('รวมค่าใช้จ่าย',    updates.totalCost);
    if (updates.budget     !== undefined) setCell('แหล่งงบประมาณ',    updates.budget);

    // อัปเดตสถานะเอกสาร
    setCell('เอกสารสรุปผล',   updates.docSummary  ? '\u2713' : '');
    setCell('สัญญา',          updates.docContract ? '\u2713' : '');
    setCell('วุฒิบัตร',       updates.docCert     ? '\u2713' : '');
    setCell('ใบเสร็จ',        updates.docReceipt  ? '\u2713' : '');
    setCell('บันทึก ชม.5',   updates.docHm5      ? '\u2713' : '');
    setCell('บันทึกขออนุมัติ',updates.docApproval ? '\u2713' : '');
    setCell('บันทึกความ+ชม.5',updates.docCombined ? '\u2713' : '');

    // อัปเดตหมายเหตุ
    if (updates.note !== undefined) {
      var existNote = String(sheet.getRange(targetRow, hMap['หมายเหตุ'] + 1).getValue() || '').trim();
      var newNote = existNote
        ? existNote + '\n[แก้ไข ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') + '] ' + (updates.note || '')
        : (updates.note || '');
      setCell('หมายเหตุ', newNote);
    }

    // ── helper: ดึง URL ทั้งหมดจาก cell (formula + note + plain value) ──
    var extractUrlsFromCell = function(cell) {
      var out = [];
      var formula = cell.getFormula();
      var note    = cell.getNote();
      var value   = String(cell.getValue() || '').trim();

      // 1. จาก HYPERLINK formula  =HYPERLINK("https://...","label")
      if (formula) {
        var mFm = formula.match(/https?:\/\/[^"\s]+/g);
        if (mFm) out = out.concat(mFm);
      }
      // 2. จาก note (เก็บ URL ทุกตัวแบบ plain ต่อบรรทัด)
      if (note) {
        note.split('\n').forEach(function(line) {
          var t = line.trim();
          if (t.indexOf('http') === 0) out.push(t);
        });
      }
      // 3. fallback: ถ้า value เป็น URL ตรงๆ
      if (out.length === 0 && value.indexOf('http') === 0) {
        out.push(value);
      }
      // deduplicate
      var seen = {};
      return out.filter(function(u) {
        if (!u || seen[u]) return false;
        seen[u] = true;
        return true;
      });
    };

    // ── helper: เขียน URL ทั้งหมดลง cell ให้ตรงกับ pattern เดิม ──
    // formula = =HYPERLINK("url_ล่าสุด","ชื่อไฟล์ล่าสุด (+N)") 
    // note    = url1\nurl2\nurl3\n...  (1 URL ต่อบรรทัด)
    var writeUrlsToCell = function(cell, fileObjArr, existingUrlArr) {
      // fileObjArr = [{name, url}, ...] ไฟล์ใหม่
      // existingUrlArr = [url, ...] URL เดิม
      var allUrls = existingUrlArr.slice();
      fileObjArr.forEach(function(f) {
        if (f.url && allUrls.indexOf(f.url) === -1) allUrls.push(f.url);
      });
      if (allUrls.length === 0) return;

      cell.clearContent();
      cell.clearNote();

      // formula ชี้ที่ไฟล์ใหม่สุด
      var lastNew   = fileObjArr[fileObjArr.length - 1];
      var safeName  = lastNew.name.replace(/"/g, '').replace(/'/g, '');
      var label     = allUrls.length > 1
        ? safeName + ' (+' + (allUrls.length - 1) + ')'
        : safeName;
      cell.setFormula('=HYPERLINK("' + lastNew.url + '","' + label + '")');

      // note เก็บทุก URL แบบ 1 URL ต่อบรรทัด
      if (allUrls.length > 1) {
        cell.setNote(allUrls.join('\n'));
      }
      Logger.log('\u2705 writeUrlsToCell: ' + allUrls.length + ' URLs');
    };

    // อัปโหลดไฟล์เพิ่มเติม (ถ้ามี)
    var fileResults = [];
    if (files && files.length > 0) {
      Logger.log('อัพโหลดไฟล์เพิ่มเติม ' + files.length + ' ไฟล์...');
      fileResults = saveFilesToDrive(files);
      Logger.log('saveFilesToDrive คืน ' + fileResults.length + ' ผลลัพธ์');

      var newOk = fileResults.filter(function(r) {
        Logger.log('  file: ' + r.name + ' url=' + (r.url || 'EMPTY') + ' err=' + (r.error || '-'));
        return r.url && r.url.indexOf('http') === 0;
      });
      Logger.log('newOk count=' + newOk.length);

      if (newOk.length > 0) {
        // อ่านชื่อไฟล์เดิม
        var nameColIdx3 = hMap['ชื่อไฟล์แนบ'];
        var oldNames = nameColIdx3 !== undefined
          ? String(sheet.getRange(targetRow, nameColIdx3 + 1).getValue() || '').trim()
          : '';

        // อ่าน URL เดิมจาก cell
        var urlColIdx3 = hMap['URL ไฟล์แนบ'];
        var existUrls  = [];
        if (urlColIdx3 !== undefined) {
          existUrls = extractUrlsFromCell(sheet.getRange(targetRow, urlColIdx3 + 1));
          Logger.log('existUrls count=' + existUrls.length);
        }

        // อัปเดตชื่อไฟล์
        var addNames = newOk.map(function(r) { return r.name; }).join(', ');
        var allNames = oldNames ? oldNames + ', ' + addNames : addNames;
        if (nameColIdx3 !== undefined) {
          sheet.getRange(targetRow, nameColIdx3 + 1).setValue(allNames);
          Logger.log('\u2705 เขียนชื่อไฟล์: ' + allNames);
        }

        // เขียน URL ลง cell
        if (urlColIdx3 !== undefined) {
          writeUrlsToCell(
            sheet.getRange(targetRow, urlColIdx3 + 1),
            newOk,
            existUrls
          );
        }
      } else {
        Logger.log('\u26a0 ไม่มีไฟล์อัพโหลดสำเร็จ — ข้ามการเขียน URL');
      }
    }

    SpreadsheetApp.flush();
    Logger.log('\u2705 flush เสร็จ');

    // ── คืน finalNames / finalUrls ให้ client ──
    var finalNames = '';
    var finalUrls  = '';
    var nameColF = hMap['ชื่อไฟล์แนบ'];
    var urlColF  = hMap['URL ไฟล์แนบ'];
    if (nameColF !== undefined) {
      finalNames = String(sheet.getRange(targetRow, nameColF + 1).getValue() || '').trim();
    }
    if (urlColF !== undefined) {
      var finalUrlArr = extractUrlsFromCell(sheet.getRange(targetRow, urlColF + 1));
      finalUrls = finalUrlArr.join('\n');
    }
    Logger.log('finalNames=' + finalNames);
    Logger.log('finalUrls=' + finalUrls.substring(0, 100));

    Logger.log('=== updateTrainingRequest SUCCESS row=' + targetRow + ' ===');
    return {
      success    : true,
      message    : 'แก้ไขสำเร็จ (แถว ' + targetRow + ')',
      fileResults: fileResults,
      fileNames  : finalNames,
      fileUrls   : finalUrls
    };

  } catch (error) {
    Logger.log('\u274c updateTrainingRequest ERROR: ' + error.toString());
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

/**
 * แก้ไข hyperlink ย้อนหลัง — สำหรับแถวที่มี URL ใน "URL ไฟล์แนบ" แต่ยังไม่เป็น HYPERLINK formula
 * รันจาก Script Editor เพื่อแก้แถวเก่าทั้งหมด
 */
function fixMissingHyperlinks() {
  try {
    var info    = getOrInitTrainingSheet();
    var sheet   = info.sheet;
    var hMap    = info.headerMap;

    // หา column index
    var urlColIdx  = hMap['URL ไฟล์แนบ'];
    var nameColIdx = hMap['ชื่อไฟล์แนบ'];

    if (urlColIdx === undefined) {
      // scan header row
      var lastCol = sheet.getLastColumn();
      var hRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      for (var i = 0; i < hRow.length; i++) {
        if (String(hRow[i]||'').trim() === 'URL ไฟล์แนบ')  urlColIdx  = i;
        if (String(hRow[i]||'').trim() === 'ชื่อไฟล์แนบ') nameColIdx = i;
      }
    }

    if (urlColIdx === undefined) {
      Logger.log('\u274c ไม่พบ column "URL ไฟล์แนบ"');
      return;
    }

    Logger.log('urlColIdx=' + urlColIdx + ' (col ' + (urlColIdx+1) + ')');

    var lastRow = sheet.getLastRow();
    var fixed   = 0;
    var skipped = 0;

    for (var r = 2; r <= lastRow; r++) {
      var cell    = sheet.getRange(r, urlColIdx + 1);
      var formula = cell.getFormula();
      var value   = String(cell.getValue() || '').trim();

      // ถ้ามี formula อยู่แล้ว → ข้าม
      if (formula && formula.indexOf('HYPERLINK') !== -1) {
        skipped++;
        continue;
      }

      // ถ้า value เป็น URL (เริ่มด้วย https://)
      if (value && value.indexOf('https://') === 0) {
        // หาชื่อไฟล์ (จาก ชื่อไฟล์แนบ column)
        var fName = '';
        if (nameColIdx !== undefined) {
          fName = String(sheet.getRange(r, nameColIdx + 1).getValue() || '').trim();
        }
        // ถ้ามีหลาย URL (คั่นด้วย \n) ใช้ URL แรก
        var urls  = value.split('\n');
        var names = fName ? fName.split(', ') : [];
        var firstUrl  = urls[0].trim();
        var firstName = (names[0] || 'ไฟล์แนบ').replace(/"/g,'').replace(/'/g,'');
        var hlFormula = '=HYPERLINK("' + firstUrl + '","' + firstName + '")';
        cell.clearContent();
        cell.setFormula(hlFormula);
        if (urls.length > 1) {
          var note = urls.map(function(u, ui) {
            return (names[ui] || u) + '\n' + u;
          }).join('\n\n');
          cell.setNote(note);
        }
        fixed++;
        Logger.log('\u2705 Row ' + r + ': "' + firstName + '" → ' + firstUrl.substring(0,60));
      }
    }

    SpreadsheetApp.flush();
    Logger.log('=== fixMissingHyperlinks เสร็จสิ้น: fixed=' + fixed + ' skipped=' + skipped + ' ===');

  } catch(e) {
    Logger.log('\u274c fixMissingHyperlinks error: ' + e.toString());
  }
}
function getRecentTrainingRequests(limit) {
  try {
    limit = limit || 10;
    var info    = getOrInitTrainingSheet();
    var sheet   = info.sheet;
    var hMap    = info.headerMap;
    var lastRow = sheet.getLastRow();

    Logger.log('getRecentTrainingRequests: sheet=' + TRAINING_REQUEST_SHEET +
               ', lastRow=' + lastRow +
               ', headerCount=' + info.headers.length);

    if (lastRow <= 1) {
      Logger.log('getRecentTrainingRequests: Sheet ว่างเปล่า');
      return { success: true, data: [] };
    }

    var dataRows = lastRow - 1; // ไม่รวม header

    // ── อ่านข้อมูลทั้งหมดในครั้งเดียว (batch) ──
    var allData = sheet.getRange(2, 1, dataRows, info.headers.length).getValues();

    // ── อ่าน formula ของ column "URL ไฟล์แนบ" แบบ batch ──
    var urlColIdx = hMap['URL ไฟล์แนบ'];
    var urlFormulas = [];
    var urlNotes    = [];
    if (urlColIdx !== undefined) {
      var urlRange = sheet.getRange(2, urlColIdx + 1, dataRows, 1);
      var rawFormulas = urlRange.getFormulas();   // batch: 1 API call
      var rawNotes    = urlRange.getNotes();      // batch: 1 API call
      for (var ri = 0; ri < dataRows; ri++) {
        urlFormulas.push(rawFormulas[ri][0] || '');
        urlNotes.push(rawNotes[ri][0]    || '');
      }
    }

    // ── เรียงตาม Timestamp จริง (ล่าสุดก่อน) ──
    var tsColIdx = hMap['Timestamp'];
    var sortIdx  = [];
    for (var si = 0; si < dataRows; si++) {
      var tsVal = (tsColIdx !== undefined) ? allData[si][tsColIdx] : null;
      var tsMs  = (tsVal instanceof Date) ? tsVal.getTime()
                : (tsVal ? new Date(tsVal).getTime() : 0);
      if (isNaN(tsMs)) tsMs = 0;
      // seqNum สำรอง
      var seqColIdxTmp = hMap['ลำดับ'];
      var seqV = seqColIdxTmp !== undefined ? (parseFloat(allData[si][seqColIdxTmp]) || 0) : si;
      sortIdx.push({ i: si, tsMs: tsMs, seq: seqV });
    }
    // Primary: Timestamp มากสุดก่อน | Secondary: seqNum มากสุดก่อน
    sortIdx.sort(function(a, b) {
      if (b.tsMs !== a.tsMs) return b.tsMs - a.tsMs;
      return b.seq - a.seq;
    });

    var results = [];

    for (var si2 = 0; si2 < sortIdx.length && results.length < limit; si2++) {
      var i = sortIdx[si2].i;
      var r = allData[i];

      var hMapRef = hMap;
      var rowRef  = r;

      var getVal = function(key) {
        var ci = hMapRef[key];
        return ci !== undefined ? rowRef[ci] : undefined;
      };
      var getByIdx = function(idx) {
        return idx < rowRef.length ? rowRef[idx] : '';
      };
      var gStr = function(key) {
        var v = getVal(key);
        return String(v === null || v === undefined ? '' : v).trim();
      };
      var gDate = function(key) {
        var v = getVal(key);
        if (!v) return '';
        try {
          var d = (v instanceof Date) ? v : new Date(v);
          return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
        } catch(ex) { return ''; }
      };
      var gFloat = function(key) { return parseFloat(getVal(key)) || 0; };
      var gTs = function(key) {
        var v = getVal(key);
        if (!v) return '';
        try {
          var d = (v instanceof Date) ? v : new Date(v);
          return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm');
        } catch(ex) { return ''; }
      };

      var name = gStr('ชื่อ-สกุล') || gStr('ชื่อ') || String(getByIdx(1) || '').trim();
      if (!name) continue;

      var topic = gStr('เรื่อง/หลักสูตร') || gStr('หลักสูตร') || gStr('เรื่อง') || String(getByIdx(8) || '').trim();
      var savedAt = gTs('Timestamp') || (getByIdx(0) ? (function(){
        try { return Utilities.formatDate(new Date(getByIdx(0)),'Asia/Bangkok','yyyy-MM-dd HH:mm'); } catch(ex){ return ''; }
      })() : '');

      // ── ดึง URL จาก batch arrays ──
      var fileUrlsStr = '';
      if (urlColIdx !== undefined) {
        var formula = urlFormulas[i] || '';
        var note    = urlNotes[i]    || '';
        var urls    = [];

        // 1. จาก HYPERLINK formula
        if (formula && formula.toUpperCase().indexOf('HYPERLINK') !== -1) {
          var mF = formula.match(/https?:\/\/[^"'\s]+/g);
          if (mF) urls = urls.concat(mF);
        }
        // 2. จาก cell note (กรณีไฟล์หลายตัว)
        if (note) {
          var mN = note.match(/https?:\/\/[^\s\n]+/g);
          if (mN) urls = urls.concat(mN);
        }
        // 3. fallback: plain URL ใน cell value
        if (urls.length === 0) {
          var plain = String(rowRef[urlColIdx] || '').trim();
          if (plain.indexOf('http') === 0) urls.push(plain);
        }
        // deduplicate
        var seen = {};
        urls = urls.filter(function(u) { if (seen[u]) return false; seen[u] = true; return true; });
        fileUrlsStr = urls.join('\n');
      }

      // อ่าน seqNum จากคอลัมน์ "ลำดับ" จริงในชีท (ไม่ hardcode BASE_START)
      var seqRaw = gStr('ลำดับ');
      var seqNum = seqRaw !== '' ? seqRaw : String(i + 2); // i+2 = row จริง (1-based + skip header)
      results.push({
        savedAt    : savedAt,
        tsMs       : sortIdx[si2].tsMs,
        seqNum     : seqNum,
        name       : name,
        position   : gStr('ตำแหน่ง'),
        mission    : gStr('กลุ่มภารกิจ'),
        dept       : gStr('กลุ่มงาน'),
        topic      : topic,
        startDate  : gDate('วันที่เริ่ม'),
        endDate    : gDate('วันที่สิ้นสุด'),
        venue      : gStr('สถานที่'),
        province   : gStr('จังหวัด'),
        totalCost  : gFloat('รวมค่าใช้จ่าย'),
        budget     : gStr('แหล่งงบประมาณ'),
        fileNames  : gStr('ชื่อไฟล์แนบ'),
        fileUrls   : fileUrlsStr,
        // [FIX] รองรับทั้ง '✓' และชื่อไฟล์แนบ (เช่น 'สรุปงาน.pdf') — ถือว่ามีเอกสารเมื่อเซลล์ไม่ว่าง
        docSummary : gStr('เอกสารสรุปผล').length > 0,
        docContract: gStr('สัญญา').length        > 0,
        docCert    : gStr('วุฒิบัตร').length     > 0,
        docReceipt : gStr('ใบเสร็จ').length      > 0,
        docHm5     : gStr('บันทึก ชม.5').length  > 0,
        docApproval: gStr('บันทึกขออนุมัติ').length > 0,
        docCombined: gStr('บันทึกความ+ชม.5').length > 0,
        lineId     : gStr('LINE ID'),
        recorder   : gStr('ผู้บันทึก'),
        hm5        : gStr('เลขที่ ชม.5'),
        memo       : gStr('เลขที่บันทึก'),
        note       : gStr('หมายเหตุ'),
        rowIndex   : i + 2,   // +2 เพราะ 1-based และข้ามแถว header
        costReg    : gFloat('ค่าลงทะเบียน'),
        costTravel : gFloat('ค่าเบี้ยเลี้ยง'),
        costHotel  : gFloat('ค่าที่พัก'),
        costTrans  : gFloat('ค่าพาหนะ'),
        costProj   : gFloat('เงินโครงการ')
      });
    }

    Logger.log('\u2705 getRecentTrainingRequests: คืน ' + results.length + ' รายการ');
    return { success: true, data: results };

  } catch (error) {
    Logger.log('\u274c getRecentTrainingRequests: ' + error.toString());
    return { success: false, message: 'getRecentTrainingRequests error: ' + error.toString() };
  }
}




/**
 * ============================================
 * === backfillFileUrls() ===
 * สแกนทุกแถวในชีท "อบรม ปีงบประมาณ 2569"
 * อ่าน HYPERLINK formula จากคอลัมน์ "ชื่อไฟล์แนบ"
 * แล้วเขียน URL plain text ลงคอลัมน์ "URL ไฟล์แนบ"
 * (สำหรับแถวที่ยังไม่มี URL ในคอลัมน์นั้น)
 *
 * ▶ รันจาก Apps Script Editor ครั้งเดียว เพื่อ backfill ข้อมูลเก่า
 * ▶ หรือเรียกผ่าน Web App ก็ได้ (ใช้เวลาตามจำนวนแถว)
 * ============================================
 */
function backfillFileUrls() {
  try {
    var info    = getOrInitTrainingSheet();
    var sheet   = info.sheet;
    var hMap    = info.headerMap;
    var lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      Logger.log('backfillFileUrls: Sheet ว่าง ไม่มีข้อมูล');
      return { success: true, updated: 0, message: 'Sheet ว่าง' };
    }

    var dataRows   = lastRow - 1;
    var nameColIdx = hMap['ชื่อไฟล์แนบ'];
    var urlColIdx  = hMap['URL ไฟล์แนบ'];

    if (nameColIdx === undefined || urlColIdx === undefined) {
      Logger.log('backfillFileUrls: ไม่พบคอลัมน์ "ชื่อไฟล์แนบ" หรือ "URL ไฟล์แนบ"');
      return { success: false, message: 'ไม่พบคอลัมน์ที่ต้องการ' };
    }

    Logger.log('backfillFileUrls: สแกน ' + dataRows + ' แถว, nameCol=' + (nameColIdx+1) + ', urlCol=' + (urlColIdx+1));

    // อ่าน formula ของ "ชื่อไฟล์แนบ" (batch)
    var nameRange    = sheet.getRange(2, nameColIdx + 1, dataRows, 1);
    var nameFormulas = nameRange.getFormulas();
    var nameNotes    = nameRange.getNotes();
    var nameValues   = nameRange.getValues();

    // อ่าน URL ปัจจุบันใน "URL ไฟล์แนบ" (batch)
    var urlRange    = sheet.getRange(2, urlColIdx + 1, dataRows, 1);
    var urlFormulas = urlRange.getFormulas();
    var urlValues   = urlRange.getValues();

    var updatedCount = 0;
    var skippedCount = 0;
    var noUrlCount   = 0;

    // เตรียม array สำหรับ batch write (copy ค่าปัจจุบันก่อน ไม่แตะแถวที่มีข้อมูลอยู่แล้ว)
    var outputValues = urlValues.map(function(row) { return [row[0]]; });

    for (var i = 0; i < dataRows; i++) {
      // ถ้า URL ไฟล์แนบมีข้อมูลอยู่แล้ว — ข้ามเลย
      var existingFormula = urlFormulas[i][0] || '';
      var existingValue   = String(urlValues[i][0] || '').trim();
      if (existingFormula !== '' || existingValue !== '') {
        skippedCount++;
        continue;
      }

      // ดึง URL จากคอลัมน์ "ชื่อไฟล์แนบ"
      var urls = [];

      // 1. จาก HYPERLINK formula เช่น =HYPERLINK("https://drive.google.com/...","389.pdf")
      var nameFm = nameFormulas[i][0] || '';
      if (nameFm && nameFm.toUpperCase().indexOf('HYPERLINK') !== -1) {
        var mFm = nameFm.match(/https?:\/\/[^"'\s]+/g);
        if (mFm) urls = urls.concat(mFm);
      }

      // 2. จาก note ของ "ชื่อไฟล์แนบ" (กรณีไฟล์หลายตัว เก็บ URL ใน note)
      var nameNote = nameNotes[i][0] || '';
      if (nameNote) {
        var mNote = nameNote.match(/https?:\/\/[^\s\n]+/g);
        if (mNote) urls = urls.concat(mNote);
      }

      // 3. fallback: ถ้า cell value เป็น URL ตรงๆ
      var nameVal = String(nameValues[i][0] || '').trim();
      if (urls.length === 0 && nameVal.indexOf('http') === 0) {
        urls.push(nameVal);
      }

      if (urls.length === 0) {
        noUrlCount++;
        continue; // ไม่มี URL จริงๆ (แถวที่ยังไม่มีไฟล์แนบ)
      }

      // deduplicate
      var seen = {};
      urls = urls.filter(function(u) {
        if (!u || seen[u]) return false;
        seen[u] = true;
        return true;
      });

      // เขียนลง output array — batch write ครั้งเดียวด้านล่าง
      outputValues[i] = [urls.join('\n')];
      updatedCount++;

      Logger.log('Row ' + (i + 2) + ': ' + urls.length + ' URL(s) → ' + urls[0]);
    }

    // Batch write ครั้งเดียว — ประหยัด API call มากกว่า setRange ใน loop
    if (updatedCount > 0) {
      urlRange.setValues(outputValues);
    }

    SpreadsheetApp.flush();

    var msg = 'backfillFileUrls เสร็จ: อัปเดต ' + updatedCount + ' แถว, ข้าม ' + skippedCount + ' แถว (มีข้อมูลอยู่แล้ว), ไม่มีไฟล์ ' + noUrlCount + ' แถว';
    Logger.log('✅ ' + msg);
    return { success: true, updated: updatedCount, skipped: skippedCount, noUrl: noUrlCount, message: msg };

  } catch (e) {
    Logger.log('❌ backfillFileUrls error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}


/**
 * Helper: alias สำหรับ backward compat — เรียก debugAndFixTrainingSheet() แทน
 */
function debugTrainingSheetHeaders() {
  debugAndFixTrainingSheet();
}

/**
 * ลบไฟล์จาก Google Drive
 */
function deleteDriveFile(fileId) {
  // Validate fileId before calling Drive API
  if (!fileId) {
    return { success: false, message: 'ไม่ได้ระบุ File ID' };
  }
  var id = String(fileId).trim();
  if (id === '' || id === 'undefined' || id === 'null') {
    return { success: false, message: 'File ID ไม่ถูกต้อง: "' + id + '"' };
  }
  // Google Drive File IDs are typically 28-44 alphanumeric chars
  if (!/^[a-zA-Z0-9_\-]{10,60}$/.test(id)) {
    return { success: false, message: 'File ID ไม่ถูกรูปแบบ: "' + id + '"' };
  }
  try {
    var file = DriveApp.getFileById(id);
    file.setTrashed(true);
    Logger.log('\u2705 ลบไฟล์สำเร็จ: ' + id);
    return { success: true, message: 'ลบไฟล์สำเร็จ' };
  } catch (e) {
    Logger.log('\u274c deleteDriveFile id=' + id + ' error: ' + e.toString());
    return { success: false, message: 'ลบไม่สำเร็จ: ' + e.message };
  }
}


