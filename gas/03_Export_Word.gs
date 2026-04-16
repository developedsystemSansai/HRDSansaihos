// ============================================
// === 9. exportPersonalSelectedToWord(exportData) ===
// สร้างไฟล์ Google Sheets สรุปรายการอบรมส่วนบุคคล (ใช้ SpreadsheetApp แทน DocumentApp)
// ============================================
function exportPersonalSelectedToWord(exportData) {
  try {
    if (!exportData || !exportData.summary || !exportData.registrations) {
      return { success: false, message: 'ข้อมูลไม่ครบถ้วน' };
    }

    var s   = exportData.summary;
    var reg = exportData.registrations;
    var fileName = 'รายงานอบรม_' + (s.fullName || 'บุคลากร') + '_' +
                   Utilities.formatDate(new Date(), 'Asia/Bangkok', 'ddMMyyyy');

    var thaiMonths = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
    function fmtDateP(ds) {
      if (!ds) return '-';
      try {
        var d = ds instanceof Date ? ds : new Date(ds);
        if (isNaN(d.getTime())) return String(ds);
        return d.getDate() + ' ' + thaiMonths[d.getMonth()] + ' ' + (d.getFullYear() + 543);
      } catch(e) { return String(ds); }
    }

    // สร้าง Spreadsheet ใหม่
    var ss   = SpreadsheetApp.create(fileName);
    var sh   = ss.getActiveSheet();
    sh.setName('รายงาน');

    // ── หัว ──
    sh.getRange('A1').setValue('รายงานสรุปการพัฒนาศักยภาพบุคลากร โรงพยาบาลสันทราย');
    sh.getRange('A1').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange('A1:E1').merge();

    sh.getRange('A2').setValue('ชื่อ-สกุล: ' + (s.fullName || '-'));
    sh.getRange('A3').setValue('ตำแหน่ง: ' + (s.position || '-'));
    sh.getRange('A4').setValue('กลุ่มภารกิจ: ' + (s.missionGroup || '-') + '   กลุ่มงาน: ' + (s.department || '-'));
    sh.getRange('A5').setValue('จำนวนรายการที่เลือก: ' + reg.length + ' รายการ   รวม: ' + (s.totalHours || 0).toFixed(1) + ' ชั่วโมง');
    sh.getRange('A2:E5').setFontSize(12);

    // ── header ตาราง ──
    var hdrRow = 7;
    var headers = ['ลำดับ', 'หลักสูตร/เรื่อง', 'วันที่', 'สถานที่', 'ชม.'];
    sh.getRange(hdrRow, 1, 1, 5).setValues([headers])
      .setFontWeight('bold').setFontSize(12)
      .setBackground('#1e3a8a').setFontColor('#ffffff')
      .setHorizontalAlignment('center');

    // ── ข้อมูล ──
    var rows = [];
    reg.forEach(function(r, idx) {
      var hrsDisplay = String(r.hours || 0);
      if (String(r.category || '').trim() === NO_CREDIT_CAT) hrsDisplay += ' *';
      var dateStr = fmtDateP(r.date) + (r.endDate ? ' - ' + fmtDateP(r.endDate) : '');
      rows.push([
        idx + 1,
        String(r.topic || '-'),
        dateStr,
        String(r.location || '-'),
        hrsDisplay
      ]);
    });

    if (rows.length > 0) {
      sh.getRange(hdrRow + 1, 1, rows.length, 5).setValues(rows).setFontSize(12);
      // zebra stripe
      for (var ri = 0; ri < rows.length; ri++) {
        if (ri % 2 === 1) {
          sh.getRange(hdrRow + 1 + ri, 1, 1, 5).setBackground('#eff6ff');
        }
      }
    }

    // ── สรุป ──
    var sumRow = hdrRow + rows.length + 2;
    sh.getRange(sumRow, 1).setValue('รวมทั้งสิ้น ' + reg.length + ' รายการ   ' + (s.totalHours || 0).toFixed(1) + ' ชั่วโมง')
      .setFontWeight('bold').setFontSize(12);
    sh.getRange(sumRow, 1, 1, 5).merge();

    sh.getRange(sumRow + 1, 1).setValue(
      '* รายการที่มีเครื่องหมาย * คือ ประชุมคณะกรรมการ/คณะทำงาน/ประชุมชี้แจงฯ บันทึกชั่วโมงไว้เป็นหลักฐาน แต่ไม่นับรวมในหน่วยกิตสะสม'
    ).setFontStyle('italic').setFontSize(11).setFontColor('#64748b');
    sh.getRange(sumRow + 1, 1, 1, 5).merge();

    // ── ลายเซ็น ──
    var sigRow = sumRow + 3;
    sh.getRange(sigRow, 1).setValue('(ลงชื่อ)....................................................');
    sh.getRange(sigRow + 1, 1).setValue('     (' + (s.fullName || '...........................') + ')');
    sh.getRange(sigRow + 2, 1).setValue('วันที่ ' + fmtDateP(new Date()));

    // ── column width ──
    sh.setColumnWidth(1, 60);
    sh.setColumnWidth(2, 280);
    sh.setColumnWidth(3, 130);
    sh.setColumnWidth(4, 180);
    sh.setColumnWidth(5, 60);

    // ── เส้นขอบตาราง ──
    if (rows.length > 0) {
      sh.getRange(hdrRow, 1, rows.length + 1, 5)
        .setBorder(true, true, true, true, true, true);
    }

    // ย้ายไปยัง ATTACHMENTS folder ถ้ากำหนดไว้
    try {
      var file = DriveApp.getFileById(ss.getId());
      var folder = DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID);
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    } catch(fe) {
      Logger.log('⚠️ ย้ายไฟล์ไป folder ไม่ได้ (ไม่ใช่ error หลัก): ' + fe.toString());
    }

    var url = ss.getUrl();
    Logger.log('✅ exportPersonalSelectedToWord: ' + url);
    return { success: true, url: url, id: ss.getId() };

  } catch (e) {
    Logger.log('❌ exportPersonalSelectedToWord: ' + e);
    return { success: false, message: e.toString() };
  }
}

// ============================================
// === 10. exportDashboardToWord(year, month, missionGroup) ===
// สร้างไฟล์ Google Sheets สรุปภาพรวม Dashboard (ใช้ SpreadsheetApp แทน DocumentApp)
// ============================================
function exportDashboardToWord(year, month, missionGroup) {
  try {
    var result = getFilteredDashboard(year, month, missionGroup);
    if (!result.success) return result;

    var reg  = result.data.registrations;
    var summ = result.data.summary;

    var titleParts = [];
    if (year)         titleParts.push('ปี ' + (parseInt(year) + 543));
    if (month)        titleParts.push('เดือน ' + month);
    if (missionGroup) titleParts.push(missionGroup);
    var titleStr = titleParts.length ? titleParts.join(' / ') : 'ทั้งหมด';

    var fileName = 'สรุป Dashboard_' + titleStr + '_' +
                   Utilities.formatDate(new Date(), 'Asia/Bangkok', 'ddMMyyyy');

    // สร้าง Spreadsheet ใหม่
    var ss = SpreadsheetApp.create(fileName);
    var sh = ss.getActiveSheet();
    sh.setName('Dashboard');

    // ── หัว ──
    sh.getRange('A1').setValue('รายงานสรุปการพัฒนาศักยภาพบุคลากร โรงพยาบาลสันทราย — ' + titleStr);
    sh.getRange('A1').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange('A1:F1').merge();

    sh.getRange('A2').setValue('จำนวนรายการทั้งหมด: ' + summ.numRegistrations + ' รายการ   รวมชั่วโมง: ' + (summ.totalHours || 0).toFixed(1) + ' ชั่วโมง');
    sh.getRange('A2').setFontSize(12);
    sh.getRange('A2:F2').merge();

    if (reg.length === 0) {
      sh.getRange('A4').setValue('ไม่พบข้อมูลตามเงื่อนไขที่กำหนด').setFontStyle('italic');
    } else {
      // สรุปรายบุคคล
      var byPerson = {};
      reg.forEach(function(r) {
        var k = r.fullName;
        if (!byPerson[k]) byPerson[k] = { name: k, position: r.position || '-', dept: r.department || '-', count: 0, hours: 0 };
        byPerson[k].count++;
        byPerson[k].hours += r.hours || 0;
      });
      var personList = Object.values(byPerson).sort(function(a, b) { return b.hours - a.hours; });

      sh.getRange('A4').setValue('สรุปรายบุคคล (' + personList.length + ' คน)')
        .setFontWeight('bold').setFontSize(14);
      sh.getRange('A4:F4').merge();

      var hdrRow = 5;
      var headers = ['ลำดับ', 'ชื่อ-สกุล', 'ตำแหน่ง', 'กลุ่มงาน', 'จำนวนครั้ง', 'รวม ชม.'];
      sh.getRange(hdrRow, 1, 1, 6).setValues([headers])
        .setFontWeight('bold').setFontSize(12)
        .setBackground('#1e3a8a').setFontColor('#ffffff')
        .setHorizontalAlignment('center');

      var pRows = personList.map(function(p, i) {
        return [i + 1, p.name, p.position, p.dept, p.count, parseFloat(p.hours.toFixed(1))];
      });

      sh.getRange(hdrRow + 1, 1, pRows.length, 6).setValues(pRows).setFontSize(12);

      // zebra stripe
      for (var ri = 0; ri < pRows.length; ri++) {
        if (ri % 2 === 1) sh.getRange(hdrRow + 1 + ri, 1, 1, 6).setBackground('#eff6ff');
      }

      // เส้นขอบ
      sh.getRange(hdrRow, 1, pRows.length + 1, 6).setBorder(true, true, true, true, true, true);

      // column width
      sh.setColumnWidth(1, 60);
      sh.setColumnWidth(2, 220);
      sh.setColumnWidth(3, 180);
      sh.setColumnWidth(4, 180);
      sh.setColumnWidth(5, 90);
      sh.setColumnWidth(6, 80);

      // ── sheet 2: รายละเอียดทั้งหมด ──
      var sh2 = ss.insertSheet('รายละเอียด');
      var h2 = ['ลำดับ', 'ชื่อ-สกุล', 'ตำแหน่ง', 'กลุ่มงาน', 'หลักสูตร/เรื่อง', 'วันที่', 'ถึงวันที่', 'สถานที่', 'ชม.', 'ประเภท'];
      sh2.getRange(1, 1, 1, h2.length).setValues([h2])
        .setFontWeight('bold').setFontSize(12)
        .setBackground('#1e3a8a').setFontColor('#ffffff')
        .setHorizontalAlignment('center');

      var thaiM = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
      function fmtD2(ds) {
        if (!ds) return '-';
        try { var d = ds instanceof Date ? ds : new Date(ds);
          if (isNaN(d.getTime())) return String(ds);
          return d.getDate() + ' ' + thaiM[d.getMonth()] + ' ' + (d.getFullYear() + 543); } catch(e) { return '-'; }
      }
      var d2Rows = reg.map(function(r, i) {
        return [i + 1, r.fullName || '-', r.position || '-', r.department || '-',
                r.topic || '-', fmtD2(r.date), fmtD2(r.endDate), r.location || '-',
                r.hours || 0, r.category || '-'];
      });
      if (d2Rows.length > 0) {
        sh2.getRange(2, 1, d2Rows.length, h2.length).setValues(d2Rows).setFontSize(11);
      }
      sh2.autoResizeColumns(1, h2.length);
    }

    // timestamp
    var lastRow = sh.getLastRow() + 2;
    sh.getRange(lastRow, 1).setValue('พิมพ์เมื่อ: ' +
      Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') + ' น.')
      .setFontStyle('italic').setFontSize(11).setFontColor('#64748b');

    // ย้ายไปยัง ATTACHMENTS folder ถ้ากำหนดไว้
    try {
      var file = DriveApp.getFileById(ss.getId());
      var folder = DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID);
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    } catch(fe) {
      Logger.log('⚠️ ย้ายไฟล์ไป folder ไม่ได้ (ไม่ใช่ error หลัก): ' + fe.toString());
    }

    var url = ss.getUrl();
    Logger.log('✅ exportDashboardToWord: ' + url);
    return { success: true, url: url, id: ss.getId() };

  } catch (e) {
    Logger.log('❌ exportDashboardToWord: ' + e);
    return { success: false, message: e.toString() };
  }
}
/**
 * รับ password จาก client → คืน session token ถ้าถูก
 * Token = HMAC-like string ที่มีอายุ 8 ชั่วโมง
 * Client เก็บ token ใน sessionStorage แล้วส่งมาทุกครั้งที่เรียก Admin function
 */
function adminLogin(password) {
  if (!password || password !== getAdminPassword()) {
    return { success: false, message: 'รหัสผ่านไม่ถูกต้อง' };
  }
  // สร้าง token = hash ของ (password + วันชั่วโมงปัจจุบัน)
  // GAS ไม่มี crypto โดยตรง ใช้ Utilities.computeDigest แทน
  var now      = new Date();
  var hourSlot = now.getFullYear() + '-' + now.getMonth() + '-' + now.getDate() + '-' + now.getHours();
  var raw      = getAdminPassword() + '|' + hourSlot;
  var bytes    = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
  var token    = bytes.map(function(b) {
    return ('0' + (b & 0xff).toString(16)).slice(-2);
  }).join('').substring(0, 32);

  return { success: true, token: token };
}

/**
 * ตรวจสอบ token ที่ client ส่งมา
 * คืน true/false — ใช้ก่อนทุก Admin-only function
 */
function verifyAdminToken(token) {
  if (!token) return false;
  // คำนวณ token ของ slot ปัจจุบัน
  var now      = new Date();
  var hourSlot = now.getFullYear() + '-' + now.getMonth() + '-' + now.getDate() + '-' + now.getHours();
  var raw      = getAdminPassword() + '|' + hourSlot;
  var bytes    = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
  var expected = bytes.map(function(b) {
    return ('0' + (b & 0xff).toString(16)).slice(-2);
  }).join('').substring(0, 32);

  if (token === expected) return true;

  // ยอมรับ token ของ slot ก่อนหน้า 1 ชั่วโมง (กันกรณีเปลี่ยนชั่วโมงพอดี)
  var prev     = new Date(now.getTime() - 3600000);
  var prevSlot = prev.getFullYear() + '-' + prev.getMonth() + '-' + prev.getDate() + '-' + prev.getHours();
  var rawPrev  = getAdminPassword() + '|' + prevSlot;
  var bytesPrev = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, rawPrev);
  var expectedPrev = bytesPrev.map(function(b) {
    return ('0' + (b & 0xff).toString(16)).slice(-2);
  }).join('').substring(0, 32);

  return token === expectedPrev;
}

function verifyAdminTokenPrevHour(token) {
  // alias — SHA-256 verifyAdminToken already checks prev hour slot internally
  return verifyAdminToken(token);
}

function getLogoImage() {
  var fileId = '1qIREbhnqt9n4xWbdqe5V8lnypb4DKtVS';
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  var base64 = Utilities.base64Encode(blob.getBytes());
  var mimeType = blob.getContentType();
  return 'data:' + mimeType + ';base64,' + base64;
}

