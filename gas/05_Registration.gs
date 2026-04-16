// ============================================
// === PART 1: ระบบลงทะเบียน ===
// ============================================

function saveRegistration(data, files) {
  try {
    if (!data || !data.date || !data.fullName || !data.position || !data.missionGroup ||
        !data.department || !data.category || !data.topic || !data.hours ||
        !data.location || !data.time) {
      return { success: false, message: 'ข้อมูลไม่ครบถ้วน' };
    }

    var spreadsheet = getOrCreateSheet();
    var sheet = spreadsheet.getSheetByName('Registrations');
    var hours = parseFloat(data.hours);

    if (isNaN(hours) || hours <= 0) {
      return { success: false, message: 'จำนวนชั่วโมงไม่ถูกต้อง (ต้องมากกว่า 0)' };
    }

    var dateParts = data.date.split('-');
    var year  = parseInt(dateParts[0], 10);
    var month = parseInt(dateParts[1], 10) - 1;
    var day   = parseInt(dateParts[2], 10);
    var utcDate = new Date(Date.UTC(year, month, day));

    var endDateParts = data.endDate ? data.endDate.split('-') : null;
    var utcEndDate   = null;
    if (endDateParts && endDateParts.length === 3) {
      utcEndDate = new Date(Date.UTC(
        parseInt(endDateParts[0], 10),
        parseInt(endDateParts[1], 10) - 1,
        parseInt(endDateParts[2], 10)
      ));
    }

    if (utcEndDate && utcEndDate < utcDate) {
      return { success: false, message: 'วันที่สิ้นสุดต้องมากกว่าหรือเท่ากับวันที่เริ่มต้น' };
    }

    var lineUserId = String(data.lineUserId || '').trim();

    // อัพโหลดไฟล์แนบก่อน เพื่อได้ URL สำหรับบันทึกลง Sheet พร้อมกัน
    var attachResults = [];
    var fileNames = '';
    var fileUrls  = '';
    if (files && files.length > 0) {
      try {
        attachResults = saveFilesToDrive(files);
        var okFiles2 = attachResults.filter(function(r) { return r && r.url; });
        fileNames = attachResults.map(function(r) { return r.name; }).join(', ');
        fileUrls  = okFiles2.map(function(r) { return r.url; }).join('\n');
        Logger.log('✅ อัพโหลดไฟล์แนบ register: ' + okFiles2.length + '/' + attachResults.length + ' ไฟล์');
      } catch(uploadErr) {
        Logger.log('⚠️ อัพโหลดไฟล์แนบล้มเหลว: ' + uploadErr.toString());
      }
    }

    // บันทึกลง Registrations (17 คอลัมน์ รวม LINE User ID + ไฟล์แนบ)
    var newRow = [
      new Date(),
      utcDate,
      utcEndDate,
      String(data.time).trim(),
      String(data.fullName).trim(),
      String(data.position).trim(),
      String(data.missionGroup).trim(),
      String(data.department).trim(),
      String(data.category).trim(),
      String(data.topic).trim(),
      hours,
      String(data.summary    || '').trim(),
      String(data.suggestion || '').trim(),
      String(data.location   || '').trim(),
      lineUserId,
      fileNames,
      fileUrls
    ];

    sheet.appendRow(newRow);

    // ถ้ามี URL ไฟล์แนบ → ใส่เป็น HYPERLINK ในคอลัมน์ URL ไฟล์แนบ
    if (fileUrls) {
      var lastRow = sheet.getLastRow();
      var hdr2 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var urlColIdx = hdr2.indexOf('URL ไฟล์แนบ');
      if (urlColIdx !== -1) {
        var urlList = fileUrls.split('\n').filter(function(u) { return u.trim(); });
        if (urlList.length === 1) {
          var fname = fileNames.split(',')[0].trim();
          sheet.getRange(lastRow, urlColIdx + 1).setFormula(
            '=HYPERLINK("' + urlList[0] + '","' + fname + '")'
          );
        } else {
          // หลายไฟล์ → ใส่ plain text URL ทั้งหมด (Sheet ไม่รองรับ multi-hyperlink ใน cell เดียว)
          sheet.getRange(lastRow, urlColIdx + 1).setValue(fileUrls);
        }
      }
    }

    SpreadsheetApp.flush();

    // ถ้าผู้ลงทะเบียนกรอก LINE User ID → upsert ลง Sheet ข้อมูลพนักงาน ด้วย
    if (lineUserId) {
      saveEmployeeLineData({
        fullName   : String(data.fullName).trim(),
        lineId     : lineUserId,
        status     : 'เปิดใช้งาน',
        position   : String(data.position    || '').trim(),
        mission    : String(data.missionGroup || '').trim(),
        department : String(data.department  || '').trim(),
        _internal  : true   // ← bypass admin token check
      }, null);
    }

    var wordResult = createSingleRegistrationWord(data);

    var okFiles = attachResults.filter(function(r) { return r && r.url; });
    return {
      success    : true,
      message    : 'บันทึกข้อมูลเรียบร้อยแล้ว',
      wordDoc    : wordResult.success ? { url: wordResult.url, id: wordResult.id } : null,
      attachFiles: okFiles,
      fileNames  : fileNames,
      fileUrls   : fileUrls
    };
  } catch (error) {
    Logger.log('Error in saveRegistration: ' + error.toString());
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function createSingleRegistrationWord(data) {
  try {
    var fileName = 'สรุปรายงานผล_' + data.fullName + '_' + formatDateForFileName(data.date);

    // สร้าง Spreadsheet แทน DocumentApp (ไม่ต้องการ scope documents)
    var ss   = SpreadsheetApp.create(fileName);
    var sh   = ss.getActiveSheet();
    sh.setName('บันทึกข้อความ');

    // ── หัว ──
    sh.getRange('A1').setValue('บันทึกข้อความ');
    sh.getRange('A1').setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange('A1:B1').merge();

    sh.getRange('A2').setValue('ส่วนราชการ  กลุ่มงาน' + data.department + ' โรงพยาบาลสันทราย');
    sh.getRange('A2:B2').merge().setFontSize(13);

    sh.getRange('A3').setValue('ที่ ชม................./                                           วันที่ ' + formatDateThaiLong(data.date));
    sh.getRange('A3:B3').merge().setFontSize(13);

    sh.getRange('A4').setValue('เรื่อง  สรุปรายงานผลการเข้าร่วมประชุม/ อบรม/ สัมมนา');
    sh.getRange('A4:B4').merge().setFontSize(13);

    sh.getRange('A5').setValue('เรียน  ผู้อำนวยการโรงพยาบาลสันทราย (ผ่านหัวหน้ากลุ่มงาน/รองผู้อำนวยการกลุ่มภารกิจ)');
    sh.getRange('A5:B5').merge().setFontSize(13);

    var content = '      ข้าพเจ้า ' + data.fullName + ' ตำแหน่ง' + data.position +
                  ' ได้เข้าร่วม' + data.category.toLowerCase() + ' ' + data.topic +
                  ' ระหว่างวันที่ ' + formatThaiDate(data.date) + ' ' +
                  (data.endDate ? 'ถึง ' + formatThaiDate(data.endDate) : '') +
                  ' เวลา ' + data.time + ' น. ณ ' + data.location +
                  ' เสร็จเรียบร้อยแล้วนั้น ในการนี้จึงขอสรุปผลการเข้าร่วมประชุม/อบรม/สัมมนา ดังนี้';
    sh.getRange('A6').setValue(content);
    sh.getRange('A6:B6').merge().setFontSize(13).setWrap(true);

    sh.getRange('A7').setValue('1. สรุปสาระสำคัญ').setFontWeight('bold').setFontSize(13);
    sh.getRange('A7:B7').merge();
    sh.getRange('A8').setValue('      ' + (data.summary || '(ไม่มีข้อมูลสรุป)'));
    sh.getRange('A8:B8').merge().setFontSize(13).setWrap(true);

    sh.getRange('A9').setValue('2. ข้อคิดเห็น/ ข้อเสนอแนะ/ ข้อที่จะนำมาปฏิบัติ').setFontWeight('bold').setFontSize(13);
    sh.getRange('A9:B9').merge();
    sh.getRange('A10').setValue('      ' + (data.suggestion || '(ไม่มีข้อมูลเพิ่มเติม)'));
    sh.getRange('A10:B10').merge().setFontSize(13).setWrap(true);

    sh.getRange('A11').setValue('3. งบประมาณที่ใช้ในการประชุม/ อบรม/ สัมมนา/ ศึกษาดูงานครั้งนี้').setFontWeight('bold').setFontSize(13);
    sh.getRange('A11:B11').merge();

    var budgetLines = [
      ['(  ) ไม่ใช้งบประมาณ'],
      ['(  ) เงินบำรุงโรงพยาบาลสันทราย'],
      ['      · ค่าลงทะเบียน                                   บาท'],
      ['      · ค่าเบี้ยเลี้ยง                                         บาท'],
      ['      · ค่าที่พัก                                                 บาท'],
      ['      · ค่าพาหนะเดินทาง                            บาท'],
      ['      · อื่น ๆ                                                      บาท'],
      ['      · รวมทั้งสิ้น                                              บาท'],
      ['(  ) งบประมาณแหล่งอื่นๆ จาก...................................................................จำนวน          บาท'],
      ['(  ) จากผู้จัด'],
      ['(  ) ไม่ประสงค์จะขอเบิกเงิน']
    ];
    var budgetStart = 12;
    budgetLines.forEach(function(line, i) {
      var r = sh.getRange(budgetStart + i, 1);
      r.setValue(line[0]).setFontSize(13);
      sh.getRange(budgetStart + i, 1, 1, 2).merge();
    });

    var sigStart = budgetStart + budgetLines.length + 2;
    sh.getRange(sigStart, 1).setValue('(ลงชื่อ)....................................................').setFontSize(13);
    sh.getRange(sigStart, 1, 1, 2).merge();
    sh.getRange(sigStart + 1, 1).setValue('     (' + data.fullName + ')').setFontSize(13);
    sh.getRange(sigStart + 1, 1, 1, 2).merge();
    sh.getRange(sigStart + 2, 1).setValue('ผู้เข้าร่วมอบรม').setFontSize(13);
    sh.getRange(sigStart + 2, 1, 1, 2).merge();

    sh.getRange(sigStart + 4, 1).setValue('(ลงชื่อ)................................................…').setFontSize(13);
    sh.getRange(sigStart + 4, 1, 1, 2).merge();
    sh.getRange(sigStart + 5, 1).setValue('     (..........................................)').setFontSize(13);
    sh.getRange(sigStart + 5, 1, 1, 2).merge();
    sh.getRange(sigStart + 6, 1).setValue('หัวหน้ากลุ่มงาน').setFontSize(13);
    sh.getRange(sigStart + 6, 1, 1, 2).merge();

    sh.getRange(sigStart + 8, 1).setValue('(ลงชื่อ).....................………....................').setFontSize(13);
    sh.getRange(sigStart + 8, 1, 1, 2).merge();
    sh.getRange(sigStart + 9, 1).setValue('     (...................................)').setFontSize(13);
    sh.getRange(sigStart + 9, 1, 1, 2).merge();
    sh.getRange(sigStart + 10, 1).setValue('รองผู้อำนวยการ..................').setFontSize(13);
    sh.getRange(sigStart + 10, 1, 1, 2).merge();

    sh.getRange(sigStart + 12, 1).setValue('(ลงชื่อ)...................................................').setFontSize(13);
    sh.getRange(sigStart + 12, 1, 1, 2).merge();
    sh.getRange(sigStart + 13, 1).setValue('     (นายวรวุฒิ   โฆวัชรกุล)').setFontSize(13);
    sh.getRange(sigStart + 13, 1, 1, 2).merge();
    sh.getRange(sigStart + 14, 1).setValue('ผู้อำนวยการโรงพยาบาลสันทราย').setFontSize(13);
    sh.getRange(sigStart + 14, 1, 1, 2).merge();

    // column width
    sh.setColumnWidth(1, 600);
    sh.setColumnWidth(2, 200);

    // row heights for wrapped text
    sh.setRowHeight(6, 80);
    sh.setRowHeight(8, 60);
    sh.setRowHeight(10, 60);

    // ย้ายไปยัง ATTACHMENTS folder ถ้ากำหนดไว้
    try {
      var file = DriveApp.getFileById(ss.getId());
      var folder = DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID);
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    } catch(fe) {
      Logger.log('⚠️ ย้ายไฟล์ไป folder ไม่ได้ (ไม่ใช่ error หลัก): ' + fe.toString());
    }

    return {
      success: true,
      url: ss.getUrl(),
      id: ss.getId(),
      message: 'สร้างเอกสารสำเร็จ'
    };

  } catch (error) {
    Logger.log('Error in createSingleRegistrationWord: ' + error.toString());
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function addWatermark(body) {
  // ฟังก์ชันนี้เคยใช้กับ DocumentApp แต่ปัจจุบันเปลี่ยนไปใช้ SpreadsheetApp แล้ว
  // คงไว้เพื่อ backward compatibility ไม่ทำงานและไม่ throw error
  try {
    Logger.log('ℹ️ addWatermark: ข้ามเนื่องจากใช้ SpreadsheetApp แล้ว');
  } catch (error) {
    Logger.log('Error in addWatermark: ' + error.toString());
  }
}

// ============================================
// === Date Formatting Functions ===
// ============================================

function formatDateForFileName(dateStr) {
  try {
    var date = new Date(dateStr);
    var day = String(date.getDate()).padStart(2, '0');
    var month = String(date.getMonth() + 1).padStart(2, '0');
    var year = date.getFullYear();
    return day + month + year;
  } catch (error) {
    Logger.log('Error in formatDateForFileName: ' + error.toString());
    return '';
  }
}

function formatThaiDate(dateStr) {
  if (!dateStr) return '';
  try {
    var date = new Date(dateStr);
    var day = date.getDate();
    var month = date.getMonth();
    var year = date.getFullYear() + 543;
    var thaiMonths = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
                      'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
    return day + ' ' + thaiMonths[month] + ' ' + year;
  } catch (error) {
    Logger.log('Error in formatThaiDate: ' + error.toString());
    return '';
  }
}

function formatDateThaiLong(dateStr) {
  if (!dateStr) return '';
  try {
    var date = new Date(dateStr);
    var day = date.getDate();
    var month = date.getMonth();
    var year = date.getFullYear() + 543;
    var thaiMonths = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
                      'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
    return day + ' ' + thaiMonths[month] + ' ' + year;
  } catch (error) {
    Logger.log('Error in formatDateThaiLong: ' + error.toString());
    return '';
  }
}

function formatDateThaiShort(date) {
  try {
    var day = String(date.getDate()).padStart(2, '0');
    var month = String(date.getMonth() + 1).padStart(2, '0');
    var year = date.getFullYear() + 543;
    return day + '/' + month + '/' + year;
  } catch (error) {
    Logger.log('Error in formatDateThaiShort: ' + error.toString());
    return '';
  }
}


