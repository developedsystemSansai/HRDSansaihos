// ============================================
// === Sheet Management ===
// ============================================

function getOrCreateSheet() {
  try {
    var spreadsheet;
    if (SPREADSHEET_ID) {
      spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    } else {
      spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    }
    if (!spreadsheet) {
      throw new Error('ไม่พบ Spreadsheet — ตรวจสอบ SPREADSHEET_ID ในโค้ด');
    }
    setupSheets(spreadsheet);
    return spreadsheet;
  } catch (error) {
    console.error('Error in getOrCreateSheet:', error.toString());
    throw new Error('ไม่สามารถเข้าถึง Spreadsheet ได้: ' + error.toString());
  }
}

/**
 * setupSheets — สร้าง/ตรวจสอบ Sheet ที่จำเป็นทั้งหมด
 * รับ spreadsheet object โดยตรง หรือเรียกโดยไม่ส่ง parameter
 * (ในกรณีรัน manual จาก Apps Script editor)
 */
function setupSheets(spreadsheet) {
  try {
    // ถ้าเรียกโดยไม่ส่ง parameter (เช่น รันเอง) ให้ดึง spreadsheet เอง
    if (!spreadsheet || typeof spreadsheet.getSheetByName !== 'function') {
      if (SPREADSHEET_ID) {
        spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      } else {
        spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      }
    }
    if (!spreadsheet) {
      throw new Error('ไม่พบ Spreadsheet — ตรวจสอบ SPREADSHEET_ID');
    }
    var registrationsSheet = spreadsheet.getSheetByName('Registrations');
    if (!registrationsSheet) {
      registrationsSheet = spreadsheet.insertSheet('Registrations');
      var headers = [
        'Timestamp', 'Date', 'End Date', 'Time', 'Full Name', 'Position',
        'Mission Group', 'Department', 'Category', 'Topic', 'Hours',
        'Summary', 'Suggestion', 'Location', 'LINE User ID',
        'ชื่อไฟล์แนบ', 'URL ไฟล์แนบ'
      ];
      registrationsSheet.getRange(1, 1, 1, 17).setValues([headers]).setFontWeight('bold');
    } else {
      // ถ้า Sheet มีอยู่แล้ว ตรวจว่ามีคอลัมน์ที่ต้องการหรือยัง
      var existingHeaders = registrationsSheet.getRange(1, 1, 1, registrationsSheet.getLastColumn()).getValues()[0];
      if (existingHeaders.indexOf('LINE User ID') === -1) {
        var nextCol = registrationsSheet.getLastColumn() + 1;
        registrationsSheet.getRange(1, nextCol).setValue('LINE User ID').setFontWeight('bold');
        Logger.log('✅ เพิ่มคอลัมน์ LINE User ID ใน Registrations');
        existingHeaders = registrationsSheet.getRange(1, 1, 1, registrationsSheet.getLastColumn()).getValues()[0];
      }
      if (existingHeaders.indexOf('ชื่อไฟล์แนบ') === -1) {
        var nextColA = registrationsSheet.getLastColumn() + 1;
        registrationsSheet.getRange(1, nextColA).setValue('ชื่อไฟล์แนบ').setFontWeight('bold');
        Logger.log('✅ เพิ่มคอลัมน์ ชื่อไฟล์แนบ ใน Registrations');
        existingHeaders = registrationsSheet.getRange(1, 1, 1, registrationsSheet.getLastColumn()).getValues()[0];
      }
      if (existingHeaders.indexOf('URL ไฟล์แนบ') === -1) {
        var nextColB = registrationsSheet.getLastColumn() + 1;
        registrationsSheet.getRange(1, nextColB).setValue('URL ไฟล์แนบ').setFontWeight('bold');
        Logger.log('✅ เพิ่มคอลัมน์ URL ไฟล์แนบ ใน Registrations');
      }
    }

    registrationsSheet.getRange('B2:B').setNumberFormat('dd/mm/yyyy');
    registrationsSheet.getRange('C2:C').setNumberFormat('dd/mm/yyyy');

    // Sheet: ข้อมูลพนักงาน (สำหรับ LINE Messaging API)
    var employeeSheet = spreadsheet.getSheetByName('ข้อมูลพนักงาน');
    if (!employeeSheet) {
      employeeSheet = spreadsheet.insertSheet('ข้อมูลพนักงาน');
      var empHeaders = [
        'รหัสพนักงาน', 'ชื่อ-สกุล', 'ตำแหน่ง', 'ระดับ', 'กลุ่มภารกิจ',
        'กลุ่มงาน', 'เบอร์โทร', 'อีเมล', 'LINE User ID',
        'สถานะแจ้งเตือน', 'วันที่อัปเดต', 'หมายเหตุ'
      ];
      employeeSheet.getRange(1, 1, 1, 12).setValues([empHeaders]).setFontWeight('bold');
    }

    // Sheet: ตั้งค่าระบบ
    var configSheet = spreadsheet.getSheetByName('ตั้งค่าระบบ');
    if (!configSheet) {
      configSheet = spreadsheet.insertSheet('ตั้งค่าระบบ');
      var configHeaders = ['หัวข้อการตั้งค่า', 'ค่าที่ตั้ง', 'คำอธิบาย'];
      configSheet.getRange(1, 1, 1, 3).setValues([configHeaders]).setFontWeight('bold');

      var configData = [
        ['แจ้งเตือนก่อนอบรม (วัน)', '7',  'แจ้งเตือนก่อนอบรม N วัน'],
        ['แจ้งเตือนหลังอบรม (วัน)', '7',  'เตือนส่งเอกสารสรุปผล หลังอบรมเสร็จ N วัน'],
        ['เวลาส่งแจ้งเตือน (ชั่วโมง)', '8', 'ชั่วโมงที่ส่งแจ้งเตือน (0-23) เช่น 8 = 08:00'],
        ['ชื่อโรงพยาบาล', 'โรงพยาบาลสันทราย', ''],
        ['ปีงบประมาณปัจจุบัน', '2568', ''],
        ['LINE Channel Access Token', '', 'วางค่า Channel Access Token จาก LINE Developers Console'],
        ['Google Drive Folder ID', '', 'ID ของโฟลเดอร์ที่ต้องการเก็บเอกสาร Word (ไม่บังคับ)']
      ];
      configSheet.getRange(2, 1, configData.length, 3).setValues(configData);
    }

    // Sheet: Log การแจ้งเตือน
    var logSheet = spreadsheet.getSheetByName('Log_การแจ้งเตือน');
    if (!logSheet) {
      logSheet = spreadsheet.insertSheet('Log_การแจ้งเตือน');
      var logHeaders = ['วันที่', 'เวลา', 'ประเภทแจ้งเตือน', 'ผู้รับ', 'หลักสูตร', 'สถานะ', 'Response Code', 'ข้อความ Error'];
      logSheet.getRange(1, 1, 1, 8).setValues([logHeaders]).setFontWeight('bold');
    }

    // ลบ Sheet1 ถ้ามี
    var defaultSheet = spreadsheet.getSheetByName('Sheet1');
    if (defaultSheet && spreadsheet.getSheets().length > 1) {
      spreadsheet.deleteSheet(defaultSheet);
    }
  } catch (error) {
    Logger.log('Error in setupSheets: ' + error.toString());
    throw error;
  }
}

