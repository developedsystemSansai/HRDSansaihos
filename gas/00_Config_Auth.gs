/**
 * ⚠️ รันฟังก์ชันนี้ 1 ครั้งจาก Apps Script Editor (▶ Run) ก่อน/หลัง Deploy ใหม่
 * เพื่อ grant permission DriveApp ให้กับ Web App
 * ขั้นตอน: Script Editor → เลือก initDriveAccess → กด Run → Allow permissions
 */
function initDriveAccess() {
  try {
    var folder = DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID);
    Logger.log('✅ DriveApp ทำงานได้ปกติ');
    Logger.log('✅ โฟลเดอร์: ' + folder.getName() + ' (id=' + ATTACHMENTS_FOLDER_ID + ')');
    // สร้างไฟล์ทดสอบแล้วลบทันที
    var testBlob = Utilities.newBlob('HRD Drive Access Test ' + new Date().toISOString(), 'text/plain', '_hrd_test_.txt');
    var testFile = folder.createFile(testBlob);
    testFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var testUrl = 'https://drive.google.com/file/d/' + testFile.getId() + '/view?usp=sharing';
    Logger.log('✅ สร้างไฟล์ทดสอบสำเร็จ: ' + testUrl);
    testFile.setTrashed(true);
    Logger.log('✅ ลบไฟล์ทดสอบแล้ว');
    Logger.log('✅ DriveApp พร้อมใช้งาน! Deploy Web App ได้เลย');
  } catch (e) {
    Logger.log('❌ DriveApp Error: ' + e.toString());
    Logger.log('  → ตรวจสอบ ATTACHMENTS_FOLDER_ID: ' + ATTACHMENTS_FOLDER_ID);
    Logger.log('  → ถ้า "not authorized" กด Run อีกครั้งแล้วกด Allow');
  }
}

/**
 * ทดสอบอัพโหลดไฟล์ PDF จำลองไปยัง Drive
 * รันจาก Script Editor เพื่อยืนยันก่อน deploy
 */
function testDriveUpload() {
  try {
    var folder = DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID);
    Logger.log('\u2705 โฟลเดอร์: ' + folder.getName());
    var fakePdfB64 = 'JVBERi0xLjQKMSAwIG9iago8PAovVHlwZSAvQ2F0YWxvZwovUGFnZXMgMiAwIFIKPj4KZW5kb2JqCjIgMCBvYmoKPDwKL1R5cGUgL1BhZ2VzCi9LaWRzIFszIDAgUl0KL0NvdW50IDEKPJ4KZW5kb2JqCjMgMCBvYmoKPDwKL1R5cGUgL1BhZ2UKL1BhcmVudCAyIDAgUgovTWVkaWFCb3ggWzAgMCA2MTIgNzkyXQo+PgplbmRvYmoKeHJlZgowIDQKMDAwMDAwMDAwMCA2NTUzNSBmIAowMDAwMDAwMDA5IDAwMDAwIG4gCjAwMDAwMDAwNTggMDAwMDAgbiAKMDAwMDAwMDExNSAwMDAwMCBuIAp0cmFpbGVyCjw8Ci9TaXplIDQKL1Jvb3QgMSAwIFIKPj4Kc3RhcnR4cmVmCjE5MAolJUVPRgo=';
    var bytes = Utilities.base64Decode(fakePdfB64);
    var blob  = Utilities.newBlob(bytes, 'application/pdf', 'test_hrd_upload.pdf');
    var file  = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=sharing';
    Logger.log('\u2705 อัพโหลดสำเร็จ! URL: ' + url);
    file.setTrashed(true);
    Logger.log('\u2705 ระบบอัพโหลดพร้อมใช้งาน');
  } catch (e) {
    Logger.log('\u274c testDriveUpload Error: ' + e.toString());
  }
}

/**
 * ตรวจสอบสถานะ Drive + Spreadsheet ทั้งหมด (รันเพื่อ diagnose)
 */
function checkSystemStatus() {
  Logger.log('=== HRD System Status Check ===');
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log('\u2705 SpreadsheetApp: ' + ss.getName());
    var ts = ss.getSheetByName('\u0e2d\u0e1a\u0e23\u0e21 \u0e1b\u0e35\u0e07\u0e1a\u0e1b\u0e23\u0e30\u0e21\u0e32\u0e13 2569');
    Logger.log(ts ? '\u2705 Sheet อบรมฯ: ' + ts.getLastRow() + ' แถว' : '\u26a0\ufe0f ยังไม่มี Sheet อบรมฯ');
  } catch (e) { Logger.log('\u274c SpreadsheetApp: ' + e.message); }
  try {
    var folder = DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID);
    Logger.log('\u2705 DriveApp โฟลเดอร์: ' + folder.getName());
  } catch (e) {
    Logger.log('\u274c DriveApp: ' + e.message);
    Logger.log('   \u2192 รัน initDriveAccess() แล้ว grant permission ก่อน');
  }
  try { Logger.log('\u2705 Script user: ' + Session.getActiveUser().getEmail()); } catch(e) {}
  Logger.log('=== เสร็จสิ้น ===');
}

/**
 * ▶ รัน quickDriveTest() เพื่อทดสอบสิทธิ์อัพโหลด Drive ทีละขั้นตอน
 * ดู Log เพื่อระบุว่าปัญหาอยู่ที่ไหน
 */
function quickDriveTest() {
  Logger.log('--- Quick Drive Test ---');
  Logger.log('Account: ' + Session.getActiveUser().getEmail());

  // Test 1: createFile ใน root Drive
  try {
    var f1 = DriveApp.createFile(
      Utilities.newBlob('root test', 'text/plain', '_root_test.txt')
    );
    Logger.log('\u2705 createFile in root: OK — id=' + f1.getId());
    f1.setTrashed(true);
  } catch (e) {
    Logger.log('\u274c createFile in root: FAIL — ' + e.message);
    Logger.log('  \u2192 DriveApp scope ยังไม่ได้รับ authorization');
    Logger.log('  \u2192 ตรวจสอบ appsscript.json และกด Run อีกครั้ง → Allow');
    return;
  }

  // Test 2: createFile ใน folder เป้าหมาย
  try {
    var folder = DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID);
    Logger.log('\u2705 folder: ' + folder.getName());
    var f2 = folder.createFile(
      Utilities.newBlob('folder test', 'text/plain', '_folder_test.txt')
    );
    Logger.log('\u2705 createFile in folder: OK — id=' + f2.getId());
    f2.setTrashed(true);
    Logger.log('');
    Logger.log('\ud83c\udf89 ทุกอย่างพร้อม! Deploy \u2192 New version ได้เลย');
  } catch (e) {
    Logger.log('\u274c createFile in folder: FAIL — ' + e.message);
    Logger.log('');
    Logger.log('\u26a0 root Drive สร้างได้ แต่ folder นี้ไม่ได้');
    Logger.log('  \u2192 Folder สร้างโดย account อื่น หรือสิทธิ์ไม่ถึง Editor');
    Logger.log('  \u2192 รัน createNewAttachmentsFolder() แล้ว copy ID ใหม่มาใส่ ATTACHMENTS_FOLDER_ID');
  }
}

// ============================================
// === ตั้งค่าหลัก (แก้ไขค่าเหล่านี้ก่อนใช้งาน) ===
// ============================================

var SPREADSHEET_ID  = '1g5dxhE-9tT0vqA3chSZ6hLsjeHkafAqgpO-bdbUtwFM';

function getAdminPassword() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
}

function getStaffPassword() {
  return PropertiesService.getScriptProperties().getProperty('STAFF_PASSWORD');
}

function loginAdmin(password) {
  if (password !== getAdminPassword()) {
    return { success: false };
  }

  var token = Utilities.getUuid();

  CacheService.getScriptCache().put(token, "admin", 3600); // 1 ชม.

  return { success: true, token: token };
}
function loginStaff(password) {
  if (password !== getStaffPassword()) {
    return { success: false };
  }

  var token = Utilities.getUuid();

  CacheService.getScriptCache().put(token, "staff", 3600); // 1 ชม.

  return { success: true, token: token };
}
function verifyStaffToken(token) {
  return CacheService.getScriptCache().get(token) === "staff";
}
function login(role, password) {
  if (role === 'admin' && password === getAdminPassword()) {
    var token = Utilities.getUuid();
    CacheService.getScriptCache().put(token, "admin", 3600);
    return { success: true, token: token, role: 'admin' };
  }

  if (role === 'staff' && password === getStaffPassword()) {
    var token = Utilities.getUuid();
    CacheService.getScriptCache().put(token, "staff", 3600);
    return { success: true, token: token, role: 'staff' };
  }

  return { success: false };
}
// ── โฟลเดอร์ Google Drive สำหรับเก็บเอกสารแนบ ──
// ⚠️  ถ้า error "Access denied" → รัน createNewAttachmentsFolder() แล้วเปลี่ยน ID ที่นี่
var ATTACHMENTS_FOLDER_ID = '1Lr86HW_cF8pocuxi-E0cDTMNI_sijqPQ';

// ── หมวดหมู่ที่ไม่นับหน่วยกิตสะสม (บันทึกชั่วโมงไว้แต่ไม่รวมใน totalHours) ──
var NO_CREDIT_CAT = 'ประชุมคณะกรรมการ/คณะทำงาน/ประชุมชี้แจง/ประชุมผู้รับผิดชอบงาน ฯ';

// ============================================
// === 🔧 Drive Folder Setup Functions ===
// ============================================

/**
 * ▶ รัน STEP 1 ก่อน: diagnoseFolderPermission()
 *   ดูว่าปัญหาอยู่ที่ไหน
 *
 * ▶ รัน STEP 2 ถ้าจำเป็น: createNewAttachmentsFolder()
 *   สร้างโฟลเดอร์ใหม่ใน My Drive แล้ว copy ID มาใส่ ATTACHMENTS_FOLDER_ID
 *
 * ▶ รัน STEP 3: initDriveAccess()
 *   ยืนยันว่าทุกอย่างพร้อม แล้ว Deploy → New version
 */

/**
 * STEP 1 — ตรวจสอบสิทธิ์โฟลเดอร์ปัจจุบัน
 */
function diagnoseFolderPermission() {
  Logger.log('=== Folder Permission Diagnosis ===');
  Logger.log('FOLDER_ID: ' + ATTACHMENTS_FOLDER_ID);
  var folder;
  try {
    folder = DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID);
    Logger.log('\u2705 Step 1 getFolderById: OK — ' + folder.getName());
  } catch (e) {
    Logger.log('\u274c Step 1 getFolderById: FAIL — ' + e.message);
    Logger.log('   \u2192 โฟลเดอร์ไม่มีอยู่ หรือไม่ได้แชร์ให้ account นี้');
    return;
  }
  try {
    var owner = folder.getOwner();
    Logger.log('\ud83d\udc64 Owner: ' + (owner ? owner.getEmail() : '?'));
    Logger.log('\ud83d\udc64 Me:    ' + Session.getActiveUser().getEmail());
  } catch(e) {}
  try {
    var t = folder.createFile(Utilities.newBlob('test','text/plain','_perm_test.txt'));
    t.setTrashed(true);
    Logger.log('\u2705 Step 2 createFile: OK — มีสิทธิ์เขียน ใช้งานได้เลย!');
    Logger.log('   \u2192 ปัญหาน่าจะอยู่ที่ appsscript.json หรือ Web App scope');
    Logger.log('   \u2192 ตรวจสอบ appsscript.json และ Redeploy → New version');
  } catch (e) {
    Logger.log('\u274c Step 2 createFile: FAIL — ' + e.message);
    Logger.log('');
    Logger.log('\ud83d\udd27 สาเหตุ: สิทธิ์เป็นแค่ "Viewer" ไม่ใช่ "Editor"');
    Logger.log('   \u2192 รัน createNewAttachmentsFolder() เพื่อสร้างโฟลเดอร์ใหม่');
  }
}

/**
 * STEP 2 — สร้างโฟลเดอร์ใหม่ใน My Drive (ถ้าโฟลเดอร์เดิมไม่มีสิทธิ์เขียน)
 * หลังรันแล้ว: copy ID จาก Log มาแทนใน ATTACHMENTS_FOLDER_ID บรรทัด 93
 */
function createNewAttachmentsFolder() {
  try {
    var newFolder = DriveApp.createFolder('HRD_Attachments_2569');
    var newId = newFolder.getId();
    // ทดสอบสร้างไฟล์ทันที
    var t = newFolder.createFile(Utilities.newBlob('test','text/plain','_test.txt'));
    t.setTrashed(true);
    Logger.log('\u2705 สร้างโฟลเดอร์ใหม่สำเร็จ!');
    Logger.log('');
    Logger.log('\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550');
    Logger.log('\u2605 FOLDER ID ใหม่ (copy นี้): ' + newId);
    Logger.log('\u2605 URL: https://drive.google.com/drive/folders/' + newId);
    Logger.log('');
    Logger.log('ขั้นตอนต่อไป:');
    Logger.log('1. Copy ID: ' + newId);
    Logger.log('2. แก้บรรทัด ATTACHMENTS_FOLDER_ID ใน รหัส.gs ให้เป็น ID ใหม่');
    Logger.log('3. กด Save');
    Logger.log('4. รัน initDriveAccess() อีกครั้ง');
    Logger.log('5. Deploy \u2192 New version');
    Logger.log('\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550');
  } catch (e) {
    Logger.log('\u274c Error: ' + e.toString());
  }
}

