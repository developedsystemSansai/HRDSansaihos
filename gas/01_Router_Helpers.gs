// ============================================
// === HTML Rendering — แยก URL ผู้ใช้ / Admin ===
// ============================================
//
//  URL ปกติ   : ?  (ไม่มี parameter)      → โหลด index.html  (4 แท็บ ไม่มีพนักงาน)
//  URL Admin  : ?mode=admin               → โหลด index.html  พร้อมส่ง isAdmin=true
//                                           แล้วหน้าจะแสดงปุ่มล็อกอินแท็บพนักงาน
//
//  ข้อดี: ใช้ไฟล์ HTML เดียว ไม่ต้อง deploy หลายครั้ง
//  ข้อดี: Admin Password ตรวจสอบฝั่ง Server (GAS) ไม่ใช่แค่ JS

function doGet(e) {
  var action = e.parameter.action;
  var callback = e.parameter.callback;

  var output;

  if (action === 'dashboard') {
    output = getDashboardData();
  } else if (action === 'detail') {
    output = getUserDetail(e.parameter.name);
  }

  if (action) {
    var json = JSON.stringify(output || {});
    if (callback) {
      return ContentService
        .createTextOutput(callback + "(" + json + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  var template = HtmlService.createTemplateFromFile('index');
  template.isAdminMode = (e.parameter.mode === 'admin');

  return template.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
// ============================================
// === Helper Function สำหรับหาคอลัมน์ (รองรับชื่อไทย/อังกฤษ/ดึงตาม Index สำรอง) ===
// ============================================
function getColIdx(hdr, enName, thName, fallbackIdx) {
  var idx = hdr.indexOf(enName);
  if (idx === -1 && thName) idx = hdr.indexOf(thName);
  return idx !== -1 ? idx : fallbackIdx;
}
// ============================================
// === 1. getAllUniqueStaffNames() ===
// ดึงรายชื่อพนักงานไม่ซ้ำจาก Sheet "อบรม ปีงบประมาณ 2569"
// ============================================
