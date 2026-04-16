function getAllUniqueStaffNames() {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('อบรม ปีงบประมาณ 2569'); 
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, data: [] };
    }

    var data = sheet.getDataRange().getValues();
    var hdr  = data[0].map(function(h) { return String(h || '').trim(); });
    
    // ใช้ Header ภาษาไทยตามชีท อบรม ปีงบประมาณ 2569
    var cName = hdr.indexOf('ชื่อ-สกุล');
    var cPos  = hdr.indexOf('ตำแหน่ง');
    var cMis  = hdr.indexOf('กลุ่มภารกิจ');
    var cDept = hdr.indexOf('กลุ่มงาน');
    var cLine = hdr.indexOf('LINE ID');

    if (cName < 0) {
      return { success: false, message: 'ไม่พบคอลัมน์ "ชื่อ-สกุล" ใน Sheet อบรม ปีงบประมาณ 2569' };
    }

    var personMap = {};
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][cName] || '').trim();
      if (!name) continue;
      var lineId = cLine >= 0 ? String(data[i][cLine] || '').trim() : '';
      
      if (!personMap[name] || lineId) {
        personMap[name] = {
          fullName    : name,
          position    : cPos  >= 0 ? String(data[i][cPos]  || '').trim() : '',
          missionGroup: cMis  >= 0 ? String(data[i][cMis]  || '').trim() : '',
          department  : cDept >= 0 ? String(data[i][cDept] || '').trim() : '',
          lineId      : lineId
        };
      }
    }

    var results = Object.values(personMap);
    results.sort(function(a, b) { return a.fullName.localeCompare(b.fullName, 'th'); });
    Logger.log('✅ getAllUniqueStaffNames: ' + results.length + ' คน จาก อบรม ปีงบประมาณ 2569');
    return { success: true, data: results };
  } catch (e) {
    Logger.log('❌ getAllUniqueStaffNames: ' + e);
    return { success: false, message: e.toString() };
  }
}
// ============================================
// === 2. getEmployeeLineData(token) ===
// ดึง LINE ID จาก Sheet "อบรม ปีงบประมาณ 2569"
// ============================================
function getEmployeeLineData(token) {
  if (!verifyAdminToken(token)) {
    return { success: false, message: 'ไม่มีสิทธิ์ดูข้อมูลนี้ กรุณา Login ใหม่' };
  }
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    // ── อ่าน LINE ID จากชีท "อบรม ปีงบประมาณ 2569" ──
    var trainSheet = ss.getSheetByName('อบรม ปีงบประมาณ 2569');
    var lineFromTrain = {}; 
    if (trainSheet && trainSheet.getLastRow() > 1) {
      var tData = trainSheet.getDataRange().getValues();
      var tHdr  = tData[0].map(function(h) { return String(h || '').trim(); });
      var tName = tHdr.indexOf('ชื่อ-สกุล');
      var tLine = tHdr.indexOf('LINE ID');
      if (tName >= 0 && tLine >= 0) {
        for (var i = 1; i < tData.length; i++) {
          var n = String(tData[i][tName] || '').trim();
          var l = String(tData[i][tLine] || '').trim();
          if (n && l) lineFromTrain[n] = l;
        }
      }
    }

    // ── อ่านสถานะแจ้งเตือนจาก Sheet "ข้อมูลพนักงาน" ──
    var statusMap = {};
    var empSheet  = ss.getSheetByName('ข้อมูลพนักงาน');
    if (empSheet && empSheet.getLastRow() > 1) {
      var eData   = empSheet.getDataRange().getValues();
      var eHdr    = eData[0].map(function(h) { return String(h || '').trim(); });
      var eName   = eHdr.indexOf('ชื่อ-สกุล');
      var eStat   = eHdr.indexOf('สถานะแจ้งเตือน');
      var eLine   = eHdr.indexOf('LINE User ID');
      if (eName >= 0) {
        for (var j = 1; j < eData.length; j++) {
          var en = String(eData[j][eName] || '').trim();
          if (!en) continue;
          statusMap[en] = eStat >= 0 ? String(eData[j][eStat] || 'เปิดใช้งาน').trim() : 'เปิดใช้งาน';
          if (!lineFromTrain[en] && eLine >= 0) {
            var el = String(eData[j][eLine] || '').trim();
            if (el) lineFromTrain[en] = el;
          }
        }
      }
    }

    var results = Object.keys(lineFromTrain).map(function(name) {
      return {
        fullName: name,
        lineId  : lineFromTrain[name],
        status  : statusMap[name] || 'เปิดใช้งาน'
      };
    });
    results.sort(function(a, b) { return a.fullName.localeCompare(b.fullName, 'th'); });

    Logger.log('✅ getEmployeeLineData: ' + results.length + ' รายการ');
    return { success: true, data: results };
  } catch (e) {
    Logger.log('❌ getEmployeeLineData: ' + e);
    return { success: false, message: e.toString() };
  }
}

// ============================================
// === 3. saveEmployeeLineData(data, token) ===
// ============================================
function saveEmployeeLineData(data, token) {
  if (!data._internal && !verifyAdminToken(token)) {
    return { success: false, message: 'ไม่มีสิทธิ์แก้ไขข้อมูลนี้ กรุณา Login ใหม่' };
  }
  try {
    var ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
    var fullName = String(data.fullName || '').trim();
    var lineId   = String(data.lineId   || '').trim();
    var status   = String(data.status   || 'เปิดใช้งาน').trim();

    // ── 1. เขียน LINE ID กลับไปทุกแถวใน "อบรม ปีงบประมาณ 2569" ──
    var trainSheet = ss.getSheetByName('อบรม ปีงบประมาณ 2569');
    var updatedRows = 0;
    if (trainSheet && trainSheet.getLastRow() > 1) {
      var tData = trainSheet.getDataRange().getValues();
      var tHdr  = tData[0].map(function(h) { return String(h || '').trim(); });
      var tName = tHdr.indexOf('ชื่อ-สกุล');
      var tLine = tHdr.indexOf('LINE ID');
      if (tName >= 0 && tLine >= 0) {
        for (var i = 1; i < tData.length; i++) {
          var rowName = String(tData[i][tName] || '').trim();
          if (rowName === fullName) {
            trainSheet.getRange(i + 1, tLine + 1).setValue(lineId);
            updatedRows++;
          }
        }
        if (updatedRows > 0) SpreadsheetApp.flush();
        Logger.log('✅ อัปเดต LINE ID ใน อบรม ปีงบประมาณ 2569: ' + updatedRows + ' แถว');
      }
    }

    // ── 2. บันทึก/อัปเดตสถานะใน Sheet "ข้อมูลพนักงาน" ──
    var empSheet = ss.getSheetByName('ข้อมูลพนักงาน');
    if (!empSheet) {
      empSheet = ss.insertSheet('ข้อมูลพนักงาน');
      var headers = ['ชื่อ-สกุล', 'LINE User ID', 'สถานะแจ้งเตือน', 'ตำแหน่ง', 'กลุ่มภารกิจ', 'กลุ่มงาน', 'อัปเดตล่าสุด'];
      empSheet.getRange(1, 1, 1, headers.length)
              .setValues([headers]).setFontWeight('bold')
              .setBackground('#1e3a8a').setFontColor('#ffffff');
      empSheet.setFrozenRows(1);
    }

    var eData   = empSheet.getLastRow() > 1 ? empSheet.getRange(2, 1, empSheet.getLastRow() - 1, empSheet.getLastColumn()).getValues() : [];
    var eHdr    = empSheet.getRange(1, 1, 1, empSheet.getLastColumn()).getValues()[0].map(function(h) { return String(h || '').trim(); });
    var eName   = eHdr.indexOf('ชื่อ-สกุล');
    var eLine   = eHdr.indexOf('LINE User ID');
    var eStat   = eHdr.indexOf('สถานะแจ้งเตือน');
    var ePos    = eHdr.indexOf('ตำแหน่ง');
    var eMis    = eHdr.indexOf('กลุ่มภารกิจ');
    var eDept   = eHdr.indexOf('กลุ่มงาน');
    var eUpd    = eHdr.indexOf('อัปเดตล่าสุด');

    var targetEmpRow = -1;
    for (var j = 0; j < eData.length; j++) {
      if (String(eData[j][eName >= 0 ? eName : 0] || '').trim() === fullName) {
        targetEmpRow = j + 2;
        break;
      }
    }

    var now = new Date();
    if (targetEmpRow > 0) {
      if (eLine >= 0) empSheet.getRange(targetEmpRow, eLine + 1).setValue(lineId);
      if (eStat >= 0) empSheet.getRange(targetEmpRow, eStat + 1).setValue(status);
      if (eUpd  >= 0) empSheet.getRange(targetEmpRow, eUpd  + 1).setValue(now);
    } else {
      var newRow = new Array(Math.max(eHdr.length, 7)).fill('');
      if (eName >= 0) newRow[eName] = fullName;
      if (eLine >= 0) newRow[eLine] = lineId;
      if (eStat >= 0) newRow[eStat] = status;
      if (ePos  >= 0) newRow[ePos]  = String(data.position   || '').trim();
      if (eMis  >= 0) newRow[eMis]  = String(data.mission    || '').trim();
      if (eDept >= 0) newRow[eDept] = String(data.department || '').trim();
      if (eUpd  >= 0) newRow[eUpd]  = now;
      empSheet.appendRow(newRow);
    }
    SpreadsheetApp.flush();

    Logger.log('✅ saveEmployeeLineData: ' + fullName + ' lineId=' + lineId + ' status=' + status);
    return { success: true };
  } catch (e) {
    Logger.log('❌ saveEmployeeLineData: ' + e);
    return { success: false, message: e.toString() };
  }
}
// ============================================
// === addManualEmployee(data, token) ===
// เพิ่มพนักงานใหม่ที่ไม่มีข้อมูลในระบบ (อบรมกรณีพิเศษ)
// ============================================
function addManualEmployee(data, token) {
  if (!verifyAdminToken(token)) {
    return { success: false, message: 'ไม่มีสิทธิ์ดำเนินการ กรุณา Login ใหม่' };
  }
  try {
    var ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
    var fullName = String(data.fullName   || '').trim();
    var lineId   = String(data.lineId     || '').trim();
    var position = String(data.position   || '').trim();
    var mission  = String(data.mission    || '').trim();
    var dept     = String(data.department || '').trim();
    var status   = String(data.status     || 'เปิดใช้งาน').trim();

    if (!fullName) return { success: false, message: 'กรุณาระบุชื่อ-สกุล' };

    // ── ตรวจสอบ/สร้าง Sheet "ข้อมูลพนักงาน" ──
    var empSheet = ss.getSheetByName('ข้อมูลพนักงาน');
    if (!empSheet) {
      empSheet = ss.insertSheet('ข้อมูลพนักงาน');
      var headers = ['ชื่อ-สกุล','LINE User ID','สถานะแจ้งเตือน','ตำแหน่ง','กลุ่มภารกิจ','กลุ่มงาน','อัปเดตล่าสุด'];
      empSheet.getRange(1, 1, 1, headers.length)
              .setValues([headers]).setFontWeight('bold')
              .setBackground('#1e3a8a').setFontColor('#ffffff');
      empSheet.setFrozenRows(1);
    }

    var eHdr = empSheet.getRange(1, 1, 1, empSheet.getLastColumn()).getValues()[0]
                       .map(function(h) { return String(h || '').trim(); });
    var eName = eHdr.indexOf('ชื่อ-สกุล');
    var eLine = eHdr.indexOf('LINE User ID');
    var eStat = eHdr.indexOf('สถานะแจ้งเตือน');
    var ePos  = eHdr.indexOf('ตำแหน่ง');
    var eMis  = eHdr.indexOf('กลุ่มภารกิจ');
    var eDept = eHdr.indexOf('กลุ่มงาน');
    var eUpd  = eHdr.indexOf('อัปเดตล่าสุด');

    // ── เช็คซ้ำ ──
    var eData = empSheet.getLastRow() > 1
      ? empSheet.getRange(2, 1, empSheet.getLastRow() - 1, empSheet.getLastColumn()).getValues()
      : [];

    for (var j = 0; j < eData.length; j++) {
      if (String(eData[j][eName >= 0 ? eName : 0] || '').trim() === fullName) {
        var ur = j + 2;
        if (eLine >= 0 && lineId)   empSheet.getRange(ur, eLine + 1).setValue(lineId);
        if (eStat >= 0)             empSheet.getRange(ur, eStat + 1).setValue(status);
        if (ePos  >= 0 && position) empSheet.getRange(ur, ePos  + 1).setValue(position);
        if (eMis  >= 0 && mission)  empSheet.getRange(ur, eMis  + 1).setValue(mission);
        if (eDept >= 0 && dept)     empSheet.getRange(ur, eDept + 1).setValue(dept);
        if (eUpd  >= 0)             empSheet.getRange(ur, eUpd  + 1).setValue(new Date());
        SpreadsheetApp.flush();
        Logger.log('✅ addManualEmployee (อัปเดต): ' + fullName);
        return { success: true, action: 'updated' };
      }
    }

    // ── เพิ่มใหม่ ──
    var newRow = new Array(Math.max(eHdr.length, 7)).fill('');
    if (eName >= 0) newRow[eName] = fullName;
    if (eLine >= 0) newRow[eLine] = lineId;
    if (eStat >= 0) newRow[eStat] = status;
    if (ePos  >= 0) newRow[ePos]  = position;
    if (eMis  >= 0) newRow[eMis]  = mission;
    if (eDept >= 0) newRow[eDept] = dept;
    if (eUpd  >= 0) newRow[eUpd]  = new Date();
    empSheet.appendRow(newRow);
    SpreadsheetApp.flush();

    // ── sync LINE ID ไปชีทอบรมฯ ถ้ามีชื่อตรงกัน ──
    if (lineId) {
      var trainSheet = ss.getSheetByName('อบรม ปีงบประมาณ 2569');
      if (trainSheet && trainSheet.getLastRow() > 1) {
        var tData = trainSheet.getDataRange().getValues();
        var tHdr  = tData[0].map(function(h) { return String(h || '').trim(); });
        var tName = tHdr.indexOf('ชื่อ-สกุล');
        var tLine = tHdr.indexOf('LINE ID');
        if (tName >= 0 && tLine >= 0) {
          for (var k = 1; k < tData.length; k++) {
            if (String(tData[k][tName] || '').trim() === fullName) {
              trainSheet.getRange(k + 1, tLine + 1).setValue(lineId);
            }
          }
          SpreadsheetApp.flush();
        }
      }
    }

    Logger.log('✅ addManualEmployee (เพิ่มใหม่): ' + fullName);
    return { success: true, action: 'added' };
  } catch (e) {
    Logger.log('❌ addManualEmployee: ' + e);
    return { success: false, message: e.toString() };
  }
}
// ============================================
// === 4. searchStaffByName(term) ===
// ============================================
function searchStaffByName(term) {
  try {
    if (!term || term.length < 2) return { success: true, data: [] };
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Registrations');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };

    var data      = sheet.getDataRange().getValues();
    var hdr       = data[0].map(function(h) { return String(h || '').trim(); });
    
    var cName = getColIdx(hdr, 'Full Name', 'ชื่อ-สกุล', 4);
    var cPos  = getColIdx(hdr, 'Position', 'ตำแหน่ง', 5);
    var cMis  = getColIdx(hdr, 'Mission Group', 'กลุ่มภารกิจ', 6);
    var cDept = getColIdx(hdr, 'Department', 'กลุ่มงาน', 7);
    var termLower = term.toLowerCase();

    var seen = {}, results = [];
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][cName] || '').trim();
      if (!name || seen[name]) continue;
      if (name.toLowerCase().indexOf(termLower) === -1) continue;
      
      seen[name] = true;
      results.push({
        fullName    : name,
        position    : String(data[i][cPos]  || '').trim(),
        missionGroup: String(data[i][cMis]  || '').trim(),
        department  : String(data[i][cDept] || '').trim()
      });
      if (results.length >= 10) break;
    }
    return { success: true, data: results };
  } catch (e) { return { success: false, message: e.toString() }; }
}
// ============================================
// === 5. getPersonalData(name) ===
// ============================================
function getPersonalData(name) {
  try {
    if (!name) return { success: false, message: 'กรุณาระบุชื่อ' };
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Registrations');
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, data: { summary: { totalRegistrations: 0, totalHours: 0 }, registrations: [] } };
    }

    var data  = sheet.getDataRange().getValues();
    var hdr   = data[0].map(function(h) { return String(h || '').trim(); });
    
    var cName = getColIdx(hdr, 'Full Name', 'ชื่อ-สกุล', 4);
    var cDate = getColIdx(hdr, 'Date', 'วันที่', 1);
    var cEnd  = getColIdx(hdr, 'End Date', 'ถึงวันที่', 2);
    var cTime = getColIdx(hdr, 'Time', 'เวลา', 3);
    var cPos  = getColIdx(hdr, 'Position', 'ตำแหน่ง', 5);
    var cMis  = getColIdx(hdr, 'Mission Group', 'กลุ่มภารกิจ', 6);
    var cDept = getColIdx(hdr, 'Department', 'กลุ่มงาน', 7);
    var cCat  = getColIdx(hdr, 'Category', 'หมวดหมู่', 8);
    var cTopic= getColIdx(hdr, 'Topic', 'หัวข้อ/หลักสูตร', 9);
    var cHrs  = getColIdx(hdr, 'Hours', 'จำนวนชั่วโมง', 10);
    var cLoc  = getColIdx(hdr, 'Location', 'สถานที่', 13);
    var cSum  = getColIdx(hdr, 'Summary', 'สรุป', 11);
    var cSug  = getColIdx(hdr, 'Suggestion', 'ข้อเสนอแนะ', 12);
    var cFN   = hdr.indexOf('ชื่อไฟล์แนบ');
    var cFU   = hdr.indexOf('URL ไฟล์แนบ');

    var nameLower = name.toLowerCase();
    var results = [];
    var totalHours = 0;

    for (var i = 1; i < data.length; i++) {
      var rowName = String(data[i][cName] || '').trim();
      if (rowName.toLowerCase() !== nameLower) continue;

      var hrs = parseFloat(data[i][cHrs]) || 0;
      var rowCat = String(data[i][cCat] || '').trim();
      if (rowCat !== NO_CREDIT_CAT) totalHours += hrs;

      var d = data[i][cDate] instanceof Date ? data[i][cDate] : new Date(data[i][cDate]);
      var dStr = isNaN(d.getTime()) ? '' : Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
      
      var e = data[i][cEnd] instanceof Date ? data[i][cEnd] : new Date(data[i][cEnd]);
      var eStr = isNaN(e.getTime()) ? '' : Utilities.formatDate(e, 'Asia/Bangkok', 'yyyy-MM-dd');

      results.push({
        fullName    : rowName,
        date        : dStr,
        endDate     : eStr,
        time        : String(data[i][cTime] || '').trim(),
        position    : String(data[i][cPos]  || '').trim(),
        missionGroup: String(data[i][cMis]  || '').trim(),
        department  : String(data[i][cDept] || '').trim(),
        category    : String(data[i][cCat]  || '').trim(),
        topic       : String(data[i][cTopic]|| '').trim(),
        hours       : hrs,
        location    : String(data[i][cLoc]  || '').trim(),
        summary     : String(data[i][cSum]  || '').trim(),
        suggestion  : String(data[i][cSug]  || '').trim(),
        fileNames   : cFN !== -1 ? String(data[i][cFN] || '').trim() : '',
        fileUrls    : cFU !== -1 ? String(data[i][cFU] || '').trim() : ''
      });
    }

    return {
      success: true,
      data: {
        summary: { fullName: name, totalRegistrations: results.length, totalHours: totalHours },
        registrations: results
      }
    };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getFilteredDashboard(year, month, missionGroup) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Registrations');
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, data: { summary: { numRegistrations: 0, totalHours: 0 }, registrations: [] } };
    }

    var data  = sheet.getDataRange().getValues();
    var hdr   = data[0].map(function(h) { return String(h || '').trim(); });
    
    var cName = getColIdx(hdr, 'Full Name', 'ชื่อ-สกุล', 4);
    var cDate = getColIdx(hdr, 'Date', 'วันที่', 1);
    var cEnd  = getColIdx(hdr, 'End Date', 'ถึงวันที่', 2);
    var cTime = getColIdx(hdr, 'Time', 'เวลา', 3);
    var cPos  = getColIdx(hdr, 'Position', 'ตำแหน่ง', 5);
    var cMis  = getColIdx(hdr, 'Mission Group', 'กลุ่มภารกิจ', 6);
    var cDept = getColIdx(hdr, 'Department', 'กลุ่มงาน', 7);
    var cCat  = getColIdx(hdr, 'Category', 'หมวดหมู่', 8);
    var cTopic= getColIdx(hdr, 'Topic', 'หัวข้อ/หลักสูตร', 9);
    var cHrs  = getColIdx(hdr, 'Hours', 'จำนวนชั่วโมง', 10);
    var cLoc  = getColIdx(hdr, 'Location', 'สถานที่', 13);
    var cFN   = hdr.indexOf('ชื่อไฟล์แนบ');
    var cFU   = hdr.indexOf('URL ไฟล์แนบ');

    var filterYear  = year  ? parseInt(year)  : 0;
    var filterMonth = month ? parseInt(month) : 0;
    var filterMis   = missionGroup ? String(missionGroup).trim() : '';

    var results = [];
    var totalHours = 0;

    for (var i = 1; i < data.length; i++) {
      var dateVal = data[i][cDate];
      if (!dateVal) continue;
      var d = dateVal instanceof Date ? dateVal : new Date(dateVal);
      if (isNaN(d.getTime())) continue;

      var rowYear  = d.getFullYear();
      var rowMonth = d.getMonth() + 1;

      if (filterYear  && rowYear  !== filterYear)  continue;
      if (filterMonth && rowMonth !== filterMonth) continue;

      var rowMis = String(data[i][cMis] || '').trim();
      if (filterMis && rowMis !== filterMis) continue;

      var hrs = parseFloat(data[i][cHrs]) || 0;
      var rowCat = String(data[i][cCat] || '').trim();
      if (rowCat !== NO_CREDIT_CAT) totalHours += hrs;

      var e = data[i][cEnd] instanceof Date ? data[i][cEnd] : new Date(data[i][cEnd]);
      var eStr = isNaN(e.getTime()) ? '' : Utilities.formatDate(e, 'Asia/Bangkok', 'yyyy-MM-dd');

      results.push({
        fullName    : String(data[i][cName]  || '').trim(),
        date        : Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd'),
        endDate     : eStr,
        time        : String(data[i][cTime]  || '').trim(),
        position    : String(data[i][cPos]   || '').trim(),
        missionGroup: rowMis,
        department  : String(data[i][cDept]  || '').trim(),
        category    : String(data[i][cCat]   || '').trim(),
        topic       : String(data[i][cTopic] || '').trim(),
        hours       : hrs,
        location    : String(data[i][cLoc]   || '').trim(),
        fileNames   : cFN !== -1 ? String(data[i][cFN] || '').trim() : '',
        fileUrls    : cFU !== -1 ? String(data[i][cFU] || '').trim() : ''
      });
    }

    return {
      success: true,
      data: {
        summary      : { numRegistrations: results.length, totalHours: totalHours },
        registrations: results
      }
    };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// ============================================
// === 7. getSimpleDashboard() ===
// ============================================
function getSimpleDashboard() {
  return getFilteredDashboard('', '', '');
}

// ============================================
// === 8. getGroupSummary(groupName) ===
// ============================================
function getGroupSummary(groupName) {
  try {
    if (!groupName) return { success: false, message: 'กรุณาระบุกลุ่มภารกิจ' };
    var result = getFilteredDashboard('', '', groupName);
    if (!result.success) return result;
    
    return {
      success: true,
      data: {
        summary: {
          numRegistrations: result.data.summary.numRegistrations,
          totalHours      : result.data.summary.totalHours
        },
        registrations: result.data.registrations
      }
    };
  } catch (e) { return { success: false, message: e.toString() }; }
}
