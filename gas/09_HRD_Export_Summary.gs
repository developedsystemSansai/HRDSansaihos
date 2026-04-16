 */
function getHrdSummaryDataEnhanced(params) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var ts = ss.getSheetByName('อบรม ปีงบประมาณ 2569');
    
    if (!ts) {
      return { success: false, message: 'ไม่พบ Sheet อบรมฯ', data: [], summary: _emptySum() };
    }
 
    var vals = ts.getDataRange().getValues();
    var hdr  = vals[0];
 
    function col(name, alt, def) {
      var i = hdr.indexOf(name);
      if (i < 0 && alt) i = hdr.indexOf(alt);
      return (i < 0 ? def : i);
    }
 
    function gStr(row, c) { return (c < 0 || !row[c]) ? '' : String(row[c]).trim(); }
    function gNum(row, c) { var v = parseFloat(row[c] || 0); return isNaN(v) ? 0 : v; }
    function gDate(row, c) {
      if (c < 0) return null;
      var v = row[c];
      if (!v) return null;
      if (v instanceof Date) return v;
      if (typeof v === 'string') {
        var parts = v.split('/');
        if (parts.length === 3) {
          return new Date(+parts[2], +parts[1] - 1, +parts[0]);
        }
      }
      return null;
    }
 
    function fmtDate(d) {
      if (!d || !(d instanceof Date)) return '';
      var dd = ('0' + d.getDate()).slice(-2);
      var mm = ('0' + (d.getMonth()+1)).slice(-2);
      var yy = d.getFullYear();
      return dd + '/' + mm + '/' + yy;
    }
 
    var cName  = col('ชื่อ-สกุล');
    var cPos   = col('ตำแหน่ง');
    var cMis   = col('กลุ่มภารกิจ');
    var cDept  = col('กลุ่มงาน');
    var cTopic = col('เรื่อง/หลักสูตร', 'หลักสูตร/เรื่อง');
    var cSDate = col('วันที่เริ่ม');
    var cEDate = col('วันที่สิ้นสุด');
    var cVenue = col('สถานที่');
    var cProv  = col('จังหวัด');
    var cHm5   = col('เลขที่ ชม.5');
    var cMemo  = col('เลขที่บันทึก');
    var cBudget= col('แหล่งงบประมาณ');
    var cCostReg = col('ค่าลงทะเบียน');
    var cCostTrav= col('ค่าเบี้ยเลี้ยง');
    var cCostHotel=col('ค่าที่พัก');
    var cCostTrans=col('ค่าพาหนะ');
    var cCostProj =col('เงินโครงการ');
    var cTotalCost=col('รวมค่าใช้จ่าย');
    var cRec   = col('ผู้บันทึก');
    var cTs    = col('วันที่บันทึก', 'Timestamp');
    
    // คอลัมน์เอกสาร - เพิ่มการตรวจสอบไฟล์แนบ
    var cFileName = col('ชื่อไฟล์แนบ', '', -1);
    var cFileURL  = col('URL ไฟล์แนบ', '', -1);
    var cDocSum   = col('เอกสารสรุปผล', '', -1);
    var cDocCon   = col('สัญญา', '', -1);
    var cDocCert  = col('วุฒิบัตร', '', -1);
    var cDocRec   = col('ใบเสร็จ', '', -1);
    var cDocHm5   = col('บันทึก ชม.5', '', -1);
    var cDocApp   = col('บันทึกขออนุมัติ', '', -1);
    var cDocCom   = col('บันทึกความ+ชม.5', '', -1);
 
    var filterStart = params.filterStart ? new Date(params.filterStart) : null;
    var filterEnd   = params.filterEnd   ? new Date(params.filterEnd)   : null;
    var results = [];
 
    for (var i = 1; i < vals.length; i++) {
      var row = vals[i];
      if (!row[cName] && !row[cTopic]) continue;
 
      var sDate = gDate(row, cSDate);
      if (filterStart && sDate && sDate < filterStart) continue;
      if (filterEnd   && sDate && sDate > filterEnd)   continue;
 
      var name    = gStr(row, cName);
      var mission = gStr(row, cMis);
      var hm5     = gStr(row, cHm5);
      var memo    = gStr(row, cMemo);
      var budget  = gStr(row, cBudget);
      var totalCost = gNum(row, cTotalCost);
      var costReg   = gNum(row, cCostReg);
      var costTrav  = gNum(row, cCostTrav);
      var costHotel = gNum(row, cCostHotel);
      var costTrans = gNum(row, cCostTrans);
      var costProj  = gNum(row, cCostProj);
      
      var topic = gStr(row, cTopic);
      
      // 🆕 ตรวจสอบไฟล์แนบจากคอลัมน์
      var hasFileName = cFileName >= 0 && gStr(row, cFileName).length > 0;
      var hasFileURL  = cFileURL  >= 0 && gStr(row, cFileURL).length > 0;
      var hasAttachment = hasFileName || hasFileURL;
      
      // 🆕 ตรวจสอบเอกสารสรุปผลจาก Registrations
      var docSummaryFromReg = checkRegistrationDocuments(name, topic, sDate);
      
      // ตรวจสอบเอกสารแต่ละประเภท
      var docSummary = (cDocSum >= 0 && gStr(row, cDocSum).length > 0) || docSummaryFromReg;
      var docContract = cDocCon >= 0 && gStr(row, cDocCon).length > 0;
      var docCert = cDocCert >= 0 && gStr(row, cDocCert).length > 0;
      var docReceipt = cDocRec >= 0 && gStr(row, cDocRec).length > 0;
      
      // 🆕 ตรวจสอบ ชม.5, บันทึกขออนุมัติ, บันทึกความ+ชม.5 โดยดูจากทั้งเซลล์และไฟล์แนบ
      var docHm5 = (cDocHm5 >= 0 && gStr(row, cDocHm5).length > 0) || hasAttachment;
      var docApproval = (cDocApp >= 0 && gStr(row, cDocApp).length > 0) || hasAttachment;
      var docCombined = (cDocCom >= 0 && gStr(row, cDocCom).length > 0) || hasAttachment;
 
      results.push({
        name     : name,
        position : gStr(row, cPos),
        mission  : mission,
        dept     : gStr(row, cDept),
        topic    : topic,
        startDate: fmtDate(sDate),
        endDate  : fmtDate(gDate(row, cEDate)),
        venue    : gStr(row, cVenue),
        province : gStr(row, cProv),
        hm5      : hm5,
        memo     : memo,
        budget   : budget,
        totalCost: totalCost,
        costReg  : costReg,
        costTravel: costTrav,
        costHotel: costHotel,
        costTrans: costTrans,
        costProj : costProj,
        docSummary : docSummary,
        docContract: docContract,
        docCert    : docCert,
        docReceipt : docReceipt,
        docHm5     : docHm5,
        docApproval: docApproval,
        docCombined: docCombined,
        recorder : gStr(row, cRec),
        savedAt  : fmtDate(gDate(row, cTs))
      });
    }
 
    var summary = _buildSummaryEnhanced(results);
    Logger.log('✅ getHrdSummaryDataEnhanced: ' + results.length + ' รายการ');
    return { success: true, data: results, summary: summary };
 
  } catch(e) {
    Logger.log('❌ getHrdSummaryDataEnhanced: ' + e.toString());
    return { success: false, message: e.toString(), data: [], summary: _emptySum() };
  }
}

/**
 * _buildSummaryEnhanced(data)
 * สร้าง object สรุปสถิติจาก array ข้อมูล พร้อมสรุปตามหมวดหมู่
 */
function _buildSummaryEnhanced(data) {
  var totalCost = 0;
  var personSet = {};
  var missionMap = {};
  var budgetMap  = {};
  var categoryMap = {}; // 🆕 เพิ่มการจัดกลุ่มตามหมวดหมู่
  var docComplete = 0;
 
  data.forEach(function(r) {
    totalCost += (r.totalCost || 0);
    if (r.name)    personSet[r.name] = true;
    if (r.mission) {
      missionMap[r.mission] = (missionMap[r.mission] || 0) + 1;
      
      // 🆕 สรุปตามหมวดหมู่/กลุ่มภารกิจ
      if (!categoryMap[r.mission]) {
        categoryMap[r.mission] = {
          count: 0,
          totalCost: 0,
          docComplete: 0,
          persons: {}
        };
      }
      categoryMap[r.mission].count++;
      categoryMap[r.mission].totalCost += (r.totalCost || 0);
      if (r.docHm5 || r.docApproval || r.docCombined) {
        categoryMap[r.mission].docComplete++;
      }
      if (r.name) {
        categoryMap[r.mission].persons[r.name] = true;
      }
    }
    if (r.budget)  budgetMap[r.budget] = (budgetMap[r.budget] || 0) + (r.totalCost || 0);
    if (r.docHm5 || r.docApproval || r.docCombined) docComplete++;
  });
 
  return {
    totalRecords : data.length,
    totalPersons : Object.keys(personSet).length,
    totalCost    : totalCost,
    docComplete  : docComplete,
    missionBreakdown : missionMap,
    budgetBreakdown  : budgetMap,
    categoryBreakdown: categoryMap // 🆕
  };
}
 
function _emptySum() {
  return { 
    totalRecords: 0, 
    totalPersons: 0, 
    totalCost: 0, 
    docComplete: 0, 
    missionBreakdown: {}, 
    budgetBreakdown: {},
    categoryBreakdown: {}
  };
}

// ============================================================
// === exportHrdToSheet(params)
// === สร้าง Sheet สรุปผลใหม่ใน Spreadsheet และ return URL
// === (ทางเลือก: ถ้าต้องการ export ผ่าน GAS backend แทน client-side)
// ============================================================
function exportHrdToSheet(params) {
  try {
    var result = getHrdSummaryData(params);
    if (!result.success || result.data.length === 0) {
      return { success: false, message: 'ไม่พบข้อมูลในช่วงเวลาที่เลือก' };
    }

    var data    = result.data;
    var summary = result.summary;
    var label   = String(params.periodLabel || 'Export').replace(/[\/\\:*?"<>|]/g,'_').substring(0,40);
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ─── ลบ Sheet เก่าถ้ามี ───
    var sheetName = 'Export_' + label;
    var existing  = ss.getSheetByName(sheetName);
    if (existing) ss.deleteSheet(existing);

    var sheet = ss.insertSheet(sheetName);
    var now   = new Date();

    // ─── หัวรายงาน ───
    sheet.getRange('A1').setValue('สรุปผลการพัฒนาบุคลากร (HRD)');
    sheet.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#1e3a8a');
    sheet.getRange('A2').setValue('โรงพยาบาลสันทราย · ช่วงเวลา: ' + label);
    sheet.getRange('A3').setValue('พิมพ์วันที่: ' + Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm'));
    sheet.getRange('A4').setValue('');

    // ─── แถวสรุปสถิติ ───
    var sumRow = 5;
    sheet.getRange(sumRow, 1).setValue('📋 จำนวนรายการ').setFontWeight('bold');
    sheet.getRange(sumRow, 2).setValue(summary.totalRecords);
    sheet.getRange(sumRow, 3).setValue('👤 บุคลากร').setFontWeight('bold');
    sheet.getRange(sumRow, 4).setValue(summary.totalPersons + ' คน');
    sheet.getRange(sumRow, 5).setValue('💰 รวมค่าใช้จ่าย').setFontWeight('bold');
    sheet.getRange(sumRow, 6).setValue(summary.totalCost);
    sheet.getRange(sumRow, 6).setNumberFormat('#,##0.00');
    sheet.getRange(sumRow, 7).setValue('📋 เอกสารครบ').setFontWeight('bold');
    sheet.getRange(sumRow, 8).setValue(summary.docComplete + ' รายการ');

    // ─── Header ตาราง ───
    var headers = [
      'ลำดับ','ชื่อ-สกุล','ตำแหน่ง','กลุ่มภารกิจ','กลุ่มงาน',
      'หลักสูตร/เรื่อง','วันที่เริ่ม','วันที่สิ้นสุด','สถานที่','จังหวัด',
      'เลขที่ ชม.5','เลขที่บันทึก','แหล่งงบ',
      'ค่าลงทะเบียน','ค่าเบี้ยเลี้ยง','ค่าที่พัก','ค่าพาหนะ','เงินโครงการ','รวม',
      'สรุป','สัญญา','วุฒิ','ใบเสร็จ','ชม.5','ขออนุมัติ','รวม+ชม.5',
      'ผู้บันทึก'
    ];
    var hdrRow = 7;
    var hdrRange = sheet.getRange(hdrRow, 1, 1, headers.length);
    hdrRange.setValues([headers]);
    hdrRange.setBackground('#1e3a8a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12);
    sheet.setFrozenRows(hdrRow);

    // ─── ข้อมูล ───
    var rows = data.map(function(r, i) {
      function bool(v) { return v ? '✓' : ''; }
      return [
        i+1, r.name||'', r.position||'', r.mission||'', r.dept||'',
        r.topic||'', r.startDate||'', r.endDate||'', r.venue||'', r.province||'',
        r.hm5||'', r.memo||'', r.budget||'',
        r.costReg||0, r.costTravel||0, r.costHotel||0, r.costTrans||0, r.costProj||0, r.totalCost||0,
        bool(r.docSummary), bool(r.docContract), bool(r.docCert), bool(r.docReceipt),
        bool(r.docHm5), bool(r.docApproval), bool(r.docCombined),
        r.recorder||''
      ];
    });

    if (rows.length > 0) {
      var dataRange = sheet.getRange(hdrRow+1, 1, rows.length, headers.length);
      dataRange.setValues(rows);
      // จัดรูปแบบคอลัมน์เงิน (ลำดับ 14–19 = index 13–18)
      for (var c = 14; c <= 19; c++) {
        sheet.getRange(hdrRow+1, c, rows.length, 1).setNumberFormat('#,##0.00');
      }
      // สลับสีแถว
      for (var r2 = 0; r2 < rows.length; r2++) {
        var bg = (r2 % 2 === 0) ? '#f8fafc' : '#ffffff';
        sheet.getRange(hdrRow+1+r2, 1, 1, headers.length).setBackground(bg);
      }
    }

    // ─── Auto-resize ───
    sheet.autoResizeColumns(1, headers.length);

    // ─── Sheet สรุปตามงบประมาณ ───
    var budgetData = Object.keys(summary.budgetBreakdown).map(function(k) {
      return [k, summary.budgetBreakdown[k]];
    });
    if (budgetData.length > 0) {
      var budgetStartRow = hdrRow + rows.length + 3;
      sheet.getRange(budgetStartRow, 1).setValue('📊 สรุปตามแหล่งงบประมาณ').setFontWeight('bold').setFontColor('#1e3a8a').setFontSize(13);
      sheet.getRange(budgetStartRow+1, 1).setValue('แหล่งงบ').setFontWeight('bold').setBackground('#e2e8f0');
      sheet.getRange(budgetStartRow+1, 2).setValue('รวม (บาท)').setFontWeight('bold').setBackground('#e2e8f0');
      budgetData.forEach(function(row, idx) {
        sheet.getRange(budgetStartRow+2+idx, 1).setValue(row[0]);
        sheet.getRange(budgetStartRow+2+idx, 2).setValue(row[1]).setNumberFormat('#,##0.00');
      });
    }

    SpreadsheetApp.flush();

    var url = 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID
            + '/edit#gid=' + sheet.getSheetId();

    Logger.log('✅ exportHrdToSheet: สร้าง Sheet "' + sheetName + '" สำเร็จ url=' + url);
    return { success: true, url: url, sheetName: sheetName, summary: summary };

  } catch(e) {
    Logger.log('❌ exportHrdToSheet: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}
// ============================================================
// === 🆕 exportHrdToSheetEnhanced(params)
// === สร้าง Sheet สรุปผลใหม่พร้อมตรวจสอบเอกสารและสรุปตามหมวดหมู่
// ============================================================
function exportHrdToSheetEnhanced(params) {
  try {
    var result = getHrdSummaryDataEnhanced(params);
    if (!result.success || result.data.length === 0) {
      return { success: false, message: 'ไม่พบข้อมูลในช่วงเวลาที่เลือก' };
    }
 
    var data    = result.data;
    var summary = result.summary;
    var label   = String(params.periodLabel || 'Export').replace(/[\/\\:*?"<>|]/g,'_').substring(0,40);
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
 
    // ─── ลบ Sheet เก่าถ้ามี ───
    var sheetName = 'Export_' + label;
    var existing  = ss.getSheetByName(sheetName);
    if (existing) ss.deleteSheet(existing);
 
    var sheet = ss.insertSheet(sheetName);
    var now   = new Date();
 
    // ─── หัวรายงาน ───
    sheet.getRange('A1').setValue('สรุปผลการพัฒนาบุคลากร (HRD)');
    sheet.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#1e3a8a');
    sheet.getRange('A2').setValue('โรงพยาบาลสันทราย · ช่วงเวลา: ' + label);
    sheet.getRange('A3').setValue('พิมพ์วันที่: ' + Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm'));
    sheet.getRange('A4').setValue('');
 
    // ─── แถวสรุปสถิติ ───
    var sumRow = 5;
    sheet.getRange(sumRow, 1).setValue('📋 จำนวนรายการ').setFontWeight('bold');
    sheet.getRange(sumRow, 2).setValue(summary.totalRecords);
    sheet.getRange(sumRow, 3).setValue('👤 บุคลากร').setFontWeight('bold');
    sheet.getRange(sumRow, 4).setValue(summary.totalPersons + ' คน');
    sheet.getRange(sumRow, 5).setValue('💰 รวมค่าใช้จ่าย').setFontWeight('bold');
    sheet.getRange(sumRow, 6).setValue(summary.totalCost);
    sheet.getRange(sumRow, 6).setNumberFormat('#,##0.00');
    sheet.getRange(sumRow, 7).setValue('📋 เอกสารครบ').setFontWeight('bold');
    sheet.getRange(sumRow, 8).setValue(summary.docComplete + ' รายการ');
 
    // ─── Header ตาราง ───
    var headers = [
      'ลำดับ','ชื่อ-สกุล','ตำแหน่ง','กลุ่มภารกิจ','กลุ่มงาน',
      'หลักสูตร/เรื่อง','วันที่เริ่ม','วันที่สิ้นสุด','สถานที่','จังหวัด',
      'เลขที่ ชม.5','เลขที่บันทึก','แหล่งงบ',
      'ค่าลงทะเบียน','ค่าเบี้ยเลี้ยง','ค่าที่พัก','ค่าพาหนะ','เงินโครงการ','รวม',
      'เอกสารสรุปผล','สัญญา','วุฒิบัตร','ใบเสร็จ','ชม.5','บันทึกขออนุมัติ','บันทึกความ+ชม.5',
      'ผู้บันทึก'
    ];
    var hdrRow = 7;
    var hdrRange = sheet.getRange(hdrRow, 1, 1, headers.length);
    hdrRange.setValues([headers]);
    hdrRange.setBackground('#1e3a8a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12);
    sheet.setFrozenRows(hdrRow);
 
    // ─── ข้อมูล ───
    var rows = data.map(function(r, i) {
      function bool(v) { return v ? '✓' : ''; }
      return [
        i+1, r.name||'', r.position||'', r.mission||'', r.dept||'',
        r.topic||'', r.startDate||'', r.endDate||'', r.venue||'', r.province||'',
        r.hm5||'', r.memo||'', r.budget||'',
        r.costReg||0, r.costTravel||0, r.costHotel||0, r.costTrans||0, r.costProj||0, r.totalCost||0,
        bool(r.docSummary), bool(r.docContract), bool(r.docCert), bool(r.docReceipt),
        bool(r.docHm5), bool(r.docApproval), bool(r.docCombined),
        r.recorder||''
      ];
    });
 
    if (rows.length > 0) {
      var dataRange = sheet.getRange(hdrRow+1, 1, rows.length, headers.length);
      dataRange.setValues(rows);
      // จัดรูปแบบคอลัมน์เงิน (ลำดับ 14–19 = index 13–18)
      for (var c = 14; c <= 19; c++) {
        sheet.getRange(hdrRow+1, c, rows.length, 1).setNumberFormat('#,##0.00');
      }
      // สลับสีแถว
      for (var r2 = 0; r2 < rows.length; r2++) {
        var bg = (r2 % 2 === 0) ? '#f8fafc' : '#ffffff';
        sheet.getRange(hdrRow+1+r2, 1, 1, headers.length).setBackground(bg);
      }
    }
 
    // ─── Auto-resize ───
    sheet.autoResizeColumns(1, headers.length);
 
    var nextStartRow = hdrRow + rows.length + 3;
 
    // ─── 🆕 สรุปตามหมวดหมู่/กลุ่มภารกิจ ───
    var categoryData = Object.keys(summary.categoryBreakdown).map(function(k) {
      var cat = summary.categoryBreakdown[k];
      return [
        k, 
        cat.count, 
        Object.keys(cat.persons).length,
        cat.totalCost,
        cat.docComplete
      ];
    });
    
    if (categoryData.length > 0) {
      sheet.getRange(nextStartRow, 1).setValue('📊 สรุปตามหมวดหมู่/กลุ่มภารกิจ')
        .setFontWeight('bold').setFontColor('#1e3a8a').setFontSize(13);
      
      var catHeaders = ['กลุ่มภารกิจ', 'จำนวนรายการ', 'จำนวนบุคลากร', 'รวมค่าใช้จ่าย (บาท)', 'เอกสารครบ'];
      var catHeaderRange = sheet.getRange(nextStartRow+1, 1, 1, catHeaders.length);
      catHeaderRange.setValues([catHeaders]);
      catHeaderRange.setFontWeight('bold').setBackground('#e2e8f0');
      
      categoryData.forEach(function(row, idx) {
        sheet.getRange(nextStartRow+2+idx, 1).setValue(row[0]);
        sheet.getRange(nextStartRow+2+idx, 2).setValue(row[1]);
        sheet.getRange(nextStartRow+2+idx, 3).setValue(row[2]);
        sheet.getRange(nextStartRow+2+idx, 4).setValue(row[3]).setNumberFormat('#,##0.00');
        sheet.getRange(nextStartRow+2+idx, 5).setValue(row[4]);
      });
      
      nextStartRow = nextStartRow + categoryData.length + 4;
    }
 
    // ─── สรุปตามงบประมาณ ───
    var budgetData = Object.keys(summary.budgetBreakdown).map(function(k) {
      return [k, summary.budgetBreakdown[k]];
    });
    
    if (budgetData.length > 0) {
      sheet.getRange(nextStartRow, 1).setValue('📊 สรุปตามแหล่งงบประมาณ')
        .setFontWeight('bold').setFontColor('#1e3a8a').setFontSize(13);
      sheet.getRange(nextStartRow+1, 1).setValue('แหล่งงบ').setFontWeight('bold').setBackground('#e2e8f0');
      sheet.getRange(nextStartRow+1, 2).setValue('รวม (บาท)').setFontWeight('bold').setBackground('#e2e8f0');
      
      budgetData.forEach(function(row, idx) {
        sheet.getRange(nextStartRow+2+idx, 1).setValue(row[0]);
        sheet.getRange(nextStartRow+2+idx, 2).setValue(row[1]).setNumberFormat('#,##0.00');
      });
    }
 
    SpreadsheetApp.flush();
 
    var url = 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID
            + '/edit#gid=' + sheet.getSheetId();
 
    Logger.log('✅ exportHrdToSheetEnhanced: สร้าง Sheet "' + sheetName + '" สำเร็จ url=' + url);
    return { success: true, url: url, sheetName: sheetName, summary: summary };
 
  } catch(e) {
    Logger.log('❌ exportHrdToSheetEnhanced: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}
 
// ============================================================
// === 🔄 ฟังก์ชัน wrapper เพื่อความเข้ากันได้กับโค้ดเดิม
// === เรียกใช้ฟังก์ชันใหม่แทนฟังก์ชันเดิม
// ============================================================
 
function getHrdSummaryData(params) {
  return getHrdSummaryDataEnhanced(params);
}
 
function exportHrdToSheet(params) {
  return exportHrdToSheetEnhanced(params);
}
// ============================================================
// === 🆕 ฟังก์ชันตรวจสอบไฟล์แนบจาก Registrations
// === ใช้ตรวจสอบว่ามีเอกสารสรุปผลจากผู้ลงทะเบียนหรือไม่
// ============================================================
 
/**
 * checkRegistrationDocuments(fullName, topic, startDate)
 * ตรวจสอบว่ามีเอกสารสรุปผลจาก sheet Registrations หรือไม่
 * 
 * @param {string} fullName - ชื่อ-สกุล
 * @param {string} topic - หัวข้อการอบรม
 * @param {Date} startDate - วันที่เริ่ม
 * @return {boolean} - มีเอกสารสรุปผล = true
 */
function checkRegistrationDocuments(fullName, topic, startDate) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var regSheet = ss.getSheetByName('Registrations');
    
    if (!regSheet) {
      Logger.log('⚠️ ไม่พบ sheet Registrations');
      return false;
    }
    
    var data = regSheet.getDataRange().getValues();
    var headers = data[0];
    
    // หาตำแหน่งคอลัมน์
    var colFullName = headers.indexOf('Full Name');
    var colTopic = headers.indexOf('Topic');
    var colDateIn = headers.indexOf('Date-in');
    var colFileName = headers.indexOf('ชื่อไฟล์แนบ');
    var colFileURL = headers.indexOf('URL ไฟล์แนบ');
    
    if (colFullName < 0 || colTopic < 0) {
      return false;
    }
    
    // ค้นหาแถวที่ตรงกัน
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // ตรวจสอบความตรงกัน
      var nameMatch = String(row[colFullName] || '').trim() === String(fullName || '').trim();
      var topicMatch = String(row[colTopic] || '').trim() === String(topic || '').trim();
      
      if (nameMatch && topicMatch) {
        // ตรวจสอบว่ามีไฟล์แนบหรือไม่
        var hasFileName = colFileName >= 0 && String(row[colFileName] || '').trim().length > 0;
        var hasFileURL = colFileURL >= 0 && String(row[colFileURL] || '').trim().length > 0;
        
        if (hasFileName || hasFileURL) {
          return true;
        }
      }
    }
    
    return false;
    
  } catch (e) {
    Logger.log('❌ checkRegistrationDocuments error: ' + e.toString());
    return false;
  }
}
