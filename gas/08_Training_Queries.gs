// ============================================
// === getAllTrainingData() — ดึงทุกแถว ไม่มีเงื่อนไขกรอง ===
// เรียงจากลำดับมากสุด (ล่าสุด) ไปน้อยสุด
// ใช้สำหรับ: รายการล่าสุด (sf tab) และ ชม.5 tab
// ============================================
function getAllTrainingData() {
  try {
    var info    = getOrInitTrainingSheet();
    var sheet   = info.sheet;
    var hMap    = info.headerMap;
    var lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return { success: true, data: [] };
    }

    var dataRows = lastRow - 1;

    // batch read ทุก column
    var allData = sheet.getRange(2, 1, dataRows, info.headers.length).getValues();

    // batch read URL formula + note + value จากคอลัมน์ "URL ไฟล์แนบ"
    var urlColIdx   = hMap['URL ไฟล์แนบ'];
    var nameColIdx  = hMap['ชื่อไฟล์แนบ'];
    var urlFormulas = [], urlNotes = [], urlValues2 = [];
    var nameFormulas2 = [], nameNotes2 = [];

    if (urlColIdx !== undefined) {
      var uRange = sheet.getRange(2, urlColIdx + 1, dataRows, 1);
      var uFm    = uRange.getFormulas();
      var uNt    = uRange.getNotes();
      var uVl    = uRange.getValues();
      for (var ri = 0; ri < dataRows; ri++) {
        urlFormulas.push(uFm[ri][0] || '');
        urlNotes.push(uNt[ri][0]    || '');
        urlValues2.push(String(uVl[ri][0] || '').trim());
      }
    }

    // batch read HYPERLINK formula จากคอลัมน์ "ชื่อไฟล์แนบ" เพื่อ fallback
    if (nameColIdx !== undefined) {
      var nRange = sheet.getRange(2, nameColIdx + 1, dataRows, 1);
      var nFm    = nRange.getFormulas();
      var nNt    = nRange.getNotes();
      for (var ri2 = 0; ri2 < dataRows; ri2++) {
        nameFormulas2.push(nFm[ri2][0] || '');
        nameNotes2.push(nNt[ri2][0]    || '');
      }
    }

    // helper สร้าง row object
    function buildRecord(i) {
      var rowRef  = allData[i];
      var hMapRef = hMap;

      var getVal = function(key) {
        var ci = hMapRef[key];
        return ci !== undefined ? rowRef[ci] : undefined;
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

      // ── ดึง URL — ลำดับความสำคัญ: URL column → ชื่อไฟล์แนบ formula → note ──
      var urls = [];

      if (urlColIdx !== undefined) {
        // 1. จาก "URL ไฟล์แนบ" formula
        var uf = urlFormulas[i] || '';
        if (uf && uf.toUpperCase().indexOf('HYPERLINK') !== -1) {
          var mUF = uf.match(/https?:\/\/[^"'\s]+/g);
          if (mUF) urls = urls.concat(mUF);
        }
        // 2. จาก "URL ไฟล์แนบ" note
        var un = urlNotes[i] || '';
        if (un) {
          var mUN = un.match(/https?:\/\/[^\s\n]+/g);
          if (mUN) urls = urls.concat(mUN);
        }
        // 3. จาก "URL ไฟล์แนบ" plain value (หลัง backfill)
        var uv = urlValues2[i] || '';
        if (uv && uv.indexOf('http') === 0) {
          uv.split('\n').forEach(function(u) {
            var t = u.trim();
            if (t) urls.push(t);
          });
        }
      }

      // 4. fallback: ดึงจาก HYPERLINK formula ของ "ชื่อไฟล์แนบ"
      if (urls.length === 0 && nameColIdx !== undefined) {
        var nf = nameFormulas2[i] || '';
        if (nf && nf.toUpperCase().indexOf('HYPERLINK') !== -1) {
          var mNF = nf.match(/https?:\/\/[^"'\s]+/g);
          if (mNF) urls = urls.concat(mNF);
        }
        var nn = nameNotes2[i] || '';
        if (nn) {
          var mNN = nn.match(/https?:\/\/[^\s\n]+/g);
          if (mNN) urls = urls.concat(mNN);
        }
      }

      // deduplicate
      var seen = {};
      urls = urls.filter(function(u) {
        if (!u || seen[u]) return false;
        seen[u] = true;
        return true;
      });

      var seqRaw = gStr('ลำดับ');
      var seqNum = seqRaw !== '' ? seqRaw : String(i + 2);

      return {
        savedAt    : gTs('Timestamp'),
        seqNum     : seqNum,
        name       : gStr('ชื่อ-สกุล') || gStr('ชื่อ'),
        position   : gStr('ตำแหน่ง'),
        mission    : gStr('กลุ่มภารกิจ'),
        dept       : gStr('กลุ่มงาน'),
        topic      : gStr('เรื่อง/หลักสูตร') || gStr('หลักสูตร') || gStr('เรื่อง'),
        startDate  : gDate('วันที่เริ่ม'),
        endDate    : gDate('วันที่สิ้นสุด'),
        venue      : gStr('สถานที่'),
        province   : gStr('จังหวัด'),
        totalCost  : gFloat('รวมค่าใช้จ่าย'),
        budget     : gStr('แหล่งงบประมาณ'),
        fileNames  : gStr('ชื่อไฟล์แนบ'),
        fileUrls   : urls.join('\n'),
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
        rowIndex   : i + 2,
        costReg    : gFloat('ค่าลงทะเบียน'),
        costTravel : gFloat('ค่าเบี้ยเลี้ยง'),
        costHotel  : gFloat('ค่าที่พัก'),
        costTrans  : gFloat('ค่าพาหนะ'),
        costProj   : gFloat('เงินโครงการ')
      };
    }

    // เรียงจากลำดับมากสุด (ล่าสุด) ไปน้อยสุด — รองรับชีทที่เรียงลำดับใดก็ได้
    var seqColIdx3 = hMap['ลำดับ'];
    var allRecords = [];
    for (var i = 0; i < dataRows; i++) {
      var rec = buildRecord(i);
      if (!rec.name && !rec.seqNum) continue;
      var seqV = seqColIdx3 !== undefined ? (parseFloat(allData[i][seqColIdx3]) || 0) : 0;
      allRecords.push({ rec: rec, seq: seqV });
    }
    allRecords.sort(function(a, b) { return b.seq - a.seq; });
    var results = allRecords.map(function(x) { return x.rec; });

    Logger.log('\u2705 getAllTrainingData: คืน ' + results.length + ' รายการ');
    return { success: true, data: results };

  } catch (error) {
    Logger.log('\u274c getAllTrainingData: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}
// ============================================================
// === getRecentFromRegistrations(limit) ===
// ดึงรายการล่าสุดจาก Sheet Registrations (แท็บบันทึก)
// เรียงตาม Timestamp ล่าสุดก่อน
// ============================================================
function getRecentFromRegistrations(limit) {
  try {
    limit = limit || 20;
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Registrations');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };

    var data  = sheet.getDataRange().getValues();
    var hdr   = data[0].map(function(h) { return String(h || '').trim(); });

    var cName  = getColIdx(hdr, 'Full Name',     'ชื่อ-สกุล',       4);
    var cDate  = getColIdx(hdr, 'Date',           'วันที่',           1);
    var cEnd   = getColIdx(hdr, 'End Date',       'ถึงวันที่',        2);
    var cTime  = getColIdx(hdr, 'Time',           'เวลา',             3);
    var cPos   = getColIdx(hdr, 'Position',       'ตำแหน่ง',         5);
    var cMis   = getColIdx(hdr, 'Mission Group',  'กลุ่มภารกิจ',     6);
    var cDept  = getColIdx(hdr, 'Department',     'กลุ่มงาน',        7);
    var cCat   = getColIdx(hdr, 'Category',       'หมวดหมู่',        8);
    var cTopic = getColIdx(hdr, 'Topic',          'หัวข้อ/หลักสูตร', 9);
    var cHrs   = getColIdx(hdr, 'Hours',          'จำนวนชั่วโมง',    10);
    var cLoc   = getColIdx(hdr, 'Location',       'สถานที่',          13);
    var cTs    = getColIdx(hdr, 'Timestamp',      'Timestamp',        0);
    var cFN    = hdr.indexOf('ชื่อไฟล์แนบ');
    var cFU    = hdr.indexOf('URL ไฟล์แนบ');

    // เรียงตาม Timestamp ล่าสุดก่อน
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][cName] || '').trim();
      if (!name) continue;
      var tsVal = data[i][cTs];
      var tsMs  = (tsVal instanceof Date) ? tsVal.getTime()
                : (tsVal ? new Date(tsVal).getTime() : 0);
      if (isNaN(tsMs)) tsMs = 0;
      rows.push({ rowIdx: i, tsMs: tsMs });
    }
    rows.sort(function(a, b) { return b.tsMs - a.tsMs; });

    var results = [];
    var take = Math.min(limit, rows.length);
    for (var ti = 0; ti < take; ti++) {
      var r   = data[rows[ti].rowIdx];
      var hrs = parseFloat(r[cHrs]) || 0;

      var d    = r[cDate] instanceof Date ? r[cDate] : new Date(r[cDate]);
      var dStr = isNaN(d.getTime()) ? '' : Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
      var e    = r[cEnd]  instanceof Date ? r[cEnd]  : new Date(r[cEnd]);
      var eStr = isNaN(e.getTime()) ? '' : Utilities.formatDate(e, 'Asia/Bangkok', 'yyyy-MM-dd');
      var ts   = r[cTs]   instanceof Date ? r[cTs]   : new Date(r[cTs]);
      var tsStr = isNaN(ts.getTime()) ? '' : Utilities.formatDate(ts, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm');

      results.push({
        fullName    : String(r[cName]  || '').trim(),
        date        : dStr,
        endDate     : eStr,
        time        : String(r[cTime]  || '').trim(),
        position    : String(r[cPos]   || '').trim(),
        missionGroup: String(r[cMis]   || '').trim(),
        department  : String(r[cDept]  || '').trim(),
        category    : String(r[cCat]   || '').trim(),
        topic       : String(r[cTopic] || '').trim(),
        hours       : hrs,
        location    : String(r[cLoc]   || '').trim(),
        savedAt     : tsStr,
        tsMs        : rows[ti].tsMs,
        fileNames   : cFN !== -1 ? String(r[cFN] || '').trim() : '',
        fileUrls    : cFU !== -1 ? String(r[cFU] || '').trim() : ''
      });
    }

    Logger.log('✅ getRecentFromRegistrations: คืน ' + results.length + ' รายการ');
    return { success: true, data: results };
  } catch(e) {
    Logger.log('❌ getRecentFromRegistrations: ' + e);
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// === getRecentRecords(limit) — alias → getRecentTrainingRequests ===
// เรียกจาก sfLoadRecentRecords() ในแท็บสรุปอบรม (HRD)
// ============================================================
function getRecentRecords(limit) {
  return getRecentTrainingRequests(limit || 15);
}

// ============================================================
// === trackLinkClick(linkId) ===
// บันทึกจำนวนคลิก link ใน Learn Sidebar และ News Sidebar
// เรียกจาก index.html → trackLearnClick(linkId)
// เก็บใน ScriptProperties: "link_view_" + linkId
// ============================================================
function trackLinkClick(linkId) {
  try {
    if (!linkId) return;
    var props = PropertiesService.getScriptProperties();
    var key   = 'link_view_' + linkId;
    var cur   = parseInt(props.getProperty(key) || '0', 10);
    props.setProperty(key, String(cur + 1));
    Logger.log('✅ trackLinkClick: ' + linkId + ' → ' + (cur + 1));
  } catch (e) {
    Logger.log('❌ trackLinkClick: ' + e.toString());
  }
}

// ============================================================
// === getLinkViews() ===
// คืนค่า object { linkId: count } ทั้งหมดที่มี prefix "link_view_"
// เรียกจาก index.html → loadLearnViewCounts()
// ============================================================
function getLinkViews() {
  try {
    var props = PropertiesService.getScriptProperties().getProperties();
    var result = {};
    Object.keys(props).forEach(function(key) {
      if (key.indexOf('link_view_') === 0) {
        var linkId = key.replace('link_view_', '');
        result[linkId] = parseInt(props[key] || '0', 10);
      }
    });
    Logger.log('✅ getLinkViews: คืน ' + Object.keys(result).length + ' รายการ');
    return result;
  } catch (e) {
    Logger.log('❌ getLinkViews: ' + e.toString());
    return {};
  }
}
// ============================================================
// === HRD Export Summary Functions
// === เพิ่มใน รหัส.gs (วางต่อท้ายไฟล์)
// === โรงพยาบาลสันทราย · พัฒนาโดย HRD Team 2569
// ============================================================
/**
 * getHrdSummaryDataEnhanced(params)
 * ปรับปรุงจาก getHrdSummaryData เดิม เพื่อตรวจสอบไฟล์แนบอย่างละเอียด
