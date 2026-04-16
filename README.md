# HRD Web App — โรงพยาบาลสันทราย

ระบบบันทึกหน่วยกิตและติดตามผลการพัฒนาบุคลากร (HRD) งบประมาณ 2569

## 📁 โครงสร้างไฟล์

```
hrd-webapp/
├── appsscript.json              # manifest: OAuth scopes, timezone, webapp config
├── index.html                   # UI หลัก (Single Page App)
├── gas/
│   ├── 00_Config_Auth.gs        # ค่าคงที่, Auth (login, token), Drive setup utils
│   ├── 01_Router_Helpers.gs     # doGet(), include(), getColIdx()
│   ├── 02_Staff_Data.gs         # Staff data: getAllUniqueStaffNames, getEmployeeLineData,
│   │                            #   saveEmployeeLineData, addManualEmployee,
│   │                            #   searchStaffByName, getPersonalData, getFilteredDashboard,
│   │                            #   getSimpleDashboard, getGroupSummary
│   ├── 03_Export_Word.gs        # exportPersonalSelectedToWord, exportDashboardToWord
│   ├── 04_Admin.gs              # adminLogin, verifyAdminToken, ฟังก์ชัน admin อื่นๆ
│   ├── 05_Registration.gs       # PART 1: ระบบลงทะเบียนอบรม
│   ├── 06_LINE_Notify.gs        # PART 2: LINE Messaging API + Trigger setup
│   ├── 07_Training_Drive.gs     # PART 7: saveTrainingRequest, updateTrainingRequest,
│   │                            #   saveFilesToDrive, getOrInitTrainingSheet, etc.
│   ├── 08_Training_Queries.gs   # getAllTrainingData, getRecentTrainingRequests,
│   │                            #   getRecentFromRegistrations, trackLinkClick, getLinkViews
│   └── 09_HRD_Export_Summary.gs # getHrdSummaryDataEnhanced, exportHrdToSheet,
│                                #   exportHrdToSheetEnhanced, checkRegistrationDocuments
└── docs/
    └── (เอกสารเพิ่มเติม)
```

## ⚙️ การตั้งค่าก่อนใช้งาน

1. เปิด [Google Apps Script](https://script.google.com) → สร้าง Project ใหม่
2. อัปโหลดไฟล์ทั้งหมดใน `gas/` เข้า Project (หรือใช้ [clasp](https://github.com/google/clasp))
3. แก้ไขค่าใน `gas/00_Config_Auth.gs`:
   ```javascript
   var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID';
   var ATTACHMENTS_FOLDER_ID = 'YOUR_DRIVE_FOLDER_ID';
   ```
4. ตั้งค่า Script Properties (`Project Settings → Script properties`):
   - `ADMIN_PASSWORD` — รหัสผ่านแอดมิน
   - `STAFF_PASSWORD` — รหัสผ่านเจ้าหน้าที่
   - `LINE_CHANNEL_ACCESS_TOKEN` — LINE Messaging API token
5. รัน `initDriveAccess()` ครั้งเดียวเพื่อ grant DriveApp permission
6. Deploy → New deployment → Web app → Execute as: Me, Access: Anyone

## 🛠️ เครื่องมือที่แนะนำ (clasp)

```bash
npm install -g @google/clasp
clasp login
clasp create --type webapp --title "HRD Sansai"
clasp push
```

## 📋 Sheet ที่ใช้งาน

| Sheet Name | คำอธิบาย |
|---|---|
| `อบรม ปีงบประมาณ 2569` | ข้อมูลการอบรมหลัก |
| `Registrations` | ข้อมูลการลงทะเบียน |
| `Config` | ค่าตั้งค่าระบบ |

## 👥 พัฒนาโดย
HRD Team — โรงพยาบาลสันทราย ปีงบประมาณ 2569
