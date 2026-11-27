
// ================= ตั้งค่าระบบ =================
const SHEET_ID = '1UUqHcRHOOuzVBJqrk8VCLJ7J81LzgEkJ2dCfLQrTJfA';
const SHEET_NAME = 'Applicants';
const ADMIN_USER = 'admin';
const ADMIN_PASS = 'admin5678';

// ใส่ ID โฟลเดอร์ Google Drive สำหรับเก็บรูปและเอกสาร (ปล่อยว่างให้ระบบสร้างอัตโนมัติ)
const DRIVE_FOLDER_ID = '';

// ใส่ File ID ของไฟล์โลโก้โรงเรียนใน Google Drive (หรืออัปโหลดผ่านปุ่มในหน้าแอดมินได้ทันที)
const SCHOOL_LOGO_FILE_ID = '';
const DEFAULT_LOGO_FILENAME = 'apple-icon.png';
// =============================================

const HEADERS = [
  'รหัสใบสมัคร', 'วันเวลาที่สมัคร', 'วันที่(สถิติ)', 'ระดับชั้น', 'แผนการเรียน', 'รอบ/ความสามารถพิเศษ',
  'เลขบัตรประชาชน', 'คำนำหน้า', 'ชื่อ', 'นามสกุล', 'โรงเรียนที่จบ', 'เบอร์โทรผู้สมัคร',
  'ชื่อ-สกุล บิดา', 'อาชีพบิดา', 'เบอร์โทรบิดา',
  'ชื่อ-สกุล มารดา', 'อาชีพมารดา', 'เบอร์โทรมารดา',
  'ชื่อ-สกุล ผู้ปกครอง', 'ความสัมพันธ์', 'อาชีพผู้ปกครอง', 'เบอร์โทรผู้ปกครอง',
  'ลิงก์รูปถ่าย', 'ลิงก์ ปพ.', 'ลิงก์สำเนาบัตร', 'ลิงก์สำเนาทะเบียนบ้าน',
  'สถานะเอกสาร', 'หมายเหตุเจ้าหน้าที่', 'วันเวลาที่ตรวจสอบ'
];

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  template.schoolLogoDataUri = getSchoolLogoDataUri_();
  template.webAppUrl = ScriptApp.getService().getUrl() || '';
  template.initialAppId = (e && e.parameter && e.parameter.id) ? String(e.parameter.id).trim() : '';

  return template.evaluate()
    .setTitle('ระบบรับสมัครนักเรียน ม.1 และ ม.4 | โรงเรียนท่ายางวิทยา')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSchoolLogoDataUri_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const savedId = props.getProperty('SCHOOL_LOGO_FILE_ID') || SCHOOL_LOGO_FILE_ID;
    let file = null;

    if (savedId) {
      try {
        file = DriveApp.getFileById(savedId);
      } catch (e) {
        file = null;
      }
    }

    if (!file) {
      const files = DriveApp.getFilesByName(DEFAULT_LOGO_FILENAME);
      if (files.hasNext()) {
        file = files.next();
      }
    }

    if (!file) return '';
    const blob = file.getBlob();
    const contentType = blob.getContentType() || 'image/png';
    return 'data:' + contentType + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (err) {
    return '';
  }
}

function getDriveImageDataUri_(url) {
  if (!url) return '';
  try {
    let fileId = '';
    const m1 = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    const m2 = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (m1) fileId = m1[1];
    else if (m2) fileId = m2[1];
    if (!fileId) return url;

    const blob = DriveApp.getFileById(fileId).getBlob();
    const contentType = blob.getContentType() || 'image/jpeg';
    return 'data:' + contentType + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    return url;
  }
}

function getSheet_() {

  //const data = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  //const ss = SpreadsheetApp.getActiveSpreadsheet();
 // let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function getUploadFolder_() {
  if (DRIVE_FOLDER_ID) {
    return DriveApp.getFolderById(DRIVE_FOLDER_ID);
  }
  const folderName = 'TYW_Admission_Files';
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  const newFolder = DriveApp.createFolder(folderName);
  newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return newFolder;
}

function saveBase64File_(fileObj, prefixName, isImageThumbnail) {
  if (!fileObj || !fileObj.data) return '';
  const folder = getUploadFolder_();
  const base64Data = fileObj.data.split(',')[1];
  const contentType = fileObj.type || 'application/octet-stream';
  const decoded = Utilities.base64Decode(base64Data);
  const safeName = prefixName + '_' + fileObj.name;
  const blob = Utilities.newBlob(decoded, contentType, safeName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  if (isImageThumbnail) {
    return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w600';
  }
  return file.getUrl();
}

function rowToObject_(row, rowNumber) {
  return {
    rowNumber: rowNumber,
    id: String(row[0] || ''),
    timestamp: String(row[1] || ''),
    dateKey: String(row[2] || ''),
    level: String(row[3] || ''),
    plan: String(row[4] || ''),
    subPlan: String(row[5] || ''),
    nationalId: String(row[6] || '').replace(/^'/, ''),
    prefix: String(row[7] || ''),
    firstName: String(row[8] || ''),
    lastName: String(row[9] || ''),
    prevSchool: String(row[10] || ''),
    phone: String(row[11] || '').replace(/^'/, ''),
    fatherName: String(row[12] || ''),
    fatherJob: String(row[13] || ''),
    fatherPhone: String(row[14] || '').replace(/^'/, ''),
    motherName: String(row[15] || ''),
    motherJob: String(row[16] || ''),
    motherPhone: String(row[17] || '').replace(/^'/, ''),
    guardianName: String(row[18] || ''),
    guardianRelation: String(row[19] || ''),
    guardianJob: String(row[20] || ''),
    guardianPhone: String(row[21] || '').replace(/^'/, ''),
    photoUrl: String(row[22] || ''),
    docPorPorUrl: String(row[23] || ''),
    docIdCardUrl: String(row[24] || ''),
    docHouseRegUrl: String(row[25] || ''),
    docStatus: String(row[26] || 'รอตรวจสอบ'),
    adminNote: String(row[27] || ''),
    verifiedAt: String(row[28] || '')
  };
}

// 1. บันทึกใบสมัครใหม่
function submitApplication(formData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!formData.firstName || !formData.lastName || !formData.prevSchool || !formData.phone) {
      throw new Error('กรุณากรอก ชื่อ, นามสกุล, โรงเรียนที่จบ และเบอร์โทรศัพท์');
    }

    const sheet = getSheet_();
    const now = new Date();
    const thaiYear = now.getFullYear() + 543;
    const timestamp = Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/') + thaiYear + ' ' +
      Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm');
    const dateKey = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd');

    const lastRow = sheet.getLastRow();
    const seq = String(lastRow).padStart(4, '0');
    const appId = 'TYW' + String(thaiYear).slice(-2) + '-' + seq;
    const filePrefix = appId + '_' + formData.firstName;

    const photoUrl = saveBase64File_(formData.photoFile, filePrefix + '_PHOTO', true);
    const docPorPorUrl = saveBase64File_(formData.docPorPorFile, filePrefix + '_PORPOR', false);
    const docIdCardUrl = saveBase64File_(formData.docIdCardFile, filePrefix + '_IDCARD', false);
    const docHouseRegUrl = saveBase64File_(formData.docHouseRegFile, filePrefix + '_HOUSEREG', false);

    const newRow = [
      appId,
      timestamp,
      dateKey,
      formData.level || 'ม.1',
      formData.plan || '',
      formData.subPlan || '',
      formData.nationalId ? "'" + formData.nationalId : '',
      formData.prefix || '',
      formData.firstName.trim(),
      formData.lastName.trim(),
      formData.prevSchool.trim(),
      "'" + formData.phone.trim(),
      formData.fatherName || '',
      formData.fatherJob || '',
      formData.fatherPhone ? "'" + formData.fatherPhone : '',
      formData.motherName || '',
      formData.motherJob || '',
      formData.motherPhone ? "'" + formData.motherPhone : '',
      formData.guardianName || '',
      formData.guardianRelation || '',
      formData.guardianJob || '',
      formData.guardianPhone ? "'" + formData.guardianPhone : '',
      photoUrl,
      docPorPorUrl,
      docIdCardUrl,
      docHouseRegUrl,
      'รอตรวจสอบ',
      '',
      ''
    ];

    sheet.appendRow(newRow);

    const record = rowToObject_(newRow, sheet.getLastRow());
    record.photoDataUri = formData.photoFile && formData.photoFile.data ? formData.photoFile.data : '';

    return {
      success: true,
      appId: appId,
      record: record
    };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// 2. ข้อมูลสถิติสำหรับหน้าแดชบอร์ด
function getDashboardStats() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  const emptyResult = {
    total: 0, m1Total: 0, m4Total: 0, verifiedCount: 0,
    daily: {}, m1Plans: {}, m4Plans: {}
  };
  if (lastRow <= 1) return emptyResult;

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const stats = {
    total: data.length,
    m1Total: 0,
    m4Total: 0,
    verifiedCount: 0,
    daily: {},
    m1Plans: {},
    m4Plans: {}
  };

  data.forEach(row => {
    const dateKey = String(row[2] || 'ไม่ระบุวันที่');
    const level = String(row[3] || '');
    const plan = String(row[4] || 'ไม่ระบุแผน');
    const subPlan = String(row[5] || '');
    const status = String(row[26] || '');
    const fullPlan = subPlan ? plan + ' (' + subPlan + ')' : plan;

    if (!stats.daily[dateKey]) {
      stats.daily[dateKey] = { m1: 0, m4: 0, total: 0 };
    }
    stats.daily[dateKey].total++;

    if (level === 'ม.1') {
      stats.m1Total++;
      stats.daily[dateKey].m1++;
      stats.m1Plans[fullPlan] = (stats.m1Plans[fullPlan] || 0) + 1;
    } else if (level === 'ม.4') {
      stats.m4Total++;
      stats.daily[dateKey].m4++;
      stats.m4Plans[fullPlan] = (stats.m4Plans[fullPlan] || 0) + 1;
    }

    if (status === 'อนุมัติตรวจสอบแล้ว') {
      stats.verifiedCount++;
    }
  });

  return stats;
}

// 3. ค้นหาใบสมัครเพื่อสั่งพิมพ์/ดาวน์โหลด
function searchApplicantsForPrint(criteria) {
  if (!criteria) return [];

  const mode = criteria.mode || 'name';
  const fn = String(criteria.firstName || '').trim().toLowerCase();
  const ln = String(criteria.lastName || '').trim().toLowerCase();
  const kw = String(criteria.keyword || '').trim().toLowerCase();

  if (mode === 'name') {
    if (!fn || !ln) return [];
  } else {
    if (kw.length < 3) return [];
  }

  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const results = [];

  for (let i = data.length - 1; i >= 0; i--) {
    const obj = rowToObject_(data[i], i + 2);
    let matched = false;

    if (mode === 'name') {
      matched = obj.firstName.toLowerCase().includes(fn) && obj.lastName.toLowerCase().includes(ln);
    } else {
      matched =
        obj.nationalId.toLowerCase().includes(kw) ||
        obj.prevSchool.toLowerCase().includes(kw) ||
        obj.id.toLowerCase().includes(kw);
    }

    if (matched) {
      obj.photoDataUri = obj.photoUrl ? getDriveImageDataUri_(obj.photoUrl) : '';
      results.push(obj);
      if (results.length >= 10) break;
    }
  }
  return results;
}

// ดึงข้อมูลใบสมัคร 1 รายการตามรหัสใบสมัคร (ใช้สำหรับปุ่มพิมพ์ในหน้าแอดมิน และสแกน QR Code)
function getApplicantByIdForPrint(appId) {
  const id = String(appId || '').trim();
  if (!id) return null;

  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === id.toLowerCase()) {
      const obj = rowToObject_(data[i], i + 2);
      obj.photoDataUri = obj.photoUrl ? getDriveImageDataUri_(obj.photoUrl) : '';
      return obj;
    }
  }
  return null;
}

// 4. ระบบยืนยันตัวตนแอดมิน
function adminLogin(username, password) {
  if (username === ADMIN_USER && password === ADMIN_PASS) {
    return { success: true };
  }
  return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}

function verifyAdmin_(auth) {
  if (!auth || auth.username !== ADMIN_USER || auth.password !== ADMIN_PASS) {
    throw new Error('ไม่มีสิทธิ์เข้าถึงระบบแอดมิน');
  }
}

// 5. แอดมินอัปโหลด/เปลี่ยนโลโก้โรงเรียน
function uploadSchoolLogoAdmin(auth, fileObj) {
  verifyAdmin_(auth);
  if (!fileObj || !fileObj.data) {
    return { success: false, message: 'ไม่พบไฟล์รูปภาพที่อัปโหลด' };
  }
  try {
    const folder = getUploadFolder_();
    const base64Data = fileObj.data.split(',')[1];
    const contentType = fileObj.type || 'image/png';
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, 'SCHOOL_LOGO_' + fileObj.name);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    PropertiesService.getScriptProperties().setProperty('SCHOOL_LOGO_FILE_ID', file.getId());
    return { success: true, logoDataUri: fileObj.data };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// 6. ดึงรายชื่อผู้สมัครทั้งหมด (สำหรับแอดมิน)
function getAllApplicantsAdmin(auth) {
  verifyAdmin_(auth);
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return data.map((row, idx) => rowToObject_(row, idx + 2)).reverse();
}

// 7. แอดมินแก้ไขข้อมูลผู้สมัครครบทุกช่อง (รวมอัปโหลดไฟล์แทนที่)
function updateApplicantAdmin(auth, appData) {
  verifyAdmin_(auth);
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const idx = ids.indexOf(appData.id);
    if (idx === -1) throw new Error('ไม่พบข้อมูลใบสมัครนี้');

    const rowNum = idx + 2;
    const existing = sheet.getRange(rowNum, 1, 1, HEADERS.length).getValues()[0];
    const filePrefix = appData.id + '_' + (appData.firstName || 'EDIT');

    // หากแอดมินอัปโหลดไฟล์ใหม่ให้บันทึกแทนที่ หากไม่ได้เลือกให้คงลิงก์เดิมไว้
    const photoUrl = appData.photoFile
      ? saveBase64File_(appData.photoFile, filePrefix + '_PHOTO', true)
      : existing[22];
    const docPorPorUrl = appData.docPorPorFile
      ? saveBase64File_(appData.docPorPorFile, filePrefix + '_PORPOR', false)
      : existing[23];
    const docIdCardUrl = appData.docIdCardFile
      ? saveBase64File_(appData.docIdCardFile, filePrefix + '_IDCARD', false)
      : existing[24];
    const docHouseRegUrl = appData.docHouseRegFile
      ? saveBase64File_(appData.docHouseRegFile, filePrefix + '_HOUSEREG', false)
      : existing[25];

    const now = new Date();
    const thaiYear = now.getFullYear() + 543;
    const nowStr = Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/') + thaiYear + ' ' +
      Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm');

    const updatedRow = [
      existing[0], existing[1], existing[2],
      appData.level || existing[3],
      appData.plan || '',
      appData.subPlan || '',
      appData.nationalId ? "'" + appData.nationalId : '',
      appData.prefix || '',
      appData.firstName || '',
      appData.lastName || '',
      appData.prevSchool || '',
      appData.phone ? "'" + appData.phone : '',
      appData.fatherName || '',
      appData.fatherJob || '',
      appData.fatherPhone ? "'" + appData.fatherPhone : '',
      appData.motherName || '',
      appData.motherJob || '',
      appData.motherPhone ? "'" + appData.motherPhone : '',
      appData.guardianName || '',
      appData.guardianRelation || '',
      appData.guardianJob || '',
      appData.guardianPhone ? "'" + appData.guardianPhone : '',
      photoUrl,
      docPorPorUrl,
      docIdCardUrl,
      docHouseRegUrl,
      appData.docStatus || existing[26],
      appData.adminNote || '',
      appData.docStatus === 'อนุมัติตรวจสอบแล้ว' ? (existing[28] || nowStr) : ''
    ];

    sheet.getRange(rowNum, 1, 1, HEADERS.length).setValues([updatedRow]);
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// 8. แอดมินอนุมัติ/อัปเดตสถานะเอกสาร
function updateDocStatusAdmin(auth, appId, status, note) {
  verifyAdmin_(auth);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const idx = ids.indexOf(appId);
    if (idx === -1) throw new Error('ไม่พบข้อมูลใบสมัครนี้');

    const rowNum = idx + 2;
    const now = new Date();
    const thaiYear = now.getFullYear() + 543;
    const nowStr = Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/') + thaiYear + ' ' +
      Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm');

    sheet.getRange(rowNum, 27, 1, 3).setValues([[
      status,
      note || '',
      status === 'อนุมัติตรวจสอบแล้ว' ? nowStr : ''
    ]]);
    return { success: true, verifiedAt: status === 'อนุมัติตรวจสอบแล้ว' ? nowStr : '' };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// 9. แอดมินลบข้อมูลผู้สมัคร
function deleteApplicantAdmin(auth, appId) {
  verifyAdmin_(auth);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const idx = ids.indexOf(appId);
    if (idx === -1) throw new Error('ไม่พบข้อมูลใบสมัครนี้');

    sheet.deleteRow(idx + 2);
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}
