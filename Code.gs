// ============================================================
// ระบบฐานข้อมูลศิษย์เก่าและฎีกา - Google Apps Script Backend
// ============================================================

const SHEET_ID = '17Q755BxZtGNg5dJr0VBOoKYu_bJtpuJ6syAFVUyJI5M'; 
const DRIVE_FOLDER_ID = '1ZhrROoLGRSdeW-MpxFLfohRr8Z1uaHGO'; 

// รายชื่อแท็บใน Google Sheets
const SHEETS = {
  ADMINS: 'admins',
  STAFF: 'staff',
  ALUMNI: 'alumni',
  BRANCHES: 'branches',
  SLIDES: 'slides',
  NEWS: 'news',
  SETTINGS: 'settings',
  VISITORS: 'visitors',
  DEEKA: 'deeka' // แท็บเก็บข้อมูลฎีกาเก่าที่ปรับปรุงใหม่
};

// ============================================================
// MAIN ENTRY POINT
// ============================================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ระบบฐานข้อมูลศิษย์เก่าและข้อมูลฎีกาเก่า')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doPost(e) {
  const params = e.parameter;
  const callback = params.callback;
  const action = params.action;
  let result;
  try {
    result = processAction(action, params);
  } catch (err) {
    result = { success: false, message: err.message };
  }
  const jsonp = callback + '(' + JSON.stringify(result) + ')';
  return ContentService.createTextOutput(jsonp)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// แหล่งรวม Action และเรียกใช้งานผ่าน Client
function processAction(action, params) {
  recordVisitor();
  switch (action) {
    // ระบบยืนยันตัวตน (Authentication)
    case 'login': return login(params);
    case 'logout': return { success: true };

    // ระบบจัดการแอดมิน
    case 'getAdmins': return getAdmins();
    case 'addAdmin': return addAdmin(params);
    case 'editAdmin': return editAdmin(params);
    case 'deleteAdmin': return deleteAdmin(params);
    case 'resetAdminPassword': return resetAdminPassword(params);

    // ระบบจัดการเจ้าหน้าที่
    case 'getStaff': return getStaff();
    case 'addStaff': return addStaff(params);
    case 'editStaff': return editStaff(params);
    case 'deleteStaff': return deleteStaff(params);
    case 'resetStaffPassword': return resetStaffPassword(params);
    case 'updateStaffProfile': return updateStaffProfile(params);

    // ระบบจัดการสาย/สาขาที่เรียน
    case 'getBranches': return getBranches();
    case 'addBranch': return addBranch(params);
    case 'editBranch': return editBranch(params);
    case 'deleteBranch': return deleteBranch(params);

    // ระบบศิษย์เก่า (Alumni Management)
    case 'getAlumni': return getAlumni(params);
    case 'addAlumni': return addAlumni(params);
    case 'editAlumni': return editAlumni(params);
    case 'deleteAlumni': return deleteAlumni(params);
    case 'approveAlumni': return approveAlumni(params);
    case 'resetAlumniPassword': return resetAlumniPassword(params);
    case 'updateAlumniProfile': return updateAlumniProfile(params);
    case 'searchAlumni': return searchAlumni(params);

    // ปรับปรุงใหม่: ขบวนการลงทะเบียนศิษย์เก่า ร่วมรุ่น + OTP
    case 'sendOTPToPhone': return sendOTPToPhone(params);
    case 'verifyOTPCode': return verifyOTPCode(params);
    case 'finalRegisterAlumni': return finalRegisterAlumni(params);

    // ปรับปรุงใหม่: ค้นหาข้อมูลฎีกาเก่า
    case 'searchDeekaDocuments': return searchDeekaDocuments(params);

    // ระบบสไลด์หน้าแรก
    case 'getSlides': return getSlides();
    case 'addSlide': return addSlide(params);
    case 'editSlide': return editSlide(params);
    case 'deleteSlide': return deleteSlide(params);

    // ระบบข่าวสารคลังโรงเรียน
    case 'getNews': return getNews();
    case 'addNews': return addNews(params);
    case 'editNews': return editNews(params);
    case 'deleteNews': return deleteNews(params);

    // หน้าตั้งค่าระบบทั่วไป
    case 'getSettings': return getSettings();
    case 'saveSettings': return saveSettings(params);

    // รายงานสถิติแดชบอร์ด
    case 'getAlumniByBranch': return getAlumniByBranch();
    case 'getVisitorStats': return getVisitorStats();

    // บันทึกรูปภาพศิษย์เก่าขึ้น Drive
    case 'uploadImage': return uploadImageToDrive(params);

    default: return { success: false, message: 'ไม่พบ Action: ' + action };
  }
}

// ============================================================
// INITIALIZATION - ตรวจสอบและสร้างโครงสร้างชีตระบบทั้งหมด
// ============================================================
function initSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheetDefs = [
    { name: SHEETS.ADMINS, headers: ['id','username','password','name','email','createdAt'] },
    { name: SHEETS.STAFF, headers: ['id','username','password','name','email','phone','createdAt'] },
    { name: SHEETS.ALUMNI, headers: ['id','username','password','firstName','lastName','nickname','graduateYear','generation','branch','birthdate','phone','mobile','email','website','address','photoUrl','photoLink','company','position','workAddress','status','createdAt'] },
    { name: SHEETS.BRANCHES, headers: ['id','name','createdAt'] },
    { name: SHEETS.SLIDES, headers: ['id','title','imageUrl','link','createdAt'] },
    { name: SHEETS.NEWS, headers: ['id','title','content','imageUrl','createdAt','createdBy'] },
    { name: SHEETS.SETTINGS, headers: ['key','value'] },
    { name: SHEETS.VISITORS, headers: ['date','count'] },
    { name: SHEETS.DEEKA, headers: ['deekaId','year','category','detail','amount','payee','date','fileUrl'] } // เพิ่มส่วนชีตฎีกา
  ];
  sheetDefs.forEach(def => {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      sheet.appendRow(def.headers);
    }
  });

  // สร้างบัญชี Admin เริ่มต้นกรณีชีตยังว่างอยู่
  const adminSheet = ss.getSheetByName(SHEETS.ADMINS);
  if (adminSheet.getLastRow() <= 1) {
    adminSheet.appendRow([generateId(), 'admin', hashPassword('admin1234'), 'ผู้ดูแลระบบ', 'admin@school.ac.th', new Date().toISOString()]);
  }

  // ตั้งค่าข้อมูลเว็บไซต์ทั่วไปเริ่มต้น
  const settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
  if (settingsSheet.getLastRow() <= 1) {
    settingsSheet.appendRow(['siteName', 'ระบบฐานข้อมูลศิษย์เก่าโรงเรียน']);
    settingsSheet.appendRow(['logoUrl', '']);
    settingsSheet.appendRow(['footerText', 'Copyright © 2026 ระบบฐานข้อมูลศิษย์เก่าและคลังข้อมูลสารสนเทศ | พัฒนาโดย KKnasawan']);
  }
  return { success: true, message: 'สร้างชีตระบบเริ่มต้นเรียบร้อยแล้ว' };
}

// ============================================================
// UTILITIES FUNCTIONS
// ============================================================
function generateId() {
  return Utilities.getUuid();
}

function hashPassword(pw) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw);
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

function sheetToObjects(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const [headers, ...rows] = sheet.getDataRange().getValues();
  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) return i + 1;
  }
  return -1;
}

function recordVisitor() {
  try {
    const sheet = getSheet(SHEETS.VISITORS);
    const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === today) {
        sheet.getRange(i + 1, 2).setValue(Number(data[i][1]) + 1);
        return;
      }
    }
    sheet.appendRow([today, 1]);
  } catch (e) {}
}

// ============================================================
// AUTHENTICATION LOGIC
// ============================================================
function login(params) {
  const { username, password } = params;
  const hashed = hashPassword(password);

  // ตรวจสอบข้อมูลสิทธิ์ แอดมิน
  const admins = sheetToObjects(getSheet(SHEETS.ADMINS));
  const admin = admins.find(a => a.username === username && a.password === hashed);
  if (admin) return { success: true, role: 'admin', user: { id: admin.id, name: admin.name, username: admin.username, email: admin.email } };

  // ตรวจสอบข้อมูลสิทธิ์ เจ้าหน้าที่
  const staffList = sheetToObjects(getSheet(SHEETS.STAFF));
  const staff = staffList.find(s => s.username === username && s.password === hashed);
  if (staff) return { success: true, role: 'staff', user: { id: staff.id, name: staff.name, username: staff.username, email: staff.email } };

  // ตรวจสอบข้อมูลสิทธิ์ ศิษย์เก่า
  const alumniList = sheetToObjects(getSheet(SHEETS.ALUMNI));
  const alumni = alumniList.find(a => a.username === username && a.password === hashed);
  if (alumni) {
    if (alumni.status !== 'approved') return { success: false, message: 'บัญชีศิษย์เก่าของคุณยังไม่ได้รับการอนุมัติใช้งาน' };
    return { success: true, role: 'alumni', user: { id: alumni.id, name: alumni.firstName + ' ' + alumni.lastName, username: alumni.username, email: alumni.email } };
  }

  return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}

// ============================================================
// ADMIN MANAGEMENT
// ============================================================
function getAdmins() {
  const admins = sheetToObjects(getSheet(SHEETS.ADMINS)).map(a => ({ ...a, password: '***' }));
  return { success: true, data: admins };
}

function addAdmin(params) {
  const { username, password, name, email } = params;
  const sheet = getSheet(SHEETS.ADMINS);
  const existing = sheetToObjects(sheet);
  if (existing.find(a => a.username === username)) return { success: false, message: 'Username นี้มีอยู่แล้วในระบบ' };
  sheet.appendRow([generateId(), username, hashPassword(password), name, email, new Date().toISOString()]);
  return { success: true, message: 'เพิ่มผู้ดูแลระบบสำเร็จ' };
}

function editAdmin(params) {
  const { id, name, email } = params;
  const sheet = getSheet(SHEETS.ADMINS);
  const row = findRowById(sheet, id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูลผู้ใช้' };
  sheet.getRange(row, 4).setValue(name);
  sheet.getRange(row, 5).setValue(email);
  return { success: true, message: 'แก้ไขผู้ดูแลระบบสำเร็จ' };
}

function deleteAdmin(params) {
  const sheet = getSheet(SHEETS.ADMINS);
  const row = findRowById(sheet, params.id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูล' };
  sheet.deleteRow(row);
  return { success: true, message: 'ลบข้อมูลสำเร็จ' };
}

function resetAdminPassword(params) {
  const { id, newPassword } = params;
  const sheet = getSheet(SHEETS.ADMINS);
  const row = findRowById(sheet, id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูล' };
  sheet.getRange(row, 3).setValue(hashPassword(newPassword));
  return { success: true, message: 'รีเซ็ตรหัสผ่านสำเร็จ' };
}

// ============================================================
// STAFF MANAGEMENT
// ============================================================
function getStaff() {
  const data = sheetToObjects(getSheet(SHEETS.STAFF)).map(s => ({ ...s, password: '***' }));
  return { success: true, data };
}

function addStaff(params) {
  const { username, password, name, email, phone } = params;
  const sheet = getSheet(SHEETS.STAFF);
  const existing = sheetToObjects(sheet);
  if (existing.find(s => s.username === username)) return { success: false, message: 'Username นี้มีอยู่แล้วในระบบ' };
  sheet.appendRow([generateId(), username, hashPassword(password), name, email, phone || '', new Date().toISOString()]);
  return { success: true, message: 'เพิ่มเจ้าหน้าที่เรียบร้อย' };
}

function editStaff(params) {
  const { id, name, email, phone } = params;
  const sheet = getSheet(SHEETS.STAFF);
  const row = findRowById(sheet, id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูล' };
  sheet.getRange(row, 4).setValue(name);
  sheet.getRange(row, 5).setValue(email);
  sheet.getRange(row, 6).setValue(phone || '');
  return { success: true, message: 'แก้ไขข้อมูลเจ้าหน้าที่สำเร็จ' };
}

function deleteStaff(params) {
  const sheet = getSheet(SHEETS.STAFF);
  const row = findRowById(sheet, params.id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูล' };
  sheet.deleteRow(row);
  return { success: true, message: 'ลบเจ้าหน้าที่สำเร็จ' };
}

function resetStaffPassword(params) {
  const { id, newPassword } = params;
  const sheet = getSheet(SHEETS.STAFF);
  const row = findRowById(sheet, id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูล' };
  sheet.getRange(row, 3).setValue(hashPassword(newPassword));
  return { success: true, message: 'รีเซ็ตรหัสผ่านสำเร็จ' };
}

function updateStaffProfile(params) {
  const { id, name, email, phone } = params;
  const sheet = getSheet(SHEETS.STAFF);
  const row = findRowById(sheet, id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูล' };
  sheet.getRange(row, 4).setValue(name);
  sheet.getRange(row, 5).setValue(email);
  sheet.getRange(row, 6).setValue(phone || '');
  return { success: true, message: 'บันทึกแก้ไขข้อมูลส่วนตัวสำเร็จ' };
}

// ============================================================
// BRANCH MANAGEMENT
// ============================================================
function getBranches() {
  return { success: true, data: sheetToObjects(getSheet(SHEETS.BRANCHES)) };
}

function addBranch(params) {
  const sheet = getSheet(SHEETS.BRANCHES);
  const existing = sheetToObjects(sheet);
  if (existing.find(b => b.name === params.name)) return { success: false, message: 'สาขาวิชา/แผนการเรียนนี้ได้รับการบันทึกอยู่แล้ว' };
  sheet.appendRow([generateId(), params.name, new Date().toISOString()]);
  return { success: true, message: 'เพิ่มสาขาวิชาสำเร็จ' };
}

function editBranch(params) {
  const { id, name } = params;
  const sheet = getSheet(SHEETS.BRANCHES);
  const row = findRowById(sheet, id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูล' };
  sheet.getRange(row, 2).setValue(name);
  return { success: true, message: 'ปรับปรุงชื่อสาขาสำเร็จ' };
}

function deleteBranch(params) {
  const sheet = getSheet(SHEETS.BRANCHES);
  const row = findRowById(sheet, params.id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูล' };
  sheet.deleteRow(row);
  return { success: true, message: 'ลบสาขาวิชาสำเร็จ' };
}

// ============================================================
// ALUMNI MANAGEMENT
// ============================================================
function getAlumni(params) {
  let data = sheetToObjects(getSheet(SHEETS.ALUMNI)).map(a => ({ ...a, password: '***' }));
  if (params && params.status) data = data.filter(a => a.status === params.status);
  return { success: true, data };
}

function addAlumni(params) {
  const sheet = getSheet(SHEETS.ALUMNI);
  const existing = sheetToObjects(sheet);
  if (existing.find(a => a.username === params.username)) return { success: false, message: 'Username นี้ได้รับการลงทะเบียนในระบบแล้ว' };
  sheet.appendRow([
    generateId(), params.username, hashPassword(params.password || 'alumni1234'),
    params.firstName || '', params.lastName || '', params.nickname || '',
    params.graduateYear || '', params.generation || '', params.branch || '',
    params.birthdate || '', params.phone || '', params.mobile || '',
    params.email || '', params.website || '', params.address || '',
    params.photoUrl || '', params.photoLink || '',
    params.company || '', params.position || '', params.workAddress || '',
    'approved', new Date().toISOString()
  ]);
  return { success: true, message: 'เพิ่มประวัติศิษย์เก่าสำเร็จ' };
}

function editAlumni(params) {
  const sheet = getSheet(SHEETS.ALUMNI);
  const row = findRowById(sheet, params.id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูล' };
  const cols = ['id','username','password','firstName','lastName','nickname','graduateYear','generation','branch','birthdate','phone','mobile','email','website','address','photoUrl','photoLink','company','position','workAddress','status','createdAt'];
  cols.forEach((col, i) => {
    if (col !== 'id' && col !== 'password' && col !== 'status' && col !== 'createdAt' && params[col] !== undefined) {
      sheet.getRange(row, i + 1).setValue(params[col]);
    }
  });
  return { success: true, message: 'แก้ไขข้อมูลสำเร็จ' };
}

function updateAlumniProfile(params) {
  const sheet = getSheet(SHEETS.ALUMNI);
  const row = findRowById(sheet, params.id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูลโปรไฟล์' };
  const fieldsMap = { firstName:4, lastName:5, nickname:6, graduateYear:7, generation:8, branch:9, birthdate:10, phone:11, mobile:12, email:13, website:14, address:15, photoUrl:16, photoLink:17, company:18, position:19, workAddress:20 };
  Object.entries(fieldsMap).forEach(([key, col]) => {
    if (params[key] !== undefined) sheet.getRange(row, col).setValue(params[key]);
  });
  return { success: true, message: 'บันทึกแก้ไขข้อมูลส่วนตัวสำเร็จแล้ว' };
}

function deleteAlumni(params) {
  const sheet = getSheet(SHEETS.ALUMNI);
  const row = findRowById(sheet, params.id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูล' };
  sheet.deleteRow(row);
  return { success: true, message: 'ลบรายชื่อสำเร็จ' };
}

function approveAlumni(params) {
  const sheet = getSheet(SHEETS.ALUMNI);
  const row = findRowById(sheet, params.id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูล' };
  sheet.getRange(row, 21).setValue('approved');
  return { success: true, message: 'อนุมัติการสมัครศิษย์เก่าสำเร็จ' };
}

function resetAlumniPassword(params) {
  const { id, newPassword } = params;
  const sheet = getSheet(SHEETS.ALUMNI);
  const row = findRowById(sheet, id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูล' };
  sheet.getRange(row, 3).setValue(hashPassword(newPassword));
  return { success: true, message: 'รีเซ็ตรหัสผ่านศิษย์เก่าสำเร็จ' };
}

function searchAlumni(params) {
  const q = (params.q || '').toLowerCase();
  let data = sheetToObjects(getSheet(SHEETS.ALUMNI))
    .filter(a => a.status === 'approved')
    .filter(a => (a.firstName + ' ' + a.lastName + ' ' + a.branch).toLowerCase().includes(q))
    .map(a => ({ ...a, password: '***' }));
  return { success: true, data };
}

// ============================================================
// OTP & REGISTRATION FLOW (ปรับปรุงใหม่ตามที่กำหนด)
// ============================================================

// ส่งรหัส OTP 4 หลักทางข้อความจำลอง / แคช
function sendOTPToPhone(params) {
  const { mobile } = params;
  const sheet = getSheet(SHEETS.ALUMNI);
  const alumniList = sheetToObjects(sheet);
  
  // ตรวจสอบเบอร์ซ้ำในชีตก่อนส่ง
  const existing = alumniList.find(a => a.mobile === mobile);
  if (existing) {
    return { success: false, message: 'เบอร์โทรศัพท์มือถือนี้เคยถูกใช้สมัครไปแล้ว' };
  }

  // สุ่มเลข 4 หลัก
  const otpCode = Math.floor(1000 + Math.random() * 9000).toString();
  
  // เก็บไว้ในแคชระบบมีอายุการกรอก 5 นาที (300 วินาที)
  const cache = CacheService.getScriptCache();
  cache.put('OTP_' + mobile, otpCode, 300);

  // LOG รหัส OTP สำหรับใช้ในการ Debug ใช้งาน (หากใช้ SMS API จริง สามารถเพิ่มโค้ดยิง GET/POST ตรงนี้ได้)
  Logger.log(`[OTP Verification] เบอร์มือถือ: ${mobile} => รหัส OTP คือ: ${otpCode}`);

  return {
    success: true,
    message: 'จำลองการส่งรหัส OTP เรียบร้อยแล้ว (มีอายุการยืนยันตัวตน 5 นาที)',
    debugOtp: otpCode // ส่งค่าออกไปให้ทดสอบกรอกได้ทันทีแบบไม่มี SMS Gateway
  };
}

// ตรวจรหัส OTP
function verifyOTPCode(params) {
  const { mobile, otp } = params;
  const cache = CacheService.getScriptCache();
  const savedOtp = cache.get('OTP_' + mobile);

  if (!savedOtp) {
    return { success: false, message: 'รหัส OTP หมดอายุความปลอดภัยแล้ว กรุณาส่งใหม่อีกครั้ง' };
  }
  if (savedOtp === otp) {
    return { success: true, message: 'ตรวจสอบรหัส OTP ถูกต้องสมบูรณ์' };
  } else {
    return { success: false, message: 'รหัส OTP ไม่ถูกต้อง กรุณากรอกใหม่อีกครั้ง' };
  }
}

// บันทึกข้อมูลศิษย์เก่ารอบสุดท้ายหลังจากผ่านยืนยัน OTP
function finalRegisterAlumni(params) {
  const sheet = getSheet(SHEETS.ALUMNI);
  const existing = sheetToObjects(sheet);
  
  if (existing.find(a => a.username === params.username)) {
    return { success: false, message: 'ชื่อผู้ใช้ (Username) นี้ได้รับการใช้งานไปแล้วในระบบ' };
  }

  const newId = generateId();
  sheet.appendRow([
    newId,
    params.username,
    hashPassword(params.password),
    params.firstName || '',
    params.lastName || '',
    '', // nickname ปล่อยว่างให้ศิษย์เก่าไปอัปเดตในหน้าโปรไฟล์
    params.graduateYear || '',
    params.generation || '',
    '', // branch
    params.birthdate || '',
    '', // phone
    params.mobile || '',
    '', '', '', '', '', '', '', '', // ข้อมูลบริษัท/ตำแหน่ง/รูปภาพ ปล่อยว่างเพื่อเติมภายหลัง
    'approved', // อนุมัติให้อัตโนมัติหลังผ่าน OTP
    new Date().toISOString()
  ]);

  return {
    success: true,
    message: 'บันทึกข้อมูลและส่งข้อมูลสำเร็จแล้ว!',
    data: { username: params.username, password: params.password }
  };
}

// ============================================================
// DEEKA OLD SYSTEM (ระบบคลังสืบค้นฎีกาเก่า - ปรับปรุงใหม่)
// ============================================================
function searchDeekaDocuments(params) {
  try {
    const sheet = getSheet(SHEETS.DEEKA);
    if (!sheet) return { success: true, data: [] };

    const data = sheetToObjects(sheet);
    const q = (params.q || '').toLowerCase().trim();
    const year = (params.year || '').trim();

    const filtered = data.filter(doc => {
      const matchYear = year ? String(doc.year) === year : true;
      const matchText = q ? (
        String(doc.deekaId).toLowerCase().includes(q) ||
        String(doc.detail).toLowerCase().includes(q) ||
        String(doc.payee).toLowerCase().includes(q) ||
        String(doc.category).toLowerCase().includes(q)
      ) : true;
      return matchYear && matchText;
    });

    return { success: true, data: filtered };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการดึงข้อมูลฎีกา: ' + err.message };
  }
}

// ============================================================
// SLIDES & NEWS & SETTINGS
// ============================================================
function getSlides() {
  return { success: true, data: sheetToObjects(getSheet(SHEETS.SLIDES)) };
}

function addSlide(params) {
  getSheet(SHEETS.SLIDES).appendRow([generateId(), params.title || '', params.imageUrl || '', params.link || '', new Date().toISOString()]);
  return { success: true, message: 'เพิ่มสไลด์สำเร็จ' };
}

function editSlide(params) {
  const sheet = getSheet(SHEETS.SLIDES);
  const row = findRowById(sheet, params.id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูล' };
  sheet.getRange(row, 2).setValue(params.title || '');
  sheet.getRange(row, 3).setValue(params.imageUrl || '');
  sheet.getRange(row, 4).setValue(params.link || '');
  return { success: true, message: 'แก้ไขสำเร็จ' };
}

function deleteSlide(params) {
  const sheet = getSheet(SHEETS.SLIDES);
  const row = findRowById(sheet, params.id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูล' };
  sheet.deleteRow(row);
  return { success: true, message: 'ลบสำเร็จ' };
}

function getNews() {
  return { success: true, data: sheetToObjects(getSheet(SHEETS.NEWS)).reverse() };
}

function addNews(params) {
  getSheet(SHEETS.NEWS).appendRow([generateId(), params.title || '', params.content || '', params.imageUrl || '', new Date().toISOString(), params.createdBy || '']);
  return { success: true, message: 'เพิ่มข่าวสารสารสนเทศสำเร็จ' };
}

function editNews(params) {
  const sheet = getSheet(SHEETS.NEWS);
  const row = findRowById(sheet, params.id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูล' };
  sheet.getRange(row, 2).setValue(params.title || '');
  sheet.getRange(row, 3).setValue(params.content || '');
  sheet.getRange(row, 4).setValue(params.imageUrl || '');
  return { success: true, message: 'แก้ไขสำเร็จ' };
}

function deleteNews(params) {
  const sheet = getSheet(SHEETS.NEWS);
  const row = findRowById(sheet, params.id);
  if (row < 0) return { success: false, message: 'ไม่พบข้อมูล' };
  sheet.deleteRow(row);
  return { success: true, message: 'ลบสำเร็จ' };
}

function getSettings() {
  const sheet = getSheet(SHEETS.SETTINGS);
  const data = sheetToObjects(sheet);
  const settings = {};
  data.forEach(row => { settings[row.key] = row.value; });
  return { success: true, data: settings };
}

function saveSettings(params) {
  const sheet = getSheet(SHEETS.SETTINGS);
  const data = sheet.getDataRange().getValues();
  const keys = ['siteName', 'logoUrl', 'footerText'];
  keys.forEach(key => {
    if (params[key] !== undefined) {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(params[key]);
          found = true;
          break;
        }
      }
      if (!found) sheet.appendRow([key, params[key]]);
    }
  });
  return { success: true, message: 'ปรับปรุงข้อมูลทั่วไปสำเร็จ' };
}

// ============================================================
// REPORTS & GRAPHS
// ============================================================
function getAlumniByBranch() {
  const alumni = sheetToObjects(getSheet(SHEETS.ALUMNI)).filter(a => a.status === 'approved');
  const branches = sheetToObjects(getSheet(SHEETS.BRANCHES));
  const result = branches.map(b => ({
    branch: b.name,
    count: alumni.filter(a => a.branch === b.name).length
  }));
  return { success: true, data: result };
}

function getVisitorStats() {
  const data = sheetToObjects(getSheet(SHEETS.VISITORS));
  return { success: true, data };
}

// ============================================================
// GOOGLE DRIVE IMAGE UPLOAD
// ============================================================
function uploadImageToDrive(params) {
  try {
    const { base64Data, mimeType, fileName, alumniName } = params;
    const rootFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

    let subFolder;
    const folders = rootFolder.getFoldersByName(alumniName);
    if (folders.hasNext()) {
      subFolder = folders.next();
    } else {
      subFolder = rootFolder.createFolder(alumniName);
    }

    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    const file = subFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId = file.getId();
    // ปรับโครงสร้าง URL ให้เสถียรสำหรับการเป็นหน้ากากแสดงผลของ <img> แท็ก
    const directUrl = 'https://docs.google.com/uc?export=view&id=' + fileId;
    return { success: true, url: directUrl, fileId };
  } catch (e) {
    return { success: false, message: e.message };
  }
}