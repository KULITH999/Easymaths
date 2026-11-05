// =====================================================
// EASY MATHS — API WRAPPER (for Vercel frontend)
// =====================================================

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ success: false, message: 'Invalid JSON body.' }); }

  const action = body.action;
  const args = body.args || [];
  let result;

  try {
    switch (action) {
      case 'signUpUser':                 result = signUpUser(args[0]); break;
      case 'loginUser':                  result = loginUser(args[0], args[1]); break;
      case 'checkSession':               result = checkSession(args[0]); break;
      case 'logoutUser':                 result = logoutUser(args[0]); break;

      case 'adminLogin':                 result = adminLogin(args[0], args[1]); break;
      case 'checkAdminSession':          result = checkAdminSession(args[0]); break;
      case 'adminLogout':                result = adminLogout(args[0]); break;

      case 'adminGetStudents':           result = adminGetStudents(args[0]); break;
      case 'adminSetPaymentStatus':      result = adminSetPaymentStatus(args[0], args[1], args[2]); break;
      case 'adminSetStudentAccess':      result = adminSetStudentAccess(args[0], args[1], args[2]); break;
      case 'adminDeleteStudent':         result = adminDeleteStudent(args[0], args[1]); break;
      case 'adminChangeStudentPassword': result = adminChangeStudentPassword(args[0], args[1], args[2]); break;
      case 'checkStudentAccess':         result = checkStudentAccess(args[0]); break;

      case 'adminGetContent':            result = adminGetContent(args[0], args[1]); break;
      case 'adminSaveLinkContent':       result = adminSaveLinkContent(args[0], args[1], args[2]); break;
      case 'adminDeleteContent':         result = adminDeleteContent(args[0], args[1], args[2]); break;
      case 'adminUploadMaterial':        result = adminUploadMaterial(args[0], args[1]); break;
      case 'adminUpdateMaterialDetails': result = adminUpdateMaterialDetails(args[0], args[1], args[2], args[3], args[4]); break;
      case 'studentGetContent':          result = studentGetContent(args[0], args[1]); break;
      case 'studentDownloadMaterial':    result = studentDownloadMaterial(args[0], args[1]); break;
      case 'studentDownloadMaterialById':result = studentDownloadMaterialById(args[0], args[1]); break;

      case 'adminGetClassDetails':       result = adminGetClassDetails(args[0]); break;
      case 'adminSaveClassDetails':      result = adminSaveClassDetails(args[0], args[1]); break;
      case 'adminDeleteClassDetails':    result = adminDeleteClassDetails(args[0], args[1]); break;
      case 'studentGetClassDetails':     result = studentGetClassDetails(args[0]); break;

      default:
        result = { success: false, message: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, message: String(err && err.message ? err.message : err) };
  }

  return json_(result);
}

function doGet(e) {
  return json_({ ok: true, message: 'EASY MATHS API is running.' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================================
// YOUR ORIGINAL CODE (unchanged below this line)
// =====================================================

const SPREADSHEET_ID = '18CaVU8sxdJg8e9zdA6N54T7IAUiQKNFa1UJ1Oxv9ri4';
const USERS_SHEET_NAME = 'Users';
const SESSION_TTL_SECONDS = 21600;

function getDatabaseSheet_() {
  let ss;
  if (SPREADSHEET_ID && SPREADSHEET_ID !== 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE') {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  if (!ss) throw new Error('Google Sheet ID is not set.');
  let sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET_NAME);
    sheet.getRange(1, 1, 1, 8).setValues([[
      'User ID','Name','Email','WhatsApp Number','District','Contact Number','Password Hash','Created At'
    ]]);
    sheet.setFrozenRows(1);
  }
  const dbHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  if (dbHeaders.indexOf('Grade') === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Grade');
  }
  sheet.getRange('D:D').setNumberFormat('@');
  sheet.getRange('F:F').setNumberFormat('@');
  return sheet;
}

function signUpUser(data) {
  data = data || {};
  const name = String(data.name || '').trim();
  const email = normalizeEmail_(data.email);
  const whatsapp = normalizePhone_(data.whatsapp);
  const district = String(data.district || '').trim();
  const contact = normalizePhone_(data.contact);
  const password = String(data.password || '');
  const confirmPassword = String(data.confirmPassword || '');
  const grade = String(data.grade || '').trim();

  if (!name || !email || !whatsapp || !district || !contact || !grade || !password || !confirmPassword) {
    return { success: false, message: 'Please fill all fields.' };
  }
  if (!/^Grade (6|7|8|9|10|11)$/.test(grade)) {
    return { success: false, message: 'Please select a valid grade.' };
  }
  if (!isValidEmail_(email)) {
    return { success: false, message: 'Please enter a valid email address.' };
  }
  if (password.length < 6) {
    return { success: false, message: 'Password must contain at least 6 characters.' };
  }
  if (password !== confirmPassword) {
    return { success: false, message: 'Passwords do not match.' };
  }
  if (whatsapp.length < 7 || contact.length < 7) {
    return { success: false, message: 'Please enter valid phone numbers.' };
  }

  const sheet = getDatabaseSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const lastRow = sheet.getLastRow();
    const values = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues()
      : [];
    for (let i = 0; i < values.length; i++) {
      const rowEmail = normalizeEmail_(values[i][2]);
      const rowWhatsapp = normalizePhone_(values[i][3]);
      const rowContact = normalizePhone_(values[i][5]);
      if (rowEmail === email) return { success: false, message: 'This email is already registered.' };
      if (rowWhatsapp === whatsapp) return { success: false, message: 'This WhatsApp number is already registered.' };
      if (rowContact === contact) return { success: false, message: 'This contact number is already registered.' };
    }
    const userId = Utilities.getUuid();
    const salt = Utilities.getUuid();
    const passwordHash = hashPassword_(password, salt);
    const userHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const gradeColumn = userHeaders.indexOf('Grade');
    const newUserRow = new Array(sheet.getLastColumn()).fill('');
    newUserRow[0] = userId;
    newUserRow[1] = name;
    newUserRow[2] = email;
    newUserRow[3] = whatsapp;
    newUserRow[4] = district;
    newUserRow[5] = contact;
    newUserRow[6] = salt + ':' + passwordHash;
    newUserRow[7] = new Date();
    newUserRow[gradeColumn] = grade;
    sheet.appendRow(newUserRow);
    return { success: true, message: 'Account created successfully.' };
  } finally {
    lock.releaseLock();
  }
}

function loginUser(identifier, password) {
  identifier = String(identifier || '').trim();
  password = String(password || '');
  if (!identifier || !password) {
    return { success: false, message: 'Enter your email/WhatsApp number and password.' };
  }
  const sheet = getDatabaseSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'No registered accounts found.' };
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const userHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const gradeColumn = userHeaders.indexOf('Grade');
  const normalizedEmail = normalizeEmail_(identifier);
  const normalizedPhone = normalizePhoneForLogin_(identifier);
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const email = normalizeEmail_(row[2]);
    const whatsapp = normalizePhoneForLogin_(row[3]);
    if (email === normalizedEmail || whatsapp === normalizedPhone) {
      const stored = String(row[6] || '');
      const parts = stored.split(':');
      if (parts.length !== 2) return { success: false, message: 'Account password data is invalid.' };
      const salt = parts[0];
      const storedHash = parts[1];
      const enteredHash = hashPassword_(password, salt);
      if (!constantTimeEqual_(enteredHash, storedHash)) {
        return { success: false, message: 'Incorrect password.' };
      }
      const user = {
        userId: String(row[0]),
        name: String(row[1]),
        email: String(row[2]),
        whatsapp: String(row[3]),
        district: String(row[4]),
        contact: String(row[5]),
        grade: gradeColumn >= 0 ? String(row[gradeColumn] || '') : ''
      };
      const token = Utilities.getUuid() + '-' + Utilities.getUuid();
      CacheService.getScriptCache().put('session_' + token, JSON.stringify(user), SESSION_TTL_SECONDS);
      return { success: true, message: 'Login successful.', token: token, user: user };
    }
  }
  return { success: false, message: 'Email or WhatsApp number not found.' };
}

function checkSession(token) {
  token = String(token || '').trim();
  if (!token) return { success: false };
  const cached = CacheService.getScriptCache().get('session_' + token);
  if (!cached) return { success: false };
  try { return { success: true, user: JSON.parse(cached) }; }
  catch (e) { return { success: false }; }
}

function logoutUser(token) {
  token = String(token || '').trim();
  if (token) CacheService.getScriptCache().remove('session_' + token);
  return { success: true };
}

function normalizeEmail_(value) { return String(value || '').trim().toLowerCase(); }
function normalizePhone_(value) { return String(value || '').replace(/[^0-9+]/g, ''); }
function normalizePhoneForLogin_(value) {
  let phone = String(value || '').replace(/[^0-9]/g, '');
  if (phone.startsWith('0')) phone = phone.substring(1);
  return phone;
}
function isValidEmail_(email) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email); }

function hashPassword_(password, salt) {
  const input = salt + ':' + password;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  return digest.map(function (byte) {
    const value = byte < 0 ? byte + 256 : byte;
    return ('0' + value.toString(16)).slice(-2);
  }).join('');
}

function constantTimeEqual_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  let result = 0;
  for (let i = 0; i < a.length; i++) result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return result === 0;
}

function adminLogin(username, password) {
  const props = PropertiesService.getScriptProperties();
  const savedUsername = props.getProperty('ADMIN_USERNAME');
  const savedPassword = props.getProperty('ADMIN_PASSWORD');
  if (!savedUsername || !savedPassword) return { success: false, message: 'Admin settings not found.' };
  if (String(username) !== savedUsername || String(password) !== savedPassword) {
    return { success: false, message: 'Invalid admin login.' };
  }
  const token = Utilities.getUuid() + '-' + Utilities.getUuid();
  CacheService.getScriptCache().put('admin_' + token, 'AUTHORIZED', SESSION_TTL_SECONDS);
  return { success: true, token: token };
}

function checkAdminSession(token) {
  token = String(token || '').trim();
  if (!token) return { success: false };
  const valid = CacheService.getScriptCache().get('admin_' + token);
  return { success: valid === 'AUTHORIZED' };
}

function adminLogout(token) {
  token = String(token || '').trim();
  if (token) CacheService.getScriptCache().remove('admin_' + token);
  return { success: true };
}

function ensureStudentAccessColumns_() {
  const sheet = getDatabaseSheet_();
  const headers = ['Payment Status', 'Access Status', 'Approved At', 'Approved By'];
  headers.forEach(function (header) {
    const lastCol = sheet.getLastColumn();
    const existing = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    if (existing.indexOf(header) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
    }
  });
  return sheet;
}

function adminGetStudents(adminToken) {
  if (!checkAdminSession(adminToken).success) return { success: false, message: 'Unauthorized admin.' };
  const sheet = ensureStudentAccessColumns_();
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return { success: true, students: [] };
  const headers = values[0];
  const paymentCol = headers.indexOf('Payment Status');
  const accessCol = headers.indexOf('Access Status');
  const students = values.slice(1).map(function (row, index) {
    return {
      rowNumber: index + 2,
      userId: row[0], name: row[1], email: row[2], whatsapp: row[3],
      district: row[4], contact: row[5],
      grade: row[headers.indexOf('Grade')] || '',
      paymentStatus: row[paymentCol] || 'PENDING',
      accessStatus: row[accessCol] || 'LOCKED'
    };
  });
  return { success: true, students: students };
}

function adminSetPaymentStatus(adminToken, rowNumber, status) {
  if (!checkAdminSession(adminToken).success) return { success: false, message: 'Unauthorized admin.' };
  status = String(status || '').toUpperCase();
  if (!['APPROVED', 'PENDING', 'REJECTED'].includes(status)) {
    return { success: false, message: 'Invalid payment status.' };
  }
  const sheet = ensureStudentAccessColumns_();
  rowNumber = Number(rowNumber);
  if (!Number.isInteger(rowNumber) || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    return { success: false, message: 'Invalid student.' };
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const paymentCol = headers.indexOf('Payment Status') + 1;
  const accessCol = headers.indexOf('Access Status') + 1;
  const approvedAtCol = headers.indexOf('Approved At') + 1;
  const approvedByCol = headers.indexOf('Approved By') + 1;
  sheet.getRange(rowNumber, paymentCol).setValue(status);
  if (status === 'APPROVED') {
    sheet.getRange(rowNumber, accessCol).setValue('ACTIVE');
    sheet.getRange(rowNumber, approvedAtCol).setValue(new Date());
    sheet.getRange(rowNumber, approvedByCol).setValue('admin');
  } else {
    sheet.getRange(rowNumber, accessCol).setValue('LOCKED');
    sheet.getRange(rowNumber, approvedAtCol).clearContent();
    sheet.getRange(rowNumber, approvedByCol).clearContent();
  }
  return { success: true, message: 'Student updated successfully.' };
}

function adminSetStudentAccess(adminToken, rowNumber, enabled) {
  if (!checkAdminSession(adminToken).success) return { success: false, message: 'Unauthorized admin.' };
  const sheet = ensureStudentAccessColumns_();
  rowNumber = Number(rowNumber);
  if (!Number.isInteger(rowNumber) || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    return { success: false, message: 'Invalid student.' };
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const paymentCol = headers.indexOf('Payment Status') + 1;
  const accessCol = headers.indexOf('Access Status') + 1;
  if (enabled === true) {
    const paymentStatus = sheet.getRange(rowNumber, paymentCol).getValue();
    if (paymentStatus !== 'APPROVED') {
      return { success: false, message: 'Approve the payment before enabling access.' };
    }
    sheet.getRange(rowNumber, accessCol).setValue('ACTIVE');
  } else {
    sheet.getRange(rowNumber, accessCol).setValue('LOCKED');
  }
  return { success: true, message: 'Access updated successfully.' };
}

function checkStudentAccess(token) {
  token = String(token || '').trim();
  if (!token) return { success: false, allowed: false, status: 'LOGIN_REQUIRED', message: 'Please login first.' };
  const cached = CacheService.getScriptCache().get('session_' + token);
  if (!cached) return { success: false, allowed: false, status: 'LOGIN_REQUIRED', message: 'Your session has expired. Please login again.' };
  let user;
  try { user = JSON.parse(cached); }
  catch (e) { return { success: false, allowed: false, status: 'LOGIN_REQUIRED', message: 'Invalid session. Please login again.' }; }
  if (!user || !user.userId) return { success: false, allowed: false, status: 'LOGIN_REQUIRED', message: 'Invalid user session.' };
  const sheet = ensureStudentAccessColumns_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, allowed: false, status: 'NOT_FOUND', message: 'Student account not found.' };
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const paymentCol = headers.indexOf('Payment Status');
  const accessCol = headers.indexOf('Access Status');
  const gradeCol = headers.indexOf('Grade');
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(user.userId)) {
      const paymentStatus = String(values[i][paymentCol] || 'PENDING').toUpperCase();
      const accessStatus = String(values[i][accessCol] || 'LOCKED').toUpperCase();
      if (paymentStatus !== 'APPROVED') {
        return { success: true, allowed: false, status: paymentStatus, message: 'Your payment is pending approval.' };
      }
      if (accessStatus !== 'ACTIVE') {
        return { success: true, allowed: false, status: 'LOCKED', message: 'Your account access is currently disabled.' };
      }
      user.grade = gradeCol >= 0 ? String(values[i][gradeCol] || '') : '';
      return { success: true, allowed: true, status: 'ACTIVE', message: 'Access granted.', user: user };
    }
  }
  return { success: false, allowed: false, status: 'NOT_FOUND', message: 'Student account not found.' };
}

const LIVE_SHEET_NAME = 'LiveClasses';
const MATERIALS_SHEET_NAME = 'StudyMaterials';
const VIDEOS_SHEET_NAME = 'VideoLessons';

function getContentSheet_(type) {
  const config = {
    live: { name: LIVE_SHEET_NAME, headers: ['ID', 'Title', 'Description', 'Link', 'Created At'] },
    materials: { name: MATERIALS_SHEET_NAME, headers: ['ID', 'Title', 'Description', 'File ID', 'File Name', 'Created At'] },
    videos: { name: VIDEOS_SHEET_NAME, headers: ['ID', 'Title', 'Description', 'Link', 'Created At'] }
  };
  const selected = config[type];
  if (!selected) throw new Error('Invalid content type.');
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(selected.name);
  if (!sheet) {
    sheet = ss.insertSheet(selected.name);
    sheet.appendRow(selected.headers.concat(['Grade']));
    sheet.setFrozenRows(1);
  } else {
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    if (currentHeaders.indexOf('Grade') === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Grade');
    }
  }
  return sheet;
}

function requireAdmin_(adminToken) {
  if (!checkAdminSession(adminToken).success) throw new Error('Unauthorized admin.');
}

function validateClassLink_(link) {
  link = String(link || '').trim();
  if (!link) throw new Error('Please enter a link.');
  const match = link.match(/^https:\/\/([^\/?#]+)(?:[\/?#]|$)/i);
  if (!match) throw new Error('Invalid link. Please enter a full HTTPS Zoom or YouTube link.');
  let authority = match[1].toLowerCase();
  authority = authority.substring(authority.lastIndexOf('@') + 1);
  const host = authority.replace(/:\d+$/, '');
  const allowed =
    host === 'zoom.us' || host.endsWith('.zoom.us') ||
    host === 'youtube.com' || host.endsWith('.youtube.com') ||
    host === 'youtu.be';
  if (!allowed) throw new Error('Only Zoom and YouTube links are allowed.');
  return link;
}

function adminGetContent(adminToken, type) {
  requireAdmin_(adminToken);
  const sheet = getContentSheet_(type);
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0] || [];
  const gradeCol = headers.indexOf('Grade');
  return {
    success: true,
    items: values.length > 1
      ? values.slice(1).map(function (row) {
          return {
            id: row[0], title: row[1], description: row[2],
            link: type === 'materials' ? '' : (row[3] || ''),
            fileName: type === 'materials' ? (row[4] || '') : '',
            grade: gradeCol >= 0 ? (row[gradeCol] || 'ALL') : 'ALL'
          };
        })
      : []
  };
}

function adminSaveLinkContent(adminToken, type, data) {
  requireAdmin_(adminToken);
  if (type !== 'live' && type !== 'videos') throw new Error('Invalid content type.');
  data = data || {};
  const title = String(data.title || '').trim();
  const description = String(data.description || '').trim();
  const link = validateClassLink_(data.link);
  const grade = String(data.grade || '').trim();
  if (!/^Grade (6|7|8|9|10|11|12|13)$/.test(grade) && grade !== 'ALL') {
    throw new Error('Please select a valid grade.');
  }
  if (!title) throw new Error('Please enter a title.');
  const sheet = getContentSheet_(type);
  const id = String(data.id || '').trim();
  if (id) {
    const values = sheet.getDataRange().getDisplayValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        sheet.getRange(i + 1, 2, 1, 3).setValues([[title, description, link]]);
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
        sheet.getRange(i + 1, headers.indexOf('Grade') + 1).setValue(grade);
        return { success: true, message: 'Content updated successfully.' };
      }
    }
    throw new Error('Content not found.');
  }
  sheet.appendRow([Utilities.getUuid(), title, description, link, new Date(), grade]);
  return { success: true, message: 'Content added successfully.' };
}

function adminDeleteContent(adminToken, type, id) {
  requireAdmin_(adminToken);
  const sheet = getContentSheet_(type);
  id = String(id || '').trim();
  if (!id) throw new Error('Invalid content ID.');
  const values = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      if (type === 'materials') {
        const fileId = values[i][3];
        if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {} }
      }
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Content deleted successfully.' };
    }
  }
  throw new Error('Content not found.');
}

function getMaterialsFolder_() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty('EASY_MATHS_MATERIALS_FOLDER');
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (e) {}
  }
  const folder = DriveApp.createFolder('EASY MATHS Study Materials');
  props.setProperty('EASY_MATHS_MATERIALS_FOLDER', folder.getId());
  return folder;
}

function adminUploadMaterial(adminToken, data) {
  requireAdmin_(adminToken);
  data = data || {};
  const title = String(data.title || '').trim();
  const description = String(data.description || '').trim();
  const fileName = String(data.fileName || 'material.pdf').trim();
  const mimeType = String(data.mimeType || '');
  const base64 = String(data.base64 || '');
  const grade = String(data.grade || '').trim();
  if (!/^Grade (6|7|8|9|10|11|12|13)$/.test(grade) && grade !== 'ALL') {
    throw new Error('Please select a valid grade.');
  }
  if (!title || !base64) throw new Error('Title and PDF file are required.');
  if (mimeType !== 'application/pdf' || !fileName.toLowerCase().endsWith('.pdf')) {
    throw new Error('Only PDF files are allowed.');
  }
  if (base64.length > 14000000) throw new Error('PDF file is too large. Maximum size is 10 MB.');
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > 10 * 1024 * 1024) throw new Error('PDF file is too large. Maximum size is 10 MB.');
  if (bytes.length < 5 || bytes[0] !== 37 || bytes[1] !== 80 || bytes[2] !== 68 || bytes[3] !== 70 || bytes[4] !== 45) {
    throw new Error('Invalid PDF file.');
  }
  const blob = Utilities.newBlob(bytes, 'application/pdf', fileName);
  const folder = getMaterialsFolder_();
  const file = folder.createFile(blob);
  const sheet = getContentSheet_('materials');
  sheet.appendRow([Utilities.getUuid(), title, description, file.getId(), fileName, new Date(), grade]);
  return { success: true, message: 'PDF uploaded successfully.' };
}

function studentGetContent(token, type) {
  const access = checkStudentAccess(token);
  if (!access.success || !access.allowed) {
    return { success: false, allowed: false, status: access.status || 'LOGIN_REQUIRED', message: access.message || 'Access denied.' };
  }
  const sheet = getContentSheet_(type);
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0] || [];
  const gradeCol = headers.indexOf('Grade');
  const studentGrade = String(access.user.grade || '').trim();
  const items = values.length > 1
    ? values.slice(1).filter(function (row) {
        const contentGrade = gradeCol >= 0 ? String(row[gradeCol] || 'ALL').trim() : 'ALL';
        return contentGrade === studentGrade;
      }).map(function (row) {
        if (type === 'materials') {
          return { id: row[0], title: row[1], description: row[2], fileName: row[4],
            grade: gradeCol >= 0 ? (row[gradeCol] || 'ALL') : 'ALL' };
        }
        return { id: row[0], title: row[1], description: row[2], link: row[3],
          grade: gradeCol >= 0 ? (row[gradeCol] || 'ALL') : 'ALL' };
      })
    : [];
  return { success: true, allowed: true, items: items };
}

function studentDownloadMaterial(token, fileId) {
  const access = checkStudentAccess(token);
  if (!access.success || !access.allowed) throw new Error(access.message || 'Access denied.');
  fileId = String(fileId || '').trim();
  if (!fileId) throw new Error('Invalid file.');
  const sheet = getContentSheet_('materials');
  const values = sheet.getDataRange().getDisplayValues();
  let found = false;
  const headers = values[0] || [];
  const gradeCol = headers.indexOf('Grade');
  for (let i = 1; i < values.length; i++) {
    if (values[i][3] === fileId) {
      const contentGrade = gradeCol >= 0 ? String(values[i][gradeCol] || 'ALL') : 'ALL';
      if (contentGrade !== 'ALL' && contentGrade !== String(access.user.grade || '')) {
        throw new Error('This material is not assigned to your grade.');
      }
      found = true;
      break;
    }
  }
  if (!found) throw new Error('Material not found.');
  const file = DriveApp.getFileById(fileId);
  if (file.getMimeType() !== 'application/pdf') throw new Error('Invalid material file.');
  const blob = file.getBlob();
  if (blob.getBytes().length > 10 * 1024 * 1024) throw new Error('File is too large to download.');
  return {
    success: true,
    fileName: file.getName(),
    mimeType: 'application/pdf',
    base64: Utilities.base64Encode(blob.getBytes())
  };
}

function adminUpdateMaterialDetails(adminToken, id, title, description, grade) {
  requireAdmin_(adminToken);
  id = String(id || '').trim();
  title = String(title || '').trim();
  description = String(description || '').trim();
  grade = String(grade || '').trim();
  if (!/^Grade (6|7|8|9|10|11|12|13)$/.test(grade) && grade !== 'ALL') {
    throw new Error('Please select a valid grade.');
  }
  if (!id || !title) throw new Error('Material ID and title are required.');
  const sheet = getContentSheet_('materials');
  const values = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[title, description]]);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
      sheet.getRange(i + 1, headers.indexOf('Grade') + 1).setValue(grade);
      return { success: true, message: 'Material details updated successfully.' };
    }
  }
  throw new Error('Material not found.');
}

function studentDownloadMaterialById(token, materialId) {
  const access = checkStudentAccess(token);
  if (!access || !access.success || !access.allowed) {
    return { success: false, message: 'Your account does not have access.' };
  }
  const sheet = getContentSheet_('materials');
  const values = sheet.getDataRange().getValues();
  let fileId = '';
  let fileName = '';
  const headers = values[0] || [];
  const gradeCol = headers.indexOf('Grade');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(materialId)) {
      const contentGrade = gradeCol >= 0 ? String(values[i][gradeCol] || 'ALL') : 'ALL';
      if (contentGrade !== 'ALL' && contentGrade !== String(access.user.grade || '')) {
        return { success: false, message: 'This material is not assigned to your grade.' };
      }
      fileId = values[i][3];
      fileName = values[i][4];
      break;
    }
  }
  if (!fileId) return { success: false, message: 'PDF not found.' };
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  if (blob.getContentType() !== 'application/pdf') return { success: false, message: 'Invalid PDF file.' };
  const bytes = blob.getBytes();
  if (bytes.length > 10 * 1024 * 1024) return { success: false, message: 'File is too large.' };
  return {
    success: true,
    fileName: fileName || file.getName(),
    mimeType: 'application/pdf',
    base64: Utilities.base64Encode(bytes)
  };
}

function adminDeleteStudent(adminToken, userId) {
  if (!checkAdminSession(adminToken).success) return { success: false, message: 'Unauthorized admin.' };
  userId = String(userId || '').trim();
  if (!userId) return { success: false, message: 'Invalid student ID.' };
  const sheet = getDatabaseSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Student not found.' };
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]) === userId) {
        sheet.deleteRow(i + 2);
        return { success: true, message: 'Student deleted successfully.' };
      }
    }
    return { success: false, message: 'Student not found.' };
  } finally {
    lock.releaseLock();
  }
}

function adminChangeStudentPassword(adminToken, userId, newPassword) {
  if (!checkAdminSession(adminToken).success) return { success: false, message: 'Unauthorized admin.' };
  userId = String(userId || '').trim();
  newPassword = String(newPassword || '');
  if (!userId || !newPassword) return { success: false, message: 'Student ID and new password are required.' };
  if (newPassword.length < 6) return { success: false, message: 'Password must contain at least 6 characters.' };
  const sheet = getDatabaseSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Student not found.' };
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]) === userId) {
        const rowNumber = i + 2;
        const newSalt = Utilities.getUuid();
        const newHash = hashPassword_(newPassword, newSalt);
        sheet.getRange(rowNumber, 7).setValue(newSalt + ':' + newHash);
        return { success: true, message: 'Student password changed successfully.' };
      }
    }
    return { success: false, message: 'Student not found.' };
  } finally {
    lock.releaseLock();
  }
}

const CLASS_DETAILS_SHEET_NAME = 'ClassDetails';

function getClassDetailsSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CLASS_DETAILS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CLASS_DETAILS_SHEET_NAME);
    sheet.appendRow(['ID', 'Title', 'Description', 'Link', 'Created At', 'Grade']);
    sheet.setFrozenRows(1);
  } else {
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    if (headers.indexOf('Grade') === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Grade');
    }
  }
  return sheet;
}

function validateClassDetailsLink_(link) {
  link = String(link || '').trim();
  if (!link) return '';
  if (!/^https:\/\/[^\s]+$/i.test(link)) throw new Error('Please enter a valid HTTPS link.');
  return link;
}

function adminGetClassDetails(adminToken) {
  requireAdmin_(adminToken);
  const sheet = getClassDetailsSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0] || [];
  const gradeCol = headers.indexOf('Grade');
  return {
    success: true,
    items: values.length > 1
      ? values.slice(1).map(function (row) {
          return {
            id: row[0] || '', title: row[1] || '', description: row[2] || '',
            link: row[3] || '', grade: gradeCol >= 0 ? (row[gradeCol] || 'ALL') : 'ALL'
          };
        })
      : []
  };
}

function adminSaveClassDetails(adminToken, data) {
  requireAdmin_(adminToken);
  data = data || {};
  const title = String(data.title || '').trim();
  const description = String(data.description || '').trim();
  const link = validateClassDetailsLink_(data.link);
  const grade = String(data.grade || '').trim();
  const id = String(data.id || '').trim();
  if (!title) throw new Error('Please enter a class details title.');
  if (!/^Grade (6|7|8|9|10|11)$/.test(grade)) throw new Error('Please select a grade from Grade 6 to Grade 11.');
  const sheet = getClassDetailsSheet_();
  if (id) {
    const values = sheet.getDataRange().getDisplayValues();
    const headers = values[0] || [];
    const gradeCol = headers.indexOf('Grade');
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === id) {
        sheet.getRange(i + 1, 2, 1, 3).setValues([[title, description, link]]);
        sheet.getRange(i + 1, gradeCol + 1).setValue(grade);
        return { success: true, message: 'Class details updated successfully.' };
      }
    }
    throw new Error('Class details not found.');
  }
  sheet.appendRow([Utilities.getUuid(), title, description, link, new Date(), grade]);
  return { success: true, message: 'Class details added successfully.' };
}

function adminDeleteClassDetails(adminToken, id) {
  requireAdmin_(adminToken);
  id = String(id || '').trim();
  if (!id) throw new Error('Invalid class details ID.');
  const sheet = getClassDetailsSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Class details deleted successfully.' };
    }
  }
  throw new Error('Class details not found.');
}

function studentGetClassDetails(token) {
  token = String(token || '').trim();
  if (!token) return { success: false, allowed: false, status: 'LOGIN_REQUIRED', message: 'Please login first.' };
  const cached = CacheService.getScriptCache().get('session_' + token);
  if (!cached) return { success: false, allowed: false, status: 'LOGIN_REQUIRED', message: 'Your session has expired. Please login again.' };
  let sessionUser;
  try { sessionUser = JSON.parse(cached); }
  catch (e) { return { success: false, allowed: false, status: 'LOGIN_REQUIRED', message: 'Invalid session. Please login again.' }; }
  if (!sessionUser || !sessionUser.userId) return { success: false, allowed: false, status: 'LOGIN_REQUIRED', message: 'Invalid user session.' };
  const usersSheet = getDatabaseSheet_();
  const userValues = usersSheet.getDataRange().getDisplayValues();
  const userHeaders = userValues[0] || [];
  const gradeCol = userHeaders.indexOf('Grade');
  let studentGrade = '';
  for (let i = 1; i < userValues.length; i++) {
    if (String(userValues[i][0]) === String(sessionUser.userId)) {
      studentGrade = gradeCol >= 0 ? String(userValues[i][gradeCol] || '').trim() : '';
      break;
    }
  }
  if (!studentGrade) return { success: false, allowed: false, status: 'STUDENT_NOT_FOUND', message: 'Student grade could not be found.' };
  const sheet = getClassDetailsSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0] || [];
  const contentGradeCol = headers.indexOf('Grade');
  const items = values.length > 1
    ? values.slice(1).filter(function (row) {
        const contentGrade = contentGradeCol >= 0 ? String(row[contentGradeCol] || 'ALL').trim() : 'ALL';
        return contentGrade === studentGrade;
      }).map(function (row) {
        return {
          id: row[0] || '', title: row[1] || '', description: row[2] || '',
          link: row[3] || '', grade: contentGradeCol >= 0 ? (row[contentGradeCol] || 'ALL') : 'ALL'
        };
      })
    : [];
  return { success: true, allowed: true, grade: studentGrade, items: items };
}