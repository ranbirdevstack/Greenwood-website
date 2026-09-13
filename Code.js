/**
 * Greenwood Public School — Serverless Web App Backend
 * Bound to Google Sheets Database
 */

const SPREADSHEET_ID = '1UF8xuxirPcGUwm3fsfSQVTghFzuloDhJJqpkjwi9454';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Greenwood Public School — Official Portal & Admin Hub')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// -------------------------------------------------------------
// 1. DATABASE & SHEET ENGINE
// -------------------------------------------------------------

function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (err) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

function getOrCreateSheet(sheetName, headers) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    const sheets = ss.getSheets();
    sheet = sheets.find(s => s.getName().trim().toLowerCase() === sheetName.trim().toLowerCase());
  }
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length) {
      sheet.appendRow(headers);
    }
  }
  return sheet;
}

function sanitizeDateString(dateVal, tz) {
  if (!dateVal) return '';
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
  }
  const s = String(dateVal).trim();
  if (s.includes('T')) return s.split('T')[0];
  return s;
}

function sanitizeTimeString(timeVal, tz) {
  if (!timeVal) return '';
  if (timeVal instanceof Date) {
    return Utilities.formatDate(timeVal, tz, 'hh:mm a');
  }
  let s = String(timeVal).trim();
  if (s.includes('1899') || s.includes('GMT')) {
    const match = s.match(/\b\d{1,2}:\d{2}(?::\d{2})?\b/);
    return match ? match[0] : '';
  }
  return s;
}

// -------------------------------------------------------------
// 2. SETTINGS & BRANDING CONFIGURATION
// -------------------------------------------------------------

function getSchoolSettings() {
  try {
    const sheet = getOrCreateSheet('Settings', ['Key', 'Value', 'Description / Used In']);
    const rows = sheet.getDataRange().getValues();
    
    const settings = {
      SchoolName: 'Greenwood Public School',
      SchoolTagline: 'Affiliated to CBSE • New Delhi',
      AcademicSession: '2026-2027',
      AffiliationNumber: '2130849',
      SchoolCode: '70142',
      CampusArea: '15 Acres',
      TotalStudents: '2,400+',
      countBoardPass: '100%',
      StudentTeacherRatio: '01:20',
      ContactPhone: '+91 542 223344',
      ContactMobile: '+91 98765 43210',
      ContactEmail: 'admissions@greenwoodpublic.edu.in',
      SchoolAddress: 'Knowledge Park, Varanasi, UP - 221005',
      OfficeHours: 'Mon - Sat (8:00 AM - 2:30 PM)',
      AdmissionTickerAlert: 'Admissions open for Session 2026-27 • Term 1 Datesheet Announced',
      admin_username: 'admin',
      admin_password: 'admin123',
      admin_mobile: '8787262194'
    };

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0]) {
        settings[String(rows[i][0]).trim()] = String(rows[i][1] || '').trim();
      }
    }

    return { success: true, data: settings };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function updateSchoolSettings(newSettings) {
  try {
    const sheet = getOrCreateSheet('Settings', ['Key', 'Value', 'Description / Used In']);
    const data = sheet.getDataRange().getValues();
    const keys = Object.keys(newSettings);

    keys.forEach(key => {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === key) {
          sheet.getRange(i + 1, 2).setValue(newSettings[key]);
          found = true;
          break;
        }
      }
      if (!found) {
        sheet.appendRow([key, newSettings[key], 'Web App Config']);
      }
    });

    return { success: true, message: 'Settings updated successfully!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// -------------------------------------------------------------
// 3. ADMIN AUTH & MOBILE-VERIFIED OTP RECOVERY
// -------------------------------------------------------------

function verifyAdminCredentials(username, password) {
  try {
    const sheet = getOrCreateSheet('Settings', ['Key', 'Value']);
    const rows = sheet.getDataRange().getValues();
    
    let storedUser = 'admin';
    let storedPass = 'admin123';

    for (let i = 1; i < rows.length; i++) {
      const key = String(rows[i][0]).trim();
      if (key === 'admin_username') storedUser = String(rows[i][1]).trim();
      if (key === 'admin_password') storedPass = String(rows[i][1]).trim();
    }

    if (username.trim() === storedUser && password.trim() === storedPass) {
      return { success: true, message: 'Authentication successful.' };
    }
    return { success: false, message: 'Invalid username or password.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function generateAdminOtpForMobile(mobileNumber) {
  try {
    const sheet = getOrCreateSheet('Settings', ['Key', 'Value']);
    const rows = sheet.getDataRange().getValues();

    let registeredMobile = '8787262194';

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === 'admin_mobile') {
        registeredMobile = String(rows[i][1]).trim();
      }
    }

    const cleanInput = String(mobileNumber || '').replace(/\D/g, '').slice(-10);
    const cleanStored = registeredMobile.replace(/\D/g, '').slice(-10);

    if (!cleanInput || cleanInput !== cleanStored) {
      return { 
        success: false, 
        message: 'Mobile number not recognized. Please enter the registered administrator number.' 
      };
    }

    const otp = Math.floor(100000 + Math.random() * 900000).toString();
    const expiry = new Date().getTime() + (10 * 60 * 1000); // 10 minutes

    const payload = {
      'admin_mobile': cleanStored,
      'admin_reset_otp': otp,
      'admin_reset_expiry': expiry.toString()
    };

    const data = sheet.getDataRange().getValues();
    Object.keys(payload).forEach(k => {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === k) {
          sheet.getRange(i + 1, 2).setValue(payload[k]);
          found = true;
          break;
        }
      }
      if (!found) sheet.appendRow([k, payload[k]]);
    });

    return { success: true, otp: otp };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function updateAdminPassword(enteredOtp, newPassword) {
  try {
    const sheet = getOrCreateSheet('Settings', ['Key', 'Value']);
    const rows = sheet.getDataRange().getValues();
    
    let savedOtp = '';
    let expiry = 0;

    for (let i = 1; i < rows.length; i++) {
      const key = String(rows[i][0]).trim();
      if (key === 'admin_reset_otp') savedOtp = String(rows[i][1]).trim();
      if (key === 'admin_reset_expiry') expiry = Number(rows[i][1]);
    }

    if (!savedOtp || !expiry) {
      return { success: false, message: 'No active OTP found. Please request a new code.' };
    }

    const now = new Date().getTime();
    if (now > expiry) {
      return { success: false, message: 'OTP has expired. Please request a new code.' };
    }

    if (String(enteredOtp).trim() !== savedOtp) {
      return { success: false, message: 'Incorrect OTP entered. Password was not updated.' };
    }

    const data = sheet.getDataRange().getValues();
    let updated = false;

    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0]).trim();
      if (key === 'admin_password') {
        sheet.getRange(i + 1, 2).setValue(newPassword.trim());
        updated = true;
      }
      if (key === 'admin_reset_otp') {
        sheet.getRange(i + 1, 2).setValue('');
      }
    }

    if (!updated) {
      sheet.appendRow(['admin_password', newPassword.trim()]);
    }

    return { success: true, message: 'Password updated successfully! Please log in with your new password.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// -------------------------------------------------------------
// 4. ADMISSIONS PIPELINE (SEPARATED ADDRESS & STATUS)
// -------------------------------------------------------------

function submitAdmissionForm(formData) {
  try {
    const sheet = getOrCreateSheet('Admissions', [
      'Timestamp', 'StudentName', 'ParentName', 'Grade', 'Phone', 'Email', 'Address', 'Status'
    ]);
    const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
    const dateStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      dateStr,
      String(formData.studentName || '').trim(),
      String(formData.parentName || '').trim(),
      String(formData.grade || '').trim(),
      String(formData.phone || '').trim(),
      String(formData.email || '').trim(),
      String(formData.address || '').trim(),
      'Pending Review'
    ]);
    return { success: true, message: 'Application submitted successfully! Your status is Pending Review.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getAdmissions() {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('Admissions');
    if (!sheet) {
      const sheets = ss.getSheets();
      sheet = sheets.find(s => s.getName().trim().toLowerCase() === 'admissions');
    }
    if (!sheet) return [];

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];
    rows.shift();

    const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
    const result = [];

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0] && !r[1]) continue;

      let dateStr = '';
      if (r[0] instanceof Date) {
        dateStr = Utilities.formatDate(r[0], tz, 'yyyy-MM-dd HH:mm');
      } else {
        dateStr = String(r[0] || '').trim();
      }

      let addressVal = String(r[6] || '').trim();
      let statusVal = r[7] ? String(r[7]).trim() : 'Pending Review';

      // Backward fallback if older single-column schema was used
      if (!r[7] && addressVal && (addressVal.includes('Review') || addressVal.includes('Scheduled') || addressVal.includes('Approved') || addressVal.includes('Enrolled') || addressVal.includes('Rejected'))) {
        statusVal = addressVal;
        addressVal = '';
      }

      result.push({
        rowNumber: i + 2,
        time: dateStr,
        student: String(r[1] || ''),
        parent: String(r[2] || ''),
        grade: String(r[3] || ''),
        phone: String(r[4] || ''),
        email: String(r[5] || ''),
        address: addressVal,
        status: statusVal
      });
    }

    return result.reverse();
  } catch (err) {
    return [];
  }
}

function updateAdmissionStatus(rowNumber, newStatus) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('Admissions');
    if (!sheet) throw new Error('Admissions sheet not found.');

    const row = Number(rowNumber);
    if (!row || row < 2) throw new Error('Invalid row index.');

    sheet.getRange(1, 7).setValue('Address');
    sheet.getRange(1, 8).setValue('Status');
    sheet.getRange(row, 8).setValue(newStatus);

    return { success: true, message: 'Status updated to "' + newStatus + '"!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// -------------------------------------------------------------
// 5. CALENDAR EVENTS ENGINE
// -------------------------------------------------------------

function getEvents() {
  try {
    const sheet = getOrCreateSheet('Events', ['Title', 'Date', 'Time', 'Location', 'Description']);
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];
    rows.shift();

    const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';

    return rows.map((r, index) => ({
      id: index + 1,
      title: String(r[0] || ''),
      date: sanitizeDateString(r[1], tz),
      time: sanitizeTimeString(r[2], tz),
      location: String(r[3] || ''),
      description: String(r[4] || '')
    }));
  } catch (err) {
    return [];
  }
}

function addCalendarEvent(eventData) {
  try {
    const sheet = getOrCreateSheet('Events', ['Title', 'Date', 'Time', 'Location', 'Description']);
    const title = String(eventData.title || '').trim();
    const date = String(eventData.date || '').trim();
    const time = String(eventData.time || '').trim();
    const location = String(eventData.location || '').trim();
    const desc = String(eventData.description || '').trim();

    if (!title || !date) throw new Error('Title and Date are required.');

    sheet.appendRow([title, date, time, location, desc]);
    return { success: true, message: 'Event successfully published!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// -------------------------------------------------------------
// 6. NOTICES & CIRCULARS ENGINE
// -------------------------------------------------------------

function getNotices() {
  try {
    const sheet = getOrCreateSheet('Notices', ['Title', 'Category', 'Date', 'Details']);
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];
    rows.shift();

    const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';

    return rows.reverse().map((r, index) => ({
      id: index + 1,
      title: String(r[0] || ''),
      category: String(r[1] || 'General'),
      date: sanitizeDateString(r[2], tz),
      details: String(r[3] || '')
    }));
  } catch (err) {
    return [];
  }
}

function addNotice(noticeData) {
  try {
    const sheet = getOrCreateSheet('Notices', ['Title', 'Category', 'Date', 'Details']);
    const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
    const dateStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    sheet.appendRow([
      String(noticeData.title || '').trim(),
      String(noticeData.category || 'General').trim(),
      dateStr,
      String(noticeData.details || '').trim()
    ]);
    return { success: true, message: 'Notice posted successfully!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// -------------------------------------------------------------
// 7. STUDENT ROSTER & VERIFICATION
// -------------------------------------------------------------

function lookupStudent(studentId) {
  try {
    const sheet = getOrCreateSheet('Students', ['ID', 'Name', 'Class', 'Attendance', 'FeeStatus', 'Result']);
    const rows = sheet.getDataRange().getValues();
    const query = String(studentId).trim().toLowerCase();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === query) {
        return {
          success: true,
          student: {
            id: rows[i][0],
            name: rows[i][1],
            class: rows[i][2],
            attendance: rows[i][3],
            feeStatus: rows[i][4],
            result: rows[i][5]
          }
        };
      }
    }
    return { success: false, message: 'No student record found with ID: ' + studentId };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getAllStudents() {
  try {
    const sheet = getOrCreateSheet('Students', ['ID', 'Name', 'Class', 'Attendance', 'FeeStatus', 'Result']);
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];
    rows.shift();
    return rows.map(r => ({
      id: r[0],
      name: r[1],
      class: r[2],
      attendance: r[3],
      feeStatus: r[4],
      result: r[5]
    }));
  } catch (err) {
    return [];
  }
}

function saveStudentRecord(student) {
  try {
    const sheet = getOrCreateSheet('Students', ['ID', 'Name', 'Class', 'Attendance', 'FeeStatus', 'Result']);
    const rows = sheet.getDataRange().getValues();
    const targetId = String(student.id || '').trim().toLowerCase();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === targetId) {
        sheet.getRange(i + 1, 2).setValue(student.name);
        sheet.getRange(i + 1, 3).setValue(student.class);
        sheet.getRange(i + 1, 4).setValue(student.attendance);
        sheet.getRange(i + 1, 5).setValue(student.feeStatus);
        sheet.getRange(i + 1, 6).setValue(student.result);
        return { success: true, message: 'Student record updated.' };
      }
    }

    sheet.appendRow([
      student.id,
      student.name,
      student.class,
      student.attendance,
      student.feeStatus,
      student.result
    ]);
    return { success: true, message: 'New student added to roster.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// -------------------------------------------------------------
// 8. CRM INQUIRIES & DRIVE MEDIA UPLOADER
// -------------------------------------------------------------

function submitInquiry(lead) {
  try {
    const sheet = getOrCreateSheet('Inquiries', ['Timestamp', 'Name', 'Email', 'Service', 'Budget', 'Details']);
    const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
    const dateStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      dateStr,
      String(lead.name || ''),
      String(lead.email || ''),
      String(lead.service || ''),
      String(lead.budget || ''),
      String(lead.details || '')
    ]);
    return { success: true, message: 'Inquiry submitted successfully!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getLeads() {
  try {
    const sheet = getOrCreateSheet('Inquiries', ['Timestamp', 'Name', 'Email', 'Service', 'Budget', 'Details']);
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];
    rows.shift();

    const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';

    return rows.reverse().map(r => ({
      timestamp: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd HH:mm') : String(r[0] || ''),
      name: String(r[1] || ''),
      email: String(r[2] || ''),
      service: String(r[3] || ''),
      budget: String(r[4] || ''),
      details: String(r[5] || '')
    }));
  } catch (err) {
    return [];
  }
}

function uploadImageToDrive(base64Data, fileName, title, category) {
  try {
    const cleanBase64 = base64Data.replace(/^data:image\/(png|jpeg|jpg|webp);base64,/, '');
    const decoded = Utilities.base64Decode(cleanBase64);
    const blob = Utilities.newBlob(decoded, 'image/jpeg', fileName || 'portal_upload.jpg');

    let folderIterator = DriveApp.getFoldersByName('School_Portal_Uploads');
    let folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.createFolder('School_Portal_Uploads');

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const publicUrl = 'https://drive.google.com/uc?export=view&id=' + file.getId();

    const gallerySheet = getOrCreateSheet('Gallery', ['Title', 'Category', 'URL', 'Timestamp']);
    const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
    gallerySheet.appendRow([
      title || 'Campus Event',
      category || 'General',
      publicUrl,
      Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')
    ]);

    return { success: true, url: publicUrl, message: 'Image uploaded and linked to gallery!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getGalleryImages() {
  try {
    const sheet = getOrCreateSheet('Gallery', ['Title', 'Category', 'URL', 'Timestamp']);
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];
    rows.shift();
    return rows.reverse().map(r => ({
      title: r[0],
      category: r[1],
      url: r[2],
      date: r[3]
    }));
  } catch (err) {
    return [];
  }
}