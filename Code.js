/**
 * Greenwood Public School — Serverless Backend Engine
 * Google Apps Script & Google Sheets CRM
 */

const SPREADSHEET_ID = '1UF8xuxirPcGUwm3fsfSQVTghFzuloDhJJqpkjwi9454';

function doGet(e) {
  // If called with ?api=admissions, return direct JSON without iframe sandbox
  if (e && e.parameter && e.parameter.api === 'admissions') {
    const data = getAdmissions();
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Default HTML render
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Greenwood Public School — Official Portal & Admin Hub')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// -------------------------------------------------------------
// 1. DATABASE & SHEET HELPERS
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
// 2. SETTINGS, BRANDING & STAT COUNTERS
// -------------------------------------------------------------

function getSchoolSettings() {
  try {
    const sheet = getOrCreateSheet('Settings', ['Key', 'Value']);
    const rows = sheet.getDataRange().getValues();
    const settings = {
      schoolName: 'Greenwood Public School',
      tagline: 'AFFILIATED TO CBSE • NEW DELHI',
      noticeTicker: 'Admissions open for Session 2026-27 • Term 1 Datesheets Published • Scholarship Test Registrations Open',
      countStudents: '2,400+',
      countBoardPass: '100%',
      countFacultyRatio: '1:20',
      countCampusArea: '15 Acres'
    };

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0]) settings[String(rows[i][0])] = String(rows[i][1]);
    }

    return { success: true, data: settings };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function updateSchoolSettings(settings) {
  try {
    const sheet = getOrCreateSheet('Settings', ['Key', 'Value']);
    const data = sheet.getDataRange().getValues();
    const keys = Object.keys(settings);

    keys.forEach(key => {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(settings[key]);
          found = true;
          break;
        }
      }
      if (!found) {
        sheet.appendRow([key, settings[key]]);
      }
    });
    return { success: true, message: 'Settings and Stat Counters updated successfully!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// -------------------------------------------------------------
// 3. CALENDAR EVENTS ENGINE
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
// 4. NOTICES & CIRCULARS ENGINE
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
// 5. STUDENT ROSTER & VERIFICATION PORTAL
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
    return { success: false, message: 'No student found with ID: ' + studentId };
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
    return { success: true, message: 'New student added.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// -------------------------------------------------------------
// 6. ADMISSIONS PIPELINE & STATUS UPDATE
// -------------------------------------------------------------

// 1. Submit form from public page (sets Address to Col G, Status to Col H)
function submitAdmissionForm(formData) {
  try {
    const sheet = getOrCreateSheet('Admissions', [
      'Timestamp', 'StudentName', 'ParentName', 'Grade', 'Phone', 'Email', 'Address', 'Status'
    ]);
    const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
    const dateStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      dateStr,                                          // Col A: Timestamp
      String(formData.studentName || '').trim(),        // Col B: Student Name
      String(formData.parentName || '').trim(),         // Col C: Parent Name
      String(formData.grade || '').trim(),              // Col D: Grade
      String(formData.phone || '').trim(),              // Col E: Phone
      String(formData.email || '').trim(),              // Col F: Email
      String(formData.address || '').trim(),            // Col G: Address
      'Pending Review'                                  // Col H: Status (Default)
    ]);
    return { success: true, message: 'Application submitted successfully! Default status: Pending Review.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// 2. Fetch both separated Address and Status for the Admin table
function getAdmissions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Admissions');
    if (!sheet) {
      const sheets = ss.getSheets();
      sheet = sheets.find(s => s.getName().trim().toLowerCase() === 'admissions');
    }
    if (!sheet) return [];

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];
    rows.shift(); // Remove header row

    const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';

    const list = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0] && !r[1]) continue;

      let dateStr = '';
      if (r[0] instanceof Date) {
        dateStr = Utilities.formatDate(r[0], tz, 'yyyy-MM-dd HH:mm');
      } else {
        dateStr = String(r[0] || '').trim();
      }

      // Col G (index 6) is Address, Col H (index 7) is Status
      let addressVal = String(r[6] || '').trim();
      let statusVal = r[7] ? String(r[7]).trim() : 'Pending Review';

      // Fallback if older data placed address directly into status column
      if (!r[7] && addressVal && (addressVal.includes('Review') || addressVal.includes('Scheduled') || addressVal.includes('Approved') || addressVal.includes('Enrolled') || addressVal.includes('Rejected'))) {
        statusVal = addressVal;
        addressVal = '';
      }

      list.push({
        rowNumber: i + 2, // 1-based row index in Sheet
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

    return list.reverse();
  } catch (err) {
    return [];
  }
}

// 3. Update only the Status (Column H / Col 8) without altering Address
function updateAdmissionStatus(rowNumber, newStatus) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Admissions');
    if (!sheet) throw new Error('Admissions sheet not found.');

    const row = Number(rowNumber);
    if (!row || row < 2) throw new Error('Invalid row number.');

    // Ensure header names
    sheet.getRange(1, 7).setValue('Address');
    sheet.getRange(1, 8).setValue('Status');

    // Update Column H (8)
    sheet.getRange(row, 8).setValue(newStatus);

    return { success: true, message: 'Status updated to "' + newStatus + '"!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}
// -------------------------------------------------------------
// 7. LEADS & INQUIRIES CRM
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

// -------------------------------------------------------------
// 8. GOOGLE DRIVE MEDIA UPLOADER & GALLERY
// -------------------------------------------------------------

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