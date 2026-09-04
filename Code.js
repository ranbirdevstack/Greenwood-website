function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Greenwood Public School | Portal & Admin')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers) sheet.appendRow(headers);
  }
  return sheet;
}

function getSchoolSettings() {
  const sheet = getOrCreateSheet('Settings', ['Key', 'Value']);
  const rows = sheet.getDataRange().getValues();
  let settings = { name: 'Greenwood Public School', tagline: 'Affiliated to CBSE • New Delhi' };
  
  if (rows.length > 1) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === 'SchoolName') settings.name = String(rows[i][1]);
      if (String(rows[i][0]).trim() === 'SchoolTagline') settings.tagline = String(rows[i][1]);
    }
  } else {
    sheet.appendRow(['SchoolName', settings.name]);
    sheet.appendRow(['SchoolTagline', settings.tagline]);
  }
  return settings;
}

function updateSchoolSettings(data) {
  const sheet = getOrCreateSheet('Settings', ['Key', 'Value']);
  const rows = sheet.getDataRange().getValues();
  let nameFound = false, tagFound = false;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === 'SchoolName') {
      sheet.getRange(i + 1, 2).setValue(data.name);
      nameFound = true;
    }
    if (String(rows[i][0]).trim() === 'SchoolTagline') {
      sheet.getRange(i + 1, 2).setValue(data.tagline);
      tagFound = true;
    }
  }
  if (!nameFound) sheet.appendRow(['SchoolName', data.name]);
  if (!tagFound) sheet.appendRow(['SchoolTagline', data.tagline]);
  return { success: true };
}

function getAdminStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  function getCount(tab) {
    const s = ss.getSheetByName(tab);
    if (!s) return 0;
    const count = s.getLastRow();
    return count > 1 ? count - 1 : 0;
  }

  return {
    students: getCount('Students'),
    events: getCount('Events'),
    notices: getCount('Notices'),
    gallery: getCount('Gallery'),
    admissions: getCount('Admissions'),
    inquiries: getCount('Inquiries')
  };
}

function getNotices() {
  const sheet = getOrCreateSheet('Notices', ['Date', 'Category', 'Title', 'Content']);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  rows.shift();
  return rows.map(r => ({
    date: r[0] ? String(r[0]) : '',
    category: r[1] ? String(r[1]) : 'General',
    title: r[2] ? String(r[2]) : '',
    content: r[3] ? String(r[3]) : ''
  })).reverse();
}

function getEvents() {
  const sheet = getOrCreateSheet('Events', ['Title', 'Date', 'Time', 'Location', 'Description']);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  rows.shift();
  
  const timeZone = Session.getScriptTimeZone() || 'Asia/Kolkata';

  return rows.map(r => {
    // 1. Clean up Date
    let dateStr = '';
    if (r[1] instanceof Date) {
      dateStr = Utilities.formatDate(r[1], timeZone, 'yyyy-MM-dd');
    } else if (r[1]) {
      dateStr = String(r[1]).split('T')[0].trim();
    }

    let timeStr = '';
    if (r[2] instanceof Date) {
      timeStr = Utilities.formatDate(r[2], timeZone, 'hh:mm a');
    } else if (r[2]) {
      timeStr = String(r[2]).trim();
    
      if (timeStr.includes('1899')) {
        const timeMatch = timeStr.match(/\d{2}:\d{2}(:\d{2})?/);
        timeStr = timeMatch ? timeMatch[0] : '';
      }
    }

    return {
      title: String(r[0] || ''),
      date: dateStr,
      time: timeStr,
      location: String(r[3] || ''),
      description: String(r[4] || '')
    };
  });
}

function getGallery() {
  const sheet = getOrCreateSheet('Gallery', ['Title', 'Category', 'ImageURL']);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  rows.shift();
  return rows.map(r => ({ 
    title: String(r[0] || ''), 
    category: String(r[1] || 'General'), 
    url: String(r[2] || '') 
  })).reverse();
}

function getAdmissions() {
  const sheet = getOrCreateSheet('Admissions', ['Timestamp', 'StudentName', 'ParentName', 'Grade', 'Phone', 'Email', 'Status']);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  rows.shift();
  return rows.map(r => ({
    time: String(r[0]),
    student: String(r[1]),
    parent: String(r[2]),
    grade: String(r[3]),
    phone: String(r[4]),
    email: String(r[5]),
    status: String(r[6])
  })).reverse();
}

function findStudentById(id) {
  const sheet = getOrCreateSheet('Students', ['StudentID', 'Name', 'Grade', 'Section', 'Attendance', 'TermResult', 'FeeStatus']);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { found: false };
  rows.shift();
  const match = rows.find(r => String(r[0]).trim().toUpperCase() === String(id).trim().toUpperCase());
  if (match) {
    return {
      found: true,
      id: String(match[0]),
      name: String(match[1]),
      grade: String(match[2]),
      section: String(match[3]),
      attendance: String(match[4]),
      result: String(match[5]),
      feeStatus: String(match[6])
    };
  }
  return { found: false };
}

function submitAdmission(data) {
  const sheet = getOrCreateSheet('Admissions', ['Timestamp', 'StudentName', 'ParentName', 'Grade', 'Phone', 'Email', 'Status']);
  sheet.appendRow([
    new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
    data.studentName,
    data.parentName,
    data.grade,
    data.phone,
    data.email,
    'Pending Review'
  ]);
  return { success: true };
}

function submitInquiry(data) {
  const sheet = getOrCreateSheet('Inquiries', ['Timestamp', 'Name', 'Email', 'Subject', 'Message']);
  sheet.appendRow([
    new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
    data.name,
    data.email,
    data.subject,
    data.message
  ]);
  return { success: true };
}

function addNotice(data) {
  getOrCreateSheet('Notices').appendRow([data.date, data.category, data.title, data.content]);
  return { success: true };
}

function addEvent(data) {
  getOrCreateSheet('Events').appendRow([data.title, data.date, data.time, data.location, data.desc]);
  return { success: true };
}

function addStudent(data) {
  getOrCreateSheet('Students').appendRow([data.id, data.name, data.grade, data.section, data.attendance, data.result, data.feeStatus]);
  return { success: true };
}

function addGalleryItem(data) {
  const sheet = getOrCreateSheet('Gallery', ['Title', 'Category', 'ImageURL']);
  sheet.appendRow([data.title, data.category, data.url]);
  return { success: true };
}

function uploadImageToDrive(fileData) {
  try {
    const contentType = fileData.type;
    const bytes = Utilities.base64Decode(fileData.base64.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileData.name);

    const folder = DriveApp.getRootFolder();
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId = file.getId();
    const publicUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1000';

    const sheet = getOrCreateSheet('Gallery', ['Title', 'Category', 'ImageURL']);
    sheet.appendRow([fileData.title, fileData.category, publicUrl]);

    return { success: true, url: publicUrl };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}
function testAuthorize() {
  DriveApp.getRootFolder();
  SpreadsheetApp.getActiveSpreadsheet();
}