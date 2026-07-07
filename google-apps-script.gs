/**
 * DIU CPC Basic Programming Course — Attendance System
 * Google Apps Script Web App
 *
 * HOW TO UPDATE:
 * 1. Replace SHEET_ID below with your Google Sheet ID
 * 2. Save → Deploy → New Deployment → Web App (Anyone, Execute as Me)
 */

// ─── CONFIGURATION ───────────────────────────────────────────────────────────
var SHEET_ID = '1G3Lk36x09JorHAVlQF_rb9JCZkxxu11_kORAJae7t08'; // Pre-configured sheet ID
var SHEET_NAME = 'Attendance';
// ─────────────────────────────────────────────────────────────────────────────

// ── Helper: get or create the sheet and add headers if empty ─────────────────
function getSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    var headers = [
      '#', 'Timestamp', 'Full Name', 'Email',
      'Student ID', 'Roll Number', 'Batch', 'Semester',
      'Department', 'CPC Membership ID',
      'Lecture Number', 'Lecture Date', 'Time Submitted',
      'Rating', 'Difficulty', 'Comment'
    ];
    sheet.appendRow(headers);

    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a73e8');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);
    sheet.setFrozenRows(1);

    var widths = [40, 160, 160, 200, 140, 80, 70, 90, 100, 150, 100, 110, 110, 70, 100, 250];
    widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  }
  return sheet;
}

// ── Helper: append one data row ───────────────────────────────────────────────
function appendRecord(data) {
  var sheet = getSheet();
  var rowNum = sheet.getLastRow(); // 1-based, header already there

  sheet.appendRow([
    rowNum,
    new Date().toLocaleString('en-BD', { timeZone: 'Asia/Dhaka' }),
    data.name        || '',
    data.email       || '',
    data.studentId   || '',
    data.roll        || '',
    data.batch       || '',
    data.semester    ? data.semester + ' Semester' : '',
    data.dept        || 'CSE',
    data.cpcMembershipId || '',
    data.lecture     ? 'Lecture #' + data.lecture : '',
    data.date        || '',
    data.time        || '',
    data.rating      || '',
    data.difficulty  || '',
    data.comment     || ''
  ]);

  // Alternate row coloring
  if (rowNum % 2 === 0) {
    sheet.getRange(sheet.getLastRow(), 1, 1, 16).setBackground('#f8f9fa');
  }
}

// ── POST handler (primary method) ─────────────────────────────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    appendRecord(data);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── GET handler (fallback — receives data via URL query params) ───────────────
function doGet(e) {
  try {
    var p = e.parameter;
    // If query params carry attendance data, save them
    if (p && p.name && p.email) {
      appendRecord({
        name:           p.name,
        email:          p.email,
        studentId:      p.studentId      || '',
        roll:           p.roll           || '',
        batch:          p.batch          || '',
        semester:       p.semester       || '',
        dept:           p.dept           || 'CSE',
        cpcMembershipId: p.cpcMembershipId || '',
        lecture:        p.lecture        || '',
        date:           p.date           || '',
        time:           p.time           || '',
        rating:         p.rating         || '',
        difficulty:     p.difficulty     || '',
        comment:        p.comment        || '',
      });
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'success', method: 'GET' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Health check
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'DIU CPC Attendance Script is running.' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
