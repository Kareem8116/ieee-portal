// ═══════════════════════════════════════════════════════════
//  IEEE KSB — Student Portal & Grade Management
//  Apps Script Backend
//  
//  SHEET STRUCTURE:
//    Sheet "Students"  : ID | Name | Email
//    Sheet "Grades"    : Student ID | Assessment | Grade | Max Grade | Date | Notes
//    Sheet "Submissions": Name | Student ID | Email | Score | Max Grade | Total Qs | Correct | Time Taken | Tab Switches | Submitted At | Quiz Title
// ═══════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1_V9NN1yTY_AmoZLx6N3NtI3FUFYDbbpclgeir9ZpoHw';
const ADMIN_KEY      = 'KSA@2006';

// Pre-seeded student roster (24 students)
const STUDENT_ROSTER = [
  { id: '47784', name: 'Mohammed Elsayed', email: '' },
  { id: '92169', name: 'Reem Abdelhamid',  email: '' },
  { id: '96125', name: 'Mostafa Haasan',   email: '' },
  { id: '31610', name: 'Mohamed Tamer',    email: '' },
  { id: '92545', name: 'Ziad khaled',      email: '' },
  { id: '95186', name: 'Menna Mosaad',     email: '' },
  { id: '35567', name: 'Amr Sabry',        email: '' },
  { id: '11006', name: 'Khloud Mahdy',     email: '' },
  { id: '32755', name: 'Mariam Shaban',    email: '' },
  { id: '52345', name: 'Muhammed Fouad',   email: '' },
  { id: '22294', name: 'Nada Waleed',      email: '' },
  { id: '71430', name: 'Ahmed Mohamed',    email: '' },
  { id: '76903', name: 'Maryam Taha',      email: '' },
  { id: '38231', name: 'Sama Saber',       email: '' },
  { id: '40106', name: 'Bothina Atef',     email: '' },
  { id: '24055', name: 'Jana fahim',       email: '' },
  { id: '73760', name: 'Eman nabih',       email: '' },
  { id: '26087', name: 'Enas Ahmed',       email: '' },
  { id: '12959', name: 'Arwa Tamer',       email: '' },
  { id: '35206', name: 'Omar ibrahim',     email: '' },
  { id: '07174', name: 'Donia Ashraf',     email: '' },
  { id: '05375', name: 'Nesreen Hossam',   email: '' },
  { id: '39236', name: 'Mariam Hamdy',     email: '' },
  { id: '28937', name: 'Habiba saad',      email: '' },
];

// ───────────────────────────────────────────────────────────
//  SETUP — Run once manually from Apps Script editor
// ───────────────────────────────────────────────────────────
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // ── Students sheet ──
  let studentsSheet = ss.getSheetByName('Students');
  if (!studentsSheet) {
    studentsSheet = ss.insertSheet('Students');
  }
  if (studentsSheet.getLastRow() === 0) {
    studentsSheet.appendRow(['ID', 'Name', 'Email']);
    studentsSheet.getRange(1, 1, 1, 3).setFontWeight('bold')
      .setBackground('#003865').setFontColor('#ffffff');
  }

  // Seed roster if sheet is empty (only header row)
  if (studentsSheet.getLastRow() <= 1) {
    STUDENT_ROSTER.forEach(s => {
      studentsSheet.appendRow([s.id, s.name, s.email]);
    });
  }

  // ── Grades sheet ──
  let gradesSheet = ss.getSheetByName('Grades');
  if (!gradesSheet) {
    gradesSheet = ss.insertSheet('Grades');
    gradesSheet.appendRow(['Student ID', 'Student Name', 'Assessment', 'Grade', 'Max Grade', 'Date', 'Notes']);
    gradesSheet.getRange(1, 1, 1, 7).setFontWeight('bold')
      .setBackground('#003865').setFontColor('#ffffff');
  }

  // ── Submissions sheet ──
  let subsSheet = ss.getSheetByName('Submissions');
  if (!subsSheet) {
    subsSheet = ss.insertSheet('Submissions');
    subsSheet.appendRow(['Name', 'Student ID', 'Email', 'Quiz Title', 'Score', 'Max Grade', 'Total Questions', 'Correct Answers', 'Time Taken (s)', 'Tab Switches', 'Submitted At']);
    subsSheet.getRange(1, 1, 1, 11).setFontWeight('bold')
      .setBackground('#003865').setFontColor('#ffffff');
  }

  SpreadsheetApp.getUi().alert('✅ Setup complete! Sheets created and students seeded.');
}

// ───────────────────────────────────────────────────────────
//  doGet — Read operations
// ───────────────────────────────────────────────────────────
function doGet(e) {
  const params  = e.parameter || {};
  const action  = params.action || '';
  const key     = params.key   || '';

  // All read actions require admin key
  if (key !== ADMIN_KEY) {
    return jsonOut({ error: 'Unauthorized' });
  }

  if (action === 'getStudents') {
    return jsonOut(getStudents());
  }

  if (action === 'getGrades') {
    const studentId = params.studentId || '';
    return jsonOut(getStudentGrades(studentId));
  }

  if (action === 'getSubmissions') {
    return jsonOut(getSubmissions());
  }

  return jsonOut({ error: 'Unknown action' });
}

// ───────────────────────────────────────────────────────────
//  doPost — Write operations
// ───────────────────────────────────────────────────────────
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch(err) {}

  const action   = body.action   || '';
  const adminKey = body.adminKey || '';

  // submitQuiz does NOT need adminKey (student-facing)
  if (action === 'submitQuiz') {
    return jsonOut(submitQuiz(body));
  }

  // All other writes need adminKey
  if (adminKey !== ADMIN_KEY) {
    return jsonOut({ error: 'Unauthorized' });
  }

  if (action === 'setGrade')   return jsonOut(setGrade(body));
  if (action === 'addStudent') return jsonOut(addStudent(body));

  return jsonOut({ error: 'Unknown action' });
}

// ───────────────────────────────────────────────────────────
//  HELPERS
// ───────────────────────────────────────────────────────────
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function getHeaders(sheet) {
  if (sheet.getLastRow() === 0) return {};
  const row = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  row.forEach((h, i) => { map[h] = i; });
  return map;
}

// ───────────────────────────────────────────────────────────
//  READ: Students list
// ───────────────────────────────────────────────────────────
function getStudents() {
  const sheet = getSheet('Students');
  if (!sheet) return { students: [] };
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { students: [] };
  const h = getHeaders(sheet);
  const students = data.slice(1).map(row => ({
    id:    String(row[h['ID']]   || '').trim(),
    name:  String(row[h['Name']] || '').trim(),
    email: String(row[h['Email']]|| '').trim(),
  })).filter(s => s.id);
  return { students };
}

// ───────────────────────────────────────────────────────────
//  READ: Student grades (by ID) — public (student portal)
//  Note: We use adminKey for validation but allow student
//        to access their own grades via portal with just ID
// ───────────────────────────────────────────────────────────
function getStudentGrades(studentId) {
  if (!studentId) return { error: 'Student ID required' };

  // Look up student name
  const studentsSheet = getSheet('Students');
  let studentName = '';
  if (studentsSheet) {
    const data = studentsSheet.getDataRange().getValues();
    const h    = getHeaders(studentsSheet);
    const row  = data.slice(1).find(r => String(r[h['ID']]).trim() === String(studentId).trim());
    if (row) studentName = String(row[h['Name']] || '');
  }

  // Also check Submissions for auto-submitted quiz scores
  const subsSheet = getSheet('Submissions');
  const quizGrades = [];
  if (subsSheet && subsSheet.getLastRow() > 1) {
    const data = subsSheet.getDataRange().getValues();
    const h    = getHeaders(subsSheet);
    data.slice(1).forEach(row => {
      if (String(row[h['Student ID']]).trim() === String(studentId).trim()) {
        quizGrades.push({
          assessment: String(row[h['Quiz Title']] || 'Quiz'),
          grade:      row[h['Score']],
          maxGrade:   row[h['Max Grade']],
          date:       row[h['Submitted At']] ? new Date(row[h['Submitted At']]).toLocaleDateString('en-GB') : '',
          notes:      `${row[h['Correct Answers']]} / ${row[h['Total Questions']]} correct`,
        });
      }
    });
  }

  // Manual grades
  const gradesSheet = getSheet('Grades');
  const manualGrades = [];
  if (gradesSheet && gradesSheet.getLastRow() > 1) {
    const data = gradesSheet.getDataRange().getValues();
    const h    = getHeaders(gradesSheet);
    data.slice(1).forEach(row => {
      if (String(row[h['Student ID']]).trim() === String(studentId).trim()) {
        manualGrades.push({
          assessment: String(row[h['Assessment']] || ''),
          grade:      row[h['Grade']],
          maxGrade:   row[h['Max Grade']],
          date:       row[h['Date']] ? new Date(row[h['Date']]).toLocaleDateString('en-GB') : '',
          notes:      String(row[h['Notes']] || ''),
        });
      }
    });
  }

  if (!studentName && quizGrades.length === 0 && manualGrades.length === 0) {
    return { error: 'لم يتم العثور على هذا الرقم. تحقق من رقمك الجامعي.' };
  }

  return {
    student: { id: studentId, name: studentName || 'Student' },
    grades:  [...quizGrades, ...manualGrades],
  };
}

// ───────────────────────────────────────────────────────────
//  READ: All quiz submissions (admin)
// ───────────────────────────────────────────────────────────
function getSubmissions() {
  const sheet = getSheet('Submissions');
  if (!sheet || sheet.getLastRow() <= 1) return { submissions: [] };
  const data = sheet.getDataRange().getValues();
  const h    = getHeaders(sheet);
  const submissions = data.slice(1).map(row => ({
    name:           String(row[h['Name']]            || ''),
    studentId:      String(row[h['Student ID']]      || ''),
    email:          String(row[h['Email']]           || ''),
    quizTitle:      String(row[h['Quiz Title']]      || ''),
    score:          row[h['Score']],
    maxGrade:       row[h['Max Grade']],
    totalQuestions: row[h['Total Questions']],
    correctAnswers: row[h['Correct Answers']],
    timeTaken:      row[h['Time Taken (s)']],
    tabSwitchCount: row[h['Tab Switches']],
    submittedAt:    row[h['Submitted At']] ? new Date(row[h['Submitted At']]).toLocaleString('en-GB') : '',
  })).reverse(); // newest first
  return { submissions };
}

// ───────────────────────────────────────────────────────────
//  WRITE: Submit quiz (student-facing, no admin key)
// ───────────────────────────────────────────────────────────
function submitQuiz(body) {
  const sheet = getSheet('Submissions');
  if (!sheet) return { error: 'Submissions sheet not found. Run setupSheets() first.' };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { return { error: 'Server busy, try again.' }; }

  try {
    const now = new Date();
    sheet.appendRow([
      body.name           || '',
      body.studentId      || '',
      body.email          || '',
      body.quizTitle      || '',
      body.score          || 0,
      body.maxGrade       || 0,
      body.totalQuestions || 0,
      body.correctAnswers || 0,
      body.timeTaken      || 0,
      body.tabSwitchCount || 0,
      now,
    ]);
    // Format submitted-at column as text
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 11).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ───────────────────────────────────────────────────────────
//  WRITE: Set / update a manual grade (admin)
// ───────────────────────────────────────────────────────────
function setGrade(body) {
  const { studentId, assessment, grade, notes } = body;
  if (!studentId || !assessment || grade === undefined) {
    return { error: 'Missing required fields.' };
  }

  // Look up student name
  const studentsSheet = getSheet('Students');
  let studentName = '';
  if (studentsSheet) {
    const data = studentsSheet.getDataRange().getValues();
    const h    = getHeaders(studentsSheet);
    const row  = data.slice(1).find(r => String(r[h['ID']]).trim() === String(studentId).trim());
    if (row) studentName = String(row[h['Name']] || '');
  }

  // Max grade now comes from the request body (flexible)
  const maxGrade = body.maxGrade !== undefined ? body.maxGrade : '';

  const sheet = getSheet('Grades');
  if (!sheet) return { error: 'Grades sheet not found. Run setupSheets() first.' };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { return { error: 'Server busy.' }; }

  try {
    // Check if entry already exists → update it
    if (sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      const h    = getHeaders(sheet);
      for (let i = 1; i < data.length; i++) {
        if (
          String(data[i][h['Student ID']]).trim() === String(studentId).trim() &&
          String(data[i][h['Assessment']]).trim()  === String(assessment).trim()
        ) {
          // Update existing row
          sheet.getRange(i + 1, h['Grade'] + 1).setValue(grade);
          sheet.getRange(i + 1, h['Notes'] + 1).setValue(notes || '');
          sheet.getRange(i + 1, h['Date']  + 1).setValue(new Date());
          return { ok: true, updated: true };
        }
      }
    }
    // Append new row
    sheet.appendRow([
      studentId,
      studentName,
      assessment,
      grade,
      maxGrade,
      new Date(),
      notes || '',
    ]);
    sheet.getRange(sheet.getLastRow(), 6).setNumberFormat('yyyy-mm-dd');
    return { ok: true, updated: false };
  } finally {
    lock.releaseLock();
  }
}

// ───────────────────────────────────────────────────────────
//  WRITE: Add student manually (admin)
// ───────────────────────────────────────────────────────────
function addStudent(body) {
  const { studentId, name, email } = body;
  if (!studentId || !name) return { error: 'ID and Name are required.' };

  const sheet = getSheet('Students');
  if (!sheet) return { error: 'Students sheet not found.' };

  // Check for duplicate
  if (sheet.getLastRow() > 1) {
    const data = sheet.getDataRange().getValues();
    const h    = getHeaders(sheet);
    if (data.slice(1).some(r => String(r[h['ID']]).trim() === String(studentId).trim())) {
      return { error: `Student ID ${studentId} already exists.` };
    }
  }

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { return { error: 'Server busy.' }; }
  try {
    sheet.appendRow([studentId, name, email || '']);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}
