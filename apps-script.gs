// ═══════════════════════════════════════════════════════════
//  IEEE KSB — Student Portal & Grade Management
//  Apps Script Backend
//  
//  SHEET STRUCTURE:
//    Sheet "Students"   : ID | Name | Email
//    Sheet "Grades"     : Student ID | Assessment | Grade | Max Grade | Date | Notes
//    Sheet "Submissions": Name | Student ID | Email | Quiz Title | Score | Max Grade | Total Questions | Correct Answers | Time Taken (s) | Tab Switches | Submitted At | Answers JSON
// ═══════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1_V9NN1yTY_AmoZLx6N3NtI3FUFYDbbpclgeir9ZpoHw';
const ADMIN_KEY      = 'KSA@2006';

// Pre-seeded student roster (21 students)
const STUDENT_ROSTER = [
  { id: '47784', name: 'Mohammed Elsayed', email: '' },
  { id: '92169', name: 'Reem Abdelhamid',  email: '' },
  { id: '96125', name: 'Mostafa Haasan',   email: '' },
  { id: '31610', name: 'Mohamed Tamer',    email: '' },
  { id: '92545', name: 'Ziad khaled',      email: '' },
  { id: '95186', name: 'Menna Mosaad',     email: '' },
  { id: '11006', name: 'Khloud Mahdy',     email: '' },
  { id: '32755', name: 'Mariam Shaban',    email: '' },
  { id: '52345', name: 'Muhammed Fouad',   email: '' },
  { id: '22294', name: 'Nada Waleed',      email: '' },
  { id: '71430', name: 'Ahmed Mohamed',    email: '' },
  { id: '76903', name: 'Maryam Taha',      email: '' },
  { id: '38231', name: 'Sama Saber',       email: '' },
  { id: '40106', name: 'Bothina Atef',     email: '' },
  { id: '24055', name: 'Jana fahim',       email: '' },
  { id: '26087', name: 'Enas Ahmed',       email: '' },
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
    subsSheet.appendRow(['Name', 'Student ID', 'Email', 'Quiz Title', 'Score', 'Max Grade', 'Total Questions', 'Correct Answers', 'Time Taken (s)', 'Tab Switches', 'Submitted At', 'Answers JSON']);
    subsSheet.getRange(1, 1, 1, 12).setFontWeight('bold')
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

  if (action === 'getAllGrades') {
    return jsonOut(getAllGrades());
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

  if (action === 'setGrade')              return jsonOut(setGrade(body));
  if (action === 'deleteGrade')           return jsonOut(deleteGrade(body));
  if (action === 'updateSubmissionGrade') return jsonOut(updateSubmissionGrade(body));
  if (action === 'deleteSubmission')      return jsonOut(deleteSubmission(body));
  if (action === 'addStudent')            return jsonOut(addStudent(body));
  if (action === 'deleteStudent')         return jsonOut(deleteStudent(body));
  if (action === 'grantRetake')           return jsonOut(grantRetake(body));

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
//  READ: All grades for grade book (admin)
// ───────────────────────────────────────────────────────────
function getAllGrades() {
  // Manual grades
  const gs = getSheet('Grades');
  const manualGrades = [];
  if (gs && gs.getLastRow() > 1) {
    const data = gs.getDataRange().getValues();
    const h    = getHeaders(gs);
    data.slice(1).forEach((row, idx) => {
      manualGrades.push({
        rowNum:      idx + 2,
        studentId:   String(row[h['Student ID']]   || '').trim(),
        studentName: String(row[h['Student Name']] || '').trim(),
        assessment:  String(row[h['Assessment']]   || '').trim(),
        grade:       row[h['Grade']],
        maxGrade:    row[h['Max Grade']],
        date:        row[h['Date']] ? new Date(row[h['Date']]).toLocaleDateString('en-GB') : '',
        notes:       String(row[h['Notes']] || ''),
      });
    });
  }

  // Quiz submissions
  const ss = getSheet('Submissions');
  const submissions = [];
  if (ss && ss.getLastRow() > 1) {
    const data = ss.getDataRange().getValues();
    const h    = getHeaders(ss);
    data.slice(1).forEach((row, idx) => {
      submissions.push({
        rowNum:      idx + 2,
        studentId:   String(row[h['Student ID']] || '').trim(),
        name:        String(row[h['Name']]       || '').trim(),
        quizTitle:   String(row[h['Quiz Title']] || '').trim(),
        score:       row[h['Score']],
        maxGrade:    row[h['Max Grade']],
        submittedAt: row[h['Submitted At']] ? new Date(row[h['Submitted At']]).toLocaleDateString('en-GB') : '',
      });
    });
  }

  return { manualGrades, submissions };
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
function normalizeStudentId(id) {
  return String(id || '').trim().replace(/^0+/, '') || '0';
}

function getStudentGrades(studentId) {
  if (!studentId) return { error: 'Student ID required' };
  const normId = normalizeStudentId(studentId);

  // Look up student name
  const studentsSheet = getSheet('Students');
  let studentName = '';
  if (studentsSheet) {
    const data = studentsSheet.getDataRange().getValues();
    const h    = getHeaders(studentsSheet);
    const row  = data.slice(1).find(r => normalizeStudentId(r[h['ID']]) === normId);
    if (row) studentName = String(row[h['Name']] || '');
  }

  // Also check Submissions for auto-submitted quiz scores and latest submission timestamps
  const subsSheet = getSheet('Submissions');
  const quizGrades = [];
  const submissionsByTitle = {};
  const submissionDetails = {};
  if (subsSheet && subsSheet.getLastRow() > 1) {
    const data = subsSheet.getDataRange().getValues();
    const h    = getHeaders(subsSheet);
    const attemptCounts = {};
    const studentRows = data.slice(1).filter(row => normalizeStudentId(row[h['Student ID']]) === normId);
    studentRows.forEach(row => {
      const title = String(row[h['Quiz Title']] || 'Quiz').trim();
      attemptCounts[title] = (attemptCounts[title] || 0) + 1;
    });
    const seenCounts = {};
    studentRows.forEach(row => {
      const title = String(row[h['Quiz Title']] || 'Quiz').trim();
      seenCounts[title] = (seenCounts[title] || 0) + 1;
      const attemptNum = seenCounts[title];
      const totalAttempts = attemptCounts[title];
      const attemptLabel = totalAttempts > 1 ? `Attempt ${attemptNum} of ${totalAttempts}` : '1 attempt';

      const subMs = row[h['Submitted At']] ? new Date(row[h['Submitted At']]).getTime() : 0;
      if (!submissionsByTitle[title] || subMs > submissionsByTitle[title]) {
        submissionsByTitle[title] = subMs;
      }
      const ansRaw = h['Answers JSON'] !== undefined ? String(row[h['Answers JSON']] || '') : '';
      let detailIdx = undefined;
      if (ansRaw) {
        try {
          if (!submissionDetails[title]) submissionDetails[title] = [];
          detailIdx = submissionDetails[title].length;
          submissionDetails[title].push({
            attemptNum,
            score:          Number(row[h['Score']] || 0),
            maxGrade:       Number(row[h['Max Grade']] || 0),
            correct:        Number(row[h['Correct Answers']] || 0),
            totalQuestions: Number(row[h['Total Questions']] || 0),
            timeTaken:      Number(row[h['Time Taken (s)']] || 0),
            submittedAt:    row[h['Submitted At']] ? new Date(row[h['Submitted At']]).toLocaleString('en-GB') : '',
            answers:        JSON.parse(ansRaw),
          });
        } catch(e) { detailIdx = undefined; }
      }
      quizGrades.push({
        assessment: title,
        grade:      row[h['Score']],
        maxGrade:   row[h['Max Grade']],
        date:       row[h['Submitted At']] ? new Date(row[h['Submitted At']]).toLocaleDateString('en-GB') : '',
        notes:      `${row[h['Correct Answers']]} / ${row[h['Total Questions']]} correct • ${attemptLabel}`,
        attemptIdx: detailIdx,
      });
    });
  }

  // Manual grades
  const gradesSheet = getSheet('Grades');
  const manualGrades = [];
  if (gradesSheet && gradesSheet.getLastRow() > 1) {
    const data = gradesSheet.getDataRange().getValues();
    const h    = getHeaders(gradesSheet);
    data.slice(1).forEach(row => {
      if (normalizeStudentId(row[h['Student ID']]) === normId) {
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

  // Retake grants (max of ALL and specific studentId per examId)
  const retakeGrants = {};
  const retakesSheet = getSheet('Retakes');
  if (retakesSheet && retakesSheet.getLastRow() > 1) {
    const data = retakesSheet.getDataRange().getValues();
    const h    = getHeaders(retakesSheet);
    data.slice(1).forEach(row => {
      const target = String(row[h['Target']] || '').trim();
      const examId = String(row[h['Exam ID']] || '').trim();
      const grantedAt = Number(row[h['Granted At']] || 0);
      if (!examId || !grantedAt) return;
      if (target === 'ALL' || normalizeStudentId(target) === normId) {
        if (!retakeGrants[examId] || grantedAt > retakeGrants[examId]) {
          retakeGrants[examId] = grantedAt;
        }
      }
    });
  }

  if (!studentName && quizGrades.length === 0 && manualGrades.length === 0) {
    return { error: 'لم يتم العثور على هذا الرقم. تحقق من رقمك الجامعي.' };
  }

  return {
    student: { id: studentId, name: studentName || 'Student' },
    grades:  [...quizGrades, ...manualGrades],
    submissionsByTitle,
    submissionDetails,
    retakeGrants,
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
  const submissions = data.slice(1).map((row, idx) => ({
    rowNum:         idx + 2,
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
    answersJson:    h['Answers JSON'] !== undefined ? String(row[h['Answers JSON']] || '') : '',
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
      body.answersJson    || '',
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
      const normId = normalizeStudentId(studentId);
      for (let i = 1; i < data.length; i++) {
        if (
          normalizeStudentId(data[i][h['Student ID']]) === normId &&
          String(data[i][h['Assessment']]).trim()     === String(assessment).trim()
        ) {
          // Update existing row
          sheet.getRange(i + 1, h['Grade'] + 1).setValue(grade);
          if (maxGrade !== '' && h['Max Grade'] !== undefined) {
            sheet.getRange(i + 1, h['Max Grade'] + 1).setValue(maxGrade);
          }
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
//  WRITE: Delete a specific manual grade (admin)
// ───────────────────────────────────────────────────────────
function deleteGrade(body) {
  const { studentId, assessment, rowNum } = body;
  const sheet = getSheet('Grades');
  if (!sheet || sheet.getLastRow() <= 1) return { error: 'Grades sheet empty.' };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { return { error: 'Server busy.' }; }

  try {
    const data = sheet.getDataRange().getValues();
    const h    = getHeaders(sheet);
    const normId = studentId ? normalizeStudentId(studentId) : '';
    const r = Number(rowNum);
    if (r >= 2 && r <= data.length) {
      const rowId = normalizeStudentId(data[r - 1][h['Student ID']]);
      const rowAs = String(data[r - 1][h['Assessment']] || '').trim();
      if ((!normId || rowId === normId) && (!assessment || rowAs === String(assessment).trim())) {
        sheet.deleteRow(r);
        return { ok: true };
      }
    }
    if (!normId || !assessment) return { error: 'Student ID and Assessment required.' };
    for (let i = data.length - 1; i >= 1; i--) {
      if (
        normalizeStudentId(data[i][h['Student ID']]) === normId &&
        String(data[i][h['Assessment']]).trim()     === String(assessment).trim()
      ) {
        sheet.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { error: 'Grade entry not found.' };
  } finally {
    lock.releaseLock();
  }
}

// ───────────────────────────────────────────────────────────
//  WRITE: Update a quiz submission grade (admin)
// ───────────────────────────────────────────────────────────
function updateSubmissionGrade(body) {
  const { rowNum, studentId, score, maxGrade } = body;
  if (!rowNum || score === undefined) return { error: 'rowNum and score are required.' };

  const sheet = getSheet('Submissions');
  if (!sheet || sheet.getLastRow() <= 1) return { error: 'Submissions sheet empty.' };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { return { error: 'Server busy.' }; }

  try {
    const r = Number(rowNum);
    if (r < 2 || r > sheet.getLastRow()) return { error: 'Submission row not found.' };
    const h = getHeaders(sheet);
    if (studentId) {
      const rowId = normalizeStudentId(sheet.getRange(r, h['Student ID'] + 1).getValue());
      if (rowId !== normalizeStudentId(studentId)) return { error: 'Row mismatch.' };
    }
    sheet.getRange(r, h['Score'] + 1).setValue(Number(score));
    if (maxGrade !== undefined && maxGrade !== '' && h['Max Grade'] !== undefined) {
      sheet.getRange(r, h['Max Grade'] + 1).setValue(Number(maxGrade));
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ───────────────────────────────────────────────────────────
//  WRITE: Delete a specific quiz submission (admin)
// ───────────────────────────────────────────────────────────
function deleteSubmission(body) {
  const { rowNum, studentId } = body;
  if (!rowNum) return { error: 'rowNum is required.' };

  const sheet = getSheet('Submissions');
  if (!sheet || sheet.getLastRow() <= 1) return { error: 'Submissions sheet empty.' };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { return { error: 'Server busy.' }; }

  try {
    const r = Number(rowNum);
    if (r < 2 || r > sheet.getLastRow()) return { error: 'Submission row not found.' };
    const h = getHeaders(sheet);
    if (studentId) {
      const rowId = normalizeStudentId(sheet.getRange(r, h['Student ID'] + 1).getValue());
      if (rowId !== normalizeStudentId(studentId)) return { error: 'Row mismatch.' };
    }
    sheet.deleteRow(r);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ───────────────────────────────────────────────────────────
//  WRITE: Delete student (admin)
// ───────────────────────────────────────────────────────────
function deleteStudent(body) {
  const { studentId } = body;
  if (!studentId) return { error: 'Student ID required.' };
  const normId = normalizeStudentId(studentId);

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { return { error: 'Server busy.' }; }

  try {
    let totalDeleted = 0;

    const deleteMatchingRows = (sheetName, colHeader) => {
      const sheet = getSheet(sheetName);
      if (!sheet || sheet.getLastRow() <= 1) return 0;
      const data = sheet.getDataRange().getValues();
      const h    = getHeaders(sheet);
      const colIdx = h[colHeader];
      if (colIdx === undefined) return 0;
      let count = 0;
      for (let i = data.length - 1; i >= 1; i--) {
        if (normalizeStudentId(data[i][colIdx]) === normId) {
          sheet.deleteRow(i + 1);
          count++;
        }
      }
      return count;
    };

    totalDeleted += deleteMatchingRows('Students', 'ID');
    totalDeleted += deleteMatchingRows('Grades', 'Student ID');
    totalDeleted += deleteMatchingRows('Submissions', 'Student ID');
    totalDeleted += deleteMatchingRows('Retakes', 'Target');

    if (totalDeleted === 0) {
      return { error: `Student ID ${studentId} not found.` };
    }
    return { ok: true, deletedRows: totalDeleted };
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

// ───────────────────────────────────────────────────────────
//  WRITE: Grant 1x retake for ALL or a specific student (admin)
// ───────────────────────────────────────────────────────────
function getOrCreateRetakesSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Retakes');
  if (!sheet) {
    sheet = ss.insertSheet('Retakes');
    sheet.appendRow(['Target', 'Exam ID', 'Exam Label', 'Granted At', 'Updated At']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold')
      .setBackground('#003865').setFontColor('#ffffff');
  }
  return sheet;
}

function grantRetake(body) {
  const target    = String(body.target || '').trim();
  const examId    = String(body.examId || '').trim();
  const examLabel = String(body.examLabel || '').trim();
  if (!target || !examId) return { error: 'Target and Exam are required.' };

  const sheet = getOrCreateRetakesSheet();
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { return { error: 'Server busy.' }; }

  try {
    const grantedAt = Date.now();
    const now = new Date();
    if (sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      const h    = getHeaders(sheet);
      const normTarget = target === 'ALL' ? 'ALL' : normalizeStudentId(target);
      for (let i = 1; i < data.length; i++) {
        const rowTarget = String(data[i][h['Target']] || '').trim();
        const normRowTarget = rowTarget === 'ALL' ? 'ALL' : normalizeStudentId(rowTarget);
        const rowExam = String(data[i][h['Exam ID']] || '').trim();
        if (normRowTarget === normTarget && rowExam === examId) {
          sheet.getRange(i + 1, h['Exam Label'] + 1).setValue(examLabel);
          sheet.getRange(i + 1, h['Granted At'] + 1).setValue(grantedAt);
          sheet.getRange(i + 1, h['Updated At'] + 1).setValue(now);
          return { ok: true, grantedAt };
        }
      }
    }
    sheet.appendRow([target, examId, examLabel, grantedAt, now]);
    return { ok: true, grantedAt };
  } finally {
    lock.releaseLock();
  }
}
