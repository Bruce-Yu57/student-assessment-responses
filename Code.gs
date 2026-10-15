const DESTINATION_SHEET_NAME = 'ReDraft';
const DESTINATION_SPREADSHEET_ID = '1BDrF1havIeF5AFNcW9KRQY4gWdBlD5Yd2hjkr98mM84';
const EXPECTED_HEADERS = ['學生班級座號', 'Exam_ID', 'Question_No', '對錯紀錄'];
const STUDENT_SOURCE_ID = '1TctKOS8qWOidvWLeuSwiQBw0BK4kJ9wGySt9vuTpygo';
const STUDENT_SOURCE_GID = 0;
const QUESTION_SOURCE_ID = '14eJB4OkA1PEeujdAU0T96q2xd_o9XKBAKAhJ3f-cMMA';
const QUESTION_SOURCE_GID = 2028107386; // draft

function getSheetById(spreadsheetId, gid) {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheets().find(candidate => candidate.getSheetId() === gid);
  if (!sheet) throw new Error('Source sheet was not found.');
  return sheet;
}

function getColumnIndex(headers, name) {
  const index = headers.indexOf(name);
  if (index < 0) throw new Error(`Required source column is missing: ${name}`);
  return index;
}

function getActiveQuestionKeys() {
  const sheet = getSheetById(QUESTION_SOURCE_ID, QUESTION_SOURCE_GID);
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) throw new Error('No quiz questions were found.');
  const headers = values[0].map(value => value.trim());
  const quizIndex = getColumnIndex(headers, 'Exam_ID');
  const questionIndex = getColumnIndex(headers, 'Question_No');
  const statusIndex = getColumnIndex(headers, '狀態');
  const keys = new Set();
  values.slice(1).forEach(row => {
    const quiz = String(row[quizIndex] || '').trim();
    const question = String(row[questionIndex] || '').trim();
    const status = String(row[statusIndex] || '').trim().toUpperCase();
    if (quiz && question && status === 'A') keys.add(`${quiz}|${question}`);
  });
  return keys;
}

function validateSourceChoices(studentCode, quiz, answers) {
  const studentSheet = getSheetById(STUDENT_SOURCE_ID, STUDENT_SOURCE_GID);
  const studentLastRow = studentSheet.getLastRow();
  const studentCodes = studentLastRow > 0
    ? studentSheet.getRange(1, 1, studentLastRow, 1).getDisplayValues().flat().map(value => value.trim())
    : [];
  if (!studentCodes.includes(studentCode)) throw new Error('Student code is not in the current roster.');

  const activeQuestions = getActiveQuestionKeys();
  answers.forEach(answer => {
    if (!activeQuestions.has(`${quiz}|${answer.question}`)) {
      throw new Error('Quiz or question is not currently active in draft.');
    }
  });
}

/**
 * Deploy as a web app that runs as the owner and is accessible to anyone.
 * Saves/upserts one row per student, exam, and question into ReDraft.
 */
function doPost(e) {
  try {
    const raw = e && e.parameter && e.parameter.payload;
    if (!raw || raw.length > 50000) throw new Error('Missing or oversized payload.');
    const payload = JSON.parse(raw);
    const studentCode = String(payload.studentCode || '').trim();
    const quiz = String(payload.quiz || '').trim();
    const answers = payload.answers;

    if (!/^\d{4,6}$/.test(studentCode)) throw new Error('Invalid student code.');
    if (!quiz || quiz.length > 100) throw new Error('Invalid Exam_ID.');
    if (!Array.isArray(answers) || answers.length < 1 || answers.length > 300) throw new Error('Invalid answer count.');

    const seen = new Set();
    const rows = answers.map(answer => {
      const question = String(answer.question || '').trim();
      const value = Number(answer.value);
      if (!question || question.length > 80 || (value !== 0 && value !== 1)) throw new Error('Invalid answer data.');
      if (seen.has(question)) throw new Error('Duplicate question.');
      seen.add(question);
      return [studentCode, quiz, question, value];
    });

    validateSourceChoices(studentCode, quiz, rows.map(row => ({question: row[2]})));

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const spreadsheet = activeSpreadsheet && activeSpreadsheet.getId() === DESTINATION_SPREADSHEET_ID
        ? activeSpreadsheet
        : SpreadsheetApp.openById(DESTINATION_SPREADSHEET_ID);
      const sheet = spreadsheet.getSheetByName(DESTINATION_SHEET_NAME);
      if (!sheet) throw new Error('Destination sheet ReDraft was not found.');
      const headers = sheet.getRange(1, 1, 1, EXPECTED_HEADERS.length).getDisplayValues()[0];
      if (EXPECTED_HEADERS.some((header, index) => headers[index] !== header)) {
        throw new Error('Destination headers do not match: 學生班級座號, Exam_ID, Question_No, 對錯紀錄.');
      }

      const lastRow = sheet.getLastRow();
      const existing = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, EXPECTED_HEADERS.length).getDisplayValues() : [];
      const rowByKey = new Map();
      existing.forEach((row, index) => rowByKey.set(`${row[0]}|${row[1]}|${row[2]}`, index + 2));

      const appendRows = [];
      rows.forEach(row => {
        const key = `${row[0]}|${row[1]}|${row[2]}`;
        const targetRow = rowByKey.get(key);
        if (targetRow) sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
        else appendRows.push(row);
      });
      if (appendRows.length) {
        sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, EXPECTED_HEADERS.length).setValues(appendRows);
      }
    } finally {
      lock.releaseLock();
    }

    return ContentService.createTextOutput(JSON.stringify({ok: true, saved: rows.length}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: String(error.message || error)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('Assessment ReDraft endpoint is ready.');
}
