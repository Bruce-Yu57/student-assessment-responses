const DESTINATION_SPREADSHEET_ID = '1BDrF1havIeF5AFNcW9KRQY4gWdBlD5Yd2hjkr98mM84';
const DESTINATION_SHEET_NAME = '對錯紀錄';
const EXPECTED_HEADERS = ['學生班級座號', '卷號', '題號', '對錯紀錄'];
const STUDENT_SOURCE_ID = '1TctKOS8qWOidvWLeuSwiQBw0BK4kJ9wGySt9vuTpygo';
const STUDENT_SOURCE_GID = 0;
const QUESTION_SOURCE_ID = '14eJB4OkA1PEeujdAU0T96q2xd_o9XKBAKAhJ3f-cMMA';
const QUESTION_SOURCE_GID = 2028107386;

function getSheetById(spreadsheetId, gid) {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheets().find(candidate => candidate.getSheetId() === gid);
  if (!sheet) throw new Error('A source sheet was not found.');
  return sheet;
}

function validateSourceChoices(studentCode, quiz, answers) {
  const studentSheet = getSheetById(STUDENT_SOURCE_ID, STUDENT_SOURCE_GID);
  const studentLastRow = studentSheet.getLastRow();
  const studentCodes = studentLastRow > 0
    ? studentSheet.getRange(1, 1, studentLastRow, 1).getDisplayValues().flat().map(value => value.trim())
    : [];
  if (!studentCodes.includes(studentCode)) throw new Error('Student code is not in the current roster.');

  const questionSheet = getSheetById(QUESTION_SOURCE_ID, QUESTION_SOURCE_GID);
  const questionLastRow = questionSheet.getLastRow();
  if (questionLastRow < 2) throw new Error('No quiz questions were found.');
  const data = questionSheet.getRange(1, 1, questionLastRow, questionSheet.getLastColumn()).getDisplayValues();
  const headings = data[0].map(value => value.trim());
  const quizIndex = headings.findIndex(value => ['卷號', '卷号', '卷'].includes(value));
  const questionIndex = headings.findIndex(value => ['題號', '题号', '題目', '題次'].includes(value));
  if (quizIndex < 0 || questionIndex < 0) throw new Error('Quiz source headers are not recognized.');
  const validQuestions = new Set();
  data.slice(1).forEach(row => {
    const rowQuiz = String(row[quizIndex] || '').trim();
    const match = String(row[questionIndex] || '').match(/\d+/);
    if (rowQuiz && match) validQuestions.add(`${rowQuiz}|${Number(match[0])}`);
  });
  answers.forEach(answer => {
    if (!validQuestions.has(`${quiz}|${Number(answer.question)}`)) throw new Error('Question is not in the selected quiz.');
  });
}

/**
 * Deploy as a Google Apps Script web app that runs as the sheet owner and is
 * accessible to anyone. The public page sends only validated answer rows.
 * Existing student/quiz/question keys are updated so a repeat submit is safe.
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
    if (!/^[A-Za-z0-9_-]{1,40}$/.test(quiz)) throw new Error('Invalid quiz number.');
    if (!Array.isArray(answers) || answers.length < 1 || answers.length > 300) throw new Error('Invalid answer count.');

    const seen = new Set();
    const rows = answers.map(answer => {
      const question = Number(answer.question);
      const value = Number(answer.value);
      if (!Number.isInteger(question) || question < 1 || question > 10000 || (value !== 0 && value !== 1)) {
        throw new Error('Invalid answer data.');
      }
      if (seen.has(question)) throw new Error('Duplicate question.');
      seen.add(question);
      return [studentCode, quiz, question, value];
    });

    validateSourceChoices(studentCode, quiz, rows.map(row => ({question: row[2], value: row[3]})));

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = SpreadsheetApp.openById(DESTINATION_SPREADSHEET_ID).getSheetByName(DESTINATION_SHEET_NAME);
      if (!sheet) throw new Error('Destination sheet was not found.');
      const headers = sheet.getRange(1, 1, 1, EXPECTED_HEADERS.length).getDisplayValues()[0];
      if (EXPECTED_HEADERS.some((header, index) => headers[index] !== header)) {
        throw new Error('Destination headers do not match the expected columns.');
      }

      const lastRow = sheet.getLastRow();
      const existing = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, EXPECTED_HEADERS.length).getValues() : [];
      const rowByKey = new Map();
      existing.forEach((row, index) => {
        rowByKey.set(`${row[0]}|${row[1]}|${row[2]}`, index + 2);
      });

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

    return ContentService.createTextOutput(JSON.stringify({ ok: true, saved: rows.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(error.message || error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('Assessment response endpoint is ready.');
}
