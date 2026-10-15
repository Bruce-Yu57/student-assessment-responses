const DESTINATION_SHEET_NAME = 'ReDraft';
const DESTINATION_SPREADSHEET_ID = '1BDrF1havIeF5AFNcW9KRQY4gWdBlD5Yd2hjkr98mM84';
const EXPECTED_HEADERS = ['學生班級座號', 'Exam_ID', 'Question_No', '對錯紀錄'];
const RECORD_HEADERS = ['學生班級座號', 'Exam_ID', 'Question_Type', 'Question_No', '對錯紀錄'];
const STUDENT_SOURCE_ID = '1TctKOS8qWOidvWLeuSwiQBw0BK4kJ9wGySt9vuTpygo';
const STUDENT_SOURCE_GID = 0;
const QUESTION_SOURCE_ID = '14eJB4OkA1PEeujdAU0T96q2xd_o9XKBAKAhJ3f-cMMA';
const QUESTION_SOURCE_GID = 2028107386; // draft
const RECORD_SHEET_NAME = 'Record';
const TEACHER_CODE_PROPERTY = 'TEACHER_ACCESS_CODE';

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

function questionTypeFromId_(questionId) {
  const prefix = String(questionId || '').match(/^([A-Za-z]+)/)?.[1].toUpperCase() || 'OTHER';
  return ({Q:'選擇題', F:'填充題', W:'計算題'})[prefix] || `${prefix}類`;
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

function getDestinationSpreadsheet_() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return activeSpreadsheet && activeSpreadsheet.getId() === DESTINATION_SPREADSHEET_ID
    ? activeSpreadsheet
    : SpreadsheetApp.openById(DESTINATION_SPREADSHEET_ID);
}

function teacherCodeMatches_(code) {
  const savedCode = PropertiesService.getScriptProperties().getProperty(TEACHER_CODE_PROPERTY);
  return Boolean(savedCode && String(code || '').trim() === savedCode);
}

function jsonp_(callback, result) {
  if (!/^[A-Za-z_$][A-Za-z0-9_$]{0,80}$/.test(String(callback || ''))) {
    return ContentService.createTextOutput('Invalid callback.');
  }
  return ContentService.createTextOutput(`${callback}(${JSON.stringify(result)});`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function readDraftForTeacher_(code) {
  if (!teacherCodeMatches_(code)) return {ok:false, error:'辨識碼不正確或尚未設定。'};
  const sheet = getDestinationSpreadsheet_().getSheetByName(DESTINATION_SHEET_NAME);
  if (!sheet) return {ok:false, error:'找不到 ReDraft 工作表。'};
  const headers = sheet.getRange(1, 1, 1, EXPECTED_HEADERS.length).getDisplayValues()[0];
  if (EXPECTED_HEADERS.some((header, index) => headers[index] !== header)) {
    return {ok:false, error:'ReDraft 欄位標題不符合預期。'};
  }
  const lastRow = sheet.getLastRow();
  const rows = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, EXPECTED_HEADERS.length).getDisplayValues()
      .filter(row => row.some(value => String(value).trim() !== ''))
    : [];
  return {ok:true, rows};
}

function promoteTeacherDraft_(payload) {
  if (!teacherCodeMatches_(payload.code)) throw new Error('辨識碼不正確或尚未設定。');
  if (payload.confirmed !== true) throw new Error('Teacher confirmation is required.');
  if (!Array.isArray(payload.rows) || payload.rows.length < 1 || payload.rows.length > 5000) {
    throw new Error('Invalid draft rows.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const spreadsheet = getDestinationSpreadsheet_();
    const draft = spreadsheet.getSheetByName(DESTINATION_SHEET_NAME);
    if (!draft) throw new Error('ReDraft was not found.');
    const draftHeaders = draft.getRange(1, 1, 1, EXPECTED_HEADERS.length).getDisplayValues()[0];
    if (EXPECTED_HEADERS.some((header, index) => draftHeaders[index] !== header)) throw new Error('ReDraft headers do not match.');

    const lastDraftRow = draft.getLastRow();
    const currentRows = lastDraftRow > 1
      ? draft.getRange(2, 1, lastDraftRow - 1, EXPECTED_HEADERS.length).getDisplayValues()
        .filter(row => row.some(value => String(value).trim() !== ''))
      : [];
    if (currentRows.length !== payload.rows.length) throw new Error('ReDraft changed while it was being reviewed. Reload it and try again.');

    const studentSheet = getSheetById(STUDENT_SOURCE_ID, STUDENT_SOURCE_GID);
    const studentLastRow = studentSheet.getLastRow();
    const validStudents = new Set(studentLastRow > 0
      ? studentSheet.getRange(1, 1, studentLastRow, 1).getDisplayValues().flat().map(value => value.trim())
      : []);
    const validQuestions = getActiveQuestionKeys();
    const seen = new Set();
    const rows = payload.rows.map((item, index) => {
      const original = Array.isArray(item.original) ? item.original.map(value => String(value ?? '')) : [];
      if (original.length !== EXPECTED_HEADERS.length || original.some((value, column) => value !== currentRows[index][column])) {
        throw new Error('ReDraft changed while it was being reviewed. Reload it and try again.');
      }
      const studentCode = String(item.studentCode || '').trim();
      const quiz = String(item.quiz || '').trim();
      const question = String(item.question || '').trim();
      const value = Number(item.value);
      if (!/^\d{4,6}$/.test(studentCode) || !validStudents.has(studentCode)) throw new Error(`第 ${index + 1} 筆學生班級座號無效。`);
      if (!quiz || !question || !validQuestions.has(`${quiz}|${question}`)) throw new Error(`第 ${index + 1} 筆考卷或題號目前不存在。`);
      if (item.value !== '0' && item.value !== '1') throw new Error(`第 ${index + 1} 筆對錯值只能是 0 或 1。`);
      const key = `${studentCode}|${quiz}|${question}`;
      if (seen.has(key)) throw new Error('修正後有重複的學生、考卷與題號。');
      seen.add(key);
      return [studentCode, quiz, question, value];
    });

    let record = spreadsheet.getSheetByName(RECORD_SHEET_NAME);
    if (!record) {
      record = spreadsheet.insertSheet(RECORD_SHEET_NAME);
      record.getRange(1, 1, 1, RECORD_HEADERS.length).setValues([RECORD_HEADERS]);
    }
    const recordHeaders = record.getRange(1, 1, 1, RECORD_HEADERS.length).getDisplayValues()[0];
    if (RECORD_HEADERS.some((header, index) => recordHeaders[index] !== header)) {
      throw new Error('Record 欄位標題須為：學生班級座號、Exam_ID、Question_Type、Question_No、對錯紀錄。');
    }

    const lastRecordRow = record.getLastRow();
    const existing = lastRecordRow > 1
      ? record.getRange(2, 1, lastRecordRow - 1, RECORD_HEADERS.length).getDisplayValues()
      : [];
    const rowByKey = new Map();
    existing.forEach((row, index) => rowByKey.set(`${row[0]}|${row[1]}|${row[2]}|${row[3]}`, index + 2));
    const appendRows = [];
    const rowsToVerify = new Map();
    rows.forEach(draftRow => {
      const row = [draftRow[0], draftRow[1], questionTypeFromId_(draftRow[2]), draftRow[2], draftRow[3]];
      const key = `${row[0]}|${row[1]}|${row[2]}|${row[3]}`;
      const targetRow = rowByKey.get(key);
      if (targetRow) record.getRange(targetRow, 1, 1, row.length).setValues([row]);
      else {
        appendRows.push(row);
        rowByKey.set(key, lastRecordRow + appendRows.length);
      }
      rowsToVerify.set(key, row);
    });
    if (appendRows.length) {
      record.getRange(record.getLastRow() + 1, 1, appendRows.length, RECORD_HEADERS.length).setValues(appendRows);
    }
    SpreadsheetApp.flush();

    for (const [key, expected] of rowsToVerify) {
      const recordRow = rowByKey.get(key);
      const actual = record.getRange(recordRow, 1, 1, RECORD_HEADERS.length).getDisplayValues()[0];
      if (actual.some((value, column) => value !== String(expected[column]))) {
        throw new Error('Record 寫入確認失敗；ReDraft 已保留。');
      }
    }

    if (lastDraftRow > 1) draft.getRange(2, 1, lastDraftRow - 1, EXPECTED_HEADERS.length).clearContent();
    SpreadsheetApp.flush();
    return {ok:true, saved:rows.length, cleared:true};
  } finally {
    lock.releaseLock();
  }
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
    if (payload.action === 'promote') {
      const result = promoteTeacherDraft_(payload);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
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
      const spreadsheet = getDestinationSpreadsheet_();
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

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.action === 'verifyTeacher') {
    const configured = Boolean(PropertiesService.getScriptProperties().getProperty(TEACHER_CODE_PROPERTY));
    const ok = configured && teacherCodeMatches_(params.code);
    return jsonp_(params.callback, {ok, error:configured ? '辨識碼不正確。' : '尚未設定教師辨識碼。'});
  }
  if (params.action === 'teacherDraft') return jsonp_(params.callback, readDraftForTeacher_(params.code));
  return ContentService.createTextOutput('Assessment ReDraft and teacher endpoint is ready.');
}
