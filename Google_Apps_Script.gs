/**
 * URD Simulator — Google Sheets Backend (v3 with Batches)
 *
 * Tabs created/used:
 *   Scores                — one row per assessment attempt
 *                            columns: datetime, empId, name, country, batch,
 *                                     day, dayName, score, pct, correct,
 *                                     total, result, decisions
 *   Batches               — one row per trainer-created batch
 *                            columns: batchId, batchName, startDate, endDate,
 *                                     createdAt, fileName, questionCount,
 *                                     activeTab, createdBy
 *   Day6_<timestamp>      — one per Day 6 CSV upload (legacy global)
 *   Batch_<batchId>       — one per batch upload, questions for that batch
 *
 * Routes:
 *   POST (no type)                  → save score (auto-migrates columns)
 *   POST {type:'upload_day6'}       → create Day6_<ts> tab and set active
 *   POST {type:'create_batch'}      → create Batch_<id> tab + Batches row
 *   GET                             → return all score records as JSON
 *   GET ?action=get_day6            → return latest Day 6 questions
 *   GET ?action=day6_status         → return Day 6 active info
 *   GET ?action=list_batches        → return all batches (for dropdown)
 *   GET ?action=get_batch&id=ID     → return questions for a specific batch
 */

const SCORES_SHEET = 'Scores';
const BATCHES_SHEET = 'Batches';
const DAY6_TAB_PREFIX = 'Day6_';
const BATCH_TAB_PREFIX = 'Batch_';
const ACTIVE_DAY6_PROP = 'ACTIVE_DAY6_TAB';

// Master Scores headers. Order is preserved when writing rows; missing
// columns in existing sheets are auto-added at the end.
const SCORE_HEADERS = [
  'datetime', 'empId', 'name', 'country', 'batch',
  'day', 'dayName', 'score', 'pct', 'correct',
  'total', 'result', 'decisions'
];

const BATCH_HEADERS = [
  'batchId', 'batchName', 'startDate', 'endDate', 'createdAt',
  'fileName', 'questionCount', 'activeTab', 'createdBy'
];

const QUESTION_HEADERS = [
  'id', 'title', 'body', 'correct_action', 'correct_reason',
  'explanation_correct', 'explanation_wrong', 'rating',
  'property_name', 'experience_type', 'date', 'policy_name', 'owner_response'
];

/* ════════════════════════════════════════
   ROUTING
   ════════════════════════════════════════ */

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body && body.type === 'upload_day6') return handleDay6Upload_(body);
    if (body && body.type === 'create_batch') return handleCreateBatch_(body);
    return handleScoreSave_(body);
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || '';
    if (action === 'get_day6') return handleDay6Fetch_();
    if (action === 'day6_status') return handleDay6Status_();
    if (action === 'list_batches') return handleListBatches_();
    if (action === 'get_batch') return handleGetBatch_(e.parameter.id);
    return handleScoresFetch_();
  } catch (err) {
    return jsonResponse_({ error: String(err) });
  }
}

/* ════════════════════════════════════════
   SCORES (with auto column-migration)
   ════════════════════════════════════════ */

function handleScoreSave_(body) {
  const sheet = getOrCreateScoresSheet_();
  // Make sure all expected columns exist (auto-migrate existing sheets).
  ensureColumns_(sheet, SCORE_HEADERS);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(function (h) {
    switch (h) {
      case 'datetime': return body.datetime || new Date().toLocaleString();
      case 'empId':    return body.empId || '';
      case 'name':     return body.name || '';
      case 'country':  return body.country || '';
      case 'batch':    return body.batch || '';
      case 'day':      return body.day || '';
      case 'dayName':  return body.dayName || '';
      case 'score':    return Number(body.score) || 0;
      case 'pct':      return Number(body.pct) || 0;
      case 'correct':  return Number(body.correct) || 0;
      case 'total':    return Number(body.total) || 0;
      case 'result':   return (Number(body.pct) >= 80) ? 'PASSED' : 'FAILED';
      case 'decisions':return JSON.stringify(body.decisions || []);
      default:         return '';
    }
  });
  sheet.appendRow(row);
  return jsonResponse_({ ok: true });
}

function handleScoresFetch_() {
  const sheet = getOrCreateScoresSheet_();
  ensureColumns_(sheet, SCORE_HEADERS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse_([]);

  const headers = data[0];
  const rows = data.slice(1).map(function (row) {
    const obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    if (obj.decisions && typeof obj.decisions === 'string') {
      try { obj.decisions = JSON.parse(obj.decisions); }
      catch (parseErr) { obj.decisions = []; }
    }
    obj.day = Number(obj.day) || obj.day;
    obj.score = Number(obj.score) || 0;
    obj.pct = Number(obj.pct) || 0;
    obj.correct = Number(obj.correct) || 0;
    obj.total = Number(obj.total) || 0;
    return obj;
  });
  return jsonResponse_(rows);
}

function getOrCreateScoresSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SCORES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SCORES_SHEET);
    sheet.appendRow(SCORE_HEADERS);
    sheet.getRange(1, 1, 1, SCORE_HEADERS.length)
      .setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, SCORE_HEADERS.length, 140);
  }
  return sheet;
}

function ensureColumns_(sheet, expectedHeaders) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(expectedHeaders);
    sheet.getRange(1, 1, 1, expectedHeaders.length)
      .setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    return;
  }
  const lastCol = sheet.getLastColumn();
  const existing = sheet.getRange(1, 1, 1, Math.max(1, lastCol)).getValues()[0];
  expectedHeaders.forEach(function (h) {
    if (existing.indexOf(h) === -1) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(h);
      sheet.getRange(1, newCol)
        .setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
      existing.push(h);
    }
  });
}

/* ════════════════════════════════════════
   DAY 6 UPLOAD (legacy, preserved)
   ════════════════════════════════════════ */

function handleDay6Upload_(body) {
  return createQuestionTab_(body.questions || [], DAY6_TAB_PREFIX, function (tabName) {
    PropertiesService.getScriptProperties().setProperty(ACTIVE_DAY6_PROP, tabName);
  });
}

function handleDay6Fetch_() {
  const tabName = PropertiesService.getScriptProperties().getProperty(ACTIVE_DAY6_PROP);
  if (!tabName) return jsonResponse_([]);
  return jsonResponse_(readQuestionTab_(tabName));
}

function handleDay6Status_() {
  const tabName = PropertiesService.getScriptProperties().getProperty(ACTIVE_DAY6_PROP);
  if (!tabName) return jsonResponse_({ active: false, count: 0 });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return jsonResponse_({ active: false, count: 0 });
  return jsonResponse_({
    active: true, tab: tabName,
    count: Math.max(0, sheet.getLastRow() - 1)
  });
}

/* ════════════════════════════════════════
   BATCH MANAGEMENT
   ════════════════════════════════════════ */

function handleCreateBatch_(body) {
  const questions = (body && body.questions) || [];
  if (!Array.isArray(questions) || questions.length === 0) {
    return jsonResponse_({ ok: false, error: 'No questions provided.' });
  }
  const batchName = String(body.batchName || '').trim();
  if (!batchName) return jsonResponse_({ ok: false, error: 'Batch name is required.' });

  const batchId = 'BATCH-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const tabName = BATCH_TAB_PREFIX + batchId;

  // Create the questions tab
  const result = createQuestionTab_(questions, BATCH_TAB_PREFIX, null, batchId);
  if (!result || !result.tab) {
    return jsonResponse_({ ok: false, error: 'Failed to create batch tab.' });
  }

  // Append to Batches sheet
  const batchesSheet = getOrCreateBatchesSheet_();
  batchesSheet.appendRow([
    batchId,
    batchName,
    body.startDate || '',
    body.endDate || '',
    new Date().toLocaleString(),
    body.fileName || '',
    questions.length,
    result.tab,
    body.createdBy || 'Admin'
  ]);

  return jsonResponse_({
    ok: true,
    batchId: batchId,
    tab: result.tab,
    count: questions.length
  });
}

function handleListBatches_() {
  const sheet = getOrCreateBatchesSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse_([]);
  const headers = data[0];
  const rows = data.slice(1).map(function (row) {
    const obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    obj.questionCount = Number(obj.questionCount) || 0;
    if (obj.startDate instanceof Date) {
      obj.startDate = Utilities.formatDate(obj.startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    if (obj.endDate instanceof Date) {
      obj.endDate = Utilities.formatDate(obj.endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    return obj;
  }).filter(function (b) { return b.batchId; });
  return jsonResponse_(rows);
}

function handleGetBatch_(batchId) {
  if (!batchId) return jsonResponse_([]);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Look up the active tab for this batch in the Batches sheet
  const batchesSheet = ss.getSheetByName(BATCHES_SHEET);
  if (!batchesSheet) return jsonResponse_([]);
  const data = batchesSheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse_([]);
  const headers = data[0];
  const idIdx = headers.indexOf('batchId');
  const tabIdx = headers.indexOf('activeTab');
  let tabName = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === batchId) {
      tabName = data[i][tabIdx];
      break;
    }
  }
  if (!tabName) return jsonResponse_([]);
  return jsonResponse_(readQuestionTab_(tabName));
}

function getOrCreateBatchesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BATCHES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(BATCHES_SHEET);
    sheet.appendRow(BATCH_HEADERS);
    sheet.getRange(1, 1, 1, BATCH_HEADERS.length)
      .setFontWeight('bold').setBackground('#00AF87').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidths(3, 7, 130);
  }
  return sheet;
}

/* ════════════════════════════════════════
   SHARED QUESTION TAB HELPERS
   ════════════════════════════════════════ */

function createQuestionTab_(questions, prefix, afterCreate, customId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stamp = customId || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
  const tabName = prefix + stamp;
  const sheet = ss.insertSheet(tabName);
  sheet.appendRow(QUESTION_HEADERS);
  sheet.getRange(1, 1, 1, QUESTION_HEADERS.length)
    .setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  const rows = questions.map(function (q, i) {
    return [
      (prefix === BATCH_TAB_PREFIX ? 7000 : 6000) + i + 1,
      q.title || '',
      q.body || '',
      q.correct_action || '',
      q.correct_reason || '',
      q.explanation_correct || '',
      q.explanation_wrong || '',
      Number(q.rating) || 3,
      q.property_name || '',
      q.experience_type || '',
      q.date || '',
      q.policy_name || '',
      q.owner_response || ''
    ];
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, QUESTION_HEADERS.length).setValues(rows);
  }
  // Column widths
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 400);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 180);
  sheet.setColumnWidth(6, 300);
  sheet.setColumnWidth(7, 300);
  sheet.setColumnWidth(8, 60);
  sheet.setColumnWidth(9, 180);
  sheet.setColumnWidth(10, 130);
  sheet.setColumnWidth(11, 100);
  sheet.setColumnWidth(12, 180);
  sheet.setColumnWidth(13, 300);

  if (afterCreate) afterCreate(tabName);
  return { ok: true, tab: tabName, count: questions.length };
}

function readQuestionTab_(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(function (row) {
    const q = {};
    headers.forEach(function (h, i) { q[h] = row[i]; });
    return q;
  }).filter(function (q) {
    return q.title && q.body && q.correct_action;
  });
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
