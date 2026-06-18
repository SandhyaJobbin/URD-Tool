/**
 * URD Simulator — Google Sheets Backend (v2 with Day 6 upload support)
 *
 * Routes:
 *   POST (default body)              → save a score record to the "Scores" tab
 *   POST {type:'upload_day6', ...}   → create a new tab "Day6_<timestamp>"
 *                                       and store the uploaded questions
 *   GET (no params)                  → return all score records as JSON
 *   GET ?action=get_day6             → return the latest uploaded Day 6 questions
 *
 * Setup steps are in Google_Sheets_Setup_Guide.md
 */

const SCORES_SHEET = 'Scores';
const DAY6_TAB_PREFIX = 'Day6_';
const ACTIVE_DAY6_PROP = 'ACTIVE_DAY6_TAB';

const SCORE_HEADERS = [
  'datetime', 'empId', 'name', 'day', 'dayName',
  'score', 'pct', 'correct', 'total', 'result', 'decisions'
];

const DAY6_HEADERS = [
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
    if (body && body.type === 'upload_day6') {
      return handleDay6Upload_(body);
    }
    return handleScoreSave_(body);
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || '';
    if (action === 'get_day6') {
      return handleDay6Fetch_();
    }
    if (action === 'day6_status') {
      return handleDay6Status_();
    }
    return handleScoresFetch_();
  } catch (err) {
    return jsonResponse_({ error: String(err) });
  }
}

/* ════════════════════════════════════════
   SCORES (existing behavior)
   ════════════════════════════════════════ */

function handleScoreSave_(body) {
  const sheet = getOrCreateScoresSheet_();
  sheet.appendRow([
    body.datetime || new Date().toLocaleString(),
    body.empId || '',
    body.name || '',
    body.day || '',
    body.dayName || '',
    Number(body.score) || 0,
    Number(body.pct) || 0,
    Number(body.correct) || 0,
    Number(body.total) || 0,
    (Number(body.pct) >= 80) ? 'PASSED' : 'FAILED',
    JSON.stringify(body.decisions || [])
  ]);
  return jsonResponse_({ ok: true });
}

function handleScoresFetch_() {
  const sheet = getOrCreateScoresSheet_();
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

/* ════════════════════════════════════════
   DAY 6 UPLOAD
   ════════════════════════════════════════ */

function handleDay6Upload_(body) {
  const questions = (body && body.questions) || [];
  if (!Array.isArray(questions) || questions.length === 0) {
    return jsonResponse_({ ok: false, error: 'No questions provided in payload.' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
  const tabName = DAY6_TAB_PREFIX + stamp;

  const sheet = ss.insertSheet(tabName);
  sheet.appendRow(DAY6_HEADERS);
  sheet.getRange(1, 1, 1, DAY6_HEADERS.length)
    .setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  const rows = questions.map(function (q, i) {
    return [
      6000 + i + 1,
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
    sheet.getRange(2, 1, rows.length, DAY6_HEADERS.length).setValues(rows);
  }

  // Set sensible column widths
  sheet.setColumnWidth(1, 50);    // id
  sheet.setColumnWidth(2, 250);   // title
  sheet.setColumnWidth(3, 400);   // body
  sheet.setColumnWidth(4, 130);   // correct_action
  sheet.setColumnWidth(5, 180);   // correct_reason
  sheet.setColumnWidth(6, 300);   // explanation_correct
  sheet.setColumnWidth(7, 300);   // explanation_wrong
  sheet.setColumnWidth(8, 60);    // rating
  sheet.setColumnWidth(9, 180);   // property_name
  sheet.setColumnWidth(10, 130);  // experience_type
  sheet.setColumnWidth(11, 100);  // date
  sheet.setColumnWidth(12, 180);  // policy_name
  sheet.setColumnWidth(13, 300);  // owner_response

  // Wrap text in body / explanation / owner_response columns
  sheet.getRange(2, 3, rows.length, 1).setWrap(true);
  sheet.getRange(2, 6, rows.length, 2).setWrap(true);
  sheet.getRange(2, 13, rows.length, 1).setWrap(true);

  // Mark this tab as the active Day 6 source
  PropertiesService.getScriptProperties().setProperty(ACTIVE_DAY6_PROP, tabName);

  return jsonResponse_({
    ok: true,
    tab: tabName,
    count: questions.length,
    timestamp: stamp
  });
}

function handleDay6Fetch_() {
  const tabName = PropertiesService.getScriptProperties().getProperty(ACTIVE_DAY6_PROP);
  if (!tabName) return jsonResponse_([]);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return jsonResponse_([]);

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse_([]);

  const headers = data[0];
  const questions = data.slice(1).map(function (row) {
    const q = {};
    headers.forEach(function (h, i) { q[h] = row[i]; });
    return q;
  }).filter(function (q) {
    return q.title && q.body && q.correct_action;
  });

  return jsonResponse_(questions);
}

function handleDay6Status_() {
  const tabName = PropertiesService.getScriptProperties().getProperty(ACTIVE_DAY6_PROP);
  if (!tabName) return jsonResponse_({ active: false, count: 0 });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return jsonResponse_({ active: false, count: 0 });

  const count = Math.max(0, sheet.getLastRow() - 1);
  return jsonResponse_({
    active: true,
    tab: tabName,
    count: count
  });
}

/* ════════════════════════════════════════
   HELPERS
   ════════════════════════════════════════ */

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
