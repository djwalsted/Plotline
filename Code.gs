/**
 * Plotline Timnath: saves survey answers into this Google Sheet.
 * Paste this whole file into Code.gs in Extensions > Apps Script, then Deploy as a Web app.
 * Tabs "Answers" and "Feedback" are created automatically on first use.
 */

var ANSWER_SHEET = 'Answers';
var FEEDBACK_SHEET = 'Feedback';
var HEADERS = [
  'Submitted', 'Building', 'What they want', 'Write-in',
  'Idea score (1-5)', 'Location score (1-5)', 'How often', 'When',
  'How they would get there', 'Spend per visit', 'Must-haves', 'Comment',
  'Town', 'Distance from Main St', 'Gets around by', 'Trail use', 'Kids at home',
  'Respondent ID', 'App data (do not edit)'
];
var BUILDINGS = ['a', 'b', 'c', 'd'];

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// The website (GitHub Pages) reads answers here: .../exec?action=answers
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === 'answers') return json_({ ok: true, result: getAnswers() });
    return json_({ ok: true, result: 'Plotline Timnath backend is running.' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

// The website sends answers and feedback here.
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'submitAnswer') return json_({ ok: true, result: submitAnswer(body.payload) });
    if (body.action === 'submitFeedback') return json_({ ok: true, result: submitFeedback(body.payload) });
    if (body.action === 'joinPanel') return json_({ ok: true, result: joinPanel(body.payload) });
    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function sheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}

// Keep text safe for a spreadsheet: limit length, block formulas.
function clean_(v, max) {
  var s = String(v == null ? '' : v).slice(0, max || 500);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function getAnswers() {
  var sh = sheet_(ANSWER_SHEET, HEADERS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var n = HEADERS.length;
  var vals = sh.getRange(2, n - 1, last - 1, 2).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] && vals[i][1]) out.push({ rid: String(vals[i][0]), data: String(vals[i][1]) });
  }
  return out;
}

function submitAnswer(p) {
  if (!p || typeof p.rid !== 'string' || !/^[a-z0-9]{6,40}$/.test(p.rid)) throw new Error('Bad respondent id');
  var r = p.row || {};
  if (BUILDINGS.indexOf(r.building) < 0) throw new Error('Unknown building');
  if (typeof p.data !== 'string' || p.data.length > 8000) throw new Error('Bad data');
  JSON.parse(p.data); // must be valid JSON

  var values = [
    new Date(), clean_(r.address, 120), clean_(r.picks, 300), clean_(r.writein, 120),
    clean_(r.concept, 2), clean_(r.location, 2), clean_(r.freq, 40), clean_(r.when, 200),
    clean_(r.how, 40), clean_(r.spend, 40), clean_(r.extras, 300), clean_(r.comment, 600),
    clean_(r.town, 40), clean_(r.band, 40), clean_(r.modes, 120), clean_(r.trail, 40), clean_(r.kids, 40),
    p.rid, p.data
  ];

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(ANSWER_SHEET, HEADERS);
    var last = sh.getLastRow();
    var n = HEADERS.length;
    var target = 0;
    if (last >= 2) {
      // One row per person per building: update it if they answer again.
      var keys = sh.getRange(2, n - 1, last - 1, 2).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (String(keys[i][0]) === p.rid) {
          try { if (JSON.parse(keys[i][1]).building === r.building) { target = i + 2; break; } } catch (e) {}
        }
      }
    }
    if (target) sh.getRange(target, 1, 1, n).setValues([values]);
    else sh.appendRow(values);
  } finally {
    lock.releaseLock();
  }
  return getAnswers();
}

function submitFeedback(p) {
  if (!p || typeof p.text !== 'string' || !p.text.trim()) throw new Error('Empty feedback');
  var sh = sheet_(FEEDBACK_SHEET, ['Submitted', 'Feedback', 'Respondent ID']);
  sh.appendRow([new Date(), clean_(p.text, 1500), clean_(p.rid, 40)]);
  return true;
}

// Resident panel: people who want to be notified about new projects.
// Kept on its own tab and never linked to survey answers.
var PANEL_SHEET = 'Panel';
var PANEL_HEADERS = ['Joined', 'Email', 'Mobile', 'Text consent', 'Consent wording', 'Town', 'ZIP', 'Nearest intersection', 'Signed up from', 'Status'];
var SMS_CONSENT = 'Text me about new surveys near me. Up to 4 texts a month. Message and data rates may apply. Reply STOP to opt out.';

function joinPanel(p) {
  if (!p) throw new Error('Missing data');
  var email = String(p.email || '').trim().toLowerCase().slice(0, 120);
  var phone = String(p.phone || '').replace(/[^0-9+]/g, '').slice(0, 16);
  if (!email && !phone) throw new Error('Email or phone required');
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Bad email');
  var sms = p.sms === true && !!phone;
  var row = [new Date(), clean_(email, 120), clean_(phone, 16), sms ? 'Yes' : 'No', sms ? SMS_CONSENT : '',
    clean_(p.town, 40), clean_(p.zip, 5), clean_(p.near, 80), clean_(p.source, 40), 'Active'];
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(PANEL_SHEET, PANEL_HEADERS);
    var last = sh.getLastRow();
    if (last >= 2) {
      var keys = sh.getRange(2, 2, last - 1, 2).getValues();
      for (var i = 0; i < keys.length; i++) {
        if ((email && String(keys[i][0]).toLowerCase() === email) || (phone && String(keys[i][1]) === phone)) {
          sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
          return true;
        }
      }
    }
    sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return true;
}
