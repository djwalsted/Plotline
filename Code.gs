/**
 * Plotline Timnath: saves survey answers into this Google Sheet.
 * Paste this whole file into Code.gs in Extensions > Apps Script, then Deploy as a Web app.
 * Tabs are created automatically on first use:
 *   Answers   - one row per person per site (private)
 *   Comments  - every comment waits here as Pending. Set Status to Approved to publish it.
 *               You may edit the comment text first (for example, to remove a name).
 *   Sources   - how people heard about Plotline
 *   Alerts    - suspicious activity (bots, sudden bursts of answers)
 *   Feedback  - site feedback, property inquiries and survey requests
 *   Panel     - people who signed up to be notified
 */

// Email alerts go to the Google account that owns this script.
var EMAIL_ALERTS = true;
// More than this many answers for one site within 10 minutes triggers an alert.
var BURST_LIMIT = 25;
// Answers submitted faster than this (milliseconds) after opening a survey are treated as bots.
var MIN_SURVEY_MS = 6000;

var ANSWER_SHEET = 'Answers';
var FEEDBACK_SHEET = 'Feedback';
var HEADERS = [
  'Submitted', 'Building', 'What they want', 'Write-in',
  'Idea score (1-5)', 'Location score (1-5)', 'How often', 'When',
  'How they would get there', 'Spend per visit', 'Must-haves', 'Comment',
  'Town', 'Distance from site', 'Gets around by', 'Trail use', 'Kids at home',
  'Respondent ID', 'App data (do not edit)'
];

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
  var approved = approvedComments_();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    if (!vals[i][0] || !vals[i][1]) continue;
    var rid = String(vals[i][0]), data = String(vals[i][1]);
    // Only publish comments that were approved in the Comments tab (using the approved wording).
    try {
      var d = JSON.parse(data);
      if (d && d.concept && d.concept.comment) {
        var ok = approved[rid + '|' + d.building];
        if (ok) d.concept.comment = ok; else delete d.concept.comment;
        data = JSON.stringify(d);
      }
    } catch (e) {}
    out.push({ rid: rid, data: data });
  }
  return out;
}

// ---------- comment review ----------
var COMMENT_SHEET = 'Comments';
var COMMENT_HEADERS = ['Submitted', 'Site', 'Comment (edit to remove names before approving)', 'Status', 'Respondent ID', 'Site ID', 'Original comment (do not edit)'];
var COMMENT_STATUSES = ['Pending', 'Approved', 'Rejected'];

function commentSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(COMMENT_SHEET);
  if (sh) return sh;
  sh = sheet_(COMMENT_SHEET, COMMENT_HEADERS);
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(COMMENT_STATUSES, true).setAllowInvalid(false).build();
  sh.getRange(2, 4, sh.getMaxRows() - 1, 1).setDataValidation(rule);
  sh.setColumnWidth(3, 420);
  // First run: queue every existing comment for review.
  var ans = sheet_(ANSWER_SHEET, HEADERS), last = ans.getLastRow(), n = HEADERS.length;
  if (last >= 2) {
    var rows = ans.getRange(2, 1, last - 1, n).getValues(), add = [];
    for (var i = 0; i < rows.length; i++) {
      try {
        var d = JSON.parse(rows[i][n - 1]);
        var c = d && d.concept && d.concept.comment;
        if (c) add.push([rows[i][0], rows[i][1], clean_(c, 600), 'Pending', String(rows[i][n - 2]), d.building, clean_(c, 600)]);
      } catch (e) {}
    }
    if (add.length) sh.getRange(2, 1, add.length, COMMENT_HEADERS.length).setValues(add);
  }
  return sh;
}

function approvedComments_() {
  var sh = commentSheet_(), last = sh.getLastRow(), map = {};
  if (last < 2) return map;
  var vals = sh.getRange(2, 1, last - 1, COMMENT_HEADERS.length).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][3] === 'Approved' && vals[i][2]) map[String(vals[i][4]) + '|' + String(vals[i][5])] = String(vals[i][2]).replace(/^'/, '');
  }
  return map;
}

// Add, update or remove the review row for one person's comment on one site.
function queueComment_(rid, site, address, comment) {
  var sh = commentSheet_(), last = sh.getLastRow(), row = 0;
  if (last >= 2) {
    var keys = sh.getRange(2, 5, last - 1, 3).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === rid && String(keys[i][1]) === site) { row = i + 2; var original = String(keys[i][2]); break; }
    }
  }
  var text = clean_(comment, 600);
  if (!text) { if (row) sh.deleteRow(row); return false; }
  if (row && original === text) return false; // unchanged: keep its current status
  var values = [[new Date(), clean_(address, 120), text, 'Pending', rid, site, text]];
  if (row) sh.getRange(row, 1, 1, values[0].length).setValues(values); else sh.appendRow(values[0]);
  return true;
}

// ---------- sources, alerts, spam ----------
var HEARD_LABELS = { facebook: 'Facebook group', nextdoor: 'Nextdoor', flyer: 'Flyer or QR code', friend: 'Friend or neighbor', plotline: 'Email or text from Plotline', other: 'Somewhere else' };

function recordSource_(rid, address, heard, src) {
  if (!heard && !src) return;
  var sh = sheet_('Sources', ['First answer', 'Site', 'How they heard', 'Arrived from link type', 'Respondent ID']);
  var last = sh.getLastRow();
  if (last >= 2) {
    var ids = sh.getRange(2, 5, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === rid) return;
  }
  sh.appendRow([new Date(), clean_(address, 120), HEARD_LABELS[heard] || clean_(heard, 40), clean_(src, 20), rid]);
}

function alert_(site, what) {
  sheet_('Alerts', ['When', 'Site', 'What happened']).appendRow([new Date(), clean_(site, 120), clean_(what, 300)]);
  notify_('alert:' + site, 'Plotline alert: ' + site, what + '\n\nSee the Alerts tab in your Plotline sheet.');
}

// Sends at most one email per topic per hour.
function notify_(key, subject, body) {
  if (!EMAIL_ALERTS) return;
  var cache = CacheService.getScriptCache();
  if (cache.get('mail:' + key)) return;
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), subject, body);
    cache.put('mail:' + key, '1', 3600);
  } catch (e) {}
}

function countBurst_(site) {
  var cache = CacheService.getScriptCache(), key = 'burst:' + site;
  var n = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(n), 600);
  return n;
}

function submitAnswer(p) {
  if (!p || typeof p.rid !== 'string' || !/^[a-z0-9]{6,40}$/.test(p.rid)) throw new Error('Bad respondent id');
  var r = p.row || {};
  if (typeof r.building !== 'string' || !/^[a-z0-9-]{1,40}$/.test(r.building)) throw new Error('Unknown building');
  if (typeof p.data !== 'string' || p.data.length > 8000) throw new Error('Bad data');
  var parsed = JSON.parse(p.data); // must be valid JSON

  // Bot traps: a hidden field people never see, or a survey finished impossibly fast.
  // Pretend it worked so bots learn nothing, but do not save it.
  if (p.hp) { alert_(r.address || r.building, 'Blocked an answer that filled in the hidden bot trap.'); return getAnswers(); }
  if (typeof p.elapsed === 'number' && p.elapsed < MIN_SURVEY_MS && !p.update) { alert_(r.address || r.building, 'Blocked an answer submitted ' + Math.round(p.elapsed / 1000) + ' seconds after opening the survey.'); return getAnswers(); }

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
    var comment = parsed && parsed.concept && parsed.concept.comment;
    if (queueComment_(p.rid, r.building, r.address, comment || '')) {
      notify_('comments', 'Plotline: comments waiting for review', 'New comments are waiting in the Comments tab of your Plotline sheet. Set Status to Approved to publish them.');
    }
    recordSource_(p.rid, r.address, parsed && parsed.heard, parsed && parsed.src);
  } finally {
    lock.releaseLock();
  }
  if (!target) {
    var burst = countBurst_(r.building);
    if (burst === BURST_LIMIT + 1) alert_(r.address || r.building, 'More than ' + BURST_LIMIT + ' new answers in 10 minutes. Check the Answers tab for repeated or unusual answers.');
  }
  return getAnswers();
}

function submitFeedback(p) {
  if (p && p.hp) { alert_('Feedback form', 'Blocked a message that filled in the hidden bot trap.'); return true; }
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
  if (p.hp) { alert_('Sign-up form', 'Blocked a sign-up that filled in the hidden bot trap.'); return true; }
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
