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
 *   Areas, Sites - the areas and sites shown on the website. Edit these to add, change or close sites.
 *   Custom answers - answers to a site's optional custom question
 *   Invitations - a log of survey invitations emailed from the Plotline menu
 *   Test answers, Test feedback, Test panel - answers from the test copy of the site
 *
 * After pasting: reload the sheet to see the "Plotline" menu.
 */

var BACKEND_VERSION = 3;
var SITE_URL = 'https://djwalsted.github.io/Plotline/';
// Required by law (CAN-SPAM) before emailing invitations to the panel. Example: 'Plotline, PO Box 123, Timnath, CO 80547'
var POSTAL_ADDRESS = '';

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
    var a = e && e.parameter && e.parameter.action, test = !!(e && e.parameter && e.parameter.test);
    if (a === 'answers') return json_({ ok: true, result: getAnswers(test) });
    if (a === 'config') return json_({ ok: true, result: getConfig() });
    if (a === 'version') return json_({ ok: true, result: { version: BACKEND_VERSION } });
    if (a === 'all') return json_({ ok: true, result: { version: BACKEND_VERSION, answers: getAnswers(test), config: getConfig() } });
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

function getAnswers(test) {
  var sh = sheet_(test ? TEST_SHEET : ANSWER_SHEET, HEADERS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var n = HEADERS.length;
  var vals = sh.getRange(2, n - 1, last - 1, 2).getValues();
  if (test) return vals.filter(function (v) { return v[0] && v[1]; }).map(function (v) { return { rid: String(v[0]), data: String(v[1]) }; });
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

  var test = !!p.test;
  if (test) { upsertAnswer_(TEST_SHEET, p, r, parsed); return getAnswers(true); }

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
    recordCustom_(p.rid, r.building, r.address, parsed && parsed.concept && parsed.concept.custom);
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
  var sh = sheet_(p.test ? 'Test feedback' : FEEDBACK_SHEET, ['Submitted', 'Feedback', 'Respondent ID']);
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
    var sh = sheet_(p.test ? 'Test panel' : PANEL_SHEET, PANEL_HEADERS);
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


// ---------- test copy of the site ----------
var TEST_SHEET = 'Test answers';
function upsertAnswer_(sheetName, p, r, parsed) {
  var values = [new Date(), clean_(r.address, 120), clean_(r.picks, 300), clean_(r.writein, 120), clean_(r.concept, 2), clean_(r.location, 2),
    clean_(r.freq, 40), clean_(r.when, 200), clean_(r.how, 40), clean_(r.spend, 40), clean_(r.extras, 300), clean_(r.comment, 600),
    clean_(r.town, 40), clean_(r.band, 40), clean_(r.modes, 120), clean_(r.trail, 40), clean_(r.kids, 40), p.rid, p.data];
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(sheetName, HEADERS), last = sh.getLastRow(), n = HEADERS.length, target = 0;
    if (last >= 2) {
      var keys = sh.getRange(2, n - 1, last - 1, 2).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (String(keys[i][0]) === p.rid) { try { if (JSON.parse(keys[i][1]).building === r.building) { target = i + 2; break; } } catch (e) {} }
      }
    }
    if (target) sh.getRange(target, 1, 1, n).setValues([values]); else sh.appendRow(values);
  } finally { lock.releaseLock(); }
}

// ---------- custom questions ----------
function recordCustom_(rid, site, address, custom) {
  if (!custom || !custom.q) return;
  var sh = sheet_('Custom answers', ['Submitted', 'Site', 'Question', 'Answer', 'Respondent ID', 'Site ID']);
  var last = sh.getLastRow(), row = 0;
  if (last >= 2) {
    var keys = sh.getRange(2, 5, last - 1, 2).getValues();
    for (var i = 0; i < keys.length; i++) if (String(keys[i][0]) === rid && String(keys[i][1]) === site) { row = i + 2; break; }
  }
  var v = [[new Date(), clean_(address, 120), clean_(custom.q, 200), clean_(custom.a, 300), rid, site]];
  if (row) sh.getRange(row, 1, 1, 6).setValues(v); else sh.appendRow(v[0]);
}

// ---------- areas and sites, edited in the sheet ----------
var AREA_HEADERS = ['Area ID (do not change)', 'Name', 'Town', 'Description', 'Center latitude', 'Center longitude', 'Map zoom (12-16)', 'Walking circle (Yes/No)', 'Walk circle latitude (optional)', 'Walk circle longitude (optional)', 'Label position on overview map (e, w, s, n, se)', 'Status (Open/Draft)'];
var SITE_HEADERS = ['Status (Open/Draft/Closed)', 'Area ID', 'Letter on map', 'Site ID (do not change)', 'Link name (letters, numbers, dashes)', 'Name', 'Listing status', 'Size', 'Price', 'Notes', 'Latitude', 'Longitude', 'Pin approximate (Yes/No)', 'Suggested idea', 'Idea description', 'Idea in a sentence (e.g. "a coffee shop")', 'Image (optional link)', 'Custom question (optional)', 'Custom answer choices (comma-separated, optional)'];
var SEED_AREAS = [["oldtown", "Old Town Timnath", "Timnath", "Storefronts and lots along Main Street in Timnath’s original downtown.", 40.5305, -104.9823, 15, "Yes", "", "", "e", "Open"], ["harmony", "I-25 & Harmony", "Timnath", "Fast-growing commercial corridor along Harmony Road at the interstate.", 40.5212, -104.9765, 13, "No", "", "", "s", "Open"], ["windsor", "Downtown Windsor", "Windsor", "Main Street (Highway 392) storefronts and pads between the shopping centers and Windsor Lake.", 40.4795, -104.9095, 14, "Yes", 40.4792, -104.9012, "s", "Open"], ["i25-392", "I-25 & Highway 392", "Windsor", "Windsor’s interstate gateway, west of town.", 40.484, -104.9845, 14, "No", "", "", "se", "Open"]];
var SEED_SITES = [["Open", "oldtown", "A", "a", "4100-main", "4100 Main St", "For lease", "About 1,627 sq ft", "Listed at $22.90/sq ft/yr NNN (about $3,100/mo)", "Standalone retail/office building with strong street visibility. A homebuilder currently lists its office here, so confirm what is available.", 40.53213, -104.98103, "No", "Coffee shop", "Espresso, drip and pastries. Opens early, closes mid-afternoon.", "a coffee shop", "img/a.jpg", "", ""], ["Open", "oldtown", "B", "c", "4109-main", "4109 Main St", "Sublease", "774–2,348 sq ft (three spaces)", "Rate not listed", "Class C building with Main Street frontage and natural light, currently office", 40.5316, -104.9829, "Yes", "Bakery & breakfast", "Bread, pastries and a short breakfast menu. Busy weekends, grab-and-go weekdays.", "a bakery and breakfast spot", "img/c.jpg", "", ""], ["Open", "oldtown", "C", "d", "4135-main", "4135 Main St", "Land for sale", "About 1.34 acres (parcel records differ)", "Listed at $955,000", "Vacant commercial land on Main Street", 40.5293, -104.983, "Yes", "Food truck park", "Rotating food trucks, shaded seating, a small play area and bike parking.", "a food truck park with a patio and play area", "img/d.jpg", "", ""], ["Open", "oldtown", "D", "b", "4138-main", "4138 Main St · Colorado Feed & Grain", "Status unclear", "2,872 sq ft end-cap, main level", "Rate not listed", "Historic 1920 building. Timnath Beerwerks runs a taproom in part of it; this listing is a separate end-cap space", 40.52898, -104.98171, "No", "Kitchen & patio", "A casual spot serving food to go with a beer from Timnath Beerwerks next door, with a family-friendly patio.", "a kitchen and patio next to the brewery", "img/b.jpg", "", ""], ["Open", "harmony", "A", "g", "ladera-i25-harmony", "Ladera · SE corner of I-25 & Harmony", "For sale or lease", "Large master-planned site (size varies by listing)", "Price on request", "Master-planned land at the interstate; listing details conflict, so confirm with the broker", 40.518, -104.9935, "Yes", "Shops, homes & a park", "Walkable blocks with ground-floor shops, apartments or townhomes above, and a central green.", "a mixed-use district with shops, homes and a park", "", "", ""], ["Open", "harmony", "B", "e", "5590-e-harmony", "5590 E Harmony Rd", "For sale or build-to-suit", "Two lots, 0.675 acres each", "About $30/sq ft", "Commercial lots under a mile from I-25, marketed to restaurants, retail and offices", 40.5225, -104.989, "Yes", "Drive-thru restaurant", "A quick-service restaurant with a drive-thru lane, open lunch through late evening.", "a fast-casual restaurant with a drive-thru", "", "", ""], ["Open", "harmony", "C", "f", "harmony-artisan-lots", "Harmony Rd & Artisan Pl lots", "Land for sale", "6.40 and 2.47 acres", "$5.16M and $1.99M ($7.15M together)", "Improved lots zoned commercial mixed-use, about five minutes from I-25", 40.5239, -104.9591, "Yes", "Grocery store", "A full-service grocery store with a pharmacy, so nearby neighborhoods do not have to drive across town.", "a full grocery store with a pharmacy", "", "", ""], ["Open", "windsor", "A", "i", "windsor-1296-main", "1296 Main St, Unit D", "For lease", "1,635 sq ft", "Rate not listed", "Street retail at Main and 13th", 40.48033, -104.9218, "No", "Ice cream & dessert", "Locally made ice cream, shakes and treats, open late on summer evenings.", "an ice cream and dessert shop", "", "", ""], ["Open", "windsor", "B", "h", "windsor-1039-main", "1039 Main St · Windsor Town Center", "For lease", "2,886 sq ft", "Listed at $14.50/sq ft/yr", "Shopping center space on Main Street; listings show different sizes, so confirm with the broker", 40.47835, -104.91535, "No", "Family restaurant", "Breakfast through dinner, booths, a kids menu and room for big tables.", "a sit-down family restaurant", "", "", ""], ["Open", "windsor", "C", "j", "windsor-101-main", "101 Main St", "For sale", "Retail pad with Main Street frontage", "Price on request", "Frontage on Main Street and Highway 392 with flexible zoning for retail, commercial or mixed use", 40.47961, -104.89718, "No", "Gas station & market", "Fuel, EV chargers and a small market with grab-and-go food.", "a gas station with a convenience market", "", "", ""], ["Open", "i25-392", "A", "k", "i25-392-nec", "NE corner of I-25 & Hwy 392", "Land for sale", "Commercial development land", "Price on request", "Interstate frontage at Windsor’s west gateway", 40.484, -104.9845, "Yes", "Hotel & restaurants", "A mid-size hotel with two or three restaurants, serving travelers and weekend events.", "a hotel with restaurants", "", "", ""]];

function configSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var a = ss.getSheetByName('Areas'), s = ss.getSheetByName('Sites');
  if (!a) { a = sheet_('Areas', AREA_HEADERS); a.getRange(2, 1, SEED_AREAS.length, AREA_HEADERS.length).setValues(SEED_AREAS); }
  if (!s) {
    s = sheet_('Sites', SITE_HEADERS); s.getRange(2, 1, SEED_SITES.length, SITE_HEADERS.length).setValues(SEED_SITES);
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Open', 'Draft', 'Closed'], true).build();
    s.getRange(2, 1, s.getMaxRows() - 1, 1).setDataValidation(rule);
  }
  return { areas: a, sites: s };
}

function getConfig() {
  var sh = configSheets_();
  var num = function (v) { var n = Number(v); return isFinite(n) && v !== '' ? n : null; };
  var yes = function (v) { return /^y/i.test(String(v)); };
  var slug = function (v) { return String(v || '').toLowerCase().replace(/[^a-z0-9-]/g, '').slice(0, 60); };
  var areas = [], sites = [];
  var la = sh.areas.getLastRow();
  if (la >= 2) sh.areas.getRange(2, 1, la - 1, AREA_HEADERS.length).getValues().forEach(function (r) {
    var id = slug(r[0]);
    if (!id || /^draft/i.test(String(r[11]))) return;
    var c = [num(r[4]), num(r[5])]; if (c[0] === null || c[1] === null) return;
    var w = [num(r[8]), num(r[9])];
    areas.push({ id: id, slug: id, name: String(r[1]), town: String(r[2]), blurb: String(r[3]), center: c, zoom: num(r[6]) || 14,
      walk: yes(r[7]), walkCenter: w[0] !== null && w[1] !== null ? w : null, labelPos: String(r[10] || 's').toLowerCase() });
  });
  var ls = sh.sites.getLastRow();
  if (ls >= 2) sh.sites.getRange(2, 1, ls - 1, SITE_HEADERS.length).getValues().forEach(function (r) {
    var status = String(r[0] || '').trim();
    if (!/^(open|closed)$/i.test(status)) return;
    var id = slug(r[3]), lat = num(r[10]), lng = num(r[11]);
    if (!id || lat === null || lng === null) return;
    var choices = String(r[18] || '').split(',').map(function (x) { return x.trim(); }).filter(String);
    sites.push({ status: status.toLowerCase(), area: slug(r[1]), letter: String(r[2] || '').slice(0, 2), id: id, slug: slug(r[4]) || id,
      name: String(r[5]), tag: String(r[6]), size: String(r[7]), price: String(r[8]), notes: String(r[9]), lat: lat, lng: lng, approx: yes(r[12]),
      idea: String(r[13]), ideaBlurb: String(r[14]), ideaLabel: String(r[15]), img: String(r[16] || ''),
      customQ: String(r[17] || ''), customChoices: choices });
  });
  return { areas: areas, sites: sites };
}

// ---------- Plotline menu: email invitations ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Plotline')
    .addItem('Email a survey invitation\u2026', 'showInviteDialog')
    .addItem('Set up Areas and Sites tabs', 'setupTabs')
    .addToUi();
}
function setupTabs() { configSheets_(); commentSheet_(); SpreadsheetApp.getUi().alert('Areas, Sites and Comments tabs are ready.'); }

function panelRows_() {
  var sh = sheet_(PANEL_SHEET, PANEL_HEADERS), last = sh.getLastRow();
  if (sh.getRange(1, 11).getValue() !== 'Last invited') sh.getRange(1, 11).setValue('Last invited').setFontWeight('bold');
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 11).getValues().map(function (r, i) {
    return { row: i + 2, email: String(r[1]).trim(), town: String(r[5]).trim(), zip: String(r[6]).trim(), status: String(r[9]).trim() || 'Active' };
  });
}
function inviteRecipients_(opts) {
  var towns = (opts.towns || []).map(String), zips = String(opts.zips || '').split(/[\s,]+/).filter(String);
  return panelRows_().filter(function (p) {
    if (!p.email || !/^active$/i.test(p.status)) return false;
    if (!towns.length && !zips.length) return true;
    return towns.indexOf(p.town) >= 0 || zips.indexOf(p.zip) >= 0;
  });
}
function inviteOptions() {
  var cfg = getConfig(), towns = {};
  panelRows_().forEach(function (p) { if (p.town) towns[p.town] = (towns[p.town] || 0) + 1; });
  return { sites: cfg.sites.filter(function (s) { return s.status === 'open'; }).map(function (s) { return { id: s.id, slug: s.slug, name: s.name, idea: s.idea }; }),
    towns: towns, quota: MailApp.getRemainingDailyQuota(), postal: !!POSTAL_ADDRESS, me: Session.getEffectiveUser().getEmail() };
}
function previewInvite(opts) { return { count: inviteRecipients_(opts).length, quota: MailApp.getRemainingDailyQuota() }; }
function sendInvite(opts) {
  var site = getConfig().sites.filter(function (s) { return s.id === opts.siteId; })[0];
  if (!site) throw new Error('Pick a site.');
  var link = SITE_URL + '?src=email#' + site.slug;
  var body = String(opts.message || '').replace(/\{link\}/g, link);
  if (body.indexOf(link) < 0) body += '\n\n' + link;
  var footer = '\n\n--\nYou are getting this because you signed up at Plotline to hear about projects near you. Reply "unsubscribe" to stop.\n' + (POSTAL_ADDRESS || '');
  var subject = String(opts.subject || 'What should go here?').slice(0, 150);
  if (opts.testOnly) { MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '[TEST] ' + subject, body + footer); return { sent: 1, test: true }; }
  if (!POSTAL_ADDRESS) throw new Error('Add your postal address to POSTAL_ADDRESS at the top of Code.gs first. It is required by law for marketing email.');
  var list = inviteRecipients_(opts), quota = MailApp.getRemainingDailyQuota();
  if (list.length > quota) throw new Error('Google allows ' + quota + ' more emails today; this would send ' + list.length + '. Narrow the towns or ZIP codes, or send the rest tomorrow.');
  var sh = sheet_(PANEL_SHEET, PANEL_HEADERS), now = new Date();
  list.forEach(function (p) { MailApp.sendEmail(p.email, subject, body + footer); sh.getRange(p.row, 11).setValue(now); });
  sheet_('Invitations', ['Sent', 'Site', 'Recipients', 'Towns', 'ZIP codes', 'Subject']).appendRow([now, site.name, list.length, (opts.towns || []).join(', '), clean_(opts.zips, 200), clean_(subject, 150)]);
  return { sent: list.length };
}
function showInviteDialog() {
  var html = HtmlService.createHtmlOutput(INVITE_HTML).setWidth(480).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Email a survey invitation');
}
var INVITE_HTML = '<!doctype html><html><head><base target="_top"><style>body{font:14px Arial,sans-serif;margin:0;padding:4px 2px;color:#18282A}label{display:block;font-weight:bold;margin:12px 0 4px}select,input,textarea{width:100%;box-sizing:border-box;font:inherit;padding:6px}textarea{height:120px}.towns label{font-weight:normal;display:inline-block;margin:2px 12px 2px 0}.towns input{width:auto}.row{display:flex;gap:8px;margin-top:14px}button{padding:8px 14px;font:inherit;cursor:pointer}.primary{background:#0D6B64;color:#fff;border:0;border-radius:6px}.note{color:#5A6B69;font-size:12px}.warn{color:#9A5B00}#msg{margin-top:10px;font-weight:bold}</style></head><body>'
 + '<label for="site">Survey</label><select id="site"></select>'
 + '<label>Who gets it</label><div class="towns" id="towns"></div><input id="zips" placeholder="Or ZIP codes, separated by commas (optional)"><p class="note">Leave all boxes unchecked and ZIP blank to email everyone on the panel. Only Active members with an email address are included.</p>'
 + '<label for="subject">Subject</label><input id="subject" value="What should go here? 2-minute survey">'
 + '<label for="message">Message</label><textarea id="message">Hi neighbor,\n\nA site near you is open for feedback on Plotline. Tell the people deciding what you would actually use. It takes about 2 minutes and never asks for your name.\n\n{link}\n\nThanks for speaking up for your neighborhood.</textarea><p class="note">{link} becomes the survey link.</p>'
 + '<p id="count" class="note"></p><p id="postal" class="note warn"></p>'
 + '<div class="row"><button id="test">Send a test to me</button><button id="send" class="primary">Send</button></div><div id="msg"></div>'
 + '<script>var O;function opts(){var t=[].slice.call(document.querySelectorAll(".towns input:checked")).map(function(x){return x.value});return{siteId:document.getElementById("site").value,towns:t,zips:document.getElementById("zips").value,subject:document.getElementById("subject").value,message:document.getElementById("message").value}}'
 + 'function count(){google.script.run.withSuccessHandler(function(r){document.getElementById("count").textContent=r.count+" recipient(s). Google allows "+r.quota+" more emails today."}).previewInvite(opts())}'
 + 'google.script.run.withSuccessHandler(function(o){O=o;var s=document.getElementById("site");o.sites.forEach(function(x){var e=document.createElement("option");e.value=x.id;e.textContent=x.name+" ("+x.idea+")";s.appendChild(e)});var t=document.getElementById("towns");Object.keys(o.towns).forEach(function(k){var l=document.createElement("label");l.innerHTML="<input type=checkbox value=\\""+k.replace(/"/g,"")+"\\"> "+k.replace(/</g,"")+" ("+o.towns[k]+")";t.appendChild(l)});if(!o.postal)document.getElementById("postal").textContent="Add your postal address to POSTAL_ADDRESS in Code.gs before sending to the panel (required by law). Test emails to yourself work now.";document.addEventListener("change",count);document.getElementById("zips").addEventListener("input",count);count()}).inviteOptions();'
 + 'function run(test){var m=document.getElementById("msg");m.textContent="Sending...";var o=opts();o.testOnly=test;google.script.run.withSuccessHandler(function(r){m.textContent=r.test?"Test sent to "+O.me+".":"Sent to "+r.sent+" people. Logged in the Invitations tab."}).withFailureHandler(function(e){m.textContent=e.message}).sendInvite(o)}'
 + 'document.getElementById("test").onclick=function(){run(true)};document.getElementById("send").onclick=function(){if(confirm("Send this invitation now?"))run(false)};</script></body></html>';
