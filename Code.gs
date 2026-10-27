/**
 * MVB Serve & Pass Stats — Google Apps Script backend
 *
 * Bound to the "MVB GAME Serve and Pass Master" Google Sheet. Practices go to a second
 * workbook ("MVB PRACTICE Serve and Pass Master") whose ID is stored in the script property
 * PRACTICE_SHEET_ID. Both workbooks have the same tabs.
 * The phone/laptop app posts JSON here; this script writes into the right workbook.
 *
 * Tabs it uses:
 *   Face Page  — season totals (formulas extended when a new session tab is created)
 *   Roster     — jersey numbers (col A) and names (col B)
 *   Sessions   — one row per game/practice tab
 *   Log        — one row per serve/pass entry (audit trail; player rows are rebuilt from it)
 *   Template   — hidden layout copied for each new session tab
 *
 * ONE-TIME SETUP (see README): set PASSCODE and PRACTICE_SHEET_ID in Project Settings >
 * Script properties, then Deploy > New deployment > Web app, execute as Me, access: Anyone.
 */

var TAB = { FACE: 'Face Page', ROSTER: 'Roster', SESSIONS: 'Sessions', LOG: 'Log', TEMPLATE: 'Template' };
var FIRST_ATTEMPT_COL = 3;   // C
var MAX_ATTEMPTS = 50;       // C..AZ
var FACE_COUNT_COL = 56;     // BD on a session tab holds the count of 0s, BE..BH = 1..4
var LOG_COLS = { ts: 1, tab: 2, set: 3, player: 4, type: 5, grade: 6, source: 7, id: 8 };
var SESSION_COLS = { tab: 1, type: 2, date: 3, opp: 4, home: 5, sets: 6, entries: 7, created: 8, updated: 9, id: 10, status: 11 };

/* ---------------------------------------------------------------- entry points */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!checkToken_(p.token)) return json_({ ok: false, error: 'bad passcode' });
  if (p.action === 'roster') return json_({ ok: true, roster: getRoster_() });
  return json_({ ok: true, ping: 'MVB stats backend', time: new Date().toISOString() });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad JSON' }); }
  if (!checkToken_(body.token)) return json_({ ok: false, error: 'bad passcode' });

  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (err) { return json_({ ok: false, error: 'busy, retry', retry: true }); }
  try {
    var ss = book_(body.session && body.session.kind);
    var out;
    switch (body.action) {
      case 'roster':       out = { roster: getRoster_() }; break;
      case 'sessions':     out = { sessions: listSessions_() }; break;
      case 'events':       out = { events: listEvents_(ss, body.session) }; break;
      case 'startSession': out = startSession_(ss, body.session); break;
      case 'entries':      out = addEntries_(ss, body.session, body.entries || []); break;
      case 'undo':         out = undoEntry_(ss, body.session, body.entryId); break;
      case 'sets':         out = updateSets_(ss, body.session); break;
      case 'deleteSession': out = deleteSession_(ss, body.session); break;
      default:             return json_({ ok: false, error: 'unknown action ' + body.action });
    }
    SpreadsheetApp.flush();
    out.ok = true;
    return json_(out);
  } catch (err) {
    var msg = String(err && err.message || err) + ' [' + body.action + (body.session ? ' ' + body.session.kind : '') + ']';
    console.error(msg + ' session=' + (body.session && body.session.id) + ' ' + (err && err.stack || ''));
    return json_({ ok: false, error: msg });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** The games workbook is the one this script is bound to; practices live in the workbook named by PRACTICE_SHEET_ID. */
function book_(kind) {
  if (kind !== 'practice') return SpreadsheetApp.getActiveSpreadsheet();
  var id = PropertiesService.getScriptProperties().getProperty('PRACTICE_SHEET_ID');
  if (!id) throw new Error('PRACTICE_SHEET_ID script property is not set (see README)');
  return SpreadsheetApp.openById(id);
}

function checkToken_(token) {
  var want = PropertiesService.getScriptProperties().getProperty('PASSCODE');
  if (!want) return true;               // no passcode configured yet: open (set one before sharing the link)
  return String(token || '') === String(want);
}

/* ---------------------------------------------------------------- roster */

function getRoster_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.ROSTER);
  var vals = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getValues();
  var out = [];
  vals.forEach(function (r) { if (typeof r[0] === 'number') out.push({ n: r[0], name: String(r[1] || '') }); });
  return out;
}

/* ---------------------------------------------------------------- read-back for the app */

/** Every game and practice known to the two workbooks (so the app can show sessions started elsewhere, e.g. selfTest). */
function listSessions_() {
  var out = [];
  var books = [['game', SpreadsheetApp.getActiveSpreadsheet()]];
  try { books.push(['practice', book_('practice')]); } catch (err) { /* practice book not configured yet */ }
  books.forEach(function (b) {
    var se = b[1].getSheetByName(TAB.SESSIONS);
    var n = se.getLastRow();
    if (n < 2) return;
    se.getRange(2, 1, n - 1, SESSION_COLS.status).getValues().forEach(function (r) {
      if (!r[SESSION_COLS.id - 1]) return;
      out.push({ id: String(r[SESSION_COLS.id - 1]), kind: r[SESSION_COLS.type - 1] === 'practice' ? 'practice' : 'game', tab: tabKey_(r[SESSION_COLS.tab - 1]),
                 date: isoDate_(r[SESSION_COLS.date - 1]), opp: String(r[SESSION_COLS.opp - 1] || ''), home: r[SESSION_COLS.home - 1] !== 'Away',
                 sets: Number(r[SESSION_COLS.sets - 1]) || 1, entries: Number(r[SESSION_COLS.entries - 1]) || 0, ended: r[SESSION_COLS.status - 1] === 'Final' });
    });
  });
  return out;
}

/** All Log entries for one session, in order, so another device can pick a game up. */
function listEvents_(ss, s) {
  var srow = findSessionRow_(ss, s && s.id);
  if (!srow) return [];
  var log = ss.getSheetByName(TAB.LOG);
  var n = log.getLastRow(), out = [];
  if (n < 2) return out;
  log.getRange(2, 1, n - 1, LOG_COLS.id).getValues().forEach(function (r) {
    if (tabKey_(r[LOG_COLS.tab - 1]) !== tabKey_(srow.tab)) return;
    var ts = r[LOG_COLS.ts - 1];
    out.push({ id: String(r[LOG_COLS.id - 1]), set: Number(r[LOG_COLS.set - 1]) || 1, player: Number(r[LOG_COLS.player - 1]), type: r[LOG_COLS.type - 1] === 'S' ? 'S' : 'P',
               grade: Number(r[LOG_COLS.grade - 1]), ts: ts instanceof Date ? ts.getTime() : Date.now() });
  });
  return out;
}

function isoDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
}

/* ---------------------------------------------------------------- sessions */

/** session = { id, kind: 'game'|'practice', date: 'YYYY-MM-DD', opp, home: bool, sets: n } */
function startSession_(ss, s) {
  if (!s || !s.id) throw new Error('session.id required');
  var existing = findSessionRow_(ss, s.id);
  if (existing) {
    extendFacePage_(ss, existing.tab);          // idempotent; repairs a partially completed first attempt
    return { tab: existing.tab, created: false };
  }

  var name = uniqueTabName_(ss, s.kind === 'practice' ? (nextNumber_(ss, 'practice', 'P') + ' ' + fmtDate_(s.date)) : (nextNumber_(ss, 'game', 'G') + ' ' + fmtDate_(s.date) + ' ' + (s.opp || '')));
  var tpl = ss.getSheetByName(TAB.TEMPLATE);
  if (!tpl) throw new Error('Template tab is missing');
  var sh = tpl.copyTo(ss).setName(name);
  sh.showSheet();
  // Register the tab first so a failure later in this function cannot create a duplicate tab on retry.
  var se = ss.getSheetByName(TAB.SESSIONS);
  se.getRange(se.getLastRow() + 1, SESSION_COLS.tab).setNumberFormat('@');
  se.appendRow([name, s.kind || 'game', s.date || '', s.kind === 'practice' ? '' : (s.opp || ''),
                s.kind === 'practice' ? '' : (s.home === false ? 'Away' : 'Home'), s.sets || 1, 0, new Date(), new Date(), s.id]);
  setStatus_(se, se.getLastRow(), s);
  try { ss.setActiveSheet(sh); ss.moveActiveSheet(ss.getNumSheets()); } catch (err) { /* ordering only */ }
  var note = (s.kind === 'practice' ? 'Practice' : ((s.home === false ? 'at ' : 'vs ') + (s.opp || ''))) + ' — ' + (s.date || '') + ' — recorded live by the stats app';
  sh.getRange(60, 1).setValue(note);
  extendFacePage_(ss, name);
  return { tab: name, created: true };
}

function updateSets_(ss, s) {
  var row = findSessionRow_(ss, s && s.id);
  if (!row) { startSession_(ss, s); row = findSessionRow_(ss, s.id); }   // its start never arrived, or it was removed: recreate
  var se = ss.getSheetByName(TAB.SESSIONS);
  se.getRange(row.row, SESSION_COLS.sets).setValue(s.sets || 1);
  se.getRange(row.row, SESSION_COLS.updated).setValue(new Date());
  setStatus_(se, row.row, s);
  return { tab: row.tab };
}

/** "Final" / "In progress" in the Sessions tab (adds the Status header if the sheet predates it). */
function setStatus_(se, row, s) {
  if (se.getRange(1, SESSION_COLS.status).getValue() === '') se.getRange(1, SESSION_COLS.status).setValue('Status').setFontWeight('bold');
  se.getRange(row, SESSION_COLS.status).setValue(s && s.ended ? 'Final' : 'In progress');
}

function findSessionRow_(ss, id) {
  var se = ss.getSheetByName(TAB.SESSIONS);
  var n = se.getLastRow();
  if (n < 2) return null;
  var vals = se.getRange(2, 1, n - 1, SESSION_COLS.id).getValues();
  for (var i = 0; i < vals.length; i++) if (String(vals[i][SESSION_COLS.id - 1]) === String(id)) return { row: i + 2, tab: tabKey_(vals[i][0]) };
  return null;
}

function nextNumber_(ss, kind, prefix) {
  var se = ss.getSheetByName(TAB.SESSIONS);
  var n = se.getLastRow(), max = 0;
  if (n >= 2) se.getRange(2, 1, n - 1, SESSION_COLS.type).getValues().forEach(function (r) {
    var m = tabKey_(r[0]).match(new RegExp('^' + prefix + '(\\d+)\\b'));
    if (r[SESSION_COLS.type - 1] === kind && m) max = Math.max(max, Number(m[1]));
  });
  return prefix + (max + 1);
}

/** Delete one game/practice: its tab, its Log rows, its Sessions row, and its terms in the Face Page formulas. */
function deleteSession_(ss, s) {
  var srow = findSessionRow_(ss, s && s.id);
  if (!srow) return { deleted: false };
  var tab = srow.tab;
  removeFromFacePage_(ss, tab);                 // before deleting the sheet, so no #REF! appears
  var sh = ss.getSheetByName(tab);
  if (sh) ss.deleteSheet(sh);
  var log = ss.getSheetByName(TAB.LOG);
  var n = log.getLastRow();
  if (n >= 2) {
    var tabs = log.getRange(2, LOG_COLS.tab, n - 1, 1).getValues();
    for (var i = tabs.length - 1; i >= 0; i--) if (tabKey_(tabs[i][0]) === tabKey_(tab)) log.deleteRow(i + 2);
  }
  ss.getSheetByName(TAB.SESSIONS).deleteRow(srow.row);
  return { deleted: true, tab: tab };
}

function removeFromFacePage_(ss, tabName) {
  var fp = ss.getSheetByName(TAB.FACE);
  var last = Math.max(fp.getLastRow(), 60);
  var colA = fp.getRange(1, 1, last, 1).getValues();
  var q = "'" + tabName.replace(/'/g, "''") + "'";
  var esc = function (x) { return x.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); };
  for (var r = 1; r <= last; r++) {
    if (typeof colA[r - 1][0] !== 'number') continue;
    var range = fp.getRange(r, 3, 1, 5), formulas = range.getFormulas()[0], changed = false;
    for (var i = 0; i < 5; i++) {
      var f = String(formulas[i] || '');
      var term = q + '!' + colLetter_(FACE_COUNT_COL + i) + r;
      var re = new RegExp('\\+' + esc(term) + '(?![0-9])');
      if (re.test(f)) { formulas[i] = f.replace(re, ''); changed = true; }
    }
    if (changed) range.setFormulas([formulas]);
  }
}

/** Tab names as stored in Log/Sessions, tolerant of Sheets having auto-converted a date-looking name into a date. */
function tabKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'MMM d');
  return String(v);
}

function uniqueTabName_(ss, base) {
  base = String(base).replace(/[\[\]\*\?\/\\:']/g, '').trim().slice(0, 31) || 'Session';
  var name = base, i = 2;
  while (ss.getSheetByName(name)) { name = (base.slice(0, 31 - (' (' + i + ')').length) + ' (' + i + ')'); i++; }
  return name;
}

function fmtDate_(iso) {
  if (!iso) return '';
  var m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return String(iso);
  return ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][Number(m[2]) - 1] + ' ' + Number(m[3]);
}

/** Append +'Tab'!BDr (… BHr) to each Face Page total formula in columns C..G. */
function extendFacePage_(ss, tabName) {
  var fp = ss.getSheetByName(TAB.FACE);
  var last = Math.max(fp.getLastRow(), 60);
  var colA = fp.getRange(1, 1, last, 1).getValues();
  var q = "'" + tabName.replace(/'/g, "''") + "'";
  var touched = 0;
  for (var r = 1; r <= last; r++) {
    if (typeof colA[r - 1][0] !== 'number') continue;
    var range = fp.getRange(r, 3, 1, 5);
    var formulas = range.getFormulas()[0];
    var changed = false;
    for (var i = 0; i < 5; i++) {
      var f = String(formulas[i] || '').trim();
      if (!f) continue;
      if (f.indexOf(q + '!') >= 0) continue;
      var term = q + '!' + colLetter_(FACE_COUNT_COL + i) + r;
      formulas[i] = f.endsWith(')') ? f.slice(0, -1) + '+' + term + ')' : f + '+' + term;
      changed = true; touched++;
    }
    if (changed) range.setFormulas([formulas]);
  }
  return touched;
}

function colLetter_(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }

/* ---------------------------------------------------------------- entries */

/** entries = [{ id, set, player, type: 'S'|'P', grade: 0-4, ts }] */
function addEntries_(ss, session, entries) {
  var srow = findSessionRow_(ss, session && session.id);
  if (!srow) { startSession_(ss, session); srow = findSessionRow_(ss, session.id); }
  var tab = srow.tab;
  var log = ss.getSheetByName(TAB.LOG);
  var known = knownEntryIds_(log);
  var rows = [], touched = {};
  entries.forEach(function (e) {
    if (!e || !e.id || known[e.id]) return;
    var player = Number(e.player), grade = Number(e.grade), type = e.type === 'S' ? 'S' : 'P';
    if (!(grade >= 0 && grade <= 4)) return;
    rows.push([e.ts ? new Date(e.ts) : new Date(), tab, Number(e.set) || 1, player, type, grade, 'app', String(e.id)]);
    touched[type + ':' + player] = true;
    known[e.id] = true;
  });
  if (rows.length) {
    var r0 = log.getLastRow() + 1;
    log.getRange(r0, LOG_COLS.tab, rows.length, 1).setNumberFormat('@');
    log.getRange(r0, 1, rows.length, rows[0].length).setValues(rows);
  }
  var skipped = rebuildRows_(ss, tab, Object.keys(touched));
  bumpSession_(ss, srow.row, session);
  return { tab: tab, added: rows.length, skipped: skipped };
}

function undoEntry_(ss, session, entryId) {
  var srow = findSessionRow_(ss, session && session.id);
  if (!srow) return { removed: false };
  var log = ss.getSheetByName(TAB.LOG);
  var n = log.getLastRow();
  if (n < 2) return { removed: false };
  var ids = log.getRange(2, LOG_COLS.id, n - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(entryId)) {
      var row = i + 2;
      var vals = log.getRange(row, 1, 1, LOG_COLS.id).getValues()[0];
      log.deleteRow(row);
      rebuildRows_(ss, srow.tab, [vals[LOG_COLS.type - 1] + ':' + vals[LOG_COLS.player - 1]]);
      bumpSession_(ss, srow.row, session);
      return { removed: true };
    }
  }
  return { removed: false };
}

function knownEntryIds_(log) {
  var n = log.getLastRow(), out = {};
  if (n < 2) return out;
  log.getRange(2, LOG_COLS.id, n - 1, 1).getValues().forEach(function (r) { if (r[0]) out[String(r[0])] = true; });
  return out;
}

function bumpSession_(ss, row, session) {
  var se = ss.getSheetByName(TAB.SESSIONS);
  var log = ss.getSheetByName(TAB.LOG);
  var tab = tabKey_(se.getRange(row, SESSION_COLS.tab).getValue());
  var n = log.getLastRow(), count = 0;
  if (n >= 2) log.getRange(2, LOG_COLS.tab, n - 1, 1).getValues().forEach(function (r) { if (tabKey_(r[0]) === tabKey_(tab)) count++; });
  se.getRange(row, SESSION_COLS.entries).setValue(count);
  se.getRange(row, SESSION_COLS.updated).setValue(new Date());
  if (session && session.sets) se.getRange(row, SESSION_COLS.sets).setValue(session.sets);
  if (session) setStatus_(se, row, session);
}

/** Rewrite the attempt cells (C..AZ) for the given "type:player" keys from the Log, in log order. */
function rebuildRows_(ss, tab, keys) {
  if (!keys.length) return [];
  var sh = ss.getSheetByName(tab) || recreateTab_(ss, tab);
  var blocks = blocks_(sh);                     // [passing, serving] -> { jersey: row }
  var log = ss.getSheetByName(TAB.LOG);
  var n = log.getLastRow();
  var all = n >= 2 ? log.getRange(2, 1, n - 1, LOG_COLS.id).getValues() : [];
  var skipped = [];
  keys.forEach(function (key) {
    var type = key.split(':')[0], player = Number(key.split(':')[1]);
    var grades = [];
    all.forEach(function (r) {
      if (tabKey_(r[LOG_COLS.tab - 1]) === tabKey_(tab) && r[LOG_COLS.type - 1] === type && Number(r[LOG_COLS.player - 1]) === player) grades.push(Number(r[LOG_COLS.grade - 1]));
    });
    var map = type === 'S' ? blocks[1] : blocks[0];
    var row = map[player];
    if (!row) { skipped.push('#' + player + ' is not in column A of the tab'); return; }
    var vals = [];
    for (var i = 0; i < MAX_ATTEMPTS; i++) vals.push(i < grades.length ? grades[i] : '');
    if (grades.length > MAX_ATTEMPTS) skipped.push('#' + player + ' ' + (type === 'S' ? 'serves' : 'passes') + ' over 50; only the first 50 are on the tab');
    sh.getRange(row, FIRST_ATTEMPT_COL, 1, MAX_ATTEMPTS).setValues([vals]);
  });
  return skipped;
}

/** A tab listed in Sessions but deleted by hand: copy the Template back under the same name and re-add it to the Face Page. */
function recreateTab_(ss, tab) {
  var tpl = ss.getSheetByName(TAB.TEMPLATE);
  if (!tpl) throw new Error('Template tab is missing');
  var sh = tpl.copyTo(ss).setName(tab);
  sh.showSheet();
  sh.getRange(60, 1).setValue('Recreated by the stats app from the Log');
  extendFacePage_(ss, tab);
  return sh;
}

/** Find the PASSING and SERVING blocks by their '#' header rows; map jersey -> row. */
function blocks_(sh) {
  var vals = sh.getRange(1, 1, Math.max(sh.getLastRow(), 60), 1).getValues();
  var out = [], cur = null;
  for (var r = 1; r <= vals.length; r++) {
    var v = vals[r - 1][0];
    if (v === '#') { cur = {}; out.push(cur); continue; }
    if (cur && typeof v === 'number') cur[v] = r;
    else if (cur && (v === '' || v === null) && Object.keys(cur).length) cur = null;
  }
  if (out.length < 2) throw new Error('could not find PASSING and SERVING blocks');
  return out;
}

/* ---------------------------------------------------------------- manual helpers (run from the editor) */

/** Rebuild every player row of one tab from the Log (use if a tab was edited by hand and you want it back in sync). */
function rebuildWholeTab(tabName, kind) {
  var ss = book_(kind || 'game');
  var sh = ss.getSheetByName(tabName);
  var b = blocks_(sh), keys = [];
  Object.keys(b[0]).forEach(function (p) { keys.push('P:' + p); });
  Object.keys(b[1]).forEach(function (p) { keys.push('S:' + p); });
  rebuildRows_(ss, tabName, keys);
}

/** Quick self-test from the editor: creates a game tab "G1 <today> SELFTEST" and adds two entries. Delete the tab after. */
function selfTest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var id = 'selftest-' + Date.now();
  var s = { id: id, kind: 'game', date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), opp: 'SELFTEST', home: true, sets: 1 };
  Logger.log(startSession_(ss, s));
  Logger.log(addEntries_(ss, s, [{ id: id + '-1', set: 1, player: 1, type: 'S', grade: 4 }, { id: id + '-2', set: 1, player: 1, type: 'P', grade: 2 }]));
  Logger.log(undoEntry_(ss, s, id + '-2'));
}

/** Same self-test against the practice workbook (checks PRACTICE_SHEET_ID). Creates a tab "P1 <today>" there. */
function selfTestPractice() {
  var s = { id: 'selftest-' + Date.now(), kind: 'practice', date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), sets: 1 };
  var ss = book_('practice');
  Logger.log(startSession_(ss, s));
  Logger.log(addEntries_(ss, s, [{ id: s.id + '-1', set: 1, player: 1, type: 'P', grade: 3 }]));
}
