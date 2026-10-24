/**
 * Triton MMR — backend for the Riftbound pairings board.
 *
 * Built the same way as the Nano3 dicing endpoint: one web app, actions on the
 * query string, CORS-safe. The board reads ratings before an event and writes
 * the event back as results come in, both with plain fetch, no tabs opening.
 *
 * SETUP, ONCE
 *   1. Open the Triton MMR spreadsheet, Extensions > Apps Script.
 *   2. Replace everything in Code.gs with this file. Save.
 *   3. Run the function `setup` once from the editor and approve the permissions
 *      prompt. This creates the tabs and proves the script can reach the sheet.
 *   4. Deploy > New deployment > Web app
 *          Execute as:      Me
 *          Who has access:  Anyone
 *      Copy the /exec URL and paste it into the board's Sheet sync card.
 *   5. After ANY later edit to this file: Deploy > Manage deployments > pencil >
 *      Version: New version > Deploy. Without that step the old code keeps running.
 *
 * ACTIONS
 *   GET  ?action=ping                 is the endpoint alive
 *   GET  ?action=ratings              every known player and their current rating
 *   GET  ?action=checkin&name=&date=&on=1|0
 *                                     tick or untick "Checked in" on the TNN Data signups
 *   POST (body = JSON, text/plain)    write one event: ratings, event log, match log
 *
 * Syncing the same event any number of times is safe: that date's log rows are
 * replaced, and every player's lifetime totals are rebuilt from the Event Log.
 * The board relies on this, since it syncs after every reported result.
 *
 * CORS: GET responses from ContentService are readable cross-origin. The board
 * POSTs with Content-Type text/plain, which is a "simple request", so the browser
 * sends it without a preflight. Do not switch the board to application/json:
 * Apps Script cannot answer the OPTIONS preflight that would trigger.
 */

/* Bump this whenever you edit the file. The pairings and standings pages show it,
   so you can tell at a glance whether the deployment is serving the new code. */
var VERSION = '5-checkout';

/* The signup form's spreadsheet, where the "Checked in" boxes live. */
var SIGNUP_SHEET_ID = '1Tsoclk1mt6Yr8r2z9vT-j5KSBpxx-phHeVTj17B89ac';
var SIGNUP_TAB = 'TNN Signups';

var RATINGS = 'Ratings';
var EVENTS  = 'Event Log';
var MATCHES = 'Match Log';

var RATING_HEADERS = ['Player', 'Rating', 'Category', 'Events', 'Matches W', 'Matches L',
                      'Games W', 'Games L', 'Peak', 'First seen', 'Last played'];
var EVENT_HEADERS  = ['Date', 'Event', 'Player', 'Category', 'Start MMR', 'End MMR', 'Change',
                      'Matches W', 'Matches L', 'Games W', 'Games L', 'Perf'];
var MATCH_HEADERS  = ['Date', 'Round', 'Table', 'Player', 'Opponent', 'Result', 'Score',
                      'Games', 'Note', 'MMR before', 'Change', 'MMR after'];

/* ---------- run this once from the editor ---------- */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet_(ss, RATINGS, RATING_HEADERS);
  sheet_(ss, EVENTS, EVENT_HEADERS);
  sheet_(ss, MATCHES, MATCH_HEADERS);
  Logger.log('Ready. Tabs in place on "%s".', ss.getName());
}

/* ---------- web app ---------- */

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'ping';
  try {
    if (action === 'ping')    return json_({ ok: true, version: VERSION, sheet: SpreadsheetApp.getActiveSpreadsheet().getName() });
    if (action === 'ratings') return json_({ ok: true, players: readRatings_() });
    if (action === 'events')  return json_({ ok: true, events: readEvents_() });
    if (action === 'event')   return json_(readEvent_(String(e.parameter.date || '')));
    if (action === 'checkin') return json_(setCheckin_(String(e.parameter.name || ''),
                                                       String(e.parameter.date || ''),
                                                       String(e.parameter.on) === '1'));
    // A payload on the query string still works, for a client that cannot POST.
    if (e.parameter && e.parameter.p) return json_(importEvent_(JSON.parse(e.parameter.p)));
    return json_({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) || (e && e.parameter && e.parameter.p) || '';
    if (!raw) return json_({ ok: false, error: 'No data in the request body.' });
    return json_(importEvent_(JSON.parse(raw)));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ---------- read ---------- */

function readRatings_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = sheet_(ss, RATINGS, RATING_HEADERS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, RATING_HEADERS.length).getValues();
  var out = [];
  vals.forEach(function (r) {
    var name = String(r[0] || '').trim();
    if (!name) return;
    out.push({
      name:     name,
      rating:   num_(r[1]),
      category: String(r[2] || ''),
      events:   num_(r[3]),
      w:        num_(r[4]),
      l:        num_(r[5])
    });
  });
  return out;
}

/** Every event date on record, newest first, for the standings page. */
function readEvents_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = sheet_(ss, EVENTS, EVENT_HEADERS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 2).getValues();
  var seen = {}, out = [];
  vals.forEach(function (r) {
    var d = dateKey_(r[0]);
    if (!d || seen[d]) return;
    seen[d] = true;
    out.push({ date: d, name: String(r[1] || '') });
  });
  out.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });
  return out;
}

/** One event: its final table plus every match row, for the standings page. */
function readEvent_(date) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!date) {
    var all = readEvents_();
    if (!all.length) return { ok: true, date: '', name: '', players: [], matches: [] };
    date = all[0].date;
  }
  var name = '';
  var players = [];
  var esh = sheet_(ss, EVENTS, EVENT_HEADERS);
  var elast = esh.getLastRow();
  if (elast > 1) {
    esh.getRange(2, 1, elast - 1, EVENT_HEADERS.length).getValues().forEach(function (r) {
      if (dateKey_(r[0]) !== date) return;
      name = name || String(r[1] || '');
      players.push({
        name: String(r[2]), category: String(r[3] || ''),
        start: num_(r[4]), end: num_(r[5]), change: num_(r[6]),
        w: num_(r[7]), l: num_(r[8]), gw: num_(r[9]), gl: num_(r[10]),
        perf: r[11] === '' ? null : num_(r[11])
      });
    });
  }
  var matches = [];
  var msh = sheet_(ss, MATCHES, MATCH_HEADERS);
  var mlast = msh.getLastRow();
  if (mlast > 1) {
    msh.getRange(2, 1, mlast - 1, MATCH_HEADERS.length).getValues().forEach(function (r) {
      if (dateKey_(r[0]) !== date) return;
      matches.push({
        round: num_(r[1]), table: r[2] === '' ? null : num_(r[2]),
        player: String(r[3]), opponent: String(r[4]), result: String(r[5]),
        score: String(r[6]), games: String(r[7]), note: String(r[8]),
        before: num_(r[9]), change: num_(r[10]), after: num_(r[11])
      });
    });
  }
  return { ok: true, date: date, name: name, players: players, matches: matches };
}

/* ---------- write ---------- */

/**
 * Tick or untick "Checked in" for one person on the signup tab. Every row with
 * their name for that event is set, so a duplicate signup cannot leave a stale
 * tick behind. Nobody on the tab (a walk-in) is not an error: there is no box.
 */
function setCheckin_(name, date, on) {
  var key = normalize_(name);
  if (!key) return { ok: false, error: 'No name given.' };
  date = dateKey_(date) || todayISO_();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.openById(SIGNUP_SHEET_ID).getSheetByName(SIGNUP_TAB);
    if (!sh) return { ok: false, error: 'No "' + SIGNUP_TAB + '" tab on the signup sheet.' };
    var last = sh.getLastRow(), width = sh.getLastColumn();
    if (last < 2) return { ok: true, version: VERSION, rows: 0 };
    var head = sh.getRange(1, 1, 1, width).getValues()[0].map(function (h) { return String(h).trim(); });
    var cName = -1, cBox = -1, cEv = [];
    head.forEach(function (h, i) {
      if (cName < 0 && /what is your (full )?name/i.test(h)) cName = i;
      if (cBox < 0 && /^checked[\s-]*in$/i.test(h)) cBox = i;
      if (/which event/i.test(h)) cEv.push(i);
    });
    if (cName < 0 || cBox < 0) return { ok: false, error: 'The signup tab needs a name column and a "Checked in" column.' };
    var vals = sh.getRange(2, 1, last - 1, width).getValues();
    var hits = 0;
    vals.forEach(function (r, i) {
      if (normalize_(r[cName]) !== key) return;
      var forDate = cEv.some(function (c) { return signupDate_(r[c], date) === date; });
      if (!forDate) return;
      sh.getRange(i + 2, cBox + 1).setValue(on);
      hits++;
    });
    return { ok: true, version: VERSION, rows: hits };
  } finally {
    lock.releaseLock();
  }
}

/** "Friday, October 23rd" as "2026-10-23", in the year of the given date. */
function signupDate_(v, like) {
  if (v instanceof Date) return dateKey_(v);
  var months = ['january','february','march','april','may','june','july','august',
                'september','october','november','december'];
  var m = /([a-z]+)\s+(\d{1,2})/i.exec(String(v || '').replace(/^[a-z]+,\s*/i, ''));
  if (!m) return '';
  var mi = months.indexOf(m[1].toLowerCase());
  if (mi < 0) return '';
  return String(like).slice(0, 4) + '-' + ('0' + (mi + 1)).slice(-2) + '-' + ('0' + m[2]).slice(-2);
}

function importEvent_(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var date = dateKey_(data.d) || todayISO_();
    var eventName = String(data.ev || 'Riftbound');

    // players: [name, category, startMMR, endMMR, change, W, L, gamesW, gamesL, perf, isJudge]
    var players = (data.p || []).map(function (row) {
      return {
        name:  String(row[0]),
        cat:   String(row[1] || ''),
        start: num_(row[2]),
        end:   num_(row[3]),
        delta: num_(row[4]),
        w:     num_(row[5]),
        l:     num_(row[6]),
        gw:    num_(row[7]),
        gl:    num_(row[8]),
        perf:  row[9] === '' || row[9] === null ? '' : num_(row[9]),
        judge: !!row[10]
      };
    });

    writeEventLog_(ss, date, eventName, players);
    var matchRows = writeMatchLog_(ss, date, players, data.m || []);
    var rated = writeRatings_(ss, date, players);

    return { ok: true, version: VERSION, date: date, event: eventName,
             ratings: rated, eventRows: players.length, matchRows: matchRows };
  } finally {
    lock.releaseLock();
  }
}

/** One row per player per event. Re-running the same date replaces its rows. */
function writeEventLog_(ss, date, eventName, players) {
  var sh = sheet_(ss, EVENTS, EVENT_HEADERS);
  var rows = players.map(function (p) {
    return [date, eventName, p.name, p.judge ? 'Judge' : p.cat, p.start, p.end, p.delta,
            p.w, p.l, p.gw, p.gl, p.perf];
  });
  replaceDate_(sh, EVENT_HEADERS.length, date, rows);
}

/**
 * One row per player per match, in round order, carrying the running rating.
 * matches: [round, table, aIndex, bIndex, result, gamesA, gamesB, deltaA, deltaB]
 * result is "a", "b", "d" for a draw, or "bye".
 */
function writeMatchLog_(ss, date, players, matches) {
  var sh = sheet_(ss, MATCHES, MATCH_HEADERS);

  var running = {};
  players.forEach(function (p, i) { running[i] = p.start; });

  var ordered = matches.slice().sort(function (x, y) { return (x[0] - y[0]) || (x[1] - y[1]); });
  var rows = [];

  ordered.forEach(function (m) {
    var round = num_(m[0]), table = num_(m[1]);
    var ai = m[2], bi = m[3], res = String(m[4]);
    var ga = num_(m[5]), gb = num_(m[6]), da = num_(m[7]), db = num_(m[8]);
    var seq = String(m[9] || '');
    var A = players[ai];
    if (!A) return;

    if (res === 'bye') {
      rows.push([date, round, '', A.name, 'Bye', 'Bye', '', '', 'Sat this round out', running[ai], 0, running[ai]]);
      return;
    }
    var B = players[bi];
    if (!B) return;

    var beforeA = running[ai], beforeB = running[bi];
    running[ai] = round1_(beforeA + da);
    running[bi] = round1_(beforeB + db);

    rows.push([date, round, table, A.name, B.name,
               res === 'a' ? 'Win' : res === 'b' ? 'Loss' : 'Draw',
               ga + '-' + gb, seqFor_(seq, 'a'), noteFor_(res === 'a', ga, gb, seq, 'a'),
               beforeA, da, running[ai]]);
    rows.push([date, round, table, B.name, A.name,
               res === 'b' ? 'Win' : res === 'a' ? 'Loss' : 'Draw',
               gb + '-' + ga, seqFor_(seq, 'b'), noteFor_(res === 'b', gb, ga, seq, 'b'),
               beforeB, db, running[bi]]);
  });

  replaceDate_(sh, MATCH_HEADERS.length, date, rows);
  return rows.length;
}

/**
 * The running ladder. Players in this event get their new rating; new players
 * are appended. Then every player's Events, W/L, games and first/last dates are
 * rebuilt from the Event Log, so re-syncing a night can never double count or
 * wipe an earlier night. Judges carry no rating.
 */
function writeRatings_(ss, date, players) {
  var sh = sheet_(ss, RATINGS, RATING_HEADERS);
  var width = RATING_HEADERS.length;
  var last = sh.getLastRow();
  var rows = last > 1 ? sh.getRange(2, 1, last - 1, width).getValues() : [];

  var at = {};
  rows.forEach(function (r, i) {
    var k = normalize_(r[0]);
    if (k) at[k] = i;
  });

  var written = 0;
  players.forEach(function (p) {
    if (p.judge) return;
    var k = normalize_(p.name);
    if (!k) return;
    var i = at[k];
    if (i === undefined) {
      rows.push([p.name, p.end, p.cat, 0, 0, 0, 0, 0, Math.max(p.start, p.end), date, date]);
      at[k] = rows.length - 1;
    } else {
      rows[i][1] = p.end;
      rows[i][2] = p.cat;
      rows[i][8] = Math.max(num_(rows[i][8]) || p.end, p.end);
    }
    written++;
  });

  /* Lifetime totals, straight from the Event Log. */
  var tot = {};
  var esh = sheet_(ss, EVENTS, EVENT_HEADERS);
  var elast = esh.getLastRow();
  if (elast > 1) {
    esh.getRange(2, 1, elast - 1, EVENT_HEADERS.length).getValues().forEach(function (r) {
      var k = normalize_(r[2]), d = dateKey_(r[0]);
      if (!k || !d || String(r[3]) === 'Judge') return;
      var t = tot[k] || (tot[k] = { ev: {}, w: 0, l: 0, gw: 0, gl: 0, first: d, last: d });
      t.ev[d] = true;
      t.w += num_(r[7]); t.l += num_(r[8]); t.gw += num_(r[9]); t.gl += num_(r[10]);
      if (d < t.first) t.first = d;
      if (d > t.last) t.last = d;
    });
  }
  rows.forEach(function (r) {
    var t = tot[normalize_(r[0])];
    if (!t) return;   // never logged, such as a rating typed in by hand: leave it be
    r[3] = Object.keys(t.ev).length;
    r[4] = t.w; r[5] = t.l; r[6] = t.gw; r[7] = t.gl;
    r[9] = t.first; r[10] = t.last;
  });

  rows.sort(function (a, b) {
    return num_(b[1]) - num_(a[1]) || String(a[0]).localeCompare(String(b[0]));
  });

  if (rows.length) {
    var clearRows = Math.max(rows.length, last - 1);
    if (clearRows > 0) sh.getRange(2, 1, clearRows, width).clearContent();
    sh.getRange(2, 1, rows.length, width).setValues(rows);
  }
  return written;
}

/* ---------- helpers ---------- */

function sheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh && name === RATINGS) {
    // The sheet may already hold Player / Rating on its first, untitled tab.
    var first = ss.getSheets()[0];
    if (String(first.getRange(1, 1).getValue() || '').trim().toLowerCase() === 'player') {
      first.setName(RATINGS);
      sh = first;
    }
  }
  if (!sh) sh = ss.insertSheet(name);

  var head = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(head[i]).trim() !== headers[i]) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
      break;
    }
  }
  return sh;
}

/**
 * Swap out every row for this date in one write: keep the other dates, add the
 * new rows, rewrite the tab. Deleting rows one at a time got slower with every
 * round; this costs the same whether the night has ten rows or five hundred.
 */
function replaceDate_(sh, width, date, rows) {
  var last = sh.getLastRow();
  var keep = last > 1
    ? sh.getRange(2, 1, last - 1, width).getValues().filter(function (r) {
        return dateKey_(r[0]) !== date && String(r[0]) !== '';
      })
    : [];
  var all = keep.concat(rows);
  if (last > 1) sh.getRange(2, 1, last - 1, width).clearContent();
  if (all.length) sh.getRange(2, 1, all.length, width).setValues(all);
}

/** Sheets turns "2026-10-23" into a date cell; read either back as "2026-10-23". */
function dateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').trim().slice(0, 10);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** "aba" becomes "W-L-W" from the given side's point of view. */
function seqFor_(seq, side) {
  if (!seq) return '';
  return seq.split('').map(function (g) {
    return g === side ? 'W' : g === 'd' ? 'D' : 'L';
  }).join('-');
}

/** How the set actually went, in the words a player would use. */
function noteFor_(won, mine, theirs, seq, side) {
  if (mine === theirs) return 'Draw';
  var lostFirst = seq && seq.charAt(0) !== side && seq.charAt(0) !== 'd';
  if (won) {
    if (theirs === 0) return 'Clean sweep: ' + mine + '-0';
    if (lostFirst)    return 'Reverse sweep: ' + mine + '-' + theirs;
    return 'Close game: ' + mine + '-' + theirs;
  }
  if (mine === 0) return 'Landslide: 0-' + theirs;
  if (!lostFirst)  return 'Reverse swept: ' + mine + '-' + theirs;
  return 'Close game: ' + mine + '-' + theirs;
}

function num_(v) { var n = Number(v); return isFinite(n) ? n : 0; }
function round1_(v) { return Math.round(v * 10) / 10; }
function normalize_(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ''); }
function todayISO_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
