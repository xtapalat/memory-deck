/**
 * Memory Deck — Apps Script backend
 *
 * Setup:
 *   1. Create a new Google Sheet.
 *   2. Extensions > Apps Script. Paste this file in, replacing everything.
 *   3. Change TOKEN below to a long random string.
 *   4. Run setup() once (grant permissions when asked).
 *   5. Deploy > New deployment > Web app.
 *        Execute as: Me
 *        Who has access: Anyone
 *      Copy the /exec URL into the app's Settings screen along with your token.
 *
 * Re-deploy (Deploy > Manage deployments > edit > New version) after any edit,
 * or the live URL keeps serving the old code.
 */

var TOKEN = 'CHANGE_ME_TO_A_LONG_RANDOM_STRING';

// Days at each rung. A card climbs, holds, or drops one step at a time.
var LADDER = [1, 3, 7, 14, 30, 60, 120, 240];

var CARDS = 'cards';
var SESSIONS = 'sessions';

var CARD_HEADERS = [
  'id', 'deck', 'front', 'back',
  'created_date', 'created_session',
  'rung', 'interval',
  'due_session', 'due_date',
  'last_reviewed_session', 'last_reviewed_date',
  'review_count', 'status',
  'hard',
  'is_grammar', 'grammar_subtype'
];

// Columns added after the sheet was first built. Each is blank on every row
// that predates it, and blank must read as "no" everywhere — see isTrue_.
var ADDED_COLS = ['hard', 'is_grammar', 'grammar_subtype'];

var GRAMMAR_SUBTYPES = [
  'verb-pattern', 'verb-irregular', 'contraction',
  'clitic', 'agreement', 'diminutive'
];

/**
 * Only the literal word TRUE counts. Blank, FALSE, 0, a stray space, or
 * anything else at all reads as false — which is what makes the backfill a
 * non-event: every row written before these columns existed is blank, and
 * blank is already the right answer for all of them.
 */
function isTrue_(v) {
  return String(v == null ? '' : v).trim().toUpperCase() === 'TRUE';
}

var SESSION_HEADERS = ['deck', 'session_number', 'date'];


/* ------------------------------------------------------------------ setup */

function setup() {
  var ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, CARDS, CARD_HEADERS);
  ensureSheet_(ss, SESSIONS, SESSION_HEADERS);

  Logger.log('Memory Deck is set up. Now deploy as a web app.');
}

/**
 * These columns arrived after the sheet did. Rather than make you run a
 * migration, this writes any missing header the first time anything touches
 * them — the rows below stay blank, which already reads as false.
 *
 * It writes a header only into a cell that is empty, so a sheet where the
 * columns were added by hand, or in another order, is left alone rather than
 * overwritten.
 */
function ensureTagColumns_(sh) {
  ADDED_COLS.forEach(function (name) {
    var col = CARD_HEADERS.indexOf(name) + 1;
    var cur = String(sh.getRange(1, col).getValue()).trim();
    if (cur === '') sh.getRange(1, col).setValue(name).setFontWeight('bold');
  });
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}


/* ----------------------------------------------------------------- routing */

function doGet(e) {
  return handle_(e.parameter || {}, (e.parameter || {}).action);
}

function doPost(e) {
  var body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Body was not valid JSON.' });
  }
  return handle_(body, body.action);
}

function handle_(params, action) {
  if (params.token !== TOKEN) {
    return json_({ ok: false, error: 'Bad token.' });
  }
  try {
    switch (action) {
      case 'state':
        return json_(state_(params.bump === '1' || params.bump === true));
      case 'addCard':
        return json_(addCard_(params));
      case 'addCards':
        return json_(addCards_(params));
      case 'applyReviews':
        return json_(applyReviews_(params.reviews || [], params.batch,
                                   params.sessions || null));
      case 'setSession':
        return json_(setSession_(params));
      case 'reschedule':
        return json_(reschedule_(params));
      case 'setHard':
        return json_(setHard_(params));
      case 'draw':
        return json_(draw_(params));
      case 'sync':
        return json_(sync_());
      default:
        return json_({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/* ---------------------------------------------------------------- sessions */

/**
 * The session clock, now kept per deck. A deck's clock advances only on a day
 * you actually reviewed at least one card in it — opening the app, or working
 * through journal, no longer ages horseshoe or babbel. Miss a fortnight in one
 * deck and that deck waits for you while the others carry on.
 */
function deckSessions_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SESSIONS);
  var last = sh.getLastRow();
  var map = {};
  if (last < 2) return map;

  var rows = sh.getRange(2, 1, last - 1, 3).getValues();
  rows.forEach(function (r) {
    var deck = String(r[0]);
    var num = Number(r[1]);
    var date = asDateString_(r[2]);
    if (!map[deck] || num >= map[deck].num) map[deck] = { num: num, date: date };
  });
  return map;
}

function sessionOf_(map, deck) {
  return map[deck] ? map[deck].num : 1;
}

/**
 * Write the clocks the phone tells us about, verbatim. The phone is the one
 * that decides when a deck moves — it knows whether the day's queue actually
 * emptied, and it knows when you pressed Pull forward — so this is a record,
 * not a decision. A number lower than what is already on file is ignored:
 * clocks do not run backwards, and a stale retry should not drag one back.
 */
function setDecks_(sessions) {
  var map = deckSessions_();
  var t = today_();
  var add = [];

  Object.keys(sessions || {}).forEach(function (deck) {
    var want = sessions[deck];
    var num = Number(want && want.num !== undefined ? want.num : want);
    if (!num) return;
    var date = (want && want.date) ? String(want.date) : t;
    var cur = map[deck];
    if (cur && Number(cur.num) >= num) return;      // nothing new to record
    add.push([deck, num, date]);
  });

  if (add.length) {
    var sh = SpreadsheetApp.getActive().getSheetByName(SESSIONS);
    sh.getRange(sh.getLastRow() + 1, 1, add.length, 3).setValues(add);
  }
  return deckSessions_();
}

/**
 * Tag or untag cards as hard. One column, written in merged runs like every
 * other bulk write here. Tagging changes nothing about scheduling — a hard
 * card comes round on its ladder exactly as before. All the tag does is make
 * the card reachable from Drill hard.
 */
function setHard_(p) {
  var ids = p.ids || [];
  if (!ids.length) return { ok: true, set: 0 };
  if (alreadyDone_(p.batch)) return { ok: true, set: 0, duplicate: true };

  var sh = SpreadsheetApp.getActive().getSheetByName(CARDS);
  ensureTagColumns_(sh);
  var col = CARD_HEADERS.indexOf('hard') + 1;
  var last = sh.getLastRow();
  if (last < 2) return { ok: true, set: 0 };

  var val = Number(p.hard) ? 1 : '';
  var want = {};
  ids.forEach(function (id) { want[String(id)] = 1; });

  var idCol = sh.getRange(2, 1, last - 1, 1).getValues();
  var rows = [];
  for (var i = 0; i < idCol.length; i++) {
    if (want[String(idCol[i][0])]) rows.push(i + 2);
  }
  if (!rows.length) return { ok: true, set: 0, missing: ids };

  var runs = [], cur = null;
  rows.forEach(function (r) {
    if (cur && r === cur.start + cur.n) cur.n++;
    else { cur = { start: r, n: 1 }; runs.push(cur); }
  });
  runs.forEach(function (r) {
    var block = [];
    for (var k = 0; k < r.n; k++) block.push([val]);
    sh.getRange(r.start, col, r.n, 1).setValues(block);
  });

  return { ok: true, set: rows.length, hard: Number(p.hard) ? 1 : 0 };
}

/** One deck's clock, moved by hand from the phone's Pull forward button. */
function setSession_(p) {
  if (!p.deck) return { ok: false, error: 'No deck.' };
  if (alreadyDone_(p.batch)) return { ok: true, duplicate: true };
  var one = {};
  one[p.deck] = { num: Number(p.num), date: p.date || today_() };
  var map = setDecks_(one);
  return { ok: true, deck: p.deck, session: sessionOf_(map, p.deck) };
}

/**
 * Move a set of cards to new days. Used by Smooth, which decides on the phone
 * which card belongs on which day; this only writes the answer down. Columns
 * 9 and 10 sit inside the block writeCardRows_ already handles, so a reshuffle
 * of six hundred cards is a few dozen ranges rather than a full rewrite.
 */
function reschedule_(p) {
  var moves = p.moves || [];
  if (!moves.length) return { ok: true, moved: 0 };
  if (alreadyDone_(p.batch)) return { ok: true, moved: 0, duplicate: true };

  var data = readCardsLean_();
  var byId = {};
  data.rows.forEach(function (c) { byId[String(c.id)] = c; });

  var changed = [], missing = [];
  moves.forEach(function (m) {
    var c = byId[String(m.id)];
    if (!c) { missing.push(m.id); return; }
    c.due_session = Number(m.due_session);
    c.due_date = m.due_date || c.due_date;
    changed.push(c);
  });

  writeCardRows_(data.sheet, changed);
  return { ok: true, moved: changed.length, missing: missing };
}

/**
 * One-time migration from the single global clock to per-deck clocks. Run this
 * once from the editor after pasting this file in. Safe to run twice.
 */
function migrateToPerDeckSessions() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SESSIONS);
  var last = sh.getLastRow();
  var width = sh.getLastColumn();

  if (width >= 3) {
    Logger.log('Already migrated — sessions has a deck column. Nothing to do.');
    return 'already migrated';
  }

  var current = 1;
  if (last >= 2) {
    var rows = sh.getRange(2, 1, last - 1, 2).getValues();
    rows.forEach(function (r) { current = Math.max(current, Number(r[0]) || 1); });
  }

  // Every existing due_session was written against the old global clock, so
  // each deck inherits that same number and nothing shifts underneath you.
  var cards = ss.getSheetByName(CARDS);
  var decks = {};
  if (cards.getLastRow() >= 2) {
    cards.getRange(2, 2, cards.getLastRow() - 1, 1).getValues()
      .forEach(function (r) { if (r[0]) decks[r[0]] = true; });
  }

  var t = today_();
  var out = [SESSION_HEADERS];
  Object.keys(decks).sort().forEach(function (d) { out.push([d, current, t]); });

  sh.clear();
  sh.getRange(1, 1, out.length, 3).setValues(out);
  // Logger, never getUi().alert(). An alert opens a modal in the spreadsheet
  // tab and blocks the run until someone clicks it, which looks exactly like
  // a hung migration when you are watching the editor instead.
  var note = 'Migrated ' + (out.length - 1) + ' decks to their own clocks, all at session ' + current + '.';
  Logger.log(note);
  return note;
}


/**
 * Run this from the editor and read the execution log. It reports cards whose
 * review_count looks higher than the number of times you actually graded them,
 * which is the fingerprint of a retried save that landed more than once.
 */
function inspectDeck(deck) {
  deck = deck || 'portugese potter';
  var rows = readCardsLean_().rows.filter(function (c) { return c.deck === deck; });
  var t = today_();

  var byCount = {}, todays = [];
  rows.forEach(function (c) {
    var n = Number(c.review_count) || 0;
    byCount[n] = (byCount[n] || 0) + 1;
    if (c.last_reviewed_date === t) todays.push(c);
  });

  Logger.log('deck: ' + deck + '  ·  ' + rows.length + ' cards');
  Logger.log('reviewed today: ' + todays.length);
  Object.keys(byCount).sort(function (a, b) { return a - b; }).forEach(function (k) {
    Logger.log('  review_count ' + k + ': ' + byCount[k] + ' cards');
  });

  var suspect = todays.filter(function (c) { return Number(c.review_count) > 1; });
  if (suspect.length) {
    Logger.log('SUSPECT — graded today but review_count > 1, so a save landed more than once:');
    suspect.slice(0, 25).forEach(function (c) {
      Logger.log('  ' + c.id + '  count=' + c.review_count + '  rung=' + c.rung +
                 '  interval=' + c.interval + '  due=' + c.due_date);
    });
  } else {
    Logger.log('No double-application found. Nothing to repair.');
  }
  return suspect.length;
}

/**
 * Undo n extra applications for cards graded today in a deck: walks the rung
 * back down and recomputes the interval. Run inspectDeck first, and pass the
 * number of EXTRA times the batch landed (e.g. 2 if review_count reads 3 where
 * it should read 1).
 */
function repairDeck(deck, extra) {
  deck = deck || 'portugese potter';
  extra = Number(extra) || 0;
  if (extra < 1) { Logger.log('Pass the number of extra applications, e.g. repairDeck("portugese potter", 2)'); return; }

  var data = readCardsLean_();
  var t = today_();
  var changed = [];

  data.rows.forEach(function (c) {
    if (c.deck !== deck || c.last_reviewed_date !== t) return;
    var rung = Math.max(0, (Number(c.rung) || 0) - extra);
    c.rung = rung;
    c.interval = LADDER[rung];
    c.due_session = Number(c.due_session) - (LADDER[Number(c.rung) || 0] || 0) + LADDER[rung];
    c.due_date = addDays_(t, LADDER[rung]);
    c.review_count = Math.max(0, (Number(c.review_count) || 0) - extra);
    changed.push(c);
  });

  writeCardRows_(data.sheet, changed);
  Logger.log('Rolled back ' + changed.length + ' cards in ' + deck + ' by ' + extra + ' application(s).');
  return changed.length;
}


/* ------------------------------------------------------------------- cards */

function readCards_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(CARDS);
  var last = sh.getLastRow();
  if (last < 2) return { sheet: sh, rows: [], values: [] };

  var values = sh.getRange(2, 1, last - 1, CARD_HEADERS.length).getValues();
  var rows = values.map(function (v, i) {
    var o = {};
    CARD_HEADERS.forEach(function (h, j) { o[h] = v[j]; });
    o.created_date = asDateString_(o.created_date);
    o.due_date = asDateString_(o.due_date);
    o.last_reviewed_date = asDateString_(o.last_reviewed_date);
    o._row = i + 2;
    return o;
  });
  return { sheet: sh, rows: rows, values: values };
}

/**
 * Counts only, and deliberately so. This runs every time the app opens, and
 * the two costly columns in the sheet are front and back — the card text. We
 * skip them entirely here and read only what it takes to count: deck,
 * due_session, due_date, status. The actual cards arrive later, from draw_(),
 * once you have said how many you want.
 */
function state_(bump) {
  var t = today_();
  var sessions = deckSessions_();
  var sh = SpreadsheetApp.getActive().getSheetByName(CARDS);
  var last = sh.getLastRow();

  var decks = {};
  if (last >= 2) {
    var n = last - 1;
    var deckCol = sh.getRange(2, 2, n, 1).getValues();
    var tail = sh.getRange(2, 9, n, 6).getValues();   // due_session .. status

    for (var i = 0; i < n; i++) {
      if (String(tail[i][5]) !== 'active') continue;
      var deck = deckCol[i][0];
      if (!decks[deck]) {
        decks[deck] = {
          deck: deck, due: 0, ahead: 0, total: 0,
          session: sessionOf_(sessions, deck)
        };
      }
      decks[deck].total++;

      var dueSession = Number(tail[i][0]);
      var reviewed = asDateString_(tail[i][3]);        // last_reviewed_date

      // The session clock is the only clock; due_date is not consulted.
      if (reviewed === t) decks[deck].ahead++;         // done today, not offered again
      else if (dueSession <= decks[deck].session) decks[deck].due++;
      else decks[deck].ahead++;
    }
  }

  return {
    ok: true,
    today: t,
    ladder: LADDER,
    decks: Object.keys(decks).sort().map(function (k) { return decks[k]; })
  };
}

/**
 * The whole collection in one call. The phone keeps a local mirror and does
 * all its scheduling against that, so this runs in the background rather than
 * in front of you. At ~1,500 cards it is a couple of hundred kilobytes.
 */
function sync_() {
  var data = readCards_();
  return {
    ok: true,
    today: today_(),
    ladder: LADDER,
    sessions: deckSessions_(),
    // So Settings can offer a way straight into the sheet when you want to
    // fix something by hand rather than through the app.
    sheetUrl: SpreadsheetApp.getActive().getUrl(),
    cards: data.rows.filter(function (c) { return String(c.status) === 'active'; })
      .map(function (c) {
        return {
          id: c.id, deck: c.deck, front: c.front, back: c.back,
          rung: Number(c.rung) || 0,
          interval: Number(c.interval) || 1,
          due_session: Number(c.due_session) || 1,
          due_date: c.due_date,
          created_date: c.created_date,
          last_reviewed_date: c.last_reviewed_date,
          review_count: Number(c.review_count) || 0,
          hard: Number(c.hard) ? 1 : 0,
          is_grammar: isTrue_(c.is_grammar) ? 1 : 0,
          grammar_subtype: String(c.grammar_subtype || '').trim()
        };
      })
  };
}

/**
 * Hand back exactly n cards for a deck, already in the order you want to see
 * them. Priority: what the session clock owes you, then cards not due yet.
 * There is deliberately no third band for cards the wall calendar has run
 * past — due_date is a note of when a card was last scheduled, not a deadline,
 * and treating it as one is how a fortnight away becomes an unpayable debt.
 * Ask for more than are due and it keeps going into the next days' worth.
 * Must match localDraw() in index.html.
 */
function draw_(p) {
  var t = today_();
  var sessions = deckSessions_();
  var data = readCards_();
  var deck = p.deck || '';
  var n = Math.max(1, Math.min(300, Number(p.n) || 10));

  var due = [], ahead = [];
  data.rows.forEach(function (c) {
    if (String(c.status) !== 'active') return;
    if (deck && c.deck !== deck) return;
    // Reviewed already today. Pulling ahead should reach into tomorrow, not
    // hand you the same card twice in one evening.
    if (c.last_reviewed_date === t) return;

    if (Number(c.due_session) <= sessionOf_(sessions, c.deck)) due.push(c);
    else ahead.push(c);
  });

  var newestFirst = function (a, b) {
    var x = String(a.created_date || ''), y = String(b.created_date || '');
    if (x !== y) return x < y ? 1 : -1;
    return String(a.id) < String(b.id) ? 1 : -1;
  };

  // Must match localDraw() in index.html.
  var inFlightFirst = function (a, b) {
    var d = (Number(a.interval) || 1) - (Number(b.interval) || 1);
    if (d) return d;
    var x = String(a.last_reviewed_date || ''), y = String(b.last_reviewed_date || '');
    if (x !== y) return x < y ? -1 : 1;
    return String(a.id) < String(b.id) ? -1 : 1;
  };
  var hasSeen = function (c) { return !!c.last_reviewed_date; };
  var notSeen = function (c) { return !c.last_reviewed_date; };

  var seen = due.filter(hasSeen).sort(inFlightFirst);
  var fresh = due.filter(notSeen).sort(newestFirst);

  ahead.sort(function (a, b) {
    var d = Number(a.due_session) - Number(b.due_session);
    return d ? d : newestFirst(a, b);
  });

  var pool = seen.concat(fresh, ahead).slice(0, n);

  return {
    ok: true,
    today: t,
    scheduled: due.length,
    available: due.length + ahead.length,
    cards: pool.map(function (c) {
      return {
        id: c.id, deck: c.deck, front: c.front, back: c.back,
        rung: Number(c.rung), interval: Number(c.interval),
        review_count: Number(c.review_count),
        created_date: c.created_date,
        last_reviewed_date: c.last_reviewed_date,
        due_session: Number(c.due_session), due_date: c.due_date,
        hard: Number(c.hard) ? 1 : 0,
        is_grammar: isTrue_(c.is_grammar) ? 1 : 0,
        grammar_subtype: String(c.grammar_subtype || '').trim(),
        ahead: Number(c.due_session) > sessionOf_(sessions, c.deck)
      };
    })
  };
}

function addCard_(p) {
  // The phone owns the clock, so take its number when it sends one. Without
  // this a card written while a clock move is still sitting in the outbox
  // lands a day behind where the phone just showed it.
  var session = Number(p.session) || sessionOf_(deckSessions_(), p.deck || 'journal');
  var t = today_();
  var sh = SpreadsheetApp.getActive().getSheetByName(CARDS);

  var deck = String(p.deck || 'journal').trim();
  var front = String(p.front || '').trim();
  var back = String(p.back || '').trim();
  if (!front) return { ok: false, error: 'A card needs something on the front.' };

  var id = nextId_(sh);

  // New cards enter at the bottom rung and come back tomorrow, not tonight.
  // The three trailing tag columns are left blank, which reads as not hard
  // and not grammar — a card typed in the app is vocabulary by default.
  sh.appendRow([
    id, deck, front, back,
    t, session,
    0, LADDER[0],
    session + 1, addDays_(t, LADDER[0]),
    '', '',
    0, 'active',
    '', '', ''
  ]);

  return { ok: true, id: id, session: session };
}

/**
 * A batch of cards from the import screen, written in one go.
 *
 * Two things matter here. First, the ids are assigned on this end: the column
 * is scanned once and counted up locally, so a hundred cards cost one read
 * rather than a hundred, and no other Claude session has to be trusted to
 * invent numbers that do not already exist.
 *
 * Second, the new cards are spread forward at `per` a day starting tomorrow.
 * Dropping eighty cards onto a single day would stall the deck outright —
 * nothing advances until a day is finished, and a day of eighty rarely is.
 */
function addCards_(p) {
  var cards = p.cards || [];
  if (!cards.length) return { ok: false, error: 'No cards in that block.' };
  if (alreadyDone_(p.batch)) return { ok: true, added: 0, duplicate: true };

  var deck = String(p.deck || 'journal').trim();
  var per = Math.max(1, Number(p.per) || 10);
  var session = Number(p.session) || sessionOf_(deckSessions_(), deck);
  var t = today_();
  var sh = SpreadsheetApp.getActive().getSheetByName(CARDS);

  // One scan of the id column, then count up in memory.
  var last = sh.getLastRow(), max = 0;
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
      var m = /^c(\d+)$/.exec(String(r[0]));
      if (m) max = Math.max(max, Number(m[1]));
    });
  }

  var rows = [], ids = [], skipped = 0;
  cards.forEach(function (c) {
    var front = String(c.front || '').trim();
    var back = String(c.back || '').trim();
    if (!front) { skipped++; return; }

    var offset = 1 + Math.floor(rows.length / per);   // day 1 is tomorrow
    var id = 'c' + pad_(++max, 4);
    ids.push(id);

    // A card may carry its own tag; otherwise it inherits the block's.
    var gsub = String(c.grammar_subtype || p.grammar_subtype || '').trim();
    var gram = (c.is_grammar === undefined)
      ? (isTrue_(p.is_grammar) || !!gsub)
      : isTrue_(c.is_grammar);
    if (!gram) gsub = '';

    rows.push([
      id, deck, front, back,
      t, session,
      0, LADDER[0],
      session + offset, addDays_(t, offset),
      '', '',
      0, 'active',
      '', gram ? 'TRUE' : '', gsub
    ]);
  });

  if (!rows.length) return { ok: false, error: 'Every card was missing a front.' };

  ensureTagColumns_(sh);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

  return {
    ok: true, added: rows.length, skipped: skipped, deck: deck,
    first: ids[0], last: ids[ids.length - 1],
    days: Math.ceil(rows.length / per), per: per, session: session
  };
}

function nextId_(sh) {
  var last = sh.getLastRow();
  var max = 0;
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
      var m = /^c(\d+)$/.exec(String(r[0]));
      if (m) max = Math.max(max, Number(m[1]));
    });
  }
  return 'c' + pad_(max + 1, 4);
}

/**
 * reviews: [{ id: 'c0142', result: 'again' | 'fuzzy' | 'got_it' | 'easy' }]
 *
 * The rung you are on carries the history, so each answer moves you relative
 * to where you already were rather than restarting from a review count.
 */

/**
 * A lean read for grading: skips front and back, which are the bulk of the
 * sheet and irrelevant to scheduling. Two narrow reads instead of one wide one.
 */
function readCardsLean_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(CARDS);
  var last = sh.getLastRow();
  if (last < 2) return { sheet: sh, rows: [] };
  var n = last - 1;

  var head = sh.getRange(2, 1, n, 2).getValues();    // id, deck
  var tail = sh.getRange(2, 7, n, 8).getValues();    // rung .. status

  var rows = [];
  for (var i = 0; i < n; i++) {
    rows.push({
      id: head[i][0], deck: head[i][1],
      rung: tail[i][0], interval: tail[i][1],
      due_session: tail[i][2], due_date: asDateString_(tail[i][3]),
      last_reviewed_session: tail[i][4],
      last_reviewed_date: asDateString_(tail[i][5]),
      review_count: tail[i][6], status: tail[i][7],
      _row: i + 2
    });
  }
  return { sheet: sh, rows: rows };
}

/**
 * Write back only the rows that changed, and only columns 7-13 (rung through
 * review_count), which happen to be contiguous. Adjacent rows are merged into
 * single ranges. Grading twenty cards used to rewrite every row and every
 * column in the sheet; this touches a few dozen cells.
 */
function writeCardRows_(sheet, cards) {
  if (!cards.length) return;
  cards.sort(function (a, b) { return a._row - b._row; });

  var runs = [], cur = null;
  cards.forEach(function (c) {
    if (cur && c._row === cur.start + cur.vals.length) {
      cur.vals.push(rowSlice_(c));
    } else {
      cur = { start: c._row, vals: [rowSlice_(c)] };
      runs.push(cur);
    }
  });

  runs.forEach(function (r) {
    sheet.getRange(r.start, 7, r.vals.length, 7).setValues(r.vals);
  });
}

function rowSlice_(c) {
  return [c.rung, c.interval, c.due_session, c.due_date,
          c.last_reviewed_session, c.last_reviewed_date, c.review_count];
}

/**
 * Reject a batch we have already processed. A hung request that the phone
 * gives up on may well have succeeded on this end, and the retry would
 * otherwise grade every card a second time and push it a second rung up.
 */
function alreadyDone_(batch) {
  if (!batch) return false;
  var props = PropertiesService.getScriptProperties();
  var seen = (props.getProperty('batches') || '').split(',');
  if (seen.indexOf(batch) !== -1) return true;
  seen.push(batch);
  props.setProperty('batches', seen.slice(-40).join(','));
  return false;
}

function applyReviews_(reviews, batch, clientSessions) {
  if (!reviews.length) return { ok: true, applied: 0 };
  if (alreadyDone_(batch)) return { ok: true, applied: 0, duplicate: true };

  var t = today_();
  var data = readCardsLean_();
  var byId = {};
  data.rows.forEach(function (c) { byId[String(c.id)] = c; });

  var touched = {};
  reviews.forEach(function (r) {
    var c = byId[String(r.id)];
    if (c) touched[c.deck] = true;
  });

  // Grading no longer advances anything. The clock moves when a deck's queue
  // empties and the next day arrives, and the phone is what notices that —
  // so it sends where each deck stands and we record it. The old behaviour
  // (bump on any review) is why a half-finished night still aged the deck.
  var sessions = clientSessions ? setDecks_(clientSessions) : deckSessions_();

  var applied = 0, missing = [], changed = [];

  reviews.forEach(function (r) {
    var c = byId[String(r.id)];
    if (!c) { missing.push(r.id); return; }

    var session = sessionOf_(sessions, c.deck);
    var rung = Number(c.rung) || 0;
    var result = String(r.result || '').toLowerCase();

    if (result === 'again') {
      // Stays in today's pile. No movement on the ladder, no credit taken.
      c.due_session = session;
      c.due_date = t;
    } else {
      if (result === 'fuzzy') rung = rung - 1;
      else if (result === 'got_it') rung = rung + 1;
      else if (result === 'easy') rung = rung + 2;
      else return; // unrecognised, leave the card alone

      rung = Math.max(0, Math.min(LADDER.length - 1, rung));
      var interval = LADDER[rung];

      c.rung = rung;
      c.interval = interval;
      c.due_session = session + interval;
      c.due_date = addDays_(t, interval);
      c.last_reviewed_session = session;
      c.last_reviewed_date = t;
      c.review_count = (Number(c.review_count) || 0) + 1;
    }

    changed.push(c);
    applied++;
  });

  writeCardRows_(data.sheet, changed);

  var out = {};
  Object.keys(touched).forEach(function (d) { out[d] = sessionOf_(sessions, d); });
  return { ok: true, applied: applied, missing: missing, sessions: out };
}


/* ------------------------------------------------------------------ dates */

function tz_() {
  return SpreadsheetApp.getActive().getSpreadsheetTimeZone();
}

function today_() {
  return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
}

function addDays_(dateStr, n) {
  var p = dateStr.split('-');
  var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  d.setDate(d.getDate() + Number(n));
  return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd');
}

function asDateString_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  }
  return String(v);
}

function pad_(n, width) {
  var s = String(n);
  while (s.length < width) s = '0' + s;
  return s;
}
