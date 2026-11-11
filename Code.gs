/**
 * Book Brawl scoring backend.
 * Lives inside the Google Sheet (Extensions > Apps Script).
 *  - Watches the roster sheet for new book titles and scores them in a "Books" tab.
 *  - Serves the standings as JSON (team names, classes and points only) for the website.
 */

const CONFIG = {
  ROSTER_SHEET: 'Sheet1',       // tab with Period | Team | First | Last | Book #1 | Book #2 ...
  BOOKS_SHEET: 'Books',         // created automatically; this is where rulings are made
  FIRST_BOOK_COLUMN: 5,         // column E
  CACHE_SECONDS: 60,
  FULL_NAMES: false,            // false shows "Jane D."; true shows "Jane Doe" on the public site
};

const MULTIPLIER = { easy: 1, medium: 1.25, hard: 1.5 };
const BOOK_HEADERS = ['Key', 'Title as typed', 'Book', 'Author', 'Words', 'Difficulty', 'Points', 'Note'];

/** Run once from the editor: creates the Books tab and the triggers, then scores existing titles. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  booksSheet_();
  ScriptApp.getProjectTriggers().forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('scoreNewBooks').timeBased().everyMinutes(10).create();
  scoreNewBooks();
}

/** Optional: wipes the Books tab and looks every title up again. This also erases any rulings. */
function rescoreAll() {
  const sh = booksSheet_();
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  scoreNewBooks();
}

/** Installable edit trigger: scores new titles right after they're typed. */
function handleEdit(e) {
  CacheService.getScriptCache().remove('standings');
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() === CONFIG.ROSTER_SHEET &&
      e.range.getLastColumn() >= CONFIG.FIRST_BOOK_COLUMN) {
    scoreNewBooks();
  }
}

/** Finds titles in the roster that aren't in the Books tab yet and estimates them. */
function scoreNewBooks() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const sh = booksSheet_();
    const last = sh.getLastRow();
    const known = new Set(
      last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().flat().map(String).filter(Boolean) : []
    );
    const pending = [...rosterTitles_()].filter(([k]) => !known.has(k));

    for (let i = 0; i < pending.length; i += 20) {
      const batch = pending.slice(i, i + 20);
      const results = estimate_(batch.map(([, raw]) => raw));
      const start = sh.getLastRow() + 1;
      const rows = batch.map(([k, raw], j) => {
        const r = results[j] || {};
        return [k, raw, r.title || '', r.author || '', r.words || '', r.difficulty || 'medium', '', r.note || ''];
      });
      sh.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
      sh.getRange(start, 7, rows.length, 1).setFormulas(rows.map((_, j) => [pointsFormula_(start + j)]));
    }
    if (pending.length) CacheService.getScriptCache().remove('standings');
  } finally {
    lock.releaseLock();
  }
}

/** Web app endpoint for the site. Returns only team names, classes and points. */
function doGet() {
  const cache = CacheService.getScriptCache();
  let json = cache.get('standings');
  if (!json) {
    json = JSON.stringify(standings_());
    cache.put('standings', json, CONFIG.CACHE_SECONDS);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ---------- Scoring ----------

function base_(w) {
  return w <= 50000 ? 1 : w <= 100000 ? 2 : w <= 150000 ? 3 : w <= 200000 ? 4 : w <= 250000 ? 5 : 6;
}

function pointsFormula_(r) {
  return `=IF(E${r}="","",IFS(E${r}<=50000,1,E${r}<=100000,2,E${r}<=150000,3,E${r}<=200000,4,E${r}<=250000,5,TRUE,6)` +
         `*IFS(F${r}="hard",1.5,F${r}="medium",1.25,TRUE,1))`;
}

function round2_(n) { return Math.round(n * 100) / 100; }

function standings_() {
  const ss = SpreadsheetApp.getActive();

  const points = new Map();
  const bs = ss.getSheetByName(CONFIG.BOOKS_SHEET);
  if (bs && bs.getLastRow() > 1) {
    for (const [k, , , , w, d] of bs.getRange(2, 1, bs.getLastRow() - 1, 6).getValues()) {
      const words = Number(w);
      if (!k || !words) continue;
      points.set(String(k), base_(words) * (MULTIPLIER[String(d).toLowerCase().trim()] || 1));
    }
  }

  const classes = new Map();
  const teams = new Map();
  const readers = [];
  const rows = ss.getSheetByName(CONFIG.ROSTER_SHEET).getDataRange().getValues().slice(1);
  for (const row of rows) {
    const period = String(row[0]).trim();
    const first = String(row[2]).trim(), last = String(row[3]).trim();
    if (!period || !(first || last)) continue;

    if (!classes.has(period)) classes.set(period, { name: period, readers: 0, books: 0, total: 0 });
    const c = classes.get(period);
    c.readers++;

    const teamName = String(row[1]).trim();
    let t = null;
    if (teamName) {
      const tk = period + '|' + teamName;
      if (!teams.has(tk)) teams.set(tk, { name: teamName, cls: period, books: 0, points: 0 });
      t = teams.get(tk);
    }
    const reader = { name: displayName_(first, last), cls: period, team: teamName, books: 0, points: 0 };
    readers.push(reader);

    for (let col = CONFIG.FIRST_BOOK_COLUMN - 1; col < row.length; col++) {
      const raw = String(row[col]).trim();
      if (!raw) continue;
      const p = points.get(key_(raw));
      if (p == null) continue;
      c.total += p; c.books++;
      if (t) { t.points += p; t.books++; }
      reader.points += p; reader.books++;
    }
  }

  return {
    updated: new Date().toISOString(),
    classes: [...classes.values()].map((c) => ({
      name: c.name, readers: c.readers, books: c.books,
      total: round2_(c.total), avg: round2_(c.readers ? c.total / c.readers : 0),
    })),
    teams: [...teams.values()].map((t) => ({ name: t.name, cls: t.cls, books: t.books, points: round2_(t.points) })),
    readers: readers.map((r) => ({ ...r, points: round2_(r.points) })),
  };
}

function displayName_(first, last) {
  if (CONFIG.FULL_NAMES) return [first, last].filter(Boolean).join(' ');
  return last ? `${first} ${last.charAt(0).toUpperCase()}.`.trim() : first;
}

// ---------- Sheet helpers ----------

function booksSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.BOOKS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.BOOKS_SHEET);
    sh.getRange(1, 1, 1, BOOK_HEADERS.length).setValues([BOOK_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.hideColumns(1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['easy', 'medium', 'hard'], true).setAllowInvalid(false).build();
    sh.getRange('F2:F').setDataValidation(rule);
  }
  return sh;
}

function key_(s) {
  return String(s).toLowerCase().normalize('NFKD').replace(/[^a-z0-9]+/g, ' ').trim();
}

function rosterTitles_() {
  const values = SpreadsheetApp.getActive().getSheetByName(CONFIG.ROSTER_SHEET).getDataRange().getValues();
  const titles = new Map();
  for (const row of values.slice(1)) {
    for (let c = CONFIG.FIRST_BOOK_COLUMN - 1; c < row.length; c++) {
      const raw = String(row[c]).trim();
      if (!raw) continue;
      const k = key_(raw);
      if (k && !titles.has(k)) titles.set(k, raw);
    }
  }
  return titles;
}

// ---------- Estimating word count and difficulty (free, no API key) ----------
//
// 1. Open Library finds the book and gives the median page count across all its editions
//    (one odd edition can't skew it) plus its subject tags.
// 2. Google Books fills gaps: page count if Open Library has none, and its category labels.
// 3. Page count x words-per-page for the kind of book gives the word count.
// 4. Subjects and categories decide difficulty using the Book Brawl rubric.
// Every row gets a note saying where its numbers came from, so rulings are quick.

const WORDS_PER_PAGE = { graphic: 60, young: 230, standard: 275 };

const PATTERNS = {
  graphic: /graphic novel|comic|manga|cartoon/i,
  young: /juvenile|young adult|children|middle grade|teen|ya fiction|school stories/i,
  academic: /philosophy|history|histor|science|politic|economic|psycholog|sociology|religio|theology|military art|strategy|criticism|essays|biography/i,
  fiction: /fiction|novel|stories|fantasy|romance|mystery|thriller|dystopia/i,
};

function estimate_(titles) {
  return titles.map((t) => {
    try { return lookupBook_(t); }
    catch (err) {
      console.warn('Lookup failed for ' + t + ': ' + err);
      return { title: t, author: '', words: '', difficulty: 'medium', note: 'Lookup failed; enter words' };
    }
  });
}

function splitTitle_(raw) {
  const m = String(raw).match(/^(.*?)\s+by\s+(.+)$/i);
  return m ? { title: m[1].trim(), author: m[2].trim() } : { title: String(raw).trim(), author: '' };
}

function fetchJson_(url) {
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { Accept: 'application/json' } });
  if (res.getResponseCode() !== 200) return null;
  try { return JSON.parse(res.getContentText()); } catch (e) { return null; }
}

function lookupBook_(raw) {
  const q = splitTitle_(raw);
  const ol = openLibrary_(q);
  const gb = googleBooks_(q, ol);

  const title = (ol && ol.title) || (gb && gb.title) || q.title;
  const author = (ol && ol.author) || (gb && gb.author) || q.author;
  const tags = [].concat(ol ? ol.subjects : [], gb ? gb.categories : []).join(' | ');

  const kind = PATTERNS.graphic.test(tags) ? 'graphic' : PATTERNS.young.test(tags) ? 'young' : 'standard';
  let pages = 0, pageSource = '';
  if (ol && ol.pages) { pages = ol.pages; pageSource = 'median of ' + ol.editions + ' editions'; }
  else if (gb && gb.pages) { pages = gb.pages; pageSource = 'Google Books'; }

  const words = pages ? Math.round((pages * WORDS_PER_PAGE[kind]) / 1000) * 1000 : '';
  const difficulty = difficulty_(kind, tags, ol);

  if (!ol && !gb) return { title: q.title, author: q.author, words: '', difficulty: 'medium', note: 'Not found; enter words and difficulty' };
  const note = pages
    ? `${pages} pages (${pageSource}); ${kind === 'standard' ? 'standard' : kind === 'young' ? 'YA/juvenile' : 'graphic'} page density`
    : 'No page count found; enter words';
  return { title, author, words, difficulty, note };
}

function difficulty_(kind, tags, ol) {
  if (kind === 'graphic' || kind === 'young') return 'easy';
  const isFiction = PATTERNS.fiction.test(tags);
  if (!isFiction && PATTERNS.academic.test(tags)) return 'hard';
  // English-language classics written before 1900 count as dense classics.
  if (isFiction && ol && ol.year && ol.year < 1900 && !/translat/i.test(tags)) return 'hard';
  return 'medium';
}

function openLibrary_(q) {
  let url = 'https://openlibrary.org/search.json?limit=8&fields=title,author_name,number_of_pages_median,edition_count,subject,first_publish_year' +
    '&title=' + encodeURIComponent(q.title);
  if (q.author) url += '&author=' + encodeURIComponent(q.author);
  let data = fetchJson_(url);
  if ((!data || !data.docs || !data.docs.length) && q.author) {
    data = fetchJson_(url.replace(/&author=[^&]*/, ''));
  }
  const docs = (data && data.docs) || [];
  if (!docs.length) return null;

  // Prefer real matches on title, then the most-published one (the actual novel, not a study guide).
  const want = key_(q.title);
  const matches = docs.filter((d) => {
    const k = key_(d.title || '');
    return k === want || k.startsWith(want + ' ') || want.startsWith(k + ' ');
  });
  const pool = matches.length ? matches : docs;
  const best = pool.slice().sort((a, b) => (b.edition_count || 0) - (a.edition_count || 0))[0];

  return {
    title: best.title || '',
    author: (best.author_name || [])[0] || '',
    pages: Number(best.number_of_pages_median) || 0,
    editions: best.edition_count || 1,
    subjects: (best.subject || []).slice(0, 40),
    year: Number(best.first_publish_year) || 0,
  };
}

function googleBooks_(q, ol) {
  let query = 'intitle:' + q.title;
  const author = q.author || (ol && ol.author);
  if (author) query += '+inauthor:' + author;
  const data = fetchJson_('https://www.googleapis.com/books/v1/volumes?maxResults=5&printType=books&q=' + encodeURIComponent(query));
  const items = (data && data.items) || [];
  if (!items.length) return null;
  const withPages = items.map((i) => i.volumeInfo || {}).filter((v) => v.pageCount);
  const v = withPages[0] || items[0].volumeInfo || {};
  const categories = [].concat(...items.map((i) => (i.volumeInfo && i.volumeInfo.categories) || []));
  return { title: v.title || '', author: (v.authors || [])[0] || '', pages: Number(v.pageCount) || 0, categories };
}
