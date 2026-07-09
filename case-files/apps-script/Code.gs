/**
 * Case Files backend — extends the existing doGet (activity logging to the
 * "Log" tab) with:
 *   - serving case/quote/fact data as JSON from a "Cases" tab, so content
 *     no longer has to be hardcoded in the HTML
 *   - a "Pending" tab + doPost endpoint that a daily research job posts new
 *     candidate entries into, for human review before anything goes live
 *
 *   GET  ?action=getCases                        -> live case data as JSON
 *   GET  ?user=...&type=...&id=...&title=...      -> log activity (as before)
 *   GET  (no params)                              -> raw Log sheet dump (as before)
 *   POST ?action=addPending  body: {"items":[...]} -> queue new entries for review
 *
 * Behavior changes from the original doGet: activity logging now wraps the
 * sheet write in a LockService lock and flushes explicitly, so concurrent
 * requests (e.g. fast double-clicks) can't race and silently drop a row —
 * that race was the likely cause of log rows going missing intermittently.
 *
 * Setup:
 *   1. Extensions > Apps Script on the Sheet, replace contents with this file.
 *   2. Run `seedCases` once from the editor (top toolbar function picker),
 *      authorize when prompted. This creates the "Cases" tab and fills it.
 *   3. Deploy > Manage deployments > pick the existing deployment > Edit
 *      (pencil icon) > New version > Deploy. This is required — editing the
 *      script does NOT update the live /exec URL until you cut a new version.
 *   4. Reload the Sheet once so the "Case Files" menu (onOpen) appears —
 *      that's how you publish or reject items sitting in "Pending".
 *
 * Reviewing new content:
 *   The daily job only ever writes to the "Pending" tab, status='pending'.
 *   In that tab, set a row's status to "approved" to publish it live, or
 *   "rejected" to discard it — then run Case Files > Publish approved
 *   pending items from the Sheet's menu (or wait for the nightly trigger,
 *   if you set one up) to apply those decisions.
 */

var CASES_SHEET_NAME = 'Cases';
var PENDING_SHEET_NAME = 'Pending';
var LOG_SHEET_NAME = 'Log';
var CASES_COLUMNS = ['section', 'id', 'title', 'year', 'json'];
var PENDING_COLUMNS = ['section', 'id', 'title', 'year', 'json', 'status', 'dateAdded'];

function doGet(e) {
  var params = (e && e.parameter) || {};
  if (params.action === 'getCases') {
    return getCasesResponse(params.user || '');
  }
  if (params.type) {
    return logActivityResponse(params);
  }
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  var rows = sheet2 ? sheet2.getDataRange().getValues() : [];
  return ContentService.createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var params = (e && e.parameter) || {};
  if (params.action === 'addPending') {
    return addPendingResponse(e);
  }
  return ContentService.createTextOutput(JSON.stringify({ error: 'unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getCasesResponse(user) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CASES_SHEET_NAME);
  var result = { solved: [], open: [], problems: [], quotes: [], facts: [], viewed: [] };

  if (sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var rows = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      for (var i = 0; i < rows.length; i++) {
        var section = rows[i][0];
        var json = rows[i][4];
        if (!json) continue;
        var obj;
        try {
          obj = JSON.parse(json);
        } catch (err) {
          continue;
        }
        if (section === 'solved') result.solved.push(obj);
        else if (section === 'teen') result.open.push(obj);
        else if (section === 'problem') result.problems.push(obj);
        else if (section === 'quote') result.quotes.push(obj);
        else if (section === 'fact') result.facts.push(obj);
      }
    }
  }

  if (user) {
    result.viewed = getViewedIds(user);
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

var VIEW_ACTIVITY_TYPES = { closed_case: 1, teen_case: 1, problem_viewed: 1, problem_picked: 1 };

/**
 * Every card-open click is already logged to the "Log" tab via
 * logActivityResponse (type = closed_case | teen_case | problem_viewed).
 * Rather than tracking "already viewed" separately (e.g. in localStorage,
 * which is per-browser and doesn't follow a user across devices), this
 * reads that existing log back out, filtered to one user's view-type
 * events, so "don't show me cards I've already opened" works from
 * whichever device/browser the user is on.
 */
function getViewedIds(user) {
  var ids = [];
  var seen = {};
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  if (!sheet) return ids;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return ids;
  var rows = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  rows.forEach(function (row) {
    var rowUser = row[1];
    var type = row[2];
    var id = row[3];
    if (rowUser === user && VIEW_ACTIVITY_TYPES[type] && id && !seen[id]) {
      seen[id] = true;
      ids.push(id);
    }
  });
  return ids;
}

function logActivityResponse(params) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(LOG_SHEET_NAME);
      sheet.appendRow(['Timestamp', 'User', 'Type', 'Item ID', 'Item Title']);
    }
    sheet.appendRow([
      new Date(),
      params.user || '',
      params.type || '',
      params.id || '',
      params.title || ''
    ]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function slugify(str) {
  return String(str)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/(^-|-$)/g, '')
    .slice(0, 40);
}

/**
 * One-time setup: creates the "Cases" tab (if missing) and fills it with
 * the current case/quote/fact data, one row per entry:
 *   section ('solved'|'teen'|'problem'|'quote'|'fact') | id | title | year | json
 * Re-running this wipes and re-seeds the tab from scratch — safe to run
 * again after editing the SEED_* arrays below, but any hand-edits made
 * directly in the Sheet (including anything published from Pending) in
 * the meantime will be lost.
 */
function seedCases() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CASES_SHEET_NAME);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet(CASES_SHEET_NAME);
  sheet.appendRow(CASES_COLUMNS);

  var rows = [];
  SEED_SOLVED.forEach(function (c) {
    rows.push(['solved', c.id, c.title, c.year || '', JSON.stringify(c)]);
  });
  SEED_OPEN.forEach(function (c) {
    rows.push(['teen', c.id, c.title, '', JSON.stringify(c)]);
  });
  SEED_PROBLEMS.forEach(function (c) {
    rows.push(['problem', c.id, c.title, '', JSON.stringify(c)]);
  });
  SEED_QUOTES.forEach(function (q) {
    var id = 'quote-' + slugify(q.text);
    var withId = { id: id, text: q.text, attr: q.attr, url: q.url || '' };
    rows.push(['quote', id, q.attr, '', JSON.stringify(withId)]);
  });
  SEED_FACTS.forEach(function (f) {
    var id = 'fact-' + slugify(f.text);
    rows.push(['fact', id, f.text.slice(0, 60), '', JSON.stringify({ id: id, text: f.text, url: f.url || '' })]);
  });

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }
  SpreadsheetApp.flush();
}

/**
 * Returns a Set-like object of every id currently present in Cases + Pending,
 * so the daily research job (and addPendingResponse) can skip duplicates.
 */
function getExistingIds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ids = {};
  [CASES_SHEET_NAME, PENDING_SHEET_NAME].forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    var idCol = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    idCol.forEach(function (row) {
      if (row[0]) ids[row[0]] = true;
    });
  });
  return ids;
}

/**
 * POST ?action=addPending, body: {"items":[{section,id,title,year,data},...]}
 * Queues new candidate entries (from the daily research job) into the
 * "Pending" tab with status='pending'. Existing ids (already in Cases or
 * Pending) are skipped rather than duplicated. Nothing here ever touches
 * the live "Cases" tab directly — publishing is a separate, human-triggered
 * step (see publishApprovedPending).
 */
function addPendingResponse(e) {
  var body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'invalid JSON body' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var items = Array.isArray(body.items) ? body.items : [];
  var validSections = { solved: 1, teen: 1, problem: 1, quote: 1, fact: 1 };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  var added = [];
  var skipped = [];
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PENDING_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(PENDING_SHEET_NAME);
      sheet.appendRow(PENDING_COLUMNS);
    }
    var existingIds = getExistingIds();
    var rows = [];
    items.forEach(function (item) {
      if (!item || !item.id || !item.section || !validSections[item.section]) {
        skipped.push((item && item.id) || '(missing id)');
        return;
      }
      if (existingIds[item.id]) {
        skipped.push(item.id);
        return;
      }
      existingIds[item.id] = true;
      added.push(item.id);
      rows.push([
        item.section,
        item.id,
        item.title || '',
        item.year || '',
        JSON.stringify(item.data || item),
        'pending',
        new Date()
      ]);
    });
    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, PENDING_COLUMNS.length).setValues(rows);
      SpreadsheetApp.flush();
    }
  } finally {
    lock.releaseLock();
  }

  return ContentService.createTextOutput(JSON.stringify({ added: added, skipped: skipped }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Menu-driven publish step. In the "Pending" tab, set a row's status to
 * "approved" to move it into the live "Cases" tab, or "rejected" to discard
 * it. Rows left as "pending" are untouched. Run from the Sheet's
 * Case Files > Publish approved pending items menu.
 */
function publishApprovedPending() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pending = ss.getSheetByName(PENDING_SHEET_NAME);
  if (!pending) return;
  var lastRow = pending.getLastRow();
  if (lastRow < 2) return;

  var data = pending.getRange(2, 1, lastRow - 1, PENDING_COLUMNS.length).getValues();
  var cases = ss.getSheetByName(CASES_SHEET_NAME);
  if (!cases) {
    cases = ss.insertSheet(CASES_SHEET_NAME);
    cases.appendRow(CASES_COLUMNS);
  }

  var toPublish = [];
  var rowsToDelete = [];
  data.forEach(function (row, i) {
    var status = String(row[5] || '').toLowerCase();
    if (status === 'approved') {
      toPublish.push([row[0], row[1], row[2], row[3], row[4]]);
      rowsToDelete.push(i + 2);
    } else if (status === 'rejected') {
      rowsToDelete.push(i + 2);
    }
  });

  if (toPublish.length) {
    cases.getRange(cases.getLastRow() + 1, 1, toPublish.length, CASES_COLUMNS.length).setValues(toPublish);
  }
  rowsToDelete
    .sort(function (a, b) { return b - a; })
    .forEach(function (rowNum) { pending.deleteRow(rowNum); });

  SpreadsheetApp.flush();

  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (err) { ui = null; }
  if (ui) {
    ui.alert('Published ' + toPublish.length + ' item(s), removed ' + rowsToDelete.length + ' row(s) from Pending.');
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Case Files')
    .addItem('Publish approved pending items', 'publishApprovedPending')
    .addToUi();
}

/**
 * Simple trigger — fires automatically on any edit to the Sheet, no setup
 * required. When you type "approved" into a row's status cell in the
 * "Pending" tab, that row is immediately moved into the live "Cases" tab.
 * Typing "rejected" discards it. Handles pasting into multiple status
 * cells at once too. publishApprovedPending() above still exists as a
 * manual batch fallback (e.g. to catch anything this trigger missed).
 */
function onEdit(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== PENDING_SHEET_NAME) return;

  var statusCol = PENDING_COLUMNS.indexOf('status') + 1;
  if (statusCol < e.range.getColumn() || statusCol > e.range.getLastColumn()) return;

  var startRow = Math.max(e.range.getRow(), 2);
  var endRow = e.range.getLastRow();
  if (startRow > endRow) return;

  var ss = e.source;
  var cases = ss.getSheetByName(CASES_SHEET_NAME);
  if (!cases) {
    cases = ss.insertSheet(CASES_SHEET_NAME);
    cases.appendRow(CASES_COLUMNS);
  }

  for (var row = endRow; row >= startRow; row--) {
    var status = String(sheet.getRange(row, statusCol).getValue() || '').toLowerCase().trim();
    if (status === 'approved') {
      var rowData = sheet.getRange(row, 1, 1, CASES_COLUMNS.length).getValues()[0];
      cases.appendRow(rowData);
      sheet.deleteRow(row);
    } else if (status === 'rejected') {
      sheet.deleteRow(row);
    }
  }
}

// ---------------------------------------------------------------------------
// Seed data — copied from the solved / open / problems arrays that used to
// live inline in case-file.html. Edit here and re-run seedCases() to update
// the Sheet, or just edit rows directly in the "Cases" tab's json column.
// ---------------------------------------------------------------------------

var SEED_SOLVED = [
  {
    id: 'apollo13', year: '1970',
    title: "Astronauts turned a disaster into a safe landing",
    teaser: "An oxygen tank blew up 200,000 miles from home.",
    story: [
      "Apollo 13 was two days into a Moon mission when a routine tank stir caused an oxygen tank to explode, killing power, water, and air in the main spacecraft.",
      "The crew crammed into the tiny lunar module Aquarius — built for 2 people for 2 days — and stretched it to keep 3 people alive for 4.",
      "When the air filters didn't fit, Mission Control built a working adapter out of duct tape, a plastic bag, and cardboard, then talked the crew through building the same thing in space.",
      "Instead of turning around, they let the ship loop around the Moon and used its gravity like a slingshot to fling themselves home, saving precious fuel. All three astronauts splashed down safely four days later."
    ],
    diagram: 'flow', steps: [['tank explodes', '200,000 miles out'], ['move to lifeboat', 'crowd into Aquarius'], ['jury-rig air filter', 'tape, bags, cardboard'], ['splash down safe', '4 days later']],
    source: "NASA, 'Apollo 13: Mission Details'; Smithsonian National Air and Space Museum; Wikipedia, cross-checked against NASA records.",
    links: [
      { label: "NASA — Apollo 13: Mission Details", url: "https://www.nasa.gov/missions/apollo/apollo-13-mission-details/" },
      { label: "Smithsonian National Air and Space Museum — Apollo 13", url: "https://airandspace.si.edu/explore/stories/apollo-missions/apollo-13" },
      { label: "Wikipedia — Apollo 13", url: "https://en.wikipedia.org/wiki/Apollo_13" }
    ]
  },
  {
    id: 'smallpox', year: '1980',
    title: "One vaccine ended a disease that terrified kings",
    teaser: "It killed 300 million people in the 20th century alone.",
    story: [
      "Smallpox had plagued humanity for at least 3,000 years, killing roughly 1 in 3 people it infected and scarring survivors for life.",
      "In 1967 the World Health Organization launched an all-out global campaign, sending health workers into deserts, jungles, and war zones with a new two-pronged needle that used far less vaccine per dose.",
      "Their strategy was almost detective-like: 'ring vaccination.' The moment a case appeared anywhere, teams raced in and vaccinated everyone around that person, stamping out the fire before it could spread.",
      "On October 26, 1977, a hospital cook in Somalia became the last person to catch smallpox naturally — and survived. On May 8, 1980, the WHO declared it gone. It remains the only human disease ever eradicated."
    ],
    diagram: 'flow', steps: [['thought incurable', '3,000 years of fear'], ['ring vaccination', 'contain each case fast'], ['last case, 1977', 'Somalia, survived'], ['eradicated, 1980', 'only disease ever wiped out']],
    source: "World Health Organization, 'The WHO and the Eradication of Smallpox'; CDC 'History of Smallpox'; Wikipedia, cross-checked.",
    links: [
      { label: "Origins (OSU) — The WHO and the Eradication of Smallpox", url: "https://origins.osu.edu/read/who-and-eradication-smallpox" },
      { label: "CDC — History of Smallpox", url: "https://www.cdc.gov/smallpox/about/history.html" },
      { label: "Wikipedia — Smallpox", url: "https://en.wikipedia.org/wiki/Smallpox" }
    ]
  },
  {
    id: 'sars', year: '2003',
    title: "A mystery illness spread to 8 countries in weeks, then stopped",
    teaser: "It started with one night in a hotel room.",
    story: [
      "In November 2002, doctors in Guangdong, China began seeing a strange, severe pneumonia nobody could identify.",
      "In February 2003, a doctor who'd treated these patients spent one night in Room 911 of Hong Kong's Metropole Hotel. He got sick and died — but guests on his floor had already been infected and flew home, carrying the illness with them.",
      "An Italian WHO doctor, Carlo Urbani, was first to recognize it wasn't ordinary pneumonia. He sounded the alarm — and later died of SARS himself.",
      "With no vaccine or cure, the world fought back with the oldest tools in medicine: find every case, isolate and quarantine contacts, and screen travelers. By July 5, 2003, the WHO declared the chain of transmission broken worldwide."
    ],
    diagram: 'flow', steps: [['disease identified', 'named SARS, March 2003'], ['global alert issued', 'WHO warns the world'], ['isolate & quarantine', 'every case, every contact'], ['contained', 'declared July 5, 2003']],
    source: "History.com, 'WHO declares SARS contained worldwide'; Wikipedia, '2002–2004 SARS outbreak'; CDC chronology, cross-checked.",
    links: [
      { label: "History.com — WHO declares SARS contained worldwide", url: "https://www.history.com/this-day-in-history/july-5/world-health-organization-declares-sars-contained-worldwide" },
      { label: "Wikipedia — 2002–2004 SARS outbreak", url: "https://en.wikipedia.org/wiki/2002%E2%80%932004_SARS_outbreak" }
    ]
  },
  {
    id: 'condor', year: '2025',
    title: "One bird was down to 27 left on the entire planet",
    teaser: "Then keepers started feeding chicks with puppets.",
    story: [
      "By 1982 only 22 California condors survived in the wild. Lead poisoning, poison traps, and habitat loss had crushed the population of North America's largest flying bird.",
      "In a bold, controversial move, scientists captured every single wild condor left on Earth. By 1987 all 27 remaining birds were in captivity — the species was extinct in the wild.",
      "Keepers fed chicks using condor-shaped hand puppets so they wouldn't imprint on humans and lose their wild instincts. Starting in 1992, captive-bred birds were released back into California, then Arizona, then Mexico.",
      "It took decades of careful breeding and tracking. As of December 2025, the worldwide population stands at 607 condors — more than half of them flying free."
    ],
    diagram: 'flow', steps: [['capture the last 22', 'wild condors, 1987'], ['breed in zoos', 'puppet-fed chicks'], ['release to the wild', 'starting 1992'], ['607 alive today', 'over half in the wild']],
    source: "U.S. Fish & Wildlife Service, 'California Condor Recovery Program'; Wikipedia, cross-checked against the 2025 FWS Population Status Report.",
    links: [
      { label: "U.S. Fish & Wildlife Service — California Condor Recovery Program", url: "https://www.fws.gov/program/california-condor-recovery" },
      { label: "Wikipedia — California Condor", url: "https://en.wikipedia.org/wiki/California_condor" }
    ]
  },
  {
    id: 'pellagra', year: '1937',
    title: "One man's stubborn idea ended a 200 year old disease",
    teaser: "Doctors were sure it was contagious. He proved them wrong.",
    story: [
      "Starting in the 1730s, a horrifying disease spread across parts of Europe and later the American South — the 'four D's': dermatitis, diarrhea, dementia, death. For 200 years, doctors were sure it was an infection.",
      "In 1914, Dr. Joseph Goldberger noticed that caretakers around sick patients never caught the disease themselves. He suspected it wasn't contagious at all — it was something missing from what poor families ate.",
      "To prove it, he ran 'filth parties' — injecting himself and volunteers, including his own wife, with infected blood and swallowing patients' skin scabs. Nobody got sick. Feeding experiments in orphanages and a Mississippi prison confirmed it: the disease came from a diet missing key nutrients.",
      "It took until 1937 for scientists to pin down the exact culprit — niacin, vitamin B3. Once flour and cornmeal were fortified with it, a disease that had haunted communities for two centuries essentially disappeared."
    ],
    diagram: 'chain',
    source: "NIH Office of History, 'Joseph Goldberger & the War on Pellagra'; UAB Libraries; Wikipedia, cross-checked.",
    links: [
      { label: "NIH Office of History — Joseph Goldberger & the War on Pellagra", url: "https://history.nih.gov/pages/viewpage.action?pageId=8883184" },
      { label: "UAB Libraries — History of Pellagra", url: "https://library.uab.edu/locations/reynolds/collections/regional-history/pellagra/history" },
      { label: "Wikipedia — Joseph Goldberger", url: "https://en.wikipedia.org/wiki/Joseph_Goldberger" }
    ]
  },
  {
    id: 'wolves', year: '1995',
    title: "Wolves changed the course of a river",
    teaser: "Bringing back one predator reshaped an entire park.",
    story: [
      "By the 1920s, wolves had been wiped out of Yellowstone National Park. Without them, elk grazed freely along riverbanks, stripping young trees before they could ever mature, and the riverbanks eroded.",
      "In 1995, scientists reintroduced 41 gray wolves from Canada. Once wolves were hunting again, elk got nervous about lingering in open river valleys where they were easy to catch.",
      "With elk avoiding those areas, willow and aspen trees finally got a chance to grow back. Their roots helped hold the soil together along the riverbanks.",
      "This is often told as a clean, guaranteed chain reaction — but real wolf biologists say it's more complicated. The wolves-affecting-elk-behavior part is solid; how much rivers actually changed because of it is still debated among ecologists."
    ],
    diagram: 'flow', steps: [['wolves return', 'reintroduced in 1995'], ['elk get nervous', 'avoid open riverbanks'], ['plants regrow', 'willow, aspen return'], ['rivers narrow', 'roots hold banks firm']],
    source: "National Park Service wolf restoration records; Ripple & Beschta (2012); International Wolf Center (2025) critical review, cross-checked.",
    links: [
      { label: "International Wolf Center — Do Wolves Really Change Rivers?", url: "https://wolf.org/wolf-info/international-wolf-magazine/do-wolves-really-change-rivers-2/" },
      { label: "Discover Wildlife — Yellowstone wolf reintroduction", url: "https://www.discoverwildlife.com/animal-facts/mammals/yellowstone-wolf-reintroduction" }
    ]
  },
  {
    id: 'ddt', year: '1972',
    title: "One woman's writing stopped a toxic pesticide",
    teaser: "A book about silent birdsong took down a 'miracle' chemical.",
    story: [
      "By the 1950s, America was spraying DDT everywhere to kill insects, seen as a miracle chemical. But it built up in soil and water for years, then worked its way up the food chain into birds.",
      "In 1962, marine biologist Rachel Carson published Silent Spring, describing a future with no birdsong because DDT was thinning bird eggshells so badly they cracked before hatching, pushing eagles and falcons toward extinction.",
      "The chemical industry attacked her viciously, but a review ordered by President Kennedy backed her science completely.",
      "It took a decade of fighting — including a boycott led by farmworkers under Cesar Chavez — before the newly formed EPA banned DDT in 1972. Bald eagles have made a dramatic comeback since."
    ],
    diagram: 'flow', steps: [['DDT sprayed everywhere', 'seen as a miracle chemical'], ['Silent Spring published', 'Rachel Carson, 1962'], ['a decade of pushback', 'scientists, farmworkers fight'], ['DDT banned', 'EPA, 1972']],
    source: "EPA 'DDT — A Brief History and Status'; Wikipedia's DDT entry; Pesticide Action Network's 'The DDT Story', cross-checked.",
    links: [
      { label: "EPA — DDT: A Brief History and Status", url: "https://www.epa.gov/ingredients-used-pesticide-products/ddt-brief-history-and-status" },
      { label: "Wikipedia — DDT", url: "https://en.wikipedia.org/wiki/DDT" },
      { label: "Pesticide Action Network — The DDT Story", url: "https://www.panna.org/resources/ddt-story/" }
    ]
  },
  {
    id: 'y2k', year: '2000',
    title: "The world thought computers would crash at midnight",
    teaser: "Old code could have mistaken 2000 for 1900.",
    story: [
      "Decades ago, computer memory was so expensive that programmers abbreviated years to two digits — 1985 was just stored as '85'. Nobody expected that code to still be running when the year rolled over to '00'.",
      "By the 1990s, that old code was everywhere — running banks, power plants, airlines, and hospitals. If a computer read '00' as 1900 instead of 2000, calculations could break, transactions could fail, and critical systems could behave unpredictably.",
      "A programmer named Peter de Jager sounded the alarm in a 1993 magazine article, and it snowballed into one of the largest coordinated tech efforts in history. Companies and governments spent an estimated $300–500 billion combing through code line by line and fixing dates worldwide.",
      "At midnight on January 1, 2000, city after city crossed into the new millennium — and almost nothing happened. Planes didn't fall, banks kept running, power stayed on. It wasn't a hoax; it was a threat that got fixed before it could strike."
    ],
    diagram: 'flow', steps: [['two-digit years', 'memory-saving shortcut'], ['bug discovered', 'flagged in 1993'], ['global fix effort', '$300–500 billion spent'], ['midnight passes safely', 'Jan 1, 2000']],
    source: "TIME, '20 Years Later, the Y2K bug seems like a joke'; Wikipedia's Year 2000 problem entry; Britannica, 'Y2K bug', cross-checked.",
    links: [
      { label: "TIME — 20 Years Later, the Y2K bug seems like a joke", url: "https://time.com/5752129/y2k-bug-history/" },
      { label: "Wikipedia — Year 2000 problem", url: "https://en.wikipedia.org/wiki/Year_2000_problem" },
      { label: "Britannica — Y2K bug", url: "https://www.britannica.com/technology/Y2K-bug" }
    ]
  },
  {
    id: 'genome', year: '2003',
    title: "Scientists finished mapping the entire human blueprint",
    teaser: "3 billion letters of DNA, read one by one.",
    story: [
      "Every cell in your body carries a genome — around 3 billion chemical 'letters' of DNA that spell out the instructions for building and running a human being. In 1990, an international team set out to read every single one of them.",
      "It was one of the most ambitious science projects ever attempted, sometimes compared to splitting the atom or landing on the Moon. Thousands of researchers across the US, UK, France, Germany, Japan, and China worked together, funded by governments pooling billions of dollars.",
      "Midway through, a private company called Celera Genomics, led by scientist Craig Venter, launched a rival effort using a faster method — turning the project into a race. Eventually the public project and Celera published their results together in 2001.",
      "On April 14, 2003 — fittingly, the 50th anniversary of the discovery of DNA's double-helix shape — the Human Genome Project was declared essentially complete, two years ahead of schedule. It's since made possible everything from personalized medicine to genetic disease screening."
    ],
    diagram: 'flow', steps: [['project launches', '1990, global team'], ['a rival races them', 'Celera Genomics, 1998'], ['rough draft', 'announced 2000'], ['genome completed', 'April 14, 2003']],
    source: "National Human Genome Research Institute (genome.gov); Wikipedia's Human Genome Project entry; Britannica, cross-checked.",
    links: [
      { label: "genome.gov — International Consortium Completes Human Genome Project", url: "https://www.genome.gov/11006929/2003-release-international-consortium-completes-hgp" },
      { label: "Wikipedia — Human Genome Project", url: "https://en.wikipedia.org/wiki/Human_Genome_Project" },
      { label: "Britannica — Human Genome Project", url: "https://www.britannica.com/event/Human-Genome-Project" }
    ]
  },
  {
    id: 'enigma', year: '1940',
    title: "A math puzzle helped win a world war",
    teaser: "An 'unbreakable' code, cracked by pure logic.",
    story: [
      "During World War II, German forces encrypted their military messages using a typewriter-like machine called Enigma, scrambling every letter through a set of rotating wheels. The settings changed daily, creating trillions of possible combinations — the Germans considered it unbreakable.",
      "At a secret British estate called Bletchley Park, a mathematician named Alan Turing was recruited to help crack it. Turing designed an electromechanical machine called the Bombe, which used logic and educated guesses — like knowing German messages often included the word for 'weather' — to rapidly rule out impossible settings and narrow in on the real ones.",
      "Turing built on earlier codebreaking work by Polish mathematicians, who'd cracked an older version of Enigma years before and shared their knowledge with Britain and France once Germany invaded Poland.",
      "By early 1940, Bletchley Park was reading German military communications regularly. Historians estimate this shortened the war by as much as two years, saving countless lives — though the secret was kept for 30 years, and Turing never lived to see his work publicly celebrated."
    ],
    diagram: 'flow', steps: [['Enigma seems unbreakable', 'trillions of settings'], ['Turing designs the Bombe', 'logic + educated guesses'], ['codes cracked', 'early 1940'], ['war shortened', 'by an estimated 2 years']],
    source: "Imperial War Museums, 'How Alan Turing Cracked The Enigma Code'; Bletchley Park; Wikipedia's Bletchley Park entry, cross-checked.",
    links: [
      { label: "Imperial War Museums — How Alan Turing Cracked The Enigma Code", url: "https://www.iwm.org.uk/history/how-alan-turing-cracked-the-enigma-code" },
      { label: "Wikipedia — Bletchley Park", url: "https://en.wikipedia.org/wiki/Bletchley_Park" }
    ]
  },
  {
    id: 'delta-works', year: '1986',
    title: "A country redesigned itself so the sea couldn't drown it again",
    teaser: "1,836 people died. Engineers made sure it wouldn't happen twice.",
    story: [
      "On the night of January 31, 1953, a massive North Sea storm smashed into the Netherlands, a country where over a quarter of the land sits below sea level. Dikes burst in nearly 500 places. By morning, 1,836 people were dead and 200,000 hectares of land lay underwater.",
      "Within weeks, the Dutch government convened a commission of engineers led by Johan van Veen, who had already been warning about this exact risk for years. Their answer was almost unthinkably ambitious: instead of just rebuilding dikes, they'd redesign the entire southwestern coastline.",
      "The result, called the Delta Works, closed off major river estuaries with a system of dams, sluices, and massive movable storm surge barriers — shortening the Dutch coastline that needed defending by about 700 kilometers. The largest barrier, the Oosterscheldekering, spans nearly 9 kilometers and can lower its gates in about an hour when a dangerous storm approaches.",
      "It took over 30 years to complete and is recognized by the American Society of Civil Engineers as one of the Seven Wonders of the Modern World. The design has protected millions of people ever since, though as sea levels keep rising, Dutch engineers are already planning what comes next."
    ],
    diagram: 'flow', steps: [['storm floods the coast', '1,836 dead, 1953'], ['engineers redesign the coastline', 'Delta Commission formed'], ['barriers built', 'shortened coast by 700km'], ['millions protected', 'completed 1986']],
    source: "Institution of Civil Engineers; Britannica's Delta Works entry; Watersnoodmuseum, cross-checked.",
    links: [
      { label: "Institution of Civil Engineers — Delta Works", url: "https://www.ice.org.uk/what-is-civil-engineering/infrastructure-projects/delta-works" },
      { label: "Britannica — Delta Works", url: "https://www.britannica.com/event/Delta-Works" }
    ]
  },
  {
    id: 'ozone', year: '1987',
    title: "A hole in the sky was actually closing",
    teaser: "The world banned one chemical, and it worked.",
    story: [
      "In 1985, British scientists in Antarctica discovered something alarming: a massive hole had opened up in the ozone layer, the thin shield high in the atmosphere that blocks harmful UV radiation from the sun. Without it, skin cancer and eye damage rates would skyrocket worldwide.",
      "Scientists traced the cause to chlorofluorocarbons, or CFCs — everyday chemicals used in spray cans, refrigerators, and air conditioners. A single chlorine atom released from CFCs could destroy over 100,000 ozone molecules before breaking down.",
      "In 1987, nearly every country on Earth signed the Montreal Protocol, agreeing to phase out CFCs entirely. It's the only United Nations treaty ever ratified by all 198 member parties.",
      "It worked. By 2005, ozone-depleting chemical use had dropped 90-95% among signing countries, and by 2019 scientists recorded the smallest ozone hole since 1982. It's still healing — full recovery over Antarctica isn't expected until around 2066 — but it's the clearest proof that global cooperation can fix a planet-sized problem."
    ],
    diagram: 'flow', steps: [['ozone hole found', 'discovered, 1985'], ['CFCs identified', 'the chemical culprit'], ['treaty signed', 'Montreal Protocol, 1987'], ['ozone recovering', 'smallest hole since 1982, 2019']],
    source: "NASA/NOAA ozone recovery reports; UN Environment Programme; Wikipedia's Montreal Protocol entry, cross-checked.",
    links: [
      { label: "NOAA — 4 facts about ozone and the Montreal Protocol", url: "https://www.noaa.gov/stories/4-facts-you-might-not-know-about-ozone-and-montreal-protocol" },
      { label: "UN News — Ozone layer recovery is on track", url: "https://news.un.org/en/story/2023/01/1132277" },
      { label: "Wikipedia — Montreal Protocol", url: "https://en.wikipedia.org/wiki/Montreal_Protocol" }
    ]
  },
  {
    id: 'polio', year: '1955',
    title: "One vaccine ended a disease that paralyzed thousands of kids a year",
    teaser: "He refused to patent it so everyone could have it.",
    story: [
      "In the early 1950s, polio was one of the most feared diseases in America. In 1952 alone, it paralyzed over 18,000 people and killed more than 3,000 — many of them children who ended up needing leg braces, wheelchairs, or breathing machines called iron lungs just to survive.",
      "A researcher named Jonas Salk spent seven years developing a vaccine using a killed version of the virus — safe, but still able to teach the body to fight it. He tested it first on himself, his lab team, and his own children before it ever reached the public.",
      "In 1954, nearly 2 million children volunteered as 'Polio Pioneers' in the largest medical field trial in history. On April 12, 1955, the results were announced to the world: the vaccine was safe and up to 90% effective.",
      "Salk never patented it. When asked who owned the rights, he famously replied, 'Well, the people, I would say. There is no patent. Could you patent the sun?' Within a year, U.S. polio deaths dropped by half, and by 1995 the disease had been eliminated across the entire Western Hemisphere."
    ],
    diagram: 'flow', steps: [['polio terrifies families', '18,000+ paralyzed in 1952'], ['Salk develops vaccine', '7 years of research'], ['2 million kids trial it', 'Polio Pioneers, 1954'], ['vaccine declared safe', 'April 12, 1955']],
    source: "Salk Institute, '70 years of the Salk vaccine'; WHO, 'History of polio vaccination'; Wikipedia's Jonas Salk entry, cross-checked.",
    links: [
      { label: "Salk Institute — The day polio met its match", url: "https://www.salk.edu/news-release/the-day-polio-met-its-match-celebrating-70-years-of-the-salk-vaccine/" },
      { label: "WHO — History of polio vaccination", url: "https://www.who.int/news-room/spotlight/history-of-vaccination/history-of-polio-vaccination" },
      { label: "Wikipedia — Jonas Salk", url: "https://en.wikipedia.org/wiki/Jonas_Salk" }
    ]
  },
  {
    id: 'guinea-worm', year: '2025',
    title: "A worm-borne disease went from 3.5 million cases to just 10",
    teaser: "No vaccine, no medicine — just changed habits.",
    story: [
      "Guinea worm disease is nightmarish: you drink contaminated water, and about a year later a meter-long worm slowly emerges through a painful blister in your skin, usually on your leg. In 1986, an estimated 3.5 million people caught it every year across 21 countries in Africa and Asia.",
      "Former U.S. President Jimmy Carter took on eradicating it as a personal mission through the Carter Center. There's no vaccine and no cure — the only way to stop it is breaking the worm's life cycle by changing behavior.",
      "Workers taught millions of people to filter their drinking water through simple cloth screens, isolated contaminated water sources, and paid villagers to report new cases early so health workers could respond fast.",
      "It's worked almost unbelievably well. By 2025, only 10 human cases were reported worldwide. If eradication is confirmed, Guinea worm will become just the second human disease ever wiped out, after smallpox — and the first ever beaten without any medicine at all."
    ],
    diagram: 'flow', steps: [['3.5 million cases', '21 countries, 1986'], ['Carter Center takes on it', 'no vaccine, no cure'], ['teach water filtering', 'cloth screens, village by village'], ['10 cases left', '2025, nearly eradicated']],
    source: "The Carter Center, 'Guinea Worm Disease Reaches All-Time Low' (2026); Smithsonian Magazine; CBS News, cross-checked.",
    links: [
      { label: "The Carter Center — Guinea Worm Disease press release", url: "https://www.cartercenter.org/news/guinea-worm-announcement/" },
      { label: "Smithsonian Magazine — Jimmy Carter and Guinea worm", url: "https://www.smithsonianmag.com/smart-news/jimmy-carter-worked-to-eradicate-the-vicious-guinea-worm-parasite-slashing-cases-by-the-millions-180985791/" }
    ]
  },
  {
    id: 'titanic-solas', year: '1914',
    title: "A ship's disaster rewrote the rules of the sea",
    teaser: "Not enough lifeboats for everyone on board.",
    story: [
      "When the Titanic sank in April 1912, it had lifeboat space for only about half the people on board — and that actually met the legal requirements of the time, which were based on a ship's weight, not how many passengers it carried.",
      "The disaster killed over 1,500 people, and investigations in both Britain and the U.S. exposed just how outdated maritime safety rules were. Ships also weren't required to keep constant radio watch — a nearby ship, the Californian, never heard Titanic's distress calls because its radio operator had gone to bed.",
      "In 1914, nations came together in London to create SOLAS — the International Convention for the Safety of Life at Sea. It required enough lifeboats for every single person aboard, mandatory lifeboat drills, 24-hour radio watch on passenger ships, and a new International Ice Patrol to track icebergs.",
      "SOLAS has been updated many times since, but it's still in force today and is considered the most important maritime safety treaty in the world. Every time you see a ship carry exactly enough lifeboats for its passengers, that's the Titanic's legacy."
    ],
    diagram: 'flow', steps: [['Titanic sinks', '1912, half enough lifeboats'], ['investigations expose gaps', 'no 24hr radio watch'], ['nations meet in London', '1913-14 conference'], ['SOLAS treaty signed', 'lifeboats for all, 1914']],
    source: "International Maritime Organization; Library of Congress 'The Titanic and the Law'; Wikipedia, cross-checked.",
    links: [
      { label: "Library of Congress — The Titanic and the Law: Safety and Science", url: "https://blogs.loc.gov/law/2024/04/the-titanic-and-the-law-safety-and-science/" },
      { label: "Wikipedia — Changes in safety practices after the sinking of the Titanic", url: "https://en.wikipedia.org/wiki/Changes_in_safety_practices_after_the_sinking_of_the_Titanic" }
    ]
  },
  {
    id: 'wannacry', year: '2017',
    title: "A 22 year old stopped a global cyberattack by accident",
    teaser: "He registered a random web address for a few dollars.",
    story: [
      "On May 12, 2017, a ransomware called WannaCry began tearing through computers worldwide, locking people out of their files and demanding payment to unlock them. Within hours it had hit over 200,000 computers in more than 150 countries, crippling parts of the UK's National Health Service.",
      "Marcus Hutchins, a 22-year-old self-taught security researcher working from his bedroom in England, was analyzing the malware's code when he noticed something odd — it was trying to connect to a strange, unregistered web address before doing anything else.",
      "On a hunch, he bought that domain name for a few dollars, the way researchers sometimes do to track threats. He had no idea what would happen next: registering it accidentally triggered a hidden 'kill switch' built into the malware, and the attack's spread ground to a halt almost immediately.",
      "Hutchins didn't consider himself a hero — he said he was just 'doing his bit to stop botnets.' His quick thinking is credited with slowing the attack enough to prevent it from devastating the United States the way it had Europe and Asia."
    ],
    diagram: 'flow', steps: [['ransomware spreads fast', '200,000+ computers, 150 countries'], ['researcher spots odd code', 'a strange, unregistered URL'], ['buys the domain', 'a hunch, a few dollars'], ['kill switch triggers', 'attack stops spreading']],
    source: "Wikipedia's WannaCry ransomware attack entry; NBC News; Marcus Hutchins Wikipedia entry, cross-checked.",
    links: [
      { label: "Wikipedia — WannaCry ransomware attack", url: "https://en.wikipedia.org/wiki/WannaCry_ransomware_attack" },
      { label: "NBC News — Marcus Hutchins 'Saved the U.S.' from WannaCry", url: "https://www.nbcnews.com/storyline/hacking-of-america/marcus-hutchins-saved-u-s-wannacry-cyberattack-bedroom-compter-n759931" }
    ]
  },
  {
    id: 'insulin', year: '1922',
    title: "A disease that was a death sentence became a daily routine",
    teaser: "Its discoverer refused to profit from it.",
    story: [
      "Before 1921, being diagnosed with type 1 diabetes was essentially a death sentence, especially for children. Doctors could only put patients on starvation diets to buy them a little more time.",
      "In May 1921, a Canadian surgeon named Frederick Banting and a medical student, Charles Best, began experimenting at the University of Toronto, trying to extract a substance from the pancreas that could control blood sugar. It took months of failed attempts and a lot of dead lab dogs before they finally isolated something that worked: insulin.",
      "On January 11, 1922, they gave the first injection to Leonard Thompson, a 14-year-old dying of diabetes. The first dose caused an allergic reaction, but a purified second dose twelve days later worked dramatically — his blood sugar normalized and he lived another 13 years.",
      "The scientists refused to patent insulin for personal profit, selling the rights for a token $1 so it could be produced widely. Within a year it was in mass production, transforming a fatal disease into a manageable one for millions of people."
    ],
    diagram: 'flow', steps: [['diabetes = death sentence', 'before 1921'], ['Banting and Best experiment', 'University of Toronto, 1921'], ['insulin isolated', 'extracted from the pancreas'], ['first patient saved', 'Leonard Thompson, 1922']],
    source: "Nobel Prize official history; Diabetes UK; Penn Today's '100 years of insulin', cross-checked.",
    links: [
      { label: "NobelPrize.org — The miracle discovery that reversed the diabetes death sentence", url: "https://www.nobelprize.org/the-miracle-discovery-that-reversed-the-diabetes-death-sentence/" },
      { label: "Diabetes UK — Who discovered insulin?", url: "https://www.diabetes.org.uk/our-research/about-our-research/our-impact/discovery-of-insulin" }
    ]
  },
  {
    id: 'cholera-pump', year: '1854',
    title: "A doctor traced a killer disease to a single water pump",
    teaser: "He removed a handle, and the story became a legend.",
    story: [
      "In the summer of 1854, cholera tore through the Soho district of London, killing over 600 people in about 10 days. At the time, most doctors believed disease spread through 'bad air' rather than anything in water.",
      "A physician named John Snow disagreed. He'd been suspicious of the 'bad air' theory for years, and when the outbreak hit, he began mapping every single death in the neighborhood, marking each one on a street map. A striking pattern emerged: the deaths clustered tightly around one specific water pump on Broad Street.",
      "Snow convinced skeptical local officials to remove the pump's handle so nobody could use it anymore. The famous story says the epidemic stopped almost overnight — though historians have since found it's more complicated: the outbreak was already declining by the time the handle came off, since many residents had already fled the area.",
      "Even so, Snow's careful mapping and water-based theory turned out to be correct, and decades later scientists confirmed cholera really is spread through contaminated water. Snow is now considered the founder of modern epidemiology — the science of tracking how diseases spread."
    ],
    diagram: 'flow', steps: [['cholera outbreak hits Soho', '600+ deaths, 1854'], ['Snow maps every death', 'pattern points to one pump'], ['pump handle removed', 'Sept 8, 1854'], ['water theory confirmed', 'decades later, by science']],
    source: "Wikipedia's 1854 Broad Street cholera outbreak entry; London Museum; PMC medical history review, cross-checked (note: the 'instant stop' version of this story is a popular simplification — the outbreak was already declining before the handle was removed).",
    links: [
      { label: "London Museum — John Snow: Cholera & the Broad Street pump", url: "https://www.londonmuseum.org.uk/collections/london-stories/john-snow-cholera-broad-street-pump/" },
      { label: "Wikipedia — 1854 Broad Street cholera outbreak", url: "https://en.wikipedia.org/wiki/1854_Broad_Street_cholera_outbreak" }
    ]
  },
  {
    id: 'chile-miners', year: '2010',
    title: "33 miners were trapped half a mile underground for 69 days",
    teaser: "Nobody knew if they were alive for over two weeks.",
    story: [
      "On August 5, 2010, a section of the San José mine in Chile's Atacama Desert collapsed, trapping 33 miners more than 2,300 feet underground — deeper than the height of the Empire State Building.",
      "For 17 agonizing days, nobody on the surface knew if any of them had survived. The miners, trapped in a small refuge chamber, rationed just two spoonfuls of tuna and a sip of milk every two days to stretch their tiny emergency food supply.",
      "When a rescue drill finally broke through, it pulled back up a note taped to the drill bit: 'We are well in the shelter, the 33.' Engineers from Chile, NASA, and around the world raced to design a way to bring them up alive, eventually building a narrow rescue capsule called Fénix.",
      "On October 13, 2010, all 33 miners were pulled to the surface one by one, live on television, watched by an estimated one billion people worldwide. It became one of the most celebrated rescue operations in history."
    ],
    diagram: 'flow', steps: [['mine collapses', '2,300 ft underground, Aug 5'], ['17 days of silence', 'miners ration tiny food supply'], ['note found on drill', '"we are well," day 17'], ['all 33 rescued', 'Oct 13, 2010, live worldwide']],
    source: "Britannica, 'Chile mine rescue of 2010'; NASA oral histories; Wikipedia's 2010 Copiapó mining accident entry, cross-checked.",
    links: [
      { label: "NASA — Chilean Miners Rescue Oral Histories", url: "https://www.nasa.gov/history/history-publications-and-resources/oral-histories/chilean-miners-rescue/" },
      { label: "Wikipedia — 2010 Copiapó mining accident", url: "https://en.wikipedia.org/wiki/2010_Copiap%C3%B3_mining_accident" }
    ]
  },
  {
    id: 'hubble-mirror', year: '1993',
    title: "A telescope that cost billions launched blurry",
    teaser: "Astronauts gave it glasses, 350 miles up.",
    story: [
      "When NASA launched the Hubble Space Telescope in 1990, it was supposed to deliver the sharpest views of the universe ever captured. Instead, its very first images came back blurry — an embarrassing, expensive disaster that made headlines and late-night jokes for years.",
      "Engineers discovered the problem: Hubble's primary mirror had been ground to the wrong shape by a tiny amount — about 1/50th the width of a human hair — during manufacturing. That tiny flaw was enough to scatter the light Hubble collected, ruining its focus.",
      "Since Hubble was designed to be serviced by astronauts in orbit, NASA didn't need to bring it back to Earth to fix it. In December 1993, seven astronauts flew up on the Space Shuttle and performed a record five consecutive spacewalks, installing a corrective optics package nicknamed COSTAR — essentially a set of glasses for the telescope.",
      "It worked spectacularly. The 'fixed' Hubble went on to capture some of the most iconic images in the history of astronomy and is still operating today, over three decades later."
    ],
    diagram: 'flow', steps: [['Hubble launches blurry', 'flawed mirror, 1990'], ['flaw diagnosed', "1/50th a hair's width off"], ['astronauts fly to fix it', '5 spacewalks, Dec 1993'], ['Hubble sees clearly', 'iconic images ever since']],
    source: "NASA Science, 'Hubble's Comeback Story'; Smithsonian National Air and Space Museum; CBS News, cross-checked.",
    links: [
      { label: "NASA Science — Hubble's Comeback Story", url: "https://science.nasa.gov/mission/hubble/impacts-and-benefits/comeback-story/" },
      { label: "Smithsonian — Hubble Space Telescope", url: "https://airandspace.si.edu/explore/stories/hubble-space-telescope" }
    ]
  },
  {
    id: 'great-smog', year: '1956',
    title: "A five day fog killed thousands, then changed the law",
    teaser: "People couldn't see their own feet.",
    story: [
      "In December 1952, a thick, toxic smog settled over London and refused to lift for five straight days. Combining coal smoke with fog, it turned the air a sickly yellow-black — in some places, people genuinely could not see their own feet.",
      "Buses stopped running except for the Underground. Ambulances couldn't navigate the streets, so sick people had to find their own way to hospitals. Even indoor movie theaters had to close because audiences couldn't see the screen through smog that seeped inside buildings.",
      "At first, officials estimated around 4,000 people had died from it. But researchers later dug deeper into the health records and found the true toll was far worse — likely around 12,000 deaths, mostly from lung infections triggered by the polluted air.",
      "The disaster forced Britain's government to act. In 1956, Parliament passed the Clean Air Act, creating smoke-free zones and restricting coal burning in homes and factories. It marked one of the first major environmental laws in modern history — and London's skies have never been that deadly since."
    ],
    diagram: 'flow', steps: [['smog blankets London', '5 days, Dec 1952'], ['thousands fall ill', "can't see their own feet"], ['true death toll found', 'up to 12,000, later research'], ['Clean Air Act passed', '1956, coal restricted']],
    source: "Wikipedia's Great Smog of London entry; London Museum, 'The Great Smog of 1952'; Britannica, cross-checked.",
    links: [
      { label: "London Museum — The Great Smog of 1952", url: "https://www.londonmuseum.org.uk/collections/london-stories/the-great-smog-of-1952/" },
      { label: "Wikipedia — Great Smog of London", url: "https://en.wikipedia.org/wiki/Great_Smog_of_London" }
    ]
  }
];

var SEED_OPEN = [
  {
    id: 'gitanjali-rao', title: "An 11 year old built a device to catch poisoned water",
    story: [
      "When Gitanjali Rao was 10, she watched the news with her parents about the Flint, Michigan water crisis — where lead had contaminated the water supply, poisoning thousands of children.",
      "She couldn't stop thinking about it. She'd read about carbon nanotube sensors used to detect gases in the air, and had an idea: could the same kind of sensor detect lead in water instead? Over more than a year of research and testing, she built a device she called Tethys, named after the Greek goddess of fresh water.",
      "Tethys uses a 3D-printed box about the size of a deck of cards containing carbon nanotubes that react to lead in water, sending results to a smartphone app in seconds — far faster than sending samples to a lab. She won the 2017 Discovery Education 3M Young Scientist Challenge for it, at age 11, and was later granted a real U.S. patent for the invention at 15.",
      "She's still working on it — partnering with scientists at Denver Water to refine the device and test it against other contaminants, with the goal of making fast, affordable water testing available to anyone worried about what's coming out of their tap."
    ],
    source: "USPTO, 'One girl's commitment: Gitanjali Rao's Journey of Innovation'; NPR; Wikipedia's Gitanjali Rao entry, cross-checked.",
    links: [
      { label: "USPTO — One girl's commitment", url: "https://www.uspto.gov/learning-and-resources/journeys-innovation/field-stories/one-girls-commitment" },
      { label: "Wikipedia — Gitanjali Rao (inventor)", url: "https://en.wikipedia.org/wiki/Gitanjali_Rao_(inventor)" }
    ]
  },
  {
    id: 'fionn-ferreira', title: "A 16 year old used magnets to pull plastic out of water",
    story: [
      "Fionn Ferreira grew up kayaking along the coast of West Cork, Ireland. One day he noticed a rock covered in oil residue from a spill — and stuck to that oil were tiny bits of plastic.",
      "He started wondering: plastic and oil are both 'non-polar' in chemistry, meaning they're attracted to each other. Could that be used to pull microplastics — tiny plastic fragments smaller than 5 millimeters — out of water?",
      "At 16, he began experimenting with ferrofluid, a magnetic liquid made from oil and iron oxide powder. When mixed into water, the ferrofluid bonds to microplastics; then a magnet can pull both the ferrofluid and the trapped plastic out, leaving cleaner water behind. Across more than 1,000 tests on ten different types of plastic, his method removed over 85% of microplastics on average.",
      "His project won the 2019 Google Science Fair's global grand prize. He's now a doctoral researcher and still working toward his real goal: scaling the method up so it could be used in wastewater treatment plants before microplastics ever reach the ocean."
    ],
    source: "World Economic Forum, 'This Google Science Fair winner is using a magnetic liquid'; Wikipedia; fionnferreira.com, cross-checked.",
    links: [
      { label: "World Economic Forum — Fionn Ferreira's magnetic liquid", url: "https://www.weforum.org/stories/2020/01/ocean-plastic-pollution-magnetic-liquid-irish-scientist/" },
      { label: "Fionn Ferreira — official site", url: "https://www.fionnferreira.com/" }
    ]
  },
  {
    id: 'boyan-slat', title: "A 16 year old saw more plastic than fish, and decided to fix it",
    story: [
      "While scuba diving in Greece at 16, Boyan Slat noticed something disturbing: there were more plastic bags floating around him than actual fish. The question stuck with him — why doesn't anyone just clean this up?",
      "He turned it into a high school research project, and learned that ocean currents naturally sweep floating plastic into five giant 'garbage patches' worldwide, the biggest being the Great Pacific Garbage Patch. Cleaning it with boats and nets, he calculated, would take thousands of years and cost billions of dollars.",
      "His idea: instead of chasing the plastic, build a passive floating barrier that lets ocean currents bring the plastic to it. At 18, after his TEDx talk about the idea went viral, he dropped out of his aerospace engineering degree with just €300 in savings and founded The Ocean Cleanup.",
      "It hasn't been easy — early prototypes broke apart or failed to hold plastic, and plenty of scientists were skeptical it could ever work at scale. But by the end of 2024, The Ocean Cleanup had pulled over 20 million kilograms of plastic from oceans and rivers, and Slat's team is aiming to remove 90% of floating ocean plastic by 2040."
    ],
    source: "The Ocean Cleanup, 'How it all Began'; TIME Magazine; Wikipedia's Boyan Slat entry, cross-checked.",
    links: [
      { label: "The Ocean Cleanup — How it all Began", url: "https://theoceancleanup.com/milestones/how-it-all-began/" },
      { label: "Wikipedia — Boyan Slat", url: "https://en.wikipedia.org/wiki/Boyan_Slat" }
    ]
  },
  {
    id: 'deepika-kurup', title: "A 14 year old built a sun-powered water cleaner",
    story: [
      "On a trip to visit her grandparents in India, 14-year-old Deepika Kurup watched local children collecting dirty water from the street to drink, cook, and wash with. It stuck with her — kids her own age, drinking water that could make them seriously sick.",
      "Back home in New Hampshire, she spent three months of her summer, skipping camp and vacation, reading graduate-level research papers on water purification instead.",
      "She landed on a method called solar photocatalysis: combining titanium dioxide and zinc oxide, two cheap, non-toxic minerals that, when hit by sunlight, produce compounds that kill bacteria in water. No electricity, no chemicals — just sun. In her tests, it wiped out total coliform bacteria within about 15 minutes of sunlight exposure.",
      "Her invention won America's Top Young Scientist award in 2012. She's since founded a nonprofit, Catalyst for World Water, to get the technology into the hands of communities that need it most — and is now studying medicine at Stanford."
    ],
    source: "Wikipedia's Deepika Kurup entry; Fast Company; TED speaker bio, cross-checked.",
    links: [
      { label: "Wikipedia — Deepika Kurup", url: "https://en.wikipedia.org/wiki/Deepika_Kurup" },
      { label: "Fast Company — Meet the 14-year-old girl", url: "https://www.fastcompany.com/2681073/meet-the-14-year-old-girl-who-developed-a-low-cost-water-purification-system" }
    ]
  },
  {
    id: 'vinisha-umashankar', title: "A 12 year old redesigned a cart to stop burning charcoal",
    story: [
      "Walking home from school in Tamil Nadu, India, Vinisha Umashankar noticed something she'd seen a hundred times before without really seeing it: a street vendor dumping burnt charcoal from his ironing cart. India has an estimated 10 million of these carts, each burning about 11 pounds of charcoal a day to heat clothes irons.",
      "Curious, she started researching — and learned that all that charcoal burning was contributing to deforestation and filling vendors' lungs with smoke every single day.",
      "At just 12 years old, she spent six months teaching herself college-level physics from textbooks to design a solution: a cart with solar panels on its roof that could power the iron directly, storing extra energy in a battery for cloudy days or nighttime. She partnered with India's National Innovation Foundation to build a working, patented prototype called Iron-Max.",
      "Her invention became a finalist for the £1 million Earthshot Prize, and she's spoken at the United Nations climate summit COP26, urging world leaders to act. She's now working to get her solar carts manufactured and into vendors' hands across India."
    ],
    source: "TIME, 'Earthshot Finalist, Age 14, Invented Solar Ironing Cart'; NPR; Earthshot Prize official bio, cross-checked.",
    links: [
      { label: "TIME — Vinisha Umashankar's solar ironing cart", url: "https://time.com/6101003/earthshot-prize-india-solar-ironing/" },
      { label: "NPR — A 15-year-old girl invented a solar ironing cart", url: "https://www.npr.org/sections/goatsandsoda/2021/11/03/1050227033/a-15-year-old-girl-invented-a-solar-ironing-cart-thats-winning-global-respect" }
    ]
  },
  {
    id: 'mari-copeny', title: "An 8 year old's letter got a president to visit her city",
    story: [
      "In 2014, the city of Flint, Michigan switched its water source to save money — and didn't properly treat the new water, letting lead leach out of old pipes into the drinking water of roughly 100,000 people, including thousands of children.",
      "Mari Copeny, known locally as 'Little Miss Flint,' was 8 years old and living through it. In 2016, she wrote a letter to President Obama, telling him she was one of the kids affected and that she'd been marching in protest to speak up for the other kids in her city.",
      "Obama wrote back — and then flew to Flint to meet her in person, later approving $100 million in federal relief funds for the city's water infrastructure.",
      "Mari didn't stop there. She's since raised hundreds of thousands of dollars to distribute water filters and bottled water to Flint families, donated over a million water bottles, and become a national voice speaking out about water crises in other communities facing the same danger. As she put it: 'If they don't give you a seat at the table, stand on it with a megaphone.'"
    ],
    source: "Wikipedia's Amariyanna Copeny entry; NUHW; YR Media interview, cross-checked.",
    links: [
      { label: "Wikipedia — Amariyanna Copeny", url: "https://en.wikipedia.org/wiki/Amariyanna_Copeny" },
      { label: "YR Media — Little Miss Flint's Call to Action", url: "https://yr.media/health/little-miss-flint-fighting-environmental-racism/" }
    ]
  },
  {
    id: 'pythagorean-teens', title: "Two teenagers solved a math problem called 'impossible' for 2,000 years",
    story: [
      "In December 2022, a Louisiana high school offered students a bonus math challenge with a $500 prize: use trigonometry to prove the Pythagorean theorem — the ancient rule about right triangles that most people learn in geometry class.",
      "There was a catch nobody told them about: mathematicians had long considered this specific task circular and impossible. Trigonometry itself is built on the Pythagorean theorem, so using it to prove the theorem seemed like using a fact to prove itself.",
      "Seniors Calcea Johnson and Ne'Kiya Jackson didn't know that reputation — they just worked the problem, sometimes on weekends and holidays, and came up with a genuinely new proof using a geometric construction they nicknamed the 'waffle cone.' They presented it to the American Mathematical Society in 2023.",
      "They didn't stop at one. In 2024, now in college, they published a peer-reviewed paper in the American Mathematical Monthly with 10 total trigonometric proofs of the theorem — only the third and fourth such proofs ever found before theirs, out of a problem that had stood since ancient Greece."
    ],
    source: "Science News; Smithsonian Magazine; CBS News 60 Minutes, cross-checked.",
    links: [
      { label: "Science News — Two teenagers have once again proved an ancient math rule", url: "https://www.sciencenews.org/article/teenagers-pythagorean-theorem-math" },
      { label: "Smithsonian Magazine — Two High Schoolers Found an 'Impossible' Proof", url: "https://www.smithsonianmag.com/smart-news/two-high-schoolers-found-an-impossible-proof-for-a-2000-year-old-math-rule-then-they-discovered-nine-more-180985357/" }
    ]
  },
  {
    id: 'easton-lachappelle', title: "A 14 year old built a robotic hand out of Legos",
    story: [
      "Growing up in rural Colorado with no robotics classes nearby, Easton LaChappelle taught himself electronics and coding from YouTube videos. At 14, he built his first robotic hand using Legos, fishing line, and electrical tubing, controlled by a repurposed Nintendo Power Glove.",
      "The project that changed everything happened at a science fair, where he met a 7-year-old girl wearing a prosthetic arm that cost $80,000 — and that she'd need to replace over and over as she grew, at the same staggering price each time.",
      "Easton realized his hobby could actually help someone. He switched from Legos to 3D printing, and by 17 had built a fully functional robotic prosthetic arm — controllable by an EEG headset that reads brainwaves — for under $500, a tiny fraction of the industry price.",
      "He skipped college to found Unlimited Tomorrow, a company building affordable, custom 3D-printed prosthetic arms for around $8,000 instead of tens of thousands. His technology has since been shown to a sitting U.S. president and used by NASA engineers."
    ],
    source: "CNN; Society for Science; Fast Company, cross-checked.",
    links: [
      { label: "CNN — Young entrepreneur creates robotic limbs controlled by the mind", url: "https://www.cnn.com/2020/08/31/health/prosthetic-unlimited-tomorrow-robotics-scn-trnd/" },
      { label: "Society for Science — This alum prints prosthetics of the future", url: "https://www.societyforscience.org/blog/this-alum-prints-prosthetics-of-the-future/" }
    ]
  }
];

var SEED_PROBLEMS = [
  {
    id: 'plastic', title: "A trash patch bigger than Texas floats in the Pacific",
    story: [
      "Between California and Hawaii sits the Great Pacific Garbage Patch — a vast zone where ocean currents trap floating plastic. Researchers estimate it covers about 1.6 million square kilometers, roughly three times the size of France, and holds somewhere around 1.8 trillion pieces of plastic.",
      "It isn't a solid floating island like people picture. Most of it is 'confetti' — tiny microplastics smaller than a fingernail, spread through the water like pepper in soup, mixed with larger debris like fishing nets, which make up close to half the total mass.",
      "Cleanup is brutally hard: NOAA estimates it would take 67 ships a full year to clear less than 1% of just the North Pacific. Nets big enough to catch the plastic also scoop up the marine life living in and around it. No country has taken responsibility for cleaning it up, since it sits in international waters."
    ],
    source: "The Ocean Cleanup, 'The Great Pacific Garbage Patch'; NOAA Marine Debris Program; Wikipedia's Great Pacific Garbage Patch entry, cross-checked.",
    links: [
      { label: "The Ocean Cleanup — The Great Pacific Garbage Patch", url: "https://theoceancleanup.com/great-pacific-garbage-patch/" },
      { label: "NOAA — How Big Is the Great Pacific Garbage Patch?", url: "https://response.restoration.noaa.gov/about/media/how-big-great-pacific-garbage-patch-science-vs-myth.html" },
      { label: "Wikipedia — Great Pacific Garbage Patch", url: "https://en.wikipedia.org/wiki/Great_Pacific_Garbage_Patch" }
    ]
  },
  {
    id: 'amr', title: "Doctors are running out of medicines that still work",
    story: [
      "Since penicillin's discovery in 1928, antibiotics have saved millions of lives. But every time we use one, the bacteria that happen to survive multiply — slowly training germs to beat our best drugs. Scientists call the survivors 'superbugs.'",
      "Meanwhile, the pipeline of brand-new antibiotics has slowed to a trickle, partly because developing them isn't very profitable for drug companies.",
      "Drug-resistant infections already kill at least 1.27 million people worldwide every year and are linked to nearly 5 million deaths total, according to a major Lancet study. Left unchecked, that toll could climb toward 10 million deaths a year by 2050 — comparable to cancer."
    ],
    source: "World Health Organization fact sheet on antimicrobial resistance; CDC 'Facts and Stats'; The Lancet study reported via CNN (2024).",
    links: [
      { label: "WHO — Antimicrobial Resistance fact sheet", url: "https://www.who.int/news-room/fact-sheets/detail/antimicrobial-resistance" },
      { label: "CDC — Antimicrobial Resistance Facts and Stats", url: "https://www.cdc.gov/antimicrobial-resistance/data-research/facts-stats/index.html" },
      { label: "CNN — 40 million could die from superbug infections by 2050", url: "https://www.cnn.com/2024/09/16/health/antibiotic-resistant-superbug-infections-2050-wellness/index.html" }
    ]
  },
  {
    id: 'debris', title: "A swarm of junk circles Earth at 17,000 mph",
    story: [
      "Since Sputnik launched in 1957, humanity has left behind dead satellites, spent rocket stages, and fragments from collisions — all still orbiting. Space agencies currently track roughly 30,000-plus objects larger than 10 centimeters, and estimate over a million more between 1-10 cm that are too small to track individually.",
      "At orbital speeds — around 7 to 8 kilometers per second — even a 1 cm paint fleck hits with the force of a hand grenade. A single 2009 collision between a dead Russian satellite and a working Iridium satellite alone created over 2,300 trackable fragments.",
      "Scientists worry about a scenario called Kessler Syndrome: once debris gets dense enough, collisions create more debris, which causes more collisions, in a runaway cascade that could make some orbits unusable for generations. No large-scale cleanup mission has succeeded yet."
    ],
    source: "NASA Orbital Debris Program Office FAQ; ESA Space Debris Office statistics; Scientific American (2026) reporting on Space-Track.org data.",
    links: [
      { label: "NASA Orbital Debris Program Office — FAQ", url: "https://orbitaldebris.jsc.nasa.gov/faq/" },
      { label: "ESA Space Debris User Portal — statistics", url: "https://sdup.esoc.esa.int/discosweb/statistics/" },
      { label: "Scientific American — Almost half of objects in orbit are junk", url: "https://www.scientificamerican.com/article/almost-half-of-the-objects-in-earths-orbit-is-junk-and-thats-only-the-stuff-we-know-about/" }
    ]
  },
  {
    id: 'alzheimers', title: "Someone's memories are slipping away and no cure exists",
    story: [
      "Alzheimer's disease slowly destroys brain cells involved in memory and thinking. It's the most common cause of dementia, and today an estimated 7.4 million Americans over 65 live with it — a number expected to nearly double by 2050 without a breakthrough.",
      "For decades, there was no way to treat the disease itself, only its symptoms. That changed recently: two new drugs, Leqembi and Kisunla, were approved to actually clear the harmful amyloid protein plaques building up in the brain, and can modestly slow decline in early-stage patients.",
      "But neither drug is a cure — they don't reverse damage already done, and a major 2026 review called their effect on cognition 'trivial.' Researchers are now exploring entirely different angles, including targeting a second protein called tau and testing treatment years before symptoms even start."
    ],
    source: "Alzheimer's Association, '2026 Alzheimer's Disease Facts and Figures'; Mayo Clinic; World Economic Forum (2026), cross-checked.",
    links: [
      { label: "Alzheimer's Association — Treatments for Alzheimer's & Dementia", url: "https://www.alz.org/alzheimers-dementia/treatments" },
      { label: "Mayo Clinic — Alzheimer's treatments: What's on the horizon?", url: "https://www.mayoclinic.org/diseases-conditions/alzheimers-disease/in-depth/alzheimers-treatments/art-20047780" },
      { label: "World Economic Forum — Alzheimer's disease: 9 recent breakthroughs", url: "https://www.weforum.org/stories/2026/04/alzheimers-disease-breakthroughs-health/" }
    ]
  },
  {
    id: 'malaria', title: "A mosquito bite still kills someone every minute",
    story: [
      "Malaria is caused by a parasite spread through the bite of infected mosquitoes. In 2024 alone, it caused an estimated 282 million cases and about 610,000 deaths worldwide — roughly 95% of them in Africa, and most of them children under 5.",
      "Since 2000, tools like insecticide-treated bed nets, better treatments, and two new vaccines rolled out since 2021 have helped avert an estimated 14 million deaths. New tools alone prevented around 1 million deaths in 2024.",
      "But progress has stalled: cases actually rose by about 9 million between 2023 and 2024, and drug-resistant strains of the malaria parasite are now confirmed or suspected in at least 8 African countries, threatening the medicines that have worked for decades."
    ],
    source: "WHO fact sheet on malaria; World malaria report 2025; CDC, cross-checked.",
    links: [
      { label: "WHO — Fact sheet about malaria", url: "https://www.who.int/news-room/fact-sheets/detail/malaria" },
      { label: "WHO — World malaria report 2025", url: "https://www.who.int/teams/global-malaria-programme/reports/world-malaria-report-2025" }
    ]
  },
  {
    id: 'coral-bleaching', title: "84% of the world's coral reefs have been hit by heat stress",
    story: [
      "Coral reefs get their color from tiny algae living inside them. When ocean water gets too hot, corals expel that algae in a stress response called bleaching — turning stark white and, if the heat doesn't ease up, dying.",
      "Between January 2023 and early 2025, the fourth global mass bleaching event on record hit at least 84% of the world's reefs across 82 countries — the most widespread bleaching ever documented, and far worse than the previous record of 68% set in 2014-2017.",
      "Reefs cover less than 1% of the ocean floor but support around a quarter of all marine species and provide food or income for roughly a billion people. Scientists estimate the world has already lost about half its coral cover since the 1950s."
    ],
    source: "International Coral Reef Initiative; UNEP; NOAA Coral Reef Watch, cross-checked.",
    links: [
      { label: "ICRI — 84% of the world's coral reefs impacted", url: "https://icriforum.org/4gbe-2025/" },
      { label: "UNEP — The world's corals are bleaching", url: "https://www.unep.org/news-and-stories/story/worlds-corals-are-bleaching-heres-why-and-what-it-means-oceans-future" }
    ]
  },
  {
    id: 'pfas', title: "'Forever chemicals' are in almost everyone's bloodstream",
    story: [
      "PFAS are a family of over 12,000 human-made chemicals used since the 1940s in things like non-stick pans, waterproof clothing, firefighting foam, and fast food wrappers, prized for resisting heat, water, and stains.",
      "The problem is right there in their nickname: 'forever chemicals.' Their carbon-fluorine bond is so strong that they barely break down in the environment or the human body — some can persist for thousands of years.",
      "Studies have found detectable PFAS in the blood of 97-99% of Americans tested, and nearly half of U.S. tap water is estimated to contain at least one type. They've been linked to cancer, reduced fertility, and developmental problems in children. There's no known safe level of exposure, and cleaning them out of water supplies is still technically difficult and expensive."
    ],
    source: "EPA; National Institute of Environmental Health Sciences; Johns Hopkins Bloomberg School of Public Health, cross-checked.",
    links: [
      { label: "EPA — Human health and environmental risks of PFAS", url: "https://www.epa.gov/pfas/our-current-understanding-human-health-and-environmental-risks-pfas" },
      { label: "Johns Hopkins — What to Know About PFAS", url: "https://publichealth.jhu.edu/2024/what-to-know-about-pfas" }
    ]
  },
  {
    id: 'bee-collapse', title: "Honeybee colonies keep mysteriously disappearing",
    story: [
      "Since 2006, beekeepers around the world have reported a strange phenomenon: worker bees simply vanish from otherwise healthy hives, leaving the queen and immature bees behind with no one to feed them. It's called Colony Collapse Disorder.",
      "No single cause has ever been proven. Researchers point to a tangled mix of pesticides (especially a class called neonicotinoids), parasitic mites, viruses, and poor nutrition, all stacking up to weaken bee immune systems at once.",
      "The stakes are enormous: honeybees are responsible for pollinating roughly a third of the food humans eat and contribute an estimated $15 billion to the U.S. economy alone. U.S. beekeepers have lost 30-45% of their colonies in a typical winter since CCD began, compared to 15-20% before it — a heavy, ongoing toll even though the most extreme early collapse rates have eased somewhat."
    ],
    source: "USDA Agricultural Research Service; US EPA; National Invasive Species Information Center, cross-checked.",
    links: [
      { label: "USDA ARS — Honey Bee Health", url: "https://www.ars.usda.gov/oc/br/ccd/index/" },
      { label: "EPA — Colony Collapse Disorder", url: "https://www.epa.gov/pollinator-protection/colony-collapse-disorder" }
    ]
  },
  {
    id: 'water-scarcity', title: "Nearly 4 billion people don't have enough water some months",
    story: [
      "Only about 0.5% of all the water on Earth is fresh, liquid, and actually reachable by people — the rest is locked in oceans, ice caps, or deep underground. That tiny slice has to support a growing global population.",
      "Right now, an estimated 2.2 billion people lack access to safely managed drinking water, and roughly 4 billion people experience severe water scarcity for at least one month every year. Agriculture alone uses about 70% of all freshwater withdrawn worldwide.",
      "A major 2026 UN scientific assessment concluded the world has pushed past a 'safe operating space' for freshwater entirely, warning of a new era of chronic water insecurity. More than half of the world's large lakes have shrunk since the early 1990s, and droughts affected over 1.4 billion people between 2002 and 2021 alone."
    ],
    source: "UN World Water Development Report; WHO/UNICEF; United Nations University, cross-checked.",
    links: [
      { label: "United Nations University — World Enters 'Era of Global Water Bankruptcy'", url: "https://unu.edu/inweh/news/world-enters-era-of-global-water-bankruptcy" },
      { label: "WHO — 1 in 4 people globally still lack access to safe drinking water", url: "https://www.who.int/news/item/26-08-2025-1-in-4-people-globally-still-lack-access-to-safely-managed-drinking-water" }
    ]
  },
  {
    id: 'wildfires', title: "Wildfires now burn twice as much forest as 20 years ago",
    story: [
      "Fire has always been part of many forest ecosystems. But data going back to 2001 shows forest fires now burn more than twice as much tree cover globally each year as they did two decades ago — and four of the five worst years for global forest fires have happened since 2020.",
      "Scientists point to climate change as the main driver: hotter temperatures, earlier snowmelt, and drier vegetation create longer, more intense fire seasons. A single degree Celsius of warming could increase the area burned per year by up to 600% in some Western U.S. forest types.",
      "The consequences ripple outward — Canada burned nearly 7.8 million hectares of forest in 2023, about six times its 20-year average, while wildfire smoke now travels thousands of miles, harming air quality and health far from the flames themselves."
    ],
    source: "NASA Science; World Resources Institute; NOAA, cross-checked.",
    links: [
      { label: "World Resources Institute — New Data Confirms: Forest Fires Are Getting Worse", url: "https://www.wri.org/insights/global-trends-forest-fires" },
      { label: "NASA Science — Wildfires and Climate Change", url: "https://science.nasa.gov/earth/explore/wildfires-and-climate-change/" }
    ]
  },
  {
    id: 'road-safety', title: "A road crash kills someone every 26 seconds worldwide",
    story: [
      "Road traffic crashes kill about 1.19 million people every year globally — and they're the single leading cause of death for children and young people aged 5 to 29.",
      "The burden falls extremely unevenly: 92% of road deaths happen in low- and middle-income countries, even though those countries have only about 60% of the world's vehicles. Pedestrians, cyclists, and motorcyclists — the most vulnerable road users — make up over half of all deaths.",
      "The United Nations set a goal to cut road deaths in half by 2030, but progress has been slow: deaths fell only about 5% between 2010 and 2021, even as safer-vehicle technology and seatbelt laws spread. Fewer than 50 countries have laws addressing all five biggest risk factors — speed, drunk driving, helmets, seatbelts, and child restraints."
    ],
    source: "WHO Global status report on road safety 2023; CDC Global Road Safety; PAHO/WHO, cross-checked.",
    links: [
      { label: "WHO — Road traffic injuries fact sheet", url: "https://www.who.int/news-room/fact-sheets/detail/road-traffic-injuries" },
      { label: "PAHO/WHO — Road safety remains urgent global issue", url: "https://www.paho.org/en/news/13-12-2023-despite-notable-progress-road-safety-remains-urgent-global-issue" }
    ]
  },
  {
    id: 'fusion-energy', title: "Scientists can already fuse atoms — just not cheaply enough yet",
    story: [
      "Nuclear fusion is the same reaction that powers the sun: smashing light atoms together to release enormous energy, with no long-lived radioactive waste and no risk of a meltdown. If it worked at scale, it could mean nearly limitless clean electricity.",
      "In December 2022, the National Ignition Facility became the first lab to achieve 'ignition' — producing more fusion energy than the laser energy used to trigger it. Since then, private companies have raised nearly $10 billion chasing the next milestone: sustained net electricity to an actual power grid.",
      "Nobody has done that yet. The huge international ITER project in France won't attempt full fusion until 2039. Private companies like Commonwealth Fusion and Helion are racing to get there sooner, but even optimistic timelines put real, grid-connected fusion power in the early-to-mid 2030s at best — a technology that has famously felt '30 years away' for decades."
    ],
    source: "CNN; Earth911's 'The State of Fusion Energy in 2026'; EveryCRSReport congressional summary, cross-checked.",
    links: [
      { label: "CNN — This company says nuclear fusion could finally power the grid", url: "https://www.cnn.com/2026/04/30/climate/nuclear-fusion-real-world-electricity-grid" },
      { label: "Earth911 — The State of Fusion Energy in 2026", url: "https://earth911.com/eco-tech/the-state-of-fusion-energy-in-2026-real-reactors-real-grids-real-caveats/" }
    ]
  }
];

var SEED_QUOTES = [
  { text: "Could you patent the sun?", attr: "Jonas Salk, on why he never patented the polio vaccine", url: "https://www.salk.edu/news-release/the-day-polio-met-its-match-celebrating-70-years-of-the-salk-vaccine/" },
  { text: "I'm just someone doing my bit to stop botnets.", attr: "Marcus Hutchins, after stopping the WannaCry cyberattack", url: "https://www.nbcnews.com/storyline/hacking-of-america/marcus-hutchins-saved-u-s-wannacry-cyberattack-bedroom-compter-n759931" },
  { text: "Why can't we just clean this up?", attr: "Boyan Slat, age 16, founder of The Ocean Cleanup", url: "https://theoceancleanup.com/milestones/how-it-all-began/" },
  { text: "The only way to do great work is to love what you do.", attr: "Steve Jobs, Stanford commencement address, 2005", url: "https://news.stanford.edu/2005/06/14/jobs-061505/" },
  { text: "If I can do it, you can do it. Anyone can do it.", attr: "Gitanjali Rao, TIME's first Kid of the Year", url: "https://www.uspto.gov/learning-and-resources/journeys-innovation/field-stories/one-girls-commitment" },
  { text: "Somewhere, something incredible is waiting to be known.", attr: "Carl Sagan", url: "https://en.wikipedia.org/wiki/Carl_Sagan" },
  { text: "There is no stop button. There is no magic fix.", attr: "Vinisha Umashankar, teen inventor of a solar ironing cart", url: "https://time.com/6101003/earthshot-prize-india-solar-ironing/" },
  { text: "If they don't give you a seat at the table, stand on it with a megaphone.", attr: "Mari Copeny, 'Little Miss Flint'", url: "https://yr.media/health/little-miss-flint-fighting-environmental-racism/" },
  { text: "We are all equal now.", attr: "Luis Urzúa, shift supervisor of the trapped Chilean miners, 2010", url: "https://en.wikipedia.org/wiki/2010_Copiap%C3%B3_mining_accident" },
  { text: "That's one small step for man, one giant leap for mankind.", attr: "Neil Armstrong, first Moon landing, 1969", url: "https://en.wikipedia.org/wiki/Apollo_11" },
  { text: "I have not failed. I've just found 10,000 ways that won't work.", attr: "Thomas Edison", url: "https://en.wikipedia.org/wiki/Thomas_Edison" },
  { text: "Nothing in life is to be feared, it is only to be understood.", attr: "Marie Curie", url: "https://en.wikipedia.org/wiki/Marie_Curie" },
  { text: "The important thing is not to stop questioning.", attr: "Albert Einstein", url: "https://en.wikipedia.org/wiki/Albert_Einstein" },
  { text: "It always seems impossible until it's done.", attr: "Nelson Mandela", url: "https://en.wikipedia.org/wiki/Nelson_Mandela" },
  { text: "An innovation's true potential is understood only when it reaches people.", attr: "Vinisha Umashankar", url: "https://www.npr.org/sections/goatsandsoda/2021/11/03/1050227033/a-15-year-old-girl-invented-a-solar-ironing-cart-thats-winning-global-respect" }
];

var SEED_FACTS = [
  { text: "there's a piece of hardware still sitting untouched on the Moon — Apollo 11's landing stage, left there in 1969. No wind, no rain, so it hasn't moved an inch in over 55 years.", url: "https://en.wikipedia.org/wiki/Apollo_11" },
  { text: "a flight recorder isn't actually black — it's painted bright orange so rescue teams can spot it in wreckage.", url: "https://en.wikipedia.org/wiki/Flight_recorder" },
  { text: "penicillin, the first antibiotic, was discovered by accident when a scientist noticed mold killing bacteria in a forgotten petri dish.", url: "https://en.wikipedia.org/wiki/Penicillin" },
  { text: "the International Space Station orbits Earth roughly every 90 minutes — astronauts on board see 16 sunrises and sunsets a day.", url: "https://en.wikipedia.org/wiki/International_Space_Station" },
  { text: "the largest living bird in North America, the California condor, can glide for hours without flapping its wings once.", url: "https://en.wikipedia.org/wiki/California_condor" },
  { text: "octopuses have three hearts, and two of them stop beating whenever the octopus swims.", url: "https://en.wikipedia.org/wiki/Octopus" },
  { text: "a bolt of lightning is roughly five times hotter than the surface of the sun.", url: "https://en.wikipedia.org/wiki/Lightning" },
  { text: "honey found in ancient Egyptian tombs, thousands of years old, is technically still edible.", url: "https://en.wikipedia.org/wiki/Honey" },
  { text: "bananas are technically classified as berries, botanically speaking.", url: "https://en.wikipedia.org/wiki/Banana" },
  { text: "sharks have been around longer than trees — by tens of millions of years.", url: "https://en.wikipedia.org/wiki/Shark" },
  { text: "wombats are the only animal known to produce cube-shaped droppings.", url: "https://en.wikipedia.org/wiki/Wombat" },
  { text: "octopus blood is blue, because it's built around copper instead of iron.", url: "https://en.wikipedia.org/wiki/Hemocyanin" },
  { text: "the first computer 'bug' was an actual moth found stuck in a machine in 1947.", url: "https://en.wikipedia.org/wiki/Software_bug" },
  { text: "astronauts can temporarily grow a couple of centimeters taller while in microgravity.", url: "https://en.wikipedia.org/wiki/Effect_of_spaceflight_on_the_human_body" },
  { text: "sea otters hold hands while sleeping so they don't drift apart from each other.", url: "https://en.wikipedia.org/wiki/Sea_otter" }
];
