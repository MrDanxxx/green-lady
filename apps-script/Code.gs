// Green Lady cider - Google Sheet backend (Google Apps Script).
//
// Paste this whole file into the Sheet's Extensions > Apps Script editor.
// 1. Change CREW_PIN below to your own number.
// 2. Run "setup" once (creates the tabs with headings and one example row each).
// 3. Deploy > New deployment > Web app: Execute as "Me", Who has access "Anyone".
//
// Crew actions (Fill, Reading, Moved, Bottled, Note, Empty) need the PIN.
// Bottle ratings from friends need no PIN.
// Anything in the Barrels, Blends and Log tabs can be seen by anyone who has
// the page link. Ratings are shown only as counts (names and comments stay
// private in the Sheet).

const CREW_PIN = "1234";   // <-- CHANGE THIS

const TABS = {
  Barrels: ["Barrel", "Size", "Type", "Location", "Notes"],
  Blends:  ["Blend", "Vintage", "Juiced", "Apples", "Gravity", "pH", "Yeast", "Notes"],
  Log:     ["Time", "Barrel", "Action", "Blend", "Gravity", "pH", "Bottles", "Priming", "Location", "Note", "By"],
  Ratings: ["Time", "Bottle", "Rating", "Tags", "Name", "Comment"],
};
// Sheet tab names (left = what the page calls it, right = the tab name in YOUR Sheet).
const TAB_NAMES = {
  Barrels: "Barrels",
  Blends:  "Blends",
  Log:     "Fermentation and bottling",
  Ratings: "Ratings",
};
const ACTIONS = ["Fill", "Reading", "Moved", "Bottled", "Note", "Empty"];
const RATINGS = ["Good", "OK", "Poor"];

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(TABS).forEach(function (name) {
    tab(ss, name);
  });
  const blank = ss.getSheetByName("Sheet1");
  if (blank && blank.getLastRow() === 0) ss.deleteSheet(blank);

  const barrels = tab(ss, "Barrels");
  if (barrels.getLastRow() === 1) barrels.appendRow(["B01", "25 L", "Plastic", "Home", "Example row - change or delete"]);
  const blends = tab(ss, "Blends");
  if (blends.getLastRow() === 1) blends.appendRow(["26A", 26, "2026-10-04",
    "Dabinett 40% - Home orchard; Bramley 60% - Neighbour's tree", 1.052, 3.5, "Mangrove Jack's M02",
    "Example row - change or delete"]);
  blends.getRange("E:E").setNumberFormat("0.000");
  tab(ss, "Log").getRange("E:E").setNumberFormat("0.000");
}

function doGet() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const out = {};
  Object.keys(TABS).forEach(function (name) {
    const sh = ss.getSheetByName(TAB_NAMES[name] || name);
    if (!sh) { out[name] = []; return; }
    const rows = sh.getDataRange().getValues();
    const head = rows.shift();
    out[name] = rows
      .filter(function (r) { return r.some(function (x) { return x !== ""; }); })
      .map(function (r) {
        const o = {};
        head.forEach(function (h, i) {
          let x = r[i];
          if (x instanceof Date) x = Utilities.formatDate(x, tz, "yyyy-MM-dd HH:mm");
          if (x !== "" && h) o[h] = x;
        });
        return o;
      });
  });
  // Friends' names and comments stay private.
  out.Ratings = out.Ratings.map(function (r) { return { Bottle: r.Bottle, Rating: r.Rating, Tags: r.Tags || "" }; });
  return reply(out);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActive();
    const now = new Date();

    if (d.kind === "pin") return reply({ ok: String(d.pin) === String(CREW_PIN) });

    if (d.kind === "rating") {
      if (RATINGS.indexOf(d.rating) < 0) return reply({ ok: false, error: "Unknown rating" });
      tab(ss, "Ratings").appendRow([now, clean(d.bottle, 30).toUpperCase(), d.rating,
        clean(d.tags, 200), clean(d.name, 40), clean(d.comment, 500)]);
      return reply({ ok: true });
    }

    if (d.kind === "log") {
      if (String(d.pin) !== String(CREW_PIN)) return reply({ ok: false, error: "Wrong crew PIN" });
      if (ACTIONS.indexOf(d.action) < 0) return reply({ ok: false, error: "Unknown action" });
      tab(ss, "Log").appendRow([now, clean(d.barrel, 10).toUpperCase(), d.action,
        clean(d.blend, 20).toUpperCase(), num(d.gravity), num(d.ph), num(d.bottles), clean(d.priming, 40),
        clean(d.location, 60), clean(d.note, 500), clean(d.by, 40)]);
      return reply({ ok: true });
    }

    return reply({ ok: false, error: "Unknown request" });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Finds a tab by its name in TAB_NAMES; creates it with headings if it's missing.
function tab(ss, name) {
  const real = TAB_NAMES[name] || name;
  let sh = ss.getSheetByName(real);
  if (!sh) sh = ss.insertSheet(real);
  if (sh.getLastRow() === 0) {
    sh.appendRow(TABS[name]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, TABS[name].length).setFontWeight("bold").setBackground("#e6f0d6");
  }
  return sh;
}

function clean(s, max) {
  s = String(s == null ? "" : s).slice(0, max).trim();
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // never let typed text become a formula
  return s;
}
function num(x) {
  if (x === "" || x == null) return "";
  const n = Number(x);
  return isFinite(n) ? n : "";
}
function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
