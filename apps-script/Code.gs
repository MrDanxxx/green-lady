// Green Lady cider - Google Sheet backend (Google Apps Script).
//
// Paste this whole file into the Sheet's Extensions > Apps Script editor.
// 1. Change CREW_PIN below to your own number.
// 2. Run "setup" once (creates the tabs with headings and one example row each).
// 3. Deploy > New deployment > Web app: Execute as "Me", Who has access "Anyone".
//    After pasting a newer version: Deploy > Manage deployments > pencil >
//    Version: New version > Deploy (keeps the same web address).
//
// Columns are matched by their HEADING, in any order, so you can rearrange
// them or add your own. Missing columns are added on the right when needed.
// Accepted headings (any capitals) are listed in ALIASES below.
//
// Barrel actions (Fill, Reading, Moved, Bottled, Note, Empty), adding/editing
// barrels and creating/editing blends all need the PIN.
// Bottle ratings from friends need no PIN.
// Anything in the Barrels, Blends and Log tabs can be seen by anyone who has
// the page link. Ratings are shown only as counts (names and comments stay
// private in the Sheet).

const CREW_PIN = "1234";   // <-- CHANGE THIS

const TABS = {
  Barrels: ["Barrel", "Size", "Type", "Location", "Notes"],
  Blends:  ["Blend", "Vintage", "Juiced", "Apples", "Gravity", "pH", "Yeast", "Notes"],
  Log:     ["Time", "Barrel", "Action", "Blend", "Yeast", "Gravity", "pH", "Bottles", "Priming", "Location", "Notes", "By"],
  Ratings: ["Time", "Bottle", "Rating", "Tags", "Name", "Comment"],
};
// Sheet tab names (left = what the page calls it, right = the tab name in YOUR Sheet).
const TAB_NAMES = {
  Barrels: "Barrels",
  Blends:  "Blends",
  Log:     "Fermentation and bottling",
  Ratings: "Ratings",
};
// Other headings accepted for each column (lower case).
const ALIASES = {
  Barrels: { Barrel: ["vessel"], Notes: ["note"] },
  Blends:  { Juiced: ["date juiced", "date"], Gravity: ["sg", "starting gravity", "og"], pH: ["starting ph"], Yeast: ["yiest"], Notes: ["note"] },
  Log:     { Time: ["date", "when", "date/time"], Barrel: ["vessel"], Action: ["what"], Yeast: ["yiest"], Gravity: ["sg"],
             Location: ["where"], Notes: ["note", "comment", "comments"], By: ["who", "name"] },
  Ratings: { Time: ["date"] },
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
    const head = rows.shift().map(function (h) { return canon(name, h); });
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
      appendByHeading(tab(ss, "Ratings"), "Ratings", { Time: now, Bottle: clean(d.bottle, 30).toUpperCase(), Rating: d.rating,
        Tags: clean(d.tags, 200), Name: clean(d.name, 40), Comment: clean(d.comment, 500) });
      return reply({ ok: true });
    }

    if (d.kind === "log") {
      if (String(d.pin) !== String(CREW_PIN)) return reply({ ok: false, error: "Wrong PIN number" });
      if (ACTIONS.indexOf(d.action) < 0) return reply({ ok: false, error: "Unknown action" });
      appendByHeading(tab(ss, "Log"), "Log", { Time: now, Barrel: clean(d.barrel, 10).toUpperCase(), Action: d.action,
        Blend: clean(d.blend, 20).toUpperCase(), Yeast: clean(d.yeast, 60), Gravity: num(d.gravity), pH: num(d.ph),
        Bottles: num(d.bottles), Priming: clean(d.priming, 40), Location: clean(d.location, 60),
        Notes: clean(d.note, 500), By: clean(d.by, 40) });
      // Keep the Barrels tab's Location up to date when a barrel moves.
      if (d.action === "Moved" && d.location) {
        upsertByHeading(tab(ss, "Barrels"), "Barrels", "Barrel",
          { Barrel: clean(d.barrel, 10).toUpperCase(), Location: clean(d.location, 60) }, "edit");
      }
      return reply({ ok: true });
    }

    if (d.kind === "barrel") {
      if (String(d.pin) !== String(CREW_PIN)) return reply({ ok: false, error: "Wrong PIN number" });
      const id = clean(d.barrel, 10).toUpperCase();
      if (!/^[BD]\d{2,3}$/.test(id)) return reply({ ok: false, error: "Barrel numbers look like B21 or D03" });
      const r = upsertByHeading(tab(ss, "Barrels"), "Barrels", "Barrel", { Barrel: id, Size: numOrText(d.size, 20),
        Type: clean(d.type, 40), Location: clean(d.location, 60), Notes: clean(d.notes, 500) }, d.isNew ? "new" : "edit");
      if (!r.ok) return reply(r);
      if (d.moved && d.location) appendByHeading(tab(ss, "Log"), "Log", { Time: now, Barrel: id, Action: "Moved",
        Location: clean(d.location, 60), By: clean(d.by, 40) });
      return reply({ ok: true });
    }

    if (d.kind === "blend") {
      if (String(d.pin) !== String(CREW_PIN)) return reply({ ok: false, error: "Wrong PIN number" });
      const id = clean(d.blend, 10).toUpperCase();
      if (!/^\d{2}[A-Z]{1,2}$/.test(id)) return reply({ ok: false, error: "Blend codes look like 26D" });
      const r = upsertByHeading(tab(ss, "Blends"), "Blends", "Blend", { Blend: id, Juiced: dayDate(d.juiced),
        Apples: clean(d.apples, 1000), Gravity: num(d.gravity), pH: num(d.ph), Yeast: clean(d.yeast, 60),
        Notes: clean(d.notes, 1000) }, d.isNew ? "new" : "edit", { Juiced: "d mmm yyyy", Gravity: "0.000" });
      return reply(r);
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

// The standard name for a column heading, e.g. "Yiest" -> "Yeast", "Size (litres)" -> "Size".
function canon(tabKey, h) {
  const t = String(h == null ? "" : h).trim(), low = t.toLowerCase();
  if (!t) return "";
  const std = TABS[tabKey] || [], al = ALIASES[tabKey] || {};
  for (let i = 0; i < std.length; i++) {
    const c = std[i];
    if (low === c.toLowerCase() || (al[c] || []).indexOf(low) >= 0) return c;
    if (c === "Size" && low.indexOf("size") === 0) return c;
  }
  return t;
}

// Finds the column for each value by heading.
// Adds a column on the right for any (non-empty) value whose heading isn't there yet.
function columnsFor(sh, tabKey, values) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const where = {};
  head.forEach(function (h, i) { const c = canon(tabKey, h); if (c && !(c in where)) where[c] = i; });
  Object.keys(values).forEach(function (k) {
    if (values[k] === "" || values[k] == null || k in where) return;
    const col = head.length;
    head.push(k);
    sh.getRange(1, col + 1).setValue(k).setFontWeight("bold");
    where[k] = col;
  });
  return { head: head, where: where };
}

// Writes one new row, putting each value under its matching heading.
function appendByHeading(sh, tabKey, values) {
  const cols = columnsFor(sh, tabKey, values);
  const row = cols.head.map(function () { return ""; });
  Object.keys(values).forEach(function (k) { if (k in cols.where) row[cols.where[k]] = values[k]; });
  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}

// Updates the row whose key column (e.g. Barrel = "B07") matches, or adds it.
// mode "new": refuse if it already exists.  mode "edit": refuse if it doesn't.
// Only the given columns are touched; an empty value clears that cell.
function upsertByHeading(sh, tabKey, keyName, values, mode, formats) {
  const cols = columnsFor(sh, tabKey, values);
  const id = String(values[keyName]).trim().toUpperCase();
  const kc = cols.where[keyName];
  let row = 0;
  if (kc != null && sh.getLastRow() > 1) {
    const keys = sh.getRange(2, kc + 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim().toUpperCase() === id) { row = i + 2; break; }
    }
  }
  if (row && mode === "new") return { ok: false, error: id + " already exists" };
  if (!row && mode === "edit") return { ok: false, error: id + " isn't in the Sheet" };
  if (!row) row = sh.getLastRow() + 1;
  Object.keys(values).forEach(function (k) {
    if (!(k in cols.where)) return;
    const cell = sh.getRange(row, cols.where[k] + 1);
    cell.setValue(values[k]);
    if (formats && formats[k] && values[k] !== "") cell.setNumberFormat(formats[k]);
  });
  return { ok: true };
}

// "2026-11-04" from the page's date picker -> a real date (no US/UK mix-ups).
function dayDate(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ""));
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12) : "";
}
function numOrText(x, max) {
  const n = num(x);
  return n === "" ? clean(x, max) : n;
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
