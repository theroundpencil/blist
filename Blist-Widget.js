// Variables used by Scriptable.
// icon-color: yellow; icon-glyph: book-open;

// =============================================
//  BLIST - iOS Home Screen Widget
//  Reads stats from a public Google Sheet.
//  Shows a random book cover that rotates
//  every hour. Tap to open Blist.
// =============================================

// -- SETUP (two things to fill in) --

// 1. Your Google Sheet ID (from the spreadsheet URL)
//    URL looks like: https://docs.google.com/spreadsheets/d/SHEET_ID_HERE/edit
const SHEET_ID = "PASTE_YOUR_SHEET_ID_HERE";

// 2. Your Blist web app URL (so tapping the widget opens Blist)
const BLIST_URL = "PASTE_YOUR_APPS_SCRIPT_URL_HERE";

// -- COLORS --
var YELLOW = new Color("#FFD60A");
var GRAY   = new Color("#8E8E93");
var WHITE  = new Color("#FFFFFF");
var BG     = new Color("#111111");

// -- FETCH STATS from public Google Sheet --
async function fetchStats() {
  try {
    var url = "https://docs.google.com/spreadsheets/d/" + SHEET_ID + "/gviz/tq?tqx=out:csv&sheet=Widget";
    var req = new Request(url);
    req.timeoutInterval = 15;
    var csv = await req.loadString();

    // CSV format: row 0 = headers, row 1 = data
    // Column A = stats_json (big JSON string), Column B = updated_at
    var lines = csv.split("\n");
    if (lines.length < 2) return null;

    // Extract JSON from column A — CSV wraps it in quotes and escapes internal quotes
    var raw = lines[1];
    var jsonStr = extractFirstCSVField(raw);
    jsonStr = jsonStr.replace(/^"/, "").replace(/"$/, "").replace(/""/g, '"');

    return JSON.parse(jsonStr);
  } catch (e) {
    console.error("Fetch failed: " + e.message);
    return null;
  }
}

// Extract first field from a CSV line (handles quoted fields with commas inside)
function extractFirstCSVField(line) {
  if (line.charAt(0) !== '"') {
    var idx = line.indexOf(",");
    return idx >= 0 ? line.substring(0, idx) : line;
  }
  // Quoted field — find closing quote (not doubled)
  var i = 1;
  while (i < line.length) {
    if (line.charAt(i) === '"') {
      if (i + 1 < line.length && line.charAt(i + 1) === '"') {
        i += 2; // escaped quote
      } else {
        return line.substring(0, i + 1); // closing quote
      }
    } else {
      i++;
    }
  }
  return line; // fallback
}

// -- FETCH COVER IMAGE --
async function fetchCover(url) {
  try {
    var req = new Request(url);
    req.timeoutInterval = 10;
    return await req.loadImage();
  } catch (e) {
    console.error("Cover failed: " + e.message);
    return null;
  }
}

// -- BUILD WIDGET --
async function createWidget(stats) {
  var family = config.widgetFamily || "small";
  var w = new ListWidget();
  w.backgroundColor = BG;
  w.url = BLIST_URL;

  if (!stats) {
    w.setPadding(14, 14, 14, 14);
    var t = w.addText("Blist");
    t.font = Font.boldSystemFont(16);
    t.textColor = YELLOW;
    w.addSpacer(6);
    var m = w.addText("Could not load stats.\nCheck Sheet ID & sharing.");
    m.font = Font.systemFont(11);
    m.textColor = GRAY;
    return w;
  }

  // Load cover
  var coverImg = null;
  if (stats.randomBook && stats.randomBook.coverUrl) {
    coverImg = await fetchCover(stats.randomBook.coverUrl);
  }

  if (family === "small") {
    return buildSmall(w, stats, coverImg);
  } else {
    return buildMedium(w, stats, coverImg);
  }
}

// =============================================
//  SMALL — Full-bleed cover + overlay stats
// =============================================
function buildSmall(w, s, coverImg) {
  w.setPadding(0, 0, 0, 0);

  if (coverImg) {
    w.backgroundImage = coverImg;
    var g = new LinearGradient();
    g.locations = [0, 0.3, 0.65, 1.0];
    g.colors = [
      new Color("#000000", 0.0),
      new Color("#000000", 0.0),
      new Color("#000000", 0.55),
      new Color("#000000", 0.92)
    ];
    w.backgroundGradient = g;
  }

  w.addSpacer();

  var bottom = w.addStack();
  bottom.layoutVertically();
  bottom.setPadding(0, 12, 10, 12);

  // Book info
  if (s.randomBook) {
    var title = bottom.addText(s.randomBook.title);
    title.font = Font.boldSystemFont(11);
    title.textColor = WHITE;
    title.lineLimit = 1;
    bottom.addSpacer(2);
    var author = bottom.addText(s.randomBook.author || "");
    author.font = Font.systemFont(9);
    author.textColor = new Color("#CCCCCC");
    author.lineLimit = 1;
    bottom.addSpacer(6);
  }

  // Stat pills
  var row = bottom.addStack();
  row.centerAlignContent();
  addPill(row, String(s.total), "Total");
  row.addSpacer(4);
  addPill(row, String(s.byStatus["Reading"] || 0), "Reading");
  row.addSpacer(4);
  addPill(row, String(s.byOwnership["Buy"] || 0), "Buy");

  return w;
}

// =============================================
//  MEDIUM — Cover left + stats grid right
// =============================================
function buildMedium(w, s, coverImg) {
  w.setPadding(12, 12, 12, 12);

  var main = w.addStack();
  main.centerAlignContent();

  // Left: cover
  if (coverImg) {
    var imgStack = main.addStack();
    imgStack.cornerRadius = 8;
    imgStack.size = new Size(80, 120);
    var img = imgStack.addImage(coverImg);
    img.imageSize = new Size(80, 120);
    img.cornerRadius = 8;
    img.applyFillingContentMode();
  } else {
    var ph = main.addStack();
    ph.backgroundColor = new Color("#FFD60A", 0.15);
    ph.cornerRadius = 8;
    ph.size = new Size(80, 120);
    ph.layoutVertically();
    ph.addSpacer();
    var phT = ph.addText("Blist");
    phT.font = Font.boldSystemFont(14);
    phT.textColor = YELLOW;
    phT.centerAlignText();
    ph.addSpacer();
  }

  main.addSpacer(12);

  // Right: stats
  var right = main.addStack();
  right.layoutVertically();

  // Header row
  var header = right.addStack();
  header.centerAlignContent();
  var logo = header.addText("Blist");
  logo.font = Font.boldSystemFont(16);
  logo.textColor = YELLOW;
  header.addSpacer();
  var avg = header.addText("* " + (s.avgRating || "-"));
  avg.font = Font.mediumSystemFont(12);
  avg.textColor = WHITE;

  right.addSpacer(4);

  // Book title + author
  if (s.randomBook) {
    var bTitle = right.addText(s.randomBook.title);
    bTitle.font = Font.mediumSystemFont(10);
    bTitle.textColor = new Color("#CCCCCC");
    bTitle.lineLimit = 1;
    var bAuthor = right.addText(s.randomBook.author || "");
    bAuthor.font = Font.systemFont(9);
    bAuthor.textColor = GRAY;
    bAuthor.lineLimit = 1;
  }

  right.addSpacer(6);

  // Stats row 1
  var r1 = right.addStack();
  addStatCell(r1, String(s.total), "Total");
  r1.addSpacer();
  addStatCell(r1, String(s.byStatus["Reading"] || 0), "Reading");
  r1.addSpacer();
  addStatCell(r1, String(s.byStatus["To Read"] || 0), "To Read");
  r1.addSpacer();
  addStatCell(r1, String(s.byStatus["Read"] || 0), "Read");

  right.addSpacer(4);

  // Stats row 2
  var r2 = right.addStack();
  addStatCell(r2, String(s.byOwnership["Own"] || 0), "Own");
  r2.addSpacer();
  addStatCell(r2, String(s.byOwnership["Buy"] || 0), "Buy");
  r2.addSpacer();
  addStatCell(r2, String(s.recommendCount || 0), "Liked");
  r2.addSpacer();
  addStatCell(r2, String(s.byStatus["Don't Read"] || 0), "Skip");

  return w;
}

// -- HELPERS --
function addPill(stack, value, label) {
  var pill = stack.addStack();
  pill.backgroundColor = new Color("#000000", 0.55);
  pill.cornerRadius = 10;
  pill.setPadding(3, 7, 3, 7);
  pill.layoutVertically();
  pill.centerAlignContent();
  var v = pill.addText(value);
  v.font = Font.boldSystemFont(12);
  v.textColor = WHITE;
  v.centerAlignText();
  var l = pill.addText(label);
  l.font = Font.systemFont(7);
  l.textColor = GRAY;
  l.centerAlignText();
}

function addStatCell(stack, value, label) {
  var cell = stack.addStack();
  cell.layoutVertically();
  cell.centerAlignContent();
  var v = cell.addText(value);
  v.font = Font.boldSystemFont(15);
  v.textColor = WHITE;
  var l = cell.addText(label);
  l.font = Font.mediumSystemFont(8);
  l.textColor = GRAY;
}

// -- RUN --
var stats = await fetchStats();
var widget = await createWidget(stats);

if (config.runsInWidget) {
  Script.setWidget(widget);
} else {
  var family = config.widgetFamily || "small";
  if (family === "medium") {
    widget.presentMedium();
  } else {
    widget.presentSmall();
  }
}

Script.complete();
