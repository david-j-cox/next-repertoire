/*
 * The Next Repertoire - workshop backend.
 * Google Apps Script web app bound to a Google Sheet.
 *
 * Setup: see SETUP.md. Set the facilitator passcode once by running setAdminKey() below.
 *
 * GET  ?action=state              -> {gates:{unlocked:[],current:""}, updated}
 * GET  ?action=summary            -> aggregate counts for the live panels on participant pages
 * GET  ?action=shift&table=A       -> that table's context shift, only after the reveal is opened
 * GET  ?action=myTable&pid=p-xxx   -> that participant's assigned table, once assignments are shown
 * POST {action:"assign", key, capacity} -> sort everyone into tables from the interest poll
 * GET  ?action=admin&key=PASSCODE -> counts and recent rows for the facilitator panel
 * POST {action:"submit", form, pid, table, data:{...}}
 * POST {action:"setGates", key, gates:{unlocked:[],current:"",show:[]}}
 */

var FORMS = ["pre", "demographics", "quadrant", "interests", "table_reports", "commit", "post"];
var MAX_ROWS_RETURNED = 600;

/*
 * Run this once from the editor to set the facilitator passcode.
 * Put the passcode on the `key` line, Run, then set it back to "change-me".
 * Running it with the placeholder still in place does nothing, so an accidental
 * Run cannot wipe the passcode you already set.
 */
function setAdminKey() {
  var key = "change-me";
  if (!key || key === "change-me") {
    Logger.log("Nothing changed. Put your passcode on the key line in setAdminKey, then Run again.");
    return;
  }
  PropertiesService.getScriptProperties().setProperty("ADMIN_KEY", key);
  Logger.log("Facilitator passcode set.");
}

/* Reports whether a passcode is set, without printing it. */
function checkAdminKeyStatus() {
  var k = PropertiesService.getScriptProperties().getProperty("ADMIN_KEY");
  Logger.log(!k ? "No passcode set."
    : (k === "change-me" ? "WARNING: the passcode is still the placeholder \"change-me\". Set a real one."
                         : "A passcode is set (" + k.length + " characters)."));
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || "state";
  if (action === "state") return json({ gates: getGates(), updated: getUpdated() });
  if (action === "summary") return json(summaryData());
  if (action === "shift") return json(shiftFor(p.table));
  if (action === "myTable") return json(myTable(p.pid));
  if (action === "admin") {
    if (!checkKey(p.key)) return json({ error: "unauthorized" });
    return json(adminData());
  }
  return json({ error: "unknown action" });
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ error: "bad json" }); }
  if (body.action === "submit") return json(submit(body));
  if (body.action === "assign") {
    if (!checkKey(body.key)) return json({ error: "unauthorized" });
    return json(assignTables(body.capacity));
  }
  if (body.action === "setGates") {
    if (!checkKey(body.key)) return json({ error: "unauthorized" });
    setGates(body.gates || {});
    return json({ ok: true, gates: getGates() });
  }
  return json({ error: "unknown action" });
}

/* ---------- gates ---------- */
function configSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("config");
  if (!sh) {
    sh = ss.insertSheet("config");
    sh.getRange("A1:B2").setValues([["gates", JSON.stringify({ unlocked: [], current: "welcome", show: [] })], ["updated", new Date().toISOString()]]);
  }
  return sh;
}
function getGates() {
  var cache = CacheService.getScriptCache();
  var c = cache.get("gates");
  if (c) return JSON.parse(c);
  var raw = configSheet().getRange("B1").getValue();
  var g;
  try { g = JSON.parse(raw); } catch (err) { g = { unlocked: [], current: "welcome" }; }
  if (!g || !g.unlocked) g = { unlocked: [], current: "welcome", show: [] };
  if (!g.show) g.show = [];
  cache.put("gates", JSON.stringify(g), 5);
  return g;
}
function getUpdated() {
  var v = configSheet().getRange("B2").getValue();
  return v ? String(v) : "";
}
function setGates(g) {
  var clean = { unlocked: Array.isArray(g.unlocked) ? g.unlocked.map(String) : [], current: String(g.current || "welcome"), show: Array.isArray(g.show) ? g.show.map(String) : [] };
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = configSheet();
    sh.getRange("B1").setValue(JSON.stringify(clean));
    sh.getRange("B2").setValue(new Date().toISOString());
    CacheService.getScriptCache().remove("gates");
  } finally { lock.releaseLock(); }
}

/* ---------- submissions ---------- */
function submit(body) {
  var form = String(body.form || "");
  if (FORMS.indexOf(form) < 0) return { error: "unknown form" };
  var data = body.data && typeof body.data === "object" ? body.data : {};
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(form);
    if (!sh) { sh = ss.insertSheet(form); sh.appendRow(["timestamp", "pid", "table"]); sh.setFrozenRows(1); }
    var headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 3)).getValues()[0].map(String);
    var keys = Object.keys(data);
    keys.forEach(function (k) {
      if (headers.indexOf(k) < 0) { headers.push(k); sh.getRange(1, headers.length).setValue(k); }
    });
    var row = headers.map(function (h) {
      if (h === "timestamp") return new Date();
      if (h === "pid") return String(body.pid || "");
      if (h === "table") return String(body.table || "");
      var v = data[h];
      if (v === undefined || v === null) return "";
      return typeof v === "object" ? JSON.stringify(v) : v;
    });
    sh.appendRow(row);
  } finally { lock.releaseLock(); }
  return { ok: true };
}

/* ---------- facilitator data ---------- */
function adminData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = { counts: {}, rows: {}, gates: getGates(), assigned: assignmentTally() };
  FORMS.forEach(function (f) {
    var sh = ss.getSheetByName(f);
    if (!sh || sh.getLastRow() < 2) { out.counts[f] = 0; out.rows[f] = []; return; }
    var n = sh.getLastRow() - 1;
    out.counts[f] = n;
    var start = Math.max(2, sh.getLastRow() - MAX_ROWS_RETURNED + 1);
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    var vals = sh.getRange(start, 1, sh.getLastRow() - start + 1, sh.getLastColumn()).getValues();
    out.rows[f] = vals.map(function (r) {
      var o = {};
      headers.forEach(function (h, i) { o[h] = r[i] instanceof Date ? r[i].toISOString() : r[i]; });
      return o;
    });
  });
  return out;
}

/* Aggregates for participant-facing live panels. No free text except table-level reports. Cached 10 s. */
function summaryData() {
  var cache = CacheService.getScriptCache();
  var c = cache.get("summary");
  if (c) return JSON.parse(c);
  var d = adminData();
  var out = { counts: d.counts, interests: latestPerPid(d.rows.interests || []).map(function (r) { return { first: r["fint-first"], second: r["fint-second"], role: r["fint-role"] }; }),
              demographics: (d.rows.demographics || []).map(function (r) { return { role: r["fdem-role"], setting: r["fdem-setting"], years: r["fdem-years"], org: r["fdem-org"], ai: r["fdem-ai"] }; }),
              pre: (d.rows.pre || []).map(function (r) { return { c1: r["fpre-c1"], c2: r["fpre-c2"], c3: r["fpre-c3"], c4: r["fpre-c4"] }; }),
              post: (d.rows.post || []).map(function (r) { return { c1: r["fpost-c1"], c2: r["fpost-c2"], c3: r["fpost-c3"], c4: r["fpost-c4"] }; }),
              table_reports: d.rows.table_reports || [] };
  cache.put("summary", JSON.stringify(out), 10);
  return out;
}
function latestPerPid(rows) {
  var m = {}; rows.forEach(function (r) { m[r.pid || Math.random()] = r; });
  return Object.keys(m).map(function (k) { return m[k]; });
}

function assignmentTally() {
  var map = assignmentMap(), counts = {}, n = 0;
  TABLE_IDS.forEach(function (t) { counts[t] = 0; });
  Object.keys(map).forEach(function (pid) {
    if (counts[map[pid].table] !== undefined) counts[map[pid].table]++;
    n++;
  });
  return { n: n, counts: counts };
}

/* Sealed context shifts. Held server-side so the text is not in the page source before the reveal. */
var SHIFTS = {
  A: { t: "The source has changed",
       c: "The workflow now receives an audio transcript containing overlapping speakers and an ambiguous speaker label. A sentence could have been said by either the caregiver or the clinician. The draft treats it as a confirmed clinician observation.",
       q: "Where does uncertainty remain visible? Who verifies the source before the statement enters the record?" },
  B: { t: "The implementer changed",
       c: "The family member who practiced the proposed routine is temporarily unavailable. A new caregiver can participate only briefly and has not received training. The drafted plan still assumes the original implementation conditions.",
       q: "Which part of the workflow must pause or change before implementation? Whose input is now missing?" },
  C: { t: "A preference was inferred",
       c: "The system ranks goals using a structured intake form. A caregiver explains that the form was completed with language support and that the ranking does not reflect their stated priority. The team had treated the form as an authoritative preference measure.",
       q: "How will the workflow obtain and preserve meaningful preference information without treating an inferred score as consent?" },
  D: { t: "The expert is unavailable",
       c: "A new staff member can correctly repeat the AI-generated explanation, but a qualified supervisor is unexpectedly unavailable during the planned rehearsal. The workflow assumes that a completed lesson is enough to begin independent performance.",
       q: "What is the smallest safe practice opportunity now? What requires observed performance or supervisor availability?" },
  E: { t: "The population shifted",
       c: "The organization adds a service setting with different referral patterns and missing-data practices. The prediction tool is unchanged. Its overall dashboard remains stable, but no one has tested performance for the new setting.",
       q: "What evidence is needed before these predictions inform decisions in the new setting?" },
  F: { t: "The incentive changed",
       c: "Management begins rewarding schedulers for accepting the first AI-generated schedule. A family requests a different arrangement. Staff can technically override the recommendation, but overrides now lower their performance score.",
       q: "Does the human override still function? What technical and organizational changes are both needed?" }
};
function shiftFor(table) {
  var g = getGates();
  if (g.show.indexOf("reveal") < 0) return { sealed: true };
  var key = String(table || "").toUpperCase();
  if (!SHIFTS[key]) return { sealed: false, unknown: true };
  return { sealed: false, table: key, shift: SHIFTS[key] };
}

/* ---------- table assignment ---------- */
var TABLE_IDS = ["A", "B", "C", "D", "E", "F"];

/**
 * Sort everyone who answered the interest poll into tables.
 * Honors first choice, then second choice, then fills the emptiest table,
 * keeping every table within `capacity` (default: an even split).
 * Earlier responses are seated first when a table is oversubscribed.
 */
function assignTables(capacity) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("interests");
  if (!sh || sh.getLastRow() < 2) return { error: "No interest-poll responses yet." };

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var col = {};
  headers.forEach(function (h, i) { col[h] = i; });

  var latest = {};
  vals.forEach(function (r) {
    var pid = String(r[col["pid"]] || "");
    if (!pid) return;
    latest[pid] = {
      pid: pid,
      ts: r[col["timestamp"]] instanceof Date ? r[col["timestamp"]].getTime() : 0,
      first: String(r[col["fint-first"]] || "").toUpperCase(),
      second: String(r[col["fint-second"]] || "").toUpperCase()
    };
  });
  var people = Object.keys(latest).map(function (k) { return latest[k]; });
  people.sort(function (a, b) { return a.ts - b.ts; });
  if (!people.length) return { error: "No interest-poll responses yet." };

  var cap = Number(capacity) > 0 ? Number(capacity) : Math.ceil(people.length / TABLE_IDS.length);
  var counts = {};
  TABLE_IDS.forEach(function (t) { counts[t] = 0; });

  function seat(p, t, basis) { counts[t]++; p.table = t; p.basis = basis; }

  people.forEach(function (p) {
    if (counts[p.first] !== undefined && counts[p.first] < cap) seat(p, p.first, "first choice");
  });
  people.forEach(function (p) {
    if (!p.table && counts[p.second] !== undefined && counts[p.second] < cap) seat(p, p.second, "second choice");
  });
  people.forEach(function (p) {
    if (p.table) return;
    var t = TABLE_IDS.slice().sort(function (a, b) { return counts[a] - counts[b]; })[0];
    seat(p, t, "balanced");
  });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var a = ss.getSheetByName("assignments");
    if (!a) { a = ss.insertSheet("assignments"); }
    a.clear();
    a.appendRow(["timestamp", "pid", "table", "basis"]);
    a.setFrozenRows(1);
    var now = new Date();
    a.getRange(2, 1, people.length, 4).setValues(people.map(function (p) {
      return [now, p.pid, p.table, p.basis];
    }));
    CacheService.getScriptCache().remove("assign");
  } finally { lock.releaseLock(); }

  var basis = { "first choice": 0, "second choice": 0, "balanced": 0 };
  people.forEach(function (p) { basis[p.basis]++; });
  return { ok: true, n: people.length, capacity: cap, counts: counts, basis: basis };
}

/* The participant's own seat. Withheld until the facilitator shows assignments. */
function myTable(pid) {
  var g = getGates();
  if (g.show.indexOf("assignment") < 0) return { pending: true };
  pid = String(pid || "");
  if (!pid) return { pending: true };
  var map = assignmentMap();
  if (!map[pid]) return { pending: true, unassigned: true };
  return { table: map[pid].table, basis: map[pid].basis };
}

function assignmentMap() {
  var cache = CacheService.getScriptCache();
  var c = cache.get("assign");
  if (c) return JSON.parse(c);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("assignments");
  var map = {};
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(function (r) {
      if (r[1]) map[String(r[1])] = { table: String(r[2]), basis: String(r[3]) };
    });
  }
  cache.put("assign", JSON.stringify(map), 15);
  return map;
}

/* ---------- helpers ---------- */
function checkKey(k) {
  var stored = PropertiesService.getScriptProperties().getProperty("ADMIN_KEY");
  return !!stored && String(k || "") === stored;
}
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
