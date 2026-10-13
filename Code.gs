/*
 * The Next Repertoire - workshop backend.
 * Google Apps Script web app bound to a Google Sheet.
 *
 * Setup: see SETUP.md. Set the facilitator passcode once by running setAdminKey() below.
 *
 * GET  ?action=state              -> {gates:{unlocked:[],current:""}, updated}
 *      &group=A-1&since=ms        -> also {group:{g, now, d:{fieldId:{v,t,p}}}}, the table's shared entries changed after `since`
 * POST {action:"groupPut", group, pid, fields:{...}} -> store shared table entries (table sync)
 * GET  ?action=summary            -> aggregate counts for the live panels on participant pages
 * GET  ?action=shift&table=A       -> that table's context shift, only after the reveal is opened
 * GET  ?action=myTable&pid=p-xxx   -> that participant's assigned table and group, once assignments are shown
 * POST {action:"assign", key, tableSize, maxTables} -> sort everyone into topic tables (A-F) and table groups (A-1, A-2)
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
  if (action === "state") {
    var st = { gates: getGates(), updated: getUpdated() };
    if (p.group) st.group = groupDelta(p.group, p.since);
    return json(st);
  }
  if (action === "summary") return json(summaryData());
  if (action === "shift") return json(shiftFor(p.table));
  if (action === "myTable") return json(myTable(p.pid));
  if (action === "admin") {
    if (!checkKey(p.key)) return json({ error: "unauthorized" });
    try { flushGroups(); } catch (err) {}
    return json(adminData());
  }
  return json({ error: "unknown action" });
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ error: "bad json" }); }
  if (body.action === "submit") return json(submit(body));
  if (body.action === "groupPut") return json(groupPut(body));
  if (body.action === "assign") {
    if (!checkKey(body.key)) return json({ error: "unauthorized" });
    return json(assignTables(body));
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
  var out = { counts: d.counts, interests: latestPerPid(d.rows.interests || []).map(function (r) { return { first: r["fint-first"], second: r["fint-second"], third: r["fint-third"], role: r["fint-role"] }; }),
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
  var map = assignmentMap(), counts = {}, groups = {}, n = 0;
  TABLE_IDS.forEach(function (t) { counts[t] = 0; });
  Object.keys(map).forEach(function (pid) {
    if (counts[map[pid].table] !== undefined) counts[map[pid].table]++;
    if (map[pid].group) groups[map[pid].group] = (groups[map[pid].group] || 0) + 1;
    n++;
  });
  return { n: n, counts: counts, groups: groups };
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
 * Sort everyone who answered the interest poll into topic tables (A-F), then split
 * each topic into at most `maxTables` physical tables of about `tableSize` people
 * (groups A-1, A-2, ...). Honors first choice, then second, then third, then fills
 * the emptiest topic. Earlier responses are seated first when a topic is full.
 * Within a topic, people are dealt round-robin by expected role so each table gets a mix.
 */
function assignTables(opts) {
  opts = opts || {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("interests");
  if (!sh || sh.getLastRow() < 2) return { error: "No interest-poll responses yet." };

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var col = {};
  headers.forEach(function (h, i) { col[h] = i; });
  function cell(r, h) { return col[h] === undefined ? "" : String(r[col[h]] || ""); }

  var latest = {};
  vals.forEach(function (r) {
    var pid = cell(r, "pid");
    if (!pid) return;
    latest[pid] = {
      pid: pid,
      ts: r[col["timestamp"]] instanceof Date ? r[col["timestamp"]].getTime() : 0,
      first: cell(r, "fint-first").toUpperCase(),
      second: cell(r, "fint-second").toUpperCase(),
      third: cell(r, "fint-third").toUpperCase(),
      role: cell(r, "fint-role")
    };
  });
  var people = Object.keys(latest).map(function (k) { return latest[k]; });
  people.sort(function (a, b) { return a.ts - b.ts; });
  if (!people.length) return { error: "No interest-poll responses yet." };

  var size = Number(opts.tableSize) > 0 ? Math.floor(Number(opts.tableSize)) : 6;
  var maxT = Number(opts.maxTables) > 0 ? Math.floor(Number(opts.maxTables)) : 2;
  var cap = size * maxT;
  var overflow = false;
  if (cap * TABLE_IDS.length < people.length) { cap = Math.ceil(people.length / TABLE_IDS.length); overflow = true; }

  var counts = {};
  TABLE_IDS.forEach(function (t) { counts[t] = 0; });
  function seat(p, t, basis) { counts[t]++; p.table = t; p.basis = basis; }

  [["first", "first choice"], ["second", "second choice"], ["third", "third choice"]].forEach(function (pass) {
    people.forEach(function (p) {
      var t = p[pass[0]];
      if (!p.table && counts[t] !== undefined && counts[t] < cap) seat(p, t, pass[1]);
    });
  });
  people.forEach(function (p) {
    if (p.table) return;
    var t = TABLE_IDS.slice().sort(function (a, b) { return counts[a] - counts[b]; })[0];
    seat(p, t, "balanced");
  });

  /* split each topic into physical tables */
  var groups = {};
  TABLE_IDS.forEach(function (t) {
    var members = people.filter(function (p) { return p.table === t; });
    if (!members.length) return;
    var k = Math.min(maxT, Math.max(1, Math.ceil(members.length / size)));
    members.sort(function (a, b) { return a.role < b.role ? -1 : (a.role > b.role ? 1 : a.ts - b.ts); });
    members.forEach(function (p, i) {
      p.group = t + "-" + ((i % k) + 1);
      groups[p.group] = (groups[p.group] || 0) + 1;
    });
  });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var a = ss.getSheetByName("assignments");
    if (!a) { a = ss.insertSheet("assignments"); }
    a.clear();
    a.appendRow(["timestamp", "pid", "table", "basis", "group"]);
    a.setFrozenRows(1);
    var now = new Date();
    a.getRange(2, 1, people.length, 5).setValues(people.map(function (p) {
      return [now, p.pid, p.table, p.basis, p.group];
    }));
    CacheService.getScriptCache().remove("assign");
  } finally { lock.releaseLock(); }

  var basis = { "first choice": 0, "second choice": 0, "third choice": 0, "balanced": 0 };
  people.forEach(function (p) { basis[p.basis]++; });
  return { ok: true, n: people.length, tableSize: size, maxTables: maxT, overflow: overflow, counts: counts, groups: groups, basis: basis };
}

/* The participant's own seat. Withheld until the facilitator shows assignments. */
function myTable(pid) {
  var g = getGates();
  if (g.show.indexOf("assignment") < 0) return { pending: true };
  pid = String(pid || "");
  if (!pid) return { pending: true };
  var map = assignmentMap();
  if (!map[pid]) return { pending: true, unassigned: true };
  return { table: map[pid].table, basis: map[pid].basis, group: map[pid].group || "" };
}

function assignmentMap() {
  var cache = CacheService.getScriptCache();
  var c = cache.get("assign");
  if (c) return JSON.parse(c);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("assignments");
  var map = {};
  if (sh && sh.getLastRow() > 1) {
    var w = Math.max(4, Math.min(5, sh.getLastColumn()));
    sh.getRange(2, 1, sh.getLastRow() - 1, w).getValues().forEach(function (r) {
      if (r[1]) map[String(r[1])] = { table: String(r[2]), basis: String(r[3]), group: String(r[4] || "") };
    });
  }
  cache.put("assign", JSON.stringify(map), 15);
  return map;
}

/* ---------- table sync ----------
 * Each physical table (A-1, A-2, ...) has one shared document of field values.
 * It lives in the script cache for speed and is copied to the group_notes tab
 * (on the first write after a quiet minute, and whenever the facilitator panel refreshes).
 */
var GROUP_TTL = 21600;   /* 6 hours, the cache maximum */
function cleanGroup(g) {
  g = String(g || "").toUpperCase();
  return /^[A-F]-[1-9]$/.test(g) ? g : "";
}
function groupsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("group_notes");
  if (!sh) { sh = ss.insertSheet("group_notes"); sh.appendRow(["group", "doc", "updated"]); sh.setFrozenRows(1); }
  return sh;
}
function groupDoc(g) {
  var c = CacheService.getScriptCache().get("grp:" + g);
  if (c) return JSON.parse(c);
  var sh = groupsSheet();
  if (sh.getLastRow() > 1) {
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === g) {
        try { var d = JSON.parse(rows[i][1]); CacheService.getScriptCache().put("grp:" + g, rows[i][1], GROUP_TTL); return d; } catch (err) { return {}; }
      }
    }
  }
  return {};
}
function writeGroupRows(docs) {
  var names = Object.keys(docs);
  if (!names.length) return;
  var sh = groupsSheet();
  var existing = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  var now = new Date();
  names.forEach(function (g) {
    var i = existing.indexOf(g);
    if (i >= 0) sh.getRange(i + 2, 2, 1, 2).setValues([[docs[g], now]]);
    else { sh.appendRow([g, docs[g], now]); existing.push(g); }
  });
}
function groupPut(body) {
  var g = cleanGroup(body.group);
  if (!g) return { error: "no group" };
  var fields = body.fields && typeof body.fields === "object" ? body.fields : null;
  if (!fields) return { error: "no fields" };
  var cache = CacheService.getScriptCache();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { error: "busy" };
  var now = Date.now(), s;
  try {
    var doc = groupDoc(g);
    Object.keys(fields).slice(0, 200).forEach(function (k) {
      if (!/^f[\w-]{1,40}$/.test(k)) return;
      var v = fields[k];
      v = typeof v === "boolean" ? v : String(v === null || v === undefined ? "" : v).slice(0, 4000);
      doc[k] = { v: v, t: now, p: String(body.pid || "").slice(0, 40) };
    });
    s = JSON.stringify(doc);
    if (s.length > 95000) return { error: "too large" };
    cache.put("grp:" + g, s, GROUP_TTL);
    if (!cache.get("grpsaved:" + g)) {
      var one = {}; one[g] = s;
      writeGroupRows(one);
      cache.put("grpsaved:" + g, "1", 60);
    }
  } finally { lock.releaseLock(); }
  return { ok: true, t: now };
}
/* Entries changed after `since` (ms). Clients ask with an overlap window and ignore repeats. */
function groupDelta(group, since) {
  var g = cleanGroup(group);
  if (!g) return { g: "", now: Date.now(), d: {} };
  var doc = groupDoc(g), t = Number(since) || 0, d = {};
  Object.keys(doc).forEach(function (k) { if (doc[k] && doc[k].t > t) d[k] = doc[k]; });
  return { g: g, now: Date.now(), d: d };
}
/* Copy every cached table document to the group_notes tab. Also runnable from the editor. */
function flushGroups() {
  var keys = [];
  TABLE_IDS.forEach(function (t) { for (var i = 1; i <= 4; i++) keys.push("grp:" + t + "-" + i); });
  var got = CacheService.getScriptCache().getAll(keys), docs = {};
  Object.keys(got).forEach(function (k) { docs[k.slice(4)] = got[k]; });
  writeGroupRows(docs);
}

/* ---------- helpers ---------- */
function checkKey(k) {
  var stored = PropertiesService.getScriptProperties().getProperty("ADMIN_KEY");
  return !!stored && String(k || "") === stored;
}
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
