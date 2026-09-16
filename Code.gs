/*
 * The Next Repertoire - workshop backend.
 * Google Apps Script web app bound to a Google Sheet.
 *
 * Setup: see SETUP.md. Set the facilitator passcode once by running setAdminKey() below.
 *
 * GET  ?action=state              -> {gates:{unlocked:[],current:""}, updated}
 * GET  ?action=summary            -> aggregate counts for the live panels on participant pages
 * GET  ?action=admin&key=PASSCODE -> counts and recent rows for the facilitator panel
 * POST {action:"submit", form, pid, table, data:{...}}
 * POST {action:"setGates", key, gates:{unlocked:[],current:"",show:[]}}
 */

var FORMS = ["pre", "demographics", "quadrant", "interests", "table_reports", "commit", "post"];
var MAX_ROWS_RETURNED = 600;

/* Run this once from the editor to set the facilitator passcode. */
function setAdminKey() {
  var key = "change-me";
  PropertiesService.getScriptProperties().setProperty("ADMIN_KEY", key);
  Logger.log("Facilitator passcode set.");
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || "state";
  if (action === "state") return json({ gates: getGates(), updated: getUpdated() });
  if (action === "summary") return json(summaryData());
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
  var out = { counts: {}, rows: {}, gates: getGates() };
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

/* ---------- helpers ---------- */
function checkKey(k) {
  var stored = PropertiesService.getScriptProperties().getProperty("ADMIN_KEY");
  return !!stored && String(k || "") === stored;
}
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
