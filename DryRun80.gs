/*
 * 80-person table-sort dry run for the new mix rules. Paste into the Sheet's Apps Script
 * project (ChironWorkshop) as a new file (DryRun80.gs). No redeploy needed.
 * Run dryRun80() and read the log. It adds 80 fake interest-poll rows (pid starts with "dryrun-"),
 * runs the real sort with tables of 6 and up to 5 tables per topic, logs each table's mix,
 * then removes the fake rows and clears the assignments tab. Real responses are kept.
 * Fake room: one organization sends 10 people, a few send 3 to 5, the rest come alone;
 * about 1 in 4 people are Leadership or BCBA; topic B is the most popular.
 */
function dryRun80() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("interests");
  if (!sh) { sh = ss.insertSheet("interests"); sh.appendRow(["timestamp", "pid", "table"]); sh.setFrozenRows(1); }
  var headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 3)).getValues()[0].map(String);
  ["fint-first", "fint-second", "fint-third", "fint-role", "fint-position", "fint-org"].forEach(function (h) {
    if (headers.indexOf(h) < 0) { headers.push(h); sh.getRange(1, headers.length).setValue(h); }
  });
  var T = ["A", "B", "C", "D", "E", "F"], firsts = "BBBBACDEFBACDEFBBACDEF";
  var POS = ["Leadership or administration", "BCBA-D or BCBA", "BCaBA", "RBT or behavior technician", "Student or trainee", "Other"];
  var posCycle = [1, 3, 3, 2, 0, 3, 4, 3, 1, 5, 3, 2];
  var now = Date.now(), rows = [];
  for (var i = 0; i < 80; i++) {
    var f = firsts.charAt(i % firsts.length), s = T[(T.indexOf(f) + 1 + (i % 3)) % 6], th = T[(T.indexOf(f) + 3) % 6];
    var org = i < 10 ? "org-bigclinic" : i < 15 ? "org-midA" : i < 19 ? "org-midB" : i < 22 ? "org-midC" : "org-solo" + i;
    var r = headers.map(function () { return ""; });
    r[headers.indexOf("timestamp")] = new Date(now + i * 1000);
    r[headers.indexOf("pid")] = "dryrun-" + (i + 1);
    r[headers.indexOf("fint-first")] = f; r[headers.indexOf("fint-second")] = s; r[headers.indexOf("fint-third")] = th;
    r[headers.indexOf("fint-role")] = ["Designer", "Implementer", "Supporter", "Not sure"][i % 4];
    r[headers.indexOf("fint-position")] = POS[posCycle[i % posCycle.length]];
    r[headers.indexOf("fint-org")] = org;
    rows.push(r);
  }
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  SpreadsheetApp.flush();
  try {
    var res = assignTables({ tableSize: 6, maxTables: 5 });
    if (!res.ok) { Logger.log("Sort error: " + JSON.stringify(res)); return; }
    Logger.log(res.n + " seated. " + JSON.stringify(res.basis));
    Object.keys(res.mix).sort().forEach(function (g) {
      var m = res.mix[g];
      Logger.log(g + ": " + m.n + " people, " + m.drivers + " Leadership/BCBA, max " + m.maxSameOrg + " from one org");
    });
    Logger.log(res.flags.length ? "FLAGS: " + res.flags.join("; ") : "No flags: every table has a Leadership or BCBA participant and no more than 2 from one organization.");
  } finally {
    var all = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues() : [];
    var pidCol = headers.indexOf("pid");
    var keep = all.filter(function (r) { return String(r[pidCol]).indexOf("dryrun-") !== 0; });
    if (all.length) sh.getRange(2, 1, all.length, headers.length).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, headers.length).setValues(keep);
    var a2 = ss.getSheetByName("assignments"); if (a2) a2.clear();
    CacheService.getScriptCache().remove("assign");
    Logger.log("Cleaned up: fake rows removed and assignments tab cleared.");
  }
}
