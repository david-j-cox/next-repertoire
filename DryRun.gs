/*
 * Table-sort dry run. Paste into the Sheet's Apps Script project as a new file (DryRun.gs).
 * No redeploy needed. Run dryRunSort() from the editor and read the log.
 * It adds 14 fake interest-poll rows (pid starts with "dryrun-"), runs the real sort
 * with tables of 3, logs the result, then removes the fake rows and clears the assignments tab.
 */
function dryRunSort() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("interests");
  /* Same layout the site creates on the first real submission. */
  if (!sh) { sh = ss.insertSheet("interests"); sh.appendRow(["timestamp", "pid", "table"]); sh.setFrozenRows(1); Logger.log("Created the interests tab."); }
  var headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 3)).getValues()[0].map(String);
  ["fint-first", "fint-second", "fint-third", "fint-role"].forEach(function (h) {
    if (headers.indexOf(h) < 0) { headers.push(h); sh.getRange(1, headers.length).setValue(h); }
  });

  var picks = [
    ["A", "B", "C"], ["A", "C", "D"], ["A", "B", "E"], ["A", "D", "F"], ["A", "B", "C"], ["A", "E", "F"], ["A", "C", "B"],
    ["B", "A", "C"], ["C", "D", "A"], ["D", "A", "B"], ["E", "F", "A"], ["F", "E", "D"], ["B", "C", "D"], ["E", "A", "B"]
  ];
  var roles = ["Designer", "Implementer", "Supporter", "Not sure"];
  var now = Date.now();
  var rows = picks.map(function (p, i) {
    var r = headers.map(function () { return ""; });
    r[headers.indexOf("timestamp")] = new Date(now + i * 1000);
    r[headers.indexOf("pid")] = "dryrun-" + (i + 1);
    r[headers.indexOf("fint-first")] = p[0];
    r[headers.indexOf("fint-second")] = p[1];
    r[headers.indexOf("fint-third")] = p[2];
    r[headers.indexOf("fint-role")] = roles[i % roles.length];
    return r;
  });
  var start = sh.getLastRow() + 1;
  sh.getRange(start, 1, rows.length, headers.length).setValues(rows);
  SpreadsheetApp.flush();

  try {
    var res = assignTables({ tableSize: 3, maxTables: 2 });
    Logger.log("Sort result: " + JSON.stringify(res));
    var a = ss.getSheetByName("assignments");
    if (a && a.getLastRow() > 1) {
      a.getRange(2, 1, a.getLastRow() - 1, 5).getValues().forEach(function (r) {
        Logger.log(r[1] + " -> table " + r[2] + " (" + r[3] + "), seat " + r[4]);
      });
    }
    Logger.log("Expected: 7 people want A, so A is capped at 6 and splits into A-1 and A-2; the 7th person gets their second choice.");
  } finally {
    var all = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues() : [];
    var pidCol = headers.indexOf("pid");
    /* Keep real rows, drop fake ones. Rewrites instead of deleting, since Sheets will not delete every unfrozen row. */
    var keep = all.filter(function (r) { return String(r[pidCol]).indexOf("dryrun-") !== 0; });
    if (all.length) sh.getRange(2, 1, all.length, headers.length).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, headers.length).setValues(keep);
    var a2 = ss.getSheetByName("assignments");
    if (a2) a2.clear();
    CacheService.getScriptCache().remove("assign");
    Logger.log("Cleaned up: fake rows removed and assignments tab cleared.");
  }
}
