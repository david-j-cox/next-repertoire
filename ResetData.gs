/*
 * Clear all workshop responses before a run-through or the live session.
 * Paste into the Sheet's Apps Script project as a new file (ResetData.gs). No redeploy needed.
 * Run previewReset() first to see what would be cleared, then resetAll().
 * Keeps each tab and its header row. Clears every response row, assignments, and table notes,
 * empties the backend cache, and turns off all room-result switches (sections you opened stay open).
 */
var RESET_TABS = FORMS.concat(["assignments", "group_notes"]);

function previewReset() { reset_(false); }
function resetAll() { reset_(true); }

function reset_(write) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  RESET_TABS.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { Logger.log(name + ": no tab yet."); return; }
    var n = Math.max(0, sh.getLastRow() - 1);
    Logger.log(name + ": " + n + " row" + (n === 1 ? "" : "s") + (write ? " cleared." : " to clear."));
    if (write && n > 0) sh.getRange(2, 1, n, Math.max(1, sh.getLastColumn())).clearContent();
  });
  var g = getGates();
  Logger.log("Room-result switches on now: " + (g.show.join(", ") || "none") + (write ? ". Turning all off." : ". Will turn all off."));
  if (!write) { Logger.log("Preview only. Nothing changed. Run resetAll() to clear."); return; }

  setGates({ unlocked: g.unlocked, current: g.current, show: [] });
  var keys = ["assign", "summary", "gates"];
  TABLE_IDS.forEach(function (t) { for (var i = 1; i <= 9; i++) keys.push("grp:" + t + "-" + i, "grpsaved:" + t + "-" + i); });
  CacheService.getScriptCache().removeAll(keys);
  Logger.log("Done. Cache cleared. Responses, seats, and table notes are empty.");
}
