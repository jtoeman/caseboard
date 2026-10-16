// Lives as its own file in the same Apps Script project as caseboard-reminders.gs,
// whose doPost hands {"kind":"movieleague_theme"} posts to saveThemeIdea_.
// The Movie League theme ideas. Body: {"kind":"movieleague_theme","idea":"...",
// "nickname":"...","player":"<anonymous id>","where":"past"|"suggest"}.
// Each idea is one row on the "Theme Suggestions" tab (made on first use).
var THEME_IDEAS_TAB = 'Theme Suggestions';

function saveThemeIdea_(body) {
  var idea = themeCell_(body.idea, 120);
  if (!idea) return ContentService.createTextOutput('empty idea');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(THEME_IDEAS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(THEME_IDEAS_TAB);
    sheet.appendRow(['Time', 'Idea', 'Nickname', 'Player id', 'Where']);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date(), idea, themeCell_(body.nickname, 40), themeCell_(body.player, 64), themeCell_(body.where, 20)]);
  return ContentService.createTextOutput('ok');
}

// Trimmed, length-capped text that Sheets won't treat as a formula.
function themeCell_(value, max) {
  var s = String(value || '').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
