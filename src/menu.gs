// ════════════════════════════════════════════════════════
// MENU & TRIGGERS
// ════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙ Master CHI')
    .addItem('📊 Build all trend sheets', 'buildAllTrends')
    .addSeparator()
    .addItem('🔗 Create Testing Scorecard list (first run only)', 'pushToClickUp')
    .addItem('🔧 Setup fields on Testing list', 'setupFieldsOnTestingList')
    .addItem('🔄 Update Testing Scorecard', 'updateTestingScorecard')
    .addSeparator()
    .addItem('⏱ Set up 24-hour auto-sync', 'setupDailySync')
    .addItem('⏹ Stop auto-sync', 'removeDailySync')
    .addSeparator()
    .addItem('🔍 Diagnose Testing list fields', 'diagnoseTestingList')
    .addToUi();
}

function setupDailySync() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'updateTestingScorecard') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('updateTestingScorecard').timeBased().everyDays(1).atHour(3).create();
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Auto-sync active — runs daily at 3 AM.', '✅');
}

function removeDailySync() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'updateTestingScorecard') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(
    removed > 0 ? '✅ Auto-sync stopped.' : 'No trigger found.',
    removed > 0 ? '✅' : 'ℹ️'
  );
}
