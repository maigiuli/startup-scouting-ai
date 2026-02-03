
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Startup Scouting AI')
    .addItem('Scouting accelerators', 'scoutingAccelerators')
    .addItem('Find portfolio pages for each accelerator', 'updateAllPortfolioUrls')
    .addItem('Update startups of accelerators', 'updateStartups')
    .addItem('Enrich missing information', 'enrichMissingData')
    .addItem('Generate missing value proposition', 'generateValueProposition')
    .addToUi();
}

function writeLog(status, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Logs");
  if (!logSheet) {
    logSheet = ss.insertSheet("Logs");
    logSheet.appendRow(["Timestamp", "Status", "Message"]);
    logSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#d9d9d9");
  }
  
  logSheet.appendRow([new Date(), status, message]);
}
 


