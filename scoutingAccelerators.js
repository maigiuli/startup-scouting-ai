
function scoutingAccelerators() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Accelerators');

  const url = "https://rankings.ft.com/incubator-accelerator-programmes-europe"; 

  
  try {
    let lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      sheet.appendRow(["Name", "Website", "Location"]);
      sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#f3f3f3");
    }

    let startIndex = (lastRow <= 1) ? 1 : lastRow + 1; 

    console.log("Sheet has " + lastRow + " lines. I start to read the HTML from the position: " + startIndex);
    writeLog("INFO", "Sheet has " + lastRow + " lines. I start to read the HTML from the position: " + startIndex);
    if (lastRow==151) writeLog("INFO", "No more accelerators to be added")


    const response = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
    const content = response.getContentText();
    const $ = Cheerio.load(content);

    const rows = content.split('<tr');
    let foundCount = 0;
    const MAX_BATCH = 10; 
  

    for (let i = startIndex; i < rows.length && foundCount < MAX_BATCH; i++) {      
      const location = $('tbody tr:nth-child('+ i + ') td:nth-child(5)').first().text()
      const country = $('tbody tr:nth-child('+ i + ') td:nth-child(6)').first().text()
      const fullLocation = location + ", " + country

      const aTag = $('tbody tr:nth-child('+ i + ') a').first()
      const website = aTag.attr("href")
      const name = aTag.text()
      
      if (name && website){
        sheet.appendRow([name, website, fullLocation]);        
        foundCount++; 
      }
    }
  } catch (e) {
    console.error("Error in scouting: " + e.toString());
  }
}
   