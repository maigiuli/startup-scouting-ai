
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
    const response = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
    const content = response.getContentText();
    const $ = Cheerio.load(content);

    const rows = content.split('<tr');
    const existingWebsites= getExistingWebsites(sheet);
    let foundCount = 0;
    const MAX_BATCH = 10; 

    for (let i = 1; i < rows.length && foundCount < MAX_BATCH; i++) {      
      const location = $('tbody tr:nth-child('+ i + ') td:nth-child(5)').first().text()
      const country = $('tbody tr:nth-child('+ i + ') td:nth-child(6)').first().text()
      const fullLocation = location + ", " + country
      const aTag = $('tbody tr:nth-child('+ i + ') a').first()
      const website = aTag.attr("href")
      const name = aTag.text()

      const normalized = normalizeUrl(website);

      if (!name || !normalized) continue;

      if (existingWebsites.has(normalized)) {
        writeLog("INFO", "Skipped duplicate: " + website);
        continue;
      }
      const rowIndex = findFirstEmptyRow(sheet, 2);
      sheet.getRange(rowIndex, 1, 1, 3).setValues([[name, website, fullLocation]]); 
      existingWebsites.add(normalized);      
      foundCount++; 
      
    }
  } catch (e) {
    console.error("Error in scouting: " + e.toString());
  }
}

function getExistingWebsites(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();

  const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const set = new Set();

  values.forEach(row => {
    const rawUrl = row[0];
    const normalized = normalizeUrl(rawUrl);
    if (normalized) {
      set.add(normalized);
    }
  });
  return set;
}
function normalizeUrl(url) {
  if (!url) return null;
  return url
    .toLowerCase()
    .trim()
    .replace(/^https?:\/\//, '')
    .replace(/^www\./, '')
    .replace(/\/$/, '');
}

function findFirstEmptyRow(sheet, keyColumnIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;

  const range = sheet.getRange(2, keyColumnIndex, lastRow - 1, 1);
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) {
      return i + 2; // 
    }
  }
  return lastRow + 1; 
}


