
function updateAllPortfolioUrls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Accelerators"); 

  let targetCell = sheet.getRange("D1");
  targetCell.setValue("Portfolio page");
  targetCell.setFontWeight("bold");
  targetCell.setBackground("#f3f3f3"); 
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("'Accelerators' not found!");
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; 
  
  const range = sheet.getRange(2, 1, lastRow - 1, 4); 
  const data = range.getValues();
  const results = [];

  for (let i = 0; i < data.length; i++) {
    let originalUrl = data[i][1]; //B(Sito)
    let existingPortfolio = data[i][3]; //D  (Portofolio)
    let foundPortfolio = "";

    if (existingPortfolio && existingPortfolio !== "") {
      writeLog("INFO", "Line " + (i + 2) + ": Portfolio is already there skip");
      continue; 
    }
    if (originalUrl && originalUrl !== "") {
      try {
        let domainUrl = cleanToDomain(originalUrl);
        writeLog("INFO", "Analyzing (" + (i + 2) + "): " + domainUrl);
        foundPortfolio = findPortfolioUrl(domainUrl);
        
        if(foundPortfolio){
          writeLog("INFO", "Found page: " + foundPortfolio); 
        }else{
          foundPortfolio = "Not found";
        }

        sheet.getRange(i + 2, 4).setValue(foundPortfolio);
      } catch (e) {
        writeLog("ERROR", "Error at line " + (i + 2) + ": " + e.toString());
        foundPortfolio = "Error : " + e.message;
      }
    }
    results.push([foundPortfolio]);
  }
}

function getHostName(url) {
  const match = url.match(/^(?:https?:\/\/)?(?:www\.)?([^\/:?#]+)/i);
  return match ? match[1] : null;
}

function cleanToDomain(url) {
  if (!url) return "";
  const regex = /^(https?:\/\/[^\/]+)/i;
  const match = url.match(regex);
  if (match && match[1]) {
    return match[1] + "/"; 
  }
  return url; 
}
function findPortfolioUrl(baseUrl) {
  try {
    const rootUrl = baseUrl.endsWith('/') ? baseUrl : baseUrl + '/';
    const options = {
      "validateHttpsCertificates": false,
      "muteHttpExceptions": true,
      "followRedirects": true, 
      "headers": {
          "Accept-Language": "en-US,en;q=0.9",
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }
    };
    const response = UrlFetchApp.fetch(rootUrl, options);
    const hostName = getHostName(rootUrl)
    const content = response.getContentText();
    const $ = Cheerio.load(content);
    const aTags = $('a')
    const keywords = ["portfolio", "startup", "start-up", "proyectos", "companies"];
    
    for(const a of aTags){
      const href = a.attribs["href"] || ""
      console.log(href)
      const isMatch = keywords.some(kw => href.includes(kw))
      if(isMatch){
        console.log(href)

        if(href.startsWith('/')) return "www." + hostName + href
        if(href.includes(hostName)) return href
      }
    }
    return null; 
  } catch (e) {
    Logger.log("Error in searching portfolio for " + baseUrl + ": " + e.toString());
    writeLog("ERROR", "Error in searching portfolio for " + baseUrl + ": " + e.toString() )
    return null;
  }
}

