
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

function scoutingAccelerators() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Accelerators');

  const url = "https://rankings.ft.com/incubator-accelerator-programmes-europe"; 
  
  try {
    let lastRow = sheet.getLastRow();
    let startIndex = (lastRow <= 1) ? 1 : lastRow+1; 

    console.log("Sheet has " + lastRow + " lines. I start to read the HTML from the position: " + startIndex);
    writeLog("INFO", "Sheet has " + lastRow + " lines. I start to read the HTML from the position: " + startIndex);
    if (lastRow==151) writeLog("INFO", "No more accelerators to be added")
    const response = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
    const html = response.getContentText();
    
    const rows = html.split('<tr');
    let foundCount = 0;
    const MAX_BATCH = 10; 
    let results = [];

    for (let i = 1; i < rows.length && foundCount < MAX_BATCH; i++) {
      let row = rows[i];
      
      const cells = row.split(/<td[^>]*>/i);
      
      if (cells.length > 6) {
        if (i < startIndex) continue;
        
        const linkMatch = /<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/i.exec(cells[2]);
        
        
        let name = linkMatch ? linkMatch[2].replace(/<[^>]*>/g, "").trim() : "";
        let website = linkMatch ? linkMatch[1].trim() : "";
        let city = cells[5].replace(/<[^>]*>/g, "").trim();
        let country = cells[6].replace(/<[^>]*>/g, "").trim();
        
        let location = (city && country) ? `${city}, ${country}` : (city || country || "N/A");

        if (name && website) {
          results.push([name, website, location]);
          foundCount++; 
        }
      }
      
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Name", "Website", "Location"]);
      sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#f3f3f3");
    }

    if (results.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, results.length, 3).setValues(results);
      console.log("Batch added. ");
      writeLog("INFO", "Batch added. "); 
    }

  } catch (e) {
    console.error("Error in scouting: " + e.toString());
  }
}
    

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
    let existingPortfolio = data[i][3]; //D (Portfolio)
    let foundPortfolio = "";

    if (existingPortfolio && existingPortfolio !== "") {
      Logger.log("Line " + (i + 2) + ": Portfolio is already there skip");
      writeLog("INFO", "Line " + (i + 2) + ": Portfolio is already there skip");
      continue; 
    }
    if (originalUrl && originalUrl !== "") {
      try {
        let domainUrl = cleanToDomain(originalUrl);
        
        Logger.log("Analyzing (" + (i + 2) + "): " + domainUrl);
        writeLog("INFO", "Analyzing (" + (i + 2) + "): " + domainUrl);
        
        
        foundPortfolio = findPortfolioUrl(domainUrl);
        
        if (!foundPortfolio) {
          foundPortfolio = "Not found";
        }
        sheet.getRange(i + 2, 4).setValue(foundPortfolio);
      } catch (e) {
        Logger.log("Error at line " + (i + 2) + ": " + e.toString());
        writeLog("ERROR", "Error at line " + (i + 2) + ": " + e.toString());
        foundPortfolio = "Error : " + e.message;
      }
    }
    
    
    results.push([foundPortfolio]);
  }


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
    const html = response.getContentText();
    const htmlLower = html.toLowerCase();
    const linkRegex = /<a\s+(?:[^>]*?\s+)?href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
    
    
    let match;
    const keywords = ["portfolio", "startup", "start-up", "companies"];
    
    while ((match = linkRegex.exec(html)) !== null) {
      let href = match[1];
      let linkText = match[2].replace(/<[^>]*>/g, '').toLowerCase().trim();
      if (href === "#" || href.trim() === "") continue;
      const isMatch = keywords.some(kw => linkText.includes(kw) || href.includes(kw));

      if (isMatch) {
        if (href.startsWith('/')) {
          href = rootUrl.replace(/\/$/, '') + match[1]; 
        } else if (!href.startsWith('http')) {
          href = rootUrl + match[1];
        } else {
          href = match[1];
        }
        
        if (href.includes(baseUrl) && !href.includes('linkedin') && !href.includes('twitter')) {
          Logger.log("Founded page: " + href + " (from text: " + linkText + ")");
          writeLog("INFO", "Founded page: " + href + " (from text: " + linkText + ")"); 
          return href; 
        
      }
    }
    }
   
    return null; 
  } catch (e) {
    Logger.log("Error in searching portfolio for " + baseUrl + ": " + e.toString());
    writeLog("ERROR", "Error in searching portfolio for " + baseUrl + ": " + e.toString() )
    return null;
  }
}


function updateStartups() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Accelerators');
  const sheet_startups = ss.getSheetByName('startups');
  const data = sheet.getDataRange().getValues();
  const data_startups = sheet_startups.getDataRange().getValues();
  const existingNames = data_startups.map(row => row[0].toLowerCase().trim());

   if (sheet_startups.getLastRow() === 0) {
    sheet_startups.appendRow(["Startup Name", "Website", "Location", "Accelerator"]);
    const headerRange = sheet_startups.getRange("A1:D1");
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#f3f3f3");
  }

  for (let i = 1; i < data.length; i++) {
    let acceleratorName = data[i][0];
    let portfolioUrl = data[i][3];
    
    if (!portfolioUrl || portfolioUrl.includes("Website not reachable")|| portfolioUrl.includes("Not found")) continue;

    try {
      let html = UrlFetchApp.fetch(portfolioUrl, {"muteHttpExceptions": true}).getContentText();
      html = html.replace(/<(script|style|svg|footer|nav)[\s\S]*?<\/\1>/gi, "");
      html = html.replace(/<svg[\s\S]*?<\/svg>/gi, "")
      html = html.replace(/class="[^"]*"/gi, "");
      
      let responseAI = askAIForStartups(html.substring(0, 100000), acceleratorName);
      
      let result = JSON.parse(responseAI);

      
      if (result.startups && result.startups.length > 0) {
        result.startups.forEach(startup => {
          if (!existingNames.includes(startup.name.toLowerCase().trim())) {
          sheet_startups.appendRow([
            startup.name, 
            startup.website, 
            startup.location, 
            acceleratorName
          ]);
          existingNames.push(startup.name.toLowerCase().trim());
          Logger.log("Added " + result.startups.length + " startup per " + acceleratorName);
          writeLog("INFO", "Added " + result.startups.length + " startup per " + acceleratorName);
          }else{
            Logger.log("Skipping already present startup " + startup.name + " per " + acceleratorName);
            writeLog("INFO", "Skipping already present startup " + startup.name + " per " + acceleratorName)
          }
        });
        
      }
    } catch (e) {
      Logger.log("Error for " + acceleratorName + ": " + e.message);
      writeLog("ERROR", "Error for " + acceleratorName + ": " + e.message)
    }
    Utilities.sleep(1000);
  }
}

function askAIForStartups(contentToAnalyze, acceleratorName) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  const apiUrl = "https://api.openai.com/v1/chat/completions";

  const prompt = `Analyze the portfolio text of "${acceleratorName}". 
  Extract EVERY startup. For each, find: Name, Official Website URL, and Location.
  
  STRICT RULES:
  - Respond ONLY with a JSON object: {"startups": [{"name": "...", "website": "...", "location": "..."}]}.
  - For the website, provide the FULL URL 
    1. Priority: Official external website (e.g., https://www.startup.com).
    2. Alternative: If the external site is not visible, provide the internal relative link found in the source but always providing the FULL path (e.g., https://www.accelerators.com/our-companies/name)
  - If data is missing, use "".
  - Search in image "alt" tags, "h3", "h4", or JSON titles.
  
  Text: ${contentToAnalyze}`;

  const payload = {
    "model": "gpt-4o-mini",
    "messages": [
      {"role": "system", "content": "You are a data extraction tool. Output only valid JSON."},
      {"role": "user", "content": prompt}
    ],
    "response_format": { "type": "json_object" }, 
    "temperature": 0
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {"Authorization": "Bearer " + apiKey},
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(apiUrl, options);
  return JSON.parse(response.getContentText()).choices[0].message.content;
}


function generateValueProposition() {


  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Startups');
  const apiKey = PropertiesService
  .getScriptProperties()
  .getProperty('OPENAI_API_KEY');
  let targetCell = sheet.getRange("E1");
  if (targetCell.getValue() === "") {
  targetCell.setValue("Value Proposition");
  targetCell.setFontWeight("bold");
  targetCell.setBackground("#f3f3f3"); 
  }

  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) { 
    const vp = data[i][4]; 
    if (!vp) {
      const nome = data[i][0];
      const website = data[i][1];
      const location = data[i][2];
      const acc = data[i][3];

      const prompt = `Generate a value proposition clear for the startup ${nome} with website ${website} and location ${location} that used the accelerator ${acc} using this scheme: Startup <X> helps <Target Y> do <What W> so that <Benefit Z>.`;

      const response = UrlFetchApp.fetch(
        "https://api.openai.com/v1/chat/completions",
        {
          method: "post",
          headers: {
            "Authorization": "Bearer " + apiKey,
            "Content-Type": "application/json"
          },
          payload: JSON.stringify({
            model: "gpt-4o-mini",
            messages: [
              { role: "system", content: "You are a startup analyst and venture capital." },
              { role: "user", content: prompt }
            ],
            temperature: 0.4
          })
        }
      );

      const json = JSON.parse(response.getContentText());
      const vpTesto = json.choices[0].message.content;

      sheet.getRange(i+1, 5).setValue(vpTesto); 
      
    }else{
      console.log("Skip: vp already there.")
      writeLog("INFO", "Skip: vp already there.")
    }
  }
}


function enrichMissingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('startups');
  const data = sheet.getDataRange().getValues();
  
  let promptItems = [];
  const MAX_PER_BATCH = 10; 

  for (let i = 1; i < data.length; i++) {
  
    if (promptItems.length >= MAX_PER_BATCH) break;

    let name = data[i][0];
    let website = data[i][1];
    let location = data[i][2];
    let acc = data[i][3];

    const isWebsiteMissing = !website || website.toString().trim() === "" || website.includes(acc.split(" ").join("").toLowerCase()) || website.includes(acc.split(" ").join(".").toLowerCase());
    
    const isLocationMissing = !location || location.toString().trim() === "" || location.toString().toLowerCase() === "n/a";

    if (isWebsiteMissing || isLocationMissing) {
      promptItems.push({ 
        name: name, 
        currentWebsite: website || "unknown", 
        currentLocation: location || "unknown" 
      });
      writeLog("INFO", "Starting enrichment for " + promptItems.length + " startups.");
    }
  }

  if (promptItems.length === 0) {
    Logger.log("All data is complete. ");
    writeLog("INFO", "All data is complete. ");
    return;
  }

  const prompt = `You are a startup database expert. Complete the missing information (official website and headquarters city/country) for the following startups.
  Return ONLY a JSON object with this structure: 
  {"enrichedStartups": [{"name": "...", "website": "...", "location": "..."}]}
  
  Startups List: ${JSON.stringify(promptItems)}`;

  try {
    
    const responseAI = askAI(prompt);
    if (!responseAI) throw new Error("No response received from AI.");

    const cleanJson = responseAI.replace(/```json/g, "").replace(/```/g, "").trim();
    const result = JSON.parse(cleanJson);

    if (result.enrichedStartups && result.enrichedStartups.length > 0) {
      
      result.enrichedStartups.forEach(enriched => {
        for (let j = 1; j < data.length; j++) {
          
          if (data[j][0].toLowerCase().trim() === enriched.name.toLowerCase().trim()) {
            
            if (enriched.website && enriched.website !== "unknown") {
              sheet.getRange(j + 1, 2).setValue(enriched.website);
              
            }
            if (enriched.location && enriched.location !== "unknown") {
              sheet.getRange(j + 1, 3).setValue(enriched.location);
            }
            writeLog("SUCCESS", "Successfully enriched: " + enriched.name);
            Logger.log("Successfully enriched: " + enriched.name);
            break; 
          }
        }
      });
    }
  } catch (e) {
    writeLog("ERROR", "Critical failure: " + e.message);
    Logger.log("Critical Error during enrichment: " + e.message);
  }
}

function askAI(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  const url = "https://api.openai.com/v1/chat/completions";

  const payload = {
    "model": "gpt-4o-mini",
    "messages": [
      {
        "role": "system", 
        "content": "You are a data assistant. Return only raw JSON. No conversational text."
      },
      {
        "role": "user", 
        "content": prompt
      }
    ],
    "temperature": 0.1 
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "Authorization": "Bearer " + apiKey
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.choices && json.choices.length > 0) {
      return json.choices[0].message.content;
    } else {
      Logger.log("AI API Error: " + response.getContentText());
      writeLog("ERROR", "AI API Error: " + response.getContentText()); 
      return null;
    }
  } catch (e) {
    Logger.log("Network Error: " + e.message);
    writeLog("ERROR", "Network Error: " + re.message); 

    return null;
  }
}