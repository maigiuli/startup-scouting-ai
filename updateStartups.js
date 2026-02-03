
function updateStartups() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_accelerators = ss.getSheetByName('Accelerators');
  const sheet_startups = ss.getSheetByName('startups');
  const data = sheet_accelerators.getDataRange().getValues();
  const data_startups = sheet_startups.getDataRange().getValues();
  const existingNames = data_startups.map(row => row[0].toLowerCase().trim());

  if (!data[0][4]) {
    sheet_accelerators
      .getRange("E1")
      .setValue("Processed")
      .setFontWeight("bold")
      .setBackground("#f3f3f3");
  }

   if (sheet_startups.getLastRow() === 0) {
    sheet_startups.appendRow(["Startup Name", "Website", "Location", "Accelerator"]);
    const headerRange = sheet_startups.getRange("A1:D1");
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#f3f3f3");
  }
  const MAX_BATCH = 5;
  let processedCount = 0;

  for (let i = 1; i < data.length; i++) {

    if (processedCount >= MAX_BATCH) break;

    const acceleratorName = data[i][0];
    const portfolioUrl = data[i][3];
    const processedFlag = data[i][4]; 

    if (processedFlag === "YES") continue;

    if (!portfolioUrl || portfolioUrl.includes("Not found")) {
      sheet_accelerators.getRange(i + 1, 5).setValue("YES");
      continue;
    }
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
          
          }else{
            Logger.log("Skipping already present startup " + startup.name + " per " + acceleratorName);
            writeLog("INFO", "Skipping already present startup " + startup.name + " per " + acceleratorName)
          }
        });
        Logger.log("Added " + result.startups.length + " startup per " + acceleratorName);
          writeLog("INFO", "Added " + result.startups.length + " startup per " + acceleratorName);
        
      }
       sheet_accelerators.getRange(i + 1, 5).setValue("YES");
      processedCount++;
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
  - Be careful, IMPORTANT: Some startups are only listed as logos in a horizontal carousel. 
  - They are often list near to the words "alumni", "startups", "companies"
  - DO NOT INCLUDE activities or programme's names even tough sometimes they are listed in the startup or portfolio page. 
  - If data is missing, use "".
  - Search in image "alt" tags, "h3", "h4", or JSON titles.

  WEBSITE
  - For the website, provide the FULL URL. Priority: Official external website (e.g., https://www.startup.com).
  
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
