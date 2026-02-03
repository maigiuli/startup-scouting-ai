
function enrichMissingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('startups');
  const data = sheet.getDataRange().getValues();
  
  let promptItems = [];
  const MAX_PER_BATCH = 50; 


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
            
            const websiteValue =
              enriched.website && enriched.website.trim() !== ""
                ? enriched.website
                : "unknown";

            const locationValue =
              enriched.location && enriched.location.trim() !== ""
                ? enriched.location
                : "unknown";

            sheet.getRange(j + 1, 2).setValue(websiteValue);
            sheet.getRange(j + 1, 3).setValue(locationValue);
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