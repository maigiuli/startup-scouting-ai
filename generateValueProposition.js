
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
  const MAX_BATCH = 40;
  let processedCount = 0;

  let startRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (!data[i][4] || data[i][4].toString().trim() === "") {
      startRow = i;
      break;
    }
  }

  if (startRow === -1) {
    Logger.log("All value propositions are filled.");
    writeLog("INFO", "All value propositions are filled.");
    return;
  }
  
  for (let i = startRow; i < data.length; i++) { 
    if (processedCount >= MAX_BATCH) break;
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
      processedCount++;
      
    }else{
      console.log("Skip: vp already there.")
      writeLog("INFO", "Skip: vp already there.")
    }
  }
}

