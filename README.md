# Startup Scouting AI – Prototype

This repository contains a prototype that automates the scouting of 10 Top accelerators and startups and generates a synthetic value proposition for startups using Google Sheets, Google Apps Script, and LLM via API.

---

## Google Sheet

[Link to google sheet](https://docs.google.com/spreadsheets/d/1_qYL7gVXxr6ltYNpBeAFCOSWWLntlFvwEFX2AzuPDaE/edit?gid=2060480799#gid=2060480799)

The spreadsheet contains two main sheets:

- Accelerators, with columns:
  - `Name` 
  - `Website`
  - `Location`
  - `Portfolio Page`
  - `Processed`

- Startups, with columns:
  - `Name` 
  - `Website`
  - `Location`
  - `Accelerator` 
  - `Value Proposition`


When the sheet is opened, a menu called “Startup Scouting AI” appears in the top navigation bar. From this menu, the entire workflow can be executed without interacting with the code.

## Startup scouting AI Menu 

- **scoutingAccelerators**: Adds new accelerators to the Accelerators sheet, using as a source this [link](https://rankings.ft.com/incubator-accelerator-programmes-europe) (Financial Times Europe's Leading Start-Up Hubs Ranking). The function is made so to add accelerators in batch of 10 items and never add the same item multiple times. The script saves name, website and location of each accelerator and adds an heading line if it is the first execution of the function. I used the Cheerio library to deal with URL parsing and normalization. 'UrlFetchApp' calls are always handled with a try-catch scheme to avoid interruptions of the code. 
- **updateAllPortfolioUrls**: Find portfolio pages for each accelerator present in the Accelerators sheet and save the link in the fourth column of the same sheet. The function attempts to identify portfolio / alumni / companies pages using simple URL heuristics. An heading line is added if it is the first execution of the function. The function calls two other functions, **cleanToDomain**, that eventually cleans the link from potentially dangerous extra pages, and **findPortfolioUrl**, the one that actually does the job, using as keywords "portfolio", "startup", "start-up" and "companies" . 
- **updateStartups**: Update startups related to each accelerator in the list and report them in the "startups" page (name, website, location, accelerator). The function calls an auxiliary function **askAIForStartups** that uses a LLM prompt in order to extract the startups from the portfolio page. The LLM uses a `gpt-4o-mini` model, a constraint on the response format, and a temperature = 0 to avoid hallucinations.
- **enrichMissingData**: Find missing information for the listed startups. Many website do not contain the link to the startup page or their location. The function provides missing information in order to have a more complete dataset to be given to the next function.  
- **generateValueProposition**: Generate missing value proposition. The function visits the startup website and generates a concise sentence using an LLM prompt. Here the temperature is slightly higher (0.4). 

In all the functions, apiKey is always managed with `PropertiesService` for security. 
  
---

## Setup with API Key (LLM)

The LLM API key is stored securely using Google Apps Script `PropertiesService` and is **not committed to the repository**.

Set the API key once by running the following snippet in the Apps Script editor:

```js
PropertiesService.getScriptProperties()
  .setProperty("OPENAI_API_KEY", "your-api-key");
