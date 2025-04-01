function onOpen() {
  SpreadsheetApp.getUi().createMenu("Open dialog").addItem("Open dialog", "openDialog").addToUi();
}

function openDialog() {
  const html = HtmlService.createTemplateFromFile("index");
  html.apiKey = PropertiesService.getScriptProperties().getProperty("apiKey") || "";
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(1000).setHeight(1000), "Roadmap Generator as Gemini");
}

function doGemini(obj) {
  try {
    if (obj.apiKey && obj.apiKey != "") {
      PropertiesService.getScriptProperties().setProperty("apiKey", obj.apiKey);
    } else {
      throw new Error("Please set your API key and try again.");
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const object = {
      ...obj,
      spreadsheetId: ss.getId(),
      sheetId: sheet.getSheetId(),
    };
    const g = new GenerateRoadmap(object).generateContent();
    g.putValuesToSpreadsheet();
    return "Done.";
  } catch (e) {
    console.error(e.stack)
    return `Error: ${e.message}`;
  }
}
