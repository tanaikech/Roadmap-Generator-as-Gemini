/**
 * Class object for generating roadmap using Gemini API.
 * Author: Kanshi Tanaike
 * version 1.0.0
 * @class
 */
class GenerateRoadmap {
  /**
   * @param {Object} object Object using this script.
   * @param {String} object.apiKey API key for Gemini API.
   * @param {String} object.goal Goal of the roadmap.
   * @param {String|Object} object.description Description of the roadmap.
   * @param {String} object.spreadsheetId Spreadsheet ID.
   * @param {String} object.sheetId Sheet ID.
   * @param {String} object.dateFormat Date format to use in JSON schema.
   */
  constructor(object) {
    const { apiKey, goal, description, spreadsheetId, sheetId, dateFormat = "yyyy/MM/dd HH:mm:ss" } = object;

    /** @private */
    this.apiKey = apiKey;
    /** @private */
    this.spreadsheetId = spreadsheetId;
    /** @private */
    this.sheetId = sheetId;
    /** @private */
    this.keys = ["task", "startDateTime", "endDateTime", "duration", "dependencies", "note"];
    /** @private */
    this.keyLen = this.keys.length;
    /** @private */
    this.delayTime = 3600;
    /** @private */
    this.generatedContent;
    const properties = {
      title: { description: "The goal of this roadmap.", type: "string" },
      startOfTask: { description: `Start of this roadmap. Format is '${dateFormat}'.`, type: "string" },
      endOfTask: { description: `End of this roadmap. Format is '${dateFormat}'.`, type: "string" },
      times: { description: `Times for the time axis of the first row. Format is '${dateFormat}'. First, confirm 'startDateTime' and 'endDateTime' for each task. And then, these values are required to be given to include 'startDateTime' and 'endDateTime' of all tasks in the roadmap.`, type: "array", items: { type: "string" } },
      tasks: {
        type: "array",
        items: {
          type: "object",
          properties: {
            task: { description: "Task name", type: "string" },
            startDateTime: { description: `Start date time of this task. Format is '${dateFormat}'.`, type: "string" },
            endDateTime: { description: `End date time of this task. Format is '${dateFormat}'.`, type: "string" },
            duration: { description: "Duration of this task. This unit is second.", type: "number" },
            dependencies: { description: "Dependencies for other tasks.", type: "array", items: { type: "string" } },
            note: { description: "Description of this task.", type: "string" },
          },
          required: ["task", "startDateTime", "endDateTime", "duration", "dependencies", "note"],
          additionalProperties: false,
        }
      }
    };
    if (description) {
      if (typeof description == "object") {
        /** @private */
        this.jsonSchema = description;
      } else if (typeof description == "string") {
        const jsonSchema = { description, type: "object", properties };
        /** @private */
        this.prompt = [
          "The following text is a sample roadmap. Understand it carefully. Resolve the following text to task, startDateTime, endDateTime, dependencies, and note. Return the result by following the JSON schema.",
          JSON.stringify(jsonSchema),
        ].join("\n")
      } else {
        throw new Error("Please set valid description.");
      }
    } else if (goal) {
      const now = new Date();
      const start = new Date(now.getTime() + (this.delayTime * 1000));
      /** @private */
      this.description = [
        `The goal of this roadmap is ${goal}.`,
        `Create a JSON object to confirm ${goal} as a roadmap of the Gantt chart.`,
        `The start of this roadmap is ${Utilities.formatDate(start, Session.getScriptTimeZone(), dateFormat)}. Set this as the start date time of this roadmap.`,
      ].join("\n");
      /** @private */
      this.jsonSchema = { description: this.description, type: "object", properties };
    } else {
      throw new Error("Please set valid goal or description.");
    }
  }

  /**
  * ### Description
  * Generate content using Gemini API.
  *
  * @return {GenerateRoadmap}
  */
  generateContent() {
    const g = new GeminiWithFiles.geminiWithFiles({
      apiKey: this.apiKey,
      model: "models/gemini-2.0-flash-exp",
      generationConfig: { responseMimeType: "application/json" },
      tools: [{ googleSearch: {} }],
    });
    let res;
    if (this.jsonSchema) {
      res = g.generateContent({ jsonSchema: this.jsonSchema });
    } else {
      res = g.generateContent({ q: this.prompt });
    }
    if (!res.times || !Array.isArray(res.times) || res.times.length == 0) {
      throw new Error("Gemini couldn't generate a valid roadmap. Try again.");
    }
    const timesAr = Array(res.times.length).fill(null);
    const values = [
      [res.title, ...Array(this.keyLen - 1).fill(null), ...timesAr],
      [...this.keys, ...res.times],
      ...res.tasks.map(o => [...this.keys.map(k => Array.isArray(o[k]) ? o[k].join(",") : o[k]), ...timesAr])
    ];
    this.generatedContent = { res, values };
    return this;
  }

  /**
  * ### Description
  * Create gantt chart on Google Sheets.
  * This method can be used for the format generated by only the original JSON schema. Please be careful about this.
  *
  * @return {void}
  */
  putValuesToSpreadsheet() {
    const { res, values } = this.generatedContent;
    const sheet = SpreadsheetApp.openById(this.spreadsheetId).getSheetById(this.sheetId);
    const rule = SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied("=AND(G$2>=$B3,G$2<=$C3)")
      .setBackground("#ff0000")
      .setRanges([sheet.getRange(3, this.keyLen + 1, values.length - 1, res.times.length)])
      .build();
    sheet.clear().getRange(1, 1, values.length, values[0].length).setValues(values);
    sheet.setConditionalFormatRules([rule]);
    sheet.autoResizeColumns(1, sheet.getMaxColumns());
  }

  /**
  * ### Description
  * Get generated content.
  *
  * @type {Object} Gererated content.
  */
  get getGeneratedContent() {
    return this.generatedContent.res;
  }
}
