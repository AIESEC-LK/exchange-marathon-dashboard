function createDailyTrigger() {
  // Delete existing triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "masterSaveFunction") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger("masterSaveFunction")
    .timeBased()
    .atHour(19)
    .nearMinute(55) 
    .everyDays(1)
    .create();
}

function masterSaveFunction() {
  saveExchangeData();
  saveDailyScores();
  copyRangeToSheet();
}

function getNextAvailableColumn(sheet) {
  var lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return 2; // Default to column 2 if nothing exists
  
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  for (var col = lastColumn - 1; col >= 0; col--) {
    if (headers[col] !== "" && headers[col] !== null) {
      return col + 2; // Move to the next empty column
    }
  }

  return 2; // If no date headers exist, start from column 2
}

function saveExchangeData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var marathonSheet = ss.getSheetByName("Marathon_Exchange_Data");
  var historySheet1 = ss.getSheetByName("Cumulative_Exchange_Data");

  var entities = marathonSheet.getRange("C43:C52").getValues();
  var apd = marathonSheet.getRange("D43:D52").getValues();
  var apl = marathonSheet.getRange("E43:E52").getValues();
  var su = marathonSheet.getRange("F43:F52").getValues();
  var date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  // Check if today's date column already exists
  var headers = historySheet1.getRange(1, 1, 1, historySheet1.getLastColumn()).getValues()[0];
  var dateExists = headers.some(header => {
    if (header instanceof Date) {
      var headerDate = Utilities.formatDate(header, Session.getScriptTimeZone(), "yyyy-MM-dd");
      return headerDate === date;
    } else if (typeof header === "string") {
      return header.trim() === date.trim();
    }
    return false; // Skip non-date and non-string headers
  });

  if (dateExists) {
    Logger.log("Date already exists. Skipping update.");
    return;
  }

  var newColumn = getNextAvailableColumn(historySheet1);
  historySheet1.getRange(1, newColumn).setValue(date); // Set date as the header

  for (var i = 0; i < entities.length; i++) {
    var entityName = entities[i][0];

    if (entityName) {
      var apl_row = historySheet1.createTextFinder("APL").findNext(); // Find APL row
      var apd_row = historySheet1.createTextFinder("APD").findNext(); // Find APD row
      var su_row = historySheet1.createTextFinder("SU").findNext(); // Find row SU row
      if (apl_row) {
        historySheet1.getRange(apl_row.getRow()+i, newColumn).setValue(apl[i][0]); // Update existing row
        historySheet1.getRange(apd_row.getRow()+i, newColumn).setValue(apd[i][0]); // Update existing row
        historySheet1.getRange(su_row.getRow()+i, newColumn).setValue(su[i][0]); // Update existing row
      } else {
        var lastRow = historySheet.getLastRow() + 1;
        historySheet1.getRange(lastRow, 1).setValue(entityName); 
        historySheet1.getRange(lastRow, newColumn).setValue(apd[i][0]); // Add score
      }
    }
  }
}

function saveDailyScores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dailySheet = ss.getSheetByName("Daily_Score");
  var historySheet2 = ss.getSheetByName("Daily_Score_History");

  var entities = dailySheet.getRange("B17:B26").getValues();
  var scores = dailySheet.getRange("C17:C26").getValues();
  var date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  // Check if today's date column already exists
  var headers = historySheet2.getRange(1, 1, 1, historySheet2.getLastColumn()).getValues()[0];
  var dateExists = headers.some(header => {
    if (header instanceof Date) {
      var headerDate = Utilities.formatDate(header, Session.getScriptTimeZone(), "yyyy-MM-dd");
      return headerDate === date;
    } else if (typeof header === "string") {
      return header.trim() === date.trim();
    }
    return false; // Skip non-date and non-string headers
  });

  if (dateExists) {
    Logger.log("Date already exists. Skipping update.");
    return;
  }

  var newColumn = getNextAvailableColumn(historySheet2);
  historySheet2.getRange(1, newColumn).setValue(date); // Set date as the header

  for (var i = 0; i < entities.length; i++) {
    var entityName = entities[i][0];

    if (entityName) {
      var row = historySheet2.createTextFinder(entityName).findNext(); // Find row by entity name
      if (row) {
        historySheet2.getRange(row.getRow(), newColumn).setValue(scores[i][0]); // Update existing row
      } else {
        var lastRow = historySheet2.getLastRow();
        historySheet2.getRange(lastRow, 1).setValue(entityName); 
        historySheet2.getRange(lastRow, newColumn).setValue(scores[i][0]); // Add score
      }
    }
  }
}

function copyRangeToSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var SOURCE_SHEET = "Streamlit"; 
  var DEST_SHEET = "Streamlit2";
  var RANGE_TO_COPY = "A1:F61";

  var sourceSheet = ss.getSheetByName(SOURCE_SHEET);
  var destSheet = ss.getSheetByName(DEST_SHEET);

  // Get data from source range
  var data = sourceSheet.getRange(RANGE_TO_COPY).getValues();

  // Clear old data in destination sheet before copying
  destSheet.getRange(RANGE_TO_COPY).clearContent();

  // Paste new data
  destSheet.getRange(RANGE_TO_COPY).setValues(data);

  Logger.log("Data copied successfully!");
}
