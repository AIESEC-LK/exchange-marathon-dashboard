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

// Function to create a daily trigger
function createDailyTrigger2() {
  // Delete existing triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "copyRangeToSheet") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger("copyRangeToSheet")
    .timeBased()
    .atHour(20)
    .nearMinute(35) 
    .everyDays(1)
    .create();
}

