function onOpen() {
  createTrigger();
}

function createTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == 'masterFunction') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('masterFunction').timeBased().everyMinutes(5).create();
}

function masterFunction() {
  fetchData();
}

function getCurrentDate() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function fetchData() {
  var today = getCurrentDate();
  var url1 = `https://analytics.api.aiesec.org/v2/applications/analyze.json?access_token=ArA7-7UZV6QGhqnouAfwZSA49g4pxNxmRxt0giHjBxA&start_date=$2025-03-11&end_date=${today}&performance_v3%5Boffice_id%5D=1623`;

  var response = UrlFetchApp.fetch(url1);
  var jsonData = JSON.parse(response.getContentText());
  processExchangeData(jsonData);
}

function processExchangeData(data) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh0 = ss.getSheetByName("Total_Raw_Exchange_Data");
  var sheetData = [];

  for (var key in data) {
    if (data.hasOwnProperty(key) && ['2204', '872', '1340', '222', '221', '2175', '4535', '2186', '2188', '5490'].includes(key)) {
      var section = data[key];

      for (var subKey in section) {
        if (section.hasOwnProperty(subKey)) {
          var item = section[subKey];
          var concatenatedKey = key + "_" + subKey; // Concatenate key with subKey
          sheetData.push([concatenatedKey, '']);
          sheetData.push([item.doc_count, item.doc_count]);
        }
      }
    }
  }
  if (sheetData.length > 0) {
    let targetRange = sh0.getRange(2, 1, sheetData.length, 2);
    targetRange.setValues(sheetData);
  }
}
