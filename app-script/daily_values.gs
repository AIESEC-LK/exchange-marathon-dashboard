// API data
const apiUrl = 'https://analytics.api.aiesec.org/v2/applications/analyze';
const authToken = '';

// Constants
const officesList = [
    { id: 222, name: 'CC' },
    { id: 872, name: 'CN' },
    { id: 1340, name: 'CS' },
    { id: 2204, name: 'Kandy' },
    { id: 4535, name: 'NIBM' },
    { id: 2186, name: 'NSBM' },
    { id: 5490, name: 'Rajarata' },
    { id: 2175, name: 'Ruhuna' },
    { id: 2188, name: 'SLIIT' },
    { id: 221, name: 'USJ' }
];

const patternList = [
  {name: "iGTa", pattern: /^i_.*_[8]$/},
  {name: "iGTe", pattern: /^i_.*_[9]$/},
  {name: "iGV", pattern: /^i_.*_[7]$/},

  {name: "oGTa", pattern: /^o_.*_[8]$/},
  {name: "oGTe", pattern: /^o_.*_[9]$/},
  {name: "oGV", pattern: /^o_.*_[7]$/}
];

// Configs
const today = new Date();
const offsetInMilliseconds = 5.5 * 60 * 60 * 1000; // 5.5 hours in milliseconds
const adjustedTime = new Date(today.getTime() + offsetInMilliseconds);

// Format the date in `YYYY-MM-DD` and time in `HH:MM:SS`
const reportStartDate = adjustedTime.toISOString().split('T')[0];
const reportStartTime = adjustedTime.toISOString().split('T')[1].slice(0, 8);

// console.log(`Date in GMT+5:30: ${reportStartDate}`);
// console.log(`Time in GMT+5:30: ${reportStartTime}`);

const reportEndDate = '2024-11-17';

const worksheetName = "Daily Numbers";

const statusKeys = [
  "applied",
  "approved",
];

const reportHeaders = [
  "Office",
  "Category",
  "Applied",
  "Approved",
];

// Helper functions
function retrieveData(reportStartDate, reportEndDate, initiative) {
  const requestUrl = `${apiUrl}?access_token=${authToken}&start_date=${reportStartDate}&end_date=${reportEndDate}&performance_v3[office_id]=${1623}`;
  const response = UrlFetchApp.fetch(requestUrl).getContentText();
  const parsedData = JSON.parse(response);
  return parsedData;
}

// doc_count -> APL values
// applicants.value -> PPL values
function parseData(apiResult) {
  let parsedValues = {}

  patternList.forEach((patternItem) => {
    let resultsObj = {}

    const matches = Object.entries(apiResult).filter(([key, value]) => patternItem.pattern.test(key));

    matches.forEach((match) => {
      statusKeys.forEach((statusKey) => {
        if(match[0].includes(statusKey)){
          resultsObj[statusKey] = resultsObj[statusKey] ? resultsObj[statusKey] : 0 + (match[1]?.applicants?.value || 0)
        };
      });
    })

    parsedValues[patternItem.name] = resultsObj
  })

  return parsedValues;
}

function initializeSheet() {
  const worksheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(worksheetName);

  if (!worksheet) {
    throw new Error($`Sheet with name ${worksheetName} does not exist.`);
  }

  worksheet.getRange(1, 1, 1 , reportHeaders.length).setValues([reportHeaders]); 
}

function addRowToSheet(rowIdx, rowInfo){
    const worksheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(worksheetName);

    worksheet.getRange(1 + rowIdx, 1, 1 , rowInfo.length).setValues([rowInfo]); 
}

// =================

function executeProcess(){
  console.log("Starting process...");
  initializeSheet();

  let outputData = {}
  let apiData = retrieveData(reportStartDate, reportEndDate);

  console.log("Fetching data...")
  officesList.forEach((office) => {
    let officeData = apiData[office.id.toString()]
    console.log(officeData)
    const parsedResult = parseData(officeData);

    outputData[office.name] = parsedResult;
  });

  // console.log(outputData);
  console.log("Writing to sheet...");

  officesList.forEach((office, idx1) => {
    patternList.forEach((patternItem, idx2) => {
      const columnValues = statusKeys.map((statusKey) => outputData[office.name][patternItem.name][statusKey]);

      const rowInfo = [
        office.name,
        patternItem.name,
        ...columnValues
      ];

      addRowToSheet((idx1 * patternList.length) + (idx2 + 1), rowInfo);
    });
  });

  console.log("Completed writing to sheet.");
}
