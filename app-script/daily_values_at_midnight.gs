// API data
const endpointUrl = 'https://analytics.api.aiesec.org/v2/applications/analyze';
const token = '';

// Constants
const branchList = [
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

const categoryPatterns = [
  {name: "iGTa", pattern: /^i_.*_[8]$/},
  {name: "iGTe", pattern: /^i_.*_[9]$/},
  {name: "iGV", pattern: /^i_.*_[7]$/},

  {name: "oGTa", pattern: /^o_.*_[8]$/},
  {name: "oGTe", pattern: /^o_.*_[9]$/},
  {name: "oGV", pattern: /^o_.*_[7]$/}
];

// Configs
const today1 = new Date();

const offsetInMilliseconds1 = 5.5 * 60 * 60 * 1000; // 5.5 hours in milliseconds
const adjustedTime1 = new Date(today1.getTime() + offsetInMilliseconds1);

// Format the date in `YYYY-MM-DD` and time in `HH:MM:SS`
const dataStartDate = adjustedTime1.toISOString().split('T')[0];
const dataStartTime = adjustedTime1.toISOString().split('T')[1].slice(0, 8);

// console.log(`Date in GMT+5:30: ${dataStartDate}`);
// console.log(`Time in GMT+5:30: ${dataStartTime}`);

const dataEndDate = '2024-11-17';

const reportSheetName = "Daily Summary at Midnight";

const statusTypes = [
  "applied",
  "approved",
];

const tableHeaders = [
  "Branch",
  "Category",
  "Applied",
  "Approved",
];

// Helper functions
function getData(dataStartDate, dataEndDate, activity) {
  const queryUrl = `${endpointUrl}?access_token=${token}&start_date=${dataStartDate}&end_date=${dataEndDate}&performance_v3[office_id]=${1623}`;
  const responseText = UrlFetchApp.fetch(queryUrl).getContentText();
  const responseData = JSON.parse(responseText);
  return responseData;
}

// doc_count -> APL values
// applicants.value -> PPL values
function extractValues(apiData) {
  let extractedValues = {}

  categoryPatterns.forEach((patternObj) => {
    let categoryData = {}

    const matchingEntries = Object.entries(apiData).filter(([key, value]) => patternObj.pattern.test(key));

    matchingEntries.forEach((entry) => {
      statusTypes.forEach((statusType) => {
        if(entry[0].includes(statusType)){
          categoryData[statusType] = categoryData[statusType] ? categoryData[statusType] : 0 + (entry[1]?.applicants?.value || 0)
        };
      });
    })

    extractedValues[patternObj.name] = categoryData
  })

  return extractedValues;
}

function setupSheet() {
  const reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(reportSheetName);

  if (!reportSheet) {
    throw new Error($`Sheet with name ${reportSheetName} does not exist.`);
  }

  reportSheet.getRange(1, 1, 1 , tableHeaders.length).setValues([tableHeaders]); 
}

function insertRowToSheet(rowNumber, rowData){
    const reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(reportSheetName);

    reportSheet.getRange(1 + rowNumber, 1, 1 , rowData.length).setValues([rowData]); 
}

// =================

function runProcess(){
  console.log("Process started...");
  setupSheet();

  let finalData = {}
  let fetchedData = getData(dataStartDate, dataEndDate);

  console.log("Data fetching...")
  branchList.forEach((branch) => {
    let branchData = fetchedData[branch.id.toString()]
    console.log(branchData)
    const parsedData = extractValues(branchData);

    finalData[branch.name] = parsedData;
  });

  // console.log(finalData);
  console.log("Writing to sheet...");

  branchList.forEach((branch, index1) => {
    categoryPatterns.forEach((patternObj, index2) => {
      const dynamicCols = statusTypes.map((statusType) => finalData[branch.name][patternObj.name][statusType]);

      const rowData = [
        branch.name,
        patternObj.name,
        ...dynamicCols
      ];

      insertRowToSheet((index1 * categoryPatterns.length) + (index2 + 1), rowData);
    });
  });

  console.log("Sheet writing completed.");
}

