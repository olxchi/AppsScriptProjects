// We have a folder alerts 
// https://drive.google.com/drive/folders/0000000000000000000000000000000
const ALERTS_FOLDER = DriveApp.getFolderById('0000000000000000000000000000000');

// This folder has spreadsheets that hold all the KPI alerts for different sites
// Each document has a title of the format 'sitename kpi_triggered'
// the date we need is the date the spreadsheet was created - 1
// if more than 1 kpi is triggered for the same site at the same time, we need to append with a comma to the cell that already exists

// for each file in Alerts folder
// read the title and seperate sitename and the kpi that's triggered
// find the date the file was created and subtract 1 from it
// check if that date already exists in row 1, check if that site already exists in column B
// if date doesn't exist, add a new date to the leftmost column closest to our sites (column C - we'd insert a new third column)
// if site doesn't exist, append new site to bottom of column B

// Get the locations of the given site/date cell we need, at that site/date cell insert kpi_triggered name 

// make the original alert spreadsheet file read only

// We can have a dictionary of Site names that populate while we fill the spreadsheet
// {"EDTNAB02-IOTSMF-01" : cell , }

// get the active spreadsheet 
// var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1"); 

// get the cell object 
// var cell = sheet.getRange("C2");

// given a cell object, get the value inside it 
// var value = cell.getValue();

// set the value of a cell
// cell.setValue("Hello World")
var cellMap = {} // site --> date --> cell ex. { "EDTNAB02-IOTSMF-01": { "2025-04-21": "C3", "2025-04-22": "D3" } } 

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Get Alerts Info')
  .addItem('Fill Empty Sheets With Alerts', 'fillSheet')
  .addToUi();
}

function parseFileName(filename) {
  // function splits filename into site and kpi 
  // const filename = "EDTNAB02-IOTSMF-01 max_ul_throughput.csv";

  // remove the .csv extension
  let nameNoExt = filename.replace(/\.csv$/i, "");

  // split into two parts 
  const split_title = nameNoExt.split(/\s+/);

  return split_title // site = split_title[0], kpi = split_title[1]

}


function updateSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Test Sheet");
  
  // Get the latest file from the alerts folder
  const files = ALERTS_FOLDER.getFiles();
  const file = files.next();
  
  // Get file details
  const fileName = file.getName();
  const split_title = parseFileName(fileName);
  const siteName = split_title[0];
  const kpiName = split_title[1];
  
  // Calculate the date we need (one day before creation date)
  var creationDate = file.getDateCreated();
  var dayBefore = new Date(creationDate.getTime() - 24 * 60 * 60 * 1000);
  var dateToCheck = Utilities.formatDate(dayBefore, 'EST', 'yyyy-MM-dd');
  
  // Get all current data
  var data = sheet.getDataRange().getValues();
  var headerRow = data[0];
  
  // Find column index of the date (starting from index 2 which is column C)
  var dateColIndex = -1;
  
  // Check each column header starting from column C (index 2)
  for(var i = 2; i < headerRow.length; i++) {
    if(headerRow[i] instanceof Date) {
      var headerDate = Utilities.formatDate(headerRow[i], 'EST', 'yyyy-MM-dd');
      //Logger.log("Comparing " + dateToCheck + " with header date " + headerDate);
      if(headerDate === dateToCheck) {
        dateColIndex = i;
        break;
      }
    }
  }
  
  // Find row index of the site
  var siteRowIndex = -1;
  for(var i = 1; i < data.length; i++) {
    if(data[i][1] === siteName) {
      siteRowIndex = i;
      break;
    }
  }

  //Logger.log("Site: " + siteName);
  //Logger.log("Date to check: " + dateToCheck);
  //Logger.log("Date column index: " + dateColIndex);
  //Logger.log("Site row index: " + siteRowIndex);

  if(siteRowIndex !== -1) {
    if(dateColIndex !== -1) {
      // Both site and date exist - update existing cell
      var cell = sheet.getRange(siteRowIndex + 1, dateColIndex + 1);
      var currentValue = cell.getValue();
      var newValue = currentValue ? currentValue + ", " + kpiName : kpiName;
      cell.setValue(newValue);
      
      // Update cellMap
      if(!cellMap[siteName]) {
        cellMap[siteName] = {};
      }
      cellMap[siteName][dateToCheck] = cell.getA1Notation();
    } else {
      // Site exists but date doesn't - create new column
      sheet.insertColumnAfter(2);
      sheet.getRange(1, 3).setValue(dateToCheck);
      var cell = sheet.getRange(siteRowIndex + 1, 3);
      cell.setValue(kpiName);
      
      // Update cellMap
      if(!cellMap[siteName]) {
        cellMap[siteName] = {};
      }
      cellMap[siteName][dateToCheck] = cell.getA1Notation();
    }
  } else {
    // Site doesn't exist - add new row
    var newRowIndex = sheet.getLastRow() + 1;
    sheet.getRange(newRowIndex, 1).setValue(newRowIndex - 1);
    sheet.getRange(newRowIndex, 2).setValue(siteName);
    
    if(dateColIndex !== -1) {
      // Date exists - add to existing column
      var cell = sheet.getRange(newRowIndex, dateColIndex + 1);
      cell.setValue(kpiName);
      
      // Update cellMap
      if(!cellMap[siteName]) {
        cellMap[siteName] = {};
      }
      cellMap[siteName][dateToCheck] = cell.getA1Notation();
    } else {
      // Date doesn't exist - create new column
      sheet.insertColumnAfter(2);
      sheet.getRange(1, 3).setValue(dateToCheck);
      var cell = sheet.getRange(newRowIndex, 3);
      cell.setValue(kpiName);
      
      // Update cellMap
      if(!cellMap[siteName]) {
        cellMap[siteName] = {};
      }
      cellMap[siteName][dateToCheck] = cell.getA1Notation();
    }
  }

  // Maintain formatting
  sheet.autoResizeColumns(1, sheet.getLastColumn());
 //// sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold");
  sheet.setFrozenColumns(2);
  sheet.getRange(1, 1, sheet.getLastRow(), 2).setFontWeight("bold");
}
//testing every 5 min trigger
// function createTrigger() {
//   ScriptApp.newTrigger('updateSheet')
//     .timeBased()
//     .everyMinutes(5)
//     .create();
// }
//trigger that runs updatesheet every day at 8am est
function createTimeTrigger() {
  ScriptApp.newTrigger('updateSheet')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone("America/New_York")
    .create();
}

function fillSheet() {
  // get our alerts folder
 // const ALERTS_FOLDER = DriveApp.getFolderById('1stSrrI1PL9nReAU4wJ8c402BQnYOGaIQ');
 // open spreadsheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Test Sheet"); // change this 

  sheet.clearContents();
  sheet.getRange(1,1).setValue("Index");
  sheet.getRange(1,2).setValue("Site");
  
  // const dateRow = 1;
  // var numbersCol = 1; // column A
  // const siteColumn = 2; // Column B

  // var siteRowMap = {}; // key: site name, value: row number site is found
  var dateColNum = {} // key: date, value: column number date is found
  //made function global
  //var cellMap = {} // site --> date --> cell ex. { "EDTNAB02-IOTSMF-01": { "2025-04-21": "C3", "2025-04-22": "D3" } } 

  var siteDateKPIs = {} // site --> date --> [kpis]
  

  // get a list of all files in our alerts folder
  const files = ALERTS_FOLDER.getFiles();
  let count = 0;

  // for every file in our folder, get the name of file, the date it was created and fill our big dictionary
  while (files.hasNext()) {
    const file = files.next();
    const file_name = file.getName();

    // split the file name into kpi name and site name
    split_title = parseFileName(file_name);
    const SiteName = split_title[0];
    const kpiName = split_title[1];

    // get the date of the alert
    var creationDate = file.getDateCreated();
    var dayBefore = new Date(creationDate.getTime() - 24 * 60 * 60 * 1000);
    var adjDate = Utilities.formatDate(dayBefore, 'EST', 'yyyy-MM-dd');

    // fill our dictionaries

    // if site date not in dict, add it with smaller dict for date and kpi
    if(!siteDateKPIs[SiteName]) {
      siteDateKPIs[SiteName] = {};
    }

    if(!siteDateKPIs[SiteName][adjDate]) {
      siteDateKPIs[SiteName][adjDate] = [];
      siteDateKPIs[SiteName][adjDate].push(kpiName);
    }
    
 //   console.log(siteDateKPIs);
    count += 1;
 //   if (count > 10) { break }
  }
  
  // create an array of all the dates in our folder
  var dateDicts = Object.values(siteDateKPIs);
  var dates = dateDicts.flatMap(obj => Object.keys(obj)) 
  var uniqueDates = [...new Set(dates)].sort().reverse() ; // get rid of duplicates 


  // create an array of all our sites in our folder
  var allSites = Object.keys(siteDateKPIs).sort();

  // start writing to our sheet
  // set date headers in row 1 starting from column C (col 3)
  sheet.getRange(1, 3,1, uniqueDates.length).setValues([uniqueDates]);
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold");
  uniqueDates.forEach((date, i) => {
    dateColNum[date] = i + 3; 
  });

  // set our sites and values
  allSites.forEach((site, i) => {
    var row = i + 2
    sheet.getRange(row, 1).setValue(i + 1); //sets index value in col 1
    sheet.getRange(row, 2).setValue(site); // sets site name in col 2

    Object.entries(siteDateKPIs[site]).forEach(([date, kpis]) => {
      var col = dateColNum[date]; // find col index for date
      var cell = sheet.getRange(row, col);
      cell.setValue(kpis.join(', '));

      // populate cell map for easy lookup
      if (!cellMap[site]) {
        cellMap[site] = {};
      }
      var cellA1 = cell.getA1Notation;
      cellMap[site][date] = cellA1;
    });

  });

  // wrap text
  sheet.autoResizeColumns(1, sheet.getLastColumn());
 // sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).setWrap(true);

  // freeze index and site, bold site
  sheet.setFrozenColumns(2);
  sheet.getRange(1, 1, sheet.getLastRow(), 2).setFontWeight("bold"); 

//  console.log(uniqueDates); // 1219
  console.log(allSites);
  console.log(cellMap);

return cellMap;
  
}
