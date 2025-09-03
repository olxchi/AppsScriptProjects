// each person has their own dollar per hour value. default: $111.6
// user should be able to interact and change dollar per hour value
const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet(); 

//excel output sheet
const sheet = SPREADSHEET.getSheetByName("Excel Output");

// dollar per hour sheet
const DPHSHEET = SPREADSHEET.getSheetByName("Dollar Per Hour Value");

// raw data sheet
const RAWDATA = SPREADSHEET.getSheetByName("Raw Data");

// sheet with all possible codes/names
var LISTCODES = SPREADSHEET.getSheetByName("List Of Known Time Codes");

// list of Team Members
// create a team member object with fields: dphval, fn, ln, ID
// not currently in use, but could be helpful in the future
function TeamMember(fn, ln, tID) {
  this.dphval = 111.60; //default
  this.fn = fn;
  this.ln = ln;
  this.tID = tID; 
  this.projects = {}; // key: project, val: hours spent on proj
}

// global list of team members
var TEAM = [];

// update dph fxn 
function setDPH(TeamMember, newDPHVal) {
  TeamMember.dphval = newDPHVal;
}

function getMemberFromID(tID) {
  for (var i = 0; i < TEAM.length; i++) {
    if (TEAM[i].tID === tID) {
      return TEAM[i];
    }
  }
  return -1;
}

// do these things whenever the spreadsheet is opened
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Cleanup')
  .addItem('Remove Duplicates + Leading Zeros','lockExecution')
  .addToUi();
}


// do these things whenever the spreadsheet is edited 
function onEdit() {
  RAWDATA.getRange("A2").setFormula('=UNIQUE(FILTER(\'Excel Output\'!AF4:AF, \'Excel Output\'!AF4:AF <> ""))');
  SpreadsheetApp.flush();
  fillProgramNames();
}

/* removes duplicates, removes leading zeros from time codes*/
/* also, populates dictionary with key information from raw data */

function removeDuplicates() {
  // var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var data = sheet.getDataRange().getValues();
  var newData = [];
  var seen = {};

  // todo: get rid of first two rows (date + blank)

  // loop through each row, if that row hasn't been seen before, add it to our list of seen data and also list of fresh data, 
  for (var i = 0; i < data.length; i++) {
    var row = data[i].join('');
    if(!seen[row]) {
      newData.push(data[i]);
      seen[row] = true;
    } else {
      Logger.log(row);
    }
  }

  sheet.clearContents();
  sheet.getRange(1,1, newData.length, newData[0].length).setValues(newData);

  // removes leading zeros once duplicates are here (REMEMBER: keep at end or else clear contents will mess it up)
  leadingZeros();

}

// removes leading zeros from row AE in the spreadsheet, called in removeDuplicates 
function leadingZeros() {
  var lastrow = sheet.getLastRow();
  var range = sheet.getRange("AE4:AE" + lastrow);
  var values = range.getValues();

  var modifiedValues = values.map(function(row) {
    // check if cell is empty (skip), use regex to remove leading zeroes
    var cell = row[0].toString();
    if (cell !== '') {
      return [cell.replace(/^0+/,'')];
    } else {
      return [cell];
    }
   // return [row[0].toString().replace(/^0+/,'')];
  });

//  Logger.log(modifiedValues);

  var range2 = sheet.getRange("AE4:AE" + lastrow);
  range2.setValues(modifiedValues);
  
}

// returns a list of team member objects
function createTMObjects() {
  var data = DPHSHEET.getRange("A2:D" + DPHSHEET.getLastRow()).getValues();

  for (let i = 0; i < data.length; i++) {
    var tID = data[i][0];
    var fn = data[i][1];
    var ln = data[i][2];
    var dph = data[i][3];

    var mymem = new TeamMember(fn, ln, tID);
    TEAM.push(mymem);
  }

  Logger.log(TEAM);

}
// used to fill sheet "Raw Data" with corresponding timecode from "List of Known Timecodes"
function fillProgramNames() {
  // gete all data from 'list of known time codes'
  var timeCodeData = LISTCODES.getRange("A2:B").getValues();

  // create a dictionary to map code to program name
  var timeCodeDict = {}; // key: code, value: program name

  // for each row, if code exists, add to dict
  for (var i = 0; i < timeCodeData.length; i++) {
    var program = timeCodeData[i][0]; // Column A is Program
    var code = timeCodeData[i][1]; // Column B is Receiver WBS Element/Rec Order Code

    if (code) {
      timeCodeDict[code] = program;
    }

  }

  // now get data from 'raw data' - this is actually the unique codes reflected in excel output

  var data = RAWDATA.getRange("A2:B").getValues();

  // iterate through each row in "Raw Data"
  for (var j = 0; j < data.length; j++) {
    var rawDataCode = data[j][0]; // code is in Column A
    if (rawDataCode && timeCodeDict[rawDataCode]) { // if code exists in dict
      data[j][1] = timeCodeDict[rawDataCode]; // fill program in column B
    } else {
      data[j][1] = ""; //no match = empty cell
    }
    
  }

  // write code back to raw data sheet
  RAWDATA.getRange("A2:B" + (data.length + 1)).setValues(data);

}


// combines column AD and AE bc they're basically the same thing
function combineWBSREC() {
  var lastRow = sheet.getLastRow();
  var wbsCol = 30; //AD
  var recCol = 31; //AE
  var combineCol = 32; // index for new col

  var headers = sheet.getRange(3, 1, 1, sheet.getLastColumn()).getValues()[0];
  var combineColIndex = headers.indexOf("Receiver WBS Element/Rec Order Code") + 1;

  if (combineColIndex === 0) {
    sheet.insertColumnAfter(recCol);
    combineColIndex = combineCol;
    sheet.getRange(3, combineCol).setValue("Receiver WBS Element/Rec Order Code");

  }

  var range = sheet.getRange(4, wbsCol, lastRow - 3, 2); // col AD and AE

  var values = range.getValues();

  var combinedValues = [];

  for (var i = 0; i < values.length; i++) {
    var wbsVal = values[i][0]; // value from col AD
    var recVal = values[i][1]; // AE
    var combinedVal = wbsVal || recVal;

    combinedValues.push([combinedVal]);

  }

  sheet.getRange(4, combineCol, combinedValues.length, 1).setValues(combinedValues);

}

// makes it so that race conditions don't delete data
function lockExecution() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(100000);
    removeDuplicates();
    combineWBSREC();

  } catch (error) {
    console.log(error);
    SpreadsheetApp.getUi().alert("An error occurred. Wait till the script finishes running to try and run it again. / Une erreur s’est produite. Attendez que le script finisse de s’exécuter pour essayer de l’exécuter à nouveau.");

  } finally {
    lock.releaseLock();
  }

}
 
