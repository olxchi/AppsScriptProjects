/* Runs a query that selects all proxy data server users of the last 7 days. Scheduled to run monthly.  */

const FOLDERID = '000000000000000000000000000000';

// extracts data into a sheet (removed refresh component, refresh scheduled on connected sheet already and it's too long to run effectively here)
function refreshAndExtract() {

  // make the connected sheet the active one
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Connected Sheet 1'), true);

  // refresh the data and wait for completion
  SpreadsheetApp.enableAllDataSourcesExecution();
//  spreadsheet.getActiveSheet().asDataSourceSheet().refreshData();
//  spreadsheet.getActiveSheet().asDataSourceSheet().waitForCompletion(300);

  // create table from data source, and then put it on new sheet (new sheet immediately becomes active one)
  var dataSource = spreadsheet.getActiveSheet().asDataSourceSheet().getDataSource();
  var dataSourceTable = dataSource.createDataSourceTableOnNewSheet();
  dataSourceTable.syncAllColumns();
}

// creates a spreadsheet, moves it into folder
function moveFileToFolder() {
  // active sheet should be the extracted one
  SpreadsheetApp.flush();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var newSpreadsheet = SpreadsheetApp.create('Proxy Server Users: ' + Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy"));
  Utilities.sleep(2000);
  var file = DriveApp.getFileById(newSpreadsheet.getId());

  // copy sheet to new spreadsheet
  SpreadsheetApp.flush(); // this ensures the new spreadsheet actually exists by telling spreadsheet to wait till updated
  sheet.copyTo(newSpreadsheet);

    // remove the default 'Sheet1' from the new spreadsheet
  var defaultSheet = newSpreadsheet.getSheetByName('Sheet1');
  if (defaultSheet) {
    newSpreadsheet.deleteSheet(defaultSheet);
  }

  //move file to folder
  var folder = DriveApp.getFolderById(FOLDERID);
  file.moveTo(folder);

  Logger.log('Sheet copied and new spreadsheet moved to the folder successfully.');

}

function main() {
  refreshAndExtract();
  moveFileToFolder();
}

function createMonthlyTrigger() {
  // Create a time-driven trigger that runs on the first of every month at 1 AM
  ScriptApp.newTrigger('main')
    .timeBased()
    .onMonthDay(3)           // Run on the first day of the month
    .atHour(10)               // Run at 1 AM (Actually now runs at 8AM)
    .create();
}