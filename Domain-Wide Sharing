// create a script that populates a google sheet filled with all files an individual has that is shared domain wide
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Domain-Wide Files')
  .addItem('Get Domain-Wide Shared Files', 'lockExecution')
  .addToUi();
}

//function makeSheet() {
//  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//  var sheet = spreadsheet.insertSheet('Shared Files');
//  sheet.clear();
// sheet.appendRow(['File Name', 'Link to File']);
//  sheet.activate();
//  return sheet;
//}

function makeSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Shared File(s)";
  var sheet = null;
  var currentDate = new Date();
  var formattedDate = Utilities.formatDate(currentDate, spreadsheet.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  while (sheet == null) {
    var nameToCheck = sheetName + " " + formattedDate;
    var existingSheet = spreadsheet.getSheetByName(nameToCheck);
    if (existingSheet == null) {
      sheet = spreadsheet.insertSheet(nameToCheck);
    }
    formattedDate = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }

  sheet.clear();
  sheet.appendRow(['File Name', 'Link to File']);
  sheet.activate();
  return sheet;
}


function lockExecution() {
  var lock = LockService.getScriptLock();
 // lock.waitLock(15000);
  try {
    lock.waitLock(5000);
    findDomainWide();
  } catch(error) {
    console.log(error);
    SpreadsheetApp.getUi().alert("An error ocurred with your script. Wait till the script finishes running to try and run it again.");

  } finally {
    lock.releaseLock();
  }

}

function findDomainWide() {
  var start = Date.now();
  // first, get a list of files accessible by the user
  mysheet = makeSheet();
  var nextPageToken = ''; // initialize next page token
  
  do {
    var files = Drive.Files.list({
      corpora: 'default', // correct corpora value confusing, play around with this
      q: 'trashed=false',
      includeItemsFromAllDrives: false,
      supportsAllDrives: false, // include shared drive
      pageToken: nextPageToken, // use next page token
      pageSize: 1000
      }); 

      // loop through all files
    for(var i = 0; i < files.items.length; i++) {
      var file = files.items[i];

      // loop through the permissions for each file.
      // check if manager (organizer) or owner
      // then once that is established, check if file is discoverable and shared to the domain
      
      var isDomainShared = false;

      if (file.userPermission.role === 'owner'){
        try {
          var permissions = Drive.Permissions.list(file.id);
          for (var j = 0; j < permissions.items.length; j++) {
          var permission = permissions.items[j];

          if (isDomainShared === false) {
            if (permission.type === 'domain') {
              isDomainShared = true;
            }

          }

        }
        } catch(error) {
            continue;

        }
      }
      
      if (isDomainShared) {
        mysheet.appendRow([file.title, file.alternateLink]);
        }

      }
    nextPageToken = files.nextPageToken; // set next page token to next page
    } while(nextPageToken);

    var end = Date.now();
    var timetaken = (end-start)/60000+" minutes";
    mysheet.appendRow([' ', ' ']);
    mysheet.appendRow(['Time Taken To Complete', timetaken]);
}
