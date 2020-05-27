function cpAvailabilityFrom() { 
  //  =importrange("https://docs.google.com/spreadsheets/d/1kf13iLaI642Nhi8wr7c41p-UNDX76I419a3DC36AP-c", "Availability!A2:H114")
  //  A2:H120
  var startRow = 2;
  var startCol = 1; // A
  var endRow = 120;
  var endCol = 7;   // G
  var numRows = endRow - startRow + 1;
  var numColumns = endCol - startCol + 1;
  
//  var startCell = 'A2';
//  var endCell = 'G120';
  
  var sourceSheetName = 'Availability';
  var targetSheetName = 'Availability';
  
  var sourceSSdocID = '1kf13iLaI642Nhi8wr7c41p-UNDX76I419a3DC36AP-c';
  
  var ssSource = SpreadsheetApp.openById(sourceSSdocID); // getActiveSpreadsheet();
  var sourceSheet = ssSource.getSheetByName(sourceSheetName); 
//  var sourcRange = sourceSheet.getRange(startCell + ":" + endCell).getValues();
  var sourcRange = sourceSheet.getRange(startRow, startCol, numRows, numColumns).getValues();
  var srcUpdateTime = sourceSheet.getRange("K1").getValue();
  
  var ssTarget = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ssTarget.getSheetByName(targetSheetName);
  var targetRange = targetSheet.getRange(startRow, startCol, numRows, numColumns).setValues(sourcRange); // .activate(); // .getValues();
  var trgUpdateTime = targetSheet.getRange("P1").setValue(srcUpdateTime); 

  var end = new Date(); // Get script ending time
//  var lastExecDate = "Updated on " + end.getDate() + "/" + (end.getMonth() +1 ) +  "/" + end.getFullYear() + " at " + end.getHours() + ":" + end.getMinutes() + ":" + end.getSeconds();
//  ssTarget.getRange("U1").setValue(lastExecDate);
  var curDate = end.getDate();
  var triggerDate = 1;
  if (curDate == triggerDate) {
    var crtNewSheet = createNewSheet(triggerDate);
  }
}

function createNewSheet(triggerDate) {
  var end = new Date(); // Get script ending time
  var curDate = end.getDate();

  if (curDate == triggerDate) {
    var sourceSSdocID = '1xGulUuODWgcO0rq4FCelGMFGmYZ4W-8QTwITBgty5m0'; // "List of Properties and Descriptions"
    var sourceSheetName = 'RatesCopy';
    var ssSource = SpreadsheetApp.openById(sourceSSdocID); // getActiveSpreadsheet();
    var sourceSheet = ssSource.getSheetByName(sourceSheetName); 
    var srcMonth = sourceSheet.getRange("B2").getValue();
    var srcYear = sourceSheet.getRange("B1").getValue();
    var newSheetName = srcMonth.toString()+srcYear.toString();
//    newSheetName = "new2-"+ newSheetName;
    
    var ssTarget = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ssTarget.getSheets()[0].activate(); // Note that the JavaScript index is 0, but this will log 1
    var firstSheetName = sheet.getName();
//    Logger.log("firstSheetName = " + firstSheetName + "; newSheetName = " + newSheetName );
//    Exception: A sheet with the name "new2-May20" already exists. Please enter another name. (line 71, file "Code")
    if (firstSheetName != newSheetName) {
      ssTarget.duplicateActiveSheet();
      ssTarget.getActiveSheet().setName(newSheetName);
    } 
  }
  return "done";
}

function onOpen(e){
  var menuItems = [{name: "Update Availability", functionName: "cpAvailabilityFrom"}];
  SpreadsheetApp.getActive().addMenu("ELM Menu", menuItems);
  cpAvailabilityFrom();
}
