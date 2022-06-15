/*
#=================================================================================================
#title           :Roster.gs
#document        :"1. Roster 2020 ELM"
#document's link : https://docs.google.com/spreadsheets/d/1gSagEZltDqyXnuLlWyk-mSNnf92tEQkjxuacvc44EhA/edit#gid=1724961346
#description     :This script creates new fortnight sheets, puts the respective date's header,
                  moves old sheets into archive, copy list of RO's (Responsible Officer's)
                  from expired fortnight to the new one.
#author          :Ilya KIM
#email           :egift7@gmail.com
#date            :2021-07-12
#version         :0.1
#usage           :
#notes           :
#                :
#=================================================================================================

ToDo:
1. Triger to run auto creation every second Friday
2. MovingOut put to respective cells of Row = 16
   Check the changes in MO and mend the respective sheets (make a permanent values to compare with?)
3. remove Availability By Unit Type (rows 33-45 ?)

4. Manage sheets by ID

MO 7BM_(28/7/2021) new should be 30/10/2021

*/

//triggers that works on certain event
function onOpen(e){
  //var checkTabColor=setSheetTabColor();
  // Add a custom menu to the active spreadsheet, including a separator and a sub-menu.
  SpreadsheetApp.getUi()
  .createMenu('<ELM>')
  .addItem('Fill MO, MI dates to respective cells', 'setMOdatesFromRangePassToSetDatesToRespectiveCells')
  .addSeparator()
//  .addSeparator()
  .addItem('Create new fortnight tab', 'activateNewFortnightArchiveOld')
  .addSeparator()
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Miscellaneous...')
//              .addItem('Put color code into colored celles', 'putValueToColoredCell')
//              .addItem('Clear weeks sheets', 'ClearRange')
              .addItem('About...', 'getAbout')
//              .addSeparator()
//              .addItem('test---save/download file', 'doGet')              
             )
  .addToUi();
  }

function getAbout() {
  Browser.msgBox('This is an ELM\'s fortnight Roster.');
}


function triggerActivateNewFortnightArchiveOld() {
  // ============== create new fortnight sheet every second Thursday/Friday =================
  //  var triggerDayIs = 'Thu'; // 'Fri' , 'Sat', 'Mon','Tue','Wed','Thu','Fri','Sat.'
  var dTodayIs = new Date(); // Get script ending time
  var strStartDate = dTodayIs.toString();
  var partsStrStartDate =strStartDate.split(' '); 
  var sStartWeekDay = partsStrStartDate[0];
  
  if (sStartWeekDay == 'Thu' || sStartWeekDay == 'Mon' ) {
    activateNewFortnightArchiveOld();
  }
}

function setDatesToRespectiveCells(sUnitName, sDateToSet, objDateToSet, intRowToInsert) {  // , srcSheet, objStartDateIs
  // it's enough to use only two fortnights: 'currentSN' and 'plus1SN' => don't need to scan all sheets.
  var intPeriodLastingIs = 14;
  var intPeriodStartIndexIs = 0;
  var  intStartColIs = 4;
  var arrAllTabSheets = getAllRelevantPeriodsToArray();
  var sCurSheetNme = arrAllTabSheets['currentSN'];
  var sCurrStartDateIs = arrAllTabSheets['currentStartDate'];
  var sPlus1SheetNme = arrAllTabSheets['plus1SN'];
  var sPlus1StartDateIs = arrAllTabSheets['plus1StartDate'];
//  var objStartDateIs = sCurrStartDateIs; 
  var objStartDateIs =  new Date(sCurrStartDateIs); 
  
  //        startPeriodConvertedDate = new Date(startPeriodformatedMMDDYYYY);
//  let objDateToSet = arrMovingOutUnitsDates[cntUnitsOfAgrStatus][2];
  
  var srcSheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sCurSheetNme).activate();
  var srcSheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sPlus1SheetNme).activate();
  var diffInDaysActualMIMinus1stFortnightDay = diffDate1Date2(objStartDateIs, objDateToSet)*(-1); // differences should also point to column
  
  if ( diffInDaysActualMIMinus1stFortnightDay >= intPeriodStartIndexIs && diffInDaysActualMIMinus1stFortnightDay <= (2*intPeriodLastingIs - 1)) {
//    Logger.log("!!! diffDays =  " + diffInDaysActualMIMinus1stFortnightDay + " ; srcSheetNameIs =  " + srcSheet2.getSheetName() +  " ; intRowToInsert =  " + intRowToInsert.toString() + " ; sUnitName =  " + sUnitName);
    if ( diffInDaysActualMIMinus1stFortnightDay >= intPeriodStartIndexIs && diffInDaysActualMIMinus1stFortnightDay <= (intPeriodLastingIs - 1) ) { 
      var srcSheetNameIs = srcSheet1;
      var intColToInsert = intStartColIs + diffInDaysActualMIMinus1stFortnightDay;
    } else if ( diffInDaysActualMIMinus1stFortnightDay >= intPeriodLastingIs && diffInDaysActualMIMinus1stFortnightDay <= (2*intPeriodLastingIs - 1) ) {
      var srcSheetNameIs = srcSheet2;
      var intColToInsert = intStartColIs + diffInDaysActualMIMinus1stFortnightDay - intPeriodLastingIs;
    } else {
      var srcSheetNameIs = "DateIsToFarOrPassedAlready";
      var intColToInsert = "DateIsToFarOrPassedAlready";
    }
    if ( srcSheetNameIs != "DateIsToFarOrPassedAlready" ) {
//      Logger.log("objStartDateIs =  " + objStartDateIs.toString() + " ; objDateToSet =  " + objDateToSet.toString());
//      Logger.log(" ; diffInDaysActualMIMinus1stFortnightDay =  " + diffInDaysActualMIMinus1stFortnightDay);
//      Logger.log(" >>>>>>>>>>>>> diffDays =  " + diffInDaysActualMIMinus1stFortnightDay + " ; srcSheetNameIs =  " + srcSheetNameIs.getSheetName() +  " ; intRowToInsert =  " + intRowToInsert.toString() + " ; intColToInsert =  " + intColToInsert.toString() + " ; sUnitName =  " + sUnitName);
      srcSheetNameIs.activate();
      let sCopiedExistedValue = srcSheetNameIs.getRange(intRowToInsert, intColToInsert).getValue();
      if (sCopiedExistedValue.length > 1) { // if (sCopiedExistedValue !="") {
        sUnitName = sCopiedExistedValue + "\r\n" + sUnitName;
      }
//      let sPrevMOdatePreserv = srcSheetNameIs.getRange(intRowToInsert+1, intColToInsert).getValue();
      srcSheetNameIs.getRange(intRowToInsert, intColToInsert).setValue(sUnitName); // row, column
    } 
  }
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sCurSheetNme).activate();
}

function convertDateToString(objDate) {
  var arrMonths = {'Jan':'01', 'Feb':'02', 'Mar':'03', 'Apr':'04', 'May':'05', 'Jun':'06', 'Jul':'07', 'Aug':'08', 'Sep':'09', 'Oct':'10', 'Nov':'11', 'Dec':'12'};
  // Mon Jul 05 2021 00:00:00 GMT+1000 (Australian Eastern Standard Time)
  var sVarType = typeof objDate; 
  Logger.log("sVarType " + sVarType + " ; objDate " + objDate);
  if (sVarType == 'object') { // 'object', 'string' 
    var strStartDate = objDate.toString();
    var partsStrStartDate =strStartDate.split(' '); 
    var sStartWeekDay = partsStrStartDate[0];
    var sStartMonth = partsStrStartDate[1];
    var sStartMonthNum = arrMonths[sStartMonth];
    var sStartDay = partsStrStartDate[2];
    var sStartYear = partsStrStartDate[3];
    var strDateIs = sStartDay + "/" + sStartMonthNum + "/" + sStartYear;
  } else {
    var strDateIs = objDate;
  }
  return strDateIs;
}
//       getStartingRow("PreT - SU (Entered Manually)",srcsPlus1SN, 3,10);
function getStartingRow(sFindVal, srcSheet, colStart, rowStart) {
  var arrGetSearchRange = srcSheet.getRange( rowStart, colStart, 60, 1).getValues();
  for (var i = 0; i < arrGetSearchRange.length; i++ ) {
    let strSrcMOdatesTmp = arrGetSearchRange[i][0].toString();
    
//    Logger.log("strSrcMOdatesTmp = " + strSrcMOdatesTmp);
    if ( sFindVal.trim() == strSrcMOdatesTmp.trim()) {
       var sRowSearchedValLocation = i + rowStart + 1;
//      Logger.log("i = " + i + " ; sRowSearchedValLocation = " + sRowSearchedValLocation);
      //  i = 32 ; sRowSearchedValLocation = 35
    }
  }
  return sRowSearchedValLocation;
}

function setMOdatesFromRangePassToSetDatesToRespectiveCells() {
//  var spreadsheet = SpreadsheetApp.getActive();
  // it's enough to use only two fortnights: 'currentSN' and 'plus1SN' => don't need to scan all sheets.
  //  arrAllPeriods = {'currentStartDate':'', 'currentSN':'', 'plus1StartDate':'', 'plus1SN':'', 'plus2StartDate':'', 'plus2SN':''};
  var arrAllTabSheets = getAllRelevantPeriodsToArray();
  //  Logger.log("arrAllTabSheets = " + arrAllTabSheets);
  var sCurSheetNme = arrAllTabSheets['currentSN'].toString();
  
  //  Logger.log("sCurSheetNme = !!! " + sCurSheetNme);
  
  var sPlus1SN = arrAllTabSheets['plus1SN'].toString();
  //  var sCurrStartDateIs = arrAllTabSheets['currentStartDate'];
  var intRowToInsert = 0;
  const  sStartColumn = 3, sNumRows = 7, sNumColumns = 1;
  // "PreT - SU (Entered Manually)"
  var srcsPlus1SN = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sPlus1SN).activate();
  var sStartRows_Plus1SN = getStartingRow("Agreement's Status",srcsPlus1SN, 3,3); // should be 36
  
  const intStartColIToClear = 4, numRowsToClear = 7, numColumnsToClear = 14;
  srcsPlus1SN.getRange(sStartRows_Plus1SN, intStartColIToClear, numRowsToClear, numColumnsToClear).clearContent();
  
  var srcSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sCurSheetNme).activate();
  var sStartRow_srcSheet = getStartingRow("Agreement's Status",srcSheet, 3,10);
  srcSheet.getRange(sStartRow_srcSheet, intStartColIToClear, numRowsToClear, numColumnsToClear).clearContent();
  
  let iRowNumDiff = parseInt(sStartRow_srcSheet) - parseInt(sStartRows_Plus1SN);
  let iRowStartInsetDel = 25;
//  Logger.log("iRowNumDiff = >>> " + iRowNumDiff);
  if (iRowNumDiff > 0) {
    //    srcsPlus1SN.insertRows(rowIndex, numRows);
    srcsPlus1SN.insertRows(iRowStartInsetDel, iRowNumDiff);
  }  else if (iRowNumDiff < 0) {
    //    srcsPlus1SN.deleteRows(rowPosition, howMany);
    srcsPlus1SN.deleteRows(iRowStartInsetDel, Math.abs(iRowNumDiff));
//    spreadsheet.getRange('A44').activate();
  }
  var sCurrentMOdateUpdated = "", sPrevMOdatePreserved = "";
              
  var arr2DRangeOfUnitsDates = srcSheet.getRange(sStartRow_srcSheet, sStartColumn, sNumRows, sNumColumns).getValues();  // (row, column, numRows, numColumns);
  var arrAgrStatesToRendere = ["COL - NT", "COL - SU", "COL - AG", "COL - MI", "MO - MO", "Moving Out"]; // , "Available" , "Vacant"
  for (var i = 0; i < arrAgrStatesToRendere.length; i++ ) {
    var sAgrStateIs = arrAgrStatesToRendere[i].toString().trim();
//    intRowToInsert = sStartRow;
    var iRowCnt = 0;
    for (var j = 0; j < arr2DRangeOfUnitsDates.length; j++ ) {
//      Logger.log("[j] = " + j.toString() + " ; arr2DRangeOfUnitsDates[j] = " + arr2DRangeOfUnitsDates[j].toString());
      //  [j] = 0 ; arr2DRangeOfUnitsDates[i] = COL - NT : 16J24_(30/4/2021), 3J15_(4/5/2021), [21-07-19 19:04:28:815 AEST] 
//      Logger.log(" ; arr2DRangeOfUnitsDates[0][1] = " + arr2DRangeOfUnitsDates[0][1] + " ; arr2DRangeOfUnitsDates[1][1] = " + arr2DRangeOfUnitsDates[1][1]);
      
      let strSrcMOdatesTmp = arr2DRangeOfUnitsDates[j].toString();
      //      Logger.log("strSrcMOdatesTmp = " + strSrcMOdatesTmp.toString());
      let arrSrcMOdatesTmp = strSrcMOdatesTmp.split(":");
      let strAgrStatus2DRange = arrSrcMOdatesTmp[0].toString().trim();
      if ( strAgrStatus2DRange == sAgrStateIs ) {
//        Logger.log(" ; strAgrStatus2DRange = " + strAgrStatus2DRange.toString() + " ; arr2DRangeOfUnitsDates[1] = " + arr2DRangeOfUnitsDates[1].toString() + " ; arr2DRangeOfUnitsDates[2] = " + arr2DRangeOfUnitsDates[2].toString());
        intRowToInsert = sStartRow_srcSheet + iRowCnt;
        var arrMovingOutUnitsDates = getMOdatesToArray(strSrcMOdatesTmp);
        for (var cntUnitsOfAgrStatus = 0; cntUnitsOfAgrStatus < arrMovingOutUnitsDates.length; cntUnitsOfAgrStatus++ ) { 
          let sUnitName = arrMovingOutUnitsDates[cntUnitsOfAgrStatus][0];
          let sDateToSet = arrMovingOutUnitsDates[cntUnitsOfAgrStatus][1];
          let objDateToSet = arrMovingOutUnitsDates[cntUnitsOfAgrStatus][2];
          // arrUnitsDatesMO[0] =  16J24,30/4/2021,Fri Apr 30 2021 00:00:00 GMT+1000 (Australian Eastern Standard Time)
          // setDatesToRespectiveCells(sUnitName, sDateToSet, objDateToSet, intRowToInsert, srcSheet, objStartDateIs)
//          setDatesToRespectiveCells(arrMovingOutUnitsDates[0][0], arrMovingOutUnitsDates[0][1], arrMovingOutUnitsDates[0][2], intRowToInsert, srcSheet, sCurrStartDateIs); //(sUnitName, sDateToSet);
//          Logger.log(" ; sUnitName = " + sUnitName + " ; sDateToSet = " + sDateToSet.toString() + " ; objDateToSet = " + objDateToSet.toString()+ " ; intRowToInsert = " + intRowToInsert.toString());
          if (sAgrStateIs == 'MO - MO') {
            // let sPrevMOdatePreserved = srcSheet.getRange(row, column, numRows, numColumns).getValues() // Object[][]
            let arrCurrentMOdateUpdated = srcSheet.getRange(intRowToInsert, intStartColIToClear,1 ,numColumnsToClear).getValues();
//            Logger.log("10. arrCurrentMOdateUpdated = " + arrCurrentMOdateUpdated);
            let arrPrevMOdatePreserved = srcSheet.getRange(intRowToInsert + 1, intStartColIToClear,1 ,numColumnsToClear).getValues();
            Logger.log("20. arrCurrentMOdateUpdated[0].length = " + arrCurrentMOdateUpdated[0].length);
            for (var cntCol=0; cntCol < arrCurrentMOdateUpdated[0].length; cntCol++) {
              sCurrentMOdateUpdated = arrCurrentMOdateUpdated[0][cntCol];
              sPrevMOdatePreserved = arrPrevMOdatePreserved[0][cntCol];
              Logger.log("1. sCurrentMOdateUpdated = " + sCurrentMOdateUpdated + "; sPrevMOdatePreserved = " + sPrevMOdatePreserved);
//              Logger.log("2. sPrevMOdatePreserved = " + sPrevMOdatePreserved);
              if (sCurrentMOdateUpdated != "" && sCurrentMOdateUpdated != sPrevMOdatePreserved) {
                var sNewPreservedMO = sPrevMOdatePreserved + sCurrentMOdateUpdated;
                srcSheet.getRange(intRowToInsert + 1, intStartColIToClear+cntCol).setValue(sNewPreservedMO);
                Logger.log("3. sNewPreservedMO = " + sNewPreservedMO);
              }
            }
          }
          
          setDatesToRespectiveCells(sUnitName, sDateToSet, objDateToSet, intRowToInsert); // , srcSheet, objStartDateIs
        }
      }
      iRowCnt++;
    }
  }
//  Logger.log("setMOdatesFromRangePassToSetDatesToRespectiveCells : work is done ");
}

function getMOdatesToArray(strSrcMOdates) {  // arr2DRangeOfUnitsDates
////  "COL - NT", "COL - SU", "COL - AG", "COL - MI", 'MO - MO'
  var arrSrcMOdatesTmp = strSrcMOdates.split(":"); // arrSrcMOdatesTmp = Moving Out , 2J24_(31/7/2021), 22D_(4/8/2021), 7BM_(28/7/2021), 7FM_(3/8/2021), 
  var arrSrcMO_UnitsAndDates = arrSrcMOdatesTmp[1].split("), ");  // arrSrcMOdates =  2J24_(31/7/2021),22D_(4/8/2021),7BM_(28/7/2021),7FM_(3/8/2021),
  var arrUnitsDatesMO =[];
  for (var i = 0; i < arrSrcMO_UnitsAndDates.length - 1; i++ ) {
    if (!arrUnitsDatesMO[i]) arrUnitsDatesMO[i] = [];
    arrUnitsDatesMO[i] = arrSrcMO_UnitsAndDates[i].split("_(");
    arrUnitsDatesMO[i][2] = strDateToObject(arrUnitsDatesMO[i][1]);
    // Logger.log("[i] = " + i + " ; arrUnitsDatesMO[i] = " + arrUnitsDatesMO[i]); // [i] = 0 ; arrUnitsDatesMO[i] =  2J24,31/7/2021
    //// [i] = 0 ; arrUnitsDatesMO[i][0] =  2J24 ; arrUnitsDatesMO[i][1] = 31/7/2021
    // Logger.log("[i] = " + i + " ; arrUnitsDatesMO[i][0] = " + arrUnitsDatesMO[i][0]  + " ; arrUnitsDatesMO[i][1] = " + arrUnitsDatesMO[i][1]  + " ; arrUnitsDatesMO[i][2] = " + arrUnitsDatesMO[i][2]);
    //// [i] = 3 ; arrUnitsDatesMO[i][0] = 7FM ; arrUnitsDatesMO[i][1] = 3/8/2021 ; arrUnitsDatesMO[i][2] = Tue Aug 03 2021 00:00:00 GMT+1000 (Australian Eastern Standard Time)
  }
//  Logger.log("arrUnitsDatesMO[0] = " + arrUnitsDatesMO[0]);
  return arrUnitsDatesMO;
  
}

function getAllRelevantPeriodsToArray() {
  var arrExistingSheets = [];
  var cntSheetTabNum = 0;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
//  var sheet = ss.getSheets()[0]; // Note that the JavaScript index is 0, but this will log 1
  var sheetName = ""; // sheet.getName();
  
  var dateSheetNameParserArray = [];
  var sheetLoop = 0;
  var startDayFromSN = "";
  var startMonthFromSN = "";
  
  var endDayFromSN = "";
  var endMonthFromSN = "";
  
  var endYearFromSN = "";
  var startYearFromSN = "";
  
  var startPeriodformatedMMDDYYYY = "";
  var startPeriodConvertedDate = "";
  
  var endPeriodformatedMMDDYYYY = "";
  var endPeriodConvertedDate = "";
  
  var NewStartPeriod = '';
  var arrAllPeriods = {'minus2StartDate':'', 'minus2SN':'', 'minus1StartDate':'', 'minus1SN':'', 'currentStartDate':'', 'currentSN':'', 'plus1StartDate':'', 'plus1SN':'', 'plus2StartDate':'', 'plus2SN':'', 'plus3StartDate':'', 'plus3SN':'', 'plus4StartDate':'', 'plus4SN':''};
  var now_objDate = new Date;
  now_objDate.setHours(0, 0, 0, 0);
//  const d = new Date();
//d.setHours(15, 35, 1);
  var objNewStartPeriod  = new Date;
  
  for (var i = 0; i < sheets.length ; i++ ) {
    sheetLoop = sheets[i];
    sheetName = sheetLoop.getSheetName();
    dateSheetNameParserArray = sheetName.split("-");
    // 05July-18July-21  //ddmmm-ddmmm-YY
    if (dateSheetNameParserArray.length == 3) {
      startDayFromSN = dateSheetNameParserArray[0].toString().substring(0,2);
      startMonthFromSN = dateSheetNameParserArray[0].toString().substring(2,dateSheetNameParserArray[0].length);
      
      endDayFromSN = dateSheetNameParserArray[1].toString().substring(0,2);
      endMonthFromSN = dateSheetNameParserArray[1].toString().substring(2,dateSheetNameParserArray[1].length);
      
      startYearFromSN = dateSheetNameParserArray[2].toString();
      endYearFromSN = dateSheetNameParserArray[2].toString();
      
//      Logger.log("==== >>>> dateSheetNameParserArray => " + dateSheetNameParserArray);
//      Logger.log("==== >>>> startDayFromSN => " + startDayFromSN);
//      Logger.log("==== >>>> startMonthFromSN => " + startMonthFromSN);
//      Logger.log("==== >>>> startYearFromSN => " + startYearFromSN);
//      
      if (parseInt(dateSheetNameParserArray[0])>18 && startMonthFromSN == "Dec") {
        startYearFromSN = (dateSheetNameParserArray[2] - 1).toString();
//        Logger.log("33333.startYearFromSN =>>> " + startYearFromSN);
      }
      //      startPeriodformatedMMDDYYYY = startDayFromSN + '/' + startMonthFromSN + '/' + startYearFromSN;
      startPeriodformatedMMDDYYYY = startMonthFromSN + '/' + startDayFromSN + '/' + startYearFromSN;
      
      
//      Logger.log("startPeriodformatedMMDDYYYY => " + startPeriodformatedMMDDYYYY);
      startPeriodConvertedDate = new Date(startPeriodformatedMMDDYYYY);
      
//      endPeriodformatedMMDDYYYY = endDayFromSN + '/' + endMonthFromSN + '/' + endYearFromSN;
      endPeriodformatedMMDDYYYY = endMonthFromSN + '/' + endDayFromSN + '/' + endYearFromSN;
      endPeriodConvertedDate = new Date(endPeriodformatedMMDDYYYY);
//      now_objDate.setHours(0, 0, 0, 0);
//      Logger.log("cntSheetTabNum => " + cntSheetTabNum);
      if (!arrExistingSheets[cntSheetTabNum]) arrExistingSheets[cntSheetTabNum] = [];
      arrExistingSheets[cntSheetTabNum][0] = sheetName.toString();
//      Logger.log("arrExistingSheets[cntSheetTabNum][0] => " + arrExistingSheets[cntSheetTabNum][0]);
      arrExistingSheets[cntSheetTabNum][1] = startPeriodConvertedDate.toString();
      arrExistingSheets[cntSheetTabNum][2] = endPeriodConvertedDate.toString();
      cntSheetTabNum++;
      
      var iPeriod = 14;

//      Logger.log("now_objDate => " + now_objDate.getTime() + " ; startPeriodConvertedDate => " + startPeriodConvertedDate + " ; endPeriodConvertedDate => " + endPeriodConvertedDate.getTime());
//      Logger.log("now_objDate >= startPeriodConvertedDate = " + (now_objDate >= startPeriodConvertedDate) + " ; now_objDate <= endPeriodConvertedDate = " + (now_objDate <= endPeriodConvertedDate));
//      if ( now_objDate >= startPeriodConvertedDate && now_objDate.getTime() <= endPeriodConvertedDate.getTime() ) { 
      if ( now_objDate >= startPeriodConvertedDate && now_objDate <= endPeriodConvertedDate ) {   
        
//        Logger.log("1111.startPeriodConvertedDate = " + startPeriodConvertedDate);
        NewStartPeriod = getNewStartPeriod(startPeriodConvertedDate,-2*iPeriod); // 21/6/2021
//        Logger.log("2222.NewStartPeriod >>>> = >>>> " + NewStartPeriod); // Thu Sep 02 2021 12:19:53 GMT+1000 (Australian Eastern Standard Time)
        arrAllPeriods['minus2StartDate'] = NewStartPeriod;
        objNewStartPeriod  = new Date(NewStartPeriod);
        arrAllPeriods['minus2SN'] = getTabSheetName(objNewStartPeriod);

        
        NewStartPeriod = getNewStartPeriod(startPeriodConvertedDate,-1*iPeriod);
        arrAllPeriods['minus1StartDate'] = NewStartPeriod.toString();
        objNewStartPeriod  = new Date(NewStartPeriod);
        arrAllPeriods['minus1SN'] = getTabSheetName(objNewStartPeriod);
        
        
        NewStartPeriod = getNewStartPeriod(startPeriodConvertedDate,1*iPeriod);
        arrAllPeriods['plus1StartDate'] = NewStartPeriod.toString();
        objNewStartPeriod  = new Date(NewStartPeriod);
        arrAllPeriods['plus1SN'] = getTabSheetName(objNewStartPeriod);
        
        
        NewStartPeriod = getNewStartPeriod(startPeriodConvertedDate,2*iPeriod);
        arrAllPeriods['plus2StartDate'] = NewStartPeriod.toString();
        objNewStartPeriod  = new Date(NewStartPeriod);
        arrAllPeriods['plus2SN'] = getTabSheetName(objNewStartPeriod);
        
        NewStartPeriod = getNewStartPeriod(startPeriodConvertedDate,3*iPeriod);
        arrAllPeriods['plus3StartDate'] = NewStartPeriod.toString();
        objNewStartPeriod  = new Date(NewStartPeriod);
        arrAllPeriods['plus3SN'] = getTabSheetName(objNewStartPeriod);
        
        
        NewStartPeriod = getNewStartPeriod(startPeriodConvertedDate,4*iPeriod);
        arrAllPeriods['plus4StartDate'] = NewStartPeriod.toString();
        objNewStartPeriod  = new Date(NewStartPeriod);
        arrAllPeriods['plus4SN'] = getTabSheetName(objNewStartPeriod);
        
        NewStartPeriod = getNewStartPeriod(startPeriodConvertedDate,0*iPeriod);
        arrAllPeriods['currentStartDate'] = NewStartPeriod.toString();
        objNewStartPeriod  = new Date(NewStartPeriod);
        arrAllPeriods['currentSN'] = getTabSheetName(objNewStartPeriod);  
        
//        if (!arrAllPeriods['arrExistingSheets']) arrAllPeriods['arrExistingSheets'] = [];
        arrAllPeriods['arrExistingSheets'] = arrExistingSheets;
//        Logger.log("arrAllPeriods => " + arrAllPeriods);
        //  Logger.log("arrAllPeriods['arrExistingSheets'] =>>> " + arrAllPeriods['arrExistingSheets']);
      } // if current tab
    } // if fortnight tab
  } // loop i
  
//  Logger.log("arrAllPeriods = " + arrAllPeriods);
//  Logger.log("arrExistingSheets => " + arrExistingSheets);
  //  arrAllPeriods['arrExistingSheets'] = undefined
  
  return arrAllPeriods; // arrExistingSheets
}

function activateNewFortnightArchiveOld() {
  var arrAllRelevantSheets = getAllRelevantPeriodsToArray();
  var arrAllExistingSheets = arrAllRelevantSheets['arrExistingSheets'];
  var rowStart = 1, columnStart = 4;
  var startPeriodConvertedDate = arrAllRelevantSheets['currentStartDate'];
  var sheetName = arrAllRelevantSheets['currentSN'];
  
  // fill the header of current fortnight by dates
  var dateArr = getDateToFillHeader(startPeriodConvertedDate);
//  Logger.log("sheetName: " + sheetName + "\n");
//  Logger.log("dateArr: " + dateArr + "\n");
//  Logger.log("dateArr[0]: " + dateArr[0] + "\n");
//  //   TypeError: Cannot read property 'length' of undefined
//  Logger.log("dateArr.length: " + dateArr.length + "\n");
//  Logger.log("dateArr[0].length: " + dateArr[0].length + "\n");
////  TypeError: Cannot read property 'getRange' of null
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(rowStart, columnStart, dateArr.length, dateArr[0].length).setValues(dateArr); // .getRange(row, column, numRows, numColumns)
  createArchiveSheets(arrAllRelevantSheets, arrAllExistingSheets);
}

function createArchiveSheets(arrAllPeriods = [], arrExistingSheets = []) {
// TypeError: Cannot read property 'activate' of null
  var rosterSS = SpreadsheetApp.getActiveSpreadsheet();

  if (rosterSS.getSheetByName(arrAllPeriods['minus2SN']) != null) {
    //    Logger.log("moveTabToArchive(arrAllPeriods['minus2SN']): " + arrAllPeriods['minus2SN'] + "\n");
    moveTabToArchive(arrAllPeriods['minus2SN']);  // moveTabToArchive(sheetTabNameToMove)
  }
  
  // ==== just in case, check and move all old roster's sheet to archive ==== 
  var diffInDays = 0;
  for (var i = 0; i < arrExistingSheets.length ; i++ ) {
//    Logger.log("TabName: " + arrExistingSheets[i] + "\n");
    diffInDays = diffDate1Date2(arrAllPeriods['currentStartDate'], arrExistingSheets[i][1]);
    //    Logger.log("diffInDays: " + diffInDays + "\n");
    if (diffInDays >= 25) {
      //      Logger.log("moveTabToArchive(arrExistingSheets[i][0]): " + arrExistingSheets[i][0] + "\n");
      moveTabToArchive(arrExistingSheets[i][0]);  // moveTabToArchive(sheetTabNameToMove)
    }
  }
  // ==== just in case, check and move all old roster's sheet to archive ==== 
  var intSheetIndexPos = 1;
  createSheets(rosterSS, arrAllPeriods['plus1SN'], arrAllPeriods['plus1StartDate'], arrAllPeriods['currentSN'], intSheetIndexPos++);
  createSheets(rosterSS, arrAllPeriods['plus2SN'], arrAllPeriods['plus2StartDate'], arrAllPeriods['currentSN'], intSheetIndexPos++);
  createSheets(rosterSS, arrAllPeriods['plus3SN'], arrAllPeriods['plus3StartDate'], arrAllPeriods['currentSN'], intSheetIndexPos++);
//  createSheets(rosterSS, arrAllPeriods['plus4SN'], arrAllPeriods['plus4StartDate'], arrAllPeriods['currentSN'], intSheetIndexPos++);
  if (rosterSS.getSheetByName(arrAllPeriods['minus1SN']) != null) {
    rosterSS.getSheetByName(arrAllPeriods['minus1SN']).activate();
    rosterSS.moveActiveSheet(intSheetIndexPos++);
  }
  
  if (rosterSS.getSheetByName(arrAllPeriods['currentSN']) != null) {
    let srcStartRow = 2, srcStartCol = 1, srcNumRows = 29, srcNumCol = 22;
    
    let srcCurrentSheet = rosterSS.getSheetByName(arrAllPeriods['currentSN']).activate();
    rosterSS.moveActiveSheet(1);
    let srcRange = srcCurrentSheet.getRange(srcStartRow, srcStartCol, srcNumRows, srcNumCol).getValues(); // .setValues(srcRange); 
    
    let dstPlus1Sheet = rosterSS.getSheetByName(arrAllPeriods['plus1SN']).activate();
    rosterSS.moveActiveSheet(2);
    let dstRange = dstPlus1Sheet.getRange(srcStartRow, srcStartCol, srcNumRows, srcNumCol).setValues(srcRange);  
    
    let dstPlus2Sheet = rosterSS.getSheetByName(arrAllPeriods['plus2SN']).activate();
    rosterSS.moveActiveSheet(3);
    dstPlus2Sheet.getRange(srcStartRow, srcStartCol, srcNumRows, srcNumCol).setValues(srcRange); 
    
  }
  rosterSS.getSheetByName(arrAllPeriods['currentSN']).activate();
}

function createSheets(rosterSS, sNewTabSheetName, startPeriodConvertedDate, sCurrentTabSheetName, iNewTabSheetIndex) {
  Logger.log("sNewTabSheetName = " + sNewTabSheetName + " ; startPeriodConvertedDate = " + startPeriodConvertedDate  + " ; sCurrentTabSheetName = " +  sCurrentTabSheetName  + " ; iNewTabSheetIndex = " +  iNewTabSheetIndex);
  if (rosterSS.getSheetByName(sNewTabSheetName) == null) {
    rosterSS.getSheetByName(sCurrentTabSheetName).activate();
    rosterSS.duplicateActiveSheet();
    rosterSS.renameActiveSheet(sNewTabSheetName);
    //  clear UnitsAgreements range's content
    let row = 36, column = 4, numRows = 28, numColumns = 14;
    rosterSS.getActiveSheet().getRange(row, column, numRows, numColumns).clearContent();
  }
  rosterSS.getSheetByName(sNewTabSheetName).activate();
  rosterSS.moveActiveSheet(iNewTabSheetIndex);
  let rowStart = 1, columnStart = 4;
  let dateArr = getDateToFillHeader(startPeriodConvertedDate);
  rosterSS.getSheetByName(sNewTabSheetName).getRange(rowStart, columnStart, dateArr.length, dateArr[0].length).setValues(dateArr); // .getRange(row, column, numRows, numColumns)
  rosterSS.getSheetByName(sCurrentTabSheetName).activate();
}

function moveTabToArchive(sheetTabNameToMove) {
  var rosterSS = SpreadsheetApp.getActiveSpreadsheet();
  
  var allRostersArchiveID_Val = '12ZQjv9nUjb0W8KKpvrrYh0u_6lT2vmU7DY_W1zhorsc';
  var rosterArchiveSS = SpreadsheetApp.openById(allRostersArchiveID_Val);
  
  if (rosterArchiveSS.getSheetByName(sheetTabNameToMove) == null) {
    var sheetLast = rosterSS.getSheetByName(sheetTabNameToMove);
    var newArcLastSheet = sheetLast.copyTo(rosterArchiveSS);
    
    var whatissheetname = newArcLastSheet.getSheetName();
//    Logger.log("whatissheetname = " + whatissheetname); //
    // A sheet with the name "24Jul.-06Aug.-17" already exists. Please enter another name. (line 532, file "Code")    
    newArcLastSheet.setName(sheetTabNameToMove);
    newArcLastSheet.activate();
    newArcLastSheet.showSheet();
    rosterArchiveSS.moveActiveSheet(1); // move newly created sheet to first position
    //      rosterSS.deleteSheet(sheet)
    rosterSS.deleteSheet(sheetLast); //or instead of delete the sheet - rename and hide it?
    //      SpreadsheetApp.openById("your spreadsheet id").insertSheet("ToTo");
  }
}

function getNewStartPeriod(startDay, cntj) {
  // Mon Jul 05 2021 00:00:00 GMT+1000 (Australian Eastern Standard Time)
  var newDate='';
  var secondDate = new Date();
  
  var sVarType = typeof startDay; 
  if (sVarType == 'string') {
    var dateToHandle = convertDateFormat(startDay); 
  } else {
    var dateToHandle = startDay;
  }
////  someDate.setDate(someDate.getDate() + numberOfDaysToAdd); 
//  secondDate.setDate(dateToHandle.getDate() + cntj);
  
  var secondDate = new Date(dateToHandle.getTime()+(cntj*24*60*60*1000));
  
  newDate = secondDate;
//  Logger.log("getNewStartPeriod.startDay is ==>> " + startDay);
//  Logger.log("getNewStartPeriod.cntj is ==>> " + cntj);
//  Logger.log("getNewStartPeriod.dateToHandle is ==>> " + dateToHandle);
//  Logger.log("getNewStartPeriod.secondDate is ==>> " + secondDate);
  
  return newDate;
}

function getTabSheetName(startPeriod) {
  // 05July-18July-21  //ddmmm-ddmmm-YY
//  Logger.log("getTabSheetName.startPeriod is ==>> " + startPeriod);
//  var sEndPeriod = getNewStartPeriod(startPeriod, 13);
  var sEndPeriod = addDays(startPeriod, 13);  // addDays(date, days)
  var endPeriod = new Date(sEndPeriod);
//  Logger.log("getTabSheetName.endPeriod is =====>>>> " + endPeriod);
  
  var arrMonths = {'Jan':'01', 'Feb':'02', 'Mar':'03', 'Apr':'04', 'May':'05', 'Jun':'06', 'Jul':'07', 'Aug':'08', 'Sep':'09', 'Oct':'10', 'Nov':'11', 'Dec':'12'};

  // Mon Jul 05 2021 00:00:00 GMT+1000 (Australian Eastern Standard Time)
  var strStartDate = startPeriod.toString();
  var partsStrStartDate =strStartDate.split(' '); 
  var sStartWeekDay = partsStrStartDate[0];
  var sStartMonth = partsStrStartDate[1];
  var sStartMonthNum = arrMonths[sStartMonth];
  var sStartDay = partsStrStartDate[2];
  var sStartYear = partsStrStartDate[3];
  
  var strEndDate = endPeriod.toString();
  var partsStrEndDate =strEndDate.split(' '); 
  var sEndWeekDay = partsStrEndDate[0];
  var sEndMonth = partsStrEndDate[1];
  var sEndMonthNum = arrMonths[sEndMonth];
  var sEndDay = partsStrEndDate[2];
  var sEndYear = partsStrEndDate[3];
  
//  06Jul-undefinedundefined-undefined
  // 05July-18July-21  //ddmmm-ddmmm-YY
  var sSheetName = sStartDay + sStartMonth + "-" + sEndDay + sEndMonth + "-" + sEndYear;
  
//  Logger.log("getTabSheetName.sSheetName is !!! " + sSheetName);
  
  return sSheetName;

}

function getDateToFillHeader(startDay) {
// 24/05/2021   dd/mm/yyyy
//  var saDays = ['1st              MONDAY                           ','Mon.','Tue.','Wed.','Thu.','Fri.','Sat.'];
  //  var saDays = ['SUNDAY.', 'MONDAY','Tue.','Wed.','Thu.','Fri.','Sat.'];
  var saDays = ['MONDAY','TUESDAY','WEDNSDAY','THUSDAY','FRIDAY','SATURDAY','SUNDAY'];
  var saDays = ['Mon.','Tue.','Wed.','Thu.','Fri.','Sat.','Sun.'];
  var arrMonths = {'Jan':'01', 'Feb':'02', 'Mar':'03', 'Apr':'04', 'May':'05', 'Jun':'06', 'Jul':'07', 'Aug':'08', 'Sep':'09', 'Oct':'10', 'Nov':'11', 'Dec':'12'};
  
  var prefWeekCnt = ['1st', '2nd'];
  var dateArr = [];
//  var setDate = "";
  // typeof operator maps an operand to one of six values: "string", "number", "object", "function", "undefined" and "boolean".
  var sVarType = typeof startDay; 
//  var i=0;
  var cntj=0;
  var cntsaDays=0;
  var cntprefWeekCnt=0;
  if (sVarType == 'string') {
    var dateToHandle = convertDateFormat(startDay); 
  } else {
    var dateToHandle = startDay;
  }
  var secondDate = new Date();
  
  for (cntprefWeekCnt = 0; cntprefWeekCnt < prefWeekCnt.length; cntprefWeekCnt++) {
    for (cntsaDays = 0; cntsaDays < saDays.length; cntsaDays++) {
      
      var sSecondDate = addDays(startDay, cntj);  // addDays(date, days)
      var secondDate = new Date(sSecondDate);      
      
      var strDate = secondDate.toString();
      //      Logger.log('strDate is :' + strDate); // Mon Oct 30 2017 00:00:00 GMT+1100 (AEDT)
      var partsNextDate =strDate.split(' '); 
      var sWeekDay = partsNextDate[0];
      var sMonth = partsNextDate[1];
      var sMonthNumIs = arrMonths[sMonth];
      //      var sMonthNum = (parseInt(sMonth) + 1).toString();
      var sDay = partsNextDate[2];
      var sYear = partsNextDate[3];
//      Logger.log('strDate is :' + strDate.toString());
//      Logger.log('cntj is :' + cntj.toString() + ' ; sMonth is :' + sMonth.toString() + ' ; sMonthNum is :' + sMonthNumIs.toString() + ' ; sDay is :' + sDay.toString());
      
      var sDay_next = sWeekDay + "-" + sDay + "-" + sMonth;
      ////      setDate = dateToHandle + i*24*3600*1000;
      ////      setDate = new Date(dateToHandle);
      //      Logger.log('cntj is :' + cntj);
      //      Logger.log('secondDate is :' + secondDate);
      //      dateArr[cntj] = 'User Specific AA4 Templates'
      /* ==============================================================
      setValues accepts(and getValues() returns):
            1 argument of type:
      Object[][] a two dimensional array of objects
            It does NOT accept a 1 dimensional array. A range is always two dimensional, regardless of the range height or width.
      
      If A1:A2 is the range, then corresponding values array would be like:
            [[1],[3]]
      ============================================================== */
      if (!dateArr[0]) dateArr[0] = [];
//      dateArr[0][cntj] = prefWeekCnt[cntprefWeekCnt] + " " + saDays[cntsaDays] + " \r\n" + secondDate.getDate().toString() + "/" + (secondDate.getMonth() + 1).toString() + "/" + secondDate.getFullYear().toString();
      dateArr[0][cntj] = prefWeekCnt[cntprefWeekCnt] + " " + saDays[cntsaDays] + " \r\n" + sDay + "/" + sMonthNumIs + "/" + sYear;
      
      cntj++;
//      dateArr[cntj] = prefWeekCnt[cntprefWeekCnt] + " " + saDays[cntsaDays] + " " + sDay + "/" + (secondDate.getMonth() + 1).toString() + "/" + secondDate.getFullYear().toString();
      //      } // end i loop
    } // end cntsaDays loop
  } // end prefWeekCnt
  
  return dateArr;
  
}


function getDateToFillHeader123(startPeriodDayIs) {
  //  var saDays = ['SUNDAY.', 'MONDAY','Tue.','Wed.','Thu.','Fri.','Sat.'];
//  var saDays = ['MONDAY','TUESDAY','WEDNSDAY','THUSDAY','FRIDAY','SATURDAY','SUNDAY'];
  var saDays = ['Mon.','Tue.','Wed.','Thu.','Fri.','Sat.','Sun.'];
  var arrMonths = {'Jan':'01', 'Feb':'02', 'Mar':'03', 'Apr':'04', 'May':'05', 'Jun':'06', 'Jul':'07', 'Aug':'08', 'Sep':'09', 'Oct':'10', 'Nov':'11', 'Dec':'12'};
  var prefWeekCnt = ['1st', '2nd'];
  var dateArr = [];
  // typeof operator maps an operand to one of six values: "string", "number", "object", "function", "undefined" and "boolean".
  var sVarType = typeof startPeriodDayIs; 
//  var i=0;
  if (sVarType == 'string') {
    var dateToHandle = convertDateFormat(startPeriodDayIs); 
  } else {
    var dateToHandle = startPeriodDayIs;
  }
//  var setDate = '';
  
  var cntprefWeekCnt=0;
  var cntsaDays=0;
  var cntJdays=0;
  var secondDate = dateToHandle;
  
  for (cntprefWeekCnt = 0; cntprefWeekCnt < prefWeekCnt.length; cntprefWeekCnt++) {
//    secondDate = dateToHandle;
    for (cntsaDays = 0; cntsaDays < saDays.length; cntsaDays++) {
      
      var sDay_next = sWeekDay + "-" + sDay + "-" + sMonth;
      
      //      secondDate = dateToHandle;
      //      var result = new Date(date);
      //      result.setDate(result.getDate() + days);
      secondDate.setDate(dateToHandle.getDate() + cntJdays);
      var strDate = secondDate.toString();
      var partsNextDate = strDate.split(' '); 
      var sWeekDay = partsNextDate[0];
      var sMonth = partsNextDate[1];
      var sMonthNum = arrMonths[sMonth];
      var sDay = partsNextDate[2];
      var sYear = partsNextDate[3];
      Logger.log('strDate is :' + strDate.toString());
      Logger.log('cntj is :' + cntj.toString() + ' ; sMonth is :' + sMonth.toString() + ' ; sMonthNum is :' + sMonthNumIs.toString() + ' ; sDay is :' + sDay.toString());
      
      if (!dateArr[0]) dateArr[0] = [];
      //      dateArr[0][cntJdays] = prefWeekCnt[cntprefWeekCnt] + " " + saDays[cntsaDays] + " \r\n" + secondDate.getDate().toString() + "/" + (secondDate.getMonth() + 1).toString() + "/" + secondDate.getFullYear().toString();
//      dateArr[0][cntJdays] = prefWeekCnt[cntprefWeekCnt] + " " + saDays[cntsaDays] + " " + sDay + "/" + (secondDate.getMonth() + 1).toString() + "/" + secondDate.getFullYear().toString();
//      dateArr[0][cntJdays] = prefWeekCnt[cntprefWeekCnt] + " " + saDays[cntsaDays] + " \r\n" + sDay + "/" + sMonth + "-" + sMonthNum + "/" + sYear;
      dateArr[0][cntJdays] = prefWeekCnt[cntprefWeekCnt] + " " + saDays[cntsaDays] + " \r\n" + sDay + "/" + sMonthNum + "/" + sYear;
      cntJdays++;
      
    } // end cntsaDays loop
  } // end prefWeekCnt
  
  return dateArr;
  
}

function diffDate1Date2(oDate1, oDate2) {
  // convert date format dd/mm/yyyy to mm/dd/yyyy
  var oDate1 =  new Date(oDate1); 
  var oDate2 =  new Date(oDate2); 
  var diffDate1Date2 = 0;
//  Logger.log("1 ; oDate1 is " + oDate1 + " ; oDate2 is " + oDate2);
//  oDate1 = strDateToObject(oDate1);
//  oDate2 = strDateToObject(oDate2);
//  Logger.log("2 ; oDate1 is " + oDate1 + " ; oDate2 is " + oDate2);
//Logger.log("3 ; oDate1.getTime() is " + oDate1.getTime() + " ; oDate2.getTime() is " + oDate2.getTime());
  var diff_in_days = Math.floor((oDate1.getTime() - oDate2.getTime())/(24*3600*1000));
//  Logger.log("diff_in_days is " + diff_in_days);
  return diff_in_days; // return object - Date
}

function addDays(date, days) {
  var result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}


function convertDateFormat(sDDMMYYY) {  
  // convert date format dd/mm/yyyy to mm/dd/yyyy
  var datearray = sDDMMYYY.split("/");
  var formatedMMDDYYYY = datearray[1] + '/' + datearray[0] + '/' + datearray[2];
  var convertedDate = new Date(formatedMMDDYYYY);
  //  Logger.log("formatedMMDDYYYY is " + formatedMMDDYYYY);
  return convertedDate; // return object - Date
}

// ==== date sandbox =====

//function convertDateFormat(sDDMMYYY) {  // var result = new Date(date);
function strDateToObject(strDate) {
  //undefined
  // ======================
//  var initialDate = '30/10/2017';
  var sVarType = typeof strDate; 
  if (sVarType == 'string') {
    var parts =strDate.toString().split('/'); 
    //please put attention to the month (parts[0]), Javascript counts months from 0: // Apr '04/03/2014'
    // January - 0, February - 1, etc
    var objDate = new Date(parts[2],parts[1]-1,parts[0]); // yyyy,mm,dd
    //  Logger.log('objDate is :' + objDate); // 
  } else {
    var objDate = new Date(); // today's date
  }
  return objDate;
}


function getWeekDay(srcDate) {
  var saDays = ['Sun','Mon.','Tue.','Wed.','Thu.','Fri.','Sat.'];
  // typeof operator maps an operand to one of six values: "string", "number", "object", "function", "undefined" and "boolean".
  var sVarType = typeof srcDate; 
  if (sVarType == 'string') {
    var dateToHandle = convertDateFormat(srcDate); 
  } else {
    var dateToHandle = srcDate;
  }
//  var dateToHandle = srcDate;
  var iIsDay = dateToHandle.getDay();
  var getDayStr = saDays[iIsDay];
  return getDayStr;
}

function tmp() {
  srcDate = "21/06/2020";
  var convertedDate = new Date(srcDate);
  var wDay = getWeekDay(convertedDate);
//    var wDay = getWeekDay(srcDate);

  Logger.log("wDay is " + wDay);
  
  var new_now = new Date;
  var now = new Date(new_now.getTime());
  var fbsheet = SpreadsheetApp.getActive().getSheetByName('Facebook Advertising Schedule');
    
  var days_ago1 = parseInt(fbsheet.getRange("J1").getValue());
  
  var strdate = fb_sheet_data[r][c].toString().substring(1 + fb_sheet_data[r][c].toString().lastIndexOf(" "));
  var cellday   = pad(parseInt(strdate.toString().substring(0, strdate.indexOf("."))), 2);
  var cellmonth = pad(parseInt(strdate.toString().substring(strdate.indexOf(".")+1)), 2);
  var celldate = new Date(now.getYear().toString() + "-" + cellmonth + "-" + cellday);
  var diff_in_days = Math.floor((now - celldate)/(24*3600*1000));
  //Logger.log([fb_sheet_data[r][c], r, c, strdate, celldate, diff_in_days]);
  if       (diff_in_days >= days_ago1 && diff_in_days < days_ago2) {
  } else if(diff_in_days >= days_ago4) {
    fbsheet.getRange(row_offset + r, col_offset + c).setBackground(bg4);
  }
}

function test_addDays() {
  var initialDate = '30/10/2017';
  var parts =initialDate.split('/'); 
  //please put attention to the month (parts[0]), Javascript counts months from 0: // Apr '04/03/2014'
  // January - 0, February - 1, etc
  
  var mydate = new Date(parts[2],parts[1]-1,parts[0]); // yyyy,mm,dd
  Logger.log('mydate is :' + mydate); // 
  
  for (var i = 0; i < 14 ; i++ ) {
    var nextDaytmp = addDays(mydate,i);
//    [17-10-28 18:20:02:832 AEDT] nextDaytmp is :Mon Oct 30 2017 00:00:00 GMT+1100 (AEDT)
    
    var strDate = nextDaytmp.toString();
    Logger.log('strDate is :' + strDate); // Mon Oct 30 2017 00:00:00 GMT+1100 (AEDT)
    var partsNextDate =strDate.split(' '); 
    var sWeekDay = partsNextDate[0];
    var sMonth = partsNextDate[1];
    var sDay = partsNextDate[2];
    var sYear = partsNextDate[3];
    //    Logger.log('sWeekDay is :' + sWeekDay);
    //    Logger.log('sMonth is :' + sMonth);
    //    Logger.log('sDay is :' + sDay);
    //    Logger.log('sYear is :' + sYear);
//    // that works perfectly fine
//    var wd = nextDaytmp.getDay() 	// Get the weekday as a number (0-6)
//    var dd = nextDaytmp.getDate();
//    var mm = nextDaytmp.getMonth() + 1;
//    var y = nextDaytmp.getFullYear();
////    Logger.log('dd is :' + dd);
////    Logger.log('mm is :' + mm);
////    Logger.log('y is :' + y);
  }
}



// ============================================

function daylyUnitsAgrStausUpdate() {
  // https://docs.google.com/spreadsheets/d/1ISEVWWoNLm4ro76WPzYxIOodkIIZ5ZsUcaUEvTmL6gY/edit#gid=1691623190 // 'as_at_for_creator_v1.gss'
  //  "as_at_for_creator_v1.gss" => "1ISEVWWoNLm4ro76WPzYxIOodkIIZ5ZsUcaUEvTmL6gY" , "fortnight_periods", "B1:K85", col "9" is a source of values for tab name
  // ========== "as_at_for_creator_v1.gss" ============
  var srcStartRow = 1, srcStartCol = 11 , srcNumRows = 85, srcNumCol = 10;
  var srcOfPeriodsSSdocID = '1ISEVWWoNLm4ro76WPzYxIOodkIIZ5ZsUcaUEvTmL6gY'; // 'as_at_for_creator_v1.gss'
  var srcOfPeriodsSheetName = 'fortnight_periods';
  var srcOfPeriodsSS = SpreadsheetApp.openById(srcOfPeriodsSSdocID); // getActiveSpreadsheet();
  var srcOfPeriodsSheet = srcOfPeriodsSS.getSheetByName(srcOfPeriodsSheetName); 
  var dstRange = srcOfPeriodsSheet.getRange(srcStartRow, srcStartCol, srcNumRows, srcNumCol).setValues(srcRange);  
  //  ========= fortnight tab ========
  // D11:D29 => Total amount
  // E11:E29 => Status 
  // #AgrStatusToken# 
  // #UnitsStatusToken#
  var dstStartRow= 2, dstStartCol = 2, dstNumRows = 29, dstNumCol = 60;
  var dstSheetName = 'AgrCopy';
  var dstSS = SpreadsheetApp.getActiveSpreadsheet();
  var arrAllSheets = srcSS.getSheets();
  var dstSheet = srcSS.getSheetByName(srcSheetName); 
  var dstRange = srcSheet.getRange().getValues(dstStartRow, dstStartCol, dstNumRows, dstNumCol);
  
  //  1st              MONDAY                           06/07/2020
    var end = new Date(); // Get script ending time
//    var timeElapsed = (end.getTime() - start.getTime()) / 1000;
    var lastExecDate = "Updated on " + end.getDate() + "/" + (end.getMonth() +1 ) +  "/" + end.getFullYear() + " at " + end.getHours() + ":" + end.getMinutes() + ":" + end.getSeconds();
    

}

function myFunction() {
  var str_l = ["COL - NT", "COL - SU", "COL - AG", "COL - MI", "LIP - CT", "LIP - AMO", "LIP - LFR", "LIP - ITV", "LIP - LE", "LIP - LT", "MO - PreMO", "MO - MO", "MO - CO", "MO - LFR", "MO - ITV", "MO - NTV", "MO - EVCT", "EOL - BOL", "EOL - NTV/EVCT", "EOL - Dispute", "EOL - VCAT", "EOL - PT", "EOL - Archive", "EOL - Cancelled", "Offer"];
  length = str_l.length;
  Logger.log("length is : " + length);
  
}


function main() {
  triggerActivateNewFortnightArchiveOld();
  setMOdatesFromRangePassToSetDatesToRespectiveCells();
}


// ==== get sheet by its ID ====

//  var ssSheetID = rosterSS.getSheetByName(arrAllPeriods['currentSN']).getSheetId(); // integer
//  Logger.log("tst currentSN ssSheetID: " + ssSheetID + "\n"); // tst currentSN ssSheetID: 1605837246
//  var ssIndex = rosterSS.getSheetByName(arrAllPeriods['currentSN']).getIndex(); // tst currentSN ssIndex: 1
//  Logger.log("tst currentSN ssIndex: " + ssIndex + "\n");
  
//  function getSheetById(id) {
//  return SpreadsheetApp.getActive().getSheets().filter(
//    function(s) {return s.getSheetId() === id;}
//  )[0];
//}
//
//var sheet = getSheetById(123456789);
//  In example above gid=910184575, so you can do something like this:
//
//function getSheetById_test(){
//  var sheet = getSheetById(910184575);
//  Logger.log(sheet.getName());
//}
//
//function getSheetById(gid){
//  for each (var sheet in SpreadsheetApp.getActive().getSheets()) {
//    if(sheet.getSheetId()==gid){
//      return sheet;
//    }
//  }
//}

//  function createArchiveSheets(arrAllPeriods = [], arrExistingSheets = []) {
// associative array looping
  //  var keys = [];
  //      arrExistingSheets[cntSheetTabNum][0] = sheetName;
  //      arrExistingSheets[cntSheetTabNum][1] = startPeriodConvertedDate;
  //      arrExistingSheets[cntSheetTabNum][2] = endPeriodConvertedDate;
  //  for(var k in arrAllPeriods) {
  ////    keys.push(k+':'+arrAllPeriods[k]);
  //    
  //  }
  //  for(var k in arrAllPeriods) keys.push(k+':'+arrAllPeriods[k]);
  //  Logger.log("total " + keys.length + "\n" + keys.join('\n'));
  //  Logger.log("arrExistingSheets.length: " + arrExistingSheets.length + "\n");
  //  
  ////    var rosterSS_ID = "1gSagEZltDqyXnuLlWyk-mSNnf92tEQkjxuacvc44EhA";
  ////  var rosterSS = SpreadsheetApp.openById(rosterSS_ID);
//  }

      ////      sSetDate = dateToHandle + i*24*3600*1000;
      ////      sSetDate = new Date(dateToHandle);
      /* ==============================================================
      setValues accepts(and getValues() returns):
            1 argument of type:
      Object[][] a two dimensional array of objects
            It does NOT accept a 1 dimensional array. A range is always two dimensional, regardless of the range height or width.
      
      If A1:A2 is the range, then corresponding values array would be like:
            [[1],[3]]
      ============================================================== */


