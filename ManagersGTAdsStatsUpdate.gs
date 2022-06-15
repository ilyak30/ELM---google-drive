// 'Managers copy of List of Properties and Descriptions - Management'
// https://docs.google.com/spreadsheets/d/1cqkX9QE2QlSPuruoWS7PwejV3rKugoT5ShODVSIyWpA/edit#gid=1820459955

/*
#=================================================================================================
#title           :ManagersGTAdsStatsUpdate
#document        :"Managers copy of List of Properties and Descriptions - Management"
#document's link : https://docs.google.com/spreadsheets/d/1cqkX9QE2QlSPuruoWS7PwejV3rKugoT5ShODVSIyWpA/edit#gid=1820459955
#description     :This script updates RosterAgrUnitStatus,
                  
                  
#author          :Ilya KIM
#email           :egift7@gmail.com
#date            :2021-07-12
#version         :0.1
#usage           :
#notes           :
#                :
#=================================================================================================

ToDo:
1. Manage sheets by ID


*/


// Globals
// Illia can we have this broken up by unit type - there are 10 unit types - Share house single room / share house twin room / studio / 1BR / 2BR/ 3BR/4BR/5BR/6BR/7BR/8BR
// Share house single room / share house twin room / studio / 1BR / 2BR/ 3BR/4BR/5BR/6BR/7BR/8BR
//
// =VLOOKUP(A55,RatesCopy!A:G,6,0)
// =INDEX(RatesCopy!F:F, MATCH(A4, RatesCopy!A:A, 0), 1)
// =COUNTIF(C5:C, "Available")
// ='All Current'!A5
// ='All Current'!B5
//
// RatesCopy sheet uses a formula
//  
// A5 (A + E1) =importrange("https://docs.google.com/spreadsheets/d/1WUU4pvanWH5_iSUY2hPyXLfrTyR22MAHkDhRxSl5KTg", B2 & B1 & "!" & "A" & E1 & ":" & "A")
// B5 (E2 + E1) =importrange("https://docs.google.com/spreadsheets/d/1WUU4pvanWH5_iSUY2hPyXLfrTyR22MAHkDhRxSl5KTg", B2 & B1 & "!" & E2 & E1 & ":" & E2) 
// C5 (I + E1) =importrange("https://docs.google.com/spreadsheets/d/1WUU4pvanWH5_iSUY2hPyXLfrTyR22MAHkDhRxSl5KTg", B2 & B1 & "!" & "I" & E1 & ":" & "I")
// D5 (E + E1) =importrange("https://docs.google.com/spreadsheets/d/1WUU4pvanWH5_iSUY2hPyXLfrTyR22MAHkDhRxSl5KTg", B2 & B1 & "!" & "E" & E1 & ":" & "E")
// E5 (F  + E1) =importrange("https://docs.google.com/spreadsheets/d/1WUU4pvanWH5_iSUY2hPyXLfrTyR22MAHkDhRxSl5KTg", B2 & B1 & "!" & "F" & E1 & ":" & "F")
// F5 (H + E1) =importrange("https://docs.google.com/spreadsheets/d/1WUU4pvanWH5_iSUY2hPyXLfrTyR22MAHkDhRxSl5KTg", B2 & B1 & "!" & "H" & E1 & ":" & "H")
// G5 (B + E1) =importrange("https://docs.google.com/spreadsheets/d/1WUU4pvanWH5_iSUY2hPyXLfrTyR22MAHkDhRxSl5KTg", B2 & B1 & "!" & "B" & E1 & ":" & "B")
// it's better to make a script to copy those data instead of formula

// "#ERROR!" , "#NUM!" , "#NAME?"

var RC = 0; // index of "RatesCopy" in sheet_names[] array
var PR = 1; // "Private Rooms"
var TR = 2; // "Twin Rooms"
//var SL = 1; // "Single Leases - obsolete - by Gene 

var COLUMN_PC = 1;
var ROW_START = 2;

var COLUMN_SL_BILLS = 8;
var DESC = ['T', 'Y', 'R'];
//var sheet_names = ["Private Rooms", "Single Leases", "Twin Rooms", "RatesCopy"];
var sheet_names = ["RatesCopy", "Private Rooms", "Twin Rooms"];
var numOfSheets = sheet_names.length;

var fbsheet = SpreadsheetApp.getActive().getSheetByName('Facebook Advertising Schedule');
var FB_RANGE = ["I5", "BE107"]

var sheets = [];
//for(var sheet_id=0; sheet_id<sheet_names.length; ++sheet_id)
for(var sheet_id=0; sheet_id<numOfSheets; ++sheet_id)  
  sheets.push(SpreadsheetApp.getActive().getSheetByName(sheet_names[sheet_id]));

var sheets_data = [];
//for(var sheet_id=0; sheet_id<sheet_names.length; ++sheet_id) {
for(var sheet_id=0; sheet_id<numOfSheets; ++sheet_id) {  
  sheets_data.push(sheets[sheet_id].getDataRange().getValues());
}

//Logger.log("sheets_data[0][5][0].Values is " + sheets_data[0][5][0]);
//Logger.log("sheets_data[0][5][1].Values is " + sheets_data[0][5][1]);

function adjustUI() {
// "All Current"
// "Unit Information"
  var srcAllCur_sn = 'All Current';
  var srcUI_sn = 'Unit Information';
  
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss - spread sheet
  
  var ssAllCur = ss.getSheetByName(srcAllCur_sn);
  var ssUI = ss.getSheetByName(srcUI_sn);
  
  var ssAllCurStartRow = 1;
  var ssAllCurLastRow = ssAllCur.getLastRow();
  var ssAllCurStartCol = 1;
  var ssAllCurLastCol = ssAllCur.getLastColumn();
  var arrAllCur = ssAllCur.getRange(ssAllCurStartRow, ssAllCurStartCol, ssAllCurLastRow , ssAllCurLastCol).getValues();
  
  var ssUIStartCol = 1;
  var ssUILastRow = ssUI.getLastRow();
  var ssUIStartCol = 1;
  var ssUILastCol = ssUI.getLastColumn();
  var arrUI = ssUI.getRange(ssUIStartCol, ssUIStartCol, ssUILastRow , ssUILastCol).getValues();
  
  // Logger.log('arrPTasks : ssPTasksLastRow = ' + ssPTasksLastRow + ': ssPTasksLastColumn = ' + ssPTasksLastColumn + '\n');
  
  for (var row_i = 2; row_i < ssUILastCol; row_i++ ) {  // ssUILastCol = 116
    UIName = arrUI[row_i][0].trim();

    for (var col_j = 1; col_j < ssUILastCol; col_j++ ) {  // ssUILastColumn = 87
      
      UIName = arrUI[1][col_j].trim();               // var colProjName = 0; // column with projects name 
      valCellException = arrUI[row_i][col_j].trim();
     
        for (var rowPT_x = 1; rowPT_x < arrPTasks.length ; rowPT_x++ ) {
          PTaskNameFromArray = arrPTasks[rowPT_x][2].trim();
          posName = PTaskNameFromArray.indexOf(composedPTaskName);

          if  (  posName == 0 ) {
              arrPTasks[rowPT_x][12] = "_delete_;";
            arrPTasks.splice(rowPT_x , 1); // Removes the 1 element of arrPTasks from rowPT_x position
              cntDeletedPTasks++;
              arrUI[row_i][col_j] = "N/A_del_";
          } else {
            arrUI[row_i][col_j] = "N/A_chk_";
          }
        } // end 3rd for

      } // end 2nd for
  } // end 1st for  
  
  
  ssfinalPTCreated.getRange(1, 1, arrPTasks.length, arrPTasks[0].length).setValues(arrPTasks); 
  ssUI.getRange(1, 1, arrUI.length, arrUI[10].length).setValues(arrUI);
  
}  

function updateRosterAgrUnitStatus() {
  // 1st              MONDAY                           22/06/2020
  //  "1. Roster 2020 ELM" => https://docs.google.com/spreadsheets/d/1gSagEZltDqyXnuLlWyk-mSNnf92tEQkjxuacvc44EhA/edit#gid=638668067
  // docID = "1gSagEZltDqyXnuLlWyk-mSNnf92tEQkjxuacvc44EhA"
  // tab "AgregatedStatusInfo"
  // B2 - UnitsStatus, B14 - AgreementStatus
  //  NEW 22Jun-05Jul-20, 06Jul-19Jul-20, Roster_Template_base
  //  A12 => "#UnitsStatusToken#"   ; insert to  A13:A19
  //  A21 => "#AgrStatusToken#"     ; insert to  A23:A30
  // D - total number
  // E - status 
  // F - comma separated list of units
  var srcStartRow = 2, srcStartCol = 12 , srcNumRows = 75, srcNumCol = 60; // srcNumRows = 29
  var dstStartRow= 2, dstStartCol = 3, dstNumRows = 75, dstNumCol = 60;   // dstNumRows = 29
  var srcSheetName = 'AgrCopy';
  var srcSS = SpreadsheetApp.getActiveSpreadsheet();
  var srcSheet = srcSS.getSheetByName(srcSheetName); 
  var srcRange = srcSheet.getRange(srcStartRow, srcStartCol, srcNumRows, srcNumCol).getValues();
//  var srcUpdateTime = srcSheet.getRange(srcTimeCell).getValue();
  //  var ssSource = SpreadsheetApp.openById(sourceSSdocID); // getActiveSpreadsheet();
  //  var srcRange = srcSheet.getRange(startCell + ":" + endCell).getValues();
  
  // ========== ""1. Roster 2020 ELM"" ============
  var dstAgrUnitStatusSSdocID = '1gSagEZltDqyXnuLlWyk-mSNnf92tEQkjxuacvc44EhA';   // "1. Roster 2020 ELM"
  var dstAgrUnitStatusSheetName = 'AgregatedStatusInfo';
  var dstAgrUnitStatusSS = SpreadsheetApp.openById(dstAgrUnitStatusSSdocID); // getActiveSpreadsheet();
  var dstAgrUnitStatusSheet = dstAgrUnitStatusSS.getSheetByName(dstAgrUnitStatusSheetName); 
  
  var dstRange = dstAgrUnitStatusSheet.getRange(dstStartRow, dstStartCol, dstNumRows, dstNumCol).setValues(srcRange);
  
  var firstSheet = dstAgrUnitStatusSS.getSheets()[0].activate(); // Note that the JavaScript index is 0, but this will log 1
//  var firstSheetName = firstSheet.getName();
  var srcDateRange = firstSheet.getRange(1, 5, 40, 20).getValues();
  
//  srcRange[isRow][isCol] // D2:R10 => // column 5 search 'Moving Out'; 7CM_(20/12/2020)
  
  
}

function RegularExp() {
  var string = "Hello 2019 World, Welcome to Weird Geek 2019 1st           THURSDAY                             06/11/2020";
//  var regExp = new RegExp("([a-z]+\\s[a-z]+\\,\\s[a-z]+\\s[a-z]+\\s[a-z]+\\s[a-z]+)","gi"); 
  //   [0-9]+
//  var regExp = new RegExp("([0-9]+\\s[a-z]+)","gi"); 
//    var regExp = new RegExp("(\\d+\\s[a-z]+)","gi"); 
//    var regExp = new RegExp("(\\d+\/\\d+\/\\d+)","gi");  // 26/11/2020
//  var regExp = new RegExp("([1-9]+\/\\d+\/\\d+)","gi");  // 6/11/2020
//    var regExp = new RegExp("([0-9]\\d\/\\d+\/\\d+)","gi");  // 6/11/2020
    var regExp = new RegExp("(\\d+\/\\d+\/\\d+)","gi");  // 06/11/2020

  
  var lastName = regExp.exec(string)[1];
  Logger.log(lastName); 
}

function setDateForMO() {
//  var searchRange = "M2:AP11";
//  var cursheet = SpreadsheetApp.getActive().getSheetByName("AgrCopy");
//  var arrSearchRange = cursheet.getRange(searchRange).getValues();
//  
//  for (iRow=0;iRow<=10;iRow++) {
//    arrSearchRange[]
//    
//  }
//  


}

function getSummaryArrayOfUnitsAgrStatus(arrSearchRange, arrSearchPatern, arrExcept, ExeptedCol) {
//  "### arrSearchRange is: " + arrSearchRange + 
//  ### arrSearchPatern is: [object Object] ### arrExcept is: Suspended
//  Logger.log(" ### arrSearchPatern is: " + arrSearchPatern + " ### arrExcept is: " + arrExcept);
  var arrRez = [];
  var iRowToScan = 0;
  var irow = 0;
  var jcol =0;
  var icol = 0;
  var getCellVal = "";
  var getCellExeptVal = "";
  var getTotalVal = "";
  var iseek = 0;
  var sPaternVal = "";
  var sSoughtCellVal = "";
  var colcnt = 2; // leave 0 
  //  var maxColNum = 0;
  var colToCheck = 0;
  var retColVal = "";
  var colStart = 0;
  var sizeOfArray = arrSearchRange.length >> 1; // get the whole part of devision to 2
  
  var sExeptedVal = "";
  var sExeptionVal = "";
  var ecnt = 0;
  var inewrowcnt = 0;
  var sSummaryRow = "";
  var bOnlyThatStatus = false;
  var sMOdate;
  
//  Logger.log("sizeOfArray is: " + sizeOfArray);
  //     {"rowstart":2, "rowend":8,"colstart":10, "coltocheck":2,"returncol":1}
  colStart = arrSearchPatern["colstart"];
  
  sPaternVal = arrSearchRange[1][colStart];
//  Logger.log("colStart is: " + colStart + "sPaternVal is: " + sPaternVal );
//  //     var arrSearchPaternUnitStatus = {"rowstart":2, "rowend":11,"colstart":10, "coltocheck":1,"returncol":0};
//  //      var arrAgrStatusExceptions = ["EOL - Archive"];
//  //  var arrUnitStatusExceptions = ["Suspended"];
  
  for (iseek=0;iseek<=arrSearchPatern["rowend"]-arrSearchPatern["rowstart"];iseek++) {
    sPaternVal = arrSearchRange[iseek+arrSearchPatern["rowstart"]-1][colStart];
    
    for (ecnt = 0;ecnt<arrExcept.length;ecnt++) {
      sExeptionVal = arrExcept[ecnt];
      if (ExeptedCol > 0) {
        sExeptedVal = sExeptionVal;
        Logger.log("sExeptedVal is: " + sExeptedVal);
      }
      //    Logger.log("iseek is: " + iseek + " sPaternVal is: " + sPaternVal);
      
      if (sPaternVal == sExeptionVal) {
        sPaternVal = "";
        //        Logger.log("ecnt is: " + ecnt + " sExeptionVal is: " + sExeptionVal + " sPaternVal is: " + sPaternVal);
        //        sExeptedVal = "";
        //        break;
      }
      
    }
    if (sPaternVal != "") {
      if (!arrRez[inewrowcnt]) arrRez[inewrowcnt] = [];
      arrRez[inewrowcnt][colcnt] = sPaternVal;
      sSummaryRow = sPaternVal + " : ";
      //       sSummaryRow = "";
      colToCheck = arrSearchPatern["coltocheck"];
      retColVal = arrSearchPatern["returncol"];
      for (irow=0;irow<arrSearchRange.length;irow++) {
        getCellVal = arrSearchRange[irow][colToCheck];
        getCellExeptVal = arrSearchRange[irow][ExeptedCol];
        if (ExeptedCol > 0 && getCellExeptVal == sExeptedVal) { 
          bOnlyThatStatus = true;
        } else {
          bOnlyThatStatus = false;
        }
        if (ExeptedCol == 0) { 
          bOnlyThatStatus = true;
        }
        
        if (getCellVal == sPaternVal && bOnlyThatStatus) {
          sSoughtCellVal = arrSearchRange[irow][retColVal];
          colcnt++;
// that is a crutch - an awkward inclusion 
          sMOdate = arrSearchRange[irow][4];
          // || sPaternVal == "COL - MI"
          if ( sMOdate != "" && (sPaternVal == "Moving Out" || sPaternVal == "MO - MO") ) {
            
            var dayMO = sMOdate.getDate();
            var monthMO = sMOdate.getMonth() + 1;
            var yearMO = sMOdate.getFullYear();
            sSoughtCellVal = sSoughtCellVal + "_(" + dayMO + "/" + monthMO + "/" + yearMO +")";
          }
// that is a crutch - an awkward inclusion 
          
          sSummaryRow = sSummaryRow + sSoughtCellVal + ", ";
          arrRez[inewrowcnt][colcnt] = sSoughtCellVal;
//          if (maxColNum<colcnt) maxColNum=colcnt;
        }
      }
      
      arrRez[inewrowcnt][0] = colcnt - 2; // - 1
      arrRez[inewrowcnt][1] = sSummaryRow;
      
      for (jcol=colcnt+1;jcol<sizeOfArray-1;jcol++) {
        arrRez[inewrowcnt][jcol] = "";
      }
      
      colcnt = 2;
      inewrowcnt++;
    }
  }
  
  return arrRez;
  
}

function getSortedArray(arrToBeSorted, arrSortPatern, colToSort) {
  var arrSortedRez = [];
  var icnt, jcnt, zcnt;
  var srcTargetArrVal, srcSortPatern;
  zcnt = 0;
  
  for (icnt=0;icnt<arrSortPatern.length;icnt++) {
    srcSortPatern = arrSortPatern[icnt];
    for (jcnt=0;jcnt<arrToBeSorted.length;jcnt++) {
      srcTargetArrVal = arrToBeSorted[jcnt][colToSort];  
      if (srcSortPatern==srcTargetArrVal) {
        if (!arrSortedRez[zcnt]) arrSortedRez[zcnt] = [];
        arrSortedRez[zcnt] = arrToBeSorted[jcnt];
        zcnt++;
      }
    }
  }
  return arrSortedRez;
}

function setSummaryOfUnitsAgrByStatus() {
  //  If you call your function with a reference to a range of cells as an argument (like =DOUBLE(A1:B10)), the argument will be a two-dimensional array of the cells' values.
  var searchRange = "A1:K120";
//  var searchRange = "A1:G120";
//  var arrSearchUnitRange = {"rowstart":2, "rowend":120};
//  var arrSearchAgreementRange = {"rowstart":2, "rowend":120};
  var arrSearchPaternUnitStatus = {"rowstart":2, "rowend":11,"colstart":10, "coltocheck":1,"returncol":0};
  var arrSearchPaternAgreementStatus = {"rowstart":14, "rowend":30, "colstart":10, "coltocheck":2,"returncol":0};
  var arrSearchPaternUnitType = {"rowstart":43, "rowend":56, "colstart":10, "coltocheck":5,"returncol":0};
  var arrSearchPaternUnitTypeAvailOnly = {"rowstart":60, "rowend":73, "colstart":10, "coltocheck":5,"returncol":0};
  // instead of taken value from sheet accordint arrSearchPaternUnitStatus and arrSearchPaternAgreementStatus
  //  the lists arrSortByAgrStatPartern and arrSortByUnitsStatPartern can be used 
  var arrSortByAgrStatPartern = ["COL - NT", "COL - SU", "COL - AG", "COL - MI", "LIP - AMO", "LIP - LFR", "LIP - ITV", "LIP - LE", "LIP - LT", "MO - PreMO", "MO - MO", "MO - CO", "MO - LFR", "MO - ITV", "MO - NTV", "MO - EVCT", "EOL - BOL", "EOL - NTV/EVCT", "EOL - Dispute", "EOL - VCAT", "LIP - CT", "EOL - PT", "EOL - Archive", "EOL - Cancelled", "Offer"];
  var arrSortByUnitsStatPartern = ["Vacant", "Available", "Moving Out", "Looking for Replacement", "Lease Extension", "Occupied", "Renovation"];   
  var arrAgrStatusExceptions = ["EOL - Archive"];
  var arrUnitStatusExceptions = ["Suspended"];
  var arrUnitTypeExceptions = ["Available"];
  var ExeptedCol = 1;
  var arrSortUnitType = ["ST", "SR", "DR", "TWR", "1 BR", "2 BR", "3 BR", "4 BR", "5 BR", "6 BR", "7 BR", "8 BR"];
  // "SR", "TW", "ST","1BR","2BR","3BR","4BR","5BR","6BR","7BR","8BR"   // from RatesCard
  // "ST", "SR", "1 BR", "4 BR", "5 BR", "8 BR", "2 BR", "TWR", "3 BR", "7 BR", "6 BR", "DR" // from vTiger

  var colToSort = 2;
  var curdoc = SpreadsheetApp.getActiveSpreadsheet(); 
  var cursheet = SpreadsheetApp.getActive().getSheetByName("AgrCopy");
  
  var arrSearchRange = cursheet.getRange(searchRange).getValues();
  
  //  var getRezValArray = cursheet.getRange(arrSearchPaternUnitStatus["rowstart"] + 1, arrSearchPaternUnitStatus["colstart"] + 2, 8, 60).getValues();
  var getRezValArray;
  
  // ==== Unit Status Table
  getRezValArray = getSummaryArrayOfUnitsAgrStatus(arrSearchRange, arrSearchPaternUnitStatus, arrUnitStatusExceptions, 0); // return array
  getRezValArray = getSortedArray(getRezValArray, arrSortByUnitsStatPartern, colToSort);

//  //  cursheet.getRange(arrSearchPaternUnitStatus["rowstart"] + 1, arrSearchPaternUnitStatus["colstart"] + 2, getRezValArray.length, getRezValArray[1].length).clear();
  var rowT = arrSearchPaternUnitStatus["rowstart"]; // + 1
  var columnT = arrSearchPaternUnitStatus["colstart"] + 2; // + 2
  var numRowsT = getRezValArray.length;
  var numColumnsT = getRezValArray[1].length;
  var valuesT = getRezValArray;
  
  cursheet.getRange(rowT, columnT, 10, 55).activate().clear();
  var setVal = cursheet.getRange(rowT, columnT, numRowsT, numColumnsT).setValues(getRezValArray);
  // ==== Agreement Status Table
  getRezValArray = getSummaryArrayOfUnitsAgrStatus(arrSearchRange, arrSearchPaternAgreementStatus, arrAgrStatusExceptions, 0);
  getRezValArray = getSortedArray(getRezValArray, arrSortByAgrStatPartern, colToSort);
  //  var setVal = cursheet.getRange(arrSearchPaternAgreementStatus["rowstart"] + 1, arrSearchPaternAgreementStatus["colstart"] + 2 , getRezValArray.length, getRezValArray[1].length).clear();
  
  cursheet.getRange(arrSearchPaternAgreementStatus["rowstart"] , arrSearchPaternAgreementStatus["colstart"] + 2 , 15, 55).activate().clear();
  setVal = cursheet.getRange(arrSearchPaternAgreementStatus["rowstart"] , arrSearchPaternAgreementStatus["colstart"] + 2 , getRezValArray.length, getRezValArray[1].length).setValues(getRezValArray); 
  // ==== Units Type Table
  getRezValArray = getSummaryArrayOfUnitsAgrStatus(arrSearchRange, arrSearchPaternUnitType, arrAgrStatusExceptions, 0);
  getRezValArray = getSortedArray(getRezValArray, arrSortUnitType, colToSort);
  
  cursheet.getRange(arrSearchPaternUnitType["rowstart"] , arrSearchPaternUnitType["colstart"] + 2, 13, 55).activate().clear();
  setVal = cursheet.getRange(arrSearchPaternUnitType["rowstart"] , arrSearchPaternUnitType["colstart"] + 2 , getRezValArray.length, getRezValArray[1].length).setValues(getRezValArray); 
  
  //  // ==== Available only Status Units Type Table 
  getRezValArray = getSummaryArrayOfUnitsAgrStatus(arrSearchRange, arrSearchPaternUnitTypeAvailOnly, arrUnitTypeExceptions, ExeptedCol);
  getRezValArray = getSortedArray(getRezValArray, arrSortUnitType, colToSort);
  
  cursheet.getRange(arrSearchPaternUnitTypeAvailOnly["rowstart"] , arrSearchPaternUnitTypeAvailOnly["colstart"] + 2, 13, 55).activate().clear();
  setVal = cursheet.getRange(arrSearchPaternUnitTypeAvailOnly["rowstart"] , arrSearchPaternUnitTypeAvailOnly["colstart"] + 2 , getRezValArray.length, getRezValArray[1].length).setValues(getRezValArray); 
//  
//  Logger.log(typeof getRezValArray);
//  Logger.log("getRezValArray is: " + getRezValArray);
//  Logger.log("getRezValArray[0][0] is: " + getRezValArray[0][0]);
//  Logger.log("getRezValArray[0][1] is: " + getRezValArray[0][1]);
//  Logger.log("getRezValArray[1][1] is: " + getRezValArray[1][1]);
//  Logger.log("getRezValArray[2][2] is: " + getRezValArray[2][2]);
////  Logger.log("getRezValArray[6][0] is: " + getRezValArray[6][0]);
////  Logger.log("getRezValArray[6][1] is: " + getRezValArray[6][1]);
////  Logger.log("getRezValArray[6][58] is: " + getRezValArray[6][58]);
//  Logger.log("getRezValArray is: " + getRezValArray);
//  
  
  //  Exception: The number of columns in the data does not match the number of columns in the range. The data has 0 but the range has 119. (line 195, file "ManagersGTAds")
  //  Exception: The parameters (SpreadsheetApp.Range) don't match the method signature for SpreadsheetApp.Range.setValues. (line 189, file "ManagersGTAds")
  //  Cannot convert Array to number[][]. (line 189, file "ManagersGTAds")
  // the reason was that there was no [0] index : final aray getRezValArray[][] was defined as getRezValArray[1-8][0-59]  and getRezValArray[12-22][0-59] , i.e. 

}



function toR1C1str(reference) {
  var range = SpreadsheetApp.getActiveSheet().getRange(reference);
  var row = range.getRow();
  var column = range.getColumn();
  var start = 'R' + row + 'C' + column;
  var rows = range.getNumRows();
  var columns = range.getNumColumns();
  var end = ((rows * columns) == 1)?'':':R' + (row + rows - 1) + 'C' + (column + columns - 1);
  return start + end;
}

function toR1C1(reference) {
  var R1C1str = toR1C1str(reference);
  var row = parseInt(R1C1str.substring(1, R1C1str.indexOf("C")));
  var col = parseInt(R1C1str.substring(R1C1str.indexOf("C")+1));
  return [row, col];
}


function pad(num, size) {
    var s = "000000000" + num;
    return s.substr(s.length-size);
}

function facebook_clear_all_backgrounds() {
  var fbsheet = SpreadsheetApp.getActive().getSheetByName('Facebook Advertising Schedule');
  fbsheet.getRange(FB_RANGE[0] + ":" + FB_RANGE[1]).setBackground("");
}

function fill_fb_cell() {
  var fbsheet = SpreadsheetApp.getActive().getSheetByName('Facebook Advertising Schedule');
  var cell_range = toR1C1(fbsheet.getActiveRange().getA1Notation())
  var offsets = toR1C1(FB_RANGE[0]);
  var lastRC =  toR1C1(FB_RANGE[1]);
  //Logger.log([cell_range, offsets, lastRC]);
  if(cell_range[0] >= offsets[0] && cell_range[0] <= lastRC[0] &&
     cell_range[1] >= offsets[1] && cell_range[1] <= lastRC[1]) {
    var user_name = fbsheet.getRange("P1").getValue();
    var previous_str = fbsheet.getActiveRange().getValue();
    var new_now = new Date;
    var now = new Date(new_now.getTime());
    var day = now.getDate();
    var month = now.getMonth() + 1;
    if(previous_str.length == 0) {
      fbsheet.getActiveRange().setValue(user_name + " " + day + "." + month);
    } else {
      fbsheet.getActiveRange().setValue(previous_str + ", " + user_name + " " + day + "." + month);
    }
  }
}

function facebook_date_colorizer() {
  // Facebook Advertising Schedule
  // I:AS 5:107
  // J1 - days ago
  var offsets = toR1C1(FB_RANGE[0]);
  var lastRC =  toR1C1(FB_RANGE[1]);
  var row_offset = offsets[0];
  var col_offset = offsets[1];
  var lastRCrow_offset = lastRC[0];
  var lastRCcol_offset = lastRC[1];
  var new_now = new Date;
  var now = new Date(new_now.getTime());
  var fbsheet = SpreadsheetApp.getActive().getSheetByName('Facebook Advertising Schedule');
    
  var days_ago1 = parseInt(fbsheet.getRange("J1").getValue());
  var days_ago2 = parseInt(fbsheet.getRange("J2").getValue());
  var days_ago3 = parseInt(fbsheet.getRange("M1").getValue());
  var days_ago4 = parseInt(fbsheet.getRange("M2").getValue());
  var bg1 = fbsheet.getRange("J1").getBackground();
  var bg2 = fbsheet.getRange("J2").getBackground();
  var bg3 = fbsheet.getRange("M1").getBackground();
  var bg4 = fbsheet.getRange("M2").getBackground();
  //Logger.log([days_ago1, days_ago2, days_ago3]);
  var fb_sheet_data = fbsheet.getRange(row_offset, col_offset, lastRCrow_offset-row_offset+1, lastRCcol_offset-col_offset+1).getValues();
  //Logger.log(fb_sheet_data);

  for(var r=0; r<fb_sheet_data.length; ++r) {
    for(var c=0; c<fb_sheet_data[r].length; ++c) {
      if(fb_sheet_data[r][c] != "") {
        var strdate = fb_sheet_data[r][c].toString().substring(1 + fb_sheet_data[r][c].toString().lastIndexOf(" "));
        var cellday   = pad(parseInt(strdate.toString().substring(0, strdate.indexOf("."))), 2);
        var cellmonth = pad(parseInt(strdate.toString().substring(strdate.indexOf(".")+1)), 2);
        var celldate = new Date(now.getYear().toString() + "-" + cellmonth + "-" + cellday);
        var diff_in_days = Math.floor((now - celldate)/(24*3600*1000));
        //Logger.log([fb_sheet_data[r][c], r, c, strdate, celldate, diff_in_days]);
        if       (diff_in_days >= days_ago1 && diff_in_days < days_ago2) {
          fbsheet.getRange(row_offset + r, col_offset + c).setBackground(bg1);
        } else if(diff_in_days >= days_ago2 && diff_in_days < days_ago3) {
          fbsheet.getRange(row_offset + r, col_offset + c).setBackground(bg2);
        } else if(diff_in_days >= days_ago3 && diff_in_days < days_ago4) {
          fbsheet.getRange(row_offset + r, col_offset + c).setBackground(bg3);
        } else if(diff_in_days >= days_ago4) {
          fbsheet.getRange(row_offset + r, col_offset + c).setBackground(bg4);
        }
      }
    }
  }
}

function add_dinamic_stats() {
  // UnitsQntDynStats; // "UnitsDynamicStats";  
  // "UnitsPricesDynStats"; // "DynamicStats";
  // "1xGulUuODWgcO0rq4FCelGMFGmYZ4W-8QTwITBgty5m0" // "List of Properties and Descriptions"
  
  var srcCol = 1; // 1 - update Units; 2 - update Prices
  var targetsheetName = "UnitsQntDynStats"; // "UnitsDynamicStats";
  updateUnitsPricesDynamicStats(targetsheetName, srcCol);
  
  var srcCol = 2; // 1 - update Units; 2 - update Prices  
  var targetsheetName = "UnitsPricesDynStats"; // "DynamicStats";
  updateUnitsPricesDynamicStats(targetsheetName, srcCol);
  
//  Exception: Unexpected error while getting the method or property removeChart on object SpreadsheetApp.Sheet.
  
  addstat();
  update_allnotes();
  facebook_date_colorizer();
  setSummaryOfUnitsAgrByStatus();
  setDateForMO();
  updateRosterAgrUnitStatus();
}

function updateUnitsPricesDynamicStats(targetsheetName, srcCol) {

  var sourcesheetName = 'Calculations'; 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourcesheet = ss.getSheetByName(sourcesheetName); 
  var targetsheet = ss.getSheetByName(targetsheetName);
  
  var allcurrentdata = sourcesheet.getRange("Y2:AC14").getValues();
  
  //  var currdate_row = 12;
  var dates = targetsheet.getRange("A2:A").getValues();
  var new_now = new Date;
  var now = new Date(new_now.getTime());
  
  for(var row=0; row<dates.length; ++row) {
    if(dates[row][0].getDate() == now.getDate() && dates[row][0].getMonth() == now.getMonth() && dates[row][0].getYear() == now.getYear()) {
//    if(dates[row][0].getDate() == now.getDate()) {
      // set the active row to the current date position
      break;
    }
  }
  var sheetrow = (row + 2).toString();
  var col_names = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O"];  
  var varAU = '';
  
  for(var cnt=0; cnt<8; ++cnt) {
    varAU = allcurrentdata[cnt][srcCol];
    // possible errors "#ERROR!" , "#NUM!" , "#NAME?" , "#N/A"
    if (varAU.toString().indexOf("#") < 0) {
      targetsheet.getRange(col_names[cnt+1] + sheetrow).setValue(varAU);
    }
  }
  
  SetChartsRange_HideRows(targetsheetName, row+2, 13, 32, "A", "L");
}


function addstat() {
  // "#ERROR!" , "#NUM!" , "#NAME?"
  // sometime it gets either ERRORS or "0" . Trace it. if such behaviour will repeat probably need to change 
  // 1 - trigger set to " every 6 or 8 hours" and ad the condition checker to that function 
  // ( if cell.value is not a number than run the function or check the value of source cells)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('All Current'); // .getSheets()[0];

  
  var allcurrentdata = sheet.getRange("E1:E2").getValues();
  var available_rate = allcurrentdata[0][0];
  var mo_rate = allcurrentdata[1][0];
  
//  var currdate_row = 12;
  var dates = SpreadsheetApp.getActive().getSheetByName("AvailStats").getRange("A2:A").getValues();
  var new_now = new Date;
  var now = new Date(new_now.getTime());
  for(var row=0; row<dates.length; ++row) {
    if(dates[row][0].getDate() == now.getDate() && dates[row][0].getMonth() == now.getMonth() && dates[row][0].getYear() == now.getYear()) {
      
      break;
    }
  }
  var sheetrow = (row + 2).toString();
//  Logger.log("available_rate is: " + available_rate + " mo_rate " + mo_rate);
//  Logger.log("available_rate is: " + available_rate.toString().indexOf("#") );
//  Logger.log("mo_rate is: " + mo_rate.toString().indexOf("#") );
  if (available_rate.toString().indexOf("#") < 0 && mo_rate.toString().indexOf("#") < 0) {
    SpreadsheetApp.getActive().getSheetByName("AvailStats").getRange("B" + sheetrow).setValue(available_rate);
    SpreadsheetApp.getActive().getSheetByName("AvailStats").getRange("C" + sheetrow).setValue(mo_rate);
  }
  SetChartsRange_HideRows("AvailStats", row+2, 5, 35, "A", "D");
}

function SetChartsRange_HideRows(sname, curdaterow, colpos, period, rangestart, rangeend) {
//  Logger.log("curdaterow = " + curdaterow.toString() + "  period = " + period.toString() ); 
  var spreadsheet = SpreadsheetApp.getActive().getSheetByName(sname); //.activate(); // .getRange("A2:A").getValues();
  //spreadsheet.getRange('A2').activate();
//  var sheet = spreadsheet.getActiveSheet();
  var charts = spreadsheet.getCharts();
  var chart = charts[charts.length - 1];
  var startrow = curdaterow - period;
//  if (curdaterow > period) {
//    startrow = curdaterow - period; // Number(curdaterow), parseInt("234",10)
//  } 

  var endrow = curdaterow  ;
  
  var rangeToHide = spreadsheet.getRange( '2:' + startrow.toString()); //.activate();
//  spreadsheet.getRange( 'A1:A' + startrow.toString()).activate();
//    spreadsheet.getRange( startrow.toString() + ':' + endrow.toString()).activate();

  spreadsheet.hideRows(rangeToHide.getRow(), rangeToHide.getNumRows());
//  spreadsheet.hideRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
//  spreadsheet.getRange(startrow.toString() + ':' + startrow.toString() ); // .activate();
  var rangeToChart =spreadsheet.getRange(startrow.toString() + ':' + startrow.toString() ); // .activate();
  var col2name = spreadsheet.getRange('B1');
  
  //  
  try {
    spreadsheet.removeChart(chart);
  } catch(err) {
//    Logger.log("err is: " + err);
    
//    err is: Exception: Unexpected error while getting the method or property removeChart on object SpreadsheetApp.Sheet.
//    if (!rezToArray[0]) rezToArray[0] = [];
//    rezToArray[0][0] = "#empty#";
////    Logger.log("rezToArray[0][0] is: " + rezToArray[0][0]);
  } 
  
  chart = spreadsheet.newChart();
  chart.asLineChart();
  chart.addRange(spreadsheet.getRange(rangestart + startrow.toString() + ":" + rangeend + endrow.toString()));
  chart.setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS);
  chart.setTransposeRowsAndColumns(false);
  chart.setNumHeaders(1);
  chart.setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH);
  chart.setOption('bubble.stroke', '#000000');
  chart.setOption('useFirstColumnAsDomain', true);
  chart.setOption('domainAxis.direction', 1);
//  .setOption('series.0.labelInLegend', '$11,692')
// ......
//  .setOption('series.6.labelInLegend', '$0')
  chart.setOption('height', 471);
  chart.setOption('width', 756);
  chart.setPosition(startrow, colpos, 32, 0);
//  chart.build();
//  
  var headerdata = spreadsheet.getRange(rangestart +  "1:" + rangeend + "1").getValues();
  var headerSize = headerdata[0].length;
//  Logger.log("headerSize = " + headerSize.toString()); 
  var varColName = '';
  var varSeriesName = '';
  
  for(var cnt=1; cnt<headerSize; ++cnt) {
    varColName = headerdata[0][cnt];
    varSeriesName = 'series.' + (cnt - 1).toString() + '.labelInLegend';
    chart.setOption(varSeriesName, varColName);
  }
  
  spreadsheet.insertChart(chart.build());
}

function rrunit(pc) {
  var pc_found = false;
  var prev_pc = "";
  var prev_rate = 0;
  var people;
  people_live = [0, 0, 0]; 
  for(var r=5; r<sheets_data[RC].length; ++r) {
    if(sheets_data[RC][r][0] == pc) {
      pc_found = true;
      var rate = sheets_data[RC][r][1];
      if(typeof rate != "number") rate = 0;
      
      if(typeof sheets_data[RC][r][2] == "number") {
        people = sheets_data[RC][r][2];
        if(sheets_data[RC][r][0] == prev_pc) {
          if(people_live[0] > 0)
            people_live = [people, rate, prev_rate]; 
        } else {
          people_live = [people, rate, prev_rate]; 
        }
      } else {
        people = 0;
        people_live = [0, rate, rate];
      }
      
      prev_rate = rate;
      prev_pc   = pc;
    }
    
    if(pc_found && sheets_data[RC][r][0] != pc ) {
      break;
    }
  }
  
  total_rate = people_live[1] * people_live[0];
  if(people_live[0] == 0)
    total_rate = people_live[1];
  if(people_live[0] == 3)
    total_rate = people_live[2] * 2;  // 2x rate for 2 people
  
  //Logger.log(total_rate)
  return total_rate;
}

// Return Rate
// replacement for $" & round(INDEX(RatesCopy!B:B, match(INDEX(A:A, ROW(), 1), RatesCopy!A:A,0)+1)) & " per person per week based on 2 persons sharing
function rr(people) {  
  if (people == undefined || people == 0)
    var rc_data_search_people = "UNIT";
  else
    var rc_data_search_people = people;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var pc_row = ss.getActiveSheet().getActiveRange().getRow();
  var pc = s.getRange(pc_row, 1).getValue();
  var rate = "";
  
  for(var r=0; r<sheets_data[RC].length; ++r) {
    if(sheets_data[RC][r][0] == pc) {
      if(typeof sheets_data[RC][r][2] == "number")
        sheets_data[RC][r][2] = parseInt(sheets_data[RC][r][2]);

      if(sheets_data[RC][r][2] == rc_data_search_people) {
        rate = Math.round(sheets_data[RC][r][1]);
        break;
      }
    }
  }
  return rate;
}

function calcAvailableAndMOTotalValue() {
//  sheet_calc = SpreadsheetApp.getActive().getSheetByName("Calculations");
  
  people_live = {};
  var prev_pc = "";
  var prev_rate = 0;
  
  // Status:
  // Other:     0
  // Available: 1
  // MO:        2  
  
  //   people rate prev_rate Status
  // [ 0      1    2         3      ]
  for(var r=5; r<sheets_data[RC].length; ++r) {
    var rate = sheets_data[RC][r][1];
    var status = 0;
    // TypeError: sheets_data[RC][r][5].toLowerCase is not a function (line 239, file "GTAds")
    // .toLowerCase function only exists on strings. You can call toString() on anything in javascript to get a string representation. Putting this all together: var temp = ans.toString().toLowerCase();
    //if(sheets_data[RC][r][5].toString().toLowerCase() == "available")
    if(sheets_data[RC][r][5] == "Available" || sheets_data[RC][r][5] == "Vacant")
      status = 1;
    //if(sheets_data[RC][r][5].toString().toLowerCase() == "moving out")
    if(sheets_data[RC][r][5] == "Moving Out")
      status = 2;
      
    if(typeof sheets_data[RC][r][2] == "number") {
      
      people = sheets_data[RC][r][2];
      if(sheets_data[RC][r][0] == prev_pc) {
        if(people_live[prev_pc][0] > 0)
          people_live[sheets_data[RC][r][0]] = [people, rate, prev_rate, status]; 
      } else {
        people_live[sheets_data[RC][r][0]] = [people, rate, prev_rate, status]; 
      }
    } else {
      people = 0;
      people_live[sheets_data[RC][r][0]] = [0, rate, rate, status];
    }

    prev_rate = rate;
    prev_pc   = sheets_data[RC][r][0];
  }
    
  //sheet_calc.getRange("B:D").clear();
  //var row_counter = 1;
  var sum_rates_of_total_available = 0;
  var sum_rates_of_total_mo = 0;
  for(pc in people_live) {
    total_rate = people_live[pc][1] * people_live[pc][0];
//    Logger.log(["total_rate = " + total_rate]);
    if(people_live[pc][0] == 0)
      total_rate = people_live[pc][1];
    if(people_live[pc][0] == 3)
      total_rate = people_live[pc][2] * 2;  // 2x rate for 2 people
   
    if(people_live[pc][3] == 1)
      sum_rates_of_total_available += total_rate;
    if(people_live[pc][3] == 2)
      sum_rates_of_total_mo += total_rate;
    //sheet_calc.getRange("B" + row_counter.toString()).setValue(total_rate);
    //sheet_calc.getRange("C" + row_counter.toString()).setValue(people_live[pc][0]);
    //sheet_calc.getRange("D" + row_counter.toString()).setValue(people_live[pc][3]);
    //row_counter++;
  }
//  Logger.log(["sum_rates_of_total_available = " + sum_rates_of_total_available]);
//  Logger.log(["sum_rates_of_total_mo = " + sum_rates_of_total_mo]);
//  //var sum_rates_availble_and_mo = sum_rates_of_total_available + sum_rates_of_total_mo;
//  //sheet_calc.getRange("E1:G1").setValues([[sum_rates_of_total_available, sum_rates_of_total_mo, sum_rates_availble_and_mo]]);
//  //Logger.log([sum_rates_of_total_available, sum_rates_of_total_mo]);
  return [sum_rates_of_total_available, sum_rates_of_total_mo];
}

function generate_notes(sheet_row_number, sheet_id) {
  var sheets1 = [];
  var sheets_data1 = [];
  sheets1.push(SpreadsheetApp.getActive().getSheetByName('Unit Information'));
  sheets_data1.push(sheets1[0].getDataRange().getValues());
//}
  
//  var sheet_pc = sheets_data1[sheet_id][sheet_row_number][COLUMN_PC];
  var sheet_pc = sheets_data1[0][sheet_row_number][COLUMN_PC];
  var str_val = "";
  if(sheet_pc.length > 0) {
    var max_people = 0;
    var min_people = 100000;
    
    // column is:
    // 0  1    2
    // PC Rate People/Unit
    var sheet_pc_founded = false;
    var rates = {};
    for(var r=0; r<sheets_data[RC].length; ++r) {
      if(sheets_data[RC][r][0] == sheet_pc) {
        sheet_pc_founded = true;
        if(typeof sheets_data[RC][r][2] == "number")
          sheets_data[RC][r][2] = parseInt(sheets_data[RC][r][2]);
        rates[sheets_data[RC][r][2]] = Math.round(sheets_data[RC][r][1]);
        
        if (sheets_data[RC][r][2] < min_people)
          min_people = sheets_data[RC][r][2];
        
        if (sheets_data[RC][r][2] > max_people)
          max_people = sheets_data[RC][r][2];
      } else if(sheet_pc_founded === true) {
        break;
      }
    }
    
    if (sheet_pc_founded === true) {
      if(min_people != max_people)
        str_val += "The house is suitable for a group of " + min_people + " to " + max_people + " persons sharing:\n";
      for(var p=max_people; p>=min_people; --p)
        str_val += "$" + rates[p] + " per person per week based on " + p + " persons sharing\n";
      
      if (sheet_id == SL) {
        var bills = sheets_data[sheet_id][sheet_row_number][COLUMN_SL_BILLS];
        if(bills == "Water")
          str_val += "Price includes water.\n";
        if(bills == "All")
          str_val += "Pricing includes all bills – with nothing more to pay!\n";
      } else {
        str_val += "Pricing includes all bills – with nothing more to pay!\n";
      }
    } 
    //Logger.log(str_val);
  }
  return str_val;
}

// Returns property rates strings for people
function pr() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_id = sheet_names.indexOf(ss.getActiveSheet().getName());
  var pc_row = ss.getActiveSheet().getActiveRange().getRow();
  
  var pc = ss.getActiveSheet().getRange(pc_row, 1).getValue();

  Logger.log("ss.getActiveSheet().getName() " + ss.getActiveSheet().getName());  
  Logger.log("pc_row " + pc_row);
  Logger.log("sheet_id " + sheet_id);
  
  
  str_val = "\n" + generate_notes(pc_row-1, sheet_id);
  str_val += "\nFlexible lease terms and simple application process.\nTo confirm availability, pricing and to arrange a personal inspection contact us today: 0421 127 291 - Easy Living Melbourne.\n" +
             "When enquiring please provide your best phone contact number and personal email address so that we can contact you to answer your enquiry directly.\n" +
             "Please note that pricing is subject to change on a daily basis.\n" +
             "Hope to speak to you soon and welcome the opportunity of providing you with a home.\n\n" + "PC: " + pc;
  
  return str_val;
}

function update_allnotes() {
  for(var sheet_id=0; sheet_id < RC; ++sheet_id) {
    var notes_arr = [];
    for(var sheet_row_number=1; sheet_row_number<sheets_data[sheet_id].length; ++sheet_row_number)
      notes_arr.push(generate_notes(sheet_row_number, sheet_id));
  
    for(var r=0; r < parseInt(sheets_data[sheet_id].length) - ROW_START + 1; ++r) {
      var celln = DESC[sheet_id] + (r + ROW_START).toString();
      Logger.log(celln);
      if((notes_arr[r]).length > 0)
        sheets[sheet_id].getRange(celln).setNote(notes_arr[r]);
    }
  }

  var today = new Date();
  str_notes_updated = "Notes updated: " + Utilities.formatDate(today, "Australia/Melbourne", "dd MMMM HH:mm:ss")
  for(var sheet_id=0; sheet_id < RC; ++sheet_id) {
    sheets[sheet_id].getRange(DESC[sheet_id] + '1').setNote(str_notes_updated);
  }
}


//function put_ads_on_gumtree() {
//  Browser.msgBox("TEST")
//}

function onOpen(e){
  var menuItems = [{name: "Update Notes Desc", functionName: "update_allnotes"}, {name: "Facebook Data Colorizer", functionName: "facebook_date_colorizer"}, {name: "Update Availability Statistics", functionName: "addstat"}, {name: "->Update Dynamic Stats", functionName: "add_dinamic_stats"}];
  SpreadsheetApp.getActive().addMenu("ELM Menu", menuItems);
  update_allnotes();
  addstat();
  facebook_date_colorizer();
}

//function copyValidations(sheetName, rowFrom, rowTo, column, numColumns) {
// /* Copy validations of all cells in rowFrom to cells in rowTo, starting at column.\*/
//
//  var sh = SpreadsheetApp.getActiveSpreadsheet();
//  var ss = sh.getSheetByName(sheetName);
//
//  for (var i = column;i<= column + numColumns;i++) {
//    var cellFrom = ss.getRange(rowFrom, i);
//    var cellTo = ss.getRange(rowTo, i);
//    copyCellDataValidation(cellFrom, cellTo); 
//  }
//}


//function onEdit(e){
//}

//function onOpen(e){
//  //var checkTabColor=setSheetTabColor();
//  // Add a custom menu to the active spreadsheet, including a separator and a sub-menu.
//  SpreadsheetApp.getUi()
//  .createMenu('<ELM>')
//  .addItem('Create new as_at_for document (according Settings page)', 'create_as_at_for')
//  .addSeparator()
////  .addItem('Create MYOB-file', '--importDataForMYOB_and_MoveToArchive')
//  .addSeparator()
//  //.addItem('Move schedules to archive', 'moveFilesToArchive')
//  .addSeparator()
//  .addSubMenu(SpreadsheetApp.getUi().createMenu('Miscellaneous...')
////              .addItem('Put color code into colored celles', 'putValueToColoredCell')
////              .addItem('Clear weeks sheets', 'ClearRange')
//              .addItem('Set tabs color', 'setSheetTabColor')
//              .addSeparator()
////              .addItem('test---save/download file', 'doGet')              
//             )
//  .addToUi();
//  }


//function add_dinamic_stats0_obsolete() {
//  // sometime it gets either ERRORS or "0" . Trace it. if such behaviour will repeat probably need to change 
//  // 1 - trigger set to " every 6 or 8 hours" and ad the condition checker to that function 
//  // ( if cell.value is not a number than run the function or check the value of source cells)
//  
//   // "#ERROR!" , "#NUM!" , "#NAME?" , "#N/A"
//  
//  var sourcesheetName = 'Calculations'; 
//  
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var sourcesheet = ss.getSheetByName(sourcesheetName); 
//  var targetsheet = ss.getSheetByName(targetsheetName);
//  
//  var allcurrentdata = sourcesheet.getRange("W2:AA13").getValues();
//  
//  //  var currdate_row = 12;
//  var dates = targetsheet.getRange("A2:A").getValues();
//  var new_now = new Date;
//  var now = new Date(new_now.getTime());
//  
//  for(var row=0; row<dates.length; ++row) {
//    if(dates[row][0].getDate() == now.getDate() && dates[row][0].getMonth() == now.getMonth() && dates[row][0].getYear() == now.getYear()) {
////    if(dates[row][0].getDate() == now.getDate()) {
//      // set the active row to the current date position
//      break;
//    }
//  }
////  Logger.log("row is: " + row);
//  
//  var sheetrow = (row + 2).toString();
//  var col_names = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O"];  
//  //B var available_rate = allcurrentdata[0][0] = 'Available       allcurrentdata[0][2] = AU$
//  //C     allcurrentdata[1][0] = 'Moving Out                       allcurrentdata[1][2] = AU$
//  //D     allcurrentdata[2][0] = 'Occupied  
//  //E     allcurrentdata[3][0] = 'Lease Extension
//  //F var mo_rate = allcurrentdata[4][0] = 'Looking for Replacement
//  //G     allcurrentdata[5][0] = 'Renovation
//  //H     allcurrentdata[6][0] = '
//  //I     allcurrentdata[7][0] = '  
//    
//  //  Logger.log("row is: " + row);
//  var varAU = '';
//  
//  for(var cnt=0; cnt<8; ++cnt) {
//    varAU = allcurrentdata[cnt][2];
//    // possible errors "#ERROR!" , "#NUM!" , "#NAME?" , "#N/A"
//    if (varAU.toString().indexOf("#") < 0) {
//      targetsheet.getRange(col_names[cnt+1] + sheetrow).setValue(varAU);
//    }
//  }
//  
//  SetChartsRange_HideRows(targetsheetName, row+2, 12, 32, "A", "K");
//}


