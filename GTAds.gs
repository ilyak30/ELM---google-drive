// Globals
// =VLOOKUP(A55,RatesCopy!A:G,6,0)
// =INDEX(RatesCopy!F:F, MATCH(A4, RatesCopy!A:A, 0), 1)
// =COUNTIF(C5:C, "Available")
// ='All Current'!A5
// ='All Current'!B5

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

function fillRatesCopySheet() {
  // PC	        A + E1  =importrange("1WUU4pvanWH5_iSUY2hPyXLfrTyR22MAHkDhRxSl5KTg", B2 & B1 & "!" & "A" & E1 & ":" & "A")
  // Rate	    E2+ E1  =importrange("1WUU4pvanWH5_iSUY2hPyXLfrTyR22MAHkDhRxSl5KTg", B2 & B1 & "!" &  E2 & E1 & ":" & E2)
  // People	    I + E1  =importrange("1WUU4pvanWH5_iSUY2hPyXLfrTyR22MAHkDhRxSl5KTg", B2 & B1 & "!" & "I" & E1 & ":" & "I")
  // Actual MO	E + E1  =importrange("1WUU4pvanWH5_iSUY2hPyXLfrTyR22MAHkDhRxSl5KTg", B2 & B1 & "!" & "E" & E1 & ":" & "E")
  // EOL Date	F + E1  =importrange("1WUU4pvanWH5_iSUY2hPyXLfrTyR22MAHkDhRxSl5KTg", B2 & B1 & "!" & "F" & E1 & ":" & "F")
  // Unit	    H + E1  =importrange("1WUU4pvanWH5_iSUY2hPyXLfrTyR22MAHkDhRxSl5KTg", B2 & B1 & "!" & "H" & E1 & ":" & "H")
  // Type       B + E1  =importrange("1WUU4pvanWH5_iSUY2hPyXLfrTyR22MAHkDhRxSl5KTg", B2 & B1 & "!" & "B" & E1 & ":" & "B")
  // 						
// 1A96G	470	   1		       18/06/2020	Occupied	1BR
//  =importrange("https://docs.google.com/spreadsheets/d/1WUU4pvanWH5_iSUY2hPyXLfrTyR22MAHkDhRxSl5KTg", B2 & B1 & "!" & "A" & E1 & ":" & "A")
// Year= B1	20		    StartRow=E1	3
// Month=B2	April		CyanColumn=E2	P
  var endSourceRow = 300;
  var endSourceCol = 7;   // G
  
  var startTargetRow = 5;
  var startTargetCol = 1; // A
  var endTargetCol = 7;   // G
  var numRows = endTargetRow - startTargetRow + 1;
  var numColumns = endTargetCol - startTargetCol + 1;
  
  var sourceSSdocID = '1WUU4pvanWH5_iSUY2hPyXLfrTyR22MAHkDhRxSl5KTg';
  var sourceSheetName = ''; 
  var targetSheetName = "RatesCopy";
  
  var ssTarget = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ssTarget.getSheetByName(targetSheetName);
  var targetRangeSourceInfo = targetSheet.getRange("A1:F2").getValues(); 
  
  sourceSheetName = targetRangeSourceInfo[1][1].toString() + targetRangeSourceInfo[0][1].toString(); "April"+"20"
  var startSourceCol = [0]; // A
  
  var startSourceRow = targetRangeSourceInfo[0][4];  // "3"
  var endTargetRow = endSourceRow + startTargetRow - startSourceRow;
  
  var sourceOfRatesCol = targetRangeSourceInfo[1][4].toString(); // "P"
  var startSourceCol = ["A",sourceOfRatesCol,"I","E","F","H","B"]; // A
  var arrSizeSourceCol = startSourceCol.length;
  
  
  var ssSource = SpreadsheetApp.openById(sourceSSdocID); // getActiveSpreadsheet();
  var sourceSheet = ssSource.getSheetByName(sourceSheetName); 
//  var sourceSheet = ssSource.getSheetByName("April20"); 
  Logger.log("sourceSheetName is: " + sourceSheetName); // April20
  Logger.log("targetStartRow is: " + startTargetRow);    // 5
  Logger.log("sourceOfRatesCol is: " + sourceOfRatesCol);  // "P"
  Logger.log("arrSizeSourceCol is: " + arrSizeSourceCol);  // 
  
  //  var sourcPC = sourceSheet.getRange(startSourceRow, startSourceCol, numRows, 1).getValues(); // A + E1
  //  var sourcRange = sourceSheet.getRange(startSourceRow, startSourceCol, numRows, numColumns).getValues();
  
  for (var i = 0;i< arrSizeSourceCol;i++) {
    var sColName = startSourceCol[i];
     Logger.log("sColName is: " + sColName); 
    for (var irow = 0;irow<= 1;irow++) {
      
      targetsheet.getRange(sColName + sheetrow).setValue(allcurrentdata[cnt][2]);
      var sourcRange = sourceSheet.getRange("A5:A30", "C5:C30", "F5:F30").getValues();   // , "C5:C30", "F5:F30"
    }
  }
  Logger.log("sourcRange is: [2][1]" + sourcRange[0][1]);
  
//  var targetRangeToInsert = targetSheet.getRange(startTargetRow, startTargetCol, numRows, numColumns).setValues(sourcRange); 

}

function copyValidations(sheetName, rowFrom, rowTo, column, numColumns) {
 /* Copy validations of all cells in rowFrom to cells in rowTo, starting at column.\*/

  var sh = SpreadsheetApp.getActiveSpreadsheet();
  var ss = sh.getSheetByName(sheetName);

  for (var i = column;i<= column + numColumns;i++) {
    var cellFrom = ss.getRange(rowFrom, i);
    var cellTo = ss.getRange(rowTo, i);
    copyCellDataValidation(cellFrom, cellTo); 
  }
}

function add_dinamic_stats() {
  // sometime it gets either ERRORS or "0" . Trace it. if such behaviour will repeat probably need to change 
  // 1 - trigger set to " every 6 or 8 hours" and ad the condition checker to that function 
  // ( if cell.value is not a number than run the function or check the value of source cells)
  
   // "#ERROR!" , "#NUM!" , "#NAME?" , "#N/A"
  
  var sourcesheetName = 'Calculations'; 
  var targetsheetName = "DynamicStats";
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourcesheet = ss.getSheetByName(sourcesheetName); 
  var targetsheet = ss.getSheetByName(targetsheetName);
  
  var allcurrentdata = sourcesheet.getRange("W2:AA13").getValues();
  
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
//  Logger.log("row is: " + row);
  
  var sheetrow = (row + 2).toString();
  var col_names = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O"];  
  //B var available_rate = allcurrentdata[0][0] = 'Available       allcurrentdata[0][2] = AU$
  //C     allcurrentdata[1][0] = 'Moving Out                       allcurrentdata[1][2] = AU$
  //D     allcurrentdata[2][0] = 'Occupied  
  //E     allcurrentdata[3][0] = 'Lease Extension
  //F var mo_rate = allcurrentdata[4][0] = 'Looking for Replacement
  //G     allcurrentdata[5][0] = 'Renovation
  //H     allcurrentdata[6][0] = '
  //I     allcurrentdata[7][0] = '  
    
  //  Logger.log("row is: " + row);
  var varAU = '';
  
  for(var cnt=0; cnt<8; ++cnt) {
    varAU = allcurrentdata[cnt][2];
    // possible errors "#ERROR!" , "#NUM!" , "#NAME?" , "#N/A"
    if (varAU.toString().indexOf("#") < 0) {
      targetsheet.getRange(col_names[cnt+1] + sheetrow).setValue(varAU);
    }
  }
  
  SetChartsRange_HideRows(targetsheetName, row+2, 12, 32, "A", "K");
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
  
//   Logger.log("startrow is: " + startrow);
  
  var rangeToHide = spreadsheet.getRange( '2:' + startrow.toString()); //.activate();
//  spreadsheet.getRange( 'A1:A' + startrow.toString()).activate();
//    spreadsheet.getRange( startrow.toString() + ':' + endrow.toString()).activate();

  spreadsheet.hideRows(rangeToHide.getRow(), rangeToHide.getNumRows());
//  spreadsheet.hideRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
//  spreadsheet.getRange(startrow.toString() + ':' + startrow.toString() ); // .activate();
  var rangeToChart =spreadsheet.getRange(startrow.toString() + ':' + startrow.toString() ); // .activate();
  var col2name = spreadsheet.getRange('B1');
  
//  spreadsheet.removeChart(chart);  
//  chart = spreadsheet.newChart()
//  .asLineChart()
//  .addRange(spreadsheet.getRange(rangestart + startrow.toString() + ":" + rangeend + endrow.toString()))
//  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
//  .setTransposeRowsAndColumns(false)
//  .setNumHeaders(1)
//  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
//  .setOption('bubble.stroke', '#000000')
//  .setOption('useFirstColumnAsDomain', true)
//  
//    
//    .setOption('domainAxis.direction', 1)
//
////  .setOption('series.0.labelInLegend', '$11,692')
////  .setOption('series.1.labelInLegend', '$1,026')
////  .setOption('series.2.labelInLegend', '$19,523')
////  .setOption('series.3.labelInLegend', '$1,786')
////  .setOption('series.4.labelInLegend', '$4,810')
////  .setOption('series.5.labelInLegend', '$314')
////  .setOption('series.6.labelInLegend', '$0')
//  
//  .setOption('height', 471)
//  .setOption('width', 756)
//  .setPosition(startrow, colpos, 32, 0)
//  .build();
  

  spreadsheet.removeChart(chart);
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
//  .setOption('series.1.labelInLegend', '$1,026')
//  .setOption('series.2.labelInLegend', '$19,523')
//  .setOption('series.3.labelInLegend', '$1,786')
//  .setOption('series.4.labelInLegend', '$4,810')
//  .setOption('series.5.labelInLegend', '$314')
//  .setOption('series.6.labelInLegend', '$0')
  chart.setOption('height', 471);
  chart.setOption('width', 756);
  chart.setPosition(startrow, colpos, 32, 0);
//  chart.build();
//  
  var headerdata = spreadsheet.getRange(rangestart +  "1:" + rangeend + "1").getValues();
  var headerSize = headerdata[0].length;
  Logger.log("headerSize = " + headerSize.toString()); 
  var varColName = '';
  var varSeriesName = '';
  
  for(var cnt=1; cnt<headerSize; ++cnt) {
    varColName = headerdata[0][cnt];
    varSeriesName = 'series.' + (cnt - 1).toString() + '.labelInLegend';
    chart.setOption(varSeriesName, varColName);
  }
  
//  //chart.build();
  
  spreadsheet.insertChart(chart.build());
//    spreadsheet.insertChart(chart);

}

function rrunit(pc) {
  var pc_found = false;
  var prev_pc = "";
  var prev_rate = 0;
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
  sheet_calc = SpreadsheetApp.getActive().getSheetByName("Calculations");
  
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
    rate = sheets_data[RC][r][1];
    status = 0;
    // TypeError: sheets_data[RC][r][5].toLowerCase is not a function (line 239, file "GTAds")
    // .toLowerCase function only exists on strings. You can call toString() on anything in javascript to get a string representation. Putting this all together: var temp = ans.toString().toLowerCase();
    //if(sheets_data[RC][r][5].toString().toLowerCase() == "available")
    if(sheets_data[RC][r][5] == "Available")
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
  //var sum_rates_availble_and_mo = sum_rates_of_total_available + sum_rates_of_total_mo;
  //sheet_calc.getRange("E1:G1").setValues([[sum_rates_of_total_available, sum_rates_of_total_mo, sum_rates_availble_and_mo]]);
  //Logger.log([sum_rates_of_total_available, sum_rates_of_total_mo]);
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

function createNewRateSheet() {
//  1. Rates/Add/Availabiltiy/Stats
//  https://docs.google.com/spreadsheets/d/1WUU4pvanWH5_iSUY2hPyXLfrTyR22MAHkDhRxSl5KTg/edit#gid=515246521
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('N1').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('UnitStats'), true);
  spreadsheet.duplicateActiveSheet();
  spreadsheet.getActiveSheet().setName('new name of UnitStats');
}


function put_ads_on_gumtree() {
  Browser.msgBox("TEST")
}

function onOpen(e){
  var menuItems = [{name: "Update Notes Desc", functionName: "update_allnotes"}, {name: "Facebook Data Colorizer", functionName: "facebook_date_colorizer"}, {name: "Update Availability Statistics", functionName: "addstat"}, {name: "Update Dinamic Stats", functionName: "add_dinamic_stats"}];
  SpreadsheetApp.getActive().addMenu("ELM Menu", menuItems);
  update_allnotes();
  addstat();
  facebook_date_colorizer();
}
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
