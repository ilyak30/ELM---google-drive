// Globals

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
  fbsheet.getRange(FB_RANGE[0] + ":" + FB_RANGE[1]).setBackground("");
}

function fill_fb_cell() {
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
        var strdate = fb_sheet_data[r][c].substring(1 + fb_sheet_data[r][c].lastIndexOf(" "));
        var cellday   = pad(parseInt(strdate.substring(0, strdate.indexOf("."))), 2);
        var cellmonth = pad(parseInt(strdate.substring(strdate.indexOf(".")+1)), 2);
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

function addstat() {
  // sometime it gets either ERRORS or "0" . Trace it. if such behaviour will repeat probably need to change 
  // 1 - trigger set to " every 6 or 8 hours" and ad the condition checker to that function 
  // ( if cell.value is not a number than run the function or check the value of source cells)
  var allcurrentdata = SpreadsheetApp.getActive().getSheetByName("All Current").getRange("E1:E2").getValues();
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
//  Logger.log("row is: " + row);
  var sheetrow = (row + 2).toString();
  SpreadsheetApp.getActive().getSheetByName("AvailStats").getRange("B" + sheetrow).setValue(available_rate);
  SpreadsheetApp.getActive().getSheetByName("AvailStats").getRange("C" + sheetrow).setValue(mo_rate);
  
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

function put_ads_on_gumtree() {
  Browser.msgBox("TEST")
}

function onOpen(e){
  var menuItems = [{name: "Update Notes Desc", functionName: "update_allnotes"}];
  SpreadsheetApp.getActive().addMenu("ELM Menu", menuItems);
  update_allnotes();
//  addstat();
  facebook_date_colorizer();
}
//function onEdit(e){
//}
