function myMySQLFetchData() { 

  var ssNameToInsert = 'Availability';
  var rowInsertStart = 3;
  var colInsertStart = 1;
  
  var conn = Jdbc.getConnection('jdbc:mysql://xxx.xxx.xxx.xxx:3306/dbname', { 
    user: 'username', 
    password: 'password'
  });
  
  var stmt = conn.createStatement();
  var start = new Date(); // Get script starting time
  
  var sql_from_part = ' FROM vtiger_units INNER JOIN vtiger_properties ON vtiger_units.propertyid = vtiger_properties.propertyid  INNER JOIN vtiger_crmentity ON vtiger_units.unitid = vtiger_crmentity.crmid INNER JOIN vtiger_unitscf ON vtiger_units.unitid = vtiger_unitscf.unitid  WHERE vtiger_crmentity.deleted=0 AND   (  (( vtiger_units.unit_type <> "NONE" AND vtiger_units.unit_type IS NOT NULL AND vtiger_units.unit_type <> "Administrative") ))  ';
  var sql_rows_num = 'SELECT count(*) as cnt ' + sql_from_part;

  var sql_u = 'SELECT vtiger_units.code as "Code", vtiger_units.unit_status as "Status", vtiger_unitscf.cf_2832 as "Last Agreement Status" , vtiger_unitscf.cf_2834 as "Last EOL date", vtiger_unitscf.cf_929 as "Actual Moving out Date", vtiger_units.unit_type as "Type", vtiger_properties.name as "Related to Property" ' + sql_from_part + ' ORDER BY modifiedtime DESC';

  var rs1 = stmt.executeQuery(sql_rows_num);
  rs1.absolute(1);// .next();
  var rowcnt = rs1.getString('cnt');

  var rs = stmt.executeQuery(sql_u);
      
// Google Sheets Details       
  var doc = SpreadsheetApp.getActiveSpreadsheet(); // Returns the currently active spreadsheet
//  
  var dstsheet = SpreadsheetApp.getActive().getSheetByName(ssNameToInsert);
//      var cell = dstsheet.getRange('a6');
  
  dstsheet.getRange('A2:G120').activate();
  dstsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  dstsheet.getRange('A1').activate();
  
  
  var row = 0;
  var getCount = rs.getMetaData().getColumnCount(); // Mysql table column name count.
  var colcnt = rs.getMetaData().getColumnCount();

  var row = 0; 
  var rezToArray = []; // = new Array[rowcnt];
  
  while (rs.next()) {
    for (var col = 0; col < colcnt; col++) { 
      if (!rezToArray[row]) rezToArray[row] = [];
      rezToArray[row][col] = rs.getString(col + 1);
    }
    row++;
  }

  dstsheet.getRange(rowInsertStart, colInsertStart, rezToArray.length , rezToArray[0].length).setValues(rezToArray);

  rs1.close();
  rs.close();
  stmt.close();
  conn.close();
  var end = new Date(); // Get script ending time
//  Logger.log('start of export = ' + start.getTime() + ' ; end of export = ' + end.getTime() + ' ; start.getMonth = ' + start.getMonth() + ' ; end = ' + end ); 
  var timeElapsed = (end.getTime() - start.getTime()) / 1000;
  
  var lastExecDate = "Updated on " + end.getDate() + "/" + (end.getMonth() +1 ) +  "/" + end.getFullYear() + " at " + end.getHours() + ":" + end.getMinutes() + ":" + end.getSeconds();
  
  dstsheet.getRange("K1").setValue(lastExecDate);
  dstsheet.getRange("K2").setValue("It's taken " + timeElapsed  + 'sec.');
  
  SpreadsheetApp.flush();
} 

function onOpen(e){
  var menuItems = [{name: "Update Availability from vTiger", functionName: "myMySQLFetchData"}];
  SpreadsheetApp.getActive().addMenu("ELM Menu", menuItems);
  myMySQLFetchData();
  
}

