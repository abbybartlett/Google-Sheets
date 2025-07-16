// this script ID: 	1J3Ln44TEfS3qRURHBj6QONHMs44n7xtExMenY_gExdyvoDvZsu2TFD2L (lives at https://script.google.com/home/projects/1J3Ln44TEfS3qRURHBj6QONHMs44n7xtExMenY_gExdyvoDvZsu2TFD2L/edit) 
// spreadsheet Template using this script: Arrange - https://script.google.com/d/MChwD2qGnIg6jZgfYQTuf3TZfUhHPiFYJ/edit?mid=ACjPJvGvddkwDFqlitzluHfCr69CABaN7l6ovA35Rn-vuXz-pG0NAulE3Yv8GJ9EzINfEnPOjVWPxfogPhKmL3WX8N8TgWMTg3w7xIEQH_FFVhhb7RoyA9nf0r6vTvHpUcOxU41F5lwNZRc&uiv=2 
// ^ redirects to script bound to sheet "zz_UsageReport TEMPLATE": https://docs.google.com/spreadsheets/d/1rpIvHN0vKI8_cYb7CxQDcAFij170csIinrfzoI2gLr0/edit?gid=263469272#gid=263469272 

var ss = SpreadsheetApp.getActiveSpreadsheet();
var ssId = ss.getId();
var sheet = ss.getActiveSheet();
var sheetId = sheet.getSheetId();
var reportSheet = ss.getSheetByName('report');
var maxRows = sheet.getMaxRows();
var lastRow = sheet.getLastRow();

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Script Menu')
      .addItem('1. Format Usage Report', 'arrangeUsageReport')
      /*.addSubMenu(ui.createMenu('Step-by-Step Formatting')
        .addItem('Group by OS', 'groupByOS')
        .addItem('Remove empty columns', 'deleteEmptyColumns')
        .addItem('Set column headers', 'setHeader')
        .addItem('Copy Col A & C' , 'copyAppInfo')
        .addItem('Set Operating System', 'setOS')
        .addItem('Move iOS info below Android', 'moveAppleRange')
        .addItem('sort Col C, then Col A', 'sort'))
      */
      .addItem('2. Copy AppDB Reporting Sheet', 'copyAppDb')
      .addSeparator()
      .addItem('Sort for App Downloads Report', 'sortColumn1')
      .addItem('Sort for All Teams Report', 'sortColumn2')
      .addSeparator()
      .addSubMenu(ui.createMenu('More options...')
/*      .addItem('Show College', 'showCollege')
        .addItem('College apps', 'showCollege')
        .addItem('All Live apps', 'showLive')
        .addItem('ATR apps', 'showATR')
        .addItem('Show all rows', 'showRows')
        .addSeparator()
        .addItem('Copy/Paste values on report', 'setReportValues')
        .addItem('Format Report', 'formatReport')
        .addSeparator()
        */
        .addItem('Copy AppDB App Store Sheet','copyAppStoreInfo'))
      .addToUi();
      
  //reportSheet.showRows(1, maxRows);
  //copyAppDb();
  getColors();
}

function getColors() {
  var colors = {
    "royal blue": "#1362BB",
    "white": "#ffffff",
    "charcoal":"#474747",
    "black" : "#000000",
    "lightGray" : "#D8D8D8",
    "green" : "#069D72",
    "aqua" : "#17D9E5",
    "yellow" : "#ffd100",
    "rose" : "#ff4d4d",
    "tangerine" : "#FF822A"
    }
  var properties1 = PropertiesService.getScriptProperties();
  properties1.setProperties(colors);
  return properties1.getProperties();
};


// added 06-13-2017 to sort columns on any sheet
function sortColumn1() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getActiveSheet();

  sheet.sort(2);
  sheet.sort(4);
  sheet.sort(5, false);
  sheet.sort(6);
};
function sortColumn2() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getActiveSheet();

  sheet.sort(2);
  sheet.sort(4);
  sheet.sort(7, false);
};

// added 11-28-2017 to remove the issue of having to Allow Access to AppDB spreadsheet to get keys
function copyAppDb(ssId) {
  //appDB sheet info
  var ss = SpreadsheetApp.openById("1YeliZXviL0jweAyGtfnXhF-OCuYXhRo40HzD85UKwOM");
  var dbSheet = ss.getSheetByName("Reporting");
  var numRows = dbSheet.getMaxRows();
  var numColumns = dbSheet.getLastColumn();
  var data = dbSheet.getSheetValues(1, 1, numRows, numColumns);
  
  //this sheet's info, where the data will be pasted
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  var thisSs = SpreadsheetApp.openById(ssId);
  var thisSheet = thisSs.getSheetByName("appDB");
  //var thisSheet = thisSs.getSheetByName("copy of Reporting");
  if(!thisSheet){
    thisSs.insertSheet("appDB",1);
    thisSs.getSheetByName("appDB").getRange("A1").setValue("Analytics Reporting Info");
    thisSs.getSheetByName("appDB").setTabColor("black");
    };
  //dbSheet.copyTo(thisSs);
  
  var dataRange = thisSheet.getRange(1, 1, numRows, numColumns);
  
  dataRange.setValues(data);
  thisSheet.setColumnWidth(1,220);
  thisSheet.setFrozenColumns(2);
  thisSheet.setFrozenRows(1);
  thisSheet.setName("appDB");
};

// added 12-11-2017 to copy App Store info
function copyAppStoreInfo(ssId) {
  //appDB sheet info
  var ss = SpreadsheetApp.openById("1YeliZXviL0jweAyGtfnXhF-OCuYXhRo40HzD85UKwOM");
  var dbSheet = ss.getSheetByName("AppStore");
  var numRows = dbSheet.getMaxRows();
  var numColumns = dbSheet.getLastColumn();
  var data = dbSheet.getSheetValues(1, 1, numRows, numColumns);
  
  //this sheet's info, where the data will be pasted
  var thisSs = SpreadsheetApp.openById(ssId);  
  var thisSheet = thisSs.getSheetByName("appStoreDB");
  if(!thisSheet){
    thisSs.insertSheet("appStoreDB",1);
    thisSs.getSheetByName("appStoreDB").getRange("A1").setValue("App Store Info");
    thisSs.getSheetByName("appStoreDB").setTabColor("black");
    };
  var dataRange = thisSheet.getRange(1, 1, numRows, numColumns);
  dataRange.setValues(data);
  thisSheet.setColumnWidth(1,220);
  thisSheet.setFrozenColumns(2);
  thisSheet.setFrozenRows(1);
};

/* -------------------------------------------------------
    everything below this point is used to rearrange and 
    organize the weekly Usage Report from the server 
-------------------------------------------------------- */

function groupByOS() {
  sheet.getRange('D1:D').moveTo(sheet.getRange('AA1'));
  sheet.getRange('F1:F').moveTo(sheet.getRange('AB1'));
  sheet.getRange('H1:H').moveTo(sheet.getRange('AC1'));
  sheet.getRange('J1:J').moveTo(sheet.getRange('AD1'));
  sheet.getRange('L1:L').moveTo(sheet.getRange('AE1'));
  sheet.getRange('N1:N').moveTo(sheet.getRange('AF1'));
  sheet.getRange('P1:P').moveTo(sheet.getRange('AG1'));
  sheet.getRange('R1:R').moveTo(sheet.getRange('AH1'));
  sheet.getRange('T1:T').moveTo(sheet.getRange('AI1'));
  sheet.getRange('V1:V').moveTo(sheet.getRange('AJ1'));  
};

function deleteEmptyColumns() {
  sheet.deleteColumn(4);
  sheet.deleteColumn(5);
  sheet.deleteColumn(6);
  sheet.deleteColumn(7);
  sheet.deleteColumn(8);
  sheet.deleteColumn(9);
  sheet.deleteColumn(10);
  sheet.deleteColumn(11);
  sheet.deleteColumn(12);
  sheet.deleteColumn(13);
  sheet.insertColumnAfter(1);
  sheet.setFrozenRows(1);
  
};

function setHeader() {
  var headerValues = [
    ["App Name", "Operating System", "Channel ID", "Users", "Registered Users", "Promos Redeemed", "Churn", 'NonConverted', 'Registered-NonImported', 'Users-NonImported', 'Imported Users', 'Converted Users', 'Unique Devices']
  ];
  var range = sheet.getRange('A1:M1');
  range.setValues(headerValues);
};

function copyAppInfo() {
  var rangeToCopy = sheet.getRange('A:A');
  rangeToCopy.copyTo(sheet.getRange(1, 15));
  var rangeToCopyNext = sheet.getRange('C:C');
  rangeToCopyNext.copyTo(sheet.getRange(1, 17));
};

function setOS() {
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var valuesA = [];
  var valuesI = [];
  
  for (var i = 0; i <= numRows - 1; i++) {
    var valuesA = ["Android"];
    var valuesI = ["iOS"];
  }
  
  var range = sheet.getRange(2, 2, numRows - 1);
  range.setValue(valuesA);
  var range = sheet.getRange(2, 16, numRows - 1);
  range.setValue(valuesI);
};

function moveAppleRange() {
  var numRows = sheet.getDataRange().getNumRows();
  var lastRow = sheet.getLastRow();
  var rangeToCopy = sheet.getRange(2, 15, numRows, 13);
  rangeToCopy.copyTo(sheet.getRange(lastRow + 1, 1));
  
  var range = sheet.getRange('N1:AA');
  range.clearContent();
};

function sort() {
  sheet.sort(3);
  sheet.sort(1);
};

function deleteEmptyRows() {
  var lastRow = sheet.getLastRow();
  var maxRows = sheet.getMaxRows();
  sheet.deleteRows(lastRow+1, maxRows-lastRow)
};
  
function arrangeUsageReport() {
  groupByOS();
  deleteEmptyColumns();
  setHeader();
  copyAppInfo();
  setOS();
  moveAppleRange();
  sort();
  deleteEmptyRows();
};


//     MORE OPTIONS section added 12/19/2017 
//
// *** FORMULAS ADDED *** showCollege, showLive, showATR, setReportValues, formatReport
//
//     @showCollege == hide all rows where column "Type" ≠ "College"
//     @showLive == hide all rows where column "Status" ≠ "1. Live"
//     @showATR == hide all rows where column "ATR?" ≠ "1. Live"
//     @showRows == unhide all rows
//     @setReportValues == copy/paste values for columns A:C & H:L on "report" sheet
//     @formatReport == apply conditional formatting to rows on "report" sheet

/*function showCollege() {
  Logger.log(maxRows);
  Logger.log(lastRow);
  var data = reportSheet.getRange('E:E').getValues();
  
  for(var i=0; i<lastRow+1; i++){
      Logger.log(data[i][0].toString());
    if(data[i][0].toString() !== 'College'){
      reportSheet.hideRows(i+1);
    }
  }
  //reportSheet.unhideRow(1);
};
  
function showRows() {
  sheet.showRows(1, maxRows);
}
*/
