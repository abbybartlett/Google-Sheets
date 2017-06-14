function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Script Menu')
      .addItem('Format Usage Report', 'arrangeUsageReport')
      .addSubMenu(ui.createMenu('Step-by-Step Formatting')
        .addItem('Group by OS', 'groupByOS')
        .addItem('Remove empty columns', 'deleteEmptyColumns')
        .addItem('Set column headers', 'setHeader')
        .addItem('Copy Col A & C' , 'copyAppInfo')
        .addItem('Set Operating System', 'setOS')
        .addItem('Move iOS info below Android', 'moveAppleRange')
        .addItem('sort Col C, then Col A', 'sort'))
      .addSeparator()
      .addItem('Sort Columns B, A, C', 'sortColumn')
      .addToUi();
}
  
var sheet = SpreadsheetApp.getActiveSheet();

// added 06-13-2017 to sort columns on any sheet
function sortColumn() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getActiveSheet();

  sheet.sort(2);
  sheet.sort(1);
  sheet.sort(3, false);
}
/* -------------------------------------------------------
    everything below this point is used to rearrange and 
    organize the weekly Usage Report from the server 
-------------------------------------------------------- */

function groupByOS() {
  sheet.getRange('D1:D').moveTo(sheet.getRange('K1'));
  sheet.getRange('F1:F').moveTo(sheet.getRange('L1'));
  sheet.getRange('H1:H').moveTo(sheet.getRange('M1'));
  sheet.getRange('J1:J').moveTo(sheet.getRange('N1'));
  sheet.insertColumnsAfter(10, 3);
};

function deleteEmptyColumns() {
  sheet.deleteColumn(4);
  sheet.deleteColumn(5);
  sheet.deleteColumn(6);
  sheet.insertColumnAfter(1);
  sheet.setFrozenRows(1);
};

function setHeader() {
  var headerValues = [
    ["App Name", "Operating System", "Channel ID", "Users", "Registered Users", "Promos Redeemed", "Churn"]
  ];
  var range = sheet.getRange('A1:G1');
  range.setValues(headerValues);
};

function copyAppInfo() {
  var rangeToCopy = sheet.getRange('A:A');
  rangeToCopy.copyTo(sheet.getRange(1, 9));
  var rangeToCopyNext = sheet.getRange('C:C');
  rangeToCopyNext.copyTo(sheet.getRange(1, 11));
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
  var range = sheet.getRange(2, 10, numRows - 1);
  range.setValue(valuesI);
};

function moveAppleRange() {
  var numRows = sheet.getDataRange().getNumRows();
  var lastRow = sheet.getLastRow();
  var rangeToCopy = sheet.getRange(2, 9, numRows, 7);
  rangeToCopy.copyTo(sheet.getRange(lastRow + 1, 1));
  
  var range = sheet.getRange('I1:O');
  range.clearContent();
};

function sort() {
  sheet.sort(3);
  sheet.sort(1);
};

function arrangeUsageReport() {
  groupByOS();
  deleteEmptyColumns();
  setHeader();
  copyAppInfo();
  setOS();
  moveAppleRange();
  sort();
};
