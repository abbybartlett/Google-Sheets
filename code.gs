/* 
Arrange Template script ID for library: 	
1J3Ln44TEfS3qRURHBj6QONHMs44n7xtExMenY_gExdyvoDvZsu2TFD2L
*/

var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
var sheet = SpreadsheetApp.getActiveSheet();
var sheetId = sheet.getSheetId();

function onOpen(){
  arrange.onOpen();
};

function arrangeUsageReport(){
  arrange.arrangeUsageReport();
};

function sortColumn1(){
  arrange.sortColumn1();
};

function sortColumn2(){
  arrange.sortColumn2();
};

function copyAppDb() {
  arrange.copyAppDb(ssId);
};

function copyAppStoreInfo() {
  arrange.copyAppStoreInfo(ssId);
};

function getColors(){
  arrange.getColors();
};

function showRows() {
  arrange.showRows();
};

function showCollege(){
  arrange.showCollege();
};

function deleteEmptyRows() {
  arrange.deleteEmptyRows()
};
