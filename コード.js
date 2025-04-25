let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var activeSheet = activeSpreadsheet.getActiveSheet();
let scriptSheet = activeSpreadsheet.getSheetByName('スクリプト');
let dataExtractionSheet = activeSpreadsheet.getSheetByName('データ抽出');

function onOpen() {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('データ取得', [
      {name: 'ソースコード作成', functionName: 'generateSourceCodeForMaleData'},
    ]);
}

function generateSourceCodeForMaleData() {
  var dataValues = new Array(1000);
  var dataFoundFlags = new Array(1000);
  for (var i = 0; i < 100 ; i++) {
    dataValues[i] = "";
    dataFoundFlags[i] = 0;
  }
  
  var dataSheetLastRow, dataSheetLastColumn;
  dataSheetLastRow = dataExtractionSheet.getLastRow();
  dataSheetLastColumn = dataExtractionSheet.getLastColumn();
  dataSheetValues = dataExtractionSheet.getRange(1, 1, 1, dataSheetLastColumn).getValues();
  
  scriptSheet.clear();
  clearDataExtractionSheet();
  const activeSheetLastRow = activeSheet.getLastRow();
  const activeSheetLastColumn = activeSheet.getLastColumn();
  
  let genderCode = "x";
  let rowNumber = 1;
  
  const activeSheetValues = activeSheet.getRange(1, 1, activeSheetLastRow, activeSheetLastColumn).getValues();
  for (var i = 0; i < activeSheetValues.length; i++) {
    for (var j = 0; j < activeSheetValues[i].length; j++) {
      let cellText = activeSheetValues[i][j].toString();
      let firstCharacter = cellText.slice(0, 1);
      
      if(cellText !== "" && firstCharacter == genderCode){
        let itemNumber = Number(cellText.slice(1));
        dataFoundFlags[itemNumber - 1] += 1;
        if(dataFoundFlags[itemNumber - 1] == 1){
          dataValues[itemNumber - 1] += `activeSheetValues[${i}][${j}]`;
        }else{
          dataValues[itemNumber - 1] += ` + " " + activeSheetValues[${i}][${j}]`;
        }
        let columnTitle = dataSheetValues[0][itemNumber - 1];
        scriptSheet.getRange(itemNumber, 1).setValue(`data1[r${rowNumber}][${itemNumber-1}] = ${dataValues[itemNumber - 1]}// ${columnTitle}`);
      }
    }
  }
}


function clearDataExtractionSheet() {
  var dataSheetLastRow, dataSheetLastColumn;
  dataSheetLastRow = dataExtractionSheet.getLastRow();
  dataSheetLastColumn = dataExtractionSheet.getLastColumn();
  if (dataSheetLastRow > 1){
    dataExtractionSheet.getRange(2, 1, dataSheetLastRow - 1, dataSheetLastColumn).clearContent();
  }
}
