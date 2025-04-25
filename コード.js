let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var activeSheet = spreadsheet.getSheetByName('申込み用紙');
let scriptSheet = spreadsheet.getSheetByName('スクリプト');
let dataExtractionSheet = spreadsheet.getSheetByName('データ抽出');

function onOpen() {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('データ取得', [
      {name: 'ソースコード作成', functionName: 'generateMaleDataSourceCode'},
    ]);
}

function generateMaleDataSourceCode() {
  var dataValues = new Array(1000);
  var occurrenceCount = new Array(1000);
  for (var i = 0; i < 100; i++) {
    dataValues[i] = "";
    occurrenceCount[i] = 0;
  }
  
  var lastRowDataSheet = dataExtractionSheet.getLastRow();
  var lastColDataSheet = dataExtractionSheet.getLastColumn();
  var headerValues = dataExtractionSheet.getRange(1, 1, 1, lastColDataSheet).getValues();
  
  scriptSheet.clear();
  clearDataExtractionSheet();
  
  const lastRowActiveSheet = activeSheet.getLastRow();
  const lastColActiveSheet = activeSheet.getLastColumn();
  
  let categoryIdentifier = "x";
  let rowIndex = 1;
  
  const values = activeSheet.getRange(1, 1, lastRowActiveSheet, lastColActiveSheet).getValues();
  
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      let cellValue = values[i][j].toString();
      let firstCharacter = cellValue.slice(0, 1);
      
      if (cellValue !== "" && firstCharacter == categoryIdentifier) {
        let itemNumber = Number(cellValue.slice(1));
        
        occurrenceCount[itemNumber - 1] += 1;
        
        // 値の位置情報を記録
        if (occurrenceCount[itemNumber - 1] == 1) {
          // 最初の出現時は単純に位置を記録
          dataValues[itemNumber - 1] = {
            row: i,
            col: j
          };
        } else if (occurrenceCount[itemNumber - 1] == 2) {
          // 2回目の出現時は配列に変換して両方の位置を保存
          let firstOccurrence = dataValues[itemNumber - 1];
          dataValues[itemNumber - 1] = [
            { row: firstOccurrence.row, col: firstOccurrence.col },
            { row: i, col: j }
          ];
        } else if (occurrenceCount[itemNumber - 1] > 2) {
          // 3回目以降の出現時は配列に追加
          dataValues[itemNumber - 1].push({ row: i, col: j });
        }
        
        let headerLabel = headerValues[0][itemNumber - 1];
        
        // 条件付き出力形式の構築
        let outputCode = "";
        
        if (occurrenceCount[itemNumber - 1] == 1) {
          // 1つの値のみの場合
          outputCode = `values01[${dataValues[itemNumber - 1].row}][${dataValues[itemNumber - 1].col}]`;
        } else if (occurrenceCount[itemNumber - 1] == 2) {
          // 2つの値がある場合、条件付き連結
          let pos1 = dataValues[itemNumber - 1][0];
          let pos2 = dataValues[itemNumber - 1][1];
          outputCode = `if(values01[${pos1.row}][${pos1.col}] != ""){values01[${pos1.row}][${pos1.col}] + " " + values01[${pos2.row}][${pos2.col}]}`;
        } else {
          // 3つ以上の値がある場合、最初の2つだけ使用
          let pos1 = dataValues[itemNumber - 1][0];
          let pos2 = dataValues[itemNumber - 1][1];
          outputCode = `if(values01[${pos1.row}][${pos1.col}] != ""){values01[${pos1.row}][${pos1.col}] + " " + values01[${pos2.row}][${pos2.col}]}`;
        }
        
        scriptSheet.getRange(itemNumber, 1).setValue(
          `data1[r${rowIndex}][${itemNumber - 1}] = ${outputCode}// ${headerLabel}`
        );
      }
    }
  }
}

function clearDataExtractionSheet() {
  var lastRowDataSheet = dataExtractionSheet.getLastRow();
  var lastColDataSheet = dataExtractionSheet.getLastColumn();
  
  if (lastRowDataSheet > 1) {
    dataExtractionSheet.getRange(2, 1, lastRowDataSheet - 1, lastColDataSheet).clearContent();
  }
}
