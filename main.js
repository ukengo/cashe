function myFunction() {
  const spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A2820').activate();
}

function insert() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Аркуш1 (копия)')
  //const ui = sp.getUi();
  //const response = ui.alert('Выбрана последняя запись?', ui.ButtonSet.YES_NO);
  // if (response == ui.Button.YES) {
  const rangeA1 = sheet.getRange(1, 1)
  let a1 = rangeA1.getValue()
  sheet.getRange(1, 3).setValue(a1)
  rangeA1.setValue(a1 - 2 + (sheet.getActiveRange().getRow()))
  // }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Меню')
    .addItem('Insert', 'insert')
    .addItem('Insert Rows', 'insertRows')
    .addToUi();
  let cache = CacheService.getScriptCache();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let row = cache.get("row");
  let column = cache.get("column");
  sheet.getRange(+row, +column).activate();
  SpreadsheetApp.flush();
}

function onEdit(e) {
  const ss = e.source
  const range = e.range
  const evalue = e.value
  const row = range.getRow();
  const column = range.getColumn();
  let cache = CacheService.getScriptCache();
  cache.putAll({
    "row": `${row}`,
    "column": `${column}`
  })
  const sheetMelnik = ss.getSheetByName('Аркуш1 (копия)')
  const sheetFriedman = ss.getSheetByName('Лист1')
  const sheetName = ss.getActiveSheet().getName()
  const valueA1 = sheetMelnik.getRange(1, 1).getValue()
  const rangeCellEdit = sheetFriedman.getRange(valueA1 + row - 2, column)
  if (sheetName == 'Аркуш1 (копия)' && e.range.getRow() != 1 && e.range.getRow() != 2) {
    rangeCellEdit.setValue(evalue)
    range.clearContent()
  }  
}

function insertRows() {
  const ss = SpreadsheetApp.getActive()
  const sheetMelnik = ss.getSheetByName('Аркуш1 (копия)')
  const sheetFriedman = ss.getSheetByName('Лист1')
  const sheetName = ss.getActiveSheet().getName()
  const valueA1 = sheetMelnik.getRange(1, 1).getValue()
  const row = sheetMelnik.getActiveRange().getRow()
   const ui = SpreadsheetApp.getUi()
  if (sheetName == 'Аркуш1 (копия)') {
   
    const response = ui.prompt(`Укажите количество строк`)
   if (response.getSelectedButton() == ui.Button.OK) {
      rowSum = response.getResponseText();

      sheetFriedman.insertRows(valueA1 + row - 1, rowSum)
    }
 }
}

function forUprav() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const list1 = ss.getSheetByName('Лист1');
  const list1Copy = ss.getSheetByName('forUprav');
  list1Copy.getDataRange().clear()
  const arr = list1.getRange('A1:F').getValues();
  const len = arr.map(x => x[0]).indexOf('PRHisAGc');
  let data = list1.getRange('A1:F' + len).getValues();
  for (let i = 0; i < data.length - 1; i++) {
    if (data[i + 1][0] == 0 && data[i][0]) {
      data[i + 1][0] = data[i][0];
    }
  }
  list1Copy.getRange(1, 1, data.length, data[0].length).setValues(data); 
}


/*
function onEdit(e){
  let row = e.range.getRow();
  let column = e.range.getColumn();

  let cache = CacheService.getScriptCache();
  cache.putAll({
    "row": `${row}`,
    "column": `${column}`
  });
}

function onOpen() {
  let cache = CacheService.getScriptCache();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let row = cache.get("row");
  let column = cache.get("column");
  sheet.getRange(+row, +column).activate();
  SpreadsheetApp.flush();

} */