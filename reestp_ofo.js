function InsertRowCache(data) {
  const url = ('https://docs.google.com/spreadsheets/d/1Y6ukb_iQ0JpPt2nTGvZx0HjWMJ2ACWD0iQDlV9Rjr7s/edit#gid=658828820');
  const sheet = 'Лист1';
  const key = 'PRHisAGc';
  const colA = 'A';
  const colB = 'B';
  const ss = SpreadsheetApp.openByUrl(url);
  const spredSheet = ss.getSheetByName(sheet);
  if(formatDateDDtireMMtireYYYY((preIndex(spredSheet, key, colA)[1][0])) != formatDateDDtireMMtireYYYY(new Date())) {
    spredSheet.insertRows((preIndex(spredSheet, key, colB)[0])+1, 2);
    spredSheet.getRange((preIndex(spredSheet, key, colB)[0])+2,1,1,1).setValue(new Date());
  } 
  spredSheet.insertRows((preIndex(spredSheet, key, colA)[0])+1);
  spredSheet.getRange((preIndex(spredSheet, key, colA)[0])+1,2,1,4).setValues(data);  
  const quantBlankRow = cellIndex(spredSheet, key)-preIndex(spredSheet, key, colB)[0]-1
  console.log(quantBlankRow);
  if(quantBlankRow > 3) {
    spredSheet.deleteRows(preIndex(spredSheet, key, colB)[0]+1, quantBlankRow-3);
  }
}


// ищет индекс предпоследней заполненной ячейки и ее значение в указанной колонке. Колонка указывается буквой
function preIndex(spredSheet, key, col) {
  const index = cellIndex(spredSheet, key);
  const dataArr = spredSheet.getRange(col+'1:'+col+''+index).getValues();
  for(i = dataArr.length-2; i>=0; i--) {
    if(dataArr[i] != '') {
      return [i+1,dataArr[i]]
    }
  }
}

// ищет номер строки указанной ячейки 
function cellIndex(spredSheet, key) {
  const dataArr = spredSheet.getRange('A1:A').getValues();
  return dataArr.flat().indexOf(key)+1;
}

//форматирование даты в вид dd-mm-гггг 
function formatDateDDtireMMtireYYYY(date){
    return ((date.getDate()) < 10 ? '0' : '') + (date.getDate()) + '-'
          + ((date.getMonth() + 1) < 10 ? '0' : '') + (date.getMonth() + 1) + '-'
          +  date.getFullYear(); 
}













/* function preIndex() {
  const url = ('https://docs.google.com/spreadsheets/d/1Y6ukb_iQ0JpPt2nTGvZx0HjWMJ2ACWD0iQDlV9Rjr7s/edit#gid=658828820');
  const sheet = 'Лист1';
  const key = 'PRHisAGc';
  const col = 'A';
  const ss = SpreadsheetApp.openByUrl(url);
  const spredSheet = ss.getSheetByName(sheet);
  const index = cellIndex(spredSheet, key);
  console.log(col+'1:'+col+''+index)
 // const dataArr = spredSheet.getRange(col+'1:'+col+''+index).getValues();
  /* for(i = dataArr.length-2; i>=0; i--) {
    if(dataArr[i] != '') {
      return [i+1,dataArr[i]]
    }
  } 
} */