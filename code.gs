function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('NumberTable')
      .setTitle('Number Table');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function getTableSize(dRow, dCol) {
  initialize(parseInt(dRow, 10), parseInt(dCol, 10));
}

function backupCommand() {
  var a = 15;
  initialize(a, a);
}

function initialize(dRow, dCol) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  sheet.clear();
  sheet.deleteColumns(2, sheet.getMaxColumns() - 1);
  sheet.deleteRows(2, sheet.getMaxRows() - 1);
  
  while(sheet.getMaxColumns() < dCol) {
    sheet.insertColumns(sheet.getMaxColumns());
  }
  
  while(sheet.getMaxRows() < dRow) {
    sheet.insertRows(sheet.getMaxRows());
  }
  
  for(var a = 1; a <= sheet.getMaxColumns(); a++) {
    sheet.setColumnWidth(a, 40);
  }
  
  for(var a = 1; a <= sheet.getMaxRows(); a++) {
    sheet.setRowHeight(a, 21);
  }
  
  var range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  
  range.getCell(1, 1).setValue("1");
  range.setHorizontalAlignment("center");
  range.setVerticalAlignment("middle");
  
  for(var a = 2; a <= sheet.getMaxRows(); a++) {
    range.getCell(a, 1).setFormula("=R[-1]C[0]+1");
  }
  
  for(var a = 2; a <= sheet.getMaxColumns(); a++) {
    range.getCell(1, a).setFormula("=R[0]C[-1]+1");
  }
  
  for(var a = 2; a <= sheet.getMaxRows(); a++) {
    for(var b = 2; b <= sheet.getMaxColumns(); b++) {
      range.getCell(a, b).setFormula("=R[-" + (a - 1) + "]C[0]*R[0]C[-" + (b - 1) + "]");
    }
  }
  
  var rangeVals = range.getValues();
  
  for(var a = 1; (a - 1) < rangeVals.length; a++) {
    for(var b = 1; (b - 1) < rangeVals[a - 1].length; b++) {
      if(isPrime(rangeVals[a - 1][b - 1])) {
        range.getCell(a, b).setBackgroundRGB(242, 211, 111).setNote(range.getCell(a, b).getNote() + "Prime number\n");
      }
      else if(Math.pow(Math.round(Math.sqrt(rangeVals[a - 1][b - 1])), 2) === rangeVals[a - 1][b - 1]) {
        range.getCell(a, b).setBackgroundRGB(200, 255, 142).setNote(range.getCell(a, b).getNote() + "Perfect square - Root: " + Math.sqrt(rangeVals[a - 1][b - 1]) + "\n");
      }
      else {
        range.getCell(a, b).setBackgroundRGB(198, 198, 198);
      }
      
      if(rangeVals[a - 1][b - 1] % 2 === 0) {
        range.getCell(a, b).setFontColor("Purple")/*.setNote(range.getCell(a, b).getNote() + "Even number\n")*/;
      }
      else {
        range.getCell(a, b).setFontColor("Red")/*.setNote(range.getCell(a, b).getNote() + "Odd number\n")*/;
      }
    }
  }
  
  
  
  for(var a = 5; a <= sheet.getMaxRows(); a+=5) {
    sheet.getRange(a, 1, 1, sheet.getMaxColumns()).setBorder(true, true, true, true, null, null);
  }
  
  for(var a = 5; a <= sheet.getMaxColumns(); a+=5) {
    sheet.getRange(1, a, sheet.getMaxRows(), 1).setBorder(true, true, true, true, null, null);
  }
  
  for(var a = 5; a <= sheet.getMaxRows(); a+=5) {
    for(var b = 5; b <= sheet.getMaxColumns(); b+=5) {
      sheet.getRange(a, b).setBorder(false, false, false, false, null, null);
    }
  }
  
  sheet.insertColumnBefore(1);
  sheet.insertRowBefore(1);
  sheet.insertColumnsAfter(sheet.getMaxColumns(), 1);
  sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  
  sheet.setColumnWidth(1, 20);
  sheet.setRowHeight(1, 20);
  sheet.setColumnWidth(sheet.getMaxColumns(), 20);
  sheet.setRowHeight(sheet.getMaxRows(), 20);
  sheet.getRange(2, 1, (sheet.getMaxRows() - 1)).merge().setBackgroundRGB(0, 0, 0).setBorder(true, true, true, true, true, true);
  sheet.getRange(1, 1, 1, (sheet.getMaxColumns() - 1)).merge().setBackgroundRGB(0, 0, 0).setBorder(true, true, true, true, true, true);
  sheet.getRange(sheet.getMaxRows(), 2, 1, (sheet.getMaxColumns() - 1)).merge().setBackgroundRGB(0, 0, 0).setBorder(true, true, true, true, true, true);
  sheet.getRange(1, sheet.getMaxColumns(), (sheet.getMaxRows() - 1), 1).merge().setBackgroundRGB(0, 0, 0).setBorder(true, true, true, true, true, true);
  
  sheet.insertColumnBefore(1);
  
  sheet.getRange(1, 1, 1, 1).setValue("Prime numbers").setFontColor("Black").setBackgroundRGB(242, 211, 111).setBorder(true, true, true, true, true, true);
  sheet.getRange(2, 1, 1, 1).setValue("Perfect squares").setFontColor("Black").setBackgroundRGB(200, 255, 142).setBorder(true, true, true, true, true, true);
  sheet.getRange(3, 1, 1, 1).setValue("Even numbers").setFontColor("Purple").setBackgroundRGB(198, 198, 198).setBorder(true, true, true, true, true, true);
  sheet.getRange(4, 1, 1, 1).setValue("Odd numbers").setFontColor("Red").setBackgroundRGB(198, 198, 198).setBorder(true, true, true, true, true, true);
  sheet.getRange(5, 1, (sheet.getMaxRows() - 4), 1).merge().setBackgroundRGB(0, 0, 0).setBorder(true, true, true, true, true, true);
  
  sheet.autoResizeColumn(1);
  
  return true;
}

function isPrime(n) {
  if(typeof n !== "number") return false;
  if(Math.floor(n) !== n) return false;
  if(n <= 1) return false;
  if(n <= 3) return true;
  if(n % 2 === 0 || n % 3 === 0) return false;
  for(var i = 5; i*i <= n; i += 6) {
    if(n % i === 0  || n % (i + 2) === 0) return false;
  }
  return true;
}
