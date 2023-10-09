function writeHelloWorld() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  sheet.getRange('A1').setValue('Hello World');
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Hello World Menu')
    .addItem('Write Hello World', 'writeHelloWorld')
    .addToUi();
}
