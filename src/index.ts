import { renderConfig } from './configuration';

// @ts-ignore
function goConfig() {
  renderConfig();
}

// @ts-ignore
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Sensitivity Analysis')
    .addItem('Configure', 'goConfig')
    .addToUi();
}

// @ts-ignore
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
