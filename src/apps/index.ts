/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/ban-ts-comment */

import { renderConfig } from './configuration';
import { renderConfirm, executeSensitivityAnalysis } from './run';

// @ts-ignore
function goConfig() {
  renderConfig();
}

// @ts-ignore
function goRun() {
  renderConfirm();
}

// @ts-ignore
function runSensitivityAnalysis(payload) {
  return executeSensitivityAnalysis(payload);
}

// @ts-ignore
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Sensitivity Analysis')
    .addItem('Configure', 'goConfig')
    .addSeparator()
    .addItem('Run', 'goRun')
    .addToUi();
}

// @ts-ignore
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// @ts-ignore
function processJSONData(payload) {
  // Process the JSON object payload
  // Insert logic here

  return { success: true };
}
