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
function getConfigValues() {
  const properties = PropertiesService.getDocumentProperties();
  const config = properties.getProperties();

  return config;
}

// @ts-ignore
function setConfigValues({
  config,
}: {
  config: {
    scenarioSwitcher: string;
    modelOutput: string;
    pessimisticColumn: string;
    baseColumn: string;
    optimisticColumn: string;
  };
}) {
  // TODO: server-side validation

  const properties = PropertiesService.getDocumentProperties();
  properties.setProperties(config);

  return { success: true };
}
