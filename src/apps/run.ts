export const renderConfirm = () => {
  const html = HtmlService.createTemplateFromFile('app/assets/confirm.html');
  const htmlBody = html.evaluate().getContent();

  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(HtmlService.createHtmlOutput(htmlBody), 'Confirm');
};

type ExecuteSensitivityAnalysis = {
  runInNewSheet?: boolean;
};

export const executeSensitivityAnalysis = ({
  runInNewSheet = false,
}: ExecuteSensitivityAnalysis = {}) => {
  const properties = PropertiesService.getDocumentProperties();

  const { scenarioSwitcherColumnIndex, modelOutputCellIndex } =
    properties.getProperties();

  // TODO: configurable scenarios
  const scenarios = ['Base', 'Pessimistic', 'Optimistic'];

  console.log(
    `Running Sensitivity Analysis with the following config: ${JSON.stringify(
      {
        runInNewSheet,
        scenarioSwitcherColumnIndex,
        modelOutputCellIndex,
        scenarios,
      },
      null,
      2,
    )}`,
  );

  // Configurations
  // --------------------------------------------------------------------------

  // Check that we have all the required config values

  const missingValues = [
    ['scenario switcher', scenarioSwitcherColumnIndex],
    ['model output cell', modelOutputCellIndex],
  ]
    .filter(([, value]) => !value)
    .map(([name]) => name);

  if (missingValues.length > 0) {
    console.error(
      `The following configuration values are missing: ${missingValues.join(
        ', ',
      )}. Please use the "Configure" menu to set them.`,
    );

    return {
      success: false,
      message: `The following configuration values are missing: ${missingValues.join(
        ', ',
      )}. Please use the "Configure" menu to set them.`,
    };
  }

  // New Sheet
  // --------------------------------------------------------------------------

  // Are we running in a new sheet? If so, duplicate the current sheet

  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let newSheet: GoogleAppsScript.Spreadsheet.Sheet;
  if (runInNewSheet) {
    newSheet = sheet
      .copyTo(spreadsheet)
      .setName(`${sheetName} - Sensitivity Analysis`)
      .activate();
  } else {
    newSheet = sheet;
  }

  // TODO: store that we ran in a new sheet or didn't run in a new sheet in the
  // properties, so that we can default to the last used option

  // Scenario Switcher Column
  // --------------------------------------------------------------------------

  // Find the scenario switcher column

  console.log(
    `Finding the scenario switcher column (${scenarioSwitcherColumnIndex})`,
  );

  const scenarioSwitcherColumn = newSheet.getRange(
    `${scenarioSwitcherColumnIndex}:${scenarioSwitcherColumnIndex}`,
  );

  // Does the column have any dropdowns?

  console.log('Validating the scenario switcher column');

  const rules = scenarioSwitcherColumn.getDataValidations();

  console.log('Rules:', JSON.stringify(rules, null, 2));

  // Validate the dropdowns

  const dropdownsExist = rules.some(
    (rule) =>
      rule[0]?.getCriteriaType() ===
      SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST,
  );

  if (!dropdownsExist) {
    console.error(
      `The Scenario Switcher column (${scenarioSwitcherColumnIndex}) does not have any dropdowns.`,
    );

    return {
      success: false,
      message: `The Scenario Switcher column (${scenarioSwitcherColumnIndex}) does not have any dropdowns.`,
    };
  }

  console.log(
    `Found some dropdowns. Validating that they have the correct values ${scenarios.join(
      ', ',
    )}.`,
  );

  const validDropdowns = rules.every(
    (rule) =>
      // Either it's not a dropdown
      rule[0]?.getCriteriaType() !==
        SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST ||
      // Or it is a dropdown and all its values are the scenarios
      (rule[0]?.getCriteriaValues()[0].length === scenarios.length &&
        rule[0]
          ?.getCriteriaValues()[0]
          .filter((value: string) => scenarios.includes(value)).length ===
          scenarios.length),
  );

  if (!validDropdowns) {
    console.error(
      `The Scenario Switcher column (${scenarioSwitcherColumnIndex}) does not have the correct values (${scenarios.join(
        ', ',
      )}).`,
    );

    return {
      success: false,
      message: `The Scenario Switcher column (${scenarioSwitcherColumnIndex}) does not have the correct values (${scenarios.join(
        ', ',
      )}).`,
    };
  }

  console.log('Validated the dropdowns.');

  // Model Output Cell
  // --------------------------------------------------------------------------

  // Find the model output cell

  console.log('Finding the model output cell');

  const modelOutputCell = newSheet.getRange(modelOutputCellIndex);

  // Does the cell have a non-empty value?

  console.log('Validating the model output cell');

  const modelOutputCellValue = modelOutputCell.getValue();

  if (
    modelOutputCellValue === '' ||
    modelOutputCellValue === null ||
    modelOutputCellValue === undefined
  ) {
    console.error(
      `The Model Output cell (${modelOutputCellIndex}) does not have a value.`,
    );

    return {
      success: false,
      message: `The Model Output cell (${modelOutputCellIndex}) does not seem to have a value. Make sure it is the output of the model and is non-empty.`,
    };
  }

  // Done!
  // --------------------------------------------------------------------------

  return { success: true };
};
