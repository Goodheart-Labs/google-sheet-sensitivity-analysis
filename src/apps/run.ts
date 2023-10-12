import { error } from './utils';

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
  // Pull in the config values
  // --------------------------------------------------------------------------

  const properties = PropertiesService.getDocumentProperties();

  const {
    scenarioSwitcherColumnIndex,
    modelOutputCellIndex,
    baseColumnColumnIndex,
    pessimisticColumnColumnIndex,
    optimisticColumnColumnIndex,
  } = properties.getProperties();

  // Prepare the scenario columns - this is a little overengineered right now
  // but will be useful when we add support for more than these three scenarios
  //
  // TODO: configurable scenarios

  const scenarioColumnColumnIndexes = [
    { name: 'Base', index: baseColumnColumnIndex },
    { name: 'Pessimistic', index: pessimisticColumnColumnIndex },
    { name: 'Optimistic', index: optimisticColumnColumnIndex },
  ];

  console.log(
    `Running Sensitivity Analysis with the following config: ${JSON.stringify(
      {
        runInNewSheet,
        scenarioSwitcherColumnIndex,
        modelOutputCellIndex,
        scenarioColumnColumnIndexes,
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
    .concat(
      scenarioColumnColumnIndexes.map(({ name, index }) => [
        name.toLowerCase() + ' column',
        index,
      ]),
    )
    .filter(([, value]) => !value)
    .map(([name]) => name);

  if (missingValues.length > 0) {
    return error(
      `The following configuration values are missing: ${missingValues.join(
        ', ',
      )}. Please use the "Configure" menu to set them.`,
    );
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

  // Get the row index of every dropdown

  const trueIfRowHasDropdownFalseOtherwiseArray = rules.map(
    (rule) =>
      rule[0]?.getCriteriaType() ===
      SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST,
  );

  if (!trueIfRowHasDropdownFalseOtherwiseArray.filter(Boolean).length) {
    return error(
      `The Scenario Switcher column (${scenarioSwitcherColumnIndex}) does not have any dropdowns.`,
    );
  }

  const scenarioNames = Object.values(scenarioColumnColumnIndexes).map(
    ({ name }) => name,
  );

  console.log(
    `Found some dropdowns. Validating that they have the correct values: ${scenarioNames.join(
      ', ',
    )}`,
  );

  // Validate the dropdowns

  const validDropdowns = rules
    .map(
      (rule) =>
        rule[0]?.getCriteriaType() ===
          SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST &&
        rule[0]?.getCriteriaValues(),
    )
    .filter((values): values is [string[], boolean] => !!values)
    .every(
      ([values]) =>
        values.length === scenarioNames.length &&
        values.filter((v) => scenarioNames.includes(v)).length ===
          scenarioNames.length,
    );

  if (!validDropdowns) {
    return error(
      `The Scenario Switcher column (${scenarioSwitcherColumnIndex}) does not have the correct values (${scenarioNames.join(
        ', ',
      )}).`,
    );
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
    return error(
      `The Model Output cell (${modelOutputCellIndex}) does not seem to have a value. Make sure it is the output of the model and is non-empty.`,
    );
  }

  // Scenario Columns
  // --------------------------------------------------------------------------

  for (let i = 0; i < scenarioColumnColumnIndexes.length; i++) {
    const { name, index } = scenarioColumnColumnIndexes[i];
    // Are there any numerical values in the column?

    console.log(`Validating the ${name} column`);

    const scenarioColumn = newSheet.getRange(`${index}:${index}`);

    const scenarioColumnValues = scenarioColumn.getValues();

    const scenarioColumnNumericalValues = scenarioColumnValues
      .flat()
      .filter((value) => typeof value === 'number');

    if (scenarioColumnNumericalValues.length === 0) {
      return error(
        `The ${name} column (${i}) does not have any numerical values.`,
      );
    }

    // Make sure each row with a dropdown has a value

    for (
      let row = 0;
      row < trueIfRowHasDropdownFalseOtherwiseArray.length;
      row++
    ) {
      const trueIfRowHasDropdown = trueIfRowHasDropdownFalseOtherwiseArray[row];
      if (!trueIfRowHasDropdown) continue;

      console.log(`Checking that row ${row + 1} has a numerical value`);
      console.log('Value:', scenarioColumnValues[row]);

      const value = scenarioColumnValues[row][0];

      if (typeof value !== 'number') {
        return error(
          `Row ${
            row + 1
          } in the ${name} column (${row}) does not have a numerical value.`,
        );
      }

      console.log(`Row ${row + 1} has a numerical value`);
    }

    console.log(`Validated the ${name} column`);
  }

  // Okay, we're good to go!
  // --------------------------------------------------------------------------

  console.log('All validations passed. Running the sensitivity analysis.');

  // Loop through each row with a dropdown

  // For each row, loop through each scenario

  // For each scenario:
  // - set the scenario switcher to that scenario
  // - get the value of the model output cell
  // - set the value of the scenario output cell in this row to the value of the model output cell

  // Keep going until we run out of rows with dropdowns

  // Done!
  // --------------------------------------------------------------------------

  return { success: true };
};
