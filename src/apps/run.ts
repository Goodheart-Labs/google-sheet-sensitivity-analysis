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
    baseInputColumnColumnIndex,
    pessimisticInputColumnColumnIndex,
    optimisticInputColumnColumnIndex,
    baseOutputColumnColumnIndex,
    pessimisticOutputColumnColumnIndex,
    optimisticOutputColumnColumnIndex,
  } = properties.getProperties();

  // Prepare the scenario columns - this is a little overengineered right now
  // but will be useful when we add support for more than these three scenarios
  //
  // TODO: configurable scenarios

  const scenarioColumnColumnIndexes = [
    {
      name: 'Base',
      base: true,
      inputIndex: baseInputColumnColumnIndex,
      outputIndex: baseOutputColumnColumnIndex,
    },
    {
      name: 'Pessimistic',
      base: false,
      inputIndex: pessimisticInputColumnColumnIndex,
      outputIndex: pessimisticOutputColumnColumnIndex,
    },
    {
      name: 'Optimistic',
      base: false,
      inputIndex: optimisticInputColumnColumnIndex,
      outputIndex: optimisticOutputColumnColumnIndex,
    },
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
      scenarioColumnColumnIndexes.flatMap(
        ({ name, inputIndex, outputIndex }) => [
          [name.toLowerCase() + ' input column', inputIndex],
          [name.toLowerCase() + ' output column', outputIndex],
        ],
      ),
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
      .setName(
        `${sheetName} - Sensitivity Analysis (${
          new Date().toDateString() + ' ' + new Date().toLocaleTimeString()
        })`,
      )
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

  // Scenario Input Columns
  // --------------------------------------------------------------------------

  for (let i = 0; i < scenarioColumnColumnIndexes.length; i++) {
    const { name, inputIndex } = scenarioColumnColumnIndexes[i];
    // Are there any numerical values in the column?

    console.log(`Validating the ${name} input column`);

    const scenarioColumn = newSheet.getRange(`${inputIndex}:${inputIndex}`);

    const scenarioColumnValues = scenarioColumn.getValues();

    const scenarioColumnNumericalValues = scenarioColumnValues
      .flat()
      .filter((value) => typeof value === 'number');

    if (scenarioColumnNumericalValues.length === 0) {
      return error(
        `The ${name} input column (${i}) does not have any numerical values.`,
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
          } in the ${name} input column (${row}) does not have a numerical value.`,
        );
      }

      console.log(`Row ${row + 1} has a numerical value`);
    }

    console.log(`Validated the ${name} input column`);
  }

  // Okay, we're good to go!
  // --------------------------------------------------------------------------

  console.log('All validations passed. Running the sensitivity analysis.');

  // Set all of the scenario switcher cells to the scenario marked base: true

  const resetEverythingToBase = () => {
    const baseScenario = scenarioColumnColumnIndexes.find(({ base }) => base);
    if (!baseScenario) {
      return error('Could not find the base scenario column.');
    }

    const baseScenarioName = baseScenario.name;

    for (
      let row = 0;
      row < trueIfRowHasDropdownFalseOtherwiseArray.length;
      row++
    ) {
      const trueIfRowHasDropdown = trueIfRowHasDropdownFalseOtherwiseArray[row];
      if (!trueIfRowHasDropdown) continue;

      const scenarioSwitcherCell = scenarioSwitcherColumn.getCell(row + 1, 1);
      scenarioSwitcherCell.setValue(baseScenarioName);
    }

    console.log(`Reset everything to the ${baseScenarioName} scenario.`);
  };

  // Loop through each row with a dropdown

  for (
    let row = 0;
    row < trueIfRowHasDropdownFalseOtherwiseArray.length;
    row++
  ) {
    const trueIfRowHasDropdown = trueIfRowHasDropdownFalseOtherwiseArray[row];
    if (!trueIfRowHasDropdown) continue;

    // Reset everything to the base scenario

    resetEverythingToBase();

    // For each row, loop through each scenario

    for (let i = 0; i < scenarioColumnColumnIndexes.length; i++) {
      const { name, outputIndex } = scenarioColumnColumnIndexes[i];

      // Set the scenario switcher to that scenario

      const scenarioSwitcherCell = scenarioSwitcherColumn.getCell(row + 1, 1);

      console.log(
        `Setting the scenario switcher cell (${scenarioSwitcherCell.getA1Notation()}) to ${name}`,
      );

      scenarioSwitcherCell.setValue(name);

      // Get the value of the model output cell

      const modelOutputCellValue = modelOutputCell.getValue();

      console.log(
        `The value of the model output cell (${modelOutputCell.getA1Notation()}) is ${modelOutputCellValue}`,
      );

      // Set the scenario output cell to that value

      const scenarioOutputCell = newSheet.getRange(`${outputIndex}${row + 1}`);

      console.log(
        `Setting the scenario output cell (${scenarioOutputCell.getA1Notation()}) to ${modelOutputCellValue}`,
      );

      scenarioOutputCell.setValue(modelOutputCellValue);
    }
  }

  // Reset everything to the base scenario one last time

  resetEverythingToBase();

  // Done!
  // --------------------------------------------------------------------------

  return { success: true };
};
