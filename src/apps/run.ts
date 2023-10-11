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
}: ExecuteSensitivityAnalysis) => {
  const properties = PropertiesService.getDocumentProperties();
  const { scenarioSwitcher } = properties.getProperties();

  // Are we running in a new sheet? If so, duplicate the current sheet

  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let newSheet: GoogleAppsScript.Spreadsheet.Sheet;
  if (runInNewSheet) {
    newSheet = sheet.copyTo(spreadsheet);
    newSheet.setName(`${sheetName} - Sensitivity Analysis`);
  } else {
    newSheet = sheet;
  }

  // Find the scenario switcher column

  const scenarioSwitcherColumn = newSheet
    .getRange(1, 1, 1, newSheet.getMaxColumns())
    .getValues()[0]
    .indexOf(scenarioSwitcher);

  // Does the column have any dropdowns?

  const rules = newSheet
    .getRange(2, scenarioSwitcherColumn + 1, newSheet.getMaxRows() - 1, 1)
    .getDataValidations();

  // Validate the dropdowns

  const dropdownsExist = rules.some(
    (rule) =>
      rule[0]?.getCriteriaType() ===
      SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST,
  );

  if (!dropdownsExist) {
    return {
      success: false,
      message: `The Scenario Switcher column (${scenarioSwitcher}) does not have any dropdowns.`,
    };
  }

  const dropdowns = rules.every(
    (rule) =>
      rule[0]?.getCriteriaType() !==
        SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST ||
      (rule[0]?.getCriteriaValues().includes('Base') &&
        rule[0]?.getCriteriaValues().includes('Pessimistic') &&
        rule[0]?.getCriteriaValues().includes('Optimistic')),
  );

  if (!dropdowns) {
    return {
      success: false,
      message: `The Scenario Switcher column (${scenarioSwitcher}) does not have the correct values (Base, Pessimistic, Optimistic).`,
    };
  }

  return { success: true };
};
