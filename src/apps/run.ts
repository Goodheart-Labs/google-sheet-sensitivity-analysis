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
  console.log(runInNewSheet);

  return { success: true };
};
