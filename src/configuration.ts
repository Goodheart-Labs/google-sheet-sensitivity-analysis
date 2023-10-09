export const renderConfig = () => {
  const html = HtmlService.createTemplateFromFile('app/assets/config.html');
  const htmlBody = html.evaluate().getContent();

  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(HtmlService.createHtmlOutput(htmlBody), 'Configure');
};
