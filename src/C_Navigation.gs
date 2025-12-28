/**
 * Navigation helpers for sidebar forms.
 */
function showPaymentForm() {
  const html = HtmlService.createHtmlOutputFromFile('F_Transaction')
    .setTitle('New Payment')
    .setWidth(420);

  SpreadsheetApp.getUi().showSidebar(html);
}
