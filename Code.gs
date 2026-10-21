/**
 * Adds the "Automations" menu when the spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Automations')
    .addItem('Generate invoices', 'generateInvoices')
    .addItem('Send weekly summary', 'sendWeeklySummary')
    .addToUi();
}

