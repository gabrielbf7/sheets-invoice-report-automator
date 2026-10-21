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


/**
 * Emails the weekly sales summary to the spreadsheet owner.
 * TODO: implement.
 */
function sendWeeklySummary() {
  SpreadsheetApp.getActive().toast('Send weekly summary: coming soon.', 'Automations');
}