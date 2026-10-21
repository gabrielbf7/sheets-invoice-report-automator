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
 * Generates one PDF invoice per client with pending orders.
 * TODO: implement.
 */
function generateInvoices() {
  SpreadsheetApp.getActive().toast('Generate invoices: coming soon.', 'Automations');
}

/**
 * Emails the weekly sales summary to the spreadsheet owner.
 * TODO: implement.
 */
function sendWeeklySummary() {
  SpreadsheetApp.getActive().toast('Send weekly summary: coming soon.', 'Automations');
}