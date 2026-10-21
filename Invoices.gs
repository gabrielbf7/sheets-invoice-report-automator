/**
 * Invoices.gs — Invoice generation for pending orders.
 */

const ORDERS_SHEET_NAME = 'Orders';
const CLIENTS_SHEET_NAME = 'Clients';
const TEMPLATE_SHEET_NAME = 'Invoice Template';
const INVOICES_FOLDER_NAME = 'Invoices';
const INVOICE_DUE_DAYS = 30;   // payment terms: Net 30
const INVOICE_MAX_ITEMS = 10;  // template rows B13:E22

// Zero-based column indexes in the Orders sheet
const ORDER_COLS = {
  ORDER_ID: 0, DATE: 1, CLIENT: 2, EMAIL: 3, ITEM: 4, QTY: 5,
  UNIT_PRICE: 6, TOTAL: 7, STATUS: 8, INVOICE_NO: 9, INVOICE_LINK: 10
};

/**
 * Menu action: generate one PDF invoice per client with pending orders,
 * save it to the Invoices folder and mark the orders as invoiced.
 */
function generateInvoices() {
  const ss = SpreadsheetApp.getActive();
  const groups = getPendingOrdersByClient_();
  const clients = Object.keys(groups);

  if (clients.length === 0) {
    ss.toast('No pending orders', 'Generate invoices');
    return;
  }

  let nextNumber = getNextInvoiceNumber_();

  clients.forEach(client => {
    const orders = groups[client];
    const invoiceNo = formatInvoiceNumber_(nextNumber++);

    fillInvoiceTemplate_(client, orders, invoiceNo);
    const pdf = exportInvoicePdf_(invoiceNo, client);
    markOrdersInvoiced_(orders, invoiceNo, pdf.getUrl());

    Logger.log(invoiceNo + ' — ' + client + ' (' + orders.length + ' order(s))');
    Utilities.sleep(2000); // pause between exports to avoid rate limits
  });

  Logger.log('Invoices generated: ' + clients.length);
  ss.toast(clients.length + ' invoice(s) generated', 'Generate invoices');
}

/**
 * Reads the Orders sheet and returns pending orders grouped by client:
 * { "Client Name": [ {rowNumber, orderId, date, item, qty, unitPrice, total}, ... ] }
 * rowNumber is kept so we can mark the orders as invoiced later.
 */
function getPendingOrdersByClient_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ORDERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const groups = {};

  for (let i = 1; i < data.length; i++) { // skip header row
    const row = data[i];
    if (String(row[ORDER_COLS.STATUS]).trim() !== 'pending') continue;

    const client = row[ORDER_COLS.CLIENT];
    if (!groups[client]) groups[client] = [];
    groups[client].push({
      rowNumber: i + 1,
      orderId: row[ORDER_COLS.ORDER_ID],
      date: row[ORDER_COLS.DATE],
      item: row[ORDER_COLS.ITEM],
      qty: row[ORDER_COLS.QTY],
      unitPrice: row[ORDER_COLS.UNIT_PRICE],
      total: row[ORDER_COLS.TOTAL]
    });
  }
  return groups;
}

/**
 * Returns the next invoice number as an integer (highest existing INV-#### + 1).
 */
function getNextInvoiceNumber_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ORDERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const match = String(data[i][ORDER_COLS.INVOICE_NO]).match(/^INV-(\d+)$/);
    if (match) max = Math.max(max, Number(match[1]));
  }
  return max + 1;
}

/**
 * Formats an integer as an invoice number, e.g. 135 → "INV-0135".
 */
function formatInvoiceNumber_(n) {
  return 'INV-' + String(n).padStart(4, '0');
}

/**
 * Looks up a client in the Clients sheet (Client | Email | Company | Address | City).
 */
function getClientInfo_(clientName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CLIENTS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === clientName) {
      return { client: data[i][0], company: data[i][2], address: data[i][3], city: data[i][4] };
    }
  }
  return { client: clientName, company: '', address: '', city: '' }; // fallback
}

/**
 * Writes invoice header, bill-to and line items into the Invoice Template sheet.
 */
function fillInvoiceTemplate_(clientName, orders, invoiceNo) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TEMPLATE_SHEET_NAME);
  const info = getClientInfo_(clientName);
  const today = new Date();
  const due = new Date(today.getTime() + INVOICE_DUE_DAYS * 24 * 60 * 60 * 1000);

  sheet.getRange('E3').setValue(invoiceNo);
  sheet.getRange('E4:E5').setValues([[today], [due]]).setNumberFormat('yyyy-mm-dd');
  sheet.getRange('B7:B10').setValues([[info.client], [info.company], [info.address], [info.city]]);

  sheet.getRange('B13:E22').clearContent(); // clear previous items
  const rows = orders.slice(0, INVOICE_MAX_ITEMS)
    .map(o => [o.item, o.qty, o.unitPrice, o.total]);
  sheet.getRange(13, 2, rows.length, 4).setValues(rows); // starts at B13

  SpreadsheetApp.flush(); // make sure changes are written before export
}

/**
 * Returns the "Invoices" folder next to the spreadsheet, creating it if needed.
 * Found by name, so no private folder IDs are stored in the code.
 */
function getInvoicesFolder_() {
  const ssFile = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const parents = ssFile.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  const existing = parent.getFoldersByName(INVOICES_FOLDER_NAME);
  return existing.hasNext() ? existing.next() : parent.createFolder(INVOICES_FOLDER_NAME);
}

/**
 * Exports only the Invoice Template sheet as a PDF and saves it to the Invoices folder.
 * Retries with increasing waits if Google returns 429 (too many requests).
 */
function exportInvoicePdf_(invoiceNo, clientName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(TEMPLATE_SHEET_NAME);

  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export' +
    '?format=pdf&gid=' + sheet.getSheetId() +
    '&size=letter&portrait=true&fitw=true' +
    '&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED' +
    '&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5';

  const maxAttempts = 5;
  let waitMs = 5000;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true // handle errors ourselves instead of crashing
    });
    const code = response.getResponseCode();

    if (code === 200) {
      const fileName = invoiceNo + ' - ' + clientName + '.pdf';
      const blob = response.getBlob().setName(fileName);
      return getInvoicesFolder_().createFile(blob);
    }

    if (code === 429 && attempt < maxAttempts) {
      Logger.log('Rate limited on ' + invoiceNo + ', retrying in ' + waitMs / 1000 + 's');
      Utilities.sleep(waitMs);
      waitMs *= 2; // 5s, 10s, 20s, 40s
      continue;
    }

    throw new Error('PDF export failed for ' + invoiceNo + ' (HTTP ' + code + ')');
  }
}

/**
 * Sets Status = invoiced, Invoice No and Invoice Link for each order row.
 */
function markOrdersInvoiced_(orders, invoiceNo, pdfUrl) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ORDERS_SHEET_NAME);
  orders.forEach(o => {
    // Columns I:K (Status, Invoice No, Invoice Link) = sheet columns 9..11
    sheet.getRange(o.rowNumber, ORDER_COLS.STATUS + 1, 1, 3)
      .setValues([['invoiced', invoiceNo, pdfUrl]]);
  });
}