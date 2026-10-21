/**
 * Invoices.gs — Invoice generation for pending orders.
 */

const ORDERS_SHEET_NAME = 'Orders';
const CLIENTS_SHEET_NAME = 'Clients';
const TEMPLATE_SHEET_NAME = 'Invoice Template';
const INVOICE_DUE_DAYS = 30;   // payment terms: Net 30
const INVOICE_MAX_ITEMS = 10;  // template rows B13:E22

// Zero-based column indexes in the Orders sheet
const ORDER_COLS = {
  ORDER_ID: 0, DATE: 1, CLIENT: 2, EMAIL: 3, ITEM: 4, QTY: 5,
  UNIT_PRICE: 6, TOTAL: 7, STATUS: 8, INVOICE_NO: 9, INVOICE_LINK: 10
};

/**
 * Menu action: generate one invoice per client with pending orders.
 * Current step: logs pending orders and fills the template for the first client.
 */
function generateInvoices() {
  const groups = getPendingOrdersByClient_();
  const clients = Object.keys(groups);
  let totalOrders = 0;

  clients.forEach(client => {
    const orders = groups[client];
    totalOrders += orders.length;
    Logger.log(client + ': ' + orders.length + ' order(s) — ' +
      orders.map(o => o.orderId).join(', '));
  });
  Logger.log('Clients: ' + clients.length + ' | Pending orders: ' + totalOrders);

  if (clients.length === 0) {
    SpreadsheetApp.getActive().toast('No pending orders', 'Generate invoices');
    return;
  }

  // Step 2: fill the template for the first client only (visual check)
  const firstClient = clients[0];
  const invoiceNo = getNextInvoiceNumber_();
  fillInvoiceTemplate_(firstClient, groups[firstClient], invoiceNo);
  Logger.log('Template filled for ' + firstClient + ' (' + invoiceNo + ')');
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
 * Returns the next invoice number (highest existing INV-#### in Orders + 1).
 */
function getNextInvoiceNumber_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ORDERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const match = String(data[i][ORDER_COLS.INVOICE_NO]).match(/^INV-(\d+)$/);
    if (match) max = Math.max(max, Number(match[1]));
  }
  return 'INV-' + String(max + 1).padStart(4, '0');
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

  SpreadsheetApp.flush(); // make sure changes are written before export (next step)
}