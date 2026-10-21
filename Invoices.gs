/**
 * Invoices.gs — Invoice generation for pending orders.
 */

const ORDERS_SHEET_NAME = 'Orders';

// Zero-based column indexes in the Orders sheet
const ORDER_COLS = {
  ORDER_ID: 0, DATE: 1, CLIENT: 2, EMAIL: 3, ITEM: 4, QTY: 5,
  UNIT_PRICE: 6, TOTAL: 7, STATUS: 8, INVOICE_NO: 9, INVOICE_LINK: 10
};

/**
 * Menu action: generate one invoice per client with pending orders.
 * Step 1: only reads and logs pending orders grouped by client.
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
  SpreadsheetApp.getActive().toast(
    clients.length + ' client(s) with pending orders', 'Generate invoices');
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