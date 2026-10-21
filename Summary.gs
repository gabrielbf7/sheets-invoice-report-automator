/**
 * Weekly sales summary: total sold, number of invoices and top clients
 * for the last SUMMARY_DAYS days (today included).
 */

const SUMMARY_DAYS = 7;
const SUMMARY_TOP_CLIENTS = 5;

function sendWeeklySummary() {
  const stats = getWeeklyStats_();
  Logger.log(JSON.stringify(stats, null, 2));
}

function getWeeklyStats_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Orders');
  const [header, ...rows] = sheet.getDataRange().getValues();
  const iDate = header.indexOf('Date');
  const iClient = header.indexOf('Client');
  const iTotal = header.indexOf('Total');
  const iInvoice = header.indexOf('Invoice No');

  const end = new Date();
  const start = new Date(end);
  start.setDate(start.getDate() - (SUMMARY_DAYS - 1));
  start.setHours(0, 0, 0, 0);

  let totalSold = 0;
  let orderCount = 0;
  const invoices = new Set();
  const byClient = {};

  rows.forEach(row => {
    const date = new Date(row[iDate]);
    if (isNaN(date) || date < start || date > end) return;

    const amount = Number(row[iTotal]) || 0;
    totalSold += amount;
    orderCount++;
    if (row[iInvoice]) invoices.add(row[iInvoice]);
    byClient[row[iClient]] = (byClient[row[iClient]] || 0) + amount;
  });

  const topClients = Object.entries(byClient)
    .sort((a, b) => b[1] - a[1])
    .slice(0, SUMMARY_TOP_CLIENTS)
    .map(([client, total]) => ({ client, total: Math.round(total * 100) / 100 }));

  const tz = Session.getScriptTimeZone();
  return {
    from: Utilities.formatDate(start, tz, 'yyyy-MM-dd'),
    to: Utilities.formatDate(end, tz, 'yyyy-MM-dd'),
    orders: orderCount,
    totalSold: Math.round(totalSold * 100) / 100,
    invoices: invoices.size,
    topClients: topClients,
  };
}