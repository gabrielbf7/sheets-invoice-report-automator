/**
 * Weekly sales summary: total sold, number of invoices and top clients
 * for the last SUMMARY_DAYS days (today included).
 */

const SUMMARY_DAYS = 7;
const SUMMARY_TOP_CLIENTS = 5;

function sendWeeklySummary() {
  const stats = getWeeklyStats_();
  const recipient = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  const subject = `Weekly sales summary (${stats.from} to ${stats.to})`;

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: buildSummaryHtml_(stats),
  });

  Logger.log(`Summary sent: ${subject}`);
  SpreadsheetApp.getActive().toast('Weekly summary sent to your email.', 'Automations', 5);
}

function buildSummaryHtml_(stats) {
  const usd = n => '$' + n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  const rows = stats.topClients.map((c, i) =>
    `<tr><td>${i + 1}</td><td>${c.client}</td><td style="text-align:right">${usd(c.total)}</td></tr>`
  ).join('');

  return `
    <h2>Weekly sales summary</h2>
    <p>${stats.from} to ${stats.to}</p>
    <ul>
      <li><b>Total sold:</b> ${usd(stats.totalSold)}</li>
      <li><b>Orders:</b> ${stats.orders}</li>
      <li><b>Invoices:</b> ${stats.invoices}</li>
    </ul>
    <h3>Top clients</h3>
    <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse">
      <tr><th>#</th><th>Client</th><th>Total</th></tr>
      ${rows}
    </table>
    <p style="color:#888;font-size:12px">Sent automatically by Sheets Invoice &amp; Report Automator.</p>`;
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

function enableWeeklySummary() {
  removeSummaryTriggers_();
  ScriptApp.newTrigger('sendWeeklySummary')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
  SpreadsheetApp.getActive().toast('Weekly summary enabled: every Monday, 8-9 AM.', 'Automations', 5);
}

function disableWeeklySummary() {
  const removed = removeSummaryTriggers_();
  SpreadsheetApp.getActive().toast(`Weekly summary disabled (${removed} trigger(s) removed).`, 'Automations', 5);
}

function removeSummaryTriggers_() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendWeeklySummary');
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  return triggers.length;
}