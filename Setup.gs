/**
 * Setup.gs — builds the demo spreadsheet with fake data.
 * Run setupDemo() once. WARNING: it overwrites the Orders, Clients
 * and Invoice Template sheets.
 */

const SEED = 42;                 // fixed seed -> same data every run
const NUM_CLIENTS = 50;
const TOTAL_ORDERS = 200;
const PENDING_ORDERS = 25;       // orders from the last 7 days
const MAX_PENDING_PER_CLIENT = 10; // must fit the invoice template (10 lines)
const DAYS_BACK = 60;
const CURRENCY_FORMAT = '$#,##0.00';
const DATE_FORMAT = 'yyyy-mm-dd';

const FIRST_NAMES = ['Olivia', 'Liam', 'Emma', 'Noah', 'Ava', 'Ethan', 'Mia', 'Lucas',
  'Sophia', 'Mason', 'Grace', 'Henry', 'Chloe', 'Owen', 'Lily', 'Jack', 'Zoe', 'Leo',
  'Nora', 'Caleb'];
const LAST_NAMES = ['Carter', 'Hughes', 'Bennett', 'Foster', 'Reed', 'Morgan', 'Hayes',
  'Brooks', 'Price', 'Ward', 'Sullivan', 'Perry', 'Coleman', 'Jenkins', 'Barnes',
  'Fisher', 'Hunt', 'Palmer', 'Spencer', 'Wells'];
const COMPANY_WORDS = ['Blue Harbor', 'Maple', 'Summit', 'Riverstone', 'Golden Oak',
  'Pine Hill', 'Silver Lake', 'Redwood', 'Cedar', 'Lakeside'];
const COMPANY_TYPES = ['Bakery', 'Studio', 'Fitness', 'Dental', 'Consulting', 'Coffee',
  'Realty', 'Florist', 'Auto Repair', 'Law Office'];
const STREETS = ['Main St', 'Oak Ave', 'Elm St', 'Park Blvd', 'Lake Dr', 'Hill Rd',
  'River Rd', 'Market St', 'Sunset Ave', 'Maple Ln'];
const CITIES = ['Austin, TX', 'Denver, CO', 'Portland, OR', 'Raleigh, NC', 'Tampa, FL',
  'Columbus, OH', 'Phoenix, AZ', 'Madison, WI', 'Boise, ID', 'Richmond, VA'];
const ITEMS = [
  { name: 'Website maintenance', price: 120, maxQty: 2 },
  { name: 'Logo design', price: 250, maxQty: 1 },
  { name: 'SEO audit', price: 300, maxQty: 1 },
  { name: 'Hosting (monthly)', price: 25, maxQty: 6 },
  { name: 'Consulting hour', price: 80, maxQty: 5 },
  { name: 'Copywriting (page)', price: 60, maxQty: 4 },
  { name: 'Social media pack', price: 150, maxQty: 2 },
  { name: 'Product photo', price: 15, maxQty: 10 },
  { name: 'Email campaign', price: 90, maxQty: 3 },
  { name: 'Support ticket', price: 40, maxQty: 4 }
];

function setupDemo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rng = makeRng_(SEED);

  const clients = buildClients_(rng);
  const orders = buildOrders_(rng, clients);

  writeOrders_(getOrCreateSheet_(ss, 'Orders', 0), orders);
  writeClients_(getOrCreateSheet_(ss, 'Clients', 1), clients);
  buildInvoiceTemplate_(getOrCreateSheet_(ss, 'Invoice Template', 2));
  removeEmptyDefaultSheets_(ss, ['Orders', 'Clients', 'Invoice Template']);

  ss.setActiveSheet(ss.getSheetByName('Orders'));
  ss.toast(`Demo ready: ${clients.length} clients, ${orders.length} orders ` +
    `(${PENDING_ORDERS} pending).`, 'Setup', 8);
}

// ---------- Data generation ----------

function buildClients_(rng) {
  const names = uniquePairs_(rng, FIRST_NAMES, LAST_NAMES, NUM_CLIENTS);
  const companies = uniquePairs_(rng, COMPANY_WORDS, COMPANY_TYPES, NUM_CLIENTS);

  const clients = names.map(([first, last], i) => ({
    name: `${first} ${last}`,
    email: `${first}.${last}@example.com`.toLowerCase(),
    company: `${companies[i][0]} ${companies[i][1]}`,
    address: `${randInt_(rng, 100, 9999)} ${pick_(rng, STREETS)}`,
    city: pick_(rng, CITIES),
    // Some clients buy more: 10 heavy buyers, 15 regular, 25 occasional
    weight: i < 10 ? 5 : i < 25 ? 2 : 1
  }));
  return clients.sort((a, b) => a.name.localeCompare(b.name));
}

function buildOrders_(rng, clients) {
  const orders = [];
  const pendingCount = {};

  for (let i = 0; i < TOTAL_ORDERS; i++) {
    const isPending = i < PENDING_ORDERS;
    const daysAgo = isPending ? randInt_(rng, 0, 6) : randInt_(rng, 7, DAYS_BACK - 1);

    let client;
    do {
      client = weightedPick_(rng, clients);
    } while (isPending && (pendingCount[client.name] || 0) >= MAX_PENDING_PER_CLIENT);
    if (isPending) pendingCount[client.name] = (pendingCount[client.name] || 0) + 1;

    const item = pick_(rng, ITEMS);
    orders.push({
      daysAgo, client, item,
      qty: randInt_(rng, 1, item.maxQty),
      status: isPending ? 'pending' : 'invoiced',
      invoiceNo: ''
    });
  }

  // Oldest first, then give each order an ID
  orders.sort((a, b) => b.daysAgo - a.daysAgo);

  // History: one invoice per client per week for old orders
  const invoiceByKey = {};
  let nextInvoice = 1;
  orders.forEach(o => {
    if (o.status !== 'invoiced') return;
    const key = `${o.client.name}|${Math.floor(o.daysAgo / 7)}`;
    if (!invoiceByKey[key]) {
      invoiceByKey[key] = 'INV-' + String(nextInvoice++).padStart(4, '0');
    }
    o.invoiceNo = invoiceByKey[key];
  });
  return orders;
}

// ---------- Sheet writers ----------

function writeOrders_(sheet, orders) {
  const headers = ['Order ID', 'Date', 'Client', 'Email', 'Item', 'Qty',
    'Unit Price', 'Total', 'Status', 'Invoice No', 'Invoice Link'];
  const today = new Date();

  const rows = orders.map((o, i) => [
    'ORD-' + String(i + 1).padStart(4, '0'),
    new Date(today.getFullYear(), today.getMonth(), today.getDate() - o.daysAgo),
    o.client.name, o.client.email, o.item.name, o.qty, o.item.price,
    o.qty * o.item.price, o.status, o.invoiceNo, ''
  ]);

  resetSheet_(sheet);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  sheet.getRange(2, 2, rows.length, 1).setNumberFormat(DATE_FORMAT);
  sheet.getRange(2, 7, rows.length, 2).setNumberFormat(CURRENCY_FORMAT);

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['pending', 'invoiced'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 9, rows.length, 1).setDataValidation(statusRule);

  styleHeader_(sheet, headers.length);
}

function writeClients_(sheet, clients) {
  const headers = ['Client', 'Email', 'Company', 'Address', 'City'];
  const rows = clients.map(c => [c.name, c.email, c.company, c.address, c.city]);

  resetSheet_(sheet);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  styleHeader_(sheet, headers.length);
}

function buildInvoiceTemplate_(sheet) {
  resetSheet_(sheet);
  sheet.setHiddenGridlines(true);

  // Header
  sheet.getRange('B2').setValue('INVOICE').setFontSize(24).setFontWeight('bold');
  sheet.getRange('B3:B4').setValues([['Demo Studio LLC'], ['billing@example.com']]);
  sheet.getRange('D3:D5').setValues([['Invoice No:'], ['Issue date:'], ['Due date:']])
    .setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('E4:E5').setNumberFormat(DATE_FORMAT);

  // Bill to: B7:B10 = client, company, address, city (filled by the script)
  sheet.getRange('B6').setValue('Bill to:').setFontWeight('bold');

  // Items: B13:E22 (max 10 lines)
  sheet.getRange('B12:E12').setValues([['Item', 'Qty', 'Unit Price', 'Total']])
    .setFontWeight('bold').setBackground('#e8eaed');
  sheet.getRange('B12:E22').setBorder(true, true, true, true, false, false);
  sheet.getRange('D13:E22').setNumberFormat(CURRENCY_FORMAT);

  // Total
  sheet.getRange('D24').setValue('TOTAL').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('E24').setFormula('=SUM(E13:E22)')
    .setNumberFormat(CURRENCY_FORMAT).setFontWeight('bold');

  sheet.getRange('B26').setValue('Thank you for your business!').setFontStyle('italic');

  sheet.setColumnWidth(1, 20);
  sheet.setColumnWidth(2, 260);
  sheet.setColumnWidths(3, 3, 110);
}

// ---------- Helpers ----------

function getOrCreateSheet_(ss, name, index) {
  return ss.getSheetByName(name) || ss.insertSheet(name, index);
}

function resetSheet_(sheet) {
  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
}

function styleHeader_(sheet, numCols) {
  sheet.getRange(1, 1, 1, numCols).setFontWeight('bold').setBackground('#e8eaed');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, numCols);
}

// Deletes leftover empty sheets (e.g. the default "Sheet1")
function removeEmptyDefaultSheets_(ss, keepNames) {
  ss.getSheets().forEach(s => {
    if (!keepNames.includes(s.getName()) && s.getLastRow() === 0) ss.deleteSheet(s);
  });
}

// Seeded random number generator (mulberry32): same seed -> same sequence
function makeRng_(seed) {
  let a = seed >>> 0;
  return function () {
    a = (a + 0x6D2B79F5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function randInt_(rng, min, max) {
  return min + Math.floor(rng() * (max - min + 1));
}

function pick_(rng, list) {
  return list[Math.floor(rng() * list.length)];
}

function weightedPick_(rng, list) {
  const total = list.reduce((sum, x) => sum + x.weight, 0);
  let r = rng() * total;
  for (const x of list) {
    r -= x.weight;
    if (r < 0) return x;
  }
  return list[list.length - 1];
}

// Returns n unique [a, b] combinations in random (seeded) order
function uniquePairs_(rng, listA, listB, n) {
  const pairs = [];
  listA.forEach(a => listB.forEach(b => pairs.push([a, b])));
  for (let i = pairs.length - 1; i > 0; i--) {
    const j = Math.floor(rng() * (i + 1));
    [pairs[i], pairs[j]] = [pairs[j], pairs[i]];
  }
  return pairs.slice(0, n);
}