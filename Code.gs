/**
 * Globe AMAX Retailer Portal — normalized database backend
 * Database schema: Customers + Transactions + SIM_Inventory + Stock_In_Log + AUDIT_LOG
 * Spreadsheet timezone: Asia/Manila
 */
const SPREADSHEET_ID = '1oQZJwvQlmAHhbGMsya2wvsIXUIfbYEUP8jOhWGJFpHI';
const PORTAL_TIME_ZONE = 'Asia/Manila';
const CUSTOMER_SHEET = 'Customers';
const TRANSACTION_SHEET = 'Transactions';
const SIM_INVENTORY_SHEET = 'SIM_Inventory';
const STOCK_IN_SHEET = 'Stock_In_Log';
const AUDIT_SHEET = 'AUDIT_LOG';
const BACKUP_FOLDER_NAME = 'Sales & Load Portal Backups';
const BACKUP_PREFIX = 'Globe_AMAX_Retailer_Backup_';
const BACKUP_RETENTION_DAYS = 30;
const MAX_OFFLINE_TRANSACTIONS = 5000;
const SIM_BRANDS = ['Globe', 'TM', 'GOMO', 'RT SIM'];
const PAYMENT_STATUSES = ['Cash', 'Credit', 'Gcash Payment', 'Pending', 'Paid', 'Loaded'];
const CUSTOMER_HEADERS = ['Customer_ID', 'Customer_Name', 'Barangay', 'Phone_Number', 'Customer_Type'];
const TRANSACTION_HEADERS = ['Transaction_ID', 'Timestamp_PHT', 'Customer_ID', 'Service_Type', 'SIM_Brand', 'GCash_Type', 'Quantity', 'Unit_Price', 'Total_Amount', 'Payment_Status', 'SMS_Command', 'Notes'];
const SIM_HEADERS = ['SIM_Brand', 'Beginning_Inventory', 'Notes'];
const STOCK_HEADERS = ['StockIn_ID', 'Timestamp_PHT', 'SIM_Brand', 'Quantity', 'Notes'];
const AUDIT_HEADERS = ['Timestamp_PHT', 'Actor', 'Action', 'Entity_ID', 'Sheet', 'Row', 'Field', 'Old_Value', 'New_Value_or_Details'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Globe AMAX Retailer Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function getSpreadsheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (ss.getSpreadsheetTimeZone() !== PORTAL_TIME_ZONE) ss.setSpreadsheetTimeZone(PORTAL_TIME_ZONE);
  return ss;
}
function dateKey_(date) { return Utilities.formatDate(date || new Date(), PORTAL_TIME_ZONE, 'yyyy-MM-dd'); }
function money_(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function round2_(n) { return Math.round((Number(n) + Number.EPSILON) * 100) / 100; }
function clean_(value) { return String(value == null ? '' : value).trim(); }

/** Run once from the Apps Script editor after pasting this project. */
function initializePortalDatabase() {
  const ss = getSpreadsheet_();
  const customerSheet = ss.getSheetByName(CUSTOMER_SHEET);
  if (!customerSheet) throw new Error('Customers sheet is missing. Copy the uploaded database workbook into the linked Google spreadsheet first.');
  assertHeaders_(customerSheet, CUSTOMER_HEADERS, true);
  ensureDataSheet_(ss, TRANSACTION_SHEET, TRANSACTION_HEADERS);
  ensureDataSheet_(ss, SIM_INVENTORY_SHEET, SIM_HEADERS);
  ensureDataSheet_(ss, STOCK_IN_SHEET, STOCK_HEADERS);
  ensureDataSheet_(ss, AUDIT_SHEET, AUDIT_HEADERS);
  ensureSimBrands_();
  formatDataSheet_(ss.getSheetByName(CUSTOMER_SHEET), CUSTOMER_HEADERS.length);
  formatDataSheet_(ss.getSheetByName(TRANSACTION_SHEET), TRANSACTION_HEADERS.length);
  formatDataSheet_(ss.getSheetByName(SIM_INVENTORY_SHEET), SIM_HEADERS.length);
  formatDataSheet_(ss.getSheetByName(STOCK_IN_SHEET), STOCK_HEADERS.length);
  formatDataSheet_(ss.getSheetByName(AUDIT_SHEET), AUDIT_HEADERS.length);
  appendAudit_(ss, 'DATABASE_INITIALIZED', SPREADSHEET_ID, TRANSACTION_SHEET, 1, 'Schema', '', 'Normalized schema initialized');
  return {
    success: true,
    spreadsheet: ss.getName(),
    spreadsheetUrl: ss.getUrl(),
    timeZone: ss.getSpreadsheetTimeZone(),
    customers: Math.max(0, customerSheet.getLastRow() - 1),
    transactions: Math.max(0, ss.getSheetByName(TRANSACTION_SHEET).getLastRow() - 1),
    sheets: [CUSTOMER_SHEET, TRANSACTION_SHEET, SIM_INVENTORY_SHEET, STOCK_IN_SHEET, AUDIT_SHEET]
  };
}

function assertHeaders_(sheet, expected, allowEmpty) {
  const lastColumn = Math.max(sheet.getLastColumn(), expected.length);
  const actual = sheet.getRange(1, 1, 1, Math.min(lastColumn, expected.length)).getDisplayValues()[0].map(clean_);
  const empty = actual.every(v => !v);
  if (empty && allowEmpty) {
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    return;
  }
  if (expected.some((h, i) => actual[i] !== h)) {
    throw new Error('Unexpected headers in ' + sheet.getName() + '. Expected: ' + expected.join(' | ') + '. No rows were changed.');
  }
}

function ensureDataSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  assertHeaders_(sheet, headers, true);
  return sheet;
}

function formatDataSheet_(sheet, width) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, width)
    .setFontWeight('bold').setBackground('#071a3b').setFontColor('#ffffff')
    .setWrap(true);
  if (!sheet.getFilter() && sheet.getLastRow() > 1) {
    sheet.getRange(1, 1, sheet.getLastRow(), width).createFilter();
  }
}

function getCustomerSheet_() {
  const sheet = getSpreadsheet_().getSheetByName(CUSTOMER_SHEET);
  if (!sheet) throw new Error('Customers sheet not found. Run initializePortalDatabase().');
  assertHeaders_(sheet, CUSTOMER_HEADERS, false);
  return sheet;
}
function getTransactionsSheet_() {
  const ss = getSpreadsheet_();
  return ensureDataSheet_(ss, TRANSACTION_SHEET, TRANSACTION_HEADERS);
}

function getCustomers_() {
  const sh = getCustomerSheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, CUSTOMER_HEADERS.length).getDisplayValues()
    .map((r, i) => ({
      row: i + 2,
      customerId: clean_(r[0]),
      owner: clean_(r[1]),
      barangay: clean_(r[2]),
      number: clean_(r[3]),
      customerType: clean_(r[4])
    }))
    .filter(c => c.customerId && (c.owner || c.number));
}
function customerMap_() {
  const map = {};
  getCustomers_().forEach(c => { map[c.customerId] = c; });
  return map;
}

function searchCustomers(query) {
  const q = clean_(query).toLowerCase();
  if (q.length < 2) return [];
  return getCustomers_().filter(c => [c.customerId, c.owner, c.barangay, c.number, c.customerType].join(' ').toLowerCase().includes(q)).slice(0, 40);
}

function validateCustomer_(customer) {
  if (!customer) throw new Error('Pumili muna ng customer.');
  const customerId = clean_(customer.customerId || customer.id);
  if (!customerId) throw new Error('Missing Customer_ID. Mag-search ulit.');
  const sh = getCustomerSheet_();
  if (sh.getLastRow() < 2) throw new Error('No customer records are available. Import the Customers sheet first.');
  const found = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().flat();
  const index = found.findIndex(id => clean_(id) === customerId);
  if (index < 0) throw new Error('Customer record not found. Mag-search ulit.');
  const row = index + 2;
  const r = sh.getRange(row, 1, 1, CUSTOMER_HEADERS.length).getDisplayValues()[0].map(clean_);
  return { row, customerId: r[0], owner: r[1], barangay: r[2], number: r[3], customerType: r[4] };
}

function getAuditSheet_(ss) {
  return ensureDataSheet_(ss, AUDIT_SHEET, AUDIT_HEADERS);
}
function appendAudit_(ss, action, entityId, sheetName, rowNumber, field, oldValue, newValue) {
  const sh = getAuditSheet_(ss);
  let actor = '';
  try { actor = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  sh.appendRow([new Date(), actor, action, entityId || '', sheetName || '', rowNumber || '', field || '',
    oldValue == null ? '' : String(oldValue), newValue == null ? '' : String(newValue)]);
  sh.getRange(sh.getLastRow(), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

function getSimInventorySheet_(ss) {
  const sh = ensureDataSheet_(ss, SIM_INVENTORY_SHEET, SIM_HEADERS);
  ensureSimBrands_();
  return sh;
}
function ensureSimBrands_() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SIM_INVENTORY_SHEET) || ensureDataSheet_(ss, SIM_INVENTORY_SHEET, SIM_HEADERS);
  const existing = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().flat().map(clean_) : [];
  SIM_BRANDS.forEach(brand => { if (!existing.includes(brand)) sh.appendRow([brand, 0, '']); });
}
function getStockInRows_() {
  const ss = getSpreadsheet_();
  const sh = ensureDataSheet_(ss, STOCK_IN_SHEET, STOCK_HEADERS);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, STOCK_HEADERS.length).getValues();
}

function getSimInventoryData_() {
  const ss = getSpreadsheet_();
  const sh = getSimInventorySheet_(ss);
  const stockInByBrand = {};
  getStockInRows_().forEach(r => {
    const brand = clean_(r[2]);
    stockInByBrand[brand] = (stockInByBrand[brand] || 0) + (Number(r[3]) || 0);
  });
  const baseRows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues() : [];
  const base = {};
  baseRows.forEach(r => { base[clean_(r[0])] = Number(r[1]) || 0; });
  const soldByBrand = {};
  getAllTransactions_().forEach(t => {
    if (t.type === 'Sim Card') soldByBrand[t.item] = (soldByBrand[t.item] || 0) + Number(t.quantity || 0);
  });
  return SIM_BRANDS.map(brand => {
    const beginning = base[brand] || 0;
    const stockIn = stockInByBrand[brand] || 0;
    const sold = soldByBrand[brand] || 0;
    return { brand, beginning, stockIn, sold, ending: beginning + stockIn - sold };
  });
}
function getSimInventorySummary() { return { success: true, items: getSimInventoryData_() }; }

function setSimBeginningInventory(brand, quantity) {
  const qty = Number(quantity);
  if (!SIM_BRANDS.includes(brand) || !Number.isInteger(qty) || qty < 0 || qty > 1000000) {
    throw new Error('Maglagay ng valid na beginning inventory (whole number, zero o higit pa).');
  }
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const current = getSimInventoryData_().find(x => x.brand === brand);
    if (current && qty + current.stockIn < current.sold) throw new Error('Hindi puwedeng mas mababa sa bilang ng SIM na naibenta ang Beginning Inventory.');
    const sh = getSimInventorySheet_(getSpreadsheet_());
    const brands = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).getDisplayValues().flat().map(clean_);
    const index = brands.indexOf(brand);
    if (index < 0) throw new Error('Hindi makita ang SIM brand.');
    const cell = sh.getRange(index + 2, 2);
    const oldValue = Number(cell.getValue()) || 0;
    cell.setValue(qty);
    appendAudit_(getSpreadsheet_(), 'SIM_BEGINNING_INVENTORY_CHANGED', brand, sh.getName(), index + 2, 'Beginning_Inventory', oldValue, qty);
    return getSimInventorySummary();
  } finally { lock.releaseLock(); }
}

function addSimStockIn(brand, quantity, requestId) {
  const qty = Number(quantity);
  if (!SIM_BRANDS.includes(brand) || !Number.isInteger(qty) || qty < 1 || qty > 1000000) {
    throw new Error('Maglagay ng valid na Stock In quantity (whole number na higit sa zero).');
  }
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const sh = ensureDataSheet_(ss, STOCK_IN_SHEET, STOCK_HEADERS);
    const stockInId = requestId ? 'STOCK-' + clean_(requestId) : 'STOCK-' + Utilities.getUuid();
    if (sh.getLastRow() > 1) {
      const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().flat();
      if (ids.includes(stockInId)) return getSimInventorySummary();
    }
    sh.appendRow([stockInId, new Date(), brand, qty, 'Portal stock-in']);
    sh.getRange(sh.getLastRow(), 2).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    appendAudit_(ss, 'SIM_STOCK_IN', stockInId, sh.getName(), sh.getLastRow(), 'Quantity', '', brand + ': ' + qty);
    return getSimInventorySummary();
  } finally { lock.releaseLock(); }
}

function transactionIdFromRequest_(requestId) {
  const safe = clean_(requestId).replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 90);
  return safe ? 'TXN-' + safe : 'TXN-' + Utilities.getUuid();
}
function findTransactionById_(id) {
  const sh = getTransactionsSheet_();
  if (sh.getLastRow() < 2) return null;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().flat();
  const index = ids.findIndex(v => clean_(v) === clean_(id));
  if (index < 0) return null;
  const row = sh.getRange(index + 2, 1, 1, TRANSACTION_HEADERS.length).getValues()[0];
  const record = formatStoredTransaction_(row, index + 2);
  record.success = true;
  record.sheet = TRANSACTION_SHEET;
  return record;
}

function saveSale_(customer, sale, requestId) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = getSpreadsheet_();
    const id = transactionIdFromRequest_(requestId);
    const duplicate = findTransactionById_(id);
    if (duplicate) return duplicate;
    const person = validateCustomer_(customer);
    const unitPrice = Number(sale.unitPrice), quantity = Number(sale.quantity);
    if (!Number.isFinite(unitPrice) || unitPrice <= 0 || unitPrice > 100000 || !Number.isInteger(quantity) || quantity < 1 || quantity > 100000) {
      throw new Error('Maglagay ng valid na unit price at whole-number quantity.');
    }
    if (sale.type === 'Sim Card') {
      if (!SIM_BRANDS.includes(sale.item)) throw new Error('Pumili ng valid na SIM brand.');
      const stock = getSimInventoryData_().find(x => x.brand === sale.item);
      if (!stock || stock.ending < quantity) throw new Error('Kulang ang ' + sale.item + ' SIM stock. Available: ' + (stock ? stock.ending : 0) + ' pcs.');
    }
    const allowedStatuses = PAYMENT_STATUSES;
    const paymentStatus = allowedStatuses.includes(sale.paymentStatus) ? sale.paymentStatus : 'Cash';
    const extraFee = Number(sale.extraFee || 0);
    if (!Number.isFinite(extraFee) || extraFee < 0 || extraFee > 100000) throw new Error('Invalid additional fee.');
    const total = round2_(unitPrice * quantity + extraFee);
    const now = new Date();
    const serviceType = sale.type === 'Simcard' ? 'Sim Card' : (sale.type === 'Load' ? 'Amax' : sale.type);
    const simBrand = serviceType === 'Sim Card' ? sale.item : '';
    const gcashType = serviceType === 'GCash' ? (sale.gcashType || 'GCash') : '';
    const smsCommand = serviceType === 'Amax' ? 'LoadRT/' + person.number.replace(/\D/g, '') + '/' + String(unitPrice) : '';
    const notes = extraFee ? 'Service fee: ' + extraFee.toFixed(2) : '';
    const sh = getTransactionsSheet_();
    sh.appendRow([id, now, person.customerId, serviceType, simBrand, gcashType, quantity, unitPrice, total, paymentStatus, smsCommand, notes]);
    const row = sh.getLastRow();
    sh.getRange(row, 2).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sh.getRange(row, 8, 1, 2).setNumberFormat('"₱"#,##0.00');
    setTransactionRowColor_(sh, row, paymentStatus);
    appendAudit_(ss, 'TRANSACTION_CREATED', id, sh.getName(), row, 'Transaction', '', JSON.stringify({ customerId: person.customerId, serviceType, simBrand, quantity, total, paymentStatus }));
    return {
      success: true, id, date: Utilities.formatDate(now, PORTAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss'),
      reportDateKey: dateKey_(now), sheet: TRANSACTION_SHEET, customerId: person.customerId,
      barangay: person.barangay, owner: person.owner, number: person.number, type: serviceType,
      item: simBrand || (serviceType === 'GCash' ? gcashType : serviceType), priceType: sale.priceType || '',
      unitPrice, quantity, paymentStatus, amount: total, totalSales: total, serviceFee: extraFee,
      smsCommand
    };
  } finally { lock.releaseLock(); }
}

function saveTransaction(customer, amount, paymentStatus, requestId) {
  const value = Number(amount);
  if (!Number.isFinite(value) || value <= 0 || value > 100000) throw new Error('Maglagay ng valid na load amount.');
  return saveSale_(customer, { type: 'Amax', item: 'Amax', unitPrice: value, quantity: 1, paymentStatus }, requestId);
}
function saveGcashTransaction(customer, amount, paymentStatus, requestId) {
  const value = Number(amount);
  if (!Number.isFinite(value) || value <= 0 || value > 100000) throw new Error('Maglagay ng valid na GCash amount.');
  return saveSale_(customer, { type: 'GCash', item: 'GCash', gcashType: 'GCash', unitPrice: value, quantity: 1,
    extraFee: round2_(value * 0.01), paymentStatus }, requestId);
}
function saveSimcardTransaction(customer, brand, unitPrice, quantity, paymentStatus, requestId) {
  if (!SIM_BRANDS.includes(brand)) throw new Error('Pumili ng valid na SIM brand.');
  return saveSale_(customer, { type: 'Sim Card', item: brand, unitPrice, quantity, priceType: 'Manual', paymentStatus }, requestId);
}

function getTransactionCustomerMap_() { return customerMap_(); }
function formatStoredTransaction_(r, rowNumber) {
  const customers = getTransactionCustomerMap_();
  const customer = customers[clean_(r[2])] || { customerId: clean_(r[2]), owner: '[Customer not found]', barangay: '', number: '' };
  const stamp = r[1] instanceof Date ? r[1] : new Date(r[1]);
  const validStamp = stamp && !isNaN(stamp.getTime());
  const date = validStamp ? Utilities.formatDate(stamp, PORTAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss') : clean_(r[1]);
  const type = clean_(r[3]);
  const item = type === 'Sim Card' ? clean_(r[4]) : (type === 'GCash' ? (clean_(r[5]) || 'GCash') : 'Amax');
  return {
    id: clean_(r[0]), timestamp: validStamp ? stamp : null,
    reportDateKey: validStamp ? Utilities.formatDate(stamp, PORTAL_TIME_ZONE, 'yyyy-MM-dd') : '',
    date, customerId: clean_(r[2]), sourceSheet: TRANSACTION_SHEET,
    barangay: customer.barangay, owner: customer.owner, number: customer.number,
    type, item, priceType: '', unitPrice: Number(r[7]) || 0, quantity: Number(r[6]) || 0,
    paymentStatus: clean_(r[9]) || 'Cash', amount: Number(r[8]) || 0,
    smsCommand: clean_(r[10]), notes: clean_(r[11]), rowNumber: rowNumber || 0
  };
}
function getAllTransactions_() {
  const sh = getTransactionsSheet_();
  if (sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, TRANSACTION_HEADERS.length).getValues();
  const customers = customerMap_();
  return rows.map((r, i) => {
    const customer = customers[clean_(r[2])] || { customerId: clean_(r[2]), owner: '[Customer not found]', barangay: '', number: '' };
    const stamp = r[1] instanceof Date ? r[1] : new Date(r[1]);
    const validStamp = stamp && !isNaN(stamp.getTime());
    const type = clean_(r[3]);
    return {
      id: clean_(r[0]), timestamp: validStamp ? stamp : null,
      reportDateKey: validStamp ? Utilities.formatDate(stamp, PORTAL_TIME_ZONE, 'yyyy-MM-dd') : '',
      date: validStamp ? Utilities.formatDate(stamp, PORTAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss') : clean_(r[1]),
      customerId: clean_(r[2]), barangay: customer.barangay, owner: customer.owner, number: customer.number,
      type, item: type === 'Sim Card' ? clean_(r[4]) : (type === 'GCash' ? (clean_(r[5]) || 'GCash') : 'Amax'),
      priceType: '', unitPrice: Number(r[7]) || 0, quantity: Number(r[6]) || 0,
      paymentStatus: clean_(r[9]) || 'Cash', amount: Number(r[8]) || 0,
      smsCommand: clean_(r[10]), notes: clean_(r[11]), rowNumber: i + 2
    };
  }).filter(t => t.id);
}
function formatTransaction_(t) {
  return { id: t.id, date: t.date, reportDateKey: t.reportDateKey || '', customerId: t.customerId,
    barangay: t.barangay, owner: t.owner, number: t.number, type: t.type, item: t.item,
    priceType: t.priceType || '', unitPrice: t.unitPrice, quantity: t.quantity,
    paymentStatus: t.paymentStatus || 'Cash', amount: t.amount, smsCommand: t.smsCommand || '' };
}

function getPeriodReport(period, selectedDate) {
  const todayKey = dateKey_(new Date());
  let startKey, endKey;
  if (period === 'date') {
    const key = clean_(selectedDate || todayKey);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(key)) throw new Error('Invalid report date.');
    startKey = key; endKey = key;
  } else {
    const p = todayKey.split('-').map(Number);
    const todayUtc = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
    let start = new Date(todayUtc);
    if (period === 'week') {
      const day = start.getUTCDay();
      start.setUTCDate(start.getUTCDate() - ((day + 6) % 7));
    } else if (period === 'month') {
      start = new Date(Date.UTC(p[0], p[1] - 1, 1));
    } else {
      period = 'day';
    }
    const end = new Date(todayUtc); end.setUTCDate(end.getUTCDate() + 1);
    startKey = Utilities.formatDate(start, 'UTC', 'yyyy-MM-dd');
    endKey = Utilities.formatDate(end, 'UTC', 'yyyy-MM-dd');
  }
  const inclusiveEnd = period === 'date' ? endKey : Utilities.formatDate(new Date(Date.parse(endKey + 'T00:00:00Z') - 86400000), 'UTC', 'yyyy-MM-dd');
  const all = getAllTransactions_();
  const transactions = all.filter(t => t.reportDateKey >= startKey && t.reportDateKey <= inclusiveEnd)
    .sort((a, b) => (b.timestamp ? b.timestamp.getTime() : 0) - (a.timestamp ? a.timestamp.getTime() : 0));
  const summaries = { Amax: { count: 0, amount: 0 }, 'Sim Card': { count: 0, amount: 0 }, GCash: { count: 0, amount: 0 } };
  const breakdown = {};
  const outstanding = transactions.filter(t => ['Credit', 'Pending'].includes(t.paymentStatus));
  transactions.forEach(t => {
    const category = t.type === 'Amax' ? 'Amax' : t.type;
    if (!summaries[category]) summaries[category] = { count: 0, amount: 0 };
    summaries[category].count++;
    summaries[category].amount += t.amount;
    const key = [t.type, t.item].filter(Boolean).join(' • ');
    if (!breakdown[key]) breakdown[key] = { label: key, quantity: 0, transactions: 0, totalSales: 0 };
    breakdown[key].quantity += t.quantity;
    breakdown[key].transactions++;
    breakdown[key].totalSales += t.amount;
  });
  Object.keys(summaries).forEach(k => summaries[k].amount = money_(summaries[k].amount));
  return {
    period, summaries, startDate: startKey, endDate: inclusiveEnd,
    totalCount: transactions.length,
    totalAmount: money_(transactions.reduce((sum, t) => sum + t.amount, 0)),
    outstandingCount: outstanding.length,
    outstandingAmount: money_(outstanding.reduce((sum, t) => sum + t.amount, 0)),
    items: Object.keys(breakdown).map(k => ({ ...breakdown[k], totalSales: money_(breakdown[k].totalSales) })),
    transactions: transactions.map(formatTransaction_)
  };
}
function getReportForDate(dateKey) { return getPeriodReport('date', dateKey); }
function getDailySummary() { return getPeriodReport('day'); }

function getCustomerTransactionHistory(customerNumber, transactionDate) {
  const target = clean_(customerNumber);
  const date = clean_(transactionDate);
  if (!target || !/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('A customer number and valid transaction date are required.');
  return getAllTransactions_().filter(t => t.number === target && t.reportDateKey === date)
    .sort((a, b) => (b.timestamp ? b.timestamp.getTime() : 0) - (a.timestamp ? a.timestamp.getTime() : 0))
    .map(formatTransaction_);
}

function setTransactionRowColor_(sheet, row, status) {
  let color = '#000000';
  if (status === 'Credit') color = '#b77900';
  else if (status === 'Pending') color = '#c2410c';
  else if (status === 'Gcash Payment') color = '#0891b2';
  sheet.getRange(row, 1, 1, TRANSACTION_HEADERS.length).setFontColor(color);
}
function updateTransactionPaymentStatus(transactionId, paymentStatus) {
  if (!clean_(transactionId) || !PAYMENT_STATUSES.includes(paymentStatus)) throw new Error('Invalid transaction or payment status.');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const sh = getTransactionsSheet_();
    if (sh.getLastRow() < 2) throw new Error('Transaction not found. Refresh the report and try again.');
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().flat();
    const index = ids.findIndex(id => clean_(id) === clean_(transactionId));
    if (index < 0) throw new Error('Transaction not found. Refresh the report and try again.');
    const row = index + 2;
    const cell = sh.getRange(row, 10);
    const oldStatus = clean_(cell.getValue()) || 'Cash';
    cell.setValue(paymentStatus);
    setTransactionRowColor_(sh, row, paymentStatus);
    if (oldStatus !== paymentStatus) appendAudit_(ss, 'PAYMENT_STATUS_CHANGED', transactionId, sh.getName(), row, 'Payment_Status', oldStatus, paymentStatus);
    return { success: true, id: transactionId, paymentStatus };
  } finally { lock.releaseLock(); }
}

function getRecentTransactions() {
  return getAllTransactions_().sort((a, b) => (b.timestamp ? b.timestamp.getTime() : 0) - (a.timestamp ? a.timestamp.getTime() : 0))
    .slice(0, 15).map(t => {
      const x = formatTransaction_(t);
      x.amount = '₱' + Number(t.amount || 0).toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      return x;
    });
}
function getTransactionSummary() {
  const rows = getAllTransactions_();
  return { totalCount: rows.length, totalAmount: money_(rows.reduce((sum, t) => sum + t.amount, 0)) };
}
function getTransactionIntegrityReport() {
  const ss = getSpreadsheet_();
  const rows = getAllTransactions_();
  const ids = {}, duplicateIds = [], orphanCustomerIds = [];
  const validCustomers = customerMap_();
  rows.forEach(t => {
    if (ids[t.id]) duplicateIds.push({ id: t.id, rows: [ids[t.id], t.rowNumber] });
    else ids[t.id] = t.rowNumber;
    if (!validCustomers[t.customerId]) orphanCustomerIds.push({ id: t.id, customerId: t.customerId });
  });
  return { transactionRows: rows.length, uniqueTransactionIds: Object.keys(ids).length,
    duplicateIdCount: duplicateIds.length, duplicateIds, orphanCustomerCount: orphanCustomerIds.length,
    orphanCustomerIds, customerCount: Math.max(0, getCustomerSheet_().getLastRow() - 1), timeZone: ss.getSpreadsheetTimeZone() };
}

/** Read-only local cache payload. Capped to avoid oversized browser storage. */
function getOfflineBootstrap() {
  const customers = getCustomers_();
  const transactions = getAllTransactions_().sort((a, b) => (b.timestamp ? b.timestamp.getTime() : 0) - (a.timestamp ? a.timestamp.getTime() : 0))
    .slice(0, MAX_OFFLINE_TRANSACTIONS).map(formatTransaction_);
  return { success: true, timeZone: PORTAL_TIME_ZONE, fetchedAt: new Date().toISOString(), customers, transactions, inventory: getSimInventoryData_() };
}
function getDatabaseUrl() { return 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/edit'; }
function getPortalStatus() {
  const ss = getSpreadsheet_();
  return { success: true, spreadsheetName: ss.getName(), spreadsheetUrl: ss.getUrl(), timeZone: ss.getSpreadsheetTimeZone(), date: dateKey_(new Date()), schema: 'normalized-v2' };
}

// Optional backup protection. Run setupDatabaseProtection() once and authorize Drive access.
function getBackupFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PORTAL_BACKUP_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);
  props.setProperty('PORTAL_BACKUP_FOLDER_ID', folder.getId());
  return folder;
}
function setupDatabaseProtection() {
  const ss = getSpreadsheet_();
  getAuditSheet_(ss);
  const folder = getBackupFolder_();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'runScheduledDatabaseBackup').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('runScheduledDatabaseBackup').timeBased().everyDays(1).atHour(2).create();
  const first = createDatabaseBackup_();
  appendAudit_(ss, 'BACKUP_SETUP', '', '', '', 'Schedule', '', 'Daily backup configured; retention ' + BACKUP_RETENTION_DAYS + ' days');
  return { success: true, message: 'Daily backup installed. Initial backup: ' + first.name, folderUrl: folder.getUrl() };
}
function runScheduledDatabaseBackup() { return createDatabaseBackup_(); }
function createDatabaseBackup_() {
  const ss = getSpreadsheet_(), folder = getBackupFolder_();
  const stamp = Utilities.formatDate(new Date(), PORTAL_TIME_ZONE, 'yyyy-MM-dd_HHmmss');
  const name = BACKUP_PREFIX + stamp;
  const copy = DriveApp.getFileById(SPREADSHEET_ID).makeCopy(name, folder);
  const cutoff = Date.now() - BACKUP_RETENTION_DAYS * 86400000;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getId() !== copy.getId() && file.getName().indexOf(BACKUP_PREFIX) === 0 && file.getDateCreated().getTime() < cutoff) file.setTrashed(true);
  }
  appendAudit_(ss, 'DATABASE_BACKUP', '', '', '', 'Backup_File', '', name + ' | ' + copy.getUrl());
  return { success: true, name, url: copy.getUrl() };
}
