/**
 * Biozed Finance Manager — PASTE THIS ENTIRE FILE into Apps Script Code.gs
 * (replace the empty myFunction). Then Run → setupFinanceManager
 */


// ========== Config.gs ==========

/**
 * Biozed Finance Manager — Configuration
 * Version 1.0
 */

var CONFIG = {
  OWNER: 'BIOZED HOSSAIN',
  APP_NAME: 'Personal Finance Manager',
  VERSION: '1.0',
  CURRENCY: 'BDT',
  CURRENCY_FORMAT: '৳#,##0.00',
  DATE_FORMAT: 'dd MMM yyyy',
  EMERGENCY_BALANCE: 1000,
  BUDGET_ALERT_PCT: 0.2,
  DEFAULT_MONTHLY_INCOME: 30000,
  FOOTER_LINE_1: 'Money is a tool.',
  FOOTER_LINE_2: 'Every Taka must have a purpose.',
  FOOTER_LINE_3: 'Spend intentionally. Save consistently.',
  BUILT_FOR: 'Built for Biozed Hossain'
};

var SHEETS = {
  DASHBOARD: 'Dashboard',
  INCOME: 'Income',
  EXPENSE: 'Expense Tracker',
  BUDGET: 'Monthly Budget',
  INSTALLMENT: 'Installment Tracker',
  DEBT: 'Debt Tracker',
  SAVINGS: 'Savings',
  REPORTS: 'Reports',
  SETTINGS: 'Settings',
  CATEGORIES: 'Categories',
  CHART_DATA: '_ChartData'
};

var INCOME_SOURCES = [
  'Salary',
  'Freelance',
  'Bonus',
  'Gift',
  'Refund',
  'Other'
];

var PAYMENT_METHODS = [
  'Cash',
  'bKash',
  'Nagad',
  'Rocket',
  'Bank',
  'Card',
  'Other'
];

var EXPENSE_CATEGORIES = [
  'Market',
  'Rent',
  'Electricity',
  'Gas',
  'Internet',
  'Transport',
  'Medical',
  'Education',
  'Kids',
  'Clothing',
  'Restaurant',
  'Shopping',
  'Entertainment',
  'Gift',
  'Installment',
  'Loan',
  'Emergency',
  'Investment',
  'Business',
  'Office',
  'Other'
];

var SUB_CATEGORIES = {
  Market: ['Grocery', 'Vegetables', 'Fruits', 'Meat/Fish', 'Household', 'Other'],
  Rent: ['House Rent', 'Office Rent', 'Other'],
  Electricity: ['Home Bill', 'Office Bill', 'Other'],
  Gas: ['Cylinder', 'Pipeline', 'Other'],
  Internet: ['Home', 'Mobile Data', 'Other'],
  Transport: ['Bus', 'CNG', 'Uber/Pathao', 'Fuel', 'Other'],
  Medical: ['Doctor', 'Medicine', 'Hospital', 'Other'],
  Education: ['Tuition', 'Books', 'Fees', 'Other'],
  Kids: ['Food', 'Clothes', 'Toys', 'School', 'Other'],
  Clothing: ['Self', 'Family', 'Other'],
  Restaurant: ['Lunch', 'Dinner', 'Cafe', 'Other'],
  Shopping: ['Online', 'Store', 'Other'],
  Entertainment: ['Movie', 'Outing', 'Subscription', 'Other'],
  Gift: ['Family', 'Friends', 'Occasion', 'Other'],
  Installment: ['Phone', 'Laptop', 'Appliance', 'Other'],
  Loan: ['Personal', 'Bank', 'Family', 'Other'],
  Emergency: ['Medical', 'Repair', 'Other'],
  Investment: ['Savings', 'Stock', 'Crypto', 'Other'],
  Business: ['Supplies', 'Marketing', 'Other'],
  Office: ['Supplies', 'Snacks', 'Other'],
  Other: ['Misc']
};

/** Default monthly budgets (BDT) — edit on Monthly Budget sheet */
var DEFAULT_BUDGETS = {
  Market: 15000,
  Rent: 8000,
  Electricity: 1500,
  Gas: 1200,
  Internet: 1000,
  Transport: 3000,
  Medical: 2000,
  Education: 2000,
  Kids: 3000,
  Clothing: 1500,
  Restaurant: 2000,
  Shopping: 2000,
  Entertainment: 1500,
  Gift: 1000,
  Installment: 5000,
  Loan: 0,
  Emergency: 1000,
  Investment: 2000,
  Business: 0,
  Office: 500,
  Other: 1000
};

var COLORS = {
  HEADER_BG: '#0f172a',
  HEADER_FG: '#ffffff',
  ACCENT: '#2563eb',
  INCOME: '#16a34a',
  EXPENSE: '#dc2626',
  BALANCE: '#0369a1',
  BUDGET: '#7c3aed',
  OVER: '#dc2626',
  SAVINGS: '#059669',
  INSTALLMENT: '#ea580c',
  DEBT: '#2563eb',
  WARN: '#f59e0b',
  LIGHT: '#f8fafc',
  ROW_ALT: '#f1f5f9',
  GREEN_BG: '#dcfce7',
  RED_BG: '#fee2e2',
  YELLOW_BG: '#fef9c3',
  ORANGE_BG: '#ffedd5',
  BLUE_BG: '#dbeafe',
  WHITE: '#ffffff',
  BORDER: '#cbd5e1',
  MUTED: '#64748b'
};

var HEALTH = {
  EXCELLENT: { min: 90, label: 'Excellent', color: '#16a34a', emoji: '🟢' },
  GOOD: { min: 70, label: 'Good', color: '#ca8a04', emoji: '🟡' },
  ATTENTION: { min: 50, label: 'Needs Attention', color: '#ea580c', emoji: '🟠' },
  CRITICAL: { min: 0, label: 'Critical', color: '#dc2626', emoji: '🔴' }
};


// ========== Code.gs ==========

/**
 * Biozed Finance Manager — Main Entry & Menu
 * Open Extensions → Apps Script, paste all .gs files, then run setupFinanceManager
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 Finance Manager')
    .addItem('🚀 Setup / Rebuild Workbook', 'setupFinanceManager')
    .addItem('🔄 Refresh Dashboard', 'refreshDashboard')
    .addItem('📊 Rebuild Charts', 'rebuildCharts')
    .addItem('📝 Add Sample Data', 'addSampleData')
    .addSeparator()
    .addItem('📅 Set Active Month to Current', 'setActiveMonthToCurrent')
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();
}

/**
 * One-click setup: creates all sheets, formulas, formatting, charts.
 * Safe to re-run — preserves Income / Expense / Savings / Debt / Installment data rows.
 */
function setupFinanceManager() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Building Biozed Finance Manager…', 'Setup', 5);

  ensureAllSheets_(ss);
  buildCategoriesSheet_(ss);
  buildSettingsSheet_(ss);
  buildIncomeSheet_(ss);
  buildExpenseSheet_(ss);
  buildBudgetSheet_(ss);
  buildInstallmentSheet_(ss);
  buildDebtSheet_(ss);
  buildSavingsSheet_(ss);
  buildReportsSheet_(ss);
  buildChartDataSheet_(ss);
  buildDashboard_(ss);

  applyAllConditionalFormatting_(ss);
  installNamedRanges_(ss);
  rebuildCharts();
  refreshDashboard();

  ss.setActiveSheet(ss.getSheetByName(SHEETS.DASHBOARD));
  ss.toast('Setup complete. Enter data on Income & Expense Tracker.', 'Done ✓', 8);
  SpreadsheetApp.getUi().alert(
    'Biozed Finance Manager v' + CONFIG.VERSION,
    'Workbook is ready.\n\n' +
      '1. Open Settings — set Monthly Income & Fixed Costs\n' +
      '2. Open Monthly Budget — adjust category budgets\n' +
      '3. Log income on Income sheet\n' +
      '4. Log expenses on Expense Tracker\n' +
      '5. Dashboard updates automatically (or use 💰 Finance Manager → Refresh)\n\n' +
      'Financial Health Score appears on the Dashboard.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function showAbout() {
  SpreadsheetApp.getUi().alert(
    CONFIG.OWNER + ' — ' + CONFIG.APP_NAME,
    CONFIG.FOOTER_LINE_1 + '\n' +
      CONFIG.FOOTER_LINE_2 + '\n' +
      CONFIG.FOOTER_LINE_3 + '\n\n' +
      CONFIG.BUILT_FOR + '\nVersion ' + CONFIG.VERSION,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function setActiveMonthToCurrent() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = ss.getSheetByName(SHEETS.SETTINGS);
  if (!settings) {
    SpreadsheetApp.getUi().alert('Run Setup first.');
    return;
  }
  var now = new Date();
  settings.getRange('B3').setValue(now.getFullYear());
  settings.getRange('B4').setValue(now.getMonth() + 1);
  refreshDashboard();
  ss.toast('Active month set to ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM yyyy'), 'Month', 4);
}

/** Recalculate on edit of key sheets */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    var name = sheet.getName();
    var watch = [
      SHEETS.INCOME,
      SHEETS.EXPENSE,
      SHEETS.BUDGET,
      SHEETS.INSTALLMENT,
      SHEETS.DEBT,
      SHEETS.SAVINGS,
      SHEETS.SETTINGS
    ];
    if (watch.indexOf(name) === -1) return;

    if (name === SHEETS.EXPENSE && e.range.getColumn() === 2 && e.range.getRow() >= 2) {
      setSubCategoryValidation_(sheet, e.range.getRow());
    }

    // Light debounce via PropertiesService timestamp
    var props = PropertiesService.getDocumentProperties();
    var last = Number(props.getProperty('lastRefresh') || 0);
    var now = Date.now();
    if (now - last < 2000) return;
    props.setProperty('lastRefresh', String(now));

    refreshDashboardLight_();
  } catch (err) {
    // Never block user edits
    console.error(err);
  }
}


// ========== Sheets.gs ==========

/**
 * Sheet creation helpers & data-entry sheets
 */

function ensureAllSheets_(ss) {
  var order = [
    SHEETS.DASHBOARD,
    SHEETS.INCOME,
    SHEETS.EXPENSE,
    SHEETS.BUDGET,
    SHEETS.INSTALLMENT,
    SHEETS.DEBT,
    SHEETS.SAVINGS,
    SHEETS.REPORTS,
    SHEETS.SETTINGS,
    SHEETS.CATEGORIES,
    SHEETS.CHART_DATA
  ];
  order.forEach(function (name) {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });
  // Reorder
  order.forEach(function (name, i) {
    var sh = ss.getSheetByName(name);
    if (sh) ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
  });
  // Hide chart helper
  var chart = ss.getSheetByName(SHEETS.CHART_DATA);
  if (chart) chart.hideSheet();
}

function styleHeaderRow_(sheet, cols, bg) {
  var range = sheet.getRange(1, 1, 1, cols);
  range
    .setBackground(bg || COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_FG)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 32);
}

function clearDataKeepHeader_(sheet, startRow, numCols) {
  var last = sheet.getLastRow();
  if (last >= startRow) {
    sheet.getRange(startRow, 1, last - startRow + 1, numCols).clearContent();
  }
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/* ===================== CATEGORIES ===================== */

function buildCategoriesSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.CATEGORIES);
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([['Category', 'Sub Category', 'Income Source', 'Payment Method']]);
  styleHeaderRow_(sh, 4);

  var catRows = EXPENSE_CATEGORIES.map(function (c) { return [c]; });
  sh.getRange(2, 1, catRows.length, 1).setValues(catRows);

  var subRows = [];
  EXPENSE_CATEGORIES.forEach(function (cat) {
    var subs = SUB_CATEGORIES[cat] || ['Other'];
    subs.forEach(function (s) {
      subRows.push([cat, s]);
    });
  });
  // Put flat sub list for dependent dropdown helper in cols E-F
  sh.getRange(1, 5, 1, 2).setValues([['CatKey', 'SubKey']]);
  sh.getRange(1, 5, 1, 2).setBackground(COLORS.MUTED).setFontColor(COLORS.WHITE).setFontWeight('bold');
  if (subRows.length) {
    sh.getRange(2, 5, subRows.length, 2).setValues(subRows);
  }

  var incomeRows = INCOME_SOURCES.map(function (s) { return [s]; });
  sh.getRange(2, 3, incomeRows.length, 1).setValues(incomeRows);

  var payRows = PAYMENT_METHODS.map(function (p) { return [p]; });
  sh.getRange(2, 4, payRows.length, 1).setValues(payRows);

  sh.setColumnWidths(1, 6, 140);
  sh.getRange('A1:D' + Math.max(EXPENSE_CATEGORIES.length + 1, 22))
    .setBorder(true, true, true, true, true, true, COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
}

/* ===================== SETTINGS ===================== */

function buildSettingsSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.SETTINGS);
  var existing = {};
  try {
    existing.income = sh.getRange('B6').getValue();
    existing.fixed = sh.getRange('B7').getValue();
    existing.year = sh.getRange('B3').getValue();
    existing.month = sh.getRange('B4').getValue();
  } catch (e) {}

  sh.clear();
  sh.getRange('A1').setValue('⚙️ SETTINGS').setFontSize(16).setFontWeight('bold').setFontColor(COLORS.HEADER_BG);
  sh.getRange('A2').setValue('Change these values — Dashboard & formulas update automatically.');

  var now = new Date();
  var year = existing.year || now.getFullYear();
  var month = existing.month || (now.getMonth() + 1);
  var income = existing.income || CONFIG.DEFAULT_MONTHLY_INCOME;
  var fixed = existing.fixed || 0;

  var rows = [
    ['Active Year', year, 'e.g. 2026'],
    ['Active Month (1-12)', month, 'e.g. 7 = July'],
    ['Owner Name', CONFIG.OWNER, ''],
    ['Monthly Income (BDT)', income, 'Used for Safe Daily Budget'],
    ['Fixed Monthly Costs (BDT)', fixed, 'Rent + bills you always pay'],
    ['Emergency Balance Alert', CONFIG.EMERGENCY_BALANCE, 'Warn if balance below this'],
    ['Budget Alert Threshold %', CONFIG.BUDGET_ALERT_PCT * 100, 'Warn when remaining < this %'],
    ['Currency', CONFIG.CURRENCY, ''],
    ['Version', CONFIG.VERSION, '']
  ];
  sh.getRange(3, 1, rows.length, 3).setValues(rows);
  sh.getRange('A3:A11').setFontWeight('bold').setBackground(COLORS.LIGHT);
  sh.getRange('B3:B11').setBackground(COLORS.BLUE_BG);
  sh.getRange('B6').setNumberFormat(CONFIG.CURRENCY_FORMAT);
  sh.getRange('B7').setNumberFormat(CONFIG.CURRENCY_FORMAT);
  sh.getRange('B8').setNumberFormat(CONFIG.CURRENCY_FORMAT);

  sh.getRange('A13').setValue('Safe Daily Budget Formula').setFontWeight('bold');
  sh.getRange('A14').setValue(
    '=(Monthly Income − Fixed Costs − Sum of category budgets that are "fixed") / days in month\n' +
    'Implemented as: (B6 − B7 − discretionary buffer from remaining flexible budget) / days.\n' +
    'Dashboard uses: MAX(0, (Income − Fixed − Budgeted Variable) / DaysInMonth) with Income from Settings.'
  );
  sh.getRange('A14').setWrap(true);
  sh.setRowHeight(14, 60);

  sh.getRange('A16').setValue('Health Score Weights').setFontWeight('bold');
  sh.getRange(17, 1, 5, 2).setValues([
    ['Within Budget', 30],
    ['Daily Logging', 20],
    ['No Overspend', 20],
    ['Savings Deposits', 15],
    ['Installments On Time', 15]
  ]);

  sh.setColumnWidth(1, 240);
  sh.setColumnWidth(2, 160);
  sh.setColumnWidth(3, 280);
}

/* ===================== INCOME ===================== */

function buildIncomeSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.INCOME);
  var oldData = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues() : [];

  sh.clear();
  var headers = ['Date', 'Source', 'Description', 'Amount', 'Payment Method', 'Notes'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sh, headers.length, COLORS.INCOME);

  if (oldData.length && oldData[0][0] !== '') {
    sh.getRange(2, 1, oldData.length, 6).setValues(oldData);
  }

  // Validation
  var sourceRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(SHEETS.CATEGORIES).getRange('C2:C50'), true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('B2:B1000').setDataValidation(sourceRule);

  var payRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(SHEETS.CATEGORIES).getRange('D2:D50'), true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('E2:E1000').setDataValidation(payRule);

  sh.getRange('A2:A1000').setNumberFormat(CONFIG.DATE_FORMAT);
  sh.getRange('D2:D1000').setNumberFormat(CONFIG.CURRENCY_FORMAT);

  sh.setColumnWidths(1, 1, 110);
  sh.setColumnWidth(2, 120);
  sh.setColumnWidth(3, 220);
  sh.setColumnWidth(4, 110);
  sh.setColumnWidth(5, 130);
  sh.setColumnWidth(6, 200);

  // Totals row hint
  sh.getRange('H1').setValue('Month Total').setFontWeight('bold');
  sh.getRange('H2').setFormula(
    '=SUMIFS(D:D,A:A,">="&DATE(Settings!$B$3,Settings!$B$4,1),' +
    'A:A,"<"&EDATE(DATE(Settings!$B$3,Settings!$B$4,1),1))'
  ).setNumberFormat(CONFIG.CURRENCY_FORMAT).setBackground(COLORS.GREEN_BG);
}

/* ===================== EXPENSE TRACKER ===================== */

function buildExpenseSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.EXPENSE);
  var cols = 10;
  var oldData = [];
  if (sh.getLastRow() > 1) {
    oldData = sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues();
  }

  sh.clear();
  var headers = [
    'Date', 'Category', 'Sub Category', 'Description',
    'Budget', 'Actual', 'Difference', 'Status', 'Payment', 'Notes'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sh, headers.length, COLORS.EXPENSE);

  // Restore user-entered columns only (A-F, I-J); formulas in G-H
  if (oldData.length && oldData[0][0] !== '') {
    for (var r = 0; r < oldData.length; r++) {
      sh.getRange(r + 2, 1, 1, 6).setValues([[
        oldData[r][0], oldData[r][1], oldData[r][2], oldData[r][3], oldData[r][4], oldData[r][5]
      ]]);
      sh.getRange(r + 2, 9, 1, 2).setValues([[oldData[r][8], oldData[r][9]]]);
    }
  }

  // Formulas for first 500 rows
  for (var i = 2; i <= 501; i++) {
    // Difference = Budget - Actual
    sh.getRange('G' + i).setFormula(
      '=IF(OR(E' + i + '="",F' + i + '=""),"",E' + i + '-F' + i + ')'
    );
    // Status
    sh.getRange('H' + i).setFormula(
      '=IF(OR(E' + i + '="",F' + i + '=""),"",' +
      'IF(F' + i + '>E' + i + ',"🔴 OVER",' +
      'IF(F' + i + '<E' + i + ',"🟢 UNDER","🟡 PERFECT")))'
    );
  }

  var catRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(SHEETS.CATEGORIES).getRange('A2:A50'), true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('B2:B1000').setDataValidation(catRule);

  // Sub-category: free text + note (full dependent dropdown needs Apps Script onEdit)
  var payRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(SHEETS.CATEGORIES).getRange('D2:D50'), true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('I2:I1000').setDataValidation(payRule);

  sh.getRange('A2:A1000').setNumberFormat(CONFIG.DATE_FORMAT);
  sh.getRange('E2:G1000').setNumberFormat(CONFIG.CURRENCY_FORMAT);

  sh.setColumnWidth(1, 100);
  sh.setColumnWidth(2, 120);
  sh.setColumnWidth(3, 120);
  sh.setColumnWidth(4, 180);
  sh.setColumnWidth(5, 90);
  sh.setColumnWidth(6, 90);
  sh.setColumnWidth(7, 100);
  sh.setColumnWidth(8, 110);
  sh.setColumnWidth(9, 100);
  sh.setColumnWidth(10, 160);

  sh.getRange('L1').setValue('Month Total Actual').setFontWeight('bold');
  sh.getRange('L2').setFormula(
    '=SUMIFS(F:F,A:A,">="&DATE(Settings!$B$3,Settings!$B$4,1),' +
    'A:A,"<"&EDATE(DATE(Settings!$B$3,Settings!$B$4,1),1))'
  ).setNumberFormat(CONFIG.CURRENCY_FORMAT).setBackground(COLORS.RED_BG);
}

/** Dependent Sub Category dropdown from Categories map */
function setSubCategoryValidation_(sheet, row) {
  var cat = sheet.getRange(row, 2).getValue();
  var cell = sheet.getRange(row, 3);
  if (!cat) {
    cell.clearDataValidations();
    return;
  }
  var subs = SUB_CATEGORIES[cat] || ['Other'];
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(subs, true)
    .setAllowInvalid(true)
    .build();
  cell.setDataValidation(rule);
}

/* ===================== MONTHLY BUDGET ===================== */

function buildBudgetSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.BUDGET);
  // Preserve custom budget amounts if present
  var custom = {};
  if (sh.getLastRow() > 1) {
    var prev = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    prev.forEach(function (r) {
      if (r[0]) custom[r[0]] = r[1];
    });
  }

  sh.clear();
  var headers = ['Category', 'Monthly Budget', 'Used', 'Remaining', '% Used', 'Status'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sh, headers.length, COLORS.BUDGET);

  var rows = EXPENSE_CATEGORIES.map(function (cat, idx) {
    var row = idx + 2;
    var amt = custom[cat] != null ? custom[cat] : (DEFAULT_BUDGETS[cat] || 0);
    return [cat, amt];
  });
  sh.getRange(2, 1, rows.length, 2).setValues(rows);

  for (var i = 0; i < EXPENSE_CATEGORIES.length; i++) {
    var r = i + 2;
    // Used = sum of Actual for category in active month
    sh.getRange('C' + r).setFormula(
      '=SUMIFS(\'Expense Tracker\'!$F:$F,\'Expense Tracker\'!$B:$B,A' + r + ',' +
      '\'Expense Tracker\'!$A:$A,">="&DATE(Settings!$B$3,Settings!$B$4,1),' +
      '\'Expense Tracker\'!$A:$A,"<"&EDATE(DATE(Settings!$B$3,Settings!$B$4,1),1))'
    );
    sh.getRange('D' + r).setFormula('=B' + r + '-C' + r);
    sh.getRange('E' + r).setFormula('=IF(B' + r + '=0,"—",C' + r + '/B' + r + ')');
    sh.getRange('F' + r).setFormula(
      '=IF(B' + r + '=0,"—",' +
      'IF(D' + r + '<0,"🔴 Over",' +
      'IF(E' + r + '>=0.9,"🟡 Near Limit","🟢 OK")))'
    );
  }

  var n = EXPENSE_CATEGORIES.length + 1;
  sh.getRange(n + 1, 1, 1, 4).setValues([['TOTAL', '', '', '']]);
  sh.getRange('A' + (n + 1)).setFontWeight('bold');
  sh.getRange('B' + (n + 1)).setFormula('=SUM(B2:B' + n + ')').setFontWeight('bold');
  sh.getRange('C' + (n + 1)).setFormula('=SUM(C2:C' + n + ')').setFontWeight('bold');
  sh.getRange('D' + (n + 1)).setFormula('=SUM(D2:D' + n + ')').setFontWeight('bold');

  sh.getRange('B2:D' + (n + 1)).setNumberFormat(CONFIG.CURRENCY_FORMAT);
  sh.getRange('E2:E' + n).setNumberFormat('0%');

  sh.setColumnWidth(1, 140);
  sh.setColumnWidths(2, 5, 120);
}

/* ===================== INSTALLMENT ===================== */

function buildInstallmentSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.INSTALLMENT);
  var oldData = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues() : [];

  sh.clear();
  var headers = ['Item / Month Label', 'Amount', 'Due Date', 'Paid', 'Remaining', 'Status'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sh, headers.length, COLORS.INSTALLMENT);

  if (oldData.length && oldData[0][0] !== '') {
    for (var r = 0; r < oldData.length; r++) {
      sh.getRange(r + 2, 1, 1, 4).setValues([[oldData[r][0], oldData[r][1], oldData[r][2], oldData[r][3]]]);
    }
  }

  for (var i = 2; i <= 101; i++) {
    sh.getRange('E' + i).setFormula('=IF(B' + i + '="","",B' + i + '-D' + i + ')');
    sh.getRange('F' + i).setFormula(
      '=IF(B' + i + '="","",' +
      'IF(D' + i + '>=B' + i + ',"🟢 Paid",' +
      'IF(AND(D' + i + '>0,D' + i + '<B' + i + '),"🟠 Partial",' +
      'IF(C' + i + '<TODAY(),"🔴 Overdue","🔵 Pending"))))'
    );
  }

  sh.getRange('B2:E100').setNumberFormat(CONFIG.CURRENCY_FORMAT);
  sh.getRange('C2:C100').setNumberFormat(CONFIG.DATE_FORMAT);

  sh.getRange('H1').setValue('Total Remaining').setFontWeight('bold');
  sh.getRange('H2').setFormula('=SUMIF(E2:E100,">0")').setNumberFormat(CONFIG.CURRENCY_FORMAT);
  sh.getRange('H3').setValue('Paid Count');
  sh.getRange('I3').setFormula('=COUNTIF(F2:F100,"🟢 Paid")');
  sh.getRange('H4').setValue('Total Items');
  sh.getRange('I4').setFormula('=COUNTA(A2:A100)');

  sh.setColumnWidth(1, 180);
  sh.setColumnWidths(2, 5, 110);
}

/* ===================== DEBT ===================== */

function buildDebtSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.DEBT);
  var oldData = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues() : [];

  sh.clear();
  var headers = ['Person', 'Amount', 'Paid', 'Remaining', 'Date', 'Status'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sh, headers.length, COLORS.DEBT);

  if (oldData.length && oldData[0][0] !== '') {
    for (var r = 0; r < oldData.length; r++) {
      sh.getRange(r + 2, 1, 1, 3).setValues([[oldData[r][0], oldData[r][1], oldData[r][2]]]);
      sh.getRange(r + 2, 5).setValue(oldData[r][4]);
    }
  }

  for (var i = 2; i <= 101; i++) {
    sh.getRange('D' + i).setFormula('=IF(B' + i + '="","",B' + i + '-C' + i + ')');
    sh.getRange('F' + i).setFormula(
      '=IF(B' + i + '="","",' +
      'IF(C' + i + '>=B' + i + ',"Completed",' +
      'IF(C' + i + '>0,"Partial","Pending")))'
    );
  }

  sh.getRange('B2:D100').setNumberFormat(CONFIG.CURRENCY_FORMAT);
  sh.getRange('E2:E100').setNumberFormat(CONFIG.DATE_FORMAT);

  sh.getRange('H1').setValue('Debt Remaining').setFontWeight('bold');
  sh.getRange('H2').setFormula('=SUMIF(D2:D100,">0")').setNumberFormat(CONFIG.CURRENCY_FORMAT);

  sh.setColumnWidth(1, 160);
  sh.setColumnWidths(2, 5, 110);
}

/* ===================== SAVINGS ===================== */

function buildSavingsSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.SAVINGS);
  var oldData = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues() : [];

  sh.clear();
  var headers = ['Date', 'Goal', 'Deposit', 'Withdraw', 'Balance'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sh, headers.length, COLORS.SAVINGS);

  if (oldData.length && oldData[0][0] !== '') {
    for (var r = 0; r < oldData.length; r++) {
      sh.getRange(r + 2, 1, 1, 4).setValues([[oldData[r][0], oldData[r][1], oldData[r][2], oldData[r][3]]]);
    }
  }

  // Running balance
  sh.getRange('E2').setFormula('=IF(A2="","",IFERROR(N(C2),0)-IFERROR(N(D2),0))');
  for (var i = 3; i <= 501; i++) {
    sh.getRange('E' + i).setFormula(
      '=IF(A' + i + '="","",E' + (i - 1) + '+IFERROR(N(C' + i + '),0)-IFERROR(N(D' + i + '),0))'
    );
  }

  sh.getRange('A2:A500').setNumberFormat(CONFIG.DATE_FORMAT);
  sh.getRange('C2:E500').setNumberFormat(CONFIG.CURRENCY_FORMAT);

  sh.getRange('G1').setValue('Current Balance').setFontWeight('bold');
  sh.getRange('G2').setFormula('=IFERROR(LOOKUP(2,1/(E2:E500<>""),E2:E500),0)')
    .setNumberFormat(CONFIG.CURRENCY_FORMAT)
    .setBackground(COLORS.GREEN_BG);

  sh.getRange('G3').setValue('Total Deposits');
  sh.getRange('H3').setFormula('=SUM(C2:C500)').setNumberFormat(CONFIG.CURRENCY_FORMAT);

  sh.setColumnWidth(1, 110);
  sh.setColumnWidth(2, 160);
  sh.setColumnWidths(3, 3, 110);
}

/* ===================== REPORTS ===================== */

function buildReportsSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.REPORTS);
  sh.clear();

  sh.getRange('A1').setValue('📊 AUTOMATIC REPORTS').setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setValue('All values pull from active month in Settings. Refresh Dashboard to sync chart data.')
    .setFontColor(COLORS.MUTED);

  sh.getRange('A3').setValue('── Monthly Report ──').setFontWeight('bold').setFontColor(COLORS.ACCENT);
  sh.getRange('A4').setValue('Total Income');
  sh.getRange('B4').setFormula(
    '=SUMIFS(Income!D:D,Income!A:A,">="&DATE(Settings!$B$3,Settings!$B$4,1),Income!A:A,"<"&EDATE(DATE(Settings!$B$3,Settings!$B$4,1),1))'
  );
  sh.getRange('A5').setValue('Total Expense');
  sh.getRange('B5').setFormula(
    '=SUMIFS(\'Expense Tracker\'!F:F,\'Expense Tracker\'!A:A,">="&DATE(Settings!$B$3,Settings!$B$4,1),\'Expense Tracker\'!A:A,"<"&EDATE(DATE(Settings!$B$3,Settings!$B$4,1),1))'
  );
  sh.getRange('A6').setValue('Net Balance (Income − Expense)');
  sh.getRange('B6').setFormula('=B4-B5');
  sh.getRange('A7').setValue('Total Budget');
  sh.getRange('B7').setFormula('=\'Monthly Budget\'!B' + (EXPENSE_CATEGORIES.length + 2));
  sh.getRange('A8').setValue('Budget Used %');
  sh.getRange('B8').setFormula('=IF(B7=0,0,B5/B7)');

  sh.getRange('A10').setValue('── Weekly Report ──').setFontWeight('bold').setFontColor(COLORS.ACCENT);
  sh.getRange('A11').setValue('This Week Expense');
  sh.getRange('B11').setFormula(
    '=SUMIFS(\'Expense Tracker\'!F:F,\'Expense Tracker\'!A:A,">="&TODAY()-WEEKDAY(TODAY(),2)+1,\'Expense Tracker\'!A:A,"<"&TODAY()-WEEKDAY(TODAY(),2)+8)'
  );

  sh.getRange('A13').setValue('── Expense Insights ──').setFontWeight('bold').setFontColor(COLORS.ACCENT);
  sh.getRange('A14').setValue('Highest Single Expense');
  sh.getRange('B14').setFormula(
    '=IFERROR(MAXIFS(\'Expense Tracker\'!F:F,\'Expense Tracker\'!A:A,">="&DATE(Settings!$B$3,Settings!$B$4,1),\'Expense Tracker\'!A:A,"<"&EDATE(DATE(Settings!$B$3,Settings!$B$4,1),1)),0)'
  );
  sh.getRange('A15').setValue('Highest Expense Category');
  sh.getRange('B15').setFormula(
    '=IFERROR(INDEX(\'Expense Tracker\'!B:B,MATCH(B14,\'Expense Tracker\'!F:F,0)),"—")'
  );
  sh.getRange('A16').setValue('Highest Expense Description');
  sh.getRange('B16').setFormula(
    '=IFERROR(INDEX(\'Expense Tracker\'!D:D,MATCH(B14,\'Expense Tracker\'!F:F,0)),"—")'
  );
  sh.getRange('A17').setValue('Top Category (by spend)');
  sh.getRange('B17').setFormula(
    '=IFERROR(INDEX(\'Monthly Budget\'!A2:A22,MATCH(MAX(\'Monthly Budget\'!C2:C22),\'Monthly Budget\'!C2:C22,0)),"—")'
  );
  sh.getRange('A18').setValue('Top Category Amount');
  sh.getRange('B18').setFormula('=MAX(\'Monthly Budget\'!C2:C22)');
  sh.getRange('A19').setValue('Most Expensive Day (date of highest)');
  sh.getRange('B19').setFormula(
    '=IFERROR(INDEX(\'Expense Tracker\'!A:A,MATCH(B14,\'Expense Tracker\'!F:F,0)),"—")'
  );
  sh.getRange('A20').setValue('Average Daily Expense');
  sh.getRange('B20').setFormula('=IFERROR(B5/DAY(EOMONTH(DATE(Settings!$B$3,Settings!$B$4,1),0)),0)');
  sh.getRange('A21').setValue('Average Weekly Expense');
  sh.getRange('B21').setFormula('=B20*7');
  sh.getRange('A22').setValue("Today's Expense");
  sh.getRange('B22').setFormula('=SUMIF(\'Expense Tracker\'!A:A,TODAY(),\'Expense Tracker\'!F:F)');
  sh.getRange('A23').setValue('Expense Trend (vs last month)');
  sh.getRange('B23').setFormula(
    '=IFERROR(B5-SUMIFS(\'Expense Tracker\'!F:F,\'Expense Tracker\'!A:A,">="&EDATE(DATE(Settings!$B$3,Settings!$B$4,1),-1),\'Expense Tracker\'!A:A,"<"&DATE(Settings!$B$3,Settings!$B$4,1)),0)'
  );

  sh.getRange('B4:B7').setNumberFormat(CONFIG.CURRENCY_FORMAT);
  sh.getRange('B8').setNumberFormat('0.0%');
  sh.getRange('B11').setNumberFormat(CONFIG.CURRENCY_FORMAT);
  sh.getRange('B14').setNumberFormat(CONFIG.CURRENCY_FORMAT);
  sh.getRange('B18:B23').setNumberFormat(CONFIG.CURRENCY_FORMAT);
  sh.getRange('B19').setNumberFormat(CONFIG.DATE_FORMAT);

  sh.setColumnWidth(1, 280);
  sh.setColumnWidth(2, 180);
  sh.getRange('A4:B8').setBackground(COLORS.LIGHT);
  sh.getRange('A11:B11').setBackground(COLORS.YELLOW_BG);
  sh.getRange('A14:B23').setBackground(COLORS.BLUE_BG);
}

function installNamedRanges_(ss) {
  // Optional — skip if permissions noisy; formulas use sheet refs already
}


// ========== Dashboard.gs ==========

/**
 * Dashboard — fully formula-driven, no manual editing required
 */

function buildDashboard_(ss) {
  var sh = ss.getSheetByName(SHEETS.DASHBOARD);
  sh.clear();
  sh.clearCharts();
  sh.setHiddenGridlines(true);

  // ===== HEADER =====
  sh.getRange('A1:L1').merge()
    .setValue(CONFIG.OWNER)
    .setFontSize(22)
    .setFontWeight('bold')
    .setFontColor(COLORS.HEADER_FG)
    .setBackground(COLORS.HEADER_BG)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 42);

  sh.getRange('A2:L2').merge()
    .setValue(CONFIG.APP_NAME)
    .setFontSize(12)
    .setFontColor('#94a3b8')
    .setBackground(COLORS.HEADER_BG)
    .setHorizontalAlignment('center');

  sh.getRange('A3:L3').merge()
    .setFormula(
      '="Month : "&TEXT(DATE(Settings!$B$3,Settings!$B$4,1),"MMMM yyyy")&' +
      '"    ·    Today\'s Cash Position    ·    Last Updated: "&TEXT(NOW(),"dd MMM yyyy HH:mm")'
    )
    .setFontSize(11)
    .setFontWeight('bold')
    .setBackground(COLORS.ACCENT)
    .setFontColor(COLORS.WHITE)
    .setHorizontalAlignment('center');
  sh.setRowHeight(3, 28);

  // ===== FINANCIAL HEALTH SCORE =====
  sh.getRange('A5').setValue('Financial Health Score').setFontWeight('bold').setFontSize(12);
  sh.getRange('A6:B7').merge()
    .setValue('') // filled by refreshDashboard
    .setFontSize(28)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(COLORS.LIGHT);
  sh.setRowHeight(6, 24);
  sh.setRowHeight(7, 24);
  sh.getRange('C5').setValue('Score Label');
  sh.getRange('C6:E7').merge()
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sh.getRange('F5').setValue('Score Breakdown');
  sh.getRange('F6').setValue('Budget · Logging · Overspend · Savings · Installments')
    .setFontColor(COLORS.MUTED).setFontSize(9);
  sh.getRange('F7').setValue(''); // detail from script

  // ===== KPI CARDS (labels + formula values) =====
  var kpis = [
    { cell: 'A9', label: '💰 Total Income', formulaCell: 'A10', color: COLORS.GREEN_BG, formula: incomeFormula_() },
    { cell: 'C9', label: '💸 Total Expense', formulaCell: 'C10', color: COLORS.RED_BG, formula: expenseFormula_() },
    { cell: 'E9', label: '🏦 Current Balance', formulaCell: 'E10', color: COLORS.BLUE_BG, formula: '=A10-C10' },
    { cell: 'G9', label: '📉 Budget Remaining', formulaCell: 'G10', color: COLORS.LIGHT, formula: "='Monthly Budget'!D" + (EXPENSE_CATEGORIES.length + 2) },
    { cell: 'I9', label: '🔴 Over Budget', formulaCell: 'I10', color: COLORS.RED_BG, formula: '=MAX(0,C10-\'Monthly Budget\'!B' + (EXPENSE_CATEGORIES.length + 2) + ')' },
    { cell: 'A12', label: '🟢 Savings', formulaCell: 'A13', color: COLORS.GREEN_BG, formula: "=IFERROR(Savings!G2,0)" },
    { cell: 'C12', label: '🟠 Installment Left', formulaCell: 'C13', color: COLORS.ORANGE_BG, formula: "=IFERROR('Installment Tracker'!H2,0)" },
    { cell: 'E12', label: '🔵 Debt Remaining', formulaCell: 'E13', color: COLORS.BLUE_BG, formula: "=IFERROR('Debt Tracker'!H2,0)" }
  ];

  kpis.forEach(function (k) {
    sh.getRange(k.cell).setValue(k.label).setFontWeight('bold').setFontSize(10).setBackground(k.color);
    var fc = sh.getRange(k.formulaCell);
    fc.setFormula(k.formula)
      .setNumberFormat(CONFIG.CURRENCY_FORMAT)
      .setFontSize(16)
      .setFontWeight('bold')
      .setBackground(k.color);
  });

  // Merge label+value visual blocks
  [['A9', 'B9'], ['C9', 'D9'], ['E9', 'F9'], ['G9', 'H9'], ['I9', 'J9'],
   ['A12', 'B12'], ['C12', 'D12'], ['E12', 'F12'],
   ['A10', 'B10'], ['C10', 'D10'], ['E10', 'F10'], ['G10', 'H10'], ['I10', 'J10'],
   ['A13', 'B13'], ['C13', 'D13'], ['E13', 'F13']].forEach(function (pair) {
    try { sh.getRange(pair[0] + ':' + pair[1]).merge(); } catch (e) {}
  });

  // ===== ALERTS =====
  sh.getRange('A15').setValue('⚠ Alerts').setFontWeight('bold').setFontSize(12).setFontColor(COLORS.OVER);
  sh.getRange('A16:L16').merge().setFormula(
    '=IF(E10<Settings!$B$8,"⚠ Low Cash — Balance below emergency threshold.","")'
  ).setFontColor(COLORS.OVER).setFontWeight('bold').setFontSize(12);

  sh.getRange('A17:L17').merge().setFormula(
    '=IF(I10>0,"🚨 You spent "&TEXT(I10,"#,##0")&" BDT extra this month.","")'
  ).setFontColor(COLORS.OVER).setFontWeight('bold');

  sh.getRange('A18:L18').merge().setFormula(
    '=IF(AND(G10>0,G10/\'Monthly Budget\'!B' + (EXPENSE_CATEGORIES.length + 2) + '<Settings!$B$9/100),' +
    '"🟠 Budget Alert — less than "&Settings!$B$9&"% of budget remaining.","")'
  ).setFontColor(COLORS.WARN).setFontWeight('bold');

  // ===== SMART INSIGHTS =====
  sh.getRange('A20').setValue('💡 Smart Insights').setFontWeight('bold').setFontSize(12);

  sh.getRange('A21').setValue('Largest Expense');
  sh.getRange('B21').setFormula('=Reports!B15');
  sh.getRange('C21').setFormula('=Reports!B14').setNumberFormat(CONFIG.CURRENCY_FORMAT);

  sh.getRange('A22').setValue("Today's Spending");
  sh.getRange('B22').setFormula('=Reports!B22').setNumberFormat(CONFIG.CURRENCY_FORMAT);

  sh.getRange('A23').setValue('This Week');
  sh.getRange('B23').setFormula('=Reports!B11').setNumberFormat(CONFIG.CURRENCY_FORMAT);

  sh.getRange('A24').setValue('Safe Daily Budget');
  // (Monthly Income - Fixed Costs) / days in month
  sh.getRange('B24').setFormula(
    '=MAX(0,(Settings!$B$6-Settings!$B$7)/DAY(EOMONTH(DATE(Settings!$B$3,Settings!$B$4,1),0)))'
  ).setNumberFormat('"৳"#,##0"/day"').setFontWeight('bold').setBackground(COLORS.GREEN_BG);

  sh.getRange('A25:L25').merge().setFormula(
    '=IF(Reports!B22>B24,' +
    '"🔴 Today\'s spending exceeded your limit by "&TEXT(Reports!B22-B24,"#,##0")&" BDT.",' +
    'IF(Reports!B22>0,"🟢 Today\'s spending is within your safe daily limit.","ℹ️ No expenses logged today."))'
  ).setFontWeight('bold');

  // ===== PROGRESS BARS (text-based + %) =====
  sh.getRange('A27').setValue('📈 Progress').setFontWeight('bold').setFontSize(12);

  sh.getRange('A28').setValue('Monthly Budget');
  sh.getRange('B28').setFormula(
    '=IFERROR(Reports!B5/Reports!B7,0)'
  ).setNumberFormat('0%');
  sh.getRange('C28:G28').merge().setFormula(
    '=REPT("█",ROUND(MIN(1,MAX(0,B28))*10))&REPT("░",10-ROUND(MIN(1,MAX(0,B28))*10))&" "&TEXT(B28,"0%")'
  ).setFontFamily('Courier New');

  sh.getRange('A29').setValue('Installment');
  sh.getRange('B29').setFormula(
    '=IFERROR(\'Installment Tracker\'!I3/MAX(1,\'Installment Tracker\'!I4),0)'
  ).setNumberFormat('0%');
  sh.getRange('C29:G29').merge().setFormula(
    '=IFERROR(\'Installment Tracker\'!I3&"/"&\'Installment Tracker\'!I4&" Paid  "&' +
    'REPT("█",ROUND(MIN(1,MAX(0,B29))*10))&REPT("░",10-ROUND(MIN(1,MAX(0,B29))*10)),"—")'
  ).setFontFamily('Courier New');

  sh.getRange('A30').setValue('Savings Goal');
  // Target = 20% of monthly income
  sh.getRange('B30').setFormula(
    '=IFERROR(MIN(1,Savings!G2/MAX(1,Settings!$B$6*0.2)),0)'
  ).setNumberFormat('0%');
  sh.getRange('C30:G30').merge().setFormula(
    '=REPT("█",ROUND(MIN(1,MAX(0,B30))*10))&REPT("░",10-ROUND(MIN(1,MAX(0,B30))*10))&" "&TEXT(B30,"0%")'
  ).setFontFamily('Courier New');
  // ===== CHART ANCHORS =====
  sh.getRange('A32').setValue('📉 Charts').setFontWeight('bold').setFontSize(12);
  sh.getRange('A33').setValue('Expense By Category (Pie) · Daily Expense (Line) · Income vs Expense (Column)')
    .setFontColor(COLORS.MUTED).setFontSize(10);

  // Chart data lives on _ChartData; charts placed below
  // Space for charts rows 34–52

  // ===== FOOTER =====
  sh.getRange('A54:L54').merge()
    .setValue(CONFIG.FOOTER_LINE_1 + '  ' + CONFIG.FOOTER_LINE_2)
    .setHorizontalAlignment('center')
    .setFontStyle('italic')
    .setFontColor(COLORS.MUTED);
  sh.getRange('A55:L55').merge()
    .setValue(CONFIG.FOOTER_LINE_3 + '  ·  ' + CONFIG.BUILT_FOR + '  ·  Version ' + CONFIG.VERSION)
    .setHorizontalAlignment('center')
    .setFontColor(COLORS.MUTED)
    .setFontSize(10);

  // Protect note
  sh.getRange('A57:L57').merge()
    .setValue('ℹ️ This sheet is auto-calculated. Enter data on Income, Expense Tracker, Budget, Installment, Debt & Savings.')
    .setFontColor(COLORS.ACCENT)
    .setFontSize(9);

  // Column widths
  for (var c = 1; c <= 12; c++) {
    sh.setColumnWidth(c, 90);
  }
  sh.setColumnWidth(1, 160);
}

function incomeFormula_() {
  return '=SUMIFS(Income!D:D,Income!A:A,">="&DATE(Settings!$B$3,Settings!$B$4,1),Income!A:A,"<"&EDATE(DATE(Settings!$B$3,Settings!$B$4,1),1))';
}

function expenseFormula_() {
  return "=SUMIFS('Expense Tracker'!F:F,'Expense Tracker'!A:A,\">=\"&DATE(Settings!$B$3,Settings!$B$4,1),'Expense Tracker'!A:A,\"<\"&EDATE(DATE(Settings!$B$3,Settings!$B$4,1),1))";
}

/**
 * Custom function for progress bar — works as sheet formula =progressBar_(0.68)
 */
function progressBar_(ratio) {
  var r = Number(ratio) || 0;
  if (r < 0) r = 0;
  if (r > 1) r = 1;
  var filled = Math.round(r * 10);
  var empty = 10 - filled;
  var bar = '';
  for (var i = 0; i < filled; i++) bar += '█';
  for (var j = 0; j < empty; j++) bar += '░';
  return bar + '  ' + Math.round(r * 100) + '%';
}

function refreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  buildChartDataSheet_(ss);
  updateHealthScoreOnDashboard_(ss);
  applyBudgetAlertFormatting_(ss);
  ss.toast('Dashboard refreshed', '✓', 3);
}

function refreshDashboardLight_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  buildChartDataSheet_(ss);
  updateHealthScoreOnDashboard_(ss);
}

function applyBudgetAlertFormatting_(ss) {
  var dash = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!dash) return;
  var remaining = dash.getRange('G10').getValue();
  var totalBudget = ss.getSheetByName(SHEETS.BUDGET)
    .getRange(EXPENSE_CATEGORIES.length + 2, 2).getValue();
  var threshold = ss.getSheetByName(SHEETS.SETTINGS).getRange('B9').getValue() / 100;
  var cell = dash.getRange('G10');
  if (totalBudget > 0 && remaining / totalBudget < threshold && remaining >= 0) {
    cell.setBackground(COLORS.ORANGE_BG);
  } else if (remaining < 0) {
    cell.setBackground(COLORS.RED_BG);
  } else {
    cell.setBackground(COLORS.LIGHT);
  }

  var balance = dash.getRange('E10').getValue();
  var emergency = ss.getSheetByName(SHEETS.SETTINGS).getRange('B8').getValue();
  if (balance < emergency) {
    dash.getRange('E10').setBackground(COLORS.RED_BG);
  } else {
    dash.getRange('E10').setBackground(COLORS.BLUE_BG);
  }
}


// ========== HealthScore.gs ==========

/**
 * Financial Health Score (0–100)
 *
 * Weights (Settings A17:B21):
 *  Within Budget 30 | Daily Logging 20 | No Overspend 20 | Savings 15 | Installments 15
 */

function updateHealthScoreOnDashboard_(ss) {
  var score = calculateHealthScore_(ss);
  var dash = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!dash) return;

  var band = healthBand_(score.total);
  dash.getRange('A6').setValue(score.total)
    .setFontColor(band.color)
    .setBackground(COLORS.LIGHT);
  dash.getRange('C6')
    .setValue(band.emoji + ' ' + score.total + ' — ' + band.label)
    .setFontColor(band.color)
    .setBackground(COLORS.LIGHT);
  dash.getRange('F7').setValue(
    'Budget ' + score.parts.budget + '/30 · Log ' + score.parts.logging +
    '/20 · Overspend ' + score.parts.overspend + '/20 · Save ' +
    score.parts.savings + '/15 · Install ' + score.parts.installments + '/15'
  ).setFontSize(9).setFontColor(COLORS.MUTED);
}

function healthBand_(score) {
  if (score >= HEALTH.EXCELLENT.min) return HEALTH.EXCELLENT;
  if (score >= HEALTH.GOOD.min) return HEALTH.GOOD;
  if (score >= HEALTH.ATTENTION.min) return HEALTH.ATTENTION;
  return HEALTH.CRITICAL;
}

function calculateHealthScore_(ss) {
  var settings = ss.getSheetByName(SHEETS.SETTINGS);
  var year = settings.getRange('B3').getValue();
  var month = settings.getRange('B4').getValue();
  var start = new Date(year, month - 1, 1);
  var end = new Date(year, month, 1);
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var daysInMonth = new Date(year, month, 0).getDate();
  var dayOfMonth = (today.getFullYear() === year && today.getMonth() + 1 === month)
    ? today.getDate()
    : daysInMonth;

  // --- Budget adherence (30) ---
  var budgetSh = ss.getSheetByName(SHEETS.BUDGET);
  var totalBudget = Number(budgetSh.getRange(EXPENSE_CATEGORIES.length + 2, 2).getValue()) || 0;
  var used = Number(budgetSh.getRange(EXPENSE_CATEGORIES.length + 2, 3).getValue()) || 0;
  var budgetPts = 30;
  if (totalBudget > 0) {
    var ratio = used / totalBudget;
    if (ratio <= 1) {
      budgetPts = Math.round(30 * (1 - Math.max(0, ratio - 0.7) / 0.3 * 0.3));
      // Full 30 if under 70% used; taper to ~21 at 100%; if under budget keep high
      if (ratio <= 0.85) budgetPts = 30;
      else if (ratio <= 1) budgetPts = Math.round(30 - (ratio - 0.85) / 0.15 * 8);
      else budgetPts = Math.max(0, Math.round(15 - (ratio - 1) * 30));
    } else {
      budgetPts = Math.max(0, Math.round(12 - (ratio - 1) * 40));
    }
  }

  // --- Daily logging (20) — unique days with expenses / days elapsed ---
  var exp = ss.getSheetByName(SHEETS.EXPENSE);
  var last = exp.getLastRow();
  var loggedDays = {};
  if (last >= 2) {
    var dates = exp.getRange(2, 1, last - 1, 1).getValues();
    var actuals = exp.getRange(2, 6, last - 1, 1).getValues();
    for (var i = 0; i < dates.length; i++) {
      var d = dates[i][0];
      if (!(d instanceof Date)) continue;
      if (d >= start && d < end && actuals[i][0] !== '' && actuals[i][0] != null) {
        loggedDays[Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd')] = true;
      }
    }
  }
  var uniqueDays = Object.keys(loggedDays).length;
  var loggingPts = dayOfMonth > 0
    ? Math.round(20 * Math.min(1, uniqueDays / dayOfMonth))
    : 0;

  // --- Avoid overspending (20) — count OVER rows vs total ---
  var overCount = 0;
  var lineCount = 0;
  if (last >= 2) {
    var statuses = exp.getRange(2, 8, last - 1, 1).getValues();
    var cats = exp.getRange(2, 1, last - 1, 1).getValues();
    for (var j = 0; j < statuses.length; j++) {
      var dd = cats[j][0];
      if (!(dd instanceof Date)) continue;
      if (dd >= start && dd < end && statuses[j][0]) {
        lineCount++;
        if (String(statuses[j][0]).indexOf('OVER') !== -1) overCount++;
      }
    }
  }
  var overspendPts = 20;
  if (lineCount > 0) {
    overspendPts = Math.round(20 * (1 - overCount / lineCount));
  }

  // --- Savings deposits this month (15) ---
  var sav = ss.getSheetByName(SHEETS.SAVINGS);
  var savLast = sav.getLastRow();
  var depositSum = 0;
  if (savLast >= 2) {
    var sDates = sav.getRange(2, 1, savLast - 1, 1).getValues();
    var deps = sav.getRange(2, 3, savLast - 1, 1).getValues();
    for (var k = 0; k < sDates.length; k++) {
      var sd = sDates[k][0];
      if (sd instanceof Date && sd >= start && sd < end) {
        depositSum += Number(deps[k][0]) || 0;
      }
    }
  }
  var incomeTarget = Number(settings.getRange('B6').getValue()) || CONFIG.DEFAULT_MONTHLY_INCOME;
  var saveGoal = incomeTarget * 0.1; // 10% of income ideal
  var savingsPts = saveGoal > 0
    ? Math.round(15 * Math.min(1, depositSum / saveGoal))
    : (depositSum > 0 ? 15 : 0);

  // --- Installments on time (15) ---
  var inst = ss.getSheetByName(SHEETS.INSTALLMENT);
  var iLast = inst.getLastRow();
  var instPts = 15;
  if (iLast >= 2) {
    var statusesI = inst.getRange(2, 6, iLast - 1, 1).getValues();
    var amounts = inst.getRange(2, 2, iLast - 1, 1).getValues();
    var totalI = 0;
    var good = 0;
    for (var m = 0; m < amounts.length; m++) {
      if (!amounts[m][0]) continue;
      totalI++;
      var st = String(statusesI[m][0] || '');
      if (st.indexOf('Paid') !== -1 || st.indexOf('Partial') !== -1 || st.indexOf('Pending') !== -1) {
        if (st.indexOf('Overdue') === -1) good++;
      }
    }
    instPts = totalI > 0 ? Math.round(15 * (good / totalI)) : 15;
  }

  var total = budgetPts + loggingPts + overspendPts + savingsPts + instPts;
  total = Math.max(0, Math.min(100, total));

  return {
    total: total,
    parts: {
      budget: budgetPts,
      logging: loggingPts,
      overspend: overspendPts,
      savings: savingsPts,
      installments: instPts
    }
  };
}


// ========== Charts.gs ==========

/**
 * Conditional formatting & chart data / charts
 */

function applyAllConditionalFormatting_(ss) {
  applyExpenseFormatting_(ss);
  applyBudgetFormatting_(ss);
  applyDebtFormatting_(ss);
}

function applyExpenseFormatting_(ss) {
  var sh = ss.getSheetByName(SHEETS.EXPENSE);
  sh.clearConditionalFormatRules();
  var range = sh.getRange('A2:J1000');

  var over = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2>$E2')
    .setBackground(COLORS.RED_BG)
    .setRanges([range])
    .build();

  var near = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($E2<>"",$F2<>"",$F2<=$E2,$F2>=$E2*0.9)')
    .setBackground(COLORS.YELLOW_BG)
    .setRanges([range])
    .build();

  var under = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($E2<>"",$F2<>"",$F2<$E2*0.9)')
    .setBackground(COLORS.GREEN_BG)
    .setRanges([range])
    .build();

  sh.setConditionalFormatRules([over, near, under]);
}

function applyBudgetFormatting_(ss) {
  var sh = ss.getSheetByName(SHEETS.BUDGET);
  sh.clearConditionalFormatRules();
  var n = EXPENSE_CATEGORIES.length;
  var range = sh.getRange(2, 1, n, 6);

  var over = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$D2<0')
    .setBackground(COLORS.RED_BG)
    .setRanges([range])
    .build();

  var warn = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($B2>0,$E2>=0.9,$D2>=0)')
    .setBackground(COLORS.YELLOW_BG)
    .setRanges([range])
    .build();

  var ok = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($B2>0,$D2>0,$E2<0.9)')
    .setBackground(COLORS.GREEN_BG)
    .setRanges([range])
    .build();

  sh.setConditionalFormatRules([over, warn, ok]);
}

function applyDebtFormatting_(ss) {
  var sh = ss.getSheetByName(SHEETS.DEBT);
  sh.clearConditionalFormatRules();
  var range = sh.getRange('A2:F100');

  var done = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="Completed"')
    .setBackground(COLORS.GREEN_BG)
    .setRanges([range])
    .build();

  var partial = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="Partial"')
    .setBackground(COLORS.YELLOW_BG)
    .setRanges([range])
    .build();

  var pending = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="Pending"')
    .setBackground(COLORS.ORANGE_BG)
    .setRanges([range])
    .build();

  sh.setConditionalFormatRules([done, partial, pending]);
}

/* ===================== CHART DATA ===================== */

function buildChartDataSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.CHART_DATA) || ss.insertSheet(SHEETS.CHART_DATA);
  sh.clear();
  sh.hideSheet();

  var settings = ss.getSheetByName(SHEETS.SETTINGS);
  var year = settings.getRange('B3').getValue();
  var month = settings.getRange('B4').getValue();
  var start = new Date(year, month - 1, 1);
  var end = new Date(year, month, 1);
  var daysInMonth = new Date(year, month, 0).getDate();

  // Category totals for pie
  sh.getRange('A1:B1').setValues([['Category', 'Amount']]);
  var budget = ss.getSheetByName(SHEETS.BUDGET);
  var catData = budget.getRange(2, 1, EXPENSE_CATEGORIES.length, 3).getValues();
  var pieRows = [];
  catData.forEach(function (r) {
    if (Number(r[2]) > 0) pieRows.push([r[0], Number(r[2])]);
  });
  if (!pieRows.length) pieRows.push(['No data', 0]);
  sh.getRange(2, 1, pieRows.length, 2).setValues(pieRows);

  // Daily expense for line chart
  sh.getRange('D1:E1').setValues([['Day', 'Expense']]);
  var daily = [];
  for (var d = 1; d <= daysInMonth; d++) daily.push([d, 0]);

  var exp = ss.getSheetByName(SHEETS.EXPENSE);
  var last = exp.getLastRow();
  if (last >= 2) {
    var rows = exp.getRange(2, 1, last - 1, 6).getValues();
    rows.forEach(function (r) {
      var dt = r[0];
      var amt = Number(r[5]) || 0;
      if (dt instanceof Date && dt >= start && dt < end && amt) {
        daily[dt.getDate() - 1][1] += amt;
      }
    });
  }
  sh.getRange(2, 4, daily.length, 2).setValues(daily);

  // Income vs Expense for column
  sh.getRange('G1:I1').setValues([['Metric', 'Income', 'Expense']]);
  var income = 0;
  var expense = 0;
  var inc = ss.getSheetByName(SHEETS.INCOME);
  var iLast = inc.getLastRow();
  if (iLast >= 2) {
    var iRows = inc.getRange(2, 1, iLast - 1, 4).getValues();
    iRows.forEach(function (r) {
      if (r[0] instanceof Date && r[0] >= start && r[0] < end) income += Number(r[3]) || 0;
    });
  }
  daily.forEach(function (x) { expense += x[1]; });
  sh.getRange('G2:I2').setValues([['This Month', income, expense]]);
}

function rebuildCharts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  buildChartDataSheet_(ss);
  var dash = ss.getSheetByName(SHEETS.DASHBOARD);
  var data = ss.getSheetByName(SHEETS.CHART_DATA);

  // Remove existing charts
  dash.getCharts().forEach(function (c) {
    dash.removeChart(c);
  });

  var pieLast = Math.max(data.getLastRow(), 2);
  // Only category rows (col A)
  var pieEnd = 2;
  while (pieEnd <= pieLast && data.getRange(pieEnd, 1).getValue() !== '') pieEnd++;
  pieEnd = Math.max(pieEnd - 1, 2);

  var pieChart = dash.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(data.getRange(1, 1, pieEnd, 2))
    .setOption('title', 'Expense By Category')
    .setOption('pieHole', 0.35)
    .setPosition(34, 1, 0, 0)
    .build();
  dash.insertChart(pieChart);

  var days = new Date(
    ss.getSheetByName(SHEETS.SETTINGS).getRange('B3').getValue(),
    ss.getSheetByName(SHEETS.SETTINGS).getRange('B4').getValue(),
    0
  ).getDate();

  var lineChart = dash.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(data.getRange(1, 4, days + 1, 2))
    .setOption('title', 'Daily Expense')
    .setOption('curveType', 'function')
    .setOption('legend', { position: 'none' })
    .setPosition(34, 5, 0, 0)
    .build();
  dash.insertChart(lineChart);

  var colChart = dash.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(data.getRange(1, 7, 2, 3))
    .setOption('title', 'Income vs Expense')
    .setPosition(34, 9, 0, 0)
    .build();
  dash.insertChart(colChart);

  ss.toast('Charts rebuilt', '📊', 3);
}


// ========== SampleData.gs ==========

/**
 * Optional sample data for July 2026 demo
 */

function addSampleData() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert(
    'Add Sample Data?',
    'This adds demo rows for Income, Expenses, Installments, Debt & Savings for the active month. Existing rows are kept.',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = ss.getSheetByName(SHEETS.SETTINGS);
  var year = settings.getRange('B3').getValue();
  var month = settings.getRange('B4').getValue();

  // Income
  var income = ss.getSheetByName(SHEETS.INCOME);
  var iRow = Math.max(income.getLastRow() + 1, 2);
  income.getRange(iRow, 1, 2, 6).setValues([
    [new Date(year, month - 1, 1), 'Salary', 'Monthly salary', 30000, 'Bank', 'Primary'],
    [new Date(year, month - 1, 10), 'Freelance', 'Side project', 5000, 'bKash', '']
  ]);

  // Expenses
  var exp = ss.getSheetByName(SHEETS.EXPENSE);
  var eRow = Math.max(exp.getLastRow() + 1, 2);
  var samples = [
    [new Date(year, month - 1, 2), 'Market', 'Grocery', 'Rice & oil', 500, 650, , , 'Cash', ''],
    [new Date(year, month - 1, 3), 'Transport', 'CNG', 'Office commute', 200, 180, , , 'Cash', ''],
    [new Date(year, month - 1, 5), 'Restaurant', 'Lunch', 'Office lunch', 300, 280, , , 'bKash', ''],
    [new Date(year, month - 1, 8), 'Electricity', 'Home Bill', 'DESCO bill', 1500, 1450, , , 'Bank', ''],
    [new Date(year, month - 1, 12), 'Internet', 'Home', 'Monthly wifi', 1000, 1000, , , 'bKash', ''],
    [new Date(year, month - 1, 15), 'Market', 'Grocery', 'Weekly market', 2000, 1850, , , 'Cash', ''],
    [new Date(year, month - 1, 18), 'Medical', 'Medicine', 'Pharmacy', 500, 420, , , 'Cash', ''],
    [new Date(year, month - 1, 20), 'Entertainment', 'Subscription', 'Streaming', 500, 550, , , 'Card', ''],
    [new Date(year, month - 1, 22), 'Kids', 'School', 'Tuition', 1500, 1500, , , 'Cash', '']
  ];
  // Write A-F and I-J; formulas already in G-H
  samples.forEach(function (s, idx) {
    var r = eRow + idx;
    exp.getRange(r, 1, 1, 6).setValues([[s[0], s[1], s[2], s[3], s[4], s[5]]]);
    exp.getRange(r, 9, 1, 2).setValues([[s[8], s[9]]]);
  });

  // Installments
  var inst = ss.getSheetByName(SHEETS.INSTALLMENT);
  var nRow = Math.max(inst.getLastRow() + 1, 2);
  inst.getRange(nRow, 1, 3, 4).setValues([
    ['Phone EMI — Month 6/9', 3500, new Date(year, month - 1, 5), 3500],
    ['Phone EMI — Month 7/9', 3500, new Date(year, month - 1, 5), 0],
    ['Laptop EMI — Month 2/6', 5000, new Date(year, month - 1, 12), 2500]
  ]);

  // Debt
  var debt = ss.getSheetByName(SHEETS.DEBT);
  var dRow = Math.max(debt.getLastRow() + 1, 2);
  debt.getRange(dRow, 1, 2, 3).setValues([
    ['Rahim', 5000, 2000],
    ['Office Fund', 2000, 2000]
  ]);
  debt.getRange(dRow, 5).setValue(new Date(year, month - 1, 1));
  debt.getRange(dRow + 1, 5).setValue(new Date(year, month - 1, 8));

  // Savings
  var sav = ss.getSheetByName(SHEETS.SAVINGS);
  var sRow = Math.max(sav.getLastRow() + 1, 2);
  sav.getRange(sRow, 1, 2, 4).setValues([
    [new Date(year, month - 1, 1), 'Emergency Fund', 3000, 0],
    [new Date(year, month - 1, 15), 'Emergency Fund', 2000, 0]
  ]);

  rebuildCharts();
  refreshDashboard();
  ui.alert('Sample data added. Open Dashboard to see KPIs, alerts & charts.');
}

