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
