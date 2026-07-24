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
