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
