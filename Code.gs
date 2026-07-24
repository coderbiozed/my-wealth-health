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
    .addItem('🌐 Show Sync Setup Help', 'showSyncHelp')
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();
}

function showSyncHelp() {
  SpreadsheetApp.getUi().alert(
    'Netlify ↔ Sheet Sync',
    '1. Deploy → New deployment → Web app\n' +
      '2. Execute as: Me · Who has access: Anyone\n' +
      '3. Copy the /exec URL\n' +
      '4. Open https://my-wealth-health.netlify.app/ → Settings\n' +
      '5. Paste URL (and optional token from Settings!B12)\n' +
      '6. Test connection → Save to Sheet\n\n' +
      'Website edits will write into this workbook.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
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
