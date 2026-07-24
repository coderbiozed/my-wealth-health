/**
 * Web App API — sync Netlify frontend ↔ Google Sheets
 *
 * Deploy: Deploy → New deployment → Web app
 *   Execute as: Me
 *   Who has access: Anyone
 * Copy the Web App URL into the Netlify app Settings → Sheet API URL
 */

var SYNC_TOKEN_CELL = 'B12'; // Settings!B12 optional token
var API_URL_HINT_CELL = 'B13';

function doGet(e) {
  try {
    var params = (e && e.parameter) || {};
    if (!checkToken_(params.token)) {
      return jsonOut_({ ok: false, error: 'Unauthorized' });
    }
    var action = params.action || 'load';
    if (action === 'ping') {
      return jsonOut_({ ok: true, app: 'My Wealth Health', version: CONFIG.VERSION });
    }
    return jsonOut_(exportState_());
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    if (!checkToken_(body.token)) {
      return jsonOut_({ ok: false, error: 'Unauthorized' });
    }
    var action = body.action || 'save';
    if (action === 'load') return jsonOut_(exportState_());
    if (action === 'save') {
      if (!body.data) return jsonOut_({ ok: false, error: 'Missing data' });
      importState_(body.data);
      return jsonOut_({ ok: true, savedAt: new Date().toISOString() });
    }
    if (action === 'ping') {
      return jsonOut_({ ok: true, app: 'My Wealth Health', version: CONFIG.VERSION });
    }
    return jsonOut_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function checkToken_(token) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = ss.getSheetByName(SHEETS.SETTINGS);
  if (!settings) return true;
  ensureSyncSettingsRows_(settings);
  var expected = String(settings.getRange(SYNC_TOKEN_CELL).getValue() || '').trim();
  if (!expected) return true; // no token configured = open
  return String(token || '').trim() === expected;
}

function ensureSyncSettingsRows_(settings) {
  if (!settings.getRange('A12').getValue()) {
    settings.getRange('A12').setValue('Sync Token (optional)').setFontWeight('bold').setBackground(COLORS.LIGHT);
    settings.getRange(SYNC_TOKEN_CELL).setBackground(COLORS.BLUE_BG);
    settings.getRange('C12').setValue('If set, Netlify app must send the same token');
  }
  if (!settings.getRange('A13').getValue()) {
    settings.getRange('A13').setValue('Web App URL (paste in Netlify Settings)').setFontWeight('bold').setBackground(COLORS.LIGHT);
    settings.getRange(API_URL_HINT_CELL).setBackground(COLORS.YELLOW_BG);
    settings.getRange('C13').setValue('Deploy → Web app → copy URL → paste in website Settings');
  }
}

/** Export sheet data as frontend state JSON */
function exportState_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAllSheets_(ss);
  ensureSyncSettingsRows_(ss.getSheetByName(SHEETS.SETTINGS));

  var settingsSh = ss.getSheetByName(SHEETS.SETTINGS);
  var settings = {
    year: Number(settingsSh.getRange('B3').getValue()) || new Date().getFullYear(),
    month: Number(settingsSh.getRange('B4').getValue()) || (new Date().getMonth() + 1),
    owner: String(settingsSh.getRange('B5').getValue() || CONFIG.OWNER),
    monthlyIncome: Number(settingsSh.getRange('B6').getValue()) || 30000,
    fixedCosts: Number(settingsSh.getRange('B7').getValue()) || 0,
    emergencyBalance: Number(settingsSh.getRange('B8').getValue()) || 1000,
    budgetAlertPct: Number(settingsSh.getRange('B9').getValue()) || 20
  };

  var budgets = {};
  var bud = ss.getSheetByName(SHEETS.BUDGET);
  var bLast = bud.getLastRow();
  if (bLast >= 2) {
    var bRows = bud.getRange(2, 1, bLast - 1, 2).getValues();
    bRows.forEach(function (r) {
      if (r[0]) budgets[String(r[0])] = Number(r[1]) || 0;
    });
  }
  EXPENSE_CATEGORIES.forEach(function (c) {
    if (budgets[c] == null) budgets[c] = DEFAULT_BUDGETS[c] || 0;
  });

  var income = readIncome_(ss);
  var expenses = readExpenses_(ss);
  var installments = readInstallments_(ss);
  var debts = readDebts_(ss);
  var savings = readSavings_(ss);

  return {
    ok: true,
    data: {
      settings: settings,
      budgets: budgets,
      income: income,
      expenses: expenses,
      installments: installments,
      debts: debts,
      savings: savings
    },
    loadedAt: new Date().toISOString()
  };
}

function importState_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureWorkbookForApi_(ss);

  var settingsSh = ss.getSheetByName(SHEETS.SETTINGS);
  ensureSyncSettingsRows_(settingsSh);
  var s = data.settings || {};
  if (s.year != null) settingsSh.getRange('B3').setValue(Number(s.year));
  if (s.month != null) settingsSh.getRange('B4').setValue(Number(s.month));
  if (s.owner != null) settingsSh.getRange('B5').setValue(s.owner);
  if (s.monthlyIncome != null) settingsSh.getRange('B6').setValue(Number(s.monthlyIncome));
  if (s.fixedCosts != null) settingsSh.getRange('B7').setValue(Number(s.fixedCosts));
  if (s.emergencyBalance != null) settingsSh.getRange('B8').setValue(Number(s.emergencyBalance));
  if (s.budgetAlertPct != null) settingsSh.getRange('B9').setValue(Number(s.budgetAlertPct));

  writeBudgets_(ss, data.budgets || {});
  writeIncome_(ss, data.income || []);
  writeExpenses_(ss, data.expenses || []);
  writeInstallments_(ss, data.installments || []);
  writeDebts_(ss, data.debts || []);
  writeSavings_(ss, data.savings || []);

  try {
    rebuildCharts();
    refreshDashboard();
  } catch (err) {
    // charts optional during API write
  }
}

/** Quiet setup for API (no UI alerts) */
function ensureWorkbookForApi_(ss) {
  ensureAllSheets_(ss);
  if (!ss.getSheetByName(SHEETS.SETTINGS) || !ss.getSheetByName(SHEETS.SETTINGS).getRange('A1').getValue()) {
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
  }
  ensureSyncSettingsRows_(ss.getSheetByName(SHEETS.SETTINGS));
}

function toIsoDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(v);
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  return s;
}

function toSheetDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return v;
  var s = String(v).slice(0, 10);
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return v;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function clearDataRows_(sheet, startRow, cols) {
  var last = sheet.getLastRow();
  if (last >= startRow) {
    sheet.getRange(startRow, 1, last - startRow + 1, cols).clearContent();
  }
}

function readIncome_(ss) {
  var sh = ss.getSheetByName(SHEETS.INCOME);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 6).getValues();
  return rows.filter(function (r) { return r[0] || r[3]; }).map(function (r, i) {
    return {
      id: 'inc-' + i + '-' + toIsoDate_(r[0]),
      date: toIsoDate_(r[0]),
      source: r[1] || '',
      description: r[2] || '',
      amount: Number(r[3]) || 0,
      payment: r[4] || '',
      notes: r[5] || ''
    };
  });
}

function writeIncome_(ss, list) {
  var sh = ss.getSheetByName(SHEETS.INCOME);
  clearDataRows_(sh, 2, 6);
  if (!list.length) return;
  var values = list.map(function (r) {
    return [toSheetDate_(r.date), r.source || '', r.description || '', Number(r.amount) || 0, r.payment || '', r.notes || ''];
  });
  sh.getRange(2, 1, values.length, 6).setValues(values);
  sh.getRange(2, 1, values.length, 1).setNumberFormat(CONFIG.DATE_FORMAT);
  sh.getRange(2, 4, values.length, 1).setNumberFormat(CONFIG.CURRENCY_FORMAT);
}

function readExpenses_(ss) {
  var sh = ss.getSheetByName(SHEETS.EXPENSE);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 10).getValues();
  return rows.filter(function (r) { return r[0] || r[5] !== ''; }).map(function (r, i) {
    return {
      id: 'exp-' + i + '-' + toIsoDate_(r[0]),
      date: toIsoDate_(r[0]),
      category: r[1] || '',
      subCategory: r[2] || '',
      description: r[3] || '',
      budget: r[4] === '' || r[4] == null ? '' : Number(r[4]),
      actual: r[5] === '' || r[5] == null ? '' : Number(r[5]),
      payment: r[8] || '',
      notes: r[9] || ''
    };
  });
}

function writeExpenses_(ss, list) {
  var sh = ss.getSheetByName(SHEETS.EXPENSE);
  clearDataRows_(sh, 2, 10);
  if (!list.length) return;
  var n = list.length;
  var left = list.map(function (r) {
    return [
      toSheetDate_(r.date),
      r.category || '',
      r.subCategory || '',
      r.description || '',
      r.budget === '' || r.budget == null ? '' : Number(r.budget),
      r.actual === '' || r.actual == null ? '' : Number(r.actual)
    ];
  });
  var right = list.map(function (r) {
    return [r.payment || '', r.notes || ''];
  });
  sh.getRange(2, 1, n, 6).setValues(left);
  sh.getRange(2, 9, n, 2).setValues(right);

  // Restore Difference + Status formulas
  for (var i = 0; i < n; i++) {
    var row = i + 2;
    sh.getRange('G' + row).setFormula(
      '=IF(OR(E' + row + '="",F' + row + '=""),"",E' + row + '-F' + row + ')'
    );
    sh.getRange('H' + row).setFormula(
      '=IF(OR(E' + row + '="",F' + row + '=""),"",' +
      'IF(F' + row + '>E' + row + ',"🔴 OVER",' +
      'IF(F' + row + '<E' + row + ',"🟢 UNDER","🟡 PERFECT")))'
    );
  }
  sh.getRange(2, 1, n, 1).setNumberFormat(CONFIG.DATE_FORMAT);
  sh.getRange(2, 5, n, 3).setNumberFormat(CONFIG.CURRENCY_FORMAT);
}

function readInstallments_(ss) {
  var sh = ss.getSheetByName(SHEETS.INSTALLMENT);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 4).getValues();
  return rows.filter(function (r) { return r[0] || r[1]; }).map(function (r, i) {
    return {
      id: 'inst-' + i,
      label: r[0] || '',
      amount: Number(r[1]) || 0,
      dueDate: toIsoDate_(r[2]),
      paid: Number(r[3]) || 0
    };
  });
}

function writeInstallments_(ss, list) {
  var sh = ss.getSheetByName(SHEETS.INSTALLMENT);
  clearDataRows_(sh, 2, 6);
  if (!list.length) return;
  var values = list.map(function (r) {
    return [r.label || '', Number(r.amount) || 0, toSheetDate_(r.dueDate), Number(r.paid) || 0];
  });
  sh.getRange(2, 1, values.length, 4).setValues(values);
  for (var i = 0; i < values.length; i++) {
    var row = i + 2;
    sh.getRange('E' + row).setFormula('=IF(B' + row + '="","",B' + row + '-D' + row + ')');
    sh.getRange('F' + row).setFormula(
      '=IF(B' + row + '="","",' +
      'IF(D' + row + '>=B' + row + ',"🟢 Paid",' +
      'IF(AND(D' + row + '>0,D' + row + '<B' + row + '),"🟠 Partial",' +
      'IF(C' + row + '<TODAY(),"🔴 Overdue","🔵 Pending"))))'
    );
  }
  sh.getRange(2, 2, values.length, 3).setNumberFormat(CONFIG.CURRENCY_FORMAT);
  sh.getRange(2, 3, values.length, 1).setNumberFormat(CONFIG.DATE_FORMAT);
}

function readDebts_(ss) {
  var sh = ss.getSheetByName(SHEETS.DEBT);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 5).getValues();
  return rows.filter(function (r) { return r[0] || r[1]; }).map(function (r, i) {
    return {
      id: 'debt-' + i,
      person: r[0] || '',
      amount: Number(r[1]) || 0,
      paid: Number(r[2]) || 0,
      date: toIsoDate_(r[4])
    };
  });
}

function writeDebts_(ss, list) {
  var sh = ss.getSheetByName(SHEETS.DEBT);
  clearDataRows_(sh, 2, 6);
  if (!list.length) return;
  var left = list.map(function (r) {
    return [r.person || '', Number(r.amount) || 0, Number(r.paid) || 0];
  });
  var dates = list.map(function (r) { return [toSheetDate_(r.date)]; });
  sh.getRange(2, 1, left.length, 3).setValues(left);
  sh.getRange(2, 5, dates.length, 1).setValues(dates);
  for (var i = 0; i < list.length; i++) {
    var row = i + 2;
    sh.getRange('D' + row).setFormula('=IF(B' + row + '="","",B' + row + '-C' + row + ')');
    sh.getRange('F' + row).setFormula(
      '=IF(B' + row + '="","",' +
      'IF(C' + row + '>=B' + row + ',"Completed",' +
      'IF(C' + row + '>0,"Partial","Pending")))'
    );
  }
  sh.getRange(2, 2, list.length, 3).setNumberFormat(CONFIG.CURRENCY_FORMAT);
  sh.getRange(2, 5, list.length, 1).setNumberFormat(CONFIG.DATE_FORMAT);
}

function readSavings_(ss) {
  var sh = ss.getSheetByName(SHEETS.SAVINGS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 4).getValues();
  return rows.filter(function (r) { return r[0] || r[2] || r[3]; }).map(function (r, i) {
    return {
      id: 'sav-' + i + '-' + toIsoDate_(r[0]),
      date: toIsoDate_(r[0]),
      goal: r[1] || '',
      deposit: Number(r[2]) || 0,
      withdraw: Number(r[3]) || 0
    };
  });
}

function writeSavings_(ss, list) {
  var sh = ss.getSheetByName(SHEETS.SAVINGS);
  clearDataRows_(sh, 2, 5);
  if (!list.length) return;
  var values = list.map(function (r) {
    return [toSheetDate_(r.date), r.goal || '', Number(r.deposit) || 0, Number(r.withdraw) || 0];
  });
  sh.getRange(2, 1, values.length, 4).setValues(values);
  sh.getRange('E2').setFormula('=IF(A2="","",IFERROR(N(C2),0)-IFERROR(N(D2),0))');
  for (var i = 3; i <= values.length + 1; i++) {
    sh.getRange('E' + i).setFormula(
      '=IF(A' + i + '="","",E' + (i - 1) + '+IFERROR(N(C' + i + '),0)-IFERROR(N(D' + i + '),0))'
    );
  }
  sh.getRange(2, 1, values.length, 1).setNumberFormat(CONFIG.DATE_FORMAT);
  sh.getRange(2, 3, values.length, 3).setNumberFormat(CONFIG.CURRENCY_FORMAT);
}

function writeBudgets_(ss, budgets) {
  var sh = ss.getSheetByName(SHEETS.BUDGET);
  var last = Math.max(sh.getLastRow(), EXPENSE_CATEGORIES.length + 1);
  var rows = sh.getRange(2, 1, EXPENSE_CATEGORIES.length, 1).getValues();
  for (var i = 0; i < rows.length; i++) {
    var cat = rows[i][0];
    if (cat && budgets[cat] != null) {
      sh.getRange(i + 2, 2).setValue(Number(budgets[cat]) || 0);
    }
  }
}
