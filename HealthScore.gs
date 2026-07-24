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
