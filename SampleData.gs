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
