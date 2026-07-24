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
