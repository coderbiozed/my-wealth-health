/**
 * My Wealth Health — Biozed Hossain
 * Local HTML frontend with localStorage persistence
 */

const STORAGE_KEY = 'myWealthHealth_v1';

const INCOME_SOURCES = ['Salary', 'Freelance', 'Bonus', 'Gift', 'Refund', 'Other'];
const PAYMENT_METHODS = ['Cash', 'bKash', 'Nagad', 'Rocket', 'Bank', 'Card', 'Other'];
const EXPENSE_CATEGORIES = [
  'Market', 'Rent', 'Electricity', 'Gas', 'Internet', 'Transport', 'Medical',
  'Education', 'Kids', 'Clothing', 'Restaurant', 'Shopping', 'Entertainment',
  'Gift', 'Installment', 'Loan', 'Emergency', 'Investment', 'Business', 'Office', 'Other'
];
const SUB_CATEGORIES = {
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
const DEFAULT_BUDGETS = {
  Market: 15000, Rent: 8000, Electricity: 1500, Gas: 1200, Internet: 1000,
  Transport: 3000, Medical: 2000, Education: 2000, Kids: 3000, Clothing: 1500,
  Restaurant: 2000, Shopping: 2000, Entertainment: 1500, Gift: 1000,
  Installment: 5000, Loan: 0, Emergency: 1000, Investment: 2000, Business: 0,
  Office: 500, Other: 1000
};

const TITLES = {
  dashboard: ['Dashboard', "Today's cash position — auto calculated"],
  income: ['Income', 'Add and update income sources'],
  expense: ['Expense Tracker', 'Log daily spending with budget vs actual'],
  budget: ['Monthly Budget', 'Set category limits — used amount updates live'],
  installment: ['Installment Tracker', 'Track EMIs and remaining amounts'],
  debt: ['Debt Tracker', 'People you owe / are owed'],
  savings: ['Savings', 'Deposits, withdrawals, running balance'],
  reports: ['Reports', 'Monthly & weekly insights'],
  settings: ['Settings', 'Income, alerts, and preferences']
};

let state = loadState();
let currentView = 'dashboard';
let editCtx = null;
let charts = {};

function uid() {
  return crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(16).slice(2);
}

function defaultState() {
  const now = new Date();
  return {
    settings: {
      year: now.getFullYear(),
      month: now.getMonth() + 1,
      owner: 'BIOZED HOSSAIN',
      monthlyIncome: 30000,
      fixedCosts: 0,
      emergencyBalance: 1000,
      budgetAlertPct: 20
    },
    budgets: { ...DEFAULT_BUDGETS },
    income: [],
    expenses: [],
    installments: [],
    debts: [],
    savings: []
  };
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultState();
    return { ...defaultState(), ...JSON.parse(raw) };
  } catch {
    return defaultState();
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function money(n) {
  const v = Number(n) || 0;
  return '৳' + v.toLocaleString('en-BD', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
}

function parseDate(s) {
  if (!s) return null;
  const d = new Date(s + (s.length === 10 ? 'T00:00:00' : ''));
  return Number.isNaN(d.getTime()) ? null : d;
}

function isoDate(d = new Date()) {
  const x = d instanceof Date ? d : new Date(d);
  const y = x.getFullYear();
  const m = String(x.getMonth() + 1).padStart(2, '0');
  const day = String(x.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function monthBounds(year, month) {
  const start = new Date(year, month - 1, 1);
  const end = new Date(year, month, 1);
  return { start, end, days: new Date(year, month, 0).getDate() };
}

function inActiveMonth(dateStr) {
  const d = parseDate(dateStr);
  if (!d) return false;
  const { start, end } = monthBounds(state.settings.year, state.settings.month);
  return d >= start && d < end;
}

function monthLabel() {
  return new Date(state.settings.year, state.settings.month - 1, 1)
    .toLocaleString('en-US', { month: 'long', year: 'numeric' });
}

/* ========== CALCULATIONS ========== */

function compute() {
  const { year, month, monthlyIncome, fixedCosts, emergencyBalance, budgetAlertPct } = state.settings;
  const { days } = monthBounds(year, month);

  const incomeRows = state.income.filter((r) => inActiveMonth(r.date));
  const expenseRows = state.expenses.filter((r) => inActiveMonth(r.date));

  const totalIncome = incomeRows.reduce((s, r) => s + (Number(r.amount) || 0), 0);
  const totalExpense = expenseRows.reduce((s, r) => s + (Number(r.actual) || 0), 0);
  const balance = totalIncome - totalExpense;

  const budgetRows = EXPENSE_CATEGORIES.map((cat) => {
    const budget = Number(state.budgets[cat]) || 0;
    const used = expenseRows
      .filter((e) => e.category === cat)
      .reduce((s, e) => s + (Number(e.actual) || 0), 0);
    const remaining = budget - used;
    const pct = budget > 0 ? used / budget : 0;
    let status = '—';
    if (budget > 0) {
      if (remaining < 0) status = 'over';
      else if (pct >= 0.9) status = 'near';
      else status = 'ok';
    }
    return { category: cat, budget, used, remaining, pct, status };
  });

  const totalBudget = budgetRows.reduce((s, r) => s + r.budget, 0);
  const budgetRemaining = totalBudget - totalExpense;
  const overBudget = Math.max(0, totalExpense - totalBudget);

  const savingsBalance = state.savings.reduce(
    (bal, r) => bal + (Number(r.deposit) || 0) - (Number(r.withdraw) || 0),
    0
  );

  const installmentLeft = state.installments.reduce((s, r) => {
    const rem = (Number(r.amount) || 0) - (Number(r.paid) || 0);
    return s + Math.max(0, rem);
  }, 0);

  const paidInst = state.installments.filter((r) => (Number(r.paid) || 0) >= (Number(r.amount) || 0) && r.amount).length;
  const totalInst = state.installments.filter((r) => r.amount).length;

  const debtRemaining = state.debts.reduce((s, r) => {
    const rem = (Number(r.amount) || 0) - (Number(r.paid) || 0);
    return s + Math.max(0, rem);
  }, 0);

  const today = isoDate();
  const todaySpend = state.expenses
    .filter((e) => e.date === today)
    .reduce((s, e) => s + (Number(e.actual) || 0), 0);

  const weekStart = (() => {
    const d = new Date();
    const day = (d.getDay() + 6) % 7;
    d.setHours(0, 0, 0, 0);
    d.setDate(d.getDate() - day);
    return d;
  })();
  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekEnd.getDate() + 7);
  const weekSpend = state.expenses.reduce((s, e) => {
    const d = parseDate(e.date);
    if (d && d >= weekStart && d < weekEnd) return s + (Number(e.actual) || 0);
    return s;
  }, 0);

  const safeDaily = Math.max(0, (monthlyIncome - fixedCosts) / days);

  let highest = null;
  expenseRows.forEach((e) => {
    if (!highest || Number(e.actual) > Number(highest.actual)) highest = e;
  });

  const topCat = budgetRows.slice().sort((a, b) => b.used - a.used)[0];

  const daily = Array.from({ length: days }, (_, i) => ({ day: i + 1, amount: 0 }));
  expenseRows.forEach((e) => {
    const d = parseDate(e.date);
    if (d) daily[d.getDate() - 1].amount += Number(e.actual) || 0;
  });

  const byCategory = budgetRows.filter((r) => r.used > 0).map((r) => ({ cat: r.category, amount: r.used }));

  const health = calcHealth(expenseRows, budgetRows, totalBudget, totalExpense, days);

  const alerts = [];
  if (balance < emergencyBalance) alerts.push({ type: 'danger', text: '⚠ Low Cash — Balance below emergency threshold.' });
  if (overBudget > 0) alerts.push({ type: 'danger', text: `🚨 You spent ${money(overBudget)} extra this month.` });
  if (budgetRemaining >= 0 && totalBudget > 0 && budgetRemaining / totalBudget < budgetAlertPct / 100) {
    alerts.push({ type: 'warn', text: `🟠 Budget Alert — less than ${budgetAlertPct}% of budget remaining.` });
  }
  if (todaySpend > safeDaily) {
    alerts.push({ type: 'danger', text: `🔴 Today's spending exceeded your limit by ${money(todaySpend - safeDaily)}.` });
  } else if (todaySpend > 0) {
    alerts.push({ type: 'info', text: "🟢 Today's spending is within your safe daily limit." });
  }

  return {
    totalIncome, totalExpense, balance, totalBudget, budgetRemaining, overBudget,
    savingsBalance, installmentLeft, debtRemaining, paidInst, totalInst,
    todaySpend, weekSpend, safeDaily, highest, topCat, daily, byCategory,
    budgetRows, expenseRows, incomeRows, health, alerts, days
  };
}

function calcHealth(expenseRows, budgetRows, totalBudget, totalExpense, days) {
  const { year, month, monthlyIncome } = state.settings;
  const today = new Date();
  const dayOfMonth = (today.getFullYear() === year && today.getMonth() + 1 === month)
    ? today.getDate()
    : days;

  let budgetPts = 30;
  if (totalBudget > 0) {
    const ratio = totalExpense / totalBudget;
    if (ratio <= 0.85) budgetPts = 30;
    else if (ratio <= 1) budgetPts = Math.round(30 - ((ratio - 0.85) / 0.15) * 8);
    else budgetPts = Math.max(0, Math.round(12 - (ratio - 1) * 40));
  }

  const logged = new Set(expenseRows.map((e) => e.date).filter(Boolean));
  const loggingPts = dayOfMonth > 0 ? Math.round(20 * Math.min(1, logged.size / dayOfMonth)) : 0;

  let overCount = 0;
  let lineCount = 0;
  expenseRows.forEach((e) => {
    const b = Number(e.budget);
    const a = Number(e.actual);
    if (e.budget === '' || e.budget == null || e.actual === '' || e.actual == null) return;
    lineCount++;
    if (a > b) overCount++;
  });
  const overspendPts = lineCount ? Math.round(20 * (1 - overCount / lineCount)) : 20;

  const { start, end } = monthBounds(year, month);
  const depositSum = state.savings.reduce((s, r) => {
    const d = parseDate(r.date);
    if (d && d >= start && d < end) return s + (Number(r.deposit) || 0);
    return s;
  }, 0);
  const saveGoal = (monthlyIncome || 30000) * 0.1;
  const savingsPts = saveGoal > 0 ? Math.round(15 * Math.min(1, depositSum / saveGoal)) : (depositSum > 0 ? 15 : 0);

  let good = 0;
  let totalI = 0;
  state.installments.forEach((r) => {
    if (!r.amount) return;
    totalI++;
    const rem = (Number(r.amount) || 0) - (Number(r.paid) || 0);
    const due = parseDate(r.dueDate);
    const overdue = rem > 0 && due && due < new Date(isoDate());
    if (!overdue) good++;
  });
  const installPts = totalI ? Math.round(15 * (good / totalI)) : 15;

  const total = Math.max(0, Math.min(100, budgetPts + loggingPts + overspendPts + savingsPts + installPts));
  let band = { label: 'Critical', color: '#b42318', emoji: '🔴' };
  if (total >= 90) band = { label: 'Excellent', color: '#1b7a45', emoji: '🟢' };
  else if (total >= 70) band = { label: 'Good', color: '#a16207', emoji: '🟡' };
  else if (total >= 50) band = { label: 'Needs Attention', color: '#c2410c', emoji: '🟠' };

  return {
    total,
    band,
    parts: { budget: budgetPts, logging: loggingPts, overspend: overspendPts, savings: savingsPts, installments: installPts }
  };
}

function expenseStatus(budget, actual) {
  if (budget === '' || budget == null || actual === '' || actual == null) return { key: '', label: '—' };
  const b = Number(budget);
  const a = Number(actual);
  if (a > b) return { key: 'over', label: '🔴 OVER' };
  if (a < b) return { key: 'under', label: '🟢 UNDER' };
  return { key: 'perfect', label: '🟡 PERFECT' };
}

function installmentStatus(row) {
  const amount = Number(row.amount) || 0;
  const paid = Number(row.paid) || 0;
  if (!amount) return { key: '', label: '—' };
  if (paid >= amount) return { key: 'done', label: '🟢 Paid' };
  if (paid > 0) return { key: 'partial', label: '🟠 Partial' };
  const due = parseDate(row.dueDate);
  if (due && due < new Date(isoDate())) return { key: 'overdue', label: '🔴 Overdue' };
  return { key: 'pending', label: '🔵 Pending' };
}

function debtStatus(row) {
  const amount = Number(row.amount) || 0;
  const paid = Number(row.paid) || 0;
  if (!amount) return { key: '', label: '—' };
  if (paid >= amount) return { key: 'done', label: 'Completed' };
  if (paid > 0) return { key: 'partial', label: 'Partial' };
  return { key: 'pending', label: 'Pending' };
}

/* ========== UI HELPERS ========== */

function toast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.hidden = false;
  clearTimeout(toast._t);
  toast._t = setTimeout(() => { el.hidden = true; }, 2200);
}

function destroyCharts() {
  Object.values(charts).forEach((c) => c.destroy());
  charts = {};
}

function renderAlerts(alerts) {
  const box = document.getElementById('alerts');
  if (!alerts.length) {
    box.hidden = true;
    box.innerHTML = '';
    return;
  }
  box.hidden = false;
  box.innerHTML = alerts.map((a) => `<div class="alert ${a.type}">${a.text}</div>`).join('');
}

/* ========== VIEWS ========== */

function render() {
  const [title, sub] = TITLES[currentView];
  document.getElementById('viewTitle').textContent = title;
  document.getElementById('viewSubtitle').textContent = sub + ' · ' + monthLabel();
  document.getElementById('activeMonth').value =
    `${state.settings.year}-${String(state.settings.month).padStart(2, '0')}`;

  document.querySelectorAll('.nav-btn').forEach((b) => {
    b.classList.toggle('active', b.dataset.view === currentView);
  });

  destroyCharts();
  const stats = compute();
  renderAlerts(currentView === 'dashboard' || currentView === 'reports' ? stats.alerts : []);

  const root = document.getElementById('viewRoot');
  const map = {
    dashboard: renderDashboard,
    income: renderIncome,
    expense: renderExpense,
    budget: renderBudget,
    installment: renderInstallment,
    debt: renderDebt,
    savings: renderSavings,
    reports: renderReports,
    settings: renderSettings
  };
  root.innerHTML = map[currentView](stats);
  bindViewEvents(stats);
  if (currentView === 'dashboard') mountDashboardCharts(stats);
}

function renderDashboard(s) {
  const pctBudget = s.totalBudget ? Math.min(1, s.totalExpense / s.totalBudget) : 0;
  const pctInst = s.totalInst ? s.paidInst / s.totalInst : 0;
  const pctSave = Math.min(1, s.savingsBalance / Math.max(1, state.settings.monthlyIncome * 0.2));
  const p = s.health.parts;

  return `
  <div class="grid two" style="margin-bottom:0.9rem">
    <div class="panel health" style="--score:${s.health.total};--score-color:${s.health.band.color}">
      <div class="health-score"><span>${s.health.total}</span></div>
      <div class="health-meta">
        <h3>${s.health.band.emoji} ${s.health.band.label}</h3>
        <p>Financial Health Score · ${monthLabel()}</p>
        <div class="health-parts">
          Budget ${p.budget}/30 · Logging ${p.logging}/20 · Overspend ${p.overspend}/20<br/>
          Savings ${p.savings}/15 · Installments ${p.installments}/15
        </div>
      </div>
    </div>
    <div class="panel">
      <div class="progress-list">
        <div class="progress-row">
          <label><span>Monthly Budget</span><span>${Math.round(pctBudget * 100)}%</span></label>
          <div class="bar"><i style="width:${pctBudget * 100}%"></i></div>
        </div>
        <div class="progress-row">
          <label><span>Installment</span><span>${s.paidInst}/${s.totalInst || 0} Paid</span></label>
          <div class="bar"><i style="width:${pctInst * 100}%;background:linear-gradient(90deg,#c2410c,#f59e0b)"></i></div>
        </div>
        <div class="progress-row">
          <label><span>Savings Goal (20% income)</span><span>${Math.round(pctSave * 100)}%</span></label>
          <div class="bar"><i style="width:${pctSave * 100}%;background:linear-gradient(90deg,#0f766e,#2dd4bf)"></i></div>
        </div>
        <div class="insight" style="border:0;padding-top:0.4rem">
          <span>Safe Daily Budget</span><strong>${money(s.safeDaily)}/day</strong>
        </div>
      </div>
    </div>
  </div>

  <div class="grid kpi" style="margin-bottom:0.9rem">
    <div class="panel kpi income"><div class="label">💰 Total Income</div><div class="value">${money(s.totalIncome)}</div></div>
    <div class="panel kpi expense"><div class="label">💸 Total Expense</div><div class="value">${money(s.totalExpense)}</div></div>
    <div class="panel kpi balance"><div class="label">🏦 Current Balance</div><div class="value">${money(s.balance)}</div></div>
    <div class="panel kpi budget"><div class="label">📉 Budget Remaining</div><div class="value">${money(s.budgetRemaining)}</div></div>
    <div class="panel kpi over"><div class="label">🔴 Over Budget</div><div class="value">${money(s.overBudget)}</div></div>
    <div class="panel kpi save"><div class="label">🟢 Savings</div><div class="value">${money(s.savingsBalance)}</div></div>
    <div class="panel kpi inst"><div class="label">🟠 Installment Left</div><div class="value">${money(s.installmentLeft)}</div></div>
    <div class="panel kpi debt"><div class="label">🔵 Debt Remaining</div><div class="value">${money(s.debtRemaining)}</div></div>
  </div>

  <div class="grid three" style="margin-bottom:0.9rem">
    <div class="panel"><h3 style="margin:0 0 0.5rem;font-family:var(--display)">Expense By Category</h3><div class="chart-box"><canvas id="chartPie"></canvas></div></div>
    <div class="panel"><h3 style="margin:0 0 0.5rem;font-family:var(--display)">Daily Expense</h3><div class="chart-box"><canvas id="chartLine"></canvas></div></div>
    <div class="panel"><h3 style="margin:0 0 0.5rem;font-family:var(--display)">Income vs Expense</h3><div class="chart-box"><canvas id="chartCol"></canvas></div></div>
  </div>

  <div class="panel">
    <h3 style="margin:0 0 0.6rem;font-family:var(--display)">Smart Insights</h3>
    <div class="insight-list">
      <div class="insight"><span>Largest Expense</span><strong>${s.highest ? `${s.highest.category} · ${money(s.highest.actual)}` : '—'}</strong></div>
      <div class="insight"><span>Today's Spending</span><strong>${money(s.todaySpend)}</strong></div>
      <div class="insight"><span>This Week</span><strong>${money(s.weekSpend)}</strong></div>
      <div class="insight"><span>Top Category</span><strong>${s.topCat && s.topCat.used ? `${s.topCat.category} · ${money(s.topCat.used)}` : '—'}</strong></div>
    </div>
  </div>
  <p class="footer-note">Money is a tool. Every Taka must have a purpose. Spend intentionally. Save consistently.<br/>Built for Biozed Hossain · Version 1.0</p>`;
}

function mountDashboardCharts(s) {
  const pieData = s.byCategory.length ? s.byCategory : [{ cat: 'No data', amount: 1 }];
  charts.pie = new Chart(document.getElementById('chartPie'), {
    type: 'doughnut',
    data: {
      labels: pieData.map((x) => x.cat),
      datasets: [{
        data: pieData.map((x) => x.amount),
        backgroundColor: ['#0d7a5f', '#1d4e89', '#b45309', '#b42318', '#0f766e', '#4d7c0f', '#7c2d12', '#334155', '#0891b2', '#a16207', '#be123c']
      }]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 10 } } } } }
  });

  charts.line = new Chart(document.getElementById('chartLine'), {
    type: 'line',
    data: {
      labels: s.daily.map((d) => d.day),
      datasets: [{
        label: 'Expense',
        data: s.daily.map((d) => d.amount),
        borderColor: '#b42318',
        backgroundColor: 'rgba(180,35,24,0.12)',
        fill: true,
        tension: 0.35,
        pointRadius: 2
      }]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
  });

  charts.col = new Chart(document.getElementById('chartCol'), {
    type: 'bar',
    data: {
      labels: ['This Month'],
      datasets: [
        { label: 'Income', data: [s.totalIncome], backgroundColor: '#1b7a45' },
        { label: 'Expense', data: [s.totalExpense], backgroundColor: '#b42318' }
      ]
    },
    options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true } } }
  });
}

function tableActions(id) {
  return `<div class="actions">
    <button type="button" class="btn ghost sm" data-edit="${id}">Edit</button>
    <button type="button" class="btn danger sm" data-del="${id}">Del</button>
  </div>`;
}

function renderIncome() {
  const rows = state.income.slice().sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  const body = rows.length
    ? rows.map((r) => `<tr>
        <td>${r.date || ''}</td><td>${r.source || ''}</td><td>${r.description || ''}</td>
        <td class="num">${money(r.amount)}</td><td>${r.payment || ''}</td><td>${r.notes || ''}</td>
        <td>${tableActions(r.id)}</td>
      </tr>`).join('')
    : `<tr><td colspan="7" class="empty">No income yet. Click Add Income.</td></tr>`;

  return `<div class="panel">
    <div class="section-head">
      <h3>Income entries</h3>
      <button type="button" class="btn primary" data-add="income">+ Add Income</button>
    </div>
    <div class="table-wrap"><table class="data">
      <thead><tr><th>Date</th><th>Source</th><th>Description</th><th class="num">Amount</th><th>Payment</th><th>Notes</th><th></th></tr></thead>
      <tbody>${body}</tbody>
    </table></div>
  </div>`;
}

function renderExpense() {
  const rows = state.expenses.slice().sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  const body = rows.length
    ? rows.map((r) => {
      const diff = (Number(r.budget) || 0) - (Number(r.actual) || 0);
      const st = expenseStatus(r.budget, r.actual);
      const rowClass = st.key === 'over' ? 'row-over' : st.key === 'perfect' ? 'row-near' : st.key === 'under' ? 'row-under' : '';
      return `<tr class="${rowClass}">
        <td>${r.date || ''}</td><td>${r.category || ''}</td><td>${r.subCategory || ''}</td><td>${r.description || ''}</td>
        <td class="num">${money(r.budget)}</td><td class="num">${money(r.actual)}</td>
        <td class="num">${money(diff)}</td>
        <td><span class="badge ${st.key}">${st.label}</span></td>
        <td>${r.payment || ''}</td><td>${r.notes || ''}</td>
        <td>${tableActions(r.id)}</td>
      </tr>`;
    }).join('')
    : `<tr><td colspan="11" class="empty">No expenses yet. Click Add Expense.</td></tr>`;

  return `<div class="panel">
    <div class="section-head">
      <h3>Expense tracker</h3>
      <button type="button" class="btn primary" data-add="expense">+ Add Expense</button>
    </div>
    <div class="table-wrap"><table class="data">
      <thead><tr>
        <th>Date</th><th>Category</th><th>Sub</th><th>Description</th>
        <th class="num">Budget</th><th class="num">Actual</th><th class="num">Diff</th><th>Status</th><th>Payment</th><th>Notes</th><th></th>
      </tr></thead>
      <tbody>${body}</tbody>
    </table></div>
  </div>`;
}

function renderBudget(s) {
  const body = s.budgetRows.map((r) => {
    const rowClass = r.status === 'over' ? 'row-over' : r.status === 'near' ? 'row-near' : r.status === 'ok' ? 'row-under' : '';
    const badge = r.status === 'over' ? '<span class="badge over">🔴 Over</span>'
      : r.status === 'near' ? '<span class="badge near">🟡 Near Limit</span>'
      : r.status === 'ok' ? '<span class="badge ok">🟢 OK</span>' : '—';
    return `<tr class="${rowClass}">
      <td class="col-cat">${r.category}</td>
      <td class="num col-budget"><input type="number" min="0" step="1" class="budget-input" data-budget-cat="${r.category}" value="${r.budget}" /></td>
      <td class="num">${money(r.used)}</td>
      <td class="num">${money(r.remaining)}</td>
      <td class="num">${r.budget ? Math.round(r.pct * 100) + '%' : '—'}</td>
      <td class="col-status">${badge}</td>
    </tr>`;
  }).join('');

  return `<div class="panel">
    <div class="section-head">
      <h3>Category budgets · ${monthLabel()}</h3>
      <button type="button" class="btn primary" id="saveBudgets">Save budgets</button>
    </div>
    <div class="table-wrap"><table class="data budget-table">
      <thead><tr>
        <th class="col-cat">Category</th>
        <th class="num col-budget">Monthly Budget</th>
        <th class="num">Used</th>
        <th class="num">Remaining</th>
        <th class="num">% Used</th>
        <th class="col-status">Status</th>
      </tr></thead>
      <tbody>${body}
        <tr class="total-row">
          <td class="col-cat"><strong>TOTAL</strong></td>
          <td class="num col-budget"><strong>${money(s.totalBudget)}</strong></td>
          <td class="num"><strong>${money(s.totalExpense)}</strong></td>
          <td class="num"><strong>${money(s.budgetRemaining)}</strong></td>
          <td class="num"><strong>${s.totalBudget ? Math.round((s.totalExpense / s.totalBudget) * 100) + '%' : '—'}</strong></td>
          <td class="col-status"></td>
        </tr>
      </tbody>
    </table></div>
  </div>`;
}

function renderInstallment() {
  const rows = state.installments;
  const body = rows.length
    ? rows.map((r) => {
      const rem = (Number(r.amount) || 0) - (Number(r.paid) || 0);
      const st = installmentStatus(r);
      return `<tr>
        <td>${r.label || ''}</td><td class="num">${money(r.amount)}</td><td>${r.dueDate || ''}</td>
        <td class="num">${money(r.paid)}</td><td class="num">${money(rem)}</td>
        <td><span class="badge ${st.key}">${st.label}</span></td>
        <td>${tableActions(r.id)}</td>
      </tr>`;
    }).join('')
    : `<tr><td colspan="7" class="empty">No installments yet.</td></tr>`;

  return `<div class="panel">
    <div class="section-head">
      <h3>Installments</h3>
      <button type="button" class="btn primary" data-add="installment">+ Add Installment</button>
    </div>
    <div class="table-wrap"><table class="data">
      <thead><tr><th>Item / Month</th><th class="num">Amount</th><th>Due Date</th><th class="num">Paid</th><th class="num">Remaining</th><th>Status</th><th></th></tr></thead>
      <tbody>${body}</tbody>
    </table></div>
  </div>`;
}

function renderDebt() {
  const rows = state.debts;
  const body = rows.length
    ? rows.map((r) => {
      const rem = (Number(r.amount) || 0) - (Number(r.paid) || 0);
      const st = debtStatus(r);
      return `<tr>
        <td>${r.person || ''}</td><td class="num">${money(r.amount)}</td><td class="num">${money(r.paid)}</td>
        <td class="num">${money(rem)}</td><td>${r.date || ''}</td>
        <td><span class="badge ${st.key}">${st.label}</span></td>
        <td>${tableActions(r.id)}</td>
      </tr>`;
    }).join('')
    : `<tr><td colspan="7" class="empty">No debts yet.</td></tr>`;

  return `<div class="panel">
    <div class="section-head">
      <h3>Debt tracker</h3>
      <button type="button" class="btn primary" data-add="debt">+ Add Debt</button>
    </div>
    <div class="table-wrap"><table class="data">
      <thead><tr><th>Person</th><th class="num">Amount</th><th class="num">Paid</th><th class="num">Remaining</th><th>Date</th><th>Status</th><th></th></tr></thead>
      <tbody>${body}</tbody>
    </table></div>
  </div>`;
}

function renderSavings() {
  let bal = 0;
  const rows = state.savings.slice().sort((a, b) => (a.date || '').localeCompare(b.date || ''));
  const body = rows.length
    ? rows.map((r) => {
      bal += (Number(r.deposit) || 0) - (Number(r.withdraw) || 0);
      return `<tr>
        <td>${r.date || ''}</td><td>${r.goal || ''}</td>
        <td class="num">${money(r.deposit)}</td><td class="num">${money(r.withdraw)}</td>
        <td class="num">${money(bal)}</td>
        <td>${tableActions(r.id)}</td>
      </tr>`;
    }).join('')
    : `<tr><td colspan="6" class="empty">No savings entries yet.</td></tr>`;

  return `<div class="panel">
    <div class="section-head">
      <h3>Savings · Balance ${money(bal)}</h3>
      <button type="button" class="btn primary" data-add="savings">+ Add Entry</button>
    </div>
    <div class="table-wrap"><table class="data">
      <thead><tr><th>Date</th><th>Goal</th><th class="num">Deposit</th><th class="num">Withdraw</th><th class="num">Balance</th><th></th></tr></thead>
      <tbody>${body}</tbody>
    </table></div>
  </div>`;
}

function renderReports(s) {
  const avgDaily = s.days ? s.totalExpense / s.days : 0;
  return `<div class="grid two">
    <div class="panel">
      <h3 style="margin:0 0 0.7rem;font-family:var(--display)">Monthly Report</h3>
      <div class="insight-list">
        <div class="insight"><span>Total Income</span><strong>${money(s.totalIncome)}</strong></div>
        <div class="insight"><span>Total Expense</span><strong>${money(s.totalExpense)}</strong></div>
        <div class="insight"><span>Net Balance</span><strong>${money(s.balance)}</strong></div>
        <div class="insight"><span>Budget Used %</span><strong>${s.totalBudget ? Math.round((s.totalExpense / s.totalBudget) * 100) + '%' : '—'}</strong></div>
      </div>
    </div>
    <div class="panel">
      <h3 style="margin:0 0 0.7rem;font-family:var(--display)">Weekly & Insights</h3>
      <div class="insight-list">
        <div class="insight"><span>This Week Expense</span><strong>${money(s.weekSpend)}</strong></div>
        <div class="insight"><span>Today's Expense</span><strong>${money(s.todaySpend)}</strong></div>
        <div class="insight"><span>Highest Expense</span><strong>${s.highest ? money(s.highest.actual) + ' · ' + (s.highest.description || s.highest.category) : '—'}</strong></div>
        <div class="insight"><span>Top Category</span><strong>${s.topCat && s.topCat.used ? s.topCat.category + ' · ' + money(s.topCat.used) : '—'}</strong></div>
        <div class="insight"><span>Avg Daily Expense</span><strong>${money(avgDaily)}</strong></div>
        <div class="insight"><span>Avg Weekly Expense</span><strong>${money(avgDaily * 7)}</strong></div>
        <div class="insight"><span>Safe Daily Budget</span><strong>${money(s.safeDaily)}/day</strong></div>
      </div>
    </div>
  </div>`;
}

function renderSettings() {
  const st = state.settings;
  return `<div class="panel" style="max-width:640px">
    <div class="section-head"><h3>Preferences</h3></div>
    <form id="settingsForm" class="form-grid">
      <div class="field"><label>Owner</label><input name="owner" value="${st.owner}" /></div>
      <div class="field"><label>Monthly Income (BDT)</label><input name="monthlyIncome" type="number" min="0" value="${st.monthlyIncome}" /></div>
      <div class="field"><label>Fixed Monthly Costs</label><input name="fixedCosts" type="number" min="0" value="${st.fixedCosts}" /></div>
      <div class="field"><label>Emergency Balance Alert</label><input name="emergencyBalance" type="number" min="0" value="${st.emergencyBalance}" /></div>
      <div class="field"><label>Budget Alert Threshold %</label><input name="budgetAlertPct" type="number" min="1" max="100" value="${st.budgetAlertPct}" /></div>
      <div class="field full"><button type="submit" class="btn primary">Save Settings</button></div>
    </form>
    <p class="muted" style="margin-top:1rem;font-size:0.85rem">Active month is controlled from the top bar. Data is saved in this browser (localStorage).</p>
  </div>`;
}

/* ========== FORMS ========== */

function openForm(type, item = null) {
  editCtx = { type, id: item?.id || null };
  const dialog = document.getElementById('formDialog');
  const fields = document.getElementById('dialogFields');
  document.getElementById('dialogTitle').textContent = (item ? 'Edit ' : 'Add ') + type;

  const opts = (arr, sel) => arr.map((o) => `<option value="${o}" ${o === sel ? 'selected' : ''}>${o}</option>`).join('');
  const today = isoDate();

  if (type === 'income') {
    fields.innerHTML = `<div class="form-grid">
      <div class="field"><label>Date</label><input name="date" type="date" required value="${item?.date || today}" /></div>
      <div class="field"><label>Source</label><select name="source">${opts(INCOME_SOURCES, item?.source || 'Salary')}</select></div>
      <div class="field full"><label>Description</label><input name="description" value="${item?.description || ''}" /></div>
      <div class="field"><label>Amount</label><input name="amount" type="number" min="0" step="1" required value="${item?.amount ?? ''}" /></div>
      <div class="field"><label>Payment</label><select name="payment">${opts(PAYMENT_METHODS, item?.payment || 'Bank')}</select></div>
      <div class="field full"><label>Notes</label><input name="notes" value="${item?.notes || ''}" /></div>
    </div>`;
  } else if (type === 'expense') {
    const cat = item?.category || 'Market';
    fields.innerHTML = `<div class="form-grid">
      <div class="field"><label>Date</label><input name="date" type="date" required value="${item?.date || today}" /></div>
      <div class="field"><label>Category</label><select name="category" id="expCategory">${opts(EXPENSE_CATEGORIES, cat)}</select></div>
      <div class="field"><label>Sub Category</label><select name="subCategory" id="expSub">${opts(SUB_CATEGORIES[cat] || ['Other'], item?.subCategory)}</select></div>
      <div class="field"><label>Description</label><input name="description" value="${item?.description || ''}" /></div>
      <div class="field"><label>Budget (planned)</label><input name="budget" type="number" min="0" step="1" value="${item?.budget ?? ''}" /></div>
      <div class="field"><label>Actual</label><input name="actual" type="number" min="0" step="1" required value="${item?.actual ?? ''}" /></div>
      <div class="field"><label>Payment</label><select name="payment">${opts(PAYMENT_METHODS, item?.payment || 'Cash')}</select></div>
      <div class="field"><label>Notes</label><input name="notes" value="${item?.notes || ''}" /></div>
    </div>`;
  } else if (type === 'installment') {
    fields.innerHTML = `<div class="form-grid">
      <div class="field full"><label>Item / Month Label</label><input name="label" required value="${item?.label || ''}" /></div>
      <div class="field"><label>Amount</label><input name="amount" type="number" min="0" required value="${item?.amount ?? ''}" /></div>
      <div class="field"><label>Due Date</label><input name="dueDate" type="date" value="${item?.dueDate || today}" /></div>
      <div class="field"><label>Paid</label><input name="paid" type="number" min="0" value="${item?.paid ?? 0}" /></div>
    </div>`;
  } else if (type === 'debt') {
    fields.innerHTML = `<div class="form-grid">
      <div class="field"><label>Person</label><input name="person" required value="${item?.person || ''}" /></div>
      <div class="field"><label>Date</label><input name="date" type="date" value="${item?.date || today}" /></div>
      <div class="field"><label>Amount</label><input name="amount" type="number" min="0" required value="${item?.amount ?? ''}" /></div>
      <div class="field"><label>Paid</label><input name="paid" type="number" min="0" value="${item?.paid ?? 0}" /></div>
    </div>`;
  } else if (type === 'savings') {
    fields.innerHTML = `<div class="form-grid">
      <div class="field"><label>Date</label><input name="date" type="date" required value="${item?.date || today}" /></div>
      <div class="field"><label>Goal</label><input name="goal" value="${item?.goal || 'Emergency Fund'}" /></div>
      <div class="field"><label>Deposit</label><input name="deposit" type="number" min="0" value="${item?.deposit ?? ''}" /></div>
      <div class="field"><label>Withdraw</label><input name="withdraw" type="number" min="0" value="${item?.withdraw ?? ''}" /></div>
    </div>`;
  }

  dialog.showModal();

  const catSel = document.getElementById('expCategory');
  if (catSel) {
    catSel.addEventListener('change', () => {
      const sub = document.getElementById('expSub');
      const list = SUB_CATEGORIES[catSel.value] || ['Other'];
      sub.innerHTML = opts(list, list[0]);
    });
  }
}

function collectForm() {
  const form = document.getElementById('entityForm');
  const data = {};
  form.querySelectorAll('[name]').forEach((el) => {
    data[el.name] = el.type === 'number' ? (el.value === '' ? '' : Number(el.value)) : el.value.trim();
  });
  return data;
}

function saveForm(e) {
  e.preventDefault();
  if (!editCtx) return;
  const data = collectForm();
  const { type, id } = editCtx;
  const listKey = {
    income: 'income',
    expense: 'expenses',
    installment: 'installments',
    debt: 'debts',
    savings: 'savings'
  }[type];

  if (id) {
    const idx = state[listKey].findIndex((x) => x.id === id);
    if (idx >= 0) state[listKey][idx] = { ...state[listKey][idx], ...data };
  } else {
    state[listKey].push({ id: uid(), ...data });
  }
  saveState();
  document.getElementById('formDialog').close();
  toast('Saved');
  render();
}

function deleteItem(type, id) {
  if (!confirm('Delete this entry?')) return;
  const listKey = {
    income: 'income',
    expense: 'expenses',
    installment: 'installments',
    debt: 'debts',
    savings: 'savings'
  }[type];
  state[listKey] = state[listKey].filter((x) => x.id !== id);
  saveState();
  toast('Deleted');
  render();
}

function bindViewEvents() {
  const root = document.getElementById('viewRoot');

  root.querySelector('[data-add]')?.addEventListener('click', (e) => {
    openForm(e.currentTarget.dataset.add);
  });

  root.querySelectorAll('[data-edit]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const id = btn.dataset.edit;
      const type = currentView === 'expense' ? 'expense'
        : currentView === 'income' ? 'income'
        : currentView === 'installment' ? 'installment'
        : currentView === 'debt' ? 'debt'
        : 'savings';
      const listKey = { income: 'income', expense: 'expenses', installment: 'installments', debt: 'debts', savings: 'savings' }[type];
      const item = state[listKey].find((x) => x.id === id);
      if (item) openForm(type, item);
    });
  });

  root.querySelectorAll('[data-del]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const type = currentView === 'expense' ? 'expense'
        : currentView === 'income' ? 'income'
        : currentView === 'installment' ? 'installment'
        : currentView === 'debt' ? 'debt'
        : 'savings';
      deleteItem(type, btn.dataset.del);
    });
  });

  document.getElementById('saveBudgets')?.addEventListener('click', () => {
    root.querySelectorAll('[data-budget-cat]').forEach((input) => {
      state.budgets[input.dataset.budgetCat] = Number(input.value) || 0;
    });
    saveState();
    toast('Budgets saved');
    render();
  });

  document.getElementById('settingsForm')?.addEventListener('submit', (e) => {
    e.preventDefault();
    const fd = new FormData(e.target);
    state.settings.owner = fd.get('owner');
    state.settings.monthlyIncome = Number(fd.get('monthlyIncome')) || 0;
    state.settings.fixedCosts = Number(fd.get('fixedCosts')) || 0;
    state.settings.emergencyBalance = Number(fd.get('emergencyBalance')) || 0;
    state.settings.budgetAlertPct = Number(fd.get('budgetAlertPct')) || 20;
    saveState();
    toast('Settings saved');
    render();
  });
}

/* ========== SAMPLE / EXPORT ========== */

function addSampleData() {
  const { year, month } = state.settings;
  const d = (day) => `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;

  state.income.push(
    { id: uid(), date: d(1), source: 'Salary', description: 'Monthly salary', amount: 30000, payment: 'Bank', notes: '' },
    { id: uid(), date: d(10), source: 'Freelance', description: 'Side project', amount: 5000, payment: 'bKash', notes: '' }
  );
  state.expenses.push(
    { id: uid(), date: d(2), category: 'Market', subCategory: 'Grocery', description: 'Rice', budget: 500, actual: 650, payment: 'Cash', notes: '' },
    { id: uid(), date: d(3), category: 'Transport', subCategory: 'CNG', description: 'Office', budget: 200, actual: 180, payment: 'Cash', notes: '' },
    { id: uid(), date: d(5), category: 'Restaurant', subCategory: 'Lunch', description: 'Lunch', budget: 300, actual: 280, payment: 'bKash', notes: '' },
    { id: uid(), date: d(8), category: 'Electricity', subCategory: 'Home Bill', description: 'DESCO', budget: 1500, actual: 1450, payment: 'Bank', notes: '' },
    { id: uid(), date: d(15), category: 'Market', subCategory: 'Grocery', description: 'Weekly market', budget: 2000, actual: 1850, payment: 'Cash', notes: '' },
    { id: uid(), date: d(20), category: 'Entertainment', subCategory: 'Subscription', description: 'Streaming', budget: 500, actual: 550, payment: 'Card', notes: '' }
  );
  state.installments.push(
    { id: uid(), label: 'Phone EMI — 6/9', amount: 3500, dueDate: d(5), paid: 3500 },
    { id: uid(), label: 'Phone EMI — 7/9', amount: 3500, dueDate: d(5), paid: 0 }
  );
  state.debts.push(
    { id: uid(), person: 'Rahim', amount: 5000, paid: 2000, date: d(1) }
  );
  state.savings.push(
    { id: uid(), date: d(1), goal: 'Emergency Fund', deposit: 3000, withdraw: 0 },
    { id: uid(), date: d(15), goal: 'Emergency Fund', deposit: 2000, withdraw: 0 }
  );
  saveState();
  toast('Sample data added');
  currentView = 'dashboard';
  render();
}

function exportData() {
  const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `my-wealth-health-${state.settings.year}-${state.settings.month}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
  toast('Exported');
}

function importData(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const data = JSON.parse(reader.result);
      state = { ...defaultState(), ...data };
      saveState();
      toast('Imported');
      render();
    } catch {
      toast('Invalid JSON file');
    }
  };
  reader.readAsText(file);
}

/* ========== INIT ========== */

function init() {
  const params = new URLSearchParams(location.search);
  if (params.get('view') && TITLES[params.get('view')]) {
    currentView = params.get('view');
  }

  function setSidebarOpen(open) {
    const sidebar = document.getElementById('sidebar');
    const backdrop = document.getElementById('sidebarBackdrop');
    sidebar.classList.toggle('open', open);
    if (backdrop) backdrop.hidden = !open;
  }

  document.getElementById('nav').addEventListener('click', (e) => {
    const btn = e.target.closest('.nav-btn');
    if (!btn) return;
    currentView = btn.dataset.view;
    setSidebarOpen(false);
    render();
  });

  document.getElementById('menuToggle').addEventListener('click', () => {
    const open = !document.getElementById('sidebar').classList.contains('open');
    setSidebarOpen(open);
  });

  document.getElementById('sidebarBackdrop')?.addEventListener('click', () => {
    setSidebarOpen(false);
  });

  document.getElementById('activeMonth').addEventListener('change', (e) => {
    const [y, m] = e.target.value.split('-').map(Number);
    state.settings.year = y;
    state.settings.month = m;
    saveState();
    render();
  });

  document.getElementById('btnSample').addEventListener('click', addSampleData);
  document.getElementById('btnExport').addEventListener('click', exportData);
  document.getElementById('btnImport').addEventListener('click', () => document.getElementById('importFile').click());
  document.getElementById('importFile').addEventListener('change', (e) => {
    if (e.target.files[0]) importData(e.target.files[0]);
    e.target.value = '';
  });

  document.getElementById('dialogClose').addEventListener('click', () => document.getElementById('formDialog').close());
  document.getElementById('dialogCancel').addEventListener('click', () => document.getElementById('formDialog').close());
  document.getElementById('entityForm').addEventListener('submit', saveForm);

  render();
}

init();
