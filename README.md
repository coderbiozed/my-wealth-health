# Biozed Finance Manager — My Wealth Health

Fully dynamic personal finance system for **Biozed Hossain** — Google Sheets backend + HTML frontend.

## HTML Frontend (easiest)

Open `index.html` in your browser (double-click or Live Server).

- Dashboard, Income, Expense, Budget, Installments, Debt, Savings, Reports, Settings
- Add / Edit / Delete with forms
- Charts, health score, alerts — all live
- Data saved in browser `localStorage`
- Export / Import JSON backup
- **Sample data** button to try it instantly

```bash
# optional local server
cd "my wealth health" && python3 -m http.server 8765
# then open http://localhost:8765
```

## Google Sheets Setup (5 minutes)

1. Open [Google Sheets](https://sheets.google.com) → **Blank spreadsheet**
2. Rename it to `Biozed Finance Manager` (or `My Wealth Health`)
3. Go to **Extensions → Apps Script**
4. Delete any default `Code.gs` content
5. Create these files in the Apps Script editor (File → New → Script) and paste matching contents from this folder:

| Apps Script file | Copy from |
|------------------|-----------|
| `Config.gs` | `Config.gs` |
| `Code.gs` | `Code.gs` |
| `Sheets.gs` | `Sheets.gs` |
| `Dashboard.gs` | `Dashboard.gs` |
| `HealthScore.gs` | `HealthScore.gs` |
| `Charts.gs` | `Charts.gs` |
| `SampleData.gs` | `SampleData.gs` |

6. Optional: File → Project settings → upload `appsscript.json` timezone (`Asia/Dhaka`)
7. Save → select function **`setupFinanceManager`** → click **Run**
8. Authorize the script when prompted
9. Return to the spreadsheet — all sheets are created

After setup, use the menu: **💰 Finance Manager**

---

## Workbook Structure

```
📁 Biozed Finance Manager
├── Dashboard          ← auto KPIs, alerts, health score, charts (no manual edit)
├── Income
├── Expense Tracker
├── Monthly Budget
├── Installment Tracker
├── Debt Tracker
├── Savings
├── Reports            ← auto monthly/weekly insights
├── Settings           ← active month, income, alert thresholds
└── Categories         ← powers all dropdowns
```

---

## How to Use Daily

1. **Settings** — set Active Year/Month, Monthly Income (e.g. 30,000), Fixed Costs
2. **Monthly Budget** — adjust category budgets (defaults included)
3. **Income** — log salary / freelance / etc.
4. **Expense Tracker** — log each spend:
   - `Budget` = planned for that purchase
   - `Actual` = what you paid
   - Status & Difference calculate automatically
5. Open **Dashboard** (or menu → Refresh Dashboard)

---

## Smart Features Included

| Feature | Behavior |
|---------|----------|
| Auto Status | 🔴 OVER / 🟢 UNDER / 🟡 PERFECT |
| Conditional formatting | Red / yellow (≥90%) / green on expenses & budget |
| Budget remaining alert | Orange when remaining &lt; threshold (default 20%) |
| Low cash alert | ⚠ when balance &lt; emergency amount (default ৳1,000) |
| Overspend alert | 🚨 extra BDT spent vs budget |
| Safe daily budget | `(Income − Fixed Costs) ÷ days in month` |
| Today over limit | Warns if today's spend &gt; safe daily budget |
| Financial Health Score | 0–100 from budget, logging, overspend, savings, installments |
| Charts | Pie (category) · Line (daily) · Column (income vs expense) |
| Progress bars | Budget % · Installments paid · Savings goal |

### Health Score bands

- 🟢 90–100 Excellent  
- 🟡 70–89 Good  
- 🟠 50–69 Needs Attention  
- 🔴 Below 50 Critical  

---

## Menu Actions

- **Setup / Rebuild Workbook** — rebuild structure (keeps your data rows)
- **Refresh Dashboard** — recalc health score + chart data
- **Rebuild Charts** — recreate pie / line / column charts
- **Add Sample Data** — demo rows for the active month
- **Set Active Month to Current** — jump Settings to today

---

## Footer

> Money is a tool.  
> Every Taka must have a purpose.  
> Spend intentionally. Save consistently.  
> Built for Biozed Hossain · Version 1.0

---

## Tips

- Change month in **Settings!B3/B4** — Dashboard & Reports follow automatically
- `_ChartData` is a hidden helper sheet — leave it alone
- Re-running Setup is safe; Income / Expense / Savings / Debt / Installment values are preserved
- For a new month: update Settings month, keep logging — history stays in the sheets
