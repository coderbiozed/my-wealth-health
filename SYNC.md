# Sync Netlify ↔ Google Sheets

Website: https://my-wealth-health.netlify.app/

Edits in the website save into your Google Sheet through an Apps Script **Web App**.

## 1. Prepare the Sheet

1. Create a Google Sheet (or open your existing Biozed Finance Manager workbook)
2. **Extensions → Apps Script**
3. Paste all `.gs` files (including new **`Api.gs`**) — or paste updated `ALL_IN_ONE.gs`
4. Run **`setupFinanceManager`** once (authorize)
5. Optional: set a sync token in **Settings!B12**

## 2. Deploy Web App

1. In Apps Script: **Deploy → New deployment**
2. Type: **Web app**
3. Description: `My Wealth Health API`
4. Execute as: **Me**
5. Who has access: **Anyone**
6. **Deploy** → copy the Web App URL  
   (ends with `/exec`)

## 3. Connect the website

1. Open https://my-wealth-health.netlify.app/
2. Go to **Settings**
3. Paste **Sheet API URL**
4. Paste the same **Sync Token** if you set Settings!B12
5. Keep **Push edits to Sheet** checked
6. Click **Test connection** → should say Connected
7. Click **Save Settings**
8. Click **Save to Sheet now** (or add any expense — auto-saves after ~1s)

## 4. How sync works

| Action | What happens |
|--------|----------------|
| Add/Edit/Delete on website | Saves to browser + pushes to Sheet |
| **Load from Sheet** | Overwrites website data from Sheet |
| **Save to Sheet now** | Full push of current website data |
| Phone + laptop | Same Sheet = shared data (after both use the same API URL) |

## Notes

- First deploy: authorize Google when prompted
- If you change the script, **Deploy → Manage deployments → Edit → New version**
- Local data always stays in the browser as backup (Export JSON)
