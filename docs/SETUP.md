# Line Waste Log App — Setup (about 15 minutes)

## 1. Create the email + logging service
1. Create a new Google Sheet named **FC Line Waste Log – Submissions**.
2. **Extensions ▸ Apps Script** → delete the sample code → paste all of `apps-script/Code.gs`.
3. At the top, set:
   - `RECIPIENTS` — management email(s), comma-separated (work Outlook addresses are fine)
   - `SHARED_KEY` — any phrase; it must match the app
4. Click **Save**, choose `authorizeAndTest` in the function menu, click **Run**, and approve the permissions. A test email should arrive.
5. **Deploy ▸ New deployment ▸ ⚙ Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy**, then copy the **Web app URL** (ends in `/exec`).

> If you edit `Code.gs` later: **Deploy ▸ Manage deployments ▸ ✏ ▸ Version: New version**. The URL doesn't change.

## 2. Connect the app
Pick one:
- **Option A (recommended):** open `index.html` in a text editor and fill in `ENDPOINT_URL`, `SHARED_KEY` and `MANAGER_PIN` in the `CONFIG` block at the top of the script. Every device picks these up automatically.
- **Option B (per device):** open the app, tap **⚙︎**, enter the manager PIN (default `2468`), paste the URL, and tap **Send test**.

## 3. Put it on the kitchen tablets
Host the HTML file the same way as your other apps (GitHub Pages works well). Then on each tablet, open the link and use **Add to Home Screen** so it opens like an app.

## How it works each night
1. The chef enters their name, badge # and the business date.
2. The **Outlets** screen lists PFC, Pronto, Little Wok and Agave. Each one shows *Not started*, *In progress* or *✓ Done*.
3. For each outlet, the chef sets that outlet's closing time and counts every item (+/−, or type an amount). **Finish outlet** won't go through while any blank items remain.
4. **Review & submit** unlocks when all outlets are done. The chef reviews the count and taps **Submit tonight's count**.

## What management gets: one weekly report
- Every night's count fills that day's column on this week's **"Week M-D to M-D"** tab in the Google Sheet, laid out like the Excel template (Mon–Sun across the top).
- **After Sunday's count is submitted,** one email goes out covering Monday through Sunday. It includes:
  - the week's total cost, items wasted, units and days logged
  - any days with no count
  - the highest-cost outlet and highest-waste day
  - 10+ flags
  - an outlet-by-day table and the top 10 items by cost
  - each outlet's items day by day
  - all chef notes from the week
- If Sunday is corrected and re-sent, an **UPDATED** email goes out. If nobody submits Sunday, the report goes out with the first count of the next week and says Sunday was missing.
- To send a week's report early, open Apps Script and run `sendWeeklyReportNow`.
- To go back to a nightly email as well, set `DAILY_EMAIL = true` in the script.
- If Wi-Fi drops, the count is saved on the phone and sent automatically when the connection returns. Duplicates are ignored.

## Changing items or costs
Items, units and cost per unit are in the `STATIONS` list inside the HTML. They were copied from your template's *Waste Summary* tab. Edit names or costs there. Status colors follow the template's key: 🟢 under 5 units, 🟠 5–9, 🔴 10 or more (`AMBER_AT` / `RED_AT` in both files).
