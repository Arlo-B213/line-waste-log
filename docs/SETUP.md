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
3. For each outlet, the chef sets that outlet's closing time and counts every item (+/−, or type an amount). **Finish outlet** won't go through while any item is blank; blank items are highlighted, and one tap can set the rest to 0.
4. **Review & send report** stays locked until **all outlets are done**. The chef reviews everything and taps **Send report**, which sends **one email** to management covering every outlet.

- The email has the night's total cost, an outlet-by-outlet summary (closing time, units, cost, % of total, top $ item), 🔴 flags for anything at 10+ units, and a detailed breakdown for each outlet with the chef's notes.
- Everything is logged to the Google Sheet: a **Submissions** tab (one row per outlet per night) and a **Waste Log** tab (one row per wasted item).
- Progress is saved on the tablet as the chef goes. If the app gets closed partway through, it picks up where they left off.
- If Wi-Fi drops, the report is saved and sent automatically when the connection returns. Duplicates are ignored.
- A manager can clear an unfinished count under ⚙︎ → **Clear tonight's count**.

## Changing items or costs
Items, units and cost per unit are in the `STATIONS` list inside the HTML. They were copied from your template's *Waste Summary* tab. Edit names or costs there. Status colors follow the template's key: 🟢 under 5 units, 🟠 5–9, 🔴 10 or more (`AMBER_AT` / `RED_AT` in both files).
