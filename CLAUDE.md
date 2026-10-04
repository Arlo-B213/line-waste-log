# Line Waste Log — project guide for Claude

Chefs use this tablet app for the end-of-night line waste count across the Pechanga Resort Casino food court. They count **every outlet** and submit each night. Each submission is logged to a Google Sheet, including a Mon–Sun weekly tab. Management gets **one weekly email** after Sunday's count is submitted; there is no nightly email (`DAILY_EMAIL = false` in `Code.gs`).

## Layout
- `index.html` — the whole app in one self-contained file: HTML, CSS and vanilla JS with no build step. It is hosted on GitHub Pages and added to the tablets' home screens.
- `apps-script/Code.gs` — the Google Apps Script web app backend. `doPost` checks the shared key, removes duplicates by submission `id`, writes to the `Submissions` sheet (one row per outlet) and the `Waste Log` sheet (one row per wasted item), and sends the HTML email with `MailApp`.
- `docs/SETUP.md` — deployment steps for the manager.
- `docs/FC - Line Waste Log (template).xlsx` — the original Excel template. The item list, UOMs and cost/unit come from its *📊 Waste Summary* tab.
- `tools/e2e_test.py` — a Playwright end-to-end test. It fills all 4 outlets, checks that sending is blocked until all are done, and intercepts the POST.
- `tools/sheet_mock_test.js` — runs `Code.gs` against an in-memory Google Sheets mock: simulates a week of nightly posts (including a re-sent night, a duplicate, a new menu item and a week rollover), then prints the weekly tab and writes `mock_email.html`.
- `tools/render_email.js` — renders the email HTML from a saved payload by running `Code.gs` in Node with mocked `MailApp`/`Utilities`.

## Key rules / behavior
- Outlets: PFC Product (Pechanga Fried Chicken), Pronto Product, Little Wok, Agave. They are defined in `STATIONS` and `STATION_META` in `index.html`.
- An outlet is "done" only when every item has a number (0 allowed) and the chef tapped **Finish outlet**. **Review & send** stays disabled until all outlets are done.
- Status colors follow the template key: 🟢 <5 units, 🟠 5–9, 🔴 10+. The thresholds are `AMBER_AT`/`RED_AT`, set in both `index.html` and `Code.gs`, and must stay in sync.
- **Shifts:** the start screen picks Day, Swing or Closing (default guessed from the time: 5–14h Day, 14–20h Swing, else Closing). Closing is the required full count above. Day/Swing are optional logs of only what was thrown out: no outlet or item is required, and the payload carries `shift` plus only outlets/items with qty > 0. Each shift has its own session key (`wl_session` closing, `wl_shift_day`, `wl_shift_swing`). In `Code.gs`, a non-closing `shift` is written to the **Shift Log** tab only (never the week tab, Submissions or Waste Log) and appears in its own "Day & Swing shifts" section of the weekly email, excluded from the closing totals.
- **American Classic** is `shiftOnly` in `STATIONS`: it shows on Day/Swing logs only, never in the closing count. Its item list is empty until pricing arrives, so it shows as "Item list coming soon"; add items to its `cats` to enable it.
- Weight UOMs (Lbs/LBS/pan) step by 0.5 and each/slice step by 1.
- The session (all outlets for one business date) is kept in `localStorage` under `wl_session`. Unsent reports go in `wl_outbox` and are retried on `online` and on load. Before 5 AM, the business date defaults to the previous day.
- The POST uses `Content-Type: text/plain` so Apps Script needs no CORS preflight.
- ⚙︎ settings are behind `CONFIG.MANAGER_PIN`. They cover the endpoint URL, shared key, whether chefs see $ costs (off by default), retrying unsent reports, and clearing tonight's count.
- **Weekly log:** each post also fills a Mon–Sun tab named `Week M-D to M-D`, with the newest tab first. The tab has a week summary at the top, then one block per outlet (`■ Outlet` header, item rows, `Daily Waste Cost ($)`, `Closing Time`, `Counted By (Badge)`). Re-sending a night overwrites that day's column. New menu items are inserted above the block's cost row. The **weekly report email** is sent by `sendWeeklyReport_` when a Sunday count arrives. If Sunday is re-sent, an "UPDATED" copy goes out. If Sunday is never submitted, the report goes out "late" with the first count of the next week. Sent status is kept in cell A3 of the week tab. To send the current week's report by hand, run `sendWeeklyReportNow()` from the editor.
- Payload shape: `{id, key, chef, badge, date, submittedAt, totals:{itemsWasted,units,cost}, stations:[{station, stationLabel, closeTime, notes, totals, items:[{category,item,uom,qty,costPerUnit,cost,targetPct}]}]}`. If you change it, update `Code.gs` too.

## Working on it
- Preview the app by opening `index.html` in a browser; there is no server or build.
- Test with `python tools/e2e_test.py`, which needs `pip install playwright` and `playwright install chromium`.
- Preview the email with `node tools/render_email.js <payload.json> out.html`.
- After changing `Code.gs`, the manager must paste it into Apps Script and use **Deploy ▸ Manage deployments ▸ New version**. The URL doesn't change.
- Keep the navy/gold look (`--bg #0c1629`, `--gold #d4af37`), large touch targets (at least 44px), and phone widths down to 360px with no horizontal scroll.
