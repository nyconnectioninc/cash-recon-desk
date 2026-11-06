# Cash Recon Desk

The dashboard for the daily **RQ cashout vs. Brinks deposit** recon.

- **Frontend:** `index.html`, served by GitHub Pages. It holds no data. Everything loads from the API after sign-in.
- **Backend:** the **Brinks Cash Recon Engine** Apps Script project (owner `jonsanchez0009@gmail.com`).
  - `apps-script/Code.gs`: the engine. It runs every 15 min from 6:00 to 11:59 AM, ingests the RQ file and the Brinks emails (**Device Activity**: every safe type, plus the Deposit Report for comments/variance), reconciles, nets late drops and sends the alert.
  - `apps-script/Api.gs` + `apps-script/Api2.gs` (*Features.gs* in the project): the dashboard API (web app). Deployed as **Execute as Me · Anyone**. Every call except sign-in needs a signed 30-day session for an Active user on the **Users** tab.
- **Data:** Google Sheet *Brinks Deposit Recon – Control Sheet* (tabs: Store Map, Daily_Log, RQ_Raw, Brinks_Raw, Config, Users, Desk_Activity, Deposit_Allocations, Day_Locks, Access_Requests, Run_Log).

## Brinks data

- **Device Activity report** (main source): bag drops *and* smart-safe (bill feeder) deposits, with the Brinks **Location #** on every row. Location # = Store Map **Brinks ID**, so stores that share a Brinks Store ID (Berlin, Broadway, 8th Ave) are told apart exactly.
- **Deposit Report** (kept running): adds comments, verified $ and variance to the same bags (merged by barcode).
- Smart-safe deposits have no barcode; they're keyed `SD-<Location #>-<time>-<amount>`. $1 install tests are ignored.
- **New installs:** the first deposit under a `WZ####` Store ID fills that store's Brinks ID on the Store Map, sets it Active and emails the admin.
- **Manual exports:** drop a Device Activity / Deposit Report export (xlsx or Google Sheet) in Drive › Brinks Recon › Brinks Imports. It loads on the next run. `brinksCatchUp()` loads it now and re-runs the last 7 days.
- Config: `DEVICE_QUERY` (Gmail search for the Device Activity email), `BRINKS_QUERY`, `BRINKS_IMPORT_PATH`.

## Roles (Users tab)

| Role | Sees | Can edit |
|---|---|---|
| Admin | All stores | Adjustments, notes, team access |
| Editor | All stores | Adjustments, notes |
| Super Viewer | All stores | — |
| DM Viewer | Stores where Store Map **DM Email** = their email (or DM names listed in Scope) | — |
| Store Viewer | Stores where Store Map **SM Email** = their email (or WZ #s listed in Scope) | — |

Unchecking **Active** locks a person out immediately.

## What the page does

- **Daily recon:** pick a business day, see the counts and $ for each status, a 21-day trend, and a filterable store table. Editors can **close the day** (locks edits); only an Admin can reopen it.
- **Store detail:** RQ cashouts and the Brinks bags counted for the day, engine notes, 30-day history, the adjustment form, plus:
  - **Extra $ matcher:** when a day is over, it finds the earlier short or missing days that add up to the extra and applies them in one tap.
  - **Combined deposit:** one drop covering several days, with a required reason. Store closed, Brinks missed pickup, and safe malfunction are excused on the scorecard; manager error counts as late. Every allocation can be undone.
- **Still open:** unresolved items from the last 30 days.
- **Safe tracker:** cash that should be in each safe since its last Brinks pickup, compared with RQ cashouts, with a limit warning. Pickup days and per-store limits live on the Store Map (*Pickup Days*, *Safe Limit*).
- **Scorecard:** missing, short/over, late, and excused counts by store or DM over 5, 10, or 20 business days.
- **Store calendar:** month view per store, color-coded by status.
- **Stores:** every store and its history.
- **Export:** Excel, Google Sheet (saved to Drive › Brinks Recon › Exports), or PDF for any date range and scope.
- **People & access (Admin):** access requests (people ask from the sign-in page; @wireless-zone.com / @nyconnectionli.com only), approve or deny, roles and scope, a permissions table, and settings (tolerance, safe limit).

Adjustments write to Daily_Log (*Manual Adj $, Adj Reason, Adj By, Note*). The engine never overwrites them. Every change is logged on **Desk_Activity**.

## Updating

- **Page:** edit `index.src.html`, run `python3 build.py` (it inlines the logo into `index.html`), and push. Pages updates in about a minute.
- **API / engine:** paste into Apps Script, then **Deploy → Manage deployments → Edit → Version: New version → Deploy**. The URL stays the same.
