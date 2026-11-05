# Cash Recon Desk

The dashboard for the daily **RQ cashout vs. Brinks deposit** recon.

- **Frontend:** `index.html`, served by GitHub Pages. It holds no data. Everything loads from the API after sign-in.
- **Backend:** the **Brinks Cash Recon Engine** Apps Script project (owner `jonsanchez0009@gmail.com`).
  - `apps-script/Code.gs`: the engine. It runs every 15 min from 6:00 to 11:59 AM, ingests the RQ file and the Brinks email, reconciles, nets late drops and sends the alert.
  - `apps-script/Api.gs`: the dashboard API (web app). Deployed as **Execute as Me · Anyone**. Every call except sign-in needs a signed 30-day session for an Active user on the **Users** tab.
- **Data:** Google Sheet *Brinks Deposit Recon – Control Sheet* (tabs: Store Map, Daily_Log, RQ_Raw, Brinks_Raw, Config, Users, Desk_Activity, Run_Log).

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

- **Daily recon:** pick a business day, see reconciled / missing / short / over / bag-no-cashout counts and $, a 21-day trend, and a filterable store table.
- **Store detail:** RQ cashouts, the Brinks bags counted for that day (who made each one, its barcode, Brinks ID and comment), engine notes, a 30-day history, and (for editors) an adjustment and note form.
- **Still open:** unresolved items from the last 30 days.
- **Stores:** every store and its day-by-day history.
- **Team access:** add people, set roles and scope (Admins).

Adjustments write to Daily_Log (*Manual Adj $, Adj Reason, Adj By, Note*). The engine never overwrites them. Every change is logged on **Desk_Activity**.

## Updating

- **Page:** edit `index.src.html`, run `python3 build.py` (it inlines the logo into `index.html`), and push. Pages updates in about a minute.
- **API / engine:** paste into Apps Script, then **Deploy → Manage deployments → Edit → Version: New version → Deploy**. The URL stays the same.
