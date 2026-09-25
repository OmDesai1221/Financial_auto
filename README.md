# Astra Chemtech — Monthly Excel Update Tool

A small local web app that automates merging each month's GL trial-balance
export ("Sheet1") into two existing Excel workbooks:

- **PNL** — `P&L STATEMENT.xlsx`, a single-sheet Profit & Loss statement.
- **ACPLACT** — `acplact.xlsx`, a two-sheet FIFO costing schedule
  (`FIFO-sch` detail + `FIFO - P & L` summary).

Both are matched by **GL Code** (set once, permanent — never re-derived from
description text). You can run either transformation, or both together, from
one monthly upload.

## What it does each month

1. You upload the new month's Sheet1 export.
2. It matches every GL-coded row by code, inserts any brand-new GL rows in
   the correct position, adds the new month as a column (shifting older
   months over), and carries forward safe formula patterns (subtotals,
   fiscal-YTD accruals, the FIFO stock rollforward, etc.) for rows that have
   no Sheet1 source at all.
3. Before anything is saved, you get a **review page** — every affected row
   is listed with where its value came from, and anything with no reliable
   source is left for you to type in by hand (not guessed at).
4. Only when you click **Confirm** does it write to the real files — a
   timestamped backup of the previous state is kept automatically, and a
   one-click **Undo** reverts the most recent change.

It also supports **"Replace latest month"** for when corrected data arrives
after a month's already been processed, without adding a duplicate column.

## Requirements

- Python 3.9+
- Your own copies of `P&L STATEMENT.xlsx` and `acplact.xlsx` (these are
  **not** included in this repo — see below)

## Setup

```bash
git clone https://github.com/OmDesai1221/<repo-name>.git
cd <repo-name>

python3 -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

pip install -r requirements.txt
```

## Put your real files in place

This repo ships **code only** — no financial data. Before running it, copy
your own workbooks into the project root (same folder as this README):

```
<repo-name>/
├── P&L STATEMENT.xlsx      ← your real P&L workbook
├── acplact.xlsx             ← your real ACPLACT workbook
└── app/
```

These two files are gitignored, so they'll never accidentally get committed
or overwritten by a `git pull`.

## Running it

```bash
cd app
python3 app.py
```

Open **http://localhost:5050** in a browser. You'll be asked to sign in first:

- **Username:** `kishor`
- **Password:** `Shree@1977` (default — changeable from within the app via "Change password", which requires the current password to confirm)

The password is stored locally (hashed) in `app/auth_data.json`, gitignored — a changed password persists across restarts but is specific to whichever machine it was changed on.

- Check the boxes for whichever transformation(s) you want to run.
- Choose **New month** (adds a new column) or **Replace latest month**
  (corrects the month you already processed).
- Upload the new Sheet1 file and submit.
- Review the changes, then **Confirm & save** (or **Cancel** to discard).

Leave the terminal running while you use the app; closing it (Ctrl+C) stops
the server.

## Letting a client use it remotely (no domain needed)

Since this is meant for a single person to use once a month, the simplest
zero-cost way to give them a link is a quick tunnel — no account, no config:

```bash
brew install cloudflared      # one-time
cloudflared tunnel --url http://localhost:5050
```

This prints a temporary public URL like `https://random-words.trycloudflare.com`.
Send that link to whoever needs it; it only works while both the app and the
tunnel are running on your machine, and a fresh link is generated every time
you restart the tunnel.

## Project structure

```
app/
├── app.py                Flask routes — upload, review, confirm, undo
├── transform_type1.py    PNL transformation logic
├── transform_type2.py    ACPLACT transformation logic
├── templates/            HTML pages (index, review, result)
├── uploads/               each Sheet1 file you upload (gitignored)
├── outputs/               finished files ready to download (gitignored)
├── history/               automatic pre-change backups, for Undo (gitignored)
└── pending/                in-progress merge awaiting your Confirm (gitignored)
```

## Notes

- GL Code is the permanent match key — once a row is created from a
  particular code, its description never changes again even if a future
  Sheet1 spells it differently.
- "Undo" reverts every file changed in the most recent run together (if you
  ran both transformations at once, Undo restores both; if you ran only one,
  it only touches that one).
- The tool needs a login (see above). Only share the tunnel link with the intended user, and don't leave the tunnel running longer than needed.
- The first page has a **Download the file from last time** section for re-fetching the most recent finished PNL / ACPLACT files.
- ACPLACT: opening stock (`X1…` codes) is carried forward unchanged; closing stock (`R3…`) follows the sign used the previous month; Sheet1 codes not linked to any row are listed on the review page with their amounts.
