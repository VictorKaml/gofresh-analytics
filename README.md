# GoFresh Sales Dashboard — Desktop App

An Electron wrapper around your existing `GoFresh_Sales_Dashboard.html` /
`generate_dashboard.py` dashboard. Excel parsing is still done by Python
(`src/aggregate_workbook.py`, using `openpyxl`) — the same `aggregate()`
logic as `generate_dashboard.py` — but it now runs as a subprocess launched
by the Electron main process, rather than inside the dashboard's own code.
**Requires Python 3 with `openpyxl` installed** on the machine running the
app (`pip install openpyxl`).

## What it does

- On launch, it automatically loads the bundled workbook at `data/default.xlsx`
  and renders the full dashboard (all tabs, charts, tables — identical to the
  HTML your Python script generates).
- **File → Open Excel File…** (or the toolbar button, or `Ctrl+O`) lets you
  point it at any other `.xlsx`/`.xlsm` workbook that has a `data all` sheet
  with the expected columns. The dashboard re-renders instantly.

## Project structure

```
gofresh-dashboard-app/
├── main.js                 Electron main process (window, menu, file dialog)
├── preload.js               Secure bridge between main process and UI
├── index.html                App shell: toolbar + iframe that hosts the dashboard
├── src/aggregate_workbook.py  Python subprocess: same aggregate() logic as generate_dashboard.py
├── assets/dashboard-template.html   Your dashboard HTML, with __PAYLOAD__ placeholder
├── data/default.xlsx         Bundled default workbook (swap this file to change the default)
└── package.json               Dependencies + electron-builder config (Windows/nsis)
```

## How it works

1. `main.js` shells out to `python3 src/aggregate_workbook.py <path>` (via
   Node's `child_process.execFileSync`), which reads the chosen `.xlsx` with
   `openpyxl`, finds the `data all` sheet, auto-detects the header row, and
   prints the compact `{ lookups, rows, n_rows, n_invoices }` payload as JSON
   on stdout — the exact same logic as `aggregate()` in
   `generate_dashboard.py`. Progress/diagnostic messages go to stderr so
   stdout stays pure JSON. `main.js` parses that JSON back into an object.
2. `main.js` takes that payload, does `JSON.stringify()`, and substitutes it
   into `assets/dashboard-template.html` in place of `__PAYLOAD__` — this is
   the exact same template your Python `write_html()` step uses. All the
   chart/table aggregation logic lives in that template's embedded
   JavaScript, untouched.
3. The rendered HTML is written to a temp file in the app's user-data folder
   and loaded into an `<iframe>` in the app window.

Because all the heavy-lifting front-end logic (KPIs, charts, filters,
sortable tables) is unchanged from your original file, the app's behavior is
identical to opening the generated `.html` in a browser — just wrapped in a
proper desktop window with a file picker.

## Requirements

- [Node.js](https://nodejs.org) 18+ (includes npm) — only needed on the
  *build* machine, not for people running the packaged app.
- **Python 3 with `openpyxl` installed** — needed on *every* machine that
  runs the app (dev or packaged), since `aggregate_workbook.py` runs as a
  subprocess each time a workbook is loaded. On Windows the app looks for
  `python` then `py` on PATH; on macOS/Linux it looks for `python3` then
  `python`. If neither is found, the app shows an error asking the user to
  install Python 3 and run `pip install openpyxl`.

## Setup (first time)

```bash
cd gofresh-dashboard-app
npm install
pip install openpyxl
```

## Run in development

```bash
npm start
```

This launches the app via Electron directly from source — good for quick
testing/tweaks.

## Build a Windows installer

```bash
npm run dist
```

This uses `electron-builder` to produce a Windows installer (`.exe`, NSIS)
in the `dist/` folder. Double-click it to install the app like any other
Windows program (Start Menu + Desktop shortcuts are created automatically).

> The first `npm run dist` run downloads Electron's prebuilt binaries and
> build tools — it needs an internet connection and can take a few minutes.

## Swapping the default workbook

Replace `data/default.xlsx` with your latest export (keep the filename, or
update `DEFAULT_XLSX` in `main.js`), then rebuild. Users can also always use
**File → Open Excel File…** to point at any workbook without rebuilding.

## Customizing the dashboard itself

Any visual/behavioral changes should be made in
`assets/dashboard-template.html` (same file your Python script embeds) —
`main.js` just fills in the data, it doesn't touch layout or chart logic.

## Notes on the required workbook format

Same as the original script: a sheet named `data all` (any capitalization)
with at least these columns somewhere in the header row (row 1 can be a
summary row, it's skipped):

```
Date, Hour, Weekday, ShopName, Region, ProductName, Segment, Category, Qty, Amount
```

Optional columns used if present: `InvoiceDate` (date fallback), `Volume`,
`InvoiceNo`.
