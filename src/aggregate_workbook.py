#!/usr/bin/env python3
"""
__VERSION__ = "3.0.0"
aggregate_workbook.py

Reads a GoFresh sales workbook and prints the compact
{ lookups, bva, rows, n_rows, n_invoices } JSON payload to stdout.

If a "budgetpershop" sheet is present, its Budget vs Actual data is read
and included under the "bva" key (matching generate_dashboard.py's BVA
tab); if the sheet is absent, "bva" is still emitted with empty
entries/detail so the dashboard's BVA tab can show "no budget loaded".

This is the same aggregate()/build_payload() logic as
generate_dashboard.py, factored out so the Electron app can shell out to
it (via main.js) instead of using the JS/SheetJS port in parseWorkbook.js.

Uses python-calamine (a Rust-backed reader) rather than openpyxl. openpyxl
parses cells one at a time in pure Python, which measured ~45k cells/sec on
a 700k-row / 12-column workbook (~183s total). calamine reads the same file
in ~13-14s end-to-end — about 13x faster — with identical output, since it
still hands back plain Python values (str/float/datetime.date/None) per
cell, so the aggregation logic below is unchanged.

IMPORTANT: stdout must contain ONLY the JSON payload — main.js parses it
directly. All progress/diagnostic messages go to stderr instead.

Usage:
    python aggregate_workbook.py <path-to-workbook.xlsx>

Exit codes:
    0  success — JSON payload printed to stdout
    1  a handled error occurred (missing sheet, missing columns, no rows,
       etc.) — a one-line human-readable message is printed to stderr
"""

import sys
import json
import itertools

REQUIRED_COLUMNS = [
    "Date", "Hour", "Weekday", "ShopName", "Region",
    "ProductName", "Segment", "Category", "Qty", "Amount",
]


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def progress(percent, status):
    """Emit a machine-readable progress line on stderr.

    percent is an int 0-100, or None for "can't estimate yet" (indeterminate
    — main.js shows a spinner + the status text instead of a percentage).
    """
    p = str(percent) if percent is not None else "-"
    print(f"PROGRESS:{p}:{status}", file=sys.stderr, flush=True)


def fail(msg):
    log(f"ERROR: {msg}")
    sys.exit(1)


def numify(v):
    if v is None or v == "":
        return 0.0
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0.0


def idx(store, val):
    """Intern a value into a dict-as-ordered-set; return its index."""
    if val is None or val == "":
        return -1
    if val not in store:
        store[val] = len(store)
    return store[val]


def aggregate(in_path):
    try:
        import python_calamine
    except ImportError:
        fail(
            "The 'python-calamine' package is required but not installed. "
            "Run: pip install python-calamine"
        )

    log(f"Reading {in_path} ...")
    progress(2, "Opening workbook")
    try:
        wb = python_calamine.CalamineWorkbook.from_path(in_path)
    except FileNotFoundError:
        fail(f"File not found: {in_path}")
    except OSError as e:
        fail(f"File not found: {in_path} ({e})")
    except python_calamine.CalamineError as e:
        fail(f"Could not open workbook: {e}")
    except Exception as e:
        fail(f"Could not open workbook: {e}")

    SHEET_NAME = "data all"
    match = next((n for n in wb.sheet_names if n.strip().lower() == SHEET_NAME), None)
    if not match:
        fail(f"Sheet 'data all' not found. Sheets present: {wb.sheet_names}")
    ws = wb.get_sheet_by_name(match)
    log(f"  Reading from sheet: '{match}'")
    progress(8, "Found data sheet")

    LOOK_FOR = {"Date", "Amount", "ShopName", "Segment"}
    header_row = None
    headers = None
    for row_num, row in enumerate(itertools.islice(ws.iter_rows(), 5), start=1):
        found = sum(1 for v in row if v in LOOK_FOR)
        if found >= 3:
            header_row = row_num
            headers = row
            break
    if header_row is None:
        fail(
            "Could not find a header row (looked for Date / Amount / "
            "ShopName / Segment in the first 5 rows)."
        )
    log(f"  Header row detected: row {header_row}")
    progress(12, "Detected header row")

    SAFETY_MAX_ROW = 1_100_000
    # header_row is 1-indexed; skip that many rows (0-indexed) of a fresh
    # iterator to land right after the header, same as openpyxl's
    # min_row=header_row + 1 did.
    rows_iter = itertools.islice(ws.iter_rows(), header_row, SAFETY_MAX_ROW)
    col = {h: i for i, h in enumerate(headers) if h}

    missing = [r for r in REQUIRED_COLUMNS if r not in col]
    if missing:
        fail(f"Required columns missing from 'data all': {missing}")

    has_invoice_date = "InvoiceDate" in col
    has_volume = "Volume" in col
    has_invoice_no = "InvoiceNo" in col

    # Rough total for percentage math — calamine reports the sheet's exact
    # dimensions up front (unlike openpyxl, which can leave max_row
    # unavailable for some files), so this is normally exact rather than an
    # estimate. Fall back to an indeterminate spinner if unavailable.
    try:
        total_est = max((ws.height or 0) - header_row, 0)
    except Exception:
        total_est = 0

    shops, regs, prods, segs, cats, dates, invs = {}, {}, {}, {}, {}, {}, {}
    prod_cat = {}   # prodIdx -> catIdx (first seen), used to map budget lines to categories
    rows = []
    scanned = 0
    last_reported_pct = None

    for r in rows_iter:
        scanned += 1
        # calamine represents a blank cell as "" (not None, unlike
        # openpyxl), including a fully-blank row coming back as a
        # full-width row of "" rather than None/[] — check both.
        if r is None or r[0] is None or r[0] == "":
            continue

        date = r[col["Date"]]
        if not date and has_invoice_date:
            date = r[col["InvoiceDate"]]
        if not date or not hasattr(date, "strftime"):
            continue  # blank or non-date (text) cell — skip

        d = date.strftime("%Y-%m-%d")

        p_idx = idx(prods, r[col["ProductName"]])
        c_idx = idx(cats, r[col["Category"]])
        if p_idx >= 0 and c_idx >= 0 and p_idx not in prod_cat:
            prod_cat[p_idx] = c_idx

        rows.append([
            idx(dates, d),
            r[col["Hour"]] if r[col["Hour"]] not in (None, "") else -1,
            r[col["Weekday"]] if r[col["Weekday"]] not in (None, "") else -1,
            idx(shops, r[col["ShopName"]]),
            idx(regs, r[col["Region"]]),
            p_idx,
            idx(segs, r[col["Segment"]]),
            c_idx,
            round(numify(r[col["Qty"]]), 2),
            round(numify(r[col["Amount"]]), 0),
            round(numify(r[col["Volume"]]) if has_volume else 0.0, 2),
            idx(invs, r[col["InvoiceNo"]]) if has_invoice_no else -1,
        ])

        if scanned % 500 == 0:
            if total_est > 0:
                frac = min(scanned / total_est, 1.0)
                pct = 15 + int(frac * 75)  # scanning spans 15%..90%
                if pct != last_reported_pct:
                    progress(pct, f"Scanning rows ({scanned:,} / {total_est:,})")
                    last_reported_pct = pct
            else:
                progress(None, f"Scanning rows ({scanned:,} so far)")

    log(f"  Scan complete. Scanned {scanned:,} rows, kept {len(rows):,}.")
    progress(90, f"Scanned {scanned:,} rows")

    if not rows:
        fail(f"No usable transaction rows found (scanned {scanned} rows). "
             f"Check that the Date column contains real dates.")

    sorted_dates = sorted(dates.keys())
    remap = {old_idx: new_idx for new_idx, d in enumerate(sorted_dates)
             for old_idx in [dates[d]]}
    for row in rows:
        row[0] = remap[row[0]]

    # ---- Budget vs Actual: read the budgetpershop sheet if present ----
    progress(92, "Reading budget sheet")
    bva_entries = {}
    bva_detail = []
    BVA_SEGS = ["Chicken", "Beef", "Egg", "Trading", "Oil"]
    if "budgetpershop" in wb.sheet_names:
        bws = wb.get_sheet_by_name("budgetpershop")
        prod_lookup = {name: i for name, i in prods.items()}
        shop_lookup = {name: i for name, i in shops.items()}
        cats_list = list(cats.keys())
        cat_norm = {}
        for name, i in cats.items():
            key = (name or "").strip().lower()
            cat_norm.setdefault(key, i)
        va_idx = cat_norm.get("value added", -1)
        n_budget = 0
        header_found = False
        for row in itertools.islice(bws.iter_rows(), 4000):
            if not header_found:
                if row and len(row) >= 2 and row[0] == "Shop" and row[1] == "Product":
                    header_found = True
                continue
            # calamine trims trailing blank cells, so rows may be shorter than
            # the sheet's full width. Pad with "" to match openpyxl behaviour.
            if not row or row[0] in (None, ""):
                continue
            if len(row) < 24:
                row = list(row) + [""] * (24 - len(row))
            shop_name, prod_name, seg = row[0], row[1], row[2]
            if seg not in BVA_SEGS:
                continue
            dailies = [numify(v) for v in row[11:16]]
            if not any(dailies):
                continue
            s_idx = shop_lookup.get(shop_name, -1)
            if s_idx < 0:
                continue
            if prod_name and "floor top-up" in str(prod_name).lower():
                c_idx = cat_norm.get("oil", -1)
                p_idx = idx(prods, "Oil — A-grade floor top-up")
            else:
                p_idx = prod_lookup.get(prod_name, -1)
                if p_idx < 0:
                    p_idx = idx(prods, prod_name)  # budget-only product (not yet sold)
                c_idx = prod_cat.get(p_idx, -1)
                cname = (cats_list[c_idx].strip().lower() if 0 <= c_idx < len(cats_list) else "")
                if cname == "value added" and va_idx >= 0:
                    c_idx = va_idx
            key = (s_idx, BVA_SEGS.index(seg), c_idx)
            acc = bva_entries.setdefault(key, [0.0] * 5)
            for j in range(5):
                acc[j] += dailies[j]
            try:
                price = float(row[23]) if row[23] not in (None, '', '#N/A') else 0.0
            except (TypeError, ValueError):
                price = 0.0
            bva_detail.append([s_idx, p_idx, BVA_SEGS.index(seg), c_idx]
                              + [round(v, 2) for v in dailies]
                              + [round(price, 2)])
            n_budget += 1
        log(f"  budgetpershop: {n_budget:,} budget lines read -> {len(bva_entries):,} shop x segment x category entries")
        if n_budget == 0:
            log("  WARNING: budgetpershop sheet exists but no values could be read.")
    else:
        log("  budgetpershop sheet not found - Budget vs Actual tab will show 'no budget loaded'")

    progress(96, "Building dashboard payload")

    return {
        "lookups": {
            "shops": list(shops.keys()),
            "regs": list(regs.keys()),
            "prods": list(prods.keys()),
            "segs": list(segs.keys()),
            "cats": list(cats.keys()),
            "dates": sorted_dates,
        },
        "bva": {
            "segs": BVA_SEGS,
            "units": {"Chicken": "kg", "Beef": "kg", "Egg": "packs", "Trading": "units", "Oil": "L"},
            "months": ["2026-08", "2026-09", "2026-10", "2026-11", "2026-12"],
            "month_labels": ["Aug 2026", "Sep 2026", "Oct 2026", "Nov 2026", "Dec 2026"],
            "days_in_month": [31, 30, 31, 30, 31],
            "entries": [[k[0], k[1], k[2]] + [round(v, 2) for v in vals]
                        for k, vals in sorted(bva_entries.items())],
            "detail": bva_detail,
        },
        "rows": rows,
        "n_rows": len(rows),
        "n_invoices": len(invs),
    }


def main():
    if len(sys.argv) != 2:
        fail("usage: aggregate_workbook.py <path-to-workbook.xlsx>")

    payload = aggregate(sys.argv[1])
    progress(99, "Finalizing")
    # Only the JSON payload goes to stdout — everything else is on stderr.
    sys.stdout.write(json.dumps(payload, default=str))
    sys.stdout.flush()


if __name__ == "__main__":
    main()