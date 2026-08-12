#!/usr/bin/env python3
"""
aggregate_workbook.py

Reads a GoFresh sales workbook and prints the compact
{ lookups, rows, n_rows, n_invoices } JSON payload to stdout.

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

        rows.append([
            idx(dates, d),
            r[col["Hour"]] if r[col["Hour"]] not in (None, "") else -1,
            r[col["Weekday"]] if r[col["Weekday"]] not in (None, "") else -1,
            idx(shops, r[col["ShopName"]]),
            idx(regs, r[col["Region"]]),
            idx(prods, r[col["ProductName"]]),
            idx(segs, r[col["Segment"]]),
            idx(cats, r[col["Category"]]),
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