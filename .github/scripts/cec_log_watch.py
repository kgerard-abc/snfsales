#!/usr/bin/env python3
"""Watch the CEC Log tab of the INTERNAL Google Sheet for new or edited rows.

Usage: cec_log_watch.py [path-to-internal.xlsx]     (with no path, the sheet is downloaded)
Compares columns A to L of the CEC Log tab now with the saved copy in pipeline/cec_log_snapshot.json.
Writes GitHub Actions outputs: cec_changed=true|false, cec_first=true|false. Prints which rows changed.
Used by the "Watch RIDOH's list of reviews" workflow so the sales page rebuilds after the weekly CEC Log check
adds a row, or after kg fills in a row's buyer group or principals.
"""
import sys, os, re, json, datetime as dt, urllib.request
import openpyxl

SHEET_ID = os.environ.get("MASTER_SHEET_ID", "1PAgtaNxnLerzBGVojlxgp-yB7N89msak6giJh4Sg6Wk")
SNAP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "pipeline", "cec_log_snapshot.json")
NCOLS = 12          # columns A to L: ID, facility, town, beds, applicant, buyer group, principals, source, type, outcome, date, year

def out(k, v):
    print(f"{k}={v}")
    if os.environ.get("GITHUB_OUTPUT"):
        with open(os.environ["GITHUB_OUTPUT"], "a") as f: f.write(f"{k}={v}\n")

def cell_text(v):
    if v is None: return ""
    if isinstance(v, (dt.datetime, dt.date)): return f"{v:%Y-%m-%d}"
    if isinstance(v, float) and v == int(v): v = int(v)
    return re.sub(r"\s+", " ", str(v)).strip()

def read_rows(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if "CEC Log" not in wb.sheetnames: return {}
    rows = {}
    for r in wb["CEC Log"].iter_rows(min_row=2, values_only=True):
        cells = [cell_text(v) for v in list(r)[:NCOLS]]
        cells += [""] * (NCOLS - len(cells))
        if not cells[0]: continue                       # no CEC ID: skip blank and stray rows
        key = cells[0]; n = 2
        while key in rows: key = f"{cells[0]}|{n}"; n += 1
        rows[key] = cells
    return rows

def main():
    if len(sys.argv) > 1: path = sys.argv[1]
    else:
        path = "internal.xlsx"
        try:
            data = urllib.request.urlopen(urllib.request.Request(f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx", headers={"User-Agent": "abc-ri-update/1.0"}), timeout=120).read()
        except Exception as e:
            print(f"::warning::Could not download the INTERNAL sheet ({e}). Will try again next time."); out("cec_changed", "false"); out("cec_first", "false"); return
        if data[:2] != b"PK":
            print("::warning::The INTERNAL sheet did not download as a spreadsheet (check it is shared as 'Anyone with the link can view'). Will try again next time."); out("cec_changed", "false"); out("cec_first", "false"); return
        open(path, "wb").write(data)
    now = read_rows(path)
    old = json.load(open(SNAP)) if os.path.exists(SNAP) else None
    if len(now) < 100:
        print(f"::warning::Only {len(now)} rows found on the CEC Log tab; not trusting it. Nothing changed."); out("cec_changed", "false"); out("cec_first", "false"); return
    if old is not None and len(now) < len(old) - 5:
        print(f"::warning::CEC Log dropped from {len(old)} to {len(now)} rows; not trusting it. Nothing changed."); out("cec_changed", "false"); out("cec_first", "false"); return
    if old is None:
        json.dump(now, open(SNAP, "w"), indent=0, sort_keys=True, ensure_ascii=False); print(f"First run: saved {len(now)} CEC Log rows."); out("cec_changed", "true"); out("cec_first", "true"); return
    added = [k for k in now if k not in old]; gone = [k for k in old if k not in now]; chg = [k for k in now if k in old and now[k] != old[k]]
    for k in added: print("NEW ROW:", k, "|", now[k][1][:50], "|", now[k][4][:60])
    for k in gone: print("ROW GONE:", k)
    for k in chg:
        for j, (a, b) in enumerate(zip(old[k], now[k])):
            if a != b: print(f"CHANGED: {k} | column {j + 1}: {a[:70]!r} -> {b[:70]!r}")
    if added or gone or chg:
        json.dump(now, open(SNAP, "w"), indent=0, sort_keys=True, ensure_ascii=False); out("cec_changed", "true")
    else: print(f"No change in {len(now)} CEC Log rows."); out("cec_changed", "false")
    out("cec_first", "false")

main()
