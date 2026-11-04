#!/usr/bin/env python3
"""Watch RIDOH's own list of reviews (Annual Report of Reviews) for changes to Change in Effective Control rows.

Usage: ridoh_watch.py [path-to-ridoh.xlsx]     (with no path, the sheet is downloaded)
Compares the CEC rows now with the saved copy in pipeline/ridoh_cec_snapshot.json.
Writes GitHub Actions outputs: changed=true|false, first=true|false. Prints which rows changed.
"""
import sys, os, re, json, datetime as dt, urllib.request
import openpyxl

SHEET_ID = "1fdRP2rh34eluTT13k8EIqjJUSnGTH2U5yGD-KwRFi-M"
SNAP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "pipeline", "ridoh_cec_snapshot.json")

def out(k, v):
    print(f"{k}={v}")
    if os.environ.get("GITHUB_OUTPUT"):
        with open(os.environ["GITHUB_OUTPUT"], "a") as f: f.write(f"{k}={v}\n")

def cell_text(c):
    v = c.value
    if v is None: return ""
    if isinstance(v, (dt.datetime, dt.date)): return f"{v:%Y-%m-%d}"
    return re.sub(r"\s+", " ", str(v)).strip()

def read_rows(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    this = dt.date.today().year; rows = {}
    for sh in (str(this), str(this - 1)):
        if sh not in wb.sheetnames: continue
        data = [list(r) for r in wb[sh].iter_rows()]
        i = next((k for k, r in enumerate(data) if r and str(r[0].value or "").strip().lower() == "change in effective control"), None)
        if i is None: continue
        k = i + 2
        while k < len(data) and data[k] and data[k][0].value not in (None, ""):
            cells = [cell_text(c) for c in data[k][:11]]
            key = f"{sh}|{cells[0]}"; n = 2
            while key in rows: key = f"{sh}|{cells[0]}|{n}"; n += 1
            rows[key] = cells; k += 1
    return rows

def main():
    if len(sys.argv) > 1: path = sys.argv[1]
    else:
        path = "ridoh.xlsx"
        try:
            data = urllib.request.urlopen(urllib.request.Request(f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx", headers={"User-Agent": "abc-ri-update/1.0"}), timeout=120).read()
        except Exception as e:
            print(f"::warning::Could not download RIDOH's list ({e}). Will try again next time."); out("changed", "false"); out("first", "false"); return
        if data[:2] != b"PK":
            print("::warning::RIDOH's list did not download as a spreadsheet. Will try again next time."); out("changed", "false"); out("first", "false"); return
        open(path, "wb").write(data)
    now = read_rows(path)
    if len(now) < 5:
        print(f"::warning::Only {len(now)} Change in Effective Control rows found in RIDOH's list; not trusting it. Nothing changed."); out("changed", "false"); out("first", "false"); return
    old = json.load(open(SNAP)) if os.path.exists(SNAP) else None
    if old is None:
        json.dump(now, open(SNAP, "w"), indent=0, sort_keys=True, ensure_ascii=False); print(f"First run: saved {len(now)} rows."); out("changed", "true"); out("first", "true"); return
    added = [k for k in now if k not in old]; gone = [k for k in old if k not in now]; chg = [k for k in now if k in old and now[k] != old[k]]
    for k in added: print("NEW ROW:", k)
    for k in gone: print("ROW GONE:", k)
    for k in chg:
        for j, (a, b) in enumerate(zip(old[k], now[k])):
            if a != b: print(f"CHANGED: {k} | column {j + 1}: {a[:70]!r} -> {b[:70]!r}")
    if added or gone or chg:
        json.dump(now, open(SNAP, "w"), indent=0, sort_keys=True, ensure_ascii=False); out("changed", "true")
    else: print(f"No change in {len(now)} rows."); out("changed", "false")
    out("first", "false")

main()
