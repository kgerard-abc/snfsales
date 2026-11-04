#!/usr/bin/env python3
"""Build the sales page and the workbook from the history files and the master Google Sheet.

Usage: build_site.py <master-sheet-id>
Writes: index.html (repo root, replaces the old page), pipeline/out/RI_Nursing_Homes_CMS_and_CEC.xlsx
The logo already in the repo is left alone.
"""
import gzip, io, os, shutil, sys, urllib.request
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(HERE, "scripts"))
import ri_nh_build as B

sheet_id = sys.argv[1]
if "PUT_" in sheet_id:
    print("ERROR: the master Google Sheet id has not been set in the workflow yet."); sys.exit(1)
out = os.path.join(HERE, "out"); std = os.path.join(out, "std")
shutil.rmtree(out, ignore_errors=True); os.makedirs(std)
for n in ("providerinfo", "healthcitations", "penalties", "ownership"):
    with gzip.open(os.path.join(HERE, "history", n + ".csv.gz"), "rb") as f, open(os.path.join(std, n + ".csv"), "wb") as g:
        shutil.copyfileobj(f, g)
url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
try:
    data = urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "abc-ri-update/1.0"}), timeout=120).read()
except Exception as e:
    print(f"ERROR: could not download the master Google Sheet ({e}). Is it shared as 'Anyone with the link can view'?"); sys.exit(1)
if data[:2] != b"PK":
    print("ERROR: the Google Sheet link did not return a spreadsheet file. Check the sharing setting."); sys.exit(1)
master = os.path.join(out, "RI_Nursing_Home_CECs_master.xlsx")
open(master, "wb").write(data)
# RIDOH's own list of reviews is the source for where each pending application stands. Best effort: if it cannot be fetched, the CEC Log text is used.
RIDOH_SHEET_ID = "1fdRP2rh34eluTT13k8EIqjJUSnGTH2U5yGD-KwRFi-M"
try:
    rd = urllib.request.urlopen(urllib.request.Request(f"https://docs.google.com/spreadsheets/d/{RIDOH_SHEET_ID}/export?format=xlsx", headers={"User-Agent": "abc-ri-update/1.0"}), timeout=120).read()
    if rd[:2] == b"PK":
        B.RIDOH_PATH = os.path.join(out, "RIDOH_reviews.xlsx"); open(B.RIDOH_PATH, "wb").write(rd)
    else: print("WARNING: the RIDOH list did not download as a spreadsheet; pending status comes from the CEC Log.")
except Exception as e:
    print(f"WARNING: could not download the RIDOH list of reviews ({e}); pending status comes from the CEC Log.")
import openpyxl
names = openpyxl.load_workbook(master, read_only=True).sheetnames
need = ["CEC Log", "Event Overrides", "CCN Map", "Management Agreements"]
if [n for n in need if n not in names]:
    print(f"ERROR: the Google Sheet is missing tabs: {[n for n in need if n not in names]}. Is it the INTERNAL copy?"); sys.exit(1)
r = B.run(out, cec_path=master, std_dir=std)
shutil.copy(os.path.join(out, "public", "index.html"), os.path.join(ROOT, "index.html"))
print("Built index.html and the workbook.")
