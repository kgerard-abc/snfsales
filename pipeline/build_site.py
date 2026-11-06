#!/usr/bin/env python3
"""Build the sales page and the workbook from the history files and the master Google Sheet.

Usage: build_site.py <master-sheet-id>
Writes: index.html (repo root, replaces the old page), pipeline/out/RI_Nursing_Homes_CMS_and_CEC.xlsx
The logo already in the repo is left alone.
"""
import gzip, io, json, os, shutil, sys, urllib.request
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
# The SNF dashboard's Facility extras sheet: the home names and owner group names both sites use. Best effort.
EXTRAS_SHEET_ID = "1PXFDq_55whZPcW0ODuJYsD3NGcJnsvgzYkgV0y8maGQ"
try:
    ex = urllib.request.urlopen(urllib.request.Request(f"https://docs.google.com/spreadsheets/d/{EXTRAS_SHEET_ID}/export?format=xlsx", headers={"User-Agent": "abc-ri-update/1.0"}), timeout=120).read()
    if ex[:2] == b"PK":
        B.EXTRAS_PATH = os.path.join(out, "facility_extras.xlsx"); open(B.EXTRAS_PATH, "wb").write(ex)
    else: print("WARNING: the Facility extras sheet did not download as a spreadsheet; CMS names are used.")
except Exception as e:
    print(f"WARNING: could not download the Facility extras sheet ({e}); CMS names are used.")
# CMS's national figures for each nursing home chain. Best effort.
import chains
CHAINS_CSV = os.path.join(ROOT, "pipeline", "cms_chains.csv")
try:
    B.CHAINS = chains.fetch()
    B.CHAINS_ASOF = chains.version()
    chains.save(B.CHAINS, CHAINS_CSV, B.CHAINS_ASOF)
    print(f"CMS chain file: {len(B.CHAINS)} chains, as of {B.CHAINS_ASOF}.")
except Exception as e:
    B.CHAINS, B.CHAINS_ASOF = chains.load(CHAINS_CSV)
    print(f"WARNING: could not download CMS's chain performance file ({e}); using the saved copy ({len(B.CHAINS or [])} chains).")
import openpyxl
names = openpyxl.load_workbook(master, read_only=True).sheetnames
need = ["CEC Log", "Event Overrides", "CCN Map", "Management Agreements"]
if [n for n in need if n not in names]:
    print(f"ERROR: the Google Sheet is missing tabs: {[n for n in need if n not in names]}. Is it the INTERNAL copy?"); sys.exit(1)
r = B.run(out, cec_path=master, std_dir=std)
shutil.copy(os.path.join(out, "public", "index.html"), os.path.join(ROOT, "index.html"))
# printable one-page summaries for the Health Services Council, one per pending buyer group
cdir = os.path.join(ROOT, "council")
shutil.rmtree(cdir, ignore_errors=True)
if os.path.isdir(os.path.join(out, "public", "council")):
    shutil.copytree(os.path.join(out, "public", "council"), cdir)
# sales.json: a small file the SNF dashboard reads to show each home's pending or most recent sale and link here.
# One entry per home (by CMS number): pending applications and every completed sale, newest first, with buyer names
# matching the dashboard's owner groups. The dashboard shows the newest sale and uses the rest to find when the
# current owner took over. Chain reorganizations are left out.
import json
D = r["D"]; homes = {}
for s_ in sorted(D["sales"], key=lambda x: x.get("date") or "", reverse=True):
    if s_.get("kind") != "sale" or not s_.get("fid") or str(s_["fid"]).startswith("cec:") or not s_.get("date"):
        continue
    pend = s_.get("outcome") == "Pending"
    homes.setdefault(s_["fid"], []).append(dict(date=s_["date"], date_note=s_.get("date_note"), buyer=s_.get("buyer") or s_.get("who") or None, pending=pend, mgmt_start=s_.get("mgmt_start")))
feed = dict(built=D["meta"].get("built"), source="https://snfsales.abc-ri.org/", homes=homes)
with open(os.path.join(ROOT, "sales.json"), "w", encoding="utf-8") as f:
    json.dump(feed, f, ensure_ascii=False, separators=(",", ":"))
print(f"Wrote sales.json ({len(homes)} homes).")
print("Built index.html and the workbook.")
