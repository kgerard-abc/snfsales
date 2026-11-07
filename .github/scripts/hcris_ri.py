"""Download CMS cost report files (HCRIS) and keep only Rhode Island nursing homes.

Usage: python3 hcris_ri.py SNF24FY2025 SNF10FY2025 ...
Saves pipeline/hcris/<name>_{rpt,nmrc,alpha}.csv with only the RI reports (provider numbers starting with 41).
The 2540-10 form files are named SNF10FY<year>; the new 2540-24 form (periods ending on or after Sept. 30, 2025) is SNF24FY<year>.
"""
import csv, io, sys, urllib.request, zipfile
from pathlib import Path

OUT = Path("pipeline/hcris")
OUT.mkdir(parents=True, exist_ok=True)
for name in sys.argv[1:]:
    url = f"https://downloads.cms.gov/FILES/HCRIS/{name}.ZIP"
    print("downloading", url, flush=True)
    req = urllib.request.Request(url, headers={"User-Agent": "abc-ri-data/1.0"})
    data = urllib.request.urlopen(req, timeout=600).read()
    z = zipfile.ZipFile(io.BytesIO(data))
    print(" ", [(i.filename, i.file_size) for i in z.infolist()])
    files = {}
    for i in z.infolist():
        low = i.filename.lower()
        for kind in ("rpt", "nmrc", "alpha"):
            if low.endswith(f"_{kind}.csv"): files[kind] = i.filename
    keep = set()
    with z.open(files["rpt"]) as f, open(OUT / f"{name}_rpt.csv", "w", newline="") as o:
        w = csv.writer(o)
        for r in csv.reader(io.TextIOWrapper(f, "latin-1")):
            if len(r) > 2 and r[2].strip().startswith("41"):
                w.writerow(r); keep.add(r[0])
    print(f"  {len(keep)} Rhode Island reports")
    for kind in ("nmrc", "alpha"):
        n = 0
        with z.open(files[kind]) as f, open(OUT / f"{name}_{kind}.csv", "w", newline="") as o:
            w = csv.writer(o)
            for r in csv.reader(io.TextIOWrapper(f, "latin-1")):
                if r and r[0] in keep: w.writerow(r); n += 1
        print(f"  {kind}: {n} rows")
