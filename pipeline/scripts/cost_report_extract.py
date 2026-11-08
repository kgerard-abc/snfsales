"""Copy one home's rows out of the CMS cost report files, unchanged, one set of files per report.

Usage: python3 pipeline/scripts/cost_report_extract.py 415035 [SNF10 SNF24]

The optional list of form prefixes says which CMS files to read (default SNF10, form CMS-2540-10; SNF24 is the newer
form CMS-2540-24, used for cost reports beginning in 2025).

CMS's HCRIS files (SNF10FY<year>.ZIP) hold three CSV files for every nursing home in the country: _rpt (one row per
report), _nmrc (every number on every worksheet), and _alpha (every piece of text). The _nmrc file alone is about
500 MB, too big for Excel. This copies only the rows for the given CMS number's reports, exactly as they appear in
pipeline/hcris (which holds CMS's rows unchanged, Rhode Island only), into
pipeline/cost_report_views/<CCN>_cms_rows/<period end>_<record number>_{rpt,nmrc,alpha}.csv.
The only thing added is a header row with CMS's own column names; CMS's files have no header row.
"""
import csv, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
HC = ROOT / "hcris"
ccn = sys.argv[1]
FORMS = sys.argv[2:] or ["SNF10"]
OUT = ROOT / "cost_report_views" / f"{ccn}_cms_rows"
OUT.mkdir(parents=True, exist_ok=True)
HEAD = {
    "rpt": ["RPT_REC_NUM", "PRVDR_CTRL_TYPE_CD", "PRVDR_NUM", "NPI", "RPT_STUS_CD", "FY_BGN_DT", "FY_END_DT", "PROC_DT",
            "INITL_RPT_SW", "LAST_RPT_SW", "TRNSMTL_NUM", "FI_NUM", "ADR_VNDR_CD", "FI_CREAT_DT", "UTIL_CD", "NPR_DT",
            "SPEC_IND", "FI_RCPT_DT"],
    "nmrc": ["RPT_REC_NUM", "WKSHT_CD", "LINE_NUM", "CLMN_NUM", "ITM_VAL_NUM"],
    "alpha": ["RPT_REC_NUM", "WKSHT_CD", "LINE_NUM", "CLMN_NUM", "ALPHNMRC_ITM_TXT"],
}
for rpt in sorted(p for f in FORMS for p in HC.glob(f"{f}FY*_rpt.csv")):
    name = rpt.name.replace("_rpt.csv", "")
    for r in csv.reader(open(rpt, encoding="latin-1")):
        if r[2].strip() != ccn: continue
        rec, end = r[0], r[6]
        stem = f"{end[6:]}-{end[:2]}-{end[3:5]}_{rec}"
        with open(OUT / f"{stem}_rpt.csv", "w", newline="") as o:
            w = csv.writer(o); w.writerow(HEAD["rpt"]); w.writerow(r)
        for kind in ("nmrc", "alpha"):
            with open(HC / f"{name}_{kind}.csv", encoding="latin-1") as f, open(OUT / f"{stem}_{kind}.csv", "w", newline="") as o:
                w = csv.writer(o); w.writerow(HEAD[kind])
                n = 0
                for row in csv.reader(f):
                    if row and row[0] == rec: w.writerow(row); n += 1
        print(name, rec, end, n)
