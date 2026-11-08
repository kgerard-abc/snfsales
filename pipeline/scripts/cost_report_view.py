"""Lay out one home's Medicare cost reports (form CMS-2540-10) as readable worksheets, straight from the CMS files.

Usage: python3 pipeline/scripts/cost_report_view.py 415035 > pipeline/cost_report_views/415035.md

Reads every pipeline/hcris/SNF10FY<year>_{rpt,nmrc,alpha}.csv (the Rhode Island rows of CMS's HCRIS files, saved by
the "Download Rhode Island cost reports from CMS" workflow) and prints, for each report the home filed:
  - the report's CMS record number, period, and which CMS file it is in
  - Worksheet S-2 Part I line 4: the name the home filed under
  - Worksheet A-8-1 Part A: costs paid to related organizations, line by line, all columns
  - Worksheet A-8-1 Part B: the owners that make each organization related, with ownership shares
  - Worksheet G-3 lines 1 to 3: total patient revenue, contractual allowances, and net patient revenue
Every number is printed as it appears in the CMS file, with its worksheet line and column, so anyone can match it
against CMS's own download. Nothing is added up or rounded here.
"""
import csv, sys
from collections import defaultdict
from pathlib import Path

HC = Path(__file__).resolve().parents[1] / "hcris"
ccn = sys.argv[1]

def money(v):
    try:
        x = float(v)
    except (TypeError, ValueError):
        return v or ""
    return f"({abs(x):,.0f})" if x < 0 else f"{x:,.0f}"

def ln(code):  # '00100' -> '1', '00102' -> '1.02'
    a, b = int(code[:3]), int(code[3:])
    return f"{a}" if b == 0 else f"{a}.{b:02d}"

reports = []
for rpt in sorted(HC.glob("SNF10FY*_rpt.csv")):
    name = rpt.name.replace("_rpt.csv", "")
    for r in csv.reader(open(rpt, encoding="latin-1")):
        if r[2].strip() == ccn:
            reports.append(dict(file=name, rec=r[0], status=r[4], begin=r[5], end=r[6], received=r[17] if len(r) > 17 else "",
                                processed=r[7]))
STATUS = {"1": "as submitted", "2": "settled without audit", "3": "settled with audit", "4": "reopened", "5": "amended"}

def load(name, kind, recs):
    d = defaultdict(dict)
    for r in csv.reader(open(HC / f"{name}_{kind}.csv", encoding="latin-1")):
        if r[0] in recs:
            d[(r[0], r[1])][(r[2], r[3])] = r[4]
    return d

by_file = defaultdict(set)
for x in reports: by_file[x["file"]].add(x["rec"])
nm, al = {}, {}
for f, recs in by_file.items():
    nm.update(load(f, "nmrc", recs)); al.update(load(f, "alpha", recs))

reports.sort(key=lambda x: (x["begin"][-4:], x["begin"]))
print(f"# Medicare cost reports for CMS number {ccn}, as filed\n")
print("Source: CMS Healthcare Cost Report Information System (HCRIS), skilled nursing facility form CMS-2540-10, "
      "files SNF10FY<year>.ZIP from https://www.cms.gov/data-research/statistics-trends-and-reports/cost-reports/cost-reports-fiscal-year. "
      "Each figure below is copied from the CMS file with its worksheet, line, and column. Negative numbers are in parentheses.\n")
for x in reports:
    rec = x["rec"]
    A = nm.get((rec, "A810001"), {}); At = al.get((rec, "A810001"), {})
    B = nm.get((rec, "A810002"), {}); Bt = al.get((rec, "A810002"), {})
    G = nm.get((rec, "G300000"), {})
    S = al.get((rec, "S200001"), {})
    print(f"## {x['begin']} to {x['end']}\n")
    print(f"- CMS report record number: **{rec}**, in file {x['file']}")
    print(f"- Report status in that file: {STATUS.get(x['status'], x['status'])}")
    nm_name = S.get(("00400", "00100"), "")
    if nm_name: print(f"- Name filed under (Worksheet S-2, line 4, column 1): {nm_name}")
    print()
    print("**Worksheet A-8-1, Part A: costs incurred and adjustments for services furnished by related organizations**\n")
    print("| Line | Col 1: Wkst A line | Col 2: cost center | Col 3: expense item | Col 4: allowable cost | Col 5: amount included in Wkst A | Col 6: adjustment |")
    print("| --- | --- | --- | --- | --- | --- | --- |")
    for line in sorted({k[0] for k in list(A) + list(At)}):
        g = lambda c: A.get((line, c), "")
        t = lambda c: At.get((line, c), "")
        print(f"| {ln(line)} | {g('00100')} | {t('00200')} | {t('00300')} | {money(g('00400'))} | {money(g('00500'))} | {money(g('00600'))} |")
    if not A and not At: print("| | | | none reported | | | |")
    print()
    print("**Worksheet A-8-1, Part B: interrelationship to related organizations**\n")
    print("| Line | Col 1: symbol | Col 2: name (owner of provider) | Col 3: % of provider | Col 4: related organization | Col 5: % of related organization | Col 6: type of business |")
    print("| --- | --- | --- | --- | --- | --- | --- |")
    for line in sorted({k[0] for k in list(B) + list(Bt)}):
        g = lambda c: B.get((line, c), "") or Bt.get((line, c), "")
        print(f"| {ln(line)} | {g('00100')} | {g('00200')} | {g('00300')} | {g('00400')} | {g('00500')} | {g('00600')} |")
    print()
    print("**Worksheet G-3: statement of revenues and expenses, column 1**\n")
    print("| Line | Item | Amount |")
    print("| --- | --- | --- |")
    for line, item in (("00100", "Total patient revenues"), ("00200", "Less contractual allowances and discounts"), ("00300", "Net patient revenues")):
        print(f"| {ln(line)} | {item} | {money(G.get((line, '00100'), ''))} |")
    print()
