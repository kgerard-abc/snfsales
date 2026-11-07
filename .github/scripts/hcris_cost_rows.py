"""Turn the Rhode Island cost report files in pipeline/hcris into rows for the INTERNAL sheet's "Cost Reports by Year" tab.

Usage: python3 .github/scripts/hcris_cost_rows.py SNF24FY2025 SNF10FY2025 > rows.csv

Works for both CMS forms. Periods ending on or after Sept. 30, 2025 are filed on the new form, CMS-2540-24. What moved:
- Worksheet G-3: lines 1 to 5 are unchanged; net income moved from line 31 to line 32.
- Worksheet A-8-1 Part I: a new column 4 was added, so allowable cost is column 5 (was 4), the amount included in
  Worksheet A is column 6 (was 5), and the adjustment is column 7 (was 6). The total line is 100 (was 10).
- Worksheet A-8-1 Part II: names moved to columns 3 (person), 5 (related organization), and 8 (type of business).
- Worksheet S-3 Part I: Medicare and Medicaid HMO days are now on line 2; total days (column 7) and bed days (column 2)
  are still on line 1, and Medicaid days here include line 2.
Checked on November 7, 2026: with the 2540-10 settings this reproduces all 97 FY2023 and FY2024 single-report rows on
the tab exactly (revenue, expenses, margins, related-party amounts, occupancy, Medicaid share, and parties listed).
"""
import csv, re, json, sys
from collections import defaultdict
H = 'pipeline/hcris/'
FORM = {  # where each item lives on each form
  '10': dict(ni_line='03100', a81_total='01000', a81_allow='00400', a81_rec='00500', a81_adj='00600',
             pb_ind='00100', pb_name='00200', pb_org='00400', pb_type='00600', s3_total='00100'),
  '24': dict(ni_line='03200', a81_total='10000', a81_allow='00500', a81_rec='00600', a81_adj='00700',
             pb_ind='00100', pb_name='00300', pb_org='00500', pb_type='00800', s3_total='00100'),
}
RENT = re.compile(r'RENT|LEASE', re.I)
BLDG = re.compile(r'BLDG|BUILDING', re.I)
MGMT = re.compile(r'MGMT|MANAGEMENT|MGT|ADMIN|HOME OFFICE|CONSULT', re.I)
def load(name):
    form = '24' if name.startswith('SNF24') else '10'; F = FORM[form]
    rpt = {r[0]: r for r in csv.reader(open(H + name + '_rpt.csv'))}
    num = defaultdict(dict); txt = defaultdict(dict)
    for r in csv.reader(open(H + name + '_nmrc.csv')): num[r[0]][(r[1], r[2], r[3])] = float(r[4])
    for r in csv.reader(open(H + name + '_alpha.csv')): txt[r[0]][(r[1], r[2], r[3])] = r[4].strip()
    out = []
    for rec, r in rpt.items():
        n, t = num[rec], txt[rec]
        g = lambda l: n.get(('G300000', l, '00100'))
        items = sorted({k[1] for k in n if k[0] == 'A810001' and k[1] != F['a81_total']})
        rec_amt = lambda l: n.get(('A810001', l, F['a81_rec']), 0.0)
        label = lambda l: t.get(('A810001', l, '00300'), '')
        rp_total = n.get(('A810001', F['a81_total'], F['a81_rec']))
        if rp_total is None and items: rp_total = sum(rec_amt(l) for l in items)
        center = lambda l: t.get(('A810001', l, '00200'), '')
        rent = sum(rec_amt(l) for l in items if RENT.search(label(l)) or BLDG.search(center(l))) if items else None
        mgmt = sum(rec_amt(l) for l in items if MGMT.search(label(l))) if items else None
        adj = n.get(('A810001', F['a81_total'], F['a81_adj']))
        parties = []
        for l in sorted({k[1] for k in t if k[0] == 'A810002'}):
            nm, org = t.get(('A810002', l, F['pb_name']), ''), t.get(('A810002', l, F['pb_org']), '')
            typ = t.get(('A810002', l, F['pb_type']), '')
            right = (org + (f' ({typ})' if typ else '')).strip()
            if nm and right: parties.append(f'{nm} -> {right}')
            elif right: parties.append(f'-> {right}')
            elif nm: parties.append(nm)
        s3 = lambda c: n.get(('S300001', F['s3_total'], c))
        days, beddays, mcaid = s3('00700'), s3('00200'), s3('00500')
        if form == '24':  # Medicaid HMO days are on line 2 of the new form
            mcaid = (mcaid or 0) + (n.get(('S300001', '00200', '00500')) or 0)
        out.append(dict(rec=rec, ccn=r[2].strip(), form=form, fy_begin=r[5], fy_end=r[6], status=r[4],
                        rev=g('00300'), exp=g('00400'), nipc=g('00500'), ni=g(F['ni_line']),
                        rp=rp_total, rent=rent, mgmt=mgmt, adj=adj, days=days, beddays=beddays, mcaid=mcaid,
                        parties='; '.join(dict.fromkeys(parties))))
    return out
if __name__ == '__main__':
    rows = [x for n in sys.argv[1:] for x in load(n)]
    w = csv.DictWriter(sys.stdout, fieldnames=list(rows[0]))
    w.writeheader(); w.writerows(rows)
