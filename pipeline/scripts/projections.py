"""What buyers projected in their change in effective control (CEC) applications, compared with what happened.

Every nursing home CEC application has an Appendix A with Table 2 (staff by job, in full-time equivalents) and
Table 3 (revenue and expenses) for the buyer's first three years. pipeline/projections.json holds those projected
figures for each completed sale, copied from the application page and checked by eye (cec_budgets.csv is where
they were first found; its numbers are not used directly because some tables did not read cleanly).

The actual figures for the same years come from:
  - nurse staffing: CMS's daily payroll records (pipeline/pbj_quarters.csv, from pbj_years.py), hours worked by
    RNs (including the director of nursing and RN administrators), LPNs, and aides, turned into full-time
    equivalents at 2,080 hours a year;
  - money: the home's Medicare cost reports (the same figures as the "Money paid to companies the business owners
    also control" table): net patient revenue, operating margin, rent paid to a landlord the business owners also
    control, and management fees paid to a company they also control.

A year is compared only if the buyer was in charge for all or nearly all of it (from its management start, if
earlier than the sale, to the next owner's start), and only if the actual figure covers a whole year.
"""
import csv, json, os
from collections import defaultdict

HOURS_PER_FTE = 2080
MIN_DAYS = 270          # payroll records must cover at least this many usable days of the year (three quarters)
FIELDS = ('nurse_fte', 'rn_fte', 'revenue', 'margin', 'rent', 'mgmt')


def _pbj(path):
    """(ccn, year) -> dict of FTEs, from the quarterly payroll file."""
    if not os.path.exists(path): return {}
    acc = defaultdict(lambda: defaultdict(float))
    for r in csv.DictReader(open(path)):
        a = acc[(r['ccn'], int(r['year']))]
        a['days'] += int(r['days']) - int(r['gap_days']); a['quarters'] += 1
        for k in ('rn_hours', 'lpn_hours', 'aide_hours', 'resident_days', 'contract_hours'): a[k] += float(r[k] or 0)
    out = {}
    for k, a in acc.items():
        if a['quarters'] < 3 or a['days'] < MIN_DAYS: continue   # a quarter with no file at all means the home did not submit it (CMS footnote 12)
        per_year = 365.0 / a['days'] / HOURS_PER_FTE
        tot = a['rn_hours'] + a['lpn_hours'] + a['aide_hours']
        out[k] = dict(nurse_fte=round(tot * per_year, 1), rn_fte=round(a['rn_hours'] * per_year, 1),
                      residents=round(a['resident_days'] / a['days'], 1) if a['days'] else None,
                      hprd=round(tot / a['resident_days'], 2) if a['resident_days'] else None,
                      contract_pct=round(a['contract_hours'] / tot * 100, 1) if tot else None, quarters=int(a['quarters']))
    return out


def _in_charge(t, year):
    """True when tenure t covers all or nearly all of the calendar year (starts by January 31, runs to December 1)."""
    start, end = t.get('counted_from') or t.get('start'), t.get('end')
    if not start or start > f'{year}-01-31': return False
    if end and end < f'{year}-12-01': return False
    return True


def _mixed(m, i, t):
    """True when the cost report year adds together two reports and the first of them ended on or before the buyer's
    start: part of that year's money is the seller's (The Friendly Home's 2024 includes the seller's January 2023 to
    January 30, 2024 report)."""
    fye = (m.get('fye') or [None] * 999)[i]
    if not fye or ',' not in fye or not t: return False
    start = t.get('counted_from') or t.get('start') or ''
    for d in fye.split(','):
        d = d.strip()
        try: iso = f'{d[6:10]}-{d[0:2]}-{d[3:5]}'
        except Exception: continue
        if len(d) == 10 and iso <= start: return True
    return False


def build(D, proj_path, pbj_path):
    if not os.path.exists(proj_path): return None   # the home's CMS number is taken from the sale on this page, by CEC id
    P = json.load(open(proj_path, encoding='utf-8'))
    pbj = _pbj(pbj_path)
    money = (D.get('money') or {}).get('by_facility', {})
    fac = {f['id']: f for f in D.get('facilities', [])}
    ten = {}
    for t in D.get('tenures', []):
        if t.get('cec'): ten[(t['fid'], t['cec'])] = t
    ccn_of = {x['cec']: x['fid'] for x in D.get('sales', []) if x.get('kind') == 'sale' and x.get('fid')}
    rows, last_pbj = [], max((y for (_, y) in pbj), default=None)
    for s in P.get('sales', []):
        cec = s['cec']; ccn = ccn_of.get(cec)
        if not ccn: continue
        t = ten.get((ccn, cec))
        m = money.get(ccn) or {}
        yrs_m = {y: i for i, y in enumerate(m.get('years', []))}
        staffing_ok, staffing_note, past = s.get('staffing_ok', True), s.get('staffing_note'), s.get('past')
        if past:
            pp = pbj.get((ccn, int(past['year'])))
            past = dict(past, payroll=pp['nurse_fte'] if pp else None)
            if pp and staffing_ok and not (0.8 <= pp['nurse_fte'] / past['nurse_fte'] <= 1.25):
                staffing_ok = False
                staffing_note = (f"The application's own count for a year the seller ran the home ({past['nurse_fte']:g} nurse staff in {past['year']}, page {past['page']}) "
                                 f"does not match CMS's payroll records for that year ({pp['nurse_fte']:g}), so its staff counts do not measure the same thing and are not compared.")
        years = []
        for y in sorted(s.get('years', {}), key=int):
            pr = s['years'][y]; yi = int(y)
            if t is None or not _in_charge(t, yi): continue
            act = {}
            st = pbj.get((ccn, yi))
            if st: act.update(nurse_fte=st['nurse_fte'], rn_fte=st['rn_fte'], hprd=st['hprd'], contract_pct=st['contract_pct'], staff_quarters=st['quarters'])
            i = yrs_m.get(yi)
            if i is not None and _mixed(m, i, t): i = None   # the year adds a report from before the buyer took charge
            if i is not None and not (m.get('partial') or [False] * 99)[i]:
                act.update(revenue=m['rev'][i], margin=m['margin'][i], rent=m['rent'][i], mgmt=m['mgmt'][i], expenses=(m.get('exp') or [None] * 999)[i])
            if not any(act.get(k) is not None and pr.get(k) is not None for k in FIELDS if staffing_ok or not k.endswith('_fte')): continue
            pm = pr.get('margin')
            if pm is None and pr.get('profit') is not None and pr.get('revenue'): pm = round(pr['profit'] / pr['revenue'], 4)
            years.append(dict(year=yi, projected={**{k: pr.get(k) for k in FIELDS}, 'margin': pm}, actual=act))
        cap = None
        if s.get('cost_cap'):
            first = next((y for y in years if y['actual'].get('expenses') is not None), None)
            cap = dict(s['cost_cap'], year=first['year'] if first else None, actual=first['actual']['expenses'] if first else None)
        f = fac.get(ccn, {})
        rows.append(dict(cec=cec, fid=ccn, name=f.get('name') or s.get('home'), town=f.get('town'), buyer=s.get('buyer') or (t or {}).get('group'),
                         start=(t or {}).get('counted_from'), via=(t or {}).get('via'), document=s.get('document'), link=s.get('link'), archive=s.get('archive'),
                         staff_page=s.get('staff_page'), budget_page=s.get('budget_page'), filed=s.get('filed'), note=s.get('note'),
                         staffing_ok=staffing_ok, staffing_note=staffing_note, past=past, cost_cap=cap,
                         first_year=min(map(int, s.get('years', {}) or [0])), years=years,
                         waiting=(not years), landlord=m.get('landlord')))
    rows.sort(key=lambda r: (r['start'] or '9999'))
    return dict(rows=rows, last_staffing_year=last_pbj, hours_per_fte=HOURS_PER_FTE)
