"""Nurse staff hours by job, per Rhode Island home and quarter, from CMS's daily payroll records
("Payroll Based Journal Daily Nurse Staffing", data.cms.gov), saved to pipeline/pbj_quarters.csv.

Each run downloads only quarters not saved yet (all quarters from 2017 on the first time). The page uses
these to compare the nurse staffing a buyer projected in its application with what the payroll records
later showed. Days with fewer than 1.5 nurse hours per resident are left out as likely gaps in the
records (the same rule as the dashboard's staffing detail), and counted in gap_days.

  python3 pipeline/scripts/pbj_years.py
"""
import csv, datetime, json, os, re, sys, time, urllib.error, urllib.parse, urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'pbj_quarters.csv')
CATALOG = 'https://data.cms.gov/data.json'
TITLE = 'Payroll Based Journal Daily Nurse Staffing'
GROUPS = {'rn': ['RNDON', 'RNadmin', 'RN'], 'lpn': ['LPNadmin', 'LPN'], 'aide': ['CNA', 'NAtrn', 'MedAide']}
GAP_HPRD = 1.5
PAGE = 5000
FIRST_YEAR = 2017
COLS = ['ccn', 'year', 'quarter', 'days', 'gap_days', 'resident_days', 'rn_hours', 'lpn_hours', 'aide_hours', 'contract_hours']


def get(url, tries=4):
    last = None
    for n in range(tries):
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (compatible; abc-ri-update/1.1; +https://snfsales.abc-ri.org)', 'Accept': 'application/json'})
        try:
            with urllib.request.urlopen(req, timeout=180) as r: return r.read().decode('utf-8-sig', errors='replace')
        except Exception as e:
            last = f'{type(e).__name__}: {e} for {url[:100]}'
        print(f'try {n + 1} of {tries} failed: {last}', flush=True)
        if n < tries - 1: time.sleep([20, 60, 180][min(n, 2)])
    print('ERROR:', last); sys.exit(1)


def quarters():
    cat = json.loads(get(CATALOG)); out = {}
    for d in cat.get('dataset', []):
        if d.get('title', '').split(' : ')[0].strip().lower() != TITLE.lower(): continue
        for dist in d.get('distribution', []):
            url = dist.get('accessURL') or ''
            if dist.get('format', '').upper() != 'API' or '/data-api/' not in url: continue
            t = dist.get('temporal') or d.get('temporal') or ''
            start = t[0].get('startDate', '') if isinstance(t, list) and t and isinstance(t[0], dict) else str(t).split('/')[0]
            if re.match(r'\d{4}-\d{2}-\d{2}', start): out[(int(start[:4]), (int(start[5:7]) - 1) // 3 + 1)] = url
    if not out: print('ERROR: no quarterly API links for', TITLE); sys.exit(1)
    return out


def num(v):
    try: return float(v)
    except (TypeError, ValueError): return 0.0


def summarize(url, y, q):
    rows, off = [], 0
    while True:
        page = json.loads(get(f"{url}?{urllib.parse.urlencode({'filter[STATE]': 'RI', 'size': PAGE, 'offset': off})}"))
        rows += page; off += PAGE
        if len(page) < PAGE: break
    homes = {}
    for r in rows:
        r = {k.lower(): v for k, v in r.items()}   # older quarters use lower-case column names
        if str(r.get('state', 'RI')).upper() != 'RI': continue
        h = {g: sum(num(r.get(f'hrs_{c}'.lower())) for c in cs) for g, cs in GROUPS.items()}
        ctr = sum(num(r.get(f'hrs_{c}_ctr'.lower())) for cs in GROUPS.values() for c in cs)
        cen = num(r.get('mdscensus'))
        a = homes.setdefault(str(r['provnum']).strip().zfill(6), dict(days=0, gap_days=0, resident_days=0.0, rn_hours=0.0, lpn_hours=0.0, aide_hours=0.0, contract_hours=0.0))
        a['days'] += 1
        if cen > 0 and sum(h.values()) / cen < GAP_HPRD: a['gap_days'] += 1; continue
        a['resident_days'] += cen; a['contract_hours'] += ctr
        for g in GROUPS: a[g + '_hours'] += h[g]
    return [dict(ccn=c, year=y, quarter=q, **{k: (round(v, 1) if isinstance(v, float) else v) for k, v in a.items()}) for c, a in sorted(homes.items())]


def main():
    saved = list(csv.DictReader(open(OUT))) if os.path.exists(OUT) else []
    have = {(int(r['year']), int(r['quarter'])) for r in saved}
    avail = quarters()
    todo = [k for k in sorted(avail) if k[0] >= FIRST_YEAR and k not in have]
    if not todo: print('NOTHING NEW'); return
    for y, q in todo:
        rows = summarize(avail[(y, q)], y, q)
        if len(rows) < 40: print(f'ERROR: Q{q} {y} has only {len(rows)} Rhode Island homes; not saving'); sys.exit(1)
        saved += rows; print(f'Q{q} {y}: {len(rows)} homes', flush=True)
        saved.sort(key=lambda r: (int(r['year']), int(r['quarter']), r['ccn']))
        with open(OUT, 'w', newline='') as f:
            w = csv.DictWriter(f, fieldnames=COLS); w.writeheader(); w.writerows(saved)


if __name__ == '__main__':
    main()
