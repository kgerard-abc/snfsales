# =====================================================================================
# ri_nh_build.py - builds the combined CMS + CEC workbook and the public web page
# from the standardized CMS files (combined_standardized/) and the CEC workbook.
# Everything is derived from the data: rerun after each new CMS release.
# =====================================================================================
import os, re, io, json, math, shutil, datetime as dt
import numpy as np
import pandas as pd

DEFAULT_CONFIG = dict(
    anchor='2019-08-26',            # first snapshot on the quarterly grid (grid = every 3 months from here)
    sff_as_zero=True,               # months CMS leaves a home unrated because it is a Special Focus Facility (footnote 18) count as 0 stars
    min_plausible_hprd=1.5,         # reported nurse hours per resident per day below this are treated as a payroll-reporting problem (blanked, flagged)
    event_overrides_seed=[          # (CEC ID, treat as, date, public label, note) - written to the Event Overrides tab the first time only
        ('CEC-077', 'closure', '2021-09-01', 'Closed September 2021', 'Per the workbook owner, this home closed in September 2021 and was not sold. Not counted as a sale.'),
    ],
    max_plausible_turnover=90,      # nursing staff turnover (%) at or above this is CMS's 'lowest possible score' for missing payroll data: blanked and flagged
    ccn_notes_seed={'Pawtuxet Village Care and Rehabilitation Center': 'Closed April 2012. No CCN exists.'},   # notes added to the CCN Map the first time a name appears
    closed_seed={'415107': 'Closed January 2025', '415051': 'Closed 2025'},   # homes CMS still lists that have closed (CCN: 'Closed <Month> <Year>' or 'Closed <Year>'); a 'Closed ...' note in the CCN Map does the same and wins
    max_gap_days=45,
    series_start='2019-04-09',      # first CMS file after the April 2019 star-rating method change (statewide average fell 3.8 -> 3.35 in one month)                # a grid date needs a CMS file within this many days
    page_title='Rhode Island nursing homes: quality trends and what happened when homes were sold',
    page_credit='Built and maintained by ABC-RI from CMS and RIDOH public records',   # shown in the page footer
    page_contact='corrections@abc-ri.org',   # shown on the page for corrections
    data_cautions=[                 # (title, text) shown on the workbook Read Me
        ('AdviniaCare Orchard, LLC (CCN 415059)', 'CMS average residents per day falls from 126.5 in the June 2026 file to 30.1 in the July 2026 file and stays there, so its latest occupancy shows as 18%. Per the workbook owner, the facility temporarily evacuated most floors, so this reflects that event rather than a reporting change.'),
    ],
    public_show_principals=False,   # False: public page shows companies only (drops names in parentheses and after the arrow)
    public_include_types=('Ownership/operator change', 'Corporate parent change (multi-facility deal)'),
)

def nz(v):
    if v is None: return None
    if isinstance(v, float) and np.isnan(v): return None
    if v is pd.NaT: return None
    try:
        if pd.isna(v): return None
    except Exception: pass
    return v
def key_(x): return re.sub(r'[^A-Z0-9]', '', str(x).upper()) if nz(x) is not None else None
def lab_(s): return pd.Timestamp(s).strftime('%b %Y')
def nice(n):
    n = str(n).strip()
    if n.isupper():
        n = n.title()
        n = re.sub(r"'S\b", "'s", n); n = re.sub(r"\bLlc\b", "LLC", n); n = re.sub(r"\bIi\b", "II", n); n = re.sub(r"\bIii\b", "III", n); n = re.sub(r"\bInc\b", "Inc", n)
        n = re.sub(r"(?<=\s)(And|Of|At|The|For)(?=\s)", lambda m: m.group(1).lower(), n)
    return n
def clean_lbn(x):
    x = nz(x)
    if x is None: return None
    x = str(x).strip()
    return None if (not x or 'not available' in x.lower()) else x
def serial(d):   # excel date serial for a date-like
    return (pd.Timestamp(d) - pd.Timestamp('1899-12-30')).days

NUMCOLS = ['overall_rating','health_inspection_rating','staffing_rating','qm_rating','rn_staffing_rating','long_stay_qm_rating','short_stay_qm_rating',
           'reported_total_nurse_staffing_hours_per_resident_per_day','reported_rn_staffing_hours_per_resident_per_day','adjusted_total_nurse_staffing_hours_per_resident_per_day',
           'number_of_certified_beds','average_number_of_residents_per_day','number_of_fines','total_amount_of_fines_in_dollars','number_of_payment_denials','total_number_of_penalties',
           'total_nursing_staff_turnover','registered_nurse_turnover']

def load_cms(std_dir):
    """Read the four standardized CMS files the builder needs."""
    rd = lambda n, **k: pd.read_csv(os.path.join(std_dir, n), dtype=str, low_memory=False, **k)
    p = rd('providerinfo.csv')
    for c in NUMCOLS:
        if c in p.columns: p[c] = pd.to_numeric(p[c], errors='coerce')
        else: p[c] = np.nan
    for c in ['legal_business_name','chain_name','special_focus_status','abuse_icon','ownership_type','provider_changed_ownership_in_last_12_months','provider_name','city_town','county_parish']:
        if c not in p.columns: p[c] = np.nan
    hall = rd('healthcitations.csv'); pall = rd('penalties.csv'); own = rd('ownership.csv')
    return dict(p=p.sort_values(['ccn_std','snapshot_date']), hall=hall, pall=pall, own=own)

def make_grid(snaps, anchor, max_gap):
    """Quarterly grid every 3 months from `anchor`; annual points are every 4th (Aug/Sep). Returns points, flags."""
    latest = pd.Timestamp(snaps[-1]); pts, ann, k = [], [], 0
    while True:
        t = pd.Timestamp(anchor) + pd.DateOffset(months=3 * k)
        best = min(snaps, key=lambda s: abs((pd.Timestamp(s) - t).days))
        if abs((pd.Timestamp(best) - t).days) > max_gap and t > latest: break
        if t > latest + pd.Timedelta(days=max_gap): break
        if not pts or best != pts[-1]:
            pts.append(best); ann.append(k % 4 == 0)
        k += 1
    return pts, ann

def drop_removed(df_all, keys, date_col, snaps_arr, label=''):
    """A record CMS stopped listing while its date was still inside CMS's window (< 2 years old) was removed by CMS."""
    lastseen = df_all.groupby(keys)['snapshot_date'].max().rename('last_seen_snap').reset_index()
    lim = (pd.to_datetime(lastseen[date_col]) + pd.DateOffset(years=2)).dt.strftime('%Y-%m-%d')
    lastseen['limit'] = lim.where(lim < snaps_arr[-1], snaps_arr[-1])
    idx = np.searchsorted(snaps_arr, lastseen['limit'].values, side='right') - 1
    lastseen['cover_snap'] = np.where(idx >= 0, snaps_arr[np.clip(idx, 0, None)], '')
    lastseen['removed'] = (lastseen['cover_snap'] != '') & (lastseen['cover_snap'] > lastseen['last_seen_snap']) & (lastseen[date_col] >= snaps_arr[0])
    return lastseen.loc[lastseen.removed, keys]

def unique_citations_penalties(inp, snaps_arr):
    hall, pall = inp['hall'], inp['pall']
    h_snaps = np.array(sorted(hall.snapshot_date.unique()))       # each dataset is judged against its own releases
    p_snaps = np.array(sorted(pall.snapshot_date.unique()))
    k4 = ['ccn_std','survey_date','deficiency_prefix','deficiency_tag_number']
    hu = hall.sort_values('snapshot_date').drop_duplicates(k4, keep='last').copy()
    rem = drop_removed(hall[['snapshot_date'] + k4], k4, 'survey_date', h_snaps)
    hu = hu.merge(rem.assign(_rm=1), on=k4, how='left'); hu = hu[hu['_rm'].isna()].drop(columns='_rm')
    hu['serious'] = hu['scope_severity_code'].isin(list('GHIJKL'))
    hu['abuse'] = hu['deficiency_category'].fillna('').str.startswith('Freedom from Abuse')
    kp3 = ['ccn_std','penalty_date','penalty_type']
    last_snap = pall.groupby(kp3)['snapshot_date'].transform('max')
    pu = pall[pall['snapshot_date'] == last_snap].copy()
    remp = drop_removed(pall[['snapshot_date'] + kp3], kp3, 'penalty_date', p_snaps)
    pu = pu.merge(remp.assign(_rm=1), on=kp3, how='left'); pu = pu[pu['_rm'].isna()].drop(columns='_rm')
    pu['fine_amount'] = pd.to_numeric(pu['fine_amount'], errors='coerce')
    return hu, pu

# ------------------------------------------------------------------ facility model
def build_model(inp, cfg):
    p = inp['p'].copy(); hall, pall, own = inp['hall'], inp['pall'], inp['own']
    p['sff_nr'] = (p['overall_rating_footnote'].astype(str).str.replace('.0', '', regex=False) == '18') if 'overall_rating_footnote' in p.columns else False
    if cfg.get('sff_as_zero', True):
        for c in ['overall_rating', 'health_inspection_rating', 'staffing_rating', 'qm_rating', 'rn_staffing_rating', 'long_stay_qm_rating', 'short_stay_qm_rating']:
            if c in p.columns: p.loc[p['sff_nr'] & p[c].isna(), c] = 0.0
    HC = 'reported_total_nurse_staffing_hours_per_resident_per_day'
    p['hprd_flag'] = p[HC] < cfg.get('min_plausible_hprd', 1.5)
    M_hprd = p[p['hprd_flag']][['ccn_std', 'snapshot_date', HC]].copy()
    p['turn_flag'] = p['total_nursing_staff_turnover'] >= cfg.get('max_plausible_turnover', 90)
    M_turn = p[p['turn_flag']][['ccn_std', 'snapshot_date', 'total_nursing_staff_turnover']].copy()
    for c in ('total_nursing_staff_turnover', 'registered_nurse_turnover'): p.loc[p['turn_flag'], c] = np.nan
    for c in (HC, 'reported_rn_staffing_hours_per_resident_per_day', 'adjusted_total_nurse_staffing_hours_per_resident_per_day'):
        p.loc[p['hprd_flag'], c] = np.nan
    snaps = sorted(p.snapshot_date.unique()); snaps_arr = np.array(snaps); latest = snaps[-1]
    M = dict(cfg=cfg, snaps=snaps, latest=latest, p=p, hprd_rows=M_hprd, turn_rows=M_turn)
    pts, ann = make_grid(snaps, cfg['anchor'], cfg['max_gap_days'])
    M['qpts'], M['is_ann'] = pts, ann
    hu, pu = unique_citations_penalties(inp, snaps_arr); M['hu'], M['pu'] = hu, pu
    M['ds_latest'] = {k: (str(v['snapshot_date'].max()) if 'snapshot_date' in v.columns and len(v) else None)
                      for k, v in (('Ratings, staffing and turnover', p), ('Inspection citations', inp['hall']), ('Fines and penalties', inp['pall']), ('Owner lists', inp['own']))}
    M['ds_behind'] = {k: v for k, v in M['ds_latest'].items() if v and v < latest}
    # ---- facilities
    g = p.groupby('ccn_std')['snapshot_date'].agg(first_seen='min', last_seen='max')
    last_row = p.groupby('ccn_std').tail(1).set_index('ccn_std')
    fac = last_row.join(g)
    fac['cms_name'] = fac['provider_name'].map(nice); fac['city'] = fac['city_town'].str.title(); fac['county'] = fac['county_parish'].str.title()
    fac['active'] = fac['last_seen'] == latest
    fac['status'] = [('Active' if a else f"Not in CMS data after {lab_(l)}") for a, l in zip(fac['active'], fac['last_seen'])]
    def messy(n):
        l = n.lower(); return (' dba ' in l) or ('operations' in l) or ('operator' in l) or len(n) == 50 or 'legal business name' in l
    names_hist = {}
    for ccn, d in p.groupby('ccn_std'):
        seq = []
        for s, nm in zip(d.snapshot_date, d.provider_name):
            if nz(nm) is None: continue
            k = key_(nm)
            if seq and seq[-1]['key'] == k: seq[-1]['last'] = s
            else: seq.append(dict(key=k, name=nice(nm), first=s, last=s))
        names_hist[ccn] = seq
    def pick(ccn):
        seq = names_hist[ccn]; rec = [x for x in seq if x['last'] >= '2025-08-01' and not messy(x['name'])]
        return (rec or seq)[-1]['name']
    fac['name'] = [pick(c) for c in fac.index]
    fac['chain'] = fac['chain_name'].map(lambda v: (v.title() if isinstance(v, str) and v.isupper() else v))
    lb = p.assign(lbn=p['legal_business_name'].map(clean_lbn)); lb['lbn'] = lb.groupby('ccn_std')['lbn'].ffill()
    fac['operator'] = lb.groupby('ccn_std')['lbn'].last().map(lambda v: nice(v) if nz(v) else None)
    fac = fac.sort_values(['active', 'name'], ascending=[False, True])
    M['fac'], M['order'], M['names_hist'], M['lb'] = fac, fac.index.tolist(), names_hist, lb
    # ---- point-in-time pivots
    def piv(col, scale=1.0):
        d = p[p.snapshot_date.isin(pts)].pivot(index='ccn_std', columns='snapshot_date', values=col).reindex(M['order'])
        d = d.reindex(columns=pts); return d * scale if scale != 1.0 else d
    M['piv'] = piv
    # ---- calendar-quarter aggregates of citations / fines
    first_year = int(pd.Timestamp(pts[0]).year); ly = pd.Timestamp(latest)
    years = list(range(first_year, ly.year + 1)); lq = (ly.month - 1) // 3 + 1
    qpairs = [(y, q) for y in years for q in (1, 2, 3, 4) if (y, q) <= (ly.year, lq)]
    M['years'], M['qpairs'], M['last_q'] = years, qpairs, lq
    order = M['order']
    def qpiv(df, datecol, agg='size', valcol=None):
        d = df.copy(); dtm = pd.to_datetime(d[datecol]); d['_y'] = dtm.dt.year; d['_q'] = (dtm.dt.month - 1) // 3 + 1
        d = d[d['_y'].isin(years) & d['ccn_std'].isin(order)]
        gg = d.groupby(['ccn_std', '_y', '_q']); s = gg.size() if agg == 'size' else gg[valcol].sum()
        t = s.unstack(['_y', '_q']).reindex(index=order, columns=pd.MultiIndex.from_tuples(qpairs), fill_value=0).fillna(0)
        qstart = {(y, q): f'{y}-{(q - 1) * 3 + 1:02d}-01' for y, q in qpairs}
        m = pd.DataFrame({k: (fac['last_seen'].reindex(order) >= v) for k, v in qstart.items()}, index=order)
        return t.where(m)
    M['CQ'] = dict(cit_all=qpiv(hu, 'survey_date'), cit_ser=qpiv(hu[hu.serious], 'survey_date'), cit_abu=qpiv(hu[hu.abuse], 'survey_date'),
                   fines_amt=qpiv(pu[pu.penalty_type == 'Fine'], 'penalty_date', agg='sum', valcol='fine_amount'), fines_n=qpiv(pu[pu.penalty_type == 'Fine'], 'penalty_date'),
                   denials=qpiv(pu[pu.penalty_type == 'Payment Denial'], 'penalty_date'))
    # ---- monthly window (all snapshots from the first grid point on)
    win = p[p.snapshot_date >= cfg['series_start']].copy()
    win = win.merge(fac[['name', 'city']], left_on='ccn_std', right_index=True, how='left')
    win['lbn'] = win['legal_business_name'].map(clean_lbn).map(lambda v: nice(v) if nz(v) else None)
    M['win'] = win
    # ---- operator grid (forward-filled legal business name)
    lb2 = lb[lb.snapshot_date.isin(pts)].pivot(index='ccn_std', columns='snapshot_date', values='lbn').reindex(order).reindex(columns=pts)
    M['op'] = lb2.map(lambda v: nice(v) if nz(v) else None)
    return M

# ------------------------------------------------------------------ change events (monthly resolution)
def build_events(M, inp):
    p, fac, snaps = M['p'].copy(), M['fac'], M['snaps']; snap_idx = {s: i for i, s in enumerate(snaps)}
    pi = p.copy(); events = []
    def add(s, c, cat, what, b, a, note=''): events.append((s, c, cat, what, '' if nz(b) is None else str(b), '' if nz(a) is None else str(a), note))
    def changes(col, cat, what, keyfn=key_, start=None, disp=lambda v: v, require_both=True):
        for c, g in pi.groupby('ccn_std'):
            g = g[g.snapshot_date >= start] if start else g; prev_raw, prev_key = None, None
            for s, v in zip(g.snapshot_date, g[col]):
                k = keyfn(v) if nz(v) is not None else None
                if k is None:
                    if not require_both and prev_key is not None: add(s, c, cat, what, disp(prev_raw), '')
                    if not require_both: prev_raw, prev_key = None, None
                    continue
                if prev_key is not None and k != prev_key: add(s, c, cat, what, disp(prev_raw), disp(v))
                prev_raw, prev_key = v, k
    pi['lbn'] = pi['legal_business_name'].map(clean_lbn)
    changes('lbn', 'Operator change', 'Legal business name (operator) changed', disp=nice)
    fp = lambda v: 'FP' if str(v).lower().startswith('for profit') else ('NP' if str(v).lower().startswith('non profit') else 'GOV')
    changes('ownership_type', 'Profit status change', 'For-profit / non-profit status changed', keyfn=fp)
    changes('chain_name', 'Chain change', 'Chain affiliation changed', disp=nice)
    changes('provider_name', 'Name change', 'Facility name in CMS listing changed', disp=nice)
    changes('special_focus_status', 'Special focus status', 'Special focus status changed', keyfn=lambda v: str(v).strip().upper(), start='2019-07-09', require_both=False)
    changes('abuse_icon', 'Abuse icon', 'Abuse icon changed', keyfn=lambda v: str(v).strip().upper(), start='2019-10-09')
    changes('number_of_certified_beds', 'Certified beds change', 'Certified beds changed', keyfn=lambda v: float(v) if nz(v) is not None else None, disp=lambda v: int(float(v)))
    changes('overall_rating', 'Rating change', 'Overall star rating changed', keyfn=lambda v: float(v), disp=lambda v: int(float(v)))
    for c, g in pi.groupby('ccn_std'):
        prev = None
        for s, v in zip(g.snapshot_date, g['provider_changed_ownership_in_last_12_months']):
            v = str(v).strip().upper() if nz(v) is not None else None
            if prev == 'N' and v == 'Y': add(s, c, 'Change of ownership (CMS flag)', 'CMS flags a change of ownership in the last 12 months', 'N', 'Y', 'Flag stays Y for 12 months after the change')
            if v: prev = v
    for c, r in fac.iterrows():
        if not r['active']:
            d = pi[pi.ccn_std == c].snapshot_date.max(); nxt = [s for s in snaps if s > d][0]
            add(nxt, c, 'Left CMS data', 'Facility no longer listed in CMS provider file', f'Last listed {d}', '', 'Closed, merged, or no longer certified (not stated in the data)')
    ev = pd.DataFrame(events, columns=['snap', 'ccn', 'category', 'what', 'before', 'after', 'note'])
    # drop quick reversals (A -> B -> A within 3 monthly files) in the noisy categories
    keep = pd.Series(True, index=ev.index)
    for (c, cat), g in ev[ev.category.isin(['Operator change', 'Profit status change', 'Chain change', 'Name change', 'Special focus status', 'Abuse icon'])].groupby(['ccn', 'category']):
        rows = list(g.sort_values('snap').itertuples())
        for a, b in zip(rows, rows[1:]):
            if keep[a.Index] and keep[b.Index] and snap_idx[b.snap] - snap_idx[a.snap] <= 3 and key_(b.after) == key_(a.before): keep[a.Index] = False; keep[b.Index] = False
    ev = ev[keep]
    # ---- owner-list events: 5%+ direct/indirect owners and managerial control only
    o = inp['own'].copy(); o['nm'] = o.owner_name.map(lambda x: re.sub(r'[^A-Z0-9]', '', str(x).upper())); o['role'] = o.role_played_by_owner_or_manager_in_facility
    STABLE = {'5% OR GREATER DIRECT OWNERSHIP INTEREST': 'direct owner', '5% OR GREATER INDIRECT OWNERSHIP INTEREST': 'indirect owner', 'OPERATIONAL/MANAGERIAL CONTROL': 'managerial control', 'DIRECT OWNERSHIP INTEREST': 'direct owner', 'INDIRECT OWNERSHIP INTEREST': 'indirect owner'}
    o = o[o.role.isin(STABLE)]; osn = sorted(o.snapshot_date.unique()); oidx = {s: i for i, s in enumerate(osn)}
    life = o.groupby('ccn_std')['snapshot_date'].agg(['min', 'max']); lrow = o.sort_values('snapshot_date').groupby(['ccn_std', 'nm']).tail(1).set_index(['ccn_std', 'nm'])
    oev = []
    for (c, n), x in o.groupby(['ccn_std', 'nm']):
        ii = sorted(set(oidx[s] for s in x.snapshot_date)); runs, st = [], ii[0]
        for a, b in zip(ii, ii[1:]):
            if b - a - 1 > 2: runs.append((st, a)); st = b
        runs.append((st, ii[-1])); first, last_ = oidx[life.loc[c, 'min']], oidx[life.loc[c, 'max']]; d = lrow.loc[(c, n)]
        roles = ', '.join(sorted(set(STABLE[r] for r in x.role))); assoc = str(d.association_date).replace('since ', '') if nz(d.association_date) else ''
        label = f"{nice(d.owner_name)} ({roles}" + (f", associated since {assoc}" if assoc else '') + ')'
        for a, b in runs:
            if a > first: oev.append((osn[a], c, 'added', label))
            if b < last_ - 2: oev.append((osn[b + 1], c, 'removed', label))
    oe = pd.DataFrame(oev, columns=['snap', 'ccn', 'kind', 'label']); tot = oe.groupby('snap').size(); bulk = set(tot[tot >= 40].index); rows = []
    for (s, c), g in oe.groupby(['snap', 'ccn']):
        add_ = '; '.join(g[g.kind == 'added'].label); rem = '; '.join(g[g.kind == 'removed'].label)
        what = ('Owners / controllers added: ' + add_ if add_ else '') + (' | ' if add_ and rem else '') + ('Removed: ' + rem if rem else '')
        if len(what) > 700: what = what[:697] + '...'
        note = f'Bulk update: {tot[s]} owner-list changes across RI facilities in this monthly file, so this may reflect a CMS data refresh rather than a sale.' if s in bulk else ''
        rows.append((s, c, 'Owner list change', what, '', '', note))
    ev = pd.concat([ev, pd.DataFrame(rows, columns=ev.columns)], ignore_index=True); ev['ccn'] = ev['ccn'].astype(str)
    OWN_CATS = ['Change of ownership (CMS flag)', 'Operator change', 'Profit status change', 'Chain change', 'Owner list change']; STRONG = ['Change of ownership (CMS flag)', 'Operator change', 'Profit status change']
    M['ev_own'] = ev[ev.category.isin(OWN_CATS)].copy(); M['ev_other'] = ev[~ev.category.isin(OWN_CATS)].sort_values(['snap', 'ccn'], ascending=[False, True]); M['strong'] = STRONG
    return M

# ------------------------------------------------------------------ CEC workbook reading
def _v(c):
    v = c.value
    return None if (isinstance(v, str) and v.startswith('=')) else v
def _hmap(ws, row=1):
    return {re.sub(r'\s+', ' ', str(ws.cell(row=row, column=c).value or '')).strip(): c for c in range(1, ws.max_column + 1) if ws.cell(row=row, column=c).value}
def _col(h, *starts):
    for k, c in h.items():
        if any(k.startswith(s) for s in starts): return c
    return None
def _date(v):
    v = nz(v)
    if v is None: return None
    if isinstance(v, (dt.datetime, dt.date, pd.Timestamp)): return pd.Timestamp(v).normalize()
    s = str(v).strip()
    for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%Y-%m-%d %H:%M:%S'):
        try: return pd.Timestamp(dt.datetime.strptime(s, fmt))
        except Exception: pass
    return None
def _ccn(v):
    v = nz(v)
    if v is None: return None
    s = re.sub(r'\D', '', str(v))
    return s.zfill(6) if s else None

def read_cec(wb):
    out = {}
    ws = wb['CEC Log']; h = _hmap(ws); rows = []
    cols = dict(id=_col(h, 'CEC ID'), fac=_col(h, 'Facility'), town=_col(h, 'Town'), beds=_col(h, 'Beds'), applicant=_col(h, 'Applicant'), buyer=_col(h, 'Buyer group'), type=_col(h, 'Transaction type'),
                outcome=_col(h, 'Outcome'), decision=_col(h, 'Decision or status date'), year=_col(h, 'Year'), chow=_col(h, 'CMS takeover date'), seller=_col(h, 'Seller per CMS'))
    for r in range(2, ws.max_row + 1):
        cid = _v(ws.cell(row=r, column=cols['id']))
        if not cid: continue
        rows.append(dict(row=r, cec=str(cid).strip(), facility=_v(ws.cell(row=r, column=cols['fac'])), town=_v(ws.cell(row=r, column=cols['town'])), applicant=_v(ws.cell(row=r, column=cols['applicant'])),
                         buyer=_v(ws.cell(row=r, column=cols['buyer'])), type=_v(ws.cell(row=r, column=cols['type'])), outcome=_v(ws.cell(row=r, column=cols['outcome'])),
                         decision=_date(_v(ws.cell(row=r, column=cols['decision']))), chow=_date(_v(ws.cell(row=r, column=cols['chow']))), year=_v(ws.cell(row=r, column=cols['year'])),
                         seller=_v(ws.cell(row=r, column=cols['seller']))))
    out['log'] = pd.DataFrame(rows); out['log_cols'] = cols
    ws = wb['Post-Sale Tracker']; rows = []
    for r in range(3, ws.max_row + 1):
        cid = _v(ws.cell(row=r, column=1))
        if not cid: continue
        rows.append(dict(row=r, cec=str(cid).strip(), facility=_v(ws.cell(row=r, column=2)), new_owner=_v(ws.cell(row=r, column=3)), outcome=_v(ws.cell(row=r, column=5)),
                         decision=_date(_v(ws.cell(row=r, column=6))), chow=_date(_v(ws.cell(row=r, column=9)))))
    out['tracker'] = pd.DataFrame(rows); out['tracker_last_row'] = ws.max_row
    ws = wb['Facility Summary']; h = _hmap(ws); rows = []
    for r in range(2, ws.max_row + 1):
        n = _v(ws.cell(row=r, column=_col(h, 'Facility')))
        if n: rows.append(dict(row=r, name=str(n).strip(), other=_v(ws.cell(row=r, column=_col(h, 'Other names'))), town=_v(ws.cell(row=r, column=_col(h, 'Town')))))
    out['fs'] = pd.DataFrame(rows)
    pairs = {}
    for sh in ('CMS Current (RI)', 'Cost Reports by Year'):
        if sh in wb.sheetnames:
            ws = wb[sh]; h = _hmap(ws); cc, cf = _col(h, 'CCN'), _col(h, 'Facility (this workbook)', 'Facility')
            if cc and cf:
                for r in range(2, ws.max_row + 1):
                    a, b = _ccn(_v(ws.cell(row=r, column=cc))), _v(ws.cell(row=r, column=cf))
                    if a and b: pairs.setdefault(str(b).strip(), a)
    out['pairs'] = pairs
    existing = {}
    if 'CCN Map' in wb.sheetnames:
        ws = wb['CCN Map']
        for r in range(5, ws.max_row + 1):
            n = _v(ws.cell(row=r, column=1))
            if n: existing[str(n).strip()] = (_ccn(_v(ws.cell(row=r, column=2))), _v(ws.cell(row=r, column=5)))
    out['existing_map'] = existing
    cr = []
    if 'Cost Reports by Year' in wb.sheetnames:
        ws = wb['Cost Reports by Year']; h = _hmap(ws)
        g = lambda name: _col(h, name)
        cols = dict(ccn=g('CCN'), fac=g('Facility'), yr=g('FY end year'), rev=g('Net patient revenue'), exp=g('Operating expenses'), ni=g('Net income'), marg=g('Operating margin'),
                    rp=g('Related-party costs recorded'), rent=g('Rent / building costs paid to related party'), mgmt=g('Management / admin fees paid to related party'), disallow=g('Markup disallowed'), parties=g('Related parties listed'))
        for r in range(2, ws.max_row + 1):
            c_, y_ = _ccn(_v(ws.cell(row=r, column=cols['ccn']))), _v(ws.cell(row=r, column=cols['yr']))
            if not c_ or not y_: continue
            num = lambda k: (pd.to_numeric(_v(ws.cell(row=r, column=cols[k])), errors='coerce') if cols[k] else np.nan)
            cr.append(dict(ccn=c_, year=int(y_), rev=num('rev'), exp=num('exp'), ni=num('ni'), margin=num('marg'), rp=num('rp'), rent=num('rent'), mgmt=num('mgmt'), disallow=num('disallow'),
                           parties=(_v(ws.cell(row=r, column=cols['parties'])) if cols['parties'] else None)))
    out['costs'] = pd.DataFrame(cr)
    mg = {}
    if 'Management Agreements' in wb.sheetnames:
        ws = wb['Management Agreements']
        for r in range(5, ws.max_row + 1):
            cid = _v(ws.cell(row=r, column=1))
            if cid: mg[str(cid).strip()] = dict(start=_date(_v(ws.cell(row=r, column=3))), who=_v(ws.cell(row=r, column=4)), src=_v(ws.cell(row=r, column=5)))
    out['mgmt'] = mg; out['mgmt_present'] = 'Management Agreements' in wb.sheetnames
    ov = {}; out['overrides_present'] = 'Event Overrides' in wb.sheetnames
    if out['overrides_present']:
        ws = wb['Event Overrides']
        for r in range(5, ws.max_row + 1):
            cid = _v(ws.cell(row=r, column=1))
            if cid: ov[str(cid).strip()] = dict(treat=str(_v(ws.cell(row=r, column=2)) or '').strip().lower(), date=_date(_v(ws.cell(row=r, column=3))), label=_v(ws.cell(row=r, column=4)), note=_v(ws.cell(row=r, column=5)))
    out['overrides'] = ov
    return out

def build_ccn_map(cec, M):
    """One row per facility name used anywhere in the CEC tabs. User entries in the CCN Map tab always win."""
    names = []
    for df, c in ((cec['fs'], 'name'), (cec['log'], 'facility'), (cec['tracker'], 'facility')):
        for n in df[c].dropna().astype(str).str.strip(): 
            if n not in names: names.append(n)
    known = {}   # normalized name -> set(ccn)
    for ccn, seq in M['names_hist'].items():
        for x in seq: known.setdefault(x['key'], set()).add(ccn)
        known.setdefault(key_(M['fac'].loc[ccn, 'name']), set()).add(ccn)
    rows = []
    for n in names:
        if n in cec['existing_map'] and (cec['existing_map'][n][0] or cec['existing_map'][n][1] is not None):
            ccn, note = cec['existing_map'][n]; how = 'entered in CCN Map'
        elif n in cec['pairs']: ccn, note, how = cec['pairs'][n], None, 'from CMS Current / Cost Reports tabs'
        else:
            cand = known.get(key_(n), set()); ccn, note, how = (list(cand)[0], None, 'name matches a CMS name') if len(cand) == 1 else (None, None, 'no CCN')
        cms_name = M['fac'].loc[ccn, 'name'] if ccn in M['fac'].index else ''
        if note is None and n in M['cfg'].get('ccn_notes_seed', {}) and n not in cec['existing_map']: note = M['cfg']['ccn_notes_seed'][n]
        rows.append(dict(name=n, ccn=ccn, cms_name=cms_name, how=how, note=note))
    return rows

MONTHS_ = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']

def closed_notes(mp, cfg):
    """CCN -> (label, first day of the month, year_only) for homes marked 'Closed <Month> <Year>' or 'Closed <Year>'
    in the CCN Map notes or cfg closed_seed. Used for homes CMS keeps listing after they close.
    With only a year, the date is December 1 of that year and year_only is True."""
    out = {}
    def take(ccn, text):
        t = str(text or '')
        m = re.match(r'\s*closed\s+([a-z]+)\s+(\d{4})', t, re.I)
        if ccn and m and m.group(1).lower() in MONTHS_:
            out[str(ccn)] = (f'Closed {m.group(1).title()} {m.group(2)}', f'{m.group(2)}-{MONTHS_.index(m.group(1).lower()) + 1:02d}-01', False)
            return
        m = re.match(r'\s*closed\s+(?:in\s+)?(\d{4})\b', t, re.I)
        if ccn and m:
            out[str(ccn)] = (f'Closed {m.group(1)}', f'{m.group(1)}-12-01', True)
    for c, t in (cfg.get('closed_seed') or {}).items(): take(c, t)
    for r in mp: take(r.get('ccn'), r.get('note'))
    return out

# ------------------------------------------------------------------ sale timing and outcomes
SALE_TYPES = ('Ownership/operator change', 'Corporate parent change (multi-facility deal)', 'Estate / internal ownership restructuring')
MET = dict(overall='overall_rating', hi='health_inspection_rating', staffing='staffing_rating', qm='qm_rating', hours='reported_total_nurse_staffing_hours_per_resident_per_day',
           rn='reported_rn_staffing_hours_per_resident_per_day', turnover='total_nursing_staff_turnover')

def T(x):
    if x is None or x is pd.NaT: return None
    try:
        if pd.isna(x): return None
    except Exception: pass
    return pd.Timestamp(x)

def compute_sales(M, cec, ccn_of):
    p = M['p']; latest = pd.Timestamp(M['latest']); first = pd.Timestamp(M['snaps'][0]); floor = pd.Timestamp(M['cfg']['series_start'])
    pf = p[p.snapshot_date >= M['cfg']['series_start']]   # only files comparable with today's rating method (same data the workbook's CMS Monthly Data tab holds)
    snaps_by = {c: pd.to_datetime(d.snapshot_date).values for c, d in pf.groupby('ccn_std')}
    val = {}
    for c, d in pf.groupby('ccn_std'):
        dd = d.set_index('snapshot_date')
        val[c] = {k: dd[col].to_dict() for k, col in MET.items()}
        val[c]['turnover'] = {s: (v / 100 if nz(v) is not None else None) for s, v in val[c]['turnover'].items()}
    log = cec['log'].set_index('cec'); rows = []
    for t in cec['tracker'].itertuples():
        lg = log.loc[t.cec] if t.cec in log.index else None
        ttype = lg['type'] if lg is not None else None; buyer = lg['buyer'] if lg is not None else None; applicant = lg['applicant'] if lg is not None else None
        chow = t.chow if nz(t.chow) is not None else None
        dec = t.decision if nz(t.decision) is not None else None
        evd, src = (chow, 'CMS takeover date') if nz(chow) is not None else ((dec, 'RIDOH decision date') if nz(dec) is not None else (None, ''))
        ovr = cec.get('overrides', {}).get(t.cec); treat = ovr['treat'] if ovr else None
        if treat in ('closure', 'not a sale'): ttype = 'Closure / not a sale'
        elif treat == 'sale': ttype = 'Ownership/operator change'
        elif treat == 'parent company change': ttype = 'Corporate parent change (multi-facility deal)'
        rows.append(dict(row=t.row, cec=t.cec, facility=t.facility, override=treat, ccn=ccn_of.get(t.facility), type=ttype, outcome=t.outcome, buyer=buyer, applicant=applicant or t.new_owner, new_owner=t.new_owner, evd=nz(evd), src=src, decision=nz(dec), chow=nz(chow)))
    S = pd.DataFrame(rows)
    counts = S.outcome.fillna('') != 'Pending'
    for i, r in S.iterrows():
        S.loc[i, 'has_data'] = bool(r.ccn in snaps_by) if r.ccn else False
    def prev_snap(ccn, d):
        a = snaps_by.get(ccn); 
        if a is None or d is None: return None
        m = a < np.datetime64(d)
        return pd.Timestamp(a[m][-1]) if m.any() else None
    def next_snap(ccn, d, tol=92):
        a = snaps_by.get(ccn)
        if a is None or d is None: return None
        m = a >= np.datetime64(d)
        if not m.any(): return None
        s = pd.Timestamp(a[m][0]); return s if (s - d).days <= tol else None
    out = []
    S['evd'] = pd.to_datetime(S['evd']); S['ccn'] = S['ccn'].where(S['ccn'].notna(), None)
    for i, r in S.iterrows():
        d = {k: None for k in ['before', 'y1', 'y2', 'exit', 'next_event']}
        ev = T(r.evd); ccn = nz(r.ccn)
        if r.has_data and ev is not None and (nz(r.outcome) or '') != 'Pending' and nz(r.override) not in ('closure', 'not a sale'):
            b = prev_snap(ccn, ev); d['before'] = b if (b is not None and b >= floor) else None
            others = S[(S.ccn == ccn) & (S.evd > ev + pd.Timedelta(days=1)) & (S.outcome.fillna('') != 'Pending') & S.type.isin(SALE_TYPES)]
            ne = T(others.evd.min()) if len(others) else None; d['next_event'] = ne
            for k, m in (('y1', 12), ('y2', 24)):
                s2 = next_snap(ccn, ev + pd.DateOffset(months=m))
                d[k] = s2 if (s2 is not None and (ne is None or s2 < ne)) else None
            a = snaps_by[ccn]; mm = a < np.datetime64(ne) if ne is not None else np.ones(len(a), bool)
            d['exit'] = pd.Timestamp(a[mm][-1]) if mm.any() and d['before'] is not None else None
        out.append(d)
    for k in ['before', 'y1', 'y2', 'exit', 'next_event']: S[k] = pd.to_datetime(pd.Series([o[k] for o in out], index=S.index))
    # metric values at those snapshots
    for pre in ('before', 'y1', 'y2', 'exit'):
        for m in MET:
            S[f'{m}_{pre}'] = [(val[r.ccn][m].get(r[pre].strftime('%Y-%m-%d')) if (nz(r.ccn) is not None and T(r[pre]) is not None) else None) for _, r in S.iterrows()]
    hu, pu = M['hu'], M['pu']; hu = hu.assign(sd=pd.to_datetime(hu.survey_date)); pu = pu[pu.penalty_type == 'Fine'].assign(pd_=pd.to_datetime(pu[pu.penalty_type == 'Fine'].penalty_date))
    horizon = first - pd.DateOffset(years=3)
    def window(r, k):
        ev = T(r.evd)
        if ev is None or not r.has_data: return None
        if k == 'before':
            return (ev - pd.DateOffset(months=12), ev) if (T(r['before']) is not None and ev - pd.DateOffset(months=12) >= horizon) else None
        if k == 'y1': return (ev, ev + pd.DateOffset(months=12)) if T(r['y1']) is not None else None
        if k == 'y2': return (ev + pd.DateOffset(months=12), ev + pd.DateOffset(months=24)) if T(r['y2']) is not None else None
    for k in ('before', 'y1', 'y2'):
        ca, cs, fa = [], [], []
        for _, r in S.iterrows():
            w = window(r, k)
            if w is None: ca.append(None); cs.append(None); fa.append(None); continue
            h2 = hu[(hu.ccn_std == r.ccn) & (hu.sd >= w[0]) & (hu.sd < w[1])]; f = pu[(pu.ccn_std == r.ccn) & (pu.pd_ >= w[0]) & (pu.pd_ < w[1])]
            ca.append(len(h2)); cs.append(int(h2.serious.sum())); fa.append(float(f.fine_amount.sum()))
        S[f'cit_{k}'], S[f'ser_{k}'], S[f'fines_{k}'] = ca, cs, fa
    # first CMS-detected operator change at/after (event - 6 months)
    strong = M['ev_own'][M['ev_own'].category.isin(['Operator change', 'Change of ownership (CMS flag)'])].copy(); strong['d'] = pd.to_datetime(strong.snap)
    det, lag = [], []
    for _, r in S.iterrows():
        if T(r.evd) is None or nz(r.ccn) is None or nz(r.override) in ('closure', 'not a sale'): det.append(None); lag.append(None); continue
        s2 = strong[(strong.ccn == r.ccn) & (strong.category == 'Operator change') & (strong.d >= r.evd - pd.DateOffset(months=6))]
        det.append(s2.d.min() if len(s2) else None); lag.append((int((s2.d.min() - r.evd).days) if (len(s2) and (s2.d.min() - r.evd).days <= 730) else None))
    S['cms_detected'], S['cms_lag'] = pd.to_datetime(pd.Series(det, index=S.index)), lag
    S['pre_method'] = False
    S['counts'] = counts.values & S.type.isin(SALE_TYPES).values
    return S, val, snaps_by

# ------------------------------------------------------------------ names, cohort analysis, public data
def former_names(M, cec, ccn_of):
    """Per CCN: list of other names (CMS listings over time + the CEC workbook's standard/other names)."""
    fac, nh = M['fac'], M['names_hist']; out = {}
    cec_names = {}
    for r in cec['fs'].itertuples():
        c = ccn_of.get(r.name) or ('nocc:' + r.name)
        lst = [(r.name, 'CEC standard name')] + [(x.strip(), 'CEC record') for x in str(nz(r.other) or '').split(';') if x.strip()]
        cec_names.setdefault(c, []).extend(lst)
    keys_by = {}
    for c in fac.index:
        disp = key_(fac.loc[c, 'name']); seen = {disp}; lst = []
        for x in reversed(nh[c]):
            if x['key'] in seen: continue
            seen.add(x['key']); lst.append(dict(name=x['name'], src='CMS listing', first=x['first'], last=x['last']))
        for n, src in cec_names.get(c, []):
            if key_(n) in seen: continue
            seen.add(key_(n)); lst.append(dict(name=n, src=src, first=None, last=None))
        out[c] = lst
    for c, items in cec_names.items():   # CEC facilities without CMS data
        if c not in fac.index:
            seen, lst = set(), []
            for n, src in items:
                if key_(n) in seen: continue
                seen.add(key_(n)); lst.append(dict(name=n, src=src, first=None, last=None))
            out[c] = lst
    return out
def names_text(lst, first_is_current=False):
    parts = []
    for x in lst:
        if x['src'] == 'CMS listing' and x['last']: parts.append(f"{x['name']} (CMS name through {lab_(x['last'])})")
        elif x['src'] == 'CEC standard name': parts.append(f"{x['name']} (name used in the CEC record)")
        else: parts.append(x['name'])
    return '; '.join(parts)

def event_study(M, S, val, metric, sale_rows, qrange=range(-4, 9), tol=45):
    """Average change from the last snapshot before each sale, for sold homes vs comparable homes that were not sold."""
    fac = M['fac']; snaps = pd.to_datetime(pd.Series(M['snaps'])); latest = pd.Timestamp(M['latest'])
    evd_by = {}
    for r in S[S.counts].itertuples():
        if T(r.evd) is not None and nz(r.ccn): evd_by.setdefault(r.ccn, []).append(T(r.evd))
    for r in M['ev_own'][M['ev_own'].category.isin(['Operator change', 'Change of ownership (CMS flag)'])].itertuples(): evd_by.setdefault(r.ccn, []).append(pd.Timestamp(r.snap))
    def near(t):
        i = (snaps - t).abs().idxmin(); return snaps[i].strftime('%Y-%m-%d') if abs((snaps[i] - t).days) <= tol else None
    sold = {q: [] for q in qrange}; ctrl = {q: [] for q in qrange}; per_event = []
    for r in sale_rows.itertuples():
        ccn, b = r.ccn, T(r.before)
        if b is None or b < pd.Timestamp(M['cfg']['series_start']): continue
        v0 = val[ccn][metric].get(b.strftime('%Y-%m-%d'))
        if nz(v0) is None: continue
        ne = T(r.next_event)
        controls = [c for c in fac.index if c != ccn and not any((b - pd.DateOffset(months=12)) <= d <= (b + pd.DateOffset(months=24)) for d in evd_by.get(c, []))]
        for q in qrange:
            tgt = b + pd.DateOffset(months=3 * q)
            if tgt > latest + pd.Timedelta(days=tol): continue
            s = near(tgt)
            if s is None: continue
            if q > 0 and ne is not None and pd.Timestamp(s) >= ne: continue
            v = val[ccn][metric].get(s)
            if nz(v) is None: continue
            cd = [val[c][metric].get(s) - val[c][metric].get(b.strftime('%Y-%m-%d')) for c in controls
                  if nz(val[c][metric].get(s)) is not None and nz(val[c][metric].get(b.strftime('%Y-%m-%d'))) is not None]
            if len(cd) < 5: continue
            sold[q].append(v - v0); ctrl[q].append(float(np.mean(cd)))
    res = []
    for q in qrange:
        res.append(dict(q=q, months=3 * q, sold=(round(float(np.mean(sold[q])), 3) if sold[q] else None), ctrl=(round(float(np.mean(ctrl[q])), 3) if ctrl[q] else None), n=len(sold[q]),
                        worse=None, same=None, better=None))
    return res, sold, ctrl

def cohort_summary(sold, ctrl, metric, higher_better=True):
    out = {}
    for q, label in ((4, '12'), (8, '24')):
        d = sold.get(q, []); c = ctrl.get(q, [])
        if not d: out[label] = None; continue
        sgn = 1 if higher_better else -1
        out[label] = dict(n=len(d), sold=round(float(np.mean(d)), 2), ctrl=round(float(np.mean(c)), 2), worse=int(sum(1 for x in d if sgn * x < -1e-9)), same=int(sum(1 for x in d if abs(x) <= 1e-9)), better=int(sum(1 for x in d if sgn * x > 1e-9)))
    return out


# ------------------------------------------------------------------ money: related-party costs around a sale
MONEY = dict(rp='Related-party costs', rent='Rent and building costs paid to a related party', mgmt='Management and admin fees paid to a related party', margin='Operating margin')
def cost_panel(cec, S):
    """Per-home cost-report shares of net patient revenue, and same-home before/after comparisons around the first sale."""
    c = cec.get('costs')
    if c is None or not len(c): return None
    c = c.copy()
    for k in ('rp', 'rent', 'mgmt'): c[k + '_sh'] = c[k] / c['rev']
    c['margin_sh'] = c['margin']
    by = {ccn: g.set_index('year') for ccn, g in c.groupby('ccn')}
    sold = S[S.counts & S.evd.notna() & S.ccn.notna()].sort_values('evd')
    first = sold.groupby('ccn').first().reset_index()[['ccn', 'evd', 'cec', 'facility']]
    sold_ccns = set(first.ccn)
    out = dict(years=sorted(int(y) for y in c.year.dropna().unique()), by_facility={}, compare={}, n_homes=0)
    for ccn, g in by.items():
        out['by_facility'][ccn] = dict(years=[int(y) for y in g.index], rev=[_r(v, 0) for v in g['rev']], rp=[_r(v, 0) for v in g['rp']], rent=[_r(v, 0) for v in g['rent']], mgmt=[_r(v, 0) for v in g['mgmt']],
                                        rp_sh=[_r(v, 4) for v in g['rp_sh']], rent_sh=[_r(v, 4) for v in g['rent_sh']], mgmt_sh=[_r(v, 4) for v in g['mgmt_sh']], margin=[_r(v, 4) for v in g['margin_sh']],
                                        parties=[(None if nz(v) is None else str(v)[:400]) for v in g['parties']])
    # same-home before/after (year before the sale vs 1, 2, 3 years after)
    for k in ('rp_sh', 'rent_sh', 'mgmt_sh', 'margin'):
        res = {}
        for after in (1, 2, 3):
            pairs = []
            for r in first.itertuples():
                g = by.get(r.ccn); sy = T(r.evd).year
                if g is None: continue
                a, b = sy - 1, sy + after
                if a in g.index and b in g.index and nz(g.loc[a, k]) is not None and nz(g.loc[b, k]) is not None and not (pd.isna(g.loc[a, k]) or pd.isna(g.loc[b, k])):
                    pairs.append((float(g.loc[a, k]), float(g.loc[b, k]), r.ccn))
            if len(pairs) >= 5:
                res[str(after)] = dict(n=len(pairs), before=round(float(np.mean([x[0] for x in pairs])), 4), after=round(float(np.mean([x[1] for x in pairs])), 4), up=int(sum(1 for x in pairs if x[1] > x[0])))
        out['compare'][k] = res
    # homes never sold, same calendar years, as a comparison
    ctrl = {}
    for k in ('rp_sh', 'rent_sh', 'mgmt_sh', 'margin'):
        rows = c[~c.ccn.isin(sold_ccns)].groupby('year')[k].mean()
        ctrl[k] = {int(y): _r(v, 4) for y, v in rows.items() if nz(v) is not None}
    out['never_sold'] = ctrl; out['n_homes'] = len(first); out['last_year'] = (max([y for y in out['years'] if int((c.year == y).sum()) >= 20], default=(max(out['years']) if out['years'] else None))); out['n_facilities'] = len(by)
    out['first_sale'] = {r.ccn: dict(year=T(r.evd).year, date=_d(r.evd), cec=r.cec) for r in first.itertuples()}
    sale_years = {}
    for r in sold.itertuples(): sale_years.setdefault(r.ccn, set()).add(T(r.evd).year)
    out['trend'] = _cost_trend(c, sale_years)
    return out

def _cost_trend(c, sale_years):
    """Balanced panel: the homes that filed a cost report in EVERY well-covered year, split by whether
    they changed hands during the window. Shares are dollars summed across the group over revenue summed
    across the same group, so one small home cannot swing the line."""
    cover = c.groupby('year').ccn.nunique()
    if not len(cover): return None
    floor = max(20, 0.75 * float(cover.max()))                   # a year where most homes have filed yet; the newest year is always thin
    ok = sorted(int(y) for y in cover[cover >= floor].index)
    if len(ok) < 4: return None
    yrs, last = [ok[-1]], ok[-1]                                 # the contiguous run ending at the newest well-covered year
    while last - 1 in ok: last -= 1; yrs.insert(0, last)
    if len(yrs) < 4: return None
    sets = [set(c.loc[c.year == y, 'ccn']) for y in yrs]
    core = set.intersection(*sets)
    if len(core) < 20: return None
    win = set(yrs)                                               # "changed hands" means during the window shown, not at any point since 2007
    a = {x for x in core if sale_years.get(x) and (sale_years[x] & win)}
    b = core - a
    if len(a) < 5 or len(b) < 5: return None
    def series(grp, col):
        vals = []
        for y in yrs:
            x = c[(c.year == y) & (c.ccn.isin(grp))]
            rev = float(x['rev'].sum())
            vals.append(_r(float(x[col].sum()) / rev * 100, 2) if rev > 0 else None)
        return vals
    def dollars(grp, col):
        return [_r(float(c.loc[(c.year == y) & (c.ccn.isin(grp)), col].sum()), 0) for y in yrs]
    return dict(years=yrs, n_sold=len(a), n_other=len(b),
                rent=dict(sold=series(a, 'rent'), other=series(b, 'rent')),
                rp=dict(sold=series(a, 'rp'), other=series(b, 'rp')),
                rent_usd=dict(sold=dollars(a, 'rent'), other=dollars(b, 'rent')))

# ------------------------------------------------------------------ public page data
TYPE_LABEL = {'Ownership/operator change': 'Sold / new operator', 'Corporate parent change (multi-facility deal)': 'Chain reorganization', 'Estate / internal ownership restructuring': 'Estate / internal restructuring',
              'Real estate only': 'Real estate only', 'Unclear from report': 'Unclear from report'}
def _r(v, n=2):
    v = nz(v); return None if v is None else round(float(v), n)
def _d(x): 
    x = T(x); return None if x is None else x.strftime('%Y-%m-%d')

def build_public_data(M, S, val, cec, mp, ccn_of, fn, cfg):
    def pubtxt(x, applicant=False):
        x = nz(x)
        if x is None: return None
        x = re.sub(r'\s+', ' ', str(x)).strip()
        if cfg.get('public_show_principals'): return x.replace(' \u2192 ', ' \u2192 ')
        if applicant: x = x.split('\u2192')[0]
        x = re.sub(r'\s*\([^)]*\)', '', x)
        return x.strip(' -\u2014') or None
    p, fac, latest = M['p'], M['fac'], M['latest']; dates = [d for d in M['snaps'] if d >= cfg['series_start']]
    cec_ccn = ccn_of; log = cec['log']
    # statewide monthly series
    g = p.groupby('snapshot_date'); st = dict(n=[], nr=[], avg_overall=[], dist=[], avg_hours=[], avg_rn=[], avg_turnover=[], avg_staffing=[], avg_hi=[], avg_qm=[])
    for d in dates:
        x = g.get_group(d); ov = x.overall_rating.dropna()
        st['n'].append(int(len(x))); st['avg_overall'].append(_r(ov.mean(), 3))
        st['nr'].append(int((x['overall_rating_footnote'].astype(str).str.replace('.0', '', regex=False) == '18').sum()) if 'overall_rating_footnote' in x.columns else 0)
        st['dist'].append([int((ov == k).sum()) for k in range(1, 6)])
        st['avg_hours'].append(_r(x.reported_total_nurse_staffing_hours_per_resident_per_day.mean(), 3)); st['avg_rn'].append(_r(x.reported_rn_staffing_hours_per_resident_per_day.mean(), 3))
        tv = x.total_nursing_staff_turnover.dropna(); st['avg_turnover'].append(_r(tv.mean(), 1) if len(tv) >= 20 else None)
        st['avg_staffing'].append(_r(x.staffing_rating.mean(), 3)); st['avg_hi'].append(_r(x.health_inspection_rating.mean(), 3)); st['avg_qm'].append(_r(x.qm_rating.mean(), 3))
    years = M['years']; CQ = M['CQ']
    sy = dict(years=years, serious=[int(CQ['cit_ser'][[q for q in M['qpairs'] if q[0] == y]].sum().sum()) for y in years],
              citations=[int(CQ['cit_all'][[q for q in M['qpairs'] if q[0] == y]].sum().sum()) for y in years],
              fines=[int(round(CQ['fines_amt'][[q for q in M['qpairs'] if q[0] == y]].sum().sum())) for y in years], last_year_quarters=M['last_q'])
    # facilities
    piv = {c: d.set_index('snapshot_date') for c, d in p.groupby('ccn_std')}
    events_by = {}
    def ev_add(key, e): events_by.setdefault(key, []).append(e)
    OVR = cec.get('overrides', {})
    for r in log.itertuples():
        d = T(r.chow) or T(r.decision); ovr = OVR.get(r.cec)
        if ovr and ovr.get('date'): d = T(ovr['date'])
        if d is None: continue
        key = cec_ccn.get(r.facility) or ('cec:' + str(r.facility))
        rtype = r.type
        if ovr and ovr['treat'] == 'sale': rtype = 'Ownership/operator change'
        if ovr and ovr['treat'] == 'parent company change': rtype = 'Corporate parent change (multi-facility deal)'
        if ovr and ovr['treat'] in ('closure', 'not a sale'):
            ev_add(key, dict(date=_d(d), kind='closed', cec=r.cec, label=nz(ovr.get('label')) or 'Closed (not a sale)', who='', buyer=None, outcome=None, src='Note from the workbook owner (state review record ' + r.cec + ' exists)', date_note='Date from the workbook note'))
            continue
        kind = 'sale' if rtype in ('Ownership/operator change',) else ('parent' if rtype == 'Corporate parent change (multi-facility deal)' else ('internal' if rtype == 'Estate / internal ownership restructuring' else 'realestate'))
        who = pubtxt(r.applicant, True) or ''
        ev_add(key, dict(date=_d(d), kind=kind, cec=r.cec, label=TYPE_LABEL.get(rtype, rtype or 'CEC'), who=who, buyer=pubtxt(r.buyer), outcome=nz(r.outcome), src='RIDOH Change in Effective Control record ' + r.cec,
                         date_note=('CMS takeover date' if T(r.chow) else 'RIDOH decision date')))
    for r in M['ev_own'][M['ev_own'].category.isin(['Operator change', 'Change of ownership (CMS flag)'])].itertuples():
        ev_add(r.ccn, dict(date=r.snap, kind='cms', label='CMS shows a new operator' if r.category == 'Operator change' else 'CMS flags a change of ownership', who=(f'{r.before} \u2192 {r.after}' if r.before else ''), buyer=None, outcome=None, src='CMS Care Compare files (first monthly file showing the change)', date_note='First CMS file showing it'))
    for r in M['ev_other'][M['ev_other'].category == 'Left CMS data'].itertuples():
        if not any(e['kind'] == 'closed' for e in events_by.get(r.ccn, [])):
            ev_add(r.ccn, dict(date=r.snap, kind='cms', label='No longer in the CMS files', who='closed, merged or decertified; the data does not say which', buyer=None, outcome=None, src='CMS Care Compare files (first monthly file without the home)', date_note='First CMS file without the home'))
    closed_map = closed_notes(mp, cfg)
    for c_, (lab_c, d_c, yo_c) in closed_map.items():
        if yo_c: continue   # no month known, so no dated timeline event; the status still says "Closed <Year>"
        if c_ in fac.index and not any(e['kind'] == 'closed' for e in events_by.get(c_, [])):
            ev_add(c_, dict(date=d_c, kind='closed', cec='', label=lab_c, who='', buyer=None, outcome=None, src='Note from the workbook owner', date_note='Month from the workbook note'))
    hu, pu = M['hu'], M['pu']
    facs = []; cur = fac
    for c in M['order']:
        row = fac.loc[c]; d = piv[c]; lr = d.iloc[-1]
        if c in closed_map:   # a closed home CMS still lists: show it as of the last CMS file before it closed
            upto = pd.Timestamp(closed_map[c][1]) + pd.offsets.MonthEnd(0); d_b = d[pd.to_datetime(d.index) <= upto]
            if closed_map[c][2]:   # year only: the last file that year that still reported residents
                d_r = d_b[pd.to_numeric(d_b['average_number_of_residents_per_day'], errors='coerce').notna()]
                if len(d_r): d_b = d_r
            if len(d_b): lr = d_b.iloc[-1]
        def ser(col, sc=1.0, nd=2): return [(_r(d[col].get(s) * sc, nd) if nz(d[col].get(s)) is not None else None) for s in dates]
        yrs = M['years']
        def ann(df): return [int(sum(df.loc[c, q] for q in M['qpairs'] if q[0] == y and nz(df.loc[c, q]) is not None)) for y in yrs]
        rec = dict(id=c, ccn=c, name=row['name'], town=row['city'], county=row['county'], active=bool(row['active']), status=row['status'], last_seen=row['last_seen'],
                   beds=_r(lr.number_of_certified_beds, 0), owner=row['operator'], chain=nz(row['chain']), ownership_type=nz(lr.ownership_type), sff=nz(lr.special_focus_status), abuse=(str(nz(lr.abuse_icon)).upper() == 'Y'),
                   cms_name=row['cms_name'], former=[dict(name=x['name'], src=x['src'], to=x['last']) for x in fn.get(c, [])],
                   now=dict(overall=_r(lr.overall_rating, 0), hi=_r(lr.health_inspection_rating, 0), staffing=_r(lr.staffing_rating, 0), qm=_r(lr.qm_rating, 0), hours=_r(lr.reported_total_nurse_staffing_hours_per_resident_per_day, 2),
                            rn=_r(lr.reported_rn_staffing_hours_per_resident_per_day, 2), turnover=_r(lr.total_nursing_staff_turnover, 0), asof=(str(pd.Timestamp(lr.name).date()) if c in closed_map else row['last_seen'])),
                   s=dict(overall=ser('overall_rating', 1, 0), hi=ser('health_inspection_rating', 1, 0), staffing=ser('staffing_rating', 1, 0), qm=ser('qm_rating', 1, 0), hours=ser('reported_total_nurse_staffing_hours_per_resident_per_day', 1, 2),
                          rn=ser('reported_rn_staffing_hours_per_resident_per_day', 1, 2), turnover=ser('total_nursing_staff_turnover', 1, 0),
                          tf=[(1 if (s_ in d.index and bool(d['turn_flag'].get(s_))) else 0) for s_ in dates],
                          hf=[(1 if (s_ in d.index and bool(d['hprd_flag'].get(s_))) else 0) for s_ in dates],
                          nr=[(1 if ('overall_rating_footnote' in d.columns and str(d['overall_rating_footnote'].get(s_)).replace('.0', '') == '18') else 0) for s_ in dates]),
                   yr=dict(serious=ann(CQ['cit_ser']), citations=ann(CQ['cit_all']), fines=[int(round(x)) for x in ann(CQ['fines_amt'])]),
                   events=sorted(events_by.get(c, []), key=lambda e: e['date'], reverse=True), has_cms=True)
        facs.append(rec)
    # CEC-only facilities (no CCN or no CMS data since 2019)
    for r in mp:
        if r['ccn'] and r['ccn'] in fac.index: continue
        key = r['ccn'] or ('cec:' + r['name'])
        fs = cec['fs'][cec['fs'].name == r['name']]; town = nz(fs.iloc[0].town) if len(fs) else None
        m_ = re.match(r'\s*(Closed[^.;(]*)', str(r.get('note') or '')); stat = m_.group(1).strip() if m_ else 'No CMS data in this dataset (not in the CMS files since 2019)'
        facs.append(dict(id=key, ccn=r['ccn'], name=r['name'], town=town, county=None, active=None, status=stat, last_seen=None, beds=None, owner=None, chain=None, ownership_type=None, sff=None, abuse=False,
                         cms_name='', former=[dict(name=x['name'], src=x['src'], to=None) for x in fn.get(r['ccn'] or ('nocc:' + r['name']), []) if key_(x['name']) != key_(r['name'])], now=None, s=None, yr=None,
                         events=sorted(events_by.get(key, []), key=lambda e: e['date'], reverse=True), has_cms=False))
    for cid, o in OVR.items():
        if o['treat'] in ('closure', 'not a sale') and nz(o.get('label')):
            lr = log[log.cec == cid]
            if len(lr):
                fid_ = cec_ccn.get(lr.iloc[0].facility) or ('cec:' + str(lr.iloc[0].facility)); f_ = next((x for x in facs if x['id'] == fid_), None)
                if f_ is not None and o['treat'] == 'closure': f_['status'] = o['label']
    facs = [f for f in facs if f['has_cms'] or f['events']]
    # sales table
    rows = []; pub = S[S.counts & S.type.isin(cfg['public_include_types'])].sort_values('evd', ascending=False)
    fid = lambda r: (r.ccn if nz(r.ccn) and r.ccn in fac.index else 'cec:' + str(r.facility))
    for r in pub.itertuples():
        f = next((x for x in facs if x['id'] == fid(r)), None)
        rows.append(dict(cec=r.cec, fid=fid(r), date=_d(r.evd), date_note=r.src, name=(f['name'] if f else r.facility), town=(f['town'] if f else None), kind=('sale' if r.type == 'Ownership/operator change' else 'parent'),
                         label=TYPE_LABEL.get(r.type, r.type), buyer=pubtxt(r.buyer), who=pubtxt(r.applicant, True) or '', outcome=nz(r.outcome), has_cms=bool(r.has_data), before_date=_d(r.before), cms_op=_d(r.cms_detected),
                         m={k: dict(before=_r(getattr(r, f'{k}_before'), 2), y1=_r(getattr(r, f'{k}_y1'), 2), y2=_r(getattr(r, f'{k}_y2'), 2), exit=_r(getattr(r, f'{k}_exit'), 2)) for k in ('overall', 'staffing', 'hours', 'turnover')},
                         ser=dict(before=nz(r.ser_before), y1=nz(r.ser_y1), y2=nz(r.ser_y2)), fines=dict(before=nz(r.fines_before), y1=nz(r.fines_y1), y2=nz(r.fines_y2))))
    rows = [r for r in rows if r['date']]          # a sale with no recorded date is left out of the table
    for r in rows:
        for k in ('ser', 'fines'):
            r[k] = {a: (None if nz(b) is None else round(float(b))) for a, b in r[k].items()}
    money = cost_panel(cec, S)
    MG = cec.get('mgmt', {})
    for cid, m in MG.items():
        if not T(m.get('start')): continue
        lr = log[log.cec == cid]
        if not len(lr): continue
        key = cec_ccn.get(lr.iloc[0].facility) or ('cec:' + str(lr.iloc[0].facility))
        ev_add(key, dict(date=_d(m['start']), kind='mgmt', label='Buyer took over management', who=pubtxt(m.get('who')) or '', buyer=None, outcome=None, src=nz(m.get('src')) or ('Management agreement noted for ' + cid), date_note='Management agreement start'))
    for f_ in facs: f_['events'] = sorted(events_by.get(f_['id'], f_['events']), key=lambda e: e['date'], reverse=True)
    for r in rows: r['mgmt_start'] = (_d(MG[r['cec']]['start']) if (r['cec'] in MG and T(MG[r['cec']].get('start'))) else None)
    # same-home levels before vs after (for the before/after chart)
    def levels(metric, col):
        out = {}
        base = S[S.counts & S.type.isin(cfg['public_include_types']) & S.before.notna()]
        for k, lbl in (('y1', '1'), ('y2', '2')):
            pr = [(getattr(r, f'{metric}_before'), getattr(r, f'{metric}_{k}')) for r in base.itertuples() if nz(getattr(r, f'{metric}_before')) is not None and nz(getattr(r, f'{metric}_{k}')) is not None]
            if len(pr) >= 5: out[lbl] = dict(n=len(pr), before=round(float(np.mean([x[0] for x in pr])), 3), after=round(float(np.mean([x[1] for x in pr])), 3), worse=int(sum(1 for x in pr if x[1] < x[0])))
        return out
    lev = {m: levels(m, None) for m in ('overall', 'staffing', 'hours', 'rn', 'turnover')}
    # event studies
    study = {}
    for setname, types in (('sold', ('Ownership/operator change',)), ('all', tuple(cfg['public_include_types']))):
        sales = S[S.counts & S.type.isin(types) & S.before.notna() & ~S.pre_method]; study[setname] = dict(n_events=int(len(sales)), n_homes=int(sales.ccn.nunique()), metrics={})
        for m, hb in (('overall', True), ('staffing', True), ('hours', True), ('turnover', False)):
            res, sold, ctrl = event_study(M, S, val, m, sales); study[setname]['metrics'][m] = dict(points=res, summary=cohort_summary(sold, ctrl, m, hb))
    # owner groups: stars at acquisition vs 1 year, 2 years and latest (each comparison uses only homes with data at that point)
    grp = {}
    for r in S[S.counts & S.type.isin(cfg['public_include_types']) & S.before.notna()].itertuples():
        grp.setdefault(pubtxt(r.buyer) or 'Other / not stated', []).append(r)
    groups = []
    for g, rs in grp.items():
        homes = []
        for r in sorted(rs, key=lambda x: x.evd):
            f_ = fac.loc[r.ccn, 'name'] if r.ccn in fac.index else r.facility
            yrs = round((T(r.exit) - T(r.evd)).days / 365.25, 1) if T(r.exit) is not None else None
            homes.append(dict(name=f_, date=_d(r.evd), before=_r(r.overall_before, 0), y1=_r(r.overall_y1, 0), y2=_r(r.overall_y2, 0), latest=_r(r.overall_exit, 0), yrs=yrs))
        def pair(k):
            sub_ = [h for h in homes if h['before'] is not None and h[k] is not None]
            return dict(n=len(sub_), before=round(float(np.mean([h['before'] for h in sub_])), 2), after=round(float(np.mean([h[k] for h in sub_])), 2)) if sub_ else None
        ys = [h['yrs'] for h in homes if h['yrs'] is not None and h['latest'] is not None]
        groups.append(dict(group=g, n=len(homes), y1=pair('y1'), y2=pair('y2'), latest=pair('latest'), yrs=(round(float(np.median(ys)), 1) if ys else None), homes=homes))
    groups.sort(key=lambda x: -x['n'])
    # headline numbers
    ov0, ov1 = st['dist'][0], st['dist'][-1]; low = lambda dd, nr: (nr + dd[0] + dd[1]) / max(1, sum(dd) + nr)
    first_i = 0
    head = dict(n_now=st['n'][-1], n_then=st['n'][0], low_now=round(low(ov1, st['nr'][-1]) * 100), low_then=round(low(ov0, st['nr'][0]) * 100), avg_now=st['avg_overall'][-1], avg_then=st['avg_overall'][0],
                hours_now=st['avg_hours'][-1], hours_then=st['avg_hours'][0], rn_now=st['avg_rn'][-1], rn_then=st['avg_rn'][0], turnover_now=st['avg_turnover'][-1],
                sales_since=int(len([r for r in rows if r['date'] and r['date'] >= dates[0] and r['outcome'] != 'Pending'])), homes_left=int((fac['last_seen'] != fac['last_seen'].max()).sum()))   # left the CMS files (a closed home CMS still lists is not counted here)
    pyr = {}
    for r in rows:
        if r['date'] and r['date'] >= dates[0] and r['outcome'] != 'Pending': pyr[int(r['date'][:4])] = pyr.get(int(r['date'][:4]), 0) + 1
    yrs_all = list(range(int(dates[0][:4]), int(latest[:4]) + 1))
    head['per_year'] = [dict(year=y, n=pyr.get(y, 0)) for y in yrs_all]
    tv0 = next(((d, v) for d, v in zip(dates, st['avg_turnover']) if v is not None), None)
    head['turnover_first'] = (dict(date=tv0[0], value=tv0[1]) if tv0 else None)
    meta = dict(behind={k: lab_(v) for k, v in M['ds_behind'].items()}, title=cfg['page_title'], credit=cfg['page_credit'], contact=cfg['page_contact'], latest=latest, first=dates[0], updated=lab_(latest), built=dt.date.today().isoformat(), n_homes_ever=int(len(fac)))
    return dict(meta=meta, dates=dates, state=st, state_years=sy, facilities=facs, sales=rows, study=study, groups=groups, head=head, money=money, levels=lev)

# ------------------------------------------------------------------ workbook: CMS tabs
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.chart import LineChart, BarChart, Reference

MANAGED = ['CMS Read Me', 'CMS Summary', 'RI Trends (Annual)', 'RI Trends (Quarterly)', 'Ownership Changes', 'Operator by Quarter', 'Other Changes', 'Overall Rating', 'Health Inspection', 'Staffing Rating',
           'QM Rating', 'Nurse Hours', 'RN Hours', 'Turnover', 'Occupancy', 'Citations - All', 'Citations - Serious', 'Citations - Abuse', 'Fines', 'Sale Timing', 'CMS Citations', 'CMS Penalties', 'CMS Monthly Data', 'Grid Flags']
FONT = 'Arial'
F_NORM = Font(name=FONT, size=10); F_BOLD = Font(name=FONT, size=10, bold=True); F_TITLE = Font(name=FONT, size=14, bold=True, color='1F3864'); F_NOTE = Font(name=FONT, size=9, italic=True, color='595959')
F_HDR = Font(name=FONT, size=10, bold=True, color='FFFFFF'); F_GREY = Font(name=FONT, size=10, color='7F7F7F'); F_INPUT = Font(name=FONT, size=10, color='0000FF')
FILL_ANN = PatternFill('solid', start_color='1F3864'); FILL_Q = PatternFill('solid', start_color='2E75B6'); FILL_GRP = PatternFill('solid', start_color='D9E1F2'); FILL_STAT = PatternFill('solid', start_color='F2F2F2'); FILL_INPUT = PatternFill('solid', start_color='FFF2CC')
TOP = Border(top=Side(style='medium', color='1F3864')); CTR = Alignment(horizontal='center'); CENTER_W = Alignment(horizontal='center', vertical='center', wrap_text=True)
CH = '+0.00;-0.00;0.00'; CHI = '+0;-0;0'

def _val(v):
    if v is None or (isinstance(v, float) and np.isnan(v)): return None
    if isinstance(v, np.integer): return int(v)
    if isinstance(v, np.floating): return float(v)
    if isinstance(v, pd.Timestamp): return v.to_pydatetime()
    return v
def _put(ws, r, c, v, font=None, fmt=None, fill=None, align=None, border=None):
    cell = ws.cell(row=r, column=c, value=_val(v)); cell.font = font or F_NORM
    if fmt: cell.number_format = fmt
    if fill: cell.fill = fill
    if align: cell.alignment = align
    if border: cell.border = border
    return cell
def _title(ws, t, note): _put(ws, 1, 1, t, F_TITLE); _put(ws, 2, 1, note, F_NOTE); ws.sheet_view.showGridLines = False
def _hdr(ws, r, c, text, annual=True): _put(ws, r, c, text, F_HDR, fill=FILL_ANN if annual else FILL_Q, align=CENTER_W)
STAR = ColorScaleRule(start_type='num', start_value=1, start_color='F8696B', mid_type='num', mid_value=3, mid_color='FFEB84', end_type='num', end_value=5, end_color='63BE7B')
def _scale(): return ColorScaleRule(start_type='percentile', start_value=5, start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='percentile', end_value=95, end_color='63BE7B')
def _rev(): return ColorScaleRule(start_type='percentile', start_value=5, start_color='63BE7B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='percentile', end_value=95, end_color='F8696B')

def write_cms_tabs(wb, M, S, cec, mp, ccn_of, fn, cfg):
    fac, order, qpts, is_ann, p = M['fac'], M['order'], M['qpts'], M['is_ann'], M['p']; NQ = len(qpts); N = len(order); HDR, FIRST = 4, 5; LAST = FIRST + N - 1
    years, qpairs, CQ, latest = M['years'], M['qpairs'], M['CQ'], M['latest']; lastq = M['last_q']
    plabels = [lab_(s) for s in qpts]
    ann_idx = [i for i in range(NQ) if is_ann[i]]
    if ann_idx[-1] != NQ - 1: ann_idx = ann_idx + [NQ - 1]
    alabels = [plabels[i] for i in ann_idx]
    ylabels = [str(y) for y in years[:-1]] + [str(years[-1]) if lastq == 4 else f"{years[-1]} (to {pd.Timestamp(latest).strftime('%b')})"]
    SR = {}
    LEG = 'Dark headers = annual snapshot; lighter = quarterly. Quarterly columns are shown; click the small 1 button at the top-left of the sheet (or the - above the columns) to collapse to annual only, and 2 to expand again.'
    note_std = 'Snapshots are CMS monthly files nearest each quarter (see CMS Read Me). 0 = CMS did not rate the home because it is in the Special Focus Facility (SFF) program; it is counted as 0 stars in averages and counts. Blank = not in the CMS file that month, or no value reported. Grey facilities are no longer in the CMS data. ' + LEG
    def left_cols(ws):
        for i, h in enumerate(['Facility', 'CCN', 'City', 'Status']): _hdr(ws, HDR, i + 1, h)
        for i, ccn in enumerate(order):
            r = FIRST + i; row = fac.loc[ccn]; g = F_GREY if not row['active'] else F_NORM
            _put(ws, r, 1, row['name'], g); _put(ws, r, 2, str(ccn), g, '@'); _put(ws, r, 3, row['city'], g); _put(ws, r, 4, row['status'], g)
        for c, w in zip('ABCD', [46, 9, 17, 30]): ws.column_dimensions[c].width = w
    def data_block(ws, c0, df, fmt, width=11):
        for i in range(N):
            for j in range(df.shape[1]): _put(ws, FIRST + i, c0 + j, df.iloc[i, j], fmt=fmt, align=CTR)
        for j in range(df.shape[1]): ws.column_dimensions[L(c0 + j)].width = width
    def collapse(ws, a, b, summary):
        ws.column_dimensions.group(L(a), L(b), outline_level=1, hidden=False)   # quarterly columns shown by default; the outline buttons collapse them to annual
    def point_groups(ws, c0):
        prev = None
        for i in range(NQ):
            if not is_ann[i]: continue
            if prev is not None and i - 1 >= prev + 1: collapse(ws, c0 + prev + 1, c0 + i - 1, c0 + i)
            prev = i
    def stat_block(ws, c0, n, specs, start):
        rows = {}
        for k, (label, tmpl, sfmt) in enumerate(specs):
            r = start + k; rows[label] = r; bd = TOP if k == 0 else None
            _put(ws, r, 1, label, F_BOLD, fill=FILL_STAT, border=bd); [_put(ws, r, c, None, fill=FILL_STAT, border=bd) for c in range(2, 5)]
            for j in range(n):
                f = tmpl.replace('{c}', L(c0 + j)).replace('{F}', str(FIRST)).replace('{Z}', str(LAST))
                for lb, rr in rows.items(): f = f.replace('{r_' + lb + '}', str(rr))
                _put(ws, r, c0 + j, f, F_BOLD, sfmt, FILL_STAT, CTR, bd)
        return rows
    def qpiv(col, scale=1.0):
        d = M['piv'](col, scale); return d
    def point_sheet(name, ttl, note, df, fmt, kind, change_fmt, tab):
        ws = wb.create_sheet(name); ws.sheet_properties.tabColor = tab; _title(ws, ttl, note); left_cols(ws); c0 = 5
        for j in range(NQ): _hdr(ws, HDR, c0 + j, plabels[j], is_ann[j])
        data_block(ws, c0, df, fmt); cc = c0 + NQ
        heads = [f'Change {plabels[0]} to {plabels[-1]}', f'Change last 12 months ({plabels[max(0, NQ - 5)]} to {plabels[-1]})', f'Change last quarter ({plabels[-2]} to {plabels[-1]})']
        pairs = [(L(c0), L(c0 + NQ - 1)), (L(c0 + max(0, NQ - 5)), L(c0 + NQ - 1)), (L(c0 + NQ - 2), L(c0 + NQ - 1))]
        for k, (h, (x, y)) in enumerate(zip(heads, pairs)):
            _hdr(ws, HDR, cc + k, h); ws.column_dimensions[L(cc + k)].width = 17
            for i in range(N): r = FIRST + i; _put(ws, r, cc + k, f'=IF(OR({x}{r}="",{y}{r}=""),"",{y}{r}-{x}{r})', fmt=change_fmt, align=CTR)
        ws.row_dimensions[HDR].height = 54; rng = f'{L(c0)}{FIRST}:{L(c0 + NQ - 1)}{LAST}'
        if kind == 'star':
            ws.conditional_formatting.add(rng, STAR)
            specs = [('RI average', '=IFERROR(AVERAGE({c}${F}:{c}${Z}),"")', '0.00'), ('Facilities with a rating', '=COUNT({c}${F}:{c}${Z})', '0')] + [('0 stars (SFF, not rated)', '=COUNTIF({c}${F}:{c}${Z},0)', '0')] + [(f'{k} star' + ('s' if k > 1 else ''), '=COUNTIF({c}${F}:{c}${Z},' + str(k) + ')', '0') for k in range(1, 6)]
            specs += [('% rated 0-2 stars', '=IFERROR(({c}{r_0 stars (SFF, not rated)}+{c}{r_1 star}+{c}{r_2 stars})/{c}{r_Facilities with a rating},"")', '0%'), ('% rated 4-5 stars', '=IFERROR(({c}{r_4 stars}+{c}{r_5 stars})/{c}{r_Facilities with a rating},"")', '0%')]
        elif kind == 'hours':
            ws.conditional_formatting.add(rng, _scale()); specs = [('RI average', '=IFERROR(AVERAGE({c}${F}:{c}${Z}),"")', '0.00'), ('RI median', '=IFERROR(MEDIAN({c}${F}:{c}${Z}),"")', '0.00'), ('Facilities reporting', '=COUNT({c}${F}:{c}${Z})', '0')]
        else:
            ws.conditional_formatting.add(rng, _rev()); specs = [('RI average', '=IFERROR(AVERAGE({c}${F}:{c}${Z}),"")', '0%'), ('RI median', '=IFERROR(MEDIAN({c}${F}:{c}${Z}),"")', '0%'), ('Facilities reporting', '=COUNT({c}${F}:{c}${Z})', '0')]
        rows = stat_block(ws, c0, NQ, specs, LAST + 2); ws.freeze_panes = ws.cell(row=FIRST, column=c0); ws.auto_filter.ref = f'A{HDR}:{L(cc + 2)}{LAST}'; point_groups(ws, c0)
        SR[name] = dict(rows=rows, c0=c0, chg=L(cc))
    point_sheet('Overall Rating', 'Overall 5-star rating - Rhode Island nursing homes', note_std + ' 5 = best.', qpiv('overall_rating'), '0', 'star', CHI, '2E75B6')
    point_sheet('Health Inspection', 'Health inspection star rating', note_std + ' 5 = best.', qpiv('health_inspection_rating'), '0', 'star', CHI, '2E75B6')
    point_sheet('Staffing Rating', 'Staffing star rating', note_std + ' 5 = best. CMS revised the staffing rating method during this period.', qpiv('staffing_rating'), '0', 'star', CHI, '2E75B6')
    point_sheet('QM Rating', 'Quality measures (QM) star rating', note_std + ' 5 = best.', qpiv('qm_rating'), '0', 'star', CHI, '2E75B6')
    point_sheet('Nurse Hours', 'Total nurse staffing hours per resident per day (reported)', note_std + ' RN + LPN + nurse aide hours from CMS payroll data. Higher = more staff time per resident.', qpiv('reported_total_nurse_staffing_hours_per_resident_per_day'), '0.00', 'hours', CH, '548235')
    point_sheet('RN Hours', 'Registered nurse (RN) staffing hours per resident per day (reported)', note_std, qpiv('reported_rn_staffing_hours_per_resident_per_day'), '0.00', 'hours', CH, '548235')
    point_sheet('Turnover', 'Total nursing staff turnover (annual %)', 'CMS publishes turnover only from 2022, so earlier columns are blank. ' + note_std + ' Lower = more stable staff.', qpiv('total_nursing_staff_turnover', 0.01), '0%', 'turnover', '+0%;-0%;0%', '548235')
    # ---- occupancy
    ws = wb.create_sheet('Occupancy'); ws.sheet_properties.tabColor = '548235'; _title(ws, 'Certified beds, average residents, and occupancy', note_std + ' Occupancy = average residents per day / certified beds (formula).'); left_cols(ws)
    b1 = 5; b2 = b1 + NQ + 1; b3 = b2 + NQ + 1
    for b, t in [(b1, 'Certified beds'), (b2, 'Average residents per day'), (b3, 'Occupancy (residents / beds)')]:
        _put(ws, 3, b, t, F_BOLD, fill=FILL_GRP); [_put(ws, 3, b + j, None, fill=FILL_GRP) for j in range(1, NQ)]
        for j in range(NQ): _hdr(ws, HDR, b + j, plabels[j], is_ann[j])
    data_block(ws, b1, qpiv('number_of_certified_beds'), '#,##0'); data_block(ws, b2, qpiv('average_number_of_residents_per_day'), '#,##0.0')
    for j in range(NQ): ws.column_dimensions[L(b3 + j)].width = 11
    for i in range(N):
        for j in range(NQ): r = FIRST + i; bc, rc = L(b1 + j), L(b2 + j); _put(ws, r, b3 + j, f'=IF(OR({bc}{r}="",{rc}{r}="",{bc}{r}=0),"",{rc}{r}/{bc}{r})', fmt='0%', align=CTR)
    ws.conditional_formatting.add(f'{L(b3)}{FIRST}:{L(b3 + NQ - 1)}{LAST}', _scale()); occ_rows = {}; start = LAST + 2
    for k, (label, kind) in enumerate([('Total certified beds', 'beds'), ('Total average residents', 'res'), ('Statewide occupancy', 'occ'), ('Average facility occupancy', 'avg')]):
        r = start + k; occ_rows[label] = r; bd = TOP if k == 0 else None; _put(ws, r, 1, label, F_BOLD, fill=FILL_STAT, border=bd); [_put(ws, r, c, None, fill=FILL_STAT, border=bd) for c in range(2, 5)]
        for j in range(NQ):
            bc, rc, oc = L(b1 + j), L(b2 + j), L(b3 + j)
            if kind == 'beds': _put(ws, r, b1 + j, f'=SUM({bc}{FIRST}:{bc}{LAST})', F_BOLD, '#,##0', FILL_STAT, CTR, bd)
            if kind == 'res': _put(ws, r, b2 + j, f'=SUM({rc}{FIRST}:{rc}{LAST})', F_BOLD, '#,##0', FILL_STAT, CTR)
            if kind == 'occ': _put(ws, r, b3 + j, f'=IFERROR({rc}{start + 1}/{bc}{start},"")', F_BOLD, '0%', FILL_STAT, CTR)
            if kind == 'avg': _put(ws, r, b3 + j, f'=IFERROR(AVERAGE({oc}{FIRST}:{oc}{LAST}),"")', F_BOLD, '0%', FILL_STAT, CTR)
    ws.freeze_panes = ws.cell(row=FIRST, column=5); ws.auto_filter.ref = f'A{HDR}:{L(b3 + NQ - 1)}{LAST}'
    for b in (b1, b2, b3): point_groups(ws, b)
    SR['Occupancy'] = dict(rows=occ_rows, b1=b1, b2=b2, b3=b3)
    # ---- citations / fines by quarter and year
    def year_sheet(name, ttl, note, blocks, tab):
        ws = wb.create_sheet(name); ws.sheet_properties.tabColor = tab
        _title(ws, ttl, note + ' Year columns add up their quarters (formula). Quarters are shown; click the small 1 button at the top-left (or the - above the columns) to collapse to years only.'); left_cols(ws); ws.row_dimensions[HDR].height = 30
        c0 = 5; info = {}
        for key, (dfq, fmt, bt, sfmt) in blocks.items():
            col = c0; ycol = {}; qcol = {}
            for y in years:
                for (yy, q) in [x for x in qpairs if x[0] == y]: qcol[(yy, q)] = col; _hdr(ws, HDR, col, f'Q{q} {yy}', False); col += 1
                ycol[y] = col; _hdr(ws, HDR, col, ylabels[years.index(y)]); col += 1
            tcol = col; _hdr(ws, HDR, tcol, f'Total {years[0]}-{years[-1]}')
            _put(ws, 3, c0, bt, F_BOLD, fill=FILL_GRP); [_put(ws, 3, c, None, fill=FILL_GRP) for c in range(c0 + 1, tcol + 1)]
            for i in range(N):
                r = FIRST + i
                for (yy, q), cn in qcol.items(): _put(ws, r, cn, dfq.loc[order[i], (yy, q)], fmt=fmt, align=CTR)
                for y in years:
                    qs = [qcol[x] for x in qpairs if x[0] == y]; _put(ws, r, ycol[y], f'=IF(COUNT({L(qs[0])}{r}:{L(qs[-1])}{r})=0,"",SUM({L(qs[0])}{r}:{L(qs[-1])}{r}))', F_BOLD, fmt, align=CTR)
                _put(ws, r, tcol, '=SUM(' + ','.join(f'{L(ycol[y])}{r}' for y in years) + ')', F_BOLD, fmt, align=CTR)
            for c in range(c0, tcol + 1): ws.column_dimensions[L(c)].width = 11 if c != tcol else 13
            rows = {}; start = LAST + 2
            for k, (label, tmpl, sf) in enumerate([('RI total', '=SUM({c}${F}:{c}${Z})', sfmt), ('Facilities with at least one', '=COUNTIF({c}${F}:{c}${Z},">0")', '0')]):
                r = start + k; rows[label] = r; bd = TOP if k == 0 else None; _put(ws, r, 1, label, F_BOLD, fill=FILL_STAT, border=bd); [_put(ws, r, c, None, fill=FILL_STAT, border=bd) for c in range(2, 5)]
                for c in range(c0, tcol + 1): _put(ws, r, c, tmpl.replace('{c}', L(c)).replace('{F}', str(FIRST)).replace('{Z}', str(LAST)), F_BOLD, sf, FILL_STAT, CTR, bd)
            info[key] = dict(qcol=qcol, ycol=ycol, total_col=L(tcol), rows=rows)
            for y in years:
                qs = [qcol[x] for x in qpairs if x[0] == y]; collapse(ws, qs[0], qs[-1], ycol[y])
            c0 = tcol + 2
        ws.freeze_panes = ws.cell(row=FIRST, column=5); ws.auto_filter.ref = f'A{HDR}:{L(c0 - 2)}{LAST}'; SR[name] = info
    cnote = 'Counted by the date of the inspection (survey), from CMS health citation records; a citation that appears in many monthly files is counted once, and citations CMS later removed are excluded. Blank = no longer in CMS data.'
    year_sheet('Citations - All', 'Health citations (deficiencies) by quarter and year', cnote, {'a': (CQ['cit_all'], '0', 'All health citations', '#,##0')}, 'C00000')
    year_sheet('Citations - Serious', 'Serious health citations (actual harm or immediate jeopardy, scope/severity G-L)', cnote + ' G-I = actual harm; J-L = immediate jeopardy.', {'a': (CQ['cit_ser'], '0', 'Serious citations (G-L)', '#,##0')}, 'C00000')
    year_sheet('Citations - Abuse', 'Abuse, neglect and exploitation citations', cnote + ' Category: "Freedom from Abuse, Neglect, and Exploitation".', {'a': (CQ['cit_abu'], '0', 'Abuse / neglect / exploitation citations', '#,##0')}, 'C00000')
    year_sheet('Fines', 'Federal fines and payment denials', 'Counted by penalty date, from CMS penalty records; the most recent version of each penalty is used (CMS sometimes adjusts amounts or adds penalties later), and penalties CMS later removed are excluded. Blank = no longer in CMS data.',
               {'amt': (CQ['fines_amt'], '$#,##0', 'Fines ($)', '$#,##0'), 'n': (CQ['fines_n'], '0', 'Number of fines', '#,##0'), 'den': (CQ['denials'], '0', 'Payment denials (count)', '#,##0')}, 'C00000')
    return dict(SR=SR, plabels=plabels, ann_idx=ann_idx, alabels=alabels, ylabels=ylabels, left_cols=left_cols, point_groups=point_groups, FIRST=FIRST, LAST=LAST, HDR=HDR, N=N, NQ=NQ)

# ------------------------------------------------------------------ workbook: logs, summary, trends, data tabs
LOG_HDRS = [('Date (CMS file, or CEC date)', 15), ('Facility', 40), ('CCN', 9), ('Category', 28), ('What changed', 80), ('Before', 34), ('After', 34), ('Notes', 60)]
CEC_CAT = 'CEC decision (RIDOH)'

def cec_events_df(cec, ccn_of):
    rows = []
    for r in cec['log'].itertuples():
        d = T(r.chow) or T(r.decision)
        if d is None: continue
        who = re.sub(r'\s+', ' ', str(nz(r.applicant) or '')).strip()
        rows.append((d.strftime('%Y-%m-%d'), ccn_of.get(r.facility) or '', CEC_CAT, f"{r.type or ''}: {who}".strip(': '), nz(r.seller) or '', nz(r.buyer) or '',
                     f"{r.cec}; {r.outcome or ''}; date = " + ('CMS takeover date' if T(r.chow) else 'RIDOH decision date') , r.facility))
    return pd.DataFrame(rows, columns=['snap', 'ccn', 'category', 'what', 'before', 'after', 'note', 'fname'])

def write_logs_and_summary(wb, M, S, cec, mp, ccn_of, fn, cfg, ctx):
    fac, order, qpts, is_ann, p = M['fac'], M['order'], M['qpts'], M['is_ann'], M['p']; NQ, N, FIRST, LAST, HDR = ctx['NQ'], ctx['N'], ctx['FIRST'], ctx['LAST'], ctx['HDR']; SR = ctx['SR']
    plabels, alabels, ylabels = ctx['plabels'], ctx['alabels'], ctx['ylabels']; years, latest = M['years'], M['latest']; left_cols, point_groups = ctx['left_cols'], ctx['point_groups']
    ce = cec_events_df(cec, ccn_of); ctx['ce'] = ce
    # ---- Operator by Quarter (+ hidden Grid Flags marking quarters that hold a CEC decision)
    ws = wb.create_sheet('Operator by Quarter'); ws.sheet_properties.tabColor = '7030A0'
    _title(ws, 'Who operates each facility, by quarter (CMS "legal business name")', 'Amber = the operator differs from the previous column (a possible change of hands). Purple = a RIDOH Change in Effective Control decision (or CMS takeover date) falls in that period. In the annual view each year is compared with the prior year. Blank = not reported. Dates are on the Ownership Changes tab.')
    left_cols(ws); c0 = 5
    for j in range(NQ): _hdr(ws, HDR, c0 + j, plabels[j], is_ann[j]); ws.column_dimensions[L(c0 + j)].width = 26
    op = M['op']
    for i in range(N):
        for j in range(NQ): _put(ws, FIRST + i, c0 + j, op.iloc[i, j], Font(name=FONT, size=9))
    gf = wb.create_sheet('Grid Flags'); gf.sheet_state = 'hidden'; _put(gf, 1, 1, 'Helper for the purple highlight on Operator by Quarter: 1 = a CEC decision falls in the period ending at that column.', F_NOTE)
    cd = ce[ce.ccn != ''].copy(); cd['d'] = pd.to_datetime(cd.snap); by = {c: list(g.d) for c, g in cd.groupby('ccn')}
    prev_ann = None; prev_of = {}
    for j in range(NQ):
        prev_of[j] = (j - 1 if not is_ann[j] else prev_ann)
        if is_ann[j]: prev_ann = j
    for i, ccn in enumerate(order):
        for j in range(NQ):
            pj = prev_of[j]
            if pj is None: continue
            lo, hi = pd.Timestamp(qpts[pj]), pd.Timestamp(qpts[j])
            if any(lo < d <= hi for d in by.get(ccn, [])): gf.cell(row=FIRST + i, column=c0 + j, value=1)
    amber = PatternFill('solid', start_color='FFC000', end_color='FFC000'); purple = PatternFill('solid', start_color='D9D2E9', end_color='D9D2E9')
    for j in range(1, NQ):
        col = L(c0 + j); pj = prev_of[j]
        if pj is None: continue
        prev = L(c0 + pj)
        ws.conditional_formatting.add(f'{col}{FIRST}:{col}{LAST}', FormulaRule(formula=[f'AND({col}{FIRST}<>"",{prev}{FIRST}<>"",{col}{FIRST}<>{prev}{FIRST})'], fill=amber, font=Font(name=FONT, size=9, bold=True)))
        ws.conditional_formatting.add(f'{col}{FIRST}:{col}{LAST}', FormulaRule(formula=[f"'Grid Flags'!{col}{FIRST}=1"], fill=purple, font=Font(name=FONT, size=9, bold=True)))
    ws.freeze_panes = ws.cell(row=FIRST, column=c0); ws.auto_filter.ref = f'A{HDR}:{L(c0 + NQ - 1)}{LAST}'; point_groups(ws, c0)
    # ---- change logs
    def log_sheet(name, ttl, note, df, tab):
        ws = wb.create_sheet(name); ws.sheet_properties.tabColor = tab; _title(ws, ttl, note)
        for j, (h, w) in enumerate(LOG_HDRS): _hdr(ws, HDR, j + 1, h); ws.column_dimensions[L(j + 1)].width = w
        ws.row_dimensions[HDR].height = 30; wrap = Alignment(wrap_text=True, vertical='top'); top = Alignment(vertical='top'); nt = Font(name=FONT, size=9, italic=True, color='595959')
        for i, r_ in enumerate(df.itertuples()):
            r = FIRST + i; nm = fac.loc[r_.ccn, 'name'] if r_.ccn in fac.index else (getattr(r_, 'fname', '') or '')
            _put(ws, r, 1, dt.datetime.strptime(r_.snap, '%Y-%m-%d'), fmt='yyyy-mm-dd', align=top); _put(ws, r, 2, nm, align=top); _put(ws, r, 3, str(r_.ccn) if r_.ccn else '', fmt='@', align=top)
            _put(ws, r, 4, r_.category, align=top); _put(ws, r, 5, r_.what, align=wrap); _put(ws, r, 6, r_.before, align=wrap); _put(ws, r, 7, r_.after, align=wrap); _put(ws, r, 8, r_.note, nt, align=wrap)
        last = FIRST + max(len(df), 1) - 1; ws.freeze_panes = ws.cell(row=FIRST, column=3); ws.auto_filter.ref = f'A{HDR}:H{last}'; return dict(last=last, n=len(df))
    ev = M['ev_own'].copy(); ev['fname'] = ''
    ev = pd.concat([ev, ce[ev.columns.intersection(ce.columns).tolist()].reindex(columns=ev.columns, fill_value='')], ignore_index=True)
    ev['_p'] = (~ev.category.isin(M['strong'] + [CEC_CAT])).astype(int)
    own_info = log_sheet('Ownership Changes', 'Ownership and operator changes: CMS data plus RIDOH Change in Effective Control (CEC) decisions',
        'Listed first (newest first): "CEC decision (RIDOH)" rows from the CEC Log (date = CMS takeover date where known, else the state decision date) and the strong CMS signals: "Operator change" (legal business name), "Change of ownership (CMS flag)", "Profit status change". Then supporting changes: "Owner list change" (5%+ owners, managerial control) and "Chain change". For CMS rows the date is the first monthly CMS file showing the change (CMS can lag the real date). Filter by Category. Officer / director lists are excluded because CMS re-coded those roles in Dec 2024, Jun 2025 and May 2026.',
        ev.sort_values(['_p', 'snap', 'ccn'], ascending=[True, False, True]).drop(columns='_p'), '7030A0')
    oth_info = log_sheet('Other Changes', 'Other changes detected in the monthly CMS files', 'Facility name in the CMS listing, special focus status, abuse icon, certified beds, overall star rating, and facilities that left the CMS data. Date = first monthly file showing the change. Filter by Category.', M['ev_other'].assign(fname=''), '7030A0')
    ctx['own_info'] = own_info
    # ---- CMS Summary
    ws = wb.create_sheet('CMS Summary'); ws.sheet_properties.tabColor = '1F3864'
    _title(ws, 'Rhode Island nursing homes - summary (latest CMS data)', f'Latest = {lab_(latest)} CMS file, or the last file listing the facility. Star ratings 1-5 (5 = best). Changes compare {alabels[0]} with {alabels[-1]}. Citation and fine totals cover {years[0]} to {lab_(latest)}. Click a header arrow to sort or filter.')
    spec = [('name', 'Facility', 40), ('ccn', 'CCN', 9), ('former', 'Former / other names', 52), ('city', 'City', 16), ('county', 'County', 13), ('status', 'Status', 27), ('own', 'Ownership type', 30), ('chain', 'Chain (CMS, 2025+)', 24), ('oper', 'Operator (CMS legal business name)', 34),
            ('beds', 'Certified beds', 10), ('res', 'Avg residents / day', 11), ('occ', 'Occupancy', 11), ('ovr', 'Overall rating', 9), ('hi', 'Health inspection', 10), ('stf', 'Staffing rating', 9), ('qm', 'QM rating', 9),
            ('hrs', 'Total nurse hrs / resident / day', 12), ('rn', 'RN hrs / resident / day', 12), ('turn', 'Nursing staff turnover', 11), ('sff', 'Special focus status', 15), ('abu', 'Abuse icon', 8),
            ('c_ovr', f'Overall rating change {alabels[0]} to {alabels[-1]}', 13), ('c_hi', 'Health inspection change', 12), ('c_stf', 'Staffing rating change', 12), ('c_hrs', 'Nurse hrs change', 11),
            ('cit_s', f'Serious citations (G-L) {years[0]}-{years[-1]}', 13), ('cit_a', f'All citations {years[0]}-{years[-1]}', 11), ('cit_ab', 'Abuse / neglect citations', 12), ('fin', f'Fines ($) {years[0]}-{years[-1]}', 13), ('den', 'Payment denials', 10),
            ('chows', 'Ownership / operator changes seen in CMS data (strong signals)', 16), ('lastchow', 'Most recent CMS-detected change', 14), ('cec', 'CEC decisions on record (RIDOH)', 12), ('cms', 'Name in latest CMS file', 46)]
    CI = {k: i + 1 for i, (k, _, _) in enumerate(spec)}; CL = {k: L(v) for k, v in CI.items()}
    for k, h, w in spec: _hdr(ws, HDR, CI[k], h); ws.column_dimensions[CL[k]].width = w
    ws.row_dimensions[HDR].height = 66
    lk = lambda sheet, col, r: f"=IFERROR(INDEX('{sheet}'!${col}${FIRST}:${col}${LAST},MATCH($B{r},'{sheet}'!$B${FIRST}:$B${LAST},0)),\"\")"
    pl = p.sort_values('snapshot_date').groupby('ccn_std').tail(1).set_index('ccn_std'); OC = "'Ownership Changes'!"; ol = own_info['last']
    rng = lambda col: f"{OC}${col}${FIRST}:${col}${ol}"; STR = M['strong']
    for i, ccn in enumerate(order):
        r = FIRST + i; row = fac.loc[ccn]; x = pl.loc[ccn]; g = F_GREY if not row['active'] else F_NORM
        _put(ws, r, CI['name'], row['name'], g); _put(ws, r, CI['ccn'], str(ccn), g, '@'); _put(ws, r, CI['former'], names_text(fn.get(ccn, [])), Font(name=FONT, size=9, color=('7F7F7F' if not row['active'] else '000000')), align=Alignment(wrap_text=True, vertical='top'))
        _put(ws, r, CI['city'], row['city'], g); _put(ws, r, CI['county'], row['county'], g); _put(ws, r, CI['status'], row['status'], g); _put(ws, r, CI['own'], x['ownership_type'], g); _put(ws, r, CI['chain'], row['chain'], g); _put(ws, r, CI['oper'], row['operator'], g)
        _put(ws, r, CI['beds'], x['number_of_certified_beds'], fmt='#,##0', align=CTR); _put(ws, r, CI['res'], x['average_number_of_residents_per_day'], fmt='#,##0.0', align=CTR)
        _put(ws, r, CI['occ'], f'=IF(OR({CL["beds"]}{r}="",{CL["res"]}{r}="",{CL["beds"]}{r}=0),"",{CL["res"]}{r}/{CL["beds"]}{r})', fmt='0%', align=CTR)
        for k, c_ in [('ovr', 'overall_rating'), ('hi', 'health_inspection_rating'), ('stf', 'staffing_rating'), ('qm', 'qm_rating')]: _put(ws, r, CI[k], x[c_], fmt='0', align=CTR)
        _put(ws, r, CI['hrs'], x['reported_total_nurse_staffing_hours_per_resident_per_day'], fmt='0.00', align=CTR); _put(ws, r, CI['rn'], x['reported_rn_staffing_hours_per_resident_per_day'], fmt='0.00', align=CTR)
        tv = x['total_nursing_staff_turnover']; _put(ws, r, CI['turn'], None if pd.isna(tv) else tv / 100, fmt='0%', align=CTR); _put(ws, r, CI['sff'], x['special_focus_status'], align=CTR); _put(ws, r, CI['abu'], x['abuse_icon'], align=CTR)
        _put(ws, r, CI['c_ovr'], lk('Overall Rating', SR['Overall Rating']['chg'], r), fmt=CHI, align=CTR); _put(ws, r, CI['c_hi'], lk('Health Inspection', SR['Health Inspection']['chg'], r), fmt=CHI, align=CTR)
        _put(ws, r, CI['c_stf'], lk('Staffing Rating', SR['Staffing Rating']['chg'], r), fmt=CHI, align=CTR); _put(ws, r, CI['c_hrs'], lk('Nurse Hours', SR['Nurse Hours']['chg'], r), fmt=CH, align=CTR)
        _put(ws, r, CI['cit_s'], lk('Citations - Serious', SR['Citations - Serious']['a']['total_col'], r), fmt='0', align=CTR); _put(ws, r, CI['cit_a'], lk('Citations - All', SR['Citations - All']['a']['total_col'], r), fmt='0', align=CTR)
        _put(ws, r, CI['cit_ab'], lk('Citations - Abuse', SR['Citations - Abuse']['a']['total_col'], r), fmt='0', align=CTR); _put(ws, r, CI['fin'], lk('Fines', SR['Fines']['amt']['total_col'], r), fmt='$#,##0', align=CTR); _put(ws, r, CI['den'], lk('Fines', SR['Fines']['den']['total_col'], r), fmt='0', align=CTR)
        _put(ws, r, CI['chows'], '=' + '+'.join(f'COUNTIFS({rng("C")},$B{r},{rng("D")},"{c_}")' for c_ in STR), fmt='0', align=CTR)
        mx = ','.join(f'_xlfn.MAXIFS({rng("A")},{rng("C")},$B{r},{rng("D")},"{c_}")' for c_ in STR); _put(ws, r, CI['lastchow'], f'=IF({CL["chows"]}{r}=0,"",MAX({mx}))', fmt='mmm yyyy', align=CTR)
        _put(ws, r, CI['cec'], f'=COUNTIFS({rng("C")},$B{r},{rng("D")},"{CEC_CAT}")', fmt='0', align=CTR); _put(ws, r, CI['cms'], row['cms_name'], g)
    for k in ('ovr', 'hi', 'stf', 'qm'): ws.conditional_formatting.add(f'{CL[k]}{FIRST}:{CL[k]}{LAST}', STAR)
    ws.conditional_formatting.add(f'{CL["hrs"]}{FIRST}:{CL["hrs"]}{LAST}', _scale()); ws.conditional_formatting.add(f'{CL["turn"]}{FIRST}:{CL["turn"]}{LAST}', _rev())
    ws.conditional_formatting.add(f'{CL["cit_s"]}{FIRST}:{CL["cit_s"]}{LAST}', ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFF', end_type='max', end_color='F8696B'))
    ws.conditional_formatting.add(f'{CL["fin"]}{FIRST}:{CL["fin"]}{LAST}', ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFF', end_type='percentile', end_value=95, end_color='F8696B'))
    ws.conditional_formatting.add(f'{CL["chows"]}{FIRST}:{CL["chows"]}{LAST}', ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFF', end_type='max', end_color='B4A7D6'))
    r = LAST + 2; nsp = len(spec); _put(ws, r, 1, 'RI total / average (active facilities; citation, fine and change columns include all)', F_BOLD, fill=FILL_STAT, border=TOP); [_put(ws, r, c, None, fill=FILL_STAT, border=TOP) for c in range(2, nsp + 1)]
    E = f'${CL["status"]}${FIRST}:${CL["status"]}${LAST}'
    for k, (f, fm) in {'beds': ('SUMIF', '#,##0'), 'res': ('SUMIF', '#,##0'), 'ovr': ('AVERAGEIF', '0.00'), 'hi': ('AVERAGEIF', '0.00'), 'stf': ('AVERAGEIF', '0.00'), 'qm': ('AVERAGEIF', '0.00'), 'hrs': ('AVERAGEIF', '0.00'), 'rn': ('AVERAGEIF', '0.00'), 'turn': ('AVERAGEIF', '0%')}.items():
        _put(ws, r, CI[k], f'=IFERROR({f}({E},"Active",{CL[k]}{FIRST}:{CL[k]}{LAST}),"")', F_BOLD, fm, FILL_STAT, CTR, TOP)
    for k, fm in {'cit_s': '#,##0', 'cit_a': '#,##0', 'cit_ab': '#,##0', 'fin': '$#,##0', 'den': '#,##0', 'chows': '#,##0', 'cec': '#,##0'}.items(): _put(ws, r, CI[k], f'=IFERROR(SUM({CL[k]}{FIRST}:{CL[k]}{LAST}),"")', F_BOLD, fm, FILL_STAT, CTR, TOP)
    _put(ws, r, CI['occ'], f'=IFERROR({CL["res"]}{r}/{CL["beds"]}{r},"")', F_BOLD, '0%', FILL_STAT, CTR, TOP)
    ws.freeze_panes = ws.cell(row=FIRST, column=2); ws.auto_filter.ref = f'A{HDR}:{L(nsp)}{LAST}'; ctx['sum_last'] = LAST

def write_trends(wb, M, ctx):
    NQ, N, FIRST, LAST, HDR, SR = ctx['NQ'], ctx['N'], ctx['FIRST'], ctx['LAST'], ctx['HDR'], ctx['SR']; plabels, alabels, ylabels = ctx['plabels'], ctx['alabels'], ctx['ylabels']; is_ann = M['is_ann']; years, qpairs = M['years'], M['qpairs']
    def trend_sheet(name, idxs, plab, ys, tab, wide):
        ws = wb.create_sheet(name); ws.sheet_properties.tabColor = tab; n = len(idxs)
        _title(ws, f'Rhode Island statewide trends - {"quarterly snapshots" if wide else "annual snapshots"}', 'All figures are formulas over the facility tabs. Facilities that left the CMS data drop out of averages after they leave. ' + ('Annual view: "RI Trends (Annual)".' if wide else 'Quarterly detail: "RI Trends (Quarterly)".'))
        _put(ws, 3, 1, 'Snapshots (point in time)', F_BOLD, fill=FILL_GRP); [_put(ws, 3, c, None, fill=FILL_GRP) for c in range(2, n + 2)]; _hdr(ws, HDR, 1, 'Metric')
        for j, k in enumerate(idxs): _hdr(ws, HDR, 2 + j, plab[j], is_ann[k])
        ws.column_dimensions['A'].width = 46
        for c in range(2, n + 2): ws.column_dimensions[L(c)].width = 10 if wide else 12
        t = {}; r = FIRST
        refp = lambda sh, key, k: f"='{sh}'!{L(SR[sh]['c0'] + k)}{SR[sh]['rows'][key]}"
        lines = [('Facilities with an overall rating', 'Overall Rating', 'Facilities with a rating', '0'), ('Average overall rating (SFF = 0)', 'Overall Rating', 'RI average', '0.00'), ('% of facilities rated 0-2 stars (SFF = 0)', 'Overall Rating', '% rated 0-2 stars', '0%'), ('% of facilities rated 4-5 stars', 'Overall Rating', '% rated 4-5 stars', '0%'),
                 ('Average health inspection rating', 'Health Inspection', 'RI average', '0.00'), ('Average staffing rating', 'Staffing Rating', 'RI average', '0.00'), ('Average QM rating', 'QM Rating', 'RI average', '0.00'), ('Average total nurse hours / resident / day', 'Nurse Hours', 'RI average', '0.00'), ('Average RN hours / resident / day', 'RN Hours', 'RI average', '0.00')]
        for label, sh, key, fm in lines:
            _put(ws, r, 1, label); t[label] = r
            for j, k in enumerate(idxs): _put(ws, r, 2 + j, refp(sh, key, k), fmt=fm, align=CTR)
            r += 1
        _put(ws, r, 1, 'Average nursing staff turnover (published from 2022)'); t['turnover'] = r; rr = SR['Turnover']['rows']['RI average']
        for j, k in enumerate(idxs): _put(ws, r, 2 + j, f"=IF('Turnover'!{L(SR['Turnover']['c0'] + k)}{rr}=\"\",\"n/a\",'Turnover'!{L(SR['Turnover']['c0'] + k)}{rr})", fmt='0%', align=CTR)
        r += 1; _put(ws, r, 1, 'Statewide occupancy (residents / certified beds)'); t['occ'] = r
        for j, k in enumerate(idxs): _put(ws, r, 2 + j, f"='Occupancy'!{L(SR['Occupancy']['b3'] + k)}{SR['Occupancy']['rows']['Statewide occupancy']}", fmt='0%', align=CTR)
        r += 1; _put(ws, r, 1, 'Certified beds statewide'); t['beds'] = r
        for j, k in enumerate(idxs): _put(ws, r, 2 + j, f"='Occupancy'!{L(SR['Occupancy']['b1'] + k)}{SR['Occupancy']['rows']['Total certified beds']}", fmt='#,##0', align=CTR)
        r += 2; _put(ws, r, 1, 'By ' + ('quarter' if wide else 'calendar year') + ' (inspection / penalty date)', F_BOLD, fill=FILL_GRP); [_put(ws, r, c, None, fill=FILL_GRP) for c in range(2, len(ys) + 2)]
        r += 1; _hdr(ws, r, 1, 'Metric'); yh = r
        for j, (lb, _, ann) in enumerate(ys): _hdr(ws, r, 2 + j, lb, ann)
        for c in range(2, len(ys) + 2): ws.column_dimensions[L(c)].width = max(ws.column_dimensions[L(c)].width or 0, 10 if wide else 12)
        r += 1
        for label, sh, key, fm in [('All health citations', 'Citations - All', 'a', '#,##0'), ('Serious citations (G-L: harm / immediate jeopardy)', 'Citations - Serious', 'a', '#,##0'), ('Abuse / neglect / exploitation citations', 'Citations - Abuse', 'a', '#,##0'), ('Federal fines ($)', 'Fines', 'amt', '$#,##0'), ('Number of fines', 'Fines', 'n', '#,##0'), ('Payment denials', 'Fines', 'den', '#,##0')]:
            _put(ws, r, 1, label); t[label] = r; inf = SR[sh][key]
            for j, (lb, ref, ann) in enumerate(ys): _put(ws, r, 2 + j, f"='{sh}'!{L(inf['ycol'][ref] if not wide else inf['qcol'][ref])}{inf['rows']['RI total']}", fmt=fm, align=CTR)
            r += 1
        ws.freeze_panes = 'B5'; cr = r + 2; w_ = 27 if wide else 17
        def line(title_, keys, anchor, ymin=None, ymax=None, fmt=None):
            ch = LineChart(); ch.title = title_; ch.height = 8; ch.width = w_
            for kk in keys: ch.add_data(Reference(ws, min_col=1, max_col=n + 1, min_row=t[kk], max_row=t[kk]), from_rows=True, titles_from_data=True)
            ch.set_categories(Reference(ws, min_col=2, max_col=n + 1, min_row=HDR, max_row=HDR))
            if ymin is not None: ch.y_axis.scaling.min = ymin
            if ymax is not None: ch.y_axis.scaling.max = ymax
            if fmt: ch.y_axis.number_format = fmt
            for sr in ch.series: sr.smooth = False
            ch.y_axis.delete = False; ch.x_axis.delete = False; ch.legend.position = 'b'; ws.add_chart(ch, anchor)
        def bars(title_, key, anchor, fmt=None):
            ch = BarChart(); ch.type = 'col'; ch.title = title_; ch.height = 8; ch.width = w_
            ch.add_data(Reference(ws, min_col=1, max_col=len(ys) + 1, min_row=t[key], max_row=t[key]), from_rows=True, titles_from_data=True); ch.set_categories(Reference(ws, min_col=2, max_col=len(ys) + 1, min_row=yh, max_row=yh))
            if fmt: ch.y_axis.number_format = fmt
            ch.y_axis.delete = False; ch.x_axis.delete = False; ch.legend = None; ws.add_chart(ch, anchor)
        second = 'M' if wide else 'F'
        line('Average star ratings (SFF = 0)', ['Average overall rating (SFF = 0)', 'Average health inspection rating', 'Average staffing rating', 'Average QM rating'], f'A{cr}', 0, 5)
        line('Share of facilities rated 0-2 vs 4-5 stars', ['% of facilities rated 0-2 stars (SFF = 0)', '% of facilities rated 4-5 stars'], f'{second}{cr}', 0, None, '0%')
        line('Average nurse staffing hours per resident per day', ['Average total nurse hours / resident / day', 'Average RN hours / resident / day'], f'A{cr + 17}', 0)
        bars('Serious citations (G-L)', 'Serious citations (G-L: harm / immediate jeopardy)', f'{second}{cr + 17}'); bars('Federal fines ($)', 'Federal fines ($)', f'A{cr + 34}', '$#,##0'); line('Facilities with an overall rating', ['Facilities with an overall rating'], f'{second}{cr + 34}', 0)
        SR[name] = t
    trend_sheet('RI Trends (Annual)', ctx['ann_idx'], ctx['alabels'], [(ylabels[i], y, True) for i, y in enumerate(years)], '1F3864', False)
    trend_sheet('RI Trends (Quarterly)', list(range(NQ)), ctx['plabels'], [(f'Q{q} {y}', (y, q), q == 3) for (y, q) in qpairs], '1F3864', True)

MD_COLS = [('Snapshot date', 'snapshot_date', 'date'), ('CCN', 'ccn_std', 'txt'), ('Facility', 'name', 'txt'), ('City', 'city', 'txt'), ('Overall rating', 'overall_rating', '0'), ('Health inspection', 'health_inspection_rating', '0'), ('Staffing rating', 'staffing_rating', '0'), ('QM rating', 'qm_rating', '0'),
           ('Long-stay QM', 'long_stay_qm_rating', '0'), ('Short-stay QM', 'short_stay_qm_rating', '0'), ('Total nurse hrs / resident / day', 'reported_total_nurse_staffing_hours_per_resident_per_day', '0.00'), ('RN hrs / resident / day', 'reported_rn_staffing_hours_per_resident_per_day', '0.00'),
           ('Certified beds', 'number_of_certified_beds', '#,##0'), ('Avg residents / day', 'average_number_of_residents_per_day', '#,##0.0'), ('Nursing staff turnover', 'total_nursing_staff_turnover', '0%'), ('RN turnover', 'registered_nurse_turnover', '0%'),
           ('Fines, number (CMS 3-yr)', 'number_of_fines', '0'), ('Fines, $ (CMS 3-yr)', 'total_amount_of_fines_in_dollars', '$#,##0'), ('Payment denials (CMS 3-yr)', 'number_of_payment_denials', '0'), ('Special focus status', 'special_focus_status', 'txt'), ('Abuse icon', 'abuse_icon', 'txt'),
           ('Ownership type', 'ownership_type', 'txt'), ('Operator (legal business name)', 'lbn', 'txt'), ('Chain', 'chain_name', 'txt'), ('CMS: changed ownership in last 12 months', 'provider_changed_ownership_in_last_12_months', 'txt'), ('Lookup key (CCN|date serial)', 'key', 'txt')]
MD_COL = {k: L(i + 1) for i, (_, k, _) in enumerate(MD_COLS)}; MD_ROWS = 40000

def write_data_tabs(wb, M, ctx):
    win = M['win'].sort_values(['snapshot_date', 'name']).copy(); win['key'] = [f'{c}|{serial(d)}' for c, d in zip(win.ccn_std, win.snapshot_date)]
    ws = wb.create_sheet('CMS Monthly Data'); ws.sheet_properties.tabColor = '7F7F7F'
    for j, (h, k, f) in enumerate(MD_COLS): _hdr(ws, 1, j + 1, h)
    ws.row_dimensions[1].height = 54
    for i, row in enumerate(win.itertuples(index=False)):
        r = i + 2; rd = row._asdict() if hasattr(row, '_asdict') else None
    cols = {k: win[k].tolist() for _, k, _ in MD_COLS}
    for i in range(len(win)):
        r = i + 2
        for j, (h, k, f) in enumerate(MD_COLS):
            v = cols[k][i]
            if f == 'date': v = dt.datetime.strptime(v, '%Y-%m-%d'); fm = 'yyyy-mm-dd'
            elif f == 'txt': fm = '@' if k in ('ccn_std', 'key') else None
            else:
                fm = f
                if k in ('total_nursing_staff_turnover', 'registered_nurse_turnover') and nz(v) is not None: v = v / 100
            _put(ws, r, j + 1, v, fmt=fm)
    for j, wd in enumerate([12, 9, 42, 16, 9, 10, 9, 9, 9, 9, 12, 11, 9, 10, 11, 9, 10, 12, 11, 14, 8, 30, 36, 24, 14, 16]): ws.column_dimensions[L(j + 1)].width = wd
    ws.freeze_panes = 'E2'; ws.auto_filter.ref = f'A1:{L(len(MD_COLS))}{len(win) + 1}'; ctx['n_monthly'] = len(win)
    hu, pu = M['hu'], M['pu']; fac = M['fac']
    ws = wb.create_sheet('CMS Citations'); ws.sheet_properties.tabColor = '7F7F7F'
    for j, (h, w) in enumerate([('CCN', 9), ('Inspection (survey) date', 14), ('Tag', 8), ('Scope / severity', 10), ('Category', 44), ('Serious (G-L) = 1', 10), ('Abuse category = 1', 10), ('Facility', 40)]): _hdr(ws, 1, j + 1, h); ws.column_dimensions[L(j + 1)].width = w
    ws.row_dimensions[1].height = 42; h2 = hu.sort_values(['survey_date', 'ccn_std'], ascending=[False, True])
    for i, r_ in enumerate(h2.itertuples()):
        r = i + 2; _put(ws, r, 1, str(r_.ccn_std), fmt='@'); _put(ws, r, 2, dt.datetime.strptime(r_.survey_date, '%Y-%m-%d'), fmt='yyyy-mm-dd'); _put(ws, r, 3, f"{r_.deficiency_prefix or ''}{r_.deficiency_tag_number or ''}"); _put(ws, r, 4, r_.scope_severity_code)
        _put(ws, r, 5, r_.deficiency_category); _put(ws, r, 6, int(r_.serious)); _put(ws, r, 7, int(r_.abuse)); _put(ws, r, 8, fac.loc[r_.ccn_std, 'name'] if r_.ccn_std in fac.index else '')
    ws.freeze_panes = 'A2'; ws.auto_filter.ref = f'A1:H{len(h2) + 1}'; ctx['n_cit'] = len(h2)
    ws = wb.create_sheet('CMS Penalties'); ws.sheet_properties.tabColor = '7F7F7F'
    for j, (h, w) in enumerate([('CCN', 9), ('Penalty date', 13), ('Type', 16), ('Fine amount ($)', 14), ('Payment denial start', 14), ('Payment denial days', 12), ('Facility', 40)]): _hdr(ws, 1, j + 1, h); ws.column_dimensions[L(j + 1)].width = w
    ws.row_dimensions[1].height = 42; p2 = pu.sort_values(['penalty_date', 'ccn_std'], ascending=[False, True])
    for i, r_ in enumerate(p2.itertuples()):
        r = i + 2; _put(ws, r, 1, str(r_.ccn_std), fmt='@'); _put(ws, r, 2, dt.datetime.strptime(r_.penalty_date, '%Y-%m-%d'), fmt='yyyy-mm-dd'); _put(ws, r, 3, r_.penalty_type); _put(ws, r, 4, r_.fine_amount, fmt='$#,##0')
        _put(ws, r, 5, getattr(r_, 'payment_denial_start_date', None)); _put(ws, r, 6, getattr(r_, 'payment_denial_length_in_days', None)); _put(ws, r, 7, fac.loc[r_.ccn_std, 'name'] if r_.ccn_std in fac.index else '')
    ws.freeze_panes = 'A2'; ws.auto_filter.ref = f'A1:G{len(p2) + 1}'; ctx['n_pen'] = len(p2)

# ------------------------------------------------------------------ workbook: CCN Map, Sale Timing, tracker formulas, CEC tab additions
TRK = "'Post-Sale Tracker'!"; ST = "'Sale Timing'!"; MDS = "'CMS Monthly Data'!"; LOG = "'CEC Log'!"
def _mdr(col): return f"{MDS}${col}$2:${col}${MD_ROWS}"

def write_ccn_map(wb, mp, ctx):
    if 'CCN Map' in wb.sheetnames: del wb['CCN Map']
    ws = wb.create_sheet('CCN Map'); ws.sheet_properties.tabColor = 'BF9000'; ws.sheet_view.showGridLines = False
    _put(ws, 1, 1, 'CCN Map: which CMS certification number (CCN) belongs to each facility name in the CEC tabs', F_TITLE)
    _put(ws, 2, 1, 'Yellow cells with blue text are yours to edit. A CCN typed here always wins. Leave the CCN blank for a facility that has no CCN (it then gets no CMS numbers, which is correct). New facility names added to the CEC tabs are appended automatically on the next update. Rename a facility? Change the name in column A to match the CEC tabs.', F_NOTE)
    for j, (h, w) in enumerate([('Facility name (as used in the CEC tabs)', 50), ('CCN', 11), ('Name in latest CMS file', 46), ('How matched', 30), ('Notes', 60)]): _hdr(ws, 4, j + 1, h); ws.column_dimensions[L(j + 1)].width = w
    ws.row_dimensions[4].height = 30; rows = list(mp)
    for i, r_ in enumerate(rows):
        r = 5 + i; note = r_['note']
        if r_['how'] == 'no CCN' and note is None: note = 'No CCN exists for this facility'
        _put(ws, r, 1, r_['name'], F_INPUT, fill=FILL_INPUT); _put(ws, r, 2, r_['ccn'], F_INPUT, '@', FILL_INPUT, CTR)
        _put(ws, r, 3, f"=IF(B{r}=\"\",\"\",IFERROR(INDEX('CMS Summary'!$A$5:$A$400,MATCH(B{r},'CMS Summary'!$B$5:$B$400,0)),\"not in CMS files since 2019\"))", fill=FILL_STAT)
        _put(ws, r, 4, r_['how'], F_GREY, fill=FILL_STAT); _put(ws, r, 5, note, F_INPUT, fill=FILL_INPUT)
    for r in range(5 + len(rows), 5 + len(rows) + 15): [_put(ws, r, c, None, F_INPUT, '@' if c == 2 else None, FILL_INPUT) for c in (1, 2, 5)]
    ws.freeze_panes = 'A5'

def write_event_overrides(wb, cec, ctx, cfg):
    if 'Event Overrides' in wb.sheetnames: del wb['Event Overrides']
    ov = cec['overrides']
    if not cec.get('overrides_present'):
        ov = {c: dict(treat=t, date=pd.Timestamp(d) if d else None, label=l, note=n) for c, t, d, l, n in cfg.get('event_overrides_seed', [])}; cec['overrides'] = ov
    ws = wb.create_sheet('Event Overrides'); ws.sheet_properties.tabColor = 'BF9000'; ws.sheet_view.showGridLines = False
    _put(ws, 1, 1, 'Event Overrides: correct how a CEC Log / tracker row is treated', F_TITLE)
    _put(ws, 2, 1, 'Yellow cells with blue text are yours to edit. Use this when a CEC row is not really a sale (for example a home that closed instead). Treat as must be one of: closure, not a sale, sale, parent company change. "closure" and "not a sale" rows are left out of the sales table, the before/after comparison and the tracker numbers; they stay on the facility timeline (with the label and date you give). Applied on the next update.', F_NOTE)
    for j, (h, w) in enumerate([('CEC ID', 12), ('Treat as', 22), ('Date (optional)', 14), ('Public label (shown on the page)', 36), ('Note', 90)]): _hdr(ws, 4, j + 1, h); ws.column_dimensions[L(j + 1)].width = w
    ws.row_dimensions[4].height = 30; r = 5
    for cid, o in ov.items():
        _put(ws, r, 1, cid, F_INPUT, fill=FILL_INPUT); _put(ws, r, 2, o['treat'], F_INPUT, fill=FILL_INPUT); _put(ws, r, 3, o.get('date'), F_INPUT, 'yyyy-mm-dd', FILL_INPUT); _put(ws, r, 4, o.get('label'), F_INPUT, fill=FILL_INPUT); _put(ws, r, 5, o.get('note'), F_INPUT, fill=FILL_INPUT, align=Alignment(wrap_text=True, vertical='top')); r += 1
    for rr in range(r, r + 20): [_put(ws, rr, c, None, F_INPUT, 'yyyy-mm-dd' if c == 3 else None, FILL_INPUT) for c in range(1, 6)]
    ws.freeze_panes = 'A5'

def write_mgmt_agreements(wb, cec, M, S, inp_own, ctx):
    """Editable tab for management agreements (not published anywhere; filed with RIDOH under 216-RICR-40-10-1 s.1.7.2)."""
    if 'Management Agreements' in wb.sheetnames: del wb['Management Agreements']
    ws = wb.create_sheet('Management Agreements'); ws.sheet_properties.tabColor = 'BF9000'; ws.sheet_view.showGridLines = False
    _put(ws, 1, 1, 'Management Agreements: when a buyer took over running the home', F_TITLE)
    _put(ws, 2, 1, 'Yellow cells with blue text are yours to edit. A buyer often manages a home under a management agreement months before the state approves the sale, so the "before" ratings can already be the buyer\'s. RIDOH holds these contracts (216-RICR-40-10-1 s.1.7.2 requires them to be filed, with 30 days notice before a new manager starts) but does not publish them: request them under the Access to Public Records Act. Rows with no start date are suggestions from CMS\'s owner list, where a party with managerial control appears before the decision; confirm before relying on them.', F_NOTE)
    for j, (h, w) in enumerate([('CEC ID', 12), ('Facility', 40), ('Management start date', 16), ('Manager (company)', 38), ('Source / notes', 70)]): _hdr(ws, 4, j + 1, h); ws.column_dimensions[L(j + 1)].width = w
    ws.row_dimensions[4].height = 30
    have = cec.get('mgmt', {}); rows = []
    for cid, m in have.items():
        lr = cec['log'][cec['log'].cec == cid]
        rows.append((cid, (lr.iloc[0].facility if len(lr) else ''), m.get('start'), m.get('who'), m.get('src')))
    if not cec.get('mgmt_present'):
        own = inp_own.copy(); own['role'] = own.role_played_by_owner_or_manager_in_facility
        own = own[own['role'].str.upper().str.contains('MANAGERIAL', na=False)].copy()
        own['assoc'] = pd.to_datetime(own['association_date'].astype(str).str.replace('since ', '', regex=False), errors='coerce')
        own = own[own['owner_type'].astype(str).str.upper().str.contains('ORGANIZATION', na=False)] if 'owner_type' in own.columns else own
        last = own.sort_values('snapshot_date').groupby(['ccn_std', 'owner_name']).agg(assoc=('assoc', 'last')).reset_index()
        for r in S[S.counts & S.evd.notna() & S.ccn.notna()].itertuples():
            g = last[(last.ccn_std == r.ccn) & last.assoc.notna()]
            for x in g.itertuples():
                if r.evd - pd.DateOffset(months=24) <= x.assoc < r.evd:
                    rows.append((r.cec, r.facility, None, nice(x.owner_name), f'Suggestion: CMS lists this company in managerial control since {x.assoc.date()}, before the {r.src.lower()} of {r.evd.date()}. Confirm with the CEC application or an APRA request, then enter the start date.'))
    r0 = 5
    for cid, fac_, start, who, src in rows:
        _put(ws, r0, 1, cid, F_INPUT, fill=FILL_INPUT); _put(ws, r0, 2, fac_, F_INPUT, fill=FILL_INPUT); _put(ws, r0, 3, start, F_INPUT, 'yyyy-mm-dd', FILL_INPUT)
        _put(ws, r0, 4, who, F_INPUT, fill=FILL_INPUT); _put(ws, r0, 5, src, F_INPUT, fill=FILL_INPUT, align=Alignment(wrap_text=True, vertical='top')); r0 += 1
    for rr in range(r0, r0 + 20): [_put(ws, rr, c, None, F_INPUT, 'yyyy-mm-dd' if c == 3 else None, FILL_INPUT) for c in range(1, 6)]
    ws.freeze_panes = 'A5'; ctx['n_mgmt'] = len(rows)

def write_sale_timing(wb, cec, M, ctx):
    ws = wb.create_sheet('Sale Timing'); ws.sheet_properties.tabColor = 'BF9000'; last = max(cec['tracker_last_row'], 159); ctx['st_last'] = last
    lc = cec['log_cols']; idc, tyc = L(lc['id']), L(lc['type']); LR = 500
    _put(ws, 1, 1, 'Sale Timing: which CMS snapshot each Post-Sale Tracker row uses (row numbers match the tracker)', F_TITLE)
    heads = [('CEC ID', 9), ('Facility', 38), ('CCN', 9), ('Event date used', 13), ('Date basis', 20), ('Counts as an ownership change (1/0)', 13), ('Before: last CMS snapshot before the event', 15), ('Year 1: first snapshot 12+ months after', 15), ('Year 2: first snapshot 24+ months after', 15),
             ('Next ownership change at this CCN', 15), ('Exit: last snapshot before the next change (or latest)', 16), ('CMS first shows an operator change', 15), ('Days from event to CMS operator change (max 730)', 14), ('Notes', 60), ('Override (from the Event Overrides tab)', 16)]
    for j, (h, w) in enumerate(heads): _hdr(ws, 2, j + 1, h); ws.column_dimensions[L(j + 1)].width = w
    ws.row_dimensions[2].height = 66
    _put(ws, 1, 17, 'Latest CMS snapshot', F_BOLD); _put(ws, 1, 18, f"=MAX({_mdr('A')})", F_BOLD, 'yyyy-mm-dd')
    _put(ws, 2, 17, 'Star ratings comparable from', F_BOLD); _put(ws, 2, 18, dt.datetime.strptime(M['cfg']['series_start'], '%Y-%m-%d'), F_INPUT, 'yyyy-mm-dd', FILL_INPUT)
    _put(ws, 3, 17, 'Citation / fine records reliable from', F_BOLD); _put(ws, 3, 18, f"=EDATE(DATE({int(M['snaps'][0][:4])},{int(M['snaps'][0][5:7])},{int(M['snaps'][0][8:10])}),-36)", F_BOLD, 'yyyy-mm-dd')
    ws.column_dimensions['Q'].width = 34; ws.column_dimensions['R'].width = 13
    cnt = lambda t: f"COUNTIFS({LOG}${idc}$2:${idc}${LR},$A{{r}},{LOG}${tyc}$2:${tyc}${LR},\"{t}\")"
    A, Bc, Cc, Dc = _mdr('A'), _mdr('B'), None, None
    for r in range(3, last + 1):
        f = {}
        f['A'] = f'=IF({TRK}A{r}="","",{TRK}A{r})'; f['B'] = f'=IF(A{r}="","",{TRK}B{r})'
        f['C'] = f"=IF(A{r}=\"\",\"\",IFERROR(INDEX('CCN Map'!$B$5:$B$400,MATCH(B{r},'CCN Map'!$A$5:$A$400,0))&\"\",\"\"))"
        f['D'] = f'=IF(A{r}="","",IF(ISNUMBER({TRK}I{r}),{TRK}I{r},IF(ISNUMBER({TRK}F{r}),{TRK}F{r},IFERROR(DATE(VALUE(RIGHT({TRK}F{r},4)),VALUE(LEFT({TRK}F{r},2)),VALUE(MID({TRK}F{r},4,2))),""))))'
        f['E'] = f'=IF(D{r}="","",IF(ISNUMBER({TRK}I{r}),"CMS takeover date","RIDOH decision date"))'
        f['O'] = f"=IF(A{r}=\"\",\"\",IFERROR(INDEX('Event Overrides'!$B$5:$B$200,MATCH(A{r},'Event Overrides'!$A$5:$A$200,0))&\"\",\"\"))"
        f['F'] = f'=IF(A{r}="","",IF($O{r}<>"",IF(OR($O{r}="sale",$O{r}="parent company change"),1,0),IF(AND({TRK}E{r}<>"Pending",(' + '+'.join(cnt(t).format(r=r) for t in SALE_TYPES) + ')>0),1,0)))'
        f['G'] = f'=IF(OR($A{r}="",$C{r}="",$D{r}="",$O{r}="closure",$O{r}="not a sale"),"",IF(COUNTIFS({Bc},$C{r},{A},"<"&$D{r})=0,"",_xlfn.MAXIFS({A},{Bc},$C{r},{A},"<"&$D{r})))'
        f['J'] = f'=IF(OR($A{r}="",$C{r}="",$D{r}=""),"",IF(COUNTIFS($C$3:$C${last},$C{r},$D$3:$D${last},">"&($D{r}+1),$F$3:$F${last},1)=0,"",_xlfn.MINIFS($D$3:$D${last},$C$3:$C${last},$C{r},$D$3:$D${last},">"&($D{r}+1),$F$3:$F${last},1)))'
        for col, m in (('H', 12), ('I', 24)):
            tgt = f'EDATE($D{r},{m})'; mn = f'_xlfn.MINIFS({A},{Bc},$C{r},{A},">="&{tgt})'
            f[col] = f'=IF(OR($A{r}="",$C{r}="",$D{r}="",$O{r}="closure",$O{r}="not a sale"),"",IF(COUNTIFS({Bc},$C{r},{A},">="&{tgt})=0,"",IF({mn}>{tgt}+92,"",IF(AND($J{r}<>"",{mn}>=$J{r}),"",{mn}))))'
        f['K'] = f'=IF($G{r}="","",IF($J{r}="",_xlfn.MAXIFS({A},{Bc},$C{r}),IF(COUNTIFS({Bc},$C{r},{A},"<"&$J{r})=0,"",_xlfn.MAXIFS({A},{Bc},$C{r},{A},"<"&$J{r}))))'
        oc = lambda col: f"'Ownership Changes'!${col}$5:${col}$3000"; cnd = f'{oc("C")},$C{r},{oc("D")},"Operator change",{oc("A")},">="&EDATE($D{r},-6)'
        f['L'] = f'=IF(OR($C{r}="",$D{r}="",$O{r}="closure",$O{r}="not a sale"),"",IF(COUNTIFS({cnd})=0,"",_xlfn.MINIFS({oc("A")},{cnd})))'
        f['M'] = f'=IF($L{r}="","",IF($L{r}-$D{r}>730,"",$L{r}-$D{r}))'
        f['N'] = f'=IF(A{r}="","",IF($O{r}<>"","Override: "&$O{r}&" (see the Event Overrides tab)",IF(C{r}="","No CCN, so no CMS numbers",IF(D{r}="","No event date",IF(G{r}="","No CMS rating before the event (event before "&TEXT($R$2,"mmm yyyy")&", or facility not in the CMS files)",IF(H{r}="","Year 1 not available (too recent, left CMS files, or another ownership change came first)",IF(I{r}="","Year 2 not available","")))))))'
        for col, v in f.items():
            fm = 'yyyy-mm-dd' if col in 'DGHIJKL' else ('0' if col in 'FM' else ('@' if col in 'CO' else None))
            _put(ws, r, ord(col) - 64, v, F_NORM, fm, FILL_STAT, CTR if col not in 'BN' else None)
    ws.freeze_panes = 'C3'

BLOCKS_POINT = [(14, 'E', ['G', 'H', 'I', 'K']), (19, 'F', ['G', 'H', 'I']), (23, 'K', ['G', 'H', 'I']), (27, 'L', ['G', 'H', 'I']), (31, 'O', ['G', 'H', 'I'])]
def write_tracker(wb, cec, ctx):
    ws = wb['Post-Sale Tracker']; last = ctx['st_last']; n_filled = 0; n_skipped = 0
    CIT, PEN = "'CMS Citations'!", "'CMS Penalties'!"
    def cit(col): return f"{CIT}${col}$2:${col}$30000"
    def pen(col): return f"{PEN}${col}$2:${col}$30000"
    def put_formula(r, c, f):
        nonlocal n_filled, n_skipped
        cell = ws.cell(row=r, column=c); v = cell.value
        if v is None or (isinstance(v, str) and v.startswith('=') and "'Sale Timing'!" in v):
            cell.value = f; cell.fill = FILL_STAT; f0 = cell.font; cell.font = Font(name=f0.name or FONT, size=f0.size or 10, color='000000'); n_filled += 1
        else: n_skipped += 1
    for r in range(3, last + 1):
        for c0, mdcol, snapcols in BLOCKS_POINT:
            for k, sc in enumerate(snapcols):
                sd = f"{ST}${sc}{r}"; key = f"{ST}$C{r}&\"|\"&{sd}"; m = f"MATCH({key},{_mdr(MD_COL['key'])},0)"; ix = f"INDEX({_mdr(mdcol)},{m})"
                put_formula(r, c0 + k, f'=IF({sd}="","",IFERROR(IF({ix}="","",{ix}),""))')
        D_ = f"{ST}$D{r}"; C_ = f"{ST}$C{r}"
        for c0, kind in ((35, 'all'), (39, 'serious'), (43, 'fines')):
            wins = [('G', f'EDATE({D_},-12)', D_, f'EDATE({D_},-12)>={ST}$R$3'), ('H', D_, f'EDATE({D_},12)', None), ('I', f'EDATE({D_},12)', f'EDATE({D_},24)', None)]
            for k, (sc, lo, hi, extra) in enumerate(wins):
                sd = f"{ST}${sc}{r}"
                if kind == 'fines': body = f'SUMIFS({pen("D")},{pen("A")},{C_},{pen("C")},"Fine",{pen("B")},">="&{lo},{pen("B")},"<"&{hi})'
                elif kind == 'serious': body = f'COUNTIFS({cit("A")},{C_},{cit("F")},1,{cit("B")},">="&{lo},{cit("B")},"<"&{hi})'
                else: body = f'COUNTIFS({cit("A")},{C_},{cit("B")},">="&{lo},{cit("B")},"<"&{hi})'
                put_formula(r, c0 + k, (f'=IF({sd}="","",IF({extra},{body},""))' if extra else f'=IF({sd}="","",{body})'))
    ctx['trk_filled'], ctx['trk_skipped'] = n_filled, n_skipped

def write_cec_summary_cols(wb, mp, fn, ccn_of, cec, M, ctx):
    ws = wb['Facility Summary']; hdr_cell = ws.cell(row=1, column=1); c0 = ws.max_column + 1
    # remove columns from an earlier run (same headers), then add
    heads = ['CCN (from CCN Map)', 'Former / other names (CMS files + CEC record)', 'CMS-detected ownership / operator changes (2019+)', 'Latest CMS overall rating']
    existing = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}
    cols = [existing.get(h, c0 + i) for i, h in enumerate(heads)]
    for i, h in enumerate(heads):
        c = cols[i]; cell = ws.cell(row=1, column=c, value=h); cell.font = Font(name=FONT, size=10, bold=True, color='FFFFFF'); cell.fill = PatternFill('solid', start_color='1F3864'); cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.column_dimensions[L(c)].width = [11, 60, 16, 13][i]
    for r in range(2, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if not name: continue
        ccn = ccn_of.get(str(name).strip()); items = fn.get(ccn or ('nocc:' + str(name).strip()), []) if not (ccn in M['fac'].index) else fn.get(ccn, [])
        vals = [f"=IFERROR(INDEX('CCN Map'!$B$5:$B$400,MATCH($A{r},'CCN Map'!$A$5:$A$400,0))&\"\",\"\")", names_text(items) if items else None,
                f"=IF({L(cols[0])}{r}=\"\",\"\",COUNTIFS('Ownership Changes'!$C$5:$C$3000,{L(cols[0])}{r},'Ownership Changes'!$D$5:$D$3000,\"Operator change\")+COUNTIFS('Ownership Changes'!$C$5:$C$3000,{L(cols[0])}{r},'Ownership Changes'!$D$5:$D$3000,\"Change of ownership (CMS flag)\"))",
                f"=IF({L(cols[0])}{r}=\"\",\"\",IFERROR(INDEX('CMS Summary'!$M$5:$M$400,MATCH({L(cols[0])}{r},'CMS Summary'!$B$5:$B$400,0)),\"\"))"]
        for i, v in enumerate(vals):
            cell = ws.cell(row=r, column=cols[i], value=v); cell.font = Font(name=FONT, size=10); cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal=('center' if i != 1 else None)); cell.fill = PatternFill('solid', start_color='F2F2F2')
    ws.auto_filter.ref = f'A1:{L(max(cols))}{ws.max_row}'

# ------------------------------------------------------------------ read me + main entry point
def write_read_me(wb, M, S, cec, mp, cfg, ctx):
    ws = wb.create_sheet('CMS Read Me'); ws.sheet_properties.tabColor = '1F3864'; ws.sheet_view.showGridLines = False; ws.column_dimensions['A'].width = 30; ws.column_dimensions['B'].width = 130
    r = [1]
    def line(a, b='', bold=False):
        _put(ws, r[0], 1, a, F_BOLD if (bold or b) else F_NORM, align=Alignment(vertical='top', wrap_text=True)); _put(ws, r[0], 2, b, F_NORM, align=Alignment(vertical='top', wrap_text=True)); r[0] += 1
    def sec(t): r[0] += 1; _put(ws, r[0], 1, t, Font(name=FONT, size=11, bold=True, color='FFFFFF'), fill=FILL_ANN); _put(ws, r[0], 2, None, fill=FILL_ANN); r[0] += 1
    _put(ws, 1, 1, 'CMS data added to the CEC workbook', F_TITLE); r[0] = 2
    line('', f"CMS Care Compare monthly files for Rhode Island, {lab_(M['snaps'][0])} to {lab_(M['latest'])} ({len(M['snaps'])} files), combined with the RIDOH CEC tabs. Built {dt.date.today().isoformat()}. Facilities in CMS data: {len(M['fac'])} ({int(M['fac'].active.sum())} active). CEC tabs (Read Me, Facility Summary, CEC Log, Post-Sale Tracker, Chain Ratings, CMS Current, Cost Reports) are yours; the tabs listed below as generated are rebuilt on every update.")
    sec('Update each quarter')
    for a, b in [('1. Download', 'Download the newest CMS nursing-home archive zip(s) and put them in the same Drive folder as before.'), ('2. Run the notebook', 'Open cms_ri_extract.ipynb in Colab and choose Runtime > Run all. It only reads files it has not seen before, then rebuilds this workbook and the public page (about a few minutes).'),
                 ('3. Your edits are kept', 'Anything you typed in the CEC tabs, the CCN Map (yellow cells) and the Post-Sale Tracker is kept. Tracker cells that hold a number you typed are never overwritten; grey cells are formulas and refresh.'), ('4. New CEC rows', 'Add new rows to CEC Log and Post-Sale Tracker as usual (fill in the CEC ID, facility, dates and outcome). New facility names are added to the CCN Map automatically; type a CCN there if CMS has one.'),
                 ('5. Publish', 'Upload public/index.html to your web page (replacing the old file).')]: line(a, b)
    sec('Tabs the update generates (do not type in these)')
    for a, b in [('CMS Summary', 'One row per CMS facility: latest ratings, staffing, citations, fines, former names, CMS-detected and CEC changes.'), ('RI Trends (Annual / Quarterly)', 'Statewide averages and charts.'),
                 ('Overall Rating ... Occupancy', 'Facility-by-quarter grids. Quarterly columns are shown by default; the outline buttons at the top-left collapse them to annual.'), ('Citations - All / Serious / Abuse, Fines', 'By quarter, adding up to calendar years.'),
                 ('Ownership Changes', 'Every CEC decision and every change seen in the CMS files, newest first, strong signals first.'), ('Operator by Quarter', 'Who CMS lists as the operator. Amber = operator changed; purple = a CEC decision falls in that period.'), ('Other Changes', 'Name, status, beds, rating changes, facilities that left the CMS data.'),
                 ('Sale Timing', 'For each Post-Sale Tracker row: the event date and which CMS snapshots are used for Before / Year 1 / Year 2 / Exit. Rows line up with the tracker.'), ('CMS Monthly Data, CMS Citations, CMS Penalties', 'The underlying records the formulas read.')]: line(a, b)
    sec('Tabs you edit')
    line('CCN Map', 'Facility name to CCN. Blank CCN = no CCN exists (no CMS numbers are shown for that facility, and none are made up). Your entries always win over automatic matches.')
    line('Management Agreements', 'When a buyer took over running a home, which is often months before the sale is approved. Blank rows seeded from CMS\'s owner list are suggestions to confirm; RIDOH holds the contracts under 216-RICR-40-10-1 s.1.7.2 but does not publish them, so request them under the Access to Public Records Act. Dates entered here appear on the public page timeline and beside each sale.')
    line('Event Overrides', 'Use when a CEC row is not really a sale (for example a home that closed instead). The row then drops out of the sales table, the before/after comparison and the tracker numbers, but stays on the facility timeline with the label you give. Seeded with CEC-077 (BalloLife), which closed in September 2021.')
    line('Facility Summary (CEC)', 'Four columns were added at the right: CCN, former / other names, CMS-detected changes and latest CMS rating. Column B ("Other names in the record") is unchanged.')
    line('Post-Sale Tracker', 'The rating, staffing, turnover, citation and fine blocks are now grey formula cells that read the CMS tabs. If you type a number over one of them, that number stays and the formula is not restored on later updates. Ombudsman columns are still manual.')
    sec('How the Post-Sale Tracker numbers are chosen')
    for a, b in [('Event date', 'The CMS takeover date if the tracker has one (column I), otherwise the RIDOH decision date (column F).'), ('Before', 'The last CMS monthly file before the event date. Note: the tracker headers for staffing, turnover, citations and fines say "12 mo prior"; ratings, staffing hours, turnover use this last-file-before value, and citation and fine counts use the 12 months before the event.'),
                 ('Year 1 / Year 2', 'The first CMS file at least 12 / 24 months after the event (within about 3 months of that date). Left blank if the next ownership change at that CCN came first, the facility left the CMS files, or the date has not arrived yet.'),
                 ('At exit, or current', 'The last CMS file before the next ownership change at the same CCN; if there is none, the latest file.'), ('Citations and fines', 'Counted by inspection / penalty date in the 12 months before, months 0-12 after (Year 1) and months 12-24 after (Year 2).'),
                 ('CMS shows an operator change', "The first CMS file, up to 6 months before the event or later (limit 730 days), whose operator (legal business name) differs from the earlier one. CMS usually shows a new operator months after the state's decision.")]: line(a, b)
    sec('Cautions')
    line('Rating method changes', f"CMS changed how star ratings are calculated in April 2019: the statewide average fell from 3.8 to 3.35 in one month. Ratings before {lab_(cfg['series_start'])} are not comparable with later ones, so the tracker and the public page start from {lab_(cfg['series_start'])}; sales before then have no Before value (they need pre-2019 CMS archives and a different method). CMS also adjusted staffing and quality measures later.")
    line('Sales are not causes', 'Before/after numbers show what happened, not why. Many homes were already declining before sale. The public page compares sold homes with homes that were not sold over the same dates.')
    line('CMS timing', 'CMS files lag events. Owner lists are incomplete or re-coded at times (officer/director roles were re-coded in Dec 2024, Jun 2025 and May 2026 and are excluded). Citation and fine records are the latest version of each record CMS still lists; CMS lists about 3 years, so counts before early 2016 are not available.')
    line('Special Focus Facilities', 'CMS publishes no star ratings for homes in its Special Focus Facility (SFF) program (footnote 18 in the provider file). Those months are recorded as 0 stars, which is not a CMS score, so SFF status counts in averages, in the share of low-rated homes and in before/after comparisons. Example: Hebert Nursing Home is 0 for all of 2022. To leave them blank instead, set sff_as_zero=False in the notebook settings.')
    hr = M['hprd_rows']
    if len(hr):
        parts = []
        for c, g in hr.groupby('ccn_std'):
            nm = M['fac'].loc[c, 'name'] if c in M['fac'].index else c
            parts.append(f"{nm}: {len(g)} monthly files ({lab_(g.snapshot_date.min())} to {lab_(g.snapshot_date.max())}), reported {g.iloc[:, 2].min():.2f} hours")
        line('Staffing hours left blank', f"Reported nurse hours below {cfg.get('min_plausible_hprd', 1.5)} per resident per day are treated as a payroll-reporting problem, not real staffing, and are blank in the hours grids, averages and comparisons. CMS marked these months with footnote 12 on the staffing rating and gave the home 1 staffing star; CMS's published rules downgrade a home to 1 star when it fails to submit payroll data, fails an audit, or reports 7 or more days with no RN, and the file does not say which applied. " + '; '.join(parts) + '.')
    tr_ = M['turn_rows']
    if len(tr_):
        parts = []
        for c, g in tr_.groupby('ccn_std'):
            nm = M['fac'].loc[c, 'name'] if c in M['fac'].index else c
            parts.append(f"{nm}: {len(g)} monthly files ({lab_(g.snapshot_date.min())} to {lab_(g.snapshot_date.max())}), {g.iloc[:, 2].min():.0f}% to {g.iloc[:, 2].max():.0f}%")
        line('Turnover left blank', f"Nursing staff turnover at or above {cfg.get('max_plausible_turnover', 90)}% is treated as CMS's lowest-possible score for missing payroll data, not real turnover, and is blank here. CMS's published rules assign that lowest score for the affected period. " + '; '.join(parts) + '.')
    if M['ds_behind']:
        line('Some CMS files are older than others', 'The newest monthly file does not always include every CMS dataset. In this build the latest data is: ' + '; '.join(f'{k} {lab_(v)}' for k, v in M['ds_latest'].items()) + '. Anything drawn from an older file (for example citation and fine counts) stops at that date, even though the ratings run later.')
    for t, x in cfg['data_cautions']: line(t, x)
    sec('Why measures can point in different directions')
    for a, b in [('Staffing up, stars down', "The overall star rating starts from the health inspection rating and is then adjusted up or down by the staffing and quality-measure ratings. A home can add staff and still lose stars if its inspection results or quality measures got worse. Check the Health Inspection and QM Rating tabs for the same quarters."),
                 ('Stars down, serious citations and fines down', "The health inspection rating is built from the last three years of inspections, with the most recent weighted most, and it counts all deficiencies by scope and severity, not just serious ones (G-L). The citation tabs here count by inspection date, so a rating can fall in a year with fewer serious citations. Fines are not part of the star rating at all."),
                 ('Hours vs staffing stars', "The hours columns are CMS's reported (unadjusted) hours per resident per day. The staffing star compares case-mix-adjusted hours with national thresholds, so the two do not always move together.")]: line(a, b)
    sec('Public page')
    line('public/index.html', 'A single self-contained web page (no server needed): statewide trends, the sold-vs-not-sold comparison, a searchable table of every sale, a facility lookup (searchable by former names) with ownership timelines and sale markers, owner groups, and CSV downloads. It shows companies only by default (no individuals\' names); set public_show_principals=True in the notebook to change that. Homes with no CCN appear with their state review record only.')
    ws.sheet_view.zoomScale = 100

def run(out_dir, cec_path=None, config=None, std_dir=None, verbose=True):
    import openpyxl; cfg = {**DEFAULT_CONFIG, **(config or {})}; std_dir = std_dir or os.path.join(out_dir, 'combined_standardized'); say = print if verbose else (lambda *a, **k: None)
    out_xlsx = os.path.join(out_dir, 'RI_Nursing_Homes_CMS_and_CEC.xlsx'); pub_dir = os.path.join(out_dir, 'public'); os.makedirs(pub_dir, exist_ok=True)
    src = cec_path or (out_xlsx if os.path.exists(out_xlsx) else None)
    if src is None:
        import glob; c = sorted(glob.glob(os.path.join(out_dir, 'RI_Nursing_Home_CECs*.xlsx')) + glob.glob(os.path.join(out_dir, '*CEC*.xlsx')))
        if not c: raise FileNotFoundError('Put your CEC workbook (RI_Nursing_Home_CECs_....xlsx) in the output folder or pass cec_path=')
        src = c[-1]
    say('Reading CMS files...'); inp = load_cms(std_dir); M = build_events(build_model(inp, cfg), inp)
    say(f'  {len(M["snaps"])} monthly files, {len(M["fac"])} facilities, latest {M["latest"]}'); say('Reading CEC workbook:', os.path.basename(src)); wb = openpyxl.load_workbook(src); cec = read_cec(wb)
    if not cec['overrides_present']:
        cec['overrides'] = {c: dict(treat=t, date=pd.Timestamp(d) if d else None, label=l, note=n) for c, t, d, l, n in cfg.get('event_overrides_seed', [])}
    mp = build_ccn_map(cec, M); ccn_of = {r['name']: r['ccn'] for r in mp}
    for c_, (lab_c, _d_c, _yo) in closed_notes(mp, cfg).items():   # closed homes CMS still lists
        if c_ in M['fac'].index: M['fac'].loc[c_, 'active'] = False; M['fac'].loc[c_, 'status'] = lab_c
    S, val, snaps_by = compute_sales(M, cec, ccn_of); fn = former_names(M, cec, ccn_of)
    say(f'  {len(cec["tracker"])} tracker rows, {sum(1 for r in mp if not r["ccn"])} facilities with no CCN, {int(S.before.notna().sum())} sales with a CMS "before" snapshot')
    D = build_public_data(M, S, val, cec, mp, ccn_of, fn, cfg); page = build_page(D, os.path.join(pub_dir, 'index.html')); say('Public page:', page)
    say('Writing workbook...')
    for n in MANAGED:
        if n in wb.sheetnames: del wb[n]
    ctx = write_cms_tabs(wb, M, S, cec, mp, ccn_of, fn, cfg); write_logs_and_summary(wb, M, S, cec, mp, ccn_of, fn, cfg, ctx); write_trends(wb, M, ctx); write_data_tabs(wb, M, ctx)
    write_ccn_map(wb, mp, ctx); write_event_overrides(wb, cec, ctx, cfg); write_mgmt_agreements(wb, cec, M, S, inp['own'], ctx); write_sale_timing(wb, cec, M, ctx); write_tracker(wb, cec, ctx); write_cec_summary_cols(wb, mp, fn, ccn_of, cec, M, ctx); write_read_me(wb, M, S, cec, mp, cfg, ctx)
    cec_tabs = [n for n in wb.sheetnames if n not in MANAGED and n not in ('CCN Map', 'Event Overrides', 'Management Agreements')]; front = []
    for n in cec_tabs:
        front.append(n)
        if n == 'Read Me': front.append('CMS Read Me')
    if 'CMS Read Me' not in front: front.insert(0, 'CMS Read Me')
    if 'Chain Ratings' in front: i = front.index('Chain Ratings') + 1
    else: i = len(front)
    front[i:i] = ['Sale Timing', 'CCN Map', 'Management Agreements', 'Event Overrides']
    rest = [n for n in MANAGED if n in wb.sheetnames and n not in front]; order = front + rest
    wb._sheets = [wb[n] for n in order]; wb.active = 0
    for w in wb.worksheets: w.sheet_view.tabSelected = (w.title == order[0])
    wb.calculation.fullCalcOnLoad = True; wb.save(out_xlsx); say('Workbook:', out_xlsx)
    return dict(workbook=out_xlsx, page=page, M=M, S=S, D=D, ctx=ctx, val=val)

# ------------------------------------------------------------------ public web page (single self-contained HTML file)
PAGE_TEMPLATE = r'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>__TITLE__</title>
<link rel="icon" href="abcri-logo.png">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Public+Sans:wght@400;500;600;700;800&display=swap">
<style>
:root{--bg:#F6F3EA;--fg:#1E2A32;--muted:#4E5860;--card:#EAE5D7;--line:#8F8570;--accent:#1E2A32;--gold:#93701F;--accent-soft:#E2DCCA;--sold:#92471F;--ctrl:#4E5860;--teal:#2F6B7A;--warn:#E9E1CC;--warn-line:#93701F;
--s0:#8C8575;--s1:#6E1A20;--s2:#C4612F;--s3:#D4A82E;--s4:#5F9E6E;--s5:#1F5137;--chip-fg:#fff;--link:#1E2A32;--chart-bg:#ffffff;--grid:#D9D2C2;
box-sizing:border-box;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#191E1A;--fg:#E8E4D8;--muted:#A2A99C;--card:#20251F;--line:#6E6F5A;--accent:#E8E4D8;--gold:#C9A44C;--accent-soft:#2E342D;--sold:#C98757;--ctrl:#A2A99C;--teal:#7FB6C4;--warn:#302219;--warn-line:#C9A44C;--link:#E8E4D8;--s0:#6E6F5A;--s1:#D98891;--s2:#C98757;--s3:#C9A44C;--s4:#82AC94;--s5:#5FA37C;--chip-fg:#17201A;--chart-bg:#171C16;--grid:#333A31}}
:root[data-theme="dark"]{--bg:#191E1A;--fg:#E8E4D8;--muted:#A2A99C;--card:#20251F;--line:#6E6F5A;--accent:#E8E4D8;--gold:#C9A44C;--accent-soft:#2E342D;--sold:#C98757;--ctrl:#A2A99C;--teal:#7FB6C4;--warn:#302219;--warn-line:#C9A44C;--link:#E8E4D8;--s0:#6E6F5A;--s1:#D98891;--s2:#C98757;--s3:#C9A44C;--s4:#82AC94;--s5:#5FA37C;--chip-fg:#17201A;--chart-bg:#171C16;--grid:#333A31}
html{scroll-scroll-behavior:smooth;scroll-padding-top:env(safe-area-inset-top,0px)}
*,*::before,*::after{box-sizing:inherit}
body{margin:0;background:var(--bg);color:var(--fg);font:16px/1.55 "Public Sans",system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif}
a{color:var(--link)}
.wrap{max-width:1080px;margin:0 auto;padding:0 18px}
header.masthead{background:var(--bg);color:var(--fg);padding:34px 0 0}
.site-logo-link{display:inline-block;text-decoration:none}
.site-logo{height:72px;width:auto;display:block;margin:0 0 20px}
@media (max-width:560px){.site-logo{height:56px}header.masthead{padding-top:22px}}
header.masthead h1{margin:0 0 8px;font-size:clamp(1.5rem,4.2vw,2.25rem);line-height:1.14;font-weight:800;letter-spacing:-.015em}
header.masthead p{margin:0;max-width:760px;color:var(--muted);font-size:1.02rem}
header.masthead .wrap{padding-bottom:16px;border-bottom:1px solid var(--line)}
.asof{display:inline-block;margin-top:14px;padding:3px 10px;border:1px solid var(--line);background:var(--card);color:var(--fg);border-radius:2px;font-size:.85rem;font-weight:600}
nav.jump{position:sticky;top:0;background:var(--card);border-bottom:1px solid var(--line);z-index:5;padding-top:env(safe-area-inset-top,0px)}
nav.jump .wrap{display:flex;gap:6px;overflow-x:auto;padding-top:8px;padding-bottom:8px}
nav.jump a{white-space:nowrap;text-decoration:none;padding:5px 12px;border-radius:999px;font-size:.92rem;color:var(--fg);border:1px solid var(--line)}
nav.jump a:hover,nav.jump a:focus-visible{background:var(--accent-soft)}
section{padding:34px 0 8px}
h2{font-size:1.5rem;margin:0 0 6px;line-height:1.25}
h3{font-size:1.15rem;margin:22px 0 8px}
.lede{color:var(--muted);margin:0 0 16px;max-width:760px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;margin:14px 0 6px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px}
.card .big{font-size:2rem;font-weight:700;line-height:1.1}
.card .lbl{font-size:.9rem;color:var(--muted);margin-top:2px}
.card .cmp{font-size:.88rem;margin-top:6px}
.panel{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px;margin:12px 0}
.tabs{display:flex;flex-wrap:wrap;gap:6px;margin:8px 0 10px}
button.tab,.btn{font:inherit;font-size:.92rem;padding:6px 12px;border-radius:8px;border:1px solid var(--line);background:var(--card);color:var(--fg);cursor:pointer}
button.tab[aria-pressed="true"]{background:var(--accent);color:#fff;border-color:var(--accent)}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]) button.tab[aria-pressed="true"]{background:#3a5a97;color:#fff;border-color:#3a5a97}}
button:focus-visible,a:focus-visible,input:focus-visible,select:focus-visible,summary:focus-visible{outline:2px solid var(--gold);outline-offset:2px}
.chartbox{position:relative;width:100%;background:var(--chart-bg);border:1px solid var(--line);border-radius:3px;padding:10px 12px}
.chartbox svg{width:100%;height:auto;display:block;overflow:visible}
.tip{position:absolute;pointer-events:none;background:var(--card);color:var(--fg);border:1px solid var(--line);border-radius:8px;padding:6px 9px;font-size:.82rem;box-shadow:0 4px 14px rgba(0,0,0,.18);z-index:3;display:none;min-width:140px}
.tip b{display:block;margin-bottom:2px}
.legend{display:flex;flex-wrap:wrap;gap:6px 16px;font-size:.86rem;margin:6px 0 0;color:var(--muted)}
.legend i{display:inline-block;width:18px;height:0;border-top:3px solid;vertical-align:middle;margin-right:5px}
.legend i.dash{border-top-style:dashed}
.legend i.sq{width:13px;height:13px;border:0;border-radius:3px;margin-right:6px;vertical-align:-2px}
.legend i.dot{border-top-style:dotted;border-top-width:4px}
ol.timeline .kind.mgmt{background:#fdf0cf;color:#6b4e07;border-color:#e3c76a}
ol.timeline .kind.closed{background:#e4e4e4;color:#222;border-color:#999}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]) ol.timeline .kind.closed{background:#3a3f47;color:#eee;border-color:#6b7480}}
.legend i.vert{width:0;height:14px;border-top:0;border-left:3px solid;margin-right:7px}.legend i.vert.dash{border-left-style:dashed}
.note{background:var(--warn);border:1px solid var(--warn-line);border-radius:10px;padding:10px 14px;font-size:.92rem;margin:12px 0}
.callout{background:var(--accent-soft);border-left:5px solid var(--accent);border-radius:8px;padding:12px 16px;margin:12px 0}
.callout p{margin:0 0 6px}.callout p:last-child{margin:0}
.grid2{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:14px}
.small{font-size:.86rem;color:var(--muted)}
table{border-collapse:collapse;width:100%;font-size:.9rem}
caption{text-align:left;color:var(--muted);font-size:.86rem;padding:4px 0}
th,td{padding:7px 9px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}
th{position:sticky;top:0;background:var(--card);font-weight:600;white-space:nowrap}
th button{all:unset;cursor:pointer;font-weight:600}
th button:focus-visible{outline:2px solid var(--gold)}
td.num,th.num{text-align:center;white-space:nowrap}
.tblwrap{overflow-x:auto;max-height:640px;border:1px solid var(--line);border-radius:10px;background:var(--card)}
.chip{display:inline-block;min-width:1.6em;padding:1px 6px;border-radius:6px;text-align:center;font-weight:700;color:var(--chip-fg)}
.chip.s0{background:var(--s0);color:#fff}.chip.s3,.chip.s4{color:#17201A}
.chip.s1{background:var(--s1)}.chip.s2{background:var(--s2)}.chip.s3{background:var(--s3)}.chip.s4{background:var(--s4)}.chip.s5{background:var(--s5)}
.chip.na{background:transparent;color:var(--muted);font-weight:400}
.arrow{color:var(--muted);margin:0 2px}
.linkbtn{all:unset;color:var(--link);cursor:pointer;text-decoration:underline}
.linkbtn:focus-visible{outline:2px solid var(--gold)}
.row{display:flex;flex-wrap:wrap;gap:10px 18px;align-items:center;margin:8px 0}
input[type=search],input[type=text],select{font:inherit;padding:8px 10px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--fg);min-width:240px;max-width:100%}
label.chk{display:flex;gap:6px;align-items:center;font-size:.92rem}
ul.results{list-style:none;margin:4px 0 0;padding:0;border:1px solid var(--line);border-radius:8px;background:var(--card);max-height:260px;overflow:auto}
ul.results li button{all:unset;display:block;width:100%;padding:8px 12px;cursor:pointer;box-sizing:border-box}
ul.results li button:hover,ul.results li button:focus-visible{background:var(--accent-soft)}
ul.results .sub{color:var(--muted);font-size:.84rem}
.badges{display:flex;flex-wrap:wrap;gap:8px;margin:8px 0}
.badge{border:1px solid var(--line);border-radius:10px;padding:6px 10px;background:var(--card);font-size:.88rem}
.badge b{display:block;font-size:1.1rem}
ol.timeline{list-style:none;padding:0;margin:8px 0}
ol.timeline li{display:grid;grid-template-columns:96px 1fr;gap:10px;padding:8px 0;border-bottom:1px solid var(--line)}
ol.timeline .kind{display:inline-block;font-size:.75rem;padding:1px 7px;border-radius:999px;border:1px solid var(--line);margin-right:6px}
ol.timeline .kind.sale{background:#fde6d7;color:#7a2a06;border-color:#e9b492}.kind.parent{background:#e3ebfb;color:#1b3b7a;border-color:#a9bde6}.kind.cms{background:#eef0f2;color:#3d4753}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]) ol.timeline .kind.sale{background:#4a2814;color:#ffd2b5}:root:not([data-theme="light"]) .kind.parent{background:#1f2f52;color:#c9d9ff}:root:not([data-theme="light"]) .kind.cms{background:#2a323d;color:#d0d7e0}}
footer{margin:44px 0 30px;padding-top:18px;border-top:1px solid var(--line);color:var(--muted);font-size:.9rem}
details{margin:8px 0}summary{cursor:pointer;font-weight:600}
.hidden{display:none!important}
@media (max-width:640px){ol.timeline li{grid-template-columns:1fr}.card .big{font-size:1.7rem}}
/* ABC-RI house style: small square corners, thin rules, no soft shadows */
.card,.panel,.tblwrap,.note,.callout,.badge,.tip,ul.results,input[type=search],input[type=text],select,button.tab,.btn,.chip,ol.timeline .kind,.sq{border-radius:3px}
.tip,.sugg{box-shadow:0 3px 10px rgba(30,42,50,.18)}
nav.jump a{border-radius:2px;font-weight:500}
nav.jump a.on,button.tab[aria-pressed="true"]{font-weight:600}
h1,h2,h3{font-weight:700;letter-spacing:-.005em}
h2{font-weight:700}
.eyebrow{font-size:12px;letter-spacing:.06em;text-transform:uppercase;color:var(--muted);font-weight:600}
footer a{color:var(--fg)}
@media print{nav.jump,.tabs,.row button{display:none}.panel,.card{break-inside:avoid}}
</style>
</head>
<body>
<header class="masthead"><div class="wrap">
<a class="site-logo-link" href="https://abc-ri.org"><img class="site-logo" src="abcri-logo.png" width="420" height="160" alt="ABC-RI: Advocates for Better Care in Rhode Island"></a>
<h1 id="ttl"></h1>
<p>What CMS's public data shows about Rhode Island nursing homes since 2019, and how homes' ratings, staffing and inspections changed after they were sold.</p>
<span class="asof" id="asof"></span>
</div></header>
<nav class="jump" aria-label="Sections"><div class="wrap">
<a href="#glance">At a glance</a><a href="#trends">Statewide trends</a><a href="#sold">When homes were sold</a><a href="#lookup">Look up a home</a><a href="#groups">Owner groups</a><a href="#about">About the data</a>
</div></nav>
<main class="wrap">
<section id="glance"><h2>At a glance</h2><p class="lede" id="glance-lede"></p><div class="cards" id="cards"></div></section>

<section id="trends"><h2>Statewide trends</h2>
<p class="lede">Every Rhode Island nursing home in CMS's monthly Care Compare files. Charts start in April 2019, the first month after CMS changed how star ratings are calculated.</p>
<div class="panel"><div class="tabs" role="group" aria-label="Choose a measure" id="trend-tabs"></div>
<div class="chartbox" id="trend-chart"></div><div class="legend" id="trend-legend"></div><p class="small" id="trend-note"></p></div>
<div class="grid2"><div class="panel"><h3 style="margin-top:0">Serious citations per year</h3><div class="chartbox" id="ser-chart"></div><p class="small">Citations for actual harm or immediate jeopardy (scope/severity G-L), by inspection date. The latest year is partial.</p></div>
<div class="panel"><h3 style="margin-top:0">Federal fines per year</h3><div class="chartbox" id="fine-chart"></div><p class="small">Fines by penalty date, latest version of each fine. The latest year is partial.</p></div></div>
</section>

<section id="sold"><h2>What happened when homes were sold</h2>
<p class="lede">Rhode Island's Health Department reviews each change in a nursing home's effective control. This compares each home's ratings and staffing before and after such a change, using the last CMS rating before the sale as the starting point, and compares them with homes that were not sold over the same periods.</p>
<div class="grid2"><div class="panel"><h3 style="margin-top:0;font-size:1rem">Changes of control approved, by year</h3><div class="chartbox" id="peryear-chart" style="max-width:420px"></div><p class="small" id="peryear-note"></p></div>
<div class="panel" id="money-panel"><h3 style="margin-top:0;font-size:1rem">Rent paid to a landlord the owners also control</h3><div class="chartbox" id="money-chart"></div><div class="legend" id="money-legend"></div><p class="small" id="money-note"></p></div></div>
<div class="panel" id="ba-panel"><h3 style="margin-top:0">The same homes, before and after</h3>
<p class="lede" id="ba-lede"></p><div class="chartbox" id="ba-chart"></div><p class="small" id="ba-count"></p><p class="small" id="ba-note"></p></div>

<h3>One measure at a time</h3>
<div class="panel">
<div class="row"><div class="tabs" role="group" aria-label="Which changes" id="set-tabs" style="margin:0"></div><div class="tabs" role="group" aria-label="Choose a measure" id="metric-tabs" style="margin:0"></div></div>
<p class="small" id="study-howto" style="margin:4px 0 0"></p>
<div class="chartbox" id="study-chart"></div>
<div class="legend"><span><i class="sq" style="background:var(--sold)"></i>Homes that were sold</span><span><i class="sq" style="background:var(--ctrl)"></i>Comparable homes that were not sold</span></div>
<div class="callout" id="study-text" aria-live="polite"></div>
<details><summary>See the month-by-month timeline</summary><p class="small">Each point is the average change from the last CMS rating before the sale (that rating is 0, at "Sale"). Points left of "Sale" show the lead-up.</p><div class="chartbox" id="study-line"></div><div class="legend"><span><i style="border-color:var(--sold)"></i>Homes that were sold</span><span><i class="dash" style="border-color:var(--ctrl)"></i>Comparable homes that were not sold</span></div></details>
<p class="small" id="study-note"></p></div>

<h3>Every sale, one row each</h3>
<div class="row"><input type="search" id="sale-q" placeholder="Search by home, town or owner" aria-label="Search sales"><label class="chk"><input type="checkbox" id="sale-parent" checked> Include chain reorganizations</label><label class="chk"><input type="checkbox" id="sale-cms"> Only homes with CMS data</label><button class="btn" id="dl-sales">Download this table (CSV)</button></div>
<div class="tblwrap"><table id="sales-tbl"><caption>Dates are the CMS takeover date where CMS lists one, otherwise the state's decision date. Stars are CMS overall ratings (5 is best).</caption><thead></thead><tbody></tbody></table></div>
<p class="small">Blank cells mean CMS data is not available for that period: the sale was before CMS data starts in April 2019, is too recent for a year to have passed, or the home left the CMS files.</p>
</section>

<section id="lookup"><h2>Look up a home</h2>
<p class="lede">Search by current or former name, or town.</p>
<div class="row"><input type="search" id="fac-q" placeholder="Type a name or town" aria-label="Search homes" autocomplete="off"></div>
<ul class="results hidden" id="fac-results" role="listbox"></ul>
<div id="profile"></div>
</section>

<section id="groups"><h2>Owner groups</h2>
<p class="lede">Average CMS overall stars for the homes each group acquired: at the last rating before the sale, then one year later, two years later, and at the latest rating under that owner. Each comparison uses only homes that have reached that point, so recent purchases drop out of the one- and two-year columns; the count under each average says how many homes it covers. Click a group to see every home.</p>
<div class="tblwrap"><table id="groups-tbl"><thead><tr><th>Owner group</th><th class="num">Sales analyzed</th><th class="num">At acquisition → after 1 year</th><th class="num">At acquisition → after 2 years</th><th class="num">At acquisition → latest</th></tr></thead><tbody></tbody></table></div>
</section>

<section id="about"><h2>About the data</h2>
<div class="panel" id="about-body"></div>
<div class="row"><button class="btn" id="dl-monthly">Download monthly home-level data (CSV)</button></div>
</section>
<footer id="foot"></footer>
</main>
<script id="data" type="application/json">__DATA__</script>
<script>
"use strict";
const D = JSON.parse(document.getElementById('data').textContent);
const $ = (s, r) => (r || document).querySelector(s);
const el = (tag, attrs, ...kids) => { const n = document.createElement(tag); for (const k in (attrs || {})) { if (k === 'class') n.className = attrs[k]; else if (k === 'text') n.textContent = attrs[k]; else if (k.startsWith('on')) n.addEventListener(k.slice(2), attrs[k]); else if (attrs[k] !== null && attrs[k] !== undefined) n.setAttribute(k, attrs[k]); } for (const c of kids.flat()) { if (c === null || c === undefined || c === false) continue; n.append(c.nodeType ? c : document.createTextNode(String(c))); } return n; };
const NS = 'http://www.w3.org/2000/svg';
const sv = (tag, attrs, ...kids) => { const n = document.createElementNS(NS, tag); for (const k in (attrs || {})) n.setAttribute(k, attrs[k]); for (const c of kids.flat()) { if (c) n.append(c.nodeType ? c : document.createTextNode(String(c))); } return n; };
const MS = d => Date.parse(d + 'T00:00:00Z');
const fmtMon = d => new Date(MS(d)).toLocaleDateString('en-US', { month: 'short', year: 'numeric', timeZone: 'UTC' });
const fmtDay = d => new Date(MS(d)).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric', timeZone: 'UTC' });
const f1 = v => v == null ? '\u2013' : v.toFixed(1), f2 = v => v == null ? '\u2013' : v.toFixed(2), money = v => v == null ? '\u2013' : '$' + Math.round(v).toLocaleString('en-US');
const css = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const chip = v => v == null ? el('span', { class: 'chip na', text: '\u2013' }) : el('span', { class: 'chip s' + Math.round(v), text: String(Math.round(v)), 'aria-label': Math.round(v) + ' stars' });
const sign = (v, d) => (v > 0 ? '+' : v < 0 ? '\u2212' : '') + Math.abs(v).toFixed(d == null ? 2 : d);

/* ---------- charts ---------- */
function niceTicks(lo, hi, n) { const span = hi - lo || 1, step0 = span / n, mag = Math.pow(10, Math.floor(Math.log10(step0))), r = step0 / mag, step = (r < 1.5 ? 1 : r < 3 ? 2 : r < 7 ? 5 : 10) * mag; const a = Math.floor(lo / step) * step, out = []; for (let v = a; ; v += step) { out.push(+v.toFixed(10)); if (v >= hi - step * 0.001) break; } return out; }
function drawChart(box, o) {
  box.replaceChildren(); o.series = o.series || []; const W = Math.max(300, Math.min(760, box.clientWidth || 760)), H = o.height || 300, L = o.left || 46, R = 14, T = 14, B = 30, pw = W - L - R, ph = H - T - B;
  const xs = o.x, x0 = o.xmin != null ? o.xmin : Math.min(...xs), x1 = o.xmax != null ? o.xmax : Math.max(...xs);
  const all = o.series.flatMap(s => s.values).filter(v => v != null); let y0 = o.ymin != null ? o.ymin : Math.min(...all), y1 = o.ymax != null ? o.ymax : Math.max(...all);
  if (o.ymin == null && o.zero) y0 = Math.min(0, y0); if (y1 === y0) y1 = y0 + 1;
  const yt = niceTicks(y0, y1, 5); if (o.ymin == null) y0 = Math.min(y0, yt[0]); if (o.ymax == null) y1 = Math.max(y1, yt[yt.length - 1]);
  const X = v => L + (v - x0) / (x1 - x0 || 1) * pw, Y = v => T + ph - (v - y0) / (y1 - y0 || 1) * ph;
  const svg = sv('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': o.aria || 'Chart' });
  yt.forEach(t => { if (t < y0 - 1e-9 || t > y1 + 1e-9) return; svg.append(sv('line', { x1: L, x2: W - R, y1: Y(t), y2: Y(t), stroke: t === 0 && o.zero ? css('--line') : css('--grid'), 'stroke-width': t === 0 && o.zero ? 1.6 : 1 }), sv('text', { x: L - 6, y: Y(t) + 4, 'text-anchor': 'end', 'font-size': 11, fill: css('--muted') }, o.yfmt ? o.yfmt(t) : String(t))); });
  const xt = o.xticks || (() => { const a = new Date(x0), out = []; for (let y = a.getUTCFullYear() + 1; ; y++) { const t = Date.UTC(y, 0, 1); if (t > x1) break; out.push({ v: t, l: String(y) }); } return out; })();
  xt.forEach(t => { svg.append(sv('line', { x1: X(t.v), x2: X(t.v), y1: T, y2: T + ph, stroke: css('--grid'), 'stroke-dasharray': '2 4' }), sv('text', { x: X(t.v), y: H - 10, 'text-anchor': 'middle', 'font-size': 11, fill: css('--muted') }, t.l)); });
  (o.bands || []).forEach(b => { const xa = X(b.x0), xb = X(b.x1); svg.append(sv('rect', { x: xa, y: T, width: Math.max(2, xb - xa), height: ph, fill: css('--line'), opacity: .55 }, sv('title', {}, b.label)), (xb - xa) > 46 ? sv('text', { x: (xa + xb) / 2, y: T + ph / 2, 'text-anchor': 'middle', 'font-size': 11, fill: css('--muted') }, b.short || 'Not rated') : null); });
  if (o.stack) {
    // stacked area: one band per category, far more legible than one thin bar per month
    const nb = o.stack[0].length;
    const cum = o.stack.map(col => { const tot = col.reduce((a, b) => a + b, 0) || 1; let acc = 0; return col.map(c => { acc += c; return acc / tot * 100; }); });
    const MAXGAP = 70 * 864e5, runs = []; let cur = [0];
    for (let i = 1; i < xs.length; i++) { if (xs[i] - xs[i - 1] > MAXGAP) { runs.push(cur); cur = [i]; } else cur.push(i); }
    runs.push(cur);
    const edge = (run, k) => run.map((i, j) => (j ? 'L' : 'M') + X(xs[i]).toFixed(1) + ',' + Y(cum[i][k]).toFixed(1)).join('');
    for (let k = nb - 1; k >= 0; k--) runs.forEach(run => {                      // back to front: each fill covers the one below
      if (run.length < 2) return;
      const d = edge(run, k) + 'L' + X(xs[run[run.length - 1]]).toFixed(1) + ',' + Y(0).toFixed(1) + 'L' + X(xs[run[0]]).toFixed(1) + ',' + Y(0).toFixed(1) + 'Z';
      svg.append(sv('path', { d, fill: o.stackColors ? o.stackColors[k] : css('--s' + (k + 1)) }));
    });
    for (let k = 0; k < nb - 1; k++) runs.forEach(run => {                       // thin surface rule between bands
      if (run.length < 2) return;
      svg.append(sv('path', { d: edge(run, k), fill: 'none', stroke: css('--chart-bg'), 'stroke-width': 1 }));
    });
  } else {
    [...o.series].sort((a, b) => (a.z || 0) - (b.z || 0)).forEach(s => { let d = '', pen = false; s.values.forEach((v, i) => { if (v == null) { pen = false; return; } d += (pen ? 'L' : 'M') + X(xs[i]).toFixed(1) + ' ' + Y(v).toFixed(1); pen = true; });
      svg.append(sv('path', { d, fill: 'none', stroke: s.color, 'stroke-width': s.width || 2.2, 'stroke-dasharray': s.dash || '', 'stroke-linejoin': 'round' }));
      if (s.dots) s.values.forEach((v, i) => { if (v != null) svg.append(sv('circle', { cx: X(xs[i]), cy: Y(v), r: s.dots === true ? 3 : s.dots, fill: s.color })); }); });
  }
  (o.markers || []).forEach(m => { if (m.x < x0 || m.x > x1) return; svg.append(sv('line', { x1: X(m.x), x2: X(m.x), y1: T, y2: T + ph, stroke: m.color, 'stroke-width': m.w || 1.6, 'stroke-dasharray': m.dash || '' }), sv('circle', { cx: X(m.x), cy: T + 4, r: 4, fill: m.color }, sv('title', {}, m.title || ''))); });
  const cross = sv('line', { y1: T, y2: T + ph, stroke: css('--fg'), 'stroke-width': 1, opacity: 0, 'pointer-events': 'none' }); svg.append(cross);
  const tip = el('div', { class: 'tip' }); box.append(svg, tip);
  const move = ev => { const r = svg.getBoundingClientRect(), cx = ((ev.touches ? ev.touches[0].clientX : ev.clientX) - r.left) / r.width * W; let bi = 0, bd = 1e18; xs.forEach((v, i) => { const d = Math.abs(X(v) - cx); if (d < bd) { bd = d; bi = i; } });
    cross.setAttribute('x1', X(xs[bi])); cross.setAttribute('x2', X(xs[bi])); cross.setAttribute('opacity', .6);
    const lines = [el('b', { text: o.xlabel ? o.xlabel(xs[bi], bi) : String(xs[bi]) })];
    if (o.stack) { const tot = o.stack[bi].reduce((a, b) => a + b, 0); o.stack[bi].forEach((c, k) => lines.push(el('div', {}, `${o.stackLabels ? o.stackLabels[k] : (k + 1) + ' star' + (k ? 's' : '')}: ${c} (${Math.round(c / (tot || 1) * 100)}%)`))); }
    else o.series.forEach(s => { const v = s.values[bi]; lines.push(el('div', {}, el('span', { style: `color:${s.color}`, text: '\u25CF ' }), `${s.name}: ${v == null ? 'no data' : (o.tfmt ? o.tfmt(v) : (o.yfmt ? o.yfmt(v) : v))}`)); });
    if (o.extra) lines.push(...o.extra(bi));
    tip.replaceChildren(...lines); tip.style.display = 'block'; const left = X(xs[bi]) / W * r.width; tip.style.left = Math.min(Math.max(4, left + 12), r.width - tip.offsetWidth - 4) + 'px'; tip.style.top = '6px'; };
  const leave = () => { tip.style.display = 'none'; cross.setAttribute('opacity', 0); };
  svg.addEventListener('mousemove', move); svg.addEventListener('touchstart', move, { passive: true }); svg.addEventListener('touchmove', move, { passive: true }); svg.addEventListener('mouseleave', leave); svg.addEventListener('touchend', leave);
}
function dumbbells(box, rows, aria) {
  box.replaceChildren();
  const W = Math.max(320, Math.min(760, box.clientWidth || 760)), narrow = W < 520;
  const rowH = narrow ? 86 : 64, padT = 40, padB = 12, L = narrow ? 8 : 250, R = narrow ? 8 : 86;
  const x0 = L + (narrow ? 86 : 40), x1 = W - R - (narrow ? 86 : 70), H = padT + rows.length * rowH + padB;
  const svg = sv('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': aria });
  svg.append(sv('text', { x: x0, y: 18, 'text-anchor': 'middle', 'font-size': 12, 'font-weight': 700, fill: css('--muted') }, 'Year before'),
             sv('text', { x: x1, y: 18, 'text-anchor': 'middle', 'font-size': 12, 'font-weight': 700, fill: css('--muted') }, 'Two years after'));
  rows.forEach((r, i) => {
    const top = padT + i * rowH, y = top + (narrow ? 52 : 26);
    const diff = r.after - r.before, eps = r.eps == null ? 0 : r.eps;
    const dir = Math.abs(diff) <= eps ? 'flat' : (r.neutral ? 'neutral' : ((r.higherBad ? diff > 0 : diff < 0) ? 'bad' : 'good'));
    const col = dir === 'bad' ? css('--sold') : dir === 'good' ? css('--s4') : css('--ctrl');
    svg.append(sv('text', { x: narrow ? 0 : 0, y: narrow ? top + 14 : y - 4, 'font-size': 13, 'font-weight': 600, fill: css('--fg') }, r.label),
               sv('text', { x: 0, y: narrow ? top + 30 : y + 13, 'font-size': 11, fill: css('--muted') }, r.sub));
    svg.append(sv('line', { x1: x0, x2: x1, y1: y, y2: y, stroke: col, 'stroke-width': 5, opacity: .3, 'stroke-linecap': 'round' }));
    const mid = (x0 + x1) / 2;
    svg.append(sv('path', { d: `M${mid - 7} ${y - 6}L${mid + 7} ${y}L${mid - 7} ${y + 6}Z`, fill: col }));
    svg.append(sv('circle', { cx: x0, cy: y, r: 7, fill: css('--ctrl') }, sv('title', {}, 'Year before: ' + r.fmt(r.before))),
               sv('circle', { cx: x1, cy: y, r: 8, fill: col }, sv('title', {}, 'Two years after: ' + r.fmt(r.after))));
    svg.append(sv('text', { x: x0 - 12, y: y + 5, 'text-anchor': 'end', 'font-size': 13.5, fill: css('--muted') }, r.fmt(r.before)),
               sv('text', { x: x1 + 12, y: y + 5, 'font-size': 14, 'font-weight': 700, fill: col }, r.fmt(r.after)));
    if (!narrow) svg.append(sv('text', { x: W, y: y + 5, 'text-anchor': 'end', 'font-size': 11.5, fill: css('--muted') }, r.n + ' homes'));
    else svg.append(sv('text', { x: W, y: top + 14, 'text-anchor': 'end', 'font-size': 11.5, fill: css('--muted') }, r.n + ' homes'));
  });
  box.append(svg);
}
function groupedBars(box, groups, fmt, aria, zeroNote) {
  box.replaceChildren(); const W = Math.max(300, Math.min(760, box.clientWidth || 760)), H = 300, L = 54, R = 12, T = 26, B = 58, pw = W - L - R, ph = H - T - B;
  const vals = groups.flatMap(g => [g.sold, g.ctrl]).filter(v => v != null); let lo = Math.min(0, ...vals), hi = Math.max(0, ...vals); const pad = (hi - lo) * .15 || .1; lo -= pad; hi += pad;
  const Y = v => T + ph - (v - lo) / (hi - lo) * ph, svg = sv('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': aria }); const yt = niceTicks(lo, hi, 5);
  yt.forEach(t => { if (t < lo - 1e-9 || t > hi + 1e-9) return; svg.append(sv('line', { x1: L, x2: W - R, y1: Y(t), y2: Y(t), stroke: css('--grid'), 'stroke-width': Math.abs(t) < 1e-9 ? 0 : 1 }), sv('text', { x: L - 6, y: Y(t) + 4, 'text-anchor': 'end', 'font-size': 11, fill: css('--muted') }, fmt(t))); });
  svg.append(sv('line', { x1: L, x2: W - R, y1: Y(0), y2: Y(0), stroke: css('--fg'), 'stroke-width': 1.5 }));
  const gw = pw / groups.length, bw = Math.min(70, gw * .32);
  groups.forEach((g, i) => { const cx = L + gw * i + gw / 2;
    [[g.sold, css('--sold'), -bw - 3, 'Sold homes'], [g.ctrl, css('--ctrl'), 3, 'Comparable homes not sold']].forEach(b => { if (b[0] == null) return; const y0 = Y(0), y1 = Y(b[0]), top = Math.min(y0, y1), hgt = Math.max(1.5, Math.abs(y1 - y0));
      svg.append(sv('rect', { x: cx + b[2], y: top, width: bw, height: hgt, fill: b[1], rx: 2 }, sv('title', {}, b[3] + ': ' + fmt(b[0]))), sv('text', { x: cx + b[2] + bw / 2, y: b[0] >= 0 ? top - 5 : top + hgt + 13, 'text-anchor': 'middle', 'font-size': 12, 'font-weight': 700, fill: css('--fg') }, fmt(b[0]))); });
    svg.append(sv('text', { x: cx, y: H - 30, 'text-anchor': 'middle', 'font-size': 12.5, 'font-weight': 600, fill: css('--fg') }, g.label), sv('text', { x: cx, y: H - 14, 'text-anchor': 'middle', 'font-size': 11, fill: css('--muted') }, g.n == null ? 'not enough homes' : g.n + ' homes')); });
  box.append(svg);
}
function barChart(box, labels, values, color, fmt, aria, integer, opts) {
  opts = opts || {};
  box.replaceChildren(); const W = Math.max(260, Math.min(opts.maxW || 520, box.clientWidth || 520)), H = opts.height || 240, L = 54, R = 10, T = 12, B = 28, pw = W - L - R, ph = H - T - B, mx = Math.max(...values, 1); let yt = niceTicks(0, mx, 4); if (integer) yt = yt.filter(t => Number.isInteger(t)); if (integer && yt[yt.length - 1] < mx) yt.push(yt[yt.length - 1] + 1); const top = yt[yt.length - 1];
  const svg = sv('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': aria }); const Y = v => T + ph - v / top * ph, bw = pw / labels.length;
  yt.forEach(t => svg.append(sv('line', { x1: L, x2: W - R, y1: Y(t), y2: Y(t), stroke: css('--grid') }), sv('text', { x: L - 6, y: Y(t) + 4, 'text-anchor': 'end', 'font-size': 11, fill: css('--muted') }, fmt(t))));
  values.forEach((v, i) => { svg.append(sv('rect', { x: L + i * bw + bw * .15, y: Y(v), width: bw * .7, height: T + ph - Y(v), fill: color }, sv('title', {}, labels[i] + ': ' + fmt(v))), sv('text', { x: L + i * bw + bw / 2, y: H - 10, 'text-anchor': 'middle', 'font-size': 11, fill: css('--muted') }, labels[i])); });
  box.append(svg);
}

/* ---------- header + cards ---------- */
$('#ttl').textContent = D.meta.title; document.title = D.meta.title; $('#asof').textContent = 'CMS data through ' + D.meta.updated;
{ const bh = D.meta.behind || {}, ks = Object.keys(bh);
  if (ks.length) { const p = el('p', { class: 'small', style: 'margin:10px 0 0;opacity:.9;color:inherit' }, 'Not every CMS file is updated in the same month. Latest here: ' + ks.map(k => k.toLowerCase() + ' ' + bh[k]).join(', ') + '.'); $('#asof').parentNode.append(p); } }
const H0 = D.head;
$('#glance-lede').textContent = `Comparing CMS's ${fmtMon(D.dates[0])} and ${fmtMon(D.dates[D.dates.length - 1])} files.`;
[[H0.low_now + '%', 'of homes rated 0, 1 or 2 stars overall (0 = special focus facility, not rated)', `${H0.low_then}% in ${fmtMon(D.dates[0])}`],
 [f2(H0.avg_now), 'average overall star rating (5 is best; special focus facilities count as 0)', `${f2(H0.avg_then)} in ${fmtMon(D.dates[0])}`],
 [f2(H0.hours_now), 'nurse staffing hours per resident per day', `${f2(H0.hours_then)} in ${fmtMon(D.dates[0])}; RN hours ${f2(H0.rn_then)} \u2192 ${f2(H0.rn_now)}`],
 [H0.turnover_now == null ? '\u2013' : Math.round(H0.turnover_now) + '%', 'average nursing staff turnover', H0.turnover_first ? Math.round(H0.turnover_first.value) + '% when CMS began publishing it in ' + fmtMon(H0.turnover_first.date) : 'CMS began publishing this in 2022'],
 [String(H0.n_now), 'homes in the latest CMS file', `${H0.n_then} in ${fmtMon(D.dates[0])}; ${H0.homes_left} left the files`],
 [String(H0.sales_since), 'changes of control approved since April 2019', 'See the count for each year below']].forEach(c => $('#cards').append(el('div', { class: 'card' }, el('div', { class: 'big', text: c[0] }), el('div', { class: 'lbl', text: c[1] }), el('div', { class: 'cmp', text: c[2] }))));

/* ---------- statewide trends ---------- */
const X = D.dates.map(MS), st = D.state; let trendKey = 'mix';
const TRENDS = { mix: { label: 'Star ratings mix' }, avg: { label: 'Average stars' }, hours: { label: 'Nurse staffing hours' }, turn: { label: 'Staff turnover' } };
function drawTrend() {
  const box = $('#trend-chart'), lg = $('#trend-legend'), note = $('#trend-note'); lg.replaceChildren(); const xl = (v, i) => fmtMon(D.dates[i]);
  if (trendKey === 'mix') { drawChart(box, { x: X, stack: st.dist.map((d, i) => [st.nr[i]].concat(d)), stackColors: [0, 1, 2, 3, 4, 5].map(k => css('--s' + k)), stackLabels: ['0 stars (not rated: special focus facility)', '1 star', '2 stars', '3 stars', '4 stars', '5 stars'], ymin: 0, ymax: 100, yfmt: v => v + '%', xlabel: xl, aria: 'Share of homes at each star rating over time' });
    [0, 1, 2, 3, 4, 5].forEach(k => lg.append(el('span', {}, el('i', { style: `border-color:var(--s${k});border-top-width:9px` }), k === 0 ? '0 = not rated (special focus facility)' : k + (k === 1 ? ' star' : ' stars')))); note.textContent = 'One point per monthly CMS file; the band breaks where the archive has no file for a month (CMS did not publish one, or it is missing). Share of Rhode Island homes at each overall star rating in each file. Gray at the bottom is homes CMS did not rate because they are in its Special Focus Facility program, counted as 0 stars. Rating calculations have been revised by CMS over time.'; }
  if (trendKey === 'avg') { drawChart(box, { x: X, ymin: 1, ymax: 5, yfmt: v => v.toFixed(1), xlabel: xl, tfmt: v => v.toFixed(2), aria: 'Average star ratings over time', series: [
      { name: 'Overall', color: css('--accent'), values: st.avg_overall, width: 3 }, { name: 'Health inspection', color: css('--s2'), values: st.avg_hi }, { name: 'Staffing', color: css('--s4'), values: st.avg_staffing }, { name: 'Quality measures', color: css('--ctrl'), values: st.avg_qm, dash: '5 4' }] });
    [['Overall', '--accent'], ['Health inspection', '--s2'], ['Staffing', '--s4'], ['Quality measures (dashed)', '--ctrl']].forEach(a => lg.append(el('span', {}, el('i', { class: a[0].includes('dashed') ? 'dash' : '', style: `border-color:var(${a[1]})` }), a[0]))); note.textContent = 'Average of the four CMS star ratings across homes in each file (5 is best). Special focus facilities, which CMS leaves unrated, count as 0.'; }
  if (trendKey === 'hours') { drawChart(box, { x: X, ymin: 0, yfmt: v => v.toFixed(1), tfmt: v => v.toFixed(2), xlabel: xl, aria: 'Average nurse staffing hours per resident per day', series: [{ name: 'All nurses (RN + LPN + aides)', color: css('--accent'), values: st.avg_hours, width: 3 }, { name: 'Registered nurses (RN)', color: css('--s2'), values: st.avg_rn }] });
    lg.append(el('span', {}, el('i', { style: 'border-color:var(--accent)' }), 'All nurses (RN + LPN + aides)'), el('span', {}, el('i', { style: 'border-color:var(--s2)' }), 'Registered nurses (RN)')); note.textContent = 'Average hours of nursing care per resident per day reported to CMS through payroll records.'; }
  if (trendKey === 'turn') { drawChart(box, { x: X, ymin: 0, yfmt: v => Math.round(v) + '%', tfmt: v => Math.round(v) + '%', xlabel: xl, aria: 'Average nursing staff turnover', series: [{ name: 'Nursing staff turnover', color: css('--sold'), values: st.avg_turnover, width: 3 }] });
    lg.append(el('span', {}, el('i', { style: 'border-color:var(--sold)' }), 'Average nursing staff turnover')); note.textContent = 'CMS began publishing turnover in 2022, so earlier months are blank.'; }
  [...$('#trend-tabs').children].forEach(b => b.setAttribute('aria-pressed', b.dataset.k === trendKey));
}
Object.keys(TRENDS).forEach(k => $('#trend-tabs').append(el('button', { class: 'tab', 'data-k': k, 'aria-pressed': 'false', text: TRENDS[k].label, onclick: () => { trendKey = k; drawTrend(); } })));
const yl = D.state_years.years.map((y, i) => i === D.state_years.years.length - 1 && D.state_years.last_year_quarters < 4 ? y + '*' : String(y));
barChart($('#ser-chart'), yl, D.state_years.serious, css('--s1'), v => String(Math.round(v)), 'Serious citations per year', true); barChart($('#fine-chart'), yl, D.state_years.fines, css('--sold'), v => v >= 1e6 ? '$' + (v / 1e6).toFixed(1) + 'M' : '$' + Math.round(v / 1e3) + 'K', 'Federal fines per year');
drawTrend();

/* ---------- when homes were sold ---------- */
const METRICS = { overall: { label: 'Overall stars', unit: 'stars', scale: 1, d: 2, hb: true }, staffing: { label: 'Staffing stars', unit: 'staffing stars', scale: 1, d: 2, hb: true }, hours: { label: 'Nurse hours per resident per day', unit: 'nurse hours per resident per day', scale: 1, d: 2, hb: true }, turnover: { label: 'Staff turnover', unit: 'percentage points of nursing staff turnover', scale: 100, d: 0, hb: false } };
let setKey = 'sold', metKey = 'overall';
function drawStudy() {
  const st_ = D.study[setKey], m = METRICS[metKey], pts = st_.metrics[metKey].points, sm = st_.metrics[metKey].summary, at = mo => pts.find(p => p.months === mo) || {};
  const sgn = (v, d) => (v > 0 ? '+' : v < 0 ? '\u2212' : '') + Math.abs(v).toFixed(d), dd = m.d, fmt = v => sgn(v * m.scale, m.scale === 100 ? 0 : dd);
  const pre = at(-12), y1 = at(12), y2 = at(24);
  const groups = [{ label: 'Year before the sale', n: pre.n || null, sold: pre.sold == null ? null : -pre.sold, ctrl: pre.ctrl == null ? null : -pre.ctrl }, { label: 'First year after', n: y1.n || null, sold: y1.sold, ctrl: y1.ctrl }, { label: 'First two years after', n: y2.n || null, sold: y2.sold, ctrl: y2.ctrl }]
    .map(g => ({ ...g, sold: g.sold == null ? null : g.sold * m.scale, ctrl: g.ctrl == null ? null : g.ctrl * m.scale }));
  $('#study-howto').textContent = `Each bar is the average change in ${m.unit} over that period. ` + (m.hb ? 'Below zero means it fell.' : 'Above zero means it rose, which is worse.') + ' "Year before" is the change leading up to the sale; the two "after" bars are the change since the sale.';
  groupedBars($('#study-chart'), groups, v => (v > 0 ? '+' : v < 0 ? '\u2212' : '') + Math.abs(v).toFixed(m.scale === 100 ? 0 : dd), 'Average change in ' + m.label + ' before and after a sale, sold homes compared with homes not sold');
  const lp = pts.filter(p => p.sold != null);
  drawChart($('#study-line'), { x: lp.map(p => p.months), xmin: -12, xmax: 24, series: [{ name: 'Sold homes', color: css('--sold'), values: lp.map(p => p.sold * m.scale), width: 3 }, { name: 'Comparable homes not sold', color: css('--ctrl'), values: lp.map(p => p.ctrl * m.scale), dash: '6 4' }],
    ymin: null, zero: true, yfmt: v => (v > 0 ? '+' : v < 0 ? '\u2212' : '') + Math.abs(v).toFixed(m.scale === 100 ? 0 : dd), tfmt: v => sign(v, m.d), xlabel: v => v === 0 ? 'Last rating before the sale' : (v < 0 ? Math.abs(v) + ' months before' : v + ' months after'), height: 300,
    xticks: [-12, -6, 0, 6, 12, 18, 24].map(v => ({ v, l: v === 0 ? 'Sale' : (v > 0 ? '+' + v + ' mo' : v + ' mo') })), extra: bi => [el('div', { class: 'small', text: 'Homes in this comparison: ' + lp[bi].n })], aria: 'Month-by-month average change in ' + m.label });
  const t = $('#study-text'); t.replaceChildren(); const one = sm['12'], two = sm['24'], fmtd = v => sign(v * m.scale, m.d);
  if (one) t.append(el('p', {}, el('b', { text: 'One year after the sale: ' }), `${one.n} homes changed by an average of ${fmtd(one.sold)} ${m.unit}, compared with ${fmtd(one.ctrl)} for comparable homes that were not sold over the same periods. ${one.worse} of ${one.n} moved in the wrong direction, ${one.same} were unchanged, ${one.better} moved in the right direction.`));
  if (two) t.append(el('p', {}, el('b', { text: 'Two years after: ' }), `${two.n} homes changed by ${fmtd(two.sold)} ${m.unit} on average, compared with ${fmtd(two.ctrl)} for comparable homes.`));
  if (pre.sold != null && Math.abs(pre.sold * m.scale) >= (m.scale === 100 ? 2 : 0.15) && (m.hb ? pre.sold > 0 : pre.sold < 0)) t.append(el('p', {}, el('b', { text: 'Many were already slipping before the sale: ' }), `in the year leading up to the sale these homes lost ${Math.abs(pre.sold * m.scale).toFixed(m.scale === 100 ? 0 : dd)} ${m.unit} on average.`));
  if (!one) t.append(el('p', { text: 'Not enough homes with data for this measure yet.' }));
  $('#study-note').textContent = `Based on ${st_.n_events} sales (${st_.n_homes} homes) with a CMS rating before the sale, in files from April 2019. Months a home was in CMS's Special Focus Facility program count as 0 stars. "Comparable homes" are homes with no change of control or new operator from a year before to two years after the sale. A buyer sometimes runs a home under a management agreement before the state approves the sale, and CMS can take months to show a new operator, so the "before" rating may already reflect the buyer's management. This shows what happened, not why: ratings also move with inspections, staffing rules and the pandemic, which is why the comparison bars matter.`;
  [...$('#set-tabs').children].forEach(b => b.setAttribute('aria-pressed', b.dataset.k === setKey)); [...$('#metric-tabs').children].forEach(b => b.setAttribute('aria-pressed', b.dataset.k === metKey));
}
[['sold', 'Homes sold to a new operator'], ['all', 'Including chain reorganizations']].forEach(a => $('#set-tabs').append(el('button', { class: 'tab', 'data-k': a[0], 'aria-pressed': 'false', text: a[1], onclick: () => { setKey = a[0]; drawStudy(); } })));
Object.keys(METRICS).forEach(k => $('#metric-tabs').append(el('button', { class: 'tab', 'data-k': k, 'aria-pressed': 'false', text: METRICS[k].label, onclick: () => { metKey = k; drawStudy(); } })));
drawStudy();
(function () { const py = D.head.per_year || []; if (!py.length) return;
  const lastY = py[py.length - 1].year, labs = py.map(p => p.year === lastY ? p.year + '*' : String(p.year));
  barChart($('#peryear-chart'), labs, py.map(p => p.n), css('--accent'), v => String(Math.round(v)), 'Changes of control approved by year', true, { maxW: 420, height: 170 });
  $('#peryear-note').textContent = 'Counted by the date used on this page (the CMS takeover date where CMS lists one, otherwise the state decision date), so a deal approved late in one year can fall in the next. ' + lastY + ' is partial, and ' + py[0].year + ' starts in April.'; })();

/* where the money goes: rent to a landlord under the same ownership, sold homes vs the rest */
function drawMoney() {
  const t = D.money && D.money.trend, panel = $('#money-panel');
  if (!t || !panel) { if (panel) panel.classList.add('hidden'); return; }
  const n = t.years.length - 1, pc = v => v.toFixed(1) + '%';
  const span = t.years[0] + '\u2013' + t.years[n];
  const sold = { name: 'Changed hands', color: css('--sold'), values: t.rent.sold, width: 3, dots: true, z: 2 };
  const other = { name: 'Did not', color: css('--teal'), values: t.rent.other, width: 2.4, dash: '7 4', dots: true, z: 1 };
  drawChart($('#money-chart'), {
    x: t.years, xmin: t.years[0] - .25, xmax: t.years[n] + .25, height: 210, left: 40, ymin: 0, zero: true,
    xticks: t.years.map(y => ({ v: y, l: "'" + String(y).slice(2) })),
    yfmt: v => v + '%', tfmt: pc, xlabel: v => 'Fiscal year ' + v, series: [sold, other],
    aria: 'Rent paid to a related landlord as a share of patient revenue, from ' + t.years[0] + ' to ' + t.years[n] + ', for homes that changed hands compared with homes that did not'
  });
  $('#money-legend').replaceChildren(
    el('span', {}, el('i', { style: 'border-color:' + css('--sold') }), 'Changed hands in these years (' + t.n_sold + ' homes)'),
    el('span', {}, el('i', { class: 'dash', style: 'border-color:' + css('--teal') }), 'Did not (' + t.n_other + ')'));
  const s0 = t.rent.sold[0], s1 = t.rent.sold[n], o0 = t.rent.other[0], o1 = t.rent.other[n];
  const mil = v => '$' + (v / 1e6).toFixed(1) + ' million';
  const usd = t.rent_usd && t.rent_usd.sold, dollars = usd ? ' In dollars, rent at those ' + t.n_sold + ' homes went from ' + mil(usd[0]) + ' a year to ' + mil(usd[n]) + '.' : '';
  $('#money-note').replaceChildren(
    el('b', { text: 'At the ' + t.n_sold + ' homes that changed hands between ' + span + ', rent paid to a landlord under the same ownership went from ' + pc(s0) + ' of patient revenue to ' + pc(s1) + '. At the ' + t.n_other + ' that did not change hands, it stayed near ' + pc(o1) + '. ' }),
    'These are the same ' + (t.n_sold + t.n_other) + ' homes in every year, so this is not a change in which homes report. Each figure is rent and building costs the home itself reported paying to a property company its own owners control, on Worksheet A-8-1 of its Medicare cost report; rent to an unrelated landlord is not counted.' + dollars +
    ' Cost reports run about two years behind the CMS ratings above.');
}
drawMoney();

/* before/after, same homes */
(function () {
  const lv = D.levels, mn = D.money, rows = [], pct = v => Math.round(v * 100) + '%', st1 = v => v.toFixed(1), hr = v => v.toFixed(2);
  const add = (o) => { if (o) rows.push(o); };
  if (lv.overall && lv.overall['2']) add({ label: 'Overall star rating', sub: 'CMS rating, 5 is best', before: lv.overall['2'].before, after: lv.overall['2'].after, n: lv.overall['2'].n, fmt: st1, higherBad: false, eps: 0.1 });
  if (lv.staffing && lv.staffing['2']) add({ label: 'Staffing star rating', sub: 'CMS rating, 5 is best', before: lv.staffing['2'].before, after: lv.staffing['2'].after, n: lv.staffing['2'].n, fmt: st1, higherBad: false, eps: 0.1 });
  if (lv.hours && lv.hours['2']) add({ label: 'Nurse hours per resident per day', sub: 'reported hours, all nursing staff', before: lv.hours['2'].before, after: lv.hours['2'].after, n: lv.hours['2'].n, fmt: hr, higherBad: false, eps: 0.1 });
  if (lv.rn && lv.rn['2']) add({ label: 'RN hours per resident per day', sub: 'registered nurses only', before: lv.rn['2'].before, after: lv.rn['2'].after, n: lv.rn['2'].n, fmt: hr, higherBad: false, eps: 0.05 });
  if (mn && mn.compare) {
    const c = mn.compare;
    if (c.rp_sh && c.rp_sh['2']) add({ label: 'Money paid to related parties', sub: 'share of the home\u2019s patient revenue', before: c.rp_sh['2'].before, after: c.rp_sh['2'].after, n: c.rp_sh['2'].n, fmt: pct, higherBad: true, eps: 0.01 });
    if (c.rent_sh && c.rent_sh['2']) add({ label: 'Rent paid to a related party', sub: 'share of patient revenue', before: c.rent_sh['2'].before, after: c.rent_sh['2'].after, n: c.rent_sh['2'].n, fmt: pct, higherBad: true, eps: 0.01 });
  }
  if (!rows.length) { $('#ba-panel').classList.add('hidden'); return; }
  $('#ba-lede').textContent = 'Each row follows the same homes from the year before the sale (left) to two years after (right). Orange means the measure moved the wrong way for residents; gray means it barely moved.';
  $('#ba-count').textContent = 'A row counts only homes with a usable figure at both ends, so the counts differ: a home sold too recently has no "two years after" yet, CMS leaves staffing blank when payroll data is missing, and the money rows come from annual Medicare cost reports, which cover ' + ((D.money && D.money.n_facilities) || 0) + ' of these homes and run through fiscal year ' + ((D.money && D.money.last_year) || '') + '.';
  dumbbells($('#ba-chart'), rows, 'Care and money measures for the same homes, year before the sale compared with two years after');
  const c = (D.money && D.money.compare) || {}, bits = [];
  if (c.rp_sh && c.rp_sh['2']) bits.push(`Related-party payments rose for ${c.rp_sh['2'].up} of the ${c.rp_sh['2'].n} homes with cost reports for both years.`);
  if (c.rent_sh && c.rent_sh['2']) bits.push(`Rent paid to a related party rose for ${c.rent_sh['2'].up} of ${c.rent_sh['2'].n}.`);
  const ns = (D.money && D.money.never_sold && D.money.never_sold.rp_sh) || {}, yrs = Object.keys(ns).sort();
  if (yrs.length) bits.push(`Homes that were not sold paid about ${Math.round(ns[yrs[yrs.length - 1]] * 100)}% of revenue to related parties in ${yrs[yrs.length - 1]}, little changed since ${yrs[0]}.`);
  $('#ba-note').textContent = bits.join(' ') + ' Nurse hours barely moved, so the extra money did not go into more nursing time at the bedside. Related-party figures come from the homes\u2019 own Medicare cost reports: payments to companies the owners also control, such as a landlord or a management company. Money moved this way leaves the home\u2019s books as a cost, so a home can look unprofitable while its owners are paid.';
})();

/* sales table */
let sortKey = 'date', sortDir = -1;
const cols = [['date', 'Date'], ['name', 'Home'], ['label', 'Type'], ['buyer', 'New owner'], ['overall', 'Overall stars: before \u2192 1 yr \u2192 2 yrs'], ['hours', 'Nurse hours: before \u2192 1 yr'], ['ser', 'Serious citations: year before \u2192 year after'], ['fines', 'Fines: year before \u2192 year after']];
$('#sales-tbl thead').append(el('tr', {}, cols.map(c => el('th', { scope: 'col', class: ['overall', 'hours', 'ser', 'fines'].includes(c[0]) ? 'num' : '' }, ['date', 'name', 'label', 'buyer'].includes(c[0]) ? el('button', { text: c[1], onclick: () => { sortDir = sortKey === c[0] ? -sortDir : 1; sortKey = c[0]; drawSales(); }, 'aria-label': 'Sort by ' + c[1] }) : c[1]))));
function trio(a, b, c, fmt, cell) { const parts = [a, b, c].filter((v, i) => i === 0 || v != null || (i === 1)); return parts; }
function drawSales() {
  const q = $('#sale-q').value.trim().toLowerCase(), incP = $('#sale-parent').checked, onlyC = $('#sale-cms').checked;
  let rows = D.sales.filter(r => (incP || r.kind === 'sale') && (!onlyC || r.has_cms) && (!q || [r.name, r.town, r.buyer, r.who].join(' ').toLowerCase().includes(q)));
  rows.sort((a, b) => { const A = a[sortKey] == null ? '' : a[sortKey], B = b[sortKey] == null ? '' : b[sortKey]; return (A < B ? -1 : A > B ? 1 : 0) * sortDir; });
  const tb = $('#sales-tbl tbody'); tb.replaceChildren();
  rows.forEach(r => { const m = r.m;
    const newOwner = el('td', {}, r.buyer ? el('b', { text: r.buyer }) : null, r.buyer ? el('br') : null, el('span', { class: 'small', text: r.who }));
    const stars = el('td', { class: 'num' }, r.has_cms && m.overall.before != null ? [chip(m.overall.before), el('span', { class: 'arrow', text: '\u2192' }), chip(m.overall.y1), el('span', { class: 'arrow', text: '\u2192' }), chip(m.overall.y2)] : el('span', { class: 'small', text: r.has_cms ? 'no rating before sale' : 'no CMS data' }));
    const hrs = el('td', { class: 'num' }, r.has_cms && m.hours.before != null ? [f2(m.hours.before), el('span', { class: 'arrow', text: '\u2192' }), f2(m.hours.y1)] : '\u2013');
    const ser = el('td', { class: 'num' }, r.ser.before != null ? [String(r.ser.before), el('span', { class: 'arrow', text: '\u2192' }), r.ser.y1 == null ? '\u2013' : String(r.ser.y1)] : '\u2013');
    const fin = el('td', { class: 'num' }, r.fines.before != null ? [money(r.fines.before), el('span', { class: 'arrow', text: '\u2192' }), r.fines.y1 == null ? '\u2013' : money(r.fines.y1)] : '\u2013');
    tb.append(el('tr', {}, el('td', { class: 'num', title: r.date_note }, fmtDay(r.date), r.mgmt_start ? el('div', { class: 'small', text: 'Buyer managing from ' + fmtMon(r.mgmt_start) }) : null, r.cms_op ? el('div', { class: 'small', text: 'CMS records updated ' + fmtMon(r.cms_op) }) : null), el('td', {}, D.facilities.some(f => f.id === r.fid) ? el('button', { class: 'linkbtn', text: r.name, onclick: () => showFacility(r.fid, true) }) : r.name, r.town ? el('div', { class: 'small', text: r.town }) : null), el('td', {}, r.label, r.outcome === 'Pending' ? el('div', { class: 'small', text: 'Pending' }) : null), newOwner, stars, hrs, ser, fin)); });
  if (!rows.length) tb.append(el('tr', {}, el('td', { colspan: 8, text: 'No sales match.' })));
}
['#sale-q', '#sale-parent', '#sale-cms'].forEach(s => $(s).addEventListener('input', drawSales)); drawSales();

/* ---------- facility lookup ---------- */
const FAC = D.facilities.slice().sort((a, b) => a.name.localeCompare(b.name));
const norm = s => (s || '').toLowerCase().replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
function facSearch() {
  const q = norm($('#fac-q').value), ul = $('#fac-results'); ul.replaceChildren(); if (q.length < 2) { ul.classList.add('hidden'); return; }
  const hits = FAC.filter(f => norm([f.name, f.town, f.cms_name, f.owner, ...(f.former || []).map(x => x.name)].join(' ')).includes(q)).slice(0, 12);
  hits.forEach(f => { const fm = (f.former || []).find(x => norm(x.name).includes(q) && !norm(f.name).includes(q)); ul.append(el('li', {}, el('button', { onclick: () => { showFacility(f.id, true); ul.classList.add('hidden'); } }, el('div', {}, f.name), el('div', { class: 'sub', text: [f.town, f.active === false ? (String(f.status || '').startsWith('Closed') ? f.status : 'no longer in CMS files') : null, fm ? 'formerly ' + fm.name : null].filter(Boolean).join(' \u00B7 ') })))); });
  if (!hits.length) ul.append(el('li', {}, el('div', { style: 'padding:8px 12px', text: 'No homes match.' }))); ul.classList.remove('hidden');
}
$('#fac-q').addEventListener('input', facSearch);
const KIND_LBL = { mgmt: 'Buyer took over management', closed: 'Closed', sale: 'Sold / new operator', parent: 'Chain reorganization', internal: 'Internal restructuring', realestate: 'Real estate only', cms: 'CMS data' };
let curFac = null;
function showFacility(id, scroll) {
  const f = D.facilities.find(x => x.id === id); if (!f) return; curFac = id; const box = $('#profile'); box.replaceChildren();
  const gsec = $('#groups'); if (gsec) gsec.classList.add('hidden');            // statewide owner-group table is not about this home
  const closeFac = () => { curFac = null; box.replaceChildren(); if (gsec) gsec.classList.remove('hidden'); if (location.hash.startsWith('#home=')) history.replaceState(null, '', location.pathname + location.search + '#lookup'); };
  const nowB = (label, v, fmt) => el('div', { class: 'badge' }, label, el('b', {}, v == null ? '\u2013' : (fmt ? fmt(v) : v)));
  const head = el('div', { class: 'panel' }, el('button', { class: 'btn', style: 'float:right;margin-left:12px', text: 'Close', onclick: closeFac }), el('h3', { style: 'margin-top:0', text: f.name }), el('div', { class: 'small', text: [f.town, f.county ? f.county + ' County' : null, f.status, f.beds ? f.beds + ' certified beds' : null].filter(Boolean).join(' \u00B7 ') }));
  if (f.former && f.former.length) head.append(el('p', {}, el('b', { text: 'Also known as / formerly: ' }), f.former.map(x => x.name + (x.src === 'CMS listing' && x.to ? ` (CMS name through ${fmtMon(x.to)})` : x.src === 'CEC standard name' ? ' (name used in the state review record)' : '')).join('; ')));
  if (f.owner || f.chain) head.append(el('p', { class: 'small', text: ['CMS lists the operator as ' + (f.owner || 'not stated'), f.chain ? 'chain: ' + f.chain : null, f.ownership_type].filter(Boolean).join(' \u00B7 ') }));
  if (f.now) { head.append(el('div', { class: 'badges' }, el('div', { class: 'badge' }, 'Overall stars', el('div', {}, chip(f.now.overall))), el('div', { class: 'badge' }, 'Health inspection', el('div', {}, chip(f.now.hi))), el('div', { class: 'badge' }, 'Staffing', el('div', {}, chip(f.now.staffing))), el('div', { class: 'badge' }, 'Quality measures', el('div', {}, chip(f.now.qm))),
      nowB('Nurse hours / resident / day', f.now.hours, f2), nowB('Nursing staff turnover', f.now.turnover, v => Math.round(v) + '%')));
    head.append(el('p', { class: 'small', text: `As of the ${fmtMon(f.now.asof)} CMS file.` + (f.sff ? ` CMS special focus status: ${f.sff}.` : '') + (f.abuse ? ' CMS displays an abuse icon for this home.' : '') })); }
  box.append(head);
  if (f.events.length) { const ol = el('ol', { class: 'timeline' }); f.events.forEach(e => ol.append(el('li', {}, el('div', { class: 'small', title: e.date_note || '', text: fmtDay(e.date) }), el('div', {}, el('span', { class: 'kind ' + (e.kind === 'cms' ? 'cms' : e.kind === 'sale' ? 'sale' : e.kind === 'parent' ? 'parent' : e.kind === 'closed' ? 'closed' : e.kind === 'mgmt' ? 'mgmt' : ''), text: KIND_LBL[e.kind] || e.label }), e.kind === 'closed' ? el('span', { text: e.label }) : e.label && e.kind !== 'cms' ? el('span', { text: [e.buyer, e.who].filter(Boolean).join(' \u2014 ') }) : el('span', { text: e.label + (e.who ? ': ' + e.who : '') }), e.outcome ? el('span', { class: 'small', text: ' \u00B7 ' + e.outcome }) : null, el('div', { class: 'small', text: e.src }))))); box.append(el('div', { class: 'panel' }, el('h3', { style: 'margin-top:0', text: 'Ownership timeline' }), el('p', { class: 'small', text: 'State-reviewed changes and changes visible in CMS data, newest first.' }), ol)); }
  else box.append(el('div', { class: 'panel' }, el('p', { text: 'No ownership changes on record for this home.' })));
  if (f.has_cms) {
    const mk = []; f.events.forEach(e => { if (e.kind === 'cms' || e.kind === 'realestate') return; const x = MS(e.date); mk.push({ x, color: e.kind === 'sale' ? css('--sold') : (e.kind === 'closed' ? css('--fg') : css('--accent')), dash: e.kind === 'internal' ? '3 3' : '', title: KIND_LBL[e.kind] + ', ' + fmtDay(e.date) }); });
    const xf = (v, i) => fmtMon(D.dates[i]);
    const c1 = el('div', { class: 'panel' }, el('h3', { style: 'margin-top:0', text: 'Star ratings over time' }), el('div', { class: 'chartbox', id: 'fc1' }), el('div', { class: 'legend' }, ['Overall', 'Health inspection', 'Staffing', 'Quality measures (dotted)'].map((n, i) => el('span', {}, el('i', { class: i === 3 ? 'dot' : '', style: `border-color:${[css('--accent'), css('--s2'), css('--s4'), css('--ctrl')][i]}` }), n)), el('span', {}, el('i', { class: 'vert', style: `border-color:${css('--sold')}` }), 'Sale / new operator (vertical line)'), el('span', {}, el('i', { class: 'vert dash', style: `border-color:${css('--accent')}` }), 'Other change of control (vertical dashed line)')), el('p', { class: 'small', text: 'Where two ratings are the same, the lines sit on top of each other and the darker Overall line is drawn last. A gray band at 0 means CMS did not publish star ratings for that home in those months (special focus facility), so 0 means not rated, not a CMS score.' }));
    const c2 = el('div', { class: 'panel' }, el('h3', { style: 'margin-top:0', text: 'Nurse staffing hours per resident per day' }), el('div', { class: 'chartbox', id: 'fc2' }), el('div', { class: 'legend' }, el('span', {}, el('i', { style: `border-color:${css('--accent')}` }), 'All nurses'), el('span', {}, el('i', { style: `border-color:${css('--s2')}` }), 'RNs'), el('span', {}, el('i', { class: 'dash', style: `border-color:${css('--ctrl')}` }), 'Statewide average, all nurses (dashed)')), el('p', { class: 'small', text: 'Hours are CMS reported hours per resident per day, not adjusted for how sick residents are. A gray band means CMS published an implausibly low figure that month (a payroll-reporting problem), so it is left out.' }));
    const c3 = el('div', { class: 'panel' }, el('h3', { style: 'margin-top:0', text: 'Nursing staff turnover' }), el('div', { class: 'chartbox', id: 'fc3' }), el('div', { class: 'legend' }, el('span', {}, el('i', { style: `border-color:${css('--sold')}` }), 'This home'), el('span', {}, el('i', { class: 'dash', style: `border-color:${css('--ctrl')}` }), 'Statewide average (dashed)')));
    box.append(c1, el('div', { class: 'grid2' }, c2, c3), el('div', { class: 'grid2' }, el('div', { class: 'panel' }, el('h3', { style: 'margin-top:0', text: 'Serious citations per year' }), el('div', { class: 'chartbox', id: 'fc4' })), el('div', { class: 'panel' }, el('h3', { style: 'margin-top:0', text: 'Federal fines per year' }), el('div', { class: 'chartbox', id: 'fc5' }))));
    const runs = (flags, label) => { const out = []; let a = null; flags.forEach((v, i) => { if (v && a === null) a = i; if ((!v || i === flags.length - 1) && a !== null) { out.push({ x0: X[a], x1: X[v ? i : i - 1], label }); a = null; } }); return out; };
    const bands = runs(f.s.nr, '0 = not rated by CMS (special focus facility)');
    const z0 = arr => arr.map((v, i) => (v == null && f.s.nr[i]) ? 0 : v);
    drawChart($('#fc1'), { x: X, ymin: 0, ymax: 5, yfmt: v => String(v), tfmt: v => v === 0 ? '0 (not rated: special focus facility)' : String(v), xlabel: xf, markers: mk, bands, extra: bi => f.s.nr[bi] ? [el('div', { class: 'small', text: 'CMS did not rate this home this month (special focus facility)' })] : [], aria: f.name + ' star ratings over time', series: [{ name: 'Overall', color: css('--accent'), values: z0(f.s.overall), width: 3.4, z: 3 }, { name: 'Health inspection', color: css('--s2'), values: z0(f.s.hi), width: 2.4, z: 2 }, { name: 'Staffing', color: css('--s4'), values: z0(f.s.staffing), width: 2.4, z: 1 }, { name: 'Quality measures', color: css('--ctrl'), values: z0(f.s.qm), dash: '2 5', width: 3.2, z: 0 }] });
    drawChart($('#fc2'), { x: X, ymin: 0, yfmt: v => v.toFixed(1), tfmt: v => v.toFixed(2), xlabel: xf, markers: mk, height: 260, bands: runs(f.s.hf, 'Payroll data problem (CMS gave 1-star staffing): hours left out'), extra: bi => f.s.hf[bi] ? [el('div', { class: 'small', text: 'CMS reported an implausibly low figure this month (payroll data problem); left out here.' })] : [], aria: f.name + ' nurse staffing hours', series: [{ name: 'All nurses', color: css('--accent'), values: f.s.hours, width: 3 }, { name: 'RNs', color: css('--s2'), values: f.s.rn }, { name: 'Statewide avg (all nurses)', color: css('--ctrl'), values: st.avg_hours, dash: '5 4', width: 1.6 }] });
    drawChart($('#fc3'), { x: X, ymin: 0, bands: runs(f.s.tf, 'Payroll data problem: CMS assigns its worst turnover score; left out'), extra: bi => f.s.tf[bi] ? [el('div', { class: 'small', text: 'CMS published an implausible figure this month (payroll data problem); left out here.' })] : [], yfmt: v => Math.round(v) + '%', tfmt: v => Math.round(v) + '%', xlabel: xf, markers: mk, height: 260, aria: f.name + ' nursing staff turnover', series: [{ name: 'This home', color: css('--sold'), values: f.s.turnover, width: 3 }, { name: 'Statewide avg', color: css('--ctrl'), values: st.avg_turnover, dash: '5 4', width: 1.6 }] });
    barChart($('#fc4'), yl, f.yr.serious, css('--s1'), v => String(Math.round(v)), 'Serious citations per year', true); barChart($('#fc5'), yl, f.yr.fines, css('--sold'), v => v >= 1e6 ? '$' + (v / 1e6).toFixed(1) + 'M' : '$' + Math.round(v / 1e3) + 'K', 'Fines per year');
  } else box.append(el('div', { class: 'note', text: 'This home does not appear in the CMS files used here (April 2019 onward), so only the state review record is shown.' }));
  if (scroll) box.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/* ---------- owner groups ---------- */
{ const tb = $('#groups-tbl tbody'); const cell = (p, label) => el('td', { class: 'num' }, p ? [chipAvg(p.before), el('span', { class: 'arrow', text: '\u2192' }), chipAvg(p.after), el('div', { class: 'small', text: p.n + (p.n === 1 ? ' home' : ' homes') })] : el('span', { class: 'small', text: 'not yet' }));
  function chipAvg(v) { return el('b', { text: v.toFixed(1) }); }
  D.groups.forEach(g => { const detail = el('tr', { class: 'hidden' }, el('td', { colspan: 5 }, el('table', {}, el('thead', {}, el('tr', {}, ['Home', 'Sale date', 'At acquisition', 'After 1 year', 'After 2 years', 'Latest (years since)'].map(x => el('th', { text: x })))),
      el('tbody', {}, g.homes.map(h2 => el('tr', {}, el('td', { text: h2.name }), el('td', { text: fmtDay(h2.date) }), el('td', {}, chip(h2.before)), el('td', {}, chip(h2.y1)), el('td', {}, chip(h2.y2)), el('td', {}, chip(h2.latest), h2.yrs != null ? el('span', { class: 'small', text: '  (' + h2.yrs + ' yrs)' }) : null)))))));
    const btn = el('button', { class: 'linkbtn', text: g.group, 'aria-expanded': 'false', onclick: () => { detail.classList.toggle('hidden'); btn.setAttribute('aria-expanded', String(!detail.classList.contains('hidden'))); } });
    tb.append(el('tr', {}, el('td', {}, btn, el('div', { class: 'small', text: 'click to see each home' })), el('td', { class: 'num', text: String(g.n) }), cell(g.y1), cell(g.y2), el('td', { class: 'num' }, g.latest ? [chipAvg(g.latest.before), el('span', { class: 'arrow', text: '\u2192' }), chipAvg(g.latest.after), el('div', { class: 'small', text: g.latest.n + ' homes' + (g.yrs != null ? ', typically ' + g.yrs + ' yrs later' : '') })] : '\u2013')), detail); }); }

/* ---------- about + downloads ---------- */
const ab = $('#about-body');
[['Where the numbers come from', 'CMS Nursing Home Care Compare monthly files (star ratings, staffing, turnover, inspection citations, fines, ownership fields), Rhode Island homes only, and the Rhode Island Department of Health\u2019s annual reports of Change in Effective Control (CEC) reviews.'],
 ['What "sold" means here', 'A change of control approved by the state, most often a sale to a new operator. "Chain reorganizations" are the records the state log marks as a corporate parent change rather than a sale to a different buyer; since April 2019 that is only the 2022 Tryko-to-Marquis reorganization of four homes. Every other change of control on this page is a sale to a new owner or operator. Estate and internal restructurings and real-estate-only deals are shown on a home\u2019s timeline but not counted as sales.'],
 ['Dates', 'The sale date is CMS\u2019s takeover date when CMS lists one, otherwise the state\u2019s decision date. CMS files usually show a new operator several months later. The date shown in the CMS-data timeline entries is when CMS\u2019s files first showed the change.'],
 ['Star ratings', 'CMS revised how star ratings are calculated in April 2019 (the statewide average fell from 3.8 to 3.35 in a single month) and has adjusted its methods since. Charts start after that change. Ratings are one measure of a home, not a full picture of care.'],
 ['Why measures can point in different directions', 'The overall star rating starts from the health inspection rating and is then adjusted up or down by the staffing and quality-measure ratings, so a home can add staff and still lose stars if its inspection results or quality measures got worse. The health inspection rating itself is built from the last three years of inspections, most recent weighted most, and counts all deficiencies by scope and severity, not just serious ones, so it can fall in a year with fewer serious citations. Fines are not part of the star rating. The staffing hours on this page are CMS reported hours; the staffing star compares case-mix-adjusted hours with national thresholds.'],
 ['Payroll data problems', 'When a home fails to submit payroll-based staffing data, fails an audit, or reports 7 or more days with no RN, CMS gives it 1 staffing star, and the hours it publishes can be implausibly low. Reported hours under 1.5 per resident per day, and turnover of 90% or more (CMS assigns its worst turnover score when payroll data is missing), are treated as that kind of problem and left out (gray bands on the hours and turnover charts).'],
 ['Homes without a star rating', 'CMS does not publish star ratings for homes in its Special Focus Facility program, which targets homes with a history of serious quality problems. Those months are counted as 0 stars (0 is not a CMS score) so that a home being in the program shows up in statewide averages, the share of low-rated homes, the star-mix chart (gray layer) and the sold-versus-not-sold comparison, including in the downloadable CSVs.'],
 ['Money paid to related parties', 'Each home files a Medicare cost report that lists payments to related parties: companies the same owners control, such as the landlord or a management company. Those payments are costs to the home, so a home can report thin or negative margins while its owners are paid through rent and fees. The before-and-after chart compares the same homes between the year before the sale and two years after, which is why its home counts are smaller than the ratings comparison.'],
 ['"CMS records updated" dates', 'CMS\u2019s files are not updated the day a sale closes. Where a sale has one, this page shows the first monthly CMS file listing the new operator; across these sales that came a median of about eight months after the date of record. It is a reporting lag, not a second event, and it matters for two reasons: ratings in those months were still filed under the old operator\u2019s name, and a CMS operator change with no matching state record is a sign of a change that may not have gone through review.'],
 ['Management agreements and timing', 'A buyer sometimes takes over day-to-day management of a home under a management agreement before the state approves the sale, and CMS can take months to list a new operator. Where a management start date is known it is shown on the home’s timeline and beside the sale. So the last rating before a sale may already reflect the buyer\u2019s management, and the sales table shows when CMS first lists the new operator. These agreements are not published. They have been requested through APRA from RIDOH.'],
 ['The comparison line', 'Comparable homes are those with no change of control or new operator from a year before to two years after each sale. They are averaged over the same calendar dates as the sold homes, so statewide trends (the pandemic, staffing rules, rating changes) appear in both lines.'],
 ['What this does not show', 'Why ratings changed. A sale can happen because a home was struggling, and many homes were already slipping before the sale. With a few dozen sales, averages can be moved by a handful of homes; the counts are shown.'],
 ['Homes not shown in charts', 'Homes with no CMS data in these files (for example some Brookdale nursing beds and one pediatric facility) appear with their state review record only. Homes that closed or left the CMS files are marked.'],
 ['Corrections', '']].forEach(a => ab.append(el('details', { open: a[0] === 'Where the numbers come from' ? '' : null }, el('summary', { text: a[0] }), a[0] === 'Corrections' ? el('p', {}, D.meta.contact ? ['Send corrections to ', el('a', { href: 'mailto:' + D.meta.contact, text: D.meta.contact }), '.'] : 'Please report corrections to the organization that published this page.') : el('p', { text: a[1] }))));
$('#foot').append(el('p', {}, el('b', {}, el('a', { href: 'https://abc-ri.org', text: 'ABC-RI' })), ' \u00B7 Advocates for Better Care in Rhode Island'),
  el('p', {}, [D.meta.credit, 'Data through ' + D.meta.updated + '. Page built ' + D.meta.built + '.'].filter(Boolean).join(' \u00B7 '), D.meta.contact ? [' \u00B7 Corrections: ', el('a', { href: 'mailto:' + D.meta.contact, text: D.meta.contact })] : null));
function download(name, rows) { const esc = v => { v = v == null ? '' : String(v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }; const blob = new Blob([rows.map(r => r.map(esc).join(',')).join('\n')], { type: 'text/csv' }); const a = el('a', { href: URL.createObjectURL(blob), download: name }); document.body.append(a); a.click(); a.remove(); }
$('#dl-sales').addEventListener('click', () => download('ri_nursing_home_sales.csv', [['cec_id', 'date', 'date_source', 'home', 'town', 'type', 'outcome', 'buyer_group', 'applicant', 'stars_before', 'stars_1yr', 'stars_2yr', 'nurse_hrs_before', 'nurse_hrs_1yr', 'serious_cites_before', 'serious_cites_after', 'fines_before', 'fines_after']].concat(D.sales.map(r => [r.cec, r.date, r.date_note, r.name, r.town, r.label, r.outcome, r.buyer, r.who, r.m.overall.before, r.m.overall.y1, r.m.overall.y2, r.m.hours.before, r.m.hours.y1, r.ser.before, r.ser.y1, r.fines.before, r.fines.y1]))));
$('#dl-monthly').addEventListener('click', () => { const rows = [['ccn', 'home', 'file_date', 'overall_stars', 'health_inspection_stars', 'staffing_stars', 'quality_measure_stars', 'nurse_hours_per_resident_day', 'rn_hours_per_resident_day', 'nursing_staff_turnover_pct']]; D.facilities.filter(f => f.has_cms).forEach(f => D.dates.forEach((d, i) => { if (f.s.overall[i] != null || f.s.hours[i] != null) rows.push([f.ccn, f.name, d, f.s.overall[i], f.s.hi[i], f.s.staffing[i], f.s.qm[i], f.s.hours[i], f.s.rn[i], f.s.turnover[i]]); })); download('ri_nursing_home_monthly.csv', rows); });
let lastW = window.innerWidth; let rz = null;
window.addEventListener('resize', () => { clearTimeout(rz); rz = setTimeout(() => { if (Math.abs(window.innerWidth - lastW) < 30) return; lastW = window.innerWidth; drawTrend(); drawStudy(); drawMoney(); barChart($('#ser-chart'), yl, D.state_years.serious, css('--s1'), v => String(Math.round(v)), 'Serious citations per year', true); barChart($('#fine-chart'), yl, D.state_years.fines, css('--sold'), v => v >= 1e6 ? '$' + (v / 1e6).toFixed(1) + 'M' : '$' + Math.round(v / 1e3) + 'K', 'Federal fines per year'); if (curFac) showFacility(curFac, false); }, 200); });
if (location.hash.startsWith('#home=')) { showFacility(decodeURIComponent(location.hash.slice(6))); }
</script>
</body>
</html>
'''

LOGO_PNG_B64 = "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"

def build_page(D, out_path):
    """Writes index.html and, beside it, abcri-logo.png (the page links to it as a sibling file)."""
    js = json.dumps(D, separators=(',', ':'), default=lambda o: None, ensure_ascii=False).replace('</', '<\\/')
    html = PAGE_TEMPLATE.replace('__TITLE__', re.sub(r'<[^>]+>', '', D['meta']['title'])).replace('__DATA__', js)
    with open(out_path, 'w', encoding='utf-8') as f: f.write(html)
    import base64 as _b64
    logo = os.path.join(os.path.dirname(out_path) or '.', 'abcri-logo.png')
    with open(logo, 'wb') as f: f.write(_b64.b64decode(LOGO_PNG_B64))
    return out_path
