"""Additions to the public page data, made after build_public_data().

- Names: each home gets the name ABC-RI uses on the SNF dashboard (the Facility extras sheet's display_name),
  and each buyer gets the owner group name the dashboard uses (pipeline/group_names.json), so the two sites match.
- Tenures: for every sale, when the buyer began running the home (the earlier of the sale date and the start of any
  management agreement) and when it stopped (the next sale to a different group), with the citations and fines CMS
  recorded in between.
- Track records: for each owner group, its record across every Rhode Island home it has run, counted only while it
  ran them, plus CMS's national chain figures when CMS lists the group as a chain.

Kept separate from ri_nh_build.py so it can be tested on its own:
    enrich(D, hu, pu, extras=None, chains=None, names=None)
D is the page data, hu and pu the CMS citation and penalty tables, extras a DataFrame of the Facility extras sheet,
chains the rows of CMS's Nursing Home Chain Performance Measures, names the contents of group_names.json.
"""
import re
from datetime import date

import pandas as pd

IJ = set("JKL")
HARM = set("GHIJKL")
CHAIN_FIELDS = {   # CMS column -> short key
    "Number of facilities": "homes", "Number of states and territories with operations": "states",
    "Number of Special Focus Facilities (SFF)": "sff", "Number of SFF candidates": "sff_cand",
    "Percentage of facilities with an abuse icon": "abuse_pct", "Average overall 5-star rating": "overall",
    "Average health inspection rating": "health", "Average staffing rating": "staffing",
    "Average total nurse hours per resident day": "hours", "Average total nursing staff turnover percentage": "turnover",
    "Average amount of fines in dollars": "fines_avg", "Total amount of fines in dollars": "fines_total",
}


def _num(v):
    try:
        return float(str(v).replace(",", "").replace("$", ""))
    except (TypeError, ValueError):
        return None


def _clean_name(n):
    """'Respiratory and Rehabilitation Center of Ri' -> '... of RI'; 'X, LLC' -> 'X'."""
    if not n:
        return n
    n = re.sub(r",?\s+(LLC|L\.L\.C\.|Inc\.?|LP|Corp\.?)$", "", str(n).strip(), flags=re.I)
    return re.sub(r"\b(Ri)\b", "RI", n)


def _t(d):
    return pd.Timestamp(d) if d else None


def enrich(D, hu, pu, extras=None, chains=None, names=None):
    names = names or {}
    alias = names.get("aliases", {})
    fill = names.get("fill", {})
    g = lambda b: alias.get(b, b) if b else b

    # ---------------- names ----------------
    disp, owner_group = {}, {}
    if extras is not None:
        for r in extras.fillna("").itertuples():
            c = str(r.ccn).replace(".0", "").strip()
            if c:
                if str(r.display_name).strip():
                    disp[c] = str(r.display_name).strip()
                if str(getattr(r, "owner_group", "")).strip():
                    owner_group[c] = str(r.owner_group).strip()
    renamed = {}
    for f in D["facilities"]:
        new = disp.get(f["id"]) or _clean_name(f["name"])
        if new != f["name"]:
            renamed[f["name"]] = new
            f.setdefault("legal_name", f["name"])
            f["name"] = new
        if f["id"] in owner_group:
            f["owner_group"] = owner_group[f["id"]]
    fac = {f["id"]: f for f in D["facilities"]}
    for s in D["sales"]:
        f = fac.get(s.get("fid"))
        s["name"] = f["name"] if f else _clean_name(renamed.get(s["name"], s["name"]))
        if not s.get("buyer") and s.get("cec") in fill:
            s["buyer"] = fill[s["cec"]]
        s["buyer"] = g(s.get("buyer"))
    for e_list in (f.get("events") or [] for f in D["facilities"]):
        for e in e_list:
            if e.get("buyer"):
                e["buyer"] = g(e["buyer"])
    merged = {}
    for grp in D.get("groups", []):
        grp["group"] = g(grp["group"])
        for h in grp.get("homes", []):
            h["name"] = renamed.get(h["name"], _clean_name(h["name"]))
        merged.setdefault(grp["group"], []).append(grp)
    if any(len(v) > 1 for v in merged.values()):   # two old names now mean one group: keep the larger, add the homes
        out = []
        for name, gs in merged.items():
            gs.sort(key=lambda x: -x["n"])
            for extra in gs[1:]:
                gs[0]["homes"] += extra["homes"]
                gs[0]["n"] += extra["n"]
            out.append(gs[0])
        D["groups"] = out

    # ---------------- tenures ----------------
    hu = hu.assign(sd=pd.to_datetime(hu.survey_date, errors="coerce"))
    hu = hu.drop_duplicates(["ccn_std", "survey_date", "deficiency_tag_number", "scope_severity_code"])
    fines = pu[pu.penalty_type == "Fine"].assign(pd_=pd.to_datetime(pu.penalty_date, errors="coerce"))
    fines = fines.drop_duplicates(["ccn_std", "penalty_date", "fine_amount"])
    latest = pd.Timestamp(D["dates"][-1])
    floor = pd.Timestamp(D["dates"][0]) - pd.DateOffset(years=3)   # CMS files from April 2019 reach back about 3 years
    dates = [pd.Timestamp(x) for x in D["dates"]]

    def stars_at(fid, when, before=True):
        f = fac.get(fid)
        if not f or not f.get("s"):
            return None
        v = f["s"].get("overall") or []
        idx = [i for i, d in enumerate(dates) if (d < when if before else d >= when) and i < len(v) and v[i] is not None]
        if not idx:
            return None
        return v[idx[-1] if before else idx[0]]

    by_home = {}
    for s in D["sales"]:
        if s.get("fid") in fac and s.get("date") and s.get("kind") in ("sale", "parent") and s.get("buyer"):
            # A pending sale counts only when the buyer is already managing the home under a management agreement:
            # from that date the home's record is the buyer's, not the seller's.
            if s.get("outcome") == "Pending" and not (s.get("mgmt_start") and _t(s["mgmt_start"]) <= latest):
                continue
            by_home.setdefault(s["fid"], []).append(s)
    tenures = []

    def add_tenure(fid, grp, start, end, cec, via=None, owner_from=None, pending=False):
        f = fac[fid]
        lo, hi = max(start, floor) if start is not None else floor, min(end or latest, latest) + pd.Timedelta(days=1)
        if hi <= lo:
            return
        h = hu[(hu.ccn_std == fid) & (hu.sd >= lo) & (hu.sd < hi)]
        p = fines[(fines.ccn_std == fid) & (fines.pd_ >= lo) & (fines.pd_ < hi)]
        sc = h.scope_severity_code.astype(str)
        current = end is None or end > latest     # a sale CMS doesn't show yet hasn't changed who runs the home
        tenures.append(dict(
            fid=fid, name=f["name"], group=grp, cec=cec, start=start.strftime("%Y-%m-%d") if start is not None else None,
            end=end.strftime("%Y-%m-%d") if (end is not None and not current) else None, counted_from=lo.strftime("%Y-%m-%d"),
            years=round((hi - lo).days / 365.25, 1), current=current,
            via=via, owner_from=owner_from, pending=bool(pending),   # via: took over 'management' (an agreement before the sale) or 'ownership' (the sale); None before the records
            stars_start=stars_at(fid, start) if start is not None else None,
            stars_end=(f.get("now") or {}).get("overall") if current else stars_at(fid, end),
            harm=int(sc.isin(HARM).sum()), ij=int(sc.isin(IJ).sum()), fines=float(p.fine_amount.sum()), n_fines=int(len(p))))

    for fid, ss in by_home.items():
        ss.sort(key=lambda x: x["date"])
        for i, s in enumerate(ss):
            if i and g(ss[i - 1]["buyer"]) == s["buyer"]:
                continue   # a reorganization inside the same group does not start a new tenure
            start, via = _t(s["date"]), "ownership"
            if s.get("mgmt_start") and _t(s["mgmt_start"]) < start:
                start, via = _t(s["mgmt_start"]), "management"
            nxt = next((x for x in ss[i + 1:] if x["buyer"] != s["buyer"]), None)
            end = _t(nxt["date"]) if nxt else None
            if nxt and nxt.get("mgmt_start") and _t(nxt["mgmt_start"]) < end:
                end = _t(nxt["mgmt_start"])
            f = fac[fid]
            if end is None and not f.get("active", True):
                end = _t(f.get("last_seen"))
            if end is not None and end <= floor:
                continue   # entirely before CMS's records
            pend = s.get("outcome") == "Pending"
            add_tenure(fid, s["buyer"], start, end, s.get("cec"), via=via, owner_from=None if pend else s["date"], pending=pend)
    # Homes a group runs now with no sale on record (it has run them since before the state's records): counted from
    # the start of CMS's records. Only when the home has no recorded sale at all, so nothing is guessed.
    for fid, grp in owner_group.items():
        if fid in fac and fid not in by_home and grp not in ("Nonprofit", "") and fac[fid].get("active", True):
            add_tenure(fid, grp, None, None, None)
    D["tenures"] = tenures

    # ---------------- track records ----------------
    chain_rows = {str(r.get("Chain", "")).strip().upper(): r for r in (chains or [])}
    natl = chain_rows.get("NATIONAL")
    manual_chain = names.get("cms_chain", {})
    recs = {}
    for t in tenures:
        recs.setdefault(t["group"], []).append(t)
    track = {}
    for grp, ts in recs.items():
        cur = [t for t in ts if t["current"]]
        both = [t for t in ts if t["stars_start"] is not None and t["stars_end"] is not None]
        r = dict(group=grp, homes=len({t["fid"] for t in ts}), homes_now=len(cur), years=round(sum(t["years"] for t in ts), 1),
                 harm=sum(t["harm"] for t in ts), ij=sum(t["ij"] for t in ts), fines=round(sum(t["fines"] for t in ts)),
                 n_stars=len(both),
                 stars_start=round(sum(t["stars_start"] for t in both) / len(both), 1) if both else None,
                 stars_end=round(sum(t["stars_end"] for t in both) / len(both), 1) if both else None,
                 first=min((t["start"] or t["counted_from"]) for t in ts), before_records=any(t["start"] is None for t in ts))
        # CMS chain: the chain CMS lists for the group's current homes, or one named in group_names.json
        cn = manual_chain.get(grp)
        if not cn:
            cms = [fac[t["fid"]].get("chain") for t in cur if fac[t["fid"]].get("chain")]
            cn = max(set(cms), key=cms.count) if cms else None
        row = chain_rows.get(str(cn).upper()) if cn else None
        if row:
            r["chain"] = dict(name=row.get("Chain"), **{k: _num(row.get(c)) for c, k in CHAIN_FIELDS.items()})
        track[grp] = r
    # pending buyers with no Rhode Island record still get a national line if group_names.json names their CMS chain
    for s in D["sales"]:
        b = s.get("buyer")
        if s.get("outcome") == "Pending" and b and b not in track:
            row = chain_rows.get(str(manual_chain.get(b, "")).upper())
            track[b] = dict(group=b, homes=0, homes_now=0, years=0, harm=0, ij=0, fines=0, n_stars=0, stars_start=None, stars_end=None, first=None)
            if row:
                track[b]["chain"] = dict(name=row.get("Chain"), **{k: _num(row.get(c)) for c, k in CHAIN_FIELDS.items()})
    D["track"] = track
    if natl:
        D["chain_national"] = {k: _num(natl.get(c)) for c, k in CHAIN_FIELDS.items()}
    D["chain_asof"] = names.get("_chain_asof")
    return D
