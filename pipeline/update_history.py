#!/usr/bin/env python3
"""Add the newest CMS month to the Rhode Island history files used by the sales page.

Downloads the four CMS files the builder uses (ProviderInfo, HealthCitations, Penalties, Ownership)
from CMS's public Provider Data catalog, keeps the Rhode Island rows, standardizes the column names the
same way the Colab notebook does, and appends them to pipeline/history/*.csv.gz with a snapshot_date.

Prints "NOTHING NEW" and exits 0 if that CMS month is already in the history.
Exits 1 with a plain message if anything looks wrong, so nothing is published.
Usage: update_history.py [--force]
"""
import csv, gzip, io, json, os, re, sys, time, urllib.error, urllib.request
from datetime import datetime
import pandas as pd

CATALOG = "https://data.cms.gov/provider-data/api/1/metastore/schemas/dataset/items?show-reference-ids=false"
HERE = os.path.dirname(os.path.abspath(__file__))
HIST = os.path.join(HERE, "history")
FILES = {"providerinfo": "ProviderInfo", "healthcitations": "HealthCitations",
         "penalties": "Penalties", "ownership": "Ownership"}
STATE_COLS = ["state", "provider_state"]
CCN_COLS = ["cms_certification_number_ccn", "federal_provider_number", "provnum", "ccn", "provider_id"]


def die(msg):
    print("ERROR:", msg)
    sys.exit(1)


def get(url, tries=4):
    """Fetch a URL, retrying a few times (CMS's site has short outages). CMS refuses requests that
    don't look like a normal web request, so send a browser-style User-Agent and Accept header.
    On final failure, stop with a one-line error that says what CMS answered."""
    waits = [20, 60, 180]
    last = None
    for n in range(tries):
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0 (compatible; abc-ri-update/1.1; +https://snfsales.abc-ri.org)",
            "Accept": "application/json, text/csv, */*"})
        try:
            with urllib.request.urlopen(req, timeout=300) as r:
                return r.read()
        except urllib.error.HTTPError as e:
            body = e.read()[:300].decode("utf-8", errors="replace").replace("\n", " ")
            last = f"CMS answered HTTP {e.code} {e.reason} for {url[:90]} ({body.strip()[:200]})"
        except Exception as e:
            last = f"{type(e).__name__}: {e} for {url[:90]}"
        print(f"try {n + 1} of {tries} failed: {last}", flush=True)
        if n < tries - 1:
            time.sleep(waits[min(n, len(waits) - 1)])
    die(last)


def norm_col(c):
    return re.sub(r"[^a-z0-9]+", "_", str(c).strip().lower()).strip("_")


def to_std(raw_bytes, snapshot_date, source):
    """Raw CMS CSV bytes -> RI-only frame in the history schema."""
    text = raw_bytes.decode("utf-8-sig", errors="replace")
    df = pd.read_csv(io.StringIO(text), dtype=str, low_memory=False)
    df.columns = [norm_col(c) for c in df.columns]
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    ccn = next((c for c in CCN_COLS if c in df.columns), None)
    if ccn is None:
        die(f"{source}: no facility ID column found. Columns: {list(df.columns)[:12]}")
    st = next((c for c in STATE_COLS if c in df.columns), None)
    keep = df[st].fillna("").str.strip().str.upper().eq("RI") if st else df[ccn].fillna("").str.strip().str.zfill(6).str.startswith("41")
    df = df[keep].copy()
    df.insert(0, "ccn_std", df[ccn].astype(str).str.strip().str.replace(r"\.0$", "", regex=True).str.zfill(6))
    df.insert(0, "snapshot_date", snapshot_date)
    df.insert(2, "snapshot", f"nursing-homes_{snapshot_date[:4]}")
    df.insert(3, "source_path", source)
    return df


def append(name, new):
    path = os.path.join(HIST, name + ".csv.gz")
    old = pd.read_csv(path, dtype=str, low_memory=False)
    snap = new["snapshot_date"].iloc[0]
    old = old[old["snapshot_date"] != snap]            # re-running a month replaces it
    missing = [c for c in old.columns if c not in new.columns]
    extra = [c for c in new.columns if c not in old.columns]
    if len(missing) > 0.25 * len(old.columns):
        die(f"{name}: new CMS file is missing {len(missing)} of {len(old.columns)} expected columns, e.g. {missing}. CMS may have renamed columns.")
    if extra:
        print(f"NOTE {name}: {len(extra)} new CMS column(s) not kept: {extra[:6]}")
    if missing:
        print(f"NOTE {name}: {len(missing)} expected column(s) absent this month (left blank): {missing}")
    out = pd.concat([old, new.reindex(columns=old.columns)], ignore_index=True)
    out.to_csv(path, index=False, compression="gzip")
    print(f"{name}: added {len(new)} RI rows for {snap}; history now {len(out)} rows")


def main(force=False):
    items = json.loads(get(CATALOG))
    # Look at every CSV the catalog lists and match on the file name itself.
    # The nursing home files are named like NH_ProviderInfo_Sep2026.csv.
    cands, seen = {}, []
    for i in items:
        mod = str(i.get("modified", ""))[:10]
        for d in i.get("distribution", []) or []:
            if isinstance(d, dict) and isinstance(d.get("data"), dict):
                d = d["data"]                      # the catalog nests the real entry under "data"
            if not isinstance(d, dict):
                continue
            u = d.get("downloadURL") or ""
            path = u.split("?")[0]
            if not path.lower().endswith(".csv"):
                continue
            base = os.path.basename(path)
            if re.search(r"NH[_-]|nursing", base, re.I):
                seen.append(base)
            for key, frag in FILES.items():
                if re.search(r"(^|[_-])NH[_-]" + frag + r"([_.]|$)", base, re.I):
                    cands.setdefault(key, []).append((mod, u, base))
    miss = [k for k in FILES if k not in cands]
    if miss:
        die(f"CMS catalog did not list: {miss}. CSV names seen that look like nursing home files: {sorted(set(seen))[:12]}")
    urls, newest = {}, ""
    for key, lst in cands.items():
        mod, u, base = sorted(lst)[-1]          # newest if the catalog lists more than one
        urls[key] = (u, base)
        newest = max(newest, mod)
    print("CMS files: " + ", ".join(sorted(b for _, b in urls.values())))
    snapshot_date = newest or datetime.now().strftime("%Y-%m-%d")
    raws = {k: get(u) for k, (u, _) in urls.items()}
    # which month is this? read it from the provider file
    probe = to_std(raws["providerinfo"], snapshot_date, urls["providerinfo"][1])
    col = next((c for c in ("processing_date", "filedate") if c in probe.columns), None)
    if col is None or probe[col].dropna().empty:
        die("could not read the processing date from the CMS provider file")
    month = pd.to_datetime(probe[col], errors="coerce", format="mixed").dt.strftime("%Y-%m").mode().iloc[0]
    # Use the date CMS stamps on the data (like the earlier months), not the catalog's "modified" date.
    pdate = pd.to_datetime(probe[col], errors="coerce", format="mixed").dropna()
    snapshot_date = pdate.dt.strftime("%Y-%m-%d").mode().iloc[0]
    have = pd.read_csv(os.path.join(HIST, "providerinfo.csv.gz"), dtype=str, usecols=["snapshot_date"])["snapshot_date"]
    have_months = {d[:7] for d in have.unique()}
    if month in have_months and not force:
        print(f"NOTHING NEW: CMS month {month} is already in the history.")
        return
    print(f"catalog date {newest}; data date {snapshot_date}; CMS month {month}; newest saved {have.max()}; RI rows {len(probe)}")
    if snapshot_date <= have.max() and not force:
        die(f"new data is dated {snapshot_date}, not later than the newest saved ({have.max()}). Use force if this is intended.")
    if len(probe) < 50:
        die(f"only {len(probe)} Rhode Island providers found in the new CMS file (expected about 70-90)")
    for key, (u, base) in urls.items():
        append(key, to_std(raws[key], snapshot_date, base))
    print(f"READY: added CMS month {month} (snapshot {snapshot_date}).")


def inspect():
    raw = get(CATALOG)
    print(f"catalog bytes {len(raw)}")
    items = json.loads(raw)
    if isinstance(items, dict):
        print("top-level keys:", list(items)[:10])
        items = items.get("data") or items.get("results") or []
    print(f"items {len(items)}")
    if items:
        print("item keys:", list(items[0])[:20])
    nh = [i for i in items if "nursing" in json.dumps(i.get("title", "")).lower()]
    print(f"items with 'nursing' in the title: {len(nh)}")
    for i in nh[:4]:
        print("TITLE:", str(i.get("title"))[:90], "| modified", i.get("modified"))
        for d in (i.get("distribution") or [])[:4]:
            if isinstance(d, dict):
                print("   dist keys:", list(d)[:8], "| url:", str(d.get("downloadURL") or d.get("accessURL") or d.get("identifier"))[:140])
            else:
                print("   dist (not a dict):", str(d)[:140])
    if not nh and items:
        print("sample titles:", [str(i.get("title"))[:40] for i in items[:6]])


if __name__ == "__main__":
    if "--inspect" in sys.argv:
        inspect()
    else:
        main(force="--force" in sys.argv)
