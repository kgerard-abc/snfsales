#!/usr/bin/env python3
"""ABC-RI site health check.

Run from the repo root:   python3 .github/scripts/check_site.py [--live]

It works out which of the three sites it is looking at, checks the data for
problems, and exits 1 if it finds any ERROR. WARNINGs are printed but do not
fail the run. With --live it also fetches the published site.
"""
import datetime as dt
import json
import os
import re
import sys
import urllib.request
from collections import Counter

errors, warns = [], []
TODAY = dt.date.today()


def err(msg):
    errors.append(msg)


def warn(msg):
    warns.append(msg)


def read(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


def months_old(ym):
    y, m = int(ym[:4]), int(ym[5:7])
    return (TODAY.year - y) * 12 + TODAY.month - m


def stray_tokens(text, label):
    for tok in ("NaN", "undefined", "Infinity"):
        n = len(re.findall(r"(?<![A-Za-z])" + tok + r"(?![A-Za-z])", text))
        if n:
            err(f"{label}: {n} stray {tok} value(s) in the data")


def check_cname(expect):
    if not os.path.exists("CNAME"):
        err("CNAME file is missing, so the custom domain will drop off")
    elif read("CNAME").strip() != expect:
        err(f'CNAME says "{read("CNAME").strip()}", expected "{expect}"')


def json_script(html, script_id):
    m = re.search(r'<script id="%s" type="application/json">(.*?)</script>' % script_id, html, re.S)
    if not m:
        err(f'index.html has no <script id="{script_id}"> data block')
        return None
    try:
        return json.loads(m.group(1))
    except Exception as e:
        err(f"data block {script_id} is not valid JSON: {e}")
        return None


# ---------------------------------------------------------------- SNF dashboard
def check_snf():
    for f in ("index.html", "data.js"):
        if not os.path.exists(f):
            err(f"{f} is missing")
            return
    check_cname("snfdashboard.abc-ri.org")
    html, js = read("index.html"), read("data.js")
    if 'src="data.js"' not in html:
        err("index.html does not load data.js")
    stray_tokens(js, "data.js")
    m = re.match(r"\s*window\.RINH_DATA\s*=\s*", js)
    if not m:
        err("data.js does not start with window.RINH_DATA=")
        return
    try:
        d = json.loads(js[m.end():].rstrip().rstrip(";"))
    except Exception as e:
        err(f"data.js is not valid JSON: {e}")
        return

    months, meta, facs = d["months"], d["meta"], d["facilities"]
    if months != sorted(months) or len(set(months)) != len(months):
        err("months list is not sorted and unique")
    if meta["latest"] != months[-1]:
        err(f"meta.latest {meta['latest']} is not the last month {months[-1]}")
    label = dt.date(int(meta["latest"][:4]), int(meta["latest"][5:7]), 1).strftime("%b %Y")
    if meta.get("latest_label") != label:
        err(f"latest_label '{meta.get('latest_label')}' should be '{label}'")
    lag = months_old(meta["latest"])
    if lag >= 3:
        err(f"CMS data is {lag} months old (latest {meta['latest']}): the monthly update is overdue")
    elif lag == 2:
        warn(f"Latest CMS month is {meta['latest']}: next update due")

    ccns = [f["ccn"] for f in facs if f["ccn"]]
    dupes = [c for c, n in Counter(ccns).items() if n > 1]
    if dupes:
        err(f"duplicate CCNs: {dupes}")
    cc_set = set(ccns)
    metric_keys = [m_["key"] for m_ in d["metrics"]] + ["sff", "abuse"]
    for ccn, ser in d["series"].items():
        if ccn not in cc_set:
            err(f"series has CCN {ccn} that is not in facilities")
        for k, arr in ser.items():
            if k not in metric_keys:
                warn(f"series {ccn} has unknown metric {k}")
            if len(arr) != len(months):
                err(f"series {ccn}/{k} has {len(arr)} values for {len(months)} months")
            elif k in ("overall", "health", "qm", "staffing"):
                bad = [v for v in arr if v is not None and not (0 <= v <= 5)]
                if bad:
                    err(f"series {ccn}/{k} has ratings outside 0 to 5: {bad[:3]}")
    no_xy, no_grp = [], []
    for f in facs:
        if f["status"] == "Open" and f["last_month"] != meta["latest"]:
            err(f"{f['name']} ({f['ccn']}) is Open but its last month is {f['last_month']}, not {meta['latest']}")
        if f.get("lat") is None or f.get("lon") is None:
            no_xy.append(f["name"])
        elif not (41.1 < f["lat"] < 42.1 and -71.95 < f["lon"] < -71.0):
            err(f"{f['name']} map coordinates fall outside Rhode Island")
        if not f.get("x", {}).get("owner_group"):
            no_grp.append(f["name"])
    if no_xy:
        warn(f"{len(no_xy)} facilities have no map coordinates: {', '.join(no_xy)}")
    if no_grp:
        warn(f"{len(no_grp)} facilities have no owner_group: {', '.join(no_grp)}")
    # harm summaries: every key must match a real citation, and none may be blank
    cit_keys = set()
    for ccn, rows in d["citations"].items():
        for r in rows:
            cit_keys.add(f"{ccn}|{r[0]}|{r[1]}")
    for key, text in d["harm_summaries"].items():
        if not str(text).strip():
            err(f"harm summary {key} is blank")
        if key not in cit_keys:
            warn(f"harm summary {key} matches no citation")
    # recent harm-level citations with no plain-language summary
    cutoff = (TODAY - dt.timedelta(days=120)).isoformat()
    missing = []
    for ccn, rows in d["citations"].items():
        for r in rows:
            if r[0] >= cutoff and r[2] in ("G", "H", "I", "J", "K", "L") and f"{ccn}|{r[0]}|{r[1]}" not in d["harm_summaries"]:
                missing.append(f"{ccn}|{r[0]}|{r[1]}")
    if missing:
        warn(f"{len(missing)} harm citation(s) from the last 120 days have no summary yet (first: {missing[0]})")
    if not os.path.exists("abcri-logo.png") and "data:image" not in html:
        err("logo is neither a file nor embedded")


# ---------------------------------------------------------------- CEC sales site
def check_sales():
    if not os.path.exists("index.html"):
        err("index.html is missing")
        return
    check_cname("snfsales.abc-ri.org")
    html = read("index.html")
    d = json_script(html, "data")
    if d is None:
        return
    stray_tokens(json.dumps(d), "page data")
    if "abcri-logo.png" in html and not os.path.exists("abcri-logo.png"):
        err("index.html uses abcri-logo.png but the file is not in the repo (broken logo)")
    dates = d["dates"]
    if dates != sorted(dates):
        err("dates are not in order")
    n = len(dates)
    meta = d["meta"]
    if meta["latest"] != dates[-1]:
        err(f"meta.latest {meta['latest']} is not the last date {dates[-1]}")
    lag = months_old(meta["latest"])
    if lag >= 3:
        err(f"CMS data is {lag} months old (latest {meta['latest']}): the monthly update is overdue")
    exp = dt.date.fromisoformat(meta["latest"]).strftime("%b %Y")
    if meta.get("updated") != exp:
        err(f"meta.updated '{meta.get('updated')}' should be '{exp}'")
    ids = [f["id"] for f in d["facilities"]]
    if len(ids) != len(set(ids)):
        err("duplicate facility ids")
    for f in d["facilities"]:
        for k, arr in (f.get("s") or {}).items():
            if len(arr) != n:
                err(f"{f['name']}: series {k} has {len(arr)} points for {n} dates")
    fid_set = set(ids)
    cecs = [s["cec"] for s in d["sales"]]
    if len(cecs) != len(set(cecs)):
        err("duplicate CEC ids in sales")
    nodate = [s["cec"] for s in d["sales"] if not parse(s.get("date"))]
    if nodate:
        err(f"{len(nodate)} sales have no date and may show as 'Invalid Date' on the page: {', '.join(nodate)}")
    for s in d["sales"]:
        if s.get("has_cms") and s.get("fid") not in fid_set:
            err(f"{s['cec']} points to facility {s.get('fid')} which is not in facilities")
    if d["state"]["n"] and len(d["state"]["n"]) != n:
        err("state series length does not match dates")
    if "corrections@abc-ri.org" not in html:
        warn("corrections address not found on the page")


def parse(s):
    try:
        return dt.date.fromisoformat(str(s)[:10])
    except Exception:
        return None


# ---------------------------------------------------------------- ALF inspections
def check_alf():
    if not os.path.exists("index.html"):
        err("index.html is missing")
        return
    check_cname("alinspections.abc-ri.org")
    html = read("index.html")
    data = json_script(html, "facility-data")
    if data is None:
        return
    stray_tokens(json.dumps(data), "facility data")
    if os.path.exists("maptest.html"):
        warn("maptest.html (a test copy of the site) is public in the repo: delete it when no longer needed")
    keys = [f["key"] for f in data]
    if len(keys) != len(set(keys)):
        err(f"duplicate facility keys: {[k for k, n in Counter(keys).items() if n > 1]}")
    lic = [f.get("license_no") for f in data if f.get("license_no")]
    if len(lic) != len(set(lic)):
        err(f"duplicate license numbers: {[k for k, n in Counter(lic).items() if n > 1]}")
    newest = None
    for f in data:
        nm = f["name"]
        if f["badge"] not in ("most_serious", "serious", "minor", "clean", "unlicensed"):
            err(f"{nm}: unknown badge {f['badge']}")
        if f.get("license_no") and not re.fullmatch(r"ALR\d{5}", f["license_no"]):
            err(f"{nm}: license number looks wrong: {f['license_no']}")
        rep = f.get("reports") or []
        if f["report_count"] != len(rep):
            err(f"{nm}: report_count {f['report_count']} but {len(rep)} reports listed")
        ds = [r["date"] for r in rep if parse(r.get("date"))]
        if len(ds) != len(rep):
            err(f"{nm}: a report has a missing or bad date")
        if ds and f.get("most_recent_date") != max(ds):
            err(f"{nm}: most_recent_date {f.get('most_recent_date')} but newest report is {max(ds)}")
        if ds:
            top = parse(max(ds))
            newest = top if newest is None else max(newest, top)
        tiers = {p.get("tier") for r in rep for p in r.get("paragraphs", [])} | {r.get("tier") for r in rep}
        if f["badge"] == "most_serious" and "most_serious" not in tiers:
            err(f"{nm}: badged most_serious but no report is tiered most_serious")
        if f["badge"] == "clean" and tiers & {"serious", "most_serious"}:
            err(f"{nm}: badged clean but has serious findings")
        for r in rep:
            if not str(r.get("link", "")).startswith("http"):
                err(f"{nm}: report {r.get('date')} has no link")
        lat, lng = f.get("lat"), f.get("lng")
        if lat is None or lng is None:
            warn(f"{nm}: no map coordinates")
        elif not (41.12 <= lat <= 42.03 and -71.92 <= lng <= -71.08):
            err(f"{nm}: map coordinates fall outside Rhode Island ({lat}, {lng})")
        if not f.get("address") and f["badge"] != "unlicensed":
            warn(f"{nm}: no address")
    m = re.search(r"Data last updated <strong>([^<]+)</strong>", html)
    if m:
        shown = parse_long_date(m.group(1))
        if shown and newest and newest > shown:
            err(f"page says data updated {m.group(1)} but a report dated {newest} is newer")
        if shown and (TODAY - shown).days > 60:
            warn(f"'Data last updated' line is {(TODAY - shown).days} days old ({m.group(1)})")
    else:
        warn("'Data last updated' line not found")


def parse_long_date(s):
    try:
        return dt.datetime.strptime(s.strip(), "%B %d, %Y").date()
    except Exception:
        return None


# ---------------------------------------------------------------- live check
LIVE = {
    "snf": ("https://snfdashboard.abc-ri.org/", ("https://snfdashboard.abc-ri.org/data.js",)),
    "sales": ("https://snfsales.abc-ri.org/", ()),
    "alf": ("https://alinspections.abc-ri.org/", ()),
}


def check_live(kind):
    page, extras = LIVE[kind]
    for url in (page,) + extras:
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "abc-ri-site-check"})
            with urllib.request.urlopen(req, timeout=30) as r:
                body = r.read()
                if r.status != 200:
                    err(f"{url} returned HTTP {r.status}")
                elif len(body) < 10000:
                    err(f"{url} is suspiciously small ({len(body)} bytes)")
        except Exception as e:
            err(f"could not load {url}: {e}")


MATCH = {"sales": ("index.html", "owners.html", "sales.json")}


def check_match(kind):
    """The published copy of each file should be the same as the copy here."""
    import time
    base = LIVE[kind][0]

    def differing():
        out = []
        for name in MATCH.get(kind, ()):
            try:
                req = urllib.request.Request(base + name + "?t=%d" % time.time(), headers={"User-Agent": "abc-ri-site-check", "Cache-Control": "no-cache"})
                with urllib.request.urlopen(req, timeout=60) as r:
                    live = r.read()
            except Exception as e:
                out.append(f"{name} (could not load: {e})")
                continue
            with open(name, "rb") as f:
                if f.read() != live:
                    out.append(name)
        return out

    diff = differing()
    if diff:
        print("Published copy differs for: " + ", ".join(diff) + ". Looking again in 12 minutes.")
        time.sleep(720)
        diff = differing()
    if diff:
        err("the published site does not match this repository: " + ", ".join(diff))


def detect():
    if os.path.exists("data.js"):
        return "snf"
    if os.path.exists("index.html"):
        h = read("index.html")
        if 'id="facility-data"' in h:
            return "alf"
        if 'id="data"' in h:
            return "sales"
    return None


def main():
    kind = detect()
    if not kind:
        print("ERROR: could not tell which site this is. Run from the repo root.")
        return 1
    {"snf": check_snf, "sales": check_sales, "alf": check_alf}[kind]()
    if "--live" in sys.argv:
        check_live(kind)
    if "--match" in sys.argv:
        check_match(kind)
    print(f"Site: {kind}   Date: {TODAY}")
    for w in warns:
        print("WARNING:", w)
    for e in errors:
        print("ERROR:  ", e)
    print(f"{len(errors)} error(s), {len(warns)} warning(s)")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
