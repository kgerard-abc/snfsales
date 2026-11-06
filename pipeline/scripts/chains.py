"""CMS's Nursing Home Chain Performance Measures: national figures for each chain (all its homes in every state).

fetch() returns a list of rows, one per chain, plus a row whose Chain is "National". The build uses it to show a
buyer's national record, and saves a copy as pipeline/cms_chains.csv so anyone can look up the name CMS uses for a
chain (to add it under "cms_chain" in pipeline/group_names.json). If CMS can't be reached, the saved copy is used.
Dataset: https://data.cms.gov/quality-of-care/nursing-home-chain-performance-measures
"""
import csv
import json
import os
import urllib.request

ID = "c51c7172-8406-4963-b4aa-48b695225033"
API = f"https://data.cms.gov/data-api/v1/dataset/{ID}/data"
VIEWER = f"https://data.cms.gov/data-api/v1/dataset/{ID}/data-viewer?size=1"


def _get(url):
    req = urllib.request.Request(url, headers={"User-Agent": "abc-ri-update/1.0"})
    return json.loads(urllib.request.urlopen(req, timeout=120).read())


def fetch():
    rows, off = [], 0
    while True:
        page = _get(f"{API}?size=5000&offset={off}")
        rows += page
        if len(page) < 5000:
            return rows
        off += 5000


def version():
    """The date of this release (YYYY-MM-DD), read from the newest file name CMS lists for the dataset
    (for example Chain_Performance_20260610.csv), or None."""
    import re
    try:
        txt = json.dumps(_get(f"https://data.cms.gov/data-api/v1/dataset-resources/{ID}"))
    except Exception:
        return None
    ds = sorted(set(re.findall(r"_(20[2-9]\d)(0[1-9]|1[0-2])([0-3]\d)\.csv", txt)))
    return f"{ds[-1][0]}-{ds[-1][1]}-{ds[-1][2]}" if ds else None


def save(rows, path, asof=None):
    cols = list(rows[0].keys()) if rows else []
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["as_of"] + cols)
        for r in rows:
            w.writerow([asof or ""] + [r.get(c, "") for c in cols])


def load(path):
    if not os.path.exists(path):
        return None, None
    with open(path, encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    asof = rows[0].pop("as_of", None) if rows else None
    for r in rows[1:]:
        r.pop("as_of", None)
    return rows, asof or None
