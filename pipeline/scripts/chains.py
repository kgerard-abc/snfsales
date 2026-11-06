"""CMS's Nursing Home Chain Performance Measures: national figures for each chain (all its homes in every state).

fetch() returns a list of rows, one per chain, plus a row whose Chain is "National". The build uses it to show a
buyer's national record. Dataset: https://data.cms.gov/quality-of-care/nursing-home-chain-performance-measures
"""
import json
import urllib.request

API = "https://data.cms.gov/data-api/v1/dataset/c51c7172-8406-4963-b4aa-48b695225033/data"


def fetch():
    rows, off = [], 0
    while True:
        req = urllib.request.Request(f"{API}?size=5000&offset={off}", headers={"User-Agent": "abc-ri-update/1.0"})
        page = json.loads(urllib.request.urlopen(req, timeout=120).read())
        rows += page
        if len(page) < 5000:
            return rows
        off += 5000
