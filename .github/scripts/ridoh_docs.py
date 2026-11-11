#!/usr/bin/env python3
"""Collect the documents RIDOH links from each Change in Effective Control row, and tie them to CEC IDs.

Usage: ridoh_docs.py [--ridoh ridoh.xlsx] [--internal internal.xlsx]
        (a missing file is downloaded; ridoh_watch.py and cec_log_watch.py leave both in the working folder)

RIDOH's "Annual Report of Reviews" has one tab per year. In each Change in Effective Control row, these columns carry links:
  A  applicant name, linked to the application          E  review start, linked to the initiation of review notice
  G  Council recommendation, linked to the Council report H  state agency decision, linked to the decision letter
  I, J  the two Health Services Council meetings, linked to the meeting video (older years: audio)
Every year tab is read, not just the last two, because an application stays listed on later tabs as it moves along.

Each RIDOH row is tied to CEC IDs through the CEC Log tab of the INTERNAL sheet: the log's "Report tabs listed in" column
says which year tabs to look on, and the row is matched by RIDOH's description (the log keeps RIDOH's wording) or, when the
description is only "Change in Effective Control", by the applicant's name. A row that can't be matched for certain is left
out and listed under "unmatched" rather than guessed. Rows for other kinds of applications (hospitals, home care, and so on)
never match a CEC Log row, so only nursing home documents are kept.

Saves pipeline/ridoh_cec_docs.json, one entry per document (keyed by Drive id): type, date, link, RIDOH's link text, the
RIDOH rows that link it, the CEC IDs it belongs to, and the date it was first seen. Also lists CEC Log cells (columns O to U)
that are blank but that RIDOH's links can fill, and cells whose link differs from RIDOH's current one.

Writes GitHub Actions outputs: docs_changed, docs_first, new_count, new_ids (Drive ids to archive), and new_docs.md (the
notice text) when there are new documents.
"""
import os, re, sys, json, argparse, difflib, datetime as dt, urllib.request
import openpyxl

RIDOH_SHEET = "1fdRP2rh34eluTT13k8EIqjJUSnGTH2U5yGD-KwRFi-M"
INTERNAL_SHEET = os.environ.get("MASTER_SHEET_ID", "1PAgtaNxnLerzBGVojlxgp-yB7N89msak6giJh4Sg6Wk")
ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..")
SNAP = os.path.join(ROOT, "pipeline", "ridoh_cec_docs.json")
ARCHIVE = os.path.join(ROOT, "pipeline", "cec_archive_manifest.json")
NOTICE = "new_docs.md"
SITE = "https://snfsales.abc-ri.org/"

# RIDOH column (0-based) -> document type
DOC_COLS = {0: "application", 4: "notice", 6: "report", 7: "decision", 8: "video", 9: "video"}
TYPE_LABEL = {"application": "Application", "notice": "Notice that review began", "report": "Council report",
              "decision": "RIDOH decision letter", "video": "Council meeting video", "audio": "Council meeting audio"}
# CEC Log columns RIDOH's links can fill, by header (exact text)
LOG_COLS = [("application", "Application"), ("application_earlier", "Application (earlier filing)"),
            ("notice", "Initiation of review notice"), ("report", "HSC report"), ("decision", "Decision letter"),
            ("video", "HSC meeting video"), ("video2", "HSC meeting video (2nd)")]
FILLABLE_TEXT = re.compile(r"^\s*(tbd|n/?a|pending|none)?\s*$", re.I)   # a cell holding only this can be filled

def out(k, v):
    print(f"{k}={v}")
    if os.environ.get("GITHUB_OUTPUT"):
        with open(os.environ["GITHUB_OUTPUT"], "a") as f: f.write(f"{k}={v}\n")

def clean(v):
    if v is None: return ""
    if isinstance(v, (dt.datetime, dt.date)): return f"{v:%Y-%m-%d}"
    if isinstance(v, float) and v == int(v): v = int(v)
    return re.sub(r"\s+", " ", str(v)).strip()

def norm(t):
    return " ".join(re.sub(r"[^a-z0-9 ]+", " ", str(t or "").lower()).split())

def drive_ref(url):
    """(kind, id): kind is file, folder, gdoc:document|spreadsheets|presentation, or other."""
    if "/folders/" in url or "folderview" in url:
        m = re.search(r"(?:/folders/|[?&]id=)([\w-]{15,})", url)
        return ("folder", m.group(1)) if m else ("other", None)
    m = re.search(r"/file/d/([\w-]{15,})", url) or re.search(r"drive\.google\.com/(?:open|uc)\?(?:.*&)?id=([\w-]{15,})", url)
    if m: return ("file", m.group(1))
    m = re.search(r"docs\.google\.com/(document|spreadsheets|presentation)/d/([\w-]{15,})", url)
    if m: return ("gdoc:" + m.group(1), m.group(2))
    return ("other", None)

def canon(url):
    """One standard address per file or folder. Older Drive links carry a resourcekey that is needed to open them, so it is kept."""
    kind, fid = drive_ref(url)
    rk = re.search(r"[?&]resourcekey=([\w-]+)", url)
    rk = f"?resourcekey={rk.group(1)}" if rk else ""
    if kind == "file": return f"https://drive.google.com/file/d/{fid}/view{rk}"
    if kind == "folder": return f"https://drive.google.com/drive/folders/{fid}{rk}"
    if kind == "folder": return f"https://drive.google.com/drive/folders/{fid}"
    if kind.startswith("gdoc:"): return f"https://docs.google.com/{kind.split(':', 1)[1]}/d/{fid}"
    return url.strip()

def doc_key(url):
    kind, fid = drive_ref(url)
    return fid or canon(url)

MONTHS = {m: i + 1 for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"])}
DATE_RE = re.compile(r"(?P<mon>Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+(?P<d1>\d{1,2})(?:st|nd|rd|th)?,?\s+(?P<y1>\d{4})"
                     r"|(?P<m2>\d{1,2})/(?P<d2>\d{1,2})/(?P<y2>\d{4}|\d{2})\b", re.I)

def dates_in(text):
    """Every date written in the text, in order, as ISO strings."""
    found = []
    for m in DATE_RE.finditer(text or ""):
        try:
            if m.group("mon"): d = dt.date(int(m.group("y1")), MONTHS[m.group("mon")[:3].lower()], int(m.group("d1")))
            else:
                y = int(m.group("y2")); y += 2000 if y < 100 else 0
                d = dt.date(y, int(m.group("m2")), int(m.group("d2")))
            found.append(d.isoformat())
        except ValueError:
            pass
    return found

def doc_date(dtype, text, filed):
    """The date that goes with a document. A Council report cell lists the meeting dates and the adoption date last;
    the other cells give their date first. An application has no date in its link, so its first filing date is used."""
    if dtype == "application":
        if isinstance(filed, (dt.datetime, dt.date)): return f"{filed:%Y-%m-%d}"
        ds = dates_in(clean(filed)); return ds[0] if ds else ""
    ds = dates_in(text)
    if not ds: return ""
    return ds[-1] if dtype == "report" else ds[0]

def download(sheet_id, path, what):
    req = urllib.request.Request(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx", headers={"User-Agent": "abc-ri-update/1.0"})
    data = urllib.request.urlopen(req, timeout=180).read()
    if data[:2] != b"PK": raise RuntimeError(f"{what} did not download as a spreadsheet")
    open(path, "wb").write(data)

def read_ridoh(path):
    """Every Change in Effective Control row on every year tab, with its links in columns A, E, G, H, I, J."""
    wb = openpyxl.load_workbook(path, data_only=True)   # not read-only: read-only mode does not load cell links
    rows = []
    for sh in wb.sheetnames:
        if not sh.strip().isdigit(): continue
        data = list(wb[sh].iter_rows())
        i = next((k for k, r in enumerate(data) if r and clean(r[0].value).lower() == "change in effective control"), None)
        if i is None: continue
        k = i + 2
        while k < len(data) and data[k] and clean(data[k][0].value):
            r = list(data[k]) + [None] * 12
            links = []
            for j, dtype in DOC_COLS.items():
                c = r[j]
                if c is None or c.hyperlink is None or not c.hyperlink.target: continue
                url = c.hyperlink.target.strip(); text = clean(c.value)
                if dtype == "video" and re.search(r"\baudio\b", text, re.I): dtype = "audio"
                links.append(dict(type=dtype, col="ABCDEFGHIJK"[j], url=canon(url), key=doc_key(url),
                                  drive_kind=drive_ref(url)[0], text=text, date=doc_date(dtype, text, r[2].value if r[2] is not None else None)))
            rows.append(dict(tab=sh.strip(), row=data[k][0].row, applicant=clean(r[0].value), desc=clean(r[1].value) if r[1] is not None else "", links=links))
            k += 1
    return rows

def read_log(path):
    """CEC Log rows: id, facility, applicant, description, the RIDOH tabs it is listed on, and its links in columns O to U."""
    wb = openpyxl.load_workbook(path, data_only=True)
    if "CEC Log" not in wb.sheetnames: return []
    ws = wb["CEC Log"]
    hdr = {clean(c.value): c.column for c in ws[1] if clean(c.value)}
    def col(*starts):
        for name, c in hdr.items():
            if any(name.startswith(s) for s in starts): return c
    cid, fac, app, desc, tabs, year = col("CEC ID"), col("Facility"), col("Applicant"), col("Description"), col("Report tabs listed in"), col("Year")
    lcols = {k: hdr.get(h) for k, h in LOG_COLS}
    out_ = []
    for r in range(2, ws.max_row + 1):
        v = clean(ws.cell(row=r, column=cid).value)
        if not v: continue
        links = {}
        for k, c in lcols.items():
            if not c: continue
            cell = ws.cell(row=r, column=c)
            links[k] = dict(col=openpyxl.utils.get_column_letter(c), header=dict(LOG_COLS)[k], text=clean(cell.value),
                            url=(cell.hyperlink.target.strip() if cell.hyperlink is not None and cell.hyperlink.target else ""))
        t = sorted(set(re.findall(r"\b(?:19|20)\d\d\b", clean(ws.cell(row=r, column=tabs).value) if tabs else "")))
        if not t and year: t = re.findall(r"\b(?:19|20)\d\d\b", clean(ws.cell(row=r, column=year).value))
        out_.append(dict(id=v, sheet_row=r, fac=clean(ws.cell(row=r, column=fac).value) if fac else "",
                         applicant=clean(ws.cell(row=r, column=app).value) if app else "",
                         desc=clean(ws.cell(row=r, column=desc).value) if desc else "", tabs=t, links=links))
    return out_

def generic(desc):
    """True when a description says nothing that tells applications apart ("Change in Effective Control")."""
    n = norm(desc)
    return len(n) < 30 or n in ("change in effective control", "cec", "change in effective control application")

def match(c, rows):
    """The RIDOH row on one tab that goes with a CEC Log row, or (None, reason).
    First by link: when the log's application link (column O or P) is linked from exactly one row on the tab, that is the row.
    Otherwise by RIDOH's description (the log keeps RIDOH's wording) or the applicant's name."""
    mine = {doc_key(v["url"]) for k, v in c["links"].items() if k in ("application", "application_earlier") and v["url"]}
    if mine:
        hits = [e for e in rows if any(l["type"] == "application" and l["key"] in mine for l in e["links"])]
        if len(hits) == 1: return hits[0], ""
    a1 = norm(c["applicant"].split("→")[0])
    bare = lambda t: " ".join(w for w in t.split() if w not in ("llc", "inc", "l", "c"))
    llc = bare(norm(re.split(r"\s+d/?b/?a\s+", c["applicant"].split("→")[0], flags=re.I)[0]))
    scored = []
    for e in rows:
        ds = 0.0 if generic(c["desc"]) or generic(e["desc"]) else difflib.SequenceMatcher(None, norm(c["desc"]), norm(e["desc"])).ratio()
        a2 = norm(e["applicant"])
        ap = 0.9 if a1 and a2 and len(min(a1, a2, key=len)) >= 8 and (a2.startswith(a1) or a1.startswith(a2)) else 0.0
        # RIDOH sometimes puts the buyer group first ("Kadima - Bayberry Rehabilitation & Nursing LLC"), so also accept its
        # name containing the log's LLC name (the part before "d/b/a").
        if not ap and len(llc) >= 15 and llc in bare(a2): ap = 0.85
        tie = difflib.SequenceMatcher(None, a1, a2).ratio()   # only used to break ties
        scored.append((max(ds, ap), tie, e))
    if not scored: return None, "no Change in Effective Control rows on that tab"
    scored.sort(key=lambda x: (x[0], x[1]), reverse=True)
    best = scored[0]
    if best[0] < 0.75: return None, "no row matched its description or applicant"
    if len(scored) > 1 and best[0] - scored[1][0] < 0.02 and abs(best[1] - scored[1][1]) < 0.02:
        return None, "two rows matched equally well: " + best[2]["applicant"][:40] + " / " + scored[1][2]["applicant"][:40]
    return best[2], ""

def build(ridoh_rows, log_rows):
    by_tab = {}
    for e in ridoh_rows: by_tab.setdefault(e["tab"], []).append(e)
    docs, unmatched, per_cec = {}, [], {}
    for c in log_rows:
        for tab in c["tabs"]:
            if tab not in by_tab:
                continue
            e, why = match(c, by_tab[tab])
            if e is None:
                unmatched.append(dict(cec=c["id"], tab=tab, reason=why)); continue
            for l in e["links"]:
                d = docs.setdefault(l["key"], dict(type=l["type"], label=TYPE_LABEL[l["type"]], date=l["date"], text=l["text"], url=l["url"],
                                                   drive_kind=l["drive_kind"], rows=[], cec=[]))
                ref = dict(tab=e["tab"], row=e["row"], col=l["col"], applicant=e["applicant"])
                if ref not in d["rows"]: d["rows"].append(ref)
                if c["id"] not in d["cec"]: d["cec"].append(c["id"])
                d.setdefault("text_by_cec", {})[c["id"]] = l["text"]   # one file can sit on several rows with different link text
                # The same file linked in two places (a video for two meetings' rows): keep the newest tab's text and date.
                if e["tab"] >= max(x["tab"] for x in d["rows"]): d.update(type=l["type"], label=TYPE_LABEL[l["type"]], date=l["date"] or d["date"], text=l["text"])
                per_cec.setdefault(c["id"], []).append((e["tab"], l))
    for d in docs.values():
        d["cec"].sort(); d["rows"].sort(key=lambda x: (x["tab"], x["row"], x["col"]))
    return docs, unmatched, per_cec

def log_gaps(log_rows, per_cec):
    """CEC Log cells (O to U) that RIDOH's links can fill, and cells whose link differs from RIDOH's current one."""
    gaps, differs = [], []
    for c in log_rows:
        got = per_cec.get(c["id"])
        if not got or not c["links"]: continue
        have = {doc_key(v["url"]) for v in c["links"].values() if v["url"]}
        newest = lambda t: sorted([(tab, l) for tab, l in got if l["type"] == t], key=lambda x: (x[0], x[1]["date"]), reverse=True)
        want = {}
        apps = []
        for tab, l in newest("application"):
            if l["key"] not in [a["key"] for a in apps]: apps.append(l)
        if apps: want["application"] = apps[0]
        if len(apps) > 1: want["application_earlier"] = apps[1]
        for t in ("notice", "report", "decision"):
            n = newest(t)
            if n: want[t] = n[0][1]
        vids = newest("video") or newest("audio")
        if vids:
            top = vids[0][0]
            same = sorted({l["key"]: l for tab, l in vids if tab == top}.values(), key=lambda l: l["date"])
            if same: want["video"] = same[0]
            if len(same) > 1: want["video2"] = same[1]
        for k, l in want.items():
            cell = c["links"].get(k)
            if not cell: continue
            if cell["url"]:
                if doc_key(cell["url"]) != l["key"] and l["key"] not in have:
                    differs.append(dict(cec=c["id"], sheet_row=c["sheet_row"], column=cell["col"], header=cell["header"],
                                        sheet_url=cell["url"], ridoh_url=l["url"], ridoh_text=l["text"]))
            elif FILLABLE_TEXT.match(cell["text"]) and l["key"] not in have:
                gaps.append(dict(cec=c["id"], sheet_row=c["sheet_row"], column=cell["col"], header=cell["header"],
                                 url=l["url"], text=l["text"] or TYPE_LABEL[l["type"]]))
    return gaps, differs

def nice_date(iso):
    try: d = dt.date.fromisoformat(iso); return f"{d:%B} {d.day}, {d.year}"
    except Exception: return "date not given"

def notice(new, docs, log_rows, gaps):
    fac = {c["id"]: c["fac"] for c in log_rows}
    ids = sorted({x for k in new for x in docs[k]["cec"]})
    lines = ["@kgerard-abc RIDOH posted " + (f"{len(new)} new documents" if len(new) != 1 else "a new document") +
             " for nursing home applications on its Annual Report of Reviews.", ""]
    for k in sorted(new, key=lambda k: (docs[k]["cec"], docs[k]["date"])):
        d = docs[k]
        who = "; ".join(f"{x} {fac.get(x, '')}".strip() for x in d["cec"])
        txt = d["text_by_cec"].get(d["cec"][0]) or d["text"] or "Open it"
        lines.append(f"- **{who}**: {d['label']}, {nice_date(d['date'])}. [{txt}]({d['url']})")
    lines += ["", "What happens next, on its own:",
              "- The sales page rebuilds, and these show under \"CEC documents\" for each sale (" + SITE + ").",
              "- A copy of each is saved to the Internet Archive (the \"Archive CEC documents\" workflow, run for just these files).",
              ""]
    g = [x for x in gaps if x["cec"] in ids]
    if g:
        lines += [f"The CEC Log has {len(g)} blank link cells these can fill. The weekly CEC Log check fills them on Wednesday:", ""]
        lines += [f"- {x['cec']}, column {x['column']} ({x['header']})" for x in g]
        lines.append("")
    lines.append("This notice was opened by the \"Watch RIDOH's list of reviews\" workflow. Close it when you have seen it.")
    return "\n".join(lines), ids

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--ridoh"); ap.add_argument("--internal")
    a = ap.parse_args()
    rp = a.ridoh or "ridoh.xlsx"; ip = a.internal or "internal.xlsx"
    try:
        if not os.path.exists(rp): download(RIDOH_SHEET, rp, "RIDOH's list")
        if not os.path.exists(ip): download(INTERNAL_SHEET, ip, "The INTERNAL sheet")
    except Exception as ex:
        print(f"::warning::Could not get the sheets ({ex}). Will try again next time."); out("docs_changed", "false"); out("docs_first", "false"); out("new_count", 0); out("new_ids", ""); return
    ridoh_rows = read_ridoh(rp); log_rows = read_log(ip)
    print(f"RIDOH: {len(ridoh_rows)} Change in Effective Control rows on {len({r['tab'] for r in ridoh_rows})} year tabs, "
          f"{sum(len(r['links']) for r in ridoh_rows)} links in columns A, E, G, H, I, J. CEC Log: {len(log_rows)} rows.")
    docs, unmatched, per_cec = build(ridoh_rows, log_rows)
    gaps, differs = log_gaps(log_rows, per_cec)
    counts = {}
    for d in docs.values(): counts[d["type"]] = counts.get(d["type"], 0) + 1
    print(f"Nursing home documents: {len(docs)} (" + ", ".join(f"{k} {v}" for k, v in sorted(counts.items())) + f") for {len(per_cec)} CEC IDs.")
    for u in unmatched: print(f"  Not matched: {u['cec']} on the {u['tab']} tab ({u['reason']})")
    print(f"CEC Log: {len(gaps)} blank link cells RIDOH can fill, {len(differs)} cells whose link differs from RIDOH's.")
    for x in differs: print(f"  Differs: {x['cec']} column {x['column']}: sheet {x['sheet_url'][:70]} / RIDOH {x['ridoh_url'][:70]}")

    if len(log_rows) < 100 or len(docs) < 50:
        print(f"::warning::Only {len(log_rows)} CEC Log rows and {len(docs)} documents found; not trusting it. Nothing changed.")
        out("docs_changed", "false"); out("docs_first", "false"); out("new_count", 0); out("new_ids", ""); return
    old = json.load(open(SNAP)) if os.path.exists(SNAP) else None
    if old is not None and len(docs) < 0.7 * len(old.get("docs", {})):
        print(f"::warning::Documents dropped from {len(old['docs'])} to {len(docs)}; not trusting it. Nothing changed.")
        out("docs_changed", "false"); out("docs_first", "false"); out("new_count", 0); out("new_ids", ""); return

    today = dt.date.today().isoformat(); od = (old or {}).get("docs", {})
    new = []
    for k, d in docs.items():
        o = od.get(k)
        d["first_seen"] = (o or {}).get("first_seen", today)
        if o is None or set(d["cec"]) - set(o.get("cec", [])): new.append(k)
    first = old is None
    if first:
        # First run: everything is "new" but nobody needs a notice for years of old documents. Notice only what RIDOH dated
        # in the last 3 weeks, and hand the archive anything it has no record of.
        arch = json.load(open(ARCHIVE)).get("links", {}) if os.path.exists(ARCHIVE) else {}
        known = {doc_key(u) for u in arch}
        cutoff = (dt.date.today() - dt.timedelta(days=21)).isoformat()
        to_archive = [k for k in docs if k not in known and drive_ref(docs[k]["url"])[0] != "other"]
        new = [k for k in docs if docs[k]["date"] >= cutoff]
    else:
        to_archive = [k for k in new if drive_ref(docs[k]["url"])[0] != "other"]
    gone = [k for k in od if k not in docs]
    for k in gone: print(f"  No longer linked by RIDOH: {od[k].get('label')} {od[k].get('text', '')[:60]} ({', '.join(od[k].get('cec', []))})")
    for k in new: print(f"  NEW: {', '.join(docs[k]['cec'])} | {docs[k]['label']} | {docs[k]['date']} | {docs[k]['text'][:70]}")

    snap = dict(updated=dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%MZ"),
                about="Documents RIDOH links from its Annual Report of Reviews, for nursing home Change in Effective Control applications. Made by .github/scripts/ridoh_docs.py.",
                docs=docs, unmatched=unmatched, cec_log_gaps=gaps, cec_log_differs=differs)
    strip = lambda s: {k: v for k, v in (s or {}).items() if k != "updated"}
    changed = first or strip(snap) != strip(old)
    if changed:
        json.dump(snap, open(SNAP, "w"), indent=1, sort_keys=True, ensure_ascii=False)
    print("Documents changed." if changed else "No change in the documents.")
    out("docs_changed", str(changed).lower()); out("docs_first", str(first).lower())
    out("new_count", len(new)); out("new_ids", ",".join(to_archive))
    if new:
        body, ids = notice(new, docs, log_rows, gaps)
        open(NOTICE, "w").write(body)
        out("new_cec", ", ".join(ids[:6]) + (" and more" if len(ids) > 6 else ""))

main()
