#!/usr/bin/env python3
"""Copy the documents linked from RIDOH's Annual Report of Reviews (Change in Effective Control rows) to the Internet Archive.

Usage:
  archive_cec.py [--xlsx path] [--dry-run] [--limit N] [--years 2025,2026] [--only text] [--max-mb 3000] [--all-cec]

Only nursing home applications are archived unless --all-cec is given. RIDOH's sheet has no facility type column, so a row
counts as a nursing home when its applicant or description matches a facility or applicant in the CEC Log
(pipeline/cec_log_snapshot.json), or uses nursing home wording, and is not clearly home care, imaging, a hospital, or a
surgery center. Rows for an applicant that matched in another year are kept too. Fix any miss in
pipeline/cec_archive_scope.json: {"include": ["text"], "exclude": ["text"]}, matched against the applicant name.

What it does:
  1. Gets RIDOH's sheet (downloads it, or uses --xlsx) and pulls every link out of the Change in Effective Control rows
     on each year tab.
  2. Groups the links into one Internet Archive item per applicant per year, named ri-cec-YEAR-applicant.
  3. Downloads each Drive file and uploads it to its item. Folders, private files, missing files, and anything over
     --max-mb are not uploaded. They are listed in the manifest with a status so they can be handled another way.
  4. Saves pipeline/cec_archive_manifest.json after every file, so a stopped run picks up where it left off.

With --dry-run nothing is downloaded or uploaded. It only reads the sheet and reports what it would do.
Secrets IA_ACCESS and IA_SECRET (environment variables) are needed for real uploads.
"""
import os, re, sys, json, time, argparse, shutil, hashlib, tempfile, datetime as dt, urllib.request, urllib.error
from pathlib import Path
import openpyxl

SHEET_ID = "1fdRP2rh34eluTT13k8EIqjJUSnGTH2U5yGD-KwRFi-M"
ROOT = Path(__file__).resolve().parents[2]
MANIFEST = ROOT / "pipeline" / "cec_archive_manifest.json"
HEADER_TEXT = "change in effective control"

def log(*a): print(*a, flush=True)

def clean(v):
    if v is None: return ""
    if isinstance(v, (dt.datetime, dt.date)): return f"{v:%Y-%m-%d}"
    return re.sub(r"\s+", " ", str(v)).strip()

def slug(s, n=60):
    s = re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")
    return s[:n].strip("-") or "item"

def download_sheet(path):
    req = urllib.request.Request(f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx",
                                 headers={"User-Agent": "abc-ri-update/1.0"})
    data = urllib.request.urlopen(req, timeout=180).read()
    if data[:2] != b"PK":
        raise RuntimeError("RIDOH's list did not download as a spreadsheet.")
    Path(path).write_bytes(data)

def drive_ref(url):
    """Return (kind, id). kind is file, folder, gdoc, or other."""
    if "/folders/" in url or "folderview" in url:
        m = re.search(r"(?:/folders/|[?&]id=)([\w-]{15,})", url)
        return ("folder", m.group(1)) if m else ("other", None)
    m = re.search(r"/file/d/([\w-]{15,})", url) or re.search(r"drive\.google\.com/(?:open|uc)\?(?:.*&)?id=([\w-]{15,})", url)
    if m: return ("file", m.group(1))
    m = re.search(r"docs\.google\.com/(document|spreadsheets|presentation)/d/([\w-]{15,})", url)
    if m: return ("gdoc:" + m.group(1), m.group(2))
    return ("other", None)

def collect(xlsx):
    """Every link in the Change in Effective Control rows of every year tab."""
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    recs, skipped_tabs = [], []
    for sh in wb.sheetnames:
        if not sh.strip().isdigit(): continue
        ws = wb[sh]
        rows = list(ws.iter_rows())
        i = next((k for k, r in enumerate(rows) if r and clean(r[0].value).lower() == HEADER_TEXT), None)
        if i is None:
            skipped_tabs.append(sh); continue
        labels = [clean(c.value) for c in rows[i + 1]] if i + 1 < len(rows) else []
        k = i + 2
        while k < len(rows) and rows[k] and clean(rows[k][0].value):
            applicant = clean(rows[k][0].value)
            for j, c in enumerate(rows[k]):
                link = c.hyperlink.target if c.hyperlink is not None and c.hyperlink.target else None
                if not link: continue
                kind, fid = drive_ref(link)
                recs.append({"year": sh, "sheet_row": c.row, "applicant": applicant,
                             "desc": clean(rows[k][1].value) if len(rows[k]) > 1 else "",
                             "label": (labels[j] if j < len(labels) and labels[j] else f"Column {j + 1}"),
                             "link_text": clean(c.value), "url": link, "kind": kind, "id": fid})
            k += 1
    return recs, skipped_tabs

KEYWORDS = re.compile(r"nursing home|skilled nursing|nursing (and|&) rehab|rehabilitation (and|&) (health ?care|care)|"
                      r"rehab(ilitation)? (center|and)|nursing center|elderwood|\bopco\b|\bpropco\b", re.I)
NOT_NURSING_HOME = re.compile(r"home (nursing|health)|home ?care|visiting|staffing|imaging|\bpet\b|hospital|ambulatory|fasc|"
                              r"endoscopy|urolog|infusion|surgery center|brown (health|medicine|university)|lifespan|"
                              r"sleep|orthop|dialysis|pharmacy|hospice", re.I)

def norm(s):
    s = s.lower(); s = re.sub(r"\(.*?\)", " ", s); s = s.replace("&", " and ")
    s = re.sub(r"\b(llc|inc|l\.l\.c|corp|corporation|ltd|dba|d/b/a|the)\b", " ", s)
    return re.sub(r"[^a-z0-9]+", " ", s).strip()

def log_terms():
    p = ROOT / "pipeline" / "cec_log_snapshot.json"
    terms = set()
    if not p.exists(): return terms
    for v in json.loads(p.read_text()).values():
        for t in (v[1], v[4]):
            for part in re.split(r"\s*(?:\u2192|;| / | d/b/a )\s*", t):
                n = norm(part)
                if len(n) >= 8: terms.add(n)
    return terms

def scope_overrides():
    p = ROOT / "pipeline" / "cec_archive_scope.json"
    d = json.loads(p.read_text()) if p.exists() else {}
    return [x.lower() for x in d.get("include", [])], [x.lower() for x in d.get("exclude", [])]

def nursing_home_filter(recs):
    """Split records into kept and dropped. Returns (kept, dropped, reasons) where reasons maps year|applicant to a reason."""
    terms = log_terms(); inc, exc = scope_overrides(); why = {}
    for r in recs:
        key = f"{r['year']}|{r['applicant']}"
        if key in why: continue
        a = r["applicant"].lower(); text = r["applicant"] + " " + r["desc"]
        if any(x in a for x in exc): why[key] = (False, "excluded by scope file")
        elif any(x in a for x in inc): why[key] = (True, "included by scope file")
        elif any(t in norm(text) for t in terms): why[key] = (True, "matches CEC Log")
        elif KEYWORDS.search(text) and not NOT_NURSING_HOME.search(text): why[key] = (True, "nursing home wording")
        else: why[key] = (False, "not a nursing home")
    # An applicant that matched in one year is kept in every year (some rows only say "Change in Effective Control").
    good = {norm(k.split("|", 1)[1]) for k, (ok, _) in why.items() if ok}
    bad = {norm(k.split("|", 1)[1]) for k, (ok, w) in why.items() if not ok and w == "excluded by scope file"}
    for k, (ok, w) in list(why.items()):
        n = norm(k.split("|", 1)[1])
        if not ok and w == "not a nursing home" and n in good and n not in bad: why[k] = (True, "same applicant matched in another year")
    kept = [r for r in recs if why[f"{r['year']}|{r['applicant']}"][0]]
    dropped = [r for r in recs if not why[f"{r['year']}|{r['applicant']}"][0]]
    return kept, dropped, why

def report_scope(why):
    keep = sorted(k for k, (ok, _) in why.items() if ok); drop = sorted(k for k, (ok, _) in why.items() if not ok)
    lines = [f"Nursing homes kept ({len(keep)} applications):"] + [f"  {k}  [{why[k][1]}]" for k in keep]
    lines += [f"Left out as not nursing homes ({len(drop)} applications):"] + [f"  {k}" for k in drop]
    log("\n".join(lines))
    if os.environ.get("GITHUB_STEP_SUMMARY"):
        with open(os.environ["GITHUB_STEP_SUMMARY"], "a") as f:
            f.write(f"### Scope check\n\n**Kept ({len(keep)}):**\n\n" + "\n".join(f"- {k}" for k in keep) +
                    f"\n\n**Left out ({len(drop)}):**\n\n" + "\n".join(f"- {k}" for k in drop) + "\n\n")

def load_manifest():
    if MANIFEST.exists(): return json.loads(MANIFEST.read_text())
    return {"links": {}, "items": {}}

def save_manifest(m):
    m["updated"] = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%MZ")
    MANIFEST.parent.mkdir(parents=True, exist_ok=True)
    MANIFEST.write_text(json.dumps(m, indent=1, sort_keys=True, ensure_ascii=False))

def item_ids(recs, manifest):
    """One stable item name per (year, applicant). Names already in the manifest are kept."""
    ids = dict(manifest.get("items", {}))
    taken = set(ids.values())
    for r in recs:
        key = f"{r['year']}|{r['applicant']}"
        if key in ids: continue
        base = f"ri-cec-{r['year']}-{slug(r['applicant'], 70)}"
        name, n = base, 2
        while name in taken: name = f"{base}-{n}"; n += 1
        ids[key] = name; taken.add(name)
    return ids

def safe_name(name):
    name = re.sub(r"[^\w.\-() ]", "_", name).strip() or "file"
    if len(name) <= 150: return name
    stem, dot, ext = name.rpartition(".")
    return (stem[:145] + "." + ext) if dot and len(ext) <= 5 else name[:150]

def pdf_links(path):
    """Web links in a PDF: clickable links and web addresses printed in the text."""
    urls = set()
    try:
        from pypdf import PdfReader
        for page in PdfReader(str(path)).pages:
            for annot in (page.get("/Annots") or []):
                o = annot.get_object(); act = o.get("/A")
                act = act.get_object() if act is not None else {}
                if act.get("/URI"): urls.add(str(act["/URI"]).strip())
            urls.update(u.rstrip(".,;)") for u in re.findall(r"https?://[^\s<>\"]+", page.extract_text() or ""))
    except Exception as e:
        log(f"   could not read links in {Path(path).name}: {str(e)[:120]}")
    return sorted(urls)

def fetch(rec, workdir):
    """Download one record. Returns (list of paths, status, note)."""
    kind, fid = rec["kind"], rec["id"]
    out = Path(workdir) / "dl"; shutil.rmtree(out, ignore_errors=True); out.mkdir(parents=True)
    if kind == "other": return [], "other_link", "Not a Drive link"
    try:
        if kind == "folder":
            import gdown
            got = gdown.download_folder(id=fid, output=str(out) + "/", quiet=True, remaining_ok=True)
            files = [Path(g.path) for g in (got or []) if Path(g.path).exists()]
            if not files: return [], "failed", "folder came back empty or could not be read"
            note = "Drive lists at most 50 files per folder and this one has 50, so it may be incomplete" if len(files) >= 50 else ""
            return files, "ok", note
        if kind == "file":
            import gdown
            got = gdown.download(id=fid, output=str(out) + "/", quiet=True)
            if not got: return [], "failed", "gdown returned nothing"
            return [Path(got)], "ok", ""
        if kind.startswith("gdoc:"):
            what = kind.split(":", 1)[1]
            fmt = "xlsx" if what == "spreadsheets" else "pdf"
            req = urllib.request.Request(f"https://docs.google.com/{what}/d/{fid}/export?format={fmt}", headers={"User-Agent": "abc-ri-update/1.0"})
            data = urllib.request.urlopen(req, timeout=180).read()
            p = out / f"{fid}.{fmt}"; p.write_bytes(data)
            return [p], "ok", ""
    except urllib.error.HTTPError as e:
        if e.code in (401, 403): return [], "not_public", f"HTTP {e.code}"
        if e.code == 404: return [], "missing", "HTTP 404"
        return [], "failed", f"HTTP {e.code}"
    except Exception as e:
        msg = str(e); low = msg.lower()
        if "permission" in low or "anyone with the link" in low or "sign in" in low: return [], "not_public", msg[:200]
        if "not found" in low or "404" in low: return [], "missing", msg[:200]
        return [], "failed", msg[:200]
    return [], "failed", "unhandled kind"

def with_extension(path):
    """Drive sometimes hands back a file whose name has no ending. Work out the type from the file and add it."""
    path = Path(path)
    if path.suffix: return path
    head = path.read_bytes()[:16] if path.exists() else b""
    ext = None
    if head.startswith(b"%PDF"): ext = ".pdf"
    elif head[:4] == b"PK\x03\x04":
        import zipfile
        try: names = zipfile.ZipFile(path).namelist()
        except Exception: names = []
        ext = (".docx" if any(n.startswith("word/") for n in names) else ".xlsx" if any(n.startswith("xl/") for n in names)
               else ".pptx" if any(n.startswith("ppt/") for n in names) else ".zip")
    elif head[4:8] == b"ftyp": ext = ".mp4"
    elif head[:3] == b"\xff\xd8\xff": ext = ".jpg"
    elif head[:8] == b"\x89PNG\r\n\x1a\n": ext = ".png"
    elif head[:4] == b"\xd0\xcf\x11\xe0": ext = ".doc"
    if not ext: return path
    new = path.with_name(path.name + ext); path.rename(new); return new

def sha256(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""): h.update(chunk)
    return h.hexdigest()

def upload(identifier, path, filename, md, access, secret):
    from internetarchive import get_item
    item = get_item(identifier)
    target = Path(path).with_name(filename)
    if target != Path(path): Path(path).rename(target)
    resp = item.upload(files=[str(target)], metadata=(None if item.exists else md), access_key=access, secret_key=secret,
                       retries=8, retries_sleep=30, verbose=False, validate_identifier=True)
    return all(r.status_code == 200 for r in resp)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx"); ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--limit", type=int, default=0); ap.add_argument("--years", default="")
    ap.add_argument("--only", default=""); ap.add_argument("--max-mb", type=int, default=3000)
    ap.add_argument("--all-cec", action="store_true")
    a = ap.parse_args()

    xlsx = a.xlsx
    if not xlsx:
        xlsx = "ridoh.xlsx"; log("Downloading RIDOH's list of reviews..."); download_sheet(xlsx)
    recs, skipped = collect(xlsx)
    if a.years: recs = [r for r in recs if r["year"] in {y.strip() for y in a.years.split(",")}]
    if a.only: recs = [r for r in recs if a.only.lower() in r["applicant"].lower()]
    log(f"Found {len(recs)} links in Change in Effective Control rows of all kinds. Tabs with no such section: {', '.join(skipped) or 'none'}.")
    if not a.all_cec:
        recs, dropped, why = nursing_home_filter(recs)
        log(f"Keeping {len(recs)} links from nursing home applications. Leaving out {len(dropped)} links from other kinds of applications.")
        report_scope(why)

    # One entry per distinct URL. The first row that uses a link owns it.
    seen, uniq = set(), []
    for r in recs:
        if r["url"] in seen: continue
        seen.add(r["url"]); uniq.append(r)
    kinds, labels = {}, {}
    for r in uniq:
        kinds[r["kind"]] = kinds.get(r["kind"], 0) + 1
        lab = re.sub(r"[*]+", "", r["label"]).strip(); labels[lab] = labels.get(lab, 0) + 1
    log(f"{len(uniq)} different links: " + ", ".join(f"{k} {v}" for k, v in sorted(kinds.items())))
    log("Links by column: " + "; ".join(f"{k} {v}" for k, v in sorted(labels.items(), key=lambda x: -x[1])))

    manifest = load_manifest()
    manifest["items"] = item_ids(uniq, manifest)
    def done(r):
        e = manifest["links"].get(r["url"], {})
        if e.get("status") != "archived": return False
        return e.get("kind") == "folder" or Path(e.get("file", "")).suffix != ""
    queue = [r for r in uniq if not done(r)]
    log(f"{len(uniq) - len(queue)} already archived. {len(queue)} to do. Links found inside HSC reports (public comments) are added as the run goes.")

    if a.dry_run:
        for r in queue[: a.limit or 25]:
            log(f"  would do: {manifest['items'][r['year'] + '|' + r['applicant']]} | {r['label']} | {r['kind']} | {r['url']}")
        log("Dry run: nothing downloaded, uploaded, or saved. Links inside HSC reports can only be found when the reports are downloaded.")
        return

    access, secret = os.environ.get("IA_ACCESS"), os.environ.get("IA_SECRET")
    if not access or not secret:
        sys.exit("IA_ACCESS and IA_SECRET are not set. Add them as repository secrets.")

    workdir = tempfile.mkdtemp(); dl = Path(workdir) / "dl"; count = 0; i = 0
    while i < len(queue):
        r = queue[i]; i += 1
        if a.limit and count >= a.limit: break
        key = f"{r['year']}|{r['applicant']}"; ident = manifest["items"][key]
        entry = {"year": r["year"], "applicant": r["applicant"], "label": r["label"], "link_text": r["link_text"],
                 "sheet_row": r["sheet_row"], "kind": r["kind"], "item": ident, "checked": dt.date.today().isoformat()}
        if r.get("linked_from"): entry["linked_from"] = r["linked_from"]
        prior = manifest["links"].get(r["url"], {}) or {}
        old_file = prior.get("file") if prior.get("status") == "archived" else None
        paths, status, note = fetch(r, workdir)
        stored = []
        if status == "ok":
            paths = [with_extension(p) for p in paths]
            # Public comments: the HSC report links to a Drive folder of them. Add any Drive links found in the report.
            if r["label"].lower().startswith("hsc recommendation") and r["kind"] != "folder":
                found = []
                for p in paths:
                    if p.read_bytes()[:4] != b"%PDF": continue
                    for u in pdf_links(p):
                        k2, id2 = drive_ref(u)
                        if k2 == "other" or u in seen: continue
                        seen.add(u); found.append(u)
                        queue.append({**{x: r[x] for x in ("year", "sheet_row", "applicant")}, "label": "Public comments (linked from HSC report)",
                                      "link_text": "", "url": u, "kind": k2, "id": id2, "desc": r.get("desc", ""), "linked_from": r["url"]})
                if found:
                    entry["links_found_in_report"] = found; log(f"   found {len(found)} Drive link(s) inside this report")
            used = set()
            for e in manifest["links"].values():
                if e.get("item") == ident and e is not prior:
                    used.add(e.get("file")); used.update(f.get("file") for f in e.get("files", []))
            md = {"mediatype": "texts", "title": f"{r['applicant']} ({r['year']}): Change in Effective Control materials",
                  "creator": "Rhode Island Department of Health, Office of Health Systems Development (public filings)",
                  "date": r["year"], "subject": ["Rhode Island", "nursing home", "change in effective control", "certificate of need"],
                  "description": ("Public documents from RIDOH's Annual Report of Reviews for this Change in Effective Control "
                                  "application, copied by Advocates for Better Care in Rhode Island (ABC-RI) so they stay available "
                                  "if RIDOH's links stop working.")}
            for p in paths:
                rel = p.relative_to(dl) if r["kind"] == "folder" else Path(p.name)
                fname = safe_name(" - ".join(rel.parts))
                if fname in used: fname = safe_name(f"{(r['id'] or 'x')[:6]}-{fname}")
                used.add(fname)
                size = p.stat().st_size
                if size / 1e6 > a.max_mb:
                    stored.append({"file": fname, "status": "too_large", "bytes": size}); continue
                digest = sha256(p); ok, err = False, ""
                try: ok = upload(ident, p, fname, md, access, secret)
                except Exception as e: err = str(e)[:200]
                stored.append({"file": fname, "status": "archived" if ok else "failed", "bytes": size, "sha256": digest,
                               "archive_url": f"https://archive.org/download/{ident}/{fname.replace(' ', '%20')}",
                               **({"error": err} if err else {})})
            good = [f for f in stored if f["status"] == "archived"]
            if r["kind"] == "folder":
                entry.update({"files": stored, "archive_url": f"https://archive.org/details/{ident}", "archived_on": dt.date.today().isoformat()})
                status = "archived" if stored and len(good) == len(stored) else ("partial" if good else "failed")
                if note: entry["note_extra"] = note
            elif stored and good:
                f = good[0]
                entry.update({k: f[k] for k in ("file", "bytes", "sha256", "archive_url")}); entry["archived_on"] = dt.date.today().isoformat()
                status = "archived"
                if old_file and old_file != f["file"]:
                    try:
                        from internetarchive import get_item
                        get_item(ident).get_file(old_file).delete(access_key=access, secret_key=secret, cascade_delete=True)
                    except Exception as e:
                        log(f"   could not remove the old copy {old_file}: {str(e)[:120]}")
            elif stored and stored[0]["status"] == "too_large":
                status, note = "too_large", f"{stored[0]['bytes'] / 1e6:.0f} MB, over the {a.max_mb} MB limit"
            else:
                status = "failed"; note = note or "upload did not return success"
        entry.update({"status": status, "note": note})
        manifest["links"][r["url"]] = entry
        save_manifest(manifest); count += 1
        extra = f" ({len(stored)} files)" if r["kind"] == "folder" and stored else ""
        log(f"[{count}] {status:10} {ident} | {r['label']}{extra} {('| ' + note) if note else ''}")
        shutil.rmtree(dl, ignore_errors=True)
        time.sleep(1)

    counts = {}
    for e in manifest["links"].values(): counts[e["status"]] = counts.get(e["status"], 0) + 1
    log("Totals in manifest: " + ", ".join(f"{k} {v}" for k, v in sorted(counts.items())))
    if os.environ.get("GITHUB_STEP_SUMMARY"):
        with open(os.environ["GITHUB_STEP_SUMMARY"], "a") as f:
            f.write("### CEC archive totals\n\n" + "\n".join(f"- {k}: {v}" for k, v in sorted(counts.items())) + "\n")

main()
