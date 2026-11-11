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
  2. Groups the links into one Internet Archive item per application, named ri-cec-FIRSTYEAR-applicant. An application that
     takes years is listed on several year tabs, with the report and decision only on the last one. Rows with the same
     applicant name, in the same position among that applicant's rows, in tabs no more than 3 years apart, are one application.
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

def canon(url):
    """One standard address per Drive file, folder, or Google doc, so .../view and .../view?usp=sharing match."""
    kind, fid = drive_ref(url)
    if kind == "file": return f"https://drive.google.com/file/d/{fid}"
    if kind == "folder": return f"https://drive.google.com/drive/folders/{fid}"
    if kind.startswith("gdoc:"): return f"https://docs.google.com/{kind.split(':', 1)[1]}/d/{fid}"
    return url

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
                             "label": (labels[j] if j < len(labels) and labels[j] else (f"{labels[j - 1]} (more)" if 0 < j <= len(labels) and labels[j - 1] else f"Column {j + 1}")),
                             "link_text": clean(c.value), "url": canon(link), "orig_url": link, "kind": kind, "id": fid})
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

_APPS = None
def load_manifest_for_apps():
    global _APPS
    _APPS = load_manifest(); return _APPS

def load_manifest():
    m = json.loads(MANIFEST.read_text()) if MANIFEST.exists() else {}
    for k in ("links", "apps"): m.setdefault(k, {})
    m.setdefault("cleanup", [])
    for u in list(m["links"]):
        cu = canon(u)
        if cu != u:
            e = m["links"].pop(u)
            if cu not in m["links"]: m["links"][cu] = e
    for e in m["links"].values():
        if e.get("linked_from"): e["linked_from"] = canon(e["linked_from"])
    return m

def save_manifest(m):
    m["updated"] = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%MZ")
    MANIFEST.parent.mkdir(parents=True, exist_ok=True)
    MANIFEST.write_text(json.dumps(m, indent=1, sort_keys=True, ensure_ascii=False))

def assign_applications(recs, manifest):
    """Tie rows on different year tabs together as one application and give each application one item name.
    Sets r['item'], r['first_year'], r['last_year'] on every record."""
    rows = {}
    for r in recs: rows.setdefault((r["year"], r["sheet_row"]), r)
    by_name = {}
    for (yr, row), r in rows.items(): by_name.setdefault(norm(r["applicant"]), []).append((int(yr), row, r))
    apps = manifest.setdefault("apps", {}); taken = set(apps.values())
    for n, lst in by_name.items():
        per_year = {}
        for yr, row, r in sorted(lst, key=lambda x: (x[0], x[1])): per_year.setdefault(yr, []).append(r)
        occ_max = max(len(v) for v in per_year.values())
        for occ in range(occ_max):
            chain = []
            for yr in sorted(per_year):
                if occ < len(per_year[yr]):
                    if chain and yr - chain[-1][0] > 3:
                        flush(chain, n, occ, occ_max > 1, apps, taken); chain = []
                    chain.append((yr, per_year[yr][occ]))
            if chain: flush(chain, n, occ, occ_max > 1, apps, taken)
    for r in recs:
        r["item"] = r["_chain_item"]; r["first_year"] = r["_first"]; r["last_year"] = r["_last"]

def flush(chain, n, occ, several, apps, taken):
    first, last = chain[0][0], chain[-1][0]
    key = f"{n}|{occ}|{first}"
    if key not in apps:
        latest = chain[-1][1]
        text = latest["applicant"] + ((" " + re.sub(r"(?i)change in effective control\W*", "", latest.get("desc", ""))[:50]) if several else "")
        base = f"ri-cec-{first}-{slug(text, 70)}"; name, k = base, 2
        while name in taken: name = f"{base}-{k}"; k += 1
        apps[key] = name; taken.add(name)
    for _, r in chain:
        for x in RECS_BY_ROW[(r["year"], r["sheet_row"])]:
            x["_chain_item"] = apps[key]; x["_first"] = str(first); x["_last"] = str(last)

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
            import gdown, inspect
            # Newer versions of gdown dropped the remaining_ok option, so only pass it when it is accepted.
            extra = {"remaining_ok": True} if "remaining_ok" in inspect.signature(gdown.download_folder).parameters else {}
            got = gdown.download_folder(id=fid, output=str(out) + "/", quiet=True, **extra)
            # Some versions of gdown return file paths, others return objects that have a .path
            paths_ = [Path(getattr(g, "path", g)) for g in (got or [])]
            files = [p for p in paths_ if p.is_file()]
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
        msg = re.sub(r"\s+", " ", str(e)); low = msg.lower()
        if ("cannot retrieve the public link" in low or "many accesses" in low or "failed to retrieve file url" in low
                or "failed to download the following files" in low):
            return [], "blocked", "Drive refused the download. The file is private, or Drive is limiting downloads. It is tried again on the next run."
        if "permission" in low or "anyone with the link" in low or "sign in" in low: return [], "not_public", msg[:200]
        if "not found" in low or "404" in low: return [], "missing", msg[:200]
        return [], "failed", msg[:200]
    return [], "failed", "unhandled kind"

def has_extension(name):
    """True for a real file ending like .pdf or .docx. Dates such as 2025.09.26 are not endings."""
    return re.fullmatch(r"\.[A-Za-z][A-Za-z0-9]{1,4}", Path(name).suffix) is not None

GIVE_UP = 8

def with_extension(path):
    """Drive sometimes hands back a file whose name has no ending. Work out the type from the file and add it."""
    path = Path(path)
    if has_extension(path.name): return path
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
    """Upload one local file under exactly the name given. True only if archive.org accepted it."""
    from internetarchive import get_item
    item = get_item(identifier)
    resp = item.upload(files={filename: str(path)}, metadata=(None if item.exists else md), access_key=access, secret_key=secret,
                       retries=8, retries_sleep=30, verbose=False, validate_identifier=True)
    if resp and all(r.status_code == 200 for r in resp): return True
    if not resp:   # nothing was sent: fine if archive.org already holds this exact file
        item = get_item(identifier)
        return any(f.get("name") == filename and int(f.get("size") or -1) == os.path.getsize(path) for f in item.files)
    raise RuntimeError("archive.org answered " + ", ".join(str(r.status_code) for r in resp))

def child_rec(r, u):
    """A link found inside an HSC report, filed in the same item as the report."""
    k2, id2 = drive_ref(u)
    return {**{x: r[x] for x in ("year", "sheet_row", "applicant", "item", "first_year", "last_year")},
            "label": "Public comments (linked from HSC report)", "link_text": "", "url": canon(u), "orig_url": u,
            "kind": k2, "id": id2, "desc": r.get("desc", ""), "linked_from": r["url"]}

def try_delete(item_name, file_name, access, secret):
    """Remove one file from archive.org. True if it is gone (removed, or was never there)."""
    try:
        from internetarchive import get_item
        it = get_item(item_name)
        if not it.exists: return True
        f = it.get_file(file_name)
        if f is None: return True
        resp = f.delete(access_key=access, secret_key=secret, cascade_delete=True)
        if getattr(resp, "status_code", 200) >= 300: raise RuntimeError(f"HTTP {resp.status_code}")
        log(f"   removed {file_name} from {item_name}"); return True
    except Exception as e:
        log(f"   could not remove {file_name} from {item_name}: {str(e)[:120]}"); return False

def ia_tasks(item, access, secret):
    """The item's task queue on archive.org: running, waiting, and failed tasks, then the most recent finished ones."""
    lines = []
    for params in ("catalog=1&history=0&limit=10", "catalog=0&history=1&limit=6"):
        try:
            req = urllib.request.Request(f"https://archive.org/services/tasks.php?identifier={item}&{params}",
                                         headers={"Authorization": f"LOW {access}:{secret}", "User-Agent": "abc-ri-update/1.0"})
            v = json.load(urllib.request.urlopen(req, timeout=60)).get("value", {}) or {}
            for t in (v.get("catalog") or []) + (v.get("history") or []):
                lines.append("    task: " + ", ".join(f"{k}={str(t.get(k))[:70]}" for k in ("cmd", "submittime", "finished", "wait_admin", "status", "args") if t.get(k) not in (None, "", {}, [])))
        except Exception as ex:
            lines.append(f"    could not read the task queue ({str(ex)[:80]})"); break
    return lines

IA_OWN_FILES = re.compile(r"(_archive\.torrent|_files\.xml|_meta\.xml|_meta\.sqlite|_reviews\.xml|__ia_thumb\.jpg)$")

def verify():
    """Ask archive.org what each item holds and compare it with the manifest. Files the manifest says are archived but
    that are not on archive.org are marked not_on_archive, so the next full run uploads them again."""
    m = load_manifest(); items = {}
    for u, e in m["links"].items():
        if e.get("item"): items.setdefault(e["item"], []).append((u, e))
    log(f"Checking {len(items)} items on archive.org. New uploads can take several minutes to show up, so run this a while after an upload.")
    missing_total = 0; checked = 0
    for item, lst in sorted(items.items()):
        try:
            req = urllib.request.Request(f"https://archive.org/metadata/{item}", headers={"User-Agent": "abc-ri-update/1.0"})
            data = json.load(urllib.request.urlopen(req, timeout=60))
        except Exception as ex:
            log(f"  {item}: could not read it ({str(ex)[:80]})"); continue
        have = {f["name"] for f in data.get("files", [])}
        if not have: log(f"  {item}: not found on archive.org, or has no files")
        expected = set()
        for u, e in lst:
            if e.get("status") not in ("archived", "partial", "not_on_archive"): continue
            names = [e["file"]] if e.get("file") else [f["file"] for f in e.get("files", []) if f.get("status") == "archived"]
            expected.update(names); checked += len(names)
            gone = [n for n in names if n not in have]
            if gone:
                missing_total += len(gone); e["status"] = "not_on_archive"; e["note"] = "not found on archive.org: " + "; ".join(gone)[:150]
                for n in gone: log(f"  MISSING in {item}: {n}")
            elif e.get("status") == "not_on_archive":
                e["status"] = "archived"; e["note"] = ""
        if os.environ.get("IA_ACCESS") and os.environ.get("IA_SECRET") and any(e.get("status") == "not_on_archive" for _, e in lst):
            log(f"  Task queue for {item}:"); [log(l) for l in ia_tasks(item, os.environ["IA_ACCESS"], os.environ["IA_SECRET"])]
        stray = sorted(n for n in have - expected if not IA_OWN_FILES.search(n))
        for n in stray:
            twin = any(x.startswith(n + ".") for x in expected)
            log(f"  on archive.org but not in the record, {item}: {n}" + ("  (a copy with a file ending exists, so this one is queued for removal)" if twin else ""))
            c = {"item": item, "file": n}
            if twin and c not in m["cleanup"]: m["cleanup"].append(c)
    save_manifest(m)
    log(f"Checked {checked} files. Missing from archive.org: {missing_total}. Missing ones will be uploaded again by the next full run.")
    if os.environ.get("GITHUB_STEP_SUMMARY"):
        with open(os.environ["GITHUB_STEP_SUMMARY"], "a") as f: f.write(f"### Verify\n\nChecked {checked} files. Missing from archive.org: {missing_total}.\n")

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx"); ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--limit", type=int, default=0); ap.add_argument("--years", default="")
    ap.add_argument("--only", default=""); ap.add_argument("--max-mb", type=int, default=3000)
    ap.add_argument("--all-cec", action="store_true"); ap.add_argument("--verify", action="store_true")
    a = ap.parse_args()
    if a.verify: verify(); return

    xlsx = a.xlsx
    if not xlsx:
        xlsx = "ridoh.xlsx"; log("Downloading RIDOH's list of reviews..."); download_sheet(xlsx)
    recs, skipped = collect(xlsx)
    global RECS_BY_ROW
    RECS_BY_ROW = {}
    for r in recs: RECS_BY_ROW.setdefault((r["year"], r["sheet_row"]), []).append(r)
    assign_applications(recs, load_manifest_for_apps())   # before the year and name filters, so the grouping is the same on every run
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

    manifest = _APPS
    def done(r):
        e = manifest["links"].get(r["url"], {})
        if e.get("status") != "archived" or e.get("item") != r["item"]: return False
        if r["label"].lower().startswith("hsc recommendation") and not e.get("scanned"): return False
        return e.get("kind") == "folder" or has_extension(e.get("file", ""))
    queue = [r for r in uniq if not done(r)]
    # Links found inside a report on an earlier run that were never archived (for example, the run stopped first).
    for r in uniq:
        e = manifest["links"].get(r["url"], {})
        if r["label"].lower().startswith("hsc recommendation") and e.get("links_found_in_report"):
            for u in e["links_found_in_report"]:
                if canon(u) in seen: continue
                seen.add(canon(u)); c = child_rec(r, u)
                if not done(c): queue.append(c)
    tries = lambda q: int((manifest["links"].get(q["url"]) or {}).get("attempts", 0))
    gave_up = [q for q in queue if tries(q) >= GIVE_UP]
    queue = sorted([q for q in queue if tries(q) < GIVE_UP], key=tries)
    if gave_up: log(f"{len(gave_up)} links failed {GIVE_UP} or more times and are skipped. They are listed in the manifest for a person to check.")
    log(f"{len(uniq) - len([q for q in queue if not q.get('linked_from')])} already archived. {len(queue)} to do. Links found inside HSC reports (public comments) are added as the run goes.")

    if a.dry_run:
        for r in queue[: a.limit or 25]:
            log(f"  would do: {r['item']} | {r['label']} | {r['kind']} | {r['url']}")
        log("Dry run: nothing downloaded, uploaded, or saved. Links inside HSC reports can only be found when the reports are downloaded.")
        return

    access, secret = os.environ.get("IA_ACCESS"), os.environ.get("IA_SECRET")
    if not access or not secret:
        sys.exit("IA_ACCESS and IA_SECRET are not set. Add them as repository secrets.")

    # Files queued for removal from archive.org (for example copies filed without a file ending).
    left = [c for c in manifest.get("cleanup", []) if not try_delete(c["item"], c["file"], access, secret)]
    manifest["cleanup"] = left; save_manifest(manifest)

    workdir = tempfile.mkdtemp(); dl = Path(workdir) / "dl"; count = 0
    big = lambda x: re.search(r"meeting (audio|video)|recording", x["label"], re.I) is not None
    blocked_streak = 0; paused = False
    while queue:
        r = queue.pop(next((k for k, x in enumerate(queue) if not big(x)), 0))   # documents first, large recordings last
        if a.limit and count >= a.limit: break
        prior = manifest["links"].get(r["url"], {}) or {}
        ident = r["item"]
        entry = {"year": r["year"], "applicant": r["applicant"], "label": r["label"], "link_text": r["link_text"],
                 "sheet_row": r["sheet_row"], "kind": r["kind"], "item": ident, "years": f"{r['first_year']}-{r['last_year']}", "checked": dt.date.today().isoformat()}
        if r.get("linked_from"): entry["linked_from"] = r["linked_from"]
        old_file = prior.get("file") if prior.get("status") == "archived" else None
        paths, status, note = fetch(r, workdir)
        stored = []
        if status == "ok":
            paths = [with_extension(p) for p in paths]
            if any(p.read_bytes()[:200].lstrip().lower().startswith((b"<!doctype html", b"<html")) for p in paths):
                status, note, paths = "blocked", "Drive sent a web page instead of the file (usually its large-file warning). Tried again on the next run.", []
        if status == "ok":
            # Public comments: the HSC report links to a Drive folder of them. Add any Drive links found in the report.
            if r["label"].lower().startswith("hsc recommendation") and r["kind"] != "folder":
                found, every = [], []
                for p in paths:
                    if p.read_bytes()[:4] != b"%PDF": continue
                    for u in pdf_links(p):
                        every.append(u)
                        k2, id2 = drive_ref(u)
                        cu = canon(u)
                        if k2 == "other" or cu in seen: continue
                        seen.add(cu); found.append(cu)
                        queue.append(child_rec(r, u))
                entry["scanned"] = True; entry["report_urls"] = every
                log(f"   links inside this report: {len(every)} in all, {len(found)} new Drive link(s)")
                for u in every: log(f"      {u}")
                if found: entry["links_found_in_report"] = found
            used = set()
            for e in manifest["links"].values():
                if e.get("item") == ident and e is not prior:
                    used.add(e.get("file")); used.update(f.get("file") for f in e.get("files", []))
            md = {"mediatype": "texts", "title": f"{r['applicant']} ({r['first_year'] if r['first_year'] == r['last_year'] else r['first_year'] + ' to ' + r['last_year']}): Change in Effective Control materials",
                  "creator": "Rhode Island Department of Health, Office of Health Systems Development (public filings)",
                  "date": r["last_year"], "subject": ["Rhode Island", "nursing home", "change in effective control", "certificate of need"],
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
                except Exception as e:
                    err = str(e)[:200]
                    if "syntax error detected in pdf" in str(e).lower():
                        import zipfile
                        zp = p.with_name(p.name + ".zip")
                        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z: z.write(p, fname)
                        fname = fname + ".zip"
                        try: ok = upload(ident, zp, fname, md, access, secret); err = "" if ok else err
                        except Exception as e2: err = str(e2)[:200]
                        if ok: note = "archive.org would not take this PDF as is (it says the file is damaged), so it is stored inside a zip file"
                stored.append({"file": fname, "status": "archived" if ok else "failed", "bytes": size, "sha256": digest,
                               "archive_url": f"https://archive.org/download/{ident}/{fname.replace(' ', '%20')}",
                               **({"error": err} if err else {})})
            good = [f for f in stored if f["status"] == "archived"]
            if r["kind"] == "folder":
                entry.update({"files": stored, "archive_url": f"https://archive.org/details/{ident}", "archived_on": dt.date.today().isoformat()})
                status = "archived" if stored and len(good) == len(stored) else ("partial" if good else "failed")
                if note: entry["note_extra"] = note
                if prior.get("status") == "archived" and prior.get("item") not in (None, ident):
                    for pf in prior.get("files", []):
                        if not try_delete(prior["item"], pf["file"], access, secret):
                            manifest.setdefault("cleanup", []).append({"item": prior["item"], "file": pf["file"]})
            elif stored and good:
                f = good[0]
                entry.update({k: f[k] for k in ("file", "bytes", "sha256", "archive_url")}); entry["archived_on"] = dt.date.today().isoformat()
                status = "archived"
                if old_file and (old_file != f["file"] or prior.get("item") != ident):
                    if not try_delete(prior.get("item") or ident, old_file, access, secret):
                        manifest.setdefault("cleanup", []).append({"item": prior.get("item") or ident, "file": old_file})
            elif stored and stored[0]["status"] == "too_large":
                status, note = "too_large", f"{stored[0]['bytes'] / 1e6:.0f} MB, over the {a.max_mb} MB limit"
            else:
                errs = "; ".join(x["error"] for x in stored if x.get("error"))
                status = "failed"; note = note or errs or "upload did not return success"
        entry.update({"status": status, "note": note})
        if status not in ("archived", "partial"): entry["attempts"] = int(prior.get("attempts", 0)) + 1
        manifest["links"][r["url"]] = entry
        blocked_streak = blocked_streak + 1 if status == "blocked" else 0
        save_manifest(manifest); count += 1
        extra = f" ({len(stored)} files)" if r["kind"] == "folder" and stored else ""
        log(f"[{count}] {status:10} {ident} | {r['label']}{extra} {('| ' + note) if note else ''}")
        shutil.rmtree(dl, ignore_errors=True)
        time.sleep(1)
        if blocked_streak >= 4:
            if not paused:
                log("Drive has refused 4 downloads in a row. Pausing 20 minutes, then trying again."); paused = True
                blocked_streak = 0; time.sleep(1200)
            else:
                log("Drive is still refusing downloads. Stopping here so it is not hammered. Run again in a few hours; everything not yet archived is picked up."); break

    counts = {}
    for e in manifest["links"].values(): counts[e["status"]] = counts.get(e["status"], 0) + 1
    log("Totals in manifest: " + ", ".join(f"{k} {v}" for k, v in sorted(counts.items())))
    if os.environ.get("GITHUB_STEP_SUMMARY"):
        with open(os.environ["GITHUB_STEP_SUMMARY"], "a") as f:
            f.write("### CEC archive totals\n\n" + "\n".join(f"- {k}: {v}" for k, v in sorted(counts.items())) + "\n")

main()
