#!/usr/bin/env python3
"""Pull the budget tables out of every nursing home CEC application.

RIDOH's Change in Effective Control application has the same two tables in Appendix A for each home:
  Table 2: staff (FTEs) and payroll with fringes, by job, for three past years, the current year's budget,
           and the buyer's first three projected years.
  Table 3: revenue and expenses for the same seven years.
The past years are the seller's actual figures, so they are often more recent than Medicare cost reports.

What it does:
  1. Reads pipeline/cec_archive_manifest.json for every application link ("Applicant" column on RIDOH's list).
  2. Downloads each application from our Internet Archive copy (or from RIDOH's Drive if it was not archived).
  3. Finds the Table 2 and Table 3 pages with pdftotext, then reads those pages with pdfplumber, using where each
     number sits on the page to tell which year it belongs to.
  4. Matches each home to its CEC Log row (pipeline/cec_log_snapshot.json) to get the CEC ID and outcome.
  5. Writes pipeline/cec_budgets.csv (one row per number) and pipeline/cec_budgets_report.md (what was found,
     and every check that failed, so problems can be fixed by hand).

Usage: cec_budgets.py [--only text] [--years 2025,2026] [--keep-pdfs folder]
"""
import argparse, csv, json, re, shutil, subprocess, sys, tempfile, time, urllib.request, urllib.error
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
MANIFEST = ROOT / "pipeline" / "cec_archive_manifest.json"
CEC_LOG = ROOT / "pipeline" / "cec_log_snapshot.json"
OUT_CSV = ROOT / "pipeline" / "cec_budgets.csv"
OUT_REPORT = ROOT / "pipeline" / "cec_budgets_report.md"
UA = {"User-Agent": "abc-ri-update/1.0"}
DEBUG = []  # with --debug, the raw lines of tables that failed a check go to pipeline/cec_budgets_debug.txt

T2_START = "total number of ftes"
T3_START = "complete the following table for the facility"
T3_END = ("net patient revenues (dollar value", "please provide net patient revenues")
YEAR = re.compile(r"^(?:FY[:\-]?)?\(?((?:19|20)\d\d)\)?:?$", re.I)
NUM = re.compile(r"^\(?-?\$?\s?-?\(?(?:[\d,]+(?:\.\d+)?|\.\d+)\)?$")
DASHES = {"-", "–", "—", "$-", "$–"}
COLS7 = ["actual", "actual", "actual", "budget", "projected", "projected", "projected"]

def log(*a): print(*a, flush=True)

def norm(s): return re.sub(r"\s+", " ", (s or "")).strip()

# ---------- what each line is ----------

FIN_CATS = [  # checked in order; first match wins
    ("operating_profit", r"operating\s*profit|profit\s*/?\s*\(?loss|net (income|loss)"),
    ("total_revenue", r"total\s+revenue"),
    ("total_expenses", r"total\s+expense"),
    ("net_patient_revenue", r"net\s+patient\s+revenue"),
    ("payroll", r"payroll|salar"),
    ("bad_debt", r"bad\s+debt"),
    ("management_and_other", r"(mgmt|management).*(tax|r&m|repair|residual)|(tax|r&m|repair).*(mgmt|management)"),
    ("management_fee", r"mgmt|management|home office|central (office|services)"),
    ("consulting", r"consult"),
    ("leasehold", r"leasehold"),
    ("rent", r"lease|rent"),
    ("supplies", r"suppl"),
    ("office", r"office|admin|general"),
    ("utilities", r"utilit"),
    ("insurance", r"insur"),
    ("interest", r"interest"),
    ("depreciation", r"depreci|amortiz"),
    ("provider_assessment", r"provider\s+(assessment|tax)|user\s+fee|bed\s+tax"),
    ("contracted_services", r"contract|purchased|agency"),
    ("taxes", r"tax"),
]
STAFF_CATS = [
    ("total", r"^totals?\b"),
    ("medical_director", r"medical\s+director"),
    ("physicians", r"physician"),
    ("director_of_nursing", r"director\s+of\s+nurs|\bdon\b|\bdns\b"),
    ("administrator", r"administrator"),
    ("admin_nursing", r"administrative\s+nurs|nurse\s+manag|unit\s+manag|mds"),
    ("rn", r"\brns?\b|registered"),
    ("lpn", r"\blpns?\b|practical"),
    ("aides", r"aide|cna|nursing\s+assist"),
    ("ot_st", r"\bots?\b.*\bsts?\b|occupational.*speech"),
    ("pt", r"\bpts?\b|physical\s+therap"),
    ("ot", r"\bots?\b|occupational"),
    ("speech", r"speech|\bsts?\b|\bslps?\b"),
    ("clerical", r"clerical|office|business"),
    ("dietary_housekeeping", r"dietary.*housekeep|housekeep.*dietary"),
    ("housekeeping", r"housekeep|laundry|environmental"),
    ("dietary", r"dietary|food|kitchen"),
    ("maintenance", r"maint"),
    ("social_services", r"social"),
    ("recreation", r"recreat|activit"),
    ("respiratory", r"respirat"),
    ("marketing", r"market|admission"),
]

def categorize(label, cats, section=None):
    low = label.lower()
    if cats is FIN_CATS and section == "revenues" and not re.search(r"total|net\s+patient", low):
        return "other_revenue"
    for cat, pat in cats:
        if re.search(pat, low):
            return cat
    return "other"

# ---------- numbers ----------

def parse_num(tok):
    t = tok.replace("$", "").replace(" ", "").replace("#", "")
    if t in ("", "-", "–", "—"): return None
    neg = t.startswith("(") and t.endswith(")") or t.startswith("-") or t.startswith("(-")
    t = t.strip("()").lstrip("-").strip("()")
    try: v = float(t.replace(",", ""))
    except ValueError: return None
    return -v if neg else v

def is_num(tok):
    return bool(NUM.match(tok)) or tok in DASHES or bool(re.fullmatch(r"#\d+(\.\d+)?", tok)) or tok in ("#", "$")

# ---------- reading pages ----------

def page_texts(pdf):
    """Plain text of every page, using pdftotext (fast). Returns a list, one string per page."""
    r = subprocess.run(["pdftotext", "-q", str(pdf), "-"], capture_output=True, timeout=600)
    return r.stdout.decode("utf-8", "replace").split("\f")

def lines_with_words(pdf, pageno):
    """Words on one page (1-based) grouped into lines, each word with its x position."""
    import pdfplumber
    with pdfplumber.open(str(pdf)) as doc:
        if pageno > len(doc.pages): return []
        words = doc.pages[pageno - 1].extract_words(keep_blank_chars=False, x_tolerance=1.5, y_tolerance=2)
    words.sort(key=lambda w: (round(w["top"]), w["x0"]))
    lines, cur, top = [], [], None
    for w in words:
        if top is not None and abs(w["top"] - top) > 3:
            lines.append(cur); cur = []
        if not cur: top = w["top"]
        cur.append({"t": w["text"], "x0": w["x0"], "x1": w["x1"], "xc": (w["x0"] + w["x1"]) / 2})
    if cur: lines.append(cur)
    for ln in lines: ln.sort(key=lambda w: w["x0"])
    return lines

def text_of(ln): return " ".join(w["t"] for w in ln)

def year_columns(ln):
    """x positions of the fiscal years named on a header line, merging 'FY:' and '2023' when split."""
    out = []
    for w in ln:
        m = YEAR.match(w["t"].replace(" ", ""))
        if m: out.append((int(m.group(1)), w["xc"]))
    return out

# ---------- Table 3: revenue and expenses ----------

def parse_table3(lines, start_idx):
    """Rows from the line after the Table 3 question to the profit/loss line. Returns (years, rows, notes)."""
    years, rows, notes = [], [], []
    section, pending, i = "revenues", [], start_idx + 1
    first_x = None
    while i < len(lines):
        ln = lines[i]; low = text_of(ln).lower()
        if any(e in low for e in T3_END): break
        yc = year_columns(ln)
        if len(yc) >= 3 and not rows:
            years = yc if len(yc) > len(years) else years
            first_x = min(x for _, x in years) - 25
            i += 1; continue
        if re.match(r"^\s*revenues?\b", low) and not [w for w in ln if is_num(w["t"])]:
            section = "revenues"; i += 1; continue
        if re.match(r"^\s*expenses?\b", low) and not [w for w in ln if is_num(w["t"]) and parse_num(w["t"]) is not None]:
            section = "expenses"; pending = []; i += 1; continue
        if not years:
            i += 1; continue
        nums = [w for w in ln if w["xc"] >= first_x and is_num(w["t"])]
        label_words = [w["t"] for w in ln if w["xc"] < first_x]
        label = norm(" ".join(label_words))
        if not nums:
            if label and rows and (label[:1].islower() or label.startswith(")") or rows[-1]["label"].count("(") > rows[-1]["label"].count(")")):
                rows[-1]["label"] = norm(rows[-1]["label"] + " " + label)
            elif label and not re.fullmatch(r"[-–\s]*", label):
                pending.append(label)
            i += 1; continue
        if not label and i + 1 < len(lines):  # numbers sit between two halves of a wrapped label
            nxt = lines[i + 1]
            if not [w for w in nxt if is_num(w["t"])] and not year_columns(nxt):
                label = norm(" ".join(w["t"] for w in nxt if w["xc"] < first_x)); i += 1
        label = norm(" ".join(pending + [label])); pending = []
        vals = {}
        for w in nums:
            yr = min(years, key=lambda y: abs(y[1] - w["xc"]))
            v = parse_num(w["t"])
            if yr[0] in vals and v is not None and vals[yr[0]] is not None:
                notes.append(f"two numbers fell under {yr[0]} on line '{label}'")
            if v is not None or yr[0] not in vals: vals[yr[0]] = v
        rows.append({"label": label, "section": section, "vals": vals})
        if re.search(r"operating|profit|loss", label, re.I): break
        if section == "revenues" and re.search(r"total\s+revenue", label, re.I): section = "expenses"
        i += 1
    return [y for y, _ in sorted(years, key=lambda t: t[1])], rows, notes

# ---------- Table 2: staff and payroll ----------

def parse_table2(lines, start_idx, stop_idx):
    """Two blocks (past and current years, then projected years). Returns (blocks, notes)."""
    blocks, cur, notes, pending = [], None, [], []
    for ln in lines[start_idx + 1:stop_idx]:
        yc = year_columns(ln)
        if len(yc) >= 2:
            cur = {"years": sorted(yc, key=lambda t: t[1]), "rows": []}; blocks.append(cur); pending = []; continue
        if cur is None: continue
        first_x = cur["years"][0][1] - 45
        nums = [w for w in ln if w["xc"] >= first_x and is_num(w["t"])]
        label = norm(" ".join(w["t"] for w in ln if w["xc"] < first_x))
        if not nums:
            if label and not re.search(r"personnel|number|payroll|fringes|ftes|categories|fiscal|projected|budgeted|^of$", label, re.I):
                pending.append(label)
            continue
        label = norm(" ".join(pending + [label])); pending = []
        if not label: continue
        vals = defaultdict(dict)
        for w in nums:
            yr = min(cur["years"], key=lambda y: abs(y[1] - w["xc"]))[0]
            v = parse_num(w["t"])
            if v is None: continue
            kind = "fte" if (w["t"].startswith("#") or (v < 1000 and "$" not in w["t"])) else "payroll"
            vals[yr][kind] = v
        cur["rows"].append({"label": label, "vals": dict(vals)})
    return blocks, notes

# ---------- homes and the CEC Log ----------

STOP = set("the of and at in on a an llc inc d b a dba center centre nursing rehabilitation rehab health healthcare care home "
           "facility skilled operations operating operator opco propco opco, associates limited lp ri rhode island for".split())

def name_tokens(s):
    return {t for t in re.findall(r"[a-z0-9']+", (s or "").lower()) if t not in STOP and len(t) > 1}

def load_cec_log():
    rows = json.loads(CEC_LOG.read_text())
    return [{"id": r[0], "facility": r[1], "town": r[2], "buyer": r[4], "outcome": r[9], "date": r[10],
             "year": int(r[11]) if str(r[11]).isdigit() else None} for r in rows.values()]

def match_home(home, applicant, app_year, cec):
    toks = name_tokens(home)
    best, score = None, 0
    for r in cec:
        if r["year"] is None or not (app_year - 1 <= r["year"] <= app_year + 2): continue
        ft = name_tokens(r["facility"]) | name_tokens(r["town"])
        # The home name decides; the applicant name only breaks ties (one PDF can hold several homes).
        s = len(toks & ft) * 2 + min(len(name_tokens(applicant) & (name_tokens(r["buyer"]) | ft)), 2) * 0.25
        if s > score: best, score = r, s
    return (best, score) if score >= 2 else (None, score)

def home_name(texts, pageno):
    """Facility name from the nearest earlier cover page ('Legal Name of Facility') or table heading."""
    for p in range(pageno - 1, max(-1, pageno - 60), -1):
        t = texts[p]
        m = re.search(r"Legal\s+Name\s+of\s+(?:the\s+)?Facility:?\s*(.+?)(?:\n\s*\n|Date of Submission|MM/DD)", t, re.S | re.I)
        if m:
            nm = norm(m.group(1))
            m2 = re.search(r"d/b/a\s+(.+)", nm, re.I)
            return norm(m2.group(1) if m2 else nm)[:120]
    return ""

# ---------- downloads ----------

def http_get(url, dest, tries=3):
    for a in range(tries):
        try:
            req = urllib.request.Request(url, headers=UA)
            with urllib.request.urlopen(req, timeout=600) as r, open(dest, "wb") as f:
                shutil.copyfileobj(r, f, 1 << 20)
            return True
        except urllib.error.HTTPError as e:
            if e.code in (401, 403, 404): return False
            time.sleep(5 * (a + 1))
        except Exception:
            time.sleep(5 * (a + 1))
    return False

def sources(manifest, only, years):
    """One entry per application document: (applicant, year, item, label url, list of download urls)."""
    links = manifest["links"]
    other_files = {v.get("file") for v in links.values() if v.get("label") != "Applicant"}
    seen, out = set(), []
    for url, v in sorted(links.items(), key=lambda kv: (kv[1].get("year", ""), kv[1].get("applicant", ""))):
        if v.get("label") != "Applicant": continue
        if only and only.lower() not in v.get("applicant", "").lower(): continue
        if years and v.get("year") not in years: continue
        key = v.get("sha256") or url
        if key in seen: continue
        seen.add(key)
        urls = []
        if v.get("status") == "archived" and v.get("archive_url"):
            urls = [v["archive_url"]]
        elif v.get("status") == "archived" and v.get("kind") == "folder" and v.get("item"):
            try:
                meta = json.load(urllib.request.urlopen(urllib.request.Request(
                    f"https://archive.org/metadata/{v['item']}", headers=UA), timeout=60))
                urls = [f"https://archive.org/download/{v['item']}/{urllib.parse.quote(f['name'])}"
                        for f in meta.get("files", []) if f["name"].lower().endswith(".pdf") and f["name"] not in other_files]
            except Exception as e:
                log("  could not list", v["item"], e)
        elif v.get("kind") == "file":
            fid = re.search(r"/file/d/([\w-]+)", url)
            if fid: urls = [f"https://drive.usercontent.google.com/download?id={fid.group(1)}&export=download&confirm=t"]
        out.append({"applicant": v.get("applicant", ""), "year": int(v.get("year") or 0), "item": v.get("item", ""),
                    "link": url, "status": v.get("status", ""), "kind": v.get("kind", ""), "urls": urls})
    return out

# ---------- one PDF ----------

def extract(pdf, src, cec, out_rows, report):
    texts = page_texts(pdf)
    if sum(len(t.strip()) for t in texts[:20]) < 200:
        report.append(f"- **No text in the PDF** ({len(texts)} pages). It is a scan and needs OCR.")
        return 0
    flat = [norm(t).lower() for t in texts]
    t3_pages = [i + 1 for i, t in enumerate(flat) if T3_START in t]
    t2_pages = [i + 1 for i, t in enumerate(flat) if T2_START in t]
    if not t3_pages and not t2_pages:
        report.append(f"- No Appendix A tables found in {len(texts)} pages.")
        return 0
    found = 0
    for p3 in t3_pages:
        lines = lines_with_words(pdf, p3) + lines_with_words(pdf, p3 + 1)
        idx = next((k for k, ln in enumerate(lines) if T3_START in text_of(ln).lower()), None)
        if idx is None: continue
        years, rows, notes = parse_table3(lines, idx)
        dump = lambda n: DEBUG.append(f"===== {src['year']} {src['applicant']} | page {p3} | {'; '.join(n)}\n" + "\n".join(
            " ".join(f"{w['t']}@{w['xc']:.0f}" for w in ln) for ln in lines[idx: idx + 70]))
        home = home_name(texts, p3)
        heading = ""
        for ln in lines[idx: idx + 4]:
            m = re.match(r"(.+?)\s+Past\s+Three", text_of(ln))
            if m and not re.search(r"fiscal|years|budget", m.group(1), re.I): heading = norm(m.group(1))
        home = heading or home or src["applicant"]
        r, score = match_home(home, src["applicant"], src["year"], cec)
        cols = COLS7 if len(years) == 7 else None
        if not cols:
            notes.append(f"found {len(years)} year columns ({years}), expected 7; column types left blank")
        # Check: expense lines should add up to Total Expenses in each year.
        tot = next((x for x in rows if categorize(x["label"], FIN_CATS, x["section"]) == "total_expenses"), None)
        if tot:
            for y in years:
                parts = [x["vals"].get(y) for x in rows if x["section"] == "expenses"
                         and categorize(x["label"], FIN_CATS, x["section"]) not in ("total_expenses", "operating_profit")]
                s = sum(v for v in parts if v)
                t = tot["vals"].get(y)
                if t and abs(s - t) > max(2, abs(t) * 0.005):
                    notes.append(f"{y}: expense lines add to {s:,.0f} but Total Expenses says {t:,.0f}")
        for x in rows:
            cat = categorize(x["label"], FIN_CATS, x["section"])
            for k, y in enumerate(years):
                if y not in x["vals"]: continue
                out_rows.append({**base(src, p3, home, r), "table": "finances", "section": x["section"], "line": x["label"],
                                 "category": cat, "fy": y, "column": cols[k] if cols else "", "fte": "",
                                 "amount": fmt(x["vals"][y])})
        if notes: dump(notes)
        found += 1
        report.append(f"- Table 3, page {p3}: **{home}** ({match_text(r, score)}), years {years[0] if years else '?'}-{years[-1] if years else '?'}, "
                      f"{len(rows)} lines." + ("" if not notes else " **Check:** " + "; ".join(notes)))
    for p2 in t2_pages:
        nxt = next((p for p in t3_pages if p >= p2), p2 + 2)
        lines = []
        for p in range(p2, min(nxt, p2 + 3) + 1): lines += lines_with_words(pdf, p)
        idx = next((k for k, ln in enumerate(lines) if T2_START in text_of(ln).lower()), None)
        stop = next((k for k, ln in enumerate(lines) if T3_START in text_of(ln).lower()), len(lines))
        if idx is None: continue
        blocks, notes = parse_table2(lines, idx, stop)
        home = home_name(texts, p2) or src["applicant"]
        for ln in lines[idx: idx + 4]:
            m = re.match(r"(.+?)\s+Past\s+Three", text_of(ln))
            if m and not re.search(r"fiscal|years|budget|personnel", m.group(1), re.I): home = norm(m.group(1))
        r, score = match_home(home, src["applicant"], src["year"], cec)
        nrows = 0
        for b, blk in enumerate(blocks[:2]):
            yrs = [y for y, _ in blk["years"]]
            kinds = (["actual"] * (len(yrs) - 1) + ["budget"]) if b == 0 else ["projected"] * len(yrs)
            tot = next((x for x in blk["rows"] if categorize(x["label"], STAFF_CATS) == "total"), None)
            for y in yrs:
                if tot and tot["vals"].get(y, {}).get("payroll"):
                    s = sum(x["vals"].get(y, {}).get("payroll") or 0 for x in blk["rows"] if x is not tot)
                    t = tot["vals"][y]["payroll"]
                    if abs(s - t) > max(2, t * 0.005): notes.append(f"{y}: payroll lines add to {s:,.0f} but Totals says {t:,.0f}")
            for x in blk["rows"]:
                cat = categorize(x["label"], STAFF_CATS)
                for k, y in enumerate(yrs):
                    v = x["vals"].get(y)
                    if not v: continue
                    out_rows.append({**base(src, p2, home, r), "table": "staff", "section": "", "line": x["label"],
                                     "category": cat, "fy": y, "column": kinds[k], "fte": fmt(v.get("fte")),
                                     "amount": fmt(v.get("payroll"))})
                    nrows += 1
        if len(blocks) < 2: notes.append(f"found {len(blocks)} year blocks, expected 2 (past and projected)")
        if notes:
            DEBUG.append(f"===== {src['year']} {src['applicant']} | page {p2} | {'; '.join(notes)}\n" + "\n".join(
                " ".join(f"{w['t']}@{w['xc']:.0f}" for w in ln) for ln in lines[idx: min(stop, idx + 90)]))
        found += 1
        report.append(f"- Table 2, page {p2}: **{home}** ({match_text(r, score)}), {nrows} numbers."
                      + ("" if not notes else " **Check:** " + "; ".join(notes)))
    return found

def base(src, page, home, r):
    return {"applicant": src["applicant"], "application_year": src["year"], "archive_item": src["item"],
            "link": src["link"], "pdf_page": page, "home": home,
            "cec_id": r["id"] if r else "", "cec_facility": r["facility"] if r else "",
            "cec_outcome": r["outcome"] if r else ""}

def match_text(r, score):
    return f"{r['id']}, {r['facility']}, {r['outcome']}" if r else "no CEC Log match, check by hand"

def fmt(v):
    if v is None: return ""
    return str(int(v)) if float(v).is_integer() else f"{v:.2f}"

FIELDS = ["applicant", "application_year", "archive_item", "link", "pdf_page", "home", "cec_id", "cec_facility",
          "cec_outcome", "table", "section", "line", "category", "fy", "column", "fte", "amount"]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", default="")
    ap.add_argument("--years", default="")
    ap.add_argument("--keep-pdfs", default="")
    ap.add_argument("--debug", action="store_true")
    a = ap.parse_args()
    years = {y.strip() for y in a.years.split(",") if y.strip()}
    manifest = json.loads(MANIFEST.read_text())
    cec = load_cec_log()
    srcs = sources(manifest, a.only, years)
    log(f"{len(srcs)} application documents to read")
    out_rows, report = [], []
    summary = defaultdict(int)
    tmp = Path(tempfile.mkdtemp())
    for n, src in enumerate(srcs, 1):
        report.append(f"\n### {src['year']}: {src['applicant']}\n\n[RIDOH link]({src['link']})"
                      + (f" | [our copy](https://archive.org/details/{src['item']})" if src["status"] == "archived" else ""))
        if not src["urls"]:
            report.append(f"- Not read: no copy to download (archive status: {src['status']}, {src['kind']}).")
            summary["no copy"] += 1; continue
        got_any = 0
        for k, url in enumerate(src["urls"]):
            pdf = tmp / f"{n}_{k}.pdf"
            log(f"[{n}/{len(srcs)}] {src['year']} {src['applicant'][:60]}")
            if not http_get(url, pdf):
                report.append(f"- Download failed: {url}"); continue
            if pdf.read_bytes()[:5] != b"%PDF-":
                report.append(f"- Not a PDF: {url}"); pdf.unlink(); continue
            try:
                got_any += extract(pdf, src, cec, out_rows, report)
            except Exception as e:
                report.append(f"- Error while reading: {e}")
            if a.keep_pdfs:
                Path(a.keep_pdfs).mkdir(parents=True, exist_ok=True); shutil.copy(pdf, Path(a.keep_pdfs) / pdf.name)
            pdf.unlink(missing_ok=True)
        summary["tables found" if got_any else "no tables found"] += 1
    shutil.rmtree(tmp, ignore_errors=True)
    out_rows.sort(key=lambda r: (r["application_year"], r["applicant"], r["home"], r["table"], r["pdf_page"]))
    with open(OUT_CSV, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS); w.writeheader(); w.writerows(out_rows)
    homes = {(r["applicant"], r["home"]) for r in out_rows}
    head = [
        "# Budget tables from CEC applications",
        "",
        f"Made by .github/scripts/cec_budgets.py on {time.strftime('%B %-d, %Y')}. The numbers are in cec_budgets.csv.",
        "",
        "Each application's Appendix A has Table 2 (staff and payroll by job) and Table 3 (revenue and expenses) for three "
        "past years (the seller's actual figures), the current year's budget, and the buyer's first three projected years. "
        "Applicants label the \"Other\" expense lines in their own words, so the category column is a best guess from the "
        "line's wording; the line column keeps the wording as printed.",
        "",
        "A **Check** note means a table did not add up or was not laid out as expected. Look at the page before using "
        "those numbers. Homes are matched to the CEC Log by name; \"sale never closed\" in cec_outcome means the sale "
        "did not go through (for example Kadima, 2022), so leave those out of any before-and-after comparison.",
        "",
        f"Applications read: {len(srcs)}. With tables: {summary['tables found']}. Without tables: {summary['no tables found']}. "
        f"No copy to read: {summary['no copy']}. Homes with numbers: {len(homes)}. Numbers saved: {len(out_rows)}.",
    ]
    OUT_REPORT.write_text("\n".join(head + report) + "\n")
    if a.debug:
        (ROOT / "pipeline" / "cec_budgets_debug.txt").write_text("\n\n".join(DEBUG) + "\n")
    log("\n".join(head[-1:]))

if __name__ == "__main__":
    import urllib.parse
    main()
