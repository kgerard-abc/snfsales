#!/usr/bin/env python3
"""Temporary: find completed Appendix B compliance reports in the archived nursing home CEC files (2023 to 2026).
Downloads every PDF in each Internet Archive item listed in items.json, reads each page's text (OCR for pages with
no text layer), flags compliance report pages, and saves page text, candidate page images, and candidate pages as PDFs."""
import json, os, re, subprocess, sys, time, gzip, urllib.request, urllib.parse
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
HERE = Path(__file__).parent
OUT = HERE / "out"; (OUT / "text").mkdir(parents=True, exist_ok=True); (OUT / "img").mkdir(exist_ok=True); (OUT / "pdf").mkdir(exist_ok=True)
TMP = Path("/tmp/appb"); TMP.mkdir(exist_ok=True)
UA = {"User-Agent": "abc-ri-appendix-b/1.0"}
CAND = re.compile(r"compliance\s+report|to\s+be\s+completed\s+by\s+the\s+appropriate|reviewer.{0,3}s\s+(name|signature)|appendix\s*b\b|no\s+records?\s+(were\s+)?found|change\s+in\s+effective\s+control\s+application", re.I)

def log(*a): print(*a, flush=True)
def curl(url, dest):
    for a in range(4):
        r = subprocess.run(["curl", "-sSL", "--fail", "--retry", "3", "--connect-timeout", "30", "--max-time", "1800", "-A", UA["User-Agent"], "-o", str(dest), url], capture_output=True, text=True)
        if r.returncode == 0: return True
        log("  curl", r.returncode, r.stderr[:200]); time.sleep(10 * (a + 1))
    return None

def get(url, dest=None, tries=4):
    for a in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=900) as r:
                if dest is None: return r.read()
                with open(dest, "wb") as f:
                    while True:
                        b = r.read(1 << 20)
                        if not b: break
                        f.write(b)
                return True
        except urllib.error.HTTPError as e:
            if e.code in (401, 403, 404): log("  HTTP", e.code, url); return None
            time.sleep(5 * (a + 1))
        except Exception as e:
            log("  retry", a, e); time.sleep(5 * (a + 1))
    return None

def pages_text(pdf):
    r = subprocess.run(["pdftotext", "-layout", "-q", str(pdf), "-"], capture_output=True, timeout=1800)
    t = r.stdout.decode("utf-8", "replace").split("\f")
    if t and not t[-1].strip(): t = t[:-1]
    return t

def npages(pdf):
    r = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True)
    m = re.search(r"Pages:\s+(\d+)", r.stdout); return int(m.group(1)) if m else 0

def ocr(args):
    try: return _ocr(args)
    except Exception as e: return args[1], f"(OCR failed: {e})"

def _ocr(args):
    pdf, p = args
    base = TMP / f"o{os.getpid()}_{abs(hash((str(pdf), p)))}"
    subprocess.run(["pdftoppm", "-r", "150", "-scale-to", "1800", "-gray", "-f", str(p), "-l", str(p), str(pdf), str(base)], capture_output=True, timeout=120)
    imgs = list(TMP.glob(base.name + "*"))
    if not imgs: return p, ""
    r = subprocess.run(["tesseract", str(imgs[0]), "stdout", "--psm", "3"], capture_output=True, timeout=120, env={**os.environ, "OMP_THREAD_LIMIT": "1"})
    for i in imgs: i.unlink()
    return p, r.stdout.decode("utf-8", "replace")

def main():
    shard, nsh = int(os.environ.get("SHARD", "0")), int(os.environ.get("NSHARDS", "1"))
    items = [it for i, it in enumerate(json.load(open(HERE / "items.json"))) if i % nsh == shard]
    done = set()
    for rp in OUT.glob("report*.json"):
        for r in json.load(open(rp)):
            if r.get("pages") is not None: done.add((r["item"], r.get("file")))
    seen, report = {}, []
    RPT = OUT / f"report_{shard}.json"
    for item in items:
        meta = get(f"https://archive.org/metadata/{item}")
        if not meta: log("no item", item); report.append({"item": item, "error": "no item"}); continue
        meta = json.loads(meta)
        for f in meta.get("files", []):
            name = f["name"]
            if not name.lower().endswith(".pdf"): continue
            md5 = f.get("md5")
            if (item, name) in done: continue
            if md5 in seen:
                report.append({"item": item, "file": name, "same_as": seen[md5]}); continue
            key = f"{item}/{name}"; seen[md5] = key
            tag = re.sub(r"[^\w]+", "_", item[7:40] + "_" + name[:60]).strip("_") + "_" + (md5 or "x")[:6]
            pdf = TMP / "doc.pdf"
            url = f"https://archive.org/download/{item}/{urllib.parse.quote(name)}"
            log("==", key, f.get("size"))
            t0 = time.time()
            if not curl(url, pdf): report.append({"item": item, "file": name, "error": "download failed"}); continue
            try:
                n = npages(pdf); texts = pages_text(pdf)
            except Exception as e:
                report.append({"item": item, "file": name, "error": str(e)}); continue
            texts += [""] * (n - len(texts))
            log(f"   downloaded and read text in {time.time()-t0:.0f}s, {n} pages")
            need = [i + 1 for i, t in enumerate(texts) if len(re.sub(r"\s", "", t)) < 60]
            ocred = set()
            if need:
                with ThreadPoolExecutor(4) as ex:
                    for k, (p, t) in enumerate(ex.map(ocr, [(pdf, p) for p in need])):
                        texts[p - 1] = t; ocred.add(p)
                        if k % 25 == 0: log(f"   OCR {k}/{len(need)} at {time.time()-t0:.0f}s")
            cands = []
            for i, t in enumerate(texts):
                p = i + 1
                if CAND.search(t):
                    tpl = bool(re.search(r"Reviewer.{0,3}s\s+Name:?\s*_{6,}", t))
                    cands.append({"page": p, "ocr": p in ocred, "blank_template": tpl,
                                  "received_stamp": bool(re.search(r"RECEIVED", t)), "head": re.sub(r"\s+", " ", t)[:600]})
            with gzip.open(OUT / "text" / (tag + ".txt.gz"), "wt") as g:
                for i, t in enumerate(texts): g.write(f"===== PAGE {i+1}{' OCR' if i+1 in ocred else ''} =====\n{t}\n")
            # images and single-page pdfs for compliance-report pages that are not the blank template, plus the next 6 pages if scanned
            keep = set()
            for c in cands:
                if re.search(r"compliance\s+report|reviewer", c["head"], re.I) and not c["blank_template"]:
                    keep.add(c["page"])
                    if c["ocr"]: keep.update(range(c["page"] + 1, min(n, c["page"] + 6) + 1))
                if c["received_stamp"]: keep.add(c["page"])
            for p in sorted(keep)[:80]:
                subprocess.run(["pdftoppm", "-r", "60", "-gray", "-jpeg", "-f", str(p), "-l", str(p), "-singlefile", str(pdf), str(OUT / "img" / f"{tag}_p{p}")], capture_output=True)
                subprocess.run(["qpdf", str(pdf), "--pages", ".", str(p), "--", str(OUT / "pdf" / f"{tag}_p{p}.pdf")], capture_output=True)
            report.append({"item": item, "file": name, "tag": tag, "md5": md5, "pages": n, "ocr_pages": len(ocred), "candidates": cands, "kept": sorted(keep)})
            log(f"   {n} pages, {len(ocred)} OCR, {len(cands)} candidate pages, kept {sorted(keep)[:20]}")
            json.dump(report, open(RPT, "w"), indent=1)
    json.dump(report, open(RPT, "w"), indent=1)

if __name__ == "__main__": main()
