#!/usr/bin/env python3
"""Temporary: copy one RIDOH Drive file that RIDOH's list doesn't link (the October 9, 2025 AdviniaCare application,
which Mike Dexter of RIDOH linked in an email) to the existing archive.org item for that application, and read its
Appendix B pages."""
import os, sys, json, gzip, hashlib, re, subprocess, urllib.request
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
import scan
DRIVE_ID = "1JdDx3vudfO7DuWgc7il6hXB1vg0fMk85"
ITEM = "ri-cec-2025-adviniacare-newport-bellevue-ave-llc"
NAME = "2025.10.09 AdviniaCare CEC Applications & Exhibits resubmitted.pdf"
OUT = Path(__file__).parent / "out"
pdf = Path("/tmp/oct9.pdf")
url = f"https://drive.usercontent.google.com/download?id={DRIVE_ID}&export=download&confirm=t"
if not scan.curl(url, pdf): sys.exit("download failed")
head = pdf.read_bytes()[:5]
if head != b"%PDF-": sys.exit(f"not a PDF: {head!r}")
sha = hashlib.sha256(pdf.read_bytes()).hexdigest()
md5 = hashlib.md5(pdf.read_bytes()).hexdigest()
n = scan.npages(pdf); texts = scan.pages_text(pdf); texts += [""] * (n - len(texts))
need = [i + 1 for i, t in enumerate(texts) if len(re.sub(r"\s", "", t)) < 60]
from concurrent.futures import ThreadPoolExecutor
with ThreadPoolExecutor(4) as ex:
    for p, t in ex.map(scan.ocr, [(pdf, p) for p in need]): texts[p - 1] = t
with gzip.open(OUT / "text" / "2025_adviniacare_2025_10_09_resubmitted_Oct9.txt.gz", "wt") as g:
    for i, t in enumerate(texts): g.write(f"===== PAGE {i+1}{' OCR' if i+1 in need else ''} =====\n{t}\n")
hits = [{"page": i + 1, "head": re.sub(r"\s+", " ", t)[:300]} for i, t in enumerate(texts)
        if re.search(r"to be completed by the ap|compliance report|nevada|carson city|no records|hcqc", t, re.I)]
from internetarchive import get_item
access, secret = os.environ["IA_ACCESS"], os.environ["IA_SECRET"]
item = get_item(ITEM)
resp = item.upload(files={NAME: str(pdf)}, access_key=access, secret_key=secret, retries=8, retries_sleep=30, verbose=False)
ok = bool(resp) and all(r.status_code == 200 for r in resp)
if not resp:
    item = get_item(ITEM); ok = any(f.get("name") == NAME and int(f.get("size") or -1) == pdf.stat().st_size for f in item.files)
res = {"drive_id": DRIVE_ID, "item": ITEM, "file": NAME, "bytes": pdf.stat().st_size, "sha256": sha, "md5": md5, "pages": n,
       "ocr_pages": len(need), "uploaded": ok, "archive_url": f"https://archive.org/download/{ITEM}/{NAME.replace(' ', '%20')}", "hits": hits}
json.dump(res, open(OUT / "backup_oct9.json", "w"), indent=1)
print(json.dumps({k: v for k, v in res.items() if k != "hits"}, indent=1)); print(len(hits), "Appendix B or Nevada pages")
if not ok: sys.exit("upload failed")
