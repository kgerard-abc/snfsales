"""Temporary: save the text of a few application pages (pdftotext -layout) so the budget figures can be checked by hand."""
import json, subprocess, sys, tempfile
from pathlib import Path
sys.path.insert(0, ".github/scripts")
import cec_budgets as cb
PAGES = json.loads(Path(".github/scripts/page_check.json").read_text())
man = json.loads(cb.MANIFEST.read_text())
out = Path("pipeline/page_check"); out.mkdir(parents=True, exist_ok=True)
for src in cb.sources(man, "", None):
    fid = src["link"].split("/d/")[-1].split("/")[0]
    if fid not in PAGES: continue
    with tempfile.TemporaryDirectory() as d:
        for k, url in enumerate(src["urls"]):
            pdf = Path(d) / f"{k}.pdf"
            if not cb.http_get(url, pdf): print("no download", url); continue
            for p in PAGES[fid]:
                r = subprocess.run(["pdftotext", "-layout", "-f", str(p), "-l", str(p + 1), str(pdf), "-"], capture_output=True, text=True)
                (out / f"{fid}_{p}.txt").write_text(r.stdout)
                print(fid, p, len(r.stdout))
