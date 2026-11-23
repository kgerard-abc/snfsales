"""Upload RIDOH nurse and nursing assistant discipline orders to one archive.org item."""
import os, sys, glob, time
from internetarchive import get_item

ITEM = "ri-doh-nurse-aide-discipline-orders"
MD = {
    "title": "Rhode Island Department of Health disciplinary orders: nursing assistants, medication aides, and nurses (copied November 23, 2026)",
    "mediatype": "texts",
    "creator": "Rhode Island Department of Health",
    "contributor": "Advocates for Better Care in Rhode Island (ABC-RI)",
    "date": "2026-11-23",
    "subject": ["Rhode Island", "nursing assistants", "nurses", "professional discipline", "long-term care"],
    "source": "https://datahealth.ri.gov/lists/disciplinaryactions/",
    "description": ("Copies of every disciplinary order RIDOH posted under Nursing Assistant and Nursing "
                    "(RN, LPN, APRN) as of November 23, 2026: 429 PDFs from https://datahealth.ri.gov/discipline/. "
                    "index.csv lists each action as RIDOH shows it. The text folder has OCR text of each order, "
                    "which may contain errors; the PDF is the record."),
}
access, secret = os.environ.get("IA_ACCESS"), os.environ.get("IA_SECRET")
if not access or not secret:
    sys.exit("IA_ACCESS and IA_SECRET are not set for this repository.")
files = {}
for p in sorted(glob.glob("discipline/pdf/*.pdf")):
    files["pdf/" + os.path.basename(p)] = p
for p in sorted(glob.glob("discipline/text/*.txt")):
    files["ocr-text/" + os.path.basename(p)] = p
files["index.csv"] = "discipline/index.csv"
item = get_item(ITEM)
have = {f.get("name"): int(f.get("size") or -1) for f in (item.files or [])}
todo = {k: v for k, v in files.items() if have.get(k) != os.path.getsize(v)}
print(f"{len(files)} files; {len(files) - len(todo)} already on archive.org; uploading {len(todo)}", flush=True)
first = not item.exists
ok = fail = 0
for name, path in todo.items():
    for attempt in range(3):
        try:
            r = item.upload(files={name: path}, metadata=(MD if first else None), access_key=access, secret_key=secret,
                            retries=8, retries_sleep=30, verbose=False, validate_identifier=True)
            if r and all(x.status_code == 200 for x in r):
                ok += 1; first = False; break
            raise RuntimeError(",".join(str(x.status_code) for x in r))
        except Exception as e:
            print("retry", name, str(e)[:120], flush=True); time.sleep(20)
            item = get_item(ITEM)
    else:
        fail += 1; print("FAILED", name, flush=True)
print(f"uploaded {ok}, failed {fail}. https://archive.org/details/{ITEM}")
sys.exit(1 if fail else 0)
