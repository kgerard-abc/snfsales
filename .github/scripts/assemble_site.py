"""Copy the files the public site needs into a folder. Stops if anything looks wrong."""
import os, re, shutil, subprocess, sys

OUT = sys.argv[1] if len(sys.argv) > 1 else "_site_out"
FILES = [
    "index.html", "owners.html", "sales.json",
    "sitemap.xml", "robots.txt",
    "favicon.ico", "favicon.png", "apple-touch-icon.png", "abcri-logo.png",
]
DIRS = ["home"]
NOT_SITE = {".github", ".gitignore", "pipeline", "CNAME", "LICENSE"}
CUSTOM_DOMAIN = ""

# Every tracked top-level file or folder must be listed above, one way or the other.
tracked = subprocess.run(["git", "ls-files", "--cached", "--others", "--exclude-standard"], capture_output=True, text=True, check=True).stdout.split("\n")
tracked = [p for p in tracked if p and os.path.exists(p)]
top = {p.split("/")[0] for p in tracked if p}
unlisted = sorted(top - {f.split("/")[0] for f in FILES} - set(DIRS) - NOT_SITE)
if unlisted:
    sys.exit("Not listed in assemble_site.py (add to FILES, DIRS, or NOT_SITE): " + ", ".join(unlisted))
if not CUSTOM_DOMAIN and not os.path.exists("CNAME"):
    sys.exit("No custom domain is set here. Stopping so the site keeps its address.")

out_abs = os.path.abspath(OUT)
if out_abs == os.getcwd() or os.getcwd().startswith(out_abs + os.sep):
    sys.exit("Output folder must not be the repository folder.")
shutil.rmtree(OUT, ignore_errors=True)
os.makedirs(OUT)
for f in FILES:
    if not os.path.isfile(f):
        sys.exit(f"Missing site file: {f}")
    shutil.copy2(f, os.path.join(OUT, f))
for d in DIRS:
    extra = sorted(p for p in tracked if p.startswith(d + "/") and os.path.basename(p) != "index.html")
    if extra:
        sys.exit(f"Unexpected file under {d}/: {extra[0]}")
    for p in tracked:
        if p.startswith(d + "/"):
            os.makedirs(os.path.dirname(os.path.join(OUT, p)), exist_ok=True)
            shutil.copy2(p, os.path.join(OUT, p))
if CUSTOM_DOMAIN:
    with open(os.path.join(OUT, "CNAME"), "w") as fh:
        fh.write(CUSTOM_DOMAIN)

# Every local file a page refers to must be in the output.
ref = re.compile(r"""(?:href|src)=["']([^"'#?]+)|fetch\(['"]([^'"]+)""")
missing = set()
for root, _, names in os.walk(OUT):
    for n in names:
        if not n.endswith((".html", ".js")):
            continue
        p = os.path.join(root, n)
        for m in ref.finditer(open(p, encoding="utf-8", errors="ignore").read()):
            t = next(g for g in m.groups() if g)
            if re.match(r"^(https?:|mailto:|tel:|data:|//|javascript:)", t) or "${" in t or "'" in t or "+" in t:
                continue
            target = os.path.normpath(os.path.join(OUT, t.lstrip("/")) if t.startswith("/") else os.path.join(root, t))
            if not os.path.isfile(target) and not os.path.isfile(os.path.join(target, "index.html")):
                missing.add(f"{os.path.relpath(p, OUT)} -> {t}")
if missing:
    sys.exit("Pages point to files that are not published:\n  " + "\n  ".join(sorted(missing)))
print(f"{sum(len(n) for _, _, n in os.walk(OUT))} files ready in {OUT}")
