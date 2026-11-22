"""
Write a plain web page for every nursing home on the sales site, so search engines can find each home by name.

The sales site is one page that opens a home at an address ending in #home=CCN. Google ignores everything
after the #, so a search for a home's name can only land on the front page. This script reads the data
embedded in index.html (the same data the page uses) and writes, for each home with a CMS number:

    home/CCN/index.html   (snfsales.abc-ri.org/home/415035/)

leading with who has run the home and what happened under each owner (from the owner tenures the page's
owner groups table uses: star ratings when the owner took over and now or when it left, immediate jeopardy
and harm citations, and fines while it was in charge), any pending sale, the sales and changes of control
on record, and money paid to companies the business owners also control (Medicare cost reports). A button
opens the home's full record (/#home=CCN).

It also writes home/index.html (every home, A to Z, by town), sitemap.xml, and robots.txt.

    python3 pipeline/scripts/home_pages.py

pipeline/build_site.py runs it after each build, so the monthly update and the RIDOH watcher's rebuilds
keep the pages current. Nothing here changes index.html.
"""
import html
import json
import os
import re
import shutil

SITE = "https://snfsales.abc-ri.org"
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
          "October", "November", "December"]
# Same labels as KIND_LBL and LANDLORD_TXT in the page template (ri_nh_build.py).
KIND_LBL = {"mgmt": "Buyer took over management", "closed": "Closed", "unclosed": "Approved, never completed",
            "sale": "Sold / new operator", "parent": "Chain reorganization", "internal": "Internal restructuring",
            "realestate": "Real estate only"}
LANDLORD_TXT = {
    "yes": "Yes. The home reports rent to a related landlord on Worksheet A-8-1 of its Medicare cost report, which "
           "means the people who own the operator also own the company it pays rent to.",
    "likely": "Likely. The state application was filed as a pair, one company to run the home and one to buy the "
              "building, by the same applicant. The newest cost report on file does not yet show rent to a related "
              "landlord, usually because it predates the sale.",
    "partly": "Partly. The application shows the same people in both the operating and the real estate company, but "
              "not the same ownership shares.",
    "no": "No. The real estate was bought by, or stays with, people other than the operator’s owners.",
    "no related rent": "No related-party rent reported. The home either owns its own building or rents from an "
                       "unrelated landlord; the cost report alone cannot tell those apart.",
    "unknown": "Unknown. The application does not say who bought the building and the cost report shows no "
               "related-party rent.",
}
SKIP = {"Tavares Pediatric Center"}
# Mailing addresses often name a village; list homes under the town, as the dashboard does (keep in step with
# VILLAGE_TOWN and TOWN_BY_CCN in snfdashboard's index.html).
VILLAGE_TOWN = {"Riverside": "East Providence", "Wakefield": "South Kingstown", "Greenville": "Smithfield",
                "Greene": "Coventry", "Manville": "Lincoln", "Pascoag": "Burrillville", "North Scituate": "Scituate"}
TOWN_BY_CCN = {"415107": "Exeter", "415054": "South Kingstown"}


def town_of(f):
    t = f.get("town") or ""
    town = TOWN_BY_CCN.get(str(f.get("id"))) or VILLAGE_TOWN.get(t) or t
    return town, (f"{town} ({t})" if t and town != t else town)   # not included in anything on the ABC-RI sites

E = html.escape
RI = "Rhode&nbsp;Island"
DASH = "\u2013"


def keep_ri(s):
    return E(s).replace("Rhode Island", RI)


def long_date(d):
    y, m, day = d[:10].split("-")
    return f"{MONTHS[int(m) - 1]} {int(day)}, {y}"


def mon(d):
    return f"{MONTHS[int(d[5:7]) - 1]} {d[:4]}"


def money(x):
    return "${:,.0f}".format(x)


def plural(n, one, many=None):
    return f"{n} {one if n == 1 else (many or one + 's')}"


def stars(v):
    if v is None:
        return None
    v = int(round(v))
    return "0 (special focus facility, not rated by CMS)" if v == 0 else f"{v}"


def load_data(path):
    t = open(path, encoding="utf-8").read()
    m = re.search(r'<script id="data" type="application/json">(.*?)</script>', t, re.S)
    if not m:
        raise SystemExit("index.html has no data block")
    return json.loads(m.group(1))


CSS = """
:root{--bg:#F6F3EA;--panel:#EAE5D7;--ink:#1E2A32;--ink-2:#4E5860;--line:#8F8570;--link:#1E2A32}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.55 "Public Sans",system-ui,-apple-system,"Segoe UI",Arial,sans-serif}
a{color:var(--link)}
.wrap{max-width:820px;margin:0 auto;padding:16px 20px 40px}
header.top{display:flex;align-items:center;gap:14px;padding:8px 0 14px;border-bottom:1px solid var(--line)}
header.top img{height:44px;width:auto}
header.top .site{font-weight:700;font-size:1.05rem;text-decoration:none}
h1{font-size:1.7rem;line-height:1.2;margin:22px 0 4px}
h2{font-size:1.15rem;margin:28px 0 8px}
h3{font-size:1rem;margin:0 0 4px}
.sub{color:var(--ink-2);margin:0 0 4px}
.card{background:var(--panel);border:1px solid var(--line);padding:14px 16px;margin:14px 0}
.card p{margin:6px 0}
.btn{display:inline-block;background:var(--ink);color:#fff;text-decoration:none;font-weight:700;padding:12px 18px;min-height:44px;margin:10px 0}
dl{display:grid;grid-template-columns:max-content 1fr;gap:6px 16px;margin:0}
dt{font-weight:600}
dd{margin:0}
.era{border-top:1px solid var(--line);padding:14px 0 6px}
.era p{margin:4px 0}
ul.plain{padding-left:20px}
ul.plain li{margin:0 0 8px}
.tbl{overflow-x:auto}
table{border-collapse:collapse;width:100%;font-size:.95rem}
th,td{text-align:left;padding:6px 8px;border-bottom:1px solid var(--line);vertical-align:top}
th.num{text-align:right}
td.num{text-align:right;white-space:nowrap}
td .sh{display:block;font-size:.85rem;color:var(--ink-2)}
.note{font-size:.92rem;color:var(--ink-2)}
footer{margin-top:34px;padding-top:14px;border-top:1px solid var(--line);font-size:.92rem;color:var(--ink-2)}
.towns h2{font-size:1.05rem;margin:20px 0 4px}
.towns ul{margin:0;padding-left:20px}
.towns li{margin:2px 0}
@media (max-width:520px){dl{grid-template-columns:1fr}dt{margin-top:6px}h1{font-size:1.45rem}}
"""


def page(title, desc, canonical, body):
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{E(title)}</title>
<meta name="description" content="{E(desc)}">
<link rel="canonical" href="{canonical}">
<meta property="og:type" content="website">
<meta property="og:site_name" content="Advocates for Better Care in Rhode Island">
<meta property="og:title" content="{E(title)}">
<meta property="og:description" content="{E(desc)}">
<meta property="og:url" content="{canonical}">
<meta property="og:image" content="{SITE}/abcri-logo.png">
<meta name="twitter:card" content="summary">
<link rel="icon" href="/favicon.ico" sizes="any"><link rel="icon" type="image/png" href="/favicon.png"><link rel="apple-touch-icon" href="/apple-touch-icon.png">
<style>{CSS}</style>
</head>
<body>
<div class="wrap">
<header class="top"><a href="https://www.abc-ri.org/" title="ABC-RI home page"><img src="/abcri-logo.png" alt="ABC-RI logo"></a>
<a class="site" href="/">What happened when {RI} nursing homes were sold</a></header>
<main>
{body}
</main>
<footer>
<p><a href="https://www.abc-ri.org/">Advocates for Better Care <span style="white-space:nowrap">in {RI} (ABC-RI)</span></a> is an independent advocacy organization dedicated to protecting the rights and improving the care of long-term care residents in {RI}.</p>
<p>Our public data sites: <a href="https://snfdashboard.abc-ri.org/">Nursing home dashboard</a> &middot; <a href="/">Nursing home sales</a> &middot; <a href="https://alinspections.abc-ri.org/">Assisted living inspections</a> &middot; <a href="/home/">Every home, A to Z</a></p>
<p>Built and maintained by ABC-RI from CMS and RIDOH public records. It is not produced, reviewed, or endorsed by CMS, the {RI} Department of Health (RIDOH), or any state agency.</p>
<p>Spotted an error? Email <strong>corrections@abc-ri.org</strong>.</p>
</footer>
</div>
<script data-goatcounter="https://abc-ri-snfsales.goatcounter.com/count" async src="//gc.zgo.at/count.js"></script>
</body>
</html>
"""


def took_over(x):
    """'took over management in January 2024' / 'took over ownership in March 2021' (kg's wording)."""
    if not x.get("start"):
        return None
    return f"took over {'management' if x.get('via') == 'management' else 'ownership'} {long_date(x['start'])}"


def tenure_html(x, latest):
    g = x["group"] or "Owner not named"
    counted_all = not x.get("counted_from") or (x.get("start") and x["counted_from"] <= x["start"])
    yrs = f" ({x['years']:.1f} years{' so far' if x['current'] else ''})" if counted_all and x.get("years") is not None else ""
    if not x.get("start"):
        first = "Ran the home before the state records on this site begin."
    else:
        first = f"{E(g)} {took_over(x)}"
        if x.get("via") == "management" and x.get("pending"):
            first += ", while its purchase is pending with the state"
        elif x.get("via") == "management" and x.get("owner_from"):
            first += f", before the sale was approved on {long_date(x['owner_from'])}"
        first += "."
    if x.get("end"):
        by = (f", when {E(x['end_by'])} took over {'management' if x.get('end_via') == 'management' else 'ownership'}"
              if x.get("end_by") else "")
        then = f"Ran it until {long_date(x['end'])}{by}{yrs}."
    else:
        then = f"Still in charge as of {mon(latest)}{yrs}."
    when = first + " " + then
    s0, s1 = stars(x.get("stars_start")), stars(x.get("stars_end"))
    lines = [f"<h3>{E(g)}</h3>", f"<p>{when}</p>"]
    if s0 is not None and s1 is not None:
        d = int(round(x["stars_end"])) - int(round(x["stars_start"]))
        chg = "no change" if d == 0 else f"{'up' if d > 0 else 'down'} {plural(abs(d), 'star')}"
        lines.append(f"<p><strong>Overall stars (5 is best):</strong> {s0} when it took over, {s1} "
                     f"{'now' if x['current'] else 'when it left'} ({chg}).</p>")
    elif s1 is not None:
        lines.append(f"<p><strong>Overall stars (5 is best):</strong> {s1} {'now' if x['current'] else 'when it left'}."
                     + (" No rating from before it took over is in the CMS files on this site." if x.get("start") else "")
                     + "</p>")
    insp = (f"<strong>While it ran the home:</strong> {plural(x['ij'], 'immediate jeopardy citation')}, "
            f"{plural(x['harm'], 'citation')} for harm to a resident, and "
            + (f"{plural(x['n_fines'], 'federal fine')} totaling {money(x['fines'])}." if x.get("n_fines") else "no federal fines."))
    lines.append(f"<p>{insp}</p>")
    if x.get("counted_from") and x.get("start") and x["counted_from"] > x["start"]:
        lines.append(f'<p class="note">Counted from {mon(x["counted_from"])}, when the CMS files on this site begin.</p>')
    elif not x.get("start") and x.get("counted_from"):
        lines.append(f'<p class="note">Counted from {mon(x["counted_from"])}, when the CMS files on this site begin.</p>')
    return '<div class="era">' + "".join(lines) + "</div>"


def home_page(D, f, tens, sales, latest):
    fid, name = f["id"], f["name"]
    full = f"/#home={fid}"
    town, place = town_of(f)
    closed = f.get("active") is False or str(f.get("status") or "").lower().startswith("closed")
    pending = [s for s in sales if s.get("outcome") == "Pending" and s.get("kind") == "sale"]
    body = [f"<h1>{keep_ri(name)}</h1>",
            f'<p class="sub">Nursing home in {E(place)}, {RI}' + (f" &middot; {E(f['status'])}" if closed else "") + "</p>"]
    desc = []

    if pending:
        items = []
        for p in pending:
            mg = next((x for x in tens if x.get("pending") and x.get("cec") == p.get("cec") and x.get("via") == "management"), None)
            t = f"A sale to <strong>{E(p.get('buyer') or p.get('who') or 'a new owner')}</strong> is pending with the state ({E(p['cec'])})."
            if mg:
                t += f" {E(mg['group'])} has managed the home since {long_date(mg['start'])}."
            items.append(f"<p>{t}</p>")
            desc.append(f"Sale to {p.get('buyer') or p.get('who')} pending.")
        body.append('<div class="card">' + "".join(items) + "</div>")

    cur = next((x for x in tens if x.get("current")), None)
    if cur and cur.get("group"):
        desc.insert(0, f"Run by {cur['group']}" + (f" since {mon(cur['start'])}." if cur.get("start") else "."))

    if tens:
        body.append("<h2>Who has run this home, and what happened</h2>")
        body.append('<p class="note">Newest first. An owner’s time starts when it took over management, if a '
                    "management agreement came before the sale, otherwise when it took over ownership. Star ratings "
                    "are from the last CMS file before it took over and the latest file, or the last before it left. "
                    "A Special Focus Facility counts as 0 stars.</p>")
        for x in sorted(tens, key=lambda x: x.get("start") or "", reverse=True):
            body.append(tenure_html(x, latest))
        ij = sum(x["ij"] for x in tens if x.get("current"))
        if cur and cur.get("start") and ij:
            desc.append(f"{plural(ij, 'immediate jeopardy citation')} since then.")
    body.append(f'<p><a class="btn" href="{full}">See the full record, with charts and documents</a></p>')

    # Sales and other changes of control.
    ev = [e for e in f.get("events") or [] if e.get("kind") in KIND_LBL]
    ev += [dict(u, date=u.get("date")) for u in f.get("undated") or [] if u.get("pending")]
    if ev:
        lis = []
        for e in sorted(ev, key=lambda e: e.get("date") or "9999", reverse=True):
            when = "Pending" if e.get("outcome") == "Pending" or e.get("pending") else (long_date(e["date"]) if e.get("date") else str(e.get("year") or ""))
            if e.get("date_note") == "Approximate date" and e.get("date") and not e.get("pending"):
                when += " (approximate)"
            who = e.get("buyer") or e.get("who") or ""
            if e.get("buyer") and e.get("who") and e["who"] != e["buyer"]:
                who += f" ({e['who']})"
            lis.append(f"<li><strong>{E(when)}</strong>: {E(KIND_LBL[e['kind']])}" + (f", {E(who)}" if who else "")
                       + (f" ({E(e['cec'])})" if e.get("cec") else "") + "</li>")
        body.append("<h2>Sales and changes of control on record</h2>"
                    f'<ul class="plain">{"".join(lis)}</ul>'
                    '<p class="note">From RIDOH’s Change in Effective Control records. The full record links the '
                    "applications and decision letters.</p>")

    # Money paid to companies the business owners also control.
    b = (D.get("money") or {}).get("by_facility", {}).get(fid)
    if b and b.get("years"):
        rows = []
        for k, y in enumerate(b["years"]):
            if b["rev"][k] is None and b["rp"][k] is None:
                continue
            tag = " (more than a year)" if (b.get("long") or [False] * 99)[k] else (" (part of a year)" if (b.get("partial") or [False] * 99)[k] else "")
            def cell(v, sh):
                if v is None:
                    return '<td class="num">–</td>'
                return f'<td class="num">{money(v)}' + (f'<span class="sh">{sh * 100:.1f}% of revenue</span>' if sh is not None else "") + "</td>"
            rows.append(f"<tr><td>{y}{tag}</td><td class=\"num\">{money(b['rev'][k]) if b['rev'][k] is not None else '–'}</td>"
                        + cell(b["rp"][k], b["rp_sh"][k]) + cell(b["rent"][k], b["rent_sh"][k]) + "</tr>")
        rows = rows[-6:]
        if rows:
            body.append("<h2>Money paid to companies the business owners also control</h2>"
                        '<div class="tbl"><table><thead><tr><th>Year</th><th class="num">Revenue</th>'
                        '<th class="num">Paid to all companies the business owners also control</th>'
                        '<th class="num">Rent paid to a landlord the business owners also control</th></tr></thead>'
                        f"<tbody>{''.join(rows)}</tbody></table></div>"
                        '<p class="note">From the home’s Medicare cost reports (Worksheet A-8-1, '
                        '<a href="https://www.cms.gov/data-research/statistics-trends-and-reports/cost-reports">CMS HCRIS files</a>). '
                        "Revenue is net patient revenue. The full record has every year, management fees, and margins.</p>")
            if b.get("from_summary"):
                body.append('<p class="note">Reports for this home come from the cost report summary, which has no '
                            "rent line, so rent shows as a dash.</p>")
            if b.get("landlord") and b["landlord"].lower() in LANDLORD_TXT:
                body.append(f"<p><strong>Is the landlord the same people as the operator?</strong> {E(LANDLORD_TXT[b['landlord'].lower()])}</p>")

    # About the home.
    facts = []
    if f.get("beds"):
        facts.append(("Certified beds", str(int(f["beds"]))))
    if f.get("county"):
        facts.append(("County", E(f["county"])))
    if f.get("owner_group"):
        facts.append(("Owner group", E(f["owner_group"])))
    if f.get("owner"):
        facts.append(("Operator named by CMS", E(f["owner"])))
    now = f.get("now") or {}
    if f.get("has_cms") and now.get("asof"):
        o = now.get("overall")
        if f.get("sff") == "SFF" or (o is None and f.get("sff")):
            facts.append((f"Overall stars ({mon(now['asof'])})", "0 (special focus facility, not rated by CMS)"))
        elif o is not None:
            facts.append((f"Overall stars ({mon(now['asof'])})", f"{int(o)} out of 5"))
        if now.get("hours") is not None:
            facts.append(("Nurse hours per resident per day", f"{now['hours']:.2f}"))
    former = [x["name"] for x in f.get("former") or [] if x.get("name") and x["name"].strip().lower() != name.strip().lower()]
    if former:
        facts.append(("Former names", E("; ".join(dict.fromkeys(former)))))
    if fid.isdigit():
        facts.append(("CMS number", E(fid)))
    if facts:
        body.append("<h2>About this home</h2><div class=\"card\"><dl>" +
                    "".join(f"<dt>{E(k)}</dt><dd>{v}</dd>" for k, v in facts) + "</dl>"
                    + (f'<p class="note">Ratings, staffing, and inspections for this home: '
                       f'<a href="https://snfdashboard.abc-ri.org/#f-{fid}">our nursing home dashboard</a>.</p>' if f.get("has_cms") else "")
                    + "</div>")
    body.append(f'<p class="note">CMS files through {E(mon(latest))}; RIDOH records updated as they are posted. '
                f'<a href="/">See what happened across every {RI} nursing home sale</a>.</p>')

    title = f"{name}, {town}, RI: sales, owners, and what happened under each"
    d = (f"{name}, a nursing home in {place}, Rhode Island. " + " ".join(desc) +
         " Who has owned and run it, star ratings and citations under each owner, and money paid to companies the business owners also control.")
    return page(title, d.replace("  ", " "), f"{SITE}/home/{fid}/", "\n".join(body))


def index_page(entries, latest):
    by_town = {}
    for f, cur in entries:
        by_town.setdefault(town_of(f)[0] or "Other", []).append((f, cur))
    parts = [f"<h1>Every {RI} nursing home, A to Z: sales and owners</h1>",
             "<p>Each page shows who has run the home and what happened under each owner: star ratings, immediate "
             "jeopardy and harm citations, fines, sales on record, and money paid to companies the business owners "
             f"also control. CMS files through {E(mon(latest))}.</p>",
             '<p><a class="btn" href="/">See what happened when homes were sold</a></p>', '<div class="towns">']
    for town in sorted(by_town):
        lis = "".join(f'<li><a href="/home/{f["id"]}/">{keep_ri(f["name"])}</a>'
                      + (f' <span class="note">({E(cur)})</span>' if cur else "") + "</li>"
                      for f, cur in sorted(by_town[town], key=lambda x: x[0]["name"].lower()))
        parts.append(f"<h2>{E(town)}</h2><ul>{lis}</ul>")
    parts.append("</div>")
    parts.append('<p class="note">The name in parentheses is the owner group running the home now, or the home’s '
                 "status if it is closed.</p>")
    return page("Every Rhode Island nursing home, A to Z: sales and owners",
                "A page for every Rhode Island nursing home, by town: who has owned and run it, what happened under "
                "each owner, sales on record, and money paid to companies the business owners also control.",
                f"{SITE}/home/", "\n".join(parts))


def main():
    D = load_data(os.path.join(ROOT, "index.html"))
    latest = D["meta"]["latest"]
    out = os.path.join(ROOT, "home")
    if os.path.isdir(out):
        shutil.rmtree(out)
    os.makedirs(out)
    tens_by, sales_by = {}, {}
    for x in D.get("tenures") or []:
        tens_by.setdefault(x["fid"], []).append(x)
    for s in D.get("sales") or []:
        sales_by.setdefault(s.get("fid"), []).append(s)
    entries, urls = [], [f"{SITE}/", f"{SITE}/owners.html", f"{SITE}/home/"]
    for f in D["facilities"]:
        fid = str(f.get("id") or "")
        if not fid.isdigit() or f["name"] in SKIP:
            continue
        text = home_page(D, f, tens_by.get(fid, []), sales_by.get(fid, []), latest)
        os.makedirs(os.path.join(out, fid))
        with open(os.path.join(out, fid, "index.html"), "w", encoding="utf-8") as fh:
            fh.write(text)
        closed = f.get("active") is False or str(f.get("status") or "").lower().startswith("closed")
        cur = next((x["group"] for x in tens_by.get(fid, []) if x.get("current") and x.get("group")), None)
        entries.append((f, f["status"] if closed else (cur or f.get("owner_group"))))
        urls.append(f"{SITE}/home/{fid}/")
    with open(os.path.join(out, "index.html"), "w", encoding="utf-8") as fh:
        fh.write(index_page(entries, latest))
    built = D["meta"].get("built") or latest
    with open(os.path.join(ROOT, "sitemap.xml"), "w", encoding="utf-8") as fh:
        fh.write('<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n')
        for u in urls:
            fh.write(f"  <url><loc>{u}</loc><lastmod>{built}</lastmod></url>\n")
        fh.write("</urlset>\n")
    with open(os.path.join(ROOT, "robots.txt"), "w", encoding="utf-8") as fh:
        fh.write(f"User-agent: *\nAllow: /\n\nSitemap: {SITE}/sitemap.xml\n")
    print(f"Wrote {len(entries)} home pages, home/index.html, sitemap.xml ({len(urls)} addresses), and robots.txt.")


if __name__ == "__main__":
    main()
