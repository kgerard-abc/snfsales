"""One-page printable summaries for the Health Services Council: one per buyer group with a pending application.

build(D, hu, pu, out_dir) writes, for each pending buyer group:
    council/<slug>.html        one letter-size page, 11pt Helvetica, black on white
    council/<slug>-large.html  the same content in large print (18pt), several pages
and council/index.html listing them. Everything comes from the page data D (after enrich) and CMS's citation (hu)
and penalty (pu) tables, so the sheets update with every build. Short "who is behind it" notes come from
group_names.json ("about"), which ABC-RI writes and checks by hand.
"""
import html
import os
import re

import pandas as pd

IJ, HARM = set("JKL"), set("GHIJKL")


def esc(s):
    return html.escape(str(s if s is not None else ""))


def ri(s):
    """Keep 'Rhode Island' on one line."""
    return s.replace("Rhode Island", "Rhode&nbsp;Island")


def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")


def money(v):
    return "$" + f"{round(v or 0):,}"


def plural(n, word, many=None):
    return f"{n:,} {word if n == 1 else (many or word + 's')}"


def stars(v):
    return "–" if v is None else f"{v:.1f}".rstrip("0").rstrip(".") if float(v).is_integer() else f"{v:.1f}"


def chain_name(cms, g):
    """CMS writes chain names in capitals. Use ABC-RI's spelling when it is the same name, otherwise title case."""
    k = lambda x: re.sub(r"[^a-z0-9]", "", x.lower())
    return g if k(cms) == k(g) else (cms.title() if cms.isupper() else cms)


def nice_date(iso):
    t = pd.Timestamp(iso)
    return t.strftime("%B ") + str(t.day) + t.strftime(", %Y")


def month_year(iso):
    return pd.Timestamp(iso).strftime("%B %Y")


ABC = ('<span class="nw">Advocates for Better Care</span> <span class="nw">in Rhode&nbsp;Island</span>')

CSS = """
@page { size: letter; margin: 0.5in 0.6in; }
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
body { font-family: Helvetica, Arial, sans-serif; color: #000; background: #F6F3EA; font-size: 11pt; line-height: 1.28; }
.paper { background: #fff; max-width: 8.5in; margin: 16px auto; padding: 0.55in 0.6in; border: 1px solid #000; }
.screen-only { max-width: 8.5in; margin: 12px auto 0; font-size: 11pt; }
@media print {
  body { background: #fff; }
  .paper { margin: 0; padding: 0; border: none; max-width: none; }
  .screen-only { display: none; }
}
.nw { white-space: nowrap; }
.top { display: flex; justify-content: space-between; gap: 12px; border-bottom: 1px solid #000; padding-bottom: 4px; margin-bottom: 8px; font-size: 0.9em; }
h1 { font-size: 1.45em; line-height: 1.15; margin: 0 0 4px; }
h2 { font-size: 1.08em; margin: 10px 0 3px; break-after: avoid; page-break-after: avoid; }
p { margin: 0 0 6px; orphans: 3; widows: 3; }
ul { margin: 0 0 6px; padding-left: 1.2em; }
li { margin: 0 0 2px; orphans: 3; widows: 3; break-inside: avoid; }
table { width: 100%; border-collapse: collapse; margin: 2px 0 4px; font-size: 0.93em; }
th, td { text-align: left; vertical-align: top; padding: 3px 5px 3px 0; border-bottom: 1px solid #000; }
th { font-weight: bold; border-bottom: 1.5px solid #000; }
td.n, th.n { text-align: right; }
tr { break-inside: avoid; page-break-inside: avoid; }
.small { font-size: 0.86em; }
.chart { margin: 4px 0 2px; break-inside: avoid; page-break-inside: avoid; }
.chart svg { width: 100%; height: auto; display: block; }
.home { break-inside: avoid; page-break-inside: avoid; border-top: 1px solid #000; padding-top: 6px; margin-top: 8px; }
.foot { border-top: 1px solid #000; margin-top: 10px; padding-top: 4px; }
"""
LARGE = """
body { font-size: 18pt; line-height: 1.35; }
@page { margin: 0.6in; }
h1 { font-size: 1.35em; }
h2 { font-size: 1.1em; margin-top: 18px; }
.top { display: block; }
.top span { display: block; }
.small, table { font-size: 1em; }
"""


def bar_chart(rows, large=False):
    """Horizontal bars of average CMS overall stars, 0 to 5. rows = [(label, value or None, note)].
    Large print puts each label on its own line above its bar, so the text can stay big."""
    rows = [r for r in rows if r[1] is not None]
    if not rows:
        return ""
    if large:
        W, fs, bh, lw = 520, 19, 24, 0      # drawn about 7.3 inches wide, so 19 units is about 19pt
        step = fs + 6 + bh + 14
    else:
        W, fs, bh, lw = 720, 12.5, 15, 250
        step = bh + 6
    top, pw = 8, W - lw - 60
    H = top + len(rows) * step + fs + 14
    x = lambda v: lw + 4 + v / 5 * pw
    out = [f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="Average CMS overall star rating, 0 to 5" '
           f'font-family="Helvetica, Arial, sans-serif" font-size="{fs}">']
    for t in range(6):
        if large and t:   # no grid lines in large print, so they don't cross the labels
            out.append(f'<text x="{x(t):.1f}" y="{H - 4}" text-anchor="middle">{t}</text>'); continue
        out.append(f'<line x1="{x(t):.1f}" x2="{x(t):.1f}" y1="{top - 4}" y2="{H - fs - 8}" stroke="#000" '
                   f'stroke-width="{1 if t == 0 else 0.4}"{"" if t == 0 else " stroke-dasharray=\"2 3\""}/>')
        out.append(f'<text x="{x(t):.1f}" y="{H - 4}" text-anchor="middle">{t}</text>')
    for i, (lab, v, _) in enumerate(rows):
        y = top + i * step
        if large:
            out.append(f'<text x="{x(0):.1f}" y="{y + fs:.1f}">{esc(lab)}</text>')
            y += fs + 6
        else:
            out.append(f'<text x="{lw - 6}" y="{y + bh * 0.75:.1f}" text-anchor="end">{esc(lab)}</text>')
        out.append(f'<rect x="{x(0):.1f}" y="{y:.1f}" width="{max(1.5, x(v) - x(0)):.1f}" height="{bh}" fill="#000"/>')
        out.append(f'<text x="{x(v) + 6:.1f}" y="{y + bh * 0.75:.1f}" font-weight="bold">{v:.1f}</text>')
    out.append("</svg>")
    return "".join(out)


def build(D, hu, pu, out_dir, names=None):
    names = names or {}
    about = names.get("about", {})
    os.makedirs(out_dir, exist_ok=True)
    fac = {f["id"]: f for f in D["facilities"]}
    latest = pd.Timestamp(D["meta"]["latest"])
    cut = latest - pd.DateOffset(years=3)
    hu = hu.assign(sd=pd.to_datetime(hu.survey_date, errors="coerce")).drop_duplicates(
        ["ccn_std", "survey_date", "deficiency_tag_number", "scope_severity_code"])
    fines = pu[pu.penalty_type == "Fine"].assign(pd_=pd.to_datetime(pu.penalty_date, errors="coerce")).drop_duplicates(
        ["ccn_std", "penalty_date", "fine_amount"])

    def three_years(fid):
        h = hu[(hu.ccn_std == fid) & (hu.sd >= cut)]
        sc = h.scope_severity_code.astype(str)
        p = fines[(fines.ccn_std == fid) & (fines.pd_ >= cut)]
        return int(sc.isin(IJ).sum()), int(sc.isin(HARM).sum()), float(p.fine_amount.sum())

    def overall(f):
        now = f.get("now") or {}
        return 0.0 if f.get("sff") == "SFF" else now.get("overall")

    active = [f for f in D["facilities"] if f.get("active", True) and f.get("has_cms") and overall(f) is not None]
    ri_avg = sum(overall(f) for f in active) / len(active) if active else None
    ri_hours = (D.get("state") or {}).get("avg_hours")
    if isinstance(ri_hours, list):   # a monthly series: use the newest month
        ri_hours = next((v for v in reversed(ri_hours) if v is not None), None)
    natl = D.get("chain_national") or {}

    groups = {}
    for s in D["sales"]:
        if s.get("outcome") == "Pending" and s.get("buyer") and s.get("fid") in fac:
            groups.setdefault(s["buyer"], []).append(s)
    made = []
    for g, ss in sorted(groups.items(), key=lambda kv: (-len(kv[1]), kv[0])):
        ss.sort(key=lambda s: fac[s["fid"]]["name"])
        tr = (D.get("track") or {}).get(g) or {}
        homes = []
        for s in ss:
            f = fac[s["fid"]]
            ij, harm, fn = three_years(s["fid"])
            now = f.get("now") or {}
            homes.append(dict(s=s, f=f, ij=ij, harm=harm, fines=fn, overall=overall(f), hours=now.get("hours"),
                              sff=f.get("sff"), abuse=bool(f.get("abuse"))))
        # the buyer's Rhode Island homes now (homes it runs today, not the ones it is buying)
        cur = [t for t in (D.get("tenures") or []) if t["group"] == g and t["current"]]
        cur_stars = [overall(fac[t["fid"]]) for t in cur if t["fid"] in fac and overall(fac[t["fid"]]) is not None]
        made.append(dict(g=g, slug=slug(g), n=len(homes)))
        for large in (False, True):
            page = render(g, homes, tr, cur, cur_stars, ri_avg, ri_hours, natl, about.get(g), D, latest, cut, large)
            fn = os.path.join(out_dir, slug(g) + ("-large" if large else "") + ".html")
            open(fn, "w", encoding="utf-8").write(page)
    # index
    items = "".join(f'<li>{esc(m["g"])} ({plural(m["n"], "home")}): <a href="{m["slug"]}.html">one page</a>, '
                    f'<a href="{m["slug"]}-large.html">large print</a></li>' for m in made)
    open(os.path.join(out_dir, "index.html"), "w", encoding="utf-8").write(
        f'<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">'
        f'<title>Pending sales: printable summaries</title><style>{CSS}</style></head><body><div class="paper">'
        f'<div class="top"><span>{ABC}</span></div><h1>Pending nursing home sales: printable summaries</h1>'
        f'<p>One page for each buyer with an application before the Health Services Council, plus a large-print version. '
        f'Updated with each CMS release; the newest is from {month_year(latest)}.</p><ul>{items}</ul>'
        f'<p><a href="../#pending-h">Back to the sales page</a></p></div></body></html>')
    return made


def render(g, homes, tr, cur, cur_stars, ri_avg, ri_hours, natl, about, D, latest, cut, large):
    n = len(homes)
    pos = g + ("’" if g.endswith("s") else "’s")
    title = f"{g}: applications to take over {plural(n, 'Rhode Island nursing home')}"
    filed = sorted({h["s"]["date"] for h in homes})
    L = []
    L.append(f'<div class="top"><span>{ABC}</span><span class="nw">Data as of {month_year(latest)}</span></div>')
    L.append(f"<h1>{ri(esc(title))}</h1>")
    status = homes[0]["s"].get("status") if len({h["s"].get("status") for h in homes}) == 1 else None
    first = f"Filed {nice_date(filed[0])}" + (f" to {nice_date(filed[-1])}" if len(filed) > 1 else "")
    L.append(f'<p class="small">{esc(first)}. ' + (esc(status) + "." if status else "") + "</p>")
    if about:
        L.append(f"<p>{ri(esc(about))}</p>")

    # ---- the homes today
    L.append("<h2>The homes today</h2>")
    if not large:
        L.append('<table><thead><tr><th>Home</th><th class="n">Beds</th><th class="n">CMS stars</th>'
                 '<th class="n">Nurse hours per resident per day</th><th class="n">Immediate jeopardy citations</th>'
                 '<th class="n">Harm citations</th><th class="n">Fines</th></tr></thead><tbody>')
        for h in homes:
            f = h["f"]
            st = "Special Focus Facility" if h["sff"] == "SFF" else stars(h["overall"])
            notes = [x for x in ("Special Focus candidate" if h["sff"] == "SFF Candidate" else None,
                                 "CMS abuse icon" if h["abuse"] else None) if x]
            L.append(f'<tr><td><b>{esc(f["name"])}</b>, {esc(f["town"])}<br><span class="small">{esc(h["s"]["cec"])}: '
                     f'{esc(h["s"].get("who") or "")}{"; " + ", ".join(notes) if notes else ""}</span></td>'
                     f'<td class="n">{int(f["beds"]) if f.get("beds") else chr(8211)}</td><td class="n">{st}</td>'
                     f'<td class="n">{"%.2f" % h["hours"] if h["hours"] is not None else chr(8211)}</td><td class="n">{h["ij"]}</td>'
                     f'<td class="n">{h["harm"]}</td><td class="n">{money(h["fines"])}</td></tr>')
        L.append("</tbody></table>")
    else:
        for h in homes:
            f = h["f"]
            st = "Special Focus Facility (no star rating)" if h["sff"] == "SFF" else f'{stars(h["overall"])} of 5'
            L.append(f'<div class="home"><p><b>{esc(f["name"])}</b>, {esc(f["town"])}</p><ul>'
                     f'<li>Application {esc(h["s"]["cec"])}: {esc(h["s"].get("who") or "")}</li>'
                     f'<li>Beds: {int(f["beds"]) if f.get("beds") else chr(8211)}</li><li>CMS overall stars: {st}</li>'
                     + (f'<li>Nurse hours per resident per day: {h["hours"]:.2f}</li>' if h["hours"] is not None else "")
                     + f'<li>Immediate jeopardy citations: {h["ij"]}</li><li>Harm citations: {h["harm"]}</li>'
                     f'<li>Fines: {money(h["fines"])}</li>'
                     + ("<li>CMS Special Focus candidate</li>" if h["sff"] == "SFF Candidate" else "")
                     + ("<li>CMS abuse icon</li>" if h["abuse"] else "") + "</ul></div>")
    L.append(f'<p class="small">Citations and fines are from the last three years ({month_year(cut)} to {month_year(latest)}). '
             f'Harm citations include immediate jeopardy. Rhode&nbsp;Island average: {ri_hours:.2f} nurse hours per resident per day.</p>'
             if ri_hours else "")

    # ---- the buyer's record
    L.append(f"<h2>{ri(esc(pos))} record</h2><ul>")
    if tr.get("homes"):
        since = month_year(tr["first"]) if tr.get("first") else None
        L.append(f'<li>In Rhode&nbsp;Island: runs {plural(len(cur), "home")} now'
                 + (f' ({tr["homes"]} in all since {since})' if tr["homes"] != len(cur) and since else (f", since {since}" if since else "")) + ".</li>")
        if tr.get("stars_start") is not None:
            L.append(f'<li>Average CMS stars at its Rhode&nbsp;Island homes: {tr["stars_start"]:.1f} when it took over, '
                     f'{tr["stars_end"]:.1f} now or when it left.</li>')
        L.append(f'<li>While it ran them: {plural(tr["ij"], "immediate jeopardy citation")}, {plural(tr["harm"], "harm citation")}, '
                 f'and {money(tr["fines"])} in CMS fines, over {tr["years"]:g} years of running homes in all.</li>')
    else:
        L.append("<li>No Rhode&nbsp;Island nursing homes on record.</li>")
    c = tr.get("chain")
    if c:
        bits = [f'{plural(int(c.get("homes") or 0), "home")} in {plural(int(c.get("states") or 0), "state")}']
        if c.get("overall") is not None:
            bits.append(f'average {c["overall"]:.1f} stars (national average {natl.get("overall", 0):.1f})')
        if c.get("sff"):
            bits.append(plural(int(c["sff"]), "Special Focus Facility", "Special Focus Facilities"))
        if c.get("abuse_pct") is not None:
            bits.append(f'{c["abuse_pct"]:g}% of homes with an abuse icon (national {natl.get("abuse_pct", 0):g}%)')
        L.append(f'<li>Nationally, CMS lists {esc(chain_name(c["name"], g))} as a chain of ' + ", ".join(bits) + ".</li>")
    else:
        L.append("<li>CMS’s national list of nursing home chains has no entry for this buyer.</li>")
    L.append("</ul>")
    if tr.get("homes"):
        L.append('<p class="small">Counted from when the buyer began running each home: the start of a management agreement if there was one, '
                 'otherwise the sale date. CMS\u2019s files reach back to about April 2016.</p>')

    # ---- one chart
    short = g if len(g) < 26 else "Buyer"
    rows = [(f"The {plural(len(homes), 'home')} being sold", (sum(h["overall"] for h in homes if h["overall"] is not None) /
             max(1, len([h for h in homes if h["overall"] is not None]))) if homes else None, ""),
            (f"{short}\u2019s Rhode Island homes now",
             sum(cur_stars) / len(cur_stars) if cur_stars else None, ""),
            (f"{short}\u2019s homes nationally (CMS)", (c or {}).get("overall"), ""),
            ("All Rhode Island homes", ri_avg, ""),
            ("All U.S. homes (CMS)", natl.get("overall"), "")]
    ch = bar_chart(rows, large)
    if ch:
        L.append(f'<h2>Average CMS overall stars</h2><div class="chart">{ch}</div>'
                 '<p class="small">A Special Focus Facility counts as 0 stars, because CMS gives it no rating while it is in the program.</p>')

    # ---- footer
    L.append(f'<p class="small foot">Sources: CMS Care Compare '
             f'nursing home files, {month_year(latest)}; CMS Nursing Home Chain Performance Measures'
             + (f', {esc(D["chain_asof"])}' if D.get("chain_asof") else "") + '; RIDOH change in effective control applications. '
             f'More at <span class="nw">snfsales.abc-ri.org</span>. Corrections: <span class="nw">corrections@abc-ri.org</span>.</p>')
    body = "\n".join(x for x in L if x)
    other = f'{slug(g)}{"" if large else "-large"}.html'
    return (f'<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">'
            f'<title>{esc(g)}: pending sale summary{" (large print)" if large else ""}</title>'
            f'<style>{CSS}{LARGE if large else ""}</style></head><body>'
            f'<p class="screen-only"><a href="index.html">All summaries</a> · <a href="{other}">'
            f'{"Regular print" if large else "Large print"}</a> · Use your browser’s Print command to print.</p>'
            f'<div class="paper">{body}</div></body></html>')
