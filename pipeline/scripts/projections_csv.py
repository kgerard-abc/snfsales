"""Write projected-figures.csv: each buyer's projected figures with the document and page they came from."""
import csv, json, os, sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
HEAD = ["State review (CEC) number", "Nursing home", "Buyer", "Application", "Link to application",
        "Staff table (PDF page)", "Budget table (PDF page)", "Year",
        "Projected nurse staff (full-time equivalents)", "Projected RNs (full-time equivalents)",
        "Projected net patient revenue ($)", "Projected operating profit or loss ($)",
        "Projected rent ($)", "Projected management fees ($)",
        "Limit on first-year operating costs in the state's approval ($)", "Link to Health Services Council report"]


def write(src=None, out=None):
    src = src or os.path.join(ROOT, "pipeline", "projections.json")
    out = out or os.path.join(ROOT, "projected-figures.csv")
    sales = json.load(open(src, encoding="utf-8"))["sales"]
    rows = []
    for s in sales:
        cap = s.get("cost_cap") or {}
        for i, (year, y) in enumerate(sorted(s["years"].items())):
            rows.append([s["cec"], s["home"], s["buyer"], s["document"], s["link"], s.get("staff_page", ""), s.get("budget_page", ""), year,
                         y.get("nurse_fte", ""), y.get("rn_fte", ""), y.get("revenue", ""), y.get("profit", ""), y.get("rent", ""), y.get("mgmt", ""),
                         cap.get("amount", "") if i == 0 else "", cap.get("report", "") if i == 0 else ""])
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(HEAD)
        w.writerows([["" if v is None else v for v in r] for r in rows])
    return len(rows)


if __name__ == "__main__":
    print(f"Wrote projected-figures.csv ({write(*sys.argv[1:3])} rows).")
