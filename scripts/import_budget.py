"""Import federal budget data from the FGS Budget Master workbook.

Usage (from the repository root):
    python scripts/import_budget.py "path/to/Somalia_FGS_Budget_Master_FY2021-2027.xlsx"

Reads the Annual_Summary and Sectors sheets and writes data/raw/mof_budget.json,
which scripts/build.py turns into the "Federal budget" tab.

Only figures from published documents should go on the public site. FY2027 comes
from a Cabinet-approved bill that Parliament has not yet passed, so it is left out
unless you pass --include-proposed.
"""
import json
import os
import sys
from datetime import date

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "data", "raw", "mof_budget.json")

# Workbook label -> (code, display name)
LINES = {
    "Total revenue & grants": ("REV_TOTAL", "Total revenue and grants"),
    "Domestic revenue": ("REV_DOM", "Domestic revenue"),
    "Tax revenue": ("REV_TAX", "Tax revenue"),
    "Taxes on income & profits": ("TAX_INC", "Taxes on income and profits"),
    "Taxes on goods & services": ("TAX_GS", "Taxes on goods and services"),
    "Taxes on international trade": ("TAX_TRADE", "Taxes on international trade"),
    "Other taxes": ("TAX_OTH", "Other taxes"),
    "Non-tax revenue": ("REV_NONTAX", "Non-tax revenue"),
    "Grants (all donor)": ("GRANTS", "Grants (all donors)"),
    "Budget support": ("GR_BS", "Budget support grants"),
    "Project grants": ("GR_PROJ", "Project grants"),
    "Total expenditure": ("EXP_TOTAL", "Total expenditure"),
    "(A) Recurrent / Government-funded": ("EXP_REC", "Recurrent (government-funded) expenditure"),
    "(B) Donor-funded projects": ("EXP_PROJ", "Donor-funded project expenditure"),
    "Overall balance (revenue & grants - expenditure)": ("BAL", "Overall balance"),
}
# Recurrent sub-lines appear twice (recurrent and projects); we take the first (recurrent) block.
REC_LINES = {
    "Compensation of employees": ("REC_WAGES", "Compensation of employees (recurrent)"),
    "Use of goods & services": ("REC_GS", "Use of goods and services (recurrent)"),
    "Grants / transfers (incl. FMS)": ("REC_TRANSF", "Transfers incl. to member states (recurrent)"),
}
SECTORS = {
    "Administration": ("SEC_ADMIN", "Administration"),
    "Defence & Security": ("SEC_SEC", "Defence and security"),
    "Economic Services": ("SEC_ECON", "Economic services"),
    "Social Services": ("SEC_SOC", "Social services"),
    "Total appropriation": ("SEC_TOTAL", "Total appropriation"),
}


def main(path, include_proposed):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    ws = wb["Annual_Summary"]
    rows = list(ws.iter_rows(values_only=True))
    hdr = next(r for r in rows if r and r[0] == "Line item")
    hdr_i = rows.index(hdr)
    sub = rows[hdr_i + 1]
    cols = []  # (col index, year, kind)
    year = None
    for i, (h, s) in enumerate(zip(hdr, sub)):
        if isinstance(h, str) and h.startswith("FY"):
            year = int(h[2:6])
            proposed = "proposed" in h.lower()
        if s in ("Budget", "Actual") and year:
            cols.append((i, year, "B" if s == "Budget" else "A", proposed))
    src_row = next((r for r in rows if r and r[0] == "Source of each column"), None)

    series, meta, sources = {}, {}, {}
    seen_rec = set()

    def put(code, name, kind, yr, v, src):
        key = f"{code}_{kind}"
        meta[key] = {"name": name + (" (approved budget)" if kind == "B" else " (actual)"), "unit": "US$ millions"}
        series.setdefault(key, {"SOM": {}})["SOM"][str(yr)] = round(float(v), 2)
        if src:
            sources[f"{yr} {'budget' if kind == 'B' else 'actual'}"] = src

    in_rec = False
    for r in rows[hdr_i + 2:]:
        label = (r[0] or "").strip() if isinstance(r[0], str) else None
        if not label:
            continue
        if label.startswith("(A) Recurrent"):
            in_rec = True
        if label.startswith("(B) Donor"):
            in_rec = False
        target = LINES.get(label) or (REC_LINES.get(label) if in_rec and label not in seen_rec else None)
        if label in REC_LINES and in_rec:
            seen_rec.add(label)
        if label.startswith("CHECKS"):
            break
        if not target:
            continue
        code, name = target
        for ci, yr, kind, proposed in cols:
            if proposed and not include_proposed:
                continue
            v = r[ci]
            if v is None or (kind == "A" and v == 0):  # 0 = placeholder for missing actuals
                continue
            put(code, name, kind, yr, v, src_row[ci] if src_row else None)

    ws2 = wb["Sectors"]
    srows = list(ws2.iter_rows(values_only=True))
    shdr = next(r for r in srows if r and r[0] == "Sector")
    for r in srows[srows.index(shdr) + 1:]:
        if r[0] in SECTORS:
            code, name = SECTORS[r[0]]
            for i, h in enumerate(shdr[1:], start=1):
                if not h or r[i] is None:
                    continue
                yr = int(str(h)[2:6])
                if "*" in str(h) and not include_proposed:
                    continue
                put(code, name, "B", yr, r[i], None)
        if r[0] and str(r[0]).startswith("Check"):
            break

    out = {"_meta": {"source": "Ministry of Finance, Federal Government of Somalia: Appropriation Acts and budget performance reports; author's compilation",
                     "workbook": os.path.basename(path), "accessed": date.today().isoformat(),
                     "includes_proposed": include_proposed, "unit": "US$ millions",
                     "column_sources": sources},
           "meta": meta, "series": series}
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=0)
    print(f"Wrote {OUT}: {len(series)} series" + (" (FY2027 proposed included)" if include_proposed else " (FY2027 proposed left out)"))


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        sys.exit(__doc__)
    main(args[0], "--include-proposed" in sys.argv)
