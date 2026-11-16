"""Build the Somalia data hub from raw source files.

Reads:   data/raw/imf_weo_datamapper.json   (IMF WEO via DataMapper)
         data/raw/worldbank_wdi.json          (World Bank WDI, optional)
         data/manual/*.csv                    (your own series, optional)
Writes:  data/clean/<table>.csv               (one tidy table per topic)
         data/somalia.json                    (registry read by data.html)
         downloads/<table>.xlsx               (if openpyxl is installed)
         downloads/<table>.tex                (booktabs LaTeX tables)

Uses only the Python standard library (openpyxl is optional, for Excel).
Run from the repository root:  python scripts/build.py
"""
import csv
import json
import os
from datetime import date

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(ROOT, "data", "raw")
CLEAN = os.path.join(ROOT, "data", "clean")
MANUAL = os.path.join(ROOT, "data", "manual")
DL = os.path.join(ROOT, "downloads")

# First year shaded as "estimate / projection" for IMF WEO series.
# Check this against the WEO vintage each time you refresh the data.
IMF_PROJ_FROM = 2025

COUNTRIES = {
    "SOM": "Somalia", "KEN": "Kenya", "ETH": "Ethiopia", "UGA": "Uganda",
    "TZA": "Tanzania", "RWA": "Rwanda", "DJI": "Djibouti", "SSQ": "Sub-Saharan Africa",
}

# ---- Series definitions -------------------------------------------------
IMF_SERIES = {
    "NGDP_RPCH":   ("Real GDP growth", "annual % change"),
    "NGDPD":       ("Nominal GDP", "US$ billions"),
    "NGDPDPC":     ("GDP per capita", "current US$"),
    "PCPIPCH":     ("Inflation, average consumer prices", "annual % change"),
    "BCA_NGDPD":   ("Current account balance", "% of GDP"),
    "GGXCNL_NGDP": ("General government overall balance", "% of GDP"),
    "GGXWDG_NGDP": ("General government gross debt", "% of GDP"),
    "LP":          ("Population", "millions"),
}

WB_SERIES = {
    "BX.TRF.PWKR.CD.DT":   ("Personal remittances received", "current US$"),
    "BX.TRF.PWKR.DT.GD.ZS": ("Personal remittances received", "% of GDP"),
    "DT.ODA.ODAT.CD":      ("Net official development assistance received", "current US$"),
    "DT.ODA.ODAT.PC.ZS":   ("Net ODA received per capita", "current US$"),
    "BX.KLT.DINV.CD.WD":   ("Foreign direct investment, net inflows", "current US$"),
    "NE.EXP.GNFS.ZS":      ("Exports of goods and services", "% of GDP"),
    "NE.IMP.GNFS.ZS":      ("Imports of goods and services", "% of GDP"),
    "SP.POP.GROW":         ("Population growth", "annual %"),
    "SP.URB.TOTL.IN.ZS":   ("Urban population", "% of total"),
    "SP.POP.0014.TO.ZS":   ("Population aged 0-14", "% of total"),
    "SP.DYN.LE00.IN":      ("Life expectancy at birth", "years"),
    "SP.DYN.TFRT.IN":      ("Fertility rate", "births per woman"),
    "SI.POV.DDAY":         ("Poverty headcount at $3.00 a day (2021 PPP)", "% of population"),
    "EG.ELC.ACCS.ZS":      ("Access to electricity", "% of population"),
    "IT.CEL.SETS.P2":      ("Mobile subscriptions", "per 100 people"),
    "NV.AGR.TOTL.ZS":      ("Agriculture, value added", "% of GDP"),
}

# ---- Tabs (topics) shown on the web page ----------------------------------
# Each table lists series as (source, code, country).
TABS = [
    {"id": "growth", "title": "Growth & output",
     "intro": "Real growth, the size of the economy and income per person.",
     "tables": [
        {"id": "som_output", "title": "Somalia: output and income",
         "series": [("imf", "NGDP_RPCH", "SOM"), ("imf", "NGDPD", "SOM"), ("imf", "NGDPDPC", "SOM"), ("imf", "LP", "SOM")],
         "charts": [("imf", "NGDP_RPCH", "SOM", "bar"), ("imf", "NGDPDPC", "SOM", "line")]}]},
    {"id": "prices", "title": "Prices",
     "intro": "Consumer price inflation. Somalia is largely dollarised, so imported prices drive inflation.",
     "tables": [
        {"id": "som_prices", "title": "Somalia: inflation",
         "series": [("imf", "PCPIPCH", "SOM")],
         "charts": [("imf", "PCPIPCH", "SOM", "bar")]}]},
    {"id": "external", "title": "External sector",
     "intro": "The current account, remittances, trade and investment flows.",
     "tables": [
        {"id": "som_external", "title": "Somalia: external balances",
         "series": [("imf", "BCA_NGDPD", "SOM"), ("wb", "BX.TRF.PWKR.CD.DT", "SOM"), ("wb", "BX.TRF.PWKR.DT.GD.ZS", "SOM"),
                    ("wb", "NE.EXP.GNFS.ZS", "SOM"), ("wb", "NE.IMP.GNFS.ZS", "SOM"), ("wb", "BX.KLT.DINV.CD.WD", "SOM")],
         "charts": [("imf", "BCA_NGDPD", "SOM", "bar"), ("wb", "BX.TRF.PWKR.CD.DT", "SOM", "line")]}]},
    {"id": "fiscal", "title": "Public finance",
     "intro": "The general government overall balance as estimated by the IMF. Detailed federal budget data are in the Federal budget tab.",
     "tables": [
        {"id": "som_fiscal", "title": "Somalia: general government",
         "series": [("imf", "GGXCNL_NGDP", "SOM")],
         "charts": [("imf", "GGXCNL_NGDP", "SOM", "bar")]}]},
    {"id": "budget", "title": "Federal budget",
     "intro": "The Federal Government of Somalia's approved budgets and actual outturns, compiled from the annual Appropriation Acts and the Ministry of Finance's budget performance reports. US$ millions unless stated.",
     "tables": [
        {"id": "fgs_revenue_actual", "title": "Revenue and grants: actual outturn",
         "series": [("mof", c + "_A", "SOM") for c in ["REV_TOTAL", "REV_DOM", "REV_TAX", "TAX_INC", "TAX_GS", "TAX_TRADE", "TAX_OTH", "REV_NONTAX", "GRANTS", "GR_BS", "GR_PROJ"]],
         "charts": [("mof", ["REV_DOM_A", "GRANTS_A"], "SOM", "lines", "Domestic revenue and grants (actual)"),
                    ("mof", ["TAX_TRADE_A", "TAX_GS_A", "TAX_INC_A"], "SOM", "lines", "Tax revenue by type (actual)")]},
        {"id": "fgs_expenditure_actual", "title": "Expenditure and balance: actual outturn",
         "series": [("mof", c + "_A", "SOM") for c in ["EXP_TOTAL", "EXP_REC", "REC_WAGES", "REC_GS", "REC_TRANSF", "EXP_PROJ", "BAL"]],
         "charts": [("mof", ["EXP_REC_A", "EXP_PROJ_A"], "SOM", "lines", "Recurrent and project spending (actual)"),
                    ("mof", "BAL_A", "SOM", "bar")]},
        {"id": "fgs_budget_approved", "title": "Approved budgets",
         "series": [("mof", c + "_B", "SOM") for c in ["REV_TOTAL", "REV_DOM", "REV_TAX", "REV_NONTAX", "GRANTS", "EXP_TOTAL", "EXP_REC", "REC_WAGES", "EXP_PROJ", "BAL"]],
         "charts": []},
        {"id": "fgs_sectors", "title": "Allocation by sector (approved budget)",
         "series": [("mof", c + "_B", "SOM") for c in ["SEC_ADMIN", "SEC_SEC", "SEC_ECON", "SEC_SOC", "SEC_TOTAL"]],
         "charts": [("mof", ["SEC_ADMIN_B", "SEC_SEC_B", "SEC_ECON_B", "SEC_SOC_B"], "SOM", "lines", "Allocation by sector")]},
        {"id": "fgs_ratios", "title": "Key ratios and budget execution",
         "series": [("calc", c, "SOM") for c in ["DOMREV_GDP", "TAX_GDP", "GRANTS_GDP", "EXP_GDP", "DOM_SHARE", "WAGE_DOM", "TRADE_TAX",
                                                  "EXEC_DOM", "EXEC_GRANTS", "EXEC_EXP", "EXEC_REC", "EXEC_PROJ"]],
         "charts": [("calc", "DOMREV_GDP", "SOM", "bar"),
                    ("calc", ["EXEC_DOM", "EXEC_GRANTS", "EXEC_PROJ"], "SOM", "lines", "Budget execution (actual as % of approved)")]},
     ]},
    {"id": "aid", "title": "Aid & development finance",
     "intro": "Official development assistance received, in total and per person.",
     "tables": [
        {"id": "som_aid", "title": "Somalia: development assistance",
         "series": [("wb", "DT.ODA.ODAT.CD", "SOM"), ("wb", "DT.ODA.ODAT.PC.ZS", "SOM")],
         "charts": [("wb", "DT.ODA.ODAT.CD", "SOM", "line")]}]},
    {"id": "people", "title": "People & welfare",
     "intro": "Population, demography, poverty and access to basic services.",
     "tables": [
        {"id": "som_people", "title": "Somalia: population and welfare",
         "series": [("imf", "LP", "SOM"), ("wb", "SP.POP.GROW", "SOM"), ("wb", "SP.URB.TOTL.IN.ZS", "SOM"),
                    ("wb", "SP.POP.0014.TO.ZS", "SOM"), ("wb", "SP.DYN.LE00.IN", "SOM"), ("wb", "SP.DYN.TFRT.IN", "SOM"),
                    ("wb", "SI.POV.DDAY", "SOM"), ("wb", "EG.ELC.ACCS.ZS", "SOM"), ("wb", "IT.CEL.SETS.P2", "SOM")],
         "charts": [("imf", "LP", "SOM", "line"), ("wb", "SP.URB.TOTL.IN.ZS", "SOM", "line")]}]},
    {"id": "region", "title": "Regional comparison",
     "intro": "Somalia against its East African neighbours and the Sub-Saharan Africa average.",
     "tables": [
        {"id": "peer_growth", "title": "Real GDP growth (annual %)", "peer": "NGDP_RPCH",
         "charts": [("imf", "NGDP_RPCH", "PEERS", "multi")]},
        {"id": "peer_inflation", "title": "Inflation (annual %)", "peer": "PCPIPCH",
         "charts": [("imf", "PCPIPCH", "PEERS", "multi")]},
        {"id": "peer_gdppc", "title": "GDP per capita (current US$)", "peer": "NGDPDPC",
         "charts": [("imf", "NGDPDPC", "PEERS", "multi")]},
        {"id": "peer_balance", "title": "Overall fiscal balance (% of GDP)", "peer": "GGXCNL_NGDP", "charts": []},
        {"id": "peer_ca", "title": "Current account balance (% of GDP)", "peer": "BCA_NGDPD", "charts": []},
        {"id": "peer_debt", "title": "Government gross debt (% of GDP)", "peer": "GGXWDG_NGDP", "charts": []},
     ]},
]

SNAPSHOT = [("imf", "NGDP_RPCH", "SOM"), ("imf", "PCPIPCH", "SOM"), ("imf", "NGDPDPC", "SOM"),
            ("imf", "BCA_NGDPD", "SOM"), ("imf", "GGXCNL_NGDP", "SOM"), ("imf", "LP", "SOM"),
            ("calc", "DOMREV_GDP", "SOM"), ("mof", "REV_DOM_A", "SOM"),
            ("wb", "BX.TRF.PWKR.DT.GD.ZS", "SOM"), ("wb", "DT.ODA.ODAT.CD", "SOM")]

PEER_YEARS = list(range(2018, 2028))
SOM_FROM = 2012
SOM_TO = 2031


# ---- Loading ------------------------------------------------------------------
def load_json(name):
    path = os.path.join(RAW, name)
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


MOF_META = {}
CALC_META = {}


def series_index(imf, wb, mof=None):
    """Return {(source, code, country): {year: value}}."""
    idx = {}
    if mof:
        MOF_META.update(mof.get("meta", {}))
        for code, by_c in mof.get("series", {}).items():
            for c, vals in by_c.items():
                idx[("mof", code, c)] = {int(y): v for y, v in vals.items() if v is not None}
    if imf:
        for code, by_c in imf.items():
            if code.startswith("_"):
                continue
            for c, vals in by_c.items():
                idx[("imf", code, c)] = {int(y): v for y, v in vals.items() if v is not None}
    if wb:
        for code, by_c in wb.get("series", {}).items():
            for c, vals in by_c.items():
                idx[("wb", code, c)] = {int(y): v for y, v in vals.items() if v is not None}
    return idx


def meta(src, code):
    if src == "mof":
        m = MOF_META.get(code, {"name": code, "unit": "US$ millions"})
        return {"name": m["name"], "unit": m["unit"], "code": code,
                "source": "Ministry of Finance budget documents (author's compilation)"}
    if src == "calc":
        name, unit = CALC_META.get(code, (code, ""))
        return {"name": name, "unit": unit, "code": code,
                "source": "Author's calculation from Ministry of Finance and IMF data"}
    name, unit = (IMF_SERIES if src == "imf" else WB_SERIES).get(code, (code, ""))
    source = "IMF World Economic Outlook" if src == "imf" else "World Bank, World Development Indicators"
    return {"name": name, "unit": unit, "source": source, "code": code}


def add_calc(idx):
    """Derived budget indicators (author's calculations)."""
    g = lambda code: idx.get(("mof", code, "SOM"), {})
    gdp = idx.get(("imf", "NGDPD", "SOM"), {})  # US$ billions
    def ratio(code, name, unit, num, den, scale=100.0):
        out = {y: round(num[y] / den[y] * scale, 1) for y in num if y in den and den[y]}
        if out:
            CALC_META[code] = (name, unit)
            idx[("calc", code, "SOM")] = out
    ratio("DOMREV_GDP", "Domestic revenue (actual)", "% of GDP", g("REV_DOM_A"), gdp, 100 / 1000)
    ratio("TAX_GDP", "Tax revenue (actual)", "% of GDP", g("REV_TAX_A"), gdp, 100 / 1000)
    ratio("GRANTS_GDP", "Grants (actual)", "% of GDP", g("GRANTS_A"), gdp, 100 / 1000)
    ratio("EXP_GDP", "Total expenditure (actual)", "% of GDP", g("EXP_TOTAL_A"), gdp, 100 / 1000)
    ratio("DOM_SHARE", "Domestic revenue share of revenue and grants (actual)", "%", g("REV_DOM_A"), g("REV_TOTAL_A"))
    ratio("WAGE_DOM", "Wage bill as share of domestic revenue (actual)", "%", g("REC_WAGES_A"), g("REV_DOM_A"))
    ratio("TRADE_TAX", "Trade taxes as share of tax revenue (actual)", "%", g("TAX_TRADE_A"), g("REV_TAX_A"))
    ratio("EXEC_DOM", "Execution rate: domestic revenue", "% of budget", g("REV_DOM_A"), g("REV_DOM_B"))
    ratio("EXEC_GRANTS", "Execution rate: grants", "% of budget", g("GRANTS_A"), g("GRANTS_B"))
    ratio("EXEC_EXP", "Execution rate: total expenditure", "% of budget", g("EXP_TOTAL_A"), g("EXP_TOTAL_B"))
    ratio("EXEC_REC", "Execution rate: recurrent expenditure", "% of budget", g("EXP_REC_A"), g("EXP_REC_B"))
    ratio("EXEC_PROJ", "Execution rate: donor-funded projects", "% of budget", g("EXP_PROJ_A"), g("EXP_PROJ_B"))


# ---- Writers --------------------------------------------------------------------
def fmt(v, unit):
    if v is None:
        return ""
    if "US$" in unit and "billions" not in unit and abs(v) >= 1e6:
        return f"{v/1e6:,.1f}m"
    if abs(v) >= 1000:
        return f"{v:,.0f}"
    return f"{v:.1f}"


def write_csv(path, header, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def write_tex(path, title, header, rows, note):
    cols = "l" + "r" * (len(header) - 1)
    esc = lambda s: str(s).replace("%", r"\%").replace("&", r"\&").replace("$", r"\$").replace("_", r"\_")
    lines = [r"% Generated by scripts/build.py", r"\begin{table}[htbp]\centering",
             rf"\caption{{{esc(title)}}}", rf"\begin{{tabular}}{{{cols}}}", r"\toprule",
             " & ".join(esc(h) for h in header) + r" \\", r"\midrule"]
    for r in rows:
        lines.append(" & ".join(esc(x) for x in r) + r" \\")
    lines += [r"\bottomrule", r"\end{tabular}", rf"\par\smallskip\footnotesize {esc(note)}", r"\end{table}"]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def write_xlsx(path, title, header, rows, note):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font
    except ImportError:
        return False
    wb = Workbook()
    ws = wb.active
    ws.title = "data"
    ws.append([title]); ws["A1"].font = Font(bold=True, size=12)
    ws.append([])
    ws.append(header)
    for c in ws[3]:
        c.font = Font(bold=True)
    for r in rows:
        ws.append(r)
    ws.append([]); ws.append([note])
    ws.column_dimensions["A"].width = 48
    wb.save(path)
    return True


# ---- Build ----------------------------------------------------------------------
def main():
    for d in (CLEAN, DL, MANUAL):
        os.makedirs(d, exist_ok=True)
    imf = load_json("imf_weo_datamapper.json")
    wb = load_json("worldbank_wdi.json")
    mof = load_json("mof_budget.json")
    idx = series_index(imf, wb, mof)
    add_calc(idx)

    accessed = {"imf": (imf or {}).get("_meta", {}).get("accessed"),
                "wb": (wb or {}).get("_meta", {}).get("accessed"),
                "mof": (mof or {}).get("_meta", {}).get("accessed")}
    registry = {"title": "Somalia economic data", "built": date.today().isoformat(),
                "accessed": accessed, "imf_proj_from": IMF_PROJ_FROM,
                "countries": COUNTRIES, "series": {}, "tabs": [], "snapshot": [], "pending": [],
                "budget_sources": (mof or {}).get("_meta", {}).get("column_sources", {})}

    def add_series(key):
        src, code, c = key
        if key not in idx or not idx[key]:
            return False
        sid = f"{src}:{code}:{c}"
        if sid not in registry["series"]:
            m = meta(src, code)
            m.update({"country": c, "data": [[y, idx[key][y]] for y in sorted(idx[key])]})
            m["proj_from"] = IMF_PROJ_FROM if src == "imf" else None
            registry["series"][sid] = m
        return True

    for key in SNAPSHOT:
        if add_series(key):
            registry["snapshot"].append(f"{key[0]}:{key[1]}:{key[2]}")

    for tab in TABS:
        t = {k: tab[k] for k in ("id", "title", "intro")}
        t["tables"] = []
        for tbl in tab["tables"]:
            if "peer" in tbl:
                code = tbl["peer"]
                header = ["Country"] + [str(y) for y in PEER_YEARS]
                rows = []
                for c, cname in COUNTRIES.items():
                    s = idx.get(("imf", code, c))
                    if not s:
                        continue
                    add_series(("imf", code, c))
                    rows.append([cname] + [s.get(y) for y in PEER_YEARS])
                if not rows:
                    continue
                unit = IMF_SERIES[code][1]
                note = f"Source: IMF World Economic Outlook (DataMapper), accessed {accessed['imf']}. Values from {IMF_PROJ_FROM} are IMF estimates or projections."
                series_ids = [f"imf:{code}:{c}" for c in COUNTRIES if ("imf", code, c) in idx]
            else:
                avail = [k for k in tbl["series"] if add_series(k)]
                missing = [k for k in tbl["series"] if k not in avail]
                for k in missing:
                    registry["pending"].append({"tab": tab["id"], "code": k[1], "name": meta(k[0], k[1])["name"], "source": meta(k[0], k[1])["source"]})
                if not avail:
                    continue
                years = sorted({y for k in avail for y in idx[k] if SOM_FROM <= y <= SOM_TO})
                header = ["Indicator", "Unit"] + [str(y) for y in years]
                rows = [[meta(*k[:2])["name"], meta(*k[:2])["unit"]] + [idx[k].get(y) for y in years] for k in avail]
                srcs = sorted({meta(*k[:2])["source"] for k in avail})
                note = "Sources: " + "; ".join(srcs) + "."
                if any(k[0] == "imf" for k in avail):
                    note += f" IMF values from {IMF_PROJ_FROM} are estimates or projections."
                if any(k[0] in ("mof", "calc") for k in avail):
                    note += " Budget figures from the Appropriation Act of each year; actual outturns from later budget laws and MoF performance reports (see Sources & methods)."
                series_ids = [f"{k[0]}:{k[1]}:{k[2]}" for k in avail]
                unit = None

            # files
            write_csv(os.path.join(CLEAN, tbl["id"] + ".csv"), header, rows)
            disp_rows = [[r[0]] + ([r[1]] if unit is None else []) +
                         [fmt(v, (r[1] if unit is None else unit)) for v in r[(2 if unit is None else 1):]] for r in rows]
            write_tex(os.path.join(DL, tbl["id"] + ".tex"), tbl["title"], header, disp_rows, note)
            has_xlsx = write_xlsx(os.path.join(DL, tbl["id"] + ".xlsx"), tbl["title"], header, rows, note)

            charts = []
            for ch in tbl.get("charts", []):
                src, code, c, kind = ch[:4]
                if isinstance(code, list):
                    ids = [f"{src}:{cc}:{c}" for cc in code if f"{src}:{cc}:{c}" in registry["series"]]
                    if ids:
                        charts.append({"kind": kind, "series": ids, "title": ch[4], "unit": meta(src, code[0])["unit"]})
                    continue
                if c == "PEERS":
                    ids = [f"{src}:{code}:{cc}" for cc in COUNTRIES if f"{src}:{code}:{cc}" in registry["series"]]
                else:
                    ids = [f"{src}:{code}:{c}"] if f"{src}:{code}:{c}" in registry["series"] else []
                if ids:
                    charts.append({"kind": kind, "series": ids, "title": meta(src, code)["name"], "unit": meta(src, code)["unit"]})

            t["tables"].append({"id": tbl["id"], "title": tbl["title"], "header": header, "rows": rows,
                                "unit": unit, "note": note, "series": series_ids, "charts": charts,
                                "files": {"csv": f"data/clean/{tbl['id']}.csv", "tex": f"downloads/{tbl['id']}.tex",
                                          **({"xlsx": f"downloads/{tbl['id']}.xlsx"} if has_xlsx else {})}})

        # manual (your own) series
        for m in tab.get("manual", []):
            path = os.path.join(MANUAL, m + ".csv")
            if os.path.exists(path):
                with open(path, encoding="utf-8") as f:
                    rows = list(csv.reader(f))
                body = [r for r in rows[1:] if any(x.strip() for x in r[1:])]
                if body:
                    t["tables"].append({"id": m, "title": "Federal budget (own compilation)", "header": rows[0], "rows": body,
                                        "unit": None, "note": "Source: Ministry of Finance budget documents; author's compilation.",
                                        "series": [], "charts": [], "files": {"csv": f"data/manual/{m}.csv"}})
                else:
                    registry["pending"].append({"tab": tab["id"], "code": m, "name": "Federal budget series (your MoF compilation)", "source": "Ministry of Finance"})
        registry["tabs"].append(t)

    with open(os.path.join(ROOT, "data", "somalia.json"), "w", encoding="utf-8") as f:
        json.dump(registry, f, ensure_ascii=False, separators=(",", ":"))
    n_tables = sum(len(t["tables"]) for t in registry["tabs"])
    print(f"Built {len(registry['series'])} series, {n_tables} tables; {len(registry['pending'])} series pending.")


if __name__ == "__main__":
    main()
