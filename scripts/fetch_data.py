"""Download the latest public data into data/raw/.

Sources
  IMF World Economic Outlook  -> data/raw/imf_weo_datamapper.json
  World Bank WDI              -> data/raw/worldbank_wdi.json

Standard library only. Run from the repository root:
    python scripts/fetch_data.py            # both sources
    python scripts/fetch_data.py --wb-only  # only the World Bank
If a source is unreachable, its existing raw file is kept unchanged.
"""
import json
import os
import sys
import time
import urllib.request
from datetime import date

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(ROOT, "data", "raw")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build import COUNTRIES, IMF_SERIES, WB_SERIES  # noqa: E402

UA = {"User-Agent": "somalia-data-hub/1.0 (research use)"}


def get_json(url, tries=3, timeout=60):
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers=UA)
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return json.loads(r.read().decode("utf-8"))
        except Exception as e:  # noqa: BLE001
            print(f"  attempt {i+1} failed: {e}")
            time.sleep(3 * (i + 1))
    return None


def fetch_imf():
    out = {"_meta": {"source": "IMF World Economic Outlook via IMF DataMapper API",
                     "accessed": date.today().isoformat(), "countries": list(COUNTRIES)}}
    for code in IMF_SERIES:
        print(f"IMF {code}")
        j = get_json(f"https://www.imf.org/external/datamapper/api/v1/{code}/{'/'.join(COUNTRIES)}")
        if not j:
            return None
        vals = j.get("values", {}).get(code, {})
        out[code] = {c: {y: v for y, v in s.items() if int(y) >= 1990} for c, s in vals.items() if c in COUNTRIES}
    return out


def fetch_wb():
    out = {"_meta": {"source": "World Bank World Development Indicators API",
                     "accessed": date.today().isoformat()}, "series": {}}
    for code in WB_SERIES:
        print(f"WB {code}")
        j = get_json(f"https://api.worldbank.org/v2/country/SOM/indicator/{code}?format=json&per_page=200&date=1960:{date.today().year}")
        if not j or len(j) < 2 or j[1] is None:
            print(f"  no data for {code}")
            continue
        out["series"][code] = {"SOM": {d["date"]: d["value"] for d in j[1] if d["value"] is not None}}
    return out if out["series"] else None


def save(name, obj):
    with open(os.path.join(RAW, name), "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=0)
    print(f"Saved data/raw/{name}")


if __name__ == "__main__":
    os.makedirs(RAW, exist_ok=True)
    if "--wb-only" not in sys.argv:
        imf = fetch_imf()
        save("imf_weo_datamapper.json", imf) if imf else print("IMF unreachable; kept existing file.")
    wb = fetch_wb()
    save("worldbank_wdi.json", wb) if wb else print("World Bank unreachable; kept existing file.")
