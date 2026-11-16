# mustafemoh.github.io

Personal website of Mustafe Abdi Mohamed (`index.html`) and an open Somalia economic data resource (`data.html`), published with GitHub Pages at https://mustafemoh.github.io.

## What's here

| Path | Contents |
|---|---|
| `index.html` | Home page: bio, experience, publications, education and contact. |
| `data.html` | The web page (tabs, charts, tables, downloads). Reads `data/somalia.json`. |
| `data/raw/` | Untouched downloads from each source (IMF, World Bank). |
| `data/raw/mof_budget.json` | Federal budget series imported from your FGS Budget Master workbook. |
| `data/manual/` | Any other own compilations (CSV), shown in the tab that lists them. |
| `data/clean/` | One tidy CSV per table, produced by the build. |
| `data/somalia.json` | Registry of every series and table; the page reads this. |
| `downloads/` | Excel (`.xlsx`) and LaTeX (`.tex`, booktabs) versions of each table. |
| `scripts/fetch_data.py` | Downloads the latest IMF and World Bank data. |
| `scripts/build.py` | Rebuilds every output from the raw and manual data. |
| `.github/workflows/update-data.yml` | Runs fetch + build on the 1st of each month. |

## Everyday use

```bash
python scripts/fetch_data.py   # refresh raw data (needs internet)
python scripts/build.py        # rebuild tables, downloads and data/somalia.json
```

Then commit and push. To preview locally: `python -m http.server`, then open http://localhost:8000/.

## Updating the federal budget data

The **Federal budget** tab is built from the `Somalia_FGS_Budget_Master` workbook (Annual_Summary and Sectors sheets). After updating the workbook:

```bash
python scripts/import_budget.py "path/to/Somalia_FGS_Budget_Master_FY2021-2027.xlsx"
python scripts/build.py
```

FY2027 (a Cabinet-approved bill not yet passed by Parliament) is left out by default. Once the 2027 budget is enacted and published, add `--include-proposed` to the import command, or update the workbook header.

Derived ratios (domestic revenue as % of GDP, execution rates, wage bill as % of domestic revenue) are calculated in `add_calc()` in `scripts/build.py`.

To add a new IMF or World Bank series, add its code to `IMF_SERIES` or `WB_SERIES` in `scripts/build.py`, then list it in the relevant tab in `TABS`.

## Each refresh, check

- `IMF_PROJ_FROM` in `scripts/build.py`: the first year the IMF treats as an estimate or projection in the current WEO vintage.
- That no non-public data has been added.

## Licence and citation

Original data remain subject to their providers' terms (IMF and World Bank data are open for reuse with attribution). Cite the original sources, and this collection as: Mohamed, M. A. (2026). *Somalia economic data*.

This is personal research and does not represent the views of any employer.
