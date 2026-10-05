# IEX BESS Arbitrage

Solar + BESS arbitrage model for P7 (Tumkur), P8 (Bidar) and P12 (Kunnoor 4). The BESS is sized from CTU night
connectivity, charged from own solar or from IEX, and sold on IEX (DAM / GDAM / RTM, 15-minute blocks).

## Outputs (project folder)

| File | What it is |
|---|---|
| `index.html` | Home page for Vercel — identical to `bess_arbitrage_full_standalone.html` |
| `bess_arbitrage_full_standalone.html` | Full two-tab model (Solar + BESS, and 2 cycles/day CATL 12,000 IEX-only). Double-click to open. |
| `solar_bess_iex_arbitrage.html` | Solar + BESS case only |
| `two_cycle_catl_iex.html` | 2 cycles/day CATL IEX-only case only |
| `bess_arbitrage_model.html` | Same as the full model, in the format used for the published claude.ai page |
| `BESS_IEX_Financial_Model.xlsx` | Excel model: Assumptions sheet + live 25-year model sheets (also the template for the app's Generate button) |
| `xlsx_map.json` | Which Excel cell holds each input; used by the Generate button |

Inputs: `DAM_/GDAM_/RTM_15min_2025-09-01_to_2026-09-23.xlsx` (IEX MCP, 15-minute).

## Generate Excel from the app

On the hosted site the sidebar has **Financial model (Excel)**: *Generate from current inputs* downloads the
workbook filled with everything set in the app — plant and connectivity, duration, charging case, markets,
windows, capex, technical and finance inputs, booster and equity-cost switches, CATL inputs and windows, and
any prices dragged on the chart. Own-solar dispatch is recomputed in the browser for the current inputs.
The Cover shows headline results immediately; Excel, Numbers and Google Sheets recalculate the rest on open.
*Download the default model* gets the file as built. (Hidden when a page is opened as a local file.)

## Deploy (Vercel)

The repo is a static site with no build step. `index.html` is the full two-tab model.

- Import the GitHub repo in Vercel → **Framework Preset: Other**, leave Build Command and Output Directory empty, Root Directory `./`.
- Pages: `/` (full model) · `/solar` (Solar + BESS only) · `/two-cycle` (CATL 2-cycle only) · `/model` (same as `/`).
- `.vercelignore` keeps `source/` and the raw IEX price files out of the deployment; the Excel model is downloadable at `/BESS_IEX_Financial_Model.xlsx`.
- After any change: run `python3 source/build.py`, commit and push; Vercel redeploys automatically.

## Access

Access is managed by Vercel, not by the app. In Vercel → Project → Settings → **Deployment Protection**, turn on
**Vercel Authentication** for the deployments you want protected (including Production), then invite users to the
Vercel team/project. Invited users sign in with their Vercel account to open the site; anyone else is refused.

## Source (`source/`)

| File | Role |
|---|---|
| `template.html` | The model: UI, calculations and charts. **Edit this** to change the model or its defaults. |
| `standalone.js` | Extra script that turns the full model into the CATL-only standalone |
| `extract_prices.py` | Reads the IEX xlsx files → `prices.json` |
| `build.py` | `template.html` + `prices.json` → the five HTML files (incl. `index.html`) |
| `export_dispatch.js` | Runs the model's own JS → `export.json` (own-solar dispatch + reference results for Excel) |
| `mkxlsx.py` | `prices.json` + `export.json` → `BESS_IEX_Financial_Model.xlsx` and `xlsx_map.json` |
| `add_cached_values.py` | Stores calculated values in the Excel so previewers (Quick Look, Drive, phones) show numbers |
| `evalx.py` | Optional check: recalculates the Excel and reports formula errors |

## Rebuild everything

Run from the project folder:

```bash
python3 source/extract_prices.py      # only if the IEX price files change
python3 source/build.py
node source/export_dispatch.js
python3.12 source/mkxlsx.py           # needs openpyxl
python3.12 source/add_cached_values.py   # needs: pip install formulas
```

`export_dispatch.js` and `mkxlsx.py` contain their own copy of the default inputs. If you change defaults in
`template.html`, update the `base` object in `export_dispatch.js` and the values in `mkxlsx.py` to match.

After `build.py`, republish `bess_arbitrage_model.html` (and the two standalones) to keep the claude.ai pages current.
