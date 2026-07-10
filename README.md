# DC Discrepancy Dashboards

Eight cross-linked dashboard pages built from three data-center workbooks:

| Source | File | Key tabs |
|--------|------|----------|
| Client / demand | `Client_Intelligence Factory Anaysis_June11_2026 (1) - v2.xlsx` | `US Shortfall Analysis`, `Power Summary`, `Compute Demand`, `AI Power_US`, `AI Power_Global` |
| Supply model | `AI-Data-Center-Model-CLIENT-April-7-noSKU.xlsx` | `NA Data Center Supply`, `Hyperscalers & Neoclouds`, `Power Requirements per Chip`, `AI CPU Demand calculations` |
| MS Nvidia server model | `Request_NV AI Server Model_052126 2.xlsx` | `ODM share analysis`, `GB200-300 Tracker`, `GPU Roadmap`, `Nvidia Rack Layout` |

**Page 1 (`dashboard.html`):** does the facility-level supply model's US *Planned + UC* capacity match
the client's *Power Shortfall Before New "Time to Power" Solutions* (**37.71 GW**)?

**Page 2 (`chips_dashboard.html`):** chip demand (which SKUs, how many), power per chip, the derived
chip→data-center power bridge, and what kind of capacity it lands in (self-build vs leased vs neocloud).

**Page 3 (`compute_dashboard.html`):** compute (ExaFLOPs) and energy (TWh) over time — Nvidia by
generation, derived GPU-vs-ASIC splits (labeled assumptions, validated against the client's own
aggregate), the training wall, and AI power demand by company and customer type.

**Page 4 (`synthesis_dashboard.html`):** the synthesis — verdicts with evidence on (a) whether delays
damp chip demand (they defer it: ~40 forecast vintages show "slower now, bigger later"; median 9-month
modeled slippage; constrained demand runs ~13% below unconstrained in 2028) and (b) whether chip supply
outpaces buildout (yes — the model's own NA balance runs a cumulative deficit every year 2023–2030),
plus a dedicated section on where the three spreadsheets disagree (~64 GW spread on the same question).

**Page 5 (`geo_dashboard.html`):** the buildout made geographic and deployment-typed. An **offline US map**
(facility Lat/Long plotted with vendored d3-geo + us-atlas, no online tiles) with circles sized by GW of capacity
coming online 2026–28, color-toggled by construction status or by deployment type. A **second map** overlays the
EIA natural-gas pipeline network and colors each UC/planned facility by distance to the nearest pipeline
(≤5 / 5–25 / >25 mi) — a proxy for how readily a site could add on-site gas generation (the shortfall analysis's
"Nat Gas Turbines" solution); ~66% of 2026–28 capacity sits within 5 mi of a pipeline. **Both maps
are pan/zoomable** (scroll or `+`/`−`/`Reset`; markers keep a constant screen size, borders stay crisp).
Then two charts split by
**Neocloud (pink) / Leasing (blue) / Self-built (lavender)**: **GW of Chips** (demand — magnitude from the client
`AI Power_US` tab, mix from the supply model's `Hyperscalers & Neoclouds` tab; a *derived cross-source blend*) vs
**Datacenter capacity** (supply — `NA Data Center Supply` facilities). Toggles: quarterly/annual and
absolute / net-additions / YoY %. Deployment class per facility is a labeled heuristic (neocloud operator/tenant/
GPU-cloud tag → neocloud; hyperscaler-type → self-built; colocation → leasing).

**Page 6 (`flows_dashboard.html`):** four **D3 v7** views of the same 2026–28 facility cohort (all client-side
from `dashboard_data.json`, no new extraction): (1) a **Sankey** of GW flowing Company → Deployment type →
Pipeline-proximity bucket (hover to isolate); (2) a **zoomable circle-pack** of State → Company (click to zoom);
(3) a **streamgraph** of capacity by deployment type over time (annual/quarterly toggle); (4) a **beeswarm** of
facilities placed by go-live year, radius ∝ GW, colored by deployment type.

**Page 7 (`chippower_dashboard.html`):** the **GW and TWh needed by the chips themselves**, drawn from the
*demand-side* workbooks only (client + MS Nvidia, not the supply model): (1) global chip **server power by vendor
family** (Nvidia / AMD / Intel / Google TPU / AWS / custom), 2021–2028, from `AI Power_Global`
(volume × per-server Max Input Power); (2) **cumulative chip DC power capacity (GW), global vs US** (122 GW global /
69 GW US by 2028); (3) **energy (TWh)** — actual installed-fleet vs incremental new-chip load, global & US
(802 / 455 TWh in 2028); (4) the per-chip driver — **watts per accelerator** (`GPU Roadmap`) and **kW per NVL72 rack**
(`Nvidia Rack Layout`).

**Page 8 (`delayed_dashboard.html`):** the **cancelled / delayed** cut, scoped to sites whose *original* online
date (`NA Data Center Supply` col BY, "Start of operations") fell in **2026–2028** — plotted on the same **offline
zoomable US map** the Geography page uses (d3-geo AlbersUSA + us-atlas, with the faint EIA gas-pipeline network as
basemap). The model has *no* status flag, so status is derived from the model's *Live Date* (col CC): **cancelled**
= a far-future sentinel (≥2034) or blank while capacity is still planned → drawn **red**; **delayed** = the live
date slips >3 months past the original → drawn on a **yellow→red gradient** by the length of the slip. Because most
real cancellations were killed *pre-construction* by local opposition (so the model never carried them), **externally
-reported, cited** cancellations/delays are overlaid as **ringed markers** (hover for the source) — including a
cluster in **Indiana** (Google Franklin Township, the New Carlisle rejection, QTS Porter County, Agincourt
Valparaiso, Surge/Hancock County) plus PW Digital Gateway VA, Microsoft Caledonia WI, Tract Mooresville NC, AWS
Becker MN, Crusoe "Project Jade" WY, Oracle/OpenAI Abilene TX, and STACK/Oracle "Project Jupiter" NM (→2029). In the
2026–28 cohort the model shows **~140 GW delayed** (median 10.6-month slip) and **~27 GW cancelled/indefinite**;
that GW is translated to **energy** at the model's own **~8.4 TWh per GW-year** (8760h × 75% utilization × PUE 1.28
— the "GW-hours → TWh" / **Chip TWh/GW** intensity used on pages 3 & 7): ~1,177 TWh/yr deferred + ~229 TWh/yr lost.
The map filters by status (delayed / cancelled / cited) and by **phase — Under Construction vs Planned** (col BV vs
CD). A **shortfall-effect** card asks whether any of this contests the client waterfall's "under construction" line
(demand 67.6 − DC UC 14.9 − grid 15.0 = **37.7 GW** shortfall) — and the answer is **no**: of the UC originally due
2026–28, **19.0 GW still lands by 2028** and only **0.66 GW slips past** it (none cancelled), and the model shows
**26.3 GW** of US UC total, *more* than the 14.9 GW credited. The slip is almost entirely in **Planned (paper)
capacity** — **92.0 GW pushed past 2028 + 27.2 GW cancelled** — which the waterfall never counted, so the 37.7 GW
shortfall stands; what erodes is the supply model's apparent *surplus* over it (page 1's ~6× "coverage" from a
228 GW planned pipeline). Also a delay-length histogram, a **cited table** with source
links, and macro pullbacks (Baird's ~25 US cancellations in 2025 — Indiana ~8 / ~5 GW; Microsoft ~200 MW leases /
TD Cowen; AWS / Wells Fargo; Data Center Watch's $130B/75+ blocked in Q1 2026), with a SemiAnalysis caveat on the
disputed "half of 2026 is cancelled" claim.

## Run

```bash
pip install openpyxl      # one-time
python build_dashboard.py
```

This reads the three workbooks and writes:

- **`dashboard.html`** — supply vs shortfall (page 1).
- **`chips_dashboard.html`** — chips, power & buildout (page 2).
- **`compute_dashboard.html`** — compute & energy over time (page 3).
- **`synthesis_dashboard.html`** — delays, the race, and cross-file disagreements (page 4).
- **`geo_dashboard.html`** — offline (zoomable) maps + GW-of-chips vs datacenter-capacity by deployment type (page 5).
- **`flows_dashboard.html`** — four D3 views: Sankey, circle-pack, streamgraph, beeswarm (page 6).
- **`chippower_dashboard.html`** — GW & TWh needed by the chips, global & US, from the client + MS files (page 7).
- **`delayed_dashboard.html`** — map of cancelled / delayed sites (original online 2026–28) and their GW / TWh impact, with cited sources (page 8).
- **`dashboard_data.json`** — all extracted + reconciled figures.

All pages are self-contained and work offline (Chart.js inlined; pages 5–6 also inline the full
D3 v7 bundle; page 5 adds topojson-client, a us-atlas states TopoJSON and the EIA gas-pipeline
network; page 6 adds d3-sankey). A built-in self-check aborts if any anchor cell reference drifts
from the expected values.

## What page 1 shows

1. **Shortfall waterfall** — demand → less under-construction → less grid access → shortfall → "Time to Power" solutions → net (Low/Midpoint/High toggle).
2. **Coverage** — model US Under-Construction and Planned+UC vs the 37.71 GW "needed" line.
3. **YoY capacity** — US installed capacity 2017–2032, hyperscaler vs colocation, 2026–28 band.
4. **Assumption gaps** — model vs client on under-construction and 2026–28 additions.
5. **Pipeline concentration** — top US states and companies by Planned+UC.

## What page 2 shows

- **A. Chip demand (global):** shipments by family 2023–27 (Hopper→Blackwell→Rubin); NVL72 rack
  shipments by SKU + monthly ramp; Nvidia GPUs vs custom ASICs; cumulative ExaFLOPs sold vs needed.
- **B. Power:** watts per chip by generation from two independent sources (they diverge at VR300/R300:
  4,600W client vs 3,600W MS); a **derived** bridge — racks × measured rack-kW × PUE 1.3 → GW — against
  US buildout; US DC power trajectory 2023–28.
- **C. Data center type (NA/US):** self-build vs leasing vs contracted-neocloud per hyperscaler;
  neocloud + AI-lab (OpenAI/Anthropic/xAI/MSI/GDM) capacity ramps; rack allocations by customer
  segment; US Planned+UC by tenant tag (~76% untagged).

## Key findings (as built)

- Client assumes **14.85 GW** under construction; the model shows **26.3 GW** (~77% higher).
- Client US demand 2026–28 = **67.6 GW**; the model adds **94.1 GW** of US capacity over the same window.
- The **37.71 GW** shortfall is dwarfed by the model's US planned pipeline (**~202 GW**, ~6× coverage) —
  the real question is how much of that "planned" capacity actually gets built and powered.
- A year of NVL72 racks alone implies **5 / 15 / 29 GW** of facility power (2025/26/27, PUE 1.3) —
  by 2027 rivaling the client's *entire* US additions (22.5 GW).
- Frontier labs' identified data centers total **~43 GW by 2028** (OpenAI 14.7, GDM 9.8, Anthropic 8.9) —
  a bigger new capacity class than all neoclouds combined (10.5 GW).
- Oracle is ~99% leased in NA by 2028; Microsoft layers **2.9 GW of contracted neocloud** capacity on
  top of self-builds.

## Caveats surfaced in the dashboard

- **Vintage:** model is dated April 7; shortfall analysis is June 11, 2026.
- **Scope:** supply is filtered to US-only building-level facilities.
- **Sectors:** the supply model classifies only *hyperscaler* vs *colocation* — it has no "bitcoin miner"
  or "neo cloud" type. In the shortfall tab, bitcoin sites are a *supply solution* (Solution #4), not demand.

## Files

| File | Purpose |
|------|---------|
| `build_dashboard.py` | Extraction, derivation, reconciliation, HTML generation (all pages) |
| `dashboard_template.html` | Page 1 template (shortfall charts) |
| `chips_template.html` | Page 2 template (chips & buildout charts) |
| `compute_template.html` | Page 3 template (compute & energy charts) |
| `synthesis_template.html` | Page 4 template (verdicts, delays, discrepancies) |
| `geo_template.html` | Page 5 template (offline zoomable maps + by-deployment-type charts) |
| `flows_template.html` | Page 6 template (Sankey, circle-pack, streamgraph, beeswarm) |
| `chippower_template.html` | Page 7 template (chip GW/TWh demand: global & US, per-chip driver) |
| `delayed_template.html` | Page 8 template (cancelled/delayed US map — original-2026–28 sites, GW & TWh impact, cited sources) |
| `vendor/chart.umd.js` | Vendored Chart.js v4 (inlined for offline use) |
| `vendor/d3.min.js` | Vendored full D3 v7 (geo, zoom, force, hierarchy…; inlined into pages 5–6) |
| `vendor/d3-sankey.min.js` | Vendored d3-sankey plugin (inlined into page 6) |
| `vendor/topojson-client.min.js`, `vendor/us-states-10m.json` | topojson + US states TopoJSON (inlined into page 5) |
| `vendor/us-gas-pipelines.json` | EIA natural-gas pipeline network, simplified to 2dp (inlined into page 5; also used to compute per-facility pipeline distance) |
| `dashboard.html` / `chips_dashboard.html` / `compute_dashboard.html` / `synthesis_dashboard.html` / `geo_dashboard.html` / `flows_dashboard.html` / `chippower_dashboard.html` / `delayed_dashboard.html` / `dashboard_data.json` | Generated output |
