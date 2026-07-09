# DC Discrepancy Dashboards

Five cross-linked dashboard pages built from three data-center workbooks:

| Source | File | Key tabs |
|--------|------|----------|
| Client / demand | `Client_Intelligence Factory Anaysis_June11_2026 (1) - v2.xlsx` | `US Shortfall Analysis`, `Power Summary`, `Compute Demand`, `AI Power_US` |
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
coming online 2026–28, color-toggled by construction status or by deployment type. Then two charts split by
**Neocloud (pink) / Leasing (blue) / Self-built (lavender)**: **GW of Chips** (demand — magnitude from the client
`AI Power_US` tab, mix from the supply model's `Hyperscalers & Neoclouds` tab; a *derived cross-source blend*) vs
**Datacenter capacity** (supply — `NA Data Center Supply` facilities). Toggles: quarterly/annual and
absolute / net-additions / YoY %. Deployment class per facility is a labeled heuristic (neocloud operator/tenant/
GPU-cloud tag → neocloud; hyperscaler-type → self-built; colocation → leasing).

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
- **`geo_dashboard.html`** — offline map + GW-of-chips vs datacenter-capacity by deployment type (page 5).
- **`dashboard_data.json`** — all extracted + reconciled figures.

All pages are self-contained (Chart.js inlined; page 5 also inlines d3-geo, topojson-client and a
us-atlas states TopoJSON) and work offline. A built-in self-check aborts if any anchor cell reference
drifts from the expected values.

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
| `geo_template.html` | Page 5 template (offline map + by-deployment-type charts) |
| `vendor/chart.umd.js` | Vendored Chart.js v4 (inlined for offline use) |
| `vendor/d3-array.min.js`, `vendor/d3-geo.min.js`, `vendor/topojson-client.min.js`, `vendor/us-states-10m.json` | Vendored map libs + US states TopoJSON (inlined into page 5) |
| `dashboard.html` / `chips_dashboard.html` / `compute_dashboard.html` / `synthesis_dashboard.html` / `geo_dashboard.html` / `dashboard_data.json` | Generated output |
