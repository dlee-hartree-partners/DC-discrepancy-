#!/usr/bin/env python3
"""
Build a standalone HTML dashboard visualizing the discrepancy between:

  * Client_Intelligence Factory Anaysis_June11_2026 (1) - v2.xlsx
        -> tab "US Shortfall Analysis"   (demand / "needed" side)
  * AI-Data-Center-Model-CLIENT-April-7-noSKU.xlsx
        -> tab "NA Data Center Supply"   (facility-level supply side)

The central question: does the supply model's US "Planned + UC" capacity match the
"Power Shortfall Before New 'Time to Power' Solutions" (37.71 GW) from the client's
shortfall analysis?

Run:  python build_dashboard.py
Writes: dashboard_data.json  and  dashboard.html  (self-contained, works offline).
"""

import json
import os

import openpyxl
from openpyxl.utils import column_index_from_string as _cidx

HERE = os.path.dirname(os.path.abspath(__file__))

CLIENT_FILE = os.path.join(HERE, "Client_Intelligence Factory Anaysis_June11_2026 (1) - v2.xlsx")
MODEL_FILE = os.path.join(HERE, "AI-Data-Center-Model-CLIENT-April-7-noSKU.xlsx")
NV_FILE = os.path.join(HERE, "Request_NV AI Server Model_052126 2.xlsx")
VENDOR_CHARTJS = os.path.join(HERE, "vendor", "chart.umd.js")
TEMPLATE_FILE = os.path.join(HERE, "dashboard_template.html")
CHIPS_TEMPLATE_FILE = os.path.join(HERE, "chips_template.html")
COMPUTE_TEMPLATE_FILE = os.path.join(HERE, "compute_template.html")
SYNTHESIS_TEMPLATE_FILE = os.path.join(HERE, "synthesis_template.html")

OUT_JSON = os.path.join(HERE, "dashboard_data.json")
OUT_HTML = os.path.join(HERE, "dashboard.html")
OUT_CHIPS_HTML = os.path.join(HERE, "chips_dashboard.html")
OUT_COMPUTE_HTML = os.path.join(HERE, "compute_dashboard.html")
OUT_SYNTHESIS_HTML = os.path.join(HERE, "synthesis_dashboard.html")


def col(letter):
    """0-based column index for a spreadsheet column letter."""
    return _cidx(letter) - 1


def num(x):
    return float(x) if isinstance(x, (int, float)) else 0.0


# --------------------------------------------------------------------------- #
# Extraction: US Shortfall Analysis (Client file)
# --------------------------------------------------------------------------- #
def extract_shortfall(wb):
    ws = wb["US Shortfall Analysis"]
    # Random cell access is fine here (tiny tab); pull into a grid first.
    grid = {}
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=20, values_only=True), start=1):
        for c, v in enumerate(row):
            grid[(r, c)] = v

    def g(letter, row):
        return num(grid.get((row, col(letter))))

    solutions = []
    labels = {
        10: "Nat Gas Turbines",
        11: "Bloom Energy Fuel Cells",
        12: "Site DC at Op'l Nuclear Plant",
        13: "Convert Bitcoin Sites",
    }
    for r, name in labels.items():
        solutions.append({
            "name": name,
            "low": g("C", r),
            "mid": g("D", r),
            "high": g("E", r),
            "prob": g("F", r),
        })

    data = {
        "demand_2026_28": g("C", 3),                 # 67.56
        "less_under_construction": g("C", 4),        # -14.85
        "less_grid_access": g("C", 5),               # -15.0
        "shortfall_before_solutions": g("C", 6),     # 37.71  <-- the "needed" target
        "solutions": solutions,
        "weighted_solutions": {"low": g("C", 15), "mid": g("D", 15), "high": g("E", 15)},
        "net_shortfall": {"low": g("C", 16), "mid": g("D", 16), "high": g("E", 16)},
        "shortfall_pct": {"low": g("C", 17), "mid": g("D", 17), "high": g("E", 17)},
    }
    return data


# --------------------------------------------------------------------------- #
# Extraction: NA Data Center Supply (Model file)
# --------------------------------------------------------------------------- #
YEARS = list(range(2017, 2033))            # 2017..2032
YEAR_COLS = {yr: col(chr(ord("K") + i)) for i, yr in enumerate(YEARS)}  # K..Z


def _norm_type(t):
    if not isinstance(t, str):
        return "other"
    s = t.strip().lower()
    if "hyper" in s:
        return "hyperscaler"
    if "colo" in s:
        return "colocation"
    return "other"


# Tenant classification for the ~10% of US rows tagged with an estimated
# tenant (col DB) or end user (col CZ).  Names normalized to lower-case.
NEOCLOUD_TENANTS = {
    "coreweave", "lambda", "fluidstack", "nebius", "tensorwave", "nscale",
    "iren cloud", "voltage park", "crusoe cloud", "vultr", "togetherai",
    "core42", "cerebras", "bitdeer cloud", "megaspeed", "taiga cloud",
}
HYPERSCALER_TENANTS = {"microsoft", "meta", "google", "aws", "oracle", "apple", "nvidia"}
AI_LAB_USERS = {
    "openai", "openai training", "anthropic", "anthropic training", "gdm",
    "msi", "mai", "xai", "xai inference", "poolside", "bytedance",
}


def _tenant_class(end_user, tenant):
    """Classify a leaf row by its Mosaic-theory tags; None if untagged."""
    t = tenant.strip().lower() if isinstance(tenant, str) else ""
    u = end_user.strip().lower() if isinstance(end_user, str) else ""
    if t in NEOCLOUD_TENANTS:
        return "neocloud_leased"
    if t in HYPERSCALER_TENANTS:
        return "hyperscaler_leased"
    if u in AI_LAB_USERS:
        return "ai_lab_tagged"
    if t or u:
        return "other_tagged"
    return None


def extract_supply(wb):
    ws = wb["NA Data Center Supply"]

    cA, cBV, cCD, cCE = col("A"), col("BV"), col("CD"), col("CE")
    cType, cCountry, cState, cCompany = col("CX"), col("CT"), col("CP"), col("CV")
    cEndUser, cTenant = col("CZ"), col("DB")
    c2025, c2028 = YEAR_COLS[2025], YEAR_COLS[2028]
    maxcol = max(cA, cBV, cCD, cCE, cType, cCountry, cState, cCompany,
                 cEndUser, cTenant, c2028, *YEAR_COLS.values())

    tot = {"uc": 0.0, "planned": 0.0, "pluc": 0.0, "delta_25_28": 0.0}
    by_type = {"hyperscaler": dict(uc=0.0, planned=0.0, pluc=0.0),
               "colocation": dict(uc=0.0, planned=0.0, pluc=0.0),
               "other": dict(uc=0.0, planned=0.0, pluc=0.0)}
    # per-year capacity (MW) split by type, US-only
    year_by_type = {t: {yr: 0.0 for yr in YEARS} for t in ("hyperscaler", "colocation", "other")}
    by_state, by_company = {}, {}
    tenant_gw = {"neocloud_leased": 0.0, "hyperscaler_leased": 0.0,
                 "ai_lab_tagged": 0.0, "other_tagged": 0.0, "untagged": 0.0}
    n_leaf = 0

    for row in ws.iter_rows(min_row=6, values_only=True):
        if len(row) <= maxcol:
            continue
        a = row[cA]
        # leaf = individual building, identified by a UUID-style token in column A
        if not (isinstance(a, str) and len(a) >= 8 and "-" in a):
            continue
        country = row[cCountry]
        if not (isinstance(country, str) and country.strip().upper() in ("USA", "US", "UNITED STATES")):
            continue

        n_leaf += 1
        uc, planned, pluc = num(row[cBV]), num(row[cCD]), num(row[cCE])
        t = _norm_type(row[cType])
        tot["uc"] += uc
        tot["planned"] += planned
        tot["pluc"] += pluc
        tot["delta_25_28"] += num(row[c2028]) - num(row[c2025])
        by_type[t]["uc"] += uc
        by_type[t]["planned"] += planned
        by_type[t]["pluc"] += pluc
        for yr in YEARS:
            year_by_type[t][yr] += num(row[YEAR_COLS[yr]])

        tc = _tenant_class(row[cEndUser], row[cTenant])
        tenant_gw[tc if tc else "untagged"] += pluc

        st = row[cState]
        if isinstance(st, str) and st.strip():
            by_state[st.strip()] = by_state.get(st.strip(), 0.0) + pluc
        co = row[cCompany]
        if isinstance(co, str) and co.strip():
            by_company[co.strip()] = by_company.get(co.strip(), 0.0) + pluc

    def gw(v):
        return round(v / 1000.0, 3)

    def top(d, n=10):
        items = sorted(d.items(), key=lambda kv: -kv[1])[:n]
        return [{"name": k, "gw": gw(v)} for k, v in items]

    year_total = {yr: sum(year_by_type[t][yr] for t in year_by_type) for yr in YEARS}

    data = {
        "n_leaf": n_leaf,
        "totals_gw": {k: gw(v) for k, v in tot.items()},
        "tenant_gw": {k: gw(v) for k, v in tenant_gw.items()},
        "by_type_gw": {t: {k: gw(v) for k, v in d.items()} for t, d in by_type.items()},
        "years": YEARS,
        "year_series_gw": {
            "hyperscaler": [gw(year_by_type["hyperscaler"][yr]) for yr in YEARS],
            "colocation": [gw(year_by_type["colocation"][yr]) for yr in YEARS],
            "other": [gw(year_by_type["other"][yr]) for yr in YEARS],
            "total": [gw(year_total[yr]) for yr in YEARS],
        },
        "top_states": top(by_state),
        "top_companies": top(by_company),
    }
    return data


# --------------------------------------------------------------------------- #
# Extraction: chips & buildout (all three workbooks)
# --------------------------------------------------------------------------- #
def _grid(ws, max_row, max_col=60):
    """Read a bounded rectangle into a {(row, col0): value} dict (1-based rows)."""
    g = {}
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=max_row, values_only=True), start=1):
        for c, v in enumerate(row[:max_col]):
            if v is not None:
                g[(r, c)] = v
    return g


def _parse_watts(s):
    """'1,400W' -> 1400 ; returns None for multi-value or non-numeric cells."""
    if isinstance(s, (int, float)):
        return float(s)
    if not isinstance(s, str) or "|" in s or "\n" in s:
        return None
    digits = "".join(ch for ch in s if ch.isdigit())
    return float(digits) if digits else None


def extract_nv(path):
    """Morgan Stanley Nvidia AI-server model: chips, racks, customers, TDP, rack kW."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)

    # -- ODM share analysis: chip families (rows 2-7) + customer rack blocks --
    g = _grid(wb["ODM share analysis"], 220, 10)
    chip_years = [int(g[(1, c)]) for c in range(col("D"), col("H") + 1)]   # 2023..2027
    chip_families = []
    for r in range(2, 8):
        name = g.get((r, col("B")))
        if not isinstance(name, str):
            continue
        chip_families.append({
            "name": name.strip(),
            "values": [round(num(g.get((r, c))), 1) for c in range(col("D"), col("H") + 1)],
        })

    def customer_block(r_first, r_last):
        out = []
        for r in range(r_first, r_last + 1):
            name = g.get((r, col("B")))
            if not isinstance(name, str):
                continue
            name = name.strip()
            if name in ("Mix", "Other non-hyperscalers"):   # subtotal of the rows below it
                continue
            out.append({"name": name, "kracks": round(num(g.get((r, col("D")))), 3)})
        return out

    customers = {
        "2025": customer_block(171, 183),
        "2026": customer_block(203, 215),
        "total_2025": round(num(g.get((170, col("D")))), 3),
        "total_2026": round(num(g.get((202, col("D")))), 3),
    }

    # -- GB200-300 Tracker: monthly ramp + annual racks by SKU --
    g = _grid(wb["GB200-300 Tracker"], 60, 62)
    months, monthly = [], []
    yr = ""
    for c in range(col("C"), col("Z") + 1):
        lab = g.get((39, c))
        if not isinstance(lab, str):
            continue
        if c == col("C"):
            yr = "25"
        if c == col("O"):
            yr = "26"
        months.append(f"{lab} '{yr}")
        monthly.append(round(num(g.get((40, c))), 0))

    rack_years = [int(g[(39, c)]) for c in (col("BE"), col("BF"), col("BG"))]  # 2025..2027
    rack_sku_rows = [(40, "GB200 NVL72"), (41, "GB300 NVL72"), (42, "VR200 NVL72"), (43, "VR300")]
    rack_by_sku = [{
        "name": name,
        "values": [round(num(g.get((r, c))), 0) for c in (col("BE"), col("BF"), col("BG"))],
    } for r, name in rack_sku_rows]
    rack_totals = [round(num(g.get((44, c))), 0) for c in (col("BE"), col("BF"), col("BG"))]

    # -- GPU Roadmap: TDP per SKU (row 8) --
    g = _grid(wb["GPU Roadmap"], 10, 14)
    ms_tdp = []
    for c in range(col("C"), col("M") + 1):
        sku = g.get((3, c))
        w = _parse_watts(g.get((8, c)))
        if isinstance(sku, str) and w:
            ms_tdp.append({"sku": sku.split("\n")[0].strip(), "watts": w})

    # -- Nvidia Rack Layout: per-rack power usage kW (row 55) --
    g = _grid(wb["Nvidia Rack Layout"], 56, 14)
    rack_kw = {
        "GB200 NVL72": round(num(g.get((55, col("C")))), 2),
        "GB300 NVL72": round(num(g.get((55, col("E")))), 2),
        "VR200 NVL72": round(num(g.get((55, col("G")))), 2),
    }

    wb.close()
    return {
        "chip_years": chip_years, "chip_families": chip_families,
        "customers": customers,
        "months": months, "monthly_racks": monthly,
        "rack_years": rack_years, "rack_by_sku": rack_by_sku, "rack_totals": rack_totals,
        "ms_tdp": ms_tdp, "rack_kw": rack_kw,
    }


def extract_chip_watts(wb):
    """Client model 'Power Requirements per Chip': watts per SKU (unit col H == 'W').

    The tab lists ~150 SKUs across three repeated blocks (Nvidia/TPU/Trainium/Maia/
    AMD/boutique chips); a prior version of this function stopped at row 50 and
    silently dropped everything past the Nvidia+TPU section (all of Microsoft Maia,
    AMD's MI-series, AWS Inferentia3/Trainium4, and several boutique chips). Scan
    the full sheet and keep only the first occurrence of each SKU name so the
    later repeated blocks (some are '%' assumptions, not watts) don't overwrite it.
    """
    ws = wb["Power Requirements per Chip"]
    cE, cH, cK, cW = col("E"), col("H"), col("K"), col("W")
    seen = set()
    out = []
    for row in ws.iter_rows(min_row=11, max_row=ws.max_row, values_only=True):
        if len(row) <= cW:
            continue
        name, unit = row[cE], row[cH]
        if not (isinstance(name, str) and isinstance(unit, str) and unit.strip() == "W"):
            continue
        key = name.strip()
        if key in seen:
            continue
        watts = max((num(row[c]) for c in range(cK, cW + 1)), default=0.0)
        if watts > 0:
            seen.add(key)
            out.append({"sku": key, "watts": watts})
    return out


HN_YEARS = list(range(2017, 2033))  # 'Hyperscalers & Neoclouds' year cols K..Z (row 4)


def extract_buildout(wb):
    """'Hyperscalers & Neoclouds' company blocks + AI-lab MW from 'AI CPU Demand calculations'."""
    ws = wb["Hyperscalers & Neoclouds"]
    cB, cC, cD = col("B"), col("C"), col("D")
    ycols = [YEAR_COLS[yr] for yr in HN_YEARS]  # same K..Z layout as the supply tab

    def series(row):
        return [round(num(row[c]), 1) for c in ycols]

    hyperscalers, neoclouds = [], []
    global_neocloud = None
    block = hyperscalers
    current = None
    for row in ws.iter_rows(min_row=5, max_row=140, values_only=True):
        if len(row) <= max(ycols):
            continue
        b, c, d = row[cB], row[cC], row[cD]
        if isinstance(b, str) and "Global Neocloud" in b:
            block = neoclouds
            global_neocloud = series(row)
            current = None
            continue
        if isinstance(c, str) and c.strip() not in ("", "company"):
            name = c.strip()
            if "Neocloud capacity" in name:      # cross-reference row, not a company
                current = None
                continue
            current = {"name": name, "total": series(row), "sub": {}}
            block.append(current)
            continue
        if current is not None and isinstance(d, str) and d.strip():
            current["sub"][d.strip()] = series(row)

    # AI labs: rows with 'Identified Datacenters' in col C
    ws = wb["AI CPU Demand calculations"]
    ai_labs = []
    for row in ws.iter_rows(min_row=40, max_row=70, values_only=True):
        if len(row) <= max(ycols):
            continue
        c = row[cC]
        if isinstance(c, str) and "Identified Datacenter" in c:
            name = c.replace("Identified Datacenters", "").replace("Identified Datacenter", "").strip()
            ai_labs.append({"name": name, "values": series(row)})

    return {"years": HN_YEARS, "hyperscalers": hyperscalers, "neoclouds": neoclouds,
            "global_neocloud": global_neocloud, "ai_labs": ai_labs}


def extract_client_chips(wb):
    """Client file: global GPU/ASIC volumes, US DC power (Part IIB), ExaFLOPs need vs supply,
    TWh series, PUE/utilization, per-SKU volume rows, ExaFLOPs by generation, training wall."""
    g = _grid(wb["Power Summary"], 100, 20)
    years = list(range(2023, 2029))                      # cols D..I and L..Q

    def brow(r, c1="D", c2="I"):
        return [round(num(g.get((r, c))), 1) for c in range(col(c1), col(c2) + 1)]

    total_gpus = brow(43)
    total_asics = brow(63)
    # r57 'China/Ascend/MAIA 100/Other' is a grab-bag (Huawei Ascend, China parts,
    # MAIA, unspecified "Other"). Strip it so the ASIC line is comparable merchant
    # silicon (Google TPU, AWS Trainium, Meta/OpenAI/MSFT/xAI ASICs).
    asics_china_ascend_other = brow(57)
    total_asics_ex_china = [round(a - b, 1) for a, b in zip(total_asics, asics_china_ascend_other)]
    us_power_gw = brow(46, "L", "Q")
    us_additions_gw = brow(47, "L", "Q")

    # -- TWh rows (Part IIA global / IIB US) --
    twh_genai_global = brow(10, "L", "Q")            # 27.2 -> 801.6
    twh_base_global = brow(14, "L", "Q")             # 341 -> 549
    twh_genai_us = brow(39, "L", "Q")                # 16.3 -> 455.1
    global_dc_gw = brow(18, "L", "Q")                # 56 -> 205.6

    # -- PUE / utilization located by label (fixed addresses drift) --
    pue, util = None, None
    for r in range(1, 30):
        b = g.get((r, col("B")))
        if isinstance(b, str):
            if "Data Center PUE" in b:
                pue = [round(num(g.get((r, c))), 3) for c in range(col("D"), col("I") + 1)]
            elif "GPU Server Utilization Rate" in b:
                util = num(g.get((r, col("C"))))
    if pue is None or not util:
        raise SystemExit("Could not locate PUE / utilization rows by label in Power Summary")

    # -- per-SKU volume rows (labels col B) for the derived energy split --
    def vol_rows(r1, r2, skip=()):
        out = []
        for r in range(r1, r2 + 1):
            b = g.get((r, col("B")))
            if isinstance(b, str) and r not in skip:
                out.append({"row": r, "label": b.strip(), "values": brow(r)})
        return out

    gpu_vol_rows = vol_rows(24, 42)
    asic_vol_rows = vol_rows(45, 62, skip=(57,))     # China/Ascend/MAIA/Other excluded

    # -- ExaFLOPs by generation (rows 81-85) --
    exa_by_gen = []
    for r in range(81, 86):
        b = g.get((r, col("B")))
        if isinstance(b, str):
            exa_by_gen.append({"name": b.strip().strip('"'), "values": brow(r, "E", "I")})
    exa_total = brow(90, "E", "I")

    g = _grid(wb["Compute Demand"], 120, 10)
    exa_years = [2026, 2027, 2028]

    def crow(r, c1="G", c2="I", d=1):
        return [round(num(g.get((r, c))), d) for c in range(col(c1), col(c2) + 1)]

    return {
        "years": years,
        "total_gpus": total_gpus, "total_asics": total_asics,
        "total_asics_ex_china": total_asics_ex_china,
        "asics_china_ascend_other": asics_china_ascend_other,
        "us_power_gw": us_power_gw, "us_additions_gw": us_additions_gw,
        "twh_genai_global": twh_genai_global, "twh_base_global": twh_base_global,
        "twh_genai_us": twh_genai_us, "global_dc_gw": global_dc_gw,
        "pue": pue, "utilization": util,
        "gpu_vol_rows": gpu_vol_rows, "asic_vol_rows": asic_vol_rows,
        "exa_gen_years": [2024, 2025, 2026, 2027, 2028],
        "exa_by_gen": exa_by_gen, "exa_total": exa_total,
        "exa_years": exa_years,
        "exa_cumulative": crow(40),
        "exa_needed": crow(44),
        "exa_shortage": crow(45),
        "exa_shortage_pct": crow(46, d=4),
        # training wall (years E..I = 2026-2030 for r104; E..G = 2026-28 for r110/111)
        "training_years": [2026, 2027, 2028, 2029, 2030],
        "training_power_mw": crow(104, "E", "I"),
        "training_share_years": [2026, 2027, 2028],
        "training_share": crow(111, "E", "G", d=4),
        "bw_equiv_sold": crow(110, "E", "G"),
    }


# --------------------------------------------------------------------------- #
# Extraction: pages 3/4 — customer demand, unconstrained, balance, revisions,
# slippage, per-company additions
# --------------------------------------------------------------------------- #
def extract_customer_demand(wb):
    """'AI Demand by Customer': constrained demand aggregates + per-company AI power."""
    g = _grid(wb["AI Demand by Customer"], 300, 30)
    years = list(range(2017, 2033))                          # K..Z, header row 4
    ycols = [YEAR_COLS[yr] for yr in years]

    def series(r):
        return [round(num(g.get((r, c))), 1) for c in ycols]

    companies = [{"name": n, "values": series(r)} for r, n in
                 [(223, "Microsoft"), (224, "Meta"), (225, "Google"),
                  (226, "Amazon"), (227, "Oracle"), (228, "Apple")]]
    # neocloud/China blocks only carry 2023-2026 (cols Q..T)
    def qt(r):
        return [round(num(g.get((r, c))), 1) for c in (col("Q"), col("R"), col("S"), col("T"))]

    return {
        "years": years,
        "total_units_k": series(62),
        "install_base_k": series(117),
        "ai_power_mw": series(177),           # constrained cumulative AI critical IT power
        "hyperscaler_ai_mw": series(229),
        "companies": companies,
        "type_years": [2023, 2024, 2025, 2026],
        "type_hyperscaler": [round(num(g.get((229, c))), 1) for c in (col("Q"), col("R"), col("S"), col("T"))],
        "type_neoclouds": qt(258),
        "type_china": qt(290),
        "type_total": [round(num(g.get((177, c))), 1) for c in (col("Q"), col("R"), col("S"), col("T"))],
        "neo_coreweave": qt(249), "neo_tesla": qt(252), "neo_xai": qt(255),
    }


def extract_unconstrained(wb):
    """'Unconstrained AI Demand': chip-supply-driven demand curve (years J=2014..Z=2030)."""
    g = _grid(wb["Unconstrained AI Demand"], 370, 30)
    years = list(range(2014, 2031))
    ycols = [col("J") + i for i in range(len(years))]
    return {
        "years": years,
        "ai_power_mw": [round(num(g.get((365, c))), 1) for c in ycols],
        "total_units_k": [round(num(g.get((124, c))), 1) for c in ycols],
    }


def extract_na_balance(wb):
    """'NA DC Supply-Demand': the model's own supply-vs-demand balance (K=2017..X=2030)."""
    g = _grid(wb["NA DC Supply-Demand"], 160, 30)
    years = list(range(2017, 2031))
    ycols = [col("K") + i for i in range(len(years))]

    def series(r):
        return [round(num(g.get((r, c))), 1) for c in ycols]

    return {
        "years": years,
        "incr_demand_mw": series(154),
        "capacity_added_mw": series(155),
        "surplus_deficit_mw": series(156),
        "cum_surplus_deficit_mw": series(108),
        "total_demand_mw": series(100),
        "total_capacity_mw": series(91),
    }


def extract_revisions(wb):
    """'Revisions vs prior versions': Total Global Capacity vintage ladder (rows 11-51)
    + Oracle Overseas % revision row (r844)."""
    ws = wb["Revisions vs prior versions"]
    vintages = []
    oracle_pct = None
    for r, row in enumerate(ws.iter_rows(min_row=11, max_row=850, values_only=True), start=11):
        if len(row) <= col("X"):
            continue
        gg = row[col("G")]
        if r <= 51 and isinstance(gg, str) and gg.strip() != "% revision":
            vintages.append({
                "label": gg.strip(),
                "y2025": round(num(row[col("S")]), 0),
                "y2026": round(num(row[col("T")]), 0),
                "y2027": round(num(row[col("U")]), 0),
                "y2028": round(num(row[col("V")]), 0),
                "y2030": round(num(row[col("X")]), 0),
            })
            if gg.strip() == "First version":
                pass                                      # last row of the block
        if r == 844:                                       # Oracle Overseas '% revision'
            oracle_pct = [round(num(row[c]), 3) for c in
                          (col("T"), col("U"), col("V"), col("W"), col("X"))]
    vintages.reverse()                                     # oldest -> newest for charting
    return {"vintages": vintages,
            "oracle_pct_years": [2026, 2027, 2028, 2029, 2030],
            "oracle_pct": oracle_pct}


def extract_slippage(wb):
    """'NA Data Center Supply': modeled go-live slippage (Actual Live minus nameplate), US leaf rows."""
    import datetime
    ws = wb["NA Data Center Supply"]
    cA, cBY, cCB, cCT = col("A"), col("BY"), col("CB"), col("CT")
    diffs = []
    for row in ws.iter_rows(min_row=6, values_only=True):
        if len(row) <= max(cA, cBY, cCB, cCT):
            continue
        a = row[cA]
        if not (isinstance(a, str) and len(a) >= 8 and "-" in a):
            continue
        ct = row[cCT]
        if not (isinstance(ct, str) and ct.strip().upper() == "USA"):
            continue
        by, cb = row[cBY], row[cCB]
        if isinstance(by, datetime.datetime) and isinstance(cb, datetime.datetime):
            diffs.append((cb - by).days / 30.44)
    diffs.sort()
    n = len(diffs)
    buckets = {"~0": 0, "1-3 mo": 0, "4-6 mo": 0, "7-12 mo": 0, ">12 mo": 0}
    for d in diffs:
        if d <= 0.5:
            buckets["~0"] += 1
        elif d <= 3.5:
            buckets["1-3 mo"] += 1
        elif d <= 6.5:
            buckets["4-6 mo"] += 1
        elif d <= 12.5:
            buckets["7-12 mo"] += 1
        else:
            buckets[">12 mo"] += 1
    return {"n": n,
            "median_mo": round(diffs[n // 2], 1) if n else None,
            "mean_mo": round(sum(diffs) / n, 1) if n else None,
            "buckets": buckets,
            "pct_ge_7mo": round(100 * (buckets["7-12 mo"] + buckets[">12 mo"]) / n, 1) if n else None}


def extract_dashboard_adds(wb):
    """Model 'Dashboard' rows 17-25: per-company annual capacity additions (MW/yr, 2018-2032)."""
    g = _grid(wb["Dashboard"], 30, 30)
    years = list(range(2018, 2033))                        # L..Z
    ycols = [col("L") + i for i in range(len(years))]
    companies = []
    for r in range(17, 26):
        name = g.get((r, col("C")))
        if isinstance(name, str):
            companies.append({"name": name.strip(),
                              "values": [round(num(g.get((r, c))), 1) for c in ycols]})
    return {"years": years, "companies": companies}


# --------------------------------------------------------------------------- #
# Derivations: energy split (TWh by chip class) and compute split (GPU vs ASIC)
# --------------------------------------------------------------------------- #
# Watts per volume row (Power Summary row number -> W per chip).
# 'sourced' = taken from the model file's 'Power Requirements per Chip' /
# MS 'GPU Roadmap'; others are labeled analogs (assumptions).
ENERGY_WATTS = {
    # GPU rows
    24: ("H100", 700, True), 25: ("A100", 400, True), 26: ("B100/B200", 1000, True),
    27: ("R100 Rubin", 2300, True), 28: ("Rubin Ultra", 3600, True), 29: ("Feynman", 7000, True),
    30: ("R100+3", 3600, False), 31: ("Other NVDA", 500, False), 32: ("AMD Other", 500, False),
    # AMD MI-series: named rows found in 'Power Requirements per Chip' rows 99-105
    # (a prior version of this table guessed these; corrected against the sourced values)
    33: ("MI300X", 750, True), 34: ("MI325X", 1000, True), 35: ("MI350", 1400, True),
    36: ("MI400", 2500, True), 37: ("MI500 (inferred '+1' gen)", 2600, True), 38: ("MI '+2' gen (extrapolated)", 2700, False),
    # Intel: named rows 82-86 (Gaudi guesses happened to match; Falcon Shores corrected)
    39: ("Habana Gaudi2", 600, True), 40: ("Habana Gaudi3", 900, True), 41: ("Falcon Shores", 1200, True),
    42: ("Future Intel", 1000, False),
    # ASIC rows (r57 China/Ascend excluded upstream)
    45: ("TPU pre-v5 (v4 Pufferfish)", 200, True), 46: ("TPU v5p Viperfish", 550, True), 47: ("TPU v5e Viperlite", 300, True),
    48: ("TPU v6 Ghostlite", 390, True), 49: ("TPU v7 Ironwood", 980, True), 50: ("TPU v8AX Sunfish", 1200, True),
    51: ("TPU v9 Pumafish", 1500, True),
    # AWS: rows 75-81, 87-92 give named Trainium/Inferentia gens once the full
    # sheet (not just rows 11-50) is scanned
    52: ("Trainium2 Teton PD", 500, True), 53: ("Inferentia2/Trainium1", 200, True),
    54: ("Trainium3 Teton PD", 570, True), 55: ("Trainium4 (UALink/NVLink)", 1200, True),
    56: ("Trainium '2+3' gen (extrapolated)", 1400, False),
    # ByteDance/Meta/OpenAI: named rows 54-63 give real (and surprising) values —
    # ByteDance and early MTIA are far lower-power than the flagship-chip guess
    # this table used before; OpenAI's own 'Titan' chip is much higher-power
    58: ("ByteDance Custom ASIC Gen1/2", 90, True),
    59: ("Meta MTIA (Gen3-5 avg)", 1000, True),
    60: ("OpenAI 'Titan' (avg Titan1/2)", 1800, True),
    61: ("Microsoft Maia (avg of 5 named gens)", 962, True),
    62: ("xAI", 800, False),  # no named xAI chip found anywhere in the file
}
HOURS_PER_YEAR = 8760.0


def _watt_base(vol_rows, cumulative):
    """Sum of units x W per year (GW of chip power). cumulative=True uses the trailing
    install base (chips keep running once bought; no retirement within the window)."""
    n = 6                                                   # 2023..2028
    base = [0.0] * n
    for vr in vol_rows:
        w = ENERGY_WATTS.get(vr["row"])
        if w is None:
            continue
        run = 0.0
        for i in range(n):
            units_k = vr["values"][i]
            if cumulative:
                run += units_k
                base[i] += run * 1000 * w[1] / 1e9          # GW
            else:
                base[i] += units_k * 1000 * w[1] / 1e9
    return base


def derive_energy_split(client_chips):
    """TWh by chip class = install base x W x utilization x 8760h x PUE.
    Cumulative install base (no retirement). Validated against the client's
    own aggregate GenAI TWh (Power Summary r10)."""
    util = client_chips["utilization"]
    pue = client_chips["pue"]
    gpu_gw = _watt_base(client_chips["gpu_vol_rows"], cumulative=True)
    asic_gw = _watt_base(client_chips["asic_vol_rows"], cumulative=True)
    to_twh = lambda gw_list: [round(g * util * p * HOURS_PER_YEAR / 1000.0, 1)
                              for g, p in zip(gw_list, pue)]
    gpu_twh, asic_twh = to_twh(gpu_gw), to_twh(asic_gw)
    derived_total = [round(a + b, 1) for a, b in zip(gpu_twh, asic_twh)]
    client_total = client_chips["twh_genai_global"]
    ratio = [round(d / c, 2) if c else None for d, c in zip(derived_total, client_total)]
    return {"years": client_chips["years"], "gpu_twh": gpu_twh, "asic_twh": asic_twh,
            "derived_total_twh": derived_total, "client_total_twh": client_total,
            "validation_ratio": ratio, "utilization": util, "pue": pue}


def derive_compute_split(client_chips):
    """ASIC ExaFLOPs estimated at efficiency parity with same-year Nvidia:
    asic_exa[y] = nv_exa[y] x (asic annual watt-base / nvidia annual watt-base).
    Nvidia watt-base uses the Nvidia-only volume rows (24-31)."""
    nv_rows = [vr for vr in client_chips["gpu_vol_rows"] if vr["row"] <= 31]
    nv_gw = _watt_base(nv_rows, cumulative=False)
    asic_gw = _watt_base(client_chips["asic_vol_rows"], cumulative=False)
    # exa years are 2024-2028 -> volume years index 1..5
    nv_exa = client_chips["exa_total"]
    asic_exa = []
    for i, exa in enumerate(nv_exa):
        vi = i + 1                                          # 2024 -> volume index 1
        asic_exa.append(round(exa * asic_gw[vi] / nv_gw[vi], 1) if nv_gw[vi] else 0.0)
    return {"years": client_chips["exa_gen_years"], "nv_exa": nv_exa, "asic_exa_parity": asic_exa}


# Generation-level pairing of the two independent watts-per-chip tables.
# (client_substring, ms_substring) matched case-insensitively; None = absent.
WATTS_PAIRS = [
    ("A100", "a100", "a100"),
    ("H100 (SXM)", "h100 sxm", "h100"),
    ("B200", "b200 ", "b200"),
    ("GB200 NVL72", "gb200 nvl72", "gb200 nvl72"),
    ("GB300 NVL72", "gb300 nvl72", "gb300 nvl72"),
    ("VR200 NVL144", "vr200 nvl144", "vr200 nvl144"),
    ("VR300 / R300", "r300", "vr300"),
    ("F200", "f200", None),
]


def build_watts_compare(model_watts, ms_tdp):
    def find(items, key_field, sub):
        if sub is None:
            return None
        for it in items:
            s = it[key_field].lower()
            # exact-ish: substring match but avoid 'CPX' variants when not asked for
            if sub in s and ("cpx" in sub or "cpx" not in s):
                return it["watts"]
        return None

    out = []
    for gen, c_sub, m_sub in WATTS_PAIRS:
        cw = find(model_watts, "sku", c_sub)
        mw = find(ms_tdp, "sku", m_sub)
        if cw is not None or mw is not None:
            out.append({"gen": gen, "client": cw, "ms": mw})
    return out


BRIDGE_PUE = 1.3


def derive_bridge(nv, client_chips, supply):
    """Chip-implied DC power: racks/yr x per-rack usage kW x PUE -> GW/yr (global racks).

    VR300 has no rack-power figure in the NV file; approximated at VR200 kW.
    HGX/DGX 8-GPU servers are excluded (rack tracker covers NVL72-scale only),
    so this understates total chip-driven demand.
    """
    kw = dict(nv["rack_kw"])
    kw["VR300"] = kw["VR200 NVL72"]
    years = nv["rack_years"]                                   # [2025, 2026, 2027]
    implied = []
    for i in range(len(years)):
        mw = sum(s["values"][i] * kw[s["name"]] for s in nv["rack_by_sku"]) / 1000.0
        implied.append(round(mw * BRIDGE_PUE / 1000.0, 2))     # GW incl. PUE

    cy = client_chips["years"]
    client_adds = [client_chips["us_additions_gw"][cy.index(y)] for y in years]
    st = supply["year_series_gw"]["total"]
    sy = supply["years"]
    model_adds = [round(st[sy.index(y)] - st[sy.index(y - 1)], 2) for y in years]

    return {"years": years, "pue": BRIDGE_PUE, "rack_kw": kw,
            "implied_gw": implied, "client_us_additions_gw": client_adds,
            "model_us_additions_gw": model_adds}


# --------------------------------------------------------------------------- #
# Reconciliation
# --------------------------------------------------------------------------- #
def reconcile(shortfall, supply):
    needed = shortfall["shortfall_before_solutions"]
    us_uc = supply["totals_gw"]["uc"]
    us_pluc = supply["totals_gw"]["pluc"]
    us_planned = supply["totals_gw"]["planned"]
    us_add_26_28 = supply["totals_gw"]["delta_25_28"]

    return {
        "needed_shortfall": needed,                       # 37.71
        "client_uc": abs(shortfall["less_under_construction"]),   # 14.85
        "model_us_uc": us_uc,                             # ~26.28
        "client_demand": shortfall["demand_2026_28"],     # 67.56
        "model_us_additions_26_28": us_add_26_28,         # ~94.11
        "model_us_planned": us_planned,                   # ~202
        "model_us_pluc": us_pluc,                         # ~228
        "residual_pluc_vs_needed": round(us_pluc - needed, 2),
        "coverage_ratio_pluc": round(us_pluc / needed, 2) if needed else None,
        "coverage_ratio_additions": round(us_add_26_28 / shortfall["demand_2026_28"], 2)
        if shortfall["demand_2026_28"] else None,
    }


# --------------------------------------------------------------------------- #
# Self-check (fails loudly if a cell reference drifts)
# --------------------------------------------------------------------------- #
def self_check(shortfall, supply, nv=None, buildout=None, client_chips=None, extra_checks=None):
    checks = [
        ("shortfall 37.71", shortfall["shortfall_before_solutions"], 37.71, 0.1),
        ("demand 67.56", shortfall["demand_2026_28"], 67.56, 0.1),
        ("US UC 26.28", supply["totals_gw"]["uc"], 26.28, 0.5),
        ("US Planned+UC 228.38", supply["totals_gw"]["pluc"], 228.38, 1.0),
        ("US additions 26-28 94.11", supply["totals_gw"]["delta_25_28"], 94.11, 1.0),
    ]
    if nv is not None:
        gb300_2026 = next(s for s in nv["rack_by_sku"] if s["name"] == "GB300 NVL72")["values"][1]
        checks += [
            ("NV racks total 2025 = 28918", nv["rack_totals"][0], 28918, 2),
            ("NV GB300 racks 2026 = 73260", gb300_2026, 73260, 2),
            ("NV rack kW GB200 = 131.67", nv["rack_kw"]["GB200 NVL72"], 131.67, 0.1),
        ]
    if buildout is not None:
        msft = next(h for h in buildout["hyperscalers"] if h["name"] == "Microsoft")
        sb_na_2025 = msft["sub"]["Self-build - North America"][HN_YEARS.index(2025)]
        checks.append(("MSFT Self-build NA 2025 = 4485", sb_na_2025, 4485.1, 5))
    if client_chips is not None:
        checks += [
            ("Total GPUs 2028 = 9701k", client_chips["total_gpus"][-1], 9701.05, 2),
            ("US DC power 2028 = 119.4 GW", client_chips["us_power_gw"][-1], 119.4, 0.5),
        ]
    if extra_checks:
        checks += extra_checks
    problems = []
    for name, got, want, tol in checks:
        if abs(got - want) > tol:
            problems.append(f"  {name}: expected ~{want}, got {got}")
    if problems:
        raise SystemExit("SELF-CHECK FAILED (cell references may have drifted):\n" + "\n".join(problems))
    print("Self-check passed: all anchor values match planning figures.")


# --------------------------------------------------------------------------- #
# HTML rendering
# --------------------------------------------------------------------------- #
def render_html(data, template_file):
    with open(VENDOR_CHARTJS, "r", encoding="utf-8") as f:
        chartjs = f.read()
    with open(template_file, "r", encoding="utf-8") as f:
        template = f.read()
    payload = json.dumps(data)
    html = template.replace("/*__CHARTJS__*/", chartjs).replace("__DATA__", payload)
    return html


def main():
    for p in (CLIENT_FILE, MODEL_FILE, NV_FILE, VENDOR_CHARTJS, TEMPLATE_FILE,
              CHIPS_TEMPLATE_FILE, COMPUTE_TEMPLATE_FILE, SYNTHESIS_TEMPLATE_FILE):
        if not os.path.exists(p):
            raise SystemExit(f"Missing required file: {p}")

    print("Extracting client file (shortfall + chips) ...")
    client_wb = openpyxl.load_workbook(CLIENT_FILE, read_only=True, data_only=True)
    shortfall = extract_shortfall(client_wb)
    client_chips = extract_client_chips(client_wb)
    client_wb.close()

    print("Extracting model file (supply + buildout + chip watts) ...")
    model_wb = openpyxl.load_workbook(MODEL_FILE, read_only=True, data_only=True)
    supply = extract_supply(model_wb)
    buildout = extract_buildout(model_wb)
    model_watts = extract_chip_watts(model_wb)
    print("Extracting model file (demand, balance, revisions, slippage, additions) ...")
    customer = extract_customer_demand(model_wb)
    unconstrained = extract_unconstrained(model_wb)
    balance = extract_na_balance(model_wb)
    revisions = extract_revisions(model_wb)
    slippage = extract_slippage(model_wb)
    dash_adds = extract_dashboard_adds(model_wb)
    model_wb.close()

    print("Extracting NV server model (chips, racks, customers) ...")
    nv = extract_nv(NV_FILE)

    extra = [
        ("constrained AI power 2028 = 110410 MW", customer["ai_power_mw"][customer["years"].index(2028)], 110409.6, 5),
        ("unconstrained AI power 2028 = 126407 MW", unconstrained["ai_power_mw"][unconstrained["years"].index(2028)], 126406.8, 5),
        ("NA balance surplus 2028 = +7711 MW", balance["surplus_deficit_mw"][balance["years"].index(2028)], 7711.0, 5),
        ("NA total demand 2028 = 89913 MW", balance["total_demand_mw"][balance["years"].index(2028)], 89912.8, 5),
        ("GenAI TWh 2028 = 801.64", client_chips["twh_genai_global"][-1], 801.64, 0.5),
        ("US GenAI TWh 2028 = 455.07", client_chips["twh_genai_us"][-1], 455.07, 0.5),
        ("training share 2028 = 1.3228", client_chips["training_share"][-1], 1.3228, 0.01),
        ("ExaFLOPs total 2028 = 570733", client_chips["exa_total"][-1], 570733.1, 2),
        ("slippage median = 9.0 mo", slippage["median_mo"], 9.0, 0.5),
        ("vintage current 2026 = 100913", revisions["vintages"][-1]["y2026"], 100913, 2),
        ("vintage first 2030 = 140851", revisions["vintages"][0]["y2030"], 140851, 2),
        ("Oracle pct 2030 = -0.435", revisions["oracle_pct"][-1], -0.435, 0.005),
    ]
    self_check(shortfall, supply, nv, buildout, client_chips, extra)

    rec = reconcile(shortfall, supply)
    bridge = derive_bridge(nv, client_chips, supply)
    energy_split = derive_energy_split(client_chips)
    compute_split = derive_compute_split(client_chips)
    data = {"shortfall": shortfall, "supply": supply, "reconciliation": rec,
            "chips": {"nv": nv, "client": client_chips, "model_watts": model_watts,
                      "watts_compare": build_watts_compare(model_watts, nv["ms_tdp"]),
                      "buildout": buildout, "bridge": bridge},
            "compute": {"energy_split": energy_split, "compute_split": compute_split,
                        "customer": customer},
            "synthesis": {"unconstrained": unconstrained, "balance": balance,
                          "revisions": revisions, "slippage": slippage,
                          "dash_adds": dash_adds},
            "meta": {
                "client_file": os.path.basename(CLIENT_FILE),
                "model_file": os.path.basename(MODEL_FILE),
                "nv_file": os.path.basename(NV_FILE),
                "client_asof": "June 11, 2026",
                "model_asof": "April 7",
            }}

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(render_html(data, TEMPLATE_FILE))

    with open(OUT_CHIPS_HTML, "w", encoding="utf-8") as f:
        f.write(render_html(data, CHIPS_TEMPLATE_FILE))

    with open(OUT_COMPUTE_HTML, "w", encoding="utf-8") as f:
        f.write(render_html(data, COMPUTE_TEMPLATE_FILE))

    with open(OUT_SYNTHESIS_HTML, "w", encoding="utf-8") as f:
        f.write(render_html(data, SYNTHESIS_TEMPLATE_FILE))

    print("\n=== Reconciliation summary (GW) ===")
    print(f"  Needed (Power Shortfall before solutions): {rec['needed_shortfall']}")
    print(f"  Client Under-Construction assumption:      {rec['client_uc']}")
    print(f"  Model US Under-Construction:               {rec['model_us_uc']}")
    print(f"  Client demand 2026-28:                     {rec['client_demand']}")
    print(f"  Model US capacity added 2026-28:           {rec['model_us_additions_26_28']}")
    print(f"  Model US Planned:                          {rec['model_us_planned']}")
    print(f"  Model US Planned + UC:                     {rec['model_us_pluc']}")
    print(f"  Planned+UC vs needed (residual):           {rec['residual_pluc_vs_needed']}  "
          f"(coverage x{rec['coverage_ratio_pluc']})")
    print("\n=== Chips & buildout summary ===")
    print(f"  NVL72 racks 2025/26/27:                    {nv['rack_totals']}")
    print(f"  Chip-implied power GW (PUE {bridge['pue']}):           {bridge['implied_gw']}")
    print(f"  US DC power GW 2023-28 (client):           {client_chips['us_power_gw']}")
    print(f"  Neocloud companies found:                  {len(buildout['neoclouds'])}")
    print(f"  AI labs found:                             {[x['name'] for x in buildout['ai_labs']]}")
    print(f"  US tenant-tag GW: {supply['tenant_gw']}")
    print("\n=== Compute & synthesis summary ===")
    print(f"  Derived TWh (GPU+ASIC) 2023-28:            {energy_split['derived_total_twh']}")
    print(f"  Validation ratio vs client GenAI TWh:      {energy_split['validation_ratio']}")
    print(f"  ASIC parity ExaFLOPs 2024-28:              {compute_split['asic_exa_parity']}")
    print(f"  Constr vs unconstr AI power 2028 (GW):     "
          f"{customer['ai_power_mw'][customer['years'].index(2028)]/1000:.1f} vs "
          f"{unconstrained['ai_power_mw'][unconstrained['years'].index(2028)]/1000:.1f}")
    print(f"  Vintages captured:                         {len(revisions['vintages'])} "
          f"({revisions['vintages'][0]['label']} -> {revisions['vintages'][-1]['label']})")
    print(f"  Slippage: median {slippage['median_mo']} mo, {slippage['pct_ge_7mo']}% >=7mo, n={slippage['n']}")
    print(f"\nWrote:\n  {OUT_JSON}\n  {OUT_HTML}\n  {OUT_CHIPS_HTML}\n  {OUT_COMPUTE_HTML}\n  {OUT_SYNTHESIS_HTML}")


if __name__ == "__main__":
    main()
