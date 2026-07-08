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

OUT_JSON = os.path.join(HERE, "dashboard_data.json")
OUT_HTML = os.path.join(HERE, "dashboard.html")
OUT_CHIPS_HTML = os.path.join(HERE, "chips_dashboard.html")


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
    """Client model 'Power Requirements per Chip': watts per SKU (unit col H == 'W')."""
    ws = wb["Power Requirements per Chip"]
    cE, cH, cK, cW = col("E"), col("H"), col("K"), col("W")
    out = []
    for row in ws.iter_rows(min_row=11, max_row=50, values_only=True):
        if len(row) <= cW:
            continue
        name, unit = row[cE], row[cH]
        if not (isinstance(name, str) and isinstance(unit, str) and unit.strip() == "W"):
            continue
        watts = max((num(row[c]) for c in range(cK, cW + 1)), default=0.0)
        if watts > 0:
            out.append({"sku": name.strip(), "watts": watts})
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
    """Client file: global GPU/ASIC volumes, US DC power (Part IIB), ExaFLOPs need vs supply."""
    g = _grid(wb["Power Summary"], 70, 20)
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

    g = _grid(wb["Compute Demand"], 50, 10)
    exa_years = [2026, 2027, 2028]

    def crow(r, c1="G", c2="I"):
        return [round(num(g.get((r, c))), 1) for c in range(col(c1), col(c2) + 1)]

    return {
        "years": years,
        "total_gpus": total_gpus, "total_asics": total_asics,
        "total_asics_ex_china": total_asics_ex_china,
        "asics_china_ascend_other": asics_china_ascend_other,
        "us_power_gw": us_power_gw, "us_additions_gw": us_additions_gw,
        "exa_years": exa_years,
        "exa_cumulative": crow(40),
        "exa_needed": crow(44),
        "exa_shortage": crow(45),
        "exa_shortage_pct": crow(46),
    }


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
def self_check(shortfall, supply, nv=None, buildout=None, client_chips=None):
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
    for p in (CLIENT_FILE, MODEL_FILE, NV_FILE, VENDOR_CHARTJS, TEMPLATE_FILE, CHIPS_TEMPLATE_FILE):
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
    model_wb.close()

    print("Extracting NV server model (chips, racks, customers) ...")
    nv = extract_nv(NV_FILE)

    self_check(shortfall, supply, nv, buildout, client_chips)

    rec = reconcile(shortfall, supply)
    bridge = derive_bridge(nv, client_chips, supply)
    data = {"shortfall": shortfall, "supply": supply, "reconciliation": rec,
            "chips": {"nv": nv, "client": client_chips, "model_watts": model_watts,
                      "watts_compare": build_watts_compare(model_watts, nv["ms_tdp"]),
                      "buildout": buildout, "bridge": bridge},
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
    print(f"\nWrote:\n  {OUT_JSON}\n  {OUT_HTML}\n  {OUT_CHIPS_HTML}")


if __name__ == "__main__":
    main()
