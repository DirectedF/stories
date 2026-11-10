"""
State GDP vs. countries
=======================

How do the economies of U.S. states compare with the world's national economies?

This one script does the whole job, top to bottom:

  1. Download the publishers' files into ./raw/ (BEA state GDP, World Bank GDP,
     World Bank country list, plus two BEA files used only for checks).
  2. Pick the latest year that BOTH publishers cover (2025 first, then 2024).
  3. Build one ranking table: every country plus the 50 states and DC.
  4. Run the checks (sums, cross-source comparison, spot check, counts).
  5. Write combined_ranking.csv, two charts and findings.md.

Run it from anywhere, in isolated mode (-I) because ./raw/ holds downloaded files:

    python -I analysis.py      (any Python 3.11+ with pandas, requests, matplotlib, openpyxl)
    ... -I analysis.py --refresh     (re-download every source file first)

Sources
  - U.S. Bureau of Economic Analysis (BEA), table SAGDP2 "GDP in current dollars",
    all-industry total (LineCode 1), millions of current dollars.
  - World Bank, World Development Indicators, NY.GDP.MKTP.CD "GDP (current US$)".
"""

import argparse
import datetime as dt
import hashlib
import io
import json
import re
import sys
import zipfile
from pathlib import Path

import matplotlib

matplotlib.use("Agg")  # draw to files only, no window
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
import openpyxl
import pandas as pd
import requests

# --------------------------------------------------------------------------
# Settings
# --------------------------------------------------------------------------

OUT = Path(__file__).resolve().parent  # the story folder
RAW = OUT / "raw"                      # downloads only; never put scripts here

SOURCES = {
    # BEA bulk file: every SAGDP table for every state, 1997 onward.
    "bea_sagdp": ("https://apps.bea.gov/regional/zip/SAGDP.zip", "SAGDP.zip"),
    # BEA quarterly state GDP (SQGDP2). Used only to cross-check California.
    "bea_sqgdp": ("https://apps.bea.gov/regional/zip/SQGDP.zip", "SQGDP.zip"),
    # BEA vintage history of national GDP. Used only to explain the
    # BEA vs World Bank U.S. difference (check b).
    "bea_vintage": (
        "https://www.bea.gov/sites/default/files/2026-09/gdp-gdi-vintage-history.xlsx",
        "gdp-gdi-vintage-history.xlsx",
    ),
    # World Bank GDP for every economy and aggregate, 2024 and 2025.
    "wb_gdp": (
        "https://api.worldbank.org/v2/country/all/indicator/NY.GDP.MKTP.CD"
        "?date=2024:2025&format=json&per_page=20000",
        "wb_gdp.json",
    ),
    # World Bank economy list. Its "region" field marks aggregates.
    "wb_countries": (
        "https://api.worldbank.org/v2/country?format=json&per_page=400",
        "wb_countries.json",
    ),
}

CANDIDATE_YEARS = [2025, 2024]   # newest first; use the first one both sources cover
WB_MIN_COUNTRIES = 150           # a World Bank year "counts" only with this many values

# California spot check (check c). Recorded by hand from BEA's own published page,
# not from the bulk file, so it is an independent read of the same estimate.
BEA_PUBLISHED_CA = {
    "year": 2025,
    "millions": 4_297_449.8,
    "page": "BEA BEARFACTS, California fact sheet, Gross Domestic Product tab",
    "quote": "In 2025, California current-dollar GDP was 4,297,449.8 million "
             "and ranked 1st in the United States. (GDP last published on September 30, 2026)",
    "url": "https://apps.bea.gov/regional/bearfacts/",
    "checked": "2026-10-23",
}
# BEA's 2026 annual update date. Vintages before this date are "pre-revision".
BEA_ANNUAL_UPDATE = dt.date(2026, 9, 30)

# Chart style (from the brief)
STATE_COLOR = "#2a78d6"
COUNTRY_COLOR = "#9aa1aa"
TEXT_COLOR = "#121417"
TITLE_COLOR = "#13294b"
MUTED_TEXT = "#4a5160"
SOURCE_LINE = "Sources: U.S. Bureau of Economic Analysis; World Bank, World Development Indicators"
TOP_N = 25  # bars in the top-economies chart

# World Bank names are kept in the CSV for traceability. Charts and findings use
# these common short names instead.
DISPLAY_NAMES = {
    "Russian Federation": "Russia",
    "Korea, Rep.": "South Korea",
    "Turkiye": "Türkiye",
    "Iran, Islamic Rep.": "Iran",
    "Egypt, Arab Rep.": "Egypt",
    "Hong Kong SAR, China": "Hong Kong SAR (China)",
    "Macao SAR, China": "Macao SAR (China)",
    "Viet Nam": "Vietnam",
    "Slovak Republic": "Slovakia",
}
# A few names read better with "the" in a sentence.
WITH_THE = {"United States", "United Kingdom", "Netherlands", "Philippines", "United Arab Emirates",
            "Dominican Republic"}

# 2-digit state FIPS codes 01..56 (with gaps) plus 11 = DC. US total is 00000.
STATE_FIPS_PATTERN = r"^(0[1-9]|[1-5][0-9])000$"


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------

def display(name):
    return DISPLAY_NAMES.get(name, name)


def in_sentence(name):
    n = display(name)
    return f"the {n}" if n in WITH_THE else n


def trillions(usd, digits=2):
    return f"${usd / 1e12:,.{digits}f} trillion"


def billions(usd, digits=1):
    return f"${usd / 1e9:,.{digits}f} billion"


def ordinal(n):
    suffix = "th" if 10 <= n % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def join_and(items, word="and"):
    items = list(items)
    if len(items) <= 1:
        return "".join(items)
    return ", ".join(items[:-1]) + f" {word} " + items[-1]


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


# --------------------------------------------------------------------------
# Step 1: download
# --------------------------------------------------------------------------

def download_all(refresh):
    """Fetch each source into ./raw/ unless it is already there (or --refresh)."""
    RAW.mkdir(exist_ok=True)
    for key, (url, filename) in SOURCES.items():
        target = RAW / filename
        if target.exists() and not refresh:
            print(f"  using existing raw/{filename}")
            continue
        print(f"  downloading {url}")
        resp = requests.get(url, timeout=120, headers={"User-Agent": "state-gdp-vs-countries/1.0"})
        resp.raise_for_status()
        target.write_bytes(resp.content)
        print(f"    -> raw/{filename} ({len(resp.content):,} bytes)")


# --------------------------------------------------------------------------
# Step 2: read each source
# --------------------------------------------------------------------------

def read_bea_annual():
    """BEA SAGDP2 all-areas CSV -> (every industry line, all-industry total rows (LineCode 1),
    file name, "Last updated" note). The industry lines are used only by check (a)."""
    with zipfile.ZipFile(RAW / "SAGDP.zip") as z:
        name = next(n for n in z.namelist() if re.match(r"SAGDP2__ALL_AREAS_\d{4}_\d{4}\.csv$", n))
        df = pd.read_csv(io.BytesIO(z.read(name)), encoding="latin-1", dtype=str)
    for col in ["GeoFIPS", "GeoName", "LineCode", "Description", "Unit"]:
        df[col] = df[col].str.strip().str.strip('"')
    # The last rows of the file are notes, e.g. "Last updated: September 30, 2026-- ...".
    notes = df["GeoFIPS"][df["GeoName"].isna()].tolist()
    last_updated = next((n for n in notes if n.startswith("Last updated")), "unknown")
    total = df[df["LineCode"] == "1"].copy()
    assert (total["Unit"] == "Millions of current dollars").all(), "unexpected BEA unit"
    all_lines = df[df["LineCode"].notna()].copy()
    return all_lines, total, name, last_updated


def us_minus_states_by_line(df, year):
    """For each BEA industry line: U.S. value minus the sum of the 50 states + DC, in millions.
    BEA's footnote says the all-industry gap is overseas federal activity, so it should sit
    entirely in the federal civilian (LineCode 84) and military (85) lines."""
    col = str(year)
    vals = pd.to_numeric(df[col], errors="coerce")
    is_state = df["GeoFIPS"].str.match(STATE_FIPS_PATTERN)
    us = vals[df["GeoFIPS"] == "00000"].groupby(df["LineCode"]).sum()
    st = vals[is_state].groupby(df["LineCode"]).sum()
    return (us - st), df[df["GeoFIPS"] == "00000"].set_index("LineCode")["Description"].str.strip()


def read_bea_quarterly_average(fips, year):
    """Average of the four quarterly values (SQGDP2) for one area and year, in millions."""
    with zipfile.ZipFile(RAW / "SQGDP.zip") as z:
        name = next(n for n in z.namelist() if re.match(r"SQGDP2__ALL_AREAS_\d{4}_\d{4}\.csv$", n))
        df = pd.read_csv(io.BytesIO(z.read(name)), encoding="latin-1", dtype=str)
    for col in ["GeoFIPS", "LineCode"]:
        df[col] = df[col].str.strip().str.strip('"')
    row = df[(df["GeoFIPS"] == fips) & (df["LineCode"] == "1")].iloc[0]
    quarters = [f"{year}:Q{q}" for q in range(1, 5)]
    values = [float(row[q]) for q in quarters]
    return sum(values) / 4, dict(zip(quarters, values))


def read_wb():
    """World Bank GDP rows + economy metadata. Returns (gdp rows, metadata dict, lastupdated)."""
    gdp_doc = json.loads((RAW / "wb_gdp.json").read_text(encoding="utf-8"))
    header, rows = gdp_doc[0], gdp_doc[1]
    # One page is expected (per_page=20000). Stop if the API paginated anyway.
    assert int(header["pages"]) == 1, f"World Bank returned {header['pages']} pages; raise per_page"
    meta_doc = json.loads((RAW / "wb_countries.json").read_text(encoding="utf-8"))
    assert int(meta_doc[0]["pages"]) == 1, "World Bank country list paginated; raise per_page"
    meta = {m["id"]: m for m in meta_doc[1]}
    return rows, meta, header.get("lastupdated", "unknown")


# --------------------------------------------------------------------------
# Step 3: choose the year
# --------------------------------------------------------------------------

def choose_year(bea_total, wb_rows, wb_meta):
    """Latest year where BEA has every state + DC + U.S. and the World Bank has
    the U.S. plus at least WB_MIN_COUNTRIES non-aggregate economies."""
    countries = {k for k, m in wb_meta.items() if m["region"]["value"].strip() != "Aggregates"}
    notes = []
    for year in CANDIDATE_YEARS:
        col = str(year)
        bea_ok = col in bea_total.columns and pd.to_numeric(
            bea_total.loc[bea_total["GeoFIPS"].str.match(STATE_FIPS_PATTERN) | (bea_total["GeoFIPS"] == "00000"), col],
            errors="coerce").notna().sum() == 52
        wb_year = [r for r in wb_rows if r["date"] == col and r["countryiso3code"] in countries]
        wb_with_value = sum(r["value"] is not None for r in wb_year)
        wb_us = any(r["countryiso3code"] == "USA" and r["value"] is not None for r in wb_year)
        wb_ok = wb_us and wb_with_value >= WB_MIN_COUNTRIES
        notes.append(f"{year}: BEA {'complete' if bea_ok else 'missing'}; "
                     f"World Bank {wb_with_value} of {len(wb_year)} economies have a value")
        if bea_ok and wb_ok:
            return year, notes
    raise SystemExit("No candidate year is available in both sources:\n" + "\n".join(notes))


# --------------------------------------------------------------------------
# Step 4: build the combined ranking
# --------------------------------------------------------------------------

def build_tables(year, bea_total, wb_rows, wb_meta):
    col = str(year)

    # States + DC, converted from millions of dollars to dollars.
    st = bea_total[bea_total["GeoFIPS"].str.match(STATE_FIPS_PATTERN)].copy()
    states = pd.DataFrame({
        "name": st["GeoName"],
        "type": "state",          # the 50 states and DC (BEA's state-level areas)
        "gdp_usd": pd.to_numeric(st[col], errors="raise") * 1e6,
        "code": st["GeoFIPS"],
    })
    us_bea_usd = float(bea_total.loc[bea_total["GeoFIPS"] == "00000", col].iloc[0]) * 1e6

    # World Bank: keep economies whose region is not "Aggregates".
    # Some aggregate rows (income groups) come with a blank ISO3 code, so match
    # those through the 2-letter id in the metadata instead.
    aggregates = {k for k, m in wb_meta.items() if m["region"]["value"].strip() == "Aggregates"}
    iso2_to_id = {m["iso2Code"]: k for k, m in wb_meta.items()}
    year_rows = [r for r in wb_rows if r["date"] == col]
    for r in year_rows:
        r["_id"] = r["countryiso3code"] or iso2_to_id.get(r["country"]["id"])
    unmatched = [r["country"]["value"] for r in year_rows if r["_id"] not in wb_meta]
    non_agg = [r for r in year_rows if r["_id"] in wb_meta and r["_id"] not in aggregates]
    have = [r for r in non_agg if r["value"] is not None]
    lacking = sorted(r["country"]["value"] for r in non_agg if r["value"] is None)
    countries = pd.DataFrame({
        "name": [r["country"]["value"] for r in have],
        "type": "country",
        "gdp_usd": [float(r["value"]) for r in have],
        "code": [r["countryiso3code"] for r in have],
    })

    # rank_among_countries = 1 + number of countries with a larger GDP.
    # Same rule for a country and for a state, so ties share a rank and a state
    # never pushes a country down (states are compared with countries only).
    country_values = countries["gdp_usd"].to_numpy()
    for df in (countries, states):
        df["rank_among_countries"] = [int((country_values > v).sum()) + 1 for v in df["gdp_usd"]]

    combined = (pd.concat([countries, states], ignore_index=True)
                .sort_values(["gdp_usd", "name"], ascending=[False, True])
                .reset_index(drop=True))

    info = {
        "us_bea_usd": us_bea_usd,
        "aggregates": aggregates,
        "n_meta": len(wb_meta),
        "n_non_agg": len(non_agg),
        "lacking": lacking,
        "unmatched": unmatched,
        "wb_world_usd": next((float(r["value"]) for r in year_rows
                              if r["countryiso3code"] == "WLD" and r["value"] is not None), None),
    }
    return combined, countries, states, info


# --------------------------------------------------------------------------
# Step 5: checks
# --------------------------------------------------------------------------

def pre_revision_us_gdp(year):
    """From BEA's vintage-history workbook: for each quarter of `year`, the last
    current-dollar GDP estimate published BEFORE the annual update date; then the
    average of the four (BEA's annual GDP is the average of its quarters)."""
    wb = openpyxl.load_workbook(RAW / "gdp-gdi-vintage-history.xlsx", read_only=True, data_only=True)
    ws = wb["Vintage History"]
    quarter = None
    found = {}
    for row in ws.iter_rows(values_only=True):
        first = str(row[0]) if row[0] is not None else ""
        if re.fullmatch(r"\d{4}Q[1-4]", first):
            quarter = first
            continue
        if not quarter or not quarter.startswith(str(year)) or row[1] in (None, "Vintage"):
            continue
        m = re.match(r"([A-Z][a-z]{2}) (\d{1,2}), (\d{4})", str(row[6] or ""))
        if not m:
            continue
        vdate = dt.datetime.strptime(" ".join(m.groups()), "%b %d %Y").date()
        if vdate < BEA_ANNUAL_UPDATE:
            best = found.get(quarter)
            if best is None or vdate > best[0]:
                found[quarter] = (vdate, float(str(row[2]).replace(",", "")), row[1])
    if len(found) != 4:
        return None, found
    return sum(v[1] for v in found.values()) / 4 * 1e9, found


def run_checks(year, combined, countries, states, info, bea_lines):
    checks = []

    # (a) 50 states + DC vs BEA's own U.S. total.
    # BEA's footnote: the U.S. total also counts overseas activity (military and federal
    # civilian staff abroad), which no state contains. So test where the gap sits, line by line:
    # private industries (LineCode 2) and state and local government (86) should match to
    # rounding, and federal civilian (84) + military (85) should hold the whole gap.
    states_sum = states["gdp_usd"].sum()
    us = info["us_bea_usd"]
    gap = us - states_sum
    by_line, desc = us_minus_states_by_line(bea_lines, year)
    gap_m = gap / 1e6
    federal = by_line["84"] + by_line["85"]
    a_pass = (len(states) == 51 and gap > 0 and abs(by_line["2"]) <= 10 and abs(by_line["86"]) <= 10
              and abs(federal - gap_m) <= 10)
    checks.append({
        "id": "a", "name": "Sum of 50 states + DC vs BEA's U.S. total", "pass": a_pass,
        "detail": (f"{len(states)} areas sum to ${states_sum / 1e6:,.1f} million; BEA's U.S. line is "
                   f"${us / 1e6:,.1f} million, a gap of ${gap_m:,.1f} million ({gap / us:.2%}). "
                   f"Where the gap sits (U.S. minus sum of states, millions): private industries "
                   f"{by_line['2']:,.1f}; state and local government {by_line['86']:,.1f}; federal military "
                   f"{by_line['85']:,.1f}; federal civilian {by_line['84']:,.1f}. Military plus federal civilian = "
                   f"{federal:,.1f}, i.e. the whole gap, matching BEA's footnote that the U.S. total also counts "
                   "overseas military and federal civilian activity, which belongs to no state. Pass rule: "
                   "51 areas; private and state-and-local gaps within $10 million (rounding); military plus "
                   "federal civilian within $10 million of the total gap."),
    })

    # (b) BEA U.S. total vs World Bank U.S. value for the same year.
    wb_us = float(countries.loc[countries["code"] == "USA", "gdp_usd"].iloc[0])
    diff = us - wb_us
    pre_rev, found = pre_revision_us_gdp(year)
    if pre_rev is not None:
        vint = "; ".join(f"{q} ${v[1]:,.1f}B ({v[2]}, {v[0]:%b %d, %Y})" for q, v in sorted(found.items()))
        closer = abs(wb_us - pre_rev) < abs(wb_us - us)
        reason = (f"BEA revised {year} on {BEA_ANNUAL_UPDATE:%B %d, %Y}. BEA's quarterly estimates published "
                  f"before that date average ${pre_rev / 1e9:,.1f} billion ({vint}), "
                  f"{abs(wb_us - pre_rev) / pre_rev:.2%} from the World Bank figure. "
                  + ("Likely reason: the World Bank value predates BEA's September revision (it sits much closer "
                     "to BEA's pre-revision numbers); the World Bank does not state which BEA vintage it used."
                     if closer else
                     "The World Bank value is not closer to BEA's pre-revision numbers, so the reason is unexplained."))
    else:
        reason = "Pre-revision BEA vintages could not be read from the vintage-history file."
    checks.append({
        "id": "b", "name": "BEA U.S. total vs World Bank U.S. value", "pass": abs(diff) / us <= 0.01,
        "detail": (f"BEA ${us / 1e9:,.1f} billion vs World Bank ${wb_us / 1e9:,.1f} billion: BEA is higher by "
                   f"${diff / 1e9:,.1f} billion ({diff / us:.2%}). {reason} Pass rule: within 1%."),
    })

    # (c) California: bulk file vs BEA's published page, plus BEA's quarterly file.
    ca_bulk_m = float(states.loc[states["code"] == "06000", "gdp_usd"].iloc[0]) / 1e6
    ca_q_avg, ca_q = read_bea_quarterly_average("06000", year)
    pub = BEA_PUBLISHED_CA
    c_pass = (pub["year"] == year and abs(ca_bulk_m - pub["millions"]) < 0.05
              and abs(ca_bulk_m - ca_q_avg) <= 0.1)
    checks.append({
        "id": "c", "name": "California spot check against BEA's published figure", "pass": c_pass,
        "detail": (f"Bulk file (SAGDP2): ${ca_bulk_m:,.1f} million. {pub['page']} ({pub['url']}, checked "
                   f"{pub['checked']}) reads: \"{pub['quote']}\" Match: "
                   f"{'exact' if abs(ca_bulk_m - pub['millions']) < 0.05 else 'NO'}. Second BEA product "
                   f"(SQGDP2 quarterly file): the four {year} quarters "
                   f"({', '.join(f'{v:,.1f}' for v in ca_q.values())}) average ${ca_q_avg:,.2f} million."),
    })

    # (d) Countries kept after removing aggregates; how many lacked data.
    n_agg = len(info["aggregates"])
    kept = len(countries)
    leaked = {"WLD", "EUU", "EMU", "OED", "HIC", "LMY", "SSF", "EAS"} & set(countries["code"])
    world = info["wb_world_usd"]
    world_share = countries["gdp_usd"].sum() / world if world else None
    d_pass = (not leaked and info["n_non_agg"] == info["n_meta"] - n_agg and not info["unmatched"]
              and world_share is not None and 0.95 <= world_share <= 1.005)
    checks.append({
        "id": "d", "name": "Countries kept after removing aggregates", "pass": d_pass,
        "detail": (f"World Bank list has {info['n_meta']} entries, {n_agg} marked \"Aggregates\", leaving "
                   f"{info['n_non_agg']} economies. {kept} have a {year} value and are kept; "
                   f"{len(info['lacking'])} lack one and are excluded ({'; '.join(info['lacking'])}). "
                   f"Leak test: the kept countries sum to ${countries['gdp_usd'].sum() / 1e12:,.2f} trillion, "
                   f"{world_share:.1%} of the World Bank's own \"World\" total "
                   f"(${world / 1e12:,.2f} trillion); a leaked aggregate would push this far above 100%. "
                   "Pass rule: no aggregate codes kept and the sum is 95% to 100.5% of World."),
    })

    # (e) Sensitivity: would the countries without a value for this year change the story?
    # Uses their PREVIOUS-year World Bank value as a stand-in, for this check only.
    prev = str(year - 1)
    prev_vals = {}
    for r in json.loads((RAW / "wb_gdp.json").read_text(encoding="utf-8"))[1]:
        if r["date"] == prev and r["country"]["value"] in info["lacking"] and r["value"] is not None:
            prev_vals[r["country"]["value"]] = float(r["value"])
    alt_values = list(countries["gdp_usd"]) + list(prev_vals.values())
    alt_rank = {row.name: sum(v > row.gdp_usd for v in alt_values) + 1 for row in states.itertuples()}
    now_rank = dict(zip(states["name"], states["rank_among_countries"]))
    by_size = states.sort_values("gdp_usd", ascending=False)["name"]
    changed = [n for n in by_size if alt_rank[n] != now_rank[n]]
    max_shift = max((alt_rank[n] - now_rank[n] for n in changed), default=0)
    fifty = states[states["name"] != "District of Columbia"]
    top50_now = int((fifty["rank_among_countries"] <= 50).sum())
    top50_alt = sum(alt_rank[n] <= 50 for n in fifty["name"])
    big3_same = all(alt_rank[n] == int(states.loc[states["name"] == n, "rank_among_countries"].iloc[0])
                    for n in ["California", "Texas", "New York"])
    big_missing = {k: v for k, v in prev_vals.items() if v >= 50e9}
    checks.append({
        "id": "e", "name": f"Sensitivity: countries missing {year} data", "pass": big3_same and top50_now == top50_alt,
        "detail": (f"{len(prev_vals)} of the {len(info['lacking'])} excluded economies have a {prev} value. "
                   f"Missing ones above $50 billion in {prev}: "
                   f"{', '.join(f'{display(k)} {billions(v)}' for k, v in big_missing.items()) or 'none'}. "
                   f"If they were added at their {prev} values, {len(changed)} of 51 state/DC ranks would "
                   f"move down {'one place' if max_shift == 1 else f'by up to {max_shift} places'} "
                   f"({', '.join(changed) if len(changed) <= 6 else (f'the {len(changed)} smallest, {changed[0]} down to {changed[-1]}' if changed == list(by_size)[-len(changed):] else ', '.join(changed))}); "
                   f"California, Texas and New York {'keep' if big3_same else 'do NOT keep'} their ranks; "
                   f"states in the top 50: {top50_now} now vs {top50_alt} with the stand-ins. "
                   "Pass rule: headline numbers unchanged."),
        "changed": changed,
    })

    # (f) Completeness of the BEA side.
    f_pass = len(states) == 51 and states["gdp_usd"].notna().all() and "11000" in set(states["code"])
    checks.append({
        "id": "f", "name": "BEA completeness", "pass": f_pass,
        "detail": f"{len(states)} state-level areas (50 states + DC), none blank for {year}.",
    })
    return checks


# --------------------------------------------------------------------------
# Step 6: charts
# --------------------------------------------------------------------------

plt.rcParams.update({"font.family": "DejaVu Sans", "text.color": TEXT_COLOR,
                     "axes.edgecolor": "#d5d9de", "savefig.facecolor": "white"})


def chart_header(fig, title, subtitle):
    fig.text(0.035, 0.955, title, fontsize=13.5, fontweight="bold", color=TITLE_COLOR, va="top")
    fig.text(0.035, 0.900, subtitle, fontsize=8.6, color=MUTED_TEXT, va="top", linespacing=1.45)
    fig.legend(handles=[Patch(color=STATE_COLOR), Patch(color=COUNTRY_COLOR)],
               labels=["U.S. state", "Country"], loc="upper left", bbox_to_anchor=(0.028, 0.838),
               ncol=2, frameon=False, fontsize=8.6, handlelength=1.0, handleheight=1.0,
               columnspacing=1.6, handletextpad=0.5)
    fig.text(0.035, 0.022, SOURCE_LINE, fontsize=7.2, color=MUTED_TEXT, va="bottom")


def clean_axis(ax):
    for side in ("top", "right", "bottom"):
        ax.spines[side].set_visible(False)
    ax.spines["left"].set_color("#c9ced4")
    ax.xaxis.set_visible(False)  # every bar is labeled, so no x-axis or gridlines
    ax.tick_params(axis="y", length=0, pad=6)


def chart_top(combined, year, us_wb_usd, path):
    """Top economies, U.S. states and countries mixed, U.S. as a country left out."""
    top = combined[combined["code"] != "USA"].head(TOP_N).reset_index(drop=True)

    # Title: the countries that sit between California and the next state down.
    ca_pos = top.index[top["name"] == "California"][0]
    below = []
    for _, r in top.iloc[ca_pos + 1:].iterrows():
        if r["type"] == "state":
            break
        below.append(in_sentence(r["name"]).replace("the United Kingdom", "the U.K."))
    title = f"California alone has a bigger economy than {join_and(below[:3], 'or')}"

    fig = plt.figure(figsize=(8, 5.5), dpi=200)
    ax = fig.add_axes([0.20, 0.06, 0.74, 0.725])
    y = range(len(top))
    colors = [STATE_COLOR if t == "state" else COUNTRY_COLOR for t in top["type"]]
    ax.barh(list(y), top["gdp_usd"] / 1e12, height=0.68, color=colors, edgecolor="white", linewidth=0.6)
    ax.invert_yaxis()  # largest at top
    xmax = top["gdp_usd"].max() / 1e12
    ax.set_xlim(0, xmax * 1.1)
    ax.set_ylim(len(top) - 0.5, -0.6)
    for i, v in enumerate(top["gdp_usd"] / 1e12):
        ax.text(v + xmax * 0.008, i, f"${v:,.2f}T", va="center", ha="left", fontsize=7.4, color=TEXT_COLOR)
    ax.set_yticks(list(y))
    ax.set_yticklabels([display(n) for n in top["name"]], fontsize=8)
    for lbl, t in zip(ax.get_yticklabels(), top["type"]):
        lbl.set_fontweight("bold" if t == "state" else "normal")  # second cue besides color
    clean_axis(ax)
    chart_header(
        fig, title,
        f"The {TOP_N} largest economies in {year}, with U.S. states counted separately. "
        "Nominal GDP, current U.S. dollars.\n"
        f"The United States as a whole (${us_wb_usd / 1e12:,.2f}T, World Bank) is left out because "
        "its states are shown. States in bold.",
    )
    fig.savefig(path, dpi=200, facecolor="white")
    plt.close(fig)
    return title


def chart_small(pairs, year, n_countries, path):
    """The five smallest state economies, each next to the closest country in GDP."""
    labels, values, colors, kinds = [], [], [], []
    ypos, y = [], 0.0
    for p in pairs:
        for name, v, kind in ((p["state"], p["state_usd"], "state"), (display(p["country"]), p["country_usd"], "country")):
            labels.append(name); values.append(v); kinds.append(kind)
            colors.append(STATE_COLOR if kind == "state" else COUNTRY_COLOR)
            ypos.append(y); y += 0.82
        y += 0.75  # gap between pairs

    smallest = pairs[0]
    pct = (smallest["country_usd"] - smallest["state_usd"]) / smallest["state_usd"]
    fig = plt.figure(figsize=(8, 5.5), dpi=200)
    ax = fig.add_axes([0.20, 0.06, 0.74, 0.725])
    ax.barh(ypos, [v / 1e9 for v in values], height=0.7, color=colors, edgecolor="white", linewidth=0.6)
    xmax = max(values) / 1e9
    ax.set_xlim(0, xmax * 1.12)
    ax.set_ylim(max(ypos) + 0.6, -0.6)
    for yy, v in zip(ypos, values):
        ax.text(v / 1e9 + xmax * 0.008, yy, f"${v / 1e9:,.1f}B", va="center", ha="left", fontsize=8, color=TEXT_COLOR)
    ax.set_yticks(ypos)
    ax.set_yticklabels(labels, fontsize=8.6)
    for lbl, k in zip(ax.get_yticklabels(), kinds):
        lbl.set_fontweight("bold" if k == "state" else "normal")
    clean_axis(ax)
    title = f"Even the smallest state economy is the size of a country"
    chart_header(
        fig, title,
        f"The five smallest U.S. state economies in {year}, each paired with the country closest to it in GDP.\n"
        f"Nominal GDP, current U.S. dollars. {smallest['state']} ({billions(smallest['state_usd'])}) would rank "
        f"{ordinal(smallest['rank'])} of {n_countries} countries. States in bold.",
    )
    fig.savefig(path, dpi=200, facecolor="white")
    plt.close(fig)
    return title


# --------------------------------------------------------------------------
# Step 7: findings.md
# --------------------------------------------------------------------------

def neighbours(state_value, countries):
    """The country just above and just below a value, among countries only."""
    above = countries[countries["gdp_usd"] > state_value].sort_values("gdp_usd").head(1)
    below = countries[countries["gdp_usd"] <= state_value].sort_values("gdp_usd", ascending=False).head(1)
    return above.iloc[0], below.iloc[0]


def write_findings(year, year_notes, combined, countries, states, checks, pairs, titles,
                   bea_csv_name, bea_last_updated, wb_lastupdated, lacking):
    n = len(countries)
    s = states.set_index("name")
    fifty = states[states["name"] != "District of Columbia"]
    lines = []
    add = lines.append

    def place(name):
        v = s.loc[name, "gdp_usd"]
        r = int(s.loc[name, "rank_among_countries"])
        up, down = neighbours(v, countries)
        return v, r, up, down

    ca_v, ca_r, ca_up, ca_dn = place("California")
    tx_v, tx_r, tx_up, tx_dn = place("Texas")
    ny_v, ny_r, ny_up, ny_dn = place("New York")
    fl_v, fl_r, fl_up, fl_dn = place("Florida")
    bigger_than_ca = countries[countries["gdp_usd"] > ca_v].sort_values("gdp_usd", ascending=False)
    top50 = int((fifty["rank_among_countries"] <= 50).sum())
    top25 = int((fifty["rank_among_countries"] <= 25).sum())
    small = pairs[0]
    smaller_than_small = int((countries["gdp_usd"] < small["state_usd"]).sum())
    worst_rank = int(states["rank_among_countries"].max())

    add("# State GDP vs. countries: findings")
    add("")
    add(f"_Generated by `analysis.py` on {dt.date.today():%B %d, %Y}. Every number below is computed from the "
        "downloaded publisher files in `raw/`; rerun the script to rebuild this page._")
    add("")
    add("## Year used")
    add("")
    add(f"**{year}.** It is the latest year both publishers cover. BEA's SAGDP2 file (`{bea_csv_name}`) runs "
        f"through {year} (\"{bea_last_updated}\"). The World Bank series NY.GDP.MKTP.CD has {year} values for "
        f"{n} of the {n + len(lacking)} non-aggregate economies "
        f"(API `lastupdated` {wb_lastupdated}). Year test, newest first: " + "; ".join(year_notes) + ".")
    add("")
    add("## The answer")
    add("")
    add(f"**If each state were its own country, {top50} of the 50 would rank among the world's 50 largest "
        f"economies, led by California at {trillions(ca_v)}, smaller only than "
        f"{join_and([in_sentence(x) for x in bigger_than_ca['name']])}.**")
    add("")
    add("## Findings")
    add("")
    add(f"- California's {trillions(ca_v)} economy would rank {ordinal(ca_r)} among the {n} countries with "
        f"{year} data, between {in_sentence(ca_up['name'])} ({trillions(ca_up['gdp_usd'])}) and "
        f"{in_sentence(ca_dn['name'])} ({trillions(ca_dn['gdp_usd'])}).")
    add(f"- Texas, at {trillions(tx_v)}, would rank {ordinal(tx_r)}, between {in_sentence(tx_up['name'])} "
        f"({trillions(tx_up['gdp_usd'])}) and {in_sentence(tx_dn['name'])} ({trillions(tx_dn['gdp_usd'])}).")
    add(f"- New York, at {trillions(ny_v)}, would rank {ordinal(ny_r)}, between {in_sentence(ny_up['name'])} "
        f"({trillions(ny_up['gdp_usd'])}) and {in_sentence(ny_dn['name'])} ({trillions(ny_dn['gdp_usd'])}).")
    add(f"- Florida, at {trillions(fl_v)}, would rank {ordinal(fl_r)}, squeezed between "
        f"{in_sentence(fl_up['name'])} ({trillions(fl_up['gdp_usd'])}) and {in_sentence(fl_dn['name'])} "
        f"({trillions(fl_dn['gdp_usd'])}).")
    add(f"- {top50} of the 50 states would rank in the world's top 50 economies, and {top25} would rank in the "
        f"top 25 (each state compared with countries only, the United States included).")
    add(f"- {small['state']}, the smallest state economy at {billions(small['state_usd'])}, is closest in size "
        f"to {in_sentence(small['country'])} ({billions(small['country_usd'])}) and would still rank "
        f"{ordinal(small['rank'])} of {n}, larger than {smaller_than_small} countries; no state or DC ranks "
        f"below {ordinal(worst_rank)}.")
    add("")
    add("### How to read the ranks")
    add("")
    add(f"- `rank_among_countries` = 1 + the number of countries with a larger {year} GDP. States are slotted "
        "into the country list one at a time, so states never push each other down, and the United States "
        "itself counts as a country (rank 1). Ties would share a rank.")
    add(f"- \"Countries\" are the {n} World Bank economies (not aggregates) with a {year} value. The World "
        "Bank calls them economies; the list includes some territories, such as Hong Kong SAR (China) and "
        "Puerto Rico.")
    add("- The `state` type in `combined_ranking.csv` covers the 50 states and the District of Columbia "
        "(BEA's state-level areas). Findings about \"the 50 states\" leave DC out.")
    add("- Names: the CSV keeps World Bank names. The charts and this page use short names: "
        + "; ".join(f"{k} → {v}" for k, v in DISPLAY_NAMES.items() if k != v) + ".")
    add("")
    add("### Numbers behind the findings")
    add("")
    add("| Place | Type | GDP (US$) | As published | Rank among countries |")
    add("|---|---|---:|---|---:|")
    named = ["California", "Texas", "New York", "Florida"] + [p["state"] for p in pairs]
    rows = []
    for name in dict.fromkeys(named):
        r = s.loc[name]
        rows.append((name, "state", r["gdp_usd"], f"BEA, ${r['gdp_usd'] / 1e6:,.1f} million", int(r["rank_among_countries"])))
    cnames = list(bigger_than_ca["name"]) + [x["name"] for x in (ca_up, ca_dn, tx_up, tx_dn, ny_up, ny_dn, fl_up, fl_dn)] + [p["country"] for p in pairs]
    c = countries.set_index("name")
    for name in dict.fromkeys(cnames):
        r = c.loc[name]
        rows.append((display(name), "country", r["gdp_usd"], "World Bank, US$ (as shown)", int(r["rank_among_countries"])))
    for name, t, v, src, rk in sorted(rows, key=lambda x: -x[2]):
        add(f"| {name} | {t} | {v:,.0f} | {src} | {rk} |")
    add("")
    add("### Smallest states and their closest country")
    add("")
    add("| State | State GDP | Closest country | Country GDP | Difference |")
    add("|---|---:|---|---:|---:|")
    for p in pairs:
        d = p["country_usd"] - p["state_usd"]
        add(f"| {p['state']} | {billions(p['state_usd'])} | {display(p['country'])} | {billions(p['country_usd'])} | "
            f"{'+' if d >= 0 else '−'}{billions(abs(d))} ({d / p['state_usd']:+.1%}) |")
    add("")
    add("## Checks")
    add("")
    for ch in checks:
        add(f"- **({ch['id']}) {ch['name']}: {'PASS' if ch['pass'] else 'FAIL'}.** {ch['detail']}")
    add("")
    add("## Limits")
    add("")
    add("- **Exchange rates, not purchasing power.** Country GDP is converted to dollars at market exchange "
        "rates, so a currency swing moves a country up or down with no change in what it produces. Measured "
        "at purchasing power parity, many lower-income countries would look larger. States are already in dollars.")
    add(f"- **Estimates get revised, and some are missing.** BEA revised its {year} state figures on September 30, "
        f"2026 and will revise them again; World Bank {year} values are first estimates and will also change. {len(lacking)} "
        f"economies have no {year} value and are left out, including the United Arab Emirates "
        "(see check e for what that changes).")
    add("- **States are not separate economies.** They share one currency, one central bank and one federal "
        "budget, and trade freely with each other. The comparison is about size, not about what a state could "
        "do on its own.")
    add("- **Two publishers, two vintages.** State values are BEA's latest; country values, including the "
        "United States, are the World Bank's. The two U.S. totals differ slightly (check b). State GDP also "
        "leaves out U.S. activity abroad (check a).")
    add("")
    add("## Credits")
    add("")
    m = re.search(r"Last updated: ([A-Z][a-z]+ \d{1,2}, \d{4})", bea_last_updated)
    bea_release = m.group(1) if m else "date not found in file"
    add("**U.S. Bureau of Economic Analysis** (public domain; BEA's FAQ asks for a credit such as "
        "\"Source: U.S. Bureau of Economic Analysis\", and its citing guidelines ask for table name, URL, "
        "access date and the vintage of the estimate):")
    add("")
    add("- Short (chart or caption): Source: U.S. Bureau of Economic Analysis")
    add(f"- Full: U.S. Bureau of Economic Analysis, \"SAGDP2 Gross Domestic Product (GDP) by state: All "
        f"industry total, millions of current dollars,\" {year} estimates released {bea_release}, "
        f"https://apps.bea.gov/regional/zip/SAGDP.zip (accessed {file_date('SAGDP.zip'):%B %d, %Y}).")
    add("")
    add("**World Bank** (CC BY 4.0; the World Bank's Terms of Use for Datasets ask for the format "
        "\"The World Bank: Dataset name: Data source (if known)\", and CC BY 4.0 requires saying what was changed; "
        "use must not imply World Bank endorsement):")
    add("")
    add("- Short (chart or caption): World Bank, World Development Indicators (CC BY 4.0)")
    add("- Full: The World Bank: World Development Indicators, GDP (current US$) (NY.GDP.MKTP.CD): Country "
        "official statistics, National Statistical Organizations and/or Central Banks; National Accounts data "
        "files, Organisation for Economic Co-operation and Development (OECD); Staff estimates, World Bank (WB). "
        f"Accessed {file_date('wb_gdp.json'):%B %d, %Y}. Licensed under CC BY 4.0 "
        "(https://creativecommons.org/licenses/by/4.0/). Changes: combined with BEA state estimates and re-ranked.")
    add("")
    add("## Source URLs used")
    add("")
    for key, (url, filename) in SOURCES.items():
        add(f"- {url}  \n  saved as `raw/{filename}`, {file_date(filename):%Y-%m-%d %H:%M} local time, "
            f"SHA-256 `{sha256(RAW / filename)[:16]}…`")
    add(f"- {BEA_PUBLISHED_CA['url']} (California fact sheet, Gross Domestic Product tab; read "
        f"{BEA_PUBLISHED_CA['checked']} for check c)")
    add("- https://www.bea.gov/news/2026/gdp-third-estimate-industries-corporate-profits-state-gdp-and-state-personal-income-2nd "
        "(release of September 30, 2026)")
    add("- https://www.bea.gov/help/faq/147 and https://www.bea.gov/help/guidelines-for-citing-bea (BEA citation guidance)")
    add("- https://www.worldbank.org/en/about/legal/terms-of-use-for-datasets (World Bank attribution format, CC BY 4.0)")
    add("")
    add("## Chart titles")
    add("")
    for t in titles:
        add(f"- {t}")
    add("")
    (OUT / "findings.md").write_text("\n".join(lines), encoding="utf-8")


def file_date(filename):
    return dt.datetime.fromtimestamp((RAW / filename).stat().st_mtime)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    parser.add_argument("--refresh", action="store_true", help="re-download every source file")
    args = parser.parse_args()

    print("1. Downloading")
    download_all(args.refresh)

    print("2. Reading sources")
    bea_lines, bea_total, bea_csv_name, bea_last_updated = read_bea_annual()
    wb_rows, wb_meta, wb_lastupdated = read_wb()

    print("3. Choosing the year")
    year, year_notes = choose_year(bea_total, wb_rows, wb_meta)
    for note in year_notes:
        print("   ", note)
    print(f"   -> using {year}")

    print("4. Building the ranking")
    combined, countries, states, info = build_tables(year, bea_total, wb_rows, wb_meta)
    out_cols = ["name", "type", "gdp_usd", "rank_among_countries", "code"]
    csv = combined[out_cols].copy()
    csv["gdp_usd"] = csv["gdp_usd"].round(0).astype("int64")
    csv.to_csv(OUT / "combined_ranking.csv", index=False, encoding="utf-8")
    print(f"   {len(countries)} countries + {len(states)} states/DC -> combined_ranking.csv")

    print("5. Running checks")
    checks = run_checks(year, combined, countries, states, info, bea_lines)
    for ch in checks:
        print(f"   ({ch['id']}) {'PASS' if ch['pass'] else 'FAIL'}  {ch['name']}")

    print("6. Drawing charts")
    # Five smallest states (DC excluded), each with the country nearest in GDP.
    fifty = states[states["name"] != "District of Columbia"].sort_values("gdp_usd")
    pairs = []
    others = countries[countries["code"] != "USA"]
    for r in fifty.head(5).itertuples():
        nearest = others.iloc[(others["gdp_usd"] - r.gdp_usd).abs().argsort().iloc[0]]
        pairs.append({"state": r.name, "state_usd": r.gdp_usd, "rank": r.rank_among_countries,
                      "country": nearest["name"], "country_usd": nearest["gdp_usd"]})
    us_wb = float(countries.loc[countries["code"] == "USA", "gdp_usd"].iloc[0])
    t1 = chart_top(combined, year, us_wb, OUT / "chart_top_economies.png")
    t2 = chart_small(pairs, year, len(countries), OUT / "chart_small_states.png")

    print("7. Writing findings.md")
    write_findings(year, year_notes, combined, countries, states, checks, pairs, [t1, t2],
                   bea_csv_name, bea_last_updated, wb_lastupdated, info["lacking"])

    failed = [c["id"] for c in checks if not c["pass"]]
    print("Done." if not failed else f"Done, but checks FAILED: {', '.join(failed)}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
