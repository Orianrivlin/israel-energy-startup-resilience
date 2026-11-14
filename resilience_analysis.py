"""Investment-based resilience of the Israeli energy-tech startup sector:
exogenous versus endogenous shocks - analysis code.

This script produces every table and figure of the results in the paper
"Investment-Based Resilience of the Israeli Energy Startup Sector: A Comparative
Analysis of Exogenous and Endogenous Shocks".

What it does
  1. Reads three IVC Research Center exports (financing rounds, the company
     population and exits). The data are commercial and are NOT included in
     this repository; see README.md for the expected files and columns.
  2. Builds the quarterly normalized performance series P(t) = Y(t) / S(t),
     where Y(t) is the number of financing rounds in quarter t and S(t) is the
     stock of active HT companies (S(y) = S(y-1) + founded - ceased - acquired).
  3. Computes the resilience-curve metrics M1-M6 for the six selected shocks
     (three exogenous, three endogenous).
  4. Runs the in-time placebo test (B = 500 draws, fixed seed) and the exact
     permutation test of the endogenous-exogenous difference (all 20 splits).
  5. Runs the robustness checks: reported amounts instead of deal counts, clean
     core of four shocks, fixed 8-quarter windows and half-year aggregation.
  6. Writes three Excel workbooks (with editable charts), PNG figures and a
     JSON dump of all results to the output folder.

Usage
  python resilience_analysis.py --data <folder with the three IVC .xlsx files> --out <output folder>
"""
import argparse
import json
import textwrap
import warnings
from itertools import combinations
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import AreaChart, BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint, Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RegularTextRun
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# Input and output folders; set from the command line (see parse_args at the end of the file).
DATA = None   # folder with the three IVC .xlsx files
OUT = None    # output folder
FIG = None    # OUT / "figures"

QUARTERS = pd.period_range("2013Q1", "2026Q1", freq="Q")
LAST_USABLE = pd.Period("2023Q3", "Q")  # last quarter before the Iron Swords War
MAX_BASE = 8      # maximum pre-shock baseline quarters
MIN_BASE = 2      # minimum clean baseline quarters for a (pseudo-)shock
MAX_POST = 8      # maximum post-shock quarters for M6
REC = 0.95        # recovery threshold for M3
B = 500           # placebo draws
SEED = 20261030

# Full event list. Selected events are analysed; all events (selected or not)
# define the quarters excluded from baselines, post windows and placebo onsets.
EVENTS = [
    # code, name, class, direction, t0, t1, selected
    ("EX1", "Operation Protective Edge", "Exogenous", "Adverse", "2014Q3", "2014Q3", True),
    ("PARIS", "Paris Agreement (adoption and signing)", "Exogenous", "Positive", "2015Q4", "2016Q2", False),
    ("EX2", "COVID-19 initial outbreak and first lockdown", "Exogenous", "Adverse", "2020Q1", "2020Q2", True),
    ("GOTW", "Operation Guardian of the Walls", "Exogenous", "Adverse", "2021Q2", "2021Q2", False),
    ("EX3", "Global monetary tightening and Russia-Ukraine macro shock", "Exogenous", "Adverse", "2022Q1", "2022Q4", True),
    ("JUD", "Judicial overhaul crisis", "Exogenous", "Adverse", "2023Q1", "2023Q3", False),
    ("SWORDS", "Iron Swords War (ongoing)", "Exogenous", "Adverse", "2023Q4", "2026Q1", False),
    ("EN1", "Natural Gas Framework approval and HCJ invalidation", "Endogenous", "Adverse", "2015Q3", "2016Q2", True),
    ("EN2", "Leviathan Final Investment Decision (FID)", "Endogenous", "Adverse", "2017Q1", "2017Q1", True),
    ("REFORM", "Electricity market reform (Gov. Decision 3859)", "Endogenous", "Positive", "2018Q2", "2018Q3", False),
    ("LEVPROD", "Leviathan commercial production", "Endogenous", "Adverse", "2019Q4", "2020Q1", False),
    ("GD465", "Government Decision 465 (30% renewables target)", "Endogenous", "Positive", "2020Q4", "2020Q4", False),
    ("GRID", "Transmission grid bottleneck (onset: IEC capacity maps)", "Endogenous", "Adverse", "2020Q4", "2021Q1", False),
    ("STORAGE", "Storage integration regulation and first tenders", "Endogenous", "Positive", "2022Q1", "2022Q4", False),
    ("EN3", "Karish commercial gas production", "Endogenous", "Adverse", "2022Q4", "2022Q4", True),
]
CLEAN_CORE = ["EX1", "EX2", "EN1", "EN2"]
SHORT_NAMES = {
    "EX1": "Operation Protective Edge", "EX2": "COVID-19 first lockdown", "EX3": "Monetary tightening & Ukraine",
    "EN1": "Natural Gas Framework", "EN2": "Leviathan FID", "EN3": "Karish production",
}
# Sources of the metric definitions (full references in the paper).
METRIC_SOURCES = {
    "P(t), P~(t) = P(t)/P_base": "Normalized performance p = P/R with a fixed reference (static normalization): "
                                 "Poulin & Kane (2021), Sec. 3.3",
    "M1": "Depth of impact J = 1 - p (residual performance = min p): Poulin & Kane (2021), Sec. 4.2; "
          "cf. resistance as maximum fall, Iacobucci & Perugini (2021)",
    "M2": "Duration-based metric (onset to lowest performance): Poulin & Kane (2021), Sec. 4.3",
    "M3": "Recovery duration defined by a threshold (e.g. 95%): Poulin & Kane (2021), Sec. 4.3; "
          "95% restoration example: Bruneau et al. (2004), Table A2",
    "M4": "Recovery rate, linear approximation: Poulin & Kane (2021), Sec. 4.5",
    "M5": "Loss of resilience R = integral[100 - Q(t)]dt, adapted to a quarterly sum: Bruneau et al. (2004); "
          "'cumulative impact': Poulin & Kane (2021), Sec. 4.4",
    "M6": "Restored performance relative to pre-event level: Poulin & Kane (2021), Sec. 4.2; "
          "renewal (resumption of the pre-shock path or hysteretic shift): Martin (2012), pp. 8-9, 11-12",
    "Placebo test": "In-time placebo and p-value as share of placebo effects >= actual: Abadie et al. (2015), pp. 499-500",
    "Group permutation test": "Exact permutation (randomization) inference over all 20 splits of the six shocks: "
                              "Abadie et al. (2015), p. 500",
}
PERM_METRICS = ["M1", "M5"]  # M1 primary, M5 secondary; M6 is defined for only two shocks per group

INPUT_FILES = ["Energy High-Tech Capital Raising.xlsx", "Energy High-Tech Companies & MNCs.xlsx",
               "Energy High-Tech Exits.xlsx"]

POS = {q: i for i, q in enumerate(QUARTERS)}


def qidx(label):
    return POS[pd.Period(label, "Q")]


class Grid:
    """Time grid (quarters or half-years) on which the series and metrics are computed."""

    def __init__(self, labels, period_of, last_usable, max_base, max_post, fixed_len=None, unit="quarters"):
        self.labels = labels            # period labels, e.g. "2014Q3" or "2014H2"
        self.period_of = period_of      # quarter label -> index on this grid
        self.last = last_usable         # last usable period index (before the Iron Swords War)
        self.max_base = max_base
        self.max_post = max_post
        self.fixed_len = fixed_len      # None = window ends before the next selected shock
        self.unit = unit

    def event_periods(self, exclude=None):
        """Indices of periods inside any listed event window [t0, t1], except `exclude`."""
        out = set()
        for code, _, _, _, t0, t1, _ in EVENTS:
            if code != exclude:
                out.update(range(self.period_of(t0), self.period_of(t1) + 1))
        return out


def half_of(label):
    q = pd.Period(label, "Q")
    return (q.year - QUARTERS[0].year) * 2 + (q.quarter - 1) // 2


HALVES = [f"{y}H{h}" for y in range(QUARTERS[0].year, QUARTERS[-1].year + 1) for h in (1, 2)][:half_of(str(QUARTERS[-1])) + 1]
QUARTER_GRID = Grid([str(q) for q in QUARTERS], qidx, POS[LAST_USABLE], MAX_BASE, MAX_POST)
FIXED_GRID = Grid([str(q) for q in QUARTERS], qidx, POS[LAST_USABLE], MAX_BASE, MAX_POST, fixed_len=8)
HALF_GRID = Grid(HALVES, half_of, half_of("2023Q2"), MAX_BASE // 2, MAX_POST // 2, unit="half-years")


def event_quarters(exclude=None):
    return QUARTER_GRID.event_periods(exclude)


# ----------------------------------------------------------------------------
# Data
# ----------------------------------------------------------------------------
def load_data():
    deals = pd.read_excel(DATA / "Energy High-Tech Capital Raising.xlsx")
    deals = deals.dropna(subset=["Deal Date"]).copy()
    deals["Quarter"] = deals["Deal Date"].dt.to_period("Q")

    # The last row of each IVC export holds the applied filters; it has no founding year.
    comp = pd.read_excel(DATA / "Energy High-Tech Companies & MNCs.xlsx")
    comp.columns = [c.strip() for c in comp.columns]
    comp = comp.dropna(subset=["Est. Year"]).drop_duplicates("Company ID").copy()

    exits = pd.read_excel(DATA / "Energy High-Tech Exits.xlsx")
    exits.columns = [c.strip() for c in exits.columns]
    exits = exits.dropna(subset=["Exit Date"]).copy()
    return deals, comp, exits


def build_population(comp, exits):
    """HT companies (the financing file contains HT companies only) with their exit year from the stock."""
    pop = comp[comp["Company Type"] == "HT"].copy()
    first_acq = exits[exits["Exit Type"] == "Acquired"].groupby("Company ID")["Exit Date"].min()
    is_acq = pop["Company Status"] == "Acquired/Merged"
    pop["Acquired Year"] = np.where(is_acq, pop["Company ID"].map(first_acq).dt.year, np.nan)
    info = dict(all_companies=len(comp), mnc_excluded=int((comp["Company Type"] == "MNC").sum()),
                ht_companies=len(pop), acquired_status=int(is_acq.sum()),
                acquired_with_date=int(pop["Acquired Year"].notna().sum()),
                ceased_with_year=int(pop["Ceased Year"].notna().sum()))
    return pop, info


def build_series(deals, pop):
    years = range(QUARTERS[0].year, QUARTERS[-1].year + 1)
    est = pop["Est. Year"].astype(int)
    ceased = pop["Ceased Year"].dropna().astype(int)
    acquired = pop["Acquired Year"].dropna().astype(int)
    founded = {y: int((est == y).sum()) for y in years}
    closed = {y: int((ceased == y).sum()) for y in years}
    bought = {y: int((acquired == y).sum()) for y in years}
    stock, s = {}, 0
    for y in years:  # S(y) = S(y-1) + founded(y) - ceased(y) - acquired(y), year-end stock
        s = s + founded[y] - closed[y] - bought[y]
        stock[y] = s

    n = deals.groupby("Quarter").size().reindex(QUARTERS, fill_value=0)
    amt = deals.groupby("Quarter")["Deal Amount ($m)"].sum(min_count=1).reindex(QUARTERS).fillna(0.0)
    n_amt = deals.dropna(subset=["Deal Amount ($m)"]).groupby("Quarter").size().reindex(QUARTERS, fill_value=0)
    ser = pd.DataFrame({
        "Quarter": [str(q) for q in QUARTERS],
        "Deals": n.values,
        "Deals with reported amount": n_amt.values,
        "Reported amount ($m)": amt.values.round(3),
        "Active stock S(t)": [stock[q.year] for q in QUARTERS],
    })
    ser["P(t) deals"] = ser["Deals"] / ser["Active stock S(t)"]
    ser["P(t) amount"] = ser["Reported amount ($m)"] / ser["Active stock S(t)"]
    annual = pd.DataFrame({"Year": list(years), "Founded": [founded[y] for y in years],
                           "Ceased": [closed[y] for y in years], "Acquired": [bought[y] for y in years],
                           "Active stock (year-end)": [stock[y] for y in years]})
    return ser, annual


def build_half_series(deals, annual):
    half = deals["Quarter"].astype(str).map(half_of)
    n = half.value_counts().reindex(range(len(HALVES)), fill_value=0)
    stock = dict(zip(annual["Year"], annual["Active stock (year-end)"]))
    hs = pd.DataFrame({"Half-year": HALVES, "Deals": n.values,
                       "Active stock S(t)": [stock[int(h[:4])] for h in HALVES]})
    hs["P(t) deals"] = hs["Deals"] / hs["Active stock S(t)"]
    return hs


# ----------------------------------------------------------------------------
# Metrics
# ----------------------------------------------------------------------------
def curve_metrics(p, t0, tend, base_idx):
    """M1-M5 for a window [t0, tend] given baseline quarter indices."""
    p_base = float(np.mean([p[i] for i in base_idx]))
    rel = {t: p[t] / p_base for t in range(t0, tend + 1)}
    t_trough = min(rel, key=lambda t: (rel[t], t))
    m1 = 1 - rel[t_trough]
    if rel[t_trough] >= REC:
        status, t_rec = "No contraction below 95%", t0
        m2, m3, m4 = 0, 0, None
    else:
        t_rec = next((t for t in range(t_trough + 1, tend + 1) if rel[t] >= REC), None)
        m2 = t_trough - t0
        if t_rec is None:
            status, m3, m4 = "Censored (not recovered in window)", None, None
        else:
            status = "Recovered"
            m3 = t_rec - t0
            m4 = (rel[t_rec] - rel[t_trough]) / (t_rec - t_trough)
    last = t_rec if status == "Recovered" or status.startswith("No") else tend
    m5 = sum(max(0.0, 1 - rel[t]) for t in range(t0, last + 1))
    return dict(p_base=p_base, rel=rel, t_trough=t_trough, t_rec=t_rec if status == "Recovered" else None,
                M1=m1, M2=m2, M3=m3, M4=m4, M5=m5, status=status)


def analyse_events(p, grid=QUARTER_GRID):
    sel = [e for e in EVENTS if e[6]]
    sel_sorted = sorted(sel, key=lambda e: grid.period_of(e[4]))
    lab = grid.labels
    results = []
    for k, (code, name, cls, _, t0s, t1s, _) in enumerate(sel_sorted):
        t0, t1 = grid.period_of(t0s), grid.period_of(t1s)
        excl = grid.event_periods(exclude=code)
        if grid.fixed_len:
            tend = min(t0 + grid.fixed_len - 1, grid.last)
        else:
            nxt = grid.period_of(sel_sorted[k + 1][4]) - 1 if k + 1 < len(sel_sorted) else grid.last
            tend = min(nxt, grid.last)
        cand = [i for i in range(t0 - grid.max_base, t0) if i >= 0]
        base = [i for i in cand if i not in excl]
        post = [i for i in range(t1 + 1, t1 + 1 + grid.max_post) if i <= tend and i not in excl]
        m = curve_metrics(p, t0, tend, base)
        m6 = (np.mean([p[i] for i in post]) / m["p_base"]) if post else None
        concurrent = [e[0] for e in EVENTS if e[0] != code and
                      max(grid.period_of(e[4]), t0 - grid.max_base) <= min(grid.period_of(e[5]), tend)]
        results.append(dict(code=code, name=name, cls=cls, t0=t0s, t1=t1s, t0_label=lab[t0], tend=lab[tend],
                            window_len=tend - t0 + 1,
                            base=[lab[i] for i in base],
                            base_excluded=[lab[i] for i in cand if i in excl],
                            post=[lab[i] for i in post], M6=m6, concurrent=concurrent,
                            **{k2: v for k2, v in m.items() if k2 != "rel"},
                            t_rec_label=lab[m["t_rec"]] if m["t_rec"] is not None else None,
                            rel={lab[t]: v for t, v in m["rel"].items()},
                            t0_idx=t0, tend_idx=tend))
    return results


def placebo(p, results, rng, grid=QUARTER_GRID):
    allq = grid.event_periods()
    onsets_all = [s for s in range(len(grid.labels)) if s not in allq and s <= grid.last]
    out = {}
    for r in results:
        L = r["window_len"]
        draws = []
        for s in onsets_all:
            if s + L - 1 > grid.last:
                continue
            base = [i for i in range(s - grid.max_base, s) if i >= 0 and i not in allq]
            if len(base) < MIN_BASE:
                continue
            m = curve_metrics(p, s, s + L - 1, base)
            draws.append((grid.labels[s], m["M1"], m["M5"]))
        if not draws:
            out[r["code"]] = dict(n_onsets=0, onsets=[], p_M1=None, p_M5=None, M1_draws=[], M5_draws=[])
            continue
        pick = rng.integers(0, len(draws), size=B)
        m1d = np.array([draws[i][1] for i in pick])
        m5d = np.array([draws[i][2] for i in pick])
        out[r["code"]] = dict(n_onsets=len(draws), onsets=[d[0] for d in draws],
                              p_M1=float(np.mean(m1d >= r["M1"])), p_M5=float(np.mean(m5d >= r["M5"])),
                              M1_draws=m1d.tolist(), M5_draws=m5d.tolist(),
                              onset_values=[(d[0], d[1], d[2]) for d in draws])
    return out


def group_summary(results, codes=None, unit="quarters"):
    rows = []
    for cls in ["Exogenous", "Endogenous"]:
        g = [r for r in results if r["cls"] == cls and (codes is None or r["code"] in codes)]
        rec = [r for r in g if r["M3"] is not None]
        m6 = [r["M6"] for r in g if r["M6"] is not None]
        rows.append({
            "Group": cls, "Events": ", ".join(r["code"] for r in g), "N": len(g),
            "Mean M1 (depth)": np.mean([r["M1"] for r in g]),
            f"Mean M2 ({unit} to trough)": np.mean([r["M2"] for r in g]),
            "Recovered (n/N)": f"{len(rec)}/{len(g)}",
            f"Mean M3 of recovered ({unit})": np.mean([r["M3"] for r in rec]) if rec else None,
            "Mean M5 (cumulative loss)": np.mean([r["M5"] for r in g]),
            "Mean M6 (renewal)": np.mean(m6) if m6 else None,
        })
    return pd.DataFrame(rows)


def group_permutation(results, metric):
    """Exact permutation test of the endogenous-minus-exogenous difference in group means.

    Every split of the six shocks into two groups of three is a possible labelling under the null hypothesis that
    the label is unrelated to the metric. p_two counts splits with |difference| >= observed; p_one counts splits with
    difference >= observed (endogenous larger)."""
    vals = {r["code"]: r[metric] for r in results}
    codes = list(vals)
    endo = {r["code"] for r in results if r["cls"] == "Endogenous"}
    k = len(endo)

    def diff(group):
        return (np.mean([vals[c] for c in group]) - np.mean([vals[c] for c in codes if c not in group]))

    obs = diff(endo)
    diffs = [diff(set(g)) for g in combinations(codes, k)]
    tol = 1e-12
    return {"Metric": metric, "Observed difference (Endo - Exo)": obs, "Splits": len(diffs),
            "p one-sided (Endo > Exo)": float(np.mean([d >= obs - tol for d in diffs])),
            "p two-sided": float(np.mean([abs(d) >= abs(obs) - tol for d in diffs])),
            "Minimum attainable two-sided p": 2 / len(diffs),
            "Rank of endogenous group": int(sum(d > obs + tol for d in diffs)) + 1}


# ----------------------------------------------------------------------------
# Output helpers
# ----------------------------------------------------------------------------
HEAD_FILL = PatternFill("solid", fgColor="D9E1F2")
SOURCE_NOTE = "Source: Author(s), 2026."


def write_df(ws, df, r0=1, c0=1, title=None, note=SOURCE_NOTE, fmt=None):
    r = r0
    if title:
        ws.cell(r, c0, title).font = Font(bold=True, size=12)
        r += 1
    for j, col in enumerate(df.columns):
        c = ws.cell(r, c0 + j, col)
        c.font, c.fill = Font(bold=True), HEAD_FILL
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for i, row in enumerate(df.itertuples(index=False), start=1):
        for j, v in enumerate(row):
            if isinstance(v, (np.floating, float)) and np.isnan(v):
                v = None
            if isinstance(v, np.generic):
                v = v.item()
            c = ws.cell(r + i, c0 + j, "n/a" if v is None else v)
            if isinstance(v, float):
                c.number_format = (fmt or {}).get(df.columns[j], "0.000")
    end = r + len(df)
    if note:
        ws.cell(end + 1, c0, note).font = Font(italic=True, size=9)
    for j, col in enumerate(df.columns):
        letter = get_column_letter(c0 + j)
        ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width or 0, min(48, max(12, len(str(col)) + 2)))
    return end + 3


def fmt_num(v, nd=3):
    return "n/a" if v is None else round(float(v), nd)


def metrics_table(results, pl, version, unit="quarters"):
    rows = []
    for r in results:
        rows.append({
            "Code": r["code"], "Event": r["name"], "Class": r["cls"], "t0": r["t0_label"], "t1": r["t1"],
            "Window end": r["tend"], f"Baseline {unit} (n)": len(r["base"]),
            "P_base": r["p_base"], "M1 Shock depth": r["M1"], f"M2 {unit.capitalize()} to trough": r["M2"],
            f"M3 Recovery time ({unit})": r["M3"] if r["M3"] is not None else "Censored",
            "M4 Recovery velocity": r["M4"] if r["M4"] is not None else "n/a",
            "M5 Cumulative loss": r["M5"],
            "M6 Renewal (post-shock level)": r["M6"] if r["M6"] is not None else "n/a",
            "Status": r["status"],
            "Placebo p (M1)": pl[r["code"]]["p_M1"], "Placebo p (M5)": pl[r["code"]]["p_M5"],
            "Placebo onsets (distinct)": pl[r["code"]]["n_onsets"],
            "Y(t) version": version,
        })
    return pd.DataFrame(rows)


# Figures are defined once (panel data + titles) and drawn twice: as native,
# editable Excel charts and as PNG previews with the same content and styling.
BLUE, ORANGE, GREY, LOSS = "1F4E79", "C55A11", "808080", "F4CCCC"
EXO_SHADE, ENDO_SHADE = "DCE6F2", "F8D7D7"
PANELS = [("EX1", "a"), ("EN1", "b"), ("EX2", "c"), ("EN2", "d"), ("EX3", "e"), ("EN3", "f")]
GROUP_COLS = [("Mean M1 (depth)", "(a) Mean M1 (shock depth)"),
              ("Mean M5 (cumulative loss)", "(b) Mean M5 (cumulative loss)")]
SERIES_TITLE = "Normalized financing activity P(t), 2013Q1-2026Q1"


def panel_titles(r, letter):
    status = r["status"].split(" (")[0]
    return (f"({letter}) {r['code']}: {SHORT_NAMES[r['code']]} (t0 = {r['t0']})",
            f"M1 = {r['M1']:.2f}   M5 = {r['M5']:.2f}   {status}")


def panel_rows(r):
    """Quarter, P~, baseline, threshold, hidden area base, loss area for t0..window end."""
    rel = r["rel"]
    qs = list(rel.keys())
    loss_end = qs.index(r["t_rec_label"]) if r["t_rec_label"] is not None else len(qs) - 1
    rows = []
    for i, q in enumerate(qs):
        v = rel[q]
        loss = (1 - v) if (v < 1 and i <= loss_end) else 0.0
        rows.append((q, round(v, 4), 1.0, REC, round(min(v, 1.0), 4), round(loss, 4)))
    return rows


def y_limits(values, include_one=True):
    lo = min(values + ([REC] if include_one else []))
    hi = max(values + ([1.0] if include_one else []))
    return np.floor((lo - 0.1) * 10) / 10, np.ceil((hi + 0.1) * 10) / 10


def plot_panel(r, letter, path):
    rows = panel_rows(r)
    qs = [x[0] for x in rows]
    ys = np.array([x[1] for x in rows])
    base = np.array([x[4] for x in rows])
    loss = np.array([x[5] for x in rows])
    xs = np.arange(len(qs))
    t1, t2 = panel_titles(r, letter)
    fig, ax = plt.subplots(figsize=(5.2, 3.3))
    # Same construction as the Excel stacked area: from min(P~, 1) up by the loss amount.
    ax.fill_between(xs, base, base + loss, color="#" + LOSS, label="Loss area (M5)", lw=0)
    ax.plot(xs, ys, marker="o", ms=4, c="#" + BLUE, lw=1.8, label="P~(t)")
    ax.plot(xs, [1.0] * len(xs), ls="--", c="#" + GREY, lw=1.2, label="Baseline (1.0)")
    ax.plot(xs, [REC] * len(xs), ls=":", c="#" + GREY, lw=1.2, label="Recovery threshold (0.95)")
    ax.set_ylim(*y_limits(list(ys)))
    ax.set_xticks(xs)
    ax.set_xticklabels(qs, rotation=45, fontsize=7)
    ax.tick_params(axis="y", labelsize=7)
    ax.set_ylabel("P~(t) relative to baseline", fontsize=7)
    ax.set_title(f"{t1}\n{t2}", fontsize=8)
    ax.legend(fontsize=6, loc="upper center", bbox_to_anchor=(0.5, -0.32), ncol=4, frameon=False)
    fig.tight_layout()
    fig.savefig(path, dpi=300)
    plt.close(fig)


def plot_groups(summary, path):
    fig, axes = plt.subplots(1, 2, figsize=(9, 3.4))
    for ax, (col, title) in zip(axes, GROUP_COLS):
        ax.bar(summary["Group"], summary[col], color=["#" + BLUE, "#" + ORANGE], width=0.55)
        ax.axhline(0, c="black", lw=0.8)
        ax.set_title(title, fontsize=10)
        ax.set_ylim(0, max(summary[col]) * 1.2)
        for i, v in enumerate(summary[col]):
            ax.text(i, v, f"{v:.2f}", ha="center", va="bottom", fontsize=9)
    fig.tight_layout()
    fig.savefig(path, dpi=300)
    plt.close(fig)


def shade_values(ymax):
    exo, endo = [], []
    for i in range(len(QUARTERS)):
        hit = [e for e in EVENTS if e[6] and qidx(e[4]) <= i <= qidx(e[5])]
        exo.append(ymax if any(e[2] == "Exogenous" for e in hit) else None)
        endo.append(ymax if any(e[2] == "Endogenous" for e in hit) else None)
    return exo, endo


def plot_series(ser, path):
    ymax = float(np.ceil(ser["P(t) deals"].max() * 100) / 100)
    exo, endo = shade_values(ymax)
    xs = np.arange(len(ser))
    fig, ax = plt.subplots(figsize=(10, 3.4))
    ax.bar(xs, [v or 0 for v in exo], width=1.0, color="#" + EXO_SHADE, label="Exogenous shock window")
    ax.bar(xs, [v or 0 for v in endo], width=1.0, color="#" + ENDO_SHADE, label="Endogenous shock window")
    ax.plot(xs, ser["P(t) deals"], c="#" + BLUE, marker="o", ms=3, lw=1.5, label="P(t) = rounds / active firms")
    for code, _, _, _, t0, _, sel in EVENTS:
        if sel:
            ax.text(qidx(t0), ymax * 0.97, code, fontsize=7, va="top")
    ax.set_ylim(0, ymax)
    ax.set_xlim(-0.5, len(xs) - 0.5)
    ax.set_xticks(xs[::4])
    ax.set_xticklabels(ser["Quarter"][::4], fontsize=8)
    ax.set_ylabel("P(t)", fontsize=8)
    ax.set_title(SERIES_TITLE, fontsize=10)
    ax.legend(fontsize=7, loc="upper center", bbox_to_anchor=(0.5, -0.12), ncol=3, frameon=False)
    fig.tight_layout()
    fig.savefig(path, dpi=300)
    plt.close(fig)


def chart_title(lines):
    paras = []
    for text, size, bold in lines:
        cp = CharacterProperties(sz=size * 100, b=bold)
        paras.append(Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=text)]))
    return Title(tx=Text(rich=RichText(p=paras)), overlay=False)


def line_props(color, dash=None, width_pt=1.75):
    return GraphicalProperties(ln=LineProperties(solidFill=color, prstDash=dash, w=int(width_pt * 12700)))


def show_axes(chart, y_fmt="0.0"):
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.number_format = y_fmt
    if chart.legend is not None:
        chart.legend.overlay = False


def number_series(*charts):
    """Give series unique idx/order across combined charts (needed for legend entries)."""
    k = 0
    for ch in charts:
        for s in ch.series:
            s.idx = s.order = k
            k += 1


def add_panel_chart(ws_data, ws_fig, col, r, letter, anchor):
    rows = panel_rows(r)
    # The stacked-area base series is invisible; its name is a blank so its legend slot stays empty
    # (Excel ignores legend-entry deletion written by openpyxl).
    heads = ["Quarter", "P~(t)", "Baseline (1.0)", "Recovery threshold (0.95)", " ", "Loss area (M5)"]
    ws_data.cell(1, col, f"Figure 1{letter}: {r['code']}")
    for j, h in enumerate(heads):
        ws_data.cell(2, col + j, h)
    for i, row in enumerate(rows, start=3):
        for j, v in enumerate(row):
            ws_data.cell(i, col + j, v)
    last = 2 + len(rows)
    cats = Reference(ws_data, min_col=col, min_row=3, max_row=last)

    area = AreaChart()
    area.grouping = "stacked"
    area.add_data(Reference(ws_data, min_col=col + 4, max_col=col + 5, min_row=2, max_row=last), titles_from_data=True)
    area.set_categories(cats)
    base_s, loss_s = area.series
    base_s.graphicalProperties = GraphicalProperties(noFill=True, ln=LineProperties(noFill=True))
    loss_s.graphicalProperties = GraphicalProperties(solidFill=LOSS, ln=LineProperties(noFill=True))

    line = LineChart()
    line.add_data(Reference(ws_data, min_col=col + 1, max_col=col + 3, min_row=2, max_row=last), titles_from_data=True)
    line.set_categories(cats)
    p_s, b_s, t_s = line.series
    p_s.graphicalProperties = line_props(BLUE, width_pt=2)
    p_s.marker = Marker(symbol="circle", size=6)
    p_s.marker.graphicalProperties = GraphicalProperties(solidFill=BLUE, ln=LineProperties(solidFill=BLUE))
    b_s.graphicalProperties = line_props(GREY, "dash", 1.25)
    t_s.graphicalProperties = line_props(GREY, "sysDot", 1.25)
    for s in (b_s, t_s):
        s.marker = Marker(symbol="none")
    for s in line.series:
        s.smooth = False

    number_series(area, line)
    area += line
    t1, t2 = panel_titles(r, letter)
    area.title = chart_title([(t1, 10, True), (t2, 9, False)])
    area.y_axis.title = chart_title([("P~(t) relative to baseline", 8, False)])
    lo, hi = y_limits([x[1] for x in rows])
    area.y_axis.scaling.min, area.y_axis.scaling.max = float(lo), float(hi)
    area.y_axis.majorGridlines = None
    area.legend.position = "b"
    area.height, area.width = 8.5, 14
    show_axes(area)
    ws_fig.add_chart(area, anchor)
    return col + len(heads) + 1


def add_group_charts(ws_data, ws_fig, col, summary):
    ws_data.cell(1, col, "Figure 2: group means (all six events)")
    ws_data.cell(2, col, "Group")
    for j, (src, _) in enumerate(GROUP_COLS, start=1):
        ws_data.cell(2, col + j, src)
    for i, row in enumerate(summary.to_dict("records"), start=3):
        ws_data.cell(i, col, row["Group"])
        for j, (src, _) in enumerate(GROUP_COLS, start=1):
            ws_data.cell(i, col + j, round(float(row[src]), 2))
    for j, (_, title) in enumerate(GROUP_COLS, start=1):
        bc = BarChart()
        bc.add_data(Reference(ws_data, min_col=col + j, min_row=2, max_row=4), titles_from_data=True)
        bc.set_categories(Reference(ws_data, min_col=col, min_row=3, max_row=4))
        s = bc.series[0]
        for idx, color in enumerate([BLUE, ORANGE]):
            pt = DataPoint(idx=idx)
            pt.graphicalProperties = GraphicalProperties(solidFill=color)
            s.dPt.append(pt)
        s.dLbls = DataLabelList(showVal=True, showSerName=False, showCatName=False, showLegendKey=False,
                                showPercent=False, showLeaderLines=False)
        s.dLbls.position = "outEnd"
        bc.title = chart_title([(title, 10, True)])
        bc.legend = None
        bc.gapWidth = 80
        bc.y_axis.scaling.min = 0
        bc.y_axis.majorGridlines = None
        bc.height, bc.width = 8.5, 14
        show_axes(bc)
        ws_fig.add_chart(bc, f"{'A' if j == 1 else 'J'}58")
    return col + len(GROUP_COLS) + 2


def add_series_chart(ws_data, ws_fig, col, ser):
    ymax = float(np.ceil(ser["P(t) deals"].max() * 100) / 100)
    exo, endo = shade_values(ymax)
    heads = ["Quarter", "Exogenous shock window", "Endogenous shock window", "P(t) = rounds / active firms"]
    ws_data.cell(1, col, "Figure 3: quarterly series")
    for j, h in enumerate(heads):
        ws_data.cell(2, col + j, h)
    for i, (q, a, b, p) in enumerate(zip(ser["Quarter"], exo, endo, ser["P(t) deals"]), start=3):
        ws_data.cell(i, col, q)
        ws_data.cell(i, col + 1, a or 0)
        ws_data.cell(i, col + 2, b or 0)
        ws_data.cell(i, col + 3, round(float(p), 5))
    last = 2 + len(ser)
    cats = Reference(ws_data, min_col=col, min_row=3, max_row=last)
    bars = BarChart()
    bars.grouping = "standard"
    bars.overlap = 100
    bars.gapWidth = 0
    bars.add_data(Reference(ws_data, min_col=col + 1, max_col=col + 2, min_row=2, max_row=last), titles_from_data=True)
    bars.set_categories(cats)
    bars.series[0].graphicalProperties = GraphicalProperties(solidFill=EXO_SHADE, ln=LineProperties(noFill=True))
    bars.series[1].graphicalProperties = GraphicalProperties(solidFill=ENDO_SHADE, ln=LineProperties(noFill=True))
    line = LineChart()
    line.add_data(Reference(ws_data, min_col=col + 3, min_row=2, max_row=last), titles_from_data=True)
    line.set_categories(cats)
    s = line.series[0]
    s.graphicalProperties = line_props(BLUE, width_pt=1.75)
    s.marker = Marker(symbol="circle", size=4)
    s.marker.graphicalProperties = GraphicalProperties(solidFill=BLUE, ln=LineProperties(solidFill=BLUE))
    s.smooth = False
    number_series(bars, line)
    bars += line
    bars.title = chart_title([(SERIES_TITLE, 10, True)])
    bars.y_axis.title = chart_title([("P(t)", 8, False)])
    bars.y_axis.scaling.min, bars.y_axis.scaling.max = 0.0, ymax
    bars.y_axis.majorGridlines = None
    bars.legend.position = "b"
    bars.height, bars.width = 8.5, 28
    show_axes(bars, "0.00")
    ws_fig.add_chart(bars, "A77")


# ----------------------------------------------------------------------------
# Robustness checks (separate workbook)
# ----------------------------------------------------------------------------
SPEC_LABELS = {"primary": "Primary (quarters, window to next shock)", "amount": "Reported amounts",
               "fixed8q": "Fixed 8-quarter window", "half": "Half-year aggregation"}


def plot_spec_comparison(comp_df, path):
    fig, axes = plt.subplots(1, 2, figsize=(10, 3.6))
    specs = list(comp_df["Specification"])
    xs = np.arange(len(specs))
    for ax, (metric, title) in zip(axes, [("M1", "(a) Mean M1 (shock depth)"), ("M5", "(b) Mean M5 (cumulative loss)")]):
        ax.bar(xs - 0.18, comp_df[f"Exogenous {metric}"], width=0.36, color="#" + BLUE, label="Exogenous")
        ax.bar(xs + 0.18, comp_df[f"Endogenous {metric}"], width=0.36, color="#" + ORANGE, label="Endogenous")
        for x, a, b in zip(xs, comp_df[f"Exogenous {metric}"], comp_df[f"Endogenous {metric}"]):
            ax.text(x - 0.18, a, f"{a:.2f}", ha="center", va="bottom", fontsize=7)
            ax.text(x + 0.18, b, f"{b:.2f}", ha="center", va="bottom", fontsize=7)
        ax.set_xticks(xs)
        ax.set_xticklabels([textwrap.fill(s, 16) for s in specs], fontsize=7)
        ax.set_title(title, fontsize=10)
        ax.set_ylim(0, max(comp_df[f"Exogenous {metric}"].max(), comp_df[f"Endogenous {metric}"].max()) * 1.25)
    axes[0].legend(fontsize=7, frameon=False)
    fig.tight_layout()
    fig.savefig(path, dpi=300)
    plt.close(fig)


def add_spec_charts(ws_data, ws_fig, comp_df):
    heads = ["Specification", "Exogenous M1", "Endogenous M1", "Exogenous M5", "Endogenous M5"]
    for j, h in enumerate(heads, start=1):
        ws_data.cell(1, j, h)
    for i, row in enumerate(comp_df.to_dict("records"), start=2):
        for j, h in enumerate(heads, start=1):
            v = row[h]
            ws_data.cell(i, j, round(float(v), 2) if not isinstance(v, str) else v)
    last = 1 + len(comp_df)
    for k, (metric, title, anchor) in enumerate([("M1", "(a) Mean M1 (shock depth)", "A4"),
                                                  ("M5", "(b) Mean M5 (cumulative loss)", "J4")]):
        bc = BarChart()
        bc.type = "col"
        col0 = 2 if metric == "M1" else 4
        bc.add_data(Reference(ws_data, min_col=col0, max_col=col0 + 1, min_row=1, max_row=last), titles_from_data=True)
        bc.set_categories(Reference(ws_data, min_col=1, min_row=2, max_row=last))
        for s, color in zip(bc.series, [BLUE, ORANGE]):
            s.graphicalProperties = GraphicalProperties(solidFill=color)
            s.dLbls = DataLabelList(showVal=True, showSerName=False, showCatName=False, showLegendKey=False,
                                    showPercent=False, showLeaderLines=False)
            s.dLbls.position = "outEnd"
        bc.title = chart_title([(title, 10, True)])
        bc.legend.position = "b"
        bc.gapWidth = 80
        bc.y_axis.scaling.min = 0
        bc.y_axis.majorGridlines = None
        bc.height, bc.width = 8.5, 14
        show_axes(bc)
        ws_fig.add_chart(bc, anchor)


def run_robustness(ser, annual, deals, res, pl, res_amt, summ_full, summ_core, summ_amt):
    rng = np.random.default_rng(SEED)
    p_q = ser["P(t) deals"].to_numpy()
    res_fix = analyse_events(p_q, FIXED_GRID)
    pl_fix = placebo(p_q, res_fix, rng, FIXED_GRID)
    hs = build_half_series(deals, annual)
    p_h = hs["P(t) deals"].to_numpy()
    res_half = analyse_events(p_h, HALF_GRID)
    pl_half = placebo(p_h, res_half, rng, HALF_GRID)
    rob = {
        "fixed8q": dict(results=res_fix, placebo=pl_fix, summary_full=group_summary(res_fix),
                        summary_core=group_summary(res_fix, CLEAN_CORE)),
        "half": dict(results=res_half, placebo=pl_half, summary_full=group_summary(res_half, unit="half-years"),
                     summary_core=group_summary(res_half, CLEAN_CORE, unit="half-years")),
    }

    # comparison across specifications (group means of M1 and M5, all six shocks)
    specs = [("primary", summ_full), ("amount", summ_amt), ("fixed8q", rob["fixed8q"]["summary_full"]),
             ("half", rob["half"]["summary_full"])]
    comp_rows = []
    for key, s in specs:
        g = {row["Group"]: row for row in s.to_dict("records")}
        comp_rows.append({"Specification": SPEC_LABELS[key],
                          "Exogenous M1": g["Exogenous"]["Mean M1 (depth)"], "Endogenous M1": g["Endogenous"]["Mean M1 (depth)"],
                          "Exogenous M5": g["Exogenous"]["Mean M5 (cumulative loss)"],
                          "Endogenous M5": g["Endogenous"]["Mean M5 (cumulative loss)"],
                          "Exogenous recovered": g["Exogenous"]["Recovered (n/N)"],
                          "Endogenous recovered": g["Endogenous"]["Recovered (n/N)"],
                          "Exogenous M6": g["Exogenous"]["Mean M6 (renewal)"],
                          "Endogenous M6": g["Endogenous"]["Mean M6 (renewal)"]})
    comp_df = pd.DataFrame(comp_rows)
    all_pl = {"primary": (res, pl), "fixed8q": (res_fix, pl_fix), "half": (res_half, pl_half)}
    sig_rows = []
    for key, (rr, pp) in all_pl.items():
        for x in rr:
            sig_rows.append({"Specification": SPEC_LABELS[key], "Code": x["code"], "M1": x["M1"], "p (M1)": pp[x["code"]]["p_M1"],
                             "M5": x["M5"], "p (M5)": pp[x["code"]]["p_M5"], "Status": x["status"].split(" (")[0],
                             "Distinct pseudo-onsets": pp[x["code"]]["n_onsets"]})
    rob["comparison"] = comp_df

    wb = Workbook()
    ws = wb.active
    ws.title = "README"
    lines = [
        "Robustness checks - same data, same metrics (M1-M6), same baseline and exclusion rules as the primary analysis.",
        "Check 1 - Fixed 8-quarter window: every shock is measured over t0 ... t0+7 (capped at 2023Q3), instead of "
        "ending before the next selected shock. Placebo windows also have length 8.",
        "Check 2 - Half-year aggregation: rounds are summed per half-year (H1 = Q1-Q2, H2 = Q3-Q4); S(t) as in the "
        "primary analysis; baseline up to 4 clean half-years (= 8 quarters); post window up to 4 half-years; "
        "window ends before the next selected shock; last usable half-year 2023H1.",
        "Event dates are mapped to the half-year that contains them. All other rules are unchanged.",
        f"Placebo: B = {B} draws with replacement, seed {SEED}. Source of all tables: Author(s), 2026.",
    ]
    for i, t in enumerate(lines, start=1):
        ws.cell(i, 1, t)
    ws.column_dimensions["A"].width = 150

    ws = wb.create_sheet("Comparison")
    r = write_df(ws, comp_df, title="Table R1. Group means across specifications (all six shocks)")
    write_df(ws, pd.DataFrame(sig_rows), r0=r, title="Table R2. Shock-level M1, M5 and placebo p-values across specifications")

    for key, unit in [("fixed8q", "quarters"), ("half", "half-years")]:
        v = rob[key]
        name = "Fixed8Q" if key == "fixed8q" else "HalfYear"
        ws = wb.create_sheet(f"{name}_Metrics")
        write_df(ws, metrics_table(v["results"], v["placebo"], "Deal count", unit),
                 title=f"Metrics - {SPEC_LABELS[key]}")
        ws = wb.create_sheet(f"{name}_ExoEndo")
        r = write_df(ws, v["summary_full"], title=f"Group means - {SPEC_LABELS[key]} - all six shocks")
        write_df(ws, v["summary_core"], r0=r, title=f"Group means - {SPEC_LABELS[key]} - clean core (EX1, EX2 vs EN1, EN2)")
        ws = wb.create_sheet(f"{name}_Windows")
        write_df(ws, pd.DataFrame([{"Code": x["code"], "t0": x["t0_label"], "Window": f"{x['t0_label']}-{x['tend']}",
                                    "Baseline used": ", ".join(x["base"]), "Baseline excluded": ", ".join(x["base_excluded"]),
                                    "Post (M6)": ", ".join(x["post"])} for x in v["results"]]),
                 title=f"Windows - {SPEC_LABELS[key]}")
    ws = wb.create_sheet("HalfYear_Series")
    write_df(ws, hs, title="Half-year series", fmt={"P(t) deals": "0.0000"})

    ws_d = wb.create_sheet("Chart_Data")
    ws_f = wb.create_sheet("Figures")
    ws_f.cell(1, 1, "Figure R1. Exogenous vs endogenous group means across specifications (editable). Source: Author(s), 2026.")
    add_spec_charts(ws_d, ws_f, comp_df)
    wb.save(OUT / "Robustness_Checks.xlsx")
    plot_spec_comparison(comp_df, FIG / "FigureR1_robustness_comparison.png")
    return rob


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
def main():
    OUT.mkdir(parents=True, exist_ok=True)
    FIG.mkdir(exist_ok=True)
    deals, comp, exits = load_data()
    pop, pop_info = build_population(comp, exits)
    ser, annual = build_series(deals, pop)
    rng = np.random.default_rng(SEED)

    p_deals = ser["P(t) deals"].to_numpy()
    p_amt = ser["P(t) amount"].to_numpy()
    res = analyse_events(p_deals)
    res_amt = analyse_events(p_amt)
    pl = placebo(p_deals, res, rng)
    pl_amt = placebo(p_amt, res_amt, rng)

    summ_full = group_summary(res)
    summ_core = group_summary(res, CLEAN_CORE)
    summ_amt = group_summary(res_amt)

    # --- descriptive tables
    d = deals.copy()
    d["Year"] = d["Deal Date"].dt.year
    yearly = d.groupby("Year").agg(Deals=("Deal ID", "size"),
                                   **{"Deals with reported amount": ("Deal Amount ($m)", "count"),
                                      "Reported amount ($m)": ("Deal Amount ($m)", "sum"),
                                      "Median reported deal ($m)": ("Deal Amount ($m)", "median")}).reset_index()
    yearly = yearly.merge(annual, on="Year", how="left")
    stage = d.groupby("Deal Stage").agg(Deals=("Deal ID", "size"),
                                        **{"Reported amount ($m)": ("Deal Amount ($m)", "sum"),
                                           "Median reported deal ($m)": ("Deal Amount ($m)", "median")}).reset_index()
    stage["Share of deals"] = stage["Deals"] / stage["Deals"].sum()
    dtype = d["Deal Type"].value_counts().rename_axis("Deal type").reset_index(name="Deals")
    status = pop["Company Status"].value_counts().rename_axis("Company status").reset_index(name="Companies")
    sector = pop["Sector"].value_counts().rename_axis("Sector").reset_index(name="Companies")
    ex_ht = exits[exits["Company ID"].isin(pop["Company ID"])]
    exit_tab = ex_ht.groupby([ex_ht["Exit Date"].dt.year, "Exit Type"]).size().unstack(fill_value=0)
    exit_tab = exit_tab.reindex(columns=["Acquired", "IPO"], fill_value=0).rename_axis("Year").reset_index()
    overview = pd.DataFrame({
        "Item": ["Financing rounds (dated)", "Rounds with reported amount", "Total reported amount ($m)",
                 "Mean reported round ($m)", "Median reported round ($m)", "Companies with at least one round",
                 "Companies in population file (unique)", "  of which MNC (excluded from S(t))",
                 "HT companies in S(t)", "  ceased (with ceased year)", "  acquired (removed in acquisition year)",
                 "Exit events of HT companies (acquisitions / IPOs)", "Founding years", "Deal dates"],
        "Value": [len(d), int(d["Deal Amount ($m)"].notna().sum()), round(float(d["Deal Amount ($m)"].sum()), 1),
                  round(float(d["Deal Amount ($m)"].mean()), 2), round(float(d["Deal Amount ($m)"].median()), 2),
                  int(d["Company ID"].nunique()), pop_info["all_companies"], pop_info["mnc_excluded"],
                  pop_info["ht_companies"], pop_info["ceased_with_year"],
                  f"{pop_info['acquired_with_date']} of {pop_info['acquired_status']} with Acquired/Merged status",
                  f"{int((ex_ht['Exit Type'] == 'Acquired').sum())} / {int((ex_ht['Exit Type'] == 'IPO').sum())}",
                  f"{int(pop['Est. Year'].min())}-{int(pop['Est. Year'].max())}",
                  f"{d['Deal Date'].min():%Y-%m-%d} to {d['Deal Date'].max():%Y-%m-%d}"],
    })

    # --- workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "README"
    readme = [
        "Resilience of Israeli energy-tech startups: exogenous vs endogenous shocks - results workbook",
        "Generated by resilience_analysis.py from the IVC input files (read-only). All tables: Source: Author(s), 2026.",
        "Data: IVC Research Center exports - Capital Raising (rounds), Companies & MNCs (population), Exits (acquisition dates).",
        "Primary performance variable: P(t) = number of financing rounds in quarter t / active firm stock S(t).",
        "Robustness: P(t) built from reported round amounts ($m); rounds without a reported amount count as 0.",
        "S(t): HT companies only. Year-end stock S(y) = S(y-1) + founded(y) - ceased(y) - acquired(y), held constant "
        "within each year; acquired = Acquired/Merged status, year of first acquisition in the Exits file.",
        f"Baseline: mean P(t) over up to {MAX_BASE} pre-shock quarters, excluding quarters inside any other listed event window [t0, t1].",
        "Measurement window: t0 to the quarter before the next selected shock; the last window ends 2023Q3 (before the Iron Swords War).",
        f"Recovery: first quarter after the trough with P~(t) >= {REC}. Not reached inside the window = Censored.",
        f"M6: mean P(t) over up to {MAX_POST} quarters after t1 (inside the window, excluding other event quarters) / baseline mean.",
        f"Placebo: pseudo-shocks at quiet quarters (outside all listed event windows, up to 2023Q3, >= {MIN_BASE} clean baseline quarters),",
        f"  same window length as the real shock; B = {B} draws with replacement (seed {SEED}); p = share of draws with M >= actual M.",
        "Metric sources: " + "; ".join(f"{k}: {v}" for k, v in METRIC_SOURCES.items()),
    ]
    for i, line in enumerate(readme, start=1):
        ws.cell(i, 1, line)
    ws.column_dimensions["A"].width = 150

    ws = wb.create_sheet("Descriptive")
    r = write_df(ws, overview, title="Table 1. Dataset overview")
    r = write_df(ws, yearly, r0=r, title="Table 2. Financing rounds and firm population by year")
    r = write_df(ws, stage, r0=r, title="Table 3. Financing rounds by development stage", fmt={"Share of deals": "0.0%"})
    r = write_df(ws, dtype, r0=r, title="Table 4. Financing rounds by round type")
    r = write_df(ws, status, r0=r, title="Table 5. HT companies by current status")
    r = write_df(ws, sector, r0=r, title="Table 6. HT companies by sector (descriptive only; not used in the results)")
    write_df(ws, exit_tab, r0=r, title="Table 6a. Exits of HT companies by year (IVC Exits file)")

    ws = wb.create_sheet("Quarterly_Series")
    write_df(ws, ser, title="Quarterly series, 2013Q1-2026Q1", fmt={"P(t) deals": "0.0000", "P(t) amount": "0.0000"})

    ws = wb.create_sheet("Event_Windows")
    ev_rows = []
    for e in EVENTS:
        rr = next((x for x in res if x["code"] == e[0]), None)
        ev_rows.append({
            "Code": e[0], "Event": e[1], "Class": e[2], "Direction": e[3], "t0": e[4], "t1": e[5],
            "Selected": "Yes" if e[6] else "No (exclusion only)",
            "Baseline quarters used": ", ".join(rr["base"]) if rr else "",
            "Baseline quarters excluded": ", ".join(rr["base_excluded"]) if rr else "",
            "Measurement window": f"{rr['t0']}-{rr['tend']}" if rr else "",
            "Post quarters (M6)": ", ".join(rr["post"]) if rr else "",
            "Concurrent listed events": ", ".join(rr["concurrent"]) if rr else "",
        })
    write_df(ws, pd.DataFrame(ev_rows), title="Event list, windows and exclusions")

    ws = wb.create_sheet("Metrics")
    mt = metrics_table(res, pl, "Deal count")
    r = write_df(ws, mt, title="Table 7. Resilience-curve metrics of the six selected shocks (primary: deal count)")
    write_df(ws, pd.DataFrame([{"Metric": k, "Source": v} for k, v in METRIC_SOURCES.items()]), r0=r,
             title="Metric definitions: sources")

    ws = wb.create_sheet("Exo_vs_Endo")
    r = write_df(ws, summ_full, title="Table 8. Exogenous vs endogenous shocks - all six events (deal count)")
    r = write_df(ws, summ_core, r0=r, title="Table 9. Exogenous vs endogenous shocks - clean core (EX1, EX2 vs EN1, EN2)")
    write_df(ws, summ_amt, r0=r, title="Table 10. Robustness - all six events, reported-amount P(t)")

    ws = wb.create_sheet("Robustness_Amount")
    write_df(ws, metrics_table(res_amt, pl_amt, "Reported amount"),
             title="Table 11. Resilience-curve metrics using reported amounts ($m) instead of deal counts")

    ws = wb.create_sheet("Placebo")
    prow = []
    for x in res:
        q = pl[x["code"]]
        prow.append({"Code": x["code"], "Window length (quarters)": x["window_len"],
                     "Distinct pseudo-onsets": q["n_onsets"], "Pseudo-onset quarters": ", ".join(q["onsets"]),
                     "Actual M1": x["M1"], "Placebo mean M1": float(np.mean(q["M1_draws"])) if q["M1_draws"] else None,
                     "p (M1)": q["p_M1"], "Actual M5": x["M5"],
                     "Placebo mean M5": float(np.mean(q["M5_draws"])) if q["M5_draws"] else None, "p (M5)": q["p_M5"]})
    write_df(ws, pd.DataFrame(prow), title="Table 12. In-time placebo test (deal count)")

    # --- native charts
    ws = wb.create_sheet("Chart_Data")
    figs = wb.create_sheet("Figures")
    notes = [
        "Editable Excel charts (same content as the PNG previews in the figures folder).",
        "To use in Word: click a chart > Ctrl+C > in Word: Paste > Keep Source Formatting & Embed Workbook. "
        "Add the caption and 'Source: Author(s), 2026.' in Word.",
        "Figure 1 (rows 4-56): panels a-f, exogenous shocks left, endogenous shocks right.",
        "Figure 2 (row 58): group means.   Figure 3 (row 77): quarterly series.",
    ]
    for i, text in enumerate(notes, start=1):
        figs.cell(i, 1, text)
    by_code = {x["code"]: x for x in res}
    col = 1
    for k, (code, letter) in enumerate(PANELS):
        anchor = f"{'A' if k % 2 == 0 else 'J'}{5 + (k // 2) * 18}"
        col = add_panel_chart(ws, figs, col, by_code[code], letter, anchor)
        plot_panel(by_code[code], letter, FIG / f"Figure1{letter}_{code}.png")
    col = add_group_charts(ws, figs, col, summ_full)
    add_series_chart(ws, figs, col, ser)

    xlsx = OUT / "Resilience_Results_Summary.xlsx"
    wb.save(xlsx)

    for old in ["Figure1_resilience_curves.png"]:
        (FIG / old).unlink(missing_ok=True)
    plot_groups(summ_full, FIG / "Figure2_exo_vs_endo.png")
    plot_series(ser, FIG / "Figure3_quarterly_series.png")

    rob = run_robustness(ser, annual, deals, res, pl, res_amt, summ_full, summ_core, summ_amt)

    # --- exact permutation test of the group difference (all six shocks, each specification)
    perm_specs = [("primary", res), ("amount", res_amt), ("fixed8q", rob["fixed8q"]["results"]),
                  ("half", rob["half"]["results"])]
    perm = pd.DataFrame([{"Specification": SPEC_LABELS[key], **group_permutation(rr, m)}
                         for key, rr in perm_specs for m in PERM_METRICS])
    wb_p = Workbook()
    ws = wb_p.active
    ws.title = "README"
    for i, line in enumerate([
        "Exact permutation test of the difference between endogenous and exogenous shocks (group means, all six shocks).",
        "All C(6,3) = 20 splits of the six shocks into two groups of three are enumerated; no random draws are used.",
        "Null hypothesis: the exogenous/endogenous label is unrelated to the metric (labels are exchangeable).",
        "p two-sided = share of splits with |difference| >= |observed|; with 20 balanced splits the minimum attainable "
        "two-sided p is 2/20 = 0.10. p one-sided is reported for completeness only: the hypothesis (H1) predicted the "
        "opposite direction, so the two-sided p is the one interpreted.",
        "M1 is the primary metric and M5 the secondary one. M6 is not tested (defined for only two shocks per group).",
        "Source: " + METRIC_SOURCES["Group permutation test"] + ". Source of the table: Author(s), 2026.",
    ], start=1):
        ws.cell(i, 1, line)
    ws.column_dimensions["A"].width = 150
    ws = wb_p.create_sheet("Results")
    write_df(ws, perm, title="Table P1. Exact permutation test: endogenous minus exogenous group means")
    wb_p.save(OUT / "Group_Permutation_Test.xlsx")

    def strip(plc):
        return {k: {kk: vv for kk, vv in v.items() if not kk.endswith("_draws")} for k, v in plc.items()}

    dump = dict(results=res, results_amount=res_amt, placebo=strip(pl), placebo_amount=strip(pl_amt),
                summary_full=summ_full.to_dict("records"), summary_core=summ_core.to_dict("records"),
                summary_amount=summ_amt.to_dict("records"), overview=overview.to_dict("records"),
                yearly=yearly.to_dict("records"), stage=stage.to_dict("records"), population=pop_info,
                exits=exit_tab.to_dict("records"),
                robustness={k: dict(results=v["results"], placebo=strip(v["placebo"]),
                                    summary_full=v["summary_full"].to_dict("records"),
                                    summary_core=v["summary_core"].to_dict("records"))
                            for k, v in rob.items() if k != "comparison"},
                robustness_comparison=rob["comparison"].to_dict("records"),
                group_permutation=perm.to_dict("records"))
    with open(OUT / "results.json", "w", encoding="utf-8") as fh:
        json.dump(dump, fh, ensure_ascii=False, indent=1, default=lambda o: o.item() if isinstance(o, np.generic) else str(o))
    print("Wrote", xlsx)


def parse_args():
    ap = argparse.ArgumentParser(description="Resilience analysis of Israeli energy-tech startups "
                                             "(exogenous vs endogenous shocks).")
    ap.add_argument("--data", required=True, type=Path,
                    help="folder that contains the three IVC .xlsx files (see README.md)")
    ap.add_argument("--out", default=Path("output"), type=Path,
                    help="folder for the results (created if missing; default: ./output)")
    args = ap.parse_args()
    missing = [f for f in INPUT_FILES if not (args.data / f).is_file()]
    if missing:
        ap.error("missing input file(s) in --data: " + ", ".join(missing))
    return args


if __name__ == "__main__":
    _args = parse_args()
    DATA, OUT = _args.data, _args.out
    FIG = OUT / "figures"
    main()
