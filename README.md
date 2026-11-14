# Investment-Based Resilience of the Israeli Energy Startup Sector

Analysis code for the paper *"Investment-Based Resilience of the Israeli Energy Startup Sector: A Comparative Analysis of Exogenous and Endogenous Shocks"*.

The script `resilience_analysis.py` produces every table and figure in the results of the paper from three data exports of the IVC Research Center.

> **Data are not included.** The IVC data are commercial and licensed, so they cannot be redistributed. To reproduce the results you need your own access to IVC and the three exports described below. This repository contains no data, no results and no derived outputs.

## What the code computes

1. **Performance series.** The series is P(t) = Y(t) / S(t), computed for each quarter from 2013Q1 to 2026Q1.
   - Y(t) is the number of financing rounds in quarter t.
   - S(t) is the stock of active high-tech (HT) companies. The year-end stock is S(y) = S(y−1) + founded(y) − ceased(y) − acquired(y), and it is held constant within each year.
   - Multinational companies (MNC) are excluded from the stock.
   - An acquired company leaves the stock in the year of its first acquisition in the Exits file.
2. **Shocks.** The script analyses six pre-selected adverse shocks:
   - three exogenous: EX1, EX2 and EX3;
   - three endogenous: EN1, EN2 and EN3.

   The other listed events are used only to exclude quarters from baselines, post-shock periods and placebo onsets. The event list (`EVENTS`) is defined in the code.
3. **Resilience metrics.** For each shock, the baseline is the mean P(t) over up to 8 clean pre-shock quarters. The relative performance is P̃(t) = P(t) / baseline. The metrics are:

   | Metric | Meaning |
   |---|---|
   | M1 | Depth of the shock (1 − minimum P̃) |
   | M2 | Time from onset to trough |
   | M3 | Recovery time (first quarter after the trough with P̃ ≥ 0.95; censored if not reached) |
   | M4 | Recovery velocity |
   | M5 | Cumulative loss (sum of 1 − P̃ over the window) |
   | M6 | Renewal: mean post-shock level relative to the baseline |

4. **In-time placebo test.** Pseudo-shocks are placed at quiet quarters, using the same window length as the real shock. There are B = 500 draws with replacement, seed 20261030. The p-value is the share of draws with M ≥ actual M.
5. **Exact permutation test.** The test compares the endogenous and exogenous groups' means of M1 and M5. It enumerates all C(6,3) = 20 splits of the six shocks, so no random draws are used.
6. **Robustness checks:**
   - P(t) built from reported amounts instead of deal counts;
   - a clean core of four shocks (EX1 and EX2 vs EN1 and EN2);
   - fixed 8-quarter windows;
   - half-year aggregation.

## Input files (not included)

Put the three IVC exports, with exactly these file names, in one folder:

| File | Columns used |
|---|---|
| `Energy High-Tech Capital Raising.xlsx` | `Deal Date`, `Deal ID`, `Deal Amount ($m)`, `Deal Stage`, `Deal Type`, `Company ID` |
| `Energy High-Tech Companies & MNCs.xlsx` | `Company ID`, `Est. Year`, `Company Type` (`HT` / `MNC`), `Company Status` (incl. `Acquired/Merged`), `Ceased Year`, `Sector` |
| `Energy High-Tech Exits.xlsx` | `Company ID`, `Exit Date`, `Exit Type` (`Acquired` / `IPO`) |

- Rows without a `Deal Date`, `Est. Year` or `Exit Date` are dropped. This removes the filter-description row at the end of each IVC export.
- `Sector` is used for a descriptive table only.

## How to run

The code requires Python 3.10 or later.

```bash
pip install -r requirements.txt
python resilience_analysis.py --data path/to/ivc_data --out output
```

- `--data` is the folder with the three input files.
- `--out` is the output folder. It is created if it is missing, and the default is `./output`.

The run is deterministic: the placebo test uses a fixed seed, and the permutation test is exact.

## Outputs and where they appear in the paper

All outputs are written to the `--out` folder. The figures go into its `figures/` subfolder.

| Output | Content | Paper |
|---|---|---|
| `Resilience_Results_Summary.xlsx`, sheet *Descriptive* | Dataset overview; rounds and firm population by year | Table 3.1, Table 4.1 |
| `Resilience_Results_Summary.xlsx`, sheet *Event_Windows* | Event list, baselines, windows and exclusions | Table 3.2, Table A1 |
| `Resilience_Results_Summary.xlsx`, sheets *Metrics* and *Placebo* | M1–M6 by shock, with placebo p-values | Table 4.2 |
| `Resilience_Results_Summary.xlsx`, sheet *Exo_vs_Endo* | Group means (all six shocks; clean core; amounts) | Table 4.3 |
| `Group_Permutation_Test.xlsx` | Exact permutation test (M1, M5; all specifications) | Table 4.4 |
| `Robustness_Checks.xlsx`, sheet *Comparison* | Group means across specifications | Table 4.5 |
| `figures/Figure3_quarterly_series.png` | Quarterly P(t), 2013Q1–2026Q1 | Figure 1 |
| `figures/Figure1a_*.png` … `Figure1f_*.png` | Resilience curves of the six shocks | Figure 2 |
| `figures/Figure2_exo_vs_endo.png` | Mean metrics: exogenous vs endogenous | Figure 3 |
| `figures/FigureR1_robustness_comparison.png` | Robustness of the comparison | Figure 4 |
| `results.json` | All computed values in machine-readable form | — |

The file names of the figures follow the order in which they were produced, not the figure numbers in the paper. The table above gives the mapping. The Excel workbooks also contain editable versions of all charts.

## Notice

The output files are derived from licensed IVC data. Do not commit them to a public repository. The included `.gitignore` excludes data files, Excel files and outputs.
