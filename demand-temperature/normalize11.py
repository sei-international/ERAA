"""
The purpose of this script is to execute Steps 2-4 of the USAID (2024) method per region.
This code generates 7 regions x 3 future years x 16 climate years (2010-2025) of hourly load.

It performs the following steps:

1. Calculate the normal temperature (per region, per calendar day)
       average of that day's mean temperature over NORMAL_YEARS (2010-2025).

2. Normalize the base year (2025)
       for each day d in 2025:  delta_d = f(T_base,d) - f(T_normal,d)            (MW)
       normalized load(h) = actual load(h) - delta_d   for all 24 hours of day d

3. Scale to each future year
       k = target / normalized peak   (TARGET_TYPE = "peak", as in the report, MW)
       scaled load = normalized load x k
   A region without a target keeps k = 1 (no growth) and a warning is printed.

4. Climate years (put the temperature effect back)
       for climate year y and day d:  delta_d = f(T_y,d) - f(T_normal,d)  (MW)
       load(h) = scaled load(h) + delta_d
   The base-year calendar is kept (weekdays, holidays); only the weather changes,
   matched by calendar date. 29 February is dropped from leap climate years -> 8,760 h.

Outputs (normalized_output/)
    climate_year_load_<future year>.xlsx      one workbook per future year
        Summary          per region x climate year: annual GWh, peak MW, peak hour, min MW
        <region> tabs    rows = 8,760 hours (UTC+2, base-year calendar),
                         columns = climate years (MW)
    base_and_inputs.xlsx
        Base year        per region: actual, temperature effect, normalized load (MW)
        Scaling          per region x future year: target and scale factor k
        Normal temperature, Polynomials used, Notes
    temperature_values/daily_temperature_by_region.csv   cache of daily temperatures
"""

#Imports
import sys
from pathlib import Path

import numpy as np
import pandas as pd

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))
# same load cleaning / time alignment (UTC+2) as the fits
from new_fit_demand_temperature import (LOAD_FILE, TEMP_DIR, read_right_bank,
                                        read_left_bank, read_temperature, match_regions)

#Settings/global variables
BASE_YEAR = 2025
CLIMATE_YEARS = range(2010, 2026)          # weather years applied in step 4
NORMAL_YEARS = range(2010, 2026)           # years averaged for the "normal" temperature (used in step 2)
FUTURE_YEARS = [2028, 2030, 2035]

TARGET_TYPE = "peak"                       # "peak" (MW)
# One target per future year for all 7 regions together: MW (peak) or GWh (energy).
# None = no scaling (k = 1).
TARGETS = {
    2028: 962,
    2030: 1023,
    2035: 1044,
}

FIT_FILE = BASE / "fit_output" / "demand_temperature_cubic_2021_2025.xlsx"
OUT_DIR = BASE / "normalized_output"
DAILY_CACHE = TEMP_DIR / "daily_temperature_by_region.csv"
COEF_COLS = ["avg_a3", "avg_a2", "avg_a1", "avg_a0"]
# --------------------------------------------------------------------------


#Helper functions
def load_polynomials():
    p = pd.read_excel(FIT_FILE, sheet_name="Average cubic").set_index("region")
    return p[COEF_COLS + ["T_valid_min_degC", "T_valid_max_degC", "years_used"]]


def make_f(row):
    """f(T) in MW for one region: averaged cubic, held flat outside the valid range."""
    lo, hi = row["T_valid_min_degC"], row["T_valid_max_degC"]
    c = row[COEF_COLS].values.astype(float)
    return lambda T: np.polyval(c, np.clip(np.asarray(T, float), lo, hi))


def daily_temperatures(years):
    """Daily mean temperature (deg C) per region, from the yearly workbooks (cached)."""
    files = {y: TEMP_DIR / f"temperature_by_region_{y}.xlsx" for y in years}
    missing = [y for y, f in files.items() if not f.exists()]
    if missing:
        sys.exit(f"ERROR: temperature workbooks missing for {missing}.")
    newest = max(f.stat().st_mtime for f in files.values())
    if DAILY_CACHE.exists() and DAILY_CACHE.stat().st_mtime > newest:
        cached = pd.read_csv(DAILY_CACHE, index_col=0, parse_dates=True)
        if set(years) <= set(cached.index.year):
            return cached
    parts = []
    for y in years:
        print(f"reading temperature {y} ...")
        parts.append(read_temperature(y).resample("D").mean())
    daily = pd.concat(parts)
    daily.to_csv(DAILY_CACHE)
    return daily


def base_year_load():
    rb = pd.read_excel(LOAD_FILE, sheet_name="Right bank", header=None)
    lb = pd.read_excel(LOAD_FILE, sheet_name="Left bank", header=None)
    fixes = []
    load = read_right_bank(BASE_YEAR, rb, fixes)
    load["LEFT"] = read_left_bank(BASE_YEAR, lb, fixes)
    full = pd.date_range(f"{BASE_YEAR}-01-01", f"{BASE_YEAR}-12-31 23:00", freq="h")
    return load.reindex(full).interpolate(limit=3)


def by_calendar_day(daily):
    """Index daily values by 'MM-DD' and drop 29 February."""
    d = daily[~((daily.index.month == 2) & (daily.index.day == 29))].copy()
    d.index = d.index.strftime("%m-%d")
    return d


#MAIN
def main():
    OUT_DIR.mkdir(exist_ok=True)
    polys = load_polynomials()
    regions = list(polys.index)

    years_needed = sorted(set(CLIMATE_YEARS) | set(NORMAL_YEARS) | {BASE_YEAR})
    tdaily = daily_temperatures(years_needed)

    load = base_year_load()
    m = match_regions(list(load.columns), list(tdaily.columns))
    load = load.rename(columns=m)
    missing = [r for r in regions if r not in load.columns or r not in tdaily.columns]
    if missing:
        sys.exit(f"ERROR: no load or temperature for regions {missing}")
    load = load[regions]
    if load.isna().any().any():
        print("WARNING: gaps left in base-year load after interpolation:", load.isna().sum().to_dict())

    # 1. normal temperature per calendar day (365 days)
    normal_src = by_calendar_day(tdaily[tdaily.index.year.isin(NORMAL_YEARS)][regions])
    normal = normal_src.groupby(level=0).mean().sort_index()

    day_key = load.index.strftime("%m-%d")                 # calendar day of every base-year hour
    t_base = by_calendar_day(tdaily[tdaily.index.year == BASE_YEAR][regions])

    outputs = {fy: {} for fy in FUTURE_YEARS}
    base_tabs, scaling, summary = {}, [], []
    normalized_all, deltas_all = {}, {}
    for r in regions:
        f = make_f(polys.loc[r])
        f_normal = pd.Series(f(normal[r].values), index=normal.index)          # MW per calendar day

        # 2. normalize the base year: remove the base year's temperature effect (MW)
        delta_base = (pd.Series(f(t_base[r].values), index=t_base.index) - f_normal) \
            .reindex(day_key).values
        actual = load[r]
        normalized = actual - delta_base
        base_tabs[r] = pd.DataFrame({"actual_MW": actual, "temperature_effect_MW": delta_base,
                                     "normalized_MW": normalized})

        # temperature effect of every climate year (MW per hour, same value for all 24 hours of a day)
        deltas = {}
        for y in CLIMATE_YEARS:
            t_y = by_calendar_day(tdaily[tdaily.index.year == y][r])
            d = (pd.Series(f(t_y.values), index=t_y.index) - f_normal).reindex(day_key).values
            if np.isnan(d).any():
                print(f"WARNING: {r} {y}: {int(np.isnan(d).sum() / 24)} day(s) without temperature "
                      f"-> normal temperature used")
                d = np.where(np.isnan(d), 0.0, d)
            deltas[y] = d
        normalized_all[r], deltas_all[r] = normalized, deltas

    # 3. one scale factor per future year, from the sum of all regions (coincident peak / energy)
    national = pd.DataFrame(normalized_all).sum(axis=1)
    ref = national.max()
    ks = {}
    for fy in FUTURE_YEARS:
        target = TARGETS.get(fy)
        if target is None:
            print(f"WARNING: no {TARGET_TYPE} target for {fy} -> no scaling (k = 1)")
            ks[fy] = 1.0
        else:
            ks[fy] = target / ref
        scaling.append(dict(future_year=fy, target_type=TARGET_TYPE, target=target,
                            normalized_reference_all_regions=ref, scale_factor_k=ks[fy]))
        print(f"{fy}: target {target} | normalized {TARGET_TYPE} of all regions {ref:.1f} | k = {ks[fy]:.3f}")

    for r in regions:
        normalized, deltas = normalized_all[r], deltas_all[r]
        for fy in FUTURE_YEARS:
            k = ks[fy]
            scaled = normalized * k

            # 4. add each climate year's temperature effect (MW, not scaled)
            cols = {}
            for y in CLIMATE_YEARS:
                cols[y] = scaled.values + deltas[y]
                s = pd.Series(cols[y], index=load.index)
                summary.append(dict(future_year=fy, region=r, climate_year=y,
                                    peak_hour=s.idxmax().strftime("%m-%d %H:%M"), min_MW=s.min()))
            outputs[fy][r] = pd.DataFrame(cols, index=pd.Index(load.index.strftime("%Y-%m-%d %H:%M"),
                                                               name=f"hour_{BASE_YEAR}_calendar_UTC+2"))
            peaks = [c.max() for c in cols.values()]
            print(f"{r:>6} {fy}: k = {k:.3f} | climate-year peaks {min(peaks):7.1f} - {max(peaks):7.1f} MW")

    summ = pd.DataFrame(summary)

    # ---- Excel: one workbook per future year
    for fy in FUTURE_YEARS:
        path = OUT_DIR / f"climate_year_load_{fy}.xlsx"
        with pd.ExcelWriter(path, engine="openpyxl") as xw:
            summ[summ.future_year == fy].round(2).to_excel(xw, sheet_name="Summary", index=False)
            for r in regions:
                outputs[fy][r].round(3).to_excel(xw, sheet_name=str(r)[:31])
            for ws in xw.sheets.values():
                ws.freeze_panes = "B2"
                ws.column_dimensions["A"].width = 20
        print(f"written {path}")

    # ---- Excel: base year and inputs
    notes = pd.DataFrame({"note": [
        f"Base year {BASE_YEAR}; climate years {min(CLIMATE_YEARS)}-{max(CLIMATE_YEARS)}; normal temperature = "
        f"mean of {min(NORMAL_YEARS)}-{max(NORMAL_YEARS)} per calendar day.",
        "f(T) = averaged cubic per region in MW (fit_demand_temperature.py), held flat outside its valid "
        "temperature range.",
        "Normalize: load - [f(T_base) - f(T_normal)] per day (MW, same for all 24 hours).",
        f"Scale: one target per future year; k = target / {TARGET_TYPE} of the sum of all regions' normalized "
        "load; the same k is applied to every region. Years without a target: k = 1.",
        "Climate year y: scaled load + [f(T_y) - f(T_normal)] per day (MW, not scaled by k).",
        "Rows follow the base-year calendar (UTC+2, 8,760 h); weather matched by calendar date; 29 Feb dropped.",
        "Load in A, B, D, E, F is NET of distributed PV.",
    ]})
    path = OUT_DIR / "base_and_inputs.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        base = pd.concat(base_tabs, axis=1)
        base.columns = [f"{r} {c}" for r, c in base.columns]
        base.index = base.index.strftime("%Y-%m-%d %H:%M")
        base.index.name = "hour_UTC+2"
        base.round(4).to_excel(xw, sheet_name="Base year")
        pd.DataFrame(scaling).round(4).to_excel(xw, sheet_name="Scaling", index=False)
        normal.round(3).rename_axis("calendar_day").to_excel(xw, sheet_name="Normal temperature")
        polys.to_excel(xw, sheet_name="Polynomials used")
        notes.to_excel(xw, sheet_name="Notes", index=False)
        for ws in xw.sheets.values():
            ws.freeze_panes = "B2"
            ws.column_dimensions["A"].width = 20
    print(f"written {path}")


if __name__ == "__main__":
    main()