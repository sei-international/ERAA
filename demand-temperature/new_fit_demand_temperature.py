"""
The purpose of this script is to fit the relationship between daily electricity demand & temperature for 
every region of Moldova, 2021-2025, using USAID 2024 method. 
This script returns one cubic polynomial per region per year then averaged across the years.

1. Daily values. For each region and each calendar day (fixed UTC+2):
       daily load        = mean of the day's 24 hourly loads (MW)
       daily temperature = mean of the day's 24 hourly regional temperatures (deg C)

2. Yearly cubic. Least-squares fit for every region and year:
       load(T) = a3*T^3 + a2*T^2 + a1*T + a0                       (MW)

3. Average cubic per region = mean of the yearly MW coefficients (a3, a2, a1, a0).
   Valid range = temperatures covered by every year used (highest yearly minimum to
   lowest yearly maximum). Outside it, hold the curve flat at the edge value:
       f(T) = polyval(avg coefficients, clip(T, T_valid_min, T_valid_max))

Time alignment: ALL data is put into UTC+2 (no daylight saving):
    * Right bank: already fixed UTC+2. 
    * Left bank: local clock with daylight saving (Europe/Chisinau). Summer hours are
      moved one hour earlier; the non-existent spring hour (a filler copy in the file)
      is dropped; the autumn hour that occurs twice in local time is restored - split in
      half if the two hours were summed into one row, else interpolated.

Outputs (fit_output/)
    demand_temperature_cubic_2021_2025.xlsx
        Average cubic   per region: averaged MW coefficients, valid range, years used
        Cubic by year   per region and year: coefficients (MW), R2,
                        temperature range, status (used / excluded and why)
        Daily data      the daily points used in the fits
        Data fixes      corrections made to the load data (spikes, DST)
   
"""

import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

#Settings
BASE = Path(__file__).resolve().parent
LOAD_FILE = BASE / "Historical hourly electricity load.xlsx"
TEMP_DIR = BASE / "temperature_values"
OUT_DIR = BASE / "fit_output"
YEARS = range(2021, 2026)
FIXED_TZ = "Etc/GMT-2"            # fixed UTC+2
LEFT_BANK_TZ = "Europe/Chisinau"  # local time with daylight savings
DEGREE = 3

LEFT_BANK_TAB = None   # name of the left-bank tab in the temperature workbooks;
                       # None = the one tab that isn't A-F

# Region-years left out of the fits
EXCLUDE = {
    (2024, "A"): "2024 is a scaled copy of 2023 (r = 0.999)",
    (2024, "B"): "2024 is an exact copy of 2023",
    (2024, "D"): "2024 is an exact copy of 2023",
    (2024, "E"): "2024 is an exact copy of 2023",
    (2024, "F"): "2024 is an exact copy of 2023",
}
# Days left out of the fits (events / filled data)
EXCLUDE_DAYS = {
    "2022-11-15": "grid emergency (strikes on Ukraine's grid)",
    "2022-11-23": "blackout",
    "2022-10-30": "left bank: whole day copied from 2023-10-29",
}
# Single-hour errors, given in the original timestamps of the load data
SPIKES = [
    ("C", "2025-02-04 18:30"), ("C", "2025-11-20 15:30"), ("C", "2025-12-12 03:30"),
    ("LEFT", "2023-09-09 02:30"),
]

#Helper functions
def to_timestamps(col):
    """Excel serial dates or date strings -> timestamps."""
    num = pd.to_numeric(col, errors="coerce")
    if num.notna().mean() > 0.9:
        return pd.to_datetime(num, unit="D", origin="1899-12-30").dt.round("min")
    return pd.to_datetime(col, errors="coerce")


def year_of(cell):
    try:
        return int(float(cell))
    except (TypeError, ValueError):
        return None


def read_right_bank(year, raw, fixes):
    ## The right bank is already in fixed UTC+2, so the timestamps use .floor to round an hour down (ie 7:30 --> 7:00)
    col0 = next((c for c in raw.columns if year_of(raw.iat[0, c]) == year), None)
    if col0 is None:
        return None
    block = raw.iloc[2:, col0:col0 + 7]
    t = to_timestamps(block.iloc[:, 0])
    ok = t.notna()
    names = [str(n).replace("Region", "").strip() for n in raw.iloc[1, col0 + 1:col0 + 7]]
    df = pd.DataFrame(block.loc[ok].iloc[:, 1:].astype(float).values,
                      index=t[ok].values, columns=names)
    for reg, ts in SPIKES:
        ts = pd.Timestamp(ts)
        if reg in df.columns and ts in df.index:
            fixes.append(dict(series=reg, time=ts, value_MW=round(df.at[ts, reg], 1),
                              action="single-hour error -> interpolated"))
            df.at[ts, reg] = np.nan
    df.index = df.index.floor("h")             
    df = df[~df.index.duplicated()]
    return df.interpolate(limit=3)


def read_left_bank(year, raw, fixes):
    col = next((c for c in raw.columns if year_of(raw.iat[1, c]) == year), None)
    if col is None:
        return None
    t = to_timestamps(raw.iloc[2:, col])
    v = pd.to_numeric(raw.iloc[2:, col + 1], errors="coerce")
    ok = t.notna()
    s = pd.Series(v[ok].values, index=t[ok].values, dtype=float)
    for reg, ts in SPIKES:
        ts = pd.Timestamp(ts)
        if reg == "LEFT" and ts in s.index:
            fixes.append(dict(series="Left bank", time=ts, value_MW=round(s[ts], 1),
                              action="single-hour error -> interpolated"))
            s[ts] = np.nan

    # local clock with DST -> fixed UTC+2
    label = s.index.floor("h")
    local = label.tz_localize(LEFT_BANK_TZ,
                              ambiguous=np.ones(len(label), bool),   # repeated hour = summer time
                              nonexistent="NaT")
    for ts in label[local.isna()]:
        fixes.append(dict(series="Left bank", time=ts, value_MW=round(s[label == ts].iloc[0], 1),
                          action="spring DST: non-existent local hour (filler copy) dropped"))
    keep = ~local.isna()
    s = pd.Series(s.values[keep], index=local[keep].tz_convert(FIXED_TZ).tz_localize(None))
    s = s[~s.index.duplicated()]

    # autumn: the hour lost by the shift -> split a summed (doubled) neighbour, or interpolate
    full = pd.date_range(f"{year}-01-01", f"{year}-12-31 23:00", freq="h")
    s = s.reindex(full)
    h = pd.Timedelta("1h")
    for gap in s.index[s.isna()]:
        if gap.month != 10:
            continue
        split = False
        for nb, other in ((gap - h, gap - 2 * h), (gap + h, gap + 2 * h)):
            if nb in s.index and other in s.index and \
                    s[nb] > 1.6 * np.nanmean([s[other], s.get(gap + (gap - nb), np.nan)]):
                fixes.append(dict(series="Left bank", time=nb, value_MW=round(s[nb], 1),
                                  action="autumn DST: two hours summed in one row -> split in half"))
                s[gap] = s[nb] = s[nb] / 2
                split = True
                break
        if not split:
            fixes.append(dict(series="Left bank", time=gap, value_MW=np.nan,
                              action="autumn DST: missing hour after shift -> interpolated"))
    return s.interpolate(limit=3)


def read_load():
    rb = pd.read_excel(LOAD_FILE, sheet_name="Right bank", header=None)
    lb = pd.read_excel(LOAD_FILE, sheet_name="Left bank", header=None)
    fixes, out = [], {}
    for y in YEARS:
        r = read_right_bank(y, rb, fixes)
        l = read_left_bank(y, lb, fixes)
        if r is None or l is None:
            print(f"{y}: load data not found - skipped")
            continue
        r["LEFT"] = l
        out[y] = r
    return out, pd.DataFrame(fixes)


def read_temperature(year):
    path = TEMP_DIR / f"temperature_by_region_{year}.xlsx"
    if not path.exists():
        print(f"{year}: {path.name} not found - skipped")
        return None
    tabs = pd.read_excel(path, sheet_name=None, index_col=0)
    temps = {}
    for reg, df in tabs.items():
        df.index = pd.to_datetime(df.index)
        temps[reg] = df["region_mean_K"] - 273.15  # K -> deg C (average for the region for that hour)
    return pd.DataFrame(temps)


def match_regions(load_cols, temp_cols):
    """Load column name -> temperature tab name."""
    m = {c: c for c in load_cols if c in temp_cols}
    left = LEFT_BANK_TAB or next(iter(set(temp_cols) - set(m.values())), None)
    if left is None:
        sys.exit("ERROR: can't tell which temperature tab is the left bank - set LEFT_BANK_TAB.")
    m["LEFT"] = left
    return m


#MAIN
def main():
    OUT_DIR.mkdir(exist_ok=True)
    loads, fixes = read_load()
    year_rows, daily_all, regions = [], [], None
    bad_days = pd.to_datetime(list(EXCLUDE_DAYS))
    a_cols = [f"a{k}" for k in range(DEGREE, -1, -1)]

    # 1-2. daily values and one cubic per region per year
    for y, load in loads.items():
        temp = read_temperature(y)
        if temp is None:
            continue
        m = match_regions(list(load.columns), temp.columns)
        load = load.rename(columns=m)
        regs = list(m.values())
        regions = regions or regs

        dl = load[regs].resample("D").mean()
        dt = temp[regs].resample("D").mean()
        days = dl.index.intersection(dt.index)

        for r in regs:
            d = pd.DataFrame({"T_degC": dt.loc[days, r], "load_MW": dl.loc[days, r]}).dropna()
            reason = EXCLUDE.get((y, r), "")
            d["used_in_fit"] = (~d.index.isin(bad_days)) & (reason == "")
            daily_all.append(d.assign(region=r, year=y))
            if reason:
                year_rows.append(dict(region=r, year=y, status=f"excluded: {reason}"))
                print(f"{y} {r:>6}: excluded ({reason})")
                continue

            f = d[d.used_in_fit]
            T, L = f["T_degC"].values, f["load_MW"].values
            c = np.polyfit(T, L, DEGREE)
            mean_l = L.mean()
            r2 = 1 - np.sum((L - np.polyval(c, T)) ** 2) / np.sum((L - mean_l) ** 2)
            row = dict(region=r, year=y, status="used", n_days=len(f), R2=r2,
                       mean_load_MW=mean_l, T_min_degC=T.min(), T_max_degC=T.max())
            row.update(dict(zip(a_cols, c)))
            year_rows.append(row)
            print(f"{y} {r:>6}: R2 {r2:.3f}")

    if not year_rows:
        sys.exit("Nothing to fit - check the temperature workbooks exist.")
    by_year = pd.DataFrame(year_rows)
    daily = pd.concat(daily_all)
    daily.index.name = "date"

    # 3. average of the yearly cubics (MW coefficients), per region
    avg_rows = []
    for r in regions:
        f = by_year[(by_year.region == r) & (by_year.status == "used")]
        if f.empty:
            continue
        t_lo, t_hi = f.T_min_degC.max(), f.T_max_degC.min()        # range covered by every year
        c_avg = f[a_cols].mean().values
        row = dict(region=r, years_used=", ".join(map(str, f.year)), n_years=len(f),
                   T_valid_min_degC=t_lo, T_valid_max_degC=t_hi)
        row.update({f"avg_{k}": v for k, v in zip(a_cols, c_avg)})
        avg_rows.append(row)
    avg = pd.DataFrame(avg_rows)
    print()
    print(avg.round(6).to_string(index=False))

    #print so I can check the output in the console
    print("\nAverage polynomial per region (load in MW, T in deg C):")
    for _, a in avg.iterrows():
        a3, a2, a1, a0 = (a[f"avg_{k}"] for k in a_cols)
        print(f"  {a.region:>6}: load(T) = {a3:+.6g}*T^3 {a2:+.6g}*T^2 {a1:+.6g}*T {a0:+.6g}"
              f"   valid {a.T_valid_min_degC:.1f} to {a.T_valid_max_degC:.1f} deg C"
              f"   (years {a.years_used})")

    #writing output in excel
    path = OUT_DIR / "demand_temperature_cubic_2021_2025.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        avg.round(8).to_excel(xw, sheet_name="Average cubic", index=False)
        by_year.round(8).to_excel(xw, sheet_name="Cubic by year", index=False)
        daily.round(3).to_excel(xw, sheet_name="Daily data")
        fixes.to_excel(xw, sheet_name="Data fixes", index=False)
        for ws in xw.sheets.values():
            ws.freeze_panes = "B2"
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = \
                    max(10, min(60, len(str(col[0].value)) + 2))
    print(f"\nwritten {path}")


if __name__ == "__main__":
    main()
