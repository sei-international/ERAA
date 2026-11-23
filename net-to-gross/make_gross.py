"""
The purpose of this script is to convert the NET hourly regional electricity load in
"Historical_hourly_electricity_load.xlsx" to GROSS load
Net values in the provided data ("Historical_hourly_electricity_load.xlsx") already include distribution losses.

For every hour h of year y in region r:
    DG(h)    = solar_CF_r(h) x solar_MW(y) x solar_share_r  +  wind_CF_r(h) x wind_MW(y) x wind_share_r
    gross(h) = net(h) x (1 + TRANSMISSION_LOSS + OWN_USE) + DG(h) -- this is the governing equation

    solar_CF_r, wind_CF_r   mean hourly capacity factor of the region's grid cells
                            (solar_cf_by_region_<year>.xlsx / wind_cf_by_region_<year>.xlsx)
    solar_MW, wind_MW       reported distributed capacity in year y (settings below)
    *_share_r               region's share of distributed capacity (RES.xlsx, 'Connected in 2026', columns K and L)
    TRANSMISSION_LOSS       using an estimated value-- this should be replaced if possible
    OWN_USE                 own use of transmission-connected power plants, as a fraction of net load
                            (placeholder - replace). National quantity, spread over regions in
                            proportion to their load.

The left bank (Dnestroenergo) has no distributed capacity in RES.xlsx, so its DG is 0
and only the transmission losses and own use are added.
"""

#Imports

import sys
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.styles import Font, PatternFill

#Settings and global variables
BASE = Path(__file__).resolve().parent
LOAD_FILE = BASE / "Historical hourly electricity load.xlsx"
RES_FILE = BASE / "RES.xlsx"
CF_DIR = BASE                                  # folder with the capacity factor files
SOLAR_PATTERN = "solar_cf_by_region_{year}.xlsx"
WIND_PATTERN = "wind_cf_by_region_{year}.xlsx"
OUT_FILE = BASE / "Historical_hourly_electricity_load_gross.xlsx"            # net + gross + summary
OUT_FILE_GROSS_ONLY = BASE / "Historical_hourly_electricity_load_gross_only.xlsx"  # gross only, original layout
YEARS = [2021, 2022, 2023, 2024, 2025]

# Reported distributed generation capacity per year (MW).
# Distribution-connected only (same basis as the regional shares); not transmission-connected plants.
SOLAR_MW = {2021: 10, 2022: 37.8, 2023: 115.8, 2024: 146.1, 2025: 181}
WIND_MW = {2021: 87.6, 2022: 105.9, 2023: 121.3, 2024: 160.8, 2025: 160.8}

# stand in value for transmission losses
TRANSMISSION_LOSS = 0.02
# PLACEHOLDER own use of transmission-connected power plants (fraction of net load) - not sourced, replace
OWN_USE = 0.03
GROSS_FACTOR = 1 + TRANSMISSION_LOSS + OWN_USE   # both are fractions of the net load

RIGHT_REGIONS = ["A", "B", "C", "D", "E", "F"]
LEFT_CF_TAB = "Dnestroenergo"
# --------------------------------------------------------------------------


def check_capacities():
    """Stop if any year's capacity has not been filled in."""
    missing = [f"{name} {y}" for name, cap in (("SOLAR_MW", SOLAR_MW), ("WIND_MW", WIND_MW))
               for y in YEARS if cap.get(y) is None]
    if missing:
        sys.exit("ERROR: fill in the distributed capacity for: " + ", ".join(missing))


def to_timestamp(v):
    """Excel serial number or datetime -> pandas Timestamp (or None)."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return pd.Timestamp("1899-12-30") + pd.to_timedelta(v, unit="D")
    try:
        return pd.Timestamp(v)
    except (TypeError, ValueError):
        return None


def read_shares():
    """Regional shares of distributed solar (K) and wind (L) capacity."""
    ws = openpyxl.load_workbook(RES_FILE, data_only=True)["Connected in 2026"]
    solar, wind = {}, {}
    for row in range(3, 9):
        name = str(ws.cell(row, 6).value).replace("Region", "").strip()    # column F
        solar[name] = float(ws.cell(row, 11).value)                         # column K
        wind[name] = float(ws.cell(row, 12).value)                          # column L
    for label, s in (("solar", solar), ("wind", wind)):
        if abs(sum(s.values()) - 1) > 1e-6:
            sys.exit(f"ERROR: {label} shares in RES.xlsx do not add up to 1")
    return solar, wind


def read_cf_year(pattern, year):
    """Hourly capacity factor per region (mean of the region's cells), indexed by local
    UTC+2 hour, covering the local year (stitched from this and the previous file)."""
    parts = []
    for y in (year - 1, year):
        path = CF_DIR / pattern.format(year=y)
        if not path.exists():
            if y == year:
                sys.exit(f"ERROR: {path.name} not found")
            print(f"  {path.name} not found - first 2 local hours of {year} set to 0")
            continue
        tabs = pd.read_excel(path, sheet_name=None)
        df = pd.DataFrame({
            tab: d.drop(columns=["time_utc", "time_local_UTC+2"]).mean(axis=1).values
            for tab, d in tabs.items()
        }, index=pd.to_datetime(next(iter(tabs.values()))["time_local_UTC+2"]))
        parts.append(df)
    cf = pd.concat(parts)
    cf = cf[~cf.index.duplicated()]
    full = pd.date_range(f"{year}-01-01", f"{year}-12-31 23:00", freq="h")
    return cf.reindex(full).fillna(0.0)


def main():
    check_capacities()
    solar_share, wind_share = read_shares()
    wb = openpyxl.load_workbook(LOAD_FILE)
    rb, lb = wb["Right bank"], wb["Left bank"]
    rbg = wb.copy_worksheet(rb)
    rbg.title = "Right bank (gross)"
    lbg = wb.copy_worksheet(lb)
    lbg.title = "Left bank (gross)"

    # formulas in the 2021 block refer to 'Left bank' -> point them to the gross sheet
    for row in rbg.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("=") and "'Left bank'!" in c.value:
                c.value = c.value.replace("'Left bank'!", "'Left bank (gross)'!")

    diag = []
    for year in YEARS:
        print(f"{year}: reading capacity factors ...")
        cf_s = read_cf_year(SOLAR_PATTERN, year)
        cf_w = read_cf_year(WIND_PATTERN, year)

        # ---- right bank: find the year block (year in row 1, region names in row 2)
        col0 = next(c for c in range(1, rb.max_column + 1)
                    if isinstance(rb.cell(1, c).value, (int, float)) and int(rb.cell(1, c).value) == year)
        reg_cols = {}
        for c in range(col0 + 1, col0 + 7):
            reg_cols[str(rb.cell(2, c).value).replace("Region", "").strip()] = c
        for reg in RIGHT_REGIONS:
            c = reg_cols[reg]
            dg_s = cf_s[reg] * SOLAR_MW[year] * solar_share[reg]
            dg_w = cf_w[reg] * WIND_MW[year] * wind_share[reg]
            net_sum = gross_sum = dg_sum = 0.0
            neg_net = neg_gross = 0
            for row in range(3, rb.max_row + 1):
                ts = to_timestamp(rb.cell(row, col0).value)
                net = rb.cell(row, c).value
                if ts is None or not isinstance(net, (int, float)):
                    continue
                hour = ts.floor("h")
                dg = dg_s.get(hour, 0.0) + dg_w.get(hour, 0.0)
                gross = net * GROSS_FACTOR + dg
                rbg.cell(row, c).value = float(gross)
                net_sum += net
                gross_sum += gross
                dg_sum += dg
                neg_net += net < 0
                neg_gross += gross < 0
            diag.append(dict(year=year, region=reg, net_GWh=net_sum / 1000, DG_GWh=dg_sum / 1000,
                             gross_GWh=gross_sum / 1000, neg_hours_net=neg_net, neg_hours_gross=neg_gross))

        # LEFT BANK: no distributed capacity -> transmission losses and own use only
        lcol = next(c for c in range(1, lb.max_column + 1)
                    if isinstance(lb.cell(2, c).value, (int, float)) and int(lb.cell(2, c).value) == year)
        net_sum = gross_sum = 0.0
        for row in range(3, lb.max_row + 1):
            net = lb.cell(row, lcol + 1).value
            if isinstance(net, (int, float)):
                gross = net * GROSS_FACTOR
                lbg.cell(row, lcol + 1).value = float(gross)
                net_sum += net
                gross_sum += gross
        diag.append(dict(year=year, region="Left bank", net_GWh=net_sum / 1000, DG_GWh=0.0,
                         gross_GWh=gross_sum / 1000, neg_hours_net=0, neg_hours_gross=0))
    wb.save(OUT_FILE)
    d = pd.DataFrame(diag)
    print()
    print(d.round(1).to_string(index=False))
    print(f"\nwritten {OUT_FILE}")
    write_gross_only()
    print(f"written {OUT_FILE_GROSS_ONLY}")


def write_gross_only():
    """Copy of OUT_FILE with only the gross sheets, renamed 'Right bank' / 'Left bank'."""
    wb = openpyxl.load_workbook(OUT_FILE)
    for name in ("Right bank", "Left bank"):
        del wb[name]
    wb["Right bank (gross)"].title = "Right bank"
    wb["Left bank (gross)"].title = "Left bank"
    for row in wb["Right bank"].iter_rows():
        for c in row:
            if isinstance(c.value, str) and "'Left bank (gross)'!" in c.value:
                c.value = c.value.replace("'Left bank (gross)'!", "'Left bank'!")
    ws = wb.create_sheet("Gross notes")
    cap_text = ", ".join(f"{y}: {SOLAR_MW[y]} MW solar / {WIND_MW[y]} MW wind" for y in YEARS)
    lines = [
        "GROSS hourly load (MW), same layout as the original load workbook "
        "(2025-2021 blocks, Regions A-F, Left bank).",
        "gross = net x (1 + transmission loss + own use) + distributed solar + distributed wind. "
        "Made with make_gross.py.",
        "Distributed generation = regional capacity factor (mean of the region's grid cells) x "
        "reported distributed capacity of the year x regional share.",
        "Regional shares: RES.xlsx, 'Connected in 2026', columns K-L. Left bank: no distributed "
        "capacity listed -> DG = 0.",
        "Distribution losses are already in the net data and are not added again.",
        f"STAND-IN: transmission loss = {TRANSMISSION_LOSS:.1%}.",
        f"PLACEHOLDER: own use of transmission-connected power plants = {OWN_USE:.1%} of net load, "
        "spread over regions in proportion to their load.",
        "'Data notes (SEI)' describes the original NET data; its flagged cells carry over into these "
        "gross values.",
    ]
    for i, t in enumerate(lines, 1):
        ws.cell(i, 1, t).font = Font(name="Arial", bold=(i == 1))
    ws.column_dimensions["A"].width = 140
    wb.save(OUT_FILE_GROSS_ONLY)



if __name__ == "__main__":
    main()
