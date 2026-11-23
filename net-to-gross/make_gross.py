"""
The purpose of this script is to convert the NET hourly regional electricity load in "Historical_hourly_electricity_load.xlsx" to GROSS load, by adding distributed generation (DG) and transmission losses.

For every hour h of year y in region r:
    DG(h)    = solar_CF_r(h) x solar_MW(y) x solar_share_r  +  wind_CF_r(h) x wind_MW(y) x wind_share_r
    gross(h) = (net(h) + DG(h)) x (1 + TRANSMISSION_LOSS)

    solar_CF_r, wind_CF_r   mean hourly capacity factor of the region's grid cells
                            (solar_cf_by_region_<year>.xlsx / wind_cf_by_region_<year>.xlsx)
    solar_MW, wind_MW       total distributed capacity in year y (settings below)
    *_share_r               region's share of distributed capacity
                            (RES.xlsx, 'Connected in 2026', columns K and L)
    TRANSMISSION_LOSS       stand-in value (settings below) - replace with a sourced figure

The left bank (Dnestroenergo) has no distributed capacity in RES.xlsx, so its DG is 0 and
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
LOAD_FILE = BASE / "Historical_hourly_electricity_load.xlsx"
RES_FILE = BASE / "RES.xlsx"
CF_DIR = BASE                                  # folder with the capacity factor files
SOLAR_PATTERN = "solar_cf_by_region_{year}.xlsx"
WIND_PATTERN = "wind_cf_by_region_{year}.xlsx"
OUT_FILE = BASE / "Historical_hourly_electricity_load_gross.xlsx"            # net + gross + summary
OUT_FILE_GROSS_ONLY = BASE / "Historical_hourly_electricity_load_gross_only.xlsx"  # gross only, original layout
YEARS = [2021, 2022, 2023, 2024, 2025]

# Distributed capacity (MW). 
SOLAR_MW_LATEST = 793.38
WIND_MW_LATEST = 220.40
# Capacity in each year as a fraction of the latest value.
# STAND-IN: national trajectory from the NECP table in RES.xlsx (sheet 'NECP'):
#   solar = PV including prosumers, wind = wind power plants, each divided by its 2025 value.
NECP_SOLAR = {2021: 19.9, 2022: 68.1, 2023: 199.9, 2024: 395.4, 2025: 430.3}
NECP_WIND = {2021: 87.6, 2022: 105.9, 2023: 121.3, 2024: 160.8, 2025: 160.8}
SOLAR_MW = {y: SOLAR_MW_LATEST * NECP_SOLAR[y] / NECP_SOLAR[2025] for y in YEARS}
WIND_MW = {y: WIND_MW_LATEST * NECP_WIND[y] / NECP_WIND[2025] for y in YEARS}

# stand in value for transmission losses
TRANSMISSION_LOSS = 0.02

RIGHT_REGIONS = ["A", "B", "C", "D", "E", "F"]
LEFT_CF_TAB = "Dnestroenergo"                    
# --------------------------------------------------------------------------


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
                gross = (net + dg) * (1 + TRANSMISSION_LOSS)
                rbg.cell(row, c).value = float(gross)
                net_sum += net
                gross_sum += gross
                dg_sum += dg
                neg_net += net < 0
                neg_gross += gross < 0
            diag.append(dict(year=year, region=reg, net_GWh=net_sum / 1000, DG_GWh=dg_sum / 1000,
                             gross_GWh=gross_sum / 1000, neg_hours_net=neg_net, neg_hours_gross=neg_gross))

        # ---- left bank: no distributed capacity -> transmission losses only
        lcol = next(c for c in range(1, lb.max_column + 1)
                    if isinstance(lb.cell(2, c).value, (int, float)) and int(lb.cell(2, c).value) == year)
        net_sum = gross_sum = 0.0
        for row in range(3, lb.max_row + 1):
            net = lb.cell(row, lcol + 1).value
            if isinstance(net, (int, float)):
                gross = net * (1 + TRANSMISSION_LOSS)
                lbg.cell(row, lcol + 1).value = float(gross)
                net_sum += net
                gross_sum += gross
        diag.append(dict(year=year, region="Left bank", net_GWh=net_sum / 1000, DG_GWh=0.0,
                         gross_GWh=gross_sum / 1000, neg_hours_net=0, neg_hours_gross=0))

    write_summary(wb, rb, lb, solar_share, wind_share, pd.DataFrame(diag))
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
    for name in ("Right bank", "Left bank", "Gross summary"):
        del wb[name]
    wb["Right bank (gross)"].title = "Right bank"
    wb["Left bank (gross)"].title = "Left bank"
    for row in wb["Right bank"].iter_rows():
        for c in row:
            if isinstance(c.value, str) and "'Left bank (gross)'!" in c.value:
                c.value = c.value.replace("'Left bank (gross)'!", "'Left bank'!")
    ws = wb.create_sheet("Gross notes")
    lines = [
        "GROSS hourly load (MW), same layout as the original load workbook "
        "(2025-2021 blocks, Regions A-F, Left bank).",
        "gross = (net + distributed solar + distributed wind) x (1 + transmission loss). "
        "Made with make_gross_load.py.",
        "Distributed generation = regional capacity factor (mean of the region's grid cells) x "
        "national distributed capacity x regional share.",
        "Regional shares: RES.xlsx, 'Connected in 2026', columns K-L. Left bank: no distributed "
        "capacity listed -> DG = 0.",
        "Distribution losses are already in the net data and are not added again.",
        f"STAND-IN: transmission loss = {TRANSMISSION_LOSS:.1%}).",
        f"STAND-IN: capacity per year = {SOLAR_MW_LATEST} MW solar / {WIND_MW_LATEST} MW wind x NECP "
        "trajectory year/2025 (RES.xlsx 'NECP'). Replace with actual yearly capacity when available.",
        "'Data notes (SEI)' describes the original NET data; its flagged cells carry over into these "
        "gross values.",
        "2021 block: AJ = sum of gross Regions A-F + gross Left bank; AK = original hard-coded total "
        "(unchanged); AL = AK - AJ.",
    ]
    for i, t in enumerate(lines, 1):
        ws.cell(i, 1, t).font = Font(name="Arial", bold=(i == 1))
    ws.column_dimensions["A"].width = 140
    wb.save(OUT_FILE_GROSS_ONLY)


def write_summary(wb, rb, lb, solar_share, wind_share, diag):
    ws = wb.create_sheet("Gross summary")
    bold, arial = Font(name="Arial", bold=True), Font(name="Arial")
    blue = Font(name="Arial", color="0000FF")
    yellow = PatternFill("solid", fgColor="FFFF00")
    put = lambda r, c, v, f=arial: setattr(ws.cell(r, c, v), "font", f)

    put(1, 1, "Gross load: gross = (net + distributed solar + distributed wind) x (1 + transmission loss)", bold)
    put(2, 1, "Values in this sheet's input table were used by make_gross_load.py. "
              "To change them, edit the script settings and run it again.")

    put(4, 1, "Inputs", bold)
    put(5, 1, "Transmission loss (fraction of net + DG)")
    put(5, 2, TRANSMISSION_LOSS, blue)
    ws.cell(5, 2).number_format = "0.0%"
    ws.cell(5, 2).fill = yellow
    put(5, 3, "STAND-IN value, not from a source. Replace with a sourced figure (e.g. Moldelectrica or ANRE annual report).")

    put(7, 1, "Year", bold)
    put(7, 2, "Distributed solar (MW)", bold)
    put(7, 3, "Distributed wind (MW)", bold)
    put(7, 4, "Source", bold)
    for i, y in enumerate(YEARS):
        put(8 + i, 1, str(y))
        put(8 + i, 2, round(SOLAR_MW[y], 2), blue)
        put(8 + i, 3, round(WIND_MW[y], 2), blue)
        ws.cell(8 + i, 2).fill = ws.cell(8 + i, 3).fill = yellow
        put(8 + i, 4, f"Latest capacity from user ({SOLAR_MW_LATEST} MW solar, {WIND_MW_LATEST} MW wind) x NECP "
                      f"trajectory {y}/2025 (RES.xlsx 'NECP'). STAND-IN - replace with actual capacity per year.")

    r0 = 8 + len(YEARS) + 1
    put(r0, 1, "Region", bold)
    put(r0, 2, "Solar share", bold)
    put(r0, 3, "Wind share", bold)
    put(r0, 4, "Source", bold)
    for i, reg in enumerate(RIGHT_REGIONS + ["Left bank"]):
        put(r0 + 1 + i, 1, reg)
        put(r0 + 1 + i, 2, solar_share.get(reg, 0.0), blue)
        put(r0 + 1 + i, 3, wind_share.get(reg, 0.0), blue)
        ws.cell(r0 + 1 + i, 2).number_format = ws.cell(r0 + 1 + i, 3).number_format = "0.0%"
        put(r0 + 1 + i, 4, "RES.xlsx, 'Connected in 2026', columns K-L (distribution-connected)"
            if reg != "Left bank" else "No distributed capacity listed in RES.xlsx -> DG = 0")

    # ---- annual energy table: net and gross with formulas, DG from the script
    t0 = r0 + 9
    put(t0, 1, "Annual energy (GWh)", bold)
    heads = ["Year", "Region", "Net (GWh)", "Gross (GWh)", "Gross - net (GWh)",
             "of which DG (GWh, script)", "Negative hours, net", "Negative hours, gross"]
    for j, h in enumerate(heads):
        put(t0 + 1, 1 + j, h, bold)
    r = t0 + 2
    for year in YEARS:
        col0 = next(c for c in range(1, rb.max_column + 1)
                    if isinstance(rb.cell(1, c).value, (int, float)) and int(rb.cell(1, c).value) == year)
        lcol = next(c for c in range(1, lb.max_column + 1)
                    if isinstance(lb.cell(2, c).value, (int, float)) and int(lb.cell(2, c).value) == year)
        cols = {str(rb.cell(2, c).value).replace("Region", "").strip(): ("Right bank", c)
                for c in range(col0 + 1, col0 + 7)}
        cols["Left bank"] = ("Left bank", lcol + 1)
        for reg in RIGHT_REGIONS + ["Left bank"]:
            sheet, c = cols[reg]
            L = openpyxl.utils.get_column_letter(c)
            rng = f"{L}3:{L}{rb.max_row}"
            dg = diag[(diag.year == year) & (diag.region == reg)].iloc[0]
            put(r, 1, str(year))
            put(r, 2, reg)
            put(r, 3, f"=SUM('{sheet}'!{rng})/1000")
            put(r, 4, f"=SUM('{sheet} (gross)'!{rng})/1000")
            put(r, 5, f"=D{r}-C{r}")
            put(r, 6, round(float(dg.DG_GWh), 3))
            put(r, 7, f"=COUNTIF('{sheet}'!{rng},\"<0\")")
            put(r, 8, f"=COUNTIF('{sheet} (gross)'!{rng},\"<0\")")
            for j in range(3, 7):
                ws.cell(r, j).number_format = "#,##0.0"
            r += 1
    ws.column_dimensions["A"].width = 38
    for L in "BCDEFGH":
        ws.column_dimensions[L].width = 22


if __name__ == "__main__":
    main()
