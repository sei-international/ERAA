"""
The purpose of script is to return hourly solar capacity factors (CF) for EVERY ERA5 grid cell, 2010-2026,
organised by region (A-F + Left Bank)."""

from pathlib import Path
from assigning_cells_to_regions import run_regional
 
# ---------------------------------------------------------------- settings
BASE = Path(__file__).resolve().parent                 # RES availability curves
CUTOUT_DIR = BASE.parent / "cutouts"
OUT_DIR = BASE / "solar_cf_output"                     # solar results in their own folder
REGIONS_FILE = BASE.parent / "Moldova_ERAA_regions_full_partition.geojson"
YEARS = range(2010, 2027)
LOCAL_TZ = "Etc/GMT-2"                                 # fixed UTC+2
MIN_OVERLAP = 0.25
 
PANEL = "CSi"                       #crystalline silicon, Huld (2010) model
ORIENTATION = "latitude_optimal"    #fixed, south-facing, optimal tilt
TRACKING = None                     #None = fixed; "horizontal" = single-axis
EFFICIENCY_FACTOR = 0.956           #default total system loss is 14% for Moldova (https://re.jrc.ec.europa.eu/pvg_tools/en/tools.html. 
                                    #atlite's default interverter efficiency is 90% (which is already accounted for in the method), so the remaining losses are 4% (0.9*0.96 = 0.864, ~14% loss)
# --------------------------------------------------------------------------
 
 
def solar_cf(cutout):
    """Hourly solar PV CF for every grid cell of one cutout."""
    cf = cutout.pv(panel=PANEL,
                   orientation=ORIENTATION,
                   tracking=TRACKING,
                   aggregate_time=None)                # keep every hour
    return cf * EFFICIENCY_FACTOR
 
 
if __name__ == "__main__":
    run_regional(make_cf=solar_cf, name="solar", years=YEARS,
                 cutout_dir=CUTOUT_DIR, out_dir=OUT_DIR, regions_file=REGIONS_FILE,
                 attrs={"panel": PANEL, "orientation": str(ORIENTATION),
                        "tracking": str(TRACKING), "efficiency_factor": EFFICIENCY_FACTOR})
 