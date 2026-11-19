"""
The purpose of script is to return hourly wind capacity factors (CF) for EVERY ERA5 grid cell, 2010-2026,
organised by region (A-F + Left Bank).

This script assumes the use of 
Vestas V90-2.0 MW GridStreamer power curve (OEP id 133), 
hub height 105 m, cut-in 4 m/s, rated 12 m/s, cut-out 25 m/s”, 
with smoothing applied on the power curve to account for wind farm effects (https://repository.utm.md/bitstream/handle/5014/8142/SIELMEN_2015_p543_546.pdf?sequence=1&isAllowed=y)"""

from pathlib import Path 
from assigning_cells_to_regions import run_regional

# ---------------------------------------------------------------- settings
BASE = Path(__file__).resolve().parent                 #RES availability curves
CUTOUT_DIR = BASE.parent / "cutouts"
OUT_DIR = BASE / "windcf_output"
REGIONS_FILE = BASE.parent / "Moldova_ERAA_regions_full_partition.geojson"
YEARS = range(2010, 2027)
LOCAL_TZ = "Etc/GMT-2"                                 #fixed UTC+2
MIN_OVERLAP = 0.25
EFFICIENCY_FACTOR = 0.875                            #overall losses ~12.5% in Moldova (https://repository.utm.md/bitstream/handle/5014/8142/SIELMEN_2015_p543_546.pdf?sequence=1&isAllowed=y)
# --------------------------------------------------------------------------
from atlite.resource import get_oedb_windturbineconfig
TURBINE = get_oedb_windturbineconfig(133) # 133 is Vestas V90/2000/GS, which is the same turbine model used in (https://repository.utm.md/bitstream/handle/5014/8142/SIELMEN_2015_p543_546.pdf?sequence=1&isAllowed=y)
TURBINE["hub_height"] = 105   # m, from the paper

import numpy as np

V = np.array(TURBINE["V"], dtype=float)
POW = np.array(TURBINE["POW"], dtype=float)
keep = V <= 25
V = np.concatenate(([0.0, 3.99], V[keep], [25.01]))
POW = np.concatenate(([0.0, 0.0], POW[keep], [0.0]))

TURBINE["V"], TURBINE["POW"] = V, POW


def wind_cf(cutout):
    """Hourly wind CF for every grid cell of one cutout."""
    cf = cutout.wind(turbine=TURBINE,
                     interpolation_method="power",     #hourly shear exponent
                     smooth=True,                      #wind-farm power curve
                     aggregate_time=None)              #keep every hour (does not generate yearly average)
    return cf * EFFICIENCY_FACTOR
 
 
if __name__ == "__main__":
    run_regional(make_cf=wind_cf, name="wind", years=YEARS,
             cutout_dir=CUTOUT_DIR, out_dir=OUT_DIR, regions_file=REGIONS_FILE,
             attrs={"turbine": "Vestas V90/2000/GS (OEP id 133)",
                    "hub_height_m": TURBINE["hub_height"],
                    "cut_in_cut_out_m_s": "4 / 25",
                    "efficiency_factor": EFFICIENCY_FACTOR})