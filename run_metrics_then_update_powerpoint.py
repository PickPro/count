#!/usr/bin/env python3
"""Runs the existing classification script, then fills slide 2 from its Excel output.

Put all of these in the same Scripts folder:
- metric_level_color_counts_hardcoded.py   (existing working script)
- populate_slide2_from_excel.py             (PowerPoint updater)
- Final Deck.pptx                           (original PowerPoint)

The first script keeps its own hardcoded source Excel paths, as already
configured on your workstation. The second script reads the generated
'LatestFile_with_levels_and_summary.xlsx' from the Scripts folder.

Run:
    python run_metrics_then_update_powerpoint.py
"""

from pathlib import Path
import runpy
import sys

FOLDER = Path(__file__).resolve().parent
FIRST_SCRIPT = FOLDER / "metric_level_color_counts_hardcoded.py"
POWERPOINT_SCRIPT = FOLDER / "populate_slide2_from_excel.py"

for path in (FIRST_SCRIPT, POWERPOINT_SCRIPT):
    if not path.is_file():
        raise FileNotFoundError(f"Missing required script: {path}")

original_argv = sys.argv
try:
    print("STEP 1/2: Generate metric level/color summary...")
    sys.argv = [str(FIRST_SCRIPT)]
    runpy.run_path(str(FIRST_SCRIPT), run_name="__main__")

    print("STEP 2/2: Read generated Excel and populate PowerPoint slide 2...")
    sys.argv = [str(POWERPOINT_SCRIPT)]
    runpy.run_path(str(POWERPOINT_SCRIPT), run_name="__main__")
finally:
    sys.argv = original_argv

print("COMPLETE: Summary workbook generated and PowerPoint updated.")
