"""
Probe: DataObjectBase.GetDatasetName() on DataPlot obtained via GraphLayer.GetDataPlots()
"""
import sys
import os

REPO_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, REPO_ROOT)

import origin_pro_support as ops
from origin_pro_support.layer.enums import XYPlotType

proj_path = os.path.join(REPO_ROOT, "test_codes", "test_datasetname.opju")
origin = ops.OriginInstance(proj_path)

try:
    # ── create workbook with data ────────────────────────────────────────────
    wb = origin.new_workbook("ProbeBook")
    ws = wb.get_layer(0)
    ws.add_column_from_data([1.0, 2.0, 3.0, 4.0, 5.0], lname="X")
    ws.add_column_from_data([10.0, 20.0, 30.0, 40.0, 50.0], lname="Y")

    # ── create graph and plot ────────────────────────────────────────────────
    gp = origin.new_graph("ProbeGraph", XYPlotType.SCATTER)
    gl = gp.get_layer(0)
    gl.add_xy_plot(ws, 0, 1, XYPlotType.SCATTER)

    # ── get raw OriginExt GraphLayer ─────────────────────────────────────────
    raw_gl = gl._obj
    print("GraphLayer type:", type(raw_gl))

    # Get DataPlots collection
    dp_col = raw_gl.GetDataPlots()
    print("DataPlots collection type:", type(dp_col))
    print("DataPlots count (len):", len(dp_col))

    # Try index 0
    print("\n--- dp_col[0] ---")
    try:
        dp = dp_col[0]
        print("  type:", type(dp))
        print("  GetDatasetName():", repr(dp.GetDatasetName()))
        print("  DatasetName prop:", repr(dp.DatasetName))
        print("  GetName():", repr(dp.GetName()))
        print("  GetRange():", repr(dp.GetRange()))
    except Exception as e:
        print("  ERROR with index 0:", e)

    # Try index 1
    print("\n--- dp_col[1] ---")
    try:
        dp1 = dp_col[1]
        print("  type:", type(dp1))
        print("  GetDatasetName():", repr(dp1.GetDatasetName()))
    except Exception as e:
        print("  ERROR with index 1:", e)

    # Try iteration
    print("\n--- iteration ---")
    try:
        for i, dp_i in enumerate(dp_col):
            print(f"  [{i}] GetDatasetName():", repr(dp_i.GetDatasetName()))
            print(f"  [{i}] GetRange():", repr(dp_i.GetRange()))
    except Exception as e:
        print("  iteration ERROR:", e)

finally:
    origin.close(save_flag=True)
    print("\nDone")
