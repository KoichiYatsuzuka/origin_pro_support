"""
Integration test: verify DataPlot.symbol_size and DataPlot.symbol_kind.

Tests cover:
  [1]  symbol_size get/set roundtrip on a scatter plot
  [2]  symbol_size setter rejects non-positive values (ValueError)
  [3]  symbol_kind get/set roundtrip on a scatter plot for several MarkerShape values
  [4]  symbol_kind setter raises an error when called on a line-only plot (no marker)

Run from repo root:
    python test_codes/test_marker.py
"""
import sys
import os

REPO_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, REPO_ROOT)

import origin_pro_support as ops
from origin_pro_support.layer.enums import XYPlotType, GroupMode
from origin_pro_support.layer.graph_layer import DataPlot
from origin_pro_support.layer.enums import MarkerShape
from origin_pro_support.base import OriginCommandResponceError


def run():
    proj_path = os.path.join(REPO_ROOT, "test_codes", "test_marker.opju")
    if os.path.exists(proj_path):
        os.remove(proj_path)
        print(f"Removed existing project: {proj_path}")

    print("=== Starting Origin ===")
    origin = ops.OriginInstance(proj_path)
    try:
        _test_body(origin)
        print("\n=== ALL CHECKS PASSED ===")
        return 0
    except Exception as e:
        print(f"\n=== FAILED: {e} ===")
        import traceback; traceback.print_exc()
        return 1
    finally:
        origin.save()
        origin.close()
        print("Origin closed.")


def _test_body(origin: ops.OriginInstance):
    wbook = origin.new_workbook("MarkerTestBook")
    assert wbook is not None
    ws = wbook.get_layer(0)
    assert ws is not None

    x_data = [1.0, 2.0, 3.0, 4.0, 5.0]
    y_data = [2.0, 4.0, 3.0, 5.0, 1.0]
    x_col = ws.add_column_from_data(x_data, lname="X")
    y_col = ws.add_column_from_data(y_data, lname="Y")

    # ── scatter graph (has symbols) ───────────────────────────────────────
    scatter_page = origin.new_graph("MarkerTestScatter", XYPlotType.SCATTER)
    assert scatter_page is not None
    scatter_layer = scatter_page.get_layer(0)
    assert scatter_layer is not None

    scatter_plot = scatter_layer.add_xy_plot(ws, x_col.index, y_col.index, XYPlotType.SCATTER)
    assert isinstance(scatter_plot, DataPlot)

    scatter_layer.group_plots(GroupMode.NONE)

    # ── [1] symbol_size get/set roundtrip ────────────────────────────────
    print("\n[1] symbol_size get/set roundtrip ...")
    for size in (5.0, 10.0, 3.5):
        scatter_plot.symbol_size = size
        got = scatter_plot.symbol_size
        print(f"  set {size} -> got {got}")
        assert got == size, f"Expected {size}, got {got}"
    print("  OK")

    # ── [2] symbol_size rejects non-positive values ───────────────────────
    print("\n[2] symbol_size rejects non-positive values ...")
    for bad in (0.0, -1.0, -0.1):
        try:
            scatter_plot.symbol_size = bad
            assert False, f"Expected ValueError for symbol_size={bad} was NOT raised"
        except ValueError as e:
            print(f"  ValueError for {bad}: {e}")
    print("  OK")

    # ── [3] symbol_kind get/set roundtrip ─────────────────────────────────
    print("\n[3] symbol_kind get/set roundtrip ...")
    for shape in (MarkerShape.SQUARE, MarkerShape.CIRCLE, MarkerShape.TRIANGLE_UP,
                  MarkerShape.DIAMOND, MarkerShape.STAR):
        scatter_plot.symbol_kind = shape
        got = scatter_plot.symbol_kind
        print(f"  set {shape.name} -> got {got.name}")
        assert got == shape, f"Expected {shape}, got {got}"
    print("  OK")

    # ── line graph (no symbols) ────────────────────────────────────────────
    line_page = origin.new_graph("MarkerTestLine", XYPlotType.LINE)
    assert line_page is not None
    line_layer = line_page.get_layer(0)
    assert line_layer is not None

    line_plot = line_layer.add_xy_plot(ws, x_col.index, y_col.index, XYPlotType.LINE)
    assert isinstance(line_plot, DataPlot)

    line_layer.group_plots(GroupMode.NONE)

    # ── [4] symbol_kind/symbol_size on line-only plot must raise ValueError ─
    print("\n[4] symbol_kind/symbol_size raises ValueError on line-only plot ...")
    try:
        line_plot.symbol_kind = MarkerShape.CIRCLE
        assert False, "Expected ValueError for symbol_kind on line plot was NOT raised"
    except ValueError as e:
        print(f"  symbol_kind setter raised ValueError as expected: {e}")

    try:
        _ = line_plot.symbol_kind
        assert False, "Expected ValueError for symbol_kind getter on line plot was NOT raised"
    except ValueError as e:
        print(f"  symbol_kind getter raised ValueError as expected: {e}")

    try:
        line_plot.symbol_size = 10.0
        assert False, "Expected ValueError for symbol_size on line plot was NOT raised"
    except ValueError as e:
        print(f"  symbol_size setter raised ValueError as expected: {e}")

    try:
        _ = line_plot.symbol_size
        assert False, "Expected ValueError for symbol_size getter on line plot was NOT raised"
    except ValueError as e:
        print(f"  symbol_size getter raised ValueError as expected: {e}")
    print("  OK")


if __name__ == "__main__":
    sys.exit(run())
