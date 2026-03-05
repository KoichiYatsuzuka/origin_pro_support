"""
Integration test: verify DataPlot.line_style and DataPlot.line_width.

Tests cover:
  [1]  line_style get/set roundtrip on a line plot
  [2]  line_width get/set roundtrip on a line plot
  [3]  line_width setter rejects non-positive values (ValueError)
  [4]  line_style setter raises ValueError on a scatter-only plot (no line)
  [5]  line_width setter raises ValueError on a scatter-only plot (no line)
  [6]  line_style and line_width work on a line+symbol plot

Run from repo root:
    python test_codes/test_line.py
"""
import sys
import os

REPO_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, REPO_ROOT)

import origin_pro_support as ops
from origin_pro_support.layer.enums import XYPlotType, GroupMode, LineStyle
from origin_pro_support.layer.graph_layer import DataPlot
from origin_pro_support.base import OriginCommandResponceError


def run():
    proj_path = os.path.join(REPO_ROOT, "test_codes", "test_line.opju")
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
    wbook = origin.new_workbook("LineTestBook")
    assert wbook is not None
    ws = wbook.get_layer(0)
    assert ws is not None

    x_data = [1.0, 2.0, 3.0, 4.0, 5.0]
    y_data = [2.0, 4.0, 3.0, 5.0, 1.0]
    x_col = ws.add_column_from_data(x_data, lname="X")
    y_col = ws.add_column_from_data(y_data, lname="Y")

    # ── line graph ────────────────────────────────────────────────────────
    line_page = origin.new_graph("LineTestLine", XYPlotType.LINE)
    assert line_page is not None
    line_layer = line_page.get_layer(0)
    assert line_layer is not None

    line_plot = line_layer.add_xy_plot(ws, x_col.index, y_col.index, XYPlotType.LINE)
    assert isinstance(line_plot, DataPlot)

    line_layer.group_plots(GroupMode.NONE)

    # ── [1] line_style get/set roundtrip ─────────────────────────────────
    print("\n[1] line_style get/set roundtrip ...")
    for style in (LineStyle.SOLID, LineStyle.DASH, LineStyle.DOT,
                  LineStyle.DASH_DOT, LineStyle.DASH_DOT_DOT):
        line_plot.line_style = style
        got = line_plot.line_style
        print(f"  set {style.name} -> got {got.name}")
        assert got == style, f"Expected {style}, got {got}"
    print("  OK")

    # ── [2] line_width get/set roundtrip ──────────────────────────────────
    print("\n[2] line_width get/set roundtrip ...")
    for width in (1.0, 2.0, 0.5):
        line_plot.line_width = width
        got = line_plot.line_width
        print(f"  set {width} -> got {got}")
        assert got == width, f"Expected {width}, got {got}"
    print("  OK")

    # ── [3] line_width rejects non-positive values ─────────────────────
    print("\n[3] line_width rejects non-positive values ...")
    for bad in (0.0, -1.0, -0.1):
        try:
            line_plot.line_width = bad
            assert False, f"Expected ValueError for line_width={bad} was NOT raised"
        except ValueError as e:
            print(f"  ValueError for {bad}: {e}")
    print("  OK")

    # ── scatter graph (no line) ────────────────────────────────────────────
    scatter_page = origin.new_graph("LineTestScatter", XYPlotType.SCATTER)
    assert scatter_page is not None
    scatter_layer = scatter_page.get_layer(0)
    assert scatter_layer is not None

    scatter_plot = scatter_layer.add_xy_plot(ws, x_col.index, y_col.index, XYPlotType.SCATTER)
    assert isinstance(scatter_plot, DataPlot)

    scatter_layer.group_plots(GroupMode.NONE)

    # ── [4] line_style on scatter-only plot must raise ValueError ─────────
    print("\n[4] line_style raises ValueError on scatter-only plot ...")
    try:
        scatter_plot.line_style = LineStyle.SOLID
        assert False, "Expected ValueError for line_style setter on scatter plot was NOT raised"
    except ValueError as e:
        print(f"  line_style setter raised ValueError as expected: {e}")

    try:
        _ = scatter_plot.line_style
        assert False, "Expected ValueError for line_style getter on scatter plot was NOT raised"
    except ValueError as e:
        print(f"  line_style getter raised ValueError as expected: {e}")
    print("  OK")

    # ── [5] line_width on scatter-only plot must raise ValueError ─────────
    print("\n[5] line_width raises ValueError on scatter-only plot ...")
    try:
        scatter_plot.line_width = 1.0
        assert False, "Expected ValueError for line_width setter on scatter plot was NOT raised"
    except ValueError as e:
        print(f"  line_width setter raised ValueError as expected: {e}")

    try:
        _ = scatter_plot.line_width
        assert False, "Expected ValueError for line_width getter on scatter plot was NOT raised"
    except ValueError as e:
        print(f"  line_width getter raised ValueError as expected: {e}")
    print("  OK")

    # ── line+symbol graph ─────────────────────────────────────────────────
    ls_page = origin.new_graph("LineTestLineSymbol", XYPlotType.LINE_SYMBOL)
    assert ls_page is not None
    ls_layer = ls_page.get_layer(0)
    assert ls_layer is not None

    ls_plot = ls_layer.add_xy_plot(ws, x_col.index, y_col.index, XYPlotType.LINE_SYMBOL)
    assert isinstance(ls_plot, DataPlot)

    ls_layer.group_plots(GroupMode.NONE)

    # ── [6] line_style and line_width work on line+symbol plot ───────────
    print("\n[6] line_style and line_width on line+symbol plot ...")
    ls_plot.line_style = LineStyle.DASH
    got_style = ls_plot.line_style
    print(f"  line_style set DASH -> got {got_style.name}")
    assert got_style == LineStyle.DASH, f"Expected DASH, got {got_style}"

    ls_plot.line_width = 2.0
    got_width = ls_plot.line_width
    print(f"  line_width set 2.0 -> got {got_width}")
    assert got_width == 2.0, f"Expected 2.0, got {got_width}"
    print("  OK")


if __name__ == "__main__":
    sys.exit(run())
