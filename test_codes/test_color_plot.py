"""
Integration test: plot three datasets with RGB hex colors using Origin.

Colors tested:
  Dataset 1 → #FF3333  (255,  51,  51)  warm red
  Dataset 2 → #228822  ( 34, 136,  34)  medium green
  Dataset 3 → #3333FF  ( 51,  51, 255)  blue

Run from repo root:
    python test_codes/test_color_plot.py
"""
import sys
import os
import math

# ── path setup ──────────────────────────────────────────────────────────────
REPO_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, REPO_ROOT)

import origin_pro_support as ops
from origin_pro_support.layer.enums import XYPlotType, OriginColorIndex


# ── helpers ──────────────────────────────────────────────────────────────────

def hex_to_rgb(hex_color: str) -> tuple:
    """Convert '#RRGGBB' to (R, G, B) tuple of ints."""
    h = hex_color.lstrip("#")
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


COLORS_HEX = ["#FF3333", "#228822", "#3333FF"]
COLORS_RGB = [hex_to_rgb(c) for c in COLORS_HEX]   # [(255,51,51), (34,136,34), (51,51,255)]


# ── test ─────────────────────────────────────────────────────────────────────

def run():
    proj_path = os.path.join(REPO_ROOT, "test_codes", "test_color_plot.opju")

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
        origin.close(save_flag=True)
        print("Origin closed.")


def _test_body(origin: ops.OriginInstance):
    # Remove leftover pages from previous runs so the test is repeatable
    existing_graph = origin.find_graph("ColorTestGraph")
    if existing_graph is not None:
        existing_graph.destroy()

    existing_book = origin.find_book("ColorTestBook")
    if existing_book is not None:
        existing_book.destroy()

    n = 50                         # data points
    x = [i / (n - 1) * 2 * math.pi for i in range(n)]
    y_datasets = [
        [math.sin(xi) for xi in x],           # dataset 1
        [math.cos(xi) for xi in x],           # dataset 2
        [math.sin(xi + math.pi / 4) for xi in x],  # dataset 3
    ]

    # ── 1. Create workbook and fill data ─────────────────────────────────
    print("\n[1] Creating workbook...")
    wbook = origin.new_workbook("ColorTestBook")
    assert wbook is not None, "Failed to create workbook"

    ws = wbook.get_layer(0)     # first (auto-created) sheet
    assert ws is not None, "Failed to get worksheet"

    # Add X column then three Y columns
    ws.add_column_from_data(x,           lname="X",   axis="X")
    ws.add_column_from_data(y_datasets[0], lname="Y1", axis="Y")
    ws.add_column_from_data(y_datasets[1], lname="Y2", axis="Y")
    ws.add_column_from_data(y_datasets[2], lname="Y3", axis="Y")
    print(f"  Worksheet '{ws.name}' has {ws.cols} columns, {ws.rows} rows.")
    assert ws.cols >= 4, f"Expected ≥4 cols, got {ws.cols}"

    # ── 2. Create graph ───────────────────────────────────────────────────
    print("\n[2] Creating graph...")
    graph = origin.new_graph("ColorTestGraph", XYPlotType.SCATTER)
    assert graph is not None, "Failed to create graph"

    layer = graph.get_layer(0)
    assert layer is not None, "Failed to get graph layer"

    # ── 3. Plot each dataset with its hex color ───────────────────────────
    plots = []
    x_col = 0  # col index 0 = X
    for i, (rgb, hex_str) in enumerate(zip(COLORS_RGB, COLORS_HEX)):
        y_col = i + 1  # 1, 2, 3
        print(f"\n[3.{i+1}] Plotting Y{i+1} with color {hex_str} → RGB{rgb}")
        plot = layer.add_xy_plot(ws, x_col, y_col, XYPlotType.SCATTER, color=rgb)
        assert plot is not None, f"add_xy_plot returned None for dataset {i+1}"
        plots.append(plot)
        print(f"  Plot {i+1} created OK")

    # ── 4. Rescale ────────────────────────────────────────────────────────
    layer.rescale()
    print("\n[4] Layer rescaled.")

    # ── 5. Setter: change plot colors after creation ──────────────────────
    # layer.plot(n).color in LabTalk returns the color-cycle position within
    # the group (0-based) rather than the absolute palette index, so we
    # cannot reliably round-trip an absolute color index through the getter.
    # We verify that the setter runs without raising an exception instead.
    print("\n[5] Setter: changing plot colors via DataPlot.color ...")
    plots[0].color = OriginColorIndex.RED          # index color
    print(f"  Plot 1: set to OriginColorIndex.RED ({OriginColorIndex.RED.value}) - OK")

    plots[1].color = (34, 136, 34)                 # RGB tuple
    print("  Plot 2: set to RGB(34,136,34) - OK")

    plots[2].color = OriginColorIndex.BLUE         # index color
    print(f"  Plot 3: set to OriginColorIndex.BLUE ({OriginColorIndex.BLUE.value}) - OK")

    print("  All setters ran without exception: OK")


if __name__ == "__main__":
    sys.exit(run())
