"""
Integration test: verify Legend manipulation via GraphLayer.get_legend().

Tests cover:
  [1]  get_legend() returns a Legend instance
  [2]  visible property: hide and show the legend
  [3]  font_size property: get and set
  [4]  font_size setter rejects non-positive values
  [5]  background property: get and set
  [6]  set_position() and get_position()
  [7]  reset_position() runs without error
  [8]  set_layout(HORIZONTAL) and set_layout(VERTICAL) run without error
  [9]  reconstruct() runs without error
  [10] update() runs without error
  [11] text property: get and set

Run from repo root:
    python test_codes/test_legend.py
"""
import sys
import os

REPO_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, REPO_ROOT)

import origin_pro_support as ops
from origin_pro_support.layer.enums import XYPlotType, LegendLayout
from origin_pro_support.layer.graph_layer import Legend


def run():
    proj_path = os.path.join(REPO_ROOT, "test_codes", "test_legend.opju")
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
    wbook = origin.new_workbook("LegendTestBook")
    assert wbook is not None
    ws = wbook.get_layer(0)
    assert ws is not None

    x_data = [1.0, 2.0, 3.0, 4.0, 5.0]
    y_data = [2.0, 4.0, 3.0, 5.0, 1.0]
    x_col = ws.add_column_from_data(x_data, lname="X")
    y_col = ws.add_column_from_data(y_data, lname="Y")

    graph_page = origin.new_graph("LegendTestGraph", XYPlotType.LINE_SYMBOL)
    assert graph_page is not None
    layer = graph_page.get_layer(0)
    assert layer is not None

    layer.add_xy_plot(ws, x_col.index, y_col.index, XYPlotType.LINE_SYMBOL)

    # ── [1] get_legend() returns Legend ──────────────────────────────────
    print("\n[1] get_legend() returns a Legend instance ...")
    leg = layer.get_legend()
    assert isinstance(leg, Legend), f"Expected Legend, got {type(leg)}"
    print("  OK")

    # ── [2] visible property ─────────────────────────────────────────────
    print("\n[2] visible: hide and show ...")
    leg.visible = False
    assert leg.visible is False, "Expected visible=False after hide"
    leg.visible = True
    assert leg.visible is True, "Expected visible=True after show"
    print("  OK")

    # ── [3] font_size get/set ────────────────────────────────────────────
    print("\n[3] font_size get/set ...")
    leg.font_size = 16
    fs = leg.font_size
    print(f"  font_size = {fs}")
    assert fs == 16, f"Expected 16, got {fs}"
    print("  OK")

    # ── [4] font_size rejects non-positive ───────────────────────────────
    print("\n[4] font_size rejects non-positive value ...")
    try:
        leg.font_size = 0
        assert False, "Expected ValueError was NOT raised"
    except ValueError as e:
        print(f"  ValueError raised as expected: {e}")
    print("  OK")

    # ── [5] background get/set ───────────────────────────────────────────
    print("\n[5] background get/set ...")
    leg.background = 0
    bg = leg.background
    print(f"  background = {bg}")
    assert bg == 0, f"Expected 0, got {bg}"
    leg.background = 1
    bg = leg.background
    assert bg == 1, f"Expected 1, got {bg}"
    print("  OK")

    # ── [6] set_position / get_position ─────────────────────────────────
    print("\n[6] set_position / get_position ...")
    leg.set_position(1.0, 4.5)
    pos = leg.get_position()
    print(f"  position = {pos}")
    assert len(pos) == 2, f"Expected 2-tuple, got {pos}"
    print("  OK")

    # ── [7] reset_position ───────────────────────────────────────────────
    print("\n[7] reset_position() ...")
    leg.reset_position()
    print("  OK (no exception)")

    # ── [8] set_layout ───────────────────────────────────────────────────
    print("\n[8] set_layout(HORIZONTAL) / set_layout(VERTICAL) ...")
    leg.set_layout(LegendLayout.HORIZONTAL)
    leg.set_layout(LegendLayout.VERTICAL)
    print("  OK (no exception)")

    # ── [9] reconstruct ──────────────────────────────────────────────────
    print("\n[9] reconstruct() ...")
    leg.reconstruct()
    print("  OK (no exception)")

    # ── [10] update ──────────────────────────────────────────────────────
    print("\n[10] update() ...")
    leg.update()
    print("  OK (no exception)")

    # ── [11] text get/set ────────────────────────────────────────────────
    print("\n[11] text get/set ...")
    original_text = leg.text
    print(f"  original text = '{original_text}'")
    leg.text = "Custom Legend"
    new_text = leg.text
    print(f"  new text = '{new_text}'")
    assert "Custom Legend" in new_text, f"Expected 'Custom Legend' in text, got '{new_text}'"
    print("  OK")


if __name__ == "__main__":
    sys.exit(run())
