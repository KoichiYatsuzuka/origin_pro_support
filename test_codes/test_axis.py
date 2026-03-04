"""
Integration test: verify Axis manipulation via GraphLayer.get_axis().

Tests cover:
  [1]  get_axis() returns an Axis instance for X and Y
  [2]  get_range() / set_range(): set and verify roundtrip
  [3]  get_reverse() / set_reverse(): set True then False and verify
  [4]  get_major_tick() / set_major_tick(): set each TickType and verify
  [5]  get_minor_tick() / set_minor_tick(): set each TickType and verify
  [6]  show_opposite_axis() / hide_opposite_axis() / get_opposite_axis_visible() for X
  [7]  show_opposite_axis() / hide_opposite_axis() / get_opposite_axis_visible() for Y
  [8]  show_opposite_axis() raises ValueError for Z axis
  [9]  label_text property: get and set (X and Y)
  [10] show_label() / hide_label() / get_label_visible() for X axis
  [11] show_label() / hide_label() / get_label_visible() for Y axis
  [12] get_major_tick_spacing() / set_major_tick_spacing(): int and float roundtrip, reject non-positive
  [13] get_minor_tick_count() / set_minor_tick_count(): int roundtrip, reject negative

Run from repo root:
    python test_codes/test_axis.py
"""
import sys
import os

REPO_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, REPO_ROOT)

import origin_pro_support as ops
from origin_pro_support.layer.enums import XYPlotType, AxisType, TickType
from origin_pro_support.layer.graph_layer import Axis


def run():
    proj_path = os.path.join(REPO_ROOT, "test_codes", "test_axis.opju")
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
    wbook = origin.new_workbook("AxisTestBook")
    assert wbook is not None
    ws = wbook.get_layer(0)
    assert ws is not None

    x_data = [1.0, 2.0, 3.0, 4.0, 5.0]
    y_data = [2.0, 4.0, 3.0, 5.0, 1.0]
    x_col = ws.add_column_from_data(x_data, lname="X")
    y_col = ws.add_column_from_data(y_data, lname="Y")

    graph_page = origin.new_graph("AxisTestGraph", XYPlotType.LINE_SYMBOL)
    assert graph_page is not None
    layer = graph_page.get_layer(0)
    assert layer is not None

    layer.add_xy_plot(ws, x_col.index, y_col.index, XYPlotType.LINE_SYMBOL)

    ax_x = layer.get_axis(AxisType.X)
    ax_y = layer.get_axis(AxisType.Y)

    # ── [1] get_axis() returns Axis ───────────────────────────────────────
    print("\n[1] get_axis() returns Axis instances ...")
    assert isinstance(ax_x, Axis), f"Expected Axis, got {type(ax_x)}"
    assert isinstance(ax_y, Axis), f"Expected Axis, got {type(ax_y)}"
    print("  OK")

    # ── [2] get_range / set_range ─────────────────────────────────────────
    print("\n[2] get_range() / set_range() ...")
    ax_x.set_range(0.0, 10.0)
    rng = ax_x.get_range()
    print(f"  X range = {rng}")
    assert rng[0] == 0.0, f"Expected min=0.0, got {rng[0]}"
    assert rng[1] == 10.0, f"Expected max=10.0, got {rng[1]}"

    ax_y.set_range(-5.0, 5.0)
    rng_y = ax_y.get_range()
    print(f"  Y range = {rng_y}")
    assert rng_y[0] == -5.0, f"Expected min=-5.0, got {rng_y[0]}"
    assert rng_y[1] == 5.0, f"Expected max=5.0, got {rng_y[1]}"
    print("  OK")

    # ── [3] get_reverse / set_reverse ─────────────────────────────────────
    print("\n[3] get_reverse() / set_reverse() ...")
    ax_x.set_reverse(True)
    rev = ax_x.get_reverse()
    print(f"  reverse after set True = {rev}")
    assert rev is True, f"Expected True, got {rev}"

    ax_x.set_reverse(False)
    rev = ax_x.get_reverse()
    print(f"  reverse after set False = {rev}")
    assert rev is False, f"Expected False, got {rev}"
    print("  OK")

    # ── [4] get_major_tick / set_major_tick ───────────────────────────────
    print("\n[4] get_major_tick() / set_major_tick() ...")
    for tick in (TickType.NONE, TickType.IN, TickType.OUT, TickType.IN_OUT):
        ax_x.set_major_tick(tick)
        result = ax_x.get_major_tick()
        print(f"  set {tick.name} -> got {result.name}")
        assert result == tick, f"Expected {tick}, got {result}"
    print("  OK")

    # ── [5] get_minor_tick / set_minor_tick ───────────────────────────────
    print("\n[5] get_minor_tick() / set_minor_tick() ...")
    for tick in (TickType.NONE, TickType.IN, TickType.OUT, TickType.IN_OUT):
        ax_y.set_minor_tick(tick)
        result = ax_y.get_minor_tick()
        print(f"  set {tick.name} -> got {result.name}")
        assert result == tick, f"Expected {tick}, got {result}"
    print("  OK")

    # ── [6] opposite axis visibility: X -> x2 ────────────────────────────
    print("\n[6] show/hide opposite axis (X -> x2) ...")
    ax_x.show_opposite_axis()
    visible = ax_x.get_opposite_axis_visible()
    print(f"  visible after show = {visible}")
    assert visible is True, f"Expected True, got {visible}"

    ax_x.hide_opposite_axis()
    visible = ax_x.get_opposite_axis_visible()
    print(f"  visible after hide = {visible}")
    assert visible is False, f"Expected False, got {visible}"
    print("  OK")

    # ── [7] opposite axis visibility: Y -> y2 ────────────────────────────
    print("\n[7] show/hide opposite axis (Y -> y2) ...")
    ax_y.show_opposite_axis()
    visible = ax_y.get_opposite_axis_visible()
    print(f"  visible after show = {visible}")
    assert visible is True, f"Expected True, got {visible}"

    ax_y.hide_opposite_axis()
    visible = ax_y.get_opposite_axis_visible()
    print(f"  visible after hide = {visible}")
    assert visible is False, f"Expected False, got {visible}"
    print("  OK")

    # ── [8] opposite axis raises ValueError for Z ─────────────────────────
    print("\n[8] show_opposite_axis() raises ValueError for Z axis ...")
    ax_z = layer.get_axis(AxisType.Z)
    try:
        ax_z.show_opposite_axis()
        assert False, "Expected ValueError was NOT raised"
    except ValueError as e:
        print(f"  ValueError raised as expected: {e}")
    print("  OK")

    # ── [9] label_text get/set ────────────────────────────────────────────
    print("\n[9] label_text get/set ...")
    ax_x.label_text = "Time (s)"
    lbl = ax_x.label_text
    print(f"  X label_text = '{lbl}'")
    assert lbl == "Time (s)", f"Expected 'Time (s)', got '{lbl}'"

    ax_y.label_text = "Amplitude"
    lbl_y = ax_y.label_text
    print(f"  Y label_text = '{lbl_y}'")
    assert lbl_y == "Amplitude", f"Expected 'Amplitude', got '{lbl_y}'"
    print("  OK")

    # ── [10] show_label / hide_label / get_label_visible: X ───────────────
    print("\n[10] show_label() / hide_label() / get_label_visible() for X ...")
    ax_x.hide_label()
    vis = ax_x.get_label_visible()
    print(f"  label_visible after hide = {vis}")
    assert vis is False, f"Expected False, got {vis}"

    ax_x.show_label()
    vis = ax_x.get_label_visible()
    print(f"  label_visible after show = {vis}")
    assert vis is True, f"Expected True, got {vis}"
    print("  OK")

    # ── [11] show_label / hide_label / get_label_visible: Y ───────────────
    print("\n[11] show_label() / hide_label() / get_label_visible() for Y ...")
    ax_y.hide_label()
    vis = ax_y.get_label_visible()
    print(f"  label_visible after hide = {vis}")
    assert vis is False, f"Expected False, got {vis}"

    ax_y.show_label()
    vis = ax_y.get_label_visible()
    print(f"  label_visible after show = {vis}")
    assert vis is True, f"Expected True, got {vis}"
    print("  OK")

    # ── [12] get_major_tick_spacing / set_major_tick_spacing ──────────────
    print("\n[12] get_major_tick_spacing() / set_major_tick_spacing() ...")
    ax_x.set_range(0.0, 10.0)

    ax_x.set_major_tick_spacing(2)
    sp = ax_x.get_major_tick_spacing()
    print(f"  spacing after set 2 (int) = {sp}")
    assert sp == 2.0, f"Expected 2.0, got {sp}"

    ax_x.set_major_tick_spacing(0.5)
    sp = ax_x.get_major_tick_spacing()
    print(f"  spacing after set 0.5 (float) = {sp}")
    assert sp == 0.5, f"Expected 0.5, got {sp}"

    try:
        ax_x.set_major_tick_spacing(0)
        assert False, "Expected ValueError was NOT raised"
    except ValueError as e:
        print(f"  ValueError for 0 raised as expected: {e}")

    try:
        ax_x.set_major_tick_spacing(-1.0)
        assert False, "Expected ValueError was NOT raised"
    except ValueError as e:
        print(f"  ValueError for -1.0 raised as expected: {e}")
    print("  OK")

    # ── [13] get_minor_tick_count / set_minor_tick_count ──────────────────
    print("\n[13] get_minor_tick_count() / set_minor_tick_count() ...")
    ax_y.set_minor_tick_count(4)
    cnt = ax_y.get_minor_tick_count()
    print(f"  count after set 4 = {cnt}")
    assert cnt == 4, f"Expected 4, got {cnt}"

    ax_y.set_minor_tick_count(0)
    cnt = ax_y.get_minor_tick_count()
    print(f"  count after set 0 = {cnt}")
    assert cnt == 0, f"Expected 0, got {cnt}"

    try:
        ax_y.set_minor_tick_count(-1)
        assert False, "Expected ValueError was NOT raised"
    except ValueError as e:
        print(f"  ValueError for -1 raised as expected: {e}")
    print("  OK")


if __name__ == "__main__":
    sys.exit(run())
