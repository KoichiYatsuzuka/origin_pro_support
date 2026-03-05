"""
Verification test for sample_UV-vis.py

Runs the sample to generate sample_UV-vis.opju, then reopens the file and
reads back all settings that the sample was supposed to apply, asserting they
match the expected values.

Run from the repository root:
    python sample/test_sample_UV-vis.py
"""
import sys
import os
import subprocess
import math
import traceback

REPO_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, REPO_ROOT)

import origin_pro_support as ops
from origin_pro_support.layer.enums import (
    XYPlotType,
    AxisType,
    OriginColorIndex,
    LineStyle,
)

SAMPLE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJ_PATH  = os.path.join(SAMPLE_DIR, "sample_UV-vis.opju")

# Page size tolerances (inches)
# NOTE: When an opju is reopened, Origin may apply a printer-resolution
# correction to page.width (typically +5-6%).  We verify that the page is
# close to the target size; an absolute tolerance of 0.5 inch covers that.
PAGE_WIDTH_EXPECTED  = 8.0   # inches
PAGE_HEIGHT_EXPECTED = 6.0   # inches
PAGE_TOL = 0.5   # inch tolerance (allows for printer-size correction)

# Layer geometry tolerances (%)
LAYER_LEFT_EXPECTED   = 15.0
LAYER_TOP_EXPECTED    =  5.0
LAYER_WIDTH_EXPECTED  = 80.0
LAYER_HEIGHT_EXPECTED = 80.0
SCALE_TOL = 0.5

# Origin RGB color encoding: 0x1000000 | (R + G*256 + B*65536)
_COLOR_BB33BB = 0x1000000 + 187 + 51 * 256 + 187 * 65536


def _approx(a: float, b: float, tol: float) -> bool:
    return abs(a - b) <= tol


def run() -> int:
    print("=== Step 1: generate sample_UV-vis.opju via sample ===")
    sample_script = os.path.join(SAMPLE_DIR, "sample_UV-vis.py")
    result = subprocess.run(
        [sys.executable, sample_script],
        capture_output=True, text=True
    )
    print(result.stdout)
    if result.returncode != 0:
        print(result.stderr)
        print("=== FAILED: sample script exited with non-zero code ===")
        return 1

    if not os.path.exists(PROJ_PATH):
        print(f"=== FAILED: {PROJ_PATH} was not created ===")
        return 1

    print("\n=== Step 2: open opju and verify settings ===")
    origin = ops.OriginInstance(PROJ_PATH)
    try:
        _verify(origin)
        print("\n=== ALL CHECKS PASSED ===")
        return 0
    except AssertionError as e:
        print(f"\n=== FAILED (assertion): {e} ===")
        traceback.print_exc()
        return 1
    except Exception as e:
        print(f"\n=== FAILED (exception): {e} ===")
        traceback.print_exc()
        return 1
    finally:
        origin.save()
        origin.close()
        print("Origin closed.")


def _verify(origin: ops.OriginInstance) -> None:
    _verify_graph1(origin)
    _verify_graph2(origin)


# ---------------------------------------------------------------------------
# Graph 1 verification
# ---------------------------------------------------------------------------

def _verify_graph1(origin: ops.OriginInstance) -> None:
    print("\n--- Graph 1: UVvis_Spectra_Graph ---")

    gpage = origin.find_graph("UVvis_Spectra_Graph")
    assert gpage is not None, "Graph 'UVvis_Spectra_Graph' not found"

    # ── page size ────────────────────────────────────────────────────────
    print("[1] Page size ...")
    w = gpage.get_width()
    h = gpage.get_height()
    print(f"  width={w}, expected~{PAGE_WIDTH_EXPECTED}")
    print(f"  height={h}, expected~{PAGE_HEIGHT_EXPECTED}")
    assert _approx(w, PAGE_WIDTH_EXPECTED, PAGE_TOL), \
        f"Page width mismatch: got {w}, expected ~{PAGE_WIDTH_EXPECTED}"
    assert _approx(h, PAGE_HEIGHT_EXPECTED, PAGE_TOL), \
        f"Page height mismatch: got {h}, expected ~{PAGE_HEIGHT_EXPECTED}"
    print("  OK")

    layer = gpage.get_layer(0)
    assert layer is not None, "Layer 0 not found in UVvis_Spectra_Graph"

    # ── layer geometry ───────────────────────────────────────────────────
    print("[2] Layer geometry ...")
    x, y, lw, lh = layer.get_scale()
    print(f"  left={x:.2f}% top={y:.2f}% width={lw:.2f}% height={lh:.2f}%")
    assert _approx(x,  LAYER_LEFT_EXPECTED,   SCALE_TOL), f"layer.left mismatch: {x}"
    assert _approx(y,  LAYER_TOP_EXPECTED,    SCALE_TOL), f"layer.top mismatch: {y}"
    assert _approx(lw, LAYER_WIDTH_EXPECTED,  SCALE_TOL), f"layer.width mismatch: {lw}"
    assert _approx(lh, LAYER_HEIGHT_EXPECTED, SCALE_TOL), f"layer.height mismatch: {lh}"
    print("  OK")

    # ── plots ────────────────────────────────────────────────────────────
    plots = layer.data_plots
    assert len(plots) == 3, f"Expected 3 plots, got {len(plots)}"
    plot_0, plot_50, plot_100 = plots[0], plots[1], plots[2]

    # ── line styles ──────────────────────────────────────────────────────
    print("[3] Line styles ...")
    s0   = plot_0.line_style
    s50  = plot_50.line_style
    s100 = plot_100.line_style
    print(f"  plot_0={s0.name}, plot_50={s50.name}, plot_100={s100.name}")
    assert s0   == LineStyle.DOT,   f"plot_0 line_style: expected DOT, got {s0.name}"
    assert s50  == LineStyle.SOLID, f"plot_50 line_style: expected SOLID, got {s50.name}"
    assert s100 == LineStyle.SOLID, f"plot_100 line_style: expected SOLID, got {s100.name}"
    print("  OK")

    # ── line widths ──────────────────────────────────────────────────────
    print("[4] Line widths ...")
    w0   = plot_0.line_width
    w50  = plot_50.line_width
    w100 = plot_100.line_width
    print(f"  plot_0={w0}, plot_50={w50}, plot_100={w100}")
    assert w0   == 2.0, f"plot_0 line_width: expected 2.0, got {w0}"
    assert w50  == 2.0, f"plot_50 line_width: expected 2.0, got {w50}"
    assert w100 == 2.0, f"plot_100 line_width: expected 2.0, got {w100}"
    print("  OK")

    # ── colors ───────────────────────────────────────────────────────────
    print("[5] Plot colors ...")
    c0   = plot_0.color
    c50  = plot_50.color
    c100 = plot_100.color
    print(f"  plot_0 color={c0} (expected {OriginColorIndex.BLACK.value})")
    print(f"  plot_50 color={c50} (expected {OriginColorIndex.RED.value})")
    print(f"  plot_100 color={c100} (expected {OriginColorIndex.BLUE.value})")
    assert c0   == OriginColorIndex.BLACK.value, \
        f"plot_0 color: expected {OriginColorIndex.BLACK.value} (BLACK), got {c0}"
    assert c50  == OriginColorIndex.RED.value, \
        f"plot_50 color: expected {OriginColorIndex.RED.value} (RED), got {c50}"
    assert c100 == OriginColorIndex.BLUE.value, \
        f"plot_100 color: expected {OriginColorIndex.BLUE.value} (BLUE), got {c100}"
    print("  OK")

    # ── axis labels ──────────────────────────────────────────────────────
    print("[6] Axis labels ...")
    ax_x = layer.get_axis(AxisType.X)
    ax_y = layer.get_axis(AxisType.Y)
    lbl_x = ax_x.label_text
    lbl_y = ax_y.label_text
    print(f"  X='{lbl_x}', Y='{lbl_y}'")
    assert lbl_x == "Wavelength [nm]", f"X label: expected 'Wavelength [nm]', got '{lbl_x}'"
    assert lbl_y == "Absorbance [-]",  f"Y label: expected 'Absorbance [-]', got '{lbl_y}'"
    print("  OK")

    # ── axis tick spacing ────────────────────────────────────────────────
    print("[7] X major tick spacing ...")
    spacing = ax_x.get_major_tick_spacing()
    print(f"  spacing={spacing} (expected 200)")
    assert spacing == 200.0, f"X major tick spacing: expected 200.0, got {spacing}"
    print("  OK")

    print("[8] X minor tick count ...")
    minor = ax_x.get_minor_tick_count()
    print(f"  minor_count={minor} (expected 1)")
    assert minor == 1, f"X minor tick count: expected 1, got {minor}"
    print("  OK")


# ---------------------------------------------------------------------------
# Graph 2 verification
# ---------------------------------------------------------------------------

def _verify_graph2(origin: ops.OriginInstance) -> None:
    print("\n--- Graph 2: UVvis_600nm_Graph ---")

    gpage = origin.find_graph("UVvis_600nm_Graph")
    assert gpage is not None, "Graph 'UVvis_600nm_Graph' not found"

    # ── page size ────────────────────────────────────────────────────────
    print("[1] Page size ...")
    w = gpage.get_width()
    h = gpage.get_height()
    print(f"  width={w}, expected~{PAGE_WIDTH_EXPECTED}")
    print(f"  height={h}, expected~{PAGE_HEIGHT_EXPECTED}")
    assert _approx(w, PAGE_WIDTH_EXPECTED, PAGE_TOL), \
        f"Page width mismatch: got {w}, expected ~{PAGE_WIDTH_EXPECTED}"
    assert _approx(h, PAGE_HEIGHT_EXPECTED, PAGE_TOL), \
        f"Page height mismatch: got {h}, expected ~{PAGE_HEIGHT_EXPECTED}"
    print("  OK")

    layer = gpage.get_layer(0)
    assert layer is not None, "Layer 0 not found in UVvis_600nm_Graph"

    # ── layer geometry ───────────────────────────────────────────────────
    print("[2] Layer geometry ...")
    x, y, lw, lh = layer.get_scale()
    print(f"  left={x:.2f}% top={y:.2f}% width={lw:.2f}% height={lh:.2f}%")
    assert _approx(x,  LAYER_LEFT_EXPECTED,   SCALE_TOL), f"layer.left mismatch: {x}"
    assert _approx(y,  LAYER_TOP_EXPECTED,    SCALE_TOL), f"layer.top mismatch: {y}"
    assert _approx(lw, LAYER_WIDTH_EXPECTED,  SCALE_TOL), f"layer.width mismatch: {lw}"
    assert _approx(lh, LAYER_HEIGHT_EXPECTED, SCALE_TOL), f"layer.height mismatch: {lh}"
    print("  OK")

    plots = layer.data_plots
    assert len(plots) == 1, f"Expected 1 plot, got {len(plots)}"
    plot = plots[0]

    # ── color ────────────────────────────────────────────────────────────
    print("[3] Plot color (#BB33BB) ...")
    c = plot.color
    print(f"  color={c} (expected {_COLOR_BB33BB})")
    assert c == _COLOR_BB33BB, \
        f"plot color: expected {_COLOR_BB33BB} (RGB 187,51,187), got {c}"
    print("  OK")

    # ── symbol size ──────────────────────────────────────────────────────
    print("[4] Symbol size ...")
    sz = plot.symbol_size
    print(f"  symbol_size={sz} (expected 12.0)")
    assert sz == 12.0, f"symbol_size: expected 12.0, got {sz}"
    print("  OK")

    # ── line width ───────────────────────────────────────────────────────
    print("[5] Line width ...")
    lwidth = plot.line_width
    print(f"  line_width={lwidth} (expected 2.0)")
    assert lwidth == 2.0, f"line_width: expected 2.0, got {lwidth}"
    print("  OK")

    # ── axis labels ──────────────────────────────────────────────────────
    print("[6] Axis labels ...")
    ax_x = layer.get_axis(AxisType.X)
    ax_y = layer.get_axis(AxisType.Y)
    lbl_x = ax_x.label_text
    lbl_y = ax_y.label_text
    print(f"  X='{lbl_x}', Y='{lbl_y}'")
    assert lbl_x == "Time [min]",     f"X label: expected 'Time [min]', got '{lbl_x}'"
    assert lbl_y == "Absorbance [-]", f"Y label: expected 'Absorbance [-]', got '{lbl_y}'"
    print("  OK")

    # ── axis tick spacing ────────────────────────────────────────────────
    print("[7] X major tick spacing ...")
    spacing = ax_x.get_major_tick_spacing()
    print(f"  spacing={spacing} (expected 100)")
    assert spacing == 100.0, f"X major tick spacing: expected 100.0, got {spacing}"
    print("  OK")

    print("[8] X minor tick count ...")
    minor = ax_x.get_minor_tick_count()
    print(f"  minor_count={minor} (expected 1)")
    assert minor == 1, f"X minor tick count: expected 1, got {minor}"
    print("  OK")


if __name__ == "__main__":
    sys.exit(run())
