"""
Integration test: set and read back GraphLayer drawing scale (position + size).

Tests verified:
  1. set_scale() + get_scale() round-trip matches within tolerance
  2. Multiple different scale configurations are applied correctly
  3. Default scale (before any set_scale call) can be read without error

Run from repo root:
    python test_codes/test_scale.py
"""
import sys
import os

# ── path setup ──────────────────────────────────────────────────────────────
REPO_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, REPO_ROOT)

import origin_pro_support as ops
from origin_pro_support.layer.enums import XYPlotType

# ── tolerance for float comparison (Origin returns % values) ─────────────────
TOLERANCE = 0.5   # accept up to 0.5 percentage-point error


def approx_eq(a: float, b: float, tol: float = TOLERANCE) -> bool:
    return abs(a - b) <= tol


# ── test ─────────────────────────────────────────────────────────────────────

def run():
    proj_path = os.path.join(REPO_ROOT, "test_codes", "test_scale.opju")

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
    # ── cleanup leftovers from previous runs ──────────────────────────────
    existing = origin.find_graph("ScaleTestGraph")
    if existing is not None:
        existing.destroy()
        print("  Removed leftover ScaleTestGraph")

    # ── 1. Create a blank graph ───────────────────────────────────────────
    print("\n[1] Creating graph...")
    graph = origin.new_graph("ScaleTestGraph", XYPlotType.LINE)
    assert graph is not None, "Failed to create graph"

    layer = graph.get_layer(0)
    assert layer is not None, "Failed to get layer 0"
    print(f"  Graph '{graph.name}' created with layer id={layer.id}")

    # ── 2. Read default scale (must not raise) ────────────────────────────
    print("\n[2] Reading default scale...")
    default_scale = layer.get_scale()
    assert len(default_scale) == 4, "get_scale() must return a 4-tuple"
    x0, y0, w0, h0 = default_scale
    print(f"  Default scale: x={x0:.2f}%, y={y0:.2f}%, w={w0:.2f}%, h={h0:.2f}%")

    # ── 3. Set scale A and verify round-trip ──────────────────────────────
    print("\n[3] Setting scale A: x=10, y=10, w=60, h=55 ...")
    layer.set_scale(10, 10, 60, 55)
    xa, ya, wa, ha = layer.get_scale()
    print(f"  Read back: x={xa:.2f}%, y={ya:.2f}%, w={wa:.2f}%, h={ha:.2f}%")

    assert approx_eq(xa, 10), f"x mismatch: expected ~10, got {xa:.2f}"
    assert approx_eq(ya, 10), f"y mismatch: expected ~10, got {ya:.2f}"
    assert approx_eq(wa, 60), f"width mismatch: expected ~60, got {wa:.2f}"
    assert approx_eq(ha, 55), f"height mismatch: expected ~55, got {ha:.2f}"
    print("  Scale A verified OK")

    # ── 4. Set scale B (different values) and verify ──────────────────────
    print("\n[4] Setting scale B: x=20, y=30, w=50, h=40 ...")
    layer.set_scale(20, 30, 50, 40)
    xb, yb, wb, hb = layer.get_scale()
    print(f"  Read back: x={xb:.2f}%, y={yb:.2f}%, w={wb:.2f}%, h={hb:.2f}%")

    assert approx_eq(xb, 20), f"x mismatch: expected ~20, got {xb:.2f}"
    assert approx_eq(yb, 30), f"y mismatch: expected ~30, got {yb:.2f}"
    assert approx_eq(wb, 50), f"width mismatch: expected ~50, got {wb:.2f}"
    assert approx_eq(hb, 40), f"height mismatch: expected ~40, got {hb:.2f}"
    print("  Scale B verified OK")

    # ── 5. Verify that A and B are different (scale actually changed) ──────
    print("\n[5] Verifying scale changed between A and B ...")
    assert not approx_eq(xa, xb, tol=1.0), "x did not change between A and B"
    assert not approx_eq(ya, yb, tol=1.0), "y did not change between A and B"
    assert not approx_eq(wa, wb, tol=1.0), "width did not change between A and B"
    assert not approx_eq(ha, hb, tol=1.0), "height did not change between A and B"
    print("  Scale change confirmed OK")

    # ── 6. Restore to default-like scale ──────────────────────────────────
    print("\n[6] Restoring scale to default-like (x=15, y=15, w=70, h=70) ...")
    layer.set_scale(15, 15, 70, 70)
    xr, yr, wr, hr = layer.get_scale()
    print(f"  Read back: x={xr:.2f}%, y={yr:.2f}%, w={wr:.2f}%, h={hr:.2f}%")
    assert approx_eq(xr, 15), f"x mismatch: expected ~15, got {xr:.2f}"
    assert approx_eq(yr, 15), f"y mismatch: expected ~15, got {yr:.2f}"
    assert approx_eq(wr, 70), f"width mismatch: expected ~70, got {wr:.2f}"
    assert approx_eq(hr, 70), f"height mismatch: expected ~70, got {hr:.2f}"
    print("  Restore verified OK")


if __name__ == "__main__":
    sys.exit(run())
