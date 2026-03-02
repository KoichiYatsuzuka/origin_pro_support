"""
Unit tests for color-related functionality.

These tests verify the pure-Python logic (OriginColorIndex, color_to_lt_str)
and do NOT require Origin to be running.
"""
import sys
import os
import importlib.util

# Load layer/enums.py directly to avoid relative-import chain (base.py → OriginExt)
_enums_path = os.path.join(os.path.dirname(__file__), "..", "layer", "enums.py")
_spec = importlib.util.spec_from_file_location("layer.enums", _enums_path)
_mod = importlib.util.module_from_spec(_spec)
sys.modules["layer.enums"] = _mod  # must register before exec so @dataclass can resolve the module
_spec.loader.exec_module(_mod)
OriginColorIndex = _mod.OriginColorIndex
color_to_lt_str = _mod.color_to_lt_str


# ─── color_to_lt_str ────────────────────────────────────────────────────────

def test_int_index():
    assert color_to_lt_str(1) == "1"
    assert color_to_lt_str(24) == "24"


def test_named_constant():
    assert color_to_lt_str(OriginColorIndex.BLACK) == "1"
    assert color_to_lt_str(OriginColorIndex.RED) == "2"
    assert color_to_lt_str(OriginColorIndex.BLUE) == "4"
    assert color_to_lt_str(OriginColorIndex.DARK_GRAY) == "24"


def test_rgb_tuple():
    assert color_to_lt_str((240, 208, 0)) == "color(240,208,0)"
    assert color_to_lt_str((0, 0, 0)) == "color(0,0,0)"
    assert color_to_lt_str((255, 255, 255)) == "color(255,255,255)"


# ─── OriginColorIndex ────────────────────────────────────────────────────────

def test_all_indices_are_unique():
    values = [c.value for c in OriginColorIndex]
    assert len(values) == len(set(values)), "Duplicate color index detected"


def test_index_range():
    for c in OriginColorIndex:
        assert 1 <= c.value <= 24, f"{c.name} has out-of-range index {c.value}"


# ─── runner ──────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    tests = [
        test_int_index,
        test_named_constant,
        test_rgb_tuple,
        test_all_indices_are_unique,
        test_index_range,
    ]
    passed = 0
    failed = 0
    for t in tests:
        try:
            t()
            print(f"  PASS  {t.__name__}")
            passed += 1
        except Exception as e:
            print(f"  FAIL  {t.__name__}: {e}")
            failed += 1
    print(f"\n{passed} passed, {failed} failed")
    sys.exit(failed)
