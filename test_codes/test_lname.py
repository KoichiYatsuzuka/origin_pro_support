"""
Integration test: verify add_column_from_data long name behaviour.

Tests cover every data-type path and lname scenario:
  [1]  1-D list  – explicit lname
  [2]  1-D list  – lname=None  → auto list_N
  [3]  1-D list  – second auto call → list_{N+1} (continues from existing)
  [4]  1-D ndarray – explicit lname
  [5]  1-D ndarray – lname=None → auto (continues)
  [6]  2-D list  – explicit list[str] lname
  [7]  2-D list  – lname=None → auto list_N, list_{N+1}
  [8]  2-D list  – lname length mismatch → ValueError
  [9]  2-D ndarray – explicit list[str] lname
  [10] 2-D ndarray – lname=None → auto
  [11] pd.Series  – named series → series name used; lname arg ignored (warn)
  [12] pd.Series  – unnamed series → auto list_N
  [13] pd.DataFrame – column names used; lname arg ignored (warn)

Run from repo root:
    python test_codes/test_lname.py
"""
import sys
import os
import io

REPO_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, REPO_ROOT)

import numpy as np
import pandas as pd

import origin_pro_support as ops
from origin_pro_support.layer.enums import XYPlotType


# ── helpers ──────────────────────────────────────────────────────────────────

def get_lname(col) -> str:
    """Read the long name that Origin actually stored for a column."""
    return col.long_name


def run():
    proj_path = os.path.join(REPO_ROOT, "test_codes", "test_lname.opju")
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
    # cleanup leftovers
    for name in ("LNameTestBook",):
        b = origin.find_book(name)
        if b is not None:
            b.destroy()
            print(f"  Removed leftover '{name}'")

    wbook = origin.new_workbook("LNameTestBook")
    assert wbook is not None
    ws = wbook.get_layer(0)
    assert ws is not None

    data1d = [1.0, 2.0, 3.0]
    data1d_np = np.array([4.0, 5.0, 6.0])
    data2d = [[1, 2], [3, 4], [5, 6]]          # 3 rows × 2 cols
    data2d_np = np.array([[7, 8], [9, 10], [11, 12]])

    # ── [1] 1-D list with explicit lname ─────────────────────────────────
    print("\n[1] 1-D list, lname='my_x' ...")
    col = ws.add_column_from_data(data1d, lname="my_x")
    lname = get_lname(col)
    print(f"  long_name = '{lname}'")
    assert lname == "my_x", f"Expected 'my_x', got '{lname}'"
    print("  OK")

    # ── [2] 1-D list, lname=None → auto list_1 ──────────────────────────
    print("\n[2] 1-D list, lname=None (auto) ...")
    col2 = ws.add_column_from_data(data1d)
    lname2 = get_lname(col2)
    print(f"  long_name = '{lname2}'")
    assert lname2.startswith("list_"), f"Expected 'list_N', got '{lname2}'"
    n2 = int(lname2.split("_")[1])
    print(f"  Auto number: {n2} OK")

    # ── [3] second 1-D list, lname=None → list_{n2+1} ───────────────────
    print("\n[3] 1-D list, lname=None (2nd auto, should be list_{n2+1}) ...")
    col3 = ws.add_column_from_data(data1d)
    lname3 = get_lname(col3)
    print(f"  long_name = '{lname3}'")
    assert lname3 == f"list_{n2 + 1}", f"Expected 'list_{n2+1}', got '{lname3}'"
    print("  OK")

    # ── [4] 1-D ndarray, explicit lname ─────────────────────────────────
    print("\n[4] 1-D ndarray, lname='arr_col' ...")
    col4 = ws.add_column_from_data(data1d_np, lname="arr_col")
    lname4 = get_lname(col4)
    print(f"  long_name = '{lname4}'")
    assert lname4 == "arr_col", f"Expected 'arr_col', got '{lname4}'"
    print("  OK")

    # ── [5] 1-D ndarray, lname=None → continues numbering ───────────────
    print("\n[5] 1-D ndarray, lname=None (auto, continues) ...")
    col5 = ws.add_column_from_data(data1d_np)
    lname5 = get_lname(col5)
    print(f"  long_name = '{lname5}'")
    assert lname5.startswith("list_"), f"Expected 'list_N', got '{lname5}'"
    n5 = int(lname5.split("_")[1])
    assert n5 == n2 + 2, f"Expected list_{n2+2}, got '{lname5}'"
    print(f"  Auto number: {n5} OK")

    # ── [6] 2-D list, explicit list[str] lname ───────────────────────────
    print("\n[6] 2-D list, lname=['colA', 'colB'] ...")
    cols6 = ws.add_column_from_data(data2d, lname=["colA", "colB"])
    assert isinstance(cols6, list) and len(cols6) == 2
    ln6 = [get_lname(c) for c in cols6]
    print(f"  long_names = {ln6}")
    assert ln6 == ["colA", "colB"], f"Expected ['colA','colB'], got {ln6}"
    print("  OK")

    # ── [7] 2-D list, lname=None → auto list_N, list_{N+1} ──────────────
    print("\n[7] 2-D list, lname=None (auto) ...")
    cols7 = ws.add_column_from_data(data2d)
    assert isinstance(cols7, list) and len(cols7) == 2
    ln7 = [get_lname(c) for c in cols7]
    print(f"  long_names = {ln7}")
    assert ln7[0].startswith("list_"), f"Expected list_N, got '{ln7[0]}'"
    n7 = int(ln7[0].split("_")[1])
    assert ln7[1] == f"list_{n7 + 1}", f"Expected 'list_{n7+1}', got '{ln7[1]}'"
    print(f"  Auto numbers: {n7}, {n7+1} OK")

    # ── [8] 2-D list, lname length mismatch → ValueError ─────────────────
    print("\n[8] 2-D list, mismatched lname (expect ValueError) ...")
    try:
        ws.add_column_from_data(data2d, lname=["only_one"])
        assert False, "Expected ValueError was NOT raised"
    except ValueError as e:
        print(f"  ValueError raised as expected: {e}")
        print("  OK")

    # ── [9] 2-D ndarray, explicit list[str] lname ────────────────────────
    print("\n[9] 2-D ndarray, lname=['npA', 'npB'] ...")
    cols9 = ws.add_column_from_data(data2d_np, lname=["npA", "npB"])
    assert isinstance(cols9, list) and len(cols9) == 2
    ln9 = [get_lname(c) for c in cols9]
    print(f"  long_names = {ln9}")
    assert ln9 == ["npA", "npB"], f"Expected ['npA','npB'], got {ln9}"
    print("  OK")

    # ── [10] 2-D ndarray, lname=None → auto ─────────────────────────────
    print("\n[10] 2-D ndarray, lname=None (auto) ...")
    cols10 = ws.add_column_from_data(data2d_np)
    assert isinstance(cols10, list) and len(cols10) == 2
    ln10 = [get_lname(c) for c in cols10]
    print(f"  long_names = {ln10}")
    assert ln10[0].startswith("list_"), f"Expected list_N, got '{ln10[0]}'"
    n10 = int(ln10[0].split("_")[1])
    assert ln10[1] == f"list_{n10 + 1}", f"Expected list_{n10+1}, got '{ln10[1]}'"
    print(f"  Auto numbers: {n10}, {n10+1} OK")

    # ── [11] pd.Series with name → series name used; lname arg ignored ───
    print("\n[11] pd.Series (name='series_col'), lname='ignored' (expect warning) ...")
    s = pd.Series([1.0, 2.0, 3.0], name="series_col")
    import io as _io
    buf = _io.StringIO()
    old_stdout = sys.stdout
    sys.stdout = buf
    col11 = ws.add_column_from_data(s, lname="ignored")
    sys.stdout = old_stdout
    warn_out = buf.getvalue()
    print(f"  Warning output: '{warn_out.strip()}'")
    assert "Warning" in warn_out, "Expected warning for ignored lname on pd.Series"
    lname11 = get_lname(col11)
    print(f"  long_name = '{lname11}'")
    assert lname11 == "series_col", f"Expected 'series_col', got '{lname11}'"
    print("  OK")

    # ── [12] pd.Series without name → auto list_N ────────────────────────
    print("\n[12] pd.Series (no name), lname=None (auto) ...")
    s_unnamed = pd.Series([7.0, 8.0, 9.0])  # name=None
    col12 = ws.add_column_from_data(s_unnamed)
    lname12 = get_lname(col12)
    print(f"  long_name = '{lname12}'")
    assert lname12.startswith("list_"), f"Expected 'list_N', got '{lname12}'"
    print("  OK")

    # ── [13] pd.DataFrame → column names used; lname arg ignored (warn) ──
    print("\n[13] pd.DataFrame (cols 'alpha','beta'), lname=['x','y'] (expect warning) ...")
    df = pd.DataFrame({"alpha": [1, 2, 3], "beta": [4, 5, 6]})
    buf2 = _io.StringIO()
    sys.stdout = buf2
    cols13 = ws.add_column_from_data(df, lname=["x", "y"])
    sys.stdout = old_stdout
    warn_out2 = buf2.getvalue()
    print(f"  Warning output: '{warn_out2.strip()}'")
    assert "Warning" in warn_out2, "Expected warning for ignored lname on pd.DataFrame"
    assert isinstance(cols13, list) and len(cols13) == 2
    ln13 = [get_lname(c) for c in cols13]
    print(f"  long_names = {ln13}")
    assert ln13 == ["alpha", "beta"], f"Expected ['alpha','beta'], got {ln13}"
    print("  OK")


if __name__ == "__main__":
    sys.exit(run())
