"""
Probe: 同一データレンジを複数回プロットした場合の DataPlot 識別子の挙動調査

ケース:
  [A] 同一列を2回プロット → GetDatasetName / GetRange / GetName の各値
  [B] 異なる列を1回ずつプロット → 比較のための基準

Run from repo root:
    python test_codes/test_datasetname_dup.py
"""
import sys
import os

REPO_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, REPO_ROOT)

import origin_pro_support as ops
from origin_pro_support.layer.enums import XYPlotType

proj_path = os.path.join(REPO_ROOT, "test_codes", "test_datasetname_dup.opju")
origin = ops.OriginInstance(proj_path)

def dump_plots(raw_gl, label: str):
    dp_col = raw_gl.GetDataPlots()
    n = len(dp_col)
    print(f"\n  [{label}] プロット数: {n}")
    for i in range(n):
        dp = dp_col[i]
        if dp is None:
            print(f"    [{i}] None (範囲外)")
            break
        print(f"    [{i}] GetDatasetName() = {repr(dp.GetDatasetName())}")
        print(f"    [{i}] GetName()        = {repr(dp.GetName())}")
        print(f"    [{i}] GetRange()       = {repr(dp.GetRange())}")

try:
    # ── ワークブック作成 ──────────────────────────────────────────────────────
    wb = origin.new_workbook("DupBook")
    ws = wb.get_layer(0)
    ws.add_column_from_data([1.0, 2.0, 3.0, 4.0, 5.0], lname="X")
    ws.add_column_from_data([10.0, 20.0, 30.0, 40.0, 50.0], lname="Y1")
    ws.add_column_from_data([5.0, 15.0, 25.0, 35.0, 45.0], lname="Y2")

    # ── ケースA: 同一列(Y1)を2回プロット ──────────────────────────────────────
    print("\n=== ケースA: 同一列(Y1)を2回プロット ===")
    gp_a = origin.new_graph("DupGraphA", XYPlotType.SCATTER)
    gl_a = gp_a.get_layer(0)
    gl_a.add_xy_plot(ws, 0, 1, XYPlotType.SCATTER)   # Y1 (col index 1)
    gl_a.add_xy_plot(ws, 0, 1, XYPlotType.SCATTER)   # Y1 再度
    dump_plots(gl_a._obj, "Y1+Y1")

    # ── ケースB: 異なる列(Y1, Y2)を1回ずつプロット ───────────────────────────
    print("\n=== ケースB: 異なる列(Y1, Y2)を1回ずつプロット ===")
    gp_b = origin.new_graph("DupGraphB", XYPlotType.SCATTER)
    gl_b = gp_b.get_layer(0)
    gl_b.add_xy_plot(ws, 0, 1, XYPlotType.SCATTER)   # Y1
    gl_b.add_xy_plot(ws, 0, 2, XYPlotType.SCATTER)   # Y2
    dump_plots(gl_b._obj, "Y1+Y2")

    # ── ケースA での安定性確認: GetDatasetName で区別できるか ─────────────────
    print("\n=== ケースA: GetDatasetName で2つのプロットを区別できるか ===")
    dp_col_a = gl_a._obj.GetDataPlots()
    names_a = [dp_col_a[i].GetDatasetName() for i in range(2) if dp_col_a[i] is not None]
    print(f"  names = {names_a}")
    if len(set(names_a)) == 1:
        print("  → 同一の DatasetName → GetDatasetName では区別不能")
    else:
        print("  → DatasetName が異なる → GetDatasetName で区別可能")

    ranges_a = [dp_col_a[i].GetRange() for i in range(2) if dp_col_a[i] is not None]
    print(f"  ranges = {ranges_a}")
    if len(set(ranges_a)) == 1:
        print("  → 同一の Range → GetRange でも区別不能")
    else:
        print("  → Range が異なる → GetRange で区別可能")

finally:
    origin.close(save_flag=True)
    print("\nDone")
