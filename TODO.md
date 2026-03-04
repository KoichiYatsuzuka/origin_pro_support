# TODO: 未実装機能リスト

プロット作成後に操作できない機能の実装タスク一覧。

---

## 未実装（❌）

- [x] 1. **凡例の編集**
  - `GraphLayer.get_legend()` → `Legend` インスタンスを返す
  - 対象クラス: `GraphLayer`, `Legend` (新規), `LegendLayout` enum (新規)
  - LabTalk `legend` オブジェクト経由で `visible`, `text`, `font_size`, `background`, 位置, レイアウト, 再構成を制御

- [x] 2. **主目盛・副目盛の編集**
  - `TickType` 列挙型は `layer/enums.py` に定義済みだが、設定メソッドが存在しない
  - 対象クラス: `Axis`
  - 実装方針: LabTalk `x.majorTick`, `x.minorTick` 等を `Axis` クラスのメソッドとして追加

- [x] 3. **上軸・右軸の表示**
  - 上軸（Top）・右軸（Right）の表示/非表示切替メソッドがない
  - 対象クラス: `Axis` または `GraphLayer`
  - 実装方針: LabTalk `x2.showAxes`, `y2.showAxes` 等を使った show/hide メソッドを追加

- [x] 4. **軸ラベルの編集**
  - 軸タイトル文字列の get/set メソッドがない
  - 対象クラス: `Axis`
  - 実装方針: LabTalk `x.label.text$`, `y.label.text$` 等を使ったプロパティを追加

- [x] 5. **軸ラベル非表示**
  - 軸ラベルの表示/非表示切替メソッドがない
  - 対象クラス: `Axis`
  - 実装方針: LabTalk `x.showLabel` 等を使った show/hide メソッドを追加

- [x] 6. **マーカーサイズ編集**
  - `DataPlot` にシンボルサイズの get/set プロパティがない
  - 対象クラス: `DataPlot`
  - 実装方針: LabTalk `layer.plot(n).symbol.size` を使ったプロパティを追加

- [ ] 7. **線種編集**
  - `DataPlot` に線スタイル（実線/破線等）の get/set プロパティがない
  - 対象クラス: `DataPlot`
  - 実装方針: LabTalk `layer.plot(n).linestyle` を使ったプロパティを追加（線種の列挙型も定義）

- [ ] 8. **線の太さ編集**
  - `DataPlot` に線幅の get/set プロパティがない
  - 対象クラス: `DataPlot`
  - 実装方針: LabTalk `layer.plot(n).linewidth` を使ったプロパティを追加

---

## 実装あり・要検証（⚠️）

- [x] 9. **軸範囲の編集**（要バグ修正）
  - `Axis.get_range()` / `set_range()` は実装済みだが、`self._obj.LT_get_float()` / `self._obj.LT_execute()` を使っており、`self._obj` が raw OriginExt GraphLayer のため動作しない可能性が高い
  - 対象クラス: `Axis`
  - 修正方針: `Axis` クラスに `api_core` を保持させ、`api_core.LT_execute()` / `api_core.LT_get_var()` に置き換える

- [x] 10. **マーカー形状編集**（要検証）
  - `DataPlot.shape_list` (get/set) は `DataPlot_GetShapeList` / `DataPlot_SetShapeList` 経由で実装済み
  - 単一プロットのシンボル形状を個別に設定できるか未検証
  - 対象クラス: `DataPlot`
  - 確認方針: 実際に `shape_list` を設定してシンボルが変わるか統合テストで確認
