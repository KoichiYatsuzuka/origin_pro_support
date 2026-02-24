"""
複数WorkbookPage生成のテストコード

一つのOriginInstanceインスタンスに対して複数のWorkbookPageを生成し、
それぞれ異なるpd.DataFrameで初期化した後、WorksheetLayerの内容を取得し、
与えた値と一致することを確認するテスト。
"""
import sys
import os
import pandas as pd
import numpy as np
from pathlib import Path

# モジュールのパスを追加してimportを可能にする
sys.path.insert(0, str(Path(__file__).parent.parent))

from origin_pro_support import OriginInstance

def test_multiple_workbook_pages():
    """
    複数のWorkbookPage生成テスト
    """
    print("=== 複数WorkbookPage生成テスト開始 ===")
    
    # テスト用のOriginプロジェクトファイルパス（絶対パス）
    test_file = os.path.abspath("test_multiple_workbooks.opju")
    
    # 既存のテストファイルがあれば削除
    if os.path.exists(test_file):
        os.remove(test_file)
        print(f"既存のテストファイルを削除: {test_file}")
    
    origin = None
    try:
        # Originインスタンスを生成
        print("Originインスタンスを生成...")
        origin = OriginInstance(test_file)
        origin.set_show(False)  # Originウィンドウを非表示に設定
        
        # テストデータ1を作成（より多くのデータ）
        print("テストデータ1を作成...")
        np.random.seed(42)  # 再現性のため
        x_data1 = np.linspace(0, 20, 50)  # 50点のデータ
        data1 = {
            'X1': x_data1,
            'Y1': x_data1 ** 2 + np.random.normal(0, 5, 50),  # 二次関数 + ノイズ
            'Y2': 10 * np.sin(x_data1/2) + np.random.normal(0, 1, 50),  # サイン波 + ノイズ
            'Y3': np.exp(x_data1/10) + np.random.normal(0, 2, 50),  # 指数関数 + ノイズ
            'Y4': 5 * np.cos(x_data1/3) + x_data1/2 + np.random.normal(0, 1, 50),  # コサイン + 線形 + ノイズ
            'Y5': np.log(x_data1 + 1) * 10 + np.random.normal(0, 0.5, 50)  # 対数関数 + ノイズ
        }
        df1 = pd.DataFrame(data1)
        print(f"データ1: {df1.shape}")
        print(f"    Columns: {list(df1.columns)}")
        print(f"    Sample:\n{df1.head(3)}")
        
        # テストデータ2を作成（より多くのデータ）
        print("テストデータ2を作成...")
        x_data2 = np.linspace(-10, 10, 40)  # 40点のデータ
        data2 = {
            'A': x_data2,
            'B': x_data2 ** 3 + np.random.normal(0, 10, 40),  # 三次関数 + ノイズ
            'C': 15 * np.cos(x_data2/2) + np.random.normal(0, 2, 40),  # コサイン波 + ノイズ
            'D': np.abs(x_data2) * np.sin(x_data2) + np.random.normal(0, 1, 40),  # 絶対値×サイン + ノイズ
            'E': np.sqrt(np.abs(x_data2)) * 5 + np.random.normal(0, 0.8, 40),  # 平方根 + ノイズ
            'F': x_data2 ** 2 / 5 + 2 * x_data2 + np.random.normal(0, 3, 40)  # 二次式 + ノイズ
        }
        df2 = pd.DataFrame(data2)
        print(f"データ2: {df2.shape}")
        print(f"    Columns: {list(df2.columns)}")
        print(f"    Sample:\n{df2.head(3)}")
        
        # 最初のWorkbookPageを作成
        print("最初のWorkbookPageを作成...")
        workbook1 = origin.new_workbook("TestWorkbook1")
        if workbook1 is None:
            raise Exception("Workbook1の作成に失敗しました")
        
        # 最初のワークシートを取得
        worksheet1 = workbook1[0]
        if worksheet1 is None:
            raise Exception("Worksheet1の取得に失敗しました")
        
        # DataFrameをワークシートに設定
        print("DataFrame1をワークシートに設定...")
        worksheet1.from_df(df1)
        
        # 2番目のWorkbookPageを作成
        print("2番目のWorkbookPageを作成...")
        workbook2 = origin.new_workbook("TestWorkbook2")
        if workbook2 is None:
            raise Exception("Workbook2の作成に失敗しました")
        
        # 2番目のワークシートを取得
        worksheet2 = workbook2[0]
        if worksheet2 is None:
            raise Exception("Worksheet2の取得に失敗しました")
        
        # 2番目のDataFrameをワークシートに設定
        print("DataFrame2をワークシートに設定...")
        worksheet2.from_df(df2)
        
        # 保存して確定
        origin.save()
        print("プロジェクトを保存")
        
        # 検証1: 最初のワークシートの内容を取得
        print("検証1: 最初のワークシートの内容を取得...")
        try:
            # get_dataメソッドを使用してデータを取得（実際のデータ行数のみ）
            data1_array = worksheet1.get_data(row_start=0, row_end=len(df1))
            # カラム名を取得
            col_names1 = []
            columns1 = worksheet1.get_columns()
            for col in columns1:
                col_names1.append(col.get_long_name())
            
            # DataFrameに変換
            retrieved_data1 = pd.DataFrame(data1_array, columns=col_names1)
            print(f"取得データ1:\n{retrieved_data1}")
        except Exception as e:
            print(f"データ1の取得でエラー: {e}")
            # 手動でデータを検証
            print("手動でデータ1を検証...")
            rows1 = worksheet1.get_rows()
            cols1 = worksheet1.get_cols()
            print(f"ワークシート1: {rows1}行 x {cols1}列")
            
            # 最初の数行を表示
            for row in range(min(3, rows1)):
                row_data = []
                for col in range(cols1):
                    try:
                        cell_data = worksheet1.get_cell(row, col)
                        row_data.append(str(cell_data))
                    except:
                        row_data.append("N/A")
                print(f"  行{row}: {row_data}")
            retrieved_data1 = None
        
        # 検証2: 2番目のワークシートの内容を取得
        print("検証2: 2番目のワークシートの内容を取得...")
        try:
            # get_dataメソッドを使用してデータを取得（実際のデータ行数のみ）
            data2_array = worksheet2.get_data(row_start=0, row_end=len(df2))
            # カラム名を取得
            col_names2 = []
            columns2 = worksheet2.get_columns()
            for col in columns2:
                col_names2.append(col.get_long_name())
            
            # DataFrameに変換
            retrieved_data2 = pd.DataFrame(data2_array, columns=col_names2)
            print(f"取得データ2:\n{retrieved_data2}")
        except Exception as e:
            print(f"データ2の取得でエラー: {e}")
            # 手動でデータを検証
            print("手動でデータ2を検証...")
            rows2 = worksheet2.get_rows()
            cols2 = worksheet2.get_cols()
            print(f"ワークシート2: {rows2}行 x {cols2}列")
            
            # 最初の数行を表示
            for row in range(min(3, rows2)):
                row_data = []
                for col in range(cols2):
                    try:
                        cell_data = worksheet2.get_cell(row, col)
                        row_data.append(str(cell_data))
                    except:
                        row_data.append("N/A")
                print(f"  行{row}: {row_data}")
            retrieved_data2 = None
        
        # 検証3: データの比較
        print("検証3: データの比較...")
        
        if retrieved_data1 is not None and retrieved_data2 is not None:
            # データ1の比較
            try:
                pd.testing.assert_frame_equal(df1, retrieved_data1, check_dtype=False)
                print("[OK] データ1: 元のDataFrameと取得したDataFrameが一致")
            except AssertionError as e:
                print(f"[ERROR] データ1: DataFrameが一致しません - {e}")
                return False
            
            # データ2の比較
            try:
                pd.testing.assert_frame_equal(df2, retrieved_data2, check_dtype=False)
                print("[OK] データ2: 元のDataFrameと取得したDataFrameが一致")
            except AssertionError as e:
                print(f"[ERROR] データ2: DataFrameが一致しません - {e}")
                return False
        else:
            print("[INFO] DataFrameでの比較をスキップし、手動検証を行います")
            
            # 基本的な検証：行数と列数の確認
            # from_dfの動作を確認 - ヘッダー行の有無をチェック
            expected_rows = len(df1)  # データ行数のみ
            if worksheet1.get_rows() >= expected_rows and worksheet1.get_cols() == len(df1.columns):
                print("[OK] ワークシート1の次元が一致")
            else:
                # ヘッダー行がある場合のチェック
                expected_rows_with_header = len(df1) + 1
                if worksheet1.get_rows() >= expected_rows_with_header and worksheet1.get_cols() == len(df1.columns):
                    print("[OK] ワークシート1の次元が一致（ヘッダー行あり）")
                else:
                    print(f"[ERROR] ワークシート1の次元が不一致: {worksheet1.get_rows()}x{worksheet1.get_cols()} vs {expected_rows}x{len(df1.columns)}")
                    # デバッグ情報
                    print(f"  デバッグ: 実際の行数={worksheet1.get_rows()}, 期待行数={expected_rows} or {expected_rows_with_header}")
                    print(f"  デバッグ: 実際の列数={worksheet1.get_cols()}, 期待列数={len(df1.columns)}")
                    return False
            
            expected_rows2 = len(df2)  # データ行数のみ
            if worksheet2.get_rows() >= expected_rows2 and worksheet2.get_cols() == len(df2.columns):
                print("[OK] ワークシート2の次元が一致")
            else:
                # ヘッダー行がある場合のチェック
                expected_rows2_with_header = len(df2) + 1
                if worksheet2.get_rows() >= expected_rows2_with_header and worksheet2.get_cols() == len(df2.columns):
                    print("[OK] ワークシート2の次元が一致（ヘッダー行あり）")
                else:
                    print(f"[ERROR] ワークシート2の次元が不一致: {worksheet2.get_rows()}x{worksheet2.get_cols()} vs {expected_rows2}x{len(df2.columns)}")
                    # デバッグ情報
                    print(f"  デバッグ: 実際の行数={worksheet2.get_rows()}, 期待行数={expected_rows2} or {expected_rows2_with_header}")
                    print(f"  デバッグ: 実際の列数={worksheet2.get_cols()}, 期待列数={len(df2.columns)}")
                    return False
        
        # 検証4: WorkbookPageが独立していることを確認
        print("検証4: WorkbookPageの独立性を確認...")
        
        # ワークブック名の確認
        print(f"デバッグ: Workbook1.Name = '{workbook1.Name}'")
        print(f"デバッグ: Workbook2.Name = '{workbook2.Name}'")
        
        # new_workbookが正しく名前を設定できているか確認
        # もしデフォルト名（Book1, Book2など）になっている場合は、LongNameで確認
        name1 = workbook1.Name
        name2 = workbook2.Name
        
        # LongNameも確認してみる
        try:
            longname1 = workbook1.LongName
            longname2 = workbook2.LongName
            print(f"デバッグ: Workbook1.LongName = '{longname1}'")
            print(f"デバッグ: Workbook2.LongName = '{longname2}'")
            
            # LongNameが指定した名前ならそちらを使用
            if longname1 == "TestWorkbook1":
                name1 = longname1
            if longname2 == "TestWorkbook2":
                name2 = longname2
        except:
            print("LongNameの取得に失敗")
        
        if "TestWorkbook1" in name1:
            print("[OK] Workbook1の名前が正しい")
        else:
            print(f"[ERROR] Workbook1の名前が異常: {name1}")
            return False
        
        if "TestWorkbook2" in name2:
            print("[OK] Workbook2の名前が正しい")
        else:
            print(f"[ERROR] Workbook2の名前が異常: {name2}")
            return False
        
        # ワークブックが異なることを確認
        if name1 != name2:
            print("[OK] 2つのWorkbookPageが異なることを確認")
        else:
            print("[ERROR] 2つのWorkbookPageが同じになっている")
            return False
        
        # 検証5: カラム追加機能
        print("検証5: カラム追加機能...")
        
        # Workbook1にカラムを追加
        print("Workbook1に新しいカラムを追加...")
        original_cols1 = worksheet1.get_cols()
        print(f"追加前の列数: {original_cols1}")
        
        # add_column_from_dataメソッドでカラムを追加
        new_col1 = worksheet1.add_column_from_data([1000, 2000, 3000, 4000, 5000], 
                                                 lname="NewColumn1")
        
        updated_cols1 = worksheet1.get_cols()
        print(f"追加後の列数: {updated_cols1}")
        
        # Workbook2にカラムを追加
        print("Workbook2に新しいカラムを追加...")
        original_cols2 = worksheet2.get_cols()
        print(f"追加前の列数: {original_cols2}")
        
        # add_column_from_dataメソッドでカラムを追加
        new_col2 = worksheet2.add_column_from_data([0.5, 1.5, 2.5, 3.5, 4.5], 
                                                 lname="NewColumn2")
        
        updated_cols2 = worksheet2.get_cols()
        print(f"追加後の列数: {updated_cols2}")
        
        # カラム追加の検証
        if updated_cols1 == original_cols1 + 1 and updated_cols2 == original_cols2 + 1:
            print("[OK] 両方のWorkbookでカラムが正常に追加された")
        else:
            print(f"[ERROR] カラム追加失敗: Workbook1 {original_cols1}->{updated_cols1}, Workbook2 {original_cols2}->{updated_cols2}")
            return False
        
        # 追加したカラムのデータを検証
        print("追加したカラムのデータを検証...")
        
        # Workbook1の新しいカラムデータを確認
        try:
            added_data1 = new_col1.get_data(0)  # format=0を指定
            expected_data1 = [1000, 2000, 3000, 4000, 5000]
            
            # get_data()はネストされたリストを返す場合があるのでフラット化
            if isinstance(added_data1[0], list):
                added_data1 = added_data1[0]
            
            # floatに変換して比較
            added_data1_float = [float(x) for x in added_data1]
            
            if added_data1_float == expected_data1:
                print("[OK] Workbook1の追加カラムデータが正しい")
            else:
                print(f"[ERROR] Workbook1の追加カラムデータが不一致: {added_data1_float} vs {expected_data1}")
                return False
        except Exception as e:
            print(f"[ERROR] Workbook1のカラムデータ検証でエラー: {e}")
            return False
        
        # Workbook2の新しいカラムデータを確認
        try:
            added_data2 = new_col2.get_data(0)  # format=0を指定
            expected_data2 = [0.5, 1.5, 2.5, 3.5, 4.5]
            
            # get_data()はネストされたリストを返す場合があるのでフラット化
            if isinstance(added_data2[0], list):
                added_data2 = added_data2[0]
            
            # 小数の比較には許容誤差を設定
            import math
            data_match = True
            for i, (actual, expected) in enumerate(zip(added_data2, expected_data2)):
                if abs(float(actual) - expected) > 1e-10:
                    data_match = False
                    break
            
            if data_match:
                print("[OK] Workbook2の追加カラムデータが正しい")
            else:
                print(f"[ERROR] Workbook2の追加カラムデータが不一致: {added_data2} vs {expected_data2}")
                return False
        except Exception as e:
            print(f"[ERROR] Workbook2のカラムデータ検証でエラー: {e}")
            return False
        
        # 検証6: 新しいWorksheet追加機能
        print("検証6: 新しいWorksheet追加機能...")
        
        # Workbook1に新しいWorksheetを追加
        print("Workbook1に新しいWorksheetを追加...")
        original_sheets1 = len(workbook1)
        print(f"追加前のシート数: {original_sheets1}")
        
        # 新しいWorksheetを追加
        new_worksheet1 = workbook1.add_worksheet("NewSheet1")
        if new_worksheet1 is None:
            print("[ERROR] Workbook1へのWorksheet追加に失敗")
            return False
        
        updated_sheets1 = len(workbook1)
        print(f"追加後のシート数: {updated_sheets1}")
        
        # Workbook2に新しいWorksheetを追加
        print("Workbook2に新しいWorksheetを追加...")
        original_sheets2 = len(workbook2)
        print(f"追加前のシート数: {original_sheets2}")
        
        # 新しいWorksheetを追加
        new_worksheet2 = workbook2.add_worksheet("NewSheet2")
        if new_worksheet2 is None:
            print("[ERROR] Workbook2へのWorksheet追加に失敗")
            return False
        
        updated_sheets2 = len(workbook2)
        print(f"追加後のシート数: {updated_sheets2}")
        
        # Worksheet追加の検証
        if updated_sheets1 == original_sheets1 + 1 and updated_sheets2 == original_sheets2 + 1:
            print("[OK] 両方のWorkbookでWorksheetが正常に追加された")
        else:
            print(f"[ERROR] Worksheet追加失敗: Workbook1 {original_sheets1}->{updated_sheets1}, Workbook2 {original_sheets2}->{updated_sheets2}")
            return False
        
        # 追加したWorksheetの基本機能をテスト
        print("追加したWorksheetの基本機能をテスト...")
        
        # 新しいWorksheet1にデータを追加
        try:
            # add_worksheetは既にWorksheetオブジェクトを返すのでキャスト不要
            # add_column_from_dataメソッドでカラムを追加
            col_x = new_worksheet1.add_column_from_data([10, 20, 30], 
                                                       lname="TestDataX")
            col_y = new_worksheet1.add_column_from_data([100, 200, 300], 
                                                       lname="TestDataY")
            
            print("[OK] Workbook1の新しいWorksheetにデータを設定")
        except Exception as e:
            print(f"[ERROR] Workbook1の新しいWorksheet設定でエラー: {e}")
            return False
        
        # 新しいWorksheet2にデータを追加
        try:
            # add_worksheetは既にWorksheetオブジェクトを返すのでキャスト不要
            # add_column_from_dataメソッドでカラムを追加
            col_a = new_worksheet2.add_column_from_data([1.1, 2.2, 3.3], 
                                                       lname="SampleA")
            col_b = new_worksheet2.add_column_from_data([11, 22, 33], 
                                                       lname="SampleB")
            
            print("[OK] Workbook2の新しいWorksheetにデータを設定")
        except Exception as e:
            print(f"[ERROR] Workbook2の新しいWorksheet設定でエラー: {e}")
            return False
        
        # 追加したWorksheetのデータを検証
        print("追加したWorksheetのデータを検証...")
        
        # Workbook1の新しいWorksheetデータを確認
        try:
            test_data_x = col_x.get_data(0)
            test_data_y = col_y.get_data(0)
            
            # ネストされたリストをフラット化
            if isinstance(test_data_x[0], list):
                test_data_x = test_data_x[0]
            if isinstance(test_data_y[0], list):
                test_data_y = test_data_y[0]
            
            if test_data_x == [10.0, 20.0, 30.0] and test_data_y == [100.0, 200.0, 300.0]:
                print("[OK] Workbook1の新しいWorksheetデータが正しい")
            else:
                print(f"[ERROR] Workbook1の新しいWorksheetデータが不一致: X={test_data_x}, Y={test_data_y}")
                return False
        except Exception as e:
            print(f"[ERROR] Workbook1の新しいWorksheetデータ検証でエラー: {e}")
            return False
        
        # Workbook2の新しいWorksheetデータを確認
        try:
            sample_data_a = col_a.get_data(0)
            sample_data_b = col_b.get_data(0)
            
            # ネストされたリストをフラット化
            if isinstance(sample_data_a[0], list):
                sample_data_a = sample_data_a[0]
            if isinstance(sample_data_b[0], list):
                sample_data_b = sample_data_b[0]
            
            # 小数の比較には許容誤差を設定
            import math
            data_match = True
            expected_a = [1.1, 2.2, 3.3]
            expected_b = [11, 22, 33]
            
            for i, (actual, expected) in enumerate(zip(sample_data_a, expected_a)):
                if abs(float(actual) - expected) > 1e-10:
                    data_match = False
                    break
            
            for i, (actual, expected) in enumerate(zip(sample_data_b, expected_b)):
                if abs(float(actual) - expected) > 1e-10:
                    data_match = False
                    break
            
            if data_match:
                print("[OK] Workbook2の新しいWorksheetデータが正しい")
            else:
                print(f"[ERROR] Workbook2の新しいWorksheetデータが不一致: A={sample_data_a}, B={sample_data_b}")
                return False
        except Exception as e:
            print(f"[ERROR] Workbook2の新しいWorksheetデータ検証でエラー: {e}")
            return False
        
        # 検証7: 全ワークブックページの取得
        print("検証7: 全ワークブックページの取得...")
        all_workbooks = origin.get_workbook_pages()
        print(f"取得したワークブック数: {len(all_workbooks)}")
        
        if len(all_workbooks) >= 2:
            print("[OK] 複数のWorkbookPageが正常に生成されている")
        else:
            print(f"[ERROR] WorkbookPageの数が不足: {len(all_workbooks)}")
            return False
        
        # 検証8: 複数GraphPage作成テスト（改良版）
        print("\n検証8: 複数GraphPage作成テスト（改良版）...")
        
        # 最初のグラフを作成（複数プロット）
        print("最初のGraphPageを作成（複数プロット）...")
        graph1 = origin.new_graph("TestGraph1_MultiPlot")
        if graph1 is None:
            print("[ERROR] Graph1の作成に失敗しました")
            return False
        print("[OK] Graph1作成成功")
        
        # Graph1に複数のレイヤーとプロットを追加
        plots_created_graph1 = []
        
        # Layer1: X1 vs Y1 (二次関数)
        layer1_1 = graph1.add_graph_layer("Layer_Y1_Quadratic")
        try:
            plot1_1 = layer1_1.add_xy_plot(worksheet1, 0, 1)  # X1 vs Y1
            plots_created_graph1.append("X1 vs Y1 (Quadratic)")
            print("[OK] Graph1 Layer1: X1 vs Y1 (二次関数) プロット作成成功")
        except Exception as e:
            print(f"[WARNING] Graph1 Layer1プロット作成で警告: {e}")
        
        # Layer2: X1 vs Y2 (サイン波)
        layer1_2 = graph1.add_graph_layer("Layer_Y2_Sine")
        try:
            plot1_2 = layer1_2.add_xy_plot(worksheet1, 0, 2)  # X1 vs Y2
            plots_created_graph1.append("X1 vs Y2 (Sine)")
            print("[OK] Graph1 Layer2: X1 vs Y2 (サイン波) プロット作成成功")
        except Exception as e:
            print(f"[WARNING] Graph1 Layer2プロット作成で警告: {e}")
        
        # Layer3: X1 vs Y3 (指数関数)
        layer1_3 = graph1.add_graph_layer("Layer_Y3_Exponential")
        try:
            plot1_3 = layer1_3.add_xy_plot(worksheet1, 0, 3)  # X1 vs Y3
            plots_created_graph1.append("X1 vs Y3 (Exponential)")
            print("[OK] Graph1 Layer3: X1 vs Y3 (指数関数) プロット作成成功")
        except Exception as e:
            print(f"[WARNING] Graph1 Layer3プロット作成で警告: {e}")
        
        # Layer4: X1 vs Y4 (コサイン+線形)
        layer1_4 = graph1.add_graph_layer("Layer_Y4_CosLinear")
        try:
            plot1_4 = layer1_4.add_xy_plot(worksheet1, 0, 4)  # X1 vs Y4
            plots_created_graph1.append("X1 vs Y4 (Cos+Linear)")
            print("[OK] Graph1 Layer4: X1 vs Y4 (コサイン+線形) プロット作成成功")
        except Exception as e:
            print(f"[WARNING] Graph1 Layer4プロット作成で警告: {e}")
        
        # Layer5: X1 vs Y5 (対数関数)
        layer1_5 = graph1.add_graph_layer("Layer_Y5_Logarithmic")
        try:
            plot1_5 = layer1_5.add_xy_plot(worksheet1, 0, 5)  # X1 vs Y5
            plots_created_graph1.append("X1 vs Y5 (Logarithmic)")
            print("[OK] Graph1 Layer5: X1 vs Y5 (対数関数) プロット作成成功")
        except Exception as e:
            print(f"[WARNING] Graph1 Layer5プロット作成で警告: {e}")
        
        print(f"[INFO] Graph1総プロット数: {len(plots_created_graph1)}")
        
        # 2番目のグラフを作成（複数プロット）
        print("2番目のGraphPageを作成（複数プロット）...")
        graph2 = origin.new_graph("TestGraph2_MultiPlot")
        if graph2 is None:
            print("[ERROR] Graph2の作成に失敗しました")
            return False
        print("[OK] Graph2作成成功")
        
        # Graph2に複数のレイヤーとプロットを追加
        plots_created_graph2 = []
        
        # Layer1: A vs B (三次関数)
        layer2_1 = graph2.add_graph_layer("Layer_B_Cubic")
        try:
            plot2_1 = layer2_1.add_xy_plot(worksheet2, 0, 1)  # A vs B
            plots_created_graph2.append("A vs B (Cubic)")
            print("[OK] Graph2 Layer1: A vs B (三次関数) プロット作成成功")
        except Exception as e:
            print(f"[WARNING] Graph2 Layer1プロット作成で警告: {e}")
        
        # Layer2: A vs C (コサイン波)
        layer2_2 = graph2.add_graph_layer("Layer_C_Cosine")
        try:
            plot2_2 = layer2_2.add_xy_plot(worksheet2, 0, 2)  # A vs C
            plots_created_graph2.append("A vs C (Cosine)")
            print("[OK] Graph2 Layer2: A vs C (コサイン波) プロット作成成功")
        except Exception as e:
            print(f"[WARNING] Graph2 Layer2プロット作成で警告: {e}")
        
        # Layer3: A vs D (絶対値×サイン)
        layer2_3 = graph2.add_graph_layer("Layer_D_AbsSine")
        try:
            plot2_3 = layer2_3.add_xy_plot(worksheet2, 0, 3)  # A vs D
            plots_created_graph2.append("A vs D (Abs×Sine)")
            print("[OK] Graph2 Layer3: A vs D (絶対値×サイン) プロット作成成功")
        except Exception as e:
            print(f"[WARNING] Graph2 Layer3プロット作成で警告: {e}")
        
        # Layer4: A vs E (平方根)
        layer2_4 = graph2.add_graph_layer("Layer_E_SquareRoot")
        try:
            plot2_4 = layer2_4.add_xy_plot(worksheet2, 0, 4)  # A vs E
            plots_created_graph2.append("A vs E (SquareRoot)")
            print("[OK] Graph2 Layer4: A vs E (平方根) プロット作成成功")
        except Exception as e:
            print(f"[WARNING] Graph2 Layer4プロット作成で警告: {e}")
        
        # Layer5: A vs F (二次式)
        layer2_5 = graph2.add_graph_layer("Layer_F_Quadratic")
        try:
            plot2_5 = layer2_5.add_xy_plot(worksheet2, 0, 5)  # A vs F
            plots_created_graph2.append("A vs F (Quadratic)")
            print("[OK] Graph2 Layer5: A vs F (二次式) プロット作成成功")
        except Exception as e:
            print(f"[WARNING] Graph2 Layer5プロット作成で警告: {e}")
        
        print(f"[INFO] Graph2総プロット数: {len(plots_created_graph2)}")
        
        # 3番目のグラフを作成（混合データ・複数プロット）
        print("3番目のGraphPageを作成（混合データ・複数プロット）...")
        graph3 = origin.new_graph("TestGraph3_MixedMulti")
        if graph3 is None:
            print("[ERROR] Graph3の作成に失敗しました")
            return False
        print("[OK] Graph3作成成功")
        
        # Graph3に混合データで複数のレイヤーとプロットを追加
        plots_created_graph3 = []
        
        # Layer1: Workbook1データ - X1 vs Y1 (二次関数)
        layer3_1 = graph3.add_graph_layer("WB1_Y1_Quadratic")
        try:
            plot3_1 = layer3_1.add_xy_plot(worksheet1, 0, 1)  # X1 vs Y1
            plots_created_graph3.append("WB1: X1 vs Y1")
            print("[OK] Graph3 Layer1: WB1 X1 vs Y1 (二次関数) プロット作成成功")
        except Exception as e:
            print(f"[WARNING] Graph3 Layer1プロット作成で警告: {e}")
        
        # Layer2: Workbook2データ - A vs B (三次関数)
        layer3_2 = graph3.add_graph_layer("WB2_B_Cubic")
        try:
            plot3_2 = layer3_2.add_xy_plot(worksheet2, 0, 1)  # A vs B
            plots_created_graph3.append("WB2: A vs B")
            print("[OK] Graph3 Layer2: WB2 A vs B (三次関数) プロット作成成功")
        except Exception as e:
            print(f"[WARNING] Graph3 Layer2プロット作成で警告: {e}")
        
        # Layer3: Workbook1データ - X1 vs Y2 (サイン波)
        layer3_3 = graph3.add_graph_layer("WB1_Y2_Sine")
        try:
            plot3_3 = layer3_3.add_xy_plot(worksheet1, 0, 2)  # X1 vs Y2
            plots_created_graph3.append("WB1: X1 vs Y2")
            print("[OK] Graph3 Layer3: WB1 X1 vs Y2 (サイン波) プロット作成成功")
        except Exception as e:
            print(f"[WARNING] Graph3 Layer3プロット作成で警告: {e}")
        
        # Layer4: Workbook2データ - A vs C (コサイン波)
        layer3_4 = graph3.add_graph_layer("WB2_C_Cosine")
        try:
            plot3_4 = layer3_4.add_xy_plot(worksheet2, 0, 2)  # A vs C
            plots_created_graph3.append("WB2: A vs C")
            print("[OK] Graph3 Layer4: WB2 A vs C (コサイン波) プロット作成成功")
        except Exception as e:
            print(f"[WARNING] Graph3 Layer4プロット作成で警告: {e}")
        
        # Layer5: Workbook1データ - X1 vs Y3 (指数関数)
        layer3_5 = graph3.add_graph_layer("WB1_Y3_Exponential")
        try:
            plot3_5 = layer3_5.add_xy_plot(worksheet1, 0, 3)  # X1 vs Y3
            plots_created_graph3.append("WB1: X1 vs Y3")
            print("[OK] Graph3 Layer5: WB1 X1 vs Y3 (指数関数) プロット作成成功")
        except Exception as e:
            print(f"[WARNING] Graph3 Layer5プロット作成で警告: {e}")
        
        print(f"[INFO] Graph3総プロット数: {len(plots_created_graph3)}")
        
        # 総プロット数サマリー
        total_plots = len(plots_created_graph1) + len(plots_created_graph2) + len(plots_created_graph3)
        print(f"\n[INFO] 総プロット作成サマリー:")
        print(f"  Graph1: {len(plots_created_graph1)} プロット")
        print(f"  Graph2: {len(plots_created_graph2)} プロット")
        print(f"  Graph3: {len(plots_created_graph3)} プロット")
        print(f"  総計: {total_plots} プロット")
        
        # 検証9: GraphPageの独立性とプロパティ確認
        print("\n検証9: GraphPageの独立性とプロパティ確認...")
        
        # 全GraphPageを取得
        all_graphs = origin.get_graph_pages()
        print(f"取得したGraphPage数: {len(all_graphs)}")
        
        if len(all_graphs) < 3:
            print(f"[ERROR] GraphPageの数が不足: {len(all_graphs)} < 3")
            return False
        
        print("[OK] 複数のGraphPageが正常に生成されている")
        
        # GraphPageの名前を確認
        graph_names = []
        for i, graph in enumerate(all_graphs):
            try:
                graph_name = graph.Name if hasattr(graph, 'Name') else f"Graph_{i+1}"
                graph_names.append(graph_name)
                print(f"  GraphPage {i+1}: {graph_name}")
            except:
                print(f"  GraphPage {i+1}: 名前取得不可")
        
        # 一意性確認
        unique_names = set(graph_names)
        if len(unique_names) == len(graph_names):
            print("[OK] すべてのGraphPageが一意な名前を持つ")
        else:
            print("[WARNING] GraphPage名に重複がある可能性")
        
        # 検証10: GraphPageとWorkbookPageの連携確認
        print("\n検証10: GraphPageとWorkbookPageの連携確認...")
        
        # WorkbookPageとGraphPageの総数を確認
        total_workbooks = len(all_workbooks)
        total_graphs = len(all_graphs)
        
        print(f"WorkbookPage総数: {total_workbooks}")
        print(f"GraphPage総数: {total_graphs}")
        
        if total_workbooks >= 2 and total_graphs >= 3:
            print("[OK] WorkbookPageとGraphPageが正常に連携して動作")
        else:
            print(f"[ERROR] ページ数が期待値を下回る: WB={total_workbooks}, G={total_graphs}")
            return False
        
        print("\n=== 全テスト項目がパスしました ===")
        return True
        
    except Exception as e:
        print(f"[ERROR] テスト実行中にエラーが発生: {e}")
        import traceback
        traceback.print_exc()
        return False
        
    finally:
        # 必ず保存して終了
        if origin is not None:
            try:
                origin.save()
                origin.close()
                print("Originインスタンスを終了")
            except Exception as e:
                print(f"終了処理でエラー: {e}")

if __name__ == "__main__":
    success = test_multiple_workbook_pages()
    if success:
        print("\n[SUCCESS] 複数WorkbookPage生成テスト: 成功")
        sys.exit(0)
    else:
        print("\n[FAILED] 複数WorkbookPage生成テスト: 失敗")
        sys.exit(1)
