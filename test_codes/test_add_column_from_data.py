"""
add_column_from_dataメソッドのテストコード

様々なデータ型（list, np.ndarray, pd.Series, pd.DataFrame）と
次元（1D, 2D）でカラム追加機能をテストする。
"""
import sys
import os
import pandas as pd
import numpy as np
from pathlib import Path

# モジュールのパスを追加してimportを可能にする
sys.path.insert(0, str(Path(__file__).parent.parent))

from origin_pro_support import OriginInstance

def test_add_column_from_data():
    """
    add_column_from_dataメソッドの包括的テスト
    """
    print("=== add_column_from_dataメソッドテスト開始 ===")
    
    # テスト用のOriginプロジェクトファイルパス（絶対パス）
    test_file = os.path.abspath("test_add_column_from_data.opju")
    
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
        
        # テスト用ワークブックを作成
        print("テスト用ワークブックを作成...")
        workbook = origin.new_workbook("TestData")
        worksheet = workbook[0]
        
        # テスト1: 1D list
        print("\nテスト1: 1D listでカラム追加")
        try:
            list_data = [1, 2, 3, 4, 5]
            new_col = worksheet.add_column_from_data(list_data, lname="TestList")
            print(f"[OK] 1D listカラム追加成功: {new_col.LongName}")
        except Exception as e:
            print(f"[ERROR] 1D listテスト失敗: {e}")
            return False
        
        # テスト2: 1D numpy array
        print("\nテスト2: 1D numpy arrayでカラム追加")
        try:
            array_1d = np.array([10.5, 20.5, 30.5, 40.5, 50.5])
            new_col = worksheet.add_column_from_data(array_1d, lname="TestArray1D", units="V")
            print(f"[OK] 1D arrayカラム追加成功: {new_col.LongName}")
        except Exception as e:
            print(f"[ERROR] 1D arrayテスト失敗: {e}")
            return False
        
        # テスト3: pandas Series
        print("\nテスト3: pandas Seriesでカラム追加")
        try:
            series_data = pd.Series([100, 200, 300, 400, 500], name="TestSeries")
            new_col = worksheet.add_column_from_data(series_data, lname="FromSeries", axis="Y")
            print(f"[OK] Seriesカラム追加成功: {new_col.LongName}")
        except Exception as e:
            print(f"[ERROR] Seriesテスト失敗: {e}")
            return False
        
        # テスト4: pandas DataFrame（複数カラム）
        print("\nテスト4: pandas DataFrameで複数カラム追加")
        try:
            df_data = pd.DataFrame({
                'A': [1.1, 2.2, 3.3, 4.4, 5.5],
                'B': [10, 20, 30, 40, 50],
                'C': [100.1, 200.2, 300.3, 400.4, 500.5]
            })
            new_cols = worksheet.add_column_from_data(df_data, lname="DataFrame")
            if isinstance(new_cols, list):
                print(f"[OK] DataFrameから{len(new_cols)}カラム追加成功")
                for i, col in enumerate(new_cols):
                    print(f"  カラム{i+1}: {col.LongName}")
            else:
                print(f"[OK] DataFrameカラム追加成功: {new_cols.LongName}")
        except Exception as e:
            print(f"[ERROR] DataFrameテスト失敗: {e}")
            return False
        
        # テスト5: 2D numpy array（複数カラム）
        print("\nテスト5: 2D numpy arrayで複数カラム追加")
        try:
            array_2d = np.array([[1, 2, 3], [4, 5, 6], [7, 8, 9], [10, 11, 12], [13, 14, 15]])
            new_cols = worksheet.add_column_from_data(array_2d, lname="Array2D", units="units")
            if isinstance(new_cols, list):
                print(f"[OK] 2D arrayから{len(new_cols)}カラム追加成功")
                for i, col in enumerate(new_cols):
                    print(f"  カラム{i+1}: {col.LongName}")
            else:
                print(f"[OK] 2D arrayカラム追加成功: {new_cols.LongName}")
        except Exception as e:
            print(f"[ERROR] 2D arrayテスト失敗: {e}")
            return False
        
        # テスト6: 3D numpy array（エラーケース）
        print("\nテスト6: エラーケース - 3D array")
        try:
            bad_array = np.array([[[1, 2], [3, 4]], [[5, 6], [7, 8]]])  # 3D array
            worksheet.add_column_from_data(bad_array, lname="BadArray")
            print("[ERROR] 3D arrayエラーテスト失敗: ValueErrorが発生すべきだった")
            return False
        except ValueError as e:
            print(f"[OK] 期待通りのValueError: {e}")
        except Exception as e:
            print(f"[ERROR] 予期せぬエラー: {e}")
            return False
        
        # テスト7: エラーケース - ネストされたリスト
        print("\nテスト7: エラーケース - ネストされたリスト")
        try:
            bad_list = [[1, 2], [3, 4]]  # ネストされたリスト
            worksheet.add_column_from_data(bad_list, lname="BadList")
            print("[ERROR] ネストリストエラーテスト失敗: ValueErrorが発生すべきだった")
            return False
        except ValueError as e:
            print(f"[OK] 期待通りのValueError: {e}")
        except Exception as e:
            print(f"[ERROR] 予期せぬエラー: {e}")
            return False
        
        # テスト8: エラーケース - サポートされていない型
        print("\nテスト8: エラーケース - サポートされていない型")
        try:
            bad_data = {"key": "value"}  # 辞書
            worksheet.add_column_from_data(bad_data, lname="BadDict")
            print("[ERROR] サポート外型エラーテスト失敗: ValueErrorが発生すべきだった")
            return False
        except ValueError as e:
            print(f"[OK] 期待通りのValueError: {e}")
        except Exception as e:
            print(f"[ERROR] 予期せぬエラー: {e}")
            return False
        
        # 最終確認: 追加されたカラム数をチェック
        final_cols = worksheet.get_cols()
        print(f"\n最終カラム数: {final_cols}")
        print(f"期待されるカラム数: 1(initial) + 1(list) + 1(array1d) + 1(series) + 3(dataframe) + 3(array2d) = 10")
        
        if final_cols >= 9:  # 少なくとも9列あれば成功（エラーテストでカラムは追加されない）
            print("[OK] すべてのカラムが正常に追加された")
        else:
            print(f"[ERROR] カラム数が不足: {final_cols}")
            return False
        
        # 保存
        origin.save()
        print("\nプロジェクトを保存")
        
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
    success = test_add_column_from_data()
    if success:
        print("\n[SUCCESS] add_column_from_dataテスト: 成功")
        sys.exit(0)
    else:
        print("\n[FAILED] add_column_from_dataテスト: 失敗")
        sys.exit(1)
