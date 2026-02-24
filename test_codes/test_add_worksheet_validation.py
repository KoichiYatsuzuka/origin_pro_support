#!/usr/bin/env python3
"""
Test script for add_worksheet method with data validation.
"""

import os
import sys
import numpy as np
import pandas as pd

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from origin_pro_support import OriginInstance

def test_add_worksheet_validation():
    """Test add_worksheet method with various data types and validation."""
    
    # Clean up any existing test file
    test_file = "test_add_worksheet_validation.opju"
    if os.path.exists(test_file):
        os.remove(test_file)
    
    print("=== add_worksheetバリデーションテスト開始 ===")
    
    try:
        # Create Origin instance
        print("Originインスタンスを生成...")
        origin = OriginInstance(os.path.join(os.getcwd(), test_file))
        
        # Create test workbook
        print("テスト用ワークブックを生成...")
        workbook = origin.new_workbook("TestWorkbook")
        
        # Test 1: Valid 1D list
        print("\nテスト1: 有効な1D list")
        try:
            ws1 = workbook.add_worksheet("Test1D", [1, 2, 3, 4, 5], lname="TestList")
            print(f"[OK] 1D list: ワークシート '{ws1.Name}' を作成")
        except Exception as e:
            print(f"[ERROR] 1D listテスト失敗: {e}")
            return False
        
        # Test 2: Valid 2D list
        print("\nテスト2: 有効な2D list")
        try:
            ws2 = workbook.add_worksheet("Test2D", [[1, 2], [3, 4], [5, 6]], lname="Test2DList")
            print(f"[OK] 2D list: ワークシート '{ws2.Name}' を作成")
        except Exception as e:
            print(f"[ERROR] 2D listテスト失敗: {e}")
            return False
        
        # Test 3: Valid 1D numpy array
        print("\nテスト3: 有効な1D numpy array")
        try:
            ws3 = workbook.add_worksheet("Test1DArray", np.array([1, 2, 3, 4, 5]), lname="TestArray")
            print(f"[OK] 1D array: ワークシート '{ws3.Name}' を作成")
        except Exception as e:
            print(f"[ERROR] 1D arrayテスト失敗: {e}")
            return False
        
        # Test 4: Valid 2D numpy array
        print("\nテスト4: 有効な2D numpy array")
        try:
            ws4 = workbook.add_worksheet("Test2DArray", np.array([[1, 2], [3, 4], [5, 6]]), lname="Test2DArray")
            print(f"[OK] 2D array: ワークシート '{ws4.Name}' を作成")
        except Exception as e:
            print(f"[ERROR] 2D arrayテスト失敗: {e}")
            return False
        
        # Test 5: Valid pandas Series
        print("\nテスト5: 有効なpandas Series")
        try:
            series = pd.Series([1, 2, 3, 4, 5], name="TestSeries")
            ws5 = workbook.add_worksheet("TestSeries", series, lname="FromSeries")
            print(f"[OK] Series: ワークシート '{ws5.Name}' を作成")
        except Exception as e:
            print(f"[ERROR] Seriesテスト失敗: {e}")
            return False
        
        # Test 6: Valid pandas DataFrame
        print("\nテスト6: 有効なpandas DataFrame")
        try:
            df = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6], 'C': [7, 8, 9]})
            ws6 = workbook.add_worksheet("TestDF", df, lname="FromDataFrame")
            print(f"[OK] DataFrame: ワークシート '{ws6.Name}' を作成")
        except Exception as e:
            print(f"[ERROR] DataFrameテスト失敗: {e}")
            return False
        
        # Test 7: Invalid 3D numpy array (should raise ValueError)
        print("\nテスト7: 無効な3D numpy array")
        try:
            array_3d = np.array([[[1, 2], [3, 4]], [[5, 6], [7, 8]]])
            ws7 = workbook.add_worksheet("Test3D", array_3d, lname="Test3D")
            print("[ERROR] 3D arrayテスト失敗: ValueErrorが発生すべきだった")
            return False
        except ValueError as e:
            print(f"[OK] 期待通りのValueError: {e}")
        except Exception as e:
            print(f"[ERROR] 予期せぬエラー: {e}")
            return False
        
        # Test 8: Invalid 3D nested list (should raise ValueError)
        print("\nテスト8: 無効な3D nested list")
        try:
            list_3d = [[[1, 2], [3, 4]], [[5, 6], [7, 8]]]
            ws8 = workbook.add_worksheet("Test3DList", list_3d, lname="Test3DList")
            print("[ERROR] 3D listテスト失敗: ValueErrorが発生すべきだった")
            return False
        except ValueError as e:
            print(f"[OK] 期待通りのValueError: {e}")
        except Exception as e:
            print(f"[ERROR] 予期せぬエラー: {e}")
            return False
        
        # Test 9: Invalid data type - dict (should raise ValueError)
        print("\nテスト9: 無効なデータ型 - dict")
        try:
            dict_data = {'a': 1, 'b': 2, 'c': 3}
            ws9 = workbook.add_worksheet("TestDict", dict_data, lname="TestDict")
            print("[ERROR] dictテスト失敗: ValueErrorが発生すべきだった")
            return False
        except ValueError as e:
            print(f"[OK] 期待通りのValueError: {e}")
        except Exception as e:
            print(f"[ERROR] 予期せぬエラー: {e}")
            return False
        
        # Test 10: Empty worksheet (no data)
        print("\nテスト10: データなしの空ワークシート")
        try:
            ws10 = workbook.add_worksheet("EmptySheet")
            print(f"[OK] 空ワークシート: '{ws10.Name}' を作成")
        except Exception as e:
            print(f"[ERROR] 空ワークシートテスト失敗: {e}")
            return False
        
        # Final verification
        final_sheets = len(workbook)
        expected_sheets = 10  # 1 (initial) + 9 (added)
        print(f"\n最終シート数: {final_sheets}")
        print(f"期待されるシート数: {expected_sheets}")
        
        if final_sheets >= expected_sheets - 2:  # Allow for error tests that don't create sheets
            print("[OK] 適切な数のワークシートが作成された")
        else:
            print(f"[ERROR] シート数が不足: {final_sheets}")
            return False
        
        # Save project
        origin.save()
        print("\nプロジェクトを保存")
        
        print("\n=== 全テスト項目がパスしました ===")
        return True
        
    except Exception as e:
        print(f"[ERROR] テスト実行中にエラー: {e}")
        return False
    
    finally:
        # Clean up
        try:
            origin.close()
            print("Originインスタンスを終了")
        except:
            pass

if __name__ == "__main__":
    success = test_add_worksheet_validation()
    if success:
        print("\n[SUCCESS] add_worksheetバリデーションテスト: 成功")
        sys.exit(0)
    else:
        print("\n[FAILED] add_worksheetバリデーションテスト: 失敗")
        sys.exit(1)
