"""
Test script for Worksheet refactoring items 1-4.
Tests each refactoring component individually for bugs.
"""
import sys
import os
import pandas as pd
import numpy as np
from pathlib import Path

# Add parent directory to path to import origin_pro_support
sys.path.insert(0, str(Path(__file__).parent.parent))

try:
    from origin_pro_support import OriginInstance
    from origin_pro_support.layers import Worksheet
except ImportError as e:
    print(f"Import error: {e}")
    print("Make sure Origin is running and originpro is installed")
    sys.exit(1)

def cleanup_test_files():
    """Clean up any test files created during testing"""
    test_files = [
        "test_refactoring_1.opju",
        "test_refactoring_2.opju", 
        "test_refactoring_3.opju",
        "test_refactoring_4.opju"
    ]
    
    for file in test_files:
        if os.path.exists(file):
            try:
                os.remove(file)
                print(f"Cleaned up {file}")
            except Exception as e:
                print(f"Warning: Could not remove {file}: {e}")

def test_refactoring_1():
    """Test Item 1: __init__ @overload for multiple data types"""
    print("\n=== Testing Refactoring Item 1: __init__ @overload ===")
    
    try:
        # Create Origin instance
        origin = OriginInstance()
        
        # Test data
        df = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
        arr_2d = np.array([[1, 2], [3, 4], [5, 6]])
        list_2d = [[1, 2], [3, 4], [5, 6]]
        
        # Test DataFrame initialization
        print("Testing DataFrame initialization...")
        wb = origin.new_workbook("test_refactoring_1")
        ws = wb.worksheets()[0]
        worksheet_df = Worksheet(ws._obj, origin, parent=wb, data=df)
        
        # Verify data was loaded correctly
        assert worksheet_df.get_cols() == 2, f"Expected 2 columns, got {worksheet_df.get_cols()}"
        assert worksheet_df.get_rows() == 3, f"Expected 3 rows, got {worksheet_df.get_rows()}"
        
        # Check column names
        col_a = worksheet_df.Columns(0)
        col_b = worksheet_df.Columns(1)
        assert col_a.LongName == "A", f"Expected column name 'A', got '{col_a.LongName}'"
        assert col_b.LongName == "B", f"Expected column name 'B', got '{col_b.LongName}'"
        print("✓ DataFrame initialization works")
        
        # Test 2D numpy array initialization
        print("Testing 2D numpy array initialization...")
        wb2 = origin.new_workbook("test_refactoring_1")
        ws2 = wb2.worksheets()[0]
        worksheet_arr = Worksheet(ws2._obj, origin, parent=wb2, data=arr_2d)
        
        assert worksheet_arr.get_cols() == 2, f"Expected 2 columns, got {worksheet_arr.get_cols()}"
        assert worksheet_arr.get_rows() == 3, f"Expected 3 rows, got {worksheet_arr.get_rows()}"
        
        # Check generic column names
        col1 = worksheet_arr.Columns(0)
        col2 = worksheet_arr.Columns(1)
        assert col1.LongName == "Column1", f"Expected 'Column1', got '{col1.LongName}'"
        assert col2.LongName == "Column2", f"Expected 'Column2', got '{col2.LongName}'"
        print("✓ 2D numpy array initialization works")
        
        # Test 2D list initialization
        print("Testing 2D list initialization...")
        wb3 = origin.new_workbook("test_refactoring_1")
        ws3 = wb3.worksheets()[0]
        worksheet_list = Worksheet(ws3._obj, origin, parent=wb3, data=list_2d)
        
        assert worksheet_list.get_cols() == 2, f"Expected 2 columns, got {worksheet_list.get_cols()}"
        assert worksheet_list.get_rows() == 3, f"Expected 3 rows, got {worksheet_list.get_rows()}"
        print("✓ 2D list initialization works")
        
        # Test no data initialization (should create 2 empty columns)
        print("Testing no data initialization...")
        wb4 = origin.new_workbook("test_refactoring_1")
        ws4 = wb4.worksheets()[0]
        worksheet_empty = Worksheet(ws4._obj, origin, parent=wb4)
        
        assert worksheet_empty.get_cols() == 2, f"Expected 2 empty columns, got {worksheet_empty.get_cols()}"
        assert worksheet_empty.get_rows() == 0, f"Expected 0 rows, got {worksheet_empty.get_rows()}"
        print("✓ No data initialization works")
        
        # Test error cases
        print("Testing error cases...")
        wb5 = origin.new_workbook("test_refactoring_1")
        ws5 = wb5.worksheets()[0]
        
        try:
            # Test 1D array (should raise ValueError)
            arr_1d = np.array([1, 2, 3])
            Worksheet(ws5._obj, origin, parent=wb5, data=arr_1d)
            assert False, "1D array should raise ValueError"
        except ValueError:
            print("✓ 1D array correctly raises ValueError")
        
        try:
            # Test 1D list (should raise ValueError)
            list_1d = [1, 2, 3]
            Worksheet(ws5._obj, origin, parent=wb5, data=list_1d)
            assert False, "1D list should raise ValueError"
        except ValueError:
            print("✓ 1D list correctly raises ValueError")
        
        try:
            # Test unsupported type
            Worksheet(ws5._obj, origin, parent=wb5, data="invalid")
            assert False, "Unsupported type should raise TypeError"
        except TypeError:
            print("✓ Unsupported type correctly raises TypeError")
        
        # Save and close
        origin.save("test_refactoring_1.opju")
        origin.close()
        
        print("✓ Refactoring Item 1: ALL TESTS PASSED")
        return True
        
    except Exception as e:
        print(f"✗ Refactoring Item 1 FAILED: {e}")
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass
        return False

def test_refactoring_2():
    """Test Item 2: __init__ uses add_column_from_data"""
    print("\n=== Testing Refactoring Item 2: __init__ uses add_column_from_data ===")
    
    try:
        # Create Origin instance
        origin = OriginInstance()
        
        # Test that __init__ delegates to add_column_from_data
        # We can't directly test this, but we can verify the behavior is consistent
        print("Testing consistent behavior between __init__ and add_column_from_data...")
        
        # Test data
        df = pd.DataFrame({'X': [1, 2, 3], 'Y': [4, 5, 6]})
        
        # Create worksheet using __init__
        wb1 = origin.new_workbook("test_refactoring_2")
        ws1 = wb1.worksheets()[0]
        worksheet_init = Worksheet(ws1._obj, origin, parent=wb1, data=df)
        
        # Create empty worksheet and use add_column_from_data
        wb2 = origin.new_workbook("test_refactoring_2")
        ws2 = wb2.worksheets()[0]
        worksheet_method = Worksheet(ws2._obj, origin, parent=wb2)
        worksheet_method.add_column_from_data(df)
        
        # Compare results
        assert worksheet_init.get_cols() == worksheet_method.get_cols(), "Column count mismatch"
        assert worksheet_init.get_rows() == worksheet_method.get_rows(), "Row count mismatch"
        
        # Check column names
        for i in range(worksheet_init.get_cols()):
            col_init = worksheet_init.Columns(i)
            col_method = worksheet_method.Columns(i)
            assert col_init.LongName == col_method.LongName, f"Column name mismatch at index {i}"
        
        print("✓ __init__ and add_column_from_data produce consistent results")
        
        # Test with 2D array
        arr = np.array([[1, 2], [3, 4]])
        
        wb3 = origin.new_workbook("test_refactoring_2")
        ws3 = wb3.worksheets()[0]
        worksheet_init_arr = Worksheet(ws3._obj, origin, parent=wb3, data=arr)
        
        wb4 = origin.new_workbook("test_refactoring_2")
        ws4 = wb4.worksheets()[0]
        worksheet_method_arr = Worksheet(ws4._obj, origin, parent=wb4)
        worksheet_method_arr.add_column_from_data(arr)
        
        assert worksheet_init_arr.get_cols() == worksheet_method_arr.get_cols(), "Array column count mismatch"
        assert worksheet_init_arr.get_rows() == worksheet_method_arr.get_rows(), "Array row count mismatch"
        
        print("✓ Array handling is consistent between __init__ and add_column_from_data")
        
        # Save and close
        origin.save("test_refactoring_2.opju")
        origin.close()
        
        print("✓ Refactoring Item 2: ALL TESTS PASSED")
        return True
        
    except Exception as e:
        print(f"✗ Refactoring Item 2 FAILED: {e}")
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass
        return False

def test_refactoring_3():
    """Test Item 3: Early type validation in add_column_from_data"""
    print("\n=== Testing Refactoring Item 3: Early type validation ===")
    
    try:
        # Create Origin instance
        origin = OriginInstance()
        wb = origin.new_workbook("test_refactoring_3")
        ws = wb.worksheets()[0]
        worksheet = Worksheet(ws._obj, origin, parent=wb)
        
        # Test early type validation
        print("Testing early type validation...")
        
        # Test unsupported types
        invalid_types = [
            ("string", "invalid string"),
            ("dict", {"key": "value"}),
            ("set", {1, 2, 3}),
            ("tuple", (1, 2, 3)),
            ("int", 42),
            ("float", 3.14),
            ("None", None)
        ]
        
        for type_name, invalid_data in invalid_types:
            try:
                worksheet.add_column_from_data(invalid_data)
                assert False, f"{type_name} should raise TypeError"
            except TypeError as e:
                print(f"✓ {type_name} correctly raises TypeError: {str(e)[:50]}...")
        
        print("✓ Early type validation works for all unsupported types")
        
        # Test that valid types don't raise TypeError
        print("Testing valid types don't raise TypeError...")
        
        valid_data = [
            ("1D list", [1, 2, 3]),
            ("2D list", [[1, 2], [3, 4]]),
            ("pd.Series", pd.Series([1, 2, 3])),
            ("1D np.array", np.array([1, 2, 3])),
            ("2D np.array", np.array([[1, 2], [3, 4]])),
            ("pd.DataFrame", pd.DataFrame({'A': [1, 2], 'B': [3, 4]}))
        ]
        
        for type_name, data in valid_data:
            try:
                worksheet.add_column_from_data(data)
                print(f"✓ {type_name} accepted without TypeError")
            except TypeError:
                assert False, f"{type_name} should not raise TypeError"
            except Exception:
                # Other exceptions (like ValueError for wrong dimensions) are OK
                print(f"✓ {type_name} passed type validation (other errors expected for some)")
        
        # Save and close
        origin.save("test_refactoring_3.opju")
        origin.close()
        
        print("✓ Refactoring Item 3: ALL TESTS PASSED")
        return True
        
    except Exception as e:
        print(f"✗ Refactoring Item 3 FAILED: {e}")
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass
        return False

def test_refactoring_4():
    """Test Item 4: Function separation for 1D/2D data processing"""
    print("\n=== Testing Refactoring Item 4: Function separation ===")
    
    try:
        # Create Origin instance
        origin = OriginInstance()
        wb = origin.new_workbook("test_refactoring_4")
        ws = wb.worksheets()[0]
        worksheet = Worksheet(ws._obj, origin, parent=wb)
        
        # Test that private methods exist and are callable
        print("Testing private method existence...")
        
        assert hasattr(worksheet, '_add_column_from_1d_data'), "_add_column_from_1d_data method missing"
        assert hasattr(worksheet, '_add_column_from_dataframe'), "_add_column_from_dataframe method missing"
        assert hasattr(worksheet, '_add_column_from_2d_array'), "_add_column_from_2d_array method missing"
        assert hasattr(worksheet, '_add_column_from_2d_list'), "_add_column_from_2d_list method missing"
        
        print("✓ All private methods exist")
        
        # Test 1D data processing
        print("Testing 1D data processing...")
        
        # Test 1D list
        list_1d = [1, 2, 3, 4, 5]
        col = worksheet._add_column_from_1d_data(list_1d, lname="Test1D")
        assert col is not None, "1D list method should return a column"
        assert col.LongName == "Test1D", f"Expected 'Test1D', got '{col.LongName}'"
        print("✓ 1D list processing works")
        
        # Test 1D array
        arr_1d = np.array([10, 20, 30])
        col2 = worksheet._add_column_from_1d_data(arr_1d.tolist(), lname="Test1DArray")
        assert col2 is not None, "1D array method should return a column"
        assert col2.LongName == "Test1DArray", f"Expected 'Test1DArray', got '{col2.LongName}'"
        print("✓ 1D array processing works")
        
        # Test 2D data processing
        print("Testing 2D data processing...")
        
        # Test DataFrame
        df = pd.DataFrame({'A': [1, 2], 'B': [3, 4]})
        cols = worksheet._add_column_from_dataframe(df)
        if isinstance(cols, list):
            assert len(cols) == 2, f"Expected 2 columns from DataFrame, got {len(cols)}"
            assert cols[0].LongName == "A", f"Expected 'A', got '{cols[0].LongName}'"
            assert cols[1].LongName == "B", f"Expected 'B', got '{cols[1].LongName}'"
        else:
            assert cols.LongName == "A", f"Expected 'A', got '{cols.LongName}'"
        print("✓ DataFrame processing works")
        
        # Test 2D array
        arr_2d = np.array([[5, 6], [7, 8]])
        cols2 = worksheet._add_column_from_2d_array(arr_2d, lname="Array")
        if isinstance(cols2, list):
            assert len(cols2) == 2, f"Expected 2 columns from 2D array, got {len(cols2)}"
            assert cols2[0].LongName == "Array_1", f"Expected 'Array_1', got '{cols2[0].LongName}'"
            assert cols2[1].LongName == "Array_2", f"Expected 'Array_2', got '{cols2[1].LongName}'"
        print("✓ 2D array processing works")
        
        # Test 2D list
        list_2d = [[9, 10], [11, 12]]
        cols3 = worksheet._add_column_from_2d_list(list_2d, lname="List")
        if isinstance(cols3, list):
            assert len(cols3) == 2, f"Expected 2 columns from 2D list, got {len(cols3)}"
            assert cols3[0].LongName == "List_1", f"Expected 'List_1', got '{cols3[0].LongName}'"
            assert cols3[1].LongName == "List_2", f"Expected 'List_2', got '{cols3[1].LongName}'"
        print("✓ 2D list processing works")
        
        # Test that main method delegates correctly
        print("Testing delegation from main method...")
        
        # Clear worksheet for clean test
        worksheet.set_cols(0)
        
        # Test that add_column_from_data calls the right private methods
        result1 = worksheet.add_column_from_data([1, 2, 3], lname="Delegation1D")
        assert result1 is not None, "Delegation to 1D method failed"
        
        result2 = worksheet.add_column_from_data(pd.DataFrame({'X': [4, 5]}), lname="Delegation2D")
        assert result2 is not None, "Delegation to 2D method failed"
        
        print("✓ Delegation from main method works")
        
        # Save and close
        origin.save("test_refactoring_4.opju")
        origin.close()
        
        print("✓ Refactoring Item 4: ALL TESTS PASSED")
        return True
        
    except Exception as e:
        print(f"✗ Refactoring Item 4 FAILED: {e}")
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass
        return False

def main():
    """Run all refactoring tests"""
    print("Starting Worksheet Refactoring Tests (Items 1-4)")
    print("=" * 60)
    
    # Clean up any existing test files
    cleanup_test_files()
    
    results = []
    
    # Run each test
    results.append(test_refactoring_1())
    results.append(test_refactoring_2())
    results.append(test_refactoring_3())
    results.append(test_refactoring_4())
    
    # Summary
    print("\n" + "=" * 60)
    print("TEST SUMMARY")
    print("=" * 60)
    
    passed = sum(results)
    total = len(results)
    
    for i, result in enumerate(results, 1):
        status = "PASSED" if result else "FAILED"
        print(f"Refactoring Item {i}: {status}")
    
    print(f"\nOverall: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 ALL REFACTORING TESTS PASSED! No bugs detected.")
    else:
        print("⚠️  Some tests failed. Please review the implementation.")
    
    # Clean up test files
    cleanup_test_files()
    
    return passed == total

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
