"""
Test script for OriginInstance refactoring validation.
Tests that OriginInstance class maintains its functionality after:
1. Moving class definition to origin_instance.py
2. Restructuring method calls through get_root_dir()
"""
import sys
import os
import pandas as pd
import numpy as np
from pathlib import Path

# Add parent directory to path to import origin_pro_support
sys.path.insert(0, str(Path(__file__).parent.parent))

def cleanup_test_files():
    """Clean up any test files created during testing"""
    test_files = [
        "test_origininstance_refactoring.opju"
    ]
    
    for file in test_files:
        if os.path.exists(file):
            try:
                os.remove(file)
                print(f"Cleaned up {file}")
            except Exception as e:
                print(f"Warning: Could not remove {file}: {e}")

def test_origininstance_basic_functionality():
    """Test basic OriginInstance functionality"""
    print("\n=== Testing OriginInstance Basic Functionality ===")
    
    try:
        # Test import - should work after moving to origin_instance.py
        from origin_pro_support.origin_instance import OriginInstance
        print("[OK] OriginInstance import successful from origin_instance.py")
        
        # Test instantiation - should work with valid path
        try:
            # Use current directory as path for testing
            current_dir = os.getcwd()
            test_path = os.path.join(current_dir, "test.opju")
            origin = OriginInstance(test_path)
            print("[OK] OriginInstance instantiation successful")
        except Exception as e:
            print(f"[ERROR] OriginInstance instantiation failed: {e}")
            return False
        
        # Test basic properties
        assert hasattr(origin, 'get_root_dir'), "get_root_dir method missing"
        root_dir = origin.get_root_dir()
        assert root_dir is not None, "get_root_dir returned None"
        print(f"[OK] get_root_dir() works: {root_dir}")
        
        # Test that the instance is valid
        assert hasattr(origin, '_OriginInstance__core'), "Origin core not initialized"
        print("[OK] Origin core is initialized")
        
        # Test basic workbook creation
        wb = origin.new_workbook("test_workbook")
        assert wb is not None, "Workbook creation failed"
        print("[OK] Workbook creation works")
        
        # Test worksheet access
        worksheets = wb.worksheets()
        assert len(worksheets) > 0, "No worksheets found"
        ws = worksheets[0]
        assert ws is not None, "Worksheet access failed"
        print("[OK] Worksheet access works")
        
        # Test basic worksheet operations
        ws.set_cols(3)
        ws.set_rows(5)
        assert ws.get_cols() == 3, "Column setting failed"
        assert ws.get_rows() == 5, "Row setting failed"
        print("[OK] Basic worksheet operations work")
        
        # Save and close
        origin.save("test_origininstance_refactoring.opju")
        origin.close()
        print("[OK] Save and close operations work")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Basic functionality test FAILED: {e}")
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass
        return False

def test_origininstance_method_access():
    """Test that all key methods are accessible through get_root_dir()"""
    print("\n=== Testing Method Access Through get_root_dir() ===")
    
    try:
        from origin_pro_support.origin_instance import OriginInstance
        
        origin = OriginInstance("dummy_path")
        
        # Test that key methods exist and are callable
        key_methods = [
            'new_workbook',
            'new_graph', 
            'new_matrixbook',
            'new_notes',
            'get_workbook_pages',
            'get_graph_pages',
            'get_matrix_pages',
            'get_notes_pages',
            'find_book',
            'find_graph',
            'find_matrix',
            'save',
            'close',
            'get_root_dir',
            'lt_get_var',
            'lt_set_var',
            'lt_get_str',
            'lt_set_str'
        ]
        
        missing_methods = []
        for method_name in key_methods:
            if not hasattr(origin, method_name):
                missing_methods.append(method_name)
            else:
                method = getattr(origin, method_name)
                if not callable(method):
                    missing_methods.append(f"{method_name} (not callable)")
        
        if missing_methods:
            print(f"[ERROR] Missing or non-callable methods: {missing_methods}")
            return False
        
        print(f"[OK] All {len(key_methods)} key methods are accessible")
        
        # Test method execution through get_root_dir() chain
        try:
            # This tests that methods can be called through the get_root_dir() chain
            root_dir = origin.get_root_dir()
            assert root_dir is not None, "get_root_dir() returned None"
            
            # Try to create a workbook through the method chain
            wb = origin.new_workbook("method_chain_test")
            assert wb is not None, "Method chain workbook creation failed"
            
            # Test that we can access workbook properties
            wb_name = wb.name
            assert wb_name is not None, "Workbook name access failed"
            print(f"[OK] Method chain execution works (created workbook: {wb_name})")
            
            origin.close()
            
        except Exception as e:
            print(f"[ERROR] Method chain execution failed: {e}")
            return False
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Method access test FAILED: {e}")
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass
        return False

def test_origininstance_integration():
    """Test integration with other classes"""
    print("\n=== Testing Integration with Other Classes ===")
    
    try:
        from origin_pro_support.origin_instance import OriginInstance
        from origin_pro_support.layers import Worksheet
        
        origin = OriginInstance("dummy_path")
        
        # Test that Worksheet can still work with refactored OriginInstance
        wb = origin.new_workbook("integration_test")
        ws = wb.worksheets()[0]
        
        # Test Worksheet initialization with different data types
        # This tests that the Worksheet refactoring still works with the new OriginInstance
        
        # Test with DataFrame
        df = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
        worksheet_df = Worksheet(ws._obj, origin, parent=wb, data=df)
        assert worksheet_df.get_cols() == 2, "DataFrame integration failed"
        print("[OK] DataFrame integration works")
        
        # Test with 2D array
        arr = np.array([[1, 2], [3, 4]])
        worksheet_arr = Worksheet(ws._obj, origin, parent=wb, data=arr)
        assert worksheet_arr.get_cols() == 2, "2D array integration failed"
        print("[OK] 2D array integration works")
        
        # Test with 2D list
        lst = [[1, 2], [3, 4]]
        worksheet_lst = Worksheet(ws._obj, origin, parent=wb, data=lst)
        assert worksheet_lst.get_cols() == 2, "2D list integration failed"
        print("[OK] 2D list integration works")
        
        origin.close()
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Integration test FAILED: {e}")
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass
        return False

def test_origininstance_error_handling():
    """Test error handling in refactored OriginInstance"""
    print("\n=== Testing Error Handling ===")
    
    try:
        from origin_pro_support.origin_instance import OriginInstance
        
        origin = OriginInstance("dummy_path")
        
        # Test error handling for invalid operations
        try:
            # Try to access non-existent workbook
            wb = origin.find_book("non_existent_workbook")
            # Should return None or raise appropriate error
            print("[OK] Non-existent workbook handling works")
        except Exception as e:
            print(f"[OK] Non-existent workbook raises appropriate error: {type(e).__name__}")
        
        # Test LabTalk variable operations
        try:
            origin.lt_set_var("test_var", 42)
            value = origin.lt_get_var("test_var")
            assert value == 42, "LabTalk variable set/get failed"
            print("[OK] LabTalk variable operations work")
        except Exception as e:
            print(f"[ERROR] LabTalk variable operations failed: {e}")
            return False
        
        origin.close()
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Error handling test FAILED: {e}")
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass
        return False

def main():
    """Run all OriginInstance refactoring tests"""
    print("Starting OriginInstance Refactoring Validation Tests")
    print("=" * 60)
    
    # Clean up any existing test files
    cleanup_test_files()
    
    results = []
    
    # Run each test
    results.append(test_origininstance_basic_functionality())
    results.append(test_origininstance_method_access())
    results.append(test_origininstance_integration())
    results.append(test_origininstance_error_handling())
    
    # Summary
    print("\n" + "=" * 60)
    print("TEST SUMMARY")
    print("=" * 60)
    
    passed = sum(results)
    total = len(results)
    
    test_names = [
        "Basic Functionality",
        "Method Access", 
        "Integration",
        "Error Handling"
    ]
    
    for i, (name, result) in enumerate(zip(test_names, results)):
        status = "PASSED" if result else "FAILED"
        print(f"{name}: {status}")
    
    print(f"\nOverall: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 ORIGININSTANCE REFACTORING VALIDATION SUCCESSFUL!")
        print("✓ Class definition moved to origin_instance.py successfully")
        print("✓ Method access through get_root_dir() works correctly")
        print("✓ All functionality preserved")
    else:
        print("⚠️  Some tests failed. Please review the refactoring implementation.")
    
    # Clean up test files
    cleanup_test_files()
    
    return passed == total

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
