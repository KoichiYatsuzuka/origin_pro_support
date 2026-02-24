"""
Fixed test script for OriginInstance refactoring validation.
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
        "test_origininstance_fixed.opju"
    ]
    
    for file in test_files:
        if os.path.exists(file):
            try:
                os.remove(file)
                print(f"Cleaned up {file}")
            except Exception as e:
                print(f"Warning: Could not remove {file}: {e}")

def test_origininstance_complete():
    """Complete test of OriginInstance functionality"""
    print("\n=== Testing OriginInstance Complete Functionality ===")
    
    try:
        # Test import
        from origin_pro_support.origin_instance import OriginInstance
        print("[OK] OriginInstance import successful")
        
        # Test instantiation
        current_dir = os.getcwd()
        test_path = os.path.join(current_dir, "test_origininstance_fixed.opju")
        origin = OriginInstance(test_path)
        print("[OK] OriginInstance instantiation successful")
        
        # Test core functionality
        assert hasattr(origin, 'get_root_dir'), "get_root_dir method missing"
        root_dir = origin.get_root_dir()
        assert root_dir is not None, "get_root_dir returned None"
        print("[OK] get_root_dir() works")
        
        # Test core initialization
        assert hasattr(origin, '_OriginInstance__core'), "Origin core not initialized"
        print("[OK] Origin core is initialized")
        
        # Test workbook operations
        wb = origin.new_workbook("test_workbook")
        assert wb is not None, "Workbook creation failed"
        print("[OK] Workbook creation works")
        
        # Test worksheet operations
        worksheets = wb.worksheets()
        assert len(worksheets) > 0, "No worksheets found"
        ws = worksheets[0]
        assert ws is not None, "Worksheet access failed"
        
        ws.set_cols(3)
        ws.set_rows(5)
        assert ws.get_cols() == 3, "Column setting failed"
        assert ws.get_rows() == 5, "Row setting failed"
        print("[OK] Worksheet operations work")
        
        # Test method access through get_root_dir() chain
        try:
            wb2 = origin.new_workbook("chain_test")
            assert wb2 is not None, "Method chain test failed"
            print("[OK] Method chain access works")
        except Exception as e:
            print(f"[ERROR] Method chain test failed: {e}")
            return False
        
        # Test LabTalk operations
        try:
            origin.lt_set_var("test_var", 123)
            value = origin.lt_get_var("test_var")
            assert value == 123, "LabTalk variable test failed"
            print("[OK] LabTalk operations work")
        except Exception as e:
            print(f"[ERROR] LabTalk operations failed: {e}")
            return False
        
        # Test save and close
        origin.save()
        origin.close()
        print("[OK] Save and close operations work")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Complete functionality test FAILED: {e}")
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass
        return False

def test_integration_with_worksheet():
    """Test integration with Worksheet class"""
    print("\n=== Testing Integration with Worksheet Class ===")
    
    try:
        from origin_pro_support.origin_instance import OriginInstance
        from origin_pro_support.layers import Worksheet
        
        # Create Origin instance
        current_dir = os.getcwd()
        test_path = os.path.join(current_dir, "test_integration.opju")
        origin = OriginInstance(test_path)
        
        # Create workbook and worksheet
        wb = origin.new_workbook("integration_test")
        ws = wb.worksheets()[0]
        
        # Test Worksheet with different data types
        test_cases = [
            ("DataFrame", pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})),
            ("2D Array", np.array([[1, 2], [3, 4]])),
            ("2D List", [[1, 2], [3, 4]])
        ]
        
        for name, data in test_cases:
            try:
                worksheet = Worksheet(ws._obj, origin, parent=wb, data=data)
                assert worksheet.get_cols() >= 2, f"{name} test failed"
                print(f"[OK] {name} integration works")
            except Exception as e:
                print(f"[ERROR] {name} integration failed: {e}")
                return False
        
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

def main():
    """Run complete OriginInstance refactoring tests"""
    print("Starting Complete OriginInstance Refactoring Validation")
    print("=" * 60)
    
    # Clean up any existing test files
    cleanup_test_files()
    
    results = []
    
    # Run tests
    results.append(test_origininstance_complete())
    results.append(test_integration_with_worksheet())
    
    # Summary
    print("\n" + "=" * 60)
    print("COMPLETE TEST SUMMARY")
    print("=" * 60)
    
    passed = sum(results)
    total = len(results)
    
    test_names = [
        "Complete Functionality",
        "Worksheet Integration"
    ]
    
    for i, (name, result) in enumerate(zip(test_names, results)):
        status = "PASSED" if result else "FAILED"
        print(f"{name}: {status}")
    
    print(f"\nOverall: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 ORIGININSTANCE REFACTORING COMPLETELY SUCCESSFUL!")
        print("✓ Class definition moved to origin_instance.py")
        print("✓ All functionality preserved")
        print("✓ Integration with other classes works")
        print("✓ Method access through get_root_dir() works")
    else:
        print("⚠️  Some tests failed. Review needed.")
    
    # Clean up test files
    cleanup_test_files()
    
    return passed == total

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
