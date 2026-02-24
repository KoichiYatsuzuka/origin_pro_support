"""
Test current state of new_workbook method.
"""
import sys
import os
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

def test_current_state():
    """Test new_workbook method in current state"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        
        # Create instance
        current_dir = Path.cwd()
        test_path = current_dir / 'test_current_state.opju'
        origin = OriginInstance(str(test_path))
        print('SUCCESS: OriginInstance created')
        
        # Test new_workbook method
        print('\n=== Testing new_workbook method ===')
        wb = origin.new_workbook('current_test')
        print(f'new_workbook returned: {wb}')
        
        if wb is not None:
            print('SUCCESS: Workbook created!')
            try:
                worksheets = wb.worksheets()
                print(f'Worksheets: {worksheets}')
                
                # Test workbook methods
                print(f'Workbook name: {wb.get_name()}')
                print(f'Workbook long name: {wb.get_long_name()}')
                
            except Exception as e:
                print(f'Error testing workbook methods: {e}')
        else:
            print('FAILED: Workbook is None')
        
        origin.close()
        return wb is not None
        
    except Exception as e:
        print(f'ERROR: {e}')
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass
        return False

def test_folder_direct():
    """Test calling new_workbook directly on folder"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        
        # Create instance
        current_dir = Path.cwd()
        test_path = current_dir / 'test_folder_direct.opju'
        origin = OriginInstance(str(test_path))
        print('SUCCESS: OriginInstance created')
        
        # Get root folder and test directly
        print('\n=== Testing folder direct method ===')
        root_folder = origin.get_root_dir()
        print(f'Root folder path: {root_folder.get_path()}')
        
        wb = root_folder.new_workbook('folder_test')
        print(f'Direct new_workbook returned: {wb}')
        
        if wb is not None:
            print('SUCCESS: Direct workbook created!')
        else:
            print('FAILED: Direct workbook is None')
        
        origin.close()
        return wb is not None
        
    except Exception as e:
        print(f'ERROR: {e}')
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass
        return False

def main():
    """Run all tests"""
    print("=== Current State Test ===")
    print("=" * 40)
    
    results = []
    
    # Test 1: OriginInstance.new_workbook
    results.append(test_current_state())
    
    print("\n" + "=" * 40)
    
    # Test 2: Direct folder method
    results.append(test_folder_direct())
    
    # Summary
    print("\n" + "=" * 40)
    print("TEST SUMMARY")
    print("=" * 40)
    
    test_names = [
        "OriginInstance.new_workbook",
        "Direct folder.new_workbook"
    ]
    
    for i, (name, result) in enumerate(zip(test_names, results)):
        status = "PASSED" if result else "FAILED"
        print(f"{name}: {status}")
    
    passed = sum(results)
    total = len(results)
    print(f"\nOverall: {passed}/{total} tests passed")
    
    if passed == total:
        print("All tests PASSED!")
    else:
        print("Some tests FAILED")
    
    return passed == total

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
