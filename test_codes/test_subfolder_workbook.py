"""
Test subfolder creation and workbook creation in subfolders.
"""
import sys
import os
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

def test_subfolder_creation():
    """Test creating subfolder and workbook in subfolder"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        
        # Create instance
        current_dir = Path.cwd()
        import time
        timestamp = int(time.time())
        test_path = current_dir / f'test_subfolder_{timestamp}.opju'
        origin = OriginInstance(str(test_path))
        print('SUCCESS: OriginInstance created')
        
        # Get root folder
        root_folder = origin.get_root_dir()
        print(f'Root folder path: {root_folder.get_path()}')
        
        # Create subfolder
        print('\n=== Creating subfolder ===')
        subfolder = root_folder.create_folder('TestSubFolder')
        print(f'Subfolder created: {subfolder.get_path()}')
        
        # Create workbook in subfolder
        print('\n=== Creating workbook in subfolder ===')
        wb = subfolder.new_workbook('subfolder_workbook')
        print(f'Subfolder workbook: {wb}')
        
        if wb is not None:
            print('SUCCESS: Workbook created in subfolder!')
        else:
            print('FAILED: Subfolder workbook is None')
        
        # Create another workbook in root folder to test cd back
        print('\n=== Creating workbook in root folder (after subfolder) ===')
        wb2 = root_folder.new_workbook('root_after_subfolder')
        print(f'Root workbook: {wb2}')
        
        if wb2 is not None:
            print('SUCCESS: Workbook created in root after subfolder!')
        else:
            print('FAILED: Root workbook after subfolder is None')
        
        origin.close()
        return wb is not None and wb2 is not None
        
    except Exception as e:
        print(f'ERROR: {e}')
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass
        return False

def test_multiple_subfolders():
    """Test multiple subfolders with workbooks"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        
        # Create instance
        current_dir = Path.cwd()
        import time
        timestamp = int(time.time())
        test_path = current_dir / f'test_multiple_{timestamp}.opju'
        origin = OriginInstance(str(test_path))
        print('SUCCESS: OriginInstance created')
        
        # Get root folder
        root_folder = origin.get_root_dir()
        
        # Create multiple subfolders
        print('\n=== Creating multiple subfolders ===')
        sub1 = root_folder.create_folder('SubFolder1')
        sub2 = root_folder.create_folder('SubFolder2')
        
        print(f'Subfolder1 path: {sub1.get_path()}')
        print(f'Subfolder2 path: {sub2.get_path()}')
        
        # Create workbooks in each subfolder
        print('\n=== Creating workbooks in subfolders ===')
        wb1 = sub1.new_workbook('workbook1')
        wb2 = sub2.new_workbook('workbook2')
        
        print(f'Workbook1 in SubFolder1: {wb1}')
        print(f'Workbook2 in SubFolder2: {wb2}')
        
        # Test root folder still works
        print('\n=== Testing root folder still works ===')
        wb_root = root_folder.new_workbook('root_final')
        print(f'Root final workbook: {wb_root}')
        
        success = (wb1 is not None and wb2 is not None and wb_root is not None)
        
        origin.close()
        return success
        
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
    print("=== Subfolder Workbook Test ===")
    print("=" * 45)
    
    results = []
    
    # Test 1: Basic subfolder creation and workbook
    results.append(test_subfolder_creation())
    
    print("\n" + "=" * 45)
    
    # Test 2: Multiple subfolders
    results.append(test_multiple_subfolders())
    
    # Summary
    print("\n" + "=" * 45)
    print("TEST SUMMARY")
    print("=" * 45)
    
    test_names = [
        "Basic subfolder test",
        "Multiple subfolders test"
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
