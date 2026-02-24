"""
Test root folder workbook creation without directory movement.
"""
import sys
import os
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

def test_root_folder_simple():
    """Test simple workbook creation in root folder"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        
        # Create instance
        current_dir = Path.cwd()
        test_path = current_dir / 'test_root_simple.opju'
        origin = OriginInstance(str(test_path))
        print('SUCCESS: OriginInstance created')
        
        # Get root folder
        root_folder = origin.get_root_dir()
        print(f'Root folder path: {root_folder.get_path()}')
        
        # Test simple newbook command without cd
        print('\n=== Testing simple newbook command ===')
        cmd = f'newbook name:="simple_test"'
        print(f'Executing: {cmd}')
        
        # Execute the command directly on root folder
        root_folder._folder.Execute(cmd)
        print('Command executed')
        
        # Wait and check
        import time
        time.sleep(1.0)
        
        # Check if workbook was created
        print('\n=== Checking results ===')
        try:
            # Check all workbooks in project
            app = root_folder._folder.GetApplication()
            all_workbooks = list(app.GetWorksheetPages())
            print(f'Total workbooks in project: {len(all_workbooks)}')
            
            for i, wb in enumerate(all_workbooks):
                print(f'  Workbook {i}: Name="{wb.Name}", LongName="{wb.LongName}"')
                
                if wb.Name == 'simple_test' or wb.LongName == 'simple_test':
                    print(f'  -> FOUND matching workbook!')
                    from origin_pro_support.pages import WorkbookPage
                    workbook_obj = WorkbookPage(wb._obj)
                    print(f'  -> WorkbookPage object created: {workbook_obj}')
                    
                    # Test worksheets
                    worksheets = workbook_obj.worksheets()
                    print(f'  -> Worksheets: {worksheets}')
                    
                    origin.close()
                    return True
            
            print('  -> No matching workbook found')
            
        except Exception as e:
            print(f'Error checking workbooks: {e}')
        
        # Also check folder pages
        try:
            pages = list(root_folder._folder.PageBases())
            print(f'Pages in root folder: {len(pages)}')
            for i, page in enumerate(pages):
                print(f'  Page {i}: Name="{page.Name}", LongName="{page.LongName}"')
        except Exception as e:
            print(f'Error checking folder pages: {e}')
        
        origin.close()
        return False
        
    except Exception as e:
        print(f'ERROR: {e}')
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass
        return False

def test_with_origininstance_method():
    """Test using OriginInstance.new_workbook method"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        
        # Create instance
        current_dir = Path.cwd()
        test_path = current_dir / 'test_origin_method.opju'
        origin = OriginInstance(str(test_path))
        print('SUCCESS: OriginInstance created')
        
        # Test OriginInstance.new_workbook method
        print('\n=== Testing OriginInstance.new_workbook ===')
        wb = origin.new_workbook('origin_method_test')
        print(f'OriginInstance.new_workbook returned: {wb}')
        
        if wb is not None:
            print('SUCCESS: Workbook created via OriginInstance!')
            worksheets = wb.worksheets()
            print(f'Worksheets: {worksheets}')
            origin.close()
            return True
        else:
            print('FAILED: OriginInstance.new_workbook returned None')
            origin.close()
            return False
        
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
    print("=== Root Folder Simple Workbook Creation Test ===")
    print("=" * 60)
    
    results = []
    
    # Test 1: Simple command execution
    results.append(test_root_folder_simple())
    
    print("\n" + "=" * 60)
    
    # Test 2: OriginInstance method
    results.append(test_with_origininstance_method())
    
    # Summary
    print("\n" + "=" * 60)
    print("TEST SUMMARY")
    print("=" * 60)
    
    test_names = [
        "Simple Command Execution",
        "OriginInstance Method"
    ]
    
    for i, (name, result) in enumerate(zip(test_names, results)):
        status = "PASSED" if result else "FAILED"
        print(f"{name}: {status}")
    
    passed = sum(results)
    total = len(results)
    print(f"\nOverall: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 ROOT FOLDER WORKBOOK CREATION WORKS!")
    else:
        print("⚠️  Some tests failed")
    
    return passed == total

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
