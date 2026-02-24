"""
Debug test for OriginInstance to identify the issue.
"""
import sys
import os
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

def debug_origininstance():
    """Debug OriginInstance to find the issue"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        print("SUCCESS: OriginInstance imported")
        
        # Create instance
        current_dir = os.getcwd()
        test_path = os.path.join(current_dir, "debug_test.opju")
        origin = OriginInstance(test_path)
        print("SUCCESS: OriginInstance created")
        
        # Test get_root_dir
        root_dir = origin.get_root_dir()
        print(f"SUCCESS: get_root_dir() returned: {type(root_dir)}")
        print(f"Root dir methods: {[m for m in dir(root_dir) if not m.startswith('_')]}")
        
        # Check if create_workbook exists
        if hasattr(root_dir, 'create_workbook'):
            print("SUCCESS: create_workbook method exists")
            
            # Try to create workbook
            try:
                wb = root_dir.create_workbook("debug_wb")
                print(f"SUCCESS: Workbook created: {wb}")
                print(f"Workbook type: {type(wb)}")
                
                if wb is not None:
                    print("SUCCESS: Workbook is not None")
                    # Test worksheets method
                    if hasattr(wb, 'worksheets'):
                        print("SUCCESS: worksheets method exists")
                        worksheets = wb.worksheets()
                        print(f"SUCCESS: worksheets() returned: {worksheets}")
                    else:
                        print("ERROR: worksheets method missing")
                else:
                    print("ERROR: Workbook is None")
                    
            except Exception as e:
                print(f"ERROR: create_workbook failed: {e}")
        else:
            print("ERROR: create_workbook method missing")
            print(f"Available methods: {[m for m in dir(root_dir) if 'create' in m.lower()]}")
        
        origin.close()
        
    except Exception as e:
        print(f"ERROR: {e}")
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass

if __name__ == "__main__":
    debug_origininstance()
