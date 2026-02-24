"""
Debug create_workbook method specifically.
"""
import sys
import os
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

def debug_createworkbook():
    """Debug create_workbook method"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        from origin_pro_support.folder import Folder
        
        # Create instance
        current_dir = os.getcwd()
        test_path = os.path.join(current_dir, "debug_create.opju")
        origin = OriginInstance(test_path)
        
        root_dir = origin.get_root_dir()
        print(f"Root folder path: {root_dir.get_path()}")
        
        # Debug the create_workbook process
        name = "debug_wb"
        template = ""
        
        print(f"Creating workbook: {name}")
        
        # Check if page already exists
        if root_dir.has_page(name):
            print(f"ERROR: Page {name} already exists")
        else:
            print("Page doesn't exist, proceeding...")
        
        # Execute LabTalk command
        cmd = f'cd "{root_dir.get_path()}"; newbook name:="{name}" template:="{template}"'
        print(f"Executing command: {cmd}")
        
        try:
            root_dir._folder.Execute(cmd.strip())
            print("LabTalk command executed successfully")
        except Exception as e:
            print(f"LabTalk command failed: {e}")
        
        # Check for created pages
        print("Checking for created pages...")
        pages = list(root_dir._folder.PageBases())
        print(f"Found {len(pages)} pages:")
        
        for i, page in enumerate(pages):
            print(f"  Page {i}: Name='{page.Name}', LongName='{page.LongName}', Type={type(page)}")
            if page.Name == name or page.LongName == name:
                print(f"  -> Found matching page!")
                try:
                    from origin_pro_support.pages import WorkbookPage
                    wb = WorkbookPage(page._obj)
                    print(f"  -> WorkbookPage created: {wb}")
                    print(f"  -> WorkbookPage type: {type(wb)}")
                except Exception as e:
                    print(f"  -> WorkbookPage creation failed: {e}")
        
        # Try the actual method
        print("\nTrying actual create_workbook method...")
        wb = root_dir.create_workbook("test_method")
        print(f"create_workbook returned: {wb}")
        
        origin.close()
        
    except Exception as e:
        print(f"ERROR: {e}")
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass

if __name__ == "__main__":
    debug_createworkbook()
