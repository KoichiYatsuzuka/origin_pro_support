"""
Debug LabTalk command execution and page detection.
"""
import sys
import os
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

def debug_labtalk():
    """Debug LabTalk command execution"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        
        # Create instance
        current_dir = os.getcwd()
        test_path = os.path.join(current_dir, "debug_labtalk.opju")
        origin = OriginInstance(test_path)
        
        print("=== Testing direct LabTalk execution ===")
        
        # Test direct LabTalk command
        try:
            cmd = 'newbook name:="debug_wb"'
            print(f"Executing: {cmd}")
            origin.__core.LT_execute(cmd)
            print("LabTalk command executed successfully")
        except Exception as e:
            print(f"LabTalk command failed: {e}")
        
        # Wait
        import time
        time.sleep(0.5)
        
        # Check pages
        print("=== Checking pages ===")
        try:
            pages = origin.__core.GetWorksheetPages()
            print(f"GetWorksheetPages() returned {len(pages)} pages:")
            
            for i, page in enumerate(pages):
                print(f"  Page {i}: Name='{page.Name}', LongName='{page.LongName}'")
                
            # Try to find our page
            for page in pages:
                if page.Name == "debug_wb" or page.LongName == "debug_wb":
                    print(f"Found debug_wb page!")
                    from origin_pro_support.pages import WorkbookPage
                    wb = WorkbookPage(page._obj)
                    print(f"WorkbookPage created: {wb}")
                    break
            else:
                print("debug_wb page not found")
                
        except Exception as e:
            print(f"GetWorksheetPages failed: {e}")
        
        # Try PageBases from root folder
        print("=== Checking PageBases from root ===")
        try:
            root_folder = origin.__core.GetRootFolder()
            pages = list(root_folder.PageBases())
            print(f"PageBases() returned {len(pages)} pages:")
            
            for i, page in enumerate(pages):
                print(f"  Page {i}: Name='{page.Name}', LongName='{page.LongName}', Type={type(page)}")
                
        except Exception as e:
            print(f"PageBases failed: {e}")
        
        origin.close()
        
    except Exception as e:
        print(f"ERROR: {e}")
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass

if __name__ == "__main__":
    debug_labtalk()
