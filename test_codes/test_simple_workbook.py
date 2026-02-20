#!/usr/bin/env python3
"""
Simple test to debug workbook creation.
"""

import os
import sys
from pathlib import Path

# Add the parent directory to the path to import the library
sys.path.insert(0, str(Path(__file__).parent.parent))

from origin_pro_support import OriginInstance


def test_simple_workbook():
    """Test simple workbook creation."""
    print("Testing simple workbook creation...")
    
    # Set up project path
    project_path = os.path.join(os.path.dirname(__file__), "test_simple.opju")
    
    # Delete existing file if it exists
    if os.path.exists(project_path):
        print(f"Removing existing file: {project_path}")
        os.remove(project_path)
    
    origin = None
    try:
        print("Creating Origin instance...")
        origin = OriginInstance(project_path)
        origin.set_show(True)
        
        print("Getting root folder...")
        root_folder = origin.get_root_dir()
        print(f"Root folder path: {root_folder.Path}")
        
        print("Creating workbook...")
        workbook = root_folder.create_workbook("TestWorkbook")
        print(f"Workbook created: {workbook}")
        
        if workbook:
            print(f"Workbook name: {workbook.Name}")
            print(f"Workbook type: {type(workbook)}")
            print(f"Number of layers: {len(workbook)}")
            
            if len(workbook) > 0:
                worksheet = workbook[0]
                print(f"Worksheet: {worksheet}")
                print(f"Worksheet type: {type(worksheet)}")
                print(f"Worksheet dimensions: {worksheet.get_rows()} x {worksheet.get_cols()}")
        else:
            print("ERROR: Workbook creation returned None")
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        if origin:
            print("Closing Origin instance...")
            origin.close(False)  # Don't save


if __name__ == "__main__":
    test_simple_workbook()
