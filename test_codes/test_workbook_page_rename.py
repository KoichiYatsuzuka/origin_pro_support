#!/usr/bin/env python3
"""
Test script to verify WorksheetPage was successfully renamed to WorkbookPage.
"""

import os
import sys
from pathlib import Path

# Add the parent directory to the path to import the library
sys.path.insert(0, str(Path(__file__).parent.parent))

import origin_pro_support as ops

def test_workbook_page_rename():
    """Test that WorksheetPage was renamed to WorkbookPage."""
    print("Testing WorksheetPage → WorkbookPage rename...")
    
    # Set up project path
    project_path = os.path.join(os.path.dirname(__file__), "test_workbook_page_rename.opju")
    
    # Delete existing file if it exists
    if os.path.exists(project_path):
        print(f"Removing existing file: {project_path}")
        os.remove(project_path)
    
    origin = None
    try:
        print("Creating Origin instance...")
        origin = ops.OriginInstance(project_path)
        origin.set_show(True)
        
        # Test 1: Check that WorkbookPage exists in the module
        print("\n1. Checking WorkbookPage class exists...")
        assert hasattr(ops, 'WorkbookPage'), "WorkbookPage class not found in module"
        print(f"   [OK] WorkbookPage found: {ops.WorkbookPage}")
        
        # Test 2: Check that WorksheetPage no longer exists
        print("\n2. Checking WorksheetPage class no longer exists...")
        assert not hasattr(ops, 'WorksheetPage'), "WorksheetPage class still exists (should be renamed)"
        print("   [OK] WorksheetPage correctly removed")
        
        # Test 3: Test creating a workbook and getting WorkbookPage
        print("\n3. Testing workbook creation returns WorkbookPage...")
        workbook = origin.new_workbook("TestWorkbook")
        assert isinstance(workbook, ops.WorkbookPage), f"Expected WorkbookPage, got {type(workbook)}"
        print(f"   [OK] Created workbook: {workbook}")
        print(f"   [OK] Type is correct: {type(workbook).__name__}")
        
        # Test 4: Test that workbook methods work correctly
        print("\n4. Testing WorkbookPage methods...")
        worksheets = list(workbook)
        print(f"   [OK] Can iterate over worksheets: {len(worksheets)} worksheets found")
        
        if worksheets:
            worksheet = worksheets[0]
            print(f"   [OK] First worksheet: {worksheet}")
            # Test that worksheet.get_page() returns WorkbookPage
            parent_page = worksheet.get_page()
            assert isinstance(parent_page, ops.WorkbookPage), f"Expected WorkbookPage, got {type(parent_page)}"
            print(f"   [OK] Worksheet parent is WorkbookPage: {type(parent_page).__name__}")
        
        print("\n" + "="*60)
        print("[SUCCESS] ALL TESTS PASSED - WorksheetPage successfully renamed to WorkbookPage!")
        print("="*60)
        
    except Exception as e:
        print(f"\n[ERROR] TEST FAILED: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        if origin:
            print("Closing Origin instance...")
            origin.close(False)

if __name__ == "__main__":
    test_workbook_page_rename()
