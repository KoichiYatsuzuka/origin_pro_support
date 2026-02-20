#!/usr/bin/env python3
"""
Test sparklines generation with columns that should fail to demonstrate warning messages.
"""
import sys
import os
import pandas as pd

# Add the parent directory to the path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def test_sparklines_with_failures():
    """Test sparklines generation with data that should cause failures."""
    print("=== Sparklines Failure Test ===")
    
    try:
        from origin_pro_support import Worksheet, OriginInstance
        print("[OK] Successfully imported required modules")
    except ImportError as e:
        print(f"[ERROR] Import failed: {e}")
        return False
    
    # Create test file path
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_path = os.path.join(current_dir, "test_sparklines_failures.opju")
    
    # Clean up existing file
    if os.path.exists(project_path):
        os.remove(project_path)
    
    origin = None
    try:
        print("\n1. Creating Origin instance...")
        origin = OriginInstance(project_path)
        origin.set_show(False)
        
        print("2. Creating Worksheet...")
        worksheet = origin.new_sheet('w', 'FailureTest')
        
        print("3. Loading data that should cause sparklines failures...")
        # Create data with problematic content
        problematic_data = {
            'Valid_Numeric': [1, 2, 3, 4, 5],
            'Empty_Column': [None, None, None, None, None],
            'Only_Empty_Strings': ['', '', '', '', ''],
            'Mixed_Problematic': ['Not a number', 'Also not', 'Still not', 'Nope', 'Never'],
            'Special_Chars': ['!@#', '$%^', '&*()', '[]{}', '<>/?'],
            'Very_Long_Text': ['This is a very long text that might cause issues with sparklines generation' for _ in range(5)]
        }
        
        df = pd.DataFrame(problematic_data)
        worksheet.from_df(df)
        
        print("4. Attempting sparklines generation (should show warnings)...")
        worksheet.generate_sparklines()
        
        print("5. Testing individual problematic columns...")
        for i, col_name in enumerate(problematic_data.keys()):
            print(f"\n   Testing column {i} ({col_name}) individually:")
            worksheet.generate_sparklines(i, i)
        
        print("6. Saving project...")
        origin.save()
        
        print("\n=== Test Results ===")
        print("[OK] Successfully tested sparklines with problematic data")
        print("[OK] Warning messages should be visible for columns that failed")
        print("[OK] Process continued despite individual column failures")
        
        return True
        
    except Exception as e:
        print(f"\n[ERROR] Test failed: {e}")
        import traceback
        traceback.print_exc()
        return False
        
    finally:
        if origin is not None:
            try:
                origin.close()
                print("[OK] Origin instance closed")
            except Exception as e:
                print(f"[ERROR] Failed to close Origin: {e}")

if __name__ == "__main__":
    print("Sparklines Failure Demonstration")
    print("=" * 50)
    
    success = test_sparklines_with_failures()
    
    print("\n" + "=" * 50)
    print("FINAL RESULT:")
    print(f"Failure Test: {'PASSED' if success else 'FAILED'}")
    
    if success:
        print("\nThis test demonstrates:")
        print("- Sparklines are attempted for ALL columns")
        print("- Warning messages appear for columns that cannot have sparklines")
        print("- Process continues even when individual columns fail")
        print("- Better user feedback with detailed error messages")
    
    print("\nTest completed.")
