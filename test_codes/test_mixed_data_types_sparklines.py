#!/usr/bin/env python3
"""
Test sparklines generation with mixed data types (integers, floats, strings).
This test creates a worksheet with various data types and attempts sparklines generation.
"""
import sys
import os
import numpy as np
import pandas as pd

# Add the parent directory to the path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def test_mixed_data_types_sparklines():
    """Test sparklines with mixed data types: integers, floats, strings."""
    print("=== Mixed Data Types Sparklines Test ===")
    
    try:
        from origin_pro_support import Worksheet, OriginInstance
        print("[OK] Successfully imported required modules")
    except ImportError as e:
        print(f"[ERROR] Import failed: {e}")
        return False
    
    # Create test file path
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_path = os.path.join(current_dir, "test_mixed_data_types.opju")
    
    # Clean up existing file
    if os.path.exists(project_path):
        os.remove(project_path)
    
    origin = None
    try:
        print("\n1. Creating Origin instance...")
        origin = OriginInstance(project_path)
        origin.set_show(False)  # Hide window for automated test
        
        print("2. Creating Worksheet...")
        worksheet = origin.new_sheet('w', 'MixedDataTypes')
        
        print("3. Creating mixed data types (integers, floats, strings)...")
        # Create comprehensive test data with various data types
        mixed_data = {
            # Integer columns
            'Positive_Integers': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
            'Negative_Integers': [-1, -2, -3, -4, -5, -6, -7, -8, -9, -10],
            'Zero_Integers': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            'Mixed_Integers': [10, -5, 0, 15, -10, 20, -15, 25, -20, 30],
            
            # Float columns
            'Positive_Floats': [1.1, 2.2, 3.3, 4.4, 5.5, 6.6, 7.7, 8.8, 9.9, 10.0],
            'Negative_Floats': [-1.1, -2.2, -3.3, -4.4, -5.5, -6.6, -7.7, -8.8, -9.9, -10.0],
            'Small_Floats': [0.001, 0.002, 0.003, 0.004, 0.005, 0.006, 0.007, 0.008, 0.009, 0.010],
            'Large_Floats': [1000.5, 2000.5, 3000.5, 4000.5, 5000.5, 6000.5, 7000.5, 8000.5, 9000.5, 10000.5],
            
            # Scientific notation
            'Scientific_Positive': [1.23e-3, 2.34e-3, 3.45e-3, 4.56e-3, 5.67e-3, 6.78e-3, 7.89e-3, 8.90e-3, 9.01e-3, 1.12e-2],
            'Scientific_Negative': [-1.23e3, -2.34e3, -3.45e3, -4.56e3, -5.67e3, -6.78e3, -7.89e3, -8.90e3, -9.01e3, -1.12e4],
            
            # String columns
            'Short_Strings': ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'],
            'Medium_Strings': ['Apple', 'Banana', 'Cherry', 'Date', 'Elderberry', 'Fig', 'Grape', 'Honeydew', 'Kiwi', 'Lemon'],
            'Long_Strings': ['Very long string one', 'Very long string two', 'Very long string three', 'Very long string four', 'Very long string five', 'Very long string six', 'Very long string seven', 'Very long string eight', 'Very long string nine', 'Very long string ten'],
            'Numeric_Strings': ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'],
            'Mixed_Strings': ['1', 'Two', '3.3', 'Four', '5', 'Six', '7.7', 'Eight', '9', 'Ten'],
            
            # Special character strings
            'Special_Chars': ['!@#', '$%^', '&*()', '[]{}', '<>/?', '|\\', '~`', '+-=', '""', "''"],
            'Unicode_Chars': ['α', 'β', 'γ', 'δ', 'ε', 'ζ', 'η', 'θ', 'ι', 'κ'],
            
            # Mixed data types in same column
            'Mixed_Column': [1, 'Two', 3.3, 'Four', 5, 'Six', 7.7, 'Eight', 9, 'Ten'],
            
            # Edge cases
            'Empty_Values': [None, '', None, '', None, '', None, '', None, ''],
            'Boolean_Values': [True, False, True, False, True, False, True, False, True, False]
        }
        
        print(f"   Created {len(mixed_data)} columns with different data types")
        
        # Load data into worksheet
        df = pd.DataFrame(mixed_data)
        worksheet.from_df(df)
        print("[OK] Mixed data loaded successfully")
        
        print("\n4. Attempting sparklines generation for ALL mixed data types...")
        print("   This will test sparklines generation for each data type:")
        worksheet.generate_sparklines()
        
        print("\n5. Testing individual column groups by data type...")
        
        # Test integer columns
        print("\n   Testing INTEGER columns (0-3):")
        worksheet.generate_sparklines(0, 3)
        
        # Test float columns  
        print("\n   Testing FLOAT columns (4-9):")
        worksheet.generate_sparklines(4, 9)
        
        # Test string columns
        print("\n   Testing STRING columns (10-16):")
        worksheet.generate_sparklines(10, 16)
        
        # Test mixed and edge case columns
        print("\n   Testing MIXED and EDGE CASE columns (17-19):")
        worksheet.generate_sparklines(17, 19)
        
        print("\n6. Testing refresh functionality...")
        worksheet.refresh_sparklines()
        
        print("\n7. Adding more mixed data and testing...")
        # Add additional mixed data columns
        additional_data = {
            'More_Integers': [100, 200, 300, 400, 500],
            'More_Floats': [100.1, 200.2, 300.3, 400.4, 500.5],
            'More_Strings': ['Hundred', 'Two Hundred', 'Three Hundred', 'Four Hundred', 'Five Hundred']
        }
        
        for i, (col_name, data) in enumerate(additional_data.items()):
            col_index = 20 + i
            worksheet.from_list(col_index, data, lname=col_name)
        
        print("   Testing additional mixed columns (20-22):")
        worksheet.generate_sparklines(20, 22)
        
        print("\n8. Saving project...")
        origin.save()
        print("[OK] Project saved successfully")
        
        print("\n" + "="*60)
        print("MIXED DATA TYPES TEST RESULTS:")
        print("="*60)
        print("[OK] Integer columns: Various integer values tested")
        print("[OK] Float columns: Positive, negative, small, large floats tested")
        print("[OK] Scientific notation: Both positive and negative tested")
        print("[OK] String columns: Short, medium, long strings tested")
        print("[OK] Numeric strings: String representations of numbers tested")
        print("[OK] Mixed strings: Combination of numbers and text tested")
        print("[OK] Special characters: Various symbols tested")
        print("[OK] Unicode characters: Greek letters tested")
        print("[OK] Mixed columns: Multiple data types in same column tested")
        print("[OK] Edge cases: Empty values and boolean values tested")
        print("[OK] Additional data: Dynamic column addition tested")
        print("[OK] Refresh functionality: Sparklines refresh tested")
        
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

def analyze_sparklines_results():
    """Analyze and summarize the sparklines test results."""
    print("\n" + "="*60)
    print("SPARKLINES ANALYSIS:")
    print("="*60)
    print("Based on the test results above, observe the following:")
    print()
    print("1. SUCCESSFUL COLUMNS:")
    print("   - Look for '[OK] Sparklines generated' messages")
    print("   - Note which data types successfully generate sparklines")
    print()
    print("2. FAILED COLUMNS:")
    print("   - Look for '[WARNING] Could not generate sparklines' messages")
    print("   - Note which data types fail and the error messages")
    print()
    print("3. ORIGIN'S CAPABILITIES:")
    print("   - Origin's sparklines function may handle more data types than expected")
    print("   - Some string data might be convertible to numeric for sparklines")
    print("   - Error messages provide insight into Origin's limitations")
    print()
    print("4. PRACTICAL IMPLICATIONS:")
    print("   - The new approach attempts ALL columns, maximizing opportunities")
    print("   - Users get clear feedback on what works and what doesn't")
    print("   - Process continues even if some columns fail")

if __name__ == "__main__":
    print("Mixed Data Types Sparklines Test")
    print("=" * 60)
    print("This test creates a comprehensive worksheet with:")
    print("- Integers (positive, negative, zero, mixed)")
    print("- Floats (positive, negative, small, large, scientific)")
    print("- Strings (short, medium, long, numeric, mixed)")
    print("- Special characters and Unicode")
    print("- Mixed data types and edge cases")
    print("=" * 60)
    
    # Run the main test
    success = test_mixed_data_types_sparklines()
    
    # Analyze results
    analyze_sparklines_results()
    
    print("\n" + "=" * 60)
    print("FINAL RESULT:")
    if success:
        print("[SUCCESS] MIXED DATA TYPES TEST: PASSED")
        print("\nThe test successfully demonstrated:")
        print("- Sparklines generation attempted for ALL data types")
        print("- Clear success/failure messaging for each column")
        print("- Robust error handling and process continuation")
        print("- Comprehensive coverage of different data scenarios")
    else:
        print("[FAILED] MIXED DATA TYPES TEST: FAILED")
    
    print("\nTest completed.")
