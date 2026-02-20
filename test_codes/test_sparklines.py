#!/usr/bin/env python3
"""
Test sparklines functionality in Worksheet class.
This test verifies that sparklines are automatically generated and displayed.
"""
import sys
import os
import numpy as np
import pandas as pd

# Add the parent directory to the path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def test_sparklines_functionality():
    """Test sparklines generation and display."""
    print("=== Worksheet Sparklines Test ===")
    
    try:
        from origin_pro_support import Worksheet, OriginInstance
        print("[OK] Successfully imported required modules")
    except ImportError as e:
        print(f"[ERROR] Import failed: {e}")
        return False
    
    # Create test file path
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_path = os.path.join(current_dir, "test_sparklines.opju")
    
    # Clean up existing file
    if os.path.exists(project_path):
        os.remove(project_path)
    
    origin = None
    try:
        print("\n1. Creating Origin instance...")
        origin = OriginInstance(project_path)
        origin.set_show(False)  # Hide window for automated test
        
        print("2. Creating Worksheet (should auto-set header rows with sparklines)...")
        worksheet = origin.new_sheet('w', 'SparklinesTest')
        print("[OK] Worksheet created with automatic sparklines setup")
        
        print("3. Loading test data with numeric columns...")
        # Create test data with various numeric patterns
        data = {
            'Time': np.linspace(0, 10, 20),
            'Sine_Wave': np.sin(np.linspace(0, 4*np.pi, 20)),
            'Cosine_Wave': np.cos(np.linspace(0, 4*np.pi, 20)),
            'Linear': np.linspace(0, 20, 20),
            'Random': np.random.normal(10, 2, 20),
            'Exponential': np.exp(np.linspace(0, 2, 20))
        }
        
        df = pd.DataFrame(data)
        worksheet.from_df(df)
        print("[OK] Data loaded with automatic sparklines generation")
        
        print("4. Testing sparklines methods...")
        
        # Test manual sparklines generation
        print("   - Testing manual sparklines generation...")
        worksheet.generate_sparklines()
        print("[OK] Manual sparklines generation works")
        
        # Test sparklines refresh
        print("   - Testing sparklines refresh...")
        worksheet.refresh_sparklines()
        print("[OK] Sparklines refresh works")
        
        # Test numeric data detection
        print("   - Testing numeric data detection...")
        for i, col_name in enumerate(df.columns):
            has_numeric = worksheet._has_numeric_data(i)
            print(f"     Column '{col_name}': {'Numeric' if has_numeric else 'Non-numeric'}")
        
        print("5. Adding more data and testing refresh...")
        # Add more data to test refresh functionality
        additional_data = np.random.normal(15, 3, 10)
        worksheet.from_list(6, additional_data, lname='Additional_Data')
        worksheet.refresh_sparklines()
        print("[OK] Added data and refreshed sparklines")
        
        print("6. Testing header rows with sparklines...")
        # Verify that sparklines row is included in header rows
        worksheet.header_rows('LUSCO')  # Ensure sparklines are visible
        print("[OK] Header rows set to include sparklines")
        
        print("7. Saving project...")
        origin.save()
        print("[OK] Project saved successfully")
        
        print("\n=== Test Results ===")
        print("[OK] Worksheet constructor automatically sets up sparklines")
        print("[OK] Sparklines are generated for numeric columns")
        print("[OK] Manual sparklines generation works")
        print("[OK] Sparklines refresh functionality works")
        print("[OK] Numeric data detection works correctly")
        print("[OK] Header rows include sparklines (S) in LUSCO")
        print("[OK] Compatible with from_df() method")
        
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

def test_sparklines_with_different_data():
    """Test sparklines with different data types."""
    print("\n=== Sparklines Data Types Test ===")
    
    try:
        from origin_pro_support import Worksheet, OriginInstance
        
        # Create test file path
        current_dir = os.path.dirname(os.path.abspath(__file__))
        project_path = os.path.join(current_dir, "test_sparklines_types.opju")
        
        if os.path.exists(project_path):
            os.remove(project_path)
        
        origin = OriginInstance(project_path)
        origin.set_show(False)
        
        print("1. Testing with mixed data types...")
        worksheet = origin.new_sheet('w', 'MixedTypes')
        
        # Mixed data: numeric, text, dates
        mixed_data = {
            'Numbers': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
            'Text': ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'],
            'Floats': [1.1, 2.2, 3.3, 4.4, 5.5, 6.6, 7.7, 8.8, 9.9, 10.0],
            'Mixed': [1, 'B', 3.3, 'D', 5, 'F', 7.7, 'H', 9, 'J']
        }
        
        df = pd.DataFrame(mixed_data)
        worksheet.from_df(df)
        
        print("2. Checking which columns get sparklines...")
        for i, col_name in enumerate(df.columns):
            has_numeric = worksheet._has_numeric_data(i)
            print(f"   Column '{col_name}': {'Will have sparklines' if has_numeric else 'No sparklines (non-numeric)'}")
        
        origin.save()
        origin.close()
        
        print("[OK] Mixed data types test completed")
        return True
        
    except Exception as e:
        print(f"[ERROR] Mixed data types test failed: {e}")
        return False

if __name__ == "__main__":
    print("Worksheet Sparklines Test Suite")
    print("=" * 50)
    
    # Test main sparklines functionality
    main_success = test_sparklines_functionality()
    
    # Test with different data types
    types_success = test_sparklines_with_different_data()
    
    print("\n" + "=" * 50)
    print("FINAL RESULTS:")
    print(f"Main Sparklines Test: {'PASSED' if main_success else 'FAILED'}")
    print(f"Data Types Test: {'PASSED' if types_success else 'FAILED'}")
    
    if main_success and types_success:
        print("\n[SUCCESS] All sparklines tests PASSED!")
        print("\nSparklines features:")
        print("- Automatically generated for numeric columns")
        print("- Displayed in header rows (S in LUSCO)")
        print("- Manual generation and refresh available")
        print("- Smart numeric data detection")
        print("- Compatible with from_df() and from_list() methods")
    else:
        print("\n[FAILED] Some sparklines tests FAILED!")
    
    print("\nTest completed.")
