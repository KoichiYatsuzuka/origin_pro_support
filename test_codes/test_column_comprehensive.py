"""
Comprehensive test for Origin Pro Support column functionality.
Tests Workbook/Worksheet creation, DataFrame loading, and column long name handling.
"""

import pandas as pd
import origin_pro_support as ops
import os

def test_column_functionality():
    """Test complete column workflow with proper parent references."""
    
    print("=== Origin Pro Support Column Test ===\n")
    
    # Create test DataFrame with meaningful column names
    df = pd.DataFrame({
        'Time (seconds)': [0.0, 1.0, 2.0, 3.0, 4.0],
        'Temperature (C)': [20.5, 21.2, 22.8, 23.1, 24.3],
        'Pressure (kPa)': [101.3, 101.2, 101.4, 101.3, 101.5],
        'Humidity (%)': [45.2, 46.1, 47.5, 48.0, 47.8]
    })
    
    print("1. DataFrame created with columns:")
    for i, col in enumerate(df.columns):
        print(f"   {i}: '{col}'")
    print()
    
    # Setup test file
    test_file = os.path.join(os.getcwd(), "column_test.opju")
    if os.path.exists(test_file):
        os.remove(test_file)
    
    origin = None
    try:
        print("2. Creating Origin instance...")
        origin = ops.OriginInstance(test_file)
        
        print("3. Creating Workbook and Worksheet...")
        workbook = origin.new_book(long_name="Test Workbook")
        worksheet = workbook[0]
        print(f"   Workbook: {workbook.Name}")
        print(f"   Worksheet: {worksheet.Name}")
        
        print("4. Loading DataFrame data...")
        worksheet.from_df(df)
        print("   Data loaded with column long names")
        
        print("5. Testing column access methods:")
        
        # Test get_columns()
        print("   Testing get_columns()...")
        columns = worksheet.get_columns()
        print(f"   Retrieved {len(columns)} columns")
        
        # Test Columns property
        print("   Testing Columns property...")
        columns_prop = worksheet.Columns
        print(f"   Columns collection has {len(columns_prop)} columns")
        
        # Test iteration
        print("   Testing column iteration...")
        iterated_cols = list(worksheet)
        print(f"   Iterated {len(iterated_cols)} columns")
        
        # Test indexing
        print("   Testing column indexing...")
        indexed_col = worksheet[0]
        print(f"   Indexed column: {indexed_col.Name}")
        
        print("\n6. Verifying column properties:")
        all_correct = True
        for i, column in enumerate(columns):
            expected = str(df.columns[i])
            actual = column.LongName
            
            print(f"   Column {i}:")
            print(f"     Short name: '{column.Name}'")
            print(f"     Long name:  '{actual}'")
            print(f"     Expected:   '{expected}'")
            
            if actual == expected:
                print(f"     Status:     OK")
            else:
                print(f"     Status:     FAIL")
                all_correct = False
            
            # Test parent access
            try:
                parent = column.parent
                print(f"     Parent:     {parent.Name}")
            except Exception as e:
                print(f"     Parent:     ERROR - {e}")
                all_correct = False
            
            print()
        
        print("7. Testing data access:")
        for i, column in enumerate(columns):
            try:
                data = column.get_data(0)
                if data:
                    print(f"   Column '{column.LongName}': Data accessible ({len(data)} points)")
                else:
                    print(f"   Column '{column.LongName}': No data")
                    all_correct = False
            except Exception as e:
                print(f"   Column '{column.LongName}': Data access error - {e}")
                all_correct = False
        
        print("\n=== Test Results ===")
        if all_correct:
            print("SUCCESS: All tests passed!")
            print("- Column objects have proper parent references")
            print("- Long names are correctly set from DataFrame")
            print("- All column access methods work")
            print("- Data access through columns works")
        else:
            print("FAILURE: Some tests failed!")
        
        return all_correct
        
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        return False
        
    finally:
        if origin:
            try:
                origin.save()
                origin.close()
                print("\nOrigin file saved and closed")
            except:
                pass

if __name__ == "__main__":
    success = test_column_functionality()
    print(f"\nTest {'PASSED' if success else 'FAILED'}")
