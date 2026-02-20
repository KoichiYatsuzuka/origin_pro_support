#!/usr/bin/env python3
"""
Test script to demonstrate the new header_rows functionality in Worksheet class.
This script shows how to use the header_rows parameter in the constructor and
the header_rows() method to control the display of column label rows.
"""

import sys
import os

# Add the parent directory to the path to import modules
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.insert(0, parent_dir)

try:
    from layers import Worksheet
    IMPORT_SUCCESS = True
except ImportError as e:
    print(f"Import failed (expected without Origin running): {e}")
    IMPORT_SUCCESS = False

def test_header_rows_functionality():
    """
    Test the header_rows functionality.
    
    Note: This is a demonstration of the API. In a real scenario, you would
    need an actual Origin instance and worksheet objects.
    """
    
    print("=== Worksheet header_rows Functionality Test ===")
    print()
    
    # Example 1: Worksheet constructor now automatically sets header rows with sparklines
    print("1. Worksheet constructor automatically sets header rows:")
    print("   worksheet = Worksheet(worksheet_obj)")
    print("   -> Automatically shows: Long Name, Units, Sparklines, F(x), Comments")
    print("   -> Uses header_rows('LUSCO') internally")
    print("   -> Automatically generates sparklines for numeric columns")
    print()
    
    # Example 2: Different header row specifications (using method directly)
    print("2. Manual header rows control using header_rows() method:")
    print("   worksheet.header_rows('L')      -> Show only Long Name")
    print("   worksheet.header_rows('U')      -> Show only Units") 
    print("   worksheet.header_rows('C')      -> Show only Comments")
    print("   worksheet.header_rows('LU')     -> Show Long Name + Units")
    print("   worksheet.header_rows('LUC')    -> Show Long Name + Units + Comments")
    print("   worksheet.header_rows('')       -> Hide all label rows")
    print()
    
    # Example 3: Default automatic setting
    print("3. Default automatic header rows setting:")
    print("   worksheet = Worksheet(worksheet_obj)")
    print("   -> Equivalent to: worksheet.header_rows('LUSCO')")
    print("   -> Shows: L=Long Name, U=Units, S=Sparklines, O=F(x), C=Comments")
    print()
    
    # Example 4: Column label row characters reference
    print("4. Column Label Row Characters (from Origin documentation):")
    label_rows = {
        'L': 'Long Name',
        'U': 'Units', 
        'C': 'Comments',
        'G': 'Short Name',
        'E': 'Sampling Interval',
        'F': 'Filter',
        'O': 'F(x)= (Formula)',
        'S': 'Sparklines',
        'T': 'Tick-indexed Dataset',
        'I': 'Indices',
        'M': 'Missing Values'
    }
    
    for char, description in label_rows.items():
        print(f"   '{char}': {description}")
    print()
    
    print("5. Usage with from_df method:")
    print("   worksheet = Worksheet(worksheet_obj)  # Auto-sets LUSCO")
    print("   df = pd.DataFrame({'Time': [1,2,3], 'Value': [10,20,30]})")
    print("   worksheet.from_df(df)  # Column names become Long Names")
    print("   -> Shows Long Name, Units, Sparklines, F(x), and Comments rows")
    print()
    
    print("6. Changing default display:")
    print("   worksheet = Worksheet(worksheet_obj)  # Shows LUSCO with sparklines")
    print("   worksheet.header_rows('LUC')          # Change to L-U-C only")
    print("   worksheet.header_rows('')             # Hide all")
    print()
    
    print("7. Sparklines-specific methods:")
    print("   worksheet.generate_sparklines()       # Generate sparklines for numeric columns")
    print("   worksheet.refresh_sparklines()        # Refresh sparklines after data changes")
    print("   worksheet._has_numeric_data(col_idx)  # Check if column has numeric data")
    print()
    
    print("8. Automatic sparklines features:")
    print("   - Generated automatically for numeric columns")
    print("   - Created when using from_df() method")
    print("   - Refreshed when data changes")
    print("   - Only shown for columns with numeric data")
    print()
    
    print("=== Test completed ===")
    print("Note: This is a demonstration. Actual testing requires Origin to be running.")

if __name__ == "__main__":
    test_header_rows_functionality()
