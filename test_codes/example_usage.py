#!/usr/bin/env python3
"""
Example usage of the Worksheet class with automatic header rows display.

This example demonstrates how the Worksheet constructor now automatically
displays Long Name, Units, Sparklines, F(x), and Comments rows.
"""

# Example 1: Create Worksheet - header rows are set automatically
# worksheet = Worksheet(worksheet_obj)
# -> Automatically shows: Long Name, Units, Sparklines, F(x), Comments

# Example 2: Manual control after construction
# worksheet = Worksheet(worksheet_obj)  # Auto-sets LUSCO
# worksheet.header_rows('LUC')          # Change to show only L-U-C
# worksheet.header_rows('')             # Hide all

# Example 3: Similar to Origin Sample #3 behavior
# wks_2 = Worksheet(worksheet_obj)  # Shows LUSCO automatically
# df = pd.DataFrame({'Time': [1,2,3,4,5], 'Dataset 1': [0, 100.3, 210.7, 343, 4.56]})
# wks_2.from_df(df)

# Example 4: Different header row configurations
# worksheet.header_rows('L')      # Show only Long Name
# worksheet.header_rows('LU')     # Show Long Name + Units
# worksheet.header_rows('LUC')    # Show Long Name + Units + Comments
# worksheet.header_rows('LUSCO')  # Show Long Name + Units + Sparklines + F(x) + Comments
# worksheet.header_rows('')       # Hide all label rows

# Example 5: Sparklines-specific operations
# worksheet = Worksheet(worksheet_obj)  # Auto-generates sparklines
# worksheet.generate_sparklines()       # Manual generation
# worksheet.refresh_sparklines()        # Refresh after data changes
# has_numeric = worksheet._has_numeric_data(0)  # Check column 0

# Example 6: Loading data with automatic sparklines
# worksheet = Worksheet(worksheet_obj)
# df = pd.DataFrame({'X': [1,2,3,4,5], 'Y': [10,20,15,25,30]})
# worksheet.from_df(df)  # Automatically creates sparklines for X and Y columns

print("Worksheet automatic header rows with sparklines functionality has been implemented!")
print()
print("Key features:")
print("1. Worksheet constructor automatically sets header_rows('LUSCO')")
print("2. Shows: Long Name, Units, Sparklines, F(x), Comments")
print("3. Automatically generates sparklines for numeric columns")
print("4. Can still be changed manually using header_rows() method")
print("5. Compatible with existing Worksheet methods like from_df()")
print()
print("Default behavior:")
print("- L: Long Name")
print("- U: Units") 
print("- S: Sparklines (automatically generated for numeric data)")
print("- O: F(x)= (Formula)")
print("- C: Comments")
print()
print("Sparklines methods:")
print("- worksheet.generate_sparklines()  # Manual sparklines generation")
print("- worksheet.refresh_sparklines()   # Refresh after data changes")
print("- worksheet._has_numeric_data()    # Check for numeric data")
print()
print("See test_codes/test_sparklines.py for detailed sparklines testing.")
