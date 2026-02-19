#%%
import pandas as pd
import origin_pro_support as ops
import os

# Create test data
df = pd.DataFrame({
    'Time': [1, 2, 3, 4, 5],
    'Values': [10, 20, 30, 40, 50]
})

# Create Origin instance
origin_file = os.path.join(os.getcwd(), "test_pythonic.opju")
origin = ops.OriginInstance(origin_file)

# Create a new workbook and get worksheet
book = origin.new_book()
wks = book[0]

# Load data
wks.from_df(df)

print("=== Testing Pythonic Naming Conventions ===")

# Get a column to test
col = wks[0]

# Test 1: Pythonic property access
print("\n1. Testing snake_case properties:")
print(f"   col.name: {col.name}")
print(f"   col.long_name: {col.long_name}")
print(f"   col.type: {col.type}")
print(f"   col.units: {col.units}")
print(f"   col.comments: {col.comments}")
print(f"   col.parent: {type(col.parent)}")

# Test 2: Setting properties with snake_case
print("\n2. Setting properties with snake_case:")
col.units = "seconds"
col.comments = "Time measurement"
print(f"   col.units: {col.units}")
print(f"   col.comments: {col.comments}")

# Test 3: Backward compatibility with PascalCase
print("\n3. Testing backward compatibility (PascalCase):")
print(f"   col.Name: {col.Name}")
print(f"   col.LongName: {col.LongName}")
print(f"   col.Type: {col.Type}")
print(f"   col.Units: {col.Units}")
print(f"   col.Comments: {col.Comments}")
print(f"   col.Parent: {type(col.Parent)}")

# Test 4: Setting with PascalCase still works
print("\n4. Setting properties with PascalCase:")
col.Name = "A_modified"
col.Units = "modified_units"
print(f"   col.name: {col.name} (should be 'A_modified')")
print(f"   col.units: {col.units} (should be 'modified_units')")

# Test 5: Both naming conventions are synchronized
print("\n5. Testing synchronization between naming conventions:")
col.long_name = "New Long Name"
print(f"   col.long_name: {col.long_name}")
print(f"   col.LongName: {col.LongName} (should match)")

col.Comments = "New comment via PascalCase"
print(f"   col.comments: {col.comments}")
print(f"   col.Comments: {col.Comments} (should match)")

# Save and close
origin.save()
origin.close()

print("\n=== Pythonic naming test completed successfully! ===")
