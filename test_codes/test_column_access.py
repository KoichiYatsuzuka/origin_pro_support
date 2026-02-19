#%%
import pandas as pd
import origin_pro_support as ops
import os

# Create test data
df = pd.DataFrame({
    'Time': [1, 2, 3, 4, 5],
    'Dataset 1': [0, 100.3, 210.7, 343, 4.56],
    'Dataset 2': [1, 107, 211.3, 307.8, 445]
})

# Create Origin instance
origin_file = os.path.join(os.getcwd(), "test_column_access.opju")
origin = ops.OriginInstance(origin_file)

# Create a new workbook and get worksheet
book = origin.new_book()
wks = book[0]

# Load data
wks.from_df(df)

print("=== Testing Column Access Methods ===")

# Test 1: Index access
print("\n1. Testing index access wks[0]:")
try:
    col = wks[0]
    print(f"   Success: {type(col)}")
    print(f"   Column name: {col.Name}")
    print(f"   Column long name: {col.LongName}")
except Exception as e:
    print(f"   Error: {e}")

# Test 2: Column collection access
print("\n2. Testing wks.Columns(0):")
try:
    col = wks.Columns(0)
    print(f"   Success: {type(col)}")
    print(f"   Column name: {col.Name}")
    print(f"   Column long name: {col.LongName}")
except Exception as e:
    print(f"   Error: {e}")

# Test 3: get_columns method
print("\n3. Testing wks.get_columns():")
try:
    cols = wks.get_columns()
    print(f"   Success: {len(cols)} columns")
    for i, col in enumerate(cols[:3]):  # Show first 3
        print(f"   Column {i}: {col.Name} ({col.LongName})")
except Exception as e:
    print(f"   Error: {e}")

# Test 4: Iteration
print("\n4. Testing iteration over worksheet:")
try:
    for i, col in enumerate(wks):
        if i >= 3:  # Show first 3
            break
        print(f"   Column {i}: {type(col)} - {col.Name}")
except Exception as e:
    print(f"   Error: {e}")

# Test 5: Column data access
print("\n5. Testing column data access:")
try:
    col = wks[0]
    data = col.get_data(0)  # Get data in default format
    print(f"   Success: {len(data)} data points")
    print(f"   First 3 values: {data[:3]}")
except Exception as e:
    print(f"   Error: {e}")

# Test 6: Column properties
print("\n6. Testing column properties:")
try:
    col = wks[1]
    print(f"   Name: {col.Name}")
    print(f"   LongName: {col.LongName}")
    print(f"   Units: {col.Units}")
    print(f"   Comments: {col.Comments}")
    print(f"   Type: {col.Type}")
except Exception as e:
    print(f"   Error: {e}")

# Test 7: Setting column properties
print("\n7. Testing setting column properties:")
try:
    col = wks[2]
    col.Units = "test units"
    col.Comments = "test comment"
    print(f"   Success: Units={col.Units}, Comments={col.Comments}")
except Exception as e:
    print(f"   Error: {e}")

# Save and close
origin.save()
origin.close()

print("\n=== Column access test completed ===")
