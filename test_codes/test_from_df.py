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
origin_file = os.path.join(os.getcwd(), "test_from_df.opju")
origin = ops.OriginInstance(origin_file)

# Create a new workbook
book = origin.new_book()

# Get the first worksheet
wks = book[0]

# Test from_df method
print("Testing from_df method...")
wks.from_df(df)

# Test from_list method
print("Testing from_list method...")
wks.from_list(3, [0.1, 0.15, 0.1, 0.2, 0.125], 
              lname='Error', 
              units='arb. units', 
              comments='Error bars',
              axis='E')

# Test from_list with column letter
wks.from_list('E', [10, 20, 30, 40, 50], 
              lname='Additional Data', 
              units='mm')

print("Data loaded successfully!")
print(f"Worksheet has {wks.get_cols()} columns and {wks.get_rows()} rows")

# Save and close
origin.save()
origin.close()

print("Test completed successfully!")
