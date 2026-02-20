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

# Delete existing file if it exists
if os.path.exists(origin_file):
    print(f"Removing existing file: {origin_file}")
    os.remove(origin_file)

origin = None
try:
    print("Creating Origin instance...")
    origin = ops.OriginInstance(origin_file)
    
    print("Creating new workbook...")
    # Create a new workbook
    book = origin.new_book()
    
    print("Getting first worksheet...")
    # Get the first worksheet
    wks = book[0]
    
    # Test from_df method
    print("Testing from_df method...")
    wks.from_df(df)
    print("[OK] from_df method completed successfully")
    
    # Test from_list method
    print("Testing from_list method...")
    wks.from_list(3, [0.1, 0.15, 0.1, 0.2, 0.125], 
                  lname='Error', 
                  units='arb. units', 
                  comments='Error bars',
                  axis='E')
    print("[OK] from_list method (numeric index) completed successfully")
    
    # Test from_list with column letter
    wks.from_list('E', [10, 20, 30, 40, 50], 
                  lname='Additional Data', 
                  units='mm')
    print("[OK] from_list method (column letter) completed successfully")
    
    print("Data loaded successfully!")
    print(f"Worksheet has {wks.get_cols()} columns and {wks.get_rows()} rows")
    
    print("Test completed successfully!")
    
except Exception as e:
    print(f"[ERROR] Test failed with error: {e}")
    import traceback
    traceback.print_exc()
    
finally:
    # Always save and close if origin instance exists
    if origin is not None:
        try:
            print("Saving Origin file...")
            origin.save()
            print("[OK] File saved successfully")
        except Exception as save_error:
            print(f"[ERROR] Failed to save: {save_error}")
        
        try:
            print("Closing Origin instance...")
            origin.close()
            print("[OK] Origin instance closed successfully")
        except Exception as close_error:
            print(f"[ERROR] Failed to close: {close_error}")

print("Test execution finished.")
