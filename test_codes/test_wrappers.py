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
origin_file = os.path.join(os.getcwd(), "test_wrappers.opju")
origin = ops.OriginInstance(origin_file)

# Test new_workbook returns wrapped WorkbookPage
book = origin.new_workbook("TestWorkbook")
print(f"Book created: {book}")
print(f"Book is WorkbookPage: {isinstance(book, ops.WorkbookPage)}")

# Test new_sheet returns wrapped Worksheet
print("\nTesting new_sheet...")
sheet = origin.new_sheet()
print(f"Sheet type: {type(sheet)}")
print(f"Sheet is Worksheet: {isinstance(sheet, ops.Worksheet)}")

# Test from_df method
print("\nTesting from_df method...")
sheet.from_df(df)
print(f"Worksheet has {sheet.get_cols()} columns and {sheet.get_rows()} rows")

# Test from_list method
print("\nTesting from_list method...")
sheet.from_list(3, [0.1, 0.15, 0.1, 0.2, 0.125], 
                lname='Error', 
                units='arb. units', 
                comments='Error bars',
                axis='E')

# Test column access returns wrapped Column
print("\nTesting column access...")
col = sheet[0]
print(f"Column type: {type(col)}")
print(f"Column is Column: {isinstance(col, ops.Column)}")
print(f"Column long name: {col.LongName}")

# Test new_graph returns wrapped GraphPage
print("\nTesting new_graph...")
graph = origin.new_graph()
print(f"Graph type: {type(graph)}")
print(f"Graph is GraphPage: {isinstance(graph, ops.GraphPage)}")

# Test new_notes returns wrapped NotePage
print("\nTesting new_notes...")
notes = origin.new_notes("Test Notes")
print(f"Notes type: {type(notes)}")
print(f"Notes is NotePage: {isinstance(notes, ops.NotePage)}")

# Save and close
origin.save()
origin.close()

print("\nAll wrapper tests completed successfully!")
