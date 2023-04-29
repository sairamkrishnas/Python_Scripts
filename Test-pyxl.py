import os

import pandas as pd
from openpyxl import load_workbook

# read the two Excel files into two separate DataFrames
df1 = pd.read_excel(os.getcwd() + '/file1.xlsx')
df2 = pd.read_excel(os.getcwd() + '/file2.xlsx')

# compare the two DataFrames and highlight the differences in red
diff = df1.compare(df2).fillna('')

# write the differences to a new Excel file
with pd.ExcelWriter(os.getcwd() + '/output.xlsx') as writer:
    diff.to_excel(writer, sheet_name='Differences', index=False)

# change the tab color to red
book = load_workbook('path/to/output.xlsx')
writer = pd.ExcelWriter('path/to/output.xlsx', engine='openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
ws = book['Differences']
ws.sheet_properties.tabColor = 'FF0000'
writer.save()
