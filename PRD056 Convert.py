# Program that cleans up and aggregates data by "Series" in the form of an excel pivot table output

import pandas as pd
import numpy as np
import openpyxl

filename = '//in-fileserver/Business_Office/Grant/Inventory Costing Project - For Jack/PRD056V 3.1.22'
df = pd.read_csv(filename + '.csv', skiprows=8, engine='python', encoding='latin1') # change filename as needed

df = df.rename(columns={'Cat': 'Series'})
df_pivot = df.pivot_table(index = ['Series'], values = ['Value'], aggfunc ='sum')
df_pivot.reset_index

excelfilename = ("//in-fileserver/Business_Office/Grant/Inventory Costing Project - For Jack/PRD056_Value_By_Series " + input('Enter date in format "mm.dd.yy":') + '.xlsx')
writer = pd.ExcelWriter(excelfilename, engine = 'xlsxwriter')

df_pivot.to_excel(writer, sheet_name="Pivot")

writer.save()

# Formatting Excel
wb = openpyxl.load_workbook(filename=excelfilename)
worksheet = wb.active

#adds number formatting to specific column
for cell in worksheet["B"]:
  cell.number_format = '$#,##0'

#auto-adjusts column widths
dims = {}
for row in worksheet.rows:
  for cell in row:
    if cell.value:
      dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
for col, value in dims.items():
  worksheet.column_dimensions[col].width = value

wb.save(filename=excelfilename)
