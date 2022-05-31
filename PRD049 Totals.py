from importlib.resources import path
import pandas as pd
import tabula

path1 = 'V:/EOP_workpapers/Books/AAL/2022/04- Apr/PRD049' # Change path for each new month

new_file_path = path1 + ' TOTALS.xlsx'

df = tabula.io.read_pdf(input_path=path1+'.pdf', pages="1-25", multiple_tables=False, columns=[280, 350, 440], guess=False)
df = df[0]
df = df.rename(columns={df.columns[0]: 'Total_Name',
  df.columns[1]: 'On_Hand_Qty',
  df.columns[2]: 'Total_Value',
  df.columns[3]: 'TO_DELETE'}
  )
df = df.drop(columns=['On_Hand_Qty', 'TO_DELETE'])
df = df[df['Total_Name'].notna()]
df = df[df["Total_Name"].str.contains('Total')]
df = df.reset_index().drop(columns='index')

df.to_excel(new_file_path)