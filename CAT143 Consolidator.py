from math import comb
import pandas as pd
import glob
import os
import openpyxl

new_files = glob.glob("V:/EOP_workpapers/Books/AAL/2022/CAT143/Save CAT143 Downloads Here/CAT143*")

main_df = pd.read_excel("V:/EOP_workpapers/Books/AAL/2022/CAT143/CAT143_Consolidated_As-Of Mar 2022.xlsx") # CHANGE THIS LINE to current filename each month
main_df = main_df.drop(columns="Unnamed: 0")

for file in new_files:
  df = pd.read_csv(file, skiprows=8)  # read in file
  df = df[["Catalog Format", "Order Type", "# of Shipped Orders"]]
  df = pd.pivot_table(df,index=["Catalog Format"],aggfunc=sum)  # Pivot data by catalog format
  date_column_df = pd.read_csv(file, skiprows=8).assign(report_date=os.path.basename(file))  # read in file again for separate date column df
  date_column_df = date_column_df.drop_duplicates(subset=['Catalog Format'])
  date_column_df["report_date"] = date_column_df["report_date"].str.replace("CAT143 - ","").str.replace(".csv","")
  date_column_df = date_column_df.drop(columns= ['Order Type', '# of Shipped Orders'])
  combined_df = pd.merge(df, date_column_df, how = "left", on = "Catalog Format")  # merge date column df with original df
  combined_df["report_date"] = pd.to_datetime(combined_df['report_date'], format= '%b %y')
  combined_df = combined_df.rename(columns={"# of Shipped Orders": "Orders Shipped"})
  combined_df.loc["ttl"] = combined_df.sum(numeric_only=True, axis=0)  # Create row for totals
  combined_df.at["ttl", "report_date"] = combined_df.at[1, "report_date"]  # place value in df at specific cell location
  combined_df.at["ttl", "Catalog Format"] = "TOTAL SHIPPED - MONTH"  # place text in df at specific cell location
  main_df = pd.concat([main_df, combined_df], axis=0)   # add new monthly files to original dataframe # USE BRACKETS WHEN USING CONCAT ON DATAFRAMES #

# Exporting to Excel
out_path = "V:/EOP_workpapers/Books/AAL/2022/CAT143/CAT143_Consolidated_As-Of "
month = input('Enter as-of date for filename, in format "mmm yyyy":')
writer = pd.ExcelWriter(out_path + month + ".xlsx", engine='xlsxwriter')
main_df.to_excel(writer, sheet_name="Totals By Month")
writer.save()

# Formatting Excel
file = out_path + month + ".xlsx"
wb = openpyxl.load_workbook(filename=file)
worksheet = wb.active

for cell in worksheet["D"]:
  cell.number_format = "MMM-YY"

#auto-adjusts column widths
dims = {}
for row in worksheet.rows:
  for cell in row:
    if cell.value:
      dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
for col, value in dims.items():
  worksheet.column_dimensions[col].width = value

wb.save(filename=file)