from importlib.resources import path
import pandas as pd
import glob
import os

files = glob.glob("//ann-srv-acct/Accounting/EOP_workpapers/Books/AAL/2022/02- Feb/CN3008R/*.csv") # Change this line


df = pd.concat([pd.read_csv(fp, skiprows=7).assign(Cont_Code=os.path.basename(fp)) for fp in files])
df = df[["Cont_Code", "Product", "Description", "Count", "Total $", "Unit Cost $", "Cost of Goods $"]]
df["Cont_Code"] = df["Cont_Code"].str.replace("CN3008R_","").str.replace(".csv","")
df = df[df.Product != "Total"]

# Exporting to Excel
out_path = "//ann-srv-acct/Accounting/EOP_workpapers/Books/AAL/2022/02- Feb/CN3008R/CN3008R_Consolidated"
month = input('Enter month in format "Jan":')
writer = pd.ExcelWriter(out_path + "_" + month + ".xlsx", engine='xlsxwriter')
df.to_excel(writer, sheet_name="Sheet1")
writer.save()