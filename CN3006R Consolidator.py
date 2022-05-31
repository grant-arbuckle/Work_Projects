# This program consolidates csv files that were previously processed one at a time into one file, with an added "Cont_code" column to identify entries from each file

from importlib.resources import path
import pandas as pd
import glob
import os

files = glob.glob("//ann-srv-acct/Accounting/EOP_workpapers/Books/AAL/2022/02- Feb/CN3006R/CN3006R_*.csv") # Change this line   

df = pd.concat([pd.read_csv(fp, skiprows=7).assign(Cont_Code=os.path.basename(fp)) for fp in files])
df = df[["Cont_Code", "Type", "Description", "Quantity", "Amount", "Unnamed: 4", "U.S.", "Non-U.S."]].drop(columns= ["Unnamed: 4", "U.S.", "Non-U.S."])
df["Cont_Code"] = df["Cont_Code"].str.replace("CN3006R_","").str.replace(".csv","")

# Exporting to Excel
out_path = "//ann-srv-acct/Accounting/EOP_workpapers/Books/AAL/2022/02- Feb/CN3006R/CN3006R_Consolidated"
month = input('Enter month in format "Jan":')
writer = pd.ExcelWriter(out_path + "_" + month + ".xlsx", engine='xlsxwriter')
df.to_excel(writer, sheet_name="Sheet1")
writer.save()
