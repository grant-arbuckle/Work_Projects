#THIS PROGRAM IS UNDER DEVELOPMENT
import pandas as pd
import numpy as np
import openpyxl

directory = "//in-fileserver/Business_Office/Cash Manager/Daily Deposit Reports/DM3030/2022/Python Reports/"
user_file_name="AAL_DM3030 May Consolidated"  # change this as needed
excel_path = directory + user_file_name + ".xlsx"

df = pd.read_excel(excel_path, sheet_name="Sheet1") # change filename as needed
df = df.drop(["Unnamed: 0"], axis=1)
df2 = df.pivot_table(index = ['TYPE'], values = ['Total'], aggfunc ='sum')

excelfilename = excel_path
writer = pd.ExcelWriter(excelfilename, engine = 'xlsxwriter')

#df = df.drop(columns = ['Amount 1', 'Amount 2'])

df.to_excel(writer, sheet_name="Data")
df2.to_excel(writer, sheet_name="Cont_Type Summary")
writer.save()
