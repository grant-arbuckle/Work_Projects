#This program aggregates the data from the "DM3030 1 & 2" program in the form of pivot tables in an excel ouput

import pandas as pd
import numpy as np

df = pd.read_excel('//in-fileserver/Business_Office/Cash Manager/Daily Deposit Reports/DM3030/2022/Python Reports/AAL_DM3030_April 04.23.22' + '.xlsx', sheet_name="Sheet1") # change filename as needed
df2 = df.pivot_table(index = ['TYPE'], values = ['AMOUNT'], aggfunc ='sum')
df3 = df.pivot_table(index = ['Cont_Name'], values = ['AMOUNT'], aggfunc='sum')

excelfilename = '//in-fileserver/Business_Office/Cash Manager/Daily Deposit Reports/DM3030/2022/Python Reports/AAL_DM3030_April 04.23.22' + '.xlsx'
writer = pd.ExcelWriter(excelfilename, engine = 'xlsxwriter')

df = df.drop(columns = ['Amount 1', 'Amount 2'])

df.to_excel(writer, sheet_name="Data")
df2.to_excel(writer, sheet_name="Cont_Type Summary")
df3.to_excel(writer, sheet_name="Cont_Summary")
writer.save()
