# This program is used to format Paycom data exports to compare to one of our SQL reports I wrote that queries Business Central data to conduct analysis

from csv import excel
from ctypes import wstring_at
import pandas as pd
from openpyxl import load_workbook

###file path and engine setup
book = load_workbook("//in-fileserver/Business_Office/Business Central/Jet Reports/SFG Files/SFG Jet Reports/Intra-Month Labor Report.xlsx")
writer = pd.ExcelWriter("//in-fileserver/Business_Office/Business Central/Jet Reports/SFG Files/SFG Jet Reports/Intra-Month Labor Report.xlsx", engine = 'openpyxl')
writer.book = book

#df manipulation start
df = pd.read_excel(r'C:\Users\Grant_Arbuckle\Desktop\Python\Inputs\Labor_Allocation_Report_2.1-2.12.xlsx', sheet_name="Labor Allocation ")   # change filename as needed

regular_pay = df[df.Typecode.isin(['BER', 'BON', 'CDP', 'CPD', 'FTH', 'JRY', 'LCK', 'MSC', 'OTH', 'P', 'R', 'SFT', 'SPO', 'UNP', 'V'])]
regular_pay = regular_pay.drop(columns=['Department Desc', 'Payroll Profile Code', 'Payroll Profile Desc', 'Project Code', 'Project Desc', 'Supervisor Code', 'Supervisor Desc', 'Task Codes Code', 'Task Codes Desc', 'Full-Time/Part-Time Code', 'Full-Time/Part-Time Desc', 'clientcode', 'EE Code', 'EE Name', 'Typecode', 'Typedesc', 'Register Type', 'Hours/Units', 'Wages'])

OT_pay = df[df.Typecode.isin(['O'])]
OT_pay = OT_pay.drop(columns=['Department Desc', 'Payroll Profile Code', 'Payroll Profile Desc', 'Project Code', 'Project Desc', 'Supervisor Code', 'Supervisor Desc', 'Task Codes Code', 'Task Codes Desc', 'Full-Time/Part-Time Code', 'Full-Time/Part-Time Desc', 'clientcode', 'EE Code', 'EE Name', 'Typecode', 'Typedesc', 'Register Type', 'Hours/Units', 'Wages'])

###Pivots
regular_pay_pivot = regular_pay.pivot_table(index = ['Department Code'], values = ['Amount'], aggfunc ='sum')
regular_pay_pivot.columns = ['Regular_Pay']
OT_pay_pivot = OT_pay.pivot_table(index=['Department Code'], values=['Amount'], aggfunc ='sum')
OT_pay_pivot.columns = ['OT_Pay']

###Data frames to excel
writer.sheets = dict((ws.title, ws) for ws in book.worksheets) 
regular_pay_pivot.to_excel(writer, sheet_name="Paycom Sums", startcol=1, startrow=1)
OT_pay_pivot.to_excel(writer, sheet_name="Paycom Sums", startcol=3, startrow=1)
df.to_excel(writer, sheet_name="Paycom Raw Data", startrow=1)
writer.save()
