# This program consolidates several days worth of of pdf files, converts them to csv, then aggregates the data in the form of pivot tables in an excel output

import tabula
import os, PyPDF2
import pandas as pd
from numpy import isin
from openpyxl import Workbook
import openpyxl
import glob

# 1st piece
os.chdir("//in-fileserver/Business_Office/Cash Manager/Daily Deposit Reports/DM3030/2022/04 Apr/")
user_file_name="AAL_DM3030 Strip Pack Club 04.23.22"  # change this as needed

pdf2merge = []
for filename in os.listdir('.'):
  if filename.startswith('AAL_DM3030'):     #change this as needed
    pdf2merge.append(filename)

pdfWriter = PyPDF2.PdfFileWriter()

for filename in pdf2merge:
  pdfFileObj = open(filename,'rb')
  pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
  for pageNum in range(pdfReader.numPages):
    pageObj = pdfReader.getPage(pageNum)
    pdfWriter.addPage(pageObj)
pdfOutput = open("//in-fileserver/Business_Office/Cash Manager/Daily Deposit Reports/DM3030/2022/Python Reports/AAL_DM3030 Strip Pack Club 04.23.22"+'.pdf', 'wb')
pdfWriter.write(pdfOutput)
pdfOutput.close()

#convert to csv
new_file_name = (user_file_name + '.csv')
df = tabula.convert_into(("//in-fileserver/Business_Office/Cash Manager/Daily Deposit Reports/DM3030/2022/Python Reports/AAL_DM3030 Strip Pack Club 04.23.22" + '.pdf'), "//in-fileserver/Business_Office/Cash Manager/Daily Deposit Reports/DM3030/2022/Python Reports/" + new_file_name, output_format='csv', pages='all')

# 2nd piece

# adding column for file date
files = glob.glob("//in-fileserver/Business_Office/Cash Manager/Daily Deposit Reports/DM3030/2022/Python Reports/April DM3030/*.csv") # Change this line for new month
df = pd.concat([pd.read_csv(fp).assign(report_date=os.path.basename(fp)) for fp in files])
df["report_date"] = df["report_date"].str.replace(".csv","")
########
df = df.drop(columns = ['U.S.U.S.', 'Unnamed: 4', 'U.S.', 'CanadaCanada', 'Unnamed: 7', 'Canada  DebitCR Memo/'])
df = df.dropna(subset=['Unnamed: 0'])
df = df.rename(columns = {
  'Unnamed: 0': 'Cont_Name',
  'Unnamed: 1': 'Amount 1',
  'Unnamed: 2': 'Amount 2'
})

#print(df['Cont_Name'].unique())
df = df.loc[~df.Cont_Name.isin(['ications', 'Agency', 'Total', 'Total cash - Publications', 'inuities', 'Series total', '*** Series Credits ***', '**** Total Credits ****', 'Total cash - Continuities', 'Specialized Service Modules', 'Total cash - Specialized Servi' 'Catalog sales', 'Product fulfillment payments', 'Miscellaneous credits', "Total cash - Specialized Servi", "Catalog sales", "Insufficient check payment", "Annie's Crochet Specials", "Just CrossStitch", "Digital Just CrossStitch", "Digital Crochet Specials", "Annie's Creative Studio", "Knit & Crochet Now! All Access"])]
# Removing these made one-day DM3030 tie out, but caused the output to be lower than the pdf for the DM3030 Monthly file: 
#   "Annie's Crochet Specials", "Just CrossStitch", "Digital Just CrossStitch", "Digital Crochet Specials", "Annie's Creative Studio", "Knit & Crochet Now! All Access"

df["AMOUNT"] = pd.NaT
df["TYPE"] = pd.NaT
df.to_excel("//in-fileserver/Business_Office/Cash Manager/Daily Deposit Reports/DM3030/2022/Python Reports/AAL_DM3030 Strip Pack Club 04.23.22" + ".xlsx")
wb = openpyxl.load_workbook(filename='//in-fileserver/Business_Office/Cash Manager/Daily Deposit Reports/DM3030/2022/Python Reports/AAL_DM3030 Strip Pack Club 04.23.22.xlsx')
worksheet = wb.active
worksheet['F2'] = '=IF(OR(C2="",C2="DTP"),D2,C2)'
worksheet['G2'] = "=VLOOKUP(B2,'[DM3030 Vlookup sheet.xlsx]Sheet1'!$A:$B,2,FALSE)"  #copy of Vlookup sheet is needed to run this program
worksheet['A1'] = 'DELETE THIS COLUMN'
wb.save(filename="//in-fileserver/Business_Office/Cash Manager/Daily Deposit Reports/DM3030/2022/Python Reports/AAL_DM3030 Strip Pack Club 04.23.22.xlsx")
