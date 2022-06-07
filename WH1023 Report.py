import tabula
import os, PyPDF2
import pandas as pd
import glob

# Declare file name and directory paths
user_file_name="WH1023 07.21-04.22 V2"  # change this as needed
path1 = "V:/EOP_workpapers/Books/WMI/Python Reports/WH1023 Monthly Files/" # change as needed
path2 = "V:/EOP_workpapers/Books/WMI/Python Reports/WH1023 07.21-04.22" # change as needed
output_path = path2 + ".xlsx"

# Merge page 1 of pdfs
os.chdir(path1)
pdf2merge = []
for filename in os.listdir('.'):
  pdf2merge.append(filename)

pdfWriter = PyPDF2.PdfFileWriter()

for filename in pdf2merge:
  pdfFileObj = open(filename,'rb')
  pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
  for pageNum in range(1,pdfReader.numPages):
    pageObj = pdfReader.getPage(pageNum)
    pdfWriter.addPage(pageObj)
pdfOutput = open(path2 +'.pdf', 'wb') # change as needed
pdfWriter.write(pdfOutput)
pdfOutput.close()

#Read in pdf as dataframe
df = tabula.io.read_pdf(input_path=path2+'.pdf', pages="all", multiple_tables=False, columns=[100, 110, 400, 450, 500, 600, 700, 790], guess=False)
df = df[0]

#clean and filter data
df = df.rename(columns={
  df.columns[0]: 'Customer_Number',
  df.columns[1]: 'Customer_Type',
  df.columns[2]: 'Business_Name',
  df.columns[3]: 'Invoice_Number',
  df.columns[4]: 'Order_Type',
  df.columns[5]: 'Date',
  df.columns[6]: 'S/H',
  df.columns[7]: 'S/T',
  df.columns[8]: 'Amount_Due',
  })

df = df[df['Customer_Number'].notna()]
df_len_5 = df[(df['Customer_Number'].map(len) == 5)]
df_len_4 = df[(df['Customer_Number'].map(len) == 4)]
df = pd.concat([df_len_5, df_len_4])
df = df.reset_index().drop(columns='index')

print(df.tail(25))

df.to_excel(output_path)