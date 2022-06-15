import tabula
import pandas as pd
import os, PyPDF2

# Consolidate daily files into one pdf
directory = "//in-fileserver/Business_Office/Cash Manager/Daily Deposit Reports/DM3030/2022/Python Reports/May DM3030/"  # change this as needed
user_file_name="AAL_DM3030 May Consolidated"  # change this as needed
path_with_pdf = "//in-fileserver/Business_Office/Cash Manager/Daily Deposit Reports/DM3030/2022/Python Reports/" + user_file_name + ".pdf"

os.chdir(directory)
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
pdfOutput = open(path_with_pdf, 'wb')
pdfWriter.write(pdfOutput)
pdfOutput.close()
