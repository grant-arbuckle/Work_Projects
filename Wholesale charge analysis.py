# Program that consolidates and parses data from PDF files, uses For loops to read in proper count of columns (PDFs are unstructured), cleans it up and converts to csv

from requests import head
import tabula
import os, PyPDF2
import pandas as pd
import glob

# Declare file name and directory paths
user_file_name="WH1023 07.21-04.22 V2"  # change this as needed
path1 = "V:/EOP_workpapers/Books/WMI/Python Reports/WH1023 Monthly Files/" # change as needed
path2 = "V:/EOP_workpapers/Books/WMI/Python Reports/WH1023 07.21-04.22 V2" # change as needed
path3 = "V:/EOP_workpapers/Books/WMI/Python Reports/"

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

#convert to csv, clean dataframe
new_file_path = path2 + '.csv'
tabula.convert_into((path2 + '.pdf'), new_file_path, output_format='csv', pages='all') # change as needed
files = glob.glob(new_file_path) # change as needed

#prepare csv for reading in proper amount of columns
largest_column_count = 0
file_delimiter = ','

with open(new_file_path, 'r') as temp_f:
  lines = temp_f.readlines()

  for l in lines:
    column_count = len(l.split(file_delimiter)) + 1

    largest_column_count = column_count if largest_column_count < column_count else largest_column_count

column_names = [i for i in range(0, largest_column_count)]

#read in csv
df = pd.read_csv(new_file_path,header=None,delimiter=file_delimiter,names=column_names)

#clean and filter data
df = df.rename(columns={
  0:'Customer Number', 1:'Customer Type', 2:'Business Name', 3:'Invoice Number', 4:'Order Type', 5:'Date', 6:'S/H', 7:'S/T', 8:'Amount Due', 9:'placeholder', 10:'placeholder', 11:'IF VALUES HERE SHIFT DATA LEFT ONE COL'})
df = df.replace(",","",regex=True)

#filter out unneeded/total rows
prefixes = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
df = df[df['Customer Number'].str.startswith(tuple(prefixes))]

#re-save file
df.to_csv(new_file_path)
