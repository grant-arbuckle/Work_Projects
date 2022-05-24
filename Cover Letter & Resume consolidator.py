# Used the tabula and PyPDF2 modules to consolidate two pdf files

import tabula
import os, PyPDF2

#user_pdf_location=
os.chdir("C:/Users/Grant_Arbuckle/Desktop/Cover Letter & Resume/")
user_file_name="Grant Arbuckle Cover Letter & Resume"  # change this as needed

pdf2merge = []
for filename in os.listdir('.'):
  pdf2merge.append(filename)

pdfWriter = PyPDF2.PdfFileWriter()

for filename in pdf2merge:
  pdfFileObj = open(filename,'rb')
  pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
  for pageNum in range(pdfReader.numPages):
    pageObj = pdfReader.getPage(pageNum)
    pdfWriter.addPage(pageObj)
pdfOutput = open(user_file_name+'.pdf', 'wb')
pdfWriter.write(pdfOutput)
pdfOutput.close()
