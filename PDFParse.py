#### Code Writtent in - Python 3.4.3 ####
#!/usr/bin/python3
__author__ = "Rajesh Shinde"
__version__ = "1.0"
__email__ = "rajesh27071992@gmail.com"
__status__ = "Development"

#### Imports ######

import PyPDF2
import re
import xlwt
import os

''' Get names of all PDF files present in PWD in a list
All the PDF's from which we need to extract similar data
can be kept under this folder ''' 

pdfFiles = []

for f in os.listdir():
    if (f.endswith(".pdf")):
        pdfFiles.append(f)
#print(pdfFiles)
    

####### Write the header line in Excel ##############
new_excel = xlwt.Workbook(encoding='utf-8')
sheet1 = new_excel.add_sheet("Result_Data")

String1 = """Disease Description,Specific Indication,
            Molecular Abnormality,Test,Chromosome,Gene Symbol,
            Test Detects,Methodology,
            NCCN Category of Evidence,
            Specimen Types,NCCN Recommendation - Clinical Decision,
            Test Purpose,When to Test,
            Guideline Page with Test Recommendation,Notes"""

split_string = String1.split(",")
print(split_string[0])

for i in range(len(split_string)):
    sheet1.write(0,i,label=split_string[i])
#new_excel.save("Result_excel.xlsx")

#### for page 1 - all extraction of data #####

def WriteDataToExcel(pageReading, row_number):
    test = re.compile(r'(Disease Description:)(.*)(Specific Indication:)'
                  '(.*)(Molecular Abnormality:)(.*)'
                  '(Test:)(.*)(Chromosome:)(.*)(Gene Symbol:)(.*)'
                  '(Test Detects:)(.*)'
                  '(Methodology:)(.*)'
                  '(NCCN Category of Evidence:)(.*)'
                  '(Specimen Types:)(.*)'
                  '(NCCN Recommendation - Clinical Decision:)(.*)'
                  '(Test Purpose:)(.*)'
                  '(When to Test:)(.*)'
                  '(Guideline Page with Test Recommendation:)(.*)'
                  '(Notes:)([a-zA-Z0-9\s+].*)?(\.!"\#$)?'
                  , re.MULTILINE)

    match = re.search(test, pageReading.extractText())

    for j in range(2,len((split_string*2))+2,2):
        if not match.group(j):
            sheet1.write(row_number,int(j/2)-1,label=" ")
        else:
            sheet1.write(row_number,int(j/2)-1,label=match.group(j))
    new_excel.save("Result_excel.xlsx")



rowNum = 1

for filename in pdfFiles:
    #print(filename)
    fh = open(filename, "rb") ### binary mode
    ####### Read PDF and Extract data fields ##############
    pdfReader = PyPDF2.PdfFileReader(fh)
    print(pdfReader.numPages)

    ################ Running loop through each page of file -
    ### Extracting and writing data of page in Excel ################

    pageNum = 0
    #rowNum = 1
    while pageNum < pdfReader.numPages:
        #print(pageNum)
        pageReading1 = pdfReader.getPage(pageNum)
        WriteDataToExcel(pageReading1, rowNum)
        pageNum = pageNum + 1
        rowNum = rowNum + 1

