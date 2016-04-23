# Parsing-PDF-in-Python
This is python code which can be used to parse and extract data from PDF's containing NCCN Biomarker data.

- The libraries used in the Python programs are as follows:
PyPDF2 - for parsing PDF and extracting text
re - for regular expressions
os - for getting all the PDF files present in the give folder
xlwt - for writting the extracted data to excel file. 

- All the above libraries + Python 3 is needed for running this program. 
- Program can be simply run by keeping multiple PDF files containing similar information (this is because regex is used for
 data extraction) in one folder where Python program is saved. 
- Then simply run the program through editor or command line and you will see that the result excel file is generated at the
 same location.

