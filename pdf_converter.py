# importing required modules
import tabula
import PyPDF2
import pandas as pd
import os
import glob
import csv
import xlwt


# creating a pdf file object
pdf_path = '812285_electronicsreliabilityreport.pdf'
pdfFileObj = open(pdf_path, 'rb')

# creating a pdf reader object
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
file_reader = open("pdf_text.txt","w")

# creating a page object
for i in range(pdfReader.numPages):
    pageObj = pdfReader.getPage(i)
    # extracting text from page and writing in text file
    file_reader.write(pageObj.extractText())

# converts the table into csv format
multi_tables = tabula.read_pdf(pdf_path, pages="all", multiple_tables=True,output_format="csv",lattice=True)
for i in range(len(multi_tables)):
    si = "pdf_tables"+str(i)+".csv"
    multi_tables[i].to_csv(si)


# closing the pdf file object
pdfFileObj.close()
file_reader.close()


# New excel workbook
wb = xlwt.Workbook()
for csvfile in glob.glob(os.path.join('.', '*.csv')):

    fpath = csvfile.split("/", 1)
    fname = fpath[1].split(".", 1) ## fname[0] should be our worksheet name
    # Creates a new sheet in workbook
    ws = wb.add_sheet(fname[0])
    with open(csvfile, 'rt') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                ws.write(r, c, col)
wb.save('final_excel.xls')

# Removes all the .csv files
files=glob.glob('*.csv')
for filename in files:
    os.unlink(filename)

# Concatinated table for all the pages.
tabula.convert_into(pdf_path, "all_together.csv",pages='all', output_format="csv", stream=True)
