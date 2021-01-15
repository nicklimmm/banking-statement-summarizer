from PyPDF2 import PdfFileReader
from pikepdf import Pdf
from getpass import getpass
import re
import openpyxl
import os
import csv
import pandas as pd

password = getpass("Enter password: ")

cwd = os.getcwd()   # cwd = current working directory

temp_dir = cwd + '\\temp.pdf'
csv_dir = cwd + '\\summary_csv.csv'
excel_dir = cwd + '\\summary_excel.xlsx'

csv_file = open(csv_dir, 'w', newline='')   # open (or create if not available) csv file and write inside it
writer = csv.writer(csv_file)
writer.writerow(['YEAR', 'MONTH', 'INFLOW', 'OUTFLOW', 'NETFLOW'])

data_dict = dict()  # to store the data based on the year (key = year, value = [month, inflow, outflow, netflow, interest, avg_balance])

for file_name in os.listdir(cwd):
    if file_name.startswith("FRANK") and file_name.endswith(".pdf"):
        file_dir = cwd + '\\' + file_name

        try:
            temp_file = Pdf.open(file_dir, password=password)         # since PyPDF2 cannot open the encrypted file, we use pikepdf
            temp_file.save(temp_dir)                                  # to open the file and create a copy in a temporary pdf file that
            pdf_obj = PdfFileReader(temp_dir)                         # is already decrypted
        except:
            print('Wrong password! Please rerun the script again to create the summary.')
            exit()

        date_created = pdf_obj.getDocumentInfo()['/CreationDate']
        year = int(date_created[2:6])
        month = int(date_created[6:8]) - 1  # the statement is created 1 month later, so we decrement by 1 to get the actual data
        if month == 0:
            month = 12  # handles error when the date created is 1 (that is the statement for December)

        num_pages = pdf_obj.getNumPages()                   # to handle different number of pages in each file, and
        if num_pages == 3:                                  # the summary lies in the back pages of the file
            page_obj = pdf_obj.getPage(num_pages - 3)       
        else:
            page_obj = pdf_obj.getPage(num_pages - 2)
            
        text = page_obj.extractText().encode('ascii').decode('ascii')       # using regex to find the necessary details
        pattern = r'[0-9,\.]+\s+[0-9,\.]+\s+[0-9,\.]+\s+[0-9,\.]+[0-9]'
        result = re.findall(pattern, text)

        inflow, outflow, interest, avg_balance = result[-1].replace(',', '').split()
        netflow = round(float(inflow) - float(outflow), 2)
        
        if year not in data_dict:
            data_dict[year] = []
        data_dict[year].append((month, inflow, outflow, netflow))


for year in sorted(data_dict):                  # sort everything based on the date and write it into the csv file in order
    data_dict[year].sort()
    for month, *data in data_dict[year]:
        writer.writerow([year, month, *data])
    writer.writerow([])

csv_file.close()

csv_file = pd.read_csv(csv_dir)                 # copy the data into the excel file
excel_file = pd.ExcelWriter(excel_dir)
csv_file.to_excel(excel_file, index=False)

excel_file.save()

os.remove(temp_dir)     # remove the temporary pdf file to prevent anyone to access the unencrypted pdf
