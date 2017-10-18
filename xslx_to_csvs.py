'''
XSLX to CSV transformer
Lei Mao
10/18/2017
University of Chicago

Python 3
Turn every sheets in the xslx file to multiple csv files.
'''

from openpyxl import load_workbook
import csv


file_name = 'data.xlsx'

# Load the whole excel workbook
workbook = load_workbook(file_name)

# Get all the sheet names
worksheets = workbook.get_sheet_names()

def sheet_export(worksheet, exported_file_name, first_line_disposal = False):
    # Read the data from sheet and export to csv file
    with open(exported_file_name, 'w', newline='', encoding='utf-8') as csvfile:
        spamwriter = csv.writer(csvfile, delimiter=',', 
            quotechar='|', quoting=csv.QUOTE_MINIMAL)

        for i, row in enumerate(worksheet.rows):
            if (i == 0) and (first_line_disposal == True):
                continue
            line = list()
            for cell in row:
                line.append(cell.value)
            spamwriter.writerow(line)

for sheet in worksheets:
    exported_file_name = sheet + '.csv'
    sheet_export(worksheet = workbook[sheet], exported_file_name = exported_file_name, first_line_disposal = True)