import pandas as pd
import xlsxwriter as xlsx
import openpyxl as pxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

nameFolder = 'MERGE'
nameOutput = 'Merged.xlsx'

def mergeXLSX(nameFolder, nameOutput):
    data = [
        ['Bahnar', 'Vietnamese', 'Source', 'Chapter'],
    ]
    
    #List all files in the folder
    files = os.listdir(nameFolder)
    for file in files:
        if file.endswith('.xlsx'):
            print('Processing file:', file)
            df1 = pd.read_excel(nameFolder + '/' + file)
            for row in range(df1.shape[0]):
                for col in range(df1.shape[1] - 3):
                    if df1.iat[row, col+1] == None:
                        continue
                    cell1 = str(df1.iat[row, col])
                    cell2 = str(df1.iat[row, col+1])
                    cell3 = str(df1.iat[row, col+2])
                    cell4 = str(df1.iat[row, col+3])
                    data.append([cell1, cell2, cell3, cell4])
            print('Done processing file:', file)
    print('Writing to file:', nameOutput)
    workbook = xlsx.Workbook(nameOutput)
    worksheet = workbook.add_worksheet()
    for row_num, data_row in enumerate(data):
        for col_num, data_cell in enumerate(data_row):
            worksheet.write(row_num, col_num, data_cell)
    workbook.close()
    print('Done!')
    
mergeXLSX(nameFolder, nameOutput)
    