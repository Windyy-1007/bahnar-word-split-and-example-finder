import pandas as pd
import xlsxwriter as xlsx
import openpyxl as pxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

inputFile = "Merged.xlsx"
outputFile = "Enhanced.xlsx"

def checkSpace(word):
    word = word.replace('\xa0', ' ')
    if word[0] == ' ' or word[0] in '-_?!.,;:■.\\|/<>[]@#$%^&*+=~`':
        if len(word) > 1:
            return checkSpace(word[1:])
    if len(word) > 2:
        if word[0] in '0123456789' and word[1] != ' ':
            return checkSpace(word[1:])
    return word

def checkNonSuitableCharacters(word):
    #If the word contains any of these in a set, replace it with a blank ''
    word.replace('■', '')
    word.replace('*', '')
    word.replace('⦁', '')
    word.replace('©', '')
    word.replace('°', '')
    word.replace('¬', '')
    word.replace('†', '')
    word.replace('‡', '')
    word.replace('®', '')
    word.replace('•', '')
    word.replace('脫', '')
    return word
    

def enhanceWord(word):
    word = checkSpace(word)
    word = checkNonSuitableCharacters(word)
    return word
    

def enhanceFile(inputFile):
    data = [
        ['Bahnar', 'Vietnamese', 'Source', 'Chapter'],
    ]
    df1 = pd.read_excel(inputFile)
    for row in range(df1.shape[0]):
        for col in range(df1.shape[1] - 3):
            if df1.iat[row, col+1] == None:
                continue
            cell1 = str(df1.iat[row, col])
            cell2 = str(df1.iat[row, col+1])
            cell3 = str(df1.iat[row, col+2])
            cell4 = str(df1.iat[row, col+3])
            
            # If any cell is empty, add "-" to that cell
            if (not cell1.strip() or cell1 == 'nan') and (not cell2.strip() or cell2 == 'nan'):
                continue
            if not cell1.strip() or cell1 == 'nan':
                cell1 = ' '
            if not cell2.strip() or cell2 == 'nan':
                cell2 = ' '
            if not cell3.strip() or cell3 == 'nan':
                cell3 = ' '
            if not cell4.strip() or cell4 == 'nan':
                cell4 = ' '
            
            data.append([enhanceWord(cell1), enhanceWord(cell2), enhanceWord(cell3), enhanceWord(cell4)])
            print('Enhanced row:', data[-1])
    return data

def setCharacters(inputfile):
    charSet = set()
    df1 = pd.read_excel(inputfile)
    for row in range(df1.shape[0]):
        for col in range(df1.shape[1]):
            cell = str(df1.iat[row, col])
            for char in cell:
                charSet.add(char)
    return charSet

def setThreeFirstCharacters(inputfile):
    charSet = set()
    df1 = pd.read_excel(inputfile)
    for row in range(df1.shape[0]):
        for col in range(df1.shape[1]):
            cell = str(df1.iat[row, col])
            if len(cell) > 3:
                cell = cell[:3]
            #If first character is a normal character in the alphabet, skip it
            if cell[0].isalpha():
                continue
            for char in cell:
                charSet.add(char)
    return charSet

def writeToFile(data, outputFile):
    workbook = xlsx.Workbook(outputFile)
    worksheet = workbook.add_worksheet()
    for row_num, data_row in enumerate(data):
        for col_num, data_cell in enumerate(data_row):
            worksheet.write(row_num, col_num, data_cell)
    workbook.close()

def main():
    data = enhanceFile(inputFile)
    writeToFile(data, outputFile)
    print('Enhanced file:', outputFile)
    
main()
