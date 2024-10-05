import pandas as pd
import xlsxwriter as xlsx
import re

def check_space(word):
    if word[0] == ' ':
        return word[1:]
    return word

def dataFromXLSX(file):
    data = [
        ['A', 'B', 'C'],
    ]
    df1 = pd.read_excel(file)
    # For each row, check 3 cells
    for row in range(df1.shape[0]):
        for col in range(df1.shape[1] - 2):
            if df1.iat[row, col+1] == None:
                continue
            # Make sure cell 1, 2, 3 are strings
            cell1 = str(df1.iat[row, col])
            cell2 = str(df1.iat[row, col+1])
            # If cell2 is empty, skip to next row
            if not cell2.strip() or not cell1.strip():
                continue
            # Add row's content to data
            if [cell2] == '':
                continue
            cell1 = check_space(cell1)
            cell2 = check_space(cell2)
            data.append([cell1, cell2])
    return data

def dataFromTexT(text):
    # Make sure the text is a string
    if not isinstance(text, str):
        return
    # From a textfile, split the text into a list of sentences (by . ? or !)
    text = text.replace('\n', ' ')
    sentences_list = re.split(r'[.!?]\s*', text)
    
    # Combine each 2 sentences into an element of a list
    double_sentences_list = []
    for i in range(0, len(sentences_list), 2):
        if i + 1 < len(sentences_list):
            double_sentences_list.append([sentences_list[i], sentences_list[i+1]])
        else:
            double_sentences_list.append([sentences_list[i], ''])
    return double_sentences_list

def dataToXLSX(data, outputfile):
    # Write data to output file
    workbook = xlsx.Workbook(outputfile)
    worksheet = workbook.add_worksheet()

    for row_num, row_data in enumerate(data):
        for col_num, cell_data in enumerate(row_data):
            worksheet.write(row_num, col_num, cell_data)
    workbook.close()
    return