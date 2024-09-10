import pandas as pd
import xlsxwriter as xlsx

def check_duplicate_rows(file):
    df = pd.read_excel(file)
    df.drop_duplicates(inplace=True)
    df.to_excel(file, index=False)

def check_space(word):
    if word[0] == ' ':
        return word[1:]
    return word

def check_empty_cells(file):
    #Delete rows that have at least one empty cell
    df = pd.read_excel(file)
    df.dropna(inplace=True)
    df.to_excel(file, index=False)

def compare_excel(file, output_file):
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
            cell1 = str(df1.iat[row, col]).split()
            cell2 = str(df1.iat[row, col+1])
            cell3 = str(df1.iat[row, col+2]).split()
            # If cell2 is empty, skip to next row
            if not cell2.strip():
                continue
            #check_space(cell2)
            if ',' not in cell2:
                # Add row's content to data
                if [cell2] == '':
                    continue
                cell2 = check_space(cell2)
                data.append([cell1, [cell2], cell3])
            else:
                # Split cell 2 by comma
                cell2_list = cell2.split(',')
                # Add row's content to data
                for i in range(0, len(cell2_list)):
                    if [cell2_list[i]] == '':
                        continue
                    cell2_list[i] = check_space(cell2_list[i])
                    data.append([cell1, [cell2_list[i]], cell3])

    # Write data to output file using xlsxwriter
    workbook = xlsx.Workbook(output_file)
    worksheet = workbook.add_worksheet()

    for row_num, row_data in enumerate(data):
        for col_num, cell_data in enumerate(row_data):
            worksheet.write(row_num, col_num, ' '.join(cell_data))

    workbook.close()

def separate_sentences(file, output_file):
    data = [
        ['A', 'B'],
    ]
    df1 = pd.read_excel(file)
    # For each row, check 3 cells
    for row in range(df1.shape[0]):
        for col in range(df1.shape[1] - 1):
            if df1.iat[row, col+1] == None:
                continue
            # Make sure cell 1, 2 are strings
            cell1 = str(df1.iat[row, col])
            cell2 = str(df1.iat[row, col+1])
            # If cell2 is empty, skip to next row
            if not cell2 or not cell1:
                continue
            # Separate cell1 and cell2 by period
            cell1_list = cell1.split('. ')
            cell2_list = cell2.split('. ')
            # Add row's content to data
            for i in range(0, len(cell1_list)):
                if [cell1_list[i]] == '':
                    continue
                if [cell2_list[i]] == '':
                    continue
                cell1_list[i] = check_space(cell1_list[i])
                cell2_list[i] = check_space(cell2_list[i])
                data.append([cell1_list[i], cell2_list[i]])
                

    # Write data to output file using xlsxwriter
    workbook = xlsx.Workbook(output_file)
    worksheet = workbook.add_worksheet()

    for row_num, row_data in enumerate(data):
        for col_num, cell_data in enumerate(row_data):
            worksheet.write(row_num, col_num, ''.join(cell_data))

    workbook.close()

file1 = r"D:\Projects\bahnar-word-split-and-example-finder\input.xlsx"
output_file = r"D:\Projects\bahnar-word-split-and-example-finder\output.xlsx"

compare_excel(file1, output_file)
check_duplicate_rows(output_file)
check_empty_cells(output_file)