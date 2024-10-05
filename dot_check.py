import pandas as pd
import xlsxwriter as xlsx

# Enter the path to the input file and the output file
# Then run

inputfile = r'library/Từ điển mức câu_Gia Lai.csv'
outputfile = r'output/Gia Lai.xlsx'
InputdataType = 'csv'

def check_space(word):
    if word[0] == ' ':
        return word[1:]
    return word

def dataFromCSV(file):
    data = [
        ['A', 'B', 'C'],
    ]
    df1 = pd.read_csv(file)
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
            
def dotEndCheck(data):
    # Check all rows, if a row NOT end with a dot, a question mark, or an exclamation mark, add a dot to the end of that row.
    count = 0
    for row in range(1, len(data)):
        if data[row][1][-1] not in ['.', '?', '!', '."', ':', '.\'', ',']:
            data[row][1] += '.'
            count += 1
    print(f'Added {count} dots to the end of sentences.')
    return data

def writeDataToXLSX(data, output_file):
    # Write data to output file
    workbook = xlsx.Workbook(output_file)
    worksheet = workbook.add_worksheet()

    for row_num, row_data in enumerate(data):
        for col_num, cell_data in enumerate(row_data):
            worksheet.write(row_num, col_num, cell_data)
    workbook.close()
    return

def runner(input_file, output_file, dataType):
    if dataType == 'csv':
        data = dataFromCSV(input_file)
    elif dataType == 'xlsx':
        data = dataFromXLSX(input_file)
    else:
        print('Invalid data type.')
        return
    data = dotEndCheck(data)
    writeDataToXLSX(data, output_file)
    return

runner(inputfile, outputfile, InputdataType)