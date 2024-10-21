import pandas as pd
import xlsxwriter as xlsx

inputfile = 'library/Từ điển mức câu_Gia Lai.csv'
outputfileXLSX = 'output/GiaLai.xlsx'
outputfileCSV = 'output/GiaLai.csv'


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
            cell3 = str(df1.iat[row, col+2])
            # If cell2 is empty, skip to next row
            if not cell2.strip() or not cell1.strip() or not cell3.strip():
                continue
            # Add row's content to data
            cell1 = check_space(cell1)
            cell2 = check_space(cell2)
            cell3 = check_space(cell3)
            data.append([cell1, cell2, cell3])
    return data

def combineSentences(data):
    # For each row, combine cell1 with other cell1s, cell2 with other cell2s if their cell3s are the same
    combined = []
    DataSize = len(data)
    for row in range(1, len(data)):
        if(row % 200 == 0):
            print('Row {} out of {}'.format(row, DataSize))
        for row2 in range(row + 1, len(data)):
            # Ensure both rows have at least 3 elements
            if len(data[row]) >= 3 and len(data[row2]) >= 3:
                if data[row][2] == data[row2][2]:
                    # String in cell1 of row and row2 are combined into a new string and stored in cell1 of new row
                    combined.append([data[row][0] + ' ' + data[row2][0], data[row][1] + ' ' + data[row2][1], data[row][2]])
            else:
                print('Row {} or {} has less than 3 elements'.format(row, row2))
    combinedSize = len(combined)
    print('Old data size: {}, New data size: {}'.format(DataSize, combinedSize))
    return combined

def dataToXLSX(data, outputfile):
    # Write data to output file
    workbook = xlsx.Workbook(outputfile)
    worksheet = workbook.add_worksheet()

    for row_num, row_data in enumerate(data):
        for col_num, cell_data in enumerate(row_data):
            worksheet.write(row_num, col_num, cell_data)
    workbook.close()
    return

def dataToCSV(data, outputfile):
    # Write data to output file
    df = pd.DataFrame(data)
    df.to_csv(outputfile, index=False)
    return


def main():
    data = dataFromCSV(inputfile)
    combined = combineSentences(data)
    dataToCSV(combined, outputfileCSV)
    return

if __name__ == '__main__':
    main()