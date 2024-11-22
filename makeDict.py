# import phrasal
# import search
# import csv
import pandas as pd
import xlsxwriter as xlsx
from tqdm import tqdm

# Flow:

# 1. Read vocab file (has 2 columns: Bahnar and Vietnamese), phrasal them and put pairs of Bahnar and Vietnamese into the correct datatset
dtBinhDinh = []
dtGiaLai = []
dtKonTum = []

def txtToData(file):
    # File has 1 column: Viet
    data = []
    with open(file, 'r', encoding='utf-8') as f:
        for line in tqdm(f, desc="Reading txt file"):
            data.append(line.strip())
    return data

def XLSXtoData(file):
    # File has 5 columns: Bahnar, Vietnamese, ExampleV, ExampleB, Source
    data = []
    df1 = pd.read_excel(file)
    # For each row, check 2 cells
    for row in tqdm(range(df1.shape[0]), desc="Reading XLSX file"):
        if df1.iat[row, 0] == None:
            continue
        # Make sure cell 1, 2 are strings
        cell1 = str(df1.iat[row, 0])
        cell2 = str(df1.iat[row, 1])
        cell3 = str(df1.iat[row, 2])
        cell4 = str(df1.iat[row, 3])
        cell5 = str(df1.iat[row, 4])
        # If cell2 is empty, skip to next row
        if not cell1.strip() or not cell2.strip():
            continue
        # Add row's content to data
        data.append([cell1, cell2, cell3, cell4, cell5])
    return data

def dataToXLSX(data, outputfile):
    # Write data to output file
    workbook = xlsx.Workbook(outputfile)
    worksheet = workbook.add_worksheet()

    for row_num, row_data in tqdm(enumerate(data), desc="Writing to XLSX file", total=len(data)):
        for col_num, cell_data in enumerate(row_data):
            worksheet.write(row_num, col_num, cell_data)
    workbook.close()
    return

# 2. Add examples for dataset: Binh Dinh, Gia Lai, Kon Tum

# search.search('library/dictionary/Từ điển mức câu_Bình Định (4.10.2024).csv', 'library/dictionary/Từ điển mức từ_Bình Định (7.10.2024).xlsx', 'output/Output_BinhDinh.xlsx')

# 3. Merge all datasets into one: If Vietnamse word do not exist in the seed dictionary, skip to next word.

dtBinhDinh = XLSXtoData('output/Output_BinhDinh.xlsx')
dtGiaLai = XLSXtoData('output/Output_GiaLai.xlsx')
dtKonTum = XLSXtoData('output/Output_KonTum.xlsx')
dtVietnamese = txtToData('library/dictionary/viet.txt')


# For each Vietnamese word, check if its translation exists in all 3 datasets. If yes, add to finalData, If no, that column is "---". If all 3 columns are "---", skip to next word.
      
# Add example: First priority is Binh Dinh, then Gia Lai, then Kon Tum. If no example is found, 3 last columns are "---"  

finalData = [
    ['Vietnamese', 'Binh Dinh', 'Gia Lai', 'Kon Tum', 'ExampleV', 'ExampleB', 'Source'],
]

for viet in tqdm(dtVietnamese, desc="Processing Vietnamese words"):
    viet = viet.strip()
    bDinh = "---"
    gLai = "---"
    kTum = "---"
    exampleV = "---"
    exampleB = "---"
    source = "---"
    for row in dtBinhDinh:
        if row[1] == viet:
            bDinh = row[0]
            exampleV = row[2]
            exampleB = row[3]
            source = row[4]
            break
    for row in dtGiaLai:
        if row[1] == viet:
            gLai = row[0]
            if (exampleV == "N/a" and exampleB == "N/a" and source == "N/a"):
                exampleV = row[2]
                exampleB = row[3]
                source = row[4]
            break
    for row in dtKonTum:
        if row[1] == viet:
            kTum = row[0]
            if (exampleV == "N/a" and exampleB == "N/a" and source == "N/a"):
                exampleV = row[2]
                exampleB = row[3]
                source = row[4]
            break
    if bDinh == "---" and gLai == "---" and kTum == "---":
        continue
    if exampleV == "N/a" and exampleB == "N/a" and source == "N/a":
        exampleV = "---"
        exampleB = "---"
        source = "---"
    finalData.append([viet, bDinh, gLai, kTum, exampleV, exampleB, source])

# Replace source with actual label
for i in tqdm(range(1, len(finalData)), desc="Replacing source labels"):
    #If "Sử thi" is in source, the whole source is replaced with "Sử thi"
    if "Sử thi" in finalData[i][6]:
        finalData[i][6] = "st"
    elif "Kinh Thánh" in finalData[i][6]:
        finalData[i][6] = "trtr"
    elif "Điền dã" in finalData[i][6]:
        finalData[i][6] = "dd"
    elif "Văn bản" in finalData[i][6]:
        finalData[i][6] = "vb"
    else:
        finalData[i][6] = "x"

dataToXLSX(finalData, 'output/Output_Final.xlsx')
print('Done')
