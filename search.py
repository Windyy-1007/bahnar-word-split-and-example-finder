from urllib3 import *
import json
import csv
import pandas as pd
import xlsxwriter as xlsx

# Tasks
## Convert target xlsx file to json file
def convertToJSON(source = 'library.xlsx', target = 'library.json'):
    
    excel_data_df = pd.read_excel(source, sheet_name='pairs')
    
    json_str = excel_data_df.to_json()
    
    with open(target, 'w') as json_file:
        json_file.write(json_str)
        
    return

def convertToCSV(source = 'library.xlsx', target = 'library.csv'):
    excel_data_df = pd.read_excel(source, sheet_name='pairs')
    excel_data_df.to_csv(target, index = None, header=True)
    return

def deleteQuery(url = 'http://localhost:8983/solr/mycore/update?commit=true'):
    http = PoolManager()
    r = http.request('POST', url, body=b'<delete><query>*:*</query></delete>', headers={'Content-Type': 'text/xml'})
    return

## Upload json file to the solr query
def uploadToSolr(source = 'library.csv', url = 'http://localhost:8983/solr/mycore/update?commit=true'):
    http = PoolManager()
    with open(source, 'r') as file:
        data = file.read()
        r = http.request('POST', url, body=data.encode('utf-8'), headers={'Content-Type': 'application/csv'})
    return



## Implement pairing search: For each pairs of (An, Bn), search for a tuplet of (A, B) in the json file which contains the word An in A and the word Bn in B. Return B as the result.
## Save the result in a data list
def findExamples(wordSource = 'output.xlsx', url = 'http://localhost:8983/solr/mycore/select?q='):
    wordData = [
        ['Vietnamese', 'Bahnar'],
    ]
    
    df = pd.read_excel(wordSource)
    
    for row in range(df.shape[0]):
        for col in range(df.shape[1] - 2):
            
            if df.iat[row, col+1] == None:
                continue
            
            # Make sure cell 1, 2 are strings, and they are not empty
            cell1 = str(df.iat[row, col])
            cell2 = str(df.iat[row, col+1])
            if not cell1.strip() or not cell2.strip():
                continue
            # Add cell1 and cell2 to wordData
            wordData.append([cell1, cell2])
    
    http = PoolManager()
    for wordPair in wordData:
        word1 = wordPair[0]
        word2 = wordPair[1]
        matchVietList = []
        matchBahnarList = []
        resultList = []
        # CSV file
        # Search for substring word1 in the Vietnamese column, save the corresponding Bahnar column to matchVietList
        r = http.request('GET', f'{url}?q=Vietnamese:*{word1}*')
        data = csv.reader(r.data.decode('utf-8').split('\n'))
        for row in data:
            if row:
                matchVietList.append(row[0])
                print (row[0])
        # Search for substring word2 in the Bahnar column, save the corresponding Vietnamese column to matchBahnarList
        r = http.request('GET', f'{url}?q=Bahnar:*{word2}*')
        data = csv.reader(r.data.decode('utf-8').split('\n'))
        for row in data:
            if row:
                matchBahnarList.append(row[0])
                print (row[0])
                
        print(matchVietList)
        print(matchBahnarList)
        

        
        # Find the intersection of matchVietList and matchBahnarList
        result = ''
        # Find first intersection of matchVietList and matchBahnarList
        for viet in matchVietList:
            if viet in matchBahnarList:
                result = viet
                break
        resultList.append(result)
        print(resultList)
    return resultList

## Save the result in a xlsx file
def saveResult(result = [], target = 'target.xlsx'):
    workbook = xlsx.Workbook(target)
    worksheet = workbook.add_worksheet()
    for row in range(len(result)):
        worksheet.write(row, 0, result[row])
    workbook.close()
    return

convertToCSV()
deleteQuery()
uploadToSolr()
saveResult(findExamples(), 'target.xlsx')
print('Done')
