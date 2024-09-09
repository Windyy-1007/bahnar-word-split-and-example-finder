from urllib3 import *
from urllib.parse import quote as qoute
import json
import csv
import pandas as pd
import xlsxwriter as xlsx

# Tasks

## Convert the excel file to a csv file
def convertToCSV(source = 'library.xlsx', target = 'library.csv'):
    df = pd.read_excel(source, encoding='utf-8')
    df.to_csv(target, index=False, encoding='utf-8')
    return

def deleteQuery(url = 'http://localhost:8983/solr/mycore/update?commit=true'):
    http = PoolManager()
    r = http.request('POST', url, body=b'<delete><query>*:*</query></delete>', headers={'Content-Type': 'text/xml'})
    return

## Upload json file to the solr query
def uploadToSolr(source = 'library.csv', url = 'http://localhost:8983/solr/mycore/update?commit=true'):
    http = PoolManager()
    with open(source, 'r', encoding='utf-8', errors='replace') as file:
        data = file.read()
        r = http.request('POST', url, body=data.encode('utf-8'), headers={'Content-Type': 'application/csv'})
    return



## Implement pairing search: For each pairs of (An, Bn), search for a tuplet of (A, B) in the json file which contains the word An in A and the word Bn in B. Return B as the result.
## Save the result in a data list
def findExamples(wordSource = 'output.xlsx', url = 'http://localhost:8983/solr/mycore/select?q='):
    wordData = []
    resultList = []

    
    df = pd.read_excel(wordSource)
    
    print("Start searching")
    for row in range(df.shape[0]):
        for col in range(df.shape[1] - 2):
            print("Row: ", row, "Col: ", col)
            if df.iat[row, col+1] == None:
                print("None")
                continue
            
            # Make sure cell 1, 2 are strings, and they are not empty
            cell1 = str(df.iat[row, col])
            cell2 = str(df.iat[row, col+1])
            print("Appedning: ", cell1.encode("utf8"), cell2.encode("utf8"))
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
        print ("Looking for pair: ", word1.encode("utf8"), word2.encode("utf8"))
        # CSV file
        # Search for substring word1 in the Vietnamese column, save the corresponding Bahnar column to matchVietList
        r = http.request('GET', f'{url}_Vietnamese%3A*{qoute(word1)}*')
        data = csv.reader(r.data.decode('utf-8').split('\n'))
        for row in data:
            if row:
                matchVietList.append(row[0])
        # Search for substring word2 in the Bahnar column, save the corresponding Vietnamese column to matchBahnarList
        r = http.request('GET', f'{url}Bahnar%3A*{qoute(word2)}*')
        data = csv.reader(r.data.decode('utf-8').split('\n'))
        for row in data:
            if row:
                matchBahnarList.append(row[0])
        
        # Empty file matchVietList and matchBahnarList:
        open('matchVietList.txt', 'w').close()
        open('matchBahnarList.txt', 'w').close()
                
        #Save matchVietList and matchBahnarList to a txt file (for testing purpose)
        with open('matchVietList.txt', 'a', encoding="utf-8") as file:
            for line in matchVietList:
                file.write(line + '\n')
        with open('matchBahnarList.txt', 'a', encoding="utf-8") as file:
            for line in matchBahnarList:
                file.write(line + '\n')
                        
        
        
        resultBahnar = ''
        resultViet = ''
        result = ['A', 'B']
        # For each line in matchVietList.txt, search for the corresponding line in matchBahnarList.txt. If found, append the result to resultList
        with open('matchVietList.txt', 'r', encoding="utf-8") as file:
            for line in file:
                with open('matchBahnarList.txt', 'r', encoding="utf-8") as file2:
                    for line2 in file2:
                        if line == line2 and "Bahnar" in line:
                            resultBahnar = line2
                            resultViet = line
                            result = [resultViet, resultBahnar]
                            print("Result found")
                            break
                if result:  
                    break   
        
        # Save result to a txt file (for testing purpose)
        with open('result.txt', 'a', encoding="utf-8") as file:
            file.write(result + '\n')
        
        # Append result to resultList
        resultList.append(result)
        print("Added result: ", result.encode("utf8"))
    return resultList

## Save the resultList to a target xlsx file
def saveResult(resultList, target = 'target.xlsx'):
    workbook = xlsx.Workbook(target)
    worksheet = workbook.add_worksheet()
    for row in range(len(resultList)):
        worksheet.write(row, 0, resultList[row])
    workbook.close()
    return

deleteQuery()
uploadToSolr()
saveResult(findExamples(), 'target.xlsx')
print('Done')
