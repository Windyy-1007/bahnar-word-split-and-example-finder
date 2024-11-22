from urllib3 import *
from urllib.parse import quote as qoute
import json
import csv
import pandas as pd
import xlsxwriter as xlsx
from tqdm import tqdm  # Add tqdm for progress bar


# Tasks

## Convert the excel file to a csv file
def convertToCSV(source='library.xlsx', target='library.csv'):
    df = pd.read_excel(source, encoding='utf-8')
    df.to_csv(target, index=False, encoding='utf-8')
    return

## Delete query
def deleteQuery(url = 'http://localhost:8983/solr/mycore/update?commit=true'):
    http = PoolManager()
    r = http.request('POST', url, body=b'<delete><query>*:*</query></delete>', headers={'Content-Type': 'text/xml'})
    return

## Upload json file to the solr query
def uploadToSolr(source='library.csv', url='http://localhost:8983/solr/mycore/update?commit=true'):
    http = PoolManager()
    with open(source, 'r', encoding='utf-8', errors='replace') as file:
        data = file.read()
        r = http.request('POST', url, body=data.encode('utf-8'), headers={'Content-Type': 'application/csv'})
    return

def upWord(word):
    # If word[0] is ', keep it and check next letter instead
    if word[0] == "'":
        return word[0] + upWord(word[1:])
    return word[0].upper() + word[1:]
    

## Implement pairing search: For each pairs of (An, Bn), search for a tuplet of (A, B) in the json file which contains the word An in A and the word Bn in B. Return B as the result.
## Save the result in a data list
def findExamples(wordSource = 'output.xlsx', url = 'http://localhost:8983/solr/mycore/select?indent=true&q.op=AND&q='):
    wordData = []
    resultList = []

    
    df = pd.read_excel(wordSource)
    
    print("Start searching")
    for row in tqdm(range(df.shape[0]), desc="Processing rows"):  # Add progress bar
        for col in range(df.shape[1] - 2):
            #print("Row: ", row, "Col: ", col)
            if df.iat[row, col+1] == None:
                print("None")
                continue
            
            # Make sure cell 1, 2 are strings, and they are not empty
            cell1 = str(df.iat[row, col])
            cell2 = str(df.iat[row, col+1])
            #print("Appedning: ", cell1.encode("utf8"), cell2.encode("utf8"))
            if not cell1.strip() or not cell2.strip():
                continue
            # Add cell1 and cell2 to wordData
            wordData.append([cell1, cell2])
    
    http = PoolManager()
    for wordPair in tqdm(wordData, desc="Searching word pairs"):  # Add progress bar
        count = 0
        word1 = wordPair[0]
        #word1 = ' '
        word2 = wordPair[1]
        
        # Replace space with these string below to enable phrasal searches.
        word1Re = word1.replace(' ', '*\n_bahnar:*')
        word2Re = word2.replace(' ', '*\nvietnamese:*')
        # matchVietList = []
        # matchbahnarList = []
        # # CSV file
        # # Search for substring word1 in the vietnamese column, save the corresponding bahnar column to matchVietList
        # r = http.request('GET', f'{url}_bahnar%3A*{word1}*&fl=_bahnar,vietnamese,reference&rows=400&indent=true&wt=csv&useParams=')
            
        # data = csv.reader(r.data.decode('utf-8').split('\n'))
        
        # for row in data:
        #     if row and '_bahnar' not in row:
        #         matchVietList.append(row)                

        # # Search for substring word2 in the bahnar column, save the corresponding vietnamese column to matchbahnarList
        # r = http.request('GET', f'{url}vietnamese%3A*{word2}*&fl=_bahnar,vietnamese,reference&rows=400&indent=true&wt=csv&useParams=')        
        # data = csv.reader(r.data.decode('utf-8').split('\n'))
        
        # for row in data:
        #     if row and '_bahnar' not in row:
        #         matchbahnarList.append(row)

        # # Convert word1 and word2 back to normal
        # word1 = word1.replace('*\n_bahnar:*', ' ')
        # word2 = word2.replace('*\nvietnamese:*', ' ')

        # # Find the intersection of matchVietList and matchbahnarList
        # result = [word1, word2, 'N/a', 'N/a', 'N/a']
        # for key in matchVietList:
        #     if key in matchbahnarList:
        #         result = [word1, word2, key[0], key[1], key[2]]
        #         result[2] = result[2].replace('\\,',',')
        #         result[3] = result[3].replace('\\,',',')
        #         result[4] = result[4].replace('\\,',',')
        #         break
        
        # Add word1 and word2 to the query, AND operator is used to search for both words
        # Format similar to: http://localhost:8983/solr/mycore/select?indent=true&q.op=AND&q=_bahnar%3Aba%2C%20vietnamese%3Al%C3%BAa&rows=100&useParams=&wt=csv
        r = http.request('GET', f'{url}_bahnar%3A{qoute(word1Re)}%2C%20vietnamese%3A{qoute(word2Re)}&rows=500&useParams=&wt=csv&fl=_bahnar,vietnamese,reference')
        data = csv.reader(r.data.decode('utf-8').split('\n'))
        result = [word1, word2, 'N/a', 'N/a', 'N/a']
        for row in data:
            if row and '_bahnar' not in row:
                result = [word1, word2, row[0], row[1], row[2]]
                result[2] = result[2].replace('\\,',',')
                result[3] = result[3].replace('\\,',',')
                result[4] = result[4].replace('\\,',',')
                break
        
        # Append result to resultList
        resultList.append(result)

        # Create an ulitmate list that contains word1, word2 and result list
    return resultList

# Save the result list to a xlsx file
def saveResult(resultList, target='target.xlsx'):
    workbook = xlsx.Workbook(target)
    worksheet = workbook.add_worksheet()
    row = 0
    col = 0

    for result in resultList:
        worksheet.write(row, col, result[0])
        worksheet.write(row, col+1, result[1])
        worksheet.write(row, col+2, result[2])
        worksheet.write(row, col+3, result[3])
        worksheet.write(row, col+4, result[4])
        row += 1
    workbook.close()
    return

def search(library='library/dictionary/Từ điển mức câu_Kon Tum.csv', input='library/dictionary/Từ điển mức từ_Kon Tom (22.9.2024).xlsx', target='output/Output_KonTum.xlsx'):
    deleteQuery()
    uploadToSolr(library)
    resultList = findExamples(input)
    saveResult(resultList, target)
    
search()