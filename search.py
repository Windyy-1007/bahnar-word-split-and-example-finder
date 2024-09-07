from urllib3 import *
import json
import pandas as pd
import xlsxwriter as xlsx

# Tasks
## Convert target xlsx file to json file

## Upload json file to the solr query

## Implement pairing search: For each pairs of (An, Bn), search for a tuplet of (A, B) in the json file which contains the word An in A and the word Bn in B. Return B as the result.
## Save the result in a data list

## Save the result in the third column of the existing xlsx file