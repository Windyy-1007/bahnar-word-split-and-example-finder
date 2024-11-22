import sys
import json

# Set console encoding to UTF-8
sys.stdout.reconfigure(encoding='utf-8')

# Allow utf-8 encoding
with open('library/bana_to_viet.json', 'r', encoding='utf-8') as f:
    viet_bana_data = json.load(f)
    
# Print first 5 items in the dictionary

# From viet_bana_data, create a two column dictionary (array)
dict_arr =[
    {'Bana', 'Viet'},
]

for key, value in viet_bana_data.items():
    dict_arr.append([key, value])
    

# Print first 5 items in the dictionary array
print(dict_arr[:5])