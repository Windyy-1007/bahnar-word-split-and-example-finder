import json

# Load VietBana JSON data
with open('VietBana.json', 'r', encoding='utf-8') as f:
    viet_bana_data = json.load(f)

# Read VietDict.txt file
with open('VietDict.txt', 'r', encoding='utf-8') as f:
    viet_dict_words = f.read().splitlines()

# Count matching Viet words
count = 0
for word in viet_dict_words:
    if word in viet_bana_data:
        count += 1

# Print the count
print(f"Number of matching Viet words: {count}")
