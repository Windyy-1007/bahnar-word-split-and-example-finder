import json
import sys

def translate_sentence(sentence, dictionary):
    words = sentence.split()
    translated_words = []
    for word in words:
        if word in dictionary:
            translated_words.append(dictionary[word]['vi_words'][0])
            # If word contain utf-8 words like ŭ and cannot be found in the dictionary, change it become u and search again
        elif 'ŭ' in word or 'ŭ' in word:
            word = word.replace('ŭ', 'u')
            translated_words.append(dictionary[word]['vi_words'][0])
        else:
            translated_words.append(word)
    return ' '.join(translated_words)

def load_dictionary(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        dictionary = json.load(file)
    return dictionary

if __name__ == "__main__":
    dictionary = load_dictionary('library/bana_to_viet.json')
    sentence = "'năr tơ muôh kăn hơ trŭh"
    translated_sentence = translate_sentence(sentence, dictionary)
    sys.stdout.reconfigure(encoding='utf-8')
    print(translated_sentence) 

