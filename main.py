import re
import pandas as pd
import string
from selenium import webdriver
from selenium.webdriver.common.by import By
from nltk.tokenize import sent_tokenize, word_tokenize

def initialize_stop_words(stop_words_files):
    stop_words = set()
    for stop_file in stop_words_files:
        with open(stop_file, "r") as sf:
            for line in sf:
                stop_words.add(line.strip())
    return stop_words

def read_content_from_url(driver, url):
    driver.get(url)
    article_title = driver.find_element(By.XPATH, '//h1[@class="entry-title"]').text
    article_texts = driver.find_element(By.XPATH, '//div[@class="td-post-content tagdiv-type"]').text
    return article_title, article_texts

def clean_text(text, stop_words):
    words_in_text = text.split()
    remaining_words = [word for word in words_in_text if word.lower() not in stop_words]
    return " ".join(remaining_words)

def syllable_count(word):
    count = 0
    vowels = "aeiouy"
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
            if word.endswith("e"):
                count -= 1
    if count == 0:
        count += 1
    return count

def is_complex(word):
    return syllable_count(word) >= 3

def count_syllables(word):
    
    if len(word) == 2 and word.isupper():
        return 2

    word = word.lower()   
    word = re.sub(r'[^a-z]', '', word)
    
    vowels = "aeiouy"
    syllables = 0

    vowel_groups = re.findall(r'[aeiouy]+', word)
    syllables += len(vowel_groups)
    
    if word.endswith("e") and len(word) > 1 and word[-2] not in vowels and not word.endswith("le"):
        syllables -= 1

    if word.endswith("le") and len(word) > 2 and word[-3] not in vowels:
        syllables += 1

    if word.endswith("es") and len(word) > 2 and word[-3] not in vowels:
        syllables -= 1
    elif word.endswith("ed") and len(word) > 2 and word[-3] not in vowels:
        syllables -= 1

    if syllables == 0:
        syllables = 1

    return syllables

def process_text_file(file_path):
    syllable_counts = {}
    words = re.findall(r'\b\w+\b', file_path)
    for word in words:
        syllable_counts[word] = count_syllables(word)
    return syllable_counts


def syl_count(file_path):
    syllable_counts = process_text_file(file_path)
    return syllable_counts

def analyze_text(file_content, positive_words, negative_words, stop_words):
    remaining_words = clean_text(file_content, stop_words)
    tokens = word_tokenize(remaining_words)
    
    positive_score = sum(1 for word in tokens if word in positive_words)
    negative_score = sum(1 for word in tokens if word in negative_words)
    
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(tokens) + 0.000001)
    
    sentences = sent_tokenize(file_content)
    words = word_tokenize(file_content)
    
    average_sentence_length = len(words) / len(sentences)
    
    complex_words = [word for word in words if is_complex(word)]
    percentage_complex_words = len(complex_words) / len(words)
    
    fog_index = 0.4 * (average_sentence_length + (percentage_complex_words * 100))
    
    average_number_of_words_per_sentence = len(words) / len(sentences)
    
    complex_word_count = len(complex_words)
    
    word_count = len(remaining_words.split())
    
    pronoun_pattern = re.compile(r'\b(I|we|my|ours|us)\b', re.IGNORECASE)
    pronouns_found = pronoun_pattern.findall(file_content)
    personal_pronoun_count = sum(1 for pronoun in pronouns_found if pronoun.lower() != "us")

    syllables_per_word = syl_count(file_content)
    
    cleaned_words = [word for word in words if word not in string.punctuation]
    total_characters = sum(len(word) for word in cleaned_words)
    total_words = len(cleaned_words)
    average_word_length = total_characters / total_words if total_words > 0 else 0
    
    return {
        'Positive Score': positive_score,
        'Negative Score': -negative_score,
        'Polarity Score': polarity_score,
        'Subjectivity Score': subjectivity_score,
        'Avg Sentence Length': average_sentence_length,
        'Percentage Complex Words': percentage_complex_words * 100,
        'Fog Index': fog_index,
        'Avg Words Per Sentence': average_number_of_words_per_sentence,
        'Complex Word Count': complex_word_count,
        'Word Count': word_count,
        'Syllables Per Word': syllables_per_word,
        'Personal Pronouns': personal_pronoun_count,
        'Avg Word Length': average_word_length
    }

stop_words_files = [
    "StopWords/StopWords_Auditor.txt",
    "StopWords/StopWords_Currencies.txt",
    "StopWords/StopWords_DatesandNumbers.txt",
    "StopWords/StopWords_Generic.txt",
    "StopWords/StopWords_GenericLong.txt",
    "StopWords/StopWords_Geographic.txt",
    "StopWords/StopWords_Names.txt"
]
stop_words = initialize_stop_words(stop_words_files)

with open("D:/Web Scrapping/MasterDictionary/negative-words.txt", "r") as file:
    negative_words = file.read().split()
with open("D:/Web Scrapping/MasterDictionary/positive-words.txt", "r") as file:
    positive_words = file.read().split()

filtered_positive_words = [word for word in positive_words if word not in stop_words]
filtered_negative_words = [word for word in negative_words if word not in stop_words]

input_data = pd.read_excel('D:\Web Scrapping\Input.xlsx')
url_ids = input_data['URL_ID']
urls = input_data['URL']

driver = webdriver.Chrome()
driver.maximize_window()

results = []

for url_id, url in zip(url_ids, urls):
    try:
        article_title, article_texts = read_content_from_url(driver, url)
        file_content = article_title + "\n\n\n\n\n\n" + article_texts
        analysis_result = analyze_text(file_content, filtered_positive_words, filtered_negative_words, stop_words)
        
        result = {
            'URL_ID': url_id,
            'URL': url,
            **analysis_result
        }
        results.append(result)
    except Exception as e:
        print(f"Error processing URL {url}: {e}")

driver.quit()

results_df = pd.DataFrame(results)
results_df.to_excel('output.xlsx', index=False)

print("Analysis complete. Results saved to output.xlsx.")
