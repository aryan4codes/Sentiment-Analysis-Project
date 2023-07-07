import requests
from bs4 import BeautifulSoup
import re
import nltk
from nltk.corpus import stopwords
import string
import pandas as pd

nltk.download('stopwords')

def extract_article_text(url):
    # Make a request to the URL
    response = requests.get(url)
    
    # Create a BeautifulSoup object
    soup = BeautifulSoup(response.content, 'html.parser')

    # Extract the article title
    title = soup.find('title').get_text().strip()
    
    # Extract the article text
    article_text = ""
    for paragraph in soup.find_all('p'):
        text = paragraph.get_text().strip()
        if text:
            article_text += text + " "

    return title, article_text

def calculate_metrics(article_text):
    # Remove stopwords
    stop_words_files = ['StopWords_Auditor.txt', 'StopWords_Currencies.txt', 'StopWords_DatesandNumbers.txt', 'StopWords_Generic.txt', 'StopWords_GenericLong.txt', 'StopWords_Geographic.txt', 'StopWords_Names.txt']

    stop_words = set()

    # Load stop words from each file
    for file in stop_words_files:
        with open(file, 'r') as f:
           stop_words.update(line.strip() for line in f)

    words = article_text.split()
    cleaned_words = [word.lower() for word in words if word.lower() not in stop_words]

    # Remove punctuation
    cleaned_words = [word.translate(str.maketrans('', '', string.punctuation)) for word in cleaned_words]

    # Calculate positive score
    positive_words = set(line.strip() for line in open('positive-words.txt'))
    positive_score = sum(1 for word in cleaned_words if word in positive_words)

    # Calculate negative score
    negative_words = set(line.strip() for line in open('negative-words.txt'))
    negative_score = abs(sum(1 for word in cleaned_words if word in negative_words)) 

    # Calculate polarity score and subjectivity score using the provided formulas
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(cleaned_words) + 0.000001)

    # Calculate average sentence length
    sentences = re.split(r'[.!?]+', article_text)
    avg_sentence_length = sum(len(s.split()) for s in sentences) / len(sentences)

    # Calculate percentage of complex words
    complex_word_count = len([w for w in cleaned_words if syllable_count(w) > 2]) # no of complex words
    percentage_complex_words = (complex_word_count / len(cleaned_words)) * 100

    # Calculate FOG index
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)

    # Calculate average number of words per sentence
    avg_words_per_sentence = len(cleaned_words) / len(sentences)

    # Calculate word count
    word_count = len(cleaned_words)

    # Calculate syllables per word
    syllables_per_word = sum(syllable_count(w) for w in cleaned_words) / len(cleaned_words)

    # Calculate personal pronouns
    personal_pronouns = len([w for w in cleaned_words if w.lower() in ['i', 'me', 'my', 'mine', 'myself', 'we', 'ours','us']])

    # Calculate average word length
    avg_word_length = sum(len(w) for w in cleaned_words) / len(cleaned_words)

    return positive_score, negative_score, polarity_score, subjectivity_score, avg_sentence_length, percentage_complex_words, fog_index, avg_words_per_sentence, complex_word_count, word_count, syllables_per_word, personal_pronouns, avg_word_length

def syllable_count(word):
    word = word.lower()
    if len(word) <= 3:
        return 1
    count = 0
    vowels = 'aeiouy'
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
    if word.endswith('e'):
        count -= 1
    if word.endswith('le') and len(word) > 2 and word[-3] not in vowels:
        count += 1
    return count


# Read URLs from Excel sheet
df = pd.read_excel('Input.xlsx', sheet_name='Sheet1')
urls = df['URL'].tolist()

# Create lists to store the results
titles = []
positive_scores = []
negative_scores = []
polarity_scores = []
subjectivity_scores = []
avg_sentence_lengths = []
percentage_complex_words = []
fog_indices = []
avg_words_per_sentences = []
complex_word_counts = []
word_counts = []
syllables_per_words = []
personal_pronouns = []
avg_word_lengths = []

# Process each URL
for url in urls:
    # Extract article text
    title, article_text = extract_article_text(url)
    
    # Calculate metrics
    metrics = calculate_metrics(article_text)
    
    # Append the results to the respective lists
    titles.append(url)
    positive_scores.append(metrics[0])
    negative_scores.append(metrics[1])
    polarity_scores.append(metrics[2])
    subjectivity_scores.append(metrics[3])
    avg_sentence_lengths.append(metrics[4])
    percentage_complex_words.append(metrics[5])
    fog_indices.append(metrics[6])
    avg_words_per_sentences.append(metrics[7])
    complex_word_counts.append(metrics[8])
    word_counts.append(metrics[9])
    syllables_per_words.append(metrics[10])
    personal_pronouns.append(metrics[11])
    avg_word_lengths.append(metrics[12])

# Create a DataFrame from the results
data = {
    'URLs': titles,
    'Positive Score': positive_scores,
    'Negative Score': negative_scores,
    'Polarity Score': polarity_scores,
    'Subjectivity Score': subjectivity_scores,
    'Average Sentence Length': avg_sentence_lengths,
    'Percentage of Complex Words': percentage_complex_words,
    'FOG Index': fog_indices,
    'Average Words per Sentence': avg_words_per_sentences,
    'Complex Word Count': complex_word_counts,
    'Word Count': word_counts,
    'Syllables per Word': syllables_per_words,
    'Personal Pronouns': personal_pronouns,
    'Average Word Length': avg_word_lengths
}
df_output = pd.DataFrame(data)

# Write the DataFrame to an Excel file
df_output.to_excel('Output Data Structure - Aryan Rajpurkar.xlsx', index=False, sheet_name='Sheet1')
