import pandas as pd
import spacy
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
from newspaper import Article
from openpyxl import Workbook
import requests
import re
from textblob import TextBlob
from textstat import textstat


# Read the input Excel file
input_file = "Input.xlsx"
df = pd.read_excel(input_file)

# Create an Excel workbook for the output
output_file = "Output Data Structure.xlsx"
wb = Workbook()
ws = wb.active

# Define the headers for the output
headers = [
    "URL_ID",
    "URL",
    "Title",
    "POSITIVE SCORE",
    "NEGATIVE SCORE",
    "POLARITY SCORE",
    "SUBJECTIVITY SCORE",
    "AVG SENTENCE LENGTH",
    "PERCENTAGE OF COMPLEX WORDS",
    "FOG INDEX",
    "AVG NUMBER OF WORDS PER SENTENCE",
    "COMPLEX WORD COUNT",
    "WORD COUNT",
    "SYLLABLE PER WORD",
    "PERSONAL PRONOUNS",
    "AVG WORD LENGTH",
]

# Write the headers to the output file
ws.append(headers)

# Load spaCy's English tokenizer
nlp = spacy.load("en_core_web_sm")

# Initialize the VADER sentiment analyzer
analyzer = SentimentIntensityAnalyzer()

# Function to analyze text
def analyze_text(text):
    doc = nlp(text)

    positive_score = 0
    negative_score = 0
    avg_sentence_length = 0
    percentage_complex_words = 0
    fog_index = 0
    avg_words_per_sentence = 0
    complex_word_count = 0
    word_count = 0
    syllable_per_word = 0
    personal_pronouns = 0
    avg_word_length = 0

    for sentence in doc.sents:
        sentiment = analyzer.polarity_scores(sentence.text)
        positive_score += sentiment['pos']
        negative_score += sentiment['neg']
        avg_sentence_length += len(sentence) / len(list(doc.sents))
        words = [token.text for token in sentence if not token.is_punct]
        word_count += len(words)
        avg_words_per_sentence += len(words)
        for word in words:
            syllable_per_word += textstat.syllable_count(word)
            avg_word_length += len(word)
            if len(word) > 2:
                complex_word_count += 1
        personal_pronouns += len(re.findall(r'\b(I|we|my|ours|us)\b', sentence.text, re.IGNORECASE))

    avg_sentence_length = round(avg_sentence_length, 2)
    avg_words_per_sentence = round(avg_words_per_sentence / len(list(doc.sents)), 2)

    percentage_complex_words = round(complex_word_count / word_count, 2)
    fog_index = round(0.4 * (avg_sentence_length + percentage_complex_words), 2)
    avg_word_length = round(avg_word_length / word_count, 2)
    syllable_per_word = round(syllable_per_word / word_count, 2)

    polarity_score = (positive_score - negative_score) / (positive_score + negative_score + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (word_count + 0.000001)

    return [
        positive_score,
        negative_score,
        polarity_score,
        subjectivity_score,
        avg_sentence_length,
        percentage_complex_words,
        fog_index,
        avg_words_per_sentence,
        complex_word_count,
        word_count,
        syllable_per_word,
        personal_pronouns,
        avg_word_length,
    ]

# Initialize a list to store URLs that encountered errors
error_urls = []

# Iterate through URLs and analyze the text
for index, row in df.iterrows():
    url = str(row["URL"])  # Convert the URL to a string
    try:
        response = requests.get(url)
        if response.status_code != 200:
            error_urls.append(url)
            print(f"Error: Unable to retrieve URL {url}")
            continue  # Skip this URL and continue with the next one

        article = Article(url)
        article.download()
        article.parse()
        text = article.text

        results = analyze_text(text)

        # Append the URL and analysis results to the output file
        row_data = [row["URL_ID"], url, article.title] + results
        ws.append(row_data)

        # Add a print statement for each successful URL analysis
        print(f"Analysis completed for URL {url}")

    except Exception as e:
        error_urls.append(url)
        print(f"Error: {str(e)} for URL {url}")
        continue  # Skip this URL and continue with the next one

# Save the output file
wb.save(output_file)

print(f"Analysis completed. Results saved to {output_file}")
print(f"URLs with errors: {error_urls}")




