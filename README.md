# URL-Text-Analysis-and-Sentiment-Scoring-Project

This Python script is designed to analyze the content of web pages from a given list of URLs, perform sentiment analysis, and calculate various text metrics. The analysis includes sentiment scoring, readability measurements, and linguistic features. The results are then stored in an Excel file for further analysis and reporting.

Table of Contents
Introduction
Prerequisites
Getting Started
Code Description
Results
License
Introduction
This code performs the following tasks:

Reads a list of URLs and corresponding URL IDs from an input Excel file (Input.xlsx).
Analyzes the content of each URL, extracts text, and performs a series of text analyses, including:
Sentiment analysis using VADER Sentiment Analyzer.
Calculation of readability metrics, such as average sentence length and Fog Index.
Analysis of linguistic features like the percentage of complex words and personal pronouns.
Stores the results in an Excel file (Output Data Structure.xlsx) for further analysis and reporting.
Handles errors when retrieving URLs and continues the analysis for other URLs.
Prerequisites
Before running the code, ensure you have the following prerequisites in place:

Python 3.x
Required Python libraries: pandas, spacy, vaderSentiment, newspaper3k, openpyxl, requests, re, textblob, textstat
Getting Started
Clone or download this repository to your local machine.

Create an input Excel file (Input.xlsx) with the following columns:

"URL_ID" (unique identifier for each URL)
"URL" (the URL of the web page to analyze)
Open a terminal or command prompt.

Navigate to the directory containing the script.

Run the following command to execute the code:

Copy code
python url_text_analysis.py
Make sure to adjust the input file and output file paths if necessary.

Code Description
The code file (url_text_analysis.py) is organized as follows:

Import necessary Python libraries for text analysis, sentiment scoring, and web scraping.
Read the input Excel file containing a list of URLs.
Create an Excel workbook for the output and define the output headers.
Define a function to analyze text and calculate various text metrics.
Iterate through the URLs, analyze the content, and store the results in the output Excel file.
Handle errors when retrieving URLs and continue the analysis for other URLs.

Results
After running the code, you will have an Excel file (Output Data Structure.xlsx) containing the results of text analysis and sentiment scoring for the provided URLs. The file will include various metrics and scores for each URL.

