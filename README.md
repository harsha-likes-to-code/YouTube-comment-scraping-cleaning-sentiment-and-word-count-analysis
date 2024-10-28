# YouTube Comments Analysis on Venom Movie Trailer

## Overview
This project involves scraping and analyzing comments from the trailer of the new Venom movie. The analysis focuses on cleaning the comments, translating them to English, performing sentiment analysis, and conducting a word count analysis. The final results are stored in Excel files for further examination.

## Project Structure
- `commentspreprocessor.py`: This script handles the preprocessing of the comments, including:
  - Cleaning the comments (removing URLs, special characters, and converting to lowercase).
  - Translating comments to English using the Google Translate API.
  - Performing sentiment analysis using the TextBlob library.
  - Saving the cleaned comments and sentiment scores in an Excel file.

- `wordcount_analysis.py`: This script performs word count analysis on the cleaned comments and stores the results in a separate Excel file.

- `processed_comments.xlsx`: This is the input file containing the raw comments collected from the trailer.

- `cleaned_comments_with_sentiment.xlsx`: This is the output file that contains the original comments, translated comments, cleaned comments, and their corresponding sentiment scores.

- `word_count_results.xlsx`: This is the output file that contains the word count analysis results.

## Requirements
To run this project, ensure you have the following libraries installed:
- `pandas`
- `textblob`
- `googletrans`
- `openpyxl`
- `re`

You can install these libraries using pip:

```bash
pip install pandas textblob googletrans openpyxl
