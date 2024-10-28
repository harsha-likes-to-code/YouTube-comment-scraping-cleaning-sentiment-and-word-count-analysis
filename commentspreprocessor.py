# Step 1: Import Necessary Libraries
from textblob import TextBlob
import pandas as pd
import re
from googletrans import Translator

# Step 2: Initialize the Translator
translator = Translator()

# Step 3: Define the clean_comments Function
def clean_comments(comments):
    cleaned_comments = []
    for comment in comments:
        # Ensure the comment is a string
        if isinstance(comment, str):
            # Remove URLs
            comment = re.sub(r'http\S+|www\S+|https\S+', '', comment, flags=re.MULTILINE)
            # Remove special characters and numbers
            comment = re.sub(r'[^a-zA-Z\s]', '', comment)
            # Convert to lowercase
            comment = comment.lower()
            cleaned_comments.append(comment)
        else:
            cleaned_comments.append('')  # Append an empty string for non-string comments
    return cleaned_comments

# Step 4: Define the get_sentiment_scores Function
def get_sentiment_scores(comments):
    positive_scores = []
    negative_scores = []
    sentiment_scores = []
    for comment in comments:
        blob = TextBlob(comment)
        sentiment_score = blob.sentiment.polarity
        sentiment_scores.append(sentiment_score)  # Store the sentiment score for each comment

        if sentiment_score > 0:
            positive_scores.append(sentiment_score)
        elif sentiment_score < 0:
            negative_scores.append(sentiment_score)

    return sentiment_scores, positive_scores, negative_scores

# Step 5: Load Your Comments Data
df = pd.read_excel('processed_comments.xlsx')  # Load the cleaned comments Excel file
df.dropna(subset=['Processed_Comments'], inplace=True)  # Drop rows with NaN in 'Processed_Comments'
comments = df['Processed_Comments'].tolist()  # Get the comments as a list

# Step 6: Translate Comments to English
translated_comments = []
for comment in comments:
    if isinstance(comment, str) and comment.strip():  # Check if comment is non-empty string
        try:
            translated = translator.translate(comment, dest='en')
            if hasattr(translated, 'text'):
                translated_comments.append(translated.text)
            else:
                print(f"Translation returned unexpected result for comment '{comment}'")
                translated_comments.append(comment)  # Keep the original comment if translation fails
        except Exception as e:
            print(f"Translation error for comment '{comment}': {e}")
            translated_comments.append(comment)  # Keep the original comment if translation fails
    else:
        print(f"Non-string or empty comment encountered: {comment}")
        translated_comments.append('')  # Append an empty string for non-string or empty comments

# Step 7: Clean Comments
cleaned_comments = clean_comments(translated_comments)

# Step 8: Get Sentiment Scores
sentiment_scores, positive_scores, negative_scores = get_sentiment_scores(cleaned_comments)

# Step 9: Create a DataFrame to Store Results
results_df = pd.DataFrame({
    'Original Comments': comments,
    'Translated Comments': translated_comments,
    'Cleaned Comments': cleaned_comments,
    'Sentiment Scores': sentiment_scores
})

# Step 10: Save to Excel
results_df.to_excel('cleaned_comments_with_sentiment.xlsx', index=False)

# Step 11: Print Summary
print(f'Number of Positive Comments: {len(positive_scores)}')
print(f'Number of Negative Comments: {len(negative_scores)}')
print(f'Positive Scores: {positive_scores}')
print(f'Negative Scores: {negative_scores}')








