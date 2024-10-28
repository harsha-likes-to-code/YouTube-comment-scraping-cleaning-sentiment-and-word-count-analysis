import pandas as pd
from collections import Counter
import spacy

# Load the English model
nlp = spacy.load("en_core_web_sm")

# Step 1: Load the Cleaned Comments Data
df = pd.read_excel('cleaned_comments_with_sentiment.xlsx')  # Load the cleaned comments Excel file
cleaned_comments = df['Cleaned Comments'].tolist()  # Get the cleaned comments as a list

# Step 2: Initialize a Counter for Word Counts
word_count = Counter()

# Step 3: Count Nouns
for comment in cleaned_comments:
    # Ensure the comment is a valid string
    if isinstance(comment, str):
        doc = nlp(comment)  # Process the comment with spaCy
        nouns = [token.text for token in doc if token.pos_ == "NOUN"]  # Keep only nouns
        word_count.update(nouns)  # Update the counter with the nouns from the comment
    else:
        print(f'Non-string comment encountered: {comment}')  # Log non-string comments

# Step 4: Convert Word Count to DataFrame
word_count_df = pd.DataFrame(word_count.items(), columns=['Word', 'Count'])

# Step 5: Filter out words with a count of 1 and common stop words
stop_words = set([
    'the', 'and', 'a', 'to', 'of', 'in', 'for', 'is', 'on', 'that', 
    'it', 'with', 'as', 'was', 'this', 'at', 'by', 'an', 'be', 'are', 
    'not', 'or', 'if', 'from', 'but', 'we', 'they', 'you', 'their', 
    'he', 'she', 'him', 'her', 'my', 'me', 'your', 'his', 'its', 
    'our', 'us', 'what', 'which', 'who', 'whom', 'when', 'where', 
    'why', 'how', 'can', 'could', 'would', 'should', 'may', 'might', 'href', 'd'
])

# Filter out stop words and words with a count of 1
filtered_word_count_df = word_count_df[~word_count_df['Word'].isin(stop_words) & (word_count_df['Count'] > 1)]

# Step 6: Save the Filtered Word Count DataFrame to an Excel File
filtered_word_count_df.sort_values(by='Count', ascending=False, inplace=True)  # Sort by count
filtered_word_count_df.to_excel('filtered_noun_count_analysis.xlsx', index=False)

# Step 7: Print Summary
print(f'Total unique nouns before filtering: {len(word_count)}')
print(f'Total unique nouns after filtering: {len(filtered_word_count_df)}')
print(f'Saved filtered noun count analysis to "filtered_noun_count_analysis.xlsx"')





