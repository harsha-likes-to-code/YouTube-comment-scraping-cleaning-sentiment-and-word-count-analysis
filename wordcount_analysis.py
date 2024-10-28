import pandas as pd
from collections import Counter
import re

# Step 1: Load the Cleaned Comments Data
df = pd.read_excel('cleaned_comments_with_sentiment.xlsx')  # Load the cleaned comments Excel file
cleaned_comments = df['Cleaned Comments'].tolist()  # Get the cleaned comments as a list

# Step 2: Initialize a Counter for Word Counts
word_count = Counter()

# Step 3: Count Words
for comment in cleaned_comments:
    # Ensure the comment is a valid string
    if isinstance(comment, str):
        words = re.findall(r'\b\w+\b', comment)  # Extract words using regex
        word_count.update(words)  # Update the counter with the words from the comment
    else:
        print(f'Non-string comment encountered: {comment}')  # Log non-string comments

# Step 4: Convert Word Count to DataFrame
word_count_df = pd.DataFrame(word_count.items(), columns=['Word', 'Count'])

# Step 5: Save the Word Count DataFrame to an Excel File
word_count_df.sort_values(by='Count', ascending=False, inplace=True)  # Sort by count
word_count_df.to_excel('word_count_analysis.xlsx', index=False)

# Step 6: Print Summary
print(f'Total unique words: {len(word_count)}')
print(f'Saved word count analysis to "word_count_analysis.xlsx"')

