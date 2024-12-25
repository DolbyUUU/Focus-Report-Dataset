import pandas as pd
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import jieba
import re

# Load the data
df = pd.read_excel('focus_report_index.xlsx')

# Ensure 'DATE' column is datetime
df['DATE'] = pd.to_datetime(df['DATE'])

# Set 'DATE' column as index
df.set_index('DATE', inplace=True)

# Define the wordcloud
wordcloud = WordCloud(font_path='C:/Windows.old/Windows/Fonts/simhei.ttf', background_color="white", width=800, height=400)

# Load the Chinese stopwords
with open('chinese_stopwords.txt', 'r', encoding='utf8') as f:
    stopwords = [line.strip() for line in f]

# Function to remove specific strings
def remove_specific_strings(text):
    patterns = ["（一）", "（二）", "（三）", "（四）", "（五）", "（六）", "（七）", "（八）", "（九）", "（十）", "（上）", "（下）"]
    for pattern in patterns:
        text = re.sub(pattern, '', str(text))
    return text

# Loop through each year
for year in range(2003, 2024):
    # Filter rows for the current year
    yearly_data = df[df.index.year == year]
    
    # Join all the titles for the current year, converting all elements to strings first
    yearly_text = ' '.join(map(str, yearly_data['TITLE']))

    # Remove specific strings from the text
    yearly_text = remove_specific_strings(yearly_text)

    # Only generate a word cloud if yearly_text is not empty
    if yearly_text.strip():
        # Segment the text
        seg_list = jieba.cut(yearly_text, cut_all=False)
        
        # Filter out stopwords and punctuation
        seg_list = [word for word in seg_list if word not in stopwords and not word.isspace()]

        # Generate word cloud
        wordcloud.generate(' '.join(seg_list))
    
        # Save the word cloud as an image file
        wordcloud.to_file(f'wordcloud_{year}.png')
    else:
        print(f"No data available for the year {year}. Skipped word cloud generation.")