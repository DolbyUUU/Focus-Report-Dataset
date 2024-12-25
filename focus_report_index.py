import pandas as pd
from snownlp import SnowNLP
import matplotlib.pyplot as plt
import re
from scipy.stats import pearsonr

# Read data from Excel file
my_file = 'focus_report_data.xlsx'
df = pd.read_excel(my_file)

# Function to preprocess and clean text
def clean_text_cn(text):
    # Define patterns to remove
    patterns = ["（一）", "（二）", "（三）", "（四）", "（五）", "（六）", "（七）", "（八）", "（九）", "（十）", "（上）", "（下）"]
    for pattern in patterns:
        text = re.sub(pattern, '', str(text))
    # Remove punctuation and numbers
    text = re.sub('[^\u4e00-\u9fa5]', '', str(text))
    return text

# Apply text cleaning
df['cleaned_TITLE'] = df['TITLE'].apply(clean_text_cn)

# Function to get sentiment
def get_sentiment_cn(text):
    if pd.isnull(text) or text.strip() == "":
        return None
    return SnowNLP(text).sentiments

# Create a new column "INDEX" that applies the function on the cleaned "TITLE" column
df['INDEX'] = df['cleaned_TITLE'].apply(get_sentiment_cn)

# Remove rows with INDEX being None or 0
df = df[df['INDEX'].notna() & (df['INDEX'] != 0)]

# Sort the DataFrame by the "INDEX" column in descending order
df_sorted = df.sort_values(by='INDEX', ascending=False)

# Print the 10 news titles with highest sentiment index
print("10 news titles with highest sentiment index:")
print(df_sorted[['TITLE', 'INDEX']].head(10))

# Sort the DataFrame by the "INDEX" column in ascending order for lowest sentiment index
df_sorted_asc = df.sort_values(by='INDEX', ascending=True)

# Print the 10 news titles with lowest non-zero sentiment index
print("10 news titles with lowest non-zero sentiment index:")
print(df_sorted_asc[['TITLE', 'INDEX']].head(10))

# Write the DataFrame to the Excel file
df.to_excel('focus_report_index.xlsx', index=False)


# Calculate the length of each title and store it in a new column
df['title_length'] = df['cleaned_TITLE'].apply(lambda x: len(x))

# Remove rows with INDEX being None or 0
df = df[df['INDEX'].notna() & (df['INDEX'] != 0)]

# Calculate the correlation between the length of the title and the sentiment index
correlation = df['title_length'].corr(df['INDEX'])
print(f"Correlation between length of title and sentiment index: {correlation}")

# If you want the p-value as well, you can use the pearsonr function from scipy.stats
corr, p_value = pearsonr(df['title_length'], df['INDEX'])
print(f"Correlation is {corr} with a p-value of {p_value}")