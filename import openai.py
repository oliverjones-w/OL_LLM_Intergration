from openai import OpenAI

client = OpenAI(api_key=api_key)
import pandas as pd
from fuzzywuzzy import fuzz
import re

# Function to read API key from a file
def read_api_key(file_path):
    with open(file_path, 'r') as file:
        return file.read().strip()

# Read the OpenAI API key
api_key = read_api_key('../OpenAI_API_Key.txt')

# Define a mapping of common nicknames to formal names
nickname_mapping = {
    'Charlie': 'Charles',
    'Chuck': 'Charles',
    'Bill': 'William',
    'Will': 'William',
    'Billy': 'William',
    'Bob': 'Robert',
    'Rob': 'Robert',
    'Bobby': 'Robert',
    'Rich': 'Richard',
    'Dick': 'Richard',
    # Add more mappings as needed
}

def standardize_name(name):
    # Remove content within parentheses
    name = re.sub(r'\(.*?\)', '', name).strip()
    # Replace nicknames with formal names
    name_parts = name.split()
    for i, part in enumerate(name_parts):
        if part in nickname_mapping:
            name_parts[i] = nickname_mapping[part]
    return ' '.join(name_parts)

def categorize_data(unstructured_text):
    response = client.completions.create(engine="text-davinci-003",
    prompt=f"Extract the following information from the text and assign to these categories: Name, Title, Strategy, Products, Firm, Region: \"{unstructured_text}\"",
    max_tokens=100,
    n=1,
    stop=None,
    temperature=0)
    extracted_text = response.choices[0].text.strip()
    return extracted_text

def parse_extracted_text(extracted_text):
    categories = ["Name", "Title", "Strategy", "Products", "Firm", "Region"]
    data = {}
    for category in categories:
        data[category] = ""
        for line in extracted_text.split("\n"):
            if line.startswith(category):
                data[category] = line.split(":")[1].strip()
    return data

# Load the Excel files
df1 = pd.read_excel('path_to_first_excel_file.xlsx')
df2 = pd.read_excel('path_to_second_excel_file.xlsx')

# Apply GPT-4 categorization and parse the results
df1_structured = []
for unstructured_text in df1['Unstructured'].tolist():
    extracted_text = categorize_data(unstructured_text)
    structured_data = parse_extracted_text(extracted_text)
    df1_structured.append(structured_data)

df1_structured = pd.DataFrame(df1_structured)

# Now, let's apply the same categorization to the second dataset
df2_structured = []
for unstructured_text in df2['Unstructured'].tolist():
    extracted_text = categorize_data(unstructured_text)
    structured_data = parse_extracted_text(extracted_text)
    df2_structured.append(structured_data)

df2_structured = pd.DataFrame(df2_structured)

def calculate_match_score(row1, row2):
    name1 = standardize_name(row1['Name'])
    name2 = standardize_name(row2['Name'])
    name_score = fuzz.token_sort_ratio(name1, name2)
    firm_score = fuzz.token_sort_ratio(row1['Firm'], row2['Firm'])
    title_score = fuzz.token_sort_ratio(row1['Title'], row2['Title'])
    strategy_score = fuzz.token_sort_ratio(row1['Strategy'], row2['Strategy'])
    products_score = fuzz.token_sort_ratio(row1['Products'], row2['Products'])
    region_score = fuzz.token_sort_ratio(row1['Region'], row2['Region'])
    return (name_score, firm_score, title_score, strategy_score, products_score, region_score)

def overall_score(scores, weights):
    return sum(s * w for s, w in zip(scores, weights)) / sum(weights)

# Weights for each field
weights = [0.3, 0.1, 0.2, 0.15, 0.15, 0.1]  # Adjust these weights as needed

matches = []
for _, row1 in df1_structured.iterrows():
    best_match = None
    best_score = 0
    for _, row2 in df2_structured.iterrows():
        scores = calculate_match_score(row1, row2)
        score = overall_score(scores, weights)
        if score > best_score:
            best_score = score
            best_match = {**row2, 'Similarity Score': score}
    matches.append({**row1, **best_match})

# Convert the matches to a DataFrame for better visualization
match_df = pd.DataFrame(matches)

# Save the results to a new Excel file
match_df.to_excel('matched_names_structured.xlsx', index=False)
