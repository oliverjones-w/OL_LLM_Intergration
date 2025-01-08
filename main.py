import os
from openai import OpenAI
import pandas as pd
from fuzzywuzzy import fuzz
import re

# Initialize the OpenAI client
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

if not client.api_key:
    raise ValueError("The OPENAI_API_KEY environment variable is not set.")
else:
    print(f"API Key is set")

# Function to interact with OpenAI API and print raw response
def categorize_data(unstructured_text):
    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {
                "role": "user",
                "content": f"Extract the following information from the text and assign to these categories: Name, Title, Investment Strategy, Financial Products, Firm, Region, Location: \"{unstructured_text}\"",
            }
        ]
    )
    # Print the raw response (remove "#" to print raw response)
    #print("Raw response from API:")
    #print(response.model_dump_json(indent=2))
    
    # Extract the message content
    extracted_text = response.choices[0].message.content.strip()
    return extracted_text

# Function to parse the extracted text
def parse_extracted_text(extracted_text):
    categories = ["Name", "Title", "Investment Strategy", "Financial Products", "Firm", "Region", "Location"]
    data = {}
    for category in categories:
        data[category] = ""
        for line in extracted_text.split("\n"):
            if line.startswith(category):
                data[category] = line.split(":")[1].strip()
    return data

# Function to test the categorize_data function in the terminal
def test_categorize_data():
    while True:
        prompt = input("Enter the text to categorize (or 'exit' to quit): ")
        if prompt.lower() == 'exit':
            break
        extracted_text = categorize_data(prompt)
        parsed_data = parse_extracted_text(extracted_text)
        print("Extracted Data:")
        for key, value in parsed_data.items():
            print(f"{key}: {value}")

if __name__ == "__main__":
    test_categorize_data()
