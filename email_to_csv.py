import os
import csv
from datetime import datetime
import win32com.client
from dotenv import load_dotenv
from langchain_openai.llms import OpenAI
import openai
import pandas as pd

def update_emails_to_csv(account_name, subfolder_name, csv_path):
    """
    Automate the retrieval of emails from Outlook, saving them to a CSV for further processing.
    """
    try:
        # Initialize Outlook application
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Locate the root folder (your email account)
        root_folder = None
        for folder in outlook.Folders:
            if folder.Name == account_name:
                root_folder = folder
                break


        if root_folder is None:
            print(f"Account folder '{account_name}' not found.")
            return
        
        # Locate the target subfolder within the Inbox
        subfolder = None
        for folder in root_folder.Folders:
            if folder.Name == subfolder_name:
                subfolder = folder
                break

        if subfolder is None:
            print(f"Subfolder '{subfolder_name}' not found under 'Inbox'.")
            return

        # Write headers and initialize the CSV (overwrite mode)
        with open(csv_path, mode='w', newline='', encoding='utf-8') as csv_file:
            csv_writer = csv.writer(csv_file)
            # Write column headers
            csv_writer.writerow(["EmailID", "DateTime Sent", "Subject", "Sender", "Body"])

        # Read existing email IDs to avoid duplicates
        processed_ids = set()
        if os.path.isfile(csv_path):
            with open(csv_path, mode='r', encoding='utf-8') as csv_read_file:
                csv_reader = csv.reader(csv_read_file)
                next(csv_reader, None)  # Skip the header row
                for row in csv_reader:
                    if row:  # Ensure row is not empty
                        processed_ids.add(row[0])  # Assuming the first column is the email ID

        # Append new email data to the CSV
        with open(csv_path, mode='a', newline='', encoding='utf-8') as csv_file:
            csv_writer = csv.writer(csv_file)

            # Loop through emails in the subfolder
            for email in subfolder.Items:
                try:
                    email_id = email.EntryID
                    if email_id in processed_ids:
                        continue  # Skip already processed emails

                    sent_time = email.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                    subject = email.Subject or "(No Subject)"
                    sender = email.SenderName or "(No Sender)"
                    body = email.Body.strip().replace('\n', ' ').replace('\r', '')

                    # Write email data to the CSV
                    csv_writer.writerow([email_id, sent_time, subject, sender, body])
                    extracted_text = categorize_data(body, subfolder_name)
                    print(extracted_text)
                    # Parse the string into a dictionary
                    extracted_text_dict = {}
                    for line in extracted_text.strip().split('\n'):
                        if ':' in line:  # Ensure the line has a key-value structure
                            key, value = line.split(':', 1)
                            extracted_text_dict[key.strip()] = value.strip()
                    # Convert the dictionary to a DataFrame
                    extracted_text_df = pd.DataFrame([extracted_text_dict])
                    
                    # extracted_text_df = pd.DataFrame(extracted_text)
                    print(extracted_text_df)
                    # export_to_csv(extracted_text, parsed_table_csv)
                except Exception as email_error:
                    print(f"Error processing email: {email_error}")

        print(f"Emails successfully updated in '{csv_path}'.")
    
    except Exception as e:
        print(f"An error occurred: {e}")



# Load environment variables from .env file
load_dotenv()

my_api_key = os.getenv("OPENAI_API_KEY")
openai.api_key = my_api_key

if not my_api_key:
    raise ValueError("OPENAI_API_KEY environment variable is not set.")
else:
    print(f"API Key is set")


# Initialize the OpenAI client
client = OpenAI(api_key=my_api_key)

# Function to interact with OpenAI API
def categorize_data(unstructured_text, subfolder_name):
    if subfolder_name == "HFReturns": # edit according to exact folder name
        response = openai.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {
                    "role": "user",
                    "content": f"Extract the following information from the text and assign to these categories: Fund Name, Date, Monthly Return, YTD, Annualized Return, Strategy: \"{unstructured_text}\"", # edit columns according to preference
                }
            ]
        )
    elif subfolder_name == "People Moves": # edit according to exact folder name
        response = openai.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {
                    "role": "user",
                    "content": f"Extract the following information from the text and assign to these categories: Name, Function, Current Firm, Current Title, Date Joined, Current Location, Former Firm, Former Title, Date Left, Former Location, Note: \"{unstructured_text}\"",
                }
            ]
    )

    # Extract the message content
    extracted_text = response.choices[0].message.content.strip()
    return extracted_text

# Configuration
account_name = "intern2@baystreetadvisorsllc.com"  # Replace with your account name
csv_path = r"C:\Users\briel\Downloads\baystreet\folder\contacts.csv"
parsed_table_csv = r"C:\Users\briel\Downloads\baystreet\folder\parsed_table.csv"

# Fetch emails and update CSV
subfolder_name = "People Moves" # Replace with your subfolder name
update_emails_to_csv(account_name, subfolder_name, csv_path)
subfolder_name = "HFReturns" # Replace with your subfolder name
update_emails_to_csv(account_name, subfolder_name, csv_path)