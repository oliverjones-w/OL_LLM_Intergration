import os
import csv
import win32com.client
from dotenv import load_dotenv
from langchain_openai.llms import OpenAI
import openai
import pandas as pd

# Initialize an empty list to store the processed values
processed_list = []
extracted_text_dict = {}

combined_pm_df = pd.DataFrame()
combined_hf_df = pd.DataFrame()


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





def update_emails_to_csv(account_name, subfolder_name, csv_path):
    """
    Automate the retrieval of emails from Outlook, saving them to a CSV for further processing.
    """
    global extracted_text_dict
    global combined_pm_df
    global combined_hf_df
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
                    # Process the body to keep only text before "The information contained in this email"
                    if "The information contained in this email" in email.Body:
                        body = email.Body.split("The information contained in this email", 1)[0].strip()
                    else:
                        body = email.Body.strip()
                    # Replace newline and carriage return characters for consistency
                    body = body.replace('\n', ' ').replace('\r', '')

                    # Write email data to the CSV
                    csv_writer.writerow([email_id, sent_time, subject, sender, body])
                    extracted_text, extracted_folder = categorize_data(body, subfolder_name)
                    # Parse the string into a dictionary
                    for line in extracted_text.strip().split('\n'):
                        if ':' in line:  # Ensure the line has a key-value structure
                            key, value = line.split(':', 1)
                            extracted_text_dict[key.strip()] = value.strip()
                    # Convert the dictionary to a DataFrame
                    extracted_text_df = pd.DataFrame([extracted_text_dict])
                    if extracted_folder == "HFReturns":
                        combined_hf_df = pd.concat([combined_hf_df, extracted_text_df], ignore_index=True)
                    elif extracted_folder == "People Moves":
                        combined_pm_df = pd.concat([combined_pm_df, extracted_text_df], ignore_index=True)                   
                except Exception as email_error:
                    print(f"Error processing email: {email_error}")
        
    except Exception as e:
        print(f"An error occurred: {e}")

# Function to interact with OpenAI API
def categorize_data(unstructured_text, subfolder_name):
    if subfolder_name == "HFReturns": # edit according to exact folder name
        folder = "HFReturns"
        response = openai.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {
                    "role": "user",
                    "content": f"Extract the following information from the text and assign to these categories: "
                    f"Fund Name, Date, Monthly Return, YTD, Annualized Return, Strategy. "
                    f"For Monthly Return, YTD, and Annualized Return, please provide only one percantage. "
                    f"Strategy should be one sentence maximum. "
                    f"Please clean up the data as well so there are no stray characters and no dashes: \"{unstructured_text}\"", # edit columns according to preference
                }
            ]
        )
    elif subfolder_name == "People Moves": # edit according to exact folder name
        folder = "People Moves"
        response = openai.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {
                    "role": "user",
                    "content": f"Extract the following information from the text and assign to these categories: "
                    f"Name, Function, Current Firm, Current Title, Date Joined, Current Location, Former Firm, Former Title, Date Left, Former Location, Notes. "
                    f"Please clean up the data as well so there are no stray characters and no dashes: \"{unstructured_text}\"", # edit columns according to preference
                }
            ]
    )

    # Extract the message content
    extracted_text = response.choices[0].message.content.strip()
    return extracted_text, folder

# Configuration
account_name = "intern2@baystreetadvisorsllc.com"  # Replace with your account name
csv_path = r"C:\Users\briel\Downloads\baystreet\folder\contacts.csv"
parsed_hf_table_csv = r"C:\Users\briel\Downloads\baystreet\folder\parsed_hf_table.csv"
parsed_pm_table_csv = r"C:\Users\briel\Downloads\baystreet\folder\parsed_pm_table.csv"

# Fetch emails and update CSV
subfolder_name = "People Moves" # Replace with your subfolder name
update_emails_to_csv(account_name, subfolder_name, csv_path)
subfolder_name = "HFReturns" # Replace with your subfolder name
update_emails_to_csv(account_name, subfolder_name, csv_path)
combined_hf_df.to_csv(parsed_hf_table_csv, mode='w', index=False, header=True)
combined_pm_df.to_csv(parsed_pm_table_csv, mode='w', index=False, header=True)