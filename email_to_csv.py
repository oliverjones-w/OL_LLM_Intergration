import os
import csv
from datetime import datetime
import win32com.client

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

        # Locate the Inbox folder
        inbox_folder = None
        for folder in root_folder.Folders:
            if folder.Name == "Inbox":
                inbox_folder = folder
                break

        if inbox_folder is None:
            print(f"Inbox folder not found under account '{account_name}'.")
            return

        # Locate the target subfolder within the Inbox
        subfolder = None
        for folder in inbox_folder.Folders:
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
                except Exception as email_error:
                    print(f"Error processing email: {email_error}")

        print(f"Emails successfully updated in '{csv_path}'.")
    
    except Exception as e:
        print(f"An error occurred: {e}")


# Configuration
account_name = "ojones@baystreetadvisorsllc.com"  # Replace with your account name
subfolder_name = "HFReturns"                     # Replace with your subfolder name
csv_path = r"C:\Users\BSA-OliverJ'22\OneDrive\Desktop\OneDrive\Programming\Projects\OL_LLM_Integration\OL_LLM_Intergration\HFReturnsEmailOutput.csv"

# Fetch emails and update CSV
update_emails_to_csv(account_name, subfolder_name, csv_path)
