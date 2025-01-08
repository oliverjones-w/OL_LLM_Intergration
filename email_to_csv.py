import os
import csv
from datetime import datetime
import win32com.client

def update_emails_to_csv(subfolder_name, csv_path):
    """
    Collect new emails from an Outlook subfolder and update CSV file
    """
    try:
        # Initialize Outlook application
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        """
        # Access the Inbox folder
        inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox folder
        """
        # Locate the target subfolder
        subfolder = None
        for folder in outlook.Folders:
            if folder.Name == subfolder_name:
                subfolder = folder

        if subfolder is None:
            print(f"Subfolder '{subfolder_name}' not found.")
            return

        with open(csv_path, mode='a', newline='', encoding='utf-8') as csv_file:
            csv_writer = csv.writer(csv_file)

            # Get today's date for filtering
            today = datetime.now().date()

            # Loop through emails in the subfolder
            for email in subfolder.Items:
                received_date = email.ReceivedTime.date()
                if received_date == today:
                    subject = email.Subject
                    body = email.Body.strip().replace('\n', ' ').replace('\r', '')

                    # Write email details to the CSV
                    csv_writer.writerow([subject, body])

        print(f"Emails successfully saved to '{csv_path}'.")
    
    except Exception as e:
        print(f"An error occurred: {e}")

subfolder_name = "Test"   # Replace
csv_path = r"C:\Users\briel\Downloads\baystreet\folder\contacts.csv"  # Replace

# Fetch emails and update CSV
update_emails_to_csv(subfolder_name, csv_path)