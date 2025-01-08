import os
import win32com.client

def save_emails_from_outlook(subfolder_name, save_directory):
    """Access emails from a subfolder in Outlook and save contents to a specified directory."""
    try:
        # Initialize Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        # Access Inbox folder
        inbox = outlook.GetDefaultFolder(6)  # 6 refers to the Inbox folder
        
        # Find the target subfolder
        subfolder = None
        for folder in inbox.Folders:
            if folder.Name == subfolder_name:
                subfolder = folder
        
        if subfolder is None:
            print(f"Subfolder '{subfolder_name}' not found.")
            return

        # Create save directory if it doesn't exist
        if not os.path.exists(save_directory):
            os.makedirs(save_directory)
        
        # Loop through emails in the subfolder
        for idx, email in enumerate(subfolder.Items, start=1):
            subject = email.Subject
            received_time = email.ReceivedTime.strftime("%Y-%m-%d_%H-%M-%S")
            file_name = f"email_{idx}_{received_time}.txt"
            file_path = os.path.join(save_directory, file_name)
            
            # Save email subject and body to a text file
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(f"Subject: {subject}\n")
                f.write(f"Received: {email.ReceivedTime}\n\n")
                f.write(email.Body)
            
    except Exception as e:
        print(f"An error occurred: {e}")

# Example usage
subfolder_name = "TargetSubfolder"  # Replace with the name of your Outlook subfolder
save_directory = r"C:\Users\YourUsername\Documents\OutlookEmails"  # Replace with your desired save path
save_emails_from_outlook(subfolder_name, save_directory)