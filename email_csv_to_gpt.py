import pandas as pd
import openai
import csv

# This code was part of my email_to_csv.py originally but only because I didn't have access to the emails in Outlook and was using a csv printout of them


extracted_text_dict = {}

# Path to your CSV file
emails_csv = r"C:\Users\briel\Downloads\hfreturnsemailoutput.csv"

# Path to your CSV file
output_file = r"C:\Users\briel\Downloads\processed_emails.csv"

# Read the CSV into a DataFrame
df = pd.read_csv(emails_csv)

# Check if the "Body" column exists
if "Body" in df.columns:
    # Extract text after "FW: " in the "Body" column
    df["Processed_Body"] = df["Body"].apply(
        lambda x: (
            x.split("<http://www.baystreetadvisorsllc.com/>", 1)[1].split("The information contained in this email", 1)[0].strip()
            if "<http://www.baystreetadvisorsllc.com/>" in x and 
            (x.rfind("<http://www.baystreetadvisorsllc.com/>") > x.rfind("Subject: "))
            else x.split("Subject: ", 1)[1].split("The information contained in this email", 1)[0].strip()
            if "Subject: " in x else None
        )
    )

    # Drop rows where "Processed_Body" is NaN (if needed)
    df = df.dropna(subset=["Processed_Body"])
    
    # Save the processed DataFrame to a new CSV file
    df.to_csv(output_file, index=False)
    print(f"Processed data saved to: {output_file}")
else:
    print("The column 'Body' does not exist in the CSV.")

with open(output_file, mode='r', encoding='utf-8', errors='ignore') as infile:

    reader = csv.reader(infile)
    next(reader)
    for row in reader:
        # Clean each field in the row
        cleaned_row = [''.join(char for char in field if char.isprintable()) if isinstance(field, str) else field for field in row]
        unstructured_text = cleaned_row[5]

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
        extracted_text = response.choices[0].message.content.strip()
        print(extracted_text)
        for line in extracted_text.strip().split('\n'):
            if ':' in line:  # Ensure the line has a key-value structure
                key, value = line.split(':', 1)
                extracted_text_dict[key.strip()] = value.strip()
            # Convert the dictionary to a DataFrame
        extracted_text_df = pd.DataFrame([extracted_text_dict])
        combined_hf_df = pd.concat([combined_hf_df, extracted_text_df], ignore_index=True)