from pathlib import Path

import pandas as pd  # pip install pandas openpyxl
import win32com.client as win32  # pip install pywin32

# Locate examples files & create output directory
EXCEL_FILE_PATH = Path.cwd() / "Financial_Data.xlsx"
ATTACHMENT_DIR = Path.cwd() / "Attachments"

ATTACHMENT_DIR.mkdir(exist_ok=True)

# Load financial data into dataframe
data = pd.read_excel(EXCEL_FILE_PATH, sheet_name="Data")

# Get unique values from any particular column
column_name = "Country"
unique_values = data[column_name].unique()

# Query/Filter the dataframe and export the filtered dataframe as an Excel file
for unique_value in unique_values:
    data_output = data.query(f"{column_name} == @unique_value & Year==2021")
    output_path = ATTACHMENT_DIR / f"{unique_value}_2021.xlsx"
    data_output.to_excel(output_path, sheet_name=unique_value, index=False)

# Load email distribution list into dataframe
email_list = pd.read_excel(EXCEL_FILE_PATH, sheet_name="Email_List")

# Iterate over email distribution list & send emails via Outlook App
outlook = win32.Dispatch("outlook.application")
for index, row in email_list.iterrows():
    mail = outlook.CreateItem(0)
    mail.To = row["Email"]
    mail.CC = row["CC"]
    mail.Subject = f"Financial Report for: {row['Country']}"
    # mail.Body = "Message body"
    mail.HTMLBody = f"""
                    <b>Hi {row['Name']}</b>,<br><br>
                    Please find attached the report for {row['Country']}.<br><br>
                    Best Regards,<br>
                    Sven
                    """
    attachment_path = str(ATTACHMENT_DIR / f"{row['Country']}_2021.xlsx")
    mail.Attachments.Add(Source=attachment_path)

    mail.Display()

    # Uncomment to send email
    # mail.Send()
