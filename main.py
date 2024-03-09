# Script developed by Sarthak Gupta
# GitHub: sarthakghere
# Date: 9 Mar, 2024

import gspread
from google.oauth2.service_account import Credentials
import os
import dotenv
import smtplib
import csv
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# Load environment variables from .env file
dotenv.load_dotenv()

# Get required information from environment variables
SPREADSHEET_ID = os.getenv("spreadsheet_id")
SPREADHSEET_NAME = os.getenv("spreadsheet_name")
CREDENTIALS_JSON = os.getenv("path")

# Lists to store paths and names of images and PDFs
IMAGES_PATH = []
PDF_PATH = []

# Email subject
SUBJECT = "YOUR SUBJECT HERE"

# Name of the column containing entity names in the spreadsheet
NAME_COLUMN = "SCHOOLS"

# Name of the column containing recipient emails in the spreadsheet
EMAIL_COLUMN = "Email-id"

# Function to send an email
def send_mail(email, message, name):
    # Get email credentials from environment variables
    my_email = os.getenv("email")
    password = os.getenv("pass")

    # Connect to SMTP server and send email
    connection = smtplib.SMTP("smtp.gmail.com")
    connection.starttls()
    connection.login(user=my_email, password=password)
    connection.sendmail(from_addr=my_email, to_addrs=email, msg=message)
    connection.close()

    # Log sent email information to CSV and text files
    row = [name, email]
    with open("emails_sent.csv", "a", newline='') as file:
        writer = csv.writer(file)
        writer.writerow(row)
    with open("emails_sent.txt", "a") as file:
        file.write("Email sent to: " + email + "\n")
    print("Email sent to: " + email)

# Function to retrieve data from the Google Spreadsheet
def get_data_from_spreadsheet():
    creds = Credentials.from_service_account_file(CREDENTIALS_JSON, scopes=['https://www.googleapis.com/auth/spreadsheets'])
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SPREADHSEET_NAME)
    return sheet

# Function to generate the email message
def generate_message(row_data):
    message = MIMEMultipart()
    message['From'] = os.getenv("email")
    message['To'] = row_data[EMAIL_COLUMN]
    message['Subject'] = SUBJECT

    # Read email body from an HTML file
    with open("./body.html") as file:
        content = file.readlines()
        body = ""
        for item in content:
            body += item

    # Attach email body as HTML
    message.attach(MIMEText(body, 'html'))

    # Attach PDF files
    for pdf_path in PDF_PATH:
        with open(pdf_path, 'rb') as attachment:
            pdf_part = MIMEApplication(attachment.read(), Name=os.path.basename(pdf_path))

        pdf_part['Content-Disposition'] = f'attachment; filename={os.path.basename(pdf_path)}'
        message.attach(pdf_part)

    # Attach image files
    for img_path in IMAGES_PATH:
        with open(img_path, 'rb') as attachment:
            image_part = MIMEImage(attachment.read(), name=os.path.basename(img_path))

        message.attach(image_part)

    return message.as_string()

# Function to get the index of the last processed row from a file
def get_last_processed_row():
    try:
        with open("id.txt", "r") as file:
            return int(file.read())
    except FileNotFoundError:
        return 0

# Function to update the last processed row index in a file
def update_last_processed_row(row_index):
    with open("id.txt", "w") as file:
        file.write(str(row_index))

# Main function
def main():
    # Get the index of the last processed row
    last_processed_row = get_last_processed_row()
    print("Fetching emails...")

    # Connect to the Google Spreadsheet and fetch all records
    spreadsheet = get_data_from_spreadsheet()
    spreadsheet_data = spreadsheet.get_all_records()
    row_count = len(spreadsheet_data)

    # Process each row if there are new rows to process
    if row_count > last_processed_row:
        for i in range(last_processed_row, row_count):
            print(f"Row:\n{spreadsheet_data[i]}")
            row_data = spreadsheet_data[i]
            print(f"Sending email to {row_data[NAME_COLUMN]}")
            
            # Generate email message and send email
            message = generate_message(row_data=row_data)
            send_mail(row_data[EMAIL_COLUMN], message=message, name=row_data[NAME_COLUMN])
            
            # Update the last processed row index
            last_processed_row = i + 1
            update_last_processed_row(last_processed_row)

# Run the main function if the script is executed
if __name__ == "__main__":
    main()
