# PyEmail

## Overview

This script automates the process of sending emails to entities listed in a Google Spreadsheet. It retrieves data from the spreadsheet, generates personalized emails, and sends them with attached PDFs and images.

## Features

- **Email Customization:** Utilizes an HTML template to customize the email body.
- **Attachment Support:** Attaches PDF files and images to the emails.
- **Logging:** Logs sent emails to CSV and text files for reference.

## Prerequisites

Before running the script, ensure you have the following:

- Google Spreadsheet with relevant data.
- `.env` file with necessary environment variables.
- `credentials.json` for Google Sheets API access.
- Properly formatted HTML email body in a file named `body.html`.
- PDF files and image files to be attached listed in the script.

## Setup

1. Clone the repository:

   ```bash
   git clone https://github.com/sarthakghere/PyEmail.git
   ```

2. Populate the `.env` file with your credentials:

   ```
   spreadsheet_id=YOUR_SPREADSHEET_ID
   spreadsheet_name=YOUR_SPREADSHEET_NAME
   path=YOUR_CREDENTIALS_JSON_PATH
   email=YOUR_EMAIL
   pass=YOUR_EMAIL_APP_PASSWORD
   ```

3. Customize the email subject and column names in the script:

   ```python
   SUBJECT = "YOUR SUBJECT HERE"
   NAME_COLUMN = "SCHOOLS"
   EMAIL_COLUMN = "Email-id"
   ```

4. Add PDF and image paths to the respective lists in the script.

## Usage

Run the script:

```bash
python main.py
```

The script will fetch emails from the Google Spreadsheet, send personalized emails, and log the sent emails.

## Author

- **Sarthak Gupta**
  - GitHub: [sarthakghere](https://github.com/sarthakghere)

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

--- 
