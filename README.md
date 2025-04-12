# Send Mailer

## Description

This is a simple mailer that reads data from Google Sheets and can send emails using Python.

## Prerequisites
Python 3.9+ is required to run this project. You can download it from:
- [Python](https://www.python.org/downloads/)

## Installation

Clone the repository and navigate to the project directory in the terminal.

```bash
git clone https://github.com/ConducereIT/send-mailer/
cd send-mailer
```

Create a virtual environment and install the dependencies:

```bash
python -m venv venv
source venv/bin/activate  # On Windows use: venv\Scripts\activate
pip install -r requirements.txt
```

## Configuration

To use the mailer, you need to configure the following environment variables in a `.env` file:

```env
GOOGLE_SHEET_ID=your_google_sheet_id # must have a column named "email"
SEND_MAIL_HOST=smtp.gmail.com
SEND_MAIL_USER=your_email@gmail.com
SEND_MAIL_PASS=your_app_password
SEND_MAIL_SERVICE=Gmail
SEND_MAIL_SUBJECT=your_email_subject # the subject of the email
TEMPLATE_PATH=path/to/your/index.html # the path to the template html file
REPLACE_ARRAY_NAMES=name,phone,etc # the name of the columns in the sheet that will be replaced in the template
```

## Usage

To run the application:

```bash
python app.py
```

The application will:
1. Read data from the specified Google Sheet
2. Check for duplicate email addresses
3. Display all found email addresses and names
4. Log all operations to `app.log`
5. Send emails to the found email addresses with the specified subject and template

## Features

- Reads data from Google Sheets
- Validates email addresses
- Checks for duplicate emails
- Detailed logging
- Configurable through environment variables
- Sends emails to the found email addresses with the specified subject and template

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details
