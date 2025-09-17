import csv
import re
import time
import smtplib
from email.message import EmailMessage
import requests

# -------------------------
# CONFIGURATION
# -------------------------

# Google Sheet CSV link
CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRSKtp9d8U5F-5V4VDi78oCrbiA8FoJSAEdDWXCHGN-2seP98RBSLORLMRDpL1X8KqsWtawMgnbwhBw/pub?gid=0&single=true&output=csv'

# Gmail setup
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
EMAIL_ADDRESS = 'gautamsharma99067@gmail.com'
EMAIL_PASSWORD = 'pzqn ttku fdbl eyzn'

# Email content
EMAIL_SUBJECT = "Hackathon Registration Confirmation"
EMAIL_BODY_TEMPLATE = """
Hi {name},

You have been successfully registered for the hackathon. We look forward to your participation!

Best regards,
Hackathon Team
"""

# Delay between emails
EMAIL_DELAY = 2

# Path to local file to store updated CSV
LOCAL_CSV_FILE = 'participants_updated.csv'


# -------------------------
# FUNCTIONS
# -------------------------

def is_valid_email(email):
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return re.match(pattern, email) is not None


def send_email(recipient, subject, body):
    msg = EmailMessage()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = recipient
    msg['Subject'] = subject
    msg.set_content(body)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.send_message(msg)


def download_csv(url):
    response = requests.get(url)
    response.raise_for_status()
    decoded_content = response.content.decode('utf-8')
    return list(csv.reader(decoded_content.splitlines()))


def save_csv(file_path, rows):
    with open(file_path, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(rows)


def main():
    print("Downloading CSV from Google Sheets...")
    rows = download_csv(CSV_URL)
    print(f"Loaded {len(rows) - 1} participants.")

    updated = False
    for index, row in enumerate(rows[1:], start=1):  # Skip header
        name = row[0].strip()
        email = row[1].strip()
        status = row[2].strip() if len(row) > 2 else ""

        if not email:
            print(f"Row {index}: No email provided, skipping.")
            continue

        if not is_valid_email(email):
            print(f"Row {index}: Invalid email '{email}', skipping.")
            continue

        if status.lower() == "sent":
            print(f"Row {index}: Already sent, skipping.")
            continue

        try:
            body = EMAIL_BODY_TEMPLATE.format(name=name)
            send_email(email, EMAIL_SUBJECT, body)
            print(f"Row {index}: Email sent to {email}")

            # Update status
            if len(row) < 3:
                row.append("Sent")
            else:
                row[2] = "Sent"

            updated = True
            time.sleep(EMAIL_DELAY)

        except Exception as e:
            print(f"Row {index}: Failed to send email to {email}. Error: {e}")

    if updated:
        print(f"Saving updated CSV to {LOCAL_CSV_FILE}")
        save_csv(LOCAL_CSV_FILE, rows)
        print("Done! Please upload this file manually to Google Sheets if needed.")
    else:
        print("No updates made. All emails already sent or skipped.")


# -------------------------
# RUN
# -------------------------

if __name__ == "__main__":
    main()
