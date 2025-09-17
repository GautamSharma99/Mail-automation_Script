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
CSV_URL = ''

# Gmail setup
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
EMAIL_ADDRESS = 'ambitiontwenty25@gmail.com'
EMAIL_PASSWORD = 'iumd gwxp vpzz tcfp'

# Email content
EMAIL_SUBJECT = "Congratulations on Your Hackathon Registration!🎉"
EMAIL_BODY_HTML = """
<html>
  <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
    <p>Dear Participant,</p>

    <p>We’re excited to have you join us for the upcoming 
    <b>amBITion v2 Hackathon!</b></p>

    <p>Please note that the event dates have been rescheduled to 
    <b>4th &amp; 5th October, 2025</b> due to participants’ request. 
    Mark your calendars and get ready for two days of 
    <b>innovation, creativity, and collaboration</b>.</p>

    <h3 style="color:#2E86C1;">Do’s ✅</h3>
    <ul>
      <li><b>Bring</b> your laptops, chargers, extension chords, and any other required equipment.</li>
      <li><b>Carry</b> a valid ID card for registration.</li>
      <li><b>Work collaboratively</b> and respect your teammates’ ideas.</li>
      <li><b>Be open</b> to mentorship sessions and ask questions.</li>
      <li><b>Take breaks</b>, stay hydrated, and manage your time well.</li>
      <li><b>Follow</b> the submission guidelines and deadlines provided during the event.</li>
    </ul>

    <h3 style="color:#C0392B;">Don’ts ❌</h3>
    <ul>
      <li><b>Do not</b> indulge in plagiarism or copy existing solutions.</li>
      <li><b>Do not</b> engage in any form of misconduct, disrespect, or disruptive behavior.</li>
      <li><b>Avoid</b> hardcoding or using pre-built projects as submissions.</li>
      <li><b>Do not forget</b> to follow the code of conduct and respect the venue rules.</li>
    </ul>

    <p>We’ll be sharing the <b>detailed schedule, venue, and other logistics</b> soon. 
    Meanwhile, keep brainstorming your ideas and get ready to build something impactful!</p>

    <p>If you have any questions, feel free to reach out to us at:<br>
    📧 <a href="mailto:iamnayanchawhan@gmail.com">iamnayanchawhan@gmail.com</a><br>
    📞 <b>Nayan G Chawhan</b> - 9620247096</p>

    <p><b>IMPORTANT - JOIN THE WHATSAPP GROUP FOR IMPORTANT UPDATES:</b><br>
    👉 <a href="https://chat.whatsapp.com/DbsPstsG42PDktulphHBzh?mode=ems_copy_c" 
    style="color:#1E8449; font-weight:bold;">Click here to join WhatsApp Group</a></p>

    <p>Looking forward to seeing your <b>energy and creativity</b> on 
    <b>4th &amp; 5th October!</b></p>

    <p>Best Regards,<br>
    <b>Team amBITion v2</b></p>
  </body>
</html>
"""


EMAIL_BODY_TEMPLATE = """
Dear {name},

We’re excited to have you join us for the upcoming amBITion v2 Hackathon!

Please note that the event dates have been rescheduled to 4th & 5th October, 2025 due to participants' request. Mark your calendars and get ready for two days of innovation, creativity, and collaboration.

To ensure a smooth and enjoyable experience, here are some Do’s and Don’ts for the hackathon:

Do’s:
- Bring your laptops, chargers, extension chords and any other required equipment.
- Carry a valid ID card for registration.
- Work collaboratively and respect your teammates’ ideas.
- Be open to mentorship sessions and ask questions.
- Take breaks, stay hydrated, and manage your time well.
- Follow the submission guidelines and deadlines provided during the event.

Don’ts:
- Do not indulge in plagiarism or copy existing solutions.
- Do not engage in any form of misconduct, disrespect, or disruptive behavior.
- Avoid hardcoding or using pre-built projects as submissions.
- Do not forget to follow the code of conduct and respect the venue rules.

We’ll be sharing the detailed schedule, venue, and other logistics soon. Meanwhile, keep brainstorming your ideas and get ready to build something impactful!

If you have any questions, feel free to reach out to us at:
Email: iamnayanchawhan@gmail.com
Nayan G Chawhan - 9620247096

Looking forward to seeing your energy and creativity on 4th & 5th October!

IMPORTANT - JOIN THE WHATSAPP GROUP FOR IMPORTANT UPDATES:
https://chat.whatsapp.com/DbsPstsG42PDktulphHBzh?mode=ems_copy_c

Best Regards,
Team amBITion v2
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


def send_email(recipient, subject, body, html_body):
    msg = EmailMessage()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = recipient
    msg['Subject'] = subject
    msg.set_content(body)  # plain-text fallback
    msg.add_alternative(html_body, subtype='html')  # HTML version

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
            send_email(email, EMAIL_SUBJECT, EMAIL_BODY_TEMPLATE, EMAIL_BODY_HTML)
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
