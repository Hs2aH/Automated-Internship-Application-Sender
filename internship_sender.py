import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import time
from datetime import datetime

# === CONFIGURATION ===
EXCEL_FILE = "Industrial Training Tri 3 2023 Company List.xlsx"
CV_FILE = "CV MMU student (ALHAWBANI HUSAM).pdf"                  # your CV filename
REFERENCE_FILE = "MMU Student REFERENCE LETTER.pdf"   # your reference letter filename
YOUR_EMAIL = "Husam.howbani.2002@gmail.com"
YOUR_PASSWORD = "pwsuofarpjrsmfnk"       # 16-char Gmail App Password (no spaces)
LOG_FILE = "REAL sent_log.csv"

# === MODES ===
DRY_RUN = False       # False = real send
TEST_MODE = False     # False = real HR emails

# === SAFETY PROMPT ===
confirm = input("⚠️ WARNING: This will send emails to ALL companies in the Excel file. Continue? (y/n): ")
if confirm.lower() != "y":
    print("❌ Operation cancelled. No emails sent.")
    exit()

# Load Excel data
df = pd.read_excel(EXCEL_FILE, sheet_name=0)

# Setup SMTP
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(YOUR_EMAIL, YOUR_PASSWORD)

# List to store log entries
log_entries = []

# Count total emails to send
total = len(df)
sent_count = 0

for index, row in df.iterrows():
    company = row["Company Name"]
    hr_name = row["Company Contact Person( Intern recruitment, person in charge of the internship....)"]
    email = row["Company Contact Email"]

    if pd.isna(email):
        log_entries.append([company, hr_name, None, "Skipped - No email", datetime.now()])
        print(f"⚠️ Skipped {company} - no email found")
        continue  # skip if no email

    sent_count += 1
    print(f"\n📨 Sending email {sent_count} of {total} → {company}")

    # Greeting logic
    if pd.notna(hr_name):
        greeting = f"Hello and Dear {hr_name},"
    else:
        greeting = f"Hello and Dear {company} HR Department,"

    # Email subject & body
    subject = f"Internship Request - {company}"
    body = f"""### Internship Request

{greeting}

My name is Husam, and I'm a 23-year-old student about to graduate from MMU with a Bachelor's degree in Computer Science, specializing in Cyber Security and Networks.

I am confident in my skills and fully prepared to excel in technical interviews. Unlike many students who submit their CVs and wait, I actively pursue opportunities because I do not have the luxury of failure.

Please take a look at my attached resume and reference letter. If my profile resonates with your needs, I am available for an interview at your convenience.

Best regards,  
Husam Al-howbani
"""

    try:
        # Create email
        msg = MIMEMultipart()
        msg['From'] = YOUR_EMAIL
        msg['To'] = email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Attach files (CV + Reference Letter)
        for file in [CV_FILE, REFERENCE_FILE]:
            if os.path.exists(file):
                with open(file, "rb") as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                filename = os.path.basename(file)
                part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
                msg.attach(part)
                print(f"   📎 Attached: {filename}")
            else:
                print(f"   ⚠️ File not found: {file}")

        # Send email
        server.send_message(msg)
        print(f"✅ Email {sent_count}/{total} sent to {email} (for company {company})")

        # Add success entry
        log_entries.append([company, hr_name, email, "Sent", datetime.now()])

        # Pause before next email
        time.sleep(5)

    except Exception as e:
        print(f"❌ Failed to send to {email} ({company}): {e}")
        log_entries.append([company, hr_name, email, f"Failed - {e}", datetime.now()])

# Quit SMTP
server.quit()

# Save log to CSV
log_df = pd.DataFrame(log_entries, columns=["Company", "HR Name", "Email", "Status", "Timestamp"])
log_df.to_csv(LOG_FILE, index=False)

print(f"\n📌 Log saved to {LOG_FILE}")

