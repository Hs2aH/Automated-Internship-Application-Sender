# Automated Internship Application Sender
This Python script automates the process of sending out internship applications to multiple companies. It reads company and contact information from an Excel file, composes personalized emails, and attaches your CV and a reference letter before sending them out.

### 🚀 Features
+ Bulk Emailing: Reads a list of companies and their contact emails from a .xlsx file.

+ Personalization: Automatically generates a personalized greeting for each email, using the contact person's name if available.

+ Attachment Handling: Attaches your CV and a reference letter to each email.

+ Safety First: Includes a safety prompt before sending to prevent accidental mass emails.

+ Dry Run & Test Modes: Allows you to test the script without sending real emails or to send to a test email address.

+ Activity Logging: Creates a log file (.csv) to track the status of each email sent.

### ⚙️ Requirements
    Python 3.x

The following Python libraries: pandas, smtplib, email

1. A .xlsx file with your company list.

2. Your CV and reference letter in PDF format.

3. A Gmail account with a 16-character App Password generated specifically for this script.

### 🛠️ Setup and Configuration
1. Generate a Gmail App Password:

2. Go to your Google Account settings > Security.

3. Find "2-Step Verification" and ensure it's on.

4. Under "2-Step Verification," find "App passwords."

5. Create a new app password and copy the 16-character code. This is what you'll use for YOUR_PASSWORD in the script. Do not use your main Gmail password.

6. Clone the Repository:

       git clone https://github.com/your-username/your-repo-name.git
7. cd your-repo-name

8. Install Dependencies:

        pip install pandas openpyxl

9. Place Your Files:

10. Place your Excel file (e.g., Industrial Training Tri 3 2023 Company List.xlsx) in the project directory.

11. Place your CV and reference letter PDFs in the same directory.

12. Edit the Configuration:

13. Open internship_sender.py in a text editor.

14. Update the CONFIGURATION section with your specific file names and email details:

        EXCEL_FILE = "Industrial Training Tri 3 2023 Company List.xlsx"
        CV_FILE = "CV MMU student (ALHAWBANI HUSAM).pdf"
        REFERENCE_FILE = "MMU Student REFERENCE LETTER.pdf"
        YOUR_EMAIL = "your.email@gmail.com"
        YOUR_PASSWORD = "your-16-char-app-password"
        LOG_FILE = "sent_log.csv"

Review the Email Body:

17. Customize the body variable in the script to reflect your personal message, skills, and contact information.

### 🚀 Usage
+ To run the script, open your terminal or command prompt and execute:

      python internship_sender.py

+ A safety prompt will appear asking for confirmation before sending. Type y and press enter to proceed.

+ The script will print real-time status updates to the console for each email sent or skipped.

+ Once completed, a log file named sent_log.csv will be created in the directory, which you can open with Excel or a similar program to review the results.

### ⚠️ Important Notes
+ Be mindful of email sending limits. Sending too many emails in a short period might flag your account as spam.

+ Customize the subject and body to fit your needs. The provided text is a starting point.

+ The script assumes your Excel file has columns named exactly as specified in the code. Adjust the column names in the script if yours are different.
