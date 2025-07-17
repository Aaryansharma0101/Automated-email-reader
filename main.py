from email_reader import connect_to_mailbox, fetch_emails
from utils import save_emails_to_excel
import os
import time

# 🔐 Get email and password from user
EMAIL = input("Enter your email address: ").strip()
PASSWORD = input("Enter your app password: ").strip()

# 🔢 How many emails to scan
try:
    LIMIT = int(input("How many latest emails to scan (e.g., 30): ").strip())
except ValueError:
    LIMIT = 30  # default fallback

# 📨 Subject filter (can be comma-separated)
subject_input = input("Enter subject keywords (comma separated): ").strip()
SUBJECT_FILTER = [s.strip() for s in subject_input.split(",") if s.strip()]

ATTACHMENT_FOLDER = "attachments"

print("Connecting to mailbox...")
mailbox = connect_to_mailbox(EMAIL, PASSWORD)

print("Fetching emails...")
emails = fetch_emails(mailbox, SUBJECT_FILTER, ATTACHMENT_FOLDER, limit=LIMIT)

print("Saving to Excel...")
save_emails_to_excel(emails, "output.xlsx")

# ✅ Wait for disk flush
time.sleep(2)

# ✅ Open the Excel file if it exists
excel_path = os.path.abspath("output.xlsx")
if os.path.exists(excel_path):
    print("📄 File saved at:", excel_path)
    os.startfile(excel_path)
else:
    print("⚠️ Excel file was not found.")

print("✅ All done!")

