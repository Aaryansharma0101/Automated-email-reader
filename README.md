This project automates the process of reading emails from both Inbox and Spam folders using IMAP. It supports filtering emails by subject keywords, extracting unique types of content, handling attachments, and exporting email data to an Excel file for analysis.

🚀 Features
✅ Reads both read and unread emails.

🔍 Filters emails based on specific subject keywords.

📦 Supports attachments (e.g., PDFs, images, etc.).

🗂️ Extracts and saves unique content types only (e.g., avoiding duplicates).

📤 Reads from Inbox and Spam folders.

📊 Exports parsed data to Excel (.xlsx) for offline use.

💡 Ideal for automating inbox triage or large-scale email parsing.

🛠️ Tech Stack
Python 3.x

imaplib, email, openpyxl, os, re, base64, datetime# Automated-email-reader
