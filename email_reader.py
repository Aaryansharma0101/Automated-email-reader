import imaplib
import email
from email.header import decode_header
from utils import decode_mime_words, save_attachment
import os

def connect_to_mailbox(email_user, email_pass):
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(email_user, email_pass)
    return mail

# ✅ Updated: Added `limit=30` as a parameter
def fetch_emails(mail, subject_filter, attachment_dir, limit=30):
    seen_ids = set()
    results = []

    for folder in ['INBOX', '[Gmail]/Spam']:
        mail.select(folder)
        status, data = mail.search(None, "ALL")
        if status != "OK":
            continue

        # ✅ Updated: use the limit provided by user
        email_ids = data[0].split()[-limit:]
        print(f"📨 Scanning {len(email_ids)} latest emails in folder: {folder}")

        for num in email_ids:
            try:
                _, msg_data = mail.fetch(num, "(RFC822)")
                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)

                subject = decode_mime_words(msg["Subject"])
                from_ = decode_mime_words(msg.get("From"))
                to_ = decode_mime_words(msg.get("To"))
                msg_id = msg.get("Message-ID")

                # ✅ Subject filter supports multiple keywords
                if not subject or not any(word.lower() in subject.lower() for word in subject_filter):
                    continue
                if msg_id in seen_ids:
                    continue
                seen_ids.add(msg_id)

                body = ""
                attachment_path = ""

                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))

                        if "attachment" in content_disposition:
                            attachment_path = save_attachment(part, attachment_dir)
                        elif content_type == "text/plain":
                            try:
                                body = part.get_payload(decode=True).decode()
                            except:
                                pass
                else:
                    body = msg.get_payload(decode=True).decode()

                results.append({
                    "FROM": from_,
                    "TO": to_,
                    "SUBJECT": subject,
                    "BODY": body,
                    "ATTACHMENT": attachment_path
                })

                print(f"✔️ {subject}")

            except Exception as e:
                print(f"❌ Failed to fetch email: {e}")

    return results
