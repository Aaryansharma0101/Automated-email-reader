import streamlit as st
from email_reader import connect_to_mailbox, fetch_emails
from utils import save_emails_to_excel
import os
import time

# Page config
st.set_page_config(page_title="📧 Gmail to Excel Exporter", layout="centered")

# Title
st.title("📬 Gmail to Excel Exporter")
st.caption("Extract emails with specific subjects and save them to Excel")

# Inputs
with st.form("email_form"):
    email = st.text_input("📧 Gmail Address")
    password = st.text_input("🔑 App Password", type="password")
    subject_input = st.text_input("🔍 Subject Filter (comma-separated)", value="GitHub")
    limit = st.number_input("📦 Number of Emails to Fetch", min_value=1, value=30, step=1)

    submitted = st.form_submit_button("🚀 Fetch Emails")

if submitted:
    if not email or not password:
        st.error("Please enter both your email and password.")
    else:
        try:
            with st.spinner("Connecting and fetching emails..."):
                mailbox = connect_to_mailbox(email, password)
                subjects = [s.strip() for s in subject_input.split(",") if s.strip()]
                emails = fetch_emails(mailbox, subjects, "attachments", limit)

                if emails:
                    output_file = "output.xlsx"
                    save_emails_to_excel(emails, output_file)
                    time.sleep(1)

                    st.success(f"{len(emails)} emails saved to {output_file}")

                    with open(output_file, "rb") as f:
                        st.download_button("📥 Download Excel File", f, file_name=output_file)

                else:
                    st.warning("No matching emails found.")
        except Exception as e:
            st.error(f"❌ Error: {e}")
