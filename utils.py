import os
import pandas as pd
import subprocess
from email.header import decode_header

# 🔍 Decode MIME-encoded words (like Subject, From)
def decode_mime_words(s):
    if not s:
        return ""
    decoded = decode_header(s)
    return ''.join(
        str(part[0], part[1] or 'utf-8') if isinstance(part[0], bytes) else str(part[0])
        for part in decoded
    )

# 📎 Save attachments and return file path
def save_attachment(part, folder):
    if not os.path.exists(folder):
        os.makedirs(folder)

    filename = decode_mime_words(part.get_filename())
    filepath = os.path.join(folder, filename)

    with open(filepath, "wb") as f:
        f.write(part.get_payload(decode=True))

    return filepath

# 💾 Save email data to Excel and open it automatically
def save_emails_to_excel(data, output_file="output.xlsx"):
    try:
        if not data:
            print("⚠️ No emails to save.")
            return

        # Create DataFrame
        df = pd.DataFrame(data)

        # Save using ExcelWriter (safer for formatting/multiple sheets)
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        abs_path = os.path.abspath(output_file)
        print(f"📄 File saved at: {abs_path}")

        # ✅ Open the file using subprocess
        try:
            subprocess.Popen(['start', abs_path], shell=True)
            print("📊 Excel file opened successfully!")
        except Exception as e:
            print(f"⚠️ Could not open Excel automatically: {e}")
            print(f"➡️ Please open it manually from: {abs_path}")

    except Exception as e:
        print(f"❌ Failed to save Excel: {e}")
