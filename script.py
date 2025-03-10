import pymongo
import certifi
import pandas as pd
from datetime import datetime
import smtplib
import os
import json
from email.message import EmailMessage

# Load environment variables from GitHub Actions Secrets
MONGO_URI = os.getenv("MONGO_URI")
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT"))
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")

# Load department mappings from environment variable
dept_mappings = json.loads(os.getenv("DEPT_MAPPINGS"))

# MongoDB Connection
client = pymongo.MongoClient(MONGO_URI, tlsCAFile=certifi.where())
db = client["studentDB"]
collection = db["latecomers"]

# Fetch data from MongoDB
data = list(collection.find())
df = pd.DataFrame(data)
df.drop(['_id', '__v'], axis=1, inplace=True)
df['date'] = pd.to_datetime(df['date'])

# Filter today's latecomers
today_date = datetime.today().strftime('%Y-%m-%d')
df_today = df[df['date'].dt.strftime('%Y-%m-%d') == today_date]

# Generate Excel files for each department
saved_files = {}
for dept, email in dept_mappings.items():
    df_dept = df_today[df_today['department'] == dept]
    if not df_dept.empty:
        filename = f"{dept}_late_comers_{today_date}.xlsx"
        df_dept.to_excel(filename, index=False)
        saved_files[dept] = filename
        print(f"Saved: {filename}")

# Email sending function
def send_email(receiver_email, subject, body, attachment_path):
    msg = EmailMessage()
    msg["From"] = SENDER_EMAIL
    msg["To"] = receiver_email
    msg["Subject"] = subject
    msg.set_content(body)

    with open(attachment_path, "rb") as file:
        msg.add_attachment(file.read(), maintype="application",
                           subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           filename=os.path.basename(attachment_path))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
        print(f"Email sent to {receiver_email} with {attachment_path}")

# Send emails
for dept, email in dept_mappings.items():
    if email and dept in saved_files:
        send_email(email, f"Latecomers List - {dept} ({today_date})",
                   "Attached is the latecomers' list for today.", saved_files[dept])
