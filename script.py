import pymongo
import certifi
import pandas as pd
from datetime import datetime
import smtplib
import os
import json
from email.message import EmailMessage

MONGO_URI = os.getenv("MONGO_URI")
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT"))
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
ALL_MAIL = os.getenv("ALL_MAIL")
dept_mappings = json.loads(os.getenv("DEPT_MAPPINGS"))

client = pymongo.MongoClient(MONGO_URI, tlsCAFile=certifi.where())
db = client["studentDB"]
collection = db["latecomers"]

data = list(collection.find())
if not data:
    print("No data found. Skipping process.")
    exit()

df = pd.DataFrame(data)
df.drop(['_id', '__v'], axis=1, inplace=True, errors='ignore')
df['date'] = pd.to_datetime(df['date'])

today_date = datetime.today().strftime('%d-%m-%Y')
df_today = df[df['date'].dt.strftime('%d-%m-%Y') == today_date]

if df_today.empty:
    print(f"No latecomers found for {today_date}. Skipping further processing.")
    exit()

saved_files = {}
with pd.ExcelWriter(f"Latecomers_{today_date}.xlsx", engine="xlsxwriter") as writer:
    for dept, email in dept_mappings.items():
        df_dept = df_today[df_today['department'] == dept]
        if not df_dept.empty:
            df_dept.to_excel(writer, sheet_name=dept, index=False)
            saved_files[dept] = f"{dept}_late_comers_{today_date}.xlsx"
            df_dept.to_excel(saved_files[dept], index=False)

    consolidated_filename = f"Latecomers_{today_date}.xlsx"
    writer.close()

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

    os.remove(attachment_path)

for dept, email in dept_mappings.items():
    if email and dept in saved_files:
        send_email(email, f"Latecomers List - {dept} {today_date}",
                   "Attached is the latecomers' list for today.", saved_files[dept])

send_email(ALL_MAIL, f"Latecomers Consolidated Report {today_date}",
           "Attached is the consolidated latecomers' report for all departments.", consolidated_filename)
