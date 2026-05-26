import os
import glob
import smtplib

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

sender = os.environ['SMTP_USERNAME']
password = os.environ['SMTP_PASSWORD']

to_emails = os.environ['RECEIVER_EMAIL'].split(',')

cc_emails = os.environ.get('CC_EMAIL', '').split(',')

all_recipients = to_emails + cc_emails

msg = MIMEMultipart()

msg['From'] = sender

msg['To'] = ', '.join(to_emails)

msg['Cc'] = ', '.join(cc_emails)

msg['Subject'] = 'Azure Advisor Weekly Report'

body = """
Hello,

Please find attached Azure Advisor Report.

Regards,
Cloud Operations Team
"""

msg.attach(MIMEText(body, 'plain'))

excel_file = glob.glob("*.xlsx")[0]

with open(excel_file, "rb") as attachment:

    part = MIMEBase("application", "octet-stream")

    part.set_payload(attachment.read())

encoders.encode_base64(part)

part.add_header(
    "Content-Disposition",
    f"attachment; filename={excel_file}",
)

msg.attach(part)

server = smtplib.SMTP(
    'smtp.office365.com',
    587
)

server.starttls()

server.login(sender, password)

server.sendmail(
    sender,
    all_recipients,
    msg.as_string()
)

server.quit()

print("Mail Sent Successfully")