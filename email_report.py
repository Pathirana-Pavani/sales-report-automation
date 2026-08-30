"""Sends the generated report files by email via Gmail SMTP.

Sender credentials come from Streamlit secrets (st.secrets), never from
user input or source code — configure them in Streamlit Cloud's app
settings -> Secrets, or locally in .streamlit/secrets.toml (gitignored):

    gmail_email = "your-sender@gmail.com"
    gmail_app_password = "xxxx xxxx xxxx xxxx"

The app password comes from Google Account -> Security -> App passwords
(requires 2-Step Verification enabled on that Gmail account).
"""
import smtplib
from email.message import EmailMessage


def send_report_email(sender_email, sender_password, recipient_email, subject, body, attachments):
    """attachments: list of (filename, bytes, mime_maintype, mime_subtype)."""
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = recipient_email
    msg.set_content(body)

    for filename, data, maintype, subtype in attachments:
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(sender_email, sender_password)
        smtp.send_message(msg)
