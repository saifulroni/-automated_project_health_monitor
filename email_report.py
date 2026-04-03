import smtplib
from email.message import EmailMessage
import os

def send_email_report(sender_email, app_password, recipient_email, pdf_path):
    if not os.path.exists(pdf_path):
        print(f"PDF not found: {pdf_path}")
        return

    msg = EmailMessage()
    msg["Subject"] = "Weekly PMO Health Report"
    msg["From"] = sender_email
    msg["To"] = recipient_email
    msg.set_content(
        "Hello,\n\n"
        "Please find attached the latest weekly PMO health report PDF.\n\n"
        "Regards,\n"
        "Lemu Tajin Project Health Monitor"
    )

    with open(pdf_path, "rb") as f:
        file_data = f.read()
        file_name = os.path.basename(pdf_path)

    msg.add_attachment(
        file_data,
        maintype="application",
        subtype="pdf",
        filename=file_name
    )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(sender_email, app_password)
        smtp.send_message(msg)

    print(f"Email sent successfully to {recipient_email}")