import smtplib
from email.message import EmailMessage

def send_email():
    # Set up sender and recipient
    sender_email = "lumi@mmt.my"
    recipient_email = "shanjif@mmt.my"
    subject = "Test Email from Python"
    body = "Hello! This is a test email sent using Python and Office365 SMTP."

    # Create the email
    msg = EmailMessage()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.set_content(body)

    # Office365 SMTP server details
    smtp_server = "smtp.office365.com"
    smtp_port = 587

    # Login credentials (use app password if MFA enabled)
    password = "Paranskanda@33"

    # Send the email
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Secure the connection
            server.login(sender_email, password)
            server.send_message(msg)
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")
