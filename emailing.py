import smtplib #Used for sending email
from email.mime.multipart import MIMEMultipart #Allows to send bith text and attachment
from email.mime.text import MIMEText #Used to send plain text.
from email.mime.base import MIMEBase # handling binary file attachments (like PDFs, images, etc.) in your emails.
from email import encoders #

def send_email(sender_email, sender_password, recipient_email, subject, body, certificate_path, ppt_path):
    # SMTP server configuration
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587

    # Create the email message
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = recipient_email
    message['Subject'] = subject

    # Attach the body of the email
    message.attach(MIMEText(body, 'plain'))

    # Attach the e-certificate
    certificate_attachment = open(certificate_path, 'rb')
    certificate_mime = MIMEBase('application', 'octet-stream')
    certificate_mime.set_payload((certificate_attachment).read())
    encoders.encode_base64(certificate_mime)
    certificate_mime.add_header('Content-Disposition', 'attachment', filename=certificate_path)
    message.attach(certificate_mime)

    # Attach the PowerPoint presentation
    ppt_attachment = open(ppt_path, 'rb')
    ppt_mime = MIMEBase('application', 'octet-stream')
    ppt_mime.set_payload((ppt_attachment).read())
    encoders.encode_base64(ppt_mime)
    ppt_mime.add_header('Content-Disposition', 'attachment', filename=ppt_path)
    message.attach(ppt_mime)

    # Connect to the SMTP server and send the email
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(message)

    print("Email sent successfully!")

# Provide your email details and paths to the certificate and PPT
sender_email = 'xyz@gmail.com'
sender_password = 'xyz'
recipient_email = 'xyz@gmail.com'
subject = 'xyz'
body = f'''
Dear {Student},

Hello world
'''
certificate_path = 'xyz.jpg'
ppt_path = 'xyz.pptx'

# Send the email
send_email(sender_email, sender_password, recipient_email, subject, body, certificate_path, ppt_path)