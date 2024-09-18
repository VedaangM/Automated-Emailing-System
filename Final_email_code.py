import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import openpyxl
def split_excel_file(file_path):
    # Load the Excel file
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # Extract data from each row and column
    data = []
    for row in sheet.iter_rows(values_only=True):
        row_data = []
        for cell in row:
            row_data.append(cell)
        data.append(row_data)

    return data
def list_files_by_time(directory):
    # Get a list of all files in the directory
    files = os.listdir(directory)

    # Create a list of file paths with their corresponding modification times
    file_times = [(os.path.join(directory, file), os.path.getmtime(os.path.join(directory, file))) for file in files]

    # Sort the list of file paths by modification time (oldest files first)
    sorted_files = sorted(file_times, key=lambda x: x[1])

    # Create an array to store the file paths
    file_paths = []

    # Append the file paths to the array
    for file_path, modification_time in sorted_files:
        file_paths.append(file_path)

    return file_paths
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

file_path = r'C:\Users\vedaa\Documents\project\pythonProject\Bluck_email\Bluck_email\~$WORK SHOP CERTIFICATE.xlsx'
result = split_excel_file(file_path)
folder_path = r'C:\Users\vedaa\Documents\project\pythonProject\Bluck_email\Bluck_email\Final_Certificate'
sorted_file_paths = list_files_by_time(folder_path)
for i in range(4,56):
    path= sorted_file_paths[i]
    print(path)
    name= result[i][1]
    email=result[i][3]
    Send=f'name={name} ad email={email} and path ={path}'
    print(Send)
    sender_email = 'innovation4u.2014@gmail.com'
    sender_password = 'cakkdaqdxaqfzikl'
    recipient_email = email
    subject = 'eCertificate and PPT Attachment'
    body = f'''
    Dear {name},
    Hello World
    '''
    certificate_path = path
    ppt_path = "xyz.pptx"

    # Send the email
    send_email(sender_email, sender_password, recipient_email, subject, body, certificate_path, ppt_path)
