import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def get_email_addresses(file_path, sheet_name, column_index):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb[sheet_name]

    email_addresses = []
    for row in sheet.iter_rows(values_only=True):
        email = row[column_index - 1]  
        email_addresses.append(email)

    return email_addresses[1:] 


def send_email(sender_email, sender_password, receiver_email, subject, body):
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = subject

    message.attach(MIMEText(body, 'plain'))

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, message.as_string())


sender_email = ''
sender_password = ''

subject = 'Your Subject'
body = 'Your Email Body'

file_path = r''
sheet_name = 'Sheet1'
column_index = 1  


email_addresses = get_email_addresses(file_path, sheet_name, column_index)

for receiver_email in email_addresses:
    send_email(sender_email, sender_password, receiver_email, subject, body)
    print(f"Email sent to {receiver_email}")

print("All emails sent successfully.")