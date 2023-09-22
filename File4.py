import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders

# Replace these placeholders with your email credentials
sender_email = 'spammailkb4@gmail.com'
sender_password = 'ekqd cjfa tjnf xeqk'

# Replace 'your_excel_file.xlsx' with the actual file path of your modified Excel sheet
excel_file_path = 'Details.xlsx'

# Load the Excel sheet into a Pandas DataFrame
try:
    df = pd.read_excel(excel_file_path)
except FileNotFoundError:
    print(f"Error: The Excel file '{excel_file_path}' was not found.")
    exit(1)

# Assuming your Excel columns are named 'Name', 'Email', and 'Photo1', 'Photo2', 'Photo3', update these column names if needed.
name_column = 'Name'
email_column = 'Email'

# Create an SMTP session for sending the email
try:
    session = smtplib.SMTP('smtp.gmail.com', 587)
    session.starttls()
    session.login(sender_email, sender_password)
except smtplib.SMTPAuthenticationError:
    print("Error: Failed to authenticate with the SMTP server. Check your email credentials.")
    exit(1)
except Exception as e:
    print(f"Error: An unexpected error occurred while setting up the SMTP session: {str(e)}")
    exit(1)

# Iterate through rows to extract information and send emails
for index, row in df.iterrows():
    name = row[name_column]
    email = row[email_column]
    
    # Create a multipart message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = email
    msg['Subject'] = 'Your Photos'

    # Add text to the email
    body = f"Hello {name},\n\nPlease find your photos attached."
    msg.attach(MIMEText(body, 'plain'))

    # Iterate through photo columns (assuming columns are named 'Photo1', 'Photo2', 'Photo3', etc.)
    for column in df.columns:
        if column.startswith('Photo'):
            photo_id = row[column]
            if not pd.isna(photo_id):
                # Extract the numeric part from the PhotoID
                numeric_photo_id = ''.join(filter(str.isdigit, str(photo_id)))
                # Construct the actual file name by adding the prefix
                actual_photo_id = f'KDS_{numeric_photo_id}.jpg'
                
                # Attach the corresponding photo to the email
                photo_file_path = f'D:/Documents/Works/Images/{actual_photo_id}'  # Replace with the actual path to your photos
                with open(photo_file_path, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload((attachment).read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f"attachment; filename={actual_photo_id}")
                    msg.attach(part)

    # Send the email
    try:
        text = msg.as_string()
        session.sendmail(sender_email, email, text)
        print(f"Email sent successfully to {email}")
    except Exception as e:
        print(f"Error: Failed to send email to {email}: {str(e)}")

# Quit the SMTP session
session.quit()