import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage

# Load both sheets into DataFrames
# Replace the file path with a generic placeholder for the Excel file
sheets = pd.read_excel(r'C:\Path\To\Your\File\Company_Master_File.xlsx', sheet_name=['Contact Details', 'Confidential'])

# Access each sheet as a DataFrame
contact_details = sheets['Contact Details']
confidential = sheets['Confidential']

# Get today's date
today = datetime.now().date()

def send_wishes(firstname, lastname, to_email, cc_emails=None, image_path=None):
    from_email = "your_email@example.com"  # Replace with your email
    from_password = "your_email_password"  # Replace with your email password or app-specific password
    subj = "Best Wishes from [Your Company]"
    body = (f"Dear {firstname} {lastname},\n\n"
            f"Here's Wishing you a very Happy Birthday on behalf of our team at [Your Company]. "
            f"We hope you have an amazing day filled with joy and celebration!\n\n"
            f"Best regards,\n"
            f"The [Your Company] HR Team\n"
            f"www.companywebsite.com")
    
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subj
    
    # Add CC emails if provided
    if cc_emails:
        msg['Cc'] = ', '.join(cc_emails)  # Join CC emails with commas
        to_email = [to_email] + cc_emails  # Include CC in the recipient list

    msg.attach(MIMEText(body, 'plain'))

    if image_path:
        try:
            with open(image_path, 'rb') as img_file:
                image = MIMEImage(img_file.read())
                image.add_header('Content-Disposition', 'attachment', filename=image_path.split('/')[-1])
                msg.attach(image)
        except Exception as e:
            print(f"Error attaching image: {e}")

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, from_password)
        text = msg.as_string()
        server.sendmail(from_email, to_email, text)  # Send to all recipients
        server.quit()
        return "Email sent successfully!"
    except Exception as e:
        return f"Failed to send email. Error: {e}"

# Specify dummy CC emails
cc_emails = ['dummy_cc1@example.com', 'dummy_cc2@example.com']

# Replace the image path with a generic placeholder
image_path = r'C:\Path\To\Your\Image\Birthday_Card.jpg'

# Iterate through the Confidential sheet to find matching DOBs
for index, row in confidential.iterrows():
    emp_id = row['Emp_Id']
    # Handle DOB format (e.g., "October 24, 1985")
    dob = pd.to_datetime(row['DOB'], format='%B %d, %Y').date()
    firstname = row['First Name']
    lastname = row['Last Name']

    # Check if today matches the birthday
    if dob.month == today.month and dob.day == today.day:
        # Match the Emp_Id with the Contact Details sheet to get the email
        email_row = contact_details[contact_details['Emp_Id'] == emp_id]
        
        if not email_row.empty:
            to_email = email_row['P_Email1'].values[0]  # Get the email address
            # Send birthday wishes with CC and image attachment
            result = send_wishes(firstname, lastname, to_email, cc_emails, image_path)
            print(f"Birthday wishes sent to {firstname} {lastname}!")
        else:
            print(f"No email found for Emp_Id: {emp_id}")
