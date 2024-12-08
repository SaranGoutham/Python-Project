# Birthday Wishes Automation Script

This script automates the process of sending birthday wishes to employees using their details stored in an Excel file. It utilizes Python libraries such as `pandas`, `datetime`, and `smtplib` for data manipulation and email handling.

---

## Features
- Reads employee data from two sheets in an Excel file:  
  - **Contact Details**: Contains email addresses and other contact information.  
  - **Confidential**: Contains employee details like `Emp_Id`, `First Name`, `Last Name`, and `DOB`.
- Matches today's date with employees' birthdays and sends personalized birthday wishes via email.
- Supports CC recipients and image attachments (e.g., a birthday card).
- Provides feedback on the success or failure of email delivery.

---

## Requirements
- Python 3.7+
- Libraries:
  - `pandas`
  - `datetime`
  - `smtplib`
  - `email`
- Access to an SMTP server (e.g., Gmail) for sending emails.

---

## Setup

### 1. Install Dependencies
Install the required Python libraries using pip:
```bash
pip install pandas
```

### 2. Configure the Script
- Update the file paths:
  - Replace `C:\Path\To\Your\File\Company_Master_File.xlsx` with the path to your Excel file.
  - Replace `C:\Path\To\Your\Image\Birthday_Card.jpg` with the path to the birthday card image.
- Replace placeholder values with your details:
  - `your_email@example.com`: Your email address.
  - `your_email_password`: Your email password or app-specific password.
  - Add CC email addresses if needed.

### 3. Excel File Structure
Ensure the Excel file has the following structure:

#### **Sheet: Contact Details**
| Emp_Id | P_Email1            | Other Columns |
|--------|---------------------|---------------|
| 101    | john.doe@example.com| ...           |

#### **Sheet: Confidential**
| Emp_Id | First Name | Last Name | DOB            | Other Columns |
|--------|------------|-----------|----------------|---------------|
| 101    | John       | Doe       | October 24, 1985| ...           |

---

## Usage
1. Run the script:
   ```bash
   python birthday_wishes.py
   ```
2. The script will:
   - Check today's date against the DOB column in the Confidential sheet.
   - Match employee IDs to find email addresses in the Contact Details sheet.
   - Send a personalized email with the birthday wish, including CC recipients and image attachments.

---

## Troubleshooting
- **SMTP Error:** Ensure you have enabled "Less Secure Apps" or used an app-specific password for your email account.
- **Date Format Issues:** Confirm that the DOB in the Excel file matches the `'%B %d, %Y'` format (e.g., "October 24, 1985").
- **File Path Issues:** Use absolute paths for the Excel file and image to avoid errors.

---

## Disclaimer
- This script is for educational and internal use only. Ensure compliance with data privacy policies before using it in production.
