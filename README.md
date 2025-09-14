# Birthday Email System

A comprehensive Python application that automatically sends personalized birthday emails to employees across multiple companies. The system supports company-specific branding, SMTP configurations, and includes robust logging and error handling.

## Features

- **Multi-Company Support**: Configure up to 4 different companies with unique SMTP settings and branding
- **Smart Birthday Detection**: Automatically identifies today's birthdays from Excel data
- **Employee Status Filtering**: Filter by Active, Terminated, or Both employee statuses
- **Company-Specific Branding**: Custom email templates, images, and sender information per company
- **Email Validation & Deduplication**: Ensures clean recipient lists
- **Comprehensive Logging**: Detailed logs with CSV tracking of all send attempts
- **Rate Limiting**: Built-in delays to prevent spam detection
- **Dry Run Mode**: Test configuration without sending actual emails
- **Error Recovery**: Continues processing even if individual sends fail

## Requirements

- Python 3.7+
- pandas
- python-dotenv
- openpyxl (for Excel file reading)

Install dependencies:
```bash
pip install pandas python-dotenv openpyxl
```

## Project Structure

```
birthday-email-system/
├── birthday_email_system.py    # Main application
├── .env                       # Environment configuration
├── birthday_emails.log        # Application logs
├── C:/logs/                   # Send attempt logs (CSV format)
├── assets/                    # Company birthday card images
│   ├── Company1BdayCard.jpg
│   ├── Company2BdayCard.jpg
│   ├── Company3BdayCard.jpg
│   └── Company4BdayCard.jpg
└── data/                      # Excel files with employee data
    ├── COMPANY1_MASTER_EXCEL_HR_2025.xlsx
    ├── COMPANY2_MASTER_EXCEL_HR_2025.xlsx
    ├── COMPANY3_MASTER_EXCEL_HR_2025.xlsx
    └── COMPANY4_MASTER_EXCEL_HR_2025.xlsx
```

## Excel File Format

Each company's Excel file must contain three sheets:

### 1. Confidential Sheet
Required columns:
- `Emp_Id`: Employee ID
- `First_Name`: Employee's first name
- `Last_Name`: Employee's last name
- `DOB`: Date of birth (any recognizable date format)

### 2. Contact Details Sheet
Required columns:
- `Emp_Id`: Employee ID (matching Confidential sheet)
- `First_Name`: Employee's first name
- `Last_Name`: Employee's last name
- `P_Email1`: Primary email address

### 3. Employee Status Sheet
Required columns:
- `Emp_Id`: Employee ID (matching other sheets)
- `First_Name`: Employee's first name
- `Last_Name`: Employee's last name
- `P_Status`: Employee status ('A' for Active, 'T' for Terminated)

## Configuration

### Step 1: Create Environment File

Copy `.example.env` to `.env` and configure your settings:

```bash
cp .example.env .env
```

### Step 2: Global Configuration

```env
# =========================================================================
# === GLOBAL CONFIGURATION ===
# These settings apply to all companies unless overridden.
# =========================================================================

# Optional generic attachment path
ATTACH_PATH=""

# Pacing settings (in seconds)
DELAY_BETWEEN_SENDS=0.5
DELAY_BETWEEN_COMPANIES=2.5

# Employee status filter: 'A' (Active), 'T' (Terminated), or 'BOTH'
P_STATUS_FILTER=A

# Global email defaults
EMAIL_CC=hr@yourcompany.com
EMAIL_BCC=security@yourcompany.com
TEAM_NAME_TEMPLATE="{company} HR Team"
SUBJECT_TEMPLATE="🎉 Happy Birthday, {first_name}!"
EMAIL_REPUTATION_DOMAIN="yourcompany.com"

# Test mode - set to "false" for live operation
DRY_RUN=true
```

### Step 3: Company-Specific Configuration

Configure each company with their unique SMTP settings:

```env
# --- Company 1 ---
COMPANY1_SMTP_HOST="smtp.company1.com"
COMPANY1_SMTP_PORT=587
COMPANY1_SMTP_USER="birthday@company1.com"
COMPANY1_SMTP_PASS="your_secure_password"
COMPANY1_EMAIL_CC="hr@company1.com"
COMPANY1_EMAIL_BCC=""
COMPANY1_TEAM_NAME_TEMPLATE="The {company} HR Team"
COMPANY1_SUBJECT_TEMPLATE="Happy Birthday, {first_name}! From the {company} Team"
COMPANY1_EMAIL_REPUTATION_DOMAIN="company1.com"

# --- Company 2 ---
COMPANY2_SMTP_HOST="smtp.company2.com"
COMPANY2_SMTP_PORT=587
COMPANY2_SMTP_USER="noreply@company2.net"
COMPANY2_SMTP_PASS="your_secure_password"
COMPANY2_EMAIL_CC="admin@company2.net, manager@company2.net"
COMPANY2_SUBJECT_TEMPLATE="A {company} Birthday Wish for You, {first_name}!"

# Continue for Company 3 and Company 4...
```

### Step 4: Configure File Paths

Update the paths in `birthday_email_system.py`:

```python
# Update company image paths
self.company_images = {
    'Company1': r"C:\path\to\your\assets\Company1BdayCard.jpg",
    'Company2': r"C:\path\to\your\assets\Company2BdayCard.jpg",
    'Company3': r"C:\path\to\your\assets\Company3BdayCard.jpg",
    'Company4': r"C:\path\to\your\assets\Company4BdayCard.jpg",
}

# Update Excel file paths in main()
excel_files = [
    r"C:\path\to\your\excel_files\COMPANY1_MASTER_EXCEL_HR_2025.xlsx",
    r"C:\path\to\your\excel_files\COMPANY2_MASTER_EXCEL_HR_2025.xlsx",
    r"C:\path\to\your\excel_files\COMPANY3_MASTER_EXCEL_HR_2025.xlsx",
    r"C:\path\to\your\excel_files\COMPANY4_MASTER_EXCEL_HR_2025.xlsx"
]
```

## Usage

### Testing (Dry Run Mode)

First, test your configuration without sending actual emails:

```bash
# Ensure DRY_RUN=true in your .env file
python birthday_email_system.py
```

### Live Operation

When ready for live operation:

1. Set `DRY_RUN=false` in your `.env` file
2. Run the system:

```bash
python birthday_email_system.py
```

### Automated Daily Execution

Set up a scheduled task or cron job to run daily:

**Windows Task Scheduler:**
- Program: `python`
- Arguments: `C:\path\to\birthday_email_system.py`
- Schedule: Daily at your preferred time

**Linux/Mac Cron:**
```bash
# Run daily at 9:00 AM
0 9 * * * /usr/bin/python3 /path/to/birthday_email_system.py
```

## Email Template Customization

### Template Variables

Available placeholders in templates:
- `{company}`: Company name (detected from filename)
- `{first_name}`: Employee's first name

### Email Structure

Each birthday email includes:
- **Plain text version** for compatibility
- **HTML version** with professional styling
- **Company-specific birthday card** (attached image)
- **Optional generic attachment** (if ATTACH_PATH is set)

### Sample Email Content

```
Subject: 🎉 Happy Birthday, John!

Dear John,

Here's wishing you a very Happy Birthday on behalf of our Company1 Team. 
Hope you are having a Blast !!

Be Blessed! Have a Great day!

The Company1 HR Team
WWW.Company1.com
```

## Logging and Monitoring

### Application Logs

- **File**: `birthday_emails.log`
- **Content**: Detailed application flow, errors, and processing information
- **Rotation**: Appends to existing log file

### Send Attempt Logs

- **Location**: `C:/logs/birthday_sends_YYYY-MM-DD.csv`
- **Content**: CSV format with columns:
  - Timestamp
  - Recipient
  - First_Name
  - Source_File
  - Company
  - Status (Sent/Failed/Dry Run)
  - Response
  - Message_ID
  - Spam_Score

### Monitoring Success

Check the logs to verify:
1. **File Processing**: Confirm all Excel files were loaded successfully
2. **Birthday Detection**: Verify today's birthdays were identified
3. **Email Validation**: Check for any invalid email addresses
4. **Send Status**: Review success/failure rates per company

## SMTP Configuration

### Supported SMTP Settings

- **Port 587**: STARTTLS (most common)
- **Port 465**: SSL/TLS
- **Authentication**: Username/password required

### Common SMTP Providers

**Gmail:**
```env
SMTP_HOST="smtp.gmail.com"
SMTP_PORT=587
SMTP_USER="your_email@gmail.com"
SMTP_PASS="your_app_password"  # Use App Password, not regular password
```

**Outlook/Office 365:**
```env
SMTP_HOST="smtp.office365.com"
SMTP_PORT=587
SMTP_USER="your_email@yourdomain.com"
SMTP_PASS="your_password"
```

**Custom/Corporate SMTP:**
```env
SMTP_HOST="mail.yourcompany.com"
SMTP_PORT=587  # or 465 for SSL
SMTP_USER="notifications@yourcompany.com"
SMTP_PASS="your_secure_password"
```

## Security Considerations

### Environment Variables

- Store sensitive credentials in `.env` file
- **Never commit `.env` file to version control**
- Add `.env` to your `.gitignore` file

### Email Security

- Use dedicated email accounts for automated sending
- Enable App Passwords where required (Gmail, etc.)
- Configure SPF/DKIM records for better deliverability
- Set appropriate `EMAIL_REPUTATION_DOMAIN` for Message-ID alignment

### Data Protection

- Ensure Excel files are stored securely
- Limit access to log files containing email addresses
- Consider encryption for sensitive employee data

## Troubleshooting

### Common Issues

**1. SMTP Connection Failed**
```
Error: SMTP connection failed for Company1: [Errno 11001] getaddrinfo failed
```
- Verify SMTP host and port settings
- Check network connectivity
- Confirm firewall isn't blocking SMTP ports

**2. Authentication Failed**
```
Error: (535, '5.7.3 Authentication unsuccessful')
```
- Verify username and password
- Check if App Password is required
- Ensure SMTP user has send permissions

**3. No Birthdays Found**
```
Info: No birthdays on 2025-09-14. Exiting.
```
- Verify DOB column format in Excel
- Check if today's date matches any birthdays
- Confirm P_STATUS_FILTER settings

**4. Invalid Email Addresses**
```
Warning: People with invalid email addresses:
   Emp_Id=123, Name=John, Email=invalid_email
```
- Clean up email data in Excel files
- Check for typos in P_Email1 column

### Debug Mode

For detailed debugging:

1. Set `DRY_RUN=true`
2. Check `birthday_emails.log` for detailed processing information
3. Verify configuration loading in logs
4. Review recipient lists before live sending

### Log Analysis

Monitor these log patterns:

- **Success**: `"Sent birthday email to email@domain.com"`
- **Connection**: `"Connected to SMTP server for Company1"`
- **Validation**: `"After email validation and deduplication: X recipients"`
- **Errors**: `"Error"` or `"Failed"` entries

## Support

For issues or questions:

1. Check the logs first (`birthday_emails.log`)
2. Verify configuration in `.env` file
3. Test with `DRY_RUN=true` mode
4. Review Excel file format and data quality

## License

This project is provided as-is for internal business use. Modify and adapt as needed for your organization's requirements.
