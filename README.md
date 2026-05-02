# Automated Email Send Script

A Python-based bulk email sending tool that reads recipient addresses from an Excel spreadsheet and sends personalized HTML emails with PDF attachments via Gmail SMTP.

## Features

- Reads email addresses from Excel files
- Sends personalized HTML-formatted emails
- Attaches PDF documents (CV/resume) automatically
- Tracks send status (Sent/Failed) in the Excel file
- Safe to re-run without sending duplicate emails
- Loads credentials from `.env` file

## Prerequisites

- Python 3.7 or higher
- A Gmail account with 2-Factor Authentication enabled
- Gmail App Password (not your regular Gmail password)

## Setup

### 1. Clone the Repository

```bash
git clone https://github.com/ThushanMadu/automated_emailSend_script.git
cd automated_emailSend_script
```

### 2. Install Dependencies

```bash
pip install pandas openpyxl
```

### 3. Create `.env` File

Create a `.env` file in the project root with your Gmail credentials:

```env
YOUR_EMAIL=yourgmail@gmail.com
APP_PASSWORD=your16characterapppassword
```

**Important:** The `.env` file is gitignored to protect your credentials.

### 4. Generate Gmail App Password

1. Go to your Google Account settings
2. Navigate to Security > 2-Step Verification (must be enabled)
3. Go to App Passwords
4. Generate a new app password for "Mail"
5. Use this 16-character password in your `.env` file

### 5. Prepare Your Files

Place these files in the project root:

- **Excel file** (`test.xlsx`): Must contain a column named exactly `"Email Address"`
- **Your CV**: Add your CV/resume as a PDF file and update the `CV_FILE` variable in `emails.py` with the filename

## Usage

Run the script:

```bash
python emails.py
```

The script will:
1. Load all email addresses from the Excel file
2. Send emails one by one via Gmail SMTP
3. Update the Excel file with send status (Sent/Failed)
4. Skip already processed emails on re-runs

## Configuration

Edit these variables in `emails.py` to customize:

| Variable | Description | Default |
|----------|-------------|---------|
| `CV_FILE` | Path to your CV PDF | `your_cv.pdf` |
| `EXCEL_FILE` | Path to your email list | `test.xlsx` |
| `SUBJECT` | Email subject line | Internship inquiry subject |
| `HTML_BODY` | Email content (HTML) | Cover letter template |

## Excel File Format

Your Excel file should have at least this column:

| Email Address |
|---------------|
| company1@email.com |
| company2@email.com |

The script automatically adds a `Status` column to track sent emails.

## Project Structure

```
automated_emailSend_script/
├── emails.py                    # Main script
├── test.xlsx                    # Recipient email list
├── your_cv.pdf                  # Your CV attachment (add yours)
├── .env                         # Credentials (create this)
├── .gitignore
└── README.md
```

## Security Notes

- Never commit your `.env` file
- Keep your Gmail App Password secure
- The script uses SMTP SSL (port 465) for encrypted connections

## Troubleshooting

**"Missing YOUR_EMAIL or APP_PASSWORD"**
- Check that your `.env` file exists and contains both variables

**"Column 'Email Address' not found"**
- Ensure your Excel column is named exactly `"Email Address"`

**SMTP authentication errors**
- Verify your App Password is correct
- Ensure 2-Factor Authentication is enabled on your Google account

## License

MIT License - see [LICENSE](LICENSE) file for details

## Author

Thushan Madarasinghe

- GitHub: https://github.com/ThushanMadu
- Portfolio: https://thushanmadu.me
- LinkedIn: https://linkedin.com/in/thushan-madarasinghe-420810222
