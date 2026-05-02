import os
import smtplib
from email.message import EmailMessage

import pandas as pd


def load_env_file(env_path):
    """Load simple KEY=VALUE pairs from a .env file into os.environ."""
    if not os.path.exists(env_path):
        return

    with open(env_path, "r", encoding="utf-8") as env_file:
        for raw_line in env_file:
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue

            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value


load_env_file(os.path.join(os.path.dirname(__file__), ".env"))

# === CONFIGURATION ===
YOUR_EMAIL = os.environ.get("YOUR_EMAIL", "")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

if not YOUR_EMAIL or not APP_PASSWORD:
    raise ValueError(
         "Missing YOUR_EMAIL or APP_PASSWORD environment variables. "
         "Set them in the process environment or in the local .env file."
    )

# Filename of your CV (must be in the same directory as the script, or provide full path)
CV_FILE = "Thushan_Madarasinghe.pdf"
# Filename of your Excel file with emails (must be in the same directory as the script, or provide full path)
# Ensure the column containing emails is named exactly "Email Address" or update the script accordingly.
EXCEL_FILE = "test.xlsx"
# Filename of your Cover Letter (This variable is no longer used for attachment,
# but kept here for reference if you want to include the text in the body)
# COVER_LETTER_FILE = "Thushan_Madarasinghe_coverLetter.pdf" # Commented out as it's not attached

# --- Email Content ---
# Subject line for the email
SUBJECT = " Inquiry Regarding Internship Opportunities – Computer Science Undergraduate"

# HTML body of the email
# This includes the humanized cover letter text you refined earlier.
HTML_BODY = f"""
<html>
  <body style="font-family: Arial, sans-serif; line-height: 1.6;">
    <p>Dear Sir/Madam,</p>

    <p>Hope you're having a good week.</p>

    <p>My name is Thushan Madarasinghe, a Computer Science undergraduate at the Informatics Institute of Technology (IIT), affiliated with the University of Westminster. I'm writing to express my strong interest in an internship opportunity with your esteemed organization.</p>

    <p>I'm genuinely impressed by your company's commitment to excellence and nurturing new talent. I'm eager to contribute to the impactful work you're doing, which feels like a fantastic next step for me.</p>

    <p>My studies have provided a solid foundation in Object-Oriented Programming, Algorithms, and Data Structures. I'm proficient in technologies including Java, React Native, Node.js, Express.js, React, MongoDB, SQL, and HTML/CSS, and I'm comfortable with Git for version control. I also have a good understanding of backend development, including APIs and database management.</p>

    <p>I enjoy tackling new technologies and figuring things out across the full stack, bringing different components together to build functional solutions.</p>

    <p>Key projects I've worked on include:</p>
    <ul>
      <li><strong>GoviShakthi:</strong> An AI-powered MERN stack app with LLM integration for product recommendations.</li>
      <li><strong>FinTrack:</strong> A personal finance tracker built with the MERN stack.</li>
      <li><strong>Real-Time Ticketing System:</strong> Developed using Node.js, React.js, and WebSockets.</li>
      <li><strong>Plane Management System:</strong> A project built using Java.</li>
    </ul>

    <p>My involvement with the IEEE Computer Society at university has also enhanced my communication, organization, and teamwork skills through various events.</p>

    <p>I'm eager to bring my energy, technical skills, and passion for learning to your team. Please find my CV attached for your review. I would be grateful for the chance to discuss how I could contribute. If there aren't any suitable openings right now, I'd be thankful if you'd keep my application in mind and let me know about any future opportunities that might come up.</p>

    <p>Thank you for your time and consideration.</p>

    <p>Sincerely,</p>

    <p><strong>Thushan Madarasinghe</strong><br>
    +94 70 392 1791<br>
    <a href="mailto:{YOUR_EMAIL}">{YOUR_EMAIL}</a><br>
    <a href="https://github.com/ThushanMadu">GitHub</a> |
    <a href="https://thushanmadu.me">Portfolio</a> |
    <a href="https://linkedin.com/in/thushan-madarasinghe-420810222">LinkedIn</a>
    </p>
  </body>
</html>
"""

# --- Function to send a single email ---
def send_email(to_email):
    """Sends an email to a single recipient."""
    msg = EmailMessage()
    msg['Subject'] = SUBJECT
    msg['From'] = YOUR_EMAIL
    msg['To'] = to_email

    # Add the HTML body
    msg.add_alternative(HTML_BODY, subtype='html')

    # Attach CV file
    if not os.path.exists(CV_FILE):
        print(f"Error: CV file not found at {CV_FILE}")
        return False # Indicate failure
    with open(CV_FILE, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=os.path.basename(CV_FILE))

    # Send the message using SMTP
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(YOUR_EMAIL, APP_PASSWORD)
            smtp.send_message(msg)
        return True # Indicate success
    except Exception as e:
        print(f"SMTP Error: {e}")
        return False # Indicate failure


# === Main execution block ===
if __name__ == "__main__":
    # Check if Excel file exists
    if not os.path.exists(EXCEL_FILE):
        print(f"Error: Excel file not found at {EXCEL_FILE}")
    else:
        # Load emails from Excel
        try:
            data = pd.read_excel(EXCEL_FILE)

            # Add a 'Status' column if it doesn't exist
            if 'Status' not in data.columns:
                data['Status'] = '' # Initialize with empty strings

            # Check if the "Email Address" column exists
            if "Email Address" not in data.columns:
                print(f"Error: Column 'Email Address' not found in {EXCEL_FILE}. Please check the column name.")
            else:
                print(f"Loaded {len(data)} rows from {EXCEL_FILE}")
                # Iterate through each row (each email address)
                for index, row in data.iterrows():
                    email = row["Email Address"] # Get email from the specified column
                    # Only attempt to send if the email is not empty and status is not already 'Sent' or 'Failed'
                    if pd.notna(email) and str(email).strip() != "" and row['Status'] not in ['Sent', 'Failed']:
                        email_str = str(email).strip() # Ensure email is a string and strip whitespace
                        print(f"Attempting to send to: {email_str}")
                        if send_email(email_str):
                            print(f"✅ Sent successfully to {email_str}")
                            data.loc[index, 'Status'] = 'Sent' # Update status in DataFrame
                        else:
                            print(f"❌ Failed to send to {email_str}")
                            data.loc[index, 'Status'] = 'Failed' # Update status in DataFrame
                    elif pd.notna(email) and str(email).strip() != "":
                         print(f"Skipping row {index} ({email}): Status already '{row['Status']}'.")
                    else:
                        print(f"Skipping row {index}: No valid email address found.")

                # Save the updated DataFrame back to the Excel file
                try:
                    data.to_excel(EXCEL_FILE, index=False) # index=False prevents writing the DataFrame index as a column
                    print(f"Updated status in {EXCEL_FILE}")
                except Exception as e:
                    print(f"Error saving updated Excel file: {e}")


        except FileNotFoundError:
             # This case is already handled by os.path.exists, but good to have
            print(f"Error: Excel file not found at {EXCEL_FILE}")
        except Exception as e:
            print(f"An error occurred while processing the Excel file: {e}")

