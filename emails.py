"""
Automated email sender that reads recipient addresses from an Excel file
and sends a personalised HTML email with a CV attachment to each one.
"""

import logging
import os
import re
import smtplib
import time
from dataclasses import dataclass
from email.message import EmailMessage
from typing import Optional

import pandas as pd
from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
EMAIL_COLUMN = "Email Address"
STATUS_SENT = "Sent"
STATUS_FAILED = "Failed"
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465
SEND_DELAY_SECONDS = 2          # Pause between sends to avoid rate-limiting
EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
@dataclass
class Config:
    """All runtime settings loaded from the environment / .env file."""
    your_email: str
    app_password: str
    cv_file: str
    excel_file: str
    subject: str

    @classmethod
    def from_env(cls) -> "Config":
        """Load and validate configuration from environment variables."""
        load_dotenv()

        your_email = os.getenv("YOUR_EMAIL", "")
        app_password = os.getenv("APP_PASSWORD", "")

        if not your_email or not app_password:
            raise ValueError(
                "YOUR_EMAIL and APP_PASSWORD must be set in the .env file."
            )

        return cls(
            your_email=your_email,
            app_password=app_password,
            cv_file=os.getenv("CV_FILE", "CV.pdf"),
            excel_file=os.getenv("EXCEL_FILE", "emails.xlsx"),
            subject=os.getenv(
                "SUBJECT",
                "Inquiry Regarding Internship Opportunities – Computer Science Undergraduate",
            ),
        )


# ---------------------------------------------------------------------------
# Email content
# ---------------------------------------------------------------------------
HTML_BODY = """
<html>
  <body style="font-family: Arial, sans-serif; line-height: 1.6;">
    <p>Dear Sir/Madam,</p>

    <p>Hope you're having a good week.</p>

    <p>My name is Thushan Madarasinghe, a Computer Science undergraduate at the
    Informatics Institute of Technology (IIT), affiliated with the University of
    Westminster. I'm writing to express my strong interest in an internship
    opportunity with your esteemed organization.</p>

    <p>I'm genuinely impressed by your company's commitment to excellence and
    nurturing new talent. I'm eager to contribute to the impactful work you're
    doing, which feels like a fantastic next step for me.</p>

    <p>My studies have provided a solid foundation in Object-Oriented Programming,
    Algorithms, and Data Structures. I'm proficient in technologies including Java,
    React Native, Node.js, Express.js, React, MongoDB, SQL, and HTML/CSS, and I'm
    comfortable with Git for version control. I also have a good understanding of
    backend development, including APIs and database management.</p>

    <p>I enjoy tackling new technologies and figuring things out across the full
    stack, bringing different components together to build functional solutions.</p>

    <p>Key projects I've worked on include:</p>
    <ul>
      <li><strong>GoviShakthi:</strong> An AI-powered MERN stack app with LLM
          integration for product recommendations.</li>
      <li><strong>FinTrack:</strong> A personal finance tracker built with the
          MERN stack.</li>
      <li><strong>Real-Time Ticketing System:</strong> Developed using Node.js,
          React.js, and WebSockets.</li>
      <li><strong>Plane Management System:</strong> A project built using Java.</li>
    </ul>

    <p>My involvement with the IEEE Computer Society at university has also enhanced
    my communication, organization, and teamwork skills through various events.</p>

    <p>I'm eager to bring my energy, technical skills, and passion for learning to
    your team. Please find my CV attached for your review. I would be grateful for
    the chance to discuss how I could contribute. If there aren't any suitable
    openings right now, I'd be thankful if you'd keep my application in mind and
    let me know about any future opportunities that might come up.</p>

    <p>Thank you for your time and consideration.</p>

    <p>Sincerely,</p>

    <p><strong>Thushan Madarasinghe</strong><br>
    +94 70 392 1791<br>
    <a href="mailto:thushanmadu2003@gmail.com">thushanmadu2003@gmail.com</a><br>
    <a href="https://github.com/ThushanMadu">GitHub</a> |
    <a href="https://thushanmadu.me">Portfolio</a> |
    <a href="https://linkedin.com/in/thushan-madarasinghe-420810222">LinkedIn</a>
    </p>
  </body>
</html>
"""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def is_valid_email(address: str) -> bool:
    """Return True if *address* looks like a well-formed email address."""
    return bool(EMAIL_REGEX.match(address))


def build_message(cfg: Config, to_email: str) -> Optional[EmailMessage]:
    """
    Construct and return an EmailMessage ready to send.

    Returns None if the CV attachment file cannot be found.
    """
    if not os.path.exists(cfg.cv_file):
        logger.error("CV file not found: %s", cfg.cv_file)
        return None

    msg = EmailMessage()
    msg["Subject"] = cfg.subject
    msg["From"] = cfg.your_email
    msg["To"] = to_email
    msg.add_alternative(HTML_BODY, subtype="html")

    with open(cfg.cv_file, "rb") as fh:
        msg.add_attachment(
            fh.read(),
            maintype="application",
            subtype="pdf",
            filename=os.path.basename(cfg.cv_file),
        )

    return msg


def send_email(cfg: Config, to_email: str) -> bool:
    """
    Send a single email to *to_email*.

    Returns True on success, False on any failure.
    """
    msg = build_message(cfg, to_email)
    if msg is None:
        return False

    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as smtp:
            smtp.login(cfg.your_email, cfg.app_password)
            smtp.send_message(msg)
        return True
    except smtplib.SMTPAuthenticationError:
        logger.error("Authentication failed. Check YOUR_EMAIL and APP_PASSWORD.")
        return False
    except smtplib.SMTPException as exc:
        logger.error("SMTP error sending to %s: %s", to_email, exc)
        return False


def load_recipients(excel_file: str) -> pd.DataFrame:
    """
    Load the recipients DataFrame from *excel_file*.

    Raises FileNotFoundError or ValueError on problems.
    """
    if not os.path.exists(excel_file):
        raise FileNotFoundError(f"Excel file not found: {excel_file}")

    data = pd.read_excel(excel_file)

    if EMAIL_COLUMN not in data.columns:
        raise ValueError(
            f"Column '{EMAIL_COLUMN}' not found in {excel_file}. "
            "Please check the column name."
        )

    if "Status" not in data.columns:
        data["Status"] = ""

    return data


def save_recipients(data: pd.DataFrame, excel_file: str) -> None:
    """Persist the updated *data* back to *excel_file*."""
    data.to_excel(excel_file, index=False)
    logger.info("Updated status saved to %s", excel_file)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> None:
    """Entry point: load config, iterate recipients, send emails."""
    cfg = Config.from_env()

    try:
        data = load_recipients(cfg.excel_file)
    except (FileNotFoundError, ValueError) as exc:
        logger.error("%s", exc)
        return

    logger.info("Loaded %d rows from %s", len(data), cfg.excel_file)

    sent = skipped = failed = 0

    for index, row in data.iterrows():
        raw_email = row[EMAIL_COLUMN]

        # Skip blank cells
        if pd.isna(raw_email) or str(raw_email).strip() == "":
            logger.debug("Row %d: empty email — skipping.", index)
            skipped += 1
            continue

        address = str(raw_email).strip()

        # Skip already-processed rows
        if row["Status"] in (STATUS_SENT, STATUS_FAILED):
            logger.info("Row %d (%s): already '%s' — skipping.", index, address, row["Status"])
            skipped += 1
            continue

        # Validate format
        if not is_valid_email(address):
            logger.warning("Row %d: '%s' is not a valid email address — skipping.", index, address)
            skipped += 1
            continue

        logger.info("Sending to: %s", address)
        if send_email(cfg, address):
            logger.info("✅  Sent successfully to %s", address)
            data.loc[index, "Status"] = STATUS_SENT
            sent += 1
        else:
            logger.error("❌  Failed to send to %s", address)
            data.loc[index, "Status"] = STATUS_FAILED
            failed += 1

        time.sleep(SEND_DELAY_SECONDS)

    try:
        save_recipients(data, cfg.excel_file)
    except OSError as exc:
        logger.error("Could not save Excel file: %s", exc)

    logger.info("Done — sent: %d | failed: %d | skipped: %d", sent, failed, skipped)


if __name__ == "__main__":
    main()
