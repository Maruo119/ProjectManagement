"""Outlook email sending module."""
from win32com.client import Dispatch
from config import MAIL_SUBJECT, MAIL_BODY_TEMPLATE
from utils import setup_logger, is_valid_email

logger = setup_logger(__name__)


class OutlookSender:
    """Send reminder emails via Outlook."""

    def __init__(self):
        """Initialize Outlook COM connection."""
        try:
            self.outlook = Dispatch("Outlook.Application")
            logger.info("Connected to Outlook")
        except Exception as e:
            logger.error(f"Failed to connect to Outlook: {e}")
            raise

    def send_reminder(self, team_no: str, recipients: list[str]) -> bool:
        """
        Send reminder email to team.

        Args:
            team_no: Team number (used in email body)
            recipients: List of email addresses

        Returns:
            bool: True if email sent successfully, False otherwise
        """
        try:
            # Validate recipients
            valid_recipients = [
                email for email in recipients if is_valid_email(email)
            ]

            if not valid_recipients:
                logger.warning(f"No valid email addresses found for team {team_no}")
                return False

            # Create mail item
            mail = self.outlook.CreateItem(0)  # 0 = olMailItem

            # Set recipients (TO field)
            recipients_str = ";".join(valid_recipients)
            mail.To = recipients_str

            # Set subject and body
            mail.Subject = MAIL_SUBJECT
            mail.Body = MAIL_BODY_TEMPLATE.format(team_no=team_no)

            # Send email
            mail.Send()

            logger.info(
                f"Successfully sent reminder email for team {team_no} to: {recipients_str}"
            )
            return True

        except Exception as e:
            logger.error(f"Failed to send email for team {team_no}: {e}")
            return False

    def close(self):
        """Close Outlook connection."""
        logger.info("Outlook connection closed")
