"""Outlook email sending module."""
from win32com.client import Dispatch
from config import (
    PM_SUMMARY_SUBJECT,
    PM_SUMMARY_BODY_TEMPLATE,
)
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

    def send_reminder(
        self,
        team_no: str,
        recipients: list[str],
        project_name: str,
        cc_email: str,
        requests: list[dict],
    ) -> int:
        """
        Send reminder emails to team (one email per request).

        Args:
            team_no: Team number
            recipients: List of email addresses
            project_name: Project name for email body
            cc_email: Email address for CC field
            requests: List of request dicts with 'id', 'content', 'deadline'

        Returns:
            int: Number of emails sent successfully
        """
        try:
            # Validate recipients
            valid_recipients = [
                email for email in recipients if is_valid_email(email)
            ]

            if not valid_recipients:
                logger.warning(f"No valid email addresses found for team {team_no}")
                return 0

            recipients_str = ";".join(valid_recipients)
            sent_count = 0

            # Send one email per request
            for req in requests:
                try:
                    # Create mail item
                    mail = self.outlook.CreateItem(0)  # 0 = olMailItem

                    # Set recipients
                    mail.To = recipients_str
                    if cc_email and is_valid_email(cc_email):
                        mail.Cc = cc_email

                    # Set subject with project name and request id
                    mail.Subject = f"{project_name}　依頼#{req['id']}　リマインド"

                    # Build body with request details
                    body_lines = [f"お疲れ様です。"]
                    body_lines.append(f"{project_name}について、以下の依頼についてまだご回答をいただいていません。")
                    body_lines.append("お手数ですが、お早めにご対応いただきますようお願いいたします。")
                    body_lines.append("")
                    body_lines.append(f"依頼ID：{req['id']}")
                    body_lines.append(f"依頼内容：{req['content']}")
                    body_lines.append(f"期限：{req['deadline']}")
                    body_lines.append("")
                    body_lines.append("詳細は以下のExcelファイルをご参照ください。")
                    body_lines.append("D:\\ProjectManagement\\PJ依頼事項管理表.xlsx")
                    body_lines.append("")
                    body_lines.append("よろしくお願いいたします。")

                    mail.Body = "\n".join(body_lines)

                    # Send email
                    mail.Send()
                    sent_count += 1

                    logger.info(
                        f"Sent reminder email for team {team_no}, request {req['id']}"
                    )

                except Exception as e:
                    logger.error(
                        f"Failed to send email for team {team_no}, request {req['id']}: {e}"
                    )

            if sent_count > 0:
                logger.info(
                    f"Successfully sent {sent_count} reminder email(s) for team {team_no} to: {recipients_str}"
                )
            return sent_count

        except Exception as e:
            logger.error(f"Failed to process reminder emails for team {team_no}: {e}")
            return 0

    def send_pm_summary(
        self, pm_email: str, project_name: str, team_summary: list[dict]
    ) -> bool:
        """
        Send summary email to PM with teams needing response.

        Args:
            pm_email: PM's email address
            project_name: Project name
            team_summary: List of dicts with 'team_no' and 'requests'

        Returns:
            bool: True if email sent successfully, False otherwise
        """
        try:
            if not pm_email or not is_valid_email(pm_email):
                logger.warning(f"Invalid PM email address: {pm_email}")
                return False

            # Build team summary section
            team_lines = []
            for team_info in team_summary:
                team_no = team_info["team_no"]
                requests = team_info["requests"]
                team_lines.append(f"【{team_no}】")
                for req in requests:
                    team_lines.append(f"  - 依頼ID: {req['id']} / {req['content']}")
                team_lines.append("")

            team_summary_str = "\n".join(team_lines)

            # Create mail item
            mail = self.outlook.CreateItem(0)  # 0 = olMailItem

            # Set recipient
            mail.To = pm_email

            # Set subject and body
            mail.Subject = PM_SUMMARY_SUBJECT.format(project_name=project_name)
            mail.Body = PM_SUMMARY_BODY_TEMPLATE.format(
                project_name=project_name, team_summary=team_summary_str
            )

            # Send email
            mail.Send()

            logger.info(f"Successfully sent PM summary email to: {pm_email}")
            return True

        except Exception as e:
            logger.error(f"Failed to send PM summary email to {pm_email}: {e}")
            return False

    def close(self):
        """Close Outlook connection."""
        logger.info("Outlook connection closed")
