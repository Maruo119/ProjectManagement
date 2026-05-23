"""Main script for sending reminder emails."""
from excel_reader import ExcelReader
from outlook_sender import OutlookSender
from utils import setup_logger

logger = setup_logger(__name__)


def main():
    """Main process flow."""
    reader = None
    sender = None

    try:
        # Step 1: Read Excel file
        logger.info("="*50)
        logger.info("Starting reminder mail process")
        logger.info("="*50)

        reader = ExcelReader()

        # Step 2: Get project information
        logger.info("Step 1: Extracting project information...")
        project_info = reader.get_project_info()

        # Step 3: Get pending requests (status = "依頼中")
        logger.info("Step 2: Extracting pending requests...")
        pending_requests = reader.get_pending_requests()
        if not pending_requests:
            logger.warning("No pending requests found")
            return

        # Step 4: Get teams needing response
        logger.info("Step 3: Extracting teams needing response...")
        teams_needing_response = reader.get_teams_needing_response(pending_requests)
        if not teams_needing_response:
            logger.warning("No teams need response")
            return

        # Step 5: Initialize Outlook sender
        logger.info("Step 4: Initializing Outlook connection...")
        sender = OutlookSender()

        # Step 6: Send reminder emails (one per team)
        logger.info("Step 5: Sending reminder emails...")
        sent_count = 0
        for team_info in teams_needing_response:
            team_no = team_info["team_no"]
            requests = team_info["requests"]
            logger.info(f"Processing team: {team_no}")
            contacts = reader.get_team_contacts(team_no)

            if contacts:
                if sender.send_reminder(
                    team_no,
                    contacts,
                    project_info["project_name"],
                    project_info["cc_email"],
                    requests,
                ):
                    sent_count += 1
            else:
                logger.warning(f"No contacts found for team {team_no}")

        # Step 7: Send PM summary email
        logger.info("Step 6: Sending PM summary email...")
        if sender.send_pm_summary(
            project_info["pm_email"],
            project_info["project_name"],
            teams_needing_response,
        ):
            logger.info("PM summary email sent successfully")
        else:
            logger.warning("Failed to send PM summary email")

        logger.info("="*50)
        logger.info(f"Process completed. Sent {sent_count} reminder emails + 1 PM summary.")
        logger.info("="*50)

    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)
        raise

    finally:
        if reader:
            reader.close()
        if sender:
            sender.close()


if __name__ == "__main__":
    main()
