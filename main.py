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

        # Step 2: Get pending request IDs (status = "依頼中")
        logger.info("Step 1: Extracting pending request IDs...")
        pending_request_ids = reader.get_pending_request_ids()
        if not pending_request_ids:
            logger.warning("No pending requests found")
            return

        # Step 3: Get teams needing response
        logger.info("Step 2: Extracting teams needing response...")
        teams_needing_response = reader.get_teams_needing_response(
            pending_request_ids
        )
        if not teams_needing_response:
            logger.warning("No teams need response")
            return

        # Step 4: Initialize Outlook sender
        logger.info("Step 3: Initializing Outlook connection...")
        sender = OutlookSender()

        # Step 5: Send reminder emails (one per team)
        logger.info("Step 4: Sending reminder emails...")
        sent_count = 0
        for team_no in teams_needing_response:
            logger.info(f"Processing team: {team_no}")
            contacts = reader.get_team_contacts(team_no)

            if contacts:
                if sender.send_reminder(team_no, contacts):
                    sent_count += 1
            else:
                logger.warning(f"No contacts found for team {team_no}")

        logger.info("="*50)
        logger.info(f"Process completed. Sent {sent_count} reminder emails.")
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
