"""Excel file reading and data extraction module."""
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from config import (
    EXCEL_FILE,
    SHEET_SUMMARY,
    SHEET_RESPONSE,
    SHEET_TEAM,
    SUMMARY_HEADER_ROW,
    SUMMARY_REQUEST_ID_COL,
    SUMMARY_STATUS_COL,
    RESPONSE_HEADER_ROW,
    RESPONSE_REQUEST_ID_COL,
    RESPONSE_TEAM_NO_COL,
    RESPONSE_CONTENT_COL,
    TEAM_HEADER_ROW,
    TEAM_NO_COL,
    TEAM_CONTACT_START_COL,
    TEAM_CONTACT_END_COL,
    TARGET_STATUS,
)
from utils import setup_logger, clean_value

logger = setup_logger(__name__)


class ExcelReader:
    """Read and extract data from Excel file."""

    def __init__(self):
        """Initialize ExcelReader."""
        try:
            self.workbook = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
            self.sheet_summary = self.workbook.worksheets[SHEET_SUMMARY]
            self.sheet_response = self.workbook.worksheets[SHEET_RESPONSE]
            self.sheet_team = self.workbook.worksheets[SHEET_TEAM]
            logger.info(f"Loaded Excel file: {EXCEL_FILE}")
        except Exception as e:
            logger.error(f"Failed to load Excel file: {e}")
            raise

    def get_pending_request_ids(self) -> list[str]:
        """Extract request IDs with status='依頼中' from summary sheet."""
        request_ids = []
        try:
            for row in range(SUMMARY_HEADER_ROW + 1, self.sheet_summary.max_row + 1):
                status = clean_value(
                    self.sheet_summary.cell(row, SUMMARY_STATUS_COL).value
                )
                if status == TARGET_STATUS:
                    request_id = clean_value(
                        self.sheet_summary.cell(row, SUMMARY_REQUEST_ID_COL).value
                    )
                    if request_id:
                        request_ids.append(request_id)
            logger.info(f"Found {len(request_ids)} pending request IDs: {request_ids}")
            return request_ids
        except Exception as e:
            logger.error(f"Error extracting pending request IDs: {e}")
            raise

    def get_teams_needing_response(self, request_ids: list[str]) -> list[str]:
        """
        Extract team numbers from response sheet where:
        - request ID is in the given list
        - response content (column E) is blank
        """
        team_nos = []
        try:
            for row in range(RESPONSE_HEADER_ROW + 1, self.sheet_response.max_row + 1):
                req_id = clean_value(
                    self.sheet_response.cell(row, RESPONSE_REQUEST_ID_COL).value
                )
                response_content = clean_value(
                    self.sheet_response.cell(row, RESPONSE_CONTENT_COL).value
                )

                if req_id in request_ids and response_content == "":
                    team_no = clean_value(
                        self.sheet_response.cell(row, RESPONSE_TEAM_NO_COL).value
                    )
                    if team_no:
                        team_nos.append(team_no)
                        logger.info(
                            f"Found team needing response: {team_no} (request ID: {req_id})"
                        )

            # Remove duplicates while preserving order
            unique_team_nos = list(dict.fromkeys(team_nos))
            logger.info(
                f"Total unique teams needing response: {len(unique_team_nos)}"
            )
            return unique_team_nos
        except Exception as e:
            logger.error(f"Error extracting teams needing response: {e}")
            raise

    def get_team_contacts(self, team_no: str) -> list[str]:
        """Extract email addresses for a given team number."""
        contacts = []
        try:
            for row in range(TEAM_HEADER_ROW + 1, self.sheet_team.max_row + 1):
                current_team_no = clean_value(
                    self.sheet_team.cell(row, TEAM_NO_COL).value
                )
                if current_team_no == team_no:
                    # Extract all contact columns (C to G)
                    for col in range(TEAM_CONTACT_START_COL, TEAM_CONTACT_END_COL + 1):
                        email = clean_value(self.sheet_team.cell(row, col).value)
                        if email and email != "":
                            contacts.append(email)
                    logger.info(
                        f"Found {len(contacts)} contacts for team {team_no}: {contacts}"
                    )
                    break
            else:
                logger.warning(f"Team {team_no} not found in team definition sheet")

            return contacts
        except Exception as e:
            logger.error(f"Error extracting team contacts for {team_no}: {e}")
            raise

    def close(self):
        """Close workbook."""
        self.workbook.close()
        logger.info("Excel file closed")
