# Project Management Reminder Mail System

## Overview

This project automates the process of sending reminder emails for pending project requests. It reads project request information from an Excel file, identifies requests nearing their deadlines, and sends timely reminders to relevant teams via Outlook. A summary email is also sent to the project manager.

## Key Features

- **Automated Reminder Emails**: Sends reminders at configurable deadline intervals (3 days before, 1 day before, on deadline)
- **Dynamic Subject Lines**: Customizable email subjects based on deadline timing
- **Team-based Distribution**: Routes reminders to appropriate team members based on Excel configuration
- **PM Summary Emails**: Provides project managers with a consolidated status summary
- **Logging**: Comprehensive logging for monitoring and troubleshooting

## Requirements

- Python 3.8+
- Windows environment with Outlook installed
- Excel file with proper structure (see [Excel File Structure](#excel-file-structure))

## Setup

### 1. Install Dependencies

```powershell
pip install -r requirements.txt
```

### 2. Configure Excel File

Ensure your Excel file (`PJ依頼事項管理表.xlsx`) contains the following sheets in order:
- Sheet 1: 台帳 (Summary)
- Sheet 2: 回答一覧 (Response List)
- Sheet 3: team定義 (Team Definition)
- Sheet 4: PJ情報 (Project Information)

### 3. Environment Configuration

Create a `.env` file in the project root with necessary configurations (if required for your setup).

## Excel File Structure

### Sheet 1: 台帳 (Summary)
| Column | Field | Required |
|--------|-------|----------|
| B | Request ID | Yes |
| C | Content | Yes |
| D | Deadline | Yes |
| E | Status | Yes (target: "依頼中") |

### Sheet 2: 回答一覧 (Response List)
| Column | Field | Required |
|--------|-------|----------|
| B | Request ID | Yes |
| D | Team No | Yes |
| E | Content | Yes |

### Sheet 3: team定義 (Team Definition)
| Column | Field | Required |
|--------|-------|----------|
| B | Team No | Yes |
| C-G | Contact Persons (Manager_1 to Manager_5) | At least one |

### Sheet 4: PJ情報 (Project Information)
| Row | Field | Column |
|-----|-------|--------|
| 2 | Project Name | C |
| 3 | PM Email Address | C |
| 4 | CC Email Address | C |

## Usage

### Running the Script

```powershell
python main.py
```

The script will:
1. Read the Excel file
2. Extract project information
3. Identify pending requests by deadline timing
4. Send reminder emails to relevant teams
5. Send a summary email to the project manager
6. Log all actions

### Output

- **Console Output**: Progress and status messages
- **Log File**: Detailed logs saved to `logs/reminder_mail.log`

## Project Structure

```
ProjectManagement/
├── main.py                    # Main execution script
├── config.py                  # Configuration and constants
├── excel_reader.py            # Excel file processing logic
├── outlook_sender.py          # Outlook email sending logic
├── utils.py                   # Utility functions (logging, etc.)
├── requirements.txt           # Python dependencies
├── PJ依頼事項管理表.xlsx      # Excel data file
├── logs/                      # Log files directory
└── README.md                  # This file
```

## Module Descriptions

### main.py
Main orchestration script that coordinates the entire reminder email workflow.

### config.py
Contains all configuration constants including:
- File paths and sheet indices
- Column mappings for Excel sheets
- Email templates and subjects
- Logging configuration

### excel_reader.py
Handles all Excel file operations:
- Reading project information
- Extracting pending requests
- Filtering requests by deadline timing
- Retrieving team contact information

### outlook_sender.py
Manages Outlook email operations:
- Sending reminder emails to teams
- Sending summary emails to project managers
- Email formatting and delivery

### utils.py
Provides utility functions:
- Logger setup and configuration
- Common helper functions

## Troubleshooting

### No requests matching reminder timing found
- Check that the Excel file has entries with status "依頼中"
- Verify deadline dates are properly formatted
- Confirm at least one request has a deadline matching the current date or configured timing

### No teams need response
- Verify team information is correctly linked in the 回答一覧 sheet
- Check that team definitions exist in the team定義 sheet

### Outlook connection errors
- Ensure Outlook is installed and properly configured on your system
- Verify you are running the script with the correct user account
- Check Windows user permissions for Outlook access

### Missing contact information
- Verify team contact email addresses are filled in the team定義 sheet
- Ensure at least one contact person (Manager_1 through Manager_5) is defined per team

## Logging

Logs are saved to `logs/reminder_mail.log` by default. Log level can be configured in `config.py`.

## License

Internal Project

## Author

Project Management Team
