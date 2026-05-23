"""Configuration for reminder mail system."""
import os
from pathlib import Path

# File paths
PROJECT_DIR = Path(__file__).parent
EXCEL_FILE = PROJECT_DIR / "PJ依頼事項管理表.xlsx"

# Sheet names (Japanese)
SHEET_SUMMARY = 0  # 台帳
SHEET_RESPONSE = 1  # 回答一覧
SHEET_TEAM = 2  # team定義

# Column indices (1-based in Excel, but using 0-based for openpyxl)
# 台帳 (Sheet 1)
SUMMARY_HEADER_ROW = 3
SUMMARY_REQUEST_ID_COL = 2  # B列
SUMMARY_STATUS_COL = 5  # E列
TARGET_STATUS = "依頼中"

# 回答一覧 (Sheet 2)
RESPONSE_HEADER_ROW = 2
RESPONSE_REQUEST_ID_COL = 2  # B列
RESPONSE_TEAM_NO_COL = 4  # D列
RESPONSE_CONTENT_COL = 5  # E列

# team定義 (Sheet 3)
TEAM_HEADER_ROW = 2
TEAM_NO_COL = 2  # B列
TEAM_CONTACT_START_COL = 3  # C列（管理者_1）
TEAM_CONTACT_END_COL = 7  # G列（管理者_5）

# Email settings
MAIL_SUBJECT = "【リマインド】回答依頼について"
MAIL_BODY_TEMPLATE = """お疲れ様です。

以下の依頼についてまだご回答をいただいていません。
お手数ですが、お早めにご対応いただきますようお願いいたします。

リマインド対象：{team_no}

詳細は以下のExcelファイルをご参照ください。
PJ依頼事項管理表.xlsx

よろしくお願いいたします。
"""

# Logging
LOG_DIR = PROJECT_DIR / "logs"
LOG_FILE = LOG_DIR / "reminder_mail.log"
LOG_LEVEL = "INFO"
