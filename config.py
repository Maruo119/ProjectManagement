"""Configuration for reminder mail system."""
import os
from pathlib import Path

# File paths
PROJECT_DIR = Path(__file__).parent
EXCEL_FILE = PROJECT_DIR / "PJ依頼事項管理表.xlsx"

# Sheet names (Japanese)
SHEET_PROJECT = 0  # PJ情報
SHEET_SUMMARY = 1  # 台帳
SHEET_RESPONSE = 2  # 回答一覧
SHEET_TEAM = 3  # team定義

# Column indices (1-based in Excel, but using 0-based for openpyxl)
# PJ情報 (Sheet 0)
PROJECT_DATA_ROW = 2  # Data row (assuming row 2 contains project info)
PROJECT_NAME_COL = 1  # A列
PROJECT_PM_EMAIL_COL = 2  # B列
PROJECT_CC_EMAIL_COL = 3  # C列

# 台帳 (Sheet 1)
SUMMARY_HEADER_ROW = 3
SUMMARY_REQUEST_ID_COL = 2  # B列
SUMMARY_CONTENT_COL = 3  # C列
SUMMARY_DEADLINE_COL = 4  # D列
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
REMINDER_MAIL_SUBJECT = "【リマインド】回答依頼について"
REMINDER_MAIL_BODY_TEMPLATE = """お疲れ様です。
{project_name}について、以下の依頼についてまだご回答をいただいていません。
お手数ですが、お早めにご対応いただきますようお願いいたします。


依頼ID：{request_id}
依頼内容：{content}
期限：{deadline}


詳細は以下のExcelファイルをご参照ください。
D:\\ProjectManagement\\PJ依頼事項管理表.xlsx


よろしくお願いいたします。
"""

# PM Summary Email
PM_SUMMARY_SUBJECT = "【回答状況サマリー】{project_name}"
PM_SUMMARY_BODY_TEMPLATE = """お疲れ様です。

{project_name}の回答状況をお知らせします。


【未回答チーム一覧】
{team_summary}

ご不明な点やご質問がございましたら、お気軽にお問い合わせください。

よろしくお願いいたします。
"""

# Logging
LOG_DIR = PROJECT_DIR / "logs"
LOG_FILE = LOG_DIR / "reminder_mail.log"
LOG_LEVEL = "INFO"
