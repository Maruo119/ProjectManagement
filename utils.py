"""Utility functions for reminder mail system."""
import logging
import re
from pathlib import Path
from datetime import datetime, timedelta
import holidays
from config import LOG_DIR, LOG_FILE, LOG_LEVEL


def setup_logger(name: str) -> logging.Logger:
    """Set up logger with file and console handlers."""
    LOG_DIR.mkdir(exist_ok=True)

    logger = logging.getLogger(name)
    logger.setLevel(getattr(logging, LOG_LEVEL))

    if not logger.handlers:
        formatter = logging.Formatter(
            "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        )

        file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
        file_handler.setLevel(getattr(logging, LOG_LEVEL))
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        console_handler = logging.StreamHandler()
        console_handler.setLevel(getattr(logging, LOG_LEVEL))
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

    return logger


def is_valid_email(email: str) -> bool:
    """Validate email address format."""
    if not email or not isinstance(email, str):
        return False
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(pattern, email) is not None


def remove_duplicates(items: list) -> list:
    """Remove duplicates while preserving order."""
    seen = set()
    result = []
    for item in items:
        if item and item not in seen:
            seen.add(item)
            result.append(item)
    return result


def clean_value(value) -> str:
    """Clean and normalize cell value."""
    if value is None:
        return ""
    return str(value).strip()


def is_business_day(date: datetime) -> bool:
    """Check if the given date is a business day (weekday and not a holiday)."""
    if date.weekday() >= 5:
        return False
    jp_holidays = holidays.Japan()
    return date.date() not in jp_holidays


def get_business_days_before(deadline_date: datetime, business_days: int) -> datetime:
    """Calculate the date that is N business days before the deadline."""
    current = deadline_date - timedelta(days=1)
    count = 0
    while count < business_days:
        if is_business_day(current):
            count += 1
        current -= timedelta(days=1)
    return current + timedelta(days=1)


def get_reminder_timing(deadline_date: datetime, today: datetime) -> str:
    """
    Determine reminder timing based on deadline date.

    Returns: '3days_before', '1day_before', 'on_deadline', or None
    """
    if not isinstance(deadline_date, datetime):
        return None
    if not isinstance(today, datetime):
        today = datetime.now()

    # Check if today is 3 business days before deadline
    three_days_before = get_business_days_before(deadline_date, 3)
    if today.date() == three_days_before.date():
        return "3days_before"

    # Check if today is 1 business day before deadline
    one_day_before = get_business_days_before(deadline_date, 1)
    if today.date() == one_day_before.date():
        return "1day_before"

    # Check if today is the deadline
    if is_business_day(deadline_date) and today.date() == deadline_date.date():
        return "on_deadline"

    return None
