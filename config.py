"""
System Configuration and Constants
AI-Powered Internship Task Allocation, Review and Feedback Automation System
"""

import os
from pathlib import Path
from dotenv import load_dotenv

# Base directories
BASE_DIR = Path(__file__).resolve().parent

# Load environment variables
load_dotenv(dotenv_path=BASE_DIR / ".env", override=True)
DATA_DIR = BASE_DIR / "data"
BACKUP_DIR = BASE_DIR / "backups"
OUTBOX_DIR = BASE_DIR / "outbox"
TEMPLATES_DIR = BASE_DIR / "templates"
SUBMISSIONS_DIR = BASE_DIR / "sample_submissions"

# Ensure runtime directories exist
DATA_DIR.mkdir(parents=True, exist_ok=True)
BACKUP_DIR.mkdir(parents=True, exist_ok=True)
OUTBOX_DIR.mkdir(parents=True, exist_ok=True)
SUBMISSIONS_DIR.mkdir(parents=True, exist_ok=True)

# Excel File Paths
EXCEL_FILENAME = "AI_Internship_Review_Management.xlsx"
EXCEL_PATH = BASE_DIR / EXCEL_FILENAME

# SQLite Database
DB_PATH = DATA_DIR / "internship_system.db"
DB_URL = f"sqlite:///{DB_PATH}"

# Application Settings
APP_NAME = "AI Internship Review & Allocation System"
TIMEZONE = os.getenv("TIMEZONE", "Asia/Kolkata")
EMAIL_MODE = os.getenv("EMAIL_MODE", "MOCK")  # 'MOCK' or 'SMTP'

# SMTP Configuration
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", "kasthuricute273@gmail.com")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
SMTP_USE_TLS = os.getenv("SMTP_USE_TLS", "True").lower() in ("true", "1", "yes")
SENDER_NAME = os.getenv("SENDER_NAME", "Internship Operations Team")
SENDER_EMAIL = os.getenv("SENDER_EMAIL", "kasthuricute273@gmail.com")
RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL", "kasthurikumar1105@gmail.com")
DEFAULT_MEETING_LINK = os.getenv("DEFAULT_MEETING_LINK", "https://meet.google.com/zag-rcxj-kae")
DEFAULT_GCR_LINK = os.getenv("DEFAULT_GCR_LINK", "https://meet.google.com/zag-rcxj-kae")
EMAIL_MAX_RETRIES = int(os.getenv("EMAIL_MAX_RETRIES", "3"))

# AI / LLM Configuration
LLM_PROVIDER = os.getenv("LLM_PROVIDER", "MOCK")  # 'MOCK', 'OPENAI', 'GEMINI', 'ANTHROPIC'
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-1.5-flash")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
ANTHROPIC_MODEL = os.getenv("ANTHROPIC_MODEL", "claude-3-5-sonnet-20241022")

# Workflow & Review Policy
PASSING_SCORE_THRESHOLD = int(os.getenv("PASSING_SCORE_THRESHOLD", "70"))
DEFAULT_REVIEW_INTERVAL_DAYS = int(os.getenv("DEFAULT_REVIEW_INTERVAL_DAYS", "3"))
REMINDER_HOURS_BEFORE_REVIEW = [24, 1]  # Standard reminder triggers
OVERDUE_CHECK_HOURS = int(os.getenv("OVERDUE_CHECK_HOURS", "6"))

# Roles & Task Tracks
ROLES_TRACKS = [
    "AI/ML Engineer",
    "Full Stack Web Developer",
    "Data Analyst",
    "Cloud & DevOps Engineer",
    "Mobile App Developer",
    "Cybersecurity Specialist",
]

# Task & Internship Status Enumerations
class TaskStatus:
    NOT_STARTED = "Not Started"
    ASSIGNED = "Assigned"
    IN_PROGRESS = "In Progress"
    SUBMITTED = "Submitted"
    UNDER_REVIEW = "Under Review"
    CORRECTIONS_REQUIRED = "Corrections Required"
    RESUBMISSION_PENDING = "Resubmission Pending"
    APPROVED = "Approved"
    COMPLETED = "Completed"
    OVERDUE = "Overdue"
    ON_HOLD = "On Hold"

class InternshipStatus:
    NOT_STARTED = "Not Started"
    ACTIVE = "Active"
    IN_PROGRESS = "In Progress"
    NEARING_COMPLETION = "Nearing Completion"
    COMPLETED = "Completed"
    EXTENDED = "Extended"
    TERMINATED = "Terminated"

class ReviewStatus:
    SCHEDULED = "Scheduled"
    REMINDER_SENT = "Reminder Sent"
    SUBMISSION_RECEIVED = "Submission Received"
    AI_EVALUATED = "AI Evaluated"
    AWAITING_APPROVAL = "Awaiting Approval"
    APPROVED = "Approved"
    COMPLETED = "Completed"
    MISSED = "Missed"
    RESCHEDULED = "Rescheduled"

class EmailStatus:
    PENDING = "Pending"
    SENT = "Sent"
    MOCK_DELIVERED = "Mock Delivered"
    FAILED = "Failed"
    RETRY_SCHEDULED = "Retry Scheduled"
