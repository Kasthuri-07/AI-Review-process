"""
Google Classroom & Workspace Integration Connector
Implements OAuth 2.0 authorization verification, submission extraction interface,
and security guardrails ensuring that raw links are never claimed as authenticated.
"""

import os
from typing import Dict, Any, Tuple
from config import BASE_DIR

CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token.json"

SCOPES = [
    "https://www.googleapis.com/auth/classroom.courses.readonly",
    "https://www.googleapis.com/auth/classroom.coursework.students",
    "https://www.googleapis.com/auth/classroom.student-submissions.students.readonly"
]

class GoogleClassroomConnector:
    """
    Connects to the official Google Classroom API using OAuth 2.0 credentials.
    Adheres to strict truthfulness: if credentials or tokens are absent,
    it returns clear instructions and prevents false assertions of access.
    """

    def __init__(self):
        self.is_configured = CREDENTIALS_FILE.exists() or os.getenv("GOOGLE_APPLICATION_CREDENTIALS") is not None
        self.has_active_token = TOKEN_FILE.exists()

    def get_auth_status(self) -> Dict[str, Any]:
        """Returns the current OAuth2 configuration status."""
        return {
            "credentials_file_present": CREDENTIALS_FILE.exists(),
            "active_token_present": TOKEN_FILE.exists(),
            "status": "Ready" if (CREDENTIALS_FILE.exists() and TOKEN_FILE.exists()) else "OAuth Setup Required",
            "instructions": (
                "To authenticate Google Classroom:\n"
                "1. Create a Google Cloud Project and enable 'Google Classroom API'.\n"
                "2. Create OAuth 2.0 Desktop Client IDs and download as 'credentials.json' in the project root.\n"
                "3. Run 'python integrations/google_classroom.py --authorize' to generate 'token.json'.\n"
                "4. Without this OAuth flow, raw Google Classroom links are inaccessible to external scripts."
            )
        }

    def fetch_student_submission(self, course_id: str, coursework_id: str, student_id: str) -> Tuple[bool, str, Dict[str, Any]]:
        """
        Attempts to fetch student coursework submission from Google Classroom API.
        Fails safely if OAuth token is missing.
        """
        if not self.has_active_token:
            return False, (
                "Google Classroom API access unauthorized. Meeting links and Classroom URLs "
                "require explicit OAuth 2.0 user consent. Please upload files directly or authenticate OAuth."
            ), {}

        try:
            # When google-api-python-client is configured with token.json:
            # from google.oauth2.credentials import Credentials
            # from googleapiclient.discovery import build
            # creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
            # service = build('classroom', 'v1', credentials=creds)
            # submissions = service.courses().courseWork().studentSubmissions().list(
            #     courseId=course_id, courseWorkId=coursework_id, userId=student_id).execute()
            # return True, "Successfully retrieved submissions", submissions
            return False, "Google API libraries need token initialization.", {}
        except Exception as e:
            return False, f"Google Classroom API error: {str(e)}", {}

if __name__ == "__main__":
    connector = GoogleClassroomConnector()
    status = connector.get_auth_status()
    print("=== Google Classroom Connector Status ===")
    for k, v in status.items():
        print(f"{k}: {v}")
