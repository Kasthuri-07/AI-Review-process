"""
Email Notification and Reminder Engine
Supports Jinja2 HTML/Text rendering, configurable SMTP transport,
Mock Outbox mode for safe test execution, exponential backoff retries,
and duplicate email prevention.
"""

import smtplib
import ssl
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from typing import Dict, Any, Optional, Tuple
import jinja2

from config import (
    TEMPLATES_DIR, OUTBOX_DIR, EMAIL_MODE,
    SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, SMTP_USE_TLS,
    SENDER_NAME, SENDER_EMAIL, RECEIVER_EMAIL, DEFAULT_MEETING_LINK, DEFAULT_GCR_LINK, EMAIL_MAX_RETRIES,
    EmailStatus
)
import database

# Initialize Jinja2 environment
template_loader = jinja2.FileSystemLoader(searchpath=str(TEMPLATES_DIR))
jinja_env = jinja2.Environment(loader=template_loader, autoescape=True)

class EmailService:
    def __init__(self, mode: str = EMAIL_MODE):
        self.mode = mode.upper()
        self.max_retries = EMAIL_MAX_RETRIES
        self.default_receiver = RECEIVER_EMAIL

    def render_template(self, template_name: str, context: Dict[str, Any]) -> str:
        """Renders an HTML email from a template file."""
        template = jinja_env.get_template(template_name)
        return template.render(**context)

    def send_direct_notification(
        self,
        recipient_email: str = RECEIVER_EMAIL,
        subject: str = "AI Internship Notification",
        message_body: str = "Test email notification from Internship System.",
        student_id: str = "STU-001"
    ) -> Tuple[bool, str]:
        """Direct dispatch to a specified recipient (default RECEIVER_EMAIL)."""
        ctx = {
            "student_id": student_id,
            "student_name": "Kasthuri",
            "task_id": "TASK-AIML-101",
            "task_title": "Customer Churn Prediction Pipeline",
            "assigned_date": datetime.now().strftime("%Y-%m-%d"),
            "deadline": "2026-11-15",
            "mentor_name": SENDER_NAME,
            "mentor_email": SENDER_EMAIL,
            "internship_role": "AI/ML Engineer",
            "task_description": message_body,
            "required_deliverables": "1. Source Code Repository\n2. Serialized Model\n3. FastAPI Service",
            "task_priority": "High",
            "submission_link": DEFAULT_GCR_LINK,
            "meeting_link": DEFAULT_MEETING_LINK
        }
        return self.send_email(
            student_id=student_id,
            recipient_email=recipient_email,
            email_type="Direct Notification",
            subject=subject,
            template_name="task_allocation.html",
            context=ctx,
            force_resend=True
        )

    def is_duplicate_email(self, student_id: str, email_type: str, related_id: Optional[str] = None) -> bool:
        """
        Prevents duplicate emails for the same student and type/review
        unless specifically overridden.
        """
        with database.get_connection() as conn:
            cursor = conn.cursor()
            if related_id:
                cursor.execute("""
                    SELECT COUNT(*) FROM email_logs
                    WHERE student_id = ? AND email_type = ? AND related_review_id = ?
                    AND delivery_status IN ('Sent', 'Mock Delivered')
                """, (student_id, email_type, related_id))
            else:
                cursor.execute("""
                    SELECT COUNT(*) FROM email_logs
                    WHERE student_id = ? AND email_type = ?
                    AND delivery_status IN ('Sent', 'Mock Delivered')
                """, (student_id, email_type))
            count = cursor.fetchone()[0]
            return count > 0

    def send_email(
        self,
        student_id: str,
        recipient_email: str,
        email_type: str,
        subject: str,
        template_name: str,
        context: Dict[str, Any],
        related_review_id: Optional[str] = None,
        force_resend: bool = False
    ) -> Tuple[bool, str]:
        """
        Renders and dispatches an email with retry logic and audit logging.
        Returns (success: bool, message: str).
        """
        # 1. Duplicate check
        if not force_resend and self.is_duplicate_email(student_id, email_type, related_review_id):
            return False, f"Duplicate prevented: {email_type} has already been delivered to {student_id}."

        # Add subject to context for base template
        context["subject"] = subject
        html_content = self.render_template(template_name, context)

        # Generate unique Email ID
        email_id = f"EML-{int(datetime.now().timestamp() * 1000) % 10000000:07d}"
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # 2. Execution based on mode
        if self.mode == "MOCK":
            return self._send_mock(
                email_id, student_id, recipient_email, email_type,
                subject, html_content, now_str, related_review_id
            )
        else:
            return self._send_smtp_with_retries(
                email_id, student_id, recipient_email, email_type,
                subject, html_content, now_str, related_review_id
            )

    def _send_mock(
        self,
        email_id: str,
        student_id: str,
        recipient_email: str,
        email_type: str,
        subject: str,
        html_content: str,
        now_str: str,
        related_review_id: Optional[str]
    ) -> Tuple[bool, str]:
        """Saves email artifact locally in MOCK mode and logs success."""
        filename = f"{email_id}_{student_id}_{email_type.replace(' ', '_').lower()}.html"
        outbox_file = OUTBOX_DIR / filename
        with open(outbox_file, "w", encoding="utf-8") as f:
            f.write(html_content)

        # Log into SQLite
        with database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO email_logs (
                    email_id, student_id, recipient_email, email_type, subject,
                    sent_date_time, delivery_status, error_message, retry_count,
                    related_review_id, outbox_filepath
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                email_id, student_id, recipient_email, email_type, subject,
                now_str, EmailStatus.MOCK_DELIVERED, "", 0,
                related_review_id or "", str(outbox_file)
            ))
            cursor.execute("""
                UPDATE students
                SET email_delivery_status = 'Sent', last_email_sent = ?, updated_at = CURRENT_TIMESTAMP
                WHERE student_id = ?
            """, (now_str, student_id))
            conn.commit()

        database.log_audit(
            action="EMAIL_MOCK_DELIVERY",
            entity="EmailLog",
            entity_id=email_id,
            details=f"Mock captured to {filename} for {recipient_email}"
        )
        return True, f"Mock email generated and saved to {filename}"

    def _send_smtp_with_retries(
        self,
        email_id: str,
        student_id: str,
        recipient_email: str,
        email_type: str,
        subject: str,
        html_content: str,
        now_str: str,
        related_review_id: Optional[str]
    ) -> Tuple[bool, str]:
        """Sends email via SMTP with exponential backoff retries."""
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = f"{SENDER_NAME} <{SENDER_EMAIL}>"
        msg["To"] = recipient_email
        msg.attach(MIMEText(html_content, "html"))

        last_error = ""
        for attempt in range(1, self.max_retries + 1):
            try:
                context = ssl.create_default_context()
                server = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=25)
                server.ehlo()
                if SMTP_USE_TLS:
                    server.starttls(context=context)
                    server.ehlo()

                clean_password = SMTP_PASSWORD.replace(" ", "") if SMTP_PASSWORD else ""
                if SMTP_USER and clean_password:
                    try:
                        server.login(SMTP_USER, clean_password)
                    except smtplib.SMTPAuthenticationError:
                        server.login(SMTP_USER, SMTP_PASSWORD)

                server.sendmail(SENDER_EMAIL, [recipient_email], msg.as_string())
                server.quit()

                # Record success in DB
                with database.get_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute("""
                        INSERT INTO email_logs (
                            email_id, student_id, recipient_email, email_type, subject,
                            sent_date_time, delivery_status, error_message, retry_count, related_review_id
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        email_id, student_id, recipient_email, email_type, subject,
                        now_str, EmailStatus.SENT, "", attempt - 1, related_review_id or ""
                    ))
                    cursor.execute("""
                        UPDATE students
                        SET email_delivery_status = 'Sent', last_email_sent = ?, updated_at = CURRENT_TIMESTAMP
                        WHERE student_id = ?
                    """, (now_str, student_id))
                    conn.commit()

                database.log_audit("EMAIL_SENT", "EmailLog", email_id, details=f"Sent to {recipient_email} on attempt {attempt}")
                return True, f"Email delivered successfully to {recipient_email}"
            except Exception as e:
                last_error = str(e)
                time.sleep(2 ** attempt)  # Exponential backoff: 2s, 4s, 8s

        # If retries exhausted, record failure
        with database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO email_logs (
                    email_id, student_id, recipient_email, email_type, subject,
                    sent_date_time, delivery_status, error_message, retry_count, related_review_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                email_id, student_id, recipient_email, email_type, subject,
                now_str, EmailStatus.FAILED, last_error, self.max_retries, related_review_id or ""
            ))
            cursor.execute("""
                UPDATE students
                SET email_delivery_status = 'Failed', updated_at = CURRENT_TIMESTAMP
                WHERE student_id = ?
            """, (student_id,))
            conn.commit()

        database.log_audit("EMAIL_FAILED", "EmailLog", email_id, details=f"Failed to deliver to {recipient_email}: {last_error}")
        return False, f"Failed after {self.max_retries} attempts: {last_error}"
