"""
Review Scheduling and Automated Reminder Engine
Uses APScheduler with persistent job storage, time-zone awareness,
and periodic heartbeat scans for 24h, 1h, and overdue notifications.
"""

from datetime import datetime
import pytz
from typing import Dict, Any, Optional, Tuple
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.jobstores.sqlalchemy import SQLAlchemyJobStore
from apscheduler.executors.pool import ThreadPoolExecutor

from config import DB_URL, TIMEZONE, TaskStatus, ReviewStatus
import database
from email_service import EmailService
import excel_manager

class ReviewSchedulerService:
    def __init__(self, email_service: Optional[EmailService] = None):
        self.email_service = email_service or EmailService()
        self.tz = pytz.timezone(TIMEZONE)

        # Persistent job store to recover missed jobs after restarts
        jobstores = {
            'default': SQLAlchemyJobStore(url=DB_URL)
        }
        executors = {
            'default': ThreadPoolExecutor(5)
        }
        job_defaults = {
            'coalesce': True,
            'max_instances': 2,
            'misfire_grace_time': 3600  # 1 hour grace time to catch up missed jobs safely
        }

        self.scheduler = BackgroundScheduler(
            jobstores=jobstores,
            executors=executors,
            job_defaults=job_defaults,
            timezone=self.tz
        )
        self.is_running = False

    def start(self):
        """Starts the persistent background scheduler."""
        if not self.is_running:
            try:
                # Add a recurring scan heartbeat every 5 minutes
                if not self.scheduler.get_job("heartbeat_reminder_scanner"):
                    self.scheduler.add_job(
                        func=self.process_due_reminders,
                        trigger="interval",
                        minutes=5,
                        id="heartbeat_reminder_scanner",
                        replace_existing=True
                    )
                self.scheduler.start()
                self.is_running = True
                database.log_audit("SCHEDULER_STARTED", "System", "APScheduler", details="Scheduler started with 5-minute reminder heartbeat.")
            except Exception as e:
                print(f"[SCHEDULER] Warning starting scheduler: {e}")

    def shutdown(self):
        """Safely terminates the scheduler."""
        if self.is_running:
            self.scheduler.shutdown(wait=False)
            self.is_running = False

    def schedule_review(
        self,
        student_id: str,
        review_date_str: str,
        review_time_str: str = "14:00",
        meeting_link: Optional[str] = None,
        reviewer_name: Optional[str] = None,
        sync_to_excel: bool = True
    ) -> Tuple[bool, str, Dict[str, Any]]:
        """
        Schedules a formal review for a student, sends confirmation email,
        and saves schedule in SQLite and Excel.
        """
        student = database.get_student(student_id)
        if not student:
            return False, f"Student {student_id} not found.", {}

        if not student.get("assigned_task_id"):
            return False, f"Student {student_id} does not have an assigned task yet.", {}

        rev_num = int(student.get("current_review_number") or 0) + 1
        rev_id = f"REV-{student_id}-{rev_num}"
        reviewer = reviewer_name or student["mentor_name"]
        meet_url = meeting_link or f"https://meet.google.com/rev-{student_id.lower()}-{rev_num}"

        # 1. Insert/Update in review_schedules
        with database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO review_schedules (
                    review_id, student_id, student_name, student_email, task_id,
                    review_number, review_date, review_time, time_zone, meeting_gcr_link,
                    submission_status, review_status, reviewer, feedback_status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(review_id) DO UPDATE SET
                    review_date=excluded.review_date,
                    review_time=excluded.review_time,
                    meeting_gcr_link=excluded.meeting_gcr_link,
                    reviewer=excluded.reviewer,
                    review_status='Scheduled'
            """, (
                rev_id, student_id, student["student_name"], student["student_email"],
                student["assigned_task_id"], rev_num, review_date_str, review_time_str,
                TIMEZONE, meet_url, "Submitted" if student["task_status"] in (TaskStatus.SUBMITTED, TaskStatus.UNDER_REVIEW) else "Pending",
                ReviewStatus.SCHEDULED, reviewer, "Pending"
            ))

            # Update Student Master
            cursor.execute("""
                UPDATE students
                SET next_review_date = ?, next_review_time = ?, current_review_number = ?,
                    remarks = ?, updated_at = CURRENT_TIMESTAMP
                WHERE student_id = ?
            """, (review_date_str, review_time_str, rev_num, f"Review #{rev_num} scheduled for {review_date_str} {review_time_str}", student_id))
            conn.commit()

        database.log_audit("REVIEW_SCHEDULED", "Review", rev_id, details=f"Review #{rev_num} scheduled for {student_id} on {review_date_str} {review_time_str}")

        # 2. Dispatch Confirmation Email
        email_ctx = {
            "student_id": student_id,
            "student_name": student["student_name"],
            "review_id": rev_id,
            "review_number": rev_num,
            "task_id": student["assigned_task_id"],
            "task_title": student["assigned_task_title"],
            "review_date": review_date_str,
            "review_time": review_time_str,
            "time_zone": TIMEZONE,
            "reviewer": reviewer,
            "meeting_link": meet_url,
            "submission_status": "Submitted" if student["task_status"] in (TaskStatus.SUBMITTED, TaskStatus.UNDER_REVIEW) else "Pending Submission",
            "previous_feedback": student.get("latest_ai_feedback") or "No prior reviews recorded."
        }

        sent_ok, email_msg = self.email_service.send_email(
            student_id=student_id,
            recipient_email=student["student_email"],
            email_type="Review Scheduled",
            subject=f"[{student['assigned_task_id']}] Review #{rev_num} Confirmed for {review_date_str}",
            template_name="review_scheduled.html",
            context=email_ctx,
            related_review_id=rev_id
        )

        if sync_to_excel:
            excel_manager.export_sqlite_to_excel()

        return True, f"Review {rev_id} scheduled successfully. {email_msg}", {
            "review_id": rev_id,
            "student_id": student_id,
            "review_date": review_date_str,
            "review_time": review_time_str,
            "email_sent": sent_ok
        }

    def process_due_reminders(self) -> Dict[str, Any]:
        """
        Scans all scheduled reviews and pending tasks to trigger:
        - 24-hour reminders
        - 1-hour reminders
        - Overdue task notices
        """
        now = datetime.now()
        results = {
            "reminders_24h_sent": 0,
            "reminders_1h_sent": 0,
            "overdue_notices_sent": 0
        }

        # 1. Fetch data in a clean read transaction
        with database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT * FROM review_schedules
                WHERE review_status IN ('Scheduled', 'Reminder Sent')
            """)
            reviews = [dict(r) for r in cursor.fetchall()]

            today_str = now.strftime("%Y-%m-%d")
            cursor.execute("""
                SELECT * FROM students
                WHERE task_status IN ('Assigned', 'In Progress')
                AND task_deadline < ? AND task_deadline != ''
            """, (today_str,))
            overdue_students = [dict(s) for s in cursor.fetchall()]

        # 2. Process review reminders without holding outer database lock
        for r in reviews:
            rev_datetime_str = f"{r['review_date']} {r['review_time']}"
            try:
                rev_dt = datetime.strptime(rev_datetime_str, "%Y-%m-%d %H:%M")
            except ValueError:
                continue

            diff = rev_dt - now
            hours_remaining = diff.total_seconds() / 3600.0

            # 24-Hour Reminder Check (window: between 22 and 26 hours)
            if 22 <= hours_remaining <= 26 and not r.get("reminder_24h_sent"):
                ctx = {
                    "student_id": r["student_id"],
                    "student_name": r["student_name"],
                    "review_id": r["review_id"],
                    "review_number": r["review_number"],
                    "task_id": r["task_id"],
                    "review_date": r["review_date"],
                    "review_time": r["review_time"],
                    "time_zone": r["time_zone"],
                    "reviewer": r["reviewer"],
                    "submission_status": r["submission_status"],
                    "meeting_link": r["meeting_gcr_link"]
                }
                sent, _ = self.email_service.send_email(
                    student_id=r["student_id"],
                    recipient_email=r["student_email"],
                    email_type="Review Reminder 24H",
                    subject=f"[Reminder] Review #{r['review_number']} is Tomorrow ({r['review_date']})",
                    template_name="review_reminder_24h.html",
                    context=ctx,
                    related_review_id=r["review_id"]
                )
                if sent:
                    with database.get_connection() as conn:
                        conn.execute("UPDATE review_schedules SET reminder_24h_sent = 1, review_status = 'Reminder Sent' WHERE review_id = ?", (r["review_id"],))
                        conn.commit()
                    results["reminders_24h_sent"] += 1

            # 1-Hour Reminder Check (window: between 0.5 and 1.5 hours)
            elif 0.5 <= hours_remaining <= 1.5 and not r.get("reminder_1h_sent"):
                ctx = {
                    "student_id": r["student_id"],
                    "student_name": r["student_name"],
                    "review_id": r["review_id"],
                    "review_number": r["review_number"],
                    "task_id": r["task_id"],
                    "review_date": r["review_date"],
                    "review_time": r["review_time"],
                    "time_zone": r["time_zone"],
                    "reviewer": r["reviewer"],
                    "meeting_link": r["meeting_gcr_link"]
                }
                sent, _ = self.email_service.send_email(
                    student_id=r["student_id"],
                    recipient_email=r["student_email"],
                    email_type="Review Reminder 1H",
                    subject=f"[Urgent Reminder] Review #{r['review_number']} Starts in 1 Hour",
                    template_name="review_reminder_1h.html",
                    context=ctx,
                    related_review_id=r["review_id"]
                )
                if sent:
                    with database.get_connection() as conn:
                        conn.execute("UPDATE review_schedules SET reminder_1h_sent = 1 WHERE review_id = ?", (r["review_id"],))
                        conn.commit()
                    results["reminders_1h_sent"] += 1

        # 3. Process overdue tasks
        for s in overdue_students:
            ctx = {
                "student_id": s["student_id"],
                "student_name": s["student_name"],
                "task_id": s["assigned_task_id"],
                "task_title": s["assigned_task_title"],
                "deadline": s["task_deadline"],
                "mentor_name": s["mentor_name"],
                "mentor_email": s["mentor_email"],
                "submission_link": s["task_submission_link"]
            }
            sent, _ = self.email_service.send_email(
                student_id=s["student_id"],
                recipient_email=s["student_email"],
                email_type="Overdue Notification",
                subject=f"[OVERDUE] Immediate Action Required: {s['assigned_task_id']}",
                template_name="overdue_reminder.html",
                context=ctx
            )
            if sent:
                with database.get_connection() as conn:
                    conn.execute("""
                        UPDATE students
                        SET task_status = 'Overdue', remarks = 'Task deadline passed without submission. Overdue escalation sent.'
                        WHERE student_id = ?
                    """, (s["student_id"],))
                    conn.commit()
                results["overdue_notices_sent"] += 1

        return results
