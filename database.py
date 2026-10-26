"""
Database Layer (SQLite) for Internship System
Provides persistent source of truth, transactional integrity, audit logs, and status tracking.
"""

import sqlite3
from datetime import datetime
from typing import List, Dict, Any, Optional
from config import DB_PATH

def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_PATH), timeout=30.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA busy_timeout = 30000")
    return conn

def init_db():
    """Initializes all SQLite tables if they do not exist."""
    with get_connection() as conn:
        cursor = conn.cursor()

        # 1. Students Table
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS students (
            student_id TEXT PRIMARY KEY,
            student_name TEXT NOT NULL,
            student_email TEXT NOT NULL UNIQUE,
            internship_start_date TEXT,
            internship_end_date TEXT,
            internship_role TEXT NOT NULL,
            mentor_name TEXT NOT NULL,
            mentor_email TEXT NOT NULL,
            assigned_task_id TEXT,
            assigned_task_title TEXT,
            task_description TEXT,
            task_allocation_date TEXT,
            task_deadline TEXT,
            task_status TEXT DEFAULT 'Not Started',
            current_review_number INTEGER DEFAULT 0,
            last_review_date TEXT,
            next_review_date TEXT,
            next_review_time TEXT,
            gcr_link TEXT,
            task_submission_link TEXT,
            latest_ai_feedback TEXT,
            completion_percentage REAL DEFAULT 0.0,
            internship_status TEXT DEFAULT 'Active',
            email_delivery_status TEXT DEFAULT 'Pending',
            last_email_sent TEXT,
            remarks TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        """)

        # 2. Task Catalog Table
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS task_catalog (
            task_id TEXT PRIMARY KEY,
            track_role TEXT NOT NULL,
            task_title TEXT NOT NULL,
            task_description TEXT NOT NULL,
            required_deliverables TEXT NOT NULL,
            evaluation_rubric TEXT NOT NULL,
            task_priority TEXT DEFAULT 'Medium',
            estimated_days INTEGER DEFAULT 7
        );
        """)

        # 3. Task Allocations Table
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS task_allocations (
            allocation_id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id TEXT NOT NULL,
            task_title TEXT NOT NULL,
            task_description TEXT NOT NULL,
            student_id TEXT NOT NULL,
            student_name TEXT NOT NULL,
            student_email TEXT NOT NULL,
            assigned_date TEXT NOT NULL,
            deadline TEXT NOT NULL,
            task_priority TEXT DEFAULT 'Medium',
            required_deliverables TEXT NOT NULL,
            submission_link TEXT,
            task_status TEXT DEFAULT 'Assigned',
            allocation_email_status TEXT DEFAULT 'Pending',
            FOREIGN KEY (student_id) REFERENCES students(student_id)
        );
        """)

        # 4. Review Schedules Table
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS review_schedules (
            review_id TEXT PRIMARY KEY,
            student_id TEXT NOT NULL,
            student_name TEXT NOT NULL,
            student_email TEXT NOT NULL,
            task_id TEXT NOT NULL,
            review_number INTEGER NOT NULL,
            review_date TEXT NOT NULL,
            review_time TEXT NOT NULL,
            time_zone TEXT DEFAULT 'Asia/Kolkata',
            meeting_gcr_link TEXT,
            submission_status TEXT DEFAULT 'Pending Submission',
            review_status TEXT DEFAULT 'Scheduled',
            reviewer TEXT NOT NULL,
            feedback_status TEXT DEFAULT 'Pending',
            next_review_date TEXT,
            reminder_24h_sent INTEGER DEFAULT 0,
            reminder_1h_sent INTEGER DEFAULT 0,
            overdue_notice_sent INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (student_id) REFERENCES students(student_id)
        );
        """)

        # 5. AI Review Feedback Table
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS ai_review_feedback (
            feedback_id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id TEXT NOT NULL,
            task_id TEXT NOT NULL,
            review_number INTEGER NOT NULL,
            submission_summary TEXT,
            evaluation_criteria TEXT,
            identified_mistakes TEXT,
            missing_requirements TEXT,
            strengths TEXT,
            improvement_suggestions TEXT,
            ai_generated_feedback TEXT,
            score INTEGER DEFAULT 0,
            reviewer_approval_status TEXT DEFAULT 'Pending Reviewer Approval',
            reviewer_notes TEXT,
            approved_by TEXT,
            next_actions TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            approved_at TEXT,
            FOREIGN KEY (student_id) REFERENCES students(student_id)
        );
        """)

        # 6. Email Logs Table
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS email_logs (
            email_id TEXT PRIMARY KEY,
            student_id TEXT NOT NULL,
            recipient_email TEXT NOT NULL,
            email_type TEXT NOT NULL,
            subject TEXT NOT NULL,
            sent_date_time TEXT NOT NULL,
            delivery_status TEXT NOT NULL,
            error_message TEXT,
            retry_count INTEGER DEFAULT 0,
            related_review_id TEXT,
            body_preview TEXT,
            outbox_filepath TEXT
        );
        """)

        # 7. Audit Logs Table
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS audit_logs (
            audit_id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT DEFAULT CURRENT_TIMESTAMP,
            action TEXT NOT NULL,
            entity TEXT NOT NULL,
            entity_id TEXT,
            performed_by TEXT DEFAULT 'System Automation',
            details TEXT
        );
        """)

        conn.commit()

def log_audit(action: str, entity: str, entity_id: str, performed_by: str = "System Automation", details: str = ""):
    """Records an immutable audit event."""
    with get_connection() as conn:
        conn.execute(
            "INSERT INTO audit_logs (action, entity, entity_id, performed_by, details) VALUES (?, ?, ?, ?, ?)",
            (action, entity, entity_id, performed_by, details)
        )
        conn.commit()

# --- Query & Persistence Helpers ---

def get_all_students() -> List[Dict[str, Any]]:
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM students ORDER BY student_id ASC")
        return [dict(row) for row in cursor.fetchall()]

def get_student(student_id: str) -> Optional[Dict[str, Any]]:
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM students WHERE student_id = ?", (student_id,))
        row = cursor.fetchone()
        return dict(row) if row else None

def get_all_tasks() -> List[Dict[str, Any]]:
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM task_catalog ORDER BY task_id ASC")
        return [dict(row) for row in cursor.fetchall()]

def get_all_allocations() -> List[Dict[str, Any]]:
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM task_allocations ORDER BY allocation_id DESC")
        return [dict(row) for row in cursor.fetchall()]

def get_all_reviews() -> List[Dict[str, Any]]:
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM review_schedules ORDER BY review_date ASC, review_time ASC")
        return [dict(row) for row in cursor.fetchall()]

def get_all_feedbacks() -> List[Dict[str, Any]]:
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM ai_review_feedback ORDER BY feedback_id DESC")
        return [dict(row) for row in cursor.fetchall()]

def get_all_email_logs() -> List[Dict[str, Any]]:
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM email_logs ORDER BY sent_date_time DESC")
        return [dict(row) for row in cursor.fetchall()]

def get_dashboard_metrics() -> Dict[str, Any]:
    with get_connection() as conn:
        cursor = conn.cursor()
        
        cursor.execute("SELECT COUNT(*) FROM students")
        total_students = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM students WHERE internship_status IN ('Active', 'In Progress')")
        active_interns = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM students WHERE assigned_task_id IS NOT NULL AND assigned_task_id != ''")
        tasks_assigned = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM students WHERE task_status IN ('Approved', 'Completed')")
        tasks_completed = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM students WHERE task_status IN ('Assigned', 'In Progress', 'Resubmission Pending')")
        pending_tasks = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM students WHERE task_status = 'Overdue'")
        overdue_tasks = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM review_schedules WHERE review_status IN ('Scheduled', 'Reminder Sent')")
        reviews_scheduled = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM review_schedules WHERE review_status IN ('Completed', 'Approved')")
        reviews_completed = cursor.fetchone()[0]

        today_str = datetime.now().strftime("%Y-%m-%d")
        cursor.execute("SELECT COUNT(*) FROM review_schedules WHERE review_date = ? AND review_status != 'Completed'", (today_str,))
        reviews_due_today = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM email_logs WHERE delivery_status IN ('Sent', 'Mock Delivered')")
        emails_sent = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM email_logs WHERE delivery_status = 'Failed'")
        emails_failed = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM students WHERE task_status IN ('Submitted', 'Under Review')")
        students_awaiting_feedback = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM students WHERE internship_status = 'Completed'")
        students_completed_internship = cursor.fetchone()[0]

        return {
            "total_students": total_students,
            "active_interns": active_interns,
            "tasks_assigned": tasks_assigned,
            "tasks_completed": tasks_completed,
            "pending_tasks": pending_tasks,
            "overdue_tasks": overdue_tasks,
            "reviews_scheduled": reviews_scheduled,
            "reviews_completed": reviews_completed,
            "reviews_due_today": reviews_due_today,
            "emails_sent": emails_sent,
            "emails_failed": emails_failed,
            "students_awaiting_feedback": students_awaiting_feedback,
            "students_completed_internship": students_completed_internship
        }
