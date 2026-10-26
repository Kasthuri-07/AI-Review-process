"""
System Initialization Script
Initializes the SQLite database, task catalog, creates the 100-student dataset,
and writes out 'AI_Internship_Review_Management.xlsx'.
"""

import database
from config import EXCEL_PATH
from excel_manager import (
    TASK_CATALOG_DATA,
    generate_100_students_dataset,
    create_styled_workbook
)

def initialize_all():
    print("[INIT] Initializing database schemas...")
    database.init_db()

    print("[INIT] Populating Task Catalog...")
    with database.get_connection() as conn:
        cursor = conn.cursor()
        for t in TASK_CATALOG_DATA:
            cursor.execute("""
                INSERT INTO task_catalog (
                    task_id, track_role, task_title, task_description,
                    required_deliverables, evaluation_rubric, task_priority, estimated_days
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(task_id) DO UPDATE SET
                    task_title=excluded.task_title,
                    task_description=excluded.task_description,
                    required_deliverables=excluded.required_deliverables,
                    evaluation_rubric=excluded.evaluation_rubric,
                    task_priority=excluded.task_priority,
                    estimated_days=excluded.estimated_days
            """, (
                t["task_id"], t["track_role"], t["task_title"], t["task_description"],
                t["required_deliverables"], t["evaluation_rubric"], t["task_priority"], t["estimated_days"]
            ))
        conn.commit()

    print("[INIT] Generating 100 sample student records and relationships...")
    students, allocations, reviews, feedbacks, email_logs = generate_100_students_dataset()

    print(f"[INIT] Creating styled Excel workbook at: {EXCEL_PATH}...")
    create_styled_workbook(students, allocations, reviews, feedbacks, email_logs, EXCEL_PATH)
    print(f"[INIT] Excel workbook successfully created ({EXCEL_PATH.stat().st_size / 1024:.1f} KB).")

    print("[INIT] Loading records into SQLite...")
    with database.get_connection() as conn:
        cursor = conn.cursor()
        # Insert students
        for s in students:
            cursor.execute("""
                INSERT INTO students (
                    student_id, student_name, student_email, internship_start_date, internship_end_date,
                    internship_role, mentor_name, mentor_email, assigned_task_id, assigned_task_title,
                    task_description, task_allocation_date, task_deadline, task_status,
                    current_review_number, last_review_date, next_review_date, next_review_time,
                    gcr_link, task_submission_link, latest_ai_feedback, completion_percentage,
                    internship_status, email_delivery_status, last_email_sent, remarks
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(student_id) DO NOTHING
            """, (
                s["student_id"], s["student_name"], s["student_email"], s["internship_start_date"],
                s["internship_end_date"], s["internship_role"], s["mentor_name"], s["mentor_email"],
                s["assigned_task_id"], s["assigned_task_title"], s["task_description"],
                s["task_allocation_date"], s["task_deadline"], s["task_status"],
                s["current_review_number"], s["last_review_date"], s["next_review_date"],
                s["next_review_time"], s["gcr_link"], s["task_submission_link"],
                s["latest_ai_feedback"], s["completion_percentage"], s["internship_status"],
                s["email_delivery_status"], s["last_email_sent"], s["remarks"]
            ))

        # Insert allocations
        for al in allocations:
            cursor.execute("""
                INSERT INTO task_allocations (
                    task_id, task_title, task_description, student_id, student_name,
                    student_email, assigned_date, deadline, task_priority, required_deliverables,
                    submission_link, task_status, allocation_email_status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                al["task_id"], al["task_title"], al["task_description"], al["student_id"],
                al["student_name"], al["student_email"], al["assigned_date"], al["deadline"],
                al["task_priority"], al["required_deliverables"], al["submission_link"],
                al["task_status"], al["allocation_email_status"]
            ))

        # Insert reviews
        for r in reviews:
            cursor.execute("""
                INSERT INTO review_schedules (
                    review_id, student_id, student_name, student_email, task_id,
                    review_number, review_date, review_time, time_zone, meeting_gcr_link,
                    submission_status, review_status, reviewer, feedback_status, next_review_date
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(review_id) DO NOTHING
            """, (
                r["review_id"], r["student_id"], r["student_name"], r["student_email"],
                r["task_id"], r["review_number"], r["review_date"], r["review_time"],
                r["time_zone"], r["meeting_gcr_link"], r["submission_status"],
                r["review_status"], r["reviewer"], r["feedback_status"], r["next_review_date"]
            ))

        # Insert feedback
        for fb in feedbacks:
            cursor.execute("""
                INSERT INTO ai_review_feedback (
                    student_id, task_id, review_number, submission_summary, evaluation_criteria,
                    identified_mistakes, missing_requirements, strengths, improvement_suggestions,
                    ai_generated_feedback, score, reviewer_approval_status, next_actions
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                fb["student_id"], fb["task_id"], fb["review_number"], fb["submission_summary"],
                fb["evaluation_criteria"], fb["identified_mistakes"], fb["missing_requirements"],
                fb["strengths"], fb["improvement_suggestions"], fb["ai_generated_feedback"],
                fb["score"], fb["reviewer_approval_status"], fb["next_actions"]
            ))

        # Insert email logs
        for em in email_logs:
            cursor.execute("""
                INSERT INTO email_logs (
                    email_id, student_id, recipient_email, email_type, subject,
                    sent_date_time, delivery_status, error_message, retry_count, related_review_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(email_id) DO NOTHING
            """, (
                em["email_id"], em["student_id"], em["recipient_email"], em["email_type"],
                em["subject"], em["sent_date_time"], em["delivery_status"],
                em["error_message"], em["retry_count"], em["related_review_id"]
            ))

        conn.commit()

    database.log_audit("SYSTEM_INITIALIZE", "Database", "ALL", "System", "Initialized with 100 sample students and standard tasks.")
    print("[INIT] Database initialized and populated successfully.")

if __name__ == "__main__":
    initialize_all()
