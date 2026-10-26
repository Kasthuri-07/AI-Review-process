"""
Workflow Engine & Lifecycle Orchestrator
Coordinates the multi-review iterative loop:
Review -> AI Evaluation -> Reviewer Approval -> (Corrections Required -> Next Review) OR (Task Approved)
Separates task completion from internship completion and handles missed reviews.
"""

from datetime import datetime, timedelta
from typing import Dict, Any, Optional, Tuple

from config import (
    PASSING_SCORE_THRESHOLD, DEFAULT_REVIEW_INTERVAL_DAYS, TIMEZONE,
    TaskStatus, InternshipStatus, ReviewStatus
)
import database
from email_service import EmailService
from scheduler_service import ReviewSchedulerService
import excel_manager

class WorkflowEngine:
    def __init__(
        self,
        email_service: Optional[EmailService] = None,
        scheduler_service: Optional[ReviewSchedulerService] = None
    ):
        self.email_service = email_service or EmailService()
        self.scheduler_service = scheduler_service or ReviewSchedulerService(self.email_service)

    def approve_feedback_and_advance(
        self,
        feedback_id: int,
        reviewer_name: str,
        reviewer_notes: str = "",
        score_override: Optional[int] = None,
        sync_to_excel: bool = True
    ) -> Tuple[bool, str, Dict[str, Any]]:
        """
        Human reviewer approves the AI evaluation feedback.
        Advances student along the state machine:
        - If score >= threshold -> Task Approved / Completed -> Check Internship Completion
        - If score < threshold -> Corrections Required -> Schedule Next Review -> Email student
        """
        # 1. Fetch feedback entry
        with database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM ai_review_feedback WHERE feedback_id = ?", (feedback_id,))
            row = cursor.fetchone()
            if not row:
                return False, f"Feedback ID {feedback_id} not found.", {}
            fb = dict(row)

        student_id = fb["student_id"]
        task_id = fb["task_id"]
        student = database.get_student(student_id)
        if not student:
            return False, f"Student {student_id} not found.", {}

        final_score = int(score_override if score_override is not None else fb["score"])
        is_passed = (final_score >= PASSING_SCORE_THRESHOLD)
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # 2. Update Feedback Record in DB
        with database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE ai_review_feedback
                SET score = ?,
                    reviewer_approval_status = 'Approved',
                    reviewer_notes = ?,
                    approved_by = ?,
                    approved_at = ?
                WHERE feedback_id = ?
            """, (final_score, reviewer_notes, reviewer_name, now_str, feedback_id))
            conn.commit()

        database.log_audit(
            action="FEEDBACK_APPROVED",
            entity="Feedback",
            entity_id=str(feedback_id),
            performed_by=reviewer_name,
            details=f"Score: {final_score}/100. Status: {'PASSED' if is_passed else 'CORRECTIONS_REQUIRED'}"
        )

        result_payload = {"feedback_id": feedback_id, "student_id": student_id, "score": final_score}

        # 3. Branching State Machine
        if is_passed:
            # Task Approved / Completed
            new_completion_pct = min(100.0, float(student.get("completion_percentage") or 0.0) + 50.0)
            task_status = TaskStatus.APPROVED

            # Check if student meets criteria for full internship completion:
            # (e.g. completion_percentage reaches 100% and at least 2 reviews completed)
            internship_status = student.get("internship_status", "Active")
            rev_num = int(student.get("current_review_number") or 1)
            if new_completion_pct >= 100.0 and rev_num >= 2:
                task_status = TaskStatus.COMPLETED
                internship_status = InternshipStatus.COMPLETED

            with database.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    UPDATE students
                    SET task_status = ?,
                        completion_percentage = ?,
                        internship_status = ?,
                        latest_ai_feedback = ?,
                        remarks = ?,
                        updated_at = CURRENT_TIMESTAMP
                    WHERE student_id = ?
                """, (
                    task_status,
                    new_completion_pct,
                    internship_status,
                    f"Approved by {reviewer_name} (Score: {final_score}/100). {reviewer_notes}",
                    f"Task {task_id} approved. Internship status: {internship_status}.",
                    student_id
                ))
                # Update review schedule status to Completed
                cursor.execute("""
                    UPDATE review_schedules
                    SET review_status = 'Completed', feedback_status = 'Approved'
                    WHERE student_id = ? AND task_id = ? AND review_number = ?
                """, (student_id, task_id, rev_num))
                conn.commit()

            # Dispatch Task Approved / Completed Email
            ctx = {
                "student_id": student_id,
                "student_name": student["student_name"],
                "task_id": task_id,
                "task_title": student.get("assigned_task_title", task_id),
                "reviewer": reviewer_name,
                "score": final_score,
                "completion_percentage": new_completion_pct,
                "internship_status": internship_status,
                "remarks": reviewer_notes or "Outstanding work satisfying all acceptance criteria."
            }
            self.email_service.send_email(
                student_id=student_id,
                recipient_email=student["student_email"],
                email_type="Task Completed Notice",
                subject=f"[{task_id}] Evaluation Approved - Criteria Satisfied ({final_score}/100)",
                template_name="task_completed.html",
                context=ctx
            )

            result_payload["task_status"] = task_status
            result_payload["internship_status"] = internship_status
            msg = f"Task {task_id} approved for {student_id}. Score: {final_score}."

        else:
            # Corrections Required -> Schedule Next Review Session
            next_rev_num = int(student.get("current_review_number") or 1) + 1
            next_date = (datetime.now() + timedelta(days=DEFAULT_REVIEW_INTERVAL_DAYS)).strftime("%Y-%m-%d")
            next_time = "14:00"
            next_rev_id = f"REV-{student_id}-{next_rev_num}"
            meet_link = f"https://meet.google.com/rev-{student_id.lower()}-{next_rev_num}"

            with database.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    UPDATE students
                    SET task_status = 'Corrections Required',
                        next_review_date = ?,
                        next_review_time = ?,
                        current_review_number = ?,
                        latest_ai_feedback = ?,
                        remarks = ?,
                        updated_at = CURRENT_TIMESTAMP
                    WHERE student_id = ?
                """, (
                    next_date,
                    next_time,
                    next_rev_num,
                    f"Corrections Required (Score: {final_score}/100). Review #{next_rev_num} set for {next_date}.",
                    f"Resubmission required. Review #{next_rev_num} scheduled.",
                    student_id
                ))

                # Create next review schedule entry
                cursor.execute("""
                    INSERT INTO review_schedules (
                        review_id, student_id, student_name, student_email, task_id,
                        review_number, review_date, review_time, time_zone, meeting_gcr_link,
                        submission_status, review_status, reviewer, feedback_status
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(review_id) DO UPDATE SET
                        review_date=excluded.review_date,
                        review_time=excluded.review_time,
                        review_status='Scheduled'
                """, (
                    next_rev_id, student_id, student["student_name"], student["student_email"],
                    task_id, next_rev_num, next_date, next_time, TIMEZONE,
                    meet_link, "Pending Resubmission", ReviewStatus.SCHEDULED, reviewer_name, "Pending"
                ))
                conn.commit()

            # Dispatch Structured Feedback & Next Review Email
            ctx = {
                "student_id": student_id,
                "student_name": student["student_name"],
                "task_id": task_id,
                "task_title": student.get("assigned_task_title", task_id),
                "review_number": fb["review_number"],
                "reviewer": reviewer_name,
                "score": final_score,
                "passing_score": PASSING_SCORE_THRESHOLD,
                "submission_summary": fb["submission_summary"],
                "strengths": fb["strengths"],
                "identified_mistakes": fb["identified_mistakes"],
                "missing_requirements": fb["missing_requirements"],
                "improvement_suggestions": fb["improvement_suggestions"] + (f"\n\nMentor Note: {reviewer_notes}" if reviewer_notes else ""),
                "next_actions": fb["next_actions"],
                "next_review_date": next_date,
                "next_review_time": next_time,
                "time_zone": TIMEZONE,
                "meeting_link": meet_link,
                "submission_link": student.get("task_submission_link", "")
            }

            self.email_service.send_email(
                student_id=student_id,
                recipient_email=student["student_email"],
                email_type="Review Feedback Report",
                subject=f"[{task_id}] Review #{fb['review_number']} Feedback & Follow-up Session Scheduled",
                template_name="review_feedback.html",
                context=ctx,
                related_review_id=next_rev_id
            )

            result_payload["task_status"] = TaskStatus.CORRECTIONS_REQUIRED
            result_payload["next_review_id"] = next_rev_id
            result_payload["next_review_date"] = next_date
            msg = f"Corrections required for {student_id}. Review #{next_rev_num} scheduled for {next_date}."

        if sync_to_excel:
            excel_manager.export_sqlite_to_excel()

        return True, msg, result_payload

    def record_missed_review(
        self,
        review_id: str,
        reason: str = "Student did not attend scheduled review call.",
        sync_to_excel: bool = True
    ) -> Tuple[bool, str]:
        """Handles a missed review by recording it and scheduling a rescheduled session."""
        with database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM review_schedules WHERE review_id = ?", (review_id,))
            row = cursor.fetchone()
            if not row:
                return False, f"Review {review_id} not found."
            r = dict(row)

            # Mark current as Missed
            cursor.execute("""
                UPDATE review_schedules
                SET review_status = 'Missed', feedback_status = 'Missed'
                WHERE review_id = ?
            """, (review_id,))

            student_id = r["student_id"]
            next_date = (datetime.now() + timedelta(days=2)).strftime("%Y-%m-%d")
            next_time = r["review_time"]
            rescheduled_id = f"{review_id}-RESCHED"

            # Create rescheduled session
            cursor.execute("""
                INSERT INTO review_schedules (
                    review_id, student_id, student_name, student_email, task_id,
                    review_number, review_date, review_time, time_zone, meeting_gcr_link,
                    submission_status, review_status, reviewer, feedback_status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(review_id) DO NOTHING
            """, (
                rescheduled_id, student_id, r["student_name"], r["student_email"],
                r["task_id"], r["review_number"], next_date, next_time, TIMEZONE,
                r["meeting_gcr_link"], r["submission_status"], ReviewStatus.RESCHEDULED, r["reviewer"], "Pending"
            ))

            cursor.execute("""
                UPDATE students
                SET next_review_date = ?, next_review_time = ?, remarks = ?, updated_at = CURRENT_TIMESTAMP
                WHERE student_id = ?
            """, (next_date, next_time, f"Missed Review #{r['review_number']}. Rescheduled for {next_date}.", student_id))
            conn.commit()

        database.log_audit("REVIEW_MISSED", "Review", review_id, details=f"Marked missed. Rescheduled to {next_date}: {reason}")

        # Send alert to student
        student = database.get_student(student_id)
        if student:
            self.email_service.send_email(
                student_id=student_id,
                recipient_email=student["student_email"],
                email_type="Review Missed Alert",
                subject=f"[MISSED REVIEW] Rescheduled Session for {r['task_id']}",
                template_name="review_scheduled.html",
                context={
                    "student_id": student_id,
                    "student_name": student["student_name"],
                    "review_id": rescheduled_id,
                    "review_number": r["review_number"],
                    "task_id": r["task_id"],
                    "task_title": student.get("assigned_task_title", r["task_id"]),
                    "review_date": next_date,
                    "review_time": next_time,
                    "time_zone": TIMEZONE,
                    "reviewer": r["reviewer"],
                    "meeting_link": r["meeting_gcr_link"],
                    "submission_status": r["submission_status"],
                    "previous_feedback": f"ALERT: You missed your previous review on {r['review_date']}. Please ensure attendance on {next_date}."
                },
                related_review_id=rescheduled_id
            )

        if sync_to_excel:
            excel_manager.export_sqlite_to_excel()

        return True, f"Review {review_id} marked as missed. Rescheduled to {next_date} ({rescheduled_id})."
