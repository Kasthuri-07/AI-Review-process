"""
Automatic Task Allocation Engine
Allocates curriculum tasks to students based on configured track rules or manual overrides,
generates personalized allocation emails, updates database and Excel, and prevents duplicates.
"""

from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional, Tuple
import database
from config import TaskStatus
from email_service import EmailService
import excel_manager

class TaskAllocator:
    def __init__(self, email_service: Optional[EmailService] = None):
        self.email_service = email_service or EmailService()

    def get_unassigned_students(self) -> List[Dict[str, Any]]:
        """Returns all students without an active assigned task."""
        with database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT * FROM students
                WHERE assigned_task_id IS NULL OR assigned_task_id = '' OR task_status = 'Not Started'
                ORDER BY student_id ASC
            """)
            return [dict(row) for row in cursor.fetchall()]

    def allocate_task_to_student(
        self,
        student_id: str,
        task_id: Optional[str] = None,
        force_resend_email: bool = False,
        sync_to_excel: bool = True
    ) -> Tuple[bool, str, Dict[str, Any]]:
        """
        Allocates a task to a student (automatic by role, or manually by task_id).
        Dispatches personalized email, updates SQLite, and optionally syncs to Excel.
        """
        student = database.get_student(student_id)
        if not student:
            return False, f"Student {student_id} does not exist.", {}

        # 1. Determine task to assign
        task_to_assign = None
        if task_id:
            # Manual assignment
            with database.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM task_catalog WHERE task_id = ?", (task_id,))
                row = cursor.fetchone()
                if row:
                    task_to_assign = dict(row)
                else:
                    return False, f"Task {task_id} not found in catalog.", {}
        else:
            # Automatic assignment by student's role
            role = student["internship_role"]
            with database.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM task_catalog WHERE track_role = ? ORDER BY task_id ASC", (role,))
                tasks = [dict(r) for r in cursor.fetchall()]
                if not tasks:
                    return False, f"No catalog tasks found matching role '{role}'.", {}
                
                # Pick task balancing distribution
                stu_num = int("".join(filter(str.isdigit, student_id)) or "1")
                task_to_assign = tasks[(stu_num - 1) % len(tasks)]

        # 2. Compute Dates
        now = datetime.now()
        alloc_date_str = now.strftime("%Y-%m-%d")
        est_days = task_to_assign.get("estimated_days", 7)
        deadline_date_str = (now + timedelta(days=est_days)).strftime("%Y-%m-%d")
        
        gcr_link = student["gcr_link"] or f"https://classroom.google.com/c/internship-{student['internship_role'].lower().replace(' ', '-').replace('/', '-')}-2026"
        submission_link = f"https://classroom.google.com/c/submissions/{student_id.lower()}_{task_to_assign['task_id'].lower()}"

        # 3. Update Student Record in DB
        with database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE students
                SET assigned_task_id = ?,
                    assigned_task_title = ?,
                    task_description = ?,
                    task_allocation_date = ?,
                    task_deadline = ?,
                    task_status = ?,
                    gcr_link = ?,
                    task_submission_link = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE student_id = ?
            """, (
                task_to_assign["task_id"],
                task_to_assign["task_title"],
                task_to_assign["task_description"],
                alloc_date_str,
                deadline_date_str,
                TaskStatus.ASSIGNED,
                gcr_link,
                submission_link,
                student_id
            ))

            # Record in task_allocations table
            cursor.execute("""
                INSERT INTO task_allocations (
                    task_id, task_title, task_description, student_id, student_name,
                    student_email, assigned_date, deadline, task_priority, required_deliverables,
                    submission_link, task_status, allocation_email_status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                task_to_assign["task_id"],
                task_to_assign["task_title"],
                task_to_assign["task_description"],
                student_id,
                student["student_name"],
                student["student_email"],
                alloc_date_str,
                deadline_date_str,
                task_to_assign.get("task_priority", "Medium"),
                task_to_assign["required_deliverables"],
                submission_link,
                TaskStatus.ASSIGNED,
                "Pending"
            ))
            conn.commit()

        database.log_audit(
            action="TASK_ALLOCATED",
            entity="Student",
            entity_id=student_id,
            details=f"Assigned task {task_to_assign['task_id']} to {student['student_name']}"
        )

        # 4. Dispatch Personalized Allocation Email
        email_context = {
            "student_id": student_id,
            "student_name": student["student_name"],
            "internship_role": student["internship_role"],
            "task_id": task_to_assign["task_id"],
            "task_title": task_to_assign["task_title"],
            "task_description": task_to_assign["task_description"],
            "required_deliverables": task_to_assign["required_deliverables"],
            "task_priority": task_to_assign.get("task_priority", "Medium"),
            "assigned_date": alloc_date_str,
            "deadline": deadline_date_str,
            "mentor_name": student["mentor_name"],
            "mentor_email": student["mentor_email"],
            "submission_link": submission_link
        }

        subject = f"[{task_to_assign['task_id']}] Official Internship Task Allocation & Deliverables"
        sent_ok, email_msg = self.email_service.send_email(
            student_id=student_id,
            recipient_email=student["student_email"],
            email_type="Task Allocation",
            subject=subject,
            template_name="task_allocation.html",
            context=email_context,
            force_resend=force_resend_email
        )

        # Update allocation email status in task_allocations table
        with database.get_connection() as conn:
            cursor = conn.cursor()
            status_text = "Sent" if sent_ok else "Failed"
            cursor.execute("""
                UPDATE task_allocations
                SET allocation_email_status = ?
                WHERE student_id = ? AND task_id = ?
            """, (status_text, student_id, task_to_assign["task_id"]))
            conn.commit()

        # 5. Optionally export to Excel
        if sync_to_excel:
            excel_manager.export_sqlite_to_excel()

        result_data = {
            "student_id": student_id,
            "task_id": task_to_assign["task_id"],
            "task_title": task_to_assign["task_title"],
            "deadline": deadline_date_str,
            "email_sent": sent_ok,
            "email_message": email_msg
        }

        return True, f"Allocated {task_to_assign['task_id']} to {student_id}. {email_msg}", result_data

    def auto_allocate_all_pending(self, sync_to_excel: bool = True) -> Dict[str, Any]:
        """Automatically assigns tasks and dispatches emails for all unassigned students."""
        unassigned = self.get_unassigned_students()
        results = {"total": len(unassigned), "success": 0, "failed": 0, "details": []}

        for stu in unassigned:
            ok, msg, data = self.allocate_task_to_student(
                student_id=stu["student_id"],
                sync_to_excel=False  # Batch sync at the end
            )
            if ok:
                results["success"] += 1
            else:
                results["failed"] += 1
            results["details"].append({"student_id": stu["student_id"], "success": ok, "message": msg})

        if sync_to_excel and results["success"] > 0:
            excel_manager.export_sqlite_to_excel()

        return results
