"""
Main CLI Application Entrypoint
AI-Powered Internship Task Allocation, Review and Feedback Automation System
"""

import sys
import time
import argparse

from config import EXCEL_PATH
import database
import init_system
import excel_manager
from task_allocator import TaskAllocator
from scheduler_service import ReviewSchedulerService
from ai_evaluator import AIEvaluator

def main():
    parser = argparse.ArgumentParser(
        description="AI-Powered Internship Task Allocation, Review and Feedback Automation System"
    )
    parser.add_argument("--init", action="store_true", help="Initialize SQLite DB and generate Excel workbook with 100 students.")
    parser.add_argument("--allocate", action="store_true", help="Auto-allocate tasks to all unassigned students.")
    parser.add_argument("--reminders", action="store_true", help="Scan and trigger due reminders (24h, 1h, overdue).")
    parser.add_argument("--sync-to-excel", action="store_true", help="Export latest SQLite state to Excel workbook.")
    parser.add_argument("--sync-from-excel", action="store_true", help="Ingest manual edits from Excel workbook into SQLite.")
    parser.add_argument("--run-scheduler", action="store_true", help="Keep the APScheduler running in the foreground daemon mode.")
    parser.add_argument("--evaluate", type=str, metavar="STUDENT_ID", help="Run AI evaluation on a specific student's submission.")
    parser.add_argument("--task-id", type=str, help="Task ID for manual evaluation (optional).")
    parser.add_argument("--send-test-email", action="store_true", help="Send a test notification from sender to receiver email.")

    args = parser.parse_args()

    if len(sys.argv) == 1:
        parser.print_help()
        print("\nTip: Run 'streamlit run dashboard_app.py' to launch the interactive web dashboard.")
        return

    if args.init:
        print("[CLI] Running full system initialization...")
        init_system.initialize_all()

    if args.allocate:
        print("[CLI] Executing task auto-allocation...")
        allocator = TaskAllocator()
        res = allocator.auto_allocate_all_pending(sync_to_excel=True)
        print(f"[CLI] Allocation complete. Assigned {res['success']} students. ({res['failed']} failed).")

    if args.reminders:
        print("[CLI] Scanning review schedules for due reminders and overdue tasks...")
        scheduler = ReviewSchedulerService()
        res = scheduler.process_due_reminders()
        print(f"[CLI] Reminders sent: 24h: {res['reminders_24h_sent']}, 1h: {res['reminders_1h_sent']}, Overdue: {res['overdue_notices_sent']}")

    if args.sync_to_excel:
        print(f"[CLI] Exporting SQLite state to {EXCEL_PATH}...")
        excel_manager.export_sqlite_to_excel()
        print("[CLI] Export complete.")

    if args.sync_from_excel:
        print(f"[CLI] Ingesting edits from {EXCEL_PATH} into SQLite...")
        excel_manager.sync_excel_to_sqlite()
        print("[CLI] Synchronization complete.")

    if args.evaluate:
        stu_id = args.evaluate.strip()
        print(f"[CLI] Running AI evaluation for student: {stu_id}...")
        student = database.get_student(stu_id)
        if not student:
            print(f"[ERROR] Student {stu_id} not found in database.")
            return
        task_id = args.task_id or student.get("assigned_task_id")
        if not task_id:
            print(f"[ERROR] Student {stu_id} has no assigned task.")
            return
        evaluator = AIEvaluator()
        result = evaluator.evaluate_submission(student_id=stu_id, task_id=task_id)
        print("\n" + "="*50)
        print(f"AI EVALUATION REPORT FOR {stu_id} ({task_id})")
        print("="*50)
        print(f"Score: {result['score']}/100")
        print(f"Summary: {result['submission_summary']}")
        print(f"Strengths: {result['what_was_done_correctly']}")
        print(f"Mistakes: {result['identified_mistakes']}")
        print(f"Missing: {result['missing_requirements']}")
        print(f"Corrections: {result['recommended_corrections']}")
        print(f"Suggested Next Review: {result['suggested_next_review_date']} at {result['suggested_next_review_time']}")
        print("="*50 + "\n")

    if args.send_test_email:
        from config import SENDER_EMAIL, RECEIVER_EMAIL
        from email_service import EmailService
        print(f"[CLI] Sending test email from {SENDER_EMAIL} to {RECEIVER_EMAIL} (Student: Kasthuri)...")
        service = EmailService()
        ok, msg = service.send_direct_notification(
            recipient_email=RECEIVER_EMAIL,
            subject="[TEST] AI Internship System Verification",
            message_body="This is an automated test message verifying sender and receiver email integration for Kasthuri.",
            student_id="STU-001"
        )
        print(f"[CLI] Result: {msg}")

    if args.run_scheduler:
        print("[CLI] Starting APScheduler daemon with persistent SQLite job store...")
        scheduler = ReviewSchedulerService()
        scheduler.start()
        print("[CLI] Scheduler is active. Press Ctrl+C to terminate.")
        try:
            while True:
                time.sleep(1)
        except (KeyboardInterrupt, SystemExit):
            print("\n[CLI] Shutting down scheduler gracefully...")
            scheduler.shutdown()
            print("[CLI] Scheduler stopped.")

if __name__ == "__main__":
    main()
