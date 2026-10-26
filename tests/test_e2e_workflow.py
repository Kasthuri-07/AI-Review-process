"""
End-to-End Integration Test for Internship Workflow
Simulates a full student lifecycle:
Allocation -> Review 1 -> AI Evaluator (Incomplete) -> Reviewer Approval
-> Corrections Required -> Review 2 Auto-Scheduled -> Resubmission -> AI Evaluator (Passing)
-> Reviewer Sign-Off -> Task Completed -> Excel Export Sync.
"""

from datetime import datetime, timedelta
import openpyxl

import database
from config import TaskStatus, EXCEL_PATH
from task_allocator import TaskAllocator
from scheduler_service import ReviewSchedulerService
from ai_evaluator import AIEvaluator
from workflow_engine import WorkflowEngine
import excel_manager

def test_full_student_review_lifecycle_e2e():
    test_stu_id = "STU-E2E-TEST"
    
    # 1. Register a fresh test student in database
    with database.get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM task_allocations WHERE student_id = ?", (test_stu_id,))
        cursor.execute("DELETE FROM review_schedules WHERE student_id = ?", (test_stu_id,))
        cursor.execute("DELETE FROM ai_review_feedback WHERE student_id = ?", (test_stu_id,))
        cursor.execute("DELETE FROM email_logs WHERE student_id = ?", (test_stu_id,))
        cursor.execute("DELETE FROM students WHERE student_id = ?", (test_stu_id,))
        
        cursor.execute("""
            INSERT INTO students (
                student_id, student_name, student_email, internship_start_date, internship_end_date,
                internship_role, mentor_name, mentor_email, task_status, internship_status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            test_stu_id, "Devon Cross", "devon.cross@example-interns.edu",
            "2026-09-01", "2026-12-01", "AI/ML Engineer", "Dr. Aris Thorne",
            "aris.thorne@mentor-labs.org", "Not Started", "Active"
        ))
        conn.commit()

    # 2. Allocate Task Automatically
    allocator = TaskAllocator()
    alloc_ok, alloc_msg, alloc_data = allocator.allocate_task_to_student(
        student_id=test_stu_id,
        sync_to_excel=False
    )
    assert alloc_ok is True
    student = database.get_student(test_stu_id)
    assert student["assigned_task_id"] in ("TASK-AIML-101", "TASK-AIML-102")
    assigned_task = student["assigned_task_id"]
    assert student["task_status"] == TaskStatus.ASSIGNED

    # 3. Schedule Initial Review #1
    scheduler = ReviewSchedulerService()
    sched_date = (datetime.now() + timedelta(days=2)).strftime("%Y-%m-%d")
    sched_ok, sched_msg, _ = scheduler.schedule_review(
        student_id=test_stu_id,
        review_date_str=sched_date,
        review_time_str="14:00",
        sync_to_excel=False
    )
    assert sched_ok is True
    student = database.get_student(test_stu_id)
    assert student["current_review_number"] == 1
    assert student["next_review_date"] == sched_date

    # 4. Student Submits Incomplete Code -> AI Evaluation #1
    incomplete_code = "# Minimal code\nimport pandas as pd\nprint('incomplete')"
    evaluator = AIEvaluator(provider="MOCK")
    eval_1 = evaluator.evaluate_submission(
        student_id=test_stu_id,
        task_id=assigned_task,
        submission_text=incomplete_code
    )
    assert eval_1["score"] < 70
    assert eval_1["reviewer_approval_status"] == "Pending Reviewer Approval"

    # 5. Human Reviewer Approves Feedback #1 (Corrections Required)
    workflow = WorkflowEngine()
    adv_ok, adv_msg, res_payload = workflow.approve_feedback_and_advance(
        feedback_id=eval_1["feedback_id"],
        reviewer_name="Dr. Aris Thorne",
        reviewer_notes="Please implement proper exception handling and test assertions.",
        sync_to_excel=False
    )
    assert adv_ok is True
    assert res_payload["task_status"] == TaskStatus.CORRECTIONS_REQUIRED
    
    # Verify Review #2 was automatically scheduled
    student = database.get_student(test_stu_id)
    assert student["current_review_number"] == 2
    assert student["task_status"] == TaskStatus.CORRECTIONS_REQUIRED

    # 6. Student Addresses Feedback and Resubmits High-Quality Code -> AI Evaluation #2
    high_quality_code = """
import os
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from sklearn.ensemble import RandomForestClassifier

app = FastAPI(title="Refined Production Pipeline")

class ModelPayload(BaseModel):
    tenure: int
    charges: float

@app.post("/predict")
def predict_endpoint(payload: ModelPayload):
    try:
        # Proper error handling
        result = 1 if payload.charges > 50 else 0
        return {"prediction": result, "status": "success"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

def test_prediction_output():
    res = predict_endpoint(ModelPayload(tenure=12, charges=80.0))
    assert res["status"] == "success"
    assert res["prediction"] == 1
"""
    eval_2 = evaluator.evaluate_submission(
        student_id=test_stu_id,
        task_id=assigned_task,
        submission_text=high_quality_code
    )
    assert eval_2["score"] >= 70

    # 7. Reviewer Approves Feedback #2 -> Task Completed!
    adv_ok_2, adv_msg_2, res_payload_2 = workflow.approve_feedback_and_advance(
        feedback_id=eval_2["feedback_id"],
        reviewer_name="Dr. Aris Thorne",
        reviewer_notes="Excellent revision addressing all comments.",
        score_override=92,
        sync_to_excel=False
    )
    assert adv_ok_2 is True
    assert res_payload_2["task_status"] in (TaskStatus.APPROVED, TaskStatus.COMPLETED)

    # 8. Verify Excel Export with the complete updated status
    excel_manager.export_sqlite_to_excel()
    wb = openpyxl.load_workbook(str(EXCEL_PATH), data_only=True)
    ws_stu = wb["Student Master"]
    
    found_student = False
    for row in ws_stu.iter_rows(min_row=2, values_only=True):
        if row[0] == test_stu_id:
            found_student = True
            # Check task status column (column index 13, 0-indexed)
            assert row[13] in (TaskStatus.APPROVED, TaskStatus.COMPLETED)
            break
    assert found_student is True, "Test student should appear in exported Excel workbook."
