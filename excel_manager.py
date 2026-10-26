"""
Excel Manager Module (OpenPyXL)
Generates the comprehensive 6-sheet workbook 'AI_Internship_Review_Management.xlsx'
with professional formatting, formulas, dropdowns, conditional formatting, and dashboard.
Also handles bidirectional synchronization between SQLite and Excel with safe versioned backups.
"""

import shutil
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Dict, Any, Tuple
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, Reference

from config import (
    EXCEL_PATH, BACKUP_DIR, ROLES_TRACKS,
    TaskStatus, InternshipStatus, ReviewStatus, EmailStatus
)
import database

# Styling Palettes - Modern Indigo & Slate Theme
COLOR_HEADER_NAVY = "1E293B"       # Deep Slate Navy
COLOR_HEADER_TEXT = "FFFFFF"
COLOR_ACCENT_BLUE = "2563EB"       # Royal Blue
COLOR_ACCENT_LIGHT = "EFF6FF"
COLOR_ZEBRA_LIGHT = "F8FAFC"
COLOR_BORDER = "CBD5E1"

# Status colors for conditional formatting / badge styling
COLOR_SUCCESS_BG = "DCFCE7"  # Soft green
COLOR_SUCCESS_FG = "166534"
COLOR_WARNING_BG = "FEF3C7"  # Soft amber
COLOR_WARNING_FG = "92400E"
COLOR_DANGER_BG = "FEE2E2"   # Soft red
COLOR_DANGER_FG = "991B1B"
COLOR_INFO_BG = "DBEAFE"     # Soft blue
COLOR_INFO_FG = "1E40AF"

FONT_FAMILY = "Segoe UI"
HEADER_FONT = Font(name=FONT_FAMILY, size=11, bold=True, color=COLOR_HEADER_TEXT)
HEADER_FILL = PatternFill(start_color=COLOR_HEADER_NAVY, end_color=COLOR_HEADER_NAVY, fill_type="solid")
REGULAR_FONT = Font(name=FONT_FAMILY, size=10)
BOLD_FONT = Font(name=FONT_FAMILY, size=10, bold=True)
THIN_SIDE = Side(border_style="thin", color=COLOR_BORDER)
CELL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

# Catalog of realistic tasks mapped to tracks
TASK_CATALOG_DATA = [
    {
        "task_id": "TASK-AIML-101",
        "track_role": "AI/ML Engineer",
        "task_title": "Customer Churn Prediction Pipeline with Scikit-Learn & FastAPI",
        "task_description": "Build an end-to-end churn prediction model on telecom customer data. Preprocess tabular data, train Random Forest / XGBoost, evaluate ROC-AUC, and expose inference via a FastAPI REST endpoint with Dockerfile.",
        "required_deliverables": "1. Python model training script (train.py)\n2. Serialized model artifact (.joblib)\n3. FastAPI application (app.py) with /predict\n4. Dockerfile & requirements.txt\n5. Model card & validation report PDF",
        "evaluation_rubric": "Data preprocessing (20 pts), Model tuning & ROC-AUC > 0.82 (30 pts), API design & error handling (25 pts), Documentation & Docker setup (25 pts).",
        "task_priority": "High",
        "estimated_days": 10
    },
    {
        "task_id": "TASK-AIML-102",
        "track_role": "AI/ML Engineer",
        "task_title": "RAG Knowledge Base Assistant using LangChain & FAISS",
        "task_description": "Create a domain-specific Q&A chatbot using LangChain, text chunking, FAISS vector embeddings, and an open-source LLM or OpenAI API. Provide semantic citations and handle out-of-domain queries gracefully.",
        "required_deliverables": "1. Document ingestion & chunking pipeline\n2. Vector indexing script\n3. Retrieval chain module with source citation\n4. Streamlit UI interface\n5. Architecture diagram and evaluation benchmark",
        "evaluation_rubric": "Chunking strategy (20 pts), Retrieval precision (30 pts), Hallucination guardrails (25 pts), UI polish & code quality (25 pts).",
        "task_priority": "High",
        "estimated_days": 12
    },
    {
        "task_id": "TASK-FSWD-201",
        "track_role": "Full Stack Web Developer",
        "task_title": "E-Commerce REST API with JWT Auth, Cart & Payment Gateway",
        "task_description": "Design a secure REST API backend supporting user authentication (JWT + bcrypt), product catalog with pagination, shopping cart state management, and mock payment gateway integration.",
        "required_deliverables": "1. Modular backend code (Express or FastAPI)\n2. Database migration schema (PostgreSQL/SQLite)\n3. Postman/Swagger API collection with tests\n4. Unit and integration test suite (>80% coverage)\n5. README with local deployment steps",
        "evaluation_rubric": "Auth & Security (25 pts), Database design & relational integrity (25 pts), API endpoints correctness (25 pts), Test suite & docs (25 pts).",
        "task_priority": "High",
        "estimated_days": 10
    },
    {
        "task_id": "TASK-FSWD-202",
        "track_role": "Full Stack Web Developer",
        "task_title": "Real-Time Collaboration Kanban Board with WebSockets",
        "task_description": "Develop a multi-tenant Kanban board application where updates to columns, cards, and drag-and-drop actions are synchronized in real-time across connected browser clients using WebSockets.",
        "required_deliverables": "1. Frontend single page app with drag-and-drop\n2. WebSocket server with room management\n3. Optimistic UI updates with rollback\n4. Responsive CSS layout\n5. System architecture document",
        "evaluation_rubric": "WebSocket efficiency & reconnect logic (30 pts), Drag-and-drop state sync (30 pts), Code structure & modularity (20 pts), UI responsiveness (20 pts).",
        "task_priority": "Medium",
        "estimated_days": 10
    },
    {
        "task_id": "TASK-DATA-301",
        "track_role": "Data Analyst",
        "task_title": "Automated Multi-Source Retail Sales ETL & KPI Dashboard",
        "task_description": "Extract raw CSV and JSON transaction logs, clean nulls/outliers in Pandas, load into a star-schema database, and construct an interactive dashboard featuring MoM growth, cohort retention, and customer lifetime value.",
        "required_deliverables": "1. Automated ETL python script with logging\n2. Cleaned dimensional data warehouse tables\n3. Interactive dashboard (Streamlit / Dash)\n4. Business insight briefing report (PDF/Markdown)\n5. Data dictionary & pipeline diagram",
        "evaluation_rubric": "ETL robustness & data hygiene (30 pts), Metric calculations (CLV, Cohorts) (30 pts), Visual layout & interactivity (25 pts), Executive summary quality (15 pts).",
        "task_priority": "Medium",
        "estimated_days": 8
    },
    {
        "task_id": "TASK-DATA-302",
        "track_role": "Data Analyst",
        "task_title": "Healthcare Patient Readmission Risk Factor Statistical Study",
        "task_description": "Perform rigorous hypothesis testing, exploratory data analysis, and logistic regression on clinical patient records to isolate top behavioral and demographic risk drivers for 30-day hospital readmission.",
        "required_deliverables": "1. Jupyter notebook with reproducible EDA\n2. Statistical hypothesis tests (t-tests, Chi-square)\n3. Risk scoring model with odds ratios\n4. Executive slide deck summarizing findings\n5. Clean Python script version",
        "evaluation_rubric": "Statistical validity (35 pts), Visual communication (25 pts), Insight depth & actionable recommendations (25 pts), Code reproducibility (15 pts).",
        "task_priority": "Medium",
        "estimated_days": 9
    },
    {
        "task_id": "TASK-DEVO-401",
        "track_role": "Cloud & DevOps Engineer",
        "task_title": "Multi-Environment CI/CD Pipeline with Terraform & Docker",
        "task_description": "Provision AWS/LocalStack infrastructure using modular Terraform (VPC, ECS, S3). Create a GitHub Actions workflow that executes linting, container vulnerability scanning (Trivy), automated testing, and zero-downtime deployment.",
        "required_deliverables": "1. Terraform configuration files with remote state\n2. Dockerfile with multi-stage build optimization\n3. GitHub Actions workflow .yml files\n4. Security scanning report with remediation\n5. Deployment verification guide",
        "evaluation_rubric": "Infrastructure as Code standards (30 pts), Pipeline stages & security scanning (30 pts), Docker image optimization (20 pts), Error handling & rollback logic (20 pts).",
        "task_priority": "High",
        "estimated_days": 11
    },
    {
        "task_id": "TASK-DEVO-402",
        "track_role": "Cloud & DevOps Engineer",
        "task_title": "Kubernetes Microservices Observability with Prometheus & Grafana",
        "task_description": "Deploy a 3-tier microservice onto Kubernetes (Minikube / k3s). Implement metric scraping with Prometheus, custom application dashboards in Grafana, and Slack/email Alertmanager rules for pod crash and latency spikes.",
        "required_deliverables": "1. Kubernetes deployment and service manifests\n2. Prometheus ServiceMonitor & scrape configs\n3. Exported Grafana dashboard JSON\n4. Alertmanager rules YAML\n5. Chaos testing documentation and proof of alerts",
        "evaluation_rubric": "K8s manifest hygiene (25 pts), Observability coverage (Golden signals) (35 pts), Alerting precision (20 pts), Chaos testing validation (20 pts).",
        "task_priority": "High",
        "estimated_days": 11
    },
    {
        "task_id": "TASK-MOBL-501",
        "track_role": "Mobile App Developer",
        "task_title": "Cross-Platform Habit Tracker with Offline SQLite & Push Alerts",
        "task_description": "Build a cross-platform mobile application (React Native or Flutter) that allows users to track daily habits, streaks, and completion analytics. Implement local SQLite persistence with background push notification reminders.",
        "required_deliverables": "1. Mobile application source code\n2. Local SQLite database schema and ORM layer\n3. Scheduled local notification service\n4. APK / test bundle\n5. Architecture and state management breakdown",
        "evaluation_rubric": "Offline sync & local persistence (30 pts), UI fluid animations & responsiveness (30 pts), Notification reliability (20 pts), Clean architecture (20 pts).",
        "task_priority": "Medium",
        "estimated_days": 10
    },
    {
        "task_id": "TASK-MOBL-502",
        "track_role": "Mobile App Developer",
        "task_title": "Campus Delivery App with Real-Time Geolocation Tracking",
        "task_description": "Develop a delivery tracking client that integrates map SDKs, updates user location in the background, displays delivery driver position on a map in real time, and alerts the user when arrival is within 200 meters.",
        "required_deliverables": "1. Application frontend with map integration\n2. Geofencing & proximity calculation logic\n3. State management for location streams\n4. Mock location generator for testing\n5. User guide and test videos",
        "evaluation_rubric": "Map & geolocation performance (35 pts), Battery & permission management (25 pts), State handling (25 pts), UI polish (15 pts).",
        "task_priority": "High",
        "estimated_days": 12
    },
    {
        "task_id": "TASK-CYBR-601",
        "track_role": "Cybersecurity Specialist",
        "task_title": "Automated Web Vulnerability Scanner & OWASP Top 10 Audit",
        "task_description": "Develop a Python-based security scanner targeting authorized test applications (e.g., OWASP Juice Shop). Audit for SQLi, XSS, insecure headers, and open redirects, producing a CVSS-scored executive vulnerability report.",
        "required_deliverables": "1. Modular vulnerability scanning tool in Python\n2. Custom payloads and regex verification rules\n3. HTML and JSON vulnerability report generator\n4. Remediated sample code examples\n5. Ethical hacking compliance agreement & scope doc",
        "evaluation_rubric": "Scanner accuracy & low false-positive rate (35 pts), CVSS scoring & risk categorization (25 pts), Remediation guidance (25 pts), Code safety & ethics (15 pts).",
        "task_priority": "High",
        "estimated_days": 10
    },
    {
        "task_id": "TASK-CYBR-602",
        "track_role": "Cybersecurity Specialist",
        "task_title": "SOC Intrusion Detection & SIEM Log Correlation Engine",
        "task_description": "Construct an automated log correlation parser that ingests syslog and web access logs, detects brute-force authentication attacks and lateral movement patterns, and generates automated incident response tickets.",
        "required_deliverables": "1. Log ingestion and parsing engine (regex/JSON)\n2. Sigma-compatible detection rules\n3. Alert correlation and threshold detector\n4. Incident response triage playbook\n5. Synthetic log test generator",
        "evaluation_rubric": "Rule precision & detection logic (35 pts), Parser throughput & scalability (25 pts), Triage workflow integration (25 pts), Test dataset & coverage (15 pts).",
        "task_priority": "High",
        "estimated_days": 10
    }
]

# Mentors
MENTORS = [
    {"name": "Dr. Aris Thorne", "email": "aris.thorne@mentor-labs.org", "track": "AI/ML Engineer"},
    {"name": "Prof. Elena Rostova", "email": "elena.rostova@mentor-labs.org", "track": "Full Stack Web Developer"},
    {"name": "Sarah Jenkins", "email": "sarah.jenkins@mentor-labs.org", "track": "Data Analyst"},
    {"name": "Rajiv Mehta", "email": "rajiv.mehta@mentor-labs.org", "track": "Cloud & DevOps Engineer"},
    {"name": "Marcus Chen", "email": "marcus.chen@mentor-labs.org", "track": "Mobile App Developer"},
    {"name": "Nadia Al-Mansoor", "email": "nadia.mansoor@mentor-labs.org", "track": "Cybersecurity Specialist"}
]

# Sample First and Last Names for generating 100 students
FIRST_NAMES = [
    "Aarav", "Aditi", "Alexander", "Amara", "Carlos", "Chloe", "David", "Divya", "Ethan", "Fatima",
    "Gabriel", "Grace", "Hassan", "Isabella", "James", "Jia", "Kavya", "Liam", "Maya", "Noah",
    "Olivia", "Priya", "Quentin", "Rohan", "Sophia", "Tariq", "Uma", "Victor", "Wei", "Zara",
    "Ananya", "Brian", "Camila", "Daniel", "Emma", "Farhan", "Gita", "Henry", "Ines", "Jack",
    "Kiran", "Lucas", "Mei", "Nathan", "Oscar", "Pooja", "Rahul", "Sara", "Tom", "Valeria"
]
LAST_NAMES = [
    "Patel", "Sharma", "Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis",
    "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore",
    "Jackson", "Martin", "Lee", "Perez", "Thompson", "White", "Harris", "Sanchez", "Clark", "Ramirez",
    "Lewis", "Robinson", "Walker", "Young", "Allen", "King", "Wright", "Scott", "Torres", "Nguyen",
    "Hill", "Flores", "Green", "Adams", "Nelson", "Baker", "Hall", "Rivera", "Campbell", "Mitchell"
]

def generate_100_students_dataset() -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    Generates 100 consistent, realistic sample student records,
    along with their task allocations, review schedules, AI feedbacks, and email logs.
    All addresses use example-interns.edu (explicit sample addresses).
    """
    students = []
    allocations = []
    reviews = []
    feedbacks = []
    email_logs = []

    base_start = datetime(2026, 9, 1)
    base_end = datetime(2026, 12, 1)

    tasks_by_track = {}
    for t in TASK_CATALOG_DATA:
        tasks_by_track.setdefault(t["track_role"], []).append(t)

    mentor_by_track = {m["track"]: m for m in MENTORS}

    # Status distribution across the 100 students for realistic diversity:
    # 0-14: Approved / Completed tasks (some finished internship)
    # 15-34: Under Review / Submitted (awaiting AI/Mentor feedback)
    # 35-54: Corrections Required / Resubmission Pending (Review 2 scheduled)
    # 55-79: In Progress / Assigned (working towards first review)
    # 80-89: Overdue submissions
    # 90-99: Newly Allocated / Not Started
    
    email_counter = 1000

    for i in range(1, 101):
        stu_id = f"STU-{i:03d}"
        if i == 1:
            stu_name = "Kasthuri"
            stu_email = "kasthurikumar1105@gmail.com"
        else:
            f_name = FIRST_NAMES[(i - 1) % len(FIRST_NAMES)]
            l_name = LAST_NAMES[((i - 1) * 3 + 7) % len(LAST_NAMES)]
            stu_name = f"{f_name} {l_name}"
            stu_email = f"student{i:03d}@example-interns.edu"
        role = ROLES_TRACKS[(i - 1) % len(ROLES_TRACKS)]
        mentor = mentor_by_track[role].copy()
        if i == 1:
            mentor["email"] = "kasthuricute273@gmail.com"
        track_tasks = tasks_by_track[role]
        assigned_task = track_tasks[(i - 1) % len(track_tasks)]

        start_date = (base_start + timedelta(days=(i % 15))).strftime("%Y-%m-%d")
        end_date = (base_end + timedelta(days=(i % 15))).strftime("%Y-%m-%d")
        alloc_date = (base_start + timedelta(days=2 + (i % 10))).strftime("%Y-%m-%d")
        deadline = (datetime.strptime(alloc_date, "%Y-%m-%d") + timedelta(days=assigned_task["estimated_days"])).strftime("%Y-%m-%d")

        gcr_link = f"https://classroom.google.com/c/internship-{role.lower().replace(' ', '-').replace('/', '-')}-2026"
        submission_link = f"https://classroom.google.com/c/submissions/{stu_id.lower()}_{assigned_task['task_id'].lower()}"

        # Assign state based on index tier
        if i <= 10:
            # Fully Completed Internship
            task_status = TaskStatus.COMPLETED
            internship_status = InternshipStatus.COMPLETED
            completion_pct = 100.0
            rev_num = 2
            last_rev_date = "2026-10-15"
            next_rev_date = ""
            next_rev_time = ""
            email_status = EmailStatus.SENT
            latest_ai_fb = "Exceptional submission meeting all rubric specifications. Architecture, test coverage, and documentation were outstanding."
            remarks = "Internship successfully completed with Honors. Ready for certificate issuance."
        elif i <= 20:
            # Approved Task, working on next phase
            task_status = TaskStatus.APPROVED
            internship_status = InternshipStatus.IN_PROGRESS
            completion_pct = 75.0
            rev_num = 2
            last_rev_date = "2026-10-20"
            next_rev_date = "2026-11-05"
            next_rev_time = "14:00"
            email_status = EmailStatus.SENT
            latest_ai_fb = "Core criteria approved. Minor optimization tips noted. Proceed to final capstone integration."
            remarks = "Primary task approved. Capstone briefing scheduled."
        elif i <= 40:
            # Submitted / Under Review
            task_status = TaskStatus.UNDER_REVIEW
            internship_status = InternshipStatus.IN_PROGRESS
            completion_pct = 50.0
            rev_num = 1
            last_rev_date = ""
            next_rev_date = "2026-10-27"
            next_rev_time = "11:30"
            email_status = EmailStatus.SENT
            latest_ai_fb = "Initial automated rubric evaluation pending human mentor sign-off."
            remarks = "Submission received on time. Scheduled for formal mentor review."
        elif i <= 60:
            # Corrections Required (Follow-up Review 2)
            task_status = TaskStatus.CORRECTIONS_REQUIRED
            internship_status = InternshipStatus.IN_PROGRESS
            completion_pct = 40.0
            rev_num = 1
            last_rev_date = "2026-10-18"
            next_rev_date = "2026-10-28"
            next_rev_time = "16:00"
            email_status = EmailStatus.SENT
            latest_ai_fb = "Identified missing unit tests and insufficient error handling. Review rubric items 2 and 4."
            remarks = "Corrections required on error handling and test suite. Review 2 scheduled."
        elif i <= 75:
            # In Progress / Assigned
            task_status = TaskStatus.IN_PROGRESS
            internship_status = InternshipStatus.IN_PROGRESS
            completion_pct = 25.0
            rev_num = 0
            last_rev_date = ""
            next_rev_date = "2026-10-30"
            next_rev_time = "10:00"
            email_status = EmailStatus.SENT
            latest_ai_fb = "Awaiting initial submission."
            remarks = "Student actively developing solution. Mid-point check-in scheduled."
        elif i <= 88:
            # Overdue
            task_status = TaskStatus.OVERDUE
            internship_status = InternshipStatus.IN_PROGRESS
            completion_pct = 15.0
            rev_num = 0
            last_rev_date = ""
            next_rev_date = "2026-10-26"
            next_rev_time = "17:00"
            email_status = EmailStatus.SENT
            latest_ai_fb = "No submission recorded by deadline."
            remarks = "Automated escalation email sent. Urgent submission required."
        else:
            # Newly Assigned / Not Started
            task_status = TaskStatus.ASSIGNED
            internship_status = InternshipStatus.ACTIVE
            completion_pct = 0.0
            rev_num = 0
            last_rev_date = ""
            next_rev_date = "2026-11-02"
            next_rev_time = "15:00"
            email_status = EmailStatus.PENDING
            latest_ai_fb = "Task allocated. Awaiting student onboarding."
            remarks = "Task allocation dispatched. Welcome kit sent."

        last_email_sent = (datetime.now() - timedelta(days=i % 7)).strftime("%Y-%m-%d %H:%M")

        student_record = {
            "student_id": stu_id,
            "student_name": stu_name,
            "student_email": stu_email,
            "internship_start_date": start_date,
            "internship_end_date": end_date,
            "internship_role": role,
            "mentor_name": mentor["name"],
            "mentor_email": mentor["email"],
            "assigned_task_id": assigned_task["task_id"],
            "assigned_task_title": assigned_task["task_title"],
            "task_description": assigned_task["task_description"],
            "task_allocation_date": alloc_date,
            "task_deadline": deadline,
            "task_status": task_status,
            "current_review_number": rev_num,
            "last_review_date": last_rev_date,
            "next_review_date": next_rev_date,
            "next_review_time": next_rev_time,
            "gcr_link": gcr_link,
            "task_submission_link": submission_link,
            "latest_ai_feedback": latest_ai_fb,
            "completion_percentage": completion_pct,
            "internship_status": internship_status,
            "email_delivery_status": email_status,
            "last_email_sent": last_email_sent,
            "remarks": remarks
        }
        students.append(student_record)

        # Task Allocation Record
        alloc_record = {
            "task_id": assigned_task["task_id"],
            "task_title": assigned_task["task_title"],
            "task_description": assigned_task["task_description"],
            "student_id": stu_id,
            "student_name": stu_name,
            "student_email": stu_email,
            "assigned_date": alloc_date,
            "deadline": deadline,
            "task_priority": assigned_task["task_priority"],
            "required_deliverables": assigned_task["required_deliverables"],
            "submission_link": submission_link,
            "task_status": task_status,
            "allocation_email_status": "Sent" if i <= 88 else "Pending"
        }
        allocations.append(alloc_record)

        # Review Schedule Record (if applicable)
        if next_rev_date:
            rev_id = f"REV-{stu_id}-{max(1, rev_num + 1)}"
            review_record = {
                "review_id": rev_id,
                "student_id": stu_id,
                "student_name": stu_name,
                "student_email": stu_email,
                "task_id": assigned_task["task_id"],
                "review_number": max(1, rev_num + 1),
                "review_date": next_rev_date,
                "review_time": next_rev_time or "14:00",
                "time_zone": "Asia/Kolkata",
                "meeting_gcr_link": f"https://meet.google.com/rev-{stu_id.lower()}",
                "submission_status": "Submitted" if task_status in (TaskStatus.SUBMITTED, TaskStatus.UNDER_REVIEW) else "Pending",
                "review_status": ReviewStatus.SCHEDULED if task_status != TaskStatus.COMPLETED else ReviewStatus.COMPLETED,
                "reviewer": mentor["name"],
                "feedback_status": "Pending",
                "next_review_date": (datetime.strptime(next_rev_date, "%Y-%m-%d") + timedelta(days=7)).strftime("%Y-%m-%d") if task_status != TaskStatus.COMPLETED else ""
            }
            reviews.append(review_record)

        # Sample AI Feedback records for reviewed students
        if i <= 60:
            score = 92 if i <= 10 else (84 if i <= 20 else (62 if i <= 40 else 54))
            approval_status = "Approved" if i <= 20 else "Pending Reviewer Approval"
            feedback_record = {
                "student_id": stu_id,
                "task_id": assigned_task["task_id"],
                "review_number": max(1, rev_num),
                "submission_summary": f"Completed initial implementation for {assigned_task['task_title']}.",
                "evaluation_criteria": assigned_task["evaluation_rubric"],
                "identified_mistakes": "None critical" if score >= 80 else "Missing input validation, hardcoded connection strings, and zero test coverage.",
                "missing_requirements": "None" if score >= 80 else "Required Dockerfile and automated integration tests were not found.",
                "strengths": "Modular structure, clean documentation, robust core logic implementation.",
                "improvement_suggestions": "Implement unit tests with Pytest, use environment variables for configs, and add Swagger API docstrings.",
                "ai_generated_feedback": f"The submission demonstrates good understanding of {role} workflows. Score: {score}/100. {'Pass with distinction.' if score >= 85 else 'Please address identified corrections before re-evaluation.'}",
                "score": score,
                "reviewer_approval_status": approval_status,
                "next_actions": "Proceed to next module" if score >= 70 else "Refactor error handling and re-submit for Review 2."
            }
            feedbacks.append(feedback_record)

        # Sample Email Log entries
        email_counter += 1
        email_record = {
            "email_id": f"EML-{email_counter}",
            "student_id": stu_id,
            "recipient_email": stu_email,
            "email_type": "Task Allocation" if i > 60 else "Review Feedback & Follow-up",
            "subject": f"[{assigned_task['task_id']}] Official Assignment & Internship Instructions" if i > 60 else f"Review #{rev_num} Feedback & Evaluation Report",
            "sent_date_time": last_email_sent,
            "delivery_status": "Sent" if i <= 95 else "Failed",
            "error_message": "" if i <= 95 else "SMTP Mailbox Full / Temporary Network Timeout",
            "retry_count": 0 if i <= 95 else 2,
            "related_review_id": f"REV-{stu_id}-1" if i <= 60 else ""
        }
        email_logs.append(email_record)

    return students, allocations, reviews, feedbacks, email_logs


def create_styled_workbook(
    students: List[Dict[str, Any]],
    allocations: List[Dict[str, Any]],
    reviews: List[Dict[str, Any]],
    feedbacks: List[Dict[str, Any]],
    email_logs: List[Dict[str, Any]],
    output_path: Path = EXCEL_PATH
):
    """
    Constructs the 6-worksheet Excel workbook with professional typography,
    zebra striping, explicit formulas, data validation dropdowns, and dashboard.
    """
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # ----------------------------------------------------
    # Sheet 1: Student Master
    # ----------------------------------------------------
    ws_student = wb.create_sheet(title="Student Master")
    student_headers = [
        "Student ID", "Student Name", "Student Email", "Internship Start Date", "Internship End Date",
        "Internship Role", "Mentor Name", "Mentor Email", "Assigned Task ID", "Assigned Task Title",
        "Task Description", "Task Allocation Date", "Task Deadline", "Task Status", "Current Review Number",
        "Last Review Date", "Next Review Date", "Next Review Time", "GCR / Google Classroom Link",
        "Task Submission Link", "Latest AI Feedback", "Completion Percentage", "Internship Status",
        "Email Delivery Status", "Last Email Sent", "Remarks"
    ]
    ws_student.append(student_headers)

    for stu in students:
        row = [
            stu["student_id"], stu["student_name"], stu["student_email"],
            stu["internship_start_date"], stu["internship_end_date"], stu["internship_role"],
            stu["mentor_name"], stu["mentor_email"], stu["assigned_task_id"],
            stu["assigned_task_title"], stu["task_description"], stu["task_allocation_date"],
            stu["task_deadline"], stu["task_status"], stu["current_review_number"],
            stu["last_review_date"], stu["next_review_date"], stu["next_review_time"],
            stu["gcr_link"], stu["task_submission_link"], stu["latest_ai_feedback"],
            stu["completion_percentage"], stu["internship_status"], stu["email_delivery_status"],
            stu["last_email_sent"], stu["remarks"]
        ]
        ws_student.append(row)

    # ----------------------------------------------------
    # Sheet 2: Task Allocation
    # ----------------------------------------------------
    ws_alloc = wb.create_sheet(title="Task Allocation")
    alloc_headers = [
        "Task ID", "Task Title", "Task Description", "Student ID", "Student Name",
        "Student Email", "Assigned Date", "Deadline", "Task Priority", "Required Deliverables",
        "Submission Link", "Task Status", "Allocation Email Status"
    ]
    ws_alloc.append(alloc_headers)

    for al in allocations:
        row = [
            al["task_id"], al["task_title"], al["task_description"], al["student_id"],
            al["student_name"], al["student_email"], al["assigned_date"], al["deadline"],
            al["task_priority"], al["required_deliverables"], al["submission_link"],
            al["task_status"], al["allocation_email_status"]
        ]
        ws_alloc.append(row)

    # ----------------------------------------------------
    # Sheet 3: Review Schedule
    # ----------------------------------------------------
    ws_rev = wb.create_sheet(title="Review Schedule")
    rev_headers = [
        "Review ID", "Student ID", "Student Name", "Student Email", "Task ID",
        "Review Number", "Review Date", "Review Time", "Time Zone", "Meeting / GCR Link",
        "Submission Status", "Review Status", "Reviewer", "Feedback Status", "Next Review Date"
    ]
    ws_rev.append(rev_headers)

    for r in reviews:
        row = [
            r["review_id"], r["student_id"], r["student_name"], r["student_email"],
            r["task_id"], r["review_number"], r["review_date"], r["review_time"],
            r["time_zone"], r["meeting_gcr_link"], r["submission_status"], r["review_status"],
            r["reviewer"], r["feedback_status"], r["next_review_date"]
        ]
        ws_rev.append(row)

    # ----------------------------------------------------
    # Sheet 4: AI Review Feedback
    # ----------------------------------------------------
    ws_fb = wb.create_sheet(title="AI Review Feedback")
    fb_headers = [
        "Student ID", "Task ID", "Review Number", "Submission Summary", "Evaluation Criteria",
        "Identified Mistakes", "Missing Requirements", "Strengths", "Improvement Suggestions",
        "AI-Generated Feedback", "Score", "Reviewer Approval Status", "Next Actions"
    ]
    ws_fb.append(fb_headers)

    for fb in feedbacks:
        row = [
            fb["student_id"], fb["task_id"], fb["review_number"], fb["submission_summary"],
            fb["evaluation_criteria"], fb["identified_mistakes"], fb["missing_requirements"],
            fb["strengths"], fb["improvement_suggestions"], fb["ai_generated_feedback"],
            fb["score"], fb["reviewer_approval_status"], fb["next_actions"]
        ]
        ws_fb.append(row)

    # ----------------------------------------------------
    # Sheet 5: Email Log
    # ----------------------------------------------------
    ws_email = wb.create_sheet(title="Email Log")
    email_headers = [
        "Email ID", "Student ID", "Recipient Email", "Email Type", "Subject",
        "Sent Date and Time", "Delivery Status", "Error Message", "Retry Count", "Related Review ID"
    ]
    ws_email.append(email_headers)

    for em in email_logs:
        row = [
            em["email_id"], em["student_id"], em["recipient_email"], em["email_type"],
            em["subject"], em["sent_date_time"], em["delivery_status"], em["error_message"],
            em["retry_count"], em["related_review_id"]
        ]
        ws_email.append(row)

    # ----------------------------------------------------
    # Sheet 6: Dashboard (with Formulas & Executive Cards)
    # ----------------------------------------------------
    ws_dash = wb.create_sheet(title="Dashboard")
    ws_dash.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_dash.merge_cells("A1:H2")
    title_cell = ws_dash["A1"]
    title_cell.value = "AI-POWERED INTERNSHIP AUTOMATION EXECUTIVE DASHBOARD"
    title_cell.font = Font(name=FONT_FAMILY, size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    subtitle_cell = ws_dash["A3"]
    subtitle_cell.value = f"Real-Time Metrics & Automated Audit Summary | Synchronized: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    subtitle_cell.font = Font(name=FONT_FAMILY, size=10, italic=True, color="64748B")

    # KPI Card Layout
    # Cards: [Label, Formula, ColStart, ColEnd, RowStart, RowEnd, ThemeColor]
    kpis = [
        ("Total Students", '=COUNTA(\'Student Master\'!A2:A500)', "A", "B", 5, 6, "1E3A8A"),
        ("Active Interns", '=COUNTIF(\'Student Master\'!W2:W500, "Active") + COUNTIF(\'Student Master\'!W2:W500, "In Progress")', "C", "D", 5, 6, "0D9488"),
        ("Tasks Assigned", '=COUNTIF(\'Student Master\'!I2:I500, "<>")', "E", "F", 5, 6, "2563EB"),
        ("Tasks Completed", '=COUNTIF(\'Student Master\'!N2:N500, "Completed") + COUNTIF(\'Student Master\'!N2:N500, "Approved")', "G", "H", 5, 6, "16A34A"),
        
        ("Pending Tasks", '=COUNTIF(\'Student Master\'!N2:N500, "In Progress") + COUNTIF(\'Student Master\'!N2:N500, "Assigned") + COUNTIF(\'Student Master\'!N2:N500, "Resubmission Pending")', "A", "B", 8, 9, "D97706"),
        ("Overdue Tasks", '=COUNTIF(\'Student Master\'!N2:N500, "Overdue")', "C", "D", 8, 9, "DC2626"),
        ("Reviews Scheduled", '=COUNTIF(\'Review Schedule\'!L2:L500, "Scheduled")', "E", "F", 8, 9, "7C3AED"),
        ("Reviews Completed", '=COUNTIF(\'Review Schedule\'!L2:L500, "Completed")', "G", "H", 8, 9, "059669"),

        ("Reviews Due Today", f'=COUNTIF(\'Review Schedule\'!G2:G500, "{datetime.now().strftime("%Y-%m-%d")}")', "A", "B", 11, 12, "EA580C"),
        ("Emails Sent", '=COUNTIF(\'Email Log\'!G2:G500, "Sent")', "C", "D", 11, 12, "0284C7"),
        ("Emails Failed", '=COUNTIF(\'Email Log\'!G2:G500, "Failed")', "E", "F", 11, 12, "B91C1C"),
        ("Awaiting Feedback", '=COUNTIF(\'Student Master\'!N2:N500, "Under Review") + COUNTIF(\'Student Master\'!N2:N500, "Submitted")', "G", "H", 11, 12, "9333EA"),
    ]

    for label, formula, c1, c2, r1, r2, col_hex in kpis:
        # Merge Header
        ws_dash.merge_cells(f"{c1}{r1}:{c2}{r1}")
        lbl_cell = ws_dash[f"{c1}{r1}"]
        lbl_cell.value = label.upper()
        lbl_cell.font = Font(name=FONT_FAMILY, size=9, bold=True, color="64748B")
        lbl_cell.alignment = Alignment(horizontal="center", vertical="center")
        lbl_cell.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

        # Merge Value
        ws_dash.merge_cells(f"{c1}{r2}:{c2}{r2}")
        val_cell = ws_dash[f"{c1}{r2}"]
        val_cell.value = formula
        val_cell.font = Font(name=FONT_FAMILY, size=16, bold=True, color=col_hex)
        val_cell.alignment = Alignment(horizontal="center", vertical="center")
        val_cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        # Borders
        for row_idx in range(r1, r2 + 1):
            for col_letter in [c1, c2]:
                ws_dash[f"{col_letter}{row_idx}"].border = CELL_BORDER

    # Summary Breakdown Table in Dashboard
    ws_dash["A15"].value = "INTERNSHIP STATUS BREAKDOWN"
    ws_dash["A15"].font = Font(name=FONT_FAMILY, size=11, bold=True, color=COLOR_HEADER_NAVY)
    
    ws_dash["A16"].value = "Status"
    ws_dash["B16"].value = "Student Count"
    ws_dash["A16"].font = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
    ws_dash["B16"].font = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
    ws_dash["A16"].fill = HEADER_FILL
    ws_dash["B16"].fill = HEADER_FILL

    dash_statuses = ["Completed", "Active", "In Progress", "Not Started"]
    for idx, st in enumerate(dash_statuses, start=17):
        ws_dash[f"A{idx}"].value = st
        ws_dash[f"B{idx}"].value = f'=COUNTIF(\'Student Master\'!W2:W500, "{st}")'
        ws_dash[f"A{idx}"].font = REGULAR_FONT
        ws_dash[f"B{idx}"].font = BOLD_FONT
        ws_dash[f"A{idx}"].border = CELL_BORDER
        ws_dash[f"B{idx}"].border = CELL_BORDER
        ws_dash[f"B{idx}"].alignment = Alignment(horizontal="right")

    # Add Bar Chart for Status Breakdown
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Students by Internship Status"
    chart.y_axis.title = "Number of Students"
    chart.x_axis.title = "Status"
    data_ref = Reference(ws_dash, min_col=2, min_row=16, max_row=16 + len(dash_statuses))
    cats_ref = Reference(ws_dash, min_col=1, min_row=17, max_row=16 + len(dash_statuses))
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.legend = None
    chart.width = 16
    chart.height = 8
    ws_dash.add_chart(chart, "D15")

    # ----------------------------------------------------
    # Apply Professional Styling, Auto Widths, Dropdowns
    # ----------------------------------------------------
    # Data Validation lists
    dv_task_status = DataValidation(
        type="list",
        formula1='"Not Started,Assigned,In Progress,Submitted,Under Review,Corrections Required,Resubmission Pending,Approved,Completed,Overdue,On Hold"',
        allow_blank=True
    )
    dv_intern_status = DataValidation(
        type="list",
        formula1='"Not Started,Active,In Progress,Nearing Completion,Completed,Extended,On Hold,Terminated"',
        allow_blank=True
    )
    dv_priority = DataValidation(
        type="list",
        formula1='"High,Medium,Low"',
        allow_blank=True
    )
    dv_rev_status = DataValidation(
        type="list",
        formula1='"Scheduled,Reminder Sent,Submission Received,AI Evaluated,Awaiting Approval,Approved,Completed,Missed,Rescheduled"',
        allow_blank=True
    )

    ws_student.add_data_validation(dv_task_status)
    ws_student.add_data_validation(dv_intern_status)
    dv_task_status.add("N2:N500")
    dv_intern_status.add("W2:W500")

    ws_alloc.add_data_validation(dv_priority)
    dv_priority.add("I2:I500")

    ws_rev.add_data_validation(dv_rev_status)
    dv_rev_status.add("L2:L500")

    # Style all data sheets
    data_sheets = [ws_student, ws_alloc, ws_rev, ws_fb, ws_email]
    for ws in data_sheets:
        # Style Header Row
        for cell in ws[1]:
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = CELL_BORDER
        ws.row_dimensions[1].height = 26

        # Zebra striping & cell styling
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            is_even = (row_idx % 2 == 0)
            fill_color = COLOR_ZEBRA_LIGHT if is_even else "FFFFFF"
            row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
            for cell in row:
                cell.font = REGULAR_FONT
                cell.fill = row_fill
                cell.border = CELL_BORDER
                cell.alignment = Alignment(vertical="center")

        # Auto-adjust column widths
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = 0
            for cell in col:
                val_str = str(cell.value or "")
                # Skip long multiline descriptions for width calculation to keep sheets readable
                first_line = val_str.split("\n")[0] if "\n" in val_str else val_str
                if len(first_line) > max_len:
                    max_len = len(first_line)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

        # Freeze header row
        ws.freeze_panes = "A2"

    # Save workbook
    wb.save(str(output_path))
    return output_path


def sync_excel_to_sqlite(excel_file: Path = EXCEL_PATH):
    """
    Reads the Excel workbook and updates the SQLite database with user edits.
    Guarantees synchronization when managers make changes directly in Excel.
    """
    if not excel_file.exists():
        raise FileNotFoundError(f"Workbook {excel_file} not found.")

    wb = openpyxl.load_workbook(str(excel_file), data_only=True)
    with database.get_connection() as conn:
        cursor = conn.cursor()

        # 1. Sync Student Master
        if "Student Master" in wb.sheetnames:
            ws = wb["Student Master"]
            headers = [cell.value for cell in ws[1]]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or not row[0]:
                    continue
                row_dict = dict(zip(headers, row))
                cursor.execute("""
                    INSERT INTO students (
                        student_id, student_name, student_email, internship_start_date, internship_end_date,
                        internship_role, mentor_name, mentor_email, assigned_task_id, assigned_task_title,
                        task_description, task_allocation_date, task_deadline, task_status,
                        current_review_number, last_review_date, next_review_date, next_review_time,
                        gcr_link, task_submission_link, latest_ai_feedback, completion_percentage,
                        internship_status, email_delivery_status, last_email_sent, remarks, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                    ON CONFLICT(student_id) DO UPDATE SET
                        student_name=excluded.student_name,
                        student_email=excluded.student_email,
                        internship_start_date=excluded.internship_start_date,
                        internship_end_date=excluded.internship_end_date,
                        internship_role=excluded.internship_role,
                        mentor_name=excluded.mentor_name,
                        mentor_email=excluded.mentor_email,
                        assigned_task_id=excluded.assigned_task_id,
                        assigned_task_title=excluded.assigned_task_title,
                        task_description=excluded.task_description,
                        task_allocation_date=excluded.task_allocation_date,
                        task_deadline=excluded.task_deadline,
                        task_status=excluded.task_status,
                        current_review_number=excluded.current_review_number,
                        last_review_date=excluded.last_review_date,
                        next_review_date=excluded.next_review_date,
                        next_review_time=excluded.next_review_time,
                        gcr_link=excluded.gcr_link,
                        task_submission_link=excluded.task_submission_link,
                        latest_ai_feedback=excluded.latest_ai_feedback,
                        completion_percentage=excluded.completion_percentage,
                        internship_status=excluded.internship_status,
                        email_delivery_status=excluded.email_delivery_status,
                        last_email_sent=excluded.last_email_sent,
                        remarks=excluded.remarks,
                        updated_at=CURRENT_TIMESTAMP
                """, (
                    str(row_dict.get("Student ID", "")),
                    str(row_dict.get("Student Name", "")),
                    str(row_dict.get("Student Email", "")),
                    str(row_dict.get("Internship Start Date", "") or ""),
                    str(row_dict.get("Internship End Date", "") or ""),
                    str(row_dict.get("Internship Role", "")),
                    str(row_dict.get("Mentor Name", "")),
                    str(row_dict.get("Mentor Email", "")),
                    str(row_dict.get("Assigned Task ID", "") or ""),
                    str(row_dict.get("Assigned Task Title", "") or ""),
                    str(row_dict.get("Task Description", "") or ""),
                    str(row_dict.get("Task Allocation Date", "") or ""),
                    str(row_dict.get("Task Deadline", "") or ""),
                    str(row_dict.get("Task Status", "Not Started")),
                    int(row_dict.get("Current Review Number") or 0),
                    str(row_dict.get("Last Review Date", "") or ""),
                    str(row_dict.get("Next Review Date", "") or ""),
                    str(row_dict.get("Next Review Time", "") or ""),
                    str(row_dict.get("GCR / Google Classroom Link", "") or ""),
                    str(row_dict.get("Task Submission Link", "") or ""),
                    str(row_dict.get("Latest AI Feedback", "") or ""),
                    float(row_dict.get("Completion Percentage") or 0.0),
                    str(row_dict.get("Internship Status", "Active")),
                    str(row_dict.get("Email Delivery Status", "Pending")),
                    str(row_dict.get("Last Email Sent", "") or ""),
                    str(row_dict.get("Remarks", "") or "")
                ))

        # 2. Sync Review Schedule
        if "Review Schedule" in wb.sheetnames:
            ws = wb["Review Schedule"]
            headers = [cell.value for cell in ws[1]]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or not row[0]:
                    continue
                row_dict = dict(zip(headers, row))
                cursor.execute("""
                    INSERT INTO review_schedules (
                        review_id, student_id, student_name, student_email, task_id,
                        review_number, review_date, review_time, time_zone, meeting_gcr_link,
                        submission_status, review_status, reviewer, feedback_status, next_review_date
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(review_id) DO UPDATE SET
                        review_date=excluded.review_date,
                        review_time=excluded.review_time,
                        submission_status=excluded.submission_status,
                        review_status=excluded.review_status,
                        reviewer=excluded.reviewer,
                        feedback_status=excluded.feedback_status,
                        next_review_date=excluded.next_review_date
                """, (
                    str(row_dict.get("Review ID", "")),
                    str(row_dict.get("Student ID", "")),
                    str(row_dict.get("Student Name", "")),
                    str(row_dict.get("Student Email", "")),
                    str(row_dict.get("Task ID", "")),
                    int(row_dict.get("Review Number") or 1),
                    str(row_dict.get("Review Date", "")),
                    str(row_dict.get("Review Time", "14:00")),
                    str(row_dict.get("Time Zone", "Asia/Kolkata")),
                    str(row_dict.get("Meeting / GCR Link", "")),
                    str(row_dict.get("Submission Status", "Pending")),
                    str(row_dict.get("Review Status", "Scheduled")),
                    str(row_dict.get("Reviewer", "")),
                    str(row_dict.get("Feedback Status", "Pending")),
                    str(row_dict.get("Next Review Date", "") or "")
                ))

        conn.commit()


def export_sqlite_to_excel(output_path: Path = EXCEL_PATH):
    """
    Exports latest SQLite state to the formatted Excel workbook,
    safely creating a timestamped backup of the prior version first.
    """
    if output_path.exists():
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = BACKUP_DIR / f"AI_Internship_Review_Management_{timestamp}.bak.xlsx"
        shutil.copy2(output_path, backup_file)

    students = database.get_all_students()
    allocations = database.get_all_allocations()
    reviews = database.get_all_reviews()
    feedbacks = database.get_all_feedbacks()
    email_logs = database.get_all_email_logs()

    create_styled_workbook(students, allocations, reviews, feedbacks, email_logs, output_path)
