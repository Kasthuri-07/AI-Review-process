"""
AI-Powered Internship Management & Review Automation Dashboard
Streamlit Executive Interface for Task Allocation, AI Evaluation, Review Scheduling, and Audit Logs.
"""

from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st

from config import EXCEL_PATH, ROLES_TRACKS, SUBMISSIONS_DIR
import database
from task_allocator import TaskAllocator
from scheduler_service import ReviewSchedulerService
from ai_evaluator import AIEvaluator
from workflow_engine import WorkflowEngine
import excel_manager

# Page Configuration & Styling
st.set_page_config(
    page_title="AI Internship Automation & Review System",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for modern visual excellence
st.markdown("""
<style>
    .main {
        background-color: #f8fafc;
    }
    .metric-card {
        background: white;
        border-radius: 10px;
        padding: 16px 20px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        border: 1px solid #e2e8f0;
        margin-bottom: 12px;
    }
    .metric-title {
        color: #64748b;
        font-size: 11px;
        text-transform: uppercase;
        font-weight: 700;
        letter-spacing: 0.05em;
    }
    .metric-value {
        font-size: 26px;
        font-weight: 800;
        color: #0f172a;
        margin-top: 4px;
    }
    .badge-approved { background-color: #dcfce7; color: #166534; padding: 3px 8px; border-radius: 4px; font-weight: 700; font-size: 11px; }
    .badge-pending { background-color: #fef3c7; color: #92400e; padding: 3px 8px; border-radius: 4px; font-weight: 700; font-size: 11px; }
    .badge-danger { background-color: #fee2e2; color: #991b1b; padding: 3px 8px; border-radius: 4px; font-weight: 700; font-size: 11px; }
    .badge-info { background-color: #dbeafe; color: #1e40af; padding: 3px 8px; border-radius: 4px; font-weight: 700; font-size: 11px; }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 18px;
        font-weight: 600;
        border-radius: 6px;
    }
</style>
""", unsafe_allow_html=True)

# Initialize engines
allocator = TaskAllocator()
scheduler = ReviewSchedulerService()
evaluator = AIEvaluator()
workflow = WorkflowEngine()

# Sidebar Navigation & Quick Actions
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135755.png", width=60)
    st.title("Internship Ops")
    st.caption("AI-Powered Automation Engine v2.4")
    
    st.markdown("---")
    st.markdown("### ⚡ Quick Sync")
    if st.button("🔄 Sync with Excel File", use_container_width=True):
        try:
            excel_manager.sync_excel_to_sqlite()
            st.success("Successfully synchronized SQLite with Excel workbook!")
        except Exception as e:
            st.error(f"Sync error: {e}")

    if st.button("📤 Export SQLite to Excel", use_container_width=True):
        try:
            excel_manager.export_sqlite_to_excel()
            st.success("Workbook exported and backed up!")
        except Exception as e:
            st.error(f"Export error: {e}")

    st.markdown("---")
    st.markdown("### 🔔 Automated Reminders")
    if st.button("Trigger Reminder Heartbeat", use_container_width=True):
        res = scheduler.process_due_reminders()
        st.info(f"Scan complete: 24h: {res['reminders_24h_sent']}, 1h: {res['reminders_1h_sent']}, Overdue: {res['overdue_notices_sent']}")

    st.markdown("---")
    st.caption("📁 Database: SQLite (ACID)")
    st.caption(f"📊 Excel: {EXCEL_PATH.name}")

# Main Tabs
tabs = st.tabs([
    "📊 Executive Dashboard",
    "👥 Student Master (100)",
    "🎯 Task Allocation Hub",
    "📅 Review & Reminders",
    "🤖 AI Evaluation Studio",
    "📨 Email Outbox & Logs",
    "📜 System Audit Trail"
])

# ----------------------------------------------------
# TAB 1: EXECUTIVE DASHBOARD
# ----------------------------------------------------
with tabs[0]:
    st.header("Executive Operations & Performance Dashboard")
    st.caption("Real-time telemetry tracking task completion, review velocity, and email delivery.")

    metrics = database.get_dashboard_metrics()

    # Row 1: KPI Metrics
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Total Registered Students</div>
            <div class="metric-value">{metrics['total_students']}</div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Active Interns</div>
            <div class="metric-value" style="color: #0284c7;">{metrics['active_interns']}</div>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Tasks Completed / Approved</div>
            <div class="metric-value" style="color: #16a34a;">{metrics['tasks_completed']}</div>
        </div>
        """, unsafe_allow_html=True)
    with c4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Internships Completed</div>
            <div class="metric-value" style="color: #7c3aed;">{metrics['students_completed_internship']}</div>
        </div>
        """, unsafe_allow_html=True)

    # Row 2: Operational Health
    c5, c6, c7, c8 = st.columns(4)
    with c5:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Pending / In-Progress Tasks</div>
            <div class="metric-value" style="color: #d97706;">{metrics['pending_tasks']}</div>
        </div>
        """, unsafe_allow_html=True)
    with c6:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Overdue Tasks</div>
            <div class="metric-value" style="color: #dc2626;">{metrics['overdue_tasks']}</div>
        </div>
        """, unsafe_allow_html=True)
    with c7:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Reviews Scheduled</div>
            <div class="metric-value" style="color: #2563eb;">{metrics['reviews_scheduled']}</div>
        </div>
        """, unsafe_allow_html=True)
    with c8:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Students Awaiting Feedback</div>
            <div class="metric-value" style="color: #9333ea;">{metrics['students_awaiting_feedback']}</div>
        </div>
        """, unsafe_allow_html=True)

    # Charts Row
    st.markdown("### 📈 Visual Analytics & Pipeline Health")
    ch_col1, ch_col2 = st.columns(2)

    students_df = pd.DataFrame(database.get_all_students())
    if not students_df.empty:
        with ch_col1:
            st.subheader("Task Status Distribution")
            task_counts = students_df["task_status"].value_counts().reset_index()
            task_counts.columns = ["Status", "Count"]
            st.bar_chart(data=task_counts, x="Status", y="Count")

        with ch_col2:
            st.subheader("Students by Internship Track")
            role_counts = students_df["internship_role"].value_counts().reset_index()
            role_counts.columns = ["Track", "Students"]
            st.bar_chart(data=role_counts, x="Track", y="Students")

# ----------------------------------------------------
# TAB 2: STUDENT MASTER
# ----------------------------------------------------
with tabs[1]:
    st.header("Student Master Directory")
    st.caption("Complete repository of 100 student records with granular tracking and contact links.")

    students_data = database.get_all_students()
    df_students = pd.DataFrame(students_data)

    if not df_students.empty:
        # Filters
        f1, f2, f3 = st.columns(3)
        with f1:
            role_filter = st.selectbox("Filter by Track/Role", ["All"] + ROLES_TRACKS)
        with f2:
            status_filter = st.selectbox("Filter by Task Status", ["All"] + list(df_students["task_status"].unique()))
        with f3:
            search_query = st.text_input("Search by Name or ID", "")

        filtered_df = df_students.copy()
        if role_filter != "All":
            filtered_df = filtered_df[filtered_df["internship_role"] == role_filter]
        if status_filter != "All":
            filtered_df = filtered_df[filtered_df["task_status"] == status_filter]
        if search_query:
            filtered_df = filtered_df[
                filtered_df["student_id"].str.contains(search_query, case=False, na=False) |
                filtered_df["student_name"].str.contains(search_query, case=False, na=False)
            ]

        # Display interactive table
        display_cols = [
            "student_id", "student_name", "student_email", "internship_role",
            "mentor_name", "assigned_task_id", "task_status", "current_review_number",
            "next_review_date", "completion_percentage", "internship_status"
        ]
        st.dataframe(filtered_df[display_cols], use_container_width=True, height=450)
        st.caption(f"Showing {len(filtered_df)} of {len(df_students)} students.")

# ----------------------------------------------------
# TAB 3: TASK ALLOCATION HUB
# ----------------------------------------------------
with tabs[2]:
    st.header("Task Allocation Automation")
    st.caption("Automatic track matching, load balancing, and manual task assignment with duplicate prevention.")

    unassigned = allocator.get_unassigned_students()
    
    col_u1, col_u2 = st.columns([2, 1])
    with col_u1:
        st.subheader(f"Unassigned Students Queue ({len(unassigned)})")
        if unassigned:
            u_df = pd.DataFrame(unassigned)[["student_id", "student_name", "student_email", "internship_role"]]
            st.dataframe(u_df, use_container_width=True)
            if st.button("🚀 Auto-Allocate All Pending Students", type="primary"):
                with st.spinner("Assigning tasks and dispatching personalized allocation emails..."):
                    res = allocator.auto_allocate_all_pending()
                    st.success(f"Allocated {res['success']} students successfully! ({res['failed']} failed)")
                    st.rerun()
        else:
            st.success("✅ All 100 students have an active task allocated!")

    with col_u2:
        st.subheader("Manual Task Assignment")
        with st.form("manual_alloc_form"):
            all_stu = database.get_all_students()
            stu_options = [f"{s['student_id']} - {s['student_name']} ({s['internship_role']})" for s in all_stu]
            selected_stu_str = st.selectbox("Select Student", stu_options)
            
            tasks = database.get_all_tasks()
            task_options = [f"{t['task_id']} - {t['task_title']}" for t in tasks]
            selected_task_str = st.selectbox("Select Task from Catalog", task_options)
            force_resend = st.checkbox("Force Resend Allocation Email", value=False)
            
            submit_btn = st.form_submit_button("Assign Task & Send Email")
            if submit_btn:
                target_stu_id = selected_stu_str.split(" - ")[0]
                target_task_id = selected_task_str.split(" - ")[0]
                ok, msg, _ = allocator.allocate_task_to_student(
                    student_id=target_stu_id,
                    task_id=target_task_id,
                    force_resend_email=force_resend
                )
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

    st.markdown("---")
    st.subheader("📚 Task Catalog & Rubrics")
    tasks_df = pd.DataFrame(database.get_all_tasks())
    if not tasks_df.empty:
        st.dataframe(tasks_df[["task_id", "track_role", "task_title", "task_priority", "estimated_days", "evaluation_rubric"]], use_container_width=True)

# ----------------------------------------------------
# TAB 4: REVIEW & REMINDERS
# ----------------------------------------------------
with tabs[3]:
    st.header("Review Scheduling & Milestone Tracking")
    st.caption("Manage review calendar, inspect reminders, and handle missed sessions.")

    r_col1, r_col2 = st.columns([2, 1])
    
    with r_col1:
        st.subheader("Scheduled & Past Reviews")
        reviews_data = database.get_all_reviews()
        if reviews_data:
            rev_df = pd.DataFrame(reviews_data)
            display_rev_cols = [
                "review_id", "student_id", "student_name", "task_id", "review_number",
                "review_date", "review_time", "review_status", "submission_status", "reviewer"
            ]
            st.dataframe(rev_df[display_rev_cols], use_container_width=True, height=400)
        else:
            st.info("No reviews currently on schedule.")

    with r_col2:
        st.subheader("Schedule New Review Session")
        with st.form("schedule_review_form"):
            students = database.get_all_students()
            stu_choices = [f"{s['student_id']} - {s['student_name']}" for s in students]
            chosen_stu = st.selectbox("Select Student", stu_choices)
            
            review_date = st.date_input("Review Date", min_value=datetime.now().date())
            review_time = st.time_input("Review Time", value=datetime.strptime("14:00", "%H:%M").time())
            custom_meeting_link = st.text_input("Meeting / GCR Link (optional)", "")
            custom_reviewer = st.text_input("Reviewer Name (optional)", "")

            sched_btn = st.form_submit_button("Schedule & Dispatch Confirmation")
            if sched_btn:
                target_stu_id = chosen_stu.split(" - ")[0]
                ok, msg, _ = scheduler.schedule_review(
                    student_id=target_stu_id,
                    review_date_str=review_date.strftime("%Y-%m-%d"),
                    review_time_str=review_time.strftime("%H:%M"),
                    meeting_link=custom_meeting_link or None,
                    reviewer_name=custom_reviewer or None
                )
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

    st.markdown("---")
    st.subheader("⚠️ Handle Missed Review")
    with st.expander("Record Missed Review & Trigger Reschedule"):
        m_rev_id = st.text_input("Enter Review ID to mark as Missed (e.g. REV-STU-035-1)")
        m_reason = st.text_input("Reason / Notes", "Student did not attend scheduled call.")
        if st.button("Mark Missed & Auto-Reschedule"):
            if m_rev_id:
                ok, msg = workflow.record_missed_review(m_rev_id, m_reason)
                if ok:
                    st.warning(msg)
                    st.rerun()
                else:
                    st.error(msg)

# ----------------------------------------------------
# TAB 5: AI EVALUATION STUDIO
# ----------------------------------------------------
with tabs[4]:
    st.header("AI Review & Feedback Studio")
    st.caption("AI-powered submission evaluation against rubrics with mandatory human mentor sign-off.")

    stu_list = database.get_all_students()
    stu_opts = [f"{s['student_id']} - {s['student_name']} ({s.get('assigned_task_id', 'No Task')})" for s in stu_list]
    selected_eval_stu = st.selectbox("Select Student for Submission Evaluation", stu_opts)
    eval_stu_id = selected_eval_stu.split(" - ")[0]
    eval_student = database.get_student(eval_stu_id)

    e_col1, e_col2 = st.columns([1, 1])

    with e_col1:
        st.subheader("Submission Input & Code Inspector")
        # Check if local sample file exists
        local_samples = list(SUBMISSIONS_DIR.glob(f"{eval_stu_id}*"))
        default_code = ""
        if local_samples:
            with open(local_samples[0], "r", encoding="utf-8") as f:
                default_code = f.read()
            st.info(f"Loaded existing submission file: `{local_samples[0].name}`")
        else:
            default_code = f"# Upload or paste code for {eval_stu_id}\n# Task: {eval_student.get('assigned_task_id')}\n"

        submission_code = st.text_area("Submission Code / Content", value=default_code, height=280)
        uploaded_file = st.file_uploader("Or Upload Project File (.py, .txt, .json)", type=["py", "txt", "json", "md"])

        if uploaded_file is not None:
            submission_code = uploaded_file.read().decode("utf-8", errors="ignore")
            st.success(f"Loaded uploaded file: {uploaded_file.name}")

        run_ai_btn = st.button("🤖 Run AI Rubric Evaluation", type="primary")

    with e_col2:
        st.subheader("Evaluation Results & Mentor Approval")
        
        if run_ai_btn:
            with st.spinner("Analyzing code against deliverables and grading rubric..."):
                feedback_draft = evaluator.evaluate_submission(
                    student_id=eval_stu_id,
                    task_id=eval_student.get("assigned_task_id") or "TASK-AIML-101",
                    submission_text=submission_code
                )
                st.session_state["latest_eval"] = feedback_draft
                st.success(f"AI Evaluation Completed! Draft Score: {feedback_draft['score']}/100")

        if "latest_eval" in st.session_state:
            fb = st.session_state["latest_eval"]

            st.markdown(f"**Score:** `{fb['score']}/100` | **Status:** `{fb['reviewer_approval_status']}`")
            st.markdown(f"**Summary:** {fb['submission_summary']}")
            
            with st.expander("✅ Strengths (What Was Done Correctly)", expanded=True):
                st.write(fb['what_was_done_correctly'])

            with st.expander("❌ Mistakes & Technical Issues", expanded=True):
                st.write(fb['identified_mistakes'])

            with st.expander("⚠️ Missing Requirements", expanded=True):
                st.write(fb['missing_requirements'])

            with st.expander("💡 Recommended Corrections & Rubric Notes", expanded=True):
                st.write(fb['recommended_corrections'])

            st.markdown("---")
            st.markdown("### ✍️ Human Reviewer Sign-Off & Dispatch")
            mentor_name = st.text_input("Reviewer Name", value=eval_student.get("mentor_name", "Lead Mentor"))
            mentor_score_adj = st.slider("Approve / Adjust Score", 0, 100, int(fb['score']))
            mentor_notes = st.text_area("Additional Mentor Feedback to Student", "Great effort. Please review the detailed comments above.")

            if st.button("✅ Approve & Dispatch Feedback to Student", type="primary"):
                ok, msg, res = workflow.approve_feedback_and_advance(
                    feedback_id=fb["feedback_id"],
                    reviewer_name=mentor_name,
                    reviewer_notes=mentor_notes,
                    score_override=mentor_score_adj
                )
                if ok:
                    st.balloons()
                    st.success(f"Feedback approved and dispatched! Result: {msg}")
                    del st.session_state["latest_eval"]
                    st.rerun()
                else:
                    st.error(msg)

# ----------------------------------------------------
# TAB 6: EMAIL OUTBOX & LOGS
# ----------------------------------------------------
with tabs[5]:
    st.header("Email Delivery & Outbox Inspection")
    st.caption("Inspect live delivery logs, preview rendered HTML notifications, and monitor retry counters.")

    logs = database.get_all_email_logs()
    if logs:
        df_logs = pd.DataFrame(logs)
        st.dataframe(
            df_logs[["email_id", "student_id", "recipient_email", "email_type", "subject", "sent_date_time", "delivery_status", "retry_count"]],
            use_container_width=True,
            height=300
        )

        st.markdown("---")
        st.subheader("🔍 HTML Email Previewer")
        selected_log_id = st.selectbox("Select Email to Preview", df_logs["email_id"].tolist())
        selected_log = next((l for l in logs if l["email_id"] == selected_log_id), None)
        
        if selected_log and selected_log.get("outbox_filepath") and Path(selected_log["outbox_filepath"]).exists():
            with open(selected_log["outbox_filepath"], "r", encoding="utf-8") as f:
                html_body = f.read()
            st.components.v1.html(html_body, height=500, scrolling=True)
        else:
            st.info("Direct HTML preview available for emails generated in MOCK mode.")
    else:
        st.info("No emails logged yet.")

# ----------------------------------------------------
# TAB 7: SYSTEM AUDIT TRAIL
# ----------------------------------------------------
with tabs[6]:
    st.header("System Audit Trail & Security Logs")
    st.caption("Immutable append-only transaction logs tracking allocations, reviews, feedback, and edits.")

    with database.get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM audit_logs ORDER BY audit_id DESC LIMIT 100")
        audit_rows = [dict(r) for r in cursor.fetchall()]

    if audit_rows:
        df_audit = pd.DataFrame(audit_rows)
        st.dataframe(df_audit, use_container_width=True, height=450)
    else:
        st.info("No audit logs recorded.")
