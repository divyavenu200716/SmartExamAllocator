import streamlit as st
import pandas as pd
import io
import os
from allocator import process_allocation, get_student_classes, get_active_classes_from_timetable

st.set_page_config(page_title="KASC SMART SEAT", page_icon="logo.png", layout="centered")

# Hide Streamlit Watermark, Menu, Footer, and Creator Avatar
hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
a[href*="streamlit"] {display: none !important;}
[data-testid="stHeader"] {display: none !important;}
div[class^="viewerBadge"] {display: none !important;}
div[class^="creator"] {display: none !important;}
img[alt*="Creator"] {display: none !important;}
[title*="Creator"] {display: none !important;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

st.title("🎓 KASC SMART SEAT")
st.markdown("### 🪑 Automated Seating Arrangement Generator")
st.info("👈 Please use the sidebar on the left to upload your files and configure the exam settings.")

# --- SIDEBAR ---
st.sidebar.title("📁 Configuration")

with st.sidebar.expander("📂 1. Hall Details (Excel/PDF)", expanded=False):
    hall_file = st.file_uploader("Select Hall File", type=['xlsx', 'xls', 'csv', 'pdf'], label_visibility="collapsed")
    st.markdown("*(Excel: HALL NO, TOTAL SEAT, ROWS, COLUMNS)*")

with st.sidebar.expander("📂 2. Student Details (Excel/PDF)", expanded=False):
    student_files = st.file_uploader("Select Student Files", type=['xlsx', 'xls', 'csv', 'pdf'], accept_multiple_files=True, label_visibility="collapsed")
    st.markdown("*(Upload one or multiple files)*")

with st.sidebar.expander("📂 3. Time Table (Excel/PDF)", expanded=False):
    timetable_file = st.file_uploader("Select Time Table", type=['pdf', 'xlsx', 'xls', 'csv'], label_visibility="collapsed")

with st.sidebar.expander("⚙️ 4. Exam Settings", expanded=False):
    st.markdown("*(Required for auto-detection)*")
    exam_title = st.text_input("Exam Title", value="PERIYAR UNIVERSITY THEORY EXAMINATIONS - APR/MAY")
    exam_date = st.text_input("Exam Date (e.g., 28-04-2025)", value="")
    exam_session = st.selectbox("Session", ["FN", "AN"])

selected_classes = []
if student_files:
    try:
        available_classes = get_student_classes(student_files)
        
        # Try to auto-detect from PDF
        auto_detected = []
        if timetable_file and exam_date and exam_session:
            detected = get_active_classes_from_timetable(timetable_file, exam_date, exam_session)
            if not detected:
                st.sidebar.warning("⚠️ Auto-detection couldn't find classes.")
            # Fuzzy match detected classes against available classes
            auto_set = set()
            for d in detected:
                d_clean = d.replace(' ', '').upper()
                for a in available_classes:
                    a_clean = a.replace(' ', '').upper()
                    if d_clean == a_clean or d_clean in a_clean or a_clean in d_clean:
                        auto_set.add(a)
            auto_detected = list(auto_set)
            
        with st.sidebar.expander("🎓 5. Classes taking the Exam", expanded=False):
            if auto_detected:
                st.success("✨ Classes auto-detected from Time Table!")
            else:
                st.info("💡 Enter Date & Session to auto-detect classes")
                
            selected_classes = st.multiselect("Confirm Classes", available_classes, default=auto_detected, label_visibility="collapsed")
    except Exception as e:
        st.sidebar.error(f"Could not parse student details: {e}")

# --- MAIN PAGE ---
st.markdown("---")
st.markdown("### 🚀 Final Step")
if st.button("✨ Generate Seating Arrangement", use_container_width=True):
    if hall_file and student_files and timetable_file and exam_date and selected_classes:
        with st.spinner("Generating seating arrangement..."):
            try:
                output_excel_bytes = process_allocation(hall_file, student_files, timetable_file, exam_date, exam_session, selected_classes, exam_title)
                
                st.success("🎉 Seating Arrangement Generated Successfully!")
                
                st.download_button(
                    label="⬇️ Download Seating Arrangement",
                    data=output_excel_bytes,
                    file_name="Generated_Seating_Arrangement.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            except Exception as e:
                import traceback
                st.error(f"❌ An error occurred during generation: {e}")
                st.code(traceback.format_exc())
    else:
        st.warning("⚠️ Please upload all files, fill Exam Date in the sidebar, and confirm at least one class.")
