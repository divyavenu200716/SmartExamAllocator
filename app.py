import streamlit as st
import pandas as pd
import io
import os
from allocator import process_allocation, get_student_classes, get_active_classes_from_timetable

st.set_page_config(page_title="SMART SEAT", page_icon="logo.png", layout="wide")

st.title("🎓 SMART SEAT")

st.markdown("Upload your exam details below to automatically generate the seating arrangement.")

with st.expander("ℹ️ How to format your Excel files? (Click to view)"):
    st.markdown("""
    **1. Hall Details (Excel):**
    Must have columns named exactly **`HALL NO`** and **`TOTAL SEAT`**. 
    *(Example: HALL NO: 101, TOTAL SEAT: 30)*
    
    **2. Student Details (Excel):**
    - Each Department must be in a **separate Sheet** (e.g., Sheet name: `BCA`, `CS`).
    - Inside each sheet, the header row must contain **`REG.NO`** (or `REGISTER NUMBER`).
    - You can include an optional **`YEAR`** column (e.g., I-YEAR, II-YEAR).
    
    **3. Time Table (Excel / PDF):**
    - Your standard university time table with Date, Session (FN/AN), and Subject Codes.
    """)

col1, col2, col3 = st.columns(3)
with col1:
    hall_file = st.file_uploader("1. Upload Hall Details (Excel)", type=['xlsx'])
with col2:
    student_file = st.file_uploader("2. Upload Student Details (Excel)", type=['xlsx'])
with col3:
    timetable_file = st.file_uploader("3. Upload Time Table (PDF or Excel)", type=['pdf', 'xlsx', 'xls'])

st.markdown("### 4. Exam Settings (Required for auto-detection)")
col_t, col_d, col_s = st.columns([2, 1, 1])
with col_t:
    exam_title = st.text_input("Exam Title", value="PERIYAR UNIVERSITY THEORY EXAMINATIONS - APR/MAY")
with col_d:
    exam_date = st.text_input("Exam Date (e.g., 28-04-2025)", value="")
with col_s:
    exam_session = st.selectbox("Session", ["FN", "AN"])

selected_classes = []
if student_file:
    try:
        available_classes = get_student_classes(student_file)
        
        # Try to auto-detect from PDF
        auto_detected = []
        if timetable_file and exam_date and exam_session:
            detected = get_active_classes_from_timetable(timetable_file, exam_date, exam_session)
            # Only keep detected classes that actually exist in the student file
            auto_detected = [c for c in detected if c in available_classes]
            
        st.markdown("### 5. Classes taking the Exam")
        st.info("💡 The classes below are automatically detected from the Time Table PDF based on your Date and Session! You can add or remove them if needed.")
        selected_classes = st.multiselect("Confirm Classes", available_classes, default=auto_detected)
    except Exception as e:
        st.error(f"Could not parse student details: {e}")

if st.button("Generate Seating Arrangement"):
    if hall_file and student_file and timetable_file and exam_date and selected_classes:
        with st.spinner("Generating seating arrangement..."):
            try:
                output_excel_bytes = process_allocation(hall_file, student_file, timetable_file, exam_date, exam_session, selected_classes, exam_title)
                
                st.success("✅ Seating Arrangement Generated Successfully!")
                
                st.download_button(
                    label="📥 Download Seating Arrangement",
                    data=output_excel_bytes,
                    file_name="Generated_Seating_Arrangement.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                import traceback
                st.error(f"❌ An error occurred during generation: {e}")
                st.code(traceback.format_exc())
    else:
        st.warning("⚠️ Please fill all fields and select at least one class to proceed.")
