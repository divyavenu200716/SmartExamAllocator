import streamlit as st
import pandas as pd
import io
import os
from allocator import process_allocation, get_student_classes, get_active_classes_from_timetable

st.set_page_config(page_title="SMART SEAT", page_icon="logo.png", layout="wide")

st.title("🎓 SMART SEAT")

st.markdown("Upload your exam details below to automatically generate the seating arrangement.")

col1, col2, col3 = st.columns(3)
with col1:
    hall_file = st.file_uploader("1. Upload Hall Details (Excel or PDF)", type=['xlsx', 'xls', 'csv', 'pdf'])
    with st.expander("💡 View Sample Format"):
        st.markdown("For **Excel**, your table should look like this:")
        st.table(pd.DataFrame({"HALL NO": ["101", "102"], "TOTAL SEAT": [30, 36], "ROWS": [5, 6], "COLUMNS": [6, 6]}))
        st.markdown("For **PDF**, just ensure each line has: `HallNo TotalSeat Rows Cols` (e.g. `101 30 5 6`)")

with col2:
    student_files = st.file_uploader("2. Upload Student Details (Excel or PDF)", type=['xlsx', 'xls', 'csv', 'pdf'], accept_multiple_files=True)
    with st.expander("💡 View Sample Format"):
        st.markdown("Upload **one or more Excel/PDF files**.")
        st.table(pd.DataFrame({"REG.NO": ["23UCA01", "23UCA02"], "NAME": ["Arun", "Bala"], "YEAR": ["I-YEAR", "I-YEAR"]}))

with col3:
    timetable_file = st.file_uploader("3. Upload Time Table (PDF or Excel)", type=['pdf', 'xlsx', 'xls', 'csv'])
    with st.expander("👀 View Sample Format"):
        st.markdown("Just upload the standard University **PDF** or **Excel** Time Table directly!")

st.markdown("### 4. Exam Settings (Required for auto-detection)")
col_t, col_d, col_s = st.columns([2, 1, 1])
with col_t:
    exam_title = st.text_input("Exam Title", value="PERIYAR UNIVERSITY THEORY EXAMINATIONS - APR/MAY")
with col_d:
    exam_date = st.text_input("Exam Date (e.g., 28-04-2025)", value="")
with col_s:
    exam_session = st.selectbox("Session", ["FN", "AN"])

selected_classes = []
if student_files:
    try:
        available_classes = get_student_classes(student_files)
        
        # Try to auto-detect from PDF
        auto_detected = []
        if timetable_file and exam_date and exam_session:
            detected = get_active_classes_from_timetable(timetable_file, exam_date, exam_session)
            # Fuzzy match detected classes against available classes
            auto_set = set()
            for d in detected:
                d_clean = d.replace(' ', '').upper()
                for a in available_classes:
                    a_clean = a.replace(' ', '').upper()
                    if d_clean == a_clean or d_clean in a_clean or a_clean in d_clean:
                        auto_set.add(a)
            auto_detected = list(auto_set)
            
        st.markdown("### 5. Classes taking the Exam")
        st.info("💡 The classes below are automatically detected from the Time Table PDF based on your Date and Session! You can add or remove them if needed.")
        selected_classes = st.multiselect("Confirm Classes", available_classes, default=auto_detected)
    except Exception as e:
        st.error(f"Could not parse student details: {e}")

if st.button("Generate Seating Arrangement"):
    if hall_file and student_files and timetable_file and exam_date and selected_classes:
        with st.spinner("Generating seating arrangement..."):
            try:
                output_excel_bytes = process_allocation(hall_file, student_files, timetable_file, exam_date, exam_session, selected_classes, exam_title)
                
                st.success("✅ Seating Arrangement Generated Successfully!")
                
                st.download_button(
                    label="⬇️ Download Seating Arrangement",
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
