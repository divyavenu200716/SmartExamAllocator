import streamlit as st
import pandas as pd
import io
import os
from allocator import process_allocation, get_student_classes, get_active_classes_from_timetable

st.set_page_config(page_title="KASC SMART SEAT", page_icon="logo.png", layout="centered", initial_sidebar_state="expanded")

# Initialize session state for persistence across navigation
if 'hall_file' not in st.session_state: st.session_state.hall_file = None
if 'student_files' not in st.session_state: st.session_state.student_files = None
if 'timetable_file' not in st.session_state: st.session_state.timetable_file = None
if 'exam_title' not in st.session_state: st.session_state.exam_title = "PERIYAR UNIVERSITY THEORY EXAMINATIONS - APR/MAY"
if 'exam_date' not in st.session_state: st.session_state.exam_date = ""
if 'exam_session' not in st.session_state: st.session_state.exam_session = "FN"
if 'selected_classes' not in st.session_state: st.session_state.selected_classes = []

# Hide Streamlit Watermark and Add Custom Modern CSS
custom_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
a[href*="streamlit"] {display: none !important;}
div[class^="viewerBadge"] {display: none !important;}
div[class^="creator"] {display: none !important;}
img[alt*="Creator"] {display: none !important;}
[title*="Creator"] {display: none !important;}
[data-testid="stHeader"] {display: none !important;}
.stDeployButton {display: none !important;}

/* App background with a subtle blue-black gradient */
.stApp {
    background: linear-gradient(135deg, #050505 0%, #0a1128 50%, #050505 100%);
}

/* Clean text colors */
h1, h2, h3, h4, h5, h6, p {
    color: #ffffff !important;
}

/* Sidebar styling */
[data-testid="stSidebar"] {
    background-color: #0b1320;
    border-right: 1px solid #1c2e4a;
}
.stRadio > div {
    gap: 15px;
}
.stRadio label {
    font-size: 1.1rem !important;
    font-weight: 500;
}

/* Stylish Primary Buttons */
.stButton>button {
    background: linear-gradient(45deg, #0072ff, #00c6ff);
    color: white !important;
    border: none;
    border-radius: 8px;
    box-shadow: 0 4px 15px rgba(0, 198, 255, 0.4);
    transition: all 0.3s ease;
    font-weight: bold;
}
.stButton>button:hover {
    box-shadow: 0 6px 20px rgba(0, 198, 255, 0.6);
    transform: translateY(-2px);
}

/* Expander headers and uploaders */
.stExpander, div[data-testid="stFileUploader"] {
    background-color: #121f33;
    border: 1px solid #1f3659;
    border-radius: 10px;
    box-shadow: 0 4px 6px rgba(0,0,0,0.3);
}

/* Download buttons and links */
.stDownloadButton>button {
    background: linear-gradient(45deg, #11998e, #38ef7d);
    color: white !important;
    border: none;
    box-shadow: 0 4px 15px rgba(56, 239, 125, 0.4);
}
.stDownloadButton>button:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(56, 239, 125, 0.6);
}
</style>
"""
st.markdown(custom_style, unsafe_allow_html=True)

# Sidebar Navigation
st.sidebar.title("🗂️ Menu")
menu = st.sidebar.radio("Go to:", [
    "1. Hall Details", 
    "2. Student Details", 
    "3. Time Table", 
    "4. Exam Settings", 
    "5. Classes Taking Exam",
    "6. Generate Seating"
])

# Main Page Header
col1, col2 = st.columns([1, 5])
with col1:
    try:
        st.image("logo.png", use_container_width=True)
    except:
        pass
with col2:
    st.title("🎓 KASC SMART SEAT")
st.markdown("---")

# Render View Based on Selection
if menu == "1. Hall Details":
    st.header("📂 1. Hall Details")
    st.markdown("*(Excel: HALL NO, TOTAL SEAT, ROWS, COLUMNS)*")
    hf = st.file_uploader("Select Hall File", type=['xlsx', 'xls', 'csv', 'pdf'])
    if hf is not None:
        st.session_state.hall_file = hf
    if st.session_state.hall_file:
        st.success(f"✅ Uploaded: {st.session_state.hall_file.name}")

elif menu == "2. Student Details":
    st.header("👥 2. Student Details")
    st.markdown("*(Upload one or multiple files or drag a folder)*")
    sf = st.file_uploader("Select Student Files", type=['xlsx', 'xls', 'csv', 'pdf'], accept_multiple_files=True)
    if sf: # If the list is not empty
        st.session_state.student_files = sf
    if st.session_state.student_files:
        st.success(f"✅ Uploaded {len(st.session_state.student_files)} file(s)")

elif menu == "3. Time Table":
    st.header("📅 3. Time Table")
    tf = st.file_uploader("Select Time Table", type=['pdf', 'xlsx', 'xls', 'csv'])
    if tf is not None:
        st.session_state.timetable_file = tf
    if st.session_state.timetable_file:
        st.success(f"✅ Uploaded: {st.session_state.timetable_file.name}")

elif menu == "4. Exam Settings":
    st.header("⚙️ 4. Exam Settings")
    st.session_state.exam_title = st.text_input("Exam Title", value=st.session_state.exam_title)
    st.session_state.exam_date = st.text_input("Exam Date (e.g., 28-04-2025)", value=st.session_state.exam_date)
    st.session_state.exam_session = st.selectbox("Session", ["FN", "AN"], index=0 if st.session_state.exam_session=="FN" else 1)
    st.success("✅ Settings saved!")

elif menu == "5. Classes Taking Exam":
    st.header("✅ 5. Classes Taking Exam")
    if not st.session_state.student_files:
        st.warning("⚠️ Please upload Student Details first (Step 2).")
    else:
        try:
            available_classes = get_student_classes(st.session_state.student_files)
            auto_detected = []
            
            if st.session_state.timetable_file and st.session_state.exam_date and st.session_state.exam_session:
                detected = get_active_classes_from_timetable(st.session_state.timetable_file, st.session_state.exam_date, st.session_state.exam_session)
                if not detected:
                    st.warning("⚠️ Auto-detection couldn't find classes in Time Table.")
                auto_set = set()
                import re
                for d in detected:
                    d_clean = re.sub(r'[^A-Z0-9]', '', d.upper()).replace('IIIYEAR', '3YEAR').replace('IIYEAR', '2YEAR').replace('IYEAR', '1YEAR')
                    for a in available_classes:
                        a_clean = re.sub(r'[^A-Z0-9]', '', a.upper()).replace('IIIYEAR', '3YEAR').replace('IIYEAR', '2YEAR').replace('IYEAR', '1YEAR')
                        if d_clean == a_clean or d_clean in a_clean or a_clean in d_clean:
                            auto_set.add(a)
                auto_detected = list(auto_set)
                if auto_detected:
                    st.success("✨ Classes auto-detected from Time Table!")
            else:
                st.info("💡 Enter Time Table & Date (Steps 3 & 4) to auto-detect classes")

            # Use session_state fallback
            if not st.session_state.selected_classes and auto_detected:
                default_classes = auto_detected
            else:
                # keep previously selected if valid
                default_classes = [c for c in st.session_state.selected_classes if c in available_classes]
                if not default_classes and auto_detected: default_classes = auto_detected

            selected = st.multiselect("Confirm Classes", available_classes, default=default_classes)
            st.session_state.selected_classes = selected
            
        except Exception as e:
            st.error(f"Could not parse student details: {e}")

elif menu == "6. Generate Seating":
    st.header("🚀 6. Generate Seating Arrangement")
    
    # Check what is missing
    missing = []
    if not st.session_state.hall_file: missing.append("Hall Details (Step 1)")
    if not st.session_state.student_files: missing.append("Student Details (Step 2)")
    if not st.session_state.timetable_file: missing.append("Time Table (Step 3)")
    if not st.session_state.exam_date: missing.append("Exam Date (Step 4)")
    if not st.session_state.selected_classes: missing.append("Classes (Step 5)")

    if missing:
        st.warning(f"⚠️ Please complete the following steps first: {', '.join(missing)}")
    else:
        st.success("✅ All inputs are ready! Click the button below.")
        if st.button("✨ Generate Seating Arrangement", use_container_width=True):
            with st.spinner("Generating seating arrangement..."):
                try:
                    output_excel_bytes = process_allocation(
                        st.session_state.hall_file, 
                        st.session_state.student_files, 
                        st.session_state.timetable_file, 
                        st.session_state.exam_date, 
                        st.session_state.exam_session, 
                        st.session_state.selected_classes, 
                        st.session_state.exam_title
                    )
                    
                    st.success("🎉 Seating Arrangement Generated Successfully!")
                    st.download_button(
                        label="📥 Download Seating Arrangement",
                        data=output_excel_bytes,
                        file_name="Generated_Seating_Arrangement.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                except Exception as e:
                    import traceback
                    st.error(f"❌ An error occurred during generation: {e}")
                    st.code(traceback.format_exc())

st.markdown("---")
st.markdown("<p style='text-align: center; color: #a0aec0; font-size: 0.9rem;'>💡 Developed by <b>Divyavenugopal (Final Year AI&DS)</b></p>", unsafe_allow_html=True)
