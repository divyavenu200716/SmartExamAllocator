import streamlit as st
import pandas as pd

# Logo setup
st.logo("logo.png") 
st.title("Smart Exam Allocator")

# Session State Initialization
if "staff_logged_in" not in st.session_state: st.session_state.staff_logged_in = False
if "staff_name" not in st.session_state: st.session_state.staff_name = ""
if "students" not in st.session_state: st.session_state.students = pd.DataFrame(columns=["Name", "Dept", "Reg No", "Exam Code"])
if "rooms" not in st.session_state: st.session_state.rooms = pd.DataFrame(columns=["Room No", "Capacity", "Block"])
if "timetable" not in st.session_state: st.session_state.timetable = pd.DataFrame(columns=["Exam", "Date", "Dept", "Year", "Subject", "Code"])
if "staff" not in st.session_state: st.session_state.staff = pd.DataFrame(columns=["Staff Name", "Staff ID"])

# Helper Function for Excel Upload
def handle_upload(state_key, expected_columns):
    uploaded_file = st.file_uploader(f"Upload Excel for {state_key}", type=["xlsx", "xls"], key=f"up_{state_key}")
    if uploaded_file:
        try:
            df_new = pd.read_excel(uploaded_file)
            # Column name check
            if list(df_new.columns) == expected_columns:
                st.session_state[state_key] = pd.concat([st.session_state[state_key], df_new], ignore_index=True)
                st.success("Data uploaded successfully!")
            else:
                st.error(f"Column mismatch! Please ensure your Excel has these headers: {expected_columns}")
        except Exception as e:
            st.error(f"Error: {e}")

# Sidebar Menu
menu = st.sidebar.selectbox("Choose Module", [
    "Staff Login", "Room Allocation", "Student Details", "Exam Timetable", "Staff Allocation"
])

# --- 1. Staff Login ---
if menu == "Staff Login":
    st.header("Staff Login")
    if not st.session_state.staff_logged_in:
        name = st.text_input("Enter Name")
        s_id = st.text_input("Enter Staff ID")
        if st.button("Login"):
            st.session_state.staff_logged_in = True
            st.session_state.staff_name = name
            st.success(f"Welcome {name}!")
            st.rerun()
    else:
        st.write(f"Logged in as: **{st.session_state.staff_name}**")
        if st.button("Logout"):
            st.session_state.staff_logged_in = False
            st.rerun()

# --- 2. Room Allocation ---
elif menu == "Room Allocation":
    st.header("Room Allocation")
    # Excel Upload
    with st.expander("1. Upload Excel"):
        handle_upload("rooms", ["Room No", "Capacity", "Block"])
    # Manual Input
    with st.expander("2. Add Manually"):
        with st.form("room_form"):
            r_no = st.text_input("Room Number")
            cap = st.number_input("Capacity", min_value=1)
            blk = st.text_input("Block")
            if st.form_submit_button("Add Room"):
                new_data = pd.DataFrame([{"Room No": r_no, "Capacity": cap, "Block": blk}])
                st.session_state.rooms = pd.concat([st.session_state.rooms, new_data], ignore_index=True)
                st.success("Added!")
    st.dataframe(st.session_state.rooms)

# --- 3. Student Details ---
elif menu == "Student Details":
    st.header("Student Details")
    with st.expander("1. Upload Excel"):
        handle_upload("students", ["Name", "Dept", "Reg No", "Exam Code"])
    with st.expander("2. Add Manually"):
        with st.form("student_form"):
            name = st.text_input("Student Name")
            dept = st.text_input("Department")
            reg = st.text_input("Register Number")
            ex_code = st.text_input("Exam Code")
            if st.form_submit_button("Add Student"):
                new_data = pd.DataFrame([{"Name": name, "Dept": dept, "Reg No": reg, "Exam Code": ex_code}])
                st.session_state.students = pd.concat([st.session_state.students, new_data], ignore_index=True)
    st.dataframe(st.session_state.students)

# --- 4. Exam Timetable ---
elif menu == "Exam Timetable":
    st.header("Exam Timetable")
    with st.expander("1. Upload Excel"):
        handle_upload("timetable", ["Exam", "Date", "Dept", "Year", "Subject", "Code"])
    with st.expander("2. Add Manually"):
        with st.form("time_form"):
            ex = st.text_input("Exam Name")
            dt = st.date_input("Date")
            dp = st.text_input("Department")
            yr = st.text_input("Year")
            sub = st.text_input("Subject")
            cod = st.text_input("Subject Code")
            if st.form_submit_button("Add Exam"):
                new_data = pd.DataFrame([{"Exam": ex, "Date": str(dt), "Dept": dp, "Year": yr, "Subject": sub, "Code": cod}])
                st.session_state.timetable = pd.concat([st.session_state.timetable, new_data], ignore_index=True)
    st.dataframe(st.session_state.timetable)

# --- 5. Staff Allocation ---
elif menu == "Staff Allocation":
    st.header("Staff Allocation")
    with st.expander("1. Upload Excel"):
        handle_upload("staff", ["Staff Name", "Staff ID"])
    with st.expander("2. Add Manually"):
        with st.form("staff_form"):
            s_name = st.text_input("Staff Name")
            s_id = st.text_input("Staff ID")
            if st.form_submit_button("Allocate Staff"):
                new_data = pd.DataFrame([{"Staff Name": s_name, "Staff ID": s_id}])
                st.session_state.staff = pd.concat([st.session_state.staff, new_data], ignore_index=True)
    st.dataframe(st.session_state.staff)
