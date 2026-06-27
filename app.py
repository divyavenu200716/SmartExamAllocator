import streamlit as st
import pandas as pd

# Logo setup
st.logo("logo.png") 

st.title("Smart Exam Allocator")

# Session State Initialization (Data-vai save panni vaikka)
if "staff_logged_in" not in st.session_state: st.session_state.staff_logged_in = False
if "staff_name" not in st.session_state: st.session_state.staff_name = ""
if "students" not in st.session_state: st.session_state.students = pd.DataFrame(columns=["Name", "Dept", "Reg No", "Exam Code"])
if "rooms" not in st.session_state: st.session_state.rooms = pd.DataFrame(columns=["Room No", "Capacity", "Block"])
if "timetable" not in st.session_state: st.session_state.timetable = pd.DataFrame(columns=["Exam", "Date", "Dept", "Year", "Subject", "Code"])
if "staff" not in st.session_state: st.session_state.staff = pd.DataFrame(columns=["Staff Name", "Staff ID"])

# Sidebar Menu
menu = st.sidebar.selectbox("Choose Module", [
    "Staff Login", "Room Allocation", "Student Details", "Exam Timetable", "Staff Allocation", "Final Allocation"
])

# 1. Staff Login
if menu == "Staff Login":
    st.header("Staff Login")
    if not st.session_state.staff_logged_in:
        name = st.text_input("Enter Name")
        s_id = st.text_input("Enter Staff ID")
        if st.button("Login"):
            st.session_state.staff_logged_in = True
            st.session_state.staff_name = name
            st.success(f"Welcome {name}! Presence recorded.")
            st.rerun()
    else:
        st.write(f"Logged in as: **{st.session_state.staff_name}**")
        if st.button("Logout"):
            st.session_state.staff_logged_in = False
            st.rerun()

# 2. Room Allocation (Manual Input)
elif menu == "Room Allocation":
    st.header("Room Allocation")
    with st.form("room_form"):
        r_no = st.text_input("Room Number")
        cap = st.number_input("Capacity", min_value=1)
        blk = st.text_input("Block")
        if st.form_submit_button("Add Room"):
            st.session_state.rooms = pd.concat([st.session_state.rooms, pd.DataFrame([{"Room No": r_no, "Capacity": cap, "Block": blk}])], ignore_index=True)
    st.dataframe(st.session_state.rooms)

# 3. Student Details
elif menu == "Student Details":
    st.header("Student Details")
    with st.form("student_form"):
        name = st.text_input("Student Name")
        dept = st.text_input("Department")
        reg = st.text_input("Register Number")
        ex_code = st.text_input("Exam Code")
        if st.form_submit_button("Add Student"):
            st.session_state.students = pd.concat([st.session_state.students, pd.DataFrame([{"Name": name, "Dept": dept, "Reg No": reg, "Exam Code": ex_code}])], ignore_index=True)
    st.dataframe(st.session_state.students)

# 4. Exam Timetable
elif menu == "Exam Timetable":
    st.header("Exam Timetable")
    with st.form("time_form"):
        ex = st.text_input("Exam Name")
        dt = st.date_input("Date")
        dp = st.text_input("Department")
        yr = st.text_input("Year")
        sub = st.text_input("Subject")
        cod = st.text_input("Subject Code")
        if st.form_submit_button("Add Exam"):
            st.session_state.timetable = pd.concat([st.session_state.timetable, pd.DataFrame([{"Exam": ex, "Date": str(dt), "Dept": dp, "Year": yr, "Subject": sub, "Code": cod}])], ignore_index=True)
    st.dataframe(st.session_state.timetable)

# 5. Staff Allocation
elif menu == "Staff Allocation":
    st.header("Staff Allocation")
    with st.form("staff_form"):
        s_name = st.text_input("Staff Name")
        s_id = st.text_input("Staff ID")
        if st.form_submit_button("Allocate Staff"):
            st.session_state.staff = pd.concat([st.session_state.staff, pd.DataFrame([{"Staff Name": s_name, "Staff ID": s_id}])], ignore_index=True)
    st.dataframe(st.session_state.staff)

# 6. Final Allocation (Logic)
elif menu == "Final Allocation":
    st.header("Final Allocation Summary")
    if st.button("Generate Exam Schedule"):
        st.write("Allocation Logic Running...")
        st.write("--- Students ---")
        st.dataframe(st.session_state.students)
        st.write("--- Rooms Assigned ---")
        st.dataframe(st.session_state.rooms)
        st.success("Allocation Process Completed!")
