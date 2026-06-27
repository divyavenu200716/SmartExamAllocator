import streamlit as st
import pandas as pd
import streamlit as st

# Logo-vukku pakila indha code-a podunga
# 'logo.png' endra file unga folder-la irukkanum
try:
    st.sidebar.image("logo.png", use_container_width=True)
except:
    st.sidebar.write("Logo file not found!")# Logo setup
st.logo("logo.png") 
st.title("Smart Exam Allocator")

# Session State Initialization
if "staff_logged_in" not in st.session_state: st.session_state.staff_logged_in = False
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
            st.session_state[state_key] = pd.concat([st.session_state[state_key], df_new], ignore_index=True)
            st.success("Uploaded!")
        except Exception as e:
            st.error(f"Error: {e}")

# Sidebar Menu
menu = st.sidebar.selectbox("Choose Module", [
    "Staff Login", "Room Allocation", "Student Details", "Exam Timetable", "Staff Allocation", "Final Allocation"
])

# 1. Staff Login
if menu == "Staff Login":
    st.header("Staff Login")
    name = st.text_input("Name")
    s_id = st.text_input("Staff ID")
    if st.button("Login"):
        st.session_state.staff_logged_in = True
        st.success(f"Welcome {name}!")

# 2. Room Allocation
elif menu == "Room Allocation":
    st.header("Room Allocation")
    handle_upload("rooms", ["Room No", "Capacity", "Block"])
    with st.form("room_form"):
        r_no = st.text_input("Room No")
        cap = st.number_input("Capacity", min_value=1)
        blk = st.text_input("Block")
        if st.form_submit_button("Add Room"):
            st.session_state.rooms = pd.concat([st.session_state.rooms, pd.DataFrame([{"Room No": r_no, "Capacity": cap, "Block": blk}])], ignore_index=True)
    st.dataframe(st.session_state.rooms)

# 3. Student Details
elif menu == "Student Details":
    st.header("Student Details")
    handle_upload("students", ["Name", "Dept", "Reg No", "Exam Code"])
    with st.form("student_form"):
        name = st.text_input("Name")
        dept = st.text_input("Dept")
        reg = st.text_input("Reg No")
        ex = st.text_input("Exam Code")
        if st.form_submit_button("Add Student"):
            st.session_state.students = pd.concat([st.session_state.students, pd.DataFrame([{"Name": name, "Dept": dept, "Reg No": reg, "Exam Code": ex}])], ignore_index=True)
    st.dataframe(st.session_state.students)

# 4. Exam Timetable
elif menu == "Exam Timetable":
    st.header("Exam Timetable")
    handle_upload("timetable", ["Exam", "Date", "Dept", "Year", "Subject", "Code"])
    with st.form("time_form"):
        ex = st.text_input("Exam")
        dt = st.text_input("Date")
        dp = st.text_input("Dept")
        yr = st.text_input("Year")
        sub = st.text_input("Subject")
        cod = st.text_input("Code")
        if st.form_submit_button("Add Exam"):
            st.session_state.timetable = pd.concat([st.session_state.timetable, pd.DataFrame([{"Exam": ex, "Date": dt, "Dept": dp, "Year": yr, "Subject": sub, "Code": cod}])], ignore_index=True)
    st.dataframe(st.session_state.timetable)

# 5. Staff Allocation
elif menu == "Staff Allocation":
    st.header("Staff Allocation")
    handle_upload("staff", ["Staff Name", "Staff ID"])
    with st.form("staff_form"):
        s_name = st.text_input("Staff Name")
        s_id = st.text_input("Staff ID")
        if st.form_submit_button("Add Staff"):
            st.session_state.staff = pd.concat([st.session_state.staff, pd.DataFrame([{"Staff Name": s_name, "Staff ID": s_id}])], ignore_index=True)
    st.dataframe(st.session_state.staff)

# 6. Final Allocation Logic
elif menu == "Final Allocation":
    st.header("Final Allocation")
    if st.button("Generate Allocation"):
        if not st.session_state.students.empty and not st.session_state.rooms.empty:
            # Simple Logic: Assign students to rooms based on capacity
            all_students = st.session_state.students.copy()
            rooms = st.session_state.rooms.copy()
            
            allocations = []
            student_idx = 0
            
            for index, room in rooms.iterrows():
                capacity = room["Capacity"]
                count = 0
                while count < capacity and student_idx < len(all_students):
                    student = all_students.iloc[student_idx]
                    allocations.append({
                        "Room No": room["Room No"],
                        "Student Name": student["Name"],
                        "Reg No": student["Reg No"],
                        "Exam Code": student["Exam Code"]
                    })
                    student_idx += 1
                    count += 1
            
            st.success("Allocation Generated!")
            st.dataframe(pd.DataFrame(allocations))
        else:
            st.warning("Please add Students and Rooms first!")
