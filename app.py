import streamlit as st
import pandas as pd

st.title("Smart Exam Allocator")

# Initialize session state for login
if "staff_logged_in" not in st.session_state:
    st.session_state.staff_logged_in = False
    st.session_state.staff_name = ""

# Sidebar Menu
menu = st.sidebar.selectbox("Choose Module", [
    "Staff Login", 
    "Room Allocation", 
    "Student Upload", 
    "Exam Timetable", 
    "Staff Allocation"
])

# Helper function to handle excel upload
def upload_and_display(label):
    uploaded_file = st.file_uploader(f"Upload {label} (Excel)", type=["xlsx", "xls"])
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.success("File uploaded successfully!")
            st.dataframe(df)
        except Exception as e:
            st.error(f"Error reading file: {e}")

# Logic for each module
if menu == "Staff Login":
    st.header("Staff Login")
    
    if not st.session_state.staff_logged_in:
        name = st.text_input("Enter Name")
        staff_id = st.text_input("Enter Staff ID")
        if st.button("Login"):
            if name and staff_id:
                st.session_state.staff_logged_in = True
                st.session_state.staff_name = name
                st.success(f"Welcome {name}! Presence recorded.")
                st.rerun() # Refresh pannum
            else:
                st.error("Please enter both Name and ID")
    else:
        st.success(f"Logged in as: {st.session_state.staff_name}")
        if st.button("Logout"):
            st.session_state.staff_logged_in = False
            st.rerun()

elif menu == "Room Allocation":
    st.header("Room Allocation")
    upload_and_display("Room Data")

elif menu == "Student Upload":
    st.header("Student Upload")
    upload_and_display("Student Data")

elif menu == "Exam Timetable":
    st.header("Exam Timetable")
    upload_and_display("Exam Timetable")

elif menu == "Staff Allocation":
    st.header("Staff Allocation")
    upload_and_display("Staff Data")
