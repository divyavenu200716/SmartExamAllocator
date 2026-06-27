import streamlit as st
import pandas as pd

st.title("Smart Exam Allocator")

# 5 Options
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

# 5 Logic blocks
if menu == "Staff Login":
    st.header("Staff Login")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        st.success("Logged in successfully!")

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

