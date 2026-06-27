
    import streamlit as st
import pandas as pd

st.title("Smart Exam Allocator")

# Session state-la rooms-ai store panna
if "rooms" not in st.session_state:
    st.session_state.rooms = pd.DataFrame(columns=["Room Number", "Capacity", "Block"])

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

# Helper function for upload
def upload_and_display(label):
    uploaded_file = st.file_uploader(f"Upload {label} (Excel)", type=["xlsx", "xls"])
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.success("File uploaded successfully!")
            st.dataframe(df)
        except Exception as e:
            st.error(f"Error reading file: {e}")

# Main Logic
if menu == "Staff Login":
    st.header("Staff Login")
    if not st.session_state.staff_logged_in:
        name = st.text_input("Enter Name")
        staff_id = st.text_input("Enter Staff ID")
        if st.button("Login"):
            if name and staff_id:
                st.session_state.staff_logged_in = True
                st.session_state.staff_name = name
                st.success(f"Welcome {name}!")
                st.rerun()
            else:
                st.error("Please enter both Name and ID")
    else:
        st.success(f"Logged in as: {st.session_state.staff_name}")
        if st.button("Logout"):
            st.session_state.staff_logged_in = False
            st.rerun()

elif menu == "Room Allocation":
    st.header("Room Allocation")
    
    # Option 1: File Upload
    with st.expander("Upload Excel File"):
        upload_and_display("Room Data")
    
    st.write("---")
    
    # Option 2: Manual Input Form
    st.subheader("Add Room Manually")
    with st.form("room_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            r_no = st.text_input("Room Number")
        with col2:
            cap = st.number_input("Capacity", min_value=1)
        with col3:
            block = st.text_input("Block/Building")
        
        submitted = st.form_submit_button("Add Room")
        if submitted:
            new_room = {"Room Number": r_no, "Capacity": cap, "Block": block}
            st.session_state.rooms = pd.concat([st.session_state.rooms, pd.DataFrame([new_room])], ignore_index=True)
            st.success(f"Room {r_no} added!")

    # Display current rooms
    st.write("### Current Room List")
    st.dataframe(st.session_state.rooms)

elif menu == "Student Upload":
    st.header("Student Upload")
    upload_and_display("Student Data")

elif menu == "Exam Timetable":
    st.header("Exam Timetable")
    upload_and_display("Exam Timetable")

elif menu == "Staff Allocation":
    st.header("Staff Allocation")
    upload_and_display("Staff Data")
