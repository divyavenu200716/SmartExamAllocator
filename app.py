import streamlit as st
import pandas as pd

# Page Config
st.set_page_config(page_title="Smart Exam Allocator", layout="wide")

st.title("🎓 Smart Exam Allocator")

# Sidebar Menu
menu = st.sidebar.selectbox("Choose Module", ["Admin Panel", "Scanner"])

if menu == "Admin Panel":
    st.header("⚙️ Admin Panel")
    st.write("Add your details here.")
elif menu == "Scanner":
    st.header("🔍 Teacher Scanner")
    st.write("Scanner feature will be added soon.")
    st.logo("logo.png") import streamlit as st
import pandas as pd

# Sidebar la puthu option
menu = st.sidebar.selectbox("Choose Module", ["Admin Panel", "Scanner", "Student Upload", "Exam Timetable"])

if menu == "Exam Timetable":
    st.header("Upload Exam Timetable")
    uploaded_tt = st.file_uploader("Choose an Excel file for Timetable", type=["xlsx"])
    
    if uploaded_tt is not None:
        try:
            df_tt = pd.read_excel(uploaded_tt)
            st.write("Exam Schedule Details:")
            st.table(df_tt) # st.table use panna azhaga display aagum
        except Exception as e:
            st.error(f"Error: {e}")
            import streamlit as st
import pandas as pd

# Sidebar la puthu option
menu = st.sidebar.selectbox("Choose Module", ["Admin Panel", "Scanner", "Student Upload", "Exam Timetable"])

if menu == "Exam Timetable":
    st.header("Upload Exam Timetable")
    uploaded_tt = st.file_uploader("Choose an Excel file for Timetable", type=["xlsx"])
    
    if uploaded_tt is not None:
        try:
            df_tt = pd.read_excel(uploaded_tt)
            st.write("Exam Schedule Details:")
            st.table(df_tt) # st.table use panna azhaga display aagum
        except Exception as e:
            st.error(f"Error: {e}")