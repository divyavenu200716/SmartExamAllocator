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
    st.logo("logo.png") 