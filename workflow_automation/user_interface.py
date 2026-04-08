import streamlit as st

st.title("PO Validator")
file = st.file_uploader("Upload PDF")

if file:
    st.success("File uploaded")